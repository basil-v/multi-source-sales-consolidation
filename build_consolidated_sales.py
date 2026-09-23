"""
build_consolidated_sales.py
---------------------------
Consolidates sales, budget and forecast from many sources into one wide,
analysis-ready table, then loads it to a database.

This is a sanitised, self-contained version of a production pipeline I built to
unify a multi-entity FMCG group's commercial data. All names, codes and numbers
here are synthetic (see generate_sample_data.py). The techniques are the real
ones:

  * ingest three differently-shaped distributor files (Excel, offset-header
    Excel, CSV) plus an ERP, and harmonise them to one schema
  * net invoices against credit memos and UNION actuals across entities
  * pull Budget, Last Estimate and open-order Forecast (Plan/Tentative/Firm)
  * derive cost three ways: rolling average, standard (price list) and a
    nearest-month / 4-month-average fallback
  * build a multi-stage intercompany transfer-price value chain and compute
    gross profit at each hop (factory -> hub -> distributor -> retail RSP)
  * load the result to SQL in chunks

By default it reads the local SQLite ERP and writes to a local SQLite output,
so it runs with no external database. Set the DB_* environment variables to
point at SQL Server instead (see .env.example).
"""

import os
import sqlite3
import calendar
import numpy as np
import pandas as pd
from datetime import datetime

BASE   = os.path.dirname(os.path.abspath(__file__))
INPUTS = os.path.join(BASE, "inputs")
ERP_DB = os.path.join(BASE, "db", "erp.db")
OUT_DB = os.path.join(BASE, "db", "output.db")
OUT_TABLE = "SALES_WITH_BUDGET"

MONTH_ABBR = {i: calendar.month_abbr[i].upper() for i in range(1, 13)}
MONTH_NUM  = {v: k for k, v in MONTH_ABBR.items()}

UNIFIED_COLS = [
    "Group Customer Code", "Distributor Customer Code", "Customer Full Name",
    "Channel", "Sub-Channel", "Customer Group", "Branch/Country", "City", "Region",
    "Salesman_Group", "Salesman", "ENTITY", "Factory Code", "Distributor Code",
    "Brand", "Category", "Segment", "Base", "Weight", "Weight UOM", "SKUs per Carton",
    "PTT", "QTY", "Amount", "Month", "Year", "ACT/FC", "Customer Price Group",
    "SalesPerson Status", "Business Unit",
]

erp = sqlite3.connect(ERP_DB)

# ------------------------------------------------------------------
# 1. Masters
# ------------------------------------------------------------------
def read_master(sheet):
    # skiprows=1 drops the title row; a blank leading column reads as 'Unnamed' and is dropped
    df = pd.read_excel(os.path.join(INPUTS, "master_classification.xlsx"),
                       sheet_name=sheet, skiprows=1)
    df = df.loc[:, ~df.columns.astype(str).str.startswith("Unnamed")]
    df.columns = df.columns.str.strip()
    return df

item_master = read_master("SKUS CLASSIFICATION")
customer_master = read_master("CHANNEL CLASSIFICATION")

for df, col in [(item_master, "Distributor Code"), (item_master, "Factory Code"),
                (customer_master, "Distributor Customer Code"), (customer_master, "Group Customer Code")]:
    df[col] = df[col].astype(str).str.strip()


def attach_masters(df, item_key, cust_key, item_on="Distributor Code", cust_on="Distributor Customer Code"):
    """Classify a raw feed by joining SKU and customer masters."""
    df[item_key] = df[item_key].astype(str).str.strip()
    df[cust_key] = df[cust_key].astype(str).str.strip()
    out = df.merge(item_master, left_on=item_key, right_on=item_on, how="left", suffixes=("", "_im"))
    out = out.merge(customer_master, left_on=cust_key, right_on=cust_on, how="left", suffixes=("", "_cm"))
    return out


def to_unified(df, entity, act_fc, status, qty, amount,
               month, year, cust_code_col=None, cust_name_col="Customer Full Name"):
    """Project any classified feed onto the single unified schema."""
    df = df.copy()
    df["ENTITY"] = entity
    df["ACT/FC"] = act_fc
    df["SalesPerson Status"] = status
    df["QTY"] = pd.to_numeric(df[qty], errors="coerce")
    df["Amount"] = pd.to_numeric(df[amount], errors="coerce")
    df["Month"] = df[month]
    df["Year"] = df[year]
    df["PTT"] = df["Amount"] / df["QTY"].replace(0, np.nan)
    if cust_code_col and cust_code_col in df.columns:
        df["Group Customer Code"] = df["Group Customer Code"].fillna(df[cust_code_col])
    for c in UNIFIED_COLS:
        if c not in df.columns:
            df[c] = np.nan
    return df[UNIFIED_COLS]


frames = []

# ------------------------------------------------------------------
# 2. Distributor feeds (three shapes -> one schema)
# ------------------------------------------------------------------
# DIST_A : Excel, numeric month, cartons + value
for f, act in [("DIST_A_ACTUAL_2024.xlsx", "ACTUAL"),
               ("DIST_A_ACTUAL_2025.xlsx", "ACTUAL"),
               ("DIST_A_BUDGET_2026.xlsx", "BUDGET")]:
    raw = pd.read_excel(os.path.join(INPUTS, "dist_a", f), sheet_name="DATA")
    raw["Month"] = raw["Month"].map(MONTH_ABBR)
    cl = attach_masters(raw, "Item", "Customer ID")
    frames.append(to_unified(cl, "DIST_A", act, "Actual" if act == "ACTUAL" else "Budget",
                             "QTY Cs", "Sales Value", "Month", "Year", cust_code_col="Customer ID"))

# DIST_B : Excel with 2 title rows, month name, cartons
for f, act in [("DIST_B_ACTUAL_2025.xlsx", "ACTUAL"),
               ("DIST_B_BUDGET_2026.xlsx", "BUDGET")]:
    raw = pd.read_excel(os.path.join(INPUTS, "dist_b", f), skiprows=2)
    raw = raw.rename(columns={"MONTHNAME": "Month", "YEAR": "Year"})
    raw["Month"] = raw["Month"].astype(str).str.upper()
    cl = attach_masters(raw, "ITEM CODE", "CUSTOMER CODE")
    frames.append(to_unified(cl, "DIST_B", act, "Actual" if act == "ACTUAL" else "Budget",
                             "QTY", "SALES", "Month", "Year", cust_code_col="CUSTOMER CODE"))

# DIST_C : CSV, YYYYMM period, different column names
for f, act in [("DIST_C_ACTUAL_2025.csv", "ACTUAL"),
               ("DIST_C_BUDGET_2026.csv", "BUDGET")]:
    raw = pd.read_csv(os.path.join(INPUTS, "dist_c", f))
    raw["Month"] = raw["MonthYear"].astype(str).str[-2:].astype(int).map(MONTH_ABBR)
    raw = raw.rename(columns={"SaleYear": "Year"})
    cl = attach_masters(raw, "ProductCode", "CustomerCode")
    frames.append(to_unified(cl, "DIST_C", act, "Actual" if act == "ACTUAL" else "Budget",
                             "NetSalesQty", "NetsalesValue", "Month", "Year", cust_code_col="CustomerCode"))

# ------------------------------------------------------------------
# 3. ERP actuals: invoices netted against credit memos, unioned by entity
# ------------------------------------------------------------------
sql_history = """
WITH src AS (
    SELECT entity, customer_code, item_code,
           SUM(qty_base * 1)                                    AS QTY,
           SUM(amount / COALESCE(NULLIF(currency_factor,0),1))  AS Amount,
           strftime('%m', posting_date) AS mm, strftime('%Y', posting_date) AS Year
    FROM sales_invoice
    WHERE posting_group <> 'MVEH'
    GROUP BY entity, customer_code, item_code, mm, Year
    UNION ALL
    SELECT entity, customer_code, item_code,
           -SUM(qty_base),
           -SUM(amount / COALESCE(NULLIF(currency_factor,0),1)),
           strftime('%m', posting_date), strftime('%Y', posting_date)
    FROM sales_creditmemo
    GROUP BY entity, customer_code, item_code, strftime('%m', posting_date), strftime('%Y', posting_date)
)
SELECT entity, customer_code, item_code, SUM(QTY) AS QTY, SUM(Amount) AS Amount, mm, Year
FROM src
GROUP BY entity, customer_code, item_code, mm, Year
HAVING SUM(Amount) <> 0
"""
hist = pd.read_sql(sql_history, erp)
hist["Month"] = hist["mm"].astype(int).map(MONTH_ABBR)
hist = hist.merge(item_master, left_on="item_code", right_on="Factory Code", how="left") \
           .merge(customer_master, left_on="customer_code", right_on="Group Customer Code", how="left")
frames.append(to_unified(hist, hist["entity"], "ACTUAL", "Actual", "QTY", "Amount",
                         "Month", "Year"))
# entity comes from the data itself, so set it explicitly
frames[-1]["ENTITY"] = hist["entity"].values

# ------------------------------------------------------------------
# 4. ERP budget + last estimate
# ------------------------------------------------------------------
for budget_name, act, status in [("BUDGET_V1", "BUDGET", "Budget"),
                                 ("BUDGET_V2", "LAST ESTIMATE", "Last Estimate")]:
    q = f"""
    SELECT entity, source_no AS customer_code, item_code,
           SUM(qty) AS QTY, SUM(sales_amount) AS Amount, unit_price AS PTT,
           strftime('%m', date) AS mm, strftime('%Y', date) AS Year
    FROM item_budget_entry
    WHERE budget_name = '{budget_name}' AND qty <> 0
    GROUP BY entity, source_no, item_code, unit_price, mm, Year
    """
    b = pd.read_sql(q, erp)
    b["Month"] = b["mm"].astype(int).map(MONTH_ABBR)
    b = b.merge(item_master, left_on="item_code", right_on="Factory Code", how="left") \
         .merge(customer_master, left_on="customer_code", right_on="Group Customer Code", how="left")
    u = to_unified(b, b["entity"], act, status, "QTY", "Amount", "Month", "Year")
    u["ENTITY"] = b["entity"].values
    frames.append(u)

# ------------------------------------------------------------------
# 5. ERP forecast (open orders, Plan/Tentative/Firm)
# ------------------------------------------------------------------
sql_fc = """
SELECT entity, customer_code, item_code,
       SUM(qty_base) AS QTY, SUM(amount) AS Amount,
       strftime('%m', requested_delivery_date) AS mm,
       strftime('%Y', requested_delivery_date) AS Year,
       MAX(salesperson_status) AS st
FROM sales_line
WHERE doc_type = 1 AND qty_base <> 0
GROUP BY entity, customer_code, item_code, mm, Year
"""
fc = pd.read_sql(sql_fc, erp)
fc["Month"] = fc["mm"].astype(int).map(MONTH_ABBR)
fc["SalesPerson Status"] = fc["st"].map({0: "Plan", 1: "Tentative", 2: "Firm"})
fc = fc.merge(item_master, left_on="item_code", right_on="Factory Code", how="left") \
       .merge(customer_master, left_on="customer_code", right_on="Group Customer Code", how="left")
u = to_unified(fc, fc["entity"], "FORECAST", "Forecast", "QTY", "Amount", "Month", "Year")
u["ENTITY"] = fc["entity"].values
u["SalesPerson Status"] = fc["SalesPerson Status"].values
frames.append(u)

# ------------------------------------------------------------------
# 6. Combine everything
# ------------------------------------------------------------------
main = pd.concat(frames, ignore_index=True)
main["Factory Code"] = main["Factory Code"].astype(str).str.strip()
main["UpdatedTime"] = datetime.now()
main["Quarter"] = main["Month"].map(lambda m: {
    "JAN": "Q1", "FEB": "Q1", "MAR": "Q1", "APR": "Q2", "MAY": "Q2", "JUN": "Q2",
    "JUL": "Q3", "AUG": "Q3", "SEP": "Q3", "OCT": "Q4", "NOV": "Q4", "DEC": "Q4"}.get(str(m).upper(), ""))
print(f"combined rows: {len(main):,}")

# ------------------------------------------------------------------
# 7. Cost:  rolling average, standard (price list), nearest-month fill
# ------------------------------------------------------------------
# 7a. rolling average output cost (last 6 months) from ledger + value entry
avg_cost = pd.read_sql("""
SELECT ile.item_code AS 'Factory Code',
       SUM(ve.cost_amount) / NULLIF(SUM(ile.quantity), 0) AS AvgCost
FROM item_ledger_entry ile
JOIN value_entry ve
  ON ve.item_ledger_entry_no = ile.entry_no AND ve.item_code = ile.item_code AND ve.entity = ile.entity
WHERE ile.entry_type = 6
GROUP BY ile.item_code
""", erp)
main = main.merge(avg_cost, on="Factory Code", how="left")

# 7b. standard cost from the generic price list (source_group 10)
std_cost = pd.read_sql("""
SELECT product_no AS 'Factory Code', AVG(unit_price) AS Std_Cost
FROM price_list_line
WHERE source_group = 10
GROUP BY product_no
""", erp)
main = main.merge(std_cost, on="Factory Code", how="left")

# 7c. actual monthly output cost, for nearest-month fallback
month_cost = pd.read_sql("""
SELECT ile.item_code AS 'Factory Code', ile.entity AS ENTITY,
       strftime('%m', ile.posting_date) AS mm, strftime('%Y', ile.posting_date) AS Year,
       SUM(ve.cost_amount) / NULLIF(SUM(ile.quantity), 0) AS MonthCost
FROM item_ledger_entry ile
JOIN value_entry ve
  ON ve.item_ledger_entry_no = ile.entry_no AND ve.entity = ile.entity
WHERE ile.entry_type = 6
GROUP BY ile.item_code, ile.entity, mm, Year
""", erp)
month_cost["Month"] = month_cost["mm"].astype(int).map(MONTH_ABBR)
month_cost["key"] = month_cost["Year"].astype(str) + month_cost["mm"]

# nearest-month cost per (item): fall back to the closest available month
cost_hist = {}
for _, r in month_cost.sort_values("key").iterrows():
    cost_hist.setdefault(r["Factory Code"], []).append((r["key"], r["MonthCost"]))

def nearest_cost(item, year, month_abbr):
    hist = cost_hist.get(item)
    if not hist:
        return np.nan
    target = f"{year}{MONTH_NUM.get(str(month_abbr).upper(), 0):02d}"
    prior = [c for k, c in hist if k <= target]
    if prior:
        return prior[-1]
    return hist[0][1]           # else earliest available

main["ACT/STD"] = main.apply(
    lambda r: (r["AvgCost"] if pd.notna(r["AvgCost"])
               else nearest_cost(r["Factory Code"], r["Year"], r["Month"]))
    if r["ACT/FC"] == "ACTUAL"
    else (r["Std_Cost"] if pd.notna(r["Std_Cost"])
          else nearest_cost(r["Factory Code"], r["Year"], r["Month"])),
    axis=1)

# ------------------------------------------------------------------
# 8. Transfer-price value chain  (factory -> hub -> distributor/retail)
# ------------------------------------------------------------------
def price_lookup(sql):
    df = pd.read_sql(sql, erp)
    return dict(zip(df["Factory Code"], df["price"]))

tp_fac_hub = price_lookup("""
    SELECT product_no AS 'Factory Code', AVG(unit_price) AS price
    FROM price_list_line WHERE source_no = 'GRP-C-100' GROUP BY product_no""")
tp_hub_dist = price_lookup("""
    SELECT product_no AS 'Factory Code', AVG(unit_price) AS price
    FROM price_list_line WHERE source_no = 'GRP-C-101' AND entity = 'HUB1' GROUP BY product_no""")
tp_hub_trd = price_lookup("""
    SELECT product_no AS 'Factory Code', AVG(unit_price) AS price
    FROM price_list_line WHERE source_no = 'GRP-C-102' AND entity = 'HUB1' GROUP BY product_no""")

main["Factory_TP"] = main["Factory Code"].map(tp_fac_hub)
main["Hub_to_Dist_TP"] = main["Factory Code"].map(tp_hub_dist)
main["Hub_to_Trd_TP"] = main["Factory Code"].map(tp_hub_trd)

# 8b. RSP: weighted-latest retail shelf price per item x customer group x month
rsp = pd.read_sql("""
SELECT item_code AS 'Factory Code', group_name AS 'Customer Group',
       strftime('%m', date) AS mm, strftime('%Y', date) AS Year,
       AVG(price) AS RSP
FROM availability_report_data
WHERE price > 0
GROUP BY item_code, group_name, mm, Year
""", erp)
rsp["Month"] = rsp["mm"].astype(int).map(MONTH_ABBR)
rsp["Year"] = rsp["Year"].astype(int)
main["Year"] = pd.to_numeric(main["Year"], errors="coerce").astype("Int64")
main = main.merge(rsp[["Factory Code", "Customer Group", "Month", "Year", "RSP"]],
                  on=["Factory Code", "Customer Group", "Month", "Year"], how="left")

# ------------------------------------------------------------------
# 9. Value totals and gross profit at each stage of the chain
# ------------------------------------------------------------------
main["Total_Cost"]        = main["ACT/STD"] * main["QTY"]
main["Total_Factory_TP"]  = main["Factory_TP"] * main["QTY"]
main["Total_Hub_to_Dist"] = main["Hub_to_Dist_TP"] * main["QTY"]
main["Total_RSP"]         = main["RSP"] * main["QTY"]

main["Group_GP"]   = main["Amount"] - main["Total_Cost"]            # end sale vs factory cost
main["Factory_GP"] = main["Total_Factory_TP"] - main["Total_Cost"]  # factory margin
main["Hub_GP"]     = main["Total_Hub_to_Dist"] - main["Total_Factory_TP"]  # hub margin
main["Retail_GP"]  = main["Total_RSP"] - main["Total_Hub_to_Dist"]  # distributor/retail margin
main["GP_Percent"] = 1 - (main["Total_Cost"] / main["Amount"].replace(0, np.nan))

# ------------------------------------------------------------------
# 10. Clean numerics and load to SQL in chunks
# ------------------------------------------------------------------
float_cols = main.select_dtypes(include=["float64", "float32"]).columns
main[float_cols] = main[float_cols].round(6)
main.replace([np.inf, -np.inf], np.nan, inplace=True)

erp.close()

# Default sink: local SQLite. Set DB_SERVER etc. to target SQL Server instead.
if os.environ.get("DB_SERVER"):
    import urllib.parse
    from sqlalchemy import create_engine
    params = urllib.parse.quote_plus(
        f"DRIVER={{ODBC Driver 17 for SQL Server}};SERVER={os.environ['DB_SERVER']};"
        f"DATABASE={os.environ['DB_NAME']};UID={os.environ['DB_USERNAME']};PWD={os.environ['DB_PASSWORD']}")
    engine = create_engine(f"mssql+pyodbc:///?odbc_connect={params}", fast_executemany=True)
    sink = engine
    print(f"target: SQL Server {os.environ['DB_SERVER']}/{os.environ['DB_NAME']}")
else:
    sink = sqlite3.connect(OUT_DB)
    print(f"target: SQLite {OUT_DB}")

CHUNK = 50000
main.head(0).to_sql(OUT_TABLE, sink, if_exists="replace", index=False)
for i in range(0, len(main), CHUNK):
    main.iloc[i:i + CHUNK].to_sql(OUT_TABLE, sink, if_exists="append", index=False)
    print(f"  loaded {min(i + CHUNK, len(main)):,}/{len(main):,}")

print(f"\nDone. {len(main):,} rows written to {OUT_TABLE}.")
print("Columns:", ", ".join(main.columns))
