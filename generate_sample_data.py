"""
generate_sample_data.py
-----------------------
Creates a fully synthetic dataset for the multi-source sales consolidation demo.

Nothing here is real. It invents a fictional confectionery group ("Contoso
Confectionery Group") with five internal entities and three external distributor
feeds, plus a small SQLite database that stands in for an ERP. Running this once
produces every input the pipeline needs, so the project is self-contained and
runnable by anyone who clones it - no real database, no credentials.

Outputs:
  inputs/master_classification.xlsx   (SKU + customer master, messy headers on purpose)
  inputs/dist_a/DIST_A_*.xlsx         (distributor feed, format 1)
  inputs/dist_b/DIST_B_*.xlsx         (distributor feed, format 2, offset headers)
  inputs/dist_c/DIST_C_*.csv          (distributor feed, format 3)
  db/erp.db                           (SQLite ERP stand-in)
"""

import os
import sqlite3
import calendar
import numpy as np
import pandas as pd
from openpyxl import Workbook

rng = np.random.default_rng(42)

BASE = os.path.dirname(os.path.abspath(__file__))
INPUTS = os.path.join(BASE, "inputs")
DB_DIR = os.path.join(BASE, "db")
for d in [INPUTS, os.path.join(INPUTS, "dist_a"), os.path.join(INPUTS, "dist_b"),
          os.path.join(INPUTS, "dist_c"), DB_DIR]:
    os.makedirs(d, exist_ok=True)

# ------------------------------------------------------------------
# Reference dimensions (all invented)
# ------------------------------------------------------------------
# Five internal entities. Roles mirror a real group: factories make goods,
# a central hub buys from factories and sells onward, trading resells locally.
# FAC2 carries a 0.98 adjustment factor to show entity-specific economics.
ENTITIES = [
    ("FAC1", "factory",        1.00),
    ("FAC2", "factory",        0.98),
    ("MFG1", "manufacturing",  1.00),
    ("HUB1", "trading-hub",    1.00),
    ("TRD1", "trading",        1.00),
]

BRANDS   = ["Aurora", "Bolt", "Cocoa Cloud", "Dune"]
CATEGORY = ["CHOCOLATE", "SNACKS", "BAKERY"]
SEGMENTS = ["Premium", "Mainstream", "Value"]
BASES    = ["Wafer", "Biscuit", "Nut", "Cereal"]
CHANNELS = ["Modern Trade", "Traditional Trade", "HORECA", "E-Commerce", "Duty Free"]
REGIONS  = ["UAE", "KSA", "Kuwait", "Oman"]

MONTHS = list(range(1, 13))
YEARS  = [2024, 2025]

# ------------------------------------------------------------------
# SKU master  (internal factory code <-> distributor code mapping)
# ------------------------------------------------------------------
N_SKU = 24
skus = []
for i in range(N_SKU):
    fac_code = f"FG-{1000+i}"
    dist_code = f"D{2000+i}"          # code the distributors use for the same item
    cat = CATEGORY[i % len(CATEGORY)]
    weight = int(rng.choice([30, 45, 90, 150, 200]))
    skus.append({
        "Factory Code": fac_code,
        "Distributor Code": dist_code,
        "Barcode": f"62910{i:05d}",
        "SKU Description Manual": f"{BRANDS[i % len(BRANDS)]} {BASES[i % len(BASES)]} {weight}g",
        "SKU Description Systematic": f"{BRANDS[i % len(BRANDS)]}-{weight}",
        "Weight": weight,
        "Weight UOM": "PCS",
        "Brand Principal": "Contoso",
        "Units per SKU": 1,
        "Packaging Type": "Flowpack",
        "SKUs per Carton": int(rng.choice([12, 24, 48])),
        "Units per Carton": int(rng.choice([12, 24, 48])),
        "SKU Type": "FG",
        "Brand": BRANDS[i % len(BRANDS)],
        "Category": cat,
        "Segment": SEGMENTS[i % len(SEGMENTS)],
        "Base": BASES[i % len(BASES)],
        "Filling": rng.choice(["Caramel", "Hazelnut", "None"]),
        "Coating": rng.choice(["Milk", "Dark", "None"]),
        "Weight Range": "S" if weight < 60 else ("M" if weight < 150 else "L"),
        "Sub-Brand": f"{BRANDS[i % len(BRANDS)]} Line",
        "Item Description Systematic": f"{BRANDS[i % len(BRANDS)]}-{BASES[i % len(BASES)]}-{weight}",
        "Business Unit": cat,
        # economics used only by the generator (not written to the master sheet)
        "_base_cost": round(float(rng.uniform(1.5, 6.0)), 2),
    })
item_master = pd.DataFrame(skus)

# ------------------------------------------------------------------
# Customer master  (group code <-> distributor code, channel, geography)
# ------------------------------------------------------------------
N_CUST = 18
custs = []
for i in range(N_CUST):
    grp_code = f"GRP-C-{100+i}"
    dist_ccode = f"DC{500+i}"
    custs.append({
        "Group Customer Code": grp_code,
        "Distributor Customer Code": dist_ccode,
        "Customer Full Name": f"Customer {i+1} LLC",
        "Customer Name": f"Customer {i+1}",
        "Channel": CHANNELS[i % len(CHANNELS)],
        "Sub-Channel": rng.choice(["Key Account", "Wholesale", "Retail"]),
        "Customer Group": rng.choice(["Carousel", "Marketway", "Freshline", "Cornerstore"]),
        "Branch/Country": rng.choice(REGIONS),
        "City": rng.choice(["Dubai", "Riyadh", "Kuwait City", "Muscat"]),
        "Region": REGIONS[i % len(REGIONS)],
        "Salesman_Group": rng.choice(["North", "South", "Export"]),
        "Salesman": f"Rep {chr(65 + (i % 6))}",
    })
customer_master = pd.DataFrame(custs)

# Reserve special internal customer codes used in the value chain
HUB_AS_CUSTOMER   = "GRP-C-100"   # factories sell to the hub under this code
DIST_AS_CUSTOMER  = "GRP-C-101"   # hub sells to the external distributor under this code
TRD_AS_CUSTOMER   = "GRP-C-102"   # hub sells to trading arm under this code


# ------------------------------------------------------------------
# Helper: write a dataframe to an .xlsx sheet with an intentional
# title row + blank first column, so the pipeline has to use
# skiprows / usecols - mirroring real-world messy exports.
# ------------------------------------------------------------------
def write_messy_sheet(path, sheet, df, title):
    wb = Workbook()
    ws = wb.active
    ws.title = sheet
    ws["B1"] = title                      # row 1 = title (skiprows=1)
    for c, col in enumerate(df.columns, start=2):   # headers from column B
        ws.cell(row=2, column=c, value=col)
    for r, (_, row) in enumerate(df.iterrows(), start=3):
        for c, col in enumerate(df.columns, start=2):
            ws.cell(row=r, column=c, value=row[col])
    wb.save(path)


# Master workbook: two sheets, both with the messy header layout
master_path = os.path.join(INPUTS, "master_classification.xlsx")
item_cols = [c for c in item_master.columns if not c.startswith("_")]
write_messy_sheet(master_path, "SKUS CLASSIFICATION", item_master[item_cols],
                  "SKU CLASSIFICATION - CONFIDENTIAL (SAMPLE)")
# second sheet appended
from openpyxl import load_workbook
wb = load_workbook(master_path)
ws = wb.create_sheet("CHANNEL CLASSIFICATION")
ws["B1"] = "CHANNEL CLASSIFICATION (SAMPLE)"
for c, col in enumerate(customer_master.columns, start=2):
    ws.cell(row=2, column=c, value=col)
for r, (_, row) in enumerate(customer_master.iterrows(), start=3):
    for c, col in enumerate(customer_master.columns, start=2):
        ws.cell(row=r, column=c, value=row[col])
wb.save(master_path)
print(f"wrote {master_path}")


# ------------------------------------------------------------------
# Distributor feeds - three different real-world shapes
# ------------------------------------------------------------------
def sku_sample(n):
    return item_master.sample(n=n, random_state=int(rng.integers(0, 1e6)), replace=True)

# ---- DIST_A : clean-ish Excel, numeric month, cartons + value ----
def make_dist_a(year, budget=False):
    rows = []
    for _ in range(180):
        s = item_master.iloc[int(rng.integers(0, N_SKU))]
        c = customer_master.iloc[int(rng.integers(0, N_CUST))]
        qty_cs = int(rng.integers(5, 120))
        cptt = round(float(rng.uniform(20, 80)), 2)
        rows.append({
            "Customer ID": c["Distributor Customer Code"],
            "Item": s["Distributor Code"],
            "Month": int(rng.choice(MONTHS)),
            "Year": year,
            "QTY Cs": qty_cs,
            "Sales Value": round(qty_cs * cptt, 2),
        })
    df = pd.DataFrame(rows)
    tag = "BUDGET" if budget else "ACTUAL"
    p = os.path.join(INPUTS, "dist_a", f"DIST_A_{tag}_{year}.xlsx")
    with pd.ExcelWriter(p) as xl:
        df.to_excel(xl, sheet_name="DATA", index=False)
    print(f"wrote {p}")

make_dist_a(2024)
make_dist_a(2025)
make_dist_a(2026, budget=True)

# ---- DIST_B : Excel with 2 title rows to skip, month name, cartons ----
def make_dist_b(year, budget=False):
    rows = []
    for _ in range(150):
        s = item_master.iloc[int(rng.integers(0, N_SKU))]
        c = customer_master.iloc[int(rng.integers(0, N_CUST))]
        qty = int(rng.integers(3, 90))
        sales = round(qty * float(rng.uniform(60, 180)), 2)
        rows.append({
            "MONTHNAME": calendar.month_abbr[int(rng.choice(MONTHS))].upper(),
            "YEAR": year,
            "CUSTOMER CODE": c["Distributor Customer Code"],
            "CUSTOMER NAME": c["Customer Name"],
            "PRICEGROUP": "PARTNER-B",
            "PLANT": "P1",
            "City": c["City"],
            "ITEM CODE": s["Distributor Code"],
            "ITEM NAME": s["SKU Description Manual"],
            "QTY": qty,
            "SALES": sales,
        })
    df = pd.DataFrame(rows)
    tag = "BUDGET" if budget else "ACTUAL"
    p = os.path.join(INPUTS, "dist_b", f"DIST_B_{tag}_{year}.xlsx")
    # two junk rows on top -> pipeline reads with skiprows=2
    wb = Workbook(); ws = wb.active; ws.title = "Sheet1"
    ws["A1"] = "Partner B sell-out extract (sample)"
    ws["A2"] = "Generated automatically"
    for c, col in enumerate(df.columns, start=1):
        ws.cell(row=3, column=c, value=col)
    for r, row in enumerate(df.itertuples(index=False), start=4):
        for c, val in enumerate(row, start=1):
            ws.cell(row=r, column=c, value=val)
    wb.save(p)
    print(f"wrote {p}")

make_dist_b(2025)
make_dist_b(2026, budget=True)

# ---- DIST_C : CSV, YYYYMM period, different column names ----
def make_dist_c(year, budget=False):
    rows = []
    for _ in range(140):
        s = item_master.iloc[int(rng.integers(0, N_SKU))]
        c = customer_master.iloc[int(rng.integers(0, N_CUST))]
        qty = int(rng.integers(4, 100))
        val = round(qty * float(rng.uniform(50, 160)), 2)
        rows.append({
            "ProductCode": s["Distributor Code"],
            "BrandName": s["Brand"],
            "ChannelName": c["Channel"],
            "CustomerCode": c["Distributor Customer Code"],
            "CustomerName": c["Customer Name"],
            "ProductName": s["SKU Description Manual"],
            "MonthYear": f"{year}{int(rng.choice(MONTHS)):02d}",
            "SaleYear": year,
            "NetSalesQty": qty,
            "NetsalesValue": val,
        })
    df = pd.DataFrame(rows)
    tag = "BUDGET" if budget else "ACTUAL"
    p = os.path.join(INPUTS, "dist_c", f"DIST_C_{tag}_{year}.csv")
    df.to_csv(p, index=False)
    print(f"wrote {p}")

make_dist_c(2025)
make_dist_c(2026, budget=True)


# ------------------------------------------------------------------
# ERP stand-in  (SQLite)   -  generic table names, entity column.
# In production these were separate Business Central company
# databases unioned together; here one table + an `entity` column
# keeps the same logic readable.
# ------------------------------------------------------------------
db_path = os.path.join(DB_DIR, "erp.db")
if os.path.exists(db_path):
    os.remove(db_path)
conn = sqlite3.connect(db_path)

def dt(y, m):
    return f"{y}-{m:02d}-15"

# ---- item + uom + customer + cost ----
item_rows, uom_rows, cust_rows = [], [], []
for _, s in item_master.iterrows():
    for ent, role, factor in ENTITIES:
        unit_cost = round(s["_base_cost"] * factor, 3)
        item_rows.append((ent, s["Factory Code"], s["SKU Description Manual"],
                          "PCS", "FG", "Contoso", unit_cost))
        uom_rows.append((ent, s["Factory Code"], "PCS", 1))
        uom_rows.append((ent, s["Factory Code"], "CTN", int(s["SKUs per Carton"])))
for _, c in customer_master.iterrows():
    for ent, _, _ in ENTITIES:
        cust_rows.append((ent, c["Group Customer Code"], c["Customer Group"]))

pd.DataFrame(item_rows, columns=["entity", "item_code", "description", "base_uom",
             "inventory_posting_group", "principal", "unit_cost"]).to_sql("item", conn, index=False)
pd.DataFrame(uom_rows, columns=["entity", "item_code", "uom_code", "qty_per_uom"]).to_sql("item_uom", conn, index=False)
pd.DataFrame(cust_rows, columns=["entity", "no_", "customer_price_group"]).to_sql("customer", conn, index=False)

# ---- sales invoice + credit memo (actuals) ----
inv_rows, crm_rows = [], []
doc = 0
for _, s in item_master.iterrows():
    for ent, role, factor in ENTITIES:
        for y in YEARS:
            for m in MONTHS:
                if rng.random() < 0.35:      # sparse
                    continue
                doc += 1
                qty = int(rng.integers(50, 800))
                ptt = s["_base_cost"] * factor * rng.uniform(1.3, 2.2)
                amt = round(qty * ptt, 2)
                cust = customer_master.iloc[int(rng.integers(0, N_CUST))]["Group Customer Code"]
                inv_rows.append((ent, f"SI{doc:06d}", cust, s["Factory Code"],
                                 qty, amt, 1.0, dt(y, m), "STD"))
                if rng.random() < 0.10:      # occasional return
                    rq = int(qty * rng.uniform(0.02, 0.1))
                    crm_rows.append((ent, f"CM{doc:06d}", cust, s["Factory Code"],
                                     rq, round(rq * ptt, 2), 1.0, dt(y, m)))

pd.DataFrame(inv_rows, columns=["entity", "doc_no", "customer_code", "item_code",
             "qty_base", "amount", "currency_factor", "posting_date", "posting_group"]
            ).to_sql("sales_invoice", conn, index=False)
pd.DataFrame(crm_rows, columns=["entity", "doc_no", "customer_code", "item_code",
             "qty_base", "amount", "currency_factor", "posting_date"]
            ).to_sql("sales_creditmemo", conn, index=False)

# ---- item budget entry (budget + last estimate) ----
bud_rows = []
for budget_name in ["BUDGET_V1", "BUDGET_V2"]:      # V1=budget, V2=last estimate
    for _, s in item_master.iterrows():
        for ent, role, factor in ENTITIES:
            for m in MONTHS:
                if rng.random() < 0.5:
                    continue
                qty = int(rng.integers(40, 700))
                price = round(s["_base_cost"] * factor * rng.uniform(1.4, 2.3), 3)
                cust = customer_master.iloc[int(rng.integers(0, N_CUST))]["Group Customer Code"]
                bud_rows.append((ent, budget_name, s["Factory Code"], cust, qty,
                                 round(qty * price, 2), price, dt(2026, m), "CTN", "Rep A"))
pd.DataFrame(bud_rows, columns=["entity", "budget_name", "item_code", "source_no",
             "qty", "sales_amount", "unit_price", "date", "uom", "salesperson"]
            ).to_sql("item_budget_entry", conn, index=False)

# ---- open sales orders (forecast, with Plan/Tentative/Firm status) ----
sl_rows = []
for _, s in item_master.iterrows():
    for ent, role, factor in ENTITIES:
        for m in MONTHS:
            if rng.random() < 0.6:
                continue
            qty = int(rng.integers(20, 400))
            ptt = s["_base_cost"] * factor * rng.uniform(1.3, 2.1)
            cust = customer_master.iloc[int(rng.integers(0, N_CUST))]["Group Customer Code"]
            sl_rows.append((ent, 1, cust, s["Factory Code"], qty, round(qty * ptt, 2),
                            dt(2026, m), int(rng.choice([0, 1, 2]))))
pd.DataFrame(sl_rows, columns=["entity", "doc_type", "customer_code", "item_code",
             "qty_base", "amount", "requested_delivery_date", "salesperson_status"]
            ).to_sql("sales_line", conn, index=False)

# ---- item ledger + value entry (actual output cost) ----
ile_rows, ve_rows = [], []
entry = 0
for _, s in item_master.iterrows():
    for ent, role, factor in ENTITIES:
        if role not in ("factory", "manufacturing"):
            continue
        for y in YEARS:
            for m in MONTHS:
                if rng.random() < 0.5:
                    continue
                entry += 1
                qty = int(rng.integers(100, 1000))
                cost = round(qty * s["_base_cost"] * factor * rng.uniform(0.95, 1.05), 2)
                ile_rows.append((ent, entry, s["Factory Code"], 6, dt(y, m), qty))  # type 6 = output
                ve_rows.append((ent, entry, s["Factory Code"], cost))
pd.DataFrame(ile_rows, columns=["entity", "entry_no", "item_code", "entry_type",
             "posting_date", "quantity"]).to_sql("item_ledger_entry", conn, index=False)
pd.DataFrame(ve_rows, columns=["entity", "item_ledger_entry_no", "item_code", "cost_amount"]
            ).to_sql("value_entry", conn, index=False)

# ---- price list (transfer prices + standard cost + customer prices) ----
# source_group: 10/11/12 = customer/price-group/campaign prices (kept generic)
pl_rows = []
for _, s in item_master.iterrows():
    base = s["_base_cost"]
    # factory -> hub transfer price
    for ent, role, factor in ENTITIES:
        if role in ("factory", "manufacturing"):
            pl_rows.append((ent, s["Factory Code"], HUB_AS_CUSTOMER, "PCS",
                            round(base * factor * 1.25, 3), "AED", None, "STD", 11))
    # hub -> distributor and hub -> trading
    pl_rows.append(("HUB1", s["Factory Code"], DIST_AS_CUSTOMER, "PCS",
                    round(base * 1.55, 3), "AED", None, "STD", 11))
    pl_rows.append(("HUB1", s["Factory Code"], TRD_AS_CUSTOMER, "PCS",
                    round(base * 1.50, 3), "AED", None, "STD", 11))
    # a generic customer price on each entity (standard price list)
    for ent, role, factor in ENTITIES:
        pl_rows.append((ent, s["Factory Code"], "", "PCS",
                        round(base * factor * 1.6, 3), "AED", None, "STD", 10))
pd.DataFrame(pl_rows, columns=["entity", "product_no", "source_no", "uom_code",
             "unit_price", "currency_code", "ending_date", "price_list_code", "source_group"]
            ).to_sql("price_list_line", conn, index=False)

# ---- currency exchange rate ----
pd.DataFrame([("HUB1", "USD", 3.67, "2024-01-01"), ("HUB1", "AED", 1.0, "2024-01-01")],
             columns=["entity", "currency_code", "rate", "starting_date"]
            ).to_sql("currency_exchange_rate", conn, index=False)

# ---- availability / RSP feed (retail shelf prices by customer group) ----
rsp_rows = []
groups = ["Carousel", "Marketway", "Freshline", "Cornerstore"]
for _, s in item_master.iterrows():
    for g in groups:
        for y in YEARS:
            for m in MONTHS:
                if rng.random() < 0.6:
                    continue
                price = round(s["_base_cost"] * rng.uniform(2.0, 3.2), 2)
                rsp_rows.append((s["Factory Code"], g, dt(y, m), price))
pd.DataFrame(rsp_rows, columns=["item_code", "group_name", "date", "price"]
            ).to_sql("availability_report_data", conn, index=False)

conn.commit()
conn.close()
print(f"wrote {db_path}")
print("\nSample data generated. Run:  python build_consolidated_sales.py")
