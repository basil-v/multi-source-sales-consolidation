# Multi-Source Sales Consolidation Pipeline

A Python pipeline that unifies sales, budget and forecast data from many
differently-shaped sources into one wide, analysis-ready table and loads it to a
database. It is built to mirror a real problem in a multi-entity FMCG group:
the same commercial question was getting different answers depending on which
file someone opened.

Built with Python (pandas), SQL and SQLite. One run consolidates ~5,000 rows across Actual, Budget, Forecast and Last Estimate, and computes gross profit at every stage of the value chain.

> **Read the write-up:** [Turning many messy sources into one number everyone trusts](ARTICLE.md) — the story and design decisions behind this pipeline.

> **Note on data:** every name, code and number in this repository is synthetic.
> `generate_sample_data.py` invents a fictional confectionery group and a small
> SQLite database that stands in for an ERP, so the whole project runs locally
> with no external database and no credentials. It demonstrates the engineering
> approach, not any real company's data.

## The problem

A group with several manufacturing and trading entities, plus external
distributors, produced sales data in incompatible shapes:

- an ERP holding actuals, budgets and open orders across every entity
- distributor sell-out feeds, each with its own layout (different column names,
  header rows, month formats, cartons vs pieces)
- master files mapping distributor codes and customer codes to the group's own
  item and customer hierarchy

Before consolidation, building a single group view meant manual copy-paste every
cycle. This pipeline replaces that with one repeatable run.

## What it does

1. **Ingests three distributor formats** — a clean Excel feed, an Excel feed with
   offset header rows, and a CSV with different column names and a `YYYYMM`
   period — and harmonises them to one schema.
2. **Reads the ERP** and:
   - nets invoices against credit memos and unions actuals across entities
   - pulls Budget and Last Estimate from budget entries
   - pulls an open-order Forecast with `Plan / Tentative / Firm` status
3. **Classifies everything** against SKU and customer masters (brand, category,
   channel, region, salesperson).
4. **Derives cost three ways** — a rolling average output cost, a standard cost
   from the price list, and a nearest-month fallback when a month is missing.
5. **Builds an intercompany transfer-price value chain** and computes gross
   profit at each hop:

   ```
   factory cost -> factory→hub TP -> hub→distributor TP -> retail RSP
        └ Factory_GP ┘     └ Hub_GP ┘        └ Retail_GP ┘
   ```

6. **Computes a weighted retail selling price (RSP)** per item, customer group
   and month.
7. **Loads the result to SQL in chunks** — SQLite by default, or SQL Server if
   you set the `DB_*` environment variables.

## Techniques shown

- pandas multi-source ETL and schema harmonisation
- SQL with `UNION ALL`, invoice/credit-memo netting, `GROUP BY … HAVING`,
  currency conversion and window-style dedup
- dimensional joins against master data
- a multi-stage margin / transfer-price model
- chunked database loading with a swappable SQLite / SQL Server sink
- secrets kept out of code (environment variables only)

## Run it

```bash
pip install -r requirements.txt
python generate_sample_data.py      # creates inputs/ and db/erp.db
python build_consolidated_sales.py  # writes db/output.db  (table SALES_WITH_BUDGET)
```

Inspect the result:

```python
import sqlite3, pandas as pd
df = pd.read_sql("SELECT * FROM SALES_WITH_BUDGET", sqlite3.connect("db/output.db"))
df.groupby("ACT/FC")["Amount"].sum()
```

To target SQL Server instead of SQLite, copy `.env.example` to `.env`, fill in
the values, export them, and re-run the pipeline.

## Layout

```
generate_sample_data.py     # builds synthetic inputs + SQLite ERP stand-in
build_consolidated_sales.py # the pipeline
inputs/                     # generated distributor feeds + masters (gitignored)
db/                         # generated SQLite databases (gitignored)
requirements.txt
.env.example
```
