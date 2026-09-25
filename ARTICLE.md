# Turning many messy sources into one number everyone trusts

*A walk through a sales, budget and forecast consolidation pipeline. All data in
the project is synthetic; this is about the approach, not any company's numbers.*

## The problem worth solving

In a lot of businesses the hardest part of analytics is not the analysis. It is
that the same question gets different answers depending on who you ask and which
file they happen to open.

Picture a group with a few manufacturing entities and a couple of trading
entities, plus external distributors who sell the products on. Every one of them
produces sales data, and every one of them produces it differently:

- the ERP holds actuals, budgets and open orders, split across each legal entity
- each distributor sends a sell-out file in its own shape: different column
  names, extra header rows, months as numbers in one file and as names in
  another, cartons here and pieces there
- master files map each distributor's own item and customer codes back to the
  group's hierarchy

Before any of this is consolidated, building one group view means someone opening
a dozen files and copy-pasting, every cycle. It is slow, it drifts, and two
people doing it get two answers.

The pipeline in this project replaces that with a single repeatable run.

## The shape of the solution

The whole thing is one idea repeated: **pull each source into a common schema,
stack them, enrich them, then load the result.** Nothing clever, but the details
are where the real work lives.

```
distributor files (Excel, Excel-with-offset, CSV)  ─┐
ERP actuals (invoices netted vs credit memos)       ─┤
ERP budget + last estimate                          ─┼─> unified table ─> cost + value chain ─> database
ERP forecast (open orders: Plan/Tentative/Firm)     ─┘
```

## The parts that matter

### 1. Harmonising sources that refuse to match

Three distributor feeds, three shapes. One is a clean Excel sheet with a numeric
month. One is an Excel export with two junk rows on top before the real header.
One is a CSV with entirely different column names and the period written as
`YYYYMM`. Each gets read on its own terms, then projected onto a single schema so
that from then on the rest of the pipeline does not care where a row came from.

The lesson that keeps repeating: do not try to force sources to look alike at the
door. Read each one honestly, then translate. A small translation function per
source is far easier to maintain than one giant reader full of special cases.

### 2. Netting invoices against credit memos

Actual sales are not just invoices. Returns and corrections come through as
credit memos, and if you ignore them your revenue is overstated. The ERP query
pulls invoice lines as positive quantities and values, pulls credit-memo lines as
negative, unions them, and groups so each item, customer and month comes out net.
Across every entity in one pass.

### 3. Three views of the future

Actuals answer what happened. A commercial team also needs what is planned, so
the pipeline pulls three more things and tags each clearly:

- **Budget** and **Last Estimate**, from the ERP's budget entries
- **Forecast**, from open sales orders, carrying the order status so you can tell
  a firm commitment from a tentative plan

Everything lands in the same table with an `ACT/FC` flag, so a report can compare
actual against budget against forecast without another join.

### 4. Cost, derived three ways

Margin is only as good as the cost behind it, and cost is rarely clean. The
pipeline derives it three ways and falls back gracefully:

- a **rolling average** output cost from the ledger
- a **standard cost** from the price list
- when a month has no cost at all, the **nearest available month** for that item

That fallback matters more than it sounds. A single missing cost should not blank
out a whole row's margin, so the pipeline reaches for the closest real number
instead of giving up.

### 5. The transfer-price value chain

This is the part I find most interesting. In a group, a product can be made by
one entity, sold to a central hub, sold on to a distributor, and finally sold to a
shopper. Each hop has its own price, and margin sits at each hop.

The pipeline reconstructs that chain and computes gross profit at every stage:

```
factory cost ─> factory→hub price ─> hub→distributor price ─> retail shelf price
        └ factory margin ┘     └ hub margin ┘         └ retail margin ┘
```

Seeing margin per stage, rather than only at the group total, is what lets a
business tell where money is actually made and where it quietly leaks. A healthy
group total can hide an unhealthy stage.

### 6. A weighted retail price

For the retail end, a single item can sit at different shelf prices across
different store groups. The pipeline computes a weighted retail selling price per
item, store group and month, so the retail margin is grounded in what shoppers
actually paid, not a list price.

### 7. Loading, without drama

The final wide table is cleaned (rounded, infinities removed) and loaded in
chunks. It defaults to a local SQLite database so anyone can run the whole
project with no setup, and switches to SQL Server just by setting a few
environment variables. Credentials never live in the code.

## What I would add next

- **Pocket margin** net of promotions. List margin flatters reality when a
  product launches on heavy deals.
- **Automated data-quality checks** at ingestion, so a distributor sending a
  malformed file fails loudly instead of silently dropping rows.
- **Incremental loads** instead of a full rebuild, once the volume justifies it.

## Running it yourself

Everything is synthetic and self-contained:

```bash
pip install -r requirements.txt
python generate_sample_data.py
python build_consolidated_sales.py
```

The generator invents a fictional group and a small SQLite ERP; the pipeline
produces the consolidated table. Code and full details are in the repository:
**[multi-source-sales-consolidation](https://github.com/basil-v/multi-source-sales-consolidation)**.

## The takeaway

None of the individual steps are exotic. The value is in doing them in one place,
the same way every time, so that the group stops arguing about whose number is
right and starts talking about what to do. Agreeing on the number is half the job.
