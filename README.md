# Inventory Control Tower (Google Sheets + Apps Script)

A spreadsheet-native control tower for inventory, landed cost and cashflow. It is built for a brand with a small SKU count, high unit volume, and stock held at two international 3PL locations: one in the EU (`WH-DE`) and one in Asia (`WH-CN`).

One Apps Script file builds the whole workbook: 15 tabs, about 47,000 live formulas, and data validation plus color coding on every input. Once it is built, every computed number updates on its own when a sale, payment or shipment is entered. No script needs to run again.

This is an anonymized code sample. All names, SKUs and numbers in the demo dataset are synthetic.

## Files

| File | Purpose |
|---|---|
| `control-tower.gs` | The full build script plus the synthetic demo dataset |
| `appsscript.json` | Minimal V8 manifest |
| `README.md` | This document |

## How to run

1. Create a blank Google Sheet.
2. Open **Extensions > Apps Script**.
3. Replace the contents of `Code.gs` with `control-tower.gs` and save.
4. Optional: in **Project Settings**, tick "Show appsscript.json" and paste in `appsscript.json`.
5. Pick `setupAll` in the function dropdown and click **Run**. Grant the spreadsheet permission when prompted (the script only touches the sheet it is bound to).
6. Return to the sheet. All tabs are now built and filled with the synthetic demo data.

Other entry points:

- `setupStructureOnly()` builds an empty model with no demo data.
- `seedDemoData()` loads the demo data into an existing empty structure.
- `protectAllTabsExceptManualInput()` (optional) locks every computed tab to the file owner.
- Each `build*()` function rebuilds a single tab. Use these only during development: rebuilding one tab breaks other tabs' references to it, so once the model holds real data, rebuild the whole chain in `setupAll()` order.

Configuration lives at the top of the script:

- `LOOKBACK_DAYS` / `START_DATE`: the day grid starts at "today minus 60 days" by default, so the demo always has 60 days of history.
- `HORIZON_DAYS`: 360 by default.
- `CASHFLOW_LEAD_DAYS`: Cashflow starts 90 days earlier than the stock grid, because factory deposits are paid before any stock exists.
- `SEED_DEMO_DATA_ON_SETUP`: set this to `false` to skip the demo data.

If a run hits the Apps Script 6-minute execution limit on a slow account, run the `build*()` functions one at a time, in the order listed in `buildAllTabs_()`, then run `seedDemoData()`.

## Tab-by-tab data flow

```
SKU Master, Vendors, Inputs        reference data, dropdown sources
        |
Transactions  ------------------>  payment Actual Dates (matched by ref + amount)
        |                                   |
Purchase Orders  <------------------------- +
        |  (PO# + SKU: factory cost/unit)
Shipments  <------------------------------- +
        |  (one SKU per row, landed-cost components)
        v
Inventory Ledger (per-batch FIFO)  <---  Sales Plan per warehouse (Actual/day)
        |                                   |
        +--> Daily COGS  <------------------+
        +--> Purchase Orders: Actual Landed Cost/unit, PPV
StockModel per warehouse  <--- Shipments (IN, transit), Purchase Orders (production), Sales Plan
Landed Cost Summary       <--- Purchase Orders + Shipments
Cashflow                  <--- PO tranches, Shipment tranches, Sales Plan revenue x rates,
                               Fulfillment & Carrier Actuals
```

| Tab | Type | What it holds |
|---|---|---|
| SKU Master | input | SKU codes and carton specs. Source for every SKU dropdown. |
| Vendors | input | Factories, forwarders, 3PLs and customs brokers, with dropdown-typed fields. |
| Inputs | input | Standing config values, such as the PPV trigger threshold. Formulas read these cells instead of hardcoded numbers. |
| Transactions | input | Payment ledger: amount, currency and FX, counterparty, category, warehouse, and a free-text `PO#/Shipment Ref`. Status (paid / scheduled / overdue) is derived from the dates. |
| Purchase Orders | input + formulas | One row per PO x SKU. Holds the planned qty per warehouse, EXW price and fees, and up to 3 payment tranches (planned vs actual date, days late, paid status). Also holds the payment sum check, actual landed cost per warehouse, and PPV. |
| Shipments | input + formulas | One row per shipment, with one SKU per row. Tracks the ship method, planned and actual warehouse, dates, 2 freight tranches plus 1 customs tranche, customs status, and invoiced cost components. Computes gross and net cost/unit, implied duty rate, projected ETA, and projected cost. |
| Sales Plan (per warehouse) | input + formulas | A day grid with Plan/day and Actual/day per SKU, plus plan and actual revenue. This is the only place a demand assumption lives. |
| Fulfillment & Carrier Actuals | input | Real 3PL and carrier invoices per date and warehouse. They override the rate-based estimates in Cashflow. |
| Inventory Ledger | computed | FIFO batch detail, with one row per shipment (1:1 with Shipments) and per-SKU qty received, remaining and value. A summary shows on-hand qty and value, weighted cost, and oversold qty. |
| StockModel (per warehouse) | computed | A day grid per SKU: inbound, batch IDs, sold/day, stock, days in transit, days to production-ready, days of stock and pipeline days, plus a status color. Sidebars show the 7-day run rate vs plan and today's coverage. |
| Daily COGS | computed | FIFO-priced COGS and opening inventory value per day, SKU and warehouse, with totals and an unpriced (oversold) flag. |
| Landed Cost Summary | computed | One row per landed shipment: product cost, logistics cost, total, and total net of refundable VAT. |
| Cashflow | computed | Daily cash out by category (purchases per SKU, logistics, fulfillment, carrier), with a running total. |

Yellow cells are manual input. Every other cell is a live formula.

## Key logic

**Warehouses are never blended.** Each warehouse has its own Sales Plan and StockModel tab, and its own FIFO layers in the Ledger. Days of supply are always computed per warehouse. The two locations serve different markets with different demand, lead times and logistics costs, so a blended figure would hide a stockout in one location behind surplus in the other. PPV is split per warehouse for the same reason.

**Per-batch FIFO landed cost, as pure formulas.** Each landed shipment is one batch. Its landed cost/unit is the PO's full factory cost/unit (EXW plus per-unit fees) plus that shipment's own net logistics cost/unit. FIFO works without sorting any rows. Each batch's cumulative received qty is ranked by arrival date, with same-day ties broken by row order. Its remaining qty is then:

`MAX(0, MIN(batch qty, cumulative received - total sold))`

so the oldest batches deplete first. Daily COGS repeats the same clamp for the start and end of every historical day across all batches at once. It uses nested `IF` inside `SUMPRODUCT`, because `MIN`/`MAX` do not work element-wise there. Each unit sold is therefore priced at the real cost of the batch it came from, and a later wave costs what that wave actually cost.

**One SKU per shipment row.** A mixed container is entered as several rows, one per SKU, each carrying its own share of freight and customs. Batch cost, the Ledger, and the Landed Cost Summary therefore stay exact per SKU, with no estimated blends. One customs payment that covers a whole container lists every row's Shipment ID in its Transactions ref.

**Date-gated stock arrival.**
- A shipment becomes a FIFO batch only once its actual arrival date is filled in. Before that, it shows in StockModel as expected inbound on its ETA.
- Stock landing on day *d* is available from day *d + 1*.
- Opening inventory in Daily COGS only counts batches that had already landed.
- COGS consumption is deliberately not gated. A sale made before stock arrived is a backorder, priced at the cost of the batch that later fulfils it.
- Transit days are tested against each shipment's real departure-to-arrival window for each row's own date, not against "has it arrived yet today". Historical rows therefore keep their true in-transit values after a shipment lands.

**Unknown is never shown as zero.**
- Gross cost/unit stays blank until at least one cost component is invoiced.
- Transit and production days return blank, not 0, when an ETA or ready date is missing.
- Units that no batch can cover are flagged as "unpriced" instead of being priced at a guessed rate.

**Cashflow from payment schedules.**
- **Purchases:** every PO tranche lands on its actual date when paid and on its planned date otherwise. Inspection cost rides on tranche 2.
- **Logistics:** each shipment has 3 independent tranches (freight 1, freight 2, customs), each placed the same way. When the real customs amount is unknown, a manual customs estimate is used. When nothing at all is known yet, the ship method's historical cost/unit x qty is placed on the planned departure date.
- **Fulfillment and carrier costs:** revenue x a rate per 1000, per warehouse. A real invoice for a past date overrides the estimate for that warehouse only.

**Payments reconcile themselves.** Payment actual dates on Purchase Orders and Shipments are matched from Transactions:
- The match uses an exact token in a comma-separated ref list, so "PO-1" never matches "PO-10".
- When several payments carry the same ref, the one closest in amount wins.
- Any candidate more than 10% off the expected amount is rejected.

The paid flag, days late and payment status follow from the match.

**Columns are resolved by header name.** No formula uses a hand-counted column letter. Every reference goes through a name-to-letter lookup, so adding or reordering a column cannot silently mis-map a formula.

## Demo dataset

`seedDemoData()` writes a small, obviously synthetic scenario. All dates are relative to the run date.

- **2 purchase orders.** PO-001 covers 3 SKUs and is fully shipped. PO-002 covers SKU-1 and is still in production, which drives the production-days column.
- **8 shipment rows** across Sea, Rail, Air and Local, covering both warehouses. One Air shipment is still in transit.
- **60 days of noisy actual sales** plus a flat forward plan for each SKU and warehouse.
- **17 transactions** that mark most tranches as paid. One is paid late on purpose.
- **Synthetic rates and invoices:** fulfillment and carrier rates, plus one invoice override per warehouse.

Prices are round (2.00 / 1.50 / 0.80 per unit) and in EUR. The model adds factory cost and logistics cost per unit, so it assumes one reporting currency. Transactions in other currencies convert through their FX rate column.
