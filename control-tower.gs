/**
 * Inventory Control Tower (Google Sheets + Apps Script)
 *
 * A spreadsheet-native inventory, landed-cost and cashflow model for a small-SKU,
 * high-volume brand that holds stock at two international 3PL locations.
 *
 * Tabs built by this script:
 *   SKU Master, Vendors, Inputs, Transactions, Purchase Orders, Shipments,
 *   Sales Plan (one per warehouse), Fulfillment & Carrier Actuals,
 *   Inventory Ledger (per-batch FIFO), StockModel (one per warehouse),
 *   Cashflow, Daily COGS, Landed Cost Summary.
 *
 * Quick start:
 *   1. Open a blank Google Sheet -> Extensions -> Apps Script.
 *   2. Paste this whole file, save, then run setupAll().
 *   3. setupAll() builds every tab in dependency order and (by default) loads a
 *      small SYNTHETIC demo dataset via seedDemoData(). Use setupStructureOnly()
 *      for an empty model.
 *
 * Runs as a BOUND script: SpreadsheetApp.getActiveSpreadsheet() always resolves to
 * the sheet the script was opened from, so no spreadsheet IDs are needed.
 *
 * Color legend: yellow = manual input, everything else is a live formula. Every
 * computed tab updates by itself when inputs change; build*() functions only need
 * to be re-run when the STRUCTURE changes.
 *
 * Re-running a single build*() function rebuilds only that tab (rebuildTab_ deletes
 * and recreates it). Other tabs that reference it lose those references (#REF!, or
 * a silent fallback where IFERROR-wrapped). After the model holds real data, rebuild
 * the whole chain in setupAll() order rather than one tab in isolation.
 */

// ---------------------------------------------------------------------------
// Configuration
// ---------------------------------------------------------------------------

// The stock/sales/COGS day grid starts on START_DATE. By default it is "today
// minus LOOKBACK_DAYS" so the demo always has ~60 days of history and the rest of
// the horizon in the future. To pin a fixed start, replace START_DATE with
// e.g. new Date(<year>, <month - 1>, <day>).
const LOOKBACK_DAYS = 60;
const START_DATE = dateFromToday_(-LOOKBACK_DAYS);
const HORIZON_DAYS = 360;

// Cashflow starts earlier than the stock grid: factory deposits are paid weeks
// before the first unit lands, and a payment dated before the grid would have no
// row to land on and silently drop out of the running total.
const CASHFLOW_LEAD_DAYS = 90;
const CASHFLOW_START_DATE = addDays_(START_DATE, -CASHFLOW_LEAD_DAYS);
const CASHFLOW_HORIZON_DAYS = HORIZON_DAYS + CASHFLOW_LEAD_DAYS;

const SEED_DEMO_DATA_ON_SETUP = true;

const SCT_HDR    = "#1a1a2e";
const SCT_HDR_FG = "#ffffff";
const SCT_YELLOW = "#fff9c4"; // manual input, editable

// SKU codes must match SKU Master's SKU column. Declared early because Shipments'
// per-SKU qty columns are derived from it.
const SKUS = ["SKU-1", "SKU-2", "SKU-3"];

// Two stock locations. Days of supply, FIFO layers, sales and COGS are always kept
// per warehouse and never blended: the two locations serve different markets with
// different demand, lead times and logistics costs.
const WAREHOUSES = ["WH-DE", "WH-CN"];

function dateFromToday_(offsetDays) {
  const t = new Date();
  return new Date(t.getFullYear(), t.getMonth(), t.getDate() + offsetDays);
}

function addDays_(date, days) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate() + days);
}

// Date anchors are written as =DATE(y,m,d) and evaluated by Sheets itself, rather
// than setValue(Date), which passes an absolute instant through the script's
// timezone into the spreadsheet's and can land on the wrong calendar day.
function dateFormula_(date) {
  return `=DATE(${date.getFullYear()},${date.getMonth() + 1},${date.getDate()})`;
}

// ---------------------------------------------------------------------------
// Orchestration
// ---------------------------------------------------------------------------

/**
 * Builds every tab in dependency order, then seeds the synthetic demo dataset
 * (unless SEED_DEMO_DATA_ON_SETUP is false).
 *
 * Order matters: several tabs build formulas against other tabs' column
 * positions. Transactions comes before Purchase Orders/Shipments because their
 * payment "Actual Date" formulas read it. Cross-tab formulas that point at a tab
 * built later (e.g. Purchase Orders -> Inventory Ledger) are IFERROR-guarded and
 * resolve as soon as that tab exists.
 */
function setupAll() {
  buildAllTabs_();
  if (SEED_DEMO_DATA_ON_SETUP) seedDemoData();
  removeDefaultBlankSheet_();
  Logger.log("Setup complete.");
}

function setupStructureOnly() {
  buildAllTabs_();
  removeDefaultBlankSheet_();
  Logger.log("Structure built, no data loaded.");
}

function buildAllTabs_() {
  buildSkuMasterAndVendors();
  buildInputs();
  buildTransactions();
  buildPurchaseOrders();
  buildShipments();
  WAREHOUSES.forEach(function (wh) { buildSalesPlan(wh); });
  buildFulfillmentCarrierActuals();
  buildInventoryLedger();
  WAREHOUSES.forEach(function (wh) { buildStockModel(wh); });
  buildCashflow();
  buildDailyCogs();
  buildLandedCostSummary();
}

function removeDefaultBlankSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const blank = ss.getSheetByName("Sheet1");
  if (blank && ss.getSheets().length > 1 && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

// ---------------------------------------------------------------------------
// Shared helpers
// ---------------------------------------------------------------------------

// Deletes the named tab if it exists, then recreates it in the same position, so
// re-running a build doesn't push the tab to the end of the list.
function rebuildTab_(ss, name) {
  const existing = ss.getSheetByName(name);
  const index = existing ? existing.getIndex() - 1 : null;
  if (existing) ss.deleteSheet(existing);
  const sh = ss.insertSheet(name);
  if (index !== null) {
    ss.setActiveSheet(sh);
    ss.moveActiveSheet(index + 1);
  }
  return sh;
}

// insertSheet() creates a 1000 x 26 grid; writing past it throws. Any table wider
// than 26 columns or taller than 1000 rows must widen the grid BEFORE writing,
// including the header row. Sizes are absolute (include the header row).
function ensureGrid_(sh, rowsNeeded, colsNeeded) {
  if (sh.getMaxRows() < rowsNeeded) sh.insertRowsAfter(sh.getMaxRows(), rowsNeeded - sh.getMaxRows());
  if (sh.getMaxColumns() < colsNeeded) sh.insertColumnsAfter(sh.getMaxColumns(), colsNeeded - sh.getMaxColumns());
}

function writeTable_(sh, columns, rows) {
  sh.getRange(1, 1, 1, columns.length)
    .setValues([columns])
    .setBackground(SCT_HDR)
    .setFontColor(SCT_HDR_FG)
    .setFontWeight("bold");
  sh.setFrozenRows(1);

  if (rows.length > 0) {
    sh.getRange(2, 1, rows.length, columns.length)
      .setValues(rows)
      .setBackground(SCT_YELLOW);
  }

  sh.autoResizeColumns(1, columns.length);
}

// Every formula resolves columns by header name, never by a hand-counted letter,
// so reordering a column list cannot silently mis-map a formula.
function colIndexIn_(columns, name) {
  const idx = columns.indexOf(name);
  if (idx === -1) throw new Error("Unknown column: " + name);
  return idx + 1;
}

function colLetterIn_(columns, name) {
  return columnToLetter_(colIndexIn_(columns, name));
}

function columnToLetter_(col) {
  let letter = "";
  let rest = col;
  while (rest > 0) {
    const rem = (rest - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    rest = Math.floor((rest - 1) / 26);
  }
  return letter;
}

// "B2" -> "$B$2"
function absA1_(a1) {
  return a1.replace(/^([A-Z]+)(\d+)$/, "$$$1$$$2");
}

function letterToColumn_(letter) {
  let col = 0;
  for (let i = 0; i < letter.length; i++) {
    col = col * 26 + (letter.charCodeAt(i) - 64);
  }
  return col;
}

function writeDayGrid_(sh, firstRow, dateCol, dateLetter, startDate, days) {
  sh.getRange(firstRow, dateCol).setFormula(dateFormula_(startDate));
  const dateFormulas = [];
  for (let r = firstRow + 1; r < firstRow + days; r++) dateFormulas.push([`=${dateLetter}${r - 1}+1`]);
  sh.getRange(firstRow + 1, dateCol, days - 1, 1).setFormulas(dateFormulas);
  sh.getRange(firstRow, dateCol, days, 1).setNumberFormat("yyyy-mm-dd");
}

// ---------------------------------------------------------------------------
// SKU Master + Vendors
// ---------------------------------------------------------------------------

const SKU_MASTER_COLUMNS = [
  "SKU", "Product Name", "Bundle?", "Units per Carton", "CBM per Carton",
  "Kg gross per Carton", "Factory / Vendor",
];

const SKU_MASTER_ROWS = [
  ["SKU-1", "Core product",      false, 100, 0.050, 10.0, "Factory A"],
  ["SKU-2", "Accessory",         false, 200, 0.060, 12.0, "Factory A"],
  ["SKU-3", "Add-on consumable", false, 100, 0.030, 6.0,  "Factory A"],
];

const VENDOR_TYPES = ["Factory", "Forwarder", "3PL", "Customs Broker"];

// Product 1/2/3 are numbered slots rather than a multi-value cell: each stays a
// plain dropdown and is trivially referenceable in formulas.
const VENDORS_COLUMNS = [
  "Vendor Name", "Type", "Product 1", "Product 2", "Product 3",
  "Contact Person", "Address", "Contact Email/Phone",
  "Payment Terms (standard)", "Contract Ref", "Bank Details", "Active?",
];

const VENDORS_ROWS = [
  ["Factory A",          "Factory",        "SKU-1", "SKU-2", "SKU-3", "", "", "", "30% deposit / 70% before shipment", "", "", true],
  ["Forwarder A",        "Forwarder",      "SKU-1", "SKU-2", "SKU-3", "", "", "", "", "", "", true],
  ["3PL DE",             "3PL",            "SKU-1", "SKU-2", "SKU-3", "", "", "", "", "", "", true],
  ["3PL CN",             "3PL",            "SKU-1", "SKU-2", "SKU-3", "", "", "", "", "", "", true],
  ["Customs Broker A",   "Customs Broker", "SKU-1", "SKU-2", "SKU-3", "", "", "", "", "", "", true],
];

function buildSkuMasterAndVendors() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const shSkuMaster = rebuildTab_(ss, "SKU Master");
  writeTable_(shSkuMaster, SKU_MASTER_COLUMNS, SKU_MASTER_ROWS);
  shSkuMaster.getRange(2, colIndexIn_(SKU_MASTER_COLUMNS, "Bundle?"), SKU_MASTER_ROWS.length, 1).insertCheckboxes();

  const shVendors = rebuildTab_(ss, "Vendors");
  writeTable_(shVendors, VENDORS_COLUMNS, VENDORS_ROWS);
  // Dropdowns instead of free text everywhere a value is reused in formulas, so
  // "Factory"/"factory" style drift cannot break a lookup.
  const typeRule = SpreadsheetApp.newDataValidation().requireValueInList(VENDOR_TYPES, true).setAllowInvalid(false).build();
  shVendors.getRange(2, colIndexIn_(VENDORS_COLUMNS, "Type"), VENDORS_ROWS.length, 1).setDataValidation(typeRule);
  const productRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(shSkuMaster.getRange("A2:A"), true)
    .setAllowInvalid(false).build();
  shVendors.getRange(2, colIndexIn_(VENDORS_COLUMNS, "Product 1"), VENDORS_ROWS.length, 3).setDataValidation(productRule);
  shVendors.getRange(2, colIndexIn_(VENDORS_COLUMNS, "Active?"), VENDORS_ROWS.length, 1).insertCheckboxes();

  SpreadsheetApp.flush();
  Logger.log("SKU Master + Vendors built. %s rows / %s rows.", SKU_MASTER_ROWS.length, VENDORS_ROWS.length);
}

// ---------------------------------------------------------------------------
// Inputs (standing config values referenced by formulas, never hardcoded)
// ---------------------------------------------------------------------------

const INPUTS_SHEET_NAME = "Inputs";
const INPUTS_PPV_THRESHOLD_CELL = "B2";

function buildInputs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = rebuildTab_(ss, INPUTS_SHEET_NAME);

  sh.getRange("A1:B1").setValues([["Standing config values (referenced by formulas elsewhere)", "Value"]])
    .setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);
  sh.getRange("A2").setValue("PPV trigger: abs. deviation of actual vs standard landed cost/unit that flags a standard rebuild");
  sh.getRange(INPUTS_PPV_THRESHOLD_CELL).setValue(0.03).setBackground(SCT_YELLOW).setNumberFormat("0%");
  sh.getRange("A3").setValue("Stock/sales grid start date (set by START_DATE in the script, informational)");
  sh.getRange("B3").setFormula(dateFormula_(START_DATE)).setNumberFormat("yyyy-mm-dd");
  sh.getRange("A4").setValue("Grid horizon, days (informational)");
  sh.getRange("B4").setValue(HORIZON_DAYS);

  sh.autoResizeColumns(1, 2);
  SpreadsheetApp.flush();
  Logger.log("Inputs built.");
}

// ---------------------------------------------------------------------------
// Purchase Orders
// ---------------------------------------------------------------------------
//
// One row per PO x SKU. Planned quantity is entered per warehouse; Qty ordered is
// their sum. Up to 3 payment tranches per row, each with a Planned Date (known up
// front) and an Actual Date that is matched live from the Transactions tab, so
// "Days Late" and "Payment Status" are always current.
//
// Actual Landed Cost/unit is a qty-weighted average of the Inventory Ledger's
// per-batch landed cost across every Shipments row of this PO, filtered per
// warehouse. It exists only as a purchase-price-variance (PPV) signal against
// the standard cost; COGS is never priced from this blended figure.

const PO_SHEET_NAME = "Purchase Orders";
const PO_TEMPLATE_ROWS = 60; // formulas pre-filled this many rows down

const PO_PAYMENT_SLOT_COUNT = 3;
const PO_PAYMENT_SLOT_NUMBERS = Array.from({ length: PO_PAYMENT_SLOT_COUNT }, function (_, i) { return i + 1; });

function poWarehouseCols_(prefix) {
  return WAREHOUSES.map(function (wh) { return `${prefix} — ${wh}`; });
}

const PO_COLUMNS = [
  "PO#", "Vendor", "SKU"].concat(poWarehouseCols_("Planned Qty")).concat([
  "Planned Ready Date", "Actual Ready Date",
  "Qty ordered", "Qty Produced",
  "Qty Remaining to Produce", "Qty Remaining to Ship",
  "EXW/unit", "Lab-Test Amount", "Inspection Amount", "Add-on Amount", "Total PO value", "Full Factory Cost/unit", "Currency", "Order date",
]).concat(PO_PAYMENT_SLOT_NUMBERS.reduce(function (cols, i) {
  return cols.concat([
    `Payment ${i} — Amount`, `Payment ${i} — Planned Date`, `Payment ${i} — Actual Date`,
    `Payment ${i} — Days Late`, `Payment ${i} — Paid?`,
  ]);
}, [])).concat([
  "Payments Sum Check", "Match?", "Payment Status", "Status",
  "Contract Link", "Invoice Link", "Add-on Link",
]).concat(poWarehouseCols_("Standard Landed Cost/unit"))
  .concat(poWarehouseCols_("Actual Landed Cost/unit"))
  .concat(poWarehouseCols_("PPV/unit"))
  .concat(poWarehouseCols_("PPV %"))
  .concat(["PPV — Freight", "PPV — Duty", "PPV — Factory"])
  .concat(poWarehouseCols_("PPV Trigger"));

const PO_CURRENCIES = ["EUR", "USD", "CNY"];
const PO_STATUSES = ["Draft", "Placed", "In Production", "Partially Shipped", "Closed"];
const PPV_TRIGGER_TEXT = "REBUILD STANDARD";

const PO_MANUAL_COLUMNS = [
  "PO#", "Vendor", "SKU"].concat(poWarehouseCols_("Planned Qty")).concat([
  "Planned Ready Date", "Actual Ready Date",
  "Qty Produced",
  "EXW/unit", "Lab-Test Amount", "Inspection Amount", "Add-on Amount", "Currency", "Order date",
]).concat(PO_PAYMENT_SLOT_NUMBERS.reduce(function (cols, i) {
  return cols.concat([`Payment ${i} — Amount`, `Payment ${i} — Planned Date`]);
}, [])).concat([
  "Status", "Contract Link", "Invoice Link", "Add-on Link",
]).concat(poWarehouseCols_("Standard Landed Cost/unit"))
  .concat(["PPV — Freight", "PPV — Duty", "PPV — Factory"]);

function poCol_(name) { return colLetterIn_(PO_COLUMNS, name); }
function colIndex_(name) { return colIndexIn_(PO_COLUMNS, name); }

/**
 * Formula fragment that resolves a payment's Actual Date from the Transactions tab.
 *
 * `idExpr` is this row's PO# or Shipment ID; it is matched against Transactions'
 * free-text "PO#/Shipment Ref" column as an EXACT token inside a comma-separated
 * list. Both sides are comma-wrapped (","&x&",") so "PO-1" never matches inside
 * "PO-10", and one payment covering several rows can list them all.
 *
 * `amountExpr` is this slot's expected amount. When several transactions carry
 * the same ref (several tranches to one PO), the one closest in amount wins. If
 * even the closest is off by more than max(10% of expected, 10), it is treated as
 * no match: a same-ref-but-wrong-tranche payment must not mark a slot as paid.
 *
 * Deliberately built from inline range expressions, not LET: a LET-bound name
 * holding an array comparison undergoes implicit intersection when passed to
 * FILTER/MATCH and collapses to a single cell.
 */
function txnRefMatchDateFormula_(idExpr, amountExpr) {
  const refLetter = txnCol_("PO#/Shipment Ref");
  const paidLetter = txnCol_("Date paid");
  const eurLetter = txnCol_("Amount (EUR)");
  const lastRow = TRANSACTIONS_TEMPLATE_ROWS + 1;
  const wrappedRefs = `(","&SUBSTITUTE(Transactions!$${refLetter}$2:$${refLetter}$${lastRow},", ",",")&",")`;
  const wrappedId = `(","&${idExpr}&",")`;
  const matchCond = `ISNUMBER(SEARCH(${wrappedId},${wrappedRefs}))`;
  const diffExpr = `ABS(Transactions!$${eurLetter}$2:$${eurLetter}$${lastRow}-${amountExpr})`;
  const minDiff = `MIN(FILTER(${diffExpr},${matchCond}))`;
  return `IFERROR(IF(${minDiff}>MAX(${amountExpr}*0.1,10),"",INDEX(FILTER(Transactions!$${paidLetter}$2:$${paidLetter}$${lastRow},${matchCond}),MATCH(${minDiff},FILTER(${diffExpr},${matchCond}),0))),"")`;
}

function buildPurchaseOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("SKU Master") || !ss.getSheetByName("Vendors") || !ss.getSheetByName(INPUTS_SHEET_NAME)) {
    throw new Error("Build SKU Master + Vendors + Inputs first: run buildSkuMasterAndVendors() then buildInputs().");
  }

  const sh = rebuildTab_(ss, PO_SHEET_NAME);
  ensureGrid_(sh, PO_TEMPLATE_ROWS + 1, PO_COLUMNS.length);
  writeTable_(sh, PO_COLUMNS, []);

  const n = PO_TEMPLATE_ROWS;
  const qtyOrdered   = poCol_("Qty ordered");
  const qtyProduced  = poCol_("Qty Produced");
  const skuCol       = poCol_("SKU");
  const poNumCol     = poCol_("PO#");
  const paymentDateCols = PO_PAYMENT_SLOT_NUMBERS.map(function (i) {
    return {
      amount: poCol_(`Payment ${i} — Amount`),
      planned: poCol_(`Payment ${i} — Planned Date`),
      actual: poCol_(`Payment ${i} — Actual Date`),
      daysLate: poCol_(`Payment ${i} — Days Late`),
      paid: poCol_(`Payment ${i} — Paid?`),
    };
  });
  const totalVal = poCol_("Total PO value");
  const sumCheck = poCol_("Payments Sum Check");
  const whCols = {};
  WAREHOUSES.forEach(function (wh) {
    whCols[wh] = {
      standard: poCol_(`Standard Landed Cost/unit — ${wh}`),
      actual:   poCol_(`Actual Landed Cost/unit — ${wh}`),
      ppvUnit:  poCol_(`PPV/unit — ${wh}`),
      ppvPct:   poCol_(`PPV % — ${wh}`),
    };
  });
  const shPoLetter = shipCol_("PO#");
  const shWhLetter = shipCol_("Actual Warehouse");
  const shLastRow = SHIP_TEMPLATE_ROWS + 1;

  const remainToProduceFormulas = [];
  const remainToShipFormulas    = [];
  const sumCheckFormulas        = [];
  const matchFormulas           = [];
  const paymentStatusFormulas   = [];
  const ppvUnitFormulas = {}, ppvPctFormulas = {}, ppvTriggerFormulas = {}, actualLandedFormulas = {};
  WAREHOUSES.forEach(function (wh) {
    ppvUnitFormulas[wh] = []; ppvPctFormulas[wh] = []; ppvTriggerFormulas[wh] = []; actualLandedFormulas[wh] = [];
  });
  const daysLateFormulas = paymentDateCols.map(function () { return []; });
  const actualDateFormulas = paymentDateCols.map(function () { return []; });

  for (let r = 2; r <= n + 1; r++) {
    remainToProduceFormulas.push([`=IF(${qtyOrdered}${r}="","",${qtyOrdered}${r}-${qtyProduced}${r})`]);

    // Shipments is wide (one qty column per SKU), so which column to sum depends
    // on this row's SKU. A formula cannot pick a column letter dynamically, so one
    // SUMIFS per SKU is built and SWITCH selects the right one.
    const shipSkuSwitchCases = SKUS.map(function (sku) {
      const qtyLetter = shipCol_(shipSkuQtyCol_(sku));
      return `"${sku}",SUMIFS(Shipments!$${qtyLetter}:$${qtyLetter},Shipments!$${shPoLetter}:$${shPoLetter},${poNumCol}${r})`;
    }).join(",");
    remainToShipFormulas.push([`=IFERROR(IF(${qtyOrdered}${r}="","",${qtyOrdered}${r}-SWITCH(${skuCol}${r},${shipSkuSwitchCases},0)),${qtyOrdered}${r})`]);

    sumCheckFormulas.push([`=${paymentDateCols.map(function (c) { return `${c.amount}${r}`; }).join("+")}`]);

    matchFormulas.push([`=IF(${totalVal}${r}="","",IF(ROUND(${sumCheck}${r}-${totalVal}${r},2)=0,"OK","MISMATCH"))`]);

    // nSlots = slots with a real amount; pPaid = how many of those are paid.
    const nSlots = "(" + paymentDateCols.map(function (c) { return `(${c.amount}${r}>0)`; }).join("+") + ")";
    const pPaid  = "(" + paymentDateCols.map(function (c) { return `(${c.amount}${r}>0)*${c.paid}${r}`; }).join("+") + ")";
    paymentStatusFormulas.push([`=IF(${totalVal}${r}="","",IF(${pPaid}=0,"Not Paid",IF(${pPaid}=${nSlots},"Fully Paid","Partially Paid")))`]);

    // Days Late: once paid, Actual - Planned (negative = paid early). Unpaid and
    // past due: running lateness so far (TODAY() - Planned). Otherwise blank.
    paymentDateCols.forEach(function (c, i) {
      const formula = `=IF(AND(${c.planned}${r}<>"",${c.actual}${r}<>""),${c.actual}${r}-${c.planned}${r},`
        + `IF(AND(${c.planned}${r}<>"",${c.actual}${r}="",TODAY()>${c.planned}${r},NOT(${c.paid}${r})),TODAY()-${c.planned}${r},""))`;
      daysLateFormulas[i].push([formula]);
    });

    // An unused slot (blank amount) must not go hunting for a spurious match.
    paymentDateCols.forEach(function (c, i) {
      actualDateFormulas[i].push([`=IF(${c.amount}${r}="","",${txnRefMatchDateFormula_(`${poNumCol}${r}`, `${c.amount}${r}`)})`]);
    });

    // Actual Landed Cost/unit per warehouse. Inventory Ledger's batch rows are 1:1
    // row-aligned with Shipments (same order, same count), so Shipments' PO#/qty
    // columns pair positionally with the Ledger's same-row Landed Cost/unit. The
    // chain followed is PO -> Shipments -> Ledger. N() guards blank cells in the
    // SUMPRODUCT arithmetic.
    WAREHOUSES.forEach(function (wh) {
      const c = whCols[wh];
      const whMatch = `(Shipments!$${shWhLetter}$2:$${shWhLetter}$${shLastRow}="${wh}")`;
      const actualLandedSwitchCases = SKUS.map(function (sku) {
        const qtyLetter = shipCol_(shipSkuQtyCol_(sku));
        const costLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Landed Cost/unit"));
        const poMatch = `(Shipments!$${shPoLetter}$2:$${shPoLetter}$${shLastRow}=${poNumCol}${r})`;
        const num = `SUMPRODUCT(${poMatch}*${whMatch}*N(Shipments!$${qtyLetter}$2:$${qtyLetter}$${shLastRow})*N('${INV_LEDGER_SHEET_NAME}'!$${costLetter}$3:$${costLetter}$${SHIP_TEMPLATE_ROWS + 2}))`;
        const den = `SUMPRODUCT(${poMatch}*${whMatch}*N(Shipments!$${qtyLetter}$2:$${qtyLetter}$${shLastRow}))`;
        return `"${sku}",IF(${den}=0,"",${num}/${den})`;
      }).join(",");
      actualLandedFormulas[wh].push([`=IFERROR(SWITCH(${skuCol}${r},${actualLandedSwitchCases},""),"")`]);

      ppvUnitFormulas[wh].push([`=IF(OR(${c.standard}${r}="",${c.actual}${r}=""),"",${c.actual}${r}-${c.standard}${r})`]);
      ppvPctFormulas[wh].push([`=IF(OR(${c.ppvUnit}${r}="",${c.standard}${r}="",${c.standard}${r}=0),"",${c.ppvUnit}${r}/${c.standard}${r})`]);
      ppvTriggerFormulas[wh].push([`=IF(${c.ppvPct}${r}="","",IF(ABS(${c.ppvPct}${r})>${INPUTS_SHEET_NAME}!${absA1_(INPUTS_PPV_THRESHOLD_CELL)},"${PPV_TRIGGER_TEXT}","OK"))`]);
    });
  }

  sh.getRange(2, colIndex_("Qty Remaining to Produce"), n, 1).setFormulas(remainToProduceFormulas);
  sh.getRange(2, colIndex_("Qty Remaining to Ship"), n, 1).setFormulas(remainToShipFormulas);

  // EXW/unit is the manual contract price; Total PO value is computed from it plus
  // flat per-PO fees (lab test, inspection, add-ons).
  const exwUnit = poCol_("EXW/unit");
  const labTest = poCol_("Lab-Test Amount");
  const inspectionAmt = poCol_("Inspection Amount");
  const addon = poCol_("Add-on Amount");

  const totalValFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    totalValFormulas.push([
      `=IF(OR(${qtyOrdered}${r}="",${exwUnit}${r}=""),"",${qtyOrdered}${r}*${exwUnit}${r}+N(${labTest}${r})+N(${inspectionAmt}${r})+N(${addon}${r}))`
    ]);
  }
  sh.getRange(2, colIndex_("Total PO value"), n, 1).setFormulas(totalValFormulas).setNumberFormat("#,##0.00");
  // 4 decimals: unit prices for this kind of product run below 1.00.
  sh.getRange(2, colIndex_("EXW/unit"), n, 1).setNumberFormat("#,##0.0000");

  // Full Factory Cost/unit = Total PO value / Qty ordered, i.e. EXW/unit with its
  // per-unit share of fees folded in. The Inventory Ledger prices batches from
  // this, so fees are never lost from landed cost; EXW/unit itself stays the pure
  // contract price.
  const fullFactoryCostFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    fullFactoryCostFormulas.push([
      `=IF(OR(${qtyOrdered}${r}="",${qtyOrdered}${r}=0,${totalVal}${r}=""),"",${totalVal}${r}/${qtyOrdered}${r})`
    ]);
  }
  sh.getRange(2, colIndex_("Full Factory Cost/unit"), n, 1).setFormulas(fullFactoryCostFormulas).setNumberFormat("#,##0.0000");

  // Qty ordered = sum of the per-warehouse planned quantities (the split varies
  // per PO, so it is entered by hand, not derived from a fixed ratio).
  const plannedQtyLetters = WAREHOUSES.map(function (wh) { return poCol_(`Planned Qty — ${wh}`); });
  const qtyOrderedFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const allBlank = plannedQtyLetters.map(function (l) { return `${l}${r}=""`; }).join(",");
    const sum = plannedQtyLetters.map(function (l) { return `N(${l}${r})`; }).join("+");
    qtyOrderedFormulas.push([`=IF(AND(${allBlank}),"",${sum})`]);
  }
  sh.getRange(2, colIndex_("Qty ordered"), n, 1).setFormulas(qtyOrderedFormulas);

  sh.getRange(2, colIndex_("Payments Sum Check"), n, 1).setFormulas(sumCheckFormulas);
  sh.getRange(2, colIndex_("Match?"), n, 1).setFormulas(matchFormulas);
  sh.getRange(2, colIndex_("Payment Status"), n, 1).setFormulas(paymentStatusFormulas);
  paymentDateCols.forEach(function (c, i) {
    sh.getRange(2, letterToColumn_(c.daysLate), n, 1).setFormulas(daysLateFormulas[i]);
    sh.getRange(2, letterToColumn_(c.actual), n, 1).setFormulas(actualDateFormulas[i]).setNumberFormat("yyyy-mm-dd");
    const paidFormulas = [];
    for (let r = 2; r <= n + 1; r++) paidFormulas.push([`=${c.actual}${r}<>""`]);
    sh.getRange(2, letterToColumn_(c.paid), n, 1).setFormulas(paidFormulas);
  });
  WAREHOUSES.forEach(function (wh) {
    sh.getRange(2, colIndex_(`Actual Landed Cost/unit — ${wh}`), n, 1).setFormulas(actualLandedFormulas[wh]);
    sh.getRange(2, colIndex_(`PPV/unit — ${wh}`), n, 1).setFormulas(ppvUnitFormulas[wh]);
    sh.getRange(2, colIndex_(`PPV % — ${wh}`), n, 1).setFormulas(ppvPctFormulas[wh]);
    sh.getRange(2, colIndex_(`PPV Trigger — ${wh}`), n, 1).setFormulas(ppvTriggerFormulas[wh]);
  });

  PO_MANUAL_COLUMNS.forEach(function (name) {
    sh.getRange(2, colIndex_(name), n, 1).setBackground(SCT_YELLOW);
  });

  const vendorRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName("Vendors").getRange("A2:A"), true)
    .setAllowInvalid(false).build();
  sh.getRange(2, colIndex_("Vendor"), n, 1).setDataValidation(vendorRule);

  const skuRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName("SKU Master").getRange("A2:A"), true)
    .setAllowInvalid(false).build();
  sh.getRange(2, colIndex_("SKU"), n, 1).setDataValidation(skuRule);

  const currencyRule = SpreadsheetApp.newDataValidation().requireValueInList(PO_CURRENCIES, true).setAllowInvalid(false).build();
  sh.getRange(2, colIndex_("Currency"), n, 1).setDataValidation(currencyRule);

  const statusRule = SpreadsheetApp.newDataValidation().requireValueInList(PO_STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, colIndex_("Status"), n, 1).setDataValidation(statusRule);

  // Payment Actual Dates are formulas, so they get no date-picker validation.
  const PO_DATE_COLUMNS = ["Order date", "Planned Ready Date", "Actual Ready Date"]
    .concat(PO_PAYMENT_SLOT_NUMBERS.map(function (i) { return `Payment ${i} — Planned Date`; }));

  // Date validation is what enables the double-click calendar picker;
  // setNumberFormat only controls display.
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  PO_DATE_COLUMNS.forEach(function (name) {
    sh.getRange(2, colIndex_(name), n, 1).setDataValidation(dateRule).setNumberFormat("yyyy-mm-dd");
  });
  const PO_MONEY_COLUMNS = ["Total PO value", "Lab-Test Amount", "Inspection Amount", "Add-on Amount"]
    .concat(PO_PAYMENT_SLOT_NUMBERS.map(function (i) { return `Payment ${i} — Amount`; }))
    .concat(["Payments Sum Check"])
    .concat(poWarehouseCols_("Standard Landed Cost/unit"))
    .concat(poWarehouseCols_("Actual Landed Cost/unit"))
    .concat(poWarehouseCols_("PPV/unit"))
    .concat(["PPV — Freight", "PPV — Duty", "PPV — Factory"]);
  const PO_QTY_COLUMNS = poWarehouseCols_("Planned Qty").concat([
    "Qty ordered", "Qty Produced", "Qty Remaining to Produce", "Qty Remaining to Ship",
  ]);

  PO_MONEY_COLUMNS.forEach(function (name) {
    sh.getRange(2, colIndex_(name), n, 1).setNumberFormat("#,##0.00");
  });
  PO_QTY_COLUMNS.forEach(function (name) {
    sh.getRange(2, colIndex_(name), n, 1).setNumberFormat("#,##0");
  });
  WAREHOUSES.forEach(function (wh) {
    sh.getRange(2, colIndex_(`PPV % — ${wh}`), n, 1).setNumberFormat("0.00%");
  });
  paymentDateCols.forEach(function (c) {
    sh.getRange(2, letterToColumn_(c.daysLate), n, 1).setNumberFormat("#,##0");
  });

  // Red = problem: payment mismatch, PPV rebuild trigger, any overdue tranche.
  const rules = sh.getConditionalFormatRules();
  rules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("MISMATCH")
    .setBackground("#f8d7da")
    .setRanges([sh.getRange(2, colIndex_("Match?"), n, 1)])
    .build());
  WAREHOUSES.forEach(function (wh) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(PPV_TRIGGER_TEXT)
      .setBackground("#f8d7da")
      .setRanges([sh.getRange(2, colIndex_(`PPV Trigger — ${wh}`), n, 1)])
      .build());
  });
  paymentDateCols.forEach(function (c) {
    rules.push(SpreadsheetApp.newConditionalFormatRule()
      .whenNumberGreaterThan(0)
      .setBackground("#f8d7da")
      .setRanges([sh.getRange(2, letterToColumn_(c.daysLate), n, 1)])
      .build());
  });
  sh.setConditionalFormatRules(rules);

  SpreadsheetApp.flush();
  Logger.log("Purchase Orders built: %s template rows.", n);
}

// ---------------------------------------------------------------------------
// Shipments
// ---------------------------------------------------------------------------
//
// One row per shipment, with one qty column per SKU. Convention: ONE SKU PER ROW.
// A mixed container is entered as several rows (one per SKU) sharing a wave/
// container label, each carrying its own share of freight and customs. That keeps
// every downstream number (FIFO batch, landed cost/unit, Landed Cost Summary)
// exact per SKU instead of relying on an estimated blend.
//
// "Planned Warehouse" and "Actual Warehouse" are separate on purpose: an
// in-transit shipment has no Actual Warehouse yet, so transit and pipeline logic
// filters on Planned Warehouse while landed-stock logic filters on Actual.
//
// Cost tracking lives inline: 2 freight tranches + 1 customs tranche (each with
// Planned/Actual date), plus the invoiced landed-cost components. Gross Cost/unit
// includes refundable import VAT (it is real cash out); Net Cost/unit excludes it
// (it is the true cost once reclaimed).

const SHIPMENTS_SHEET_NAME = "Shipments";
const SHIP_TEMPLATE_ROWS = 60;

function shipSkuQtyCol_(sku) { return sku + " Qty"; }

const SHIPMENTS_COLUMNS = [
  "Shipment ID", "PO#", "Wave#", "Ship method",
  "Planned Warehouse", "Actual Warehouse",
  "Depart Date (planned)", "Actual Depart Date", "ETA (planned)", "Projected ETA", "Effective ETA", "Actual Arrival Date",
].concat(
  SKUS.map(shipSkuQtyCol_)
).concat([
  "Payment 1 (Freight) — Amount", "Payment 1 (Freight) — Planned Date", "Payment 1 (Freight) — Actual Date", "Payment 1 (Freight) — Paid?",
  "Payment 2 (Freight) — Amount", "Payment 2 (Freight) — Planned Date", "Payment 2 (Freight) — Actual Date", "Payment 2 (Freight) — Paid?",
  "Customs — Estimated Amount", "Customs — Amount", "Customs — Planned Date", "Customs — Actual Date", "Customs — Paid?",
  "Customs Status",
  "Quote Link", "Quote Date", "Invoice Link", "Invoice Date",
  "Customs Invoice Link", "Customs Invoice Date", "Customs Declaration Link",
  "Total Qty",
  "Projected Shipping+Taxes Cost/unit",
  "Customs Declared Value/unit",
  "Actual — Delivery", "Actual — Admin Fees", "Actual — Duty (non-refundable)", "Implied Duty Rate",
  "Actual — Import VAT (refundable)", "Actual — Other VAT (refundable)",
  "Gross Cost/unit", "Net Cost/unit",
]);

const SHIP_METHODS = ["Air", "Rail", "Sea", "Local"];
const SHIP_CUSTOMS_STATUSES = ["Not started", "In clearance", "Held", "Cleared"];

const SHIP_MANUAL_COLUMNS = [
  "Shipment ID", "PO#", "Wave#", "Ship method",
  "Planned Warehouse", "Actual Warehouse",
  "Depart Date (planned)", "Actual Depart Date", "ETA (planned)", "Actual Arrival Date",
].concat(
  SKUS.map(shipSkuQtyCol_)
).concat([
  "Payment 1 (Freight) — Amount", "Payment 1 (Freight) — Planned Date",
  "Payment 2 (Freight) — Amount", "Payment 2 (Freight) — Planned Date",
  "Customs — Estimated Amount", "Customs — Planned Date",
  "Customs Status",
  "Quote Link", "Quote Date", "Invoice Link", "Invoice Date",
  "Customs Invoice Link", "Customs Invoice Date", "Customs Declaration Link",
  "Customs Declared Value/unit",
  "Actual — Delivery", "Actual — Admin Fees", "Actual — Duty (non-refundable)",
  "Actual — Import VAT (refundable)", "Actual — Other VAT (refundable)",
]);

function shipCol_(name) { return colLetterIn_(SHIPMENTS_COLUMNS, name); }
function shipColIndex_(name) { return colIndexIn_(SHIPMENTS_COLUMNS, name); }

function buildShipments() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("SKU Master") || !ss.getSheetByName("Vendors") || !ss.getSheetByName(PO_SHEET_NAME)) {
    throw new Error("Build SKU Master + Vendors + Purchase Orders first.");
  }

  const sh = rebuildTab_(ss, SHIPMENTS_SHEET_NAME);
  ensureGrid_(sh, SHIP_TEMPLATE_ROWS + 1, SHIPMENTS_COLUMNS.length);
  writeTable_(sh, SHIPMENTS_COLUMNS, []);

  const n = SHIP_TEMPLATE_ROWS;

  const qtyLetters = SKUS.map(function (sku) { return shipCol_(shipSkuQtyCol_(sku)); });
  const totalQtyLetter = shipCol_("Total Qty");
  const totalQtyFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const sum = qtyLetters.map(function (l) { return `${l}${r}`; }).join("+");
    totalQtyFormulas.push([`=IF(${sum}=0,"",${sum})`]);
  }
  sh.getRange(2, shipColIndex_("Total Qty"), n, 1).setFormulas(totalQtyFormulas).setNumberFormat("#,##0");

  // Projected ETA = planned departure + this ship method's historical average
  // transit time (Actual Arrival - Actual Depart over past shipments). Kept for
  // reference only. Effective ETA, which everything downstream reads, mirrors the
  // hand-entered ETA (planned) and stays blank without one: a missing ETA should
  // surface as a visible gap, not be papered over with an estimate.
  const departPlannedLetter = shipCol_("Depart Date (planned)");
  const actualDepartLetter = shipCol_("Actual Depart Date");
  const etaPlannedLetter = shipCol_("ETA (planned)");
  const actualArrivalLetter = shipCol_("Actual Arrival Date");
  const shipMethodLetter = shipCol_("Ship method");

  const projectedEtaFormulas = [], effectiveEtaFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const transitHistory =
      `FILTER($${actualArrivalLetter}$2:$${actualArrivalLetter}$${n + 1}-$${actualDepartLetter}$2:$${actualDepartLetter}$${n + 1},` +
      `$${shipMethodLetter}$2:$${shipMethodLetter}$${n + 1}=$${shipMethodLetter}${r},` +
      `$${actualArrivalLetter}$2:$${actualArrivalLetter}$${n + 1}<>"",$${actualDepartLetter}$2:$${actualDepartLetter}$${n + 1}<>"")`;
    projectedEtaFormulas.push([`=IF($${departPlannedLetter}${r}="","",IFERROR($${departPlannedLetter}${r}+AVERAGE(${transitHistory}),""))`]);
    effectiveEtaFormulas.push([`=IF($${etaPlannedLetter}${r}="","",$${etaPlannedLetter}${r})`]);
  }
  sh.getRange(2, shipColIndex_("Projected ETA"), n, 1).setFormulas(projectedEtaFormulas).setNumberFormat("yyyy-mm-dd");
  sh.getRange(2, shipColIndex_("Effective ETA"), n, 1).setFormulas(effectiveEtaFormulas).setNumberFormat("yyyy-mm-dd");

  const delivery = shipCol_("Actual — Delivery"), admin = shipCol_("Actual — Admin Fees");
  const duty = shipCol_("Actual — Duty (non-refundable)");
  const importVat = shipCol_("Actual — Import VAT (refundable)"), otherVat = shipCol_("Actual — Other VAT (refundable)");
  const grossLetter = shipCol_("Gross Cost/unit");
  const declaredValueLetter = shipCol_("Customs Declared Value/unit");

  // Sanity check only, never feeds Gross/Net: duty is always the invoiced actual.
  // Comparing it to the declared value surfaces a wrong duty or declared-value
  // entry immediately instead of at month-end reconciliation.
  const impliedDutyFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const base = `${totalQtyLetter}${r}*${declaredValueLetter}${r}`;
    impliedDutyFormulas.push([`=IFERROR(IF(OR(${duty}${r}="",${declaredValueLetter}${r}="",${totalQtyLetter}${r}=""),"",${duty}${r}/(${base})),"")`]);
  }
  sh.getRange(2, shipColIndex_("Implied Duty Rate"), n, 1).setFormulas(impliedDutyFormulas).setNumberFormat("0.0%");

  const grossFormulas = [], netFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const componentSum = `${delivery}${r}+${admin}${r}+${duty}${r}+${importVat}${r}+${otherVat}${r}`;
    // Blank stays blank: qty is often known well before any cost is invoiced.
    // Without this guard the empty sum reads as 0 and the Ledger would book the
    // batch at factory cost only, with no signal that logistics cost is missing.
    const hasAnyActualCost = `OR(${delivery}${r}<>"",${admin}${r}<>"",${duty}${r}<>"",${importVat}${r}<>"",${otherVat}${r}<>"")`;
    grossFormulas.push([`=IF(OR(${totalQtyLetter}${r}="",${totalQtyLetter}${r}=0,NOT(${hasAnyActualCost})),"",(${componentSum})/${totalQtyLetter}${r})`]);
    // Net = Gross minus the refundable VAT share. Checks Gross's own blankness so
    // it never subtracts from an empty-string cell.
    netFormulas.push([`=IF(OR(${totalQtyLetter}${r}="",${totalQtyLetter}${r}=0,${grossLetter}${r}=""),"",${grossLetter}${r}-(${importVat}${r}+${otherVat}${r})/${totalQtyLetter}${r})`]);
  }
  sh.getRange(2, shipColIndex_("Gross Cost/unit"), n, 1).setFormulas(grossFormulas).setNumberFormat("#,##0.0000");
  sh.getRange(2, shipColIndex_("Net Cost/unit"), n, 1).setFormulas(netFormulas).setNumberFormat("#,##0.0000");

  // Historical average Gross Cost/unit for this ship method: a live estimate for
  // shipments not yet invoiced, so Cashflow always has a number to plan against.
  // Gross, not Net, because it feeds projected cash OUT. A row with no Gross yet is
  // blank and AVERAGEIFS skips it, so self-reference is harmless.
  const projFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    projFormulas.push([`=IFERROR(AVERAGEIFS($${grossLetter}$2:$${grossLetter}$${n + 1},$${shipMethodLetter}$2:$${shipMethodLetter}$${n + 1},${shipMethodLetter}${r}),"")`]);
  }
  sh.getRange(2, shipColIndex_("Projected Shipping+Taxes Cost/unit"), n, 1).setFormulas(projFormulas);

  // Freight Actual Date/Paid? are matched from Transactions by this row's
  // Shipment ID (a freight invoice belongs to a specific shipment, not a PO).
  const shipIdLetter = shipCol_("Shipment ID");
  [1, 2].forEach(function (i) {
    const amtLetter = shipCol_(`Payment ${i} (Freight) — Amount`);
    const actualFormulas = [];
    for (let r = 2; r <= n + 1; r++) {
      actualFormulas.push([`=IF(${amtLetter}${r}="","",${txnRefMatchDateFormula_(`${shipIdLetter}${r}`, `${amtLetter}${r}`)})`]);
    }
    const actualCol = shipColIndex_(`Payment ${i} (Freight) — Actual Date`);
    sh.getRange(2, actualCol, n, 1).setFormulas(actualFormulas).setNumberFormat("yyyy-mm-dd");
    const paidFormulas = [];
    const actualLetter = columnToLetter_(actualCol);
    for (let r = 2; r <= n + 1; r++) paidFormulas.push([`=${actualLetter}${r}<>""`]);
    sh.getRange(2, shipColIndex_(`Payment ${i} (Freight) — Paid?`), n, 1).setFormulas(paidFormulas);
  });

  // Customs — Amount = this row's Admin + Duty + Import VAT + Other VAT (a customs
  // broker invoice bundles them). Delivery is excluded: that is freight, tracked by
  // the two freight tranches. Blank when no component is entered yet, so an
  // unpriced row reads "unknown" rather than a confirmed zero. A single customs
  // payment covering a multi-SKU container lists every row's Shipment ID in its
  // Transactions ref, so each row finds the same payment independently.
  const customsAmtLetter = shipCol_("Customs — Amount");
  const customsAmtFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const ownSum = `N(${admin}${r})+N(${duty}${r})+N(${importVat}${r})+N(${otherVat}${r})`;
    const hasAnyCustomsCost = `OR(${admin}${r}<>"",${duty}${r}<>"",${importVat}${r}<>"",${otherVat}${r}<>"")`;
    customsAmtFormulas.push([`=IF(OR(${shipIdLetter}${r}="",NOT(${hasAnyCustomsCost})),"",${ownSum})`]);
  }
  sh.getRange(2, letterToColumn_(customsAmtLetter), n, 1).setFormulas(customsAmtFormulas).setNumberFormat("#,##0.00");

  const customsActualCol = shipColIndex_("Customs — Actual Date");
  const customsActualFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    customsActualFormulas.push([`=IF(${customsAmtLetter}${r}="","",${txnRefMatchDateFormula_(`${shipIdLetter}${r}`, `${customsAmtLetter}${r}`)})`]);
  }
  sh.getRange(2, customsActualCol, n, 1).setFormulas(customsActualFormulas).setNumberFormat("yyyy-mm-dd");
  const customsActualLetter = columnToLetter_(customsActualCol);
  const customsPaidFormulas = [];
  for (let r = 2; r <= n + 1; r++) customsPaidFormulas.push([`=${customsActualLetter}${r}<>""`]);
  sh.getRange(2, shipColIndex_("Customs — Paid?"), n, 1).setFormulas(customsPaidFormulas);

  SHIP_MANUAL_COLUMNS.forEach(function (name) {
    sh.getRange(2, shipColIndex_(name), n, 1).setBackground(SCT_YELLOW);
  });

  const poRule = SpreadsheetApp.newDataValidation()
    .requireValueInRange(ss.getSheetByName(PO_SHEET_NAME).getRange(`A2:A${PO_TEMPLATE_ROWS + 1}`), true)
    .setAllowInvalid(true).build(); // several SKU rows share one PO#, so the source range has duplicates
  sh.getRange(2, shipColIndex_("PO#"), n, 1).setDataValidation(poRule);

  const shipMethodRule = SpreadsheetApp.newDataValidation().requireValueInList(SHIP_METHODS, true).setAllowInvalid(false).build();
  sh.getRange(2, shipColIndex_("Ship method"), n, 1).setDataValidation(shipMethodRule);

  const warehouseRule = SpreadsheetApp.newDataValidation().requireValueInList(WAREHOUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, shipColIndex_("Planned Warehouse"), n, 1).setDataValidation(warehouseRule);
  sh.getRange(2, shipColIndex_("Actual Warehouse"), n, 1).setDataValidation(warehouseRule);

  // Clearance status is tracked separately from payment: customs can be paid and
  // still be held.
  const customsStatusRule = SpreadsheetApp.newDataValidation().requireValueInList(SHIP_CUSTOMS_STATUSES, true).setAllowInvalid(false).build();
  sh.getRange(2, shipColIndex_("Customs Status"), n, 1).setDataValidation(customsStatusRule);
  const shConditionalRules = sh.getConditionalFormatRules();
  shConditionalRules.push(SpreadsheetApp.newConditionalFormatRule()
    .whenTextEqualTo("Held")
    .setBackground("#f8d7da")
    .setRanges([sh.getRange(2, shipColIndex_("Customs Status"), n, 1)])
    .build());
  sh.setConditionalFormatRules(shConditionalRules);

  const dateCols = [
    "Depart Date (planned)", "Actual Depart Date", "ETA (planned)", "Actual Arrival Date",
    "Payment 1 (Freight) — Planned Date",
    "Payment 2 (Freight) — Planned Date",
    "Customs — Planned Date",
    "Quote Date", "Invoice Date", "Customs Invoice Date",
  ];
  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build();
  dateCols.forEach(function (name) {
    sh.getRange(2, shipColIndex_(name), n, 1).setDataValidation(dateRule).setNumberFormat("yyyy-mm-dd");
  });

  const moneyCols = [
    "Payment 1 (Freight) — Amount", "Payment 2 (Freight) — Amount",
    "Customs — Estimated Amount",
    "Actual — Delivery", "Actual — Admin Fees",
    "Actual — Duty (non-refundable)", "Actual — Import VAT (refundable)", "Actual — Other VAT (refundable)",
  ];
  moneyCols.forEach(function (name) {
    sh.getRange(2, shipColIndex_(name), n, 1).setNumberFormat("#,##0.00");
  });
  ["Customs Declared Value/unit", "Projected Shipping+Taxes Cost/unit"].forEach(function (name) {
    sh.getRange(2, shipColIndex_(name), n, 1).setNumberFormat("#,##0.0000");
  });

  SKUS.forEach(function (sku) {
    sh.getRange(2, shipColIndex_(shipSkuQtyCol_(sku)), n, 1).setNumberFormat("#,##0");
  });

  sh.autoResizeColumns(1, SHIPMENTS_COLUMNS.length);
  SpreadsheetApp.flush();
  Logger.log("Shipments built: %s template rows.", n);
}

// ---------------------------------------------------------------------------
// Sales Plan (one tab per warehouse)
// ---------------------------------------------------------------------------
//
// Same date grid as StockModel and Daily COGS (START_DATE, HORIZON_DAYS), so
// rows align by date. Plan/day is the forecast; Actual/day is the real daily sold
// qty. This tab is the ONLY place a demand assumption lives: StockModel never
// assumes a rate, it reads Actual/day for past dates and Plan/day for future ones.
//
// "Cumulative Sold Before" = running sum of Actual/day up to (not including) this
// row's date. Actuals only, never Plan, because COGS must never imply units sold
// that have not actually sold.

function salesPlanWideCol_(sku, sub) { return sku + " " + sub; }

const SALES_PLAN_SUBCOLS = ["Plan/day", "Actual/day", "Delta/day", "Cumulative Sold Before"];

const SALES_PLAN_COLUMNS = ["Date"].concat(
  SKUS.reduce(function (cols, sku) {
    SALES_PLAN_SUBCOLS.forEach(function (sub) { cols.push(salesPlanWideCol_(sku, sub)); });
    return cols;
  }, [])
).concat(["Plan Revenue €/day", "Actual Revenue €/day", "Delta Revenue €/day"]);

function salesPlanSheetName_(warehouse) { return "Sales Plan — " + warehouse; }

function buildSalesPlan(warehouse) {
  if (WAREHOUSES.indexOf(warehouse) === -1) throw new Error("Unknown warehouse: " + warehouse);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss.getSheetByName("SKU Master")) {
    throw new Error("Build SKU Master first: run buildSkuMasterAndVendors().");
  }

  const sh = rebuildTab_(ss, salesPlanSheetName_(warehouse));
  writeTable_(sh, SALES_PLAN_COLUMNS, []);

  const n = HORIZON_DAYS;
  writeDayGrid_(sh, 2, colIndexIn_(SALES_PLAN_COLUMNS, "Date"), colLetterIn_(SALES_PLAN_COLUMNS, "Date"), START_DATE, n);

  SKUS.forEach(function (sku) {
    const planCol = colIndexIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Plan/day"));
    const actualCol = colIndexIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day"));
    const actualLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day"));
    const planLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Plan/day"));
    const deltaCol = colIndexIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Delta/day"));

    sh.getRange(2, planCol, n, 1).setBackground(SCT_YELLOW).setNumberFormat("#,##0");
    sh.getRange(2, actualCol, n, 1).setBackground(SCT_YELLOW).setNumberFormat("#,##0");

    const deltaFormulas = [];
    for (let r = 2; r <= n + 1; r++) {
      deltaFormulas.push([`=IF(OR(${actualLetter}${r}="",${planLetter}${r}=""),"",${actualLetter}${r}-${planLetter}${r})`]);
    }
    sh.getRange(2, deltaCol, n, 1).setFormulas(deltaFormulas).setNumberFormat("#,##0;[red]-#,##0");

    // Running total anchored at 0; N() lets a not-yet-entered Actual/day read as 0.
    const cumSoldLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Cumulative Sold Before"));
    const cumSoldCol = colIndexIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Cumulative Sold Before"));
    const cumSoldFormulas = [];
    for (let r = 2; r <= n + 1; r++) {
      cumSoldFormulas.push([r === 2 ? "=0" : `=${cumSoldLetter}${r - 1}+N(${actualLetter}${r - 1})`]);
    }
    sh.getRange(2, cumSoldCol, n, 1).setFormulas(cumSoldFormulas).setNumberFormat("#,##0");
  });

  // Revenue: one blended figure per day (not per SKU), used by Cashflow to
  // estimate fulfillment and carrier costs as a rate per 1000 of revenue.
  const planRevCol = colIndexIn_(SALES_PLAN_COLUMNS, "Plan Revenue €/day");
  const actualRevCol = colIndexIn_(SALES_PLAN_COLUMNS, "Actual Revenue €/day");
  const actualRevLetter = colLetterIn_(SALES_PLAN_COLUMNS, "Actual Revenue €/day");
  const planRevLetter = colLetterIn_(SALES_PLAN_COLUMNS, "Plan Revenue €/day");
  const deltaRevCol = colIndexIn_(SALES_PLAN_COLUMNS, "Delta Revenue €/day");

  sh.getRange(2, planRevCol, n, 1).setBackground(SCT_YELLOW).setNumberFormat("#,##0.00");
  sh.getRange(2, actualRevCol, n, 1).setBackground(SCT_YELLOW).setNumberFormat("#,##0.00");

  const deltaRevFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    deltaRevFormulas.push([`=IF(OR(${actualRevLetter}${r}="",${planRevLetter}${r}=""),"",${actualRevLetter}${r}-${planRevLetter}${r})`]);
  }
  sh.getRange(2, deltaRevCol, n, 1).setFormulas(deltaRevFormulas).setNumberFormat("#,##0.00;[red]-#,##0.00");

  SpreadsheetApp.flush();
  Logger.log("Sales Plan (%s) built: %s day rows.", warehouse, n);
}

function buildSalesPlanDE() { buildSalesPlan("WH-DE"); }
function buildSalesPlanCN() { buildSalesPlan("WH-CN"); }

// ---------------------------------------------------------------------------
// Fulfillment & Carrier Actuals
// ---------------------------------------------------------------------------
//
// Real per-date, per-warehouse 3PL and carrier invoice amounts. For past dates
// these override Cashflow's rate-based estimate, per warehouse.

const FULFILLMENT_CARRIER_SHEET_NAME = "Fulfillment & Carrier Actuals";
const FULFILLMENT_CARRIER_TEMPLATE_ROWS = 400;

const FULFILLMENT_CARRIER_COLUMNS = ["Date", "Warehouse", "Fulfillment Actual Amount", "Carrier Actual Amount"];

function fulfillmentCarrierCol_(name) { return colLetterIn_(FULFILLMENT_CARRIER_COLUMNS, name); }
function fulfillmentCarrierColIndex_(name) { return colIndexIn_(FULFILLMENT_CARRIER_COLUMNS, name); }

function buildFulfillmentCarrierActuals() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = rebuildTab_(ss, FULFILLMENT_CARRIER_SHEET_NAME);
  writeTable_(sh, FULFILLMENT_CARRIER_COLUMNS, []);

  const n = FULFILLMENT_CARRIER_TEMPLATE_ROWS;
  sh.getRange(2, fulfillmentCarrierColIndex_("Date"), n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(false).build()
  ).setNumberFormat("yyyy-mm-dd");
  sh.getRange(2, fulfillmentCarrierColIndex_("Warehouse"), n, 1).setDataValidation(
    SpreadsheetApp.newDataValidation().requireValueInList(WAREHOUSES, true).setAllowInvalid(false).build()
  );
  ["Fulfillment Actual Amount", "Carrier Actual Amount"].forEach(function (name) {
    sh.getRange(2, fulfillmentCarrierColIndex_(name), n, 1).setNumberFormat("#,##0.00");
  });
  sh.getRange(2, 1, n, FULFILLMENT_CARRIER_COLUMNS.length).setBackground(SCT_YELLOW);

  sh.autoResizeColumns(1, FULFILLMENT_CARRIER_COLUMNS.length);
  SpreadsheetApp.flush();
  Logger.log("Fulfillment & Carrier Actuals built: %s template rows.", n);
}

// ---------------------------------------------------------------------------
// Inventory Ledger: per-batch FIFO landed cost, live formulas
// ---------------------------------------------------------------------------
//
// One Batch Detail row per Shipments row (1:1, same order), wide by SKU. A batch
// exists only once it has physically landed (Actual Arrival Date filled): an
// in-transit or planned shipment must not read as received stock, and with a
// blank Actual Warehouse it would otherwise pool both warehouses into one bucket.
//
// FIFO without sorting rows: for each batch,
//   Cumulative Received = qty of this SKU landed at this warehouse up to and
//                          including this batch (same-day ties broken by row order)
//   Qty Remaining       = MAX(0, MIN(batch qty, Cumulative Received - total sold))
// so the oldest batches are consumed first and the newest keep stock. Because it
// is formulas over TODAY(), the ledger updates the moment a new Actual/day is
// entered, with no script re-run.
//
// Landed Cost/unit = PO's Full Factory Cost/unit for this PO#+SKU + this
// shipment's Net Cost/unit. A row without a PO# (e.g. an internal transfer) uses
// factory cost 0; a PO# whose price is still unknown resolves blank, never a
// fabricated 0.
//
// The StockModel has a manual opening-stock seed (row 2); the Ledger has none and
// knows only what Shipments contains. If an opening seed is used, account for it
// before treating a StockModel vs Ledger gap as a bug.

const INV_LEDGER_SHEET_NAME = "Inventory Ledger";

const INV_LEDGER_SUMMARY_COLUMNS = [
  "Warehouse", "SKU", "Qty On Hand", "Value On Hand", "Weighted Cost/unit", "Oversold Qty",
];

const INV_LEDGER_SKU_SUBCOLS = ["Qty Received", "Cumulative Received", "Landed Cost/unit", "Qty Remaining", "Value Remaining"];
function invLedgerSkuCol_(sku, sub) { return sku + " " + sub; }
const INV_LEDGER_BATCH_COLUMNS = ["Shipment ID", "Warehouse", "Date Landed"].concat(
  SKUS.reduce(function (cols, sku) {
    INV_LEDGER_SKU_SUBCOLS.forEach(function (sub) { cols.push(invLedgerSkuCol_(sku, sub)); });
    return cols;
  }, [])
);

function invLedgerBatchCol_(name) { return colLetterIn_(INV_LEDGER_BATCH_COLUMNS, name); }
function invLedgerBatchColIndex_(name) { return colIndexIn_(INV_LEDGER_BATCH_COLUMNS, name); }

// Picks this row's Sales Plan tab by its warehouse value, via INDIRECT, so the
// reference stays correct if the warehouse on a row is edited later.
function salesPlanTabRefExpr_(warehouseCellExpr) {
  return `"'"&IF(${warehouseCellExpr}="${WAREHOUSES[0]}","${salesPlanSheetName_(WAREHOUSES[0])}","${salesPlanSheetName_(WAREHOUSES[1])}")&"'!"`;
}

function buildInventoryLedger() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [PO_SHEET_NAME, SHIPMENTS_SHEET_NAME].concat(WAREHOUSES.map(salesPlanSheetName_)).forEach(function (name) {
    if (!ss.getSheetByName(name)) throw new Error("Build " + name + " first.");
  });

  const sh = rebuildTab_(ss, INV_LEDGER_SHEET_NAME);

  const shLastRow = SHIP_TEMPLATE_ROWS + 1;
  const shIdLetter = shipCol_("Shipment ID");
  const shPoLetter = shipCol_("PO#");
  const shWhLetter = shipCol_("Actual Warehouse");
  const shArrLetter = shipCol_("Actual Arrival Date");
  const shNetLetter = shipCol_("Net Cost/unit");

  const poLastRow = PO_TEMPLATE_ROWS + 1;
  const poNumLetter = poCol_("PO#");
  const poSkuLetter = poCol_("SKU");
  const poFactoryCostLetter = poCol_("Full Factory Cost/unit");

  const spDateLetter = colLetterIn_(SALES_PLAN_COLUMNS, "Date");
  const spLastRow = HORIZON_DAYS + 1;

  const titleRow = 1;
  const headerRow = 2;
  const firstDataRow = headerRow + 1;
  ensureGrid_(sh, firstDataRow + SHIP_TEMPLATE_ROWS, INV_LEDGER_BATCH_COLUMNS.length + 2 + INV_LEDGER_SUMMARY_COLUMNS.length);
  sh.getRange(titleRow, 1).setFormula('="CURRENT ON-HAND (as of "&TEXT(TODAY(),"yyyy-mm-dd")&")"')
    .setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);
  sh.getRange(headerRow, 1, 1, INV_LEDGER_BATCH_COLUMNS.length).setValues([INV_LEDGER_BATCH_COLUMNS])
    .setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);
  sh.setFrozenRows(headerRow);

  const idFormulas = [], whFormulas = [], dateFormulas = [];
  for (let i = 0; i < SHIP_TEMPLATE_ROWS; i++) {
    const r = i + 2; // Shipments row aligned with this Ledger row
    idFormulas.push([`=IF(Shipments!$${shIdLetter}$${r}="","",Shipments!$${shIdLetter}$${r})`]);
    whFormulas.push([`=IF(Shipments!$${shWhLetter}$${r}="","",Shipments!$${shWhLetter}$${r})`]);
    dateFormulas.push([`=IF(Shipments!$${shArrLetter}$${r}="","",Shipments!$${shArrLetter}$${r})`]);
  }
  sh.getRange(firstDataRow, invLedgerBatchColIndex_("Shipment ID"), SHIP_TEMPLATE_ROWS, 1).setFormulas(idFormulas);
  sh.getRange(firstDataRow, invLedgerBatchColIndex_("Warehouse"), SHIP_TEMPLATE_ROWS, 1).setFormulas(whFormulas);
  sh.getRange(firstDataRow, invLedgerBatchColIndex_("Date Landed"), SHIP_TEMPLATE_ROWS, 1).setFormulas(dateFormulas).setNumberFormat("yyyy-mm-dd");

  SKUS.forEach(function (sku) {
    const shSkuLetter = shipCol_(shipSkuQtyCol_(sku));
    const cumRecLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Cumulative Received"));
    const remLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Qty Remaining"));
    const costLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Landed Cost/unit"));
    const spActualLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day"));

    const qtyFormulas = [], cumRecFormulas = [], costFormulas = [], remFormulas = [], valFormulas = [];
    for (let i = 0; i < SHIP_TEMPLATE_ROWS; i++) {
      const r = i + 2;
      const ledgerRow = firstDataRow + i;
      const thisQty = `Shipments!$${shSkuLetter}$${r}`;
      const hasLanded = `AND(${thisQty}<>"",${thisQty}>0,Shipments!$${shArrLetter}$${r}<>"")`;

      qtyFormulas.push([`=IF(${hasLanded},${thisQty},"")`]);

      const factoryCostLookup =
        `IFERROR(INDEX('${PO_SHEET_NAME}'!$${poFactoryCostLetter}$2:$${poFactoryCostLetter}$${poLastRow},` +
        `MATCH(1,('${PO_SHEET_NAME}'!$${poNumLetter}$2:$${poNumLetter}$${poLastRow}=Shipments!$${shPoLetter}$${r})*` +
        `('${PO_SHEET_NAME}'!$${poSkuLetter}$2:$${poSkuLetter}$${poLastRow}="${sku}"),0)),` +
        `IF(Shipments!$${shPoLetter}$${r}="",0,""))`;
      const netCostRef = `IF(Shipments!$${shNetLetter}$${r}="",0,Shipments!$${shNetLetter}$${r})`;
      costFormulas.push([`=IF(${hasLanded},IFERROR(${factoryCostLookup}+${netCostRef},""),"")`]);

      const cumulativeReceived =
        `SUMPRODUCT((Shipments!$${shSkuLetter}$2:$${shSkuLetter}$${shLastRow}>0)*` +
        `(Shipments!$${shWhLetter}$2:$${shWhLetter}$${shLastRow}=Shipments!$${shWhLetter}$${r})*` +
        `((Shipments!$${shArrLetter}$2:$${shArrLetter}$${shLastRow}<Shipments!$${shArrLetter}$${r})+` +
        `((Shipments!$${shArrLetter}$2:$${shArrLetter}$${shLastRow}=Shipments!$${shArrLetter}$${r})*` +
        `(ROW(Shipments!$${shArrLetter}$2:$${shArrLetter}$${shLastRow})<=ROW(Shipments!$${shArrLetter}$${r}))))*` +
        `Shipments!$${shSkuLetter}$2:$${shSkuLetter}$${shLastRow})`;
      cumRecFormulas.push([`=IF(${hasLanded},${cumulativeReceived},"")`]);

      const spTabRef = salesPlanTabRefExpr_(`Shipments!$${shWhLetter}$${r}`);
      const totalSold =
        `SUMPRODUCT((INDIRECT(${spTabRef}&"${spDateLetter}2:${spDateLetter}${spLastRow}")<TODAY())*` +
        `INDIRECT(${spTabRef}&"${spActualLetter}2:${spActualLetter}${spLastRow}"))`;

      remFormulas.push([`=IF(${hasLanded},MAX(0,MIN(${thisQty},$${cumRecLetter}$${ledgerRow}-${totalSold})),"")`]);
      valFormulas.push([`=IF(OR($${remLetter}$${ledgerRow}="",$${costLetter}$${ledgerRow}=""),"",$${remLetter}$${ledgerRow}*$${costLetter}$${ledgerRow})`]);
    }

    sh.getRange(firstDataRow, invLedgerBatchColIndex_(invLedgerSkuCol_(sku, "Qty Received")), SHIP_TEMPLATE_ROWS, 1).setFormulas(qtyFormulas).setNumberFormat("#,##0");
    sh.getRange(firstDataRow, invLedgerBatchColIndex_(invLedgerSkuCol_(sku, "Cumulative Received")), SHIP_TEMPLATE_ROWS, 1).setFormulas(cumRecFormulas).setNumberFormat("#,##0");
    sh.getRange(firstDataRow, invLedgerBatchColIndex_(invLedgerSkuCol_(sku, "Landed Cost/unit")), SHIP_TEMPLATE_ROWS, 1).setFormulas(costFormulas).setNumberFormat("#,##0.0000");
    sh.getRange(firstDataRow, invLedgerBatchColIndex_(invLedgerSkuCol_(sku, "Qty Remaining")), SHIP_TEMPLATE_ROWS, 1).setFormulas(remFormulas).setNumberFormat("#,##0");
    sh.getRange(firstDataRow, invLedgerBatchColIndex_(invLedgerSkuCol_(sku, "Value Remaining")), SHIP_TEMPLATE_ROWS, 1).setFormulas(valFormulas).setNumberFormat("#,##0.00");
  });

  // Summary (current on-hand per warehouse x SKU), to the right of Batch Detail.
  const summaryStartCol = INV_LEDGER_BATCH_COLUMNS.length + 2;
  sh.getRange(headerRow, summaryStartCol, 1, INV_LEDGER_SUMMARY_COLUMNS.length).setValues([INV_LEDGER_SUMMARY_COLUMNS])
    .setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);

  const summaryRows = [];
  WAREHOUSES.forEach(function (warehouse) {
    SKUS.forEach(function (sku) { summaryRows.push([warehouse, sku]); });
  });
  sh.getRange(firstDataRow, summaryStartCol, summaryRows.length, 2).setValues(summaryRows);

  const summaryWhLetter = columnToLetter_(summaryStartCol);
  const qtyOnHandCol = summaryStartCol + colIndexIn_(INV_LEDGER_SUMMARY_COLUMNS, "Qty On Hand") - 1;
  const valueOnHandCol = summaryStartCol + colIndexIn_(INV_LEDGER_SUMMARY_COLUMNS, "Value On Hand") - 1;
  const weightedCostCol = summaryStartCol + colIndexIn_(INV_LEDGER_SUMMARY_COLUMNS, "Weighted Cost/unit") - 1;
  const oversoldCol = summaryStartCol + colIndexIn_(INV_LEDGER_SUMMARY_COLUMNS, "Oversold Qty") - 1;
  const qtyOnHandLetter = columnToLetter_(qtyOnHandCol);
  const valueOnHandLetter = columnToLetter_(valueOnHandCol);
  const whDetailLetter = invLedgerBatchCol_("Warehouse");
  const lastDetailRow = firstDataRow + SHIP_TEMPLATE_ROWS - 1;

  const qtySummary = [], valSummary = [], costSummary = [], oversoldSummary = [];
  summaryRows.forEach(function (rowVals, i) {
    const r = firstDataRow + i;
    const sku = rowVals[1];
    const remLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Qty Remaining"));
    const valLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Value Remaining"));
    const skuQtyLetter = shipCol_(shipSkuQtyCol_(sku));
    const spActualLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day"));

    qtySummary.push([`=SUMIF($${whDetailLetter}$${firstDataRow}:$${whDetailLetter}$${lastDetailRow},$${summaryWhLetter}${r},$${remLetter}$${firstDataRow}:$${remLetter}$${lastDetailRow})`]);
    valSummary.push([`=SUMIF($${whDetailLetter}$${firstDataRow}:$${whDetailLetter}$${lastDetailRow},$${summaryWhLetter}${r},$${valLetter}$${firstDataRow}:$${valLetter}$${lastDetailRow})`]);
    costSummary.push([`=IF($${qtyOnHandLetter}${r}>0,$${valueOnHandLetter}${r}/$${qtyOnHandLetter}${r},"")`]);

    // Oversold = units sold to date beyond everything ever landed: a signal of
    // missing shipment data or genuine backorders.
    const spTabRef = salesPlanTabRefExpr_(`$${summaryWhLetter}${r}`);
    const totalSold = `SUMPRODUCT((INDIRECT(${spTabRef}&"${spDateLetter}2:${spDateLetter}${spLastRow}")<TODAY())*INDIRECT(${spTabRef}&"${spActualLetter}2:${spActualLetter}${spLastRow}"))`;
    const totalReceived = `SUMIFS(Shipments!$${skuQtyLetter}$2:$${skuQtyLetter}$${shLastRow},Shipments!$${shWhLetter}$2:$${shWhLetter}$${shLastRow},$${summaryWhLetter}${r})`;
    oversoldSummary.push([`=MAX(0,${totalSold}-${totalReceived})`]);
  });

  sh.getRange(firstDataRow, qtyOnHandCol, summaryRows.length, 1).setFormulas(qtySummary).setNumberFormat("#,##0");
  sh.getRange(firstDataRow, valueOnHandCol, summaryRows.length, 1).setFormulas(valSummary).setNumberFormat("#,##0.00");
  sh.getRange(firstDataRow, weightedCostCol, summaryRows.length, 1).setFormulas(costSummary).setNumberFormat("#,##0.0000");
  sh.getRange(firstDataRow, oversoldCol, summaryRows.length, 1).setFormulas(oversoldSummary).setNumberFormat("#,##0;[red]#,##0");

  sh.autoResizeColumns(1, summaryStartCol + INV_LEDGER_SUMMARY_COLUMNS.length - 1);
  SpreadsheetApp.flush();
  Logger.log("Inventory Ledger built: %s batch rows, %s summary rows.", SHIP_TEMPLATE_ROWS, summaryRows.length);
}

// ---------------------------------------------------------------------------
// StockModel (one tab per warehouse): day-by-day stock and days of supply
// ---------------------------------------------------------------------------
//
// One row per date, a column group per SKU:
//   IN         qty landing that day (landed by Actual Arrival, or still expected
//              by Effective ETA); a shipment whose ETA passed without an arrival
//              keeps showing on its ETA instead of silently vanishing
//   Batch      which Shipment IDs make up IN
//   /day       Actual/day for past dates, Plan/day for today and future
//   Stock      balance at the START of the day (row 2 = manual opening seed)
//   Trans d    days until the nearest in-transit shipment lands
//   Prod d     days until the nearest open PO is ready at the factory
//   Stk d      how many future days current stock covers, walking the actual
//              future daily rates (not a flat stock / today's-rate division)
//   Pipeline d Stk d + Trans d + Prod d
// Warehouses are never blended: each tab reads only its own Sales Plan tab and
// only shipments planned for / landed at its own warehouse.

const STOCK_SUBCOLS = ["IN", "Batch", "/day", "Stock", "Trans d", "Prod d", "Stk d", "Pipeline d"];

function stockCol_(sku, sub) { return sku + " " + sub; }

const STOCK_MODEL_COLUMNS = ["Date"].concat(
  SKUS.reduce(function (cols, sku) {
    STOCK_SUBCOLS.forEach(function (sub) { cols.push(stockCol_(sku, sub)); });
    return cols;
  }, [])
).concat(["Status"]);

function stockModelSheetName_(warehouse) { return "StockModel — " + warehouse; }

function buildStockModel(warehouse) {
  if (WAREHOUSES.indexOf(warehouse) === -1) throw new Error("Unknown warehouse: " + warehouse);
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHIPMENTS_SHEET_NAME, PO_SHEET_NAME, salesPlanSheetName_(warehouse)].forEach(function (name) {
    if (!ss.getSheetByName(name)) throw new Error("Build " + name + " first.");
  });

  const sh = rebuildTab_(ss, stockModelSheetName_(warehouse));
  // Sidebar boxes sit one blank column past the table and are 4 columns wide.
  const sideStartCol = STOCK_MODEL_COLUMNS.length + 2;
  ensureGrid_(sh, HORIZON_DAYS + 1, sideStartCol + 3);
  writeTable_(sh, STOCK_MODEL_COLUMNS, []);

  const n = HORIZON_DAYS;
  const dateLetter = colLetterIn_(STOCK_MODEL_COLUMNS, "Date");
  writeDayGrid_(sh, 2, colIndexIn_(STOCK_MODEL_COLUMNS, "Date"), dateLetter, START_DATE, n);

  const shipLastRow = SHIP_TEMPLATE_ROWS + 1;
  const shipWhLetter = shipCol_("Actual Warehouse");
  const shipArrLetter = shipCol_("Actual Arrival Date");
  const shipIdLetter = shipCol_("Shipment ID");
  const etaLetter = shipCol_("Effective ETA");
  const plannedWhLetter = shipCol_("Planned Warehouse");
  // Actual depart, not planned: "has it left yet" needs a real event.
  const departLetter = shipCol_("Actual Depart Date");

  const salesPlanSh = salesPlanSheetName_(warehouse);
  const salesPlanDateCol = colLetterIn_(SALES_PLAN_COLUMNS, "Date");
  const salesPlanLastRow = HORIZON_DAYS + 1;

  const poLastRow = PO_TEMPLATE_ROWS + 1;
  const poPlannedReadyLetter = poCol_("Planned Ready Date");
  const poSkuLetter = poCol_("SKU");
  const poPlannedQtyLetter = poCol_("Planned Qty — " + warehouse);
  const poQtyRemainingLetter = poCol_("Qty Remaining to Produce");

  SKUS.forEach(function (sku) {
    const letterOf = function (sub) { return colLetterIn_(STOCK_MODEL_COLUMNS, stockCol_(sku, sub)); };
    const indexOf = function (sub) { return colIndexIn_(STOCK_MODEL_COLUMNS, stockCol_(sku, sub)); };
    const inLetter = letterOf("IN"), dayLetter = letterOf("/day"), stockLetter = letterOf("Stock");
    const transLetter = letterOf("Trans d"), prodLetter = letterOf("Prod d"), stkLetter = letterOf("Stk d");

    const salesPlanDayCol = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Plan/day"));
    const salesPlanActualCol = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day"));
    const shipSkuQtyLetter = shipCol_(shipSkuQtyCol_(sku));

    const shipQtyRange = `Shipments!$${shipSkuQtyLetter}$2:$${shipSkuQtyLetter}$${shipLastRow}`;
    const shipWhRange = `Shipments!$${shipWhLetter}$2:$${shipWhLetter}$${shipLastRow}`;
    const shipArrRange = `Shipments!$${shipArrLetter}$2:$${shipArrLetter}$${shipLastRow}`;
    const shipPlannedWhRange = `Shipments!$${plannedWhLetter}$2:$${plannedWhLetter}$${shipLastRow}`;
    const shipEtaRange = `Shipments!$${etaLetter}$2:$${etaLetter}$${shipLastRow}`;
    const shipIdRange = `Shipments!$${shipIdLetter}$2:$${shipIdLetter}$${shipLastRow}`;
    const depRange = `Shipments!$${departLetter}$2:$${departLetter}$${shipLastRow}`;

    const inFormulas = [], batchFormulas = [], dayFormulas = [], stockFormulas = [];
    const transFormulas = [], prodFormulas = [], stkFormulas = [], pipeFormulas = [];

    for (let r = 2; r <= n + 1; r++) {
      const d = `${dateLetter}${r}`;

      // Landed term matches Actual Warehouse + Actual Arrival; planned term
      // matches Planned Warehouse + ETA and only while Actual Arrival is blank, so
      // the two can never double-count. Both apply to past and future dates alike.
      const landedTerm = `SUMIFS(${shipQtyRange},${shipWhRange},"${warehouse}",${shipArrRange},${d})`;
      const plannedTerm = `SUMIFS(${shipQtyRange},${shipPlannedWhRange},"${warehouse}",${shipArrRange},"",${shipEtaRange},${d})`;
      inFormulas.push([`=${landedTerm}+${plannedTerm}`]);

      // Batch uses the same two conditions as IN so the label and qty never drift.
      const landedBatchArr = `IF((${shipQtyRange}>0)*(${shipWhRange}="${warehouse}")*(${shipArrRange}=${d}),${shipIdRange},"")`;
      const plannedBatchArr = `IF((${shipQtyRange}>0)*(${shipPlannedWhRange}="${warehouse}")*(${shipArrRange}="")*(${shipEtaRange}=${d}),${shipIdRange},"")`;
      batchFormulas.push([`=IFERROR(TEXTJOIN(", ",TRUE,ARRAYFORMULA(${landedBatchArr}),ARRAYFORMULA(${plannedBatchArr})),"")`]);

      // Rolling actual/plan boundary: switches automatically as TODAY() advances.
      dayFormulas.push([`=IF(${d}<TODAY(),IFERROR(INDEX('${salesPlanSh}'!$${salesPlanActualCol}$2:$${salesPlanActualCol}$${salesPlanLastRow},MATCH(${d},'${salesPlanSh}'!$${salesPlanDateCol}$2:$${salesPlanDateCol}$${salesPlanLastRow},0)),0),IFERROR(INDEX('${salesPlanSh}'!$${salesPlanDayCol}$2:$${salesPlanDayCol}$${salesPlanLastRow},MATCH(${d},'${salesPlanSh}'!$${salesPlanDateCol}$2:$${salesPlanDateCol}$${salesPlanLastRow},0)),0))`]);

      // Stock[r] = Stock[r-1] + IN[r-1] - sold[r-1]: stock landing on day r is
      // available from day r+1. Row 2 is the manual opening seed.
      stockFormulas.push([r === 2 ? 0 : `=${stockLetter}${r - 1}+${inLetter}${r - 1}-${dayLetter}${r - 1}`]);

      // Trans d tests THIS ROW's date against each shipment's real transit window
      // [Actual Depart, effArr), where effArr = Actual Arrival if landed, else
      // Effective ETA. Testing "not arrived yet" instead would describe today's
      // state and zero out every historical row once a shipment lands.
      // Two tiers: 0 when nothing is in transit; blank (unknown) when something
      // departed but has neither an arrival nor an ETA; otherwise the minimum
      // days to arrival. Filtered by Planned Warehouse, since an in-transit
      // shipment has no Actual Warehouse yet.
      const effArrExpr = `IF(${shipArrRange}<>"",${shipArrRange},${shipEtaRange})`;
      const qtyOkExpr = `${shipQtyRange}>0`;
      const whOkExpr = `${shipPlannedWhRange}="${warehouse}"`;
      const knownExpr = `SUMPRODUCT((${qtyOkExpr})*(${whOkExpr})*(${depRange}<>"")*(${depRange}<=${d})*((${effArrExpr})<>"")*((${effArrExpr})>${d}))`;
      const unkExpr = `SUMPRODUCT((${qtyOkExpr})*(${whOkExpr})*(${depRange}<>"")*(${depRange}<=${d})*((${effArrExpr})=""))`;
      const filterExpr = `FILTER((${effArrExpr})-${d},${qtyOkExpr},${whOkExpr},${depRange}<>"",${depRange}<=${d},(${effArrExpr})<>"",(${effArrExpr})>${d})`;
      transFormulas.push([`=IF(${knownExpr}>0,MAX(0,MIN(${filterExpr})),IF(${unkExpr}>0,"",0))`]);

      // Prod d: days until the earliest still-in-production PO for this SKU with a
      // planned qty for this warehouse is ready. Same two tiers: 0 when nothing is
      // in production; blank when something is but its Planned Ready Date is
      // unknown (rather than reading as "ready now").
      const poRange = function (letter) { return `'${PO_SHEET_NAME}'!$${letter}$2:$${letter}$${poLastRow}`; };
      prodFormulas.push([`=IF(COUNTIFS(${poRange(poSkuLetter)},"${sku}",${poRange(poPlannedQtyLetter)},">0",${poRange(poQtyRemainingLetter)},">0")=0,0,IFERROR(MAX(0,MIN(FILTER(${poRange(poPlannedReadyLetter)}-${d},${poRange(poSkuLetter)}="${sku}",${poRange(poPlannedQtyLetter)}>0,${poRange(poQtyRemainingLetter)}>0,${poRange(poPlannedReadyLetter)}<>""))),""))`]);

      // Stk d: SCAN builds the cumulative future consumption from this day on;
      // count the days whose cumulative stays within current stock.
      stkFormulas.push([`=IF(${stockLetter}${r}<=0,0,SUMPRODUCT(--(SCAN(0,${dayLetter}${r}:${dayLetter}$${n + 1},LAMBDA(a,x,a+x))<=${stockLetter}${r})))`]);

      // IFERROR: Trans d / Prod d may legitimately be blank (unknown).
      pipeFormulas.push([`=IFERROR(${stkLetter}${r}+${transLetter}${r}+${prodLetter}${r},"")`]);
    }

    sh.getRange(2, indexOf("IN"), n, 1).setFormulas(inFormulas).setNumberFormat("#,##0");
    sh.getRange(2, indexOf("Batch"), n, 1).setFormulas(batchFormulas);
    sh.getRange(2, indexOf("/day"), n, 1).setFormulas(dayFormulas).setNumberFormat("#,##0");
    sh.getRange(2, indexOf("Stock"), n, 1).setFormulas(stockFormulas).setNumberFormat("#,##0");
    sh.getRange(2, indexOf("Stock"), 1, 1).setBackground(SCT_YELLOW); // manual opening stock seed
    sh.getRange(2, indexOf("Trans d"), n, 1).setFormulas(transFormulas).setNumberFormat("0");
    sh.getRange(2, indexOf("Prod d"), n, 1).setFormulas(prodFormulas).setNumberFormat("0");
    sh.getRange(2, indexOf("Stk d"), n, 1).setFormulas(stkFormulas).setNumberFormat("0.0");
    sh.getRange(2, indexOf("Pipeline d"), n, 1).setFormulas(pipeFormulas).setNumberFormat("0.0");
  });

  // Status: worst Stk d across SKUs for that date ("worst wins").
  const statusCol = colIndexIn_(STOCK_MODEL_COLUMNS, "Status");
  const statusFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    const stkRefs = SKUS.map(function (sku) { return `${colLetterIn_(STOCK_MODEL_COLUMNS, stockCol_(sku, "Stk d"))}${r}`; });
    const minStk = `MIN(${stkRefs.join(",")})`;
    statusFormulas.push([
      `=IFERROR(IF(${minStk}<=0,"🚨 STOCKOUT",IF(${minStk}<21,"🔴 CRITICAL",IF(${minStk}<45,"🟡 LOW",IF(${minStk}>90,"🔵 OVERSTOCK","🟢 OK")))),"")`
    ]);
  }
  sh.getRange(2, statusCol, n, 1).setFormulas(statusFormulas);

  // Thresholds: <21 days red, 21-45 yellow, 45-90 green, >90 blue (overstock).
  const statusRange = sh.getRange(2, statusCol, n, 1);
  const statusColorRules = [
    { text: "🚨 STOCKOUT", color: "#f8d7da" },
    { text: "🔴 CRITICAL", color: "#f8d7da" },
    { text: "🟡 LOW", color: "#fff3cd" },
    { text: "🟢 OK", color: "#d4edda" },
    { text: "🔵 OVERSTOCK", color: "#cce5ff" },
  ].map(function (rule) {
    return SpreadsheetApp.newConditionalFormatRule()
      .whenTextEqualTo(rule.text)
      .setBackground(rule.color)
      .setRanges([statusRange])
      .build();
  });

  // Same thresholds on every per-SKU day-count column, as custom-formula rules so
  // each boundary is one readable expression rather than rule-order dependent.
  const dayColorRules = [];
  SKUS.forEach(function (sku) {
    ["Trans d", "Prod d", "Stk d", "Pipeline d"].forEach(function (colName) {
      const letter = colLetterIn_(STOCK_MODEL_COLUMNS, stockCol_(sku, colName));
      const dayRange = sh.getRange(2, colIndexIn_(STOCK_MODEL_COLUMNS, stockCol_(sku, colName)), n, 1);
      const cell = `$${letter}2`;
      [
        { formula: `=AND(${cell}<>"",${cell}<21)`, color: "#f8d7da" },
        { formula: `=AND(${cell}>=21,${cell}<45)`, color: "#fff3cd" },
        { formula: `=AND(${cell}>=45,${cell}<=90)`, color: "#d4edda" },
        { formula: `=AND(${cell}<>"",${cell}>90)`, color: "#cce5ff" },
      ].forEach(function (rule) {
        dayColorRules.push(
          SpreadsheetApp.newConditionalFormatRule()
            .whenFormulaSatisfied(rule.formula)
            .setBackground(rule.color)
            .setRanges([dayRange])
            .build()
        );
      });
    });
  });
  sh.setConditionalFormatRules(sh.getConditionalFormatRules().concat(statusColorRules, dayColorRules));

  // Sidebar 1: trailing 7-day run rate (ending yesterday) vs what the plan said
  // for the same window.
  sh.getRange(1, sideStartCol, 1, 4).setValues([["SKU", "Actual/day", "Plan/day", "Rate Δ%"]])
    .setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);
  const trailingAvg = function (colLetter) {
    const range = `'${salesPlanSh}'!$${colLetter}$2:$${colLetter}$${salesPlanLastRow}`;
    const dates = `'${salesPlanSh}'!$${salesPlanDateCol}$2:$${salesPlanDateCol}$${salesPlanLastRow}`;
    return `IFERROR(AVERAGE(INDEX(${range},MATCH(TODAY()-7,${dates},0)):INDEX(${range},MATCH(TODAY()-1,${dates},0))),0)`;
  };
  const rateRows = SKUS.map(function (sku) {
    const actual = trailingAvg(colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day")));
    const plan = trailingAvg(colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Plan/day")));
    return [sku, `=${actual}`, `=${plan}`, `=IFERROR((${actual}-${plan})/${plan},"")`];
  });
  sh.getRange(2, sideStartCol, SKUS.length, 1).setValues(rateRows.map(function (r) { return [r[0]]; }));
  sh.getRange(2, sideStartCol + 1, SKUS.length, 3).setFormulas(rateRows.map(function (r) { return [r[1], r[2], r[3]]; }));
  sh.getRange(2, sideStartCol + 3, SKUS.length, 1).setNumberFormat("0.0%");

  // Sidebar 2: today's coverage (Stk d / Trans d / Prod d) per SKU.
  const coverageStartRow = SKUS.length + 4;
  sh.getRange(coverageStartRow, sideStartCol, 1, 1 + SKUS.length).setValues(
    [["📦 Coverage today (days)"].concat(SKUS)]
  ).setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);
  [
    { label: "Warehouse (Stk d)", sub: "Stk d" },
    { label: "Transit (Trans d)", sub: "Trans d" },
    { label: "Production (Prod d)", sub: "Prod d" },
  ].forEach(function (row, i) {
    const r = coverageStartRow + 1 + i;
    sh.getRange(r, sideStartCol).setValue(row.label);
    SKUS.forEach(function (sku, j) {
      const subLetter = colLetterIn_(STOCK_MODEL_COLUMNS, stockCol_(sku, row.sub));
      sh.getRange(r, sideStartCol + 1 + j).setFormula(
        `=IFERROR(INDEX($${subLetter}$2:$${subLetter}$${n + 1},MATCH(TODAY(),$${dateLetter}$2:$${dateLetter}$${n + 1},0)),"")`
      );
    });
  });

  sh.autoResizeColumns(1, sideStartCol + 3);
  SpreadsheetApp.flush();
  Logger.log("StockModel (%s) built: %s day rows.", warehouse, n);
}

// Zero-argument wrappers: the Apps Script Run menu cannot pass parameters.
function buildStockModelDE() { buildStockModel("WH-DE"); }
function buildStockModelCN() { buildStockModel("WH-CN"); }

// ---------------------------------------------------------------------------
// Daily COGS: FIFO-priced cost of goods sold and opening inventory, per day
// ---------------------------------------------------------------------------
//
// Same date grid as Sales Plan, so row r here is the same date as row r there and
// formulas reference it by row, no lookup needed. Future dates read 0 sold: a COGS
// report never projects from Plan/day.
//
// For date d, SKU S, warehouse W, per batch (all Ledger batch rows at once):
//   clamp(x)  = MIN(MAX(x, 0), batch qty), written as nested IF because MIN/MAX
//               don't operate element-wise inside SUMPRODUCT
//   rawStart  = clamp(CumulativeReceived - CumulativeSoldBefore(d))
//   rawEnd    = clamp(CumulativeReceived - CumulativeSoldBefore(d) - sold(d))
//   consumed  = rawStart - rawEnd        (this batch's share of day d's sales)
//   COGS(d)   = SUM(consumed x batch Landed Cost/unit)
//   Opening(d)= SUM(landedBefore(d) x rawStart [x cost])
// Opening Qty/Value is gated by Date Landed < d: a batch that lands later was not
// physically on hand yet. The COGS consumption is deliberately NOT gated: a sale
// before any batch landed is a backorder fulfilled later from whichever batch
// covers it, so that batch's cost prices it. This also handles a batch landing on
// day d without producing negative consumption.
//
// "Unpriced Units (oversold)" = units sold that no batch can account for. They are
// excluded from COGS (never priced at a guessed rate) and flagged red.

const DAILY_COGS_SHEET_NAME = "Daily COGS";

function dailyCogsKey_(warehouse, sku) { return warehouse + " " + sku; }

const DAILY_COGS_COLUMNS = ["Date"].concat(
  WAREHOUSES.reduce(function (cols, wh) {
    SKUS.forEach(function (sku) {
      cols.push(dailyCogsKey_(wh, sku) + " Opening Qty");
      cols.push(dailyCogsKey_(wh, sku) + " Opening Value €");
      cols.push(dailyCogsKey_(wh, sku) + " Units Sold");
      cols.push(dailyCogsKey_(wh, sku) + " COGS €");
      cols.push(dailyCogsKey_(wh, sku) + " Unpriced Units (oversold)");
    });
    cols.push(wh + " Total Opening Value €");
    cols.push(wh + " Total COGS €");
    return cols;
  }, [])
).concat(["Grand Total Opening Value €", "Grand Total COGS €"]);

function dailyCogsColIndex_(name) { return colIndexIn_(DAILY_COGS_COLUMNS, name); }

function buildDailyCogs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [SHIPMENTS_SHEET_NAME, INV_LEDGER_SHEET_NAME].concat(WAREHOUSES.map(salesPlanSheetName_)).forEach(function (name) {
    if (!ss.getSheetByName(name)) throw new Error("Build " + name + " first.");
  });

  const sh = rebuildTab_(ss, DAILY_COGS_SHEET_NAME);
  ensureGrid_(sh, HORIZON_DAYS + 1, DAILY_COGS_COLUMNS.length);
  writeTable_(sh, DAILY_COGS_COLUMNS, []);

  const n = HORIZON_DAYS;
  const dateLetter = colLetterIn_(DAILY_COGS_COLUMNS, "Date");
  writeDayGrid_(sh, 2, colIndexIn_(DAILY_COGS_COLUMNS, "Date"), dateLetter, START_DATE, n);

  const shLastRow = SHIP_TEMPLATE_ROWS + 1;
  const shWhLetter = shipCol_("Actual Warehouse");
  const ledgerFirstDataRow = 3; // Ledger: title row 1, header row 2
  const ledgerLastRow = ledgerFirstDataRow + SHIP_TEMPLATE_ROWS - 1;
  const dateLandedLetter = invLedgerBatchCol_("Date Landed");
  const dateLandedRange = `'${INV_LEDGER_SHEET_NAME}'!$${dateLandedLetter}$${ledgerFirstDataRow}:$${dateLandedLetter}$${ledgerLastRow}`;

  WAREHOUSES.forEach(function (wh) {
    const spSheetName = salesPlanSheetName_(wh);
    const whFilter = `(Shipments!$${shWhLetter}$2:$${shWhLetter}$${shLastRow}="${wh}")`;
    const cogsLetters = [], openingLetters = [];

    SKUS.forEach(function (sku) {
      const shSkuLetter = shipCol_(shipSkuQtyCol_(sku));
      const cumRecLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Cumulative Received"));
      const costLetter = invLedgerBatchCol_(invLedgerSkuCol_(sku, "Landed Cost/unit"));
      const spCumSoldLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Cumulative Sold Before"));
      const spActualLetter = colLetterIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day"));

      const qtyRange = `N(Shipments!$${shSkuLetter}$2:$${shSkuLetter}$${shLastRow})`;
      const cumRecRange = `N('${INV_LEDGER_SHEET_NAME}'!$${cumRecLetter}$${ledgerFirstDataRow}:$${cumRecLetter}$${ledgerLastRow})`;
      const costRange = `N('${INV_LEDGER_SHEET_NAME}'!$${costLetter}$${ledgerFirstDataRow}:$${costLetter}$${ledgerLastRow})`;
      const clamp = function (x) { return `IF(${x}<0,0,IF(${x}>${qtyRange},${qtyRange},${x}))`; };

      const openingQtyFormulas = [], openingValFormulas = [], unitsFormulas = [], cogsFormulas = [], unpricedFormulas = [];
      for (let r = 2; r <= n + 1; r++) {
        const soldTodayRaw = `'${spSheetName}'!$${spActualLetter}$${r}`;
        const soldToday = `N(${soldTodayRaw})`;
        const cumSoldStart = `'${spSheetName}'!$${spCumSoldLetter}$${r}`;
        const cumSoldEnd = `(${cumSoldStart}+${soldToday})`;

        const landedGateStart = `(${dateLandedRange}<>"")*(${dateLandedRange}<${dateLetter}${r})`;
        const rawStart = clamp(`(${cumRecRange}-${cumSoldStart})`);
        const rawEnd = clamp(`(${cumRecRange}-${cumSoldEnd})`);
        const openingTerm = `${whFilter}*(${landedGateStart}*${rawStart})`;
        const consumedTerm = `${whFilter}*(${rawStart}-${rawEnd})`;

        openingQtyFormulas.push([`=SUMPRODUCT(${openingTerm})`]);
        openingValFormulas.push([`=ROUND(SUMPRODUCT(${openingTerm}*${costRange}),2)`]);
        unitsFormulas.push([`=${soldTodayRaw}`]);
        cogsFormulas.push([`=ROUND(SUMPRODUCT(${consumedTerm}*${costRange}),2)`]);
        unpricedFormulas.push([`=IF(${soldTodayRaw}="","",MAX(0,${soldToday}-SUMPRODUCT(${consumedTerm})))`]);
      }

      const key = dailyCogsKey_(wh, sku);
      sh.getRange(2, dailyCogsColIndex_(key + " Opening Qty"), n, 1).setFormulas(openingQtyFormulas).setNumberFormat("#,##0");
      sh.getRange(2, dailyCogsColIndex_(key + " Opening Value €"), n, 1).setFormulas(openingValFormulas).setNumberFormat("#,##0.00");
      sh.getRange(2, dailyCogsColIndex_(key + " Units Sold"), n, 1).setFormulas(unitsFormulas).setNumberFormat("#,##0");
      sh.getRange(2, dailyCogsColIndex_(key + " COGS €"), n, 1).setFormulas(cogsFormulas).setNumberFormat("#,##0.00");
      sh.getRange(2, dailyCogsColIndex_(key + " Unpriced Units (oversold)"), n, 1).setFormulas(unpricedFormulas).setNumberFormat("#,##0;[red]#,##0");

      cogsLetters.push(colLetterIn_(DAILY_COGS_COLUMNS, key + " COGS €"));
      openingLetters.push(colLetterIn_(DAILY_COGS_COLUMNS, key + " Opening Value €"));
    });

    const whTotalFormulas = [], whOpeningTotalFormulas = [];
    for (let r = 2; r <= n + 1; r++) {
      whTotalFormulas.push([`=${cogsLetters.map(function (l) { return l + r; }).join("+")}`]);
      whOpeningTotalFormulas.push([`=${openingLetters.map(function (l) { return l + r; }).join("+")}`]);
    }
    sh.getRange(2, dailyCogsColIndex_(wh + " Total COGS €"), n, 1).setFormulas(whTotalFormulas).setNumberFormat("#,##0.00");
    sh.getRange(2, dailyCogsColIndex_(wh + " Total Opening Value €"), n, 1).setFormulas(whOpeningTotalFormulas).setNumberFormat("#,##0.00");
  });

  const whTotalLetters = WAREHOUSES.map(function (wh) { return colLetterIn_(DAILY_COGS_COLUMNS, wh + " Total COGS €"); });
  const whOpeningTotalLetters = WAREHOUSES.map(function (wh) { return colLetterIn_(DAILY_COGS_COLUMNS, wh + " Total Opening Value €"); });
  const grandTotalFormulas = [], grandOpeningTotalFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    grandTotalFormulas.push([`=${whTotalLetters.map(function (l) { return l + r; }).join("+")}`]);
    grandOpeningTotalFormulas.push([`=${whOpeningTotalLetters.map(function (l) { return l + r; }).join("+")}`]);
  }
  sh.getRange(2, dailyCogsColIndex_("Grand Total COGS €"), n, 1).setFormulas(grandTotalFormulas).setNumberFormat("#,##0.00");
  sh.getRange(2, dailyCogsColIndex_("Grand Total Opening Value €"), n, 1).setFormulas(grandOpeningTotalFormulas).setNumberFormat("#,##0.00");

  const rules = sh.getConditionalFormatRules();
  WAREHOUSES.forEach(function (wh) {
    SKUS.forEach(function (sku) {
      rules.push(SpreadsheetApp.newConditionalFormatRule()
        .whenNumberGreaterThan(0)
        .setBackground("#f8d7da")
        .setRanges([sh.getRange(2, dailyCogsColIndex_(dailyCogsKey_(wh, sku) + " Unpriced Units (oversold)"), n, 1)])
        .build());
    });
  });
  sh.setConditionalFormatRules(rules);

  sh.autoResizeColumns(1, DAILY_COGS_COLUMNS.length);
  SpreadsheetApp.flush();
  Logger.log("Daily COGS built: %s day rows.", n);
}

// ---------------------------------------------------------------------------
// Landed Cost Summary: one readable row per landed shipment
// ---------------------------------------------------------------------------
//
// A re-projection of Purchase Orders + Shipments, not a second source of truth.
// One row per Shipments row (one SKU per row by convention; if a row ever carried
// two SKUs, the first SKU in SKUS order would be picked). Blank until the
// shipment has landed, same gate as the Inventory Ledger.

const LANDED_COST_SUMMARY_SHEET_NAME = "Landed Cost Summary";
const LANDED_COST_SUMMARY_COLUMNS = [
  "Shipment ID", "Warehouse", "SKU", "Qty",
  "Product Cost/unit", "Logistics Cost/unit", "Total Cost/unit",
  "Total Cost/unit (net of refundable VAT)",
];
function landedCostSummaryCol_(name) { return colLetterIn_(LANDED_COST_SUMMARY_COLUMNS, name); }
function landedCostSummaryColIndex_(name) { return colIndexIn_(LANDED_COST_SUMMARY_COLUMNS, name); }

function buildLandedCostSummary() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [PO_SHEET_NAME, SHIPMENTS_SHEET_NAME].forEach(function (name) {
    if (!ss.getSheetByName(name)) throw new Error("Build " + name + " first.");
  });

  const sh = rebuildTab_(ss, LANDED_COST_SUMMARY_SHEET_NAME);
  writeTable_(sh, LANDED_COST_SUMMARY_COLUMNS, []);

  const n = SHIP_TEMPLATE_ROWS;
  const shIdLetter = shipCol_("Shipment ID");
  const shWhLetter = shipCol_("Actual Warehouse");
  const shArrLetter = shipCol_("Actual Arrival Date");
  const shPoLetter = shipCol_("PO#");
  const shTotalQtyLetter = shipCol_("Total Qty");
  const shGrossLetter = shipCol_("Gross Cost/unit");
  const shNetLetter = shipCol_("Net Cost/unit");

  const poLastRow = PO_TEMPLATE_ROWS + 1;
  const poNumLetter = poCol_("PO#");
  const poSkuLetter = poCol_("SKU");
  const poFactoryCostLetter = poCol_("Full Factory Cost/unit");

  const skuLetter = landedCostSummaryCol_("SKU");
  const productLetter = landedCostSummaryCol_("Product Cost/unit");
  const logisticsLetter = landedCostSummaryCol_("Logistics Cost/unit");

  const idFormulas = [], whFormulas = [], skuFormulas = [], qtyFormulas = [];
  const productFormulas = [], logisticsFormulas = [], totalFormulas = [], totalNetFormulas = [];

  for (let r = 2; r <= n + 1; r++) {
    const hasLanded = `AND(Shipments!$${shTotalQtyLetter}$${r}<>"",Shipments!$${shTotalQtyLetter}$${r}>0,Shipments!$${shArrLetter}$${r}<>"")`;

    idFormulas.push([`=IF(${hasLanded},Shipments!$${shIdLetter}$${r},"")`]);
    whFormulas.push([`=IF(${hasLanded},Shipments!$${shWhLetter}$${r},"")`]);

    const skuPick = SKUS.map(function (sku) {
      return `IF(Shipments!$${shipCol_(shipSkuQtyCol_(sku))}$${r}>0,"${sku}",`;
    }).join("") + '""' + ")".repeat(SKUS.length);
    skuFormulas.push([`=IF(${hasLanded},${skuPick},"")`]);

    qtyFormulas.push([`=IF(${hasLanded},Shipments!$${shTotalQtyLetter}$${r},"")`]);

    // Product = factory-side cost only (EXW + per-unit fees). Logistics = this
    // shipment's Gross Cost/unit. The net column swaps Gross for Net, i.e. the
    // true cost once refundable import VAT is reclaimed.
    const factoryCostLookup =
      `IFERROR(INDEX('${PO_SHEET_NAME}'!$${poFactoryCostLetter}$2:$${poFactoryCostLetter}$${poLastRow},` +
      `MATCH(1,('${PO_SHEET_NAME}'!$${poNumLetter}$2:$${poNumLetter}$${poLastRow}=Shipments!$${shPoLetter}$${r})*` +
      `('${PO_SHEET_NAME}'!$${poSkuLetter}$2:$${poSkuLetter}$${poLastRow}=$${skuLetter}${r}),0)),"")`;
    productFormulas.push([`=IF(${hasLanded},${factoryCostLookup},"")`]);

    logisticsFormulas.push([`=IF(${hasLanded},IF(Shipments!$${shGrossLetter}$${r}="","",Shipments!$${shGrossLetter}$${r}),"")`]);

    totalFormulas.push([`=IF(OR($${productLetter}${r}="",$${logisticsLetter}${r}=""),"",$${productLetter}${r}+$${logisticsLetter}${r})`]);

    const netLogisticsRef = `IF(Shipments!$${shNetLetter}$${r}="","",Shipments!$${shNetLetter}$${r})`;
    totalNetFormulas.push([`=IF(OR($${productLetter}${r}="",(${netLogisticsRef})=""),"",$${productLetter}${r}+(${netLogisticsRef}))`]);
  }

  sh.getRange(2, landedCostSummaryColIndex_("Shipment ID"), n, 1).setFormulas(idFormulas);
  sh.getRange(2, landedCostSummaryColIndex_("Warehouse"), n, 1).setFormulas(whFormulas);
  sh.getRange(2, landedCostSummaryColIndex_("SKU"), n, 1).setFormulas(skuFormulas);
  sh.getRange(2, landedCostSummaryColIndex_("Qty"), n, 1).setFormulas(qtyFormulas).setNumberFormat("#,##0");
  sh.getRange(2, landedCostSummaryColIndex_("Product Cost/unit"), n, 1).setFormulas(productFormulas).setNumberFormat("#,##0.0000");
  sh.getRange(2, landedCostSummaryColIndex_("Logistics Cost/unit"), n, 1).setFormulas(logisticsFormulas).setNumberFormat("#,##0.0000");
  sh.getRange(2, landedCostSummaryColIndex_("Total Cost/unit"), n, 1).setFormulas(totalFormulas).setNumberFormat("#,##0.0000");
  sh.getRange(2, landedCostSummaryColIndex_("Total Cost/unit (net of refundable VAT)"), n, 1).setFormulas(totalNetFormulas).setNumberFormat("#,##0.0000");

  sh.autoResizeColumns(1, LANDED_COST_SUMMARY_COLUMNS.length);
  SpreadsheetApp.flush();
  Logger.log("Landed Cost Summary built: %s rows.", n);
}

// ---------------------------------------------------------------------------
// Cashflow: consolidated day-by-day cash out
// ---------------------------------------------------------------------------
//
// One row per day from CASHFLOW_START_DATE. Blocks:
//   Purchases per SKU  PO tranches (+ Inspection Amount, riding Payment 2's date),
//                      each on its Actual Date if paid, else its Planned Date
//   Logistics          per shipment, 3 independent tranches (Freight 1, Freight 2,
//                      Customs), each on actual-else-planned date. If the real
//                      customs amount is unknown, the Customs Estimated Amount is
//                      used on the customs planned date. If NOTHING is known yet,
//                      the ship method's historical cost/unit x Total Qty is
//                      placed on the planned departure date.
//   Fulfillment/Carrier revenue x rate per 1000 (per warehouse); for past dates a
//                      real invoice in Fulfillment & Carrier Actuals overrides
//                      the estimate for that warehouse only
// Every cell is a live formula. Rate cells sit at the top of the tab (yellow).

const CASHFLOW_SHEET_NAME = "Cashflow";
const CASHFLOW_COST_KINDS = ["Fulfillment", "Carrier"];

function cashflowRateLabel_(kind, wh) { return `${kind} €/1000€ rev — ${wh}`; }
function cashflowSkuPurchaseCol_(sku) { return sku + " Purchases"; }

const CASHFLOW_COLUMNS = ["Date"]
  .concat(SKUS.map(cashflowSkuPurchaseCol_))
  .concat(["Logistics"]).concat(CASHFLOW_COST_KINDS).concat(["Total", "Running Total"]);

function buildCashflow() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [PO_SHEET_NAME, SHIPMENTS_SHEET_NAME, FULFILLMENT_CARRIER_SHEET_NAME].concat(WAREHOUSES.map(salesPlanSheetName_)).forEach(function (name) {
    if (!ss.getSheetByName(name)) throw new Error("Build " + name + " first.");
  });

  const sh = rebuildTab_(ss, CASHFLOW_SHEET_NAME);

  // Rate cells, one per (kind, warehouse), in column B. Left blank: enter the
  // contracted or observed rates (seedDemoData fills synthetic ones).
  const rateCells = {};
  let rateRow = 1;
  WAREHOUSES.forEach(function (wh) {
    CASHFLOW_COST_KINDS.forEach(function (kind) {
      sh.getRange(rateRow, 1).setValue(cashflowRateLabel_(kind, wh)).setFontWeight("bold");
      sh.getRange(rateRow, 2).setBackground(SCT_YELLOW).setNumberFormat("#,##0.00");
      rateCells[kind + "|" + wh] = `$B$${rateRow}`;
      rateRow++;
    });
  });

  const tableStartRow = rateRow + 1;
  const n = CASHFLOW_HORIZON_DAYS;
  ensureGrid_(sh, tableStartRow + n, CASHFLOW_COLUMNS.length);
  sh.getRange(tableStartRow, 1, 1, CASHFLOW_COLUMNS.length).setValues([CASHFLOW_COLUMNS])
    .setFontWeight("bold").setBackground(SCT_HDR).setFontColor(SCT_HDR_FG);
  sh.setFrozenRows(tableStartRow);

  const firstDataRow = tableStartRow + 1;
  const lastDataRow = firstDataRow + n - 1;
  const dateLetter = colLetterIn_(CASHFLOW_COLUMNS, "Date");
  writeDayGrid_(sh, firstDataRow, colIndexIn_(CASHFLOW_COLUMNS, "Date"), dateLetter, CASHFLOW_START_DATE, n);

  const poRange = function (letter) { return `'${PO_SHEET_NAME}'!$${letter}$2:$${letter}$${PO_TEMPLATE_ROWS + 1}`; };
  const shRange = function (letter) { return `'${SHIPMENTS_SHEET_NAME}'!$${letter}$2:$${letter}$${SHIP_TEMPLATE_ROWS + 1}`; };
  const fdRange = function (letter) { return `'${FULFILLMENT_CARRIER_SHEET_NAME}'!$${letter}$2:$${letter}$${FULFILLMENT_CARRIER_TEMPLATE_ROWS + 1}`; };

  // Amount lands on its Actual date when paid, else on its Planned date. The two
  // SUMIFS never both apply to the same row (Actual is either blank or not).
  const actualElsePlanned = function (amountRange, actualRange, plannedRange, d, extraCriteria) {
    const extra = extraCriteria ? "," + extraCriteria : "";
    return `SUMIFS(${amountRange}${extra},${actualRange},${d})+SUMIFS(${amountRange}${extra},${actualRange},"",${plannedRange},${d})`;
  };

  const poSkuRange = poRange(poCol_("SKU"));
  const poTranches = PO_PAYMENT_SLOT_NUMBERS.map(function (i) {
    return {
      amount: poRange(poCol_(`Payment ${i} — Amount`)),
      planned: poRange(poCol_(`Payment ${i} — Planned Date`)),
      actual: poRange(poCol_(`Payment ${i} — Actual Date`)),
    };
  });
  const inspectionRange = poRange(poCol_("Inspection Amount"));
  const payment2 = poTranches[1];

  const shipTranches = [
    { amount: shRange(shipCol_("Payment 1 (Freight) — Amount")), planned: shRange(shipCol_("Payment 1 (Freight) — Planned Date")), actual: shRange(shipCol_("Payment 1 (Freight) — Actual Date")) },
    { amount: shRange(shipCol_("Payment 2 (Freight) — Amount")), planned: shRange(shipCol_("Payment 2 (Freight) — Planned Date")), actual: shRange(shipCol_("Payment 2 (Freight) — Actual Date")) },
    { amount: shRange(shipCol_("Customs — Amount")), planned: shRange(shipCol_("Customs — Planned Date")), actual: shRange(shipCol_("Customs — Actual Date")) },
  ];
  const customsAmtRange = shRange(shipCol_("Customs — Amount"));
  const customsEstRange = shRange(shipCol_("Customs — Estimated Amount"));
  const customsPlannedRange = shRange(shipCol_("Customs — Planned Date"));
  const departPlannedRange = shRange(shipCol_("Depart Date (planned)"));
  const totalQtyRange = shRange(shipCol_("Total Qty"));
  const projectedRange = shRange(shipCol_("Projected Shipping+Taxes Cost/unit"));

  const spLastRow = HORIZON_DAYS + 1;
  const spDateCol = colLetterIn_(SALES_PLAN_COLUMNS, "Date");
  const spRevPlanCol = colLetterIn_(SALES_PLAN_COLUMNS, "Plan Revenue €/day");
  const spRevActualCol = colLetterIn_(SALES_PLAN_COLUMNS, "Actual Revenue €/day");

  const fdDateRange = fdRange(fulfillmentCarrierCol_("Date"));
  const fdWhRange = fdRange(fulfillmentCarrierCol_("Warehouse"));
  const fdAmountRanges = {
    Fulfillment: fdRange(fulfillmentCarrierCol_("Fulfillment Actual Amount")),
    Carrier: fdRange(fulfillmentCarrierCol_("Carrier Actual Amount")),
  };

  const purchaseFormulas = SKUS.map(function () { return []; });
  const logisticsFormulas = [], totalFormulas = [], runningFormulas = [];
  const costKindFormulas = {};
  CASHFLOW_COST_KINDS.forEach(function (kind) { costKindFormulas[kind] = []; });

  for (let r = firstDataRow; r <= lastDataRow; r++) {
    const d = `${dateLetter}${r}`;

    SKUS.forEach(function (sku, skuIdx) {
      const skuCrit = `${poSkuRange},"${sku}"`;
      const terms = poTranches.map(function (t) { return actualElsePlanned(t.amount, t.actual, t.planned, d, skuCrit); });
      const inspectionTerm = actualElsePlanned(inspectionRange, payment2.actual, payment2.planned, d, skuCrit);
      purchaseFormulas[skuIdx].push([`=${terms.join("+")}+${inspectionTerm}`]);
    });

    const shipActualTerm = shipTranches.map(function (t) { return actualElsePlanned(t.amount, t.actual, t.planned, d); }).join("+");
    const customsEstimateTerm =
      `SUMPRODUCT((N(${customsAmtRange})=0)*(${customsPlannedRange}=${d})*N(${customsEstRange}))`;
    const nothingKnownYet = shipTranches.map(function (t) { return `(N(${t.amount})=0)`; }).join("*") + `*(N(${customsEstRange})=0)`;
    const shipProjectedTerm =
      `SUMPRODUCT(${nothingKnownYet}*(${departPlannedRange}=${d})*N(${totalQtyRange})*N(${projectedRange}))`;
    logisticsFormulas.push([`=${shipActualTerm}+${customsEstimateTerm}+${shipProjectedTerm}`]);

    // Evaluated per warehouse, then summed: a real invoice entered for one
    // warehouse supersedes only that warehouse's estimate for that date.
    CASHFLOW_COST_KINDS.forEach(function (kind) {
      const perWh = WAREHOUSES.map(function (wh) {
        const spName = salesPlanSheetName_(wh);
        const lookup = function (col) {
          return `IFERROR(INDEX('${spName}'!$${col}$2:$${col}$${spLastRow},MATCH(${d},'${spName}'!$${spDateCol}$2:$${spDateCol}$${spLastRow},0)),0)`;
        };
        const rate = rateCells[kind + "|" + wh];
        const invoice = `SUMIFS(${fdAmountRanges[kind]},${fdDateRange},${d},${fdWhRange},"${wh}")`;
        const estimateActual = `(${lookup(spRevActualCol)}/1000*${rate})`;
        const estimatePlan = `(${lookup(spRevPlanCol)}/1000*${rate})`;
        return `IF(${d}<TODAY(),IF(${invoice}>0,${invoice},${estimateActual}),${estimatePlan})`;
      });
      costKindFormulas[kind].push([`=${perWh.join("+")}`]);
    });

    const parts = SKUS.map(cashflowSkuPurchaseCol_).concat(["Logistics"]).concat(CASHFLOW_COST_KINDS)
      .map(function (name) { return colLetterIn_(CASHFLOW_COLUMNS, name) + r; });
    totalFormulas.push([`=${parts.join("+")}`]);

    const totalLetter = colLetterIn_(CASHFLOW_COLUMNS, "Total");
    const runningLetter = colLetterIn_(CASHFLOW_COLUMNS, "Running Total");
    runningFormulas.push([r === firstDataRow ? `=${totalLetter}${r}` : `=${runningLetter}${r - 1}+${totalLetter}${r}`]);
  }

  SKUS.forEach(function (sku, skuIdx) {
    sh.getRange(firstDataRow, colIndexIn_(CASHFLOW_COLUMNS, cashflowSkuPurchaseCol_(sku)), n, 1)
      .setFormulas(purchaseFormulas[skuIdx]).setNumberFormat("#,##0.00");
  });
  sh.getRange(firstDataRow, colIndexIn_(CASHFLOW_COLUMNS, "Logistics"), n, 1).setFormulas(logisticsFormulas).setNumberFormat("#,##0.00");
  CASHFLOW_COST_KINDS.forEach(function (kind) {
    sh.getRange(firstDataRow, colIndexIn_(CASHFLOW_COLUMNS, kind), n, 1).setFormulas(costKindFormulas[kind]).setNumberFormat("#,##0.00");
  });
  sh.getRange(firstDataRow, colIndexIn_(CASHFLOW_COLUMNS, "Total"), n, 1).setFormulas(totalFormulas).setNumberFormat("#,##0.00");
  sh.getRange(firstDataRow, colIndexIn_(CASHFLOW_COLUMNS, "Running Total"), n, 1).setFormulas(runningFormulas).setNumberFormat("#,##0.00");

  sh.autoResizeColumns(1, CASHFLOW_COLUMNS.length);
  SpreadsheetApp.flush();
  Logger.log("Cashflow built: %s day rows.", n);
}

// ---------------------------------------------------------------------------
// Transactions: payment ledger that drives Paid?/Actual Date elsewhere
// ---------------------------------------------------------------------------
//
// Mostly manual. "PO#/Shipment Ref" is free text, comma-separated when one real
// payment covers several rows (e.g. one customs invoice for a multi-SKU
// container). Purchase Orders and Shipments read it via txnRefMatchDateFormula_.

const TRANSACTIONS_SHEET_NAME = "Transactions";
const TRANSACTIONS_TEMPLATE_ROWS = 500;

const TRANSACTIONS_COLUMNS = [
  "Transaction ID", "Due/Scheduled date", "Date paid", "Type", "Amount", "Currency",
  "FX rate on transaction date", "Amount (EUR)", "Counterparty",
  "Expense category", "Channel/warehouse", "Status", "Reconciled",
  "Our bank/account", "Vendor bank/account", "Supporting document", "Notes",
  "PO#/Shipment Ref",
];

// The keys of each map ARE the dropdown list, so list and colors never drift.
const TXN_TYPE_COLORS = {
  "vendor payment": "#cfe2ff", "freight": "#ffe5b4", "customs/VAT": "#f5d0fe",
  "refund": "#d1f7c4", "other": "#e0e0e0",
};
const TXN_CURRENCY_COLORS = { "EUR": "#cde7d8", "USD": "#cfe2ff", "CNY": "#ffe0b3" };
const TXN_EXPENSE_COLORS = {
  "Goods": "#cfe2ff", "Inspection": "#e2d9f3", "Freight & Customs (incl. taxes)": "#ffe5b4",
  "Fulfillment": "#b2dfdb", "Last Mile": "#ffd6e8",
};
const TXN_CHANNEL_COLORS = { "WH-DE": "#cfe2ff", "WH-CN": "#ffe0cc" };
// Status is a live formula, not a dropdown; this map only colors it.
const TXN_STATUS_COLORS = { "paid": "#d4edda", "scheduled": "#fff3cd", "overdue": "#f8d7da" };

const TXN_MANUAL_COLUMNS = [
  "Due/Scheduled date", "Date paid", "Type", "Amount", "Currency", "FX rate on transaction date",
  "Counterparty", "Expense category", "Channel/warehouse",
  "Reconciled", "Our bank/account", "Vendor bank/account", "Supporting document", "Notes",
  "PO#/Shipment Ref",
];

function txnCol_(name) { return colLetterIn_(TRANSACTIONS_COLUMNS, name); }
function txnColIndex_(name) { return colIndexIn_(TRANSACTIONS_COLUMNS, name); }

function txnAddColorRules_(sh, colIndex, colorMap, numRows) {
  const range = sh.getRange(2, colIndex, numRows, 1);
  const rules = sh.getConditionalFormatRules();
  Object.keys(colorMap).forEach(function (value) {
    rules.push(
      SpreadsheetApp.newConditionalFormatRule()
        .whenTextEqualTo(value)
        .setBackground(colorMap[value])
        .setRanges([range])
        .build()
    );
  });
  sh.setConditionalFormatRules(rules);
}

function buildTransactions() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = rebuildTab_(ss, TRANSACTIONS_SHEET_NAME);
  writeTable_(sh, TRANSACTIONS_COLUMNS, []);

  const n = TRANSACTIONS_TEMPLATE_ROWS;
  const dueLetter = txnCol_("Due/Scheduled date"), paidDateLetter = txnCol_("Date paid");
  const amtLetter = txnCol_("Amount"), curLetter = txnCol_("Currency"), fxLetter = txnCol_("FX rate on transaction date");

  const idFormulas = [], eurFormulas = [], statusFormulas = [];
  for (let r = 2; r <= n + 1; r++) {
    idFormulas.push([`=IF($${dueLetter}${r}="","","TXN-" & TEXT(ROW()-1,"0000"))`]);
    eurFormulas.push([`=IF($${amtLetter}${r}="","",IF($${curLetter}${r}="EUR",$${amtLetter}${r},$${amtLetter}${r}*$${fxLetter}${r}))`]);
    // Status derived from dates, so it can never drift out of sync with Date paid.
    statusFormulas.push([`=IF($${dueLetter}${r}="","",IF($${paidDateLetter}${r}<>"","paid",IF(TODAY()>$${dueLetter}${r},"overdue","scheduled")))`]);
  }
  sh.getRange(2, txnColIndex_("Transaction ID"), n, 1).setFormulas(idFormulas);
  sh.getRange(2, txnColIndex_("Amount (EUR)"), n, 1).setFormulas(eurFormulas).setNumberFormat("#,##0.00");
  sh.getRange(2, txnColIndex_("Status"), n, 1).setFormulas(statusFormulas);

  sh.getRange(2, txnColIndex_("Due/Scheduled date"), n, 1).setNumberFormat("yyyy-mm-dd");
  sh.getRange(2, txnColIndex_("Date paid"), n, 1).setNumberFormat("yyyy-mm-dd");
  sh.getRange(2, txnColIndex_("Amount"), n, 1).setNumberFormat("#,##0.00");
  sh.getRange(2, txnColIndex_("FX rate on transaction date"), n, 1).setNumberFormat("0.0000");

  const dateRule = SpreadsheetApp.newDataValidation().requireDate().setAllowInvalid(true).build();
  [txnColIndex_("Due/Scheduled date"), txnColIndex_("Date paid")].forEach(function (c) {
    sh.getRange(2, c, n, 1).setDataValidation(dateRule);
  });

  const dropdowns = {
    "Type": TXN_TYPE_COLORS, "Currency": TXN_CURRENCY_COLORS,
    "Expense category": TXN_EXPENSE_COLORS, "Channel/warehouse": TXN_CHANNEL_COLORS,
  };
  Object.keys(dropdowns).forEach(function (name) {
    const rule = SpreadsheetApp.newDataValidation()
      .requireValueInList(Object.keys(dropdowns[name]), true)
      .setAllowInvalid(true) // a value not yet in the list can still be typed
      .build();
    sh.getRange(2, txnColIndex_(name), n, 1).setDataValidation(rule);
  });

  sh.getRange(2, txnColIndex_("Reconciled"), n, 1).insertCheckboxes();

  txnAddColorRules_(sh, txnColIndex_("Type"), TXN_TYPE_COLORS, n);
  txnAddColorRules_(sh, txnColIndex_("Currency"), TXN_CURRENCY_COLORS, n);
  txnAddColorRules_(sh, txnColIndex_("Expense category"), TXN_EXPENSE_COLORS, n);
  txnAddColorRules_(sh, txnColIndex_("Channel/warehouse"), TXN_CHANNEL_COLORS, n);
  txnAddColorRules_(sh, txnColIndex_("Status"), TXN_STATUS_COLORS, n);

  TXN_MANUAL_COLUMNS.forEach(function (name) {
    sh.getRange(2, txnColIndex_(name), n, 1).setBackground(SCT_YELLOW);
  });

  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, TRANSACTIONS_COLUMNS.length);
  SpreadsheetApp.flush();
  Logger.log("Transactions built: %s template rows.", n);
}

// ---------------------------------------------------------------------------
// Optional: lock computed tabs
// ---------------------------------------------------------------------------
//
// Google Sheets cannot grant edit access to a single tab, so manual-input tabs
// stay editable for anyone with Editor access and every other tab is protected so
// only the file owner can edit it. Safe to re-run (protections are recreated, not
// stacked). Not called by setupAll(); run it after setup if wanted.

function protectAllTabsExceptManualInput() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const manualInputTabs = [
    PO_SHEET_NAME, SHIPMENTS_SHEET_NAME,
    FULFILLMENT_CARRIER_SHEET_NAME, INPUTS_SHEET_NAME, TRANSACTIONS_SHEET_NAME,
    CASHFLOW_SHEET_NAME, // rate cells at the top are manual input
  ].concat(WAREHOUSES.map(salesPlanSheetName_));
  const owner = Session.getEffectiveUser();

  ss.getSheets().forEach(function (sh) {
    if (manualInputTabs.indexOf(sh.getName()) !== -1) return;

    sh.getProtections(SpreadsheetApp.ProtectionType.SHEET).forEach(function (p) { p.remove(); });

    const protection = sh.protect().setDescription("Locked: edit access restricted to the file owner");
    protection.removeEditors(protection.getEditors());
    if (protection.canDomainEdit()) protection.setDomainEdit(false);
    protection.addEditor(owner);
  });

  SpreadsheetApp.flush();
  Logger.log("Protected every tab except %s.", manualInputTabs.join(", "));
}

// ---------------------------------------------------------------------------
// Demo data
// ---------------------------------------------------------------------------
//
// SYNTHETIC DEMO DATA. Every name, quantity, price, rate and date below is
// invented to make the formulas visibly do something. It does not describe any
// real company, product, supplier or shipment.
//
// All dates are relative to START_DATE (day 0 = first row of the stock grid;
// "today" is day LOOKBACK_DAYS), so the demo stays meaningful whenever it runs.
// Prices are in EUR: the model sums factory cost and logistics cost per unit, so
// it assumes one reporting currency (non-EUR transactions convert via FX rate).
//
// What the demo shows:
//   - 2 POs: PO-001 (3 SKUs, fully shipped) and PO-002 (SKU-1, in production)
//   - 8 shipment rows: Sea, Rail, Air, Local; both warehouses; one SKU per row;
//     one Air shipment still in transit (drives Trans d and planned IN)
//   - 60 days of noisy actual sales per SKU per warehouse + a flat forward plan
//   - Transactions that mark most PO/freight/customs tranches as paid (one late)
//   - Synthetic fulfillment/carrier rates + one invoice override per warehouse

const DEMO_UNIT_PRICE = { "SKU-1": 2.00, "SKU-2": 1.50, "SKU-3": 0.80 };
const DEMO_SELLING_PRICE = { "SKU-1": 9.00, "SKU-2": 19.00, "SKU-3": 6.00 };
const DEMO_DAILY_PLAN = {
  "WH-DE": { "SKU-1": 380, "SKU-2": 18, "SKU-3": 55 },
  "WH-CN": { "SKU-1": 45,  "SKU-2": 3,  "SKU-3": 8 },
};

// Day offset relative to START_DATE -> "yyyy-mm-dd". Written as text so Sheets
// parses it as a plain calendar date regardless of the script's timezone.
function demoDate_(offsetDays) {
  const d = addDays_(START_DATE, offsetDays);
  const pad = function (x) { return (x < 10 ? "0" : "") + x; };
  return `${d.getFullYear()}-${pad(d.getMonth() + 1)}-${pad(d.getDate())}`;
}

// Deterministic pseudo-random noise (same demo on every run).
function demoNoise_(seed) {
  let state = seed;
  return function () {
    state = (state * 1103515245 + 12345) % 2147483648;
    return state / 2147483648;
  };
}

function writeRowFields_(sh, row, colIndexFn, fields) {
  Object.keys(fields).forEach(function (name) {
    const value = fields[name];
    if (value !== null && value !== undefined && value !== "") sh.getRange(row, colIndexFn(name)).setValue(value);
  });
}

function seedDemoData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  [PO_SHEET_NAME, SHIPMENTS_SHEET_NAME, TRANSACTIONS_SHEET_NAME, FULFILLMENT_CARRIER_SHEET_NAME, CASHFLOW_SHEET_NAME]
    .concat(WAREHOUSES.map(salesPlanSheetName_))
    .forEach(function (name) {
      if (!ss.getSheetByName(name)) throw new Error("Tab " + name + " not found: run setupStructureOnly() first.");
    });

  seedDemoPurchaseOrders_(ss.getSheetByName(PO_SHEET_NAME));
  seedDemoShipments_(ss.getSheetByName(SHIPMENTS_SHEET_NAME));
  seedDemoTransactions_(ss.getSheetByName(TRANSACTIONS_SHEET_NAME));
  WAREHOUSES.forEach(function (wh, i) { seedDemoSales_(ss.getSheetByName(salesPlanSheetName_(wh)), wh, 1000 + i); });
  seedDemoFulfillmentCarrier_(ss.getSheetByName(FULFILLMENT_CARRIER_SHEET_NAME), ss.getSheetByName(CASHFLOW_SHEET_NAME));

  SpreadsheetApp.flush();
  Logger.log("Synthetic demo data loaded.");
}

function seedDemoPurchaseOrders_(sh) {
  const po001 = function (sku, qtyDE, qtyCN, p1, p2, inspection, stdDE, stdCN) {
    return {
      "PO#": "PO-001", "Vendor": "Factory A", "SKU": sku,
      "Planned Qty — WH-DE": qtyDE, "Planned Qty — WH-CN": qtyCN,
      "Planned Ready Date": demoDate_(-45), "Actual Ready Date": demoDate_(-44),
      "Qty Produced": qtyDE + qtyCN,
      "EXW/unit": DEMO_UNIT_PRICE[sku], "Inspection Amount": inspection,
      "Currency": "EUR", "Order date": demoDate_(-75), "Status": "Closed",
      "Payment 1 — Amount": p1, "Payment 1 — Planned Date": demoDate_(-75),
      "Payment 2 — Amount": p2, "Payment 2 — Planned Date": demoDate_(-40),
      "Standard Landed Cost/unit — WH-DE": stdDE, "Standard Landed Cost/unit — WH-CN": stdCN,
    };
  };
  const rows = [
    // 45,000 x 2.00 + 500 inspection = 90,500 = 27,000 + 63,500
    po001("SKU-1", 40000, 5000, 27000, 63500, 500, 2.60, 2.10),
    // 2,000 x 1.50 = 3,000 = 900 + 2,100
    po001("SKU-2", 1500, 500, 900, 2100, null, 2.40, 1.70),
    // 6,800 x 0.80 = 5,440 = 1,632 + 3,808
    po001("SKU-3", 6000, 800, 1632, 3808, null, 1.10, 0.90),
    {
      // 33,000 x 2.00 = 66,000 = 19,800 + 46,200; still in production (drives Prod d)
      "PO#": "PO-002", "Vendor": "Factory A", "SKU": "SKU-1",
      "Planned Qty — WH-DE": 30000, "Planned Qty — WH-CN": 3000,
      "Planned Ready Date": demoDate_(80),
      "EXW/unit": DEMO_UNIT_PRICE["SKU-1"],
      "Currency": "EUR", "Order date": demoDate_(40), "Status": "In Production",
      "Payment 1 — Amount": 19800, "Payment 1 — Planned Date": demoDate_(40),
      "Payment 2 — Amount": 46200, "Payment 2 — Planned Date": demoDate_(78),
      "Standard Landed Cost/unit — WH-DE": 2.60, "Standard Landed Cost/unit — WH-CN": 2.10,
    },
  ];
  rows.forEach(function (fields, i) { writeRowFields_(sh, i + 2, colIndex_, fields); });
}

function seedDemoShipments_(sh) {
  // One SKU per row. Offsets are days from START_DATE; arrival null = in transit.
  const ship = function (id, method, wh, dep, eta, arr, sku, qty, cost) {
    const landed = arr !== null;
    const fields = {
      "Shipment ID": id, "PO#": "PO-001", "Ship method": method,
      "Planned Warehouse": wh, "Actual Warehouse": landed ? wh : null,
      "Depart Date (planned)": demoDate_(dep), "Actual Depart Date": demoDate_(dep),
      "ETA (planned)": demoDate_(eta), "Actual Arrival Date": landed ? demoDate_(arr) : null,
      "Payment 1 (Freight) — Amount": cost.freight1, "Payment 1 (Freight) — Planned Date": demoDate_(dep),
      "Payment 2 (Freight) — Amount": cost.freight2 || null,
      "Payment 2 (Freight) — Planned Date": cost.freight2 ? demoDate_(eta) : null,
      "Customs — Estimated Amount": cost.customsEstimate || null,
      "Customs — Planned Date": cost.customs || cost.customsEstimate ? demoDate_(eta) : null,
      "Customs Status": cost.customs ? "Cleared" : (method === "Local" ? null : "Not started"),
      "Customs Declared Value/unit": method === "Local" ? null : DEMO_UNIT_PRICE[sku],
      "Actual — Delivery": landed ? (cost.freight1 + (cost.freight2 || 0)) : null,
      "Actual — Admin Fees": cost.customs ? cost.customs.admin : null,
      "Actual — Duty (non-refundable)": cost.customs ? cost.customs.duty : null,
      "Actual — Import VAT (refundable)": cost.customs ? cost.customs.importVat : null,
      "Actual — Other VAT (refundable)": cost.customs ? cost.customs.otherVat : null,
    };
    fields[shipSkuQtyCol_(sku)] = qty;
    return fields;
  };
  const customs = function (admin, duty, importVat, otherVat) {
    return { admin: admin, duty: duty, importVat: importVat, otherVat: otherVat };
  };
  const rows = [
    ship("SHP-001", "Sea",   "WH-DE", -35, 0,  0,    "SKU-1", 16000, { freight1: 3000, customs: customs(150, 1200, 5750, 30) }),
    ship("SHP-002", "Rail",  "WH-DE", 10,  35, 36,   "SKU-1", 14000, { freight1: 1500, freight2: 2700, customs: customs(150, 1200, 5850, 30) }),
    ship("SHP-003", "Air",   "WH-DE", 52,  66, null, "SKU-1", 10000, { freight1: 6300, customsEstimate: 6000 }),
    ship("SHP-004", "Air",   "WH-DE", -8,  0,  0,    "SKU-2", 1500,  { freight1: 1100, customs: customs(100, 60, 600, 20) }),
    ship("SHP-005", "Sea",   "WH-DE", -35, 0,  0,    "SKU-3", 6000,  { freight1: 900,  customs: customs(100, 150, 1000, 20) }),
    ship("SHP-006", "Local", "WH-CN", -2,  0,  0,    "SKU-1", 5000,  { freight1: 250 }),
    ship("SHP-007", "Local", "WH-CN", -2,  0,  0,    "SKU-2", 500,   { freight1: 60 }),
    ship("SHP-008", "Local", "WH-CN", -2,  0,  0,    "SKU-3", 800,   { freight1: 80 }),
  ];
  rows.forEach(function (fields, i) { writeRowFields_(sh, i + 2, shipColIndex_, fields); });
}

function seedDemoTransactions_(sh) {
  // Each payment's ref ties it to a PO# or Shipment ID; amount closeness picks the
  // tranche. PO-001 SKU-1 Payment 2 is paid 2 days late on purpose (Days Late).
  const txn = function (paid, type, amount, counterparty, expense, channel, ref) {
    return {
      "Due/Scheduled date": demoDate_(paid), "Date paid": demoDate_(paid), "Type": type,
      "Amount": amount, "Currency": "EUR", "Counterparty": counterparty,
      "Expense category": expense, "Channel/warehouse": channel, "PO#/Shipment Ref": ref,
    };
  };
  const rows = [
    txn(-75, "vendor payment", 27000, "Factory A", "Goods", null, "PO-001"),
    txn(-75, "vendor payment", 900,   "Factory A", "Goods", null, "PO-001"),
    txn(-74, "vendor payment", 1632,  "Factory A", "Goods", null, "PO-001"),
    txn(-38, "vendor payment", 63500, "Factory A", "Goods", null, "PO-001"),
    txn(-40, "vendor payment", 2100,  "Factory A", "Goods", null, "PO-001"),
    txn(-40, "vendor payment", 3808,  "Factory A", "Goods", null, "PO-001"),
    txn(-35, "freight", 3000, "Forwarder A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-001"),
    txn(-35, "freight", 900,  "Forwarder A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-005"),
    txn(-8,  "freight", 1100, "Forwarder A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-004"),
    txn(1, "customs/VAT", 7130, "Customs Broker A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-001"),
    txn(1, "customs/VAT", 780,  "Customs Broker A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-004"),
    txn(1, "customs/VAT", 1270, "Customs Broker A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-005"),
    txn(10, "freight", 1500, "Forwarder A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-002"),
    txn(37, "customs/VAT", 7230, "Customs Broker A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-002"),
    txn(41, "vendor payment", 19800, "Factory A", "Goods", null, "PO-002"),
    txn(52, "freight", 6300, "Forwarder A", "Freight & Customs (incl. taxes)", "WH-DE", "SHP-003"),
    txn(30, "vendor payment", 3000, "3PL DE", "Fulfillment", "WH-DE", null),
  ];
  rows.forEach(function (fields, i) { writeRowFields_(sh, i + 2, txnColIndex_, fields); });
}

function seedDemoSales_(sh, warehouse, seed) {
  const rand = demoNoise_(seed);
  const n = HORIZON_DAYS;
  const actualDays = LOOKBACK_DAYS; // days 0..LOOKBACK_DAYS-1 are in the past
  const planRevenue = [], actualRevenue = [];
  for (let i = 0; i < n; i++) { planRevenue.push([0]); actualRevenue.push([i < actualDays ? 0 : ""]); }

  SKUS.forEach(function (sku) {
    const plan = DEMO_DAILY_PLAN[warehouse][sku];
    const planValues = [], actualValues = [];
    for (let i = 0; i < n; i++) {
      planValues.push([plan]);
      planRevenue[i][0] += plan * DEMO_SELLING_PRICE[sku];
      if (i < actualDays) {
        const actual = Math.max(0, Math.round(plan * (0.85 + 0.3 * rand())));
        actualValues.push([actual]);
        actualRevenue[i][0] += actual * DEMO_SELLING_PRICE[sku];
      } else {
        actualValues.push([""]);
      }
    }
    sh.getRange(2, colIndexIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Plan/day")), n, 1).setValues(planValues);
    sh.getRange(2, colIndexIn_(SALES_PLAN_COLUMNS, salesPlanWideCol_(sku, "Actual/day")), n, 1).setValues(actualValues);
  });

  sh.getRange(2, colIndexIn_(SALES_PLAN_COLUMNS, "Plan Revenue €/day"), n, 1).setValues(planRevenue);
  sh.getRange(2, colIndexIn_(SALES_PLAN_COLUMNS, "Actual Revenue €/day"), n, 1).setValues(actualRevenue);
}

function seedDemoFulfillmentCarrier_(actualsSh, cashflowSh) {
  const rates = {
    "Fulfillment|WH-DE": 25, "Carrier|WH-DE": 40,
    "Fulfillment|WH-CN": 90, "Carrier|WH-CN": 0, // all-in 3PL rate, no separate carrier leg
  };
  let rateRow = 1;
  WAREHOUSES.forEach(function (wh) {
    CASHFLOW_COST_KINDS.forEach(function (kind) {
      cashflowSh.getRange(rateRow, 2).setValue(rates[kind + "|" + wh]);
      rateRow++;
    });
  });

  const rows = [
    { "Date": demoDate_(30), "Warehouse": "WH-DE", "Fulfillment Actual Amount": 3000, "Carrier Actual Amount": 5000 },
    { "Date": demoDate_(30), "Warehouse": "WH-CN", "Fulfillment Actual Amount": 1400 },
  ];
  rows.forEach(function (fields, i) { writeRowFields_(actualsSh, i + 2, fulfillmentCarrierColIndex_, fields); });
}
