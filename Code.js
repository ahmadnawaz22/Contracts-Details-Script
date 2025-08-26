/***** CONFIG *****/
// If your data lives in another spreadsheet, put its ID here.
// Leave empty "" to read the active file hosting the code.
const DATA_SPREADSHEET_ID = "1KAj7Zu6IURxiwdjAmCYl4clIFzM2Ut5tW8EJA_9DIpg"; // e.g. "1AbCde..."; keep "" to use current file
const SHEET_NAME = "Orderform";


// Header names (row 1)
const HEADERS = {
  OrderformCode: "OrderformCode", // A
  CustomerName:  "CustomerName",  // C
  Ref:           "Ref.",          // D (note the dot)
  StartDate:     "StartDate",     // E
  EndDate:       "EndDate",       // F
  SubProduct:    "SubProduct",    // I
  PaymentTerm:   "PaymentTerm",   // J   (rename here if your header is PaymentTerr)
  LicenseOrdered:"LicenseOrdered",// K
  Price:         "Price",         // N
  Amount:        "Amount",        // O
  Comments:      "Comments",      // S (optional)
  ClientID:      "ClientID"       // optional
};

// Filter out any row whose Ref. contains this text (case-insensitive)
const REF_EXCLUDE_TEXT = "Unsigned Prospect Contract";

// Caching
const CACHE_TTL_SEC = 600;   // 10 minutes
const MAX_CACHE_BYTES = 95000;

/***** MENU / UI *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Contracts")
    .addItem("Open (Sidebar)", "showSidebar")
    .addItem("Open (Large Popup)", "showPopup")
    .addItem("Open from ContractsStatus!B2", "openPopupFromStatus")
    .addToUi();
}

function showSidebar() {
  const t = HtmlService.createTemplateFromFile("Sidebar");
  t.initialCustomer = "";
  SpreadsheetApp.getUi().showSidebar(t.evaluate().setTitle("Customer Contracts"));
}

function showPopup() {
  const t = HtmlService.createTemplateFromFile("Sidebar");
  t.initialCustomer = "";
  SpreadsheetApp.getUi().showModelessDialog(
    t.evaluate().setWidth(1000).setHeight(700),
    "Customer Contracts"
  );
}

function openPopupFromStatus() {
  const sh = SpreadsheetApp.getActive().getSheetByName("ContractsStatus");
  if (!sh) throw new Error('Sheet "ContractsStatus" not found');
  const selected = (sh.getRange("B2").getDisplayValue() || "").trim();
  const t = HtmlService.createTemplateFromFile("Sidebar");
  t.initialCustomer = selected;
  SpreadsheetApp.getUi().showModelessDialog(
    t.evaluate().setWidth(1000).setHeight(700),
    "Customer Contracts"
  );
}

/***** DATA I/O (strict cross-file) *****/
function _openDataSS_() {
  if (!DATA_SPREADSHEET_ID) {
    throw new Error('DATA_SPREADSHEET_ID is empty. Set the external source file ID.');
  }
  return SpreadsheetApp.openById(DATA_SPREADSHEET_ID);
}
function _getSheet_() {
  const sh = _openDataSS_().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Tab "${SHEET_NAME}" not found in source file.`);
  return sh;
}
function _headerIndex_(headers) {
  const idx = {};
  headers.forEach((h, i) => (idx[String(h).trim()] = i));
  return idx;
}
function _asISO_(d) {
  if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d)) {
    return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return d ? String(d) : "";
}
function _readHeaders_() {
  const sh = _getSheet_();
  return sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
}
function _readRowsBlock_(maxCol) {
  const sh = _getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];
  return sh.getRange(2, 1, lastRow - 1, maxCol).getValues();
}

/***** Fast customers (only necessary columns) *****/
function getCustomers() {
  const cache = CacheService.getUserCache();
  const cacheKey = `customers:${DATA_SPREADSHEET_ID}:${SHEET_NAME}`;
  const cached = cache.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const headers = _readHeaders_();
  const idx = _headerIndex_(headers);
  const cCol = (idx[HEADERS.CustomerName] ?? -1) + 1;
  const rCol = (idx[HEADERS.Ref] ?? -1) + 1;
  if (cCol <= 0 || rCol <= 0) throw new Error("Missing CustomerName or Ref. headers.");

  const sh = _getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return [];

  const names = sh.getRange(2, cCol, lastRow - 1, 1).getValues();
  const refs  = sh.getRange(2, rCol, lastRow - 1, 1).getValues();

  const set = new Set();
  const refEx = REF_EXCLUDE_TEXT.toLowerCase();
  for (let i = 0; i < names.length; i++) {
    const ref = String(refs[i][0] || "");
    if (ref.toLowerCase().includes(refEx)) continue;
    const n = (names[i][0] || "").toString().trim();
    if (n) set.add(n);
  }
  const arr = Array.from(set).sort((a,b)=>a.localeCompare(b));

  const json = JSON.stringify(arr);
  if (json.length <= MAX_CACHE_BYTES) cache.put(cacheKey, json, CACHE_TTL_SEC);
  return arr;
}

/***** Core builder (no per-line Ref., comments kept) *****/
function _buildContractsFromRows_(rows, idx, customerName) {
  const byContract = new Map();
  const clientIds = new Set();
  const refEx = REF_EXCLUDE_TEXT.toLowerCase();
  const custKey = HEADERS.CustomerName;

  for (let i = 0; i < rows.length; i++) {
    const r = rows[i];

    const cust = String(r[idx[custKey]] || "").trim();
    if (cust !== String(customerName).trim()) continue;

    const refVal = String(r[idx[HEADERS.Ref]] || "");
    if (refVal.toLowerCase().includes(refEx)) continue;

    const code = String(r[idx[HEADERS.OrderformCode]] || "").trim();
    if (!byContract.has(code)) {
      byContract.set(code, {
        orderformCode: code,
        ref: r[idx[HEADERS.Ref]] || "",
        startDate: r[idx[HEADERS.StartDate]] || "",
        endDate: r[idx[HEADERS.EndDate]] || "",
        paymentTerm: (idx[HEADERS.PaymentTerm] ?? -1) >= 0 ? (r[idx[HEADERS.PaymentTerm]] || "") : "",
        lines: [],
        totalAmount: 0
      });
    }

    const line = {
      subProduct: r[idx[HEADERS.SubProduct]] || "",
      // refValue REMOVED from line (requested)
      comments: (idx[HEADERS.Comments] ?? -1) >= 0 ? (r[idx[HEADERS.Comments]] || "") : "",
      licenseOrdered: Number(r[idx[HEADERS.LicenseOrdered]] || 0),
      price: Number(r[idx[HEADERS.Price]] || 0),
      amount: Number(r[idx[HEADERS.Amount]] || 0)
    };

    const obj = byContract.get(code);
    obj.lines.push(line);
    obj.totalAmount += line.amount;

    if ((idx[HEADERS.ClientID] ?? -1) >= 0) {
      const cid = String(r[idx[HEADERS.ClientID]] || "").trim();
      if (cid) clientIds.add(cid);
    }
  }

  const contracts = Array.from(byContract.values()).map(c => ({
    ...c,
    startDate: _asISO_(c.startDate),
    endDate: _asISO_(c.endDate)
  })).sort((a,b) =>
    (a.startDate > b.startDate ? 1 : a.startDate < b.startDate ? -1 : a.orderformCode.localeCompare(b.orderformCode))
  );

  return {
    customer: customerName,
    clientId: Array.from(clientIds).join(", "),
    contractCount: contracts.length,
    currencyHint: "AED",
    contracts
  };
}

/***** One-shot bootstrap *****/
function bootstrap(initialCustomer) {
  const headers = _readHeaders_();
  const idx = _headerIndex_(headers);

  const needKeys = [
    HEADERS.OrderformCode, HEADERS.CustomerName, HEADERS.Ref, HEADERS.StartDate,
    HEADERS.EndDate, HEADERS.SubProduct, HEADERS.LicenseOrdered, HEADERS.Price,
    HEADERS.Amount, HEADERS.PaymentTerm, HEADERS.Comments, HEADERS.ClientID
  ].filter(k => k in idx);
  const maxCol = Math.max(...needKeys.map(k => idx[k] + 1));
  const rows = _readRowsBlock_(maxCol);

  // unique customers (after filter)
  const refEx = REF_EXCLUDE_TEXT.toLowerCase();
  const custSet = new Set();
  for (const r of rows) {
    const refVal = String(r[idx[HEADERS.Ref]] || "");
    if (refVal.toLowerCase().includes(refEx)) continue;
    const n = String(r[idx[HEADERS.CustomerName]] || "").trim();
    if (n) custSet.add(n);
  }
  const customers = Array.from(custSet).sort((a,b)=>a.localeCompare(b));

  const selected =
    initialCustomer && customers.includes(initialCustomer)
      ? initialCustomer
      : (customers[0] || "");

  const firstPayload = selected
    ? _buildContractsFromRows_(rows, idx, selected)
    : { customer: "", clientId: "", contractCount: 0, contracts: [] };

  return { customers, selected, firstPayload };
}

/***** Subsequent selections (with cache) *****/
function getCustomerContracts(customerName) {
  const cache = CacheService.getUserCache();
  const ck = `contracts:${DATA_SPREADSHEET_ID}:${SHEET_NAME}:${customerName}`;
  const hit = cache.get(ck);
  if (hit) return JSON.parse(hit);

  const headers = _readHeaders_();
  const idx = _headerIndex_(headers);

  const needKeys = [
    HEADERS.OrderformCode, HEADERS.CustomerName, HEADERS.Ref, HEADERS.StartDate,
    HEADERS.EndDate, HEADERS.SubProduct, HEADERS.LicenseOrdered, HEADERS.Price,
    HEADERS.Amount, HEADERS.PaymentTerm, HEADERS.Comments, HEADERS.ClientID
  ].filter(k => k in idx);
  const maxCol = Math.max(...needKeys.map(k => idx[k] + 1));
  const rows = _readRowsBlock_(maxCol);

  const payload = _buildContractsFromRows_(rows, idx, customerName);

  const json = JSON.stringify(payload);
  if (json.length <= MAX_CACHE_BYTES) cache.put(ck, json, CACHE_TTL_SEC);
  return payload;
}
