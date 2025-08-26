/***** CONFIG *****/
// Your data tab:
const SHEET_NAME = 'Orderform';

/**
 * Header names exactly as they appear in row 1.
 * Adjust if your sheet uses different labels.
 */
const HEADERS = {
  OrderformCode: "OrderformCode", // A
  CustomerName:  "CustomerName",  // C
  Ref:           "Ref.",          // D (note the dot)
  StartDate:     "StartDate",     // E
  EndDate:       "EndDate",       // F
  SubProduct:    "SubProduct",    // I
  PaymentTerm:   "PaymentTerm",   // J (change to 'PaymentTerr' if that's your header)
  LicenseOrdered:"LicenseOrdered",// K
  Price:         "Price",         // N
  Amount:        "Amount",        // O
  Comments:      "Comments",      // S (optional)
  ClientID:      "ClientID"       // optional column for client id
};

// Ignore any row where Ref. contains this text (case-insensitive)
const REF_EXCLUDE_TEXT = "Unsigned Prospect Contract";

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Contracts")
    .addItem("Open (Sidebar)", "showSidebar")
    .addItem("Open (Large Popup)", "showPopup")
    .addToUi();
}

function showSidebar() {
  const t = HtmlService.createTemplateFromFile("Sidebar");
  t.initialCustomer = '';
  const html = t.evaluate().setTitle("Customer Contracts");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showPopup() {
  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.initialCustomer = '';
  const html = t.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Customer Contracts');
}

/***** DATA HELPERS *****/
function _getSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found`);
  return sh;
}

function _getAllRows() {
  const sh = _getSheet();
  const values = sh.getDataRange().getValues();
  if (values.length < 2) return { headers: [], rows: [], idx: {} };

  const headers = values[0];
  const idx = {};
  headers.forEach((h, i) => { if (h) idx[String(h).trim()] = i; });

  // Required headers
  const REQUIRED = [
    HEADERS.OrderformCode,
    HEADERS.CustomerName,
    HEADERS.Ref,
    HEADERS.StartDate,
    HEADERS.EndDate,
    HEADERS.SubProduct,
    HEADERS.LicenseOrdered,
    HEADERS.Price,
    HEADERS.Amount
  ];
  const missing = REQUIRED.filter(h => !(h in idx));
  if (missing.length) throw new Error("Missing required headers: " + missing.join(", "));

  // Optional headers: set to -1 if not present
  [HEADERS.PaymentTerm, HEADERS.Comments, HEADERS.ClientID].forEach(h => {
    if (!(h in idx)) idx[h] = -1;
  });

  const rows = values.slice(1).filter(r => String(r.join("")).trim() !== "");
  return { headers, rows, idx };
}

function _val(row, idx, key) {
  const i = idx[key];
  if (i === undefined || i === -1) return "";
  return row[i];
}

/***** API: customers (unique, excluding filtered-out Ref.) *****/
function getCustomers() {
  const { rows, idx } = _getAllRows();
  const set = new Set();
  rows.forEach(r => {
    const refVal = String(_val(r, idx, HEADERS.Ref) || "");
    if (refVal.toLowerCase().includes(REF_EXCLUDE_TEXT.toLowerCase())) return;
    const name = _val(r, idx, HEADERS.CustomerName);
    if (name && String(name).trim() !== "") set.add(String(name).trim());
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

/***** API: contracts for a customer *****/
function getCustomerContracts(customerName) {
  const { rows, idx } = _getAllRows();

  const filtered = rows.filter(r => {
    const name = String(_val(r, idx, HEADERS.CustomerName) || "").trim();
    if (name !== String(customerName).trim()) return false;
    const refVal = String(_val(r, idx, HEADERS.Ref) || "");
    if (refVal.toLowerCase().includes(REF_EXCLUDE_TEXT.toLowerCase())) return false;
    return true;
  });

  // Client ID (optional)
  const clientIds = new Set();
  filtered.forEach(r => {
    const cid = String(_val(r, idx, HEADERS.ClientID) || "").trim();
    if (cid) clientIds.add(cid);
  });

  // Group by OrderformCode
  const byContract = new Map();
  filtered.forEach(r => {
    const code   = String(_val(r, idx, HEADERS.OrderformCode) || "").trim();
    const refVal = _val(r, idx, HEADERS.Ref) || "";

    if (!byContract.has(code)) {
      byContract.set(code, {
        orderformCode: code,
        ref: refVal,
        startDate: _val(r, idx, HEADERS.StartDate) || "",
        endDate: _val(r, idx, HEADERS.EndDate) || "",
        paymentTerm: _val(r, idx, HEADERS.PaymentTerm) || "",
        lines: [],
        totalAmount: 0
      });
    }

    const line = {
      subProduct: _val(r, idx, HEADERS.SubProduct) || "",
      refValue: refVal || "", // NEW: D column value on each row
      licenseOrdered: Number(_val(r, idx, HEADERS.LicenseOrdered) || 0),
      price: Number(_val(r, idx, HEADERS.Price) || 0),
      amount: Number(_val(r, idx, HEADERS.Amount) || 0),
      comments: _val(r, idx, HEADERS.Comments) || "" // NEW: S column
    };

    const obj = byContract.get(code);
    obj.lines.push(line);
    obj.totalAmount += line.amount;
  });

  function asISO(d) {
    if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d)) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    return d ? String(d) : "";
  }

  const contracts = Array.from(byContract.values()).map(c => ({
    ...c,
    startDate: asISO(c.startDate),
    endDate: asISO(c.endDate)
  }));

  contracts.sort((a,b) =>
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

/***** Optional: open from ContractsStatus!B2 *****/
function openPopupFromStatus() {
  const sh = SpreadsheetApp.getActive().getSheetByName('ContractsStatus');
  if (!sh) throw new Error('Sheet "ContractsStatus" not found');
  const selected = (sh.getRange('B2').getDisplayValue() || '').trim();

  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.initialCustomer = selected;
  const html = t.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Customer Contracts');
}
