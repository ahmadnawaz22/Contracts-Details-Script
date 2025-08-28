/***** CONFIG *****/
// If your data is in another spreadsheet, put its ID here; leave "" to use this file.
const DATA_SPREADSHEET_ID = "1KAj7Zu6IURxiwdjAmCYl4clIFzM2Ut5tW8EJA_9DIpg"; // e.g. "1AbCde..."; "" = current spreadsheet
const SHEET_NAME = "Orderform";

// Phrase used in Ref. column
const REF_EXCLUDE_TEXT = "Unsigned Prospect Contract"; // Renewals view will INCLUDE rows that contain this text

/***** FLEXIBLE HEADER NAMES *****/
const HEADER_ALIASES = {
  OrderformCode: ["OrderformCode", "OrderFormCode", "Order Form Code"],
  CustomerName:  ["CustomerName", "Customer Name"],
  Ref:           ["Ref.", "Ref"],
  StartDate:     ["StartDate", "Start Date", "Renewal Date"],
  EndDate:       ["EndDate", "End Date"],
  SubProduct:    ["SubProduct", "Sub Product", "Product", "Sub-Product"],
  PaymentTerm:   ["PaymentTerm", "PaymentTerr", "Payment Term"],
  LicenseOrdered:["LicenseOrdered", "Qty", "Quantity", "License Ordered"],
  Price:         ["Price", "Unit Price"],
  Amount:        ["Amount", "Total", "Line Total"],
  Comments:      ["Comments", "Comment", "Notes"],
  ClientID:      ["ClientID", "Client ID", "Customer ID"],
  AccountManager:["AccountManager", "Account Manager", "AM"]
};

/***** SHEET UI HELPERS (optional but handy) *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu("Contracts")
    .addItem("Open Sidebar", "showSidebar")
    .addItem("Renewals (Popup)", "showUnsignedProspects")
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createTemplateFromFile("Sidebar")
    .evaluate()
    .setTitle("Customer Contracts");
  SpreadsheetApp.getUi().showSidebar(html);
}

// Button-friendly: open Prospect.html (singular) as Renewals
function showUnsignedProspects() {
  const html = HtmlService.createHtmlOutputFromFile('Prospects') // <-- matches your file name
    .setWidth(1100)
    .setHeight(720);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Renewals');
}

// Optional alias if you ever wire a button to `showRenewals`
function showRenewals() { showUnsignedProspects(); }

/***** CORE HELPERS *****/
function _openDataSS_() {
  return DATA_SPREADSHEET_ID
    ? SpreadsheetApp.openById(DATA_SPREADSHEET_ID)
    : SpreadsheetApp.getActive();
}
function _getSheet_() {
  const sh = _openDataSS_().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Tab "${SHEET_NAME}" not found.`);
  return sh;
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
function _indexFromAliases_(headers) {
  const map = {};
  headers.forEach((h,i)=> map[String(h).trim()] = i);
  const idx = {};
  for (const [key, aliases] of Object.entries(HEADER_ALIASES)) {
    idx[key] = -1;
    for (const a of aliases) if (a in map) { idx[key] = map[a]; break; }
  }
  return idx;
}
function _require_(idx, keys) {
  const missing = keys.filter(k => idx[k] < 0);
  if (missing.length) throw new Error("Missing required headers: " + missing.join(", "));
}
function _asISO_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v ? String(v) : "";
}
function _parseDateKey_(v) {
  // normalize to numeric for reliable sorting
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) return v.getTime();
  if (typeof v === "string") {
    let d = new Date(v);
    if (!isNaN(d)) return d.getTime();
    const m = v.trim().match(/^(\d{1,2})\s+([A-Za-z]{3,})[a-z]*\s+(\d{2,4})$/);
    if (m) {
      const day = +m[1];
      const mon = ["jan","feb","mar","apr","may","jun","jul","aug","sep","oct","nov","dec"].indexOf(m[2].slice(0,3).toLowerCase());
      let year = +m[3]; if (year < 100) year += 2000;
      if (mon >= 0) { d = new Date(year, mon, day); if (!isNaN(d)) return d.getTime(); }
    }
  }
  return Number.POSITIVE_INFINITY; // blanks last
}

/***** RENEWALS (Prospect.html) *****/
// Returns contracts where Ref. CONTAINS "Unsigned Prospect Contract"
function getUnsignedProspectContracts(force) {
  const headers = _readHeaders_();
  const idx = _indexFromAliases_(headers);
  // Required for this payload:
  _require_(idx, ["OrderformCode","CustomerName","Ref","StartDate","EndDate","SubProduct","LicenseOrdered","Price","Amount"]);

  const needCols = [
    "OrderformCode","CustomerName","Ref","StartDate","EndDate",
    "SubProduct","LicenseOrdered","Price","Amount",
    "PaymentTerm","Comments","ClientID","AccountManager"
  ].filter(k => idx[k] >= 0).map(k => idx[k] + 1);
  const maxCol = Math.max(...needCols);
  const rows = _readRowsBlock_(maxCol);

  const phrase = (REF_EXCLUDE_TEXT || "").toLowerCase();
  const byContract = new Map();

  for (const r of rows) {
    const refVal = String(r[idx.Ref] || "");
    if (!refVal.toLowerCase().includes(phrase)) continue; // include ONLY renewals (unsigned prospect)

    const code = String(r[idx.OrderformCode] || "").trim();
    const customer = String(r[idx.CustomerName] || "").trim();

    if (!byContract.has(code)) {
      byContract.set(code, {
        orderformCode: code,
        customer,
        clientId: idx.ClientID >= 0 ? String(r[idx.ClientID] || "").trim() : "",
        startDate: r[idx.StartDate] || "",
        endDate: r[idx.EndDate] || "",
        paymentTerm: idx.PaymentTerm >= 0 ? (r[idx.PaymentTerm] || "") : "",
        lines: [],
        totalAmount: 0
      });
    }

    const line = {
      subProduct: r[idx.SubProduct] || "",
      comments:   idx.Comments >= 0 ? (r[idx.Comments] || "") : "",
      licenseOrdered: Number(r[idx.LicenseOrdered] || 0),
      price: Number(r[idx.Price] || 0),
      amount: Number(r[idx.Amount] || 0),
      accountManager: idx.AccountManager >= 0 ? (r[idx.AccountManager] || "") : ""
    };

    const obj = byContract.get(code);
    obj.lines.push(line);
    obj.totalAmount += line.amount;
  }

  const contracts = Array.from(byContract.values()).map(c => ({
    ...c,
    startDate: _asISO_(c.startDate),
    endDate:   _asISO_(c.endDate),
    _startKey: _parseDateKey_(c.startDate)
  })).sort((a,b) => {
    const d = a._startKey - b._startKey;
    if (d !== 0) return d;
    return a.orderformCode.localeCompare(b.orderformCode);
  });

  return { total: contracts.length, contracts };
}

/***** SIDEBAR SUPPORT (optional; safe to keep even if you only use Prospect.html) *****/
// Just the customer list (excluding renewals)
function bootstrapCustomers() {
  const headers = _readHeaders_();
  const idx = _indexFromAliases_(headers);
  _require_(idx, ["CustomerName","Ref"]);

  const sh = _getSheet_();
  const lastRow = sh.getLastRow();
  if (lastRow < 2) return { customers: [] };

  const names = sh.getRange(2, idx.CustomerName + 1, lastRow - 1, 1).getValues();
  const refs  = sh.getRange(2, idx.Ref + 1,           lastRow - 1, 1).getValues();

  const ex = (REF_EXCLUDE_TEXT || "").toLowerCase();
  const set = new Set();
  for (let i = 0; i < names.length; i++) {
    const ref = String(refs[i][0] || "");
    if (ref.toLowerCase().includes(ex)) continue;     // exclude renewals here
    const n = String(names[i][0] || "").trim();
    if (n) set.add(n);
  }
  return { customers: Array.from(set).sort((a,b)=>a.localeCompare(b)) };
}

// All customers (grouped) — excludes renewals
function getAllCustomerContracts() {
  const headers = _readHeaders_();
  const idx = _indexFromAliases_(headers);
  _require_(idx, ["OrderformCode","CustomerName","Ref","StartDate","EndDate","SubProduct","LicenseOrdered","Price","Amount"]);

  const needCols = [
    "OrderformCode","CustomerName","Ref","StartDate","EndDate",
    "SubProduct","LicenseOrdered","Price","Amount",
    "PaymentTerm","Comments","ClientID","AccountManager"
  ].filter(k => idx[k] >= 0).map(k => idx[k] + 1);
  const maxCol = Math.max(...needCols);
  const rows = _readRowsBlock_(maxCol);

  const ex = (REF_EXCLUDE_TEXT || "").toLowerCase();
  const byCustomer = new Map();

  for (const r of rows) {
    const refVal = String(r[idx.Ref] || "");
    if (refVal.toLowerCase().includes(ex)) continue; // exclude renewals here

    const cust = String(r[idx.CustomerName] || "").trim();
    if (!cust) continue;
    if (!byCustomer.has(cust)) byCustomer.set(cust, []);
    byCustomer.get(cust).push(r);
  }

  const groups = [];
  for (const [customerName, rs] of byCustomer) {
    const byContract = new Map();
    const clientIds = new Set();

    for (const r of rs) {
      const code = String(r[idx.OrderformCode] || "").trim();
      if (!byContract.has(code)) {
        byContract.set(code, {
          orderformCode: code,
          ref: r[idx.Ref] || "",
          startDate: r[idx.StartDate] || "",
          endDate: r[idx.EndDate] || "",
          paymentTerm: idx.PaymentTerm >= 0 ? (r[idx.PaymentTerm] || "") : "",
          lines: [],
          totalAmount: 0
        });
      }
      const line = {
        subProduct: r[idx.SubProduct] || "",
        comments:   idx.Comments >= 0 ? (r[idx.Comments] || "") : "",
        licenseOrdered: Number(r[idx.LicenseOrdered] || 0),
        price: Number(r[idx.Price] || 0),
        amount: Number(r[idx.Amount] || 0),
        accountManager: idx.AccountManager >= 0 ? (r[idx.AccountManager] || "") : ""
      };
      const obj = byContract.get(code);
      obj.lines.push(line);
      obj.totalAmount += line.amount;

      if (idx.ClientID >= 0) {
        const cid = String(r[idx.ClientID] || "").trim();
        if (cid) clientIds.add(cid);
      }
    }

    const contracts = Array.from(byContract.values())
      .map(c => ({ ...c, startDate: _asISO_(c.startDate), endDate: _asISO_(c.endDate) }))
      .sort((a,b) =>
        (a.startDate > b.startDate ? 1 : a.startDate < b.startDate ? -1 : a.orderformCode.localeCompare(b.orderformCode))
      );

    groups.push({
      customer: customerName,
      clientId: Array.from(clientIds).join(", "),
      contractCount: contracts.length,
      contracts
    });
  }

  groups.sort((a,b) => a.customer.localeCompare(b.customer));
  return { groups, totalGroups: groups.length };
}

// One customer (excludes renewals)
function getCustomerContracts(customerName) {
  const headers = _readHeaders_();
  const idx = _indexFromAliases_(headers);
  _require_(idx, ["OrderformCode","CustomerName","Ref","StartDate","EndDate","SubProduct","LicenseOrdered","Price","Amount"]);

  const needCols = [
    "OrderformCode","CustomerName","Ref","StartDate","EndDate",
    "SubProduct","LicenseOrdered","Price","Amount",
    "PaymentTerm","Comments","ClientID","AccountManager"
  ].filter(k => idx[k] >= 0).map(k => idx[k] + 1);
  const maxCol = Math.max(...needCols);
  const rows = _readRowsBlock_(maxCol);

  const ex = (REF_EXCLUDE_TEXT || "").toLowerCase();

  const filtered = rows.filter(r => {
    const name = String(r[idx.CustomerName] || "").trim();
    if (name !== String(customerName).trim()) return false;
    const refVal = String(r[idx.Ref] || "");
    if (refVal.toLowerCase().includes(ex)) return false; // exclude renewals here
    return true;
  });

  const byContract = new Map();
  const clientIds = new Set();

  for (const r of filtered) {
    const code = String(r[idx.OrderformCode] || "").trim();
    if (!byContract.has(code)) {
      byContract.set(code, {
        orderformCode: code,
        ref: r[idx.Ref] || "",
        startDate: r[idx.StartDate] || "",
        endDate: r[idx.EndDate] || "",
        paymentTerm: idx.PaymentTerm >= 0 ? (r[idx.PaymentTerm] || "") : "",
        lines: [],
        totalAmount: 0
      });
    }
    const line = {
      subProduct: r[idx.SubProduct] || "",
      comments:   idx.Comments >= 0 ? (r[idx.Comments] || "") : "",
      licenseOrdered: Number(r[idx.LicenseOrdered] || 0),
      price: Number(r[idx.Price] || 0),
      amount: Number(r[idx.Amount] || 0),
      accountManager: idx.AccountManager >= 0 ? (r[idx.AccountManager] || "") : ""
    };
    const obj = byContract.get(code);
    obj.lines.push(line);
    obj.totalAmount += line.amount;

    if (idx.ClientID >= 0) {
      const cid = String(r[idx.ClientID] || "").trim();
      if (cid) clientIds.add(cid);
    }
  }

  const contracts = Array.from(byContract.values())
    .map(c => ({ ...c, startDate: _asISO_(c.startDate), endDate: _asISO_(c.endDate) }))
    .sort((a,b) =>
      (a.startDate > b.startDate ? 1 : a.startDate < b.startDate ? -1 : a.orderformCode.localeCompare(b.orderformCode))
    );

  return {
    customer: customerName,
    clientId: Array.from(clientIds).join(", "),
    contractCount: contracts.length,
    contracts
  };
}
/*** BUTTON WRAPPERS (sheet-bound) ***/

// Big popup for the Contracts (Sidebar) UI
function showPopup() {
  // Replace 'Sidebar' with your file name if different
  var html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(1000)
    .setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Customer Contracts');
}

// Big popup for Renewals (Prospect.html)
function showUnsignedProspects() {
  // Your file name is singular: 'Prospect'
  var html = HtmlService.createHtmlOutputFromFile('Prospect')
    .setWidth(1100)
    .setHeight(720);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Renewals');
}

/*** (optional) quick sanity check ***/
function hello() {
  SpreadsheetApp.getUi().alert('Buttons are wired correctly!');
}
