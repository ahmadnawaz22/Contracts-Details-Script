/***** INVOICES CONFIG *****/
// Paste your data spreadsheet URL or ID for the file that has the "Invoices" tab:
const INV_DATA_SPREADSHEET_URL_OR_ID = "1KAj7Zu6IURxiwdjAmCYl4clIFzM2Ut5tW8EJA_9DIpg";
const INV_SHEET_NAME = "Invoices"; // the tab name that holds invoice rows

/***** Flexible header aliases (trim/spaces/case tolerated) *****/
const INV_HEADER_ALIASES = {
  Contract:        ["Contract", "Contract Code", "OrderformCode", "Order Form Code"],
  InvoiceNumber:   ["Invoice Number", "Invoice No", "Invoice #", "Invoice", "InvoiceNumber"],
  Customer:        ["Customer", "Customer Name", "CustomerName"],
  IssueDate:       ["Issue Date", "Issued On", "IssueDate"],
  DueDate:         ["Due Date", "Due Data", "DueDate"], // accepts the "Due Data" variant
  Description:     ["Description", "Desc"],
  AmountUSD:       ["Amount USD", "Amount (USD)", "AmountUSD", "Amount"],
  Balance:         ["Balance", "Outstanding", "Outstanding Amount", "Outstanding amount of invoice"],
  Tier:            ["Tier", "Status", "Invoice Status"],           // Receivable / Received
  Risk:            ["Risk"],                                       // NIL Value / Doubtfull
  AccountManager:  ["AccountManager", "Account Manager", "AM"],
  Issued:          ["Issued", "Is Issued", "Issued?"]
};

/***** Open the data spreadsheet (robust, matches your contracts resolver) *****/
function INV_resolveSpreadsheetId_() {
  const src = String(INV_DATA_SPREADSHEET_URL_OR_ID || "").trim();
  if (src) {
    const m = src.match(/[-\w]{25,}/);
    if (m) return m[0];
  }
  // Reuse your existing resolver if present
  try { if (typeof _resolveSpreadsheetId_ === "function") return _resolveSpreadsheetId_(); } catch(e){}
  try { if (typeof _openDataSS_ === "function") return _openDataSS_().getId(); } catch(e){}
  // Script property fallback (if you set it earlier)
  const prop = PropertiesService.getScriptProperties().getProperty("DATA_SOURCE_ID");
  if (prop) return prop;
  // Finally: active spreadsheet
  return SpreadsheetApp.getActive().getId();
}
function INV_openDataSS_() { return SpreadsheetApp.openById(INV_resolveSpreadsheetId_()); }
function INV_getSheet_() {
  const sh = INV_openDataSS_().getSheetByName(INV_SHEET_NAME);
  if (!sh) throw new Error(`Tab "${INV_SHEET_NAME}" not found in the data spreadsheet.`);
  return sh;
}

/***** Low-level helpers *****/
function INV_readHeaders_() {
  const sh = INV_getSheet_();
  return sh.getRange(1,1,1,sh.getLastColumn()).getValues()[0].map(h => String(h).trim());
}
function INV_indexFromAliases_(headers) {
  const map = {}; headers.forEach((h,i)=> map[String(h).trim()] = i);
  const idx = {};
  for (const [key, aliases] of Object.entries(INV_HEADER_ALIASES)) {
    idx[key] = -1;
    for (const a of aliases) { if (a in map) { idx[key] = map[a]; break; } }
  }
  // Require the critical fields
  const required = ["Contract","InvoiceNumber","Customer","IssueDate","DueDate","Description","AmountUSD","Balance","Tier","Risk","AccountManager","Issued"];
  const missing = required.filter(k => idx[k] < 0);
  if (missing.length) throw new Error("Missing headers on 'Invoices' tab: " + missing.join(", "));
  return idx;
}
function INV_readRows_(maxCol) {
  const sh = INV_getSheet_();
  const lr = sh.getLastRow();
  if (lr < 2) return [];
  return sh.getRange(2,1,lr-1,maxCol).getValues();
}
function INV_asISO_(v) {
  if (Object.prototype.toString.call(v) === "[object Date]" && !isNaN(v)) {
    return Utilities.formatDate(v, Session.getScriptTimeZone(), "yyyy-MM-dd");
  }
  return v ? String(v) : "";
}
function INV_yes_(v) {
  const s = String(v).trim().toLowerCase();
  return s === "yes" || s === "y" || s === "true" || s === "1";
}
// Replaced existing INV_normRisk_ with this version
function INV_normRisk_(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return "";                   // blank = "Expected" bucket
  if (s.startsWith("doubt")) return "Doubtful";   // normalize spelling
  if (s.startsWith("nil"))   return "NIL Value";  // keep as-is; not a dedicated filter option
  return String(v).trim();
}
function INV_normTier_(v) {
  const s = String(v || "").trim().toLowerCase();
  if (!s) return "";
  // Detect "Received" first so it doesn't get swallowed by "receiv..."
  if (s.startsWith("received") || s.startsWith("recv")) return "Received";
  if (s.startsWith("receivable") || s.includes("receivable value")) return "Receivable";
  // Fallbacks if wording is inside the string
  if (s.includes(" received")) return "Received";
  if (s.includes(" receivable")) return "Receivable";
  return String(v).trim();
}

/***** PUBLIC: data for Invoices.html (Issued=Yes only) *****/
function getInvoicesIssued() {
  const headers = INV_readHeaders_();
  const idx = INV_indexFromAliases_(headers);

  const needCols = Object.values(idx).filter(i => i >= 0).map(i => i + 1);
  const maxCol = Math.max.apply(null, needCols);
  const rows = INV_readRows_(maxCol);

  const items = [];
  for (const r of rows) {
    if (!INV_yes_(r[idx.Issued])) continue; // enforce "Issued = Yes"
    items.push({
      invoiceNumber:   String(r[idx.InvoiceNumber] || "").trim(),
      contract:        String(r[idx.Contract] || "").trim(),
      customer:        String(r[idx.Customer] || "").trim(),
      issueDate:       INV_asISO_(r[idx.IssueDate]),
      dueDate:         INV_asISO_(r[idx.DueDate]),
      description:     String(r[idx.Description] || ""),
      amountUSD:       Number(String(r[idx.AmountUSD] || "0").replace(/,/g,"")),
      balance:         Number(String(r[idx.Balance]   || "0").replace(/,/g,"")),
      tier:            INV_normTier_(r[idx.Tier] || ""),
      risk:            INV_normRisk_(r[idx.Risk] || ""),
      accountManager:  String(r[idx.AccountManager] || "").trim(),
      issued:          true
    });
  }

  // Sort by Customer, then DueDate asc
  items.sort((a,b)=>{
    const c = String(a.customer).localeCompare(String(b.customer));
    if (c !== 0) return c;
    const ad = a.dueDate || "9999-12-31";
    const bd = b.dueDate || "9999-12-31";
    if (ad !== bd) return ad < bd ? -1 : 1;
    return String(a.invoiceNumber).localeCompare(String(b.invoiceNumber));
  });

  // Small debug block to help the UI show what's happening if zero rows
  return {
    invoices: items,
    debug: {
      spreadsheetId: INV_resolveSpreadsheetId_(),
      sheet: INV_SHEET_NAME,
      headersSeen: headers,
      totalRows: rows.length,
      issuedYes: items.length
    }
  };
}

/***** (optional) button/menu opener *****/
function showInvoices() {
  const html = HtmlService.createHtmlOutputFromFile('Invoices')
    .setWidth(1100)
    .setHeight(720);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Invoices');
}
