/***** CONFIG *****/
const SHEET_NAME = "Orderform"; // change if needed
const HEADERS = {
  OrderformCode: "OrderformCode", // Col A
  CustomerName:  "CustomerName",  // Col C
  Ref:           "Ref.",          // Col D
  StartDate:     "StartDate",     // Col E
  EndDate:       "EndDate",       // Col F
  SubProduct:    "SubProduct",    // Col I
  PaymentTerm:   "PaymentTerm",   // Col J
  LicenseOrdered:"LicenseOrdered",// Col K
  Price:         "Price",         // Col N
  Amount:        "Amount"         // Col O
};

/***** MENU *****/
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Contracts')
    .addItem('Open (Sidebar)', 'showSidebar')
    .addItem('Open (Large Popup)', 'showPopup')   // new
    .addToUi();
}

function showSidebar() {
  const html = HtmlService.createHtmlOutputFromFile("Sidebar")
    .setTitle("Customer Contracts");
  SpreadsheetApp.getUi().showSidebar(html);
}

function showPopup() {
  const html = HtmlService.createHtmlOutputFromFile('Sidebar')
    .setWidth(1000)   // <- adjust width in pixels
    .setHeight(700);  // <- adjust height in pixels
  SpreadsheetApp.getUi().showModelessDialog(html, 'Customer Contracts'); // or showModalDialog
}

/***** DATA HELPERS *****/
function _getSheet() {
  const sh = SpreadsheetApp.getActive().getSheetByName(SHEET_NAME);
  if (!sh) throw new Error(`Sheet "${SHEET_NAME}" not found`);
  return sh;
}

function _getHeaderIndexes(headerRow) {
  const map = {};
  headerRow.forEach((h, i) => (map[h] = i));
  // sanity check – required headers
  const required = Object.values(HEADERS);
  const missing = required.filter(h => !(h in map));
  if (missing.length) {
    throw new Error("Missing required headers: " + missing.join(", "));
  }
  return map;
}

function _getAllRows() {
  const sh = _getSheet();
  const values = sh.getDataRange().getValues(); // includes header row
  if (values.length < 2) return { headers: [], rows: [] , idx:{} };
  const headers = values[0];
  const idx = _getHeaderIndexes(headers);
  const rows = values.slice(1).filter(r => String(r.join("")).trim() !== ""); // non-empty rows
  return { headers, rows, idx };
}

/***** API: customers *****/
function getCustomers() {
  const { rows, idx } = _getAllRows();
  const set = new Set();
  rows.forEach(r => {
    const name = r[idx[HEADERS.CustomerName]];
    if (name && String(name).trim() !== "") set.add(String(name).trim());
  });
  return Array.from(set).sort((a,b)=>a.localeCompare(b));
}

/***** API: contracts for a customer *****/
function getCustomerContracts(customerName) {
  const { rows, idx } = _getAllRows();
  const filtered = rows.filter(r => String(r[idx[HEADERS.CustomerName]]).trim() === String(customerName).trim());

  // group by OrderformCode
  const byContract = new Map();
  filtered.forEach(r => {
    const code = String(r[idx[HEADERS.OrderformCode]]).trim();
    if (!byContract.has(code)) {
      byContract.set(code, {
        orderformCode: code,
        ref: r[idx[HEADERS.Ref]] || "",
        startDate: r[idx[HEADERS.StartDate]] || "",
        endDate: r[idx[HEADERS.EndDate]] || "",
        paymentTerm: r[idx[HEADERS.PaymentTerm]] || "",
        lines: [],
        totalAmount: 0
      });
    }
    const line = {
      subProduct: r[idx[HEADERS.SubProduct]] || "",
      licenseOrdered: Number(r[idx[HEADERS.LicenseOrdered]] || 0),
      price: Number(r[idx[HEADERS.Price]] || 0),
      amount: Number(r[idx[HEADERS.Amount]] || 0)
    };
    const obj = byContract.get(code);
    obj.lines.push(line);
    obj.totalAmount += line.amount;
  });

  // format dates on server (return ISO, client will pretty-print)
  function asISO(d) {
    if (Object.prototype.toString.call(d) === "[object Date]" && !isNaN(d)) {
      return Utilities.formatDate(d, Session.getScriptTimeZone(), "yyyy-MM-dd");
    }
    // sometimes dates are text; just pass through
    return d ? String(d) : "";
  }

  const contracts = Array.from(byContract.values()).map(c => ({
    ...c,
    startDate: asISO(c.startDate),
    endDate: asISO(c.endDate)
  }));

  // sort contracts by StartDate then OrderformCode
  contracts.sort((a,b) => (a.startDate > b.startDate ? 1 : a.startDate < b.startDate ? -1 : a.orderformCode.localeCompare(b.orderformCode)));

  return {
    customer: customerName,
    contractCount: contracts.length,
    currencyHint: "AED",
    contracts
  };
}

// Opens the large popup (modeless dialog)
function showPopup() {
  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.initialCustomer = ''; // no preselection
  const html = t.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Customer Contracts');
}

// Opens the popup and preselects the customer typed in ContractsStatus!B2
function openPopupFromStatus() {
  const sh = SpreadsheetApp.getActive().getSheetByName('ContractsStatus');
  if (!sh) throw new Error('Sheet "ContractsStatus" not found');
  const selected = (sh.getRange('B2').getDisplayValue() || '').trim();

  const t = HtmlService.createTemplateFromFile('Sidebar');
  t.initialCustomer = selected; // pass initial selection to HTML
  const html = t.evaluate().setWidth(1000).setHeight(700);
  SpreadsheetApp.getUi().showModelessDialog(html, 'Customer Contracts');
}
