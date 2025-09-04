/** 
 * Code.gs — Google Apps Script (bound to Google Sheet)
 * 
 * Shopify → Export GST CSV (Last Month or Specific Month)
 * 
 * Generates a Google Sheet in Drive for a selected month's Shopify orders with reverse GST calculation.
 * Adds custom menu items inside “🛠️ KC Tools”:
 *   - Export GST CSV (Last Month)
 *   - Export GST CSV (Pick Month…)
 * 
 * -----------------------------
 * CONFIGURATION (EDIT THESE)
 * -----------------------------
 */
const SHOPIFY_STORE_DOMAIN      = getCellValue("Automation", "H5") + ".myshopify.com"; // e.g., "example.myshopify.com"
const SHOPIFY_API_VERSION       = getCellValue("Automation", "H7");                    // e.g., "2024-01"
const SHOPIFY_ACCESS_TOKEN      = getCellValue("Automation", "H6");                    // Admin API access token
const INCLUDE_ONLY_PAID_ORDERS  = false;                                                // include only 'paid' or 'partially_paid'
const TIMEZONE                  = "Asia/Kolkata";                                       // all date math/formatting in this TZ

// Drive folder for output (prefer ID if present)
const CSV_FOLDER_ID             = getCellValue("Automation", "H9"); // leave empty to fallback
const CSV_FOLDER_NAME           = "GST Exports";                     // fallback only

// ✅ Email recipient (leave blank to default to the active user's email)
const REPORT_EMAIL              = getCellValue("Automation", "H8"); // e.g., "me@mycompany.com" or ""

// Simple in-memory cache per execution to reduce API calls
const _customerCache = Object.create(null);

/**
 * Register KC menu items (call this from your existing onOpen menu builder).
 * Adds both "Last Month" and "Pick Month…" actions.
 */
function kcRegisterGstCsvMenuItem(menu) {
  return menu
    .addSeparator()
    .addItem('Export GST CSV (Last Month)', 'exportGstCsvForLastMonth')
    .addItem('Export GST CSV (Pick Month…)', 'exportGstCsvPromptForMonth');
}

/* ============================================================================
 * MAIN ENTRIES
 * ==========================================================================*/

/**
 * One-click export for LAST calendar month (uses TIMEZONE).
 */
function exportGstCsvForLastMonth() {
  const { year, month } = getLastMonthYearMonth(TIMEZONE);
  exportGstCsvForMonth(year, month);
}

/**
 * Prompt the user for a specific month (YYYY-MM) and export that period.
 */
function exportGstCsvPromptForMonth() {
  const ui = SpreadsheetApp.getUi();
  const resp = ui.prompt('Export for specific month', 'Enter YYYY-MM (e.g., 2025-07):', ui.ButtonSet.OK_CANCEL);
  if (resp.getSelectedButton() !== ui.Button.OK) return;

  const val = (resp.getResponseText() || '').trim();
  const m = val.match(/^(\d{4})-(\d{2})$/);
  if (!m) {
    ui.alert('Invalid format. Please enter YYYY-MM, e.g., 2025-07.');
    return;
  }
  const year  = parseInt(m[1], 10);
  const month = parseInt(m[2], 10);
  if (month < 1 || month > 12) {
    ui.alert('Month must be between 01 and 12.');
    return;
  }
  exportGstCsvForMonth(year, month);
}

/**
 * Export for a specific month (1..12) and year (e.g., 2025).
 * Builds Google Sheet directly in the configured Drive folder.
 */
function exportGstCsvForMonth(year, month) {
  try {
    // 1) Compute the selected month's time window & labels
    const { startIso, endIso, monthNameUpper, todayYmd, year: y, month: m } = getMonthRange(year, month, TIMEZONE);

    // 2) Fetch all orders across pages
    const allOrders = fetchAllOrdersPaginated(startIso, endIso);

    // 3) Filter for paid/non-cancelled if configured
    const filtered = filterOrders(allOrders);

    // 3a) Split into normal vs returned
    const returnedOrders = filtered.filter(isReturnedOrder);
    const normalOrders   = filtered.filter(o => !isReturnedOrder(o));

    // 4) Build rows for Sheets (same columns)
    const rowsOrders  = buildSheetRows(normalOrders);
    const rowsReturns = buildSheetRows(returnedOrders);

    // 5) Save Google Sheet in Drive with required filename format:
    //    GST_Invoices_{YYYYMM}_{MONTH}-{YYYYMMDD}.csv (Sheet title)
    const yyyymm   = `${pad(y,4)}${pad(m,2)}`;
    const filename = `GST_Invoices_${yyyymm}_${monthNameUpper}-${todayYmd}.csv`;

    const folder  = resolveOutputFolder();
    const { url: fileUrl } = saveAsGoogleSheetInFolder(folder, filename, rowsOrders, rowsReturns);

    // 6) Notify user
    const msg = `Exported Orders: ${normalOrders.length}, Returns: ${returnedOrders.length} for ${monthNameUpper} ${y}.\nFile: ${filename}\nLink: ${fileUrl}`;
    SpreadsheetApp.getActive().toast(msg, 'Shopify GST Export', 10);

    // 7) Email the link
    const recipient = resolveReportRecipient();
    if (recipient) {
      const subject = `Kumud Creations GST Invoices Export: ${filename}`;
      const body = [
        `Hello Team,`,
        ``,
        `The GST Invoices export file has been generated successfully.`,
        ``,
        `Month: ${monthNameUpper} ${y}`,
        `Orders Exported: ${normalOrders.length}`,
        `Returns Exported: ${returnedOrders.length}`,
        `File: ${filename}`,
        `Link: ${fileUrl}`,
        ``,
        `Regards,`,
        `Kumud Creations`
      ].join('\n');
      MailApp.sendEmail(recipient, subject, body);
    } else {
      console.warn('No REPORT_EMAIL and could not resolve active user email; skipping email send.');
    }

  } catch (err) {
    const uiMsg = `Export failed: ${err && err.message ? err.message : err}`;
    SpreadsheetApp.getActive().toast(uiMsg, 'Shopify GST Export — Error', 10);
    console.error('Shopify GST Export — Error', uiMsg);
    throw err;
  }
}

/* Resolve email recipient: prefer REPORT_EMAIL; else active user */
function resolveReportRecipient() {
  const configured = (REPORT_EMAIL || '').trim();
  if (configured) return configured;
  try {
    const active = (Session.getActiveUser().getEmail() || '').trim();
    return active || '';
  } catch (e) {
    return '';
  }
}


/* ============================================================================
 * DATE HELPERS
 * ==========================================================================*/

/**
 * Returns the start/end RFC3339 for a given (year, month) in tz,
 * plus labels for filename.
 */
function getMonthRange(year, month, tz) {
  if (!year || !month) throw new Error('getMonthRange: year and month are required');
  if (month < 1 || month > 12) throw new Error('getMonthRange: month must be 1..12');

  const lastDay = new Date(year, month, 0).getDate(); // JS: passing (year, month, 0) gives last day of 'month'

  const monthNameUpper = Utilities.formatDate(new Date(year, month - 1, 1), tz, 'MMMM').toUpperCase();
  const todayYmd       = Utilities.formatDate(new Date(), tz, 'yyyyMMddHHmm');

  // Build timezone offset like +05:30
  const midMonth = new Date(year, month - 1, Math.min(15, lastDay));
  const offsetNoColon = Utilities.formatDate(midMonth, tz, 'Z');        // e.g., +0530
  const offset        = offsetNoColon.replace(/([+-])(\d{2})(\d{2})/, '$1$2:$3');

  const startIso = `${pad(year, 4)}-${pad(month, 2)}-01T00:00:00${offset}`;
  const endIso   = `${pad(year, 4)}-${pad(month, 2)}-${pad(lastDay, 2)}T23:59:59${offset}`;

  return { startIso, endIso, monthNameUpper, todayYmd, year, month };
}

/** Convenience: compute last month in tz. */
function getLastMonthYearMonth(tz) {
  const now = new Date();
  const nowY = parseInt(Utilities.formatDate(now, tz, 'yyyy'), 10);
  const nowM = parseInt(Utilities.formatDate(now, tz, 'M'), 10);
  const lastM = (nowM === 1) ? 12 : (nowM - 1);
  const lastY = (nowM === 1) ? (nowY - 1) : nowY;
  return { year: lastY, month: lastM };
}

/* ============================================================================
 * SHOPIFY FETCH / FILTER
 * ==========================================================================*/

/**
 * Fetches all Shopify orders in [startIso, endIso] with REST pagination (Link header / page_info).
 */
function fetchAllOrdersPaginated(startIso, endIso) {
  const base = `https://${SHOPIFY_STORE_DOMAIN}/admin/api/${SHOPIFY_API_VERSION}/orders.json`;
  const params = {
    created_at_min: startIso,
    created_at_max: endIso,
    status: 'any',
    limit: '250',
    // include customer, tags, refunds to detect returns
    fields: [
      'name','created_at','billing_address','shipping_address','line_items',
      'total_price','financial_status','cancelled_at','customer','tags','customer_id','refunds'
    ].join(',')
  };

  let url = base + '?' + toQuery(params);
  const all = [];

  while (url) {
    const resp = shopifyFetch(url);
    const json = JSON.parse(resp.getContentText());
    if (Array.isArray(json.orders)) {
      all.push.apply(all, json.orders);
    }

    maybeThrottle(resp);
    url = parseNextUrlFromLinkHeader(resp);
  }

  return all;
}

/**
 * Filters out cancelled orders; if INCLUDE_ONLY_PAID_ORDERS is true, keeps only paid/partially_paid.
 */
function filterOrders(orders) {
  const allowedFinancialStatuses = INCLUDE_ONLY_PAID_ORDERS ? new Set(['paid', 'partially_paid']) : null;
  return orders.filter(o => {
    if (o.cancelled_at) return false;
    if (allowedFinancialStatuses && !allowedFinancialStatuses.has(String(o.financial_status || '').toLowerCase())) {
      return false;
    }
    return true;
  });
}

/**
 * Determine if an order is "returned".
 * Priority:
 *  1) refunds[].refund_line_items[].restock_type === 'return'  -> returned (REST, precise)
 *  2) financial_status in ['refunded','partially_refunded'] AND no refunds array -> likely returned (fallback)
 *  3) tags contain 'return' -> store convention fallback
 * Notes:
 *  - 'voided' indicates an uncaptured authorization canceled (not a return).
 */
function isReturnedOrder(order) {
  // 1) Examine refunds for explicit returned line items
  if (Array.isArray(order.refunds) && order.refunds.length > 0) {
    for (const r of order.refunds) {
      const items = (r && (r.refund_line_items || r.refundLineItems)) || [];
      for (const rli of items) {
        const restockType = String((rli && (rli.restock_type || rli.restockType)) || '').toLowerCase();
        const qty = Number(rli && rli.quantity) || 0;
        if (restockType === 'return' && qty > 0) {
          return true;
        }
      }
    }
    // If refunds exist but none are marked returned, treat as NOT returned here
    // (they might be 'cancel' or 'no_restock' refunds)
  }

  // 2) Fallback to financial status if refunds are unavailable
  const fs = String(order.financial_status || '').toLowerCase();
  if ((fs === 'refunded' || fs === 'partially_refunded') && (!order.refunds || order.refunds.length === 0)) {
    return true;
  }

  // 3) Optional tag-based fallback (store-specific process)
  const tags = String(order.tags || '').toLowerCase();
  if (tags.includes('return')) return true;

  return false;
}

/* ============================================================================
 * CUSTOMER / ADDRESS RESOLUTION
 * ==========================================================================*/

/**
 * Fetches a customer by ID from Shopify (cached). Requires read_customers scope.
 */
function fetchCustomerById(customerId) {
  if (!customerId) return null;
  const key = String(customerId);
  if (_customerCache[key]) return _customerCache[key];

  const url = `https://${SHOPIFY_STORE_DOMAIN}/admin/api/${SHOPIFY_API_VERSION}/customers/${encodeURIComponent(key)}.json?fields=first_name,last_name,default_address,email,phone`;
  const resp = shopifyFetch(url);
  const code = resp.getResponseCode();
  if (code < 200 || code >= 300) {
    console.warn(`Customer fetch failed for ID ${key}: ${code} ${resp.getContentText()}`);
    return null;
  }
  const payload = JSON.parse(resp.getContentText());
  const customer = payload && payload.customer ? payload.customer : null;
  _customerCache[key] = customer;
  return customer;
}

/**
 * Returns { customer, billAddrObj, shipAddrObj }
 * - If order has no embedded customer, fetch by order.customer_id (or order.customer.id).
 * - billAddrObj preference: order.billing_address -> customer.default_address -> order.shipping_address
 * - shipAddrObj preference: order.shipping_address -> order.billing_address -> customer.default_address
 */
function resolveAddressesForOrder(order) {
  let customer = order.customer || null;

  // Try resolve via explicit id if customer object is missing or incomplete
  const customerId = (order.customer && order.customer.id) ? order.customer.id : (order.customer_id || null);
  if (!customer && customerId) {
    customer = fetchCustomerById(customerId);
  }

  const customerDefault = customer && customer.default_address ? customer.default_address : null;

  const billAddrObj = order.billing_address || customerDefault || order.shipping_address || null;
  const shipAddrObj = order.shipping_address || order.billing_address || customerDefault || null;

  return { customer, billAddrObj, shipAddrObj };
}

/* ============================================================================
 * SHEET ROWS / SAVE
 * ==========================================================================*/

/**
 * Loads last 2000 rows (excluding header) from "Qikink Orders" sheet and
 * returns a lookup { shopifyOrderNo -> customerName }.
 * Expected columns (header names): "Order No", "Customer Name"
 * If your header is positioned elsewhere, we derive indexes from header row.
 */
function loadCustomerNameLookup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName('Qikink Orders');
  if (!sh) {
    console.warn('Qikink Orders sheet not found');
    return {};
  }

  // Read the full header to detect positions
  const headerRow = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const orderNoIdx = headerRow.indexOf('Order No');
  const customerNameIdx = headerRow.indexOf('Customer Name');

  if (orderNoIdx < 0 || customerNameIdx < 0) {
    console.warn('Qikink Orders header must include "Order No" and "Customer Name"');
    return {};
  }

  const lastRow = sh.getLastRow();
  if (lastRow < 2) return {}; // no data

  // Load only last 2000 rows of the entire row width (so column indexes are valid)
  const startRow = Math.max(2, lastRow - 2000 + 1);
  const numRows  = lastRow - startRow + 1;
  const values   = sh.getRange(startRow, 1, numRows, sh.getLastColumn()).getValues();

  const lookup = {};
  for (let i = 0; i < values.length; i++) {
    const row = values[i];
    const orderNo = String(row[orderNoIdx] || '').trim();
    const custName = String(row[customerNameIdx] || '').trim();
    if (orderNo && custName) {
      // convert the Qikink OrderNo to Shopify OrderNo format
      lookup[convertQikinkOrderNumberToShopifyNumber(orderNo)] = custName;
    }
  }
  return lookup;
}

/**
 * Build all rows for the Google Sheet.
 */
function buildSheetRows(orders) {
  const customerNameLookup = loadCustomerNameLookup(); // load once  

  const headers = [
    'Order Number',
    'Order Date',
    'Bill To Address',
    'Ship To Address',
    'Item Details',
    'Gross Total amount',
    'GST amount',
    'IGST amount',
    'CGST amount',
    'SGST amount',
    'Total (Excluding GST amount)',
    'Total (Including GST amount)',
    'Order Tags'
  ];

  const rows = [headers];

  for (const o of orders) {
    const orderName = ensureKCOrderName(o.name);
    const orderDate = Utilities.formatDate(new Date(o.created_at), TIMEZONE, 'yyyy-MMM-dd');

    const { customer, billAddrObj, shipAddrObj } = resolveAddressesForOrder(o);

    let billTo = formatAddress(billAddrObj, customer);
    let shipTo = formatAddress(shipAddrObj, customer);

    // Prepend mapped customer name if available
    const mappedName = customerNameLookup[orderName] || '';
    if (mappedName) {
      billTo = mappedName + (billTo ? ', ' + billTo : '');
      shipTo = mappedName + (shipTo ? ', ' + shipTo : '');
    }

    // For GST state logic, derive from bill-to first, else fallback to customer default
    const billStateRaw = (billAddrObj && (safeStr(billAddrObj.province) || safeStr(billAddrObj.province_code)))
      || (customer && customer.default_address && (safeStr(customer.default_address.province) || safeStr(customer.default_address.province_code)))
      || '';
    const billState = billStateRaw.toLowerCase();

    const itemsList = Array.isArray(o.line_items) ? o.line_items.map(li => {
      const vt = (li.variant_title && String(li.variant_title).trim()) ? ` ${li.variant_title}` : '';
      const qty = li.quantity != null ? ` x${li.quantity}` : '';
      return `${li.title || ''}${vt}${qty}`.trim();
    }).join(', ') : '';

    const grossNum = parseFloat(o.total_price);
    if (!isFinite(grossNum)) {
      console.warn(`Skipping order ${o.name}: invalid total_price "${o.total_price}"`);
      continue;
    }

    const { net, gst, igst, cgst, sgst } = computeGstBreakdown(grossNum, billState);
    const orderTags = safeStr(o.tags);  // Shopify returns tags as comma-separated string

    rows.push([
      orderName,
      orderDate,
      billTo,
      shipTo,
      itemsList,
      formatNumber2(grossNum),
      formatNumber2(gst),
      igst ? formatNumber2(igst) : '',
      cgst ? formatNumber2(cgst) : '',
      sgst ? formatNumber2(sgst) : '',
      formatNumber2(net),
      formatNumber2(grossNum),
      orderTags
    ]);
  }

  return rows;
}

/**
 * Create a Google Sheet directly in a given Drive folder and write rows to:
 *  - "Orders"   (normal orders)
 *  - "Returns"  (returned orders)
 * Returns { url, spreadsheetId }.
 */
function saveAsGoogleSheetInFolder(folder, filename, rowsOrders, rowsReturns) {
  const ss = SpreadsheetApp.create(filename.replace(/\.csv$/i, '')); // drop .csv if present for Sheet title
  const file = DriveApp.getFileById(ss.getId());

  // Move to target folder (single move)
  file.moveTo(folder);

  // ORDERS sheet (rename default to "Orders")
  const ordersSheet = ss.getActiveSheet();
  ordersSheet.setName('Orders');
  ordersSheet.clear();
  if (rowsOrders && rowsOrders.length) {
    ordersSheet.getRange(1, 1, rowsOrders.length, rowsOrders[0].length).setValues(rowsOrders);
    ordersSheet.autoResizeColumns(1, rowsOrders[0].length);
  }

  // RETURNS sheet (create and populate)
  let returnsSheet = ss.getSheetByName('Returns');
  if (!returnsSheet) returnsSheet = ss.insertSheet('Returns');
  returnsSheet.clear();

  const rowsR = (rowsReturns && rowsReturns.length) ? rowsReturns : [ (rowsOrders && rowsOrders[0]) ? rowsOrders[0] : [] ];
  if (rowsR.length && rowsR[0].length) {
    returnsSheet.getRange(1, 1, rowsR.length, rowsR[0].length).setValues(rowsR);
    returnsSheet.autoResizeColumns(1, rowsR[0].length);
  }

  return { url: ss.getUrl(), spreadsheetId: ss.getId() };
}

/* ============================================================================
 * NAME / ADDRESS FORMATTERS
 * ==========================================================================*/

/**
 * If order name does not start with "#KC", prepend it.
 * (Adjust if your store uses "KC" without "#")
 */
function ensureKCOrderName(name) {
  const s = String(name || '').trim();
  return /^#KC/i.test(s) ? s : `#KC${s}`;
}

function safeStr(v) { return (v == null) ? '' : String(v).trim(); }

function customerFullName(customer) {
  if (!customer) return '';
  const f = safeStr(customer.first_name);
  const l = safeStr(customer.last_name);
  return [f, l].filter(Boolean).join(' ').trim();
}

function addressFullName(addr) {
  if (!addr) return '';
  const n = safeStr(addr.name);
  if (n) return n;
  const f = safeStr(addr.first_name);
  const l = safeStr(addr.last_name);
  return [f, l].filter(Boolean).join(' ').trim();
}

/**
 * Formats address as: Customer Name, Company, Address1, Address2, City, Province/State, Zip, Country, Phone
 * - Customer Name comes from the Customer object if available, else from address fields.
 */
function formatAddress(addr, customer) {
  const leadName = customerFullName(customer) || addressFullName(addr);
  const province = addr ? (safeStr(addr.province) || safeStr(addr.province_code)) : '';
  const parts = [
    leadName,
    safeStr(addr && addr.company),
    safeStr(addr && addr.address1),
    safeStr(addr && addr.address2),
    safeStr(addr && addr.city),
    province,
    safeStr(addr && addr.zip),
    safeStr(addr && addr.country),
    safeStr(addr && addr.phone)
  ];

  const seen = new Set();
  const out = [];
  for (const p of parts) {
    const s = (p || '').trim();
    if (!s) continue;
    const key = s.toLowerCase();
    if (seen.has(key)) continue;
    seen.add(key);
    out.push(s);
  }
  return out.join(', ');
}

/* ============================================================================
 * GST LOGIC
 * ==========================================================================*/

/**
 * Reverse-calculates GST and labels as per custom rules:
 * - If gross <= 1000 → 5%, else 12%
 * - net = gross / (1 + rate), gst = gross - net
 * - If Bill-To state indicates Gujarat → IGST only
 *   else → split equally CGST + SGST
 * NOTES:
 * - State detection checks full name "Gujarat" OR code "GJ", case-insensitively.
 * - If billing state is missing/unknown, treat as NOT Gujarat (split CGST/SGST).
 */
function computeGstBreakdown(gross, billToStateOrCode) {
  const rate = (gross <= 1000) ? 0.05 : 0.12;

  let net = gross / (1 + rate);
  net = round2(net);
  let gst = round2(gross - net);

  const s = (billToStateOrCode || '').toLowerCase();
  const isGujarat = s === 'gujarat' || s === 'gj';

  let igst = 0, cgst = 0, sgst = 0;

  if (!isGujarat) {
    igst = gst;
  } else {
    cgst = round2(gst / 2);
    sgst = round2(gst - cgst); // keep sum consistent
  }

  return { rate, net, gst, igst, cgst, sgst };
}

/* ============================================================================
 * DRIVE HELPERS
 * ==========================================================================*/

function getFolderByIdStrict(id) {
  if (!id || String(id).trim() === '') {
    throw new Error('CSV_FOLDER_ID is not set. Please provide a valid Google Drive folder ID.');
  }
  try {
    const folder = DriveApp.getFolderById(id);
    folder.getName(); // force access check
    return folder;
  } catch (e) {
    throw new Error(`Invalid CSV_FOLDER_ID "${id}". Ensure the ID exists and this script has access. Details: ${e.message}`);
  }
}

function ensureOrCreateFolderByName(name) {
  const it = DriveApp.getFoldersByName(name);
  if (it.hasNext()) return it.next();
  return DriveApp.createFolder(name);
}

function resolveOutputFolder() {
  if (CSV_FOLDER_ID && String(CSV_FOLDER_ID).trim() !== '') {
    return getFolderByIdStrict(CSV_FOLDER_ID);
  }
  return ensureOrCreateFolderByName(CSV_FOLDER_NAME);
}

/* ============================================================================
 * HTTP / Shopify helpers
 * ==========================================================================*/

function shopifyFetch(url) {
  const options = {
    method: 'get',
    headers: {
      'X-Shopify-Access-Token': SHOPIFY_ACCESS_TOKEN,
      'Content-Type': 'application/json'
    },
    muteHttpExceptions: true
  };

  const maxAttempts = 6;
  let attempt = 0;
  while (true) {
    attempt++;
    const resp = UrlFetchApp.fetch(url, options);
    const code = resp.getResponseCode();

    if (code >= 200 && code < 300) {
      return resp;
    }

    // Handle rate limit or transient server errors with backoff
    if ((code === 429 || (code >= 500 && code < 600)) && attempt < maxAttempts) {
      const delayMs = Math.min(2000 * Math.pow(2, attempt - 1), 20000); // 2s, 4s, 8s, 16s, 20s...
      Utilities.sleep(delayMs);
      continue;
    }

    // Permanent error
    throw new Error(`Shopify request failed (${code}): ${resp.getContentText()}`);
  }
}

function maybeThrottle(resp) {
  const headers = resp.getAllHeaders ? resp.getAllHeaders() : resp.getHeaders();
  const h = headers || {};
  const limitRaw = h['X-Shopify-Shop-Api-Call-Limit'] || h['x-shopify-shop-api-call-limit'];
  if (limitRaw && typeof limitRaw === 'string') {
    const parts = limitRaw.split('/');
    if (parts.length === 2) {
      const used = parseInt(parts[0], 10);
      const cap  = parseInt(parts[1], 10) || 40;
      if (isFinite(used) && isFinite(cap) && used > cap - 5) {
        Utilities.sleep(800); // cool-off
      }
    }
  } else {
    Utilities.sleep(200); // small baseline delay between pages
  }
}

function parseNextUrlFromLinkHeader(resp) {
  const headers = resp.getAllHeaders ? resp.getAllHeaders() : resp.getHeaders();
  const link = headers['Link'] || headers['link'];
  if (!link) return null;

  // Example: <https://.../orders.json?limit=250&page_info=abcdef>; rel="next"
  const parts = link.split(',');
  for (const p of parts) {
    const m = p.match(/<([^>]+)>;\s*rel="next"/i);
    if (m && m[1]) {
      return m[1];
    }
  }
  return null;
}

/* ============================================================================
 * MISC UTILITIES
 * ==========================================================================*/

function formatNumber2(n) { return (Number(n).toFixed(2)); }
function round2(n)        { return Math.round(Number(n) * 100) / 100; }

function pad(num, width) {
  const s = String(num);
  return s.length >= width ? s : ('0'.repeat(width - s.length) + s);
}

function toQuery(obj) {
  const parts = [];
  for (const k in obj) {
    if (!Object.prototype.hasOwnProperty.call(obj, k)) continue;
    const v = obj[k];
    if (v === undefined || v === null) continue;
    parts.push(encodeURIComponent(k) + '=' + encodeURIComponent(String(v)));
  }
  return parts.join('&');
}

/**
 * Get a cell's display value (throws if sheet missing).
 */
function getCellValue(sheetName, a1Notation) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(sheetName);
  if (!sh) throw new Error(`Sheet "${sheetName}" not found for getCellValue(${a1Notation})`);
  return sh.getRange(a1Notation).getDisplayValue().trim();
}
