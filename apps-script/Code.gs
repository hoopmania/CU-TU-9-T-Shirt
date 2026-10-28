// CU-TU Basketball Alumni #9 — Apps Script
// 1 row per entry (one person's order), navy+white on same row

const SHEET_NAME = 'Orders';
const ADMIN_PASS = 'cutuba9';

// Columns (1-based)
const C = {
  TIMESTAMP:   1,   // วันที่
  ORDER_ID:    2,   // เลขออร์เดอร์
  BUYER_NAME:  3,   // ชื่อผู้สั่ง
  PHONE:       4,   // โทร
  LINE_ID:     5,   // LINE ID
  ENTRY_NO:    6,   // ตัวที่
  FOR_WHOM:    7,   // สำหรับ
  TEAM:        8,   // ทีม
  DIVISION:    9,   // รุ่น
  NAVY_SIZE:   10,  // สีกรม — ไซส์
  WHITE_SIZE:  11,  // สีขาว — ไซส์
  QTY:         12,  // จำนวนตัว
  DELIVERY:    13,  // วิธีรับ
  ADDRESS:     14,  // ที่อยู่
  SHIP_FEE:    15,  // ค่าส่ง
  PRICE:       16,  // ยอดเสื้อ
  GRAND_TOTAL: 17,  // ยอดรวม
  STATUS:      18,  // สถานะ
  NOTE:        19,  // หมายเหตุ
};
const NUM_COLS = 19;

const HEADERS = [
  'วันที่','เลขออร์เดอร์','ชื่อผู้สั่ง','โทร','LINE ID',
  'ตัวที่','สำหรับ','ทีม','รุ่น',
  'สีกรม — ไซส์','สีขาว — ไซส์','จำนวนตัว',
  'วิธีรับ','ที่อยู่จัดส่ง','ค่าส่ง','ยอดเสื้อ','ยอดรวม',
  'สถานะ','หมายเหตุ'
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'order';
    if (action === 'updateStatus') return handleUpdateStatus(payload);
    if (action === 'getOrders')    return handleGetOrders(payload);
    if (action === 'checkStatus')  return handleCheckStatus(payload);
    // Debug log — remove after fixing
    Logger.log('PAYLOAD: ' + JSON.stringify({
      action, orderId: payload.orderId, buyerName: payload.buyerName,
      phone: payload.phone, delivery: payload.delivery,
      address: payload.address, shippingFee: payload.shippingFee,
      grandTotal: payload.grandTotal,
      entries: (payload.entries || []).map(e => ({
        forWhom: e.forWhom, team: e.team, division: e.division,
        navyOn: e.navyOn, navySize: e.navySize,
        whiteOn: e.whiteOn, whiteSize: e.whiteSize,
      })),
    }));
    const result = handleOrder(payload);
    return result;
  } catch (err) {
    return jsonResp({ status: 'error', message: String(err.message) });
  }
}

function doGet(e) {
  return jsonResp({ status: 'ok', message: 'Use POST' });
}

// ─── Setup ───────────────────────────────────────────────────────────────────
function setupSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(SHEET_NAME);

  const hRange = sheet.getRange(1, 1, 1, NUM_COLS);
  hRange.setValues([HEADERS]);
  hRange.setFontWeight('bold')
        .setBackground('#0a1f5c')
        .setFontColor('#ffffff')
        .setHorizontalAlignment('center');
  sheet.setFrozenRows(1);

  // Force phone column to text so leading 0 is preserved
  sheet.getRange(2, C.PHONE, sheet.getMaxRows() - 1, 1).setNumberFormat('@');

  // Column widths
  sheet.setColumnWidth(C.TIMESTAMP, 140);
  sheet.setColumnWidth(C.ORDER_ID,  120);
  sheet.setColumnWidth(C.BUYER_NAME,140);
  sheet.setColumnWidth(C.PHONE,     110);
  sheet.setColumnWidth(C.ADDRESS,   200);

  SpreadsheetApp.flush();
  Logger.log('setupSheet done — ' + SHEET_NAME);
}

// ─── Order ────────────────────────────────────────────────────────────────────
function handleOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.getRange(1, 1, 1, NUM_COLS).setValues([HEADERS])
         .setFontWeight('bold').setBackground('#0a1f5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const ts = new Date();
  const orderId = String(payload.orderId || ('CU' + ts.getTime().toString().slice(-8)));
  const entries = payload.entries || [];
  const isShip = payload.delivery === 'ship';
  const shippingFee = isShip ? (payload.shippingFee || 50) : 0;

  const rows = [];
  entries.forEach((entry, i) => {
    const navySize  = (entry.navyOn  && entry.navySize)  ? entry.navySize  : '';
    const whiteSize = (entry.whiteOn && entry.whiteSize) ? entry.whiteSize : '';
    if (!navySize && !whiteSize) return; // skip empty entry

    const qty = (navySize ? 1 : 0) + (whiteSize ? 1 : 0);
    const entryPrice = qty * 300;
    const isFirstRow = rows.length === 0;

    const row = new Array(NUM_COLS).fill('');
    row[C.TIMESTAMP   - 1] = ts;
    row[C.ORDER_ID    - 1] = orderId;
    row[C.BUYER_NAME  - 1] = String(payload.buyerName || '');
    row[C.PHONE       - 1] = String(payload.phone || '');  // stored as text via setNumberFormat
    row[C.LINE_ID     - 1] = String(payload.lineId || '');
    row[C.ENTRY_NO    - 1] = i + 1;
    row[C.FOR_WHOM    - 1] = String(entry.forWhom || '');
    row[C.TEAM        - 1] = String(entry.team || '');
    row[C.DIVISION    - 1] = String(entry.division || '');
    row[C.NAVY_SIZE   - 1] = navySize;
    row[C.WHITE_SIZE  - 1] = whiteSize;
    row[C.QTY         - 1] = qty;
    row[C.DELIVERY    - 1] = isShip ? 'จัดส่ง' : 'รับเอง';
    row[C.ADDRESS     - 1] = isShip ? String(payload.address || '') : '';
    row[C.SHIP_FEE    - 1] = isFirstRow && isShip ? shippingFee : '';
    row[C.PRICE       - 1] = entryPrice;
    row[C.GRAND_TOTAL - 1] = isFirstRow ? (payload.grandTotal || 0) : '';
    row[C.STATUS      - 1] = 'รับคำสั่งซื้อแล้ว';
    row[C.NOTE        - 1] = isFirstRow ? String(payload.note || '') : '';
    rows.push(row);
  });

  if (rows.length > 0) {
    const firstRow = sheet.getLastRow() + 1;
    // Set phone column to text BEFORE writing values, so Sheets doesn't strip leading 0
    sheet.getRange(firstRow, C.PHONE, rows.length, 1).setNumberFormat('@');
    sheet.getRange(firstRow, 1, rows.length, NUM_COLS).setValues(rows);
  }

  return jsonResp({
    status: 'ok', orderId,
    _debug: {
      rowsWritten: rows.length,
      payloadReceived: {
        orderId: payload.orderId, buyerName: payload.buyerName,
        phone: payload.phone, lineId: payload.lineId,
        delivery: payload.delivery, shippingFee: payload.shippingFee,
        grandTotal: payload.grandTotal,
        entryCount: entries.length,
        entry0: entries[0] ? {
          forWhom: entries[0].forWhom, team: entries[0].team, division: entries[0].division,
          navyOn: entries[0].navyOn, navySize: entries[0].navySize,
          whiteOn: entries[0].whiteOn, whiteSize: entries[0].whiteSize,
        } : null,
      },
    },
  });
}

// ─── Update status ────────────────────────────────────────────────────────────
function handleUpdateStatus(payload) {
  if (payload.pass !== ADMIN_PASS) return jsonResp({ status: 'error', message: 'Unauthorized' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'error', message: 'Sheet not found' });
  const row = parseInt(payload.rowIndex);
  if (isNaN(row) || row < 2) return jsonResp({ status: 'error', message: 'Invalid row' });
  sheet.getRange(row, C.STATUS).setValue(String(payload.status || ''));
  return jsonResp({ status: 'ok' });
}

// ─── Get all orders (admin) ───────────────────────────────────────────────────
function handleGetOrders(payload) {
  if (payload.pass !== ADMIN_PASS) return jsonResp({ status: 'error', message: 'Unauthorized' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'ok', rows: [] });
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[C.ORDER_ID - 1]) continue; // skip empty rows
    rows.push({
      rowIndex:   i + 1,
      timestamp:  r[C.TIMESTAMP-1]   ? r[C.TIMESTAMP-1].toString() : '',
      orderId:    String(r[C.ORDER_ID-1]    || ''),
      buyerName:  String(r[C.BUYER_NAME-1]  || ''),
      phone:      String(r[C.PHONE-1]       || ''),
      lineId:     String(r[C.LINE_ID-1]     || ''),
      entryNo:    r[C.ENTRY_NO-1]    || '',
      forWhom:    String(r[C.FOR_WHOM-1]    || ''),
      team:       String(r[C.TEAM-1]        || ''),
      division:   String(r[C.DIVISION-1]    || ''),
      navySize:   String(r[C.NAVY_SIZE-1]   || ''),
      whiteSize:  String(r[C.WHITE_SIZE-1]  || ''),
      qty:        r[C.QTY-1]         || 0,
      delivery:   String(r[C.DELIVERY-1]    || ''),
      address:    String(r[C.ADDRESS-1]     || ''),
      shipFee:    r[C.SHIP_FEE-1]    || '',
      price:      r[C.PRICE-1]       || 0,
      grandTotal: r[C.GRAND_TOTAL-1] || '',
      status:     String(r[C.STATUS-1]      || ''),
    });
  }
  return jsonResp({ status: 'ok', rows });
}

// ─── Check status (customer) ─────────────────────────────────────────────────
function handleCheckStatus(payload) {
  const q = String(payload.q || '').trim();
  if (!q) return jsonResp({ status: 'ok', rows: [] });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'ok', rows: [] });
  const data = sheet.getDataRange().getValues();
  const ql = q.toLowerCase();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (!r[C.ORDER_ID - 1]) continue;
    const name  = String(r[C.BUYER_NAME-1] || '').toLowerCase();
    const phone = String(r[C.PHONE-1]      || '');
    // match with or without leading 0
    const phoneAlt = phone.startsWith('0') ? phone.slice(1) : '0' + phone;
    if (name.includes(ql) || phone.includes(q) || phoneAlt.includes(q)) {
      rows.push({
        rowIndex:   i + 1,
        orderId:    String(r[C.ORDER_ID-1]    || ''),
        buyerName:  String(r[C.BUYER_NAME-1]  || ''),
        phone:      String(r[C.PHONE-1]       || ''),
        entryNo:    r[C.ENTRY_NO-1]    || '',
        forWhom:    String(r[C.FOR_WHOM-1]    || ''),
        team:       String(r[C.TEAM-1]        || ''),
        division:   String(r[C.DIVISION-1]    || ''),
        navySize:   String(r[C.NAVY_SIZE-1]   || ''),
        whiteSize:  String(r[C.WHITE_SIZE-1]  || ''),
        qty:        r[C.QTY-1]         || 0,
        delivery:   String(r[C.DELIVERY-1]    || ''),
        address:    String(r[C.ADDRESS-1]     || ''),
        grandTotal: r[C.GRAND_TOTAL-1] || '',
        status:     String(r[C.STATUS-1]      || ''),
      });
    }
  }
  return jsonResp({ status: 'ok', rows });
}

// ─── Helper ───────────────────────────────────────────────────────────────────
function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
