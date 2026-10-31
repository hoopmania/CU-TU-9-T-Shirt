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
  NAVY_SIZE:   10,  // สีกรมท่า CU — ไซส์
  WHITE_SIZE:  11,  // สีขาว TU — ไซส์
  QTY:         12,  // จำนวนตัว
  DELIVERY:    13,  // วิธีรับ
  ADDRESS:     14,  // ที่อยู่
  SHIP_FEE:    15,  // ค่าส่ง
  PRICE:       16,  // ยอดเสื้อ
  GRAND_TOTAL: 17,  // ยอดรวม
  STATUS:      18,  // สถานะ
  NOTE:        19,  // หมายเหตุ
  SLIP_URL:    20,  // ลิงค์สลิป
  TU_NAVY_SIZE:21,  // TU สีกรมท่า — ไซส์
};
const NUM_COLS = 21;
const SLIP_FOLDER_NAME = 'CU-TU-9-Slips';

const HEADERS = [
  'วันที่','เลขออร์เดอร์','ชื่อผู้สั่ง','โทร','LINE ID',
  'ตัวที่','สำหรับ','ทีม','รุ่น',
  'สีกรมท่า CU — ไซส์','สีขาว TU — ไซส์','จำนวนตัว',
  'วิธีรับ','ที่อยู่จัดส่ง','ค่าส่ง','ยอดเสื้อ','ยอดรวม',
  'สถานะ','หมายเหตุ','ลิงค์สลิป','TU สีกรมท่า — ไซส์'
];

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'order';
    if (action === 'updateStatus')     return handleUpdateStatus(payload);
    if (action === 'getOrders')        return handleGetOrders(payload);
    if (action === 'checkStatus')      return handleCheckStatus(payload);
    if (action === 'getOrdersByPhone') return handleGetOrdersByPhone(payload);
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

  // Upload slip to Drive (first row only)
  let slipUrl = '';
  if (payload.slipBase64) {
    try {
      const base64 = String(payload.slipBase64).replace(/^data:[^;]+;base64,/, '');
      const ext = (payload.slipFileName || 'slip.jpg').split('.').pop().toLowerCase();
      const mime = ext === 'png' ? 'image/png' : (ext === 'webp' ? 'image/webp' : 'image/jpeg');
      const blob = Utilities.newBlob(Utilities.base64Decode(base64), mime, orderId + '_slip.' + ext);
      const folders = DriveApp.getFoldersByName(SLIP_FOLDER_NAME);
      const folder = folders.hasNext() ? folders.next() : DriveApp.createFolder(SLIP_FOLDER_NAME);
      const file = folder.createFile(blob);
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      slipUrl = file.getUrl();
    } catch(slipErr) {
      slipUrl = 'ERROR: ' + slipErr.message;
      Logger.log('Slip upload error: ' + slipErr.message);
    }
  }
  const isShip = payload.delivery === 'ship';
  const shippingFee = isShip ? (payload.shippingFee || 50) : 0;

  const rows = [];
  entries.forEach((entry, i) => {
    const navySize   = (entry.navyOn  && entry.navySize)  ? entry.navySize  : '';
    const whiteSize  = (entry.whiteOn && entry.whiteSize) ? entry.whiteSize : '';
    const tuNavySize = (entry.tuOn    && entry.tuSize)    ? entry.tuSize    : '';
    if (!navySize && !whiteSize && !tuNavySize) return; // skip empty entry

    const qty = (navySize ? 1 : 0) + (whiteSize ? 1 : 0) + (tuNavySize ? 1 : 0);
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
    row[C.NAVY_SIZE    - 1] = navySize;
    row[C.WHITE_SIZE   - 1] = whiteSize;
    row[C.TU_NAVY_SIZE - 1] = tuNavySize;
    row[C.QTY          - 1] = qty;
    const deliveryLabel = payload.delivery === 'ship' ? 'จัดส่ง' : payload.delivery === 'combine' ? 'รวมส่ง' : 'รับเอง';
    row[C.DELIVERY     - 1] = deliveryLabel;
    row[C.ADDRESS     - 1] = isShip ? String(payload.address || '') : '';
    row[C.SHIP_FEE    - 1] = isFirstRow && isShip ? shippingFee : '';
    row[C.PRICE       - 1] = entryPrice;
    row[C.GRAND_TOTAL - 1] = isFirstRow ? (payload.grandTotal || 0) : '';
    row[C.STATUS      - 1] = 'รับคำสั่งซื้อแล้ว';
    const combineNote = payload.combineOrderId ? `รวมกับ ${payload.combineOrderId}` : '';
    const noteVal = [String(payload.note || ''), combineNote].filter(Boolean).join(' | ');
    row[C.NOTE        - 1] = isFirstRow ? noteVal : '';
    row[C.SLIP_URL    - 1] = isFirstRow ? slipUrl : '';
    rows.push(row);
  });

  if (rows.length > 0) {
    const firstRow = sheet.getLastRow() + 1;
    // Set phone column to text BEFORE writing values, so Sheets doesn't strip leading 0
    sheet.getRange(firstRow, C.PHONE, rows.length, 1).setNumberFormat('@');
    sheet.getRange(firstRow, 1, rows.length, NUM_COLS).setValues(rows);
  }

  return jsonResp({
    status: 'ok', orderId, slipUrl,
    _debug: {
      rowsWritten: rows.length,
      slipBase64Len: payload.slipBase64 ? String(payload.slipBase64).length : 0,
      slipFileName: payload.slipFileName || null,
      slipUrl: slipUrl || null,
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
      navySize:    String(r[C.NAVY_SIZE-1]    || ''),
      whiteSize:   String(r[C.WHITE_SIZE-1]  || ''),
      tuNavySize:  String(r[C.TU_NAVY_SIZE-1]|| ''),
      qty:         r[C.QTY-1]          || 0,
      delivery:    String(r[C.DELIVERY-1]    || ''),
      address:     String(r[C.ADDRESS-1]     || ''),
      shipFee:     r[C.SHIP_FEE-1]     || '',
      price:       r[C.PRICE-1]        || 0,
      grandTotal:  r[C.GRAND_TOTAL-1]  || '',
      status:      String(r[C.STATUS-1]      || ''),
      note:        String(r[C.NOTE-1]        || ''),
    });
  }
  return jsonResp({ status: 'ok', rows });
}

// ─── Get shippable orders by phone (for combine-delivery lookup) ─────────────
function handleGetOrdersByPhone(payload) {
  const phone = String(payload.phone || '').trim();
  if (!phone) return jsonResp({ status: 'ok', orders: [] });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'ok', orders: [] });
  const data = sheet.getDataRange().getValues();
  const phoneAlt = phone.startsWith('0') ? phone.slice(1) : '0' + phone;

  // Statuses that mean the order is still open (not yet shipped/delivered)
  const CLOSED = ['จัดส่งเรียบร้อยแล้ว','รับเสื้อเรียบร้อยแล้ว'];

  const seen = {};
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const rowPhone = String(r[C.PHONE-1] || '');
    if (rowPhone !== phone && rowPhone !== phoneAlt) continue;
    const delivery = String(r[C.DELIVERY-1] || '');
    if (delivery !== 'จัดส่ง') continue; // only ship orders can be combined
    const status = String(r[C.STATUS-1] || '');
    if (CLOSED.indexOf(status) !== -1) continue; // skip already-shipped
    const orderId = String(r[C.ORDER_ID-1] || '');
    if (!orderId || seen[orderId]) continue;
    seen[orderId] = true;

    // Count shirts in this order
    let navyCnt = 0, whiteCnt = 0, tuCnt = 0;
    for (let j = 1; j < data.length; j++) {
      if (String(data[j][C.ORDER_ID-1]) !== orderId) continue;
      if (data[j][C.NAVY_SIZE-1]) navyCnt++;
      if (data[j][C.WHITE_SIZE-1]) whiteCnt++;
      if (data[j][C.TU_NAVY_SIZE-1]) tuCnt++;
    }
    const parts = [];
    if (navyCnt) parts.push('CU กรม×' + navyCnt);
    if (whiteCnt) parts.push('TU ขาว×' + whiteCnt);
    if (tuCnt)    parts.push('TU กรม×' + tuCnt);

    const ts = r[C.TIMESTAMP-1] ? new Date(r[C.TIMESTAMP-1]) : null;
    const dateStr = ts ? ts.toLocaleDateString('th-TH', {day:'2-digit',month:'2-digit',year:'numeric'}) : '';

    seen[orderId] = { orderId, shirtSummary: parts.join(' '), dateStr };
  }
  const orders = Object.values(seen).filter(v => typeof v === 'object');
  return jsonResp({ status: 'ok', orders });
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
