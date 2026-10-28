// CU-TU Basketball Alumni #9 — Apps Script
// Each order stores one row per entry (per shirt per person)

const SHEET_NAME = 'Orders';
const ADMIN_PASS = 'cutuba9';

// Column indices (1-based)
const COL = {
  TIMESTAMP: 1,
  ORDER_ID: 2,
  BUYER_NAME: 3,
  PHONE: 4,
  LINE_ID: 5,
  DELIVERY: 6,
  ADDRESS: 7,
  GRAND_TOTAL: 8,  // only on first row of order
  ENTRY_INDEX: 9,
  FOR_WHOM: 10,
  TEAM: 11,
  DIVISION: 12,
  COLOR: 13,
  SIZE: 14,
  PRICE: 15,
  STATUS: 16,
  NOTE: 17,
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'order';

    if (action === 'updateStatus') {
      return handleUpdateStatus(payload);
    }

    return handleOrder(payload);
  } catch (err) {
    return jsonResp({ status: 'error', message: err.message });
  }
}

function handleOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['วันที่', 'เลขออร์เดอร์', 'ชื่อผู้สั่ง', 'โทร', 'LINE ID', 'วิธีรับ', 'ที่อยู่', 'ยอดรวม', 'ตัวที่', 'สำหรับ', 'ทีม', 'รุ่น', 'สีเสื้อ', 'ไซส์', 'ราคา', 'สถานะ', 'หมายเหตุ'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#0a1f5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const ts = new Date();
  const orderId = payload.orderId || ('CU' + ts.getTime().toString().slice(-8));
  const entries = payload.entries || [];

  // Save slip to Drive if provided
  if (payload.slipBase64) {
    try {
      const folder = getOrCreateFolder('CU-TU-9-Slips');
      const blob = Utilities.newBlob(Utilities.base64Decode(payload.slipBase64), 'image/jpeg', orderId + '_' + (payload.slipFileName || 'slip.jpg'));
      folder.createFile(blob);
    } catch (e) { /* ignore slip save error */ }
  }

  // Write one row per entry
  entries.forEach((entry, i) => {
    const row = new Array(17).fill('');
    row[COL.TIMESTAMP - 1] = ts;
    row[COL.ORDER_ID - 1] = orderId;
    row[COL.BUYER_NAME - 1] = payload.buyerName || '';
    row[COL.PHONE - 1] = payload.phone || '';
    row[COL.LINE_ID - 1] = payload.lineId || '';
    row[COL.DELIVERY - 1] = payload.delivery || '';
    row[COL.ADDRESS - 1] = payload.address || '';
    row[COL.GRAND_TOTAL - 1] = i === 0 ? (payload.grandTotal || 0) : '';
    row[COL.ENTRY_INDEX - 1] = i + 1;
    row[COL.FOR_WHOM - 1] = entry.forWhom || '';
    row[COL.TEAM - 1] = entry.team || '';
    row[COL.DIVISION - 1] = entry.division || '';
    row[COL.COLOR - 1] = entry.color || '';
    row[COL.SIZE - 1] = entry.size || '';
    row[COL.PRICE - 1] = entry.price || 300;
    row[COL.STATUS - 1] = payload.delivery === 'ship' ? 'ส่งที่บ้าน' : 'ยังไม่รับ';
    row[COL.NOTE - 1] = i === 0 ? (payload.note || '') : '';
    sheet.appendRow(row);
  });

  return jsonResp({ status: 'ok', orderId });
}

function handleUpdateStatus(payload) {
  if (payload.pass !== ADMIN_PASS) return jsonResp({ status: 'error', message: 'Unauthorized' });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'error', message: 'Sheet not found' });

  const rowIndex = parseInt(payload.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) return jsonResp({ status: 'error', message: 'Invalid row' });

  sheet.getRange(rowIndex, COL.STATUS).setValue(payload.status);
  return jsonResp({ status: 'ok' });
}

function doGet(e) {
  try {
    const action = e.parameter.action || '';
    const pass = e.parameter.pass || '';
    const q = e.parameter.q || '';

    if (action === 'getOrders') {
      if (pass !== ADMIN_PASS) return jsonResp({ status: 'error', message: 'Unauthorized' });
      return getOrders();
    }
    if (action === 'checkStatus') {
      return checkStatus(q);
    }
    return jsonResp({ status: 'error', message: 'Unknown action' });
  } catch (err) {
    return jsonResp({ status: 'error', message: err.message });
  }
}

function getOrders() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'ok', rows: [] });

  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    rows.push({
      rowIndex: i + 1,
      timestamp: r[COL.TIMESTAMP - 1] ? r[COL.TIMESTAMP - 1].toString() : '',
      orderId: r[COL.ORDER_ID - 1],
      buyerName: r[COL.BUYER_NAME - 1],
      phone: r[COL.PHONE - 1],
      lineId: r[COL.LINE_ID - 1],
      delivery: r[COL.DELIVERY - 1],
      address: r[COL.ADDRESS - 1],
      grandTotal: r[COL.GRAND_TOTAL - 1],
      entryIndex: r[COL.ENTRY_INDEX - 1],
      forWhom: r[COL.FOR_WHOM - 1],
      team: r[COL.TEAM - 1],
      division: r[COL.DIVISION - 1],
      color: r[COL.COLOR - 1],
      size: r[COL.SIZE - 1],
      price: r[COL.PRICE - 1],
      status: r[COL.STATUS - 1],
    });
  }
  return jsonResp({ status: 'ok', rows });
}

function checkStatus(q) {
  if (!q) return jsonResp({ status: 'ok', rows: [] });
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'ok', rows: [] });

  const data = sheet.getDataRange().getValues();
  const ql = q.toLowerCase();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    const name = String(r[COL.BUYER_NAME - 1] || '').toLowerCase();
    const phone = String(r[COL.PHONE - 1] || '');
    if (name.includes(ql) || phone.includes(q)) {
      rows.push({
        rowIndex: i + 1,
        orderId: r[COL.ORDER_ID - 1],
        buyerName: r[COL.BUYER_NAME - 1],
        phone: r[COL.PHONE - 1],
        delivery: r[COL.DELIVERY - 1],
        address: r[COL.ADDRESS - 1],
        grandTotal: r[COL.GRAND_TOTAL - 1],
        entryIndex: r[COL.ENTRY_INDEX - 1],
        forWhom: r[COL.FOR_WHOM - 1],
        team: r[COL.TEAM - 1],
        division: r[COL.DIVISION - 1],
        color: r[COL.COLOR - 1],
        size: r[COL.SIZE - 1],
        status: r[COL.STATUS - 1],
      });
    }
  }
  return jsonResp({ status: 'ok', rows });
}

function getOrCreateFolder(name) {
  const folders = DriveApp.getFoldersByName(name);
  return folders.hasNext() ? folders.next() : DriveApp.createFolder(name);
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}
