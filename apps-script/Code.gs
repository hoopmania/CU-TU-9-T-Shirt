// CU-TU Basketball Alumni #9 — Apps Script
// Each order stores one row per entry (per shirt per person)

const SHEET_NAME = 'Orders';
const ADMIN_PASS = 'cutuba9';

const COL = {
  TIMESTAMP: 1, ORDER_ID: 2, BUYER_NAME: 3, PHONE: 4, LINE_ID: 5,
  DELIVERY: 6, ADDRESS: 7, GRAND_TOTAL: 8, ENTRY_INDEX: 9,
  FOR_WHOM: 10, TEAM: 11, DIVISION: 12, COLOR: 13, SIZE: 14,
  PRICE: 15, STATUS: 16, NOTE: 17,
};

function doPost(e) {
  try {
    const payload = JSON.parse(e.postData.contents);
    const action = payload.action || 'order';
    if (action === 'updateStatus') return handleUpdateStatus(payload);
    if (action === 'getOrders') return handleGetOrders(payload);
    if (action === 'checkStatus') return handleCheckStatus(payload);
    return handleOrder(payload);
  } catch (err) {
    return jsonResp({ status: 'error', message: String(err.message) });
  }
}

// Keep doGet for direct browser testing only
function doGet(e) {
  return jsonResp({ status: 'ok', message: 'Use POST requests' });
}

function handleOrder(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    const headers = ['วันที่','เลขออร์เดอร์','ชื่อผู้สั่ง','โทร','LINE ID','วิธีรับ','ที่อยู่','ยอดรวม','ตัวที่','สำหรับ','ทีม','รุ่น','สีเสื้อ','ไซส์','ราคา','สถานะ','หมายเหตุ'];
    sheet.getRange(1, 1, 1, headers.length).setValues([headers])
      .setFontWeight('bold').setBackground('#0a1f5c').setFontColor('#ffffff');
    sheet.setFrozenRows(1);
  }

  const ts = new Date();
  const orderId = payload.orderId || ('CU' + ts.getTime().toString().slice(-8));
  const entries = payload.entries || [];

  // Expand entries: each entry may have navy and/or white
  const rows = [];
  entries.forEach((entry, i) => {
    const colors = [];
    if (entry.navyOn && entry.navySize) colors.push({ color: 'กรม', size: entry.navySize });
    if (entry.whiteOn && entry.whiteSize) colors.push({ color: 'ขาว', size: entry.whiteSize });
    colors.forEach(c => {
      const row = new Array(17).fill('');
      row[COL.TIMESTAMP - 1] = ts;
      row[COL.ORDER_ID - 1] = orderId;
      row[COL.BUYER_NAME - 1] = String(payload.buyerName || '');
      row[COL.PHONE - 1] = String(payload.phone || '');
      row[COL.LINE_ID - 1] = String(payload.lineId || '');
      row[COL.DELIVERY - 1] = String(payload.delivery || '');
      row[COL.ADDRESS - 1] = String(payload.address || '');
      row[COL.GRAND_TOTAL - 1] = rows.length === 0 ? (payload.grandTotal || 0) : '';
      row[COL.ENTRY_INDEX - 1] = i + 1;
      row[COL.FOR_WHOM - 1] = String(entry.forWhom || '');
      row[COL.TEAM - 1] = String(entry.team || '');
      row[COL.DIVISION - 1] = String(entry.division || '');
      row[COL.COLOR - 1] = c.color;
      row[COL.SIZE - 1] = c.size;
      row[COL.PRICE - 1] = 300;
      row[COL.STATUS - 1] = payload.delivery === 'ship' ? 'ส่งที่บ้าน' : 'ยังไม่รับ';
      row[COL.NOTE - 1] = rows.length === 0 ? String(payload.note || '') : '';
      rows.push(row);
    });
  });

  if (rows.length > 0) {
    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, 17).setValues(rows);
  }

  return jsonResp({ status: 'ok', orderId });
}

function handleUpdateStatus(payload) {
  if (payload.pass !== ADMIN_PASS) return jsonResp({ status: 'error', message: 'Unauthorized' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'error', message: 'Sheet not found' });
  const rowIndex = parseInt(payload.rowIndex);
  if (isNaN(rowIndex) || rowIndex < 2) return jsonResp({ status: 'error', message: 'Invalid row' });
  sheet.getRange(rowIndex, COL.STATUS).setValue(String(payload.status || ''));
  return jsonResp({ status: 'ok' });
}

function handleGetOrders(payload) {
  if (payload.pass !== ADMIN_PASS) return jsonResp({ status: 'error', message: 'Unauthorized' });
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return jsonResp({ status: 'ok', rows: [] });
  const data = sheet.getDataRange().getValues();
  const rows = [];
  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    rows.push({
      rowIndex: i + 1,
      timestamp: r[COL.TIMESTAMP-1] ? r[COL.TIMESTAMP-1].toString() : '',
      orderId: String(r[COL.ORDER_ID-1] || ''),
      buyerName: String(r[COL.BUYER_NAME-1] || ''),
      phone: String(r[COL.PHONE-1] || ''),
      lineId: String(r[COL.LINE_ID-1] || ''),
      delivery: String(r[COL.DELIVERY-1] || ''),
      address: String(r[COL.ADDRESS-1] || ''),
      grandTotal: r[COL.GRAND_TOTAL-1] || 0,
      entryIndex: r[COL.ENTRY_INDEX-1] || '',
      forWhom: String(r[COL.FOR_WHOM-1] || ''),
      team: String(r[COL.TEAM-1] || ''),
      division: String(r[COL.DIVISION-1] || ''),
      color: String(r[COL.COLOR-1] || ''),
      size: String(r[COL.SIZE-1] || ''),
      price: r[COL.PRICE-1] || 0,
      status: String(r[COL.STATUS-1] || ''),
    });
  }
  return jsonResp({ status: 'ok', rows });
}

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
    const name = String(r[COL.BUYER_NAME-1] || '').toLowerCase();
    const phone = String(r[COL.PHONE-1] || '');
    if (name.includes(ql) || phone.includes(q)) {
      rows.push({
        rowIndex: i + 1,
        orderId: String(r[COL.ORDER_ID-1] || ''),
        buyerName: String(r[COL.BUYER_NAME-1] || ''),
        phone: String(r[COL.PHONE-1] || ''),
        delivery: String(r[COL.DELIVERY-1] || ''),
        address: String(r[COL.ADDRESS-1] || ''),
        grandTotal: r[COL.GRAND_TOTAL-1] || 0,
        entryIndex: r[COL.ENTRY_INDEX-1] || '',
        forWhom: String(r[COL.FOR_WHOM-1] || ''),
        team: String(r[COL.TEAM-1] || ''),
        division: String(r[COL.DIVISION-1] || ''),
        color: String(r[COL.COLOR-1] || ''),
        size: String(r[COL.SIZE-1] || ''),
        status: String(r[COL.STATUS-1] || ''),
      });
    }
  }
  return jsonResp({ status: 'ok', rows });
}

function jsonResp(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
