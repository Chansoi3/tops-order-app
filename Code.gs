/**
 * AM404 สั่งสินค้า (โปรโมชั่น & RMS) - BACKEND API
 *
 * Deploy: Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * รองรับทั้ง GET (?action=...) และ POST (body = JSON)
 */

const SHEET_ID = '10ikBHPlkfHpWtCKy3YBo77YCqzQUXOLHsT8VaPJB8Mk';

// อีเมลที่จะรับสรุปคำสั่งซื้อทุกครั้งที่มีการบันทึก
const NOTIFY_EMAIL = 'Piwirakorn@tops.co.th';

const SHEET = { PRODUCTS: 'Products', RMS: 'RMS', ORDERS: 'Orders' };

const ORDERS_HEADER = [
  'วันที่', 'รหัสสาขา', 'สาขา', 'บาร์โค้ด', 'ชื่อสินค้า', 'บรรจุภัณฑ์ต่อ DU',
  'จำนวน(DU)', 'รวมชิ้น', 'ราคาปกติ', 'ราคาโปรโมชั่น', 'Promotion Type', 'Promotion'
];

// ---------- Routing ----------
function doGet(e) { return handleRequest(e); }
function doPost(e) { return handleRequest(e); }

function handleRequest(e) {
  if (!e || !e.parameter || !e.parameter.action) {
    return textResponse('API AM404 System Ready');
  }

  const action = e.parameter.action;
  try {
    switch (action) {
      case 'getProducts':
        return jsonResponse(getAvailableProductsForBranch(e.parameter.branch));

      case 'getRMSProducts':
        return jsonResponse(getRMSProducts());

      case 'getOrders':
        return jsonResponse(getOrders(e.parameter.branch));

      case 'saveOrder': {
        const payload = parsePayload(e);
        const branch = e.parameter.branch || payload.branch;
        const msg = saveToOrdersSheet(branch, payload.orders || []);
        return jsonResponse({ success: true, message: msg });
      }

      case 'deleteOrder': {
        const payload = parsePayload(e);
        const count = deleteOrders(payload.branchCode, payload.ts, payload.barcode);
        return jsonResponse({ success: true, deleted: count, message: 'ยกเลิกคำสั่งซื้อ ' + count + ' รายการแล้ว' });
      }

      default:
        return jsonResponse({ success: false, error: 'Unknown action: ' + action });
    }
  } catch (err) {
    return jsonResponse({ success: false, error: err.toString() });
  }
}

function parsePayload(e) {
  const raw = e.parameter.payload || (e.postData && e.postData.contents) || '{}';
  return JSON.parse(raw);
}

function jsonResponse(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

function textResponse(str) {
  return ContentService.createTextOutput(str).setMimeType(ContentService.MimeType.TEXT);
}

// ---------- 1. สินค้าโปรโมชั่น (ซ่อนรายการที่สาขานี้สั่งไปแล้ว) ----------
function getAvailableProductsForBranch(branch) {
  const ss = SpreadsheetApp.openById(SHEET_ID);

  const productSheet = ss.getSheetByName(SHEET.PRODUCTS);
  if (!productSheet) throw new Error('ไม่พบชีต "' + SHEET.PRODUCTS + '"');

  const branchCode = String(branch).split('-')[0].trim();
  const orderedBarcodes = getOrderedBarcodes(ss, branchCode);

  const prodData = productSheet.getDataRange().getValues();
  const items = [];

  for (let i = 1; i < prodData.length; i++) {
    const row = prodData[i];
    const barcode = String(row[3]).trim();
    if (!barcode || orderedBarcodes.has(barcode)) continue;

    items.push({
      startDate: row[0],
      endDate: row[1],
      group: row[2],
      barcode: barcode,
      name: row[4],
      promoType: String(row[5]).trim(),
      promotionDesc: row[6] || '-',
      normalPrice: row[7],
      promoPrice: row[8],
      duSize: (row[9] !== '' && row[9] != null) ? row[9] : 1
    });
  }
  return items;
}

function getOrderedBarcodes(ss, branchCode) {
  const orderSheet = ss.getSheetByName(SHEET.ORDERS);
  const ordered = new Set();
  if (!orderSheet || orderSheet.getLastRow() < 2) return ordered;

  const data = orderSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][1]).trim() === branchCode) {
      ordered.add(String(data[i][3]).trim());
    }
  }
  return ordered;
}

// ---------- 2. สินค้า RMS (บาร์โค้ด 13 หลัก) ----------
function getRMSProducts() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const rmsSheet = ss.getSheetByName(SHEET.RMS);
  if (!rmsSheet) return [];

  const data = rmsSheet.getDataRange().getValues();
  const items = [];

  // D(3)=รหัสสินค้า 13 หลัก, F(5)=ชื่อสินค้า, H(7)=แพค
  for (let i = 1; i < data.length; i++) {
    const barcode = String(data[i][3]).trim();
    if (!barcode || barcode === 'รหัสสินค้า2') continue;

    items.push({
      barcode: barcode,
      name: String(data[i][5]).trim(),
      duSize: (data[i][7] !== '' && data[i][7] != null) ? data[i][7] : 1
    });
  }
  return items;
}

// ---------- 3. ดึงประวัติคำสั่งซื้อที่บันทึกแล้ว ----------
function getOrders(branchFilter) {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  const sheet = ss.getSheetByName(SHEET.ORDERS);
  if (!sheet || sheet.getLastRow() < 2) return [];

  const data = sheet.getDataRange().getValues();
  const code = branchFilter ? String(branchFilter).split('-')[0].trim() : '';
  const tz = Session.getScriptTimeZone();
  const out = [];

  for (let i = 1; i < data.length; i++) {
    const r = data[i];
    if (code && String(r[1]).trim() !== code) continue;

    const rawDate = r[0];
    const isDate = Object.prototype.toString.call(rawDate) === '[object Date]';

    out.push({
      ts: isDate ? rawDate.getTime() : 0,
      date: isDate ? Utilities.formatDate(rawDate, tz, 'dd/MM/yyyy HH:mm') : String(rawDate),
      branchCode: String(r[1]).trim(),
      branchName: String(r[2]).trim(),
      barcode: String(r[3]).trim(),
      name: r[4],
      duSize: r[5],
      duQty: r[6],
      totalQty: r[7],
      normalPrice: r[8],
      promoPrice: r[9],
      promoType: String(r[10]).trim(),
      promotion: r[11]
    });
  }
  return out;
}

// ---------- 4. บันทึกคำสั่งซื้อ (ล็อกกันชนกัน + เขียนทีเดียว) ----------
function saveToOrdersSheet(branch, orders) {
  if (!orders || orders.length === 0) throw new Error('ไม่มีรายการสั่งซื้อ');

  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    let sheet = ss.getSheetByName(SHEET.ORDERS);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET.ORDERS);
      sheet.appendRow(ORDERS_HEADER);
    }

    const timestamp = new Date();
    const parts = String(branch).split('-');
    const branchCode = parts[0].trim();
    const branchName = parts.slice(1).join('-').trim();

    const rows = orders.map(function (item) {
      const p = item.product || {};
      return [
        timestamp,
        branchCode,
        branchName,
        p.barcode,
        p.name,
        toNumber(p.duSize, 1),
        toNumber(item.duQty, 0),
        toNumber(item.totalQty, 0),
        (p.normalPrice != null && p.normalPrice !== '') ? p.normalPrice : '-',
        (p.promoPrice != null && p.promoPrice !== '') ? p.promoPrice : '-',
        p.promoType || 'ราคาปกติ',
        p.promotionDesc || 'RMS'
      ];
    });

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, ORDERS_HEADER.length).setValues(rows);

    // ส่งอีเมลสรุป (ไม่ให้ล้มเหลวกระทบการบันทึก)
    try {
      sendOrderEmail(branchCode, branchName, timestamp, rows);
    } catch (mailErr) {
      Logger.log('ส่งอีเมลไม่สำเร็จ: ' + mailErr);
    }

    return 'บันทึกข้อมูลจำนวน ' + rows.length + ' รายการ เข้าฐานข้อมูลเรียบร้อยแล้ว!';
  } finally {
    lock.releaseLock();
  }
}

// ---------- 5. ยกเลิก/ลบคำสั่งซื้อ ----------
// ลบทั้งบิล: ส่ง branchCode + ts | ลบรายเดียว: ส่ง branchCode + ts + barcode
function deleteOrders(branchCode, ts, barcode) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = SpreadsheetApp.openById(SHEET_ID);
    const sheet = ss.getSheetByName(SHEET.ORDERS);
    if (!sheet || sheet.getLastRow() < 2) return 0;

    const code = String(branchCode).trim();
    const targetTs = Number(ts);
    const data = sheet.getDataRange().getValues();
    const rowsToDelete = [];

    for (let i = 1; i < data.length; i++) {
      const r = data[i];
      const rowTs = (Object.prototype.toString.call(r[0]) === '[object Date]') ? r[0].getTime() : 0;
      if (String(r[1]).trim() !== code) continue;
      if (rowTs !== targetTs) continue;
      if (barcode && String(r[3]).trim() !== String(barcode).trim()) continue;
      rowsToDelete.push(i + 1); // 1-based row number
    }

    // ลบจากล่างขึ้นบนเพื่อไม่ให้ index เลื่อน
    for (let j = rowsToDelete.length - 1; j >= 0; j--) {
      sheet.deleteRow(rowsToDelete[j]);
    }
    return rowsToDelete.length;
  } finally {
    lock.releaseLock();
  }
}

// ---------- ส่งอีเมลสรุปคำสั่งซื้อ ----------
function sendOrderEmail(branchCode, branchName, timestamp, rows) {
  if (!NOTIFY_EMAIL) return;

  const tz = Session.getScriptTimeZone();
  const dateStr = Utilities.formatDate(timestamp, tz, 'dd/MM/yyyy HH:mm');
  let totalPieces = 0;

  let tableRows = '';
  rows.forEach(function (r, i) {
    totalPieces += Number(r[7]) || 0;
    tableRows +=
      '<tr>' +
      '<td style="padding:6px 8px;border:1px solid #e7eaf0;text-align:center;">' + (i + 1) + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #e7eaf0;">' + r[4] + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #e7eaf0;text-align:center;">' + r[3] + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #e7eaf0;text-align:center;">' + r[6] + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #e7eaf0;text-align:center;">' + r[7] + '</td>' +
      '<td style="padding:6px 8px;border:1px solid #e7eaf0;text-align:center;">' + r[10] + '</td>' +
      '</tr>';
  });

  const html =
    '<div style="font-family:Arial,sans-serif;color:#1f2a44;max-width:700px;">' +
      '<div style="background:#00205b;color:#fff;padding:16px 20px;border-radius:10px 10px 0 0;">' +
        '<h2 style="margin:0;font-size:18px;">คำสั่งซื้อใหม่ - AM404</h2>' +
        '<div style="font-size:13px;opacity:.85;">ระบบสั่งสินค้า AM404</div>' +
      '</div>' +
      '<div style="border:1px solid #e7eaf0;border-top:0;padding:16px 20px;border-radius:0 0 10px 10px;">' +
        '<p style="margin:4px 0;"><b>สาขา:</b> ' + branchCode + ' - ' + branchName + '</p>' +
        '<p style="margin:4px 0;"><b>วันที่/เวลา:</b> ' + dateStr + '</p>' +
        '<p style="margin:4px 0;"><b>จำนวนรายการ:</b> ' + rows.length + ' • <b>รวม:</b> ' + totalPieces + ' ชิ้น</p>' +
        '<table style="border-collapse:collapse;width:100%;margin-top:12px;font-size:13px;">' +
          '<thead><tr style="background:#f3f5fa;">' +
            '<th style="padding:6px 8px;border:1px solid #e7eaf0;">#</th>' +
            '<th style="padding:6px 8px;border:1px solid #e7eaf0;text-align:left;">ชื่อสินค้า</th>' +
            '<th style="padding:6px 8px;border:1px solid #e7eaf0;">บาร์โค้ด</th>' +
            '<th style="padding:6px 8px;border:1px solid #e7eaf0;">DU</th>' +
            '<th style="padding:6px 8px;border:1px solid #e7eaf0;">รวมชิ้น</th>' +
            '<th style="padding:6px 8px;border:1px solid #e7eaf0;">ประเภท</th>' +
          '</tr></thead>' +
          '<tbody>' + tableRows + '</tbody>' +
        '</table>' +
      '</div>' +
    '</div>';

  MailApp.sendEmail({
    to: NOTIFY_EMAIL,
    subject: '[AM404] คำสั่งซื้อใหม่ - ' + branchCode + ' ' + branchName + ' (' + rows.length + ' รายการ)',
    htmlBody: html
  });
}

function toNumber(value, fallback) {
  const n = Number(value);
  return isNaN(n) ? fallback : n;
}
