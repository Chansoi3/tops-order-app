/**
 * AM404 สั่งสินค้า (โปรโมชั่น & RMS) - BACKEND API
 *
 * Deploy: Deploy > New deployment > Web app
 *   - Execute as: Me
 *   - Who has access: Anyone
 * รองรับทั้ง GET (?action=...) และ POST (body = JSON)
 */

const SHEET_ID = '10ikBHPlkfHpWtCKy3YBo77YCqzQUXOLHsT8VaPJB8Mk';

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

      case 'saveOrder': {
        const payload = parsePayload(e);
        const branch = e.parameter.branch || payload.branch;
        const msg = saveToOrdersSheet(branch, payload.orders || []);
        return jsonResponse({ success: true, message: msg });
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

// ---------- 3. บันทึกคำสั่งซื้อ (ล็อกกันชนกัน + เขียนทีเดียว) ----------
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
    return 'บันทึกข้อมูลจำนวน ' + rows.length + ' รายการ เข้าฐานข้อมูลเรียบร้อยแล้ว!';
  } finally {
    lock.releaseLock();
  }
}

function toNumber(value, fallback) {
  const n = Number(value);
  return isNaN(n) ? fallback : n;
}
