/**
 * Golf Travel Mobility — 預約／訂單查詢／修改取消申請 後端
 * 使用方式：貼到 Google 試算表「擴充功能 → Apps Script」，取代原本全部程式碼。
 * 部署：部署 → 管理部署作業 → 編輯（鉛筆）→ 版本選「新版本」→ 部署（網址不變）。
 * 第一次使用請在編輯器上方選 setupSheet 並按「執行」一次，建立欄位與下拉選單。
 */

const SHEET_BOOKING = '預約';
const SHEET_CHANGE = '修改取消申請';
const TZ = 'Asia/Taipei';
const SITE_URL = 'https://golf-travel-mobility-gylb.vercel.app/';
const SEND_CONFIRM_EMAIL = true; // 不想寄確認信改成 false

const BOOKING_HEADERS = [
  '預約編號', '送出時間', '姓名', '手機', 'Email', '通知方式', '出發地', '目的地',
  '出發日期', '開球日期', '球袋數量', '單程/來回', '備註',
  '訂單狀態', '付款狀態', '預計收件日', '預計送達日', '物流公司', '物流追蹤碼', '客服備註'
];
const CHANGE_HEADERS = ['送出時間', '預約編號', 'Email 或手機', '申請項目', '說明', '處理狀態', '客服備註'];

const STATUS_LIST = [
  '預約已送出', '專人確認中', '待補資料', '報價已提供', '待付款', '已付款', '已安排收件',
  '已完成收件', '運送中', '清關中', '已抵達目的地', '球場／飯店已簽收', '異常處理中', '已完成服務', '已取消'
];
const PAY_LIST = ['未付款', '待付款', '已付款', '已退款', '部分退款'];
const DELIVERED_STATUS = ['球場／飯店已簽收', '已完成服務'];

/* ---------- 進入點 ---------- */

function doPost(e) {
  let data = {};
  try {
    data = JSON.parse((e && e.postData && e.postData.contents) || '{}');
  } catch (err) {
    return json_({ ok: false, error: 'bad_request' });
  }
  const type = data.type || 'booking';
  try {
    if (type === 'lookup') return json_(lookup_(data));
    if (type === 'change') return json_(saveChange_(data));
    return json_(saveBooking_(data));
  } catch (err) {
    return json_({ ok: false, error: 'server_error' });
  }
}

function doGet() {
  return json_({ ok: true, service: 'Golf Travel Mobility' });
}

/* ---------- 預約 ---------- */

function saveBooking_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(SHEET_BOOKING);
    const map = ensureHeaders_(sheet, BOOKING_HEADERS);
    const orderId = cleanId_(d.orderId) || genId_();
    const row = {
      '預約編號': orderId,
      '送出時間': Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'),
      '姓名': d.name,
      '手機': d.phone ? "'" + d.phone : '',
      'Email': d.email,
      '通知方式': d.notify,
      '出發地': d.origin,
      '目的地': d.destination,
      '出發日期': d.departDate,
      '開球日期': d.teeDate,
      '球袋數量': d.bagCount,
      '單程/來回': d.trip,
      '備註': d.note,
      '訂單狀態': '預約已送出',
      '付款狀態': '未付款'
    };
    appendByHeader_(sheet, map, row);
    if (SEND_CONFIRM_EMAIL && isEmail_(d.email)) sendConfirm_(d, orderId);
    return { ok: true, orderId: orderId };
  } finally {
    lock.releaseLock();
  }
}

function sendConfirm_(d, orderId) {
  const link = SITE_URL + '?order=' + encodeURIComponent(orderId) + '#track';
  const body = [
    (d.name || '您好') + ' 您好：',
    '',
    '已收到您的球袋運送預約需求。',
    '預約編號：' + orderId,
    '目的地：' + (d.destination || '') + '　出發日期：' + (d.departDate || ''),
    '',
    '提醒：送出預約需求不代表訂單成立。',
    '客服將確認配送條件並提供正式報價（有效 48 小時），完成付款後訂單才正式成立，並開始安排收件與配送。',
    '',
    '查詢訂單進度：' + link,
    '（請以預約編號＋預約時填寫的 Email 或手機查詢）',
    '',
    '客服服務時間：週一至週五 9:00–18:00',
    '非上班時間收到的預約，將於下一個工作時間處理。',
    'LINE 客服：https://lin.ee/XOMWpls',
    '',
    'Golf Travel Mobility'
  ].join('\n');
  try {
    MailApp.sendEmail({ to: d.email, subject: '【Golf Travel Mobility】已收到您的預約需求（' + orderId + '）', body: body });
  } catch (err) { /* 寄信失敗不影響預約寫入 */ }
}

/* ---------- 訂單查詢 ---------- */

function lookup_(d) {
  const orderId = cleanId_(d.order);
  const contact = String(d.contact || '').trim();
  if (!orderId || !contact) return { ok: false, error: 'not_found' };

  const sheet = getSheet_(SHEET_BOOKING);
  const map = ensureHeaders_(sheet, BOOKING_HEADERS);
  const last = sheet.getLastRow();
  if (last < 2) return { ok: false, error: 'not_found' };
  const values = sheet.getRange(2, 1, last - 1, sheet.getLastColumn()).getValues();
  const get = (r, h) => (map[h] ? r[map[h] - 1] : '');

  for (const r of values) {
    if (cleanId_(get(r, '預約編號')) !== orderId) continue;
    if (!contactMatches_(contact, get(r, 'Email'), get(r, '手機'))) continue;
    const status = String(get(r, '訂單狀態') || '預約已送出');
    return {
      ok: true,
      order: {
        orderId: orderId,
        name: maskName_(get(r, '姓名')),
        status: status,
        payStatus: fmt_(get(r, '付款狀態')),
        pickupDate: fmt_(get(r, '預計收件日')),
        deliveryDate: fmt_(get(r, '預計送達日')),
        carrier: fmt_(get(r, '物流公司')),
        tracking: fmt_(get(r, '物流追蹤碼')),
        delivered: DELIVERED_STATUS.indexOf(status) >= 0
      }
    };
  }
  return { ok: false, error: 'not_found' };
}

function contactMatches_(input, email, phone) {
  const i = String(input).trim().toLowerCase();
  if (i.indexOf('@') >= 0) return i === String(email || '').trim().toLowerCase();
  const a = i.replace(/\D/g, ''), b = String(phone || '').replace(/\D/g, '');
  if (a.length < 8 || b.length < 8) return false;
  return a.slice(-9) === b.slice(-9);
}

/* ---------- 修改／取消申請 ---------- */

function saveChange_(d) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    const sheet = getSheet_(SHEET_CHANGE);
    const map = ensureHeaders_(sheet, CHANGE_HEADERS);
    appendByHeader_(sheet, map, {
      '送出時間': Utilities.formatDate(new Date(), TZ, 'yyyy/MM/dd HH:mm'),
      '預約編號': cleanId_(d.order),
      'Email 或手機': d.contact ? "'" + d.contact : '',
      '申請項目': d.item,
      '說明': d.detail,
      '處理狀態': '待處理'
    });
    return { ok: true };
  } finally {
    lock.releaseLock();
  }
}

/* ---------- 一次性設定（手動執行） ---------- */

function setupSheet() {
  const b = getSheet_(SHEET_BOOKING);
  const bm = ensureHeaders_(b, BOOKING_HEADERS);
  b.setFrozenRows(1);
  setDropdown_(b, bm['訂單狀態'], STATUS_LIST);
  setDropdown_(b, bm['付款狀態'], PAY_LIST);

  const c = getSheet_(SHEET_CHANGE);
  const cm = ensureHeaders_(c, CHANGE_HEADERS);
  c.setFrozenRows(1);
  setDropdown_(c, cm['處理狀態'], ['待處理', '處理中', '已完成', '無法處理']);
}

/* ---------- 工具 ---------- */

function getSheet_(name) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(name) || ss.insertSheet(name);
}

// 依標題對應欄位；缺少的標題補在最右邊，不動既有資料
function ensureHeaders_(sheet, headers) {
  const lastCol = sheet.getLastColumn();
  let current = lastCol ? sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(String) : [];
  const missing = headers.filter(h => current.indexOf(h) < 0);
  if (missing.length) {
    sheet.getRange(1, current.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold');
    current = current.concat(missing);
  }
  const map = {};
  current.forEach((h, i) => { if (h) map[h] = i + 1; });
  return map;
}

function appendByHeader_(sheet, map, obj) {
  const row = new Array(sheet.getLastColumn()).fill('');
  Object.keys(obj).forEach(k => { if (map[k]) row[map[k] - 1] = obj[k] == null ? '' : obj[k]; });
  sheet.appendRow(row);
}

function setDropdown_(sheet, col, list) {
  if (!col) return;
  const rule = SpreadsheetApp.newDataValidation().requireValueInList(list, true).setAllowInvalid(true).build();
  sheet.getRange(2, col, Math.max(sheet.getMaxRows() - 1, 1), 1).setDataValidation(rule);
}

function genId_() {
  const chars = 'ABCDEFGHJKLMNPQRSTUVWXYZ23456789';
  let r = '';
  for (let i = 0; i < 4; i++) r += chars[Math.floor(Math.random() * chars.length)];
  return 'GTM' + Utilities.formatDate(new Date(), TZ, 'yyMMdd') + '-' + r;
}

function cleanId_(v) { return String(v || '').trim().toUpperCase().replace(/[^A-Z0-9-]/g, ''); }
function isEmail_(v) { return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(String(v || '')); }
function maskName_(v) { const s = String(v || '').trim(); return s ? s.charAt(0) + '○○' : ''; }
function fmt_(v) {
  if (v instanceof Date) return Utilities.formatDate(v, TZ, 'yyyy/MM/dd');
  return v == null ? '' : String(v);
}
function json_(o) {
  return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON);
}
