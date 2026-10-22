/**
 * ระบบหลังบ้านเว็บรับสมัครงาน — เก็บข้อมูลลง Google Sheets
 * 1 แถว = ผู้สมัคร 1 คน · 1 คอลัมน์ = ข้อมูล 1 ช่อง (ใบสมัครทุกช่อง + ผลแบบทดสอบ + ส่วน HR)
 * แถวที่ 1 = หัวคอลัมน์ภาษาไทย · แถวที่ 2 (ซ่อนไว้) = รหัสช่องที่ระบบใช้ อย่าแก้หรือลบ
 * วิธีติดตั้งดูใน README.md หัวข้อ "เชื่อม Google Sheets"
 */

// ⚠️ ตั้งรหัสผ่านสำหรับ HR ก่อน Deploy (ใช้ดู / แก้ / ลบข้อมูลผู้สมัคร)
const HR_PASSWORD = 'เปลี่ยนรหัสผ่านตรงนี้';

const SHEET_NAME = 'ผู้สมัคร';
const FIRST = 3;          // แถวแรกของข้อมูลผู้สมัคร
const MAX_CELL = 49000;   // 1 ช่องเก็บได้ไม่เกิน 50,000 ตัวอักษร

// หน้าเว็บใบสมัคร (ผู้สมัครเปิดลิงก์นี้)  ·  ใส่ ?hr=1 ท้ายลิงก์เพื่อเปิดหน้า HR
function doGet(e) {
  let html = HtmlService.createHtmlOutputFromFile('index').getContent();
  if (e && e.parameter && e.parameter.hr) html = html.replace('<body>', '<body data-hr="1">');
  return HtmlService.createHtmlOutput(html)
    .setTitle('ใบสมัครงาน - บริษัท อินฟินิตี้ บลู จำกัด')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// ใช้เมื่อโฮสต์หน้าเว็บที่อื่น (เช่น GitHub Pages) แล้วส่งข้อมูลมาที่นี่
function doPost(e) {
  const allowed = { apiSubmit, apiList, apiSaveHr, apiRemove };
  let out;
  try {
    const req = JSON.parse(e.postData.contents);
    if (!allowed[req.fn]) throw new Error('ไม่รู้จักคำสั่ง');
    out = { result: allowed[req.fn](req.payload) };
  } catch (err) {
    out = { error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

function sheet_() {
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(SHEET_NAME);
  if (!sh) sh = ss.insertSheet(SHEET_NAME);
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 2, 2).setValues([['รหัสผู้สมัคร', 'วันเวลาที่สมัคร'], ['id', 'createdAt']]);
    sh.setFrozenRows(2);
    sh.hideRows(2);
    sh.getRange('1:1').setFontWeight('bold').setBackground('#e8efff');
  }
  return sh;
}
// รหัสช่องของทุกคอลัมน์ (แถว 2)
function keys_(sh) {
  return sh.getRange(2, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
}
// เพิ่มคอลัมน์ใหม่ให้รหัสช่องที่ยังไม่มี แล้วคืนรายการรหัสทั้งหมด
function ensureCols_(sh, labels, wanted) {
  const keys = keys_(sh);
  const add = wanted.filter(k => k && keys.indexOf(k) < 0);
  if (add.length) {
    const c = keys.length + 1;
    sh.getRange(1, c, 2, add.length).setValues([add.map(k => String(labels[k] || k)), add]);
    sh.getRange(1, c, 1, add.length).setFontWeight('bold').setBackground('#e8efff');
    return keys.concat(add);
  }
  return keys;
}
// กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ Sheets ตีความเป็นสูตร
const safe_ = v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : (v == null ? '' : v);
function checkKey_(key) {
  if (HR_PASSWORD === 'เปลี่ยนรหัสผ่านตรงนี้') throw new Error('ยังไม่ได้ตั้งรหัสผ่าน HR ในไฟล์ Code.gs');
  if (String(key || '') !== HR_PASSWORD) throw new Error('รหัสผ่าน HR ไม่ถูกต้อง');
}
function findRow_(sh, id) {
  const n = sh.getLastRow() - FIRST + 1;
  if (n > 0) {
    const ids = sh.getRange(FIRST, 1, n, 1).getValues();
    for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + FIRST;
  }
  throw new Error('ไม่พบผู้สมัคร');
}
// เขียนค่าลงแถว: ข้อความเก็บเป็นข้อความ (เบอร์โทรไม่หายเลข 0) ตัวเลขผลทดสอบเก็บเป็นตัวเลข
function writeRow_(sh, row, keys, rec) {
  const vals = keys.map(k => (k in rec) ? safe_(rec[k]) : '');
  const fmts = vals.map(v => typeof v === 'number' ? '0' : '@');
  const r = sh.getRange(row, 1, 1, keys.length);
  r.setNumberFormats([fmts]);
  r.setValues([vals]);
}
const lock_ = fn => { const l = LockService.getScriptLock(); l.waitLock(20000); try { return fn(); } finally { l.releaseLock(); } };

// p = { row: {รหัสช่อง: ค่า}, labels: {รหัสช่อง: หัวคอลัมน์ภาษาไทย} }
function apiSubmit(p) {
  const rec = (p && p.row) || {};
  if (!rec.id || !rec.nameTh) throw new Error('ข้อมูลไม่ครบ');
  Object.keys(rec).forEach(k => {
    const v = rec[k];
    if (v != null && typeof v === 'object') delete rec[k];
    else if (typeof v === 'string' && v.length > MAX_CELL) rec[k] = k === 'signImg' ? '' : v.slice(0, MAX_CELL);
  });
  if (Object.keys(rec).length > 400) throw new Error('ข้อมูลไม่ถูกต้อง');
  return lock_(() => {
    const sh = sheet_();
    const keys = ensureCols_(sh, p.labels || {}, Object.keys(rec));
    writeRow_(sh, Math.max(sh.getLastRow() + 1, FIRST), keys, rec);
    return 'ok';
  });
}

// คืนรายการผู้สมัครเป็น {รหัสช่อง: ค่า}
function apiList(p) {
  checkKey_(p && p.key);
  const sh = sheet_(), keys = keys_(sh), n = sh.getLastRow() - FIRST + 1;
  if (n <= 0) return [];
  return sh.getRange(FIRST, 1, n, keys.length).getDisplayValues().map(r => {
    const o = {};
    keys.forEach((k, i) => { if (k && r[i] !== '') o[k] = r[i]; });
    return o;
  }).filter(o => o.id);
}

// p = { key, id, hr: {status, pos, ...}, labels }
function apiSaveHr(p) {
  checkKey_(p && p.key);
  return lock_(() => {
    const sh = sheet_(), row = findRow_(sh, p.id), hr = p.hr || {};
    const rec = {}; Object.keys(hr).forEach(k => rec['hr.' + k] = String(hr[k] == null ? '' : hr[k]));
    const keys = ensureCols_(sh, p.labels || {}, Object.keys(rec));
    keys.forEach((k, i) => { if (k in rec) sh.getRange(row, i + 1).setNumberFormat('@').setValue(safe_(rec[k])); });
    return 'ok';
  });
}

function apiRemove(p) {
  checkKey_(p && p.key);
  return lock_(() => { const sh = sheet_(); sh.deleteRow(findRow_(sh, p.id)); return 'ok'; });
}
