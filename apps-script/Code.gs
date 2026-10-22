/**
 * ระบบหลังบ้านเว็บรับสมัครงาน — เก็บข้อมูลลง Google Sheets
 * วิธีติดตั้งดูใน README.md หัวข้อ "เชื่อม Google Sheets"
 */

// ⚠️ ตั้งรหัสผ่านสำหรับ HR ก่อน Deploy (ใช้ดู / แก้ / ลบข้อมูลผู้สมัคร)
const HR_PASSWORD = 'เปลี่ยนรหัสผ่านตรงนี้';

const SHEET_NAME = 'ผู้สมัคร';
const HEAD = ['id', 'วันที่สมัคร', 'ชื่อ-นามสกุล', 'ชื่อเล่น', 'ตำแหน่ง 1', 'ตำแหน่ง 2', 'โทร', 'อีเมล', 'วันเกิด',
  'เหมาะกับตำแหน่ง 1 (%)', 'ตำแหน่งที่เหมาะสุด', 'DISC', 'BOSI', 'Enneagram', 'เลขเส้นทางชีวิต', 'เข้ากับ CEO',
  'สถานะ HR', 'hr', 'sig', 'data1', 'data2', 'data3', 'data4'];
const COL = name => HEAD.indexOf(name) + 1;
const CHUNK = 45000; // 1 ช่องเก็บได้ไม่เกิน 50,000 ตัวอักษร

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
    sh.appendRow(HEAD);
    sh.setFrozenRows(1);
    sh.getRange(1, 1, 1, HEAD.length).setFontWeight('bold');
  }
  return sh;
}
// กันข้อความที่ขึ้นต้นด้วย = + - @ ไม่ให้ Sheets ตีความเป็นสูตร
const safe_ = v => (typeof v === 'string' && /^[=+\-@]/.test(v)) ? "'" + v : (v == null ? '' : v);
function checkKey_(key) {
  if (HR_PASSWORD === 'เปลี่ยนรหัสผ่านตรงนี้') throw new Error('ยังไม่ได้ตั้งรหัสผ่าน HR ในไฟล์ Code.gs');
  if (String(key || '') !== HR_PASSWORD) throw new Error('รหัสผ่าน HR ไม่ถูกต้อง');
}
function findRow_(sh, id) {
  const ids = sh.getRange(2, 1, Math.max(sh.getLastRow() - 1, 1), 1).getValues();
  for (let i = 0; i < ids.length; i++) if (String(ids[i][0]) === String(id)) return i + 2;
  throw new Error('ไม่พบผู้สมัคร');
}

function apiSubmit(a) {
  if (!a || typeof a !== 'object' || !a.id || !a.nameTh) throw new Error('ข้อมูลไม่ครบ');
  const sig = /^data:image\/png;base64,/.test(a.signImg || '') && a.signImg.length < CHUNK ? a.signImg : '';
  const data = Object.assign({}, a); delete data.signImg; delete data.hr;
  const json = JSON.stringify(data);
  if (json.length > CHUNK * 4) throw new Error('ข้อมูลยาวเกินไป');
  const parts = [0, 1, 2, 3].map(i => json.slice(i * CHUNK, (i + 1) * CHUNK));
  const s = a.summary || {};
  const lock = LockService.getScriptLock(); lock.waitLock(20000);
  try {
    const sh = sheet_();
    const row = [a.id, Utilities.formatDate(new Date(), 'Asia/Bangkok', 'yyyy-MM-dd HH:mm'), a.nameTh, a.nick, a.pos1, a.pos2, a.phone, a.email, a.dob,
      s.fit1, s.best, s.disc, s.bosi, s.enn, s.lp, s.ceo, 'รอพิจารณา', '{}', sig].map(safe_).concat(parts);
    sh.getRange(sh.getLastRow() + 1, 1, 1, row.length).setNumberFormat('@').setValues([row]);
  } finally { lock.releaseLock(); }
  return 'ok';
}

function apiList(p) {
  checkKey_(p && p.key);
  const sh = sheet_();
  if (sh.getLastRow() < 2) return [];
  const rows = sh.getRange(2, 1, sh.getLastRow() - 1, HEAD.length).getValues();
  const out = [];
  rows.forEach(r => {
    try {
      const a = JSON.parse(r.slice(COL('data1') - 1).join(''));
      a.signImg = r[COL('sig') - 1] || '';
      a.hr = JSON.parse(r[COL('hr') - 1] || '{}');
      out.push(a);
    } catch (e) { /* ข้ามแถวที่ข้อมูลเสีย */ }
  });
  return out;
}

function apiSaveHr(p) {
  checkKey_(p && p.key);
  const sh = sheet_(), row = findRow_(sh, p.id), hr = p.hr || {};
  sh.getRange(row, COL('hr')).setValue(JSON.stringify(hr));
  sh.getRange(row, COL('สถานะ HR')).setValue(safe_(hr.status || 'รอพิจารณา'));
  return 'ok';
}

function apiRemove(p) {
  checkKey_(p && p.key);
  const sh = sheet_();
  sh.deleteRow(findRow_(sh, p.id));
  return 'ok';
}
