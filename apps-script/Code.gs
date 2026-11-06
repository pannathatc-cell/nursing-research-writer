/**
 * ผู้ช่วยเขียนงานวิจัยทางการพยาบาล — ตัวเชื่อม Google Drive / Google Sheets
 * พัฒนาโดย Pannathat Chanpanit, Sisaket
 *
 * หน้าที่
 *  - บันทึกไฟล์สำรองโครงการ (.json) ลงโฟลเดอร์ Drive  → โฟลเดอร์ย่อย "Projects"
 *  - บันทึกไฟล์ Word (.docx) ที่สร้างจากโปรแกรม           → โฟลเดอร์ย่อย "Word"
 *  - บันทึกทะเบียนโครงการ คลังอ้างอิง และผลตรวจอ้างอิง APA 7 ลง Google Sheets
 *  - ดึงรายการโครงการและไฟล์สำรองกลับมาเปิดในโปรแกรม
 *
 * วิธีติดตั้ง (ทำครั้งเดียว)
 *  1) เปิด Google Sheets ตาม SHEET_ID ด้านล่าง > ส่วนขยาย (Extensions) > Apps Script
 *  2) ลบโค้ดเดิมในไฟล์ Code.gs แล้ววางโค้ดนี้ทั้งหมด กดบันทึก
 *  3) เลือกฟังก์ชัน setup แล้วกด Run (เรียกใช้) > อนุญาตสิทธิ์ (Review permissions) ด้วยบัญชีของคุณ
 *  4) Deploy (ทำให้ใช้งานได้) > New deployment > ประเภท Web app
 *       Execute as: Me (ตัวคุณเอง) | Who has access: Anyone (ทุกคน)
 *     กด Deploy แล้วคัดลอก Web app URL (ลงท้ายด้วย /exec)
 *  5) ในโปรแกรม เมนู "บันทึกลง Google Drive" วาง Web app URL และ Token ด้านล่าง แล้วกดทดสอบการเชื่อมต่อ
 *  เมื่อแก้โค้ดภายหลัง: Deploy > Manage deployments > แก้ไข (ดินสอ) > Version: New version > Deploy (URL เดิมใช้ต่อได้)
 */

const FOLDER_ID = '1KDoh6OXY0-ua9b8Y7YfHpEmiKNVJkAEC';
const SHEET_ID  = '1Z65dzySp-5yF9HSYQ29kp3CJACCgk3h6UDWGvns4ZOM';
/** รหัสลับที่โปรแกรมต้องส่งมาด้วย ป้องกันคนอื่นที่รู้ URL เขียนข้อมูลลง Drive ของคุณ เปลี่ยนได้ แต่ต้องแก้ในโปรแกรมให้ตรงกัน */
const TOKEN     = 'ตั้งรหัสลับของคุณเอง';
const TZ        = 'Asia/Bangkok';

const SHEETS = {
  Projects:   ['รหัสโครงการ','ชื่อเรื่อง (ไทย)','ชื่อเรื่อง (อังกฤษ)','หัวข้อ/ประเด็น','ผู้วิจัย','รูปแบบอ้างอิง','หัวข้อที่เขียนแล้ว','จำนวนคำ','เอกสารในคลังอ้างอิง','บันทึกล่าสุด','ไฟล์สำรอง (ลิงก์)','รหัสไฟล์สำรอง','บันทึกโดย'],
  References: ['รหัสโครงการ','ชื่อโครงการ','ลำดับ','รายการอ้างอิง (APA 7)','ผู้แต่ง','ปี','ชื่อเรื่อง','วารสาร/แหล่ง','DOI','URL','ยืนยันแล้ว','แหล่งที่มา','อัปเดต'],
  RefAudit:   ['รหัสโครงการ','ชื่อโครงการ','การอ้างในเนื้อหา','รายการอ้างอิง (APA 7)','สถานะ','หัวข้อที่อ้าง','ลิงก์ตรวจสอบ','ตรวจเมื่อ'],
  Markers:    ['รหัสโครงการ','ชื่อโครงการ','หัวข้อ','ป้าย','ข้อความ','ตรวจเมื่อ'],
  Exports:    ['รหัสโครงการ','ชื่อโครงการ','ชื่อไฟล์ Word','ลิงก์','รหัสไฟล์','ขนาด (KB)','สร้างเมื่อ'],
  Log:        ['เวลา','คำสั่ง','รหัสโครงการ','ผล','รายละเอียด']
};

/* ---------- ติดตั้ง ---------- */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(SHEETS).forEach(name => ensureSheet_(ss, name));
  const root = DriveApp.getFolderById(FOLDER_ID);
  sub_(root, 'Projects'); sub_(root, 'Word');
  Logger.log('พร้อมใช้งาน: โฟลเดอร์ "' + root.getName() + '" และสเปรดชีต "' + ss.getName() + '"');
}

/* ---------- จุดรับคำสั่ง ---------- */
function doGet(e) {
  const p = (e && e.parameter) || {};
  if (!p.action) return out_({ ok: true, app: 'nursing-research-writer', msg: 'ตัวเชื่อม Google Drive ทำงานอยู่ ใช้งานผ่านโปรแกรม' });
  return handle_(p);
}
function doPost(e) {
  let body = {};
  try { body = JSON.parse((e && e.postData && e.postData.contents) || '{}'); }
  catch (err) { return out_({ ok: false, error: 'ข้อมูลที่ส่งมาไม่ใช่ JSON' }); }
  return handle_(body);
}

function handle_(b) {
  if (b.token !== TOKEN) return out_({ ok: false, error: 'Token ไม่ถูกต้อง ตรวจให้ตรงกับค่า TOKEN ใน Apps Script' });
  const lock = LockService.getScriptLock();
  try {
    lock.waitLock(20000);
    const fn = ACTIONS[b.action];
    if (!fn) return out_({ ok: false, error: 'ไม่รู้จักคำสั่ง ' + b.action });
    const res = fn(b);
    log_(b.action, b.pid || (b.project && b.project.id) || '', 'สำเร็จ', res && res.msg || '');
    return out_(Object.assign({ ok: true }, res));
  } catch (err) {
    log_(b.action, b.pid || '', 'ผิดพลาด', String(err && err.message || err));
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

const ACTIONS = {
  ping(b) {
    const ss = SpreadsheetApp.openById(SHEET_ID), root = DriveApp.getFolderById(FOLDER_ID);
    Object.keys(SHEETS).forEach(name => ensureSheet_(ss, name));
    return { folder: root.getName(), folderUrl: root.getUrl(), sheet: ss.getName(), sheetUrl: ss.getUrl(), msg: 'เชื่อมต่อแล้ว' };
  },

  /** บันทึกไฟล์สำรองโครงการ (ทับไฟล์เดิมของโครงการเดียวกัน) และอัปเดตทะเบียนโครงการ */
  saveProject(b) {
    const p = b.project || {}; if (!p.id) throw new Error('ไม่มีรหัสโครงการ');
    const folder = sub_(DriveApp.getFolderById(FOLDER_ID), 'Projects');
    const name = safe_(p.title_th || p.topic || 'โครงการ') + ' [' + p.id + '].json';
    let file = findByPid_(folder, p.id);
    if (file) { file.setContent(b.backup); file.setName(name); }
    else { file = folder.createFile(name, b.backup, 'application/json'); file.setDescription('nrw:' + p.id); }
    const m = b.metrics || {};
    upsert_('Projects', p.id, [p.id, p.title_th || '', p.title_en || '', p.topic || '', p.researcher || '', b.citeStyle || '',
      (m.filled || 0) + '/' + (m.total || 0), m.words || 0, m.refs || 0, now_(), file.getUrl(), file.getId(), b.by || '']);
    return { fileId: file.getId(), url: file.getUrl(), msg: name };
  },

  /** แทนที่คลังอ้างอิงของโครงการใน Sheet References */
  saveRefs(b) {
    const rows = (b.refs || []).map((r, i) => [b.pid, b.title || '', i + 1, r.apa || '', r.authors || '', r.year || '', r.title || '', r.journal || '', r.doi || '', r.url || '', r.verified ? 'ใช่' : 'ยังไม่ยืนยัน', r.source || '', now_()]);
    replaceRows_('References', b.pid, rows);
    return { count: rows.length, msg: rows.length + ' รายการ' };
  },

  /** แทนที่ผลตรวจการอ้างอิง (APA 7) และจุดที่ต้องหาแหล่ง ของโครงการ */
  saveAudit(b) {
    const t = now_();
    const a = (b.items || []).map(x => [b.pid, b.title || '', x.cite || '', x.apa || '', x.status || '', x.where || '', x.link || '', t]);
    const m = (b.markers || []).map(x => [b.pid, b.title || '', x.section || '', x.kind || '', x.text || '', t]);
    replaceRows_('RefAudit', b.pid, a); replaceRows_('Markers', b.pid, m);
    return { count: a.length, markers: m.length, msg: a.length + ' รายการ / ' + m.length + ' จุด' };
  },

  /** บันทึกไฟล์ Word ลงโฟลเดอร์ Word */
  saveDocx(b) {
    const folder = sub_(DriveApp.getFolderById(FOLDER_ID), 'Word');
    const bytes = Utilities.base64Decode(b.base64);
    const stamp = Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HHmm');
    const name = safe_(String(b.filename || 'รายงาน.docx').replace(/\.docx$/i, '')) + ' ' + stamp + '.docx';
    const file = folder.createFile(Utilities.newBlob(bytes, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document', name));
    sheet_('Exports').appendRow([b.pid || '', b.title || '', name, file.getUrl(), file.getId(), Math.round(bytes.length / 1024), now_()]);
    return { fileId: file.getId(), url: file.getUrl(), msg: name };
  },

  /** รายการโครงการที่สำรองไว้ใน Drive */
  listProjects() {
    const sh = sheet_('Projects'); const n = sh.getLastRow();
    const rows = n > 1 ? sh.getRange(2, 1, n - 1, SHEETS.Projects.length).getDisplayValues() : [];
    return { projects: rows.map(r => ({ id: r[0], title: r[1] || r[3] || '(ไม่มีชื่อ)', researcher: r[4], progress: r[6], updated: r[9], url: r[10], fileId: r[11] })).reverse() };
  },

  /** ดึงไฟล์สำรองโครงการกลับมา */
  loadProject(b) {
    let file = null;
    if (b.fileId) file = DriveApp.getFileById(b.fileId);
    else if (b.pid) file = findByPid_(sub_(DriveApp.getFolderById(FOLDER_ID), 'Projects'), b.pid);
    if (!file) throw new Error('ไม่พบไฟล์สำรองของโครงการนี้');
    return { backup: file.getBlob().getDataAsString('UTF-8'), name: file.getName(), msg: file.getName() };
  }
};

/* ---------- ตัวช่วย ---------- */
function out_(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }
function now_() { return Utilities.formatDate(new Date(), TZ, 'yyyy-MM-dd HH:mm:ss'); }
function safe_(s) { return String(s).replace(/[\\/:*?"<>|#\[\]]/g, '').replace(/\s+/g, ' ').trim().slice(0, 90) || 'ไม่มีชื่อ'; }
function sub_(parent, name) { const it = parent.getFoldersByName(name); return it.hasNext() ? it.next() : parent.createFolder(name); }
function findByPid_(folder, pid) {
  const it = folder.searchFiles('title contains "[' + pid + ']" and trashed = false');
  return it.hasNext() ? it.next() : null;
}
function ensureSheet_(ss, name) {
  let sh = ss.getSheetByName(name);
  if (!sh) sh = ss.insertSheet(name);
  if (sh.getLastRow() === 0) {
    sh.appendRow(SHEETS[name]);
    sh.getRange(1, 1, 1, SHEETS[name].length).setFontWeight('bold').setBackground('#DCEBEA');
    sh.setFrozenRows(1);
  }
  return sh;
}
function sheet_(name) { return ensureSheet_(SpreadsheetApp.openById(SHEET_ID), name); }
function upsert_(name, pid, row) {
  const sh = sheet_(name); const n = sh.getLastRow();
  if (n > 1) {
    const ids = sh.getRange(2, 1, n - 1, 1).getValues().map(r => String(r[0]));
    const i = ids.indexOf(String(pid));
    if (i >= 0) { sh.getRange(i + 2, 1, 1, row.length).setValues([row]); return; }
  }
  sh.appendRow(row);
}
function replaceRows_(name, pid, rows) {
  const sh = sheet_(name); const n = sh.getLastRow(), w = SHEETS[name].length;
  const keep = n > 1 ? sh.getRange(2, 1, n - 1, w).getValues().filter(r => String(r[0]) !== String(pid)) : [];
  const all = keep.concat(rows);
  if (n > 1) sh.getRange(2, 1, n - 1, w).clearContent();
  if (all.length) sh.getRange(2, 1, all.length, w).setValues(all);
}
function log_(action, pid, result, detail) {
  try { sheet_('Log').appendRow([now_(), action, pid, result, String(detail).slice(0, 300)]); } catch (e) {}
}
