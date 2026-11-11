/**
 * ผู้ช่วยเขียนงานวิจัยทางการพยาบาล — ตัวเชื่อม Google Drive / Google Sheets
 * พัฒนาโดย Pannathat Chanpanit, Sisaket
 *
 * หน้าที่
 *  - บันทึกไฟล์สำรองโครงการ (.json) ลงโฟลเดอร์ Drive  → โฟลเดอร์ย่อย "Projects"
 *  - บันทึกไฟล์ Word (.docx) ที่สร้างจากโปรแกรม           → โฟลเดอร์ย่อย "Word"
 *  - บันทึกทะเบียนโครงการ คลังอ้างอิง และผลตรวจอ้างอิง APA 7 ลง Google Sheets
 *  - ดึงรายการโครงการและไฟล์สำรองกลับมาเปิดในโปรแกรม
 *  - ระบบขอใช้งานและอนุมัติ: ผู้ใช้กรอกชื่อ อีเมล หน่วยงาน → แท็บ Users สถานะ "รออนุมัติ"
 *    แอดมินเปลี่ยนช่อง "สถานะ" ในแท็บ Users เป็น "อนุมัติ" หรือ "ระงับ" (หนึ่งแถวต่อหนึ่งอีเมลต่อหนึ่งเครื่อง)
 *  - บันทึกการใช้งานทุกครั้งที่เปิดเว็บและเปลี่ยนเมนู → แท็บ Usage และสรุปที่แท็บ Dashboard
 *  - ตั้งข้อความติดต่อผู้ดูแลและเปิด/ปิดการบังคับขออนุมัติได้ที่แท็บ Settings
 *  - บัญชีผู้ใช้และโควตา token: แอดมินสร้างบัญชี (ชื่อผู้ใช้ + รหัสผ่าน + จำนวน token) จากเมนู "จัดการผู้ใช้และโควตา token" ในโปรแกรม
 *    → แท็บ Accounts | ทุกครั้งที่ผู้ใช้เรียก AI จะหัก token และบันทึกที่แท็บ TokenLog เมื่อ token หมดจะใช้ AI ไม่ได้จนกว่าแอดมินเติม
 *    แก้ในแท็บ Accounts ได้โดยตรง: โควตา token, สถานะ (ใช้งานได้/ระงับ), วันหมดอายุ (รหัสผ่านต้องตั้งจากโปรแกรมเท่านั้น)
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
const TOKEN     = 'ใส่-token-ของคุณ';
const TZ        = 'Asia/Bangkok';
/** หลักฐานรหัสผ่านแอดมิน (ค่าแฮช ไม่ใช่รหัสผ่านจริง) ห้ามเผยแพร่ค่านี้ */
const ADMIN_PROOF = '__ADMIN_PROOF__';

const SHEETS = {
  Projects:   ['รหัสโครงการ','ชื่อเรื่อง (ไทย)','ชื่อเรื่อง (อังกฤษ)','หัวข้อ/ประเด็น','ผู้วิจัย','รูปแบบอ้างอิง','หัวข้อที่เขียนแล้ว','จำนวนคำ','เอกสารในคลังอ้างอิง','บันทึกล่าสุด','ไฟล์สำรอง (ลิงก์)','รหัสไฟล์สำรอง','บันทึกโดย'],
  References: ['รหัสโครงการ','ชื่อโครงการ','ลำดับ','รายการอ้างอิง (APA 7)','ผู้แต่ง','ปี','ชื่อเรื่อง','วารสาร/แหล่ง','DOI','URL','ยืนยันแล้ว','แหล่งที่มา','อัปเดต'],
  RefAudit:   ['รหัสโครงการ','ชื่อโครงการ','การอ้างในเนื้อหา','รายการอ้างอิง (APA 7)','สถานะ','หัวข้อที่อ้าง','ลิงก์ตรวจสอบ','ตรวจเมื่อ'],
  Markers:    ['รหัสโครงการ','ชื่อโครงการ','หัวข้อ','ป้าย','ข้อความ','ตรวจเมื่อ'],
  Exports:    ['รหัสโครงการ','ชื่อโครงการ','ชื่อไฟล์ Word','ลิงก์','รหัสไฟล์','ขนาด (KB)','สร้างเมื่อ'],
  Log:        ['เวลา','คำสั่ง','รหัสโครงการ','ผล','รายละเอียด'],
  Users:      ['อีเมล','ชื่อ-นามสกุล','หน่วยงาน','สถานะ','วันที่ขอใช้งาน','ใช้งานล่าสุด','จำนวนครั้งที่เปิด','เครื่อง','รหัสเครื่อง','หมายเหตุจากแอดมิน'],
  Usage:      ['เวลา','อีเมล','ชื่อ','เมนู','เครื่อง','รหัสเครื่อง','รหัสการใช้งาน'],
  Settings:   ['ค่า','ข้อมูล','คำอธิบาย'],
  Accounts:   ['ชื่อผู้ใช้','ชื่อ-นามสกุล','สถานะ','โควตา token','ใช้ไปแล้ว (token)','คงเหลือ (token)','วันหมดอายุ','แพ็กเกจ/จำนวนเงิน','หมายเหตุ','สร้างเมื่อ','เข้าใช้ล่าสุด','จำนวนครั้งที่เรียก AI','รหัสผ่าน (แฮช)','รหัสเซสชัน'],
  TokenLog:   ['เวลา','ชื่อผู้ใช้','ชื่อ','เมนู','token ขาเข้า','token ขาออก','รวม','คงเหลือ','โมเดล','วิธีนับ']
};
const AC = { U: 0, NAME: 1, ST: 2, QUOTA: 3, USED: 4, LEFT: 5, EXP: 6, PACK: 7, NOTE: 8, CREATED: 9, LAST: 10, CALLS: 11, HASH: 12, SID: 13 };
const AST = { OK: 'ใช้งานได้', BAN: 'ระงับ' };
const ST = { OK: 'อนุมัติ', WAIT: 'รออนุมัติ', BAN: 'ระงับ' };
const SETTING_DEFAULTS = [
  ['contact', 'กรุณาติดต่อผู้ดูแลระบบ (admin)', 'ข้อความ/ช่องทางติดต่อที่แสดงกับผู้ใช้ที่รออนุมัติหรือถูกระงับ เช่น อีเมล หรือ LINE ID'],
  ['require_approval', 'ใช่', 'ใช่ = ต้องได้รับอนุมัติก่อนใช้งาน | ไม่ = ใช้งานได้ทันทีหลังลงทะเบียน (ยังบันทึกสถิติ)'],
  ['message_wait', 'บัญชีของคุณรอการอนุมัติ', 'หัวข้อที่แสดงเมื่อรออนุมัติ'],
  ['message_ban', 'บัญชีของคุณถูกระงับการใช้งาน', 'หัวข้อที่แสดงเมื่อถูกระงับ']
];

/* ---------- ติดตั้ง ---------- */
function setup() {
  const ss = SpreadsheetApp.openById(SHEET_ID);
  Object.keys(SHEETS).forEach(name => ensureSheet_(ss, name));
  const root = DriveApp.getFolderById(FOLDER_ID);
  sub_(root, 'Projects'); sub_(root, 'Word');
  setupAccess_(ss);
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
    if (ADMIN_ACTIONS.indexOf(b.action) >= 0 && !isAdmin_(b)) return out_({ ok: false, error: 'คำสั่งนี้สำหรับผู้ดูแลระบบเท่านั้น (รหัสผ่านแอดมินไม่ถูกต้อง)' });
    if (DATA_ACTIONS.indexOf(b.action) >= 0 && !isAdmin_(b) && !session_(b)) {
      return out_({ ok: false, error: 'ต้องเข้าสู่ระบบก่อนบันทึกข้อมูล' });
    }
    const res = fn(b);
    if (b.action === 'access' || b.action === 'meter' || b.action === 'acct') return out_(Object.assign({ ok: true }, res));
    log_(b.action, b.pid || (b.project && b.project.id) || '', 'สำเร็จ', res && res.msg || '');
    return out_(Object.assign({ ok: true }, res));
  } catch (err) {
    log_(b.action, b.pid || '', 'ผิดพลาด', String(err && err.message || err));
    return out_({ ok: false, error: String(err && err.message || err) });
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
}

const DATA_ACTIONS = ['saveProject', 'saveRefs', 'saveAudit', 'saveDocx'];
const ADMIN_ACTIONS = ['listProjects', 'loadProject', 'adminList', 'adminSave', 'adminLog'];
const ACTIONS = {
  /** ผู้ใช้เข้าสู่ระบบ: {u, ph} → ออกรหัสเซสชันใหม่ (เข้าจากเครื่องใหม่ เครื่องเดิมจะถูกออกจากระบบ) */
  login(b) {
    const f = acctFind_(b.u);
    if (!f || !b.ph || String(f.row[AC.HASH]) !== String(b.ph)) throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
    const chk = acctCheck_(f.row); if (chk) throw new Error(chk);
    const sid = Utilities.getUuid();
    f.sh.getRange(f.r, AC.LAST + 1).setValue(new Date());
    f.sh.getRange(f.r, AC.SID + 1).setValue(sid);
    return Object.assign({ sid: sid, msg: 'เข้าสู่ระบบ ' + f.row[AC.U] }, acctInfo_(f.row));
  },
  /** ตรวจเซสชันและยอด token คงเหลือ */
  acct(b) {
    const f = session_(b); if (!f) return { valid: false, error: 'เซสชันหมดอายุหรือมีการเข้าสู่ระบบจากเครื่องอื่น' };
    const chk = acctCheck_(f.row); if (chk) return { valid: false, error: chk };
    return Object.assign({ valid: true }, acctInfo_(f.row));
  },
  /** หัก token หลังเรียก AI: {u, sid, tin, tout, view, model, how} */
  meter(b) {
    const f = session_(b); if (!f) return { valid: false, error: 'เซสชันหมดอายุหรือมีการเข้าสู่ระบบจากเครื่องอื่น' };
    const tin = Math.max(0, Math.round(Number(b.tin) || 0)), tout = Math.max(0, Math.round(Number(b.tout) || 0));
    const used = (Number(f.row[AC.USED]) || 0) + tin + tout, quota = Number(f.row[AC.QUOTA]) || 0, calls = (Number(f.row[AC.CALLS]) || 0) + 1;
    f.sh.getRange(f.r, AC.USED + 1, 1, 2).setValues([[used, quota - used]]);
    f.sh.getRange(f.r, AC.CALLS + 1).setValue(calls);
    sheet_('TokenLog').appendRow([new Date(), f.row[AC.U], f.row[AC.NAME], String(b.view || '').slice(0, 60), tin, tout, tin + tout, quota - used, String(b.model || '').slice(0, 60), b.how === 'est' ? 'ประมาณจากจำนวนตัวอักษร' : 'จากผู้ให้บริการ']);
    f.row[AC.USED] = used; f.row[AC.CALLS] = calls;
    const chk = acctCheck_(f.row);
    return Object.assign({ valid: !chk, error: chk || '' }, acctInfo_(f.row));
  },
  /** แอดมิน: รายชื่อบัญชีทั้งหมด (ไม่ส่งรหัสผ่าน) */
  adminList() {
    const sh = sheet_('Accounts'); const n = sh.getLastRow();
    const rows = n > 1 ? sh.getRange(2, 1, n - 1, SHEETS.Accounts.length).getValues() : [];
    const d = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd') : String(v || '');
    const dt = v => v instanceof Date ? Utilities.formatDate(v, TZ, 'yyyy-MM-dd HH:mm') : String(v || '');
    return { accounts: rows.filter(r => String(r[AC.U]).trim()).map(r => ({ u: String(r[AC.U]), name: String(r[AC.NAME]), status: String(r[AC.ST] || AST.OK),
      quota: Number(r[AC.QUOTA]) || 0, used: Number(r[AC.USED]) || 0, exp: d(r[AC.EXP]), pack: String(r[AC.PACK] || ''), note: String(r[AC.NOTE] || ''),
      created: dt(r[AC.CREATED]), last: dt(r[AC.LAST]), calls: Number(r[AC.CALLS]) || 0, hasPw: !!String(r[AC.HASH] || '') })) };
  },
  /** แอดมิน: สร้างหรือแก้ไขบัญชี {acct:{u,name,status,quota,addQuota,resetUsed,exp,pack,note,ph,kick}} */
  adminSave(b) {
    const a = b.acct || {}; const u = String(a.u || '').trim().toLowerCase();
    if (!/^[a-z0-9._-]{3,40}$/.test(u)) throw new Error('ชื่อผู้ใช้ต้องเป็น a-z 0-9 . _ - ยาว 3–40 ตัว');
    let f = acctFind_(u); const sh = sheet_('Accounts');
    if (!f) {
      if (!a.ph) throw new Error('บัญชีใหม่ต้องตั้งรหัสผ่าน');
      sh.appendRow([u, '', AST.OK, 0, 0, 0, '', '', '', new Date(), '', 0, '', '']);
      acctRule_(sh); f = acctFind_(u);
    }
    const r = f.row;
    if (a.name !== undefined) r[AC.NAME] = String(a.name);
    if (a.status) r[AC.ST] = a.status === AST.BAN ? AST.BAN : AST.OK;
    if (a.quota !== undefined && a.quota !== '') r[AC.QUOTA] = Math.max(0, Number(a.quota) || 0);
    if (a.addQuota) r[AC.QUOTA] = (Number(r[AC.QUOTA]) || 0) + Number(a.addQuota);
    if (a.resetUsed) { r[AC.USED] = 0; r[AC.CALLS] = 0; }
    if (a.exp !== undefined) r[AC.EXP] = a.exp ? new Date(a.exp + 'T23:59:59+07:00') : '';
    if (a.pack !== undefined) r[AC.PACK] = String(a.pack);
    if (a.note !== undefined) r[AC.NOTE] = String(a.note);
    if (a.ph) { r[AC.HASH] = String(a.ph); r[AC.SID] = ''; }
    if (a.kick || r[AC.ST] === AST.BAN) r[AC.SID] = '';
    r[AC.LEFT] = (Number(r[AC.QUOTA]) || 0) - (Number(r[AC.USED]) || 0);
    sh.getRange(f.r, 1, 1, SHEETS.Accounts.length).setValues([r]);
    return { msg: 'บันทึกบัญชี ' + u };
  },
  /** แอดมิน: ประวัติการใช้ token ล่าสุด {u?, limit} */
  adminLog(b) {
    const sh = sheet_('TokenLog'); const n = sh.getLastRow(); const lim = Math.min(Number(b.limit) || 200, 1000);
    if (n < 2) return { log: [] };
    const start = Math.max(2, n - 4999); let rows = sh.getRange(start, 1, n - start + 1, 10).getValues();
    if (b.u) rows = rows.filter(r => String(r[1]) === String(b.u));
    return { log: rows.slice(-lim).reverse().map(r => ({ t: r[0] instanceof Date ? Utilities.formatDate(r[0], TZ, 'yyyy-MM-dd HH:mm') : String(r[0]), u: String(r[1]), view: String(r[3]), tin: Number(r[4]) || 0, tout: Number(r[5]) || 0, left: Number(r[7]) || 0, model: String(r[8]) })) };
  },

  /** ลงทะเบียน ตรวจสิทธิ์ และบันทึกการใช้งาน: {who:{email,name,org,device,deviceId}, view, session, register} */
  access(b) {
    const w = b.who || {}; const email = String(w.email || '').trim().toLowerCase(); const dev = String(w.deviceId || '').trim();
    if (!email || !dev) return { status: 'ไม่พบ', contact: setting_('contact') };
    const sh = sheet_('Users'); const n = sh.getLastRow(); const W = SHEETS.Users.length;
    const rows = n > 1 ? sh.getRange(2, 1, n - 1, W).getValues() : [];
    let i = rows.findIndex(r => String(r[0]).trim().toLowerCase() === email && String(r[8]) === dev);
    const now = new Date();
    if (i < 0) {
      if (!b.register) return { status: 'ไม่พบ', contact: setting_('contact') };
      const st = setting_('require_approval') === 'ไม่' ? ST.OK : ST.WAIT;
      sh.appendRow([email, w.name || '', w.org || '', st, now, now, 0, w.device || '', dev, '']);
      i = rows.length; rows.push([email, w.name || '', w.org || '', st, now, now, 0, w.device || '', dev, '']);
      statusRule_(sh);
    } else if (b.register) {
      sh.getRange(i + 2, 2, 1, 2).setValues([[w.name || rows[i][1], w.org || rows[i][2]]]);
    }
    const r = rows[i]; const status = String(r[3] || ST.WAIT).trim();
    const count = (Number(r[6]) || 0) + (b.view === 'open' ? 1 : 0);
    sh.getRange(i + 2, 6, 1, 3).setValues([[now, count, w.device || r[7]]]);
    if (b.view) {
      const u = sheet_('Usage'); u.appendRow([now, email, w.name || r[1], String(b.view).slice(0, 60), w.device || '', dev, b.session || '']);
    }
    const S = { status: status, contact: setting_('contact'), note: String(r[9] || '') };
    S.title = status === ST.BAN ? setting_('message_ban') : status === ST.OK ? '' : setting_('message_wait');
    return S;
  },

  ping(b) {
    const ss = SpreadsheetApp.openById(SHEET_ID), root = DriveApp.getFolderById(FOLDER_ID);
    Object.keys(SHEETS).forEach(name => ensureSheet_(ss, name));
    if (!ss.getSheetByName('Dashboard')) setupAccess_(ss);
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

/* ---------- บัญชีผู้ใช้และโควตา token ---------- */
function isAdmin_(b) { return ADMIN_PROOF && ADMIN_PROOF.indexOf('__') !== 0 && String(b.admin || '') === ADMIN_PROOF; }
function acctFind_(u) {
  u = String(u || '').trim().toLowerCase(); if (!u) return null;
  const sh = sheet_('Accounts'); const n = sh.getLastRow(); if (n < 2) return null;
  const rows = sh.getRange(2, 1, n - 1, SHEETS.Accounts.length).getValues();
  const i = rows.findIndex(r => String(r[AC.U]).trim().toLowerCase() === u);
  return i < 0 ? null : { sh: sh, r: i + 2, row: rows[i] };
}
function session_(b) {
  if (!b || !b.u || !b.sid) return null;
  const f = acctFind_(b.u); return f && String(f.row[AC.SID]) === String(b.sid) ? f : null;
}
function acctCheck_(r) {
  if (String(r[AC.ST]).trim() === AST.BAN) return 'บัญชีนี้ถูกระงับการใช้งาน กรุณาติดต่อผู้ดูแลระบบ';
  const e = r[AC.EXP]; if (e && new Date(e).getTime() < Date.now()) return 'บัญชีนี้หมดอายุแล้ว กรุณาติดต่อผู้ดูแลระบบ';
  return '';
}
function acctInfo_(r) {
  const q = Number(r[AC.QUOTA]) || 0, u = Number(r[AC.USED]) || 0, e = r[AC.EXP];
  return { u: String(r[AC.U]), name: String(r[AC.NAME] || r[AC.U]), quota: q, used: u, left: q - u,
    exp: e instanceof Date ? Utilities.formatDate(e, TZ, 'yyyy-MM-dd') : String(e || ''), status: String(r[AC.ST] || AST.OK) };
}
function acctRule_(sh) {
  const n = Math.max(sh.getMaxRows() - 1, 1); const rng = sh.getRange(2, AC.ST + 1, n, 1);
  rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([AST.OK, AST.BAN], true).setAllowInvalid(false).build());
  sh.getRange(2, AC.QUOTA + 1, n, 3).setNumberFormat('#,##0');
  sh.getRange(2, AC.EXP + 1, n, 1).setNumberFormat('yyyy-mm-dd');
  sh.getRange(2, AC.CREATED + 1, n, 2).setNumberFormat('yyyy-mm-dd hh:mm');
}

/* ---------- ระบบขอใช้งาน ---------- */
function deviceStatus_(who) {
  const w = who || {}; const email = String(w.email || '').trim().toLowerCase(); const dev = String(w.deviceId || '');
  if (!email || !dev) return 'ไม่พบ';
  const sh = sheet_('Users'); const n = sh.getLastRow(); if (n < 2) return 'ไม่พบ';
  const r = sh.getRange(2, 1, n - 1, SHEETS.Users.length).getValues().find(x => String(x[0]).trim().toLowerCase() === email && String(x[8]) === dev);
  return r ? String(r[3]).trim() : 'ไม่พบ';
}
function setting_(key) {
  const sh = sheet_('Settings'); const n = sh.getLastRow();
  const rows = n > 1 ? sh.getRange(2, 1, n - 1, 2).getValues() : [];
  const f = rows.find(r => String(r[0]).trim() === key); if (f) return String(f[1]).trim();
  const d = SETTING_DEFAULTS.find(x => x[0] === key); return d ? d[1] : '';
}
function statusRule_(sh) {
  const n = Math.max(sh.getMaxRows() - 1, 1); const rng = sh.getRange(2, 4, n, 1);
  rng.setDataValidation(SpreadsheetApp.newDataValidation().requireValueInList([ST.WAIT, ST.OK, ST.BAN], true).setAllowInvalid(false).build());
  if (sh.getConditionalFormatRules().length) return;
  const mk = (txt, bg, fg) => SpreadsheetApp.newConditionalFormatRule().whenTextEqualTo(txt).setBackground(bg).setFontColor(fg).setRanges([rng]).build();
  sh.setConditionalFormatRules([mk(ST.OK, '#D9EFE2', '#1E6B3F'), mk(ST.WAIT, '#FBEFD5', '#8A5D00'), mk(ST.BAN, '#F8D9D7', '#9E2A22')]);
}
function setupAccess_(ss) {
  const st = ensureSheet_(ss, 'Settings');
  const have = st.getLastRow() > 1 ? st.getRange(2, 1, st.getLastRow() - 1, 1).getValues().map(r => String(r[0])) : [];
  SETTING_DEFAULTS.filter(d => have.indexOf(d[0]) < 0).forEach(d => st.appendRow(d));
  st.setColumnWidth(1, 150); st.setColumnWidth(2, 320); st.setColumnWidth(3, 520);
  const us = ensureSheet_(ss, 'Users'); statusRule_(us);
  acctRule_(ensureSheet_(ss, 'Accounts'));
  ensureSheet_(ss, 'TokenLog').getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  us.getRange('E:F').setNumberFormat('yyyy-mm-dd hh:mm');
  ensureSheet_(ss, 'Usage').getRange('A:A').setNumberFormat('yyyy-mm-dd hh:mm:ss');
  let d = ss.getSheetByName('Dashboard'); if (!d) d = ss.insertSheet('Dashboard', 0);
  d.clear();
  const F = [
    ['สรุปการใช้งาน ผู้ช่วยเขียนงานวิจัยทางการพยาบาล', ''],
    ['', ''],
    ['ผู้ใช้ทั้งหมด (บัญชี-เครื่อง)', '=COUNTA(Users!A2:A)'],
    ['อนุมัติแล้ว', '=COUNTIF(Users!D2:D,"อนุมัติ")'],
    ['รออนุมัติ', '=COUNTIF(Users!D2:D,"รออนุมัติ")'],
    ['ระงับ', '=COUNTIF(Users!D2:D,"ระงับ")'],
    ['ผู้ใช้ (อีเมลไม่ซ้ำ) วันนี้', '=IFERROR(COUNTUNIQUE(FILTER(Usage!B2:B,INT(Usage!A2:A)=TODAY())),0)'],
    ['ผู้ใช้ (อีเมลไม่ซ้ำ) 7 วันล่าสุด', '=IFERROR(COUNTUNIQUE(FILTER(Usage!B2:B,Usage!A2:A>=TODAY()-6)),0)'],
    ['ผู้ใช้ (อีเมลไม่ซ้ำ) 30 วันล่าสุด', '=IFERROR(COUNTUNIQUE(FILTER(Usage!B2:B,Usage!A2:A>=TODAY()-29)),0)'],
    ['จำนวนครั้งที่เปิดเว็บ วันนี้', '=IFERROR(COUNTIFS(Usage!D2:D,"open",Usage!A2:A,">="&TODAY()),0)'],
    ['จำนวนครั้งที่เปิดเว็บ 7 วันล่าสุด', '=IFERROR(COUNTIFS(Usage!D2:D,"open",Usage!A2:A,">="&(TODAY()-6)),0)'],
    ['จำนวนครั้งที่เปิดเว็บ ทั้งหมด', '=COUNTIF(Usage!D2:D,"open")'],
    ['', ''],
    ['บัญชีผู้ใช้ทั้งหมด', '=COUNTA(Accounts!A2:A)'],
    ['token ที่ใช้ไปทั้งหมด', '=SUM(Accounts!E2:E)'],
    ['token ที่ใช้ไป วันนี้', '=IFERROR(SUMIFS(TokenLog!G2:G,TokenLog!A2:A,">="&TODAY()),0)']
  ];
  d.getRange(1, 1, F.length, 2).setValues(F);
  d.getRange('A1').setFontSize(14).setFontWeight('bold'); d.getRange('A3:A16').setFontWeight('bold'); d.getRange('B14:B16').setNumberFormat('#,##0');
  d.getRange('D1').setValue('ผู้ใช้ที่ใช้งานบ่อย (10 อันดับ)').setFontWeight('bold');
  d.getRange('D2').setFormula('=IFERROR(QUERY(Usage!A2:D,"select B, C, count(A) where B is not null group by B, C order by count(A) desc limit 10 label B \'อีเมล\', C \'ชื่อ\', count(A) \'ครั้ง\'",0),"ยังไม่มีข้อมูล")');
  d.getRange('H1').setValue('เมนูที่ใช้บ่อย').setFontWeight('bold');
  d.getRange('H2').setFormula('=IFERROR(QUERY(Usage!A2:D,"select D, count(A) where D is not null and D <> \'open\' group by D order by count(A) desc limit 15 label D \'เมนู\', count(A) \'ครั้ง\'",0),"ยังไม่มีข้อมูล")');
  d.getRange('K1').setValue('ผู้ใช้ต่อวัน (14 วันล่าสุด)').setFontWeight('bold');
  d.getRange('K2').setFormula('=IFERROR(QUERY(Usage!A2:B,"select toDate(A), count(B) where A >= date \'"&TEXT(TODAY()-13,"yyyy-mm-dd")&"\' group by toDate(A) order by toDate(A) label toDate(A) \'วันที่\', count(B) \'ครั้ง\'",0),"ยังไม่มีข้อมูล")');
  d.setColumnWidth(1, 260); d.setColumnWidth(4, 220); d.setColumnWidth(5, 160);
}

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
