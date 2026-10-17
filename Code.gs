/** CRP — Google Apps Script (bound to Google Sheets) */
const CRP = Object.freeze({
  applications: 'Applications',
  audit: 'AuditLog',
  sessionSeconds: 21600,
  statuses: ['รอตรวจสอบ', 'อยู่ระหว่างพิจารณา', 'ผ่านการคัดเลือก', 'ไม่ผ่านการคัดเลือก'],
  checks: [
    'การกำหนดรหัสผู้สมัคร',
    'มีอายุตามกำหนด',
    'ไม่เป็นผู้ดำรงตำแหน่งทางการเมือง',
    'ไม่เป็นบุคคลล้มละลาย',
    'ไม่เป็นบุคคลต้องโทษหรือต้องคำพิพากษา',
    'รูปถ่าย',
    'วุฒิการศึกษา',
    'ผลการเรียน',
    'ทะเบียนบ้านและบัตรประชาชน',
    'สำเนาข้อมูลครอบครัว',
    'หนังสืออนุมัติ'
  ]
});

// คงหัวคอลัมน์เดิมไว้ เพื่อไม่ทำให้ข้อมูลเก่าสูญหาย
const HEADERS = [
  'ApplicationID', 'CreatedAt', 'Prefix', 'FirstName', 'LastName',
  'BirthDate', 'Phone', 'Email', 'Province', 'Education', 'Motivation',
  'Status', 'AdminNote', 'UpdatedAt',
  'Nationality', 'Ethnicity', 'ReviewStatus', 'ReviewData', 'DecisionReason', 'Region'
];
const AUDIT_HEADERS = ['At', 'Actor', 'Action', 'ApplicationID', 'Details'];
const REVIEW_OPTIONS = ['ผ่าน', 'ไม่ผ่าน'];

function setup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ต้องเปิด Apps Script จาก Google Sheets ผ่าน Extensions > Apps Script');
  ensureSheet_(ss, CRP.applications, HEADERS);
  ensureSheet_(ss, CRP.audit, AUDIT_HEADERS);
  return 'เตรียมระบบ CRP สำเร็จ (ข้อมูลเดิมยังอยู่)';
}

/** ตั้งค่าบัญชีใหม่: เปลี่ยนข้อความตัวอย่างก่อนเรียกใช้ */
function changeAdminCredentials() {
  setAdminCredentials_('CHANGE_USERNAME', 'CHANGE_STRONG_PASSWORD');
  return 'ตั้งค่าบัญชีเจ้าหน้าที่สำเร็จ';
}

function setAdminCredentials_(username, password) {
  if (!username || !password || username === 'CHANGE_USERNAME' || password === 'CHANGE_STRONG_PASSWORD') {
    throw new Error('กรุณาเปลี่ยนชื่อผู้ใช้และรหัสผ่านตัวอย่างก่อน');
  }
  const salt = Utilities.getUuid() + Utilities.getUuid();
  PropertiesService.getScriptProperties().setProperties({
    ADMIN_USER: String(username),
    ADMIN_SALT: salt,
    ADMIN_HASH: sha256_(salt + ':' + password)
  });
}

function doGet(e) {
  const page = String(e && e.parameter && e.parameter.page || '').toLowerCase() === 'admin' ? 'admin' : 'index';
  const template = HtmlService.createTemplateFromFile(page);
  template.appUrl = ScriptApp.getService().getUrl() || '';
  return template.evaluate()
    .setTitle(page === 'admin' ? 'CRP | เจ้าหน้าที่' : 'CRP | ตรวจสอบผู้สมัคร')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function apiRequest(action, payload) {
  const p = payload && typeof payload === 'object' && !Array.isArray(payload) ? payload : {};
  switch (String(action)) {
    case 'ping': return { message: 'CRP Backend Online', now: new Date().toISOString() };
    case 'publicSearch': return publicSearch_(p);
    case 'publicDetails': return publicDetails_(p);
    case 'login': return login_(p);
    case 'logout': return logout_(p);
    case 'getApplications': return getApplications_(p);
    case 'getApplication': return getApplication_(p);
    case 'addApplication': return addApplication_(p);
    case 'updateApplication': return updateApplication_(p);
    case 'saveReview': return saveReview_(p);
    case 'setApplicationStatus': return setApplicationStatus_(p);
    case 'getSummary': return getSummary_(p);
    case 'getConfig': requireAdmin_(p); return { checks: CRP.checks, statuses: CRP.statuses };
    default: throw new Error('ไม่รองรับคำสั่งนี้');
  }
}

// ---------- Authentication ----------
function login_(p) {
  const props = PropertiesService.getScriptProperties();
  const username = clean_(p.username, 100);
  const password = String(p.password || '');
  const cache = CacheService.getScriptCache();
  const failureKey = 'CRP_LOGIN_FAILURES_' + sha256_(username).slice(0, 24);
  const failures = Number(cache.get(failureKey) || 0);
  if (failures >= 10) throw new Error('เข้าสู่ระบบผิดพลาดหลายครั้ง กรุณารอ 5 นาที');
  const expected = props.getProperty('ADMIN_HASH');
  if (!expected) throw new Error('กรุณาตั้งค่าบัญชีเจ้าหน้าที่ก่อน');
  const actual = sha256_(String(props.getProperty('ADMIN_SALT') || '') + ':' + password);
  if (username !== props.getProperty('ADMIN_USER') || actual !== expected) {
    cache.put(failureKey, String(failures + 1), 300);
    throw new Error('ชื่อผู้ใช้หรือรหัสผ่านไม่ถูกต้อง');
  }
  cache.remove(failureKey);
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('CRP_SESSION_' + sha256_(token), username, CRP.sessionSeconds);
  log_('login', '', username, '');
  return { token: token, username: username, expiresIn: CRP.sessionSeconds };
}

function requireAdmin_(p) {
  const token = String(p.token || '');
  if (!/^[0-9a-f-]{72}$/i.test(token)) throw new Error('กรุณาเข้าสู่ระบบเจ้าหน้าที่');
  const user = CacheService.getScriptCache().get('CRP_SESSION_' + sha256_(token));
  if (!user) throw new Error('เซสชันหมดอายุ กรุณาเข้าสู่ระบบใหม่');
  return user;
}

function logout_(p) {
  const token = String(p.token || '');
  if (token) CacheService.getScriptCache().remove('CRP_SESSION_' + sha256_(token));
  return { success: true };
}

// ---------- Public: ค้นหาด้วยรหัสผู้สมัครเท่านั้น เพื่อป้องกันการเปิดเผยข้อมูลส่วนบุคคล ----------
function publicSearch_(p) {
  const query = normalize_(clean_(p.query, 120));
  const matching = allRows_().filter(r => !query || normalize_(r.ApplicationID + ' ' + r.Prefix + ' ' + r.FirstName + ' ' + r.LastName).includes(query)).map(publicView_);
  const offset = Math.max(0, Math.floor(Number(p.offset) || 0));
  const limit = Math.min(200, Math.max(1, Math.floor(Number(p.limit) || 100)));
  return { total: matching.length, items: matching.slice(offset, offset + limit) };
}

function publicDetails_(p) {
  const id = clean_(p.applicationId, 50);
  const r = allRows_().find(x => x.ApplicationID === id);
  if (!r) throw new Error('ไม่พบข้อมูลผู้สมัคร');
  return publicView_(r);
}

function publicView_(r) {
  return {
    applicationId: r.ApplicationID,
    fullName: [r.Prefix, r.FirstName, r.LastName].filter(Boolean).join(' '),
    province: r.Province,
    region: r.Region,
    education: r.Education,
    registeredAt: r.CreatedAt,
    registration: 'ลงทะเบียนสำเร็จ',
    status: r.Status,
    reviewStatus: r.ReviewStatus || 'รอตรวจสอบ'
  };
}

// ---------- Applications ----------
function getApplications_(p) {
  requireAdmin_(p);
  const query = normalize_(clean_(p.query, 120));
  const stage = clean_(p.stage, 30);
  const rows = allRows_().filter(r => {
    const match = !query || normalize_(r.ApplicationID + ' ' + r.Prefix + ' ' + r.FirstName + ' ' + r.LastName).includes(query);
    if (stage === 'review') return match && r.ReviewStatus !== 'ตรวจสอบแล้ว';
    if (stage === 'approval') return match && r.ReviewStatus === 'ตรวจสอบแล้ว' && !['ผ่านการคัดเลือก', 'ไม่ผ่านการคัดเลือก'].includes(r.Status);
    return match;
  });
  const offset = Math.max(0, Math.floor(Number(p.offset) || 0));
  const limit = Math.min(500, Math.max(1, Math.floor(Number(p.limit) || 200)));
  return { total: rows.length, items: rows.slice(offset, offset + limit), offset: offset, limit: limit };
}

function getApplication_(p) {
  requireAdmin_(p);
  const r = allRows_().find(x => x.ApplicationID === clean_(p.applicationId, 50));
  if (!r) throw new Error('ไม่พบข้อมูลผู้สมัคร');
  return r;
}

function addApplication_(p) {
  const admin = requireAdmin_(p);
  const input = validateInput_(p.application || {});
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const id = validateApplicationId_((p.application || {}).applicationId);
    if (findRow_(id)) throw new Error('รหัสผู้สมัครนี้มีอยู่ในระบบแล้ว กรุณาใช้รหัสอื่น');
    const now = new Date().toISOString();
    const record = Object.assign({ ApplicationID: id, CreatedAt: now }, input, {
      Status: 'รอตรวจสอบ', AdminNote: '', UpdatedAt: now,
      ReviewStatus: 'รอตรวจสอบ', ReviewData: '', DecisionReason: ''
    });
    sheet_().appendRow(HEADERS.map(k => safeCell_(record[k])));
    log_('add', id, admin, '');
    return { success: true, applicationId: id, record: record };
  } finally {
    lock.releaseLock();
  }
}

function updateApplication_(p) {
  const admin = requireAdmin_(p);
  const id = clean_(p.applicationId, 50);
  const input = validateInput_(p.application || {});
  const newId = validateApplicationId_((p.application || {}).applicationId || id);
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const match = findRow_(id);
    if (newId !== id && findRow_(newId)) throw new Error('รหัสผู้สมัครใหม่ซ้ำกับข้อมูลที่มีอยู่');
    if (!match) throw new Error('ไม่พบข้อมูลผู้สมัคร');
    const record = Object.assign({}, match.record, input, { ApplicationID: newId, UpdatedAt: new Date().toISOString() });
    writeRecord_(match, record);
    log_('update', id, admin, '');
    return { success: true, applicationId: newId, record: record };
  } finally {
    lock.releaseLock();
  }
}

// ---------- Review -> Approval ----------
function saveReview_(p) {
  const admin = requireAdmin_(p);
  const id = clean_(p.applicationId, 50);
  const checks = p.checks && typeof p.checks === 'object' ? p.checks : {};
  const result = {};
  CRP.checks.forEach(name => {
    const value = clean_(checks[name], 30);
    const allowed = name === 'หนังสืออนุมัติ' ? ['ผ่าน', 'ไม่ผ่าน', 'ไม่ระบุ'] : REVIEW_OPTIONS;
    if (!allowed.includes(value)) throw new Error('กรุณาตรวจสอบหัวข้อ: ' + name);
    result[name] = value;
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const match = findRow_(id);
    if (!match) throw new Error('ไม่พบข้อมูลผู้สมัคร');
    const record = Object.assign({}, match.record, {
      ReviewStatus: 'ตรวจสอบแล้ว',
      ReviewData: JSON.stringify(result),
      Status: ['ผ่านการคัดเลือก', 'ไม่ผ่านการคัดเลือก'].includes(match.record.Status) ? match.record.Status : 'อยู่ระหว่างพิจารณา',
      UpdatedAt: new Date().toISOString()
    });
    writeRecord_(match, record);
    log_('review', id, admin, 'ตรวจสอบแล้ว');
    return { success: true, applicationId: id, record: record };
  } finally {
    lock.releaseLock();
  }
}

function setApplicationStatus_(p) {
  const admin = requireAdmin_(p);
  const id = clean_(p.applicationId, 50);
  const status = clean_(p.status, 50);
  const reason = clean_(p.reason || p.note, 2000);
  if (!['ผ่านการคัดเลือก', 'ไม่ผ่านการคัดเลือก'].includes(status)) throw new Error('กรุณาเลือกผลผ่านหรือไม่ผ่าน');
  if (!reason) throw new Error('กรุณาระบุเหตุผลการพิจารณา');
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  try {
    const match = findRow_(id);
    if (!match) throw new Error('ไม่พบข้อมูลผู้สมัคร');
    if (match.record.ReviewStatus !== 'ตรวจสอบแล้ว') throw new Error('กรุณาตรวจสอบข้อมูลก่อนอนุมัติ');
    const record = Object.assign({}, match.record, {
      Status: status,
      AdminNote: reason,
      DecisionReason: reason,
      UpdatedAt: new Date().toISOString()
    });
    writeRecord_(match, record);
    log_('decision', id, admin, status);
    return { success: true, applicationId: id, status: status };
  } finally {
    lock.releaseLock();
  }
}

function getSummary_(p) {
  requireAdmin_(p);
  const rows = allRows_();
  const byStatus = Object.fromEntries(CRP.statuses.map(s => [s, 0]));
  const byProvince = {};
  rows.forEach(r => {
    if (Object.prototype.hasOwnProperty.call(byStatus, r.Status)) byStatus[r.Status]++;
    const province = r.Province || 'ไม่ระบุ';
    byProvince[province] = (byProvince[province] || 0) + 1;
  });
  return {
    total: rows.length,
    byStatus: byStatus,
    byProvince: byProvince,
    latest: rows.slice(0, 10).map(r => ({
      applicationId: r.ApplicationID,
      fullName: [r.Prefix, r.FirstName, r.LastName].filter(Boolean).join(' '),
      status: r.Status,
      createdAt: r.CreatedAt
    }))
  };
}

// ---------- Validation ----------
function validateApplicationId_(value) {
  const id = clean_(value, 50);
  if (!id) throw new Error('กรุณากรอกรหัสผู้สมัคร');
  if (!/^[a-zA-Z0-9_-]{2,50}$/.test(id)) {
    throw new Error('รหัสผู้สมัครใช้ได้เฉพาะตัวอักษรอังกฤษ ตัวเลข ขีดกลาง และขีดล่าง (2-50 ตัว)');
  }
  return id;
}

function validateInput_(obj) {
  const r = {
    Prefix: clean_(obj.prefix, 30),
    FirstName: clean_(obj.firstName, 100),
    LastName: clean_(obj.lastName, 100),
    BirthDate: clean_(obj.birthDate, 30),
    Phone: clean_(obj.phone, 30),
    Email: clean_(obj.email, 150),
    Province: clean_(obj.province, 100),
    Region: clean_(obj.region, 60),
    Education: clean_(obj.education, 100),
    Motivation: clean_(obj.motivation, 3000),
    Nationality: clean_(obj.nationality, 100),
    Ethnicity: clean_(obj.ethnicity, 100)
  };
  if (!r.Prefix || !r.FirstName || !r.LastName) throw new Error('กรุณาระบุคำนำหน้า ชื่อ และนามสกุล');
  if (!/^\d{10}$/.test(r.Phone)) throw new Error('เบอร์โทรศัพท์ต้องเป็นตัวเลข 10 หลัก');
  if (!r.Education || !r.Nationality || !r.Ethnicity) throw new Error('กรุณาระบุสัญชาติ เชื้อชาติ และระดับการศึกษา');
  if (!/^\d{4}-\d{2}-\d{2}$/.test(r.BirthDate)) throw new Error('กรุณาระบุวันเดือนปีเกิด');
  const birth = new Date(r.BirthDate + 'T00:00:00Z');
  if (Number.isNaN(birth.getTime()) || birth.toISOString().slice(0, 10) !== r.BirthDate || birth > new Date()) {
    throw new Error('วันเดือนปีเกิดไม่ถูกต้อง');
  }
  if (r.Region && !['ภาคเหนือ', 'ภาคตะวันออกเฉียงเหนือ', 'ภาคกลาง', 'ภาคตะวันออก', 'ภาคตะวันตก', 'ภาคใต้'].includes(r.Region)) throw new Error('ภูมิภาคไม่ถูกต้อง');
  if (r.Ethnicity && !['ไทย', 'อเมริกา', 'อื่นๆ', 'อื่น ๆ'].includes(r.Ethnicity)) throw new Error('เชื้อชาติไม่ถูกต้อง');
  if (r.Email && !/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(r.Email)) throw new Error('อีเมลไม่ถูกต้อง');
  return r;
}

// ---------- Google Sheets ----------
function sheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) throw new Error('ไม่พบ Google Sheets ที่ผูกกับสคริปต์');
  const sh = ss.getSheetByName(CRP.applications);
  if (!sh) throw new Error('กรุณารัน setup() ก่อน');
  return sh;
}

function allRows_() {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return [];
  return sh.getRange(2, 1, last - 1, HEADERS.length).getDisplayValues()
    .filter(row => row[0])
    .map(row => Object.fromEntries(HEADERS.map((key, i) => [key, row[i]])))
    .reverse();
}

function findRow_(id) {
  const sh = sheet_();
  const last = sh.getLastRow();
  if (last < 2) return null;
  const ids = sh.getRange(2, 1, last - 1, 1).getDisplayValues().flat();
  const index = ids.indexOf(id);
  if (index < 0) return null;
  const values = sh.getRange(index + 2, 1, 1, HEADERS.length).getDisplayValues()[0];
  return {
    sh: sh,
    index: index + 2,
    record: Object.fromEntries(HEADERS.map((key, i) => [key, values[i]]))
  };
}

function writeRecord_(match, record) {
  match.sh.getRange(match.index, 1, 1, HEADERS.length)
    .setValues([HEADERS.map(key => safeCell_(record[key]))]);
}

/** เพิ่มเฉพาะคอลัมน์ท้ายตารางที่ยังไม่มี ไม่ลบหรือสลับข้อมูลเดิม */
function ensureSheet_(ss, name, headers) {
  const sh = ss.getSheetByName(name) || ss.insertSheet(name);
  if (sh.getMaxColumns() < headers.length) {
    sh.insertColumnsAfter(sh.getMaxColumns(), headers.length - sh.getMaxColumns());
  }
  if (sh.getLastRow() === 0) {
    sh.getRange(1, 1, 1, headers.length).setValues([headers]);
  } else {
    const current = sh.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    for (let i = 0; i < headers.length; i++) {
      if (current[i] && current[i] !== headers[i]) {
        throw new Error('หัวคอลัมน์ ' + name + ' ไม่ตรงกันที่คอลัมน์ ' + (i + 1));
      }
      if (!current[i]) sh.getRange(1, i + 1).setValue(headers[i]);
    }
  }
  sh.setFrozenRows(1);
  sh.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#eaf5ff');
  return sh;
}

function log_(action, id, actor, details) {
  const sh = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(CRP.audit);
  if (sh) sh.appendRow([new Date().toISOString(), safeCell_(actor), safeCell_(action), safeCell_(id), safeCell_(details)]);
}

function clean_(value, max) {
  return String(value == null ? '' : value).trim().slice(0, max);
}

function normalize_(value) {
  return String(value || '').toLocaleLowerCase().replace(/\s+/g, ' ').trim();
}

function safeCell_(value) {
  const str = String(value == null ? '' : value);
  return /^[=+\-@\t\r]/.test(str) ? "'" + str : str;
}

function sha256_(str) {
  return Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, str, Utilities.Charset.UTF_8)
    .map(b => ('0' + (b & 255).toString(16)).slice(-2)).join('');
}
