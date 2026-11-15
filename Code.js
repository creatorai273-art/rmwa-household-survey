/**
 * RMWA Household & Family Survey Portal
 * Google Apps Script backend. Spreadsheet remains private to the owner.
 *
 * SECURITY NOTES:
 * 1) Store RMWA_INITIAL_ADMIN_PASSWORD and RMWA_SPREADSHEET_ID in Apps Script Script Properties.
 * 2) Never commit Script Properties, spreadsheet IDs, passwords, or household data to public source control.
 * 3) Deploy as "Execute as: Me" and allow public access only if intended; admin actions still require login.
 * 4) Keep the linked spreadsheet private.
 */

const APP_NAME = 'RMWA Household & Family Survey';
const INITIAL_ADMIN_USERNAME = 'admin';
const SESSION_SECONDS = 21600; // 6 hours
const SHEET_NAMES = ['Users', 'Households', 'Members', 'Patients', 'AuditLog'];
const BLOOD_GROUPS = ['A+','A-','B+','B-','AB+','AB-','O+','O-','Not known'];
const YES_NO = ['yes','no'];
const MOBILE_TYPES = ['Simple','Android','iPhone'];
const PROFESSIONS = ['Student','Government Job','Private Job','Unemployed','Business Owner'];
const BPS_SCALES = Array.from({length:21}, (_, i) => 'BPS-' + (i + 1));
// Sheet column formats: every date shows as 13-Jan-2026, CNIC / phone / IDs stay plain text (no lost zeros, no auto-conversion).
const DATE_FORMATS = {createdAt:'dd-mmm-yyyy hh:mm', timestamp:'dd-mmm-yyyy hh:mm', surveyDate:'dd-mmm-yyyy', jobSince:'dd-mmm-yyyy', illnessSince:'dd-mmm-yyyy'};
const TEXT_COLUMNS = ['householdId','cnic','phone'];

function doGet(e) {
  if (e && e.parameter && e.parameter.page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('Admin').setTitle('RMWA Admin Portal').addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle(APP_NAME)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

const SCHEMAS = {
  Users: ['username','passwordHash','salt','role','active','createdAt','createdBy','mustChange'],
  Households: ['householdId','createdAt','createdBy','surveyDate','village','headName','fatherName','cnic','phone','bloodGroup','education','educationYear','educationInstitute','hasJob','jobDetails','jobSince','jobLocation','department','occupation','businessName','businessLocation','landAcres','monthlyIncome','crops','orchards','buffaloes','goats','cows','bulls','camels','donkeysDogs','chickens','pigeons','otherBirds','livestockNotes','hasIllness','consent','respondentName','hasTV','hasWashingMachine','hasSewingMachine','hasSolarPanel','hasMobile','mobileTypes','hasLaptop'],
  Members: ['householdId','memberName','relationship','cnic','phone','age','gender','bloodGroup','profession','education','educationYear','institution','department','designation','bpsScale','businessName'],
  Patients: ['householdId','patientName','age','gender','relationship','disease','illnessSince','illnessType','treatmentPlace','monthlyCostPKR','expensePayer','needsAssistance','assistanceDetails','notes'],
  AuditLog: ['timestamp','username','action','recordId','details']
};

// Creates the private database (spreadsheet + all sheets with headers) if it does not exist yet. Safe to call many times.
function initDatabase_() {
  const props = PropertiesService.getScriptProperties();
  const spreadsheetId = props.getProperty('RMWA_SPREADSHEET_ID');
  let ss;
  if (spreadsheetId) {
    ss = SpreadsheetApp.openById(spreadsheetId);
  } else {
    ss = SpreadsheetApp.create('RMWA Survey Database - PRIVATE');
    ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
    props.setProperty('RMWA_SPREADSHEET_ID', ss.getId());
  }
  Object.keys(SCHEMAS).forEach(name => {
    let sh = ss.getSheetByName(name);
    if (!sh) sh = ss.insertSheet(name);
    if (sh.getLastRow() === 0) {
      sh.getRange(1,1,1,SCHEMAS[name].length).setValues([SCHEMAS[name]]);
      sh.setFrozenRows(1);
      sh.getRange(1,1,1,SCHEMAS[name].length).setFontWeight('bold').setBackground('#174c3c').setFontColor('#ffffff');
    }
  });
  // Remove default blank Sheet1 if present and empty.
  const defaultSheet = ss.getSheetByName('Sheet1');
  if (defaultSheet && ss.getSheets().length > SHEET_NAMES.length && defaultSheet.getLastRow() === 0) ss.deleteSheet(defaultSheet);
  return ss;
}

function setupRMWA() {
  const ss = initDatabase_();
  const users = ss.getSheetByName('Users');
  if (users.getLastRow() < 2) {
    const initialPassword = PropertiesService.getScriptProperties().getProperty('RMWA_INITIAL_ADMIN_PASSWORD');
    if (!initialPassword || initialPassword.length < 12) throw new Error('Set RMWA_INITIAL_ADMIN_PASSWORD (at least 12 characters) in Project Settings > Script Properties before running setupRMWA().');
    const salt = Utilities.getUuid();
    users.appendRow([INITIAL_ADMIN_USERNAME, hashPassword_(initialPassword, salt), salt, 'admin', true, new Date(), 'setup', true]);
  }
  Logger.log('RMWA setup complete. Private database: ' + ss.getUrl());
  return {message:'Setup complete. Keep this spreadsheet private.', spreadsheetUrl:ss.getUrl()};
}

// Run ONCE from the editor (Run > finalizeDatabase) on an existing database: fixes time zone, date / text formats and column widths.
function finalizeDatabase() {
  const ss = db_();
  ss.setSpreadsheetTimeZone(Session.getScriptTimeZone());
  Object.keys(SCHEMAS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh || sh.getLastColumn() < 1) return;
    const lastCol = sh.getLastColumn();
    const headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String);
    const rows = Math.max(sh.getMaxRows() - 1, 1);
    headers.forEach((h, i) => {
      const fmt = TEXT_COLUMNS.indexOf(h) >= 0 ? '@' : DATE_FORMATS[h];
      if (fmt) sh.getRange(2, i + 1, rows, 1).setNumberFormat(fmt);
    });
    sh.setFrozenRows(1);
    sh.autoResizeColumns(1, lastCol);
  });
  SpreadsheetApp.flush();
  Logger.log('Database finalized: ' + ss.getUrl());
  return 'Database finalized.';
}

// Run from the editor after setting RMWA_SPREADSHEET_ID in Project Settings > Script Properties.
function useThisSpreadsheet() {
  const id = PropertiesService.getScriptProperties().getProperty('RMWA_SPREADSHEET_ID');
  if (!id) throw new Error('Set RMWA_SPREADSHEET_ID in Project Settings > Script Properties first.');
  const ss = SpreadsheetApp.openById(id);
  initDatabase_();
  finalizeDatabase();
  Logger.log('The app now writes to: ' + ss.getUrl());
  return ss.getUrl();
}

// Run from the editor (Run > diagnoseDatabase), then open View > Logs. Shows WHICH spreadsheet the app writes to, row counts and header problems.
function diagnoseDatabase() {
  const id = PropertiesService.getScriptProperties().getProperty('RMWA_SPREADSHEET_ID');
  Logger.log('Stored spreadsheet id: ' + (id || '(none - a new one is created on first use)'));
  const ss = db_();
  Logger.log('THE APP WRITES TO: ' + ss.getUrl() + '  | time zone: ' + ss.getSpreadsheetTimeZone() + ' (script: ' + Session.getScriptTimeZone() + ')');
  Object.keys(SCHEMAS).forEach(name => {
    const sh = ss.getSheetByName(name);
    if (!sh) { Logger.log(name + ': SHEET MISSING'); return; }
    const headers = sh.getLastColumn() ? sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(String) : [];
    const missing = SCHEMAS[name].filter(c => headers.indexOf(c) < 0);
    const extra = headers.filter(c => c && SCHEMAS[name].indexOf(c) < 0);
    Logger.log(name + ': data rows = ' + Math.max(sh.getLastRow() - 1, 0) + ' | missing headers: ' + (missing.join(', ') || 'none') + ' | unexpected headers: ' + (extra.join(', ') || 'none'));
  });
}

function login(username, password) {
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  if (!username || !password) throw new Error('Enter username and password.');
  const cache = CacheService.getScriptCache();
  const attemptsKey = 'attempts_' + username;
  const attempts = Number(cache.get(attemptsKey) || 0);
  if (attempts >= 8) throw new Error('Too many attempts. Try again in 15 minutes.');
  const sh = db_().getSheetByName('Users');
  const rows = sh.getDataRange().getValues();
  const user = rows.slice(1).find(r => String(r[0]).toLowerCase() === username && String(r[4]) === 'true');
  if (!user || hashPassword_(password, String(user[2])) !== String(user[1])) {
    cache.put(attemptsKey, String(attempts + 1), 900);
    throw new Error('Incorrect username or password.');
  }
  cache.remove(attemptsKey);
  const mc = rows[0].indexOf('mustChange'), must = mc >= 0 && String(user[mc]) === 'true';
  const token = Utilities.getUuid() + Utilities.getUuid();
  cache.put('session_' + token, JSON.stringify({username:username, role:user[3], mustChange:must}), SESSION_SECONDS);
  audit_(username, 'LOGIN', '', 'Successful login');
  return {token:token, username:username, role:user[3], mustChange:must};
}

function logout(token) {
  const session = session_(token);
  CacheService.getScriptCache().remove('session_' + token);
  if (session) audit_(session.username, 'LOGOUT', '', 'User logged out');
  return true;
}

function getDashboard(token) {
  const user = requireSession_(token);
  const ss = db_();
  const households = ss.getSheetByName('Households').getLastRow() - 1;
  const members = ss.getSheetByName('Members').getLastRow() - 1;
  const patients = ss.getSheetByName('Patients').getLastRow() - 1;
  const surveyors = ss.getSheetByName('Users').getDataRange().getValues().slice(1).filter(r => r[3] === 'surveyor' && String(r[4]) === 'true').length;
  return {username:user.username, role:user.role, households:Math.max(0,households), members:Math.max(0,members), patients:Math.max(0,patients), surveyors:surveyors};
}

function createSurveyor(token, username, password) {
  const admin = requireSession_(token, 'admin');
  username = String(username || '').trim().toLowerCase();
  password = String(password || '');
  if (!/^[a-z0-9._-]{4,30}$/.test(username)) throw new Error('Username must be 4–30 characters: letters, numbers, dot, underscore or hyphen.');
  if (password.length < 12) throw new Error('Use a password of at least 12 characters.');
  const sh = db_().getSheetByName('Users');
  const rows = sh.getDataRange().getValues();
  if (rows.slice(1).some(r => String(r[0]).toLowerCase() === username)) throw new Error('That username already exists.');
  const salt = Utilities.getUuid();
  sh.appendRow([username, hashPassword_(password, salt), salt, 'surveyor', true, new Date(), admin.username]);
  audit_(admin.username, 'CREATE_SURVEYOR', username, 'Created surveyor account');
  return {message:'Surveyor account created.'};
}

function submitSurvey(token, data) {
  const user = requireSession_(token);
  return saveSurvey_(data, user.username);
}

// Public entry point: the household survey can be submitted without an account.
// Admin functions remain protected by the separate login/session checks.
function submitPublicSurvey(data) {
  if (data && data.website) throw new Error('Invalid submission.'); // honeypot
  return saveSurvey_(data, 'public');
}

function saveSurvey_(data, submittedBy) {
  if (!data || !String(data.headName || '').trim() || !String(data.village || '').trim()) {
    throw new Error('Household head name and village are required.');
  }
  if (data.consent !== 'yes') throw new Error('Respondent consent is required before saving.');
  if (String(data.headName).length > 150 || String(data.village).length > 150) throw new Error('Some fields are too long.');
  const members = Array.isArray(data.members) ? data.members.slice(0, 40) : [];
  const patients = Array.isArray(data.patients) ? data.patients.slice(0, 20) : [];
  const CNIC_RE = /^\d{5}-\d{7}-\d$/, PHONE_RE = /^\d{4}-\d{7}$/;
  const hh = data.household || {};
  [hh].concat(members).forEach(x => {
    if (String(x.cnic || '').trim() && !CNIC_RE.test(String(x.cnic).trim())) throw new Error('CNIC must be in the format 00000-0000000-0.');
    if (String(x.phone || '').trim() && !PHONE_RE.test(String(x.phone).trim())) throw new Error('Phone must be in the format 0000-0000000.');
  });
  const lock = LockService.getScriptLock();
  lock.waitLock(20000);
  const written = [];   // blocks of rows written so far, so a failed save can be rolled back completely
  try {
    const ss = db_();
    const id = 'RMWA-' + Utilities.getUuid().split('-')[0].toUpperCase();
    const h = data.household || {};
    const income = Number(h.monthlyIncome);

    // 1) Household: one row
    written.push(appendRows_(ss.getSheetByName('Households'), [{
      householdId:id, createdAt:new Date(), createdBy:submittedBy, surveyDate:parseDate_(data.surveyDate), village:data.village, headName:data.headName,
      fatherName:h.fatherName||'', cnic:h.cnic||'', phone:h.phone||'', bloodGroup:clean_(h.bloodGroup, BLOOD_GROUPS),
      education:h.education||'', educationYear:h.educationYear||'', educationInstitute:h.educationInstitute||'',
      hasJob:h.hasJob||'no', jobDetails:h.jobDetails||'', jobSince:parseOptDate_(h.jobSince), jobLocation:h.jobLocation||'', department:h.department||'',
      occupation:h.occupation||'', businessName:h.businessName||'', businessLocation:h.businessLocation||'', landAcres:h.landAcres||'',
      monthlyIncome:(isFinite(income) && h.monthlyIncome !== '' && h.monthlyIncome != null) ? income : '',
      crops:h.crops||'', orchards:h.orchards||'', buffaloes:h.buffaloes||0, goats:h.goats||0, cows:h.cows||0, bulls:h.bulls||0, camels:h.camels||0,
      donkeysDogs:h.donkeysDogs||0, chickens:h.chickens||0, pigeons:h.pigeons||0, otherBirds:h.otherBirds||0, livestockNotes:h.livestockNotes||'',
      hasIllness:data.hasIllness||'no', consent:data.consent, respondentName:data.respondentName||'',
      hasTV:clean_(h.hasTV, YES_NO)||'no', hasWashingMachine:clean_(h.hasWashingMachine, YES_NO)||'no', hasSewingMachine:clean_(h.hasSewingMachine, YES_NO)||'no',
      hasSolarPanel:clean_(h.hasSolarPanel, YES_NO)||'no', hasLaptop:clean_(h.hasLaptop, YES_NO)||'no', hasMobile:clean_(h.hasMobile, YES_NO)||'no',
      mobileTypes:String(h.mobileTypes||'').split(',').map(x => x.trim()).filter(x => MOBILE_TYPES.indexOf(x) >= 0).join(', ')
    }]));

    // 2) Family members: all rows in one write
    const memberRows = members.filter(m => m && m.name).map(m => {
      const prof = clean_(m.profession, PROFESSIONS);
      const student = prof === 'Student', job = prof === 'Government Job' || prof === 'Private Job';
      return {
        householdId:id, memberName:m.name, relationship:m.relationship||'', cnic:m.cnic||'', phone:m.phone||'', age:m.age||'', gender:m.gender||'',
        bloodGroup:clean_(m.bloodGroup, BLOOD_GROUPS), profession:prof,
        education:student ? (m.education||'') : '', educationYear:student ? (m.educationYear||'') : '', institution:student ? (m.institution||'') : '',
        department:job ? (m.department||'') : '', designation:job ? (m.designation||'') : '',
        bpsScale:(prof === 'Government Job') ? clean_(m.bpsScale, BPS_SCALES) : '',
        businessName:(prof === 'Business Owner') ? (m.businessName||'') : ''
      };
    });
    if (memberRows.length) written.push(appendRows_(ss.getSheetByName('Members'), memberRows));

    // 3) Patients: all rows in one write
    const patientRows = patients.filter(p => p && p.name && p.disease).map(p => ({
      householdId:id, patientName:p.name, age:p.age||'', gender:p.gender||'', relationship:p.relationship||'', disease:p.disease,
      illnessSince:parseOptDate_(p.illnessSince), illnessType:p.illnessType||'', treatmentPlace:p.treatmentPlace||'',
      monthlyCostPKR:Number(p.monthlyCost) || 0, expensePayer:p.expensePayer||'', needsAssistance:p.needsAssistance||'no',
      assistanceDetails:p.assistanceDetails||'', notes:p.notes||''
    }));
    if (patientRows.length) written.push(appendRows_(ss.getSheetByName('Patients'), patientRows));

    SpreadsheetApp.flush();   // data is committed to the sheet before success is reported
    // read the household row back: if the id is not really in the sheet, fail loudly instead of reporting success
    const w0 = written[0], col = w0.headers.indexOf('householdId') + 1;
    if (!col || String(w0.sh.getRange(w0.start, col).getValue()) !== id) throw new Error('Save could not be verified in the database. Please try again.');
    audit_(submittedBy, 'SUBMIT_SURVEY', id, 'Members: '+memberRows.length+'; patients: '+patientRows.length);
    return {householdId:id, message:'Survey saved successfully.'};
  } catch (err) {
    // all-or-nothing: remove whatever part of this survey was already written
    written.reverse().forEach(w => { try { w.sh.deleteRows(w.start, w.count); } catch (e) {} });
    throw err;
  } finally {
    lock.releaseLock();
  }
}

function searchHouseholds(token, query) {
  requireSession_(token);
  query = String(query || '').trim().toLowerCase();
  if (!query) return [];
  const sh = db_().getSheetByName('Households');
  const values = sh.getDataRange().getValues();
  const headers = values[0];
  const idx = {id:headers.indexOf('householdId'), village:headers.indexOf('village'), name:headers.indexOf('headName'), date:headers.indexOf('surveyDate')};
  return values.slice(1).filter(r => [r[idx.id],r[idx.village],r[idx.name]].some(v => String(v).toLowerCase().includes(query)))
    .slice(0,50).map(r => ({id:r[idx.id], village:r[idx.village], name:r[idx.name], date:r[idx.date] ? Utilities.formatDate(new Date(r[idx.date]), Session.getScriptTimeZone(), 'dd-MMM-yyyy') : ''}));
}

function getSpreadsheetLink(token) {
  requireSession_(token, 'admin');
  return db_().getUrl();
}

// Keep only allowed values for dropdown fields (anything else is saved blank)
// '2026-01-13' (from the form) -> a real Date in the script's time zone
function parseDate_(v) {
  const m = String(v || '').match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (m) return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  const d = new Date(v);
  return isNaN(d.getTime()) ? new Date() : d;
}
function clean_(value, allowed) {
  value = String(value || '').trim();
  return allowed.indexOf(value) >= 0 ? value : '';
}
// Appends several rows in ONE write, matched by column NAME (missing columns are added at the end, so older databases keep working).
// Date columns get dd-mmm-yyyy, CNIC / phone / ID columns are forced to plain text before the values are written.
function appendRows_(sh, objs) {
  const keys = [];
  objs.forEach(o => Object.keys(o).forEach(k => { if (keys.indexOf(k) < 0) keys.push(k); }));
  const lastCol = Math.max(sh.getLastColumn(), 1);
  let headers = sh.getRange(1, 1, 1, lastCol).getValues()[0].map(String).filter(String);
  const missing = keys.filter(k => headers.indexOf(k) < 0);
  if (missing.length) {
    sh.getRange(1, headers.length + 1, 1, missing.length).setValues([missing]).setFontWeight('bold').setBackground('#174c3c').setFontColor('#ffffff');
    headers = headers.concat(missing);
  }
  const start = sh.getLastRow() + 1, count = objs.length;
  const needRows = start + count - 1 - sh.getMaxRows();
  if (needRows > 0) sh.insertRowsAfter(sh.getMaxRows(), needRows);
  headers.forEach((h, i) => {
    const fmt = TEXT_COLUMNS.indexOf(h) >= 0 ? '@' : DATE_FORMATS[h];
    if (fmt) sh.getRange(start, i + 1, count, 1).setNumberFormat(fmt);
  });
  const has = (o, k) => Object.prototype.hasOwnProperty.call(o, k);
  sh.getRange(start, 1, count, headers.length).setValues(objs.map(o => headers.map(h => has(o, h) ? o[h] : '')));
  return {sh:sh, start:start, count:count, headers:headers};
}
// Optional date: '2026-01-13' -> real Date; empty -> blank; anything else (old free text such as "2 years") is kept as typed.
function parseOptDate_(v) {
  const s = String(v || '').trim();
  if (!s) return '';
  const m = s.match(/^(\d{4})-(\d{2})-(\d{2})$/);
  if (!m) return s;
  const d = new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
  return (d.getMonth() === Number(m[2]) - 1 && d.getFullYear() >= 1900) ? d : '';
}
function session_(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('session_' + token);
  if (!raw) return null;
  try { return JSON.parse(raw); } catch(e) { return null; }
}
function requireSession_(token, role, allowChange) {
  const s = session_(token);
  if (!s) throw new Error('Your session expired. Please sign in again.');
  if (s.mustChange && !allowChange) throw new Error('Please change your password first.');
  if (role && s.role !== role) throw new Error('You do not have permission for this action.');
  return s;
}
function db_() {
  const id = PropertiesService.getScriptProperties().getProperty('RMWA_SPREADSHEET_ID');
  if (!id) return initDatabase_();   // first use: the database is created automatically
  return SpreadsheetApp.openById(id);
}
function hashPassword_(password, salt) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, salt + '|' + password, Utilities.Charset.UTF_8);
  return bytes.map(b => ('0' + ((b + 256) % 256).toString(16)).slice(-2)).join('');
}
function audit_(username, action, recordId, details) {
  try { db_().getSheetByName('AuditLog').appendRow([new Date(),username,action,recordId,details]); } catch(e) {}
}

// ===================== ADMIN PORTAL (Admin.html) =====================
function getSurveyUrl() { return ScriptApp.getService().getUrl(); }
function getAdminUrl() { return ScriptApp.getService().getUrl() + '?page=admin'; }
function col_(sh, name) {
  const h = sh.getRange(1, 1, 1, Math.max(sh.getLastColumn(), 1)).getValues()[0].map(String);
  let i = h.indexOf(name);
  if (i < 0) { i = h.length; sh.getRange(1, i + 1).setValue(name).setFontWeight('bold').setBackground('#174c3c').setFontColor('#ffffff'); }
  return i + 1;
}
// Optional password reset: set a temporary RMWA_ADMIN_RESET_PASSWORD (12+ chars)
// in Script Properties, run this function once, then remove that property.
function resetAdminPassword() {
  const props = PropertiesService.getScriptProperties();
  const resetPassword = props.getProperty('RMWA_ADMIN_RESET_PASSWORD');
  if (!resetPassword || resetPassword.length < 12) throw new Error('Set RMWA_ADMIN_RESET_PASSWORD (at least 12 characters) in Script Properties first.');
  const sh = db_().getSheetByName('Users'), c = col_(sh, 'mustChange'), rows = sh.getDataRange().getValues();
  const i = rows.findIndex((r, k) => k > 0 && String(r[0]).toLowerCase() === 'admin'), salt = Utilities.getUuid();
  if (i < 0) { sh.appendRow(['admin', hashPassword_(resetPassword, salt), salt, 'admin', true, new Date(), 'reset', true]); }
  else { sh.getRange(i + 1, 2, 1, 2).setValues([[hashPassword_(resetPassword, salt), salt]]); sh.getRange(i + 1, 4, 1, 2).setValues([['admin', true]]); sh.getRange(i + 1, c).setValue(true); }
  props.deleteProperty('RMWA_ADMIN_RESET_PASSWORD');
  CacheService.getScriptCache().remove('attempts_admin');
  Logger.log('Admin password reset. Sign in and change it immediately.');
}
function changePassword(token, cur, next) {
  const s = requireSession_(token, null, true); next = String(next || '');
  if (next.length < 12 || next.toLowerCase() === s.username) throw new Error('New password must be at least 12 characters and different from the username.');
  const sh = db_().getSheetByName('Users'), rows = sh.getDataRange().getValues();
  const i = rows.findIndex((r, k) => k > 0 && String(r[0]).toLowerCase() === s.username);
  if (i < 0 || hashPassword_(String(cur || ''), String(rows[i][2])) !== String(rows[i][1])) throw new Error('Current password is incorrect.');
  const salt = Utilities.getUuid();
  sh.getRange(i + 1, 2, 1, 2).setValues([[hashPassword_(next, salt), salt]]);
  sh.getRange(i + 1, col_(sh, 'mustChange')).setValue(false);
  s.mustChange = false; CacheService.getScriptCache().put('session_' + token, JSON.stringify(s), SESSION_SECONDS);
  audit_(s.username, 'CHANGE_PASSWORD', '', 'Password changed');
  return true;
}
function rd_(name) {
  const v = db_().getSheetByName(name).getDataRange().getValues(), h = v[0].map(String), tz = Session.getScriptTimeZone();
  return v.slice(1).filter(r => r.some(x => x !== '')).map(r => { const o = {}; h.forEach((k, i) => { o[k] = r[i] instanceof Date ? Utilities.formatDate(r[i], tz, 'dd-MMM-yyyy') : r[i]; }); return o; });
}
// Employment rule: head = Employed if job / business / occupation filled, else Unemployed. Member = by profession.
function cat_(m, head) {
  if (head) return (m.hasJob === 'yes' || m.businessName || m.occupation) ? 'Employed' : 'Unemployed';
  return m.profession === 'Student' ? 'Student' : m.profession === 'Unemployed' ? 'Unemployed' : m.profession ? 'Employed' : 'Not specified';
}
function people_(H, M) {
  const hm = {}; H.forEach(h => hm[h.householdId] = h);
  const heads = H.map(h => ({hid:h.householdId, head:h.headName, village:h.village, name:h.headName, rel:'Head', age:'', phone:h.phone, cat:cat_(h, true), detail:[h.jobDetails || h.occupation || h.businessName, h.department].filter(String).join(' / ')}));
  const mem = M.map(m => { const h = hm[m.householdId] || {}; return {hid:m.householdId, head:h.headName || '', village:h.village || '', name:m.memberName, rel:m.relationship, age:m.age, phone:m.phone, cat:cat_(m, false),
    detail:(m.profession === 'Student' ? [m.education, m.institution, m.educationYear] : [m.designation || m.businessName, m.department, m.bpsScale]).filter(String).join(' / ')}; });
  return heads.concat(mem);
}
function pats_(H, P) {
  const hm = {}; H.forEach(h => hm[h.householdId] = h);
  return P.map(p => { const h = hm[p.householdId] || {}; return {hid:p.householdId, head:h.headName || '', village:h.village || '', name:p.patientName, age:p.age, disease:p.disease, type:p.illnessType, since:p.illnessSince, cost:Number(p.monthlyCostPKR) || 0, payer:p.expensePayer, assist:p.needsAssistance}; });
}
function cnt_(arr, f) { const o = {}; arr.forEach(x => { const k = f(x) || 'Not specified'; o[k] = (o[k] || 0) + 1; }); return Object.keys(o).map(k => [k, o[k]]).sort((a, b) => b[1] - a[1]); }
function stats_(H, M, P, pp) {
  const sum = k => H.reduce((a, h) => a + (Number(h[k]) || 0), 0), yes = k => H.filter(h => h[k] === 'yes').length;
  const ageG = a => a === '' || a == null ? '' : Number(a) <= 5 ? '0-5' : Number(a) <= 17 ? '6-17' : Number(a) <= 35 ? '18-35' : Number(a) <= 59 ? '36-59' : '60+';
  const inc = h => h.monthlyIncome === '' ? 'Not reported' : h.monthlyIncome < 20000 ? 'Below 20,000' : h.monthlyIncome < 50000 ? '20,000-49,999' : h.monthlyIncome < 100000 ? '50,000-99,999' : '100,000+';
  const n = c => pp.filter(p => p.cat === c).length;
  return {
    kpi:{households:H.length, people:H.length + M.length, employed:n('Employed'), unemployed:n('Unemployed'), students:n('Student'), patients:P.length, illHouseholds:H.filter(h => h.hasIllness === 'yes').length,
      avgFamily:H.length ? +((H.length + M.length) / H.length).toFixed(1) : 0, income:sum('monthlyIncome'), land:+sum('landAcres').toFixed(2), medical:P.reduce((a, p) => a + (Number(p.monthlyCostPKR) || 0), 0), assist:P.filter(p => p.needsAssistance === 'yes').length},
    villages:cnt_(H, h => h.village).slice(0, 10), gender:cnt_(M, m => m.gender), age:cnt_(M, m => ageG(m.age)), blood:cnt_(H.concat(M), x => x.bloodGroup), employment:cnt_(pp, p => p.cat),
    assets:[['TV','hasTV'],['Washing machine','hasWashingMachine'],['Sewing machine','hasSewingMachine'],['Solar panel','hasSolarPanel'],['Laptop','hasLaptop'],['Mobile phone','hasMobile']].map(a => [a[0], yes(a[1])]),
    livestock:[['Buffaloes','buffaloes'],['Cows','cows'],['Goats','goats'],['Bulls / oxen','bulls'],['Camels','camels'],['Chickens','chickens']].map(a => [a[0], sum(a[1])]),
    income:cnt_(H, inc), diseases:cnt_(P, p => p.disease).slice(0, 8), illType:cnt_(P, p => p.illnessType)
  };
}
function getPortalData(token) {
  requireSession_(token);
  const H = rd_('Households'), M = rd_('Members'), P = rd_('Patients'), pp = people_(H, M);
  return {stats:stats_(H, M, P, pp), people:pp, patients:pats_(H, P),
    households:H.map(h => ({id:h.householdId, date:h.surveyDate, village:h.village, head:h.headName, father:h.fatherName, phone:h.phone, income:h.monthlyIncome, ill:h.hasIllness,
      size:1 + M.filter(m => m.householdId === h.householdId).length}))};
}
// ---------- HTML reports (shown in the portal and converted to PDF) ----------
function esc_(s) { return String(s == null ? '' : s).replace(/[&<>"]/g, c => ({'&':'&amp;','<':'&lt;','>':'&gt;','"':'&quot;'}[c])); }
function tbl_(heads, rows) {
  return '<table class="t"><tr>' + heads.map(h => '<th>' + esc_(h) + '</th>').join('') + '</tr>' +
    (rows.length ? rows.map(r => '<tr>' + r.map(c => '<td>' + esc_(c) + '</td>').join('') + '</tr>').join('') : '<tr><td colspan="' + heads.length + '">No records</td></tr>') + '</table>';
}
function kv_(pairs) { let h = '<table class="t">'; for (let i = 0; i < pairs.length; i += 2) h += '<tr><td class="k">' + esc_(pairs[i][0]) + '</td><td>' + esc_(pairs[i][1]) + '</td>' + (pairs[i + 1] ? '<td class="k">' + esc_(pairs[i + 1][0]) + '</td><td>' + esc_(pairs[i + 1][1]) + '</td>' : '<td></td><td></td>') + '</tr>'; return h + '</table>'; }
const REPORT_CSS_ = '<style>body{font-family:Arial,sans-serif;font-size:10px;color:#17231e}h1{font-size:18px;color:#174c3c;margin:0}h2{font-size:12px;color:#174c3c;border-bottom:2px solid #174c3c;padding-bottom:3px;margin:14px 0 6px}.sub{color:#64736b;font-size:9px}.t{width:100%;border-collapse:collapse;margin-bottom:8px}.t th{background:#174c3c;color:#fff;text-align:left;padding:4px;font-size:9px}.t td{border-bottom:1px solid #dce5df;padding:4px}.t td.k{background:#f3f6f4;font-weight:bold;width:17%}.card{page-break-after:always}</style>';
function page_(title, body) {
  return '<html><head><meta charset="utf-8">' + REPORT_CSS_ + '</head><body><h1>RMWA - Razal Memon Welfare Association</h1><div class="sub">' + esc_(title) + ' &middot; Generated ' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy HH:mm') + '</div>' + body + '</body></html>';
}
function famBody_(id, H, M, P) {
  const h = H.filter(x => x.householdId === id)[0]; if (!h) throw new Error('Household not found.');
  const yn = v => v === 'yes' ? 'Yes' : 'No';
  const mem = M.filter(m => m.householdId === id), pat = P.filter(p => p.householdId === id);
  return '<h2>Family report card - ' + esc_(h.headName) + ' (' + esc_(id) + ')</h2>' +
    kv_([['Household ID', id], ['Survey date', h.surveyDate], ['Head of household', h.headName], ['Father\'s name', h.fatherName], ['Village / area', h.village], ['CNIC', h.cnic], ['Phone', h.phone], ['Blood group', h.bloodGroup],
      ['Education', [h.education, h.educationYear, h.educationInstitute].filter(String).join(' / ')], ['Employment', cat_(h, true)], ['Job / designation', [h.jobDetails, h.department, h.jobLocation].filter(String).join(' / ')], ['Occupation / business', [h.occupation, h.businessName].filter(String).join(' / ')],
      ['Monthly income (PKR)', h.monthlyIncome], ['Land (acres)', h.landAcres], ['Crops / orchards', [h.crops, h.orchards].filter(String).join(' / ')], ['Respondent', h.respondentName]]) +
    '<h2>Household assets</h2>' + kv_([['TV', yn(h.hasTV)], ['Washing machine', yn(h.hasWashingMachine)], ['Sewing machine', yn(h.hasSewingMachine)], ['Solar panel', yn(h.hasSolarPanel)], ['Laptop', yn(h.hasLaptop)], ['Mobile', yn(h.hasMobile) + (h.mobileTypes ? ' (' + h.mobileTypes + ')' : '')]]) +
    '<h2>Livestock</h2>' + kv_([['Buffaloes', h.buffaloes], ['Cows', h.cows], ['Goats', h.goats], ['Bulls / oxen', h.bulls], ['Camels', h.camels], ['Donkeys / dogs', h.donkeysDogs], ['Chickens', h.chickens], ['Pigeons', h.pigeons]]) +
    '<h2>Family members (' + mem.length + ')</h2>' + tbl_(['Name', 'Relation', 'Age', 'Gender', 'Blood', 'Status', 'Details'], mem.map(m => [m.memberName, m.relationship, m.age, m.gender, m.bloodGroup, cat_(m, false), people_([h], [m])[1].detail])) +
    '<h2>Patients (' + pat.length + ')</h2>' + tbl_(['Name', 'Age', 'Disease', 'Type', 'Since', 'Monthly cost', 'Needs help'], pat.map(p => [p.patientName, p.age, p.disease, p.illnessType, p.illnessSince, p.monthlyCostPKR, p.needsAssistance]));
}
function getFamilyHtml(token, id) {
  requireSession_(token); const H = rd_('Households');
  return '<style>' + REPORT_CSS_.replace(/<\/?style>/g, '') + '</style>' + famBody_(id, H, rd_('Members'), rd_('Patients'));
}
function reportHtml_(type, arg) {
  const H = rd_('Households'), M = rd_('Members'), P = rd_('Patients'), pp = people_(H, M), S = stats_(H, M, P, pp), pl = c => pp.filter(p => p.cat === c).map(p => [p.name, p.rel, p.age, p.head, p.village, p.phone, p.detail]);
  const ph = ['Name', 'Relation', 'Age', 'Household head', 'Village', 'Phone', 'Details'];
  if (type === 'family') return page_('Family report card', famBody_(arg, H, M, P));
  if (type === 'employed') return page_('Employed persons', tbl_(ph, pl('Employed')));
  if (type === 'unemployed') return page_('Unemployed persons', tbl_(ph, pl('Unemployed')));
  if (type === 'students') return page_('Students', tbl_(['Name', 'Relation', 'Age', 'Household head', 'Village', 'Phone', 'Class / institution / year'], pl('Student')));
  if (type === 'patients') return page_('Patients', tbl_(['Patient', 'Age', 'Disease', 'Type', 'Since', 'Monthly cost', 'Payer', 'Needs help', 'Household head', 'Village'], pats_(H, P).map(p => [p.name, p.age, p.disease, p.type, p.since, p.cost, p.payer, p.assist, p.head, p.village])));
  if (type === 'households') return page_('Household list', tbl_(['ID', 'Date', 'Village', 'Head', 'Father', 'Phone', 'Income (PKR)'], H.map(h => [h.householdId, h.surveyDate, h.village, h.headName, h.fatherName, h.phone, h.monthlyIncome])));
  const k = S.kpi, t2 = (a, label) => '<h2>' + label + '</h2>' + tbl_(['Item', 'Count'], a);
  return page_('Overall survey summary', '<h2>Key figures</h2>' + kv_([['Households', k.households], ['Total people', k.people], ['Employed', k.employed], ['Unemployed', k.unemployed], ['Students', k.students], ['Patients', k.patients], ['Average family size', k.avgFamily], ['Households with illness', k.illHouseholds],
    ['Total monthly income (PKR)', k.income], ['Land (acres)', k.land], ['Monthly medical cost (PKR)', k.medical], ['Patients needing help', k.assist]]) +
    t2(S.villages, 'Households by village') + t2(S.employment, 'Employment status') + t2(S.gender, 'Gender (family members)') + t2(S.age, 'Age groups (family members)') + t2(S.blood, 'Blood groups') +
    t2(S.assets, 'Households owning assets') + t2(S.livestock, 'Livestock (total)') + t2(S.income, 'Household income bands (PKR / month)') + t2(S.diseases, 'Top illnesses') + t2(S.illType, 'Illness type'));
}
function getReportPdf(token, type, arg) {
  requireSession_(token);
  const pdf = Utilities.newBlob(reportHtml_(String(type), String(arg || '')), 'text/html', 'r.html').getAs('application/pdf');
  return {name:'RMWA-' + type + (arg ? '-' + arg : '') + '-' + Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MMM-yyyy') + '.pdf', data:Utilities.base64Encode(pdf.getBytes())};
}