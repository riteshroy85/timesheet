/**
 * Timesheet API (Google Apps Script). The phone screen is hosted separately (GitHub Pages)
 * and calls this web app with POST requests.
 *
 * Bound to a Google Sheet with three tabs:
 *   Users      : Name | Email | PIN | Active (TRUE/FALSE) | Admin (TRUE for admins)
 *   Settings   : A1 "Registration code", B1 the code staff need to register
 *   Timesheet  : Name | Email | Date | Day | Start Time | End Time | Break (hrs) | Total Hrs | In-Hrs   (created automatically)
 * An employee is identified by name + email. Email (lower-case) is the unique key; the name must match it.
 * Break is in hours (e.g. 0.75). Total Hrs = End - Start. In-Hrs = Total Hrs - Break.
 */

var USERS_SHEET = 'Users';
var LOG_SHEET = 'Timesheet';
var SETTINGS_SHEET = 'Settings';
var MAX_FAILS = 5;           // wrong PINs allowed per email and per device before locking out...
var MAX_FAILS_GLOBAL = 30;   // ...wrong registration codes allowed across all devices
var LOCKOUT_SECONDS = 600;
var PIN_MIN = 6;             // PIN length rules for new registrations
var PIN_MAX = 8;

// Timesheet column numbers
var C = { NAME: 1, EMAIL: 2, DATE: 3, DAY: 4, START: 5, END: 6, BREAK: 7, TOTAL: 8, IN: 9 };
var HEADERS = ['Name', 'Email', 'Date', 'Day', 'Start Time', 'End Time', 'Break (hrs)', 'Total Hrs', 'In-Hrs'];
var USER_HEADERS = ['Name', 'Email', 'PIN', 'Active', 'Admin'];
var EMAIL_RE = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;

/**
 * Run once from the Apps Script editor (select "setup" > Run) on a blank Sheet.
 * Safe to run again: it also upgrades older Users / Timesheet tabs to the current columns
 * (existing people get a blank Email that you then fill in).
 * No default user is created. To make yourself the first admin, add a row to Users:
 *   Your Name | your@email | (leave PIN blank) | TRUE | TRUE
 * then open the app, tap "First time? Register" and set your PIN with the Settings!B1 code.
 */
function setup() {
  var ss = SpreadsheetApp.getActive();

  var users = ss.getSheetByName(USERS_SHEET) || ss.insertSheet(USERS_SHEET);
  if (users.getLastRow() === 0) {
    users.appendRow(USER_HEADERS);
    users.setFrozenRows(1);
  } else if (String(users.getRange('B1').getValue()) === 'PIN') {
    users.insertColumnAfter(1); // older layout: Name | PIN | Active
  }
  users.getRange(1, 1, 1, USER_HEADERS.length).setValues([USER_HEADERS]);
  users.getRange('B:C').setNumberFormat('@'); // keep leading zeros in PINs
  styleHeader_(users.getRange(1, 1, 1, USER_HEADERS.length));

  // Registration code staff need to create their own PIN. Blank it to switch registration off.
  var settings = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  if (settings.getLastRow() === 0) {
    settings.appendRow(['Registration code', Math.random().toString(36).slice(2, 8).toUpperCase()]);
    settings.getRange('B1').setNumberFormat('@');
    settings.getRange('A1').setFontWeight('bold');
    settings.setColumnWidth(1, 160);
  }

  var log = logSheet_();
  if (String(log.getRange('B1').getValue()) === 'Date') {
    log.insertColumnAfter(1); // older layout: Name | Date | ...; old rows keep a blank Email
  }
  log.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
  log.getRange('A:F').setNumberFormat('@');
  styleHeader_(log.getRange(1, 1, 1, HEADERS.length));

  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function styleHeader_(range) {
  range.setBackground('#1f5c80').setFontColor('#ffffff').setFontWeight('bold');
}

function doGet() {
  return ContentService.createTextOutput('Timesheet API is running.');
}

/**
 * Body (JSON sent as text/plain): { action, email, pin, deviceId, ... }
 *   register : name, email, code, newPin -> { name, entry }  (no pin needed; creates or claims a user)
 *   login : email + pin -> { name, admin, entry }            (entry for today; PINs need not be unique)
 *   get   : date                -> entry
 *   save  : date, start, end, breakHrs -> entry
 *   admin_list | admin_add | admin_reset | admin_active | admin_delete | admin_code   (admins only)
 */
function doPost(e) {
  try {
    var req = JSON.parse(e.postData.contents);
    var deviceId = String(req.deviceId || '').slice(0, 64);
    if (req.action === 'register') {
      var created = register_(req, deviceId);
      return json_({ ok: true, data: { name: created.name, admin: created.admin, entry: entry_(created.email, today_()) } });
    }
    var user = authenticate_(req.email, req.pin, deviceId);
    if (String(req.action).indexOf('admin_') === 0) {
      if (!user.admin) throw new Error('Admins only.');
      return json_({ ok: true, data: admin_(user, req) });
    }
    if (req.action === 'login') {
      return json_({ ok: true, data: { name: user.name, admin: user.admin, entry: entry_(user.email, today_()) } });
    } else if (req.action === 'get') {
      return json_({ ok: true, data: entry_(user.email, checkDate_(req.date)) });
    } else if (req.action === 'save') {
      return json_({ ok: true, data: save_(user, req) });
    }
    throw new Error('Unknown action.');
  } catch (err) {
    return json_({ ok: false, error: String(err.message || err) });
  }
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * First-time sign-up: needs the registration code from the Settings tab.
 * Email is the unique key. A new email creates a user. An email already listed with a blank PIN
 * (added by the admin, or PIN cleared for a reset) is claimed, but the name must match the one on file.
 * Returns { name, email }.
 */
function register_(req, deviceId) {
  var keys = [['fails_d_' + deviceId, MAX_FAILS], ['fails_code', MAX_FAILS_GLOBAL]];
  checkLockout_(keys);

  var name = String(req.name || '').replace(/\s+/g, ' ').trim();
  var email = String(req.email || '').trim().toLowerCase();
  var newPin = String(req.newPin || '').trim();
  if (!name || name.length > 40) throw new Error('Enter your name (up to 40 characters).');
  if (!EMAIL_RE.test(email) || email.length > 80) throw new Error('Enter a valid email address.');
  if (!new RegExp('^\\d{' + PIN_MIN + ',' + PIN_MAX + '}$').test(newPin)) {
    throw new Error('PIN must be ' + PIN_MIN + ' to ' + PIN_MAX + ' digits.');
  }

  var settings = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var code = settings ? String(settings.getRange('B1').getDisplayValue()).trim() : '';
  if (!code) throw new Error('Registration is switched off. Ask your admin.');
  if (String(req.code || '').trim().toUpperCase() !== code.toUpperCase()) {
    recordFail_(keys);
    throw new Error('Incorrect registration code.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
    var rows = sheet.getDataRange().getDisplayValues();
    var claimRow = 0;
    var isAdmin = false;
    for (var i = 1; i < rows.length; i++) {
      var rowName = String(rows[i][0]).trim();
      var rowEmail = String(rows[i][1]).trim().toLowerCase();
      var rowPin = String(rows[i][2]).trim();
      if (rowEmail === email) {
        if (rowPin) throw new Error('This email is already registered. Ask your admin to reset your PIN.');
        if (rowName.toLowerCase() !== name.toLowerCase()) {
          throw new Error('The name does not match the one on file for this email.');
        }
        claimRow = i + 1;
        isAdmin = isTrue_(rows[i][4]);
        name = rowName; // keep the existing spelling so old timesheet rows still match
      }
    }
    if (claimRow) {
      sheet.getRange(claimRow, 3, 1, 2).setValues([[newPin, 'TRUE']]);
    } else {
      sheet.appendRow([name, email, newPin, 'TRUE', '']);
      sheet.getRange(sheet.getLastRow(), 2, 1, 2).setNumberFormat('@').setValues([[email, newPin]]);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return { name: name, email: email, admin: isAdmin };
}

function isTrue_(v) { return String(v).trim().toUpperCase() === 'TRUE'; }

// ---------- admin ----------

/**
 * Admin actions. `req.target` is the email of the user being changed (req.email is the signed-in admin).
 * Admins cannot reset, deactivate or delete themselves here, so the last admin can't lock themselves out.
 * Returns the refreshed { users, code } list.
 */
function admin_(admin, req) {
  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var ss = SpreadsheetApp.getActive();
    var sheet = ss.getSheetByName(USERS_SHEET);
    var target = String(req.target || '').trim().toLowerCase();
    var row;

    switch (req.action) {
      case 'admin_list':
        break;

      case 'admin_add':
        var name = String(req.name || '').replace(/\s+/g, ' ').trim();
        if (!name || name.length > 40) throw new Error('Enter a name (up to 40 characters).');
        if (!EMAIL_RE.test(target) || target.length > 80) throw new Error('Enter a valid email address.');
        if (userRow_(sheet, target)) throw new Error('That email is already in the list.');
        sheet.appendRow([name, target, '', 'TRUE', '']);
        sheet.getRange(sheet.getLastRow(), 2, 1, 2).setNumberFormat('@').setValues([[target, '']]);
        break;

      case 'admin_reset':   // clears the PIN; the person registers again (name + email + code) and picks a new one
        row = userRowFor_(sheet, target, admin);
        sheet.getRange(row, 3).setValue('');
        CacheService.getScriptCache().remove('fails_e_' + target); // lift any lockout too
        break;

      case 'admin_active':
        row = userRowFor_(sheet, target, admin);
        sheet.getRange(row, 4).setValue(req.active ? 'TRUE' : 'FALSE');
        break;

      case 'admin_delete':  // removes the login only; their timesheet rows stay in the Timesheet tab
        row = userRowFor_(sheet, target, admin);
        sheet.deleteRow(row);
        break;

      case 'admin_code':
        var code = String(req.code || '').trim();
        if (code && !/^[A-Za-z0-9]{4,20}$/.test(code)) throw new Error('Code must be 4 to 20 letters or digits.');
        var settings = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
        settings.getRange('A1').setValue('Registration code');
        settings.getRange('B1').setNumberFormat('@').setValue(code.toUpperCase());
        break;

      default:
        throw new Error('Unknown action.');
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return adminState_();
}

function adminState_() {
  var ss = SpreadsheetApp.getActive();
  var rows = ss.getSheetByName(USERS_SHEET).getDataRange().getDisplayValues();
  var users = [];
  for (var i = 1; i < rows.length; i++) {
    if (!String(rows[i][1]).trim()) continue;
    users.push({
      name: String(rows[i][0]).trim(),
      email: String(rows[i][1]).trim().toLowerCase(),
      hasPin: !!String(rows[i][2]).trim(),
      active: String(rows[i][3]).toUpperCase() !== 'FALSE',
      admin: isTrue_(rows[i][4])
    });
  }
  var settings = ss.getSheetByName(SETTINGS_SHEET);
  return { users: users, code: settings ? String(settings.getRange('B1').getDisplayValue()).trim() : '' };
}

/** Row number (1-based) of the user with this email, or 0. */
function userRow_(sheet, email) {
  var rows = sheet.getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    if (String(rows[i][1]).trim().toLowerCase() === email) return i + 1;
  }
  return 0;
}

/** Like userRow_ but throws if the user doesn't exist or is the admin themselves. */
function userRowFor_(sheet, email, admin) {
  if (email === admin.email) throw new Error('You cannot change your own account here.');
  var row = userRow_(sheet, email);
  if (!row) throw new Error('User not found.');
  return row;
}

/** Writes (or overwrites) the user's row for the given date. End and break are optional. */
function save_(user, req) {
  var date = checkDate_(req.date);
  var start = String(req.start || '');
  var end = String(req.end || '');
  var timeRe = /^([01]\d|2[0-3]):[0-5]\d$/;
  if (!timeRe.test(start)) throw new Error('Enter a valid start time.');
  if (end && !timeRe.test(end)) throw new Error('Enter a valid end time.');

  var brk = '';
  if (req.breakHrs !== '' && req.breakHrs != null) {
    brk = Number(req.breakHrs);
    if (isNaN(brk) || brk < 0 || brk > 24) throw new Error('Break must be between 0 and 24 hours.');
    brk = round2_(brk);
  }

  var total = total_(start, end);
  var inHrs = total === '' ? '' : Math.max(0, round2_(total - (brk || 0)));
  if (total !== '' && brk > total) throw new Error('Break is longer than the time worked.');

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = logSheet_();
    var row = findRow_(sheet, date, user.email);
    if (!row) {
      sheet.appendRow([user.name, user.email, date, dayName_(date), '', '', '', '', '']);
      row = sheet.getLastRow();
    }
    sheet.getRange(row, C.START, 1, 5).setValues([[start, end, brk, total, inHrs]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return entry_(user.email, date);
}

// ---------- helpers ----------

/**
 * Finds the user by email and checks the PIN. PINs need not be unique: the email picks the person.
 * Returns { name, email, admin }. A wrong email and a wrong PIN give the same error, so the app does not
 * reveal which emails exist.
 */
function authenticate_(email, pin, deviceId) {
  email = String(email || '').trim().toLowerCase();
  pin = String(pin || '').trim();
  var keys = [['fails_d_' + deviceId, MAX_FAILS], ['fails_e_' + email, MAX_FAILS]];
  checkLockout_(keys);

  var rows = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET).getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    var active = String(rows[i][3]).toUpperCase() !== 'FALSE';
    var rowPin = String(rows[i][2]).trim();
    if (email && active && rowPin && String(rows[i][1]).trim().toLowerCase() === email && rowPin === pin) {
      clearFails_(keys);
      return { name: String(rows[i][0]).trim(), email: email, admin: isTrue_(rows[i][4]) };
    }
  }
  recordFail_(keys);
  throw new Error('Incorrect email or PIN.');
}

// Lockout counters live in the script cache. keys is a list of [cacheKey, maxFails].
function checkLockout_(keys) {
  var cache = CacheService.getScriptCache();
  keys.forEach(function (k) {
    if (Number(cache.get(k[0]) || 0) >= k[1]) throw new Error('Too many wrong attempts. Try again in a few minutes.');
  });
}

function recordFail_(keys) {
  var cache = CacheService.getScriptCache();
  keys.forEach(function (k) {
    cache.put(k[0], String(Number(cache.get(k[0]) || 0) + 1), LOCKOUT_SECONDS);
  });
}

function clearFails_(keys) {
  CacheService.getScriptCache().removeAll(keys.map(function (k) { return k[0]; }));
}

function logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:F').setNumberFormat('@'); // keep text, dates and times as plain text
    styleHeader_(sheet.getRange(1, 1, 1, HEADERS.length));
  }
  return sheet;
}

function today_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/** Returns the date if it is a real yyyy-MM-dd date; otherwise throws. */
function checkDate_(date) {
  date = String(date || '');
  var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(date);
  if (!m) throw new Error('Pick a valid date.');
  var d = new Date(+m[1], +m[2] - 1, +m[3], 12);
  if (d.getMonth() !== +m[2] - 1) throw new Error('Pick a valid date.');
  return date;
}

function dayName_(date) {
  var p = date.split('-');
  return Utilities.formatDate(new Date(+p[0], +p[1] - 1, +p[2], 12), Session.getScriptTimeZone(), 'EEEE');
}

function findRow_(sheet, date, email) {
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][C.DATE - 1] === date && String(data[i][C.EMAIL - 1]).toLowerCase() === email) return i + 1;
  }
  return 0;
}

function entry_(email, date) {
  var sheet = logSheet_();
  var row = findRow_(sheet, date, email);
  var empty = { date: date, day: dayName_(date), start: '', end: '', brk: '', total: '', inHrs: '' };
  if (!row) return empty;
  var v = sheet.getRange(row, C.START, 1, 5).getDisplayValues()[0];
  empty.start = v[0]; empty.end = v[1]; empty.brk = v[2]; empty.total = v[3]; empty.inHrs = v[4];
  return empty;
}

function total_(start, end) {
  if (!start || !end) return '';
  var s = String(start).split(':'), e = String(end).split(':');
  var mins = (+e[0] * 60 + +e[1]) - (+s[0] * 60 + +s[1]);
  if (mins < 0) mins += 24 * 60; // shift crossing midnight
  return round2_(mins / 60);
}

function round2_(n) { return Math.round(n * 100) / 100; }
