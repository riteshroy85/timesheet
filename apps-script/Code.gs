/**
 * Timesheet API (Google Apps Script). The phone screen is hosted separately (GitHub Pages)
 * and calls this web app with POST requests.
 *
 * Bound to a Google Sheet with two tabs:
 *   Users      : Name | PIN | Active (TRUE/FALSE)
 *   Timesheet  : Name | Date | Day | Start Time | End Time | Break (hrs) | Total Hrs | In-Hrs   (created automatically)
 * Break is in hours (e.g. 0.75). Total Hrs = End - Start. In-Hrs = Total Hrs - Break.
 */

var USERS_SHEET = 'Users';
var LOG_SHEET = 'Timesheet';
var SETTINGS_SHEET = 'Settings';
var MAX_FAILS = 5;           // wrong PINs allowed per device...
var MAX_FAILS_GLOBAL = 30;   // ...and across all devices, before locking out
var LOCKOUT_SECONDS = 600;

// Timesheet column numbers
var C = { NAME: 1, DATE: 2, DAY: 3, START: 4, END: 5, BREAK: 6, TOTAL: 7, IN: 8 };
var HEADERS = ['Name', 'Date', 'Day', 'Start Time', 'End Time', 'Break (hrs)', 'Total Hrs', 'In-Hrs'];

/**
 * Run once from the Apps Script editor (select "setup" > Run) on a blank Sheet.
 * Creates the Users and Timesheet tabs with styled headers and a sample user.
 * Safe to run again: it also refreshes the Timesheet header row.
 */
function setup() {
  var ss = SpreadsheetApp.getActive();
  var users = ss.getSheetByName(USERS_SHEET) || ss.insertSheet(USERS_SHEET);
  if (users.getLastRow() === 0) {
    users.appendRow(['Name', 'PIN', 'Active']);
    users.appendRow(['Sample User', '123456', 'TRUE']);
    users.getRange('B:B').setNumberFormat('@'); // keep leading zeros in PINs
    users.setFrozenRows(1);
  }
  styleHeader_(users.getRange(1, 1, 1, 3));

  // Settings: the registration code staff need to create their own PIN. Blank it to switch registration off.
  var settings = ss.getSheetByName(SETTINGS_SHEET) || ss.insertSheet(SETTINGS_SHEET);
  if (settings.getLastRow() === 0) {
    settings.appendRow(['Registration code', Math.random().toString(36).slice(2, 8).toUpperCase()]);
    settings.getRange('B1').setNumberFormat('@');
    settings.getRange('A1').setFontWeight('bold');
    settings.setColumnWidth(1, 160);
  }

  var log = logSheet_();
  log.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
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
 * Body (JSON sent as text/plain): { action, pin, deviceId, ... }
 *   register : name, code, newPin -> { name, entry }  (no pin needed; creates or claims a user)
 *   login : -> { name, entry }                      (entry for today)
 *   get   : date                -> entry
 *   save  : date, start, end, breakHrs -> entry
 */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var deviceId = String(req.deviceId || '').slice(0, 64);
    if (req.action === 'register') {
      var name = register_(req, deviceId);
      return json_({ ok: true, data: { name: name, entry: entry_(name, today_()) } });
    }
    var user = authenticate_(req.pin, deviceId);
    if (req.action === 'login') {
      out = { ok: true, data: { name: user.name, entry: entry_(user.name, today_()) } };
    } else if (req.action === 'get') {
      out = { ok: true, data: entry_(user.name, checkDate_(req.date)) };
    } else if (req.action === 'save') {
      out = { ok: true, data: save_(user.name, req) };
    } else {
      throw new Error('Unknown action.');
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return json_(out);
}

function json_(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * First-time sign-up: needs the registration code from the Settings tab.
 * Creates a new user, or claims an existing name whose PIN cell is blank (admin PIN reset).
 * Returns the saved name.
 */
function register_(req, deviceId) {
  checkLockout_(deviceId);

  var name = String(req.name || '').replace(/\s+/g, ' ').trim();
  var newPin = String(req.newPin || '').trim();
  if (!name || name.length > 40) throw new Error('Enter your name (up to 40 characters).');
  if (!/^\d{4,8}$/.test(newPin)) throw new Error('PIN must be 4 to 8 digits.');

  var settings = SpreadsheetApp.getActive().getSheetByName(SETTINGS_SHEET);
  var code = settings ? String(settings.getRange('B1').getDisplayValue()).trim() : '';
  if (!code) throw new Error('Registration is switched off. Ask your admin.');
  if (String(req.code || '').trim().toUpperCase() !== code.toUpperCase()) {
    recordFail_(deviceId);
    throw new Error('Incorrect registration code.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET);
    var rows = sheet.getDataRange().getDisplayValues();
    var claimRow = 0;
    for (var i = 1; i < rows.length; i++) {
      var rowName = String(rows[i][0]).trim();
      var rowPin = String(rows[i][1]).trim();
      if (rowPin === newPin) throw new Error('That PIN is already in use. Choose a different one.');
      if (rowName.toLowerCase() === name.toLowerCase()) {
        if (rowPin) throw new Error('That name is already registered. Ask your admin to reset your PIN.');
        claimRow = i + 1;
        name = rowName; // keep the existing spelling so old timesheet rows still match
      }
    }
    if (claimRow) {
      sheet.getRange(claimRow, 2, 1, 2).setValues([[newPin, 'TRUE']]);
    } else {
      sheet.appendRow([name, newPin, 'TRUE']);
      sheet.getRange(sheet.getLastRow(), 2).setNumberFormat('@').setValue(newPin);
    }
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return name;
}

/** Writes (or overwrites) the user's row for the given date. End and break are optional. */
function save_(name, req) {
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
    var row = findRow_(sheet, date, name);
    if (!row) {
      sheet.appendRow([name, date, dayName_(date), '', '', '', '', '']);
      row = sheet.getLastRow();
    }
    sheet.getRange(row, C.START, 1, 5).setValues([[start, end, brk, total, inHrs]]);
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return entry_(name, date);
}

// ---------- helpers ----------

function authenticate_(pin, deviceId) {
  pin = String(pin || '').trim();
  checkLockout_(deviceId);

  var rows = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET).getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    var active = String(rows[i][2]).toUpperCase() !== 'FALSE';
    if (active && pin && String(rows[i][1]).trim() === pin) {
      CacheService.getScriptCache().remove('fails_' + deviceId);
      return { name: String(rows[i][0]).trim() };
    }
  }
  recordFail_(deviceId);
  throw new Error('Incorrect PIN.');
}

// Wrong PINs and wrong registration codes share the same lockout counters.
function checkLockout_(deviceId) {
  var cache = CacheService.getScriptCache();
  if (Number(cache.get('fails_' + deviceId) || 0) >= MAX_FAILS || Number(cache.get('fails_all') || 0) >= MAX_FAILS_GLOBAL) {
    throw new Error('Too many wrong attempts. Try again in a few minutes.');
  }
}

function recordFail_(deviceId) {
  var cache = CacheService.getScriptCache();
  cache.put('fails_' + deviceId, String(Number(cache.get('fails_' + deviceId) || 0) + 1), LOCKOUT_SECONDS);
  cache.put('fails_all', String(Number(cache.get('fails_all') || 0) + 1), LOCKOUT_SECONDS);
}

function logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
    sheet.getRange('A:E').setNumberFormat('@'); // keep text, dates and times as plain text
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

function findRow_(sheet, date, name) {
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][C.DATE - 1] === date && data[i][C.NAME - 1] === name) return i + 1;
  }
  return 0;
}

function entry_(name, date) {
  var row = findRow_(logSheet_(), date, name);
  var empty = { date: date, day: dayName_(date), start: '', end: '', brk: '', total: '', inHrs: '' };
  if (!row) return empty;
  var v = logSheet_().getRange(row, C.START, 1, 5).getDisplayValues()[0];
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
