/**
 * Timesheet API (Google Apps Script). The phone screen is hosted separately (GitHub Pages)
 * and calls this web app with POST requests.
 *
 * Bound to a Google Sheet with two tabs:
 *   Users      : Name | PIN | Active (TRUE/FALSE)
 *   Timesheet  : Name | Date | Day | Start Time | End Time | Break | Total Hrs | In-Hrs   (created automatically)
 * Break is in minutes. Total Hrs = End - Start. In-Hrs = Total Hrs - Break.
 */

var USERS_SHEET = 'Users';
var LOG_SHEET = 'Timesheet';
var MAX_FAILS = 5;           // wrong PINs allowed per device...
var MAX_FAILS_GLOBAL = 30;   // ...and across all devices, before locking out
var LOCKOUT_SECONDS = 600;

// Timesheet column numbers
var C = { NAME: 1, DATE: 2, DAY: 3, START: 4, END: 5, BREAK: 6, TOTAL: 7, IN: 8 };

/**
 * Run once from the Apps Script editor (select "setup" > Run) on a blank Sheet.
 * Creates the Users and Timesheet tabs with styled headers and a sample user.
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
  styleHeader_(logSheet_().getRange(1, 1, 1, 8));

  var blank = ss.getSheetByName('Sheet1');
  if (blank && blank.getLastRow() === 0) ss.deleteSheet(blank);
}

function styleHeader_(range) {
  range.setBackground('#1f5c80').setFontColor('#ffffff').setFontWeight('bold');
}

function doGet() {
  return ContentService.createTextOutput('Timesheet API is running.');
}

/** Body (JSON sent as text/plain): { action: 'login'|'record', pin, deviceId, kind?, value? } */
function doPost(e) {
  var out;
  try {
    var req = JSON.parse(e.postData.contents);
    var deviceId = String(req.deviceId || '').slice(0, 64);
    if (req.action === 'login') {
      var user = authenticate_(req.pin, deviceId);
      out = { ok: true, data: { name: user.name, today: todayEntry_(user.name) } };
    } else if (req.action === 'record') {
      out = { ok: true, data: record_(req.pin, deviceId, req.kind, req.value) };
    } else {
      throw new Error('Unknown action.');
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return ContentService.createTextOutput(JSON.stringify(out)).setMimeType(ContentService.MimeType.JSON);
}

/**
 * kind 'start' | 'end': value is 'HH:mm'.
 * kind 'break': value is whole minutes (0-600).
 */
function record_(pin, deviceId, kind, value) {
  var user = authenticate_(pin, deviceId);
  var col;
  if (kind === 'start' || kind === 'end') {
    if (!/^([01]\d|2[0-3]):[0-5]\d$/.test(value)) throw new Error('Invalid time.');
    col = kind === 'start' ? C.START : C.END;
  } else if (kind === 'break') {
    value = Number(value);
    if (!(value >= 0 && value <= 600) || value % 1 !== 0) throw new Error('Invalid break.');
    col = C.BREAK;
  } else {
    throw new Error('Invalid action.');
  }

  var lock = LockService.getScriptLock();
  lock.waitLock(10000);
  try {
    var sheet = logSheet_();
    var now = new Date();
    var tz = Session.getScriptTimeZone();
    var date = Utilities.formatDate(now, tz, 'yyyy-MM-dd');
    var row = findRow_(sheet, date, user.name);
    if (!row) {
      sheet.appendRow([user.name, date, Utilities.formatDate(now, tz, 'EEEE'), '', '', '', '', '']);
      row = sheet.getLastRow();
    }
    sheet.getRange(row, col).setValue(value);

    var v = sheet.getRange(row, C.START, 1, 3).getDisplayValues()[0]; // start, end, break
    var total = total_(v[0], v[1]);
    sheet.getRange(row, C.TOTAL).setValue(total);
    sheet.getRange(row, C.IN).setValue(total === '' ? '' : Math.max(0, round2_(total - (Number(v[2]) || 0) / 60)));
    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
  return todayEntry_(user.name);
}

// ---------- helpers ----------

function authenticate_(pin, deviceId) {
  pin = String(pin || '').trim();
  var cache = CacheService.getScriptCache();
  var devKey = 'fails_' + deviceId;
  var devFails = Number(cache.get(devKey) || 0);
  var globalFails = Number(cache.get('fails_all') || 0);
  if (devFails >= MAX_FAILS || globalFails >= MAX_FAILS_GLOBAL) {
    throw new Error('Too many wrong attempts. Try again in a few minutes.');
  }

  var rows = SpreadsheetApp.getActive().getSheetByName(USERS_SHEET).getDataRange().getDisplayValues();
  for (var i = 1; i < rows.length; i++) {
    var active = String(rows[i][2]).toUpperCase() !== 'FALSE';
    if (active && pin && String(rows[i][1]).trim() === pin) {
      cache.remove(devKey);
      return { name: String(rows[i][0]).trim() };
    }
  }
  cache.put(devKey, String(devFails + 1), LOCKOUT_SECONDS);
  cache.put('fails_all', String(globalFails + 1), LOCKOUT_SECONDS);
  throw new Error('Incorrect PIN.');
}

function logSheet_() {
  var ss = SpreadsheetApp.getActive();
  var sheet = ss.getSheetByName(LOG_SHEET);
  if (!sheet) {
    sheet = ss.insertSheet(LOG_SHEET);
    sheet.appendRow(['Name', 'Date', 'Day', 'Start Time', 'End Time', 'Break (min)', 'Total Hrs', 'In-Hrs']);
    sheet.setFrozenRows(1);
    sheet.getRange('A:E').setNumberFormat('@'); // keep text, dates and times as plain text
    styleHeader_(sheet.getRange(1, 1, 1, 8));
  }
  return sheet;
}

function findRow_(sheet, date, name) {
  var data = sheet.getDataRange().getDisplayValues();
  for (var i = data.length - 1; i >= 1; i--) {
    if (data[i][C.DATE - 1] === date && data[i][C.NAME - 1] === name) return i + 1;
  }
  return 0;
}

function todayEntry_(name) {
  var sheet = logSheet_();
  var tz = Session.getScriptTimeZone();
  var row = findRow_(sheet, Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd'), name);
  if (!row) return { start: '', end: '', brk: '', total: '', inHrs: '' };
  var v = sheet.getRange(row, C.START, 1, 5).getDisplayValues()[0];
  return { start: v[0], end: v[1], brk: v[2], total: v[3], inHrs: v[4] };
}

function total_(start, end) {
  if (!start || !end) return '';
  var s = String(start).split(':'), e = String(end).split(':');
  var mins = (+e[0] * 60 + +e[1]) - (+s[0] * 60 + +s[1]);
  if (mins < 0) mins += 24 * 60; // shift crossing midnight
  return round2_(mins / 60);
}

function round2_(n) { return Math.round(n * 100) / 100; }
