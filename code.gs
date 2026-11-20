/**
 * NOVA Career Studio — Consultation form backend (Google Apps Script Web App)
 */

const SPREADSHEET_ID = ''; // <-- CONFIG PLACEHOLDER

const SHEET_NAME = 'Responses';
const HEADERS = ['Timestamp', 'Name', 'Email', 'Phone'];
const MAX_FIELD_LENGTH = 500;
const LOCK_TIMEOUT_MS = 30000;

// ─── ENTRY POINTS ───────────────────────────────────────────────────────────
function doGet(e) {
  return handleRequest_(e);
}

function doPost(e) {
  return handleRequest_(e);
}

// ─── CORE ───────────────────────────────────────────────────────────────────
function handleRequest_(e) {
  try {
    const data = readParams_(e);

    if (!data.name || !data.email || !data.phone) {
      throw new Error('Name, email and phone are all required.');
    }

    saveRow_(data);
    return jsonResponse_({ success: true });
  } catch (err) {
    return jsonResponse_({
      success: false,
      error: err && err.message ? err.message : String(err)
    });
  }
}

function readParams_(e) {
  const p = (e && e.parameter) || {};
  return {
    name: clean_(p.name),
    email: clean_(p.email),
    phone: clean_(p.phone)
  };
}

function clean_(value) {
  if (value === undefined || value === null) return '';
  return String(value).trim().slice(0, MAX_FIELD_LENGTH);
}

function getSpreadsheet_() {
  const ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  if (!ss) {
    throw new Error(
      'No spreadsheet found. Bind this script to a Google Sheet or set SPREADSHEET_ID.'
    );
  }
  return ss;
}

function getResponsesSheet_(ss) {
  let sheet = ss.getSheetByName(SHEET_NAME);

  // Auto-create the tab if it does not exist
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  // Write bold headers on a brand-new / empty tab
  if (sheet.getLastRow() === 0) {
    sheet
      .getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight('bold');
    sheet.setFrozenRows(1);
  }

  return sheet;
}

function saveRow_(data) {
  const lock = LockService.getScriptLock();
  lock.waitLock(LOCK_TIMEOUT_MS); // prevents two simultaneous submissions overwriting each other

  try {
    const ss = getSpreadsheet_();
    const sheet = getResponsesSheet_(ss);

    const timestamp = Utilities.formatDate(
      new Date(),
      ss.getSpreadsheetTimeZone(),
      'yyyy-MM-dd HH:mm:ss'
    );

    // Append as a new row. Plain-text format keeps phone numbers such as
    // "+91 98765 43210" intact and stops values starting with = + - @
    // from being interpreted as formulas.
    const row = sheet.getLastRow() + 1;
    const range = sheet.getRange(row, 1, 1, HEADERS.length);
    range.setNumberFormat('@');
    range.setValues([[timestamp, data.name, data.email, data.phone]]);

    SpreadsheetApp.flush();
  } finally {
    lock.releaseLock();
  }
}

function jsonResponse_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}