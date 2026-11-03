/**
 * Fantasy Lighting time card: the small program that lives inside the Google Sheet.
 * The iPad app sends its requests here, and this reads and writes the Hours tab.
 * This file holds no secrets. The access code is created by setup() and kept in Google's script settings.
 *
 * @OnlyCurrentDoc
 */

const HOURS_TAB = 'Hours';
const PEOPLE_TAB = 'People';
const SETUP_TAB = 'Setup';
const PAYROLL_TAB = 'Payroll';
const PAYROLL_FIRST_ROW = 6;
const HOURS_HEADERS = ['Name', 'Date', 'Start', 'End', 'Break (minutes)', 'Hours', 'Logged at', 'Entry ID'];
const COL = { name: 1, date: 2, start: 3, end: 4, breakMin: 5, hours: 6, loggedAt: 7, id: 8 };
// Written as a formula so the Hours column stays right if someone edits a row in the Sheet.
const HOURS_FORMULA = '=ROUND((MOD(TIMEVALUE(RC[-2])-TIMEVALUE(RC[-3]),1)*1440-RC[-1])/60,2)';
const MAX_PIN_TRIES = 5;
const PIN_LOCK_SECONDS = 600;

// Run this once from the script editor. Running it again keeps all hours and the same access code.
function setup() {
  const book = SpreadsheetApp.getActive();
  const first = book.getSheets()[0];
  if (first.getName() === 'Sheet1' && first.getLastRow() === 0) first.setName(HOURS_TAB);

  const hours = book.getSheetByName(HOURS_TAB) || book.insertSheet(HOURS_TAB);
  if (hours.getLastRow() === 0) {
    hours.getRange(1, 1, 1, HOURS_HEADERS.length).setValues([HOURS_HEADERS]).setFontWeight('bold');
    hours.setFrozenRows(1);
    hours.hideColumns(COL.id);
  }

  const people = book.getSheetByName(PEOPLE_TAB) || book.insertSheet(PEOPLE_TAB);
  if (people.getLastRow() === 0) {
    people.getRange(1, 1, 1, 2).setValues([['Name', 'PIN (4 digits)']]).setFontWeight('bold');
    people.setFrozenRows(1);
  }
  // Plain text keeps a PIN like 0042 from turning into the number 42.
  people.getRange('B:B').setNumberFormat('@');

  const props = PropertiesService.getScriptProperties();
  let code = props.getProperty('ACCESS_CODE');
  if (!code) {
    const random = Utilities.getUuid().replace(/-/g, '').toUpperCase().slice(0, 12);
    code = random.slice(0, 4) + '-' + random.slice(4, 8) + '-' + random.slice(8, 12);
    props.setProperty('ACCESS_CODE', code);
  }

  const notes = book.getSheetByName(SETUP_TAB) || book.insertSheet(SETUP_TAB);
  notes.clear();
  notes.getRange(1, 1, 3, 2).setValues([
    ['Access code', code],
    ['What it is', 'Type this into the iPad once, to connect it to this Sheet. Keep it private.'],
    ['Names and PINs', 'Add each person on the People tab, with a 4-digit PIN in the second column.']
  ]);
  Logger.log('Setup finished. Access code: ' + code);
}

function doGet() {
  return ContentService.createTextOutput('The Fantasy Lighting time card is running.');
}

function doPost(e) {
  let result;
  const lock = LockService.getScriptLock();
  try {
    const request = JSON.parse(e.postData.contents);
    lock.waitLock(20000);
    result = handleRequest(request);
  } catch (err) {
    if (!err.apiCode) console.error(err);
    result = { ok: false, error: err.apiCode || 'server' };
  } finally {
    lock.releaseLock();
  }
  return ContentService.createTextOutput(JSON.stringify(result)).setMimeType(ContentService.MimeType.JSON);
}

function apiError(code) {
  const err = new Error(code);
  err.apiCode = code;
  return err;
}

function handleRequest(request) {
  const expected = PropertiesService.getScriptProperties().getProperty('ACCESS_CODE');
  if (!expected || String(request.code || '').trim().toUpperCase() !== expected) throw apiError('access');
  if (request.action === 'save') saveEntry(request.entry || {}, request.pin);
  else if (request.action === 'delete') deleteEntry(request.id, request.pin);
  else if (request.action !== 'load') throw apiError('invalid');
  if (request.action !== 'load') refreshPayrollSafely();
  return loadAll();
}

// A payroll problem must never stop someone's hours from saving.
function refreshPayrollSafely() {
  try {
    updatePayroll();
  } catch (err) {
    console.error(err);
  }
}

// ---- Reading ----
function loadAll() {
  const people = readPeople().map(p => p.name);
  const sheet = hoursSheet();
  const entries = [];
  const badRows = [];
  const lastRow = sheet.getLastRow();
  if (lastRow >= 2) {
    // Display values are the text as shown in the cell, so Google never hands back date objects.
    const rows = sheet.getRange(2, 1, lastRow - 1, HOURS_HEADERS.length).getDisplayValues();
    rows.forEach((row, i) => {
      if (isBlankRow(row)) return;
      const rowNumber = i + 2;
      const entry = rowToEntry(row);
      if (!isValidEntry(entry)) {
        badRows.push(rowNumber);
        return;
      }
      // A row typed straight into the Sheet has no ID yet, and the app needs one to edit it.
      if (!entry.id) {
        entry.id = Utilities.getUuid();
        sheet.getRange(rowNumber, COL.id).setValue(entry.id);
      }
      entries.push(entry);
    });
  }
  return { ok: true, people, entries, badRows };
}

function isBlankRow(row) {
  return row.slice(0, 5).every(cell => String(cell).trim() === '');
}

function rowToEntry(row) {
  return {
    id: String(row[COL.id - 1]).trim(),
    name: String(row[COL.name - 1]).trim(),
    date: normalizeDate(row[COL.date - 1]),
    start: normalizeTime(row[COL.start - 1]),
    end: normalizeTime(row[COL.end - 1]),
    breakMin: String(row[COL.breakMin - 1]).trim() === '' ? 0 : Number(row[COL.breakMin - 1])
  };
}

// Reads only the name and PIN columns, so pay rates are never sent to the iPad.
function readPeople() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(PEOPLE_TAB);
  if (!sheet) throw apiError('setup');
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  return sheet.getRange(2, 1, lastRow - 1, 2).getDisplayValues()
    .map(row => ({ name: String(row[0]).trim(), pin: String(row[1]).trim() }))
    .filter(person => person.name);
}

function hoursSheet() {
  const sheet = SpreadsheetApp.getActive().getSheetByName(HOURS_TAB);
  if (!sheet) throw apiError('setup');
  return sheet;
}

function normalizeDate(value) {
  const text = String(value).trim();
  if (/^\d{4}-\d{2}-\d{2}$/.test(text)) return text;
  const parsed = new Date(text);
  if (!text || isNaN(parsed)) return '';
  return toIsoDate(parsed);
}

function toIsoDate(d) {
  return d.getFullYear() + '-' + pad(d.getMonth() + 1) + '-' + pad(d.getDate());
}

// Accepts how a person might type a time in the Sheet: 8:00 AM, 8:00, 17:30, 17:30:00.
function normalizeTime(value) {
  const match = String(value).trim().match(/^(\d{1,2}):(\d{2})(?::\d{2})?\s*([ap])?\.?m?\.?$/i);
  if (!match) return '';
  let hour = Number(match[1]);
  const minute = Number(match[2]);
  const meridiem = (match[3] || '').toLowerCase();
  if (meridiem && (hour < 1 || hour > 12)) return '';
  if (meridiem === 'p' && hour < 12) hour += 12;
  if (meridiem === 'a' && hour === 12) hour = 0;
  if (hour > 23 || minute > 59) return '';
  return pad(hour) + ':' + pad(minute);
}

function pad(n) {
  return String(n).padStart(2, '0');
}

function toMinutes(time) {
  const parts = time.split(':').map(Number);
  return parts[0] * 60 + parts[1];
}

function isValidEntry(entry) {
  if (!entry.name || !/^\d{4}-\d{2}-\d{2}$/.test(entry.date) || !entry.start || !entry.end) return false;
  if (entry.start === entry.end || !Number.isInteger(entry.breakMin) || entry.breakMin < 0) return false;
  // An end time earlier than the start time means the shift ran past midnight.
  const span = toMinutes(entry.end) - toMinutes(entry.start);
  return (span < 0 ? span + 1440 : span) > entry.breakMin;
}

// ---- PINs ----
function checkPin(name, pin) {
  const wanted = String(name || '').trim().toLowerCase();
  const person = readPeople().find(p => p.name.toLowerCase() === wanted);
  if (!person) throw apiError('no-person');
  const cache = CacheService.getScriptCache();
  const key = 'pin-tries-' + Utilities.base64Encode(wanted).slice(0, 200);
  const tries = Number(cache.get(key) || 0);
  if (tries >= MAX_PIN_TRIES) throw apiError('locked');
  if (!/^\d{4}$/.test(person.pin) || String(pin || '') !== person.pin) {
    cache.put(key, String(tries + 1), PIN_LOCK_SECONDS);
    throw apiError(tries + 1 >= MAX_PIN_TRIES ? 'locked' : 'pin');
  }
  cache.remove(key);
  return person;
}

// ---- Writing ----
function findRow(sheet, id) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2 || !id) return 0;
  const ids = sheet.getRange(2, COL.id, lastRow - 1, 1).getDisplayValues();
  const index = ids.findIndex(row => row[0] === id);
  return index < 0 ? 0 : index + 2;
}

function to12Hour(time) {
  const parts = time.split(':').map(Number);
  return (parts[0] % 12 || 12) + ':' + pad(parts[1]) + (parts[0] < 12 ? ' AM' : ' PM');
}

// Adds a new row, or updates the row with this ID. Sending the same entry twice after
// a dropped connection therefore can't create a duplicate.
function saveEntry(raw, pin) {
  const entry = {
    id: String(raw.id || ''),
    name: String(raw.name || '').trim(),
    date: normalizeDate(raw.date),
    start: normalizeTime(raw.start),
    end: normalizeTime(raw.end),
    breakMin: Number(raw.breakMin)
  };
  if (!/^[\w-]{1,64}$/.test(entry.id) || !isValidEntry(entry)) throw apiError('invalid');
  const person = checkPin(entry.name, pin);

  const sheet = hoursSheet();
  let row = findRow(sheet, entry.id);
  if (row) {
    const owner = String(sheet.getRange(row, COL.name).getDisplayValue()).trim().toLowerCase();
    if (owner !== person.name.toLowerCase()) throw apiError('invalid');
  } else {
    row = sheet.getLastRow() + 1;
    if (row > sheet.getMaxRows()) sheet.insertRowAfter(sheet.getMaxRows());
    sheet.getRange(row, COL.loggedAt).setValue(new Date());
    sheet.getRange(row, COL.id).setValue(entry.id);
  }
  // Plain text stops Google changing the date or times into its own format.
  sheet.getRange(row, COL.name, 1, 4).setNumberFormat('@')
    .setValues([[person.name, entry.date, to12Hour(entry.start), to12Hour(entry.end)]]);
  sheet.getRange(row, COL.breakMin).setValue(entry.breakMin);
  sheet.getRange(row, COL.hours).setFormulaR1C1(HOURS_FORMULA);
}

function deleteEntry(id, pin) {
  const sheet = hoursSheet();
  const row = findRow(sheet, String(id || ''));
  if (!row) return;
  checkPin(sheet.getRange(row, COL.name).getDisplayValue(), pin);
  sheet.deleteRow(row);
}

// ---- Payroll: hours times pay rate, before taxes, with no overtime. Only used inside the Sheet. ----
// Payroll stays current by itself: when hours are saved or deleted on the iPad, when the Sheet opens,
// and when someone edits the Hours tab, the People tab, or the week in cell B1.
function onOpen() {
  refreshPayrollSafely();
}

function onEdit(e) {
  const tab = e.range.getSheet().getName();
  const editedWeek = tab === PAYROLL_TAB && e.range.getRow() === 1 && e.range.getColumn() === 2;
  if (tab === HOURS_TAB || tab === PEOPLE_TAB || editedWeek) refreshPayrollSafely();
}

// Run this once from the script editor. Running it again keeps the pay rates and goes back to showing this week.
function setupPayroll() {
  const book = SpreadsheetApp.getActive();
  const people = book.getSheetByName(PEOPLE_TAB);
  if (!people || !book.getSheetByName(HOURS_TAB)) throw new Error('Run setup first.');
  people.getRange(1, 3).setValue('Pay rate ($ per hour)').setFontWeight('bold');

  const payroll = book.getSheetByName(PAYROLL_TAB) || book.insertSheet(PAYROLL_TAB);
  payroll.getRange(1, 2).setNumberFormat('@').setValue('');
  payroll.getRange(1, 1, 4, 1).setValues([['Other week (blank = this week)'], ['Showing'], ['Last updated'], ['Problems']]).setFontWeight('bold');
  payroll.getRange(PAYROLL_FIRST_ROW - 1, 1, 1, 4)
    .setValues([['Name', 'Hours', 'Pay rate', 'Pay before taxes']]).setFontWeight('bold');
  payroll.setColumnWidth(1, 230);
  updatePayroll();
}

function parseRate(value) {
  const rate = Number(String(value).replace(/[$,\s]/g, ''));
  return String(value).trim() !== '' && isFinite(rate) && rate > 0 ? rate : null;
}

function updatePayroll() {
  const book = SpreadsheetApp.getActive();
  const payroll = book.getSheetByName(PAYROLL_TAB);
  const people = book.getSheetByName(PEOPLE_TAB);
  const hours = book.getSheetByName(HOURS_TAB);
  if (!payroll || !people || !hours) throw new Error('Run setupPayroll first.');

  const lastOld = payroll.getLastRow();
  if (lastOld >= PAYROLL_FIRST_ROW) payroll.getRange(PAYROLL_FIRST_ROW, 1, lastOld - PAYROLL_FIRST_ROW + 1, 4).clearContent();
  payroll.getRange(3, 2).setNumberFormat('m/d/yyyy h:mm am/pm').setValue(new Date());

  // Weeks run Sunday to Saturday. A blank B1 follows the current week by itself.
  const typedWeek = String(payroll.getRange(1, 2).getDisplayValue()).trim();
  const today = new Date();
  const weekStart = typedWeek
    ? normalizeDate(typedWeek)
    : toIsoDate(new Date(today.getFullYear(), today.getMonth(), today.getDate() - today.getDay()));
  if (!weekStart) {
    payroll.getRange(2, 2).setValue('');
    payroll.getRange(4, 2).setValue('Cell B1 needs a date, like 11/8/2026, or leave it blank for this week. Nothing was calculated.');
    return;
  }
  const parts = weekStart.split('-').map(Number);
  const weekEnd = toIsoDate(new Date(parts[0], parts[1] - 1, parts[2] + 6));
  payroll.getRange(2, 2).setNumberFormat('@').setValue(weekStart + ' to ' + weekEnd);

  // One line per person, grouped ignoring capitals so a retyped name still adds up to the same person.
  const lines = {};
  const lineFor = name => {
    const key = name.toLowerCase();
    if (!lines[key]) lines[key] = { name, minutes: 0, rate: null };
    return lines[key];
  };
  const lastPerson = people.getLastRow();
  if (lastPerson >= 2) {
    people.getRange(2, 1, lastPerson - 1, 3).getDisplayValues().forEach(row => {
      const name = String(row[0]).trim();
      if (name) lineFor(name).rate = parseRate(row[2]);
    });
  }

  const problems = [];
  const badRows = [];
  const lastHour = hours.getLastRow();
  if (lastHour >= 2) {
    hours.getRange(2, 1, lastHour - 1, HOURS_HEADERS.length).getDisplayValues().forEach((row, i) => {
      if (isBlankRow(row)) return;
      const entry = rowToEntry(row);
      if (!isValidEntry(entry)) {
        badRows.push(i + 2);
      } else if (entry.date >= weekStart && entry.date <= weekEnd) {
        const span = toMinutes(entry.end) - toMinutes(entry.start);
        lineFor(entry.name).minutes += (span < 0 ? span + 1440 : span) - entry.breakMin;
      }
    });
  }
  if (badRows.length) problems.push('Hours tab row ' + badRows.join(', ') + " can't be read and is not counted.");

  let totalMinutes = 0;
  let totalPay = 0;
  const rows = Object.keys(lines).map(key => lines[key])
    .sort((a, b) => a.name.localeCompare(b.name))
    .map(line => {
      const pay = line.rate === null ? null : Math.round(line.minutes / 60 * line.rate * 100) / 100;
      totalMinutes += line.minutes;
      if (pay !== null) totalPay += pay;
      if (pay === null && line.minutes > 0) problems.push(line.name + ' has hours but no pay rate, so their pay is not in the total.');
      return [line.name, Math.round(line.minutes / 60 * 100) / 100, line.rate === null ? '' : line.rate, pay === null ? (line.minutes > 0 ? 'Needs pay rate' : '') : pay];
    });
  rows.push(['Total', Math.round(totalMinutes / 60 * 100) / 100, '', Math.round(totalPay * 100) / 100]);

  const range = payroll.getRange(PAYROLL_FIRST_ROW, 1, rows.length, 4);
  range.setValues(rows).setFontWeight('normal');
  payroll.getRange(PAYROLL_FIRST_ROW, 2, rows.length, 1).setNumberFormat('0.00');
  payroll.getRange(PAYROLL_FIRST_ROW, 3, rows.length, 2).setNumberFormat('$#,##0.00');
  payroll.getRange(PAYROLL_FIRST_ROW + rows.length - 1, 1, 1, 4).setFontWeight('bold');
  payroll.getRange(4, 2).setValue(problems.length ? problems.join(' ') : 'None');
}
