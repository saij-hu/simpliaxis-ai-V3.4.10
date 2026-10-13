/**
 * Google Apps Script Web App
 * Sales vs Target / Target Management sync
 *
 * Spreadsheet columns: Month | CRE | Source | Sales Target, plus RevenueTarget
 * in the next available column. Existing columns are preserved.
 * GET returns { success: true, salesVsTarget: [...] }
 * POST actions: setSalesTarget, setRevenueTarget, setTargets, resetAllTargets, recompute
 */

const SPREADSHEET_ID = '1NXVBkxrRzCKCR0VF6yT9b2m79pObKxAulJAsxCruQQM';
const SALES_VS_TARGET_SHEET = 'SalesVsTarget';

function doGet() {
  try {
    return respond({ success: true, salesVsTarget: readSalesVsTargetRows() });
  } catch (error) {
    console.error('SalesVsTarget GET failed:', error && error.stack ? error.stack : error);
    return respond({ success: false, salesVsTarget: [], error: errorMessage(error) });
  }
}

function doPost(e) {
  try {
    if (!e || !e.postData || typeof e.postData.contents !== 'string') {
      throw new Error('Missing POST body. Send a JSON request body.');
    }

    const body = JSON.parse(e.postData.contents);
    const action = String(body.action || '').trim();
    let savedTarget = null;

    if (action === 'setSalesTarget') {
      savedTarget = withScriptLock(function () {
        return upsertSalesVsTarget(body.month, body.cre, body.source, body.salesTarget);
      });
    } else if (action === 'setTargets') {
      savedTarget = withScriptLock(function () {
        const sales = upsertSalesVsTarget(body.month, body.cre, body.source, body.salesTarget);
        const revenue = upsertRevenueTarget(body.month, body.cre, body.source, body.revenueTarget);
        return {
          month: sales.month,
          monthKey: sales.monthKey,
          cre: sales.cre,
          source: sales.source,
          salesTarget: sales.salesTarget,
          revenueTarget: revenue.revenueTarget,
          rowNumber: sales.rowNumber
        };
      });
    } else if (action === 'setRevenueTarget') {
      savedTarget = withScriptLock(function () {
        return upsertRevenueTarget(body.month, body.cre, body.source, body.revenueTarget);
      });
    } else if (action === 'resetAllTargets') {
      withScriptLock(function () {
        resetAllTargets(body.month);
      });
    } else if (action === 'recompute') {
      // Existing dashboard computes derived achievement metrics client-side.
    } else {
      throw new Error('Unknown action: ' + (action || '(missing)'));
    }

    SpreadsheetApp.flush();
    const rows = readSalesVsTargetRows();

    // Return and verify the exact written target field before the frontend displays success.
    if (action === 'setSalesTarget' || action === 'setRevenueTarget' || action === 'setTargets') {
      const expected = savedTarget;
      const verified = rows.find(function (row) {
        return normalizeMonth(row.month) === expected.monthKey &&
          normalizeText(row.cre) === normalizeText(expected.cre) &&
          normalizeText(row.source) === normalizeText(expected.source);
      });

      if (!verified) {
        throw new Error(
          'Write completed but verification could not find the saved row for ' +
          expected.cre + ' / ' + expected.monthKey + ' / source "' + expected.source + '".'
        );
      }
      const fields = action === 'setTargets' ? ['salesTarget', 'revenueTarget'] :
        [action === 'setRevenueTarget' ? 'revenueTarget' : 'salesTarget'];
      fields.forEach(function (field) {
        if (Math.abs(Number(verified[field]) - Number(expected[field])) > 0.000001) {
          throw new Error(
            (field === 'revenueTarget' ? 'Revenue target' : 'Sales target') +
            ' verification mismatch for ' + expected.cre + ' / ' + expected.monthKey +
            '. Expected ' + expected[field] + ', read back ' + verified[field] + '.'
          );
        }
      });
      savedTarget = Object.assign({}, expected, { rowNumber: expected.rowNumber });
    }

    return respond({
      success: true,
      action: action,
      saved: savedTarget,
      salesVsTarget: rows
    });
  } catch (error) {
    console.error('SalesVsTarget POST failed:', error && error.stack ? error.stack : error);
    return respond({ success: false, salesVsTarget: [], error: errorMessage(error) });
  }
}

function respond(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function errorMessage(error) {
  return error && error.message ? String(error.message) : String(error);
}

/** Find the RevenueTarget column or create it at the next available column.
 * Existing A:O columns are preserved. This is backward-compatible with older sheets.
 */
function getRevenueTargetColumn_(sheet) {
  const lastColumn = Math.max(4, sheet.getLastColumn());
  const headers = sheet.getRange(1, 1, 1, lastColumn).getDisplayValues()[0];
  for (let i = 0; i < headers.length; i++) {
    const key = normalizeText(headers[i]).replace(/[ _-]/g, '');
    if (key === 'revenuetarget' || key === 'monthlyrevenuetarget') return i + 1;
  }
  const newColumn = lastColumn + 1;
  sheet.getRange(1, newColumn).setValue('RevenueTarget');
  return newColumn;
}

function withScriptLock(callback) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    return callback();
  } finally {
    lock.releaseLock();
  }
}

function getSpreadsheet() {
  const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
  console.log('Connected Spreadsheet: ' + ss.getName() + ' (' + ss.getId() + ')');
  return ss;
}

function getSalesVsTargetSheet() {
  const sheet = getSpreadsheet().getSheetByName(SALES_VS_TARGET_SHEET);
  if (!sheet) {
    throw new Error('Sheet "' + SALES_VS_TARGET_SHEET + '" was not found. Check the tab name.');
  }
  return sheet;
}

/**
 * Convert a month value from the browser or Google Sheets into yyyy-MM.
 * Google Sheets may return an actual Date for a date-formatted cell;
 * String(date) is not comparable to a browser value such as "2026-10".
 */
function normalizeMonth(value, timeZone) {
  if (value === null || value === undefined || value === '') return '';
  const tz = timeZone || getSpreadsheetTimeZone_();

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, tz, 'yyyy-MM');
  }

  const text = String(value).trim();
  if (!text) return '';

  // ISO month or ISO date: 2026-10, 2026-10-01
  let match = text.match(/^(\d{4})[-/](\d{1,2})(?:[-/]\d{1,2})?$/);
  if (match) {
    const month = Number(match[2]);
    if (month >= 1 && month <= 12) return match[1] + '-' + String(month).padStart(2, '0');
  }

  // Slash-formatted date text; default to India-style DD/MM/YYYY.
  match = text.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})$/);
  if (match) {
    const day = Number(match[1]);
    const month = Number(match[2]);
    if (day >= 1 && day <= 31 && month >= 1 && month <= 12) {
      return match[3] + '-' + String(month).padStart(2, '0');
    }
  }

  // Text month names, e.g. "Oct 2026"; interpret by JavaScript date parser.
  const parsed = new Date(text);
  if (!isNaN(parsed.getTime())) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM');
  }

  // Preserve an unknown value so it can be reported/diagnosed rather than hidden.
  return text;
}

function getSpreadsheetTimeZone_() {
  try {
    return getSpreadsheet().getSpreadsheetTimeZone() || Session.getScriptTimeZone() || 'Etc/UTC';
  } catch (error) {
    return Session.getScriptTimeZone() || 'Etc/UTC';
  }
}

function normalizeText(value) {
  return String(value === null || value === undefined ? '' : value)
    .normalize('NFKC')
    .trim()
    .replace(/\s+/g, ' ')
    .toLocaleLowerCase();
}

function displayText(value) {
  return String(value === null || value === undefined ? '' : value)
    .normalize('NFKC')
    .trim()
    .replace(/\s+/g, ' ');
}

function readSalesVsTargetRows() {
  const sheet = getSalesVsTargetSheet();
  const revenueTargetColumn = getRevenueTargetColumn_(sheet);
  const lastRow = sheet.getLastRow();
  const lastColumn = Math.max(4, revenueTargetColumn, sheet.getLastColumn());
  if (lastRow < 2) return [];

  const values = sheet.getRange(1, 1, lastRow, lastColumn).getValues();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();

  return values.slice(1)
    .filter(function (row) {
      return normalizeMonth(row[0], timeZone) !== '' && displayText(row[1]) !== '';
    })
    .map(function (row) {
      return {
        month: normalizeMonth(row[0], timeZone),
        cre: displayText(row[1]),
        source: displayText(row[2] || ''),
        salesTarget: finiteNumber(row[3]),
        revenueTarget: finiteNumber(row[revenueTargetColumn - 1])
      };
    });
}

function finiteNumber(value) {
  if (typeof value === 'number') return isFinite(value) ? value : 0;
  const parsed = Number(String(value === null || value === undefined ? '' : value).replace(/,/g, '').trim());
  return isFinite(parsed) ? parsed : 0;
}

function upsertSalesVsTarget(month, cre, source, salesTarget) {
  const sheet = getSalesVsTargetSheet();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const monthKey = normalizeMonth(month, timeZone);
  const creDisplay = displayText(cre);
  const sourceDisplay = displayText(source);
  const targetValue = finiteNumber(salesTarget);

  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('Invalid month "' + displayText(month) + '". Expected format YYYY-MM, e.g. 2026-10.');
  }
  if (!creDisplay) throw new Error('CRE is required to save a target.');
  if (salesTarget === null || salesTarget === undefined || String(salesTarget).trim() === '' || !isFinite(Number(String(salesTarget).replace(/,/g, '')))) {
    throw new Error('A valid numeric salesTarget is required.');
  }

  const values = sheet.getDataRange().getValues();
  const monthCreRows = [];
  for (let i = 1; i < values.length; i++) {
    const rowMonth = normalizeMonth(values[i][0], timeZone);
    const rowCre = normalizeText(values[i][1]);
    if (rowMonth === monthKey && rowCre === normalizeText(creDisplay)) {
      monthCreRows.push({ rowNumber: i + 1, source: displayText(values[i][2] || '') });
    }
  }

  let match = null;
  if (sourceDisplay) {
    match = monthCreRows.find(function (candidate) {
      return normalizeText(candidate.source) === normalizeText(sourceDisplay);
    }) || null;
  } else if (monthCreRows.length === 1) {
    // Backward compatibility: if old frontend omitted source and this CRE
    // has exactly one row for the month, update that unique row safely.
    match = monthCreRows[0];
  } else if (monthCreRows.length > 1) {
    throw new Error('Source is required: multiple targets exist for ' + creDisplay + ' in ' + monthKey + '.');
  }

  let savedSource = sourceDisplay;
  let rowNumber;
  if (match) {
    rowNumber = match.rowNumber;
    if (!savedSource) savedSource = match.source;
    // Standardize Month as text YYYY-MM so future reads are deterministic.
    sheet.getRange(rowNumber, 1).setNumberFormat('@').setValue(monthKey);
    // Preserve the existing CRE/source display unless a new source was given.
    if (!displayText(values[rowNumber - 1][1])) sheet.getRange(rowNumber, 2).setValue(creDisplay);
    if (sourceDisplay) sheet.getRange(rowNumber, 3).setValue(sourceDisplay);
    sheet.getRange(rowNumber, 4).setValue(targetValue);
  } else {
    rowNumber = sheet.getLastRow() + 1;
    sheet.getRange(rowNumber, 1, 1, 4).setNumberFormat('@');
    sheet.getRange(rowNumber, 1, 1, 4).setValues([[monthKey, creDisplay, sourceDisplay, targetValue]]);
    // Restore numeric formatting for target value.
    sheet.getRange(rowNumber, 4).setNumberFormat('#,##0.##');
  }

  SpreadsheetApp.flush();

  // Read back the exact row directly, then verify it before returning success.
  const readBack = sheet.getRange(rowNumber, 1, 1, 4).getValues()[0];
  const readMonth = normalizeMonth(readBack[0], timeZone);
  const readCre = displayText(readBack[1]);
  const readSource = displayText(readBack[2] || '');
  const readTarget = finiteNumber(readBack[3]);

  if (readMonth !== monthKey || normalizeText(readCre) !== normalizeText(creDisplay) ||
      normalizeText(readSource) !== normalizeText(savedSource) ||
      Math.abs(readTarget - targetValue) > 0.000001) {
    throw new Error('Target row read-back verification failed. Row ' + rowNumber +
      ' returned month=' + readMonth + ', CRE=' + readCre + ', source=' + readSource +
      ', salesTarget=' + readTarget + '.');
  }

  return { month: monthKey, monthKey: monthKey, cre: readCre, source: readSource, salesTarget: readTarget, rowNumber: rowNumber };
}

function upsertRevenueTarget(month, cre, source, revenueTarget) {
  const sheet = getSalesVsTargetSheet();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const revenueColumn = getRevenueTargetColumn_(sheet);
  const monthKey = normalizeMonth(month, timeZone);
  const creDisplay = displayText(cre);
  const sourceDisplay = displayText(source);
  const targetValue = finiteNumber(revenueTarget);

  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('Invalid month "' + displayText(month) + '". Expected format YYYY-MM, e.g. 2026-10.');
  }
  if (!creDisplay) throw new Error('CRE is required to save a revenue target.');
  if (revenueTarget === null || revenueTarget === undefined || String(revenueTarget).trim() === '' || !isFinite(Number(String(revenueTarget).replace(/,/g, '')))) {
    throw new Error('A valid numeric revenueTarget is required.');
  }

  const values = sheet.getDataRange().getValues();
  const monthCreRows = [];
  for (let i = 1; i < values.length; i++) {
    if (normalizeMonth(values[i][0], timeZone) === monthKey && normalizeText(values[i][1]) === normalizeText(creDisplay)) {
      monthCreRows.push({ rowNumber: i + 1, source: displayText(values[i][2] || '') });
    }
  }

  let match = null;
  if (sourceDisplay) {
    match = monthCreRows.find(function (candidate) {
      return normalizeText(candidate.source) === normalizeText(sourceDisplay);
    }) || null;
  } else if (monthCreRows.length === 1) {
    match = monthCreRows[0];
  } else if (monthCreRows.length > 1) {
    throw new Error('Source is required: multiple targets exist for ' + creDisplay + ' in ' + monthKey + '.');
  }

  let savedSource = sourceDisplay;
  let rowNumber;
  if (match) {
    rowNumber = match.rowNumber;
    if (!savedSource) savedSource = match.source;
    sheet.getRange(rowNumber, 1).setNumberFormat('@').setValue(monthKey);
    if (sourceDisplay) sheet.getRange(rowNumber, 3).setValue(sourceDisplay);
  } else {
    rowNumber = sheet.getLastRow() + 1;
    // Keep the existing columns A:O intact and set only the identity + SalesTarget default + new RevenueTarget column.
    sheet.getRange(rowNumber, 1, 1, 4).setValues([[monthKey, creDisplay, sourceDisplay, 0]]);
  }
  sheet.getRange(rowNumber, revenueColumn).setValue(targetValue).setNumberFormat('#,##0.##');
  SpreadsheetApp.flush();

  const readBack = sheet.getRange(rowNumber, 1, 1, Math.max(4, revenueColumn)).getValues()[0];
  const readMonth = normalizeMonth(readBack[0], timeZone);
  const readCre = displayText(readBack[1]);
  const readSource = displayText(readBack[2] || '');
  const readTarget = finiteNumber(readBack[revenueColumn - 1]);
  if (readMonth !== monthKey || normalizeText(readCre) !== normalizeText(creDisplay) ||
      normalizeText(readSource) !== normalizeText(savedSource) || Math.abs(readTarget - targetValue) > 0.000001) {
    throw new Error('Revenue target read-back verification failed at row ' + rowNumber +
      '. Returned month=' + readMonth + ', CRE=' + readCre + ', source=' + readSource +
      ', revenueTarget=' + readTarget + '.');
  }
  return { month: monthKey, monthKey: monthKey, cre: readCre, source: readSource, revenueTarget: readTarget, rowNumber: rowNumber };
}

function resetAllTargets(month) {
  const sheet = getSalesVsTargetSheet();
  const timeZone = sheet.getParent().getSpreadsheetTimeZone();
  const monthKey = normalizeMonth(month, timeZone);
  if (!/^\d{4}-\d{2}$/.test(monthKey)) {
    throw new Error('Invalid month "' + displayText(month) + '". Expected format YYYY-MM.');
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return;
  const revenueColumn = getRevenueTargetColumn_(sheet);
  const values = sheet.getRange(1, 1, lastRow, Math.max(4, revenueColumn, sheet.getLastColumn())).getValues();
  for (let i = 1; i < values.length; i++) {
    if (normalizeMonth(values[i][0], timeZone) === monthKey) {
      sheet.getRange(i + 1, 4).setValue(0);
      sheet.getRange(i + 1, revenueColumn).setValue(0);
    }
  }
  SpreadsheetApp.flush();
}

function testSalesVsTarget() {
  const rows = readSalesVsTargetRows();
  console.log('SalesVsTarget rows: ' + rows.length);
  console.log(JSON.stringify(rows));
}

function checkSalesVsTargetSheet() {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(SALES_VS_TARGET_SHEET);
  console.log('Spreadsheet: ' + ss.getName());
  console.log('Spreadsheet ID: ' + ss.getId());
  if (!sheet) throw new Error('Sheet "' + SALES_VS_TARGET_SHEET + '" NOT FOUND.');
  console.log('Sheet: ' + sheet.getName());
  console.log('Last row: ' + sheet.getLastRow());
  console.log('Last column: ' + sheet.getLastColumn());
  console.log('First 10 rows: ' + JSON.stringify(sheet.getDataRange().getValues().slice(0, 10)));
}


/**
 * ONE-TIME CLEANUP: run this manually from the Apps Script editor (Run > cleanupDuplicateTargets).
 * For each Month + CRE (case-insensitive) it keeps ONE row, preferring the "All Sources" row, else the
 * last row with a non-zero SalesTarget, else the last row. Blank-source duplicates created by older
 * builds are deleted, and the kept row's source is set to "All Sources". Check the log afterwards.
 */
function cleanupDuplicateTargets() {
  withScriptLock(function () {
    const sheet = getSalesVsTargetSheet();
    const tz = sheet.getParent().getSpreadsheetTimeZone();
    const revCol = getRevenueTargetColumn_(sheet);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) return;
    const values = sheet.getRange(1, 1, lastRow, Math.max(4, revCol)).getValues();
    const groups = {};
    for (let i = 1; i < values.length; i++) {
      const m = normalizeMonth(values[i][0], tz), c = normalizeText(values[i][1]);
      if (!m || !c) continue;
      (groups[m + '|' + c] = groups[m + '|' + c] || []).push(i + 1);
    }
    const toDelete = [];
    Object.keys(groups).forEach(function (key) {
      const rowsN = groups[key];
      if (rowsN.length < 2) return;
      let keep = rowsN.filter(function (n) { return normalizeText(values[n - 1][2]) === 'all sources'; }).pop();
      if (!keep) keep = rowsN.filter(function (n) { return finiteNumber(values[n - 1][3]) > 0; }).pop();
      if (!keep) keep = rowsN[rowsN.length - 1];
      const bestRev = rowsN.map(function (n) { return finiteNumber(values[n - 1][revCol - 1]); })
        .filter(function (v) { return v > 0; }).pop();
      sheet.getRange(keep, 3).setValue('All Sources');
      if (bestRev && !finiteNumber(values[keep - 1][revCol - 1])) sheet.getRange(keep, revCol).setValue(bestRev);
      rowsN.forEach(function (n) { if (n !== keep) toDelete.push(n); });
    });
    toDelete.sort(function (a, b) { return b - a; }).forEach(function (n) { sheet.deleteRow(n); });
    SpreadsheetApp.flush();
    console.log('Removed ' + toDelete.length + ' duplicate row(s).');
  });
}
