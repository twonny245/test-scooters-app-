/*
  Apps Script for the AA Scooters spreadsheet.

  Includes: customer intake, bike parts/oil data, bike photos (Drive-backed),
  and Bike Tax category lookup (used by Available Bikes to price by category).
*/

var PARTS_SHEET_NAME = 'Parts and Oil change';
var OPERATION_SHEET_NAME = 'Operation';
var BIKE_TAX_SHEET_NAME = 'Bike Tax';
var BIKES_SHEET_NAME = 'bikes';

// The "bikes" sheet has a second, separate table further down tracking
// expenses per bike per month -- same column layout as the income table
// at the top (the header row at the very top of the sheet is frozen, so
// column J still means "July" down here too), just its own list of bike
// names starting at column A, row 52 onward instead of row 2.
var BIKES_EXPENSE_SECTION_START_ROW = 52;

// ID of the Drive folder that holds one subfolder per bike, full of that
// bike's photos. Get this from the folder's URL:
// https://drive.google.com/drive/folders/THIS_PART_IS_THE_ID
// The Google account running this script must have at least Editor access
// to this folder (owning it themselves is simplest).
var PHOTOS_ROOT_FOLDER_ID = '1E11bBgY5BeohoSiDCffA1Uz4-YQJOt7U';

// =====================================================================
// ---- Post-write verification ("did it really land?") ----
//
// Every function below that writes to the spreadsheet also records WHAT
// it wrote and WHERE (sheet + row + column + the value it expects to
// find there) into VERIFY_LOG. Just before an action builds its
// response, it calls runWriteVerification(), which flushes any pending
// writes, re-reads every recorded cell fresh from the spreadsheet, and
// compares. Anything that doesn't match -- or any write that was
// skipped because its target tab/row couldn't be found -- comes back as
// a warning in the response, so the page can tell the user to go check
// the sheet manually instead of a write silently going missing.
//
// VERIFY_LOG is a script global: Apps Script runs each web-app request
// in its own isolated execution, so there's no bleed between requests.
// verifyReset() at the top of doPost clears it per-request anyway, as a
// belt-and-braces measure.
// =====================================================================
var VERIFY_LOG = [];

function verifyReset() {
  VERIFY_LOG = [];
}

// Records: "cell (row, col) on sheetName should now hold `expected`".
function verifyCell(sheetName, row, col, expected, label) {
  VERIFY_LOG.push({ kind: 'cell', sheetName: sheetName, row: row, col: col, expected: expected, label: label });
}

// Records: "cell (row, col) on sheetName should NO LONGER hold
// `oldValue`" -- used after deleting a row's cells, where whatever
// shifted up into its place can't be known in advance, but finding the
// exact same old value still sitting there means the delete probably
// didn't take.
function verifyCellChanged(sheetName, row, col, oldValue, label) {
  VERIFY_LOG.push({ kind: 'changed', sheetName: sheetName, row: row, col: col, oldValue: oldValue, label: label });
}

// Records a problem noticed at write time (e.g. the target tab doesn't
// exist, so nothing was written at all). These are always reported.
function verifyProblem(message) {
  VERIFY_LOG.push({ kind: 'problem', message: message });
}

// Normalizes a value so "the same thing" compares equal regardless of
// how Sheets stored it: Dates become dd/MM/yyyy strings (or HH:mm for
// time-only cells, which come back as dates near 1899-12-30), numbers
// and numeric-looking strings become plain rounded numbers, and other
// strings are trimmed + lowercased.
function verifyNormalize(v, tz) {
  if (v === null || v === undefined) return '';
  if (v instanceof Date) {
    if (v.getFullYear() < 1930) return Utilities.formatDate(v, tz, 'HH:mm');
    return Utilities.formatDate(v, tz, 'dd/MM/yyyy');
  }
  if (typeof v === 'number') return String(Math.round(v * 100) / 100);
  var s = v.toString().trim();
  if (s !== '' && !isNaN(Number(s))) return String(Math.round(Number(s) * 100) / 100);
  return s.toLowerCase();
}

function verifyValuesMatch(expected, actual, tz) {
  var ne = verifyNormalize(expected, tz);
  var na = verifyNormalize(actual, tz);
  if (ne === na) return true;
  if (ne !== '' && na !== '' && !isNaN(Number(ne)) && !isNaN(Number(na))) {
    return Math.abs(Number(ne) - Number(na)) < 0.01; // numeric tolerance
  }
  return false;
}

// Human-friendly rendering of a value for warning messages.
function verifyDisplay(v, tz) {
  var n = verifyNormalize(v, tz);
  return n === '' ? '(blank)' : n;
}

// Re-reads every recorded write fresh from the spreadsheet and compares.
// Returns { problems: [...], checked: N, failed: N } -- problems is
// empty when every check passed. Never throws: a check that itself
// errors becomes a problem message instead, so verification can't
// break the action.
function runWriteVerification(ss) {
  var problems = [];
  var checked = 0;
  var failed = 0;
  if (!VERIFY_LOG.length) return { problems: problems, checked: checked, failed: failed };

  try { SpreadsheetApp.flush(); } catch (flushErr) {}

  var tz = ss.getSpreadsheetTimeZone();
  for (var i = 0; i < VERIFY_LOG.length; i++) {
    var entry = VERIFY_LOG[i];
    try {
      if (entry.kind === 'problem') {
        problems.push(entry.message);
        continue;
      }
      checked++;
      var sheet = ss.getSheetByName(entry.sheetName);
      if (!sheet) {
        problems.push('CHECK FAILED (' + entry.label + '): tab "' + entry.sheetName + '" not found when re-reading.');
        continue;
      }
      var actual = sheet.getRange(entry.row, entry.col).getValue();
      var where = '"' + entry.sheetName + '" row ' + entry.row + ', column ' + columnToLetter(entry.col);
      if (entry.kind === 'cell') {
        if (!verifyValuesMatch(entry.expected, actual, tz)) {
          failed++;
          problems.push('CHECK FAILED (' + entry.label + '): expected "' + verifyDisplay(entry.expected, tz) +
            '" at ' + where + ' but found "' + verifyDisplay(actual, tz) + '" -- please check the sheet manually.');
        }
      } else if (entry.kind === 'changed') {
        if (verifyValuesMatch(entry.oldValue, actual, tz) && verifyNormalize(entry.oldValue, tz) !== '') {
          failed++;
          problems.push('CHECK (' + entry.label + '): ' + where + ' still shows "' + verifyDisplay(entry.oldValue, tz) +
            '" after the delete -- if there were two identical entries this is expected, otherwise please check the sheet manually.');
        }
      }
    } catch (checkErr) {
      problems.push('Post-write check could not run (' + (entry.label || 'unlabelled write') + '): ' + checkErr.message);
    }
  }
  VERIFY_LOG = [];
  return { problems: problems, checked: checked, failed: failed };
}

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    verifyReset();

    if (data.action === 'updateBike') {
      return updateBikeRow(data);
    }
    if (data.action === 'uploadPhoto') {
      return uploadBikePhoto(data);
    }
    if (data.action === 'deletePhoto') {
      return deleteBikePhoto(data);
    }
    if (data.action === 'markReturned') {
      return markBikeReturned(data);
    }
    if (data.action === 'extendBike') {
      return extendBikeRow(data);
    }
    if (data.action === 'closeBikeForExtend') {
      return closeBikeForExtend(data);
    }
    if (data.action === 'addExpense') {
      return addExpenseRow(data);
    }
    if (data.action === 'editExpense') {
      return editExpenseRow(data);
    }
    if (data.action === 'addIncome') {
      return addIncomeRow(data);
    }
    if (data.action === 'editIncome') {
      return editIncomeRow(data);
    }
    if (data.action === 'deleteExpense') {
      return deleteExpenseRow(data);
    }
    if (data.action === 'deleteIncome') {
      return deleteIncomeRow(data);
    }

    // ---- Customer-intake behavior ----
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    // Columns: A timestamp, B contact, C name, D nationality, E passport,
    // F bikeModel, G status (blank, unused), H rentingDateFrom, I returnDate,
    // J returnTime, K deliverToHotel, L totalPrice, M paidBy, N situation
    // (left blank here — set later by markBikeReturned/closeBikeForExtend),
    // O deposit method (blank if no deposit was taken), P source ("Direct"
    // for a normal walk-in add on this page, "Extend" when this row was
    // auto-populated by the Extend flow on Bikes Status).
    var isExtendSource = (data.source || '').toString().trim().toLowerCase() === 'extend';

    sheet.appendRow([
      Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'dd/MM/yyyy'),
      data.contact || '',
      data.name || '',
      data.nationality || '',
      data.passport || '',
      data.bikeModel || '',
      '',
      formatIsoDateToDMY(data.rentingDateFrom),
      formatIsoDateToDMY(data.returnDate),
      data.returnTime || '',
      data.deliverToHotel || '',
      data.totalPrice || '',
      data.paidBy || '',
      '',
      data.deposit || '',
      isExtendSource ? 'Extend' : 'Direct'
    ]);

    var newRow = sheet.getLastRow();
    var numCols = 16;
    var newRange = sheet.getRange(newRow, 1, 1, numCols);
    newRange.setBorder(true, true, true, true, true, true);

    // Register the key cells of the new customer row for post-write
    // verification (re-read + compared just before responding).
    verifyCell('customer', newRow, 2, data.contact || '', 'customer row: contact');
    verifyCell('customer', newRow, 3, data.name || '', 'customer row: name');
    verifyCell('customer', newRow, 6, data.bikeModel || '', 'customer row: bike');
    verifyCell('customer', newRow, 8, formatIsoDateToDMY(data.rentingDateFrom), 'customer row: renting-from date');
    verifyCell('customer', newRow, 9, formatIsoDateToDMY(data.returnDate), 'customer row: return date');
    verifyCell('customer', newRow, 12, data.totalPrice || '', 'customer row: total price');
    verifyCell('customer', newRow, 13, data.paidBy || '', 'customer row: paid by');

    var fromDate = data.rentingDateFrom ? new Date(data.rentingDateFrom + 'T00:00:00') : null;
    var toDate = data.returnDate ? new Date(data.returnDate + 'T00:00:00') : null;
    var dayCount = null;
    if (fromDate && toDate) {
      dayCount = Math.round((toDate - fromDate) / (1000 * 60 * 60 * 24));
      var fillColor = dayCount >= 30 ? '#00ffff' : '#93C47D';
      sheet.getRange(newRow, 2, 1, 3).setBackground(fillColor);
    }

    // Also log this rental as a row on the current month's income sheet
    // (e.g. "July"). Wrapped so a problem here never breaks customer intake
    // -- but no longer silently: any error comes back as a warning.
    var incomeSheetWarning = null;
    try {
      appendMonthlyIncomeRow(ss, data, dayCount);
    } catch (incomeErr) {
      incomeSheetWarning = 'Income sheet: ' + incomeErr.message;
    }

    // If (and only if) paid in cash, also log it on the "cash" sheet.
    // Wrapped so a problem here never breaks customer intake -- but no
    // longer silently: any error comes back as a warning (a cash row
    // once went missing with nothing to show for it).
    var cashSheetWarning = null;
    try {
      if ((data.paidBy || '').toString().trim().toLowerCase() === 'cash') {
        appendCashSheetRow(ss, data, dayCount);
      }
    } catch (cashErr) {
      cashSheetWarning = 'Cash sheet: ' + cashErr.message;
    }

    // Also roll this rental's amount into the bike's own running total for
    // the current month on the "bikes" sheet (e.g. "=2000+2500+2800"), so
    // that sheet keeps tracking income per bike per month. Any problem here
    // (bike name / month column not found, etc.) is surfaced as a warning
    // below rather than breaking customer intake, which has already succeeded.
    var bikesSheetWarning = null;
    try {
      addRentalAmountToBikesSheet(ss, data.bikeModel, data.totalPrice);
    } catch (bikesErr) {
      bikesSheetWarning = bikesErr.message;
      Logger.log('Bikes sheet update warning: ' + bikesErr.message);
    }

    // If paid by Wise or Revolut, add the amount into that method's
    // deposit-tracking cell on the current month's sheet (self-locating via
    // the label in column L), as a running "=X+Y" formula. If the label
    // can't be verified/found anywhere, nothing is written to column M --
    // instead a warning is logged and returned in the response, rather than
    // silently guessing at the wrong cell.
    var depositWarning = null;
    try {
      var paidByLower = (data.paidBy || '').toString().trim().toLowerCase();
      if (paidByLower === 'wise' || paidByLower === 'revolut') {
        processDepositForPayment(ss, paidByLower, data.totalPrice);
      }
    } catch (depositErr) {
      depositWarning = depositErr.message;
      Logger.log('Deposit update warning: ' + depositErr.message);
    }

    // Security deposit (the checkbox/dropdown on the intake form -- separate
    // from how the rental itself was paid). Only Scan/Wise/Revolut need a
    // logged row; Cash and Passport are just noted on the customer row and
    // need nothing else. Never applicable to an extension row (no deposit
    // is taken when extending an existing rental), so it's skipped outright
    // in that case -- this is a server-side safety net on top of the intake
    // form already hiding the deposit fields for an extension. Wrapped so a
    // problem here never breaks the rest of customer intake, which has
    // already succeeded by this point.
    var securityDepositWarning = null;
    try {
      var depositMethodLower = (data.deposit || '').toString().trim().toLowerCase();
      if (!isExtendSource && (depositMethodLower === 'scan' || depositMethodLower === 'wise' || depositMethodLower === 'revolut')) {
        logSecurityDeposit(ss, depositMethodLower, data.depositAmount, data.name);
      }
    } catch (secDepErr) {
      securityDepositWarning = secDepErr.message;
      Logger.log('Security deposit log warning: ' + secDepErr.message);
    }

    var responsePayload = { success: true };
    var warnings = [];
    if (incomeSheetWarning) warnings.push(incomeSheetWarning);
    if (cashSheetWarning) warnings.push(cashSheetWarning);
    if (depositWarning) warnings.push(depositWarning);
    if (securityDepositWarning) warnings.push(securityDepositWarning);
    if (bikesSheetWarning) warnings.push(bikesSheetWarning);

    // Post-write verification: re-read everything this intake wrote
    // (customer row, monthly income row, cash row, bikes total, deposit
    // cells) and confirm it actually landed where it should.
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);
    responsePayload.checksPassed = verification.checked - verification.failed;
    responsePayload.checksTotal = verification.checked;

    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// Converts a yyyy-MM-dd string (what <input type="date"> sends) into
// dd/MM/yyyy, matching the format already used everywhere else in the sheet.
// If the string doesn't look like yyyy-MM-dd, it's returned unchanged.
function formatIsoDateToDMY(isoStr) {
  if (!isoStr) return '';
  var m = String(isoStr).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
  if (!m) return isoStr;
  var y = m[1];
  var mo = m[2].length === 1 ? '0' + m[2] : m[2];
  var d = m[3].length === 1 ? '0' + m[3] : m[3];
  return d + '/' + mo + '/' + y;
}

// ---- Builds the "<bike> rent <N> day(s)" text used both on the monthly
// income sheet and the cash sheet, so the two stay in the same format.
// When this row came from the Extend flow (Bikes Status -> Extend), "rent"
// is swapped for "extend" instead -- e.g. "GT red 2 extend 30 days" -- so
// it's clear at a glance on those sheets which rows are fresh rentals and
// which are extensions of an existing one. ----
function buildRentalIncomeText(data, dayCount) {
  var bikeName = (data.bikeModel || '').toString().trim();
  var isExtendSource = (data.source || '').toString().trim().toLowerCase() === 'extend';
  var verb = isExtendSource ? 'extend' : 'rent';
  var text = bikeName;
  if (dayCount !== null && dayCount !== undefined && !isNaN(dayCount)) {
    text += ' ' + verb + ' ' + dayCount + (dayCount === 1 ? ' day' : ' days');
  } else {
    text += ' ' + verb;
  }
  return text;
}

// ---- The reverse of buildRentalIncomeText -- pulls the bike name back out
// of a rental/extension income description ("forza (300) rent 10 days" ->
// "forza (300)", "GT black 5 extend 43 days" -> "GT black 5"). Returns ''
// if the text doesn't look like a rent/extend line at all -- e.g. a
// manually-typed "Add Income" entry on the Accounts page -- so callers can
// tell a bike-rental income row apart from unrelated income and skip
// reconciling anything against the "bikes" sheet for the latter. ----
function extractBikeNameFromRentalIncomeText(text) {
  var m = (text || '').toString().trim().match(/^(.*?)\s+(rent|extend)\b/i);
  return m ? m[1].trim() : '';
}

// ---- Returns the sheet tab matching the current month's name (e.g.
// "July"), or null if there isn't one. Shared by anything that logs against
// the current month's sheet, so the lookup stays in one place. ----
function getCurrentMonthSheet(ss) {
  var monthName = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMMM'); // e.g. "July"
  return ss.getSheetByName(monthName);
}

// ---- Finds the first row (starting at startRow) where EVERY column
// listed in cols is blank -- not just one column checked as a stand-in
// for "the whole row is free". Every "append a new entry" writer below
// uses this before writing, so a stray value sitting in just one of the
// columns it's about to write into (left over from a manual edit, or an
// old version of this code writing to the wrong column -- exactly what
// happened with appendMonthlyIncomeRow before its column numbers were
// corrected) can never be silently overwritten.
//
// If every row up to the sheet's current used range already has
// something in at least one of `cols`, this returns the row right after
// the last used row -- i.e. a brand-new blank row, never a guess. ----
function findFullyEmptyRow(sheet, startRow, cols) {
  var minCol = cols[0], maxCol = cols[0];
  for (var c = 1; c < cols.length; c++) {
    if (cols[c] < minCol) minCol = cols[c];
    if (cols[c] > maxCol) maxCol = cols[c];
  }
  var width = maxCol - minCol + 1;
  var lastRow = sheet.getLastRow();
  var scanRows = Math.max(lastRow - startRow + 2, 1); // +1 so at least one always-blank row past the used range is included

  var values = sheet.getRange(startRow, minCol, scanRows, width).getValues();
  for (var i = 0; i < values.length; i++) {
    var rowAllBlank = true;
    for (var j = 0; j < cols.length; j++) {
      var v = values[i][cols[j] - minCol];
      if (v !== '' && v !== null && v !== undefined) { rowAllBlank = false; break; }
    }
    if (rowAllBlank) return startRow + i;
  }
  return startRow + values.length;
}

// ---- Log a new rental as a row on the current month's income sheet (e.g.
// "July"). Columns: F Date, G Income, H PAX name, I Amount, J paid. Finds
// the next row where ALL FIVE of those columns are blank via
// findFullyEmptyRow -- not just one column checked as a stand-in for "the
// whole row is free" -- so a stray value in any one of them never gets
// silently overwritten. If no sheet matches the current month name, this
// silently does nothing -- it's not created automatically, since it should
// already exist as a normal monthly tab. ----
function appendMonthlyIncomeRow(ss, data, dayCount) {
  var DATE_COL = 6;   // F
  var INCOME_COL = 7; // G
  var NAME_COL = 8;   // H
  var AMOUNT_COL = 9; // I
  var PAID_COL = 10;  // J

  var sheet = getCurrentMonthSheet(ss);
  if (!sheet) {
    // No tab for the current month -- nothing to log against, but say so
    // rather than skipping silently.
    var missingMonth = Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMMM');
    verifyProblem('Income sheet: no tab named "' + missingMonth + '" was found, so this entry was NOT logged on the monthly income sheet.');
    return;
  }

  var targetRow = findFullyEmptyRow(sheet, 2, [DATE_COL, INCOME_COL, NAME_COL, AMOUNT_COL, PAID_COL]);

  var incomeText = buildRentalIncomeText(data, dayCount);

  var paidByRaw = (data.paidBy || '').toString().trim().toLowerCase();
  var paidDisplay = paidByRaw === 'scan' ? 'QR scan' : paidByRaw;

  var amountValue = data.totalPrice !== '' && data.totalPrice !== undefined && !isNaN(Number(data.totalPrice))
    ? Number(data.totalPrice)
    : '';

  sheet.getRange(targetRow, DATE_COL, 1, 5).setValues([[
    new Date(), incomeText, data.name || '', amountValue, paidDisplay
  ]]);

  var monthName = sheet.getName();
  verifyCell(monthName, targetRow, DATE_COL, new Date(), '"' + monthName + '" income row: date');
  verifyCell(monthName, targetRow, INCOME_COL, incomeText, '"' + monthName + '" income row: description');
  verifyCell(monthName, targetRow, NAME_COL, data.name || '', '"' + monthName + '" income row: name');
  verifyCell(monthName, targetRow, AMOUNT_COL, amountValue, '"' + monthName + '" income row: amount');
  verifyCell(monthName, targetRow, PAID_COL, paidDisplay, '"' + monthName + '" income row: paid by');

  // Match the formatting (currency style, borders, banding) of the row
  // directly above, so the new row looks consistent with the rest.
  if (targetRow - 1 >= 2) {
    sheet.getRange(targetRow - 1, DATE_COL, 1, 5)
      .copyFormatToRange(sheet, DATE_COL, PAID_COL, targetRow, targetRow);
  }
}

// ---- Log a new rental as a row on the "cash" sheet, but ONLY when it was
// paid in cash. Unlike the monthly sheet, "cash" is one running log for the
// whole year (columns A income date, B income, C amount), so rows are just
// appended at the very bottom rather than looked up per month. Finds the
// next empty row via column A only, since column D ("expense"/"Tax" labels)
// can have its own unrelated entries further down. ----
function appendCashSheetRow(ss, data, dayCount) {
  appendCashSheetRowText(ss, buildRentalIncomeText(data, dayCount), data.totalPrice);
}

// ---- Same as appendCashSheetRow above, but takes the income label and
// amount directly instead of deriving them from a rental's bikeModel/
// dayCount. Shared by the rental cash logging above and by the Accounts
// page's manual "Add Income" flow (addIncomeRow below), so both paths hit
// the exact same sheet-writing logic. Returns the row number it wrote to
// (or undefined if there's no "cash" tab), so the caller can store it as a
// link for later edits/deletes to find their way back to this exact row. ----
function appendCashSheetRowText(ss, incomeText, rawAmount) {
  var DATE_COL = 1;   // A
  var INCOME_COL = 2; // B
  var AMOUNT_COL = 3; // C

  var sheet = ss.getSheetByName('cash');
  if (!sheet) {
    // No "cash" tab -- nothing to log against, but say so rather than
    // skipping silently.
    verifyProblem('Cash sheet: no tab named "cash" was found, so this entry was NOT logged on the cash sheet.');
    return;
  }

  var targetRow = findFullyEmptyRow(sheet, 2, [DATE_COL, INCOME_COL, AMOUNT_COL]);

  var amountValue = rawAmount !== '' && rawAmount !== undefined && rawAmount !== null && !isNaN(Number(rawAmount))
    ? Number(rawAmount)
    : '';

  sheet.getRange(targetRow, DATE_COL, 1, 3).setValues([[
    new Date(), incomeText, amountValue
  ]]);

  verifyCell('cash', targetRow, DATE_COL, new Date(), 'cash sheet income row: date');
  verifyCell('cash', targetRow, INCOME_COL, incomeText, 'cash sheet income row: description');
  verifyCell('cash', targetRow, AMOUNT_COL, amountValue, 'cash sheet income row: amount');

  // Match the formatting (currency style, borders) of the row directly
  // above, so the new row looks consistent with the rest.
  if (targetRow - 1 >= 2) {
    sheet.getRange(targetRow - 1, DATE_COL, 1, 3)
      .copyFormatToRange(sheet, DATE_COL, AMOUNT_COL, targetRow, targetRow);
  }
  return targetRow;
}

// ---- Same idea as appendCashSheetRowText above, but for the "cash"
// sheet's EXPENSE side -- columns E (date), F (description), G (amount),
// running independently of the income side in A-C (column D is just a
// manual highlight the user uses for their own bookkeeping, untouched
// here). Used when a cash expense is added via the Accounts page. Returns
// the row it wrote to, same reason as appendCashSheetRowText above. ----
function appendCashExpenseRowText(ss, expenseText, rawAmount) {
  var DATE_COL = 5;   // E
  var LABEL_COL = 6;  // F
  var AMOUNT_COL = 7; // G

  var sheet = ss.getSheetByName('cash');
  if (!sheet) {
    // No "cash" tab -- nothing to log against, but say so rather than
    // skipping silently.
    verifyProblem('Cash sheet: no tab named "cash" was found, so this expense was NOT logged on the cash sheet.');
    return;
  }

  var targetRow = findFullyEmptyRow(sheet, 2, [DATE_COL, LABEL_COL, AMOUNT_COL]);

  var amountValue = rawAmount !== '' && rawAmount !== undefined && rawAmount !== null && !isNaN(Number(rawAmount))
    ? Number(rawAmount)
    : '';

  sheet.getRange(targetRow, DATE_COL, 1, 3).setValues([[
    new Date(), expenseText, amountValue
  ]]);

  verifyCell('cash', targetRow, DATE_COL, new Date(), 'cash sheet expense row: date');
  verifyCell('cash', targetRow, LABEL_COL, expenseText, 'cash sheet expense row: description');
  verifyCell('cash', targetRow, AMOUNT_COL, amountValue, 'cash sheet expense row: amount');

  // Match the formatting (currency style, borders) of the row directly
  // above, so the new row looks consistent with the rest.
  if (targetRow - 1 >= 2) {
    sheet.getRange(targetRow - 1, DATE_COL, 1, 3)
      .copyFormatToRange(sheet, DATE_COL, AMOUNT_COL, targetRow, targetRow);
  }
  return targetRow;
}

// ---- Adds an amount into one of the fixed deposit-tracking cells on the
// current month's sheet -- M11 (next to the "wise(less deposit)" label in
// L11) for Wise, M12 (next to "revolut(less deposit)" in L12) for Revolut.
// These are fixed reference cells (not
// something that grows with new rental rows), and the goal is to keep a
// visible running total as a formula, e.g. "=100+300", rather than just
// silently replacing the number. If the cell is empty, the formula becomes
// "=amount". If it already holds a formula, "+amount" is appended to it. If
// it holds a plain (non-formula) number -- as these cells currently do --
// that number becomes the first term of a new "=existing+amount" formula,
// so nothing already there is lost. ----
function addAmountToDepositCell(sheet, row, col, rawAmount) {
  var amount = Number(rawAmount);
  if (rawAmount === '' || rawAmount === null || rawAmount === undefined || isNaN(amount)) return;

  var range = sheet.getRange(row, col);
  var formula = range.getFormula();

  // The evaluated value before this write -- the post-write check below
  // expects the cell to evaluate to (before + amount) afterwards.
  var beforeVal = Number(range.getValue());
  if (isNaN(beforeVal)) beforeVal = 0;

  if (formula && formula.charAt(0) === '=') {
    range.setFormula(formula + '+' + amount);
  } else {
    var currentValue = range.getValue();
    if (currentValue === '' || currentValue === null || isNaN(Number(currentValue))) {
      range.setFormula('=' + amount);
      beforeVal = 0;
    } else {
      range.setFormula('=' + Number(currentValue) + '+' + amount);
      beforeVal = Number(currentValue);
    }
  }

  verifyCell(sheet.getName(), row, col, beforeVal + amount,
    'running deposit total at "' + sheet.getName() + '" ' + columnToLetter(col) + row);
}

// ---- Same fuzzy bike-name matching used client-side in bikes.html and
// bike-name-audit.html, ported here so server-side income logging matches
// the same bike a customer row displays under (e.g. "Aerox Cool 1 (black)"
// on the customer sheet should still land on the "aerox cool 1" row of the
// "bikes" sheet, and "GT Black" should NOT be confused with "GT Black 2"). ----
function normalizeBikeNameForRentalLog(s) {
  return (s || '').toString()
    .toLowerCase()
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}
var RENTAL_LOG_DISTINGUISHING_SUFFIXES = {
  'one':1,'two':1,'three':1,'four':1,'five':1,'six':1,'seven':1,'eight':1,'nine':1,'ten':1,
  'i':1,'ii':1,'iii':1,'iv':1,'v':1,'vi':1,'vii':1,'viii':1,'ix':1,'x':1,
  '1':1,'2':1,'3':1,'4':1,'5':1,'6':1,'7':1,'8':1,'9':1,'10':1
};
function bikeNamesMatchForRentalLog(a, b) {
  var na = normalizeBikeNameForRentalLog(a);
  var nb = normalizeBikeNameForRentalLog(b);
  if (!na || !nb) return false;
  if (na === nb) return true;

  var ta = na.split(' ');
  var tb = nb.split(' ');
  var shorter = ta.length <= tb.length ? ta : tb;
  var longer = ta.length <= tb.length ? tb : ta;

  var isPrefix = true;
  for (var i = 0; i < shorter.length; i++) {
    if (shorter[i] !== longer[i]) { isPrefix = false; break; }
  }
  if (isPrefix) {
    var extra = longer.slice(shorter.length);
    for (var j = 0; j < extra.length; j++) {
      if (RENTAL_LOG_DISTINGUISHING_SUFFIXES[extra[j]]) return false;
    }
    return true;
  }

  return na.indexOf(nb) !== -1 || nb.indexOf(na) !== -1;
}

// ---- Finds the column on the "bikes" sheet matching the current month
// (e.g. "June"), by comparing header row 1 case-insensitively. The sheet's
// headers are sometimes abbreviated/mixed-case ("April", "july", "sept"),
// so this matches on the first 3 letters of the month name rather than
// requiring an exact match. Returns -1 if no column matches. ----
function findBikesSheetMonthColumn(sheet, monthName) {
  var lastCol = sheet.getLastColumn();
  var headerRow = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  var targetShort = monthName.toString().trim().toLowerCase().slice(0, 3);
  for (var c = 0; c < headerRow.length; c++) {
    var h = (headerRow[c] || '').toString().trim().toLowerCase();
    if (h && h.slice(0, 3) === targetShort) return c + 1;
  }
  return -1;
}

// ---- Finds the row on the "bikes" sheet whose column A bike name
// fuzzy-matches the given bike name, scanning from startRow (default 2,
// the income table) down to the sheet's last row. Pass
// BIKES_EXPENSE_SECTION_START_ROW to search the expense table further
// down instead -- restricting the scan to start there (rather than
// scanning the whole column) avoids accidentally matching the SAME bike
// name's row up in the income table. Returns -1 if no row matches. ----
function findBikesSheetRow(sheet, bikeName, startRow) {
  var start = startRow || 2;
  var lastRow = sheet.getLastRow();
  if (lastRow < start) return -1;
  var names = sheet.getRange(start, 1, lastRow - start + 1, 1).getValues();
  for (var i = 0; i < names.length; i++) {
    if (bikeNamesMatchForRentalLog(names[i][0], bikeName)) return i + start;
  }
  return -1;
}

// ---- Adds a rental/extension amount into the bike's cell for the current
// month on the "bikes" sheet, so each bike's monthly income total keeps
// growing as a visible running formula (e.g. "=2000+2500+2800") every time
// it's rented or extended -- same running-formula approach already used for
// the Wise/Revolut deposit totals in addAmountToDepositCell(). If the cell
// is empty, the formula becomes "=amount". If it already holds a formula,
// "+amount" is appended. If it holds a plain (non-formula) number, that
// number becomes the first term of a new "=existing+amount" formula, so
// nothing already there is lost. Throws (rather than silently doing
// nothing) if the "bikes" sheet, the bike's row, or the month's column
// can't be found, so a naming mismatch surfaces as a warning instead of
// quietly skipping the update. ----
// monthNameOverride lets a caller working against a specific (possibly
// past) month sheet -- like editIncomeRow, reconciling an edit made via
// the Accounts page's month selector -- target that exact month's column
// instead of always assuming "right now". Defaults to the real current
// month, which is what a brand-new rental/extension (doPost,
// extendBikeRow) should always use. rawAmount can be negative -- this is
// also how an edit's delta gets applied (subtract the old amount, add the
// new one), same running-formula approach as everywhere else here.
// sectionStartRowOverride lets a caller target the "bikes" sheet's
// separate EXPENSE table (BIKES_EXPENSE_SECTION_START_ROW) instead of the
// default income table (row 2) -- same sheet, same month columns (the
// header row is frozen/shared), just a different block of bike-name rows
// further down. ----
function addRentalAmountToBikesSheet(ss, bikeModel, rawAmount, monthNameOverride, sectionStartRowOverride) {
  var amount = Number(rawAmount);
  if (rawAmount === '' || rawAmount === null || rawAmount === undefined || isNaN(amount) || amount === 0) return;

  var sheet = ss.getSheetByName(BIKES_SHEET_NAME);
  if (!sheet) {
    throw new Error('Sheet named "' + BIKES_SHEET_NAME + '" not found -- bike monthly total was NOT updated.');
  }

  var bikeNameTrimmed = (bikeModel || '').toString().trim();
  if (!bikeNameTrimmed) {
    throw new Error('No bike name given -- bike monthly total was NOT updated.');
  }

  var row = findBikesSheetRow(sheet, bikeNameTrimmed, sectionStartRowOverride);
  if (row === -1) {
    throw new Error('Could not find a row for "' + bikeNameTrimmed + '" on the "' + BIKES_SHEET_NAME +
      '" sheet -- its monthly total was NOT updated.');
  }

  var monthName = monthNameOverride || Utilities.formatDate(new Date(), ss.getSpreadsheetTimeZone(), 'MMMM'); // e.g. "July"
  var col = findBikesSheetMonthColumn(sheet, monthName);
  if (col === -1) {
    throw new Error('Could not find a "' + monthName + '" column on the "' + BIKES_SHEET_NAME +
      '" sheet -- "' + bikeNameTrimmed + '"\'s monthly total was NOT updated.');
  }

  var cell = sheet.getRange(row, col);
  var formula = cell.getFormula();

  // The evaluated value before this write -- the post-write check below
  // expects the cell to evaluate to (before + amount) afterwards.
  var beforeVal = Number(cell.getValue());
  if (isNaN(beforeVal)) beforeVal = 0;

  if (formula && formula.charAt(0) === '=') {
    cell.setFormula(formula + '+' + amount);
  } else {
    var currentValue = cell.getValue();
    if (currentValue === '' || currentValue === null || isNaN(Number(currentValue))) {
      cell.setFormula('=' + amount);
      beforeVal = 0;
    } else {
      cell.setFormula('=' + Number(currentValue) + '+' + amount);
      beforeVal = Number(currentValue);
    }
  }

  verifyCell(sheet.getName(), row, col, beforeVal + amount,
    '"' + bikeNameTrimmed + '" monthly total (' + monthName + ') on the "' + BIKES_SHEET_NAME + '" sheet');
}

// ---- doGet: action = 'bikesList' -- returns every unique bike name in
// column A of the "bikes" sheet (income table, expense table, or both --
// they're expected to list the same bikes, so this just de-duplicates),
// for the Accounts page's "attach to bike" dropdown when adding/editing an
// expense. Returns [] rather than throwing if the sheet is missing, so a
// problem here never breaks the rest of the Accounts page. ----
// ---- Column A of the "bikes" sheet isn't ONLY bike names -- between the
// income table and the expense table (and at the end of each) there are
// label rows like "totals", "expense %", "Expenses" (the expense table's
// own header, sitting right above its first bike at
// BIKES_EXPENSE_SECTION_START_ROW), and "total" (closing out the expense
// table). Flags any of those so getBikesListNames can skip them rather
// than offering "total" as something to attach an expense to. ----
function looksLikeBikesSheetLabel(raw) {
  var t = (raw || '').toString().trim().toLowerCase();
  if (!t) return true; // blank -- not a bike name either.
  if (t.indexOf('total') === 0) return true; // "total", "totals"
  if (t === 'expenses' || t === 'expense' || t.indexOf('expense %') === 0) return true;
  return false;
}

function getBikesListNames() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(BIKES_SHEET_NAME);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, bikes: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var lastRow = sheet.getLastRow();
    var bikes = [];
    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
      var seen = {};
      for (var i = 0; i < values.length; i++) {
        var name = (values[i][0] || '').toString().trim();
        if (!name || looksLikeBikesSheetLabel(name)) continue;
        var key = name.toLowerCase();
        if (!seen[key]) {
          seen[key] = true;
          bikes.push(name);
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, bikes: bikes }))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Converts a 1-based column number to its spreadsheet letter (12 ->
// "L", 13 -> "M", etc.), purely for readable error messages below. ----
function columnToLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ---- Finds the row where a given label (e.g. "wise(less deposit)") lives
// in a label column. Checks the expected row first; if the label there
// doesn't match (case/whitespace-insensitive), searches the whole column
// for it instead, in case the sheet's layout shifted vertically. Returns
// the row number if found, or null if the label isn't anywhere in the
// column. ----
function findDepositRow(sheet, expectedRow, labelCol, expectedLabel) {
  function norm(s) { return (s || '').toString().trim().toLowerCase(); }
  var target = norm(expectedLabel);

  var atExpected = norm(sheet.getRange(expectedRow, labelCol).getValue());
  if (atExpected === target) return expectedRow;

  var maxRow = sheet.getMaxRows();
  var colValues = sheet.getRange(1, labelCol, maxRow, 1).getValues();
  for (var i = 0; i < colValues.length; i++) {
    if (norm(colValues[i][0]) === target) return i + 1;
  }
  return null; // Label not found anywhere in the column.
}

// ---- Entry point used by a new rental (doPost), an extension
// (extendBikeRow), and a manual Accounts income entry (addIncomeRow) alike
// for Wise/Revolut payments -- all three route through this one function,
// so fixing/verifying the target cell here fixes it everywhere at once.
//
// The label lives in column L ("wise(less deposit)" / "revolut(less
// deposit)") and the running total lives in column M, directly to its
// right -- this moved one column over from the previous K/L layout after
// a column was inserted on the sheet. findDepositRow() self-heals for the
// label moving to a different ROW, but as a second, independent safety net
// against the columns themselves shifting again in the future, this also
// re-checks -- right at the point of writing -- that the cell immediately
// to the left of the value cell actually still says the expected label. If
// it doesn't, this throws instead of silently writing into the wrong cell. ----
function processDepositForPayment(ss, paidByLower, rawAmount) {
  var LABEL_COL = 12;  // L
  var VALUE_COL = 13;  // M

  var sheet = getCurrentMonthSheet(ss);
  if (!sheet) {
    throw new Error('No sheet found for the current month -- could not update the ' + paidByLower + ' deposit total.');
  }

  var expectedRow = paidByLower === 'wise' ? 11 : 12;
  var expectedLabel = paidByLower === 'wise' ? 'wise(less deposit)' : 'revolut(less deposit)';

  var row = findDepositRow(sheet, expectedRow, LABEL_COL, expectedLabel);
  if (row === null) {
    throw new Error('Could not find a "' + expectedLabel + '" row in column ' + columnToLetter(LABEL_COL) +
      ' of the "' + sheet.getName() + '" sheet -- the ' + paidByLower + ' deposit total was NOT updated.');
  }

  // Safety check: confirm the cell directly to the left of where we're
  // about to write still says the expected label. This is deliberately
  // re-derived from VALUE_COL rather than reusing LABEL_COL, so if the two
  // ever drift apart in a future edit, this still catches it.
  var neighborCell = sheet.getRange(row, VALUE_COL - 1);
  var neighborLabel = (neighborCell.getValue() || '').toString().trim().toLowerCase();
  if (neighborLabel !== expectedLabel) {
    throw new Error('Safety check failed: ' + sheet.getName() + '!' + columnToLetter(VALUE_COL - 1) + row +
      ' does not say "' + expectedLabel + '" (found "' + neighborLabel + '" instead) -- the ' +
      paidByLower + ' deposit total was NOT updated. The column may have moved again.');
  }

  addAmountToDepositCell(sheet, row, VALUE_COL, rawAmount);
}

// ---- Security deposit tracking (the checkbox/dropdown added to the intake
// form -- NOT the same thing as processDepositForPayment above, which tracks
// running Wise/Revolut payment totals in fixed cells M11/M12). This instead
// logs each Scan/Wise/Revolut security deposit as its own row in a growing
// table on the current month's sheet:
//   Scan    -> N (date), O (amount), P (name)
//   Wise    -> Q (date), R (amount), S (name)
//   Revolut -> U (date), V (amount), W (name)
// Cash and Passport deposits need no logging here -- they're just noted on
// the customer row itself.
//
// Starting at row 2, finds the first row where both the amount and name
// cells are empty and writes there. Stops (and throws, so the intake
// response carries a warning rather than failing silently) if it reaches a
// row labelled "total" in the date column, so an existing totals row never
// gets overwritten. ----
function logSecurityDeposit(ss, methodLower, rawAmount, customerName) {
  var COLUMNS_BY_METHOD = {
    scan:    { date: 14, amount: 15, name: 16 }, // N, O, P
    wise:    { date: 17, amount: 18, name: 19 }, // Q, R, S
    revolut: { date: 21, amount: 22, name: 23 }  // U, V, W
  };
  var cols = COLUMNS_BY_METHOD[methodLower];
  if (!cols) return; // Cash/Passport/unrecognized -- nothing to log.

  var sheet = getCurrentMonthSheet(ss);
  if (!sheet) {
    throw new Error('No sheet found for the current month -- could not log the ' + methodLower + ' deposit.');
  }

  var maxRow = sheet.getMaxRows();
  var rowsToScan = maxRow - 1; // Starting from row 2.
  var dateColValues = sheet.getRange(2, cols.date, rowsToScan, 1).getValues();
  var amountColValues = sheet.getRange(2, cols.amount, rowsToScan, 1).getValues();
  var nameColValues = sheet.getRange(2, cols.name, rowsToScan, 1).getValues();

  var targetRow = null;
  for (var i = 0; i < amountColValues.length; i++) {
    var dateLabel = (dateColValues[i][0] || '').toString().trim().toLowerCase();
    if (dateLabel === 'total') break; // Don't write into or past the totals row.

    // All three cells this writes into (date, amount, name) have to be
    // blank -- not just amount+name -- so a row with a stray leftover
    // date but nothing else never gets silently overwritten.
    var dateEmpty = dateColValues[i][0] === '' || dateColValues[i][0] === null;
    var amtEmpty = amountColValues[i][0] === '' || amountColValues[i][0] === null;
    var nameEmpty = nameColValues[i][0] === '' || nameColValues[i][0] === null;
    if (dateEmpty && amtEmpty && nameEmpty) {
      targetRow = i + 2;
      break;
    }
  }

  if (!targetRow) {
    throw new Error('Could not find a free row above the totals row in the ' + methodLower +
      ' deposit section of "' + sheet.getName() + '" -- the deposit was NOT logged.');
  }

  sheet.getRange(targetRow, cols.date).setValue(new Date());
  sheet.getRange(targetRow, cols.amount).setValue(Number(rawAmount) || rawAmount || '');
  sheet.getRange(targetRow, cols.name).setValue(customerName || '');

  var secMonthName = sheet.getName();
  verifyCell(secMonthName, targetRow, cols.date, new Date(), methodLower + ' security deposit: date');
  verifyCell(secMonthName, targetRow, cols.amount, Number(rawAmount) || rawAmount || '', methodLower + ' security deposit: amount');
  verifyCell(secMonthName, targetRow, cols.name, customerName || '', methodLower + ' security deposit: customer name');
}

// ---- Update one row in the Parts and Oil change tab ----
function updateBikeRow(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet named "' + PARTS_SHEET_NAME + '" not found in this spreadsheet.');
    }

    var lastCol = sheet.getLastColumn();
    var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];

    var headerToCol = {};
    headers.forEach(function(h, i) {
      var key = (h || '').toString().trim();
      if (!key) key = 'Column ' + columnLetter(i + 1);
      headerToCol[key] = i + 1;
    });

    var bikeCol = 1;
    var lastRow = sheet.getLastRow();
    var bikeNames = sheet.getRange(1, bikeCol, lastRow, 1).getValues();

    var targetRow = -1;
    for (var r = 0; r < bikeNames.length; r++) {
      if ((bikeNames[r][0] || '').toString().trim() === (data.bike || '').toString().trim()) {
        targetRow = r + 1;
        break;
      }
    }

    if (targetRow === -1) {
      throw new Error('Bike "' + data.bike + '" not found in "' + PARTS_SHEET_NAME + '".');
    }

    var fields = data.fields || {};
    Object.keys(fields).forEach(function(headerName) {
      var col = headerToCol[headerName];
      if (col) {
        sheet.getRange(targetRow, col).setValue(fields[headerName]);
        verifyCell(PARTS_SHEET_NAME, targetRow, col, fields[headerName],
          '"' + headerName + '" for ' + data.bike);
      }
    });

    var responsePayload = { success: true };
    var verification = runWriteVerification(ss);
    if (verification.problems.length) responsePayload.warning = verification.problems.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Update one row in the customer tab: called from the "Return" button
// on the Bikes Status page. Sets the return date to the picked date, makes
// sure that date's font color is black (it's sometimes red from earlier
// conditional/manual formatting), and flips "situation" to "Returned". ----
function markBikeReturned(data) {
  try {
    var rowNumber = parseInt(data.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number.');
    }
    if (!data.returnDate) {
      throw new Error('No return date given.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var CUSTOMER_RETURN_DATE_COL = 9;  // I: Return date
    var CUSTOMER_RETURN_TIME_COL = 10; // J: Return time
    var CUSTOMER_SITUATION_COL = 14;   // N: situation

    var m = String(data.returnDate).trim().match(/^(\d{4})-(\d{1,2})-(\d{1,2})$/);
    if (!m) {
      throw new Error('Return date must be in yyyy-MM-dd format.');
    }
    var y = parseInt(m[1], 10), mo = parseInt(m[2], 10) - 1, d = parseInt(m[3], 10);
    var dateValue = new Date(y, mo, d);

    var dateCell = sheet.getRange(rowNumber, CUSTOMER_RETURN_DATE_COL);
    dateCell.setValue(dateValue);
    dateCell.setFontColor('#000000');

    // The return time cell is often left red by the same earlier
    // conditional/manual formatting as the return date, so it's normalized
    // to black here too, not just the date.
    sheet.getRange(rowNumber, CUSTOMER_RETURN_TIME_COL).setFontColor('#000000');

    sheet.getRange(rowNumber, CUSTOMER_SITUATION_COL).setValue('Returned');

    verifyCell('customer', rowNumber, CUSTOMER_RETURN_DATE_COL, dateValue, 'return date');
    verifyCell('customer', rowNumber, CUSTOMER_SITUATION_COL, 'Returned', 'situation');

    var responsePayload = { success: true };
    var verification = runWriteVerification(ss);
    if (verification.problems.length) responsePayload.warning = verification.problems.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Update one row in the customer tab: called from the "Extend" flow
// on the Bikes Status page for SHORT extensions only (under 30 days, and
// the "Extend 1 month" checkbox not ticked). Adds the given number of days
// onto whatever return date is currently in the sheet, and appends the
// newly paid amount onto the total price as a "=oldValue+amountPaid"
// formula. Longer extensions (1 month, or 30+ days) instead go through
// closeBikeForExtend() and a brand-new customer-intake row — see the
// "Extend" button's client-side logic in bikes.html. ----
function extendBikeRow(data) {
  try {
    var rowNumber = parseInt(data.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number.');
    }
    var daysToExtend = parseInt(data.daysToExtend, 10);
    if (!daysToExtend || daysToExtend <= 0) {
      throw new Error('Days to extend must be a positive number.');
    }
    var amountPaid = parseFloat(data.amountPaid);
    if (isNaN(amountPaid) || amountPaid < 0) {
      throw new Error('Amount paid must be a number.');
    }
    var paidBy = (data.paidBy || '').toString().trim();
    if (!paidBy) {
      throw new Error('Paid by is required.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var CUSTOMER_NAME_COL = 3;          // C: name
    var CUSTOMER_BIKE_COL = 6;          // F: bikeModel
    var CUSTOMER_RETURN_DATE_COL = 9;   // I: Return date
    var CUSTOMER_TOTAL_PRICE_COL = 12;  // L: total price
    var CUSTOMER_PAIDBY_COL = 13;       // M: paid by

    var dateCell = sheet.getRange(rowNumber, CUSTOMER_RETURN_DATE_COL);
    var currentDateValue = dateCell.getValue();
    var currentDate = currentDateValue instanceof Date ? new Date(currentDateValue.getTime()) : null;
    if (!currentDate) {
      // Fall back to parsing a dd/MM/yyyy string, in case the cell isn't a real Date.
      var m = String(currentDateValue).trim().match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
      if (!m) {
        throw new Error('Could not read the current return date to extend from.');
      }
      currentDate = new Date(parseInt(m[3], 10), parseInt(m[2], 10) - 1, parseInt(m[1], 10));
    }
    currentDate.setDate(currentDate.getDate() + daysToExtend);
    dateCell.setValue(currentDate);

    var priceCell = sheet.getRange(rowNumber, CUSTOMER_TOTAL_PRICE_COL);
    var currentPrice = Number(priceCell.getValue()) || 0;
    priceCell.setFormula('=' + currentPrice + '+' + amountPaid);

    // Record the payment method used for this extension on the row itself
    // (the sheet only has one "paid by" cell per row, so this reflects the
    // most recent payment — the same way the total price above rolls the
    // new amount into the existing figure).
    sheet.getRange(rowNumber, CUSTOMER_PAIDBY_COL).setValue(paidBy);

    verifyCell('customer', rowNumber, CUSTOMER_RETURN_DATE_COL, currentDate, 'extended return date');
    verifyCell('customer', rowNumber, CUSTOMER_TOTAL_PRICE_COL, currentPrice + amountPaid, 'total price after extension');
    verifyCell('customer', rowNumber, CUSTOMER_PAIDBY_COL, paidBy, 'paid by after extension');

    // Everything below mirrors what a brand-new rental (in doPost, above)
    // logs for its payment — the monthly income sheet, the cash sheet (if
    // paid in cash), and the Wise/Revolut running deposit total. A short
    // extension used to skip all of this, so it never showed up on the
    // current month's income sheet, the cash sheet, or the Wise/Revolut
    // totals. Each step is wrapped so a logging problem never rolls back
    // the extension itself, which has already been saved above.
    var bikeModel = sheet.getRange(rowNumber, CUSTOMER_BIKE_COL).getValue();
    var custName = sheet.getRange(rowNumber, CUSTOMER_NAME_COL).getValue();
    var incomeData = {
      bikeModel: bikeModel || '',
      name: custName || '',
      totalPrice: amountPaid,
      paidBy: paidBy,
      source: 'extend'
    };

    var warnings = [];

    try {
      appendMonthlyIncomeRow(ss, incomeData, daysToExtend);
    } catch (incomeErr) {
      warnings.push('Income sheet: ' + incomeErr.message);
    }

    try {
      if (paidBy.toLowerCase() === 'cash') {
        appendCashSheetRow(ss, incomeData, daysToExtend);
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    try {
      var paidByLower = paidBy.toLowerCase();
      if (paidByLower === 'wise' || paidByLower === 'revolut') {
        processDepositForPayment(ss, paidByLower, amountPaid);
      }
    } catch (depositErr) {
      warnings.push('Deposit total: ' + depositErr.message);
    }

    try {
      addRentalAmountToBikesSheet(ss, bikeModel, amountPaid);
    } catch (bikesErr) {
      warnings.push('Bikes sheet: ' + bikesErr.message);
    }

    // Post-write verification: re-read everything this extension wrote
    // and confirm it actually landed where it should.
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    var responsePayload = { success: true };
    responsePayload.checksPassed = verification.checked - verification.failed;
    responsePayload.checksTotal = verification.checked;
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Update one row in the customer tab: called from the "Extend" flow
// on the Bikes Status page when the extension is long (1-month checkbox
// ticked, or 30+ days typed in). Rather than pushing the due date on this
// row, that flow closes this booking out — using its current due date,
// left untouched, as the point it ended — and starts a brand-new rental
// record for the extension period instead. This just flips "situation" to
// "Returned" and normalizes the return date's and return time's font color
// (they're sometimes red from earlier conditional/manual formatting, same
// fix markBikeReturned applies); it deliberately does NOT change the
// return date's value, unlike markBikeReturned(). ----
function closeBikeForExtend(data) {
  try {
    var rowNumber = parseInt(data.rowNumber, 10);
    if (!rowNumber || rowNumber < 2) {
      throw new Error('Invalid row number.');
    }

    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var CUSTOMER_RETURN_DATE_COL = 9;  // I: Return date
    var CUSTOMER_RETURN_TIME_COL = 10; // J: Return time
    var CUSTOMER_SITUATION_COL = 14;   // N: situation

    sheet.getRange(rowNumber, CUSTOMER_RETURN_DATE_COL).setFontColor('#000000');
    sheet.getRange(rowNumber, CUSTOMER_RETURN_TIME_COL).setFontColor('#000000');
    sheet.getRange(rowNumber, CUSTOMER_SITUATION_COL).setValue('Returned');

    verifyCell('customer', rowNumber, CUSTOMER_SITUATION_COL, 'Returned', 'situation (close for extend)');

    var responsePayload = { success: true };
    var verification = runWriteVerification(ss);
    if (verification.problems.length) responsePayload.warning = verification.problems.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Bike Photos: Drive-backed storage ----
// One subfolder per bike, named exactly as the bike appears in the Parts
// and Oil change tab, sitting inside PHOTOS_ROOT_FOLDER_ID. Created lazily
// the first time a photo is uploaded for that bike.

function getOrCreateBikeFolder(bikeName) {
  var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
  var name = (bikeName || '').toString().trim();
  if (!name) throw new Error('No bike name given.');

  var existing = root.getFoldersByName(name);
  if (existing.hasNext()) return existing.next();

  return root.createFolder(name);
}

function fileToPhotoObject(file) {
  return {
    id: file.getId(),
    name: file.getName(),
    url: 'https://drive.google.com/thumbnail?id=' + file.getId() + '&sz=w1000'
  };
}

// ---- doPost: action = 'uploadPhoto' ----
// Expects: { action, bike, filename, mimeType, base64 }
function uploadBikePhoto(data) {
  try {
    if (!data.bike) throw new Error('No bike specified.');
    if (!data.base64) throw new Error('No image data received.');

    var folder = getOrCreateBikeFolder(data.bike);
    var bytes = Utilities.base64Decode(data.base64);
    var blob = Utilities.newBlob(bytes, data.mimeType || 'image/jpeg', data.filename || 'photo.jpg');

    var file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, photo: fileToPhotoObject(file) }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- doPost: action = 'deletePhoto' ----
// Expects: { action, fileId }
// Moves the file to Trash rather than permanently deleting it, so an
// accidental tap can still be recovered from Drive's Trash if needed.
function deleteBikePhoto(data) {
  try {
    if (!data.fileId) throw new Error('No file specified.');
    var file = DriveApp.getFileById(data.fileId);
    file.setTrashed(true);

    return ContentService
      .createTextOutput(JSON.stringify({ success: true }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- doGet: action = 'bikePhotos', param bike=<name> ----
// Returns { success, photos: [{id, name, url}, ...] }. If the bike has no
// folder yet (no photos ever uploaded), returns an empty list rather than
// an error.
function getBikePhotos(bikeName) {
  try {
    var name = (bikeName || '').toString().trim();
    if (!name) throw new Error('No bike specified.');

    var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
    var folders = root.getFoldersByName(name);
    if (!folders.hasNext()) {
      return ContentService
        .createTextOutput(JSON.stringify({ success: true, photos: [] }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    var folder = folders.next();
    var files = folder.getFiles();
    var photos = [];
    while (files.hasNext()) {
      var f = files.next();
      photos.push(fileToPhotoObject(f));
    }
    // Most recently added first.
    photos.reverse();

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, photos: photos }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- doGet: action = 'photoFolders' ----
// Returns { success, folders: [{name, count}] } — every subfolder under
// PHOTOS_ROOT_FOLDER_ID (one per bike that's ever had a photo uploaded)
// with how many files are inside. Lets a page check photo coverage across
// every bike in one round trip instead of one bikePhotos call per bike.
function getPhotoFolders() {
  try {
    var root = DriveApp.getFolderById(PHOTOS_ROOT_FOLDER_ID);
    var folders = root.getFolders();
    var result = [];
    while (folders.hasNext()) {
      var folder = folders.next();
      var count = 0;
      var files = folder.getFiles();
      while (files.hasNext()) { files.next(); count++; }
      result.push({ name: folder.getName(), count: count });
    }

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, folders: result }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function columnLetter(col) {
  var letter = '';
  while (col > 0) {
    var rem = (col - 1) % 26;
    letter = String.fromCharCode(65 + rem) + letter;
    col = Math.floor((col - 1) / 26);
  }
  return letter;
}

// ---- Used by the "Search" screen on the Customer Record page ----

var HEADER_ROWS = 1;

function doGet(e) {
  try {
    if (e.parameter.action === 'parts') {
      return getPartsData();
    }
    if (e.parameter.action === 'bikePhotos') {
      return getBikePhotos(e.parameter.bike);
    }
    if (e.parameter.action === 'photoFolders') {
      return getPhotoFolders();
    }
    if (e.parameter.action === 'accounts') {
      return getAccountsData(e.parameter.month);
    }
    if (e.parameter.action === 'bikesList') {
      return getBikesListNames();
    }

    // ---- Customer-search behavior ----
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName('customer');
    if (!sheet) {
      throw new Error('Sheet named "customer" not found in this spreadsheet.');
    }

    var values = sheet.getDataRange().getValues();
    var keys = ['timestamp','contact','name','nationality','passport','bikeModel',
                'status','rentingDateFrom','returnDate','returnTime','deliverToHotel',
                'totalPrice','paidBy','situation'];
    var tz = ss.getSpreadsheetTimeZone();

    function cellToString(key, val) {
      if (val instanceof Date) {
        if (key === 'returnTime') {
          return Utilities.formatDate(val, tz, 'HH:mm');
        }
        return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
      }
      return val !== undefined && val !== null ? String(val) : '';
    }

    var rows = values.slice(HEADER_ROWS).map(function(row, i) {
      var obj = {};
      keys.forEach(function(k, ki) {
        obj[k] = cellToString(k, row[ki]);
      });
      obj.rowNumber = HEADER_ROWS + i + 1; // 1-indexed sheet row this record lives on
      return obj;
    }).filter(function(r) { return r.name !== ''; });

    return ContentService
      .createTextOutput(JSON.stringify({ success: true, rows: rows }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Serve the Parts and Oil change tab, PLUS the Operation tab's Bike +
// Status columns (operationRows), PLUS the Bike Tax tab's Bike model +
// category columns (categoryRows), so pages can price/status bikes without
// needing separate round trips. ----
function getPartsData() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
    if (!sheet) {
      throw new Error('Sheet named "' + PARTS_SHEET_NAME + '" not found in this spreadsheet.');
    }

    var values = sheet.getDataRange().getValues();
    var tz = ss.getSpreadsheetTimeZone();
    var rawHeaders = values[0];

    var headers = rawHeaders.map(function(h, i) {
      var key = (h || '').toString().trim();
      return key || ('Column ' + columnLetter(i + 1));
    });

    function cellToString(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
      }
      return val !== undefined && val !== null ? String(val) : '';
    }

    // Read strikethrough formatting on the bike-name column (col 1) so the
    // client can tell "sold, struck through on purpose" bikes apart from
    // bikes that are just missing/mismatched data. getFontLines() returns
    // "line-through" or "none" per cell.
    var lastRow = sheet.getLastRow();
    var strikeArray = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, 1).getFontLines() : [];

    var rows = values.slice(1).map(function(row, i) {
      var obj = {};
      headers.forEach(function(h, i2) {
        obj[h] = cellToString(row[i2]);
      });
      obj.__struck = !!(strikeArray[i] && strikeArray[i][0] === 'line-through');
      return obj;
    }).filter(function(r) { return (r[headers[0]] || '').toString().trim() !== ''; });

    var operationRows = getOperationStatusRows();
    var categoryRows = getBikeTaxCategories();

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        headers: headers,
        rows: rows,
        operationRows: operationRows,
        categoryRows: categoryRows
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Read just the Bike + Status columns from the Operation tab.
// Looks up the columns by header text ("Bike" / "Status") rather than fixed
// column letters, so it keeps working if columns get added or reordered.
// Returns [] (rather than throwing) if the tab or columns are missing, so a
// problem here never breaks the rest of the Oil Change page. ----
function getOperationStatusRows() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(OPERATION_SHEET_NAME);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values.length) return [];

    var headerRow = values[0].map(function(h) {
      return (h || '').toString().trim().toLowerCase();
    });
    var bikeCol = headerRow.indexOf('bike');
    var statusCol = headerRow.indexOf('status');
    if (bikeCol === -1 || statusCol === -1) return [];

    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var bike = (values[i][bikeCol] || '').toString().trim();
      var status = (values[i][statusCol] || '').toString().trim();
      if (bike) rows.push({ bike: bike, status: status });
    }
    return rows;

  } catch (err) {
    return [];
  }
}

var ACCOUNTS_MONTH_NAMES = ['January','February','March','April','May','June',
  'July','August','September','October','November','December'];

// ---- Plain Levenshtein edit distance, used only to tolerate typos in a
// month tab's name (e.g. "Feburary" should still resolve to February). ----
function levenshteinDistance(a, b) {
  a = a || ''; b = b || '';
  var m = a.length, n = b.length;
  var dp = [];
  for (var i = 0; i <= m; i++) { dp.push([i]); }
  for (var j = 0; j <= n; j++) { dp[0][j] = j; }
  for (i = 1; i <= m; i++) {
    for (j = 1; j <= n; j++) {
      if (a.charAt(i - 1) === b.charAt(j - 1)) {
        dp[i][j] = dp[i - 1][j - 1];
      } else {
        dp[i][j] = 1 + Math.min(dp[i - 1][j - 1], dp[i - 1][j], dp[i][j - 1]);
      }
    }
  }
  return dp[m][n];
}

// ---- Scores how well a (normalized, lowercase) sheet name matches a given
// full month name. 0 = confident match: exact, or the sheet name is just an
// abbreviation/prefix of the month ("jan", "sept", "June" all score 0
// against their month). Otherwise falls back to edit distance against the
// full name, to tolerate spelling mistakes ("Feburary"). ----
function monthMatchScore(sheetNameNorm, fullMonthNameLower) {
  if (!sheetNameNorm) return 99;
  if (sheetNameNorm === fullMonthNameLower) return 0;
  if (fullMonthNameLower.indexOf(sheetNameNorm) === 0 && sheetNameNorm.length >= 3) return 0;
  return levenshteinDistance(sheetNameNorm, fullMonthNameLower);
}

// ---- Finds the sheet (visible or hidden) whose name best matches the given
// full month name ("July"), tolerating abbreviations ("Jul") and minor
// spelling mistakes. Returns null if nothing is a close enough match. ----
function findMonthSheetFuzzy(ss, fullMonthName) {
  var target = fullMonthName.toLowerCase();
  var sheets = ss.getSheets();
  var best = null;
  var bestScore = 99;
  for (var i = 0; i < sheets.length; i++) {
    var name = sheets[i].getName();
    var norm = name.toString().trim().toLowerCase();
    var score = monthMatchScore(norm, target);
    if (score < bestScore) {
      bestScore = score;
      best = sheets[i];
    }
  }
  // Allow up to a 2-character edit distance so small typos still resolve,
  // but not so loose that an unrelated sheet ("cash", "bikes") gets matched.
  if (best && bestScore <= 2) return best;
  return null;
}

// ---- Expense classification: each expense entry can be tagged as
// Business, Personal, Wages/Bike Purchase, or To Transfer. Rather than
// storing the tag anywhere new, it's represented purely as the background
// color of the expense description cell (column B) -- Business is the
// default (no fill), so old, never-classified entries read back as
// Business automatically with nothing to migrate. ----
var EXPENSE_TYPE_COLORS = {
  business: null,       // no fill -- the default, unclassified look
  personal: '#cfe2f3',  // light blue
  wages: '#f6b26b',     // orange
  transfer: '#fff2cc'   // light yellow
};

// ---- Colors the expense description cell (column B) of the given row to
// match its classification. Passing an unrecognized/blank type falls back
// to Business (no fill), so a row is never left with a stray color from a
// previous classification if the new save doesn't specify one. ----
function applyExpenseTypeColor(sheet, row, type) {
  var key = (type || 'business').toString().trim().toLowerCase();
  var color = EXPENSE_TYPE_COLORS.hasOwnProperty(key) ? EXPENSE_TYPE_COLORS[key] : null;
  sheet.getRange(row, 2).setBackground(color);
}

// ---- The reverse of applyExpenseTypeColor -- reads a cell's background
// color back and maps it to a classification, so getAccountsData can tell
// the client which type an existing entry already has (to prefill the
// dropdown on Edit). Anything that isn't one of the three colored types
// (including plain white/no-fill) reads back as Business. ----
function expenseTypeFromColor(hex) {
  var h = (hex || '').toString().trim().toLowerCase();
  if (h === EXPENSE_TYPE_COLORS.personal) return 'personal';
  if (h === EXPENSE_TYPE_COLORS.wages) return 'wages';
  if (h === EXPENSE_TYPE_COLORS.transfer) return 'transfer';
  return 'business';
}

// ---- Optional "which bike(s) is this expense split across" link, stored
// as a cell NOTE on the expense description cell (column B) -- not a new
// column, and not a stored row-number reference to another sheet either
// (the design explicitly avoided for cash-linking, since a stale row
// number can silently point at the wrong thing after a manual edit
// elsewhere). A note travels with its own cell -- if this row shifts
// (insert/delete elsewhere on this same sheet), the note moves with it
// automatically; if the row is deleted, the note simply goes with it.
//
// The note holds a JSON array of {bike, amount} pairs -- e.g. a single
// parts purchase covering several bikes can split ฿1,000 to "GT Black"
// and ฿1,500 to "GT Red" in one expense entry. Empty/no note means "not
// attached to any bike", the common case (most expenses aren't
// bike-specific). For backward compatibility, a note that's just a plain
// bike name (from before splitting existed) is treated as a single split
// covering the row's whole amount. ----
function parseExpenseBikeSplitsNote(note, fallbackAmount) {
  var trimmed = (note || '').toString().trim();
  if (!trimmed) return [];

  function cleanAmount(raw) {
    return (raw === '' || raw === null || raw === undefined || isNaN(Number(raw))) ? '' : Number(raw);
  }

  if (trimmed.charAt(0) === '[') {
    try {
      var parsed = JSON.parse(trimmed);
      if (!Array.isArray(parsed)) return [];
      return parsed
        .map(function(s) {
          return { bike: (s && s.bike || '').toString().trim(), amount: cleanAmount(s && s.amount) };
        })
        .filter(function(s) { return s.bike && s.amount !== ''; });
    } catch (e) {
      return []; // Malformed note -- treat as no link rather than throwing.
    }
  }

  // Legacy plain-bike-name note (pre-splitting) -- the row's whole amount
  // was attributed to that one bike.
  var fallback = cleanAmount(fallbackAmount);
  return fallback === '' ? [] : [{ bike: trimmed, amount: fallback }];
}

function getExpenseBikeSplits(sheet, row, fallbackAmount) {
  return parseExpenseBikeSplitsNote(sheet.getRange(row, 2).getNote(), fallbackAmount);
}

function setExpenseBikeSplits(sheet, row, splits) {
  var clean = (splits || [])
    .map(function(s) {
      var bike = (s && s.bike || '').toString().trim();
      var amt = (s && s.amount !== '' && s.amount !== null && s.amount !== undefined && !isNaN(Number(s.amount)))
        ? Number(s.amount) : '';
      return { bike: bike, amount: amt };
    })
    .filter(function(s) { return s.bike && s.amount !== ''; });
  sheet.getRange(row, 2).setNote(clean.length ? JSON.stringify(clean) : '');
}

// ---- Fixed labeled rows further down each month sheet ("personal
// expenses total" around row 149, "wages and bike purchase" around row
// 150) that track a running total of just that type of expense, as a
// formula chaining together the AMOUNT CELL of every matching expense --
// e.g. "=C29+C50+C55" -- rather than copying amounts across. Referencing
// the source cells directly means editing an expense's amount later
// updates these totals automatically, with nothing to reconcile here.
// Business expenses aren't broken out this way at all (they only count
// toward the overall "total expenses" row already on the sheet). ----
var EXPENSE_TYPE_TOTAL_LABELS = {
  personal: { row: 149, label: 'personal expenses total' },
  wages: { row: 150, label: 'wages and bike purchase' }
};

// ---- Finds the row (in column B) for a Personal/Wages running total,
// self-healing via findDepositRow (already used for the Wise/Revolut
// deposit cells) if the row has drifted from its usual spot. Returns null
// for Business/anything unrecognized, which isn't totalled this way. ----
function locateExpenseTypeTotalRow(sheet, type) {
  var def = EXPENSE_TYPE_TOTAL_LABELS[type];
  if (!def) return null;
  return findDepositRow(sheet, def.row, 2, def.label);
}

// ---- Adds (add=true) or removes (add=false) a cell reference like "C29"
// to/from the running "=C29+C50+..." formula in column C of the Personal/
// Wages total row. Verifies -- right before writing -- that column B
// still holds the expected label, same safety net already used before
// writing the Wise/Revolut deposit totals; throws instead of writing into
// the wrong cell if it doesn't match. Safe to call repeatedly -- adding an
// already-present reference, or removing one that's not there, is a
// no-op. Does nothing for Business/anything unrecognized. ----
function updateExpenseTypeTotalRef(sheet, type, refRow, add) {
  var def = EXPENSE_TYPE_TOTAL_LABELS[type];
  if (!def) return;

  var row = locateExpenseTypeTotalRow(sheet, type);
  if (row === null) {
    throw new Error('Could not find the "' + def.label + '" row in column B of "' + sheet.getName() +
      '" -- the ' + type + ' total was NOT updated.');
  }

  var actualLabel = (sheet.getRange(row, 2).getValue() || '').toString().trim().toLowerCase();
  if (actualLabel !== def.label) {
    throw new Error('Safety check failed: ' + sheet.getName() + '!B' + row + ' does not say "' + def.label +
      '" (found "' + actualLabel + '" instead) -- the ' + type + ' total was NOT updated. The row may have moved again.');
  }

  var ref = 'C' + refRow;
  var cell = sheet.getRange(row, 3);
  var formula = cell.getFormula();
  var terms;
  if (formula && formula.charAt(0) === '=') {
    terms = formula.slice(1).split('+').map(function(s) { return s.trim(); }).filter(function(s) { return s; });
  } else {
    // Not a formula yet -- if a plain number is already sitting there
    // (manually entered before this feature existed), keep it as the
    // first term rather than silently discarding it.
    var currentValue = cell.getValue();
    terms = (currentValue !== '' && currentValue !== null && !isNaN(Number(currentValue)) && Number(currentValue) !== 0)
      ? [String(Number(currentValue))]
      : [];
  }

  if (add) {
    if (terms.indexOf(ref) === -1) terms.push(ref);
  } else {
    terms = terms.filter(function(t) { return t !== ref; });
  }

  cell.setFormula(terms.length ? '=' + terms.join('+') : '');
}

// ---- Shared with getAccountsData, getAccountsFreeRow, and the
// add/edit-expense/income functions below, so the "where does the real
// data end and the totals block begin" rule stays identical everywhere
// it's used. Flags a label as a summary/totals line -- "total expenses",
// "income for month", "net profit", "% of ...", etc. ----
function looksLikeSummaryLabel(raw) {
  var t = (raw || '').toString().trim().toLowerCase();
  if (!t) return false;
  if (t.indexOf('total') === 0) return true; // "total expenses", "total income", ...
  var phrases = [
    'income for month', 'income for the month', 'income less',
    'bussiness expense', 'business expense', 'personal expense',
    'wages and bike', 'net profit', 'actual profit', '% of'
  ];
  for (var p = 0; p < phrases.length; p++) {
    if (t.indexOf(phrases[p]) !== -1) return true;
  }
  return false;
}

// ---- Serve one month's sheet for the Accounts page: the expense list
// (A Date, B expense, C amount) and the income list (E Date, F Income,
// G name, H Amount, I paid). monthIndexRaw is 0 (January) - 11 (December);
// defaults to the current month if missing/invalid. The sheet itself is
// located by fuzzy-matching its name against the month, so abbreviated or
// slightly misspelled tab names ("Jan", "Feburary") still resolve, and
// hidden tabs are included since Apps Script can read them regardless.
// Each returned row includes its real sheet row number so the Accounts
// page can send it back on addExpense/editExpense/addIncome/editIncome. ----
function getAccountsData(monthIndexRaw) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var tz = ss.getSpreadsheetTimeZone();

    var currentIndex = Number(Utilities.formatDate(new Date(), tz, 'M')) - 1;
    var monthIndex = currentIndex;
    if (monthIndexRaw !== undefined && monthIndexRaw !== null && monthIndexRaw !== '' && !isNaN(Number(monthIndexRaw))) {
      monthIndex = Math.max(0, Math.min(11, Math.round(Number(monthIndexRaw))));
    }
    var targetMonthName = ACCOUNTS_MONTH_NAMES[monthIndex];

    var sheet = findMonthSheetFuzzy(ss, targetMonthName);
    if (!sheet) {
      return ContentService
        .createTextOutput(JSON.stringify({
          success: false,
          monthIndex: monthIndex,
          month: targetMonthName,
          error: 'No sheet found matching "' + targetMonthName + '".'
        }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    function cellToString(val) {
      if (val instanceof Date) {
        return Utilities.formatDate(val, tz, 'dd/MM/yyyy');
      }
      return val !== undefined && val !== null ? val : '';
    }

    var lastRow = sheet.getLastRow();

    // Reads columns A-J together, row by row, and stops for good the moment
    // either side's label looks like a summary/totals line (see
    // looksLikeSummaryLabel above). That's the real boundary between the
    // actual expense/income rows and the personal totals block further
    // down the sheet, so once it's hit, nothing after it (on EITHER side)
    // is included, even if that side's own column still has data lower
    // down. Two consecutive fully-blank rows are also treated as the end,
    // as a fallback for months whose totals block doesn't use recognizable
    // label text.
    //
    // Column layout: A date, B expense, C amount, D payment (expense side),
    // E is a blank "CHK" spacer column (ignored), F date, G income,
    // H name, I amount, J paid (income side).
    var expenses = [];
    var income = [];
    if (lastRow >= 2) {
      var combined = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
      var expenseColColors = sheet.getRange(2, 2, lastRow - 1, 1).getBackgrounds();
      var expenseColNotes = sheet.getRange(2, 2, lastRow - 1, 1).getNotes();
      var prevBlank = false;
      for (var idx = 0; idx < combined.length; idx++) {
        var r = combined[idx];
        var eDate = r[0], expenseLabel = (r[1] || '').toString().trim(), eAmount = r[2];
        var ePayment = (r[3] || '').toString().trim();
        var iDate = r[5], incomeLabel = (r[6] || '').toString().trim(), iAmount = r[8];

        if (looksLikeSummaryLabel(expenseLabel) || looksLikeSummaryLabel(incomeLabel)) break;

        var eDateEmpty = (eDate === '' || eDate === null);
        var eAmountEmpty = (eAmount === '' || eAmount === null);
        var iDateEmpty = (iDate === '' || iDate === null);
        var iAmountEmpty = (iAmount === '' || iAmount === null);

        var rowFullyBlank = eDateEmpty && !expenseLabel && eAmountEmpty &&
          iDateEmpty && !incomeLabel && iAmountEmpty;
        if (rowFullyBlank) {
          if (prevBlank) break; // two in a row -- treat as the end
          prevBlank = true;
          continue;
        }
        prevBlank = false;

        var sheetRow = idx + 2; // combined[] is 0-based starting at sheet row 2

        if (expenseLabel || !eAmountEmpty) {
          expenses.push({
            row: sheetRow,
            date: cellToString(eDate),
            expense: expenseLabel,
            amount: (eAmountEmpty || isNaN(Number(eAmount))) ? '' : Number(eAmount),
            payment: ePayment,
            type: expenseTypeFromColor(expenseColColors[idx] && expenseColColors[idx][0]),
            bikeSplits: parseExpenseBikeSplitsNote(
              expenseColNotes[idx] && expenseColNotes[idx][0],
              (eAmountEmpty || isNaN(Number(eAmount))) ? '' : Number(eAmount)
            )
          });
        }
        if (incomeLabel || !iAmountEmpty) {
          income.push({
            row: sheetRow,
            date: cellToString(iDate),
            income: incomeLabel,
            name: (r[7] || '').toString().trim(),
            amount: (iAmountEmpty || isNaN(Number(iAmount))) ? '' : Number(iAmount),
            paidBy: (r[9] || '').toString().trim()
          });
        }
      }
    }

    return ContentService
      .createTextOutput(JSON.stringify({
        success: true,
        monthIndex: monthIndex,
        month: targetMonthName,
        sheetName: sheet.getName(),
        expenses: expenses,
        income: income
      }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Shared month-sheet lookup for the four accounts write functions
// below. Accepts either a monthIndex (0-11) or falls back to the current
// month, same rule as getAccountsData. ----
function locateAccountsSheet(monthIndexRaw) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var tz = ss.getSpreadsheetTimeZone();
  var currentIndex = Number(Utilities.formatDate(new Date(), tz, 'M')) - 1;
  var monthIndex = currentIndex;
  if (monthIndexRaw !== undefined && monthIndexRaw !== null && monthIndexRaw !== '' && !isNaN(Number(monthIndexRaw))) {
    monthIndex = Math.max(0, Math.min(11, Math.round(Number(monthIndexRaw))));
  }
  var targetMonthName = ACCOUNTS_MONTH_NAMES[monthIndex];
  var sheet = findMonthSheetFuzzy(ss, targetMonthName);
  if (!sheet) throw new Error('No sheet found matching "' + targetMonthName + '".');
  return sheet;
}

// ---- Checks whether one side (expense or income) of a given combined
// row (as returned by the getRange(...,10) reads above) is safe to write a
// new entry into without erasing anything meaningful already there.
//
// Only the "content" cells (date/label/name/amount) have to be blank --
// a leftover value sitting alone in just the payment/paid-by cell (with
// everything else on that side blank) does NOT count as an occupied row.
// A stray payment-only cell isn't a real entry, and treating it as
// "occupied" was leaving permanent gap rows: new entries would skip past
// it to the next fully-blank row instead of reusing (and overwriting) it. ----
function isAccountsSideEmpty(r, side) {
  if (side === 'expense') {
    var eDateEmpty = (r[0] === '' || r[0] === null);
    var expenseLabel = (r[1] || '').toString().trim();
    var eAmountEmpty = (r[2] === '' || r[2] === null);
    return eDateEmpty && !expenseLabel && eAmountEmpty;
  }
  var iDateEmpty = (r[5] === '' || r[5] === null);
  var incomeLabel = (r[6] || '').toString().trim();
  var nameEmpty = !(r[7] || '').toString().trim();
  var iAmountEmpty = (r[8] === '' || r[8] === null);
  return iDateEmpty && !incomeLabel && nameEmpty && iAmountEmpty;
}

// ---- Finds a row that's safe to write a brand-new expense OR income
// entry into (side is 'expense' or 'income'), packing new entries in
// right alongside the other side rather than always adding a whole new
// row. Since expense and income rows don't need to line up 1-for-1, a row
// that already has an income entry but no expense (or vice versa) is
// reused for the new entry on the empty side -- this is what keeps
// entries "one after another" with no gaps on either side.
//
// If every row up to the totals block already has this side filled, and
// the data runs straight into a "total ..." style summary row with no
// gap, a fresh row is inserted directly above that summary row -- since
// it lands inside the range most SUM()-style total formulas already
// cover, existing formulas simply expand to include it. If no summary row
// is found either, the row right after the sheet's last used row is
// used.
//
// Returns { row, inserted }. inserted is true only for the "row inserted
// above the totals block" case above -- that's the one situation where
// EVERY other row on the sheet (both sides) shifts down by one, so the
// client can't just append its new entry locally; it needs to fall back to
// a full reload instead of trusting old row numbers it already has
// cached. ----
function getAccountsFreeRow(sheet, side) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return { row: 2, inserted: false };

  var combined = sheet.getRange(2, 1, lastRow - 1, 10).getValues();
  for (var idx = 0; idx < combined.length; idx++) {
    var r = combined[idx];
    var expenseLabel = (r[1] || '').toString().trim();
    var incomeLabel = (r[6] || '').toString().trim();

    if (looksLikeSummaryLabel(expenseLabel) || looksLikeSummaryLabel(incomeLabel)) {
      var summaryRow = idx + 2;
      sheet.insertRowBefore(summaryRow);
      return { row: summaryRow, inserted: true };
    }

    if (isAccountsSideEmpty(r, side)) {
      return { row: idx + 2, inserted: false };
    }
  }
  return { row: lastRow + 1, inserted: false };
}

// =====================================================================
// Cash-sheet reconciliation for edits/deletes.
//
// No stored reference is kept anywhere for this -- an earlier version of
// this feature stored a row-number link in helper columns (AG/AH) on each
// month sheet, but that's deliberately been dropped: if someone manually
// inserts/deletes a row on the "cash" sheet later (entirely possible, since
// a lot of this data is entered by hand), a stored row number would go
// stale and could silently point at the wrong entry. Instead, whenever an
// edit or delete needs to find an entry's matching "cash" sheet row, it's
// looked up fresh, right then, by matching description + amount. In
// practice this is unique (a specific bike rented for a specific length
// starting on a specific date can't happen twice; a specific expense
// bought for a specific bike on a specific day is much the same). On the
// rare chance it isn't unique, nothing is guessed at -- the match is
// handed back to the client as a list of candidates so the user can pick
// which one, and the request is resubmitted with that exact row. ----

// ---- Searches the "cash" sheet's expense side (E-G) or income side
// (A-C) for every row whose description + amount match what's given. ----
function findCashCandidates(ss, side, text, amount) {
  var sheet = ss.getSheetByName('cash');
  if (!sheet) return [];

  var dateCol = side === 'expense' ? 5 : 1;
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];

  var expectedText = (text || '').toString().trim();
  var expectedAmountNum = (amount === '' || amount === null || amount === undefined || isNaN(Number(amount))) ? null : Number(amount);
  var tz = ss.getSpreadsheetTimeZone();

  var values = sheet.getRange(2, dateCol, lastRow - 1, 3).getValues(); // [date, label, amount]
  var candidates = [];
  for (var i = 0; i < values.length; i++) {
    var rowLabel = (values[i][1] || '').toString().trim();
    var rowAmountRaw = values[i][2];
    var rowAmountNum = (rowAmountRaw === '' || rowAmountRaw === null || isNaN(Number(rowAmountRaw))) ? null : Number(rowAmountRaw);
    var amountMatches = (rowAmountNum === null && expectedAmountNum === null) ||
      (rowAmountNum !== null && expectedAmountNum !== null && rowAmountNum === expectedAmountNum);

    if (rowLabel === expectedText && amountMatches) {
      var rawDate = values[i][0];
      var dateDisplay = rawDate instanceof Date ? Utilities.formatDate(rawDate, tz, 'dd/MM/yyyy') : (rawDate || '').toString();
      candidates.push({
        row: i + 2,
        date: dateDisplay,
        text: rowLabel,
        amount: rowAmountNum === null ? '' : rowAmountNum
      });
    }
  }
  return candidates;
}

// ---- Works out which "cash" sheet row to reconcile against:
//   - If the client already resubmitted with an explicit choice
//     (data.cashRowChoice, after the user picked from a disambiguation
//     list), that's used directly.
//   - Else searches for candidates: zero -> null (not found), one -> that
//     row, more than one -> returns a special marker so the caller can
//     bail out and ask the user to choose, before anything is written. ----
function resolveCashRow(ss, side, cashRowChoice, text, amount) {
  if (cashRowChoice) return { row: Math.round(Number(cashRowChoice)) };

  var candidates = findCashCandidates(ss, side, text, amount);
  if (candidates.length === 0) return { row: null };
  if (candidates.length === 1) return { row: candidates[0].row };
  return { needsDisambiguation: true, candidates: candidates };
}

// ---- Confirms a "cash" sheet row still holds what we expect (description
// + amount) right before it's touched -- cheap last-moment insurance in
// case something changed between the lookup above and this write. ----
function cashRowStillMatches(cashSheet, cashRow, labelCol, amountCol, expectedText, expectedAmount) {
  if (!cashRow || cashRow < 2) return false;
  var actualText = (cashSheet.getRange(cashRow, labelCol).getValue() || '').toString().trim();
  var actualAmountRaw = cashSheet.getRange(cashRow, amountCol).getValue();
  var expectedAmountNum = (expectedAmount === '' || expectedAmount === null || expectedAmount === undefined) ? null : Number(expectedAmount);
  var actualAmountNum = (actualAmountRaw === '' || actualAmountRaw === null) ? null : Number(actualAmountRaw);
  var amountMatches = (actualAmountNum === null && expectedAmountNum === null) ||
    (actualAmountNum !== null && expectedAmountNum !== null && actualAmountNum === expectedAmountNum);
  return actualText === (expectedText || '').toString().trim() && amountMatches;
}

// ---- Deletes a "cash" sheet row's 3 cells (date/description/amount) and
// shifts everything below it, in just those 3 columns, up by one -- same
// "delete cells, shift up" scoping as deleteExpenseRow/deleteIncomeRow
// below use on the month sheet itself. side is 'expense' (cash sheet
// columns E-G) or 'income' (columns A-C). ----
function deleteCashRow(ss, cashRow, side, expectedText, expectedAmount) {
  var sheet = ss.getSheetByName('cash');
  if (!sheet) throw new Error('"cash" sheet not found.');
  var dateCol = side === 'expense' ? 5 : 1;
  var labelCol = side === 'expense' ? 6 : 2;
  var amountCol = side === 'expense' ? 7 : 3;

  if (!cashRowStillMatches(sheet, cashRow, labelCol, amountCol, expectedText, expectedAmount)) {
    throw new Error('Could not confirm "cash" sheet row ' + cashRow + ' still matches this entry -- it was NOT removed. Please check/remove it manually if needed.');
  }
  sheet.getRange(cashRow, dateCol, 1, 3).deleteCells(SpreadsheetApp.Dimension.ROWS);

  // Post-write verification: the deleted entry's description should no
  // longer be sitting at this cash row (whatever shifted up replaces it).
  verifyCellChanged('cash', cashRow, labelCol, expectedText, 'deleted cash sheet row');
}

// ---- Updates a "cash" sheet row's description + amount in place (no
// shifting -- this is an edit, not a delete). ----
function updateCashRow(ss, cashRow, side, expectedOldText, expectedOldAmount, newText, newAmount) {
  var sheet = ss.getSheetByName('cash');
  if (!sheet) throw new Error('"cash" sheet not found.');
  var labelCol = side === 'expense' ? 6 : 2;
  var amountCol = side === 'expense' ? 7 : 3;

  if (!cashRowStillMatches(sheet, cashRow, labelCol, amountCol, expectedOldText, expectedOldAmount)) {
    throw new Error('Could not confirm "cash" sheet row ' + cashRow + ' still matches this entry -- it was NOT updated. Please check/update it manually if needed.');
  }
  sheet.getRange(cashRow, labelCol).setValue(newText);
  var amountValue = (newAmount === '' || newAmount === undefined || newAmount === null || isNaN(Number(newAmount))) ? '' : Number(newAmount);
  sheet.getRange(cashRow, amountCol).setValue(amountValue);

  verifyCell('cash', cashRow, labelCol, newText, 'updated cash sheet row: description');
  verifyCell('cash', cashRow, amountCol, amountValue, 'updated cash sheet row: amount');
}

// ---- action:'addExpense' -- data: { monthIndex, date (yyyy-MM-dd),
// expense, amount, payment }. Writes into a fresh row's A/B/C/D columns
// only, leaving whatever is in that row's income columns (F-J) untouched.
//
// Then, exactly like a manual income entry does in addIncomeRow below,
// routes based on payment:
//   Cash -> also logged as its own row on the "cash" sheet's expense side
//           (columns E-G).
//   Bank (or anything else) -> nothing further -- Bank is the expense-side
//           equivalent of Scan on the income side; the A-D row is enough.
// Wrapped so a problem there never rolls back the expense row itself,
// which is already saved by this point -- any issue comes back as a
// non-fatal "warning" instead. ----
function addExpenseRow(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = locateAccountsSheet(data.monthIndex);
    var freeRow = getAccountsFreeRow(sheet, 'expense');
    var row = freeRow.row;
    sheet.getRange(row, 1).setValue(formatIsoDateToDMY(data.date) || '');
    sheet.getRange(row, 2).setValue(data.expense || '');
    sheet.getRange(row, 3).setValue(
      (data.amount === '' || data.amount === undefined || data.amount === null || isNaN(Number(data.amount)))
        ? '' : Number(data.amount)
    );
    sheet.getRange(row, 4).setValue(data.payment || '');
    verifyCell(sheet.getName(), row, 1, formatIsoDateToDMY(data.date) || '', 'expense row: date');
    verifyCell(sheet.getName(), row, 2, data.expense || '', 'expense row: description');
    verifyCell(sheet.getName(), row, 3,
      (data.amount === '' || data.amount === undefined || data.amount === null || isNaN(Number(data.amount)))
        ? '' : Number(data.amount), 'expense row: amount');
    verifyCell(sheet.getName(), row, 4, data.payment || '', 'expense row: payment');
    applyExpenseTypeColor(sheet, row, data.expenseType);
    var expenseBikeSplits = Array.isArray(data.expenseBikeSplits) ? data.expenseBikeSplits : [];
    setExpenseBikeSplits(sheet, row, expenseBikeSplits);

    var warnings = [];
    var paymentLower = (data.payment || '').toString().trim().toLowerCase();
    var expenseTypeKey = (data.expenseType || 'business').toString().trim().toLowerCase();

    try {
      if (paymentLower === 'cash') {
        appendCashExpenseRowText(ss, data.expense || '', data.amount);
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    try {
      if (expenseTypeKey === 'personal' || expenseTypeKey === 'wages') {
        updateExpenseTypeTotalRef(sheet, expenseTypeKey, row, true);
      }
    } catch (typeErr) {
      warnings.push('Expense type total: ' + typeErr.message);
    }

    // If this expense is split across one or more bikes, add each split's
    // amount into that bike's cell in the "bikes" sheet's EXPENSE table
    // (starting at BIKES_EXPENSE_SECTION_START_ROW), for this accounts
    // sheet's own month -- same running-formula approach as the income
    // side. Each split is wrapped independently so one bad bike name
    // doesn't stop the rest from being applied.
    var bikeWarnings = [];
    expenseBikeSplits.forEach(function(s) {
      var bike = (s && s.bike || '').toString().trim();
      var amt = Number(s && s.amount);
      if (!bike || s.amount === '' || isNaN(amt)) return;
      try {
        addRentalAmountToBikesSheet(ss, bike, amt, sheet.getName(), BIKES_EXPENSE_SECTION_START_ROW);
      } catch (bikeErr) {
        bikeWarnings.push(bikeErr.message);
      }
    });
    if (bikeWarnings.length) warnings.push('Bikes sheet (expense): ' + bikeWarnings.join(' '));

    // Post-write verification: re-read everything this add wrote and
    // confirm it actually landed where it should.
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    // shifted: true only in the rare case a row had to be inserted above
    // the totals block, which pushes every other row on the sheet down by
    // one -- the client can't safely trust cached row numbers after that,
    // so it should do a full reload instead of a local-only update.
    var responsePayload = { success: true, row: row, shifted: freeRow.inserted };
    responsePayload.checksPassed = verification.checked - verification.failed;
    responsePayload.checksTotal = verification.checked;
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- action:'editExpense' -- data: { monthIndex, row, date, expense,
// amount, payment, cashRowChoice (optional) }. Overwrites A/B/C/D on the
// given (already-existing) row only.
//
// Snapshots the OLD description/amount/payment BEFORE writing anything, so
// the payment routing can be reconciled against the NEW values:
//   Cash -> Cash (amount or description changed) -- updates the matching
//           "cash" row in place.
//   Cash -> Bank -- removes the matching "cash" row.
//   Bank -> Cash -- adds a brand-new "cash" row.
//   Bank -> Bank -- nothing to reconcile.
// The old "cash" row (when needed) is looked up fresh by matching
// description + amount via resolveCashRow() -- see the comment above that
// section for why there's no stored reference. If that search finds more
// than one candidate, NOTHING is written yet -- this returns
// needsDisambiguation + the candidate list so the client can ask the user
// which one, then resubmit with cashRowChoice set. Every reconciliation
// step is wrapped so a problem there never rolls back the expense edit
// itself, which is already saved by this point. ----
function editExpenseRow(data) {
  try {
    if (!data.row || isNaN(Number(data.row))) throw new Error('Missing row number to edit.');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = locateAccountsSheet(data.monthIndex);
    var row = Math.round(Number(data.row));

    var oldExpense = (sheet.getRange(row, 2).getValue() || '').toString().trim();
    var oldAmountRaw = sheet.getRange(row, 3).getValue();
    var oldAmount = (oldAmountRaw === '' || oldAmountRaw === null || isNaN(Number(oldAmountRaw))) ? '' : Number(oldAmountRaw);
    var oldPaymentLower = (sheet.getRange(row, 4).getValue() || '').toString().trim().toLowerCase();
    var oldTypeKey = expenseTypeFromColor(sheet.getRange(row, 2).getBackground());
    var oldBikeSplits = getExpenseBikeSplits(sheet, row, oldAmount);

    var newAmount = (data.amount === '' || data.amount === undefined || data.amount === null || isNaN(Number(data.amount)))
      ? '' : Number(data.amount);
    var newPaymentLower = (data.payment || '').toString().trim().toLowerCase();
    var newExpenseText = (data.expense || '').toString().trim();
    var newBikeSplits = Array.isArray(data.expenseBikeSplits) ? data.expenseBikeSplits : [];

    var wasCash = oldPaymentLower === 'cash';
    var isCash = newPaymentLower === 'cash';

    // Phase 1: resolve which "cash" row (if any) needs to be updated/
    // removed, BEFORE touching this sheet -- so an ambiguous match can
    // bail out cleanly with nothing written.
    var resolvedOldCashRow = null;
    if (wasCash) {
      var resolution = resolveCashRow(ss, 'expense', data.cashRowChoice, oldExpense, oldAmount);
      if (resolution.needsDisambiguation) {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            needsDisambiguation: true,
            candidates: resolution.candidates
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      resolvedOldCashRow = resolution.row;
    }

    // Phase 2: apply the edit now that reconciliation is resolved.
    sheet.getRange(row, 1).setValue(formatIsoDateToDMY(data.date) || '');
    sheet.getRange(row, 2).setValue(data.expense || '');
    sheet.getRange(row, 3).setValue(newAmount);
    sheet.getRange(row, 4).setValue(data.payment || '');
    verifyCell(sheet.getName(), row, 1, formatIsoDateToDMY(data.date) || '', 'edited expense row: date');
    verifyCell(sheet.getName(), row, 2, data.expense || '', 'edited expense row: description');
    verifyCell(sheet.getName(), row, 3, newAmount, 'edited expense row: amount');
    verifyCell(sheet.getName(), row, 4, data.payment || '', 'edited expense row: payment');
    applyExpenseTypeColor(sheet, row, data.expenseType);
    setExpenseBikeSplits(sheet, row, newBikeSplits);
    var newTypeKey = (data.expenseType || 'business').toString().trim().toLowerCase();

    var warnings = [];

    // Reconcile the "bikes" sheet expense table -- subtract each OLD split
    // from whichever bike it was attached to, then add each NEW split to
    // whichever bike it's attached to now. Same bike, amount-only change
    // -> nets out to the delta on that one cell. Different bike (fixing a
    // misattributed expense) -> comes off the old bike's cell and goes
    // onto the new one's. A bike dropped from the split, or a new one
    // added, is naturally handled too, since old and new are reconciled
    // independently. Each split is wrapped on its own so one bad bike
    // name doesn't stop the rest from being reconciled.
    var oldBikeWarnings = [];
    oldBikeSplits.forEach(function(s) {
      try {
        addRentalAmountToBikesSheet(ss, s.bike, -s.amount, sheet.getName(), BIKES_EXPENSE_SECTION_START_ROW);
      } catch (e) {
        oldBikeWarnings.push(e.message);
      }
    });
    if (oldBikeWarnings.length) warnings.push('Bikes sheet (removing old expense): ' + oldBikeWarnings.join(' '));

    var newBikeWarnings = [];
    newBikeSplits.forEach(function(s) {
      var bike = (s && s.bike || '').toString().trim();
      var amt = Number(s && s.amount);
      if (!bike || s.amount === '' || isNaN(amt)) return;
      try {
        addRentalAmountToBikesSheet(ss, bike, amt, sheet.getName(), BIKES_EXPENSE_SECTION_START_ROW);
      } catch (e) {
        newBikeWarnings.push(e.message);
      }
    });
    if (newBikeWarnings.length) warnings.push('Bikes sheet (adding new expense): ' + newBikeWarnings.join(' '));

    // Reconcile the Personal/Wages running totals if the classification
    // changed -- an amount-only change needs nothing here, since the
    // total formula references this row's amount cell directly and picks
    // up the new value on its own.
    if (oldTypeKey !== newTypeKey) {
      try {
        if (oldTypeKey === 'personal' || oldTypeKey === 'wages') {
          updateExpenseTypeTotalRef(sheet, oldTypeKey, row, false);
        }
        if (newTypeKey === 'personal' || newTypeKey === 'wages') {
          updateExpenseTypeTotalRef(sheet, newTypeKey, row, true);
        }
      } catch (typeErr) {
        warnings.push('Expense type total: ' + typeErr.message);
      }
    }

    try {
      if (wasCash && isCash) {
        if (resolvedOldCashRow) {
          updateCashRow(ss, resolvedOldCashRow, 'expense', oldExpense, oldAmount, newExpenseText, newAmount);
        } else {
          appendCashExpenseRowText(ss, newExpenseText, newAmount);
          warnings.push('Could not find a matching "cash" sheet row for this entry -- a NEW cash row was added for the updated amount instead. Please check the "cash" sheet for a possible duplicate.');
        }
      } else if (wasCash && !isCash) {
        if (resolvedOldCashRow) {
          deleteCashRow(ss, resolvedOldCashRow, 'expense', oldExpense, oldAmount);
        } else {
          warnings.push('Could not find a matching "cash" sheet row for this entry -- if it logged one, please remove it manually.');
        }
      } else if (!wasCash && isCash) {
        appendCashExpenseRowText(ss, newExpenseText, newAmount);
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    // Post-write verification: re-read everything this edit wrote and
    // confirm it actually landed where it should.
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    var responsePayload = { success: true, row: row };
    responsePayload.checksPassed = verification.checked - verification.failed;
    responsePayload.checksTotal = verification.checked;
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Builds the text logged on the "cash" sheet for a manual Accounts
// income entry -- there's no bikeModel/dayCount to build off of here (that's
// only for rental income). Matches the rental cash-row shape exactly: just
// date + description + amount, no name column, since the "cash" sheet only
// has three columns (A date, B description, C amount). ----
function buildGeneralIncomeText(data) {
  return (data.income || '').toString().trim();
}

// ---- action:'addIncome' -- data: { monthIndex, date, income, name,
// amount, paidBy }. Writes into a fresh row's F/G/H/I/J columns only,
// leaving whatever is in that row's expense columns (A-D) untouched.
//
// Then, exactly like a new customer rental (or an extension) does in
// doPost/extendBikeRow above, routes the payment based on paidBy:
//   Cash    -> also logged as its own row on the "cash" sheet.
//   Wise    -> added into the running Wise deposit total (M11).
//   Revolut -> added into the running Revolut deposit total (M12).
//   Scan (or anything else) -> nothing further; the F-J row is enough.
// Each routing step is wrapped so a problem there never rolls back the
// income row itself, which is already saved by this point -- any issue is
// returned as a non-fatal "warning" instead. ----
function addIncomeRow(data) {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = locateAccountsSheet(data.monthIndex);
    var freeRow = getAccountsFreeRow(sheet, 'income');
    var row = freeRow.row;
    sheet.getRange(row, 6).setValue(formatIsoDateToDMY(data.date) || '');
    sheet.getRange(row, 7).setValue(data.income || '');
    sheet.getRange(row, 8).setValue(data.name || '');
    sheet.getRange(row, 9).setValue(
      (data.amount === '' || data.amount === undefined || data.amount === null || isNaN(Number(data.amount)))
        ? '' : Number(data.amount)
    );
    sheet.getRange(row, 10).setValue(data.paidBy || '');
    verifyCell(sheet.getName(), row, 6, formatIsoDateToDMY(data.date) || '', 'income row: date');
    verifyCell(sheet.getName(), row, 7, data.income || '', 'income row: description');
    verifyCell(sheet.getName(), row, 8, data.name || '', 'income row: name');
    verifyCell(sheet.getName(), row, 9,
      (data.amount === '' || data.amount === undefined || data.amount === null || isNaN(Number(data.amount)))
        ? '' : Number(data.amount), 'income row: amount');
    verifyCell(sheet.getName(), row, 10, data.paidBy || '', 'income row: paid by');

    var warnings = [];
    var paidByLower = (data.paidBy || '').toString().trim().toLowerCase();

    try {
      if (paidByLower === 'cash') {
        appendCashSheetRowText(ss, buildGeneralIncomeText(data), data.amount);
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    try {
      if (paidByLower === 'wise' || paidByLower === 'revolut') {
        processDepositForPayment(ss, paidByLower, data.amount);
      }
    } catch (depositErr) {
      warnings.push('Deposit total: ' + depositErr.message);
    }

    // Post-write verification: re-read everything this add wrote and
    // confirm it actually landed where it should.
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    // shifted: true only in the rare case a row had to be inserted above
    // the totals block -- see addExpenseRow's comment on the same field.
    var responsePayload = { success: true, row: row, shifted: freeRow.inserted };
    responsePayload.checksPassed = verification.checked - verification.failed;
    responsePayload.checksTotal = verification.checked;
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- action:'editIncome' -- data: { monthIndex, row, date, income, name,
// amount, paidBy, cashRowChoice (optional) }. Overwrites F/G/H/I/J on the
// given (already-existing) row only.
//
// Snapshots the OLD income text/amount/paidBy BEFORE writing anything, so
// both the "cash" sheet and the Wise/Revolut running totals can be
// reconciled against the NEW values:
//   Cash -> Cash (amount or description changed) -- updates the matching
//           "cash" row in place.
//   Cash -> something else -- removes the matching "cash" row.
//   something else -> Cash -- adds a brand-new "cash" row.
//   Wise/Revolut involved (before and/or after) -- subtracts the OLD amount
//           from whichever method it used to be under, then adds the NEW
//           amount to whichever method it's under now (same method, a
//           different one, or none -- each half only fires if relevant, so
//           this naturally covers an amount-only change too).
// The old "cash" row (when needed) is looked up fresh by matching
// description + amount via resolveCashRow() -- see the comment above that
// section for why there's no stored reference. If that search finds more
// than one candidate, NOTHING is written yet -- this returns
// needsDisambiguation + the candidate list so the client can ask the user
// which one, then resubmit with cashRowChoice set. Every reconciliation
// step is wrapped so a problem there never rolls back the income edit
// itself, which is already saved by this point. ----
function editIncomeRow(data) {
  try {
    if (!data.row || isNaN(Number(data.row))) throw new Error('Missing row number to edit.');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = locateAccountsSheet(data.monthIndex);
    var row = Math.round(Number(data.row));

    var oldIncome = (sheet.getRange(row, 7).getValue() || '').toString().trim();
    var oldAmountRaw = sheet.getRange(row, 9).getValue();
    var oldAmount = (oldAmountRaw === '' || oldAmountRaw === null || isNaN(Number(oldAmountRaw))) ? '' : Number(oldAmountRaw);
    var oldPaidByLower = (sheet.getRange(row, 10).getValue() || '').toString().trim().toLowerCase();
    var oldGeneralText = buildGeneralIncomeText({ income: oldIncome });

    var newAmount = (data.amount === '' || data.amount === undefined || data.amount === null || isNaN(Number(data.amount)))
      ? '' : Number(data.amount);
    var newPaidByLower = (data.paidBy || '').toString().trim().toLowerCase();
    var newGeneralText = buildGeneralIncomeText(data);

    var wasCash = oldPaidByLower === 'cash';
    var isCash = newPaidByLower === 'cash';

    // Phase 1: resolve which "cash" row (if any) needs to be updated/
    // removed, BEFORE touching this sheet -- so an ambiguous match can
    // bail out cleanly with nothing written.
    var resolvedOldCashRow = null;
    if (wasCash) {
      var resolution = resolveCashRow(ss, 'income', data.cashRowChoice, oldGeneralText, oldAmount);
      if (resolution.needsDisambiguation) {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            needsDisambiguation: true,
            candidates: resolution.candidates
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      resolvedOldCashRow = resolution.row;
    }

    // Phase 2: apply the edit now that reconciliation is resolved.
    sheet.getRange(row, 6).setValue(formatIsoDateToDMY(data.date) || '');
    sheet.getRange(row, 7).setValue(data.income || '');
    sheet.getRange(row, 8).setValue(data.name || '');
    sheet.getRange(row, 9).setValue(newAmount);
    sheet.getRange(row, 10).setValue(data.paidBy || '');
    verifyCell(sheet.getName(), row, 6, formatIsoDateToDMY(data.date) || '', 'edited income row: date');
    verifyCell(sheet.getName(), row, 7, data.income || '', 'edited income row: description');
    verifyCell(sheet.getName(), row, 8, data.name || '', 'edited income row: name');
    verifyCell(sheet.getName(), row, 9, newAmount, 'edited income row: amount');
    verifyCell(sheet.getName(), row, 10, data.paidBy || '', 'edited income row: paid by');

    var warnings = [];

    // ---- Cash reconciliation ----
    try {
      if (wasCash && isCash) {
        if (resolvedOldCashRow) {
          updateCashRow(ss, resolvedOldCashRow, 'income', oldGeneralText, oldAmount, newGeneralText, newAmount);
        } else {
          appendCashSheetRowText(ss, newGeneralText, newAmount);
          warnings.push('Could not find a matching "cash" sheet row for this entry -- a NEW cash row was added for the updated amount instead. Please check the "cash" sheet for a possible duplicate.');
        }
      } else if (wasCash && !isCash) {
        if (resolvedOldCashRow) {
          deleteCashRow(ss, resolvedOldCashRow, 'income', oldGeneralText, oldAmount);
        } else {
          warnings.push('Could not find a matching "cash" sheet row for this entry -- if it logged one, please remove it manually.');
        }
      } else if (!wasCash && isCash) {
        appendCashSheetRowText(ss, newGeneralText, newAmount);
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    // ---- Wise/Revolut reconciliation ----
    try {
      if (oldPaidByLower === 'wise' || oldPaidByLower === 'revolut') {
        processDepositForPayment(ss, oldPaidByLower, -(oldAmount === '' ? 0 : oldAmount));
      }
    } catch (revertErr) {
      warnings.push('Deposit total (removing old amount): ' + revertErr.message);
    }
    try {
      if (newPaidByLower === 'wise' || newPaidByLower === 'revolut') {
        processDepositForPayment(ss, newPaidByLower, (newAmount === '' ? 0 : newAmount));
      }
    } catch (applyErr) {
      warnings.push('Deposit total (adding new amount): ' + applyErr.message);
    }

    // ---- "bikes" sheet reconciliation ----
    // If this income row is a bike rental/extension line (built by
    // buildRentalIncomeText when the booking was first created on the
    // customer page), that bike also has a running per-month total on the
    // "bikes" sheet -- same subtract-old/add-new treatment as the
    // Wise/Revolut totals just above, using this accounts sheet's OWN
    // month (not necessarily today's real month, since a past month can
    // be edited via the Accounts page's month selector). A manually-typed
    // "Add Income" entry that doesn't look like a rental line is left
    // alone entirely -- there's nothing on the "bikes" sheet to reconcile.
    var accountsMonthName = sheet.getName();
    try {
      var oldBikeName = extractBikeNameFromRentalIncomeText(oldIncome);
      if (oldBikeName && oldAmount !== '') {
        addRentalAmountToBikesSheet(ss, oldBikeName, -oldAmount, accountsMonthName);
      }
    } catch (bikeRevertErr) {
      warnings.push('Bikes sheet (removing old amount): ' + bikeRevertErr.message);
    }
    try {
      var newBikeName = extractBikeNameFromRentalIncomeText(data.income);
      if (newBikeName && newAmount !== '') {
        addRentalAmountToBikesSheet(ss, newBikeName, newAmount, accountsMonthName);
      }
    } catch (bikeApplyErr) {
      warnings.push('Bikes sheet (adding new amount): ' + bikeApplyErr.message);
    }

    // Post-write verification: re-read everything this edit wrote and
    // confirm it actually landed where it should.
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    var responsePayload = { success: true, row: row };
    responsePayload.checksPassed = verification.checked - verification.failed;
    responsePayload.checksTotal = verification.checked;
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- action:'deleteExpense' -- data: { monthIndex, row, cashRowChoice
// (optional) }. Reverses whatever the entry originally logged (a "cash"
// sheet row, if it was paid in cash), then deletes just the A-D cells for
// this row and shifts everything below -- in those columns only -- up by
// one. The income side (F-J) of this same physical row, if any, is left
// completely untouched.
//
// If payment was Cash, the matching "cash" row is looked up fresh by
// description + amount via resolveCashRow(). If that finds more than one
// candidate, NOTHING is deleted yet -- this returns needsDisambiguation +
// the candidates so the client can ask the user which one, then resubmit
// with cashRowChoice set. ----
function deleteExpenseRow(data) {
  try {
    if (!data.row || isNaN(Number(data.row))) throw new Error('Missing row number to delete.');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = locateAccountsSheet(data.monthIndex);
    var row = Math.round(Number(data.row));

    var expense = (sheet.getRange(row, 2).getValue() || '').toString().trim();
    var amountRaw = sheet.getRange(row, 3).getValue();
    var amount = (amountRaw === '' || amountRaw === null || isNaN(Number(amountRaw))) ? '' : Number(amountRaw);
    var paymentLower = (sheet.getRange(row, 4).getValue() || '').toString().trim().toLowerCase();
    var typeKey = expenseTypeFromColor(sheet.getRange(row, 2).getBackground());
    var expenseBikeSplits = getExpenseBikeSplits(sheet, row, amount);

    var resolvedCashRow = null;
    if (paymentLower === 'cash') {
      var resolution = resolveCashRow(ss, 'expense', data.cashRowChoice, expense, amount);
      if (resolution.needsDisambiguation) {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            needsDisambiguation: true,
            candidates: resolution.candidates
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      resolvedCashRow = resolution.row;
    }

    var warnings = [];

    try {
      if (paymentLower === 'cash') {
        if (resolvedCashRow) {
          deleteCashRow(ss, resolvedCashRow, 'expense', expense, amount);
        } else {
          warnings.push('Could not find a matching "cash" sheet row for this entry -- if it logged one, please remove it manually.');
        }
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    // Remove this row's reference from the Personal/Wages running total
    // BEFORE the row's cells are deleted below -- once deleteCells shifts
    // everything up, "row" no longer belongs to this entry, so the
    // reference has to come out first while it's still correct.
    try {
      if (typeKey === 'personal' || typeKey === 'wages') {
        updateExpenseTypeTotalRef(sheet, typeKey, row, false);
      }
    } catch (typeErr) {
      warnings.push('Expense type total: ' + typeErr.message);
    }

    // Same reason -- if this expense was split across one or more bikes,
    // remove each split's contribution from that bike's cell in the
    // "bikes" sheet's expense table before the row's cells (and the note
    // carrying this link) are deleted below.
    var deleteBikeWarnings = [];
    expenseBikeSplits.forEach(function(s) {
      try {
        addRentalAmountToBikesSheet(ss, s.bike, -s.amount, sheet.getName(), BIKES_EXPENSE_SECTION_START_ROW);
      } catch (e) {
        deleteBikeWarnings.push(e.message);
      }
    });
    if (deleteBikeWarnings.length) warnings.push('Bikes sheet (expense): ' + deleteBikeWarnings.join(' '));

    // Delete just A:D for this row, shifting everything below -- in those
    // columns only -- up. The income side (F-J) of this same row, if any,
    // is left completely alone.
    sheet.getRange(row, 1, 1, 4).deleteCells(SpreadsheetApp.Dimension.ROWS);

    // Post-write verification: the deleted entry's description should no
    // longer be sitting at this row (whatever shifted up replaces it).
    verifyCellChanged(sheet.getName(), row, 2, expense, 'deleted expense row');
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    var responsePayload = { success: true };
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- action:'deleteIncome' -- data: { monthIndex, row, cashRowChoice
// (optional) }. Reverses whatever the entry originally did (a "cash" sheet
// row if paid in cash, or the running Wise/Revolut deposit total if paid
// that way), then deletes just the F-J cells for this row and shifts
// everything below -- in those columns only -- up by one. The expense side
// (A-D) of this same physical row, if any, is left completely untouched.
//
// If payment was Cash, the matching "cash" row is looked up fresh by
// description + amount via resolveCashRow(). If that finds more than one
// candidate, NOTHING is deleted yet -- this returns needsDisambiguation +
// the candidates so the client can ask the user which one, then resubmit
// with cashRowChoice set. ----
function deleteIncomeRow(data) {
  try {
    if (!data.row || isNaN(Number(data.row))) throw new Error('Missing row number to delete.');
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = locateAccountsSheet(data.monthIndex);
    var row = Math.round(Number(data.row));

    var income = (sheet.getRange(row, 7).getValue() || '').toString().trim();
    var amountRaw = sheet.getRange(row, 9).getValue();
    var amount = (amountRaw === '' || amountRaw === null || isNaN(Number(amountRaw))) ? '' : Number(amountRaw);
    var paidByLower = (sheet.getRange(row, 10).getValue() || '').toString().trim().toLowerCase();

    var resolvedCashRow = null;
    if (paidByLower === 'cash') {
      var resolution = resolveCashRow(ss, 'income', data.cashRowChoice, income, amount);
      if (resolution.needsDisambiguation) {
        return ContentService
          .createTextOutput(JSON.stringify({
            success: false,
            needsDisambiguation: true,
            candidates: resolution.candidates
          }))
          .setMimeType(ContentService.MimeType.JSON);
      }
      resolvedCashRow = resolution.row;
    }

    var warnings = [];

    try {
      if (paidByLower === 'cash') {
        if (resolvedCashRow) {
          deleteCashRow(ss, resolvedCashRow, 'income', income, amount);
        } else {
          warnings.push('Could not find a matching "cash" sheet row for this entry -- if it logged one, please remove it manually.');
        }
      }
    } catch (cashErr) {
      warnings.push('Cash sheet: ' + cashErr.message);
    }

    try {
      if (paidByLower === 'wise' || paidByLower === 'revolut') {
        processDepositForPayment(ss, paidByLower, -(amount === '' ? 0 : amount));
      }
    } catch (depositErr) {
      warnings.push('Deposit total: ' + depositErr.message);
    }

    // If this was a bike rental/extension line, reverse its contribution
    // to that bike's running per-month total on the "bikes" sheet too --
    // same reasoning as editIncomeRow's reconciliation above (this
    // accounts sheet's own month, not necessarily today's real month).
    try {
      var deletedBikeName = extractBikeNameFromRentalIncomeText(income);
      if (deletedBikeName && amount !== '') {
        addRentalAmountToBikesSheet(ss, deletedBikeName, -amount, sheet.getName());
      }
    } catch (bikeErr) {
      warnings.push('Bikes sheet: ' + bikeErr.message);
    }

    // Delete just F:J for this row, shifting everything below -- in those
    // columns only -- up. The expense side (A-D) of this same row, if any,
    // is left completely alone.
    sheet.getRange(row, 6, 1, 5).deleteCells(SpreadsheetApp.Dimension.ROWS);

    // Post-write verification: the deleted entry's description should no
    // longer be sitting at this row (whatever shifted up replaces it).
    verifyCellChanged(sheet.getName(), row, 7, income, 'deleted income row');
    var verification = runWriteVerification(ss);
    warnings = warnings.concat(verification.problems);

    var responsePayload = { success: true };
    if (warnings.length) responsePayload.warning = warnings.join(' ');

    return ContentService
      .createTextOutput(JSON.stringify(responsePayload))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    return ContentService
      .createTextOutput(JSON.stringify({ success: false, error: err.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ---- Read Bike model + category + make/model/cc/key from the Bike Tax tab.
// Looks columns up by header text so it keeps working if columns move.
// make/model/cc/key are optional — any that are blank or missing just come
// back as empty strings, since not every bike has them filled in yet.
// Returns [] (never throws) so a problem here never breaks the rest of the
// Available Bikes page. ----
function getBikeTaxCategories() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(BIKE_TAX_SHEET_NAME);
    if (!sheet) return [];

    var values = sheet.getDataRange().getValues();
    if (!values.length) return [];

    var headerRow = values[0].map(function(h) {
      return (h || '').toString().trim().toLowerCase();
    });
    var bikeCol = headerRow.indexOf('bike model');
    if (bikeCol === -1) bikeCol = headerRow.indexOf('bike');
    var catCol = headerRow.indexOf('category');
    if (bikeCol === -1 || catCol === -1) return [];

    var makeCol = headerRow.indexOf('make');
    // Look for model/cc/key/deposit starting from the make column onward, so
    // these always resolve to the group of columns sitting next to "make" —
    // not some unrelated column elsewhere in the sheet that happens to
    // share a name.
    var searchFrom = makeCol > -1 ? makeCol : 0;
    var modelCol = headerRow.indexOf('model', searchFrom);
    var ccCol = headerRow.indexOf('cc', searchFrom);
    var keyCol = headerRow.indexOf('key', searchFrom);
    var depositCol = headerRow.indexOf('deposit', searchFrom);
    if (depositCol === -1) depositCol = headerRow.indexOf('deposit');
    var boxCol = headerRow.indexOf('box', searchFrom);
    if (boxCol === -1) boxCol = headerRow.indexOf('box');
    var absCol = headerRow.indexOf('abs', searchFrom);
    if (absCol === -1) absCol = headerRow.indexOf('abs');
    var tractionCol = headerRow.indexOf('traction control', searchFrom);
    if (tractionCol === -1) tractionCol = headerRow.indexOf('traction control');

    // Plate number lives near the front of the sheet (e.g. "Plate No."),
    // not next to make/model, so it's looked up across the whole header
    // row rather than from searchFrom onward. Matched by "contains 'plate'"
    // rather than an exact string so header punctuation (e.g. the period
    // in "Plate No.") doesn't break the lookup.
    var plateCol = -1;
    for (var pc = 0; pc < headerRow.length; pc++) {
      if (headerRow[pc].indexOf('plate') !== -1) { plateCol = pc; break; }
    }

    var rows = [];
    for (var i = 1; i < values.length; i++) {
      var bike = (values[i][bikeCol] || '').toString().trim();
      if (!bike) continue;
      var cat = (values[i][catCol] || '').toString().trim();
      rows.push({
        bike: bike,
        category: cat,
        make: makeCol > -1 ? (values[i][makeCol] || '').toString().trim() : '',
        model: modelCol > -1 ? (values[i][modelCol] || '').toString().trim() : '',
        cc: ccCol > -1 ? (values[i][ccCol] || '').toString().trim() : '',
        key: keyCol > -1 ? (values[i][keyCol] || '').toString().trim() : '',
        deposit: depositCol > -1 ? (values[i][depositCol] || '').toString().trim() : '',
        box: boxCol > -1 ? (values[i][boxCol] || '').toString().trim() : '',
        abs: absCol > -1 ? (values[i][absCol] || '').toString().trim() : '',
        tractionControl: tractionCol > -1 ? (values[i][tractionCol] || '').toString().trim() : '',
        plate: plateCol > -1 ? (values[i][plateCol] || '').toString().trim() : ''
      });
    }
    return rows;

  } catch (err) {
    return [];
  }
}

// =====================================================================
// ONE-TIME MIGRATION: import photos from the "Cosmetic Damage" Google Doc
// into each bike's Drive photo folder. Each Doc tab (e.g. "RAX Red") is
// matched to a real bike name from the Parts and Oil change sheet, and
// every image embedded in that tab is copied into that bike's folder.
//
// HOW TO RUN:
//   1. In the Apps Script editor, pick "importCosmeticDamagePhotos" from
//      the function dropdown at the top (next to Run/Debug).
//   2. Click Run. First time, you'll get a permissions prompt for Google
//      Docs access — Advanced > Go to (project) (unsafe) > Allow.
//   3. When it finishes, go to View > Logs (or Executions) to see a
//      summary of what was imported per tab, and which tabs (if any)
//      couldn't be confidently matched to a bike name.
//   4. This only needs to be run once. Re-running it will import the
//      same photos again as duplicates, so don't run it twice unless
//      you've deleted the previous import first.
//
// Safe to delete this whole section afterward if you don't need it again.
// =====================================================================

var COSMETIC_DAMAGE_DOC_ID = '10YMe4YqkJHT94STf40J9fQqu2L01j2Qg0Skr3gLRcPY';

function normalizeBikeNameForImport(s) {
  return (s || '').toString()
    .toLowerCase()
    .replace(/\([^)]*\)/g, ' ')
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .replace(/\s+/g, ' ');
}

function bikeNamesMatchForImport(a, b) {
  var na = normalizeBikeNameForImport(a);
  var nb = normalizeBikeNameForImport(b);
  if (!na || !nb) return false;
  if (na === nb) return true;
  return na.indexOf(nb) !== -1 || nb.indexOf(na) !== -1;
}

function getAllBikeNamesForImport() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(PARTS_SHEET_NAME);
  if (!sheet) throw new Error('Sheet named "' + PARTS_SHEET_NAME + '" not found.');
  var values = sheet.getDataRange().getValues();
  var names = [];
  for (var i = 1; i < values.length; i++) {
    var name = (values[i][0] || '').toString().trim();
    if (name) names.push(name);
  }
  return names;
}

// Picks the best bike-name match for a Doc tab title. If several bike
// names loosely match, prefers whichever is closest in length to the tab
// title (i.e. the tightest match), rather than guessing at random.
function findBestBikeMatchForImport(tabTitle, bikeNames) {
  var candidates = bikeNames.filter(function(n) {
    return bikeNamesMatchForImport(n, tabTitle);
  });
  if (!candidates.length) return null;
  if (candidates.length === 1) return candidates[0];

  var targetLen = normalizeBikeNameForImport(tabTitle).length;
  candidates.sort(function(a, b) {
    var da = Math.abs(normalizeBikeNameForImport(a).length - targetLen);
    var db = Math.abs(normalizeBikeNameForImport(b).length - targetLen);
    return da - db;
  });
  return candidates[0];
}

// Recursively walks a Doc element tree (paragraphs, tables, table cells,
// etc.) collecting every inline image found anywhere inside it.
function collectInlineImagesForImport(element, out) {
  var type = element.getType();
  if (type === DocumentApp.ElementType.INLINE_IMAGE) {
    out.push(element.asInlineImage());
    return;
  }
  var numChildren;
  try {
    numChildren = element.getNumChildren();
  } catch (e) {
    return; // Not a container element (e.g. plain text run) — nothing to recurse into.
  }
  for (var i = 0; i < numChildren; i++) {
    collectInlineImagesForImport(element.getChild(i), out);
  }
}
