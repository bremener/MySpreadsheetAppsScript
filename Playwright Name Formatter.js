/*
 * One-time playwright-name formatter.
 *
 * Run formatAllPlaywrightNamesLastFirst() directly from the Apps Script editor.
 * This script does not define onOpen(), so it will not conflict with the
 * Monologue Database Builder menu.
 */

const PLAYWRIGHT_FORMATTER = {
  PLAYWRIGHT_HEADER: 'playwright',

  // Add or remove tab names here as needed.
  SKIP_SHEET_NAMES: [
    'Skipped files',
    'Manual Entries',
    'Manual Entry Import Report',
    'Reconciliation Report',
    'Reconciliation Exceptions',
    '_Monologue Import Cache',
    '_Monologue Manual Data'
  ]
};

function formatAllPlaywrightNamesLastFirst() {
  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const updatedSheets = [];
  let updatedNames = 0;

  spreadsheet.getSheets().forEach(function(sheet) {
    if (shouldSkipPlaywrightFormatterSheet(sheet)) {
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastColumn = sheet.getLastColumn();

    if (lastRow < 2 || lastColumn < 1) {
      return;
    }

    const headers = sheet
      .getRange(1, 1, 1, lastColumn)
      .getDisplayValues()[0]
      .map(normalizePlaywrightFormatterHeader);

    const playwrightColumnIndex = headers.indexOf(
      PLAYWRIGHT_FORMATTER.PLAYWRIGHT_HEADER
    );

    if (playwrightColumnIndex === -1) {
      return;
    }

    const playwrightRange = sheet.getRange(
      2,
      playwrightColumnIndex + 1,
      lastRow - 1,
      1
    );

    const values = playwrightRange.getValues();
    let sheetChanged = false;

    const formattedValues = values.map(function(row) {
      const original = String(row[0] || '');
      const formatted = formatOnePlaywrightName(original);

      if (formatted !== original) {
        sheetChanged = true;
        updatedNames += 1;
      }

      return [formatted];
    });

    if (sheetChanged) {
      playwrightRange.setValues(formattedValues);
      updatedSheets.push(sheet.getName());
    }
  });

  SpreadsheetApp.flush();

  SpreadsheetApp.getUi().alert(
    'Playwright formatter complete. Updated ' +
    updatedNames +
    ' name' +
    (updatedNames === 1 ? '' : 's') +
    (updatedSheets.length > 0
      ? ' in: ' + updatedSheets.join(', ') + '.'
      : '. No eligible visible tabs required changes.')
  );
}

function shouldSkipPlaywrightFormatterSheet(sheet) {
  const sheetName = sheet.getName();

  return sheet.isSheetHidden() ||
    sheetName.charAt(0) === '_' ||
    PLAYWRIGHT_FORMATTER.SKIP_SHEET_NAMES.indexOf(sheetName) !== -1;
}

function formatOnePlaywrightName(playwright) {
  const name = normalizePlaywrightFormatterText(playwright);

  if (name === '') {
    return '';
  }

  // Preserve "Last, First" entries, standardizing only comma spacing.
  if (name.indexOf(',') !== -1) {
    return name
      .replace(/\s*,\s*/g, ', ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  const parts = name.split(/\s+/);

  if (parts.length < 2) {
    return name;
  }

  const lastName = parts.pop();
  const givenNames = parts.join(' ');

  return lastName + ', ' + givenNames;
}

function normalizePlaywrightFormatterHeader(text) {
  return normalizePlaywrightFormatterText(text)
    .toLowerCase()
    .replace(/[\-_]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizePlaywrightFormatterText(text) {
  return String(text || '')
    .replace(/\s+/g, ' ')
    .trim();
}