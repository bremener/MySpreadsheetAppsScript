/*
 * CLEAN COLUMN D QUOTES — SAFE ADD-ON
 *
 * Add this code to the end of your existing MonologueDatabaseSheetBuilder
 * Apps Script project. Then save and reload the Google Sheet.
 *
 * It adds two menu commands:
 * - Preview First Line quote cleanup
 * - Apply First Line quote cleanup
 *
 * Scope and safety:
 * - Processes only valid monologue database tabs.
 * - Explicitly excludes Manual Entries, Skipped files, and Source Changes Review.
 * - Changes only Column D (First Line).
 * - Removes quotation-mark characters anywhere in Column D.
 * - Does not change Column A:C, E:H, formulas in other columns, rich-text links,
 *   smart chips, row order, filters, or sheet structure.
 * - Preview makes no changes. Apply writes only cells whose visible text changes.
 */

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Monologue Database')
    .addItem('Scan folder for new entries', 'buildMonologueDatabase')
    .addItem('Scan all folder files safely', 'forceFullMonologueRescan')
    .addSeparator()
    .addItem('Import Manual Entries', 'importManualEntries')
    .addSeparator()
    .addItem('Preview Playwright cleanup', 'previewPlaywrightCleanup')
    .addItem('Apply Playwright cleanup', 'applyPlaywrightCleanup')
    .addItem('Preview Character cleanup', 'previewCharacterCleanup')
    .addItem('Apply Character cleanup', 'applyCharacterCleanup')
    .addSeparator()
    .addItem('Preview First Line quote cleanup', 'previewFirstLineQuoteCleanup')
    .addItem('Apply First Line quote cleanup', 'applyFirstLineQuoteCleanup')
    .addSeparator()
    .addItem('Preview database duplicate removal', 'previewDatabaseDuplicateRemoval')
    .addItem('Apply database duplicate removal', 'applyDatabaseDuplicateRemoval')
    .addSeparator()
    .addItem('Build Source Changes Review', 'buildSourceChangesReview')
    .addToUi();
}

function previewFirstLineQuoteCleanup() {
  runFirstLineQuoteCleanup_(true);
}

function applyFirstLineQuoteCleanup() {
  runFirstLineQuoteCleanup_(false);
}

function runFirstLineQuoteCleanup_(dryRun) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDestinationSpreadsheet_();
    const databaseSheets = ss.getSheets().filter(function(sheet) {
      return isMonologueDatabaseSheet_(sheet);
    });

    if (databaseSheets.length === 0) {
      throw new Error(
        'No database tabs with the expected A1:H1 headers were found.'
      );
    }

    const changes = [];

    databaseSheets.forEach(function(sheet) {
      const lastRow = sheet.getLastRow();

      if (lastRow < CONFIG.DATA_START_ROW) {
        return;
      }

      const rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
      const firstLineValues = sheet
        .getRange(
          CONFIG.DATA_START_ROW,
          4,
          rowCount,
          1
        )
        .getDisplayValues()
        .flat();

      firstLineValues.forEach(function(displayValue, index) {
        const original = cleanText_(displayValue);

        if (original === '') {
          return;
        }

        const cleaned = removeQuotationMarksFromFirstLine_(original);

        if (cleaned === original) {
          return;
        }

        changes.push({
          sheetName: sheet.getName(),
          rowNumber: CONFIG.DATA_START_ROW + index,
          original: original,
          cleaned: cleaned
        });
      });
    });

    if (changes.length === 0) {
      showMessage_(
        dryRun
          ? 'First Line quote cleanup preview complete: no cells need changes.'
          : 'First Line quote cleanup complete: no cells needed changes.'
      );
      return;
    }

    const previewLimit = 30;
    const previewLines = changes
      .slice(0, previewLimit)
      .map(function(change) {
        return (
          change.sheetName +
          '!D' +
          change.rowNumber +
          '\n  "' +
          change.original +
          '" → "' +
          change.cleaned +
          '"'
        );
      });

    const extraCount = changes.length - previewLines.length;
    Logger.log(JSON.stringify(changes, null, 2));

    if (dryRun) {
      showMessage_(
        'FIRST LINE QUOTE CLEANUP PREVIEW\n\n' +
        'Column D cells that would change: ' + changes.length + '\n\n' +
        previewLines.join('\n\n') +
        (extraCount > 0
          ? '\n\n…and ' + extraCount + ' more.'
          : '') +
        '\n\nNo cells were changed.\n\n' +
        'If this preview is correct, use:\n' +
        'Monologue Database → Apply First Line quote cleanup'
      );
      return;
    }

    changes.forEach(function(change) {
      ss
        .getSheetByName(change.sheetName)
        .getRange(change.rowNumber, 4)
        .setValue(change.cleaned);
    });

    SpreadsheetApp.flush();

    showMessage_(
      'FIRST LINE QUOTE CLEANUP COMPLETE\n\n' +
      'Updated Column D cells: ' + changes.length + '\n\n' +
      previewLines.join('\n\n') +
      (extraCount > 0
        ? '\n\n…and ' + extraCount + ' more.'
        : '')
    );
  } finally {
    lock.releaseLock();
  }
}

function removeQuotationMarksFromFirstLine_(value) {
  return cleanText_(value)
    .replace(/["“”„‟«»‹›‘’‚‛`´]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}
