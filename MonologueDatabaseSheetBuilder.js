/*
 * MONOLOGUE DATABASE SHEET BUILDER — FULL FIXED VERSION
 *
 * Database tabs use A:H:
 * A Playwright | B Play / Mono Link | C Character | D First Line
 * E Identifiers | F Themes | G Considerations | H Link to Play
 *
 * Manual Entries uses A:J:
 * A:H same as database tabs | I Category | J Import Status
 */

const FOLDER_ID = '1b6Uqq3W17iKuAVs9SPoCooBF444RnCyY';
const SPREADSHEET_ID = '';

const CONFIG = {
  VISIBLE_COLUMNS: [
    'Playwright',
    'Play / Mono Link',
    'Character',
    'First Line',
    'Identifiers',
    'Themes',
    'Considerations',
    'Link to Play'
  ],
  DATA_START_ROW: 2,
  DATABASE_COLUMN_COUNT: 8,
  MANUAL_ENTRIES_SHEET_NAME: 'Manual Entries',
  MANUAL_CATEGORY_COLUMN: 9,
  MANUAL_STATUS_COLUMN: 10,
  SKIPPED_SHEET_NAME: 'Skipped files',
  SOURCE_CHANGES_REVIEW_SHEET_NAME: 'Source Changes Review',
  MAX_SHEET_NAME_LENGTH: 100,
  FIRST_LINE_WORD_COUNT: 10,
  EXCLUDE_PARENTHETICAL_STAGE_DIRECTIONS: true
};

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
    .addItem('Preview First Line quote cleanup', 'previewFirstLineQuoteCleanup')
    .addItem('Apply First Line quote cleanup', 'applyFirstLineQuoteCleanup')
    .addSeparator()
    .addItem('Review matching play titles one at a time', 'reviewMatchingPlayTitles')
    .addItem('Preview exact duplicate removal', 'previewDatabaseDuplicateRemoval')
    .addItem('Apply exact duplicate removal', 'applyDatabaseDuplicateRemoval')
    .addItem('Preview likely duplicate removal', 'previewLikelyDuplicateRemoval')
    .addItem('Apply likely duplicate removal', 'applyLikelyDuplicateRemoval')
    .addSeparator()
    .addItem('Build Source Changes Review', 'buildSourceChangesReview')
    .addItem(
  'Diagnose a Google Doc layout',
  'diagnoseGoogleDocLayout'
)
    .addSeparator()
    .addItem(
      'Build folder/database inventory',
      'buildFolderDatabaseInventory'
    )
    .addToUi();
}

function buildMonologueDatabase() {
  runSafeFolderScan_();
}

function forceFullMonologueRescan() {
  runSafeFolderScan_();
}

function runSafeFolderScan_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    validateFolderId_();
    const ss = getDestinationSpreadsheet_();
    const rootFolder = DriveApp.getFolderById(FOLDER_ID);
    const rootFolderName = rootFolder.getName();
    const outputSheet = getOrCreateSheet_(ss, makeSafeTabName_(rootFolderName));

    ensureDatabaseHeaders_(outputSheet);

    const databaseState = getDatabaseState_(outputSheet);
    const filesToProcess = [];
    collectFilesRecursively_(rootFolder, '', filesToProcess);

    const processedDocumentIds = new Set();
    const skippedFiles = [];
    const totals = { added: 0, duplicate: 0, unsupported: 0, review: 0 };

    filesToProcess.forEach(function(item) {
      try {
        const resolvedFile = resolveGoogleDocFile_(item.file);

        if (!resolvedFile) {
          skippedFiles.push({
            rootFolder: rootFolderName,
            sourceFolder: item.sourceFolder,
            fileName: item.file.getName(),
            fileUrl: item.file.getUrl(),
            reason: 'Not a native Google Doc or shortcut to a Google Doc'
          });
          totals.unsupported++;
          return;
        }

        const documentId = resolvedFile.getId();
        if (processedDocumentIds.has(documentId)) {
          totals.duplicate++;
          return;
        }
        processedDocumentIds.add(documentId);

        const info = readMonologueHeader_(resolvedFile);
        const candidateValues = [
          info.playwright,
          resolvedFile.getName(),
          info.character,
          info.firstLine,
          '', '', '', ''
        ];

        if (!hasRequiredIdentityFields_(candidateValues)) {
          skippedFiles.push({
            rootFolder: rootFolderName,
            sourceFolder: item.sourceFolder,
            fileName: resolvedFile.getName(),
            fileUrl: resolvedFile.getUrl(),
            reason: 'Could not obtain Playwright, Character, and First Line'
          });
          totals.review++;
          return;
        }

        const fingerprint = makeFingerprint_(candidateValues);
        if (
          databaseState.fingerprintRows.has(fingerprint) ||
          databaseState.ambiguousFingerprints.has(fingerprint)
        ) {
          totals.duplicate++;
          return;
        }

        appendFolderEntryAtRow_(
          outputSheet,
          databaseState.nextAppendRow,
          candidateValues,
          resolvedFile.getUrl()
        );

        addNewRowToDatabaseState_(
          databaseState,
          fingerprint,
          databaseState.nextAppendRow
        );
        databaseState.nextAppendRow++;
        totals.added++;
      } catch (error) {
        skippedFiles.push({
          rootFolder: rootFolderName,
          sourceFolder: item.sourceFolder,
          fileName: item.file.getName(),
          fileUrl: item.file.getUrl(),
          reason: error && error.message ? error.message : String(error)
        });
        totals.review++;
      }
    });

    updateSkippedFilesSheet_(ss, skippedFiles);
    SpreadsheetApp.flush();

    showMessage_(
      'Folder scan complete.\n\n' +
      'Added: ' + totals.added + '\n' +
      'Duplicates skipped: ' + totals.duplicate + '\n' +
      'Unsupported files: ' + totals.unsupported + '\n' +
      'Needs review: ' + totals.review
    );
  } finally {
    lock.releaseLock();
  }
}

function importManualEntries() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDestinationSpreadsheet_();
    const manualSheet = ss.getSheetByName(CONFIG.MANUAL_ENTRIES_SHEET_NAME);

    if (!manualSheet) {
      throw new Error(
        'Could not find a tab named "' + CONFIG.MANUAL_ENTRIES_SHEET_NAME + '".'
      );
    }

    ensureManualEntriesHeaders_(manualSheet);

    const lastRow = manualSheet.getLastRow();
    if (lastRow < CONFIG.DATA_START_ROW) {
      showMessage_('No manual entries were found.');
      return;
    }

    const rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
    const sourceDisplayRows = manualSheet
      .getRange(CONFIG.DATA_START_ROW, 1, rowCount, CONFIG.DATABASE_COLUMN_COUNT)
      .getDisplayValues();
    const categoryNames = manualSheet
      .getRange(CONFIG.DATA_START_ROW, CONFIG.MANUAL_CATEGORY_COLUMN, rowCount, 1)
      .getDisplayValues()
      .flat();
    const statuses = manualSheet
      .getRange(CONFIG.DATA_START_ROW, CONFIG.MANUAL_STATUS_COLUMN, rowCount, 1)
      .getDisplayValues()
      .flat();

    const categoryStates = new Map();
    const totals = {
      imported: 0,
      merged: 0,
      duplicateNoNewData: 0,
      review: 0,
      alreadyHandled: 0,
      blank: 0
    };

    for (let index = 0; index < rowCount; index++) {
      const sourceRow = CONFIG.DATA_START_ROW + index;
      const sourceDisplayValues = sourceDisplayRows[index];
      const categoryName = cleanText_(categoryNames[index]);
      const priorStatus = cleanText_(statuses[index]);

      if (!rowHasContent_(sourceDisplayValues)) {
        totals.blank++;
        continue;
      }

      if (
        priorStatus === 'Imported' ||
        priorStatus === 'Merged into duplicate' ||
        priorStatus === 'Duplicate — no new E:H data'
      ) {
        totals.alreadyHandled++;
        continue;
      }

      if (!hasRequiredIdentityFields_(sourceDisplayValues)) {
        setManualStatus_(
          manualSheet,
          sourceRow,
          'Needs review — complete Playwright, Character, and First Line'
        );
        totals.review++;
        continue;
      }

      if (categoryName === '') {
        setManualStatus_(manualSheet, sourceRow, 'Needs review — add Category');
        totals.review++;
        continue;
      }

      if (!categoryStates.has(categoryName)) {
        const targetSheet = ss.getSheetByName(categoryName);
        if (!targetSheet) {
          categoryStates.set(categoryName, {
            error: 'Needs review — category tab not found'
          });
        } else {
          ensureDatabaseHeaders_(targetSheet);
          categoryStates.set(categoryName, {
            targetSheet: targetSheet,
            databaseState: getDatabaseState_(targetSheet)
          });
        }
      }

      const categoryState = categoryStates.get(categoryName);
      if (categoryState.error) {
        setManualStatus_(manualSheet, sourceRow, categoryState.error);
        totals.review++;
        continue;
      }

      const fingerprint = makeFingerprint_(sourceDisplayValues);
      const databaseState = categoryState.databaseState;

      if (databaseState.ambiguousFingerprints.has(fingerprint)) {
        setManualStatus_(
          manualSheet,
          sourceRow,
          'Needs review — multiple duplicate matches'
        );
        totals.review++;
        continue;
      }

      if (databaseState.fingerprintRows.has(fingerprint)) {
        const targetRow = databaseState.fingerprintRows.get(fingerprint);
        const mergeResult = mergeMissingSupplementalData_(
          manualSheet,
          sourceRow,
          categoryState.targetSheet,
          targetRow
        );

        if (mergeResult.updatedCount > 0) {
          setManualStatus_(manualSheet, sourceRow, 'Merged into duplicate');
          totals.merged++;
        } else {
          setManualStatus_(manualSheet, sourceRow, 'Duplicate — no new E:H data');
          totals.duplicateNoNewData++;
        }
        continue;
      }

      appendManualSourceRowAtRow_(
        manualSheet,
        sourceRow,
        categoryState.targetSheet,
        databaseState.nextAppendRow
      );

      addNewRowToDatabaseState_(
        databaseState,
        fingerprint,
        databaseState.nextAppendRow
      );
      databaseState.nextAppendRow++;

      setManualStatus_(manualSheet, sourceRow, 'Imported');
      totals.imported++;
    }

    SpreadsheetApp.flush();
    showMessage_(
      'Manual import complete.\n\n' +
      'New rows imported: ' + totals.imported + '\n' +
      'Duplicates enriched: ' + totals.merged + '\n' +
      'Duplicates with no new E:H data: ' + totals.duplicateNoNewData + '\n' +
      'Needs review: ' + totals.review + '\n' +
      'Already handled: ' + totals.alreadyHandled + '\n' +
      'Blank rows: ' + totals.blank
    );
  } finally {
    lock.releaseLock();
  }
}

function getDatabaseDataRows_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return [];
  }

  return sheet
    .getRange(
      CONFIG.DATA_START_ROW,
      1,
      lastRow - CONFIG.DATA_START_ROW + 1,
      CONFIG.DATABASE_COLUMN_COUNT
    )
    .getDisplayValues();
}

function getSheetColumnDisplayValues_(sheet, column) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return [];
  }

  return sheet
    .getRange(CONFIG.DATA_START_ROW, column, lastRow - CONFIG.DATA_START_ROW + 1, 1)
    .getDisplayValues()
    .flat();
}

function getDatabaseState_(sheet) {
  const lastRow = Math.max(sheet.getLastRow(), CONFIG.DATA_START_ROW - 1);

  if (lastRow < CONFIG.DATA_START_ROW) {
    return {
      fingerprintRows: new Map(),
      ambiguousFingerprints: new Set(),
      nextAppendRow: CONFIG.DATA_START_ROW
    };
  }

  const rows = getDatabaseDataRows_(sheet);

  const fingerprintRows = new Map();
  const ambiguousFingerprints = new Set();
  let firstEmptyRow = null;

  rows.forEach(function(row, index) {
    const rowNumber = CONFIG.DATA_START_ROW + index;

    if (!rowHasContent_(row)) {
      if (firstEmptyRow === null) {
        firstEmptyRow = rowNumber;
      }
      return;
    }

    if (!hasRequiredIdentityFields_(row)) {
      return;
    }

    const fingerprint = makeFingerprint_(row);
    if (fingerprintRows.has(fingerprint)) {
      fingerprintRows.delete(fingerprint);
      ambiguousFingerprints.add(fingerprint);
      return;
    }

    if (!ambiguousFingerprints.has(fingerprint)) {
      fingerprintRows.set(fingerprint, rowNumber);
    }
  });

  return {
    fingerprintRows: fingerprintRows,
    ambiguousFingerprints: ambiguousFingerprints,
    nextAppendRow: firstEmptyRow === null ? lastRow + 1 : firstEmptyRow
  };
}

function addNewRowToDatabaseState_(databaseState, fingerprint, rowNumber) {
  if (databaseState.ambiguousFingerprints.has(fingerprint)) {
    return;
  }

  if (databaseState.fingerprintRows.has(fingerprint)) {
    databaseState.fingerprintRows.delete(fingerprint);
    databaseState.ambiguousFingerprints.add(fingerprint);
    return;
  }

  databaseState.fingerprintRows.set(fingerprint, rowNumber);
}

function makeFingerprint_(row) {
  return [
    normalizeForMatch_(row[0]),             // Playwright
    normalizeForMatch_(row[2]),             // Character
    firstMeaningfulWordsKey_(row[3], 6)    // First 6 words; comparison ignores punctuation
  ].join(' | ');
}

function makeLikelyDuplicateFingerprint_(row) {
  return [
    normalizeForMatch_(row[0]),
    normalizeForMatch_(row[2]),
    firstMeaningfulWordsKey_(row[3], 6)
  ].join(' | ');
}

/**
 * Comparison-only key for duplicate detection.
 *
 * It keeps word boundaries based on whitespace first, so contractions
 * remain one word:
 *
 * "I’m"       → one word
 * "don't"     → one word
 * "you’re"    → one word
 *
 * It ignores punctuation only after the words have been selected.
 *
 * This function never writes to Column D and never changes displayed text.
 */
function firstMeaningfulWordsKey_(value, wordCount) {
  const limit = wordCount || 10;

  const sourceWords = String(value || '')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(function(word) {
      return word !== '';
    })
    .slice(0, limit);

  return sourceWords
    .map(function(word) {
      return word
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/[\u2018\u2019]/g, "'")
        .replace(/[\u201C\u201D]/g, '"')
        .toLowerCase()

        /*
         * Remove punctuation characters from the selected word for
         * comparison, but do not split it into separate words.
         *
         * "don't" and "dont" compare as "dont".
         * "I'm" and "Im" compare as "im".
         * "shut," and "shut" compare as "shut".
         */
        .replace(/[^a-z0-9]+/g, '');
    })
    .filter(function(word) {
      return word !== '';
    })
    .join(' ');
}

function normalizeForMatch_(value) {
  return cleanText_(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/…/g, '...')
    .replace(/[\u2010-\u2015]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function hasRequiredIdentityFields_(row) {
  return [0, 2, 3].every(function(index) {
    return normalizeForMatch_(row[index]) !== '';
  });
}

function rowHasContent_(row) {
  return row.some(function(value) {
    return cleanText_(value) !== '';
  });
}

function ensureDatabaseHeaders_(sheet) {
  const headerRange = sheet.getRange(1, 1, 1, CONFIG.DATABASE_COLUMN_COUNT);
  const headers = headerRange.getDisplayValues()[0].map(cleanText_);
  const headersAreBlank = headers.every(function(header) {
    return header === '';
  });

  if (headersAreBlank) {
    headerRange.setValues([CONFIG.VISIBLE_COLUMNS]);
    return;
  }

  const headersMatch = CONFIG.VISIBLE_COLUMNS.every(function(expectedHeader, index) {
    return headers[index] === expectedHeader;
  });

  if (!headersMatch) {
    throw new Error(
      'Tab "' + sheet.getName() +
      '" does not have the expected A1:H1 database headers. No destination data was changed.'
    );
  }
}

function ensureManualEntriesHeaders_(sheet) {
  const expectedHeaders = CONFIG.VISIBLE_COLUMNS.concat(['Category', 'Import Status']);
  const headerRange = sheet.getRange(1, 1, 1, expectedHeaders.length);
  const headers = headerRange.getDisplayValues()[0].map(cleanText_);
  const headersAreBlank = headers.every(function(header) {
    return header === '';
  });

  if (headersAreBlank) {
    headerRange.setValues([expectedHeaders]);
    return;
  }

  const headersMatch = expectedHeaders.every(function(expectedHeader, index) {
    return headers[index] === expectedHeader;
  });

  if (!headersMatch) {
    throw new Error(
      'The "' + CONFIG.MANUAL_ENTRIES_SHEET_NAME +
      '" tab must have these exact A1:J1 headers:\n\n' +
      expectedHeaders.join(' | ') +
      '\n\nNo manual rows were imported.'
    );
  }
}

function ensureSheetHasRow_(sheet, requiredRow) {
  const currentMaxRows = sheet.getMaxRows();
  if (requiredRow <= currentMaxRows) {
    return;
  }
  sheet.insertRowsAfter(currentMaxRows, requiredRow - currentMaxRows);
}

function appendFolderEntryAtRow_(targetSheet, targetRow, values, documentUrl) {
  ensureSheetHasRow_(targetSheet, targetRow);

  targetSheet
    .getRange(targetRow, 1, 1, CONFIG.DATABASE_COLUMN_COUNT)
    .setValues([values]);

  const label = cleanText_(values[1]);
  const url = cleanText_(documentUrl);
  if (label !== '' && url !== '') {
    const richTextLink = SpreadsheetApp.newRichTextValue()
      .setText(label)
      .setLinkUrl(url)
      .build();
    targetSheet.getRange(targetRow, 2).setRichTextValue(richTextLink);
  }
}

function appendManualSourceRowAtRow_(sourceSheet, sourceRow, targetSheet, targetRow) {
  ensureSheetHasRow_(targetSheet, targetRow);

  sourceSheet
    .getRange(sourceRow, 1, 1, CONFIG.DATABASE_COLUMN_COUNT)
    .copyTo(
      targetSheet.getRange(targetRow, 1, 1, CONFIG.DATABASE_COLUMN_COUNT),
      SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
      false
    );
}

function mergeMissingSupplementalData_(sourceSheet, sourceRow, targetSheet, targetRow) {
  const supplementalStartColumn = 5;
  const supplementalColumnCount = 4;

  ensureSheetHasRow_(targetSheet, targetRow);

  const sourceDisplayValues = sourceSheet
    .getRange(sourceRow, supplementalStartColumn, 1, supplementalColumnCount)
    .getDisplayValues()[0];
  const targetDisplayValues = targetSheet
    .getRange(targetRow, supplementalStartColumn, 1, supplementalColumnCount)
    .getDisplayValues()[0];

  let updatedCount = 0;

  for (let offset = 0; offset < supplementalColumnCount; offset++) {
    const sourceValue = cleanText_(sourceDisplayValues[offset]);
    const targetValue = cleanText_(targetDisplayValues[offset]);

    if (sourceValue === '' || targetValue !== '') {
      continue;
    }

    targetSheet
      .getRange(targetRow, supplementalStartColumn + offset)
      .setValue(sourceValue);
    updatedCount++;
  }

  return { updatedCount: updatedCount };
}

function setManualStatus_(sheet, row, status) {
  sheet.getRange(row, CONFIG.MANUAL_STATUS_COLUMN).setValue(status);
}

function validateFolderId_() {
  if (cleanText_(FOLDER_ID) === '') {
    throw new Error('Set FOLDER_ID near the top of the script before scanning.');
  }
}

function collectFilesRecursively_(folder, relativePath, filesToProcess) {
  const stack = [{ folder: folder, relativePath: relativePath }];

  while (stack.length > 0) {
    const current = stack.pop();
    const currentPath = current.relativePath === '' ? current.folder.getName() : current.relativePath;
    const files = current.folder.getFiles();

    while (files.hasNext()) {
      const file = files.next();
      const mimeType = file.getMimeType();
      if (mimeType !== MimeType.GOOGLE_DOCS && mimeType !== MimeType.SHORTCUT) {
        continue;
      }
      filesToProcess.push({ file: file, sourceFolder: currentPath });
    }

    const subfolders = current.folder.getFolders();
    const nestedFolders = [];
    while (subfolders.hasNext()) {
      const subfolder = subfolders.next();
      nestedFolders.push({
        folder: subfolder,
        relativePath: currentPath + ' / ' + subfolder.getName()
      });
    }

    for (let index = nestedFolders.length - 1; index >= 0; index--) {
      stack.push(nestedFolders[index]);
    }
  }
}

function resolveGoogleDocFile_(file) {
  if (!file || typeof file.getMimeType !== 'function') {
    return null;
  }

  const mimeType = file.getMimeType();

  if (mimeType === MimeType.GOOGLE_DOCS) {
    return file;
  }

  if (mimeType === MimeType.SHORTCUT) {
    try {
      const targetId = file.getTargetId();
      if (!targetId) {
        return null;
      }
      const targetFile = DriveApp.getFileById(targetId);
      return targetFile.getMimeType() === MimeType.GOOGLE_DOCS ? targetFile : null;
    } catch (error) {
      return null;
    }
  }

  return null;
}

/**
 * Detects same-line, ALL-CAPS speaker cues that do not use a colon.
 *
 * Supported:
 *   SIDNEY Talk... or lecture?
 *   FRANZ. Okay. Where do I even start?
 *   ACTOR 4 I want to see the live Africa.
 *
 * Returns:
 *   {
 *     character: 'FRANZ',
 *     dialogue: 'Okay. Where do I even start?'
 *   }
 *
 * Returns null if the line is not a supported cue format.
 */
function parseInlineUppercaseSpeakerWithoutColon_(text) {
  const line = String(text || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  if (line === '') {
    return null;
  }

  /*
   * Period format:
   *
   * FRANZ. Okay. Where do I even start?
   * ACTOR 4. I want to see the live Africa.
   */
  const periodMatch = line.match(
    /^([A-Z][A-Z0-9'’()-]*(?:\s+[A-Z][A-Z0-9'’()-]*){0,3})\.\s+(.+)$/
  );

  if (periodMatch) {
    const character = periodMatch[1].trim();
    const dialogue = periodMatch[2].trim();

    /*
     * Reject all-caps headings such as:
     * ACT ONE. SCENE TWO
     */
    if (
      dialogue !== '' &&
      dialogue !== dialogue.toUpperCase()
    ) {
      return {
        character: character,
        dialogue: dialogue
      };
    }
  }

  /*
   * Space-only format:
   *
   * SIDNEY Talk... or lecture?
   * ACTOR 4 I want to see the live Africa.
   *
   * The remaining dialogue must begin with a normal sentence-like
   * character, quote, or opening parenthesis.
   */
  const spaceMatch = line.match(
    /^([A-Z][A-Z0-9'’()-]*(?:\s+[A-Z][A-Z0-9'’()-]*){0,3})\s+([A-Z“"'(].+)$/
  );

  if (spaceMatch) {
    const character = spaceMatch[1].trim();
    const dialogue = spaceMatch[2].trim();

    /*
     * Reject all-caps headings such as:
     * ACT ONE SCENE TWO
     */
    if (
      dialogue !== '' &&
      dialogue !== dialogue.toUpperCase()
    ) {
      return {
        character: character,
        dialogue: dialogue
      };
    }
  }

  return null;
}


/**
/**
 * Reads the play, playwright, character, and opening monologue excerpt
 * from one Google Doc source file.
 */
function readMonologueHeader_(file) {
  if (!file || typeof file.getId !== 'function') {
    throw new Error('No Google Doc file was supplied for parsing');
  }

  const doc = DocumentApp.openById(file.getId());
  const paragraphs = doc.getBody().getParagraphs();
  const visualLines = [];

  /*
   * Preserve both visual-line header parsing and original paragraph indexes.
   * The index mapping lets the collector inspect real text styling later.
   */
  paragraphs.forEach(function(paragraph, paragraphIndex) {
    const rawText = String(paragraph.getText() || '');

    rawText.split(/\r?\n/).forEach(function(rawLine) {
      const line = normalizeText_(rawLine);

      /*
       * Keep every nonblank visual line, including one-word character labels:
       *
       * Riley
       * Man
       * Rob
       * VICTOR
       * ACTOR 4
       */
      if (line !== '') {
        visualLines.push({
          text: line,
          paragraphIndex: paragraphIndex
        });
      }
    });
  });

  if (visualLines.length < 3) {
    throw new Error(
      'Fewer than three nonblank text lines after splitting paragraph breaks'
    );
  }

  const header = parseHeader_(visualLines.map(function(item) {
    return item.text;
  }));

  if (header.play === '') {
    throw new Error('Could not identify a play title');
  }

  if (header.playwright === '') {
    throw new Error('Could not identify a playwright');
  }

  validatePlaywright_(header.playwright);

  /*
   * Skip an exact repeated title after the header:
   *
   * Dirty Crusty
   * Clare Barron
   * Dirty Crusty  ← ignored
   * VICTOR        ← selected
   */
  const playKey = normalizeForMatch_(header.play);
  let startLine = null;

  for (let i = header.nextLineIndex; i < visualLines.length; i++) {
    if (normalizeForMatch_(visualLines[i].text) === playKey) {
      continue;
    }

    startLine = visualLines[i];
    break;
  }

  if (!startLine) {
    throw new Error(
      'No character or monologue text found after the header'
    );
  }

  const startText = startLine.text;
  const startIndex = startLine.paragraphIndex;

  /*
   * FORMAT 1: ordinary inline cue with a colon.
   *
   * Robbie: See, see? This is why...
   * ACTOR 4: I want to see the live Africa...
   */
  const inlineCue = extractInitialSpeakerCue_(startText);

  if (inlineCue.found && inlineCue.remainingText !== '') {
    if (!characterLabelHasAtMostFiveWords_(inlineCue.character)) {
      throw new Error(
        'Inline character cue exceeds five words: "' +
        inlineCue.character +
        '"'
      );
    }

    const dialogue = collectSpokenTextFromParagraphs_(
      paragraphs,
      startIndex + 1,
      inlineCue.remainingText,
      inlineCue.character
    );

    if (dialogue === '') {
      throw new Error(
        'No spoken monologue text found after inline cue "' +
        inlineCue.character +
        '".'
      );
    }

    return {
      play: header.play,
      playwright: header.playwright,
      character: inlineCue.character,
      firstLine: firstWordsWithEllipsis_(dialogue)
    };
  }

  /*
   * FORMAT 2: same-line ALL-CAPS cue without a colon.
   *
   * SIDNEY Talk... or lecture?
   * FRANZ. Okay. Where do I even start?
   */
  const inlineNoColonCue =
    parseInlineUppercaseSpeakerWithoutColon_(startText);

  if (
    inlineNoColonCue &&
    inlineNoColonCue.dialogue !== ''
  ) {
    if (
      !characterLabelHasAtMostFiveWords_(
        inlineNoColonCue.character
      )
    ) {
      throw new Error(
        'Inline no-colon character cue exceeds five words: "' +
        inlineNoColonCue.character +
        '"'
      );
    }

    const dialogue = collectSpokenTextFromParagraphs_(
      paragraphs,
      startIndex + 1,
      inlineNoColonCue.dialogue,
      inlineNoColonCue.character
    );

    if (dialogue === '') {
      throw new Error(
        'No spoken monologue text found after inline no-colon cue "' +
        inlineNoColonCue.character +
        '".'
      );
    }



    return {
      play: header.play,
      playwright: header.playwright,
      character: inlineNoColonCue.character,
      firstLine: firstWordsWithEllipsis_(dialogue)
    };
  }

  /*
   * FORMAT 3: character name on its own line.
   *
   * Man
   *
   * MAN: I wanted her car...
   *
   * Or:
   *
   * ACTOR 4
   *
   * I want to see the live Africa...
   */
  const character = cleanCharacterLabel_(startText);

  if (character === '') {
    throw new Error('Character line is blank after processing');
  }

  /*
   * Critical guard: a dialogue paragraph must never become a character.
   *
   * Example rejected:
   * Man, shut up. You never belonged with us...
   */
  if (!characterLabelHasAtMostFiveWords_(character)) {
    throw new Error(
      'Character label exceeds five words; likely dialogue instead: "' +
      character +
      '"'
    );
  }

  const dialogue = collectSpokenTextFromParagraphs_(
    paragraphs,
    startIndex + 1,
    '',
    character
  );

  if (dialogue === '') {
    throw new Error(
      'No spoken monologue text found after the character line "' +
      character +
      '".'
    );
  }

  return {
    play: header.play,
    playwright: header.playwright,
    character: character,
    firstLine: firstWordsWithEllipsis_(dialogue)
  };
}


/**
 * Finds the first non-italic spoken paragraph for an expected character.
 *
 * It stops when it reaches a cue for a different character.
 */
function findFirstNonitalicSpokenTextWithIndex_(
  paragraphs,
  startIndex,
  expectedCharacter
) {
  const expectedCharacterKey = normalizeForMatch_(
    expectedCharacter
  );

  for (let i = startIndex; i < paragraphs.length; i++) {
    const paragraphText = normalizeText_(
      paragraphs[i].getText()
    );

    if (paragraphText === '') {
      continue;
    }

    /*
     * Same-line ALL-CAPS cue without a colon:
     *
     * SIDNEY Talk...
     * FRANZ. Okay...
     *
     * If it repeats the expected character, return only the dialogue.
     * If it introduces a different character, stop collecting.
     */
    const inlineNoColonCue =
      parseInlineUppercaseSpeakerWithoutColon_(paragraphText);

    if (inlineNoColonCue) {
      const cueCharacterKey = normalizeForMatch_(
        inlineNoColonCue.character
      );

      if (cueCharacterKey === expectedCharacterKey) {
        return {
          text: inlineNoColonCue.dialogue,
          index: i
        };
      }

      return {
        text: '',
        index: -1
      };
    }

    /*
     * Ordinary colon cue:
     *
     * ROB: I want you to know...
     */
    const colonCue = extractSpeakerCue_(paragraphText);

    if (colonCue.found) {
      const cueCharacterKey = normalizeForMatch_(
        colonCue.character
      );

      /*
       * Repeated formatting for the same character:
       *
       * Rob
       * ROB: I want you to know...
       */
      if (
        cueCharacterKey === expectedCharacterKey &&
        colonCue.remainingText !== ''
      ) {
        const spokenText = getSpokenTextAfterSpeakerCue_(
          paragraphs[i],
          colonCue.endIndex
        );

        if (spokenText !== '') {
          return {
            text: spokenText,
            index: i
          };
        }
      }

      /*
       * A different cue, such as DAD:, starts another speaker.
       */
      return {
        text: '',
        index: -1
      };
    }

    /*
     * The next regular paragraph is dialogue.
     */
    const spokenText = getSpokenText_(paragraphs[i]);

    if (spokenText !== '') {
      return {
        text: spokenText,
        index: i
      };
    }
  }

  return {
    text: '',
    index: -1
  };
}

function collectSpokenTextFromParagraphs_(
  paragraphs,
  startIndex,
  initialText,
  expectedCharacter
) {
  let collected = cleanSpokenText_(initialText || '');
  let collectedWordCount = countWords_(collected);
  const expectedCharacterKey = normalizeForMatch_(expectedCharacter || '');

  for (let i = startIndex; i < paragraphs.length; i++) {
    const rawText = normalizeText_(paragraphs[i].getText());

    /*
     * Empty paragraphs only create visual spacing. Do not end dialogue.
     */
    if (rawText === '') {
      continue;
    }

    const nextCue = extractSpeakerCue_(rawText);

    if (nextCue.found) {
      const nextCharacterKey = normalizeForMatch_(nextCue.character);

      /*
       * Same-character cue:
       *
       * Man
       * MAN: I wanted her car...
       *
       * Keep the text after “MAN:”.
       */
      if (
        expectedCharacterKey !== '' &&
        nextCharacterKey === expectedCharacterKey
      ) {
        const cueDialogue = cleanSpokenText_(nextCue.remainingText);

        if (cueDialogue !== '') {
          collectedWordCount += countWords_(cueDialogue);
          collected = normalizeText_(collected + ' ' + cueDialogue);
        }

        if (collectedWordCount >= CONFIG.FIRST_LINE_WORD_COUNT) {
          break;
        }

        continue;
      }

      /*
       * Different character means the selected monologue has ended.
       */
      break;
    }

    /*
     * Strip italic runs. If this paragraph is only italic stage direction,
     * it becomes empty and scanning continues after it.
     */
    const spokenPiece = getNonItalicText_(paragraphs[i]);

    if (spokenPiece === '') {
      continue;
    }

    collectedWordCount += countWords_(spokenPiece);
    collected = normalizeText_(collected + ' ' + spokenPiece);

    if (collectedWordCount >= CONFIG.FIRST_LINE_WORD_COUNT) {
      break;
    }
  }

  return collected;
}

function getAnySpeakerCue_(text) {
  const normalized = normalizeText_(text);
  const colonCue = extractSpeakerCue_(normalized);

  if (colonCue.found) {
    return colonCue;
  }

  const firstWord = (normalized.split(/\s+/)[0] || '')
    .replace(/^[^\p{L}\p{M}]+|[^\p{L}\p{M}]+$/gu, '');

  if (!/^[\p{Lu}\p{M}]{2,}$/u.test(firstWord)) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  return extractUppercaseCueWithoutColon_(normalized);
}

function extractSpeakerCue_(text) {
  const normalized = normalizeText_(text);

  /*
   * Only examine a colon that occurs near the beginning of the paragraph.
   * A speaker label should be short; a normal spoken sentence will almost
   * always put the colon later.
   */
  const colonIndex = normalized.indexOf(':');

  if (colonIndex <= 0 || colonIndex > 40) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  const candidateLabel = normalizeText_(
    normalized.substring(0, colonIndex)
  );

  const remainingText = normalizeText_(
    normalized.substring(colonIndex + 1)
  );

  /*
   * A cue must have actual dialogue after the colon.
   */
  if (remainingText === '') {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  /*
   * Accept only a name / role-shaped label:
   *
   * MAN
   * Robbie
   * ACTOR 4
   * MRS. THOMAS
   * Young Man
   *
   * Rejects:
   * These are the facts
   * I said
   * My brother says
   * The order
   */
  if (!isLikelySpeakerLabel_(candidateLabel)) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  return {
    found: true,
    character: cleanCharacterLabel_(candidateLabel),
    remainingText: remainingText,
    endIndex: colonIndex + 1
  };
}

function extractInitialSpeakerCue_(text) {
  const normalized = normalizeText_(text);
  const colonIndex = normalized.indexOf(':');

  if (colonIndex <= 0 || colonIndex > 45) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  const candidate = cleanCharacterLabel_(
    normalized.substring(0, colonIndex)
  );
  const remainingText = normalizeText_(
    normalized.substring(colonIndex + 1)
  );

  /*
   * Called only for the initial character/dialogue line. It permits
   * title-case character labels such as Robbie and Young Man, as well as
   * ACTOR 4. It still rejects sentence-like material before the colon.
   */
  if (
    candidate === '' ||
    remainingText === '' ||
    candidate.length > 45 ||
    candidate.split(/\s+/).length > 5 ||
    !/^[\p{L}\p{M}][\p{L}\p{M}.'’\-]*(?:\s+[\p{L}\p{M}][\p{L}\p{M}.'’\-]*)*(?:\s+#?\d+)?$/u
      .test(candidate)
  ) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  return {
    found: true,
    character: candidate,
    remainingText: remainingText,
    endIndex: colonIndex + 1
  };
}

function isLikelySpeakerLabel_(text) {
  const label = cleanCharacterLabel_(text);

  if (label === '') {
    return false;
  }

  /*
   * Never accept sentence-like speaker labels.
   */
  if (
    label.length > 35 ||
    label.split(/\s+/).length > 4 ||
    /[!?;,]/.test(label)
  ) {
    return false;
  }

  /*
   * A label must consist of name/role tokens, with an optional final number:
   *
   * MAN
   * ACTOR 4
   * Young Man
   * MRS. THOMAS
   *
   * This deliberately rejects lowercase action/reporting phrases such as:
   * "I said", "She said", "These are the facts".
   */
  if (
    !/^[\p{Lu}][\p{Lu}\p{M}.'’\-]*(?:\s+[\p{Lu}][\p{Lu}\p{M}.'’\-]*)*(?:\s+#?\d+)?$/u
      .test(label)
  ) {
    return false;
  }

  return true;
}

function extractUppercaseCueWithoutColon_(text) {
  const normalized = normalizeText_(text);

  /*
   * First capture an all-caps cue with an optional final number, then capture
   * remaining dialogue.
   */
  const match = normalized.match(
    /^((?:[\p{Lu}\p{M}][\p{Lu}\p{M}.'’-]*\s+){0,5}[\p{Lu}\p{M}][\p{Lu}\p{M}.'’-]*(?:\s+\d+)?)\s+(.+)$/u
  );

  if (!match) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  const candidateCue = normalizeText_(match[1]);
  const remainingText = normalizeText_(match[2]);

  if (
    candidateCue === '' ||
    remainingText === '' ||
    !isValidCharacterLabel_(candidateCue)
  ) {
    return {
      found: false,
      character: '',
      remainingText: normalized,
      endIndex: 0
    };
  }

  return {
    found: true,
    character: cleanCharacterLabel_(candidateCue),
    remainingText: remainingText,
    endIndex: normalized.indexOf(remainingText)
  };
}

function cleanCharacterLabel_(text) {
  const result = removeLeadingLabel_(text, ['character', 'role', 'speaker']);
  return normalizeText_(result.text).replace(/[.:\s]+$/, '').trim();
}

function isValidCharacterLabel_(text) {
  const normalized = normalizeText_(text);

  /*
   * Allows:
   * Greg
   * MRS. THOMAS
   * Actor 1
   * WOMAN 2
   * MAN #3
   *
   * Does not allow:
   * 1
   * 12
   * A normal multiword dialogue sentence
   */
  return /^[\p{L}\p{M}][\p{L}\p{M}.'’-]*(?:\s+[\p{L}\p{M}][\p{L}\p{M}.'’-]*)*(?:\s+#?\d+)?$/u
    .test(normalized);
}


function getNonItalicText_(paragraph) {
  /*
   * This is the function called by your existing collector. It preserves
   * ordinary dialogue and returns an empty string for italic-only direction.
   */
  return cleanSpokenText_(
    removeItalicTextAfterIndex_(paragraph, 0)
  );
}

function getSpokenText_(paragraph) {
  const nonItalicText = cleanSpokenText_(
    removeItalicTextAfterIndex_(paragraph, 0)
  );

  if (nonItalicText !== '') {
    return nonItalicText;
  }

  return cleanSpokenText_(normalizeText_(paragraph.getText()));
}

function getSpokenTextAfterSpeakerCue_(paragraph, startAtIndex) {
  const nonItalicText = cleanSpokenText_(
    removeItalicTextAfterIndex_(paragraph, startAtIndex)
  );

  if (nonItalicText !== '') {
    return nonItalicText;
  }

  return cleanSpokenText_(
    normalizeText_(paragraph.getText().substring(startAtIndex))
  );
}

function removeItalicTextAfterIndex_(paragraph, startAtIndex) {
  const textElement = paragraph.editAsText();
  const fullText = textElement.getText();
  const attributeIndices = textElement.getTextAttributeIndices();
  let result = '';

  for (let i = 0; i < attributeIndices.length; i++) {
    const runStart = attributeIndices[i];
    const runEnd = i + 1 < attributeIndices.length
      ? attributeIndices[i + 1]
      : fullText.length;

    if (runEnd <= startAtIndex) {
      continue;
    }

    const includedStart = Math.max(runStart, startAtIndex);
    if (textElement.isItalic(runStart) !== true) {
      result += fullText.substring(includedStart, runEnd);
    }
  }

  return normalizeText_(result);
}

function characterLabelHasAtMostFiveWords_(text) {
  const label = normalizeText_(text);

  if (label === '') {
    return false;
  }

  /*
   * Count whitespace-separated label words.
   *
   * Examples:
   * "Riley"               → 1
   * "Young Man"           → 2
   * "ACTOR 4"             → 2
   * "Woman With Red Hat"  → 4
   * "Man, shut up..."     → too many words → false
   */
  const wordCount = label.split(/\s+/).filter(function(word) {
    return word !== '';
  }).length;

  return wordCount >= 1 && wordCount <= 5;
}

function cleanSpokenText_(text) {
  let cleaned = normalizeText_(text);

  if (CONFIG.EXCLUDE_PARENTHETICAL_STAGE_DIRECTIONS) {
    let previous;
    do {
      previous = cleaned;
      cleaned = cleaned.replace(/(?:\([^()]*\)|\[[^\[\]]*\])/g, ' ');
    } while (cleaned !== previous);
  }

  return normalizeText_(cleaned);
}

/**
 * Returns the first configured number of whitespace-separated words.
 *
 * Important:
 * - Preserves source punctuation exactly.
 * - Preserves apostrophes, quotes, dashes, ellipses, commas, periods, etc.
 * - Adds "..." only when the source contains more words than the excerpt.
 * - Does not remove or normalize punctuation.
 */
function firstWordsWithEllipsis_(text) {
  const words = String(text || '')
    .replace(/\s+/g, ' ')
    .trim()
    .split(' ')
    .filter(function(word) {
      return word !== '';
    });

  if (words.length === 0) {
    return '';
  }

  const wordLimit = CONFIG.FIRST_LINE_WORD_COUNT;
  const excerpt = words.slice(0, wordLimit).join(' ');

  /*
   * If the source ended at or before the word limit, return it exactly.
   * Do not add artificial ellipses.
   */
  if (words.length <= wordLimit) {
    return excerpt;
  }

  /*
   * If the final included source word already ends in an ellipsis,
   * do not add another one.
   */
  if (/(\.\.\.|…)$/.test(excerpt)) {
    return excerpt;
  }

  return excerpt + '...';
}

function countWords_(text) {
  return normalizeText_(text).split(/\s+/).filter(Boolean).length;
}

function parseHeader_(lines) {
  const firstLine = normalizeText_(lines[0] || '');
  if (firstLine === '') {
    return { play: '', playwright: '', nextLineIndex: 0 };
  }

  const secondLine = normalizeText_(lines[1] || '');
  const knownException = getKnownPlayHeaderException_(firstLine);

  if (knownException.isMatch) {
    return {
      play: knownException.play,
      playwright: cleanPlaywrightName_(secondLine),
      nextLineIndex: 2
    };
  }

  const secondLineByMatch = secondLine.match(/^\s*by\s+(.+)$/i);
  if (secondLineByMatch) {
    return {
      play: firstLine,
      playwright: cleanPlaywrightName_(secondLineByMatch[1]),
      nextLineIndex: 2
    };
  }

  const bySplit = splitAtBy_(firstLine);
  if (bySplit.hasSeparator) {
    return {
      play: bySplit.before,
      playwright: cleanPlaywrightName_(bySplit.after),
      nextLineIndex: 1
    };
  }

  const punctuationSplit = splitTitleAndAuthorAtSeparator_(firstLine);
  if (punctuationSplit.hasSeparator) {
    return {
      play: punctuationSplit.before,
      playwright: cleanPlaywrightName_(punctuationSplit.after),
      nextLineIndex: 1
    };
  }

  const labeledTitle = removeLeadingLabel_(firstLine, ['play', 'title', 'play title']);
  const labeledPlaywright = removeLeadingLabel_(secondLine, [
    'playwright', 'author', 'written by'
  ]);

  if (labeledTitle.wasLabeled && labeledPlaywright.wasLabeled) {
    return {
      play: labeledTitle.text,
      playwright: cleanPlaywrightName_(labeledPlaywright.text),
      nextLineIndex: 2
    };
  }

  return {
    play: labeledTitle.text,
    playwright: cleanPlaywrightName_(secondLine),
    nextLineIndex: 2
  };
}

function removeLeadingLabel_(text, labels) {
  const normalized = normalizeText_(text);
  if (normalized === '') {
    return { wasLabeled: false, text: '' };
  }

  const lower = normalized.toLowerCase();

  for (let i = 0; i < labels.length; i++) {
    const label = labels[i].toLowerCase();
    const labelLength = label.length;
    const prefixPatterns = [
      label + ':',
      label + ' :',
      label + '-',
      label + ' -'
    ];

    for (let j = 0; j < prefixPatterns.length; j++) {
      const prefix = prefixPatterns[j];
      if (lower.indexOf(prefix) === 0) {
        return {
          wasLabeled: true,
          text: normalized.substring(prefix.length).replace(/^\s+/, '').trim()
        };
      }
    }

    if (lower.indexOf(label + ' ') === 0) {
      const remainder = normalized.substring(labelLength).replace(/^\s+/, '');
      if (/^[\-:]/.test(remainder)) {
        return {
          wasLabeled: true,
          text: remainder.replace(/^[\-:]/, '').trim()
        };
      }
    }
  }

  return { wasLabeled: false, text: normalized };
}

function cleanPlaywrightName_(text) {
  let cleaned = normalizeText_(text)
    .trim();

  /*
   * Repair values already incorrectly inverted by an earlier script:
   *
   * "Ferrentino, By Lindsey" → "Ferrentino, Lindsey"
   * "Silverman, By Jen"      → "Silverman, Jen"
   */
  cleaned = cleaned.replace(
    /^(.+?),\s*by\s+(.+)$/i,
    '$1, $2'
  );

  /*
   * Normal source-credit cleanup:
   *
   * "By Lindsey Ferrentino" → "Lindsey Ferrentino"
   * "Written by Jen Silverman" → "Jen Silverman"
   */
  cleaned = cleaned
    .replace(
      /^(?:written\s+by|playwright|author|by)\s*[:\-–—]?\s*/i,
      ''
    )
    .trim();

  return formatPlaywrightCredit_(cleaned);
}


/**
 * Formats one or more playwright credits.
 *
 * The critical rule:
 * - If the value is already formatted as "Last, First", preserve each
 *   name exactly and only standardize spacing and "and".
 * - Do not split a formatted name at its internal comma.
 */
function formatPlaywrightCredit_(credit) {
  const normalized = normalizeText_(credit)
    .replace(/\s*&\s*/g, ' and ')
    .replace(/\s+and\s+/gi, ' and ')
    .replace(/\s*,\s*/g, ', ')
    .trim();

  if (normalized === '') {
    return '';
  }

  /*
   * Already formatted multi-name credits:
   *
   * Lhota, Barbara and Brodsky, Ira
   * Nottage, Lynn and Jacobs-Jenkins, Branden
   *
   * Preserve them. The old function mistakenly split every comma,
   * which converted "Lhota, Barbara and Brodsky, Ira" into four names.
   */
  if (
    normalized.indexOf(' and ') !== -1 &&
    normalized.match(/,/g) &&
    normalized.match(/,/g).length >= 2
  ) {
    return normalized;
  }

  /*
   * One already formatted name:
   *
   * Silverman, Jen
   * Jacobs-Jenkins, Branden
   */
  if (normalized.indexOf(' and ') === -1 && normalized.indexOf(',') !== -1) {
    return normalized;
  }

  /*
   * Multiple unformatted names:
   *
   * Jen Silverman and Lynn Nottage
   * Jen Silverman & Lynn Nottage
   */
  const names = normalized.split(/\s+and\s+/i)
    .map(function(name) {
      return formatOnePlaywrightName_(name);
    })
    .filter(function(name) {
      return name !== '';
    });

  if (names.length === 0) {
    return '';
  }

  if (names.length === 1) {
    return names[0];
  }

  if (names.length === 2) {
    return names[0] + ' and ' + names[1];
  }

  return names.slice(0, -1).join(', ') + ', and ' + names[names.length - 1];
}


/**
 * Retained for compatibility if another part of the script calls it.
 *
 * It now splits only on "and", never on commas. Commas belong to
 * "Last, First" formatting and must not be treated as separators.
 */
function splitPlaywrightCredit_(credit) {
  return normalizeText_(credit)
    .replace(/\s*&\s*/g, ' and ')
    .split(/\s+and\s+/i)
    .map(cleanText_)
    .filter(function(name) {
      return name !== '';
    });
}


/**
 * Formats one person only.
 *
 * "Jen Silverman"       → "Silverman, Jen"
 * "Silverman, Jen"      → "Silverman, Jen"
 * "Jacobs-Jenkins, Branden" remains unchanged.
 */
function formatOnePlaywrightName_(name) {
  const cleaned = normalizeText_(name)
    .replace(/^(?:written\s+by|playwright|author|by)\s*[:\-–—]?\s*/i, '')
    .trim();

  if (cleaned === '') {
    return '';
  }

  /*
   * It is already in Last, First format.
   */
  if (cleaned.indexOf(',') !== -1) {
    return cleaned.replace(/\s*,\s*/g, ', ');
  }

  const words = cleaned.split(/\s+/).filter(Boolean);

  if (words.length < 2) {
    return cleaned;
  }

  const lastName = words.pop();

  return lastName + ', ' + words.join(' ');
}

function validatePlaywright_(playwright) {
  const cleaned = cleanText_(playwright);
  if (cleaned === '') {
    throw new Error('Playwright field is blank');
  }
  if (cleaned.length > 250) {
    throw new Error('Playwright field is unusually long: ' + cleaned);
  }
}

function getKnownPlayHeaderException_(text) {
  const normalized = normalizeForMatch_(text);
  if (normalized.indexOf('antigone, presented by the girls of st. catherine') === 0) {
    return {
      isMatch: true,
      play: "Antigone, Presented by the Girls of St. Catherine's"
    };
  }
  return { isMatch: false, play: '' };
}

function splitAtBy_(text) {
  const normalized = normalizeText_(text);
  const match = normalized.match(/\s+\bby\b\s+/i);
  if (!match) {
    return { hasSeparator: false, before: '', after: '' };
  }
  return {
    hasSeparator: true,
    before: normalized.substring(0, match.index).trim(),
    after: normalized.substring(match.index + match[0].length).trim()
  };
}

function splitTitleAndAuthorAtSeparator_(text) {
  const normalized = normalizeText_(text);
  const match = normalized.match(/\s+(?:-|\/)\s+/);
  if (!match) {
    return { hasSeparator: false, before: '', after: '' };
  }

  const before = normalized.substring(0, match.index).trim();
  const after = normalized.substring(match.index + match[0].length).trim();
  if (before === '' || after === '') {
    return { hasSeparator: false, before: '', after: '' };
  }
  return { hasSeparator: true, before: before, after: after };
}

function updateSkippedFilesSheet_(ss, skippedFiles) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SKIPPED_SHEET_NAME);
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.clear();

  sheet.getRange('A1:E1').setValues([[
    'Root folder',
    'Source folder',
    'File',
    'File URL',
    'Reason skipped'
  ]]);

  if (skippedFiles.length > 0) {
    const sortedFiles = skippedFiles.slice().sort(function(a, b) {
      return a.sourceFolder.localeCompare(b.sourceFolder) ||
        a.fileName.localeCompare(b.fileName);
    });
    const rows = sortedFiles.map(function(item) {
      return [item.rootFolder, item.sourceFolder, item.fileName, item.fileUrl, item.reason];
    });

    sheet.getRange(2, 1, rows.length, 5).setValues(rows);

    const fileLinks = sortedFiles.map(function(item) {
      const builder = SpreadsheetApp.newRichTextValue().setText(cleanText_(item.fileName));
      if (cleanText_(item.fileUrl) !== '') {
        builder.setLinkUrl(cleanText_(item.fileUrl));
      }
      return [builder.build()];
    });
    sheet.getRange(2, 3, fileLinks.length, 1).setRichTextValues(fileLinks);
    sheet.getRange(1, 1, rows.length + 1, 5).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.getRange('A1:E1').setFontWeight('bold').setBackground('#f4cccc');
  sheet.getRange('B:E').setWrap(true);
  sheet.setColumnWidth(1, 220);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 260);
  sheet.setColumnWidth(4, 360);
  sheet.setColumnWidth(5, 420);
}

function previewPlaywrightCleanup() {
  runPlaywrightCleanup_(true);
}

function applyPlaywrightCleanup() {
  runPlaywrightCleanup_(false);
}

function runPlaywrightCleanup_(dryRun) {
  const ss = getDestinationSpreadsheet_();
  const databaseSheets = getMonologueDatabaseSheets_(ss);
  const changes = collectColumnCleanupChanges_(databaseSheets, 1, cleanPlaywrightName_);
  showCleanupPreviewOrApply_(ss, changes, dryRun, 'Playwright', 1);
}

function previewCharacterCleanup() {
  runCharacterCleanup_(true);
}

function applyCharacterCleanup() {
  runCharacterCleanup_(false);
}

function runCharacterCleanup_(dryRun) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const ss = getDestinationSpreadsheet_();
    const databaseSheets = getMonologueDatabaseSheets_(ss);
    const changes = collectColumnCleanupChanges_(databaseSheets, 3, cleanExistingCharacterName_);
    showCleanupPreviewOrApply_(ss, changes, dryRun, 'Character', 3);
  } finally {
    lock.releaseLock();
  }
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
    const databaseSheets = getMonologueDatabaseSheets_(ss);
    const changes = collectColumnCleanupChanges_(
      databaseSheets,
      4,
      removeQuotationMarksFromFirstLine_
    );
    showCleanupPreviewOrApply_(ss, changes, dryRun, 'First Line quote', 4);
  } finally {
    lock.releaseLock();
  }
}

function collectColumnCleanupChanges_(databaseSheets, column, cleaner) {
  const changes = [];

  databaseSheets.forEach(function(sheet) {
    const values = getSheetColumnDisplayValues_(sheet, column);
    if (values.length === 0) {
      return;
    }

    values.forEach(function(value, index) {
      const original = cleanText_(value);
      const cleaned = cleaner(original);
      if (original !== '' && cleaned !== '' && cleaned !== original) {
        changes.push({
          sheetName: sheet.getName(),
          row: CONFIG.DATA_START_ROW + index,
          original: original,
          cleaned: cleaned
        });
      }
    });
  });

  return changes;
}

function showCleanupPreviewOrApply_(ss, changes, dryRun, label, column) {
  if (changes.length === 0) {
    showMessage_(
      dryRun
        ? label + ' cleanup preview complete: no cells need changes.'
        : label + ' cleanup complete: no cells needed changes.'
    );
    return;
  }

  const previewLimit = 30;
  const columnLetter = String.fromCharCode(64 + column);
  const previewLines = changes.slice(0, previewLimit).map(function(change) {
    return change.sheetName + '!' + columnLetter + change.row +
      '\n  "' + change.original + '" → "' + change.cleaned + '"';
  });
  const extraCount = changes.length - previewLines.length;

  Logger.log(JSON.stringify(changes, null, 2));

  if (dryRun) {
    showMessage_(
      label.toUpperCase() + ' CLEANUP PREVIEW\n\n' +
      'Cells that would change: ' + changes.length + '\n\n' +
      previewLines.join('\n\n') +
      (extraCount > 0 ? '\n\n…and ' + extraCount + ' more.' : '') +
      '\n\nNo cells were changed.'
    );
    return;
  }

  const updatesBySheet = {};
  changes.forEach(function(change) {
    if (!updatesBySheet[change.sheetName]) {
      updatesBySheet[change.sheetName] = [];
    }
    updatesBySheet[change.sheetName].push({
      row: change.row,
      value: change.cleaned
    });
  });

  Object.keys(updatesBySheet).forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    updatesBySheet[sheetName].forEach(function(update) {
      sheet.getRange(update.row, column).setValue(update.value);
    });
  });

  SpreadsheetApp.flush();
  showMessage_(label + ' cleanup complete.\n\nUpdated cells: ' + changes.length);
}

function getMonologueDatabaseSheets_(ss) {
  return ss.getSheets().filter(isMonologueDatabaseSheet_);
}

function isMonologueDatabaseSheet_(sheet) {
  const excludedSheetNames = [
    CONFIG.MANUAL_ENTRIES_SHEET_NAME,
    CONFIG.SKIPPED_SHEET_NAME,
    CONFIG.SOURCE_CHANGES_REVIEW_SHEET_NAME
  ];
  if (excludedSheetNames.indexOf(sheet.getName()) !== -1) {
    return false;
  }

  const headers = sheet
    .getRange(1, 1, 1, CONFIG.DATABASE_COLUMN_COUNT)
    .getDisplayValues()[0]
    .map(cleanText_);

  return CONFIG.VISIBLE_COLUMNS.every(function(expectedHeader, index) {
    return headers[index] === expectedHeader;
  });
}

function cleanExistingCharacterName_(value) {
  return cleanText_(value).replace(/[.:\s]+$/, '').trim();
}

function removeQuotationMarksFromFirstLine_(value) {
  return cleanText_(value)
    .replace(/["“”„‟«»‹›‘’‚‛`´]/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}

function previewDatabaseDuplicateRemoval() {
  runExactDuplicateRemoval_(true);
}

function applyDatabaseDuplicateRemoval() {
  runExactDuplicateRemoval_(false);
}

function runExactDuplicateRemoval_(dryRun) {
  const databaseSheets = getMonologueDatabaseSheets_(getDestinationSpreadsheet_());
  const plan = buildDuplicateRemovalPlan_(databaseSheets, 6, 'Exact: same Playwright + Character + first 6 First Line words');
  executeDuplicateRemovalPlan_(plan, dryRun, 'EXACT DUPLICATE');
}

function previewLikelyDuplicateRemoval() {
  runLikelyDuplicateRemoval_(true);
}

function applyLikelyDuplicateRemoval() {
  runLikelyDuplicateRemoval_(false);
}

function runLikelyDuplicateRemoval_(dryRun) {
  const databaseSheets = getMonologueDatabaseSheets_(getDestinationSpreadsheet_());
  const exactPlan = buildDuplicateRemovalPlan_(
    databaseSheets,
    3,
    'Exact: same Playwright + Character + first 3 First Line words'
  );
  const exactRows = new Set();
  exactPlan.forEach(function(item) {
    item.allRows.forEach(function(rowNumber) {
      exactRows.add(item.sheetName + '|' + rowNumber);
    });
  });

  const likelyPlan = buildDuplicateRemovalPlan_(
    databaseSheets,
    6,
    'Likely: same Playwright + Character + first 6 First Line words',
    exactRows
  );

  executeDuplicateRemovalPlan_(exactPlan.concat(likelyPlan), dryRun, 'LIKELY DUPLICATE');
}

function buildDuplicateRemovalPlan_(databaseSheets, wordCount, matchType, excludedRowKeys) {
  const plan = [];
  const excluded = excludedRowKeys || new Set();

  databaseSheets.forEach(function(sheet) {
    const sheetName = sheet.getName();
    const rows = getDatabaseDataRows_(sheet);
    if (rows.length === 0) {
      return;
    }

    const groups = new Map();
    rows.forEach(function(row, index) {
      const rowNumber = CONFIG.DATA_START_ROW + index;
      if (
        excluded.has(sheetName + '|' + rowNumber) ||
        !rowHasContent_(row) ||
        !hasRequiredIdentityFields_(row)
      ) {
        return;
      }

      const key = [
        normalizeForMatch_(row[0]),
        normalizeForMatch_(row[2]),
        firstMeaningfulWordsKey_(row[3], wordCount)
      ].join(' | ');

      if (!groups.has(key)) {
        groups.set(key, []);
      }
      groups.get(key).push({
        rowNumber: rowNumber,
        displayValues: row,
        filledCount: countFilledDatabaseCells_(row)
      });
    });

    groups.forEach(function(group, fingerprint) {
      if (group.length < 2) {
        return;
      }
      plan.push(makeDuplicatePlanItem_(sheetName, matchType, fingerprint, group));
    });
  });

  return plan;
}

function makeDuplicatePlanItem_(sheetName, matchType, fingerprint, group) {
  const sortedCandidates = group.slice().sort(function(a, b) {
    if (b.filledCount !== a.filledCount) {
      return b.filledCount - a.filledCount;
    }
    return a.rowNumber - b.rowNumber;
  });

  const keeper = sortedCandidates[0];
  const removeCandidates = sortedCandidates.slice(1);
  const simulatedSupplemental = keeper.displayValues.slice(4, 8);
  const mergeSources = [];

  removeCandidates.forEach(function(candidate) {
    const supplemental = candidate.displayValues.slice(4, 8);
    const filledColumns = [];

    for (let offset = 0; offset < 4; offset++) {
      if (
        cleanText_(simulatedSupplemental[offset]) === '' &&
        cleanText_(supplemental[offset]) !== ''
      ) {
        simulatedSupplemental[offset] = supplemental[offset];
        filledColumns.push(String.fromCharCode('E'.charCodeAt(0) + offset));
      }
    }

    if (filledColumns.length > 0) {
      mergeSources.push({ sourceRow: candidate.rowNumber, filledColumns: filledColumns });
    }
  });

  return {
    sheetName: sheetName,
    matchType: matchType,
    fingerprint: fingerprint,
    keeperRow: keeper.rowNumber,
    keeperFilledCount: keeper.filledCount,
    removeRows: removeCandidates.map(function(candidate) { return candidate.rowNumber; }),
    allRows: sortedCandidates.map(function(candidate) { return candidate.rowNumber; }),
    candidates: sortedCandidates.map(function(candidate) {
      return { rowNumber: candidate.rowNumber, filledCount: candidate.filledCount };
    }),
    mergeSources: mergeSources
  };
}

function executeDuplicateRemovalPlan_(plan, dryRun, label) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    if (plan.length === 0) {
      showMessage_(
        dryRun
          ? label + ' preview complete: no duplicate groups found.'
          : label + ' removal complete: no duplicate groups were found.'
      );
      return;
    }

    const rowsToRemove = plan.reduce(function(total, item) {
      return total + item.removeRows.length;
    }, 0);
    const previewLimit = 20;
    const previewLines = plan.slice(0, previewLimit).map(function(item) {
      const scores = item.candidates.map(function(candidate) {
        return 'row ' + candidate.rowNumber + ': ' + candidate.filledCount + '/8 filled';
      }).join('; ');
      const mergeText = item.mergeSources.length > 0
        ? '\n  Fill keeper blanks: ' + item.mergeSources.map(function(source) {
            return 'from row ' + source.sourceRow + ' → ' + source.filledColumns.join(',');
          }).join('; ')
        : '';

      return item.sheetName +
        '\n  Match: ' + item.matchType +
        '\n  Keep row: ' + item.keeperRow + ' (' + item.keeperFilledCount + '/8 filled)' +
        '\n  Remove rows: ' + item.removeRows.join(', ') +
        '\n  Scores: ' + scores + mergeText;
    });

    Logger.log(JSON.stringify(plan, null, 2));

    if (dryRun) {
      showMessage_(
        label + ' REMOVAL PREVIEW\n\n' +
        'Duplicate groups: ' + plan.length + '\n' +
        'Rows that would be removed: ' + rowsToRemove + '\n\n' +
        previewLines.join('\n\n') +
        (plan.length > previewLimit
          ? '\n\n…and ' + (plan.length - previewLimit) + ' more duplicate groups.'
          : '') +
        '\n\nNo rows or cells were changed.\n\n' +
        'The row with the most filled A:H cells is kept; ties keep the topmost row.'
      );
      return;
    }

    const ui = SpreadsheetApp.getUi();
    const response = ui.alert(
      'Apply duplicate removal?',
      'This will remove ' + rowsToRemove + ' row(s) across ' + plan.length +
      ' group(s). The most complete A:H row is kept. Only blank E:H cells in a keeper may be filled before deletion.',
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      showMessage_('Duplicate removal was canceled. No rows were changed.');
      return;
    }

    const ss = getDestinationSpreadsheet_();
    plan.forEach(function(item) {
      const sheet = ss.getSheetByName(item.sheetName);
      item.removeRows.forEach(function(sourceRow) {
        mergeMissingDuplicateRowData_(sheet, sourceRow, item.keeperRow);
      });
    });

    const rowsToDeleteBySheet = new Map();
    plan.forEach(function(item) {
      if (!rowsToDeleteBySheet.has(item.sheetName)) {
        rowsToDeleteBySheet.set(item.sheetName, []);
      }
      rowsToDeleteBySheet.get(item.sheetName).push.apply(
        rowsToDeleteBySheet.get(item.sheetName),
        item.removeRows
      );
    });

    rowsToDeleteBySheet.forEach(function(rowNumbers, sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      rowNumbers.sort(function(a, b) { return b - a; }).forEach(function(rowNumber) {
        sheet.deleteRow(rowNumber);
      });
    });

    SpreadsheetApp.flush();
    showMessage_(
      label + ' REMOVAL COMPLETE\n\n' +
      'Duplicate groups resolved: ' + plan.length + '\n' +
      'Rows removed: ' + rowsToRemove
    );
  } finally {
    lock.releaseLock();
  }
}

function countFilledDatabaseCells_(row) {
  return row.slice(0, CONFIG.DATABASE_COLUMN_COUNT).reduce(function(count, value) {
    return count + (cleanText_(value) === '' ? 0 : 1);
  }, 0);
}

function mergeMissingDuplicateRowData_(sheet, sourceRow, keeperRow) {
  const supplementalStartColumn = 5;
  const supplementalColumnCount = 4;
  ensureSheetHasRow_(sheet, Math.max(sourceRow, keeperRow));

  const sourceValues = sheet
    .getRange(sourceRow, supplementalStartColumn, 1, supplementalColumnCount)
    .getDisplayValues()[0];
  const keeperValues = sheet
    .getRange(keeperRow, supplementalStartColumn, 1, supplementalColumnCount)
    .getDisplayValues()[0];

  for (let offset = 0; offset < supplementalColumnCount; offset++) {
    if (
      cleanText_(sourceValues[offset]) === '' ||
      cleanText_(keeperValues[offset]) !== ''
    ) {
      continue;
    }

    sheet
      .getRange(keeperRow, supplementalStartColumn + offset)
      .setValue(cleanText_(sourceValues[offset]));
  }
}

function buildSourceChangesReview() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    validateFolderId_();
    const ss = getDestinationSpreadsheet_();
    const rootFolder = DriveApp.getFolderById(FOLDER_ID);
    const rootFolderName = rootFolder.getName();
    const databaseSheetName = makeSafeTabName_(rootFolderName);
    const databaseSheet = ss.getSheetByName(databaseSheetName);

    if (!databaseSheet) {
      throw new Error(
        'Could not find the destination database tab "' + databaseSheetName +
        '". Run the folder scanner once first.'
      );
    }

    ensureDatabaseHeaders_(databaseSheet);
    const databaseIndex = getSourceReviewDatabaseIndex_(databaseSheet);
    const filesToProcess = [];
    collectFilesRecursively_(rootFolder, '', filesToProcess);

    const processedDocumentIds = new Set();
    const reviewRows = [];
    const totals = {
      unchanged: 0,
      changed: 0,
      noMatch: 0,
      ambiguous: 0,
      unparseable: 0,
      unsupported: 0
    };

    filesToProcess.forEach(function(item) {
      const fallbackName = item.file.getName();
      const fallbackUrl = item.file.getUrl();

      try {
        const resolvedFile = resolveGoogleDocFile_(item.file);
        if (!resolvedFile) {
          reviewRows.push(makeSourceReviewRow_({
            status: 'Unsupported source file', databaseTab: databaseSheetName,
            databaseRow: '', sourceFolder: item.sourceFolder,
            sourceFileName: fallbackName, sourceFileUrl: fallbackUrl,
            field: '', databaseValue: '', sourceValue: '',
            notes: 'Not a native Google Doc or shortcut to a Google Doc'
          }));
          totals.unsupported++;
          return;
        }

        const documentId = resolvedFile.getId();
        if (processedDocumentIds.has(documentId)) {
          return;
        }
        processedDocumentIds.add(documentId);

        const sourceInfo = readMonologueHeader_(resolvedFile);
        const sourceRowValues = [
          sourceInfo.playwright,
          resolvedFile.getName(),
          sourceInfo.character,
          sourceInfo.firstLine,
          '', '', '', ''
        ];

        if (!hasRequiredIdentityFields_(sourceRowValues)) {
          throw new Error('Parsed source did not provide Playwright, Character, and First Line');
        }

        const fingerprint = makeFingerprint_(sourceRowValues);
        if (databaseIndex.ambiguousFingerprints.has(fingerprint)) {
          reviewRows.push(makeSourceReviewRow_({
            status: 'Multiple possible database matches', databaseTab: databaseSheetName,
            databaseRow: '', sourceFolder: item.sourceFolder,
            sourceFileName: resolvedFile.getName(), sourceFileUrl: resolvedFile.getUrl(),
            field: '', databaseValue: '', sourceValue: '',
            notes: 'More than one database row has this duplicate key: ' + fingerprint
          }));
          totals.ambiguous++;
          return;
        }

        if (!databaseIndex.fingerprintRows.has(fingerprint)) {
          reviewRows.push(makeSourceReviewRow_({
            status: 'No confident database match', databaseTab: databaseSheetName,
            databaseRow: '', sourceFolder: item.sourceFolder,
            sourceFileName: resolvedFile.getName(), sourceFileUrl: resolvedFile.getUrl(),
            field: '', databaseValue: '', sourceValue: '',
            notes: 'No database row matches Playwright + Character + the first three meaningful First Line words.'
          }));
          totals.noMatch++;
          return;
        }

        const targetRecord = databaseIndex.fingerprintRows.get(fingerprint);
        const differences = getSourceDatabaseDifferences_(targetRecord, {
          playwright: sourceInfo.playwright,
          documentName: resolvedFile.getName(),
          character: sourceInfo.character,
          firstLine: sourceInfo.firstLine,
          documentUrl: resolvedFile.getUrl()
        });

        if (differences.length === 0) {
          totals.unchanged++;
          return;
        }

        differences.forEach(function(difference) {
          reviewRows.push(makeSourceReviewRow_({
            status: 'Review/update', databaseTab: databaseSheetName,
            databaseRow: targetRecord.rowNumber, sourceFolder: item.sourceFolder,
            sourceFileName: resolvedFile.getName(), sourceFileUrl: resolvedFile.getUrl(),
            field: difference.field, databaseValue: difference.databaseValue,
            sourceValue: difference.sourceValue, notes: difference.notes
          }));
        });
        totals.changed++;
      } catch (error) {
        reviewRows.push(makeSourceReviewRow_({
          status: 'Could not parse source', databaseTab: databaseSheetName,
          databaseRow: '', sourceFolder: item.sourceFolder,
          sourceFileName: fallbackName, sourceFileUrl: fallbackUrl,
          field: '', databaseValue: '', sourceValue: '',
          notes: error && error.message ? error.message : String(error)
        }));
        totals.unparseable++;
      }
    });

    writeSourceChangesReviewSheet_(ss, reviewRows);
    SpreadsheetApp.flush();

    showMessage_(
      'Source Changes Review complete.\n\n' +
      'Unchanged source files: ' + totals.unchanged + '\n' +
      'Source records with differences: ' + totals.changed + '\n' +
      'No confident database match: ' + totals.noMatch + '\n' +
      'Multiple possible matches: ' + totals.ambiguous + '\n' +
      'Could not parse: ' + totals.unparseable + '\n' +
      'Unsupported files: ' + totals.unsupported
    );
  } finally {
    lock.releaseLock();
  }
}

function getSourceReviewDatabaseIndex_(sheet) {
  const lastRow = sheet.getLastRow();
  if (lastRow < CONFIG.DATA_START_ROW) {
    return { fingerprintRows: new Map(), ambiguousFingerprints: new Set() };
  }

  const rowCount = lastRow - CONFIG.DATA_START_ROW + 1;
  const values = sheet
    .getRange(CONFIG.DATA_START_ROW, 1, rowCount, CONFIG.DATABASE_COLUMN_COUNT)
    .getDisplayValues();
  const columnBRichText = sheet
    .getRange(CONFIG.DATA_START_ROW, 2, rowCount, 1)
    .getRichTextValues();

  const fingerprintRows = new Map();
  const ambiguousFingerprints = new Set();

  values.forEach(function(row, index) {
    if (!rowHasContent_(row) || !hasRequiredIdentityFields_(row)) {
      return;
    }

    const fingerprint = makeFingerprint_(row);
    const record = {
      rowNumber: CONFIG.DATA_START_ROW + index,
      playwright: row[0],
      documentName: row[1],
      character: row[2],
      firstLine: row[3],
      documentUrl: getFirstRichTextLinkUrl_(columnBRichText[index][0])
    };

    if (fingerprintRows.has(fingerprint)) {
      fingerprintRows.delete(fingerprint);
      ambiguousFingerprints.add(fingerprint);
      return;
    }
    if (!ambiguousFingerprints.has(fingerprint)) {
      fingerprintRows.set(fingerprint, record);
    }
  });

  return { fingerprintRows: fingerprintRows, ambiguousFingerprints: ambiguousFingerprints };
}

function getSourceDatabaseDifferences_(targetRecord, sourceRecord) {
  const differences = [];

  if (normalizeSourceReviewPlaywright_(targetRecord.playwright) !== normalizeSourceReviewPlaywright_(sourceRecord.playwright)) {
    differences.push({ field: 'Playwright', databaseValue: targetRecord.playwright, sourceValue: sourceRecord.playwright, notes: 'Source Playwright differs from database Column A.' });
  }
  if (normalizeSourceReviewCharacter_(targetRecord.character) !== normalizeSourceReviewCharacter_(sourceRecord.character)) {
    differences.push({ field: 'Character', databaseValue: targetRecord.character, sourceValue: sourceRecord.character, notes: 'Source Character differs from database Column C.' });
  }
  if (normalizeSourceReviewFirstLine_(targetRecord.firstLine) !== normalizeSourceReviewFirstLine_(sourceRecord.firstLine)) {
    differences.push({ field: 'First Line', databaseValue: targetRecord.firstLine, sourceValue: sourceRecord.firstLine, notes: 'Full normalized source excerpt differs from database Column D.' });
  }
  if (normalizeSourceReviewGeneric_(targetRecord.documentName) !== normalizeSourceReviewGeneric_(sourceRecord.documentName)) {
    differences.push({ field: 'Play / Mono Link label', databaseValue: targetRecord.documentName, sourceValue: sourceRecord.documentName, notes: 'Current source document name differs from database Column B.' });
  }
  if (
    cleanText_(targetRecord.documentUrl) !== '' &&
    normalizeDriveUrlForReview_(targetRecord.documentUrl) !== normalizeDriveUrlForReview_(sourceRecord.documentUrl)
  ) {
    differences.push({ field: 'Play / Mono Link URL', databaseValue: targetRecord.documentUrl, sourceValue: sourceRecord.documentUrl, notes: 'Current source Google Doc URL differs from database Column B link.' });
  }

  return differences;
}

function normalizeSourceReviewPlaywright_(value) {
  return normalizeSourceReviewGeneric_(value)
    .replace(/^(?:written\s+by|playwright|author|by)\b\s*[:\-–—,]*\s*/i, '')
    .trim();
}

function normalizeSourceReviewCharacter_(value) {
  return normalizeSourceReviewGeneric_(value).replace(/[.:\s]+$/, '').trim();
}

function normalizeSourceReviewFirstLine_(value) {
  return normalizeSourceReviewGeneric_(value)
    .replace(/^[^a-z0-9]+/i, '')
    .replace(/[^a-z0-9]+$/i, '')
    .trim();
}

function normalizeSourceReviewGeneric_(value) {
  return cleanText_(value)
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[“”]/g, '"')
    .replace(/[‘’]/g, "'")
    .replace(/…/g, '...')
    .replace(/[\u2010-\u2015]/g, '-')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeDriveUrlForReview_(url) {
  return cleanText_(url).replace(/\/$/, '').trim();
}

function getFirstRichTextLinkUrl_(richTextValue) {
  if (!richTextValue) {
    return '';
  }
  const wholeCellUrl = richTextValue.getLinkUrl();
  if (wholeCellUrl) {
    return wholeCellUrl;
  }
  const runs = richTextValue.getRuns();
  for (let i = 0; i < runs.length; i++) {
    const runUrl = runs[i].getLinkUrl();
    if (runUrl) {
      return runUrl;
    }
  }
  return '';
}

function makeSourceReviewRow_(item) {
  return [
    item.status || '', item.databaseTab || '', item.databaseRow || '',
    item.sourceFolder || '', item.sourceFileName || '', item.sourceFileUrl || '',
    item.field || '', item.databaseValue || '', item.sourceValue || '', item.notes || ''
  ];
}

function writeSourceChangesReviewSheet_(ss, reviewRows) {
  const sheet = getOrCreateSheet_(ss, CONFIG.SOURCE_CHANGES_REVIEW_SHEET_NAME);
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }
  sheet.clear();

  const headers = [[
    'Status', 'Database tab', 'Database row', 'Source folder', 'Source file',
    'Source URL', 'Field', 'Database value', 'Source value', 'Notes'
  ]];
  sheet.getRange(1, 1, 1, headers[0].length).setValues(headers);

  if (reviewRows.length > 0) {
    sheet.getRange(2, 1, reviewRows.length, headers[0].length).setValues(reviewRows);
    const sourceLinks = reviewRows.map(function(row) {
      const builder = SpreadsheetApp.newRichTextValue().setText(cleanText_(row[4]));
      if (cleanText_(row[5]) !== '') {
        builder.setLinkUrl(cleanText_(row[5]));
      }
      return [builder.build()];
    });
    sheet.getRange(2, 5, sourceLinks.length, 1).setRichTextValues(sourceLinks);
    sheet.getRange(1, 1, reviewRows.length + 1, headers[0].length).createFilter();
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers[0].length).setFontWeight('bold').setBackground('#fff2cc');
  sheet.getRange('D:J').setWrap(true);
  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 230);
  sheet.setColumnWidth(3, 110);
  sheet.setColumnWidth(4, 280);
  sheet.setColumnWidth(5, 300);
  sheet.setColumnWidth(6, 360);
  sheet.setColumnWidth(7, 220);
  sheet.setColumnWidth(8, 360);
  sheet.setColumnWidth(9, 360);
  sheet.setColumnWidth(10, 440);
}

function getDestinationSpreadsheet_() {
  if (cleanText_(SPREADSHEET_ID) !== '') {
    return SpreadsheetApp.openById(cleanText_(SPREADSHEET_ID));
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error(
      'No destination spreadsheet found. Open Apps Script from the target Google Sheet or set SPREADSHEET_ID near the top of the script.'
    );
  }
  return ss;
}

function getOrCreateSheet_(ss, sheetName) {
  let sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  }
  return sheet;
}

function makeSafeTabName_(name) {
  let safeName = String(name)
    .replace(/[:\\/?*\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
  if (safeName === '') {
    safeName = 'Unnamed Folder';
  }
  return safeName.substring(0, CONFIG.MAX_SHEET_NAME_LENGTH);
}

function cleanText_(value) {
  return String(value === null || value === undefined ? '' : value).trim();
}

function normalizeText_(text) {
  return String(text || '').replace(/\s+/g, ' ').trim();
}

function escapeRegExp_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function showMessage_(message) {
  SpreadsheetApp.getUi().alert(message);
}

function diagnoseGoogleDocLayout() {
  const response = Browser.inputBox(
    'Diagnose Google Doc Layout',
    'Paste the full Google Docs URL or only the document ID:',
    Browser.Buttons.OK_CANCEL
  );

  if (
    response === 'cancel' ||
    cleanText_(response) === ''
  ) {
    return;
  }

  const input = cleanText_(response);

  /*
   * First, extract the ID from a normal Google Docs URL:
   * https://docs.google.com/document/d/DOCUMENT_ID/edit
   */
  const urlMatch = input.match(
    /\/document\/d\/([^\/?#\s]+)/
  );

  /*
   * Otherwise allow a raw ID pasted by itself.
   */
  const documentId = urlMatch
    ? urlMatch[1]
    : input;

  if (
    documentId.length < 20 ||
    /[\s/]/.test(documentId)
  ) {
    SpreadsheetApp.getUi().alert(
      'Could not find a valid Google Doc ID.\\n\\n' +
      'Paste either:\\n' +
      '• the whole Google Docs URL, or\\n' +
      '• only the ID between /d/ and /edit.'
    );
    return;
  }

  let doc;

  try {
    doc = DocumentApp.openById(documentId);
  } catch (error) {
    SpreadsheetApp.getUi().alert(
      'Apps Script could not open that Google Doc.\\n\\n' +
      'Check that it is a Google Doc and that the account running the script ' +
      'has permission to open it.\\n\\n' +
      'Details: ' +
      (error && error.message ? error.message : String(error))
    );
    return;
  }

  const body = doc.getBody();
  const ss = getDestinationSpreadsheet_();

  const reportName = 'Source Layout Diagnostic';
  let report = ss.getSheetByName(reportName);

  if (!report) {
    report = ss.insertSheet(reportName);
  }

  if (report.getFilter()) {
    report.getFilter().remove();
  }

  report.clear();

  report.getRange('A1:E1').setValues([[
    'Document name',
    'Element index',
    'Element type',
    'Text visible to Apps Script',
    'Notes'
  ]]);

  const rows = [];
  const childCount = body.getNumChildren();

  for (let i = 0; i < childCount; i++) {
    const child = body.getChild(i);
    const elementType = String(child.getType());
    let text = '';
    let notes = '';

    try {
      if (typeof child.getText === 'function') {
        text = String(child.getText() || '');
      } else {
        notes = 'Element has no direct getText() method';
      }
    } catch (error) {
      notes = error && error.message
        ? error.message
        : String(error);
    }

    rows.push([
      doc.getName(),
      i,
      elementType,
      text,
      notes
    ]);
  }

  if (rows.length > 0) {
    report
      .getRange(2, 1, rows.length, 5)
      .setValues(rows);
  }

  report.setFrozenRows(1);

  report
    .getRange('A1:E1')
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  report.getRange('D:E').setWrap(true);

  report.setColumnWidth(1, 280);
  report.setColumnWidth(2, 110);
  report.setColumnWidth(3, 180);
  report.setColumnWidth(4, 760);
  report.setColumnWidth(5, 360);

  if (rows.length > 0) {
    report
      .getRange(1, 1, rows.length + 1, 5)
      .createFilter();
  }

  SpreadsheetApp.getUi().alert(
    'Diagnostic complete. Open the \"' +
    reportName +
    '\" tab and copy the rows for the document here.'
  );
}

function normalizedFirstWordsForDuplicateKey_(text, wordLimit) {
  const limit = wordLimit || 3;

  return String(text || '')
    .normalize('NFKD')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim()
    .split(/\s+/)
    .filter(function(word) {
      return word !== '';
    })
    .slice(0, limit)
    .join(' ');
}

function monologueDuplicateKey_(playwright, play, character, firstLine) {
  return [
    normalizeForMatch_(playwright),
    normalizeForMatch_(play),
    normalizeForMatch_(character),
    normalizedFirstWordsForDuplicateKey_(firstLine, 3)
  ].join(' | ');
}

/**
 * Builds a report that compares every Google Doc in FOLDER_ID and all
 * subfolders against source links recorded in Column B of every database tab.
 *
 * This is report-only. It does not add, delete, modify, or move any
 * database rows or Drive files.
 */
function buildFolderDatabaseInventory() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    validateFolderId_();

    const ss = getDestinationSpreadsheet_();
    const rootFolder = DriveApp.getFolderById(FOLDER_ID);
    const rootFolderName = rootFolder.getName();

    /*
     * Get all Google Docs and shortcuts within the configured folder tree.
     */
    const folderItems = [];
    collectFilesRecursively_(rootFolder, '', folderItems);

    /*
     * Index every database source link from Column B.
     *
     * The inventory considers all sheets that have the standard A:H
     * monologue-database headers, excluding Manual Entries and reports.
     */
    const databaseIndex = buildDatabaseSourceLinkIndex_(ss);

    const seenDocumentIds = new Set();
    const inventoryRows = [];
    const totals = {
      sourceGoogleDocs: 0,
      represented: 0,
      missing: 0,
      duplicateFolderReferences: 0,
      unsupported: 0,
      databaseRowsWithoutLinks: databaseIndex.rowsWithoutSourceLinks.length
    };

    folderItems.forEach(function(item) {
      const originalFile = item.file;
      const resolvedFile = resolveGoogleDocFile_(originalFile);

      if (!resolvedFile) {
        inventoryRows.push([
          'Unsupported source file',
          rootFolderName,
          item.sourceFolder,
          originalFile.getName(),
          originalFile.getUrl(),
          '',
          '',
          '',
          '',
          'Not a native Google Doc or shortcut to a Google Doc'
        ]);
        totals.unsupported++;
        return;
      }

      const documentId = resolvedFile.getId();
      const documentUrl = resolvedFile.getUrl();

      /*
       * A Google Doc can appear in multiple folder locations through
       * shortcuts. List the extra reference but do not count it as a
       * separate missing database record.
       */
      if (seenDocumentIds.has(documentId)) {
        inventoryRows.push([
          'Duplicate folder reference',
          rootFolderName,
          item.sourceFolder,
          resolvedFile.getName(),
          documentUrl,
          documentId,
          '',
          '',
          '',
          'The same Google Doc appears more than once in the folder tree.'
        ]);
        totals.duplicateFolderReferences++;
        return;
      }

      seenDocumentIds.add(documentId);
      totals.sourceGoogleDocs++;

      const databaseMatches =
        databaseIndex.rowsByDocumentId.get(documentId) || [];

      if (databaseMatches.length === 0) {
        inventoryRows.push([
          'Missing from database',
          rootFolderName,
          item.sourceFolder,
          resolvedFile.getName(),
          documentUrl,
          documentId,
          '',
          '',
          '',
          'No database row has a Column B source link to this Google Doc.'
        ]);
        totals.missing++;
        return;
      }

      if (databaseMatches.length === 1) {
        const match = databaseMatches[0];

        inventoryRows.push([
          'Represented in database',
          rootFolderName,
          item.sourceFolder,
          resolvedFile.getName(),
          documentUrl,
          documentId,
          match.sheetName,
          match.rowNumber,
          match.playwright + ' | ' + match.character,
          ''
        ]);
        totals.represented++;
        return;
      }

      /*
       * Several rows link to the same source Doc. This may be intentional,
       * but it should be visible in the report.
       */
      inventoryRows.push([
        'Multiple database rows for source',
        rootFolderName,
        item.sourceFolder,
        resolvedFile.getName(),
        documentUrl,
        documentId,
        databaseMatches.map(function(match) {
          return match.sheetName;
        }).join(' | '),
        databaseMatches.map(function(match) {
          return match.rowNumber;
        }).join(', '),
        databaseMatches.map(function(match) {
          return match.playwright + ' | ' + match.character;
        }).join(' || '),
        'This Google Doc is linked from more than one database row.'
      ]);
    });

    /*
     * Also list database rows that cannot be verified against a folder
     * source because Column B has no Drive link.
     */
    databaseIndex.rowsWithoutSourceLinks.forEach(function(row) {
      inventoryRows.push([
        'Database row has no source link',
        '',
        '',
        '',
        '',
        '',
        row.sheetName,
        row.rowNumber,
        row.playwright + ' | ' + row.character,
        'Column B contains no Google Drive document link.'
      ]);
    });

    writeFolderDatabaseInventorySheet_(
      ss,
      inventoryRows,
      totals,
      rootFolderName
    );

    SpreadsheetApp.flush();

    showMessage_(
      'Folder/database inventory complete.\n\n' +
      'Unique Google Docs in folder tree: ' +
      totals.sourceGoogleDocs + '\n' +
      'Represented in database: ' +
      totals.represented + '\n' +
      'Missing from database: ' +
      totals.missing + '\n' +
      'Duplicate folder references: ' +
      totals.duplicateFolderReferences + '\n' +
      'Unsupported source files: ' +
      totals.unsupported + '\n' +
      'Database rows without source links: ' +
      totals.databaseRowsWithoutLinks + '\n\n' +
      'Open the "Folder Database Inventory" tab and filter the Status column.'
    );
  } finally {
    lock.releaseLock();
  }
}


/**
 * Reads all recognized database tabs and creates an index of the
 * Google Doc IDs found in the rich-text links in Column B.
 */
function buildDatabaseSourceLinkIndex_(ss) {
  const rowsByDocumentId = new Map();
  const rowsWithoutSourceLinks = [];

  ss.getSheets()
    .filter(isMonologueDatabaseSheet_)
    .forEach(function(sheet) {
      const lastRow = sheet.getLastRow();

      if (lastRow < CONFIG.DATA_START_ROW) {
        return;
      }

      const rowCount = lastRow - CONFIG.DATA_START_ROW + 1;

      const values = sheet
        .getRange(
          CONFIG.DATA_START_ROW,
          1,
          rowCount,
          CONFIG.DATABASE_COLUMN_COUNT
        )
        .getDisplayValues();

      const columnBRichText = sheet
        .getRange(
          CONFIG.DATA_START_ROW,
          2,
          rowCount,
          1
        )
        .getRichTextValues();

      values.forEach(function(row, index) {
        if (!rowHasContent_(row)) {
          return;
        }

        const rowNumber = CONFIG.DATA_START_ROW + index;
        const sourceUrl = getFirstRichTextLinkUrl_(
          columnBRichText[index][0]
        );

        const record = {
          sheetName: sheet.getName(),
          rowNumber: rowNumber,
          playwright: cleanText_(row[0]),
          playLabel: cleanText_(row[1]),
          character: cleanText_(row[2]),
          firstLine: cleanText_(row[3]),
          sourceUrl: sourceUrl
        };

        const documentId = extractGoogleDocumentId_(sourceUrl);

        if (!documentId) {
          rowsWithoutSourceLinks.push(record);
          return;
        }

        if (!rowsByDocumentId.has(documentId)) {
          rowsByDocumentId.set(documentId, []);
        }

        rowsByDocumentId.get(documentId).push(record);
      });
    });

  return {
    rowsByDocumentId: rowsByDocumentId,
    rowsWithoutSourceLinks: rowsWithoutSourceLinks
  };
}


/**
 * Gets the document ID from ordinary Google Docs URLs, including:
 *
 * https://docs.google.com/document/d/ID/edit
 * https://docs.google.com/document/d/ID/view
 *
 * Returns an empty string for non-Google-Doc links.
 */
function extractGoogleDocumentId_(url) {
  const text = cleanText_(url);

  if (text === '') {
    return '';
  }

  const match = text.match(/\/document\/d\/([^\/?#\s]+)/);

  return match ? match[1] : '';
}


/**
 * Writes the report into a dedicated inventory tab.
 */
function writeFolderDatabaseInventorySheet_(
  ss,
  inventoryRows,
  totals,
  rootFolderName
) {
  const sheetName = 'Folder Database Inventory';
  const sheet = getOrCreateSheet_(ss, sheetName);

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.clear();

  const summaryRows = [
    ['Folder/database inventory', rootFolderName],
    ['Unique Google Docs in folder tree', totals.sourceGoogleDocs],
    ['Represented in database', totals.represented],
    ['Missing from database', totals.missing],
    ['Duplicate folder references', totals.duplicateFolderReferences],
    ['Unsupported source files', totals.unsupported],
    ['Database rows without source links', totals.databaseRowsWithoutLinks]
  ];

  sheet.getRange(1, 1, summaryRows.length, 2).setValues(summaryRows);

  const headerRow = summaryRows.length + 2;

  const headers = [[
    'Status',
    'Root folder',
    'Source folder',
    'Source file',
    'Source URL',
    'Google Doc ID',
    'Database tab',
    'Database row',
    'Database identity',
    'Notes'
  ]];

  sheet
    .getRange(headerRow, 1, 1, headers[0].length)
    .setValues(headers);

  if (inventoryRows.length > 0) {
    sheet
      .getRange(
        headerRow + 1,
        1,
        inventoryRows.length,
        headers[0].length
      )
      .setValues(inventoryRows);

    /*
     * Make the source-file name a clickable link in Column D.
     */
    const sourceFileLinks = inventoryRows.map(function(row) {
      const fileName = cleanText_(row[3]);
      const fileUrl = cleanText_(row[4]);
      const builder = SpreadsheetApp.newRichTextValue()
        .setText(fileName);

      if (fileName !== '' && fileUrl !== '') {
        builder.setLinkUrl(fileUrl);
      }

      return [builder.build()];
    });

    sheet
      .getRange(headerRow + 1, 4, sourceFileLinks.length, 1)
      .setRichTextValues(sourceFileLinks);

    sheet
      .getRange(
        headerRow,
        1,
        inventoryRows.length + 1,
        headers[0].length
      )
      .createFilter();
  }

  sheet.setFrozenRows(headerRow);

  sheet
    .getRange(1, 1, summaryRows.length, 2)
    .setFontWeight('bold');

  sheet
    .getRange(headerRow, 1, 1, headers[0].length)
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.getRange('A:J').setWrap(true);

  sheet.setColumnWidth(1, 250);
  sheet.setColumnWidth(2, 220);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 360);
  sheet.setColumnWidth(6, 300);
  sheet.setColumnWidth(7, 240);
  sheet.setColumnWidth(8, 110);
  sheet.setColumnWidth(9, 320);
  sheet.setColumnWidth(10, 420);
}

/*
 * INTERACTIVE PLAY-TITLE DUPLICATE REVIEW
 *
 * Matches database rows by the displayed value in Column B:
 * Play Mono Link.
 *
 * It asks, pair by pair:
 * 1. Are these actually duplicates?
 * 2. If yes, which row should be kept?
 *
 * Before deleting, it copies only missing E:H values from the
 * deleted row into blank E:H cells in the kept row.
 *
 * It never changes rows unless you explicitly choose Delete row 1
 * or Delete row 2 for that exact pair.
 */

function reviewMatchingPlayTitles() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const ss = getDestinationSpreadsheet();
    const ui = SpreadsheetApp.getUi();

   const activeSheet = ss.getActiveSheet();

const sheets = [activeSheet];
    

    const candidates = findMatchingPlayTitlePairs_(sheets);

    if (candidates.length === 0) {
      ui.alert(
        'Matching-title review complete',
        'No matching Play Mono Link titles were found.',
        ui.ButtonSet.OK
      );
      return;
    }

    let reviewed = 0;
    let deleted = 0;
    let skipped = 0;
    let canceled = false;

    for (let i = 0; i < candidates.length; i++) {
      const pair = candidates[i];

      /*
       * Row numbers can change after earlier deletions.
       * Re-find both rows from their stable full-row signatures
       * before showing or changing the pair.
       */
      const currentPair = findCurrentTitlePair_(ss, pair);

      if (!currentPair) {
        continue;
      }

      reviewed++;

      const description =
        'Possible matching play title (' +
        reviewed +
        ' of ' +
        candidates.length +
        ')\n\n' +
        'Normalized title:\n' +
        currentPair.title +
        '\n\n' +
        formatPlayTitleCandidate_('ROW 1', currentPair.first) +
        '\n\n' +
        formatPlayTitleCandidate_('ROW 2', currentPair.second) +
        '\n\n' +
        'Are these two rows duplicates of the same monologue entry?';

      const isDuplicate = ui.alert(
        'Review matching play titles',
        description,
        ui.ButtonSet.YES_NO_CANCEL
      );

      if (isDuplicate === ui.Button.CANCEL) {
        canceled = true;
        break;
      }

      if (isDuplicate === ui.Button.NO) {
        skipped++;
        continue;
      }

      const keepChoice = ui.alert(
        'Choose the row to keep',
        'Keep Row 1, delete Row 2?\n\n' +
          'YES = Keep Row 1 / Delete Row 2\n' +
          'NO = Keep Row 2 / Delete Row 1\n' +
          'CANCEL = Do not delete either row',
        ui.ButtonSet.YES_NO_CANCEL
      );

      if (keepChoice === ui.Button.CANCEL) {
        skipped++;
        continue;
      }

      const keep = keepChoice === ui.Button.YES
        ? currentPair.first
        : currentPair.second;

      const remove = keepChoice === ui.Button.YES
        ? currentPair.second
        : currentPair.first;

      /*
       * Preserve database annotations:
       * only fill blank E:H cells in the kept row.
       */
      mergeMissingDuplicateRowData_(
        keep.sheet,
        remove.rowNumber,
        keep.rowNumber
      );

      keep.sheet.deleteRow(remove.rowNumber);
      deleted++;
    }

    SpreadsheetApp.flush();

    ui.alert(
      'Matching-title review complete',
      'Pairs reviewed: ' + reviewed +
      '\nRows deleted: ' + deleted +
      '\nPairs skipped: ' + skipped +
      (canceled
        ? '\n\nReview stopped early. You can run it again later.'
        : ''),
      ui.ButtonSet.OK
    );

  } finally {
    lock.releaseLock();
  }
}


/*
 * Finds every pair of rows in the same database tab whose Column B
 * displayed values normalize to the same title.
 *
 * It does NOT compare across separate database tabs, because those
 * may intentionally contain the same source title in different
 * categories.
 */
function findMatchingPlayTitlePairs_(databaseSheets) {
  const pairs = [];

  databaseSheets.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();

    if (lastRow < CONFIG.DATASTARTROW) {
      return;
    }

    const rowCount = lastRow - CONFIG.DATASTARTROW + 1;

    const values = sheet
      .getRange(
        CONFIG.DATASTARTROW,
        1,
        rowCount,
        CONFIG.DATABASECOLUMNCOUNT
      )
      .getDisplayValues();

    const groups = new Map();

    values.forEach(function(row, index) {
      if (!rowHasContent(row)) {
        return;
      }

      const playTitle = cleanText(row[1]);

      if (playTitle === '') {
        return;
      }

      const titleKey = normalizePlayTitleForReview_(playTitle);

      if (titleKey === '') {
        return;
      }

      const record = {
        sheetName: sheet.getName(),
        sheet: sheet,
        rowNumber: CONFIG.DATASTARTROW + index,
        title: playTitle,
        titleKey: titleKey,
        displayValues: row,
        signature: makeFullRowSignature_(row)
      };

      if (!groups.has(titleKey)) {
        groups.set(titleKey, []);
      }

      groups.get(titleKey).push(record);
    });

    groups.forEach(function(group, titleKey) {
      if (group.length < 2) {
        return;
      }

      /*
       * Make every unique row-pair available for review.
       * If three rows share the same title, you may be asked about
       * 1 vs 2, 1 vs 3, and 2 vs 3 as appropriate.
       */
      for (let firstIndex = 0; firstIndex < group.length - 1; firstIndex++) {
        for (
          let secondIndex = firstIndex + 1;
          secondIndex < group.length;
          secondIndex++
        ) {
          pairs.push({
            title: group[firstIndex].title,
            titleKey: titleKey,
            first: group[firstIndex],
            second: group[secondIndex]
          });
        }
      }
    });
  });

  return pairs;
}


/*
 * After rows may have been deleted, locate the same two original rows
 * again by their whole A:H displayed content.
 */
function findCurrentTitlePair_(ss, originalPair) {
  const sheet = ss.getSheetByName(originalPair.first.sheetName);

  if (!sheet) {
    return null;
  }

  const first = findRowBySignature_(
    sheet,
    originalPair.first.signature
  );

  const second = findRowBySignature_(
    sheet,
    originalPair.second.signature
  );

  if (!first || !second || first.rowNumber === second.rowNumber) {
    return null;
  }

  return {
    title: originalPair.title,
    first: first,
    second: second
  };
}


function findRowBySignature_(sheet, signature) {
  const lastRow = sheet.getLastRow();

  if (lastRow < CONFIG.DATASTARTROW) {
    return null;
  }

  const rowCount = lastRow - CONFIG.DATASTARTROW + 1;

  const values = sheet
    .getRange(
      CONFIG.DATASTARTROW,
      1,
      rowCount,
      CONFIG.DATABASECOLUMNCOUNT
    )
    .getDisplayValues();

  for (let index = 0; index < values.length; index++) {
    const row = values[index];

    if (makeFullRowSignature_(row) === signature) {
      return {
        sheetName: sheet.getName(),
        sheet: sheet,
        rowNumber: CONFIG.DATASTARTROW + index,
        title: cleanText(row[1]),
        displayValues: row,
        signature: signature
      };
    }
  }

  return null;
}


function makeFullRowSignature_(row) {
  return row
    .slice(0, CONFIG.DATABASECOLUMNCOUNT)
    .map(function(value) {
      return cleanText(value)
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    })
    .join('\u241F');
}


/*
 * Normalizes only for comparison:
 * - ignores case
 * - treats curly/straight apostrophes and quotation marks alike
 * - ignores punctuation, extra spaces, and a trailing "(1)"-style
 *   copy indicator in a source-document name
 *
 * Displayed sheet values remain unchanged.
 */
function normalizePlayTitleForReview_(title) {
  return cleanText(title)
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[’‘]/g, "'")
    .replace(/[“”]/g, '"')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .trim();
}

function formatPlayTitleCandidate_(label, record) {
  const row = record.displayValues;

  return (
    label +
    ': ' +
    record.sheetName +
    '!Row ' +
    record.rowNumber +
    '\n' +
    'Playwright: ' +
    cleanText(row[0]) +
    '\n' +
    'Play Mono Link: ' +
    cleanText(row[1]) +
    '\n' +
    'Character: ' +
    cleanText(row[2]) +
    '\n' +
    'First Line: ' +
    cleanText(row[3]) +
    '\n' +
    'Identifiers: ' +
    cleanText(row[4]) +
    '\n' +
    'Themes: ' +
    cleanText(row[5]) +
    '\n' +
    'Considerations: ' +
    cleanText(row[6]) +
    '\n' +
    'Link to Play: ' +
    cleanText(row[7])
  );
}
function getDestinationSpreadsheet() {
  if (
    typeof SPREADSHEETID !== 'undefined' &&
    String(SPREADSHEETID || '').trim() !== ''
  ) {
    return SpreadsheetApp.openById(
      String(SPREADSHEETID).trim()
    );
  }

  const spreadsheet = SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'No destination spreadsheet found. Open Apps Script from the target ' +
      'Google Sheet, or set SPREADSHEETID near the top of the script.'
    );
  }

  return spreadsheet;
}
function isMonologueDatabaseSheet(sheet) {
  const excludedSheetNames = [
    CONFIG.MANUALENTRIESSHEETNAME,
    CONFIG.SKIPPEDSHEETNAME,
    CONFIG.SOURCECHANGESREVIEWSHEETNAME
  ];

  if (excludedSheetNames.indexOf(sheet.getName()) !== -1) {
    return false;
  }

  const headers = sheet
    .getRange(1, 1, 1, CONFIG.DATABASECOLUMNCOUNT)
    .getDisplayValues()[0]
    .map(cleanText);

  return CONFIG.VISIBLECOLUMNS.every(function(expectedHeader, index) {
    return headers[index] === expectedHeader;
  });
}

/*
 * INTERACTIVE SAME-FILE-NAME REVIEW
 *
 * Looks only at database tabs whose A1:H1 headers match the expected
 * monologue database layout. It compares Column B, Play Mono Link,
 * exactly after light normalization (case, whitespace, quote style).
 *
 * IMPORTANT:
 * - "The Moors" and "The Moors (1)" are different.
 * - Nothing is deleted unless you explicitly choose which row to delete.
 * - Before deleting, missing E:H data is copied from the deleted row
 *   into blank E:H cells of the retained row.
 */

function reviewMatchingPlayTitles() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ui = SpreadsheetApp.getUi();

  if (!ss) {
    throw new Error(
      'Open this script from the Google Sheet that contains the database.'
    );
  }

  const sheets = ss.getSheets().filter(function(sheet) {
    return isReviewDatabaseTab_(sheet);
  });

  if (sheets.length === 0) {
    ui.alert(
      'No database tabs found',
      'No tabs were found with the expected A1:H1 database headers.',
      ui.ButtonSet.OK
    );
    return;
  }

  const candidates = getSamePlayNamePairs_(sheets);

  if (candidates.length === 0) {
    ui.alert(
      'Review complete',
      'No matching Column B file/play names were found.',
      ui.ButtonSet.OK
    );
    return;
  }

  let reviewed = 0;
  let deleted = 0;
  let skipped = 0;
  let stoppedEarly = false;

  for (let i = 0; i < candidates.length; i++) {
    const originalPair = candidates[i];

    /*
     * Re-find rows each time. This prevents prior deletions from causing
     * row-number shifts that would affect the wrong row later.
     */
    const currentPair = findCurrentReviewPair_(ss, originalPair);

    if (!currentPair) {
      continue;
    }

    reviewed++;

    const question =
      'Possible duplicate file/play name (' +
      reviewed +
      ' of ' +
      candidates.length +
      ')\n\n' +
      'Column B title: ' +
      currentPair.title +
      '\n\n' +
      describeReviewRow_('ROW 1', currentPair.first) +
      '\n\n' +
      describeReviewRow_('ROW 2', currentPair.second) +
      '\n\n' +
      'Are these duplicate database entries?';

    const duplicateAnswer = ui.alert(
      'Review matching file/play names',
      question,
      ui.ButtonSet.YES_NO_CANCEL
    );

    if (duplicateAnswer === ui.Button.CANCEL) {
      stoppedEarly = true;
      break;
    }

    if (duplicateAnswer === ui.Button.NO) {
      skipped++;
      continue;
    }

   const keepChoice = ui.prompt(
  'Select the row to keep',
  'Enter 1 to keep ROW 1 and delete ROW 2.\n' +
    'Enter 2 to keep ROW 2 and delete ROW 1.\n\n' +
    'The unselected row will be deleted.',
  ui.ButtonSet.OK_CANCEL
);


if (keepChoice.getSelectedButton() === ui.Button.CANCEL) {
  skipped++;
  continue;
}


const selectedRow = String(keepChoice.getResponseText() || '').trim();


if (selectedRow !== '1' && selectedRow !== '2') {
  ui.alert(
    'Invalid selection',
    'No row was deleted. Enter 1 for ROW 1 or 2 for ROW 2 when you run the review again.',
    ui.ButtonSet.OK
  );
  skipped++;
  continue;
}


const keepRow =
  selectedRow === '1'
    ? currentPair.first
    : currentPair.second;


const deleteRow =
  selectedRow === '1'
    ? currentPair.second
    : currentPair.first;

    copyMissingAnnotationsBeforeDelete_(
      keepRow.sheet,
      deleteRow.rowNumber,
      keepRow.rowNumber
    );

    deleteRow.sheet.deleteRow(deleteRow.rowNumber);
    deleted++;
  }

  SpreadsheetApp.flush();

  ui.alert(
    'Matching-title review finished',
    'Pairs reviewed: ' + reviewed +
      '\nRows deleted: ' + deleted +
      '\nPairs skipped: ' + skipped +
      (stoppedEarly
        ? '\n\nThe review stopped early. Run it again later to continue.'
        : ''),
    ui.ButtonSet.OK
  );
}


/*
 * A valid database tab must have exactly these eight headers in A1:H1.
 */
/*
 * Returns true for a database tab that contains the core database columns.
 *
 * It:
 * - Excludes known non-database workflow tabs.
 * - Ignores capitalization, extra spaces, punctuation, and line breaks.
 * - Accepts reasonable header-name variations.
 * - Does NOT require the eight database columns to be in A:H exactly.
 */
function isReviewDatabaseTab_(sheet) {
  const excludedNames = [
    'manual entries',
    'skipped files',
    'source changes review',
    'source layout diagnostic',
    'import staging',
    'import review',
    'import log'
  ];

  const normalizedTabName = normalizeTabLabel_(sheet.getName());

  if (excludedNames.indexOf(normalizedTabName) !== -1) {
    return false;
  }

  const lastColumn = sheet.getLastColumn();

  if (lastColumn === 0) {
    return false;
  }

  const headers = sheet
    .getRange(1, 1, 1, lastColumn)
    .getDisplayValues()[0]
    .map(normalizeTabLabel_);

  const hasPlaywright = hasAnyHeader_(headers, [
    'playwright',
    'author',
    'writer'
  ]);

  const hasPlayName = hasAnyHeader_(headers, [
    'play mono link',
    'playmonolink',
    'file play name',
    'file name',
    'play name',
    'title',
    'play title'
  ]);

  const hasCharacter = hasAnyHeader_(headers, [
    'character',
    'character name'
  ]);

  const hasFirstLine = hasAnyHeader_(headers, [
    'first line',
    'firstline',
    'opening line',
    'first words'
  ]);

  /*
   * Require the four columns that identify a real monologue database row.
   * Do NOT require annotation columns, because some legitimate database
   * tabs may not have every optional note column.
   */
  return hasPlaywright && hasPlayName && hasCharacter && hasFirstLine;
}


/*
 * Makes names safe to compare:
 * "First Line", "FIRST-LINE", and " first line " all become "first line".
 */
function normalizeTabLabel_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .toLowerCase()
    .replace(/[^a-z0-9]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


/*
 * Returns true if the normalized header list contains any allowed label.
 */
function hasAnyHeader_(headers, allowedLabels) {
  return allowedLabels.some(function(label) {
    const normalizedAllowedLabel = normalizeTabLabel_(label);

    return headers.indexOf(normalizedAllowedLabel) !== -1;
  });
}

/*
 * Group entries within each individual tab by the Column B displayed
 * text. Matching is case-insensitive and ignores harmless quote/space
 * differences, but preserves all numbers.
 *
 * Thus:
 * The Moors      != The Moors (1)
 * The Moors (1)  != The Moors (2)
 */
function getSamePlayNamePairs_(sheets) {
  const pairs = [];

  sheets.forEach(function(sheet) {
    const lastRow = sheet.getLastRow();

    if (lastRow < 2) {
      return;
    }

    const rowCount = lastRow - 1;

    const values = sheet
      .getRange(2, 1, rowCount, 8)
      .getDisplayValues();

    const groups = {};

    values.forEach(function(row, index) {
      const hasContent = row.some(function(value) {
        return String(value == null ? '' : value).trim() !== '';
      });

      if (!hasContent) {
        return;
      }

      const columnBTitle = String(row[1] == null ? '' : row[1]).trim();

      if (columnBTitle === '') {
        return;
      }

      const titleKey = normalizeReviewTitle_(columnBTitle);

      if (titleKey === '') {
        return;
      }

      const record = {
        sheetName: sheet.getName(),
        rowNumber: index + 2,
        title: columnBTitle,
        values: row,
        signature: makeReviewRowSignature_(row)
      };

      if (!groups[titleKey]) {
        groups[titleKey] = [];
      }

      groups[titleKey].push(record);
    });

    Object.keys(groups).forEach(function(titleKey) {
      const group = groups[titleKey];

      if (group.length < 2) {
        return;
      }

      for (let i = 0; i < group.length - 1; i++) {
        for (let j = i + 1; j < group.length; j++) {
          pairs.push({
            sheetName: sheet.getName(),
            title: group[i].title,
            first: group[i],
            second: group[j]
          });
        }
      }
    });
  });

  return pairs;
}


/*
 * Re-find both original records in the current sheet after earlier
 * deletes may have changed row numbers.
 */
function findCurrentReviewPair_(ss, originalPair) {
  const sheet = ss.getSheetByName(originalPair.sheetName);

  if (!sheet) {
    return null;
  }

  const first = findReviewRowBySignature_(
    sheet,
    originalPair.first.signature
  );

  const second = findReviewRowBySignature_(
    sheet,
    originalPair.second.signature
  );

  if (!first || !second || first.rowNumber === second.rowNumber) {
    return null;
  }

  return {
    title: originalPair.title,
    first: first,
    second: second
  };
}


function findReviewRowBySignature_(sheet, signature) {
  const lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return null;
  }

  const values = sheet
    .getRange(2, 1, lastRow - 1, 8)
    .getDisplayValues();

  for (let index = 0; index < values.length; index++) {
    if (makeReviewRowSignature_(values[index]) === signature) {
      return {
        sheet: sheet,
        sheetName: sheet.getName(),
        rowNumber: index + 2,
        title: String(values[index][1] == null ? '' : values[index][1]).trim(),
        values: values[index],
        signature: signature
      };
    }
  }

  return null;
}


function makeReviewRowSignature_(row) {
  return row
    .slice(0, 8)
    .map(function(value) {
      return String(value == null ? '' : value)
        .normalize('NFKD')
        .replace(/[\u0300-\u036f]/g, '')
        .replace(/\s+/g, ' ')
        .trim();
    })
    .join('\u241F');
}


/*
 * Preserve digits. Do not strip "(1)", "(2)", years, or any other
 * file-identifying number.
 */
function normalizeReviewTitle_(value) {
  return String(value == null ? '' : value)
    .normalize('NFKD')
    .replace(/[\u0300-\u036f]/g, '')
    .replace(/[‘’]/g, "'")
    .replace(/[“”]/g, '"')
    .toLowerCase()
    .replace(/\s+/g, ' ')
    .trim();
}


function describeReviewRow_(label, record) {
  const row = record.values;

  return (
    label +
    ': ' +
    record.sheetName +
    ' — spreadsheet row ' +
    record.rowNumber +
    '\n' +
    'Playwright: ' + String(row[0] || '') +
    '\n' +
    'File/play name: ' + String(row[1] || '') +
    '\n' +
    'Character: ' + String(row[2] || '') +
    '\n' +
    'First Line: ' + String(row[3] || '') +
    '\n' +
    'Identifiers: ' + String(row[4] || '') +
    '\n' +
    'Themes: ' + String(row[5] || '') +
    '\n' +
    'Considerations: ' + String(row[6] || '') +
    '\n' +
    'Link to Play: ' + String(row[7] || '')
  );
}


/*
 * Transfer only missing annotations before a confirmed deletion:
 * E = Identifiers
 * F = Themes
 * G = Considerations
 * H = Link to Play
 *
 * Existing values in the kept row are never overwritten.
 */
function copyMissingAnnotationsBeforeDelete_(
  sheet,
  sourceRowNumber,
  keepRowNumber
) {
  const sourceValues = sheet
    .getRange(sourceRowNumber, 5, 1, 4)
    .getDisplayValues()[0];

  const keepValues = sheet
    .getRange(keepRowNumber, 5, 1, 4)
    .getDisplayValues()[0];

  for (let offset = 0; offset < 4; offset++) {
    const sourceValue = String(
      sourceValues[offset] == null ? '' : sourceValues[offset]
    ).trim();

    const keepValue = String(
      keepValues[offset] == null ? '' : keepValues[offset]
    ).trim();

    if (sourceValue !== '' && keepValue === '') {
      sheet
        .getRange(sourceRowNumber, 5 + offset)
        .copyTo(
          sheet.getRange(keepRowNumber, 5 + offset),
          SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
          false
        );
    }
  }
}