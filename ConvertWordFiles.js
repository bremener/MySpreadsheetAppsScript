/*
 * WORD TO GOOGLE DOCS CONVERTER - FIXED
 *
 * Target folder:
 * 1YKDbLw9IY1ay0nHN1XzD21X71sP65Gm2
 *
 * What it does:
 * - Searches this folder and all subfolders.
 * - Finds actual .doc and .docx files.
 * - Ignores a Word file if a native Google Doc with the same base name
 *   already exists in the same folder.
 * - Converts only Word files that do not have such a native Google Doc.
 * - Keeps original Word files unchanged.
 * - Converts up to 25 unmatched files per run.
 * - Runs automatically every minute until no unmatched Word files remain.
 * - Writes results directly to a report Sheet.
 *
 * BEFORE RUNNING:
 * 1. Apps Script -> Services -> + -> Drive API -> Add.
 * 2. Save.
 * 3. Delete old converter triggers in Apps Script -> Triggers.
 * 4. Run startFinalWordConversion once.
 */

var FINAL_SOURCE_FOLDER_ID =
  '17RXoytCU5sxfotfqK52QHDp1ML7EmIRp';

var FINAL_SPREADSHEET_ID =
  '1lfqBwaXUZ62xPimJRnKgAUSebhhEmbGPFvzPMuIlCa4';

var FINAL_CONFIG = {
  INCLUDE_SUBFOLDERS: true,
  FILES_PER_RUN: 25,
  PROGRESS_SHEET_NAME: 'Word Conversion Progress',
  REPORT_SHEET_NAME: 'Word Conversion Report',
  PROPERTY_PREFIX: 'finalWordConversion.v1.'
};

var FINAL_KEYS = {
  ACTIVE: FINAL_CONFIG.PROPERTY_PREFIX + 'active',
  TOTAL_WORD_FILES: FINAL_CONFIG.PROPERTY_PREFIX + 'totalWordFiles',
  TOTAL_NEEDING_CONVERSION:
    FINAL_CONFIG.PROPERTY_PREFIX + 'totalNeedingConversion',
  PROCESSED_NEEDING_CONVERSION:
    FINAL_CONFIG.PROPERTY_PREFIX + 'processedNeedingConversion',
  CONVERTED: FINAL_CONFIG.PROPERTY_PREFIX + 'converted',
  FAILED: FINAL_CONFIG.PROPERTY_PREFIX + 'failed',
  LAST_FILE: FINAL_CONFIG.PROPERTY_PREFIX + 'lastFile',
  LAST_FOLDER: FINAL_CONFIG.PROPERTY_PREFIX + 'lastFolder'
};


/*
 * START HERE. Run manually once.
 */
function startFinalWordConversion() {
  var lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    stopFinalWordConversion_();

    var rootFolder = DriveApp.getFolderById(
      FINAL_SOURCE_FOLDER_ID
    );

    var totalWordFiles = countFinalWordFiles_(rootFolder);

    var totalNeedingConversion =
      countFinalUnmatchedWordFiles_(rootFolder);

    var properties = PropertiesService.getScriptProperties();

    properties.setProperties({
      'finalWordConversion.v1.active': 'true',
      'finalWordConversion.v1.totalWordFiles':
        String(totalWordFiles),
      'finalWordConversion.v1.totalNeedingConversion':
        String(totalNeedingConversion),
      'finalWordConversion.v1.processedNeedingConversion': '0',
      'finalWordConversion.v1.converted': '0',
      'finalWordConversion.v1.failed': '0',
      'finalWordConversion.v1.lastFile': '',
      'finalWordConversion.v1.lastFolder': ''
    });

    prepareFinalReportSheet_();
    ensureFinalWordTrigger_();
    refreshFinalWordProgress_();
  } finally {
    lock.releaseLock();
  }

  processFinalWordConversionBatch();
}


/*
 * Runs automatically once per minute.
 * Do not normally run this manually.
 */
function processFinalWordConversionBatch() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    return;
  }

  try {
    var properties = PropertiesService.getScriptProperties();

    if (
      properties.getProperty(FINAL_KEYS.ACTIVE) !== 'true'
    ) {
      removeFinalWordTriggers_();
      return;
    }

    var rootFolder = DriveApp.getFolderById(
      FINAL_SOURCE_FOLDER_ID
    );

    var batch = {
      processed: 0,
      converted: 0,
      failed: 0,
      lastFile: '',
      lastFolder: '',
      reportRows: []
    };

    processFinalFolder_(
      rootFolder,
      rootFolder.getName(),
      batch
    );

    addFinalCounter_(
      FINAL_KEYS.PROCESSED_NEEDING_CONVERSION,
      batch.processed
    );

    addFinalCounter_(
      FINAL_KEYS.CONVERTED,
      batch.converted
    );

    addFinalCounter_(
      FINAL_KEYS.FAILED,
      batch.failed
    );

    if (batch.lastFile) {
      properties.setProperty(
        FINAL_KEYS.LAST_FILE,
        batch.lastFile
      );

      properties.setProperty(
        FINAL_KEYS.LAST_FOLDER,
        batch.lastFolder
      );
    }

    appendFinalReportRows_(batch.reportRows);

    var remaining = countFinalUnmatchedWordFiles_(
      rootFolder
    );

    if (remaining === 0) {
      properties.setProperty(
        FINAL_KEYS.ACTIVE,
        'false'
      );

      removeFinalWordTriggers_();
    }

    refreshFinalWordProgress_();
  } finally {
    lock.releaseLock();
  }
}


/*
 * Scans the folder tree.
 *
 * It does NOT count existing same-title Google Docs against the batch limit.
 * The batch limit is used only for files that need conversion or fail during
 * an attempted conversion.
 */
function processFinalFolder_(folder, folderPath, batch) {
  if (batch.processed >= FINAL_CONFIG.FILES_PER_RUN) {
    return;
  }

  var nativeDocNames = getFinalNativeDocNames_(folder);
  var files = folder.getFiles();

  while (
    files.hasNext() &&
    batch.processed < FINAL_CONFIG.FILES_PER_RUN
  ) {
    var wordFile = files.next();

    if (!isFinalWordFile_(wordFile)) {
      continue;
    }

    var newGoogleDocName = getFinalGoogleDocName_(
      wordFile.getName()
    );

    var normalizedName = newGoogleDocName.toLowerCase();

    /*
     * Existing native Docs are ignored without counting toward this batch.
     */
    if (nativeDocNames[normalizedName]) {
      continue;
    }

    batch.processed += 1;
    batch.lastFile = wordFile.getName();
    batch.lastFolder = folderPath;

    try {
      var created = Drive.Files.copy(
        {
          name: newGoogleDocName,
          mimeType: MimeType.GOOGLE_DOCS,
          parents: [folder.getId()]
        },
        wordFile.getId()
      );

      var newDocumentUrl = '';

      if (created && created.id) {
        newDocumentUrl =
          'https://docs.google.com/document/d/' +
          created.id +
          '/edit';
      }

      /*
       * Prevent another duplicate in this same execution.
       */
      nativeDocNames[normalizedName] = true;

      batch.converted += 1;

      batch.reportRows.push([
        new Date(),
        folderPath,
        wordFile.getName(),
        newGoogleDocName,
        'Converted',
        'Original Word file retained.',
        newDocumentUrl
      ]);
    } catch (error) {
      batch.failed += 1;

      batch.reportRows.push([
        new Date(),
        folderPath,
        wordFile.getName(),
        newGoogleDocName,
        'Failed',
        error && error.message
          ? error.message
          : String(error),
        ''
      ]);
    }
  }

  if (!FINAL_CONFIG.INCLUDE_SUBFOLDERS) {
    return;
  }

  var subfolders = folder.getFolders();

  while (
    subfolders.hasNext() &&
    batch.processed < FINAL_CONFIG.FILES_PER_RUN
  ) {
    var subfolder = subfolders.next();

    processFinalFolder_(
      subfolder,
      folderPath + ' / ' + subfolder.getName(),
      batch
    );
  }
}


/*
 * Counts every actual Word file in the source tree.
 */
function countFinalWordFiles_(folder) {
  var count = 0;
  var files = folder.getFiles();

  while (files.hasNext()) {
    if (isFinalWordFile_(files.next())) {
      count += 1;
    }
  }

  if (!FINAL_CONFIG.INCLUDE_SUBFOLDERS) {
    return count;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    count += countFinalWordFiles_(
      subfolders.next()
    );
  }

  return count;
}


/*
 * Counts Word files which have no native Google Doc with the same base title
 * in their own folder.
 */
function countFinalUnmatchedWordFiles_(folder) {
  var count = 0;
  var nativeDocNames = getFinalNativeDocNames_(folder);
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();

    if (!isFinalWordFile_(file)) {
      continue;
    }

    var newGoogleDocName = getFinalGoogleDocName_(
      file.getName()
    );

    if (!nativeDocNames[newGoogleDocName.toLowerCase()]) {
      count += 1;
    }
  }

  if (!FINAL_CONFIG.INCLUDE_SUBFOLDERS) {
    return count;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    count += countFinalUnmatchedWordFiles_(
      subfolders.next()
    );
  }

  return count;
}


/*
 * Creates a lowercase lookup of native Google Doc names in one folder.
 */
function getFinalNativeDocNames_(folder) {
  var names = {};
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();

    if (file.getMimeType() === MimeType.GOOGLE_DOCS) {
      names[file.getName().toLowerCase()] = true;
    }
  }

  return names;
}


function isFinalWordFile_(file) {
  var mimeType = file.getMimeType();
  var fileName = file.getName().toLowerCase();

  return (
    mimeType === 'application/msword' ||
    mimeType ===
      'application/vnd.openxmlformats-officedocument.wordprocessingml.document' ||
    /\.(doc|docx)$/i.test(fileName)
  );
}


/*
 * Example.docx becomes Example.
 */
function getFinalGoogleDocName_(fileName) {
  return String(fileName)
    .replace(/\.(doc|docx)$/i, '')
    .trim();
}


function addFinalCounter_(propertyKey, amount) {
  if (!amount) {
    return;
  }

  var properties = PropertiesService.getScriptProperties();

  var oldValue = Number(
    properties.getProperty(propertyKey) || 0
  );

  properties.setProperty(
    propertyKey,
    String(oldValue + amount)
  );
}


/*
 * Creates or clears the report tab for this conversion run.
 */
function prepareFinalReportSheet_() {
  var spreadsheet = SpreadsheetApp.openById(
    FINAL_SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    FINAL_CONFIG.REPORT_SHEET_NAME
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      FINAL_CONFIG.REPORT_SHEET_NAME
    );
  }

  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  sheet.clear();

  sheet.getRange('A1:G1').setValues([[
    'Timestamp',
    'Folder',
    'Original Word file',
    'Native Google Doc name',
    'Status',
    'Details',
    'Google Doc link'
  ]]);

  sheet.setFrozenRows(1);

  sheet.getRange('A1:G1')
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.getRange('A:G').setWrap(true);

  sheet.setColumnWidth(1, 170);
  sheet.setColumnWidth(2, 280);
  sheet.setColumnWidth(3, 300);
  sheet.setColumnWidth(4, 300);
  sheet.setColumnWidth(5, 150);
  sheet.setColumnWidth(6, 420);
  sheet.setColumnWidth(7, 300);
}


/*
 * Logs this batch directly to the Sheet.
 * No large queue or results object is stored in Script Properties.
 */
function appendFinalReportRows_(rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  var spreadsheet = SpreadsheetApp.openById(
    FINAL_SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    FINAL_CONFIG.REPORT_SHEET_NAME
  );

  if (!sheet) {
    prepareFinalReportSheet_();

    sheet = spreadsheet.getSheetByName(
      FINAL_CONFIG.REPORT_SHEET_NAME
    );
  }

  var startRow = sheet.getLastRow() + 1;

  sheet
    .getRange(startRow, 1, rows.length, 7)
    .setValues(rows);

  var formulas = rows.map(function(row) {
    if (!row[6]) {
      return [''];
    }

    return [
      '=HYPERLINK("' +
      escapeFinalFormula_(row[6]) +
      '","Open Google Doc")'
    ];
  });

  sheet
    .getRange(startRow, 7, formulas.length, 1)
    .setFormulas(formulas);
}


/*
 * Creates only one automatic one-minute trigger.
 */
function ensureFinalWordTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i += 1) {
    if (
      triggers[i].getHandlerFunction() ===
      'processFinalWordConversionBatch'
    ) {
      return;
    }
  }

  ScriptApp.newTrigger(
    'processFinalWordConversionBatch'
  )
    .timeBased()
    .everyMinutes(1)
    .create();
}


function removeFinalWordTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i += 1) {
    if (
      triggers[i].getHandlerFunction() ===
      'processFinalWordConversionBatch'
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/*
 * Stops only this final converter. It does not delete source Word files or
 * converted Google Docs.
 */
function stopFinalWordConversion() {
  stopFinalWordConversion_();
  refreshFinalWordProgress_();
}


function stopFinalWordConversion_() {
  removeFinalWordTriggers_();

  var properties = PropertiesService.getScriptProperties();

  Object.keys(FINAL_KEYS).forEach(function(key) {
    properties.deleteProperty(FINAL_KEYS[key]);
  });
}


/*
 * Refreshes the Progress sheet without converting files.
 */
function refreshFinalWordProgress() {
  refreshFinalWordProgress_();
}


function refreshFinalWordProgress_() {
  var spreadsheet = SpreadsheetApp.openById(
    FINAL_SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    FINAL_CONFIG.PROGRESS_SHEET_NAME
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      FINAL_CONFIG.PROGRESS_SHEET_NAME
    );
  }

  var properties = PropertiesService.getScriptProperties();

  var active = properties.getProperty(
    FINAL_KEYS.ACTIVE
  ) === 'true';

  var totalWordFiles = Number(
    properties.getProperty(
      FINAL_KEYS.TOTAL_WORD_FILES
    ) || 0
  );

  var totalNeedingConversion = Number(
    properties.getProperty(
      FINAL_KEYS.TOTAL_NEEDING_CONVERSION
    ) || 0
  );

  var processed = Number(
    properties.getProperty(
      FINAL_KEYS.PROCESSED_NEEDING_CONVERSION
    ) || 0
  );

  var converted = Number(
    properties.getProperty(FINAL_KEYS.CONVERTED) || 0
  );

  var failed = Number(
    properties.getProperty(FINAL_KEYS.FAILED) || 0
  );

  var lastFile = properties.getProperty(
    FINAL_KEYS.LAST_FILE
  ) || '';

  var lastFolder = properties.getProperty(
    FINAL_KEYS.LAST_FOLDER
  ) || '';

  var remaining = '';

  try {
    remaining = countFinalUnmatchedWordFiles_(
      DriveApp.getFolderById(
        FINAL_SOURCE_FOLDER_ID
      )
    );
  } catch (error) {
    remaining =
      'Could not count remaining files: ' +
      (error.message || String(error));
  }

  var progress = totalNeedingConversion > 0
    ? processed / totalNeedingConversion
    : 0;

  sheet.clear();

  sheet.getRange('A1:B1').setValues([[
    'Word Conversion Progress',
    ''
  ]]);

  sheet.getRange('A3:B14').setValues([
    ['Job active', active ? 'Yes' : 'No'],
    ['Source folder ID', FINAL_SOURCE_FOLDER_ID],
    ['Total Word files found', totalWordFiles],
    ['Word files needing conversion at start',
      totalNeedingConversion],
    ['Conversion attempts processed', processed],
    ['Converted to native Google Docs', converted],
    ['Failed', failed],
    ['Still needing conversion', remaining],
    ['Last file attempted', lastFile],
    ['Last folder attempted', lastFolder],
    ['Progress', progress],
    ['Note',
      'Original Word files remain by design.']
  ]);

  sheet.getRange('B13').setNumberFormat('0.0%');

  sheet.getRange('A1:B1')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#cfe2f3');

  sheet.getRange('A3:A14')
    .setFontWeight('bold')
    .setBackground('#eaf2f8');

  sheet.getRange('A:B').setWrap(true);

  sheet.setColumnWidth(1, 320);
  sheet.setColumnWidth(2, 650);
}


function escapeFinalFormula_(text) {
  return String(text).replace(/"/g, '""');
}