/*
 * FIXED WORD TO GOOGLE DOCS CONVERTER
 *
 * Target: one selected Drive subfolder and all of its subfolders.
 *
 * What it does:
 * - Finds actual .doc and .docx files.
 * - Checks whether a native Google Doc with the same base name already exists
 *   in the SAME folder.
 * - Converts only Word files without such a native Google Doc.
 * - Keeps the original Word file.
 * - Processes up to 25 Word files per run.
 * - Continues automatically every minute until no unmatched Word files remain.
 * - Writes results directly to a spreadsheet report tab.
 *
 * Before running:
 * 1. Apps Script -> Services -> + -> Drive API -> Add.
 * 2. Save.
 * 3. Run startFixedWordConversion once.
 */

var FIXED_SOURCE_FOLDER_ID =
  '1b6Uqq3W17iKuAVs9SPoCooBF444RnCyY';

var FIXED_SPREADSHEET_ID =
  '1lfqBwaXUZ62xPimJRnKgAUSebhhEmbGPFvzPMuIlCa4';

var FIXED_CONFIG = {
  INCLUDE_SUBFOLDERS: true,
  FILES_PER_RUN: 25,
  PROGRESS_SHEET_NAME: 'Word Conversion Progress',
  REPORT_SHEET_NAME: 'Word Conversion Report',
  PROPERTY_PREFIX: 'fixedWordConversion.v1.'
};

var FIXED_KEYS = {
  ACTIVE: FIXED_CONFIG.PROPERTY_PREFIX + 'active',
  TOTAL_AT_START: FIXED_CONFIG.PROPERTY_PREFIX + 'totalAtStart',
  CHECKED: FIXED_CONFIG.PROPERTY_PREFIX + 'checked',
  CONVERTED: FIXED_CONFIG.PROPERTY_PREFIX + 'converted',
  ALREADY_NATIVE: FIXED_CONFIG.PROPERTY_PREFIX + 'alreadyNative',
  FAILED: FIXED_CONFIG.PROPERTY_PREFIX + 'failed',
  LAST_FILE: FIXED_CONFIG.PROPERTY_PREFIX + 'lastFile',
  LAST_FOLDER: FIXED_CONFIG.PROPERTY_PREFIX + 'lastFolder'
};


/*
 * START HERE.
 * Run this function manually once.
 */
function startFixedWordConversion() {
  stopFixedWordConversion_();

  var rootFolder = DriveApp.getFolderById(
    FIXED_SOURCE_FOLDER_ID
  );

  var total = countFixedWordFiles_(rootFolder);

  var properties = PropertiesService.getScriptProperties();

  properties.setProperties({
    'fixedWordConversion.v1.active': 'true',
    'fixedWordConversion.v1.totalAtStart': String(total),
    'fixedWordConversion.v1.checked': '0',
    'fixedWordConversion.v1.converted': '0',
    'fixedWordConversion.v1.alreadyNative': '0',
    'fixedWordConversion.v1.failed': '0',
    'fixedWordConversion.v1.lastFile': '',
    'fixedWordConversion.v1.lastFolder': ''
  });

  prepareFixedReportSheet_();
  ensureFixedWordTrigger_();
  refreshFixedWordProgress_();

  processFixedWordConversionBatch();
}


/*
 * Runs automatically every minute.
 * Do not normally run this manually.
 */
function processFixedWordConversionBatch() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(1000)) {
    return;
  }

  try {
    var properties = PropertiesService.getScriptProperties();

    if (
      properties.getProperty(FIXED_KEYS.ACTIVE) !== 'true'
    ) {
      removeFixedWordTriggers_();
      return;
    }

    var rootFolder = DriveApp.getFolderById(
      FIXED_SOURCE_FOLDER_ID
    );

    var batch = {
      checked: 0,
      converted: 0,
      alreadyNative: 0,
      failed: 0,
      lastFile: '',
      lastFolder: '',
      reportRows: []
    };

    processFixedFolder_(
      rootFolder,
      rootFolder.getName(),
      batch
    );

    addFixedCounter_(FIXED_KEYS.CHECKED, batch.checked);
    addFixedCounter_(FIXED_KEYS.CONVERTED, batch.converted);
    addFixedCounter_(
      FIXED_KEYS.ALREADY_NATIVE,
      batch.alreadyNative
    );
    addFixedCounter_(FIXED_KEYS.FAILED, batch.failed);

    if (batch.lastFile) {
      properties.setProperty(
        FIXED_KEYS.LAST_FILE,
        batch.lastFile
      );

      properties.setProperty(
        FIXED_KEYS.LAST_FOLDER,
        batch.lastFolder
      );
    }

    appendFixedReportRows_(batch.reportRows);

    var remaining = countFixedUnmatchedWordFiles_(rootFolder);

    if (remaining === 0) {
      properties.setProperty(FIXED_KEYS.ACTIVE, 'false');
      removeFixedWordTriggers_();
    }

    refreshFixedWordProgress_();
  } finally {
    lock.releaseLock();
  }
}


/*
 * Scans folders in order and handles at most FILES_PER_RUN Word files
 * during this one execution.
 */
function processFixedFolder_(folder, folderPath, batch) {
  if (batch.checked >= FIXED_CONFIG.FILES_PER_RUN) {
    return;
  }

  var nativeDocNames = getFixedNativeDocNames_(folder);
  var files = folder.getFiles();

  while (
    files.hasNext() &&
    batch.checked < FIXED_CONFIG.FILES_PER_RUN
  ) {
    var wordFile = files.next();

    if (!isFixedWordFile_(wordFile)) {
      continue;
    }

    batch.checked += 1;
    batch.lastFile = wordFile.getName();
    batch.lastFolder = folderPath;

    var newGoogleDocName = getFixedGoogleDocName_(
      wordFile.getName()
    );

    var normalizedName = newGoogleDocName.toLowerCase();

    if (nativeDocNames[normalizedName]) {
      batch.alreadyNative += 1;

      batch.reportRows.push([
        new Date(),
        folderPath,
        wordFile.getName(),
        newGoogleDocName,
        'Already native',
        'A same-named native Google Doc already exists in this folder.',
        ''
      ]);

      continue;
    }

    try {
      var created = Drive.Files.copy(
        {
          name: newGoogleDocName,
          mimeType: MimeType.GOOGLE_DOCS,
          parents: [folder.getId()]
        },
        wordFile.getId()
      );

      var url = created && created.id
        ? 'https://docs.google.com/document/d/' +
          created.id +
          '/edit'
        : '';

      nativeDocNames[normalizedName] = true;
      batch.converted += 1;

      batch.reportRows.push([
        new Date(),
        folderPath,
        wordFile.getName(),
        newGoogleDocName,
        'Converted',
        'Original Word file retained.',
        url
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

  if (!FIXED_CONFIG.INCLUDE_SUBFOLDERS) {
    return;
  }

  var subfolders = folder.getFolders();

  while (
    subfolders.hasNext() &&
    batch.checked < FIXED_CONFIG.FILES_PER_RUN
  ) {
    var subfolder = subfolders.next();

    processFixedFolder_(
      subfolder,
      folderPath + ' / ' + subfolder.getName(),
      batch
    );
  }
}


/*
 * Counts all actual Word files in the folder tree.
 */
function countFixedWordFiles_(folder) {
  var count = 0;
  var files = folder.getFiles();

  while (files.hasNext()) {
    if (isFixedWordFile_(files.next())) {
      count += 1;
    }
  }

  if (!FIXED_CONFIG.INCLUDE_SUBFOLDERS) {
    return count;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    count += countFixedWordFiles_(
      subfolders.next()
    );
  }

  return count;
}


/*
 * Counts Word files which do not yet have a same-title native Google Doc in
 * their own folder. This is the useful "remaining" count.
 */
function countFixedUnmatchedWordFiles_(folder) {
  var count = 0;
  var nativeDocNames = getFixedNativeDocNames_(folder);
  var files = folder.getFiles();

  while (files.hasNext()) {
    var file = files.next();

    if (!isFixedWordFile_(file)) {
      continue;
    }

    var newGoogleDocName = getFixedGoogleDocName_(
      file.getName()
    );

    if (!nativeDocNames[newGoogleDocName.toLowerCase()]) {
      count += 1;
    }
  }

  if (!FIXED_CONFIG.INCLUDE_SUBFOLDERS) {
    return count;
  }

  var subfolders = folder.getFolders();

  while (subfolders.hasNext()) {
    count += countFixedUnmatchedWordFiles_(
      subfolders.next()
    );
  }

  return count;
}


/*
 * Makes a lowercase lookup table of native Google Docs in one folder.
 */
function getFixedNativeDocNames_(folder) {
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


function isFixedWordFile_(file) {
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
 * Converts "Example.docx" into "Example".
 */
function getFixedGoogleDocName_(fileName) {
  return String(fileName)
    .replace(/\.(doc|docx)$/i, '')
    .trim();
}


function addFixedCounter_(propertyKey, amount) {
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
 * Creates the report tab and writes its headings.
 */
function prepareFixedReportSheet_() {
  var spreadsheet = SpreadsheetApp.openById(
    FIXED_SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    FIXED_CONFIG.REPORT_SHEET_NAME
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      FIXED_CONFIG.REPORT_SHEET_NAME
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
 * Appends batch results directly to the spreadsheet. This avoids the
 * Script Properties size limit.
 */
function appendFixedReportRows_(rows) {
  if (!rows || rows.length === 0) {
    return;
  }

  var spreadsheet = SpreadsheetApp.openById(
    FIXED_SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    FIXED_CONFIG.REPORT_SHEET_NAME
  );

  if (!sheet) {
    prepareFixedReportSheet_();

    sheet = spreadsheet.getSheetByName(
      FIXED_CONFIG.REPORT_SHEET_NAME
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
      escapeFixedFormula_(row[6]) +
      '","Open Google Doc")'
    ];
  });

  sheet
    .getRange(startRow, 7, formulas.length, 1)
    .setFormulas(formulas);
}


/*
 * Ensures exactly one time-driven trigger exists for this script.
 */
function ensureFixedWordTrigger_() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i += 1) {
    if (
      triggers[i].getHandlerFunction() ===
      'processFixedWordConversionBatch'
    ) {
      return;
    }
  }

  ScriptApp.newTrigger(
    'processFixedWordConversionBatch'
  )
    .timeBased()
    .everyMinutes(1)
    .create();
}


function removeFixedWordTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i += 1) {
    if (
      triggers[i].getHandlerFunction() ===
      'processFixedWordConversionBatch'
    ) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/*
 * Stops this converter and clears only this converter's saved counters.
 * It never deletes source Word files or converted Docs.
 */
function stopFixedWordConversion() {
  stopFixedWordConversion_();
  refreshFixedWordProgress_();
}


function stopFixedWordConversion_() {
  removeFixedWordTriggers_();

  var properties = PropertiesService.getScriptProperties();

  Object.keys(FIXED_KEYS).forEach(function(key) {
    properties.deleteProperty(FIXED_KEYS[key]);
  });
}


/*
 * Refreshes the progress tab.
 */
function refreshFixedWordProgress() {
  refreshFixedWordProgress_();
}


function refreshFixedWordProgress_() {
  var spreadsheet = SpreadsheetApp.openById(
    FIXED_SPREADSHEET_ID
  );

  var sheet = spreadsheet.getSheetByName(
    FIXED_CONFIG.PROGRESS_SHEET_NAME
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      FIXED_CONFIG.PROGRESS_SHEET_NAME
    );
  }

  var properties = PropertiesService.getScriptProperties();

  var active = properties.getProperty(
    FIXED_KEYS.ACTIVE
  ) === 'true';

  var total = Number(
    properties.getProperty(
      FIXED_KEYS.TOTAL_AT_START
    ) || 0
  );

  var checked = Number(
    properties.getProperty(FIXED_KEYS.CHECKED) || 0
  );

  var converted = Number(
    properties.getProperty(FIXED_KEYS.CONVERTED) || 0
  );

  var alreadyNative = Number(
    properties.getProperty(
      FIXED_KEYS.ALREADY_NATIVE
    ) || 0
  );

  var failed = Number(
    properties.getProperty(FIXED_KEYS.FAILED) || 0
  );

  var lastFile = properties.getProperty(
    FIXED_KEYS.LAST_FILE
  ) || '';

  var lastFolder = properties.getProperty(
    FIXED_KEYS.LAST_FOLDER
  ) || '';

  var remaining = '';

  try {
    remaining = countFixedUnmatchedWordFiles_(
      DriveApp.getFolderById(
        FIXED_SOURCE_FOLDER_ID
      )
    );
  } catch (error) {
    remaining =
      'Could not count remaining files: ' +
      (error.message || String(error));
  }

  sheet.clear();

  sheet.getRange('A1:B1').setValues([[
    'Word Conversion Progress',
    ''
  ]]);

  sheet.getRange('A3:B13').setValues([
    ['Job active', active ? 'Yes' : 'No'],
    ['Source folder ID', FIXED_SOURCE_FOLDER_ID],
    ['Total Word files at start', total],
    ['Checked this run', checked],
    ['Converted', converted],
    ['Already had native Google Doc', alreadyNative],
    ['Failed', failed],
    ['Still needing conversion', remaining],
    ['Last file attempted', lastFile],
    ['Last folder attempted', lastFolder],
    ['Progress', total > 0 ? checked / total : 0]
  ]);

  sheet.getRange('B13').setNumberFormat('0.0%');

  sheet.getRange('A1:B1')
    .setFontWeight('bold')
    .setFontSize(14)
    .setBackground('#cfe2f3');

  sheet.getRange('A3:A13')
    .setFontWeight('bold')
    .setBackground('#eaf2f8');

  sheet.getRange('A:B').setWrap(true);

  sheet.setColumnWidth(1, 280);
  sheet.setColumnWidth(2, 600);
}


function escapeFixedFormula_(text) {
  return String(text).replace(/"/g, '""');
}