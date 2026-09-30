/**
 * SMART-CHIP-ONLY DATABASE COMPARISON
 *
 * Finds source rows that:
 * 1. Are missing from the scripted database tab, AND
 * 2. Contain a Google Sheets smart chip in "Play / Mono Link".
 *
 * Plain text, normal rich-text hyperlinks, and =HYPERLINK() formula cells
 * are intentionally excluded.
 *
 * REQUIRED SETUP:
 * Apps Script > Services > + > Google Sheets API > Add
 */

/* =========================
   Configuration
   ========================= */

const COMPARE_CONFIG = {
  /*
   * Leave blank if this Apps Script project is bound to the spreadsheet:
   * Google Sheet > Extensions > Apps Script.
   *
   * For a standalone script, enter the spreadsheet ID here—not a gid.
   */
  spreadsheetId: '',

  /*
   * Use either:
   * - exact visible tab names, OR
   * - numeric gids as strings.
   */
  sourceTab: 'Copy of Female Serio-Comic (Contemp)',
  scriptedTab: 'Female Serio-Comic (Contemporary)',

  reviewTab: 'Rows Missing — Smart Chips Only',

  matchHeaders: [
    'Playwright',
    'Play / Mono Link',
    'Character',
    'First Line'
  ],

  smartChipHeader: 'Play / Mono Link'
};


/* =========================
   Main function
   ========================= */

/**
 * Run this function.
 */
function auditSmartChipRows() {
  const spreadsheet = openCompareSpreadsheet_();

  const sourceSheet = findCompareSheet_(
    spreadsheet,
    COMPARE_CONFIG.sourceTab
  );

  const scriptedSheet = findCompareSheet_(
    spreadsheet,
    COMPARE_CONFIG.scriptedTab
  );

  const source = readCompareTable_(sourceSheet);
  const scripted = readCompareTable_(scriptedSheet);

  const sourceColumns = mapCompareHeaders_(
    source.headers
  );

  const scriptedColumns = mapCompareHeaders_(
    scripted.headers
  );

  validateCompareColumns_(
    sourceColumns,
    sourceSheet.getName()
  );

  validateCompareColumns_(
    scriptedColumns,
    scriptedSheet.getName()
  );

  const scriptedKeys = new Set(
    scripted.rows
      .map(function(row) {
        return makeCompareKey_(
          row,
          scriptedColumns
        );
      })
      .filter(function(key) {
        return key !== '';
      })
  );

  const smartChipColumn =
    sourceColumns[
      normalizeCompareText_(
        COMPARE_CONFIG.smartChipHeader
      )
    ] + 1;

  const rowsWithSmartChips = findCompareSmartChipRows_(
    spreadsheet.getId(),
    sourceSheet,
    smartChipColumn
  );

  const missingSmartChipRows = source.rows
    .map(function(row, index) {
      return {
        rowNumber: index + 2,
        key: makeCompareKey_(
          row,
          sourceColumns
        )
      };
    })
    .filter(function(item) {
      return (
        item.key !== '' &&
        !scriptedKeys.has(item.key) &&
        rowsWithSmartChips.has(item.rowNumber)
      );
    });

  writeCompareReview_(
    spreadsheet,
    sourceSheet,
    source.headers,
    missingSmartChipRows
  );

  SpreadsheetApp.getUi().alert(
    'Smart-chip comparison complete.\n\n' +
    'Missing rows with a smart chip in "' +
    COMPARE_CONFIG.smartChipHeader +
    '": ' +
    missingSmartChipRows.length +
    '\n\nReview tab:\n' +
    COMPARE_CONFIG.reviewTab
  );
}


/* =========================
   Read and compare
   ========================= */

function readCompareTable_(sheet) {
  const rowCount = sheet.getLastRow();
  const columnCount = sheet.getLastColumn();

  if (rowCount === 0 || columnCount === 0) {
    throw new Error(
      'The tab "' + sheet.getName() + '" is empty.'
    );
  }

  const values = sheet
    .getRange(1, 1, rowCount, columnCount)
    .getDisplayValues();

  return {
    headers: values[0],
    rows: values.slice(1)
  };
}


function mapCompareHeaders_(headers) {
  return headers.reduce(function(columns, header, index) {
    const normalizedHeader = normalizeCompareText_(
      header
    );

    if (normalizedHeader !== '') {
      columns[normalizedHeader] = index;
    }

    return columns;
  }, {});
}


function validateCompareColumns_(columns, sheetName) {
  const missing = COMPARE_CONFIG.matchHeaders.filter(
    function(header) {
      return columns[
        normalizeCompareText_(header)
      ] === undefined;
    }
  );

  if (missing.length > 0) {
    throw new Error(
      'The tab "' +
      sheetName +
      '" is missing required column(s): ' +
      missing.join(', ')
    );
  }
}


function makeCompareKey_(row, columns) {
  const values = COMPARE_CONFIG.matchHeaders.map(
    function(header) {
      const columnIndex = columns[
        normalizeCompareText_(header)
      ];

      return normalizeCompareText_(
        row[columnIndex]
      );
    }
  );

  const allBlank = values.every(function(value) {
    return value === '';
  });

  return allBlank ? '' : values.join('|');
}


/* =========================
   Smart-chip detection
   ========================= */

/**
 * Returns a Set of 1-based sheet row numbers that contain a smart chip in
 * the selected column.
 *
 * The Google Sheets API returns smart chips through:
 * sheets.data.rowData.values.chipRuns
 */
function findCompareSmartChipRows_(
  spreadsheetId,
  sheet,
  columnNumber
) {
  const columnLetter = columnNumberToLetter_(
    columnNumber
  );

  const a1Range =
    quoteCompareSheetName_(sheet.getName()) +
    '!' +
    columnLetter +
    '2:' +
    columnLetter;

  let response;

  try {
    response = Sheets.Spreadsheets.get(
      spreadsheetId,
      {
        ranges: [a1Range],
        includeGridData: true,
        fields:
          'sheets(data(startRow,rowData(values(chipRuns))))'
      }
    );
  } catch (error) {
    throw new Error(
      'Could not read smart-chip metadata. Confirm that the ' +
      'Google Sheets API is enabled under Apps Script > Services. ' +
      'Google returned: ' +
      error.message
    );
  }

  const smartChipRows = new Set();

  (response.sheets || []).forEach(function(apiSheet) {
    (apiSheet.data || []).forEach(function(gridData) {
      const startRow = Number(gridData.startRow || 0);

      (gridData.rowData || []).forEach(
        function(rowData, rowOffset) {
          const hasSmartChip = (rowData.values || []).some(
            function(cell) {
              return (cell.chipRuns || []).some(
                function(chipRun) {
                  /*
                   * Any chip run counts. This includes rich-link chips,
                   * file chips, people chips, place chips, and other chip
                   * formats returned by the Sheets API.
                   */
                  return Boolean(chipRun.chip);
                }
              );
            }
          );

          if (hasSmartChip) {
            /*
             * API rows are zero-based; Google Sheets row numbers are 1-based.
             */
            smartChipRows.add(
              startRow + rowOffset + 1
            );
          }
        }
      );
    });
  });

  return smartChipRows;
}


/* =========================
   Review output
   ========================= */

/**
 * Rebuilds the review tab.
 *
 * Direct copying preserves the original source cell content and formatting.
 */
function writeCompareReview_(
  spreadsheet,
  sourceSheet,
  sourceHeaders,
  rowsToCopy
) {
  const reviewSheet = getOrCreateCompareReviewSheet_(
    spreadsheet
  );

  if (reviewSheet.getFilter()) {
    reviewSheet.getFilter().remove();
  }

  reviewSheet.clear();

  const reviewHeaders = [
    'Source Row',
    'Comparison Key'
  ].concat(sourceHeaders);

  reviewSheet
    .getRange(1, 1, 1, reviewHeaders.length)
    .setValues([reviewHeaders]);

  rowsToCopy.forEach(function(item, index) {
    const destinationRow = index + 2;

    reviewSheet
      .getRange(destinationRow, 1)
      .setValue(item.rowNumber);

    reviewSheet
      .getRange(destinationRow, 2)
      .setValue(item.key);

    sourceSheet
      .getRange(
        item.rowNumber,
        1,
        1,
        sourceHeaders.length
      )
      .copyTo(
        reviewSheet.getRange(
          destinationRow,
          3,
          1,
          sourceHeaders.length
        ),
        SpreadsheetApp.CopyPasteType.PASTE_NORMAL,
        false
      );
  });

  formatCompareReview_(
    reviewSheet,
    reviewHeaders.length,
    rowsToCopy.length
  );
}


function getOrCreateCompareReviewSheet_(spreadsheet) {
  let sheet = spreadsheet.getSheetByName(
    COMPARE_CONFIG.reviewTab
  );

  if (!sheet) {
    sheet = spreadsheet.insertSheet(
      COMPARE_CONFIG.reviewTab
    );
  }

  return sheet;
}


function formatCompareReview_(
  sheet,
  columnCount,
  resultCount
) {
  sheet.setFrozenRows(1);

  sheet
    .getRange(1, 1, 1, columnCount)
    .setFontWeight('bold')
    .setBackground('#d9ead3');

  sheet.getRange('A:ZZ').setWrap(true);

  sheet.setColumnWidth(1, 100);
  sheet.setColumnWidth(2, 430);

  for (
    let column = 3;
    column <= columnCount;
    column += 1
  ) {
    sheet.setColumnWidth(column, 230);
  }

  if (resultCount > 0) {
    sheet
      .getRange(
        1,
        1,
        resultCount + 1,
        columnCount
      )
      .createFilter();
  }
}


/* =========================
   Spreadsheet utilities
   ========================= */

function openCompareSpreadsheet_() {
  const spreadsheetId = String(
    COMPARE_CONFIG.spreadsheetId || ''
  ).trim();

  if (spreadsheetId !== '') {
    return SpreadsheetApp.openById(
      spreadsheetId
    );
  }

  const spreadsheet =
    SpreadsheetApp.getActiveSpreadsheet();

  if (!spreadsheet) {
    throw new Error(
      'No active spreadsheet was found. Open Apps Script from ' +
      'your Google Sheet or set spreadsheetId.'
    );
  }

  return spreadsheet;
}


/**
 * Finds a worksheet by exact visible tab name or numeric gid.
 */
function findCompareSheet_(
  spreadsheet,
  reference
) {
  const tabReference = String(
    reference || ''
  ).trim();

  const byName = spreadsheet.getSheetByName(
    tabReference
  );

  if (byName) {
    return byName;
  }

  if (/^\d+$/.test(tabReference)) {
    const gid = Number(tabReference);

    const byGid = spreadsheet.getSheets().find(
      function(sheet) {
        return sheet.getSheetId() === gid;
      }
    );

    if (byGid) {
      return byGid;
    }
  }

  throw new Error(
    'Could not find tab "' +
    tabReference +
    '". Use the exact visible tab name or a valid numeric gid.'
  );
}


function normalizeCompareText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/[\u2018\u2019]/g, "'")
    .replace(/[\u201C\u201D]/g, '"')
    .replace(/[\u2013\u2014]/g, '-')
    .replace(/[^a-z0-9]+/gi, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();
}


function columnNumberToLetter_(columnNumber) {
  let number = columnNumber;
  let letter = '';

  while (number > 0) {
    const remainder = (number - 1) % 26;

    letter =
      String.fromCharCode(65 + remainder) +
      letter;

    number = Math.floor(
      (number - 1) / 26
    );
  }

  return letter;
}


function quoteCompareSheetName_(sheetName) {
  return "'" +
    String(sheetName).replace(/'/g, "''") +
    "'";
}