function syncEFByACDFirstThreeWords() {
  // ============================================================
  // SETTINGS — EDIT ONLY THIS SECTION
  // ============================================================

  // Spreadsheet IDs only: do not include /edit, ?gid, or #gid.
  const SOURCE_SPREADSHEET_ID = '1E9DW0-6cskGXchQ3hesWjfneODa4cLj9wLjJOejdkUQ';
  const TARGET_SPREADSHEET_ID = '1lfqBwaXUZ62xPimJRnKgAUSebhhEmbGPFvzPMuIlCa4';

  // Exact names of the bottom tabs in each spreadsheet.
  // Enter each name once.
  const SOURCE_TAB_NAME = 'Female Dramatic (Contemp)';
  const TARGET_TAB_NAME = 'Female Dramatic (Contemporary)';

  // true = preview only; target cells are NOT changed.
  // false = source E:F is copied into matching target E:F.
  const DRY_RUN = false;

  // Use 1 if row 1 contains column headings and row 2 begins data.
  const HEADER_ROWS = 1;

  // ============================================================
  // CHECK SETTINGS
  // ============================================================

  if (
    SOURCE_TAB_NAME === 'PUT YOUR SOURCE TAB NAME HERE' ||
    TARGET_TAB_NAME === 'PUT YOUR TARGET TAB NAME HERE'
  ) {
    throw new Error(
      'Set SOURCE_TAB_NAME and TARGET_TAB_NAME to the exact names shown on the bottom sheet tabs.'
    );
  }

  // ============================================================
  // OPEN SOURCE AND TARGET TABS
  // ============================================================

  const sourceSpreadsheet = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
  const targetSpreadsheet = SpreadsheetApp.openById(TARGET_SPREADSHEET_ID);

  const sourceSheet = sourceSpreadsheet.getSheetByName(SOURCE_TAB_NAME);
  const targetSheet = targetSpreadsheet.getSheetByName(TARGET_TAB_NAME);

  if (!sourceSheet) {
    throw new Error(
      `Source tab not found: "${SOURCE_TAB_NAME}". ` +
      'Check spelling, punctuation, and spaces.'
    );
  }

  if (!targetSheet) {
    throw new Error(
      `Target tab not found: "${TARGET_TAB_NAME}". ` +
      'Check spelling, punctuation, and spaces.'
    );
  }

  Logger.log(
    `Source: "${sourceSpreadsheet.getName()}" → "${sourceSheet.getName()}"`
  );
  Logger.log(
    `Target: "${targetSpreadsheet.getName()}" → "${targetSheet.getName()}"`
  );

  const sourceRowCount = sourceSheet.getLastRow() - HEADER_ROWS;
  const targetRowCount = targetSheet.getLastRow() - HEADER_ROWS;

  if (sourceRowCount < 1) {
    throw new Error(
      `Source tab "${sourceSheet.getName()}" has no data below the header row.`
    );
  }

  if (targetRowCount < 1) {
    throw new Error(
      `Target tab "${targetSheet.getName()}" has no data below the header row.`
    );
  }

  // ============================================================
  // READ A:F FROM BOTH TABS
  //
  // Match:
  //   A = exact after cleanup
  //   C = exact after cleanup
  //   D = first 3 words after cleanup
  //
  // Ignore:
  //   B
  //
  // Copy:
  //   E:F from source → target
  // ============================================================

  const sourceData = sourceSheet
    .getRange(HEADER_ROWS + 1, 1, sourceRowCount, 6)
    .getValues();

  const targetData = targetSheet
    .getRange(HEADER_ROWS + 1, 1, targetRowCount, 6)
    .getValues();

  // ============================================================
  // MATCHING HELPERS
  // ============================================================

  function normalize(value) {
    return String(value == null ? '' : value)
      .trim()
      .replace(/\s+/g, ' ')
      .toLowerCase();
  }

  function firstThreeWords(value) {
    return normalize(value)
      .replace(/[^a-z0-9\s]/g, ' ')
      .replace(/\s+/g, ' ')
      .trim()
      .split(' ')
      .filter(Boolean)
      .slice(0, 3)
      .join(' ');
  }

  function isBlankMatchRow(row) {
    return (
      normalize(row[0]) === '' &&
      normalize(row[2]) === '' &&
      normalize(row[3]) === ''
    );
  }

  function makeMatchKey(row) {
    return [
      normalize(row[0]),      // A
      normalize(row[2]),      // C
      firstThreeWords(row[3]) // D: first three words only
    ].join(' || ');
  }

  function nonBlankEFCount(valuesEF) {
    return valuesEF.filter(value => normalize(value) !== '').length;
  }

  function efValuesDiffer(valuesA, valuesB) {
    return (
      normalize(valuesA[0]) !== normalize(valuesB[0]) ||
      normalize(valuesA[1]) !== normalize(valuesB[1])
    );
  }

  // ============================================================
  // INDEX SOURCE RECORDS
  //
  // For duplicate match keys:
  // 1. Prefer record with more nonblank E:F values.
  // 2. If equal completeness, keep the earlier source record.
  // 3. Log equally complete, conflicting E:F data for review.
  // ============================================================

  const sourceByKey = new Map();
  const duplicateSourceKeys = [];
  const equalCompletenessConflicts = [];

  sourceData.forEach((row, index) => {
    if (isBlankMatchRow(row)) {
      return;
    }

    const key = makeMatchKey(row);
    const sourceRowNumber = HEADER_ROWS + index + 1;
    const candidateEF = [row[4], row[5]];
    const candidateCompleteness = nonBlankEFCount(candidateEF);

    if (!sourceByKey.has(key)) {
      sourceByKey.set(key, {
        sourceRowNumber: sourceRowNumber,
        valuesEF: candidateEF,
        completeness: candidateCompleteness
      });
      return;
    }

    const existing = sourceByKey.get(key);

    duplicateSourceKeys.push({
      firstSeenRow: existing.sourceRowNumber,
      duplicateRow: sourceRowNumber
    });

    // Use this row only if it has more actual E:F information.
    if (candidateCompleteness > existing.completeness) {
      sourceByKey.set(key, {
        sourceRowNumber: sourceRowNumber,
        valuesEF: candidateEF,
        completeness: candidateCompleteness
      });
      return;
    }

    // If equally complete but values differ, retain the earlier row
    // and put the conflict in the execution log.
    if (
      candidateCompleteness === existing.completeness &&
      efValuesDiffer(candidateEF, existing.valuesEF)
    ) {
      equalCompletenessConflicts.push({
        keptRow: existing.sourceRowNumber,
        ignoredRow: sourceRowNumber,
        keptEF: existing.valuesEF,
        ignoredEF: candidateEF
      });
    }
  });

  if (duplicateSourceKeys.length > 0) {
    const examples = duplicateSourceKeys
      .slice(0, 20)
      .map(item => `source rows ${item.firstSeenRow} and ${item.duplicateRow}`)
      .join('; ');

    Logger.log(
      'Duplicate source keys found. For each duplicate set, the script selected ' +
      'the row with the most nonblank E:F cells; tied rows keep the earlier row. ' +
      `Examples: ${examples}`
    );
  }

  if (equalCompletenessConflicts.length > 0) {
    const examples = equalCompletenessConflicts
      .slice(0, 20)
      .map(item =>
        `kept row ${item.keptRow} [${item.keptEF.join(' | ')}], ` +
        `ignored row ${item.ignoredRow} [${item.ignoredEF.join(' | ')}]`
      )
      .join('; ');

    Logger.log(
      'WARNING: Some duplicate records had equally complete but different E:F values. ' +
      `The earlier row was retained. Review: ${examples}`
    );
  }

  // ============================================================
  // BUILD UPDATED E:F FOR TARGET
  // ============================================================

  let matchedCount = 0;
  let changedCount = 0;
  const unmatchedTargetRows = [];

  const newTargetEF = targetData.map((row, index) => {
    const targetRowNumber = HEADER_ROWS + index + 1;
    const existingEF = [row[4], row[5]];

    // Leave rows alone when match fields A/C/D are all blank.
    if (isBlankMatchRow(row)) {
      return existingEF;
    }

    const sourceMatch = sourceByKey.get(makeMatchKey(row));

    // Do not erase target E:F when no source match exists.
    if (!sourceMatch) {
      unmatchedTargetRows.push(targetRowNumber);
      return existingEF;
    }

    matchedCount++;

    if (efValuesDiffer(existingEF, sourceMatch.valuesEF)) {
      changedCount++;
    }

    return sourceMatch.valuesEF;
  });

  // ============================================================
  // REPORT RESULTS
  // ============================================================

  Logger.log(`Source records indexed: ${sourceByKey.size}`);
  Logger.log(`Target records matched: ${matchedCount}`);
  Logger.log(`Target E:F rows that would change: ${changedCount}`);
  Logger.log(`Target rows without a source match: ${unmatchedTargetRows.length}`);

  if (unmatchedTargetRows.length > 0) {
    Logger.log(
      `First 50 unmatched target rows: ${unmatchedTargetRows
        .slice(0, 50)
        .join(', ')}`
    );
  }

  // ============================================================
  // PREVIEW OR UPDATE
  // ============================================================

  if (DRY_RUN) {
    Logger.log(
      'DRY RUN COMPLETE: No target cells were changed. ' +
      'Set DRY_RUN to false only after reviewing the results.'
    );
    return;
  }

  if (changedCount === 0) {
    Logger.log('No updates needed: target E:F already matches source E:F.');
    return;
  }

  // ============================================================
  // WRITE ONLY COLUMNS E:F IN TARGET
  // ============================================================

  targetSheet
    .getRange(HEADER_ROWS + 1, 5, targetRowCount, 2)
    .setValues(newTargetEF);

  Logger.log(
    `SYNC COMPLETE: Updated target E:F for ${changedCount} row(s).`
  );
}