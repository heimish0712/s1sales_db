/****************************************************
 * SimpleEditAutomation.gs
 * 단순 onEdit 통합 + 자동입력 누락 검산·보정 - 11단계
 *
 * 목적:
 * - 주소/연면적/계약단위/견적조건 자동입력의 헤더 조회를 1회로 통합
 * - 연면적→등급, 계약단위→기본조건 변경 결과를 같은 실행의 견적 계산에 즉시 반영
 * - 단순 onEdit 대기열 누락을 기존 5분 핵심 파이프라인의 제한적 보정으로 복구
 * - 수동 검산에서는 사용자 조정 가능성이 있는 계약조건의 비어 있지 않은 값은 보존
 ****************************************************/

var AUTOEDIT_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE11',
  headerRow: 2,
  dataStartRow: 3,

  repairCursorPropertyKey: 'AUTOEDIT_REPAIR_CURSOR_V1',
  repairLastRunPropertyKey: 'AUTOEDIT_REPAIR_LAST_RUN_V1',
  repairModuleKey: 'SIMPLE_EDIT_REPAIR',
  repairSheetNames: Object.freeze([
    '마스터시트(신규)',
    '수주확정/계약완료'
  ]),
  repairMaxRowsPerRun: 250,
  repairMaxRuntimeMs: 20 * 1000,
  maxErrorLength: 1500
});


/****************************************************
 * 단순 onEdit 통합 진입점
 ****************************************************/

function AUTOEDIT_handleSimpleOnEdit_(e) {
  if (!e || !e.range) return { status: 'IGNORED_NO_EVENT' };

  var context = AUTOEDIT_buildEditContext_(e);
  if (!context || context.targetEndRow < AUTOEDIT_CONFIG.dataStartRow) {
    return { status: 'IGNORED_HEADER_OR_EMPTY' };
  }

  var rules = AUTOEDIT_selectEditRules_(context);
  if (!rules.any) {
    return { status: 'IGNORED_UNWATCHED_COLUMN' };
  }

  return AUTOEDIT_applyEditRules_(context, rules);
}


function AUTOEDIT_buildEditContext_(e) {
  var range = e.range;
  var sheet = range.getSheet();
  var lastCol = sheet.getLastColumn();

  if (lastCol < 1) return null;

  var headers = sheet
    .getRange(AUTOEDIT_CONFIG.headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  var editedStartRow = range.getRow();
  var editedEndRow = editedStartRow + range.getNumRows() - 1;
  var editedStartCol = range.getColumn();
  var editedEndCol = editedStartCol + range.getNumColumns() - 1;

  return {
    event: e,
    range: range,
    sheet: sheet,
    spreadsheet: sheet.getParent(),
    lastCol: lastCol,
    headers: headers,
    columns: AUTOEDIT_buildColumnMap_(headers),
    editedStartRow: editedStartRow,
    editedEndRow: editedEndRow,
    editedStartCol: editedStartCol,
    editedEndCol: editedEndCol,
    targetStartRow: Math.max(editedStartRow, AUTOEDIT_CONFIG.dataStartRow),
    targetEndRow: editedEndRow
  };
}


function AUTOEDIT_buildColumnMap_(headers) {
  var finalQuoteMap = buildFinalQuoteHeaderMap_(headers);

  return {
    addressCol: AUTOEDIT_findHeaderCol_(headers, ['고객사 상세 주소']),
    regionCol: AUTOEDIT_findHeaderCol_(headers, ['지역구분']),
    areaCol: AUTOEDIT_findHeaderCol_(headers, ['연면적']),
    gradeCol: AUTOEDIT_findHeaderCol_(headers, ['관리등급']),
    contractUnitCol: AUTOEDIT_findHeaderCol_(headers, ['계약단위']),
    managerCol: AUTOEDIT_findHeaderCol_(headers, ['관리자선임여부', '관리자 선임 여부']),
    maintenanceCol: AUTOEDIT_findHeaderCol_(headers, ['유지점검']),
    performanceCol: AUTOEDIT_findHeaderCol_(headers, ['성능점검']),
    finalQuoteMap: finalQuoteMap
  };
}


function AUTOEDIT_selectEditRules_(context) {
  var columns = context.columns;
  var runRegion = AUTOEDIT_isEditedColumn_(context, columns.addressCol) &&
    columns.addressCol > 0 && columns.regionCol > 0;
  var runGrade = AUTOEDIT_isEditedColumn_(context, columns.areaCol) &&
    columns.areaCol > 0 && columns.gradeCol > 0;
  var runDefaults = AUTOEDIT_isEditedColumn_(context, columns.contractUnitCol) &&
    columns.contractUnitCol > 0 && columns.managerCol > 0 &&
    columns.maintenanceCol > 0 && columns.performanceCol > 0;

  var quoteMap = columns.finalQuoteMap;
  var quoteWatchedCols = quoteMap && quoteMap.ok
    ? [
      quoteMap.areaCol,
      quoteMap.gradeCol,
      quoteMap.discountCol,
      quoteMap.contractUnitCol,
      quoteMap.managerCol,
      quoteMap.maintenanceCol,
      quoteMap.performanceCol,
      quoteMap.vatCol
    ].filter(function(col) { return col > 0; })
    : [];

  var runQuote = !!(quoteMap && quoteMap.ok) && (
    runGrade ||
    runDefaults ||
    quoteWatchedCols.some(function(col) {
      return AUTOEDIT_isEditedColumn_(context, col);
    })
  );

  return {
    region: runRegion,
    grade: runGrade,
    defaults: runDefaults,
    quote: runQuote,
    any: runRegion || runGrade || runDefaults || runQuote
  };
}


function AUTOEDIT_applyEditRules_(context, rules) {
  var sheet = context.sheet;
  var columns = context.columns;
  var startRow = context.targetStartRow;
  var endRow = context.targetEndRow;
  var numRows = endRow - startRow + 1;

  var rawRows = null;
  var displayRows = null;

  if (rules.quote) {
    rawRows = sheet.getRange(startRow, 1, numRows, context.lastCol).getValues();
    displayRows = sheet.getRange(startRow, 1, numRows, context.lastCol).getDisplayValues();
  }

  var summary = {
    status: 'COMPLETED',
    sheetName: sheet.getName(),
    startRow: startRow,
    endRow: endRow,
    rowCount: numRows,
    regionRows: 0,
    gradeRows: 0,
    defaultRows: 0,
    quoteRows: 0
  };

  if (rules.region) {
    var addresses = displayRows
      ? AUTOEDIT_extractDisplayColumn_(displayRows, columns.addressCol)
      : sheet.getRange(startRow, columns.addressCol, numRows, 1).getDisplayValues();
    var regionValues = addresses.map(function(row) {
      return [getRegionByAddress_(String(row[0] || '').trim())];
    });

    sheet.getRange(startRow, columns.regionCol, numRows, 1).setValues(regionValues);
    AUTOEDIT_applyColumnToMemory_(rawRows, displayRows, columns.regionCol, regionValues);
    summary.regionRows = numRows;
  }

  if (rules.grade) {
    var areas = displayRows
      ? AUTOEDIT_extractDisplayColumn_(displayRows, columns.areaCol)
      : sheet.getRange(startRow, columns.areaCol, numRows, 1).getDisplayValues();
    var gradeValues = areas.map(function(row) {
      return [getManagementGradeByArea_(String(row[0] || '').trim())];
    });

    sheet.getRange(startRow, columns.gradeCol, numRows, 1).setValues(gradeValues);
    AUTOEDIT_applyColumnToMemory_(rawRows, displayRows, columns.gradeCol, gradeValues);
    summary.gradeRows = numRows;
  }

  if (rules.defaults) {
    var units = rawRows
      ? AUTOEDIT_extractRawColumn_(rawRows, columns.contractUnitCol)
      : sheet.getRange(startRow, columns.contractUnitCol, numRows, 1).getValues();
    var managerValues = [];
    var maintenanceValues = [];
    var performanceValues = [];

    units.forEach(function(row) {
      var defaults = getContractDefaultsByUnit_(row[0]);
      managerValues.push([defaults.manager]);
      maintenanceValues.push([defaults.maintenance]);
      performanceValues.push([defaults.performance]);
    });

    sheet.getRange(startRow, columns.managerCol, numRows, 1).setValues(managerValues);
    sheet.getRange(startRow, columns.maintenanceCol, numRows, 1).setValues(maintenanceValues);
    sheet.getRange(startRow, columns.performanceCol, numRows, 1).setValues(performanceValues);

    AUTOEDIT_applyColumnToMemory_(rawRows, displayRows, columns.managerCol, managerValues);
    AUTOEDIT_applyColumnToMemory_(rawRows, displayRows, columns.maintenanceCol, maintenanceValues);
    AUTOEDIT_applyColumnToMemory_(rawRows, displayRows, columns.performanceCol, performanceValues);
    summary.defaultRows = numRows;
  }

  if (rules.quote) {
    var basisMap = getFinalQuoteBasisMap_(context.spreadsheet);
    if (basisMap && Object.keys(basisMap).length > 0) {
      var outputValues = [];

      for (var i = 0; i < numRows; i++) {
        var calcResult = calculateFinalQuotePriceForRow_(
          rawRows[i],
          displayRows[i],
          columns.finalQuoteMap,
          basisMap
        );
        outputValues.push([calcResult.value]);
      }

      var quoteRange = sheet.getRange(
        startRow,
        columns.finalQuoteMap.finalQuoteCol,
        numRows,
        1
      );
      quoteRange.setValues(outputValues);
      quoteRange.setNumberFormat('₩#,##0');
      summary.quoteRows = numRows;
    } else {
      summary.status = 'COMPLETED_QUOTE_BASIS_MISSING';
    }
  }

  return summary;
}


/****************************************************
 * 수동 검산·보정
 ****************************************************/

function AUTOEDIT_auditActiveSheet() {
  return AUTOEDIT_runActiveSheetAudit_(true);
}


function AUTOEDIT_repairActiveSheet() {
  return AUTOEDIT_runActiveSheetAudit_(false);
}


function AUTOEDIT_runActiveSheetAudit_(dryRun) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getActiveSheet();
  var taskName = dryRun ? '자동입력 누락 검산' : '자동입력 누락 검산·보정';

  var result = AUTOMATION_runWithModuleLeaseOrThrow_(
    AUTOEDIT_CONFIG.repairModuleKey,
    taskName,
    function() {
      var lastRow = sheet.getLastRow();
      if (lastRow < AUTOEDIT_CONFIG.dataStartRow) {
        return AUTOEDIT_makeEmptyRepairResult_(sheet, dryRun, '처리할 데이터가 없습니다.');
      }

      return AUTOEDIT_auditAndRepairSheetRange_(
        sheet,
        AUTOEDIT_CONFIG.dataStartRow,
        lastRow,
        { dryRun: dryRun, safeDefaults: true }
      );
    },
    { ttlMs: 5 * 60 * 1000, waitMs: 1000 }
  );

  SpreadsheetApp.getUi().alert(
    taskName,
    AUTOEDIT_formatRepairResult_(result),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}


function AUTOEDIT_auditAndRepairSheetRange_(sheet, startRow, endRow, options) {
  options = options || {};
  var dryRun = options.dryRun === true;
  var safeDefaults = options.safeDefaults !== false;
  var lastCol = sheet.getLastColumn();

  if (lastCol < 1 || endRow < startRow) {
    return AUTOEDIT_makeEmptyRepairResult_(sheet, dryRun, '유효한 처리 범위가 없습니다.');
  }

  var headers = sheet
    .getRange(AUTOEDIT_CONFIG.headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];
  var columns = AUTOEDIT_buildColumnMap_(headers);
  var numRows = endRow - startRow + 1;
  var rawRows = sheet.getRange(startRow, 1, numRows, lastCol).getValues();
  var displayRows = sheet.getRange(startRow, 1, numRows, lastCol).getDisplayValues();
  var basisMap = columns.finalQuoteMap && columns.finalQuoteMap.ok
    ? getFinalQuoteBasisMap_(sheet.getParent())
    : {};

  var canRegion = columns.addressCol > 0 && columns.regionCol > 0;
  var canGrade = columns.areaCol > 0 && columns.gradeCol > 0;
  var canDefaults = columns.contractUnitCol > 0 && columns.managerCol > 0 &&
    columns.maintenanceCol > 0 && columns.performanceCol > 0;
  var canQuote = !!(columns.finalQuoteMap && columns.finalQuoteMap.ok) &&
    basisMap && Object.keys(basisMap).length > 0;

  if (!canRegion && !canGrade && !canDefaults && !canQuote) {
    return AUTOEDIT_makeEmptyRepairResult_(sheet, dryRun, '지원되는 자동입력 헤더가 없습니다.');
  }

  var regionValues = canRegion ? AUTOEDIT_extractRawColumn_(rawRows, columns.regionCol) : null;
  var gradeValues = canGrade ? AUTOEDIT_extractRawColumn_(rawRows, columns.gradeCol) : null;
  var managerValues = canDefaults ? AUTOEDIT_extractRawColumn_(rawRows, columns.managerCol) : null;
  var maintenanceValues = canDefaults ? AUTOEDIT_extractRawColumn_(rawRows, columns.maintenanceCol) : null;
  var performanceValues = canDefaults ? AUTOEDIT_extractRawColumn_(rawRows, columns.performanceCol) : null;
  var quoteValues = canQuote
    ? AUTOEDIT_extractRawColumn_(rawRows, columns.finalQuoteMap.finalQuoteCol)
    : null;

  var result = {
    version: AUTOEDIT_CONFIG.version,
    status: 'COMPLETED',
    dryRun: dryRun,
    sheetName: sheet.getName(),
    startRow: startRow,
    endRow: endRow,
    scannedRows: numRows,
    regionMismatch: 0,
    customRegionPreserved: 0,
    gradeMismatch: 0,
    defaultBlankFilled: 0,
    customDefaultPreserved: 0,
    quoteMismatch: 0,
    changedCells: 0,
    supportedRules: {
      region: canRegion,
      grade: canGrade,
      defaults: canDefaults,
      quote: canQuote
    },
    note: ''
  };

  for (var i = 0; i < numRows; i++) {
    if (canRegion) {
      var addressRaw = rawRows[i][columns.addressCol - 1];
      var addressDisplay = displayRows[i][columns.addressCol - 1];
      var currentRegionRaw = rawRows[i][columns.regionCol - 1];
      var currentRegionDisplay = displayRows[i][columns.regionCol - 1];
      var addressBlank = AUTOEDIT_isBlankCell_(addressRaw, addressDisplay);
      var regionBlank = AUTOEDIT_isBlankCell_(currentRegionRaw, currentRegionDisplay);

      if (!(addressBlank && regionBlank)) {
        var expectedRegion = getRegionByAddress_(String(addressDisplay || addressRaw || '').trim());
        var regionMatches = AUTOEDIT_valuesEqual_(
          currentRegionRaw,
          currentRegionDisplay,
          expectedRegion
        );

        // 상세 주소만으로 권역을 확정할 수 없어 '주소확인필요'가 나온 경우,
        // 사용자가 직접 지정해 둔 권역은 자동 보정에서 덮어쓰지 않는다.
        if (
          !regionMatches &&
          !addressBlank &&
          expectedRegion === '주소확인필요' &&
          !regionBlank
        ) {
          result.customRegionPreserved++;
        } else if (!regionMatches) {
          result.regionMismatch++;
          result.changedCells++;
          regionValues[i][0] = expectedRegion;
          AUTOEDIT_setMemoryCell_(rawRows, displayRows, i, columns.regionCol, expectedRegion);
        }
      }
    }

    if (canGrade) {
      var areaRaw = rawRows[i][columns.areaCol - 1];
      var areaDisplay = displayRows[i][columns.areaCol - 1];
      var currentGradeRaw = rawRows[i][columns.gradeCol - 1];
      var currentGradeDisplay = displayRows[i][columns.gradeCol - 1];
      var areaBlank = AUTOEDIT_isBlankCell_(areaRaw, areaDisplay);
      var gradeBlank = AUTOEDIT_isBlankCell_(currentGradeRaw, currentGradeDisplay);

      if (!(areaBlank && gradeBlank)) {
        var expectedGrade = getManagementGradeByArea_(String(areaDisplay || areaRaw || '').trim());
        if (!AUTOEDIT_valuesEqual_(currentGradeRaw, currentGradeDisplay, expectedGrade)) {
          result.gradeMismatch++;
          result.changedCells++;
          gradeValues[i][0] = expectedGrade;
          AUTOEDIT_setMemoryCell_(rawRows, displayRows, i, columns.gradeCol, expectedGrade);
        }
      }
    }

    if (canDefaults) {
      var defaults = getContractDefaultsByUnit_(rawRows[i][columns.contractUnitCol - 1]);
      var hasValidDefaults = defaults.manager !== '' ||
        defaults.maintenance !== '' || defaults.performance !== '';

      if (hasValidDefaults) {
        var defaultSpecs = [
          { col: columns.managerCol, expected: defaults.manager, values: managerValues },
          { col: columns.maintenanceCol, expected: defaults.maintenance, values: maintenanceValues },
          { col: columns.performanceCol, expected: defaults.performance, values: performanceValues }
        ];

        defaultSpecs.forEach(function(spec) {
          var currentRaw = rawRows[i][spec.col - 1];
          var currentDisplay = displayRows[i][spec.col - 1];

          if (AUTOEDIT_isBlankCell_(currentRaw, currentDisplay)) {
            result.defaultBlankFilled++;
            result.changedCells++;
            spec.values[i][0] = spec.expected;
            AUTOEDIT_setMemoryCell_(rawRows, displayRows, i, spec.col, spec.expected);
          } else if (
            safeDefaults &&
            !AUTOEDIT_valuesEqual_(currentRaw, currentDisplay, spec.expected)
          ) {
            result.customDefaultPreserved++;
          } else if (
            !safeDefaults &&
            !AUTOEDIT_valuesEqual_(currentRaw, currentDisplay, spec.expected)
          ) {
            result.defaultBlankFilled++;
            result.changedCells++;
            spec.values[i][0] = spec.expected;
            AUTOEDIT_setMemoryCell_(rawRows, displayRows, i, spec.col, spec.expected);
          }
        });
      }
    }

    if (canQuote) {
      var quoteResult = calculateFinalQuotePriceForRow_(
        rawRows[i],
        displayRows[i],
        columns.finalQuoteMap,
        basisMap
      );
      var currentQuoteRaw = rawRows[i][columns.finalQuoteMap.finalQuoteCol - 1];
      var currentQuoteDisplay = displayRows[i][columns.finalQuoteMap.finalQuoteCol - 1];

      if (!AUTOEDIT_valuesEqual_(currentQuoteRaw, currentQuoteDisplay, quoteResult.value)) {
        result.quoteMismatch++;
        result.changedCells++;
        quoteValues[i][0] = quoteResult.value;
      }
    }
  }

  if (!dryRun && result.changedCells > 0) {
    if (canRegion && result.regionMismatch > 0) {
      sheet.getRange(startRow, columns.regionCol, numRows, 1).setValues(regionValues);
    }
    if (canGrade && result.gradeMismatch > 0) {
      sheet.getRange(startRow, columns.gradeCol, numRows, 1).setValues(gradeValues);
    }
    if (canDefaults && result.defaultBlankFilled > 0) {
      sheet.getRange(startRow, columns.managerCol, numRows, 1).setValues(managerValues);
      sheet.getRange(startRow, columns.maintenanceCol, numRows, 1).setValues(maintenanceValues);
      sheet.getRange(startRow, columns.performanceCol, numRows, 1).setValues(performanceValues);
    }
    if (canQuote && result.quoteMismatch > 0) {
      var quoteRange = sheet.getRange(
        startRow,
        columns.finalQuoteMap.finalQuoteCol,
        numRows,
        1
      );
      quoteRange.setValues(quoteValues);
      quoteRange.setNumberFormat('₩#,##0');
    }
  }

  return result;
}


/****************************************************
 * 5분 핵심 파이프라인용 제한적 자동보정
 ****************************************************/

function AUTOEDIT_runScheduledRepairSlice_() {
  var startedAtMs = Date.now();
  var lease = AUTOMATION_acquireModuleLease_(AUTOEDIT_CONFIG.repairModuleKey, {
    taskName: '자동입력 누락 제한보정',
    ttlMs: 3 * 60 * 1000,
    waitMs: 0
  });

  if (!lease.acquired) {
    return {
      status: 'SKIPPED_LEASE_BUSY',
      scannedRows: 0,
      changedCells: 0,
      reason: lease.reason || 'LEASE_BUSY'
    };
  }

  var summary = {
    version: AUTOEDIT_CONFIG.version,
    status: 'STARTED',
    startedAt: new Date(startedAtMs).toISOString(),
    finishedAt: '',
    durationMs: 0,
    scannedRows: 0,
    changedCells: 0,
    sheets: [],
    cursorBefore: null,
    cursorAfter: null,
    error: ''
  };

  try {
    var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
    var props = PropertiesService.getScriptProperties();
    var cursor = AUTOEDIT_readRepairCursor_(props);
    summary.cursorBefore = cursor;

    var remainingRows = AUTOEDIT_CONFIG.repairMaxRowsPerRun;
    var sheetNames = AUTOEDIT_CONFIG.repairSheetNames;
    var visitedSheets = 0;

    while (
      remainingRows > 0 &&
      Date.now() - startedAtMs < AUTOEDIT_CONFIG.repairMaxRuntimeMs &&
      visitedSheets < sheetNames.length
    ) {
      var sheetIndex = Math.max(0, Math.min(sheetNames.length - 1, Number(cursor.sheetIndex) || 0));
      var sheetName = sheetNames[sheetIndex];
      var sheet = ss.getSheetByName(sheetName);

      if (!sheet || sheet.getLastRow() < AUTOEDIT_CONFIG.dataStartRow) {
        summary.sheets.push({
          sheetName: sheetName,
          status: sheet ? 'EMPTY' : 'MISSING',
          scannedRows: 0,
          changedCells: 0
        });
        cursor = AUTOEDIT_advanceRepairCursor_(cursor, sheetNames.length);
        visitedSheets++;
        continue;
      }

      var startRow = Math.max(AUTOEDIT_CONFIG.dataStartRow, Number(cursor.nextRow) || AUTOEDIT_CONFIG.dataStartRow);
      var lastRow = sheet.getLastRow();

      if (startRow > lastRow) {
        cursor = AUTOEDIT_advanceRepairCursor_(cursor, sheetNames.length);
        visitedSheets++;
        continue;
      }

      var endRow = Math.min(lastRow, startRow + remainingRows - 1);
      var result = AUTOEDIT_auditAndRepairSheetRange_(
        sheet,
        startRow,
        endRow,
        { dryRun: false, safeDefaults: true }
      );

      summary.scannedRows += Number(result.scannedRows || 0);
      summary.changedCells += Number(result.changedCells || 0);
      summary.sheets.push({
        sheetName: sheetName,
        status: result.status,
        startRow: startRow,
        endRow: endRow,
        scannedRows: Number(result.scannedRows || 0),
        changedCells: Number(result.changedCells || 0),
        regionMismatch: Number(result.regionMismatch || 0),
        customRegionPreserved: Number(result.customRegionPreserved || 0),
        gradeMismatch: Number(result.gradeMismatch || 0),
        defaultBlankFilled: Number(result.defaultBlankFilled || 0),
        customDefaultPreserved: Number(result.customDefaultPreserved || 0),
        quoteMismatch: Number(result.quoteMismatch || 0)
      });

      remainingRows -= Number(result.scannedRows || 0);

      if (endRow >= lastRow) {
        cursor = AUTOEDIT_advanceRepairCursor_(cursor, sheetNames.length);
        visitedSheets++;
      } else {
        cursor.nextRow = endRow + 1;
        break;
      }
    }

    cursor.updatedAt = new Date().toISOString();
    props.setProperty(AUTOEDIT_CONFIG.repairCursorPropertyKey, JSON.stringify(cursor));
    summary.cursorAfter = cursor;
    summary.status = 'COMPLETED';
  } catch (err) {
    summary.status = 'ERROR';
    summary.error = AUTOEDIT_truncate_(AUTOMATION_errorMessage_(err), AUTOEDIT_CONFIG.maxErrorLength);
    console.error('[AUTOEDIT_runScheduledRepairSlice_] ' + summary.error, err);
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }

  summary.finishedAt = new Date().toISOString();
  summary.durationMs = Date.now() - startedAtMs;

  try {
    PropertiesService.getScriptProperties().setProperty(
      AUTOEDIT_CONFIG.repairLastRunPropertyKey,
      JSON.stringify(summary)
    );
  } catch (ignoreStatusWriteError) {
    // 자동보정 본체 성공 여부에는 영향 없음
  }

  return summary;
}


function AUTOEDIT_runScheduledRepairNow() {
  var result = AUTOEDIT_runScheduledRepairSlice_();

  SpreadsheetApp.getUi().alert(
    '자동입력 누락 제한보정',
    [
      '상태: ' + String(result.status || ''),
      '검사 행: ' + Number(result.scannedRows || 0),
      '보정 셀: ' + Number(result.changedCells || 0),
      '소요시간: ' + (Number(result.durationMs || 0) / 1000).toFixed(1) + '초',
      result.error ? ('오류: ' + result.error) : ''
    ].filter(Boolean).join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}


function AUTOEDIT_getScheduledRepairLastRun() {
  var raw = PropertiesService.getScriptProperties()
    .getProperty(AUTOEDIT_CONFIG.repairLastRunPropertyKey);
  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (err) {
    return null;
  }
}


/****************************************************
 * 보조 함수
 ****************************************************/

function AUTOEDIT_findHeaderCol_(headers, aliases) {
  var targets = aliases.map(function(alias) {
    return AUTOEDIT_normalizeHeader_(alias);
  });

  for (var i = 0; i < headers.length; i++) {
    if (targets.indexOf(AUTOEDIT_normalizeHeader_(headers[i])) >= 0) {
      return i + 1;
    }
  }

  return -1;
}


function AUTOEDIT_normalizeHeader_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, '')
    .trim();
}


function AUTOEDIT_isEditedColumn_(context, col) {
  return col > 0 && col >= context.editedStartCol && col <= context.editedEndCol;
}


function AUTOEDIT_extractRawColumn_(rows, col) {
  return rows.map(function(row) {
    return [row[col - 1]];
  });
}


function AUTOEDIT_extractDisplayColumn_(rows, col) {
  return rows.map(function(row) {
    return [String(row[col - 1] || '')];
  });
}


function AUTOEDIT_applyColumnToMemory_(rawRows, displayRows, col, values) {
  if (!rawRows || !displayRows || col < 1) return;

  for (var i = 0; i < values.length; i++) {
    AUTOEDIT_setMemoryCell_(rawRows, displayRows, i, col, values[i][0]);
  }
}


function AUTOEDIT_setMemoryCell_(rawRows, displayRows, rowIndex, col, value) {
  if (!rawRows || !displayRows || col < 1) return;
  rawRows[rowIndex][col - 1] = value;
  displayRows[rowIndex][col - 1] = value === null || typeof value === 'undefined'
    ? ''
    : String(value);
}


function AUTOEDIT_isBlankCell_(rawValue, displayValue) {
  return String(
    displayValue !== null && typeof displayValue !== 'undefined'
      ? displayValue
      : rawValue
  ).trim() === '';
}


function AUTOEDIT_valuesEqual_(rawValue, displayValue, expectedValue) {
  if (expectedValue === null || typeof expectedValue === 'undefined' || expectedValue === '') {
    return AUTOEDIT_isBlankCell_(rawValue, displayValue);
  }

  if (typeof expectedValue === 'number') {
    if (typeof rawValue === 'number' && isFinite(rawValue)) {
      return rawValue === expectedValue;
    }

    var parsed = parseFinalQuoteNumber_(rawValue, displayValue);
    return parsed !== null && parsed === expectedValue;
  }

  var actualText = String(
    displayValue !== null && typeof displayValue !== 'undefined' && String(displayValue) !== ''
      ? displayValue
      : (rawValue === null || typeof rawValue === 'undefined' ? '' : rawValue)
  ).trim();

  return actualText === String(expectedValue).trim();
}


function AUTOEDIT_makeEmptyRepairResult_(sheet, dryRun, note) {
  return {
    version: AUTOEDIT_CONFIG.version,
    status: 'SKIPPED',
    dryRun: dryRun === true,
    sheetName: sheet ? sheet.getName() : '',
    startRow: 0,
    endRow: 0,
    scannedRows: 0,
    regionMismatch: 0,
    customRegionPreserved: 0,
    gradeMismatch: 0,
    defaultBlankFilled: 0,
    customDefaultPreserved: 0,
    quoteMismatch: 0,
    changedCells: 0,
    supportedRules: {},
    note: String(note || '')
  };
}


function AUTOEDIT_formatRepairResult_(result) {
  return [
    '시트: ' + String(result.sheetName || ''),
    '상태: ' + String(result.status || ''),
    '검사 행: ' + Number(result.scannedRows || 0),
    '지역 불일치: ' + Number(result.regionMismatch || 0),
    '수동 지역값 보존: ' + Number(result.customRegionPreserved || 0),
    '등급 불일치: ' + Number(result.gradeMismatch || 0),
    '빈 계약조건 보완: ' + Number(result.defaultBlankFilled || 0) + '셀',
    '사용자 조정값 보존: ' + Number(result.customDefaultPreserved || 0) + '셀',
    '견적가 불일치: ' + Number(result.quoteMismatch || 0),
    (result.dryRun ? '예상 보정 셀: ' : '실제 보정 셀: ') + Number(result.changedCells || 0),
    result.note ? ('비고: ' + result.note) : ''
  ].filter(Boolean).join('\n');
}


function AUTOEDIT_readRepairCursor_(props) {
  var raw = props.getProperty(AUTOEDIT_CONFIG.repairCursorPropertyKey);

  if (raw) {
    try {
      var parsed = JSON.parse(raw);
      return {
        sheetIndex: Math.max(0, Number(parsed.sheetIndex) || 0),
        nextRow: Math.max(AUTOEDIT_CONFIG.dataStartRow, Number(parsed.nextRow) || AUTOEDIT_CONFIG.dataStartRow),
        cycles: Math.max(0, Number(parsed.cycles) || 0),
        updatedAt: String(parsed.updatedAt || '')
      };
    } catch (ignoreCursorParseError) {
      props.deleteProperty(AUTOEDIT_CONFIG.repairCursorPropertyKey);
    }
  }

  return {
    sheetIndex: 0,
    nextRow: AUTOEDIT_CONFIG.dataStartRow,
    cycles: 0,
    updatedAt: ''
  };
}


function AUTOEDIT_advanceRepairCursor_(cursor, sheetCount) {
  var nextSheetIndex = (Number(cursor.sheetIndex) || 0) + 1;
  var cycles = Number(cursor.cycles) || 0;

  if (nextSheetIndex >= sheetCount) {
    nextSheetIndex = 0;
    cycles++;
  }

  return {
    sheetIndex: nextSheetIndex,
    nextRow: AUTOEDIT_CONFIG.dataStartRow,
    cycles: cycles,
    updatedAt: String(cursor.updatedAt || '')
  };
}


function AUTOEDIT_truncate_(value, maxLength) {
  var text = String(value || '');
  var limit = Math.max(1, Number(maxLength) || 1);
  if (text.length <= limit) return text;
  return text.slice(0, limit - 3) + '...';
}
