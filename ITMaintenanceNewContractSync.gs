/****************************************************
 * ITMaintenanceNewContractSync.gs
 * 수주확정/계약완료 → 서무 파일의 4_ 시작 정보통신 시트
 * 신규 계약 전용 안전 이식 모듈 - 26단계
 *
 * 핵심 원칙:
 * - 계약번호를 유일키로 사용한다.
 * - 대상 계약번호 행이 비어 있을 때만 입력한다.
 * - 대상 행에 기존 업무 데이터가 있으면 절대 덮어쓰지 않는다.
 * - 대상 계약번호 행이 없을 때만 마지막 계약행 아래에 신규 행을 만든다.
 * - 대상 열번호를 고정하지 않고 6행 실제 헤더명으로 찾아 기록한다.
 * - 필수 헤더가 없거나 동일 의미 헤더가 중복된 경우에만 쓰기 전에 중단한다.
 * - 수동 편집은 통합 onEdit에서 즉시 처리하고,
 *   스크립트 생성/이벤트 누락은 5분 핵심 파이프라인에서 보정한다.
 ****************************************************/

var ITMNEW_CONFIG_2026 = Object.freeze({
  version: '2026-08-05-PHASE30-DYNAMIC-TARGET-SHEET',

  sourceSheetName: '수주확정/계약완료',
  sourceHeaderRow: 1,
  sourceStartRow: 2,

  targetSheetPrefix: '4_',
  targetHeaderRow: 6,

  // 현재 4_ 서무 시트 기준 7~8행은 수식·합계 영역이고 계약 데이터는 9행부터다.
  targetDataStartRow: 9,

  logSheetName: '_정보통신유지보수이식로그',
  logHeaders: Object.freeze([
    '처리시각',
    '계약번호',
    '원본행',
    '대상행',
    '처리경로',
    '상태',
    '입력열',
    '오류/비고',
    '버전'
  ]),

  // 계약번호를 제외한 자동입력 필드. 대상 열번호는 실행 시 헤더에서 계산한다.
  businessFieldKeys: Object.freeze([
    'region', 'vendor', 'contractGrade', 'manager', 'customerName',
    'orderDate', 'contractPeriod', 'startDate', 'endDate', 'contractMonths',
    'appointment', 'maintenance', 'performance', 'contractPrice', 'vat',
    'memo', 'invoiceEmail', 'referrer'
  ]),

  maxErrorLength: 1500
});


/****************************************************
 * 설치형 onEdit 진입점
 ****************************************************/

/**
 * 수주확정/계약완료에서 사람이 입력·붙여넣기한 행을 즉시 검사한다.
 * 필수값이 아직 덜 입력된 경우 DEFERRED를 반환해 기존 재처리 큐가 보존한다.
 */
function ITMNEW_syncFromEdit_2026(e) {
  if (!e || !e.range || !e.source) {
    return { status: 'IGNORED_INVALID_EVENT' };
  }

  var config = ITMNEW_CONFIG_2026;
  var sheet = e.range.getSheet();

  if (sheet.getName() !== config.sourceSheetName) {
    return { status: 'IGNORED_UNRELATED_SHEET' };
  }

  var firstRow = Math.max(e.range.getRow(), config.sourceStartRow);
  var lastRow = e.range.getLastRow();

  if (lastRow < config.sourceStartRow || lastRow < firstRow) {
    return { status: 'IGNORED_HEADER_EDIT' };
  }

  var lease = AUTOMATION_acquireModuleLease_('IT_MAINTENANCE_SYNC', {
    taskName: 'ITMNEW_syncFromEdit_2026',
    ttlMs: 8 * 60 * 1000,
    waitMs: 500
  });

  if (!lease.acquired) {
    return {
      status: 'DEFERRED',
      reason: lease.reason || 'LEASE_BUSY',
      error: lease.reason || '정보통신유지보수 이식 lease를 얻지 못했습니다.'
    };
  }

  try {
    var result = ITMNEW_syncSourceRowsAppendOnly_2026_(
      firstRow,
      lastRow - firstRow + 1,
      {
        route: e.__automationRetryExecution === true ? 'RETRY_QUEUE' : 'ON_EDIT',
        logWaiting: e.__automationRetryExecution !== true,
        logAlreadyExists: e.__automationRetryExecution !== true
      }
    );

    if (result.waitingRequiredFields > 0) {
      return {
        status: 'DEFERRED',
        reason: 'WAITING_REQUIRED_FIELDS',
        error: '신규 계약 이식 필수값이 아직 완성되지 않은 행이 ' +
          result.waitingRequiredFields + '건 있습니다.',
        result: result
      };
    }

    return {
      status: 'SUCCESS',
      result: result
    };
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


/****************************************************
 * 5분 파이프라인·수동 실행
 ****************************************************/

/**
 * 5분 핵심 파이프라인 3단계 진입점.
 * 대상 행을 갱신하지 않고, 대상에 실제 업무 데이터가 없는 계약번호만 이식한다.
 */
function ITMNEW_runMissingContractSyncForPipeline_2026() {
  return AUTOMATION_runWithModuleLeaseOrThrow_(
    'IT_MAINTENANCE_SYNC',
    'ITMNEW_runMissingContractSyncForPipeline_2026',
    function () {
      return ITMNEW_syncAllMissingContracts_2026_({
        route: 'CORE_PIPELINE',
        logWaiting: false,
        logAlreadyExists: false
      });
    },
    { waitMs: 1000, ttlMs: 8 * 60 * 1000 }
  );
}


/**
 * 편집기/메뉴에서 누락 계약을 즉시 이식한다.
 */
function ITMNEW_syncMissingContractsNow_2026() {
  var result = AUTOMATION_runWithModuleLeaseOrThrow_(
    'IT_MAINTENANCE_SYNC',
    'ITMNEW_syncMissingContractsNow_2026',
    function () {
      return ITMNEW_syncAllMissingContracts_2026_({
        route: 'MANUAL',
        logWaiting: true,
        logAlreadyExists: false
      });
    },
    { waitMs: 1000, ttlMs: 8 * 60 * 1000 }
  );

  try {
    SpreadsheetApp.getUi().alert(
      '신규 계약 이식 완료\n\n' +
      '원본 계약: ' + Number(result.sourceContracts || 0) + '건\n' +
      '신규 이식: ' + Number(result.copied || 0) + '건\n' +
      '준비된 빈 행 사용: ' + Number(result.filledPreparedRows || 0) + '건\n' +
      '새 행 생성: ' + Number(result.appendedRows || 0) + '건\n' +
      '기존 데이터 보호로 건너뜀: ' + Number(result.skippedExistingData || 0) + '건\n' +
      '필수값 대기: ' + Number(result.waitingRequiredFields || 0) + '건'
    );
  } catch (ignoreUiError) {}

  return result;
}


/**
 * 실제 쓰기 없이 현재 누락·대기·기존 건수를 확인한다.
 */
function ITMNEW_previewMissingContracts_2026() {
  var result = ITMNEW_buildSyncPlan_2026_({ route: 'PREVIEW' });

  try {
    SpreadsheetApp.getUi().alert(
      '신규 계약 이식 미리보기\n\n' +
      '원본 계약: ' + Number(result.sourceContracts || 0) + '건\n' +
      '이식 예정: ' + Number(result.readyToCopy || 0) + '건\n' +
      '  - 준비된 빈 행: ' + Number(result.preparedRows || 0) + '건\n' +
      '  - 신규 행 필요: ' + Number(result.appendRows || 0) + '건\n' +
      '기존 데이터 보호: ' + Number(result.skippedExistingData || 0) + '건\n' +
      '필수값 대기: ' + Number(result.waitingRequiredFields || 0) + '건\n' +
      '대상 계약번호 중복: ' + Number(result.duplicateTargetIds || 0) + '건'
    );
  } catch (ignoreUiError) {}

  return result;
}


function ITMNEW_showTransferLogSheet_2026() {
  var sheet = ITMNEW_getOrCreateLogSheet_2026_();
  sheet.showSheet();
  sheet.activate();
  return { status: 'SUCCESS', sheetName: sheet.getName() };
}


/****************************************************
 * append-only 이식 본체
 ****************************************************/

function ITMNEW_syncAllMissingContracts_2026_(options) {
  var config = ITMNEW_CONFIG_2026;
  var sourceSheet = ITMAINT_getSourceSheet_2026_();
  var lastRow = sourceSheet.getLastRow();

  if (lastRow < config.sourceStartRow) {
    return ITMNEW_emptyResult_2026_(String(options && options.route || 'UNKNOWN'));
  }

  return ITMNEW_syncSourceRowsAppendOnly_2026_(
    config.sourceStartRow,
    lastRow - config.sourceStartRow + 1,
    options || {}
  );
}


function ITMNEW_syncSourceRowsAppendOnly_2026_(startRow, rowCount, options) {
  var config = ITMNEW_CONFIG_2026;
  var route = String(options && options.route || 'UNKNOWN');
  var sourceSheet = ITMAINT_getSourceSheet_2026_();
  var masterSheet = ITMAINT_getMasterSheet_2026_();
  var targetSheet = ITMAINT_getTargetSheet_2026_();

  var sourceSchema = ITMAINT_buildSchema_2026_(
    sourceSheet,
    config.sourceHeaderRow,
    ITMAINT_getRequiredSourceHeaders_2026_(),
    '수주확정/계약완료'
  );

  var masterSchema = ITMAINT_buildSchema_2026_(
    masterSheet,
    ITMAINT_getConfig_2026_().masterHeaderRow,
    ITMAINT_getRequiredMasterHeaders_2026_(),
    '마스터시트(신규)'
  );

  var targetSchema = ITMAINT_buildSchema_2026_(
    targetSheet,
    config.targetHeaderRow,
    [],
    targetSheet.getName()
  );

  // 필수 대상 헤더를 찾지 못하거나 동일 의미 헤더가 중복되면 쓰기 전에 중단한다.
  var targetFieldMap = ITMAINT_validateTargetLayout_2026_(targetSchema);

  var sourceLastRow = sourceSheet.getLastRow();
  var safeStartRow = Math.max(config.sourceStartRow, Number(startRow) || config.sourceStartRow);
  var safeLastRow = Math.min(
    sourceLastRow,
    safeStartRow + Math.max(0, Number(rowCount) || 0) - 1
  );

  if (safeLastRow < safeStartRow) {
    return ITMNEW_emptyResult_2026_(route);
  }

  var sourceRows = sourceSheet
    .getRange(
      safeStartRow,
      1,
      safeLastRow - safeStartRow + 1,
      sourceSchema.lastCol
    )
    .getValues();

  var masterLookup = ITMAINT_buildMasterLookup_2026_(masterSheet, masterSchema);
  var targetIndex = ITMNEW_buildTargetIndex_2026_(targetSheet, targetFieldMap);
  var result = ITMNEW_emptyResult_2026_(route);
  var logs = [];

  result.sourceRows = sourceRows.length;
  result.duplicateTargetIds = targetIndex.duplicateIds.length;

  if (targetIndex.duplicateIds.length > 0) {
    throw new Error(
      targetSheet.getName() + ' 계약번호 헤더 열에 중복값이 있어 신규 이식을 중단했습니다: ' +
      targetIndex.duplicateIds.slice(0, 20).join(', ')
    );
  }

  sourceRows.forEach(function (sourceRow, offset) {
    var sourceRowNumber = safeStartRow + offset;
    var contractNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, '계약번호')
    );

    if (!contractNo) {
      result.skippedNoContractNo++;
      return;
    }

    result.sourceContracts++;

    var targetRowNumber = targetIndex.rowById[contractNo] || 0;
    var preparedExistingRow = false;

    // 기존 대상 행에 업무 데이터가 있으면 원본 필수값 상태와 관계없이 즉시 보호한다.
    // 과거 계약의 일부 원본 필드가 비어 있어도 재처리 큐에 무한 등록하지 않는다.
    if (
      targetRowNumber > 0 &&
      targetIndex.hasBusinessDataById[contractNo] === true
    ) {
      result.skippedExistingData++;

      if (options && options.logAlreadyExists) {
        logs.push(ITMNEW_makeLogRow_2026_(
          contractNo,
          sourceRowNumber,
          targetRowNumber,
          route,
          'SKIPPED_ALREADY_EXISTS',
          '',
          '대상 행에 기존 업무 데이터가 있어 자동 덮어쓰기를 차단했습니다.'
        ));
      }
      return;
    }

    var customerNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, '고객번호')
    );

    var masterRow = masterLookup.byContractNo[contractNo] ||
      masterLookup.byCustomerNo[customerNo] ||
      null;

    var targetRecord = ITMAINT_makeTargetRecord_2026_(
      sourceRow,
      sourceSchema,
      masterRow,
      masterSchema
    );

    var requiredCheck = ITMNEW_validateRequiredMappedValues_2026_(
      sourceRow,
      sourceSchema,
      targetRecord
    );

    if (!requiredCheck.ready) {
      result.waitingRequiredFields++;
      result.waitingContracts.push({
        contractNo: contractNo,
        sourceRow: sourceRowNumber,
        missing: requiredCheck.missing
      });

      if (options && options.logWaiting) {
        logs.push(ITMNEW_makeLogRow_2026_(
          contractNo,
          sourceRowNumber,
          targetRowNumber || '',
          route,
          'WAITING_REQUIRED_FIELDS',
          '',
          requiredCheck.missing.join(', ')
        ));
      }
      return;
    }

    if (targetRowNumber > 0) {
      preparedExistingRow = true;
    } else {
      targetRowNumber = ITMNEW_appendTargetTemplateRow_2026_(
        targetSheet,
        targetIndex.lastContractRow
      );
      targetIndex.lastContractRow = targetRowNumber;
      targetIndex.rowById[contractNo] = targetRowNumber;
    }

    var writeResult = ITMNEW_writeMappedTargetRow_2026_(
      targetSheet,
      targetRowNumber,
      targetRecord,
      targetFieldMap
    );
    targetIndex.hasBusinessDataById[contractNo] = true;

    result.copied++;

    if (preparedExistingRow) {
      result.filledPreparedRows++;
    } else {
      result.appendedRows++;
    }

    logs.push(ITMNEW_makeLogRow_2026_(
      contractNo,
      sourceRowNumber,
      targetRowNumber,
      route,
      'COPIED',
      String(writeResult.segments || targetFieldMap.mappingSummary || ''),
      preparedExistingRow
        ? '기존 계약번호 준비행에 신규 계약을 입력했습니다.'
        : '대상에 계약번호가 없어 새 행을 생성한 뒤 입력했습니다.'
    ));
  });

  if (logs.length > 0) {
    ITMNEW_appendLogRows_2026_(logs);
  }

  return result;
}


/**
 * 실제 쓰기 없는 미리보기 계획.
 */
function ITMNEW_buildSyncPlan_2026_(options) {
  var config = ITMNEW_CONFIG_2026;
  var sourceSheet = ITMAINT_getSourceSheet_2026_();
  var masterSheet = ITMAINT_getMasterSheet_2026_();
  var targetSheet = ITMAINT_getTargetSheet_2026_();

  var sourceSchema = ITMAINT_buildSchema_2026_(
    sourceSheet,
    config.sourceHeaderRow,
    ITMAINT_getRequiredSourceHeaders_2026_(),
    '수주확정/계약완료'
  );

  var masterSchema = ITMAINT_buildSchema_2026_(
    masterSheet,
    ITMAINT_getConfig_2026_().masterHeaderRow,
    ITMAINT_getRequiredMasterHeaders_2026_(),
    '마스터시트(신규)'
  );

  var targetSchema = ITMAINT_buildSchema_2026_(
    targetSheet,
    config.targetHeaderRow,
    [],
    targetSheet.getName()
  );

  var targetFieldMap = ITMAINT_validateTargetLayout_2026_(targetSchema);

  var result = {
    version: config.version,
    route: String(options && options.route || 'PREVIEW'),
    sourceContracts: 0,
    readyToCopy: 0,
    preparedRows: 0,
    appendRows: 0,
    skippedExistingData: 0,
    waitingRequiredFields: 0,
    duplicateTargetIds: 0
  };

  var lastRow = sourceSheet.getLastRow();
  if (lastRow < config.sourceStartRow) return result;

  var rows = sourceSheet
    .getRange(
      config.sourceStartRow,
      1,
      lastRow - config.sourceStartRow + 1,
      sourceSchema.lastCol
    )
    .getValues();

  var masterLookup = ITMAINT_buildMasterLookup_2026_(masterSheet, masterSchema);
  var targetIndex = ITMNEW_buildTargetIndex_2026_(targetSheet, targetFieldMap);
  result.duplicateTargetIds = targetIndex.duplicateIds.length;

  rows.forEach(function (sourceRow) {
    var contractNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, '계약번호')
    );

    if (!contractNo) return;
    result.sourceContracts++;

    var targetRowNumber = targetIndex.rowById[contractNo] || 0;

    if (
      targetRowNumber > 0 &&
      targetIndex.hasBusinessDataById[contractNo] === true
    ) {
      result.skippedExistingData++;
      return;
    }

    var customerNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, '고객번호')
    );
    var masterRow = masterLookup.byContractNo[contractNo] ||
      masterLookup.byCustomerNo[customerNo] ||
      null;
    var targetRecord = ITMAINT_makeTargetRecord_2026_(
      sourceRow,
      sourceSchema,
      masterRow,
      masterSchema
    );

    if (!ITMNEW_validateRequiredMappedValues_2026_(
      sourceRow,
      sourceSchema,
      targetRecord
    ).ready) {
      result.waitingRequiredFields++;
      return;
    }

    result.readyToCopy++;

    if (targetRowNumber > 0) {
      result.preparedRows++;
    } else {
      result.appendRows++;
    }
  });

  return result;
}


/****************************************************
 * 대상 행 판정·생성·쓰기
 ****************************************************/

function ITMNEW_buildTargetIndex_2026_(targetSheet, targetFieldMap) {
  var config = ITMNEW_CONFIG_2026;
  var lastRow = Math.max(targetSheet.getLastRow(), config.targetDataStartRow - 1);
  var rowById = {};
  var hasBusinessDataById = {};
  var duplicateIds = [];
  var lastContractRow = config.targetDataStartRow - 1;

  if (lastRow < config.targetDataStartRow) {
    return {
      rowById: rowById,
      hasBusinessDataById: hasBusinessDataById,
      duplicateIds: duplicateIds,
      lastContractRow: lastContractRow
    };
  }

  var lastCol = Math.max(1, targetSheet.getLastColumn());
  var rows = targetSheet
    .getRange(
      config.targetDataStartRow,
      1,
      lastRow - config.targetDataStartRow + 1,
      lastCol
    )
    .getDisplayValues();
  var contractNoIndex = Number(targetFieldMap.indexByField.contractNo);
  var businessIndexes = config.businessFieldKeys
    .map(function(fieldKey) { return targetFieldMap.indexByField[fieldKey]; })
    .filter(function(index) { return index !== undefined; });

  rows.forEach(function(row, index) {
    var id = ITMAINT_normalizeId_2026_(row[contractNoIndex]);
    var rowNumber = config.targetDataStartRow + index;

    if (!id) return;
    lastContractRow = rowNumber;

    if (rowById[id]) {
      if (duplicateIds.indexOf(id) < 0) duplicateIds.push(id);
      return;
    }

    rowById[id] = rowNumber;
    hasBusinessDataById[id] = businessIndexes.some(function(columnIndex) {
      return String(row[columnIndex] || '').trim() !== '';
    });
  });

  return {
    rowById: rowById,
    hasBusinessDataById: hasBusinessDataById,
    duplicateIds: duplicateIds,
    lastContractRow: lastContractRow
  };
}


function ITMNEW_targetRowHasBusinessData_2026_(
  targetSheet,
  rowNumber,
  targetFieldMap
) {
  var columns = ITMNEW_CONFIG_2026.businessFieldKeys
    .map(function(fieldKey) { return targetFieldMap.columnByField[fieldKey]; })
    .filter(function(column) { return Number(column) > 0; });

  for (var i = 0; i < columns.length; i++) {
    var value = targetSheet
      .getRange(rowNumber, Number(columns[i]))
      .getDisplayValue();

    if (String(value || '').trim() !== '') return true;
  }

  return false;
}


function ITMNEW_appendTargetTemplateRow_2026_(targetSheet, lastContractRow) {
  var config = ITMNEW_CONFIG_2026;
  var templateRow = Math.max(config.targetDataStartRow, Number(lastContractRow) || 0);
  var occupiedLastRow = Math.max(targetSheet.getLastRow(), templateRow);
  var newRow = Math.max(config.targetDataStartRow, occupiedLastRow + 1);

  ITMAINT_ensureTargetRows_2026_(targetSheet, newRow);

  if (templateRow >= config.targetDataStartRow && templateRow < newRow) {
    var currentLastCol = Math.max(1, targetSheet.getLastColumn());
    var templateRange = targetSheet.getRange(templateRow, 1, 1, currentLastCol);
    var destinationRange = targetSheet.getRange(newRow, 1, 1, currentLastCol);

    templateRange.copyTo(
      destinationRange,
      SpreadsheetApp.CopyPasteType.PASTE_FORMAT,
      false
    );
    templateRange.copyTo(
      destinationRange,
      SpreadsheetApp.CopyPasteType.PASTE_DATA_VALIDATION,
      false
    );
    templateRange.copyTo(
      destinationRange,
      SpreadsheetApp.CopyPasteType.PASTE_FORMULA,
      false
    );

    try {
      targetSheet.setRowHeight(newRow, targetSheet.getRowHeight(templateRow));
    } catch (ignoreRowHeightError) {}
  }

  return newRow;
}


function ITMNEW_writeMappedTargetRow_2026_(
  targetSheet,
  rowNumber,
  targetRecord,
  targetFieldMap
) {
  return ITMAINT_writeTargetRecordByHeader_2026_(
    targetSheet,
    rowNumber,
    targetRecord,
    targetFieldMap
  );
}


/****************************************************
 * 필수값·로그
 ****************************************************/

function ITMNEW_validateRequiredMappedValues_2026_(
  sourceRow,
  sourceSchema,
  targetRecord
) {
  var missing = [];

  function sourceValue(headerName) {
    return ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, headerName);
  }

  function isBlank(value) {
    return value === '' || value === null || value === undefined;
  }

  if (isBlank(sourceValue('계약번호'))) missing.push('계약번호');
  if (isBlank(sourceValue('고객사명'))) missing.push('고객사명');
  if (isBlank(sourceValue('수행사'))) missing.push('수행사');
  if (isBlank(sourceValue('선임유형'))) missing.push('선임유형');
  if (isBlank(sourceValue('계약가'))) missing.push('계약가');
  if (isBlank(sourceValue('vat'))) missing.push('VAT');
  if (isBlank(sourceValue('계약기간'))) missing.push('계약기간');

  // 계약기간 문자열 또는 마스터 일자를 통해 시작·종료일을 확정할 수 있어야 한다.
  if (isBlank(targetRecord.startDate)) missing.push('계약시작일');
  if (isBlank(targetRecord.endDate)) missing.push('계약종료일');
  if (isBlank(targetRecord.contractMonths)) missing.push('계약기간(개월)');

  return {
    ready: missing.length === 0,
    missing: missing
  };
}


function ITMNEW_getOrCreateLogSheet_2026_() {
  var ss = SYSTEMLOG_getSpreadsheet_();
  var config = ITMNEW_CONFIG_2026;
  var sheet = ss.getSheetByName(config.logSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(config.logSheetName);
  }

  var headers = config.logHeaders.slice();
  var currentHeaders = sheet
    .getRange(1, 1, 1, headers.length)
    .getDisplayValues()[0];
  var needsHeader = currentHeaders.some(function (value, index) {
    return String(value || '').trim() !== headers[index];
  });

  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  try {
    sheet.hideSheet();
  } catch (ignoreHideError) {}

  return sheet;
}


function ITMNEW_appendLogRows_2026_(rows) {
  if (!rows || rows.length === 0) return;

  var sheet = ITMNEW_getOrCreateLogSheet_2026_();
  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


function ITMNEW_makeLogRow_2026_(
  contractNo,
  sourceRow,
  targetRow,
  route,
  status,
  writtenColumns,
  detail
) {
  return [
    new Date(),
    contractNo,
    sourceRow || '',
    targetRow || '',
    route || '',
    status || '',
    writtenColumns || '',
    String(detail || '').slice(0, ITMNEW_CONFIG_2026.maxErrorLength),
    ITMNEW_CONFIG_2026.version
  ];
}


function ITMNEW_emptyResult_2026_(route) {
  return {
    version: ITMNEW_CONFIG_2026.version,
    route: route || '',
    sourceRows: 0,
    sourceContracts: 0,
    copied: 0,
    filledPreparedRows: 0,
    appendedRows: 0,
    skippedExistingData: 0,
    waitingRequiredFields: 0,
    skippedNoContractNo: 0,
    duplicateTargetIds: 0,
    waitingContracts: []
  };
}
