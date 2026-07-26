/****************************************************
 * AppointmentStatusSync.gs
 * 수행사 선임신고 현황 → 마스터 선임완료여부 동기화 - 17단계
 *
 * 소스 시트:
 * - KJ 선임신고 현황(내부용)
 * - 일신 선임신고 현황(내부용)
 * - 삼구 선임신고 현황(내부용)
 *
 * 상태 매핑:
 * - 신고 수리      → 완료
 * - 신고 접수      → 진행중(접수)
 * - 지자체 반려    → 진행중(반려)
 * - 비선임/자체선임 → 비대상
 ****************************************************/

var APPTSYNC_CONFIG = Object.freeze({
  version: '2026-07-26-PHASE17',
  sourceSheetNames: Object.freeze([
    'KJ 선임신고 현황(내부용)',
    '일신 선임신고 현황(내부용)',
    '삼구 선임신고 현황(내부용)'
  ]),
  sourceHeaderRow: 2,
  sourceDataStartRow: 3,
  sourceHeaders: Object.freeze({
    contractNo: '계약번호',
    customerNo: '고객번호',
    reportStatus: '신고 여부',
    companyName: '고객사명'
  }),

  masterSheetName: '마스터시트(신규)',
  masterHeaderRow: 2,
  masterDataStartRow: 3,
  masterHeaders: Object.freeze({
    customerNo: '고객번호',
    companyName: '회사명',
    contractNo: '발주번호',
    completionStatus: '선임완료여부'
  }),

  logSheetName: '_선임완료동기화로그',
  pendingPropertyKey: 'APPOINTMENT_STATUS_SYNC_PENDING_V1',
  lastRunPropertyKey: 'APPOINTMENT_STATUS_SYNC_LAST_RUN_V1',
  pipelineIntervalMs: 4 * 60 * 1000,
  maxLogTextLength: 1000,

  logHeaders: Object.freeze([
    '기록일시', '실행경로', '상태', '소스시트', '소스행',
    '계약번호', '고객번호', '고객사명', '신고여부', '변환값',
    '마스터행', '마스터회사명', '기존값', '신규값', '메시지', '버전'
  ])
});


/****************************************************
 * 공개 함수
 ****************************************************/

function APPTSYNC_previewAppointmentStatusSync() {
  var result = APPTSYNC_syncAll_({
    dryRun: true,
    source: 'MANUAL_PREVIEW'
  });

  SpreadsheetApp.getUi().alert(
    '선임신고 현황 동기화 미리보기',
    APPTSYNC_summaryText_(result),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}


function APPTSYNC_syncAppointmentStatusesNow() {
  var ui = SpreadsheetApp.getUi();
  var preview = APPTSYNC_syncAll_({
    dryRun: true,
    source: 'MANUAL_PREVIEW_BEFORE_WRITE'
  });

  var answer = ui.alert(
    '선임완료여부 일괄 반영',
    APPTSYNC_summaryText_(preview) +
      '\n\n마스터시트의 선임완료여부를 실제 반영하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (answer !== ui.Button.YES) return preview;

  var result = APPTSYNC_syncAll_({
    dryRun: false,
    source: 'MANUAL_FULL_SYNC'
  });

  ui.alert(
    '선임완료여부 일괄 반영 완료',
    APPTSYNC_summaryText_(result),
    ui.ButtonSet.OK
  );

  return result;
}


function APPTSYNC_showAppointmentStatusSyncLog() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = APPTSYNC_getOrCreateLogSheet_(ss);
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


/** 설치형 중앙 onEdit에서 호출 */
function APPTSYNC_handleStatusSheetEdit(e) {
  if (!e || !e.range) return { status: 'IGNORED_INVALID_EVENT' };

  var sheet = e.range.getSheet();
  var sheetName = String(sheet.getName() || '');
  if (APPTSYNC_CONFIG.sourceSheetNames.indexOf(sheetName) < 0) {
    return { status: 'IGNORED_UNRELATED_SHEET' };
  }

  var firstColumn = e.range.getColumn();
  var lastColumn = typeof e.range.getLastColumn === 'function'
    ? e.range.getLastColumn()
    : firstColumn + e.range.getNumColumns() - 1;
  var lastRow = typeof e.range.getLastRow === 'function'
    ? e.range.getLastRow()
    : e.range.getRow() + e.range.getNumRows() - 1;

  if (lastRow < APPTSYNC_CONFIG.sourceDataStartRow) {
    return { status: 'IGNORED_HEADER_ROW' };
  }

  // 계약번호·고객번호·신고 여부 변경만 즉시 동기화한다.
  if (lastColumn < 1 || firstColumn > 3) {
    return { status: 'IGNORED_UNRELATED_COLUMN' };
  }

  PropertiesService.getScriptProperties().setProperty(
    APPTSYNC_CONFIG.pendingPropertyKey,
    JSON.stringify({
      requestedAt: new Date().toISOString(),
      source: 'ON_EDIT',
      sheetName: sheetName,
      range: e.range.getA1Notation()
    })
  );

  try {
    return APPTSYNC_syncAll_({
      dryRun: false,
      source: 'ON_EDIT',
      sourceSheetName: sheetName,
      sourceRange: e.range.getA1Notation()
    });
  } catch (err) {
    return {
      status: err && err.automationLeaseBusy ? 'DEFERRED' : 'ERROR',
      reason: AUTOMATION_errorMessage_(err),
      error: AUTOMATION_errorMessage_(err)
    };
  }
}


/** 기존 5분 핵심 파이프라인의 보정 진입점 */
function APPTSYNC_runPipelineSafetySync_() {
  var props = PropertiesService.getScriptProperties();
  var pending = String(props.getProperty(APPTSYNC_CONFIG.pendingPropertyKey) || '');
  var lastRun = APPTSYNC_readJson_(props.getProperty(APPTSYNC_CONFIG.lastRunPropertyKey));
  var lastFinishedMs = Date.parse(String(lastRun && lastRun.finishedAt || ''));
  var due = !!pending || !isFinite(lastFinishedMs) ||
    Date.now() - lastFinishedMs >= APPTSYNC_CONFIG.pipelineIntervalMs;

  if (!due) {
    return {
      status: 'SKIPPED_NOT_DUE',
      pending: false,
      lastFinishedAt: lastRun && lastRun.finishedAt || ''
    };
  }

  try {
    return APPTSYNC_syncAll_({
      dryRun: false,
      source: pending ? 'PIPELINE_PENDING' : 'PIPELINE_SAFETY'
    });
  } catch (err) {
    return {
      status: err && err.automationLeaseBusy ? 'DEFERRED' : 'ERROR',
      error: AUTOMATION_errorMessage_(err)
    };
  }
}


/****************************************************
 * 전체 동기화
 ****************************************************/

function APPTSYNC_syncAll_(options) {
  options = options || {};
  var dryRun = options.dryRun === true;
  var source = String(options.source || 'UNKNOWN');

  return AUTOMATION_runWithModuleLeaseOrThrow_(
    'APPOINTMENT_STATUS_SYNC',
    'APPTSYNC_syncAll_',
    function () {
      var startedAtMs = Date.now();
      var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
      var masterSheet = ss.getSheetByName(APPTSYNC_CONFIG.masterSheetName);
      if (!masterSheet) throw new Error('마스터시트(신규)를 찾을 수 없습니다.');

      var masterContext = APPTSYNC_buildMasterContext_(masterSheet);
      var collection = APPTSYNC_collectSourceStatuses_(ss, masterContext);
      var applyResult = APPTSYNC_applyStatuses_(
        masterSheet,
        masterContext,
        collection,
        dryRun
      );

      var finishedAt = new Date().toISOString();
      var result = {
        version: APPTSYNC_CONFIG.version,
        status: applyResult.errorCount > 0
          ? 'COMPLETED_WITH_ERRORS'
          : (collection.conflictCount > 0 ? 'COMPLETED_WITH_CONFLICTS' : 'SUCCESS'),
        dryRun: dryRun,
        source: source,
        startedAt: new Date(startedAtMs).toISOString(),
        finishedAt: finishedAt,
        durationMs: Date.now() - startedAtMs,
        sourceRows: collection.sourceRowCount,
        mappedRows: collection.mappedRowCount,
        ignoredRows: collection.ignoredRowCount,
        unmatchedRows: collection.unmatchedCount,
        conflictRows: collection.conflictCount,
        desiredRows: Object.keys(collection.desiredByMasterRow).length,
        writeCount: applyResult.writeCount,
        unchangedCount: applyResult.unchangedCount,
        errorCount: applyResult.errorCount
      };

      if (!dryRun) {
        APPTSYNC_appendLogs_(ss, collection.logs.concat(applyResult.logs), source);
        var props = PropertiesService.getScriptProperties();
        props.deleteProperty(APPTSYNC_CONFIG.pendingPropertyKey);
        props.setProperty(APPTSYNC_CONFIG.lastRunPropertyKey, JSON.stringify(result));
      }

      return result;
    },
    {
      ttlMs: 3 * 60 * 1000,
      waitMs: 300,
      taskName: '선임신고 현황 → 마스터 선임완료여부 동기화'
    }
  );
}


function APPTSYNC_collectSourceStatuses_(ss, masterContext) {
  var desiredByMasterRow = {};
  var logs = [];
  var sourceRowCount = 0;
  var mappedRowCount = 0;
  var ignoredRowCount = 0;
  var unmatchedCount = 0;
  var conflictCount = 0;

  for (var si = 0; si < APPTSYNC_CONFIG.sourceSheetNames.length; si++) {
    var sheetName = APPTSYNC_CONFIG.sourceSheetNames[si];
    var sheet = ss.getSheetByName(sheetName);
    if (!sheet) {
      logs.push(APPTSYNC_makeLog_({
        status: 'SOURCE_SHEET_MISSING',
        sourceSheet: sheetName,
        message: '소스 시트를 찾을 수 없습니다.'
      }));
      conflictCount++;
      continue;
    }

    var headerMap = APPTSYNC_buildHeaderMap_(
      sheet,
      APPTSYNC_CONFIG.sourceHeaderRow,
      APPTSYNC_CONFIG.sourceHeaders
    );
    var lastRow = sheet.getLastRow();
    var rowCount = Math.max(0, lastRow - APPTSYNC_CONFIG.sourceDataStartRow + 1);
    if (!rowCount) continue;

    var firstColumn = Math.min(
      headerMap.contractNo,
      headerMap.customerNo,
      headerMap.reportStatus,
      headerMap.companyName
    );
    var lastColumn = Math.max(
      headerMap.contractNo,
      headerMap.customerNo,
      headerMap.reportStatus,
      headerMap.companyName
    );
    var values = sheet.getRange(
      APPTSYNC_CONFIG.sourceDataStartRow,
      firstColumn,
      rowCount,
      lastColumn - firstColumn + 1
    ).getDisplayValues();

    for (var ri = 0; ri < values.length; ri++) {
      var sourceRow = APPTSYNC_CONFIG.sourceDataStartRow + ri;
      var row = values[ri];
      var contractNo = APPTSYNC_valueAtAbsoluteColumn_(row, firstColumn, headerMap.contractNo);
      var customerNo = APPTSYNC_valueAtAbsoluteColumn_(row, firstColumn, headerMap.customerNo);
      var reportStatus = APPTSYNC_valueAtAbsoluteColumn_(row, firstColumn, headerMap.reportStatus);
      var companyName = APPTSYNC_valueAtAbsoluteColumn_(row, firstColumn, headerMap.companyName);

      if (!contractNo && !customerNo && !reportStatus && !companyName) continue;
      sourceRowCount++;

      var mappedStatus = APPTSYNC_mapReportStatus_(reportStatus);
      if (!mappedStatus) {
        ignoredRowCount++;
        continue;
      }

      mappedRowCount++;
      var match = APPTSYNC_matchMasterRow_(contractNo, customerNo, masterContext);

      if (match.status !== 'MATCHED') {
        unmatchedCount++;
        logs.push(APPTSYNC_makeLog_({
          status: match.status,
          sourceSheet: sheetName,
          sourceRow: sourceRow,
          contractNo: contractNo,
          customerNo: customerNo,
          companyName: companyName,
          reportStatus: reportStatus,
          mappedStatus: mappedStatus,
          message: match.message || '마스터 행을 특정하지 못했습니다.'
        }));
        continue;
      }

      var existing = desiredByMasterRow[match.record.row];
      var sourceInfo = {
        sourceSheet: sheetName,
        sourceRow: sourceRow,
        contractNo: contractNo,
        customerNo: customerNo,
        companyName: companyName,
        reportStatus: reportStatus,
        mappedStatus: mappedStatus,
        record: match.record
      };

      if (!existing) {
        desiredByMasterRow[match.record.row] = sourceInfo;
      } else if (existing.mappedStatus !== mappedStatus) {
        existing.conflict = true;
        existing.conflictMessage =
          existing.sourceSheet + ' ' + existing.sourceRow + '행=' + existing.mappedStatus +
          ' / ' + sheetName + ' ' + sourceRow + '행=' + mappedStatus;
        conflictCount++;
        logs.push(APPTSYNC_makeLog_({
          status: 'CONFLICTING_SOURCE_STATUS',
          sourceSheet: sheetName,
          sourceRow: sourceRow,
          contractNo: contractNo,
          customerNo: customerNo,
          companyName: companyName,
          reportStatus: reportStatus,
          mappedStatus: mappedStatus,
          record: match.record,
          message: existing.conflictMessage
        }));
      }
    }
  }

  return {
    desiredByMasterRow: desiredByMasterRow,
    logs: logs,
    sourceRowCount: sourceRowCount,
    mappedRowCount: mappedRowCount,
    ignoredRowCount: ignoredRowCount,
    unmatchedCount: unmatchedCount,
    conflictCount: conflictCount
  };
}


function APPTSYNC_applyStatuses_(sheet, masterContext, collection, dryRun) {
  var logs = [];
  var writeCount = 0;
  var unchangedCount = 0;
  var errorCount = 0;
  var rows = Object.keys(collection.desiredByMasterRow)
    .map(function (key) { return Number(key); })
    .sort(function (a, b) { return a - b; });

  for (var i = 0; i < rows.length; i++) {
    var item = collection.desiredByMasterRow[rows[i]];
    if (!item || item.conflict) continue;

    var current = String(item.record.completionStatus || '').trim();
    var next = String(item.mappedStatus || '').trim();

    if (current === next) {
      unchangedCount++;
      continue;
    }

    if (!dryRun) {
      try {
        sheet.getRange(item.record.row, masterContext.headerMap.completionStatus).setValue(next);
      } catch (err) {
        errorCount++;
        logs.push(APPTSYNC_makeLog_({
          status: 'WRITE_ERROR',
          sourceSheet: item.sourceSheet,
          sourceRow: item.sourceRow,
          contractNo: item.contractNo,
          customerNo: item.customerNo,
          companyName: item.companyName,
          reportStatus: item.reportStatus,
          mappedStatus: item.mappedStatus,
          record: item.record,
          oldValue: current,
          newValue: next,
          message: AUTOMATION_errorMessage_(err)
        }));
        continue;
      }
    }

    writeCount++;
    logs.push(APPTSYNC_makeLog_({
      status: dryRun ? 'WOULD_UPDATE' : 'UPDATED',
      sourceSheet: item.sourceSheet,
      sourceRow: item.sourceRow,
      contractNo: item.contractNo,
      customerNo: item.customerNo,
      companyName: item.companyName,
      reportStatus: item.reportStatus,
      mappedStatus: item.mappedStatus,
      record: item.record,
      oldValue: current,
      newValue: next,
      message: '선임신고 현황의 신고 여부를 마스터 선임완료여부로 변환했습니다.'
    }));
  }

  return {
    logs: logs,
    writeCount: writeCount,
    unchangedCount: unchangedCount,
    errorCount: errorCount
  };
}


/****************************************************
 * 마스터 인덱스·상태 변환
 ****************************************************/

function APPTSYNC_buildMasterContext_(sheet) {
  var headerMap = APPTSYNC_buildHeaderMap_(
    sheet,
    APPTSYNC_CONFIG.masterHeaderRow,
    APPTSYNC_CONFIG.masterHeaders
  );
  var lastRow = sheet.getLastRow();
  var rowCount = Math.max(0, lastRow - APPTSYNC_CONFIG.masterDataStartRow + 1);
  var byContract = {};
  var byCustomer = {};

  if (!rowCount) {
    return {
      headerMap: headerMap,
      byContract: byContract,
      byCustomer: byCustomer
    };
  }

  var customerValues = sheet.getRange(
    APPTSYNC_CONFIG.masterDataStartRow,
    headerMap.customerNo,
    rowCount,
    1
  ).getDisplayValues();
  var companyValues = sheet.getRange(
    APPTSYNC_CONFIG.masterDataStartRow,
    headerMap.companyName,
    rowCount,
    1
  ).getDisplayValues();
  var contractValues = sheet.getRange(
    APPTSYNC_CONFIG.masterDataStartRow,
    headerMap.contractNo,
    rowCount,
    1
  ).getDisplayValues();
  var completionValues = sheet.getRange(
    APPTSYNC_CONFIG.masterDataStartRow,
    headerMap.completionStatus,
    rowCount,
    1
  ).getDisplayValues();

  for (var i = 0; i < rowCount; i++) {
    var record = {
      row: APPTSYNC_CONFIG.masterDataStartRow + i,
      customerNo: APPTSYNC_normalizeId_(customerValues[i][0]),
      companyName: String(companyValues[i][0] || '').trim(),
      contractNo: APPTSYNC_normalizeId_(contractValues[i][0]),
      completionStatus: String(completionValues[i][0] || '').trim()
    };

    if (record.contractNo) APPTSYNC_pushIndex_(byContract, record.contractNo, record);
    if (record.customerNo) APPTSYNC_pushIndex_(byCustomer, record.customerNo, record);
  }

  return {
    headerMap: headerMap,
    byContract: byContract,
    byCustomer: byCustomer
  };
}


function APPTSYNC_matchMasterRow_(contractNo, customerNo, context) {
  var contractKey = APPTSYNC_normalizeId_(contractNo);
  var customerKey = APPTSYNC_normalizeId_(customerNo);
  var contractMatches = contractKey ? (context.byContract[contractKey] || []) : [];
  var customerMatches = customerKey ? (context.byCustomer[customerKey] || []) : [];

  if (contractMatches.length > 1) {
    return {
      status: 'DUPLICATE_CONTRACT_NUMBER',
      message: '마스터 발주번호가 중복입니다: ' + contractKey
    };
  }
  if (customerMatches.length > 1) {
    return {
      status: 'DUPLICATE_CUSTOMER_NUMBER',
      message: '마스터 고객번호가 중복입니다: ' + customerKey
    };
  }

  var byContractRecord = contractMatches.length === 1 ? contractMatches[0] : null;
  var byCustomerRecord = customerMatches.length === 1 ? customerMatches[0] : null;

  if (byContractRecord && byCustomerRecord && byContractRecord.row !== byCustomerRecord.row) {
    return {
      status: 'CONTRACT_CUSTOMER_CONFLICT',
      message: '계약번호와 고객번호가 서로 다른 마스터 행을 가리킵니다.'
    };
  }

  var record = byContractRecord || byCustomerRecord;
  if (!record) {
    return {
      status: 'MASTER_ROW_NOT_FOUND',
      message: '계약번호·고객번호로 마스터 행을 찾을 수 없습니다.'
    };
  }

  return { status: 'MATCHED', record: record };
}


function APPTSYNC_mapReportStatus_(value) {
  var normalized = String(value || '').replace(/[\s/]+/g, '');
  if (!normalized) return '';
  if (normalized.indexOf('신고수리') >= 0) return '완료';
  if (normalized.indexOf('신고접수') >= 0) return '진행중(접수)';
  if (normalized.indexOf('지자체반려') >= 0) return '진행중(반려)';
  if (normalized.indexOf('비선임') >= 0 || normalized.indexOf('자체선임') >= 0) return '비대상';
  return '';
}


/****************************************************
 * 로그·보조 함수
 ****************************************************/

function APPTSYNC_getOrCreateLogSheet_(ss) {
  var sheet = ss.getSheetByName(APPTSYNC_CONFIG.logSheetName);
  if (!sheet) sheet = ss.insertSheet(APPTSYNC_CONFIG.logSheetName);
  sheet.getRange(1, 1, 1, APPTSYNC_CONFIG.logHeaders.length)
    .setValues([APPTSYNC_CONFIG.logHeaders.slice()]);
  sheet.setFrozenRows(1);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}


function APPTSYNC_appendLogs_(ss, logs, source) {
  if (!logs || !logs.length) return;
  var sheet = APPTSYNC_getOrCreateLogSheet_(ss);
  var now = new Date();
  var rows = [];

  for (var i = 0; i < logs.length; i++) {
    var log = logs[i] || {};
    rows.push([
      now,
      source,
      String(log.status || ''),
      String(log.sourceSheet || ''),
      Number(log.sourceRow) || '',
      String(log.contractNo || ''),
      String(log.customerNo || ''),
      String(log.companyName || ''),
      String(log.reportStatus || ''),
      String(log.mappedStatus || ''),
      Number(log.masterRow) || '',
      String(log.masterCompanyName || ''),
      String(log.oldValue || ''),
      String(log.newValue || ''),
      APPTSYNC_truncate_(log.message),
      APPTSYNC_CONFIG.version
    ]);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, APPTSYNC_CONFIG.logHeaders.length)
    .setValues(rows);
}


function APPTSYNC_makeLog_(options) {
  options = options || {};
  var record = options.record || {};
  return {
    status: options.status || '',
    sourceSheet: options.sourceSheet || '',
    sourceRow: options.sourceRow || '',
    contractNo: options.contractNo || '',
    customerNo: options.customerNo || '',
    companyName: options.companyName || '',
    reportStatus: options.reportStatus || '',
    mappedStatus: options.mappedStatus || '',
    masterRow: record.row || '',
    masterCompanyName: record.companyName || '',
    oldValue: options.oldValue || '',
    newValue: options.newValue || '',
    message: options.message || ''
  };
}


function APPTSYNC_buildHeaderMap_(sheet, headerRow, requiredHeaders) {
  var headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  var normalizedMap = {};

  for (var i = 0; i < headers.length; i++) {
    var normalized = APPTSYNC_normalizeHeader_(headers[i]);
    if (normalized && !normalizedMap[normalized]) normalizedMap[normalized] = i + 1;
  }

  var result = {};
  for (var key in requiredHeaders) {
    if (!Object.prototype.hasOwnProperty.call(requiredHeaders, key)) continue;
    var expected = requiredHeaders[key];
    var column = normalizedMap[APPTSYNC_normalizeHeader_(expected)] || 0;
    if (!column) {
      throw new Error(
        sheet.getName() + ' 시트 ' + headerRow + '행에서 필수 헤더를 찾을 수 없습니다: ' + expected
      );
    }
    result[key] = column;
  }

  return result;
}


function APPTSYNC_normalizeHeader_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}


function APPTSYNC_normalizeId_(value) {
  var text = String(value || '').trim();
  if (!text) return '';
  if (/^\d+(\.0+)?$/.test(text)) return String(parseInt(text, 10));
  return text;
}


function APPTSYNC_pushIndex_(map, key, record) {
  if (!map[key]) map[key] = [];
  map[key].push(record);
}


function APPTSYNC_valueAtAbsoluteColumn_(row, firstColumn, absoluteColumn) {
  return String(row[absoluteColumn - firstColumn] || '').trim();
}


function APPTSYNC_readJson_(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (ignoreJsonError) {
    return null;
  }
}


function APPTSYNC_truncate_(value) {
  var text = String(value || '');
  return text.length <= APPTSYNC_CONFIG.maxLogTextLength
    ? text
    : text.slice(0, APPTSYNC_CONFIG.maxLogTextLength) + '…';
}


function APPTSYNC_summaryText_(result) {
  return [
    '상태: ' + String(result.status || ''),
    '소스 데이터행: ' + Number(result.sourceRows || 0),
    '상태 매핑행: ' + Number(result.mappedRows || 0),
    '미지정 상태 무시: ' + Number(result.ignoredRows || 0),
    '마스터 미매칭: ' + Number(result.unmatchedRows || 0),
    '소스 상태 충돌: ' + Number(result.conflictRows || 0),
    '반영 대상: ' + Number(result.desiredRows || 0),
    '입력/변경: ' + Number(result.writeCount || 0),
    '기존 동일값: ' + Number(result.unchangedCount || 0),
    result.dryRun ? '\n※ 미리보기이므로 데이터는 수정하지 않았습니다.' : ''
  ].join('\n');
}
