/****************************************************
 * InspectionScheduleSync.gs
 * 점검일정 → 마스터시트 점검예정일 동기화 - 17단계
 *
 * 대상:
 * - 소스: 점검일정
 * - 타깃: 마스터시트(신규)
 * - 유지점검예정일 / 성능점검예정일
 *
 * 원칙:
 * - 점검일정 1행=연도, 2행=월, 3행=일, 4행부터 일정
 * - "정보통신 유지" / "정보통신 성능" 문구가 있는 일정만 반영
 * - 동일 고객·동일 점검유형이 여러 날짜면 중복 제거 후 오름차순으로 모두 기록
 * - 회사명은 정규화·활성 계약 우선·주소 힌트로 보수적으로 매칭
 * - 기존 수기값은 이전 동기화값과 같을 때만 자동 갱신/삭제
 ****************************************************/

var INSPSYNC_CONFIG = Object.freeze({
  version: '2026-07-28-PHASE23',
  sourceSheetName: '점검일정',
  masterSheetName: '마스터시트(신규)',

  sourceYearRow: 1,
  sourceMonthRow: 2,
  sourceDayRow: 3,
  sourceDataStartRow: 4,
  sourceCalendarStartColumn: 3,

  masterHeaderRow: 2,
  masterDataStartRow: 3,
  masterHeaders: Object.freeze({
    customerNo: '고객번호',
    salesStatus: '현재 영업 진행 상황',
    companyName: '회사명',
    address: '고객사 상세 주소',
    contractNo: '발주번호',
    maintenanceDate: '유지점검예정일',
    performanceDate: '성능점검예정일'
  }),

  stateSheetName: '_점검일정동기화상태',
  logSheetName: '_점검일정동기화로그',
  pendingPropertyKey: 'INSPECTION_SCHEDULE_SYNC_PENDING_V1',
  lastRunPropertyKey: 'INSPECTION_SCHEDULE_SYNC_LAST_RUN_V1',
  pipelineIntervalMs: 15 * 60 * 1000,

  matchMinimumScore: 750,
  ambiguousScoreGap: 20,
  maxLogTextLength: 1000,

  typeMaintenance: 'MAINTENANCE',
  typePerformance: 'PERFORMANCE',

  aliases: Object.freeze({
    '태화지엔지': '태화지앤지',
    '오타르상가b동': '오타브상가b동',
    '대전열병합발전소': '대전열병합발전'
  }),

  stateHeaders: Object.freeze([
    '상태키', '고객번호', '발주번호', '회사명', '점검유형',
    '동기화날짜', '마스터행', '소스셀', '수정일시', '버전'
  ]),

  logHeaders: Object.freeze([
    '기록일시', '실행경로', '상태', '소스셀', '일정일자',
    '원문', '추출고객명', '고객번호', '발주번호', '마스터행',
    '회사명', '점검유형', '기존값', '신규값', '메시지', '버전'
  ])
});


/****************************************************
 * 공개 함수
 ****************************************************/

function INSPSYNC_previewInspectionScheduleSync() {
  var result = INSPSYNC_syncAll_({
    dryRun: true,
    source: 'MANUAL_PREVIEW'
  });

  INSPSYNC_safeNotify_(
    '점검일정 동기화 미리보기',
    INSPSYNC_summaryText_(result),
    10
  );

  return result;
}


function INSPSYNC_syncInspectionScheduleNow() {
  var preview = INSPSYNC_syncAll_({
    dryRun: true,
    source: 'MANUAL_PREVIEW_BEFORE_WRITE'
  });

  var confirmation = INSPSYNC_confirmWrite_(
    '점검일정 일괄 반영',
    INSPSYNC_summaryText_(preview) +
      '\n\n실제 마스터시트의 점검예정일을 반영하시겠습니까?'
  );

  if (!confirmation.confirmed) {
    Logger.log('[INSPSYNC_syncInspectionScheduleNow] 사용자 취소');
    return preview;
  }

  if (confirmation.uiUnavailable) {
    Logger.log(
      '[INSPSYNC_syncInspectionScheduleNow] UI 없는 실행 컨텍스트입니다. ' +
      '명시적 실행 함수이므로 확인창 없이 실제 반영을 계속합니다.'
    );
  }

  var result = INSPSYNC_syncAll_({
    dryRun: false,
    source: confirmation.uiUnavailable
      ? 'MANUAL_FULL_SYNC_NO_UI'
      : 'MANUAL_FULL_SYNC'
  });

  INSPSYNC_safeNotify_(
    '점검일정 일괄 반영 완료',
    INSPSYNC_summaryText_(result),
    10
  );

  return result;
}


/**
 * UI가 있는 컨텍스트에서는 확인창을 띄우고,
 * Apps Script 편집기·시간 트리거처럼 UI가 없는 컨텍스트에서는
 * 명시적으로 호출한 실행 함수라는 점을 근거로 계속 진행합니다.
 */
function INSPSYNC_confirmWrite_(title, message) {
  try {
    var ui = SpreadsheetApp.getUi();
    var answer = ui.alert(
      String(title || '점검일정 동기화'),
      String(message || ''),
      ui.ButtonSet.YES_NO
    );

    return {
      confirmed: answer === ui.Button.YES,
      uiUnavailable: false
    };
  } catch (uiError) {
    Logger.log(
      '[INSPSYNC_confirmWrite_] UI 사용 불가: ' +
      (uiError && uiError.message ? uiError.message : String(uiError))
    );

    return {
      confirmed: true,
      uiUnavailable: true,
      error: uiError && uiError.message ? uiError.message : String(uiError)
    };
  }
}


/**
 * UI alert가 가능한 경우 alert를 사용하고, 그렇지 않으면
 * 영업관리대장 toast → Logger 순서로 결과를 남깁니다.
 */
function INSPSYNC_safeNotify_(title, message, timeoutSeconds) {
  var safeTitle = String(title || '점검일정 동기화');
  var safeMessage = String(message || '');

  try {
    var ui = SpreadsheetApp.getUi();
    ui.alert(safeTitle, safeMessage, ui.ButtonSet.OK);
    return { channel: 'UI_ALERT' };
  } catch (uiError) {
    Logger.log(
      '[INSPSYNC_safeNotify_] UI 사용 불가: ' +
      (uiError && uiError.message ? uiError.message : String(uiError))
    );
  }

  try {
    var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
    if (ss && typeof ss.toast === 'function') {
      ss.toast(safeMessage, safeTitle, Number(timeoutSeconds || 8));
      return { channel: 'TOAST' };
    }
  } catch (toastError) {
    Logger.log(
      '[INSPSYNC_safeNotify_] toast 사용 불가: ' +
      (toastError && toastError.message ? toastError.message : String(toastError))
    );
  }

  Logger.log('[' + safeTitle + ']\n' + safeMessage);
  return { channel: 'LOGGER' };
}


function INSPSYNC_showInspectionScheduleSyncLog() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = INSPSYNC_getOrCreateLogSheet_(ss);
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


function INSPSYNC_showInspectionScheduleSyncState() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = INSPSYNC_getOrCreateStateSheet_(ss);
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


/** 설치형 중앙 onEdit에서 호출 */
function INSPSYNC_handleScheduleEdit(e) {
  if (!e || !e.range) return { status: 'IGNORED_INVALID_EVENT' };

  var sheet = e.range.getSheet();
  if (String(sheet.getName() || '') !== INSPSYNC_CONFIG.sourceSheetName) {
    return { status: 'IGNORED_UNRELATED_SHEET' };
  }

  var lastColumn = typeof e.range.getLastColumn === 'function'
    ? e.range.getLastColumn()
    : e.range.getColumn() + e.range.getNumColumns() - 1;
  var lastRow = typeof e.range.getLastRow === 'function'
    ? e.range.getLastRow()
    : e.range.getRow() + e.range.getNumRows() - 1;

  if (lastColumn < INSPSYNC_CONFIG.sourceCalendarStartColumn) {
    return { status: 'IGNORED_NON_CALENDAR_COLUMN' };
  }

  if (lastRow < INSPSYNC_CONFIG.sourceYearRow) {
    return { status: 'IGNORED_NON_DATA_ROW' };
  }

  PropertiesService.getScriptProperties().setProperty(
    INSPSYNC_CONFIG.pendingPropertyKey,
    JSON.stringify({
      requestedAt: new Date().toISOString(),
      source: 'ON_EDIT',
      range: e.range.getA1Notation()
    })
  );

  try {
    return INSPSYNC_syncAll_({
      dryRun: false,
      source: 'ON_EDIT',
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
function INSPSYNC_runPipelineSafetySync_() {
  var props = PropertiesService.getScriptProperties();
  var pending = String(props.getProperty(INSPSYNC_CONFIG.pendingPropertyKey) || '');
  var lastRun = INSPSYNC_readJson_(props.getProperty(INSPSYNC_CONFIG.lastRunPropertyKey));
  var lastFinishedMs = Date.parse(String(lastRun && lastRun.finishedAt || ''));
  var due = !!pending || !isFinite(lastFinishedMs) ||
    Date.now() - lastFinishedMs >= INSPSYNC_CONFIG.pipelineIntervalMs;

  if (!due) {
    return {
      status: 'SKIPPED_NOT_DUE',
      pending: false,
      lastFinishedAt: lastRun && lastRun.finishedAt || ''
    };
  }

  try {
    return INSPSYNC_syncAll_({
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

function INSPSYNC_syncAll_(options) {
  options = options || {};
  var dryRun = options.dryRun === true;
  var source = String(options.source || 'UNKNOWN');

  return AUTOMATION_runWithModuleLeaseOrThrow_(
    'INSPECTION_SCHEDULE_SYNC',
    'INSPSYNC_syncAll_',
    function () {
      var startedAtMs = Date.now();
      var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
      var sourceSheet = ss.getSheetByName(INSPSYNC_CONFIG.sourceSheetName);
      var masterSheet = ss.getSheetByName(INSPSYNC_CONFIG.masterSheetName);

      if (!sourceSheet) throw new Error('점검일정 시트를 찾을 수 없습니다.');
      if (!masterSheet) throw new Error('마스터시트(신규)를 찾을 수 없습니다.');

      var masterContext = INSPSYNC_buildMasterContext_(masterSheet);
      var parsed = INSPSYNC_parseSchedule_(sourceSheet, masterContext);
      var previousState = INSPSYNC_readState_(ss);
      var applyResult = INSPSYNC_planAndApply_(
        masterSheet,
        masterContext,
        parsed,
        previousState,
        dryRun
      );

      var finishedAt = new Date().toISOString();
      var result = {
        version: INSPSYNC_CONFIG.version,
        status: applyResult.errorCount > 0
          ? 'COMPLETED_WITH_ERRORS'
          : (applyResult.conflictCount > 0 ? 'COMPLETED_WITH_CONFLICTS' : 'SUCCESS'),
        dryRun: dryRun,
        source: source,
        startedAt: new Date(startedAtMs).toISOString(),
        finishedAt: finishedAt,
        durationMs: Date.now() - startedAtMs,
        scheduleCells: parsed.scheduleCellCount,
        matchedCells: parsed.matchedCellCount,
        unmatchedCells: parsed.unmatched.length,
        ambiguousCells: parsed.ambiguous.length,
        desiredItems: Object.keys(parsed.desiredByKey).length,
        writeCount: applyResult.writeCount,
        clearCount: applyResult.clearCount,
        unchangedCount: applyResult.unchangedCount,
        conflictCount: applyResult.conflictCount,
        errorCount: applyResult.errorCount,
        logCount: applyResult.logs.length + parsed.logs.length
      };

      if (!dryRun) {
        INSPSYNC_writeState_(ss, applyResult.nextStateRows);

        var logsToWrite = parsed.logs.concat(applyResult.logs);
        if (source.indexOf('PIPELINE_') === 0) {
          logsToWrite = logsToWrite.filter(function (log) {
            var status = String(log && log.status || '');
            return status === 'UPDATED' ||
              status === 'CLEARED_REMOVED_SCHEDULE' ||
              status === 'WRITE_ERROR' ||
              status === 'INVALID_DATE_HEADER' ||
              status === 'STATE_TARGET_NOT_FOUND';
          });
        }
        INSPSYNC_appendLogs_(ss, logsToWrite, source);

        var props = PropertiesService.getScriptProperties();
        props.deleteProperty(INSPSYNC_CONFIG.pendingPropertyKey);
        props.setProperty(INSPSYNC_CONFIG.lastRunPropertyKey, JSON.stringify(result));
      }

      return result;
    },
    {
      ttlMs: 3 * 60 * 1000,
      waitMs: 300,
      taskName: '점검일정 → 마스터 점검예정일 동기화'
    }
  );
}


/****************************************************
 * 점검일정 파싱
 ****************************************************/

function INSPSYNC_parseSchedule_(sheet, masterContext) {
  var lastRow = Math.max(INSPSYNC_CONFIG.sourceDataStartRow, sheet.getLastRow());
  var lastColumn = Math.max(INSPSYNC_CONFIG.sourceCalendarStartColumn, sheet.getLastColumn());
  var values = sheet.getRange(1, 1, lastRow, lastColumn).getDisplayValues();
  var datesByColumn = INSPSYNC_buildCalendarDates_(values, lastColumn);
  var desiredByKey = {};
  var logs = [];
  var unmatched = [];
  var ambiguous = [];
  var scheduleCellCount = 0;
  var matchedCellCount = 0;

  for (var row = INSPSYNC_CONFIG.sourceDataStartRow; row <= lastRow; row++) {
    for (var col = INSPSYNC_CONFIG.sourceCalendarStartColumn; col <= lastColumn; col++) {
      var raw = String(values[row - 1][col - 1] || '').trim();
      if (!raw) continue;

      var inspectionTypes = INSPSYNC_detectInspectionTypes_(raw);
      if (!inspectionTypes.length) continue;

      var scheduleDate = datesByColumn[col];
      var sourceCell = sheet.getRange(row, col).getA1Notation();
      scheduleCellCount++;

      if (!scheduleDate) {
        logs.push(INSPSYNC_makeLog_({
          status: 'INVALID_DATE_HEADER',
          sourceCell: sourceCell,
          raw: raw,
          message: '연도·월·일 헤더를 날짜로 해석하지 못했습니다.'
        }));
        continue;
      }

      var candidates = INSPSYNC_extractCustomerNameCandidates_(raw);
      var matchedRecords = [];
      var seenRows = {};
      var cellHadAmbiguousMatch = false;

      for (var ci = 0; ci < candidates.length; ci++) {
        var candidate = candidates[ci];
        var match = INSPSYNC_matchCustomer_(candidate, masterContext.records);

        if (match.status === 'MATCHED') {
          if (!seenRows[match.record.row]) {
            seenRows[match.record.row] = true;
            matchedRecords.push({
              candidate: candidate,
              record: match.record,
              score: match.score
            });
          }

          if (ci === 0) {
            break;
          }
        } else if (match.status === 'AMBIGUOUS') {
          cellHadAmbiguousMatch = true;
          ambiguous.push({
            sourceCell: sourceCell,
            raw: raw,
            candidate: candidate,
            matches: match.matches
          });
        }
      }

      if (!matchedRecords.length) {
        var unmatchedRow = INSPSYNC_makeLog_({
          status: cellHadAmbiguousMatch ? 'AMBIGUOUS_CUSTOMER' : 'UNMATCHED_CUSTOMER',
          sourceCell: sourceCell,
          scheduleDate: scheduleDate,
          raw: raw,
          extractedName: candidates.join(' | '),
          message: '마스터시트 고객을 안전하게 특정하지 못했습니다.'
        });
        logs.push(unmatchedRow);
        unmatched.push(unmatchedRow);
        continue;
      }

      matchedCellCount++;

      for (var mi = 0; mi < matchedRecords.length; mi++) {
        var matched = matchedRecords[mi];

        for (var ti = 0; ti < inspectionTypes.length; ti++) {
          var type = inspectionTypes[ti];
          var key = INSPSYNC_stateKey_(matched.record, type);
          var currentDesired = desiredByKey[key];
          var sourceRef = sourceCell + ':' + matched.candidate;

          if (!currentDesired || scheduleDate.getTime() < currentDesired.date.getTime()) {
            desiredByKey[key] = {
              key: key,
              type: type,
              date: scheduleDate,
              record: matched.record,
              sourceRefs: [sourceRef],
              raw: raw,
              extractedName: matched.candidate
            };
          } else if (scheduleDate.getTime() === currentDesired.date.getTime()) {
            if (currentDesired.sourceRefs.indexOf(sourceRef) < 0) {
              currentDesired.sourceRefs.push(sourceRef);
            }
          }
        }
      }
    }
  }

  return {
    desiredByKey: desiredByKey,
    logs: logs,
    unmatched: unmatched,
    ambiguous: ambiguous,
    scheduleCellCount: scheduleCellCount,
    matchedCellCount: matchedCellCount
  };
}


function INSPSYNC_buildCalendarDates_(values, lastColumn) {
  var result = {};
  var currentYear = 0;
  var currentMonth = 0;
  var previousDay = 0;

  for (var col = INSPSYNC_CONFIG.sourceCalendarStartColumn; col <= lastColumn; col++) {
    var yearText = String(values[INSPSYNC_CONFIG.sourceYearRow - 1][col - 1] || '');
    var monthText = String(values[INSPSYNC_CONFIG.sourceMonthRow - 1][col - 1] || '');
    var dayText = String(values[INSPSYNC_CONFIG.sourceDayRow - 1][col - 1] || '');

    var parsedYear = INSPSYNC_parseNumber_(yearText);
    var parsedMonth = INSPSYNC_parseNumber_(monthText);
    var parsedDay = INSPSYNC_parseNumber_(dayText);

    if (parsedYear >= 2000 && parsedYear <= 2100) currentYear = parsedYear;
    if (parsedMonth >= 1 && parsedMonth <= 12) currentMonth = parsedMonth;

    if (
      !parsedMonth && currentMonth && previousDay && parsedDay &&
      parsedDay < previousDay
    ) {
      currentMonth++;
      if (currentMonth > 12) {
        currentMonth = 1;
        currentYear++;
      }
    }

    if (currentYear && currentMonth && parsedDay >= 1 && parsedDay <= 31) {
      var date = new Date(currentYear, currentMonth - 1, parsedDay);
      if (
        date.getFullYear() === currentYear &&
        date.getMonth() === currentMonth - 1 &&
        date.getDate() === parsedDay
      ) {
        result[col] = date;
      }
    }

    if (parsedDay) previousDay = parsedDay;
  }

  return result;
}


function INSPSYNC_detectInspectionTypes_(raw) {
  var compact = String(raw || '').replace(/\s+/g, '');
  if (compact.indexOf('정보통신') < 0) return [];

  var types = [];
  if (compact.indexOf('유지') >= 0) types.push(INSPSYNC_CONFIG.typeMaintenance);
  if (compact.indexOf('성능') >= 0) types.push(INSPSYNC_CONFIG.typePerformance);
  return types;
}


function INSPSYNC_extractCustomerNameCandidates_(raw) {
  var lines = String(raw || '').replace(/\r/g, '\n').split(/\n+/);
  var cleaned = [];

  for (var i = 0; i < lines.length; i++) {
    var line = String(lines[i] || '').trim();
    if (!line) continue;
    if (line.indexOf('현황파악') >= 0) continue;

    line = line.replace(/\([^)]*정보통신[^)]*\)/g, ' ');
    line = line.replace(/정보통신\s*(유지|성능)(\s*\+\s*(유지|성능))?/g, ' ');
    line = line.replace(/^\s*(오전|오후)\s*/g, '');
    line = line.replace(/\s+/g, ' ').trim();

    if (line) cleaned.push(line);
  }

  var candidates = [];
  if (cleaned.length) candidates.push(cleaned.join(' '));

  for (var j = 0; j < cleaned.length; j++) {
    candidates.push(cleaned[j]);
    if (j + 1 < cleaned.length) candidates.push(cleaned[j] + ' ' + cleaned[j + 1]);
  }

  var unique = [];
  var seen = {};
  for (var k = 0; k < candidates.length; k++) {
    var value = String(candidates[k] || '').trim();
    if (!value) continue;
    var normalized = INSPSYNC_normalizeName_(value);
    if (!normalized || seen[normalized]) continue;
    seen[normalized] = true;
    unique.push(value);
  }

  return unique;
}


/****************************************************
 * 마스터 인덱스·매칭
 ****************************************************/

function INSPSYNC_buildMasterContext_(sheet) {
  var headerMap = INSPSYNC_buildHeaderMap_(
    sheet,
    INSPSYNC_CONFIG.masterHeaderRow,
    INSPSYNC_CONFIG.masterHeaders
  );
  var lastRow = sheet.getLastRow();
  var rowCount = Math.max(0, lastRow - INSPSYNC_CONFIG.masterDataStartRow + 1);
  var records = [];

  if (!rowCount) {
    return {
      headerMap: headerMap,
      records: records,
      lastRow: lastRow,
      rowCount: 0
    };
  }

  var customerValues = sheet.getRange(
    INSPSYNC_CONFIG.masterDataStartRow,
    headerMap.customerNo,
    rowCount,
    1
  ).getDisplayValues();
  var statusValues = sheet.getRange(
    INSPSYNC_CONFIG.masterDataStartRow,
    headerMap.salesStatus,
    rowCount,
    1
  ).getDisplayValues();
  var companyValues = sheet.getRange(
    INSPSYNC_CONFIG.masterDataStartRow,
    headerMap.companyName,
    rowCount,
    1
  ).getDisplayValues();
  var addressValues = sheet.getRange(
    INSPSYNC_CONFIG.masterDataStartRow,
    headerMap.address,
    rowCount,
    1
  ).getDisplayValues();
  var contractValues = sheet.getRange(
    INSPSYNC_CONFIG.masterDataStartRow,
    headerMap.contractNo,
    rowCount,
    1
  ).getDisplayValues();
  var dateValues = sheet.getRange(
    INSPSYNC_CONFIG.masterDataStartRow,
    headerMap.maintenanceDate,
    rowCount,
    2
  ).getValues();

  for (var i = 0; i < rowCount; i++) {
    var companyName = String(companyValues[i][0] || '').trim();
    if (!companyName) continue;

    records.push({
      row: INSPSYNC_CONFIG.masterDataStartRow + i,
      customerNo: String(customerValues[i][0] || '').trim(),
      salesStatus: String(statusValues[i][0] || '').trim(),
      companyName: companyName,
      companyNorm: INSPSYNC_normalizeName_(companyName),
      address: String(addressValues[i][0] || '').trim(),
      addressNorm: INSPSYNC_normalizeName_(addressValues[i][0] || ''),
      contractNo: String(contractValues[i][0] || '').trim(),
      maintenanceDate: dateValues[i][0],
      performanceDate: dateValues[i][1]
    });
  }

  return {
    headerMap: headerMap,
    records: records,
    lastRow: lastRow,
    rowCount: rowCount
  };
}


function INSPSYNC_matchCustomer_(candidateName, records) {
  var candidateNorm = INSPSYNC_applyAlias_(INSPSYNC_normalizeName_(candidateName));
  if (!candidateNorm || candidateNorm.length < 2) return { status: 'UNMATCHED' };

  var scored = [];

  for (var i = 0; i < records.length; i++) {
    var record = records[i];
    var masterNorm = INSPSYNC_applyAlias_(record.companyNorm);
    var score = INSPSYNC_customerMatchScore_(candidateNorm, masterNorm, record);

    if (score >= INSPSYNC_CONFIG.matchMinimumScore) {
      scored.push({ score: score, record: record });
    }
  }

  scored.sort(function (a, b) {
    if (b.score !== a.score) return b.score - a.score;
    return a.record.row - b.record.row;
  });

  if (!scored.length) return { status: 'UNMATCHED' };

  if (
    scored.length > 1 &&
    scored[0].record.row !== scored[1].record.row &&
    scored[0].score - scored[1].score <= INSPSYNC_CONFIG.ambiguousScoreGap
  ) {
    return {
      status: 'AMBIGUOUS',
      matches: scored.slice(0, 3).map(function (item) {
        return item.record.companyName + '(행 ' + item.record.row + ', 점수 ' + item.score + ')';
      })
    };
  }

  return {
    status: 'MATCHED',
    record: scored[0].record,
    score: scored[0].score
  };
}


function INSPSYNC_customerMatchScore_(candidateNorm, masterNorm, record) {
  var score = -9999;

  if (candidateNorm === masterNorm) {
    score = 1000;
  } else if (
    candidateNorm.length >= 4 && masterNorm.length >= 4 &&
    (candidateNorm.indexOf(masterNorm) >= 0 || masterNorm.indexOf(candidateNorm) >= 0)
  ) {
    var ratio = Math.min(candidateNorm.length, masterNorm.length) /
      Math.max(candidateNorm.length, masterNorm.length);
    score = 800 + Math.round(ratio * 100);

    if (candidateNorm.indexOf(masterNorm) >= 0) {
      var leftover = candidateNorm.replace(masterNorm, '');
      if (leftover.length >= 2 && record.addressNorm.indexOf(leftover) >= 0) {
        score += 180;
      }
    }
  } else {
    return score;
  }

  if (record.contractNo) score += 80;
  if (record.salesStatus === '계약완료' || record.salesStatus === '발주완료') score += 40;
  if (record.salesStatus === '수주실패') score -= 180;
  if (record.salesStatus.indexOf('장기') >= 0) score -= 100;

  return score;
}


function INSPSYNC_normalizeName_(value) {
  var text = String(value || '');
  try {
    text = text.normalize('NFKC');
  } catch (ignoreNormalizeError) {}

  return text
    .toLowerCase()
    .replace(/주식회사/g, '')
    .replace(/유한회사/g, '')
    .replace(/\(\s*주\s*\)|㈜/g, '')
    .replace(/[^0-9a-z가-힣]/g, '');
}


function INSPSYNC_applyAlias_(normalizedName) {
  return INSPSYNC_CONFIG.aliases[normalizedName] || normalizedName;
}


/****************************************************
 * 쓰기 계획·상태 보호
 ****************************************************/

function INSPSYNC_planAndApply_(sheet, masterContext, parsed, previousState, dryRun) {
  var desired = parsed.desiredByKey;
  var keys = {};
  var key;

  for (key in desired) {
    if (Object.prototype.hasOwnProperty.call(desired, key)) keys[key] = true;
  }
  for (key in previousState) {
    if (Object.prototype.hasOwnProperty.call(previousState, key)) keys[key] = true;
  }

  var writes = [];
  var logs = [];
  var nextStateRows = [];
  var writeCount = 0;
  var clearCount = 0;
  var unchangedCount = 0;
  var conflictCount = 0;
  var errorCount = 0;

  for (key in keys) {
    if (!Object.prototype.hasOwnProperty.call(keys, key)) continue;

    var item = desired[key] || null;
    var oldState = previousState[key] || null;
    var record = item ? item.record : INSPSYNC_findRecordForState_(oldState, masterContext.records);

    if (!record) {
      logs.push(INSPSYNC_makeLog_({
        status: 'STATE_TARGET_NOT_FOUND',
        message: '이전 동기화 상태의 마스터 고객을 찾지 못했습니다.',
        extractedName: oldState && oldState.companyName || ''
      }));
      errorCount++;
      continue;
    }

    var type = item ? item.type : String(oldState.type || '');
    var targetColumn = type === INSPSYNC_CONFIG.typePerformance
      ? masterContext.headerMap.performanceDate
      : masterContext.headerMap.maintenanceDate;
    var currentValue = type === INSPSYNC_CONFIG.typePerformance
      ? record.performanceDate
      : record.maintenanceDate;
    var currentKey = INSPSYNC_dateKey_(currentValue);
    var previousKey = oldState ? String(oldState.dateKey || '') : '';
    var desiredKey = item ? INSPSYNC_dateKey_(item.date) : '';
    var sourceRefs = item ? item.sourceRefs.join(', ') : (oldState && oldState.sourceRefs || '');

    if (item) {
      if (!currentKey || currentKey === desiredKey || (previousKey && currentKey === previousKey)) {
        if (currentKey !== desiredKey) {
          writes.push({ row: record.row, column: targetColumn, value: item.date });
          writeCount++;
          logs.push(INSPSYNC_makeLog_({
            status: dryRun ? 'WOULD_UPDATE' : 'UPDATED',
            sourceCell: sourceRefs,
            scheduleDate: item.date,
            raw: item.raw,
            extractedName: item.extractedName,
            record: record,
            type: type,
            oldValue: currentValue,
            newValue: item.date,
            message: '점검일정의 가장 빠른 예정일을 반영합니다.'
          }));
        } else {
          unchangedCount++;
        }

        nextStateRows.push(INSPSYNC_stateRow_(key, record, type, desiredKey, sourceRefs));
      } else {
        conflictCount++;
        logs.push(INSPSYNC_makeLog_({
          status: 'PRESERVED_MANUAL_CONFLICT',
          sourceCell: sourceRefs,
          scheduleDate: item.date,
          raw: item.raw,
          extractedName: item.extractedName,
          record: record,
          type: type,
          oldValue: currentValue,
          newValue: item.date,
          message: '마스터 현재값이 이전 자동동기화값과 달라 수기값으로 보고 보존했습니다.'
        }));
      }
    } else if (oldState) {
      if (currentKey && currentKey === previousKey) {
        writes.push({ row: record.row, column: targetColumn, value: '' });
        clearCount++;
        logs.push(INSPSYNC_makeLog_({
          status: dryRun ? 'WOULD_CLEAR' : 'CLEARED_REMOVED_SCHEDULE',
          sourceCell: sourceRefs,
          record: record,
          type: type,
          oldValue: currentValue,
          newValue: '',
          message: '점검일정에서 삭제된 기존 자동동기화값을 비웁니다.'
        }));
      } else if (currentKey) {
        conflictCount++;
        logs.push(INSPSYNC_makeLog_({
          status: 'PRESERVED_MANUAL_AFTER_REMOVAL',
          sourceCell: sourceRefs,
          record: record,
          type: type,
          oldValue: currentValue,
          newValue: '',
          message: '점검일정에서는 사라졌지만 마스터값이 수기 변경되어 보존했습니다.'
        }));
      }
    }
  }

  if (!dryRun) {
    for (var wi = 0; wi < writes.length; wi++) {
      var write = writes[wi];
      var targetRange = sheet.getRange(write.row, write.column);
      targetRange.setValue(write.value);
      if (write.value instanceof Date) {
        targetRange.setNumberFormat('yyyy. mm. dd.');
      }
    }
  }

  return {
    writeCount: writeCount,
    clearCount: clearCount,
    unchangedCount: unchangedCount,
    conflictCount: conflictCount,
    errorCount: errorCount,
    logs: logs,
    nextStateRows: nextStateRows
  };
}


function INSPSYNC_findRecordForState_(state, records) {
  if (!state) return null;

  for (var i = 0; i < records.length; i++) {
    if (state.customerNo && records[i].customerNo === state.customerNo) return records[i];
  }

  for (var j = 0; j < records.length; j++) {
    if (state.contractNo && records[j].contractNo === state.contractNo) return records[j];
  }

  return null;
}


function INSPSYNC_stateKey_(record, type) {
  var identity = record.customerNo
    ? 'CUSTOMER:' + record.customerNo
    : (record.contractNo ? 'CONTRACT:' + record.contractNo : 'ROW:' + record.row);
  return identity + '|TYPE:' + type;
}


function INSPSYNC_stateRow_(key, record, type, dateKey, sourceRefs) {
  return [
    key,
    record.customerNo,
    record.contractNo,
    record.companyName,
    type,
    dateKey,
    record.row,
    sourceRefs,
    new Date(),
    INSPSYNC_CONFIG.version
  ];
}


/****************************************************
 * 상태·로그 시트
 ****************************************************/

function INSPSYNC_readState_(ss) {
  var sheet = ss.getSheetByName(INSPSYNC_CONFIG.stateSheetName);
  var map = {};
  if (!sheet || sheet.getLastRow() < 2) return map;

  var values = sheet.getRange(2, 1, sheet.getLastRow() - 1, INSPSYNC_CONFIG.stateHeaders.length)
    .getDisplayValues();

  for (var i = 0; i < values.length; i++) {
    var key = String(values[i][0] || '').trim();
    if (!key) continue;
    map[key] = {
      key: key,
      customerNo: String(values[i][1] || '').trim(),
      contractNo: String(values[i][2] || '').trim(),
      companyName: String(values[i][3] || '').trim(),
      type: String(values[i][4] || '').trim(),
      dateKey: String(values[i][5] || '').trim(),
      masterRow: Number(values[i][6]) || 0,
      sourceRefs: String(values[i][7] || '').trim()
    };
  }

  return map;
}


function INSPSYNC_writeState_(ss, rows) {
  var sheet = INSPSYNC_getOrCreateStateSheet_(ss);
  var lastRow = sheet.getLastRow();
  if (lastRow > 1) sheet.getRange(2, 1, lastRow - 1, INSPSYNC_CONFIG.stateHeaders.length).clearContent();

  if (rows.length) {
    rows.sort(function (a, b) {
      return String(a[0]).localeCompare(String(b[0]));
    });
    sheet.getRange(2, 1, rows.length, INSPSYNC_CONFIG.stateHeaders.length).setValues(rows);
    sheet.getRange(2, 6, rows.length, 1).setNumberFormat('@');
    sheet.getRange(2, 9, rows.length, 1).setNumberFormat('yyyy-mm-dd hh:mm:ss');
  }
}


function INSPSYNC_getOrCreateStateSheet_(ss) {
  ss = SYSTEMLOG_getSpreadsheet_();
  var sheet = ss.getSheetByName(INSPSYNC_CONFIG.stateSheetName);
  if (!sheet) sheet = ss.insertSheet(INSPSYNC_CONFIG.stateSheetName);

  sheet.getRange(1, 1, 1, INSPSYNC_CONFIG.stateHeaders.length)
    .setValues([INSPSYNC_CONFIG.stateHeaders.slice()]);
  sheet.setFrozenRows(1);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}


function INSPSYNC_getOrCreateLogSheet_(ss) {
  ss = SYSTEMLOG_getSpreadsheet_();
  var sheet = ss.getSheetByName(INSPSYNC_CONFIG.logSheetName);
  if (!sheet) sheet = ss.insertSheet(INSPSYNC_CONFIG.logSheetName);

  sheet.getRange(1, 1, 1, INSPSYNC_CONFIG.logHeaders.length)
    .setValues([INSPSYNC_CONFIG.logHeaders.slice()]);
  sheet.setFrozenRows(1);
  if (!sheet.isSheetHidden()) sheet.hideSheet();
  return sheet;
}


function INSPSYNC_appendLogs_(ss, logs, source) {
  if (!logs || !logs.length) return;
  var sheet = INSPSYNC_getOrCreateLogSheet_(ss);
  var rows = [];
  var now = new Date();

  for (var i = 0; i < logs.length; i++) {
    var log = logs[i] || {};
    rows.push([
      now,
      source,
      String(log.status || ''),
      String(log.sourceCell || ''),
      log.scheduleDate || '',
      INSPSYNC_truncate_(log.raw),
      INSPSYNC_truncate_(log.extractedName),
      String(log.customerNo || ''),
      String(log.contractNo || ''),
      Number(log.masterRow) || '',
      String(log.companyName || ''),
      String(log.type || ''),
      INSPSYNC_displayDate_(log.oldValue),
      INSPSYNC_displayDate_(log.newValue),
      INSPSYNC_truncate_(log.message),
      INSPSYNC_CONFIG.version
    ]);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, INSPSYNC_CONFIG.logHeaders.length)
    .setValues(rows);
}


function INSPSYNC_makeLog_(options) {
  options = options || {};
  var record = options.record || {};
  return {
    status: options.status || '',
    sourceCell: options.sourceCell || '',
    scheduleDate: options.scheduleDate || '',
    raw: options.raw || '',
    extractedName: options.extractedName || '',
    customerNo: record.customerNo || '',
    contractNo: record.contractNo || '',
    masterRow: record.row || '',
    companyName: record.companyName || '',
    type: options.type || '',
    oldValue: options.oldValue || '',
    newValue: options.newValue || '',
    message: options.message || ''
  };
}


/****************************************************
 * 공통 보조 함수
 ****************************************************/

function INSPSYNC_buildHeaderMap_(sheet, headerRow, requiredHeaders) {
  var lastColumn = sheet.getLastColumn();
  var headers = sheet.getRange(headerRow, 1, 1, lastColumn).getDisplayValues()[0];
  var normalizedMap = {};

  for (var i = 0; i < headers.length; i++) {
    var normalized = INSPSYNC_normalizeHeader_(headers[i]);
    if (normalized && !normalizedMap[normalized]) normalizedMap[normalized] = i + 1;
  }

  var result = {};
  for (var key in requiredHeaders) {
    if (!Object.prototype.hasOwnProperty.call(requiredHeaders, key)) continue;
    var targetHeader = requiredHeaders[key];
    var column = normalizedMap[INSPSYNC_normalizeHeader_(targetHeader)] || 0;
    if (!column) {
      throw new Error(
        sheet.getName() + ' 시트 ' + headerRow + '행에서 필수 헤더를 찾을 수 없습니다: ' + targetHeader
      );
    }
    result[key] = column;
  }

  return result;
}


function INSPSYNC_normalizeHeader_(value) {
  return String(value || '').replace(/\s+/g, '').trim();
}


function INSPSYNC_parseNumber_(value) {
  var match = String(value || '').match(/\d{1,4}/);
  return match ? Number(match[0]) : 0;
}


function INSPSYNC_dateKey_(value) {
  if (!value) return '';
  var date = value instanceof Date ? value : new Date(value);
  if (!(date instanceof Date) || isNaN(date.getTime())) return '';
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
}


function INSPSYNC_displayDate_(value) {
  var key = INSPSYNC_dateKey_(value);
  return key || String(value || '');
}


function INSPSYNC_readJson_(text) {
  if (!text) return null;
  try {
    return JSON.parse(text);
  } catch (ignoreJsonError) {
    return null;
  }
}


function INSPSYNC_truncate_(value) {
  var text = String(value || '');
  return text.length <= INSPSYNC_CONFIG.maxLogTextLength
    ? text
    : text.slice(0, INSPSYNC_CONFIG.maxLogTextLength) + '…';
}


function INSPSYNC_summaryText_(result) {
  return [
    '상태: ' + String(result.status || ''),
    '점검일정 셀: ' + Number(result.scheduleCells || 0),
    '매칭 셀: ' + Number(result.matchedCells || 0),
    '미매칭: ' + Number(result.unmatchedCells || 0),
    '모호한 매칭: ' + Number(result.ambiguousCells || 0),
    '반영 대상: ' + Number(result.desiredItems || 0),
    '입력/변경: ' + Number(result.writeCount || 0),
    '삭제 반영: ' + Number(result.clearCount || 0),
    '기존 동일값: ' + Number(result.unchangedCount || 0),
    '수기값 충돌 보존: ' + Number(result.conflictCount || 0),
    result.dryRun ? '\n※ 미리보기이므로 데이터는 수정하지 않았습니다.' : ''
  ].join('\n');
}
