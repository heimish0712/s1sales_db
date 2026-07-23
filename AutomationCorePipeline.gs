/****************************************************
 * AutomationCorePipeline.gs
 * 영업관리대장 5분 핵심 데이터 동기화 파이프라인 - 3단계
 *
 * 실행 순서:
 * 1) 마스터시트(신규) → 수주확정/계약완료
 * 2) 수주확정/계약완료 → KJ·일신 수행사 고객관리
 * 3) 수주확정/계약완료 신규 계약 → 정보통신유지보수 파일
 *
 * 실패 정책:
 * - 1단계 실패: 2·3단계 모두 중단
 * - 2단계 실패: 3단계는 계속 실행
 * - 3단계 실패: 1·2단계 결과는 유지
 * - 구조 변경 전체보정 플래그는 세 단계가 모두 성공한 경우에만 삭제
 * - 실행 도중 더 새로운 구조 변경 요청이 들어오면 기존 요청만 지우지 않고 유지
 ****************************************************/

var AUTOMATION_CORE_PIPELINE_CONFIG = Object.freeze({
  version: '2026-07-22-PHASE15',
  handlerName: 'AUTOMATION_runCoreDataSyncPipeline',

  leasePropertyKey: 'AUTOMATION_CORE_SYNC_LEASE_V1',
  lastRunPropertyKey: 'AUTOMATION_CORE_SYNC_LAST_RUN_V1',

  statusSheetName: '_자동화상태',
  statusTaskKey: 'CORE_DATA_SYNC',

  // Apps Script 단일 실행 제한(최대 6분)보다 약간 길게 잡아
  // 이전 실행이 살아 있는 동안 다음 5분 트리거가 겹치지 않게 한다.
  leaseDurationMs: 7 * 60 * 1000,
  leaseLockWaitMs: 5000,
  propertyLockWaitMs: 5000,
  maxErrorLength: 1500
});


/****************************************************
 * 공개 실행 함수
 ****************************************************/

/**
 * 정식 5분 시간기반 트리거 진입점.
 */
function AUTOMATION_runCoreDataSyncPipeline() {
  var startedAtMs = Date.now();
  var runToken = AUTOMATION_createCorePipelineRunToken_();
  var activeCutoverLease = AUTOMATION_getActiveCutoverLease_();
  var leaseResult = activeCutoverLease
    ? {
      acquired: false,
      reason: '정식 트리거 전환 작업이 진행 중입니다.',
      cutoverInProgress: true,
      existingStartedAt: String(activeCutoverLease.startedAt || ''),
      existingExpiresAt: String(activeCutoverLease.expiresAt || '')
    }
    : AUTOMATION_acquireCorePipelineLease_(runToken, startedAtMs);

  var summary = {
    version: AUTOMATION_CORE_PIPELINE_CONFIG.version,
    handler: AUTOMATION_CORE_PIPELINE_CONFIG.handlerName,
    runToken: runToken,
    startedAt: new Date(startedAtMs).toISOString(),
    finishedAt: '',
    durationMs: 0,
    status: 'STARTED',
    fullSyncRequestedAtStart: false,
    fullSyncRequestAtStart: null,
    fullSyncRequestClearResult: null,
    retryQueue: null,
    autoInputRepair: null,
    maintenance: null,
    healthMonitor: null,
    lease: leaseResult,
    stageCount: 0,
    successCount: 0,
    errorCount: 0,
    skippedCount: 0,
    stages: [],
    fatalError: ''
  };

  if (!leaseResult.acquired) {
    summary.status = leaseResult.cutoverInProgress
      ? 'SKIPPED_CUTOVER_IN_PROGRESS'
      : 'SKIPPED_ALREADY_RUNNING';
    summary.skippedCount = 1;
    summary.fatalError = leaseResult.reason || '이전 핵심 동기화 실행이 아직 진행 중입니다.';
    return AUTOMATION_finalizeCorePipelineSummary_(summary, startedAtMs);
  }

  try {
    var fullSyncRequest = AUTOMATION_getCoreFullSyncRequest_();
    summary.fullSyncRequestAtStart = fullSyncRequest || null;
    summary.fullSyncRequestedAtStart = !!(fullSyncRequest && fullSyncRequest.required === true);

    try {
      summary.retryQueue = AUTOMATION_processEditRetryQueue_();
    } catch (retryQueueErr) {
      summary.retryQueue = {
        status: 'ERROR',
        error: AUTOMATION_errorMessage_(retryQueueErr)
      };
      console.error('[AUTOMATION_runCoreDataSyncPipeline][RETRY_QUEUE] ' + summary.retryQueue.error, retryQueueErr);
    }

    try {
      summary.autoInputRepair = AUTOEDIT_runScheduledRepairSlice_();
    } catch (autoInputRepairErr) {
      summary.autoInputRepair = {
        status: 'ERROR',
        error: AUTOMATION_errorMessage_(autoInputRepairErr),
        scannedRows: 0,
        changedCells: 0
      };
      console.error('[AUTOMATION_runCoreDataSyncPipeline][AUTO_INPUT_REPAIR] ' + summary.autoInputRepair.error, autoInputRepairErr);
    }

    var stage1 = AUTOMATION_runCorePipelineStage_(
      'MASTER_TO_COMPLETED',
      '마스터 → 수주확정',
      'CMS_runFullSyncForAutomationPipeline_'
    );
    summary.stages.push(stage1);

    if (stage1.status !== 'SUCCESS') {
      summary.stages.push(AUTOMATION_makeSkippedCoreStage_(
        'COMPLETED_TO_VENDOR',
        '수주확정 → 수행사',
        '1단계 실패로 후속 단계 중단'
      ));
      summary.stages.push(AUTOMATION_makeSkippedCoreStage_(
        'COMPLETED_TO_IT_MAINTENANCE',
        '수주확정 → 정보통신유지보수',
        '1단계 실패로 후속 단계 중단'
      ));
      summary.status = 'FAILED_STAGE1';
    } else {
      var stage2 = AUTOMATION_runCorePipelineStage_(
        'COMPLETED_TO_VENDOR',
        '수주확정 → 수행사',
        'vendorSyncRunFullSyncForAutomationPipeline_'
      );
      summary.stages.push(stage2);

      var stage3 = AUTOMATION_runCorePipelineStage_(
        'COMPLETED_TO_IT_MAINTENANCE',
        '수주확정 → 정보통신유지보수',
        'ITMNEW_runMissingContractSyncForPipeline_2026'
      );
      summary.stages.push(stage3);

      if (stage2.status === 'SUCCESS' && stage3.status === 'SUCCESS') {
        summary.status = 'COMPLETED';

        if (summary.fullSyncRequestedAtStart) {
          summary.fullSyncRequestClearResult =
            AUTOMATION_clearCoreFullSyncRequestIfUnchanged_(fullSyncRequest);
        }
      } else {
        summary.status = 'COMPLETED_WITH_ERRORS';
      }
    }

    if (
      summary.status === 'COMPLETED' &&
      summary.retryQueue &&
      (Number(summary.retryQueue.retried || 0) > 0 || Number(summary.retryQueue.failed || 0) > 0)
    ) {
      summary.status = 'COMPLETED_WITH_PENDING_RETRIES';
    }

    try {
      summary.maintenance = AUTOMATION_runScheduledMaintenanceIfDue_({
        hardDeadlineMs: startedAtMs + 5 * 60 * 1000
      });
    } catch (maintenanceErr) {
      summary.maintenance = {
        status: 'ERROR',
        error: AUTOMATION_errorMessage_(maintenanceErr)
      };
      console.error(
        '[AUTOMATION_runCoreDataSyncPipeline][MAINTENANCE] ' + summary.maintenance.error,
        maintenanceErr
      );
    }
  } catch (err) {
    summary.status = 'FATAL_ERROR';
    summary.fatalError = AUTOMATION_truncateCorePipelineText_(
      AUTOMATION_errorMessage_(err),
      AUTOMATION_CORE_PIPELINE_CONFIG.maxErrorLength
    );
    console.error('[AUTOMATION_runCoreDataSyncPipeline] ' + summary.fatalError, err);
  } finally {
    AUTOMATION_releaseCorePipelineLease_(runToken);
  }

  return AUTOMATION_finalizeCorePipelineSummary_(summary, startedAtMs);
}


/**
 * 메뉴/편집기에서 수동으로 실행한 뒤 결과를 팝업으로 확인한다.
 * 실제 동작은 정식 시간 트리거 핸들러와 동일하다.
 */
function AUTOMATION_runCoreDataSyncPipelineNow() {
  var result = AUTOMATION_runCoreDataSyncPipeline();
  var lines = [
    '상태: ' + result.status,
    '소요시간: ' + (result.durationMs / 1000).toFixed(1) + '초',
    '성공 단계: ' + result.successCount + '개',
    '오류 단계: ' + result.errorCount + '개',
    '건너뜀: ' + result.skippedCount + '개',
    '재처리 큐: ' + (
      result.retryQueue
        ? ('처리 ' + Number(result.retryQueue.processed || 0) +
          ' / 성공 ' + Number(result.retryQueue.succeeded || 0) +
          ' / 재시도 ' + Number(result.retryQueue.retried || 0) +
          ' / 실패 ' + Number(result.retryQueue.failed || 0))
        : '해당없음'
    ),
    '자동입력 보정: ' + (
      result.autoInputRepair
        ? (String(result.autoInputRepair.status || '') +
          ' / 검사 ' + Number(result.autoInputRepair.scannedRows || 0) +
          ' / 보정 ' + Number(result.autoInputRepair.changedCells || 0))
        : '해당없음'
    )
  ];

  result.stages.forEach(function(stage) {
    lines.push(
      stage.label + ': ' + stage.status +
      (stage.error ? ' / ' + stage.error : '')
    );
  });

  if (result.fatalError) {
    lines.push('치명 오류: ' + result.fatalError);
  }

  SpreadsheetApp.getUi().alert(
    '핵심 데이터 동기화',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}


/**
 * 마지막 핵심 동기화 결과를 Script Properties에서 조회한다.
 */
function AUTOMATION_getCoreDataSyncLastRun() {
  return AUTOMATION_readJsonProperty_(
    PropertiesService.getScriptProperties(),
    AUTOMATION_CORE_PIPELINE_CONFIG.lastRunPropertyKey
  );
}


/**
 * 숨김 상태 시트를 표시하고 이동한다.
 */
function AUTOMATION_showAutomationStatusSheet() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss.getSheetByName(AUTOMATION_CORE_PIPELINE_CONFIG.statusSheetName);

  if (!sheet) {
    throw new Error('자동화 상태 시트가 아직 없습니다. 핵심 동기화를 한 번 실행하세요.');
  }

  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


/****************************************************
 * 단계 실행
 ****************************************************/

function AUTOMATION_runCorePipelineStage_(stageKey, label, handlerName) {
  var startedAtMs = Date.now();
  var stage = {
    key: stageKey,
    label: label,
    handler: handlerName,
    status: 'RUNNING',
    startedAt: new Date(startedAtMs).toISOString(),
    finishedAt: '',
    durationMs: 0,
    result: null,
    error: ''
  };

  try {
    var handler = AUTOMATION_resolveHandler_(handlerName);

    if (typeof handler !== 'function') {
      throw new Error('핵심 동기화 단계 핸들러를 찾을 수 없습니다: ' + handlerName);
    }

    var result = handler();
    stage.result = AUTOMATION_makeCorePipelineJsonSafe_(result);
    stage.status = 'SUCCESS';
  } catch (err) {
    stage.status = 'ERROR';
    stage.error = AUTOMATION_truncateCorePipelineText_(
      AUTOMATION_errorMessage_(err),
      AUTOMATION_CORE_PIPELINE_CONFIG.maxErrorLength
    );

    console.error(
      '[AUTOMATION_runCoreDataSyncPipeline][' + stageKey + '] ' + stage.error,
      err
    );
  } finally {
    stage.finishedAt = new Date().toISOString();
    stage.durationMs = Date.now() - startedAtMs;
  }

  return stage;
}


function AUTOMATION_makeSkippedCoreStage_(stageKey, label, reason) {
  return {
    key: stageKey,
    label: label,
    handler: '',
    status: 'SKIPPED',
    startedAt: '',
    finishedAt: '',
    durationMs: 0,
    result: null,
    error: reason || ''
  };
}


/****************************************************
 * 파이프라인 lease
 ****************************************************/

function AUTOMATION_acquireCorePipelineLease_(runToken, nowMs) {
  var lock = LockService.getUserLock();

  if (!lock.tryLock(AUTOMATION_CORE_PIPELINE_CONFIG.leaseLockWaitMs)) {
    return {
      acquired: false,
      reason: '파이프라인 lease 확인용 잠금을 얻지 못했습니다.'
    };
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var key = AUTOMATION_CORE_PIPELINE_CONFIG.leasePropertyKey;
    var existing = AUTOMATION_readJsonProperty_(props, key);

    if (
      existing &&
      existing.token &&
      Number(existing.expiresAtMs) > nowMs
    ) {
      return {
        acquired: false,
        reason: '이전 파이프라인 실행 lease가 유효합니다.',
        existingToken: String(existing.token),
        existingStartedAt: existing.startedAt || '',
        existingExpiresAt: existing.expiresAt || ''
      };
    }

    var expiresAtMs = nowMs + AUTOMATION_CORE_PIPELINE_CONFIG.leaseDurationMs;
    var lease = {
      token: runToken,
      handler: AUTOMATION_CORE_PIPELINE_CONFIG.handlerName,
      startedAtMs: nowMs,
      startedAt: new Date(nowMs).toISOString(),
      expiresAtMs: expiresAtMs,
      expiresAt: new Date(expiresAtMs).toISOString()
    };

    props.setProperty(key, JSON.stringify(lease));

    return {
      acquired: true,
      token: runToken,
      expiresAt: lease.expiresAt
    };
  } finally {
    lock.releaseLock();
  }
}


function AUTOMATION_releaseCorePipelineLease_(runToken) {
  var lock = LockService.getUserLock();

  if (!lock.tryLock(AUTOMATION_CORE_PIPELINE_CONFIG.leaseLockWaitMs)) {
    console.warn('[AUTOMATION_releaseCorePipelineLease_] lease 삭제 잠금 획득 실패');
    return false;
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var key = AUTOMATION_CORE_PIPELINE_CONFIG.leasePropertyKey;
    var current = AUTOMATION_readJsonProperty_(props, key);

    if (!current) return true;

    if (String(current.token || '') !== String(runToken || '')) {
      return false;
    }

    props.deleteProperty(key);
    return true;
  } finally {
    lock.releaseLock();
  }
}


/****************************************************
 * 전체보정 요청 안전 소비
 ****************************************************/

function AUTOMATION_clearCoreFullSyncRequestIfUnchanged_(expectedRequest) {
  if (!expectedRequest) {
    return {
      cleared: false,
      reason: 'NO_REQUEST_AT_START'
    };
  }

  var lock = LockService.getUserLock();

  if (!lock.tryLock(AUTOMATION_CORE_PIPELINE_CONFIG.propertyLockWaitMs)) {
    return {
      cleared: false,
      reason: 'PROPERTY_LOCK_UNAVAILABLE'
    };
  }

  try {
    var props = PropertiesService.getScriptProperties();
    var key = AUTOMATION_DISPATCHER_CONFIG.fullSyncRequestPropertyKey;
    var current = AUTOMATION_readJsonProperty_(props, key);

    if (!current) {
      return {
        cleared: false,
        reason: 'ALREADY_CLEARED'
      };
    }

    var unchanged =
      Number(current.requestCount || 0) === Number(expectedRequest.requestCount || 0) &&
      String(current.firstRequestedAt || '') === String(expectedRequest.firstRequestedAt || '') &&
      String(current.lastRequestedAt || '') === String(expectedRequest.lastRequestedAt || '');

    if (!unchanged) {
      return {
        cleared: false,
        reason: 'NEWER_REQUEST_RETAINED',
        currentRequestCount: Number(current.requestCount || 0),
        expectedRequestCount: Number(expectedRequest.requestCount || 0)
      };
    }

    props.deleteProperty(key);

    return {
      cleared: true,
      reason: 'CLEARED_AFTER_SUCCESS'
    };
  } finally {
    lock.releaseLock();
  }
}


/****************************************************
 * 결과 집계·저장
 ****************************************************/

function AUTOMATION_finalizeCorePipelineSummary_(summary, startedAtMs) {
  var finishedAtMs = Date.now();

  summary.finishedAt = new Date(finishedAtMs).toISOString();
  summary.durationMs = Math.max(0, finishedAtMs - startedAtMs);
  summary.stageCount = summary.stages.length;
  summary.successCount = 0;
  summary.errorCount = 0;
  summary.skippedCount = String(summary.status || '').indexOf('SKIPPED_') === 0 ? 1 : 0;

  summary.stages.forEach(function(stage) {
    if (stage.status === 'SUCCESS') summary.successCount++;
    if (stage.status === 'ERROR') summary.errorCount++;
    if (stage.status === 'SKIPPED') summary.skippedCount++;
  });

  try {
    if (typeof AUTOMATION_runHealthMonitorSafe_ === 'function') {
      summary.healthMonitor = AUTOMATION_runHealthMonitorSafe_({
        source: 'CORE_PIPELINE',
        currentCoreSummary: AUTOMATION_makeCorePipelineJsonSafe_(summary),
        deadlineMs: startedAtMs + 5.5 * 60 * 1000
      });
    }
  } catch (healthMonitorErr) {
    summary.healthMonitor = {
      status: 'ERROR',
      error: AUTOMATION_errorMessage_(healthMonitorErr)
    };
    console.error(
      '[AUTOMATION_finalizeCorePipelineSummary_][HEALTH_MONITOR] ' + summary.healthMonitor.error,
      healthMonitorErr
    );
  }

  AUTOMATION_persistCorePipelineSummary_(summary);
  AUTOMATION_writeCorePipelineStatus_(summary);

  console.log('[AUTOMATION_runCoreDataSyncPipeline] ' + JSON.stringify(summary));
  return summary;
}


function AUTOMATION_persistCorePipelineSummary_(summary) {
  try {
    PropertiesService.getScriptProperties().setProperty(
      AUTOMATION_CORE_PIPELINE_CONFIG.lastRunPropertyKey,
      JSON.stringify(AUTOMATION_makeCorePipelineJsonSafe_(summary))
    );
  } catch (err) {
    console.error('[AUTOMATION_persistCorePipelineSummary_] ' + AUTOMATION_errorMessage_(err), err);
  }
}


function AUTOMATION_writeCorePipelineStatus_(summary) {
  try {
    var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();

    var sheet = ss.getSheetByName(AUTOMATION_CORE_PIPELINE_CONFIG.statusSheetName);
    var created = false;

    if (!sheet) {
      sheet = ss.insertSheet(AUTOMATION_CORE_PIPELINE_CONFIG.statusSheetName);
      created = true;
    }

    var headers = [
      '작업키', '작업명', '최근시작', '최근종료', '상태', '소요초',
      '재처리큐', '자동입력보정', '유지관리', '장애감시', '1단계', '2단계', '3단계', '전체보정요청', '플래그처리',
      '오류요약', '실행토큰', '버전'
    ];

    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    var headerMismatch = headers.some(function(header, index) {
      return String(currentHeaders[index] || '') !== header;
    });

    if (headerMismatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }

    var taskRow = 2;
    var lastRow = sheet.getLastRow();

    if (lastRow >= 2) {
      var taskKeys = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();

      for (var i = 0; i < taskKeys.length; i++) {
        if (String(taskKeys[i][0] || '') === AUTOMATION_CORE_PIPELINE_CONFIG.statusTaskKey) {
          taskRow = i + 2;
          break;
        }

        taskRow = lastRow + 1;
      }
    }

    var stageByKey = {};
    summary.stages.forEach(function(stage) {
      stageByKey[stage.key] = stage;
    });

    var errorParts = [];
    summary.stages.forEach(function(stage) {
      if (stage.error) errorParts.push(stage.label + ': ' + stage.error);
    });
    if (summary.fatalError) errorParts.push('치명오류: ' + summary.fatalError);
    if (summary.maintenance && summary.maintenance.error) {
      errorParts.push('유지관리: ' + summary.maintenance.error);
    }

    var clearResult = summary.fullSyncRequestClearResult;
    var clearText = clearResult
      ? (clearResult.cleared ? '삭제완료' : String(clearResult.reason || '유지'))
      : '해당없음';

    var row = [
      AUTOMATION_CORE_PIPELINE_CONFIG.statusTaskKey,
      '핵심 데이터 동기화',
      summary.startedAt,
      summary.finishedAt,
      summary.status,
      Math.round(summary.durationMs / 100) / 10,
      AUTOMATION_coreRetryQueueStatusText_(summary.retryQueue),
      AUTOMATION_coreAutoInputRepairStatusText_(summary.autoInputRepair),
      AUTOMATION_coreMaintenanceStatusText_(summary.maintenance),
      AUTOMATION_coreHealthMonitorStatusText_(summary.healthMonitor),
      AUTOMATION_coreStageStatusText_(stageByKey.MASTER_TO_COMPLETED),
      AUTOMATION_coreStageStatusText_(stageByKey.COMPLETED_TO_VENDOR),
      AUTOMATION_coreStageStatusText_(stageByKey.COMPLETED_TO_IT_MAINTENANCE),
      summary.fullSyncRequestedAtStart ? '예' : '아니오',
      clearText,
      AUTOMATION_truncateCorePipelineText_(errorParts.join(' | '), 2000),
      summary.runToken,
      summary.version
    ];

    sheet.getRange(taskRow, 1, 1, row.length).setValues([row]);

    if (created || headerMismatch) {
      sheet.autoResizeColumns(1, headers.length);
    }

    if (created) {
      try {
        sheet.hideSheet();
      } catch (ignoreHideError) {
        // 상태 기록 자체는 성공했으므로 숨김 실패는 무시
      }
    }
  } catch (err) {
    console.error('[AUTOMATION_writeCorePipelineStatus_] ' + AUTOMATION_errorMessage_(err), err);
  }
}


function AUTOMATION_coreRetryQueueStatusText_(retryQueue) {
  if (!retryQueue) return '해당없음';

  return AUTOMATION_truncateCorePipelineText_(
    [
      String(retryQueue.status || ''),
      '처리 ' + Number(retryQueue.processed || 0),
      '성공 ' + Number(retryQueue.succeeded || 0),
      '재시도 ' + Number(retryQueue.retried || 0),
      '실패 ' + Number(retryQueue.failed || 0)
    ].join(' / '),
    800
  );
}


function AUTOMATION_coreAutoInputRepairStatusText_(repair) {
  if (!repair) return '해당없음';

  var text = [
    String(repair.status || ''),
    '검사 ' + Number(repair.scannedRows || 0),
    '보정 ' + Number(repair.changedCells || 0)
  ].join(' / ');

  if (repair.error) text += ' / ' + repair.error;
  return AUTOMATION_truncateCorePipelineText_(text, 800);
}


function AUTOMATION_coreMaintenanceStatusText_(maintenance) {
  if (!maintenance) return '해당없음';

  var text = [
    String(maintenance.status || ''),
    '단계 ' + Number(maintenance.endStepIndex || maintenance.stepCount || 0),
    maintenance.completedCycle ? '사이클완료' : '계속예정'
  ].join(' / ');

  if (maintenance.error) text += ' / ' + maintenance.error;
  return AUTOMATION_truncateCorePipelineText_(text, 800);
}


function AUTOMATION_coreHealthMonitorStatusText_(healthMonitor) {
  if (!healthMonitor) return '해당없음';

  var text = [
    String(healthMonitor.status || ''),
    '현재장애 ' + Number(healthMonitor.activeIssueCount || 0),
    '전송 ' + Number(healthMonitor.sentCount || 0),
    '복구 ' + Number(healthMonitor.recoverySentCount || 0)
  ].join(' / ');

  if (healthMonitor.error) text += ' / ' + healthMonitor.error;
  return AUTOMATION_truncateCorePipelineText_(text, 800);
}


function AUTOMATION_coreStageStatusText_(stage) {
  if (!stage) return '';

  var text = stage.status + ' (' + (stage.durationMs / 1000).toFixed(1) + '초)';
  if (stage.error) text += ' / ' + stage.error;
  return AUTOMATION_truncateCorePipelineText_(text, 800);
}


/****************************************************
 * 공통 보조 함수
 ****************************************************/

function AUTOMATION_createCorePipelineRunToken_() {
  try {
    return Utilities.getUuid();
  } catch (ignoreUuidError) {
    return 'CORE-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
  }
}


function AUTOMATION_makeCorePipelineJsonSafe_(value) {
  if (value === null || typeof value === 'undefined') return null;

  try {
    return JSON.parse(JSON.stringify(value));
  } catch (err) {
    return {
      serializationError: AUTOMATION_errorMessage_(err),
      valueText: AUTOMATION_truncateCorePipelineText_(String(value), 1000)
    };
  }
}


function AUTOMATION_truncateCorePipelineText_(value, maxLength) {
  var text = String(value || '');
  var limit = Math.max(1, Number(maxLength) || 1);

  if (text.length <= limit) return text;
  return text.slice(0, limit - 3) + '...';
}
