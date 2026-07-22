/****************************************************
 * AutomationCutover.gs
 * 정식 13개 트리거 전환·사전점검·사후검증 - 7단계
 *
 * 목적:
 * - bang@s1samsung.com 단일 자동화 계정에서만 전환 허용
 * - 기존 트리거 삭제 전에 권한·핸들러·대상 파일·활성 작업을 사전점검
 * - 최근 30분 내 수동 핵심 데이터 동기화 성공을 전환 필수조건으로 확인
 * - 전환 중에는 TRIGGER_CUTOVER lease로 신규 자동 작업 진입 차단
 * - 정식 13개 설치 후 구조 검증과 전환 이력을 숨김 시트에 기록
 *
 * 주의:
 * - 이 파일을 배포하는 것만으로 트리거는 변경되지 않는다.
 * - 실제 전환은 AUTOMATION_executeCanonicalCutover()를 수동 실행해야 한다.
 ****************************************************/

var AUTOMATION_CUTOVER_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE7',
  cutoverLeaseModuleKey: 'TRIGGER_CUTOVER',
  cutoverLeaseTtlMs: 12 * 60 * 1000,
  cutoverLeaseWaitMs: 1000,
  statusSheetName: '_자동화전환기록',
  statePropertyKey: 'AUTOMATION_CUTOVER_STATE_V1',
  lastPreflightPropertyKey: 'AUTOMATION_CUTOVER_PREFLIGHT_V1',
  historyPropertyKey: 'AUTOMATION_CUTOVER_HISTORY_V1',
  historyLimit: 20,
  coreSuccessStatuses: Object.freeze([
    'COMPLETED',
    'COMPLETED_WITH_PENDING_RETRIES'
  ]),
  freshCoreRunMs: 30 * 60 * 1000,
  maxDetailLength: 12000
});


/****************************************************
 * 공개 실행 함수
 ****************************************************/

/**
 * 트리거를 변경하지 않고 전환 가능 상태를 점검한다.
 */
function AUTOMATION_previewCutoverReadiness() {
  TRG_assertAutomationOwner_();

  var report = AUTOMATION_buildCutoverReadiness_({
    checkActiveLeases: true,
    includeTriggerPreflight: true
  });

  AUTOMATION_storeCutoverPreflight_(report);
  AUTOMATION_appendCutoverLog_({
    cutoverId: report.cutoverId,
    phase: 'PREFLIGHT',
    status: report.ok ? 'READY' : 'BLOCKED',
    startedAt: report.generatedAt,
    finishedAt: report.generatedAt,
    triggerSummary: report.triggerSummary,
    activeLeaseCount: report.activeLeases.length,
    retrySummary: report.retryQueue,
    coreStatus: report.coreLastRun ? report.coreLastRun.status : '',
    message: AUTOMATION_cutoverReportMessage_(report),
    detail: report
  });

  AUTOMATION_showCutoverReadinessAlert_(report);
  return report;
}


/**
 * 사전점검 → 전환 가드 → 정식 13개 재설치 → 사후검증을 수행한다.
 * 실제 전환을 위한 공식 실행 함수다.
 */
function AUTOMATION_executeCanonicalCutover() {
  TRG_assertAutomationOwner_();
  TRG_assertCanonicalInstallEnabled_();

  var initialReport = AUTOMATION_buildCutoverReadiness_({
    checkActiveLeases: true,
    includeTriggerPreflight: true
  });

  AUTOMATION_storeCutoverPreflight_(initialReport);

  if (!initialReport.ok) {
    AUTOMATION_appendCutoverLog_({
      cutoverId: initialReport.cutoverId,
      phase: 'CUTOVER',
      status: 'BLOCKED_PREFLIGHT',
      startedAt: initialReport.generatedAt,
      finishedAt: new Date().toISOString(),
      triggerSummary: initialReport.triggerSummary,
      activeLeaseCount: initialReport.activeLeases.length,
      retrySummary: initialReport.retryQueue,
      coreStatus: initialReport.coreLastRun ? initialReport.coreLastRun.status : '',
      message: AUTOMATION_cutoverReportMessage_(initialReport),
      detail: initialReport
    });

    AUTOMATION_showCutoverReadinessAlert_(initialReport);
    return {
      ok: false,
      status: 'BLOCKED_PREFLIGHT',
      report: initialReport
    };
  }

  if (initialReport.alreadyCanonicalHealthy) {
    var alreadyHealthy = AUTOMATION_verifyCutoverNow_({
      cutoverId: initialReport.cutoverId,
      writeLog: true,
      showAlert: true,
      statusOverride: 'ALREADY_CANONICAL'
    });

    return {
      ok: alreadyHealthy.ok,
      status: 'ALREADY_CANONICAL',
      verification: alreadyHealthy
    };
  }

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '정식 13개 트리거 전환',
    [
      '실행 계정: ' + TRG_getEffectiveUserEmail_(),
      '현재 설치형 트리거: ' + initialReport.triggerSummary.installedTriggerCount + '개',
      '정식 계획: ' + initialReport.triggerSummary.canonicalPlannedTriggerCount + '개',
      '구형/개별: ' + initialReport.triggerSummary.legacyTriggerCount + '개',
      '고아: ' + initialReport.triggerSummary.orphanTriggerCount + '개',
      '미분류: ' + initialReport.triggerSummary.unknownTriggerCount + '개',
      '',
      '실행 순서:',
      '1. 최근 30분 내 핵심 데이터 동기화 성공 결과 재확인',
      '2. 전환 가드로 신규 자동 작업 진입 차단',
      '3. 기존 설치형 트리거 전체 삭제',
      '4. 정식 13개 트리거 설치',
      '5. 설치 구조 사후검증 및 전환 기록 저장',
      '',
      '단순 onOpen/onEdit/onSelectionChange와 웹앱 doGet/doPost는 영향을 받지 않습니다.',
      '계속하시겠습니까?'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return {
      ok: false,
      status: 'CANCELLED'
    };
  }

  var cutoverId = initialReport.cutoverId || AUTOMATION_createCutoverId_();
  var startedAtMs = Date.now();
  var state = {
    version: AUTOMATION_CUTOVER_CONFIG.version,
    cutoverId: cutoverId,
    status: 'STARTED',
    startedAt: new Date(startedAtMs).toISOString(),
    finishedAt: '',
    phase: 'ACQUIRE_CUTOVER_GUARD',
    coreSync: AUTOMATION_makeCutoverJsonSafe_(initialReport.coreLastRun),
    reinstall: null,
    verification: null,
    error: ''
  };

  AUTOMATION_saveCutoverState_(state);
  var cutoverLease = null;

  try {
    var coreResult = initialReport.coreLastRun || {};

    if (!AUTOMATION_isFreshSuccessfulCoreRun_(coreResult)) {
      throw new Error(
        '최근 30분 이내 성공한 핵심 데이터 동기화 결과가 없어 전환을 중단했습니다. ' +
        'AUTOMATION_runCoreDataSyncPipelineNow()를 먼저 별도로 실행하세요.'
      );
    }

    AUTOMATION_saveCutoverState_(state);

    cutoverLease = AUTOMATION_acquireModuleLease_(
      AUTOMATION_CUTOVER_CONFIG.cutoverLeaseModuleKey,
      {
        taskName: 'AUTOMATION_executeCanonicalCutover',
        ttlMs: AUTOMATION_CUTOVER_CONFIG.cutoverLeaseTtlMs,
        waitMs: AUTOMATION_CUTOVER_CONFIG.cutoverLeaseWaitMs
      }
    );

    if (!cutoverLease.acquired) {
      throw new Error(
        '다른 전환 작업이 진행 중이라 전환 가드를 획득하지 못했습니다: ' +
        String(cutoverLease.reason || 'LEASE_BUSY')
      );
    }

    // 가드를 잡은 뒤 다시 확인해야 전환 직전 실행 중인 다른 모듈을 놓치지 않는다.
    var guardedReport = AUTOMATION_buildCutoverReadiness_({
      checkActiveLeases: true,
      includeTriggerPreflight: true,
      excludeLeaseModules: [AUTOMATION_CUTOVER_CONFIG.cutoverLeaseModuleKey],
      cutoverId: cutoverId
    });

    if (guardedReport.activeLeases.length > 0) {
      throw new Error(
        '현재 실행 중인 자동화가 있어 전환을 중단했습니다: ' +
        guardedReport.activeLeases.map(function(item) {
          return item.moduleKey + (item.taskName ? '(' + item.taskName + ')' : '');
        }).join(', ')
      );
    }

    var nonLeaseBlockers = guardedReport.blockers.filter(function(blocker) {
      return String(blocker.code || '') !== 'ACTIVE_AUTOMATION_LEASE';
    });

    if (nonLeaseBlockers.length > 0) {
      throw new Error(
        '전환 직전 사전점검에서 차단 항목이 발견됐습니다: ' +
        nonLeaseBlockers.map(function(item) { return item.message; }).join(' / ')
      );
    }

    AUTOMATION_refreshModuleLease_(cutoverLease, true);
    state.phase = 'REINSTALL_TRIGGERS';
    AUTOMATION_saveCutoverState_(state);

    var reinstallResult = TRG_reinstallCanonicalInternal_({
      preflight: guardedReport.triggerPreflight,
      currentTriggers: ScriptApp.getProjectTriggers(),
      source: 'AUTOMATION_executeCanonicalCutover'
    });

    state.reinstall = AUTOMATION_makeCutoverJsonSafe_(reinstallResult);
    AUTOMATION_refreshModuleLease_(cutoverLease, true);

    state.phase = 'POST_VERIFY';
    AUTOMATION_saveCutoverState_(state);

    var verification = AUTOMATION_buildCutoverVerification_({
      cutoverId: cutoverId,
      ignoreCutoverLease: true
    });
    state.verification = AUTOMATION_makeCutoverJsonSafe_(verification);

    if (!verification.ok) {
      throw new Error(
        '정식 트리거 설치 후 사후검증이 실패했습니다: ' +
        verification.blockers.map(function(item) { return item.message; }).join(' / ')
      );
    }

    state.status = 'COMPLETED';
    state.phase = 'COMPLETED';
    state.finishedAt = new Date().toISOString();
    AUTOMATION_saveCutoverState_(state);
    AUTOMATION_appendCutoverHistory_(state);
    AUTOMATION_appendCutoverLog_({
      cutoverId: cutoverId,
      phase: 'CUTOVER',
      status: 'COMPLETED',
      startedAt: state.startedAt,
      finishedAt: state.finishedAt,
      triggerSummary: verification.triggerSummary,
      deletedCount: reinstallResult.deletedCount,
      createdCount: reinstallResult.createdCount,
      activeLeaseCount: 0,
      retrySummary: verification.retryQueue,
      coreStatus: coreResult.status,
      message: '최근 핵심 동기화 성공을 확인한 뒤 정식 13개 트리거 전환과 사후검증을 완료했습니다.',
      detail: state
    });

    ui.alert(
      '정식 전환 완료',
      [
        '핵심 동기화: ' + coreResult.status,
        '기존 트리거 삭제: ' + reinstallResult.deletedCount + '개',
        '정식 트리거 설치: ' + reinstallResult.createdCount + '개',
        '정식 계획 일치: ' + verification.triggerSummary.canonicalMatchedTriggerCount + '개',
        '고아·구형·미분류: 0개',
        '',
        '상세 기록은 ' + AUTOMATION_CUTOVER_CONFIG.statusSheetName + ' 시트를 확인하세요.'
      ].join('\n'),
      ui.ButtonSet.OK
    );

    return {
      ok: true,
      status: 'COMPLETED',
      cutoverId: cutoverId,
      coreSync: coreResult,
      reinstall: reinstallResult,
      verification: verification
    };
  } catch (err) {
    if (err && err.triggerReinstallResult && !state.reinstall) {
      state.reinstall = AUTOMATION_makeCutoverJsonSafe_(err.triggerReinstallResult);
    }

    state.status = 'FAILED';
    state.phase = state.phase || 'FAILED';
    state.finishedAt = new Date().toISOString();
    state.error = AUTOMATION_cutoverErrorMessage_(err);
    AUTOMATION_saveCutoverState_(state);
    AUTOMATION_appendCutoverHistory_(state);

    var failureVerification;
    try {
      failureVerification = AUTOMATION_buildCutoverVerification_({
        cutoverId: cutoverId,
        ignoreCutoverLease: true
      });
    } catch (verifyErr) {
      failureVerification = {
        ok: false,
        error: AUTOMATION_cutoverErrorMessage_(verifyErr)
      };
    }

    AUTOMATION_appendCutoverLog_({
      cutoverId: cutoverId,
      phase: 'CUTOVER',
      status: 'FAILED',
      startedAt: state.startedAt,
      finishedAt: state.finishedAt,
      triggerSummary: failureVerification.triggerSummary || initialReport.triggerSummary,
      deletedCount: state.reinstall ? Number(state.reinstall.deletedCount || 0) : 0,
      createdCount: state.reinstall ? Number(state.reinstall.createdCount || 0) : 0,
      activeLeaseCount: 0,
      retrySummary: failureVerification.retryQueue || initialReport.retryQueue,
      coreStatus: state.coreSync ? String(state.coreSync.status || '') : '',
      message: state.error,
      detail: {
        state: state,
        verification: failureVerification
      }
    });

    try {
      TRG_writeStatusSheet_(TRG_buildStatusSnapshot_());
    } catch (ignoreStatusWriteError) {
      // 전환 실패 원인을 가리지 않도록 상태시트 오류는 무시
    }

    ui.alert(
      '정식 전환 실패',
      [
        '단계: ' + state.phase,
        '오류: ' + state.error,
        '',
        '현재 트리거 상태를 자동화 관리 > 트리거 현황 열기에서 확인하세요.',
        '정식 계획이 13개보다 적다면 문제를 수정한 뒤 정식 전환 실행을 다시 수행하면 됩니다.'
      ].join('\n'),
      ui.ButtonSet.OK
    );

    return {
      ok: false,
      status: 'FAILED',
      cutoverId: cutoverId,
      error: state.error,
      state: state,
      verification: failureVerification
    };
  } finally {
    if (cutoverLease && cutoverLease.acquired) {
      AUTOMATION_releaseModuleLease_(cutoverLease);
    }
  }
}


/**
 * 현재 정식 트리거 구조와 전환 후 운영 상태를 검증한다.
 */
function AUTOMATION_verifyCutoverNow() {
  TRG_assertAutomationOwner_();
  return AUTOMATION_verifyCutoverNow_({
    cutoverId: AUTOMATION_createCutoverId_(),
    writeLog: true,
    showAlert: true
  });
}


/**
 * 숨김 전환 기록 시트를 표시한다.
 */
function AUTOMATION_showCutoverLogSheet() {
  TRG_assertAutomationOwner_();

  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(AUTOMATION_CUTOVER_CONFIG.statusSheetName);

  if (!sheet) {
    sheet = AUTOMATION_ensureCutoverLogSheet_();
  }

  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


/**
 * 마지막 전환 상태를 반환한다.
 */
function AUTOMATION_getLastCutoverState() {
  return AUTOMATION_readCutoverJsonProperty_(
    AUTOMATION_CUTOVER_CONFIG.statePropertyKey
  );
}


/****************************************************
 * 사전점검
 ****************************************************/

function AUTOMATION_buildCutoverReadiness_(options) {
  options = options || {};

  var generatedAt = new Date().toISOString();
  var cutoverId = String(options.cutoverId || AUTOMATION_createCutoverId_());
  var blockers = [];
  var warnings = [];
  var triggerPreflight = null;
  var triggerSnapshot = null;

  try {
    TRG_assertAutomationOwner_();
  } catch (ownerErr) {
    blockers.push({
      code: 'OWNER_ACCOUNT_MISMATCH',
      message: AUTOMATION_cutoverErrorMessage_(ownerErr)
    });
  }

  if (options.includeTriggerPreflight !== false) {
    try {
      triggerPreflight = TRG_preflightCanonicalInstall_();
    } catch (preflightErr) {
      blockers.push({
        code: 'TRIGGER_PREFLIGHT_FAILED',
        message: AUTOMATION_cutoverErrorMessage_(preflightErr)
      });
    }
  }

  try {
    triggerSnapshot = TRG_buildStatusSnapshot_();
  } catch (snapshotErr) {
    blockers.push({
      code: 'TRIGGER_STATUS_FAILED',
      message: AUTOMATION_cutoverErrorMessage_(snapshotErr)
    });
  }

  var excludeLeaseModules = options.excludeLeaseModules || [];
  var leaseState = options.checkActiveLeases === false
    ? { active: [], stale: [] }
    : AUTOMATION_collectCutoverLeaseState_(excludeLeaseModules);

  if (leaseState.active.length > 0) {
    blockers.push({
      code: 'ACTIVE_AUTOMATION_LEASE',
      message: '현재 실행 중인 자동화 lease가 ' + leaseState.active.length + '개 있습니다.'
    });
  }

  if (leaseState.stale.length > 0) {
    warnings.push({
      code: 'STALE_AUTOMATION_LEASE',
      message: '만료된 lease 기록이 ' + leaseState.stale.length + '개 있습니다. 해당 모듈 다음 실행 시 자동 회수됩니다.'
    });
  }

  var retryQueue = AUTOMATION_getCutoverRetryQueueSummary_();

  if (!retryQueue.schemaOk) {
    blockers.push({
      code: 'RETRY_QUEUE_SCHEMA_INVALID',
      message: retryQueue.error || '재처리 큐 헤더 구조가 올바르지 않습니다.'
    });
  }

  if (retryQueue.activeRunningCount > 0) {
    blockers.push({
      code: 'RETRY_QUEUE_RUNNING',
      message: '현재 실행 중인 재처리 작업이 ' + retryQueue.activeRunningCount + '건 있습니다.'
    });
  }

  if (retryQueue.pendingCount > 0) {
    warnings.push({
      code: 'RETRY_QUEUE_PENDING',
      message: '대기·재시도 중인 편집 재처리 작업이 ' + retryQueue.pendingCount + '건 있습니다.'
    });
  }

  if (retryQueue.failCount > 0) {
    warnings.push({
      code: 'RETRY_QUEUE_FAILED',
      message: '최종 실패 상태의 편집 재처리 작업이 ' + retryQueue.failCount + '건 있습니다.'
    });
  }

  var coreLastRun = null;
  try {
    coreLastRun = AUTOMATION_getCoreDataSyncLastRun();
  } catch (coreErr) {
    warnings.push({
      code: 'CORE_LAST_RUN_UNAVAILABLE',
      message: '마지막 핵심 동기화 결과를 읽지 못했습니다: ' + AUTOMATION_cutoverErrorMessage_(coreErr)
    });
  }

  var coreFresh = AUTOMATION_isFreshSuccessfulCoreRun_(coreLastRun);
  if (!coreFresh) {
    blockers.push({
      code: 'CORE_SYNC_NOT_RECENT',
      message: '최근 30분 이내 성공한 핵심 동기화가 없습니다. AUTOMATION_runCoreDataSyncPipelineNow()를 먼저 별도로 실행하세요.'
    });
  }

  var scriptTimezone = '';
  try {
    scriptTimezone = Session.getScriptTimeZone();
  } catch (ignoreTimezoneError) {
    scriptTimezone = '';
  }

  if (scriptTimezone && scriptTimezone !== TRG_MANAGER_CONFIG.timezone) {
    warnings.push({
      code: 'SCRIPT_TIMEZONE_MISMATCH',
      message: '스크립트 시간대가 ' + scriptTimezone + '입니다. 정식 계획 기준은 ' + TRG_MANAGER_CONFIG.timezone + '입니다.'
    });
  }

  var triggerSummary = triggerSnapshot && triggerSnapshot.summary
    ? triggerSnapshot.summary
    : AUTOMATION_emptyTriggerSummary_();
  var alreadyCanonicalHealthy = !!(
    triggerSnapshot && TRG_isCanonicalSnapshotHealthy_(triggerSnapshot)
  );

  return {
    version: AUTOMATION_CUTOVER_CONFIG.version,
    cutoverId: cutoverId,
    generatedAt: generatedAt,
    ok: blockers.length === 0,
    alreadyCanonicalHealthy: alreadyCanonicalHealthy,
    blockers: blockers,
    warnings: warnings,
    activeLeases: leaseState.active,
    staleLeases: leaseState.stale,
    retryQueue: retryQueue,
    coreLastRun: AUTOMATION_makeCutoverJsonSafe_(coreLastRun),
    coreFresh: coreFresh,
    scriptTimezone: scriptTimezone,
    triggerPreflight: triggerPreflight,
    triggerSummary: triggerSummary,
    triggerSnapshot: triggerSnapshot
  };
}


function AUTOMATION_collectCutoverLeaseState_(excludeModules) {
  var excludeMap = {};
  (excludeModules || []).forEach(function(moduleKey) {
    excludeMap[String(moduleKey || '').toUpperCase()] = true;
  });

  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var nowMs = Date.now();
  var active = [];
  var stale = [];
  var prefix = AUTOMATION_RUNTIME_CONFIG.leasePropertyPrefix;

  Object.keys(all).forEach(function(propertyKey) {
    if (propertyKey.indexOf(prefix) !== 0) return;

    var moduleKey = propertyKey.slice(prefix.length).toUpperCase();
    if (excludeMap[moduleKey]) return;

    var record = AUTOMATION_parseCutoverJson_(all[propertyKey]);
    if (!record) {
      stale.push({
        propertyKey: propertyKey,
        moduleKey: moduleKey,
        reason: 'INVALID_JSON'
      });
      return;
    }

    var item = {
      propertyKey: propertyKey,
      moduleKey: moduleKey,
      taskName: String(record.taskName || ''),
      startedAt: String(record.startedAt || ''),
      heartbeatAt: String(record.heartbeatAt || record.startedAt || ''),
      expiresAt: String(record.expiresAt || ''),
      expiresAtMs: Number(record.expiresAtMs || 0)
    };

    if (item.expiresAtMs > nowMs) active.push(item);
    else stale.push(item);
  });

  var coreLeaseKey = AUTOMATION_CORE_PIPELINE_CONFIG.leasePropertyKey;
  var coreLease = AUTOMATION_parseCutoverJson_(all[coreLeaseKey]);

  if (coreLease) {
    var coreItem = {
      propertyKey: coreLeaseKey,
      moduleKey: 'CORE_DATA_SYNC',
      taskName: String(coreLease.handler || AUTOMATION_CORE_PIPELINE_CONFIG.handlerName),
      startedAt: String(coreLease.startedAt || ''),
      heartbeatAt: String(coreLease.startedAt || ''),
      expiresAt: String(coreLease.expiresAt || ''),
      expiresAtMs: Number(coreLease.expiresAtMs || 0)
    };

    if (!excludeMap.CORE_DATA_SYNC) {
      if (coreItem.expiresAtMs > nowMs) active.push(coreItem);
      else stale.push(coreItem);
    }
  } else if (all[coreLeaseKey]) {
    stale.push({
      propertyKey: coreLeaseKey,
      moduleKey: 'CORE_DATA_SYNC',
      reason: 'INVALID_JSON'
    });
  }

  return {
    active: active,
    stale: stale
  };
}


function AUTOMATION_getCutoverRetryQueueSummary_() {
  var summary = {
    exists: false,
    schemaOk: true,
    totalCount: 0,
    pendingCount: 0,
    activeRunningCount: 0,
    staleRunningCount: 0,
    doneCount: 0,
    failCount: 0,
    otherCount: 0,
    error: ''
  };

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(AUTOMATION_RUNTIME_CONFIG.retryQueueSheetName);

    if (!sheet) return summary;
    summary.exists = true;

    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    var index = {};

    headers.forEach(function(header, i) {
      index[header] = i;
      if (String(currentHeaders[i] || '') !== String(header)) {
        summary.schemaOk = false;
      }
    });

    if (!summary.schemaOk) {
      summary.error = '재처리 큐 헤더가 정식 구조와 일치하지 않습니다.';
      return summary;
    }

    var lastRow = sheet.getLastRow();
    if (lastRow < 2) return summary;

    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var nowMs = Date.now();

    values.forEach(function(row) {
      var status = String(row[index['상태']] || '').trim().toUpperCase();
      if (!status) return;

      summary.totalCount += 1;

      if (status === 'PENDING' || status === 'RETRY') {
        summary.pendingCount += 1;
      } else if (status === 'RUNNING') {
        var lastAttemptMs = AUTOMATION_parseCutoverDateMs_(row[index['최근시도일시']]);
        if (
          lastAttemptMs &&
          nowMs - lastAttemptMs < AUTOMATION_RUNTIME_CONFIG.retryQueueRunningStaleMs
        ) {
          summary.activeRunningCount += 1;
        } else {
          summary.staleRunningCount += 1;
        }
      } else if (status === 'DONE') {
        summary.doneCount += 1;
      } else if (status === 'FAIL') {
        summary.failCount += 1;
      } else {
        summary.otherCount += 1;
      }
    });
  } catch (err) {
    summary.schemaOk = false;
    summary.error = AUTOMATION_cutoverErrorMessage_(err);
  }

  return summary;
}


function AUTOMATION_isFreshSuccessfulCoreRun_(coreLastRun) {
  if (!coreLastRun) return false;

  var status = String(coreLastRun.status || '');
  if (AUTOMATION_CUTOVER_CONFIG.coreSuccessStatuses.indexOf(status) < 0) return false;

  var finishedMs = AUTOMATION_parseCutoverDateMs_(coreLastRun.finishedAt);
  if (!finishedMs) return false;

  return Date.now() - finishedMs <= AUTOMATION_CUTOVER_CONFIG.freshCoreRunMs;
}


/****************************************************
 * 사후검증
 ****************************************************/

function AUTOMATION_verifyCutoverNow_(options) {
  options = options || {};
  var verification = AUTOMATION_buildCutoverVerification_({
    cutoverId: options.cutoverId,
    ignoreCutoverLease: false
  });

  if (options.writeLog !== false) {
    AUTOMATION_appendCutoverLog_({
      cutoverId: verification.cutoverId,
      phase: 'POST_VERIFY',
      status: options.statusOverride || (verification.ok ? 'HEALTHY' : 'UNHEALTHY'),
      startedAt: verification.generatedAt,
      finishedAt: verification.generatedAt,
      triggerSummary: verification.triggerSummary,
      activeLeaseCount: verification.activeLeases.length,
      retrySummary: verification.retryQueue,
      coreStatus: verification.coreLastRun ? verification.coreLastRun.status : '',
      message: verification.ok
        ? '정식 13개 트리거 구조가 정상입니다.'
        : verification.blockers.map(function(item) { return item.message; }).join(' / '),
      detail: verification
    });
  }

  if (options.showAlert !== false) {
    SpreadsheetApp.getUi().alert(
      '자동화 전환 사후검증',
      [
        '상태: ' + (verification.ok ? '정상' : '확인 필요'),
        '설치형 트리거: ' + verification.triggerSummary.installedTriggerCount + '개',
        '계획 일치: ' + verification.triggerSummary.canonicalMatchedTriggerCount + '개',
        '누락: ' + verification.triggerSummary.canonicalMissingTriggerCount + '개',
        '초과: ' + verification.triggerSummary.canonicalExcessTriggerCount + '개',
        '고아: ' + verification.triggerSummary.orphanTriggerCount + '개',
        '구형: ' + verification.triggerSummary.legacyTriggerCount + '개',
        '미분류: ' + verification.triggerSummary.unknownTriggerCount + '개',
        '중앙 복구요청: ' + verification.triggerSummary.repairRequestCount + '건',
        '활성 lease: ' + verification.activeLeases.length + '개',
        '재처리 대기: ' + verification.retryQueue.pendingCount + '건',
        '',
        verification.ok
          ? '정식 13개 중앙관리 구조가 정상입니다.'
          : verification.blockers.map(function(item) { return '- ' + item.message; }).join('\n')
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }

  return verification;
}


function AUTOMATION_buildCutoverVerification_(options) {
  options = options || {};
  var generatedAt = new Date().toISOString();
  var snapshot = TRG_buildStatusSnapshot_();

  if (
    TRG_isCanonicalSnapshotHealthy_(snapshot) &&
    Number(snapshot.summary.repairRequestCount || 0) > 0
  ) {
    TRG_clearCanonicalRepairRequest_();
    snapshot = TRG_buildStatusSnapshot_();
  }

  var blockers = [];
  var warnings = [];
  var exclude = options.ignoreCutoverLease
    ? [AUTOMATION_CUTOVER_CONFIG.cutoverLeaseModuleKey]
    : [];
  var leaseState = AUTOMATION_collectCutoverLeaseState_(exclude);
  var retryQueue = AUTOMATION_getCutoverRetryQueueSummary_();
  var coreLastRun = AUTOMATION_getCoreDataSyncLastRun();

  if (!TRG_isCanonicalSnapshotHealthy_(snapshot)) {
    blockers.push({
      code: 'CANONICAL_TRIGGER_MISMATCH',
      message: '정식 13개 트리거 계획과 현재 설치 상태가 일치하지 않습니다.'
    });
  }

  if (snapshot.summary.repairRequestCount > 0) {
    blockers.push({
      code: 'CANONICAL_REPAIR_REQUEST',
      message: '중앙 트리거 복구요청이 ' + snapshot.summary.repairRequestCount + '건 남아 있습니다.'
    });
  }

  if (leaseState.active.length > 0) {
    warnings.push({
      code: 'ACTIVE_AUTOMATION_LEASE',
      message: '현재 정상 실행 중인 자동화 lease가 ' + leaseState.active.length + '개 있습니다.'
    });
  }

  if (!retryQueue.schemaOk) {
    blockers.push({
      code: 'RETRY_QUEUE_SCHEMA_INVALID',
      message: retryQueue.error || '재처리 큐 구조가 올바르지 않습니다.'
    });
  }

  if (retryQueue.activeRunningCount > 0) {
    warnings.push({
      code: 'RETRY_QUEUE_RUNNING',
      message: '현재 실행 중인 재처리 작업이 ' + retryQueue.activeRunningCount + '건 있습니다.'
    });
  }

  if (retryQueue.pendingCount > 0) {
    warnings.push({
      code: 'RETRY_QUEUE_PENDING',
      message: '재처리 대기 작업이 ' + retryQueue.pendingCount + '건 있습니다.'
    });
  }

  return {
    version: AUTOMATION_CUTOVER_CONFIG.version,
    cutoverId: String(options.cutoverId || AUTOMATION_createCutoverId_()),
    generatedAt: generatedAt,
    ok: blockers.length === 0,
    blockers: blockers,
    warnings: warnings,
    triggerSummary: snapshot.summary,
    triggerSnapshot: snapshot,
    activeLeases: leaseState.active,
    staleLeases: leaseState.stale,
    retryQueue: retryQueue,
    coreLastRun: AUTOMATION_makeCutoverJsonSafe_(coreLastRun)
  };
}


/****************************************************
 * 전환 상태·이력·시트
 ****************************************************/

function AUTOMATION_saveCutoverState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    AUTOMATION_CUTOVER_CONFIG.statePropertyKey,
    JSON.stringify(AUTOMATION_compactCutoverState_(state))
  );
}


function AUTOMATION_storeCutoverPreflight_(report) {
  PropertiesService.getScriptProperties().setProperty(
    AUTOMATION_CUTOVER_CONFIG.lastPreflightPropertyKey,
    JSON.stringify(AUTOMATION_compactCutoverPreflight_(report))
  );
}


function AUTOMATION_appendCutoverHistory_(state) {
  var props = PropertiesService.getScriptProperties();
  var history = AUTOMATION_readCutoverJsonProperty_(
    AUTOMATION_CUTOVER_CONFIG.historyPropertyKey
  );

  if (!Array.isArray(history)) history = [];
  history.unshift(AUTOMATION_compactCutoverState_(state));
  history = history.slice(0, AUTOMATION_CUTOVER_CONFIG.historyLimit);

  props.setProperty(
    AUTOMATION_CUTOVER_CONFIG.historyPropertyKey,
    JSON.stringify(history)
  );
}


function AUTOMATION_appendCutoverLog_(entry) {
  try {
    var sheet = AUTOMATION_ensureCutoverLogSheet_();
    var triggerSummary = entry.triggerSummary || AUTOMATION_emptyTriggerSummary_();
    var retrySummary = entry.retrySummary || {};
    var startedMs = AUTOMATION_parseCutoverDateMs_(entry.startedAt);
    var finishedMs = AUTOMATION_parseCutoverDateMs_(entry.finishedAt);
    var durationSec = startedMs && finishedMs
      ? Math.max(0, finishedMs - startedMs) / 1000
      : '';

    var row = [[
      String(entry.cutoverId || ''),
      String(entry.phase || ''),
      String(entry.status || ''),
      String(entry.startedAt || ''),
      String(entry.finishedAt || ''),
      durationSec,
      Number(triggerSummary.installedTriggerCount || 0),
      Number(triggerSummary.canonicalPlannedTriggerCount || 0),
      Number(entry.deletedCount || 0),
      Number(entry.createdCount || 0),
      Number(triggerSummary.canonicalMatchedTriggerCount || 0),
      Number(triggerSummary.canonicalMissingTriggerCount || 0),
      Number(triggerSummary.canonicalExcessTriggerCount || 0),
      Number(triggerSummary.orphanTriggerCount || 0),
      Number(triggerSummary.legacyTriggerCount || 0),
      Number(triggerSummary.unknownTriggerCount || 0),
      Number(triggerSummary.repairRequestCount || 0),
      Number(entry.activeLeaseCount || 0),
      Number(retrySummary.pendingCount || 0),
      Number(retrySummary.failCount || 0),
      String(entry.coreStatus || ''),
      String(entry.message || '').slice(0, 2000),
      AUTOMATION_stringifyCutoverDetail_(entry.detail),
      AUTOMATION_CUTOVER_CONFIG.version
    ]];

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row[0].length).setValues(row);
    return true;
  } catch (err) {
    console.error('[AUTOMATION_appendCutoverLog_] ' + AUTOMATION_cutoverErrorMessage_(err), err);
    return false;
  }
}


function AUTOMATION_ensureCutoverLogSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var name = AUTOMATION_CUTOVER_CONFIG.statusSheetName;
  var sheet = ss.getSheetByName(name);
  var headers = [
    '전환ID', '단계', '상태', '시작일시', '종료일시', '소요초',
    '현재트리거', '정식계획', '삭제수', '설치수', '계획일치', '누락', '초과',
    '고아', '구형', '미분류', '복구요청', '활성lease', '재처리대기', '재처리실패',
    '핵심동기화', '메시지', '세부JSON', '버전'
  ];

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.hideSheet();
  } else {
    var current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    var mismatch = headers.some(function(header, index) {
      return String(current[index] || '') !== header;
    });

    if (mismatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }

  sheet.setColumnWidth(1, 210);
  sheet.setColumnWidth(2, 130);
  sheet.setColumnWidth(3, 150);
  sheet.setColumnWidth(4, 190);
  sheet.setColumnWidth(5, 190);
  sheet.setColumnWidth(21, 150);
  sheet.setColumnWidth(22, 420);
  sheet.setColumnWidth(23, 500);

  return sheet;
}


/****************************************************
 * UI·공통 유틸리티
 ****************************************************/

function AUTOMATION_showCutoverReadinessAlert_(report) {
  var lines = [
    '상태: ' + (report.ok ? '전환 가능' : '전환 차단'),
    '현재 설치형 트리거: ' + report.triggerSummary.installedTriggerCount + '개',
    '정식 계획: ' + report.triggerSummary.canonicalPlannedTriggerCount + '개',
    '계획 일치: ' + report.triggerSummary.canonicalMatchedTriggerCount + '개',
    '구형/개별: ' + report.triggerSummary.legacyTriggerCount + '개',
    '고아: ' + report.triggerSummary.orphanTriggerCount + '개',
    '미분류: ' + report.triggerSummary.unknownTriggerCount + '개',
    '활성 lease: ' + report.activeLeases.length + '개',
    '재처리 대기: ' + report.retryQueue.pendingCount + '건',
    '재처리 실패: ' + report.retryQueue.failCount + '건',
    '최근 핵심 동기화: ' + (report.coreLastRun ? String(report.coreLastRun.status || '') : '없음')
  ];

  if (report.blockers.length > 0) {
    lines.push('', '[차단 항목]');
    report.blockers.forEach(function(item) {
      lines.push('- ' + item.message);
    });
  }

  if (report.warnings.length > 0) {
    lines.push('', '[주의 항목]');
    report.warnings.forEach(function(item) {
      lines.push('- ' + item.message);
    });
  }

  SpreadsheetApp.getUi().alert(
    '자동화 전환 사전점검',
    lines.join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function AUTOMATION_cutoverReportMessage_(report) {
  var parts = [];

  if (report.blockers.length > 0) {
    parts.push('차단: ' + report.blockers.map(function(item) {
      return item.message;
    }).join(' / '));
  }

  if (report.warnings.length > 0) {
    parts.push('주의: ' + report.warnings.map(function(item) {
      return item.message;
    }).join(' / '));
  }

  if (parts.length === 0) {
    parts.push('정식 13개 트리거 전환 사전점검을 통과했습니다.');
  }

  return parts.join(' | ');
}


function AUTOMATION_compactCutoverState_(state) {
  state = state || {};
  var verification = state.verification || {};
  var reinstall = state.reinstall || {};
  var coreSync = state.coreSync || {};

  return {
    version: String(state.version || AUTOMATION_CUTOVER_CONFIG.version),
    cutoverId: String(state.cutoverId || ''),
    status: String(state.status || ''),
    phase: String(state.phase || ''),
    startedAt: String(state.startedAt || ''),
    finishedAt: String(state.finishedAt || ''),
    coreSync: {
      status: String(coreSync.status || ''),
      durationMs: Number(coreSync.durationMs || 0),
      successCount: Number(coreSync.successCount || 0),
      errorCount: Number(coreSync.errorCount || 0),
      fatalError: String(coreSync.fatalError || '').slice(0, 1000)
    },
    reinstall: {
      source: String(reinstall.source || ''),
      deletedCount: Number(reinstall.deletedCount || 0),
      createdCount: Number(reinstall.createdCount || 0)
    },
    verification: {
      ok: verification.ok === true,
      generatedAt: String(verification.generatedAt || ''),
      triggerSummary: AUTOMATION_makeCutoverJsonSafe_(verification.triggerSummary || null),
      blockerCount: Array.isArray(verification.blockers) ? verification.blockers.length : 0,
      warningCount: Array.isArray(verification.warnings) ? verification.warnings.length : 0
    },
    error: String(state.error || '').slice(0, 2000)
  };
}


function AUTOMATION_compactCutoverPreflight_(report) {
  report = report || {};
  return {
    version: String(report.version || AUTOMATION_CUTOVER_CONFIG.version),
    cutoverId: String(report.cutoverId || ''),
    generatedAt: String(report.generatedAt || ''),
    ok: report.ok === true,
    alreadyCanonicalHealthy: report.alreadyCanonicalHealthy === true,
    blockers: AUTOMATION_makeCutoverJsonSafe_(report.blockers || []),
    warnings: AUTOMATION_makeCutoverJsonSafe_(report.warnings || []),
    activeLeaseCount: Array.isArray(report.activeLeases) ? report.activeLeases.length : 0,
    staleLeaseCount: Array.isArray(report.staleLeases) ? report.staleLeases.length : 0,
    retryQueue: AUTOMATION_makeCutoverJsonSafe_(report.retryQueue || null),
    coreStatus: report.coreLastRun ? String(report.coreLastRun.status || '') : '',
    coreFresh: report.coreFresh === true,
    scriptTimezone: String(report.scriptTimezone || ''),
    triggerSummary: AUTOMATION_makeCutoverJsonSafe_(report.triggerSummary || null)
  };
}


function AUTOMATION_emptyTriggerSummary_() {
  return {
    installedTriggerCount: 0,
    canonicalPlannedTriggerCount: 13,
    canonicalMatchedTriggerCount: 0,
    canonicalMissingTriggerCount: 13,
    canonicalExcessTriggerCount: 0,
    orphanTriggerCount: 0,
    legacyTriggerCount: 0,
    unknownTriggerCount: 0,
    repairRequestCount: 0
  };
}


function AUTOMATION_createCutoverId_() {
  return 'CUTOVER-' + Utilities.formatDate(
    new Date(),
    TRG_MANAGER_CONFIG.timezone,
    'yyyyMMdd-HHmmss'
  ) + '-' + Utilities.getUuid().slice(0, 8);
}


function AUTOMATION_readCutoverJsonProperty_(propertyKey) {
  var raw = PropertiesService.getScriptProperties().getProperty(propertyKey);
  if (!raw) return null;
  return AUTOMATION_parseCutoverJson_(raw);
}


function AUTOMATION_parseCutoverJson_(raw) {
  if (!raw) return null;
  try {
    return JSON.parse(raw);
  } catch (ignoreJsonError) {
    return null;
  }
}


function AUTOMATION_parseCutoverDateMs_(value) {
  if (!value) return 0;
  var ms = new Date(value).getTime();
  return isFinite(ms) ? ms : 0;
}


function AUTOMATION_cutoverErrorMessage_(err) {
  if (!err) return '알 수 없는 오류';
  return String(err && err.message ? err.message : err);
}


function AUTOMATION_makeCutoverJsonSafe_(value) {
  if (value === null || typeof value === 'undefined') return null;

  try {
    return JSON.parse(JSON.stringify(value));
  } catch (ignoreSerializationError) {
    return String(value);
  }
}


function AUTOMATION_stringifyCutoverDetail_(detail) {
  var text = '';
  try {
    text = JSON.stringify(AUTOMATION_makeCutoverJsonSafe_(detail));
  } catch (ignoreDetailError) {
    text = String(detail || '');
  }

  if (text.length > AUTOMATION_CUTOVER_CONFIG.maxDetailLength) {
    text = text.slice(0, AUTOMATION_CUTOVER_CONFIG.maxDetailLength) + '...';
  }

  return text;
}
