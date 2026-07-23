/****************************************************
 * AutomationHealthMonitor.gs
 * 자동화 장애 감지·Discord 통지 - 13단계
 *
 * 감시 대상:
 * - 핵심 데이터 동기화 연속 실패·장시간 미실행
 * - 정식 13개 트리거 구조 불일치·중앙 복구요청
 * - 편집 재처리 큐 최종 실패·적체·멈춘 RUNNING
 * - 기능별 lease 장기 점유·손상
 * - 영업관리대장 백업 실패·장시간 미성공
 * - 메일 발송 실패 큐·발송파일 저장 큐 장애
 * - 자동 유지관리 오류
 *
 * 운영 원칙:
 * - 새 트리거를 만들지 않는다.
 * - 기존 핵심 5분 파이프라인과 Discord 1분 트리거가 제한적으로 호출한다.
 * - 동일 장애는 상태·지문·쿨다운으로 중복 폭주를 막는다.
 * - Discord 2xx 성공일 때만 알림 완료로 확정한다.
 * - 해결된 장애는 기존 장애 알림이 실제 전송된 경우에만 복구 알림을 보낸다.
 ****************************************************/

var AUTOMATION_HEALTH_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE13',

  moduleLeaseKey: 'HEALTH_MONITOR',
  moduleLeaseTtlMs: 2 * 60 * 1000,
  moduleLeaseWaitMs: 0,

  statePropertyKey: 'AUTOMATION_HEALTH_STATE_V1',
  issuePropertyPrefix: 'AUTOMATION_HEALTH_ISSUE_V1_',
  lastRunPropertyKey: 'AUTOMATION_HEALTH_LAST_RUN_V1',
  backupLastRunPropertyKey: 'AUTOMATION_BACKUP_LAST_RUN_V1',

  statusSheetName: '_자동화장애상태',
  logSheetName: '_자동화장애알림로그',

  defaultCheckIntervalMs: 10 * 60 * 1000,
  initialGraceMs: 30 * 60 * 1000,
  minimumRuntimeBudgetMs: 8 * 1000,
  maxRuntimeMs: 35 * 1000,

  coreFailureThreshold: 3,
  coreCriticalFailureThreshold: 6,
  coreStaleWarningMs: 20 * 60 * 1000,
  coreStaleCriticalMs: 60 * 60 * 1000,

  retryBacklogWarning: 10,
  retryBacklogError: 30,
  retryRunningStaleMs: 20 * 60 * 1000,

  mailArchiveBacklogWarning: 10,
  mailArchiveBacklogError: 25,
  mailArchiveRunningStaleMs: 20 * 60 * 1000,

  backupInitialGraceMs: 26 * 60 * 60 * 1000,
  backupStaleWarningMs: 26 * 60 * 60 * 1000,
  backupStaleCriticalMs: 48 * 60 * 60 * 1000,

  leaseSuspiciousAgeMs: 15 * 60 * 1000,

  repeatMsBySeverity: Object.freeze({
    CRITICAL: 60 * 60 * 1000,
    ERROR: 3 * 60 * 60 * 1000,
    WARNING: 12 * 60 * 60 * 1000
  }),

  maxDiscordContentChars: 1900,
  maxDiscordIssueItems: 10,
  maxIssueSummaryChars: 320,
  maxErrorChars: 1500,

  logRetentionDays: 365,
  logMaxRows: 5000,

  healthStatusHeaders: Object.freeze([
    '장애키', '심각도', '활성여부', '장애명', '현재요약',
    '최초감지', '최근감지', '최근알림', '알림횟수', '알림대기',
    '복구대기', '최근복구알림', '지문', '최근전송오류', '버전'
  ]),

  healthLogHeaders: Object.freeze([
    '기록일시', '전송ID', '종류', '장애키', '심각도', '장애명',
    '요약', 'Discord결과', 'HTTP코드', '출처', '버전'
  ])
});


/****************************************************
 * 공개 함수
 ****************************************************/

/**
 * 현재 장애 상태를 수집하되 Discord로는 보내지 않는다.
 */
function AUTOMATION_previewHealthStatus() {
  TRG_assertAutomationOwner_();

  var state = AUTOMATION_healthReadState_();
  var snapshot = AUTOMATION_healthCollectSnapshot_({
    source: 'MANUAL_PREVIEW',
    nowMs: Date.now(),
    state: state
  });
  var issues = AUTOMATION_healthEvaluateIssues_(snapshot, state);
  var previewState = AUTOMATION_healthJsonSafe_(state) || AUTOMATION_healthNewState_();
  AUTOMATION_healthReconcileIssues_(previewState, issues, Date.now());
  var preview = {
    status: 'PREVIEW',
    checkedAt: new Date().toISOString(),
    issueCount: issues.length,
    issues: issues,
    snapshot: snapshot
  };

  AUTOMATION_healthWriteStatusSheet_(previewState, issues, preview);
  AUTOMATION_healthShowPreviewAlert_(preview);
  return preview;
}


/**
 * 정기 점검 주기와 무관하게 현재 장애를 즉시 재평가하고,
 * 중복 방지·재알림 주기상 필요한 Discord 알림만 보낸다.
 */
function AUTOMATION_runHealthMonitorNow() {
  TRG_assertAutomationOwner_();

  var result = AUTOMATION_runHealthMonitorSafe_({
    source: 'MANUAL_RUN',
    force: true,
    deadlineMs: Date.now() + 4.5 * 60 * 1000
  });

  try {
    SpreadsheetApp.getUi().alert(
      '자동화 장애 점검',
      [
        '상태: ' + String(result.status || ''),
        '현재 장애: ' + Number(result.activeIssueCount || 0) + '건',
        '전송 대상: ' + Number(result.alertCandidateCount || 0) + '건',
        '전송 성공: ' + Number(result.sentCount || 0) + '건',
        '복구 알림: ' + Number(result.recoverySentCount || 0) + '건',
        result.error ? ('오류: ' + result.error) : ''
      ].filter(Boolean).join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {
    // 백그라운드/편집기 환경에서는 반환값과 로그로 확인한다.
  }

  return result;
}


function AUTOMATION_showHealthStatusSheet() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss.getSheetByName(AUTOMATION_HEALTH_CONFIG.statusSheetName);

  if (!sheet) {
    AUTOMATION_previewHealthStatus();
    sheet = ss.getSheetByName(AUTOMATION_HEALTH_CONFIG.statusSheetName);
  }

  if (!sheet) throw new Error('자동화 장애 상태 시트를 생성하지 못했습니다.');

  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


function AUTOMATION_showHealthAlertLogSheet() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = AUTOMATION_healthGetOrCreateLogSheet_();
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return sheet.getName();
}


/**
 * 백업 실행 함수가 성공·실패·스킵 결과를 기록할 때 사용한다.
 */
function AUTOMATION_recordBackupExecution_(result) {
  result = result || {};

  var nowIso = new Date().toISOString();
  var payload = {
    version: AUTOMATION_HEALTH_CONFIG.version,
    recordedAt: nowIso,
    status: String(result.status || 'UNKNOWN'),
    successAt: String(result.successAt || ''),
    fileName: String(result.fileName || ''),
    fileId: String(result.fileId || ''),
    fileUrl: String(result.fileUrl || ''),
    error: AUTOMATION_healthLimitText_(result.error || result.message || '', 1200)
  };

  if (payload.status === 'SUCCESS' && !payload.successAt) {
    payload.successAt = nowIso;
  }

  var props = PropertiesService.getScriptProperties();
  var previous = AUTOMATION_healthReadJsonProperty_(
    props,
    AUTOMATION_HEALTH_CONFIG.backupLastRunPropertyKey
  ) || {};

  if (!payload.successAt && previous.successAt) {
    payload.successAt = String(previous.successAt || '');
    payload.fileName = payload.fileName || String(previous.fileName || '');
    payload.fileId = payload.fileId || String(previous.fileId || '');
    payload.fileUrl = payload.fileUrl || String(previous.fileUrl || '');
  }

  props.setProperty(
    AUTOMATION_HEALTH_CONFIG.backupLastRunPropertyKey,
    JSON.stringify(payload)
  );

  return payload;
}


/****************************************************
 * 안전 진입점
 ****************************************************/

function AUTOMATION_runHealthMonitorSafe_(options) {
  options = options || {};

  try {
    return AUTOMATION_runHealthMonitor_(options);
  } catch (err) {
    var failed = {
      status: 'ERROR',
      source: String(options.source || ''),
      checkedAt: new Date().toISOString(),
      activeIssueCount: 0,
      alertCandidateCount: 0,
      sentCount: 0,
      recoverySentCount: 0,
      error: AUTOMATION_healthErrorMessage_(err)
    };

    AUTOMATION_healthRecordLastRun_(failed);
    console.error('[AUTOMATION_runHealthMonitorSafe_] ' + failed.error, err);
    return failed;
  }
}


function AUTOMATION_runHealthMonitor_(options) {
  options = options || {};

  var ownerEmail = String(
    typeof TRG_MANAGER_CONFIG !== 'undefined' && TRG_MANAGER_CONFIG
      ? TRG_MANAGER_CONFIG.automationOwnerEmail
      : 'bang@s1samsung.com'
  ).trim().toLowerCase();
  var effectiveEmail = String(
    typeof TRG_getEffectiveUserEmail_ === 'function'
      ? TRG_getEffectiveUserEmail_()
      : (Session.getEffectiveUser().getEmail() || '')
  ).trim().toLowerCase();

  if (effectiveEmail && ownerEmail && effectiveEmail !== ownerEmail) {
    return {
      status: 'SKIPPED_NOT_OWNER',
      source: String(options.source || ''),
      ownerEmail: ownerEmail,
      effectiveEmail: effectiveEmail,
      activeIssueCount: 0,
      alertCandidateCount: 0,
      sentCount: 0,
      recoverySentCount: 0
    };
  }

  var nowMs = Date.now();
  var deadlineMs = Number(options.deadlineMs || (nowMs + AUTOMATION_HEALTH_CONFIG.maxRuntimeMs));

  if (deadlineMs - nowMs < AUTOMATION_HEALTH_CONFIG.minimumRuntimeBudgetMs) {
    return {
      status: 'SKIPPED_INSUFFICIENT_TIME',
      source: String(options.source || ''),
      activeIssueCount: 0,
      alertCandidateCount: 0,
      sentCount: 0,
      recoverySentCount: 0
    };
  }

  var lease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_HEALTH_CONFIG.moduleLeaseKey,
    {
      taskName: '자동화 장애 감시',
      ttlMs: AUTOMATION_HEALTH_CONFIG.moduleLeaseTtlMs,
      waitMs: AUTOMATION_HEALTH_CONFIG.moduleLeaseWaitMs
    }
  );

  if (!lease.acquired) {
    return {
      status: lease.reason === 'CUTOVER_IN_PROGRESS'
        ? 'SKIPPED_CUTOVER_IN_PROGRESS'
        : 'SKIPPED_ALREADY_RUNNING',
      source: String(options.source || ''),
      activeIssueCount: 0,
      alertCandidateCount: 0,
      sentCount: 0,
      recoverySentCount: 0,
      reason: String(lease.reason || 'LEASE_BUSY')
    };
  }

  try {
    var state = AUTOMATION_healthReadState_();
    var currentCoreSummary = options.currentCoreSummary || null;
    var forceDue = options.force === true || AUTOMATION_healthIsCoreSummaryUnhealthy_(currentCoreSummary);
    var lastCheckedMs = AUTOMATION_healthToTimeMs_(state.lastCheckedAt);

    if (
      !forceDue &&
      lastCheckedMs > 0 &&
      nowMs - lastCheckedMs < AUTOMATION_HEALTH_CONFIG.defaultCheckIntervalMs
    ) {
      return {
        status: 'SKIPPED_NOT_DUE',
        source: String(options.source || ''),
        lastCheckedAt: String(state.lastCheckedAt || ''),
        activeIssueCount: AUTOMATION_healthCountActiveIssues_(state),
        alertCandidateCount: 0,
        sentCount: 0,
        recoverySentCount: 0
      };
    }

    if (deadlineMs - Date.now() < AUTOMATION_HEALTH_CONFIG.minimumRuntimeBudgetMs) {
      return {
        status: 'SKIPPED_INSUFFICIENT_TIME',
        source: String(options.source || ''),
        activeIssueCount: AUTOMATION_healthCountActiveIssues_(state),
        alertCandidateCount: 0,
        sentCount: 0,
        recoverySentCount: 0
      };
    }

    var snapshot = AUTOMATION_healthCollectSnapshot_({
      source: String(options.source || ''),
      nowMs: nowMs,
      deadlineMs: deadlineMs,
      state: state,
      currentCoreSummary: currentCoreSummary
    });
    var issues = AUTOMATION_healthEvaluateIssues_(snapshot, state);
    var reconciliation = AUTOMATION_healthReconcileIssues_(state, issues, nowMs);
    var sendResult = AUTOMATION_healthSendPendingAlerts_(
      reconciliation,
      snapshot,
      String(options.source || ''),
      deadlineMs
    );

    state.lastCheckedAt = new Date(nowMs).toISOString();
    state.lastSource = String(options.source || '');
    state.metrics = snapshot.metrics;
    state.lastSnapshotSummary = AUTOMATION_healthMakeSnapshotSummary_(snapshot);
    state.lastSendError = sendResult.success === false
      ? AUTOMATION_healthLimitText_(sendResult.error || '', 1200)
      : '';

    AUTOMATION_healthWriteState_(state);

    var result = {
      status: sendResult.success === false
        ? 'ALERT_SEND_FAILED_RETRY_PENDING'
        : (reconciliation.pending.length
          ? 'ALERTS_PROCESSED'
          : (issues.length ? 'ACTIVE_ISSUES_NO_NEW_ALERT' : 'HEALTHY_NO_ALERTS')),
      source: String(options.source || ''),
      checkedAt: state.lastCheckedAt,
      activeIssueCount: issues.length,
      alertCandidateCount: reconciliation.pending.length,
      selectedAlertCount: Number(sendResult.selectedCount || 0),
      sentCount: Number(sendResult.sentAlertCount || 0),
      recoverySentCount: Number(sendResult.sentRecoveryCount || 0),
      pendingAfterSendCount: Number(sendResult.remainingCount || 0),
      responseCode: Number(sendResult.responseCode || 0),
      error: sendResult.success === false ? String(sendResult.error || '') : '',
      issues: issues.map(AUTOMATION_healthIssueJsonSafe_)
    };

    AUTOMATION_healthRecordLastRun_(result);
    AUTOMATION_healthWriteStatusSheet_(state, issues, result);
    AUTOMATION_healthTrimLogSheet_();
    return result;
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


/****************************************************
 * 상태 수집
 ****************************************************/

function AUTOMATION_healthCollectSnapshot_(context) {
  context = context || {};

  var nowMs = Number(context.nowMs || Date.now());
  var state = context.state || AUTOMATION_healthNewState_();
  var snapshot = {
    checkedAt: new Date(nowMs).toISOString(),
    source: String(context.source || ''),
    metrics: state.metrics && typeof state.metrics === 'object'
      ? state.metrics
      : {},
    core: {},
    triggers: {},
    retryQueue: {},
    leases: {},
    backup: {},
    mail: {},
    maintenance: {},
    errors: []
  };

  try {
    snapshot.core = AUTOMATION_healthCollectCore_(
      context.currentCoreSummary,
      snapshot.metrics,
      nowMs,
      state
    );
  } catch (err) {
    snapshot.errors.push('핵심동기화 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  try {
    snapshot.triggers = AUTOMATION_healthCollectTriggers_();
  } catch (err) {
    snapshot.errors.push('트리거 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  try {
    snapshot.retryQueue = AUTOMATION_healthCollectRetryQueue_(nowMs);
    snapshot.retryQueue.previousActive = Number(snapshot.metrics.retryQueueActive || 0);
    snapshot.retryQueue.increase = Number(snapshot.retryQueue.active || 0) - snapshot.retryQueue.previousActive;
    snapshot.metrics.retryQueueActive = Number(snapshot.retryQueue.active || 0);
  } catch (err) {
    snapshot.errors.push('재처리 큐 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  try {
    snapshot.leases = AUTOMATION_healthCollectLeases_(nowMs);
  } catch (err) {
    snapshot.errors.push('lease 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  try {
    snapshot.backup = AUTOMATION_healthCollectBackup_(nowMs, state);
  } catch (err) {
    snapshot.errors.push('백업 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  try {
    snapshot.mail = AUTOMATION_healthCollectMailQueues_(nowMs);
    snapshot.mail.previousSendFailureUnresolved = Number(snapshot.metrics.mailSendFailureUnresolved || 0);
    snapshot.mail.previousArchiveUnresolved = Number(snapshot.metrics.mailArchiveUnresolved || 0);
    snapshot.mail.sendFailureIncrease = Number(snapshot.mail.sendFailure && snapshot.mail.sendFailure.unresolved || 0) - snapshot.mail.previousSendFailureUnresolved;
    snapshot.mail.archiveIncrease = Number(snapshot.mail.archive && snapshot.mail.archive.unresolved || 0) - snapshot.mail.previousArchiveUnresolved;
    snapshot.metrics.mailSendFailureUnresolved = Number(snapshot.mail.sendFailure && snapshot.mail.sendFailure.unresolved || 0);
    snapshot.metrics.mailArchiveUnresolved = Number(snapshot.mail.archive && snapshot.mail.archive.unresolved || 0);
  } catch (err) {
    snapshot.errors.push('메일 큐 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  try {
    snapshot.maintenance = AUTOMATION_healthCollectMaintenance_();
  } catch (err) {
    snapshot.errors.push('유지관리 상태 수집: ' + AUTOMATION_healthErrorMessage_(err));
  }

  return snapshot;
}


function AUTOMATION_healthCollectCore_(currentSummary, metrics, nowMs, state) {
  var props = PropertiesService.getScriptProperties();
  var summary = currentSummary || AUTOMATION_healthReadJsonProperty_(
    props,
    typeof AUTOMATION_CORE_PIPELINE_CONFIG !== 'undefined'
      ? AUTOMATION_CORE_PIPELINE_CONFIG.lastRunPropertyKey
      : 'AUTOMATION_CORE_SYNC_LAST_RUN_V1'
  ) || {};

  var runId = String(summary.runToken || summary.finishedAt || summary.startedAt || '');
  var previousRunId = String(metrics.coreLastRunId || '');
  var status = String(summary.status || '');
  var finishedAtMs = AUTOMATION_healthToTimeMs_(summary.finishedAt || summary.startedAt);
  var failureStreak = Number(metrics.coreFailureStreak || 0);
  var lastSuccessAt = String(metrics.coreLastSuccessAt || '');

  if (runId && runId !== previousRunId) {
    if (AUTOMATION_healthIsCoreSuccessStatus_(status)) {
      failureStreak = 0;
      lastSuccessAt = summary.finishedAt || summary.startedAt || new Date(nowMs).toISOString();
    } else if (!AUTOMATION_healthIsCoreNeutralStatus_(status)) {
      failureStreak += 1;
    }

    metrics.coreLastRunId = runId;
    metrics.coreFailureStreak = failureStreak;
    metrics.coreLastStatus = status;
    metrics.coreLastFinishedAt = String(summary.finishedAt || summary.startedAt || '');
    metrics.coreLastSuccessAt = lastSuccessAt;
  }

  var referenceAtMs = finishedAtMs || AUTOMATION_healthToTimeMs_(metrics.coreLastFinishedAt);
  var lastSuccessAtMs = AUTOMATION_healthToTimeMs_(lastSuccessAt);
  var initializedAtMs = AUTOMATION_healthToTimeMs_(state.initializedAt);

  return {
    status: status,
    runId: runId,
    finishedAt: String(summary.finishedAt || summary.startedAt || ''),
    finishedAtMs: referenceAtMs,
    ageMs: referenceAtMs ? Math.max(0, nowMs - referenceAtMs) : 0,
    failureStreak: failureStreak,
    lastSuccessAt: lastSuccessAt,
    lastSuccessAgeMs: lastSuccessAtMs ? Math.max(0, nowMs - lastSuccessAtMs) : 0,
    withinInitialGrace: initializedAtMs > 0 && nowMs - initializedAtMs < AUTOMATION_HEALTH_CONFIG.backupInitialGraceMs,
    fatalError: AUTOMATION_healthLimitText_(summary.fatalError || '', 600),
    errorCount: Number(summary.errorCount || 0),
    stages: Array.isArray(summary.stages) ? summary.stages : []
  };
}


function AUTOMATION_healthCollectTriggers_() {
  if (typeof TRG_buildStatusSnapshot_ !== 'function') {
    throw new Error('TRG_buildStatusSnapshot_ 함수를 찾을 수 없습니다.');
  }

  var snapshot = TRG_buildStatusSnapshot_();
  var summary = snapshot.summary || {};

  return {
    installed: Number(summary.installedTriggerCount || 0),
    planned: Number(summary.canonicalPlannedTriggerCount || 0),
    matched: Number(summary.canonicalMatchedTriggerCount || 0),
    missing: Number(summary.canonicalMissingTriggerCount || 0),
    excess: Number(summary.canonicalExcessTriggerCount || 0),
    orphan: Number(summary.orphanTriggerCount || 0),
    legacy: Number(summary.legacyTriggerCount || 0),
    unknown: Number(summary.unknownTriggerCount || 0),
    repairRequestCount: Number(summary.repairRequestCount || 0),
    repairRequestLastAt: String(summary.repairRequestLastAt || '')
  };
}


function AUTOMATION_healthCollectRetryQueue_(nowMs) {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss.getSheetByName(AUTOMATION_RUNTIME_CONFIG.retryQueueSheetName);
  var result = {
    total: 0,
    pending: 0,
    retry: 0,
    running: 0,
    done: 0,
    fail: 0,
    active: 0,
    staleRunning: 0,
    latestFailJobId: '',
    latestFailError: ''
  };

  if (!sheet || sheet.getLastRow() < 2) return result;

  var headers = sheet.getRange(1, 1, 1, AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders.length).getDisplayValues()[0];
  var index = AUTOMATION_makeHeaderIndex_(headers);
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(2, 1, lastRow - 1, AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders.length).getValues();
  result.total = values.length;

  values.forEach(function(row) {
    var status = String(row[index['상태'] - 1] || '').trim().toUpperCase();
    var jobId = String(row[index['작업ID'] - 1] || '');
    var lastAttemptMs = AUTOMATION_healthToTimeMs_(row[index['최근시도일시'] - 1]);

    if (status === 'PENDING') result.pending++;
    if (status === 'RETRY') result.retry++;
    if (status === 'RUNNING') {
      result.running++;
      if (lastAttemptMs && nowMs - lastAttemptMs >= AUTOMATION_HEALTH_CONFIG.retryRunningStaleMs) {
        result.staleRunning++;
      }
    }
    if (status === 'DONE') result.done++;
    if (status === 'FAIL') {
      result.fail++;
      result.latestFailJobId = jobId || result.latestFailJobId;
      result.latestFailError = AUTOMATION_healthLimitText_(
        row[index['최근오류'] - 1] || result.latestFailError,
        320
      );
    }
  });

  result.active = result.pending + result.retry + result.running;
  return result;
}


function AUTOMATION_healthCollectLeases_(nowMs) {
  var props = PropertiesService.getScriptProperties().getProperties();
  var prefix = AUTOMATION_RUNTIME_CONFIG.leasePropertyPrefix;
  var result = {
    total: 0,
    active: 0,
    expired: 0,
    corrupted: 0,
    suspiciousActive: 0,
    suspiciousKeys: [],
    corruptedKeys: []
  };

  Object.keys(props).forEach(function(key) {
    if (key.indexOf(prefix) !== 0) return;
    result.total++;

    var parsed;
    try {
      parsed = JSON.parse(props[key]);
    } catch (err) {
      result.corrupted++;
      result.corruptedKeys.push(key.substring(prefix.length));
      return;
    }

    if (!parsed || typeof parsed !== 'object' || !parsed.token) {
      result.corrupted++;
      result.corruptedKeys.push(key.substring(prefix.length));
      return;
    }

    var active = AUTOMATION_isLeaseActive_(parsed, nowMs);
    var startedAtMs = Number(parsed.startedAtMs || AUTOMATION_healthToTimeMs_(parsed.startedAt) || 0);

    if (active) {
      result.active++;
      if (startedAtMs && nowMs - startedAtMs >= AUTOMATION_HEALTH_CONFIG.leaseSuspiciousAgeMs) {
        result.suspiciousActive++;
        result.suspiciousKeys.push(key.substring(prefix.length));
      }
    } else {
      result.expired++;
    }
  });

  return result;
}


function AUTOMATION_healthCollectBackup_(nowMs, state) {
  var props = PropertiesService.getScriptProperties();
  var record = AUTOMATION_healthReadJsonProperty_(
    props,
    AUTOMATION_HEALTH_CONFIG.backupLastRunPropertyKey
  );

  if (!record || !record.successAt) {
    var discovered = AUTOMATION_healthDiscoverLatestBackup_();
    if (discovered && discovered.successAt) {
      if (record && record.status && record.status !== 'SUCCESS') {
        record = AUTOMATION_recordBackupExecution_({
          status: String(record.status || 'ERROR'),
          successAt: discovered.successAt,
          fileName: discovered.fileName,
          fileId: discovered.fileId,
          fileUrl: discovered.fileUrl,
          error: String(record.error || '')
        });
      } else {
        record = AUTOMATION_recordBackupExecution_(discovered);
      }
    }
  }

  record = record || {};
  var successAtMs = AUTOMATION_healthToTimeMs_(record.successAt);
  var recordedAtMs = AUTOMATION_healthToTimeMs_(record.recordedAt);
  var initializedAtMs = AUTOMATION_healthToTimeMs_(state.initializedAt);

  return {
    status: String(record.status || ''),
    successAt: String(record.successAt || ''),
    successAgeMs: successAtMs ? Math.max(0, nowMs - successAtMs) : 0,
    recordedAt: String(record.recordedAt || ''),
    recordedAgeMs: recordedAtMs ? Math.max(0, nowMs - recordedAtMs) : 0,
    fileName: String(record.fileName || ''),
    error: AUTOMATION_healthLimitText_(record.error || '', 500),
    missingHistory: !successAtMs,
    withinInitialGrace: initializedAtMs > 0 && nowMs - initializedAtMs < AUTOMATION_HEALTH_CONFIG.backupInitialGraceMs
  };
}


function AUTOMATION_healthDiscoverLatestBackup_() {
  if (
    typeof BACKUP_FOLDER_ID === 'undefined' ||
    !String(BACKUP_FOLDER_ID || '').trim() ||
    String(BACKUP_FOLDER_ID) === '여기에_영업관리대장_백업_폴더ID_입력'
  ) {
    return null;
  }

  var folder = DriveApp.getFolderById(String(BACKUP_FOLDER_ID));
  var files = folder.getFiles();
  var latest = null;
  var scanned = 0;

  while (files.hasNext() && scanned < 1000) {
    scanned++;
    var file = files.next();
    var name = String(file.getName() || '');

    if (
      typeof BACKUP_FILE_PREFIX !== 'undefined' &&
      BACKUP_FILE_PREFIX &&
      name.indexOf(String(BACKUP_FILE_PREFIX)) !== 0
    ) {
      continue;
    }

    var updated = file.getLastUpdated ? file.getLastUpdated() : file.getDateCreated();
    var updatedMs = AUTOMATION_healthToTimeMs_(updated);

    if (!latest || updatedMs > latest.updatedMs) {
      latest = {
        updatedMs: updatedMs,
        fileName: name,
        fileId: file.getId(),
        fileUrl: file.getUrl()
      };
    }
  }

  if (!latest) return null;

  return {
    status: 'SUCCESS',
    successAt: new Date(latest.updatedMs).toISOString(),
    fileName: latest.fileName,
    fileId: latest.fileId,
    fileUrl: latest.fileUrl,
    error: ''
  };
}


function AUTOMATION_healthCollectMailQueues_(nowMs) {
  var spreadsheetId = '';

  try {
    if (typeof CONFIG !== 'undefined' && CONFIG) {
      spreadsheetId = String(CONFIG.GENERATOR_SPREADSHEET_ID || CONFIG.MASTER_SPREADSHEET_ID || '').trim();
    }
  } catch (ignoreConfigError) {
    spreadsheetId = '';
  }

  if (!spreadsheetId) {
    throw new Error('메일 큐 스프레드시트 ID를 확인할 수 없습니다.');
  }

  var ss = SpreadsheetApp.openById(spreadsheetId);
  return {
    spreadsheetId: spreadsheetId,
    sendFailure: AUTOMATION_healthReadGenericQueue_(
      ss.getSheetByName('메일발송실패큐_DB'),
      {
        unresolvedStatuses: ['확인필요', 'RETRY', 'PENDING', 'QUEUED', 'RUNNING'],
        staleRunningMs: AUTOMATION_HEALTH_CONFIG.mailArchiveRunningStaleMs,
        nowMs: nowMs
      }
    ),
    archive: AUTOMATION_healthReadGenericQueue_(
      ss.getSheetByName(
        typeof getSentFileArchiveConfig_ === 'function'
          ? String(getSentFileArchiveConfig_().QUEUE_SHEET_NAME || '발송파일저장큐_DB')
          : '발송파일저장큐_DB'
      ),
      {
        unresolvedStatuses: ['PENDING', 'QUEUED', 'RETRY', 'RUNNING'],
        staleRunningMs: AUTOMATION_HEALTH_CONFIG.mailArchiveRunningStaleMs,
        nowMs: nowMs
      }
    )
  };
}


function AUTOMATION_healthReadGenericQueue_(sheet, options) {
  options = options || {};
  var result = {
    exists: !!sheet,
    total: 0,
    unresolved: 0,
    fail: 0,
    running: 0,
    staleRunning: 0,
    latestJobId: '',
    latestCompany: '',
    latestError: '',
    statusCounts: {}
  };

  if (!sheet || sheet.getLastRow() < 2) return result;

  var lastCol = Math.max(1, sheet.getLastColumn());
  var headers = sheet.getRange(1, 1, 1, lastCol).getDisplayValues()[0].map(function(value) {
    return String(value || '').trim();
  });
  var headerIndex = {};
  headers.forEach(function(header, index) {
    if (header) headerIndex[header] = index;
  });

  var statusIdx = AUTOMATION_healthFindHeaderIndex_(headerIndex, ['상태', 'status', 'Status']);
  var jobIdIdx = AUTOMATION_healthFindHeaderIndex_(headerIndex, ['작업ID', 'jobId', 'runId']);
  var companyIdx = AUTOMATION_healthFindHeaderIndex_(headerIndex, ['회사명', '고객사명', '건물명']);
  var errorIdx = AUTOMATION_healthFindHeaderIndex_(headerIndex, ['오류', '마지막오류', '최근오류', 'error']);
  var updatedIdx = AUTOMATION_healthFindHeaderIndex_(headerIndex, ['수정일시', '최근시도일시', '시작일시', '등록일시']);
  var lastRow = sheet.getLastRow();
  var values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
  var unresolvedMap = {};

  (options.unresolvedStatuses || []).forEach(function(status) {
    unresolvedMap[String(status || '').trim().toUpperCase()] = true;
  });

  result.total = values.length;

  values.forEach(function(row) {
    var status = statusIdx >= 0 ? String(row[statusIdx] || '').trim() : '';
    var normalizedStatus = status.toUpperCase();
    var updatedMs = updatedIdx >= 0 ? AUTOMATION_healthToTimeMs_(row[updatedIdx]) : 0;

    result.statusCounts[status || '(빈값)'] = (result.statusCounts[status || '(빈값)'] || 0) + 1;

    if (normalizedStatus === 'FAIL') result.fail++;
    if (unresolvedMap[normalizedStatus]) result.unresolved++;
    if (normalizedStatus === 'RUNNING') {
      result.running++;
      if (
        updatedMs &&
        Number(options.nowMs || Date.now()) - updatedMs >= Number(options.staleRunningMs || 0)
      ) {
        result.staleRunning++;
      }
    }

    if (normalizedStatus === 'FAIL' || unresolvedMap[normalizedStatus]) {
      if (jobIdIdx >= 0) result.latestJobId = String(row[jobIdIdx] || result.latestJobId);
      if (companyIdx >= 0) result.latestCompany = String(row[companyIdx] || result.latestCompany);
      if (errorIdx >= 0) {
        result.latestError = AUTOMATION_healthLimitText_(row[errorIdx] || result.latestError, 320);
      }
    }
  });

  return result;
}


function AUTOMATION_healthCollectMaintenance_() {
  var key = typeof AUTOMATION_MAINTENANCE_CONFIG !== 'undefined'
    ? AUTOMATION_MAINTENANCE_CONFIG.lastResultPropertyKey
    : 'AUTOMATION_MAINTENANCE_LAST_RESULT_V1';
  var result = AUTOMATION_healthReadJsonProperty_(PropertiesService.getScriptProperties(), key) || {};

  return {
    status: String(result.status || ''),
    finishedAt: String(result.finishedAt || ''),
    completedCycle: result.completedCycle === true,
    errorCount: Array.isArray(result.errors) ? result.errors.length : 0,
    latestError: AUTOMATION_healthLimitText_(
      Array.isArray(result.errors) ? result.errors.join(' | ') : (result.error || ''),
      600
    )
  };
}


/****************************************************
 * 장애 판정
 ****************************************************/

function AUTOMATION_healthEvaluateIssues_(snapshot, state) {
  var issues = [];
  var nowMs = AUTOMATION_healthToTimeMs_(snapshot.checkedAt) || Date.now();
  var core = snapshot.core || {};
  var trigger = snapshot.triggers || {};
  var retry = snapshot.retryQueue || {};
  var leases = snapshot.leases || {};
  var backup = snapshot.backup || {};
  var mail = snapshot.mail || {};
  var maintenance = snapshot.maintenance || {};

  if (Number(core.failureStreak || 0) >= AUTOMATION_HEALTH_CONFIG.coreFailureThreshold) {
    var coreSeverity = Number(core.failureStreak || 0) >= AUTOMATION_HEALTH_CONFIG.coreCriticalFailureThreshold
      ? 'CRITICAL'
      : 'ERROR';
    issues.push(AUTOMATION_healthIssue_(
      'CORE_SYNC_FAILURE_STREAK',
      coreSeverity,
      '핵심 데이터 동기화 연속 실패',
      '연속 실패 ' + Number(core.failureStreak || 0) + '회 / 최근 상태 ' + String(core.status || '확인불가'),
      coreSeverity + '|STREAK_' + (
        Number(core.failureStreak || 0) >= AUTOMATION_HEALTH_CONFIG.coreCriticalFailureThreshold ? '6_PLUS' : '3_TO_5'
      )
    ));
  }

  if (!core.withinInitialGrace) {
    var coreAgeMs = Number(core.ageMs || 0);

    if (!core.finishedAtMs || coreAgeMs >= AUTOMATION_HEALTH_CONFIG.coreStaleWarningMs) {
      var coreStaleCritical = !core.finishedAtMs || coreAgeMs >= AUTOMATION_HEALTH_CONFIG.coreStaleCriticalMs;
      issues.push(AUTOMATION_healthIssue_(
        'CORE_SYNC_STALE',
        coreStaleCritical ? 'CRITICAL' : 'ERROR',
        '핵심 동기화 장시간 미실행',
        core.finishedAtMs
          ? ('최근 실행 후 ' + AUTOMATION_healthFormatDuration_(coreAgeMs) + ' 경과')
          : '핵심 동기화 실행 이력을 확인할 수 없습니다.',
        coreStaleCritical ? 'CRITICAL' : 'ERROR'
      ));
    }
  }

  var triggerMismatch =
    Number(trigger.installed || 0) !== 13 ||
    Number(trigger.planned || 0) !== 13 ||
    Number(trigger.matched || 0) !== 13 ||
    Number(trigger.missing || 0) > 0 ||
    Number(trigger.excess || 0) > 0 ||
    Number(trigger.orphan || 0) > 0 ||
    Number(trigger.legacy || 0) > 0 ||
    Number(trigger.unknown || 0) > 0;

  if (triggerMismatch) {
    issues.push(AUTOMATION_healthIssue_(
      'TRIGGER_STRUCTURE_MISMATCH',
      'CRITICAL',
      '정식 13개 트리거 구조 이상',
      [
        '설치 ' + Number(trigger.installed || 0),
        '일치 ' + Number(trigger.matched || 0),
        '누락 ' + Number(trigger.missing || 0),
        '초과 ' + Number(trigger.excess || 0),
        '고아 ' + Number(trigger.orphan || 0),
        '구형 ' + Number(trigger.legacy || 0),
        '미분류 ' + Number(trigger.unknown || 0)
      ].join(' / '),
      [
        trigger.installed,
        trigger.matched,
        trigger.missing,
        trigger.excess,
        trigger.orphan,
        trigger.legacy,
        trigger.unknown
      ].join('|')
    ));
  }

  if (Number(trigger.repairRequestCount || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'TRIGGER_REPAIR_REQUEST',
      'ERROR',
      '정식 트리거 중앙 복구요청 발생',
      '복구요청 ' + Number(trigger.repairRequestCount || 0) + '종 / 최근 ' + String(trigger.repairRequestLastAt || '시각없음'),
      'REQUEST_KEYS_' + Number(trigger.repairRequestCount || 0)
    ));
  }

  if (Number(retry.fail || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'RETRY_QUEUE_FINAL_FAILURE',
      'ERROR',
      '편집 재처리 최종 실패',
      'FAIL ' + Number(retry.fail || 0) + '건' +
        (retry.latestFailJobId ? ' / 최근 작업 ' + retry.latestFailJobId : '') +
        (retry.latestFailError ? ' / ' + retry.latestFailError : ''),
      'FAIL_' + Number(retry.fail || 0) + '|' + String(retry.latestFailJobId || '')
    ));
  }

  if (
    Number(retry.active || 0) >= AUTOMATION_HEALTH_CONFIG.retryBacklogWarning ||
    Number(retry.increase || 0) >= AUTOMATION_HEALTH_CONFIG.retryBacklogWarning
  ) {
    var retrySeverity = Number(retry.active || 0) >= AUTOMATION_HEALTH_CONFIG.retryBacklogError
      ? 'ERROR'
      : 'WARNING';
    issues.push(AUTOMATION_healthIssue_(
      'RETRY_QUEUE_BACKLOG',
      retrySeverity,
      '편집 재처리 큐 적체',
      '활성 ' + Number(retry.active || 0) + '건 / 대기 ' + Number(retry.pending || 0) +
        ' / 재시도 ' + Number(retry.retry || 0) + ' / 실행중 ' + Number(retry.running || 0) +
        (Number(retry.increase || 0) > 0 ? ' / 직전대비 +' + Number(retry.increase || 0) : ''),
      retrySeverity + '|' + (retrySeverity === 'ERROR' ? '30_PLUS' : '10_TO_29')
    ));
  }

  if (Number(retry.staleRunning || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'RETRY_QUEUE_STALE_RUNNING',
      'ERROR',
      '재처리 큐 RUNNING 작업 정체',
      '20분 이상 RUNNING 상태 ' + Number(retry.staleRunning || 0) + '건',
      'STALE_' + Number(retry.staleRunning || 0)
    ));
  }

  if (Number(leases.corrupted || 0) > 0 || Number(leases.suspiciousActive || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'MODULE_LEASE_ABNORMAL',
      Number(leases.suspiciousActive || 0) > 0 ? 'ERROR' : 'WARNING',
      '기능별 lease 이상',
      '장기 점유 ' + Number(leases.suspiciousActive || 0) + '건 / 손상 ' + Number(leases.corrupted || 0) +
        (leases.suspiciousKeys && leases.suspiciousKeys.length
          ? ' / ' + leases.suspiciousKeys.slice(0, 5).join(', ')
          : ''),
      [
        Number(leases.suspiciousActive || 0),
        Number(leases.corrupted || 0),
        (leases.suspiciousKeys || []).slice(0, 5).join(',')
      ].join('|')
    ));
  }

  if (backup.status && backup.status !== 'SUCCESS' && backup.status.indexOf('SKIPPED_') !== 0) {
    issues.push(AUTOMATION_healthIssue_(
      'BACKUP_LAST_RUN_FAILED',
      'ERROR',
      '영업관리대장 최근 백업 실패',
      '상태 ' + String(backup.status || '') + (backup.error ? ' / ' + backup.error : ''),
      String(backup.status || '') + '|' + AUTOMATION_healthHash_(backup.error || '')
    ));
  }

  if (!backup.withinInitialGrace) {
    var backupAgeMs = Number(backup.successAgeMs || 0);

    if (backup.missingHistory || backupAgeMs >= AUTOMATION_HEALTH_CONFIG.backupStaleWarningMs) {
      var backupCritical = backup.missingHistory || backupAgeMs >= AUTOMATION_HEALTH_CONFIG.backupStaleCriticalMs;
      issues.push(AUTOMATION_healthIssue_(
        'BACKUP_STALE',
        backupCritical ? 'CRITICAL' : 'ERROR',
        '영업관리대장 백업 장시간 미성공',
        backup.missingHistory
          ? '정상 백업 이력을 확인할 수 없습니다.'
          : ('최근 성공 후 ' + AUTOMATION_healthFormatDuration_(backupAgeMs) + ' 경과 / ' + String(backup.fileName || '')),
        backupCritical ? 'CRITICAL' : 'ERROR'
      ));
    }
  }

  var sendFailure = mail.sendFailure || {};
  if (Number(sendFailure.unresolved || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'MAIL_SEND_FAILURE_QUEUE',
      'CRITICAL',
      '메일 발송 실패 확인 필요',
      '미처리 ' + Number(sendFailure.unresolved || 0) + '건' +
        (sendFailure.latestCompany ? ' / 최근 고객 ' + sendFailure.latestCompany : '') +
        (sendFailure.latestError ? ' / ' + sendFailure.latestError : ''),
      'UNRESOLVED_' + Number(sendFailure.unresolved || 0) + '|' + String(sendFailure.latestJobId || '')
    ));
  }

  var archive = mail.archive || {};
  if (Number(archive.fail || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'MAIL_ARCHIVE_QUEUE_FAILURE',
      'ERROR',
      '발송파일 저장큐 최종 실패',
      'FAIL ' + Number(archive.fail || 0) + '건' +
        (archive.latestCompany ? ' / 최근 고객 ' + archive.latestCompany : '') +
        (archive.latestError ? ' / ' + archive.latestError : ''),
      'FAIL_' + Number(archive.fail || 0) + '|' + String(archive.latestJobId || '')
    ));
  }

  if (
    Number(archive.unresolved || 0) >= AUTOMATION_HEALTH_CONFIG.mailArchiveBacklogWarning ||
    Number(mail.archiveIncrease || 0) >= AUTOMATION_HEALTH_CONFIG.mailArchiveBacklogWarning
  ) {
    var archiveSeverity = Number(archive.unresolved || 0) >= AUTOMATION_HEALTH_CONFIG.mailArchiveBacklogError
      ? 'ERROR'
      : 'WARNING';
    issues.push(AUTOMATION_healthIssue_(
      'MAIL_ARCHIVE_QUEUE_BACKLOG',
      archiveSeverity,
      '발송파일 저장큐 적체',
      '대기·재시도·실행중 ' + Number(archive.unresolved || 0) + '건' +
        (Number(mail.archiveIncrease || 0) > 0 ? ' / 직전대비 +' + Number(mail.archiveIncrease || 0) : ''),
      archiveSeverity + '|' + (archiveSeverity === 'ERROR' ? '25_PLUS' : '10_TO_24')
    ));
  }

  if (Number(archive.staleRunning || 0) > 0) {
    issues.push(AUTOMATION_healthIssue_(
      'MAIL_ARCHIVE_QUEUE_STALE_RUNNING',
      'ERROR',
      '발송파일 저장큐 RUNNING 정체',
      '20분 이상 RUNNING ' + Number(archive.staleRunning || 0) + '건',
      'STALE_' + Number(archive.staleRunning || 0)
    ));
  }

  if (
    maintenance.status === 'ERROR' ||
    maintenance.status === 'PARTIAL_ERROR' ||
    maintenance.status === 'COMPLETED_WITH_ERRORS'
  ) {
    issues.push(AUTOMATION_healthIssue_(
      'MAINTENANCE_ERROR',
      maintenance.status === 'ERROR' ? 'ERROR' : 'WARNING',
      '자동 유지관리 오류',
      '상태 ' + maintenance.status +
        (maintenance.latestError ? ' / ' + maintenance.latestError : ''),
      maintenance.status + '|' + AUTOMATION_healthHash_(maintenance.latestError || '')
    ));
  }

  if (snapshot.errors && snapshot.errors.length) {
    issues.push(AUTOMATION_healthIssue_(
      'HEALTH_COLLECTION_ERROR',
      'WARNING',
      '장애 감시 상태 수집 일부 실패',
      snapshot.errors.slice(0, 5).join(' | '),
      AUTOMATION_healthHash_(snapshot.errors.join('|'))
    ));
  }

  return issues.sort(AUTOMATION_healthIssueSort_);
}


function AUTOMATION_healthIssue_(key, severity, title, summary, fingerprint) {
  return {
    key: String(key || ''),
    severity: String(severity || 'WARNING'),
    title: String(title || ''),
    summary: AUTOMATION_healthLimitText_(summary || '', AUTOMATION_HEALTH_CONFIG.maxIssueSummaryChars),
    fingerprint: String(fingerprint || '')
  };
}


/****************************************************
 * 장애 상태·쿨다운·복구 처리
 ****************************************************/

function AUTOMATION_healthReconcileIssues_(state, currentIssues, nowMs) {
  state.issues = state.issues && typeof state.issues === 'object' ? state.issues : {};
  var currentMap = {};
  var pending = [];
  var nowIso = new Date(nowMs).toISOString();

  currentIssues.forEach(function(issue) {
    currentMap[issue.key] = issue;
    var previous = state.issues[issue.key] || {};
    var wasActive = previous.active === true;
    var fingerprintChanged = String(previous.fingerprint || '') !== String(issue.fingerprint || '');
    var severityChanged = String(previous.severity || '') !== String(issue.severity || '');
    var lastAlertMs = AUTOMATION_healthToTimeMs_(previous.lastAlertAt);
    var repeatMs = AUTOMATION_HEALTH_CONFIG.repeatMsBySeverity[issue.severity] || (12 * 60 * 60 * 1000);
    var alertDue =
      !wasActive ||
      !lastAlertMs ||
      fingerprintChanged ||
      severityChanged ||
      nowMs - lastAlertMs >= repeatMs;

    var next = {
      key: issue.key,
      severity: issue.severity,
      title: issue.title,
      summary: issue.summary,
      fingerprint: issue.fingerprint,
      active: true,
      firstDetectedAt: wasActive
        ? String(previous.firstDetectedAt || nowIso)
        : nowIso,
      lastDetectedAt: nowIso,
      lastAlertAt: String(previous.lastAlertAt || ''),
      alertCount: Number(previous.alertCount || 0),
      alertPending: Boolean(previous.alertPending) || alertDue,
      recoveryPending: false,
      lastRecoveryAlertAt: String(previous.lastRecoveryAlertAt || ''),
      lastSendError: String(previous.lastSendError || '')
    };

    state.issues[issue.key] = next;

    if (next.alertPending) {
      pending.push({
        kind: 'ALERT',
        issue: issue,
        state: next
      });
    }
  });

  Object.keys(state.issues).forEach(function(key) {
    if (currentMap[key]) return;

    var previous = state.issues[key];
    if (!previous) return;

    if (previous.active === true) {
      previous.active = false;
      previous.resolvedAt = nowIso;
      previous.lastDetectedAt = String(previous.lastDetectedAt || nowIso);
      previous.alertPending = false;
      previous.recoveryPending = Boolean(previous.lastAlertAt);
      previous.lastSendError = '';
    }

    if (previous.recoveryPending === true) {
      pending.push({
        kind: 'RECOVERY',
        issue: {
          key: key,
          severity: String(previous.severity || 'WARNING'),
          title: String(previous.title || key),
          summary: '정상 상태로 복구되었습니다.',
          fingerprint: String(previous.fingerprint || '')
        },
        state: previous
      });
    }
  });

  pending.sort(function(a, b) {
    if (a.kind !== b.kind) return a.kind === 'ALERT' ? -1 : 1;
    return AUTOMATION_healthIssueSort_(a.issue, b.issue);
  });

  return {
    state: state,
    currentIssues: currentIssues,
    pending: pending
  };
}


function AUTOMATION_healthSendPendingAlerts_(reconciliation, snapshot, source, deadlineMs) {
  var pending = reconciliation.pending || [];

  if (!pending.length) {
    return {
      success: true,
      selectedCount: 0,
      sentAlertCount: 0,
      sentRecoveryCount: 0,
      remainingCount: 0,
      responseCode: 0
    };
  }

  if (deadlineMs - Date.now() < AUTOMATION_HEALTH_CONFIG.minimumRuntimeBudgetMs) {
    return {
      success: false,
      selectedCount: 0,
      sentAlertCount: 0,
      sentRecoveryCount: 0,
      remainingCount: pending.length,
      responseCode: 0,
      error: '장애 알림 전송에 필요한 실행시간이 부족합니다.'
    };
  }

  var batch = AUTOMATION_healthBuildDiscordBatch_(pending, snapshot);
  var sendResult = AUTOMATION_healthSendDiscord_(batch.content);
  var nowIso = new Date().toISOString();
  var sendId = AUTOMATION_healthCreateId_('HEALTH');
  var sentAlertCount = 0;
  var sentRecoveryCount = 0;

  if (sendResult.success) {
    batch.items.forEach(function(item) {
      var issueState = reconciliation.state.issues[item.issue.key];
      if (!issueState) return;

      if (item.kind === 'ALERT') {
        issueState.lastAlertAt = nowIso;
        issueState.alertPending = false;
        issueState.alertCount = Number(issueState.alertCount || 0) + 1;
        issueState.lastSendError = '';
        sentAlertCount++;
      } else {
        issueState.recoveryPending = false;
        issueState.lastRecoveryAlertAt = nowIso;
        issueState.lastSendError = '';
        sentRecoveryCount++;
      }

      AUTOMATION_healthAppendLog_({
        recordedAt: nowIso,
        sendId: sendId,
        kind: item.kind,
        issue: item.issue,
        discordResult: 'SUCCESS',
        responseCode: sendResult.responseCode,
        source: source
      });
    });
  } else {
    batch.items.forEach(function(item) {
      var issueState = reconciliation.state.issues[item.issue.key];
      if (issueState) {
        issueState.lastSendError = AUTOMATION_healthLimitText_(sendResult.error || '', 1000);
      }

      AUTOMATION_healthAppendLog_({
        recordedAt: nowIso,
        sendId: sendId,
        kind: item.kind,
        issue: item.issue,
        discordResult: 'FAILED: ' + AUTOMATION_healthLimitText_(sendResult.error || '', 500),
        responseCode: sendResult.responseCode,
        source: source
      });
    });
  }

  return {
    success: sendResult.success,
    selectedCount: batch.items.length,
    sentAlertCount: sentAlertCount,
    sentRecoveryCount: sentRecoveryCount,
    remainingCount: pending.length - batch.items.length,
    responseCode: sendResult.responseCode,
    error: sendResult.error || ''
  };
}


/****************************************************
 * Discord 전송
 ****************************************************/

function AUTOMATION_healthBuildDiscordBatch_(pending, snapshot) {
  var selected = [];

  for (var i = 0; i < pending.length && selected.length < AUTOMATION_HEALTH_CONFIG.maxDiscordIssueItems; i++) {
    var candidate = selected.concat([pending[i]]);
    var remaining = pending.length - candidate.length;
    var content = AUTOMATION_healthBuildDiscordContent_(candidate, remaining, snapshot);

    if (
      selected.length > 0 &&
      content.length > AUTOMATION_HEALTH_CONFIG.maxDiscordContentChars
    ) {
      break;
    }

    selected = candidate;
  }

  if (!selected.length && pending.length) selected = [pending[0]];

  var finalContent = AUTOMATION_healthBuildDiscordContent_(
    selected,
    pending.length - selected.length,
    snapshot
  );

  if (finalContent.length > AUTOMATION_HEALTH_CONFIG.maxDiscordContentChars) {
    finalContent = finalContent.substring(0, AUTOMATION_HEALTH_CONFIG.maxDiscordContentChars - 3) + '...';
  }

  return {
    items: selected,
    content: finalContent
  };
}


function AUTOMATION_healthBuildDiscordContent_(items, remainingCount, snapshot) {
  var hasAlert = items.some(function(item) { return item.kind === 'ALERT'; });
  var hasRecovery = items.some(function(item) { return item.kind === 'RECOVERY'; });
  var lines = [];

  if (hasAlert && hasRecovery) {
    lines.push('🚨✅ **영업관리대장 자동화 장애·복구 알림**');
  } else if (hasAlert) {
    lines.push('🚨 **영업관리대장 자동화 장애 감지**');
  } else {
    lines.push('✅ **영업관리대장 자동화 복구 알림**');
  }

  lines.push('감지시각: ' + AUTOMATION_healthFormatDateTime_(snapshot.checkedAt));
  lines.push('');

  items.forEach(function(item) {
    var issue = item.issue;
    var icon = item.kind === 'RECOVERY'
      ? '✅'
      : AUTOMATION_healthSeverityIcon_(issue.severity);
    var prefix = item.kind === 'RECOVERY' ? '복구' : issue.severity;

    lines.push(icon + ' **[' + prefix + '] ' + issue.title + '**');
    lines.push('└ ' + AUTOMATION_healthLimitText_(issue.summary || '', AUTOMATION_HEALTH_CONFIG.maxIssueSummaryChars));
  });

  if (remainingCount > 0) {
    lines.push('');
    lines.push('※ 추가 알림 ' + remainingCount + '건은 다음 점검에서 이어서 전송됩니다.');
  }

  lines.push('');
  lines.push('확인: 영업관리대장 > 자동화 관리 > 장애 상태 열기');
  return lines.join('\n');
}


function AUTOMATION_healthSendDiscord_(content) {
  var webhookPropertyKey = typeof SALES_SUPPORT_ALERT_CONFIG !== 'undefined'
    ? SALES_SUPPORT_ALERT_CONFIG.WEBHOOK_PROP_KEY
    : 'SALES_SUPPORT_DISCORD_WEBHOOK_URL';
  var webhookUrl = String(
    PropertiesService.getScriptProperties().getProperty(webhookPropertyKey) || ''
  ).trim();

  if (!webhookUrl) {
    return {
      success: false,
      responseCode: 0,
      error: 'Discord 웹훅 URL이 Script Properties에 없습니다.'
    };
  }

  try {
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: content }),
      muteHttpExceptions: true
    });
    var code = Number(response.getResponseCode()) || 0;
    var body = String(response.getContentText() || '');

    if (code >= 200 && code < 300) {
      return {
        success: true,
        responseCode: code,
        error: ''
      };
    }

    return {
      success: false,
      responseCode: code,
      error: AUTOMATION_healthLimitText_('Discord HTTP ' + code + ': ' + body, 1000)
    };
  } catch (err) {
    return {
      success: false,
      responseCode: 0,
      error: AUTOMATION_healthErrorMessage_(err)
    };
  }
}


/****************************************************
 * 상태·로그 시트
 ****************************************************/

function AUTOMATION_healthReadState_() {
  var props = PropertiesService.getScriptProperties();
  var state = AUTOMATION_healthReadJsonProperty_(props, AUTOMATION_HEALTH_CONFIG.statePropertyKey);

  if (!state || typeof state !== 'object') {
    state = AUTOMATION_healthNewState_();
  }

  if (!state.initializedAt) state.initializedAt = new Date().toISOString();
  if (!state.metrics || typeof state.metrics !== 'object') state.metrics = {};
  state.issues = {};

  var allProperties = props.getProperties();
  Object.keys(allProperties).forEach(function(key) {
    if (key.indexOf(AUTOMATION_HEALTH_CONFIG.issuePropertyPrefix) !== 0) return;

    var issueKey = key.substring(AUTOMATION_HEALTH_CONFIG.issuePropertyPrefix.length);
    var issueState = AUTOMATION_healthReadJsonProperty_(props, key);

    if (!issueKey || !issueState || typeof issueState !== 'object') {
      props.deleteProperty(key);
      return;
    }

    state.issues[issueKey] = issueState;
  });

  state.version = AUTOMATION_HEALTH_CONFIG.version;
  return state;
}


function AUTOMATION_healthNewState_() {
  return {
    version: AUTOMATION_HEALTH_CONFIG.version,
    initializedAt: new Date().toISOString(),
    lastCheckedAt: '',
    lastSource: '',
    lastSendError: '',
    metrics: {},
    issues: {},
    lastSnapshotSummary: {}
  };
}


function AUTOMATION_healthWriteState_(state) {
  var props = PropertiesService.getScriptProperties();
  var safeState = AUTOMATION_healthJsonSafe_(state) || {};
  var issues = safeState.issues && typeof safeState.issues === 'object'
    ? safeState.issues
    : {};
  var rootState = {};

  Object.keys(safeState).forEach(function(key) {
    if (key === 'issues') return;
    rootState[key] = safeState[key];
  });

  props.setProperty(
    AUTOMATION_HEALTH_CONFIG.statePropertyKey,
    JSON.stringify(rootState)
  );

  var existingProperties = props.getProperties();
  var keep = {};

  Object.keys(issues).forEach(function(issueKey) {
    var propertyKey = AUTOMATION_HEALTH_CONFIG.issuePropertyPrefix + issueKey;
    keep[propertyKey] = true;
    props.setProperty(propertyKey, JSON.stringify(issues[issueKey]));
  });

  Object.keys(existingProperties).forEach(function(key) {
    if (key.indexOf(AUTOMATION_HEALTH_CONFIG.issuePropertyPrefix) !== 0) return;
    if (!keep[key]) props.deleteProperty(key);
  });
}


function AUTOMATION_healthRecordLastRun_(result) {
  PropertiesService.getScriptProperties().setProperty(
    AUTOMATION_HEALTH_CONFIG.lastRunPropertyKey,
    JSON.stringify(AUTOMATION_healthJsonSafe_(result))
  );
}


function AUTOMATION_healthWriteStatusSheet_(state, currentIssues, lastResult) {
  try {
    var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
    var name = AUTOMATION_HEALTH_CONFIG.statusSheetName;
    var sheet = ss.getSheetByName(name);
    var created = false;

    if (!sheet) {
      sheet = ss.insertSheet(name);
      created = true;
    }

    var headers = AUTOMATION_HEALTH_CONFIG.healthStatusHeaders;
    var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    var mismatch = headers.some(function(header, index) {
      return String(currentHeaders[index] || '') !== header;
    });

    if (mismatch) {
      sheet.clearContents();
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    } else if (sheet.getLastRow() > 1) {
      sheet.getRange(2, 1, sheet.getLastRow() - 1, headers.length).clearContent();
    }

    var issueKeys = Object.keys(state.issues || {}).sort(function(a, b) {
      var ia = state.issues[a] || {};
      var ib = state.issues[b] || {};
      if (ia.active !== ib.active) return ia.active ? -1 : 1;
      return AUTOMATION_healthSeverityRank_(ib.severity) - AUTOMATION_healthSeverityRank_(ia.severity);
    });
    var rows = issueKeys.map(function(key) {
      var issue = state.issues[key] || {};
      return [
        key,
        String(issue.severity || ''),
        issue.active === true ? '예' : '아니오',
        String(issue.title || ''),
        String(issue.summary || ''),
        String(issue.firstDetectedAt || ''),
        String(issue.lastDetectedAt || ''),
        String(issue.lastAlertAt || ''),
        Number(issue.alertCount || 0),
        issue.alertPending === true ? '예' : '아니오',
        issue.recoveryPending === true ? '예' : '아니오',
        String(issue.lastRecoveryAlertAt || ''),
        String(issue.fingerprint || ''),
        String(issue.lastSendError || ''),
        AUTOMATION_HEALTH_CONFIG.version
      ];
    });

    if (rows.length) {
      sheet.getRange(2, 1, rows.length, headers.length).setValues(rows);
    }

    sheet.getRange('P1').setValue('최근점검');
    sheet.getRange('Q1').setValue(String(lastResult && lastResult.checkedAt || state.lastCheckedAt || ''));
    sheet.getRange('P2').setValue('현재장애');
    sheet.getRange('Q2').setValue(Number(currentIssues && currentIssues.length || 0));
    sheet.getRange('P3').setValue('최근상태');
    sheet.getRange('Q3').setValue(String(lastResult && lastResult.status || ''));
    sheet.getRange('P4').setValue('최근전송오류');
    sheet.getRange('Q4').setValue(String(state.lastSendError || ''));

    if (created || mismatch) {
      sheet.autoResizeColumns(1, Math.max(headers.length, 17));
    }

    if (created) {
      try { sheet.hideSheet(); } catch (ignoreHideError) {}
    }
  } catch (err) {
    console.error('[AUTOMATION_healthWriteStatusSheet_] ' + AUTOMATION_healthErrorMessage_(err), err);
  }
}


function AUTOMATION_healthGetOrCreateLogSheet_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var name = AUTOMATION_HEALTH_CONFIG.logSheetName;
  var sheet = ss.getSheetByName(name);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }

  var headers = AUTOMATION_HEALTH_CONFIG.healthLogHeaders;
  var current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  var mismatch = headers.some(function(header, index) {
    return String(current[index] || '') !== header;
  });

  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  if (created) {
    try { sheet.hideSheet(); } catch (ignoreHideError) {}
  }

  return sheet;
}


function AUTOMATION_healthAppendLog_(entry) {
  try {
    var sheet = AUTOMATION_healthGetOrCreateLogSheet_();
    var issue = entry.issue || {};
    sheet.appendRow([
      String(entry.recordedAt || new Date().toISOString()),
      String(entry.sendId || ''),
      String(entry.kind || ''),
      String(issue.key || ''),
      String(issue.severity || ''),
      String(issue.title || ''),
      String(issue.summary || ''),
      String(entry.discordResult || ''),
      Number(entry.responseCode || 0),
      String(entry.source || ''),
      AUTOMATION_HEALTH_CONFIG.version
    ]);
  } catch (err) {
    console.error('[AUTOMATION_healthAppendLog_] ' + AUTOMATION_healthErrorMessage_(err), err);
  }
}


function AUTOMATION_healthTrimLogSheet_() {
  try {
    var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
    var sheet = ss.getSheetByName(AUTOMATION_HEALTH_CONFIG.logSheetName);
    if (!sheet || sheet.getLastRow() < 2) return;

    var lastRow = sheet.getLastRow();
    var values = sheet.getRange(2, 1, lastRow - 1, 1).getValues();
    var cutoffMs = Date.now() - AUTOMATION_HEALTH_CONFIG.logRetentionDays * 24 * 60 * 60 * 1000;
    var deleteCount = 0;

    for (var i = 0; i < values.length; i++) {
      var timeMs = AUTOMATION_healthToTimeMs_(values[i][0]);
      if (timeMs && timeMs < cutoffMs) deleteCount++;
      else break;
    }

    var remainingAfterAge = values.length - deleteCount;
    if (remainingAfterAge > AUTOMATION_HEALTH_CONFIG.logMaxRows) {
      deleteCount += remainingAfterAge - AUTOMATION_HEALTH_CONFIG.logMaxRows;
    }

    if (deleteCount > 0) sheet.deleteRows(2, deleteCount);
  } catch (err) {
    console.error('[AUTOMATION_healthTrimLogSheet_] ' + AUTOMATION_healthErrorMessage_(err), err);
  }
}


/****************************************************
 * 공통 보조
 ****************************************************/

function AUTOMATION_healthIsCoreSummaryUnhealthy_(summary) {
  if (!summary) return false;
  var status = String(summary.status || '');
  return !!status &&
    !AUTOMATION_healthIsCoreSuccessStatus_(status) &&
    !AUTOMATION_healthIsCoreNeutralStatus_(status);
}


function AUTOMATION_healthIsCoreSuccessStatus_(status) {
  return status === 'COMPLETED' || status === 'COMPLETED_WITH_PENDING_RETRIES';
}


function AUTOMATION_healthIsCoreNeutralStatus_(status) {
  return status === 'SKIPPED_ALREADY_RUNNING' || status === 'SKIPPED_CUTOVER_IN_PROGRESS';
}


function AUTOMATION_healthCountActiveIssues_(state) {
  return Object.keys(state.issues || {}).filter(function(key) {
    return state.issues[key] && state.issues[key].active === true;
  }).length;
}


function AUTOMATION_healthIssueSort_(a, b) {
  var severityDiff = AUTOMATION_healthSeverityRank_(b.severity) - AUTOMATION_healthSeverityRank_(a.severity);
  if (severityDiff !== 0) return severityDiff;
  return String(a.key || '').localeCompare(String(b.key || ''));
}


function AUTOMATION_healthSeverityRank_(severity) {
  var normalized = String(severity || '').toUpperCase();
  if (normalized === 'CRITICAL') return 3;
  if (normalized === 'ERROR') return 2;
  return 1;
}


function AUTOMATION_healthSeverityIcon_(severity) {
  var normalized = String(severity || '').toUpperCase();
  if (normalized === 'CRITICAL') return '🛑';
  if (normalized === 'ERROR') return '🚨';
  return '⚠️';
}


function AUTOMATION_healthFindHeaderIndex_(headerIndex, candidates) {
  for (var i = 0; i < candidates.length; i++) {
    var candidate = String(candidates[i] || '');
    if (Object.prototype.hasOwnProperty.call(headerIndex, candidate)) {
      return headerIndex[candidate];
    }
  }
  return -1;
}


function AUTOMATION_healthToTimeMs_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') return value.getTime();
  var parsed = new Date(value).getTime();
  return isFinite(parsed) ? parsed : 0;
}


function AUTOMATION_healthFormatDuration_(ms) {
  var value = Math.max(0, Number(ms || 0));
  var minutes = Math.floor(value / 60000);

  if (minutes < 60) return minutes + '분';

  var hours = Math.floor(minutes / 60);
  var remainMinutes = minutes % 60;
  if (hours < 48) return hours + '시간' + (remainMinutes ? ' ' + remainMinutes + '분' : '');

  var days = Math.floor(hours / 24);
  var remainHours = hours % 24;
  return days + '일' + (remainHours ? ' ' + remainHours + '시간' : '');
}


function AUTOMATION_healthFormatDateTime_(value) {
  var date = value instanceof Date ? value : new Date(value || Date.now());
  if (isNaN(date.getTime())) date = new Date();
  return Utilities.formatDate(date, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
}


function AUTOMATION_healthHash_(value) {
  var text = String(value || '');
  var hash = 2166136261;

  for (var i = 0; i < text.length; i++) {
    hash ^= text.charCodeAt(i);
    hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
  }

  return (hash >>> 0).toString(16);
}


function AUTOMATION_healthLimitText_(value, maxLength) {
  var text = String(value == null ? '' : value).replace(/\s+/g, ' ').trim();
  var limit = Math.max(1, Number(maxLength || 1));
  if (text.length <= limit) return text;
  if (limit <= 3) return text.substring(0, limit);
  return text.substring(0, limit - 3) + '...';
}


function AUTOMATION_healthErrorMessage_(err) {
  return AUTOMATION_healthLimitText_(
    err && err.stack ? err.stack : (err && err.message ? err.message : String(err || '')),
    AUTOMATION_HEALTH_CONFIG.maxErrorChars
  );
}


function AUTOMATION_healthCreateId_(prefix) {
  try {
    return String(prefix || 'HEALTH') + '-' + Utilities.getUuid();
  } catch (ignoreUuidError) {
    return String(prefix || 'HEALTH') + '-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
  }
}


function AUTOMATION_healthReadJsonProperty_(props, key) {
  var raw = props.getProperty(key);
  if (!raw) return null;

  try {
    var parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (err) {
    return null;
  }
}


function AUTOMATION_healthJsonSafe_(value) {
  try {
    return JSON.parse(JSON.stringify(value, function(key, item) {
      if (item instanceof Date) return item.toISOString();
      if (typeof item === 'function') return undefined;
      return item;
    }));
  } catch (err) {
    return {
      status: 'SERIALIZE_ERROR',
      value: String(value)
    };
  }
}


function AUTOMATION_healthIssueJsonSafe_(issue) {
  return {
    key: String(issue && issue.key || ''),
    severity: String(issue && issue.severity || ''),
    title: String(issue && issue.title || ''),
    summary: String(issue && issue.summary || ''),
    fingerprint: String(issue && issue.fingerprint || '')
  };
}


function AUTOMATION_healthMakeSnapshotSummary_(snapshot) {
  return {
    checkedAt: String(snapshot.checkedAt || ''),
    source: String(snapshot.source || ''),
    coreStatus: String(snapshot.core && snapshot.core.status || ''),
    coreFailureStreak: Number(snapshot.core && snapshot.core.failureStreak || 0),
    triggerInstalled: Number(snapshot.triggers && snapshot.triggers.installed || 0),
    triggerMatched: Number(snapshot.triggers && snapshot.triggers.matched || 0),
    retryActive: Number(snapshot.retryQueue && snapshot.retryQueue.active || 0),
    retryFail: Number(snapshot.retryQueue && snapshot.retryQueue.fail || 0),
    mailSendFailure: Number(snapshot.mail && snapshot.mail.sendFailure && snapshot.mail.sendFailure.unresolved || 0),
    mailArchiveFailure: Number(snapshot.mail && snapshot.mail.archive && snapshot.mail.archive.fail || 0),
    backupStatus: String(snapshot.backup && snapshot.backup.status || ''),
    backupSuccessAt: String(snapshot.backup && snapshot.backup.successAt || ''),
    collectionErrorCount: Array.isArray(snapshot.errors) ? snapshot.errors.length : 0
  };
}


function AUTOMATION_healthShowPreviewAlert_(preview) {
  var issues = preview.issues || [];
  var lines = [
    '현재 장애: ' + issues.length + '건'
  ];

  issues.slice(0, 12).forEach(function(issue) {
    lines.push('[' + issue.severity + '] ' + issue.title + ': ' + issue.summary);
  });

  if (issues.length > 12) {
    lines.push('... 외 ' + (issues.length - 12) + '건');
  }

  if (!issues.length) lines.push('감지된 자동화 장애가 없습니다.');

  try {
    SpreadsheetApp.getUi().alert(
      '자동화 장애 상태 미리보기',
      lines.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {
    Logger.log(lines.join('\n'));
  }
}
