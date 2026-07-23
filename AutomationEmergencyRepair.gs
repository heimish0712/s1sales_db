/****************************************************
 * AutomationEmergencyRepair.gs
 * PHASE16 - 운영 장애 긴급 복구
 *
 * 복구 대상:
 * 1) 구형/고아 트리거를 정식 13개 구조로 재구성
 * 2) Drive API v2 Invalid query로 FAIL 된 발송파일 저장큐 재처리 전환
 * 3) 하이웍스 wrongList 실패를 수신주소 오류로 구조화
 *
 * 주의:
 * - 소스 반영만으로 실행되지 않습니다.
 * - bang@s1samsung.com에서 명시적으로 실행해야 합니다.
 ****************************************************/

var AUTOMATION_EMERGENCY_REPAIR_CONFIG = Object.freeze({
  version: '2026-07-23-PHASE16',
  moduleLeaseKey: 'EMERGENCY_REPAIR',
  moduleLeaseTtlMs: 12 * 60 * 1000,
  moduleLeaseWaitMs: 1000,
  lastResultPropertyKey: 'AUTOMATION_EMERGENCY_REPAIR_LAST_RESULT_V1'
});


function AUTOMATION_previewEmergencyRemediation() {
  TRG_assertAutomationOwner_();

  var preview = AUTOMATION_buildEmergencyRemediationPreview_();
  var lines = [
    '현재 설치형 트리거: ' + preview.triggers.installed + '개',
    '정식 계획 일치: ' + preview.triggers.matched + '/13개',
    '누락: ' + preview.triggers.missing + '개',
    '고아: ' + preview.triggers.orphan + '개',
    '구형: ' + preview.triggers.legacy + '개',
    '',
    'Drive Invalid query 실패: ' + preview.archiveQueue.candidates + '건',
    '하이웍스 수신주소 재분류: ' + preview.mailFailureQueue.candidates + '건',
    '',
    '실제 데이터와 트리거는 변경하지 않았습니다.'
  ];

  try {
    SpreadsheetApp.getUi().alert(
      '긴급 장애 복구 미리보기',
      lines.join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {}

  return preview;
}


function AUTOMATION_executeEmergencyRemediation() {
  TRG_assertAutomationOwner_();

  var preview = AUTOMATION_buildEmergencyRemediationPreview_();
  var ui = null;
  try { ui = SpreadsheetApp.getUi(); } catch (ignoreUiError) {}

  if (ui) {
    var response = ui.alert(
      '긴급 장애 복구 실행',
      [
        '다음 작업을 실행합니다.',
        '',
        '1. 현재 설치형 트리거 ' + preview.triggers.installed + '개를 삭제하고 정식 13개로 재구성',
        '2. Drive Invalid query 최종실패 ' + preview.archiveQueue.candidates + '건을 RETRY로 전환',
        '3. 하이웍스 잘못된 수신주소 ' + preview.mailFailureQueue.candidates + '건을 수신주소확인으로 재분류',
        '',
        '단순 onEdit/onOpen과 웹앱 doGet/doPost는 삭제되지 않습니다.',
        '계속하시겠습니까?'
      ].join('\n'),
      ui.ButtonSet.YES_NO
    );

    if (response !== ui.Button.YES) {
      return { ok: false, cancelled: true, preview: preview };
    }
  }

  var emergencyLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_EMERGENCY_REPAIR_CONFIG.moduleLeaseKey,
    {
      taskName: 'AUTOMATION_executeEmergencyRemediation',
      ttlMs: AUTOMATION_EMERGENCY_REPAIR_CONFIG.moduleLeaseTtlMs,
      waitMs: AUTOMATION_EMERGENCY_REPAIR_CONFIG.moduleLeaseWaitMs
    }
  );

  if (!emergencyLease.acquired) {
    throw new Error('다른 긴급 복구 작업이 실행 중입니다: ' + String(emergencyLease.reason || 'LEASE_BUSY'));
  }

  var result = {
    ok: false,
    version: AUTOMATION_EMERGENCY_REPAIR_CONFIG.version,
    startedAt: new Date().toISOString(),
    finishedAt: '',
    preview: preview,
    mailFailureRepair: null,
    archiveQueueRepair: null,
    triggerRepair: null,
    verification: null,
    firstArchiveQueueRun: null,
    error: ''
  };

  try {
    // 1) 일반 큐 데이터부터 복구합니다. 트리거 재구성 전에도 안전하게 수행 가능합니다.
    result.mailFailureRepair = MAILOPS_reclassifyLegacyInvalidRecipientFailures_({
      maxRows: 5000
    });

    result.archiveQueueRepair = MAILOPS_requeueLegacyDriveQueryArchiveFailures_({
      force: true,
      maxRows: 20000
    });

    // 2) 구형/고아 트리거를 정식 13개로 교체합니다.
    var currentSnapshot = TRG_buildStatusSnapshot_();
    if (TRG_isCanonicalSnapshotHealthy_(currentSnapshot)) {
      result.triggerRepair = {
        skipped: true,
        reason: 'ALREADY_CANONICAL',
        installedCount: Number(currentSnapshot.summary.installedTriggerCount || 0)
      };
    } else {
      var cutoverLease = AUTOMATION_acquireModuleLease_(
        'TRIGGER_CUTOVER',
        {
          taskName: '긴급 정식 13개 트리거 재구성',
          ttlMs: 12 * 60 * 1000,
          waitMs: 1000
        }
      );

      if (!cutoverLease.acquired) {
        throw new Error('트리거 전환 가드를 획득하지 못했습니다: ' + String(cutoverLease.reason || 'LEASE_BUSY'));
      }

      try {
        var leaseState = AUTOMATION_collectCutoverLeaseState_([
          AUTOMATION_EMERGENCY_REPAIR_CONFIG.moduleLeaseKey,
          'TRIGGER_CUTOVER',
          'HEALTH_MONITOR',
          'DISCORD_SALES_SUPPORT'
        ]);

        if (leaseState.active.length > 0) {
          throw new Error(
            '현재 실행 중인 자동화가 있어 트리거 재구성을 중단했습니다: ' +
            leaseState.active.map(function(item) {
              return item.moduleKey + (item.taskName ? '(' + item.taskName + ')' : '');
            }).join(', ')
          );
        }

        var preflight = TRG_preflightCanonicalInstall_();
        result.triggerRepair = TRG_reinstallCanonicalInternal_({
          preflight: preflight,
          currentTriggers: ScriptApp.getProjectTriggers(),
          source: 'AUTOMATION_executeEmergencyRemediation'
        });
      } finally {
        AUTOMATION_releaseModuleLease_(cutoverLease);
      }
    }

    result.verification = TRG_buildStatusSnapshot_().summary;

    // 3) 복구된 저장큐를 한 배치 즉시 처리합니다. 나머지는 정식 5분 트리거가 이어서 처리합니다.
    try {
      result.firstArchiveQueueRun = processDeferredSentFileArchiveQueueV94();
    } catch (archiveRunError) {
      result.firstArchiveQueueRun = {
        ok: false,
        error: String(archiveRunError && archiveRunError.message || archiveRunError)
      };
    }

    result.ok = !!(
      result.verification &&
      Number(result.verification.installedTriggerCount || 0) === 13 &&
      Number(result.verification.canonicalMatchedTriggerCount || 0) === 13 &&
      Number(result.verification.canonicalMissingTriggerCount || 0) === 0 &&
      Number(result.verification.orphanTriggerCount || 0) === 0 &&
      Number(result.verification.legacyTriggerCount || 0) === 0
    );
    result.finishedAt = new Date().toISOString();
    AUTOMATION_saveEmergencyRemediationResult_(result);
  } catch (err) {
    result.ok = false;
    result.finishedAt = new Date().toISOString();
    result.error = String(err && err.stack || err);
    AUTOMATION_saveEmergencyRemediationResult_(result);
    throw err;
  } finally {
    AUTOMATION_releaseModuleLease_(emergencyLease);
  }

  // 복구 lease 해제 뒤 장애 상태를 즉시 다시 계산합니다.
  try {
    result.healthCheck = AUTOMATION_runHealthMonitorSafe_({
      source: 'EMERGENCY_REPAIR',
      force: true,
      deadlineMs: Date.now() + 60 * 1000
    });
    AUTOMATION_saveEmergencyRemediationResult_(result);
  } catch (ignoreHealthError) {}

  if (ui) {
    ui.alert(
      result.ok ? '긴급 장애 복구 완료' : '긴급 장애 복구 확인 필요',
      [
        '정식 트리거: ' + Number(result.verification && result.verification.installedTriggerCount || 0) + '개',
        '계획 일치: ' + Number(result.verification && result.verification.canonicalMatchedTriggerCount || 0) + '/13개',
        '수신주소 재분류: ' + Number(result.mailFailureRepair && result.mailFailureRepair.changed || 0) + '건',
        '저장큐 재처리 전환: ' + Number(result.archiveQueueRepair && result.archiveQueueRepair.requeued || 0) + '건',
        '즉시 저장큐 처리: ' + Number(result.firstArchiveQueueRun && result.firstArchiveQueueRun.processed || 0) + '건',
        '',
        '나머지 RETRY 작업은 정식 5분 트리거가 계속 처리합니다.'
      ].join('\n'),
      ui.ButtonSet.OK
    );
  }

  return result;
}


function AUTOMATION_getLastEmergencyRemediationResult() {
  var raw = PropertiesService.getScriptProperties().getProperty(
    AUTOMATION_EMERGENCY_REPAIR_CONFIG.lastResultPropertyKey
  );
  if (!raw) return null;
  try { return JSON.parse(raw); } catch (err) { return null; }
}


function AUTOMATION_buildEmergencyRemediationPreview_() {
  var triggerSnapshot = TRG_buildStatusSnapshot_();
  var triggerSummary = triggerSnapshot.summary || {};
  var mailPreview = MAILOPS_reclassifyLegacyInvalidRecipientFailures_({
    maxRows: 5000,
    dryRun: true
  });
  var archivePreview = MAILOPS_requeueLegacyDriveQueryArchiveFailures_({
    force: true,
    maxRows: 20000,
    dryRun: true
  });

  return {
    generatedAt: new Date().toISOString(),
    version: AUTOMATION_EMERGENCY_REPAIR_CONFIG.version,
    triggers: {
      installed: Number(triggerSummary.installedTriggerCount || 0),
      matched: Number(triggerSummary.canonicalMatchedTriggerCount || 0),
      missing: Number(triggerSummary.canonicalMissingTriggerCount || 0),
      excess: Number(triggerSummary.canonicalExcessTriggerCount || 0),
      orphan: Number(triggerSummary.orphanTriggerCount || 0),
      legacy: Number(triggerSummary.legacyTriggerCount || 0),
      unknown: Number(triggerSummary.unknownTriggerCount || 0),
      healthy: TRG_isCanonicalSnapshotHealthy_(triggerSnapshot)
    },
    mailFailureQueue: mailPreview,
    archiveQueue: archivePreview
  };
}


function AUTOMATION_saveEmergencyRemediationResult_(result) {
  result = result || {};
  var payload = {
    version: AUTOMATION_EMERGENCY_REPAIR_CONFIG.version,
    ok: result.ok === true,
    startedAt: String(result.startedAt || ''),
    finishedAt: String(result.finishedAt || ''),
    error: String(result.error || '').slice(0, 2000),
    triggerRepair: result.triggerRepair ? {
      skipped: result.triggerRepair.skipped === true,
      reason: String(result.triggerRepair.reason || ''),
      deletedCount: Number(result.triggerRepair.deletedCount || 0),
      createdCount: Number(result.triggerRepair.createdCount || 0)
    } : null,
    verification: result.verification ? {
      installedTriggerCount: Number(result.verification.installedTriggerCount || 0),
      canonicalMatchedTriggerCount: Number(result.verification.canonicalMatchedTriggerCount || 0),
      canonicalMissingTriggerCount: Number(result.verification.canonicalMissingTriggerCount || 0),
      canonicalExcessTriggerCount: Number(result.verification.canonicalExcessTriggerCount || 0),
      orphanTriggerCount: Number(result.verification.orphanTriggerCount || 0),
      legacyTriggerCount: Number(result.verification.legacyTriggerCount || 0),
      unknownTriggerCount: Number(result.verification.unknownTriggerCount || 0)
    } : null,
    mailFailureRepair: result.mailFailureRepair ? {
      scanned: Number(result.mailFailureRepair.scanned || 0),
      candidates: Number(result.mailFailureRepair.candidates || 0),
      changed: Number(result.mailFailureRepair.changed || 0)
    } : null,
    archiveQueueRepair: result.archiveQueueRepair ? {
      scanned: Number(result.archiveQueueRepair.scanned || 0),
      candidates: Number(result.archiveQueueRepair.candidates || 0),
      requeued: Number(result.archiveQueueRepair.requeued || 0)
    } : null,
    firstArchiveQueueRun: result.firstArchiveQueueRun ? {
      ok: result.firstArchiveQueueRun.ok !== false,
      processed: Number(result.firstArchiveQueueRun.processed || 0),
      message: String(result.firstArchiveQueueRun.message || result.firstArchiveQueueRun.error || '').slice(0, 1000)
    } : null,
    healthCheck: result.healthCheck ? {
      status: String(result.healthCheck.status || ''),
      activeIssueCount: Number(result.healthCheck.activeIssueCount || 0),
      sentCount: Number(result.healthCheck.sentCount || 0),
      recoverySentCount: Number(result.healthCheck.recoverySentCount || 0)
    } : null
  };

  PropertiesService.getScriptProperties().setProperty(
    AUTOMATION_EMERGENCY_REPAIR_CONFIG.lastResultPropertyKey,
    JSON.stringify(payload)
  );
}
