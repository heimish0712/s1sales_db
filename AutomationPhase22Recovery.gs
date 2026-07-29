/****************************************************
 * AutomationPhase22Recovery.gs
 * 근본 복구:
 * - 제거된 외부 수행사 onEdit에서 남은 재처리 큐 종료
 * - 권한/Drive 호환/lease busy 실패 중 재시도 가능한 마스터 이벤트 복구
 * - CONTRACT_SYNC 만료 lease 회수
 * - 발송파일 저장큐 Drive 호환 FAIL·정체 RUNNING 복구
 ****************************************************/

var AUTOMATION_PHASE22_RECOVERY_CONFIG = Object.freeze({
  version: '2026-07-28-PHASE22',
  recoverableRetryErrorPattern: /LEASE_BUSY|Invalid query|Invalid field selection|File not found:\s*\[object Object\]|You do not have permission to access the requested document/i,
  contractLeaseKey: 'CONTRACT_SYNC'
});

function AUTOMATION_previewPhase22RootRecovery() {
  TRG_assertAutomationOwner_();
  var report = AUTOMATION_phase22BuildRecoveryReport_(true);
  try {
    SpreadsheetApp.getUi().alert(
      'Phase22 근본 복구 미리보기',
      [
        '외부/폐기 이벤트 후보: ' + Number(report.retryQueue.obsoleteCandidates || 0) + '건',
        '마스터 재시도 후보: ' + Number(report.retryQueue.masterRecoverableCandidates || 0) + '건',
        'CONTRACT_SYNC lease: ' + (report.contractLease.exists ? (report.contractLease.stale ? '만료/정체' : '정상 실행중') : '없음'),
        '발송큐 복구 후보: ' + Number(report.mailArchive.candidates || 0) + '건',
        '발송큐 정체 RUNNING: ' + Number(report.mailArchive.staleRunningCandidates || 0) + '건',
        '',
        '실행 전 변경은 없습니다.'
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {}
  return report;
}

function AUTOMATION_executePhase22RootRecovery() {
  TRG_assertAutomationOwner_();
  var report = AUTOMATION_phase22BuildRecoveryReport_(false);
  var core = AUTOMATION_runCoreDataSyncPipeline();
  report.coreStatus = String(core && core.status || '');
  report.coreStage1 = AUTOMATION_phase22FindStage_(core, 'MASTER_TO_COMPLETED');
  report.finishedAt = new Date().toISOString();

  try {
    SpreadsheetApp.getUi().alert(
      'Phase22 근본 복구',
      [
        '외부/폐기 이벤트 종료: ' + Number(report.retryQueue.obsoleteCompleted || 0) + '건',
        '마스터 재시도 복구: ' + Number(report.retryQueue.masterRequeued || 0) + '건',
        '만료 CONTRACT_SYNC lease 회수: ' + (report.contractLease.deleted ? '예' : '아니오'),
        '발송큐 재처리 전환: ' + Number(report.mailArchive.requeued || 0) + '건',
        '발송큐 정체 RUNNING 회수: ' + Number(report.mailArchive.staleRunningCandidates || 0) + '건',
        '핵심 파이프라인: ' + report.coreStatus,
        '1단계: ' + String(report.coreStage1 && report.coreStage1.status || ''),
        '1단계 오류: ' + String(report.coreStage1 && report.coreStage1.error || '없음')
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {}

  return report;
}

function AUTOMATION_phase22BuildRecoveryReport_(dryRun) {
  var result = {
    status: dryRun ? 'PREVIEW' : 'EXECUTED',
    retryQueue: AUTOMATION_phase22RepairRetryQueue_(dryRun),
    contractLease: AUTOMATION_phase22RecoverStaleContractLease_(dryRun),
    mailArchive: MAILOPS_requeueLegacyDriveQueryArchiveFailures_({
      force: true,
      dryRun: dryRun === true,
      markerKey: 'MAIL_ARCHIVE_DRIVE_COMPAT_REPAIR_V3',
      includeCompatFailures: true,
      includeStaleRunning: true,
      maxRows: 30000
    }),
    checkedAt: new Date().toISOString(),
    version: AUTOMATION_PHASE22_RECOVERY_CONFIG.version
  };

  console.log('[AUTOMATION_phase22BuildRecoveryReport_] ' + JSON.stringify(result));
  return result;
}

function AUTOMATION_phase22RepairRetryQueue_(dryRun) {
  var lease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_phase22RepairRetryQueue_',
      ttlMs: 2 * 60 * 1000,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );
  if (!lease.acquired) throw new Error('재처리 큐 쓰기 작업이 진행 중입니다.');

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var lastRow = sheet.getLastRow();
    var result = {
      scanned: 0,
      obsoleteCandidates: 0,
      obsoleteCompleted: 0,
      masterRecoverableCandidates: 0,
      masterRequeued: 0,
      sourceSummary: {}
    };
    if (lastRow < 2) return result;

    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var masterId = String(AUTOMATION_getRuntimeMasterSpreadsheetId_() || '');
    var now = new Date();
    result.scanned = values.length;

    values.forEach(function(row, offset) {
      var status = String(row[index['상태'] - 1] || '').trim().toUpperCase();
      var sourceId = String(row[index['소스파일ID'] - 1] || '').trim();
      var handler = String(row[index['핸들러'] - 1] || '').trim();
      var errorText = String(row[index['최근오류'] - 1] || '');
      var rowNo = offset + 2;
      var summaryKey = sourceId || '(빈ID)';
      result.sourceSummary[summaryKey] = Number(result.sourceSummary[summaryKey] || 0) + 1;

      var activeOrFailed = status === 'PENDING' || status === 'RETRY' || status === 'RUNNING' || status === 'FAIL';
      if (!activeOrFailed) return;

      if (!sourceId || sourceId !== masterId) {
        result.obsoleteCandidates++;
        if (dryRun) return;
        sheet.getRange(rowNo, index['상태']).setValue('DONE');
        sheet.getRange(rowNo, index['다음시도일시']).clearContent();
        sheet.getRange(rowNo, index['완료일시']).setValue(now);
        sheet.getRange(rowNo, index['최근오류']).setValue(
          '[PHASE22 자동종료] 외부 파일 역동기화 트리거가 폐기되어 더 이상 재처리하지 않습니다. ' +
          'source=' + sourceId + ' / handler=' + handler + ' / 기존오류=' +
          errorText.replace(/\s+/g, ' ').slice(0, 1000)
        );
        result.obsoleteCompleted++;
        return;
      }

      if (
        sourceId === masterId &&
        status === 'FAIL' &&
        AUTOMATION_PHASE22_RECOVERY_CONFIG.recoverableRetryErrorPattern.test(errorText)
      ) {
        result.masterRecoverableCandidates++;
        if (dryRun) return;
        sheet.getRange(rowNo, index['상태']).setValue('RETRY');
        sheet.getRange(rowNo, index['시도횟수']).setValue(0);
        sheet.getRange(rowNo, index['다음시도일시']).setValue(now);
        sheet.getRange(rowNo, index['최근시도일시']).clearContent();
        sheet.getRange(rowNo, index['완료일시']).clearContent();
        sheet.getRange(rowNo, index['최근오류']).setValue(
          '[PHASE22 재처리복구] 원인 패치 후 마스터 이벤트 재시도. 기존오류=' +
          errorText.replace(/\s+/g, ' ').slice(0, 1200)
        );
        result.masterRequeued++;
      }
    });

    return result;
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}

function AUTOMATION_phase22RecoverStaleContractLease_(dryRun) {
  var props = PropertiesService.getScriptProperties();
  var propertyKey = AUTOMATION_RUNTIME_CONFIG.leasePropertyPrefix + AUTOMATION_PHASE22_RECOVERY_CONFIG.contractLeaseKey;
  var lease = AUTOMATION_readLeaseProperty_(props, propertyKey);
  var result = { exists: !!lease, active: false, stale: false, deleted: false, lease: lease || null };
  if (!lease) return result;

  var nowMs = Date.now();
  result.active = AUTOMATION_isLeaseActive_(lease, nowMs);
  var heartbeatMs = Number(lease.heartbeatAtMs || lease.startedAtMs || 0);
  var ageMs = heartbeatMs > 0 ? nowMs - heartbeatMs : Number.MAX_SAFE_INTEGER;
  result.stale = !result.active || ageMs >= 10 * 60 * 1000;

  if (result.stale && !dryRun) {
    props.deleteProperty(propertyKey);
    result.deleted = true;
  }
  return result;
}

function AUTOMATION_phase22FindStage_(summary, stageKey) {
  var stages = summary && Array.isArray(summary.stages) ? summary.stages : [];
  for (var i = 0; i < stages.length; i++) {
    if (String(stages[i] && stages[i].key || '') === String(stageKey || '')) return stages[i];
  }
  return null;
}
