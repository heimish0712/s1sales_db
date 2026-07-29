/****************************************************
 * AutomationPhase21Recovery.gs
 * - Drive API v2 parent reference 객체 오류 복구
 * - 미확정 지역값으로 인한 마스터→수주확정 실패 진단
 ****************************************************/

var AUTOMATION_PHASE21_RECOVERY_CONFIG = Object.freeze({
  version: '2026-07-28-PHASE21',
  parentObjectErrorPattern: /File not found:\s*\[object Object\]|removeParents[^\n]*\[object Object\]/i
});

function AUTOMATION_previewPhase21Recovery() {
  TRG_assertAutomationOwner_();
  var queue = AUTOMATION_phase21ScanRetryQueue_(true);
  var core = AUTOMATION_getCoreDataSyncLastRun() || {};
  var stage1 = AUTOMATION_phase21FindStage_(core, 'MASTER_TO_COMPLETED');
  var result = {
    status: 'PREVIEW',
    retryQueueParentObjectFailures: queue.candidates,
    coreStatus: String(core.status || ''),
    stage1Status: String(stage1 && stage1.status || ''),
    stage1Error: String(stage1 && stage1.error || ''),
    checkedAt: new Date().toISOString(),
    version: AUTOMATION_PHASE21_RECOVERY_CONFIG.version
  };

  try {
    SpreadsheetApp.getUi().alert(
      'Phase21 복구 미리보기',
      [
        '재처리 큐 [object Object] FAIL: ' + result.retryQueueParentObjectFailures + '건',
        '최근 핵심 상태: ' + (result.coreStatus || '기록없음'),
        '1단계 상태: ' + (result.stage1Status || '기록없음'),
        '1단계 오류: ' + (result.stage1Error || '기록없음')
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {}

  return result;
}

function AUTOMATION_executePhase21Recovery() {
  TRG_assertAutomationOwner_();
  var queue = AUTOMATION_phase21ScanRetryQueue_(false);
  var core = AUTOMATION_runCoreDataSyncPipeline();
  var stage1 = AUTOMATION_phase21FindStage_(core, 'MASTER_TO_COMPLETED');
  var result = {
    status: String(core.status || ''),
    retryQueueRequeued: queue.requeued,
    stage1Status: String(stage1 && stage1.status || ''),
    stage1Error: String(stage1 && stage1.error || ''),
    finishedAt: new Date().toISOString(),
    version: AUTOMATION_PHASE21_RECOVERY_CONFIG.version
  };

  try {
    SpreadsheetApp.getUi().alert(
      'Phase21 복구 실행',
      [
        '[object Object] FAIL → RETRY: ' + result.retryQueueRequeued + '건',
        '핵심 파이프라인: ' + (result.status || '확인불가'),
        '1단계: ' + (result.stage1Status || '확인불가'),
        '1단계 오류: ' + (result.stage1Error || '없음')
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {}

  return result;
}

function AUTOMATION_phase21ScanRetryQueue_(dryRun) {
  var lease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_phase21ScanRetryQueue_',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseTtlMs,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );

  if (!lease.acquired) {
    throw new Error('재처리 큐 쓰기 작업이 진행 중이라 Phase21 복구를 시작할 수 없습니다.');
  }

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var lastRow = sheet.getLastRow();
    var result = { scanned: 0, candidates: 0, requeued: 0, dryRun: dryRun === true };

    if (lastRow < 2) return result;

    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var now = new Date();
    result.scanned = values.length;

    values.forEach(function(row, offset) {
      var status = String(row[index['상태'] - 1] || '').trim().toUpperCase();
      var errorText = String(row[index['최근오류'] - 1] || '');
      if (status !== 'FAIL') return;
      if (!AUTOMATION_PHASE21_RECOVERY_CONFIG.parentObjectErrorPattern.test(errorText)) return;

      result.candidates++;
      if (dryRun === true) return;

      var rowNo = offset + 2;
      sheet.getRange(rowNo, index['상태']).setValue('RETRY');
      sheet.getRange(rowNo, index['시도횟수']).setValue(0);
      sheet.getRange(rowNo, index['다음시도일시']).setValue(now);
      sheet.getRange(rowNo, index['최근시도일시']).clearContent();
      sheet.getRange(rowNo, index['완료일시']).clearContent();
      sheet.getRange(rowNo, index['최근오류']).setValue(
        '[PHASE21 자동복구] Drive API v2 parents 객체를 부모 ID 문자열로 정규화한 뒤 재처리 예약. 기존 오류: ' +
        errorText.replace(/\s+/g, ' ').slice(0, 1200)
      );
      result.requeued++;
    });

    return result;
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}

function AUTOMATION_phase21FindStage_(summary, stageKey) {
  var stages = summary && Array.isArray(summary.stages) ? summary.stages : [];
  for (var i = 0; i < stages.length; i++) {
    if (String(stages[i] && stages[i].key || '') === String(stageKey || '')) {
      return stages[i];
    }
  }
  return null;
}
