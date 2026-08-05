/****************************************************
 * AutomationPhase26Recovery.gs
 * 2026정보통신유지보수 헤더 기반 매핑 전환 및 과거 오류 종료
 ****************************************************/

var AUTOMATION_PHASE26_CONFIG = Object.freeze({
  version: '2026-08-02-PHASE26-HEADER-DYNAMIC',
  headerErrorMarker: '2026정보통신유지보수 헤더 구조가 예상과 달라',
  headerDynamicErrorMarker: '2026정보통신유지보수 헤더 기반 매핑을 구성할 수 없어',
  headerCurrentErrorMarker: '서무 정보통신 대상 시트의 헤더 기반 매핑을 구성할 수 없어',
  retryHandler: 'ITMNEW_syncFromEdit_2026'
});


function AUTOMATION_previewPhase26HeaderRecovery() {
  var headerMapping = ITMAINT_previewTargetHeaderMapping_2026();
  var retryRows = AUTOMATION_phase26FindHeaderFailureRows_();
  var preview = ITMNEW_buildSyncPlan_2026_({ route: 'PHASE26_PREVIEW' });
  var result = {
    status: 'PREVIEW',
    headerMapping: headerMapping,
    headerFailureQueueCount: retryRows.length,
    headerFailureJobs: retryRows.slice(0, 50).map(function(item) {
      return { row: item.rowNo, jobId: item.jobId, status: item.status, error: item.error };
    }),
    syncPlan: preview,
    version: AUTOMATION_PHASE26_CONFIG.version
  };
  Logger.log('[AUTOMATION_previewPhase26HeaderRecovery] ' + JSON.stringify(result));
  return result;
}


function AUTOMATION_executePhase26HeaderRecovery() {
  var headerMapping = ITMAINT_previewTargetHeaderMapping_2026();
  var syncResult = ITMNEW_syncMissingContractsNow_2026();
  var closedQueue = AUTOMATION_phase26CloseHeaderFailureRows_();
  var pipelineResult = AUTOMATION_runCoreDataSyncPipeline();
  var healthResult = null;

  if (
    pipelineResult &&
    (pipelineResult.status === 'COMPLETED' ||
      pipelineResult.status === 'COMPLETED_WITH_PENDING_RETRIES')
  ) {
    try {
      healthResult = AUTOMATION_runHealthMonitorNow();
    } catch (healthError) {
      healthResult = { status: 'ERROR', error: AUTOMATION_errorMessage_(healthError) };
    }
  }

  var result = {
    status: 'SUCCESS',
    headerMapping: headerMapping,
    syncResult: syncResult,
    closedHeaderFailureQueueRows: closedQueue,
    pipelineStatus: String(pipelineResult && pipelineResult.status || ''),
    healthStatus: String(healthResult && healthResult.status || ''),
    version: AUTOMATION_PHASE26_CONFIG.version
  };
  Logger.log('[AUTOMATION_executePhase26HeaderRecovery] ' + JSON.stringify(result));
  return result;
}


function AUTOMATION_phase26FindHeaderFailureRows_() {
  var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
  var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
  var index = AUTOMATION_makeHeaderIndex_(headers);
  var lastRow = sheet.getLastRow();
  var result = [];
  if (lastRow < 2) return result;

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getDisplayValues();
  values.forEach(function(row, offset) {
    var status = String(row[index['상태'] - 1] || '').toUpperCase();
    var moduleName = String(row[index['모듈'] - 1] || '').toUpperCase();
    var handler = String(row[index['핸들러'] - 1] || '');
    var error = String(row[index['최근오류'] - 1] || '');
    var isHeaderError = error.indexOf(AUTOMATION_PHASE26_CONFIG.headerErrorMarker) >= 0 ||
      error.indexOf(AUTOMATION_PHASE26_CONFIG.headerDynamicErrorMarker) >= 0 ||
      error.indexOf(AUTOMATION_PHASE26_CONFIG.headerCurrentErrorMarker) >= 0;
    var isItMaintenance = moduleName === 'IT_MAINTENANCE_SYNC' ||
      handler === AUTOMATION_PHASE26_CONFIG.retryHandler;

    if (!isHeaderError || !isItMaintenance) return;
    if (status !== 'FAIL' && status !== 'RETRY' && status !== 'PENDING') return;

    result.push({
      rowNo: offset + 2,
      jobId: String(row[index['작업ID'] - 1] || ''),
      status: status,
      error: error
    });
  });
  return result;
}


function AUTOMATION_phase26CloseHeaderFailureRows_() {
  var rows = AUTOMATION_phase26FindHeaderFailureRows_();
  if (!rows.length) return { matched: 0, closed: 0 };

  var lease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_phase26CloseHeaderFailureRows_',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseTtlMs,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );
  if (!lease.acquired) {
    throw new Error('PHASE26 재처리 큐 정리 lease를 얻지 못했습니다: ' + String(lease.reason || 'LEASE_BUSY'));
  }

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var index = AUTOMATION_makeHeaderIndex_(AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders);
    var now = new Date();
    rows.forEach(function(item) {
      sheet.getRange(item.rowNo, index['상태']).setValue('DONE');
      sheet.getRange(item.rowNo, index['다음시도일시']).setValue('');
      sheet.getRange(item.rowNo, index['최근오류']).setValue(
        '[PHASE26 자동종료] 헤더 기반 전체 누락계약 동기화로 대체 처리했습니다.'
      );
      sheet.getRange(item.rowNo, index['완료일시']).setValue(now);
    });
    return { matched: rows.length, closed: rows.length };
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}
