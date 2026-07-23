/****************************************************
 * AutomationMaintenance.gs
 * 재처리 큐·기술 로그·다운로드 토큰 정리 - 12단계
 *
 * 원칙:
 * - PENDING/RETRY/RUNNING 재처리 작업은 자동 삭제하지 않는다.
 * - DONE 30일, FAIL 90일 경과 후 숨김 이력 시트로 이동한다.
 * - 다운로드 토큰은 만료·손상된 항목만 삭제한다.
 * - 업무 증빙 성격의 메일발송로그/발송파일로그는 자동 정리 대상에서 제외한다.
 * - KJ 업로드 로그의 COPIED/EXISTS_BY_NAME 최신 1건은 중복 방지 레지스트리로 영구 보존한다.
 * - 새 트리거를 추가하지 않고 기존 5분 핵심 파이프라인에서 하루 1회 제한 실행한다.
 ****************************************************/

var AUTOMATION_MAINTENANCE_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE14',

  moduleLeaseKey: 'MAINTENANCE_CLEANUP',
  moduleLeaseTtlMs: 5 * 60 * 1000,
  moduleLeaseWaitMs: 0,

  statePropertyKey: 'AUTOMATION_MAINTENANCE_STATE_V1',
  lastCompletedPropertyKey: 'AUTOMATION_MAINTENANCE_LAST_COMPLETED_V1',
  lastResultPropertyKey: 'AUTOMATION_MAINTENANCE_LAST_RESULT_V1',
  dailyIntervalMs: 20 * 60 * 60 * 1000,
  scheduledMaxRuntimeMs: 20 * 1000,
  manualMaxRuntimeMs: 4.5 * 60 * 1000,
  minimumStartBudgetMs: 8 * 1000,

  statusSheetName: '_자동화유지관리',
  retryArchiveSheetName: '_자동화재처리이력',
  retryArchiveHeadersExtra: Object.freeze(['보관일시', '보관사유']),

  retryDoneRetentionDays: 30,
  retryFailRetentionDays: 90,
  retryArchiveRetentionDays: 365,
  retryArchiveMaxRows: 20000,
  retryScheduledBatch: 300,
  retryManualBatch: 3000,

  downloadTokenPrefix: 'MAILAUTO_MULTI_DOWNLOAD_',

  genericScheduledDeleteBatch: 800,
  genericManualDeleteBatch: 8000,

  cutoverLogRetentionDays: 365,
  cutoverLogMaxRows: 500,

  kjClassifierLogRetentionDays: 90,
  kjClassifierLogMaxRows: 20000,

  kjClassifierStateRetentionDays: 365,
  kjClassifierStateMaxRows: 50000,

  kjUploadDiagnosticRetentionDays: 180,
  kjUploadDiagnosticMaxRows: 10000,

  maxSummaryErrorLength: 1500
});


/****************************************************
 * 공개 실행 함수
 ****************************************************/

/**
 * 삭제·이동 없이 현재 정리 대상만 계산한다.
 */
function AUTOMATION_previewMaintenance() {
  TRG_assertAutomationOwner_();

  var report = AUTOMATION_buildMaintenancePreview_();
  AUTOMATION_showMaintenancePreviewAlert_(report);
  return report;
}


/**
 * 보존정책에 따라 유지관리 사이클을 수동 실행한다.
 */
function AUTOMATION_runMaintenanceNow() {
  TRG_assertAutomationOwner_();

  var preview = AUTOMATION_buildMaintenancePreview_();
  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '자동화 유지관리 실행',
    [
      '다음 기술 데이터만 정리합니다.',
      '',
      '재처리 큐 보관 예정: ' + preview.retryQueue.archiveEligible + '건',
      '재처리 이력 삭제 예정: ' + preview.retryHistory.deleteEligible + '건',
      '만료·손상 다운로드 토큰: ' + preview.downloadTokens.deleteEligible + '건',
      '기술 로그 삭제 예정 합계: ' + preview.technicalLogs.deleteEligible + '건',
      '백업 휴지통 이동 예정: ' + Number(preview.backupRetention.deleteEligible || 0) + '건',
      '',
      'PENDING/RETRY/RUNNING 작업과 메일발송로그·발송파일로그는 삭제하지 않습니다.',
      '계속하시겠습니까?'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return {
      status: 'CANCELLED',
      preview: preview
    };
  }

  var result = AUTOMATION_runMaintenanceCycle_({
    mode: 'MANUAL',
    maxRuntimeMs: AUTOMATION_MAINTENANCE_CONFIG.manualMaxRuntimeMs
  });

  AUTOMATION_showMaintenanceResultAlert_(result);
  return result;
}


/**
 * 기존 5분 핵심 파이프라인에서 호출한다.
 * 마지막 완료 후 20시간이 지나지 않았고 진행 중 사이클도 없으면 실행하지 않는다.
 */
function AUTOMATION_runScheduledMaintenanceIfDue_(options) {
  options = options || {};

  var props = PropertiesService.getScriptProperties();
  var rawState = String(props.getProperty(AUTOMATION_MAINTENANCE_CONFIG.statePropertyKey) || '');
  var lastCompletedMs = AUTOMATION_maintenanceTimeMs_(
    props.getProperty(AUTOMATION_MAINTENANCE_CONFIG.lastCompletedPropertyKey)
  );
  var nowMs = Date.now();
  var due = !!rawState || !lastCompletedMs ||
    nowMs - lastCompletedMs >= AUTOMATION_MAINTENANCE_CONFIG.dailyIntervalMs;

  if (!due) {
    return {
      status: 'SKIPPED_NOT_DUE',
      lastCompletedAt: lastCompletedMs ? new Date(lastCompletedMs).toISOString() : ''
    };
  }

  var hardDeadlineMs = Number(options.hardDeadlineMs || 0);
  var availableMs = hardDeadlineMs > 0
    ? hardDeadlineMs - nowMs
    : AUTOMATION_MAINTENANCE_CONFIG.scheduledMaxRuntimeMs;

  if (availableMs < AUTOMATION_MAINTENANCE_CONFIG.minimumStartBudgetMs) {
    return {
      status: 'SKIPPED_RUNTIME_BUDGET',
      availableMs: Math.max(0, availableMs)
    };
  }

  return AUTOMATION_runMaintenanceCycle_({
    mode: 'SCHEDULED',
    maxRuntimeMs: Math.min(
      AUTOMATION_MAINTENANCE_CONFIG.scheduledMaxRuntimeMs,
      availableMs - 1000
    )
  });
}


/**
 * 숨김 재처리 이력 시트를 연다.
 */
function AUTOMATION_showRetryArchiveSheet() {
  TRG_assertAutomationOwner_();
  var sheet = AUTOMATION_getOrCreateRetryArchiveSheet_();
  sheet.showSheet();
  sheet.getParent().setActiveSheet(sheet);
  return sheet.getName();
}


/**
 * 현재 큐에 남아 있는 최종 FAIL 전체를 새 재시도 작업으로 되돌린다.
 * 이미 이력 시트로 보관된 과거 FAIL은 대상이 아니다.
 */
function AUTOMATION_requeueFailedRetryJobs() {
  TRG_assertAutomationOwner_();

  var ui = SpreadsheetApp.getUi();
  var response = ui.alert(
    '최종 실패 작업 다시 시도',
    '현재 _자동화재처리큐 시트에 남아 있는 FAIL 작업의 시도횟수를 0으로 초기화하고 RETRY 상태로 되돌립니다. 계속하시겠습니까?',
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    return { status: 'CANCELLED', requeued: 0 };
  }

  var writeLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_requeueFailedRetryJobs',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseTtlMs,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );

  if (!writeLease.acquired) {
    throw new Error('재처리 큐 쓰기 작업이 진행 중이라 FAIL 작업을 초기화하지 못했습니다.');
  }

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var lastRow = sheet.getLastRow();
    var requeued = 0;

    if (lastRow >= 2) {
      var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
      var now = new Date();

      for (var i = 0; i < values.length; i++) {
        if (String(values[i][index['상태'] - 1] || '').toUpperCase() !== 'FAIL') continue;

        var rowNo = i + 2;
        sheet.getRange(rowNo, index['상태']).setValue('RETRY');
        sheet.getRange(rowNo, index['다음시도일시']).setValue(now);
        sheet.getRange(rowNo, index['시도횟수']).setValue(0);
        sheet.getRange(rowNo, index['최근시도일시']).setValue('');
        sheet.getRange(rowNo, index['완료일시']).setValue('');
        requeued++;
      }
    }

    ui.alert('최종 실패 작업 다시 시도', 'RETRY로 되돌린 작업: ' + requeued + '건', ui.ButtonSet.OK);
    return { status: 'SUCCESS', requeued: requeued };
  } finally {
    AUTOMATION_releaseModuleLease_(writeLease);
  }
}


/**
 * 만료·손상 다운로드 토큰만 즉시 정리한다.
 */
function AUTOMATION_cleanupDownloadTokensNow() {
  TRG_assertAutomationOwner_();
  var result = AUTOMATION_cleanupExpiredDownloadTokens_();

  SpreadsheetApp.getUi().alert(
    '다운로드 토큰 정리',
    [
      '검사: ' + result.scanned + '개',
      '유효 유지: ' + result.valid + '개',
      '만료 삭제: ' + result.expiredDeleted + '개',
      '손상 삭제: ' + result.invalidDeleted + '개'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}


function AUTOMATION_showMaintenanceStatusSheet() {
  TRG_assertAutomationOwner_();
  var sheet = AUTOMATION_getOrCreateMaintenanceStatusSheet_();
  sheet.showSheet();
  sheet.getParent().setActiveSheet(sheet);
  return sheet.getName();
}


/****************************************************
 * 유지관리 사이클
 ****************************************************/

function AUTOMATION_runMaintenanceCycle_(options) {
  options = options || {};

  var mode = String(options.mode || 'SCHEDULED').toUpperCase();
  var maxRuntimeMs = Math.max(
    5000,
    Number(options.maxRuntimeMs) || AUTOMATION_MAINTENANCE_CONFIG.scheduledMaxRuntimeMs
  );
  var startedAtMs = Date.now();
  var deadlineMs = startedAtMs + maxRuntimeMs;
  var manual = mode === 'MANUAL';
  var summary = {
    version: AUTOMATION_MAINTENANCE_CONFIG.version,
    mode: mode,
    status: 'STARTED',
    startedAt: new Date(startedAtMs).toISOString(),
    finishedAt: '',
    durationMs: 0,
    cycleId: '',
    startStepIndex: 0,
    endStepIndex: 0,
    completedCycle: false,
    steps: [],
    errors: []
  };

  var lease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_MAINTENANCE_CONFIG.moduleLeaseKey,
    {
      taskName: 'AUTOMATION_runMaintenanceCycle_',
      ttlMs: AUTOMATION_MAINTENANCE_CONFIG.moduleLeaseTtlMs,
      waitMs: AUTOMATION_MAINTENANCE_CONFIG.moduleLeaseWaitMs
    }
  );

  if (!lease.acquired) {
    summary.status = 'SKIPPED_ALREADY_RUNNING';
    summary.finishedAt = new Date().toISOString();
    summary.durationMs = Date.now() - startedAtMs;
    return summary;
  }

  try {
    var state = AUTOMATION_readMaintenanceState_();
    if (!state) {
      state = {
        cycleId: AUTOMATION_createRuntimeToken_('MAINT'),
        startedAt: new Date().toISOString(),
        stepIndex: 0,
        stepAttempts: 0,
        lastError: '',
        version: AUTOMATION_MAINTENANCE_CONFIG.version
      };
      AUTOMATION_writeMaintenanceState_(state);
    }

    summary.cycleId = state.cycleId;
    summary.startStepIndex = Number(state.stepIndex || 0);

    var steps = AUTOMATION_getMaintenanceSteps_();

    while (state.stepIndex < steps.length) {
      if (Date.now() >= deadlineMs - 1000) break;

      var step = steps[state.stepIndex];
      var stepStartedAtMs = Date.now();
      var stepResult = null;

      try {
        stepResult = step.run({
          manual: manual,
          deadlineMs: deadlineMs
        }) || {};
        stepResult.status = stepResult.status || 'SUCCESS';
        state.stepAttempts = 0;
        state.lastError = '';
      } catch (err) {
        var errorText = AUTOMATION_maintenanceErrorMessage_(err);
        stepResult = {
          status: 'ERROR',
          done: false,
          error: errorText
        };
        state.stepAttempts = Number(state.stepAttempts || 0) + 1;
        state.lastError = errorText;
        summary.errors.push(step.label + ': ' + errorText);
      }

      stepResult.key = step.key;
      stepResult.label = step.label;
      stepResult.durationMs = Date.now() - stepStartedAtMs;
      summary.steps.push(AUTOMATION_maintenanceJsonSafe_(stepResult));

      if (stepResult.done === true) {
        state.stepIndex++;
        state.stepAttempts = 0;
        state.lastError = '';
        AUTOMATION_writeMaintenanceState_(state);
        continue;
      }

      AUTOMATION_writeMaintenanceState_(state);

      if (
        manual &&
        stepResult.status === 'SUCCESS' &&
        Date.now() < deadlineMs - AUTOMATION_MAINTENANCE_CONFIG.minimumStartBudgetMs
      ) {
        continue;
      }

      break;
    }

    summary.endStepIndex = Number(state.stepIndex || 0);

    if (state.stepIndex >= steps.length) {
      summary.status = summary.errors.length ? 'COMPLETED_WITH_ERRORS' : 'COMPLETED';
      summary.completedCycle = true;
      AUTOMATION_completeMaintenanceCycle_(state, summary);
    } else if (summary.errors.length) {
      summary.status = 'PARTIAL_ERROR';
    } else {
      summary.status = 'PARTIAL_CONTINUE';
    }
  } catch (err) {
    summary.status = 'ERROR';
    summary.errors.push(AUTOMATION_maintenanceErrorMessage_(err));
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }

  summary.finishedAt = new Date().toISOString();
  summary.durationMs = Date.now() - startedAtMs;
  AUTOMATION_storeMaintenanceResult_(summary);
  AUTOMATION_writeMaintenanceStatus_(summary);
  console.log('[AUTOMATION_runMaintenanceCycle_] ' + JSON.stringify(summary));
  return summary;
}


function AUTOMATION_getMaintenanceSteps_() {
  return [
    {
      key: 'DOWNLOAD_TOKENS',
      label: '만료 다운로드 토큰',
      run: function() {
        var result = AUTOMATION_cleanupExpiredDownloadTokens_();
        result.done = true;
        return result;
      }
    },
    {
      key: 'RETRY_QUEUE_ARCHIVE',
      label: '재처리 큐 완료·실패 보관',
      run: function(ctx) {
        return AUTOMATION_archiveRetryQueueBatch_(ctx.manual);
      }
    },
    {
      key: 'RETRY_HISTORY_CLEANUP',
      label: '재처리 이력 보존기간 정리',
      run: function(ctx) {
        return AUTOMATION_cleanupRetryHistoryBatch_(ctx.manual);
      }
    },
    {
      key: 'CUTOVER_LOG',
      label: '자동화 전환 기록 정리',
      run: function(ctx) {
        return AUTOMATION_cleanupGenericTechnicalLog_({
          spreadsheet: AUTOMATION_getRuntimeMasterSpreadsheet_(),
          sheetName: '_자동화전환기록',
          dateColumn: 4,
          retentionDays: AUTOMATION_MAINTENANCE_CONFIG.cutoverLogRetentionDays,
          maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.cutoverLogMaxRows,
          maxDeletes: AUTOMATION_maintenanceDeleteBatch_(ctx.manual)
        });
      }
    },
    {
      key: 'KJ_CLASSIFIER_LOG',
      label: 'KJ 서류분류 로그 정리',
      run: function(ctx) {
        return AUTOMATION_runMaintenanceStepWithModuleLease_(
          'KJ_CLASSIFIER',
          'AUTOMATION_cleanupKjClassifierLog_',
          function() {
            return AUTOMATION_cleanupGenericTechnicalLog_({
              spreadsheet: SpreadsheetApp.openById(String(KJ_DOC_CONFIG.SPREADSHEET_ID || '').trim()),
              sheetName: KJ_DOC_CONFIG.LOG_SHEET_NAME,
              dateColumn: 1,
              retentionDays: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierLogRetentionDays,
              maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierLogMaxRows,
              maxDeletes: AUTOMATION_maintenanceDeleteBatch_(ctx.manual)
            });
          }
        );
      }
    },
    {
      key: 'KJ_CLASSIFIER_STATE',
      label: 'KJ 서류분류 상태이력 정리',
      run: function(ctx) {
        return AUTOMATION_runMaintenanceStepWithModuleLease_(
          'KJ_CLASSIFIER',
          'AUTOMATION_cleanupKjClassifierState_',
          function() {
            return AUTOMATION_cleanupGenericTechnicalLog_({
              spreadsheet: SpreadsheetApp.openById(String(KJ_DOC_CONFIG.SPREADSHEET_ID || '').trim()),
              sheetName: KJ_DOC_CONFIG.STATE_SHEET_NAME,
              dateColumn: 1,
              retentionDays: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierStateRetentionDays,
              maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierStateMaxRows,
              maxDeletes: AUTOMATION_maintenanceDeleteBatch_(ctx.manual)
            });
          }
        );
      }
    },
    {
      key: 'KJ_UPLOAD_LOG',
      label: 'KJ 수행사 업로드 로그 압축',
      run: function(ctx) {
        return AUTOMATION_runMaintenanceStepWithModuleLease_(
          'KJ_VENDOR_UPLOAD',
          'AUTOMATION_cleanupKjUploadLog_',
          function() {
            return AUTOMATION_cleanupKjUploadLog_(ctx.manual);
          }
        );
      }
    },
    {
      key: 'BACKUP_RETENTION',
      label: '영업관리대장 백업 보존정책',
      run: function(ctx) {
        return AUTOMATION_cleanupBackupRetention_({
          mode: ctx.manual ? 'MANUAL' : 'SCHEDULED',
          maxTrash: ctx.manual
            ? BACKUP_RETENTION_CONFIG.manualMaxTrash
            : BACKUP_RETENTION_CONFIG.scheduledMaxTrash
        });
      }
    }
  ];
}


function AUTOMATION_completeMaintenanceCycle_(state, summary) {
  var props = PropertiesService.getScriptProperties();
  var completedAt = new Date().toISOString();

  props.setProperty(AUTOMATION_MAINTENANCE_CONFIG.lastCompletedPropertyKey, completedAt);
  props.deleteProperty(AUTOMATION_MAINTENANCE_CONFIG.statePropertyKey);
  summary.lastCompletedAt = completedAt;
}


function AUTOMATION_readMaintenanceState_() {
  var raw = String(
    PropertiesService.getScriptProperties().getProperty(
      AUTOMATION_MAINTENANCE_CONFIG.statePropertyKey
    ) || ''
  );

  if (!raw) return null;

  try {
    var parsed = JSON.parse(raw);
    if (!parsed || !parsed.cycleId || !isFinite(Number(parsed.stepIndex))) {
      throw new Error('상태 형식 오류');
    }
    return parsed;
  } catch (err) {
    PropertiesService.getScriptProperties().deleteProperty(
      AUTOMATION_MAINTENANCE_CONFIG.statePropertyKey
    );
    return null;
  }
}


function AUTOMATION_writeMaintenanceState_(state) {
  PropertiesService.getScriptProperties().setProperty(
    AUTOMATION_MAINTENANCE_CONFIG.statePropertyKey,
    JSON.stringify({
      cycleId: String(state.cycleId || ''),
      startedAt: String(state.startedAt || ''),
      stepIndex: Number(state.stepIndex || 0),
      stepAttempts: Number(state.stepAttempts || 0),
      lastError: AUTOMATION_maintenanceTruncate_(state.lastError || '', 800),
      version: AUTOMATION_MAINTENANCE_CONFIG.version
    })
  );
}


function AUTOMATION_runMaintenanceStepWithModuleLease_(moduleKey, taskName, callback) {
  var lease = AUTOMATION_acquireModuleLease_(moduleKey, {
    taskName: taskName,
    waitMs: 0
  });

  if (!lease.acquired) {
    return {
      status: 'LEASE_BUSY',
      done: false,
      module: moduleKey,
      reason: lease.reason || 'LEASE_BUSY'
    };
  }

  try {
    return callback();
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


/****************************************************
 * 재처리 큐 보관·정리
 ****************************************************/

function AUTOMATION_archiveRetryQueueBatch_(manual) {
  var writeLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_archiveRetryQueueBatch_',
      ttlMs: 2 * 60 * 1000,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );

  if (!writeLease.acquired) {
    return {
      status: 'LEASE_BUSY',
      done: false,
      archived: 0,
      eligible: 0
    };
  }

  try {
    var queueSheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var archiveSheet = AUTOMATION_getOrCreateRetryArchiveSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var lastRow = queueSheet.getLastRow();
    var batchLimit = manual
      ? AUTOMATION_MAINTENANCE_CONFIG.retryManualBatch
      : AUTOMATION_MAINTENANCE_CONFIG.retryScheduledBatch;

    if (lastRow < 2) {
      return { status: 'SUCCESS', done: true, archived: 0, eligible: 0 };
    }

    var values = queueSheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    var nowMs = Date.now();
    var doneCutoffMs = nowMs - AUTOMATION_MAINTENANCE_CONFIG.retryDoneRetentionDays * 86400000;
    var failCutoffMs = nowMs - AUTOMATION_MAINTENANCE_CONFIG.retryFailRetentionDays * 86400000;
    var eligible = [];

    values.forEach(function(row, i) {
      var status = String(row[index['상태'] - 1] || '').toUpperCase();
      if (status !== 'DONE' && status !== 'FAIL') return;

      var completedMs = AUTOMATION_maintenanceTimeMs_(row[index['완료일시'] - 1]) ||
        AUTOMATION_maintenanceTimeMs_(row[index['최근시도일시'] - 1]) ||
        AUTOMATION_maintenanceTimeMs_(row[index['최근요청일시'] - 1]);

      if (!completedMs) return;

      var cutoffMs = status === 'DONE' ? doneCutoffMs : failCutoffMs;
      if (completedMs > cutoffMs) return;

      eligible.push({
        rowNo: i + 2,
        jobId: String(row[index['작업ID'] - 1] || ''),
        status: status,
        row: row,
        reason: status + '_' + (status === 'DONE'
          ? AUTOMATION_MAINTENANCE_CONFIG.retryDoneRetentionDays
          : AUTOMATION_MAINTENANCE_CONFIG.retryFailRetentionDays) + 'D'
      });
    });

    if (!eligible.length) {
      return { status: 'SUCCESS', done: true, archived: 0, eligible: 0 };
    }

    var selected = eligible.slice(0, batchLimit);
    var existingIds = AUTOMATION_loadFirstColumnSet_(archiveSheet);
    var archiveRows = [];
    var archivedAt = new Date();

    selected.forEach(function(item) {
      if (item.jobId && existingIds[item.jobId]) return;
      archiveRows.push(item.row.concat([archivedAt, item.reason]));
      if (item.jobId) existingIds[item.jobId] = true;
    });

    if (archiveRows.length) {
      archiveSheet.getRange(
        archiveSheet.getLastRow() + 1,
        1,
        archiveRows.length,
        archiveRows[0].length
      ).setValues(archiveRows);
    }

    AUTOMATION_deleteSheetRowsDescending_(queueSheet, selected.map(function(item) {
      return item.rowNo;
    }));

    return {
      status: 'SUCCESS',
      done: eligible.length <= selected.length,
      eligible: eligible.length,
      archived: selected.length,
      appended: archiveRows.length,
      remainingEstimate: Math.max(0, eligible.length - selected.length)
    };
  } finally {
    AUTOMATION_releaseModuleLease_(writeLease);
  }
}


function AUTOMATION_cleanupRetryHistoryBatch_(manual) {
  var sheet = AUTOMATION_getOrCreateRetryArchiveSheet_();
  var dateColumn = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders.length + 1;

  return AUTOMATION_cleanupGenericTechnicalLog_({
    spreadsheet: sheet.getParent(),
    sheetName: sheet.getName(),
    dateColumn: dateColumn,
    retentionDays: AUTOMATION_MAINTENANCE_CONFIG.retryArchiveRetentionDays,
    maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.retryArchiveMaxRows,
    maxDeletes: manual
      ? AUTOMATION_MAINTENANCE_CONFIG.retryManualBatch
      : AUTOMATION_MAINTENANCE_CONFIG.retryScheduledBatch
  });
}


function AUTOMATION_getOrCreateRetryArchiveSheet_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var name = AUTOMATION_MAINTENANCE_CONFIG.retryArchiveSheetName;
  var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders.concat(
    AUTOMATION_MAINTENANCE_CONFIG.retryArchiveHeadersExtra
  );
  var sheet = ss.getSheetByName(name);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }

  var current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  var mismatch = headers.some(function(header, i) {
    return String(current[i] || '') !== header;
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


/****************************************************
 * 다운로드 토큰
 ****************************************************/

function AUTOMATION_cleanupExpiredDownloadTokens_() {
  var props = PropertiesService.getScriptProperties();
  var all = props.getProperties();
  var nowMs = Date.now();
  var result = {
    status: 'SUCCESS',
    scanned: 0,
    valid: 0,
    expiredDeleted: 0,
    invalidDeleted: 0,
    deletedKeys: []
  };

  Object.keys(all).forEach(function(key) {
    if (key.indexOf(AUTOMATION_MAINTENANCE_CONFIG.downloadTokenPrefix) !== 0) return;
    result.scanned++;

    var invalid = false;
    var expired = false;

    try {
      var payload = JSON.parse(String(all[key] || ''));
      var expiresAt = Number(payload && payload.expiresAt);
      var payloadToken = String(payload && payload.token || '');
      var keyToken = key.slice(AUTOMATION_MAINTENANCE_CONFIG.downloadTokenPrefix.length);

      invalid = !payload || !payloadToken || payloadToken !== keyToken || !isFinite(expiresAt);
      expired = !invalid && expiresAt <= nowMs;
    } catch (err) {
      invalid = true;
    }

    if (invalid || expired) {
      props.deleteProperty(key);
      result.deletedKeys.push(key);
      if (invalid) result.invalidDeleted++;
      else result.expiredDeleted++;
    } else {
      result.valid++;
    }
  });

  result.deleted = result.expiredDeleted + result.invalidDeleted;
  result.deletedKeys = result.deletedKeys.slice(0, 100);
  return result;
}


/****************************************************
 * 기술 로그 정리
 ****************************************************/

function AUTOMATION_cleanupGenericTechnicalLog_(options) {
  options = options || {};

  var ss = options.spreadsheet;
  var sheetName = String(options.sheetName || '');
  var sheet = ss && sheetName ? ss.getSheetByName(sheetName) : null;

  if (!sheet) {
    return {
      status: 'SHEET_NOT_FOUND',
      done: true,
      sheetName: sheetName,
      deleted: 0,
      eligible: 0
    };
  }

  var lastRow = sheet.getLastRow();
  var dataRows = Math.max(0, lastRow - 1);
  if (!dataRows) {
    return {
      status: 'SUCCESS',
      done: true,
      sheetName: sheetName,
      deleted: 0,
      eligible: 0,
      remainingRows: 0
    };
  }

  var dateColumn = Math.max(1, Number(options.dateColumn) || 1);
  var retentionDays = Math.max(1, Number(options.retentionDays) || 365);
  var maxDataRows = Math.max(1, Number(options.maxDataRows) || 10000);
  var maxDeletes = Math.max(1, Number(options.maxDeletes) || 500);
  var cutoffMs = Date.now() - retentionDays * 86400000;
  var dateValues = sheet.getRange(2, dateColumn, dataRows, 1).getValues();
  var candidateMap = {};
  var candidates = [];

  dateValues.forEach(function(row, i) {
    var timeMs = AUTOMATION_maintenanceTimeMs_(row[0]);
    if (timeMs && timeMs <= cutoffMs) {
      var rowNo = i + 2;
      candidateMap[rowNo] = true;
      candidates.push(rowNo);
    }
  });

  var rowsNeededForMax = Math.max(0, dataRows - maxDataRows - candidates.length);
  for (var rowNo = 2; rowNo <= lastRow && rowsNeededForMax > 0; rowNo++) {
    if (candidateMap[rowNo]) continue;
    candidateMap[rowNo] = true;
    candidates.push(rowNo);
    rowsNeededForMax--;
  }

  candidates.sort(function(a, b) { return a - b; });
  var selected = candidates.slice(0, maxDeletes);
  AUTOMATION_deleteSheetRowsDescending_(sheet, selected);

  return {
    status: 'SUCCESS',
    done: candidates.length <= selected.length,
    sheetName: sheetName,
    eligible: candidates.length,
    deleted: selected.length,
    remainingEstimate: Math.max(0, candidates.length - selected.length),
    remainingRows: Math.max(0, dataRows - selected.length)
  };
}


function AUTOMATION_cleanupKjUploadLog_(manual) {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss.getSheetByName(KJUS_CFG.LOG_SHEET_NAME);

  if (!sheet) {
    return {
      status: 'SHEET_NOT_FOUND',
      done: true,
      sheetName: KJUS_CFG.LOG_SHEET_NAME,
      deleted: 0,
      eligible: 0
    };
  }

  var lastRow = sheet.getLastRow();
  var dataRows = Math.max(0, lastRow - 1);
  if (!dataRows) {
    return {
      status: 'SUCCESS',
      done: true,
      sheetName: sheet.getName(),
      deleted: 0,
      eligible: 0
    };
  }

  var values = sheet.getRange(2, 1, dataRows, Math.max(7, KJUS_CFG.LOG_HEADER.length)).getValues();
  var seenSuccessKeys = {};
  var candidateMap = {};
  var diagnosticRows = [];
  var duplicateSuccessCount = 0;
  var diagnosticCutoffMs = Date.now() -
    AUTOMATION_MAINTENANCE_CONFIG.kjUploadDiagnosticRetentionDays * 86400000;

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var rowNo = i + 2;
    var timeMs = AUTOMATION_maintenanceTimeMs_(row[0]);
    var status = String(row[2] || '').toUpperCase();
    var sourceFileId = String(row[3] || '');
    var destFolderId = String(row[6] || '');
    var success = status === 'COPIED' || status === 'EXISTS_BY_NAME';

    if (success && sourceFileId && destFolderId) {
      var successKey = sourceFileId + '|' + destFolderId;
      if (seenSuccessKeys[successKey]) {
        candidateMap[rowNo] = true;
        duplicateSuccessCount++;
      } else {
        seenSuccessKeys[successKey] = true;
      }
      continue;
    }

    diagnosticRows.push(rowNo);
    if (timeMs && timeMs <= diagnosticCutoffMs) {
      candidateMap[rowNo] = true;
    }
  }

  diagnosticRows.sort(function(a, b) { return a - b; });
  var diagnosticAlreadyDeleted = diagnosticRows.filter(function(rowNo) {
    return !!candidateMap[rowNo];
  }).length;
  var diagnosticExtra = Math.max(
    0,
    diagnosticRows.length - diagnosticAlreadyDeleted -
      AUTOMATION_MAINTENANCE_CONFIG.kjUploadDiagnosticMaxRows
  );

  for (var d = 0; d < diagnosticRows.length && diagnosticExtra > 0; d++) {
    var diagnosticRowNo = diagnosticRows[d];
    if (candidateMap[diagnosticRowNo]) continue;
    candidateMap[diagnosticRowNo] = true;
    diagnosticExtra--;
  }

  var candidates = Object.keys(candidateMap).map(Number).sort(function(a, b) { return a - b; });
  var maxDeletes = AUTOMATION_maintenanceDeleteBatch_(manual);
  var selected = candidates.slice(0, maxDeletes);
  AUTOMATION_deleteSheetRowsDescending_(sheet, selected);

  return {
    status: 'SUCCESS',
    done: candidates.length <= selected.length,
    sheetName: sheet.getName(),
    eligible: candidates.length,
    deleted: selected.length,
    duplicateSuccessRows: duplicateSuccessCount,
    preservedSuccessKeys: Object.keys(seenSuccessKeys).length,
    remainingEstimate: Math.max(0, candidates.length - selected.length)
  };
}


/****************************************************
 * 미리보기·상태 기록
 ****************************************************/

function AUTOMATION_buildMaintenancePreview_() {
  var report = {
    generatedAt: new Date().toISOString(),
    retryQueue: AUTOMATION_previewRetryQueueMaintenance_(),
    retryHistory: AUTOMATION_previewGenericLog_({
      spreadsheet: AUTOMATION_getRuntimeMasterSpreadsheet_(),
      sheetName: AUTOMATION_MAINTENANCE_CONFIG.retryArchiveSheetName,
      dateColumn: AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders.length + 1,
      retentionDays: AUTOMATION_MAINTENANCE_CONFIG.retryArchiveRetentionDays,
      maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.retryArchiveMaxRows
    }),
    downloadTokens: AUTOMATION_previewDownloadTokens_(),
    backupRetention: AUTOMATION_getBackupRetentionPreview_(),
    technicalLogs: {
      items: []
    }
  };

  report.technicalLogs.items.push(AUTOMATION_previewGenericLog_({
    spreadsheet: AUTOMATION_getRuntimeMasterSpreadsheet_(),
    sheetName: '_자동화전환기록',
    dateColumn: 4,
    retentionDays: AUTOMATION_MAINTENANCE_CONFIG.cutoverLogRetentionDays,
    maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.cutoverLogMaxRows
  }));

  var kjSs = SpreadsheetApp.openById(String(KJ_DOC_CONFIG.SPREADSHEET_ID || '').trim());
  report.technicalLogs.items.push(AUTOMATION_previewGenericLog_({
    spreadsheet: kjSs,
    sheetName: KJ_DOC_CONFIG.LOG_SHEET_NAME,
    dateColumn: 1,
    retentionDays: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierLogRetentionDays,
    maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierLogMaxRows
  }));
  report.technicalLogs.items.push(AUTOMATION_previewGenericLog_({
    spreadsheet: kjSs,
    sheetName: KJ_DOC_CONFIG.STATE_SHEET_NAME,
    dateColumn: 1,
    retentionDays: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierStateRetentionDays,
    maxDataRows: AUTOMATION_MAINTENANCE_CONFIG.kjClassifierStateMaxRows
  }));
  report.technicalLogs.items.push(AUTOMATION_previewKjUploadLog_());

  report.technicalLogs.deleteEligible = report.technicalLogs.items.reduce(function(sum, item) {
    return sum + Number(item.deleteEligible || 0);
  }, 0);

  return report;
}


function AUTOMATION_previewRetryQueueMaintenance_() {
  var sheet = AUTOMATION_getRuntimeMasterSpreadsheet_().getSheetByName(
    AUTOMATION_RUNTIME_CONFIG.retryQueueSheetName
  );
  var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
  var index = AUTOMATION_makeHeaderIndex_(headers);
  var lastRow = sheet ? sheet.getLastRow() : 0;
  var result = {
    totalRows: Math.max(0, lastRow - 1),
    pending: 0,
    retry: 0,
    running: 0,
    done: 0,
    fail: 0,
    archiveEligible: 0
  };

  if (lastRow < 2) return result;

  var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
  var nowMs = Date.now();
  var doneCutoffMs = nowMs - AUTOMATION_MAINTENANCE_CONFIG.retryDoneRetentionDays * 86400000;
  var failCutoffMs = nowMs - AUTOMATION_MAINTENANCE_CONFIG.retryFailRetentionDays * 86400000;

  values.forEach(function(row) {
    var status = String(row[index['상태'] - 1] || '').toUpperCase();
    if (status === 'PENDING') result.pending++;
    if (status === 'RETRY') result.retry++;
    if (status === 'RUNNING') result.running++;
    if (status === 'DONE') result.done++;
    if (status === 'FAIL') result.fail++;

    if (status !== 'DONE' && status !== 'FAIL') return;
    var timeMs = AUTOMATION_maintenanceTimeMs_(row[index['완료일시'] - 1]) ||
      AUTOMATION_maintenanceTimeMs_(row[index['최근시도일시'] - 1]) ||
      AUTOMATION_maintenanceTimeMs_(row[index['최근요청일시'] - 1]);
    if (!timeMs) return;

    if (status === 'DONE' && timeMs <= doneCutoffMs) result.archiveEligible++;
    if (status === 'FAIL' && timeMs <= failCutoffMs) result.archiveEligible++;
  });

  return result;
}


function AUTOMATION_previewDownloadTokens_() {
  var all = PropertiesService.getScriptProperties().getProperties();
  var nowMs = Date.now();
  var result = {
    scanned: 0,
    valid: 0,
    expired: 0,
    invalid: 0,
    deleteEligible: 0
  };

  Object.keys(all).forEach(function(key) {
    if (key.indexOf(AUTOMATION_MAINTENANCE_CONFIG.downloadTokenPrefix) !== 0) return;
    result.scanned++;

    try {
      var payload = JSON.parse(String(all[key] || ''));
      var expiresAt = Number(payload && payload.expiresAt);
      var expectedToken = key.slice(AUTOMATION_MAINTENANCE_CONFIG.downloadTokenPrefix.length);
      var actualToken = String(payload && payload.token || '');

      if (!payload || !isFinite(expiresAt) || !actualToken || actualToken !== expectedToken) {
        result.invalid++;
      } else if (expiresAt <= nowMs) {
        result.expired++;
      } else {
        result.valid++;
      }
    } catch (err) {
      result.invalid++;
    }
  });

  result.deleteEligible = result.expired + result.invalid;
  return result;
}


function AUTOMATION_previewGenericLog_(options) {
  var ss = options.spreadsheet;
  var sheetName = String(options.sheetName || '');
  var sheet = ss && sheetName ? ss.getSheetByName(sheetName) : null;

  if (!sheet) {
    return {
      sheetName: sheetName,
      totalRows: 0,
      deleteEligible: 0,
      status: 'SHEET_NOT_FOUND'
    };
  }

  var dataRows = Math.max(0, sheet.getLastRow() - 1);
  if (!dataRows) {
    return {
      sheetName: sheetName,
      totalRows: 0,
      deleteEligible: 0,
      status: 'SUCCESS'
    };
  }

  var cutoffMs = Date.now() - Math.max(1, Number(options.retentionDays) || 365) * 86400000;
  var dates = sheet.getRange(2, Math.max(1, Number(options.dateColumn) || 1), dataRows, 1).getValues();
  var candidate = {};

  dates.forEach(function(row, i) {
    var timeMs = AUTOMATION_maintenanceTimeMs_(row[0]);
    if (timeMs && timeMs <= cutoffMs) candidate[i + 2] = true;
  });

  var maxDataRows = Math.max(1, Number(options.maxDataRows) || 10000);
  var extra = Math.max(0, dataRows - maxDataRows - Object.keys(candidate).length);
  for (var rowNo = 2; rowNo <= sheet.getLastRow() && extra > 0; rowNo++) {
    if (candidate[rowNo]) continue;
    candidate[rowNo] = true;
    extra--;
  }

  return {
    sheetName: sheetName,
    totalRows: dataRows,
    deleteEligible: Object.keys(candidate).length,
    status: 'SUCCESS'
  };
}


function AUTOMATION_previewKjUploadLog_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss.getSheetByName(KJUS_CFG.LOG_SHEET_NAME);

  if (!sheet) {
    return {
      sheetName: KJUS_CFG.LOG_SHEET_NAME,
      totalRows: 0,
      deleteEligible: 0,
      status: 'SHEET_NOT_FOUND'
    };
  }

  var dataRows = Math.max(0, sheet.getLastRow() - 1);
  if (!dataRows) {
    return {
      sheetName: sheet.getName(),
      totalRows: 0,
      deleteEligible: 0,
      status: 'SUCCESS'
    };
  }

  var values = sheet.getRange(2, 1, dataRows, Math.max(7, KJUS_CFG.LOG_HEADER.length)).getValues();
  var seen = {};
  var candidates = {};
  var diagnostics = [];
  var cutoffMs = Date.now() -
    AUTOMATION_MAINTENANCE_CONFIG.kjUploadDiagnosticRetentionDays * 86400000;

  for (var i = values.length - 1; i >= 0; i--) {
    var row = values[i];
    var rowNo = i + 2;
    var status = String(row[2] || '').toUpperCase();
    var sourceId = String(row[3] || '');
    var destId = String(row[6] || '');
    var success = status === 'COPIED' || status === 'EXISTS_BY_NAME';

    if (success && sourceId && destId) {
      var key = sourceId + '|' + destId;
      if (seen[key]) candidates[rowNo] = true;
      else seen[key] = true;
      continue;
    }

    diagnostics.push(rowNo);
    var timeMs = AUTOMATION_maintenanceTimeMs_(row[0]);
    if (timeMs && timeMs <= cutoffMs) candidates[rowNo] = true;
  }

  diagnostics.sort(function(a, b) { return a - b; });
  var diagnosticsAlreadyDeleted = diagnostics.filter(function(rowNo) {
    return !!candidates[rowNo];
  }).length;
  var extra = Math.max(
    0,
    diagnostics.length - diagnosticsAlreadyDeleted -
      AUTOMATION_MAINTENANCE_CONFIG.kjUploadDiagnosticMaxRows
  );
  for (var d = 0; d < diagnostics.length && extra > 0; d++) {
    if (candidates[diagnostics[d]]) continue;
    candidates[diagnostics[d]] = true;
    extra--;
  }

  return {
    sheetName: sheet.getName(),
    totalRows: dataRows,
    deleteEligible: Object.keys(candidates).length,
    preservedSuccessKeys: Object.keys(seen).length,
    status: 'SUCCESS'
  };
}


function AUTOMATION_showMaintenancePreviewAlert_(report) {
  SpreadsheetApp.getUi().alert(
    '자동화 유지관리 미리보기',
    [
      '재처리 큐 전체: ' + report.retryQueue.totalRows + '건',
      '- 활성(PENDING/RETRY/RUNNING): ' +
        (report.retryQueue.pending + report.retryQueue.retry + report.retryQueue.running) + '건',
      '- 보관 대상(DONE 30일 / FAIL 90일): ' + report.retryQueue.archiveEligible + '건',
      '재처리 이력 삭제 대상: ' + report.retryHistory.deleteEligible + '건',
      '만료·손상 다운로드 토큰: ' + report.downloadTokens.deleteEligible + '건',
      '기술 로그 삭제 대상: ' + report.technicalLogs.deleteEligible + '건',
      '백업 휴지통 이동 대상: ' + Number(report.backupRetention.deleteEligible || 0) + '건',
      '',
      '메일발송로그·발송파일로그는 자동 정리 대상이 아닙니다.'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function AUTOMATION_showMaintenanceResultAlert_(result) {
  var stepLines = (result.steps || []).map(function(step) {
    return '- ' + step.label + ': ' + String(step.status || '') +
      ' / 삭제·보관 ' + Number(step.deleted || step.archived || 0) + '건';
  });

  SpreadsheetApp.getUi().alert(
    '자동화 유지관리 결과',
    [
      '상태: ' + result.status,
      '사이클 완료: ' + (result.completedCycle ? '예' : '아니오'),
      '소요: ' + (Number(result.durationMs || 0) / 1000).toFixed(1) + '초',
      '',
      stepLines.join('\n'),
      result.errors && result.errors.length ? ('\n오류:\n' + result.errors.join('\n')) : ''
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}


function AUTOMATION_storeMaintenanceResult_(summary) {
  var compact = {
    version: summary.version,
    mode: summary.mode,
    status: summary.status,
    cycleId: summary.cycleId,
    startedAt: summary.startedAt,
    finishedAt: summary.finishedAt,
    durationMs: summary.durationMs,
    completedCycle: summary.completedCycle,
    startStepIndex: summary.startStepIndex,
    endStepIndex: summary.endStepIndex,
    stepCount: (summary.steps || []).length,
    errors: (summary.errors || []).slice(0, 5).map(function(item) {
      return AUTOMATION_maintenanceTruncate_(item, 400);
    })
  };

  PropertiesService.getScriptProperties().setProperty(
    AUTOMATION_MAINTENANCE_CONFIG.lastResultPropertyKey,
    JSON.stringify(compact)
  );
}


function AUTOMATION_writeMaintenanceStatus_(summary) {
  try {
    var sheet = AUTOMATION_getOrCreateMaintenanceStatusSheet_();
    var headers = [
      '작업키', '최근시작', '최근종료', '상태', '모드', '사이클ID',
      '시작단계', '종료단계', '사이클완료', '소요초', '단계요약', '오류', '버전'
    ];
    var row = [
      'TECHNICAL_DATA_MAINTENANCE',
      summary.startedAt,
      summary.finishedAt,
      summary.status,
      summary.mode,
      summary.cycleId,
      summary.startStepIndex,
      summary.endStepIndex,
      summary.completedCycle ? '예' : '아니오',
      Math.round(Number(summary.durationMs || 0) / 100) / 10,
      AUTOMATION_maintenanceTruncate_((summary.steps || []).map(function(step) {
        return step.key + ':' + step.status + ':' + Number(step.deleted || step.archived || 0);
      }).join(' | '), 3000),
      AUTOMATION_maintenanceTruncate_((summary.errors || []).join(' | '), 2000),
      AUTOMATION_MAINTENANCE_CONFIG.version
    ];

    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.setFrozenRows(1);
    sheet.getRange(2, 1, 1, row.length).setValues([row]);
  } catch (err) {
    console.error('[AUTOMATION_writeMaintenanceStatus_] ' + AUTOMATION_maintenanceErrorMessage_(err), err);
  }
}


function AUTOMATION_getOrCreateMaintenanceStatusSheet_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss.getSheetByName(AUTOMATION_MAINTENANCE_CONFIG.statusSheetName);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(AUTOMATION_MAINTENANCE_CONFIG.statusSheetName);
    created = true;
  }

  if (created) {
    try { sheet.hideSheet(); } catch (ignoreHideError) {}
  }

  return sheet;
}


/****************************************************
 * 공통 유틸리티
 ****************************************************/

function AUTOMATION_maintenanceDeleteBatch_(manual) {
  return manual
    ? AUTOMATION_MAINTENANCE_CONFIG.genericManualDeleteBatch
    : AUTOMATION_MAINTENANCE_CONFIG.genericScheduledDeleteBatch;
}


function AUTOMATION_deleteSheetRowsDescending_(sheet, rowNumbers) {
  if (!sheet || !rowNumbers || !rowNumbers.length) return 0;

  var unique = {};
  rowNumbers.forEach(function(rowNo) {
    var n = Number(rowNo);
    if (isFinite(n) && n >= 2) unique[n] = true;
  });

  var sorted = Object.keys(unique).map(Number).sort(function(a, b) { return b - a; });
  if (!sorted.length) return 0;

  var groups = [];
  var groupHigh = sorted[0];
  var groupLow = sorted[0];

  for (var i = 1; i < sorted.length; i++) {
    var current = sorted[i];
    if (current === groupLow - 1) {
      groupLow = current;
    } else {
      groups.push({ start: groupLow, count: groupHigh - groupLow + 1 });
      groupHigh = current;
      groupLow = current;
    }
  }
  groups.push({ start: groupLow, count: groupHigh - groupLow + 1 });

  groups.forEach(function(group) {
    sheet.deleteRows(group.start, group.count);
  });

  return sorted.length;
}


function AUTOMATION_loadFirstColumnSet_(sheet) {
  var set = {};
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return set;

  var values = sheet.getRange(2, 1, lastRow - 1, 1).getDisplayValues();
  values.forEach(function(row) {
    var value = String(row[0] || '');
    if (value) set[value] = true;
  });
  return set;
}


function AUTOMATION_maintenanceTimeMs_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') {
    var dateMs = value.getTime();
    return isFinite(dateMs) ? dateMs : 0;
  }

  var parsed = new Date(value).getTime();
  return isFinite(parsed) ? parsed : 0;
}


function AUTOMATION_maintenanceErrorMessage_(err) {
  if (!err) return '알 수 없는 오류';
  return err.message ? String(err.message) : String(err);
}


function AUTOMATION_maintenanceTruncate_(value, maxLength) {
  var text = String(value || '');
  var limit = Math.max(1, Number(maxLength) || 1);
  return text.length <= limit ? text : text.slice(0, Math.max(0, limit - 3)) + '...';
}


function AUTOMATION_maintenanceJsonSafe_(value) {
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (err) {
    return { status: 'SERIALIZE_ERROR', value: String(value) };
  }
}
