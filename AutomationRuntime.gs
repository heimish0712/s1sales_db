/****************************************************
 * AutomationRuntime.gs
 * 자동화 기능별 lease + 편집 재처리 큐 - 5단계
 *
 * 목적:
 * - 프로젝트 전체를 장시간 막는 ScriptLock 대신 기능별 soft lease 사용
 * - lease 획득 실패/일시 오류가 발생한 편집 이벤트를 숨김 큐에 저장
 * - 5분 핵심 동기화 파이프라인 시작 시 큐를 제한적으로 재처리
 *
 * 운영 전제:
 * - 정식 설치형 트리거는 bang@s1samsung.com 하나가 소유
 * - soft lease는 Script Properties에 모듈별 키로 저장
 * - 만료 lease는 다음 실행이 자동 회수
 ****************************************************/

var AUTOMATION_RUNTIME_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE11',

  leasePropertyPrefix: 'AUTOMATION_MODULE_LEASE_V1_',
  defaultLeaseTtlMs: 8 * 60 * 1000,
  defaultLeaseWaitMs: 0,
  leaseElectionMs: 120,
  leasePollMs: 150,
  leaseHeartbeatMs: 30 * 1000,

  retryQueueSheetName: '_자동화재처리큐',
  retryQueueProcessLeaseKey: 'RETRY_QUEUE_PROCESS',
  retryQueueWriteLeaseKey: 'RETRY_QUEUE_WRITE',
  retryQueueWriteLeaseTtlMs: 20 * 1000,
  retryQueueWriteWaitMs: 3000,
  retryQueueProcessLeaseTtlMs: 2 * 60 * 1000,
  retryQueueMaxJobsPerRun: 20,
  retryQueueMaxRuntimeMs: 45 * 1000,
  retryQueueMaxAttempts: 8,
  retryQueueRunningStaleMs: 10 * 60 * 1000,
  retryQueueMaxErrorLength: 2000,

  retryQueueHeaders: Object.freeze([
    '작업ID', '중복키', '모듈', '핸들러', '상태',
    '소스파일ID', '시트명', '시작행', '행수', '시작열', '열수',
    '최초요청일시', '최근요청일시', '다음시도일시', '시도횟수', '최대시도',
    '최근시도일시', '최근오류', '완료일시', '이벤트JSON', '버전'
  ])
});

var AUTOMATION_RUNTIME_MASTER_SPREADSHEET_ID_CACHE = '';
var AUTOMATION_RUNTIME_MASTER_SPREADSHEET_CACHE = null;


var AUTOMATION_MODULE_LEASE_DEFAULTS = Object.freeze({
  CONTRACT_SYNC: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 500 }),
  VENDOR_SYNC: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 500 }),
  IT_MAINTENANCE_SYNC: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 500 }),
  CUSTOMER_FOLDER: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 300 }),
  KJ_CLASSIFIER: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 1000 }),
  KJ_VENDOR_UPLOAD: Object.freeze({ ttlMs: 6 * 60 * 1000, waitMs: 500 }),
  BACKUP: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 500 }),
  MAIL_ARCHIVE_QUEUE: Object.freeze({ ttlMs: 5 * 60 * 1000, waitMs: 500 }),
  MAIL_ARCHIVE_WRITE: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 1000 }),
  MAIL_HISTORY: Object.freeze({ ttlMs: 8 * 60 * 1000, waitMs: 500 }),
  MAIL_REQUEST_APPEND: Object.freeze({ ttlMs: 30 * 1000, waitMs: 12000 }),
  DISCORD_SALES_SUPPORT: Object.freeze({ ttlMs: 2 * 60 * 1000, waitMs: 0 }),
  SIMPLE_EDIT_REPAIR: Object.freeze({ ttlMs: 3 * 60 * 1000, waitMs: 0 }),
  TRIGGER_CUTOVER: Object.freeze({ ttlMs: 12 * 60 * 1000, waitMs: 1000 })
});


/****************************************************
 * 기능별 soft lease
 ****************************************************/

function AUTOMATION_acquireModuleLease_(moduleKey, options) {
  options = options || {};

  var normalizedModuleKey = AUTOMATION_normalizeModuleKey_(moduleKey);

  if (
    !AUTOMATION_isCutoverLeaseExemptModule_(normalizedModuleKey) &&
    AUTOMATION_getActiveCutoverLease_()
  ) {
    return {
      acquired: false,
      moduleKey: normalizedModuleKey,
      reason: 'CUTOVER_IN_PROGRESS'
    };
  }

  var defaults = AUTOMATION_MODULE_LEASE_DEFAULTS[normalizedModuleKey] || {};
  var ttlMs = Math.max(
    1000,
    Number(options.ttlMs || defaults.ttlMs || AUTOMATION_RUNTIME_CONFIG.defaultLeaseTtlMs)
  );
  var waitMs = Math.max(
    0,
    Number(
      typeof options.waitMs !== 'undefined'
        ? options.waitMs
        : (typeof defaults.waitMs !== 'undefined'
          ? defaults.waitMs
          : AUTOMATION_RUNTIME_CONFIG.defaultLeaseWaitMs)
    ) || 0
  );
  var waitUntilMs = Date.now() + waitMs;
  var props = PropertiesService.getScriptProperties();
  var propertyKey = AUTOMATION_RUNTIME_CONFIG.leasePropertyPrefix + normalizedModuleKey;
  var taskName = String(options.taskName || normalizedModuleKey);

  while (true) {
    var nowMs = Date.now();
    var existing = AUTOMATION_readLeaseProperty_(props, propertyKey);

    if (existing && AUTOMATION_isLeaseActive_(existing, nowMs)) {
      if (nowMs < waitUntilMs) {
        Utilities.sleep(Math.min(
          AUTOMATION_RUNTIME_CONFIG.leasePollMs,
          Math.max(30, waitUntilMs - nowMs)
        ));
        continue;
      }

      return {
        acquired: false,
        moduleKey: normalizedModuleKey,
        propertyKey: propertyKey,
        reason: 'LEASE_BUSY',
        existingTaskName: String(existing.taskName || ''),
        existingStartedAt: String(existing.startedAt || ''),
        existingHeartbeatAt: String(existing.heartbeatAt || existing.startedAt || ''),
        existingExpiresAt: String(existing.expiresAt || '')
      };
    }

    if (existing) {
      props.deleteProperty(propertyKey);
    }

    var token = AUTOMATION_createRuntimeToken_('LEASE');
    var acquiredAtMs = Date.now();
    var expiresAtMs = acquiredAtMs + ttlMs;
    var leaseRecord = {
      token: token,
      moduleKey: normalizedModuleKey,
      taskName: taskName,
      startedAtMs: acquiredAtMs,
      startedAt: new Date(acquiredAtMs).toISOString(),
      heartbeatAtMs: acquiredAtMs,
      heartbeatAt: new Date(acquiredAtMs).toISOString(),
      ttlMs: ttlMs,
      expiresAtMs: expiresAtMs,
      expiresAt: new Date(expiresAtMs).toISOString(),
      version: AUTOMATION_RUNTIME_CONFIG.version
    };

    props.setProperty(propertyKey, JSON.stringify(leaseRecord));
    Utilities.sleep(AUTOMATION_RUNTIME_CONFIG.leaseElectionMs);

    var confirmed = AUTOMATION_readLeaseProperty_(props, propertyKey);

    if (confirmed && String(confirmed.token || '') === token) {
      if (
        !AUTOMATION_isCutoverLeaseExemptModule_(normalizedModuleKey) &&
        AUTOMATION_getActiveCutoverLease_()
      ) {
        var electedLease = {
          acquired: true,
          moduleKey: normalizedModuleKey,
          propertyKey: propertyKey,
          token: token
        };
        AUTOMATION_releaseModuleLease_(electedLease);

        return {
          acquired: false,
          moduleKey: normalizedModuleKey,
          propertyKey: propertyKey,
          reason: 'CUTOVER_IN_PROGRESS'
        };
      }

      return {
        acquired: true,
        moduleKey: normalizedModuleKey,
        propertyKey: propertyKey,
        token: token,
        taskName: taskName,
        ttlMs: ttlMs,
        startedAtMs: acquiredAtMs,
        startedAt: leaseRecord.startedAt,
        expiresAtMs: expiresAtMs,
        expiresAt: leaseRecord.expiresAt,
        lastHeartbeatMs: acquiredAtMs
      };
    }

    if (Date.now() >= waitUntilMs) {
      return {
        acquired: false,
        moduleKey: normalizedModuleKey,
        propertyKey: propertyKey,
        reason: 'LEASE_ELECTION_LOST'
      };
    }

    Utilities.sleep(AUTOMATION_RUNTIME_CONFIG.leasePollMs);
  }
}


function AUTOMATION_refreshModuleLease_(lease, force) {
  if (!lease || !lease.acquired || !lease.propertyKey || !lease.token) return false;

  var nowMs = Date.now();

  if (
    force !== true &&
    nowMs - Number(lease.lastHeartbeatMs || 0) < AUTOMATION_RUNTIME_CONFIG.leaseHeartbeatMs
  ) {
    return true;
  }

  var props = PropertiesService.getScriptProperties();
  var current = AUTOMATION_readLeaseProperty_(props, lease.propertyKey);

  if (!current || String(current.token || '') !== String(lease.token || '')) {
    return false;
  }

  var ttlMs = Math.max(1000, Number(lease.ttlMs) || AUTOMATION_RUNTIME_CONFIG.defaultLeaseTtlMs);
  current.heartbeatAtMs = nowMs;
  current.heartbeatAt = new Date(nowMs).toISOString();
  current.expiresAtMs = nowMs + ttlMs;
  current.expiresAt = new Date(current.expiresAtMs).toISOString();

  props.setProperty(lease.propertyKey, JSON.stringify(current));
  lease.lastHeartbeatMs = nowMs;
  lease.expiresAtMs = current.expiresAtMs;
  lease.expiresAt = current.expiresAt;

  return true;
}


function AUTOMATION_releaseModuleLease_(lease) {
  if (!lease || !lease.acquired || !lease.propertyKey || !lease.token) return false;

  var props = PropertiesService.getScriptProperties();
  var current = AUTOMATION_readLeaseProperty_(props, lease.propertyKey);

  if (!current) return true;
  if (String(current.token || '') !== String(lease.token || '')) return false;

  props.deleteProperty(lease.propertyKey);
  return true;
}


function AUTOMATION_runWithModuleLeaseOrThrow_(moduleKey, taskName, callback, options) {
  if (typeof callback !== 'function') {
    throw new Error('기능별 lease 실행 콜백이 함수가 아닙니다: ' + taskName);
  }

  options = options || {};
  options.taskName = taskName || moduleKey;

  var lease = AUTOMATION_acquireModuleLease_(moduleKey, options);

  if (!lease.acquired) {
    var err = new Error(
      '[' + String(taskName || moduleKey) + '] 다른 동일 기능 작업이 실행 중이라 시작하지 못했습니다.'
    );
    err.automationLeaseBusy = true;
    err.automationModuleKey = AUTOMATION_normalizeModuleKey_(moduleKey);
    err.automationLeaseReason = lease.reason || 'LEASE_BUSY';
    throw err;
  }

  try {
    return callback(lease);
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


function AUTOMATION_runEditHandlerWithLease_(moduleKey, handlerName, e, callback, options) {
  options = options || {};

  var lease = AUTOMATION_acquireModuleLease_(moduleKey, {
    taskName: handlerName,
    ttlMs: options.ttlMs,
    waitMs: options.waitMs
  });

  if (!lease.acquired) {
    return AUTOMATION_deferEditEvent_(
      moduleKey,
      handlerName,
      e,
      lease.reason || 'LEASE_BUSY',
      ''
    );
  }

  try {
    var callbackResult = callback(lease);

    return {
      status: 'SUCCESS',
      module: AUTOMATION_normalizeModuleKey_(moduleKey),
      handler: handlerName,
      result: AUTOMATION_makeRuntimeJsonSafe_(callbackResult),
      retryQueued: false
    };
  } catch (err) {
    var errorText = AUTOMATION_runtimeErrorMessage_(err);
    var deferred = AUTOMATION_deferEditEvent_(
      moduleKey,
      handlerName,
      e,
      'HANDLER_ERROR',
      errorText
    );

    deferred.status = 'ERROR';
    deferred.error = errorText;
    console.error('[' + handlerName + '] ' + errorText, err);
    return deferred;
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


function AUTOMATION_isCutoverLeaseExemptModule_(moduleKey) {
  var normalized = AUTOMATION_normalizeModuleKey_(moduleKey);
  return normalized === 'TRIGGER_CUTOVER' || normalized === 'RETRY_QUEUE_WRITE';
}


function AUTOMATION_getActiveCutoverLease_() {
  var propertyKey = AUTOMATION_RUNTIME_CONFIG.leasePropertyPrefix + 'TRIGGER_CUTOVER';
  var props = PropertiesService.getScriptProperties();
  var lease = AUTOMATION_readLeaseProperty_(props, propertyKey);

  if (!lease) return null;
  if (AUTOMATION_isLeaseActive_(lease, Date.now())) return lease;

  props.deleteProperty(propertyKey);
  return null;
}


function AUTOMATION_readLeaseProperty_(props, propertyKey) {
  var raw = props.getProperty(propertyKey);
  if (!raw) return null;

  try {
    var parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (err) {
    props.deleteProperty(propertyKey);
    return null;
  }
}


function AUTOMATION_isLeaseActive_(lease, nowMs) {
  if (!lease || !lease.token) return false;

  var expiresAtMs = Number(lease.expiresAtMs || 0);
  if (expiresAtMs > 0) return expiresAtMs > nowMs;

  var heartbeatAtMs = Number(lease.heartbeatAtMs || lease.startedAtMs || 0);
  var ttlMs = Math.max(
    1000,
    Number(lease.ttlMs || AUTOMATION_RUNTIME_CONFIG.defaultLeaseTtlMs)
  );

  return heartbeatAtMs > 0 && nowMs - heartbeatAtMs < ttlMs;
}


function AUTOMATION_normalizeModuleKey_(moduleKey) {
  var value = String(moduleKey || '').trim().toUpperCase().replace(/[^A-Z0-9_]+/g, '_');
  if (!value) throw new Error('자동화 모듈 키가 비어 있습니다.');
  return value;
}


/****************************************************
 * 편집 재처리 큐 등록
 ****************************************************/

function AUTOMATION_deferEditEvent_(moduleKey, handlerName, e, reason, errorText) {
  var retryExecution = !!(e && e.__automationRetryExecution === true);
  var queueResult = null;

  if (!retryExecution) {
    try {
      queueResult = AUTOMATION_enqueueEditRetry_(
        moduleKey,
        handlerName,
        e,
        reason,
        errorText
      );
    } catch (queueErr) {
      queueResult = {
        queued: false,
        reason: 'QUEUE_WRITE_ERROR',
        error: AUTOMATION_runtimeErrorMessage_(queueErr)
      };
      console.error('[AUTOMATION_deferEditEvent_] ' + queueResult.error, queueErr);
    }
  }

  return {
    status: 'DEFERRED',
    module: AUTOMATION_normalizeModuleKey_(moduleKey),
    handler: String(handlerName || ''),
    reason: String(reason || 'DEFERRED'),
    error: String(errorText || ''),
    retryQueued: !!(queueResult && queueResult.queued),
    queueResult: queueResult,
    retryExecution: retryExecution
  };
}


function AUTOMATION_enqueueEditRetry_(moduleKey, handlerName, e, reason, errorText) {
  var descriptor = AUTOMATION_captureEditEvent_(e);

  if (!descriptor) {
    return {
      queued: false,
      reason: 'INVALID_EDIT_EVENT'
    };
  }

  var moduleName = AUTOMATION_normalizeModuleKey_(moduleKey);
  var safeHandlerName = String(handlerName || '').trim();

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(safeHandlerName)) {
    throw new Error('재처리 큐 핸들러명이 올바르지 않습니다: ' + safeHandlerName);
  }

  var writeLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_enqueueEditRetry_',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseTtlMs,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );

  if (!writeLease.acquired) {
    return {
      queued: false,
      reason: 'QUEUE_WRITE_LEASE_BUSY'
    };
  }

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var now = new Date();
    var nowIso = now.toISOString();
    var dedupeKey = AUTOMATION_buildRetryDedupeKey_(moduleName, safeHandlerName, descriptor);
    var existingRow = AUTOMATION_findPendingRetryRowByDedupeKey_(sheet, index, dedupeKey);
    var eventJson = JSON.stringify(descriptor);
    var shortenedError = AUTOMATION_truncateRuntimeText_(
      errorText || reason || '',
      AUTOMATION_RUNTIME_CONFIG.retryQueueMaxErrorLength
    );

    if (existingRow > 0) {
      sheet.getRange(existingRow, index['최근요청일시']).setValue(now);
      sheet.getRange(existingRow, index['다음시도일시']).setValue(now);
      sheet.getRange(existingRow, index['최근오류']).setValue(shortenedError);
      sheet.getRange(existingRow, index['이벤트JSON']).setValue(eventJson);
      sheet.getRange(existingRow, index['상태']).setValue('RETRY');

      return {
        queued: true,
        updatedExisting: true,
        row: existingRow,
        dedupeKey: dedupeKey
      };
    }

    var jobId = AUTOMATION_createRuntimeToken_('RETRY');
    var row = [
      jobId,
      dedupeKey,
      moduleName,
      safeHandlerName,
      'PENDING',
      descriptor.sourceSpreadsheetId,
      descriptor.sheetName,
      descriptor.startRow,
      descriptor.numRows,
      descriptor.startColumn,
      descriptor.numColumns,
      now,
      now,
      now,
      0,
      AUTOMATION_RUNTIME_CONFIG.retryQueueMaxAttempts,
      '',
      shortenedError,
      '',
      eventJson,
      AUTOMATION_RUNTIME_CONFIG.version
    ];

    sheet.getRange(sheet.getLastRow() + 1, 1, 1, row.length).setValues([row]);

    return {
      queued: true,
      updatedExisting: false,
      jobId: jobId,
      dedupeKey: dedupeKey
    };
  } finally {
    AUTOMATION_releaseModuleLease_(writeLease);
  }
}


function AUTOMATION_captureEditEvent_(e) {
  if (!e || !e.range || !e.source) return null;

  try {
    var range = e.range;
    var sheet = range.getSheet();

    return {
      sourceSpreadsheetId: String(e.source.getId() || ''),
      sheetName: String(sheet.getName() || ''),
      startRow: Number(range.getRow()) || 0,
      numRows: Math.max(1, Number(range.getNumRows()) || 1),
      startColumn: Number(range.getColumn()) || 0,
      numColumns: Math.max(1, Number(range.getNumColumns()) || 1),
      capturedAt: new Date().toISOString()
    };
  } catch (err) {
    return null;
  }
}


function AUTOMATION_buildRetryDedupeKey_(moduleKey, handlerName, descriptor) {
  return [
    moduleKey,
    handlerName,
    descriptor.sourceSpreadsheetId,
    descriptor.sheetName,
    descriptor.startRow,
    descriptor.numRows,
    descriptor.startColumn,
    descriptor.numColumns
  ].join('|');
}


function AUTOMATION_findPendingRetryRowByDedupeKey_(sheet, index, dedupeKey) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var values = sheet.getRange(2, 1, lastRow - 1, AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders.length).getDisplayValues();

  for (var i = values.length - 1; i >= 0; i--) {
    var status = String(values[i][index['상태'] - 1] || '').toUpperCase();
    var key = String(values[i][index['중복키'] - 1] || '');

    if (key !== dedupeKey) continue;
    if (status === 'PENDING' || status === 'RETRY') return i + 2;
  }

  return 0;
}


/****************************************************
 * 5분 파이프라인용 재처리 큐 소비
 ****************************************************/

function AUTOMATION_processEditRetryQueue_() {
  var startedAtMs = Date.now();
  var summary = {
    status: 'STARTED',
    reserved: 0,
    processed: 0,
    succeeded: 0,
    retried: 0,
    failed: 0,
    ignored: 0,
    errors: []
  };

  var processLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueProcessLeaseKey,
    {
      taskName: 'AUTOMATION_processEditRetryQueue_',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueProcessLeaseTtlMs,
      waitMs: 0
    }
  );

  if (!processLease.acquired) {
    summary.status = 'SKIPPED_ALREADY_RUNNING';
    return summary;
  }

  try {
    var jobs = AUTOMATION_reserveRetryJobs_();
    summary.reserved = jobs.length;

    for (var i = 0; i < jobs.length; i++) {
      if (Date.now() - startedAtMs >= AUTOMATION_RUNTIME_CONFIG.retryQueueMaxRuntimeMs) {
        AUTOMATION_returnReservedRetryJob_(jobs[i], 'RUN_TIME_LIMIT');
        for (var j = i + 1; j < jobs.length; j++) {
          AUTOMATION_returnReservedRetryJob_(jobs[j], 'RUN_TIME_LIMIT');
        }
        break;
      }

      var job = jobs[i];
      var result = AUTOMATION_executeRetryJob_(job);
      summary.processed++;

      if (result.outcome === 'DONE') summary.succeeded++;
      if (result.outcome === 'IGNORED') summary.ignored++;
      if (result.outcome === 'RETRY') summary.retried++;
      if (result.outcome === 'FAIL') summary.failed++;
      if (result.error) summary.errors.push(result.error);
    }

    summary.status = summary.failed > 0 || summary.retried > 0
      ? 'COMPLETED_WITH_PENDING'
      : 'COMPLETED';
  } catch (err) {
    summary.status = 'ERROR';
    summary.errors.push(AUTOMATION_runtimeErrorMessage_(err));
    console.error('[AUTOMATION_processEditRetryQueue_] ' + summary.errors[summary.errors.length - 1], err);
  } finally {
    AUTOMATION_releaseModuleLease_(processLease);
  }

  summary.durationMs = Date.now() - startedAtMs;
  return summary;
}


function AUTOMATION_reserveRetryJobs_() {
  var writeLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_reserveRetryJobs_',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseTtlMs,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );

  if (!writeLease.acquired) return [];

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var lastRow = sheet.getLastRow();
    var jobs = [];
    var nowMs = Date.now();

    if (lastRow < 2) return jobs;

    var values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();

    for (var i = 0; i < values.length; i++) {
      if (jobs.length >= AUTOMATION_RUNTIME_CONFIG.retryQueueMaxJobsPerRun) break;

      var row = values[i];
      var status = String(row[index['상태'] - 1] || '').toUpperCase();
      var lastAttemptMs = AUTOMATION_toTimeMs_(row[index['최근시도일시'] - 1]);
      var nextAttemptMs = AUTOMATION_toTimeMs_(row[index['다음시도일시'] - 1]);
      var isStaleRunning = status === 'RUNNING' && (
        !lastAttemptMs || nowMs - lastAttemptMs >= AUTOMATION_RUNTIME_CONFIG.retryQueueRunningStaleMs
      );

      if (status !== 'PENDING' && status !== 'RETRY' && !isStaleRunning) continue;
      if (nextAttemptMs && nextAttemptMs > nowMs) continue;

      var rowNo = i + 2;
      var attempts = Math.max(0, Number(row[index['시도횟수'] - 1]) || 0) + 1;
      var job = AUTOMATION_retryRowToJob_(row, index, rowNo);
      job.attempts = attempts;

      sheet.getRange(rowNo, index['상태']).setValue('RUNNING');
      sheet.getRange(rowNo, index['시도횟수']).setValue(attempts);
      sheet.getRange(rowNo, index['최근시도일시']).setValue(new Date());
      sheet.getRange(rowNo, index['최근오류']).setValue('');

      jobs.push(job);
    }

    return jobs;
  } finally {
    AUTOMATION_releaseModuleLease_(writeLease);
  }
}


function AUTOMATION_retryRowToJob_(row, index, rowNo) {
  var descriptor = null;
  var rawEvent = String(row[index['이벤트JSON'] - 1] || '');

  try {
    descriptor = rawEvent ? JSON.parse(rawEvent) : null;
  } catch (err) {
    descriptor = null;
  }

  return {
    rowNo: rowNo,
    jobId: String(row[index['작업ID'] - 1] || ''),
    module: String(row[index['모듈'] - 1] || ''),
    handler: String(row[index['핸들러'] - 1] || ''),
    maxAttempts: Math.max(1, Number(row[index['최대시도'] - 1]) || AUTOMATION_RUNTIME_CONFIG.retryQueueMaxAttempts),
    descriptor: descriptor
  };
}


function AUTOMATION_executeRetryJob_(job) {
  var outcome = 'RETRY';
  var errorText = '';

  try {
    if (!job.descriptor) {
      throw new Error('재처리 이벤트 JSON이 없거나 손상되었습니다.');
    }

    var handler = AUTOMATION_resolveHandler_(job.handler);
    if (typeof handler !== 'function') {
      throw new Error('재처리 핸들러를 찾을 수 없습니다: ' + job.handler);
    }

    var e = AUTOMATION_rebuildEditEvent_(job.descriptor);
    var result = handler(e);
    var status = String(result && result.status || 'SUCCESS').toUpperCase();

    if (status === 'DEFERRED' || status === 'ERROR' || status === 'LEASE_BUSY') {
      errorText = String(
        result && (result.error || result.reason) || '재처리 핸들러가 다시 지연되었습니다.'
      );
      outcome = job.attempts >= job.maxAttempts ? 'FAIL' : 'RETRY';
    } else if (status.indexOf('IGNORED') === 0) {
      outcome = 'IGNORED';
    } else {
      outcome = 'DONE';
    }
  } catch (err) {
    errorText = AUTOMATION_runtimeErrorMessage_(err);
    outcome = job.attempts >= job.maxAttempts ? 'FAIL' : 'RETRY';
  }

  AUTOMATION_finalizeRetryJob_(job, outcome, errorText);

  return {
    jobId: job.jobId,
    outcome: outcome,
    error: errorText
  };
}


function AUTOMATION_rebuildEditEvent_(descriptor) {
  var source = SpreadsheetApp.openById(String(descriptor.sourceSpreadsheetId || ''));
  var sheet = source.getSheetByName(String(descriptor.sheetName || ''));

  if (!sheet) {
    throw new Error('재처리 대상 시트를 찾을 수 없습니다: ' + descriptor.sheetName);
  }

  var range = sheet.getRange(
    Number(descriptor.startRow),
    Number(descriptor.startColumn),
    Math.max(1, Number(descriptor.numRows) || 1),
    Math.max(1, Number(descriptor.numColumns) || 1)
  );

  return {
    source: source,
    range: range,
    __automationRetryExecution: true,
    __automationRetryCapturedAt: descriptor.capturedAt || ''
  };
}


function AUTOMATION_finalizeRetryJob_(job, outcome, errorText) {
  var writeLease = AUTOMATION_acquireModuleLease_(
    AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseKey,
    {
      taskName: 'AUTOMATION_finalizeRetryJob_',
      ttlMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteLeaseTtlMs,
      waitMs: AUTOMATION_RUNTIME_CONFIG.retryQueueWriteWaitMs
    }
  );

  if (!writeLease.acquired) {
    console.error('[AUTOMATION_finalizeRetryJob_] 큐 쓰기 lease를 얻지 못했습니다: ' + job.jobId);
    return false;
  }

  try {
    var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
    var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
    var index = AUTOMATION_makeHeaderIndex_(headers);
    var rowNo = AUTOMATION_findRetryJobRowById_(sheet, index, job.jobId);

    if (!rowNo) return false;

    var now = new Date();
    var status = outcome === 'DONE' || outcome === 'IGNORED' ? 'DONE' : outcome;
    var nextAttempt = '';

    if (outcome === 'RETRY') {
      nextAttempt = new Date(Date.now() + AUTOMATION_retryBackoffMs_(job.attempts));
    }

    sheet.getRange(rowNo, index['상태']).setValue(status);
    sheet.getRange(rowNo, index['다음시도일시']).setValue(nextAttempt);
    sheet.getRange(rowNo, index['최근오류']).setValue(
      AUTOMATION_truncateRuntimeText_(errorText || '', AUTOMATION_RUNTIME_CONFIG.retryQueueMaxErrorLength)
    );
    sheet.getRange(rowNo, index['완료일시']).setValue(
      status === 'DONE' || status === 'FAIL' ? now : ''
    );

    return true;
  } finally {
    AUTOMATION_releaseModuleLease_(writeLease);
  }
}


function AUTOMATION_returnReservedRetryJob_(job, reason) {
  return AUTOMATION_finalizeRetryJob_(job, 'RETRY', String(reason || 'RETRY'));
}


function AUTOMATION_retryBackoffMs_(attempts) {
  var step = Math.max(0, Math.min(5, Number(attempts || 1) - 1));
  return Math.min(30 * 60 * 1000, 60 * 1000 * Math.pow(2, step));
}


/****************************************************
 * 재처리 큐 표시/관리
 ****************************************************/

function AUTOMATION_showRetryQueueSheet() {
  var sheet = AUTOMATION_getOrCreateRetryQueueSheet_();
  sheet.showSheet();
  sheet.getParent().setActiveSheet(sheet);
  return sheet.getName();
}


function AUTOMATION_retryQueueNow() {
  var result = AUTOMATION_processEditRetryQueue_();
  SpreadsheetApp.getUi().alert(
    '자동화 재처리 큐',
    [
      '상태: ' + result.status,
      '예약: ' + result.reserved,
      '처리: ' + result.processed,
      '성공: ' + result.succeeded,
      '재시도 대기: ' + result.retried,
      '최종 실패: ' + result.failed,
      '무시: ' + result.ignored
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );
  return result;
}


function AUTOMATION_getOrCreateRetryQueueSheet_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var name = AUTOMATION_RUNTIME_CONFIG.retryQueueSheetName;
  var sheet = ss.getSheetByName(name);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }

  var headers = AUTOMATION_RUNTIME_CONFIG.retryQueueHeaders;
  var current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  var mismatch = headers.some(function(header, index) {
    return String(current[index] || '') !== header;
  });

  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    sheet.autoResizeColumns(1, headers.length);
  }

  if (created) {
    try {
      sheet.hideSheet();
    } catch (ignoreHideError) {
      // 숨김 실패는 큐 기능과 무관하므로 무시
    }
  }

  return sheet;
}


/**
 * 백그라운드 자동화가 사용할 영업관리대장 ID를 명시 설정에서만 결정한다.
 * 시간 트리거에서는 활성 스프레드시트가 없거나 예상과 다를 수 있으므로
 * getActive()/getActiveSpreadsheet() 폴백을 절대 사용하지 않는다.
 */
function AUTOMATION_getRuntimeMasterSpreadsheetId_() {
  if (AUTOMATION_RUNTIME_MASTER_SPREADSHEET_ID_CACHE) {
    return AUTOMATION_RUNTIME_MASTER_SPREADSHEET_ID_CACHE;
  }

  var props = PropertiesService.getScriptProperties();
  var propertyKey = typeof PROP_MASTER_SPREADSHEET_ID !== 'undefined'
    ? PROP_MASTER_SPREADSHEET_ID
    : 'MASTER_SPREADSHEET_ID';
  var spreadsheetId = '';
  var source = '';

  // 코드에 고정된 운영 기준값을 최우선으로 사용한다.
  // 과거 Script Property가 다른 접근 가능한 파일을 가리키더라도 잘못 채택하지 않는다.
  try {
    if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.MASTER_SPREADSHEET_ID) {
      spreadsheetId = String(CONFIG.MASTER_SPREADSHEET_ID || '').trim();
      source = 'CONFIG.MASTER_SPREADSHEET_ID';
    }
  } catch (ignoreConfigIdError) {
    spreadsheetId = '';
  }

  if (!spreadsheetId) {
    spreadsheetId = String(props.getProperty(propertyKey) || '').trim();
    source = spreadsheetId ? ('Script Properties.' + propertyKey) : '';
  }

  if (!spreadsheetId) {
    try {
      if (typeof CUSTOMER_FOLDER_CFG !== 'undefined' && CUSTOMER_FOLDER_CFG) {
        spreadsheetId = String(CUSTOMER_FOLDER_CFG.MASTER_SPREADSHEET_ID || '').trim();
        source = spreadsheetId ? 'CUSTOMER_FOLDER_CFG.MASTER_SPREADSHEET_ID' : '';
      }
    } catch (ignoreCustomerFolderConfigError) {
      spreadsheetId = '';
    }
  }

  if (!spreadsheetId) {
    throw new Error(
      '영업관리대장 스프레드시트 ID를 명시 설정에서 확인할 수 없습니다. ' +
      'CONFIG.MASTER_SPREADSHEET_ID 또는 Script Properties의 ' + propertyKey +
      '를 확인하세요.'
    );
  }

  AUTOMATION_RUNTIME_MASTER_SPREADSHEET_ID_CACHE = spreadsheetId;
  props.setProperty(propertyKey, spreadsheetId);
  console.log('[AUTOMATION_getRuntimeMasterSpreadsheetId_] ' + source + '=' + spreadsheetId);
  return spreadsheetId;
}


function AUTOMATION_getRuntimeMasterSpreadsheet_() {
  var spreadsheetId = AUTOMATION_getRuntimeMasterSpreadsheetId_();

  if (
    AUTOMATION_RUNTIME_MASTER_SPREADSHEET_CACHE &&
    AUTOMATION_RUNTIME_MASTER_SPREADSHEET_ID_CACHE === spreadsheetId
  ) {
    return AUTOMATION_RUNTIME_MASTER_SPREADSHEET_CACHE;
  }

  AUTOMATION_RUNTIME_MASTER_SPREADSHEET_CACHE = SpreadsheetApp.openById(spreadsheetId);
  AUTOMATION_RUNTIME_MASTER_SPREADSHEET_ID_CACHE = spreadsheetId;
  return AUTOMATION_RUNTIME_MASTER_SPREADSHEET_CACHE;
}


/**
 * 시간 트리거가 사용하는 주요 스프레드시트 바인딩을 비파괴 검증한다.
 */
function AUTOMATION_verifyBackgroundSpreadsheetBindings() {
  var master = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var result = {
    status: 'SUCCESS',
    masterSpreadsheetId: master.getId(),
    masterSpreadsheetName: master.getName(),
    masterRequiredSheets: {},
    externalSpreadsheets: [],
    checkedAt: new Date().toISOString(),
    errors: []
  };

  function checkExternal_(label, spreadsheetId, requiredSheetName) {
    var id = String(spreadsheetId || '').trim();
    if (!id) {
      result.errors.push(label + ' ID 누락');
      return;
    }

    try {
      var spreadsheet = SpreadsheetApp.openById(id);
      var hasRequiredSheet = requiredSheetName
        ? !!spreadsheet.getSheetByName(requiredSheetName)
        : true;

      result.externalSpreadsheets.push({
        label: label,
        spreadsheetId: id,
        spreadsheetName: spreadsheet.getName(),
        requiredSheetName: String(requiredSheetName || ''),
        requiredSheetFound: hasRequiredSheet
      });

      if (!hasRequiredSheet) {
        result.errors.push(label + ' 시트 누락: ' + requiredSheetName);
      }
    } catch (err) {
      result.errors.push(label + ' 접근 실패: ' + AUTOMATION_runtimeErrorMessage_(err));
    }
  }

  ['마스터시트(신규)', '수주확정/계약완료', '영업지원요청'].forEach(function(name) {
    result.masterRequiredSheets[name] = !!master.getSheetByName(name);
    if (!result.masterRequiredSheets[name]) {
      result.errors.push('영업관리대장 시트 누락: ' + name);
    }
  });

  try {
    var kjDocId = typeof KJ_DOC_CONFIG !== 'undefined' && KJ_DOC_CONFIG
      ? KJ_DOC_CONFIG.SPREADSHEET_ID
      : '';
    checkExternal_('KJ 고객관리', kjDocId, '고객관리');
  } catch (kjConfigErr) {
    result.errors.push('KJ 고객관리 설정 확인 실패: ' + AUTOMATION_runtimeErrorMessage_(kjConfigErr));
  }

  try {
    if (typeof TARGET_FILES !== 'undefined' && TARGET_FILES) {
      checkExternal_('KJ 수행사 고객관리', TARGET_FILES.KJ, '고객관리');
      checkExternal_('일신 수행사 고객관리', TARGET_FILES['일신'], '고객관리');
    }
  } catch (vendorConfigErr) {
    result.errors.push('수행사 파일 설정 확인 실패: ' + AUTOMATION_runtimeErrorMessage_(vendorConfigErr));
  }

  try {
    if (typeof ITMAINT_getConfig_2026_ === 'function') {
      var itConfig = ITMAINT_getConfig_2026_();
      checkExternal_(
        '2026정보통신유지보수',
        itConfig.targetSpreadsheetId,
        itConfig.targetSheetName
      );
    }
  } catch (itConfigErr) {
    result.errors.push('정보통신유지보수 설정 확인 실패: ' + AUTOMATION_runtimeErrorMessage_(itConfigErr));
  }

  try {
    if (typeof CONFIG !== 'undefined' && CONFIG) {
      checkExternal_('메일 파일생성기', CONFIG.GENERATOR_SPREADSHEET_ID, '생성대상');
    }
  } catch (generatorConfigErr) {
    result.errors.push('메일 파일생성기 설정 확인 실패: ' + AUTOMATION_runtimeErrorMessage_(generatorConfigErr));
  }

  if (result.errors.length) result.status = 'ERROR';
  console.log('[AUTOMATION_verifyBackgroundSpreadsheetBindings] ' + JSON.stringify(result));

  try {
    SpreadsheetApp.getUi().alert(
      '백그라운드 파일 바인딩 검증',
      [
        '상태: ' + result.status,
        '영업관리대장: ' + result.masterSpreadsheetName + ' (' + result.masterSpreadsheetId + ')',
        '외부 파일 확인: ' + result.externalSpreadsheets.length + '개',
        '오류: ' + result.errors.length + '건',
        result.errors.length ? ('\n' + result.errors.join('\n')) : '\n모든 명시 파일·필수 시트에 접근 가능합니다.'
      ].join('\n'),
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  } catch (ignoreUiError) {
    // 편집기 직접 실행이나 백그라운드 검증에서는 반환값·로그로 확인한다.
  }

  return result;
}


function AUTOMATION_makeHeaderIndex_(headers) {
  var index = {};
  headers.forEach(function(header, i) {
    index[header] = i + 1;
  });
  return index;
}


function AUTOMATION_findRetryJobRowById_(sheet, index, jobId) {
  var lastRow = sheet.getLastRow();
  if (lastRow < 2) return 0;

  var values = sheet.getRange(2, index['작업ID'], lastRow - 1, 1).getDisplayValues();
  for (var i = 0; i < values.length; i++) {
    if (String(values[i][0] || '') === String(jobId || '')) return i + 2;
  }
  return 0;
}


function AUTOMATION_toTimeMs_(value) {
  if (!value) return 0;
  if (Object.prototype.toString.call(value) === '[object Date]') return value.getTime();
  var parsed = new Date(value).getTime();
  return isFinite(parsed) ? parsed : 0;
}


function AUTOMATION_createRuntimeToken_(prefix) {
  try {
    return String(prefix || 'AUTO') + '-' + Utilities.getUuid();
  } catch (ignoreUuidError) {
    return String(prefix || 'AUTO') + '-' + Date.now() + '-' + Math.random().toString(36).slice(2, 10);
  }
}


function AUTOMATION_makeRuntimeJsonSafe_(value) {
  if (value === null || typeof value === 'undefined') return null;
  try {
    return JSON.parse(JSON.stringify(value));
  } catch (err) {
    return String(value);
  }
}


function AUTOMATION_runtimeErrorMessage_(err) {
  if (!err) return '알 수 없는 오류';
  return err.message ? String(err.message) : String(err);
}


function AUTOMATION_truncateRuntimeText_(value, maxLength) {
  var text = String(value || '');
  var limit = Math.max(1, Number(maxLength) || 1);
  return text.length <= limit ? text : text.slice(0, limit - 3) + '...';
}
