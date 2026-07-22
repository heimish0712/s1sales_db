/****************************************************
 * 영업지원요청 G열 변경 감지 → Discord 알림
 *
 * 감시 시트: 영업지원요청
 * 기준 ID: A열 접수번호
 * 감시 값: G열
 *
 * 9단계 안정화:
 * - 대용량 단일 Script Property 상태를 숨김 시트로 이전
 * - Discord 2xx 성공 건만 상태 확정
 * - 429/4xx/5xx/네트워크 오류는 상태를 확정하지 않고 다음 실행 재시도
 * - 1회 최대 20건 및 Discord 본문 길이 제한
 * - 구형 상태 JSON 자동 마이그레이션/손상 복구
 *
 * 트리거는 AUTOMATION_executeCanonicalCutover()로 중앙 설치·관리함.
 ****************************************************/

var SALES_SUPPORT_ALERT_CONFIG = Object.freeze({
  SHEET_NAME: '영업지원요청',
  START_ROW: 2,
  END_ROW: 1000,

  ID_COL: 1,       // A열: 접수번호
  WATCH_COL: 7,    // G열: 감시할 내용

  // 구형 단일 JSON 상태. V2 시트 마이그레이션 후 삭제함.
  STATE_PROP_KEY: 'SALES_SUPPORT_LAST_ID_G_MAP',
  WEBHOOK_PROP_KEY: 'SALES_SUPPORT_DISCORD_WEBHOOK_URL',

  STATE_SHEET_NAME: '_영업지원Discord상태',
  STATE_INIT_PROP_KEY: 'SALES_SUPPORT_DISCORD_STATE_V2_INITIALIZED',
  LAST_RUN_PROP_KEY: 'SALES_SUPPORT_DISCORD_LAST_RUN_V2',
  STATE_HEADERS: Object.freeze(['접수번호', '확정G값', '최근확정일시']),

  LEASE_KEY: 'DISCORD_SALES_SUPPORT',
  MAX_ALERT_ITEMS: 20,
  MAX_DISCORD_CONTENT_CHARS: 1900,
  MAX_VALUE_CHARS: 240,
  MAX_ERROR_CHARS: 1000
});


/**
 * 최초 1회 실행.
 * Discord 웹훅 URL을 Script Properties에 저장함.
 *
 * 웹훅 URL은 기존 운영값을 유지한다.
 */
function setSalesSupportWebhookUrlOnce() {
  var webhookUrl = 'https://discord.com/api/webhooks/1516614305305329785/7SVOGX2cOSDIyetVyPF9SGlCFj5HRNVFYASGhUvKWakGWn8xGmNt7hUfpfDWL94zcaUF';

  PropertiesService
    .getScriptProperties()
    .setProperty(SALES_SUPPORT_ALERT_CONFIG.WEBHOOK_PROP_KEY, webhookUrl);

  Logger.log('Discord 웹훅 URL 저장 완료');
}


/**
 * 현재 상태를 알림 완료 상태로 강제 저장한다.
 * 최초 도입 또는 상태를 다시 기준화할 때 수동 실행한다.
 */
function saveSalesSupportCurrentState() {
  return SALES_SUPPORT_runWithLease_(function() {
    var currentMap = getSalesSupportCurrentIdGMap_();
    var nowIso = new Date().toISOString();
    var stateMap = {};

    Object.keys(currentMap).forEach(function(id) {
      var value = String(currentMap[id] || '');
      if (!value) return;

      stateMap[id] = {
        value: value,
        confirmedAt: nowIso
      };
    });

    SALES_SUPPORT_writeStateMap_(stateMap);

    var props = PropertiesService.getScriptProperties();
    props.setProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_INIT_PROP_KEY, '1');
    props.deleteProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);

    SALES_SUPPORT_recordLastRun_({
      status: 'STATE_SAVED',
      currentCount: Object.keys(currentMap).length,
      storedCount: Object.keys(stateMap).length
    });

    Logger.log('영업지원요청 현재 상태 저장 완료: ' + Object.keys(stateMap).length + '건');

    return {
      status: 'STATE_SAVED',
      currentCount: Object.keys(currentMap).length,
      storedCount: Object.keys(stateMap).length
    };
  }, 'saveSalesSupportCurrentState');
}


/**
 * 실제 감시 함수.
 * 시간 기반 트리거가 1분마다 실행한다.
 */
function checkSalesSupportNewValues() {
  return SALES_SUPPORT_runWithLease_(function() {
    var props = PropertiesService.getScriptProperties();
    var currentMap = getSalesSupportCurrentIdGMap_();
    var stateResult = SALES_SUPPORT_loadStateMap_(currentMap);

    if (stateResult.initializedNow) {
      var initializedResult = {
        status: stateResult.status,
        currentCount: Object.keys(currentMap).length,
        storedCount: Object.keys(stateResult.map).length,
        pendingCount: 0,
        sentCount: 0
      };

      SALES_SUPPORT_recordLastRun_(initializedResult);
      return initializedResult;
    }

    var stateMap = stateResult.map;
    var stateDirty = Boolean(stateResult.needsRewrite);
    var changedItems = [];

    // 삭제된 접수번호와 G열이 비워진 접수번호는 상태에서 제거한다.
    // 이후 같은 값이 다시 입력되면 신규 변경으로 감지된다.
    Object.keys(stateMap).forEach(function(id) {
      var currentValue = Object.prototype.hasOwnProperty.call(currentMap, id)
        ? String(currentMap[id] || '')
        : '';

      if (!currentValue) {
        delete stateMap[id];
        stateDirty = true;
      }
    });

    Object.keys(currentMap).forEach(function(id) {
      var currentValue = String(currentMap[id] || '');
      if (!currentValue) return;

      var oldEntry = stateMap[id] || null;
      var oldValue = oldEntry ? String(oldEntry.value || '') : '';

      if (currentValue !== oldValue) {
        changedItems.push({
          id: id,
          oldValue: oldValue,
          newValue: currentValue
        });
      }
    });

    if (changedItems.length === 0) {
      if (stateDirty) SALES_SUPPORT_writeStateMap_(stateMap);

      var noChangeResult = {
        status: 'NO_CHANGES',
        currentCount: Object.keys(currentMap).length,
        storedCount: Object.keys(stateMap).length,
        pendingCount: 0,
        sentCount: 0,
        repairedRows: Number(stateResult.repairedRows || 0)
      };

      SALES_SUPPORT_recordLastRun_(noChangeResult);
      return noChangeResult;
    }

    var batch = SALES_SUPPORT_buildDiscordBatch_(changedItems);
    var sendResult = sendSalesSupportDiscordAlert_(
      batch.items,
      changedItems.length - batch.items.length
    );

    if (!sendResult.success) {
      var failedResult = {
        status: 'SEND_FAILED_RETRY_PENDING',
        currentCount: Object.keys(currentMap).length,
        storedCount: Object.keys(stateMap).length,
        pendingCount: changedItems.length,
        sentCount: 0,
        responseCode: sendResult.responseCode,
        retryAfterMs: sendResult.retryAfterMs,
        error: sendResult.error
      };

      // 전송 실패 시 변경 상태는 절대 확정하지 않는다.
      // 다만 삭제/공백 상태 정리는 저장해 상태 꼬임을 방지한다.
      if (stateDirty) SALES_SUPPORT_writeStateMap_(stateMap);

      SALES_SUPPORT_recordLastRun_(failedResult);
      Logger.log('Discord 알림 전송 실패. 다음 실행에서 재시도: ' + JSON.stringify(failedResult));
      return failedResult;
    }

    var confirmedAt = new Date().toISOString();
    batch.items.forEach(function(item) {
      stateMap[item.id] = {
        value: item.newValue,
        confirmedAt: confirmedAt
      };
    });

    SALES_SUPPORT_writeStateMap_(stateMap);
    props.setProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_INIT_PROP_KEY, '1');

    var successResult = {
      status: changedItems.length > batch.items.length
        ? 'SENT_PARTIAL_REMAINING'
        : 'SENT',
      currentCount: Object.keys(currentMap).length,
      storedCount: Object.keys(stateMap).length,
      pendingCount: changedItems.length - batch.items.length,
      sentCount: batch.items.length,
      responseCode: sendResult.responseCode,
      contentLength: batch.contentLength,
      repairedRows: Number(stateResult.repairedRows || 0)
    };

    SALES_SUPPORT_recordLastRun_(successResult);
    return successResult;
  }, 'checkSalesSupportNewValues');
}


/**
 * 영업지원요청 시트에서 A열 접수번호와 G열 값을 맵으로 가져온다.
 */
function getSalesSupportCurrentIdGMap_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var sheet = ss && ss.getSheetByName(SALES_SUPPORT_ALERT_CONFIG.SHEET_NAME);

  if (!sheet) {
    throw new Error('시트를 찾을 수 없음: ' + SALES_SUPPORT_ALERT_CONFIG.SHEET_NAME);
  }

  var startRow = SALES_SUPPORT_ALERT_CONFIG.START_ROW;
  var endRow = SALES_SUPPORT_ALERT_CONFIG.END_ROW;
  var watchCol = SALES_SUPPORT_ALERT_CONFIG.WATCH_COL;
  var idCol = SALES_SUPPORT_ALERT_CONFIG.ID_COL;
  var numRows = endRow - startRow + 1;

  var values = sheet
    .getRange(startRow, 1, numRows, watchCol)
    .getDisplayValues();

  var map = {};

  values.forEach(function(row) {
    var id = String(row[idCol - 1] || '').trim();
    var gValue = String(row[watchCol - 1] || '').trim();

    if (!id) return;
    map[id] = gValue;
  });

  return map;
}


/**
 * Discord로 알림을 발송하고 HTTP 성공 여부를 반환한다.
 * 2xx만 성공으로 인정한다.
 */
function sendSalesSupportDiscordAlert_(items, remainingCount) {
  var webhookUrl = PropertiesService
    .getScriptProperties()
    .getProperty(SALES_SUPPORT_ALERT_CONFIG.WEBHOOK_PROP_KEY);

  if (!webhookUrl) {
    return {
      success: false,
      responseCode: 0,
      retryAfterMs: 0,
      error: 'Discord 웹훅 URL이 저장되어 있지 않습니다. setSalesSupportWebhookUrlOnce()를 실행하십시오.'
    };
  }

  var content = SALES_SUPPORT_buildDiscordContent_(items, remainingCount);

  try {
    var response = UrlFetchApp.fetch(webhookUrl, {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify({ content: content }),
      muteHttpExceptions: true
    });

    var responseCode = Number(response.getResponseCode()) || 0;
    var responseBody = String(response.getContentText() || '');
    var headers = {};

    try {
      headers = response.getAllHeaders ? response.getAllHeaders() : {};
    } catch (ignoreHeadersError) {
      headers = {};
    }

    if (responseCode >= 200 && responseCode < 300) {
      return {
        success: true,
        responseCode: responseCode,
        retryAfterMs: 0,
        contentLength: content.length
      };
    }

    return {
      success: false,
      responseCode: responseCode,
      retryAfterMs: SALES_SUPPORT_extractRetryAfterMs_(responseBody, headers),
      error: SALES_SUPPORT_limitText_(
        'Discord HTTP ' + responseCode + ': ' + responseBody,
        SALES_SUPPORT_ALERT_CONFIG.MAX_ERROR_CHARS
      )
    };
  } catch (error) {
    return {
      success: false,
      responseCode: 0,
      retryAfterMs: 0,
      error: SALES_SUPPORT_limitText_(
        error && error.stack ? error.stack : String(error),
        SALES_SUPPORT_ALERT_CONFIG.MAX_ERROR_CHARS
      )
    };
  }
}


/**
 * 알림 1회에 포함할 건수를 Discord 길이 제한 안에서 결정한다.
 */
function SALES_SUPPORT_buildDiscordBatch_(items) {
  var maxItems = SALES_SUPPORT_ALERT_CONFIG.MAX_ALERT_ITEMS;
  var selected = [];

  for (var i = 0; i < items.length && selected.length < maxItems; i++) {
    var candidate = selected.concat([items[i]]);
    var remaining = items.length - candidate.length;
    var content = SALES_SUPPORT_buildDiscordContent_(candidate, remaining);

    if (
      selected.length > 0 &&
      content.length > SALES_SUPPORT_ALERT_CONFIG.MAX_DISCORD_CONTENT_CHARS
    ) {
      break;
    }

    selected = candidate;
  }

  if (selected.length === 0 && items.length > 0) {
    selected = [items[0]];
  }

  var finalContent = SALES_SUPPORT_buildDiscordContent_(
    selected,
    Math.max(0, items.length - selected.length)
  );

  return {
    items: selected,
    contentLength: finalContent.length
  };
}


function SALES_SUPPORT_buildDiscordContent_(items, remainingCount) {
  var lines = items.map(function(item) {
    var before = item.oldValue ? item.oldValue : '빈칸';

    return [
      '• 접수번호 ' + SALES_SUPPORT_limitText_(item.id, 100),
      '  기존: ' + SALES_SUPPORT_limitText_(before, SALES_SUPPORT_ALERT_CONFIG.MAX_VALUE_CHARS),
      '  신규: ' + SALES_SUPPORT_limitText_(item.newValue, SALES_SUPPORT_ALERT_CONFIG.MAX_VALUE_CHARS)
    ].join('\n');
  }).join('\n\n');

  var remainingText = Number(remainingCount) > 0
    ? '\n\n외 ' + Number(remainingCount) + '건은 다음 1분 실행에서 이어서 알림합니다.'
    : '';

  return '🔔 새로운 영업지원요청이 들어왔어여 ㅋㅋ\n\n' + lines + remainingText;
}


/**
 * 숨김 상태 시트를 읽는다.
 * 최초 실행이면 기존 Script Property를 마이그레이션하고,
 * 마이그레이션할 상태가 없거나 손상됐다면 현재 상태로 조용히 초기화한다.
 */
function SALES_SUPPORT_loadStateMap_(currentMap) {
  var props = PropertiesService.getScriptProperties();
  var sheet = SALES_SUPPORT_getOrCreateStateSheet_();
  var result = SALES_SUPPORT_readStateSheet_(sheet);
  var initialized = props.getProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_INIT_PROP_KEY) === '1';
  var legacyRaw = props.getProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);

  if (!initialized) {
    // V2 상태 시트는 남아 있는데 초기화 마커만 지워진 경우,
    // 기존 확정 상태를 버리지 않고 마커만 복구한다.
    if (Object.keys(result.map).length > 0) {
      props.setProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_INIT_PROP_KEY, '1');
      props.deleteProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);

      return {
        map: result.map,
        initializedNow: false,
        status: 'STATE_MARKER_REPAIRED',
        needsRewrite: result.needsRewrite,
        repairedRows: result.repairedRows
      };
    }

    var legacyMap = SALES_SUPPORT_parseLegacyStateMap_(legacyRaw);

    if (legacyMap) {
      var migratedMap = {};
      var migratedAt = new Date().toISOString();

      Object.keys(legacyMap).forEach(function(id) {
        var value = String(legacyMap[id] || '');
        if (!value) return;
        migratedMap[id] = { value: value, confirmedAt: migratedAt };
      });

      SALES_SUPPORT_writeStateMap_(migratedMap);
      props.setProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_INIT_PROP_KEY, '1');
      props.deleteProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);

      return {
        map: migratedMap,
        initializedNow: true,
        status: 'LEGACY_STATE_MIGRATED',
        needsRewrite: false,
        repairedRows: result.repairedRows
      };
    }

    // 상태가 없거나 구형 JSON이 손상된 경우 기존 데이터 폭탄을 막기 위해
    // 현재 값을 기준 상태로 저장하고 이번 실행에서는 알림하지 않는다.
    var seededMap = {};
    var seededAt = new Date().toISOString();

    Object.keys(currentMap).forEach(function(id) {
      var value = String(currentMap[id] || '');
      if (!value) return;
      seededMap[id] = { value: value, confirmedAt: seededAt };
    });

    SALES_SUPPORT_writeStateMap_(seededMap);
    props.setProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_INIT_PROP_KEY, '1');
    props.deleteProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);

    return {
      map: seededMap,
      initializedNow: true,
      status: legacyRaw ? 'CORRUPTED_LEGACY_STATE_RESEEDED' : 'STATE_INITIALIZED',
      needsRewrite: false,
      repairedRows: result.repairedRows
    };
  }

  if (legacyRaw) {
    props.deleteProperty(SALES_SUPPORT_ALERT_CONFIG.STATE_PROP_KEY);
  }

  return {
    map: result.map,
    initializedNow: false,
    status: 'STATE_LOADED',
    needsRewrite: result.needsRewrite,
    repairedRows: result.repairedRows
  };
}


function SALES_SUPPORT_getOrCreateStateSheet_() {
  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var name = SALES_SUPPORT_ALERT_CONFIG.STATE_SHEET_NAME;
  var sheet = ss.getSheetByName(name);
  var created = false;

  if (!sheet) {
    sheet = ss.insertSheet(name);
    created = true;
  }

  var headers = SALES_SUPPORT_ALERT_CONFIG.STATE_HEADERS;
  var currentHeaders = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  var mismatch = headers.some(function(header, index) {
    return String(currentHeaders[index] || '') !== header;
  });

  if (mismatch) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
  }

  if (created) {
    try {
      sheet.hideSheet();
    } catch (ignoreHideError) {
      // 숨김 실패는 알림 처리와 무관하므로 무시
    }
  }

  return sheet;
}


function SALES_SUPPORT_readStateSheet_(sheet) {
  var lastRow = sheet.getLastRow();

  if (lastRow < 2) {
    return { map: {}, needsRewrite: false, repairedRows: 0 };
  }

  var rows = sheet.getRange(2, 1, lastRow - 1, 3).getDisplayValues();
  var map = {};
  var invalidRows = 0;
  var duplicateRows = 0;

  rows.forEach(function(row) {
    var id = String(row[0] || '').trim();
    var value = String(row[1] || '').trim();
    var confirmedAt = String(row[2] || '').trim();

    if (!id || !value) {
      if (id || value || confirmedAt) invalidRows++;
      return;
    }

    if (Object.prototype.hasOwnProperty.call(map, id)) duplicateRows++;

    map[id] = {
      value: value,
      confirmedAt: confirmedAt || ''
    };
  });

  return {
    map: map,
    needsRewrite: invalidRows > 0 || duplicateRows > 0,
    repairedRows: invalidRows + duplicateRows
  };
}


function SALES_SUPPORT_writeStateMap_(stateMap) {
  var sheet = SALES_SUPPORT_getOrCreateStateSheet_();
  var ids = Object.keys(stateMap).sort();
  var rows = ids.map(function(id) {
    var entry = stateMap[id] || {};
    return [
      id,
      String(entry.value || ''),
      String(entry.confirmedAt || '')
    ];
  });
  var lastRow = sheet.getLastRow();

  if (lastRow > 1) {
    sheet.getRange(2, 1, lastRow - 1, 3).clearContent();
  }

  if (rows.length > 0) {
    sheet.getRange(2, 1, rows.length, 3).setValues(rows);
  }
}


function SALES_SUPPORT_parseLegacyStateMap_(raw) {
  if (!raw) return null;

  try {
    var parsed = JSON.parse(raw);
    if (!parsed || typeof parsed !== 'object' || Array.isArray(parsed)) return null;
    return parsed;
  } catch (ignoreParseError) {
    return null;
  }
}


function SALES_SUPPORT_extractRetryAfterMs_(responseBody, headers) {
  var bodyRetry = 0;

  try {
    var parsed = JSON.parse(responseBody || '{}');
    bodyRetry = Number(parsed.retry_after || parsed.retryAfter || 0);
  } catch (ignoreJsonError) {
    bodyRetry = 0;
  }

  if (bodyRetry > 0) {
    // Discord는 환경에 따라 초 또는 밀리초로 응답할 수 있어 작은 값은 초로 간주한다.
    return bodyRetry < 1000 ? Math.round(bodyRetry * 1000) : Math.round(bodyRetry);
  }

  headers = headers || {};
  var headerValue = headers['Retry-After'] || headers['retry-after'] || 0;
  var headerNumber = Number(headerValue) || 0;
  return headerNumber > 0 ? Math.round(headerNumber * 1000) : 0;
}


function SALES_SUPPORT_recordLastRun_(result) {
  var record = {
    recordedAt: new Date().toISOString(),
    status: String(result && result.status || ''),
    currentCount: Number(result && result.currentCount || 0),
    storedCount: Number(result && result.storedCount || 0),
    pendingCount: Number(result && result.pendingCount || 0),
    sentCount: Number(result && result.sentCount || 0),
    responseCode: Number(result && result.responseCode || 0),
    retryAfterMs: Number(result && result.retryAfterMs || 0),
    repairedRows: Number(result && result.repairedRows || 0),
    error: SALES_SUPPORT_limitText_(
      result && result.error ? String(result.error) : '',
      SALES_SUPPORT_ALERT_CONFIG.MAX_ERROR_CHARS
    )
  };

  PropertiesService.getScriptProperties().setProperty(
    SALES_SUPPORT_ALERT_CONFIG.LAST_RUN_PROP_KEY,
    JSON.stringify(record)
  );
}


function SALES_SUPPORT_runWithLease_(callback, taskName) {
  if (typeof callback !== 'function') {
    throw new Error('Discord 알림 실행 콜백이 함수가 아닙니다.');
  }

  if (typeof AUTOMATION_acquireModuleLease_ !== 'function') {
    return callback();
  }

  var lease = AUTOMATION_acquireModuleLease_(
    SALES_SUPPORT_ALERT_CONFIG.LEASE_KEY,
    {
      taskName: taskName || 'Discord 영업지원 알림',
      ttlMs: 2 * 60 * 1000,
      waitMs: 0
    }
  );

  if (!lease.acquired) {
    var skipped = {
      status: lease.reason === 'CUTOVER_IN_PROGRESS'
        ? 'SKIPPED_CUTOVER_IN_PROGRESS'
        : 'SKIPPED_ALREADY_RUNNING',
      pendingCount: 0,
      sentCount: 0,
      error: String(lease.reason || 'LEASE_BUSY')
    };

    SALES_SUPPORT_recordLastRun_(skipped);
    return skipped;
  }

  try {
    return callback();
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


function SALES_SUPPORT_limitText_(value, maxLength) {
  var text = String(value == null ? '' : value);
  var limit = Math.max(1, Number(maxLength) || 1);

  if (text.length <= limit) return text;
  if (limit <= 3) return text.substring(0, limit);
  return text.substring(0, limit - 3) + '...';
}


/**
 * 상태 시트를 수동 확인할 때 실행한다.
 */
function showSalesSupportDiscordStateSheet() {
  var sheet = SALES_SUPPORT_getOrCreateStateSheet_();
  sheet.showSheet();
  sheet.getParent().setActiveSheet(sheet);
  return sheet.getName();
}


/**
 * 테스트용. 실제 Discord 응답이 2xx가 아니면 오류를 발생시킨다.
 */
function testSalesSupportDiscordAlert() {
  var result = sendSalesSupportDiscordAlert_([
    {
      id: 'TEST-001',
      oldValue: '',
      newValue: '테스트 알림입니다'
    }
  ], 0);

  if (!result.success) {
    throw new Error(
      'Discord 테스트 실패: HTTP ' + result.responseCode + ' / ' + result.error
    );
  }

  Logger.log('Discord 테스트 성공: HTTP ' + result.responseCode);
  return result;
}
