/****************************************************
 * AutomationDispatcher.gs
 * 영업관리대장 설치형 이벤트 중앙 디스패처 - 2단계
 *
 * 목적:
 * - 영업관리대장에 개별로 설치되던 onEdit/onChange 트리거를
 *   향후 각각 하나의 중앙 진입점으로 통합한다.
 * - 현재 기존 업무 모듈의 내부 로직은 변경하지 않고,
 *   시트 판정과 호출 순서만 중앙에서 관리한다.
 * - 정식 트리거 재설치는 5분 핵심 동기화 파이프라인까지
 *   준비된 후 TriggerManager에서 일괄 수행한다.
 *
 * 주의:
 * - 이 파일을 추가하는 것만으로는 새 트리거가 설치되지 않는다.
 * - 기존 개별 트리거와 이 디스패처 트리거를 동시에 설치하면
 *   같은 업무가 중복 실행될 수 있다.
 ****************************************************/

var AUTOMATION_DISPATCHER_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE2',

  masterSheetName: '마스터시트(신규)',
  completedSheetName: '수주확정/계약완료',

  fullSyncRequestPropertyKey: 'AUTOMATION_CORE_FULL_SYNC_REQUIRED_V1',

  structuralChangeTypes: Object.freeze({
    INSERT_ROW: true,
    REMOVE_ROW: true,
    INSERT_COLUMN: true,
    REMOVE_COLUMN: true,
    INSERT_GRID: true,
    REMOVE_GRID: true,
    OTHER: true
  })
});


/****************************************************
 * 공개 설치형 트리거 진입점
 ****************************************************/

/**
 * 영업관리대장 통합 설치형 onEdit 진입점.
 *
 * 라우팅:
 * - 마스터시트(신규)
 *   1) sb01 마스터 ↔ 수주확정 동기화
 *   2) 고객사 Drive 폴더 생성/보정
 *
 * - 수주확정/계약완료
 *   1) sb01 수주확정 ↔ 마스터 동기화 및 보조 서식
 *   2) sb02 수행사 고객관리 동기화
 *   3) sb03 정보통신유지보수 동기화
 *
 * - 그 외 시트
 *   아무 작업 없이 즉시 종료
 *
 * 각 모듈 오류는 개별 기록하고 다음 모듈 실행을 계속한다.
 */
function AUTOMATION_handleSalesLedgerEdit(e) {
  var summary = AUTOMATION_createEditSummary_(e);

  if (!summary.hasValidEvent) {
    summary.status = 'IGNORED_INVALID_EVENT';
    return summary;
  }

  if (summary.sheetName === AUTOMATION_DISPATCHER_CONFIG.masterSheetName) {
    summary.route = 'MASTER_SHEET';

    AUTOMATION_runModuleSafely_(
      summary,
      'CONTRACT_MASTER_SYNC',
      'handleContractMasterSyncOnEdit',
      e
    );

    AUTOMATION_runModuleSafely_(
      summary,
      'CUSTOMER_FOLDER',
      'customerFolderInstallableOnEdit',
      e
    );

    summary.status = summary.errorCount > 0
      ? 'COMPLETED_WITH_ERRORS'
      : 'COMPLETED';

    return summary;
  }

  if (summary.sheetName === AUTOMATION_DISPATCHER_CONFIG.completedSheetName) {
    summary.route = 'COMPLETED_SHEET';

    // 순서 중요:
    // sb01이 보조값/마스터 역반영을 먼저 처리한 뒤
    // 수행사 및 정보통신유지보수 파일에 현재 행을 반영한다.
    AUTOMATION_runModuleSafely_(
      summary,
      'CONTRACT_MASTER_SYNC',
      'handleContractMasterSyncOnEdit',
      e
    );

    AUTOMATION_runModuleSafely_(
      summary,
      'VENDOR_SYNC',
      'installedOnEdit',
      e
    );

    AUTOMATION_runModuleSafely_(
      summary,
      'IT_MAINTENANCE_SYNC',
      'ITMAINT_onEditSync_2026',
      e
    );

    summary.status = summary.errorCount > 0
      ? 'COMPLETED_WITH_ERRORS'
      : 'COMPLETED';

    return summary;
  }

  summary.status = 'IGNORED_UNRELATED_SHEET';
  return summary;
}


/**
 * 영업관리대장 통합 설치형 onChange 진입점.
 *
 * 구조 변경 이벤트에서 직접 전체동기화를 실행하지 않는다.
 * 대신 Script Properties에 전체보정 필요 플래그를 기록한다.
 * 후속 3단계의 5분 핵심 동기화 파이프라인이 이 플래그를 소비한다.
 */
function AUTOMATION_handleSalesLedgerChange(e) {
  var changeType = String(e && e.changeType ? e.changeType : '').toUpperCase();

  var result = {
    version: AUTOMATION_DISPATCHER_CONFIG.version,
    changeType: changeType,
    requested: false,
    status: ''
  };

  if (!changeType) {
    result.status = 'IGNORED_INVALID_EVENT';
    return result;
  }

  if (!AUTOMATION_DISPATCHER_CONFIG.structuralChangeTypes[changeType]) {
    result.status = 'IGNORED_NON_STRUCTURAL_CHANGE';
    return result;
  }

  var request = AUTOMATION_recordCoreFullSyncRequest_(changeType, e);

  result.requested = true;
  result.status = 'FULL_SYNC_REQUESTED';
  result.request = request;

  return result;
}


/****************************************************
 * 전체보정 요청 플래그 관리
 ****************************************************/

/**
 * 구조 변경에 따른 전체보정 요청을 누적 기록한다.
 * 실제 전체동기화는 이 함수에서 실행하지 않는다.
 */
function AUTOMATION_recordCoreFullSyncRequest_(changeType, e) {
  var props = PropertiesService.getScriptProperties();
  var key = AUTOMATION_DISPATCHER_CONFIG.fullSyncRequestPropertyKey;
  var now = new Date();
  var nowIso = now.toISOString();
  var previous = AUTOMATION_readJsonProperty_(props, key) || {};

  var changeTypes = Array.isArray(previous.changeTypes)
    ? previous.changeTypes.slice()
    : [];

  if (changeTypes.indexOf(changeType) < 0) {
    changeTypes.push(changeType);
  }

  var sourceSpreadsheetId = '';

  try {
    sourceSpreadsheetId = e && e.source && typeof e.source.getId === 'function'
      ? String(e.source.getId() || '')
      : '';
  } catch (ignoreSourceIdError) {
    sourceSpreadsheetId = '';
  }

  var request = {
    required: true,
    version: AUTOMATION_DISPATCHER_CONFIG.version,
    firstRequestedAt: previous.firstRequestedAt || nowIso,
    lastRequestedAt: nowIso,
    requestCount: Math.max(0, Number(previous.requestCount) || 0) + 1,
    lastChangeType: changeType,
    changeTypes: changeTypes,
    sourceSpreadsheetId: sourceSpreadsheetId
  };

  props.setProperty(key, JSON.stringify(request));
  return request;
}


/**
 * 후속 핵심 동기화 파이프라인에서 사용할 전체보정 요청 조회 함수.
 */
function AUTOMATION_getCoreFullSyncRequest_() {
  return AUTOMATION_readJsonProperty_(
    PropertiesService.getScriptProperties(),
    AUTOMATION_DISPATCHER_CONFIG.fullSyncRequestPropertyKey
  );
}


/**
 * 전체보정 요청 여부만 반환한다.
 */
function AUTOMATION_isCoreFullSyncRequired_() {
  var request = AUTOMATION_getCoreFullSyncRequest_();
  return !!(request && request.required === true);
}


/**
 * 후속 핵심 동기화가 성공한 뒤 호출할 플래그 삭제 함수.
 * 2단계에서는 자동으로 호출하지 않는다.
 */
function AUTOMATION_clearCoreFullSyncRequest_() {
  PropertiesService
    .getScriptProperties()
    .deleteProperty(AUTOMATION_DISPATCHER_CONFIG.fullSyncRequestPropertyKey);
}


/****************************************************
 * 디스패처 내부 보조 함수
 ****************************************************/

function AUTOMATION_createEditSummary_(e) {
  var summary = {
    version: AUTOMATION_DISPATCHER_CONFIG.version,
    hasValidEvent: false,
    status: '',
    route: 'NONE',
    spreadsheetId: '',
    sheetName: '',
    row: 0,
    lastRow: 0,
    column: 0,
    lastColumn: 0,
    moduleCount: 0,
    successCount: 0,
    errorCount: 0,
    modules: []
  };

  if (!e || !e.range || !e.source) {
    return summary;
  }

  try {
    var range = e.range;
    var sheet = range.getSheet();

    summary.hasValidEvent = true;
    summary.spreadsheetId = typeof e.source.getId === 'function'
      ? String(e.source.getId() || '')
      : '';
    summary.sheetName = String(sheet.getName() || '');
    summary.row = Number(range.getRow()) || 0;
    summary.lastRow = typeof range.getLastRow === 'function'
      ? Number(range.getLastRow()) || summary.row
      : summary.row + Math.max(1, Number(range.getNumRows()) || 1) - 1;
    summary.column = Number(range.getColumn()) || 0;
    summary.lastColumn = typeof range.getLastColumn === 'function'
      ? Number(range.getLastColumn()) || summary.column
      : summary.column + Math.max(1, Number(range.getNumColumns()) || 1) - 1;
  } catch (err) {
    summary.hasValidEvent = false;
    summary.status = 'INVALID_EVENT_ACCESS';
    summary.eventError = AUTOMATION_errorMessage_(err);
  }

  return summary;
}


function AUTOMATION_runModuleSafely_(summary, moduleName, handlerName, e) {
  var startedAt = Date.now();
  var row = {
    module: moduleName,
    handler: handlerName,
    status: 'PENDING',
    durationMs: 0,
    error: ''
  };

  summary.moduleCount++;
  summary.modules.push(row);

  try {
    var handler = AUTOMATION_resolveHandler_(handlerName);

    if (typeof handler !== 'function') {
      throw new Error('핸들러를 찾을 수 없습니다: ' + handlerName);
    }

    handler(e);

    row.status = 'SUCCESS';
    summary.successCount++;
  } catch (err) {
    row.status = 'ERROR';
    row.error = AUTOMATION_errorMessage_(err);
    summary.errorCount++;

    console.error(
      '[AUTOMATION_handleSalesLedgerEdit][' + moduleName + '] ' + row.error,
      err
    );
  } finally {
    row.durationMs = Date.now() - startedAt;
  }

  return row;
}


function AUTOMATION_resolveHandler_(handlerName) {
  var safeName = String(handlerName || '');

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(safeName)) {
    return null;
  }

  try {
    if (typeof globalThis !== 'undefined' && typeof globalThis[safeName] === 'function') {
      return globalThis[safeName];
    }
  } catch (ignoreGlobalThisError) {
    // eval 보조 검사로 계속 진행
  }

  try {
    return eval('typeof ' + safeName + ' === "function" ? ' + safeName + ' : null');
  } catch (ignoreEvalError) {
    return null;
  }
}


function AUTOMATION_readJsonProperty_(props, key) {
  var raw = props.getProperty(key);

  if (!raw) return null;

  try {
    var parsed = JSON.parse(raw);
    return parsed && typeof parsed === 'object' ? parsed : null;
  } catch (err) {
    // 손상된 플래그는 다음 요청에서 정상 JSON으로 덮어쓸 수 있도록 제거한다.
    props.deleteProperty(key);
    return null;
  }
}


function AUTOMATION_errorMessage_(err) {
  if (!err) return '알 수 없는 오류';
  return err.message ? String(err.message) : String(err);
}
