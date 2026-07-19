/****************************************************
 * ContractSync.gs
 * 마스터 "수주확정/계약완료" ↔ 수행사 "고객관리" 연동
 *
 * 최종 수정 방향:
 * - 열 번호 고정 참조 제거
 * - 마스터/수행사 모두 헤더명 기준으로 동기화
 * - 수행사 파일의 추가 열, 예: "파일 확인" 열 보존
 * - 기존처럼 고객관리 시트를 clearContents()로 밀어버리지 않음
 * - 양방향 연동도 AC:AI / AB:AH 같은 열주소가 아니라 헤더명 기준
 ****************************************************/


/****************************************************
 * 0. 기본 설정
 ****************************************************/
var TARGET_FILES = {
  "KJ": "1uSj0qnAiuelxd1yuDn_7BCB8cHRePaDzJGgih144Boc",
  "일신": "1F_rc7WCrjyMIeKm4N_Kgh004738ZiADTagQG13DuVFw"
};

var MAIN_SHEET_NAME = "수주확정/계약완료";
var TARGET_SHEET_NAME = "고객관리";

var PROP_MASTER_SPREADSHEET_ID = "MASTER_SPREADSHEET_ID";

var PERIODIC_SYNC_HANDLER_NAME = "syncAllFromMasterTimeDriven";
var PERIODIC_SYNC_MINUTES = 5;

var HEADER_SCAN_MAX_ROWS = 5;


/**
 * 수행사 고객관리 시트에 동기화할 표준 헤더 순서
 * 사용자가 현재 활성 시트에 적어준 헤더 기준.
 *
 * 주의:
 * - "파일 확인"은 여기에 넣지 않음.
 * - "파일 확인"은 수행사 파일에서 별도 추가 열로 유지.
 */
var SYNC_HEADERS = [
  "계약번호",
  "고객번호",
  "계약일자(발주번호 부여일)",
  "수행사발송일자",
  "사업자등록증 저장",
  "선임신고서 저장",
  "계약서 저장",
  "지역",
  "계약담당자",
  "고객사명",
  "담당자 성함",
  "전화번호",
  "이메일 주소",
  "연면적",
  "선임유형",
  "계약가",
  "VAT",
  "수행사",
  "사업자등록번호",
  "대표자명",
  "업태",
  "종목",
  "고객사 주소",
  "계약기간",
  "비상주선임",
  "유지점검",
  "성능점검",
  "청구 등 메모",
  "선임예정일",
  "유지점검예정일",
  "성능점검예정일",
  "선임완료여부",
  "유지점검완료여부",
  "성능점검완료여부"
];


/**
 * 수행사 → 마스터 역반영 대상 헤더
 * 기존 AC:AI / AB:AH 대신 헤더명 기준.
 */
var BIDIR_HEADERS = [
  "청구 등 메모",
  "선임예정일",
  "유지점검예정일",
  "성능점검예정일",
  "선임완료여부",
  "유지점검완료여부",
  "성능점검완료여부"
];


var KEY_HEADER_ALIASES = {
  CONTRACT_ID: ["계약번호", "계약 번호", "계약No", "계약NO"],
  CUSTOMER_ID: ["고객번호", "고객 번호", "고객No", "고객NO", "고객 no", "고객 no."],
  ASSIGNEE: ["수행사", "최종수행사", "최종 수행사"]
};


/****************************************************
 * 1. 최초 1회 실행: 트리거 설치
 *
 * 실행 함수:
 * installSyncTriggers()
 ****************************************************/
function installSyncTriggers() {
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var masterId = masterSs.getId();

  PropertiesService
    .getScriptProperties()
    .setProperty(PROP_MASTER_SPREADSHEET_ID, masterId);

  deleteSyncTriggers_();

  ScriptApp
    .newTrigger("installedOnEdit")
    .forSpreadsheet(masterId)
    .onEdit()
    .create();

  for (var assignee in TARGET_FILES) {
    ScriptApp
      .newTrigger("installedOnEdit")
      .forSpreadsheet(TARGET_FILES[assignee])
      .onEdit()
      .create();
  }

  ScriptApp
    .newTrigger(PERIODIC_SYNC_HANDLER_NAME)
    .timeBased()
    .everyMinutes(PERIODIC_SYNC_MINUTES)
    .create();

  Logger.log("트리거 설치 완료: 마스터 + 수행사 파일 onEdit + " + PERIODIC_SYNC_MINUTES + "분 주기 전체 동기화");

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "동기화 트리거 설치 완료: onEdit + " + PERIODIC_SYNC_MINUTES + "분 주기 전체 동기화",
    "설치 완료",
    5
  );
}


/****************************************************
 * 2. 전체 재이관
 *
 * 기존처럼 수행사 시트를 clearContents()로 밀지 않음.
 * 헤더명 기준으로 전체 upsert + 불필요 행 정리.
 * "파일 확인" 같은 추가 열은 보존.
 ****************************************************/
function resetAndReMigrateAllData() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("다른 동기화 작업 실행 중이라 전체 재이관 스킵");
    return;
  }

  try {
    syncAllMainRowsToTargetsIncremental_();

    SpreadsheetApp.getActiveSpreadsheet().toast(
      "전체 재이관 완료: 헤더명 기준 동기화 / 추가 열 보존",
      "완료",
      5
    );
  } finally {
    lock.releaseLock();
  }
}


/****************************************************
 * 3. 시간기반 전체 재동기화
 ****************************************************/
function syncAllFromMasterTimeDriven() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("다른 동기화 작업 실행 중이라 시간기반 동기화 스킵");
    return;
  }

  try {
    syncAllMainRowsToTargetsIncremental_();
  } catch (err) {
    Logger.log("시간기반 전체 동기화 오류: " + getErrorMessage_(err));
    console.error(err);
  } finally {
    lock.releaseLock();
  }
}


/**
 * 마스터 전체를 기준으로 수행사 파일을 증분 동기화한다.
 * - 기존 수행사 파일 행은 upsert
 * - 수행사가 바뀐 경우 다른 수행사 파일에서 제거
 * - 마스터에 더 이상 없는 계약/고객번호는 수행사 파일에서 제거
 * - 수행사 파일의 추가 열은 보존
 */
function syncAllMainRowsToTargetsIncremental_() {
  var mainSheet = getMainSheet_();

  if (!mainSheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다: " + MAIN_SHEET_NAME);
  }

  var mainMeta = requireHeaderMeta_(mainSheet, [
    KEY_HEADER_ALIASES.CONTRACT_ID,
    KEY_HEADER_ALIASES.CUSTOMER_ID,
    KEY_HEADER_ALIASES.ASSIGNEE
  ]);

  var lastRow = mainSheet.getLastRow();

  if (lastRow <= mainMeta.headerRow) {
    Logger.log("마스터에 동기화할 데이터가 없습니다.");
    return;
  }

  SpreadsheetApp.flush();
  Utilities.sleep(500);

  var rowCount = lastRow - mainMeta.headerRow;

  var mainValues = mainSheet
    .getRange(mainMeta.headerRow + 1, 1, rowCount, mainMeta.lastCol)
    .getValues();

  var mainDisplayValues = mainSheet
    .getRange(mainMeta.headerRow + 1, 1, rowCount, mainMeta.lastCol)
    .getDisplayValues();

  var contractCol = findCol_(mainMeta, KEY_HEADER_ALIASES.CONTRACT_ID);
  var customerCol = findCol_(mainMeta, KEY_HEADER_ALIASES.CUSTOMER_ID);
  var assigneeCol = findCol_(mainMeta, KEY_HEADER_ALIASES.ASSIGNEE);

  var uniqueMap = {};

  for (var i = 0; i < mainValues.length; i++) {
    var sourceValues = mainValues[i];
    var sourceDisplay = mainDisplayValues[i];

    var contractId = normalizeKey(sourceDisplay[contractCol - 1]);
    var customerId = normalizeKey(sourceDisplay[customerCol - 1]);

    if (!contractId && !customerId) continue;

    var key = makeUniqueKeyFromIds_(contractId, customerId);
    var assignee = normalizeAssignee_(sourceDisplay[assigneeCol - 1]);
    var record = buildRecordFromSourceRow_(sourceValues, mainMeta);

    uniqueMap[key] = {
      assignee: assignee,
      contractId: contractId,
      customerId: customerId,
      record: record
    };
  }

  var updated = 0;
  var removedBecauseNoAssignee = 0;

  for (var uniqueKey in uniqueMap) {
    var item = uniqueMap[uniqueKey];

    if (!TARGET_FILES[item.assignee]) {
      removeRowFromAllTargetFiles_(item.contractId, item.customerId);
      removedBecauseNoAssignee++;
      continue;
    }

    upsertRowToTargetFile_(
      TARGET_FILES[item.assignee],
      item.contractId,
      item.customerId,
      item.record
    );

    removeRowFromOtherTargetFiles_(
      item.contractId,
      item.customerId,
      item.assignee
    );

    updated++;
  }

  cleanupTargetsNotInMaster_(uniqueMap);

  Logger.log(
    "전체 동기화 완료: 반영 " +
    updated +
    "건, 수행사 없음/미등록 제거 " +
    removedBecauseNoAssignee +
    "건"
  );
}


/****************************************************
 * 4. 설치형 onEdit 트리거 진입점
 ****************************************************/
function installedOnEdit(e) {
  if (!e || !e.range || !e.source) return;

  var route;

  try {
    route = vendorSyncClassifyInstalledEditEvent_(e);
  } catch (classificationError) {
    Logger.log("installedOnEdit 이벤트 판정 오류: " + getErrorMessage_(classificationError));
    console.error(classificationError);
    return;
  }

  // 관련 없는 파일/시트 편집에서는 프로젝트 전체 ScriptLock을 잡지 않음.
  if (!route) return;

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("다른 동기화 작업 실행 중이라 onEdit 동기화 스킵");
    return;
  }

  try {
    if (route.kind === "MAIN") {
      handleMainEdit_(e);
      return;
    }

    if (route.kind === "TARGET") {
      handleTargetEdit_(e);
      return;
    }
  } catch (err) {
    Logger.log("installedOnEdit 오류: " + getErrorMessage_(err));
    console.error(err);
  } finally {
    lock.releaseLock();
  }
}


/**
 * 설치형 onEdit 이벤트가 실제 수행사 동기화 대상인지 락 획득 전에 판정한다.
 */
function vendorSyncClassifyInstalledEditEvent_(e) {
  if (!e || !e.range || !e.source) return null;

  var editedSsId = String(e.source.getId() || "");
  var editedSheetName = String(e.range.getSheet().getName() || "");
  var masterId = String(getMasterSpreadsheetId_() || "");

  if (editedSsId === masterId && editedSheetName === MAIN_SHEET_NAME) {
    return {
      kind: "MAIN",
      spreadsheetId: editedSsId,
      sheetName: editedSheetName
    };
  }

  if (isTargetSpreadsheetId_(editedSsId) && editedSheetName === TARGET_SHEET_NAME) {
    return {
      kind: "TARGET",
      spreadsheetId: editedSsId,
      sheetName: editedSheetName
    };
  }

  return null;
}


/****************************************************
 * 5. 마스터 → 수행사 처리
 ****************************************************/
function handleMainEdit_(e) {
  var sheet = e.range.getSheet();
  var range = e.range;

  var mainMeta = requireHeaderMeta_(sheet, [
    KEY_HEADER_ALIASES.CONTRACT_ID,
    KEY_HEADER_ALIASES.CUSTOMER_ID,
    KEY_HEADER_ALIASES.ASSIGNEE
  ]);

  var startRow = range.getRow();
  var numRows = range.getNumRows();
  var lastEditedRow = startRow + numRows - 1;

  SpreadsheetApp.flush();
  Utilities.sleep(500);

  if (startRow <= mainMeta.headerRow && lastEditedRow >= mainMeta.headerRow) {
    syncHeaderToAllTargets_();

    if (lastEditedRow <= mainMeta.headerRow) {
      return;
    }
  }

  var firstDataRow = Math.max(startRow, mainMeta.headerRow + 1);

  for (var row = firstDataRow; row <= lastEditedRow; row++) {
    syncOneMainRowToTargets_(sheet, row, mainMeta);
  }
}


/**
 * 마스터의 특정 행 하나를 수행사 파일에 반영
 */
function syncOneMainRowToTargets_(mainSheet, row, mainMeta) {
  if (!mainMeta) {
    mainMeta = requireHeaderMeta_(mainSheet, [
      KEY_HEADER_ALIASES.CONTRACT_ID,
      KEY_HEADER_ALIASES.CUSTOMER_ID,
      KEY_HEADER_ALIASES.ASSIGNEE
    ]);
  }

  if (row <= mainMeta.headerRow) return;

  var values = mainSheet
    .getRange(row, 1, 1, mainMeta.lastCol)
    .getValues()[0];

  var displayValues = mainSheet
    .getRange(row, 1, 1, mainMeta.lastCol)
    .getDisplayValues()[0];

  var contractCol = findCol_(mainMeta, KEY_HEADER_ALIASES.CONTRACT_ID);
  var customerCol = findCol_(mainMeta, KEY_HEADER_ALIASES.CUSTOMER_ID);
  var assigneeCol = findCol_(mainMeta, KEY_HEADER_ALIASES.ASSIGNEE);

  var contractId = normalizeKey(displayValues[contractCol - 1]);
  var customerId = normalizeKey(displayValues[customerCol - 1]);

  if (!contractId && !customerId) {
    Logger.log("계약번호/고객번호가 없어 연동하지 않음. row=" + row);
    return;
  }

  var assignee = normalizeAssignee_(displayValues[assigneeCol - 1]);
  var record = buildRecordFromSourceRow_(values, mainMeta);

  if (!TARGET_FILES[assignee]) {
    removeRowFromAllTargetFiles_(contractId, customerId);
    Logger.log("수행사 없음 또는 미등록. 모든 수행사 파일에서 제거: row=" + row);
    return;
  }

  upsertRowToTargetFile_(TARGET_FILES[assignee], contractId, customerId, record);

  removeRowFromOtherTargetFiles_(contractId, customerId, assignee);
}


/**
 * 수행사 파일에 upsert
 * - 헤더명 기준으로 해당 컬럼만 씀
 * - "파일 확인" 같은 추가 열은 보존
 */
function upsertRowToTargetFile_(fileId, contractId, customerId, record) {
  var targetSheet = getTargetSheet_(fileId);
  var targetMeta = ensureTargetHeaders_(targetSheet);

  var matchingRows = findRowsInTargetByKeys_(targetSheet, targetMeta, contractId, customerId);

  var targetRow;

  if (matchingRows.length > 0) {
    targetRow = matchingRows[0];
  } else {
    targetRow = Math.max(targetSheet.getLastRow() + 1, targetMeta.headerRow + 1);
  }

  writeRecordToTargetRow_(targetSheet, targetMeta, targetRow, record);

  matchingRows = findRowsInTargetByKeys_(targetSheet, targetMeta, contractId, customerId);

  for (var i = matchingRows.length - 1; i >= 0; i--) {
    var duplicateRow = matchingRows[i];

    if (duplicateRow !== targetRow) {
      targetSheet.deleteRow(duplicateRow);
    }
  }
}


/****************************************************
 * 6. 수행사 → 마스터 처리
 ****************************************************/
function handleTargetEdit_(e) {
  var targetSheet = e.range.getSheet();
  var range = e.range;

  var targetMeta = ensureTargetHeaders_(targetSheet);
  var bidirCols = getBidirTargetColumns_(targetMeta);

  if (!rangeIntersectsAnyColumns_(range, bidirCols)) {
    return;
  }

  var startRow = range.getRow();
  var startCol = range.getColumn();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  if (startRow <= targetMeta.headerRow) {
    return;
  }

  var values = range.getValues();

  var mainSheet = getMainSheet_();
  var mainMeta = requireHeaderMeta_(mainSheet, [
    KEY_HEADER_ALIASES.CONTRACT_ID,
    KEY_HEADER_ALIASES.CUSTOMER_ID
  ]);

  for (var r = 0; r < numRows; r++) {
    var targetRowNumber = startRow + r;
    if (targetRowNumber <= targetMeta.headerRow) continue;

    var contractId = normalizeKey(
      getDisplayByAliases_(targetSheet, targetRowNumber, targetMeta, KEY_HEADER_ALIASES.CONTRACT_ID)
    );

    var customerId = normalizeKey(
      getDisplayByAliases_(targetSheet, targetRowNumber, targetMeta, KEY_HEADER_ALIASES.CUSTOMER_ID)
    );

    if (!contractId && !customerId) continue;

    var mainRowNumber = findRowInMainByKeys_(mainSheet, mainMeta, contractId, customerId);

    if (!mainRowNumber) {
      Logger.log("마스터에서 대응 행을 찾지 못함: 계약번호=" + contractId + ", 고객번호=" + customerId);
      continue;
    }

    for (var c = 0; c < numCols; c++) {
      var targetColNumber = startCol + c;
      var canonicalHeader = getCanonicalBidirHeaderByTargetCol_(targetMeta, targetColNumber);

      if (!canonicalHeader) continue;

      var mainCol = findColByHeaderName_(mainMeta, canonicalHeader);

      if (mainCol < 1) {
        Logger.log("마스터에서 역반영 대상 헤더를 찾지 못함: " + canonicalHeader);
        continue;
      }

      mainSheet
        .getRange(mainRowNumber, mainCol)
        .setValue(values[r][c]);
    }
  }
}


/****************************************************
 * 7. 레코드 생성 / 쓰기
 ****************************************************/
function buildRecordFromSourceRow_(sourceRow, sourceMeta) {
  var record = {};

  for (var i = 0; i < SYNC_HEADERS.length; i++) {
    var header = SYNC_HEADERS[i];
    var col = findColByHeaderName_(sourceMeta, header);

    record[header] = col > 0 ? sourceRow[col - 1] : "";
  }

  return record;
}


function writeRecordToTargetRow_(targetSheet, targetMeta, targetRow, record) {
  var cells = [];

  for (var i = 0; i < SYNC_HEADERS.length; i++) {
    var header = SYNC_HEADERS[i];
    var col = findColByHeaderName_(targetMeta, header);

    if (col < 1) continue;

    cells.push({
      col: col,
      value: record.hasOwnProperty(header) ? record[header] : ""
    });
  }

  writeCellsByColumnBlocks_(targetSheet, targetRow, cells);
}


/**
 * 인접한 열끼리 묶어서 setValues.
 * 중간에 "파일 확인" 같은 추가 열이 있으면 그 열은 건드리지 않음.
 */
function writeCellsByColumnBlocks_(sheet, row, cells) {
  if (!cells || !cells.length) return;

  cells.sort(function (a, b) {
    return a.col - b.col;
  });

  var blockStartCol = cells[0].col;
  var blockValues = [cells[0].value];
  var prevCol = cells[0].col;

  for (var i = 1; i < cells.length; i++) {
    var item = cells[i];

    if (item.col === prevCol + 1) {
      blockValues.push(item.value);
      prevCol = item.col;
      continue;
    }

    sheet
      .getRange(row, blockStartCol, 1, blockValues.length)
      .setValues([blockValues]);

    blockStartCol = item.col;
    blockValues = [item.value];
    prevCol = item.col;
  }

  sheet
    .getRange(row, blockStartCol, 1, blockValues.length)
    .setValues([blockValues]);
}


/****************************************************
 * 8. 헤더 처리
 ****************************************************/
function syncHeaderToAllTargets_() {
  for (var assignee in TARGET_FILES) {
    var targetSheet = getTargetSheet_(TARGET_FILES[assignee]);
    ensureTargetHeaders_(targetSheet);
  }

  Logger.log("수행사 파일 헤더 확인 완료: 기존 추가 열 보존");
}


/**
 * 수행사 시트에 SYNC_HEADERS가 없으면 추가.
 * 기존 추가 열은 절대 삭제하지 않음.
 */
function ensureTargetHeaders_(targetSheet) {
  var meta = detectHeaderMeta_(targetSheet, [
    KEY_HEADER_ALIASES.CONTRACT_ID,
    KEY_HEADER_ALIASES.CUSTOMER_ID
  ]);

  if (!meta) {
    targetSheet.getRange(1, 1, 1, SYNC_HEADERS.length).setValues([SYNC_HEADERS]);

    return detectHeaderMeta_(targetSheet, [
      KEY_HEADER_ALIASES.CONTRACT_ID,
      KEY_HEADER_ALIASES.CUSTOMER_ID
    ]);
  }

  var appended = false;

  for (var i = 0; i < SYNC_HEADERS.length; i++) {
    var header = SYNC_HEADERS[i];

    if (findColByHeaderName_(meta, header) < 1) {
      var newCol = targetSheet.getLastColumn() + 1;

      targetSheet
        .getRange(meta.headerRow, newCol)
        .setValue(header);

      appended = true;
    }
  }

  if (appended) {
    meta = detectHeaderMeta_(targetSheet, [
      KEY_HEADER_ALIASES.CONTRACT_ID,
      KEY_HEADER_ALIASES.CUSTOMER_ID
    ]);
  }

  return meta;
}


function detectHeaderMeta_(sheet, requiredAliasGroups) {
  var lastRow = sheet.getLastRow();
  var lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) return null;

  var scanRows = Math.min(HEADER_SCAN_MAX_ROWS, lastRow);

  for (var r = 1; r <= scanRows; r++) {
    var headers = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    var map = {};
    var colToHeader = {};

    for (var c = 0; c < headers.length; c++) {
      var original = String(headers[c] || "").trim();
      var normalized = vendorSyncNormalizeHeader_(original);

      if (normalized && !map[normalized]) {
        map[normalized] = c + 1;
      }

      colToHeader[c + 1] = original;
    }

    var ok = true;

    if (requiredAliasGroups && requiredAliasGroups.length) {
      for (var i = 0; i < requiredAliasGroups.length; i++) {
        if (findColFromMap_(map, requiredAliasGroups[i]) < 1) {
          ok = false;
          break;
        }
      }
    }

    if (ok) {
      return {
        headerRow: r,
        lastCol: lastCol,
        headers: headers,
        map: map,
        colToHeader: colToHeader
      };
    }
  }

  return null;
}


function requireHeaderMeta_(sheet, requiredAliasGroups) {
  var meta = detectHeaderMeta_(sheet, requiredAliasGroups);

  if (!meta) {
    throw new Error("헤더를 찾지 못했습니다. 시트명: " + sheet.getName());
  }

  return meta;
}


function findCol_(meta, aliases) {
  if (!meta || !meta.map) return -1;
  return findColFromMap_(meta.map, aliases);
}


function findColFromMap_(map, aliases) {
  var arr = Array.isArray(aliases) ? aliases : [aliases];

  for (var i = 0; i < arr.length; i++) {
    var key = vendorSyncNormalizeHeader_(arr[i]);

    if (map[key]) return map[key];
  }

  return -1;
}


function findColByHeaderName_(meta, headerName) {
  if (!meta || !meta.map) return -1;

  var key = vendorSyncNormalizeHeader_(headerName);

  return meta.map[key] || -1;
}


function getDisplayByAliases_(sheet, row, meta, aliases) {
  var col = findCol_(meta, aliases);
  if (col < 1) return "";

  return sheet.getRange(row, col).getDisplayValue();
}


function getBidirTargetColumns_(targetMeta) {
  var cols = [];

  for (var i = 0; i < BIDIR_HEADERS.length; i++) {
    var col = findColByHeaderName_(targetMeta, BIDIR_HEADERS[i]);

    if (col > 0) {
      cols.push(col);
    }
  }

  return cols;
}


function getCanonicalBidirHeaderByTargetCol_(targetMeta, targetCol) {
  var targetHeader = targetMeta.colToHeader[targetCol];

  if (!targetHeader) return "";

  var normalizedTargetHeader = vendorSyncNormalizeHeader_(targetHeader);

  for (var i = 0; i < BIDIR_HEADERS.length; i++) {
    if (vendorSyncNormalizeHeader_(BIDIR_HEADERS[i]) === normalizedTargetHeader) {
      return BIDIR_HEADERS[i];
    }
  }

  return "";
}


/****************************************************
 * 9. 행 찾기 / 중복 제거
 ****************************************************/
function findRowsInTargetByKeys_(targetSheet, targetMeta, contractId, customerId) {
  var lastRow = targetSheet.getLastRow();

  if (lastRow <= targetMeta.headerRow) return [];

  var contractCol = findCol_(targetMeta, KEY_HEADER_ALIASES.CONTRACT_ID);
  var customerCol = findCol_(targetMeta, KEY_HEADER_ALIASES.CUSTOMER_ID);

  var rowCount = lastRow - targetMeta.headerRow;

  var contractValues = contractCol > 0
    ? targetSheet.getRange(targetMeta.headerRow + 1, contractCol, rowCount, 1).getDisplayValues()
    : [];

  var customerValues = customerCol > 0
    ? targetSheet.getRange(targetMeta.headerRow + 1, customerCol, rowCount, 1).getDisplayValues()
    : [];

  var rows = [];

  for (var i = 0; i < rowCount; i++) {
    var currentContractId = contractCol > 0 ? normalizeKey(contractValues[i][0]) : "";
    var currentCustomerId = customerCol > 0 ? normalizeKey(customerValues[i][0]) : "";

    var matched =
      (contractId && currentContractId === contractId) ||
      (customerId && currentCustomerId === customerId);

    if (matched) {
      rows.push(targetMeta.headerRow + 1 + i);
    }
  }

  return rows;
}


function findRowInMainByKeys_(mainSheet, mainMeta, contractId, customerId) {
  var lastRow = mainSheet.getLastRow();

  if (lastRow <= mainMeta.headerRow) return null;

  var contractCol = findCol_(mainMeta, KEY_HEADER_ALIASES.CONTRACT_ID);
  var customerCol = findCol_(mainMeta, KEY_HEADER_ALIASES.CUSTOMER_ID);

  var rowCount = lastRow - mainMeta.headerRow;

  var contractValues = contractCol > 0
    ? mainSheet.getRange(mainMeta.headerRow + 1, contractCol, rowCount, 1).getDisplayValues()
    : [];

  var customerValues = customerCol > 0
    ? mainSheet.getRange(mainMeta.headerRow + 1, customerCol, rowCount, 1).getDisplayValues()
    : [];

  var foundRow = null;

  for (var i = 0; i < rowCount; i++) {
    var currentContractId = contractCol > 0 ? normalizeKey(contractValues[i][0]) : "";
    var currentCustomerId = customerCol > 0 ? normalizeKey(customerValues[i][0]) : "";

    var matched =
      (contractId && currentContractId === contractId) ||
      (customerId && currentCustomerId === customerId);

    if (matched) {
      foundRow = mainMeta.headerRow + 1 + i;
    }
  }

  return foundRow;
}


function removeRowFromOtherTargetFiles_(contractId, customerId, currentAssignee) {
  for (var assignee in TARGET_FILES) {
    if (assignee === currentAssignee) continue;

    removeRowFromTargetFile_(TARGET_FILES[assignee], contractId, customerId);
  }
}


function removeRowFromAllTargetFiles_(contractId, customerId) {
  for (var assignee in TARGET_FILES) {
    removeRowFromTargetFile_(TARGET_FILES[assignee], contractId, customerId);
  }
}


function removeRowFromTargetFile_(fileId, contractId, customerId) {
  var targetSheet = getTargetSheet_(fileId);
  var targetMeta = ensureTargetHeaders_(targetSheet);

  var rows = findRowsInTargetByKeys_(targetSheet, targetMeta, contractId, customerId);

  for (var i = rows.length - 1; i >= 0; i--) {
    targetSheet.deleteRow(rows[i]);
  }
}


/**
 * 마스터에 없는 계약/고객번호가 수행사 파일에 남아 있으면 제거.
 * 단, 계약번호/고객번호가 둘 다 없는 수기 행은 건드리지 않음.
 */
function cleanupTargetsNotInMaster_(uniqueMap) {
  for (var assignee in TARGET_FILES) {
    var targetSheet = getTargetSheet_(TARGET_FILES[assignee]);
    var targetMeta = ensureTargetHeaders_(targetSheet);

    var lastRow = targetSheet.getLastRow();

    if (lastRow <= targetMeta.headerRow) continue;

    var contractCol = findCol_(targetMeta, KEY_HEADER_ALIASES.CONTRACT_ID);
    var customerCol = findCol_(targetMeta, KEY_HEADER_ALIASES.CUSTOMER_ID);

    var rowCount = lastRow - targetMeta.headerRow;

    var contractValues = contractCol > 0
      ? targetSheet.getRange(targetMeta.headerRow + 1, contractCol, rowCount, 1).getDisplayValues()
      : [];

    var customerValues = customerCol > 0
      ? targetSheet.getRange(targetMeta.headerRow + 1, customerCol, rowCount, 1).getDisplayValues()
      : [];

    for (var i = rowCount - 1; i >= 0; i--) {
      var targetRowNumber = targetMeta.headerRow + 1 + i;

      var contractId = contractCol > 0 ? normalizeKey(contractValues[i][0]) : "";
      var customerId = customerCol > 0 ? normalizeKey(customerValues[i][0]) : "";

      if (!contractId && !customerId) continue;

      var key = makeUniqueKeyFromIds_(contractId, customerId);
      var item = uniqueMap[key];

      if (!item || item.assignee !== assignee) {
        targetSheet.deleteRow(targetRowNumber);
      }
    }
  }
}


/****************************************************
 * 10. 공통 유틸
 ****************************************************/
function rangeIntersectsAnyColumns_(range, cols) {
  if (!cols || !cols.length) return false;

  var rangeStartCol = range.getColumn();
  var rangeEndCol = range.getColumn() + range.getNumColumns() - 1;

  for (var i = 0; i < cols.length; i++) {
    if (rangeStartCol <= cols[i] && cols[i] <= rangeEndCol) {
      return true;
    }
  }

  return false;
}


function makeUniqueKeyFromIds_(contractId, customerId) {
  var normalizedContractId = normalizeKey(contractId);
  var normalizedCustomerId = normalizeKey(customerId);

  if (normalizedContractId) return "CONTRACT:" + normalizedContractId;
  if (normalizedCustomerId) return "CUSTOMER:" + normalizedCustomerId;

  return "";
}


function getMasterSpreadsheetId_() {
  var savedId = PropertiesService
    .getScriptProperties()
    .getProperty(PROP_MASTER_SPREADSHEET_ID);

  if (savedId) return savedId;

  return SpreadsheetApp.getActiveSpreadsheet().getId();
}


function getMainSheet_() {
  var masterId = getMasterSpreadsheetId_();
  var ss = SpreadsheetApp.openById(masterId);

  return ss.getSheetByName(MAIN_SHEET_NAME);
}


function getTargetSheet_(fileId) {
  var ss = SpreadsheetApp.openById(fileId);
  var sheet = ss.getSheetByName(TARGET_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(TARGET_SHEET_NAME);
  }

  return sheet;
}


function isTargetSpreadsheetId_(spreadsheetId) {
  for (var assignee in TARGET_FILES) {
    if (TARGET_FILES[assignee] === spreadsheetId) {
      return true;
    }
  }

  return false;
}


function deleteSyncTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();

  var handlersToDelete = {
    "installedOnEdit": true
  };

  handlersToDelete[PERIODIC_SYNC_HANDLER_NAME] = true;

  for (var i = 0; i < triggers.length; i++) {
    var handler = triggers[i].getHandlerFunction();

    if (handlersToDelete[handler]) {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


/**
 * 기존 함수명 호환용.
 */
function deleteInstalledOnEditTriggers_() {
  deleteSyncTriggers_();
}


function normalizeAssignee_(value) {
  var key = normalizeKey(value).replace(/\s+/g, "");

  if (!key) return "";

  var upperKey = key.toUpperCase();

  if (upperKey === "KJ" || key === "케이제이") {
    return "KJ";
  }

  if (key === "일신" || key === "일신정보통신") {
    return "일신";
  }

  return key;
}


function normalizeKey(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}


function vendorSyncNormalizeHeader_(value) {
  return String(value || "")
    .replace(/\s+/g, "")
    .replace(/[()［］\[\]{}]/g, "")
    .toLowerCase()
    .trim();
}


function getErrorMessage_(err) {
  return err && err.message ? err.message : String(err);
}