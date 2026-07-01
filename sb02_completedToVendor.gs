/****************************************************
 * ContractSync.gs
 * 마스터 "수주확정/계약완료" ↔ 수행사 "고객관리" 연동
 *
 * 핵심 구조
 * 1. 마스터 A:AI 중 I열만 제외하여 수행사 파일로 연동
 * 2. 수행사 구분: 마스터 S열
 *    - KJ / 케이제이 → 케이제이 파일의 "고객관리"
 *    - 일신 / 일신정보통신 → 일신 파일의 "고객관리"
 * 3. 고유키: A열 계약번호 우선, 없으면 B열 고객번호 보조
 * 4. 최초 전체 재이관 함수 있음
 * 5. 평소에는 설치형 onEdit 트리거로 자동 반영
 * 6. 다른 스크립트가 S열을 바꾸는 경우를 대비해 5분마다 전체 재동기화
 * 7. AC:AI는 양방향 연동
 *
 * 주의
 * - 마스터 AC:AI = 29~35열
 * - I열을 제외했기 때문에 수행사 파일에서는 AB:AH = 28~34열
 ****************************************************/

var TARGET_FILES = {
  "KJ": "1uSj0qnAiuelxd1yuDn_7BCB8cHRePaDzJGgih144Boc",
  "일신": "1F_rc7WCrjyMIeKm4N_Kgh004738ZiADTagQG13DuVFw"
};

var MAIN_SHEET_NAME = "수주확정/계약완료";
var TARGET_SHEET_NAME = "고객관리";

var SOURCE_FIRST_DATA_ROW = 2;

// 마스터 기준 A:AI
var SOURCE_SYNC_START_COL = 1;
var SOURCE_SYNC_END_COL = 35; // AI

// 제외 열: I열
var EXCLUDED_SOURCE_COLS = {
  9: true
};

// 고유 ID 후보
var CONTRACT_ID_COL = 1; // A열 계약번호
var CUSTOMER_ID_COL = 2; // B열 고객번호

// 수행사 정보
var ASSIGNEE_COL = 19; // S열

// 마스터 기준 양방향 구간 AC:AI
var BIDIR_SOURCE_START_COL = 29; // AC
var BIDIR_SOURCE_END_COL = 35;   // AI

// 설치 시 마스터 스프레드시트 ID 저장용
var PROP_MASTER_SPREADSHEET_ID = "MASTER_SPREADSHEET_ID";

// 다른 스크립트가 S열을 수정하는 경우 onEdit이 안 타므로 시간기반 전체 동기화 추가
var PERIODIC_SYNC_HANDLER_NAME = "syncAllFromMasterTimeDriven";
var PERIODIC_SYNC_MINUTES = 5;


/****************************************************
 * 0. 최초 1회 실행: 트리거 설치
 *
 * 실행 함수:
 * installSyncTriggers()
 *
 * 설치 내용:
 * 1. 마스터 파일 onEdit
 * 2. 수행사 파일 onEdit
 * 3. 5분마다 마스터 전체 재동기화
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
 * 1. 최초 전체 초기화 후 재이관
 *
 * 수행사 파일의 "고객관리" 시트 내용을 전부 비우고,
 * 마스터 기준으로 다시 덮어쓴다.
 *
 * 실행 함수:
 * resetAndReMigrateAllData()
 ****************************************************/
function resetAndReMigrateAllData() {
  var mainSheet = getMainSheet_();

  if (!mainSheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다: " + MAIN_SHEET_NAME);
  }

  var lastRow = mainSheet.getLastRow();

  if (lastRow < 1) {
    Logger.log("마스터 시트에 데이터가 없습니다.");
    return;
  }

  var mainValues = mainSheet
    .getRange(1, 1, lastRow, SOURCE_SYNC_END_COL)
    .getValues();

  var headerValues = buildTargetRowFromSourceRow_(mainValues[0]);

  var grouped = {};

  for (var assigneeName in TARGET_FILES) {
    grouped[assigneeName] = [];
  }

  /*
   * 같은 계약번호/고객번호가 마스터에 여러 번 있으면
   * 아래쪽 행, 즉 나중에 나온 행을 최종본으로 본다.
   */
  var uniqueMap = {};

  for (var i = 1; i < mainValues.length; i++) {
    var sourceRow = mainValues[i];

    var key = getUniqueKeyFromSourceRow_(sourceRow);
    if (!key) continue;

    var assignee = normalizeAssignee_(sourceRow[ASSIGNEE_COL - 1]);
    if (!TARGET_FILES[assignee]) continue;

    uniqueMap[key] = {
      assignee: assignee,
      rowValues: buildTargetRowFromSourceRow_(sourceRow)
    };
  }

  for (var uniqueKey in uniqueMap) {
    var item = uniqueMap[uniqueKey];
    grouped[item.assignee].push(item.rowValues);
  }

  for (var targetAssignee in TARGET_FILES) {
    var fileId = TARGET_FILES[targetAssignee];
    var targetSheet = getTargetSheet_(fileId);

    ensureEnoughColumns_(targetSheet, headerValues.length);

    // 기존 데이터 전부 삭제
    targetSheet.clearContents();

    // 헤더 작성
    targetSheet
      .getRange(1, 1, 1, headerValues.length)
      .setValues([headerValues]);

    var rows = grouped[targetAssignee];

    if (rows.length > 0) {
      targetSheet
        .getRange(2, 1, rows.length, headerValues.length)
        .setValues(rows);
    }

    Logger.log(targetAssignee + " 재이관 완료: " + rows.length + "건");
  }

  Logger.log("전체 초기화 및 재이관 완료");
}


/****************************************************
 * 1-1. 시간기반 전체 재동기화
 *
 * 다른 스크립트가 S열을 채우면 onEdit이 다시 발생하지 않으므로,
 * 이 함수가 5분마다 마스터 전체를 훑어서 수행사 파일에 재반영한다.
 ****************************************************/
function syncAllFromMasterTimeDriven() {
  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("다른 동기화 작업 실행 중이라 이번 시간기반 동기화는 스킵");
    return;
  }

  try {
    syncAllMainRowsToTargetsIncremental_();
  } catch (err) {
    Logger.log("시간기반 전체 동기화 오류: " + err.message);
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
 */
function syncAllMainRowsToTargetsIncremental_() {
  var mainSheet = getMainSheet_();

  if (!mainSheet) {
    throw new Error("마스터 시트를 찾을 수 없습니다: " + MAIN_SHEET_NAME);
  }

  var lastRow = mainSheet.getLastRow();

  if (lastRow < SOURCE_FIRST_DATA_ROW) {
    Logger.log("마스터에 동기화할 데이터가 없습니다.");
    return;
  }

  var rowCount = lastRow - SOURCE_FIRST_DATA_ROW + 1;

  // 수식 계산 지연 대응
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  var mainValues = mainSheet
    .getRange(SOURCE_FIRST_DATA_ROW, 1, rowCount, SOURCE_SYNC_END_COL)
    .getValues();

  /*
   * uniqueMap:
   * key → {
   *   assignee,
   *   contractId,
   *   customerId,
   *   rowValues
   * }
   *
   * 같은 키가 여러 번 나오면 아래쪽 행을 최종본으로 본다.
   */
  var uniqueMap = {};

  for (var i = 0; i < mainValues.length; i++) {
    var sourceRow = mainValues[i];

    var contractId = normalizeKey(sourceRow[CONTRACT_ID_COL - 1]);
    var customerId = normalizeKey(sourceRow[CUSTOMER_ID_COL - 1]);

    if (!contractId && !customerId) continue;

    var key = makeUniqueKeyFromIds_(contractId, customerId);
    var assignee = normalizeAssignee_(sourceRow[ASSIGNEE_COL - 1]);

    uniqueMap[key] = {
      assignee: assignee,
      contractId: contractId,
      customerId: customerId,
      rowValues: buildTargetRowFromSourceRow_(sourceRow)
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
      item.rowValues
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
    "시간기반 전체 동기화 완료: 반영 " +
    updated +
    "건, 수행사 없음/미등록 제거 " +
    removedBecauseNoAssignee +
    "건"
  );
}


/**
 * 마스터에 없는 계약/고객번호가 수행사 파일에 남아 있으면 제거한다.
 * 단, 수행사 파일에 수기로 넣은 별도 행도 지워질 수 있으니
 * 고객관리 시트는 마스터 기준 복제본으로 쓰는 전제다.
 */
function cleanupTargetsNotInMaster_(uniqueMap) {
  for (var assignee in TARGET_FILES) {
    var targetSheet = getTargetSheet_(TARGET_FILES[assignee]);
    var lastRow = targetSheet.getLastRow();

    if (lastRow < 2) continue;

    var idValues = targetSheet
      .getRange(2, 1, lastRow - 1, 2)
      .getDisplayValues();

    for (var i = idValues.length - 1; i >= 0; i--) {
      var targetRowNumber = i + 2;

      var contractId = normalizeKey(idValues[i][0]);
      var customerId = normalizeKey(idValues[i][1]);

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
 * 2. 설치형 onEdit 트리거 진입점
 *
 * 마스터 수정:
 * - 해당 행을 수행사 파일에 반영
 * - 수행사가 바뀌었으면 다른 수행사 파일에서 제거
 *
 * 수행사 수정:
 * - 고객관리 시트의 AB:AH 수정 시
 * - 마스터 AC:AI로 역반영
 ****************************************************/
function installedOnEdit(e) {
  if (!e || !e.range || !e.source) return;

  var lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    Logger.log("다른 동기화 작업 실행 중이라 onEdit 동기화 스킵");
    return;
  }

  try {
    var editedSs = e.source;
    var editedSheet = e.range.getSheet();
    var editedSheetName = editedSheet.getName();

    var masterId = getMasterSpreadsheetId_();
    var editedSsId = editedSs.getId();

    if (editedSsId === masterId && editedSheetName === MAIN_SHEET_NAME) {
      handleMainEdit_(e);
      return;
    }

    if (isTargetSpreadsheetId_(editedSsId) && editedSheetName === TARGET_SHEET_NAME) {
      handleTargetEdit_(e);
      return;
    }
  } catch (err) {
    Logger.log("installedOnEdit 오류: " + err.message);
    console.error(err);
  } finally {
    lock.releaseLock();
  }
}


/****************************************************
 * 3. 마스터 → 수행사 처리
 ****************************************************/
function handleMainEdit_(e) {
  var sheet = e.range.getSheet();
  var range = e.range;

  var startRow = range.getRow();
  var numRows = range.getNumRows();

  // 수식 및 다른 처리 직후의 값 반영 대기
  SpreadsheetApp.flush();
  Utilities.sleep(500);

  // 헤더가 수정된 경우 수행사 헤더도 갱신
  if (startRow === 1) {
    syncHeaderToAllTargets_();

    if (numRows === 1) {
      return;
    }
  }

  var firstDataRow = Math.max(startRow, SOURCE_FIRST_DATA_ROW);
  var lastEditedRow = startRow + numRows - 1;

  for (var row = firstDataRow; row <= lastEditedRow; row++) {
    syncOneMainRowToTargets_(sheet, row);
  }
}


/**
 * 마스터의 특정 행 하나를 수행사 파일에 반영
 */
function syncOneMainRowToTargets_(mainSheet, row) {
  var sourceRow = mainSheet
    .getRange(row, 1, 1, SOURCE_SYNC_END_COL)
    .getValues()[0];

  var contractId = normalizeKey(sourceRow[CONTRACT_ID_COL - 1]);
  var customerId = normalizeKey(sourceRow[CUSTOMER_ID_COL - 1]);

  if (!contractId && !customerId) {
    Logger.log("계약번호/고객번호가 없어 연동하지 않음. row=" + row);
    return;
  }

  var assignee = normalizeAssignee_(sourceRow[ASSIGNEE_COL - 1]);
  var rowValues = buildTargetRowFromSourceRow_(sourceRow);

  if (!TARGET_FILES[assignee]) {
    /*
     * 수행사 값이 비었거나 KJ/일신이 아니면
     * 기존 수행사 파일에 남아 있을 수 있는 행을 제거한다.
     */
    removeRowFromAllTargetFiles_(contractId, customerId);
    Logger.log("수행사 없음 또는 미등록. 모든 수행사 파일에서 제거: row=" + row);
    return;
  }

  var fileId = TARGET_FILES[assignee];

  upsertRowToTargetFile_(fileId, contractId, customerId, rowValues);

  // 수행사가 변경되었을 수 있으므로 다른 수행사 파일에서는 제거
  removeRowFromOtherTargetFiles_(contractId, customerId, assignee);
}


/**
 * 수행사 파일에 upsert
 */
function upsertRowToTargetFile_(fileId, contractId, customerId, rowValues) {
  var targetSheet = getTargetSheet_(fileId);
  var headerValues = getTargetHeaderValues_();

  ensureEnoughColumns_(targetSheet, headerValues.length);
  ensureHeader_(targetSheet, headerValues);

  var matchingRows = findRowsInTargetByKeys_(targetSheet, contractId, customerId);

  var targetRow;

  if (matchingRows.length > 0) {
    targetRow = matchingRows[0];
  } else {
    targetRow = Math.max(targetSheet.getLastRow() + 1, 2);
  }

  targetSheet
    .getRange(targetRow, 1, 1, rowValues.length)
    .setValues([rowValues]);

  // 중복 제거
  matchingRows = findRowsInTargetByKeys_(targetSheet, contractId, customerId);

  for (var i = matchingRows.length - 1; i >= 0; i--) {
    var duplicateRow = matchingRows[i];

    if (duplicateRow !== targetRow) {
      targetSheet.deleteRow(duplicateRow);
    }
  }
}


/****************************************************
 * 4. 수행사 → 마스터 처리
 *
 * 마스터 AC:AI는 수행사 파일에서 AB:AH에 해당한다.
 ****************************************************/
function handleTargetEdit_(e) {
  var targetSheet = e.range.getSheet();
  var range = e.range;

  var targetBidirStartCol = sourceColToTargetCol_(BIDIR_SOURCE_START_COL); // AC → AB
  var targetBidirEndCol = sourceColToTargetCol_(BIDIR_SOURCE_END_COL);     // AI → AH

  if (!rangeIntersectsColumns_(range, targetBidirStartCol, targetBidirEndCol)) {
    return;
  }

  var startRow = range.getRow();
  var startCol = range.getColumn();
  var numRows = range.getNumRows();
  var numCols = range.getNumColumns();

  if (startRow === 1) {
    return;
  }

  var values = range.getValues();
  var mainSheet = getMainSheet_();

  for (var r = 0; r < numRows; r++) {
    var targetRowNumber = startRow + r;
    if (targetRowNumber === 1) continue;

    var ids = targetSheet
      .getRange(targetRowNumber, 1, 1, 2)
      .getDisplayValues()[0];

    var contractId = normalizeKey(ids[0]);
    var customerId = normalizeKey(ids[1]);

    if (!contractId && !customerId) continue;

    var mainRowNumber = findRowInMainByKeys_(mainSheet, contractId, customerId);
    if (!mainRowNumber) {
      Logger.log("마스터에서 대응 행을 찾지 못함: 계약번호=" + contractId + ", 고객번호=" + customerId);
      continue;
    }

    for (var c = 0; c < numCols; c++) {
      var targetColNumber = startCol + c;

      if (targetColNumber < targetBidirStartCol || targetColNumber > targetBidirEndCol) {
        continue;
      }

      var sourceColNumber = targetColToSourceCol_(targetColNumber);

      if (sourceColNumber < BIDIR_SOURCE_START_COL || sourceColNumber > BIDIR_SOURCE_END_COL) {
        continue;
      }

      mainSheet
        .getRange(mainRowNumber, sourceColNumber)
        .setValue(values[r][c]);
    }
  }
}


/****************************************************
 * 5. 행/헤더 변환
 ****************************************************/

/**
 * 마스터 A:AI 중 I열만 제외하여 수행사용 행 생성
 */
function buildTargetRowFromSourceRow_(sourceRow) {
  var rowValues = [];

  for (var sourceCol = SOURCE_SYNC_START_COL; sourceCol <= SOURCE_SYNC_END_COL; sourceCol++) {
    if (EXCLUDED_SOURCE_COLS[sourceCol]) continue;

    rowValues.push(sourceRow[sourceCol - 1]);
  }

  return rowValues;
}


/**
 * 마스터 헤더 기준 수행사 헤더 생성
 */
function getTargetHeaderValues_() {
  var mainSheet = getMainSheet_();

  return buildTargetRowFromSourceRow_(
    mainSheet.getRange(1, 1, 1, SOURCE_SYNC_END_COL).getValues()[0]
  );
}


/**
 * 수행사 헤더 갱신
 */
function syncHeaderToAllTargets_() {
  var headerValues = getTargetHeaderValues_();

  for (var assignee in TARGET_FILES) {
    var targetSheet = getTargetSheet_(TARGET_FILES[assignee]);

    ensureEnoughColumns_(targetSheet, headerValues.length);

    targetSheet
      .getRange(1, 1, 1, headerValues.length)
      .setValues([headerValues]);
  }

  Logger.log("수행사 파일 헤더 갱신 완료");
}


function ensureHeader_(targetSheet, headerValues) {
  if (targetSheet.getLastRow() === 0) {
    targetSheet
      .getRange(1, 1, 1, headerValues.length)
      .setValues([headerValues]);
    return;
  }

  var firstCell = targetSheet.getRange(1, 1).getValue();

  if (!firstCell) {
    targetSheet
      .getRange(1, 1, 1, headerValues.length)
      .setValues([headerValues]);
  }
}


/****************************************************
 * 6. 행 찾기 / 중복 제거
 ****************************************************/

function findRowsInTargetByKeys_(targetSheet, contractId, customerId) {
  var lastRow = targetSheet.getLastRow();

  if (lastRow < 2) return [];

  var idValues = targetSheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues();

  var rows = [];

  for (var i = 0; i < idValues.length; i++) {
    var currentContractId = normalizeKey(idValues[i][0]);
    var currentCustomerId = normalizeKey(idValues[i][1]);

    var matched =
      (contractId && currentContractId === contractId) ||
      (customerId && currentCustomerId === customerId);

    if (matched) {
      rows.push(i + 2);
    }
  }

  return rows;
}


function findRowInMainByKeys_(mainSheet, contractId, customerId) {
  var lastRow = mainSheet.getLastRow();

  if (lastRow < 2) return null;

  var idValues = mainSheet
    .getRange(2, 1, lastRow - 1, 2)
    .getDisplayValues();

  var foundRow = null;

  for (var i = 0; i < idValues.length; i++) {
    var currentContractId = normalizeKey(idValues[i][0]);
    var currentCustomerId = normalizeKey(idValues[i][1]);

    var matched =
      (contractId && currentContractId === contractId) ||
      (customerId && currentCustomerId === customerId);

    if (matched) {
      // 중복이 있으면 아래쪽 행을 최종본으로 본다.
      foundRow = i + 2;
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
  var rows = findRowsInTargetByKeys_(targetSheet, contractId, customerId);

  for (var i = rows.length - 1; i >= 0; i--) {
    targetSheet.deleteRow(rows[i]);
  }
}


/****************************************************
 * 7. 열 매핑
 ****************************************************/

/**
 * 마스터 열 번호 → 수행사 열 번호
 * I열 제외 때문에 I 이후 열은 -1
 */
function sourceColToTargetCol_(sourceCol) {
  var targetCol = 0;

  for (var col = SOURCE_SYNC_START_COL; col <= sourceCol; col++) {
    if (EXCLUDED_SOURCE_COLS[col]) continue;
    targetCol++;
  }

  return targetCol;
}


/**
 * 수행사 열 번호 → 마스터 열 번호
 */
function targetColToSourceCol_(targetCol) {
  var currentTargetCol = 0;

  for (var sourceCol = SOURCE_SYNC_START_COL; sourceCol <= SOURCE_SYNC_END_COL; sourceCol++) {
    if (EXCLUDED_SOURCE_COLS[sourceCol]) continue;

    currentTargetCol++;

    if (currentTargetCol === targetCol) {
      return sourceCol;
    }
  }

  return null;
}


function rangeIntersectsColumns_(range, startCol, endCol) {
  var rangeStartCol = range.getColumn();
  var rangeEndCol = range.getColumn() + range.getNumColumns() - 1;

  return rangeStartCol <= endCol && rangeEndCol >= startCol;
}


/****************************************************
 * 8. 유틸
 ****************************************************/

function getUniqueKeyFromSourceRow_(sourceRow) {
  var contractId = normalizeKey(sourceRow[CONTRACT_ID_COL - 1]);
  var customerId = normalizeKey(sourceRow[CUSTOMER_ID_COL - 1]);

  return makeUniqueKeyFromIds_(contractId, customerId);
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


function ensureEnoughColumns_(sheet, neededCols) {
  var currentMaxCols = sheet.getMaxColumns();

  if (currentMaxCols < neededCols) {
    sheet.insertColumnsAfter(currentMaxCols, neededCols - currentMaxCols);
  }
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
 * 예전 코드에서 이 함수를 직접 호출하던 경우를 대비해 남겨둠.
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