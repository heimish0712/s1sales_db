/****************************************************
 * ContractSync.gs
 * 마스터 "수주확정/계약완료" ↔ 수행사 "고객관리" 연동
 *
 * 핵심 구조
 * 1. 마스터 A:AI 중 I열만 제외하여 수행사 파일로 연동
 * 2. 수행사 구분: 마스터 S열
 *    - KJ   → 케이제이 파일의 "고객관리"
 *    - 일신 → 일신 파일의 "고객관리"
 * 3. 고유키: A열 계약번호 우선, 없으면 B열 고객번호 보조
 * 4. 최초 전체 재이관 함수 있음
 * 5. 평소에는 설치형 onEdit 트리거로 자동 반영
 * 6. AC:AI는 양방향 연동
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


/****************************************************
 * 0. 최초 1회 실행: 트리거 설치
 *
 * 이미 installedOnEdit 트리거가 있으면 삭제 후 다시 설치함.
 * 마스터 파일 + 수행사 파일 양쪽에 onEdit 트리거를 설치한다.
 ****************************************************/
function installSyncTriggers() {
  var masterSs = SpreadsheetApp.getActiveSpreadsheet();
  var masterId = masterSs.getId();

  PropertiesService
    .getScriptProperties()
    .setProperty(PROP_MASTER_SPREADSHEET_ID, masterId);

  deleteInstalledOnEditTriggers_();

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

  Logger.log("트리거 설치 완료: 마스터 + 수행사 파일");
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

    var assignee = normalizeKey(sourceRow[ASSIGNEE_COL - 1]);
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
}


/****************************************************
 * 3. 마스터 → 수행사 처리
 ****************************************************/
function handleMainEdit_(e) {
  var sheet = e.range.getSheet();
  var range = e.range;

  var startRow = range.getRow();
  var numRows = range.getNumRows();

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

  var assignee = normalizeKey(sourceRow[ASSIGNEE_COL - 1]);
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

  if (contractId) return "CONTRACT:" + contractId;
  if (customerId) return "CUSTOMER:" + customerId;

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


function deleteInstalledOnEditTriggers_() {
  var triggers = ScriptApp.getProjectTriggers();

  for (var i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === "installedOnEdit") {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
}


function normalizeKey(value) {
  if (value === null || value === undefined) return "";
  return String(value).trim();
}
