/****************************************************
 * 1회성 행 정리 스크립트
 *
 * - A 시트: I열 값이 "KJ"인 행만 남김
 * - B 시트: I열 값이 "일신"인 행만 남김
 * - C 시트: I열 값이 "삼구"인 행만 남김
 *
 * 전제:
 * - 1행은 헤더라서 삭제하지 않음
 * - 실제 데이터는 2행부터
 ****************************************************/

function clean_A_B_C_once() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  cleanSheetByColumnI_(ss, "KJ 선임신고 현황(내부용)", "KJ");
  cleanSheetByColumnI_(ss, "일신 선임신고 현황(내부용)", "일신");
  cleanSheetByColumnI_(ss, "삼구 선임신고 현황(내부용)", "삼구");
}


/**
 * 특정 시트에서 I열 값이 keepValue와 일치하는 행만 남기고 나머지 삭제
 */
function cleanSheetByColumnI_(ss, sheetName, keepValue) {
  const sheet = ss.getSheetByName(sheetName);

  if (!sheet) {
    throw new Error(`시트를 찾을 수 없음: ${sheetName}`);
  }

  const HEADER_ROW = 1;
  const START_ROW = 2;
  const TARGET_COL = 9; // I열

  const lastRow = sheet.getLastRow();

  if (lastRow < START_ROW) {
    Logger.log(`${sheetName}: 삭제할 데이터 없음`);
    return;
  }

  const numRows = lastRow - HEADER_ROW;

  // I열 값만 읽기
  const values = sheet
    .getRange(START_ROW, TARGET_COL, numRows, 1)
    .getValues()
    .map(row => String(row[0]).trim());

  let deleteStart = null;
  let deleteCount = 0;

  // 아래에서 위로 삭제해야 행 밀림 문제 없음
  for (let i = values.length - 1; i >= 0; i--) {
    const rowNumber = START_ROW + i;
    const value = values[i];

    if (value !== keepValue) {
      if (deleteStart === null) {
        deleteStart = rowNumber;
        deleteCount = 1;
      } else if (rowNumber === deleteStart - 1) {
        deleteStart = rowNumber;
        deleteCount++;
      } else {
        sheet.deleteRows(deleteStart, deleteCount);
        deleteStart = rowNumber;
        deleteCount = 1;
      }
    }
  }

  // 마지막 삭제 묶음 처리
  if (deleteStart !== null && deleteCount > 0) {
    sheet.deleteRows(deleteStart, deleteCount);
  }

  Logger.log(`${sheetName}: I열이 "${keepValue}"인 행만 남김`);
}