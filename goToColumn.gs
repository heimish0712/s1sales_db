function goToColumnI_viaIPlus6_sameRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();

  if (!cell) return;

  const row = cell.getRow();

  // I열 + 6 = O열
  sheet.getRange(row, 15).activate();
  SpreadsheetApp.flush();

  // 바로 I열로 복귀
  sheet.getRange(row, 9).activate();
}
function goToColumnN_viaNPlus6_sameRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (!cell) return;

  const row = cell.getRow();

  // N열 + 6 = T열
  sheet.getRange(row, 24).activate();
  SpreadsheetApp.flush();

  // N열
  sheet.getRange(row, 14).activate();
}

function goToColumnY_viaYPlus6_sameRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (!cell) return;

  const row = cell.getRow();

  // Y열 + 6 = AE열
  sheet.getRange(row, 35).activate();
  SpreadsheetApp.flush();

  // Y열
  sheet.getRange(row, 25).activate();
}

function goToColumnAL_viaALPlus6_sameRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const cell = sheet.getActiveCell();
  if (!cell) return;

  const row = cell.getRow();

  // AL열 + 6 = AR열
  sheet.getRange(row, 48).activate();
  SpreadsheetApp.flush();

  // AL열
  sheet.getRange(row, 38).activate();
}
