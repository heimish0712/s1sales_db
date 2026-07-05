/*************************************************
 * IMPORTRANGE 보조시트 → 활성시트 메모 원본확인 가져오기
 *
 * 매핑 기준:
 * - 활성시트 A열
 * - __원본_IMPORT 시트 S열
 *
 * 입력 대상:
 * - 활성시트의 '마스터시트 메모 원본확인' 열
 *
 * 기준:
 * - 헤더 행: 2행
 * - 데이터 시작 행: 3행
 *************************************************/

const IMPORT_MEMO_FROM_BRIDGE_CONFIG = {
  IMPORT_SHEET_NAME: '__원본_IMPORT',

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  // 활성시트 A열
  ACTIVE_KEY_COL: 1,

  // IMPORTRANGE로 가져온 원본의 S열
  SOURCE_KEY_COL: 19,

  TARGET_HEADER_NAME: '마스터시트 메모 원본확인',
  SOURCE_VALUE_HEADER_NAME: '메모',

  KEEP_EXISTING_WHEN_NO_MATCH: true
};


function importMasterMemoOriginalCheckFromImportRangeSheet() {
  const cfg = IMPORT_MEMO_FROM_BRIDGE_CONFIG;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet();
  const importSheet = ss.getSheetByName(cfg.IMPORT_SHEET_NAME);

  if (!importSheet) {
    throw new Error(`IMPORTRANGE 보조시트를 찾을 수 없습니다: ${cfg.IMPORT_SHEET_NAME}`);
  }

  SpreadsheetApp.flush();

  const importA1Value = importSheet.getRange('A1').getDisplayValue();

  if (
    String(importA1Value).includes('#REF') ||
    String(importA1Value).includes('권한') ||
    String(importA1Value).includes('Loading') ||
    String(importA1Value).includes('로드')
  ) {
    throw new Error(
      [
        'IMPORTRANGE 보조시트가 아직 준비되지 않았습니다.',
        '',
        `시트명: ${cfg.IMPORT_SHEET_NAME}`,
        'A1 셀의 IMPORTRANGE 권한 허용 또는 로딩 완료를 먼저 확인해주세요.'
      ].join('\n')
    );
  }

  const activeLastRow = activeSheet.getLastRow();
  const activeLastCol = activeSheet.getLastColumn();

  const sourceLastRow = importSheet.getLastRow();
  const sourceLastCol = importSheet.getLastColumn();

  if (activeLastRow < cfg.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('활성시트에 가져올 데이터 행이 없습니다.');
    return;
  }

  if (sourceLastRow < cfg.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('IMPORTRANGE 보조시트에 원본 데이터 행이 없습니다.');
    return;
  }

  const activeHeaders = activeSheet
    .getRange(cfg.HEADER_ROW, 1, 1, activeLastCol)
    .getDisplayValues()[0]
    .map(normalizeHeader_);

  const targetColIndex = activeHeaders.indexOf(normalizeHeader_(cfg.TARGET_HEADER_NAME)) + 1;

  if (targetColIndex < 1) {
    throw new Error(`활성시트에서 대상 헤더를 찾을 수 없습니다: ${cfg.TARGET_HEADER_NAME}`);
  }

  const sourceHeaders = importSheet
    .getRange(cfg.HEADER_ROW, 1, 1, sourceLastCol)
    .getDisplayValues()[0]
    .map(normalizeHeader_);

  const sourceValueColIndex = sourceHeaders.indexOf(normalizeHeader_(cfg.SOURCE_VALUE_HEADER_NAME)) + 1;

  if (sourceValueColIndex < 1) {
    throw new Error(`IMPORTRANGE 보조시트에서 가져올 헤더를 찾을 수 없습니다: ${cfg.SOURCE_VALUE_HEADER_NAME}`);
  }

  const activeNumRows = activeLastRow - cfg.DATA_START_ROW + 1;
  const sourceNumRows = sourceLastRow - cfg.DATA_START_ROW + 1;

  const activeKeys = activeSheet
    .getRange(cfg.DATA_START_ROW, cfg.ACTIVE_KEY_COL, activeNumRows, 1)
    .getDisplayValues();

  const currentTargetValues = activeSheet
    .getRange(cfg.DATA_START_ROW, targetColIndex, activeNumRows, 1)
    .getDisplayValues();

  const sourceKeys = importSheet
    .getRange(cfg.DATA_START_ROW, cfg.SOURCE_KEY_COL, sourceNumRows, 1)
    .getDisplayValues();

  const sourceValues = importSheet
    .getRange(cfg.DATA_START_ROW, sourceValueColIndex, sourceNumRows, 1)
    .getDisplayValues();

  const sourceMap = new Map();

  for (let i = 0; i < sourceNumRows; i++) {
    const key = normalizeKey_(sourceKeys[i][0]);
    const value = sourceValues[i][0];

    if (!key) continue;

    // 같은 고객번호가 여러 번 있으면 비어있지 않은 값 우선
    if (value !== '') {
      sourceMap.set(key, value);
    } else if (!sourceMap.has(key)) {
      sourceMap.set(key, value);
    }
  }

  const output = [];
  let matchedCount = 0;
  let changedCount = 0;
  let noMatchCount = 0;

  for (let i = 0; i < activeNumRows; i++) {
    const key = normalizeKey_(activeKeys[i][0]);
    const oldValue = currentTargetValues[i][0];

    if (!key) {
      output.push([oldValue]);
      continue;
    }

    if (sourceMap.has(key)) {
      const newValue = sourceMap.get(key);
      output.push([newValue]);
      matchedCount++;

      if (String(oldValue) !== String(newValue)) {
        changedCount++;
      }
    } else {
      noMatchCount++;

      if (cfg.KEEP_EXISTING_WHEN_NO_MATCH) {
        output.push([oldValue]);
      } else {
        output.push(['']);
        if (oldValue !== '') changedCount++;
      }
    }
  }

  activeSheet
    .getRange(cfg.DATA_START_ROW, targetColIndex, activeNumRows, 1)
    .setValues(output);

  SpreadsheetApp.getUi().alert(
    [
      '가져오기 완료',
      '',
      `매칭된 행: ${matchedCount}건`,
      `변경된 행: ${changedCount}건`,
      `매칭 없음: ${noMatchCount}건`,
      '',
      `대상 열: ${cfg.TARGET_HEADER_NAME}`
    ].join('\n')
  );
}


function normalizeHeader_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .trim();
}


function normalizeKey_(value) {
  let text = String(value || '').trim();

  if (!text) return '';

  text = text.replace(/\.0$/, '');
  text = text.replace(/\s+/g, '');

  return text;
}