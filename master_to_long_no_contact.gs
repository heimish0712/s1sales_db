/**
 * 영업관리대장 - 마스터시트(신규) -> TM시트파일 - 장기미접촉 이관
 *
 * 실행 순서
 * 1) previewMasterToLongNoContact_()
 * 2) 생성된 __장기미접촉_이관프리뷰 시트에서 이관하지 않을 행 삭제 또는 A열 체크 해제
 * 3) transferPreviewToLongNoContact_()
 *
 * 중요
 * - 실제 이관은 프리뷰 시트의 현재 남아있는 행만 기준으로 실행합니다.
 * - 수식은 계산된 값만 가져옵니다. copyTo를 쓰지 않으므로 원본 서식은 복사되지 않습니다.
 * - 장기미접촉 빈 행 기준은 A열~L열 값이 모두 공란인 행입니다.
 */

const M2LNC_CFG = {
  // TODO: TM시트파일 URL의 /d/와 /edit 사이 ID로 교체하세요.
  TARGET_TM_SPREADSHEET_ID: '1CewglKPqsOQJqj3hHZBdLRp5_4lEIR_7yZHUDZ8grQQ',

  SOURCE_SHEET_NAME: '마스터시트(신규)',
  TARGET_SHEET_NAME: '장기미접촉',
  PREVIEW_SHEET_NAME: '__장기미접촉_이관프리뷰',
  LOG_SHEET_NAME: '장기미접촉_마스터이관로그',

  SOURCE_HEADER_ROW: 2,
  SOURCE_DATA_START_ROW: 3,
  TARGET_HEADER_ROW: 2,
  TARGET_DATA_START_ROW: 3,

  STATUS_HEADER: '현재 영업 진행 상황',
  STATUS_VALUE: '!!상태지정필요!!',

  SOURCE_TRANSFER_MARK_HEADER: '장기미접촉 이관 여부',
  SOURCE_TRANSFER_MARK_VALUE: 'O',

  TARGET_CUSTOMER_NO_HEADER: '고객번호',

  // 요청 기준: A~L이 전부 비어 있는 행만 빈 행으로 봄
  TARGET_EMPTY_CHECK_START_COL: 1,
  TARGET_EMPTY_CHECK_COL_COUNT: 12,

  // 장기미접촉은 현재 A~Z 안에 필요한 값이 다 있음
  TARGET_WRITE_START_COL: 1,
  TARGET_WRITE_COL_COUNT: 26
};

const M2LNC_PREVIEW_META_HEADERS = [
  '이관여부',
  '실행상태',
  '이관된행',
  '오류',
  '소스행',
  '기존장기미접촉행',
  '기존이관여부'
];

const M2LNC_TARGET_HEADERS = [
  '날짜',
  '지역',
  '견적담당',
  '회사명',
  '담당자',
  '전화번호',
  '직통번호',
  '이메일 주소',
  '연면적',
  '최종 견적가',
  '할인견적률(%)',
  '메모',
  '주소',
  '최종분류',
  '우선순위',
  '마지막컨택월',
  '최근컨택여부',
  '재컨택근거',
  'TM 담당자',
  'TM 컨택 내용',
  '상태값',
  '건물유형',
  '고객번호'
];

// 왼쪽: 마스터시트(신규) 헤더명, 오른쪽: 장기미접촉 헤더명
const M2LNC_SOURCE_TO_TARGET_MAP = {
  '마스터시트 최초등록일': '날짜',
  '지역구분': '지역',
  '영업담당자': '견적담당',
  '회사명': '회사명',
  '고객사 담당자': '담당자',
  '대표전화번호': '전화번호',
  '직통번호': '직통번호',
  '담당자 이메일 주소': '이메일 주소',
  '연면적': '연면적',
  '최종 견적가': '최종 견적가',
  '할인률(%)': '할인견적률(%)',
  '메모': '메모',
  '고객사 상세 주소': '주소',
  'TM 컨택 내용': 'TM 컨택 내용',
  'TM 진행 현황': '상태값',
  '건물 유형': '건물유형',
  '고객번호': '고객번호'
};

// 마스터 쪽에 같은 의미의 헤더가 없어서 이관용 기본값으로 채움
const M2LNC_FIXED_TARGET_VALUES = {
  '최종분류': '상태지정필요-마스터이관',
  '재컨택근거': '마스터 현재 영업 진행 상황=!!상태지정필요!!'
};

function previewMasterToLongNoContact_() {
  assertTargetSpreadsheetIdSet_();

  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(30000)) {
    throw new Error('다른 사용자가 이관 프리뷰를 실행 중입니다. 잠시 후 다시 실행하세요.');
  }

  try {
    const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = getRequiredSheet_(sourceSs, M2LNC_CFG.SOURCE_SHEET_NAME);
    const targetSs = SpreadsheetApp.openById(M2LNC_CFG.TARGET_TM_SPREADSHEET_ID);
    const targetSheet = getRequiredSheet_(targetSs, M2LNC_CFG.TARGET_SHEET_NAME);

    const sourceHeaderMap = getHeaderMap_(sourceSheet, M2LNC_CFG.SOURCE_HEADER_ROW);
    const targetHeaderMap = getHeaderMap_(targetSheet, M2LNC_CFG.TARGET_HEADER_ROW);
    validateRequiredHeadersForPreview_(sourceHeaderMap, targetHeaderMap);

    const targetCustomerRows = buildTargetCustomerRowsByCustomerNo_(targetSheet, targetHeaderMap);
    const sourceValues = getBodyValues_(sourceSheet, M2LNC_CFG.SOURCE_DATA_START_ROW);

    const statusCol = sourceHeaderMap[longNoContactNormalizeHeader_(M2LNC_CFG.STATUS_HEADER)];
    const markCol = sourceHeaderMap[longNoContactNormalizeHeader_(M2LNC_CFG.SOURCE_TRANSFER_MARK_HEADER)] || null;

    const previewRows = [];

    sourceValues.forEach((sourceRow, i) => {
      const sourceRowNo = M2LNC_CFG.SOURCE_DATA_START_ROW + i;
      if (isBlankRow_(sourceRow)) return;

      const status = normalizeCellText_(sourceRow[statusCol - 1]);
      if (status !== M2LNC_CFG.STATUS_VALUE) return;

      const targetRecord = buildTargetRecordFromSourceRow_(sourceRow, sourceHeaderMap);
      const customerKey = longNoContactNormalizeCustomerNo_(targetRecord['고객번호']);
      const existingTargetRows = customerKey && targetCustomerRows[customerKey]
        ? targetCustomerRows[customerKey]
        : [];

      // 이미 장기미접촉에 같은 고객번호가 있으면 기본 체크 해제.
      // 그래도 강제 이관하려면 프리뷰에서 A열을 TRUE로 바꾸면 됨.
      const shouldTransfer = existingTargetRows.length === 0;
      const existingMark = markCol ? sourceRow[markCol - 1] : '';

      previewRows.push(
        [
          shouldTransfer,
          '',
          '',
          '',
          sourceRowNo,
          existingTargetRows.join(', '),
          existingMark
        ].concat(M2LNC_TARGET_HEADERS.map(header => targetRecord[header] === undefined ? '' : targetRecord[header]))
      );
    });

    const previewSheet = getOrCreateSheet_(sourceSs, M2LNC_CFG.PREVIEW_SHEET_NAME);
    previewSheet.clear();

    const previewHeaders = M2LNC_PREVIEW_META_HEADERS.concat(M2LNC_TARGET_HEADERS);
    previewSheet.getRange(1, 1, 1, previewHeaders.length).setValues([previewHeaders]);

    if (previewRows.length > 0) {
      previewSheet.getRange(2, 1, previewRows.length, previewHeaders.length).setValues(previewRows);
      previewSheet.getRange(2, 1, previewRows.length, 1).insertCheckboxes();
    }

    previewSheet.setFrozenRows(1);
    previewSheet.autoResizeColumns(1, Math.min(previewHeaders.length, 12));

    const checkedCount = previewRows.filter(row => row[0] === true).length;
    const duplicateCount = previewRows.length - checkedCount;

    SpreadsheetApp.getUi().alert(
      '장기미접촉 이관 프리뷰 생성 완료\n\n' +
      '상태지정필요 대상: ' + previewRows.length + '건\n' +
      '기본 이관 체크됨: ' + checkedCount + '건\n' +
      '기존 장기미접촉 고객번호 중복으로 체크 해제됨: ' + duplicateCount + '건\n\n' +
      '프리뷰 시트에서 이관하지 않을 행은 삭제하거나 A열 체크를 해제한 뒤 transferPreviewToLongNoContact_()를 실행하세요.'
    );
  } finally {
    lock.releaseLock();
  }
}

function transferPreviewToLongNoContact_() {
  assertTargetSpreadsheetIdSet_();

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('다른 사용자가 장기미접촉 이관을 실행 중입니다. 잠시 후 다시 실행하세요.');
  }

  try {
    const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
    const sourceSheet = getRequiredSheet_(sourceSs, M2LNC_CFG.SOURCE_SHEET_NAME);
    const previewSheet = getRequiredSheet_(sourceSs, M2LNC_CFG.PREVIEW_SHEET_NAME);
    const targetSs = SpreadsheetApp.openById(M2LNC_CFG.TARGET_TM_SPREADSHEET_ID);
    const targetSheet = getRequiredSheet_(targetSs, M2LNC_CFG.TARGET_SHEET_NAME);

    const sourceHeaderMap = getHeaderMap_(sourceSheet, M2LNC_CFG.SOURCE_HEADER_ROW);
    const targetHeaderMap = getHeaderMap_(targetSheet, M2LNC_CFG.TARGET_HEADER_ROW);
    validateRequiredHeadersForTransfer_(sourceHeaderMap, targetHeaderMap);

    const previewRange = previewSheet.getDataRange();
    const previewValues = previewRange.getValues();
    if (previewValues.length < 2) {
      SpreadsheetApp.getUi().alert('프리뷰 데이터가 없습니다. 먼저 previewMasterToLongNoContact_()를 실행하세요.');
      return;
    }

    const previewHeaderMap = getHeaderMapFromHeaderRowValues_(previewValues[0]);
    validatePreviewHeaders_(previewHeaderMap);

    const records = [];
    const executionStatusCol = previewHeaderMap[longNoContactNormalizeHeader_('실행상태')];
    const transferredRowCol = previewHeaderMap[longNoContactNormalizeHeader_('이관된행')];
    const errorCol = previewHeaderMap[longNoContactNormalizeHeader_('오류')];
    const sourceRowCol = previewHeaderMap[longNoContactNormalizeHeader_('소스행')];
    const transferFlagCol = previewHeaderMap[longNoContactNormalizeHeader_('이관여부')];

    for (let i = 1; i < previewValues.length; i++) {
      const previewRow = previewValues[i];
      const previewRowNo = i + 1;
      if (isBlankRow_(previewRow)) continue;

      const oldStatus = normalizeCellText_(previewRow[executionStatusCol - 1]);
      if (oldStatus === '이관완료') continue;

      if (!isTruthy_(previewRow[transferFlagCol - 1])) {
        continue;
      }

      const targetRecord = {};
      M2LNC_TARGET_HEADERS.forEach(header => {
        const col = previewHeaderMap[longNoContactNormalizeHeader_(header)];
        targetRecord[header] = col ? previewRow[col - 1] : '';
      });

      const sourceRowNo = Number(previewRow[sourceRowCol - 1]);
      records.push({
        previewRowNo,
        previewIndex: i,
        sourceRowNo: Number.isFinite(sourceRowNo) ? sourceRowNo : null,
        targetRecord
      });
    }

    if (records.length === 0) {
      SpreadsheetApp.getUi().alert('이관할 프리뷰 행이 없습니다. 삭제/체크 상태를 확인하세요.');
      return;
    }

    const emptyTargetRows = findEmptyRowsByAtoL_(targetSheet, records.length);
    const targetRowsToWrite = records.map(record => buildTargetWriteRow_(record.targetRecord, targetHeaderMap));

    writeRowsToTargetSlots_(targetSheet, emptyTargetRows, targetRowsToWrite);

    const markCol = sourceHeaderMap[longNoContactNormalizeHeader_(M2LNC_CFG.SOURCE_TRANSFER_MARK_HEADER)] || null;
    const now = new Date();
    const userEmail = getActiveUserEmailSafe_();
    const logRows = [];

    records.forEach((record, i) => {
      const targetRowNo = emptyTargetRows[i];
      const previewRow = previewValues[record.previewIndex];

      previewRow[executionStatusCol - 1] = '이관완료';
      previewRow[transferredRowCol - 1] = targetRowNo;
      previewRow[errorCol - 1] = '';

      if (markCol && record.sourceRowNo) {
        sourceSheet.getRange(record.sourceRowNo, markCol).setValue(M2LNC_CFG.SOURCE_TRANSFER_MARK_VALUE);
      }

      logRows.push([
        now,
        userEmail,
        record.sourceRowNo || '',
        targetRowNo,
        record.targetRecord['고객번호'] || '',
        record.targetRecord['회사명'] || '',
        '이관완료',
        '프리뷰 기준 이관 / 값만 붙여넣기 / A~L 빈 행 사용'
      ]);
    });

    // 프리뷰 실행상태/이관된행/오류 갱신
    previewSheet.getRange(2, 1, previewValues.length - 1, previewValues[0].length)
      .setValues(previewValues.slice(1));

    appendTransferLogRows_(sourceSs, logRows);

    SpreadsheetApp.getUi().alert(
      '장기미접촉 이관 완료\n\n' +
      '이관 건수: ' + records.length + '건\n' +
      '대상 시트: ' + M2LNC_CFG.TARGET_SHEET_NAME + '\n' +
      '사용한 빈 행: ' + emptyTargetRows.join(', ')
    );
  } finally {
    lock.releaseLock();
  }
}

function buildTargetRecordFromSourceRow_(sourceRow, sourceHeaderMap) {
  const targetRecord = {};
  M2LNC_TARGET_HEADERS.forEach(header => targetRecord[header] = '');

  Object.keys(M2LNC_SOURCE_TO_TARGET_MAP).forEach(sourceHeader => {
    const targetHeader = M2LNC_SOURCE_TO_TARGET_MAP[sourceHeader];
    const sourceCol = sourceHeaderMap[longNoContactNormalizeHeader_(sourceHeader)];
    if (!sourceCol) return;
    targetRecord[targetHeader] = sourceRow[sourceCol - 1];
  });

  Object.keys(M2LNC_FIXED_TARGET_VALUES).forEach(targetHeader => {
    targetRecord[targetHeader] = M2LNC_FIXED_TARGET_VALUES[targetHeader];
  });

  return targetRecord;
}

function buildTargetWriteRow_(targetRecord, targetHeaderMap) {
  const row = new Array(M2LNC_CFG.TARGET_WRITE_COL_COUNT).fill('');

  M2LNC_TARGET_HEADERS.forEach(targetHeader => {
    const targetCol = targetHeaderMap[longNoContactNormalizeHeader_(targetHeader)];
    if (!targetCol) return;
    const writeIndex = targetCol - M2LNC_CFG.TARGET_WRITE_START_COL;
    if (writeIndex < 0 || writeIndex >= M2LNC_CFG.TARGET_WRITE_COL_COUNT) {
      throw new Error('장기미접촉 헤더 [' + targetHeader + '] 위치가 A~Z 범위를 벗어났습니다. 현재 열: ' + targetCol);
    }
    row[writeIndex] = targetRecord[targetHeader] === undefined ? '' : targetRecord[targetHeader];
  });

  return row;
}

function writeRowsToTargetSlots_(targetSheet, emptyTargetRows, targetRowsToWrite) {
  let groupStartIndex = 0;

  while (groupStartIndex < emptyTargetRows.length) {
    let groupEndIndex = groupStartIndex;
    while (
      groupEndIndex + 1 < emptyTargetRows.length &&
      emptyTargetRows[groupEndIndex + 1] === emptyTargetRows[groupEndIndex] + 1
    ) {
      groupEndIndex++;
    }

    const startRow = emptyTargetRows[groupStartIndex];
    const rowCount = groupEndIndex - groupStartIndex + 1;
    const values = targetRowsToWrite.slice(groupStartIndex, groupEndIndex + 1);

    targetSheet
      .getRange(startRow, M2LNC_CFG.TARGET_WRITE_START_COL, rowCount, M2LNC_CFG.TARGET_WRITE_COL_COUNT)
      .setValues(values);

    groupStartIndex = groupEndIndex + 1;
  }
}

function findEmptyRowsByAtoL_(sheet, requiredCount) {
  const startRow = M2LNC_CFG.TARGET_DATA_START_ROW;
  const startCol = M2LNC_CFG.TARGET_EMPTY_CHECK_START_COL;
  const colCount = M2LNC_CFG.TARGET_EMPTY_CHECK_COL_COUNT;

  let maxRows = Math.max(sheet.getMaxRows(), sheet.getLastRow(), startRow + requiredCount + 50);
  if (sheet.getMaxRows() < maxRows) {
    sheet.insertRowsAfter(sheet.getMaxRows(), maxRows - sheet.getMaxRows());
  }

  const rowCount = maxRows - startRow + 1;
  const values = sheet.getRange(startRow, startCol, rowCount, colCount).getValues();
  const emptyRows = [];

  for (let i = 0; i < values.length && emptyRows.length < requiredCount; i++) {
    if (values[i].every(isBlankCell_)) {
      emptyRows.push(startRow + i);
    }
  }

  if (emptyRows.length < requiredCount) {
    const addCount = requiredCount - emptyRows.length;
    const oldMaxRows = sheet.getMaxRows();
    sheet.insertRowsAfter(oldMaxRows, addCount);
    for (let r = oldMaxRows + 1; r <= oldMaxRows + addCount; r++) {
      emptyRows.push(r);
    }
  }

  return emptyRows;
}

function buildTargetCustomerRowsByCustomerNo_(targetSheet, targetHeaderMap) {
  const customerNoCol = targetHeaderMap[longNoContactNormalizeHeader_(M2LNC_CFG.TARGET_CUSTOMER_NO_HEADER)];
  if (!customerNoCol) {
    throw new Error('장기미접촉 시트에서 고객번호 헤더를 찾지 못했습니다.');
  }

  const lastRow = targetSheet.getLastRow();
  if (lastRow < M2LNC_CFG.TARGET_DATA_START_ROW) return {};

  const rowCount = lastRow - M2LNC_CFG.TARGET_DATA_START_ROW + 1;
  const values = targetSheet.getRange(M2LNC_CFG.TARGET_DATA_START_ROW, customerNoCol, rowCount, 1).getValues();
  const map = {};

  values.forEach((row, i) => {
    const key = longNoContactNormalizeCustomerNo_(row[0]);
    if (!key) return;
    if (!map[key]) map[key] = [];
    map[key].push(M2LNC_CFG.TARGET_DATA_START_ROW + i);
  });

  return map;
}

function appendTransferLogRows_(ss, rows) {
  if (!rows || rows.length === 0) return;

  const sheet = getOrCreateSheet_(ss, M2LNC_CFG.LOG_SHEET_NAME);
  const headers = ['실행일시', '사용자', '마스터소스행', '장기미접촉행', '고객번호', '회사명', '상태', '메시지'];

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }

  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
}

function validateRequiredHeadersForPreview_(sourceHeaderMap, targetHeaderMap) {
  const requiredSourceHeaders = [M2LNC_CFG.STATUS_HEADER].concat(Object.keys(M2LNC_SOURCE_TO_TARGET_MAP));
  assertHeadersExist_(sourceHeaderMap, requiredSourceHeaders, M2LNC_CFG.SOURCE_SHEET_NAME);
  assertHeadersExist_(targetHeaderMap, M2LNC_TARGET_HEADERS, M2LNC_CFG.TARGET_SHEET_NAME);
}

function validateRequiredHeadersForTransfer_(sourceHeaderMap, targetHeaderMap) {
  assertHeadersExist_(targetHeaderMap, M2LNC_TARGET_HEADERS, M2LNC_CFG.TARGET_SHEET_NAME);

  // 이 헤더는 있으면 이관 후 O 표시를 해주고, 없으면 이관 자체는 계속 진행 가능하게 둠
  if (!sourceHeaderMap[longNoContactNormalizeHeader_(M2LNC_CFG.SOURCE_TRANSFER_MARK_HEADER)]) {
    console.warn('마스터시트에 [' + M2LNC_CFG.SOURCE_TRANSFER_MARK_HEADER + '] 헤더가 없어 이관 완료 표시를 생략합니다.');
  }
}

function validatePreviewHeaders_(previewHeaderMap) {
  assertHeadersExist_(previewHeaderMap, M2LNC_PREVIEW_META_HEADERS.concat(M2LNC_TARGET_HEADERS), M2LNC_CFG.PREVIEW_SHEET_NAME);
}

function assertHeadersExist_(headerMap, requiredHeaders, sheetName) {
  const missing = requiredHeaders.filter(header => !headerMap[longNoContactNormalizeHeader_(header)]);
  if (missing.length > 0) {
    throw new Error('[' + sheetName + '] 시트에서 필수 헤더를 찾지 못했습니다: ' + missing.join(', '));
  }
}

function getHeaderMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();
  const headerValues = sheet.getRange(headerRow, 1, 1, lastCol).getValues()[0];
  return getHeaderMapFromHeaderRowValues_(headerValues);
}

function getHeaderMapFromHeaderRowValues_(headerValues) {
  const map = {};
  headerValues.forEach((header, index) => {
    const key = longNoContactNormalizeHeader_(header);
    if (!key) return;
    if (!map[key]) map[key] = index + 1;
  });
  return map;
}

function getBodyValues_(sheet, dataStartRow) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < dataStartRow) return [];
  return sheet.getRange(dataStartRow, 1, lastRow - dataStartRow + 1, lastCol).getValues();
}

function getRequiredSheet_(ss, sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    throw new Error('시트를 찾지 못했습니다: ' + sheetName);
  }
  return sheet;
}

function getOrCreateSheet_(ss, sheetName) {
  return ss.getSheetByName(sheetName) || ss.insertSheet(sheetName);
}

function longNoContactNormalizeHeader_(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function normalizeCellText_(value) {
  return String(value === null || value === undefined ? '' : value)
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}

function longNoContactNormalizeCustomerNo_(value) {
  let text = normalizeCellText_(value);
  if (!text) return '';
  text = text.replace(/,/g, '');
  if (/^\d+(\.0+)?$/.test(text)) {
    text = String(Number(text));
  }
  return text;
}

function isBlankCell_(value) {
  return value === '' || value === null || value === undefined;
}

function isBlankRow_(row) {
  return row.every(isBlankCell_);
}

function isTruthy_(value) {
  if (value === true) return true;
  const text = normalizeCellText_(value).toUpperCase();
  return text === 'TRUE' || text === 'Y' || text === 'YES' || text === 'O' || text === '1';
}

function getActiveUserEmailSafe_() {
  try {
    return Session.getActiveUser().getEmail() || '';
  } catch (e) {
    return '';
  }
}

function assertTargetSpreadsheetIdSet_() {
  if (!M2LNC_CFG.TARGET_TM_SPREADSHEET_ID || M2LNC_CFG.TARGET_TM_SPREADSHEET_ID === '여기에_TM시트파일_ID_입력') {
    throw new Error('M2LNC_CFG.TARGET_TM_SPREADSHEET_ID에 TM시트파일 ID를 먼저 입력하세요.');
  }
}
function previewMasterToLongNoContact() {
  return previewMasterToLongNoContact_();
}

function transferPreviewToLongNoContact() {
  return transferPreviewToLongNoContact_();
}