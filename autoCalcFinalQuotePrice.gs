/***************************************
 * 최종 견적가 자동 계산
 *
 * 기준:
 * - 헤더 행: 2행
 * - 데이터 시작 행: 3행
 * - 기준표 시트: 계약기준!A2:N10
 *
 * 계산식:
 * - 선임금액 = 관리자 선임 여부가 "선임"이면 선임단가 * 계약개월, 아니면 0
 * - 유지금액 = 유지단가 * 유지점검횟수
 * - 성능금액 = 성능단가 * 성능점검횟수
 * - 부가세가 "포함"이면 1.1, 아니면 1
 * - 최종 견적가 = 만원 단위 절사
 ***************************************/

/***************************************
 * 계약조건 수정 시 같은 행의 최종 견적가 자동 계산
 ***************************************/
function autoCalcFinalQuotePriceOnEdit_(e) {
  try {
    if (!e || !e.range) return;

    const sheet = e.range.getSheet();

    const HEADER_ROW = 2;
    const DATA_START_ROW = 3;

    const editedRange = e.range;
    const editedStartRow = editedRange.getRow();
    const editedEndRow = editedStartRow + editedRange.getNumRows() - 1;
    const editedStartCol = editedRange.getColumn();
    const editedEndCol = editedStartCol + editedRange.getNumColumns() - 1;

    const lastCol = sheet.getLastColumn();
    if (lastCol < 1) return;

    const headers = sheet
      .getRange(HEADER_ROW, 1, 1, lastCol)
      .getDisplayValues()[0];

    const headerMap = buildFinalQuoteHeaderMap_(headers);
    if (!headerMap.ok) return;

    const watchedCols = [
      headerMap.areaCol,
      headerMap.gradeCol,
      headerMap.discountCol,
      headerMap.contractUnitCol,
      headerMap.managerCol,
      headerMap.maintenanceCol,
      headerMap.performanceCol,
      headerMap.vatCol
    ].filter(col => col > 0);

    const shouldRun = watchedCols.some(col => col >= editedStartCol && col <= editedEndCol);
    if (!shouldRun) return;

    const targetStartRow = Math.max(editedStartRow, DATA_START_ROW);
    const targetEndRow = editedEndRow;
    if (targetEndRow < DATA_START_ROW) return;

    const basisMap = getFinalQuoteBasisMap_();
    if (!basisMap || Object.keys(basisMap).length === 0) return;

    recalcFinalQuotePriceRows_(sheet, targetStartRow, targetEndRow, lastCol, headerMap, basisMap);
  } catch (err) {
    console.error('[autoCalcFinalQuotePriceOnEdit_] ' + (err && err.stack ? err.stack : err));
  }
}


/***************************************
 * 활성 시트 기존 데이터 최종 견적가 일괄 계산
 *
 * 실행 함수:
 * fillFinalQuotePriceByContractConditionsOnActiveSheetOnce
 ***************************************/
function fillFinalQuotePriceByContractConditionsOnActiveSheetOnce() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < DATA_START_ROW) {
    ui.alert('처리할 데이터가 없습니다.');
    return;
  }

  if (lastCol < 1) {
    ui.alert('시트에 컬럼이 없습니다.');
    return;
  }

  const headers = sheet
    .getRange(HEADER_ROW, 1, 1, lastCol)
    .getDisplayValues()[0];

  const headerMap = buildFinalQuoteHeaderMap_(headers);
  if (!headerMap.ok) {
    ui.alert(
      '최종 견적가 계산에 필요한 헤더를 찾지 못했습니다.\n\n' +
      headerMap.missingHeaders.join('\n')
    );
    return;
  }

  const basisMap = getFinalQuoteBasisMap_();
  if (!basisMap || Object.keys(basisMap).length === 0) {
    ui.alert('계약기준 시트에서 단가 기준을 읽지 못했습니다.\n\n계약기준!A2:N10 범위를 확인해 주세요.');
    return;
  }

  const result = recalcFinalQuotePriceRows_(sheet, DATA_START_ROW, lastRow, lastCol, headerMap, basisMap);

  SpreadsheetApp.flush();

  ui.alert(
    '최종 견적가 일괄 계산 완료\n\n' +
    '시트명: ' + sheet.getName() + '\n' +
    '처리 행 수: ' + result.totalRows + '행\n' +
    '정상 계산: ' + result.calculatedRows + '행\n' +
    '입력 부족/확인 필요: ' + result.warningRows + '행\n' +
    '전체 공란 처리: ' + result.blankRows + '행'
  );
}


/***************************************
 * 여러 행 최종 견적가 계산 후 X열에 해당하는 최종 견적가 컬럼에 반영
 ***************************************/
function recalcFinalQuotePriceRows_(sheet, startRow, endRow, lastCol, headerMap, basisMap) {
  const numRows = endRow - startRow + 1;

  const rawRows = sheet
    .getRange(startRow, 1, numRows, lastCol)
    .getValues();

  const displayRows = sheet
    .getRange(startRow, 1, numRows, lastCol)
    .getDisplayValues();

  const outputValues = [];

  let calculatedRows = 0;
  let warningRows = 0;
  let blankRows = 0;

  for (let i = 0; i < numRows; i++) {
    const rawRow = rawRows[i];
    const displayRow = displayRows[i];

    const calcResult = calculateFinalQuotePriceForRow_(rawRow, displayRow, headerMap, basisMap);
    outputValues.push([calcResult.value]);

    if (calcResult.status === 'calculated') calculatedRows++;
    if (calcResult.status === 'warning') warningRows++;
    if (calcResult.status === 'blank') blankRows++;
  }

  const targetRange = sheet.getRange(startRow, headerMap.finalQuoteCol, numRows, 1);
  targetRange.setValues(outputValues);
  targetRange.setNumberFormat('₩#,##0');

  return {
    totalRows: numRows,
    calculatedRows: calculatedRows,
    warningRows: warningRows,
    blankRows: blankRows
  };
}


/***************************************
 * 행 1개의 최종 견적가 계산
 ***************************************/
function calculateFinalQuotePriceForRow_(rawRow, displayRow, headerMap, basisMap) {
  const input = {
    gradeRaw: getFinalQuoteCellRaw_(rawRow, headerMap.gradeCol),
    gradeDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.gradeCol),
    discountRaw: getFinalQuoteCellRaw_(rawRow, headerMap.discountCol),
    discountDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.discountCol),
    contractUnitRaw: getFinalQuoteCellRaw_(rawRow, headerMap.contractUnitCol),
    contractUnitDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.contractUnitCol),
    managerRaw: getFinalQuoteCellRaw_(rawRow, headerMap.managerCol),
    managerDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.managerCol),
    maintenanceRaw: getFinalQuoteCellRaw_(rawRow, headerMap.maintenanceCol),
    maintenanceDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.maintenanceCol),
    performanceRaw: getFinalQuoteCellRaw_(rawRow, headerMap.performanceCol),
    performanceDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.performanceCol),
    vatRaw: getFinalQuoteCellRaw_(rawRow, headerMap.vatCol),
    vatDisplay: getFinalQuoteCellDisplay_(displayRow, headerMap.vatCol)
  };

  const requiredFields = [
    [input.gradeRaw, input.gradeDisplay],
    [input.discountRaw, input.discountDisplay],
    [input.contractUnitRaw, input.contractUnitDisplay],
    [input.managerRaw, input.managerDisplay],
    [input.maintenanceRaw, input.maintenanceDisplay],
    [input.performanceRaw, input.performanceDisplay],
    [input.vatRaw, input.vatDisplay]
  ];

  const inputCount = requiredFields.filter(pair => isFinalQuoteInputPresent_(pair[0], pair[1])).length;

  if (inputCount === 0) {
    return { value: '', status: 'blank' };
  }

  if (inputCount < 7) {
    return { value: '변수 입력 부족', status: 'warning' };
  }

  const gradeKey = normalizeFinalQuoteKey_(input.gradeDisplay || input.gradeRaw);
  const basis = basisMap[gradeKey];
  if (!basis) {
    return { value: '#N/A', status: 'warning' };
  }

  const discountRate = parseFinalQuotePercent_(input.discountRaw, input.discountDisplay);
  const contractMonths = parseFinalQuoteNumber_(input.contractUnitRaw, input.contractUnitDisplay);
  const maintenanceCount = parseFinalQuoteNumber_(input.maintenanceRaw, input.maintenanceDisplay);
  const performanceCount = parseFinalQuoteNumber_(input.performanceRaw, input.performanceDisplay);

  if (
    discountRate === null ||
    contractMonths === null ||
    maintenanceCount === null ||
    performanceCount === null ||
    basis.managerUnit === null ||
    basis.maintenanceUnit === null ||
    basis.performanceUnit === null
  ) {
    return { value: '변수 입력 부족', status: 'warning' };
  }

  const discountFactor = 1 - (discountRate / 100);
  const vatFactor = normalizeFinalQuoteKey_(input.vatDisplay || input.vatRaw) === normalizeFinalQuoteKey_('포함') ? 1.1 : 1;

  const managerText = normalizeFinalQuoteKey_(input.managerDisplay || input.managerRaw);
  const managerAmount = managerText === normalizeFinalQuoteKey_('선임')
    ? basis.managerUnit * contractMonths
    : 0;

  const maintenanceAmount = basis.maintenanceUnit * maintenanceCount;
  const performanceAmount = basis.performanceUnit * performanceCount;

  const rawAmount = (managerAmount + maintenanceAmount + performanceAmount) * vatFactor * discountFactor;

  if (!isFinite(rawAmount) || rawAmount < 0) {
    return { value: '변수 입력 부족', status: 'warning' };
  }

  const finalAmount = Math.floor(rawAmount / 10000) * 10000;

  return { value: finalAmount, status: 'calculated' };
}


/***************************************
 * 계약기준 시트에서 등급별 단가 읽기
 *
 * 기준:
 * - A열: 관리등급
 * - D열: 선임단가
 * - E열: 유지단가
 * - F열: 성능단가
 ***************************************/
function getFinalQuoteBasisMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const basisSheet = ss.getSheetByName('계약기준');
  if (!basisSheet) return {};

  const START_ROW = 2;
  const START_COL = 1;
  const NUM_ROWS = 9;
  const NUM_COLS = 14;

  const values = basisSheet
    .getRange(START_ROW, START_COL, NUM_ROWS, NUM_COLS)
    .getDisplayValues();

  const basisMap = {};

  values.forEach(row => {
    const gradeKey = normalizeFinalQuoteKey_(row[0]);
    if (!gradeKey) return;

    basisMap[gradeKey] = {
      managerUnit: parseFinalQuoteNumber_(row[3], row[3]),
      maintenanceUnit: parseFinalQuoteNumber_(row[4], row[4]),
      performanceUnit: parseFinalQuoteNumber_(row[5], row[5])
    };
  });

  return basisMap;
}


/***************************************
 * 헤더맵 생성
 ***************************************/
function buildFinalQuoteHeaderMap_(headers) {
  const headerMap = {
    areaCol: findFinalQuoteHeaderCol_(headers, ['연면적']),
    gradeCol: findFinalQuoteHeaderCol_(headers, ['관리등급']),
    discountCol: findFinalQuoteHeaderCol_(headers, ['할인률(%)', '할인율(%)', '할인률', '할인율']),
    contractUnitCol: findFinalQuoteHeaderCol_(headers, ['계약단위']),
    managerCol: findFinalQuoteHeaderCol_(headers, ['관리자선임여부', '관리자 선임 여부', '관리자\n선임 여부']),
    maintenanceCol: findFinalQuoteHeaderCol_(headers, ['유지점검']),
    performanceCol: findFinalQuoteHeaderCol_(headers, ['성능점검']),
    finalQuoteCol: findFinalQuoteHeaderCol_(headers, ['최종 견적가', '최종견적가', '최종 견적 금액', '최종견적금액']),
    vatCol: findFinalQuoteHeaderCol_(headers, ['부가세'])
  };

  const required = [
    ['관리등급', headerMap.gradeCol],
    ['할인률(%)', headerMap.discountCol],
    ['계약단위', headerMap.contractUnitCol],
    ['관리자 선임 여부', headerMap.managerCol],
    ['유지점검', headerMap.maintenanceCol],
    ['성능점검', headerMap.performanceCol],
    ['최종 견적가', headerMap.finalQuoteCol],
    ['부가세', headerMap.vatCol]
  ];

  const missingHeaders = required
    .filter(item => item[1] < 1)
    .map(item => item[0]);

  headerMap.ok = missingHeaders.length === 0;
  headerMap.missingHeaders = missingHeaders;

  return headerMap;
}


/***************************************
 * 헤더명 기준 컬럼 찾기
 * - 공백, 줄바꿈, 괄호 안 기호 차이를 흡수
 ***************************************/
function findFinalQuoteHeaderCol_(headers, aliases) {
  const targets = aliases.map(alias => normalizeFinalQuoteHeader_(alias));

  for (let i = 0; i < headers.length; i++) {
    const current = normalizeFinalQuoteHeader_(headers[i]);
    if (targets.indexOf(current) >= 0) {
      return i + 1;
    }
  }

  return -1;
}


/***************************************
 * 헤더 정규화
 ***************************************/
function normalizeFinalQuoteHeader_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, '')
    .replace(/％/g, '%')
    .trim();
}


/***************************************
 * 값 비교용 정규화
 ***************************************/
function normalizeFinalQuoteKey_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, '')
    .trim();
}


/***************************************
 * 필수 입력값 존재 여부
 * - 숫자 0은 입력된 값으로 봄
 ***************************************/
function isFinalQuoteInputPresent_(rawValue, displayValue) {
  if (rawValue === 0) return true;
  if (rawValue instanceof Date) return true;

  const rawText = String(rawValue === null || typeof rawValue === 'undefined' ? '' : rawValue).trim();
  const displayText = String(displayValue === null || typeof displayValue === 'undefined' ? '' : displayValue).trim();

  return rawText !== '' || displayText !== '';
}


/***************************************
 * 행 배열에서 1-based 컬럼값 읽기
 ***************************************/
function getFinalQuoteCellRaw_(row, col) {
  if (!row || col < 1) return '';
  return row[col - 1];
}

function getFinalQuoteCellDisplay_(row, col) {
  if (!row || col < 1) return '';
  return row[col - 1];
}


/***************************************
 * 퍼센트 값 파싱
 *
 * 인식 예:
 * - 10
 * - 10%
 * - 14.2
 * - 0.1 + 표시값 10%
 ***************************************/
function parseFinalQuotePercent_(rawValue, displayValue) {
  const displayText = String(displayValue || '').trim();

  if (displayText) {
    const displayNum = parseFinalQuoteNumberFromText_(displayText);
    if (displayNum !== null) return displayNum;
  }

  if (typeof rawValue === 'number' && isFinite(rawValue)) {
    return rawValue;
  }

  return parseFinalQuoteNumberFromText_(rawValue);
}


/***************************************
 * 일반 숫자 파싱
 *
 * 인식 예:
 * - 12개월 → 12
 * - 2회 → 2
 * - ₩4,050,000 → 4050000
 * - 14.2 → 14.2
 ***************************************/
function parseFinalQuoteNumber_(rawValue, displayValue) {
  if (typeof rawValue === 'number' && isFinite(rawValue)) {
    return rawValue;
  }

  const displayNum = parseFinalQuoteNumberFromText_(displayValue);
  if (displayNum !== null) return displayNum;

  return parseFinalQuoteNumberFromText_(rawValue);
}


/***************************************
 * 문자열 안의 숫자 추출
 ***************************************/
function parseFinalQuoteNumberFromText_(value) {
  if (value === null || typeof value === 'undefined') return null;

  let text = String(value).trim();
  if (!text) return null;

  text = text
    .replace(/,/g, '')
    .replace(/₩/g, '')
    .replace(/원/g, '')
    .replace(/개월/g, '')
    .replace(/회/g, '')
    .replace(/%/g, '')
    .replace(/％/g, '')
    .trim();

  const match = text.match(/-?\d+(\.\d+)?/);
  if (!match) return null;

  const num = Number(match[0]);
  if (!isFinite(num)) return null;

  return num;
}
