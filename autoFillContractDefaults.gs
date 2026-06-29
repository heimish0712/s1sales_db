/***************************************
 * 계약단위 입력 시 기본 계약조건 자동 입력 + 일괄 반영
 *
 * 기준:
 * - 헤더 행: 2행
 * - 데이터 시작 행: 3행
 *
 * 헤더명:
 * - 계약단위
 * - 관리자선임여부
 * - 유지점검
 * - 성능점검
 *
 * 실제 입력값 기준:
 * - 계약단위: 숫자 6 / 12
 * - 관리자선임여부: 텍스트 "선임"
 * - 유지점검: 숫자 1 / 2
 * - 성능점검: 숫자 1
 *
 * 자동입력 기준:
 * - 6개월  → 선임 / 1 / 1
 * - 12개월 → 선임 / 2 / 1
 ***************************************/



/***************************************
 * 계약단위 수정 시 같은 행의 기본 계약조건 자동 입력
 ***************************************/
function autoFillContractDefaultsOnUnitEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const UNIT_HEADER = '계약단위';
  const MANAGER_HEADER = '관리자선임여부';
  const MAINTENANCE_HEADER = '유지점검';
  const PERFORMANCE_HEADER = '성능점검';

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

  const unitCol = findHeaderColForContractDefaults_(headers, UNIT_HEADER);
  const managerCol = findHeaderColForContractDefaults_(headers, MANAGER_HEADER);
  const maintenanceCol = findHeaderColForContractDefaults_(headers, MAINTENANCE_HEADER);
  const performanceCol = findHeaderColForContractDefaults_(headers, PERFORMANCE_HEADER);

  if (unitCol < 1 || managerCol < 1 || maintenanceCol < 1 || performanceCol < 1) return;

  // 수정된 범위에 계약단위 컬럼이 포함되지 않으면 무시
  if (unitCol < editedStartCol || unitCol > editedEndCol) return;

  const targetStartRow = Math.max(editedStartRow, DATA_START_ROW);
  const targetEndRow = editedEndRow;

  if (targetEndRow < DATA_START_ROW) return;

  const numRows = targetEndRow - targetStartRow + 1;

  // 실제 값 기준으로 읽음
  // U열이 화면상 "6개월"처럼 보여도 실제 값이 6이면 6으로 읽힘
  const unitValues = sheet
    .getRange(targetStartRow, unitCol, numRows, 1)
    .getValues();

  const managerValues = [];
  const maintenanceValues = [];
  const performanceValues = [];

  unitValues.forEach(row => {
    const contractUnit = row[0];
    const defaults = getContractDefaultsByUnit_(contractUnit);

    managerValues.push([defaults.manager]);           // 텍스트
    maintenanceValues.push([defaults.maintenance]);   // 숫자
    performanceValues.push([defaults.performance]);   // 숫자
  });

  sheet.getRange(targetStartRow, managerCol, numRows, 1).setValues(managerValues);
  sheet.getRange(targetStartRow, maintenanceCol, numRows, 1).setValues(maintenanceValues);
  sheet.getRange(targetStartRow, performanceCol, numRows, 1).setValues(performanceValues);
}


/***************************************
 * 활성 시트 기존 데이터 일괄 반영
 *
 * 실행 함수:
 * fillContractDefaultsByUnitOnActiveSheetOnce
 ***************************************/
function fillContractDefaultsByUnitOnActiveSheetOnce() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const UNIT_HEADER = '계약단위';
  const MANAGER_HEADER = '관리자선임여부';
  const MAINTENANCE_HEADER = '유지점검';
  const PERFORMANCE_HEADER = '성능점검';

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

  const unitCol = findHeaderColForContractDefaults_(headers, UNIT_HEADER);
  const managerCol = findHeaderColForContractDefaults_(headers, MANAGER_HEADER);
  const maintenanceCol = findHeaderColForContractDefaults_(headers, MAINTENANCE_HEADER);
  const performanceCol = findHeaderColForContractDefaults_(headers, PERFORMANCE_HEADER);

  const missingHeaders = [];

  if (unitCol < 1) missingHeaders.push(UNIT_HEADER);
  if (managerCol < 1) missingHeaders.push(MANAGER_HEADER);
  if (maintenanceCol < 1) missingHeaders.push(MAINTENANCE_HEADER);
  if (performanceCol < 1) missingHeaders.push(PERFORMANCE_HEADER);

  if (missingHeaders.length > 0) {
    ui.alert(
      '필수 헤더를 찾지 못했습니다.\n\n' +
      missingHeaders.join('\n')
    );
    return;
  }

  const numRows = lastRow - DATA_START_ROW + 1;

  // 실제 값 기준으로 읽음
  const unitValues = sheet
    .getRange(DATA_START_ROW, unitCol, numRows, 1)
    .getValues();

  const managerValues = [];
  const maintenanceValues = [];
  const performanceValues = [];

  let filledCount = 0;

  unitValues.forEach(row => {
    const contractUnit = row[0];
    const defaults = getContractDefaultsByUnit_(contractUnit);

    if (defaults.manager || defaults.maintenance || defaults.performance) {
      filledCount++;
    }

    managerValues.push([defaults.manager]);           // 텍스트 "선임"
    maintenanceValues.push([defaults.maintenance]);   // 숫자 1 또는 2
    performanceValues.push([defaults.performance]);   // 숫자 1
  });

  sheet.getRange(DATA_START_ROW, managerCol, numRows, 1).setValues(managerValues);
  sheet.getRange(DATA_START_ROW, maintenanceCol, numRows, 1).setValues(maintenanceValues);
  sheet.getRange(DATA_START_ROW, performanceCol, numRows, 1).setValues(performanceValues);

  SpreadsheetApp.flush();

  ui.alert(
    '계약단위 기준 기본조건 일괄 반영 완료\n\n' +
    '시트명: ' + sheet.getName() + '\n' +
    '전체 처리 행 수: ' + numRows + '행\n' +
    '기본조건 반영 행 수: ' + filledCount + '행\n\n' +
    '6개월: 선임 / 유지점검 1 / 성능점검 1\n' +
    '12개월: 선임 / 유지점검 2 / 성능점검 1'
  );
}


/***************************************
 * 계약단위 → 기본 계약조건 반환
 *
 * 반환값:
 * - manager: 텍스트
 * - maintenance: 숫자
 * - performance: 숫자
 ***************************************/
function getContractDefaultsByUnit_(contractUnitValue) {
  const unit = normalizeContractUnitValue_(contractUnitValue);

  if (unit === 6) {
    return {
      manager: '선임',
      maintenance: 1,
      performance: 1
    };
  }

  if (unit === 12) {
    return {
      manager: '선임',
      maintenance: 2,
      performance: 1
    };
  }

  return {
    manager: '',
    maintenance: '',
    performance: ''
  };
}


/***************************************
 * 계약단위 값 정규화
 *
 * 인식 예:
 * - 숫자 6
 * - 숫자 12
 * - "6"
 * - "12"
 * - "6개월"
 * - "12개월"
 * - "1년"
 * - "일년"
 * - "반년"
 ***************************************/
function normalizeContractUnitValue_(value) {
  if (typeof value === 'number') {
    if (value === 6) return 6;
    if (value === 12) return 12;
  }

  const text = String(value || '')
    .replace(/\s+/g, '')
    .trim();

  if (!text) return '';

  if (text === '6' || /6개월|반년/.test(text)) return 6;
  if (text === '12' || /12개월|1년|일년/.test(text)) return 12;

  return '';
}


/***************************************
 * 헤더명 기준 컬럼 찾기
 * - 공백, 줄바꿈 제거 후 비교
 ***************************************/
function findHeaderColForContractDefaults_(headers, headerName) {
  const target = normalizeHeaderForContractDefaults_(headerName);

  for (let i = 0; i < headers.length; i++) {
    const current = normalizeHeaderForContractDefaults_(headers[i]);

    if (current === target) {
      return i + 1;
    }
  }

  return -1;
}


/***************************************
 * 헤더 정규화
 ***************************************/
function normalizeHeaderForContractDefaults_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, '')
    .trim();
}
