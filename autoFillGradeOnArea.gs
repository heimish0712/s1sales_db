/***************************************
 * 연면적 입력 시 관리등급 자동 입력 + 일괄 반영
 *
 * 기준:
 * - 헤더 행: 2행
 * - 데이터 시작 행: 3행
 * - 연면적 헤더명: 연면적
 * - 관리등급 헤더명: 관리등급
 *
 * 등급 기준:
 * - 15,000㎡ 미만: 초급
 * - 30,000㎡ 미만: 중급
 * - 60,000㎡ 미만: 고급
 * - 60,000㎡ 이상: 특급
 ***************************************/

/***************************************
 * 연면적 수정 시 같은 행 관리등급 자동 입력
 ***************************************/
function autoFillGradeOnAreaEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const AREA_HEADER = '연면적';
  const GRADE_HEADER = '관리등급';

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

  const areaCol = findHeaderColForGradeFill_(headers, AREA_HEADER);
  const gradeCol = findHeaderColForGradeFill_(headers, GRADE_HEADER);

  // 필요한 헤더가 없으면 작동 안 함
  if (areaCol < 1 || gradeCol < 1) return;

  // 수정된 범위에 연면적 컬럼이 포함되지 않으면 무시
  if (areaCol < editedStartCol || areaCol > editedEndCol) return;

  // 데이터 시작 행 미만은 무시
  const targetStartRow = Math.max(editedStartRow, DATA_START_ROW);
  const targetEndRow = editedEndRow;

  if (targetEndRow < DATA_START_ROW) return;

  const numRows = targetEndRow - targetStartRow + 1;

  const areaValues = sheet
    .getRange(targetStartRow, areaCol, numRows, 1)
    .getDisplayValues();

  const gradeValues = areaValues.map(row => {
    const areaText = String(row[0] || '').trim();
    return [getManagementGradeByArea_(areaText)];
  });

  sheet
    .getRange(targetStartRow, gradeCol, numRows, 1)
    .setValues(gradeValues);
}


/***************************************
 * 활성 시트 기존 데이터 일괄 반영
 *
 * 실행 함수:
 * fillManagementGradeByAreaOnActiveSheetOnce
 ***************************************/
function fillManagementGradeByAreaOnActiveSheetOnce() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const AREA_HEADER = '연면적';
  const GRADE_HEADER = '관리등급';

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

  const areaCol = findHeaderColForGradeFill_(headers, AREA_HEADER);
  const gradeCol = findHeaderColForGradeFill_(headers, GRADE_HEADER);

  if (areaCol < 1) {
    ui.alert('연면적 헤더를 찾지 못했습니다: ' + AREA_HEADER);
    return;
  }

  if (gradeCol < 1) {
    ui.alert('관리등급 헤더를 찾지 못했습니다: ' + GRADE_HEADER);
    return;
  }

  const numRows = lastRow - DATA_START_ROW + 1;

  const areaValues = sheet
    .getRange(DATA_START_ROW, areaCol, numRows, 1)
    .getDisplayValues();

  const gradeValues = areaValues.map(row => {
    const areaText = String(row[0] || '').trim();
    return [getManagementGradeByArea_(areaText)];
  });

  sheet
    .getRange(DATA_START_ROW, gradeCol, numRows, 1)
    .setValues(gradeValues);

  SpreadsheetApp.flush();

  ui.alert(
    '관리등급 일괄 반영 완료\n\n' +
    '시트명: ' + sheet.getName() + '\n' +
    '처리 행 수: ' + numRows + '행\n' +
    '연면적 컬럼: ' + areaCol + '열\n' +
    '관리등급 컬럼: ' + gradeCol + '열'
  );
}


/***************************************
 * 연면적 → 관리등급 반환
 *
 * 기준:
 * - 15,000 미만: 초급
 * - 30,000 미만: 중급
 * - 60,000 미만: 고급
 * - 60,000 이상: 특급
 ***************************************/
function getManagementGradeByArea_(areaValue) {
  const area = parseAreaNumber_(areaValue);

  // 연면적이 비어 있거나 숫자로 읽히지 않으면 빈칸
  // 데이터 확인 규칙 오류 방지용
  if (area === null) return '';

  if (area < 15000) return '초급';
  if (area < 30000) return '중급';
  if (area < 60000) return '고급';

  return '특급';
}


/***************************************
 * 연면적 문자열 → 숫자 변환
 *
 * 예:
 * - 15000
 * - 15,000
 * - 15,000㎡
 * - 연면적 15,000.5㎡
 ***************************************/
function parseAreaNumber_(value) {
  if (value === null || typeof value === 'undefined') return null;

  let text = String(value).trim();
  if (!text) return null;

  text = text
    .replace(/,/g, '')
    .replace(/㎡/g, '')
    .replace(/m²/gi, '')
    .replace(/m2/gi, '')
    .replace(/제곱미터/g, '')
    .trim();

  const match = text.match(/-?\d+(\.\d+)?/);
  if (!match) return null;

  const num = Number(match[0]);

  if (isNaN(num)) return null;
  if (num < 0) return null;

  return num;
}


/***************************************
 * 헤더명 기준 컬럼 찾기
 * - 공백, 줄바꿈 제거 후 비교
 ***************************************/
function findHeaderColForGradeFill_(headers, headerName) {
  const target = normalizeHeaderForGradeFill_(headerName);

  for (let i = 0; i < headers.length; i++) {
    const current = normalizeHeaderForGradeFill_(headers[i]);

    if (current === target) {
      return i + 1;
    }
  }

  return -1;
}


/***************************************
 * 헤더 정규화
 ***************************************/
function normalizeHeaderForGradeFill_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, '')
    .trim();
}
