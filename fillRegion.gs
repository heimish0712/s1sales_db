/***************************************
 * 주소 입력 시 지역구분 자동 입력
 *
 * 기준:
 * - 헤더 행: 2행
 * - 주소 헤더명: 고객사 상세 주소
 * - 지역구분 헤더명: 지역구분
 * - 데이터 시작 행: 3행
 *
 * 동작:
 * - 사용자가 주소 칼럼을 수정하면
 * - 같은 행의 지역구분 칼럼을 자동 입력
 * - 열 순서가 바뀌어도 헤더명으로 찾아서 작동
 ***************************************/

/***************************************
 * 활성 시트 주소 기준 지역구분 일괄 입력
 *
 * 실행 함수:
 * fillRegionByAddressOnActiveSheetOnce
 *
 * 기준:
 * - 활성 시트만 대상
 * - 헤더 행: 2행
 * - 데이터 시작 행: 3행
 * - 주소 헤더명: 고객사 상세 주소
 * - 지역구분 헤더명: 지역구분
 *
 * 동작:
 * - 현재 활성 시트의 고객사 상세 주소를 전부 읽음
 * - 같은 행의 지역구분 컬럼에 권역 일괄 입력
 * - 기존 getRegionByAddress_ 함수를 그대로 사용
 ***************************************/
function fillRegionByAddressOnActiveSheetOnce() {
  const ui = SpreadsheetApp.getUi();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const ADDRESS_HEADER = '고객사 상세 주소';
  const REGION_HEADER = '지역구분';

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

  const addressCol = findHeaderColByName_(headers, ADDRESS_HEADER);
  const regionCol = findHeaderColByName_(headers, REGION_HEADER);

  if (addressCol < 1) {
    ui.alert('주소 헤더를 찾지 못했습니다: ' + ADDRESS_HEADER);
    return;
  }

  if (regionCol < 1) {
    ui.alert('지역구분 헤더를 찾지 못했습니다: ' + REGION_HEADER);
    return;
  }

  const numRows = lastRow - DATA_START_ROW + 1;

  const addressValues = sheet
    .getRange(DATA_START_ROW, addressCol, numRows, 1)
    .getDisplayValues();

  const regionValues = addressValues.map(row => {
    const address = String(row[0] || '').trim();

    // 기존 로직 그대로 사용
    return [getRegionByAddress_(address)];
  });

  sheet
    .getRange(DATA_START_ROW, regionCol, numRows, 1)
    .setValues(regionValues);

  SpreadsheetApp.flush();

  ui.alert(
    '지역구분 일괄 입력 완료\n\n' +
    '시트명: ' + sheet.getName() + '\n' +
    '처리 행 수: ' + numRows + '행\n' +
    '주소 컬럼: ' + addressCol + '열\n' +
    '지역구분 컬럼: ' + regionCol + '열'
  );
}


/***************************************
 * 헤더명 기준 컬럼 찾기
 * - 공백, 줄바꿈 제거 후 비교
 ***************************************/
function findHeaderColByName_(headers, headerName) {
  const target = normalizeHeaderForRegionFill_(headerName);

  for (let i = 0; i < headers.length; i++) {
    const current = normalizeHeaderForRegionFill_(headers[i]);

    if (current === target) {
      return i + 1;
    }
  }

  return -1;
}


/***************************************
 * 헤더 정규화
 ***************************************/
function normalizeHeaderForRegionFill_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/\s+/g, '')
    .trim();
}

/***************************************
 * onEdit 본체
 ***************************************/
function autoFillRegionOnAddressEdit_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  const HEADER_ROW = 2;
  const DATA_START_ROW = 3;

  const ADDRESS_HEADER = '고객사 상세 주소';
  const REGION_HEADER = '지역구분';

  const editedRange = e.range;
  const editedStartRow = editedRange.getRow();
  const editedEndRow = editedStartRow + editedRange.getNumRows() - 1;
  const editedStartCol = editedRange.getColumn();
  const editedEndCol = editedStartCol + editedRange.getNumColumns() - 1;

  // 헤더 행 읽기
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;

  const headers = sheet
    .getRange(HEADER_ROW, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(v => String(v || '').trim());

  const addressCol = headers.indexOf(ADDRESS_HEADER) + 1;
  const regionCol = headers.indexOf(REGION_HEADER) + 1;

  // 필요한 헤더가 없으면 작동 안 함
  if (addressCol < 1 || regionCol < 1) return;

  // 수정된 범위에 주소 칼럼이 포함되지 않으면 무시
  if (addressCol < editedStartCol || addressCol > editedEndCol) return;

  // 3행 미만은 무시
  const targetStartRow = Math.max(editedStartRow, DATA_START_ROW);
  const targetEndRow = editedEndRow;

  if (targetEndRow < DATA_START_ROW) return;

  const numRows = targetEndRow - targetStartRow + 1;

  // 주소값 읽기
  const addressValues = sheet
    .getRange(targetStartRow, addressCol, numRows, 1)
    .getDisplayValues();

  // 지역구분 산정
  const regionValues = addressValues.map(row => {
    const address = String(row[0] || '').trim();
    return [getRegionByAddress_(address)];
  });

  // 지역구분 입력
  sheet
    .getRange(targetStartRow, regionCol, numRows, 1)
    .setValues(regionValues);
}


/***************************************
 * 주소 문자열 → 권역 반환
 ***************************************/
function getRegionByAddress_(address) {
  if (!address) return '주소확인필요';

  const text = normalizeAddressText_(address);

  // 주소처럼 보이는 키워드가 없으면 주소확인필요
  if (!hasAddressLikeKeyword_(text)) {
    return '주소확인필요';
  }

  // 메모성 값이면 주소확인필요
  // 단, 실제 시도/광역시명이 있으면 주소로 인정
  if (isMemoOnlyText_(text)) {
    return '주소확인필요';
  }

  // 수도권: 서울, 경기, 인천
  if (/(서울특별시|서울시|서울|경기도|경기|인천광역시|인천시|인천)/.test(text)) {
    return '수도권';
  }

  // 강원권: 강원
  if (/(강원특별자치도|강원도|강원)/.test(text)) {
    return '강원권';
  }

  // 충청권: 대전, 세종, 충북, 충남
  if (/(대전광역시|대전시|대전|세종특별자치시|세종시|세종|충청북도|충북|충청남도|충남)/.test(text)) {
    return '충청권';
  }

  // 대구경북권: 대구, 경북
  // '해운대구' 안의 '대구' 같은 부분 문자열은 대구로 오인하지 않는다.
  const hasDaeguToken = /(^|\s)(대구광역시|대구시|대구)(?=\s|$)/.test(text);
  if (hasDaeguToken || /(경상북도|경북)/.test(text)) {
    return '대구경북권';
  }

  // 부울경권: 부산, 울산, 경남
  if (/(부산광역시|부산시|부산|울산광역시|울산시|울산|경상남도|경남)/.test(text)) {
    return '부울경권';
  }

  // 호남권: 광주, 전북, 전남
  if (/(광주광역시|광주시|광주|전라북도|전북|전라남도|전남)/.test(text)) {
    return '호남권';
  }

  // 제주권: 제주
  if (/(제주특별자치도|제주도|제주시|서귀포시|제주)/.test(text)) {
    return '제주권';
  }

  return '주소확인필요';
}


/***************************************
 * 주소 텍스트 정리
 ***************************************/
function normalizeAddressText_(value) {
  return String(value || '')
    .trim()
    .replace(/\s+/g, ' ')
    .replace(/[()]/g, ' ')
    .replace(/,/g, ' ');
}


/***************************************
 * 주소처럼 보이는 핵심 키워드 존재 여부
 ***************************************/
function hasAddressLikeKeyword_(text) {
  if (!text) return false;

  return /(서울특별시|서울시|서울|경기도|경기|인천광역시|인천시|인천|강원특별자치도|강원도|강원|대전광역시|대전시|대전|세종특별자치시|세종시|세종|충청북도|충북|충청남도|충남|대구광역시|대구시|대구|경상북도|경북|부산광역시|부산시|부산|울산광역시|울산시|울산|경상남도|경남|광주광역시|광주시|광주|전라북도|전북|전라남도|전남|제주특별자치도|제주도|제주시|서귀포시|제주|시 |군 |구 |읍 |면 |동 |리 |로 |길|번길)/.test(text);
}


/***************************************
 * 주소가 아니라 순수 메모성 값인지 판단
 ***************************************/
function isMemoOnlyText_(text) {
  if (!text) return true;

  // 실제 광역/도 단위 지역명이 있으면 주소 후보로 인정
  const hasProvinceOrCity =
    /(서울특별시|서울시|서울|경기도|경기|인천광역시|인천시|인천|강원특별자치도|강원도|강원|대전광역시|대전시|대전|세종특별자치시|세종시|세종|충청북도|충북|충청남도|충남|대구광역시|대구시|대구|경상북도|경북|부산광역시|부산시|부산|울산광역시|울산시|울산|경상남도|경남|광주광역시|광주시|광주|전라북도|전북|전라남도|전남|제주특별자치도|제주도|제주시|서귀포시|제주)/.test(text);

  if (hasProvinceOrCity) {
    return false;
  }

  // 예: "부경 - 박준영 수석 제보", "충청 - 이상진 지점장 제보"
  const memoPattern =
    /(제보|지원|수석|책임|프로|지점장|담당|소개건|진행|확인필요|주소확인|미상|없음|모름)/;

  const hasMemoWord = memoPattern.test(text);

  // 도로명/지번 주소 흔적
  const hasDetailedAddressWord =
    /(시|군|구|읍|면|동|리|로|길|번길|산단|공단)/.test(text);

  if (hasMemoWord && !hasDetailedAddressWord) {
    return true;
  }

  return false;
}
