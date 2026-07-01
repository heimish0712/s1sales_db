/****************************************************
 * A파일 "수주확정/계약완료" → B파일 "2026정보통신유지보수" 연동
 *
 * 목적:
 * - A파일의 "수주확정/계약완료" 시트 값을
 * - B파일의 "2026정보통신유지보수" 시트로 연동
 *
 * 기준:
 * - A파일 A열 계약번호 = 고유 ID
 * - B파일 A열 계약번호 = 고유 ID
 *
 * 동작:
 * - 계약번호가 이미 B파일에 있으면 A파일 기준으로 A~K 갱신
 * - 계약번호가 B파일에 없으면 새 행 추가
 * - B파일 L/M/N열은 보호, 절대 건드리지 않음
 *
 * K열 부가세:
 * - A파일 R열이 "포함"인 경우에만 계산
 * - A파일 Q열 계약가가 부가세 포함 금액이라고 보고
 * - 부가세 = Q열 계약가 / 11
 *
 * 충돌 방지:
 * - TARGET_SHEET_NAME 같은 흔한 전역 const를 사용하지 않음
 * - 기존 다른 .gs 파일과 변수명 충돌 방지
 *
 * 사용 순서:
 * 1. 이 코드 전체를 A파일 Apps Script에 붙여넣기
 * 2. ITMAINT_installTriggers_2026() 실행
 * 3. 권한 승인
 * 4. ITMAINT_initialSync_2026() 실행
 ****************************************************/


/**
 * 설정값
 * 전역 const 충돌 방지를 위해 함수 안에서만 반환함.
 */
function ITMAINT_getConfig_2026_() {
  return {
    targetSpreadsheetId: "1gDg9NNGWXb772yxJgKl2ORmXXL79iypRInN7FEbQVT4",

    sourceSheetName: "수주확정/계약완료",
    targetSheetName: "2026정보통신유지보수",

    sourceStartRow: 2,
    targetStartRow: 8,

    sourceLastCol: 28, // A~AB
    targetLastCol: 14, // A~N

    /**
     * B파일에서 자동으로 쓰는 구간
     * A~K만 처리
     * L/M/N 보호
     */
    writableSegments: [
      { startCol: 1, colCount: 11 } // A~K
    ],

    /**
     * 매핑표
     *
     * A파일 → B파일
     * A 계약번호 → A 계약번호
     * H 지역 → B 사업팀
     * S 수행사 → C 진행사
     * P 선임유형 → D 계약등급
     * K 고객사명 → E 계약처명
     * C 계약일자 → F 수주일
     * Z 비상주선임 → G 선임
     * AA 유지점검 → H 유지점검
     * AB 성능점검 → I 성능점검
     * Q 계약가 → J 계약금액
     */
    columnMap: [
      { sourceCol: 1,  targetCol: 1  }, // A → A 계약번호
      { sourceCol: 8,  targetCol: 2  }, // H → B 지역/사업팀
      { sourceCol: 19, targetCol: 3  }, // S → C 수행사/진행사
      { sourceCol: 16, targetCol: 4  }, // P → D 선임유형/계약등급
      { sourceCol: 11, targetCol: 5  }, // K → E 고객사명/계약처명
      { sourceCol: 3,  targetCol: 6  }, // C → F 계약일자/수주일
      { sourceCol: 26, targetCol: 7  }, // Z → G 비상주선임/선임
      { sourceCol: 27, targetCol: 8  }, // AA → H 유지점검
      { sourceCol: 28, targetCol: 9  }, // AB → I 성능점검
      { sourceCol: 17, targetCol: 10 }  // Q → J 계약가/계약금액
    ]
  };
}


/**
 * 최초 1회 실행
 *
 * 기존 B파일 내용을 지우지 않음.
 * 계약번호 기준으로 A~K만 갱신/추가함.
 *
 * L/M/N열은 보호.
 */
function ITMAINT_initialSync_2026() {
  ITMAINT_runWithLock_2026_(() => {
    ITMAINT_syncAllRowsWithoutClear_2026_();
  });
}


/**
 * 셀 수정 시 자동 실행
 *
 * 설치형 onEdit 트리거로 실행됨.
 */
function ITMAINT_onEditSync_2026(e) {
  if (!e || !e.range) return;

  const config = ITMAINT_getConfig_2026_();

  const editedSheet = e.range.getSheet();
  if (editedSheet.getName() !== config.sourceSheetName) return;

  const editedStartRow = e.range.getRow();
  const editedLastRow = editedStartRow + e.range.getNumRows() - 1;

  if (editedLastRow < config.sourceStartRow) return;

  const startRow = Math.max(editedStartRow, config.sourceStartRow);
  const rowCount = editedLastRow - startRow + 1;

  ITMAINT_runWithLock_2026_(() => {
    ITMAINT_syncSourceRows_2026_(startRow, rowCount);
  });
}


/**
 * 행 추가, 시트 구조 변경 등 대응
 *
 * 설치형 onChange 트리거로 실행됨.
 */
function ITMAINT_onChangeSync_2026(e) {
  if (!e) return;

  const changeType = e.changeType;

  if (
    changeType === "INSERT_ROW" ||
    changeType === "INSERT_GRID" ||
    changeType === "REMOVE_ROW" ||
    changeType === "OTHER"
  ) {
    ITMAINT_runWithLock_2026_(() => {
      ITMAINT_syncAllRowsWithoutClear_2026_();
    });
  }
}


/**
 * 5분마다 전체 재검사
 *
 * onEdit/onChange가 놓치는 경우 대비.
 * 예:
 * - 다른 스크립트가 값을 바꾸는 경우
 * - 수식 결과가 바뀌는 경우
 * - 붙여넣기/대량 수정 후 일부 트리거가 애매하게 도는 경우
 */
function ITMAINT_timeDrivenSync_2026() {
  ITMAINT_runWithLock_2026_(() => {
    ITMAINT_syncAllRowsWithoutClear_2026_();
  });
}


/**
 * 트리거 설치용
 *
 * 최초 1회 실행.
 * 기존 구버전 트리거와 신버전 트리거를 삭제한 뒤 재설치함.
 */
function ITMAINT_installTriggers_2026() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const triggerFunctionNames = [
    // 신버전
    "ITMAINT_onEditSync_2026",
    "ITMAINT_onChangeSync_2026",
    "ITMAINT_timeDrivenSync_2026",

    // 혹시 남아 있을 구버전
    "onEditSync_정보통신유지보수",
    "onChangeSync_정보통신유지보수",
    "timeDrivenSync_정보통신유지보수"
  ];

  ScriptApp.getProjectTriggers().forEach(trigger => {
    const fn = trigger.getHandlerFunction();

    if (triggerFunctionNames.indexOf(fn) !== -1) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger("ITMAINT_onEditSync_2026")
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ScriptApp.newTrigger("ITMAINT_onChangeSync_2026")
    .forSpreadsheet(ss)
    .onChange()
    .create();

  ScriptApp.newTrigger("ITMAINT_timeDrivenSync_2026")
    .timeBased()
    .everyMinutes(5)
    .create();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    "정보통신유지보수 자동 연동 트리거 설치 완료",
    "설치 완료",
    5
  );
}


/**
 * 특정 행 범위만 동기화
 */
function ITMAINT_syncSourceRows_2026_(startRow, rowCount) {
  const config = ITMAINT_getConfig_2026_();

  const sourceSheet = ITMAINT_getSourceSheet_2026_();
  const targetSheet = ITMAINT_getTargetSheet_2026_();

  const sourceValues = sourceSheet
    .getRange(startRow, 1, rowCount, config.sourceLastCol)
    .getValues();

  const targetIdMap = ITMAINT_getTargetIdMap_2026_(targetSheet);

  sourceValues.forEach(sourceRow => {
    const uniqueId = ITMAINT_normalizeId_2026_(sourceRow[0]);

    if (!uniqueId) return;

    let targetRowNumber = targetIdMap[uniqueId];

    if (!targetRowNumber) {
      targetRowNumber = ITMAINT_getFirstEmptyTargetRow_2026_(targetSheet);
      targetIdMap[uniqueId] = targetRowNumber;
    }

    const targetRow = ITMAINT_makeTargetRow_2026_(sourceRow);

    ITMAINT_writeTargetRowsWritableColumns_2026_(
      targetSheet,
      targetRowNumber,
      [targetRow]
    );
  });
}


/**
 * 전체 행 동기화
 *
 * 기존 계약번호는 갱신.
 * 없는 계약번호는 추가.
 *
 * B파일에만 있고 A파일에는 없는 계약번호는 삭제하지 않음.
 */
function ITMAINT_syncAllRowsWithoutClear_2026_() {
  const config = ITMAINT_getConfig_2026_();
  const sourceSheet = ITMAINT_getSourceSheet_2026_();

  const lastRow = sourceSheet.getLastRow();
  if (lastRow < config.sourceStartRow) return;

  const rowCount = lastRow - config.sourceStartRow + 1;

  ITMAINT_syncSourceRows_2026_(config.sourceStartRow, rowCount);
}


/**
 * A파일 1행 데이터를 B파일 A~N 구조로 변환
 *
 * 실제 입력은 A~K만 함.
 */
function ITMAINT_makeTargetRow_2026_(sourceRow) {
  const config = ITMAINT_getConfig_2026_();

  const targetRow = new Array(config.targetLastCol).fill("");

  // 기본 매핑
  config.columnMap.forEach(map => {
    targetRow[map.targetCol - 1] = sourceRow[map.sourceCol - 1];
  });

  /**
   * B파일 K열 부가세
   *
   * A파일 R열이 "포함"일 때만 계산.
   * A파일 Q열 계약가가 부가세 포함 금액이라고 보고,
   * 부가세 = Q / 11
   */
  const contractAmount = ITMAINT_parseNumber_2026_(sourceRow[17 - 1]); // Q열
  const vatStatus = String(sourceRow[18 - 1]).trim(); // R열

  if (vatStatus === "포함" && contractAmount !== null) {
    targetRow[11 - 1] = Math.round(contractAmount / 11); // K열
  } else {
    targetRow[11 - 1] = "";
  }

  return targetRow;
}


/**
 * B파일 A열 계약번호를 읽어서
 * {계약번호: 행번호} 형태로 반환
 */
function ITMAINT_getTargetIdMap_2026_(targetSheet) {
  const config = ITMAINT_getConfig_2026_();

  const lastRow = targetSheet.getLastRow();
  const idMap = {};

  if (lastRow < config.targetStartRow) return idMap;

  const idValues = targetSheet
    .getRange(
      config.targetStartRow,
      1,
      lastRow - config.targetStartRow + 1,
      1
    )
    .getValues();

  idValues.forEach((row, index) => {
    const id = ITMAINT_normalizeId_2026_(row[0]);

    if (id) {
      if (!idMap[id]) {
        idMap[id] = config.targetStartRow + index;
      }
    }
  });

  return idMap;
}


/**
 * B파일에서 TARGET_START_ROW 이후
 * A열이 비어 있는 첫 행 찾기
 */
function ITMAINT_getFirstEmptyTargetRow_2026_(targetSheet) {
  const config = ITMAINT_getConfig_2026_();

  const maxRows = targetSheet.getMaxRows();

  if (maxRows < config.targetStartRow) {
    targetSheet.insertRowsAfter(maxRows, config.targetStartRow - maxRows);
  }

  const lastRow = Math.max(targetSheet.getLastRow(), config.targetStartRow);
  const rowCount = lastRow - config.targetStartRow + 1;

  const values = targetSheet
    .getRange(config.targetStartRow, 1, rowCount, 1)
    .getValues();

  for (let i = 0; i < values.length; i++) {
    const id = ITMAINT_normalizeId_2026_(values[i][0]);

    if (!id) {
      return config.targetStartRow + i;
    }
  }

  const newRow = lastRow + 1;

  ITMAINT_ensureTargetRows_2026_(targetSheet, newRow);

  return newRow;
}


/**
 * 필요한 행 수가 부족하면 행 추가
 */
function ITMAINT_ensureTargetRows_2026_(sheet, requiredLastRow) {
  const maxRows = sheet.getMaxRows();

  if (maxRows < requiredLastRow) {
    sheet.insertRowsAfter(maxRows, requiredLastRow - maxRows);
  }
}


/**
 * 숫자 변환
 *
 * 예:
 * 1,100,000
 * ₩1,100,000
 * 1100000원
 */
function ITMAINT_parseNumber_2026_(value) {
  if (value === "" || value === null || value === undefined) return null;

  if (typeof value === "number") return value;

  const cleaned = String(value)
    .replace(/[^\d.-]/g, "")
    .trim();

  if (cleaned === "") return null;

  const num = Number(cleaned);

  if (isNaN(num)) return null;

  return num;
}


/**
 * 계약번호 정규화
 */
function ITMAINT_normalizeId_2026_(value) {
  if (value === "" || value === null || value === undefined) return "";

  return String(value).trim();
}


/**
 * A파일 원본 시트 가져오기
 */
function ITMAINT_getSourceSheet_2026_() {
  const config = ITMAINT_getConfig_2026_();

  const sourceSheet = SpreadsheetApp
    .getActiveSpreadsheet()
    .getSheetByName(config.sourceSheetName);

  if (!sourceSheet) {
    throw new Error(
      `A파일에서 "${config.sourceSheetName}" 시트를 찾을 수 없음. 실제 시트 탭 이름을 확인해.`
    );
  }

  return sourceSheet;
}


/**
 * B파일 대상 시트 가져오기
 */
function ITMAINT_getTargetSheet_2026_() {
  const config = ITMAINT_getConfig_2026_();

  const targetSheet = SpreadsheetApp
    .openById(config.targetSpreadsheetId)
    .getSheetByName(config.targetSheetName);

  if (!targetSheet) {
    throw new Error(
      `B파일에서 "${config.targetSheetName}" 시트를 찾을 수 없음. B파일 시트 탭 이름을 확인해.`
    );
  }

  return targetSheet;
}


/**
 * 연동 시 자동 처리 대상 열만 입력
 *
 * 현재 입력 대상:
 * - A~K
 *
 * 보호 대상:
 * - L
 * - M
 * - N
 */
function ITMAINT_writeTargetRowsWritableColumns_2026_(targetSheet, startRow, rows) {
  if (!rows || rows.length === 0) return;

  const config = ITMAINT_getConfig_2026_();

  config.writableSegments.forEach(segment => {
    const values = rows.map(row => {
      return row.slice(
        segment.startCol - 1,
        segment.startCol - 1 + segment.colCount
      );
    });

    targetSheet
      .getRange(
        startRow,
        segment.startCol,
        values.length,
        segment.colCount
      )
      .setValues(values);
  });
}


/**
 * 동시 실행 방지
 *
 * onEdit, onChange, 5분 트리거가 겹치면
 * 같은 행을 동시에 쓰는 꼴이 날 수 있어서 잠금 처리.
 */
function ITMAINT_runWithLock_2026_(callback) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30000)) {
    throw new Error("다른 정보통신유지보수 동기화 작업이 실행 중이라 이번 실행은 중단됨");
  }

  try {
    callback();
  } finally {
    lock.releaseLock();
  }
}

function ITMAINT_forceUnlock_2026() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();

  Object.keys(all).forEach(key => {
    if (
      key.includes("ITMAINT") &&
      (
        key.toUpperCase().includes("LOCK") ||
        key.toUpperCase().includes("RUNNING") ||
        key.toUpperCase().includes("SYNC")
      )
    ) {
      props.deleteProperty(key);
      Logger.log("삭제한 락 키: " + key);
    }
  });

  Logger.log("정보통신유지보수 동기화 락 강제 해제 완료");
}
