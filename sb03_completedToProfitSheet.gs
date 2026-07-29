/****************************************************
 * A파일 "수주확정/계약완료" → B파일 "2026정보통신유지보수" 연동
 *
 * [긴급 매핑 수정본]
 * - 열 번호가 아니라 실제 헤더명을 검증한 뒤 매핑한다.
 * - B파일의 현재 40열 구조에 맞춰 입력 열을 다시 고정한다.
 * - 수식/수동관리 열은 절대 쓰지 않는다.
 *
 * 입력 대상(B파일):
 * A:K  계약 기본정보
 * N:R  선임/점검/계약금액/VAT
 * X:Y  청구 특이사항/세금계산서 이메일
 * AJ   소개자
 *
 * 보호 대상(B파일):
 * L:M, S:W, Z:AI, AK:AN
 *
 * 추가 기준:
 * - 계약시작일/종료일/계약개월/선임·점검 횟수는
 *   "마스터시트(신규)"를 계약번호(발주번호) 우선,
 *   고객번호 차선으로 찾아 가져온다.
 * - 마스터 매칭이 없으면 수주확정 시트 값과 계약기간 문자열을
 *   안전한 범위에서만 보조값으로 사용한다.
 ****************************************************/

function ITMAINT_getConfig_2026_() {
  return {
    targetSpreadsheetId: "1gDg9NNGWXb772yxJgKl2ORmXXL79iypRInN7FEbQVT4",

    sourceSheetName: "수주확정/계약완료",
    masterSheetName: "마스터시트(신규)",
    targetSheetName: "2026정보통신유지보수",

    sourceHeaderRow: 1,
    sourceStartRow: 2,

    masterHeaderRow: 2,
    masterStartRow: 3,

    targetHeaderRow: 6,
    targetStartRow: 8,
    targetLastCol: 40, // A~AN

    /**
     * B파일에서 자동으로 쓰는 구간.
     * 그 외 열은 수식/수동값 보호를 위해 절대 쓰지 않는다.
     */
    writableSegments: [
      { startCol: 1, colCount: 11 }, // A:K
      { startCol: 14, colCount: 5 }, // N:R
      { startCol: 24, colCount: 2 }, // X:Y
      { startCol: 36, colCount: 1 }  // AJ 소개자
    ],

    /**
     * 현재 B파일 구조를 강제 검증한다.
     * 하나라도 다르면 쓰기 전에 즉시 중단한다.
     */
    expectedTargetHeaders: {
      1: "계약번호",
      2: "권역",
      3: "수행사",
      4: "계약등급",
      5: "담당자",
      6: "계약처명",
      7: "수주일",
      8: "계약서상계약기간",
      9: "계약시작일",
      10: "계약종료일",
      11: "계약기간개월",
      14: "선임",
      15: "유지점검",
      16: "성능점검",
      17: "계약서상계약금액",
      18: "부가세적용여부",
      24: "청구요청사항및계약특이사항",
      25: "세금계산서요청이메일",
      36: "소개자"
    }
  };
}


/**
 * 최초/수동 전체 동기화.
 * 기존 대상 행을 삭제하지 않고 계약번호 기준으로 갱신·추가한다.
 */
function ITMAINT_initialSync_2026() {
  // 15단계 이후 공개 초기동기화도 신규 계약 append-only 방식으로 제한한다.
  return ITMNEW_syncMissingContractsNow_2026();
}


/**
 * 수주확정/계약완료 편집 행만 동기화.
 */
function ITMAINT_onEditSync_2026(e) {
  // 구형 트리거 호환용. 실제 처리는 신규 계약 append-only 모듈로 위임한다.
  return ITMNEW_syncFromEdit_2026(e);
}


/**
 * 구형 onChange 호환 핸들러.
 * 정식 중앙 onChange에서는 전체보정 플래그를 사용한다.
 */
function ITMAINT_onChangeSync_2026(e) {
  if (!e) return { status: "IGNORED_INVALID_EVENT" };

  var changeType = String(e.changeType || "").toUpperCase();

  if (
    changeType === "INSERT_ROW" ||
    changeType === "INSERT_GRID" ||
    changeType === "REMOVE_ROW" ||
    changeType === "OTHER"
  ) {
    return ITMNEW_runMissingContractSyncForPipeline_2026();
  }

  return { status: "IGNORED_NON_STRUCTURAL_CHANGE" };
}


function ITMAINT_timeDrivenSync_2026() {
  // 구형 시간 트리거 호환용. 기존 행 갱신 없이 누락 신규 계약만 이식한다.
  return ITMNEW_runMissingContractSyncForPipeline_2026();
}


function ITMAINT_runFullSyncForAutomationPipeline_2026() {
  // 함수명은 호환을 위해 유지하되 동작은 append-only 신규 이식으로 제한한다.
  return ITMNEW_runMissingContractSyncForPipeline_2026();
}


/**
 * 특정 수주확정 행 범위 동기화.
 */
function ITMAINT_syncSourceRows_2026_(startRow, rowCount) {
  var config = ITMAINT_getConfig_2026_();

  var sourceSheet = ITMAINT_getSourceSheet_2026_();
  var masterSheet = ITMAINT_getMasterSheet_2026_();
  var targetSheet = ITMAINT_getTargetSheet_2026_();

  // 어떤 값도 쓰기 전에 세 시트의 헤더를 모두 검증한다.
  var sourceSchema = ITMAINT_buildSchema_2026_(
    sourceSheet,
    config.sourceHeaderRow,
    ITMAINT_getRequiredSourceHeaders_2026_(),
    "수주확정/계약완료"
  );

  var masterSchema = ITMAINT_buildSchema_2026_(
    masterSheet,
    config.masterHeaderRow,
    ITMAINT_getRequiredMasterHeaders_2026_(),
    "마스터시트(신규)"
  );

  var targetSchema = ITMAINT_buildSchema_2026_(
    targetSheet,
    config.targetHeaderRow,
    [],
    "2026정보통신유지보수"
  );

  ITMAINT_validateTargetLayout_2026_(targetSchema);

  var sourceLastCol = sourceSheet.getLastColumn();
  var sourceValues = sourceSheet
    .getRange(startRow, 1, rowCount, sourceLastCol)
    .getValues();

  var masterLookup = ITMAINT_buildMasterLookup_2026_(
    masterSheet,
    masterSchema
  );

  var targetIdMap = ITMAINT_getTargetIdMap_2026_(targetSheet);

  var syncedRows = 0;
  var skippedNoId = 0;
  var insertedRows = 0;
  var updatedRows = 0;
  var missingMasterRows = 0;

  sourceValues.forEach(function (sourceRow) {
    var contractNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, "계약번호")
    );

    if (!contractNo) {
      skippedNoId++;
      return;
    }

    var customerNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, "고객번호")
    );

    var masterRow = masterLookup.byContractNo[contractNo] ||
      masterLookup.byCustomerNo[customerNo] ||
      null;

    if (!masterRow) missingMasterRows++;

    var targetRowNumber = targetIdMap[contractNo];
    var existed = !!targetRowNumber;

    if (!targetRowNumber) {
      targetRowNumber = ITMAINT_getFirstEmptyTargetRow_2026_(targetSheet);
      targetIdMap[contractNo] = targetRowNumber;
    }

    var targetRow = ITMAINT_makeTargetRow_2026_(
      sourceRow,
      sourceSchema,
      masterRow,
      masterSchema
    );

    ITMAINT_writeTargetRowsWritableColumns_2026_(
      targetSheet,
      targetRowNumber,
      [targetRow]
    );

    syncedRows++;

    if (existed) {
      updatedRows++;
    } else {
      insertedRows++;
    }
  });

  return {
    sourceRows: sourceValues.length,
    syncedRows: syncedRows,
    skippedNoId: skippedNoId,
    insertedRows: insertedRows,
    updatedRows: updatedRows,
    missingMasterRows: missingMasterRows,
    mappingVersion: "HEADER_SAFE_V2"
  };
}


/**
 * 전체 행 동기화.
 */
function ITMAINT_syncAllRowsWithoutClear_2026_() {
  var config = ITMAINT_getConfig_2026_();
  var sourceSheet = ITMAINT_getSourceSheet_2026_();

  var lastRow = sourceSheet.getLastRow();

  if (lastRow < config.sourceStartRow) {
    return {
      sourceRows: 0,
      syncedRows: 0,
      skippedNoId: 0,
      insertedRows: 0,
      updatedRows: 0,
      missingMasterRows: 0,
      mappingVersion: "HEADER_SAFE_V2"
    };
  }

  return ITMAINT_syncSourceRows_2026_(
    config.sourceStartRow,
    lastRow - config.sourceStartRow + 1
  );
}


/**
 * 수주확정 1행 + 마스터 1행을 대상 A:AN 구조로 변환.
 * 실제 쓰기는 config.writableSegments만 수행한다.
 */
function ITMAINT_makeTargetRow_2026_(
  sourceRow,
  sourceSchema,
  masterRow,
  masterSchema
) {
  var config = ITMAINT_getConfig_2026_();
  var targetRow = new Array(config.targetLastCol).fill("");

  function source(headerName) {
    return ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, headerName);
  }

  function master(headerName) {
    if (!masterRow) return "";
    return ITMAINT_getByHeader_2026_(masterRow, masterSchema, headerName);
  }

  var contractNo = source("계약번호");
  var customerNo = source("고객번호");
  var contractPeriod = source("계약기간");

  var parsedPeriod = ITMAINT_parseContractPeriod_2026_(contractPeriod);

  var startDate = master("계약시작일") || parsedPeriod.startDate || "";
  var endDate = master("계약종료일") || parsedPeriod.endDate || "";

  var contractMonths = ITMAINT_parseCount_2026_(master("계약단위"));

  if (contractMonths === null) {
    contractMonths = parsedPeriod.months;
  }

  if (contractMonths === null) {
    contractMonths = ITMAINT_parseCount_2026_(source("비상주선임"));
  }

  var appointmentMonths = ITMAINT_getAppointmentMonths_2026_(
    master("관리자선임여부"),
    contractMonths,
    source("비상주선임")
  );

  var maintenanceCount = ITMAINT_parseCount_2026_(master("유지점검"));

  if (maintenanceCount === null) {
    maintenanceCount = ITMAINT_parseCount_2026_(source("유지점검"));
  }

  var performanceCount = ITMAINT_parseCount_2026_(master("성능점검"));

  if (performanceCount === null) {
    performanceCount = ITMAINT_parseCount_2026_(source("성능점검"));
  }

  // A:K
  targetRow[0] = contractNo;                                  // A 계약번호
  targetRow[1] = source("지역");                              // B 권역
  targetRow[2] = source("수행사");                            // C 수행사
  targetRow[3] = source("선임유형");                          // D 계약등급
  targetRow[4] = source("계약담당자");                        // E 담당자
  targetRow[5] = source("고객사명");                          // F 계약처명
  targetRow[6] = source("계약일자발주번호부여일");            // G 수주일
  targetRow[7] = contractPeriod ||                            // H 계약서상 계약 기간
    ITMAINT_composeContractPeriod_2026_(startDate, endDate);
  targetRow[8] = startDate;                                  // I 계약시작일
  targetRow[9] = endDate;                                    // J 계약종료일
  targetRow[10] = contractMonths === null ? "" : contractMonths; // K 계약기간(개월)

  // N:R
  targetRow[13] = appointmentMonths === null ? "" : appointmentMonths; // N 선임
  targetRow[14] = maintenanceCount === null ? "" : maintenanceCount;   // O 유지점검
  targetRow[15] = performanceCount === null ? "" : performanceCount;   // P 성능점검
  targetRow[16] = source("계약가");                                     // Q 계약서상 계약금액
  targetRow[17] = ITMAINT_normalizeVatLabel_2026_(source("vat"));       // R 부가세 적용 여부

  // X:Y
  targetRow[23] = source("청구등메모");                         // X 청구 요청사항 및 계약 특이사항
  targetRow[24] = source("세금계산서요청이메일");               // Y 세금계산서 요청 이메일

  // AJ
  targetRow[35] = source("제보자");                             // AJ 소개자

  return targetRow;
}


/**
 * 대상 현재 헤더가 예상 열과 정확히 맞는지 검사.
 * 틀리면 데이터를 한 칸도 쓰지 않고 중단한다.
 */
function ITMAINT_validateTargetLayout_2026_(targetSchema) {
  var config = ITMAINT_getConfig_2026_();
  var mismatches = [];

  Object.keys(config.expectedTargetHeaders).forEach(function (columnText) {
    var column = Number(columnText);
    var expected = ITMAINT_normalizeHeader_2026_(
      config.expectedTargetHeaders[column]
    );
    var actual = targetSchema.normalizedHeaders[column - 1] || "";

    if (actual !== expected) {
      mismatches.push(
        ITMAINT_columnToLetter_2026_(column) +
        "열: 예상=[" + config.expectedTargetHeaders[column] +
        "], 실제=[" + (targetSchema.rawHeaders[column - 1] || "") + "]"
      );
    }
  });

  if (mismatches.length) {
    throw new Error(
      "2026정보통신유지보수 헤더 구조가 예상과 달라 동기화를 중단했습니다.\n" +
      mismatches.join("\n")
    );
  }
}


/**
 * 헤더 스키마 생성 + 필수 헤더 검사.
 */
function ITMAINT_buildSchema_2026_(
  sheet,
  headerRow,
  requiredHeaders,
  label
) {
  var lastCol = sheet.getLastColumn();

  if (lastCol < 1) {
    throw new Error(label + " 시트에 헤더가 없습니다.");
  }

  var rawHeaders = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  var normalizedHeaders = rawHeaders.map(ITMAINT_normalizeHeader_2026_);
  var indexByHeader = {};

  normalizedHeaders.forEach(function (header, index) {
    if (!header) return;

    if (Object.prototype.hasOwnProperty.call(indexByHeader, header)) {
      throw new Error(
        label + " 시트에 중복 헤더가 있습니다: " + rawHeaders[index]
      );
    }

    indexByHeader[header] = index;
  });

  (requiredHeaders || []).forEach(function (requiredHeader) {
    var normalized = ITMAINT_normalizeHeader_2026_(requiredHeader);

    if (!Object.prototype.hasOwnProperty.call(indexByHeader, normalized)) {
      throw new Error(
        label + " 시트에서 필수 헤더를 찾을 수 없습니다: " + requiredHeader
      );
    }
  });

  return {
    rawHeaders: rawHeaders,
    normalizedHeaders: normalizedHeaders,
    indexByHeader: indexByHeader,
    lastCol: lastCol
  };
}


function ITMAINT_getRequiredSourceHeaders_2026_() {
  return [
    "계약번호",
    "고객번호",
    "계약일자발주번호부여일",
    "지역",
    "제보자",
    "계약담당자",
    "고객사명",
    "선임유형",
    "계약가",
    "vat",
    "수행사",
    "계약기간",
    "비상주선임",
    "유지점검",
    "성능점검",
    "청구등메모",
    "세금계산서요청이메일"
  ];
}


function ITMAINT_getRequiredMasterHeaders_2026_() {
  return [
    "고객번호",
    "발주번호",
    "계약시작일",
    "계약종료일",
    "계약단위",
    "관리자선임여부",
    "유지점검",
    "성능점검"
  ];
}


/**
 * 마스터를 계약번호 우선/고객번호 차선 조회용 맵으로 구성.
 */
function ITMAINT_buildMasterLookup_2026_(masterSheet, masterSchema) {
  var config = ITMAINT_getConfig_2026_();
  var lastRow = masterSheet.getLastRow();

  var result = {
    byContractNo: {},
    byCustomerNo: {}
  };

  if (lastRow < config.masterStartRow) return result;

  var rows = masterSheet
    .getRange(
      config.masterStartRow,
      1,
      lastRow - config.masterStartRow + 1,
      masterSchema.lastCol
    )
    .getValues();

  rows.forEach(function (row) {
    var contractNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(row, masterSchema, "발주번호")
    );

    var customerNo = ITMAINT_normalizeId_2026_(
      ITMAINT_getByHeader_2026_(row, masterSchema, "고객번호")
    );

    if (contractNo && !result.byContractNo[contractNo]) {
      result.byContractNo[contractNo] = row;
    }

    if (customerNo && !result.byCustomerNo[customerNo]) {
      result.byCustomerNo[customerNo] = row;
    }
  });

  return result;
}


function ITMAINT_getByHeader_2026_(row, schema, headerName) {
  var normalized = ITMAINT_normalizeHeader_2026_(headerName);
  var index = schema.indexByHeader[normalized];

  if (index === undefined) return "";

  return row[index];
}


function ITMAINT_normalizeHeader_2026_(value) {
  return String(value === null || value === undefined ? "" : value)
    .replace(/\s+/g, "")
    .replace(/[()（）]/g, "")
    .replace(/[·ㆍ]/g, "")
    .replace(/[_\-\/]/g, "")
    .trim()
    .toLowerCase();
}


function ITMAINT_parseContractPeriod_2026_(value) {
  var text = String(value === null || value === undefined ? "" : value).trim();

  if (!text) {
    return {
      startDate: "",
      endDate: "",
      months: null
    };
  }

  var dateParts = text.match(
    /(\d{2,4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/g
  ) || [];

  if (dateParts.length < 2) {
    return {
      startDate: "",
      endDate: "",
      months: null
    };
  }

  var startDate = ITMAINT_parseDateText_2026_(dateParts[0]);
  var endDate = ITMAINT_parseDateText_2026_(dateParts[1]);
  var months = ITMAINT_calculateContractMonths_2026_(startDate, endDate);

  return {
    startDate: startDate || "",
    endDate: endDate || "",
    months: months
  };
}


function ITMAINT_parseDateText_2026_(value) {
  var parts = String(value).match(/(\d{2,4})\D+(\d{1,2})\D+(\d{1,2})/);

  if (!parts) return null;

  var year = Number(parts[1]);
  var month = Number(parts[2]);
  var day = Number(parts[3]);

  if (year < 100) year += 2000;

  var date = new Date(year, month - 1, day);

  if (
    date.getFullYear() !== year ||
    date.getMonth() !== month - 1 ||
    date.getDate() !== day
  ) {
    return null;
  }

  return date;
}


function ITMAINT_calculateContractMonths_2026_(startDate, endDate) {
  if (
    !(startDate instanceof Date) ||
    isNaN(startDate.getTime()) ||
    !(endDate instanceof Date) ||
    isNaN(endDate.getTime())
  ) {
    return null;
  }

  var months =
    (endDate.getFullYear() - startDate.getFullYear()) * 12 +
    (endDate.getMonth() - startDate.getMonth());

  if (endDate.getDate() >= startDate.getDate()) {
    months += 1;
  }

  return months > 0 ? months : null;
}


function ITMAINT_composeContractPeriod_2026_(startDate, endDate) {
  if (!startDate || !endDate) return "";

  return (
    ITMAINT_formatDateForPeriod_2026_(startDate) +
    "~" +
    ITMAINT_formatDateForPeriod_2026_(endDate)
  );
}


function ITMAINT_formatDateForPeriod_2026_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) {
    return Utilities.formatDate(
      value,
      Session.getScriptTimeZone() || "Asia/Seoul",
      "yy.MM.dd."
    );
  }

  return String(value || "").trim();
}


function ITMAINT_parseCount_2026_(value) {
  if (value === "" || value === null || value === undefined) return null;

  if (typeof value === "number") {
    return isNaN(value) ? null : value;
  }

  var match = String(value).replace(/,/g, "").match(/-?\d+(?:\.\d+)?/);

  if (!match) return null;

  var number = Number(match[0]);

  return isNaN(number) ? null : number;
}


function ITMAINT_getAppointmentMonths_2026_(
  masterAppointmentValue,
  contractMonths,
  sourceFallback
) {
  var text = String(
    masterAppointmentValue === null ||
    masterAppointmentValue === undefined
      ? ""
      : masterAppointmentValue
  ).trim();

  if (/(미선임|선임안함|해당없음|불필요|없음|^x$|^0$)/i.test(text)) {
    return 0;
  }

  if (/선임/.test(text) && contractMonths !== null) {
    return contractMonths;
  }

  var fallback = ITMAINT_parseCount_2026_(sourceFallback);

  if (fallback !== null) return fallback;

  return contractMonths;
}


function ITMAINT_normalizeVatLabel_2026_(value) {
  var text = String(value === null || value === undefined ? "" : value)
    .replace(/\s+/g, "")
    .trim();

  if (!text) return "";

  if (text === "포함" || text === "부포" || /부가세포함/.test(text)) {
    return "부포";
  }

  if (text === "별도" || text === "부별" || /부가세별도/.test(text)) {
    return "부별";
  }

  return text;
}


/**
 * 대상 계약번호 → 행번호 맵.
 */
function ITMAINT_getTargetIdMap_2026_(targetSheet) {
  var config = ITMAINT_getConfig_2026_();
  var lastRow = targetSheet.getLastRow();
  var idMap = {};

  if (lastRow < config.targetStartRow) return idMap;

  var idValues = targetSheet
    .getRange(
      config.targetStartRow,
      1,
      lastRow - config.targetStartRow + 1,
      1
    )
    .getValues();

  idValues.forEach(function (row, index) {
    var id = ITMAINT_normalizeId_2026_(row[0]);

    if (id && !idMap[id]) {
      idMap[id] = config.targetStartRow + index;
    }
  });

  return idMap;
}


function ITMAINT_getFirstEmptyTargetRow_2026_(targetSheet) {
  var config = ITMAINT_getConfig_2026_();
  var maxRows = targetSheet.getMaxRows();

  if (maxRows < config.targetStartRow) {
    targetSheet.insertRowsAfter(
      maxRows,
      config.targetStartRow - maxRows
    );
  }

  var lastRow = Math.max(targetSheet.getLastRow(), config.targetStartRow);
  var rowCount = lastRow - config.targetStartRow + 1;

  var values = targetSheet
    .getRange(config.targetStartRow, 1, rowCount, 1)
    .getValues();

  for (var i = 0; i < values.length; i++) {
    if (!ITMAINT_normalizeId_2026_(values[i][0])) {
      return config.targetStartRow + i;
    }
  }

  var newRow = lastRow + 1;

  ITMAINT_ensureTargetRows_2026_(targetSheet, newRow);

  return newRow;
}


function ITMAINT_ensureTargetRows_2026_(sheet, requiredLastRow) {
  var maxRows = sheet.getMaxRows();

  if (maxRows < requiredLastRow) {
    sheet.insertRowsAfter(maxRows, requiredLastRow - maxRows);
  }
}


/**
 * A:K, N:R, X:Y, AJ만 기록한다.
 */
function ITMAINT_writeTargetRowsWritableColumns_2026_(
  targetSheet,
  startRow,
  rows
) {
  if (!rows || rows.length === 0) return;

  var config = ITMAINT_getConfig_2026_();

  config.writableSegments.forEach(function (segment) {
    var values = rows.map(function (row) {
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


function ITMAINT_normalizeId_2026_(value) {
  if (value === "" || value === null || value === undefined) return "";

  if (typeof value === "number" && isFinite(value)) {
    return String(value).replace(/\.0+$/, "");
  }

  return String(value).trim().replace(/\.0+$/, "");
}


function ITMAINT_columnToLetter_2026_(column) {
  var result = "";
  var value = column;

  while (value > 0) {
    var remainder = (value - 1) % 26;
    result = String.fromCharCode(65 + remainder) + result;
    value = Math.floor((value - 1) / 26);
  }

  return result;
}


function ITMAINT_getSourceSheet_2026_() {
  var config = ITMAINT_getConfig_2026_();
  var sourceSheet = AUTOMATION_getRuntimeMasterSpreadsheet_()
    .getSheetByName(config.sourceSheetName);

  if (!sourceSheet) {
    throw new Error(
      'A파일에서 "' + config.sourceSheetName +
      '" 시트를 찾을 수 없습니다.'
    );
  }

  return sourceSheet;
}


function ITMAINT_getMasterSheet_2026_() {
  var config = ITMAINT_getConfig_2026_();
  var masterSheet = AUTOMATION_getRuntimeMasterSpreadsheet_()
    .getSheetByName(config.masterSheetName);

  if (!masterSheet) {
    throw new Error(
      'A파일에서 "' + config.masterSheetName +
      '" 시트를 찾을 수 없습니다.'
    );
  }

  return masterSheet;
}


function ITMAINT_getTargetSheet_2026_() {
  var config = ITMAINT_getConfig_2026_();
  var targetSheet = SpreadsheetApp
    .openById(config.targetSpreadsheetId)
    .getSheetByName(config.targetSheetName);

  if (!targetSheet) {
    throw new Error(
      'B파일에서 "' + config.targetSheetName +
      '" 시트를 찾을 수 없습니다.'
    );
  }

  return targetSheet;
}


function ITMAINT_runWithLock_2026_(callback) {
  return AUTOMATION_runWithModuleLeaseOrThrow_(
    "IT_MAINTENANCE_SYNC",
    "ITMAINT_runWithLock_2026_",
    callback,
    { waitMs: 1000, ttlMs: 8 * 60 * 1000 }
  );
}


function ITMAINT_resetModuleLease_2026() {
  var props = PropertiesService.getScriptProperties();
  var leaseKey =
    AUTOMATION_RUNTIME_CONFIG.leasePropertyPrefix +
    "IT_MAINTENANCE_SYNC";
  var existed = !!props.getProperty(leaseKey);

  props.deleteProperty(leaseKey);

  Logger.log(
    existed
      ? "정보통신유지보수 기능별 lease를 삭제했습니다: " + leaseKey
      : "삭제할 정보통신유지보수 lease가 없습니다."
  );

  return {
    deleted: existed,
    propertyKey: leaseKey
  };
}
