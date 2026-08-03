/****************************************************
 * A파일 "수주확정/계약완료" → B파일 "2026정보통신유지보수" 연동
 *
 * [긴급 매핑 수정본]
 * - 열 번호가 아니라 실제 헤더명을 검증한 뒤 매핑한다.
 * - 대상 열번호를 고정하지 않고 6행 실제 헤더명으로 찾는다.
 * - 수식/수동관리 열은 절대 쓰지 않는다.
 *
 * 입력 대상(B파일): 계약번호·권역·수행사·계약등급·담당자·계약처명·수주일,
 * 계약기간·선임·점검·계약금액·VAT·청구메모·세금계산서 이메일·소개자.
 * 각 열은 현재 헤더명으로 탐색하고, 그 외 수식/수동관리 열은 기록하지 않는다.
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
    version: "2026-08-02-PHASE26-HEADER-DYNAMIC",
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

    /**
     * PHASE26: 대상 열번호를 전혀 고정하지 않는다.
     * 2026정보통신유지보수 6행의 실제 헤더를 정규화해 논리 필드와 연결한다.
     * 중간에 수익·정산 열이 추가되거나 소개자 열이 이동해도 헤더명만 유지되면 동작한다.
     */
    targetFieldDefinitions: {
      contractNo: {
        label: "계약번호",
        aliases: ["계약번호"]
      },
      region: {
        label: "권역",
        aliases: ["권역", "지역"]
      },
      vendor: {
        label: "수행사",
        aliases: ["수행사"]
      },
      contractGrade: {
        label: "계약등급",
        aliases: ["계약등급", "선임유형"]
      },
      manager: {
        label: "담당자",
        aliases: ["담당자", "계약담당자"]
      },
      customerName: {
        label: "계약처명",
        aliases: ["계약처명", "고객사명"]
      },
      orderDate: {
        label: "수주일",
        aliases: ["수주일", "계약일자(발주번호 부여일)", "계약일자발주번호부여일"]
      },
      contractPeriod: {
        label: "계약서 상 계약 기간",
        aliases: ["계약서 상 계약 기간", "계약서상 계약기간"]
      },
      startDate: {
        label: "계약시작일",
        aliases: ["계약시작일"]
      },
      endDate: {
        label: "계약종료일",
        aliases: ["계약종료일"]
      },
      contractMonths: {
        label: "계약기간(개월)",
        aliases: ["계약기간(개월)", "계약기간 개월"]
      },
      appointment: {
        label: "선임",
        aliases: ["선임", "비상주선임"]
      },
      maintenance: {
        label: "유지점검",
        aliases: ["유지점검"]
      },
      performance: {
        label: "성능점검",
        aliases: ["성능점검"]
      },
      contractPrice: {
        label: "계약서상 계약금액",
        aliases: ["계약서상 계약금액", "계약서 상 계약금액"]
      },
      vat: {
        label: "부가세 적용 여부",
        aliases: ["부가세 적용 여부", "부가세적용여부", "VAT"]
      },
      memo: {
        label: "청구 요청사항 및 계약 특이사항",
        aliases: ["청구 요청사항 및 계약 특이사항", "청구요청사항및계약특이사항", "청구 등 메모"]
      },
      invoiceEmail: {
        label: "세금계산서 요청 이메일",
        aliases: ["세금계산서 요청 이메일", "세금계산서요청이메일"]
      },
      referrer: {
        label: "소개자",
        aliases: ["소개자", "제보자"]
      }
    },

    targetRequiredFieldKeys: [
      "contractNo", "region", "vendor", "contractGrade", "manager",
      "customerName", "orderDate", "contractPeriod", "startDate", "endDate",
      "contractMonths", "appointment", "maintenance", "performance",
      "contractPrice", "vat", "memo", "invoiceEmail", "referrer"
    ]
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

  var targetFieldMap = ITMAINT_validateTargetLayout_2026_(targetSchema);

  var sourceLastCol = sourceSheet.getLastColumn();
  var sourceValues = sourceSheet
    .getRange(startRow, 1, rowCount, sourceLastCol)
    .getValues();

  var masterLookup = ITMAINT_buildMasterLookup_2026_(
    masterSheet,
    masterSchema
  );

  var targetIdMap = ITMAINT_getTargetIdMap_2026_(targetSheet, targetFieldMap);

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
      targetRowNumber = ITMAINT_getFirstEmptyTargetRow_2026_(targetSheet, targetFieldMap);
      targetIdMap[contractNo] = targetRowNumber;
    }

    var targetRecord = ITMAINT_makeTargetRecord_2026_(
      sourceRow,
      sourceSchema,
      masterRow,
      masterSchema
    );

    ITMAINT_writeTargetRecordByHeader_2026_(
      targetSheet,
      targetRowNumber,
      targetRecord,
      targetFieldMap
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
    mappingVersion: "HEADER_DYNAMIC_V3"
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
      mappingVersion: "HEADER_DYNAMIC_V3"
    };
  }

  return ITMAINT_syncSourceRows_2026_(
    config.sourceStartRow,
    lastRow - config.sourceStartRow + 1
  );
}


/**
 * 수주확정 1행 + 마스터 1행을 대상 논리 필드 객체로 변환.
 * 실제 열 위치는 대상 6행 헤더에서 실행 시점에 계산한다.
 */
function ITMAINT_makeTargetRecord_2026_(
  sourceRow,
  sourceSchema,
  masterRow,
  masterSchema
) {
  function source(headerName) {
    return ITMAINT_getByHeader_2026_(sourceRow, sourceSchema, headerName);
  }

  function master(headerName) {
    if (!masterRow) return "";
    return ITMAINT_getByHeader_2026_(masterRow, masterSchema, headerName);
  }

  var contractNo = source("계약번호");
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

  return {
    contractNo: contractNo,
    region: source("지역"),
    vendor: source("수행사"),
    contractGrade: source("선임유형"),
    manager: source("계약담당자"),
    customerName: source("고객사명"),
    orderDate: source("계약일자발주번호부여일"),
    contractPeriod: contractPeriod || ITMAINT_composeContractPeriod_2026_(startDate, endDate),
    startDate: startDate,
    endDate: endDate,
    contractMonths: contractMonths === null ? "" : contractMonths,
    appointment: appointmentMonths === null ? "" : appointmentMonths,
    maintenance: maintenanceCount === null ? "" : maintenanceCount,
    performance: performanceCount === null ? "" : performanceCount,
    contractPrice: source("계약가"),
    vat: ITMAINT_normalizeVatLabel_2026_(source("vat")),
    memo: source("청구등메모"),
    invoiceEmail: source("세금계산서요청이메일"),
    referrer: source("제보자")
  };
}


/**
 * 구형 내부 호출 호환용 별칭.
 * PHASE26부터 반환값은 고정 A:AN 배열이 아니라 논리 필드 객체다.
 */
function ITMAINT_makeTargetRow_2026_(
  sourceRow,
  sourceSchema,
  masterRow,
  masterSchema
) {
  return ITMAINT_makeTargetRecord_2026_(
    sourceRow,
    sourceSchema,
    masterRow,
    masterSchema
  );
}


/**
 * 대상 헤더를 실제 이름으로 찾아 논리 필드와 연결한다.
 * 틀리면 데이터를 한 칸도 쓰지 않고 중단한다.
 */
function ITMAINT_validateTargetLayout_2026_(targetSchema) {
  return ITMAINT_buildTargetFieldMap_2026_(targetSchema);
}


function ITMAINT_buildTargetFieldMap_2026_(targetSchema) {
  var config = ITMAINT_getConfig_2026_();
  var definitions = config.targetFieldDefinitions || {};
  var required = config.targetRequiredFieldKeys || Object.keys(definitions);
  var columnByField = {};
  var indexByField = {};
  var rawHeaderByField = {};
  var missing = [];
  var ambiguous = [];

  Object.keys(definitions).forEach(function(fieldKey) {
    var definition = definitions[fieldKey] || {};
    var aliases = definition.aliases || [definition.label || fieldKey];
    var matchedIndexes = [];

    aliases.forEach(function(alias) {
      var normalized = ITMAINT_normalizeHeader_2026_(alias);
      var index = targetSchema.indexByHeader[normalized];
      if (index === undefined) return;
      if (matchedIndexes.indexOf(index) < 0) matchedIndexes.push(index);
    });

    if (matchedIndexes.length === 0) {
      if (required.indexOf(fieldKey) >= 0) {
        missing.push((definition.label || fieldKey) + ' [' + aliases.join(' / ') + ']');
      }
      return;
    }

    if (matchedIndexes.length > 1) {
      ambiguous.push(
        (definition.label || fieldKey) + ': ' + matchedIndexes.map(function(index) {
          return ITMAINT_columnToLetter_2026_(index + 1) + '열 [' +
            String(targetSchema.rawHeaders[index] || '') + ']';
        }).join(', ')
      );
      return;
    }

    var matchedIndex = matchedIndexes[0];
    columnByField[fieldKey] = matchedIndex + 1;
    indexByField[fieldKey] = matchedIndex;
    rawHeaderByField[fieldKey] = String(targetSchema.rawHeaders[matchedIndex] || '');
  });

  if (missing.length || ambiguous.length) {
    var parts = [
      '2026정보통신유지보수 헤더 기반 매핑을 구성할 수 없어 동기화를 중단했습니다.'
    ];
    if (missing.length) parts.push('누락 헤더: ' + missing.join(', '));
    if (ambiguous.length) parts.push('중복·모호 헤더: ' + ambiguous.join(' | '));
    throw new Error(parts.join('\n'));
  }

  return {
    columnByField: columnByField,
    indexByField: indexByField,
    rawHeaderByField: rawHeaderByField,
    lastCol: targetSchema.lastCol,
    mappingSummary: ITMAINT_getTargetMappedColumnSummary_2026_({
      columnByField: columnByField,
      rawHeaderByField: rawHeaderByField
    })
  };
}


function ITMAINT_getTargetMappedColumnSummary_2026_(targetFieldMap) {
  var config = ITMAINT_getConfig_2026_();
  var definitions = config.targetFieldDefinitions || {};
  var columnByField = targetFieldMap && targetFieldMap.columnByField || {};
  var rawHeaderByField = targetFieldMap && targetFieldMap.rawHeaderByField || {};

  return Object.keys(columnByField)
    .map(function(fieldKey) {
      return {
        fieldKey: fieldKey,
        column: Number(columnByField[fieldKey]),
        label: String(definitions[fieldKey] && definitions[fieldKey].label || fieldKey),
        header: String(rawHeaderByField[fieldKey] || '')
      };
    })
    .sort(function(a, b) { return a.column - b.column; })
    .map(function(item) {
      return ITMAINT_columnToLetter_2026_(item.column) + ':' + item.header;
    })
    .join(', ');
}


function ITMAINT_previewTargetHeaderMapping_2026() {
  var config = ITMAINT_getConfig_2026_();
  var targetSheet = ITMAINT_getTargetSheet_2026_();
  var targetSchema = ITMAINT_buildSchema_2026_(
    targetSheet,
    config.targetHeaderRow,
    [],
    '2026정보통신유지보수'
  );
  var targetFieldMap = ITMAINT_buildTargetFieldMap_2026_(targetSchema);
  var result = {
    status: 'SUCCESS',
    spreadsheetId: config.targetSpreadsheetId,
    sheetName: config.targetSheetName,
    headerRow: config.targetHeaderRow,
    lastColumn: targetSchema.lastCol,
    mapping: targetFieldMap.mappingSummary,
    columnByField: targetFieldMap.columnByField,
    version: config.version
  };

  Logger.log('[ITMAINT_previewTargetHeaderMapping_2026] ' + JSON.stringify(result));
  return result;
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
function ITMAINT_getTargetIdMap_2026_(targetSheet, targetFieldMap) {
  var config = ITMAINT_getConfig_2026_();
  var fieldMap = targetFieldMap || ITMAINT_validateTargetLayout_2026_(
    ITMAINT_buildSchema_2026_(
      targetSheet,
      config.targetHeaderRow,
      [],
      '2026정보통신유지보수'
    )
  );
  var contractNoColumn = Number(fieldMap.columnByField.contractNo);
  var lastRow = targetSheet.getLastRow();
  var idMap = {};

  if (lastRow < config.targetStartRow) return idMap;

  var idValues = targetSheet
    .getRange(
      config.targetStartRow,
      contractNoColumn,
      lastRow - config.targetStartRow + 1,
      1
    )
    .getValues();

  idValues.forEach(function(row, index) {
    var id = ITMAINT_normalizeId_2026_(row[0]);

    if (id && !idMap[id]) {
      idMap[id] = config.targetStartRow + index;
    }
  });

  return idMap;
}


function ITMAINT_getFirstEmptyTargetRow_2026_(targetSheet, targetFieldMap) {
  var config = ITMAINT_getConfig_2026_();
  var fieldMap = targetFieldMap || ITMAINT_validateTargetLayout_2026_(
    ITMAINT_buildSchema_2026_(
      targetSheet,
      config.targetHeaderRow,
      [],
      '2026정보통신유지보수'
    )
  );
  var contractNoColumn = Number(fieldMap.columnByField.contractNo);
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
    .getRange(config.targetStartRow, contractNoColumn, rowCount, 1)
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
 * 논리 필드 객체를 현재 대상 헤더 위치에 기록한다.
 * 열 삽입·이동과 무관하며, 서로 붙어 있는 대상 열만 묶어 setValues한다.
 */
function ITMAINT_writeTargetRecordByHeader_2026_(
  targetSheet,
  rowNumber,
  targetRecord,
  targetFieldMap
) {
  var entries = Object.keys(targetRecord || {})
    .filter(function(fieldKey) {
      return targetFieldMap && targetFieldMap.columnByField[fieldKey];
    })
    .map(function(fieldKey) {
      return {
        fieldKey: fieldKey,
        column: Number(targetFieldMap.columnByField[fieldKey]),
        value: targetRecord[fieldKey]
      };
    })
    .sort(function(a, b) { return a.column - b.column; });

  if (!entries.length) return { writeOperations: 0, writtenCells: 0, segments: [] };

  var groups = [];
  entries.forEach(function(entry) {
    var lastGroup = groups.length ? groups[groups.length - 1] : null;
    if (!lastGroup || entry.column !== lastGroup.endColumn + 1) {
      groups.push({
        startColumn: entry.column,
        endColumn: entry.column,
        values: [entry.value],
        fields: [entry.fieldKey]
      });
      return;
    }
    lastGroup.endColumn = entry.column;
    lastGroup.values.push(entry.value);
    lastGroup.fields.push(entry.fieldKey);
  });

  groups.forEach(function(group) {
    targetSheet
      .getRange(rowNumber, group.startColumn, 1, group.values.length)
      .setValues([group.values]);
  });

  return {
    writeOperations: groups.length,
    writtenCells: entries.length,
    segments: groups.map(function(group) {
      return ITMAINT_columnToLetter_2026_(group.startColumn) + ':' +
        ITMAINT_columnToLetter_2026_(group.endColumn);
    })
  };
}


/**
 * 다건 호환용 래퍼.
 */
function ITMAINT_writeTargetRowsWritableColumns_2026_(
  targetSheet,
  startRow,
  records,
  targetFieldMap
) {
  if (!records || records.length === 0) return;
  var config = ITMAINT_getConfig_2026_();
  var fieldMap = targetFieldMap || ITMAINT_validateTargetLayout_2026_(
    ITMAINT_buildSchema_2026_(
      targetSheet,
      config.targetHeaderRow,
      [],
      '2026정보통신유지보수'
    )
  );

  records.forEach(function(record, index) {
    ITMAINT_writeTargetRecordByHeader_2026_(
      targetSheet,
      startRow + index,
      record,
      fieldMap
    );
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
