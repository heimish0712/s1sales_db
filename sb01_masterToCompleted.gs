/****************************************************
 * 수주확정/계약완료(A) ↔ 마스터시트(신규)(B)
 *
 * 기본 방향:
 *   B시트 마스터 → A시트 수주확정/계약완료
 *
 * 예외:
 *   - A시트 1~158행 기존값 → B시트로 1회 역연동 가능
 *   - A시트 E열 ↔ B시트 AR열, A시트 G열 ↔ B시트 AQ열 실시간 상호연동
 *
 * 핵심 원칙:
 *   - 평상시 기준 데이터는 B시트, 즉 마스터시트
 *   - A시트 B열 고객번호를 B시트 A열 고객번호에서 찾음
 *   - 열번호가 바뀌어도 헤더명 기준으로 찾음
 *   - 헤더는 1행, 2행 둘 다 검사
 ****************************************************/

const CONTRACT_MASTER_SYNC = {
  targetSheetName: "수주확정/계약완료", // A시트
  sourceSheetName: "마스터시트(신규)",   // B시트

  // 헤더가 1행 또는 2행에 있다고 했으니 둘 다 검사
  headerRows: [1, 2],

  // 실제 데이터 시작 행
  dataStartRow: 3,

  // A시트 B열에 고객번호 입력 시 자동으로 B에서 불러오기 시작할 행
  // 1~158행은 이미 값이 있으므로 고객번호 입력에 따른 자동 덮어쓰기 방지
  autoPullStartRow: 159,

  // 마스터시트가 수정될 때는 A시트 기존 1~158행도 반영해야 하므로 3행부터 반영
  masterChangeReflectStartRow: 3,

  // A시트 기존값을 B시트로 1회 역연동할 범위
  oneTimeReverse: {
    startRow: 3,
    endRow: 158,

    // false면 A시트가 빈칸인 값은 B시트를 빈칸으로 덮어쓰지 않음
    // 진짜로 빈칸까지 밀어버리고 싶으면 true로 바꾸면 됨
    writeBlanksToMaster: false
  },

  // A시트 고객번호
  targetId: {
    headers: ["고객번호"],
    fallbackLetter: "B"
  },

  // B시트 고객번호
  sourceId: {
    headers: ["고객번호"],
    fallbackLetter: "A"
  },

  // A시트에서 실시간 역연동을 허용할 열
  // 여기 지정된 열은 A → B, B → A 둘 다 반영됨.
  // 수주 E열 ↔ 마스터 AR열, 수주 G열 ↔ 마스터 AQ열만 명시 연동함.
  liveBidirectionalTargetLetters: ["E", "G"],
  liveBidirectionalWriteBlanks: true,

  fields: [
    {
      name: "수주 E열 ↔ 마스터 AR열",
      type: "direct",
      bidirectional: true,
      valueMode: "raw",
      reverseValueMode: "raw",
      target: { headers: [], fallbackLetter: "E" },
      source: { headers: [], fallbackLetter: "AR" }
    },
    {
      name: "수주 G열 ↔ 마스터 AQ열",
      type: "direct",
      bidirectional: true,
      valueMode: "raw",
      reverseValueMode: "raw",
      target: { headers: [], fallbackLetter: "G" },
      source: { headers: [], fallbackLetter: "AQ" }
    },
    {
      name: "지역",
      type: "direct",
      target: { headers: ["지역"], fallbackLetter: "H" },
      source: { headers: ["지역구분"], fallbackLetter: "D" }
    },
    {
      name: "제보자",
      type: "direct",
      target: { headers: ["제보자"], fallbackLetter: "I" },
      source: { headers: ["제보자"], fallbackLetter: "BG" }
    },
    {
      name: "계약담당자",
      type: "direct",
      target: { headers: ["계약담당자"], fallbackLetter: "J" },
      source: { headers: ["영업담당자"], fallbackLetter: "F" }
    },
    {
      name: "고객사명",
      type: "direct",
      target: { headers: ["고객사명"], fallbackLetter: "K" },
      source: { headers: ["회사명"], fallbackLetter: "G" }
    },
    {
      name: "담당자 이름",
      type: "direct",
      target: { headers: ["담당자 이름", "담당자이름"], fallbackLetter: "L" },
      source: { headers: ["고객사 담당자", "고객사담당자"], fallbackLetter: "J" }
    },
    {
      name: "전화번호",
      type: "direct",
      valueMode: "display",
      reverseValueMode: "display",
      target: { headers: ["전화번호"], fallbackLetter: "M" },
      source: { headers: ["직통번호"], fallbackLetter: "L" }
    },
    {
      name: "이메일 주소",
      type: "direct",
      target: { headers: ["이메일 주소", "이메일주소"], fallbackLetter: "N" },
      source: { headers: ["담당자 이메일 주소", "담당자이메일주소"], fallbackLetter: "M" }
    },
    {
      name: "연면적",
      type: "direct",
      target: { headers: ["연면적"], fallbackLetter: "O" },
      source: { headers: ["연면적"], fallbackLetter: "N" }
    },
    {
      name: "선임 유형",
      type: "direct",
      target: { headers: ["선임 유형", "선임유형"], fallbackLetter: "P" },
      source: { headers: ["관리등급"], fallbackLetter: "O" }
    },
    {
      name: "계약가",
      type: "direct",
      target: { headers: ["계약가"], fallbackLetter: "Q" },
      source: { headers: ["최종 견적가", "최종견적가"], fallbackLetter: "X" }
    },
    {
      name: "VAT",
      type: "direct",
      target: { headers: ["VAT", "부가세"], fallbackLetter: "R" },
      source: { headers: ["부가세"], fallbackLetter: "Z" }
    },
    {
      name: "수행사",
      type: "direct",
      target: { headers: ["수행사"], fallbackLetter: "S" },
      source: { headers: ["수행사"], fallbackLetter: "E" }
    },
    {
      name: "사업자등록번호",
      type: "direct",
      valueMode: "display",
      reverseValueMode: "display",
      target: { headers: ["사업자등록번호"], fallbackLetter: "T" },
      source: { headers: ["사업자등록번호"], fallbackLetter: "AT" }
    },
    {
      name: "대표자명",
      type: "direct",
      target: { headers: ["대표자명"], fallbackLetter: "U" },
      source: { headers: ["대표자명"], fallbackLetter: "AU" }
    },
    {
      name: "업태",
      type: "direct",
      target: { headers: ["업태"], fallbackLetter: "V" },
      source: { headers: ["업태"], fallbackLetter: "AV" }
    },
    {
      name: "종목",
      type: "direct",
      target: { headers: ["종목"], fallbackLetter: "W" },
      source: { headers: ["종목"], fallbackLetter: "AW" }
    },
    {
      name: "고객사 주소",
      type: "direct",
      target: { headers: ["고객사 주소", "고객사주소"], fallbackLetter: "X" },
      source: { headers: ["고객사 상세 주소", "고객사상세주소"], fallbackLetter: "I" }
    },

    // Y열: 계약시작일 ~ 계약종료일
    {
      name: "계약기간",
      type: "period",
      target: { headers: ["계약기간"], fallbackLetter: "Y" },
      sourceStart: { headers: ["계약시작일"], fallbackLetter: "R" },
      sourceEnd: { headers: ["계약종료일"], fallbackLetter: "S" }
    },

    // Z열: 관리자 선임 여부가 "선임"일 때만 계약단위에서 숫자만 추출
    {
      name: "비상주 선임",
      type: "conditionalExtractNumber",
      target: { headers: ["비상주 선임", "비상주선임"], fallbackLetter: "Z" },
      conditionSource: { headers: ["관리자 선임 여부", "관리자선임여부"], fallbackLetter: "U" },
      valueSource: { headers: ["계약단위"], fallbackLetter: "T" },
      conditionText: "선임",
      suffixForReverse: "개월"
    },

    // AA열: 유지점검 n회 → n
    {
      name: "유지점검",
      type: "extractNumber",
      target: { headers: ["유지점검"], fallbackLetter: "AA" },
      source: { headers: ["유지점검"], fallbackLetter: "V" },
      suffixForReverse: "회"
    },

    // AB열: 성능점검 n회 → n
    {
      name: "성능점검",
      type: "extractNumber",
      target: { headers: ["성능점검"], fallbackLetter: "AB" },
      source: { headers: ["성능점검"], fallbackLetter: "W" },
      suffixForReverse: "회"
    },
    {
      name: "청구 등 메모",
      type: "direct",
      target: { headers: ["청구 등 메모", "청구등메모"], fallbackLetter: "AC" },
      source: { headers: ["계약 사항 관련 메모", "계약사항관련메모"], fallbackLetter: "AX" }
    },
    {
      name: "선임예정일",
      type: "direct",
      target: { headers: ["선임예정일"], fallbackLetter: "AD" },
      source: { headers: ["선임예정일"], fallbackLetter: "BA" }
    },
    {
      name: "유지점검예정일",
      type: "direct",
      target: { headers: ["유지점검예정일"], fallbackLetter: "AE" },
      source: { headers: ["유지점검예정일"], fallbackLetter: "BB" }
    },
    {
      name: "성능점검예정일",
      type: "direct",
      target: { headers: ["성능점검예정일"], fallbackLetter: "AF" },
      source: { headers: ["성능점검예정일"], fallbackLetter: "BC" }
    },
    {
      name: "선임완료여부",
      type: "direct",
      target: { headers: ["선임완료여부"], fallbackLetter: "AG" },
      source: { headers: ["선임완료여부"], fallbackLetter: "BD" }
    },
    {
      name: "유지점검완료여부",
      type: "direct",
      target: { headers: ["유지점검완료여부"], fallbackLetter: "AH" },
      source: { headers: ["유지점검완료"], fallbackLetter: "BE" }
    },
    {
      name: "성능점검완료여부",
      type: "direct",
      target: { headers: ["성능점검완료여부"], fallbackLetter: "AI" },
      source: { headers: ["성능점검완료"], fallbackLetter: "BF" }
    }
  ]
};


/****************************************************
 * 1. 최초 1회 실행: 트리거 설치
 ****************************************************/
function installContractMasterSyncTrigger() {
  const ss = SpreadsheetApp.getActive();
  const handlerName = "handleContractMasterSyncOnEdit";

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (trigger.getHandlerFunction() === handlerName) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handlerName)
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  ss.toast("수주확정/계약완료 ↔ 마스터시트 동기화 트리거 설치 완료", "설치 완료", 5);
}


/****************************************************
 * 2. 설치형 onEdit 트리거 핸들러
 ****************************************************/
function handleContractMasterSyncOnEdit(e) {
  if (!e || !e.range || !e.source) return;

  const lock = LockService.getDocumentLock();

  if (!lock.tryLock(5000)) return;

  try {
    const ss = e.source;
    const range = e.range;
    const editedSheet = range.getSheet();
    const editedSheetName = editedSheet.getName();

    // A시트 보조 기능 먼저 처리
    // B열 날짜 입력, E/F/G/K 색상 처리
    handleTargetSheetExtraFeatures_(e);

    const ctx = buildContractMasterSyncContext_(ss);

    // A시트에서 고객번호 입력 시: 159행 이후만 B → A 자동 조회
    // A시트 E/G열 수정 시: 같은 고객번호를 가진 B시트 명시 열로 즉시 역반영
    if (editedSheetName === CONTRACT_MASTER_SYNC.targetSheetName) {
      const firstRow = Math.max(range.getRow(), CONTRACT_MASTER_SYNC.dataStartRow);
      const lastRow = range.getLastRow();

      if (rangeIntersectsColumn_(range, ctx.targetIdCol)) {
        for (let row = firstRow; row <= lastRow; row++) {
          if (row < CONTRACT_MASTER_SYNC.autoPullStartRow) continue;

          pullOneTargetRowFromMaster_(ctx, row, true);
        }
      }

      const affectedBidirectionalFields = getAffectedBidirectionalTargetFields_(ctx, range);

      if (affectedBidirectionalFields.length) {
        let pushed = 0;
        let skipped = 0;

        for (let row = firstRow; row <= lastRow; row++) {
          const ok = pushOneTargetRowBidirectionalFieldsToMaster_(ctx, row, affectedBidirectionalFields);
          if (ok) pushed++;
          else skipped++;
        }

        if (pushed > 0) {
          SpreadsheetApp.getActive().toast(
            `E/G 상호연동 완료: ${pushed}행 반영${skipped ? `, ${skipped}행 스킵` : ""}`,
            "수주→마스터 반영",
            3
          );
        }
      }

      return;
    }

    // B시트, 즉 마스터시트가 수정되면: 같은 고객번호를 가진 A시트 행 전체 갱신
    if (editedSheetName === CONTRACT_MASTER_SYNC.sourceSheetName) {
      const affectedSourceCols = getAffectedSourceColumns_(ctx);
      const relevant = affectedSourceCols.some(col => rangeIntersectsColumn_(range, col));

      if (!relevant) return;

      const firstRow = Math.max(range.getRow(), CONTRACT_MASTER_SYNC.dataStartRow);
      const lastRow = range.getLastRow();

      for (let sourceRow = firstRow; sourceRow <= lastRow; sourceRow++) {
        reflectOneMasterRowToAllTargetRows_(ctx, sourceRow);
      }

      return;
    }

  } catch (err) {
    console.error(err);
    SpreadsheetApp.getActive().toast("동기화 오류: " + err.message, "오류", 8);
  } finally {
    lock.releaseLock();
  }
}


/****************************************************
 * 3. A시트 1~158행 기존값 → B시트로 1회 역연동
 *
 * 이 함수는 반드시 필요할 때 딱 1번만 실행.
 * 평소에는 실행하지 마. 데이터가 다시 뒤엉킨다.
 ****************************************************/
function oneTimePushA1To158ToMaster() {
  const ss = SpreadsheetApp.getActive();
  const ctx = buildContractMasterSyncContext_(ss);

  const startRow = CONTRACT_MASTER_SYNC.oneTimeReverse.startRow;
  const endRow = Math.min(
    CONTRACT_MASTER_SYNC.oneTimeReverse.endRow,
    ctx.targetSheet.getLastRow()
  );

  let updated = 0;
  let skippedNoId = 0;
  let notFound = 0;

  for (let targetRow = startRow; targetRow <= endRow; targetRow++) {
    const idValue = getCellDisplay_(ctx.targetSheet, targetRow, ctx.targetIdCol);

    if (!idValue) {
      skippedNoId++;
      continue;
    }

    const sourceRow = findSourceRowById_(ctx, idValue);

    if (!sourceRow) {
      notFound++;
      console.log(`마스터시트에서 고객번호를 찾지 못함: A시트 ${targetRow}행 / 고객번호 ${idValue}`);
      continue;
    }

    pushOneTargetRowToMaster_(ctx, targetRow, sourceRow);
    updated++;
  }

  ss.toast(
    `A→B 1회 역연동 완료: ${updated}행 반영, 고객번호 없음 ${skippedNoId}행, 마스터 미발견 ${notFound}행`,
    "역연동 완료",
    8
  );
}


/****************************************************
 * 4. 선택 실행: A시트 전체를 마스터 기준으로 강제 갱신
 *
 * 주의:
 * 이 함수는 A시트 3행부터 끝까지 전부 B시트 기준으로 덮어씀.
 * 1~158행을 보존하고 싶으면 함부로 실행하지 마.
 ****************************************************/
function forceSyncAllTargetRowsFromMaster() {
  const ss = SpreadsheetApp.getActive();
  const ctx = buildContractMasterSyncContext_(ss);

  const startRow = CONTRACT_MASTER_SYNC.dataStartRow;
  const lastRow = ctx.targetSheet.getLastRow();

  let updated = 0;
  let skipped = 0;

  for (let targetRow = startRow; targetRow <= lastRow; targetRow++) {
    const ok = pullOneTargetRowFromMaster_(ctx, targetRow, false);
    if (ok) updated++;
    else skipped++;
  }

  ss.toast(`전체 강제 동기화 완료: ${updated}행 반영, ${skipped}행 스킵`, "동기화 완료", 8);
}


/****************************************************
 * A시트 특정 행을 고객번호 기준으로 B시트에서 가져와 반영
 ****************************************************/
function pullOneTargetRowFromMaster_(ctx, targetRow, showToast) {
  const idValue = getCellDisplay_(ctx.targetSheet, targetRow, ctx.targetIdCol);

  if (!idValue) return false;

  const sourceRow = findSourceRowById_(ctx, idValue);

  if (!sourceRow) {
    if (showToast) {
      SpreadsheetApp.getActive().toast(
        `마스터시트에서 고객번호 [${idValue}]를 찾지 못했습니다.`,
        "조회 실패",
        5
      );
    }
    return false;
  }

  writeMasterRowToTargetRow_(ctx, sourceRow, targetRow);
  refreshTargetStatusColorsIfNeeded_(ctx.targetSheet, targetRow, targetRow);

  if (showToast) {
    SpreadsheetApp.getActive().toast(
      `고객번호 [${idValue}] 정보 불러오기 완료`,
      "동기화 완료",
      3
    );
  }

  return true;
}


/****************************************************
 * B시트 특정 행 변경 → A시트 같은 고객번호 행 전체 반영
 ****************************************************/
function reflectOneMasterRowToAllTargetRows_(ctx, sourceRow) {
  const idValue = getCellDisplay_(ctx.sourceSheet, sourceRow, ctx.sourceIdCol);

  if (!idValue) return;

  const targetRows = findTargetRowsById_(ctx, idValue);

  targetRows.forEach(targetRow => {
    if (targetRow < CONTRACT_MASTER_SYNC.masterChangeReflectStartRow) return;
    writeMasterRowToTargetRow_(ctx, sourceRow, targetRow);
    refreshTargetStatusColorsIfNeeded_(ctx.targetSheet, targetRow, targetRow);
  });
}


/****************************************************
 * A시트 E/G 수정 → B시트 AR/AQ 열로 역반영
 ****************************************************/
function getAffectedBidirectionalTargetFields_(ctx, range) {
  return ctx.resolvedFields.filter(field => {
    return field.bidirectional && rangeIntersectsColumn_(range, field.targetCol);
  });
}

function pushOneTargetRowBidirectionalFieldsToMaster_(ctx, targetRow, fields) {
  if (!fields || !fields.length) return false;

  const idValue = getCellDisplay_(ctx.targetSheet, targetRow, ctx.targetIdCol);

  if (!idValue) return false;

  const sourceRow = findSourceRowById_(ctx, idValue);

  if (!sourceRow) {
    console.log(`마스터시트에서 고객번호를 찾지 못함: A시트 ${targetRow}행 / 고객번호 ${idValue}`);
    return false;
  }

  const targetLastCol = Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol);

  const raw = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getValues()[0];

  const display = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getDisplayValues()[0];

  fields.forEach(field => {
    if (field.type !== "direct") {
      throw new Error(`실시간 상호연동은 direct 타입만 지원합니다: ${field.name}`);
    }

    const value = getByMode_(
      raw,
      display,
      field.targetCol,
      field.reverseValueMode || field.valueMode || "raw"
    );

    if (!CONTRACT_MASTER_SYNC.liveBidirectionalWriteBlanks && isBlank_(value)) return;

    ctx.sourceSheet.getRange(sourceRow, field.sourceCol).setValue(value);
  });

  return true;
}

/****************************************************
 * B시트 sourceRow 값을 A시트 targetRow에 씀
 ****************************************************/
function writeMasterRowToTargetRow_(ctx, sourceRow, targetRow) {
  const sourceLastCol = Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol);

  const raw = ctx.sourceSheet
    .getRange(sourceRow, 1, 1, sourceLastCol)
    .getValues()[0];

  const display = ctx.sourceSheet
    .getRange(sourceRow, 1, 1, sourceLastCol)
    .getDisplayValues()[0];

  ctx.resolvedFields.forEach(field => {
    let value = "";

    if (field.type === "direct") {
      value = getByMode_(raw, display, field.sourceCol, field.valueMode || "raw");

    } else if (field.type === "period") {
      const start = getByMode_(raw, display, field.sourceStartCol, "display");
      const end = getByMode_(raw, display, field.sourceEndCol, "display");
      value = makePeriodText_(start, end);

    } else if (field.type === "conditionalExtractNumber") {
      const conditionValue = String(display[field.conditionSourceCol - 1] || "").trim();

      if (conditionValue === field.conditionText) {
        value = extractFirstNumber_(display[field.valueSourceCol - 1]);
      } else {
        value = "";
      }

    } else if (field.type === "extractNumber") {
      value = extractFirstNumber_(display[field.sourceCol - 1]);
    }

    ctx.targetSheet.getRange(targetRow, field.targetCol).setValue(value);
  });
}


/****************************************************
 * A시트 targetRow 값을 B시트 sourceRow에 씀
 * 1회 역연동 전용
 ****************************************************/
function pushOneTargetRowToMaster_(ctx, targetRow, sourceRow) {
  const targetLastCol = Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol);

  const raw = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getValues()[0];

  const display = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getDisplayValues()[0];

  ctx.resolvedFields.forEach(field => {
    if (field.type === "direct") {
      const value = getByMode_(raw, display, field.targetCol, field.reverseValueMode || field.valueMode || "raw");
      setSourceIfAllowed_(ctx.sourceSheet, sourceRow, field.sourceCol, value);
      return;
    }

    if (field.type === "period") {
      const periodText = String(display[field.targetCol - 1] || "").trim();

      if (!periodText && !CONTRACT_MASTER_SYNC.oneTimeReverse.writeBlanksToMaster) return;

      const parsed = splitPeriodText_(periodText);

      if (parsed.start !== null) {
        setSourceIfAllowed_(ctx.sourceSheet, sourceRow, field.sourceStartCol, parsed.start);
      }

      if (parsed.end !== null) {
        setSourceIfAllowed_(ctx.sourceSheet, sourceRow, field.sourceEndCol, parsed.end);
      }

      return;
    }

    if (field.type === "conditionalExtractNumber") {
      const n = extractFirstNumber_(display[field.targetCol - 1]);

      if ((n === "" || n === null) && !CONTRACT_MASTER_SYNC.oneTimeReverse.writeBlanksToMaster) return;

      if (n !== "" && n !== null) {
        ctx.sourceSheet.getRange(sourceRow, field.conditionSourceCol).setValue(field.conditionText);
        ctx.sourceSheet.getRange(sourceRow, field.valueSourceCol).setValue(String(n) + (field.suffixForReverse || ""));
      } else {
        ctx.sourceSheet.getRange(sourceRow, field.conditionSourceCol).setValue("");
        ctx.sourceSheet.getRange(sourceRow, field.valueSourceCol).setValue("");
      }

      return;
    }

    if (field.type === "extractNumber") {
      const n = extractFirstNumber_(display[field.targetCol - 1]);

      if ((n === "" || n === null) && !CONTRACT_MASTER_SYNC.oneTimeReverse.writeBlanksToMaster) return;

      const value = n === "" || n === null
        ? ""
        : String(n) + (field.suffixForReverse || "");

      setSourceIfAllowed_(ctx.sourceSheet, sourceRow, field.sourceCol, value);
      return;
    }
  });
}


/****************************************************
 * 컨텍스트 구성: 시트, 열 위치, 필드 매핑 해석
 ****************************************************/
function buildContractMasterSyncContext_(ss) {
  const targetSheet = ss.getSheetByName(CONTRACT_MASTER_SYNC.targetSheetName);
  const sourceSheet = ss.getSheetByName(CONTRACT_MASTER_SYNC.sourceSheetName);

  if (!targetSheet) {
    throw new Error(`A시트를 찾을 수 없습니다: ${CONTRACT_MASTER_SYNC.targetSheetName}`);
  }

  if (!sourceSheet) {
    throw new Error(`B시트를 찾을 수 없습니다: ${CONTRACT_MASTER_SYNC.sourceSheetName}`);
  }

  const targetIdCol = resolveColumn_(targetSheet, CONTRACT_MASTER_SYNC.targetId);
  const sourceIdCol = resolveColumn_(sourceSheet, CONTRACT_MASTER_SYNC.sourceId);

  const resolvedFields = CONTRACT_MASTER_SYNC.fields.map(field => {
    const targetCol = resolveColumn_(targetSheet, field.target);
    const resolved = {
      name: field.name,
      type: field.type,
      valueMode: field.valueMode || "raw",
      reverseValueMode: field.reverseValueMode || null,
      conditionText: field.conditionText || null,
      suffixForReverse: field.suffixForReverse || "",
      bidirectional: field.bidirectional === true,
      targetCol: targetCol
    };

    if (field.type === "direct" || field.type === "extractNumber") {
      if (field.source && field.source.headersFromTarget) {
        const sourceHeaders = getHeaderCandidatesFromColumn_(targetSheet, targetCol);

        if (!sourceHeaders.length) {
          throw new Error(
            `${targetSheet.getName()} ${columnNumberToLetter_(targetCol)}열의 헤더가 비어 있어 마스터시트 상호연동 열을 찾을 수 없습니다.`
          );
        }

        resolved.sourceCol = resolveColumn_(sourceSheet, {
          headers: sourceHeaders,
          fallbackLetter: field.source.fallbackLetter || null
        });
      } else {
        resolved.sourceCol = resolveColumn_(sourceSheet, field.source);
      }
    }

    if (field.type === "period") {
      resolved.sourceStartCol = resolveColumn_(sourceSheet, field.sourceStart);
      resolved.sourceEndCol = resolveColumn_(sourceSheet, field.sourceEnd);
    }

    if (field.type === "conditionalExtractNumber") {
      resolved.conditionSourceCol = resolveColumn_(sourceSheet, field.conditionSource);
      resolved.valueSourceCol = resolveColumn_(sourceSheet, field.valueSource);
    }

    return resolved;
  });

  const sourceCols = [sourceIdCol];
  const targetCols = [targetIdCol];

  resolvedFields.forEach(field => {
    targetCols.push(field.targetCol);

    if (field.sourceCol) sourceCols.push(field.sourceCol);
    if (field.sourceStartCol) sourceCols.push(field.sourceStartCol);
    if (field.sourceEndCol) sourceCols.push(field.sourceEndCol);
    if (field.conditionSourceCol) sourceCols.push(field.conditionSourceCol);
    if (field.valueSourceCol) sourceCols.push(field.valueSourceCol);
  });

  return {
    targetSheet,
    sourceSheet,
    targetIdCol,
    sourceIdCol,
    resolvedFields,
    maxSourceCol: Math.max.apply(null, sourceCols),
    maxTargetCol: Math.max.apply(null, targetCols)
  };
}


/****************************************************
 * B시트에서 고객번호 찾기
 ****************************************************/
function findSourceRowById_(ctx, idValue) {
  const normalizedId = normalizeId_(idValue);
  const lastRow = ctx.sourceSheet.getLastRow();

  if (lastRow < CONTRACT_MASTER_SYNC.dataStartRow) return null;

  const values = ctx.sourceSheet
    .getRange(
      CONTRACT_MASTER_SYNC.dataStartRow,
      ctx.sourceIdCol,
      lastRow - CONTRACT_MASTER_SYNC.dataStartRow + 1,
      1
    )
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (normalizeId_(values[i][0]) === normalizedId) {
      return CONTRACT_MASTER_SYNC.dataStartRow + i;
    }
  }

  return null;
}


/****************************************************
 * A시트에서 같은 고객번호 행들 찾기
 ****************************************************/
function findTargetRowsById_(ctx, idValue) {
  const normalizedId = normalizeId_(idValue);
  const startRow = CONTRACT_MASTER_SYNC.masterChangeReflectStartRow;
  const lastRow = ctx.targetSheet.getLastRow();

  if (lastRow < startRow) return [];

  const values = ctx.targetSheet
    .getRange(startRow, ctx.targetIdCol, lastRow - startRow + 1, 1)
    .getDisplayValues();

  const rows = [];

  for (let i = 0; i < values.length; i++) {
    if (normalizeId_(values[i][0]) === normalizedId) {
      rows.push(startRow + i);
    }
  }

  return rows;
}


/****************************************************
 * 마스터시트에서 수정 감지해야 하는 원본 열 목록
 ****************************************************/
function getAffectedSourceColumns_(ctx) {
  const cols = new Set();

  cols.add(ctx.sourceIdCol);

  ctx.resolvedFields.forEach(field => {
    if (field.sourceCol) cols.add(field.sourceCol);
    if (field.sourceStartCol) cols.add(field.sourceStartCol);
    if (field.sourceEndCol) cols.add(field.sourceEndCol);
    if (field.conditionSourceCol) cols.add(field.conditionSourceCol);
    if (field.valueSourceCol) cols.add(field.valueSourceCol);
  });

  return Array.from(cols);
}


/****************************************************
 * 헤더명 기준 열 찾기
 * 못 찾으면 fallbackLetter 사용
 ****************************************************/
function resolveColumn_(sheet, columnSpec) {
  const lastCol = Math.max(sheet.getLastColumn(), columnLetterToNumber_(columnSpec.fallbackLetter || "A"));
  const headerMap = {};

  CONTRACT_MASTER_SYNC.headerRows.forEach(rowNum => {
    const headers = sheet.getRange(rowNum, 1, 1, lastCol).getDisplayValues()[0];

    headers.forEach((header, index) => {
      const key = normalizeHeader_(header);

      if (key && !headerMap[key]) {
        headerMap[key] = index + 1;
      }
    });
  });

  const candidates = columnSpec.headers || [];

  for (const candidate of candidates) {
    const key = normalizeHeader_(candidate);

    if (headerMap[key]) {
      return headerMap[key];
    }
  }

  if (columnSpec.fallbackLetter) {
    return columnLetterToNumber_(columnSpec.fallbackLetter);
  }

  throw new Error(
    `${sheet.getName()} 시트에서 열을 찾지 못했습니다. 후보 헤더: ${candidates.join(", ")}`
  );
}

function getHeaderCandidatesFromColumn_(sheet, col) {
  const candidates = [];
  const seen = {};

  CONTRACT_MASTER_SYNC.headerRows.forEach(rowNum => {
    const value = sheet.getRange(rowNum, col).getDisplayValue();
    const text = String(value || "").trim();
    const key = normalizeHeader_(text);

    if (text && key && !seen[key]) {
      candidates.push(text);
      seen[key] = true;
    }
  });

  return candidates;
}

function columnNumberToLetter_(column) {
  let temp = Number(column);
  let letter = "";

  while (temp > 0) {
    const mod = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }

  return letter;
}


/****************************************************
 * 값 처리 유틸
 ****************************************************/
function getByMode_(raw, display, col, mode) {
  if (mode === "display") {
    return display[col - 1];
  }

  return raw[col - 1];
}


function getCellDisplay_(sheet, row, col) {
  return String(sheet.getRange(row, col).getDisplayValue() || "").trim();
}


function setSourceIfAllowed_(sheet, row, col, value) {
  if (!CONTRACT_MASTER_SYNC.oneTimeReverse.writeBlanksToMaster && isBlank_(value)) {
    return;
  }

  sheet.getRange(row, col).setValue(value);
}


function isBlank_(value) {
  return value === "" || value === null || typeof value === "undefined";
}


function makePeriodText_(start, end) {
  const s = String(start || "").trim();
  const e = String(end || "").trim();

  if (s && e) return `${s} ~ ${e}`;
  if (s) return s;
  if (e) return e;

  return "";
}


function splitPeriodText_(text) {
  const value = String(text || "").trim();

  if (!value) {
    return { start: "", end: "" };
  }

  const parts = value.split(/\s*~\s*/);

  if (parts.length >= 2) {
    return {
      start: parts[0].trim(),
      end: parts.slice(1).join("~").trim()
    };
  }

  // "~"가 없으면 계약기간을 안전하게 쪼갤 수 없으니 시작일만 넣음
  return {
    start: value,
    end: null
  };
}


function extractFirstNumber_(value) {
  const text = String(value || "").replace(/,/g, "").trim();
  const match = text.match(/-?\d+(\.\d+)?/);

  if (!match) return "";

  const n = Number(match[0]);

  return Number.isNaN(n) ? match[0] : n;
}


function normalizeId_(value) {
  return String(value || "")
    .trim()
    .replace(/\s+/g, "");
}


function normalizeHeader_(value) {
  return String(value || "")
    .replace(/\uFEFF/g, "")
    .replace(/\u00A0/g, " ")
    .trim()
    .replace(/\s+/g, "")
    .toLowerCase();
}


function rangeIntersectsColumn_(range, col) {
  return col >= range.getColumn() && col <= range.getLastColumn();
}


function columnLetterToNumber_(letter) {
  let column = 0;
  const upper = String(letter || "").toUpperCase().trim();

  for (let i = 0; i < upper.length; i++) {
    column = column * 26 + upper.charCodeAt(i) - 64;
  }

  return column;
}


/****************************************************
 * 디버그용: A159 고객번호 기준으로 테스트
 ****************************************************/
function debugContractMasterSyncA159() {
  const ss = SpreadsheetApp.getActive();
  const ctx = buildContractMasterSyncContext_(ss);

  const testRow = 159;
  const idValue = getCellDisplay_(ctx.targetSheet, testRow, ctx.targetIdCol);
  const sourceRow = idValue ? findSourceRowById_(ctx, idValue) : null;

  console.log("A시트:", ctx.targetSheet.getName());
  console.log("B시트:", ctx.sourceSheet.getName());
  console.log("A시트 고객번호 열:", ctx.targetIdCol);
  console.log("B시트 고객번호 열:", ctx.sourceIdCol);
  console.log("A159 고객번호:", idValue);
  console.log("B시트에서 찾은 행:", sourceRow);

  if (!idValue) {
    ss.toast("A159에 고객번호가 없습니다.", "디버그", 5);
    return;
  }

  if (!sourceRow) {
    ss.toast(`마스터시트에서 고객번호 [${idValue}]를 찾지 못했습니다.`, "디버그", 8);
    return;
  }

  writeMasterRowToTargetRow_(ctx, sourceRow, testRow);
  ss.toast("A159 기준 테스트 동기화 완료", "디버그 완료", 5);
}

function oneTimePushA1To158ToMaster_FAST() {
  const ss = SpreadsheetApp.getActive();
  const ctx = buildContractMasterSyncContext_(ss);

  const startRow = CONTRACT_MASTER_SYNC.oneTimeReverse.startRow;
  const endRow = Math.min(
    CONTRACT_MASTER_SYNC.oneTimeReverse.endRow,
    ctx.targetSheet.getLastRow()
  );

  const targetRowCount = endRow - startRow + 1;

  if (targetRowCount <= 0) {
    ss.toast("역연동할 A시트 행이 없습니다.", "역연동", 5);
    return;
  }

  const sourceStartRow = CONTRACT_MASTER_SYNC.dataStartRow;
  const sourceLastRow = ctx.sourceSheet.getLastRow();
  const sourceRowCount = sourceLastRow - sourceStartRow + 1;

  if (sourceRowCount <= 0) {
    ss.toast("마스터시트에 데이터가 없습니다.", "역연동", 5);
    return;
  }

  // 1. A시트 3~158행을 한 번에 읽기
  const targetLastCol = Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol);
  const targetRaw = ctx.targetSheet
    .getRange(startRow, 1, targetRowCount, targetLastCol)
    .getValues();

  const targetDisplay = ctx.targetSheet
    .getRange(startRow, 1, targetRowCount, targetLastCol)
    .getDisplayValues();

  // 2. B시트 고객번호 열을 한 번에 읽고, 고객번호 → B시트 행번호 맵 만들기
  const sourceIds = ctx.sourceSheet
    .getRange(sourceStartRow, ctx.sourceIdCol, sourceRowCount, 1)
    .getDisplayValues();

  const sourceRowById = new Map();

  sourceIds.forEach((row, index) => {
    const id = normalizeId_(row[0]);
    if (id && !sourceRowById.has(id)) {
      sourceRowById.set(id, sourceStartRow + index);
    }
  });

  // 3. 수정해야 하는 B시트 열 목록 수집
  const sourceCols = collectReverseWritableSourceColumns_(ctx);

  // 4. B시트의 필요한 열만 한 번씩 읽기
  const sourceColumnData = {};

  sourceCols.forEach(col => {
    const range = ctx.sourceSheet.getRange(sourceStartRow, col, sourceRowCount, 1);

    sourceColumnData[col] = {
      values: range.getValues(),
      formulas: range.getFormulas()
    };
  });

  let updated = 0;
  let skippedNoId = 0;
  let notFound = 0;

  // 5. 메모리 안에서만 값 변경
  for (let i = 0; i < targetRowCount; i++) {
    const idValue = targetDisplay[i][ctx.targetIdCol - 1];
    const normalizedId = normalizeId_(idValue);

    if (!normalizedId) {
      skippedNoId++;
      continue;
    }

    const sourceRow = sourceRowById.get(normalizedId);

    if (!sourceRow) {
      notFound++;
      console.log(`마스터시트에서 고객번호를 찾지 못함: A시트 ${startRow + i}행 / 고객번호 ${idValue}`);
      continue;
    }

    const sourceIndex = sourceRow - sourceStartRow;

    ctx.resolvedFields.forEach(field => {
      applyReverseFieldToColumnData_(
        field,
        targetRaw[i],
        targetDisplay[i],
        sourceColumnData,
        sourceIndex
      );
    });

    updated++;
  }

  // 6. 바뀐 B시트 열만 한 번씩 쓰기
  sourceCols.forEach(col => {
    const data = sourceColumnData[col];

    // 기존 수식은 수식 문자열로 복원
    const output = data.values.map((row, i) => {
      const formula = data.formulas[i][0];
      return [formula ? formula : row[0]];
    });

    ctx.sourceSheet
      .getRange(sourceStartRow, col, sourceRowCount, 1)
      .setValues(output);
  });

  ss.toast(
    `빠른 역연동 완료: ${updated}행 반영, 고객번호 없음 ${skippedNoId}행, 마스터 미발견 ${notFound}행`,
    "A→B 역연동 완료",
    8
  );
}


function collectReverseWritableSourceColumns_(ctx) {
  const cols = new Set();

  ctx.resolvedFields.forEach(field => {
    if (field.type === "direct") {
      cols.add(field.sourceCol);
    }

    if (field.type === "period") {
      cols.add(field.sourceStartCol);
      cols.add(field.sourceEndCol);
    }

    if (field.type === "conditionalExtractNumber") {
      cols.add(field.conditionSourceCol);
      cols.add(field.valueSourceCol);
    }

    if (field.type === "extractNumber") {
      cols.add(field.sourceCol);
    }
  });

  return Array.from(cols).sort((a, b) => a - b);
}


function applyReverseFieldToColumnData_(field, targetRawRow, targetDisplayRow, sourceColumnData, sourceIndex) {
  const writeBlanks = CONTRACT_MASTER_SYNC.oneTimeReverse.writeBlanksToMaster;

  if (field.type === "direct") {
    const value = getByMode_(
      targetRawRow,
      targetDisplayRow,
      field.targetCol,
      field.reverseValueMode || field.valueMode || "raw"
    );

    setColumnDataIfAllowed_(sourceColumnData, field.sourceCol, sourceIndex, value, writeBlanks);
    return;
  }

  if (field.type === "period") {
    const periodText = String(targetDisplayRow[field.targetCol - 1] || "").trim();

    if (!periodText && !writeBlanks) return;

    const parsed = splitPeriodText_(periodText);

    if (parsed.start !== null) {
      setColumnDataIfAllowed_(sourceColumnData, field.sourceStartCol, sourceIndex, parsed.start, writeBlanks);
    }

    if (parsed.end !== null) {
      setColumnDataIfAllowed_(sourceColumnData, field.sourceEndCol, sourceIndex, parsed.end, writeBlanks);
    }

    return;
  }

  if (field.type === "conditionalExtractNumber") {
    const n = extractFirstNumber_(targetDisplayRow[field.targetCol - 1]);

    if ((n === "" || n === null) && !writeBlanks) return;

    if (n !== "" && n !== null) {
      setColumnDataIfAllowed_(sourceColumnData, field.conditionSourceCol, sourceIndex, field.conditionText, writeBlanks);
      setColumnDataIfAllowed_(
        sourceColumnData,
        field.valueSourceCol,
        sourceIndex,
        String(n) + (field.suffixForReverse || ""),
        writeBlanks
      );
    } else {
      setColumnDataIfAllowed_(sourceColumnData, field.conditionSourceCol, sourceIndex, "", writeBlanks);
      setColumnDataIfAllowed_(sourceColumnData, field.valueSourceCol, sourceIndex, "", writeBlanks);
    }

    return;
  }

  if (field.type === "extractNumber") {
    const n = extractFirstNumber_(targetDisplayRow[field.targetCol - 1]);

    if ((n === "" || n === null) && !writeBlanks) return;

    const value = n === "" || n === null
      ? ""
      : String(n) + (field.suffixForReverse || "");

    setColumnDataIfAllowed_(sourceColumnData, field.sourceCol, sourceIndex, value, writeBlanks);
    return;
  }
}


function setColumnDataIfAllowed_(sourceColumnData, col, sourceIndex, value, writeBlanks) {
  if (!writeBlanks && isBlank_(value)) return;

  if (!sourceColumnData[col]) {
    throw new Error(`sourceColumnData에 ${col}열 데이터가 없습니다.`);
  }

  sourceColumnData[col].values[sourceIndex][0] = value;
}

/****************************************************
 * A시트 보조 기능
 *
 * 1. B열 고객번호 입력 시 C열에 입력일 자동 기재
 * 2. G열이 "저장"이 아니면 G/K 연분홍색
 * 3. E열이 "저장"이 아니면 E 연노란색
 * 4. F열이 "저장"이 아니면 F 연노란색
 ****************************************************/

const TARGET_SHEET_EXTRA_CONFIG = {
  sheetName: "수주확정/계약완료",

  headerRows: [1, 2],
  firstDataRow: 3,

  idCol: 2,       // B열
  dateCol: 3,     // C열

  eCol: 5,        // E열
  fCol: 6,        // F열
  gCol: 7,        // G열
  kCol: 11,       // K열

  savedText: "저장",

  colors: {
    pink: "#FCE4EC",   // 연한 분홍색
    yellow: "#FFF9C4", // 연한 노란색
    white: "#FFFFFF"
  }
};


function refreshTargetStatusColorsIfNeeded_(sheet, firstRow, lastRow) {
  if (!sheet || sheet.getName() !== TARGET_SHEET_EXTRA_CONFIG.sheetName) return;

  const safeFirstRow = Math.max(firstRow, TARGET_SHEET_EXTRA_CONFIG.firstDataRow);
  const safeLastRow = Math.max(lastRow, safeFirstRow);

  applyStatusColorsForRows_(sheet, safeFirstRow, safeLastRow);
}

/**
 * A시트 보조 기능 onEdit 처리
 */
function handleTargetSheetExtraFeatures_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();

  if (sheet.getName() !== TARGET_SHEET_EXTRA_CONFIG.sheetName) return;

  const range = e.range;
  const firstRow = Math.max(range.getRow(), TARGET_SHEET_EXTRA_CONFIG.firstDataRow);
  const lastRow = range.getLastRow();

  if (lastRow < TARGET_SHEET_EXTRA_CONFIG.firstDataRow) return;

  // 1. B열 입력 시 C열 날짜 자동 입력
  if (rangeIntersectsColumn_(range, TARGET_SHEET_EXTRA_CONFIG.idCol)) {
    applyInputDateForIdColumn_(sheet, firstRow, lastRow);
  }

  // 2~4. E/F/G/K 색상 갱신
  const needColorRefresh =
    rangeIntersectsColumn_(range, TARGET_SHEET_EXTRA_CONFIG.eCol) ||
    rangeIntersectsColumn_(range, TARGET_SHEET_EXTRA_CONFIG.fCol) ||
    rangeIntersectsColumn_(range, TARGET_SHEET_EXTRA_CONFIG.gCol) ||
    rangeIntersectsColumn_(range, TARGET_SHEET_EXTRA_CONFIG.kCol);

  if (needColorRefresh) {
    applyStatusColorsForRows_(sheet, firstRow, lastRow);
  }
}


/**
 * B열에 숫자값이 입력되면 C열에 오늘 날짜 입력
 *
 * - B열이 비어 있으면 C열도 비움
 * - B열이 숫자 또는 숫자로만 된 텍스트이면 날짜 입력
 * - 이미 C열에 날짜가 있더라도 B열을 다시 수정하면 오늘 날짜로 갱신
 */
function applyInputDateForIdColumn_(sheet, firstRow, lastRow) {
  const rowCount = lastRow - firstRow + 1;

  const idValues = sheet
    .getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.idCol, rowCount, 1)
    .getDisplayValues();

  const todayText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy. MM. dd."
  );

  const dateValues = idValues.map(row => {
    const value = String(row[0] || "").trim();

    if (!value) {
      return [""];
    }

    // 쉼표 제거 후 숫자 여부 판단: 1,234도 숫자로 봄
    const normalized = value.replace(/,/g, "");

    if (/^\d+(\.\d+)?$/.test(normalized)) {
      return [todayText];
    }

    // 숫자가 아니면 C열은 건드리지 않기 위해 현재값 유지가 필요함
    // 다만 setValues 구조상 기존값을 다시 읽어서 넣음
    return [null];
  });

  const dateRange = sheet.getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.dateCol, rowCount, 1);
  const currentDates = dateRange.getValues();

  const output = dateValues.map((row, i) => {
    return row[0] === null ? [currentDates[i][0]] : row;
  });

  dateRange.setValues(output);
}


/**
 * E/F/G/K 색상 처리
 *
 * G열 != "저장"이면 G열과 K열 연분홍색
 * G열 == "저장"이면 G열과 K열 흰색
 *
 * E열 != "저장"이면 E열 연노란색
 * E열 == "저장"이면 E열 흰색
 *
 * F열 != "저장"이면 F열 연노란색
 * F열 == "저장"이면 F열 흰색
 */
function applyStatusColorsForRows_(sheet, firstRow, lastRow) {
  const rowCount = lastRow - firstRow + 1;

  const eValues = sheet
    .getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.eCol, rowCount, 1)
    .getDisplayValues();

  const fValues = sheet
    .getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.fCol, rowCount, 1)
    .getDisplayValues();

  const gValues = sheet
    .getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.gCol, rowCount, 1)
    .getDisplayValues();

  const kValues = sheet
    .getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.kCol, rowCount, 1)
    .getDisplayValues();

  const yellow = TARGET_SHEET_EXTRA_CONFIG.colors.yellow;
  const pink = TARGET_SHEET_EXTRA_CONFIG.colors.pink;
  const white = TARGET_SHEET_EXTRA_CONFIG.colors.white;
  const savedText = TARGET_SHEET_EXTRA_CONFIG.savedText;

  const eBackgrounds = [];
  const fBackgrounds = [];
  const gBackgrounds = [];
  const kBackgrounds = [];

  for (let i = 0; i < rowCount; i++) {
    const e = String(eValues[i][0] || "").trim();
    const f = String(fValues[i][0] || "").trim();
    const g = String(gValues[i][0] || "").trim();
    const k = String(kValues[i][0] || "").trim();

    // K열이 완전 공란이면 모든 색상 규칙 적용 안 함
    // 기존 색도 흰색으로 제거
    if (k === "") {
      eBackgrounds.push([white]);
      fBackgrounds.push([white]);
      gBackgrounds.push([white]);
      kBackgrounds.push([white]);
      continue;
    }

    // K열에 값이 있을 때만 E/F/G/K 색상 규칙 적용
    eBackgrounds.push([e === savedText ? white : yellow]);
    fBackgrounds.push([f === savedText ? white : yellow]);

    const gColor = g === savedText ? white : pink;
    gBackgrounds.push([gColor]);
    kBackgrounds.push([gColor]);
  }

  sheet.getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.eCol, rowCount, 1)
    .setBackgrounds(eBackgrounds);

  sheet.getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.fCol, rowCount, 1)
    .setBackgrounds(fBackgrounds);

  sheet.getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.gCol, rowCount, 1)
    .setBackgrounds(gBackgrounds);

  sheet.getRange(firstRow, TARGET_SHEET_EXTRA_CONFIG.kCol, rowCount, 1)
    .setBackgrounds(kBackgrounds);
}


/**
 * 기존 행 전체 색상 한 번 정리하고 싶을 때 수동 실행
 */
function refreshAllTargetSheetStatusColors() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName(TARGET_SHEET_EXTRA_CONFIG.sheetName);

  if (!sheet) {
    throw new Error("시트를 찾을 수 없습니다: " + TARGET_SHEET_EXTRA_CONFIG.sheetName);
  }

  const firstRow = TARGET_SHEET_EXTRA_CONFIG.firstDataRow;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) {
    SpreadsheetApp.getActive().toast("색상 정리할 데이터가 없습니다.", "색상 정리", 5);
    return;
  }

  applyStatusColorsForRows_(sheet, firstRow, lastRow);

  SpreadsheetApp.getActive().toast("E/F/G/K열 색상 전체 정리 완료", "색상 정리 완료", 5);
}