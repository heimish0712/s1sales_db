/****************************************************
 * 수주확정/계약완료(A) ↔ 마스터시트(신규)(B)
 *
 * 기본 방향:
 *   마스터시트(신규) → 수주확정/계약완료
 *
 * 예외:
 *   - 수주확정/계약완료의 "사업자등록증 저장", "계약서 저장"은
 *     수주확정/계약완료 값을 최우선으로 사용한다.
 *   - 위 두 상태값은 수주확정/계약완료 → 마스터시트 방향으로 보정한다.
 *
 * 핵심 원칙:
 *   - 모든 열은 실제 헤더명으로 찾는다.
 *   - A/B/E/G/AR/AQ 같은 열주소 fallback은 사용하지 않는다.
 *   - 헤더 위치와 데이터 시작 행도 실제 시트에서 자동 탐지한다.
 *   - 필수 헤더가 없거나 중복되면 잘못된 열에 쓰지 않고 즉시 중단한다.
 ****************************************************/

const CONTRACT_MASTER_SYNC = {
  targetSheetName: "수주확정/계약완료",
  targetSheetNames: ["수주확정/계약완료", "수주확정계약완료"],
  sourceSheetName: "마스터시트(신규)",
  sourceSheetNames: ["마스터시트(신규)"],

  // 첨부 파일 기준: 수주확정/계약완료는 1행, 마스터는 2행이 헤더다.
  // 실제 운영 시트에서는 첫 10행을 검사해 헤더 행을 자동 탐지한다.
  headerSearchMaxRows: 10,
  targetHeaderRequired: [
    ["계약번호"],
    ["고객번호"],
    ["고객사명"]
  ],
  sourceHeaderRequired: [
    ["고객번호"],
    ["현재 영업 진행 상황", "현재영업진행상황"],
    ["회사명"]
  ],

  // 고객번호 입력에 따른 신규행 자동 조회는 기존 운영 규칙을 유지한다.
  autoPullStartRow: 159,

  oneTimeReverse: {
    endRow: 158,
    writeBlanksToMaster: false
  },

  targetId: { headers: ["고객번호"] },
  sourceId: { headers: ["고객번호"] },

  // 수주확정/계약완료에서 상태값을 빈칸으로 지운 경우도 유효한 수정이다.
  liveBidirectionalWriteBlanks: true,

  fields: [
    {
      name: "사업자등록증 저장",
      type: "direct",
      bidirectional: true,
      targetAuthority: true,
      valueMode: "raw",
      reverseValueMode: "raw",
      target: { headers: ["사업자등록증 저장", "사업자등록증저장"] },
      source: { headers: ["사업자등록증 저장", "사업자등록증저장"] }
    },
    {
      name: "계약서 저장",
      type: "direct",
      bidirectional: true,
      targetAuthority: true,
      valueMode: "raw",
      reverseValueMode: "raw",
      target: { headers: ["계약서 저장", "계약서저장"] },
      source: { headers: ["계약서 저장", "계약서저장"] }
    },
    {
      name: "지역",
      type: "direct",
      target: { headers: ["지역"] },
      source: { headers: ["지역구분", "지역 구분"] }
    },
    {
      name: "제보자",
      type: "direct",
      target: { headers: ["제보자"] },
      source: { headers: ["제보자"] }
    },
    {
      name: "계약담당자",
      type: "direct",
      target: { headers: ["계약담당자", "계약 담당자"] },
      source: { headers: ["영업담당자", "영업 담당자"] }
    },
    {
      name: "고객사명",
      type: "direct",
      target: { headers: ["고객사명", "고객사 명"] },
      source: { headers: ["회사명", "회사 명"] }
    },
    {
      name: "담당자 성함",
      type: "direct",
      target: { headers: ["담당자 성함", "담당자성함", "담당자 이름", "담당자이름"] },
      source: { headers: ["고객사 담당자", "고객사담당자"] }
    },
    {
      name: "전화번호",
      type: "direct",
      valueMode: "display",
      reverseValueMode: "display",
      target: { headers: ["전화번호", "전화 번호"] },
      source: { headers: ["직통번호", "직통 번호"] }
    },
    {
      name: "이메일 주소",
      type: "direct",
      target: { headers: ["이메일 주소", "이메일주소"] },
      source: { headers: ["담당자 이메일 주소", "담당자이메일주소"] }
    },
    {
      name: "연면적",
      type: "direct",
      target: { headers: ["연면적"] },
      source: { headers: ["연면적"] }
    },
    {
      name: "선임유형",
      type: "direct",
      target: { headers: ["선임유형", "선임 유형"] },
      source: { headers: ["관리등급", "관리 등급"] }
    },
    {
      name: "계약가",
      type: "direct",
      target: { headers: ["계약가", "계약 가"] },
      source: { headers: ["최종 견적가", "최종견적가"] }
    },
    {
      name: "VAT",
      type: "direct",
      target: { headers: ["VAT", "부가세"] },
      source: { headers: ["부가세"] }
    },
    {
      name: "수행사",
      type: "direct",
      target: { headers: ["수행사"] },
      source: { headers: ["수행사"] }
    },
    {
      name: "사업자등록번호",
      type: "direct",
      valueMode: "display",
      reverseValueMode: "display",
      target: { headers: ["사업자등록번호", "사업자 등록번호"] },
      source: { headers: ["사업자등록번호", "사업자 등록번호"] }
    },
    {
      name: "대표자명",
      type: "direct",
      target: { headers: ["대표자명", "대표자 명"] },
      source: { headers: ["대표자명", "대표자 명"] }
    },
    {
      name: "업태",
      type: "direct",
      target: { headers: ["업태"] },
      source: { headers: ["업태"] }
    },
    {
      name: "종목",
      type: "direct",
      target: { headers: ["종목"] },
      source: { headers: ["종목"] }
    },
    {
      name: "고객사 주소",
      type: "direct",
      target: { headers: ["고객사 주소", "고객사주소"] },
      source: { headers: ["고객사 상세 주소", "고객사상세주소"] }
    },
    {
      name: "계약기간",
      type: "period",
      target: { headers: ["계약기간", "계약 기간"] },
      sourceStart: { headers: ["계약시작일", "계약 시작일"] },
      sourceEnd: { headers: ["계약종료일", "계약 종료일"] }
    },
    {
      name: "비상주선임",
      type: "conditionalExtractNumber",
      target: { headers: ["비상주선임", "비상주 선임"] },
      conditionSource: { headers: ["관리자 선임 여부", "관리자선임여부"] },
      valueSource: { headers: ["계약단위", "계약 단위"] },
      conditionText: "선임",
      suffixForReverse: "개월"
    },
    {
      name: "유지점검",
      type: "extractNumber",
      target: { headers: ["유지점검", "유지 점검"] },
      source: { headers: ["유지점검", "유지 점검"] },
      suffixForReverse: "회"
    },
    {
      name: "성능점검",
      type: "extractNumber",
      target: { headers: ["성능점검", "성능 점검"] },
      source: { headers: ["성능점검", "성능 점검"] },
      suffixForReverse: "회"
    },
    {
      name: "청구 등 메모",
      type: "direct",
      target: { headers: ["청구 등 메모", "청구등메모"] },
      source: { headers: ["계약 사항 관련 메모", "계약사항관련메모"] }
    },
    {
      name: "선임예정일",
      type: "direct",
      target: { headers: ["선임예정일", "선임 예정일"] },
      source: { headers: ["선임예정일", "선임 예정일"] }
    },
    {
      name: "유지점검예정일",
      type: "direct",
      target: { headers: ["유지점검예정일", "유지점검 예정일"] },
      source: { headers: ["유지점검예정일", "유지점검 예정일"] }
    },
    {
      name: "성능점검예정일",
      type: "direct",
      target: { headers: ["성능점검예정일", "성능점검 예정일"] },
      source: { headers: ["성능점검예정일", "성능점검 예정일"] }
    },
    {
      name: "선임완료여부",
      type: "direct",
      target: { headers: ["선임완료여부", "선임 완료여부"] },
      source: { headers: ["선임완료여부", "선임 완료여부"] }
    },
    {
      name: "유지점검완료여부",
      type: "direct",
      target: { headers: ["유지점검완료여부", "유지점검 완료여부"] },
      source: { headers: ["유지점검완료", "유지점검 완료"] }
    },
    {
      name: "성능점검완료여부",
      type: "direct",
      target: { headers: ["성능점검완료여부", "성능점검 완료여부"] },
      source: { headers: ["성능점검완료", "성능점검 완료"] }
    }
  ]
};


/****************************************************
 * 2. 설치형 onEdit 트리거 핸들러
 ****************************************************/
function handleContractMasterSyncOnEdit(e) {
  if (!e || !e.range || !e.source) {
    return { status: 'IGNORED_INVALID_EVENT' };
  }

  const ss = e.source;
  const range = e.range;
  const editedSheetName = range.getSheet().getName();

  const isTargetSheet = CMS28_matchesSheetName_(
    editedSheetName,
    CONTRACT_MASTER_SYNC.targetSheetNames
  );
  const isSourceSheet = CMS28_matchesSheetName_(
    editedSheetName,
    CONTRACT_MASTER_SYNC.sourceSheetNames
  );

  if (!isTargetSheet && !isSourceSheet) {
    return { status: 'IGNORED_UNRELATED_SHEET' };
  }

  return AUTOMATION_runEditHandlerWithLease_(
    'CONTRACT_SYNC',
    'handleContractMasterSyncOnEdit',
    e,
    function () {
      try {
        // A시트 보조 기능 먼저 처리
        // 헤더 기준 입력일·저장상태 색상 처리
        handleTargetSheetExtraFeatures_(e);

        const ctx = buildContractMasterSyncContext_(ss);

        // A시트에서 고객번호 입력 시: 159행 이후만 B → A 자동 조회
        // 수주확정의 권위 저장상태 수정 시: 같은 고객번호의 마스터 동일 헤더로 즉시 역반영
        if (isTargetSheet) {
          const firstRow = Math.max(range.getRow(), ctx.targetDataStartRow);
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
                `저장상태 상방연동 완료: ${pushed}행 반영${skipped ? `, ${skipped}행 스킵` : ""}`,
                "수주확정→마스터",
                3
              );
            }
          }

          return {
            route: 'TARGET',
            firstRow: firstRow,
            lastRow: lastRow
          };
        }

        // B시트, 즉 마스터시트가 수정되면: 같은 고객번호를 가진 A시트 행 전체 갱신
        if (isSourceSheet) {
          const affectedSourceCols = getAffectedSourceColumns_(ctx);
          const relevant = affectedSourceCols.some(col => rangeIntersectsColumn_(range, col));

          if (!relevant) {
            return { route: 'SOURCE', relevant: false };
          }

          const firstRow = Math.max(range.getRow(), ctx.sourceDataStartRow);
          const lastRow = range.getLastRow();

          for (let sourceRow = firstRow; sourceRow <= lastRow; sourceRow++) {
            reflectOneMasterRowToAllTargetRows_(ctx, sourceRow);
            CMS27_reconcileOneSourceRowFromTargetAuthority_(ctx, sourceRow);
          }

          return {
            route: 'SOURCE',
            relevant: true,
            firstRow: firstRow,
            lastRow: lastRow
          };
        }

        return { route: 'NONE' };
      } catch (err) {
        try {
          SpreadsheetApp.getActive().toast("동기화 오류: " + err.message, "오류", 8);
        } catch (ignoreToastError) {}
        throw err;
      }
    },
    { waitMs: 500, ttlMs: 8 * 60 * 1000 }
  );
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

  const startRow = ctx.targetDataStartRow;
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

  const writeResult = writeMasterRowToTargetRow_(ctx, sourceRow, targetRow);

  if (writeResult.changedCells > 0) {
    refreshTargetStatusColorsIfNeeded_(ctx.targetSheet, targetRow, targetRow);
  }

  if (showToast) {
    SpreadsheetApp.getActive().toast(
      writeResult.changedCells > 0
        ? `고객번호 [${idValue}] 변경값 ${writeResult.changedCells}개 반영`
        : `고객번호 [${idValue}] 변경사항 없음`,
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
    if (targetRow < ctx.targetDataStartRow) return;

    const writeResult = writeMasterRowToTargetRow_(ctx, sourceRow, targetRow);

    if (writeResult.changedCells > 0) {
      refreshTargetStatusColorsIfNeeded_(ctx.targetSheet, targetRow, targetRow);
    }
  });
}


/****************************************************
 * 수주확정 저장상태 수정 → 마스터 동일 헤더로 역반영
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
  const sourceLastCol = Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol);

  const raw = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getValues()[0];

  const display = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getDisplayValues()[0];

  const sourceCurrent = ctx.sourceSheet
    .getRange(sourceRow, 1, 1, sourceLastCol)
    .getValues()[0];

  const changedCells = [];

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

    if (!CMS19_valuesEqual_(sourceCurrent[field.sourceCol - 1], value)) {
      changedCells.push({
        col: field.sourceCol,
        value: value
      });
    }
  });

  CMS19_writeChangedRowCells_(ctx.sourceSheet, sourceRow, changedCells);
  return true;
}

/****************************************************
 * 27~28단계: 수주확정 저장상태 우선권
 ****************************************************/
function CMS27_isTargetAuthoritativeField_(field) {
  return !!(field && field.targetAuthority === true);
}

function CMS27_getTargetAuthoritativeFields_(ctx) {
  return ctx.resolvedFields.filter(function(field) {
    return CMS27_isTargetAuthoritativeField_(field);
  });
}

/**
 * 마스터의 대응 헤더가 직접 수정되더라도 수주확정 저장상태를 최종값으로 되돌린다.
 * 동일 고객번호가 여러 행이면 가장 아래쪽 행을 최신 행으로 본다.
 */
function CMS27_reconcileOneSourceRowFromTargetAuthority_(ctx, sourceRow) {
  const fields = CMS27_getTargetAuthoritativeFields_(ctx);
  if (!fields.length) return { status: 'NO_AUTHORITY_FIELDS', changedCells: 0 };

  const idValue = getCellDisplay_(ctx.sourceSheet, sourceRow, ctx.sourceIdCol);
  if (!idValue) return { status: 'SKIPPED_NO_ID', changedCells: 0 };

  const targetRows = findTargetRowsById_(ctx, idValue);
  if (!targetRows.length) return { status: 'SKIPPED_TARGET_NOT_FOUND', changedCells: 0 };

  const canonicalTargetRow = targetRows[targetRows.length - 1];
  const result = CMS27_pushTargetAuthorityFieldsToMaster_(
    ctx,
    canonicalTargetRow,
    sourceRow,
    fields
  );

  result.status = result.changedCells > 0 ? 'UPDATED_FROM_TARGET' : 'UNCHANGED';
  result.customerId = idValue;
  result.targetRow = canonicalTargetRow;
  result.duplicateTargetRows = Math.max(0, targetRows.length - 1);
  return result;
}

function CMS27_pushTargetAuthorityFieldsToMaster_(ctx, targetRow, sourceRow, fields) {
  const authorityFields = fields || CMS27_getTargetAuthoritativeFields_(ctx);
  if (!authorityFields.length) return { changedCells: 0, writeOperations: 0 };

  const targetLastCol = Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol);
  const sourceLastCol = Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol);

  const targetRaw = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getValues()[0];

  const targetDisplay = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getDisplayValues()[0];

  const sourceCurrent = ctx.sourceSheet
    .getRange(sourceRow, 1, 1, sourceLastCol)
    .getValues()[0];

  const changedCells = [];

  authorityFields.forEach(function(field) {
    if (field.type !== 'direct') {
      throw new Error('수주확정 우선 필드는 direct 타입만 지원합니다: ' + field.name);
    }

    const value = getByMode_(
      targetRaw,
      targetDisplay,
      field.targetCol,
      field.reverseValueMode || field.valueMode || 'raw'
    );

    // 권위 저장상태는 빈칸도 유효한 최종값이다. 수주확정에서 지우면 마스터도 지운다.
    if (!CMS19_valuesEqual_(sourceCurrent[field.sourceCol - 1], value)) {
      changedCells.push({ col: field.sourceCol, value: value });
    }
  });

  const writeResult = CMS19_writeChangedRowCells_(ctx.sourceSheet, sourceRow, changedCells);
  return {
    changedCells: changedCells.length,
    writeOperations: writeResult.writeOperations,
    changedColumns: changedCells.map(function(cell) { return cell.col; })
  };
}

/****************************************************
 * B시트 sourceRow 값을 A시트 targetRow에 씀
 ****************************************************/
function writeMasterRowToTargetRow_(ctx, sourceRow, targetRow) {
  const sourceLastCol = Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol);
  const targetLastCol = Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol);

  const raw = ctx.sourceSheet
    .getRange(sourceRow, 1, 1, sourceLastCol)
    .getValues()[0];

  const display = ctx.sourceSheet
    .getRange(sourceRow, 1, 1, sourceLastCol)
    .getDisplayValues()[0];

  const targetCurrent = ctx.targetSheet
    .getRange(targetRow, 1, 1, targetLastCol)
    .getValues()[0];

  const changedCells = [];

  ctx.resolvedFields.forEach(field => {
    // 수주확정 우선으로 지정된 저장상태는 마스터 값으로 절대 덮어쓰지 않는다.
    if (CMS27_isTargetAuthoritativeField_(field)) return;

    let value = "";

    if (field.type === "direct") {
      value = getByMode_(raw, display, field.sourceCol, field.valueMode || "raw");

      if (field.name === "지역") {
        value = CMS21_normalizeRegionForTargetOrSkip_(value);
        if (CMS21_isSkipWriteValue_(value)) return;
      }
    } else if (field.type === "period") {
      const start = getByMode_(raw, display, field.sourceStartCol, "display");
      const end = getByMode_(raw, display, field.sourceEndCol, "display");
      value = makePeriodText_(start, end);
    } else if (field.type === "conditionalExtractNumber") {
      const conditionValue = String(display[field.conditionSourceCol - 1] || "").trim();

      if (conditionValue === field.conditionText) {
        value = extractFirstNumber_(display[field.valueSourceCol - 1]);
      } else if (conditionValue === "비선임") {
        value = 0;
      } else {
        value = "";
      }
    } else if (field.type === "extractNumber") {
      value = extractFirstNumber_(display[field.sourceCol - 1]);
    }

    if (!CMS19_valuesEqual_(targetCurrent[field.targetCol - 1], value)) {
      changedCells.push({
        col: field.targetCol,
        value: value
      });
    }
  });

  const writeResult = CMS19_writeChangedRowCells_(
    ctx.targetSheet,
    targetRow,
    changedCells
  );

  return {
    changedCells: changedCells.length,
    writeOperations: writeResult.writeOperations,
    changedColumns: changedCells.map(function(cell) { return cell.col; })
  };
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
  const targetSheet = CMS28_findSheetByNames_(
    ss,
    CONTRACT_MASTER_SYNC.targetSheetNames
  );
  const sourceSheet = CMS28_findSheetByNames_(
    ss,
    CONTRACT_MASTER_SYNC.sourceSheetNames
  );

  if (!targetSheet) {
    throw new Error(
      `수주확정/계약완료 시트를 찾을 수 없습니다. 후보: ${CONTRACT_MASTER_SYNC.targetSheetNames.join(", ")}`
    );
  }

  if (!sourceSheet) {
    throw new Error(
      `마스터시트를 찾을 수 없습니다. 후보: ${CONTRACT_MASTER_SYNC.sourceSheetNames.join(", ")}`
    );
  }

  const targetHeaderRow = CMS28_detectHeaderRow_(
    targetSheet,
    CONTRACT_MASTER_SYNC.targetHeaderRequired,
    CONTRACT_MASTER_SYNC.headerSearchMaxRows
  );
  const sourceHeaderRow = CMS28_detectHeaderRow_(
    sourceSheet,
    CONTRACT_MASTER_SYNC.sourceHeaderRequired,
    CONTRACT_MASTER_SYNC.headerSearchMaxRows
  );

  const targetHeaderIndex = CMS28_buildHeaderIndex_(targetSheet, targetHeaderRow);
  const sourceHeaderIndex = CMS28_buildHeaderIndex_(sourceSheet, sourceHeaderRow);

  const targetIdCol = resolveColumn_(
    targetSheet,
    targetHeaderIndex,
    CONTRACT_MASTER_SYNC.targetId,
    '수주확정 고객번호'
  );
  const sourceIdCol = resolveColumn_(
    sourceSheet,
    sourceHeaderIndex,
    CONTRACT_MASTER_SYNC.sourceId,
    '마스터 고객번호'
  );

  const resolvedFields = CONTRACT_MASTER_SYNC.fields.map(function(field) {
    const targetCol = resolveColumn_(
      targetSheet,
      targetHeaderIndex,
      field.target,
      field.name + ' 대상열'
    );
    const resolved = {
      name: field.name,
      type: field.type,
      valueMode: field.valueMode || 'raw',
      reverseValueMode: field.reverseValueMode || null,
      conditionText: field.conditionText || null,
      suffixForReverse: field.suffixForReverse || '',
      bidirectional: field.bidirectional === true,
      targetAuthority: field.targetAuthority === true,
      targetCol: targetCol,
      targetHeader: CMS28_getHeaderTextAtColumn_(targetHeaderIndex, targetCol)
    };

    if (field.type === 'direct' || field.type === 'extractNumber') {
      if (field.source && field.source.headersFromTarget) {
        const sourceHeaders = getHeaderCandidatesFromColumn_(
          targetSheet,
          targetHeaderRow,
          targetCol
        );

        if (!sourceHeaders.length) {
          throw new Error(
            `${targetSheet.getName()}의 ${field.name} 대상 헤더가 비어 있어 마스터 열을 찾을 수 없습니다.`
          );
        }

        resolved.sourceCol = resolveColumn_(
          sourceSheet,
          sourceHeaderIndex,
          { headers: sourceHeaders },
          field.name + ' 원본열'
        );
      } else {
        resolved.sourceCol = resolveColumn_(
          sourceSheet,
          sourceHeaderIndex,
          field.source,
          field.name + ' 원본열'
        );
      }
      resolved.sourceHeader = CMS28_getHeaderTextAtColumn_(sourceHeaderIndex, resolved.sourceCol);
    }

    if (field.type === 'period') {
      resolved.sourceStartCol = resolveColumn_(
        sourceSheet,
        sourceHeaderIndex,
        field.sourceStart,
        field.name + ' 시작일 원본열'
      );
      resolved.sourceEndCol = resolveColumn_(
        sourceSheet,
        sourceHeaderIndex,
        field.sourceEnd,
        field.name + ' 종료일 원본열'
      );
      resolved.sourceStartHeader = CMS28_getHeaderTextAtColumn_(sourceHeaderIndex, resolved.sourceStartCol);
      resolved.sourceEndHeader = CMS28_getHeaderTextAtColumn_(sourceHeaderIndex, resolved.sourceEndCol);
    }

    if (field.type === 'conditionalExtractNumber') {
      resolved.conditionSourceCol = resolveColumn_(
        sourceSheet,
        sourceHeaderIndex,
        field.conditionSource,
        field.name + ' 조건 원본열'
      );
      resolved.valueSourceCol = resolveColumn_(
        sourceSheet,
        sourceHeaderIndex,
        field.valueSource,
        field.name + ' 값 원본열'
      );
      resolved.conditionSourceHeader = CMS28_getHeaderTextAtColumn_(sourceHeaderIndex, resolved.conditionSourceCol);
      resolved.valueSourceHeader = CMS28_getHeaderTextAtColumn_(sourceHeaderIndex, resolved.valueSourceCol);
    }

    return resolved;
  });

  const sourceCols = [sourceIdCol];
  const targetCols = [targetIdCol];

  resolvedFields.forEach(function(field) {
    targetCols.push(field.targetCol);

    if (field.sourceCol) sourceCols.push(field.sourceCol);
    if (field.sourceStartCol) sourceCols.push(field.sourceStartCol);
    if (field.sourceEndCol) sourceCols.push(field.sourceEndCol);
    if (field.conditionSourceCol) sourceCols.push(field.conditionSourceCol);
    if (field.valueSourceCol) sourceCols.push(field.valueSourceCol);
  });

  return {
    targetSheet: targetSheet,
    sourceSheet: sourceSheet,
    targetHeaderRow: targetHeaderRow,
    sourceHeaderRow: sourceHeaderRow,
    targetDataStartRow: targetHeaderRow + 1,
    sourceDataStartRow: sourceHeaderRow + 1,
    targetHeaderIndex: targetHeaderIndex,
    sourceHeaderIndex: sourceHeaderIndex,
    targetIdCol: targetIdCol,
    sourceIdCol: sourceIdCol,
    resolvedFields: resolvedFields,
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

  if (lastRow < ctx.sourceDataStartRow) return null;

  const values = ctx.sourceSheet
    .getRange(
      ctx.sourceDataStartRow,
      ctx.sourceIdCol,
      lastRow - ctx.sourceDataStartRow + 1,
      1
    )
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    if (normalizeId_(values[i][0]) === normalizedId) {
      return ctx.sourceDataStartRow + i;
    }
  }

  return null;
}


/****************************************************
 * A시트에서 같은 고객번호 행들 찾기
 ****************************************************/
function findTargetRowsById_(ctx, idValue) {
  const normalizedId = normalizeId_(idValue);
  const startRow = ctx.targetDataStartRow;
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
 * 28단계: 헤더 기반 시트·열 탐색
 ****************************************************/
function CMS28_normalizeSheetName_(value) {
  return String(value || '')
    .trim()
    .replace(/[\s\/\\]+/g, '')
    .toLowerCase();
}

function CMS28_matchesSheetName_(actualName, configuredNames) {
  const actualKey = CMS28_normalizeSheetName_(actualName);
  return (configuredNames || []).some(function(name) {
    return CMS28_normalizeSheetName_(name) === actualKey;
  });
}

function CMS28_findSheetByNames_(ss, configuredNames) {
  const names = configuredNames || [];

  for (let i = 0; i < names.length; i++) {
    const direct = ss.getSheetByName(names[i]);
    if (direct) return direct;
  }

  const expected = {};
  names.forEach(function(name) {
    expected[CMS28_normalizeSheetName_(name)] = true;
  });

  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    if (expected[CMS28_normalizeSheetName_(sheets[i].getName())]) {
      return sheets[i];
    }
  }

  return null;
}

function CMS28_detectHeaderRow_(sheet, requiredHeaderGroups, maxRows) {
  const scanRows = Math.min(
    Math.max(1, Number(maxRows) || 10),
    Math.max(1, sheet.getLastRow())
  );
  const lastCol = Math.max(1, sheet.getLastColumn());
  const values = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();
  let best = null;

  for (let rowIndex = 0; rowIndex < values.length; rowIndex++) {
    const keys = {};
    values[rowIndex].forEach(function(value) {
      const key = cmsNormalizeHeader_(value);
      if (key) keys[key] = true;
    });

    let matched = 0;
    (requiredHeaderGroups || []).forEach(function(group) {
      const found = (group || []).some(function(candidate) {
        return !!keys[cmsNormalizeHeader_(candidate)];
      });
      if (found) matched++;
    });

    if (!best || matched > best.matched) {
      best = { row: rowIndex + 1, matched: matched };
    }

    if (matched === (requiredHeaderGroups || []).length) {
      return rowIndex + 1;
    }
  }

  throw new Error(
    `${sheet.getName()} 시트의 헤더 행을 찾지 못했습니다. ` +
    `필수 헤더 그룹 ${JSON.stringify(requiredHeaderGroups)} / 최다 일치 ${best ? best.matched : 0}개`
  );
}

function CMS28_buildHeaderIndex_(sheet, headerRow) {
  const lastCol = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(headerRow, 1, 1, lastCol).getDisplayValues()[0];
  const byKey = {};
  const byColumn = {};

  headers.forEach(function(header, index) {
    const text = String(header || '').trim();
    const key = cmsNormalizeHeader_(text);
    if (!key) return;

    if (!byKey[key]) byKey[key] = [];
    byKey[key].push({ col: index + 1, text: text });
    byColumn[index + 1] = text;
  });

  return {
    headerRow: headerRow,
    byKey: byKey,
    byColumn: byColumn,
    headers: headers
  };
}

function resolveColumn_(sheet, headerIndex, columnSpec, fieldLabel) {
  const candidates = (columnSpec && columnSpec.headers) || [];
  const matches = {};

  candidates.forEach(function(candidate) {
    const key = cmsNormalizeHeader_(candidate);
    const entries = headerIndex.byKey[key] || [];
    entries.forEach(function(entry) {
      matches[entry.col] = entry.text;
    });
  });

  const columns = Object.keys(matches).map(Number).sort(function(a, b) { return a - b; });

  if (columns.length === 1) return columns[0];

  if (columns.length > 1) {
    throw new Error(
      `${sheet.getName()} 시트의 ${fieldLabel || '열'} 헤더가 중복 또는 모호합니다. ` +
      `후보=[${candidates.join(', ')}], 일치=[${columns.map(function(col) { return matches[col]; }).join(', ')}]`
    );
  }

  throw new Error(
    `${sheet.getName()} 시트에서 ${fieldLabel || '필수 열'}을 찾지 못했습니다. ` +
    `후보 헤더=[${candidates.join(', ')}], 헤더행=${headerIndex.headerRow}`
  );
}

function CMS28_getHeaderTextAtColumn_(headerIndex, col) {
  return String(headerIndex.byColumn[col] || '').trim();
}

function getHeaderCandidatesFromColumn_(sheet, headerRow, col) {
  const value = sheet.getRange(headerRow, col).getDisplayValue();
  const text = String(value || '').trim();
  return text ? [text] : [];
}

/**
 * 실제 쓰기 없이 현재 헤더 위치와 매핑을 확인한다.
 */
function CMS28_previewContractMasterHeaderMapping() {
  const ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  const ctx = buildContractMasterSyncContext_(ss);
  const mapping = ctx.resolvedFields.map(function(field) {
    const row = {
      field: field.name,
      type: field.type,
      targetHeader: field.targetHeader,
      targetColumn: field.targetCol,
      authority: field.targetAuthority === true
    };

    if (field.sourceCol) {
      row.sourceHeader = field.sourceHeader;
      row.sourceColumn = field.sourceCol;
    }
    if (field.sourceStartCol) {
      row.sourceStartHeader = field.sourceStartHeader;
      row.sourceStartColumn = field.sourceStartCol;
      row.sourceEndHeader = field.sourceEndHeader;
      row.sourceEndColumn = field.sourceEndCol;
    }
    if (field.conditionSourceCol) {
      row.conditionSourceHeader = field.conditionSourceHeader;
      row.conditionSourceColumn = field.conditionSourceCol;
      row.valueSourceHeader = field.valueSourceHeader;
      row.valueSourceColumn = field.valueSourceCol;
    }
    return row;
  });

  const result = {
    status: 'OK',
    targetSheet: ctx.targetSheet.getName(),
    targetHeaderRow: ctx.targetHeaderRow,
    targetDataStartRow: ctx.targetDataStartRow,
    sourceSheet: ctx.sourceSheet.getName(),
    sourceHeaderRow: ctx.sourceHeaderRow,
    sourceDataStartRow: ctx.sourceDataStartRow,
    targetIdHeader: CMS28_getHeaderTextAtColumn_(ctx.targetHeaderIndex, ctx.targetIdCol),
    sourceIdHeader: CMS28_getHeaderTextAtColumn_(ctx.sourceHeaderIndex, ctx.sourceIdCol),
    mapping: mapping
  };

  Logger.log(JSON.stringify(result, null, 2));
  return result;
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


function cmsNormalizeHeader_(value) {
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

  const startRow = ctx.targetDataStartRow;
  const endRow = Math.min(
    CONTRACT_MASTER_SYNC.oneTimeReverse.endRow,
    ctx.targetSheet.getLastRow()
  );

  const targetRowCount = endRow - startRow + 1;

  if (targetRowCount <= 0) {
    ss.toast("역연동할 A시트 행이 없습니다.", "역연동", 5);
    return;
  }

  const sourceStartRow = ctx.sourceDataStartRow;
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
 * 수주확정/계약완료 보조 기능 — 모두 헤더명 기준
 *
 * 1. 고객번호 입력 시 계약일자(발주번호 부여일)에 입력일 자동 기재
 * 2. 계약서 저장이 "저장"이 아니면 계약서 저장/고객사명 연분홍색
 * 3. 사업자등록증 저장이 "저장"이 아니면 연노란색
 * 4. 선임신고서 저장이 "저장"이 아니면 연노란색
 ****************************************************/

const TARGET_SHEET_EXTRA_CONFIG = {
  sheetNames: ["수주확정/계약완료", "수주확정계약완료"],
  headerRequired: [
    ["고객번호"],
    ["계약일자(발주번호 부여일)"],
    ["고객사명"]
  ],
  columns: {
    id: { headers: ["고객번호"] },
    inputDate: { headers: ["계약일자(발주번호 부여일)", "계약일자"] },
    businessRegistrationSaved: { headers: ["사업자등록증 저장", "사업자등록증저장"] },
    appointmentReportSaved: { headers: ["선임신고서 저장", "선임신고서저장"] },
    contractSaved: { headers: ["계약서 저장", "계약서저장"] },
    customerName: { headers: ["고객사명", "고객사 명"] }
  },
  savedText: "저장",
  colors: {
    pink: "#FCE4EC",
    yellow: "#FFF9C4",
    white: "#FFFFFF"
  }
};

function CMS28_resolveTargetExtraContext_(sheet) {
  if (!sheet || !CMS28_matchesSheetName_(sheet.getName(), TARGET_SHEET_EXTRA_CONFIG.sheetNames)) {
    return null;
  }

  const headerRow = CMS28_detectHeaderRow_(
    sheet,
    TARGET_SHEET_EXTRA_CONFIG.headerRequired,
    CONTRACT_MASTER_SYNC.headerSearchMaxRows
  );
  const headerIndex = CMS28_buildHeaderIndex_(sheet, headerRow);
  const specs = TARGET_SHEET_EXTRA_CONFIG.columns;

  return {
    headerRow: headerRow,
    dataStartRow: headerRow + 1,
    idCol: resolveColumn_(sheet, headerIndex, specs.id, '고객번호'),
    dateCol: resolveColumn_(sheet, headerIndex, specs.inputDate, '계약일자'),
    businessRegistrationSavedCol: resolveColumn_(
      sheet,
      headerIndex,
      specs.businessRegistrationSaved,
      '사업자등록증 저장'
    ),
    appointmentReportSavedCol: resolveColumn_(
      sheet,
      headerIndex,
      specs.appointmentReportSaved,
      '선임신고서 저장'
    ),
    contractSavedCol: resolveColumn_(sheet, headerIndex, specs.contractSaved, '계약서 저장'),
    customerNameCol: resolveColumn_(sheet, headerIndex, specs.customerName, '고객사명')
  };
}

function refreshTargetStatusColorsIfNeeded_(sheet, firstRow, lastRow) {
  const extra = CMS28_resolveTargetExtraContext_(sheet);
  if (!extra) return;

  const safeFirstRow = Math.max(firstRow, extra.dataStartRow);
  const safeLastRow = Math.max(lastRow, safeFirstRow);
  applyStatusColorsForRows_(sheet, safeFirstRow, safeLastRow, extra);
}

function handleTargetSheetExtraFeatures_(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const extra = CMS28_resolveTargetExtraContext_(sheet);
  if (!extra) return;

  const range = e.range;
  const firstRow = Math.max(range.getRow(), extra.dataStartRow);
  const lastRow = range.getLastRow();

  if (lastRow < extra.dataStartRow) return;

  if (rangeIntersectsColumn_(range, extra.idCol)) {
    applyInputDateForIdColumn_(sheet, firstRow, lastRow, extra);
  }

  const needColorRefresh =
    rangeIntersectsColumn_(range, extra.businessRegistrationSavedCol) ||
    rangeIntersectsColumn_(range, extra.appointmentReportSavedCol) ||
    rangeIntersectsColumn_(range, extra.contractSavedCol) ||
    rangeIntersectsColumn_(range, extra.customerNameCol);

  if (needColorRefresh) {
    applyStatusColorsForRows_(sheet, firstRow, lastRow, extra);
  }
}

function applyInputDateForIdColumn_(sheet, firstRow, lastRow, extra) {
  const rowCount = lastRow - firstRow + 1;
  const idValues = sheet
    .getRange(firstRow, extra.idCol, rowCount, 1)
    .getDisplayValues();

  const todayText = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    "yyyy. MM. dd."
  );

  const dateValues = idValues.map(function(row) {
    const value = String(row[0] || "").trim();
    if (!value) return [""];

    const normalized = value.replace(/,/g, "");
    if (/^\d+(\.\d+)?$/.test(normalized)) return [todayText];
    return [null];
  });

  const dateRange = sheet.getRange(firstRow, extra.dateCol, rowCount, 1);
  const currentDates = dateRange.getValues();
  const output = dateValues.map(function(row, i) {
    return row[0] === null ? [currentDates[i][0]] : row;
  });

  dateRange.setValues(output);
}

function applyStatusColorsForRows_(sheet, firstRow, lastRow, extra) {
  const rowCount = lastRow - firstRow + 1;
  const businessValues = sheet
    .getRange(firstRow, extra.businessRegistrationSavedCol, rowCount, 1)
    .getDisplayValues();
  const appointmentValues = sheet
    .getRange(firstRow, extra.appointmentReportSavedCol, rowCount, 1)
    .getDisplayValues();
  const contractValues = sheet
    .getRange(firstRow, extra.contractSavedCol, rowCount, 1)
    .getDisplayValues();
  const customerValues = sheet
    .getRange(firstRow, extra.customerNameCol, rowCount, 1)
    .getDisplayValues();

  const yellow = TARGET_SHEET_EXTRA_CONFIG.colors.yellow;
  const pink = TARGET_SHEET_EXTRA_CONFIG.colors.pink;
  const white = TARGET_SHEET_EXTRA_CONFIG.colors.white;
  const savedText = TARGET_SHEET_EXTRA_CONFIG.savedText;
  const businessBackgrounds = [];
  const appointmentBackgrounds = [];
  const contractBackgrounds = [];
  const customerBackgrounds = [];

  for (let i = 0; i < rowCount; i++) {
    const business = String(businessValues[i][0] || "").trim();
    const appointment = String(appointmentValues[i][0] || "").trim();
    const contract = String(contractValues[i][0] || "").trim();
    const customer = String(customerValues[i][0] || "").trim();

    if (customer === "") {
      businessBackgrounds.push([white]);
      appointmentBackgrounds.push([white]);
      contractBackgrounds.push([white]);
      customerBackgrounds.push([white]);
      continue;
    }

    businessBackgrounds.push([business === savedText ? white : yellow]);
    appointmentBackgrounds.push([appointment === savedText ? white : yellow]);

    const contractColor = contract === savedText ? white : pink;
    contractBackgrounds.push([contractColor]);
    customerBackgrounds.push([contractColor]);
  }

  sheet.getRange(firstRow, extra.businessRegistrationSavedCol, rowCount, 1)
    .setBackgrounds(businessBackgrounds);
  sheet.getRange(firstRow, extra.appointmentReportSavedCol, rowCount, 1)
    .setBackgrounds(appointmentBackgrounds);
  sheet.getRange(firstRow, extra.contractSavedCol, rowCount, 1)
    .setBackgrounds(contractBackgrounds);
  sheet.getRange(firstRow, extra.customerNameCol, rowCount, 1)
    .setBackgrounds(customerBackgrounds);
}

function refreshAllTargetSheetStatusColors() {
  const ss = SpreadsheetApp.getActive();
  const sheet = CMS28_findSheetByNames_(ss, TARGET_SHEET_EXTRA_CONFIG.sheetNames);

  if (!sheet) {
    throw new Error(
      "수주확정/계약완료 시트를 찾을 수 없습니다: " + TARGET_SHEET_EXTRA_CONFIG.sheetNames.join(", ")
    );
  }

  const extra = CMS28_resolveTargetExtraContext_(sheet);
  const firstRow = extra.dataStartRow;
  const lastRow = sheet.getLastRow();

  if (lastRow < firstRow) {
    SpreadsheetApp.getActive().toast("색상 정리할 데이터가 없습니다.", "색상 정리", 5);
    return;
  }

  applyStatusColorsForRows_(sheet, firstRow, lastRow, extra);
  SpreadsheetApp.getActive().toast(
    "저장상태 헤더 기준 색상 정리 완료",
    "색상 정리 완료",
    5
  );
}

function normalizeRegionGroupForTarget_(value) {
  const text = String(value || "").trim();

  if (!text) return "";

  if ([
    "강원권",
    "대구경북권",
    "부울경권",
    "수도권",
    "제주권",
    "충청권",
    "호남권"
  ].includes(text)) {
    return text;
  }

  if (/서울|경기|인천|수도/.test(text)) return "수도권";
  if (/강원/.test(text)) return "강원권";
  if (/대구|경북|경상북/.test(text)) return "대구경북권";
  if (/부산|울산|경남|경상남|부울경/.test(text)) return "부울경권";
  if (/충북|충남|충청|대전|세종/.test(text)) return "충청권";
  if (/전북|전남|전라|광주|호남/.test(text)) return "호남권";
  if (/제주/.test(text)) return "제주권";

  return text;
}

/****************************************************
 * 수정 시 동기화 + 5분마다 강제 동기화
 * 트리거는 AUTOMATION_executeCanonicalCutover()로 중앙 전환·설치
 ****************************************************/



/****************************************************
 * 5분마다 자동 실행되는 핸들러
 ****************************************************/
function handleContractMasterSyncEvery5Minutes() {
  try {
    return CMS_runFullSyncForAutomationPipeline_();
  } catch (err) {
    console.error(err);
    return null;
  }
}


/**
 * 중앙 핵심 동기화 파이프라인용 1단계 진입점.
 * 기존 시간 트리거와 달리 오류를 삼키지 않고 상위 파이프라인에 전달한다.
 */
function CMS_runFullSyncForAutomationPipeline_() {
  return AUTOMATION_runWithModuleLeaseOrThrow_(
    'CONTRACT_SYNC',
    'CMS_runFullSyncForAutomationPipeline_',
    function () {
      return syncAllTargetRowsFromMaster_FAST_();
    },
    { waitMs: 15000, ttlMs: 8 * 60 * 1000 }
  );
}


/****************************************************
 * A시트 전체 변경 감지 동기화
 * - 고객번호 맵을 한 번 만든 뒤 전체를 비교한다.
 * - 현재값과 다른 셀만 기록하며, 변경이 없으면 쓰기 작업을 하지 않는다.
 ****************************************************/
function forceSyncAllTargetRowsFromMaster() {
  const ss = SpreadsheetApp.getActive();
  const result = syncAllTargetRowsFromMaster_FAST_();

  CMS5_safeToast_(
    ss,
    `변경 감지 동기화 완료: 일반 변경행 ${result.changedRows}행 / 일반 변경셀 ${result.changedCells}개 / 저장상태→마스터 ${result.targetAuthority.changedRows}행 ${result.targetAuthority.changedCells}셀 / 변경없음 ${result.unchangedRows}행 / 고객번호 없음 ${result.skippedNoId}행 / 마스터 미발견 ${result.notFound}행`,
    "변경 감지 동기화",
    8
  );
}


/****************************************************
 * 마스터시트 기준 A시트 전체 비교·변경분 동기화
 ****************************************************/
function syncAllTargetRowsFromMaster_FAST_() {
  const ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  const ctx = buildContractMasterSyncContext_(ss);

  // 권위 저장상태는 먼저 수주확정 → 마스터 방향으로 충돌을 해소한다.
  const targetAuthorityResult = CMS27_reconcileAllTargetAuthorityToMaster_(ctx);

  const targetStartRow = ctx.targetDataStartRow;
  const targetLastRow = ctx.targetSheet.getLastRow();

  if (targetLastRow < targetStartRow) {
    return {
      scannedRows: 0,
      matchedRows: 0,
      changedRows: 0,
      changedCells: 0,
      writeOperations: 0,
      skippedInvalidRegion: 0,
      skippedNoId: 0,
      notFound: 0,
      targetAuthority: targetAuthorityResult
    };
  }

  const targetRowCount = targetLastRow - targetStartRow + 1;
  const sourceStartRow = ctx.sourceDataStartRow;
  const sourceLastRow = ctx.sourceSheet.getLastRow();

  if (sourceLastRow < sourceStartRow) {
    return {
      scannedRows: targetRowCount,
      matchedRows: 0,
      changedRows: 0,
      changedCells: 0,
      writeOperations: 0,
      skippedInvalidRegion: 0,
      skippedNoId: targetRowCount,
      notFound: 0,
      targetAuthority: targetAuthorityResult
    };
  }

  const sourceRowCount = sourceLastRow - sourceStartRow + 1;
  const sourceLastCol = Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol);

  const targetIds = ctx.targetSheet
    .getRange(targetStartRow, ctx.targetIdCol, targetRowCount, 1)
    .getDisplayValues();

  const sourceIds = ctx.sourceSheet
    .getRange(sourceStartRow, ctx.sourceIdCol, sourceRowCount, 1)
    .getDisplayValues();

  const sourceIndexById = new Map();

  sourceIds.forEach((row, index) => {
    const id = normalizeId_(row[0]);

    if (id && !sourceIndexById.has(id)) {
      sourceIndexById.set(id, index);
    }
  });

  const sourceRaw = ctx.sourceSheet
    .getRange(sourceStartRow, 1, sourceRowCount, sourceLastCol)
    .getValues();

  const sourceDisplay = ctx.sourceSheet
    .getRange(sourceStartRow, 1, sourceRowCount, sourceLastCol)
    .getDisplayValues();

  const targetCols = CMS5_collectTargetColumns_(ctx);
  const targetColumnData = {};
  const changedByColumn = {};

  targetCols.forEach(col => {
    targetColumnData[col] = ctx.targetSheet
      .getRange(targetStartRow, col, targetRowCount, 1)
      .getValues();
    changedByColumn[col] = [];
  });

  let matchedRows = 0;
  let skippedNoId = 0;
  let notFound = 0;
  let changedCells = 0;
  let skippedInvalidRegion = 0;
  const changedRowIndexes = new Set();

  for (let i = 0; i < targetRowCount; i++) {
    const idValue = targetIds[i][0];
    const normalizedId = normalizeId_(idValue);

    if (!normalizedId) {
      skippedNoId++;
      continue;
    }

    const sourceIndex = sourceIndexById.get(normalizedId);

    if (typeof sourceIndex === "undefined") {
      notFound++;
      continue;
    }

    matchedRows++;

    const rawRow = sourceRaw[sourceIndex];
    const displayRow = sourceDisplay[sourceIndex];

    ctx.resolvedFields.forEach(field => {
      if (CMS27_isTargetAuthoritativeField_(field)) return;

      const value = CMS5_makeTargetValueFromMasterField_(
        field,
        rawRow,
        displayRow
      );

      if (CMS21_isSkipWriteValue_(value)) {
        if (field.name === '지역') skippedInvalidRegion++;
        return;
      }

      const currentValue = targetColumnData[field.targetCol][i][0];

      if (!CMS19_valuesEqual_(currentValue, value)) {
        changedByColumn[field.targetCol].push({
          rowOffset: i,
          value: value
        });
        changedRowIndexes.add(i);
        changedCells++;
      }
    });
  }

  let writeOperations = 0;

  targetCols.forEach(col => {
    writeOperations += CMS19_writeChangedColumnRuns_(
      ctx.targetSheet,
      targetStartRow,
      col,
      changedByColumn[col]
    );
  });

  if (changedCells > 0) {
    SpreadsheetApp.flush();
  }

  return {
    scannedRows: targetRowCount,
    matchedRows: matchedRows,
    changedRows: changedRowIndexes.size,
    changedCells: changedCells,
    writeOperations: writeOperations,
    skippedInvalidRegion: skippedInvalidRegion,
    unchangedRows: Math.max(0, matchedRows - changedRowIndexes.size),
    skippedNoId: skippedNoId,
    notFound: notFound,
    targetAuthority: targetAuthorityResult
  };
}

/**
 * 5분 안전 동기화에서도 권위 저장상태는 수주확정 값을 마스터로 보정한다.
 * 동일 고객번호가 여러 행이면 가장 아래쪽 행 하나만 사용한다.
 */
function CMS27_reconcileAllTargetAuthorityToMaster_(ctx) {
  const fields = CMS27_getTargetAuthoritativeFields_(ctx);
  const targetStartRow = ctx.targetDataStartRow;
  const sourceStartRow = ctx.sourceDataStartRow;
  const targetLastRow = ctx.targetSheet.getLastRow();
  const sourceLastRow = ctx.sourceSheet.getLastRow();

  if (!fields.length || targetLastRow < targetStartRow || sourceLastRow < sourceStartRow) {
    return {
      scannedTargetRows: 0,
      canonicalCustomers: 0,
      changedRows: 0,
      changedCells: 0,
      writeOperations: 0,
      duplicateTargetRows: 0,
      skippedNoId: 0,
      sourceNotFound: 0
    };
  }

  const targetRowCount = targetLastRow - targetStartRow + 1;
  const sourceRowCount = sourceLastRow - sourceStartRow + 1;
  const targetIds = ctx.targetSheet
    .getRange(targetStartRow, ctx.targetIdCol, targetRowCount, 1)
    .getDisplayValues();
  const sourceIds = ctx.sourceSheet
    .getRange(sourceStartRow, ctx.sourceIdCol, sourceRowCount, 1)
    .getDisplayValues();

  const sourceIndexById = new Map();
  sourceIds.forEach(function(row, index) {
    const id = normalizeId_(row[0]);
    if (id && !sourceIndexById.has(id)) sourceIndexById.set(id, index);
  });

  // 아래쪽 행이 최신이므로 같은 고객번호가 다시 나오면 덮어쓴다.
  const canonicalTargetIndexById = new Map();
  let skippedNoId = 0;
  let duplicateTargetRows = 0;
  targetIds.forEach(function(row, index) {
    const id = normalizeId_(row[0]);
    if (!id) {
      skippedNoId++;
      return;
    }
    if (canonicalTargetIndexById.has(id)) duplicateTargetRows++;
    canonicalTargetIndexById.set(id, index);
  });

  const targetLastCol = Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol);
  const sourceLastCol = Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol);
  const targetRaw = ctx.targetSheet
    .getRange(targetStartRow, 1, targetRowCount, targetLastCol)
    .getValues();
  const targetDisplay = ctx.targetSheet
    .getRange(targetStartRow, 1, targetRowCount, targetLastCol)
    .getDisplayValues();
  const sourceRaw = ctx.sourceSheet
    .getRange(sourceStartRow, 1, sourceRowCount, sourceLastCol)
    .getValues();

  const changedByColumn = {};
  fields.forEach(function(field) { changedByColumn[field.sourceCol] = []; });

  let sourceNotFound = 0;
  let changedCells = 0;
  const changedSourceIndexes = new Set();

  canonicalTargetIndexById.forEach(function(targetIndex, id) {
    const sourceIndex = sourceIndexById.get(id);
    if (typeof sourceIndex === 'undefined') {
      sourceNotFound++;
      return;
    }

    fields.forEach(function(field) {
      const value = getByMode_(
        targetRaw[targetIndex],
        targetDisplay[targetIndex],
        field.targetCol,
        field.reverseValueMode || field.valueMode || 'raw'
      );
      const current = sourceRaw[sourceIndex][field.sourceCol - 1];

      if (!CMS19_valuesEqual_(current, value)) {
        changedByColumn[field.sourceCol].push({ rowOffset: sourceIndex, value: value });
        sourceRaw[sourceIndex][field.sourceCol - 1] = value;
        changedSourceIndexes.add(sourceIndex);
        changedCells++;
      }
    });
  });

  let writeOperations = 0;
  Object.keys(changedByColumn).forEach(function(colText) {
    const col = Number(colText);
    writeOperations += CMS19_writeChangedColumnRuns_(
      ctx.sourceSheet,
      sourceStartRow,
      col,
      changedByColumn[col]
    );
  });

  if (changedCells > 0) SpreadsheetApp.flush();

  return {
    scannedTargetRows: targetRowCount,
    canonicalCustomers: canonicalTargetIndexById.size,
    changedRows: changedSourceIndexes.size,
    changedCells: changedCells,
    writeOperations: writeOperations,
    duplicateTargetRows: duplicateTargetRows,
    skippedNoId: skippedNoId,
    sourceNotFound: sourceNotFound
  };
}

/**
 * 실제 반영 없이 현재 권위 저장상태 충돌 건수를 확인한다.
 */
function CMS27_previewTargetAuthorityConflicts() {
  const ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  const ctx = buildContractMasterSyncContext_(ss);
  const fields = CMS27_getTargetAuthoritativeFields_(ctx);
  const targetStartRow = ctx.targetDataStartRow;
  const sourceStartRow = ctx.sourceDataStartRow;
  const targetLastRow = ctx.targetSheet.getLastRow();
  const sourceLastRow = ctx.sourceSheet.getLastRow();

  if (!fields.length || targetLastRow < targetStartRow || sourceLastRow < sourceStartRow) {
    const empty = { status: 'NO_DATA', conflictRows: 0, conflictCells: 0 };
    Logger.log(JSON.stringify(empty));
    return empty;
  }

  const targetRowCount = targetLastRow - targetStartRow + 1;
  const sourceRowCount = sourceLastRow - sourceStartRow + 1;
  const targetValues = ctx.targetSheet
    .getRange(targetStartRow, 1, targetRowCount, Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol))
    .getValues();
  const targetDisplay = ctx.targetSheet
    .getRange(targetStartRow, 1, targetRowCount, Math.max(ctx.targetSheet.getLastColumn(), ctx.maxTargetCol))
    .getDisplayValues();
  const sourceValues = ctx.sourceSheet
    .getRange(sourceStartRow, 1, sourceRowCount, Math.max(ctx.sourceSheet.getLastColumn(), ctx.maxSourceCol))
    .getValues();

  const sourceIndexById = new Map();
  sourceValues.forEach(function(row, index) {
    const id = normalizeId_(row[ctx.sourceIdCol - 1]);
    if (id && !sourceIndexById.has(id)) sourceIndexById.set(id, index);
  });

  const canonicalTargetIndexById = new Map();
  targetDisplay.forEach(function(row, index) {
    const id = normalizeId_(row[ctx.targetIdCol - 1]);
    if (id) canonicalTargetIndexById.set(id, index);
  });

  let conflictCells = 0;
  let sourceNotFound = 0;
  const conflictRows = [];

  canonicalTargetIndexById.forEach(function(targetIndex, id) {
    const sourceIndex = sourceIndexById.get(id);
    if (typeof sourceIndex === 'undefined') {
      sourceNotFound++;
      return;
    }

    const differences = [];
    fields.forEach(function(field) {
      const targetValue = getByMode_(
        targetValues[targetIndex],
        targetDisplay[targetIndex],
        field.targetCol,
        field.reverseValueMode || field.valueMode || 'raw'
      );
      const sourceValue = sourceValues[sourceIndex][field.sourceCol - 1];
      if (!CMS19_valuesEqual_(sourceValue, targetValue)) {
        conflictCells++;
        differences.push({
          field: field.name,
          targetValue: targetValue,
          masterValue: sourceValue
        });
      }
    });

    if (differences.length) {
      conflictRows.push({
        customerId: id,
        targetRow: targetStartRow + targetIndex,
        masterRow: sourceStartRow + sourceIndex,
        differences: differences
      });
    }
  });

  const result = {
    status: 'PREVIEW',
    conflictRows: conflictRows.length,
    conflictCells: conflictCells,
    sourceNotFound: sourceNotFound,
    sample: conflictRows.slice(0, 30)
  };
  Logger.log(JSON.stringify(result));
  return result;
}

function CMS27_reconcileTargetAuthorityNow() {
  const ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  const ctx = buildContractMasterSyncContext_(ss);
  const result = CMS27_reconcileAllTargetAuthorityToMaster_(ctx);
  Logger.log(JSON.stringify(result));
  CMS5_safeToast_(
    ss,
    '저장상태 우선값 보정 완료: 변경행 ' + result.changedRows + ' / 변경셀 ' + result.changedCells,
    '수주확정 우선값 보정',
    8
  );
  return result;
}

/****************************************************
 * 19단계: 변경값만 쓰는 동기화 유틸
 ****************************************************/
var CMS21_SKIP_WRITE_ = Object.freeze({ __cms21SkipWrite: true });

function CMS21_isSkipWriteValue_(value) {
  return !!(value && typeof value === 'object' && value.__cms21SkipWrite === true);
}

/**
 * 수주확정 '지역' 열의 데이터 유효성은 7개 권역만 허용한다.
 * 마스터가 주소확인필요/미확정/임의문구이면 기존 수주확정 값을 보존한다.
 */
function CMS21_normalizeRegionForTargetOrSkip_(value) {
  const normalized = normalizeRegionGroupForTarget_(value);
  const allowed = [
    '강원권',
    '대구경북권',
    '부울경권',
    '수도권',
    '제주권',
    '충청권',
    '호남권'
  ];

  if (!normalized || normalized === '주소확인필요') {
    return CMS21_SKIP_WRITE_;
  }

  return allowed.indexOf(normalized) !== -1
    ? normalized
    : CMS21_SKIP_WRITE_;
}


function CMS19_valuesEqual_(currentValue, desiredValue) {
  const currentBlank = isBlank_(currentValue);
  const desiredBlank = isBlank_(desiredValue);

  if (currentBlank || desiredBlank) {
    return currentBlank && desiredBlank;
  }

  if (currentValue instanceof Date || desiredValue instanceof Date) {
    if (!(currentValue instanceof Date) || !(desiredValue instanceof Date)) {
      return false;
    }

    return currentValue.getTime() === desiredValue.getTime();
  }

  if (typeof currentValue === "number" && typeof desiredValue === "number") {
    if (Number.isNaN(currentValue) && Number.isNaN(desiredValue)) return true;
    return currentValue === desiredValue;
  }

  if (typeof currentValue === "boolean" || typeof desiredValue === "boolean") {
    return currentValue === desiredValue;
  }

  return CMS19_normalizeComparableText_(currentValue) ===
    CMS19_normalizeComparableText_(desiredValue);
}

function CMS19_normalizeComparableText_(value) {
  return String(value === null || typeof value === "undefined" ? "" : value)
    .replace(/\r\n/g, "\n")
    .replace(/\r/g, "\n")
    .replace(/\u00A0/g, " ");
}

function CMS19_writeChangedRowCells_(sheet, row, cells) {
  if (!cells || !cells.length) {
    return { writeOperations: 0 };
  }

  const sorted = cells.slice().sort(function(a, b) { return a.col - b.col; });
  let writeOperations = 0;
  let blockStartCol = sorted[0].col;
  let blockValues = [sorted[0].value];
  let previousCol = sorted[0].col;

  for (let i = 1; i < sorted.length; i++) {
    const cell = sorted[i];

    if (cell.col === previousCol + 1) {
      blockValues.push(cell.value);
      previousCol = cell.col;
      continue;
    }

    sheet.getRange(row, blockStartCol, 1, blockValues.length).setValues([blockValues]);
    writeOperations++;

    blockStartCol = cell.col;
    blockValues = [cell.value];
    previousCol = cell.col;
  }

  sheet.getRange(row, blockStartCol, 1, blockValues.length).setValues([blockValues]);
  writeOperations++;

  return { writeOperations: writeOperations };
}

function CMS19_writeChangedColumnRuns_(sheet, startRow, col, changes) {
  if (!changes || !changes.length) return 0;

  const sorted = changes.slice().sort(function(a, b) {
    return a.rowOffset - b.rowOffset;
  });

  let writeOperations = 0;
  let runStartOffset = sorted[0].rowOffset;
  let runValues = [[sorted[0].value]];
  let previousOffset = sorted[0].rowOffset;

  for (let i = 1; i < sorted.length; i++) {
    const change = sorted[i];

    if (change.rowOffset === previousOffset + 1) {
      runValues.push([change.value]);
      previousOffset = change.rowOffset;
      continue;
    }

    sheet
      .getRange(startRow + runStartOffset, col, runValues.length, 1)
      .setValues(runValues);
    writeOperations++;

    runStartOffset = change.rowOffset;
    runValues = [[change.value]];
    previousOffset = change.rowOffset;
  }

  sheet
    .getRange(startRow + runStartOffset, col, runValues.length, 1)
    .setValues(runValues);
  writeOperations++;

  return writeOperations;
}


/****************************************************
 * A시트에 쓰는 열 목록 수집
 ****************************************************/
function CMS5_collectTargetColumns_(ctx) {
  const cols = new Set();

  ctx.resolvedFields.forEach(field => {
    if (CMS27_isTargetAuthoritativeField_(field)) return;
    cols.add(field.targetCol);
  });

  return Array.from(cols).sort((a, b) => a - b);
}


/****************************************************
 * 기존 writeMasterRowToTargetRow_의 값 생성 로직을
 * 빠른 동기화용으로 분리한 버전
 ****************************************************/
function CMS5_makeTargetValueFromMasterField_(field, raw, display) {
  let value = "";

  if (field.type === "direct") {
    value = getByMode_(
      raw,
      display,
      field.sourceCol,
      field.valueMode || "raw"
    );

    if (field.name === "지역") {
      value = CMS21_normalizeRegionForTargetOrSkip_(value);
    }

    return value;
  }

  if (field.type === "period") {
    const start = getByMode_(
      raw,
      display,
      field.sourceStartCol,
      "display"
    );

    const end = getByMode_(
      raw,
      display,
      field.sourceEndCol,
      "display"
    );

    return makePeriodText_(start, end);
  }

  if (field.type === "conditionalExtractNumber") {
    const conditionValue = String(
      display[field.conditionSourceCol - 1] || ""
    ).trim();

    if (conditionValue === field.conditionText) {
      return extractFirstNumber_(display[field.valueSourceCol - 1]);
    }

    if (conditionValue === "비선임") {
      return 0;
    }

    return "";
  }

  if (field.type === "extractNumber") {
    return extractFirstNumber_(display[field.sourceCol - 1]);
  }

  return "";
}


/****************************************************
 * 시간기반 트리거에서 toast가 안 먹어도 터지지 않게 처리
 ****************************************************/
function CMS5_safeToast_(ss, message, title, seconds) {
  try {
    ss.toast(message, title, seconds);
  } catch (err) {
    console.log(`${title}: ${message}`);
  }
}