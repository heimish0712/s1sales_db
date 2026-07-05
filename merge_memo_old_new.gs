/*************************************************
 * merge_memo_old_new_final_v4.gs
 * 마스터시트 메모 최종 업데이트본 생성 V4 FINAL
 *
 * 입력:
 * - 활성시트 '메모'
 * - 활성시트 '마스터시트 메모 원본확인'
 *
 * 출력:
 * - 활성시트 '마스터시트 메모 최종 업데이트본'
 *
 * 동작:
 * - 기존 최종 업데이트본 열이 있으면 헤더 유지
 * - 3행부터 데이터만 삭제 후 재기록
 *
 * 삭제 대상 TM:
 * - (유현희TM), 유현희TM, 유현희 TM
 * - (라유화TM), 라유화TM, 라유화 TM
 * - (박원경TM), 박원경TM, 박원경 TM, 박원경TM확인
 * - (김서하TM), 김서하TM, 김서하 TM, 김서하TM에게, 검토중(김서하TM
 * - (TM 미상), TM 미상
 *
 * 보존:
 * - 김서하
 * - (김서하)
 * - 김서하대리
 * - 유현경TM
 * - 노이현TM
 * - 윤현경 TM
 * - 차고은 TM
 *************************************************/

const FINAL_MEMO_V3_CONFIG = {
  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  MEMO_HEADER_NAME: '메모',
  ORIGINAL_MEMO_HEADER_NAME: '마스터시트 메모 원본확인',
  FINAL_MEMO_HEADER_NAME: '마스터시트 메모 최종 업데이트본',

  DEFAULT_YEAR: 2026,

  SORT_BY_DATE: false
};




/**
 * 메인 실행 함수
 */
function buildFinalMasterMemoUpdateColumn() {
  const cfg = FINAL_MEMO_V3_CONFIG;

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  const lastRow = sheet.getLastRow();

  if (lastRow < cfg.DATA_START_ROW) {
    SpreadsheetApp.getUi().alert('처리할 데이터 행이 없습니다.');
    return;
  }

  const headerMap = getFinalMemoHeaderMapV3_(sheet, cfg.HEADER_ROW);

  const memoCol = headerMap[normalizeFinalMemoHeaderV3_(cfg.MEMO_HEADER_NAME)];
  const originalMemoCol = headerMap[normalizeFinalMemoHeaderV3_(cfg.ORIGINAL_MEMO_HEADER_NAME)];

  if (!memoCol) {
    throw new Error(`활성시트에서 '${cfg.MEMO_HEADER_NAME}' 헤더를 찾을 수 없습니다.`);
  }

  if (!originalMemoCol) {
    throw new Error(`활성시트에서 '${cfg.ORIGINAL_MEMO_HEADER_NAME}' 헤더를 찾을 수 없습니다.`);
  }

  const finalCol = createOrGetFinalMemoColumnV3_(
    sheet,
    cfg.FINAL_MEMO_HEADER_NAME,
    cfg.HEADER_ROW
  );

  const numRows = lastRow - cfg.DATA_START_ROW + 1;

  // 기존 최종 업데이트본 데이터만 삭제. 헤더는 유지.
  sheet
    .getRange(cfg.DATA_START_ROW, finalCol, numRows, 1)
    .clearContent();

  SpreadsheetApp.flush();

  const memoValues = sheet
    .getRange(cfg.DATA_START_ROW, memoCol, numRows, 1)
    .getDisplayValues();

  const originalValues = sheet
    .getRange(cfg.DATA_START_ROW, originalMemoCol, numRows, 1)
    .getDisplayValues();

  const output = [];

  const stats = {
    processedRows: 0,
    outputRows: 0,
    targetTmRemoved: 0,
    meaninglessRemoved: 0,
    materialSendRecords: 0,
    contactRecords: 0,
    nextActionRecords: 0,
    supportRequestRecords: 0,
    rawPreservedRecords: 0,
    duplicateRemoved: 0,
    replacedByLongerDuplicate: 0,
    invalidDateIgnored: 0,
    jsDateNormalized: 0
  };

  for (let i = 0; i < numRows; i++) {
    const memoText = memoValues[i][0] || '';
    const originalText = originalValues[i][0] || '';

    const finalText = buildFinalMemoTextV3_(originalText, memoText, stats);

    output.push([finalText]);

    stats.processedRows++;

    if (finalText) {
      stats.outputRows++;
    }
  }

  sheet
    .getRange(cfg.DATA_START_ROW, finalCol, numRows, 1)
    .setValues(output);

  SpreadsheetApp.getUi().alert(
    [
      '마스터시트 메모 최종 업데이트본 재생성 완료',
      '',
      `처리 행: ${stats.processedRows}건`,
      `결과 생성 행: ${stats.outputRows}건`,
      '',
      `대상 TM 삭제: ${stats.targetTmRemoved}건`,
      `내용 없는 줄 삭제: ${stats.meaninglessRemoved}건`,
      `JS 날짜 정규화: ${stats.jsDateNormalized}건`,
      `날짜 오인식 방지/무시: ${stats.invalidDateIgnored}건`,
      '',
      `자료발송 정규화: ${stats.materialSendRecords}건`,
      `컨택이력 정규화: ${stats.contactRecords}건`,
      `다음액션 보존: ${stats.nextActionRecords}건`,
      `영업지원요청 보존: ${stats.supportRequestRecords}건`,
      `원문 보존: ${stats.rawPreservedRecords}건`,
      '',
      `중복 제거: ${stats.duplicateRemoved}건`,
      `긴 이력으로 대체: ${stats.replacedByLongerDuplicate}건`,
      '',
      `출력 열: ${cfg.FINAL_MEMO_HEADER_NAME}`
    ].join('\n')
  );
}


/*************************************************
 * 최종 메모 생성
 *************************************************/

function buildFinalMemoTextV3_(originalText, memoText, stats) {
  const originalRecords = buildMemoRecordsFromTextV3_(originalText, 'original', stats);
  const memoRecords = buildMemoRecordsFromTextV3_(memoText, 'memo', stats);

  let merged = mergeMemoRecordsV3_(originalRecords, memoRecords, stats);

  if (FINAL_MEMO_V3_CONFIG.SORT_BY_DATE) {
    merged = sortMemoRecordsByDateV3_(merged);
  }

  return merged
    .map(record => record.normalizedText)
    .filter(text => text && String(text).trim() !== '')
    .join('\n');
}


function buildMemoRecordsFromTextV3_(text, sourceName, stats) {
  const segments = splitMemoToSegmentsV3_(text, stats);
  const records = [];

  for (let i = 0; i < segments.length; i++) {
    const record = parseMemoSegmentV3_(segments[i], sourceName, stats);

    if (record && record.normalizedText) {
      records.push(record);
    }
  }

  return records;
}


/*************************************************
 * 세그먼트 분해
 *************************************************/

function splitMemoToSegmentsV3_(text, stats) {
  if (text === null || text === undefined) return [];

  stats = stats || {
    targetTmRemoved: 0,
    meaninglessRemoved: 0
  };

  const raw = String(text)
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .trim();

  if (!raw) return [];

  const lines = raw.split('\n');
  const result = [];

  for (let i = 0; i < lines.length; i++) {
    const line = cleanMemoSegmentV3_(lines[i]);

    if (!line) continue;

    // 대상 TM이 붙은 줄이라도 같은 셀/같은 줄 안에 정상 메모가 앞뒤로 섞일 수 있으므로
    // 여기서 줄 전체를 바로 삭제하지 않는다.
    // 먼저 // 단위와 날짜/차수 경계로 최대한 쪼갠 뒤, 대상 TM이 포함된 조각만 삭제한다.
    const slashParts = line.split(/\/{2,}/g);

    for (let j = 0; j < slashParts.length; j++) {
      const part = cleanMemoSegmentV3_(slashParts[j]);

      if (!part) continue;

      if (isMeaninglessMemoSegmentV3_(part) && !hasTargetTmMarkerV3_(part)) {
        stats.meaninglessRemoved++;
        continue;
      }

      // 대상 TM이 없는 기존 정형 로그는 과도한 분해 방지
      if (!hasTargetTmMarkerV3_(part) && isStrongStructuredLineV3_(part)) {
        result.push(part);
        continue;
      }

      const subParts = splitByEventBoundariesV3_(part);

      for (let k = 0; k < subParts.length; k++) {
        const sub = cleanMemoSegmentV3_(subParts[k]);

        if (!sub) continue;

        // 대상 TM 표기가 포함된 조각만 삭제한다.
        // 예: "공문내용은 알고계심 6/22-담당자 미기재건 패스함(유현희TM)"
        //     → 앞 정상문장은 보존, 6/22 TM 조각만 삭제
        if (hasTargetTmMarkerV3_(sub)) {
          stats.targetTmRemoved++;
          continue;
        }

        if (isMeaninglessMemoSegmentV3_(sub)) {
          stats.meaninglessRemoved++;
          continue;
        }

        result.push(sub);
      }
    }
  }

  return result;
}


/**
 * 한 줄 안에 여러 날짜/차수 이벤트가 붙은 경우 분해.
 * 단, 1.5콜, 안산1. 29,636 같은 숫자/소수/면적/금액 오인식 방지.
 */
function splitByEventBoundariesV3_(line) {
  const text = String(line || '').trim();

  if (!text) return [];

  const boundaries = [];

  const candidateRegex = /(?:^|[\s,;，、/])((?:\d+\s*[차치]\s*[).]?\s*)?(?:(?:20\d{2}|\d{2})\s*[.\/-]\s*\d{1,2}\s*[.\/-]\s*\d{1,2}|(?:\d{1,2}\s*월\s*\d{1,2}\s*일?)|(?:\d{1,2}\s*\/\s*\d{1,2}\s*일?)|(?:\d{1,2}\s*\.\s*\d{1,2}\s*일?)|(?:\d+\s*[차치]\s*[).])))/g;

  let match;

  while ((match = candidateRegex.exec(text)) !== null) {
    const full = match[0];
    const candidate = match[1];
    const startIndex = match.index + full.indexOf(candidate);

    if (startIndex <= 0 || startIndex >= text.length) continue;

    if (isSafeEventBoundaryV3_(text, startIndex)) {
      boundaries.push(startIndex);
    }
  }

  const uniqueBoundaries = Array.from(new Set(boundaries))
    .sort((a, b) => a - b);

  if (uniqueBoundaries.length === 0) {
    return [text];
  }

  const parts = [];
  let prev = 0;

  for (let i = 0; i < uniqueBoundaries.length; i++) {
    const idx = uniqueBoundaries[i];
    const part = text.slice(prev, idx);

    if (cleanMemoSegmentV3_(part)) {
      parts.push(part);
    }

    prev = idx;
  }

  const last = text.slice(prev);

  if (cleanMemoSegmentV3_(last)) {
    parts.push(last);
  }

  return parts;
}


function isSafeEventBoundaryV3_(text, startIndex) {
  const before = startIndex > 0 ? text[startIndex - 1] : '';
  const afterChunk = text.slice(startIndex, startIndex + 30);

  if (/\d/.test(before)) {
    return false;
  }

  if (/^\d{1,2}\s*\.\s*\d{1,2}\s*콜/i.test(afterChunk)) {
    return false;
  }

  if (/^\d{1,2}\s*\.\s*\d{1,2}\s*,\s*\d/.test(afterChunk)) {
    return false;
  }

  if (/^\d{1,2}\s*\.\s*\d{1,2}\s*(?:콜|건|개|명|평|㎡|제곱|만원|원)/.test(afterChunk)) {
    return false;
  }

  // 1.2 6/8 같은 구조에서 1.2를 날짜로 잡는 것 방지
  if (/^\d{1,2}\s*\.\s*\d{1,2}\s+\d{1,2}\s*[\/월]/.test(afterChunk)) {
    return false;
  }

  return true;
}


function cleanMemoSegmentV3_(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .replace(/^[\s,.;，、/|]+/g, '')
    .replace(/[\s,.;，、/|]+$/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}


/*************************************************
 * 삭제/무의미 판정
 *************************************************/

/**
 * 삭제 대상 TM.
 *
 * 삭제:
 * - (유현희TM), 유현희TM, 유현희 TM
 * - (라유화TM), 라유화TM, 라유화 TM
 * - (박원경TM), 박원경TM, 박원경 TM, 박원경TM확인
 * - (김서하TM), 김서하TM, 김서하 TM, 김서하TM에게, 검토중(김서하TM
 * - (TM 미상), TM 미상
 *
 * 보존:
 * - 김서하
 * - (김서하)
 * - 김서하대리
 * - 유현경TM
 * - 노이현TM
 * - 윤현경 TM
 * - 차고은 TM
 */
function hasTargetTmMarkerV3_(text) {
  const value = String(text || '');

  if (/(?:\(\s*)?(?:유현희|라유화|박원경|김서하)\s*TM(?:\s*\))?/i.test(value)) {
    return true;
  }

  if (/(?:\(\s*)?TM\s*미상(?:\s*\))?/i.test(value)) {
    return true;
  }

  return false;
}


/**
 * 내용 없는 줄 삭제.
 */
function isMeaninglessMemoSegmentV3_(text) {
  let value = String(text || '').trim();

  if (!value) return true;

  value = value
    .replace(/^"+|"+$/g, '')
    .replace(/^'+|'+$/g, '')
    .replace(/^`+|`+$/g, '')
    .trim();

  if (!value) return true;

  // 괄호/따옴표/구분기호만 남은 경우
  if (/^[()\[\]{}"'`]+$/.test(value)) return true;

  // 구분선만 있는 경우
  if (/^[=\-_*#\s]+$/.test(value)) return true;

  // 시간만 있는 경우: 08:27
  if (/^\d{1,2}:\d{2}$/.test(value)) return true;

  // 날짜만 있는 경우
  if (/^(?:20\d{2}|\d{2})[.\/-]\d{1,2}[.\/-]\d{1,2}\.?$/.test(value)) return true;
  if (/^\d{1,2}[.\/]\d{1,2}[.)]?$/.test(value)) return true;
  if (/^\d{1,2}\s*월\s*\d{1,2}\s*일?[.)]?$/.test(value)) return true;

  // JS Date 객체 문자열만 있는 경우
  if (/^(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}\s+20\d{2}\s+00:00:00\s+GMT[+-]\d{4}/i.test(value)) {
    return true;
  }

  // 단독 숫자/단독 문자 찌꺼리
  if (/^\d{1,2}$/.test(value)) return true;
  if (/^[ㄱ-ㅎㅏ-ㅣ]$/.test(value)) return true;

  // tm만 있는 경우
  if (/^tm$/i.test(value)) return true;

  // 태그 제거 후 내용이 없는 경우
  const withoutTags = value
    .replace(/^\[컨택이력\](?:\[[^\]]+\])*/g, '')
    .replace(/^\[자료발송\]/g, '')
    .replace(/^\[다음액션\]/g, '')
    .replace(/^\[\d{2}\.\d{2}\.\d{2}(?:\s+\d{1,2}:\d{2})?\]\s*\[영업지원요청[^\]]*\]/g, '')
    .replace(/\d{2}\.\d{2}\.\d{2}(?:\s+\d{1,2}:\d{2})?\.?/g, '')
    .replace(/\d{1,2}[.\/]\d{1,2}[.)]?/g, '')
    .replace(/[()\[\]{}"'`~!@#$%^&*_+=<>?:;,.，、/\-|\\\s]/g, '')
    .trim();

  if (!withoutTags) return true;

  return false;
}


/*************************************************
 * 세그먼트 파싱
 *************************************************/

function parseMemoSegmentV3_(segment, sourceName, stats) {
  const original = cleanMemoSegmentV3_(segment);

  if (!original) return null;

  if (hasTargetTmMarkerV3_(original)) {
    stats.targetTmRemoved++;
    return null;
  }

  if (isMeaninglessMemoSegmentV3_(original)) {
    stats.meaninglessRemoved++;
    return null;
  }

  // 다음액션은 그대로 보존
  if (/^\[다음액션\]/.test(original)) {
    if (isMeaninglessMemoSegmentV3_(removeKnownPrefixesV3_(original))) {
      stats.meaninglessRemoved++;
      return null;
    }

    stats.nextActionRecords++;

    return createMemoRecordV3_({
      type: '다음액션',
      round: '',
      dateInfo: extractDateInfoV3_(original, stats),
      body: original,
      normalizedText: original,
      sourceName: sourceName
    });
  }

  // 영업지원요청 접수 로그는 그대로 보존
  if (/\[영업지원요청\s*접수\s*#?\d*\]/.test(original)) {
    stats.supportRequestRecords++;

    return createMemoRecordV3_({
      type: '영업지원요청',
      round: '',
      dateInfo: extractDateInfoV3_(original, stats),
      body: original,
      normalizedText: original,
      sourceName: sourceName
    });
  }

  const structuredContact = parseStructuredContactLineV3_(original, stats);

  let working = original;
  let existingStructuredContact = false;

  if (structuredContact) {
    existingStructuredContact = true;
    working = structuredContact.body || original;
  }

  let round = structuredContact ? structuredContact.round : '';
  let dateInfo = structuredContact ? structuredContact.dateInfo : null;

  if (!round) {
    const roundExtracted = extractRoundFromTextV3_(working);
    round = roundExtracted.round;
    working = roundExtracted.text;
  }

  if (!dateInfo) {
    dateInfo = extractDateInfoV3_(working, stats);
  }

  if (dateInfo) {
    working = removeDateFromTextByDateInfoV3_(working, dateInfo);
  }

  if (!round) {
    const roundExtractedAfterDate = extractRoundFromTextV3_(working);
    round = roundExtractedAfterDate.round;
    working = roundExtractedAfterDate.text;
  }

  const body = cleanMemoBodyV3_(working);

  if (isMeaninglessMemoSegmentV3_(body)) {
    stats.meaninglessRemoved++;
    return null;
  }

  const isMaterial = detectMaterialSendMemoV3_(original, body);

  let type = '원문';

  if (isMaterial) {
    type = '자료발송';
  } else if (existingStructuredContact || round || dateInfo) {
    type = '컨택이력';
  }

  let normalizedText;

  if (existingStructuredContact && type === '컨택이력') {
    normalizedText = original;
  } else {
    normalizedText = buildNormalizedMemoLineV3_(type, round, dateInfo, body || original);
  }

  normalizedText = cleanFinalOutputLineV3_(normalizedText);

  if (!normalizedText || isMeaninglessMemoSegmentV3_(removeKnownPrefixesV3_(normalizedText))) {
    stats.meaninglessRemoved++;
    return null;
  }

  if (type === '자료발송') {
    stats.materialSendRecords++;
  } else if (type === '컨택이력') {
    stats.contactRecords++;
  } else if (type === '원문') {
    stats.rawPreservedRecords++;
  }

  return createMemoRecordV3_({
    type: type,
    round: round,
    dateInfo: dateInfo,
    body: body || original,
    normalizedText: normalizedText,
    sourceName: sourceName
  });
}


function parseStructuredContactLineV3_(text, stats) {
  const value = String(text || '').trim();

  if (!/^\[컨택이력\]/.test(value)) {
    return null;
  }

  let round = '';

  const roundMatch = value.match(/\[(\d+\s*차)\]/);

  if (roundMatch) {
    round = roundMatch[1].replace(/\s+/g, '');
  }

  const dateInfo = extractDateInfoV3_(value, stats);

  let body = value;

  if (dateInfo) {
    body = value.slice(dateInfo.endIndex);
  } else {
    body = value.replace(/^\[컨택이력\](?:\[[^\]]+\])*/g, '');
  }

  body = cleanMemoBodyV3_(body);

  return {
    round: round,
    dateInfo: dateInfo,
    body: body
  };
}


function buildNormalizedMemoLineV3_(type, round, dateInfo, body) {
  const cleanedBody = cleanMemoBodyV3_(body);

  if (!cleanedBody) return '';

  if (type === '자료발송') {
    if (dateInfo) {
      return `[자료발송] ${dateInfo.display} ${cleanedBody}`.trim();
    }

    return `[자료발송] ${cleanedBody}`.trim();
  }

  if (type === '컨택이력') {
    let prefix = '[컨택이력]';

    if (round) {
      prefix += `[${round}]`;
    }

    if (dateInfo) {
      return `${prefix} ${dateInfo.display} ${cleanedBody}`.trim();
    }

    return `${prefix} ${cleanedBody}`.trim();
  }

  return cleanedBody;
}


function cleanMemoBodyV3_(value) {
  return String(value || '')
    .replace(/^\[컨택이력\](?:\[[^\]]+\])*/g, '')
    .replace(/^\[자료발송\]/g, '')
    .replace(/^\[다음액션\]/g, '[다음액션]')
    .replace(/^[\s:：\-–—.,，、/|]+/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}


function removeKnownPrefixesV3_(value) {
  return String(value || '')
    .replace(/^\[컨택이력\](?:\[[^\]]+\])*/g, '')
    .replace(/^\[자료발송\]/g, '')
    .replace(/^\[다음액션\]/g, '')
    .replace(/^\[\d{2}\.\d{2}\.\d{2}(?:\s+\d{1,2}:\d{2})?\]\s*\[영업지원요청[^\]]*\]/g, '')
    .trim();
}


/*************************************************
 * 차수 추출
 *************************************************/

function extractRoundFromTextV3_(text) {
  let value = String(text || '').trim();

  const match = value.match(/^(\d+)\s*[차치]\s*[).]?\s*/);

  if (!match) {
    return {
      round: '',
      text: value
    };
  }

  const round = `${Number(match[1])}차`;

  value = value.slice(match[0].length).trim();

  return {
    round: round,
    text: value
  };
}


/*************************************************
 * 날짜 추출/정규화 V3
 *************************************************/

function extractDateInfoV3_(text, stats) {
  const value = String(text || '');

  if (!value.trim()) return null;

  const candidates = [];

  collectDateCandidatesV3_(value, candidates, stats);

  if (candidates.length === 0) return null;

  candidates.sort((a, b) => a.startIndex - b.startIndex);

  const picked = candidates[0];

  const yy = String(picked.year).slice(-2);
  const mm = pad2V3_(picked.month);
  const dd = pad2V3_(picked.day);

  const displayDate = `${yy}.${mm}.${dd}.`;

  return {
    year: picked.year,
    year2: yy,
    month: picked.month,
    day: picked.day,
    time: picked.time || '',
    dateKey: `${yy}.${mm}.${dd}`,
    display: picked.time ? `${displayDate.slice(0, -1)} ${picked.time}` : displayDate,
    startIndex: picked.startIndex,
    endIndex: picked.endIndex
  };
}


function collectDateCandidatesV3_(value, candidates, stats) {
  const patterns = [
    // JS Date 문자열: Fri Jun 12 2026 00:00:00 GMT+0900 ...
    {
      kind: 'jsDate',
      regex: /\b(Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+(\d{1,2})\s+(20\d{2})\s+\d{2}:\d{2}:\d{2}\s+GMT[+-]\d{4}(?:\s+\([^)]+\))?/gi,
      build: function (m) {
        return {
          year: Number(m[4]),
          month: monthNameToNumberV3_(m[2]),
          day: Number(m[3]),
          time: ''
        };
      }
    },

    // 08:27 26.06.30 08:27 형태
    {
      kind: 'timeFirstYY',
      regex: /(\d{1,2}:\d{2})\s+(\d{2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{1,2})\.?\s*(\d{1,2}:\d{2})?/g,
      build: function (m) {
        return {
          year: 2000 + Number(m[2]),
          month: Number(m[3]),
          day: Number(m[4]),
          time: m[5] || m[1]
        };
      }
    },

    // 2026. 2. 27 / 2026.02.27 10:49
    {
      kind: 'yyyy',
      regex: /(20\d{2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{1,2})\.?\s*(\d{1,2}:\d{2})?/g,
      build: function (m) {
        return {
          year: Number(m[1]),
          month: Number(m[2]),
          day: Number(m[3]),
          time: m[4] || ''
        };
      }
    },

    // 26.07.02 / 26.07.02 10:49
    {
      kind: 'yy',
      regex: /(\d{2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{1,2})\.?\s*(\d{1,2}:\d{2})?/g,
      build: function (m) {
        return {
          year: 2000 + Number(m[1]),
          month: Number(m[2]),
          day: Number(m[3]),
          time: m[4] || ''
        };
      }
    },

    // 2월10일
    {
      kind: 'koreanMonthDay',
      regex: /(\d{1,2})\s*월\s*(\d{1,2})\s*일?/g,
      build: function (m) {
        return {
          year: FINAL_MEMO_V3_CONFIG.DEFAULT_YEAR,
          month: Number(m[1]),
          day: Number(m[2]),
          time: ''
        };
      }
    },

    // 2/10, 06/10, 2/10일
    {
      kind: 'slashMonthDay',
      regex: /(\d{1,2})\s*\/\s*(\d{1,2})\s*일?(?:\s*(\d{1,2}:\d{2}))?/g,
      build: function (m) {
        return {
          year: FINAL_MEMO_V3_CONFIG.DEFAULT_YEAR,
          month: Number(m[1]),
          day: Number(m[2]),
          time: m[3] || ''
        };
      }
    },

    // 4.14, 4.1, 06.10
    {
      kind: 'dotMonthDay',
      regex: /(\d{1,2})\s*\.\s*(\d{1,2})\s*일?(?:\s*(\d{1,2}:\d{2}))?/g,
      build: function (m) {
        return {
          year: FINAL_MEMO_V3_CONFIG.DEFAULT_YEAR,
          month: Number(m[1]),
          day: Number(m[2]),
          time: m[3] || ''
        };
      }
    }
  ];

  for (let p = 0; p < patterns.length; p++) {
    const pattern = patterns[p];
    let match;

    pattern.regex.lastIndex = 0;

    while ((match = pattern.regex.exec(value)) !== null) {
      const info = pattern.build(match);

      const startIndex = match.index;
      const endIndex = match.index + match[0].length;

      if (!isValidMemoDateV3_(info.year, info.month, info.day)) {
        if (stats) stats.invalidDateIgnored++;
        continue;
      }

      if (!isSafeDateMatchContextV3_(value, startIndex, endIndex, pattern.kind, match[0])) {
        if (stats) stats.invalidDateIgnored++;
        continue;
      }

      if (pattern.kind === 'jsDate' && stats) {
        stats.jsDateNormalized++;
      }

      candidates.push({
        year: info.year,
        month: info.month,
        day: info.day,
        time: info.time || '',
        startIndex: startIndex,
        endIndex: endIndex,
        kind: pattern.kind
      });
    }
  }
}


function isSafeDateMatchContextV3_(text, startIndex, endIndex, kind, matchedText) {
  const before = startIndex > 0 ? text[startIndex - 1] : '';
  const after2 = text.slice(endIndex, endIndex + 12);
  const chunk = text.slice(startIndex, startIndex + 40);

  if (kind === 'jsDate') return true;
  if (kind === 'yyyy') return true;
  if (kind === 'yy') return true;

  // 숫자 바로 뒤에서 시작하면 날짜가 아니라 숫자 일부일 가능성이 큼
  // 예: 안산1. 29,636에서 1.29를 날짜로 잡는 것 방지
  if (/\d/.test(before)) {
    return false;
  }

  // 숫자가 바로 이어지면 날짜가 아니라 코드/면적/금액 일부일 가능성이 큼
  if (/^\d/.test(after2)) {
    return false;
  }

  // 쉼표 뒤 숫자는 면적/금액/수치일 가능성이 큼: 1. 29,636 등
  // 단, 4.24,인바운드 처럼 쉼표 뒤 문자가 오는 날짜는 허용한다.
  if (/^\s*,\s*\d/.test(after2)) {
    return false;
  }

  if (kind === 'slashMonthDay') {
    // 1/ 1년치, 유지 1/ 1년치 같은 수량 표현을 날짜로 잡지 않음
    if (/^\s*년/.test(after2)) {
      return false;
    }

    // 1/1회, 1/2회 등 점검 횟수 표현 방지
    if (/^\s*(회|건|개|명|평|㎡|제곱|만원|원)/.test(after2)) {
      return false;
    }
  }

  if (kind === 'dotMonthDay') {
    if (/콜|차콜|tm|TM/i.test(after2)) {
      return false;
    }

    if (/^\d{1,2}\s*\.\s*\d{1,2}\s*(콜|건|개|명|평|㎡|제곱|만원|원|공장|캠퍼스|동|호)/.test(chunk)) {
      return false;
    }

    if (/^\d{1,2}\s*\.\s*\d{1,2}\s*,\s*\d/.test(chunk)) {
      return false;
    }

    // 1.2 6/8, 1.5콜 류 방지
    if (/^\d{1,2}\s*\.\s*\d{1,2}\s+\d{1,2}\s*[\/월]/.test(chunk)) {
      return false;
    }
  }

  return true;
}


function removeDateFromTextByDateInfoV3_(text, dateInfo) {
  if (!dateInfo) return text;

  const value = String(text || '');

  const before = value.slice(0, dateInfo.startIndex);
  const after = value.slice(dateInfo.endIndex);

  return cleanMemoSegmentV3_(`${before} ${after}`);
}


function isValidMemoDateV3_(year, month, day) {
  if (!year || !month || !day) return false;

  if (year < 2020 || year > 2035) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;

  return true;
}


function pad2V3_(num) {
  return String(num).padStart(2, '0');
}


function monthNameToNumberV3_(monthName) {
  const map = {
    Jan: 1,
    Feb: 2,
    Mar: 3,
    Apr: 4,
    May: 5,
    Jun: 6,
    Jul: 7,
    Aug: 8,
    Sep: 9,
    Oct: 10,
    Nov: 11,
    Dec: 12
  };

  const key = String(monthName || '').slice(0, 3);

  return map[key] || 0;
}


/*************************************************
 * 자료발송 판정 V3
 *************************************************/

function detectMaterialSendMemoV3_(originalText, bodyText) {
  const original = String(originalText || '');
  const body = String(bodyText || '');
  const text = `${original} ${body}`;

  const compact = text
    .replace(/\s+/g, '')
    .replace(/[.,，、]/g, '');

  // 명백히 발송/제출하지 않은 경우는 자료발송 금지
  if (/(제출안함|발송안함|미발송|못보냄|안보냄|보내지않음|보내지않았다|보내지않았음)/.test(compact)) {
    return false;
  }

  // 요청/문의/검토만 있는 경우는 실제 발송 완료가 아님
  if (
    /(견적서?발송요청|견적서?송부요청|견적서?제출요청|견적서?재발송요청|견적서?재송부요청|견적요청|견적서요청|견적문의|견적서문의|견적검토|견적서검토|자료요청|자료발송요청|메일발송요청|발송문의|송부문의|제출문의)/.test(compact) &&
    !/(발송완료|제출완료|송부완료|전송완료|보냈|보냄|보내드림|송부함|발송함|제출함|재발송완료|재송부완료)/.test(compact)
  ) {
    return false;
  }

  const positivePatterns = [
    /견적서?.{0,12}(발송|송부|제출|전송|보냄|보냈|보내드림|발송완료|송부완료|제출완료|전송완료|발송함|송부함|제출함|전송함|재발송|재송부)/,
    /견적.{0,8}(발송|송부|제출|전송|보냄|보냈|보내드림|발송완료|송부완료|제출완료|전송완료|발송함|송부함|제출함|전송함|재발송|재송부)/,
    /메일.{0,8}(발송|송부|전송|보냄|보냈|보내드림|발송완료|송부완료|전송완료|발송함|송부함|전송함)/,
    /자료.{0,8}(발송|송부|전송|보냄|보냈|보내드림|발송완료|송부완료|전송완료|발송함|송부함|전송함)/,
    /안내자료.{0,8}(발송|송부|전송|보냄|보냈|보내드림)/,
    /단가표.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /수행사정보.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /수행사.{0,4}정보.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /샘플보고서.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /보고서샘플.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /비교견적서.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /계약서.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /사업자등록증.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /과업지시서.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /산출내역서.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /점검양식.{0,8}(발송|송부|전송|보냄|보냈|보내드림|제출)/,
    /인바운드.{0,40}(발송|송부|제출|보냄|보냈|보내드림)/,
    /(초급|중급|고급|특급|집합건물|할인견적|수기견적).{0,25}(발송|송부|제출|보냄|보냈|보내드림)/,
    /팩스발송/,
    /FAX발송/i
  ];

  for (let i = 0; i < positivePatterns.length; i++) {
    if (positivePatterns[i].test(compact)) {
      return true;
    }
  }

  return false;
}


/*************************************************
 * 레코드 생성/중복 제거/병합 V3
 *************************************************/

function createMemoRecordV3_(params) {
  const type = params.type || '원문';
  const round = params.round || '';
  const dateInfo = params.dateInfo || null;
  const body = params.body || '';
  const normalizedText = cleanFinalOutputLineV3_(params.normalizedText || '');

  const bodyKey = makeBodyKeyForFinalMemoV3_(body || normalizedText);
  const normalizedKey = makeBodyKeyForFinalMemoV3_(normalizedText);

  const dateKey = dateInfo ? dateInfo.dateKey : '';

  const dedupKey = [
    type,
    dateKey,
    round,
    bodyKey || normalizedKey
  ].join('|');

  return {
    type: type,
    round: round,
    dateInfo: dateInfo,
    dateKey: dateKey,
    body: body,
    bodyKey: bodyKey,
    normalizedKey: normalizedKey,
    normalizedText: normalizedText,
    dedupKey: dedupKey,
    sourceName: params.sourceName || ''
  };
}


function mergeMemoRecordsV3_(originalRecords, memoRecords, stats) {
  const merged = [];

  const allRecords = []
    .concat(originalRecords || [])
    .concat(memoRecords || []);

  for (let i = 0; i < allRecords.length; i++) {
    addMemoRecordUniquelyV3_(merged, allRecords[i], stats);
  }

  return merged;
}


function addMemoRecordUniquelyV3_(records, newRecord, stats) {
  if (!newRecord || !newRecord.normalizedText) return;

  for (let i = 0; i < records.length; i++) {
    const oldRecord = records[i];

    if (isSameMemoRecordV3_(oldRecord, newRecord)) {
      stats.duplicateRemoved++;

      if (
        newRecord.normalizedText.length > oldRecord.normalizedText.length + 10 &&
        newRecord.normalizedKey.length >= oldRecord.normalizedKey.length
      ) {
        records[i] = newRecord;
        stats.replacedByLongerDuplicate++;
      }

      return;
    }
  }

  records.push(newRecord);
}


/**
 * 중복 제거 원칙:
 * - 날짜가 다르면 절대 중복으로 보지 않음
 * - 날짜 있는 메모와 날짜 없는 메모도 함부로 중복 제거하지 않음
 * - 같은 날짜/같은 타입/같은 본문일 때만 제거
 */
function isSameMemoRecordV3_(a, b) {
  if (!a || !b) return false;

  if (!a.normalizedKey || !b.normalizedKey) return false;

  if (a.type !== b.type) return false;

  if (a.dateKey && b.dateKey && a.dateKey !== b.dateKey) {
    return false;
  }

  if ((a.dateKey && !b.dateKey) || (!a.dateKey && b.dateKey)) {
    return false;
  }

  if (a.dedupKey === b.dedupKey) return true;

  if (a.normalizedKey === b.normalizedKey) return true;

  const sameDate = a.dateKey === b.dateKey;
  const sameRound = (a.round || '') === (b.round || '');

  if (sameDate && sameRound && a.bodyKey && b.bodyKey) {
    const shorter = a.bodyKey.length <= b.bodyKey.length ? a.bodyKey : b.bodyKey;
    const longer = a.bodyKey.length > b.bodyKey.length ? a.bodyKey : b.bodyKey;

    if (shorter.length >= 18 && longer.indexOf(shorter) >= 0) {
      return true;
    }
  }

  return false;
}


function makeBodyKeyForFinalMemoV3_(text) {
  return String(text || '')
    .replace(/\[컨택이력\]/g, '')
    .replace(/\[자료발송\]/g, '')
    .replace(/\[다음액션\]/g, '')
    .replace(/\[영업지원요청[^\]]*\]/g, '')
    .replace(/\[\d+\s*차\]/g, '')
    .replace(/\[(전화|문자|메일|방문|기타)\]/g, '')
    .replace(/(?:Mon|Tue|Wed|Thu|Fri|Sat|Sun)\s+(?:Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s+\d{1,2}\s+20\d{2}\s+\d{2}:\d{2}:\d{2}\s+GMT[+-]\d{4}(?:\s+\([^)]+\))?/gi, '')
    .replace(/20\d{2}\s*[.\/-]\s*\d{1,2}\s*[.\/-]\s*\d{1,2}\.?/g, '')
    .replace(/\d{2}\s*[.\/-]\s*\d{1,2}\s*[.\/-]\s*\d{1,2}\.?/g, '')
    .replace(/\d{1,2}\s*월\s*\d{1,2}\s*일?/g, '')
    .replace(/\d{1,2}\s*\/\s*\d{1,2}\s*일?/g, '')
    .replace(/\d{1,2}:\d{2}/g, '')
    .replace(/^\d+\s*[차치]\s*[).]?/g, '')
    .replace(/[()\[\]{}'"“”‘’`~!@#$%^&*_+=<>?:;,.，、/\-|\\\s]/g, '')
    .trim();
}


function cleanFinalOutputLineV3_(text) {
  return String(text || '')
    .replace(/\s+/g, ' ')
    .replace(/\s+\./g, '.')
    .replace(/\[\s+/g, '[')
    .replace(/\s+\]/g, ']')
    .trim();
}


function sortMemoRecordsByDateV3_(records) {
  return records.slice().sort(function (a, b) {
    const ak = a.dateKey || '99.99.99';
    const bk = b.dateKey || '99.99.99';

    if (ak < bk) return -1;
    if (ak > bk) return 1;

    return 0;
  });
}


/*************************************************
 * 헤더/열 처리
 *************************************************/

function getFinalMemoHeaderMapV3_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();

  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const map = {};

  for (let i = 0; i < headers.length; i++) {
    const key = normalizeFinalMemoHeaderV3_(headers[i]);

    if (!key) continue;

    if (!map[key]) {
      map[key] = i + 1;
    }
  }

  return map;
}


function createOrGetFinalMemoColumnV3_(sheet, headerName, headerRow) {
  const headerMap = getFinalMemoHeaderMapV3_(sheet, headerRow);
  const normalizedHeader = normalizeFinalMemoHeaderV3_(headerName);

  if (headerMap[normalizedHeader]) {
    return headerMap[normalizedHeader];
  }

  const lastCol = sheet.getLastColumn();

  sheet.insertColumnAfter(lastCol);

  const newCol = lastCol + 1;

  const headerCell = sheet.getRange(headerRow, newCol);
  headerCell.setValue(headerName);

  if (lastCol >= 1) {
    const sourceHeader = sheet.getRange(headerRow, lastCol);
    sourceHeader.copyTo(headerCell, { formatOnly: true });
  }

  sheet.setColumnWidth(newCol, 420);

  return newCol;
}


function normalizeFinalMemoHeaderV3_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .trim();
}


/*************************************************
 * 구조 판정
 *************************************************/

function isStrongStructuredLineV3_(text) {
  const value = String(text || '').trim();

  if (/^\[컨택이력\]/.test(value)) return true;
  if (/^\[다음액션\]/.test(value)) return true;
  if (/^\[\d{2}\.\d{2}\.\d{2}(?:\s+\d{1,2}:\d{2})?\]\s*\[영업지원요청/.test(value)) return true;

  return false;
}
