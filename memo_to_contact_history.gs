/*************************************************
 * 최종 메모 → 다른 파일의 컨택이력_DB 시트로 이관
 *
 * 실행 위치:
 * - 영업관리대장 Apps Script
 *
 * 소스:
 * - 영업관리대장 현재/지정 시트
 * - '마스터시트 메모 최종 업데이트본' 열
 *
 * 타겟:
 * - 다른 스프레드시트의 '컨택이력_DB' 시트
 *
 * 주요 기능:
 * - 헤더명 기준으로 읽기/쓰기
 * - 컨택이력 / 자료발송 / 다음액션 / 영업지원요청 분리
 * - 작성일시/차수/연락수단/작성자/태그/상태 추출
 * - 클라이언트요청ID 기준 중복 append 방지
 * - 마스터메모반영 = Y
 *************************************************/

const CONTACT_HISTORY_MIGRATION_CONFIG = {
  /*************************************************
   * 여기에 타겟 파일 ID 입력
   *************************************************/
  TARGET_SPREADSHEET_ID: '12YmhhkhMCVvpBezJOzer_HlAzscLbNsumX-bidv1KWw',

  TARGET_SHEET_NAME: '컨택이력_DB',

  // 비워두면 현재 활성 시트 기준
  SOURCE_SHEET_NAME: '',

  // 영업관리대장 구조 기준
  SOURCE_HEADER_ROW: 2,
  SOURCE_DATA_START_ROW: 3,

  // 헤더가 다를 때도 대응
  FINAL_MEMO_HEADERS: [
    '마스터시트 메모 최종 업데이트본'
  ],

  CUSTOMER_NO_HEADERS: [
    '고객번호'
  ],

  COMPANY_HEADERS: [
    '회사명',
    '고객사명',
    '고객명',
    '업체명'
  ],

  TARGET_HEADER_ROW: 1,

  DEFAULT_AUTHOR: '메모이관',
  MASTER_MEMO_REFLECTED_VALUE: 'Y',

  SKIP_ROWS_WITHOUT_CUSTOMER_NO: true,
  SKIP_DUPLICATE_BY_CLIENT_REQUEST_ID: true,

  PREVIEW_SHEET_NAME: '__컨택이력_이관미리보기',
  LOG_SHEET_NAME: '__컨택이력_이관로그'
};


const CONTACT_HISTORY_DB_REQUIRED_HEADERS = [
  '이력ID',
  '고객번호',
  '회사명',
  '마스터행',
  '작성일시',
  '작성자',
  '기록구분',
  '차수',
  '연락수단',
  '태그',
  '계약진행상태',
  '컨택내용',
  '특이사항',
  '다음액션',
  '다음액션일시',
  '다음액션태그',
  '다음액션담당자',
  '마스터메모반영',
  '클라이언트요청ID'
];


/**
 * 1단계: 미리보기만 생성
 * 타겟 DB에는 쓰지 않음.
 */
function previewFinalMemoToContactHistoryDb() {
  runFinalMemoToContactHistoryDbMigration_({
    previewOnly: true
  });
}


/**
 * 2단계: 실제 이관
 * 다른 파일의 컨택이력_DB에 append.
 */
function migrateFinalMemoToContactHistoryDb() {
  runFinalMemoToContactHistoryDbMigration_({
    previewOnly: false
  });
}


/*************************************************
 * 메인 러너
 *************************************************/

function runFinalMemoToContactHistoryDbMigration_(options) {
  options = options || {};

  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  if (!cfg.TARGET_SPREADSHEET_ID || cfg.TARGET_SPREADSHEET_ID === '여기에_컨택이력_DB_있는_다른파일_ID') {
    throw new Error('CONTACT_HISTORY_MIGRATION_CONFIG.TARGET_SPREADSHEET_ID에 타겟 스프레드시트 ID를 입력해야 합니다.');
  }

  const sourceSs = SpreadsheetApp.getActiveSpreadsheet();
  const sourceSheet = getSourceSheetForContactHistoryMigration_(sourceSs);
  const sourceInfo = getSourceSheetInfoForContactHistoryMigration_(sourceSheet);

  const targetSs = SpreadsheetApp.openById(cfg.TARGET_SPREADSHEET_ID);
  const targetSheet = getOrCreateTargetContactHistorySheet_(targetSs);

  ensureTargetContactHistoryHeaders_(targetSheet);

  const targetHeaderInfo = getTargetHeaderInfo_(targetSheet);
  const existingClientRequestIds = loadExistingClientRequestIds_(targetSheet, targetHeaderInfo);

  const result = buildContactHistoryRowsFromFinalMemo_(
    sourceSheet,
    sourceInfo,
    targetHeaderInfo,
    existingClientRequestIds,
    options
  );

  if (options.previewOnly) {
    writeContactHistoryMigrationPreview_(sourceSs, result.outputHeaders, result.outputRows, result.stats);
    writeContactHistoryMigrationLog_(sourceSs, result.stats, 'PREVIEW');
    notifyContactHistoryMigration_(
      [
        '컨택이력_DB 이관 미리보기 생성 완료',
        '',
        `소스 시트: ${sourceSheet.getName()}`,
        `대상 예정 행: ${result.stats.readyRows}건`,
        `중복 제외: ${result.stats.duplicateSkipped}건`,
        `고객번호 없음 제외: ${result.stats.noCustomerSkipped}건`,
        `빈/무의미 라인 제외: ${result.stats.emptyLineSkipped}건`,
        '',
        `미리보기 시트: ${cfg.PREVIEW_SHEET_NAME}`
      ].join('\n')
    );
    return;
  }

  if (result.outputRows.length === 0) {
    writeContactHistoryMigrationLog_(sourceSs, result.stats, 'MIGRATE_NO_ROWS');
    notifyContactHistoryMigration_(
      [
        '이관할 신규 행이 없습니다.',
        '',
        `중복 제외: ${result.stats.duplicateSkipped}건`,
        `고객번호 없음 제외: ${result.stats.noCustomerSkipped}건`,
        `빈/무의미 라인 제외: ${result.stats.emptyLineSkipped}건`
      ].join('\n')
    );
    return;
  }

  appendRowsToTargetContactHistoryDb_(targetSheet, result.outputRows, result.outputHeaders);

  writeContactHistoryMigrationLog_(sourceSs, result.stats, 'MIGRATE_DONE');

  notifyContactHistoryMigration_(
    [
      '컨택이력_DB 이관 완료',
      '',
      `소스 시트: ${sourceSheet.getName()}`,
      `타겟 파일: ${targetSs.getName()}`,
      `타겟 시트: ${targetSheet.getName()}`,
      '',
      `append 완료: ${result.stats.readyRows}건`,
      `중복 제외: ${result.stats.duplicateSkipped}건`,
      `고객번호 없음 제외: ${result.stats.noCustomerSkipped}건`,
      `빈/무의미 라인 제외: ${result.stats.emptyLineSkipped}건`,
      '',
      `컨택이력: ${result.stats.byType['컨택이력'] || 0}건`,
      `자료발송: ${result.stats.byType['자료발송'] || 0}건`,
      `다음액션: ${result.stats.byType['다음액션'] || 0}건`,
      `영업지원요청: ${result.stats.byType['영업지원요청'] || 0}건`,
      `기타메모: ${result.stats.byType['기타메모'] || 0}건`
    ].join('\n')
  );
}


/*************************************************
 * 소스/타겟 시트 처리
 *************************************************/

function getSourceSheetForContactHistoryMigration_(sourceSs) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  if (cfg.SOURCE_SHEET_NAME) {
    const sheet = sourceSs.getSheetByName(cfg.SOURCE_SHEET_NAME);

    if (!sheet) {
      throw new Error(`소스 시트를 찾을 수 없습니다: ${cfg.SOURCE_SHEET_NAME}`);
    }

    return sheet;
  }

  return sourceSs.getActiveSheet();
}


function getSourceSheetInfoForContactHistoryMigration_(sheet) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  let headerRow = cfg.SOURCE_HEADER_ROW;
  let headerMap = getHeaderMapFromRow_(sheet, headerRow);

  let finalMemoCol = findFirstExistingHeaderColumn_(headerMap, cfg.FINAL_MEMO_HEADERS);

  // 혹시 헤더행이 다르면 1~5행 자동 탐색
  if (!finalMemoCol) {
    for (let r = 1; r <= Math.min(5, sheet.getLastRow()); r++) {
      const testMap = getHeaderMapFromRow_(sheet, r);
      const testFinalMemoCol = findFirstExistingHeaderColumn_(testMap, cfg.FINAL_MEMO_HEADERS);

      if (testFinalMemoCol) {
        headerRow = r;
        headerMap = testMap;
        finalMemoCol = testFinalMemoCol;
        break;
      }
    }
  }

  if (!finalMemoCol) {
    throw new Error(`소스 시트에서 '${cfg.FINAL_MEMO_HEADERS.join(' 또는 ')}' 헤더를 찾을 수 없습니다.`);
  }

  const customerNoCol = findFirstExistingHeaderColumn_(headerMap, cfg.CUSTOMER_NO_HEADERS);
  const companyCol = findFirstExistingHeaderColumn_(headerMap, cfg.COMPANY_HEADERS);

  if (!customerNoCol) {
    throw new Error(`소스 시트에서 고객번호 헤더를 찾을 수 없습니다: ${cfg.CUSTOMER_NO_HEADERS.join(', ')}`);
  }

  if (!companyCol) {
    throw new Error(`소스 시트에서 회사명/고객사명 헤더를 찾을 수 없습니다: ${cfg.COMPANY_HEADERS.join(', ')}`);
  }

  const dataStartRow = headerRow === cfg.SOURCE_HEADER_ROW
    ? cfg.SOURCE_DATA_START_ROW
    : headerRow + 1;

  return {
    headerRow: headerRow,
    dataStartRow: dataStartRow,
    headerMap: headerMap,
    finalMemoCol: finalMemoCol,
    customerNoCol: customerNoCol,
    companyCol: companyCol
  };
}


function getOrCreateTargetContactHistorySheet_(targetSs) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  let sheet = targetSs.getSheetByName(cfg.TARGET_SHEET_NAME);

  if (!sheet) {
    sheet = targetSs.insertSheet(cfg.TARGET_SHEET_NAME);
  }

  return sheet;
}


function ensureTargetContactHistoryHeaders_(targetSheet) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;
  const headerRow = cfg.TARGET_HEADER_ROW;

  const lastCol = Math.max(targetSheet.getLastColumn(), CONTACT_HISTORY_DB_REQUIRED_HEADERS.length);

  let headers = targetSheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const hasAnyHeader = headers.some(h => String(h || '').trim() !== '');

  if (!hasAnyHeader) {
    targetSheet
      .getRange(headerRow, 1, 1, CONTACT_HISTORY_DB_REQUIRED_HEADERS.length)
      .setValues([CONTACT_HISTORY_DB_REQUIRED_HEADERS]);

    targetSheet.setFrozenRows(headerRow);
    return;
  }

  const normalizedExisting = {};

  for (let i = 0; i < headers.length; i++) {
    const key = normalizeHeaderKey_(headers[i]);
    if (key) {
      normalizedExisting[key] = i + 1;
    }
  }

  let appendCol = targetSheet.getLastColumn();

  for (let i = 0; i < CONTACT_HISTORY_DB_REQUIRED_HEADERS.length; i++) {
    const header = CONTACT_HISTORY_DB_REQUIRED_HEADERS[i];
    const key = normalizeHeaderKey_(header);

    if (!normalizedExisting[key]) {
      appendCol++;
      targetSheet.getRange(headerRow, appendCol).setValue(header);
      normalizedExisting[key] = appendCol;
    }
  }

  targetSheet.setFrozenRows(headerRow);
}


function getTargetHeaderInfo_(targetSheet) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;
  const headerRow = cfg.TARGET_HEADER_ROW;
  const lastCol = targetSheet.getLastColumn();

  const headers = targetSheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const headerMap = {};

  for (let i = 0; i < headers.length; i++) {
    const key = normalizeHeaderKey_(headers[i]);

    if (!key) continue;

    if (!headerMap[key]) {
      headerMap[key] = i + 1;
    }
  }

  return {
    headerRow: headerRow,
    lastCol: lastCol,
    headers: headers,
    headerMap: headerMap
  };
}


function getHeaderMapFromRow_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();

  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const map = {};

  for (let i = 0; i < headers.length; i++) {
    const key = normalizeHeaderKey_(headers[i]);

    if (!key) continue;

    if (!map[key]) {
      map[key] = i + 1;
    }
  }

  return map;
}


function findFirstExistingHeaderColumn_(headerMap, headerNames) {
  for (let i = 0; i < headerNames.length; i++) {
    const key = normalizeHeaderKey_(headerNames[i]);

    if (headerMap[key]) {
      return headerMap[key];
    }
  }

  return 0;
}


function normalizeHeaderKey_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .trim();
}


/*************************************************
 * 이관 행 생성
 *************************************************/

function buildContactHistoryRowsFromFinalMemo_(
  sourceSheet,
  sourceInfo,
  targetHeaderInfo,
  existingClientRequestIds,
  options
) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  const stats = {
    sourceRows: 0,
    sourceCellsWithMemo: 0,
    sourceLines: 0,
    readyRows: 0,
    duplicateSkipped: 0,
    noCustomerSkipped: 0,
    emptyLineSkipped: 0,
    parseFailedSkipped: 0,
    byType: {}
  };

  const outputRows = [];
  const usedClientRequestIds = {};

  const lastRow = sourceSheet.getLastRow();
  const lastCol = sourceSheet.getLastColumn();

  if (lastRow < sourceInfo.dataStartRow) {
    return {
      outputHeaders: targetHeaderInfo.headers,
      outputRows: [],
      stats: stats
    };
  }

  const numRows = lastRow - sourceInfo.dataStartRow + 1;

  const values = sourceSheet
    .getRange(sourceInfo.dataStartRow, 1, numRows, lastCol)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const sourceRowNo = sourceInfo.dataStartRow + i;
    const row = values[i];

    const customerNo = cleanText_(row[sourceInfo.customerNoCol - 1]);
    const companyName = cleanText_(row[sourceInfo.companyCol - 1]);
    const finalMemo = String(row[sourceInfo.finalMemoCol - 1] || '').trim();

    stats.sourceRows++;

    if (!finalMemo) continue;

    stats.sourceCellsWithMemo++;

    if (cfg.SKIP_ROWS_WITHOUT_CUSTOMER_NO && !customerNo) {
      stats.noCustomerSkipped++;
      continue;
    }

    const lines = splitFinalMemoCellToLines_(finalMemo);

    for (let j = 0; j < lines.length; j++) {
      const line = cleanText_(lines[j]);

      stats.sourceLines++;

      if (isMeaninglessMigrationLine_(line)) {
        stats.emptyLineSkipped++;
        continue;
      }

      const parsed = parseFinalMemoLineForContactHistory_(line);

      if (!parsed || !parsed.recordType) {
        stats.parseFailedSkipped++;
        continue;
      }

      if (isMeaninglessMigrationLine_(parsed.contactContent) && parsed.recordType !== '다음액션') {
        stats.emptyLineSkipped++;
        continue;
      }

      const clientRequestId = createMigrationClientRequestId_(customerNo, sourceRowNo, line);

      if (
        cfg.SKIP_DUPLICATE_BY_CLIENT_REQUEST_ID &&
        (existingClientRequestIds[clientRequestId] || usedClientRequestIds[clientRequestId])
      ) {
        stats.duplicateSkipped++;
        continue;
      }

      usedClientRequestIds[clientRequestId] = true;

      const historyId = createMigrationHistoryId_(
        parsed.writtenAt,
        customerNo,
        sourceRowNo,
        line
      );

      const rowObject = createContactHistoryRowObject_({
        historyId: historyId,
        customerNo: customerNo,
        companyName: companyName,
        sourceRowNo: sourceRowNo,
        parsed: parsed,
        originalLine: line,
        clientRequestId: clientRequestId
      });

      const outputRow = makeTargetRowArray_(rowObject, targetHeaderInfo.headers);

      outputRows.push(outputRow);

      stats.readyRows++;
      stats.byType[parsed.recordType] = (stats.byType[parsed.recordType] || 0) + 1;
    }
  }

  return {
    outputHeaders: targetHeaderInfo.headers,
    outputRows: outputRows,
    stats: stats
  };
}


function splitFinalMemoCellToLines_(memoText) {
  return String(memoText || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .split('\n')
    .map(line => cleanText_(line))
    .filter(line => line !== '');
}


/*************************************************
 * 한 줄 파싱
 *************************************************/

function parseFinalMemoLineForContactHistory_(line) {
  const originalLine = cleanText_(line);

  if (!originalLine) return null;

  // [26.06.21 10:48] [영업지원요청 접수 #533] ...
  const supportRequest = parseSupportRequestLine_(originalLine);

  if (supportRequest) {
    return supportRequest;
  }

  let working = originalLine;

  const tagInfo = extractLeadingBracketTags_(working);

  working = tagInfo.remainingText;

  let recordType = tagInfo.recordType || '';
  let round = tagInfo.round || '';
  let contactMethod = tagInfo.contactMethod || '';
  let tags = tagInfo.tags || [];

  const leadingDateInfo = extractLeadingDateInfo_(working);

  let writtenAt = null;

  if (leadingDateInfo) {
    writtenAt = leadingDateInfo.date;
    working = cleanText_(working.slice(leadingDateInfo.endIndex));
  }

  const roundInfo = extractLeadingRoundInfo_(working);

  if (roundInfo.round && !round) {
    round = roundInfo.round;
    working = roundInfo.remainingText;
  }

  const authorInfo = extractAuthorFromTail_(working);
  const author = authorInfo.author;
  working = authorInfo.text;

  const body = cleanBodyText_(working);

  if (!recordType) {
    if (detectMaterialSendRecord_(originalLine)) {
      recordType = '자료발송';
    } else if (writtenAt || round || contactMethod) {
      recordType = '컨택이력';
    } else {
      recordType = '기타메모';
    }
  }

  if (!contactMethod) {
    contactMethod = detectContactMethod_(originalLine, recordType);
  }

  const autoTags = detectAutoTags_(originalLine, recordType);
  tags = mergeStringList_(tags, autoTags);

  const contractStatus = detectContractStatus_(originalLine);

  let contactContent = body;
  let nextAction = '';
  let nextActionAt = '';
  let nextActionTag = '';
  let nextActionAssignee = '';

  if (recordType === '다음액션') {
    nextAction = body;
    nextActionAt = writtenAt || '';
    nextActionTag = contactMethod || detectNextActionTag_(body);
    nextActionAssignee = author || '';
    contactContent = '';
  }

  return {
    recordType: recordType,
    writtenAt: writtenAt || '',
    author: author || '',
    round: round || '',
    contactMethod: contactMethod || '',
    tags: tags.join(', '),
    contractStatus: contractStatus || '',
    contactContent: contactContent,
    specialNote: '',
    nextAction: nextAction,
    nextActionAt: nextActionAt,
    nextActionTag: nextActionTag,
    nextActionAssignee: nextActionAssignee,
    originalLine: originalLine
  };
}


function parseSupportRequestLine_(line) {
  const match = String(line || '').match(
    /^\[(\d{2}\.\d{1,2}\.\d{1,2}(?:\s+\d{1,2}:\d{2}(?::\d{2})?)?)\]\s*\[([^\]]*영업지원요청[^\]]*)\]\s*(.*)$/
  );

  if (!match) return null;

  const dateInfo = parseDateTextToDate_(match[1]);
  const supportTag = cleanText_(match[2]);
  const bodyRaw = cleanBodyText_(match[3]);

  const authorInfo = extractAuthorFromTail_(bodyRaw);
  const body = authorInfo.text;

  const autoTags = detectAutoTags_(body, '영업지원요청');

  if (supportTag) {
    autoTags.unshift(supportTag);
  }

  return {
    recordType: '영업지원요청',
    writtenAt: dateInfo ? dateInfo.date : '',
    author: authorInfo.author || '',
    round: '',
    contactMethod: '',
    tags: uniqueStringList_(autoTags).join(', '),
    contractStatus: detectContractStatus_(body),
    contactContent: body,
    specialNote: '',
    nextAction: '',
    nextActionAt: '',
    nextActionTag: '',
    nextActionAssignee: '',
    originalLine: line
  };
}


function extractLeadingBracketTags_(text) {
  let working = String(text || '').trim();

  let recordType = '';
  let round = '';
  let contactMethod = '';
  const tags = [];

  while (true) {
    const match = working.match(/^\[([^\]]+)\]\s*/);

    if (!match) break;

    const tag = cleanText_(match[1]);

    working = working.slice(match[0].length).trim();

    if (tag === '컨택이력') {
      recordType = '컨택이력';
      continue;
    }

    if (tag === '자료발송') {
      recordType = '자료발송';
      continue;
    }

    if (tag === '다음액션') {
      recordType = '다음액션';
      continue;
    }

    if (/^\d+\s*차$/.test(tag)) {
      round = tag.replace(/\s+/g, '');
      continue;
    }

    if (isContactMethodTag_(tag)) {
      contactMethod = normalizeContactMethod_(tag);
      continue;
    }

    tags.push(tag);
  }

  return {
    recordType: recordType,
    round: round,
    contactMethod: contactMethod,
    tags: tags,
    remainingText: working
  };
}


function extractLeadingRoundInfo_(text) {
  let working = String(text || '').trim();

  const match = working.match(/^(?:\(?\s*)?(\d+)\s*[차치]\s*[).]?\s*/);

  if (!match) {
    return {
      round: '',
      remainingText: working
    };
  }

  const round = `${Number(match[1])}차`;

  working = working.slice(match[0].length).trim();

  return {
    round: round,
    remainingText: working
  };
}


/*************************************************
 * 날짜 파싱
 *************************************************/

function extractLeadingDateInfo_(text) {
  const value = String(text || '').trim();

  if (!value) return null;

  const patterns = [
    {
      regex: /^((?:20)?\d{2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{1,2})\.?\s*(\d{1,2}:\d{2}(?::\d{2})?)?/,
      build: function (m) {
        let year = Number(m[1]);

        if (year < 100) {
          year = 2000 + year;
        }

        return {
          year: year,
          month: Number(m[2]),
          day: Number(m[3]),
          time: m[4] || ''
        };
      }
    },
    {
      regex: /^(\d{1,2})\s*[.\/]\s*(\d{1,2})\.?\s*(\d{1,2}:\d{2}(?::\d{2})?)?/,
      build: function (m) {
        return {
          year: 2026,
          month: Number(m[1]),
          day: Number(m[2]),
          time: m[3] || ''
        };
      }
    },
    {
      regex: /^(\d{1,2})\s*월\s*(\d{1,2})\s*일?\s*(\d{1,2}:\d{2}(?::\d{2})?)?/,
      build: function (m) {
        return {
          year: 2026,
          month: Number(m[1]),
          day: Number(m[2]),
          time: m[3] || ''
        };
      }
    }
  ];

  for (let i = 0; i < patterns.length; i++) {
    const pattern = patterns[i];
    const match = value.match(pattern.regex);

    if (!match) continue;

    const after = value.slice(match[0].length, match[0].length + 12);

    if (!isSafeLeadingDateContext_(match[0], after)) {
      continue;
    }

    const parts = pattern.build(match);

    if (!isValidDateParts_(parts.year, parts.month, parts.day)) {
      continue;
    }

    const date = makeDateFromParts_(parts.year, parts.month, parts.day, parts.time);

    return {
      date: date,
      endIndex: match[0].length
    };
  }

  return null;
}


function parseDateTextToDate_(dateText) {
  const value = String(dateText || '').trim();

  const match = value.match(/^((?:20)?\d{2})\s*[.\/-]\s*(\d{1,2})\s*[.\/-]\s*(\d{1,2})\.?\s*(\d{1,2}:\d{2}(?::\d{2})?)?/);

  if (!match) return null;

  let year = Number(match[1]);

  if (year < 100) {
    year = 2000 + year;
  }

  const month = Number(match[2]);
  const day = Number(match[3]);
  const time = match[4] || '';

  if (!isValidDateParts_(year, month, day)) return null;

  return {
    date: makeDateFromParts_(year, month, day, time)
  };
}


function makeDateFromParts_(year, month, day, timeText) {
  let hour = 0;
  let minute = 0;
  let second = 0;

  if (timeText) {
    const tm = String(timeText).match(/^(\d{1,2}):(\d{2})(?::(\d{2}))?/);

    if (tm) {
      hour = Number(tm[1]);
      minute = Number(tm[2]);
      second = tm[3] ? Number(tm[3]) : 0;
    }
  }

  return new Date(year, month - 1, day, hour, minute, second);
}


function isValidDateParts_(year, month, day) {
  if (!year || !month || !day) return false;

  if (year < 2020 || year > 2035) return false;
  if (month < 1 || month > 12) return false;
  if (day < 1 || day > 31) return false;

  return true;
}


function isSafeLeadingDateContext_(matchedDateText, afterText) {
  const matched = String(matchedDateText || '');
  const after = String(afterText || '');

  // 1.2공장, 1/ 1년치, 1.5콜 같은 숫자/명칭 오인식 방지
  if (/^\s*(공장|창고|동|층|호|년치|개동|개소|콜|차콜|만원|원|㎡|제곱|평)/.test(after)) {
    return false;
  }

  if (/^\s*\d/.test(after) && /^\d{1,2}\s*[.\/]\s*\d{1,2}/.test(matched)) {
    return false;
  }

  return true;
}


/*************************************************
 * 작성자 추출
 *************************************************/

function extractAuthorFromTail_(text) {
  let value = String(text || '').trim();

  const match = value.match(/\(([^()]{1,20})\)\s*$/);

  if (!match) {
    return {
      author: '',
      text: value
    };
  }

  const raw = cleanText_(match[1]);
  const author = normalizeAuthorName_(raw);

  if (!author) {
    return {
      author: '',
      text: value
    };
  }

  value = value.slice(0, match.index).trim();

  return {
    author: author,
    text: value
  };
}


function normalizeAuthorName_(rawName) {
  const value = cleanText_(rawName);

  if (!value) return '';

  const exclusion = {
    '우호적': true,
    '호의적': true,
    '관심많음': true,
    '관심 많음': true,
    '검토중': true,
    '미수신': true,
    '완': true,
    '완료': true,
    '거절': true,
    '문자': true,
    '전화': true,
    '메일': true,
    '팩스': true,
    '인콜': true,
    '아웃콜': true,
    'TM확인': true
  };

  if (exclusion[value]) return '';

  const aliasMap = {
    '문': '문형진',
    '최': '최보람',
    '이': '이옥희',
    '서': '김서하',
    '경아': '김경아',
    '보람': '최보람',
    '새봄': '박새봄',
    '수원': '방수원',
    '방수원': '방수원',
    '김경아': '김경아',
    '최보람': '최보람',
    '박새봄': '박새봄',
    '김서하': '김서하',
    '이옥희': '이옥희',
    '문형진': '문형진',
    '유현희': '유현희',
    '라유화': '라유화',
    '박원경': '박원경'
  };

  if (aliasMap[value]) {
    return aliasMap[value];
  }

  if (/^[가-힣]{2,4}$/.test(value)) {
    return value;
  }

  if (/^[가-힣]{2,6}\s*(관리자|서무|대리|과장|팀장|책임|차장|부장)$/.test(value)) {
    return value;
  }

  return '';
}


/*************************************************
 * 자동 분류/태그/상태
 *************************************************/

function detectMaterialSendRecord_(text) {
  const compact = String(text || '')
    .replace(/\s+/g, '')
    .replace(/[.,，、]/g, '');

  if (/(발송안함|제출안함|미발송|안보냄|보내지않음|제출하지않음)/.test(compact)) {
    return false;
  }

  const positivePatterns = [
    /자료발송/,
    /메일발송/,
    /견적서?.{0,12}(발송|송부|제출|전송|보냄|보냈|보내드림|재발송|재송부|발송완료|제출완료)/,
    /견적.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림|재발송|재송부|발송완료|제출완료)/,
    /단가표.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /수행사.{0,10}(정보|자료).{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /샘플보고서.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /비교견적서?.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /계약서.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /사업자등록증.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /과업지시서.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
    /산출내역서.{0,10}(발송|송부|제출|전송|보냄|보냈|보내드림)/,
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


function detectContactMethod_(text, recordType) {
  const value = String(text || '');

  if (/\[전화\]/.test(value)) return '전화';
  if (/\[문자\]/.test(value)) return '문자';
  if (/\[메일\]/.test(value)) return '메일';
  if (/\[방문\]/.test(value)) return '방문';
  if (/\[팩스\]/.test(value)) return '팩스';

  if (/팩스|FAX/i.test(value)) return '팩스';
  if (/카톡|문자|SMS/i.test(value)) return '문자';
  if (/메일|이메일|mail/i.test(value)) return '메일';
  if (/방문|미팅|내방|화상회의/.test(value)) return '방문';
  if (/인콜/.test(value)) return '인콜';
  if (/아웃콜/.test(value)) return '아웃콜';
  if (/전화|통화|유선|부재|미수신|내선/.test(value)) return '전화';

  if (recordType === '자료발송') {
    return '메일';
  }

  return '';
}


function isContactMethodTag_(tag) {
  const value = cleanText_(tag);

  return [
    '전화',
    '문자',
    '메일',
    '방문',
    '팩스',
    '인콜',
    '아웃콜',
    '기타',
    '테스트발송'
  ].indexOf(value) >= 0;
}


function normalizeContactMethod_(tag) {
  const value = cleanText_(tag);

  if (value === '테스트발송') return '테스트발송';

  return value;
}


function detectAutoTags_(text, recordType) {
  const value = String(text || '');
  const tags = [];

  const tagRules = [
    ['견적서', /견적서|견적발송|견적제출|견적재발송|수기견적/],
    ['용역신청서', /용역신청서/],
    ['선임신고서', /선임신고서/],
    ['위임장', /위임장/],
    ['약관', /약관/],
    ['안내문', /안내문|안내자료|안내장/],
    ['법령요약', /법령요약|법령|공문내용|의무화/],
    ['수행사정보', /수행사\s*정보|수행사자료|수행사 자료/],
    ['샘플보고서', /샘플보고서|보고서샘플/],
    ['비교견적서', /비교견적/],
    ['계약서', /계약서/],
    ['사업자등록증', /사업자등록증/],
    ['과업지시서', /과업지시서/],
    ['산출내역서', /산출내역서/],
    ['단가표', /단가표/],
    ['팩스', /팩스|FAX/i],
    ['인바운드', /인바운드|인콜/],
    ['할인견적', /할인견적|할인가|할인율|네고/],
    ['계약진행', /계약진행|계약 진행|계약하|계약서|용역신청서 보내/],
    ['타사계약', /타사계약|다른업체|다른 업체|타업체|다른곳|다른 곳/],
    ['부재', /부재|미수신|전화안받|연결안됨/]
  ];

  for (let i = 0; i < tagRules.length; i++) {
    if (tagRules[i][1].test(value)) {
      tags.push(tagRules[i][0]);
    }
  }

  if (recordType && tags.indexOf(recordType) < 0) {
    tags.unshift(recordType);
  }

  return uniqueStringList_(tags);
}


function detectContractStatus_(text) {
  const value = String(text || '');

  if (/타사계약완료|타사선정완료|타업체.*계약|다른업체.*계약|다른곳.*계약|수주실패|탈락/.test(value)) {
    return '수주실패';
  }

  if (/계약완료|발주메일 발송완료|발주완료|나라장터.*계약|계약서.*작성완료/.test(value)) {
    return '계약완료';
  }

  if (/계약진행|계약 진행|계약하신다고|계약하겠|용역신청서.*보내|사업자등록증.*보내/.test(value)) {
    return '계약진행';
  }

  if (/검토중|비교중|결정.*예정|보고.*중|품의|결재/.test(value)) {
    return '검토중';
  }

  if (/부재|미수신|전화안받|자리에 없음|외근|휴가|연차/.test(value)) {
    return '부재';
  }

  if (/거절|하지마|전화하지|관심없|필요없|안한다|안함/.test(value)) {
    return '거절';
  }

  if (/자체선임|직접선임|자체수행|직접수행/.test(value)) {
    return '자체진행';
  }

  if (/보류|중단/.test(value)) {
    return '보류';
  }

  return '';
}


function detectNextActionTag_(text) {
  const value = String(text || '');

  if (/전화|통화|유선/.test(value)) return '전화';
  if (/문자|카톡/.test(value)) return '문자';
  if (/메일|이메일/.test(value)) return '메일';
  if (/방문|미팅/.test(value)) return '방문';

  return '';
}


/*************************************************
 * 타겟 행 오브젝트 생성
 *************************************************/

function createContactHistoryRowObject_(params) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  const parsed = params.parsed;

  const rowObject = {};

  rowObject['이력ID'] = params.historyId;
  rowObject['고객번호'] = params.customerNo;
  rowObject['회사명'] = params.companyName;
  rowObject['마스터행'] = params.sourceRowNo;
  rowObject['작성일시'] = parsed.writtenAt || '';
  rowObject['작성자'] = parsed.author || cfg.DEFAULT_AUTHOR;
  rowObject['기록구분'] = parsed.recordType || '기타메모';
  rowObject['차수'] = parsed.round || '';
  rowObject['연락수단'] = parsed.contactMethod || '';
  rowObject['태그'] = parsed.tags || '';
  rowObject['계약진행상태'] = parsed.contractStatus || '';
  rowObject['컨택내용'] = parsed.contactContent || '';
  rowObject['특이사항'] = parsed.specialNote || '';
  rowObject['다음액션'] = parsed.nextAction || '';
  rowObject['다음액션일시'] = parsed.nextActionAt || '';
  rowObject['다음액션태그'] = parsed.nextActionTag || '';
  rowObject['다음액션담당자'] = parsed.nextActionAssignee || '';
  rowObject['마스터메모반영'] = cfg.MASTER_MEMO_REFLECTED_VALUE;
  rowObject['클라이언트요청ID'] = params.clientRequestId;

  return rowObject;
}


function makeTargetRowArray_(rowObject, targetHeaders) {
  const row = [];

  for (let i = 0; i < targetHeaders.length; i++) {
    const header = String(targetHeaders[i] || '').trim();

    if (!header) {
      row.push('');
      continue;
    }

    row.push(rowObject[header] !== undefined ? rowObject[header] : '');
  }

  return row;
}


/*************************************************
 * append / preview / log
 *************************************************/

function appendRowsToTargetContactHistoryDb_(targetSheet, outputRows, outputHeaders) {
  if (!outputRows || outputRows.length === 0) return;

  const startRow = Math.max(
    targetSheet.getLastRow() + 1,
    CONTACT_HISTORY_MIGRATION_CONFIG.TARGET_HEADER_ROW + 1
  );

  targetSheet
    .getRange(startRow, 1, outputRows.length, outputHeaders.length)
    .setValues(outputRows);

  const targetHeaderInfo = getTargetHeaderInfo_(targetSheet);

  const writtenAtCol = targetHeaderInfo.headerMap[normalizeHeaderKey_('작성일시')];
  const nextActionAtCol = targetHeaderInfo.headerMap[normalizeHeaderKey_('다음액션일시')];

  if (writtenAtCol) {
    targetSheet
      .getRange(startRow, writtenAtCol, outputRows.length, 1)
      .setNumberFormat('yyyy. m. d aaa h:mm:ss');
  }

  if (nextActionAtCol) {
    targetSheet
      .getRange(startRow, nextActionAtCol, outputRows.length, 1)
      .setNumberFormat('yyyy. m. d aaa h:mm:ss');
  }
}


function writeContactHistoryMigrationPreview_(sourceSs, headers, rows, stats) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  let sheet = sourceSs.getSheetByName(cfg.PREVIEW_SHEET_NAME);

  if (!sheet) {
    sheet = sourceSs.insertSheet(cfg.PREVIEW_SHEET_NAME);
  }

  sheet.clearContents();

  const infoRows = [
    ['미리보기 생성일시', new Date()],
    ['이관 예정 행', stats.readyRows],
    ['중복 제외', stats.duplicateSkipped],
    ['고객번호 없음 제외', stats.noCustomerSkipped],
    ['빈/무의미 라인 제외', stats.emptyLineSkipped],
    [''],
    ['아래부터 실제 append 예정 데이터']
  ];

  sheet.getRange(1, 1, infoRows.length, 2).setValues(
    infoRows.map(r => [r[0] || '', r[1] || ''])
  );

  const headerStartRow = infoRows.length + 2;

  sheet
    .getRange(headerStartRow, 1, 1, headers.length)
    .setValues([headers]);

  if (rows.length > 0) {
    sheet
      .getRange(headerStartRow + 1, 1, rows.length, headers.length)
      .setValues(rows);
  }

  sheet.setFrozenRows(headerStartRow);
  sheet.autoResizeColumns(1, Math.min(headers.length, 10));
}


function writeContactHistoryMigrationLog_(sourceSs, stats, mode) {
  const cfg = CONTACT_HISTORY_MIGRATION_CONFIG;

  let sheet = sourceSs.getSheetByName(cfg.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = sourceSs.insertSheet(cfg.LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 12).setValues([[
      '일시',
      '모드',
      '소스행',
      '메모셀',
      '라인수',
      '이관대상',
      '중복제외',
      '고객번호없음제외',
      '빈라인제외',
      '파싱실패',
      '유형별',
      '비고'
    ]]);
  }

  const nextRow = sheet.getLastRow() + 1;

  sheet.getRange(nextRow, 1, 1, 12).setValues([[
    new Date(),
    mode,
    stats.sourceRows,
    stats.sourceCellsWithMemo,
    stats.sourceLines,
    stats.readyRows,
    stats.duplicateSkipped,
    stats.noCustomerSkipped,
    stats.emptyLineSkipped,
    stats.parseFailedSkipped,
    JSON.stringify(stats.byType),
    ''
  ]]);
}


function loadExistingClientRequestIds_(targetSheet, targetHeaderInfo) {
  const map = {};
  const col = targetHeaderInfo.headerMap[normalizeHeaderKey_('클라이언트요청ID')];

  if (!col) return map;

  const lastRow = targetSheet.getLastRow();
  const headerRow = CONTACT_HISTORY_MIGRATION_CONFIG.TARGET_HEADER_ROW;

  if (lastRow <= headerRow) return map;

  const values = targetSheet
    .getRange(headerRow + 1, col, lastRow - headerRow, 1)
    .getDisplayValues();

  for (let i = 0; i < values.length; i++) {
    const id = cleanText_(values[i][0]);

    if (id) {
      map[id] = true;
    }
  }

  return map;
}


/*************************************************
 * ID 생성
 *************************************************/

function createMigrationHistoryId_(dateValue, customerNo, sourceRowNo, originalLine) {
  const date = dateValue instanceof Date ? dateValue : new Date();

  const datePart = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'yyyyMMdd'
  );

  const timePart = Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'HHmmss'
  );

  const hash = shortHash_(
    [
      'HISTORY',
      customerNo,
      sourceRowNo,
      originalLine
    ].join('|'),
    8
  );

  return `CH-${datePart}-${timePart}-${hash}`;
}


function createMigrationClientRequestId_(customerNo, sourceRowNo, originalLine) {
  const hash = shortHash_(
    [
      'MASTER_MEMO_MIGRATION',
      customerNo,
      sourceRowNo,
      originalLine
    ].join('|'),
    12
  );

  return `MIG-MASTER-MEMO-${customerNo || 'NO-CUSTOMER'}-${sourceRowNo}-${hash}`;
}


function shortHash_(text, length) {
  const digest = Utilities.computeDigest(
    Utilities.DigestAlgorithm.MD5,
    String(text || ''),
    Utilities.Charset.UTF_8
  );

  let hex = '';

  for (let i = 0; i < digest.length; i++) {
    const byte = digest[i] < 0 ? digest[i] + 256 : digest[i];
    hex += ('0' + byte.toString(16)).slice(-2);
  }

  return hex.slice(0, length || 8);
}


/*************************************************
 * 텍스트 정리
 *************************************************/

function cleanText_(value) {
  return String(value || '')
    .replace(/\u00A0/g, ' ')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .trim();
}


function cleanBodyText_(value) {
  return String(value || '')
    .replace(/^[\s:：\-–—.,，、/|]+/g, '')
    .replace(/\s+/g, ' ')
    .trim();
}


function isMeaninglessMigrationLine_(text) {
  let value = cleanText_(text);

  if (!value) return true;

  value = value
    .replace(/^["“”']+|["“”']+$/g, '')
    .replace(/^`+|`+$/g, '')
    .trim();

  if (!value) return true;

  if (/^[()\[\]{}"'`]+$/.test(value)) return true;
  if (/^[=\-_*#\s]+$/.test(value)) return true;
  if (/^\d{1,2}:\d{2}$/.test(value)) return true;

  // 날짜만 있는 경우
  if (/^(?:20\d{2}|\d{2})[.\/-]\d{1,2}[.\/-]\d{1,2}\.?$/.test(value)) return true;
  if (/^\d{1,2}[.\/]\d{1,2}[.)]?$/.test(value)) return true;
  if (/^\d{1,2}월\s*\d{1,2}일?[.)]?$/.test(value)) return true;

  return false;
}


function mergeStringList_(a, b) {
  return uniqueStringList_([].concat(a || []).concat(b || []));
}


function uniqueStringList_(list) {
  const seen = {};
  const result = [];

  for (let i = 0; i < list.length; i++) {
    const item = cleanText_(list[i]);

    if (!item) continue;

    if (!seen[item]) {
      seen[item] = true;
      result.push(item);
    }
  }

  return result;
}


function notifyContactHistoryMigration_(message) {
  try {
    SpreadsheetApp.getUi().alert(message);
  } catch (err) {
    Logger.log(message);
  }
}