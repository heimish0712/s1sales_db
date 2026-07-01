/*******************************************************
 * 마스터시트 메모 내 TM 구메모 분리 이관 1회성 코드
 *
 * 목적:
 * - 활성 시트에서 실행
 * - '메모' 열 안에 과거 TM 이관으로 섞여 들어간 TM 메모만 추출
 * - 추출한 TM 메모를 'TM 진행 현황 (7/1 전)' 열로 이동
 * - 기존 '메모' 열에서는 해당 TM 메모 줄 삭제
 *
 * 중요:
 * - 복사가 아니라 이동입니다.
 * - TM 메모로 확실히 판단되는 줄만 이동합니다.
 * - 기본 판정 기준은 줄 끝/내용에 '(유현희TM)' 같은 TM suffix가 있는 경우입니다.
 *******************************************************/

const OLD_TM_MEMO_MOVE_CONFIG = {
  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  MEMO_HEADER: '메모',
  TARGET_HEADER: 'TM 진행 현황 (7/1 전)',

  // 기준 연도. 6/10 같은 날짜는 2026년으로 해석합니다.
  DEFAULT_YEAR: 2026,

  // 확실한 TM 표기만 이동합니다. 이름 뒤 TM 표기가 없는 일반 메모는 건드리지 않습니다.
  TM_NAME_CANDIDATES: [
    '유현희', '박원경', '라유화', '김서하', '방수원', '박새봄', '문형진', '이옥희', '김경아', '최보람'
  ],

  // false 권장: TM 표시 없는 다음 줄까지 임의로 끌고 가면 일반 메모를 건드릴 위험이 큽니다.
  // 단, '(유현희TM)'처럼 명시적 TM suffix가 있는 날짜 없는 줄은 직전 TM 항목 뒤에 붙입니다.
  INCLUDE_UNMARKED_CONTINUATION_AFTER_TM_LINE: false,

  LOG_SHEET_NAME: 'TM_7월전메모분리로그'
};

/**
 * 먼저 이 함수로 결과만 확인하세요. 실제 시트 값은 변경하지 않고 로그 시트에 PREVIEW로 남깁니다.
 */
function previewMoveOldTmMemosFromActiveSheet() {
  return moveOldTmMemosFromActiveSheet_(true);
}

/**
 * 실제 이동 실행 함수.
 * - 메모에서 TM 줄 삭제
 * - TM 진행 현황 (7/1 전) 열에 누적
 */
function moveOldTmMemosFromActiveSheet() {
  return moveOldTmMemosFromActiveSheet_(false);
}

function moveOldTmMemosFromActiveSheet_(dryRun) {
  const cfg = OLD_TM_MEMO_MOVE_CONFIG;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  if (!sheet) throw new Error('활성 시트를 찾지 못했습니다.');

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < cfg.DATA_START_ROW) {
    ss.toast('처리할 데이터 행이 없습니다.', 'TM 구메모 분리', 5);
    return { ok: true, processed: 0, moved: 0 };
  }

  const headers = sheet.getRange(cfg.HEADER_ROW, 1, 1, lastCol).getDisplayValues()[0];
  const headerMap = buildOldTmHeaderMap_(headers);

  const memoCol = findOldTmCol_(headerMap, [cfg.MEMO_HEADER]);
  if (!memoCol) throw new Error('활성 시트에서 [메모] 헤더를 찾지 못했습니다. 헤더 행은 ' + cfg.HEADER_ROW + '행 기준입니다.');

  let targetCol = findOldTmCol_(headerMap, [cfg.TARGET_HEADER]);
  if (!targetCol) {
    if (dryRun) {
      // preview에서는 실제 헤더를 만들지 않고 가상으로만 안내합니다.
      targetCol = lastCol + 1;
    } else {
      targetCol = lastCol + 1;
      sheet.getRange(cfg.HEADER_ROW, targetCol).setValue(cfg.TARGET_HEADER);
    }
  }

  const numRows = lastRow - cfg.DATA_START_ROW + 1;
  const memoRange = sheet.getRange(cfg.DATA_START_ROW, memoCol, numRows, 1);
  const targetRange = sheet.getRange(cfg.DATA_START_ROW, targetCol, numRows, 1);
  const memoValues = memoRange.getValues();
  const targetValues = targetRange.getValues();

  const logRows = [];
  let touchedRows = 0;
  let movedLineCount = 0;
  let duplicateLineCount = 0;

  for (let i = 0; i < numRows; i++) {
    const rowNo = cfg.DATA_START_ROW + i;
    const memo = String(memoValues[i][0] || '');
    if (!memo.trim()) continue;

    const extracted = extractOldTmEntriesFromMemo_(memo, cfg);
    if (!extracted.entries.length) continue;

    const appendResult = appendUniqueOldTmLines_(String(targetValues[i][0] || ''), extracted.renderedLines);

    memoValues[i][0] = extracted.cleanedMemo;
    targetValues[i][0] = appendResult.newText;

    touchedRows++;
    movedLineCount += appendResult.added.length;
    duplicateLineCount += appendResult.duplicated.length;

    logRows.push([
      new Date(),
      dryRun ? 'PREVIEW' : 'MOVE',
      sheet.getName(),
      rowNo,
      'OK',
      '추출 ' + extracted.entries.length + '건 / 대상 신규 ' + appendResult.added.length + '건 / 대상중복 ' + appendResult.duplicated.length + '건',
      extracted.originalLines.join('\n'),
      extracted.renderedLines.join('\n'),
      appendResult.duplicated.join('\n'),
      shortenOldTmText_(memo, 3000),
      shortenOldTmText_(extracted.cleanedMemo, 3000)
    ]);
  }

  if (!dryRun && touchedRows > 0) {
    memoRange.setValues(memoValues);
    targetRange.setValues(targetValues);
    SpreadsheetApp.flush();
  }

  if (logRows.length) writeOldTmMoveLog_(ss, logRows);

  const msg = (dryRun ? '[미리보기] ' : '') +
    'TM 구메모 분리 완료: 대상행 ' + touchedRows + '건 / 신규이동 ' + movedLineCount + '줄 / 대상중복 ' + duplicateLineCount + '줄';
  ss.toast(msg, 'TM 구메모 분리', 8);

  return {
    ok: true,
    dryRun: dryRun,
    sheetName: sheet.getName(),
    touchedRows: touchedRows,
    movedLineCount: movedLineCount,
    duplicateLineCount: duplicateLineCount,
    targetHeader: cfg.TARGET_HEADER
  };
}

function extractOldTmEntriesFromMemo_(memo, cfg) {
  const rawLines = String(memo || '').replace(/\r/g, '').split('\n');
  const cleanedLines = [];
  const entries = [];
  const originalLines = [];
  let previousWasExplicitTm = false;

  rawLines.forEach(function(line) {
    const raw = String(line || '');
    const trimmed = raw.trim();

    if (!trimmed) {
      cleanedLines.push(raw);
      previousWasExplicitTm = false;
      return;
    }

    const parsed = parseOldTmMemoLine_(trimmed, cfg);
    if (parsed.isTm) {
      originalLines.push(trimmed);

      // 날짜 없는 TM 명시 줄은 직전 TM 항목 뒤에 이어 붙입니다.
      if (!parsed.dateText && entries.length > 0) {
        const last = entries[entries.length - 1];
        last.content = joinOldTmSentence_(last.content, parsed.content);
        if (!last.tmName && parsed.tmName) last.tmName = parsed.tmName;
      } else {
        entries.push({
          dateText: parsed.dateText,
          content: parsed.content,
          tmName: parsed.tmName,
          originalLine: trimmed
        });
      }

      previousWasExplicitTm = true;
      return;
    }

    // 기본값 false. 일반 메모를 건드리지 않기 위해 명시 TM 표기가 없는 줄은 보존합니다.
    if (cfg.INCLUDE_UNMARKED_CONTINUATION_AFTER_TM_LINE === true && previousWasExplicitTm && entries.length > 0 && !looksLikeOldTmNonTmBoundary_(trimmed)) {
      originalLines.push(trimmed);
      entries[entries.length - 1].content = joinOldTmSentence_(entries[entries.length - 1].content, trimmed);
      return;
    }

    cleanedLines.push(raw);
    previousWasExplicitTm = false;
  });

  const renderedLines = entries
    .map(renderOldTmEntryLine_)
    .map(function(v) { return String(v || '').trim(); })
    .filter(Boolean);

  return {
    entries: entries,
    originalLines: originalLines,
    renderedLines: renderedLines,
    cleanedMemo: cleanupOldTmMemoAfterMove_(cleanedLines.join('\n'))
  };
}

function parseOldTmMemoLine_(line, cfg) {
  let text = String(line || '').trim();
  const tmName = extractOldTmName_(text, cfg);
  if (!tmName) {
    return { isTm: false };
  }

  text = removeOldTmSuffix_(text, cfg).trim();
  text = text.replace(/라고\s*함\.?\s*$/g, '').trim();

  const dateParsed = parseLeadingOldTmDate_(text, cfg.DEFAULT_YEAR || 2026);
  let dateText = '';
  let content = text;

  if (dateParsed && dateParsed.dateText) {
    dateText = dateParsed.dateText;
    content = dateParsed.rest;
  }

  content = String(content || '')
    .replace(/^[-–—:：)\.\s]+/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (!content) return { isTm: false };

  return {
    isTm: true,
    dateText: dateText,
    content: content,
    tmName: tmName
  };
}

function extractOldTmName_(line, cfg) {
  const text = String(line || '').trim();

  // 가장 확실한 형태: (유현희TM), (유현희 TM), (TM미상)
  const parenTm = text.match(/\(([^()]{1,30}?)(?:\s*TM)\)\s*$/i);
  if (parenTm) {
    const name = String(parenTm[1] || '').trim();
    if (!name || /^미상$/i.test(name)) return 'TM미상';
    return name.replace(/\s+/g, '');
  }

  const anyParenTm = text.match(/\(([^()]{1,30}?)(?:\s*TM)\)/i);
  if (anyParenTm) {
    const name = String(anyParenTm[1] || '').trim();
    if (!name || /^미상$/i.test(name)) return 'TM미상';
    return name.replace(/\s+/g, '');
  }

  // 괄호가 없더라도 "유현희TM"처럼 명확하면 인식
  const names = cfg.TM_NAME_CANDIDATES || [];
  for (let i = 0; i < names.length; i++) {
    const n = String(names[i] || '').trim();
    if (!n) continue;
    const re = new RegExp(escapeOldTmRegex_(n) + '\\s*TM', 'i');
    if (re.test(text)) return n;
  }

  return '';
}

function removeOldTmSuffix_(line, cfg) {
  let text = String(line || '').trim();
  text = text.replace(/\([^()]{1,30}?\s*TM\)\s*$/ig, '').trim();
  text = text.replace(/\([^()]{1,30}?\s*TM\)/ig, '').trim();
  (cfg.TM_NAME_CANDIDATES || []).forEach(function(name) {
    if (!name) return;
    const re = new RegExp(escapeOldTmRegex_(name) + '\\s*TM\\s*$', 'i');
    text = text.replace(re, '').trim();
  });
  return text;
}

function parseLeadingOldTmDate_(text, defaultYear) {
  let s = String(text || '').trim();

  // 앞의 "1차)", "2차)" 같은 회차 prefix는 날짜 파싱에서만 제거합니다.
  s = s.replace(/^\s*\d+\s*차\s*[\)\.]?\s*/g, '').trim();

  let m = s.match(/^(20\d{2}|\d{2})\s*[.\/\-년]\s*(\d{1,2})\s*[.\/\-월]\s*(\d{1,2})\s*(?:일)?\.?\s*[-–—:：)]?\s*(.*)$/);
  if (m) {
    let year = Number(m[1]);
    if (year < 100) year += 2000;
    return buildOldTmDateParseResult_(year, Number(m[2]), Number(m[3]), m[4]);
  }

  m = s.match(/^(\d{1,2})\s*(?:\/|\.|\-|월)\s*(\d{1,2})\s*(?:일)?\.?\s*[-–—:：)]?\s*(.*)$/);
  if (m) {
    return buildOldTmDateParseResult_(Number(defaultYear) || 2026, Number(m[1]), Number(m[2]), m[3]);
  }

  return null;
}

function buildOldTmDateParseResult_(year, month, day, rest) {
  if (!year || !month || !day || month < 1 || month > 12 || day < 1 || day > 31) return null;
  const yy = String(year).slice(-2);
  const mm = String(month).padStart(2, '0');
  const dd = String(day).padStart(2, '0');
  return {
    dateText: yy + '.' + mm + '.' + dd + '.',
    rest: String(rest || '').trim()
  };
}

function renderOldTmEntryLine_(entry) {
  const content = String(entry && entry.content || '').replace(/\s+/g, ' ').trim();
  if (!content) return '';
  const tm = String(entry && entry.tmName || '').trim();
  const suffix = tm ? ' (' + tm + ')' : '';
  const date = String(entry && entry.dateText || '').trim();
  return (date ? (date + ' ') : '') + content + suffix;
}

function appendUniqueOldTmLines_(oldText, newLines) {
  const oldLines = splitOldTmNonEmptyLines_(oldText);
  const out = oldLines.slice();
  const existing = {};
  oldLines.forEach(function(line) {
    const key = normalizeOldTmDuplicateKey_(line);
    if (key) existing[key] = true;
  });

  const added = [];
  const duplicated = [];

  (newLines || []).forEach(function(line) {
    const clean = String(line || '').trim();
    if (!clean) return;
    const key = normalizeOldTmDuplicateKey_(clean);
    if (key && existing[key]) {
      duplicated.push(clean);
      return;
    }
    out.push(clean);
    added.push(clean);
    if (key) existing[key] = true;
  });

  return {
    newText: out.join('\n').trim(),
    added: added,
    duplicated: duplicated
  };
}

function normalizeOldTmDuplicateKey_(line) {
  return String(line || '')
    .replace(/\r/g, '')
    .replace(/^\s*\d{2}\.\d{2}\.\d{2}\.\s*/g, '')
    .replace(/\([^()]{1,30}\)\s*$/g, '')
    .replace(/라고\s*함\.?\s*$/g, '')
    .replace(/[\s\u00a0]+/g, ' ')
    .trim()
    .toLowerCase();
}

function splitOldTmNonEmptyLines_(text) {
  return String(text || '')
    .replace(/\r/g, '')
    .split('\n')
    .map(function(v) { return String(v || '').trim(); })
    .filter(Boolean);
}

function cleanupOldTmMemoAfterMove_(text) {
  return String(text || '')
    .replace(/\r/g, '')
    .split('\n')
    .map(function(v) { return String(v || '').trim(); })
    .reduce(function(acc, line) {
      // 빈 줄 과다 방지
      if (!line) {
        if (acc.length && acc[acc.length - 1] !== '') acc.push('');
        return acc;
      }
      acc.push(line);
      return acc;
    }, [])
    .join('\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}

function looksLikeOldTmNonTmBoundary_(line) {
  const s = String(line || '').trim();
  if (!s) return true;
  if (/^\[/.test(s)) return true;
  if (/^\d+\s*차\s*[\)\.]/.test(s)) return true;
  if (/^\d{1,2}\s*[\/\.\-월]\s*\d{1,2}/.test(s)) return true;
  if (/영업지원요청|컨택이력|다음액션|파일|발송|견적|계약완료/.test(s)) return true;
  return false;
}

function joinOldTmSentence_(a, b) {
  const left = String(a || '').trim();
  const right = String(b || '').trim();
  if (!left) return right;
  if (!right) return left;
  return (left + ' ' + right).replace(/\s+/g, ' ').trim();
}

function buildOldTmHeaderMap_(headers) {
  const map = {};
  (headers || []).forEach(function(h, idx) {
    const key = normalizeOldTmHeader_(h);
    if (key && !map[key]) map[key] = idx + 1;
  });
  return map;
}

function findOldTmCol_(headerMap, aliases) {
  for (let i = 0; i < (aliases || []).length; i++) {
    const key = normalizeOldTmHeader_(aliases[i]);
    if (headerMap[key]) return headerMap[key];
  }
  return 0;
}

function normalizeOldTmHeader_(v) {
  return String(v || '')
    .replace(/[\r\n]/g, '')
    .replace(/\s+/g, '')
    .replace(/[(){}\[\]_\-·ㆍ.,\/\\]/g, '')
    .toLowerCase()
    .trim();
}

function writeOldTmMoveLog_(ss, rows) {
  const cfg = OLD_TM_MEMO_MOVE_CONFIG;
  let sheet = ss.getSheetByName(cfg.LOG_SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(cfg.LOG_SHEET_NAME);
    sheet.getRange(1, 1, 1, 11).setValues([[
      '처리시각', '모드', '시트명', '행번호', '상태', '처리요약', '원본TM줄', '이동결과', '대상중복', '기존메모일부', '정리후메모일부'
    ]]);
    sheet.setFrozenRows(1);
  }
  sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length).setValues(rows);
}

function shortenOldTmText_(text, maxLen) {
  const s = String(text || '');
  const n = Number(maxLen) || 3000;
  return s.length > n ? s.slice(0, n) + '... [truncated]' : s;
}

function escapeOldTmRegex_(text) {
  return String(text || '').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}
