/**
 * P524: 메일/발송파일 큐 및 트리거 상태 점검판
 * - 큐를 실행하거나 재처리하지 않습니다.
 * - 상태/건수/최근 오류/트리거 설치 여부만 읽어서 보여줍니다.
 */
function showMailQueueSummaryP524() {
  const ui = SpreadsheetApp.getUi();
  const cfg = (typeof getSentFileArchiveConfig_ === 'function') ? getSentFileArchiveConfig_() : {};
  const archiveQueueName = String(cfg.QUEUE_SHEET_NAME || '발송파일저장큐_DB');
  const archiveHandler = String(cfg.ASYNC_TRIGGER_HANDLER || 'processDeferredSentFileArchiveQueueV94');

  const lines = [];
  lines.push('[메일/발송파일 큐 요약 - P524]');
  lines.push('');
  lines.push(mqhSummarizeQueueSheetP524_('메일발송실패큐_DB', {
    note: '자동 재발송 금지. 사람이 확인 후 처리.',
    preferredSpreadsheetLabels: ['생성기', '현재']
  }));
  lines.push('');
  lines.push(mqhSummarizeQueueSheetP524_(archiveQueueName, {
    note: '발송파일 보관 실패/대기 건. 트리거로 자동 재시도.',
    preferredSpreadsheetLabels: ['생성기', '현재']
  }));
  lines.push('');
  lines.push(mqhSummarizeQueueSheetP524_('저장큐_DB', {
    note: '포탈 저장 fallback 큐. 같은 파일에 없으면 시트 없음으로 표시.',
    preferredSpreadsheetLabels: ['현재', '마스터', '생성기']
  }));
  lines.push('');
  lines.push(mqhSummarizeQueueSheetP524_('변경큐_DB', {
    note: '포탈 검색/변경 반영 큐. 같은 파일에 없으면 시트 없음으로 표시.',
    preferredSpreadsheetLabels: ['현재', '마스터', '생성기']
  }));
  lines.push('');
  lines.push(mqhSummarizeTriggerP524_(archiveHandler));

  ui.alert('메일/발송파일 큐 요약', lines.join('\n'), ui.ButtonSet.OK);
}

function mqhSummarizeQueueSheetP524_(sheetName, options) {
  options = options || {};
  const found = mqhFindQueueSheetsP524_(sheetName, options.preferredSpreadsheetLabels || []);
  if (!found.length) {
    return [
      sheetName,
      '- 상태: 시트 없음',
      options.note ? '- 참고: ' + options.note : ''
    ].filter(Boolean).join('\n');
  }

  return found.slice(0, 3).map(function(item) {
    return mqhSummarizeOneQueueSheetP524_(item.sheet, sheetName, item.label, options);
  }).join('\n');
}

function mqhFindQueueSheetsP524_(sheetName, preferredLabels) {
  const candidates = mqhGetCandidateSpreadsheetsP524_();
  const preferred = Array.isArray(preferredLabels) ? preferredLabels : [];
  candidates.sort(function(a, b) {
    const ai = preferred.indexOf(a.label);
    const bi = preferred.indexOf(b.label);
    return (ai < 0 ? 999 : ai) - (bi < 0 ? 999 : bi);
  });

  const out = [];
  candidates.forEach(function(item) {
    try {
      const sheet = item.ss.getSheetByName(sheetName);
      if (sheet) out.push({ label: item.label, spreadsheetId: item.ss.getId(), sheet: sheet });
    } catch (err) {}
  });
  return out;
}

function mqhGetCandidateSpreadsheetsP524_() {
  const out = [];
  const seen = {};

  function add(label, ss) {
    if (!ss || !ss.getId) return;
    const id = ss.getId();
    if (!id || seen[id]) return;
    seen[id] = true;
    out.push({ label: label, ss: ss });
  }

  try { add('현재', SpreadsheetApp.getActive()); } catch (err) {}

  const config = (typeof CONFIG !== 'undefined' && CONFIG) ? CONFIG : {};
  const generatorId = String(config.GENERATOR_SPREADSHEET_ID || '').trim();
  const masterId = String(config.MASTER_SPREADSHEET_ID || '').trim();

  if (generatorId) {
    try { add('생성기', SpreadsheetApp.openById(generatorId)); } catch (err) {}
  }
  if (masterId) {
    try { add('마스터', SpreadsheetApp.openById(masterId)); } catch (err) {}
  }

  return out;
}

function mqhSummarizeOneQueueSheetP524_(sheet, sheetName, label, options) {
  options = options || {};
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(sheet.getLastColumn(), 1);
  const title = sheetName + ' [' + label + ']';

  if (lastRow < 2) {
    return [
      title,
      '- 전체: 0건',
      '- 상태: 데이터 없음',
      options.note ? '- 참고: ' + options.note : ''
    ].filter(Boolean).join('\n');
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) {
    return String(v || '').trim();
  });
  const maxRows = Math.min(lastRow - 1, 1000);
  const startRow = Math.max(2, lastRow - maxRows + 1);
  const values = sheet.getRange(startRow, 1, maxRows, lastCol).getValues();

  const statusIdx = mqhFindHeaderIndexP524_(headers, ['상태', 'status', 'Status']);
  const statusCounts = {};
  if (statusIdx >= 0) {
    values.forEach(function(row) {
      const st = String(row[statusIdx] || '(빈값)').trim() || '(빈값)';
      statusCounts[st] = (statusCounts[st] || 0) + 1;
    });
  }

  const statusText = statusIdx >= 0
    ? Object.keys(statusCounts).sort().map(function(k) { return k + ' ' + statusCounts[k] + '건'; }).join(' / ')
    : '상태 컬럼 없음';

  const createdText = mqhFindLastNonEmptyValueP524_(values, headers, ['등록일시', 'createdAt', 'created_at', '등록시간']);
  const updatedText = mqhFindLastNonEmptyValueP524_(values, headers, ['수정일시', '완료일시', '적용일시', 'updatedAt', 'updated_at']);
  const errorText = mqhFindLastNonEmptyValueP524_(values, headers, ['마지막오류', '오류', 'error', 'errorMessage', '결과JSON', 'resultJson']);

  const lines = [];
  lines.push(title);
  lines.push('- 전체: ' + (lastRow - 1) + '건' + (maxRows < lastRow - 1 ? ' / 최근 ' + maxRows + '건 기준' : ''));
  lines.push('- 상태: ' + statusText);
  if (createdText) lines.push('- 마지막 등록: ' + createdText);
  if (updatedText) lines.push('- 마지막 수정/처리: ' + updatedText);
  if (errorText) lines.push('- 최근 오류/결과: ' + mqhClipP524_(errorText, 180));
  if (options.note) lines.push('- 참고: ' + options.note);
  return lines.join('\n');
}

function mqhFindHeaderIndexP524_(headers, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const idx = headers.indexOf(String(candidates[i] || '').trim());
    if (idx >= 0) return idx;
  }
  return -1;
}

function mqhFindLastNonEmptyValueP524_(values, headers, candidates) {
  const idx = mqhFindHeaderIndexP524_(headers, candidates);
  if (idx < 0) return '';
  for (let r = values.length - 1; r >= 0; r--) {
    const value = values[r][idx];
    const text = mqhFormatValueP524_(value);
    if (text) return text;
  }
  return '';
}

function mqhFormatValueP524_(value) {
  if (value == null || value === '') return '';
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
  }
  return String(value || '').trim();
}

function mqhClipP524_(text, maxLen) {
  const s = String(text || '').replace(/\s+/g, ' ').trim();
  const n = Number(maxLen || 180) || 180;
  return s.length > n ? s.slice(0, n) + '…' : s;
}

function mqhSummarizeTriggerP524_(handler) {
  const targetHandler = String(handler || 'processDeferredSentFileArchiveQueueV94').trim();
  let triggers = [];
  try {
    triggers = ScriptApp.getProjectTriggers().filter(function(trigger) {
      return trigger && trigger.getHandlerFunction && trigger.getHandlerFunction() === targetHandler;
    });
  } catch (err) {
    return '트리거\n- 상태: 조회 실패\n- 오류: ' + mqhClipP524_(err && err.message || err, 180);
  }

  const lines = [];
  lines.push('트리거');
  lines.push('- 대상 함수: ' + targetHandler);
  if (triggers.length === 0) {
    lines.push('- 상태: 미설치');
  } else if (triggers.length === 1) {
    lines.push('- 상태: 정상 설치됨 1개');
  } else {
    lines.push('- 상태: 중복 주의 ' + triggers.length + '개');
  }

  if (triggers.length) {
    const detail = triggers.slice(0, 5).map(function(trigger, idx) {
      let source = '';
      let eventType = '';
      try { source = trigger.getTriggerSource ? String(trigger.getTriggerSource()) : ''; } catch (err1) {}
      try { eventType = trigger.getEventType ? String(trigger.getEventType()) : ''; } catch (err2) {}
      return '#' + (idx + 1) + (eventType ? ' ' + eventType : '') + (source ? ' / ' + source : '');
    }).join(', ');
    if (detail) lines.push('- 상세: ' + detail);
  }
  return lines.join('\n');
}
