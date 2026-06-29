/**
 * 수주확정/계약완료 고객사 폴더 → S1 KJ 공유 복사 도구
 *
 * 목적
 * - 영업관리대장 > 수주확정/계약완료 시트의 고객번호를 기준으로
 * - S1 고객사 파일 관리 공유드라이브/폴더 안의 고객사 폴더를 찾고
 * - S1 KJ 공유 공유드라이브/폴더로 "복사"합니다. 이동/바로가기 아님.
 * - 재실행 시 고객번호 앞자리 기준으로 대상 고객사 폴더를 찾고,
 *   이미 복사된 sourceId는 재복사하지 않으며, 새로 추가된 파일/폴더만 이어서 복사합니다.
 *
 * 사용 전 확인
 * 1) 가능하면 KJ_COPY_CONFIG.SOURCE_ROOT_ID / DEST_ROOT_ID에 공유드라이브 ID 또는 루트 폴더 ID를 직접 넣으세요.
 *    비워두면 이름으로 자동 탐색합니다.
 * 2) 이름 자동 탐색이 실패하면 Apps Script > 서비스 > Drive API(고급 Google 서비스)를 켜거나 ID를 직접 넣으세요.
 * 3) 기존 onOpen()이 이미 있으면, 기존 onOpen 안에 addKjShareCopyMenu_(); 한 줄만 추가하세요.
 */

const KJ_COPY_CONFIG = {
  CONTRACT_SHEET_NAME: '수주확정/계약완료',

  // 가능하면 ID를 직접 넣는 것을 권장합니다. 비우면 NAME으로 탐색합니다.
  SOURCE_ROOT_ID: '', // 예: '0Axxxxxxxxxxxxxxxx' 또는 폴더 ID
  SOURCE_ROOT_NAME: 'S1 고객사 파일 관리',

  DEST_ROOT_ID: '', // 예: '0Ayyyyyyyyyyyyyyyy' 또는 폴더 ID
  DEST_ROOT_NAME: 'S1 KJ 공유',

  LOG_SHEET_NAME: 'KJ공유복사로그',
  SUMMARY_SHEET_NAME: 'KJ공유복사요약',

  // Apps Script 실행 제한 회피용. 5분 근처에서 안전하게 멈추고, 다시 실행하면 이어서 복사됩니다.
  MAX_RUNTIME_MS: 1000 * 60 * 4.7,
  LOG_FLUSH_SIZE: 80,

  // 대상 폴더 안에 같은 이름의 파일이 이미 있는데 로그가 없을 때 중복 복사 방지용입니다.
  // true: 같은 이름 파일이 있으면 복사하지 않고 로그에 등록만 함
  // false: 로그가 없으면 무조건 복사
  SKIP_EXISTING_DEST_FILE_BY_NAME: true,
  SKIP_EXISTING_DEST_FOLDER_BY_NAME: true,
};

/**
 * 기존 onOpen이 이미 있으면 이 함수 전체를 중복으로 만들지 말고,
 * 기존 onOpen() 안에 addKjShareCopyMenu_(); 만 추가하세요.
 */
function onOpen() {
  addKjShareCopyMenu_();
}

function addKjShareCopyMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('KJ공유 이관')
    .addItem('수주 고객사 폴더 복사 실행', 'copyWonCustomerFoldersToKjShare')
    .addItem('진행상황 확인', 'showKjShareCopyStatus')
    .addSeparator()
    .addItem('로그/요약 시트 준비', 'prepareKjShareCopySheets')
    .addToUi();
}

function prepareKjShareCopySheets() {
  ensureKjCopySheets_();
  SpreadsheetApp.getUi().alert('KJ 공유 복사 로그/요약 시트 준비 완료');
}

function showKjShareCopyStatus() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = ss.getSheetByName(KJ_COPY_CONFIG.SUMMARY_SHEET_NAME);
  const log = ss.getSheetByName(KJ_COPY_CONFIG.LOG_SHEET_NAME);
  const summaryLast = summary ? summary.getLastRow() : 0;
  const logLast = log ? log.getLastRow() : 0;
  SpreadsheetApp.getUi().alert(
    'KJ 공유 복사 진행상황\n\n' +
    '- 요약 고객 수: ' + Math.max(0, summaryLast - 1) + '건\n' +
    '- 로그 행 수: ' + Math.max(0, logLast - 1) + '건\n\n' +
    '상세 내용은 [' + KJ_COPY_CONFIG.SUMMARY_SHEET_NAME + '] / [' + KJ_COPY_CONFIG.LOG_SHEET_NAME + '] 시트를 확인하세요.'
  );
}

function copyWonCustomerFoldersToKjShare() {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) {
    throw new Error('다른 복사 작업이 실행 중입니다. 잠시 후 다시 실행하세요.');
  }

  const startedAt = Date.now();
  const runId = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss') + '_' + Math.random().toString(36).slice(2, 7);
  const userEmail = getActiveUserEmail_();
  const logRows = [];

  try {
    ensureKjCopySheets_();

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const customerNos = getWonCustomerNos_();
    if (!customerNos.length) {
      SpreadsheetApp.getUi().alert('수주확정/계약완료 시트에서 고객번호를 찾지 못했습니다.');
      return;
    }

    const sourceRoot = resolveRootFolder_(KJ_COPY_CONFIG.SOURCE_ROOT_ID, KJ_COPY_CONFIG.SOURCE_ROOT_NAME, '원본');
    const destRoot = resolveRootFolder_(KJ_COPY_CONFIG.DEST_ROOT_ID, KJ_COPY_CONFIG.DEST_ROOT_NAME, '대상');

    const copiedIndex = loadCopiedSourceIndex_();
    const sourceCustomerFolderMap = buildCustomerFolderMapByPrefix_(sourceRoot);
    const destCustomerFolderMap = buildCustomerFolderMapByPrefix_(destRoot);

    const summaryMap = loadSummaryMap_();
    const counters = {
      customerTotal: customerNos.length,
      customerDone: 0,
      customerMissingSource: 0,
      customerError: 0,
      folderCreated: 0,
      folderRegisteredExisting: 0,
      fileCopied: 0,
      fileRegisteredExisting: 0,
      skippedLogged: 0,
      errors: 0,
      stoppedByTime: false,
    };

    for (const customerNo of customerNos) {
      if (isTimeUp_(startedAt)) {
        counters.stoppedByTime = true;
        break;
      }

      const srcFolder = sourceCustomerFolderMap[customerNo];
      if (!srcFolder) {
        counters.customerMissingSource++;
        addLog_(logRows, runId, customerNo, 'CUSTOMER_FOLDER', 'SOURCE_FOLDER_NOT_FOUND', '', '', '', '', '', '', '원본 루트에서 앞자리 고객번호 폴더를 찾지 못함', userEmail);
        updateSummary_(summaryMap, customerNo, '', '', '', '', 'SOURCE_FOLDER_NOT_FOUND', '원본 고객사 폴더 없음');
        flushLogsIfNeeded_(logRows);
        continue;
      }

      let dstFolder = destCustomerFolderMap[customerNo];
      try {
        if (!dstFolder) {
          dstFolder = destRoot.createFolder(srcFolder.getName());
          destCustomerFolderMap[customerNo] = dstFolder;
          counters.folderCreated++;
          addLog_(logRows, runId, customerNo, 'CUSTOMER_FOLDER', 'CREATED_DEST_CUSTOMER_FOLDER', srcFolder.getId(), srcFolder.getName(), srcFolder.getName(), dstFolder.getId(), dstFolder.getName(), dstFolder.getName(), '', userEmail);
        }

        copyFolderIncremental_(srcFolder, dstFolder, {
          customerNo,
          runId,
          userEmail,
          copiedIndex,
          logRows,
          counters,
          startedAt,
          sourceBasePath: srcFolder.getName(),
          destBasePath: dstFolder.getName(),
        });

        counters.customerDone++;
        updateSummary_(summaryMap, customerNo, srcFolder.getId(), srcFolder.getName(), dstFolder.getId(), dstFolder.getName(), counters.stoppedByTime ? 'PARTIAL' : 'DONE', counters.stoppedByTime ? '시간 제한으로 일부 처리 후 중단. 다시 실행하면 이어서 처리됨.' : '복사 확인 완료');
      } catch (err) {
        counters.customerError++;
        counters.errors++;
        addLog_(logRows, runId, customerNo, 'CUSTOMER_FOLDER', 'ERROR', srcFolder.getId(), srcFolder.getName(), srcFolder.getName(), dstFolder ? dstFolder.getId() : '', dstFolder ? dstFolder.getName() : '', dstFolder ? dstFolder.getName() : '', String(err && err.message ? err.message : err), userEmail);
        updateSummary_(summaryMap, customerNo, srcFolder.getId(), srcFolder.getName(), dstFolder ? dstFolder.getId() : '', dstFolder ? dstFolder.getName() : '', 'ERROR', String(err && err.message ? err.message : err));
      }

      flushLogsIfNeeded_(logRows);
    }

    flushLogRows_(logRows);
    writeSummaryMap_(summaryMap);

    const msg = [
      'KJ 공유 고객사 폴더 복사 실행 완료',
      '',
      '대상 고객 수: ' + counters.customerTotal,
      '처리 고객 수: ' + counters.customerDone,
      '원본 폴더 없음: ' + counters.customerMissingSource,
      '고객 처리 오류: ' + counters.customerError,
      '신규 폴더 생성: ' + counters.folderCreated,
      '신규 파일 복사: ' + counters.fileCopied,
      '기존 대상 파일 등록/스킵: ' + counters.fileRegisteredExisting,
      '이미 복사된 파일/폴더 스킵: ' + counters.skippedLogged,
      counters.stoppedByTime ? '\n시간 제한 근처에서 안전 중단했습니다. 다시 실행하면 이어서 복사합니다.' : '',
      '',
      '상세 로그: ' + KJ_COPY_CONFIG.LOG_SHEET_NAME,
    ].join('\n');

    SpreadsheetApp.getUi().alert(msg);
  } finally {
    lock.releaseLock();
  }
}

function getWonCustomerNos_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(KJ_COPY_CONFIG.CONTRACT_SHEET_NAME);
  if (!sheet) throw new Error('시트를 찾을 수 없습니다: ' + KJ_COPY_CONFIG.CONTRACT_SHEET_NAME);

  const values = sheet.getDataRange().getValues();
  if (!values.length) return [];

  const headerInfo = findHeaderInfo_(values, ['고객번호', '고객 번호', 'customerNo', 'customer no']);
  if (!headerInfo) throw new Error(KJ_COPY_CONFIG.CONTRACT_SHEET_NAME + ' 시트에서 고객번호 헤더를 찾지 못했습니다.');

  const set = {};
  for (let r = headerInfo.row + 1; r < values.length; r++) {
    const no = normalizeCustomerNo_(values[r][headerInfo.col]);
    if (no) set[no] = true;
  }
  return Object.keys(set).sort(function(a, b) { return Number(a) - Number(b); });
}

function findHeaderInfo_(values, headerCandidates) {
  const normalizedCandidates = headerCandidates.map(normalizeHeader_);
  const maxRows = Math.min(values.length, 10);
  for (let r = 0; r < maxRows; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const h = normalizeHeader_(values[r][c]);
      if (normalizedCandidates.indexOf(h) >= 0) return { row: r, col: c };
    }
  }
  return null;
}

function normalizeHeader_(v) {
  return String(v == null ? '' : v).replace(/\s+/g, '').replace(/[()\[\]{}]/g, '').toLowerCase();
}

function normalizeCustomerNo_(v) {
  if (v == null || v === '') return '';
  if (typeof v === 'number') return String(Math.trunc(v));
  const s = String(v).trim();
  const m = s.match(/\d+/);
  return m ? String(Number(m[0])) : '';
}

function extractLeadingCustomerNoFromFolderName_(name) {
  const m = String(name || '').trim().match(/^(\d+)/);
  return m ? String(Number(m[1])) : '';
}

function buildCustomerFolderMapByPrefix_(rootFolder) {
  const map = {};
  const folders = rootFolder.getFolders();
  while (folders.hasNext()) {
    const folder = folders.next();
    const no = extractLeadingCustomerNoFromFolderName_(folder.getName());
    if (!no) continue;
    if (!map[no]) map[no] = folder;
  }
  return map;
}

function resolveRootFolder_(id, name, label) {
  if (id) {
    try {
      return DriveApp.getFolderById(id);
    } catch (err) {
      throw new Error(label + ' 루트 ID로 폴더/공유드라이브를 열 수 없습니다: ' + id + '\n' + err.message);
    }
  }

  // 일반 폴더명 탐색
  const folders = DriveApp.getFoldersByName(name);
  if (folders.hasNext()) return folders.next();

  // 공유드라이브명 탐색: 고급 Drive 서비스가 켜져 있으면 사용
  const driveId = findSharedDriveIdByName_(name);
  if (driveId) {
    try {
      return DriveApp.getFolderById(driveId);
    } catch (err) {
      throw new Error(label + ' 공유드라이브는 찾았지만 DriveApp으로 루트를 열 수 없습니다. CONFIG에 루트 ID를 직접 넣어주세요. 공유드라이브명: ' + name + ', ID: ' + driveId + '\n' + err.message);
    }
  }

  throw new Error(label + ' 루트를 찾지 못했습니다: ' + name + '\nCONFIG의 SOURCE_ROOT_ID / DEST_ROOT_ID에 공유드라이브 또는 폴더 ID를 직접 넣어주세요.');
}

function findSharedDriveIdByName_(name) {
  try {
    if (typeof Drive !== 'undefined' && Drive.Drives && Drive.Drives.list) {
      const res = Drive.Drives.list({
        q: "name = '" + escapeDriveQuery_(name) + "'",
        pageSize: 10,
        fields: 'drives(id,name)'
      });
      const drives = res && res.drives ? res.drives : [];
      if (drives.length) return drives[0].id;
    }
  } catch (err) {
    // v3 Drive.Drives 실패 시 v2 Teamdrives fallback 시도
  }

  try {
    if (typeof Drive !== 'undefined' && Drive.Teamdrives && Drive.Teamdrives.list) {
      const res2 = Drive.Teamdrives.list({
        q: "name = '" + escapeDriveQuery_(name) + "'",
        maxResults: 10,
        fields: 'items(id,name)'
      });
      const items = res2 && res2.items ? res2.items : [];
      if (items.length) return items[0].id;
    }
  } catch (err2) {
    // 고급 Drive 서비스 미사용/권한 문제면 null 반환
  }
  return '';
}

function escapeDriveQuery_(s) {
  return String(s || '').replace(/\\/g, '\\\\').replace(/'/g, "\\'");
}

function copyFolderIncremental_(srcFolder, dstFolder, ctx) {
  if (isTimeUp_(ctx.startedAt)) {
    ctx.counters.stoppedByTime = true;
    return;
  }

  copyFilesIncremental_(srcFolder, dstFolder, ctx);
  if (ctx.counters.stoppedByTime) return;

  const srcSubFolders = srcFolder.getFolders();
  while (srcSubFolders.hasNext()) {
    if (isTimeUp_(ctx.startedAt)) {
      ctx.counters.stoppedByTime = true;
      return;
    }

    const srcSub = srcSubFolders.next();
    const srcSubId = srcSub.getId();
    const srcSubPath = ctx.sourceBasePath + '/' + srcSub.getName();
    let dstSub = null;

    const copied = ctx.copiedIndex[srcSubId];
    if (copied && copied.destId) {
      dstSub = safeGetFolderById_(copied.destId);
      if (dstSub) {
        ctx.counters.skippedLogged++;
      }
    }

    if (!dstSub && KJ_COPY_CONFIG.SKIP_EXISTING_DEST_FOLDER_BY_NAME) {
      dstSub = getChildFolderByExactName_(dstFolder, srcSub.getName());
      if (dstSub) {
        ctx.copiedIndex[srcSubId] = { destId: dstSub.getId(), type: 'FOLDER' };
        ctx.counters.folderRegisteredExisting++;
        addLog_(ctx.logRows, ctx.runId, ctx.customerNo, 'FOLDER', 'EXISTS_BY_NAME_REGISTERED', srcSubId, srcSubPath, srcSub.getName(), dstSub.getId(), ctx.destBasePath + '/' + dstSub.getName(), dstSub.getName(), '대상에 같은 이름 폴더가 있어 기존 폴더로 연결', ctx.userEmail);
        flushLogsIfNeeded_(ctx.logRows);
      }
    }

    if (!dstSub) {
      dstSub = dstFolder.createFolder(srcSub.getName());
      ctx.copiedIndex[srcSubId] = { destId: dstSub.getId(), type: 'FOLDER' };
      ctx.counters.folderCreated++;
      addLog_(ctx.logRows, ctx.runId, ctx.customerNo, 'FOLDER', 'CREATED', srcSubId, srcSubPath, srcSub.getName(), dstSub.getId(), ctx.destBasePath + '/' + dstSub.getName(), dstSub.getName(), '', ctx.userEmail);
      flushLogsIfNeeded_(ctx.logRows);
    }

    const childCtx = Object.assign({}, ctx, {
      sourceBasePath: srcSubPath,
      destBasePath: ctx.destBasePath + '/' + dstSub.getName(),
    });
    copyFolderIncremental_(srcSub, dstSub, childCtx);
    if (ctx.counters.stoppedByTime) return;
  }
}

function copyFilesIncremental_(srcFolder, dstFolder, ctx) {
  const files = srcFolder.getFiles();
  while (files.hasNext()) {
    if (isTimeUp_(ctx.startedAt)) {
      ctx.counters.stoppedByTime = true;
      return;
    }

    const srcFile = files.next();
    const srcId = srcFile.getId();
    const srcPath = ctx.sourceBasePath + '/' + srcFile.getName();

    const copied = ctx.copiedIndex[srcId];
    if (copied && copied.destId && safeGetFileById_(copied.destId)) {
      ctx.counters.skippedLogged++;
      continue;
    }

    if (KJ_COPY_CONFIG.SKIP_EXISTING_DEST_FILE_BY_NAME) {
      const existing = getChildFileByExactName_(dstFolder, srcFile.getName());
      if (existing) {
        ctx.copiedIndex[srcId] = { destId: existing.getId(), type: 'FILE' };
        ctx.counters.fileRegisteredExisting++;
        addLog_(ctx.logRows, ctx.runId, ctx.customerNo, 'FILE', 'EXISTS_BY_NAME_REGISTERED', srcId, srcPath, srcFile.getName(), existing.getId(), ctx.destBasePath + '/' + existing.getName(), existing.getName(), '대상에 같은 이름 파일이 있어 중복 복사하지 않고 등록', ctx.userEmail);
        flushLogsIfNeeded_(ctx.logRows);
        continue;
      }
    }

    try {
      const copiedFile = srcFile.makeCopy(srcFile.getName(), dstFolder);
      ctx.copiedIndex[srcId] = { destId: copiedFile.getId(), type: 'FILE' };
      ctx.counters.fileCopied++;
      addLog_(ctx.logRows, ctx.runId, ctx.customerNo, 'FILE', 'COPIED', srcId, srcPath, srcFile.getName(), copiedFile.getId(), ctx.destBasePath + '/' + copiedFile.getName(), copiedFile.getName(), '', ctx.userEmail);
      flushLogsIfNeeded_(ctx.logRows);
    } catch (err) {
      ctx.counters.errors++;
      addLog_(ctx.logRows, ctx.runId, ctx.customerNo, 'FILE', 'ERROR', srcId, srcPath, srcFile.getName(), '', '', '', String(err && err.message ? err.message : err), ctx.userEmail);
      flushLogsIfNeeded_(ctx.logRows);
    }
  }
}

function getChildFolderByExactName_(parent, name) {
  const iter = parent.getFoldersByName(name);
  return iter.hasNext() ? iter.next() : null;
}

function getChildFileByExactName_(parent, name) {
  const iter = parent.getFilesByName(name);
  return iter.hasNext() ? iter.next() : null;
}

function safeGetFileById_(id) {
  if (!id) return null;
  try { return DriveApp.getFileById(id); } catch (err) { return null; }
}

function safeGetFolderById_(id) {
  if (!id) return null;
  try { return DriveApp.getFolderById(id); } catch (err) { return null; }
}

function ensureKjCopySheets_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let log = ss.getSheetByName(KJ_COPY_CONFIG.LOG_SHEET_NAME);
  if (!log) log = ss.insertSheet(KJ_COPY_CONFIG.LOG_SHEET_NAME);
  if (log.getLastRow() === 0) {
    log.getRange(1, 1, 1, 13).setValues([[ // 13 cols
      '일시', '실행ID', '고객번호', '구분', '상태', '소스ID', '소스경로', '소스명', '대상ID', '대상경로', '대상명', '비고', '실행자'
    ]]);
    log.setFrozenRows(1);
  }

  let summary = ss.getSheetByName(KJ_COPY_CONFIG.SUMMARY_SHEET_NAME);
  if (!summary) summary = ss.insertSheet(KJ_COPY_CONFIG.SUMMARY_SHEET_NAME);
  if (summary.getLastRow() === 0) {
    summary.getRange(1, 1, 1, 9).setValues([[ // 9 cols
      '고객번호', '원본폴더ID', '원본폴더명', '대상폴더ID', '대상폴더명', '최종상태', '최종메시지', '최종확인일시', '최종실행자'
    ]]);
    summary.setFrozenRows(1);
  }
}

function loadCopiedSourceIndex_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const log = ss.getSheetByName(KJ_COPY_CONFIG.LOG_SHEET_NAME);
  const index = {};
  if (!log || log.getLastRow() < 2) return index;

  const values = log.getRange(2, 1, log.getLastRow() - 1, Math.min(13, log.getLastColumn())).getValues();
  values.forEach(function(row) {
    const type = String(row[3] || '');
    const status = String(row[4] || '');
    const sourceId = String(row[5] || '');
    const destId = String(row[8] || '');
    if (!sourceId || !destId) return;
    if (['COPIED', 'CREATED', 'CREATED_DEST_CUSTOMER_FOLDER', 'EXISTS_BY_NAME_REGISTERED'].indexOf(status) < 0) return;
    index[sourceId] = { type: type, destId: destId, status: status };
  });
  return index;
}

function loadSummaryMap_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KJ_COPY_CONFIG.SUMMARY_SHEET_NAME);
  const map = {};
  if (!sh || sh.getLastRow() < 2) return map;
  const values = sh.getRange(2, 1, sh.getLastRow() - 1, 9).getValues();
  values.forEach(function(row) {
    const customerNo = normalizeCustomerNo_(row[0]);
    if (!customerNo) return;
    map[customerNo] = row;
  });
  return map;
}

function updateSummary_(summaryMap, customerNo, sourceId, sourceName, destId, destName, status, message) {
  summaryMap[customerNo] = [
    customerNo,
    sourceId || (summaryMap[customerNo] && summaryMap[customerNo][1]) || '',
    sourceName || (summaryMap[customerNo] && summaryMap[customerNo][2]) || '',
    destId || (summaryMap[customerNo] && summaryMap[customerNo][3]) || '',
    destName || (summaryMap[customerNo] && summaryMap[customerNo][4]) || '',
    status || '',
    message || '',
    new Date(),
    getActiveUserEmail_(),
  ];
}

function writeSummaryMap_(summaryMap) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KJ_COPY_CONFIG.SUMMARY_SHEET_NAME);
  const keys = Object.keys(summaryMap).sort(function(a, b) { return Number(a) - Number(b); });
  if (sh.getLastRow() > 1) sh.getRange(2, 1, sh.getLastRow() - 1, 9).clearContent();
  if (!keys.length) return;
  const rows = keys.map(function(k) { return summaryMap[k]; });
  sh.getRange(2, 1, rows.length, 9).setValues(rows);
  sh.getRange(2, 8, rows.length, 1).setNumberFormat('yyyy.MM.dd. HH:mm:ss');
  sh.autoResizeColumns(1, 9);
}

function addLog_(rows, runId, customerNo, itemType, status, sourceId, sourcePath, sourceName, destId, destPath, destName, memo, userEmail) {
  rows.push([
    new Date(),
    runId,
    customerNo,
    itemType,
    status,
    sourceId || '',
    sourcePath || '',
    sourceName || '',
    destId || '',
    destPath || '',
    destName || '',
    memo || '',
    userEmail || '',
  ]);
}

function flushLogsIfNeeded_(rows) {
  if (rows.length >= KJ_COPY_CONFIG.LOG_FLUSH_SIZE) flushLogRows_(rows);
}

function flushLogRows_(rows) {
  if (!rows.length) return;
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sh = ss.getSheetByName(KJ_COPY_CONFIG.LOG_SHEET_NAME);
  sh.getRange(sh.getLastRow() + 1, 1, rows.length, 13).setValues(rows);
  sh.getRange(Math.max(2, sh.getLastRow() - rows.length + 1), 1, rows.length, 1).setNumberFormat('yyyy.MM.dd. HH:mm:ss');
  rows.length = 0;
}

function isTimeUp_(startedAt) {
  return Date.now() - startedAt > KJ_COPY_CONFIG.MAX_RUNTIME_MS;
}

function getActiveUserEmail_() {
  try {
    return Session.getActiveUser().getEmail() || Session.getEffectiveUser().getEmail() || '';
  } catch (err) {
    return '';
  }
}
