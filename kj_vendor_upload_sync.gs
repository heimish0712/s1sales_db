/**
 * KJ 수행사 업로드 동기화
 *
 * 새 파일명:
 *   kj_vendor_upload_sync.gs
 *
 * 역할:
 * - 우리 공유 드라이브의 일괄 업로드 폴더 2개를 확인
 * - 파일명 본문이 "_KJ"로 끝나는 파일만 KJ 수행사 폴더로 복사
 * - 파일명은 그대로 유지
 * - 원본 파일은 수정/삭제/이동하지 않음
 * - sourceFileId + destFolderId 로그 기준으로 중복 복사 방지
 * - 로그가 없더라도 대상 폴더에 같은 파일명이 있으면 복사하지 않음
 *
 * 기존 코드 충돌 방지:
 * - 전역 함수/상수 전부 KJUS_ 접두사 사용
 * - onOpen/onEdit를 새로 선언하지 않음
 */

const KJUS_CFG = {
  LOG_SHEET_NAME: 'KJ수행사_업로드동기화_LOG',

  LOG_HEADER: [
    '처리일시',
    '구분',
    '상태',
    '원본파일ID',
    '원본파일명',
    '원본폴더ID',
    '대상폴더ID',
    '복사파일ID',
    '복사파일명',
    '메시지'
  ],

  RULES: [
    {
      label: '사업자등록증',
      sourceFolderId: '1kG8JtnwOiAVWv_gY_PZtU0Wm7EV7UxN1',
      destFolderId: '1P4tZNLzqY46wWltPhtIxb7Mqhg2yREvK'
    },
    {
      label: '계약서류',
      sourceFolderId: '1YYkWjJhMRb62bsQbREMsGYM2GeKuyZCe',
      destFolderId: '1enmgDejBzlzly5_k9DcBHOf3YK4dEkXg'
    }
  ],

  GOOGLE_FOLDER_MIME: 'application/vnd.google-apps.folder',

  LOCK_WAIT_MS: 3000,
  MAX_RUNTIME_MS: 4.5 * 60 * 1000,
  MAX_COPIES_PER_RUN: 200,
  MAX_HANDLED_KJ_FILES_PER_RUN: 300,
  LOG_FLUSH_BATCH_SIZE: 20
};


/**
 * 30분 트리거가 실행할 메인 함수.
 * 수동 실행해도 됩니다.
 */
function KJUS_runVendorUploadSync() {
  const startedMs = Date.now();
  const summary = KJUS_createEmptySummary_();
  const lease = AUTOMATION_acquireModuleLease_(
    'KJ_VENDOR_UPLOAD',
    {
      taskName: 'KJUS_runVendorUploadSync',
      waitMs: KJUS_CFG.LOCK_WAIT_MS,
      ttlMs: 6 * 60 * 1000
    }
  );

  if (!lease.acquired) {
    summary.status = 'LOCKED_SKIP';
    summary.message = '다른 KJ 수행사 업로드 동기화가 진행 중이라 이번 실행은 건너뜀';
    summary.leaseReason = lease.reason || 'LEASE_BUSY';
    Logger.log(JSON.stringify(summary));
    return summary;
  }

  let logSheet = null;
  const logRows = [];

  try {
    logSheet = KJUS_getOrCreateLogSheet_();

    const processedKeySet = KJUS_loadProcessedKeySet_(logSheet);

    for (let i = 0; i < KJUS_CFG.RULES.length; i++) {
      if (KJUS_shouldStopRun_(startedMs, summary)) {
        break;
      }

      const rule = KJUS_CFG.RULES[i];

      try {
        KJUS_syncOneRule_({
          rule,
          processedKeySet,
          logSheet,
          logRows,
          summary,
          startedMs
        });
      } catch (ruleErr) {
        summary.errors++;
        summary.ruleErrors++;

        KJUS_queueLogRow_(logSheet, logRows, [
          new Date(),
          rule.label,
          'RULE_ERROR',
          '',
          '',
          rule.sourceFolderId,
          rule.destFolderId,
          '',
          '',
          KJUS_toErrorMessage_(ruleErr)
        ]);
      }
    }

    KJUS_flushLogRows_(logSheet, logRows);

    summary.status = summary.errors > 0 ? 'DONE_WITH_ERRORS' : 'DONE';
    summary.elapsedMs = Date.now() - startedMs;

    Logger.log(JSON.stringify(summary));
    return summary;

  } catch (err) {
    summary.status = 'FATAL_ERROR';
    summary.errors++;
    summary.elapsedMs = Date.now() - startedMs;
    summary.message = KJUS_toErrorMessage_(err);

    try {
      if (logSheet) {
        KJUS_queueLogRow_(logSheet, logRows, [
          new Date(),
          '전체실행',
          'FATAL_ERROR',
          '',
          '',
          '',
          '',
          '',
          '',
          summary.message
        ]);
        KJUS_flushLogRows_(logSheet, logRows);
      }
    } catch (logErr) {
      Logger.log('KJUS fatal log write failed: ' + KJUS_toErrorMessage_(logErr));
    }

    Logger.log(JSON.stringify(summary));
    throw err;

  } finally {
    try {
      AUTOMATION_releaseModuleLease_(lease);
    } catch (releaseErr) {
      Logger.log('KJUS lease release failed: ' + KJUS_toErrorMessage_(releaseErr));
    }
  }
}


/**
 * 메뉴에서 수동 실행할 때 쓰는 함수.
 */
function KJUS_runVendorUploadSyncFromMenu() {
  const summary = KJUS_runVendorUploadSync();
  KJUS_safeUiAlert_(KJUS_formatSummary_(summary));
}




/**
 * 로그 시트만 미리 생성/정비.
 */
function KJUS_prepareVendorUploadSyncLogSheet() {
  const sheet = KJUS_getOrCreateLogSheet_();
  const message = '로그 시트 준비 완료: ' + sheet.getName();

  Logger.log(message);
  KJUS_safeUiAlert_(message);
}


/**
 * 기존 onOpen()에 선택적으로 연결할 메뉴 함수.
 *
 * 기존 onOpen() 안에 아래 한 줄만 추가하면 됩니다.
 *   KJUS_addVendorUploadSyncMenu_();
 */
function KJUS_addVendorUploadSyncMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('KJ 수행사 업로드 동기화')
    .addItem('지금 1회 실행', 'KJUS_runVendorUploadSyncFromMenu')
    .addSeparator()
    .addItem('로그 시트 준비', 'KJUS_prepareVendorUploadSyncLogSheet')
    .addToUi();
}


/**
 * 규칙 1개 처리.
 */
function KJUS_syncOneRule_(params) {
  const rule = params.rule;
  const processedKeySet = params.processedKeySet;
  const logSheet = params.logSheet;
  const logRows = params.logRows;
  const summary = params.summary;
  const startedMs = params.startedMs;

  KJUS_forEachDirectChildFile_(
    rule.sourceFolderId,
    function(file) {
      if (KJUS_shouldStopRun_(startedMs, summary)) {
        return false;
      }

      summary.scanned++;

      const sourceFileId = file.id;
      const sourceFileName = KJUS_getFileTitle_(file);

      if (!KJUS_isKjFileName_(sourceFileName)) {
        summary.skippedNotKj++;
        return true;
      }

      summary.kjCandidates++;

      const processedKey = KJUS_buildProcessedKey_(sourceFileId, rule.destFolderId);

      if (processedKeySet[processedKey]) {
        summary.skippedAlreadyLogged++;
        return true;
      }

      try {
        const existing = KJUS_findFileByExactTitleInFolder_(
          rule.destFolderId,
          sourceFileName
        );

        if (existing) {
          processedKeySet[processedKey] = true;
          summary.existsByName++;
          summary.handledKjFiles++;

          KJUS_queueLogRow_(logSheet, logRows, [
            new Date(),
            rule.label,
            'EXISTS_BY_NAME',
            sourceFileId,
            sourceFileName,
            rule.sourceFolderId,
            rule.destFolderId,
            existing.id || '',
            KJUS_getFileTitle_(existing),
            '대상 폴더에 같은 파일명이 이미 있어 복사하지 않음'
          ]);

          return true;
        }

        const copied = KJUS_copyFileToFolder_(
          sourceFileId,
          rule.destFolderId,
          sourceFileName
        );

        processedKeySet[processedKey] = true;
        summary.copied++;
        summary.handledKjFiles++;

        KJUS_queueLogRow_(logSheet, logRows, [
          new Date(),
          rule.label,
          'COPIED',
          sourceFileId,
          sourceFileName,
          rule.sourceFolderId,
          rule.destFolderId,
          copied.id || '',
          KJUS_getFileTitle_(copied) || sourceFileName,
          '복사 완료'
        ]);

        return true;

      } catch (fileErr) {
        summary.errors++;
        summary.handledKjFiles++;

        KJUS_queueLogRow_(logSheet, logRows, [
          new Date(),
          rule.label,
          'ERROR',
          sourceFileId,
          sourceFileName,
          rule.sourceFolderId,
          rule.destFolderId,
          '',
          '',
          KJUS_toErrorMessage_(fileErr)
        ]);

        return true;
      }
    },
    startedMs
  );
}


/**
 * 특정 폴더의 직계 파일만 순회.
 * 하위 폴더 재귀 탐색은 하지 않습니다.
 */
function KJUS_forEachDirectChildFile_(folderId, callback, startedMs) {
  let pageToken = null;

  const q = [
    "'" + KJUS_escapeDriveQueryText_(folderId) + "' in parents",
    'trashed = false',
    "mimeType != '" + KJUS_CFG.GOOGLE_FOLDER_MIME + "'"
  ].join(' and ');

  do {
    if (Date.now() - startedMs >= KJUS_CFG.MAX_RUNTIME_MS) {
      return false;
    }

    const params = {
      q: q,
      maxResults: 1000,
      fields: 'items(id,title,mimeType,createdDate,modifiedDate),nextPageToken',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    };

    if (pageToken) {
      params.pageToken = pageToken;
    }

    const result = Drive.Files.list(params);
    const items = result.items || [];

    for (let i = 0; i < items.length; i++) {
      const shouldContinue = callback(items[i]);

      if (shouldContinue === false) {
        return false;
      }
    }

    pageToken = result.nextPageToken;

  } while (pageToken);

  return true;
}


/**
 * 대상 폴더에 같은 파일명이 이미 있는지 확인.
 */
function KJUS_findFileByExactTitleInFolder_(folderId, fileTitle) {
  const q = [
    "'" + KJUS_escapeDriveQueryText_(folderId) + "' in parents",
    'trashed = false',
    "mimeType != '" + KJUS_CFG.GOOGLE_FOLDER_MIME + "'",
    "title = '" + KJUS_escapeDriveQueryText_(fileTitle) + "'"
  ].join(' and ');

  const result = Drive.Files.list({
    q: q,
    maxResults: 1,
    fields: 'items(id,title,mimeType)',
    supportsAllDrives: true,
    includeItemsFromAllDrives: true
  });

  const items = result.items || [];
  return items.length > 0 ? items[0] : null;
}


/**
 * 파일 1개 복사.
 * 파일명은 원본 그대로 유지합니다.
 */
function KJUS_copyFileToFolder_(sourceFileId, destFolderId, fileTitle) {
  return Drive.Files.copy(
    {
      title: fileTitle,
      parents: [{ id: destFolderId }]
    },
    sourceFileId,
    {
      supportsAllDrives: true,
      fields: 'id,title,mimeType'
    }
  );
}


/**
 * 파일명 판정.
 * 확장자 제외 후 본문 끝이 _KJ 인 경우만 true.
 */
function KJUS_isKjFileName_(fileName) {
  const name = String(fileName || '').trim();

  if (!name) {
    return false;
  }

  const stem = name
    .replace(/\.[^./\\]+$/, '')
    .trim();

  return /_KJ$/i.test(stem);
}


/**
 * 로그 시트 생성/정비.
 */
function KJUS_getOrCreateLogSheet_() {
  const ss = AUTOMATION_getRuntimeMasterSpreadsheet_();

  let sheet = ss.getSheetByName(KJUS_CFG.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(KJUS_CFG.LOG_SHEET_NAME);
  }

  const header = KJUS_CFG.LOG_HEADER;
  const headerRange = sheet.getRange(1, 1, 1, header.length);

  headerRange.setValues([header]);
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);

  return sheet;
}


/**
 * 이미 처리된 sourceFileId + destFolderId 조합 로드.
 */
function KJUS_loadProcessedKeySet_(logSheet) {
  const processed = {};
  const lastRow = logSheet.getLastRow();

  if (lastRow < 2) {
    return processed;
  }

  const values = logSheet
    .getRange(2, 1, lastRow - 1, KJUS_CFG.LOG_HEADER.length)
    .getValues();

  values.forEach(function(row) {
    const status = String(row[2] || '').trim();
    const sourceFileId = String(row[3] || '').trim();
    const destFolderId = String(row[6] || '').trim();

    if (!sourceFileId || !destFolderId) {
      return;
    }

    if (status === 'COPIED' || status === 'EXISTS_BY_NAME') {
      processed[KJUS_buildProcessedKey_(sourceFileId, destFolderId)] = true;
    }
  });

  return processed;
}


/**
 * 로그 행 버퍼 추가.
 */
function KJUS_queueLogRow_(logSheet, logRows, row) {
  logRows.push(KJUS_normalizeLogRow_(row));

  if (logRows.length >= KJUS_CFG.LOG_FLUSH_BATCH_SIZE) {
    KJUS_flushLogRows_(logSheet, logRows);
  }
}


/**
 * 로그 버퍼 실제 기록.
 */
function KJUS_flushLogRows_(logSheet, logRows) {
  if (!logRows || logRows.length === 0) {
    return;
  }

  const startRow = logSheet.getLastRow() + 1;
  const width = KJUS_CFG.LOG_HEADER.length;

  logSheet
    .getRange(startRow, 1, logRows.length, width)
    .setValues(logRows);

  logRows.length = 0;
}


/**
 * 로그 행 길이 보정.
 */
function KJUS_normalizeLogRow_(row) {
  const width = KJUS_CFG.LOG_HEADER.length;
  const normalized = [];

  for (let i = 0; i < width; i++) {
    normalized.push(i < row.length ? row[i] : '');
  }

  return normalized;
}


/**
 * 처리 키 생성.
 */
function KJUS_buildProcessedKey_(sourceFileId, destFolderId) {
  return String(sourceFileId || '') + '|' + String(destFolderId || '');
}


/**
 * Drive query 문자열 escape.
 */
function KJUS_escapeDriveQueryText_(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");
}


/**
 * Drive v2 title/name 호환.
 */
function KJUS_getFileTitle_(file) {
  if (!file) {
    return '';
  }

  return String(file.title || file.name || '').trim();
}


/**
 * 실행 중단 조건.
 */
function KJUS_shouldStopRun_(startedMs, summary) {
  if (Date.now() - startedMs >= KJUS_CFG.MAX_RUNTIME_MS) {
    summary.stoppedByTime = true;
    return true;
  }

  if (summary.copied >= KJUS_CFG.MAX_COPIES_PER_RUN) {
    summary.stoppedByCopyLimit = true;
    return true;
  }

  if (summary.handledKjFiles >= KJUS_CFG.MAX_HANDLED_KJ_FILES_PER_RUN) {
    summary.stoppedByHandleLimit = true;
    return true;
  }

  return false;
}


/**
 * 요약 객체.
 */
function KJUS_createEmptySummary_() {
  return {
    status: 'RUNNING',
    scanned: 0,
    kjCandidates: 0,
    copied: 0,
    existsByName: 0,
    skippedNotKj: 0,
    skippedAlreadyLogged: 0,
    handledKjFiles: 0,
    errors: 0,
    ruleErrors: 0,
    stoppedByTime: false,
    stoppedByCopyLimit: false,
    stoppedByHandleLimit: false,
    elapsedMs: 0,
    message: ''
  };
}


/**
 * 요약 문구.
 */
function KJUS_formatSummary_(summary) {
  return [
    'KJ 수행사 업로드 동기화 결과',
    '',
    '상태: ' + summary.status,
    '전체 스캔 파일 수: ' + summary.scanned,
    '_KJ 후보 파일 수: ' + summary.kjCandidates,
    '복사 완료: ' + summary.copied,
    '대상 동일 파일명으로 스킵: ' + summary.existsByName,
    '처리 로그 기준 스킵: ' + summary.skippedAlreadyLogged,
    '_KJ 아님 스킵: ' + summary.skippedNotKj,
    '오류: ' + summary.errors,
    '소요 ms: ' + summary.elapsedMs,
    '',
    summary.stoppedByTime ? '※ 실행시간 제한으로 일부만 처리됨' : '',
    summary.stoppedByCopyLimit ? '※ 1회 복사 개수 제한으로 일부만 처리됨' : '',
    summary.stoppedByHandleLimit ? '※ 1회 처리 개수 제한으로 일부만 처리됨' : '',
    summary.message ? '메시지: ' + summary.message : ''
  ].filter(function(line) {
    return line !== '';
  }).join('\n');
}


/**
 * 트리거 삭제 내부 함수.
 */

/**
 * UI alert 안전 호출.
 * 시간 트리거에서는 UI가 없으므로 Logger로만 남깁니다.
 */
function KJUS_safeUiAlert_(message) {
  try {
    SpreadsheetApp.getUi().alert(String(message || ''));
  } catch (err) {
    Logger.log(String(message || ''));
  }
}


/**
 * 에러 메시지 정리.
 */
function KJUS_toErrorMessage_(err) {
  if (!err) {
    return '';
  }

  const message = err.message ? err.message : String(err);
  return KJUS_truncate_(message, 1000);
}


/**
 * 문자열 길이 제한.
 */
function KJUS_truncate_(value, maxLength) {
  const text = String(value || '');

  if (text.length <= maxLength) {
    return text;
  }

  return text.slice(0, maxLength) + '...';
}
