/***** 고객사 공유드라이브 폴더 자동 생성 설정 *****
 * P004: 고객번호 기준 중복 방지/중복 폴더 병합/수동 실행 고속화.
 * - 고객번호가 같은 폴더는 한 개만 유지하고, 나머지 폴더의 내용은 기준 폴더로 병합 후 휴지통 이동.
 * - 고객사 폴더 전용 ScriptProperties soft lock을 실제 배타 락으로 사용.
 * - 수동 init/continue는 최대 5분, 최대 5,000행까지 처리하며 시트 쓰기는 일괄 반영.
 * - Drive API 일시 오류/쿼터 오류에 재시도 적용.
 *****/
const CUSTOMER_FOLDER_CFG = {
  // 마스터시트에 바인딩된 Apps Script면 빈 값 유지.
  // 독립형 Apps Script면 마스터 스프레드시트 ID 입력.
  MASTER_SPREADSHEET_ID: '',

  MASTER_SHEET_NAME: '마스터시트(신규)',

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  // 원본 고객사 파일 관리 공유드라이브명
  SHARED_DRIVE_NAME: 'S1 고객사 파일 관리',

  // 가능하면 비워둬도 됨. 이름으로 공유드라이브를 찾음.
  // 이름 조회가 안 되면 공유드라이브 URL의 folders/ 뒤 ID를 여기에 직접 입력.
  SHARED_DRIVE_ID: '',

  // 수행사 공란일 때 폴더명에 들어갈 값
  EMPTY_VENDOR_TEXT: '수행사미정',

  // 1회 수동 실행 최대 행 수. 실제 종료 기준은 아래 5분 시간 제한이 우선입니다.
  MAX_ROWS_PER_RUN: 5000,

  // Apps Script 6분 제한보다 여유를 두고 실제 작업은 최대 5분 수행합니다.
  MAX_MILLIS_PER_RUN: 5 * 60 * 1000,

  // 고객사 폴더 soft lock이 이미 점유 중일 때 기다릴 최대 시간.
  LOCK_WAIT_MILLIS: 8000,

  // 설치형 onEdit는 사람이 셀 편집할 때마다 들어오므로, 오래 기다리지 않고 조용히 포기.
  ONEDIT_LOCK_WAIT_MILLIS: 300,

  // LockService.getScriptLock()은 프로젝트 전체 락이라 다른 시간기반 동기화 함수까지 같이 막음.
  // 고객사 폴더 작업은 자체 soft lock으로만 중복 실행을 막고, 다른 업무 락은 무시함.
  SOFT_LOCK_TTL_MILLIS: 10 * 60 * 1000,

  // ScriptProperties soft lock 동시 획득 경합 시 최종 토큰 소유자만 실행하도록 확인하는 시간.
  SOFT_LOCK_ELECTION_MILLIS: 120,

  // 장시간 수동 실행 중 soft lock이 오래된 것으로 오인되지 않도록 주기적으로 갱신.
  SOFT_LOCK_HEARTBEAT_MILLIS: 30 * 1000,

  // Drive API 429/5xx/일시적 403 오류 재시도.
  DRIVE_API_MAX_RETRIES: 4,
  DRIVE_API_RETRY_BASE_MILLIS: 500,

  // 중복 폴더 안에 같은 이름의 서로 다른 파일이 있을 때 보존용 이름 뒤에 붙일 문구.
  DUPLICATE_MERGE_SUFFIX: '중복폴더병합',
  DUPLICATE_MERGE_MAX_DEPTH: 25,

  // 고객사 폴더 안에 기본 하위폴더까지 만들지 여부
  CREATE_STANDARD_SUBFOLDERS: false,

  STANDARD_SUBFOLDERS: [
    '00_최신발송본',
    '01_사업자등록증',
    '02_고객사수취서류',
    '03_계약진행서류',
    '99_기타'
  ],

  // 마스터시트에 자동으로 추가할 관리 컬럼
  OUTPUT_HEADERS: {
    folderId: '고객사폴더ID',
    folderUrl: '고객사폴더URL',
    folderName: '고객사폴더명',
    folderStatus: '고객사폴더처리상태',
    folderUpdatedAt: '고객사폴더업데이트일시'
  },

  RENAME_IF_CHANGED: true,

  LOG_SHEET_NAME: '고객사폴더_LOG',

  TZ: 'Asia/Seoul'
};


/***** 수주실패 고객사 폴더 이동 설정 *****/
const FAILED_CUSTOMER_FOLDER_CFG = {
  FAILED_PARENT_FOLDER_NAME: '수주실패',

  // 마스터시트에서 상태값 헤더 후보. 실제 시트에 있는 이름을 우선 찾음.
  STATUS_HEADER_CANDIDATES: [
    '계약진행현황',
    '진행상태',
    '상태',
    '계약상태',
    '현재 영업 진행 상황'
  ],

  // 아래 키워드가 상태값에 포함되면 수주실패로 판단.
  FAILED_STATUS_KEYWORDS: [
    '수주실패',
    '영업종료(거절)',
    '영업종료',
    '계약중도취소',
    '거절',
    '실패'
  ],

  MAX_ROWS_PER_RUN: 5000,
  MAX_MILLIS_PER_RUN: 5 * 60 * 1000,

  PROP_NEXT_ROW: 'S1_FAILED_FOLDER_MOVE_NEXT_ROW',

  OUTPUT_HEADERS: {
    failedMoveStatus: '수주실패폴더이동상태',
    failedMoveUpdatedAt: '수주실패폴더이동일시'
  }

};


/***** Lock/실행 충돌 방지 유틸 *****/

function acquireCustomerFolderLockOrReturn_(taskName, waitMs) {
  // 고객사 폴더 코드끼리만 충돌을 막는 ScriptProperties 기반 soft lock.
  // 다른 프로젝트 함수가 사용하는 LockService.getScriptLock()과는 독립적입니다.
  const cfg = CUSTOMER_FOLDER_CFG;
  const props = PropertiesService.getScriptProperties();
  const key = 'S1_CUSTOMER_FOLDER_SOFT_LOCK';
  const ttlMs = Number(cfg.SOFT_LOCK_TTL_MILLIS || (10 * 60 * 1000));
  const electionMs = Number(cfg.SOFT_LOCK_ELECTION_MILLIS || 120);
  const waitUntilMs = Date.now() + Math.max(0, Number(waitMs || 0));

  while (true) {
    const nowMs = Date.now();
    const raw = props.getProperty(key);

    if (raw) {
      try {
        const info = JSON.parse(raw);
        const heartbeatAtMs = Number(info.heartbeatAtMs || info.startedAtMs || 0);
        const ageMs = heartbeatAtMs ? nowMs - heartbeatAtMs : ttlMs + 1;

        if (heartbeatAtMs && ageMs >= 0 && ageMs < ttlMs) {
          if (nowMs < waitUntilMs) {
            Utilities.sleep(Math.min(250, Math.max(30, waitUntilMs - nowMs)));
            continue;
          }

          Logger.log(
            `[${taskName}] 다른 고객사 폴더 작업이 실행 중이므로 이번 실행을 건너뜁니다. ` +
            `점유 함수=${info.taskName || ''} / 시작=${info.startedAt || ''} / ` +
            `마지막 갱신=${info.heartbeatAt || info.startedAt || ''} / 경과초=${Math.round(ageMs / 1000)}`
          );
          return null;
        }

        Logger.log(
          `[${taskName}] 만료된 고객사 폴더 soft lock을 회수합니다. ` +
          `이전 함수=${info.taskName || ''} / 시작=${info.startedAt || ''}`
        );
      } catch (err) {
        Logger.log(`[${taskName}] 깨진 고객사 폴더 soft lock 기록을 회수합니다.`);
      }

      props.deleteProperty(key);
    }

    const token = Utilities.getUuid();
    const acquiredAtMs = Date.now();
    const acquiredAt = Utilities.formatDate(new Date(acquiredAtMs), cfg.TZ, 'yyyy-MM-dd HH:mm:ss');

    props.setProperty(key, JSON.stringify({
      token,
      taskName: taskName || '',
      startedAtMs: acquiredAtMs,
      startedAt: acquiredAt,
      heartbeatAtMs: acquiredAtMs,
      heartbeatAt: acquiredAt
    }));

    // 거의 동시에 두 실행이 진입하면 마지막으로 기록을 소유한 실행만 살아남습니다.
    Utilities.sleep(electionMs);

    let confirmed = null;
    try {
      const latestRaw = props.getProperty(key);
      confirmed = latestRaw ? JSON.parse(latestRaw) : null;
    } catch (err) {
      confirmed = null;
    }

    if (confirmed && confirmed.token === token) {
      return {
        key,
        token,
        taskName: taskName || '',
        startedAtMs: acquiredAtMs,
        lastHeartbeatMs: acquiredAtMs,

        refreshLock: function () {
          const now = Date.now();
          const heartbeatEvery = Number(cfg.SOFT_LOCK_HEARTBEAT_MILLIS || 30000);
          if (now - this.lastHeartbeatMs < heartbeatEvery) return true;

          try {
            const latestRaw = props.getProperty(key);
            if (!latestRaw) return false;

            const latest = JSON.parse(latestRaw);
            if (latest.token !== token) return false;

            latest.heartbeatAtMs = now;
            latest.heartbeatAt = Utilities.formatDate(new Date(now), cfg.TZ, 'yyyy-MM-dd HH:mm:ss');
            props.setProperty(key, JSON.stringify(latest));
            this.lastHeartbeatMs = now;
            return true;
          } catch (err) {
            Logger.log('[customer folder soft lock heartbeat 오류] ' + (err && err.message ? err.message : err));
            return false;
          }
        },

        releaseLock: function () {
          try {
            const latestRaw = props.getProperty(key);
            if (!latestRaw) return;

            const latest = JSON.parse(latestRaw);
            if (latest.token === token) {
              props.deleteProperty(key);
            }
          } catch (err) {
            Logger.log('[customer folder soft lock release 오류] ' + (err && err.message ? err.message : err));
          }
        }
      };
    }

    if (Date.now() >= waitUntilMs) {
      Logger.log(`[${taskName}] soft lock 동시 획득 경합에서 다른 실행이 우선권을 가져 이번 실행을 건너뜁니다.`);
      return null;
    }

    Utilities.sleep(100);
  }
}


function refreshCustomerFolderLock_(lock) {
  if (!lock || typeof lock.refreshLock !== 'function') return true;
  return lock.refreshLock();
}

function releaseCustomerFolderLock_(lock) {
  if (!lock) return;

  try {
    lock.releaseLock();
  } catch (err) {
    Logger.log('[releaseCustomerFolderLock_ 오류] ' + (err && err.message ? err.message : err));
  }
}


function makeCustomerFolderLockedResult_(taskName) {
  return {
    status: 'LOCKED',
    nextRow: 0,
    message:
      `[${taskName}] 다른 고객사 폴더 작업이 실행 중이라 이번 실행은 건너뛰었습니다. ` +
      `조금 뒤 같은 함수를 다시 실행하면 됩니다.`
  };
}


function reportCustomerFolderSoftLock() {
  const raw = PropertiesService.getScriptProperties().getProperty('S1_CUSTOMER_FOLDER_SOFT_LOCK');

  if (!raw) {
    Logger.log('현재 고객사 폴더 soft lock 점유 기록 없음');
    return null;
  }

  let info;
  try {
    info = JSON.parse(raw);
  } catch (err) {
    Logger.log('고객사 폴더 soft lock 기록이 깨져 있습니다: ' + raw);
    return { raw };
  }

  const heartbeatAtMs = Number(info.heartbeatAtMs || info.startedAtMs || 0);
  const ageSec = heartbeatAtMs ? Math.round((Date.now() - heartbeatAtMs) / 1000) : '';

  Logger.log(
    '현재 고객사 폴더 soft lock 점유 기록: ' +
    '함수=' + (info.taskName || '') +
    ' / 시작=' + (info.startedAt || '') +
    ' / 마지막갱신=' + (info.heartbeatAt || info.startedAt || '') +
    ' / 갱신후경과초=' + ageSec
  );

  return info;
}


function clearCustomerFolderSoftLockDebugOnly() {
  PropertiesService.getScriptProperties().deleteProperty('S1_CUSTOMER_FOLDER_SOFT_LOCK');
  Logger.log('고객사 폴더 soft lock 점유 기록 삭제 완료');
}


function logCustomerFolderDetectionResult_(detected) {
  if (!detected) return;

  if (detected.nextRow) {
    Logger.log(
      `다음 작업 시작 행 탐지 완료: ${detected.nextRow}행 / 고객번호 ${detected.customerNo || ''} / 회사명 ${detected.company || ''}` +
      ` / 사유: ${detected.reason || ''}` +
      ` / Drive 탐지 고객폴더 ${detected.driveCustomerFolderCount || 0}개` +
      ` / 시트 유효고객 ${detected.validCustomerCount || 0}건` +
      ` / 폴더ID 기재 ${detected.sheetFolderIdCount || 0}건`
    );
  } else {
    Logger.log(
      `생성 필요한 고객사 폴더 없음` +
      ` / Drive 탐지 고객폴더 ${detected.driveCustomerFolderCount || 0}개` +
      ` / 시트 유효고객 ${detected.validCustomerCount || 0}건` +
      ` / 폴더ID 기재 ${detected.sheetFolderIdCount || 0}건`
    );
  }
}


/***** 고객사 폴더 생성/보정 실행 함수 *****/

/**
 * 최초/수동 일괄 생성 함수.
 * 무조건 3행부터 돌지 않고, 공유드라이브 루트 + 수주실패 폴더를 먼저 스캔해서
 * 실제로 Drive에 고객번호 폴더가 없는 첫 행부터 시작함.
 */
function initCreateCustomerFoldersFromMaster() {
  const lock = acquireCustomerFolderLockOrReturn_(
    'initCreateCustomerFoldersFromMaster',
    CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS
  );

  if (!lock) {
    return makeCustomerFolderLockedResult_('initCreateCustomerFoldersFromMaster');
  }

  try {
    const runStartedAtMs = Date.now();
    const cfg = CUSTOMER_FOLDER_CFG;
    const props = PropertiesService.getScriptProperties();

    props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');

    const sheet = customerFolder_getMasterSheet_();

    let headerMap = customerFolder_getHeaderMap_(sheet);
    headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

    customerFolder_assertHeader_(headerMap, '고객번호');
    customerFolder_assertHeader_(headerMap, '회사명');
    customerFolder_assertHeader_(headerMap, '수행사');

    const lastRow = sheet.getLastRow();

    if (lastRow < cfg.DATA_START_ROW) {
      props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
      Logger.log('마스터시트에 처리할 데이터가 없습니다.');

      return {
        status: 'EMPTY',
        nextRow: 0,
        message: '마스터시트에 처리할 데이터가 없습니다.'
      };
    }

    // 기존 코드는 여기서 Drive 색인을 한 번 하고, continue에서 다시 한 번 해서 느렸음.
    // 이제 최초 실행에서는 Drive 색인을 1회만 만들고 이어서 처리까지 같은 색인을 사용함.
    const driveIndex = customerFolder_buildExistingCustomerFolderIndex_();
    const detected = customerFolder_detectNextCustomerFolderWorkRowFromIndex_(sheet, headerMap, driveIndex);

    if (detected.nextRow) {
      props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(detected.nextRow));
    } else {
      props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
    }

    logCustomerFolderDetectionResult_(detected);

    if (!detected.nextRow) {
      return detected;
    }

    return customerFolder_continueCreateCustomerFoldersFromMasterLocked_({
      sheet,
      headerMap,
      driveIndex,
      detected,
      lock,
      runStartedAtMs
    });

  } finally {
    releaseCustomerFolderLock_(lock);
  }
}



/**
 * 이어서 실행 함수.
 * 공유드라이브 폴더 목록을 한 번에 색인한 뒤 누락/미연결 건만 처리.
 */
function continueCreateCustomerFoldersFromMaster() {
  const lock = acquireCustomerFolderLockOrReturn_(
    'continueCreateCustomerFoldersFromMaster',
    CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS
  );

  if (!lock) {
    return makeCustomerFolderLockedResult_('continueCreateCustomerFoldersFromMaster');
  }

  try {
    return customerFolder_continueCreateCustomerFoldersFromMasterLocked_({ lock });
  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


function customerFolder_continueCreateCustomerFoldersFromMasterLocked_(options) {
  options = options || {};

  const cfg = CUSTOMER_FOLDER_CFG;
  const startedAt = Number(options.runStartedAtMs || Date.now());
  const deadlineMs = startedAt + Number(cfg.MAX_MILLIS_PER_RUN || (5 * 60 * 1000));
  const lock = options.lock || null;
  const sheet = options.sheet || customerFolder_getMasterSheet_();

  let headerMap = options.headerMap || customerFolder_getHeaderMap_(sheet);
  headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

  customerFolder_assertHeader_(headerMap, '고객번호');
  customerFolder_assertHeader_(headerMap, '회사명');
  customerFolder_assertHeader_(headerMap, '수행사');

  const props = PropertiesService.getScriptProperties();
  const lastRow = sheet.getLastRow();

  if (lastRow < cfg.DATA_START_ROW) {
    props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
    Logger.log('마스터시트에 처리할 데이터가 없습니다.');

    return {
      status: 'EMPTY',
      nextRow: 0,
      message: '마스터시트에 처리할 데이터가 없습니다.'
    };
  }

  let row = Number(props.getProperty('S1_CUSTOMER_FOLDER_NEXT_ROW') || 0);
  let driveIndex = options.driveIndex || null;

  if (!driveIndex) {
    driveIndex = customerFolder_buildExistingCustomerFolderIndex_();
  }

  // 저장된 다음 행만 믿지 않고 매번 Drive/시트를 다시 비교합니다.
  // 이전 행에서 정보가 수정되거나 중복 폴더가 새로 생겨도 continue가 놓치지 않습니다.
  const detected = options.detected || customerFolder_detectNextCustomerFolderWorkRowFromIndex_(sheet, headerMap, driveIndex);

  if (!detected.nextRow) {
    props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
    Logger.log('생성·재연결·중복정리가 필요한 고객사 폴더가 없습니다.');
    return detected;
  }

  row = detected.nextRow;
  props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(row));

  if (row > lastRow) {
    props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
    Logger.log('처리할 행이 없습니다. 이미 완료되었습니다.');

    return {
      status: 'DONE',
      nextRow: 0,
      message: '처리할 행이 없습니다. 이미 완료되었습니다.'
    };
  }

  const driveId = driveIndex.driveId;
  const lastCol = sheet.getLastColumn();
  const rowsToRead = Math.min(lastRow - row + 1, Number(cfg.MAX_ROWS_PER_RUN || 5000));
  const values = sheet.getRange(row, 1, rowsToRead, lastCol).getDisplayValues();
  const folderIdColIdx = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderId) - 1;

  let scanned = 0;
  let created = 0;
  let relinked = 0;
  let skipped = 0;
  let errors = 0;
  let duplicateCustomersFixed = 0;
  let duplicateFoldersTrashed = 0;
  let duplicateItemsMoved = 0;
  let duplicateItemsRenamed = 0;
  let identicalItemsTrashed = 0;
  let stoppedByTime = false;

  const logs = [];
  const parentCache = {};
  const pendingSheetUpdates = [];

  while (scanned < values.length) {
    if (Date.now() >= deadlineMs) {
      stoppedByTime = true;
      break;
    }

    if (!refreshCustomerFolderLock_(lock)) {
      throw new Error('고객사 폴더 soft lock 소유권을 잃어 안전하게 실행을 중단합니다.');
    }

    const rowNum = row + scanned;
    const rowData = values[scanned];

    try {
      const result = customerFolder_createOrRelinkCustomerFolderFastForRow_({
        sheet,
        rowNum,
        rowData,
        headerMap,
        driveId,
        driveIndex,
        folderIdColIdx,
        parentCache,
        deadlineMs,
        deferSheetWrite: true
      });

      const resultStatus = String(result.status || '');

      if (resultStatus === 'CREATED' || resultStatus === 'CREATED_IN_FAILED_FOLDER' || resultStatus.indexOf('RECREATED') !== -1) {
        created++;
      } else if (
        resultStatus.indexOf('RELINKED') !== -1 ||
        resultStatus.indexOf('REUSED') !== -1 ||
        resultStatus.indexOf('RENAMED') !== -1 ||
        resultStatus.indexOf('MOVED_TO_') !== -1 ||
        resultStatus.indexOf('DUPLICATES_MERGED') !== -1
      ) {
        relinked++;
      } else {
        skipped++;
      }

      if (result.sheetWrite) {
        pendingSheetUpdates.push(result.sheetWrite);
      }

      if (result.duplicateFoldersTrashed > 0) duplicateCustomersFixed++;
      duplicateFoldersTrashed += Number(result.duplicateFoldersTrashed || 0);
      duplicateItemsMoved += Number(result.duplicateItemsMoved || 0);
      duplicateItemsRenamed += Number(result.duplicateItemsRenamed || 0);
      identicalItemsTrashed += Number(result.identicalItemsTrashed || 0);

      logs.push([
        new Date(),
        rowNum,
        result.customerNo || '',
        result.company || '',
        result.vendor || '',
        result.folderName || '',
        result.folderId || '',
        result.status || '',
        result.message || ''
      ]);

    } catch (err) {
      if (err && err.code === 'CUSTOMER_FOLDER_TIME_BUDGET') {
        stoppedByTime = true;
        break;
      }

      errors++;
      logs.push([
        new Date(),
        rowNum,
        '',
        '',
        '',
        '',
        '',
        'ERROR',
        err && err.message ? err.message : String(err)
      ]);
    }

    scanned++;
  }

  customerFolder_flushFolderInfoBatch_(sheet, headerMap, pendingSheetUpdates);
  customerFolder_appendFolderLog_(logs);

  const nextRow = row + scanned;
  const hasMore = nextRow <= lastRow;

  if (hasMore) {
    props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(nextRow));

    const message =
      `이번 실행 완료: 스캔 ${scanned}건 / 신규생성 ${created}건 / 기존폴더보정 ${relinked}건 / ` +
      `스킵 ${skipped}건 / 오류 ${errors}건 / 중복고객정리 ${duplicateCustomersFixed}건 / ` +
      `중복폴더휴지통 ${duplicateFoldersTrashed}개 / 내부항목이동 ${duplicateItemsMoved}개 / ` +
      `충돌파일명변경 ${duplicateItemsRenamed}개 / 동일파일삭제 ${identicalItemsTrashed}개. ` +
      `${stoppedByTime ? '5분 처리시간에 도달했습니다. ' : ''}` +
      `다음 시작 행: ${nextRow}`;

    Logger.log(message);

    return {
      status: 'PARTIAL',
      scanned,
      created,
      relinked,
      skipped,
      errors,
      duplicateCustomersFixed,
      duplicateFoldersTrashed,
      duplicateItemsMoved,
      duplicateItemsRenamed,
      identicalItemsTrashed,
      nextRow,
      message
    };
  }

  props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');

  const message =
    `전체 완료: 스캔 ${scanned}건 / 신규생성 ${created}건 / 기존폴더보정 ${relinked}건 / ` +
    `스킵 ${skipped}건 / 오류 ${errors}건 / 중복고객정리 ${duplicateCustomersFixed}건 / ` +
    `중복폴더휴지통 ${duplicateFoldersTrashed}개 / 내부항목이동 ${duplicateItemsMoved}개 / ` +
    `충돌파일명변경 ${duplicateItemsRenamed}개 / 동일파일삭제 ${identicalItemsTrashed}개`;

  Logger.log(message);

  return {
    status: 'DONE',
    scanned,
    created,
    relinked,
    skipped,
    errors,
    duplicateCustomersFixed,
    duplicateFoldersTrashed,
    duplicateItemsMoved,
    duplicateItemsRenamed,
    identicalItemsTrashed,
    nextRow: 0,
    message
  };
}



/**
 * 영업전산/메일자동화에서 직접 호출할 함수.
 * 예: 고객 추가 저장 후 ensureCustomerFolderByCustomerNo(customerNo);
 */
function ensureCustomerFolderByCustomerNo(customerNo) {
  const lock = acquireCustomerFolderLockOrReturn_(
    'ensureCustomerFolderByCustomerNo',
    CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS
  );

  if (!lock) {
    return makeCustomerFolderLockedResult_('ensureCustomerFolderByCustomerNo');
  }

  try {
    return customerFolder_ensureCustomerFolderByCustomerNoLocked_(customerNo);
  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


function customerFolder_ensureCustomerFolderByCustomerNoLocked_(customerNo) {
const sheet = customerFolder_getMasterSheet_();

  let headerMap = customerFolder_getHeaderMap_(sheet);
  headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

  customerFolder_assertHeader_(headerMap, '고객번호');

  const target = customerFolder_cleanValue_(customerNo);
  if (!target) {
    throw new Error('고객번호가 비어 있습니다.');
  }

  const customerNoCol = customerFolder_col_(headerMap, '고객번호');
  const lastRow = sheet.getLastRow();

  if (lastRow < CUSTOMER_FOLDER_CFG.DATA_START_ROW) {
    throw new Error('마스터시트에 데이터가 없습니다.');
  }

  const values = sheet
    .getRange(
      CUSTOMER_FOLDER_CFG.DATA_START_ROW,
      customerNoCol,
      lastRow - CUSTOMER_FOLDER_CFG.DATA_START_ROW + 1,
      1
    )
    .getDisplayValues();

  const driveId = customerFolder_getSharedDriveId_();

  for (let i = 0; i < values.length; i++) {
    const rowCustomerNo = customerFolder_cleanValue_(values[i][0]);

    if (rowCustomerNo === target) {
      const rowNum = CUSTOMER_FOLDER_CFG.DATA_START_ROW + i;
      return customerFolder_ensureCustomerFolderForSheetRow_(sheet, rowNum, driveId, headerMap);
    }
  }

  throw new Error(`마스터시트에서 고객번호를 찾지 못했습니다: ${target}`);
}



/**
 * 마스터시트에서 사람이 직접 고객번호/회사명/수행사를 수정했을 때 자동 생성용.
 * installCustomerFolderOnEditTrigger()를 1번 실행해야 작동함.
 */
function handleCustomerFolderOnEdit(e) {
  if (!e || !e.range) return;

  const sheet = e.range.getSheet();
  const cfg = CUSTOMER_FOLDER_CFG;

  if (sheet.getName() !== cfg.MASTER_SHEET_NAME) return;

  const startRow = e.range.getRow();
  const endRow = startRow + e.range.getNumRows() - 1;

  if (endRow < cfg.DATA_START_ROW) return;

  let headerMap = customerFolder_getHeaderMap_(sheet);
  headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

  const targetCols = [
    customerFolder_col_(headerMap, '고객번호'),
    customerFolder_col_(headerMap, '회사명'),
    customerFolder_col_(headerMap, '수행사')
  ];

  // 상태값을 수주실패로 바꾸거나, 수주실패에서 되돌릴 때도
  // 고객사 폴더 위치가 루트 <-> 수주실패 폴더로 맞춰져야 함.
  const statusHeaderName = customerFolder_findFirstExistingHeaderName_(
    headerMap,
    FAILED_CUSTOMER_FOLDER_CFG.STATUS_HEADER_CANDIDATES
  );

  if (statusHeaderName) {
    targetCols.push(customerFolder_col_(headerMap, statusHeaderName));
  }

  const editStartCol = e.range.getColumn();
  const editEndCol = editStartCol + e.range.getNumColumns() - 1;

  const touched = targetCols.some(c => c >= editStartCol && c <= editEndCol);
  if (!touched) return;

  const driveId = customerFolder_getSharedDriveId_();

  for (let row = Math.max(startRow, cfg.DATA_START_ROW); row <= endRow; row++) {
    customerFolder_ensureCustomerFolderForSheetRow_(sheet, row, driveId, headerMap);
  }
}


// 고객사 폴더 생성 전용: 설치형 트리거로만 실행
function customerFolderInstallableOnEdit(e) {
  const lock = acquireCustomerFolderLockOrReturn_(
    'customerFolderInstallableOnEdit',
    CUSTOMER_FOLDER_CFG.ONEDIT_LOCK_WAIT_MILLIS
  );

  if (!lock) {
    return;
  }

  try {
    handleCustomerFolderOnEdit(e);
  } catch (err) {
    Logger.log('[customerFolderInstallableOnEdit 오류] ' + (err && err.stack ? err.stack : err));
  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


/**
 * 고객사 폴더 자동 생성용 설치형 onEdit 트리거 설치.
 * 최초 1회만 직접 실행.
 */
function installCustomerFolderOnEditTrigger() {
  const ss = customerFolder_getMasterSpreadsheet_();

  ScriptApp.getProjectTriggers()
    .filter(t => t.getHandlerFunction() === 'customerFolderInstallableOnEdit')
    .forEach(t => ScriptApp.deleteTrigger(t));

  ScriptApp.newTrigger('customerFolderInstallableOnEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  Logger.log('고객사 폴더 생성용 설치형 onEdit 트리거 설치 완료');
}


/***** 고객사 폴더 생성/재연결 내부 함수 *****/

function customerFolder_ensureCustomerFolderForSheetRow_(sheet, rowNum, driveId, headerMap) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const rowData = sheet.getRange(rowNum, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const customerNo = customerFolder_cleanValue_(rowData[customerFolder_col_(headerMap, '고객번호') - 1]);

  if (!customerNo) {
    return {
      status: 'SKIPPED',
      message: '고객번호 공란',
      customerNo: ''
    };
  }

  const driveIndex = customerFolder_buildTargetCustomerFolderIndex_(driveId, customerNo);

  return customerFolder_createOrRelinkCustomerFolderFastForRow_({
    sheet,
    rowNum,
    rowData,
    headerMap,
    driveId,
    driveIndex,
    folderIdColIdx: customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderId) - 1,
    parentCache: {},
    deadlineMs: Date.now() + (4 * 60 * 1000),
    deferSheetWrite: false
  });
}


function customerFolder_createOrRelinkCustomerFolderFastForRow_(params) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const sheet = params.sheet;
  const rowNum = params.rowNum;
  const rowData = params.rowData;
  const headerMap = params.headerMap;
  const driveId = params.driveId;
  const driveIndex = params.driveIndex;
  const folderIdColIdx = params.folderIdColIdx;
  const parentCache = params.parentCache || {};
  const deadlineMs = Number(params.deadlineMs || 0);

  customerFolder_assertBeforeDeadline_(deadlineMs, 1200);

  const customerNo = customerFolder_cleanValue_(rowData[customerFolder_col_(headerMap, '고객번호') - 1]);
  const company = customerFolder_cleanValue_(rowData[customerFolder_col_(headerMap, '회사명') - 1]);
  const vendorRaw = headerMap[customerFolder_normalizeHeader_('수행사')]
    ? customerFolder_cleanValue_(rowData[customerFolder_col_(headerMap, '수행사') - 1])
    : '';
  const vendor = vendorRaw || cfg.EMPTY_VENDOR_TEXT;

  if (!customerNo || !company) {
    return {
      status: 'SKIPPED',
      message: '고객번호 또는 회사명 공란',
      customerNo,
      company,
      vendor
    };
  }

  const expectedFolderName = customerFolder_buildCustomerFolderName_(customerNo, company, vendor);
  const customerNoKey = customerFolder_normalizeCustomerNoKey_(customerNo);
  const existingFolderId = customerFolder_cleanValue_(rowData[folderIdColIdx]);
  const desired = customerFolder_getCustomerDesiredParentInfoForRow_(rowData, headerMap, driveId, parentCache);

  let entries = customerFolder_getCustomerFolderEntriesFromIndex_(driveIndex, customerNoKey);
  let existingFolder = null;

  if (existingFolderId) {
    const indexedExisting = entries.find(entry => entry.folder && entry.folder.id === existingFolderId);

    if (indexedExisting) {
      existingFolder = indexedExisting.folder;
    } else {
      existingFolder = customerFolder_getDriveFile_(existingFolderId);

      const existingFolderCustomerNoKey = existingFolder
        ? customerFolder_extractCustomerNoKeyFromFolderName_(existingFolder.name)
        : '';

      if (
        existingFolder &&
        !existingFolder.trashed &&
        existingFolder.mimeType === 'application/vnd.google-apps.folder' &&
        existingFolderCustomerNoKey === customerNoKey
      ) {
        customerFolder_removeFolderIdFromDriveIndex_(driveIndex, existingFolder.id);
        entries.push(customerFolder_makeCustomerFolderIndexEntry_(existingFolder, driveId, driveIndex.failedParentFolderId));
      } else {
        // 다른 고객번호 폴더를 잘못 가리키는 ID는 병합 대상으로 사용하지 않습니다.
        existingFolder = null;
      }
    }
  }

  if (entries.length === 0) {
    const knownFailedParent = driveIndex.failedParentFolderId ? { id: driveIndex.failedParentFolderId } : null;
    const freshEntries = customerFolder_findAllCustomerFoldersByCustomerNoAnywhere_(driveId, customerNo, knownFailedParent);
    entries = freshEntries.slice();
  }

  entries = customerFolder_uniqueFolderEntries_(entries);

  let folder;
  let status;
  let mergeStats = customerFolder_emptyMergeStats_();

  if (entries.length > 0) {
    const canonicalEntry = customerFolder_chooseCanonicalCustomerFolder_(
      entries,
      existingFolderId,
      expectedFolderName,
      desired.location
    );

    folder = canonicalEntry.folder;

    if (existingFolderId && folder.id === existingFolderId) {
      status = 'EXISTING_ID';
    } else if (existingFolderId) {
      status = 'STALE_ID_RELINKED';
    } else {
      status = 'RELINKED_BY_CUSTOMER_NO';
    }

    if (entries.length > 1) {
      const mergeResult = customerFolder_mergeDuplicateCustomerFolders_({
        canonicalEntry,
        entries,
        driveId,
        deadlineMs
      });

      folder = mergeResult.folder || folder;
      mergeStats = mergeResult.stats || mergeStats;

      if (mergeStats.duplicateFoldersTrashed > 0) {
        status = customerFolder_appendCustomerFolderStatusPart_(status, 'DUPLICATES_MERGED');
      }

      if (mergeStats.partialFolders > 0) {
        status = customerFolder_appendCustomerFolderStatusPart_(status, 'DUPLICATE_MERGE_PARTIAL');
      }
    }

    if (cfg.RENAME_IF_CHANGED && folder.name !== expectedFolderName) {
      folder = customerFolder_renameDriveFile_(folder.id, expectedFolderName);
      status = customerFolder_appendCustomerFolderStatusPart_(status, 'RENAMED');
    }

    const parentResult = customerFolder_ensureCustomerFolderParentForRow_(
      folder,
      rowData,
      headerMap,
      driveId,
      parentCache
    );

    folder = parentResult.folder || folder;

    if (parentResult.moved) {
      status = customerFolder_appendCustomerFolderStatusPart_(status, parentResult.moveStatus);
    }

    customerFolder_setDriveIndexCustomerEntries_(driveIndex, customerNoKey, [{
      folder,
      location: parentResult.location || desired.location,
      parentId: parentResult.parentId || desired.parentId
    }]);

  } else {
    customerFolder_assertBeforeDeadline_(deadlineMs, 1500);

    // 실제 생성 직전 한 번 더 조회하여 외부 수동 생성 또는 경합에 의한 중복 생성을 차단합니다.
    const knownFailedParent = driveIndex.failedParentFolderId ? { id: driveIndex.failedParentFolderId } : null;
    const finalEntries = customerFolder_findAllCustomerFoldersByCustomerNoAnywhere_(driveId, customerNo, knownFailedParent);

    if (finalEntries.length > 0) {
      const canonicalEntry = customerFolder_chooseCanonicalCustomerFolder_(
        finalEntries,
        existingFolderId,
        expectedFolderName,
        desired.location
      );

      folder = canonicalEntry.folder;
      status = existingFolderId ? 'STALE_ID_FRESH_RELINKED' : 'FRESH_RELINKED_BY_CUSTOMER_NO';

      if (finalEntries.length > 1) {
        const mergeResult = customerFolder_mergeDuplicateCustomerFolders_({
          canonicalEntry,
          entries: finalEntries,
          driveId,
          deadlineMs
        });
        folder = mergeResult.folder || folder;
        mergeStats = mergeResult.stats || mergeStats;
        if (mergeStats.duplicateFoldersTrashed > 0) {
          status = customerFolder_appendCustomerFolderStatusPart_(status, 'DUPLICATES_MERGED');
        }
        if (mergeStats.partialFolders > 0) {
          status = customerFolder_appendCustomerFolderStatusPart_(status, 'DUPLICATE_MERGE_PARTIAL');
        }
      }

      if (cfg.RENAME_IF_CHANGED && folder.name !== expectedFolderName) {
        folder = customerFolder_renameDriveFile_(folder.id, expectedFolderName);
        status = customerFolder_appendCustomerFolderStatusPart_(status, 'RENAMED');
      }

      const parentResult = customerFolder_ensureCustomerFolderParentForRow_(
        folder,
        rowData,
        headerMap,
        driveId,
        parentCache
      );
      folder = parentResult.folder || folder;
      if (parentResult.moved) {
        status = customerFolder_appendCustomerFolderStatusPart_(status, parentResult.moveStatus);
      }

      customerFolder_setDriveIndexCustomerEntries_(driveIndex, customerNoKey, [{
        folder,
        location: parentResult.location || desired.location,
        parentId: parentResult.parentId || desired.parentId
      }]);

    } else {
      const parentId = desired.parentId;
      folder = customerFolder_createDriveFolder_(expectedFolderName, parentId);
      status = existingFolderId
        ? 'STALE_ID_RECREATED'
        : (parentId === driveId ? 'CREATED' : 'CREATED_IN_FAILED_FOLDER');

      customerFolder_setDriveIndexCustomerEntries_(driveIndex, customerNoKey, [{
        folder,
        location: desired.location,
        parentId
      }]);
      driveIndex.customerFolderCount = Number(driveIndex.customerFolderCount || 0) + 1;
    }
  }

  if (cfg.CREATE_STANDARD_SUBFOLDERS) {
    customerFolder_ensureStandardSubfolders_(folder.id, driveId);
  }

  const sheetWrite = customerFolder_makeFolderInfoWrite_(rowNum, folder, expectedFolderName, status);

  if (!params.deferSheetWrite) {
    customerFolder_writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, expectedFolderName, status);
  }

  const mergeMessage = mergeStats.duplicateFoldersTrashed > 0
    ? ` / 중복폴더 ${mergeStats.duplicateFoldersTrashed}개 병합·휴지통 이동` +
      ` / 내부항목 이동 ${mergeStats.itemsMoved}개` +
      ` / 파일명 충돌 보존 ${mergeStats.itemsRenamed}개` +
      ` / 동일파일 제거 ${mergeStats.identicalItemsTrashed}개`
    : '';

  return {
    status,
    message: '정상 처리' + mergeMessage,
    customerNo,
    company,
    vendor,
    folderName: expectedFolderName,
    folderId: folder.id,
    folderUrl: folder.webViewLink || `https://drive.google.com/drive/folders/${folder.id}`,
    duplicateFoldersTrashed: mergeStats.duplicateFoldersTrashed,
    duplicateItemsMoved: mergeStats.itemsMoved,
    duplicateItemsRenamed: mergeStats.itemsRenamed,
    identicalItemsTrashed: mergeStats.identicalItemsTrashed,
    sheetWrite
  };
}



function customerFolder_ensureStandardSubfolders_(customerFolderId, driveId) {
  CUSTOMER_FOLDER_CFG.STANDARD_SUBFOLDERS.forEach(name => {
    const existing = customerFolder_findChildFolder_(customerFolderId, driveId, name);
    if (!existing) {
      customerFolder_createDriveFolder_(name, customerFolderId);
    }
  });
}


function customerFolder_makeFolderInfoWrite_(rowNum, folder, folderName, status) {
  return {
    rowNum,
    folderId: folder.id,
    folderUrl: folder.webViewLink || `https://drive.google.com/drive/folders/${folder.id}`,
    folderName,
    status,
    updatedAt: Utilities.formatDate(new Date(), CUSTOMER_FOLDER_CFG.TZ, 'yyyy-MM-dd HH:mm:ss')
  };
}


function customerFolder_flushFolderInfoBatch_(sheet, headerMap, updates) {
  if (!updates || updates.length === 0) return;

  const cfg = CUSTOMER_FOLDER_CFG;
  const sorted = updates.slice().sort((a, b) => a.rowNum - b.rowNum);
  const groups = [];
  let group = [];

  sorted.forEach(update => {
    if (!group.length || update.rowNum === group[group.length - 1].rowNum + 1) {
      group.push(update);
    } else {
      groups.push(group);
      group = [update];
    }
  });
  if (group.length) groups.push(group);

  const fields = [
    ['folderId', cfg.OUTPUT_HEADERS.folderId],
    ['folderUrl', cfg.OUTPUT_HEADERS.folderUrl],
    ['folderName', cfg.OUTPUT_HEADERS.folderName],
    ['status', cfg.OUTPUT_HEADERS.folderStatus],
    ['updatedAt', cfg.OUTPUT_HEADERS.folderUpdatedAt]
  ];

  groups.forEach(items => {
    const startRow = items[0].rowNum;

    fields.forEach(([field, headerName]) => {
      sheet
        .getRange(startRow, customerFolder_col_(headerMap, headerName), items.length, 1)
        .setValues(items.map(item => [item[field] || '']));
    });
  });
}

function customerFolder_writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, folderName, status) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const write = customerFolder_makeFolderInfoWrite_(rowNum, folder, folderName, status);

  sheet.getRange(rowNum, customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderId)).setValue(write.folderId);
  sheet.getRange(rowNum, customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderUrl)).setValue(write.folderUrl);
  sheet.getRange(rowNum, customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderName)).setValue(write.folderName);
  sheet.getRange(rowNum, customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderStatus)).setValue(write.status);
  sheet.getRange(rowNum, customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderUpdatedAt)).setValue(write.updatedAt);
}


/***** 빠른 진행상황 탐지/Drive 색인 함수 *****/

function detectAndSetNextCustomerFolderRowFast() {
  const lock = acquireCustomerFolderLockOrReturn_(
    'detectAndSetNextCustomerFolderRowFast',
    CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS
  );

  if (!lock) {
    return makeCustomerFolderLockedResult_('detectAndSetNextCustomerFolderRowFast');
  }

  try {
    return customerFolder_detectAndSetNextCustomerFolderRowFastLocked_();
  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


function customerFolder_detectAndSetNextCustomerFolderRowFastLocked_() {
  const detected = customerFolder_detectNextCustomerFolderWorkRow_();
  const props = PropertiesService.getScriptProperties();

  if (detected.nextRow) {
    props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(detected.nextRow));
  } else {
    props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
  }

  logCustomerFolderDetectionResult_(detected);

  return detected;
}



/**
 * 수동 점검용.
 * 실행하면 지금 어디부터 생성해야 하는지 로그만 찍음.
 */
function reportCustomerFolderCreateProgressFast() {
  const detected = customerFolder_detectNextCustomerFolderWorkRow_();

  if (detected.nextRow) {
    Logger.log(
      `다음 작업 필요 행: ${detected.nextRow}행 / 고객번호 ${detected.customerNo || ''} / 회사명 ${detected.company || ''} / 사유: ${detected.reason || ''}`
    );
  } else {
    Logger.log('현재 생성 필요한 고객사 폴더가 없습니다.');
  }

  Logger.log(
    `Drive 탐지 고객폴더: ${detected.driveCustomerFolderCount}개 / ` +
    `시트 유효고객: ${detected.validCustomerCount}건 / ` +
    `시트 폴더ID 기재: ${detected.sheetFolderIdCount}건 / ` +
    `Drive 내 최대 고객번호: ${detected.maxCustomerNoInDrive || ''}`
  );

  return detected;
}


/**
 * 마스터시트와 공유드라이브를 비교해서 실제로 생성 필요한 첫 행을 찾음.
 * 기준은 Drive 고객번호_ 폴더 존재 여부 + 시트 관리값 누락 여부 + 수주실패 폴더 위치 일치 여부.
 */
function customerFolder_detectNextCustomerFolderWorkRow_() {
  const cfg = CUSTOMER_FOLDER_CFG;
  const sheet = customerFolder_getMasterSheet_();

  let headerMap = customerFolder_getHeaderMap_(sheet);
  headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

  customerFolder_assertHeader_(headerMap, '고객번호');
  customerFolder_assertHeader_(headerMap, '회사명');

  const driveIndex = customerFolder_buildExistingCustomerFolderIndex_();

  return customerFolder_detectNextCustomerFolderWorkRowFromIndex_(sheet, headerMap, driveIndex);
}


function customerFolder_detectNextCustomerFolderWorkRowFromIndex_(sheet, headerMap, driveIndex) {
  const cfg = CUSTOMER_FOLDER_CFG;

  headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

  customerFolder_assertHeader_(headerMap, '고객번호');
  customerFolder_assertHeader_(headerMap, '회사명');

  const lastRow = sheet.getLastRow();

  if (lastRow < cfg.DATA_START_ROW) {
    return {
      nextRow: 0,
      driveCustomerFolderCount: driveIndex ? driveIndex.customerFolderCount : 0,
      validCustomerCount: 0,
      sheetFolderIdCount: 0,
      maxCustomerNoInDrive: driveIndex ? driveIndex.maxCustomerNo : ''
    };
  }

  const lastCol = sheet.getLastColumn();
  const values = sheet
    .getRange(cfg.DATA_START_ROW, 1, lastRow - cfg.DATA_START_ROW + 1, lastCol)
    .getDisplayValues();

  const customerNoIdx = customerFolder_col_(headerMap, '고객번호') - 1;
  const companyIdx = customerFolder_col_(headerMap, '회사명') - 1;
  const vendorIdx = headerMap[customerFolder_normalizeHeader_('수행사')]
    ? customerFolder_col_(headerMap, '수행사') - 1
    : -1;
  const folderIdIdx = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderId) - 1;
  const folderUrlIdx = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderUrl) - 1;
  const folderNameIdx = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderName) - 1;
  const folderStatusIdx = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderStatus) - 1;

  let validCustomerCount = 0;
  let sheetFolderIdCount = 0;

  for (let i = 0; i < values.length; i++) {
    const rowData = values[i];
    const rowNum = cfg.DATA_START_ROW + i;

    const customerNo = customerFolder_cleanValue_(rowData[customerNoIdx]);
    const company = customerFolder_cleanValue_(rowData[companyIdx]);
    const vendorRaw = vendorIdx >= 0 ? customerFolder_cleanValue_(rowData[vendorIdx]) : '';
    const vendor = vendorRaw || cfg.EMPTY_VENDOR_TEXT;
    const expectedFolderName = customerFolder_buildCustomerFolderName_(customerNo, company, vendor);

    const folderId = customerFolder_cleanValue_(rowData[folderIdIdx]);
    const folderUrl = customerFolder_cleanValue_(rowData[folderUrlIdx]);
    const folderName = customerFolder_cleanValue_(rowData[folderNameIdx]);
    const folderStatus = customerFolder_cleanValue_(rowData[folderStatusIdx]);

    if (!customerNo || !company) {
      continue;
    }

    validCustomerCount++;

    if (folderId) {
      sheetFolderIdCount++;
    }

    const customerNoKey = customerFolder_normalizeCustomerNoKey_(customerNo);
    const mapped = driveIndex.byCustomerNo[customerNoKey];
    const allMapped = (driveIndex.allByCustomerNo && driveIndex.allByCustomerNo[customerNoKey]) || [];
    const desiredLocation = customerFolder_isFailedCustomerRowByData_(rowData, headerMap) ? 'FAILED' : 'ROOT';

    if (allMapped.length > 1) {
      return {
        nextRow: rowNum,
        customerNo,
        company,
        reason: `동일 고객번호 폴더 ${allMapped.length}개 중복 발견`,
        driveCustomerFolderCount: driveIndex.customerFolderCount,
        validCustomerCount,
        sheetFolderIdCount,
        maxCustomerNoInDrive: driveIndex.maxCustomerNo
      };
    }

    if (!mapped) {
      return {
        nextRow: rowNum,
        customerNo,
        company,
        reason: folderId
          ? '시트에는 고객사폴더ID가 있으나 Drive에 고객번호 prefix 폴더 없음'
          : 'Drive에 고객번호 prefix 폴더 없음',
        driveCustomerFolderCount: driveIndex.customerFolderCount,
        validCustomerCount,
        sheetFolderIdCount,
        maxCustomerNoInDrive: driveIndex.maxCustomerNo
      };
    }

    // Drive에는 폴더가 있는데 시트 관리값이 비었거나 오래된 경우도
    // 생성/재연결 루틴을 태워서 ID/URL/폴더명/처리상태를 복구해야 함.
    if (!folderId || !folderUrl || !folderName || !folderStatus) {
      return {
        nextRow: rowNum,
        customerNo,
        company,
        reason: 'Drive에는 고객번호 prefix 폴더가 있으나 시트 고객사폴더 관리값이 누락됨',
        driveCustomerFolderCount: driveIndex.customerFolderCount,
        validCustomerCount,
        sheetFolderIdCount,
        maxCustomerNoInDrive: driveIndex.maxCustomerNo
      };
    }

    if (folderName !== expectedFolderName) {
      return {
        nextRow: rowNum,
        customerNo,
        company,
        reason: `시트 고객사폴더명이 현재 고객정보 기준 폴더명과 다름: ${expectedFolderName}`,
        driveCustomerFolderCount: driveIndex.customerFolderCount,
        validCustomerCount,
        sheetFolderIdCount,
        maxCustomerNoInDrive: driveIndex.maxCustomerNo
      };
    }

    if (mapped.location && mapped.location !== desiredLocation) {
      return {
        nextRow: rowNum,
        customerNo,
        company,
        reason: desiredLocation === 'FAILED'
          ? '수주실패 상태이나 고객사 폴더가 루트에 있음'
          : '수주실패 상태가 아닌데 고객사 폴더가 수주실패 폴더 안에 있음',
        driveCustomerFolderCount: driveIndex.customerFolderCount,
        validCustomerCount,
        sheetFolderIdCount,
        maxCustomerNoInDrive: driveIndex.maxCustomerNo
      };
    }
  }

  return {
    nextRow: 0,
    driveCustomerFolderCount: driveIndex.customerFolderCount,
    validCustomerCount,
    sheetFolderIdCount,
    maxCustomerNoInDrive: driveIndex.maxCustomerNo
  };
}



/**
 * 공유드라이브 루트 + 수주실패 폴더 안의 고객사 폴더를 한 번에 색인.
 * 폴더명 앞의 숫자_ 를 고객번호로 봄.
 */
function customerFolder_buildExistingCustomerFolderIndex_() {
  const driveId = customerFolder_getSharedDriveId_();
  const byCustomerNo = {};
  const allByCustomerNo = {};
  let customerFolderCount = 0;
  let maxCustomerNo = 0;

  const rootFolders = customerFolder_listDirectChildFoldersPaged_(driveId, driveId);
  const failedParentFolder = customerFolder_getFailedParentFolderIfExistsFromRoot_(rootFolders, driveId);

  const addFolder = function (folder, location, parentId) {
    const customerNoKey = customerFolder_extractCustomerNoKeyFromFolderName_(folder.name);
    if (!customerNoKey) return;

    const entry = { folder, location, parentId };
    if (!allByCustomerNo[customerNoKey]) allByCustomerNo[customerNoKey] = [];
    allByCustomerNo[customerNoKey].push(entry);

    customerFolderCount++;
    maxCustomerNo = Math.max(maxCustomerNo, Number(customerNoKey) || 0);
  };

  rootFolders.forEach(folder => {
    if (failedParentFolder && folder.id === failedParentFolder.id) return;
    addFolder(folder, 'ROOT', driveId);
  });

  if (failedParentFolder && failedParentFolder.id) {
    const failedFolders = customerFolder_listDirectChildFoldersPaged_(failedParentFolder.id, driveId);
    failedFolders.forEach(folder => addFolder(folder, 'FAILED', failedParentFolder.id));
  }

  Object.keys(allByCustomerNo).forEach(customerNoKey => {
    // 기존 동작과 호환되도록 루트 우선, 그다음 오래된 폴더 우선.
    const sorted = allByCustomerNo[customerNoKey].slice().sort((a, b) => {
      if (a.location !== b.location) return a.location === 'ROOT' ? -1 : 1;
      return customerFolder_folderCreatedTimeMs_(a.folder) - customerFolder_folderCreatedTimeMs_(b.folder);
    });
    allByCustomerNo[customerNoKey] = sorted;
    byCustomerNo[customerNoKey] = sorted[0];
  });

  return {
    driveId,
    byCustomerNo,
    allByCustomerNo,
    customerFolderCount,
    duplicateCustomerNoCount: Object.keys(allByCustomerNo).filter(key => allByCustomerNo[key].length > 1).length,
    maxCustomerNo: maxCustomerNo ? String(maxCustomerNo) : '',
    rootFolderCount: rootFolders.length,
    failedParentFolderId: failedParentFolder ? failedParentFolder.id : ''
  };
}


function customerFolder_listDirectChildFoldersPaged_(parentFolderId, driveId) {
  const q = [
    `${customerFolder_driveQueryString_(parentFolderId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `trashed = false`
  ].join(' and ');

  let pageToken = '';
  const folders = [];

  do {
    let path =
      'files' +
      '?supportsAllDrives=true' +
      '&includeItemsFromAllDrives=true' +
      '&corpora=drive' +
      '&driveId=' + encodeURIComponent(driveId) +
      '&pageSize=1000' +
      '&q=' + encodeURIComponent(q) +
      '&fields=nextPageToken,files(id,name,webViewLink,trashed,parents,mimeType,createdTime,size,md5Checksum)';

    if (pageToken) {
      path += '&pageToken=' + encodeURIComponent(pageToken);
    }

    const data = customerFolder_driveFetch_(path, { method: 'get' });

    (data.files || []).forEach(file => folders.push(file));
    pageToken = data.nextPageToken || '';

  } while (pageToken);

  return folders;
}


function customerFolder_getFailedParentFolderIfExistsFromRoot_(rootFolders, driveId) {
  const failedName = FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME || '수주실패';

  const found = (rootFolders || []).find(folder => customerFolder_cleanValue_(folder.name) === failedName);

  if (found) {
    return found;
  }

  try {
    return customerFolder_findChildFolder_(driveId, driveId, failedName);
  } catch (err) {
    return null;
  }
}


function customerFolder_findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, expectedFolderName) {
  const entries = customerFolder_findAllCustomerFoldersByCustomerNoAnywhere_(driveId, customerNo);

  if (!entries.length) return null;

  const expected = customerFolder_cleanValue_(expectedFolderName);
  const exact = expected ? entries.find(entry => entry.folder.name === expected) : null;
  const chosen = exact || entries[0];

  return {
    status: chosen.location === 'FAILED' ? 'REUSED_BY_CUSTOMER_NO_FAILED' : 'REUSED_BY_CUSTOMER_NO_ROOT',
    folder: chosen.folder,
    location: chosen.location,
    parentId: chosen.parentId,
    duplicateCount: entries.length
  };
}


function customerFolder_extractCustomerNoKeyFromFolderName_(folderName) {
  const s = customerFolder_cleanValue_(folderName);
  const m = s.match(/^(\d+)_/);

  if (!m) return '';

  return customerFolder_normalizeCustomerNoKey_(m[1]);
}


function customerFolder_normalizeCustomerNoKey_(value) {
  return customerFolder_cleanValue_(value)
    .replace(/\.0$/, '')
    .replace(/,/g, '')
    .trim();
}


/***** 고객사 폴더명 일괄 업데이트 *****/

/***** 고객번호 기준 중복 폴더 병합/색인 유틸 *****/

function customerFolder_assertBeforeDeadline_(deadlineMs, reserveMs) {
  if (!deadlineMs) return;
  if (Date.now() + Number(reserveMs || 0) < deadlineMs) return;

  const err = new Error('고객사 폴더 작업의 5분 처리시간에 도달했습니다. 다음 실행에서 이어서 처리합니다.');
  err.code = 'CUSTOMER_FOLDER_TIME_BUDGET';
  throw err;
}


function customerFolder_emptyMergeStats_() {
  return {
    duplicateFoldersTrashed: 0,
    partialFolders: 0,
    itemsMoved: 0,
    itemsRenamed: 0,
    identicalItemsTrashed: 0,
    nestedFoldersMerged: 0
  };
}


function customerFolder_addMergeStats_(target, source) {
  Object.keys(target).forEach(key => {
    target[key] = Number(target[key] || 0) + Number(source[key] || 0);
  });
  return target;
}


function customerFolder_folderCreatedTimeMs_(folder) {
  const t = folder && folder.createdTime ? new Date(folder.createdTime).getTime() : 0;
  return Number.isFinite(t) && t > 0 ? t : Number.MAX_SAFE_INTEGER;
}


function customerFolder_uniqueFolderEntries_(entries) {
  const seen = {};
  const result = [];

  (entries || []).forEach(entry => {
    const id = entry && entry.folder ? entry.folder.id : '';
    if (!id || seen[id]) return;
    seen[id] = true;
    result.push(entry);
  });

  return result;
}


function customerFolder_getCustomerFolderEntriesFromIndex_(driveIndex, customerNoKey) {
  if (!driveIndex) return [];
  if (driveIndex.allByCustomerNo && driveIndex.allByCustomerNo[customerNoKey]) {
    return driveIndex.allByCustomerNo[customerNoKey].slice();
  }
  if (driveIndex.byCustomerNo && driveIndex.byCustomerNo[customerNoKey]) {
    return [driveIndex.byCustomerNo[customerNoKey]];
  }
  return [];
}


function customerFolder_setDriveIndexCustomerEntries_(driveIndex, customerNoKey, entries) {
  if (!driveIndex.allByCustomerNo) driveIndex.allByCustomerNo = {};
  if (!driveIndex.byCustomerNo) driveIndex.byCustomerNo = {};

  const unique = customerFolder_uniqueFolderEntries_(entries);
  driveIndex.allByCustomerNo[customerNoKey] = unique;

  if (unique.length) {
    driveIndex.byCustomerNo[customerNoKey] = unique[0];
  } else {
    delete driveIndex.byCustomerNo[customerNoKey];
  }
}


function customerFolder_removeFolderIdFromDriveIndex_(driveIndex, folderId) {
  if (!driveIndex || !folderId) return;

  Object.keys(driveIndex.allByCustomerNo || {}).forEach(key => {
    const filtered = (driveIndex.allByCustomerNo[key] || []).filter(entry => {
      return !(entry.folder && entry.folder.id === folderId);
    });

    if (filtered.length) {
      driveIndex.allByCustomerNo[key] = filtered;
      driveIndex.byCustomerNo[key] = filtered[0];
    } else {
      delete driveIndex.allByCustomerNo[key];
      delete driveIndex.byCustomerNo[key];
    }
  });
}


function customerFolder_makeCustomerFolderIndexEntry_(folder, driveId, failedParentFolderId) {
  const parents = folder.parents || [];
  let location = '';
  let parentId = parents.length ? parents[0] : '';

  if (failedParentFolderId && parents.indexOf(failedParentFolderId) !== -1) {
    location = 'FAILED';
    parentId = failedParentFolderId;
  } else if (parents.indexOf(driveId) !== -1) {
    location = 'ROOT';
    parentId = driveId;
  }

  return { folder, location, parentId };
}


function customerFolder_chooseCanonicalCustomerFolder_(entries, existingFolderId, expectedFolderName, desiredLocation) {
  const sorted = customerFolder_uniqueFolderEntries_(entries).slice().sort((a, b) => {
    const score = function (entry) {
      let value = 0;
      if (existingFolderId && entry.folder.id === existingFolderId) value += 10000;
      if (entry.folder.name === expectedFolderName) value += 1000;
      if (entry.location === desiredLocation) value += 100;
      return value;
    };

    const diff = score(b) - score(a);
    if (diff) return diff;

    const timeDiff = customerFolder_folderCreatedTimeMs_(a.folder) - customerFolder_folderCreatedTimeMs_(b.folder);
    if (timeDiff) return timeDiff;

    return String(a.folder.id).localeCompare(String(b.folder.id));
  });

  if (!sorted.length) {
    throw new Error('기준 고객사 폴더를 선택할 수 없습니다.');
  }

  return sorted[0];
}


function customerFolder_buildTargetCustomerFolderIndex_(driveId, customerNo) {
  const failedParent = customerFolder_findChildFolder_(
    driveId,
    driveId,
    FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME
  );
  const entries = customerFolder_findAllCustomerFoldersByCustomerNoAnywhere_(driveId, customerNo, failedParent);
  const key = customerFolder_normalizeCustomerNoKey_(customerNo);
  const allByCustomerNo = {};
  const byCustomerNo = {};

  allByCustomerNo[key] = entries;
  if (entries.length) byCustomerNo[key] = entries[0];

  return {
    driveId,
    byCustomerNo,
    allByCustomerNo,
    customerFolderCount: entries.length,
    maxCustomerNo: key,
    rootFolderCount: 0,
    failedParentFolderId: failedParent ? failedParent.id : ''
  };
}


function customerFolder_findAllCustomerFoldersByCustomerNoAnywhere_(driveId, customerNo, failedParentFolder) {
  const result = [];
  const rootFolders = customerFolder_findCustomerFoldersByCustomerNoPrefixInParent_(driveId, driveId, customerNo);
  rootFolders.forEach(folder => result.push({ folder, location: 'ROOT', parentId: driveId }));

  const failedParent = arguments.length >= 3
    ? failedParentFolder
    : customerFolder_findChildFolder_(
        driveId,
        driveId,
        FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME
      );

  if (failedParent && failedParent.id) {
    const failedFolders = customerFolder_findCustomerFoldersByCustomerNoPrefixInParent_(
      driveId,
      failedParent.id,
      customerNo
    );
    failedFolders.forEach(folder => result.push({ folder, location: 'FAILED', parentId: failedParent.id }));
  }

  return customerFolder_uniqueFolderEntries_(result);
}


function customerFolder_findCustomerFoldersByCustomerNoPrefixInParent_(driveId, parentFolderId, customerNo) {
  const prefix = customerFolder_sanitizeFolderPart_(customerFolder_normalizeCustomerNoKey_(customerNo) || customerNo) + '_';
  const q = [
    `${customerFolder_driveQueryString_(parentFolderId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `name contains ${customerFolder_driveQueryString_(prefix)}`,
    `trashed = false`
  ].join(' and ');

  let pageToken = '';
  const folders = [];

  do {
    let path =
      'files' +
      '?supportsAllDrives=true' +
      '&includeItemsFromAllDrives=true' +
      '&corpora=drive' +
      '&driveId=' + encodeURIComponent(driveId) +
      '&pageSize=1000' +
      '&q=' + encodeURIComponent(q) +
      '&fields=nextPageToken,files(id,name,webViewLink,trashed,mimeType,parents,createdTime,size,md5Checksum)';

    if (pageToken) path += '&pageToken=' + encodeURIComponent(pageToken);

    const data = customerFolder_driveFetch_(path, { method: 'get' });
    (data.files || []).forEach(file => {
      if (customerFolder_cleanValue_(file.name).startsWith(prefix)) folders.push(file);
    });
    pageToken = data.nextPageToken || '';
  } while (pageToken);

  return folders;
}


function customerFolder_mergeDuplicateCustomerFolders_(params) {
  const canonicalEntry = params.canonicalEntry;
  const entries = customerFolder_uniqueFolderEntries_(params.entries || []);
  const driveId = params.driveId;
  const deadlineMs = params.deadlineMs;
  const stats = customerFolder_emptyMergeStats_();
  const canonicalFolder = canonicalEntry.folder;

  entries.forEach(entry => {
    if (!entry.folder || entry.folder.id === canonicalFolder.id) return;

    customerFolder_assertBeforeDeadline_(deadlineMs, 1800);

    const mergeResult = customerFolder_mergeFolderContentsRecursive_(
      entry.folder.id,
      canonicalFolder.id,
      driveId,
      deadlineMs,
      0
    );

    customerFolder_addMergeStats_(stats, mergeResult.stats);

    const remaining = customerFolder_listChildrenPaged_(entry.folder.id, driveId, 1);

    if (remaining.length === 0) {
      customerFolder_trashDriveFile_(entry.folder.id);
      stats.duplicateFoldersTrashed++;
    } else {
      stats.partialFolders++;
    }
  });

  return {
    folder: canonicalFolder,
    stats
  };
}


function customerFolder_mergeFolderContentsRecursive_(sourceFolderId, targetFolderId, driveId, deadlineMs, depth) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const stats = customerFolder_emptyMergeStats_();

  if (depth > Number(cfg.DUPLICATE_MERGE_MAX_DEPTH || 25)) {
    throw new Error(`중복 폴더 병합 최대 깊이를 초과했습니다: ${sourceFolderId}`);
  }

  const sourceChildren = customerFolder_listChildrenPaged_(sourceFolderId, driveId);
  const targetChildren = customerFolder_listChildrenPaged_(targetFolderId, driveId);
  const targetByKey = {};
  const targetByName = {};
  const usedNames = {};

  targetChildren.forEach(item => {
    const key = customerFolder_childMatchKey_(item);
    if (!targetByKey[key]) targetByKey[key] = [];
    if (!targetByName[item.name]) targetByName[item.name] = [];
    targetByKey[key].push(item);
    targetByName[item.name].push(item);
    usedNames[item.name] = true;
  });

  sourceChildren.forEach(item => {
    customerFolder_assertBeforeDeadline_(deadlineMs, 1500);

    const isFolder = item.mimeType === 'application/vnd.google-apps.folder';
    const key = customerFolder_childMatchKey_(item);
    const sameTypeAndName = targetByKey[key] || [];
    const sameNameAnyType = targetByName[item.name] || [];

    if (isFolder && sameTypeAndName.length > 0) {
      const nested = customerFolder_mergeFolderContentsRecursive_(
        item.id,
        sameTypeAndName[0].id,
        driveId,
        deadlineMs,
        depth + 1
      );
      customerFolder_addMergeStats_(stats, nested.stats);

      const remains = customerFolder_listChildrenPaged_(item.id, driveId, 1);
      if (remains.length === 0) {
        customerFolder_trashDriveFile_(item.id);
        stats.nestedFoldersMerged++;
      }
      return;
    }

    if (!isFolder && sameTypeAndName.length > 0) {
      const identical = sameTypeAndName.find(target => customerFolder_areFilesIdentical_(item, target));

      if (identical) {
        customerFolder_trashDriveFile_(item.id);
        stats.identicalItemsTrashed++;
        return;
      }
    }

    // 이름이 같은 서로 다른 파일/폴더는 덮어쓰지 않고 이름을 바꿔 모두 보존합니다.
    if (sameNameAnyType.length > 0) {
      const uniqueName = customerFolder_buildMergedUniqueName_(item.name, usedNames);
      item = customerFolder_renameDriveFile_(item.id, uniqueName);
      item.parents = item.parents || [sourceFolderId];
      stats.itemsRenamed++;
    }

    customerFolder_moveDriveFileToFolder_(item.id, targetFolderId, item.parents || [sourceFolderId]);
    stats.itemsMoved++;
    usedNames[item.name] = true;

    const movedKey = customerFolder_childMatchKey_(item);
    if (!targetByKey[movedKey]) targetByKey[movedKey] = [];
    if (!targetByName[item.name]) targetByName[item.name] = [];
    targetByKey[movedKey].push(item);
    targetByName[item.name].push(item);
  });

  return { stats };
}


function customerFolder_childMatchKey_(item) {
  return customerFolder_cleanValue_(item.name) + '\u0000' + customerFolder_cleanValue_(item.mimeType);
}


function customerFolder_areFilesIdentical_(a, b) {
  if (!a || !b) return false;
  if (a.mimeType !== b.mimeType) return false;
  if (!a.md5Checksum || !b.md5Checksum) return false;
  if (a.md5Checksum !== b.md5Checksum) return false;

  const aSize = customerFolder_cleanValue_(a.size);
  const bSize = customerFolder_cleanValue_(b.size);
  return !aSize || !bSize || aSize === bSize;
}


function customerFolder_buildMergedUniqueName_(originalName, usedNames) {
  const suffix = CUSTOMER_FOLDER_CFG.DUPLICATE_MERGE_SUFFIX || '중복폴더병합';
  const name = customerFolder_cleanValue_(originalName) || '이름없음';
  const dot = name.lastIndexOf('.');
  const hasExtension = dot > 0 && dot < name.length - 1;
  const base = hasExtension ? name.slice(0, dot) : name;
  const ext = hasExtension ? name.slice(dot) : '';

  let n = 1;
  let candidate = `${base}__${suffix}_${n}${ext}`;

  while (usedNames[candidate]) {
    n++;
    candidate = `${base}__${suffix}_${n}${ext}`;
  }

  usedNames[candidate] = true;
  return candidate;
}


function customerFolder_listChildrenPaged_(parentFolderId, driveId, maxResults) {
  const q = [
    `${customerFolder_driveQueryString_(parentFolderId)} in parents`,
    `trashed = false`
  ].join(' and ');

  let pageToken = '';
  const files = [];
  const limit = Number(maxResults || 0);

  do {
    let path =
      'files' +
      '?supportsAllDrives=true' +
      '&includeItemsFromAllDrives=true' +
      '&corpora=drive' +
      '&driveId=' + encodeURIComponent(driveId) +
      '&pageSize=1000' +
      '&q=' + encodeURIComponent(q) +
      '&fields=nextPageToken,files(id,name,webViewLink,trashed,mimeType,parents,createdTime,size,md5Checksum)';

    if (pageToken) path += '&pageToken=' + encodeURIComponent(pageToken);

    const data = customerFolder_driveFetch_(path, { method: 'get' });
    (data.files || []).forEach(file => {
      if (!limit || files.length < limit) files.push(file);
    });

    if (limit && files.length >= limit) break;
    pageToken = data.nextPageToken || '';
  } while (pageToken);

  return files;
}


/**
 * 수동 점검용: 고객번호가 같은 폴더가 두 개 이상인 건만 로그로 출력합니다.
 * 실제 병합/삭제는 하지 않습니다.
 */
function reportDuplicateCustomerFoldersFast() {
  const index = customerFolder_buildExistingCustomerFolderIndex_();
  const duplicates = [];

  Object.keys(index.allByCustomerNo || {}).forEach(customerNoKey => {
    const entries = index.allByCustomerNo[customerNoKey] || [];
    if (entries.length < 2) return;

    duplicates.push({
      customerNo: customerNoKey,
      count: entries.length,
      folders: entries.map(entry => ({
        id: entry.folder.id,
        name: entry.folder.name,
        location: entry.location
      }))
    });
  });

  if (!duplicates.length) {
    Logger.log('고객번호 기준 중복 고객사 폴더가 없습니다.');
    return { status: 'DONE', duplicateCustomerCount: 0, duplicates: [] };
  }

  duplicates.forEach(item => {
    Logger.log(
      `중복 고객번호 ${item.customerNo}: ${item.count}개 / ` +
      item.folders.map(folder => `[${folder.location}] ${folder.name} (${folder.id})`).join(' | ')
    );
  });

  Logger.log(`중복 고객번호 총 ${duplicates.length}건`);
  return {
    status: 'FOUND',
    duplicateCustomerCount: duplicates.length,
    duplicates
  };
}


function manualUpdateAllCustomerFolderNames() {
  return initCreateCustomerFoldersFromMaster();
}

function continueUpdateAllCustomerFolderNames() {
  return continueCreateCustomerFoldersFromMaster();
}

/***** 고객사 폴더 내부 하위폴더 정리 *****/

function previewTrashCustomerChildFolders() {
  return customerFolder_runTrashCustomerChildFolders_({
    dryRun: true,
    onlyStandardSubfolders: false
  });
}


function trashCustomerChildFolders() {
  return customerFolder_runTrashCustomerChildFolders_({
    dryRun: false,
    onlyStandardSubfolders: false
  });
}


function trashOnlyStandardCustomerChildFolders() {
  return customerFolder_runTrashCustomerChildFolders_({
    dryRun: false,
    onlyStandardSubfolders: true
  });
}


function customerFolder_runTrashCustomerChildFolders_(options) {
  const lock = acquireCustomerFolderLockOrReturn_('customerFolder_runTrashCustomerChildFolders_', CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS);

  if (!lock) {
    return makeCustomerFolderLockedResult_('customerFolder_runTrashCustomerChildFolders_');
  }

  try {
    const cfg = CUSTOMER_FOLDER_CFG;
    const sheet = customerFolder_getMasterSheet_();

    let headerMap = customerFolder_getHeaderMap_(sheet);
    headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);

    customerFolder_assertHeader_(headerMap, '고객번호');
    customerFolder_assertHeader_(headerMap, '회사명');

    const driveId = customerFolder_getSharedDriveId_();
    const lastRow = sheet.getLastRow();

    let processedCustomers = 0;
    let targetFolders = 0;
    let trashedFolders = 0;
    let skipped = 0;
    let errors = 0;

    const logs = [];
    const standardSet = new Set(cfg.STANDARD_SUBFOLDERS || []);

    for (let rowNum = cfg.DATA_START_ROW; rowNum <= lastRow; rowNum++) {
      try {
        const row = sheet
          .getRange(rowNum, 1, 1, sheet.getLastColumn())
          .getDisplayValues()[0];

        const customerNo = customerFolder_cleanValue_(row[customerFolder_col_(headerMap, '고객번호') - 1]);
        const company = customerFolder_cleanValue_(row[customerFolder_col_(headerMap, '회사명') - 1]);

        if (!customerNo || !company) {
          skipped++;

          logs.push([
            new Date(),
            rowNum,
            customerNo,
            company,
            '',
            '',
            '',
            'SKIPPED',
            '고객번호 또는 회사명 공란'
          ]);

          continue;
        }

        const customerFolder = customerFolder_getCustomerFolderForCleanupRow_(row, rowNum, driveId, headerMap, customerNo);

        if (!customerFolder) {
          skipped++;

          logs.push([
            new Date(),
            rowNum,
            customerNo,
            company,
            '',
            '',
            '',
            'NOT_FOUND',
            '고객사 폴더를 찾지 못함'
          ]);

          continue;
        }

        processedCustomers++;

        let childFolders = customerFolder_listChildFolders_(customerFolder.id, driveId);

        if (options.onlyStandardSubfolders) {
          childFolders = childFolders.filter(f => standardSet.has(f.name));
        }

        if (childFolders.length === 0) {
          logs.push([
            new Date(),
            rowNum,
            customerNo,
            company,
            '',
            customerFolder.name,
            customerFolder.id,
            'NO_CHILD_FOLDER',
            '삭제 대상 하위폴더 없음'
          ]);

          continue;
        }

        childFolders.forEach(child => {
          targetFolders++;

          if (options.dryRun) {
            logs.push([
              new Date(),
              rowNum,
              customerNo,
              company,
              '',
              child.name,
              child.id,
              'PREVIEW',
              `삭제 예정 하위폴더 / 고객사폴더: ${customerFolder.name}`
            ]);
          } else {
            customerFolder_trashDriveFile_(child.id);
            trashedFolders++;

            logs.push([
              new Date(),
              rowNum,
              customerNo,
              company,
              '',
              child.name,
              child.id,
              'TRASHED',
              `하위폴더 휴지통 이동 완료 / 고객사폴더: ${customerFolder.name}`
            ]);
          }
        });

      } catch (err) {
        errors++;

        logs.push([
          new Date(),
          rowNum,
          '',
          '',
          '',
          '',
          '',
          'ERROR',
          err && err.message ? err.message : String(err)
        ]);
      }
    }

    customerFolder_appendFolderLog_(logs);

    const message =
      `${options.dryRun ? '미리보기' : '삭제'} 완료: ` +
      `처리 고객폴더 ${processedCustomers}개 / 대상 하위폴더 ${targetFolders}개 / 휴지통 이동 ${trashedFolders}개 / 스킵 ${skipped}건 / 오류 ${errors}건`;

    Logger.log(message);
    return {
      status: 'DONE',
      processedCustomers,
      targetFolders,
      trashedFolders,
      skipped,
      errors,
      message
    };

  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


function customerFolder_getCustomerFolderForCleanupRow_(row, rowNum, driveId, headerMap, customerNo) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const folderIdCol = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderId);
  const existingFolderId = customerFolder_cleanValue_(row[folderIdCol - 1]);

  let folder = null;

  if (existingFolderId) {
    folder = customerFolder_getDriveFile_(existingFolderId);

    if (folder && folder.trashed) {
      folder = null;
    }

    if (folder && folder.mimeType !== 'application/vnd.google-apps.folder') {
      folder = null;
    }
  }

  if (!folder) {
    const found = customerFolder_findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, '');
    if (found && found.folder) {
      folder = found.folder;
    }
  }

  return folder;
}


function customerFolder_listChildFolders_(parentFolderId, driveId) {
  return customerFolder_listChildrenPaged_(parentFolderId, driveId).filter(file => {
    return file.mimeType === 'application/vnd.google-apps.folder';
  });
}


function customerFolder_trashDriveFile_(fileId) {
  return customerFolder_driveFetch_(
    'files/' + encodeURIComponent(fileId) + '?supportsAllDrives=true&fields=id,name,trashed',
    {
      method: 'patch',
      payload: {
        trashed: true
      }
    }
  );
}


/***** 수주실패 고객사 폴더 이동 *****/

function manualMoveFailedCustomerFolders() {
  PropertiesService.getScriptProperties().deleteProperty(FAILED_CUSTOMER_FOLDER_CFG.PROP_NEXT_ROW);
  return continueMoveFailedCustomerFolders();
}


function continueMoveFailedCustomerFolders() {
  const lock = acquireCustomerFolderLockOrReturn_('continueMoveFailedCustomerFolders', CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS);

  if (!lock) {
    return makeCustomerFolderLockedResult_('continueMoveFailedCustomerFolders');
  }

  try {
    const cfg = CUSTOMER_FOLDER_CFG;
    const failedCfg = FAILED_CUSTOMER_FOLDER_CFG;

    const sheet = customerFolder_getMasterSheet_();

    let headerMap = customerFolder_getHeaderMap_(sheet);
    headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);
    headerMap = customerFolder_ensureFailedFolderOutputHeaders_(sheet, headerMap);

    customerFolder_assertHeader_(headerMap, '고객번호');
    customerFolder_assertHeader_(headerMap, '회사명');

    const statusHeaderName = customerFolder_findFirstExistingHeaderName_(headerMap, failedCfg.STATUS_HEADER_CANDIDATES);
    if (!statusHeaderName) {
      throw new Error(
        '상태값 헤더를 찾지 못했습니다. 후보: ' +
        failedCfg.STATUS_HEADER_CANDIDATES.join(', ')
      );
    }

    const driveId = customerFolder_getSharedDriveId_();
    const failedParentFolder = customerFolder_ensureFailedParentFolder_(driveId);

    const props = PropertiesService.getScriptProperties();
    const lastRow = sheet.getLastRow();

    let row = Number(props.getProperty(failedCfg.PROP_NEXT_ROW) || cfg.DATA_START_ROW);

    if (row > lastRow) {
      props.deleteProperty(failedCfg.PROP_NEXT_ROW);
      Logger.log('수주실패 폴더 이동 대상 행이 없습니다.');
      return;
    }

    let processed = 0;
    let moved = 0;
    let already = 0;
    let skipped = 0;
    let notFound = 0;
    let errors = 0;

    const startedAt = Date.now();
    const logs = [];

    while (
      row <= lastRow &&
      processed < failedCfg.MAX_ROWS_PER_RUN &&
      Date.now() - startedAt < failedCfg.MAX_MILLIS_PER_RUN
    ) {
      try {
        const result = customerFolder_moveFailedCustomerFolderForRow_({
          sheet,
          rowNum: row,
          headerMap,
          statusHeaderName,
          driveId,
          failedParentFolderId: failedParentFolder.id
        });

        if (result.status === 'MOVED') moved++;
        else if (result.status === 'ALREADY_IN_FAILED_FOLDER') already++;
        else if (result.status === 'SKIPPED') skipped++;
        else if (result.status === 'NOT_FOUND') notFound++;

        logs.push([
          new Date(),
          row,
          result.customerNo || '',
          result.company || '',
          result.vendor || '',
          result.folderName || '',
          result.folderId || '',
          result.status || '',
          result.message || ''
        ]);

      } catch (err) {
        errors++;

        logs.push([
          new Date(),
          row,
          '',
          '',
          '',
          '',
          '',
          'ERROR',
          err && err.message ? err.message : String(err)
        ]);
      }

      row++;
      processed++;
    }

    customerFolder_appendFolderLog_(logs);

    if (row <= lastRow) {
      props.setProperty(failedCfg.PROP_NEXT_ROW, String(row));

      Logger.log(
        `수주실패 폴더 이동 이번 실행 완료: 처리 ${processed}건 / 이동 ${moved}건 / 이미이동 ${already}건 / 스킵 ${skipped}건 / 미발견 ${notFound}건 / 오류 ${errors}건. ` +
        `아직 남았습니다. continueMoveFailedCustomerFolders()를 다시 실행하세요. 다음 시작 행: ${row}`
      );

    } else {
      props.deleteProperty(failedCfg.PROP_NEXT_ROW);

      Logger.log(
        `수주실패 폴더 이동 전체 완료: 처리 ${processed}건 / 이동 ${moved}건 / 이미이동 ${already}건 / 스킵 ${skipped}건 / 미발견 ${notFound}건 / 오류 ${errors}건`
      );
    }

  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


function moveFailedCustomerFolderByCustomerNo(customerNo) {
  const lock = acquireCustomerFolderLockOrReturn_(
    'moveFailedCustomerFolderByCustomerNo',
    CUSTOMER_FOLDER_CFG.LOCK_WAIT_MILLIS
  );

  if (!lock) {
    return makeCustomerFolderLockedResult_('moveFailedCustomerFolderByCustomerNo');
  }

  try {
    return customerFolder_moveFailedCustomerFolderByCustomerNoLocked_(customerNo);
  } finally {
    releaseCustomerFolderLock_(lock);
  }
}


function customerFolder_moveFailedCustomerFolderByCustomerNoLocked_(customerNo) {
const sheet = customerFolder_getMasterSheet_();

  let headerMap = customerFolder_getHeaderMap_(sheet);
  headerMap = customerFolder_ensureOutputHeaders_(sheet, headerMap);
  headerMap = customerFolder_ensureFailedFolderOutputHeaders_(sheet, headerMap);

  customerFolder_assertHeader_(headerMap, '고객번호');

  const statusHeaderName = customerFolder_findFirstExistingHeaderName_(
    headerMap,
    FAILED_CUSTOMER_FOLDER_CFG.STATUS_HEADER_CANDIDATES
  );

  if (!statusHeaderName) {
    throw new Error('상태값 헤더를 찾지 못했습니다.');
  }

  const target = customerFolder_cleanValue_(customerNo);
  if (!target) {
    throw new Error('고객번호가 비어 있습니다.');
  }

  const customerNoCol = customerFolder_col_(headerMap, '고객번호');
  const lastRow = sheet.getLastRow();

  const values = sheet
    .getRange(
      CUSTOMER_FOLDER_CFG.DATA_START_ROW,
      customerNoCol,
      lastRow - CUSTOMER_FOLDER_CFG.DATA_START_ROW + 1,
      1
    )
    .getDisplayValues();

  const driveId = customerFolder_getSharedDriveId_();
  const failedParentFolder = customerFolder_ensureFailedParentFolder_(driveId);

  for (let i = 0; i < values.length; i++) {
    const rowCustomerNo = customerFolder_cleanValue_(values[i][0]);

    if (rowCustomerNo === target) {
      const rowNum = CUSTOMER_FOLDER_CFG.DATA_START_ROW + i;

      return customerFolder_moveFailedCustomerFolderForRow_({
        sheet,
        rowNum,
        headerMap,
        statusHeaderName,
        driveId,
        failedParentFolderId: failedParentFolder.id
      });
    }
  }

  throw new Error(`마스터시트에서 고객번호를 찾지 못했습니다: ${target}`);
}



function customerFolder_moveFailedCustomerFolderForRow_(params) {
  const sheet = params.sheet;
  const rowNum = params.rowNum;
  const headerMap = params.headerMap;
  const statusHeaderName = params.statusHeaderName;
  const driveId = params.driveId;
  const failedParentFolderId = params.failedParentFolderId;

  const row = sheet
    .getRange(rowNum, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const customerNo = customerFolder_cleanValue_(row[customerFolder_col_(headerMap, '고객번호') - 1]);
  const company = customerFolder_cleanValue_(row[customerFolder_col_(headerMap, '회사명') - 1]);

  let vendor = '';
  if (headerMap[customerFolder_normalizeHeader_('수행사')]) {
    vendor = customerFolder_cleanValue_(row[customerFolder_col_(headerMap, '수행사') - 1]);
  }

  const statusValue = customerFolder_cleanValue_(row[customerFolder_col_(headerMap, statusHeaderName) - 1]);

  if (!customerNo || !company) {
    return {
      status: 'SKIPPED',
      message: '고객번호 또는 회사명 공란',
      customerNo,
      company,
      vendor
    };
  }

  if (!customerFolder_isFailedStatus_(statusValue)) {
    customerFolder_writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
      status: 'SKIPPED_NOT_FAILED_STATUS'
    });

    return {
      status: 'SKIPPED',
      message: `수주실패 상태 아님: ${statusValue}`,
      customerNo,
      company,
      vendor
    };
  }

  const folder = customerFolder_getCustomerFolderForFailedMove_(row, driveId, headerMap, customerNo);

  if (!folder) {
    customerFolder_writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
      status: 'NOT_FOUND'
    });

    return {
      status: 'NOT_FOUND',
      message: '고객사 폴더를 찾지 못함',
      customerNo,
      company,
      vendor
    };
  }

  const expectedFolderName = customerFolder_buildCustomerFolderName_(
    customerNo,
    company,
    vendor || CUSTOMER_FOLDER_CFG.EMPTY_VENDOR_TEXT
  );

  let currentFolder = folder;

  if (currentFolder.name !== expectedFolderName) {
    currentFolder = customerFolder_renameDriveFile_(currentFolder.id, expectedFolderName);
  }

  const folderWithParents = customerFolder_getDriveFileWithParents_(currentFolder.id);

  if ((folderWithParents.parents || []).indexOf(failedParentFolderId) !== -1) {
    customerFolder_writeFolderInfoToSheet_(sheet, rowNum, headerMap, currentFolder, expectedFolderName, 'ALREADY_IN_FAILED_FOLDER');
    customerFolder_writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
      status: 'ALREADY_IN_FAILED_FOLDER'
    });

    return {
      status: 'ALREADY_IN_FAILED_FOLDER',
      message: '이미 수주실패 폴더 안에 있음',
      customerNo,
      company,
      vendor,
      folderId: currentFolder.id,
      folderName: currentFolder.name
    };
  }

  customerFolder_moveDriveFileToFolder_(currentFolder.id, failedParentFolderId, folderWithParents.parents || []);

  customerFolder_writeFolderInfoToSheet_(sheet, rowNum, headerMap, currentFolder, expectedFolderName, 'MOVED_TO_FAILED_FOLDER');
  customerFolder_writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
    status: 'MOVED'
  });

  return {
    status: 'MOVED',
    message: `수주실패 폴더로 이동 완료 / 상태값: ${statusValue}`,
    customerNo,
    company,
    vendor,
    folderId: currentFolder.id,
    folderName: currentFolder.name
  };
}


function customerFolder_ensureFailedParentFolder_(driveId) {
  const name = FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME;

  const existing = customerFolder_findChildFolder_(driveId, driveId, name);

  if (existing) {
    return existing;
  }

  return customerFolder_createDriveFolder_(name, driveId);
}


function customerFolder_getCustomerFolderForFailedMove_(row, driveId, headerMap, customerNo) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const folderIdCol = customerFolder_col_(headerMap, cfg.OUTPUT_HEADERS.folderId);
  const existingFolderId = customerFolder_cleanValue_(row[folderIdCol - 1]);

  let folder = null;

  if (existingFolderId) {
    folder = customerFolder_getDriveFile_(existingFolderId);

    if (folder && folder.trashed) {
      folder = null;
    }

    if (folder && folder.mimeType !== 'application/vnd.google-apps.folder') {
      folder = null;
    }
  }

  if (!folder) {
    const found = customerFolder_findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, '');
    if (found && found.folder) {
      folder = found.folder;
    }
  }

  return folder;
}


function customerFolder_getDriveFileWithParents_(fileId) {
  const file = customerFolder_getDriveFile_(fileId);
  if (!file) {
    throw new Error(`Drive 폴더를 찾지 못했습니다: ${fileId}`);
  }
  return file;
}


function customerFolder_moveDriveFileToFolder_(fileId, targetParentId, currentParentIds) {
  const removeParents = (currentParentIds || [])
    .filter(parentId => parentId !== targetParentId)
    .join(',');

  let path =
    'files/' + encodeURIComponent(fileId) +
    '?supportsAllDrives=true' +
    '&addParents=' + encodeURIComponent(targetParentId) +
    '&fields=id,name,parents,webViewLink,mimeType,createdTime,size,md5Checksum';

  if (removeParents) {
    path += '&removeParents=' + encodeURIComponent(removeParents);
  }

  return customerFolder_driveFetch_(path, {
    method: 'patch',
    payload: {}
  });
}


function customerFolder_isFailedStatus_(statusValue) {
  const v = customerFolder_cleanValue_(statusValue);

  if (!v) return false;

  return FAILED_CUSTOMER_FOLDER_CFG.FAILED_STATUS_KEYWORDS.some(keyword => {
    const k = customerFolder_cleanValue_(keyword);
    return k && v.indexOf(k) !== -1;
  });
}


function customerFolder_ensureFailedFolderOutputHeaders_(sheet, headerMap) {
  const required = Object.values(FAILED_CUSTOMER_FOLDER_CFG.OUTPUT_HEADERS);

  required.forEach(headerName => {
    const key = customerFolder_normalizeHeader_(headerName);

    if (!headerMap[key]) {
      const newCol = sheet.getLastColumn() + 1;

      sheet
        .getRange(CUSTOMER_FOLDER_CFG.HEADER_ROW, newCol)
        .setValue(headerName)
        .setFontWeight('bold')
        .setBackground('#f4cccc');

      headerMap[key] = newCol;
    }
  });

  return headerMap;
}


function customerFolder_writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, info) {
  const nowText = Utilities.formatDate(new Date(), CUSTOMER_FOLDER_CFG.TZ, 'yyyy-MM-dd HH:mm:ss');

  sheet
    .getRange(rowNum, customerFolder_col_(headerMap, FAILED_CUSTOMER_FOLDER_CFG.OUTPUT_HEADERS.failedMoveStatus))
    .setValue(info.status || '');

  sheet
    .getRange(rowNum, customerFolder_col_(headerMap, FAILED_CUSTOMER_FOLDER_CFG.OUTPUT_HEADERS.failedMoveUpdatedAt))
    .setValue(nowText);
}


function customerFolder_findFirstExistingHeaderName_(headerMap, headerNames) {
  for (let i = 0; i < headerNames.length; i++) {
    const name = headerNames[i];

    if (headerMap[customerFolder_normalizeHeader_(name)]) {
      return name;
    }
  }

  return '';
}


function customerFolder_getCustomerDesiredParentInfoForRow_(rowData, headerMap, driveId, parentCache) {
  parentCache = parentCache || {};

  if (customerFolder_isFailedCustomerRowByData_(rowData, headerMap)) {
    if (!parentCache.failedParentFolderId) {
      const failedFolder = customerFolder_ensureFailedParentFolder_(driveId);
      parentCache.failedParentFolderId = failedFolder.id;
    }

    return {
      parentId: parentCache.failedParentFolderId,
      location: 'FAILED',
      moveStatus: 'MOVED_TO_FAILED_FOLDER'
    };
  }

  return {
    parentId: driveId,
    location: 'ROOT',
    moveStatus: 'MOVED_TO_ROOT'
  };
}


function customerFolder_ensureCustomerFolderParentForRow_(folder, rowData, headerMap, driveId, parentCache) {
  if (!folder || !folder.id) {
    return {
      folder,
      moved: false,
      parentId: '',
      location: ''
    };
  }

  const desired = customerFolder_getCustomerDesiredParentInfoForRow_(rowData, headerMap, driveId, parentCache || {});
  const folderWithParents = folder.parents && folder.parents.length
    ? folder
    : customerFolder_getDriveFileWithParents_(folder.id);
  const currentParents = folderWithParents.parents || [];
  const alreadyInTarget = currentParents.indexOf(desired.parentId) !== -1;
  const hasOnlyTarget = alreadyInTarget && currentParents.filter(parentId => parentId !== desired.parentId).length === 0;

  if (hasOnlyTarget) {
    return {
      folder: Object.assign({}, folder, folderWithParents),
      moved: false,
      parentId: desired.parentId,
      location: desired.location
    };
  }

  const movedFolder = customerFolder_moveDriveFileToFolder_(folder.id, desired.parentId, currentParents);

  return {
    folder: Object.assign({}, folder, movedFolder),
    moved: true,
    parentId: desired.parentId,
    location: desired.location,
    moveStatus: desired.moveStatus
  };
}


function customerFolder_appendCustomerFolderStatusPart_(status, part) {
  const s = customerFolder_cleanValue_(status);
  const p = customerFolder_cleanValue_(part);

  if (!p) return s;
  if (!s) return p;
  if (s.indexOf(p) !== -1) return s;

  return s + '_' + p;
}


function customerFolder_isFailedCustomerRowByData_(rowData, headerMap) {
  const statusHeaderName = customerFolder_findFirstExistingHeaderName_(
    headerMap,
    FAILED_CUSTOMER_FOLDER_CFG.STATUS_HEADER_CANDIDATES
  );

  if (!statusHeaderName) {
    return false;
  }

  const statusValue = customerFolder_cleanValue_(rowData[customerFolder_col_(headerMap, statusHeaderName) - 1]);

  return customerFolder_isFailedStatus_(statusValue);
}


/***** 스프레드시트/헤더 유틸 *****/

function customerFolder_getMasterSpreadsheet_() {
  const id = customerFolder_cleanValue_(CUSTOMER_FOLDER_CFG.MASTER_SPREADSHEET_ID);

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


function customerFolder_getMasterSheet_() {
  const ss = customerFolder_getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(CUSTOMER_FOLDER_CFG.MASTER_SHEET_NAME);

  if (!sheet) {
    throw new Error(`마스터시트를 찾지 못했습니다: ${CUSTOMER_FOLDER_CFG.MASTER_SHEET_NAME}`);
  }

  return sheet;
}


function customerFolder_getHeaderMap_(sheet) {
  const headerRow = CUSTOMER_FOLDER_CFG.HEADER_ROW;
  const lastCol = sheet.getLastColumn();

  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach((h, i) => {
    const key = customerFolder_normalizeHeader_(h);
    if (key && !map[key]) {
      map[key] = i + 1;
    }
  });

  return map;
}


function customerFolder_ensureOutputHeaders_(sheet, headerMap) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const required = Object.values(cfg.OUTPUT_HEADERS);

  required.forEach(headerName => {
    const key = customerFolder_normalizeHeader_(headerName);

    if (!headerMap[key]) {
      const newCol = sheet.getLastColumn() + 1;

      sheet
        .getRange(cfg.HEADER_ROW, newCol)
        .setValue(headerName)
        .setFontWeight('bold')
        .setBackground('#d9ead3');

      headerMap[key] = newCol;
    }
  });

  return headerMap;
}


function customerFolder_assertHeader_(headerMap, headerName) {
  if (!headerMap[customerFolder_normalizeHeader_(headerName)]) {
    throw new Error(`필수 헤더를 찾지 못했습니다: ${headerName}`);
  }
}


function customerFolder_col_(headerMap, headerName) {
  const key = customerFolder_normalizeHeader_(headerName);
  const c = headerMap[key];

  if (!c) {
    throw new Error(`헤더 컬럼을 찾지 못했습니다: ${headerName}`);
  }

  return c;
}


/***** 문자열/로그 유틸 *****/

function customerFolder_buildCustomerFolderName_(customerNo, company, vendor) {
  const customerNoPart = customerFolder_normalizeCustomerNoKey_(customerNo) || customerFolder_cleanValue_(customerNo);

  const parts = [
    customerFolder_sanitizeFolderPart_(customerNoPart),
    customerFolder_sanitizeFolderPart_(company),
    customerFolder_sanitizeFolderPart_(vendor)
  ];

  let name = parts.join('_').replace(/_+/g, '_').trim();

  if (name.length > 180) {
    name = name.slice(0, 180).trim();
  }

  return name;
}


function customerFolder_sanitizeFolderPart_(value) {
  return customerFolder_cleanValue_(value)
    .replace(/[\/\\:*?"<>|#\[\]\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function customerFolder_cleanValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}


function customerFolder_normalizeHeader_(value) {
  return customerFolder_cleanValue_(value).replace(/\s+/g, '');
}


function customerFolder_appendFolderLog_(rows) {
  if (!rows || rows.length === 0) return;

  const ss = customerFolder_getMasterSpreadsheet_();
  const name = CUSTOMER_FOLDER_CFG.LOG_SHEET_NAME;

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '일시',
      '행',
      '고객번호',
      '회사명',
      '수행사',
      '폴더명',
      '폴더ID',
      '처리결과',
      '메시지'
    ]);
    sheet.setFrozenRows(1);
  }

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


/***** Google Drive API 처리 함수 *****/

function customerFolder_getSharedDriveId_() {
  const cfg = CUSTOMER_FOLDER_CFG;

  if (customerFolder_cleanValue_(cfg.SHARED_DRIVE_ID)) {
    return customerFolder_cleanValue_(cfg.SHARED_DRIVE_ID);
  }

  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('S1_CUSTOMER_SHARED_DRIVE_ID');

  if (cached) {
    return cached;
  }

  const q = `name = ${customerFolder_driveQueryString_(cfg.SHARED_DRIVE_NAME)}`;

  const data = customerFolder_driveFetch_(
    'drives?pageSize=10&q=' + encodeURIComponent(q) + '&fields=drives(id,name)',
    { method: 'get' }
  );

  const drives = data.drives || [];

  if (drives.length === 0) {
    throw new Error(
      `공유 드라이브를 찾지 못했습니다: ${cfg.SHARED_DRIVE_NAME}. ` +
      `공유드라이브 ID를 CUSTOMER_FOLDER_CFG.SHARED_DRIVE_ID에 직접 입력하세요.`
    );
  }

  const driveId = drives[0].id;
  props.setProperty('S1_CUSTOMER_SHARED_DRIVE_ID', driveId);

  return driveId;
}


function customerFolder_findChildFolder_(parentId, driveId, folderName) {
  if (!folderName) return null;

  const q = [
    `${customerFolder_driveQueryString_(parentId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `name = ${customerFolder_driveQueryString_(folderName)}`,
    `trashed = false`
  ].join(' and ');

  const path =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=drive' +
    '&driveId=' + encodeURIComponent(driveId) +
    '&pageSize=10' +
    '&q=' + encodeURIComponent(q) +
    '&fields=files(id,name,webViewLink,trashed,mimeType,parents,createdTime,size,md5Checksum)';

  const data = customerFolder_driveFetch_(path, { method: 'get' });
  const files = data.files || [];

  return files.length ? files[0] : null;
}


function customerFolder_createDriveFolder_(folderName, parentId) {
  return customerFolder_driveFetch_(
    'files?supportsAllDrives=true&fields=id,name,webViewLink,trashed,mimeType,parents,createdTime',
    {
      method: 'post',
      payload: {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentId]
      }
    }
  );
}


function customerFolder_getDriveFile_(fileId) {
  try {
    return customerFolder_driveFetch_(
      'files/' + encodeURIComponent(fileId) +
      '?supportsAllDrives=true&fields=id,name,webViewLink,trashed,mimeType,parents,createdTime,size,md5Checksum',
      { method: 'get' }
    );
  } catch (err) {
    return null;
  }
}


function customerFolder_renameDriveFile_(fileId, newName) {
  return customerFolder_driveFetch_(
    'files/' + encodeURIComponent(fileId) +
    '?supportsAllDrives=true&fields=id,name,webViewLink,trashed,mimeType,parents,createdTime,size,md5Checksum',
    {
      method: 'patch',
      payload: {
        name: newName
      }
    }
  );
}


function customerFolder_driveFetch_(path, options) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const url = 'https://www.googleapis.com/drive/v3/' + path;
  const maxRetries = Number(cfg.DRIVE_API_MAX_RETRIES || 4);
  const baseDelay = Number(cfg.DRIVE_API_RETRY_BASE_MILLIS || 500);
  let lastError = null;

  for (let attempt = 0; attempt <= maxRetries; attempt++) {
    const params = Object.assign(
      {
        method: 'get',
        muteHttpExceptions: true,
        headers: {
          Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
        }
      },
      options || {}
    );

    if (params.payload && typeof params.payload !== 'string') {
      params.contentType = 'application/json';
      params.payload = JSON.stringify(params.payload);
    }

    try {
      const res = UrlFetchApp.fetch(url, params);
      const code = res.getResponseCode();
      const text = res.getContentText();

      if (code >= 200 && code < 300) {
        return text ? JSON.parse(text) : {};
      }

      const retryable =
        code === 429 ||
        code === 500 ||
        code === 502 ||
        code === 503 ||
        code === 504 ||
        (code === 403 && /rateLimitExceeded|userRateLimitExceeded|backendError/i.test(text));

      lastError = new Error(`Drive API 오류 ${code}: ${text}`);

      if (!retryable || attempt >= maxRetries) {
        lastError.customerFolderNoRetry = !retryable;
        throw lastError;
      }

      let retryAfterMs = 0;
      try {
        const headers = res.getHeaders ? res.getHeaders() : {};
        const retryAfter = Number(headers['Retry-After'] || headers['retry-after'] || 0);
        if (retryAfter > 0) retryAfterMs = retryAfter * 1000;
      } catch (headerErr) {
        retryAfterMs = 0;
      }

      const delay = retryAfterMs || Math.min(8000, baseDelay * Math.pow(2, attempt) + Math.floor(Math.random() * 250));
      Utilities.sleep(delay);

    } catch (err) {
      lastError = err;

      if (err && err.customerFolderNoRetry) {
        throw err;
      }

      if (attempt >= maxRetries) {
        throw err;
      }

      const delay = Math.min(8000, baseDelay * Math.pow(2, attempt) + Math.floor(Math.random() * 250));
      Utilities.sleep(delay);
    }
  }

  throw lastError || new Error('Drive API 호출 실패');
}


function customerFolder_driveQueryString_(value) {
  const s = customerFolder_cleanValue_(value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");

  return `'${s}'`;
}
