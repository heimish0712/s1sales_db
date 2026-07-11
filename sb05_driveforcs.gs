/****************************************************
 * KJ 선임신고용 서류 자동 분류 - 최적화 버전
 *
 * 설정 객체명:
 * - KJ_CS_CONFIG
 *
 * 주요 함수:
 * 1) classifyKjDocumentsNow()
 *    - 평소 자동 실행용
 *    - 최근 추가/수정 파일만 처리
 *
 * 2) classifyKjDocumentsFullScanNow()
 *    - 전체 재스캔
 *    - 계약번호 1번부터 전체 처리
 *
 * 3) classifyKjDocumentsFullScanFrom100Now()
 *    - 계약번호 100 이상만 전체 재스캔
 *
 * 4) installKjDocClassifierTrigger()
 *    - classifyKjDocumentsNow를 10분마다 자동 실행
 *
 * 필요:
 * - 고급 서비스 Drive API 활성화
 * - appsscript.json에는 Drive API 중복 없이 1개만 유지
 ****************************************************/

const KJ_CS_CONFIG = {
  SPREADSHEET_ID: '',
  SHEET_NAME: '수주확정/계약완료',

  COL_CONTRACT_NO: 1,     // A열 계약번호
  COL_CUSTOMER_NAME: 11,  // K열 고객사명
  COL_VENDOR: 19,         // S열 수행사

  TARGET_VENDOR: 'KJ',

    FOLDER_IDS: {
    CONTRACT_DOCS: '1enmgDejBzlzly5_k9DcBHOf3YK4dEkXg',
    BUSINESS_LICENSE: '1P4tZNLzqY46wWltPhtIxb7Mqhg2yREvK',
    BUILDING_REGISTER: '191gYagr7eXcJyF14wVq5V033P6G07s7B',
    TARGET_ROOT: '1M7w2-P5NVljYVV0HC0r7i6MMY4ZszGsy'
  },

  PREFIX_CATEGORY_TO_COPIED_FILE: true,
  COPY_UNCLASSIFIED_CONTRACT_FOLDER_FILES: true,

  LOG_SHEET_NAME: 'KJ서류분류로그',
  STATE_SHEET_NAME: 'KJ서류분류상태',

  TRIGGER_EVERY_MINUTES: 10,

  // 최근 파일 검색 시 안전 여유분
  LOOKBACK_MINUTES: 30,

  LAST_SCAN_PROP_KEY: 'KJ_CS_LAST_SCAN_ISO'
};


let KJ_CS_LOG_BUFFER = [];
let KJ_CS_STATE_BUFFER = [];


/**
 * 평소 자동 실행용
 * 최근 추가/수정된 파일만 처리
 */
function classifyKjDocumentsNow() {
  KJDOC_runClassifier_({
    fullScan: false
  });
}


/**
 * 수동 전체 재스캔용
 * 계약번호 1번부터 전체 처리
 */
function classifyKjDocumentsFullScanNow() {
  KJDOC_runClassifier_({
    fullScan: true
  });
}


/**
 * 수동 실행용
 * 계약번호 100 이상만 전체 재스캔
 */
function classifyKjDocumentsFullScanFrom100Now() {
  KJDOC_runFullScanFromContractNo_(1);
}


/**
 * 10분마다 자동 실행 트리거 설치
 */
function installKjDocClassifierTrigger() {
  KJDOC_deleteClassifierTriggers_();

  ScriptApp.newTrigger('classifyKjDocumentsNow')
    .timeBased()
    .everyMinutes(KJ_CS_CONFIG.TRIGGER_EVERY_MINUTES)
    .create();

  KJDOC_log_(
    'INFO',
    '',
    '',
    '',
    '',
    `${KJ_CS_CONFIG.TRIGGER_EVERY_MINUTES}분 주기 자동분류 트리거 설치 완료`
  );

  KJDOC_flushLogs_();
}


/**
 * 처리상태 초기화
 * 주의: 이걸 실행하면 기존 복사 상태 기록이 지워짐.
 * 이후 full scan을 돌리면 중복 복사될 수 있음.
 */
function resetKjDocClassifierState() {
  const ss = KJDOC_getSpreadsheet_();

  const stateSheet = ss.getSheetByName(KJ_CS_CONFIG.STATE_SHEET_NAME);
  if (stateSheet) {
    ss.deleteSheet(stateSheet);
  }

  PropertiesService.getScriptProperties()
    .deleteProperty(KJ_CS_CONFIG.LAST_SCAN_PROP_KEY);

  KJDOC_log_('INFO', '', '', '', '', 'KJ 서류분류 상태 초기화 완료');
  KJDOC_flushLogs_();
}


/**
 * 메인 실행부
 * fullScan=false: 최근 파일만
 * fullScan=true: 원본 폴더 전체 파일
 */
function KJDOC_runClassifier_(options) {
  const fullScan = !!options.fullScan;
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30 * 1000)) {
    console.log('이미 KJ 서류 분류 작업이 실행 중이라 이번 실행은 중단됨');
    return;
  }

  KJ_CS_LOG_BUFFER = [];
  KJ_CS_STATE_BUFFER = [];

  const runStartedAt = new Date();

  try {
    KJDOC_validateConfig_();

    const ss = KJDOC_getSpreadsheet_();
    const sheet = ss.getSheetByName(KJ_CS_CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`시트를 찾을 수 없음: ${KJ_CS_CONFIG.SHEET_NAME}`);
    }

    const contractMap = KJDOC_getKjContractMap_(sheet);
    const contractNos = Object.keys(contractMap);

    if (contractNos.length === 0) {
      KJDOC_log_('INFO', '', '', '', '', '처리할 KJ 계약 건 없음');
      KJDOC_flushLogs_();
      return;
    }

    const copiedState = KJDOC_loadCopiedState_();
    const targetFolderMap = KJDOC_buildTargetFolderMap_();

    const scanAfterIso = fullScan ? null : KJDOC_getScanAfterIso_();
    const foundMap = {};

    KJDOC_log_(
      'INFO',
      '',
      '',
      '',
      '',
      fullScan
        ? `전체 재스캔 시작 / 대상 KJ 계약 ${contractNos.length}건`
        : `최근 파일 스캔 시작 / 기준시각: ${scanAfterIso || '최초 실행이므로 전체'}`
    );

    const sourceGroups = KJDOC_getSourceGroups_();

    sourceGroups.forEach(group => {
      const files = KJDOC_listSourceFiles_(group.folderId, scanAfterIso);

      files.forEach(file => {
        if (file.mimeType === 'application/vnd.google-apps.folder') return;

        const contractNo = KJDOC_extractLeadingContractNo_(file.title);
        if (!contractNo) return;

        const contractInfo = contractMap[contractNo];
        if (!contractInfo) return;

        const categoryName = KJDOC_getCategoryName_(group.sourceKey, file.title);
        if (!categoryName) return;

        KJDOC_markFound_(foundMap, contractNo, categoryName);

        KJDOC_copyOneFileIfNeeded_({
          file,
          categoryName,
          contractNo,
          customerName: contractInfo.customerName,
          targetFolderMap,
          copiedState
        });
      });
    });

    if (fullScan) {
      KJDOC_logMissingCategories_(contractMap, foundMap);
    }

    KJDOC_flushState_();
    KJDOC_flushLogs_();

    if (!fullScan) {
      PropertiesService.getScriptProperties()
        .setProperty(
          KJ_CS_CONFIG.LAST_SCAN_PROP_KEY,
          KJDOC_toDriveDateString_(runStartedAt)
        );
    }

  } catch (err) {
    KJDOC_log_('ERROR', '', '', '', '', err.stack || err.message);
    KJDOC_flushState_();
    KJDOC_flushLogs_();
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/**
 * 계약번호 minContractNo 이상만 전체 재스캔
 * 예: 100 이상
 */
function KJDOC_runFullScanFromContractNo_(minContractNo) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30 * 1000)) {
    console.log('이미 KJ 서류 분류 작업이 실행 중이라 이번 실행은 중단됨');
    return;
  }

  KJ_CS_LOG_BUFFER = [];
  KJ_CS_STATE_BUFFER = [];

  try {
    KJDOC_validateConfig_();

    const ss = KJDOC_getSpreadsheet_();
    const sheet = ss.getSheetByName(KJ_CS_CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`시트를 찾을 수 없음: ${KJ_CS_CONFIG.SHEET_NAME}`);
    }

    const allContractMap = KJDOC_getKjContractMap_(sheet);
    const contractMap = {};

    Object.keys(allContractMap).forEach(contractNo => {
      const num = Number(contractNo);
      if (!Number.isFinite(num)) return;
      if (num < minContractNo) return;

      contractMap[contractNo] = allContractMap[contractNo];
    });

    const contractNos = Object.keys(contractMap)
      .sort((a, b) => Number(a) - Number(b));

    if (contractNos.length === 0) {
      KJDOC_log_(
        'INFO',
        '',
        '',
        '',
        '',
        `계약번호 ${minContractNo} 이상 KJ 계약 건 없음`
      );

      KJDOC_flushLogs_();
      return;
    }

    const copiedState = KJDOC_loadCopiedState_();
    const targetFolderMap = KJDOC_buildTargetFolderMap_();
    const sourceGroups = KJDOC_getSourceGroups_();
    const foundMap = {};

    KJDOC_log_(
      'INFO',
      '',
      '',
      '',
      '',
      `계약번호 ${minContractNo} 이상 전체 재스캔 시작 / 대상 ${contractNos.length}건`
    );

    contractNos.forEach(contractNo => {
      const contractInfo = contractMap[contractNo];

      sourceGroups.forEach(group => {
        const files = KJDOC_listSourceFilesByContractNo_(
          group.folderId,
          contractNo
        );

        files.forEach(file => {
          if (file.mimeType === 'application/vnd.google-apps.folder') return;

          const exactContractNo = KJDOC_extractLeadingContractNo_(file.title);
          if (exactContractNo !== contractNo) return;

          const categoryName = KJDOC_getCategoryName_(
            group.sourceKey,
            file.title
          );

          if (!categoryName) return;

          KJDOC_markFound_(foundMap, contractNo, categoryName);

          KJDOC_copyOneFileIfNeeded_({
            file,
            categoryName,
            contractNo,
            customerName: contractInfo.customerName,
            targetFolderMap,
            copiedState
          });
        });
      });
    });

    KJDOC_logMissingCategories_(contractMap, foundMap);

    KJDOC_flushState_();
    KJDOC_flushLogs_();

  } catch (err) {
    KJDOC_log_('ERROR', '', '', '', '', err.stack || err.message);
    KJDOC_flushState_();
    KJDOC_flushLogs_();
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/**
 * 설정 검증
 */
function KJDOC_validateConfig_() {
  const ids = KJ_CS_CONFIG.FOLDER_IDS;

  const required = [
    ['CONTRACT_DOCS', ids.CONTRACT_DOCS],
    ['BUSINESS_LICENSE', ids.BUSINESS_LICENSE],
    ['BUILDING_REGISTER', ids.BUILDING_REGISTER],
    ['TARGET_ROOT', ids.TARGET_ROOT]
  ];

  required.forEach(([key, value]) => {
    if (!value || String(value).includes('여기에_')) {
      throw new Error(`KJ_CS_CONFIG.FOLDER_IDS.${key} 폴더 ID를 먼저 입력해야 함`);
    }
  });
}


/**
 * 스프레드시트 가져오기
 */
function KJDOC_getSpreadsheet_() {
  if (KJ_CS_CONFIG.SPREADSHEET_ID && KJ_CS_CONFIG.SPREADSHEET_ID.trim()) {
    return SpreadsheetApp.openById(KJ_CS_CONFIG.SPREADSHEET_ID.trim());
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


/**
 * S열이 KJ인 계약번호 → 고객사명 맵 생성
 */
function KJDOC_getKjContractMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = Math.max(
    KJ_CS_CONFIG.COL_CONTRACT_NO,
    KJ_CS_CONFIG.COL_CUSTOMER_NAME,
    KJ_CS_CONFIG.COL_VENDOR
  );

  const map = {};

  if (lastRow < 2) return map;

  const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getDisplayValues();

  values.forEach(row => {
    const contractNo = KJDOC_normalizeContractNo_(
      row[KJ_CS_CONFIG.COL_CONTRACT_NO - 1]
    );

    const customerName = KJDOC_normalizeText_(
      row[KJ_CS_CONFIG.COL_CUSTOMER_NAME - 1]
    );

    const vendor = KJDOC_normalizeText_(
      row[KJ_CS_CONFIG.COL_VENDOR - 1]
    );

    if (!contractNo || !customerName) return;
    if (vendor !== KJ_CS_CONFIG.TARGET_VENDOR) return;

    map[contractNo] = {
      contractNo,
      customerName
    };
  });

  return map;
}


/**
 * 원본 폴더 그룹
 */
function KJDOC_getSourceGroups_() {
  return [
    {
      sourceKey: 'CONTRACT_DOCS',
      folderId: KJ_CS_CONFIG.FOLDER_IDS.CONTRACT_DOCS
    },
    {
      sourceKey: 'BUSINESS_LICENSE',
      folderId: KJ_CS_CONFIG.FOLDER_IDS.BUSINESS_LICENSE
    },
    {
      sourceKey: 'BUILDING_REGISTER',
      folderId: KJ_CS_CONFIG.FOLDER_IDS.BUILDING_REGISTER
    }
  ];
}


/**
 * 마지막 실행 기준 시각 가져오기
 */
function KJDOC_getScanAfterIso_() {
  const props = PropertiesService.getScriptProperties();
  const lastScanIso = props.getProperty(KJ_CS_CONFIG.LAST_SCAN_PROP_KEY);

  if (!lastScanIso) {
    return null;
  }

  const lastScanDate = new Date(lastScanIso);
  const lookbackDate = new Date(
    lastScanDate.getTime() - KJ_CS_CONFIG.LOOKBACK_MINUTES * 60 * 1000
  );

  return KJDOC_toDriveDateString_(lookbackDate);
}


/**
 * 원본 폴더 파일 목록
 * scanAfterIso가 있으면 최근 파일만 검색
 */
function KJDOC_listSourceFiles_(folderId, scanAfterIso) {
  const baseConditions = [
    `'${folderId}' in parents`,
    `trashed = false`
  ];

  if (scanAfterIso) {
    baseConditions.push(
      `(createdDate > '${scanAfterIso}' or modifiedDate > '${scanAfterIso}')`
    );
  }

  const q = baseConditions.join(' and ');

  return KJDOC_listFilesByQuery_(q);
}


/**
 * 특정 계약번호가 파일명에 포함된 파일만 원본 폴더에서 조회
 *
 * Drive 검색은 정규식이 안 되므로 title contains로 후보를 좁히고,
 * 파일명 앞 숫자가 정확히 같은지 다시 필터링함.
 */
function KJDOC_listSourceFilesByContractNo_(folderId, contractNo) {
  const escapedContractNo = KJDOC_escapeDriveQueryText_(contractNo);

  const q = [
    `'${folderId}' in parents`,
    `trashed = false`,
    `title contains '${escapedContractNo}'`
  ].join(' and ');

  const candidates = KJDOC_listFilesByQuery_(q);

  return candidates.filter(file => {
    return KJDOC_extractLeadingContractNo_(file.title) === String(contractNo);
  });
}


/**
 * Drive API 파일 검색
 */
function KJDOC_listFilesByQuery_(q) {
  const all = [];
  let pageToken = null;

  do {
    const result = Drive.Files.list({
      q,
      maxResults: 1000,
      pageToken,
      fields: 'items(id,title,mimeType,createdDate,modifiedDate),nextPageToken',
      supportsAllDrives: true,
      includeItemsFromAllDrives: true
    });

    if (result.items && result.items.length > 0) {
      all.push(...result.items);
    }

    pageToken = result.nextPageToken;
  } while (pageToken);

  return all;
}


/**
 * 대상 루트 폴더의 하위 폴더 목록을 한 번에 맵으로 가져오기
 */
function KJDOC_buildTargetFolderMap_() {
  const targetRootId = KJ_CS_CONFIG.FOLDER_IDS.TARGET_ROOT;

  const q = [
    `'${targetRootId}' in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `trashed = false`
  ].join(' and ');

  const folders = KJDOC_listFilesByQuery_(q);
  const map = {};

  folders.forEach(folder => {
    map[folder.title] = {
      id: folder.id,
      title: folder.title
    };
  });

  return map;
}


/**
 * 대상 폴더 가져오기 또는 생성
 */
function KJDOC_getOrCreateTargetFolder_(contractNo, customerName, targetFolderMap) {
  const folderName = KJDOC_buildTargetFolderName_(contractNo, customerName);

  if (targetFolderMap[folderName]) {
    return targetFolderMap[folderName];
  }

  const created = Drive.Files.insert(
    {
      title: folderName,
      mimeType: 'application/vnd.google-apps.folder',
      parents: [{ id: KJ_CS_CONFIG.FOLDER_IDS.TARGET_ROOT }]
    },
    null,
    {
      supportsAllDrives: true
    }
  );

  targetFolderMap[folderName] = {
    id: created.id,
    title: created.title
  };

  KJDOC_log_(
    'CREATE_FOLDER',
    contractNo,
    customerName,
    '',
    folderName,
    '대상 하위 폴더 생성'
  );

  return targetFolderMap[folderName];
}


/**
 * 파일 분류명 판단
 */
function KJDOC_getCategoryName_(sourceKey, fileName) {
  if (sourceKey === 'BUSINESS_LICENSE') {
    return '사업자등록증';
  }

  if (sourceKey === 'BUILDING_REGISTER') {
    return '건축물대장';
  }

  if (sourceKey === 'CONTRACT_DOCS') {
    const isAppointment = KJDOC_containsAny_(fileName, [
      '선임',
      '선임신고',
      '선임신고서',
      '위임',
      '위임장',
      '신고서'
    ]);

    if (isAppointment) {
      return '선임신고서_위임장';
    }

    const isContract = KJDOC_containsAny_(fileName, [
      '계약',
      '계약서'
    ]);

    if (isContract) {
      return '계약서';
    }

    if (KJ_CS_CONFIG.COPY_UNCLASSIFIED_CONTRACT_FOLDER_FILES) {
      return '계약서폴더_미분류';
    }

    return '';
  }

  return '';
}


/**
 * 파일 1개 복사
 */
function KJDOC_copyOneFileIfNeeded_(params) {
  const {
    file,
    categoryName,
    contractNo,
    customerName,
    targetFolderMap,
    copiedState
  } = params;

  const targetFolder = KJDOC_getOrCreateTargetFolder_(
    contractNo,
    customerName,
    targetFolderMap
  );

  const copiedKey = KJDOC_buildCopiedKey_(file.id, targetFolder.id);

  if (copiedState[copiedKey]) {
    return;
  }

  const targetFileName = KJDOC_buildCopiedFileName_(categoryName, file.title);

  Drive.Files.copy(
    {
      title: targetFileName,
      parents: [{ id: targetFolder.id }]
    },
    file.id,
    {
      supportsAllDrives: true
    }
  );

  copiedState[copiedKey] = true;

  KJ_CS_STATE_BUFFER.push([
    new Date(),
    copiedKey,
    file.id,
    file.title,
    file.modifiedDate || '',
    targetFolder.id,
    targetFolder.title,
    categoryName,
    targetFileName,
    contractNo,
    customerName
  ]);

  KJDOC_log_(
    categoryName === '계약서폴더_미분류' ? 'WARN' : 'COPY',
    contractNo,
    customerName,
    categoryName,
    targetFileName,
    categoryName === '계약서폴더_미분류'
      ? '계약서 폴더 내 미분류 파일 복사'
      : '복사 완료'
  );
}


/**
 * 복사 상태시트 읽기
 */
function KJDOC_loadCopiedState_() {
  const ss = KJDOC_getSpreadsheet_();
  let sheet = ss.getSheetByName(KJ_CS_CONFIG.STATE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(KJ_CS_CONFIG.STATE_SHEET_NAME);
    sheet.appendRow([
      '일시',
      '복사키',
      '원본파일ID',
      '원본파일명',
      '원본수정일',
      '대상폴더ID',
      '대상폴더명',
      '분류',
      '복사파일명',
      '계약번호',
      '고객사명'
    ]);
  }

  const lastRow = sheet.getLastRow();
  const state = {};

  if (lastRow < 2) return state;

  const values = sheet.getRange(2, 2, lastRow - 1, 1).getValues();

  values.forEach(row => {
    const key = String(row[0] || '').trim();
    if (key) state[key] = true;
  });

  return state;
}


/**
 * 복사 상태 기록 한번에 저장
 */
function KJDOC_flushState_() {
  if (!KJ_CS_STATE_BUFFER || KJ_CS_STATE_BUFFER.length === 0) return;

  const ss = KJDOC_getSpreadsheet_();
  let sheet = ss.getSheetByName(KJ_CS_CONFIG.STATE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(KJ_CS_CONFIG.STATE_SHEET_NAME);
    sheet.appendRow([
      '일시',
      '복사키',
      '원본파일ID',
      '원본파일명',
      '원본수정일',
      '대상폴더ID',
      '대상폴더명',
      '분류',
      '복사파일명',
      '계약번호',
      '고객사명'
    ]);
  }

  const startRow = sheet.getLastRow() + 1;

  sheet
    .getRange(startRow, 1, KJ_CS_STATE_BUFFER.length, KJ_CS_STATE_BUFFER[0].length)
    .setValues(KJ_CS_STATE_BUFFER);

  KJ_CS_STATE_BUFFER = [];
}


/**
 * 로그 버퍼에 추가
 */
function KJDOC_log_(status, contractNo, customerName, category, fileName, message) {
  KJ_CS_LOG_BUFFER.push([
    new Date(),
    status,
    contractNo,
    customerName,
    category,
    fileName,
    message
  ]);
}


/**
 * 로그 한번에 저장
 */
function KJDOC_flushLogs_() {
  if (!KJ_CS_LOG_BUFFER || KJ_CS_LOG_BUFFER.length === 0) return;

  const ss = KJDOC_getSpreadsheet_();
  let sheet = ss.getSheetByName(KJ_CS_CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(KJ_CS_CONFIG.LOG_SHEET_NAME);
    sheet.appendRow([
      '일시',
      '상태',
      '계약번호',
      '고객사명',
      '분류',
      '파일명',
      '메시지'
    ]);
  }

  const startRow = sheet.getLastRow() + 1;

  sheet
    .getRange(startRow, 1, KJ_CS_LOG_BUFFER.length, KJ_CS_LOG_BUFFER[0].length)
    .setValues(KJ_CS_LOG_BUFFER);

  KJ_CS_LOG_BUFFER = [];
}


/**
 * 기존 자동분류 트리거 삭제
 */
function KJDOC_deleteClassifierTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'classifyKjDocumentsNow') {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


/**
 * 찾은 파일 표시
 */
function KJDOC_markFound_(foundMap, contractNo, categoryName) {
  if (!foundMap[contractNo]) {
    foundMap[contractNo] = {};
  }

  foundMap[contractNo][categoryName] = true;
}


/**
 * 전체 재스캔 시 누락 파일 로그 기록
 */
function KJDOC_logMissingCategories_(contractMap, foundMap) {
  const requiredCategories = [
    '계약서',
    '선임신고서_위임장',
    '사업자등록증',
    '건축물대장'
  ];

  Object.keys(contractMap).forEach(contractNo => {
    const contractInfo = contractMap[contractNo];
    const found = foundMap[contractNo] || {};

    requiredCategories.forEach(categoryName => {
      if (found[categoryName]) return;

      KJDOC_log_(
        'MISSING',
        contractNo,
        contractInfo.customerName,
        categoryName,
        '',
        `${categoryName} 파일을 찾지 못함`
      );
    });
  });
}


/**
 * 복사 키 생성
 */
function KJDOC_buildCopiedKey_(sourceFileId, targetFolderId) {
  return `${sourceFileId}__${targetFolderId}`;
}


/**
 * 복사 파일명 만들기
 */
function KJDOC_buildCopiedFileName_(categoryName, originalFileName) {
  if (!KJ_CS_CONFIG.PREFIX_CATEGORY_TO_COPIED_FILE) {
    return KJDOC_sanitizeDriveName_(originalFileName);
  }

  return KJDOC_sanitizeDriveName_(`[${categoryName}] ${originalFileName}`);
}


/**
 * 대상 폴더명 만들기
 */
function KJDOC_buildTargetFolderName_(contractNo, customerName) {
  return KJDOC_sanitizeDriveName_(`${contractNo}_${customerName}`);
}


/**
 * 파일명 앞 숫자 추출
 */
function KJDOC_extractLeadingContractNo_(fileName) {
  const match = String(fileName || '').trim().match(/^0*(\d+)/);
  if (!match) return '';

  return String(Number(match[1]));
}


/**
 * 시트 계약번호 정규화
 */
function KJDOC_normalizeContractNo_(value) {
  const text = String(value || '').trim();
  if (!text) return '';

  const match = text.match(/\d+/);
  if (!match) return '';

  return String(Number(match[0]));
}


/**
 * 텍스트 정규화
 */
function KJDOC_normalizeText_(value) {
  return String(value || '').trim();
}


/**
 * 문자열 포함 검사
 */
function KJDOC_containsAny_(text, keywords) {
  const lower = String(text || '').toLowerCase();

  return keywords.some(keyword => {
    return lower.includes(String(keyword).toLowerCase());
  });
}


/**
 * Drive 파일/폴더명 안전 처리
 */
function KJDOC_sanitizeDriveName_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|#\[\]]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 180);
}


/**
 * Drive 검색어 escape
 */
function KJDOC_escapeDriveQueryText_(value) {
  return String(value || '')
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");
}


/**
 * Drive 검색 날짜 포맷
 */
function KJDOC_toDriveDateString_(date) {
  return date.toISOString().replace(/\.\d{3}Z$/, 'Z');
}

/**
 * 강제 실행용:
 * 계약번호 100 이상 전체 재스캔
 *
 * 주의:
 * - 기존 실행이 아직 돌고 있어도 무시하고 실행함
 * - 동시에 같은 파일을 처리하면 중복 복사 가능성 있음
 * - 테스트용/긴급용으로만 쓰는 게 맞음
 */
function classifyKjDocumentsFullScanFrom100ForceNow() {
  KJDOC_runFullScanFromContractNoForce_(165);
}


/**
 * 강제 실행 본체:
 * LockService를 아예 사용하지 않음
 */
function KJDOC_runFullScanFromContractNoForce_(minContractNo) {
  KJ_CS_LOG_BUFFER = [];
  KJ_CS_STATE_BUFFER = [];

  try {
    KJDOC_validateConfig_();

    const ss = KJDOC_getSpreadsheet_();
    const sheet = ss.getSheetByName(KJ_CS_CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`시트를 찾을 수 없음: ${KJ_CS_CONFIG.SHEET_NAME}`);
    }

    const allContractMap = KJDOC_getKjContractMap_(sheet);
    const contractMap = {};

    Object.keys(allContractMap).forEach(contractNo => {
      const num = Number(contractNo);
      if (!Number.isFinite(num)) return;
      if (num < minContractNo) return;

      contractMap[contractNo] = allContractMap[contractNo];
    });

    const contractNos = Object.keys(contractMap)
      .sort((a, b) => Number(a) - Number(b));

    if (contractNos.length === 0) {
      KJDOC_log_(
        'INFO',
        '',
        '',
        '',
        '',
        `강제실행: 계약번호 ${minContractNo} 이상 KJ 계약 건 없음`
      );

      KJDOC_flushLogs_();
      return;
    }

    const copiedState = KJDOC_loadCopiedState_();
    const targetFolderMap = KJDOC_buildTargetFolderMap_();
    const sourceGroups = KJDOC_getSourceGroups_();
    const foundMap = {};

    KJDOC_log_(
      'FORCE_START',
      '',
      '',
      '',
      '',
      `강제실행: 계약번호 ${minContractNo} 이상 전체 재스캔 시작 / 대상 ${contractNos.length}건`
    );

    contractNos.forEach(contractNo => {
      const contractInfo = contractMap[contractNo];

      sourceGroups.forEach(group => {
        const files = KJDOC_listSourceFilesByContractNo_(
          group.folderId,
          contractNo
        );

        files.forEach(file => {
          if (file.mimeType === 'application/vnd.google-apps.folder') return;

          const exactContractNo = KJDOC_extractLeadingContractNo_(file.title);
          if (exactContractNo !== contractNo) return;

          const categoryName = KJDOC_getCategoryName_(
            group.sourceKey,
            file.title
          );

          if (!categoryName) return;

          KJDOC_markFound_(foundMap, contractNo, categoryName);

          KJDOC_copyOneFileIfNeeded_({
            file,
            categoryName,
            contractNo,
            customerName: contractInfo.customerName,
            targetFolderMap,
            copiedState
          });
        });
      });
    });

    KJDOC_logMissingCategories_(contractMap, foundMap);

    KJDOC_log_(
      'FORCE_DONE',
      '',
      '',
      '',
      '',
      `강제실행 완료: 계약번호 ${minContractNo} 이상 전체 재스캔`
    );

    KJDOC_flushState_();
    KJDOC_flushLogs_();

  } catch (err) {
    KJDOC_log_('ERROR', '', '', '', '', err.stack || err.message);
    KJDOC_flushState_();
    KJDOC_flushLogs_();
    throw err;
  }
}

/****************************************************
 * 문서 종류별 FULL SCAN 실행 함수
 *
 * 아래 4개 함수 중 필요한 것만 골라 실행하면 됨.
 ****************************************************/


/**
 * 계약서만 전체 재스캔
 */
function classifyKjContractsFullScanNow() {
  KJDOC_runFullScanOnlyCategory_('계약서');
}


/**
 * 선임신고서 및 위임장만 전체 재스캔
 */
function classifyKjAppointmentDocsFullScanNow() {
  KJDOC_runFullScanOnlyCategory_('선임신고서_위임장');
}


/**
 * 사업자등록증만 전체 재스캔
 */
function classifyKjBusinessLicensesFullScanNow() {
  KJDOC_runFullScanOnlyCategory_('사업자등록증');
}


/**
 * 건축물대장만 전체 재스캔
 */
function classifyKjBuildingRegistersFullScanNow() {
  KJDOC_runFullScanOnlyCategory_('건축물대장');
}


/**
 * 특정 문서 종류만 전체 재스캔
 *
 * targetCategoryName:
 * - 계약서
 * - 선임신고서_위임장
 * - 사업자등록증
 * - 건축물대장
 */
function KJDOC_runFullScanOnlyCategory_(targetCategoryName) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(30 * 1000)) {
    console.log('이미 KJ 서류 분류 작업이 실행 중이라 이번 실행은 중단됨');
    return;
  }

  KJ_CS_LOG_BUFFER = [];
  KJ_CS_STATE_BUFFER = [];

  try {
    KJDOC_validateConfig_();

    const ss = KJDOC_getSpreadsheet_();
    const sheet = ss.getSheetByName(KJ_CS_CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`시트를 찾을 수 없음: ${KJ_CS_CONFIG.SHEET_NAME}`);
    }

    const contractMap = KJDOC_getKjContractMap_(sheet);
    const contractNos = Object.keys(contractMap);

    if (contractNos.length === 0) {
      KJDOC_log_(
        'INFO',
        '',
        '',
        targetCategoryName,
        '',
        '처리할 KJ 계약 건 없음'
      );

      KJDOC_flushLogs_();
      return;
    }

    const copiedState = KJDOC_loadCopiedState_();
    const targetFolderMap = KJDOC_buildTargetFolderMap_();
    const sourceGroups = KJDOC_getSourceGroupsForCategory_(targetCategoryName);
    const foundMap = {};

    KJDOC_log_(
      'INFO',
      '',
      '',
      targetCategoryName,
      '',
      `${targetCategoryName}만 전체 재스캔 시작 / 대상 KJ 계약 ${contractNos.length}건`
    );

    sourceGroups.forEach(group => {
      const files = KJDOC_listSourceFiles_(group.folderId, null);

      files.forEach(file => {
        if (file.mimeType === 'application/vnd.google-apps.folder') return;

        const contractNo = KJDOC_extractLeadingContractNo_(file.title);
        if (!contractNo) return;

        const contractInfo = contractMap[contractNo];
        if (!contractInfo) return;

        const categoryName = KJDOC_getCategoryName_(group.sourceKey, file.title);

        if (categoryName !== targetCategoryName) return;

        KJDOC_markFound_(foundMap, contractNo, categoryName);

        KJDOC_copyOneFileIfNeeded_({
          file,
          categoryName,
          contractNo,
          customerName: contractInfo.customerName,
          targetFolderMap,
          copiedState
        });
      });
    });

    KJDOC_logMissingOneCategory_(
      contractMap,
      foundMap,
      targetCategoryName
    );

    KJDOC_log_(
      'DONE',
      '',
      '',
      targetCategoryName,
      '',
      `${targetCategoryName}만 전체 재스캔 완료`
    );

    KJDOC_flushState_();
    KJDOC_flushLogs_();

  } catch (err) {
    KJDOC_log_('ERROR', '', '', targetCategoryName, '', err.stack || err.message);
    KJDOC_flushState_();
    KJDOC_flushLogs_();
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/**
 * 문서 종류별로 읽을 원본 폴더만 선택
 */
function KJDOC_getSourceGroupsForCategory_(targetCategoryName) {
  if (
    targetCategoryName === '계약서' ||
    targetCategoryName === '선임신고서_위임장'
  ) {
    return [
      {
        sourceKey: 'CONTRACT_DOCS',
        folderId: KJ_CS_CONFIG.FOLDER_IDS.CONTRACT_DOCS
      }
    ];
  }

  if (targetCategoryName === '사업자등록증') {
    return [
      {
        sourceKey: 'BUSINESS_LICENSE',
        folderId: KJ_CS_CONFIG.FOLDER_IDS.BUSINESS_LICENSE
      }
    ];
  }

  if (targetCategoryName === '건축물대장') {
    return [
      {
        sourceKey: 'BUILDING_REGISTER',
        folderId: KJ_CS_CONFIG.FOLDER_IDS.BUILDING_REGISTER
      }
    ];
  }

  throw new Error(`지원하지 않는 문서 종류임: ${targetCategoryName}`);
}


/**
 * 특정 문서 종류 1개에 대해서만 누락 로그 기록
 */
function KJDOC_logMissingOneCategory_(contractMap, foundMap, targetCategoryName) {
  Object.keys(contractMap).forEach(contractNo => {
    const contractInfo = contractMap[contractNo];
    const found = foundMap[contractNo] || {};

    if (found[targetCategoryName]) return;

    KJDOC_log_(
      'MISSING',
      contractNo,
      contractInfo.customerName,
      targetCategoryName,
      '',
      `${targetCategoryName} 파일을 찾지 못함`
    );
  });
}

/****************************************************
 * 문서 종류별 FULL SCAN 강제 실행 함수
 * LockService 무시 버전
 *
 * 주의:
 * - 기존 실행이 살아 있어도 실행됨
 * - 동시에 같은 파일을 처리하면 중복 복사 가능성 있음
 * - 테스트/긴급용으로만 사용
 ****************************************************/


/**
 * 계약서만 강제 전체 재스캔
 */
function forceKjContractsFullScanNow() {
  KJDOC_runFullScanOnlyCategoryForce_('계약서');
}


/**
 * 선임신고서 및 위임장만 강제 전체 재스캔
 */
function forceKjAppointmentDocsFullScanNow() {
  KJDOC_runFullScanOnlyCategoryForce_('선임신고서_위임장');
}


/**
 * 사업자등록증만 강제 전체 재스캔
 */
function forceKjBusinessLicensesFullScanNow() {
  KJDOC_runFullScanOnlyCategoryForce_('사업자등록증');
}


/**
 * 건축물대장만 강제 전체 재스캔
 */
function forceKjBuildingRegistersFullScanNow() {
  KJDOC_runFullScanOnlyCategoryForce_('건축물대장');
}


/**
 * 특정 문서 종류만 강제 전체 재스캔
 * LockService 사용 안 함
 */
function KJDOC_runFullScanOnlyCategoryForce_(targetCategoryName) {
  KJ_CS_LOG_BUFFER = [];
  KJ_CS_STATE_BUFFER = [];

  try {
    KJDOC_validateConfig_();

    const ss = KJDOC_getSpreadsheet_();
    const sheet = ss.getSheetByName(KJ_CS_CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`시트를 찾을 수 없음: ${KJ_CS_CONFIG.SHEET_NAME}`);
    }

    const contractMap = KJDOC_getKjContractMap_(sheet);
    const contractNos = Object.keys(contractMap);

    if (contractNos.length === 0) {
      KJDOC_log_(
        'INFO',
        '',
        '',
        targetCategoryName,
        '',
        '강제실행: 처리할 KJ 계약 건 없음'
      );

      KJDOC_flushLogs_();
      return;
    }

    const copiedState = KJDOC_loadCopiedState_();
    const targetFolderMap = KJDOC_buildTargetFolderMap_();
    const sourceGroups = KJDOC_getSourceGroupsForCategory_(targetCategoryName);
    const foundMap = {};

    KJDOC_log_(
      'FORCE_START',
      '',
      '',
      targetCategoryName,
      '',
      `강제실행: ${targetCategoryName}만 전체 재스캔 시작 / 대상 KJ 계약 ${contractNos.length}건`
    );

    sourceGroups.forEach(group => {
      const files = KJDOC_listSourceFiles_(group.folderId, null);

      files.forEach(file => {
        if (file.mimeType === 'application/vnd.google-apps.folder') return;

        const contractNo = KJDOC_extractLeadingContractNo_(file.title);
        if (!contractNo) return;

        const contractInfo = contractMap[contractNo];
        if (!contractInfo) return;

        const categoryName = KJDOC_getCategoryName_(group.sourceKey, file.title);

        if (categoryName !== targetCategoryName) return;

        KJDOC_markFound_(foundMap, contractNo, categoryName);

        KJDOC_copyOneFileIfNeeded_({
          file,
          categoryName,
          contractNo,
          customerName: contractInfo.customerName,
          targetFolderMap,
          copiedState
        });
      });
    });

    KJDOC_logMissingOneCategory_(
      contractMap,
      foundMap,
      targetCategoryName
    );

    KJDOC_log_(
      'FORCE_DONE',
      '',
      '',
      targetCategoryName,
      '',
      `강제실행 완료: ${targetCategoryName}만 전체 재스캔`
    );

    KJDOC_flushState_();
    KJDOC_flushLogs_();

  } catch (err) {
    KJDOC_log_('ERROR', '', '', targetCategoryName, '', err.stack || err.message);
    KJDOC_flushState_();
    KJDOC_flushLogs_();
    throw err;
  }
}