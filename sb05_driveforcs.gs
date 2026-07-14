/****************************************************
 * KJ 선임신고용 서류 자동 분류 - Drive API v2
 *
 * 기준
 * - 연결된 스프레드시트의 '고객관리' 시트
 * - '계약번호', '고객사명' 헤더를 이름으로 탐색
 * - 고객관리의 모든 유효 행을 대상으로
 *   계약번호_고객사명 폴더를 생성/보정
 *
 * 드라이브
 * - 공유 드라이브명: S1 KJ 공유
 * - 원본 폴더명: 계약서 / 사업자등록증 / 건축물대장
 * - 대상 루트: 박치산 대리님 선임신고용 폴더
 * - 고정 폴더 ID가 아니라 공유 드라이브명과 폴더명으로 탐색
 *
 * 중복 방지
 * - 원본 파일 ID를 중복 기준으로 사용하지 않음
 * - 대상 폴더의 실제 파일을 매번 확인
 * - 동일 MD5 또는 동일 정규화 파일명이 있으면 복사하지 않음
 * - 예전 코드가 붙인 분류 접두어도 제거한 뒤 비교
 *
 * 거의 실시간
 * - 1분 주기 시간 기반 트리거
 * - 직전 실행 시각보다 10분 앞에서 다시 조회하여 누락 방지
 *
 * 필수
 * - Apps Script 서비스 > Drive API(고급 서비스) 활성화
 * - appsscript.json의 Drive 고급 서비스 버전은 반드시 v2
 ****************************************************/

const KJ_DOC_CONFIG = Object.freeze({
  SPREADSHEET_ID: '',
  SHEET_NAME: '고객관리',

  HEADER_CONTRACT_NO: '계약번호',
  HEADER_CUSTOMER_NAME: '고객사명',
  HEADER_SCAN_ROWS: 10,

  SHARED_DRIVE_NAME: 'S1 KJ 공유',
  TARGET_ROOT_FOLDER_NAME: '박치산 대리님 선임신고용 폴더',

  SOURCE_FOLDERS: Object.freeze([
    Object.freeze({
      key: 'CONTRACT_DOCS',
      folderName: '계약서',
      required: true
    }),
    Object.freeze({
      key: 'BUSINESS_LICENSE',
      folderName: '사업자등록증',
      required: false
    }),
    Object.freeze({
      key: 'BUILDING_REGISTER',
      folderName: '건축물대장',
      required: false
    })
  ]),

  TRIGGER_EVERY_MINUTES: 1,
  SAFETY_FULL_SCAN_EVERY_HOURS: 6,
  LOOKBACK_MINUTES: 10,
  LAST_SCAN_PROP_KEY: 'KJ_DOC_V2_LAST_SCAN_ISO',

  LOG_SHEET_NAME: 'KJ서류분류로그',
  STATE_SHEET_NAME: 'KJ서류분류상태',

  FOLDER_MIME_TYPE: 'application/vnd.google-apps.folder',
  MAX_PAGE_SIZE: 1000
});

let KJ_DOC_LOG_BUFFER = [];
let KJ_DOC_STATE_BUFFER = [];


/****************************************************
 * 공개 실행 함수
 ****************************************************/

/**
 * 최초 설정 권장 함수
 * 1) 고객관리 기준 폴더 생성/이름 보정
 * 2) 1분 주기 트리거 설치
 *
 * 최초 기존 자료 전체 복사는 별도로
 * classifyKjDocumentsFullScanNow()를 1회 실행하십시오.
 */
function setupKjDocClassifier() {
  syncKjCustomerFoldersNow();
  installKjDocClassifierTrigger();
}


/**
 * 수정/복사 없이 연결 상태만 확인합니다.
 * 실행 로그에 공유 드라이브, 대상 폴더, 원본 폴더, 고객 수를 출력합니다.
 */
function checkKjDocClassifierV2() {
  KJDOCV2_validateRuntime_();

  const ss = KJDOCV2_getSpreadsheet_();
  const sheet = ss.getSheetByName(KJ_DOC_CONFIG.SHEET_NAME);
  if (!sheet) {
    throw new Error(`시트를 찾을 수 없습니다: ${KJ_DOC_CONFIG.SHEET_NAME}`);
  }

  const customerMap = KJDOCV2_readCustomerMap_(sheet);
  const sharedDrive = KJDOCV2_findSharedDriveByName_(
    KJ_DOC_CONFIG.SHARED_DRIVE_NAME
  );
  const targetRoot = KJDOCV2_findFolderByNameInDrive_(
    sharedDrive.id,
    KJ_DOC_CONFIG.TARGET_ROOT_FOLDER_NAME,
    true
  );

  const sourceFolders = KJ_DOC_CONFIG.SOURCE_FOLDERS.map(config => {
    const folder = KJDOCV2_findFolderByNameInDrive_(
      sharedDrive.id,
      config.folderName,
      false
    );

    return {
      key: config.key,
      folderName: config.folderName,
      found: !!folder,
      folderId: folder ? folder.id : ''
    };
  });

  const result = {
    apiVersion: 'v2',
    customerCount: Object.keys(customerMap).length,
    sharedDriveName: KJ_DOC_CONFIG.SHARED_DRIVE_NAME,
    sharedDriveId: sharedDrive.id,
    targetRootName: targetRoot.title,
    targetRootId: targetRoot.id,
    sourceFolders
  };

  console.log(JSON.stringify(result, null, 2));
  return result;
}

/** 평소 자동 실행 함수: 최근 추가/수정 파일만 처리 */
function classifyKjDocumentsNow() {
  KJDOCV2_run_({
    fullScan: false,
    foldersOnly: false,
    minContractNo: 1,
    sourceKeys: null,
    categoryFilter: null
  });
}

/** 전체 원본 재스캔 */
function classifyKjDocumentsFullScanNow() {
  KJDOCV2_run_({
    fullScan: true,
    foldersOnly: false,
    minContractNo: 1,
    sourceKeys: null,
    categoryFilter: null
  });
}

/** 계약번호 100 이상만 전체 재스캔 */
function classifyKjDocumentsFullScanFrom100Now() {
  KJDOCV2_run_({
    fullScan: true,
    foldersOnly: false,
    minContractNo: 100,
    sourceKeys: null,
    categoryFilter: null
  });
}

/** 고객관리 전체 기준 고객사별 폴더만 생성/보정 */
function syncKjCustomerFoldersNow() {
  KJDOCV2_run_({
    fullScan: false,
    foldersOnly: true,
    minContractNo: 1,
    sourceKeys: null,
    categoryFilter: null
  });
}

/** 계약서 원본 폴더만 전체 재스캔 */
function classifyKjContractsFullScanNow() {
  KJDOCV2_run_({
    fullScan: true,
    foldersOnly: false,
    minContractNo: 1,
    sourceKeys: ['CONTRACT_DOCS'],
    categoryFilter: null
  });
}

/** 선임신고서/위임장만 전체 재스캔 */
function classifyKjAppointmentDocsFullScanNow() {
  KJDOCV2_run_({
    fullScan: true,
    foldersOnly: false,
    minContractNo: 1,
    sourceKeys: ['CONTRACT_DOCS'],
    categoryFilter: '선임신고서_위임장'
  });
}

/** 사업자등록증만 전체 재스캔 */
function classifyKjBusinessLicensesFullScanNow() {
  KJDOCV2_run_({
    fullScan: true,
    foldersOnly: false,
    minContractNo: 1,
    sourceKeys: ['BUSINESS_LICENSE'],
    categoryFilter: null
  });
}

/** 건축물대장만 전체 재스캔 */
function classifyKjBuildingRegistersFullScanNow() {
  KJDOCV2_run_({
    fullScan: true,
    foldersOnly: false,
    minContractNo: 1,
    sourceKeys: ['BUILDING_REGISTER'],
    categoryFilter: null
  });
}

/**
 * 기존 코드의 강제실행 함수명 호환용.
 * Drive API v2판에서는 중복 방지를 위해 LockService를 무시하지 않습니다.
 */
function classifyKjDocumentsFullScanFrom100ForceNow() {
  classifyKjDocumentsFullScanFrom100Now();
}

function forceKjContractsFullScanNow() {
  classifyKjContractsFullScanNow();
}

function forceKjAppointmentDocsFullScanNow() {
  classifyKjAppointmentDocsFullScanNow();
}

function forceKjBusinessLicensesFullScanNow() {
  classifyKjBusinessLicensesFullScanNow();
}

function forceKjBuildingRegistersFullScanNow() {
  classifyKjBuildingRegistersFullScanNow();
}

/**
 * 1분 주기 최근파일 트리거 + 6시간 주기 안전 전체스캔 트리거 설치.
 * 기존 동일 핸들러 트리거는 먼저 제거합니다.
 */
function installKjDocClassifierTrigger() {
  KJDOCV2_deleteClassifierTriggers_();

  ScriptApp.newTrigger('classifyKjDocumentsNow')
    .timeBased()
    .everyMinutes(KJ_DOC_CONFIG.TRIGGER_EVERY_MINUTES)
    .create();

  ScriptApp.newTrigger('classifyKjDocumentsSafetyFullScan')
    .timeBased()
    .everyHours(KJ_DOC_CONFIG.SAFETY_FULL_SCAN_EVERY_HOURS)
    .create();

  KJ_DOC_LOG_BUFFER = [];
  KJDOCV2_log_(
    'TRIGGER_INSTALL',
    '',
    '',
    '',
    '',
    `${KJ_DOC_CONFIG.TRIGGER_EVERY_MINUTES}분 주기 최근파일 처리 + ` +
      `${KJ_DOC_CONFIG.SAFETY_FULL_SCAN_EVERY_HOURS}시간 주기 안전 전체스캔 트리거 설치 완료`
  );
  KJDOCV2_flushLogs_();
}

/** 자동 실행 누락 방지용 안전 전체스캔 */
function classifyKjDocumentsSafetyFullScan() {
  classifyKjDocumentsFullScanNow();
}

/** 자동분류 트리거 제거 */
function uninstallKjDocClassifierTrigger() {
  KJDOCV2_deleteClassifierTriggers_();

  KJ_DOC_LOG_BUFFER = [];
  KJDOCV2_log_('TRIGGER_DELETE', '', '', '', '', '자동분류 트리거 제거 완료');
  KJDOCV2_flushLogs_();
}

/**
 * 최근조회 시각만 초기화합니다.
 * 다음 자동 실행은 전체 원본을 다시 조회하지만,
 * 대상 폴더 실파일 기준 중복검사를 하므로 같은 파일을 재복사하지 않습니다.
 */
function resetKjDocClassifierState() {
  PropertiesService.getScriptProperties()
    .deleteProperty(KJ_DOC_CONFIG.LAST_SCAN_PROP_KEY);

  KJ_DOC_LOG_BUFFER = [];
  KJDOCV2_log_(
    'STATE_RESET',
    '',
    '',
    '',
    '',
    '최근조회 시각 초기화 완료. 다음 실행은 전체 조회하며 대상 실파일 기준으로 중복 방지함'
  );
  KJDOCV2_flushLogs_();
}


/****************************************************
 * 메인 처리
 ****************************************************/

function KJDOCV2_run_(options) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(25 * 1000)) {
    console.log('이미 KJ 서류 자동분류가 실행 중이므로 이번 실행은 종료합니다.');
    return;
  }

  KJ_DOC_LOG_BUFFER = [];
  KJ_DOC_STATE_BUFFER = [];

  const runStartedAt = new Date();
  const stats = {
    customers: 0,
    foldersCreated: 0,
    foldersRenamed: 0,
    sourceFiles: 0,
    copied: 0,
    duplicateName: 0,
    duplicateHash: 0,
    unmatched: 0,
    filtered: 0,
    missingOptionalSources: 0
  };

  try {
    KJDOCV2_validateRuntime_();

    const ss = KJDOCV2_getSpreadsheet_();
    const sheet = ss.getSheetByName(KJ_DOC_CONFIG.SHEET_NAME);

    if (!sheet) {
      throw new Error(`시트를 찾을 수 없습니다: ${KJ_DOC_CONFIG.SHEET_NAME}`);
    }

    const customerMap = KJDOCV2_readCustomerMap_(sheet);
    const contractNos = Object.keys(customerMap);
    stats.customers = contractNos.length;

    if (contractNos.length === 0) {
      throw new Error(
        `'${KJ_DOC_CONFIG.SHEET_NAME}' 시트에서 계약번호와 고객사명이 모두 있는 행을 찾지 못했습니다.`
      );
    }

    const sharedDrive = KJDOCV2_findSharedDriveByName_(
      KJ_DOC_CONFIG.SHARED_DRIVE_NAME
    );

    const targetRoot = KJDOCV2_findFolderByNameInDrive_(
      sharedDrive.id,
      KJ_DOC_CONFIG.TARGET_ROOT_FOLDER_NAME,
      true
    );

    const targetFolderMap = KJDOCV2_ensureCustomerFolders_({
      sharedDriveId: sharedDrive.id,
      targetRoot,
      customerMap,
      stats
    });

    if (options.foldersOnly) {
      KJDOCV2_log_(
        'DONE',
        '',
        '',
        '',
        '',
        `고객폴더 동기화 완료 / 고객 ${stats.customers}건 / 생성 ${stats.foldersCreated}건 / 이름보정 ${stats.foldersRenamed}건`
      );
      KJDOCV2_flushState_();
      KJDOCV2_flushLogs_();
      return;
    }

    const sourceGroups = KJDOCV2_resolveSourceGroups_(
      sharedDrive.id,
      options.sourceKeys,
      stats
    );

    const scanAfterIso = options.fullScan
      ? null
      : KJDOCV2_getScanAfterIso_();

    KJDOCV2_log_(
      'START',
      '',
      '',
      '',
      '',
      options.fullScan
        ? `전체 재스캔 시작 / 고객 ${stats.customers}건 / 최소 계약번호 ${options.minContractNo || 1}`
        : `최근 파일 스캔 시작 / 조회기준 ${scanAfterIso || '최초 실행이므로 전체'}`
    );

    const targetContentCache = {};
    const runSeenSourceIds = new Set();

    sourceGroups.forEach(group => {
      const sourceFiles = KJDOCV2_listDirectFiles_(
        sharedDrive.id,
        group.folder.id,
        scanAfterIso
      );

      sourceFiles.forEach(file => {
        stats.sourceFiles += 1;

        if (!file || !file.id || runSeenSourceIds.has(file.id)) return;
        runSeenSourceIds.add(file.id);

        const contractNo = KJDOCV2_extractLeadingContractNo_(file.title);
        if (!contractNo) {
          stats.unmatched += 1;
          return;
        }

        if (Number(contractNo) < Number(options.minContractNo || 1)) {
          stats.filtered += 1;
          return;
        }

        const customerInfo = customerMap[contractNo];
        if (!customerInfo) {
          stats.unmatched += 1;
          KJDOCV2_log_(
            'UNMATCHED',
            contractNo,
            '',
            KJDOCV2_getCategoryName_(group.key, file.title),
            file.title,
            '파일명 앞 계약번호가 고객관리 시트에 없음'
          );
          return;
        }

        const categoryName = KJDOCV2_getCategoryName_(group.key, file.title);
        if (options.categoryFilter && categoryName !== options.categoryFilter) {
          stats.filtered += 1;
          return;
        }

        const targetFolder = targetFolderMap[contractNo];
        if (!targetFolder) {
          throw new Error(
            `계약번호 ${contractNo}의 대상 폴더를 준비하지 못했습니다: ${customerInfo.customerName}`
          );
        }

        KJDOCV2_copyIfNeeded_({
          sharedDriveId: sharedDrive.id,
          sourceGroup: group,
          sourceFile: file,
          categoryName,
          contractNo,
          customerInfo,
          targetFolder,
          targetContentCache,
          stats
        });
      });
    });

    if (!options.fullScan) {
      PropertiesService.getScriptProperties().setProperty(
        KJ_DOC_CONFIG.LAST_SCAN_PROP_KEY,
        runStartedAt.toISOString()
      );
    }

    KJDOCV2_log_(
      'DONE',
      '',
      '',
      '',
      '',
      [
        `처리 완료`,
        `고객 ${stats.customers}건`,
        `폴더생성 ${stats.foldersCreated}건`,
        `폴더이름보정 ${stats.foldersRenamed}건`,
        `원본조회 ${stats.sourceFiles}개`,
        `신규복사 ${stats.copied}개`,
        `동일파일명건너뜀 ${stats.duplicateName}개`,
        `동일내용건너뜀 ${stats.duplicateHash}개`,
        `고객미매칭 ${stats.unmatched}개`
      ].join(' / ')
    );

    KJDOCV2_flushState_();
    KJDOCV2_flushLogs_();
  } catch (err) {
    KJDOCV2_log_(
      'ERROR',
      '',
      '',
      '',
      '',
      err && (err.stack || err.message) ? (err.stack || err.message) : String(err)
    );

    KJDOCV2_flushState_();
    KJDOCV2_flushLogs_();
    throw err;
  } finally {
    lock.releaseLock();
  }
}


/****************************************************
 * 고객관리 읽기 및 고객 폴더 동기화
 ****************************************************/

function KJDOCV2_readCustomerMap_(sheet) {
  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();

  if (lastRow < 1 || lastCol < 1) return {};

  const scanRows = Math.min(KJ_DOC_CONFIG.HEADER_SCAN_ROWS, lastRow);
  const headerArea = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();

  let headerRowIndex = -1;
  let contractColIndex = -1;
  let customerColIndex = -1;

  for (let r = 0; r < headerArea.length; r += 1) {
    const normalizedHeaders = headerArea[r].map(KJDOCV2_normalizeHeader_);
    const cIndex = normalizedHeaders.indexOf(
      KJDOCV2_normalizeHeader_(KJ_DOC_CONFIG.HEADER_CONTRACT_NO)
    );
    const nIndex = normalizedHeaders.indexOf(
      KJDOCV2_normalizeHeader_(KJ_DOC_CONFIG.HEADER_CUSTOMER_NAME)
    );

    if (cIndex >= 0 && nIndex >= 0) {
      headerRowIndex = r;
      contractColIndex = cIndex;
      customerColIndex = nIndex;
      break;
    }
  }

  if (headerRowIndex < 0) {
    throw new Error(
      `상단 ${scanRows}행 안에서 '${KJ_DOC_CONFIG.HEADER_CONTRACT_NO}', '${KJ_DOC_CONFIG.HEADER_CUSTOMER_NAME}' 헤더를 찾지 못했습니다.`
    );
  }

  const firstDataRow = headerRowIndex + 2;
  if (firstDataRow > lastRow) return {};

  const values = sheet
    .getRange(firstDataRow, 1, lastRow - firstDataRow + 1, lastCol)
    .getDisplayValues();

  const map = {};

  values.forEach((row, offset) => {
    const contractNo = KJDOCV2_normalizeContractNo_(row[contractColIndex]);
    const customerName = KJDOCV2_normalizeCustomerName_(row[customerColIndex]);

    if (!contractNo || !customerName) return;

    if (map[contractNo] && map[contractNo].customerName !== customerName) {
      throw new Error(
        `고객관리 시트에 같은 계약번호가 서로 다른 고객사명으로 중복되어 있습니다. ` +
        `계약번호 ${contractNo}: '${map[contractNo].customerName}' / '${customerName}' ` +
        `(행 ${firstDataRow + offset})`
      );
    }

    map[contractNo] = {
      contractNo,
      customerName,
      rowNo: firstDataRow + offset,
      expectedFolderName: KJDOCV2_buildTargetFolderName_(contractNo, customerName)
    };
  });

  return map;
}

function KJDOCV2_ensureCustomerFolders_(params) {
  const {
    sharedDriveId,
    targetRoot,
    customerMap,
    stats
  } = params;

  const existingFolders = KJDOCV2_listDirectFolders_(
    sharedDriveId,
    targetRoot.id
  );

  const byExactName = {};
  const byContractNo = {};

  existingFolders.forEach(folder => {
    byExactName[folder.title] = folder;

    const contractNo = KJDOCV2_extractLeadingContractNo_(folder.title);
    if (!contractNo) return;

    if (!byContractNo[contractNo]) byContractNo[contractNo] = [];
    byContractNo[contractNo].push(folder);
  });

  const result = {};
  const contractNos = Object.keys(customerMap)
    .sort((a, b) => Number(a) - Number(b));

  contractNos.forEach(contractNo => {
    const info = customerMap[contractNo];
    const expectedName = info.expectedFolderName;

    if (byExactName[expectedName]) {
      result[contractNo] = byExactName[expectedName];
      return;
    }

    const sameNoFolders = byContractNo[contractNo] || [];

    if (sameNoFolders.length === 1) {
      const oldFolder = sameNoFolders[0];
      const renamed = KJDOCV2_renameFile_(oldFolder.id, expectedName);

      const updatedFolder = {
        id: oldFolder.id,
        title: renamed.title || expectedName,
        mimeType: KJ_DOC_CONFIG.FOLDER_MIME_TYPE,
        parents: oldFolder.parents || [{ id: targetRoot.id }],
        driveId: sharedDriveId
      };

      result[contractNo] = updatedFolder;
      byExactName[expectedName] = updatedFolder;
      byContractNo[contractNo] = [updatedFolder];
      stats.foldersRenamed += 1;

      KJDOCV2_log_(
        'RENAME_FOLDER',
        contractNo,
        info.customerName,
        '',
        expectedName,
        `고객관리 최신 고객사명 기준으로 폴더명 보정: ${oldFolder.title} → ${expectedName}`
      );
      return;
    }

    if (sameNoFolders.length > 1) {
      const selected = sameNoFolders
        .slice()
        .sort((a, b) => String(a.title).localeCompare(String(b.title)))[0];

      result[contractNo] = selected;

      KJDOCV2_log_(
        'DUPLICATE_FOLDER',
        contractNo,
        info.customerName,
        '',
        selected.title,
        `같은 계약번호로 대상 폴더가 ${sameNoFolders.length}개 있어 자동 이름변경은 하지 않고 '${selected.title}' 폴더를 사용함`
      );
      return;
    }

    const created = KJDOCV2_createFolder_(targetRoot.id, expectedName);
    result[contractNo] = created;
    byExactName[expectedName] = created;
    byContractNo[contractNo] = [created];
    stats.foldersCreated += 1;

    KJDOCV2_log_(
      'CREATE_FOLDER',
      contractNo,
      info.customerName,
      '',
      expectedName,
      '고객관리 기준 대상 고객폴더 생성'
    );
  });

  return result;
}


/****************************************************
 * 공유 드라이브 및 폴더 탐색
 ****************************************************/

function KJDOCV2_findSharedDriveByName_(driveName) {
  const matches = [];
  let pageToken = null;

  do {
    const listOptions = {
      maxResults: 100,
      fields: 'nextPageToken,items(id,name)'
    };
    if (pageToken) listOptions.pageToken = pageToken;

    // Drive API v2의 drives.list 응답 배열 필드는 drives가 아니라 items입니다.
    const response = Drive.Drives.list(listOptions);

    const sharedDrives = response.items || [];
    sharedDrives.forEach(sharedDrive => {
      if (String(sharedDrive.name || '').trim() === String(driveName).trim()) {
        matches.push(sharedDrive);
      }
    });

    pageToken = response.nextPageToken || null;
  } while (pageToken);

  if (matches.length === 0) {
    throw new Error(
      `접근 가능한 공유 드라이브에서 '${driveName}'을 찾지 못했습니다. ` +
      `스크립트 실행 계정의 공유 드라이브 권한을 확인하십시오.`
    );
  }

  if (matches.length > 1) {
    throw new Error(
      `이름이 '${driveName}'인 공유 드라이브가 ${matches.length}개입니다. ` +
      `동일 이름 공유 드라이브를 정리해야 안전하게 실행할 수 있습니다.`
    );
  }

  return matches[0];
}

function KJDOCV2_findFolderByNameInDrive_(driveId, folderName, required) {
  const q = [
    `title = '${KJDOCV2_escapeDriveQueryText_(folderName)}'`,
    `mimeType = '${KJ_DOC_CONFIG.FOLDER_MIME_TYPE}'`,
    'trashed = false'
  ].join(' and ');

  const candidates = KJDOCV2_listDriveFiles_(driveId, q);
  const exact = candidates.filter(file => file.title === folderName);

  if (exact.length === 0) {
    if (!required) return null;
    throw new Error(
      `공유 드라이브 '${KJ_DOC_CONFIG.SHARED_DRIVE_NAME}'에서 폴더 '${folderName}'을 찾지 못했습니다.`
    );
  }

  if (exact.length === 1) return exact[0];

  const rootLevel = exact.filter(file => {
    return Array.isArray(file.parents) && file.parents.some(parent => parent && parent.id === driveId);
  });

  if (rootLevel.length === 1) return rootLevel[0];

  throw new Error(
    `공유 드라이브 안에 이름이 '${folderName}'인 폴더가 ${exact.length}개 있어 대상을 확정할 수 없습니다.`
  );
}

function KJDOCV2_resolveSourceGroups_(driveId, sourceKeys, stats) {
  const allowedKeys = sourceKeys && sourceKeys.length
    ? new Set(sourceKeys)
    : null;

  const groups = [];

  KJ_DOC_CONFIG.SOURCE_FOLDERS.forEach(config => {
    if (allowedKeys && !allowedKeys.has(config.key)) return;

    const folder = KJDOCV2_findFolderByNameInDrive_(
      driveId,
      config.folderName,
      config.required
    );

    if (!folder) {
      stats.missingOptionalSources += 1;
      KJDOCV2_log_(
        'SOURCE_NOT_FOUND',
        '',
        '',
        '',
        config.folderName,
        '선택 원본 폴더를 찾지 못해 이번 실행에서 제외함'
      );
      return;
    }

    groups.push({
      key: config.key,
      folderName: config.folderName,
      folder
    });
  });

  if (groups.length === 0) {
    throw new Error('처리 가능한 원본 폴더를 하나도 찾지 못했습니다.');
  }

  return groups;
}


/****************************************************
 * 원본 조회 및 중복 방지 복사
 ****************************************************/

function KJDOCV2_copyIfNeeded_(params) {
  const {
    sharedDriveId,
    sourceGroup,
    sourceFile,
    categoryName,
    contractNo,
    customerInfo,
    targetFolder,
    targetContentCache,
    stats
  } = params;

  if (!targetContentCache[targetFolder.id]) {
    targetContentCache[targetFolder.id] = KJDOCV2_loadTargetContentIndex_(
      sharedDriveId,
      targetFolder.id,
      contractNo
    );
  }

  const index = targetContentCache[targetFolder.id];
  const signature = KJDOCV2_buildDuplicateSignature_(
    contractNo,
    sourceFile.title
  );
  const md5 = String(sourceFile.md5Checksum || '').trim().toLowerCase();

  if (md5 && index.hashes.has(md5)) {
    stats.duplicateHash += 1;
    return;
  }

  if (index.signatures.has(signature)) {
    stats.duplicateName += 1;
    return;
  }

  const targetFileName = KJDOCV2_sanitizeDriveName_(sourceFile.title);

  const copied = Drive.Files.copy(
    {
      title: targetFileName,
      parents: [{ id: targetFolder.id }]
    },
    sourceFile.id,
    {
      supportsAllDrives: true,
      fields: 'id,title,md5Checksum,modifiedDate'
    }
  );

  index.signatures.add(signature);
  if (md5) index.hashes.add(md5);

  stats.copied += 1;

  const copiedKey = signature;
  KJ_DOC_STATE_BUFFER.push([
    new Date(),
    copiedKey,
    sourceFile.id,
    sourceFile.title,
    sourceFile.modifiedDate || '',
    targetFolder.id,
    targetFolder.title,
    categoryName,
    copied.title || targetFileName,
    contractNo,
    customerInfo.customerName
  ]);

  KJDOCV2_log_(
    'COPY',
    contractNo,
    customerInfo.customerName,
    categoryName,
    copied.title || targetFileName,
    `${sourceGroup.folderName} → ${targetFolder.title} 복사 완료`
  );
}

function KJDOCV2_loadTargetContentIndex_(driveId, targetFolderId, contractNo) {
  const q = [
    `'${targetFolderId}' in parents`,
    'trashed = false',
    `mimeType != '${KJ_DOC_CONFIG.FOLDER_MIME_TYPE}'`
  ].join(' and ');

  const files = KJDOCV2_listDriveFiles_(driveId, q);
  const signatures = new Set();
  const hashes = new Set();

  files.forEach(file => {
    signatures.add(
      KJDOCV2_buildDuplicateSignature_(contractNo, file.title)
    );

    const md5 = String(file.md5Checksum || '').trim().toLowerCase();
    if (md5) hashes.add(md5);
  });

  return { signatures, hashes };
}

function KJDOCV2_listDirectFiles_(driveId, parentFolderId, scanAfterIso) {
  const conditions = [
    `'${parentFolderId}' in parents`,
    'trashed = false',
    `mimeType != '${KJ_DOC_CONFIG.FOLDER_MIME_TYPE}'`
  ];

  if (scanAfterIso) {
    conditions.push(`modifiedDate > '${scanAfterIso}'`);
  }

  return KJDOCV2_listDriveFiles_(driveId, conditions.join(' and '));
}

function KJDOCV2_listDirectFolders_(driveId, parentFolderId) {
  const q = [
    `'${parentFolderId}' in parents`,
    'trashed = false',
    `mimeType = '${KJ_DOC_CONFIG.FOLDER_MIME_TYPE}'`
  ].join(' and ');

  return KJDOCV2_listDriveFiles_(driveId, q);
}

function KJDOCV2_listDriveFiles_(driveId, q) {
  const all = [];
  let pageToken = null;

  do {
    const listOptions = {
      q,
      corpora: 'drive',
      driveId,
      includeItemsFromAllDrives: true,
      supportsAllDrives: true,
      maxResults: KJ_DOC_CONFIG.MAX_PAGE_SIZE,
      fields: [
        'nextPageToken',
        'items(id,title,mimeType,parents(id,isRoot),driveId,createdDate,modifiedDate,fileSize,md5Checksum)'
      ].join(',')
    };
    if (pageToken) listOptions.pageToken = pageToken;

    const response = Drive.Files.list(listOptions);

    if (response.items && response.items.length) {
      all.push.apply(all, response.items);
    }

    pageToken = response.nextPageToken || null;
  } while (pageToken);

  return all;
}

function KJDOCV2_createFolder_(parentId, folderName) {
  return Drive.Files.insert(
    {
      title: folderName,
      mimeType: KJ_DOC_CONFIG.FOLDER_MIME_TYPE,
      parents: [{ id: parentId }]
    },
    null,
    {
      supportsAllDrives: true,
      fields: 'id,title,mimeType,parents(id,isRoot),driveId'
    }
  );
}

function KJDOCV2_renameFile_(fileId, newTitle) {
  return Drive.Files.patch(
    { title: newTitle },
    fileId,
    {
      supportsAllDrives: true,
      fields: 'id,title,mimeType,parents(id,isRoot),driveId'
    }
  );
}


/****************************************************
 * 분류 및 중복 서명
 ****************************************************/

function KJDOCV2_getCategoryName_(sourceKey, fileName) {
  if (sourceKey === 'BUSINESS_LICENSE') return '사업자등록증';
  if (sourceKey === 'BUILDING_REGISTER') return '건축물대장';

  const normalized = KJDOCV2_normalizeLooseText_(fileName);

  if (
    normalized.indexOf('선임') >= 0 ||
    normalized.indexOf('위임') >= 0
  ) {
    return '선임신고서_위임장';
  }

  if (
    normalized.indexOf('용역신청') >= 0 ||
    normalized.indexOf('신청서') >= 0 ||
    normalized.indexOf('발주서') >= 0
  ) {
    return '용역신청서';
  }

  if (normalized.indexOf('계약') >= 0) {
    return '계약서';
  }

  return '기타계약서류';
}

function KJDOCV2_buildDuplicateSignature_(contractNo, fileName) {
  return `${contractNo}__${KJDOCV2_canonicalFileName_(fileName)}`;
}

function KJDOCV2_canonicalFileName_(fileName) {
  let text = KJDOCV2_unicodeNormalize_(fileName)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .trim();

  // 예전 코드의 [분류] 접두어 및 대괄호 제거 후 남은 접두어 호환
  text = text.replace(/^\s*\[[^\]]+\]\s*/u, '');

  const oldPrefixes = [
    '계약서폴더_미분류',
    '선임신고서_위임장',
    '사업자등록증',
    '건축물대장',
    '용역신청서',
    '기타계약서류',
    '계약서'
  ];

  for (let i = 0; i < oldPrefixes.length; i += 1) {
    const prefix = oldPrefixes[i];
    if (text === prefix) {
      text = '';
      break;
    }

    if (text.indexOf(prefix + ' ') === 0) {
      text = text.slice(prefix.length).trim();
      break;
    }
  }

  const extensionMatch = text.match(/(\.[^.\s]+)$/u);
  const extension = extensionMatch ? extensionMatch[1].toLowerCase() : '';
  let base = extension ? text.slice(0, -extension.length) : text;

  // 윈도우/드라이브가 자동으로 붙인 복사본 번호는 동일 제목으로 간주
  base = base
    .replace(/\s*\((?:복사본\s*)?\d+\)\s*$/u, '')
    .replace(/\s+-\s+복사본\s*$/u, '')
    .replace(/\s+/g, ' ')
    .trim()
    .toLowerCase();

  return `${base}${extension}`;
}


/****************************************************
 * 시각/정규화/보조 함수
 ****************************************************/

function KJDOCV2_getScanAfterIso_() {
  const lastScanIso = PropertiesService.getScriptProperties()
    .getProperty(KJ_DOC_CONFIG.LAST_SCAN_PROP_KEY);

  if (!lastScanIso) return null;

  const lastScan = new Date(lastScanIso);
  if (Number.isNaN(lastScan.getTime())) return null;

  const lookback = new Date(
    lastScan.getTime() - KJ_DOC_CONFIG.LOOKBACK_MINUTES * 60 * 1000
  );

  return lookback.toISOString();
}

function KJDOCV2_getSpreadsheet_() {
  const id = String(KJ_DOC_CONFIG.SPREADSHEET_ID || '').trim();
  return id
    ? SpreadsheetApp.openById(id)
    : SpreadsheetApp.getActiveSpreadsheet();
}

function KJDOCV2_validateRuntime_() {
  if (typeof Drive === 'undefined' || !Drive.Files || !Drive.Drives) {
    throw new Error(
      'Apps Script 고급 서비스의 Drive API가 활성화되어 있지 않습니다. ' +
      '편집기 왼쪽 서비스(+)에서 Drive API를 추가하십시오.'
    );
  }

  if (
    typeof Drive.Files.insert !== 'function' ||
    typeof Drive.Files.patch !== 'function' ||
    typeof Drive.Drives.list !== 'function'
  ) {
    throw new Error(
      '현재 Drive 고급 서비스가 v2가 아닙니다. appsscript.json에서 ' +
      'serviceId=drive 항목의 version을 v2로 설정하십시오.'
    );
  }

  if (KJ_DOC_CONFIG.TRIGGER_EVERY_MINUTES !== 1) {
    throw new Error('Drive API v2판 기본 자동 실행 주기는 1분이어야 합니다.');
  }
}

function KJDOCV2_extractLeadingContractNo_(value) {
  const text = KJDOCV2_unicodeNormalize_(value).trim();
  const match = text.match(/^0*(\d+)/u);
  if (!match) return '';

  const number = Number(match[1]);
  return Number.isFinite(number) ? String(number) : '';
}

function KJDOCV2_normalizeContractNo_(value) {
  const text = KJDOCV2_unicodeNormalize_(value).trim();
  if (!text) return '';

  const match = text.match(/\d+/u);
  if (!match) return '';

  const number = Number(match[0]);
  return Number.isFinite(number) ? String(number) : '';
}

function KJDOCV2_normalizeCustomerName_(value) {
  const text = KJDOCV2_unicodeNormalize_(value)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, ' ')
    .trim();

  if (!text || text.toLowerCase() === 'nan') return '';
  return text;
}

function KJDOCV2_normalizeHeader_(value) {
  return KJDOCV2_unicodeNormalize_(value)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function KJDOCV2_normalizeLooseText_(value) {
  return KJDOCV2_unicodeNormalize_(value)
    .replace(/[\u200B-\u200D\uFEFF]/g, '')
    .replace(/\s+/g, '')
    .toLowerCase();
}

function KJDOCV2_unicodeNormalize_(value) {
  const text = String(value == null ? '' : value);
  try {
    return text.normalize('NFKC');
  } catch (err) {
    return text;
  }
}

function KJDOCV2_buildTargetFolderName_(contractNo, customerName) {
  return KJDOCV2_sanitizeDriveName_(`${contractNo}_${customerName}`);
}

function KJDOCV2_sanitizeDriveName_(value) {
  return KJDOCV2_unicodeNormalize_(value)
    .replace(/[\\/:*?"<>|]/g, ' ')
    .replace(/[\u0000-\u001F\u007F]/g, '')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 180);
}

function KJDOCV2_escapeDriveQueryText_(value) {
  return String(value == null ? '' : value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");
}

function KJDOCV2_deleteClassifierTriggers_() {
  const handlers = new Set([
    'classifyKjDocumentsNow',
    'classifyKjDocumentsSafetyFullScan'
  ]);

  ScriptApp.getProjectTriggers().forEach(trigger => {
    if (handlers.has(trigger.getHandlerFunction())) {
      ScriptApp.deleteTrigger(trigger);
    }
  });
}


/****************************************************
 * 로그/이력
 ****************************************************/

function KJDOCV2_log_(status, contractNo, customerName, category, fileName, message) {
  KJ_DOC_LOG_BUFFER.push([
    new Date(),
    status,
    contractNo,
    customerName,
    category,
    fileName,
    message
  ]);
}

function KJDOCV2_flushLogs_() {
  if (!KJ_DOC_LOG_BUFFER.length) return;

  const ss = KJDOCV2_getSpreadsheet_();
  let sheet = ss.getSheetByName(KJ_DOC_CONFIG.LOG_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(KJ_DOC_CONFIG.LOG_SHEET_NAME);
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

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      KJ_DOC_LOG_BUFFER.length,
      KJ_DOC_LOG_BUFFER[0].length
    )
    .setValues(KJ_DOC_LOG_BUFFER);

  KJ_DOC_LOG_BUFFER = [];
}

function KJDOCV2_flushState_() {
  if (!KJ_DOC_STATE_BUFFER.length) return;

  const ss = KJDOCV2_getSpreadsheet_();
  let sheet = ss.getSheetByName(KJ_DOC_CONFIG.STATE_SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(KJ_DOC_CONFIG.STATE_SHEET_NAME);
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

  sheet
    .getRange(
      sheet.getLastRow() + 1,
      1,
      KJ_DOC_STATE_BUFFER.length,
      KJ_DOC_STATE_BUFFER[0].length
    )
    .setValues(KJ_DOC_STATE_BUFFER);

  KJ_DOC_STATE_BUFFER = [];
}
