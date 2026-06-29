/***** 고객사 공유드라이브 폴더 자동 생성 설정 *****/
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

  // 1회 실행 시 처리할 최대 행 수
  MAX_ROWS_PER_RUN: 200,

  // Apps Script 6분 제한 대비. 5분 정도에서 끊고 다음 실행으로 넘김.
  MAX_MILLIS_PER_RUN: 5 * 60 * 1000,

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

  MAX_ROWS_PER_RUN: 200,
  MAX_MILLIS_PER_RUN: 5 * 60 * 1000,

  PROP_NEXT_ROW: 'S1_FAILED_FOLDER_MOVE_NEXT_ROW',

  OUTPUT_HEADERS: {
    failedMoveStatus: '수주실패폴더이동상태',
    failedMoveUpdatedAt: '수주실패폴더이동일시'
  }
};


/***** 고객사 폴더 생성/보정 실행 함수 *****/

/**
 * 최초/수동 일괄 생성 함수.
 * 무조건 3행부터 돌지 않고, 공유드라이브 루트 + 수주실패 폴더를 먼저 스캔해서
 * 실제로 Drive에 고객번호 폴더가 없는 첫 행부터 시작함.
 */
function initCreateCustomerFoldersFromMaster() {
  const props = PropertiesService.getScriptProperties();
  props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');

  const detected = detectAndSetNextCustomerFolderRowFast();

  if (!detected || !detected.nextRow) {
    Logger.log('생성 필요한 고객사 폴더가 없습니다.');
    return;
  }

  continueCreateCustomerFoldersFromMaster();
}


/**
 * 이어서 실행 함수.
 * 공유드라이브 폴더 목록을 한 번에 색인한 뒤 누락/미연결 건만 처리.
 */
function continueCreateCustomerFoldersFromMaster() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const cfg = CUSTOMER_FOLDER_CFG;
    const sheet = getMasterSheet_();

    let headerMap = getHeaderMap_(sheet);
    headerMap = ensureOutputHeaders_(sheet, headerMap);

    assertHeader_(headerMap, '고객번호');
    assertHeader_(headerMap, '회사명');
    assertHeader_(headerMap, '수행사');

    const props = PropertiesService.getScriptProperties();
    const lastRow = sheet.getLastRow();

    if (lastRow < cfg.DATA_START_ROW) {
      props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
      Logger.log('마스터시트에 처리할 데이터가 없습니다.');
      return;
    }

    let row = Number(props.getProperty('S1_CUSTOMER_FOLDER_NEXT_ROW') || 0);

    if (!row || row < cfg.DATA_START_ROW) {
      const detected = detectNextCustomerFolderWorkRow_();
      if (!detected.nextRow) {
        props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
        Logger.log('생성 필요한 고객사 폴더가 없습니다.');
        return;
      }
      row = detected.nextRow;
      props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(row));
    }

    if (row > lastRow) {
      props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');
      Logger.log('처리할 행이 없습니다. 이미 완료되었습니다.');
      return;
    }

    const driveIndex = buildExistingCustomerFolderIndex_();
    const driveId = driveIndex.driveId;

    const lastCol = sheet.getLastColumn();
    const values = sheet
      .getRange(row, 1, lastRow - row + 1, lastCol)
      .getDisplayValues();

    const folderIdColIdx = col_(headerMap, cfg.OUTPUT_HEADERS.folderId) - 1;

    let scanned = 0;
    let created = 0;
    let relinked = 0;
    let skipped = 0;
    let errors = 0;

    const startedAt = Date.now();
    const maxMillis = cfg.MAX_MILLIS_PER_RUN || (5 * 60 * 1000);
    const logs = [];
    const parentCache = {};

    while (
      scanned < values.length &&
      scanned < cfg.MAX_ROWS_PER_RUN &&
      Date.now() - startedAt < maxMillis
    ) {
      const rowNum = row + scanned;
      const rowData = values[scanned];

      try {
        const result = createOrRelinkCustomerFolderFastForRow_({
          sheet,
          rowNum,
          rowData,
          headerMap,
          driveId,
          driveIndex,
          folderIdColIdx,
          parentCache
        });

        if (result.status === 'CREATED' || result.status === 'CREATED_IN_FAILED_FOLDER') {
          created++;
        } else if (
          result.status === 'RELINKED_BY_CUSTOMER_NO' ||
          result.status === 'RELINKED_RENAMED'
        ) {
          relinked++;
        } else {
          skipped++;
        }

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

    appendFolderLog_(logs);

    const nextRow = row + scanned;

    if (nextRow <= lastRow) {
      props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(nextRow));

      Logger.log(
        `이번 실행 완료: 스캔 ${scanned}건 / 신규생성 ${created}건 / 기존폴더연결 ${relinked}건 / 스킵 ${skipped}건 / 오류 ${errors}건. ` +
        `아직 남았습니다. continueCreateCustomerFoldersFromMaster()를 다시 실행하세요. 다음 시작 행: ${nextRow}`
      );

    } else {
      props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');

      Logger.log(
        `전체 완료: 스캔 ${scanned}건 / 신규생성 ${created}건 / 기존폴더연결 ${relinked}건 / 스킵 ${skipped}건 / 오류 ${errors}건`
      );
    }

  } finally {
    lock.releaseLock();
  }
}


/**
 * 영업전산/메일자동화에서 직접 호출할 함수.
 * 예: 고객 추가 저장 후 ensureCustomerFolderByCustomerNo(customerNo);
 */
function ensureCustomerFolderByCustomerNo(customerNo) {
  const sheet = getMasterSheet_();

  let headerMap = getHeaderMap_(sheet);
  headerMap = ensureOutputHeaders_(sheet, headerMap);

  assertHeader_(headerMap, '고객번호');

  const target = cleanValue_(customerNo);
  if (!target) {
    throw new Error('고객번호가 비어 있습니다.');
  }

  const customerNoCol = col_(headerMap, '고객번호');
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

  const driveId = getSharedDriveId_();

  for (let i = 0; i < values.length; i++) {
    const rowCustomerNo = cleanValue_(values[i][0]);

    if (rowCustomerNo === target) {
      const rowNum = CUSTOMER_FOLDER_CFG.DATA_START_ROW + i;
      return ensureCustomerFolderForSheetRow_(sheet, rowNum, driveId, headerMap);
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

  let headerMap = getHeaderMap_(sheet);
  headerMap = ensureOutputHeaders_(sheet, headerMap);

  const targetCols = [
    col_(headerMap, '고객번호'),
    col_(headerMap, '회사명'),
    col_(headerMap, '수행사')
  ];

  const editStartCol = e.range.getColumn();
  const editEndCol = editStartCol + e.range.getNumColumns() - 1;

  const touched = targetCols.some(c => c >= editStartCol && c <= editEndCol);
  if (!touched) return;

  const driveId = getSharedDriveId_();

  for (let row = Math.max(startRow, cfg.DATA_START_ROW); row <= endRow; row++) {
    ensureCustomerFolderForSheetRow_(sheet, row, driveId, headerMap);
  }
}


// 고객사 폴더 생성 전용: 설치형 트리거로만 실행
function customerFolderInstallableOnEdit(e) {
  const lock = LockService.getScriptLock();

  if (!lock.tryLock(5000)) {
    return;
  }

  try {
    handleCustomerFolderOnEdit(e);
  } catch (err) {
    Logger.log('[customerFolderInstallableOnEdit 오류] ' + (err && err.stack ? err.stack : err));
  } finally {
    lock.releaseLock();
  }
}


/**
 * 고객사 폴더 자동 생성용 설치형 onEdit 트리거 설치.
 * 최초 1회만 직접 실행.
 */
function installCustomerFolderOnEditTrigger() {
  const ss = getMasterSpreadsheet_();

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

function ensureCustomerFolderForSheetRow_(sheet, rowNum, driveId, headerMap) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const row = sheet
    .getRange(rowNum, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const customerNo = cleanValue_(row[col_(headerMap, '고객번호') - 1]);
  const company = cleanValue_(row[col_(headerMap, '회사명') - 1]);
  const vendorRaw = headerMap[normalizeHeader_('수행사')]
    ? cleanValue_(row[col_(headerMap, '수행사') - 1])
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

  const folderName = buildCustomerFolderName_(customerNo, company, vendor);

  const folderIdCol = col_(headerMap, cfg.OUTPUT_HEADERS.folderId);
  const existingFolderId = cleanValue_(row[folderIdCol - 1]);

  let folder = null;
  let status = '';

  if (existingFolderId) {
    folder = getDriveFile_(existingFolderId);

    if (folder && folder.trashed) {
      folder = null;
    }

    if (folder) {
      if (cfg.RENAME_IF_CHANGED && folder.name !== folderName) {
        folder = renameDriveFile_(folder.id, folderName);
        status = 'RENAMED';
      } else {
        status = 'EXISTING_ID';
      }
    }
  }

  if (!folder) {
    const found = findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, folderName);

    if (found && found.folder) {
      folder = found.folder;
      status = found.status || 'REUSED_BY_CUSTOMER_NO';

      if (cfg.RENAME_IF_CHANGED && folder.name !== folderName) {
        folder = renameDriveFile_(folder.id, folderName);
        status = 'REUSED_RENAMED';
      }
    }
  }

  if (!folder) {
    const parentCache = {};
    const parentId = getCustomerCreateParentIdForRow_(row, headerMap, driveId, parentCache);
    folder = createDriveFolder_(folderName, parentId);
    status = parentId === driveId ? 'CREATED' : 'CREATED_IN_FAILED_FOLDER';
  }

  if (cfg.CREATE_STANDARD_SUBFOLDERS) {
    ensureStandardSubfolders_(folder.id, driveId);
  }

  writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, folderName, status);

  return {
    status,
    message: '정상 처리',
    customerNo,
    company,
    vendor,
    folderName,
    folderId: folder.id,
    folderUrl: folder.webViewLink || `https://drive.google.com/drive/folders/${folder.id}`
  };
}


function createOrRelinkCustomerFolderFastForRow_(params) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const sheet = params.sheet;
  const rowNum = params.rowNum;
  const rowData = params.rowData;
  const headerMap = params.headerMap;
  const driveId = params.driveId;
  const driveIndex = params.driveIndex;
  const folderIdColIdx = params.folderIdColIdx;
  const parentCache = params.parentCache || {};

  const customerNo = cleanValue_(rowData[col_(headerMap, '고객번호') - 1]);
  const company = cleanValue_(rowData[col_(headerMap, '회사명') - 1]);
  const vendorRaw = headerMap[normalizeHeader_('수행사')]
    ? cleanValue_(rowData[col_(headerMap, '수행사') - 1])
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

  const existingFolderId = cleanValue_(rowData[folderIdColIdx]);

  if (existingFolderId) {
    return {
      status: 'SKIPPED_EXISTING_ID',
      message: '시트에 고객사폴더ID가 이미 있음',
      customerNo,
      company,
      vendor,
      folderId: existingFolderId
    };
  }

  const expectedFolderName = buildCustomerFolderName_(customerNo, company, vendor);
  const customerNoKey = normalizeCustomerNoKey_(customerNo);

  let mapped = driveIndex.byCustomerNo[customerNoKey];

  if (mapped && mapped.folder && mapped.folder.id) {
    let folder = mapped.folder;
    let status = 'RELINKED_BY_CUSTOMER_NO';

    if (cfg.RENAME_IF_CHANGED && folder.name !== expectedFolderName) {
      folder = renameDriveFile_(folder.id, expectedFolderName);
      status = 'RELINKED_RENAMED';

      driveIndex.byCustomerNo[customerNoKey] = {
        folder,
        location: mapped.location || '',
        parentId: mapped.parentId || ''
      };
    }

    writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, expectedFolderName, status);

    return {
      status,
      message: `고객번호 prefix 기준 기존 폴더 재연결 / 위치: ${mapped.location || ''}`,
      customerNo,
      company,
      vendor,
      folderName: expectedFolderName,
      folderId: folder.id
    };
  }

  const parentId = getCustomerCreateParentIdForRow_(rowData, headerMap, driveId, parentCache);
  const folder = createDriveFolder_(expectedFolderName, parentId);

  const createdStatus = parentId === driveId ? 'CREATED' : 'CREATED_IN_FAILED_FOLDER';

  writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, expectedFolderName, createdStatus);

  driveIndex.byCustomerNo[customerNoKey] = {
    folder,
    location: parentId === driveId ? 'ROOT' : 'FAILED',
    parentId
  };
  driveIndex.customerFolderCount++;

  return {
    status: createdStatus,
    message: parentId === driveId
      ? '신규 고객사 폴더 생성 완료'
      : '수주실패 고객사 폴더 생성 완료',
    customerNo,
    company,
    vendor,
    folderName: expectedFolderName,
    folderId: folder.id
  };
}


function ensureStandardSubfolders_(customerFolderId, driveId) {
  CUSTOMER_FOLDER_CFG.STANDARD_SUBFOLDERS.forEach(name => {
    const existing = findChildFolder_(customerFolderId, driveId, name);
    if (!existing) {
      createDriveFolder_(name, customerFolderId);
    }
  });
}


function writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, folderName, status) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const folderUrl = folder.webViewLink || `https://drive.google.com/drive/folders/${folder.id}`;
  const nowText = Utilities.formatDate(new Date(), cfg.TZ, 'yyyy-MM-dd HH:mm:ss');

  sheet.getRange(rowNum, col_(headerMap, cfg.OUTPUT_HEADERS.folderId)).setValue(folder.id);
  sheet.getRange(rowNum, col_(headerMap, cfg.OUTPUT_HEADERS.folderUrl)).setValue(folderUrl);
  sheet.getRange(rowNum, col_(headerMap, cfg.OUTPUT_HEADERS.folderName)).setValue(folderName);
  sheet.getRange(rowNum, col_(headerMap, cfg.OUTPUT_HEADERS.folderStatus)).setValue(status);
  sheet.getRange(rowNum, col_(headerMap, cfg.OUTPUT_HEADERS.folderUpdatedAt)).setValue(nowText);
}


/***** 빠른 진행상황 탐지/Drive 색인 함수 *****/

function detectAndSetNextCustomerFolderRowFast() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const detected = detectNextCustomerFolderWorkRow_();
    const props = PropertiesService.getScriptProperties();

    if (detected.nextRow) {
      props.setProperty('S1_CUSTOMER_FOLDER_NEXT_ROW', String(detected.nextRow));

      Logger.log(
        `다음 작업 시작 행 탐지 완료: ${detected.nextRow}행 / 고객번호 ${detected.customerNo || ''} / 회사명 ${detected.company || ''}` +
        ` / 사유: ${detected.reason || ''}` +
        ` / Drive 탐지 고객폴더 ${detected.driveCustomerFolderCount}개` +
        ` / 시트 유효고객 ${detected.validCustomerCount}건` +
        ` / 폴더ID 기재 ${detected.sheetFolderIdCount}건`
      );
    } else {
      props.deleteProperty('S1_CUSTOMER_FOLDER_NEXT_ROW');

      Logger.log(
        `생성 필요한 고객사 폴더 없음` +
        ` / Drive 탐지 고객폴더 ${detected.driveCustomerFolderCount}개` +
        ` / 시트 유효고객 ${detected.validCustomerCount}건` +
        ` / 폴더ID 기재 ${detected.sheetFolderIdCount}건`
      );
    }

    return detected;

  } finally {
    lock.releaseLock();
  }
}


/**
 * 수동 점검용.
 * 실행하면 지금 어디부터 생성해야 하는지 로그만 찍음.
 */
function reportCustomerFolderCreateProgressFast() {
  const detected = detectNextCustomerFolderWorkRow_();

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
 * 기준은 시트 폴더ID가 아니라 Drive에 고객번호_ 폴더가 실제 존재하는지 여부.
 */
function detectNextCustomerFolderWorkRow_() {
  const cfg = CUSTOMER_FOLDER_CFG;
  const sheet = getMasterSheet_();

  let headerMap = getHeaderMap_(sheet);
  headerMap = ensureOutputHeaders_(sheet, headerMap);

  assertHeader_(headerMap, '고객번호');
  assertHeader_(headerMap, '회사명');

  const lastRow = sheet.getLastRow();

  if (lastRow < cfg.DATA_START_ROW) {
    return {
      nextRow: 0,
      driveCustomerFolderCount: 0,
      validCustomerCount: 0,
      sheetFolderIdCount: 0,
      maxCustomerNoInDrive: ''
    };
  }

  const driveIndex = buildExistingCustomerFolderIndex_();

  const lastCol = sheet.getLastColumn();
  const values = sheet
    .getRange(cfg.DATA_START_ROW, 1, lastRow - cfg.DATA_START_ROW + 1, lastCol)
    .getDisplayValues();

  const customerNoIdx = col_(headerMap, '고객번호') - 1;
  const companyIdx = col_(headerMap, '회사명') - 1;
  const folderIdIdx = col_(headerMap, cfg.OUTPUT_HEADERS.folderId) - 1;

  let validCustomerCount = 0;
  let sheetFolderIdCount = 0;

  for (let i = 0; i < values.length; i++) {
    const rowData = values[i];
    const rowNum = cfg.DATA_START_ROW + i;

    const customerNo = cleanValue_(rowData[customerNoIdx]);
    const company = cleanValue_(rowData[companyIdx]);
    const folderId = cleanValue_(rowData[folderIdIdx]);

    if (!customerNo || !company) {
      continue;
    }

    validCustomerCount++;

    if (folderId) {
      sheetFolderIdCount++;
    }

    const customerNoKey = normalizeCustomerNoKey_(customerNo);

    // 핵심: Drive에 고객번호_ 폴더가 없을 때만 생성 대상 시작 행으로 잡음.
    // 이미 Drive에 있으면, 시트 폴더ID가 비어 있어도 생성 막힘을 만들지 않음.
    if (!driveIndex.byCustomerNo[customerNoKey]) {
      return {
        nextRow: rowNum,
        customerNo,
        company,
        reason: 'Drive에 고객번호 prefix 폴더 없음',
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
function buildExistingCustomerFolderIndex_() {
  const driveId = getSharedDriveId_();

  const byCustomerNo = {};
  let customerFolderCount = 0;
  let maxCustomerNo = 0;

  const rootFolders = listDirectChildFoldersPaged_(driveId, driveId);

  rootFolders.forEach(folder => {
    const customerNoKey = extractCustomerNoKeyFromFolderName_(folder.name);

    if (!customerNoKey) return;

    if (!byCustomerNo[customerNoKey]) {
      byCustomerNo[customerNoKey] = {
        folder,
        location: 'ROOT',
        parentId: driveId
      };
    }

    customerFolderCount++;
    maxCustomerNo = Math.max(maxCustomerNo, Number(customerNoKey) || 0);
  });

  const failedParentFolder = getFailedParentFolderIfExistsFromRoot_(rootFolders, driveId);

  if (failedParentFolder && failedParentFolder.id) {
    const failedFolders = listDirectChildFoldersPaged_(failedParentFolder.id, driveId);

    failedFolders.forEach(folder => {
      const customerNoKey = extractCustomerNoKeyFromFolderName_(folder.name);

      if (!customerNoKey) return;

      // 루트에 같은 고객번호가 있으면 루트 우선.
      if (!byCustomerNo[customerNoKey]) {
        byCustomerNo[customerNoKey] = {
          folder,
          location: 'FAILED',
          parentId: failedParentFolder.id
        };
      }

      customerFolderCount++;
      maxCustomerNo = Math.max(maxCustomerNo, Number(customerNoKey) || 0);
    });
  }

  return {
    driveId,
    byCustomerNo,
    customerFolderCount,
    maxCustomerNo: maxCustomerNo ? String(maxCustomerNo) : '',
    rootFolderCount: rootFolders.length,
    failedParentFolderId: failedParentFolder ? failedParentFolder.id : ''
  };
}


function listDirectChildFoldersPaged_(parentFolderId, driveId) {
  const q = [
    `${driveQueryString_(parentFolderId)} in parents`,
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
      '&fields=nextPageToken,files(id,name,webViewLink,trashed,parents)';

    if (pageToken) {
      path += '&pageToken=' + encodeURIComponent(pageToken);
    }

    const data = driveFetch_(path, { method: 'get' });

    (data.files || []).forEach(file => folders.push(file));
    pageToken = data.nextPageToken || '';

  } while (pageToken);

  return folders;
}


function getFailedParentFolderIfExistsFromRoot_(rootFolders, driveId) {
  const failedName = FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME || '수주실패';

  const found = (rootFolders || []).find(folder => cleanValue_(folder.name) === failedName);

  if (found) {
    return found;
  }

  try {
    return findChildFolder_(driveId, driveId, failedName);
  } catch (err) {
    return null;
  }
}


function findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, expectedFolderName) {
  const exactRoot = findChildFolder_(driveId, driveId, expectedFolderName);
  if (exactRoot) {
    return {
      status: 'REUSED_BY_NAME_ROOT',
      folder: exactRoot,
      location: 'ROOT'
    };
  }

  const foundRoot = findCustomerFolderByCustomerNoPrefixInParent_(driveId, driveId, customerNo);
  if (foundRoot.status === 'FOUND') {
    return {
      status: 'REUSED_BY_CUSTOMER_NO_ROOT',
      folder: foundRoot.folder,
      location: 'ROOT'
    };
  }

  const failedParent = findChildFolder_(driveId, driveId, FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME);
  if (failedParent) {
    const foundFailed = findCustomerFolderByCustomerNoPrefixInParent_(driveId, failedParent.id, customerNo);
    if (foundFailed.status === 'FOUND') {
      return {
        status: 'REUSED_BY_CUSTOMER_NO_FAILED',
        folder: foundFailed.folder,
        location: 'FAILED'
      };
    }
  }

  return null;
}


function extractCustomerNoKeyFromFolderName_(folderName) {
  const s = cleanValue_(folderName);
  const m = s.match(/^(\d+)_/);

  if (!m) return '';

  return normalizeCustomerNoKey_(m[1]);
}


function normalizeCustomerNoKey_(value) {
  return cleanValue_(value)
    .replace(/\.0$/, '')
    .trim();
}


/***** 고객사 폴더명 일괄 업데이트 *****/

function manualUpdateAllCustomerFolderNames() {
  PropertiesService.getScriptProperties().deleteProperty('S1_CUSTOMER_FOLDER_RENAME_NEXT_ROW');
  continueUpdateAllCustomerFolderNames();
}


function continueUpdateAllCustomerFolderNames() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const cfg = CUSTOMER_FOLDER_CFG;
    const sheet = getMasterSheet_();

    let headerMap = getHeaderMap_(sheet);
    headerMap = ensureOutputHeaders_(sheet, headerMap);

    assertHeader_(headerMap, '고객번호');
    assertHeader_(headerMap, '회사명');
    assertHeader_(headerMap, '수행사');

    const driveId = getSharedDriveId_();
    const props = PropertiesService.getScriptProperties();

    const lastRow = sheet.getLastRow();
    let row = Number(props.getProperty('S1_CUSTOMER_FOLDER_RENAME_NEXT_ROW') || cfg.DATA_START_ROW);

    if (row > lastRow) {
      props.deleteProperty('S1_CUSTOMER_FOLDER_RENAME_NEXT_ROW');
      Logger.log('폴더명 업데이트 대상 행이 없습니다.');
      return;
    }

    let processed = 0;
    let renamed = 0;
    let ok = 0;
    let skipped = 0;
    let notFound = 0;
    let errors = 0;

    const startedAt = Date.now();
    const maxMillis = cfg.MAX_MILLIS_PER_RUN || (5 * 60 * 1000);

    const logs = [];

    while (
      row <= lastRow &&
      processed < cfg.MAX_ROWS_PER_RUN &&
      Date.now() - startedAt < maxMillis
    ) {
      try {
        const result = updateCustomerFolderNameForRow_(sheet, row, driveId, headerMap);

        if (result.status === 'RENAMED') renamed++;
        else if (result.status === 'OK') ok++;
        else if (result.status === 'SKIPPED') skipped++;
        else if (result.status === 'NOT_FOUND') notFound++;

        logs.push([
          new Date(),
          row,
          result.customerNo || '',
          result.company || '',
          result.vendor || '',
          result.expectedFolderName || '',
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

    appendFolderLog_(logs);

    if (row <= lastRow) {
      props.setProperty('S1_CUSTOMER_FOLDER_RENAME_NEXT_ROW', String(row));

      Logger.log(
        `이번 실행 완료: 처리 ${processed}건 / 변경 ${renamed}건 / 정상 ${ok}건 / 스킵 ${skipped}건 / 미발견 ${notFound}건 / 오류 ${errors}건. ` +
        `아직 남았습니다. continueUpdateAllCustomerFolderNames()를 다시 실행하세요. 다음 시작 행: ${row}`
      );

    } else {
      props.deleteProperty('S1_CUSTOMER_FOLDER_RENAME_NEXT_ROW');

      Logger.log(
        `전체 완료: 처리 ${processed}건 / 변경 ${renamed}건 / 정상 ${ok}건 / 스킵 ${skipped}건 / 미발견 ${notFound}건 / 오류 ${errors}건`
      );
    }

  } finally {
    lock.releaseLock();
  }
}


function updateCustomerFolderNameForRow_(sheet, rowNum, driveId, headerMap) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const row = sheet
    .getRange(rowNum, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const customerNo = cleanValue_(row[col_(headerMap, '고객번호') - 1]);
  const company = cleanValue_(row[col_(headerMap, '회사명') - 1]);
  const vendorRaw = cleanValue_(row[col_(headerMap, '수행사') - 1]);
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

  const expectedFolderName = buildCustomerFolderName_(customerNo, company, vendor);

  const folderIdCol = col_(headerMap, cfg.OUTPUT_HEADERS.folderId);
  const existingFolderId = cleanValue_(row[folderIdCol - 1]);

  let folder = null;
  let findMessage = '';

  if (existingFolderId) {
    folder = getDriveFile_(existingFolderId);

    if (folder && folder.trashed) {
      folder = null;
      findMessage = '시트의 폴더ID가 휴지통 상태';
    }
  }

  if (!folder) {
    const found = findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, expectedFolderName);

    if (found && found.folder) {
      folder = found.folder;
      findMessage = `고객번호 prefix로 폴더 발견 / 위치: ${found.location || ''}`;
    } else {
      return {
        status: 'NOT_FOUND',
        message: findMessage || '고객번호 prefix로 폴더를 찾지 못함',
        customerNo,
        company,
        vendor,
        expectedFolderName
      };
    }
  }

  let status = 'OK';
  let message = '이미 최신 폴더명';

  if (folder.name !== expectedFolderName) {
    folder = renameDriveFile_(folder.id, expectedFolderName);
    status = 'RENAMED';
    message = `폴더명 변경 완료: ${expectedFolderName}`;
  }

  writeFolderInfoToSheet_(sheet, rowNum, headerMap, folder, expectedFolderName, status);

  return {
    status,
    message: findMessage ? `${message} / ${findMessage}` : message,
    customerNo,
    company,
    vendor,
    expectedFolderName,
    folderId: folder.id,
    folderUrl: folder.webViewLink || `https://drive.google.com/drive/folders/${folder.id}`
  };
}


/***** 고객사 폴더 내부 하위폴더 정리 *****/

function previewTrashCustomerChildFolders() {
  runTrashCustomerChildFolders_({
    dryRun: true,
    onlyStandardSubfolders: false
  });
}


function trashCustomerChildFolders() {
  runTrashCustomerChildFolders_({
    dryRun: false,
    onlyStandardSubfolders: false
  });
}


function trashOnlyStandardCustomerChildFolders() {
  runTrashCustomerChildFolders_({
    dryRun: false,
    onlyStandardSubfolders: true
  });
}


function runTrashCustomerChildFolders_(options) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const cfg = CUSTOMER_FOLDER_CFG;
    const sheet = getMasterSheet_();

    let headerMap = getHeaderMap_(sheet);
    headerMap = ensureOutputHeaders_(sheet, headerMap);

    assertHeader_(headerMap, '고객번호');
    assertHeader_(headerMap, '회사명');

    const driveId = getSharedDriveId_();
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

        const customerNo = cleanValue_(row[col_(headerMap, '고객번호') - 1]);
        const company = cleanValue_(row[col_(headerMap, '회사명') - 1]);

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

        const customerFolder = getCustomerFolderForCleanupRow_(row, rowNum, driveId, headerMap, customerNo);

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

        let childFolders = listChildFolders_(customerFolder.id, driveId);

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
            trashDriveFile_(child.id);
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

    appendFolderLog_(logs);

    Logger.log(
      `${options.dryRun ? '미리보기' : '삭제'} 완료: ` +
      `처리 고객폴더 ${processedCustomers}개 / 대상 하위폴더 ${targetFolders}개 / 휴지통 이동 ${trashedFolders}개 / 스킵 ${skipped}건 / 오류 ${errors}건`
    );

  } finally {
    lock.releaseLock();
  }
}


function getCustomerFolderForCleanupRow_(row, rowNum, driveId, headerMap, customerNo) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const folderIdCol = col_(headerMap, cfg.OUTPUT_HEADERS.folderId);
  const existingFolderId = cleanValue_(row[folderIdCol - 1]);

  let folder = null;

  if (existingFolderId) {
    folder = getDriveFile_(existingFolderId);

    if (folder && folder.trashed) {
      folder = null;
    }
  }

  if (!folder) {
    const found = findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, '');
    if (found && found.folder) {
      folder = found.folder;
    }
  }

  return folder;
}


function listChildFolders_(parentFolderId, driveId) {
  const q = [
    `${driveQueryString_(parentFolderId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `trashed = false`
  ].join(' and ');

  const path =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=allDrives' +
    '&pageSize=1000' +
    '&q=' + encodeURIComponent(q) +
    '&fields=files(id,name,webViewLink,trashed)';

  const data = driveFetch_(path, { method: 'get' });
  return data.files || [];
}


function trashDriveFile_(fileId) {
  return driveFetch_(
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
  continueMoveFailedCustomerFolders();
}


function continueMoveFailedCustomerFolders() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const cfg = CUSTOMER_FOLDER_CFG;
    const failedCfg = FAILED_CUSTOMER_FOLDER_CFG;

    const sheet = getMasterSheet_();

    let headerMap = getHeaderMap_(sheet);
    headerMap = ensureOutputHeaders_(sheet, headerMap);
    headerMap = ensureFailedFolderOutputHeaders_(sheet, headerMap);

    assertHeader_(headerMap, '고객번호');
    assertHeader_(headerMap, '회사명');

    const statusHeaderName = findFirstExistingHeaderName_(headerMap, failedCfg.STATUS_HEADER_CANDIDATES);
    if (!statusHeaderName) {
      throw new Error(
        '상태값 헤더를 찾지 못했습니다. 후보: ' +
        failedCfg.STATUS_HEADER_CANDIDATES.join(', ')
      );
    }

    const driveId = getSharedDriveId_();
    const failedParentFolder = ensureFailedParentFolder_(driveId);

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
        const result = moveFailedCustomerFolderForRow_({
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

    appendFolderLog_(logs);

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
    lock.releaseLock();
  }
}


function moveFailedCustomerFolderByCustomerNo(customerNo) {
  const sheet = getMasterSheet_();

  let headerMap = getHeaderMap_(sheet);
  headerMap = ensureOutputHeaders_(sheet, headerMap);
  headerMap = ensureFailedFolderOutputHeaders_(sheet, headerMap);

  assertHeader_(headerMap, '고객번호');

  const statusHeaderName = findFirstExistingHeaderName_(
    headerMap,
    FAILED_CUSTOMER_FOLDER_CFG.STATUS_HEADER_CANDIDATES
  );

  if (!statusHeaderName) {
    throw new Error('상태값 헤더를 찾지 못했습니다.');
  }

  const target = cleanValue_(customerNo);
  if (!target) {
    throw new Error('고객번호가 비어 있습니다.');
  }

  const customerNoCol = col_(headerMap, '고객번호');
  const lastRow = sheet.getLastRow();

  const values = sheet
    .getRange(
      CUSTOMER_FOLDER_CFG.DATA_START_ROW,
      customerNoCol,
      lastRow - CUSTOMER_FOLDER_CFG.DATA_START_ROW + 1,
      1
    )
    .getDisplayValues();

  const driveId = getSharedDriveId_();
  const failedParentFolder = ensureFailedParentFolder_(driveId);

  for (let i = 0; i < values.length; i++) {
    const rowCustomerNo = cleanValue_(values[i][0]);

    if (rowCustomerNo === target) {
      const rowNum = CUSTOMER_FOLDER_CFG.DATA_START_ROW + i;

      return moveFailedCustomerFolderForRow_({
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


function moveFailedCustomerFolderForRow_(params) {
  const sheet = params.sheet;
  const rowNum = params.rowNum;
  const headerMap = params.headerMap;
  const statusHeaderName = params.statusHeaderName;
  const driveId = params.driveId;
  const failedParentFolderId = params.failedParentFolderId;

  const row = sheet
    .getRange(rowNum, 1, 1, sheet.getLastColumn())
    .getDisplayValues()[0];

  const customerNo = cleanValue_(row[col_(headerMap, '고객번호') - 1]);
  const company = cleanValue_(row[col_(headerMap, '회사명') - 1]);

  let vendor = '';
  if (headerMap[normalizeHeader_('수행사')]) {
    vendor = cleanValue_(row[col_(headerMap, '수행사') - 1]);
  }

  const statusValue = cleanValue_(row[col_(headerMap, statusHeaderName) - 1]);

  if (!customerNo || !company) {
    return {
      status: 'SKIPPED',
      message: '고객번호 또는 회사명 공란',
      customerNo,
      company,
      vendor
    };
  }

  if (!isFailedStatus_(statusValue)) {
    writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
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

  const folder = getCustomerFolderForFailedMove_(row, driveId, headerMap, customerNo);

  if (!folder) {
    writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
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

  const expectedFolderName = buildCustomerFolderName_(
    customerNo,
    company,
    vendor || CUSTOMER_FOLDER_CFG.EMPTY_VENDOR_TEXT
  );

  let currentFolder = folder;

  if (currentFolder.name !== expectedFolderName) {
    currentFolder = renameDriveFile_(currentFolder.id, expectedFolderName);
  }

  const folderWithParents = getDriveFileWithParents_(currentFolder.id);

  if ((folderWithParents.parents || []).indexOf(failedParentFolderId) !== -1) {
    writeFolderInfoToSheet_(sheet, rowNum, headerMap, currentFolder, expectedFolderName, 'ALREADY_IN_FAILED_FOLDER');
    writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
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

  moveDriveFileToFolder_(currentFolder.id, failedParentFolderId, folderWithParents.parents || []);

  writeFolderInfoToSheet_(sheet, rowNum, headerMap, currentFolder, expectedFolderName, 'MOVED_TO_FAILED_FOLDER');
  writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, {
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


function ensureFailedParentFolder_(driveId) {
  const name = FAILED_CUSTOMER_FOLDER_CFG.FAILED_PARENT_FOLDER_NAME;

  const existing = findChildFolder_(driveId, driveId, name);

  if (existing) {
    return existing;
  }

  return createDriveFolder_(name, driveId);
}


function getCustomerFolderForFailedMove_(row, driveId, headerMap, customerNo) {
  const cfg = CUSTOMER_FOLDER_CFG;

  const folderIdCol = col_(headerMap, cfg.OUTPUT_HEADERS.folderId);
  const existingFolderId = cleanValue_(row[folderIdCol - 1]);

  let folder = null;

  if (existingFolderId) {
    folder = getDriveFile_(existingFolderId);

    if (folder && folder.trashed) {
      folder = null;
    }
  }

  if (!folder) {
    const found = findExistingCustomerFolderByCustomerNoAnywhere_(driveId, customerNo, '');
    if (found && found.folder) {
      folder = found.folder;
    }
  }

  return folder;
}


function findCustomerFolderByCustomerNoPrefixInParent_(driveId, parentFolderId, customerNo) {
  const prefix = sanitizeFolderPart_(customerNo) + '_';

  const q = [
    `${driveQueryString_(parentFolderId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `name contains ${driveQueryString_(prefix)}`,
    `trashed = false`
  ].join(' and ');

  const path =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=drive' +
    '&driveId=' + encodeURIComponent(driveId) +
    '&pageSize=50' +
    '&q=' + encodeURIComponent(q) +
    '&fields=files(id,name,webViewLink,trashed)';

  const data = driveFetch_(path, { method: 'get' });
  const files = (data.files || []).filter(f => cleanValue_(f.name).startsWith(prefix));

  if (files.length === 0) {
    return {
      status: 'NOT_FOUND',
      message: `${prefix} 로 시작하는 폴더 없음`
    };
  }

  if (files.length > 1) {
    return {
      status: 'FOUND',
      folder: files[0],
      message: `동일 고객번호 prefix 폴더 ${files.length}개 발견. 첫 번째 기준 처리`
    };
  }

  return {
    status: 'FOUND',
    folder: files[0],
    message: '고객번호 prefix 폴더 발견'
  };
}


function getDriveFileWithParents_(fileId) {
  return driveFetch_(
    'files/' + encodeURIComponent(fileId) +
    '?supportsAllDrives=true&fields=id,name,webViewLink,trashed,mimeType,parents',
    { method: 'get' }
  );
}


function moveDriveFileToFolder_(fileId, targetParentId, currentParentIds) {
  const removeParents = (currentParentIds || [])
    .filter(parentId => parentId !== targetParentId)
    .join(',');

  let path =
    'files/' + encodeURIComponent(fileId) +
    '?supportsAllDrives=true' +
    '&addParents=' + encodeURIComponent(targetParentId) +
    '&fields=id,name,parents,webViewLink';

  if (removeParents) {
    path += '&removeParents=' + encodeURIComponent(removeParents);
  }

  return driveFetch_(path, {
    method: 'patch',
    payload: {}
  });
}


function isFailedStatus_(statusValue) {
  const v = cleanValue_(statusValue);

  if (!v) return false;

  return FAILED_CUSTOMER_FOLDER_CFG.FAILED_STATUS_KEYWORDS.some(keyword => {
    const k = cleanValue_(keyword);
    return k && v.indexOf(k) !== -1;
  });
}


function ensureFailedFolderOutputHeaders_(sheet, headerMap) {
  const required = Object.values(FAILED_CUSTOMER_FOLDER_CFG.OUTPUT_HEADERS);

  required.forEach(headerName => {
    const key = normalizeHeader_(headerName);

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


function writeFailedMoveInfoToSheet_(sheet, rowNum, headerMap, info) {
  const nowText = Utilities.formatDate(new Date(), CUSTOMER_FOLDER_CFG.TZ, 'yyyy-MM-dd HH:mm:ss');

  sheet
    .getRange(rowNum, col_(headerMap, FAILED_CUSTOMER_FOLDER_CFG.OUTPUT_HEADERS.failedMoveStatus))
    .setValue(info.status || '');

  sheet
    .getRange(rowNum, col_(headerMap, FAILED_CUSTOMER_FOLDER_CFG.OUTPUT_HEADERS.failedMoveUpdatedAt))
    .setValue(nowText);
}


function findFirstExistingHeaderName_(headerMap, headerNames) {
  for (let i = 0; i < headerNames.length; i++) {
    const name = headerNames[i];

    if (headerMap[normalizeHeader_(name)]) {
      return name;
    }
  }

  return '';
}


function getCustomerCreateParentIdForRow_(rowData, headerMap, driveId, parentCache) {
  if (isFailedCustomerRowByData_(rowData, headerMap)) {
    if (!parentCache.failedParentFolderId) {
      const failedFolder = ensureFailedParentFolder_(driveId);
      parentCache.failedParentFolderId = failedFolder.id;
    }

    return parentCache.failedParentFolderId;
  }

  return driveId;
}


function isFailedCustomerRowByData_(rowData, headerMap) {
  const statusHeaderName = findFirstExistingHeaderName_(
    headerMap,
    FAILED_CUSTOMER_FOLDER_CFG.STATUS_HEADER_CANDIDATES
  );

  if (!statusHeaderName) {
    return false;
  }

  const statusValue = cleanValue_(rowData[col_(headerMap, statusHeaderName) - 1]);

  return isFailedStatus_(statusValue);
}


/***** 스프레드시트/헤더 유틸 *****/

function getMasterSpreadsheet_() {
  const id = cleanValue_(CUSTOMER_FOLDER_CFG.MASTER_SPREADSHEET_ID);

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


function getMasterSheet_() {
  const ss = getMasterSpreadsheet_();
  const sheet = ss.getSheetByName(CUSTOMER_FOLDER_CFG.MASTER_SHEET_NAME);

  if (!sheet) {
    throw new Error(`마스터시트를 찾지 못했습니다: ${CUSTOMER_FOLDER_CFG.MASTER_SHEET_NAME}`);
  }

  return sheet;
}


function getHeaderMap_(sheet) {
  const headerRow = CUSTOMER_FOLDER_CFG.HEADER_ROW;
  const lastCol = sheet.getLastColumn();

  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach((h, i) => {
    const key = normalizeHeader_(h);
    if (key && !map[key]) {
      map[key] = i + 1;
    }
  });

  return map;
}


function ensureOutputHeaders_(sheet, headerMap) {
  const cfg = CUSTOMER_FOLDER_CFG;
  const required = Object.values(cfg.OUTPUT_HEADERS);

  required.forEach(headerName => {
    const key = normalizeHeader_(headerName);

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


function assertHeader_(headerMap, headerName) {
  if (!headerMap[normalizeHeader_(headerName)]) {
    throw new Error(`필수 헤더를 찾지 못했습니다: ${headerName}`);
  }
}


function col_(headerMap, headerName) {
  const key = normalizeHeader_(headerName);
  const c = headerMap[key];

  if (!c) {
    throw new Error(`헤더 컬럼을 찾지 못했습니다: ${headerName}`);
  }

  return c;
}


/***** 문자열/로그 유틸 *****/

function buildCustomerFolderName_(customerNo, company, vendor) {
  const parts = [
    sanitizeFolderPart_(customerNo),
    sanitizeFolderPart_(company),
    sanitizeFolderPart_(vendor)
  ];

  let name = parts.join('_').replace(/_+/g, '_').trim();

  if (name.length > 180) {
    name = name.slice(0, 180).trim();
  }

  return name;
}


function sanitizeFolderPart_(value) {
  return cleanValue_(value)
    .replace(/[\/\\:*?"<>|#\[\]\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function cleanValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}


function normalizeHeader_(value) {
  return cleanValue_(value).replace(/\s+/g, '');
}


function appendFolderLog_(rows) {
  if (!rows || rows.length === 0) return;

  const ss = getMasterSpreadsheet_();
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

function getSharedDriveId_() {
  const cfg = CUSTOMER_FOLDER_CFG;

  if (cleanValue_(cfg.SHARED_DRIVE_ID)) {
    return cleanValue_(cfg.SHARED_DRIVE_ID);
  }

  const props = PropertiesService.getScriptProperties();
  const cached = props.getProperty('S1_CUSTOMER_SHARED_DRIVE_ID');

  if (cached) {
    return cached;
  }

  const q = `name = ${driveQueryString_(cfg.SHARED_DRIVE_NAME)}`;

  const data = driveFetch_(
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


function findChildFolder_(parentId, driveId, folderName) {
  if (!folderName) return null;

  const q = [
    `${driveQueryString_(parentId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `name = ${driveQueryString_(folderName)}`,
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
    '&fields=files(id,name,webViewLink,trashed)';

  const data = driveFetch_(path, { method: 'get' });
  const files = data.files || [];

  return files.length ? files[0] : null;
}


function createDriveFolder_(folderName, parentId) {
  return driveFetch_(
    'files?supportsAllDrives=true&fields=id,name,webViewLink',
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


function getDriveFile_(fileId) {
  try {
    return driveFetch_(
      'files/' + encodeURIComponent(fileId) + '?supportsAllDrives=true&fields=id,name,webViewLink,trashed,mimeType',
      { method: 'get' }
    );
  } catch (err) {
    return null;
  }
}


function renameDriveFile_(fileId, newName) {
  return driveFetch_(
    'files/' + encodeURIComponent(fileId) + '?supportsAllDrives=true&fields=id,name,webViewLink,trashed',
    {
      method: 'patch',
      payload: {
        name: newName
      }
    }
  );
}


function driveFetch_(path, options) {
  const url = 'https://www.googleapis.com/drive/v3/' + path;

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

  const res = UrlFetchApp.fetch(url, params);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Drive API 오류 ${code}: ${text}`);
  }

  return text ? JSON.parse(text) : {};
}


function driveQueryString_(value) {
  const s = cleanValue_(value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");

  return `'${s}'`;
}
