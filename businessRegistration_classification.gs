/*******************************************************
 * 사업자등록증 일괄 복사.gs
 * v2 단일 실행 함수 구조
 *
 * 목적:
 * - 공유드라이브 "S1 고객사 파일 관리" 안의 "사업자등록증 일괄" 폴더에 있는
 *   사업자등록증 파일을 상위 디렉토리의 고객사별 폴더로 복사한다.
 * - 이제 복사/미리보기는 기능별 단일 함수만 반복 실행하면 이어서 처리된다.
 *
 * 기준:
 * 1) 파일명 앞머리에 계약번호가 있으면:
 *    수주확정/계약완료 시트에서 계약번호 → 고객번호를 찾아 고객사 폴더에 복사
 *
 * 2) 파일명 앞머리에 계약번호가 없으면:
 *    파일명에서 고객사명을 정규화해서 기존 고객사 폴더명 또는 시트의 회사명과 매칭
 *
 * 3) 고객번호를 찾았는데 고객사 폴더가 없으면:
 *    상위 디렉토리에 고객사 폴더를 새로 만들고 복사
 *
 * 주의:
 * - 사업자등록증 파일명 앞 번호 = 계약번호
 * - 고객사 폴더명 앞 번호 = 고객번호
 * - 두 번호를 절대 같은 번호로 보지 않음
 *******************************************************/


/***** 설정 *****/

const BR_COPY_CFG = {
  // 마스터시트에 바인딩된 Apps Script면 빈 값 유지.
  // 독립형 Apps Script면 스프레드시트 ID 입력.
  SPREADSHEET_ID: '',

  // 원본 공유드라이브명
  SHARED_DRIVE_NAME: 'S1 고객사 파일 관리',

  // 가능하면 비워둬도 됨.
  // 이름 조회 실패 시 공유드라이브 URL의 folders/ 뒤 ID를 여기에 직접 입력.
  SHARED_DRIVE_ID: '',

  // 공유드라이브 루트 바로 아래의 일괄 수집 폴더명
  SOURCE_FOLDER_NAME: '사업자등록증 일괄',

  // 매칭 실패 파일 복사 위치
  UNMATCHED_FOLDER_NAME: '사업자등록증_미매칭',

  // 계약번호 매핑용 시트명 후보
  CONTRACT_SHEET_NAMES: [
    '수주확정/계약완료',
    '수주확정',
    '계약완료'
  ],

  // 고객정보 보강용 마스터 시트명 후보
  MASTER_SHEET_NAMES: [
    '마스터시트(신규)',
    '마스터시트'
  ],

  // 헤더 자동 탐지 범위
  HEADER_SCAN_ROWS: 10,

  // 1회 실행 제한
  MAX_FILES_PER_RUN: 80,
  MAX_MILLIS_PER_RUN: 5 * 60 * 1000,

  // 같은 이름 파일이 이미 고객사 폴더에 있을 때 처리
  // false: 기존 파일 있으면 스킵. 재실행해도 중복 생성 안 됨.
  // true: 기존 파일 휴지통 이동 후 새로 복사.
  REPLACE_SAME_NAME_FILE: false,

  // 고객사명만 있고 고객번호를 못 찾은 파일 처리
  // false: 사업자등록증_미매칭 폴더로 복사
  // true: 고객번호 없는 회사명 폴더를 루트에 생성. 권장하지 않음.
  CREATE_COMPANY_FOLDER_WITHOUT_CUSTOMER_NO: false,

  // 고객번호/수행사 없을 때 폴더명 기본값
  EMPTY_VENDOR_TEXT: '수행사미정',
  UNKNOWN_CUSTOMER_NO_TEXT: '고객번호미확인',

  // 로그 시트명
  LOG_SHEET_NAME: '사업자등록증_일괄복사_LOG',

  // 진행 저장 키
  PROP_NEXT_FILE_INDEX: 'S1_BR_COPY_NEXT_FILE_INDEX',
  PROP_PREVIEW_NEXT_FILE_INDEX: 'S1_BR_PREVIEW_NEXT_FILE_INDEX',

  TZ: 'Asia/Seoul'
};



/***** 공개 실행 함수 *****/

/**
 * 사업자등록증 일괄 복사 실행 함수.
 *
 * 사용법:
 * - 이 함수만 계속 실행하면 됩니다.
 * - 처음 실행하면 1번 파일부터 시작합니다.
 * - Apps Script 시간 제한/건수 제한으로 중간에 끊기면 진행 위치를 저장합니다.
 * - 같은 함수를 다시 실행하면 저장된 다음 파일부터 이어서 처리합니다.
 * - 전체 완료되면 진행 위치를 자동 초기화합니다.
 *
 * 기존 로그 시트를 삭제한 상태에서 실행하면,
 * 과거 진행 위치가 남아 있어도 자동으로 1번 파일부터 다시 시작합니다.
 */
function runBusinessRegistrationCopy() {
  return runBusinessRegistrationCopyInternal_({
    dryRun: false,
    propKey: BR_COPY_CFG.PROP_NEXT_FILE_INDEX,
    actionLabel: '사업자등록증 복사',
    nextFunctionName: 'runBusinessRegistrationCopy'
  });
}


/**
 * 사업자등록증 일괄 복사 미리보기 실행 함수.
 *
 * 사용법:
 * - 이 함수만 계속 실행하면 됩니다.
 * - 실제 복사/폴더 생성은 하지 않고, 매칭 예상 결과만 로그에 남깁니다.
 * - 중간에 끊기면 같은 함수를 다시 실행하면 다음 파일부터 이어서 미리보기합니다.
 */
function previewBusinessRegistrationCopyTargets() {
  return runBusinessRegistrationCopyInternal_({
    dryRun: true,
    propKey: BR_COPY_CFG.PROP_PREVIEW_NEXT_FILE_INDEX,
    actionLabel: '사업자등록증 미리보기',
    nextFunctionName: 'previewBusinessRegistrationCopyTargets'
  });
}


/**
 * 진행상황 초기화.
 * - 복사/미리보기 모두 처음부터 다시 돌리고 싶을 때만 실행합니다.
 * - 로그 시트를 삭제하고 다시 시작할 때도 이 함수 한 번 실행하면 안전합니다.
 */
function resetBusinessRegistrationCopyProgress() {
  PropertiesService.getScriptProperties().deleteProperty(BR_COPY_CFG.PROP_NEXT_FILE_INDEX);
  PropertiesService.getScriptProperties().deleteProperty(BR_COPY_CFG.PROP_PREVIEW_NEXT_FILE_INDEX);
  Logger.log('사업자등록증 복사/미리보기 진행상황을 초기화했습니다.');
  return {
    ok: true,
    message: '사업자등록증 복사/미리보기 진행상황 초기화 완료'
  };
}


/**
 * 설정/폴더/시트 탐지 점검용.
 */
function reportBusinessRegistrationCopyEnvironment() {
  const ss = getSpreadsheet_();
  const driveId = getSharedDriveId_();
  const sourceFolder = getSourceFolder_(driveId);
  const sourceFiles = listDirectNonFolderFilesPaged_(sourceFolder.id, driveId);

  const indexes = buildAllIndexes_({
    ss,
    driveId
  });

  const result = {
    sharedDriveId: driveId,
    sourceFolderName: sourceFolder.name,
    sourceFolderId: sourceFolder.id,
    sourceFileCount: sourceFiles.length,
    contractIndexCount: Object.keys(indexes.contractByContractNo).length,
    customerIndexCount: Object.keys(indexes.customerByCustomerNo).length,
    companyIndexCount: Object.keys(indexes.customerByCompanyNorm).length,
    rootCustomerFolderIndexCount: Object.keys(indexes.folderByCustomerNo).length
  };

  Logger.log(`공유드라이브 ID: ${result.sharedDriveId}`);
  Logger.log(`원본 폴더: ${result.sourceFolderName} / ${result.sourceFolderId}`);
  Logger.log(`원본 파일 수: ${result.sourceFileCount}`);
  Logger.log(`계약번호 인덱스: ${result.contractIndexCount}건`);
  Logger.log(`고객번호 인덱스: ${result.customerIndexCount}건`);
  Logger.log(`고객사명 인덱스: ${result.companyIndexCount}건`);
  Logger.log(`루트 고객사 폴더 인덱스: ${result.rootCustomerFolderIndexCount}건`);

  return result;
}


/**
 * 구버전 함수명 호환용.
 * 앞으로는 runBusinessRegistrationCopy()만 실행하면 됩니다.
 */
function manualCopyBusinessRegistrationFiles() {
  return runBusinessRegistrationCopy();
}


/**
 * 구버전 함수명 호환용.
 * 앞으로는 runBusinessRegistrationCopy()만 실행하면 됩니다.
 */
function continueCopyBusinessRegistrationFiles() {
  return runBusinessRegistrationCopy();
}


/**
 * 구버전 함수명 호환용.
 * 앞으로는 previewBusinessRegistrationCopyTargets()만 실행하면 됩니다.
 */
function continuePreviewBusinessRegistrationCopyTargets() {
  return previewBusinessRegistrationCopyTargets();
}


/**
 * 복사/미리보기 공통 실행 엔진.
 * dryRun=false: 실제 복사
 * dryRun=true : 미리보기만 로그 기록
 */
function runBusinessRegistrationCopyInternal_(options) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const startedAt = Date.now();

  try {
    const cfg = BR_COPY_CFG;
    const dryRun = !!options.dryRun;
    const propKey = options.propKey;
    const actionLabel = options.actionLabel;
    const nextFunctionName = options.nextFunctionName;

    const ss = getSpreadsheet_();
    const driveId = getSharedDriveId_();

    const sourceFolder = getSourceFolder_(driveId);
    const sourceFiles = listDirectNonFolderFilesPaged_(sourceFolder.id, driveId);

    if (sourceFiles.length === 0) {
      PropertiesService.getScriptProperties().deleteProperty(propKey);
      Logger.log(`"${cfg.SOURCE_FOLDER_NAME}" 폴더에 처리할 파일이 없습니다.`);
      return {
        ok: true,
        mode: dryRun ? 'PREVIEW' : 'COPY',
        status: 'NO_SOURCE_FILES',
        message: `"${cfg.SOURCE_FOLDER_NAME}" 폴더에 처리할 파일이 없습니다.`,
        sourceFileCount: 0
      };
    }

    const props = PropertiesService.getScriptProperties();

    // 기존 로그 시트를 삭제했는데 과거 진행 위치만 남아 있으면, 1번 파일부터 다시 시작합니다.
    if (isBrCopyLogEmpty_()) {
      props.deleteProperty(propKey);
    }

    let fileIndex = Number(props.getProperty(propKey) || 0);

    if (fileIndex < 0) fileIndex = 0;

    if (fileIndex >= sourceFiles.length) {
      props.deleteProperty(propKey);
      Logger.log(`${actionLabel}: 처리할 파일이 없습니다. 이미 완료되었습니다.`);
      return {
        ok: true,
        mode: dryRun ? 'PREVIEW' : 'COPY',
        status: 'ALREADY_DONE',
        message: '처리할 파일이 없습니다. 이미 완료되었습니다.',
        sourceFileCount: sourceFiles.length,
        nextFileIndex: ''
      };
    }

    const indexes = buildAllIndexes_({
      ss,
      driveId
    });

    const logs = [];

    let processed = 0;
    let copied = 0;
    let skipped = 0;
    let unmatched = 0;
    let createdFolders = 0;
    let errors = 0;
    let previewMatched = 0;
    let previewWouldCreateFolder = 0;

    const startIndex = fileIndex;

    while (
      fileIndex < sourceFiles.length &&
      processed < cfg.MAX_FILES_PER_RUN &&
      Date.now() - startedAt < cfg.MAX_MILLIS_PER_RUN
    ) {
      const file = sourceFiles[fileIndex];

      try {
        if (dryRun) {
          const resolvedTargets = resolveTargetsForBusinessRegistrationFile_({
            file,
            indexes,
            driveId,
            dryRun: true
          });

          if (!resolvedTargets || resolvedTargets.length === 0) {
            unmatched++;
            logs.push([
              new Date(),
              fileIndex + 1,
              file.name || '',
              '',
              '',
              '',
              '',
              '',
              '',
              'PREVIEW_ERROR',
              '대상 추출 실패'
            ]);
          } else {
            resolvedTargets.forEach(resolved => {
              if (resolved.targetFolder && resolved.targetFolder.id) {
                previewMatched++;
              } else if (resolved.wouldCreateFolder) {
                previewWouldCreateFolder++;
              } else {
                unmatched++;
              }

              logs.push([
                new Date(),
                fileIndex + 1,
                file.name || '',
                resolved.contractNo || '',
                resolved.customerNo || '',
                resolved.company || '',
                resolved.vendor || '',
                resolved.targetFolder ? resolved.targetFolder.name : '',
                resolved.targetFolder ? resolved.targetFolder.id : '',
                resolved.wouldCreateFolder ? 'PREVIEW_WOULD_CREATE_FOLDER' : 'PREVIEW',
                resolved.message || ''
              ]);
            });
          }

        } else {
          const results = copyOneBusinessRegistrationFile_({
            file,
            indexes,
            driveId
          });

          (results || []).forEach(result => {
            if (result.status === 'COPIED') {
              copied++;
            } else if (result.status === 'COPIED_TO_UNMATCHED') {
              copied++;
              unmatched++;
            } else if (
              result.status === 'SKIPPED_EXISTS' ||
              result.status === 'SKIPPED_DUPLICATE_TARGET'
            ) {
              skipped++;
            } else if (result.status === 'UNMATCHED') {
              unmatched++;
            } else if (result.status === 'ERROR') {
              errors++;
            } else {
              skipped++;
            }

            if (result.createdFolder) createdFolders++;

            logs.push([
              new Date(),
              fileIndex + 1,
              file.name || '',
              result.contractNo || '',
              result.customerNo || '',
              result.company || '',
              result.vendor || '',
              result.targetFolderName || '',
              result.targetFolderId || '',
              result.status || '',
              result.message || ''
            ]);
          });
        }

      } catch (err) {
        errors++;

        logs.push([
          new Date(),
          fileIndex + 1,
          file.name || '',
          '',
          '',
          '',
          '',
          '',
          '',
          dryRun ? 'PREVIEW_ERROR' : 'ERROR',
          err && err.message ? err.message : String(err)
        ]);
      }

      fileIndex++;
      processed++;
    }

    appendBrCopyLog_(logs);

    const elapsedMs = Date.now() - startedAt;
    const done = fileIndex >= sourceFiles.length;

    if (done) {
      props.deleteProperty(propKey);
    } else {
      props.setProperty(propKey, String(fileIndex));
    }

    const summary = {
      ok: errors === 0,
      mode: dryRun ? 'PREVIEW' : 'COPY',
      status: done ? 'DONE' : 'CONTINUE_NEEDED',
      actionLabel,
      sourceFileCount: sourceFiles.length,
      startFileNo: startIndex + 1,
      nextFileNo: done ? '' : fileIndex + 1,
      processed,
      copied,
      skipped,
      unmatched,
      createdFolders,
      previewMatched,
      previewWouldCreateFolder,
      errors,
      elapsedSeconds: Math.round(elapsedMs / 1000),
      message: ''
    };

    if (dryRun) {
      summary.message = done
        ? `미리보기 전체 완료: 처리 ${processed}개 / 기존폴더매칭 ${previewMatched}개 / 폴더생성예정 ${previewWouldCreateFolder}개 / 미매칭 ${unmatched}개 / 오류 ${errors}개`
        : `미리보기 이번 실행 완료: 처리 ${processed}개 / 기존폴더매칭 ${previewMatched}개 / 폴더생성예정 ${previewWouldCreateFolder}개 / 미매칭 ${unmatched}개 / 오류 ${errors}개. 아직 남았습니다. ${nextFunctionName}()를 다시 실행하세요. 다음 파일 순번: ${fileIndex + 1}`;
    } else {
      summary.message = done
        ? `사업자등록증 복사 전체 완료: 처리 ${processed}개 / 복사 ${copied}개 / 스킵 ${skipped}개 / 미매칭 ${unmatched}개 / 신규폴더 ${createdFolders}개 / 오류 ${errors}개`
        : `사업자등록증 복사 이번 실행 완료: 처리 ${processed}개 / 복사 ${copied}개 / 스킵 ${skipped}개 / 미매칭 ${unmatched}개 / 신규폴더 ${createdFolders}개 / 오류 ${errors}개. 아직 남았습니다. ${nextFunctionName}()를 다시 실행하세요. 다음 파일 순번: ${fileIndex + 1}`;
    }

    Logger.log(summary.message);
    return summary;

  } finally {
    lock.releaseLock();
  }
}


/**
 * 로그 시트를 삭제하거나 헤더만 남긴 상태면 true.
 * 이 경우 기존 ScriptProperties에 남아 있던 next index를 자동 초기화합니다.
 */
function isBrCopyLogEmpty_() {
  const ss = getSpreadsheet_();
  const sheet = ss.getSheetByName(BR_COPY_CFG.LOG_SHEET_NAME);

  if (!sheet) return true;

  return sheet.getLastRow() <= 1;
}


/***** 핵심 처리 *****/

function copyOneBusinessRegistrationFile_(params) {
  const file = params.file;
  const indexes = params.indexes;
  const driveId = params.driveId;

  const resolvedTargets = resolveTargetsForBusinessRegistrationFile_({
    file,
    indexes,
    driveId,
    dryRun: false
  });

  if (!resolvedTargets || resolvedTargets.length === 0) {
    return [{
      status: 'UNMATCHED',
      message: '대상 추출 실패',
      contractNo: '',
      customerNo: '',
      company: '',
      vendor: '',
      targetFolderName: '',
      targetFolderId: ''
    }];
  }

  const results = [];
  const copiedTargetKeys = {};

  resolvedTargets.forEach(resolved => {
    try {
      if (!resolved.targetFolder || !resolved.targetFolder.id) {
        const unmatchedFolder = ensureUnmatchedFolder_(driveId);
        const unmatchedName = buildCopyFileNameForTarget_(file.name, resolved);

        const exists = findDirectFileByName_(unmatchedFolder.id, driveId, unmatchedName);
        if (exists && !BR_COPY_CFG.REPLACE_SAME_NAME_FILE) {
          results.push({
            status: 'SKIPPED_EXISTS',
            message: `미매칭 폴더에 같은 이름 파일이 이미 있음 / ${resolved.message || ''}`,
            contractNo: resolved.contractNo || '',
            customerNo: resolved.customerNo || '',
            company: resolved.company || '',
            vendor: resolved.vendor || '',
            targetFolderName: unmatchedFolder.name,
            targetFolderId: unmatchedFolder.id,
            copiedFileName: unmatchedName
          });
          return;
        }

        if (exists && BR_COPY_CFG.REPLACE_SAME_NAME_FILE) {
          trashDriveFile_(exists.id);
        }

        copyDriveFileToFolder_(file.id, unmatchedName, unmatchedFolder.id);

        results.push({
          status: 'COPIED_TO_UNMATCHED',
          message: resolved.message || '대상 고객사 폴더 매칭 실패',
          contractNo: resolved.contractNo || '',
          customerNo: resolved.customerNo || '',
          company: resolved.company || '',
          vendor: resolved.vendor || '',
          targetFolderName: unmatchedFolder.name,
          targetFolderId: unmatchedFolder.id,
          copiedFileName: unmatchedName
        });
        return;
      }

      const copyName = buildCopyFileNameForTarget_(file.name, resolved);
      const targetKey = resolved.targetFolder.id + '||' + copyName;

      if (copiedTargetKeys[targetKey]) {
        results.push({
          status: 'SKIPPED_DUPLICATE_TARGET',
          message: '같은 원본 파일 내에서 동일 대상 폴더/파일명 중복이라 스킵',
          contractNo: resolved.contractNo || '',
          customerNo: resolved.customerNo || '',
          company: resolved.company || '',
          vendor: resolved.vendor || '',
          targetFolderName: resolved.targetFolder.name,
          targetFolderId: resolved.targetFolder.id,
          copiedFileName: copyName,
          createdFolder: resolved.createdFolder || false
        });
        return;
      }

      copiedTargetKeys[targetKey] = true;

      const exists = findDirectFileByName_(resolved.targetFolder.id, driveId, copyName);

      if (exists && !BR_COPY_CFG.REPLACE_SAME_NAME_FILE) {
        results.push({
          status: 'SKIPPED_EXISTS',
          message: '대상 고객사 폴더에 같은 이름 파일이 이미 있어 스킵',
          contractNo: resolved.contractNo || '',
          customerNo: resolved.customerNo || '',
          company: resolved.company || '',
          vendor: resolved.vendor || '',
          targetFolderName: resolved.targetFolder.name,
          targetFolderId: resolved.targetFolder.id,
          copiedFileName: copyName,
          createdFolder: resolved.createdFolder || false
        });
        return;
      }

      if (exists && BR_COPY_CFG.REPLACE_SAME_NAME_FILE) {
        trashDriveFile_(exists.id);
      }

      copyDriveFileToFolder_(file.id, copyName, resolved.targetFolder.id);

      results.push({
        status: 'COPIED',
        message: resolved.message || '복사 완료',
        contractNo: resolved.contractNo || '',
        customerNo: resolved.customerNo || '',
        company: resolved.company || '',
        vendor: resolved.vendor || '',
        targetFolderName: resolved.targetFolder.name,
        targetFolderId: resolved.targetFolder.id,
        copiedFileName: copyName,
        createdFolder: resolved.createdFolder || false
      });

    } catch (err) {
      results.push({
        status: 'ERROR',
        message: err && err.message ? err.message : String(err),
        contractNo: resolved.contractNo || '',
        customerNo: resolved.customerNo || '',
        company: resolved.company || '',
        vendor: resolved.vendor || '',
        targetFolderName: resolved.targetFolder ? resolved.targetFolder.name : '',
        targetFolderId: resolved.targetFolder ? resolved.targetFolder.id : ''
      });
    }
  });

  return results;
}

function resolveTargetForBusinessRegistrationFile_(params) {
  const targets = resolveTargetsForBusinessRegistrationFile_(params);
  return targets && targets.length ? targets[0] : {
    message: '대상 추출 실패'
  };
}


/**
 * 파일 1개에서 복사 대상 여러 개를 반환.
 * 파일명 앞에 계약번호가 여러 개 있으면 계약번호별로 각각 대상 고객사 폴더를 만든다.
 *
 * 예:
 * - 123_124_한국방송공사 사업자등록증.pdf
 * - 123. 124. 한국방송공사 사업자등록증.pdf
 *
 * 위 경우 123, 124를 각각 계약번호로 보고
 * 수주확정/계약완료 시트에서 각 계약번호의 고객번호를 찾아
 * 각 고객사 폴더에 같은 사업자등록증을 복사한다.
 */
function resolveTargetsForBusinessRegistrationFile_(params) {
  const file = params.file;
  const indexes = params.indexes;
  const driveId = params.driveId;
  const dryRun = !!params.dryRun;

  const parsed = parseBusinessRegistrationFileName_(file.name);
  const contractNos = parsed.contractNos || (parsed.contractNo ? [parsed.contractNo] : []);

  if (contractNos.length > 0) {
    const seenContractNos = {};
    const targets = [];

    contractNos.forEach(contractNo => {
      const key = normalizeKey_(contractNo);
      if (!key || seenContractNos[key]) return;

      seenContractNos[key] = true;

      targets.push(resolveTargetByContractNo_({
        contractNo,
        parsed,
        indexes,
        driveId,
        dryRun
      }));
    });

    return targets;
  }

  return [
    resolveTargetByCompanyName_({
      parsed,
      indexes,
      driveId,
      dryRun
    })
  ];
}


function resolveTargetByContractNo_(params) {
  const contractNo = cleanValue_(params.contractNo);
  const parsed = params.parsed;
  const indexes = params.indexes;
  const driveId = params.driveId;
  const dryRun = !!params.dryRun;

  const contractKey = normalizeKey_(contractNo);
  const contract = indexes.contractByContractNo[contractKey];

  if (!contract) {
    return {
      contractNo,
      company: parsed.companyNameGuess,
      message: `파일명 앞 계약번호 ${contractNo}를 수주확정/계약완료 시트에서 찾지 못함`
    };
  }

  const customerNo = cleanValue_(contract.customerNo);
  const company = cleanValue_(contract.company || parsed.companyNameGuess);
  const vendor = cleanValue_(contract.vendor || BR_COPY_CFG.EMPTY_VENDOR_TEXT);

  if (!customerNo) {
    return {
      contractNo,
      company,
      vendor,
      message: `계약번호 ${contractNo} 행에서 고객번호가 비어 있음`
    };
  }

  const ensured = ensureCustomerFolderForBusinessReg_({
    driveId,
    indexes,
    customerNo,
    company,
    vendor,
    dryRun
  });

  return {
    contractNo,
    customerNo,
    company,
    vendor,
    targetFolder: ensured.folder,
    createdFolder: ensured.createdFolder,
    wouldCreateFolder: ensured.wouldCreateFolder || false,
    multiContractSource: (parsed.contractNos || []).length > 1,
    originalContractNos: parsed.contractNos || [contractNo],
    message: ensured.wouldCreateFolder
      ? `계약번호 ${contractNo} → 고객번호 ${customerNo}, 고객사 폴더 신규 생성 예정`
      : ensured.createdFolder
      ? `계약번호 ${contractNo} → 고객번호 ${customerNo}, 고객사 폴더 신규 생성 후 대상 지정`
      : `계약번호 ${contractNo} → 고객번호 ${customerNo}, 기존 고객사 폴더 대상 지정`
  };
}


function resolveTargetByCompanyName_(params) {
  const parsed = params.parsed;
  const indexes = params.indexes;
  const driveId = params.driveId;
  const dryRun = !!params.dryRun;

  const companyGuess = parsed.companyNameGuess;
  const companyNorm = normalizeCompanyName_(companyGuess);

  if (!companyNorm) {
    return {
      company: companyGuess,
      message: '파일명에서 고객사명을 추출하지 못함'
    };
  }

  // 1. 기존 고객사 폴더명과 매칭
  const folderMatch = findBestFolderByCompanyName_(companyNorm, indexes);

  if (folderMatch && folderMatch.folder) {
    const customerNo = folderMatch.customerNo || '';
    const customerInfo = customerNo ? indexes.customerByCustomerNo[normalizeKey_(customerNo)] : null;

    return {
      company: companyGuess,
      customerNo,
      vendor: customerInfo ? customerInfo.vendor : '',
      targetFolder: folderMatch.folder,
      message: `계약번호 없음. 파일명 고객사명 기준 기존 폴더 매칭 / 점수 ${folderMatch.score}`
    };
  }

  // 2. 시트의 회사명과 매칭해서 고객번호 확보
  const customerMatch = findBestCustomerByCompanyName_(companyNorm, indexes);

  if (customerMatch && customerMatch.customer && customerMatch.customer.customerNo) {
    const customer = customerMatch.customer;

    const ensured = ensureCustomerFolderForBusinessReg_({
      driveId,
      indexes,
      customerNo: customer.customerNo,
      company: customer.company || companyGuess,
      vendor: customer.vendor || BR_COPY_CFG.EMPTY_VENDOR_TEXT,
      dryRun
    });

    return {
      company: customer.company || companyGuess,
      customerNo: customer.customerNo,
      vendor: customer.vendor || '',
      targetFolder: ensured.folder,
      createdFolder: ensured.createdFolder,
      wouldCreateFolder: ensured.wouldCreateFolder || false,
      message: ensured.wouldCreateFolder
        ? `계약번호 없음. 회사명으로 시트 고객번호 ${customer.customerNo} 매칭 후 폴더 신규 생성 예정`
        : ensured.createdFolder
        ? `계약번호 없음. 회사명으로 시트 고객번호 ${customer.customerNo} 매칭 후 폴더 신규 생성`
        : `계약번호 없음. 회사명으로 시트 고객번호 ${customer.customerNo} 매칭 후 기존 폴더 대상 지정`
    };
  }

  // 3. 고객번호까지 못 찾은 경우
  if (BR_COPY_CFG.CREATE_COMPANY_FOLDER_WITHOUT_CUSTOMER_NO) {
    const folderName = buildCustomerFolderNameForBusinessReg_(
      BR_COPY_CFG.UNKNOWN_CUSTOMER_NO_TEXT,
      companyGuess,
      BR_COPY_CFG.EMPTY_VENDOR_TEXT
    );

    if (dryRun) {
      return {
        company: companyGuess,
        targetFolder: {
          id: '',
          name: folderName,
          webViewLink: ''
        },
        createdFolder: false,
        wouldCreateFolder: true,
        message: '고객번호 미확인. 회사명 폴더 신규 생성 예정'
      };
    }

    const folder = createDriveFolder_(folderName, driveId);

    return {
      company: companyGuess,
      targetFolder: folder,
      createdFolder: true,
      message: '고객번호 미확인. 회사명 폴더 신규 생성'
    };
  }

  return {
    company: companyGuess,
    message: '계약번호 없음. 파일명 고객사명으로 기존 폴더/시트 고객번호를 찾지 못해 미매칭 처리'
  };
}

function ensureCustomerFolderForBusinessReg_(params) {
  const driveId = params.driveId;
  const indexes = params.indexes;
  const customerNo = cleanValue_(params.customerNo);
  const company = cleanValue_(params.company);
  const vendor = cleanValue_(params.vendor) || BR_COPY_CFG.EMPTY_VENDOR_TEXT;
  const dryRun = !!params.dryRun;

  const customerNoKey = normalizeKey_(customerNo);

  if (indexes.folderByCustomerNo[customerNoKey]) {
    return {
      folder: indexes.folderByCustomerNo[customerNoKey].folder,
      createdFolder: false
    };
  }

  const folderName = buildCustomerFolderNameForBusinessReg_(customerNo, company, vendor);

  if (dryRun) {
    return {
      folder: {
        id: '',
        name: folderName,
        webViewLink: ''
      },
      createdFolder: false,
      wouldCreateFolder: true
    };
  }

  const folder = createDriveFolder_(folderName, driveId);

  const folderInfo = {
    folder,
    customerNo,
    company,
    companyNorm: normalizeCompanyName_(company),
    vendor
  };

  indexes.folderByCustomerNo[customerNoKey] = folderInfo;
  addToArrayMap_(indexes.folderByCompanyNorm, folderInfo.companyNorm, folderInfo);

  return {
    folder,
    createdFolder: true
  };
}


/***** 인덱스 생성 *****/

function buildAllIndexes_(params) {
  const ss = params.ss;
  const driveId = params.driveId;

  const customerIndex = buildCustomerIndexFromSheets_(ss);
  const contractIndex = buildContractIndexFromSheets_(ss, customerIndex);
  const folderIndex = buildCustomerFolderIndex_(driveId);

  return {
    contractByContractNo: contractIndex.contractByContractNo,
    customerByCustomerNo: customerIndex.customerByCustomerNo,
    customerByCompanyNorm: customerIndex.customerByCompanyNorm,
    folderByCustomerNo: folderIndex.folderByCustomerNo,
    folderByCompanyNorm: folderIndex.folderByCompanyNorm
  };
}


function buildContractIndexFromSheets_(ss, customerIndex) {
  const contractByContractNo = {};

  BR_COPY_CFG.CONTRACT_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const detected = detectHeaderRowAndMap_(sheet, {
      contractNo: HEADER_CANDIDATES_.contractNo,
      customerNo: HEADER_CANDIDATES_.customerNo
    });

    if (!detected) return;

    const values = sheet
      .getRange(detected.headerRow + 1, 1, Math.max(sheet.getLastRow() - detected.headerRow, 0), sheet.getLastColumn())
      .getDisplayValues();

    values.forEach(row => {
      const contractNo = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.contractNo);
      const customerNo = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.customerNo);
      const company = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.company);
      const vendor = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.vendor);

      const contractKey = normalizeKey_(contractNo);
      if (!contractKey) return;

      const customerKey = normalizeKey_(customerNo);
      const supplement = customerKey ? customerIndex.customerByCustomerNo[customerKey] : null;

      contractByContractNo[contractKey] = {
        contractNo: cleanValue_(contractNo),
        customerNo: cleanValue_(customerNo || (supplement ? supplement.customerNo : '')),
        company: cleanValue_(company || (supplement ? supplement.company : '')),
        vendor: cleanValue_(vendor || (supplement ? supplement.vendor : '')),
        sourceSheet: sheetName
      };
    });
  });

  return {
    contractByContractNo
  };
}


function buildCustomerIndexFromSheets_(ss) {
  const customerByCustomerNo = {};
  const customerByCompanyNorm = {};

  // 1. 마스터시트 우선
  BR_COPY_CFG.MASTER_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const detected = detectHeaderRowAndMap_(sheet, {
      customerNo: HEADER_CANDIDATES_.customerNo,
      company: HEADER_CANDIDATES_.company
    });

    if (!detected) return;

    addCustomersFromSheet_(sheet, detected, customerByCustomerNo, customerByCompanyNorm, sheetName);
  });

  // 2. 계약 관련 시트도 고객 인덱스 보강
  BR_COPY_CFG.CONTRACT_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const detected = detectHeaderRowAndMap_(sheet, {
      customerNo: HEADER_CANDIDATES_.customerNo,
      company: HEADER_CANDIDATES_.company
    });

    if (!detected) return;

    addCustomersFromSheet_(sheet, detected, customerByCustomerNo, customerByCompanyNorm, sheetName);
  });

  return {
    customerByCustomerNo,
    customerByCompanyNorm
  };
}


function addCustomersFromSheet_(sheet, detected, customerByCustomerNo, customerByCompanyNorm, sheetName) {
  const rowCount = Math.max(sheet.getLastRow() - detected.headerRow, 0);
  if (rowCount <= 0) return;

  const values = sheet
    .getRange(detected.headerRow + 1, 1, rowCount, sheet.getLastColumn())
    .getDisplayValues();

  values.forEach(row => {
    const customerNo = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.customerNo);
    const company = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.company);
    const vendor = getFirstByCandidatesFromRow_(row, detected.headerMap, HEADER_CANDIDATES_.vendor);

    const customerKey = normalizeKey_(customerNo);
    const companyNorm = normalizeCompanyName_(company);

    if (!customerKey || !companyNorm) return;

    const customer = {
      customerNo: cleanValue_(customerNo),
      company: cleanValue_(company),
      companyNorm,
      vendor: cleanValue_(vendor || BR_COPY_CFG.EMPTY_VENDOR_TEXT),
      sourceSheet: sheetName
    };

    if (!customerByCustomerNo[customerKey]) {
      customerByCustomerNo[customerKey] = customer;
    }

    addToArrayMap_(customerByCompanyNorm, companyNorm, customer);
  });
}


function buildCustomerFolderIndex_(driveId) {
  const folderByCustomerNo = {};
  const folderByCompanyNorm = {};

  const rootFolders = listDirectChildFoldersPaged_(driveId, driveId);

  rootFolders.forEach(folder => {
    const name = cleanValue_(folder.name);

    // 업무용 시스템 폴더는 고객사 폴더 인덱스에서 제외
    if (
      name === BR_COPY_CFG.SOURCE_FOLDER_NAME ||
      name === BR_COPY_CFG.UNMATCHED_FOLDER_NAME ||
      name === '수주실패'
    ) {
      return;
    }

    const parsed = parseCustomerFolderName_(name);
    if (!parsed.customerNo && !parsed.companyNorm) return;

    const info = {
      folder,
      customerNo: parsed.customerNo,
      company: parsed.company,
      companyNorm: parsed.companyNorm,
      vendor: parsed.vendor
    };

    if (parsed.customerNo) {
      folderByCustomerNo[normalizeKey_(parsed.customerNo)] = info;
    }

    if (parsed.companyNorm) {
      addToArrayMap_(folderByCompanyNorm, parsed.companyNorm, info);
    }
  });

  return {
    folderByCustomerNo,
    folderByCompanyNorm
  };
}


/***** 파일명/폴더명 파싱 *****/

function parseBusinessRegistrationFileName_(fileName) {
  const originalName = cleanValue_(fileName);
  const baseName = removeExtension_(originalName);

  const extracted = extractLeadingContractNosAndRest_(baseName);
  const contractNos = extracted.contractNos || [];
  const rest = extracted.rest || baseName;

  const companyNameGuess = extractCompanyNameGuessFromBusinessRegFile_(rest || baseName);

  return {
    originalName,
    baseName,
    contractNo: contractNos.length ? contractNos[0] : '',
    contractNos,
    companyNameGuess
  };
}


/**
 * 파일명 앞머리의 계약번호 여러 개를 추출.
 *
 * 지원 예:
 * - 123_124_한국방송공사 사업자등록증
 * - 123. 124. 한국방송공사 사업자등록증
 * - 123 124 한국방송공사 사업자등록증
 * - 123-124-한국방송공사 사업자등록증
 * - 123) 124) 한국방송공사 사업자등록증
 *
 * 주의:
 * - 여기서 추출하는 번호는 계약번호다.
 * - 고객사 폴더명 앞의 고객번호와 절대 혼동하지 않는다.
 */
function extractLeadingContractNosAndRest_(baseName) {
  let rest = cleanValue_(baseName);
  const contractNos = [];

  while (true) {
    const before = rest;

    // 숫자 1~6자리 + 구분자.
    // 구분자는 _, ., 공백, -, ), ], 번 등을 허용.
    const m = rest.match(/^\s*(\d{1,6})\s*(?:[_.,·ㆍ\-번\)\]]+|\s+)/);

    if (!m || !m[1]) {
      break;
    }

    contractNos.push(m[1]);
    rest = rest.slice(m[0].length).trim();

    if (!rest || rest === before) {
      break;
    }

    // 다음 글자가 숫자로 시작하지 않으면 계약번호 prefix 추출 종료.
    if (!/^\d/.test(rest)) {
      break;
    }
  }

  return {
    contractNos,
    rest
  };
}

function extractCompanyNameGuessFromBusinessRegFile_(text) {
  let s = cleanValue_(text);

  s = s
    .replace(/사업자\s*등록증/gi, ' ')
    .replace(/사업자등록증/gi, ' ')
    .replace(/등록증/gi, ' ')
    .replace(/사본/gi, ' ')
    .replace(/최신/gi, ' ')
    .replace(/본사/gi, ' ')
    .replace(/대표자/gi, ' ')
    .replace(/등본/gi, ' ')
    .replace(/\[[^\]]*\]/g, ' ')
    .replace(/\([^)]*(도로명|지번|주소|대표|사업자번호|등록번호)[^)]*\)/g, ' ')
    .replace(/[{}]/g, ' ')
    .replace(/[_\-]+/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();

  // 너무 길면 앞쪽 고객사명 부분만 우선 사용.
  // 단, (주), 주식회사 등은 normalizeCompanyName_에서 처리.
  return s;
}


function parseCustomerFolderName_(folderName) {
  const s = cleanValue_(folderName);

  // 현재: 고객번호_고객사명_수행사
  // 향후: C1_고객사명_수행사 도 허용
  const parts = s.split('_').map(p => cleanValue_(p)).filter(Boolean);

  let customerNo = '';
  let company = '';
  let vendor = '';

  if (parts.length >= 3) {
    customerNo = parts[0];
    vendor = parts[parts.length - 1];
    company = parts.slice(1, parts.length - 1).join('_');
  } else if (parts.length === 2) {
    customerNo = parts[0];
    company = parts[1];
  } else {
    company = s;
  }

  return {
    customerNo,
    company,
    companyNorm: normalizeCompanyName_(company),
    vendor
  };
}


function buildCustomerFolderNameForBusinessReg_(customerNo, company, vendor) {
  const parts = [
    sanitizeFolderPart_(customerNo),
    sanitizeFolderPart_(company),
    sanitizeFolderPart_(vendor || BR_COPY_CFG.EMPTY_VENDOR_TEXT)
  ];

  let name = parts.join('_').replace(/_+/g, '_').trim();

  if (name.length > 180) {
    name = name.slice(0, 180).trim();
  }

  return name;
}


/**
 * 복사 파일명 생성.
 * 기본은 원본 파일명을 유지한다.
 *
 * 동일 사업자등록증을 여러 계약번호/여러 고객사 폴더에 복사하는 경우에도
 * 각 고객사 폴더가 다르므로 원본명을 유지해도 충돌은 거의 없다.
 * 필요하면 여기서 계약번호별 파일명으로 바꿀 수 있다.
 */
function buildCopyFileNameForTarget_(originalFileName, resolved) {
  return cleanValue_(originalFileName);
}


/***** 매칭 *****/

function findBestFolderByCompanyName_(companyNorm, indexes) {
  return findBestCompanyMatchFromArrayMap_(companyNorm, indexes.folderByCompanyNorm, 'folder');
}


function findBestCustomerByCompanyName_(companyNorm, indexes) {
  return findBestCompanyMatchFromArrayMap_(companyNorm, indexes.customerByCompanyNorm, 'customer');
}


function findBestCompanyMatchFromArrayMap_(companyNorm, map, returnType) {
  const target = cleanValue_(companyNorm);
  if (!target) return null;

  let best = null;
  let second = null;

  Object.keys(map || {}).forEach(key => {
    const score = calcCompanyMatchScore_(target, key);
    if (score <= 0) return;

    const arr = map[key] || [];
    arr.forEach(item => {
      const candidate = Object.assign({}, item, { score });

      if (!best || candidate.score > best.score) {
        second = best;
        best = candidate;
      } else if (!second || candidate.score > second.score) {
        second = candidate;
      }
    });
  });

  if (!best || best.score < 70) {
    return null;
  }

  // 애매한 동점이면 오매칭 방지
  if (second && best.score === second.score && best.score < 100) {
    return null;
  }

  if (returnType === 'folder') {
    return {
      folder: best.folder,
      customerNo: best.customerNo || '',
      company: best.company || '',
      vendor: best.vendor || '',
      score: best.score
    };
  }

  return {
    customer: {
      customerNo: best.customerNo || '',
      company: best.company || '',
      vendor: best.vendor || ''
    },
    score: best.score
  };
}


function calcCompanyMatchScore_(a, b) {
  const x = cleanValue_(a);
  const y = cleanValue_(b);

  if (!x || !y) return 0;
  if (x === y) return 100;

  const minLen = Math.min(x.length, y.length);
  const maxLen = Math.max(x.length, y.length);

  if (minLen < 2) return 0;

  if (x.indexOf(y) !== -1 || y.indexOf(x) !== -1) {
    return Math.round(80 + (minLen / maxLen) * 15);
  }

  // 간단한 포함 토큰 점수
  let common = 0;
  for (let i = 0; i < minLen; i++) {
    if (x.indexOf(y[i]) !== -1) common++;
  }

  const ratio = common / maxLen;
  return ratio >= 0.75 ? Math.round(ratio * 70) : 0;
}


/***** 헤더 탐지 *****/

const HEADER_CANDIDATES_ = {
  contractNo: [
    '계약번호',
    '계약 번호',
    '계약NO',
    '계약 No',
    '계약No',
    '계약 No.',
    '계약고유번호',
    '수주번호',
    '수주 번호'
  ],
  customerNo: [
    '고객번호',
    '고객 번호',
    '고객NO',
    '고객 No',
    '고객No',
    '고객 No.',
    '고객ID',
    '고객 ID'
  ],
  company: [
    '회사명',
    '고객사명',
    '고객사',
    '상호',
    '업체명',
    '건물명',
    '사업장명'
  ],
  vendor: [
    '수행사',
    '최종수행사',
    '수행 업체',
    '수행업체',
    '협력사'
  ]
};


function detectHeaderRowAndMap_(sheet, requiredGroups) {
  const maxRows = Math.min(BR_COPY_CFG.HEADER_SCAN_ROWS, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();

  if (maxRows <= 0 || lastCol <= 0) return null;

  const rows = sheet.getRange(1, 1, maxRows, lastCol).getDisplayValues();

  for (let r = 0; r < rows.length; r++) {
    const headerMap = {};
    rows[r].forEach((h, i) => {
      const key = normalizeHeader_(h);
      if (key && !headerMap[key]) {
        headerMap[key] = i + 1;
      }
    });

    let ok = true;

    Object.keys(requiredGroups || {}).forEach(groupName => {
      const candidates = requiredGroups[groupName] || [];
      const found = candidates.some(name => !!headerMap[normalizeHeader_(name)]);
      if (!found) ok = false;
    });

    if (ok) {
      return {
        headerRow: r + 1,
        headerMap
      };
    }
  }

  return null;
}


function getFirstByCandidatesFromRow_(row, headerMap, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const col = headerMap[normalizeHeader_(candidates[i])];
    if (col) {
      const value = cleanValue_(row[col - 1]);
      if (value) return value;
    }
  }

  return '';
}


/***** Google Drive API *****/

function getSharedDriveId_() {
  if (cleanValue_(BR_COPY_CFG.SHARED_DRIVE_ID)) {
    return cleanValue_(BR_COPY_CFG.SHARED_DRIVE_ID);
  }

  const props = PropertiesService.getScriptProperties();
  const cacheKey = 'S1_BR_COPY_SHARED_DRIVE_ID';
  const cached = props.getProperty(cacheKey);

  if (cached) return cached;

  const q = `name = ${driveQueryString_(BR_COPY_CFG.SHARED_DRIVE_NAME)}`;

  const data = driveFetch_(
    'drives?pageSize=10&q=' + encodeURIComponent(q) + '&fields=drives(id,name)',
    { method: 'get' }
  );

  const drives = data.drives || [];

  if (drives.length === 0) {
    throw new Error(
      `공유드라이브를 찾지 못했습니다: ${BR_COPY_CFG.SHARED_DRIVE_NAME}. ` +
      `BR_COPY_CFG.SHARED_DRIVE_ID에 공유드라이브 ID를 직접 입력하세요.`
    );
  }

  const driveId = drives[0].id;
  props.setProperty(cacheKey, driveId);

  return driveId;
}


function getSourceFolder_(driveId) {
  const folder = findChildFolder_(driveId, driveId, BR_COPY_CFG.SOURCE_FOLDER_NAME);

  if (!folder) {
    throw new Error(`공유드라이브 루트에서 원본 폴더를 찾지 못했습니다: ${BR_COPY_CFG.SOURCE_FOLDER_NAME}`);
  }

  return folder;
}


function ensureUnmatchedFolder_(driveId) {
  const existing = findChildFolder_(driveId, driveId, BR_COPY_CFG.UNMATCHED_FOLDER_NAME);
  if (existing) return existing;

  return createDriveFolder_(BR_COPY_CFG.UNMATCHED_FOLDER_NAME, driveId);
}


function findChildFolder_(parentId, driveId, folderName) {
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
    '&fields=files(id,name,webViewLink,trashed,parents)';

  const data = driveFetch_(path, { method: 'get' });
  const files = data.files || [];

  return files.length ? files[0] : null;
}


function createDriveFolder_(folderName, parentId) {
  return driveFetch_(
    'files?supportsAllDrives=true&fields=id,name,webViewLink,parents',
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


function listDirectNonFolderFilesPaged_(parentFolderId, driveId) {
  const q = [
    `${driveQueryString_(parentFolderId)} in parents`,
    `mimeType != 'application/vnd.google-apps.folder'`,
    `trashed = false`
  ].join(' and ');

  let pageToken = '';
  const files = [];

  do {
    let path =
      'files' +
      '?supportsAllDrives=true' +
      '&includeItemsFromAllDrives=true' +
      '&corpora=drive' +
      '&driveId=' + encodeURIComponent(driveId) +
      '&pageSize=1000' +
      '&q=' + encodeURIComponent(q) +
      '&fields=nextPageToken,files(id,name,mimeType,webViewLink,trashed,parents,modifiedTime,size)';

    if (pageToken) {
      path += '&pageToken=' + encodeURIComponent(pageToken);
    }

    const data = driveFetch_(path, { method: 'get' });

    (data.files || []).forEach(file => files.push(file));
    pageToken = data.nextPageToken || '';

  } while (pageToken);

  files.sort((a, b) => cleanValue_(a.name).localeCompare(cleanValue_(b.name), 'ko'));

  return files;
}


function findDirectFileByName_(parentFolderId, driveId, fileName) {
  const q = [
    `${driveQueryString_(parentFolderId)} in parents`,
    `mimeType != 'application/vnd.google-apps.folder'`,
    `name = ${driveQueryString_(fileName)}`,
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
    '&fields=files(id,name,webViewLink,trashed,parents)';

  const data = driveFetch_(path, { method: 'get' });
  const files = data.files || [];

  return files.length ? files[0] : null;
}


function copyDriveFileToFolder_(sourceFileId, newName, targetFolderId) {
  return driveFetch_(
    'files/' + encodeURIComponent(sourceFileId) + '/copy?supportsAllDrives=true&fields=id,name,webViewLink,parents',
    {
      method: 'post',
      payload: {
        name: newName,
        parents: [targetFolderId]
      }
    }
  );
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


function driveFetch_(path, options) {
  const url = driveV2CompatBuildUrl_(path, false);

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
    params.payload = JSON.stringify(driveV2CompatPreparePayload_(params.payload));
  }

  const res = UrlFetchApp.fetch(url, params);
  const code = res.getResponseCode();
  const text = res.getContentText();

  if (code < 200 || code >= 300) {
    throw new Error(`Drive API 오류 ${code}: ${text}`);
  }

  return text ? driveV2CompatNormalizeResponse_(JSON.parse(text)) : {};
}


function driveQueryString_(value) {
  const s = cleanValue_(value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");

  return `'${s}'`;
}


/***** 스프레드시트/로그 *****/

function getSpreadsheet_() {
  const id = cleanValue_(BR_COPY_CFG.SPREADSHEET_ID);

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


function appendBrCopyLog_(rows) {
  if (!rows || rows.length === 0) return;

  const ss = getSpreadsheet_();
  const name = BR_COPY_CFG.LOG_SHEET_NAME;

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '일시',
      '파일순번',
      '원본파일명',
      '계약번호',
      '고객번호',
      '회사명',
      '수행사',
      '대상폴더명',
      '대상폴더ID',
      '처리결과',
      '메시지'
    ]);
    sheet.setFrozenRows(1);
  }

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


/***** 유틸 *****/

function addToArrayMap_(map, key, item) {
  const k = cleanValue_(key);
  if (!k) return;

  if (!map[k]) map[k] = [];
  map[k].push(item);
}


function cleanValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}


function normalizeHeader_(value) {
  return cleanValue_(value)
    .replace(/\s+/g, '')
    .replace(/[._\-\/]/g, '')
    .toUpperCase();
}


function normalizeKey_(value) {
  let s = cleanValue_(value);

  if (!s) return '';

  s = s.replace(/\.0$/, '').trim();

  // 순수 숫자는 0001 / 1 비교 가능하게 앞 0 제거
  if (/^\d+$/.test(s)) {
    s = String(Number(s));
  }

  return s.toUpperCase();
}


function normalizeCompanyName_(value) {
  let s = cleanValue_(value);

  if (!s) return '';

  s = removeExtension_(s);

  s = s
    .replace(/주식회사/gi, '')
    .replace(/\(주\)/gi, '')
    .replace(/㈜/gi, '')
    .replace(/주\)/gi, '')
    .replace(/\(유\)/gi, '')
    .replace(/유한회사/gi, '')
    .replace(/재단법인/gi, '')
    .replace(/사단법인/gi, '')
    .replace(/의료법인/gi, '')
    .replace(/학교법인/gi, '')
    .replace(/사업자\s*등록증/gi, '')
    .replace(/사업자등록증/gi, '')
    .replace(/등록증/gi, '')
    .replace(/사본/gi, '')
    .replace(/[0-9]+$/g, '')
    .replace(/[^가-힣a-zA-Z0-9]/g, '')
    .toUpperCase()
    .trim();

  return s;
}


function sanitizeFolderPart_(value) {
  return cleanValue_(value)
    .replace(/[\/\\:*?"<>|#\[\]\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function removeExtension_(fileName) {
  return cleanValue_(fileName).replace(/\.[^.]+$/, '');
}
