/*******************************************************
 * 계약서 일괄 복사.gs
 * v2 단일 실행 함수 구조
 *
 * 목적:
 * - 공유드라이브 "S1 고객사 파일 관리" 안의 "계약서 일괄" 폴더에 있는
 *   계약서 파일을 상위 디렉토리의 고객사별 폴더로 복사한다.
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
 * 실행 함수:
 * - runContractFileCopy() : 실제 복사. 이 함수만 반복 실행하면 이어서 처리됨.
 * - previewContractFileCopyTargets() : 미리보기. 이 함수만 반복 실행하면 이어서 처리됨.
 * - resetContractFileCopyProgress() : 진행 초기화.
 *
 * 주의:
 * - 계약서 파일명 앞 번호 = 계약번호
 * - 고객사 폴더명 앞 번호 = 고객번호
 * - 두 번호를 절대 같은 번호로 보지 않음
 *******************************************************/


/***** 설정 *****/

const CONTRACT_COPY_CFG = {
  // 마스터시트에 바인딩된 Apps Script면 빈 값 유지.
  // 독립형 Apps Script면 스프레드시트 ID 입력.
  SPREADSHEET_ID: '',

  // 원본 공유드라이브명
  SHARED_DRIVE_NAME: 'S1 고객사 파일 관리',

  // 가능하면 비워둬도 됨.
  // 이름 조회 실패 시 공유드라이브 URL의 folders/ 뒤 ID를 여기에 직접 입력.
  SHARED_DRIVE_ID: '',

  // 공유드라이브 루트 바로 아래의 일괄 수집 폴더명
  SOURCE_FOLDER_NAME: '계약서 일괄',

  // 매칭 실패 파일 복사 위치
  UNMATCHED_FOLDER_NAME: '계약서_미매칭',

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
  // false: 계약서_미매칭 폴더로 복사
  // true: 고객번호 없는 회사명 폴더를 루트에 생성. 권장하지 않음.
  CREATE_COMPANY_FOLDER_WITHOUT_CUSTOMER_NO: false,

  // 고객번호/수행사 없을 때 폴더명 기본값
  EMPTY_VENDOR_TEXT: '수행사미정',
  UNKNOWN_CUSTOMER_NO_TEXT: '고객번호미확인',

  // 로그 시트명
  LOG_SHEET_NAME: '계약서_일괄복사_LOG',

  // 진행 저장 키
  PROP_NEXT_FILE_INDEX: 'S1_CONTRACT_COPY_NEXT_FILE_INDEX',
  PROP_PREVIEW_NEXT_FILE_INDEX: 'S1_CONTRACT_PREVIEW_NEXT_FILE_INDEX',

  TZ: 'Asia/Seoul'
};



/***** 공개 실행 함수 *****/

/**
 * 계약서 일괄 복사 실행 함수.
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
function runContractFileCopy() {
  return contractCopy_runContractFileCopyInternal_({
    dryRun: false,
    propKey: CONTRACT_COPY_CFG.PROP_NEXT_FILE_INDEX,
    actionLabel: '계약서 복사',
    nextFunctionName: 'runContractFileCopy'
  });
}


/**
 * 계약서 일괄 복사 미리보기 실행 함수.
 *
 * 사용법:
 * - 이 함수만 계속 실행하면 됩니다.
 * - 실제 복사/폴더 생성은 하지 않고, 매칭 예상 결과만 로그에 남깁니다.
 * - 중간에 끊기면 같은 함수를 다시 실행하면 다음 파일부터 이어서 미리보기합니다.
 */
function previewContractFileCopyTargets() {
  return contractCopy_runContractFileCopyInternal_({
    dryRun: true,
    propKey: CONTRACT_COPY_CFG.PROP_PREVIEW_NEXT_FILE_INDEX,
    actionLabel: '계약서 미리보기',
    nextFunctionName: 'previewContractFileCopyTargets'
  });
}


/**
 * 진행상황 초기화.
 * - 복사/미리보기 모두 처음부터 다시 돌리고 싶을 때만 실행합니다.
 * - 로그 시트를 삭제하고 다시 시작할 때도 이 함수 한 번 실행하면 안전합니다.
 */
function resetContractFileCopyProgress() {
  PropertiesService.getScriptProperties().deleteProperty(CONTRACT_COPY_CFG.PROP_NEXT_FILE_INDEX);
  PropertiesService.getScriptProperties().deleteProperty(CONTRACT_COPY_CFG.PROP_PREVIEW_NEXT_FILE_INDEX);
  Logger.log('계약서 복사/미리보기 진행상황을 초기화했습니다.');
  return {
    ok: true,
    message: '계약서 복사/미리보기 진행상황 초기화 완료'
  };
}


/**
 * 설정/폴더/시트 탐지 점검용.
 */
function reportContractFileCopyEnvironment() {
  const ss = contractCopy_getSpreadsheet_();
  const driveId = contractCopy_getSharedDriveId_();
  const sourceFolder = contractCopy_getSourceFolder_(driveId);
  const sourceFiles = contractCopy_listDirectNonFolderFilesPaged_(sourceFolder.id, driveId);

  const indexes = contractCopy_buildAllIndexes_({
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
 * 앞으로는 runContractFileCopy()만 실행하면 됩니다.
 */
function manualCopyContractFiles() {
  return runContractFileCopy();
}


/**
 * 구버전 함수명 호환용.
 * 앞으로는 runContractFileCopy()만 실행하면 됩니다.
 */
function continueCopyContractFiles() {
  return runContractFileCopy();
}


/**
 * 구버전 함수명 호환용.
 * 앞으로는 previewContractFileCopyTargets()만 실행하면 됩니다.
 */
function continuePreviewContractFileCopyTargets() {
  return previewContractFileCopyTargets();
}


/**
 * 복사/미리보기 공통 실행 엔진.
 * dryRun=false: 실제 복사
 * dryRun=true : 미리보기만 로그 기록
 */
function contractCopy_runContractFileCopyInternal_(options) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const startedAt = Date.now();

  try {
    const cfg = CONTRACT_COPY_CFG;
    const dryRun = !!options.dryRun;
    const propKey = options.propKey;
    const actionLabel = options.actionLabel;
    const nextFunctionName = options.nextFunctionName;

    const ss = contractCopy_getSpreadsheet_();
    const driveId = contractCopy_getSharedDriveId_();

    const sourceFolder = contractCopy_getSourceFolder_(driveId);
    const sourceFiles = contractCopy_listDirectNonFolderFilesPaged_(sourceFolder.id, driveId);

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
    if (contractCopy_isBrCopyLogEmpty_()) {
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

    const indexes = contractCopy_buildAllIndexes_({
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
          const resolvedTargets = contractCopy_resolveTargetsForContractFileFile_({
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
          const results = contractCopy_copyOneContractFileFile_({
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

    contractCopy_appendBrCopyLog_(logs);

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
        ? `계약서 복사 전체 완료: 처리 ${processed}개 / 복사 ${copied}개 / 스킵 ${skipped}개 / 미매칭 ${unmatched}개 / 신규폴더 ${createdFolders}개 / 오류 ${errors}개`
        : `계약서 복사 이번 실행 완료: 처리 ${processed}개 / 복사 ${copied}개 / 스킵 ${skipped}개 / 미매칭 ${unmatched}개 / 신규폴더 ${createdFolders}개 / 오류 ${errors}개. 아직 남았습니다. ${nextFunctionName}()를 다시 실행하세요. 다음 파일 순번: ${fileIndex + 1}`;
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
function contractCopy_isBrCopyLogEmpty_() {
  const ss = contractCopy_getSpreadsheet_();
  const sheet = ss.getSheetByName(CONTRACT_COPY_CFG.LOG_SHEET_NAME);

  if (!sheet) return true;

  return sheet.getLastRow() <= 1;
}


/***** 핵심 처리 *****/

function contractCopy_copyOneContractFileFile_(params) {
  const file = params.file;
  const indexes = params.indexes;
  const driveId = params.driveId;

  const resolvedTargets = contractCopy_resolveTargetsForContractFileFile_({
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
        const unmatchedFolder = contractCopy_ensureUnmatchedFolder_(driveId);
        const unmatchedName = contractCopy_buildCopyFileNameForTarget_(file.name, resolved);

        const exists = contractCopy_findDirectFileByName_(unmatchedFolder.id, driveId, unmatchedName);
        if (exists && !CONTRACT_COPY_CFG.REPLACE_SAME_NAME_FILE) {
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

        if (exists && CONTRACT_COPY_CFG.REPLACE_SAME_NAME_FILE) {
          contractCopy_trashDriveFile_(exists.id);
        }

        contractCopy_copyDriveFileToFolder_(file.id, unmatchedName, unmatchedFolder.id);

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

      const copyName = contractCopy_buildCopyFileNameForTarget_(file.name, resolved);
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

      const exists = contractCopy_findDirectFileByName_(resolved.targetFolder.id, driveId, copyName);

      if (exists && !CONTRACT_COPY_CFG.REPLACE_SAME_NAME_FILE) {
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

      if (exists && CONTRACT_COPY_CFG.REPLACE_SAME_NAME_FILE) {
        contractCopy_trashDriveFile_(exists.id);
      }

      contractCopy_copyDriveFileToFolder_(file.id, copyName, resolved.targetFolder.id);

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

function contractCopy_resolveTargetForContractFileFile_(params) {
  const targets = contractCopy_resolveTargetsForContractFileFile_(params);
  return targets && targets.length ? targets[0] : {
    message: '대상 추출 실패'
  };
}


/**
 * 파일 1개에서 복사 대상 여러 개를 반환.
 * 파일명 앞에 계약번호가 여러 개 있으면 계약번호별로 각각 대상 고객사 폴더를 만든다.
 *
 * 예:
 * - 123_124_한국방송공사 계약서.pdf
 * - 123. 124. 한국방송공사 계약서.pdf
 *
 * 위 경우 123, 124를 각각 계약번호로 보고
 * 수주확정/계약완료 시트에서 각 계약번호의 고객번호를 찾아
 * 각 고객사 폴더에 같은 계약서을 복사한다.
 */
function contractCopy_resolveTargetsForContractFileFile_(params) {
  const file = params.file;
  const indexes = params.indexes;
  const driveId = params.driveId;
  const dryRun = !!params.dryRun;

  const parsed = contractCopy_parseContractFileFileName_(file.name);
  const contractNos = parsed.contractNos || (parsed.contractNo ? [parsed.contractNo] : []);

  if (contractNos.length > 0) {
    const seenContractNos = {};
    const targets = [];

    contractNos.forEach(contractNo => {
      const key = contractCopy_normalizeKey_(contractNo);
      if (!key || seenContractNos[key]) return;

      seenContractNos[key] = true;

      targets.push(contractCopy_resolveTargetByContractNo_({
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
    contractCopy_resolveTargetByCompanyName_({
      parsed,
      indexes,
      driveId,
      dryRun
    })
  ];
}


function contractCopy_resolveTargetByContractNo_(params) {
  const contractNo = contractCopy_cleanValue_(params.contractNo);
  const parsed = params.parsed;
  const indexes = params.indexes;
  const driveId = params.driveId;
  const dryRun = !!params.dryRun;

  const contractKey = contractCopy_normalizeKey_(contractNo);
  const contract = indexes.contractByContractNo[contractKey];

  if (!contract) {
    return {
      contractNo,
      company: parsed.companyNameGuess,
      message: `파일명 앞 계약번호 ${contractNo}를 수주확정/계약완료 시트에서 찾지 못함`
    };
  }

  const customerNo = contractCopy_cleanValue_(contract.customerNo);
  const company = contractCopy_cleanValue_(contract.company || parsed.companyNameGuess);
  const vendor = contractCopy_cleanValue_(contract.vendor || CONTRACT_COPY_CFG.EMPTY_VENDOR_TEXT);

  if (!customerNo) {
    return {
      contractNo,
      company,
      vendor,
      message: `계약번호 ${contractNo} 행에서 고객번호가 비어 있음`
    };
  }

  const ensured = contractCopy_ensureCustomerFolderForContractFile_({
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


function contractCopy_resolveTargetByCompanyName_(params) {
  const parsed = params.parsed;
  const indexes = params.indexes;
  const driveId = params.driveId;
  const dryRun = !!params.dryRun;

  const companyGuess = parsed.companyNameGuess;
  const companyNorm = contractCopy_normalizeCompanyName_(companyGuess);

  if (!companyNorm) {
    return {
      company: companyGuess,
      message: '파일명에서 고객사명을 추출하지 못함'
    };
  }

  // 1. 기존 고객사 폴더명과 매칭
  const folderMatch = contractCopy_findBestFolderByCompanyName_(companyNorm, indexes);

  if (folderMatch && folderMatch.folder) {
    const customerNo = folderMatch.customerNo || '';
    const customerInfo = customerNo ? indexes.customerByCustomerNo[contractCopy_normalizeKey_(customerNo)] : null;

    return {
      company: companyGuess,
      customerNo,
      vendor: customerInfo ? customerInfo.vendor : '',
      targetFolder: folderMatch.folder,
      message: `계약번호 없음. 파일명 고객사명 기준 기존 폴더 매칭 / 점수 ${folderMatch.score}`
    };
  }

  // 2. 시트의 회사명과 매칭해서 고객번호 확보
  const customerMatch = contractCopy_findBestCustomerByCompanyName_(companyNorm, indexes);

  if (customerMatch && customerMatch.customer && customerMatch.customer.customerNo) {
    const customer = customerMatch.customer;

    const ensured = contractCopy_ensureCustomerFolderForContractFile_({
      driveId,
      indexes,
      customerNo: customer.customerNo,
      company: customer.company || companyGuess,
      vendor: customer.vendor || CONTRACT_COPY_CFG.EMPTY_VENDOR_TEXT,
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
  if (CONTRACT_COPY_CFG.CREATE_COMPANY_FOLDER_WITHOUT_CUSTOMER_NO) {
    const folderName = contractCopy_buildCustomerFolderNameForContractFile_(
      CONTRACT_COPY_CFG.UNKNOWN_CUSTOMER_NO_TEXT,
      companyGuess,
      CONTRACT_COPY_CFG.EMPTY_VENDOR_TEXT
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

    const folder = contractCopy_createDriveFolder_(folderName, driveId);

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

function contractCopy_ensureCustomerFolderForContractFile_(params) {
  const driveId = params.driveId;
  const indexes = params.indexes;
  const customerNo = contractCopy_cleanValue_(params.customerNo);
  const company = contractCopy_cleanValue_(params.company);
  const vendor = contractCopy_cleanValue_(params.vendor) || CONTRACT_COPY_CFG.EMPTY_VENDOR_TEXT;
  const dryRun = !!params.dryRun;

  const customerNoKey = contractCopy_normalizeKey_(customerNo);

  if (indexes.folderByCustomerNo[customerNoKey]) {
    return {
      folder: indexes.folderByCustomerNo[customerNoKey].folder,
      createdFolder: false
    };
  }

  const folderName = contractCopy_buildCustomerFolderNameForContractFile_(customerNo, company, vendor);

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

  const folder = contractCopy_createDriveFolder_(folderName, driveId);

  const folderInfo = {
    folder,
    customerNo,
    company,
    companyNorm: contractCopy_normalizeCompanyName_(company),
    vendor
  };

  indexes.folderByCustomerNo[customerNoKey] = folderInfo;
  contractCopy_addToArrayMap_(indexes.folderByCompanyNorm, folderInfo.companyNorm, folderInfo);

  return {
    folder,
    createdFolder: true
  };
}


/***** 인덱스 생성 *****/

function contractCopy_buildAllIndexes_(params) {
  const ss = params.ss;
  const driveId = params.driveId;

  const customerIndex = contractCopy_buildCustomerIndexFromSheets_(ss);
  const contractIndex = contractCopy_buildContractIndexFromSheets_(ss, customerIndex);
  const folderIndex = contractCopy_buildCustomerFolderIndex_(driveId);

  return {
    contractByContractNo: contractIndex.contractByContractNo,
    customerByCustomerNo: customerIndex.customerByCustomerNo,
    customerByCompanyNorm: customerIndex.customerByCompanyNorm,
    folderByCustomerNo: folderIndex.folderByCustomerNo,
    folderByCompanyNorm: folderIndex.folderByCompanyNorm
  };
}


function contractCopy_buildContractIndexFromSheets_(ss, customerIndex) {
  const contractByContractNo = {};

  CONTRACT_COPY_CFG.CONTRACT_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const detected = contractCopy_detectHeaderRowAndMap_(sheet, {
      contractNo: CONTRACT_COPY_HEADER_CANDIDATES_.contractNo,
      customerNo: CONTRACT_COPY_HEADER_CANDIDATES_.customerNo
    });

    if (!detected) return;

    const values = sheet
      .getRange(detected.headerRow + 1, 1, Math.max(sheet.getLastRow() - detected.headerRow, 0), sheet.getLastColumn())
      .getDisplayValues();

    values.forEach(row => {
      const contractNo = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.contractNo);
      const customerNo = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.customerNo);
      const company = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.company);
      const vendor = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.vendor);

      const contractKey = contractCopy_normalizeKey_(contractNo);
      if (!contractKey) return;

      const customerKey = contractCopy_normalizeKey_(customerNo);
      const supplement = customerKey ? customerIndex.customerByCustomerNo[customerKey] : null;

      contractByContractNo[contractKey] = {
        contractNo: contractCopy_cleanValue_(contractNo),
        customerNo: contractCopy_cleanValue_(customerNo || (supplement ? supplement.customerNo : '')),
        company: contractCopy_cleanValue_(company || (supplement ? supplement.company : '')),
        vendor: contractCopy_cleanValue_(vendor || (supplement ? supplement.vendor : '')),
        sourceSheet: sheetName
      };
    });
  });

  return {
    contractByContractNo
  };
}


function contractCopy_buildCustomerIndexFromSheets_(ss) {
  const customerByCustomerNo = {};
  const customerByCompanyNorm = {};

  // 1. 마스터시트 우선
  CONTRACT_COPY_CFG.MASTER_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const detected = contractCopy_detectHeaderRowAndMap_(sheet, {
      customerNo: CONTRACT_COPY_HEADER_CANDIDATES_.customerNo,
      company: CONTRACT_COPY_HEADER_CANDIDATES_.company
    });

    if (!detected) return;

    contractCopy_addCustomersFromSheet_(sheet, detected, customerByCustomerNo, customerByCompanyNorm, sheetName);
  });

  // 2. 계약 관련 시트도 고객 인덱스 보강
  CONTRACT_COPY_CFG.CONTRACT_SHEET_NAMES.forEach(sheetName => {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    const detected = contractCopy_detectHeaderRowAndMap_(sheet, {
      customerNo: CONTRACT_COPY_HEADER_CANDIDATES_.customerNo,
      company: CONTRACT_COPY_HEADER_CANDIDATES_.company
    });

    if (!detected) return;

    contractCopy_addCustomersFromSheet_(sheet, detected, customerByCustomerNo, customerByCompanyNorm, sheetName);
  });

  return {
    customerByCustomerNo,
    customerByCompanyNorm
  };
}


function contractCopy_addCustomersFromSheet_(sheet, detected, customerByCustomerNo, customerByCompanyNorm, sheetName) {
  const rowCount = Math.max(sheet.getLastRow() - detected.headerRow, 0);
  if (rowCount <= 0) return;

  const values = sheet
    .getRange(detected.headerRow + 1, 1, rowCount, sheet.getLastColumn())
    .getDisplayValues();

  values.forEach(row => {
    const customerNo = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.customerNo);
    const company = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.company);
    const vendor = contractCopy_getFirstByCandidatesFromRow_(row, detected.headerMap, CONTRACT_COPY_HEADER_CANDIDATES_.vendor);

    const customerKey = contractCopy_normalizeKey_(customerNo);
    const companyNorm = contractCopy_normalizeCompanyName_(company);

    if (!customerKey || !companyNorm) return;

    const customer = {
      customerNo: contractCopy_cleanValue_(customerNo),
      company: contractCopy_cleanValue_(company),
      companyNorm,
      vendor: contractCopy_cleanValue_(vendor || CONTRACT_COPY_CFG.EMPTY_VENDOR_TEXT),
      sourceSheet: sheetName
    };

    if (!customerByCustomerNo[customerKey]) {
      customerByCustomerNo[customerKey] = customer;
    }

    contractCopy_addToArrayMap_(customerByCompanyNorm, companyNorm, customer);
  });
}


function contractCopy_buildCustomerFolderIndex_(driveId) {
  const folderByCustomerNo = {};
  const folderByCompanyNorm = {};

  const rootFolders = contractCopy_listDirectChildFoldersPaged_(driveId, driveId);

  rootFolders.forEach(folder => {
    const name = contractCopy_cleanValue_(folder.name);

    // 업무용 시스템 폴더는 고객사 폴더 인덱스에서 제외
    if (
      name === CONTRACT_COPY_CFG.SOURCE_FOLDER_NAME ||
      name === CONTRACT_COPY_CFG.UNMATCHED_FOLDER_NAME ||
      name === '수주실패'
    ) {
      return;
    }

    const parsed = contractCopy_parseCustomerFolderName_(name);
    if (!parsed.customerNo && !parsed.companyNorm) return;

    const info = {
      folder,
      customerNo: parsed.customerNo,
      company: parsed.company,
      companyNorm: parsed.companyNorm,
      vendor: parsed.vendor
    };

    if (parsed.customerNo) {
      folderByCustomerNo[contractCopy_normalizeKey_(parsed.customerNo)] = info;
    }

    if (parsed.companyNorm) {
      contractCopy_addToArrayMap_(folderByCompanyNorm, parsed.companyNorm, info);
    }
  });

  return {
    folderByCustomerNo,
    folderByCompanyNorm
  };
}


/***** 파일명/폴더명 파싱 *****/

function contractCopy_parseContractFileFileName_(fileName) {
  const originalName = contractCopy_cleanValue_(fileName);
  const baseName = contractCopy_removeExtension_(originalName);

  const extracted = contractCopy_extractLeadingContractNosAndRest_(baseName);
  const contractNos = extracted.contractNos || [];
  const rest = extracted.rest || baseName;

  const companyNameGuess = contractCopy_extractCompanyNameGuessFromContractFileFile_(rest || baseName);

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
 * - 123_124_한국방송공사 계약서
 * - 123. 124. 한국방송공사 계약서
 * - 123 124 한국방송공사 계약서
 * - 123-124-한국방송공사 계약서
 * - 123) 124) 한국방송공사 계약서
 *
 * 주의:
 * - 여기서 추출하는 번호는 계약번호다.
 * - 고객사 폴더명 앞의 고객번호와 절대 혼동하지 않는다.
 */
function contractCopy_extractLeadingContractNosAndRest_(baseName) {
  let rest = contractCopy_cleanValue_(baseName);
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

function contractCopy_extractCompanyNameGuessFromContractFileFile_(text) {
  let s = contractCopy_cleanValue_(text);

  s = s
    .replace(/사업자\s*등록증/gi, ' ')
    .replace(/계약서/gi, ' ')
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


function contractCopy_parseCustomerFolderName_(folderName) {
  const s = contractCopy_cleanValue_(folderName);

  // 현재: 고객번호_고객사명_수행사
  // 향후: C1_고객사명_수행사 도 허용
  const parts = s.split('_').map(p => contractCopy_cleanValue_(p)).filter(Boolean);

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
    companyNorm: contractCopy_normalizeCompanyName_(company),
    vendor
  };
}


function contractCopy_buildCustomerFolderNameForContractFile_(customerNo, company, vendor) {
  const parts = [
    contractCopy_sanitizeFolderPart_(customerNo),
    contractCopy_sanitizeFolderPart_(company),
    contractCopy_sanitizeFolderPart_(vendor || CONTRACT_COPY_CFG.EMPTY_VENDOR_TEXT)
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
 * 동일 계약서를 여러 계약번호/여러 고객사 폴더에 복사하는 경우에도
 * 각 고객사 폴더가 다르므로 원본명을 유지해도 충돌은 거의 없다.
 * 필요하면 여기서 계약번호별 파일명으로 바꿀 수 있다.
 */
function contractCopy_buildCopyFileNameForTarget_(originalFileName, resolved) {
  return contractCopy_cleanValue_(originalFileName);
}


/***** 매칭 *****/

function contractCopy_findBestFolderByCompanyName_(companyNorm, indexes) {
  return contractCopy_findBestCompanyMatchFromArrayMap_(companyNorm, indexes.folderByCompanyNorm, 'folder');
}


function contractCopy_findBestCustomerByCompanyName_(companyNorm, indexes) {
  return contractCopy_findBestCompanyMatchFromArrayMap_(companyNorm, indexes.customerByCompanyNorm, 'customer');
}


function contractCopy_findBestCompanyMatchFromArrayMap_(companyNorm, map, returnType) {
  const target = contractCopy_cleanValue_(companyNorm);
  if (!target) return null;

  let best = null;
  let second = null;

  Object.keys(map || {}).forEach(key => {
    const score = contractCopy_calcCompanyMatchScore_(target, key);
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


function contractCopy_calcCompanyMatchScore_(a, b) {
  const x = contractCopy_cleanValue_(a);
  const y = contractCopy_cleanValue_(b);

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

const CONTRACT_COPY_HEADER_CANDIDATES_ = {
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


function contractCopy_detectHeaderRowAndMap_(sheet, requiredGroups) {
  const maxRows = Math.min(CONTRACT_COPY_CFG.HEADER_SCAN_ROWS, sheet.getLastRow());
  const lastCol = sheet.getLastColumn();

  if (maxRows <= 0 || lastCol <= 0) return null;

  const rows = sheet.getRange(1, 1, maxRows, lastCol).getDisplayValues();

  for (let r = 0; r < rows.length; r++) {
    const headerMap = {};
    rows[r].forEach((h, i) => {
      const key = contractCopy_normalizeHeader_(h);
      if (key && !headerMap[key]) {
        headerMap[key] = i + 1;
      }
    });

    let ok = true;

    Object.keys(requiredGroups || {}).forEach(groupName => {
      const candidates = requiredGroups[groupName] || [];
      const found = candidates.some(name => !!headerMap[contractCopy_normalizeHeader_(name)]);
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


function contractCopy_getFirstByCandidatesFromRow_(row, headerMap, candidates) {
  for (let i = 0; i < candidates.length; i++) {
    const col = headerMap[contractCopy_normalizeHeader_(candidates[i])];
    if (col) {
      const value = contractCopy_cleanValue_(row[col - 1]);
      if (value) return value;
    }
  }

  return '';
}


/***** Google Drive API *****/

function contractCopy_getSharedDriveId_() {
  if (contractCopy_cleanValue_(CONTRACT_COPY_CFG.SHARED_DRIVE_ID)) {
    return contractCopy_cleanValue_(CONTRACT_COPY_CFG.SHARED_DRIVE_ID);
  }

  const props = PropertiesService.getScriptProperties();
  const cacheKey = 'S1_CONTRACT_COPY_SHARED_DRIVE_ID';
  const cached = props.getProperty(cacheKey);

  if (cached) return cached;

  const q = `name = ${contractCopy_driveQueryString_(CONTRACT_COPY_CFG.SHARED_DRIVE_NAME)}`;

  const data = contractCopy_driveFetch_(
    'drives?pageSize=10&q=' + encodeURIComponent(q) + '&fields=drives(id,name)',
    { method: 'get' }
  );

  const drives = data.drives || [];

  if (drives.length === 0) {
    throw new Error(
      `공유드라이브를 찾지 못했습니다: ${CONTRACT_COPY_CFG.SHARED_DRIVE_NAME}. ` +
      `CONTRACT_COPY_CFG.SHARED_DRIVE_ID에 공유드라이브 ID를 직접 입력하세요.`
    );
  }

  const driveId = drives[0].id;
  props.setProperty(cacheKey, driveId);

  return driveId;
}


function contractCopy_getSourceFolder_(driveId) {
  const folder = contractCopy_findChildFolder_(driveId, driveId, CONTRACT_COPY_CFG.SOURCE_FOLDER_NAME);

  if (!folder) {
    throw new Error(`공유드라이브 루트에서 원본 폴더를 찾지 못했습니다: ${CONTRACT_COPY_CFG.SOURCE_FOLDER_NAME}`);
  }

  return folder;
}


function contractCopy_ensureUnmatchedFolder_(driveId) {
  const existing = contractCopy_findChildFolder_(driveId, driveId, CONTRACT_COPY_CFG.UNMATCHED_FOLDER_NAME);
  if (existing) return existing;

  return contractCopy_createDriveFolder_(CONTRACT_COPY_CFG.UNMATCHED_FOLDER_NAME, driveId);
}


function contractCopy_findChildFolder_(parentId, driveId, folderName) {
  const q = [
    `${contractCopy_driveQueryString_(parentId)} in parents`,
    `mimeType = 'application/vnd.google-apps.folder'`,
    `name = ${contractCopy_driveQueryString_(folderName)}`,
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

  const data = contractCopy_driveFetch_(path, { method: 'get' });
  const files = data.files || [];

  return files.length ? files[0] : null;
}


function contractCopy_createDriveFolder_(folderName, parentId) {
  return contractCopy_driveFetch_(
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


function contractCopy_listDirectChildFoldersPaged_(parentFolderId, driveId) {
  const q = [
    `${contractCopy_driveQueryString_(parentFolderId)} in parents`,
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

    const data = contractCopy_driveFetch_(path, { method: 'get' });

    (data.files || []).forEach(file => folders.push(file));
    pageToken = data.nextPageToken || '';

  } while (pageToken);

  return folders;
}


function contractCopy_listDirectNonFolderFilesPaged_(parentFolderId, driveId) {
  const q = [
    `${contractCopy_driveQueryString_(parentFolderId)} in parents`,
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

    const data = contractCopy_driveFetch_(path, { method: 'get' });

    (data.files || []).forEach(file => files.push(file));
    pageToken = data.nextPageToken || '';

  } while (pageToken);

  files.sort((a, b) => contractCopy_cleanValue_(a.name).localeCompare(contractCopy_cleanValue_(b.name), 'ko'));

  return files;
}


function contractCopy_findDirectFileByName_(parentFolderId, driveId, fileName) {
  const q = [
    `${contractCopy_driveQueryString_(parentFolderId)} in parents`,
    `mimeType != 'application/vnd.google-apps.folder'`,
    `name = ${contractCopy_driveQueryString_(fileName)}`,
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

  const data = contractCopy_driveFetch_(path, { method: 'get' });
  const files = data.files || [];

  return files.length ? files[0] : null;
}


function contractCopy_copyDriveFileToFolder_(sourceFileId, newName, targetFolderId) {
  return contractCopy_driveFetch_(
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


function contractCopy_trashDriveFile_(fileId) {
  return contractCopy_driveFetch_(
    'files/' + encodeURIComponent(fileId) + '?supportsAllDrives=true&fields=id,name,trashed',
    {
      method: 'patch',
      payload: {
        trashed: true
      }
    }
  );
}


function contractCopy_driveFetch_(path, options) {
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


function contractCopy_driveQueryString_(value) {
  const s = contractCopy_cleanValue_(value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");

  return `'${s}'`;
}


/***** 스프레드시트/로그 *****/

function contractCopy_getSpreadsheet_() {
  const id = contractCopy_cleanValue_(CONTRACT_COPY_CFG.SPREADSHEET_ID);

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


function contractCopy_appendBrCopyLog_(rows) {
  if (!rows || rows.length === 0) return;

  const ss = contractCopy_getSpreadsheet_();
  const name = CONTRACT_COPY_CFG.LOG_SHEET_NAME;

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

function contractCopy_addToArrayMap_(map, key, item) {
  const k = contractCopy_cleanValue_(key);
  if (!k) return;

  if (!map[k]) map[k] = [];
  map[k].push(item);
}


function contractCopy_cleanValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}


function contractCopy_normalizeHeader_(value) {
  return contractCopy_cleanValue_(value)
    .replace(/\s+/g, '')
    .replace(/[._\-\/]/g, '')
    .toUpperCase();
}


function contractCopy_normalizeKey_(value) {
  let s = contractCopy_cleanValue_(value);

  if (!s) return '';

  s = s.replace(/\.0$/, '').trim();

  // 순수 숫자는 0001 / 1 비교 가능하게 앞 0 제거
  if (/^\d+$/.test(s)) {
    s = String(Number(s));
  }

  return s.toUpperCase();
}


function contractCopy_normalizeCompanyName_(value) {
  let s = contractCopy_cleanValue_(value);

  if (!s) return '';

  s = contractCopy_removeExtension_(s);

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
    .replace(/계약서/gi, '')
    .replace(/등록증/gi, '')
    .replace(/사본/gi, '')
    .replace(/[0-9]+$/g, '')
    .replace(/[^가-힣a-zA-Z0-9]/g, '')
    .toUpperCase()
    .trim();

  return s;
}


function contractCopy_sanitizeFolderPart_(value) {
  return contractCopy_cleanValue_(value)
    .replace(/[\/\\:*?"<>|#\[\]\r\n\t]/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}


function contractCopy_removeExtension_(fileName) {
  return contractCopy_cleanValue_(fileName).replace(/\.[^.]+$/, '');
}

/*******************************************************
 * 계약서 일괄 복사 롤백.gs
 *
 * 목적:
 * - 방금 runContractFileCopy()로 복사된 계약서 파일만 되돌림.
 * - 계약서_일괄복사_LOG 시트의 처리결과가 COPIED / COPIED_TO_UNMATCHED 인 행만 대상으로 함.
 * - 대상폴더ID + 원본파일명 기준으로 해당 폴더 안의 같은 이름 파일을 휴지통 이동.
 *
 * 안전장치:
 * - 기본은 DRY_RUN=true라 실제 삭제하지 않고 로그만 남김.
 * - 확인 후 CONTRACT_ROLLBACK_CFG.DRY_RUN = false 로 바꾸고 실행.
 * - 파일을 완전삭제하지 않고 휴지통으로 이동.
 * - SKIPPED_EXISTS는 건드리지 않음.
 *
 * 실행 함수:
 * 1) previewRollbackContractCopiedFiles()
 *    - 삭제 예정 미리보기
 *
 * 2) rollbackContractCopiedFiles()
 *    - DRY_RUN=false일 때 실제 휴지통 이동
 *
 * 3) resetRollbackContractProgress()
 *    - 롤백 진행 위치 초기화
 *******************************************************/


const CONTRACT_ROLLBACK_CFG = {
  SPREADSHEET_ID: '',

  // 계약서 복사 코드의 로그 시트명
  SOURCE_LOG_SHEET_NAME: '계약서_일괄복사_LOG',

  // 롤백 로그 시트명
  ROLLBACK_LOG_SHEET_NAME: '계약서_일괄복사_ROLLBACK_LOG',

  // 처음엔 반드시 true
  DRY_RUN: false,

  // 처리대상 상태
  TARGET_STATUSES: [
    'COPIED',
    'COPIED_TO_UNMATCHED'
  ],

  // 1회 처리 제한
  MAX_ROWS_PER_RUN: 200,
  MAX_MILLIS_PER_RUN: 5 * 60 * 1000,

  PROP_NEXT_LOG_ROW: 'S1_CONTRACT_COPY_ROLLBACK_NEXT_LOG_ROW',

  TZ: 'Asia/Seoul'
};


/**
 * 삭제 예정 미리보기.
 * 실제 휴지통 이동 안 함.
 */
function previewRollbackContractCopiedFiles() {
  const original = CONTRACT_ROLLBACK_CFG.DRY_RUN;
  CONTRACT_ROLLBACK_CFG.DRY_RUN = true;

  try {
    return rollbackContractCopiedFiles_();
  } finally {
    CONTRACT_ROLLBACK_CFG.DRY_RUN = original;
  }
}


/**
 * 계약서 일괄 복사분 롤백.
 *
 * 중요:
 * - 실제 삭제하려면 위 설정에서 DRY_RUN을 false로 바꾼 뒤 실행.
 * - 같은 함수를 다시 실행하면 다음 로그 행부터 이어서 처리.
 */
function rollbackContractCopiedFiles() {
  return rollbackContractCopiedFiles_();
}


/**
 * 롤백 진행 위치 초기화.
 */
function resetRollbackContractProgress() {
  PropertiesService.getScriptProperties().deleteProperty(CONTRACT_ROLLBACK_CFG.PROP_NEXT_LOG_ROW);
  Logger.log('계약서 복사 롤백 진행상황을 초기화했습니다.');
  return {
    ok: true,
    message: '계약서 복사 롤백 진행상황 초기화 완료'
  };
}


function rollbackContractCopiedFiles_() {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  const startedAt = Date.now();

  try {
    const cfg = CONTRACT_ROLLBACK_CFG;
    const ss = contractRollback_getSpreadsheet_();
    const logSheet = ss.getSheetByName(cfg.SOURCE_LOG_SHEET_NAME);

    if (!logSheet) {
      throw new Error('원본 로그 시트를 찾지 못했습니다: ' + cfg.SOURCE_LOG_SHEET_NAME);
    }

    const headerMap = contractRollback_getHeaderMap_(logSheet, 1);

    contractRollback_assertHeader_(headerMap, '일시');
    contractRollback_assertHeader_(headerMap, '원본파일명');
    contractRollback_assertHeader_(headerMap, '대상폴더ID');
    contractRollback_assertHeader_(headerMap, '처리결과');

    const lastRow = logSheet.getLastRow();
    if (lastRow < 2) {
      Logger.log('롤백할 로그 데이터가 없습니다.');
      return {
        ok: true,
        status: 'NO_LOG_ROWS',
        message: '롤백할 로그 데이터가 없습니다.'
      };
    }

    const props = PropertiesService.getScriptProperties();
    let rowNo = Number(props.getProperty(cfg.PROP_NEXT_LOG_ROW) || 2);

    if (rowNo < 2) rowNo = 2;

    if (rowNo > lastRow) {
      props.deleteProperty(cfg.PROP_NEXT_LOG_ROW);
      Logger.log('롤백할 남은 로그 행이 없습니다. 이미 완료되었습니다.');
      return {
        ok: true,
        status: 'ALREADY_DONE',
        message: '롤백할 남은 로그 행이 없습니다. 이미 완료되었습니다.'
      };
    }

    const targetStatusSet = {};
    cfg.TARGET_STATUSES.forEach(function(s) {
      targetStatusSet[String(s || '').trim()] = true;
    });

    const lastCol = logSheet.getLastColumn();
    const values = logSheet
      .getRange(rowNo, 1, lastRow - rowNo + 1, lastCol)
      .getDisplayValues();

    const rollbackRows = [];

    let processed = 0;
    let targetRows = 0;
    let wouldTrash = 0;
    let trashed = 0;
    let skipped = 0;
    let notFound = 0;
    let errors = 0;

    const startRowNo = rowNo;

    while (
      processed < values.length &&
      processed < cfg.MAX_ROWS_PER_RUN &&
      Date.now() - startedAt < cfg.MAX_MILLIS_PER_RUN
    ) {
      const currentRowNo = rowNo + processed;
      const row = values[processed];

      const originalFileName = contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '원본파일명') - 1]);
      const targetFolderId = contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '대상폴더ID') - 1]);
      const status = contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '처리결과') - 1]);

      const contractNo = headerMap[contractRollback_normalizeHeader_('계약번호')]
        ? contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '계약번호') - 1])
        : '';
      const customerNo = headerMap[contractRollback_normalizeHeader_('고객번호')]
        ? contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '고객번호') - 1])
        : '';
      const company = headerMap[contractRollback_normalizeHeader_('회사명')]
        ? contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '회사명') - 1])
        : '';
      const targetFolderName = headerMap[contractRollback_normalizeHeader_('대상폴더명')]
        ? contractRollback_cleanValue_(row[contractRollback_col_(headerMap, '대상폴더명') - 1])
        : '';

      try {
        if (!targetStatusSet[status]) {
          skipped++;
          rollbackRows.push([
            new Date(),
            currentRowNo,
            originalFileName,
            contractNo,
            customerNo,
            company,
            targetFolderName,
            targetFolderId,
            status,
            '',
            '',
            'SKIPPED_STATUS',
            '롤백 대상 처리결과가 아니어서 스킵'
          ]);
          processed++;
          continue;
        }

        targetRows++;

        if (!originalFileName || !targetFolderId) {
          skipped++;
          rollbackRows.push([
            new Date(),
            currentRowNo,
            originalFileName,
            contractNo,
            customerNo,
            company,
            targetFolderName,
            targetFolderId,
            status,
            '',
            '',
            'SKIPPED_INVALID_LOG',
            '원본파일명 또는 대상폴더ID가 비어 있음'
          ]);
          processed++;
          continue;
        }

        const foundFiles = contractRollback_findDirectFilesByName_(targetFolderId, originalFileName);

        if (!foundFiles.length) {
          notFound++;
          rollbackRows.push([
            new Date(),
            currentRowNo,
            originalFileName,
            contractNo,
            customerNo,
            company,
            targetFolderName,
            targetFolderId,
            status,
            '',
            '',
            'NOT_FOUND',
            '대상 폴더에서 같은 이름 파일을 찾지 못함'
          ]);
          processed++;
          continue;
        }

        foundFiles.forEach(function(file) {
          if (cfg.DRY_RUN) {
            wouldTrash++;
            rollbackRows.push([
              new Date(),
              currentRowNo,
              originalFileName,
              contractNo,
              customerNo,
              company,
              targetFolderName,
              targetFolderId,
              status,
              file.name || '',
              file.id || '',
              'DRY_RUN_WOULD_TRASH',
              '휴지통 이동 예정'
            ]);
          } else {
            contractRollback_trashDriveFile_(file.id);
            trashed++;
            rollbackRows.push([
              new Date(),
              currentRowNo,
              originalFileName,
              contractNo,
              customerNo,
              company,
              targetFolderName,
              targetFolderId,
              status,
              file.name || '',
              file.id || '',
              'TRASHED',
              '휴지통 이동 완료'
            ]);
          }
        });

      } catch (err) {
        errors++;
        rollbackRows.push([
          new Date(),
          currentRowNo,
          originalFileName,
          contractNo,
          customerNo,
          company,
          targetFolderName,
          targetFolderId,
          status,
          '',
          '',
          'ERROR',
          err && err.message ? err.message : String(err)
        ]);
      }

      processed++;
    }

    contractRollback_appendRollbackLog_(rollbackRows);

    const nextRow = rowNo + processed;
    const done = nextRow > lastRow;

    if (done) {
      props.deleteProperty(cfg.PROP_NEXT_LOG_ROW);
    } else {
      props.setProperty(cfg.PROP_NEXT_LOG_ROW, String(nextRow));
    }

    const msg = done
      ? `계약서 복사 롤백 전체 완료: 로그행 ${startRowNo}~${nextRow - 1} 처리 / 대상 ${targetRows}건 / 삭제예정 ${wouldTrash}건 / 휴지통이동 ${trashed}건 / 미발견 ${notFound}건 / 스킵 ${skipped}건 / 오류 ${errors}건`
      : `계약서 복사 롤백 이번 실행 완료: 로그행 ${startRowNo}~${nextRow - 1} 처리 / 대상 ${targetRows}건 / 삭제예정 ${wouldTrash}건 / 휴지통이동 ${trashed}건 / 미발견 ${notFound}건 / 스킵 ${skipped}건 / 오류 ${errors}건. 아직 남았습니다. rollbackContractCopiedFiles()를 다시 실행하세요. 다음 로그 행: ${nextRow}`;

    Logger.log(msg);

    return {
      ok: errors === 0,
      dryRun: cfg.DRY_RUN,
      status: done ? 'DONE' : 'CONTINUE_NEEDED',
      startRowNo,
      nextRowNo: done ? '' : nextRow,
      processed,
      targetRows,
      wouldTrash,
      trashed,
      notFound,
      skipped,
      errors,
      message: msg
    };

  } finally {
    lock.releaseLock();
  }
}


/***** Drive 처리 *****/

function contractRollback_findDirectFilesByName_(parentFolderId, fileName) {
  const q = [
    `${contractRollback_driveQueryString_(parentFolderId)} in parents`,
    `mimeType != 'application/vnd.google-apps.folder'`,
    `name = ${contractRollback_driveQueryString_(fileName)}`,
    `trashed = false`
  ].join(' and ');

  const path =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=allDrives' +
    '&pageSize=50' +
    '&q=' + encodeURIComponent(q) +
    '&fields=files(id,name,webViewLink,trashed,parents)';

  const data = contractRollback_driveFetch_(path, { method: 'get' });
  return data.files || [];
}


function contractRollback_trashDriveFile_(fileId) {
  return contractRollback_driveFetch_(
    'files/' + encodeURIComponent(fileId) + '?supportsAllDrives=true&fields=id,name,trashed',
    {
      method: 'patch',
      payload: {
        trashed: true
      }
    }
  );
}


function contractRollback_driveFetch_(path, options) {
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


function contractRollback_driveQueryString_(value) {
  const s = contractRollback_cleanValue_(value)
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");

  return `'${s}'`;
}


/***** 스프레드시트/로그 *****/

function contractRollback_getSpreadsheet_() {
  const id = contractRollback_cleanValue_(CONTRACT_ROLLBACK_CFG.SPREADSHEET_ID);

  if (id) {
    return SpreadsheetApp.openById(id);
  }

  return SpreadsheetApp.getActiveSpreadsheet();
}


function contractRollback_appendRollbackLog_(rows) {
  if (!rows || rows.length === 0) return;

  const ss = contractRollback_getSpreadsheet_();
  const name = CONTRACT_ROLLBACK_CFG.ROLLBACK_LOG_SHEET_NAME;

  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow([
      '일시',
      '원본로그행',
      '원본파일명',
      '계약번호',
      '고객번호',
      '회사명',
      '대상폴더명',
      '대상폴더ID',
      '원본처리결과',
      '삭제대상파일명',
      '삭제대상파일ID',
      '롤백결과',
      '메시지'
    ]);
    sheet.setFrozenRows(1);
  }

  sheet
    .getRange(sheet.getLastRow() + 1, 1, rows.length, rows[0].length)
    .setValues(rows);
}


function contractRollback_getHeaderMap_(sheet, headerRow) {
  const lastCol = sheet.getLastColumn();

  const headers = sheet
    .getRange(headerRow, 1, 1, lastCol)
    .getDisplayValues()[0];

  const map = {};

  headers.forEach(function(h, i) {
    const key = contractRollback_normalizeHeader_(h);
    if (key && !map[key]) {
      map[key] = i + 1;
    }
  });

  return map;
}


function contractRollback_assertHeader_(headerMap, headerName) {
  if (!headerMap[contractRollback_normalizeHeader_(headerName)]) {
    throw new Error('필수 헤더를 찾지 못했습니다: ' + headerName);
  }
}


function contractRollback_col_(headerMap, headerName) {
  const key = contractRollback_normalizeHeader_(headerName);
  const c = headerMap[key];

  if (!c) {
    throw new Error('헤더 컬럼을 찾지 못했습니다: ' + headerName);
  }

  return c;
}


function contractRollback_cleanValue_(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim();
}


function contractRollback_normalizeHeader_(value) {
  return contractRollback_cleanValue_(value)
    .replace(/\s+/g, '')
    .replace(/[._\-\/]/g, '')
    .toUpperCase();
}

