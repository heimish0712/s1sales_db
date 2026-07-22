/****************************************************
 * 영업관리대장 자동 백업
 *
 * 저장 위치:
 * 공유드라이브 > S1 영업포털 운영 > 영업관리대장 백업
 *
 * 실행 시간:
 * 매일 12:30
 * 매일 18:00
 *
 * 파일명:
 * NEW SH 영업관리대장_백업_2026년06월28일_12시30분45초
 ****************************************************/


/**
 * [필수 설정]
 * 백업 폴더 ID를 넣으세요.
 *
 * 구글드라이브에서
 * 공유드라이브 > S1 영업포털 운영 > 영업관리대장 백업
 * 폴더를 열고,
 *
 * URL이 이런 식이면:
 * https://drive.google.com/drive/folders/1AbCdEfGhIjKlMn...
 *
 * 아래 BACKUP_FOLDER_ID에 1AbCdEfGhIjKlMn... 부분만 넣으면 됩니다.
 */
const BACKUP_FOLDER_ID = '1yycNk-XMFyEzY2GC3FLuFUMw87QfN7xk';


/**
 * 백업 파일명 접두어
 */
const BACKUP_FILE_PREFIX = 'NEW SH 영업관리대장_백업';


/**
 * 실제 백업 실행 함수
 * 트리거가 이 함수를 실행합니다.
 */
function backupSalesLedger() {
  const lease = AUTOMATION_acquireModuleLease_(
    'BACKUP',
    {
      taskName: 'backupSalesLedger',
      waitMs: 500,
      ttlMs: 8 * 60 * 1000
    }
  );

  if (!lease.acquired) {
    const skipped = {
      status: 'SKIPPED_ALREADY_RUNNING',
      message: '다른 영업관리대장 백업이 실행 중이라 이번 실행은 건너뜁니다.',
      leaseReason: lease.reason || 'LEASE_BUSY'
    };
    Logger.log(JSON.stringify(skipped));
    return skipped;
  }

  try {
    if (!BACKUP_FOLDER_ID || BACKUP_FOLDER_ID === '여기에_영업관리대장_백업_폴더ID_입력') {
      throw new Error('BACKUP_FOLDER_ID가 설정되지 않았습니다. 백업 폴더 ID를 입력하세요.');
    }

    const ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
    const sourceFile = DriveApp.getFileById(ss.getId());
    const backupFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);

    const now = new Date();
    const fileName = makeBackupFileName_(now);
    const copiedFile = sourceFile.makeCopy(fileName, backupFolder);

    Logger.log('백업 완료');
    Logger.log('백업 파일명: ' + fileName);
    Logger.log('백업 파일 URL: ' + copiedFile.getUrl());

    return {
      status: 'SUCCESS',
      fileName: fileName,
      fileId: copiedFile.getId(),
      fileUrl: copiedFile.getUrl()
    };
  } catch (err) {
    Logger.log('백업 실패: ' + err.message);
    throw err;
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


/**
 * 백업 파일명 생성
 * 예:
 * NEW SH 영업관리대장_백업_2026년06월28일_12시30분45초
 */
function makeBackupFileName_(date) {
  const tz = Session.getScriptTimeZone() || 'Asia/Seoul';

  const yyyy = Utilities.formatDate(date, tz, 'yyyy');
  const MM = Utilities.formatDate(date, tz, 'MM');
  const dd = Utilities.formatDate(date, tz, 'dd');
  const HH = Utilities.formatDate(date, tz, 'HH');
  const mm = Utilities.formatDate(date, tz, 'mm');
  const ss = Utilities.formatDate(date, tz, 'ss');

  return `${BACKUP_FILE_PREFIX}_${yyyy}년${MM}월${dd}일_${HH}시${mm}분${ss}초`;
}




/**
 * 수동 테스트용 함수
 *
 * 설정 후 이 함수를 먼저 실행해서
 * 백업 폴더에 파일이 정상 생성되는지 확인하세요.
 */
function testBackupSalesLedgerNow() {
  backupSalesLedger();
}
