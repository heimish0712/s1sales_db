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
  const lock = LockService.getScriptLock();

  try {
    // 중복 실행 방지
    lock.waitLock(30000);

    if (!BACKUP_FOLDER_ID || BACKUP_FOLDER_ID === '여기에_영업관리대장_백업_폴더ID_입력') {
      throw new Error('BACKUP_FOLDER_ID가 설정되지 않았습니다. 백업 폴더 ID를 입력하세요.');
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sourceFile = DriveApp.getFileById(ss.getId());
    const backupFolder = DriveApp.getFolderById(BACKUP_FOLDER_ID);

    const now = new Date();
    const fileName = makeBackupFileName_(now);

    const copiedFile = sourceFile.makeCopy(fileName, backupFolder);

    Logger.log('백업 완료');
    Logger.log('백업 파일명: ' + fileName);
    Logger.log('백업 파일 URL: ' + copiedFile.getUrl());

  } catch (err) {
    Logger.log('백업 실패: ' + err.message);
    throw err;

  } finally {
    try {
      lock.releaseLock();
    } catch (e) {
      // lock이 없을 경우 무시
    }
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
 * 자동 백업 트리거 설치 함수
 *
 * 이 함수는 최초 1회만 직접 실행하세요.
 * 기존 backupSalesLedger 트리거를 지우고,
 * 매일 12:30 / 18:00 트리거를 새로 만듭니다.
 */
function installSalesLedgerBackupTriggers() {
  deleteSalesLedgerBackupTriggers_();

  // 매일 12:30 근처 실행
  ScriptApp.newTrigger('backupSalesLedger')
    .timeBased()
    .everyDays(1)
    .atHour(12)
    .nearMinute(30)
    .create();

  // 매일 18:00 근처 실행
  ScriptApp.newTrigger('backupSalesLedger')
    .timeBased()
    .everyDays(1)
    .atHour(18)
    .nearMinute(0)
    .create();

  Logger.log('영업관리대장 자동 백업 트리거 설치 완료');
}


/**
 * 기존 백업 트리거 삭제
 * 트리거 중복 방지용
 */
function deleteSalesLedgerBackupTriggers_() {
  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(trigger => {
    if (trigger.getHandlerFunction() === 'backupSalesLedger') {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  Logger.log('기존 backupSalesLedger 트리거 삭제 완료');
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
