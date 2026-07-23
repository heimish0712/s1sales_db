/****************************************************
 * 영업관리대장 자동 백업 + 보존정책
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
 */
const BACKUP_FOLDER_ID = '1yycNk-XMFyEzY2GC3FLuFUMw87QfN7xk';


/**
 * 백업 파일명 접두어
 */
const BACKUP_FILE_PREFIX = 'NEW SH 영업관리대장_백업';


/**
 * 백업 보존정책
 *
 * - 최근 30일: 하루 2회 백업 전부 보존
 * - 31~180일: 날짜별 최신 1개 보존
 * - 181~730일: 월별 최신 1개 보존
 * - 730일 초과: 정리 대상
 * - 위 규칙과 무관하게 최신 60개는 무조건 보존
 * - 정리 대상은 영구삭제가 아니라 Drive 휴지통으로 이동
 */
var BACKUP_RETENTION_CONFIG = Object.freeze({
  version: '2026-07-19-PHASE14',
  recentAllDays: 30,
  dailyLatestUntilDays: 180,
  monthlyLatestUntilDays: 730,
  minimumKeepCount: 60,
  scheduledMaxTrash: 50,
  manualMaxTrash: 1000,
  maxScanFiles: 5000,
  lastResultPropertyKey: 'AUTOMATION_BACKUP_RETENTION_LAST_RUN_V1',
  statusSheetName: '_백업보존상태'
});


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
    if (typeof AUTOMATION_recordBackupExecution_ === 'function') {
      AUTOMATION_recordBackupExecution_(skipped);
    }
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

    const successResult = {
      status: 'SUCCESS',
      successAt: now.toISOString(),
      fileName: fileName,
      fileId: copiedFile.getId(),
      fileUrl: copiedFile.getUrl()
    };

    if (typeof AUTOMATION_recordBackupExecution_ === 'function') {
      AUTOMATION_recordBackupExecution_(successResult);
    }

    return successResult;
  } catch (err) {
    Logger.log('백업 실패: ' + err.message);
    if (typeof AUTOMATION_recordBackupExecution_ === 'function') {
      AUTOMATION_recordBackupExecution_({
        status: 'ERROR',
        error: String(err && err.stack || err && err.message || err || '')
      });
    }
    throw err;
  } finally {
    AUTOMATION_releaseModuleLease_(lease);
  }
}


/**
 * 백업 파일명 생성
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
 */
function testBackupSalesLedgerNow() {
  backupSalesLedger();
}


/****************************************************
 * 백업 보존정책 공개 함수
 ****************************************************/

/**
 * 파일을 변경하지 않고 현재 보존·정리 대상을 계산한다.
 */
function AUTOMATION_previewBackupRetention() {
  TRG_assertAutomationOwner_();

  var preview = AUTOMATION_getBackupRetentionPreview_();
  var sample = (preview.candidateSamples || []).slice(0, 10).map(function(item) {
    return '- ' + item.name + ' [' + item.reason + ']';
  });

  SpreadsheetApp.getUi().alert(
    '백업 보존정책 미리보기',
    [
      '상태: ' + String(preview.status || ''),
      '백업 폴더 전체 파일: ' + Number(preview.scannedFiles || 0) + '개',
      '정책 대상 백업: ' + Number(preview.matchedBackups || 0) + '개',
      '보존: ' + Number(preview.keepCount || 0) + '개',
      '휴지통 이동 대상: ' + Number(preview.deleteEligible || 0) + '개',
      '정책 외 파일 보존: ' + Number(preview.ignoredFiles || 0) + '개',
      '',
      '정책: 최근 30일 전체 / 31~180일 일별 최신 / 181~730일 월별 최신',
      '최신 60개는 무조건 보존하며 정리 대상은 영구삭제하지 않습니다.',
      sample.length ? ('\n정리 대상 예시:\n' + sample.join('\n')) : '',
      preview.error ? ('\n오류: ' + preview.error) : ''
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return preview;
}


/**
 * 보존정책을 수동으로 즉시 적용한다.
 */
function AUTOMATION_runBackupRetentionNow() {
  TRG_assertAutomationOwner_();

  var preview = AUTOMATION_getBackupRetentionPreview_();
  if (preview.status !== 'SUCCESS') {
    throw new Error('백업 보존정책 미리보기 실패: ' + String(preview.error || preview.status || ''));
  }

  if (!preview.deleteEligible) {
    SpreadsheetApp.getUi().alert(
      '백업 보존정책',
      '현재 휴지통으로 이동할 백업이 없습니다.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return preview;
  }

  var response = SpreadsheetApp.getUi().alert(
    '백업 보존정책 적용',
    [
      '정리 대상 백업: ' + preview.deleteEligible + '개',
      '최신 보존 백업: ' + preview.keepCount + '개',
      '',
      '대상 파일은 영구삭제하지 않고 Drive 휴지통으로 이동합니다.',
      '계속하시겠습니까?'
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (response !== SpreadsheetApp.getUi().Button.YES) {
    return { status: 'CANCELLED', preview: preview };
  }

  var result = AUTOMATION_cleanupBackupRetention_({
    mode: 'MANUAL',
    maxTrash: BACKUP_RETENTION_CONFIG.manualMaxTrash
  });

  SpreadsheetApp.getUi().alert(
    '백업 보존정책 결과',
    [
      '상태: ' + String(result.status || ''),
      '정리 대상: ' + Number(result.eligible || 0) + '개',
      '휴지통 이동: ' + Number(result.trashed || 0) + '개',
      '실패: ' + Number(result.failed || 0) + '개',
      '남은 대상 추정: ' + Number(result.remainingEstimate || 0) + '개',
      result.errors && result.errors.length ? ('\n오류:\n' + result.errors.join('\n')) : ''
    ].join('\n'),
    SpreadsheetApp.getUi().ButtonSet.OK
  );

  return result;
}


/**
 * 최근 보존정책 실행 상태 시트를 연다.
 */
function AUTOMATION_showBackupRetentionStatusSheet() {
  TRG_assertAutomationOwner_();

  var ss = AUTOMATION_getRuntimeMasterSpreadsheet_();
  var name = BACKUP_RETENTION_CONFIG.statusSheetName;
  var sheet = ss.getSheetByName(name) || ss.insertSheet(name);
  var last = AUTOMATION_readBackupRetentionLastResult_();
  var preview = AUTOMATION_getBackupRetentionPreview_();
  var rows = [
    ['항목', '값'],
    ['정책 버전', BACKUP_RETENTION_CONFIG.version],
    ['최근 전체보존 일수', BACKUP_RETENTION_CONFIG.recentAllDays],
    ['일별 최신 보존 종료일', BACKUP_RETENTION_CONFIG.dailyLatestUntilDays],
    ['월별 최신 보존 종료일', BACKUP_RETENTION_CONFIG.monthlyLatestUntilDays],
    ['최소 보존 개수', BACKUP_RETENTION_CONFIG.minimumKeepCount],
    ['현재 정책 대상 백업', preview.matchedBackups || 0],
    ['현재 보존 개수', preview.keepCount || 0],
    ['현재 정리 대상', preview.deleteEligible || 0],
    ['최근 실행 상태', last.status || ''],
    ['최근 실행시각', last.finishedAt || ''],
    ['최근 휴지통 이동', last.trashed || 0],
    ['최근 실패', last.failed || 0],
    ['최근 오류', (last.errors || []).join(' | ')]
  ];

  sheet.clearContents();
  sheet.getRange(1, 1, rows.length, 2).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, 2);
  sheet.showSheet();
  ss.setActiveSheet(sheet);
  return name;
}


/****************************************************
 * 백업 보존정책 내부 구현
 ****************************************************/

function AUTOMATION_getBackupRetentionPreview_() {
  try {
    var plan = AUTOMATION_buildBackupRetentionPlan_();
    return AUTOMATION_summarizeBackupRetentionPlan_(plan);
  } catch (err) {
    return {
      status: 'ERROR',
      scannedFiles: 0,
      matchedBackups: 0,
      keepCount: 0,
      deleteEligible: 0,
      ignoredFiles: 0,
      candidateSamples: [],
      error: AUTOMATION_backupRetentionError_(err)
    };
  }
}


function AUTOMATION_cleanupBackupRetention_(options) {
  options = options || {};
  var mode = String(options.mode || 'SCHEDULED').toUpperCase();
  var maxTrash = Math.max(
    1,
    Number(options.maxTrash || (
      mode === 'MANUAL'
        ? BACKUP_RETENTION_CONFIG.manualMaxTrash
        : BACKUP_RETENTION_CONFIG.scheduledMaxTrash
    )) || 1
  );

  var lease = AUTOMATION_acquireModuleLease_('BACKUP', {
    taskName: 'AUTOMATION_cleanupBackupRetention_',
    waitMs: mode === 'MANUAL' ? 1000 : 0,
    ttlMs: 8 * 60 * 1000
  });

  if (!lease.acquired) {
    return {
      status: 'LEASE_BUSY',
      done: false,
      eligible: 0,
      trashed: 0,
      failed: 0,
      remainingEstimate: 0,
      reason: lease.reason || 'LEASE_BUSY'
    };
  }

  var result = {
    version: BACKUP_RETENTION_CONFIG.version,
    mode: mode,
    status: 'STARTED',
    startedAt: new Date().toISOString(),
    finishedAt: '',
    scannedFiles: 0,
    matchedBackups: 0,
    keepCount: 0,
    eligible: 0,
    trashed: 0,
    failed: 0,
    remainingEstimate: 0,
    done: false,
    errors: []
  };

  try {
    var plan = AUTOMATION_buildBackupRetentionPlan_();
    result.scannedFiles = plan.scannedFiles;
    result.matchedBackups = plan.files.length;
    result.keepCount = plan.keep.length;
    result.eligible = plan.candidates.length;

    var selected = plan.candidates.slice(0, maxTrash);
    for (var i = 0; i < selected.length; i++) {
      try {
        DriveApp.getFileById(selected[i].id).setTrashed(true);
        result.trashed++;
      } catch (err) {
        result.failed++;
        if (result.errors.length < 10) {
          result.errors.push(selected[i].name + ': ' + AUTOMATION_backupRetentionError_(err));
        }
      }
    }

    result.remainingEstimate = Math.max(0, result.eligible - result.trashed);
    result.done = result.remainingEstimate === 0 && result.failed === 0;
    result.status = result.failed
      ? 'PARTIAL_ERROR'
      : (result.done ? 'SUCCESS' : 'PARTIAL_CONTINUE');
  } catch (err) {
    result.status = 'ERROR';
    result.errors.push(AUTOMATION_backupRetentionError_(err));
  } finally {
    result.finishedAt = new Date().toISOString();
    AUTOMATION_releaseModuleLease_(lease);
    AUTOMATION_recordBackupRetentionResult_(result);
  }

  return result;
}


function AUTOMATION_buildBackupRetentionPlan_() {
  if (!BACKUP_FOLDER_ID || BACKUP_FOLDER_ID === '여기에_영업관리대장_백업_폴더ID_입력') {
    throw new Error('BACKUP_FOLDER_ID가 설정되지 않았습니다.');
  }

  var folder = DriveApp.getFolderById(String(BACKUP_FOLDER_ID));
  var iterator = folder.getFiles();
  var files = [];
  var scanned = 0;
  var ignored = 0;

  while (iterator.hasNext()) {
    scanned++;
    if (scanned > BACKUP_RETENTION_CONFIG.maxScanFiles) {
      throw new Error(
        '백업 폴더 파일이 ' + BACKUP_RETENTION_CONFIG.maxScanFiles +
        '개를 초과하여 안전상 정리를 중단했습니다.'
      );
    }

    var file = iterator.next();
    var name = String(file.getName() || '');
    if (name.indexOf(String(BACKUP_FILE_PREFIX || '')) !== 0) {
      ignored++;
      continue;
    }

    var mimeType = '';
    try { mimeType = String(file.getMimeType() || ''); } catch (ignoreMimeError) {}
    if (mimeType && mimeType !== MimeType.GOOGLE_SHEETS) {
      ignored++;
      continue;
    }

    var created = null;
    try { created = file.getDateCreated(); } catch (ignoreCreatedError) {}
    if (!created || !isFinite(created.getTime())) {
      ignored++;
      continue;
    }

    files.push({
      id: file.getId(),
      name: name,
      createdAtMs: created.getTime(),
      createdAt: created.toISOString(),
      url: file.getUrl()
    });
  }

  var classified = AUTOMATION_classifyBackupRetentionFiles_(files, Date.now());
  classified.scannedFiles = scanned;
  classified.ignoredFiles = ignored;
  return classified;
}


/**
 * 순수 분류 함수. newest-first 정렬 후 보존·정리 대상을 반환한다.
 */
function AUTOMATION_classifyBackupRetentionFiles_(inputFiles, nowMs) {
  var files = (inputFiles || []).slice().sort(function(a, b) {
    var diff = Number(b.createdAtMs || 0) - Number(a.createdAtMs || 0);
    if (diff) return diff;
    return String(b.name || '').localeCompare(String(a.name || ''));
  });

  var dayMs = 24 * 60 * 60 * 1000;
  var recentCutoffMs = nowMs - BACKUP_RETENTION_CONFIG.recentAllDays * dayMs;
  var dailyCutoffMs = nowMs - BACKUP_RETENTION_CONFIG.dailyLatestUntilDays * dayMs;
  var monthlyCutoffMs = nowMs - BACKUP_RETENTION_CONFIG.monthlyLatestUntilDays * dayMs;
  var keep = [];
  var candidates = [];
  var seenDay = {};
  var seenMonth = {};

  files.forEach(function(file, index) {
    var createdMs = Number(file.createdAtMs || 0);
    var dayKey = AUTOMATION_backupRetentionDateKey_(createdMs);
    var monthKey = AUTOMATION_backupRetentionMonthKey_(createdMs);
    var reason = '';
    var shouldKeep = false;

    if (!createdMs || !isFinite(createdMs)) {
      shouldKeep = true;
      reason = 'INVALID_DATE_SAFE_KEEP';
    } else if (index < BACKUP_RETENTION_CONFIG.minimumKeepCount) {
      shouldKeep = true;
      reason = 'MINIMUM_KEEP_' + BACKUP_RETENTION_CONFIG.minimumKeepCount;
    } else if (createdMs >= recentCutoffMs) {
      shouldKeep = true;
      reason = 'RECENT_ALL_' + BACKUP_RETENTION_CONFIG.recentAllDays + '_DAYS';
    } else if (createdMs >= dailyCutoffMs) {
      if (!seenDay[dayKey]) {
        shouldKeep = true;
        reason = 'DAILY_LATEST';
      } else {
        reason = 'DAILY_DUPLICATE';
      }
    } else if (createdMs >= monthlyCutoffMs) {
      if (!seenMonth[monthKey]) {
        shouldKeep = true;
        reason = 'MONTHLY_LATEST';
      } else {
        reason = 'MONTHLY_DUPLICATE';
      }
    } else {
      reason = 'OLDER_THAN_' + BACKUP_RETENTION_CONFIG.monthlyLatestUntilDays + '_DAYS';
    }

    var item = {
      id: String(file.id || ''),
      name: String(file.name || ''),
      createdAtMs: createdMs,
      createdAt: String(file.createdAt || (createdMs ? new Date(createdMs).toISOString() : '')),
      url: String(file.url || ''),
      reason: reason
    };

    if (shouldKeep) {
      keep.push(item);
      if (createdMs >= dailyCutoffMs && createdMs < recentCutoffMs && dayKey) {
        seenDay[dayKey] = true;
      }
      if (createdMs >= monthlyCutoffMs && createdMs < dailyCutoffMs && monthKey) {
        seenMonth[monthKey] = true;
      }
    } else {
      candidates.push(item);
    }
  });

  candidates.sort(function(a, b) {
    return Number(a.createdAtMs || 0) - Number(b.createdAtMs || 0);
  });

  return {
    files: files,
    keep: keep,
    candidates: candidates,
    scannedFiles: files.length,
    ignoredFiles: 0
  };
}


function AUTOMATION_summarizeBackupRetentionPlan_(plan) {
  return {
    status: 'SUCCESS',
    generatedAt: new Date().toISOString(),
    scannedFiles: Number(plan.scannedFiles || 0),
    matchedBackups: (plan.files || []).length,
    keepCount: (plan.keep || []).length,
    deleteEligible: (plan.candidates || []).length,
    ignoredFiles: Number(plan.ignoredFiles || 0),
    candidateSamples: (plan.candidates || []).slice(0, 20).map(function(item) {
      return {
        id: item.id,
        name: item.name,
        createdAt: item.createdAt,
        reason: item.reason
      };
    }),
    policy: {
      recentAllDays: BACKUP_RETENTION_CONFIG.recentAllDays,
      dailyLatestUntilDays: BACKUP_RETENTION_CONFIG.dailyLatestUntilDays,
      monthlyLatestUntilDays: BACKUP_RETENTION_CONFIG.monthlyLatestUntilDays,
      minimumKeepCount: BACKUP_RETENTION_CONFIG.minimumKeepCount
    },
    error: ''
  };
}


function AUTOMATION_backupRetentionDateKey_(timeMs) {
  if (!timeMs || !isFinite(Number(timeMs))) return '';
  return Utilities.formatDate(
    new Date(Number(timeMs)),
    Session.getScriptTimeZone() || 'Asia/Seoul',
    'yyyy-MM-dd'
  );
}


function AUTOMATION_backupRetentionMonthKey_(timeMs) {
  if (!timeMs || !isFinite(Number(timeMs))) return '';
  return Utilities.formatDate(
    new Date(Number(timeMs)),
    Session.getScriptTimeZone() || 'Asia/Seoul',
    'yyyy-MM'
  );
}


function AUTOMATION_recordBackupRetentionResult_(result) {
  var compact = {
    version: BACKUP_RETENTION_CONFIG.version,
    mode: String(result.mode || ''),
    status: String(result.status || ''),
    startedAt: String(result.startedAt || ''),
    finishedAt: String(result.finishedAt || ''),
    scannedFiles: Number(result.scannedFiles || 0),
    matchedBackups: Number(result.matchedBackups || 0),
    keepCount: Number(result.keepCount || 0),
    eligible: Number(result.eligible || 0),
    trashed: Number(result.trashed || 0),
    failed: Number(result.failed || 0),
    remainingEstimate: Number(result.remainingEstimate || 0),
    errors: (result.errors || []).slice(0, 5).map(function(item) {
      return String(item || '').slice(0, 500);
    })
  };

  PropertiesService.getScriptProperties().setProperty(
    BACKUP_RETENTION_CONFIG.lastResultPropertyKey,
    JSON.stringify(compact)
  );
  return compact;
}


function AUTOMATION_readBackupRetentionLastResult_() {
  var raw = String(
    PropertiesService.getScriptProperties().getProperty(
      BACKUP_RETENTION_CONFIG.lastResultPropertyKey
    ) || ''
  );
  if (!raw) return {};
  try { return JSON.parse(raw) || {}; } catch (err) { return {}; }
}


function AUTOMATION_backupRetentionError_(err) {
  return String(err && err.stack || err && err.message || err || '').slice(0, 1200);
}
