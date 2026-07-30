/****************************************************
 * SystemLogRouter.gs
 * 영업관리대장 로그·큐·기술 시트 분리 저장 / 3일 CSV 아카이브
 ****************************************************/

var SYSTEM_LOG_CONFIG = Object.freeze({
  version: '2026-07-30-PHASE25',
  masterSpreadsheetId: '1ADDJMrej-EJBw4QHkq17xefWxQw5hZ_NgQLf2BuCD8Q',
  systemSpreadsheetId: '11pP3By_m09y_VKECfkpVLzOegXj__5YtShs7r3N1WZc',
  csvArchiveFolderId: '1qr7DwtgVA20de4yfuOJidqOm-PAxXxIn',
  backupFolderId: '1z04IFEUy6FLJN4YJj6-sWsnYYGIMLg2w',
  previousBackupFolderId: '1yycNk-XMFyEzY2GC3FLuFUMw87QfN7xk',
  retentionDays: 3,
  archiveLastRunPropertyKey: 'SYSTEM_LOG_ARCHIVE_LAST_RUN_V1',
  migrationPropertyKey: 'SYSTEM_LOG_MIGRATION_PHASE25_V1',
  backupMigrationPropertyKey: 'SYSTEM_LOG_BACKUP_MIGRATION_PHASE25_V1'
});

var SYSTEM_LOG_SPREADSHEET_CACHE_ = null;

function SYSTEMLOG_getSpreadsheet_() {
  if (!SYSTEM_LOG_SPREADSHEET_CACHE_) {
    SYSTEM_LOG_SPREADSHEET_CACHE_ = SpreadsheetApp.openById(SYSTEM_LOG_CONFIG.systemSpreadsheetId);
  }
  return SYSTEM_LOG_SPREADSHEET_CACHE_;
}

function SYSTEMLOG_getOrCreateSheet_(sheetName, headers, hideSheet) {
  var ss = SYSTEMLOG_getSpreadsheet_();
  var name = String(sheetName || '').trim();
  if (!name) throw new Error('시스템 로그 시트명이 비어 있습니다.');
  var sheet = ss.getSheetByName(name);
  if (!sheet) sheet = ss.insertSheet(name);
  if (headers && headers.length) {
    var current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
    var mismatch = headers.some(function(h, i) { return String(current[i] || '') !== String(h); });
    if (mismatch) {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers.slice()]);
      sheet.setFrozenRows(1);
      sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold');
    }
  }
  if (hideSheet === true) {
    try { if (!sheet.isSheetHidden()) sheet.hideSheet(); } catch (ignore) {}
  }
  return sheet;
}

function SYSTEMLOG_isSystemSheetName_(name) {
  var n = String(name || '').trim();
  if (!n) return false;
  var explicit = {
    '_트리거현황':1, '_자동화상태':1, '_자동화전환기록':1,
    '_자동화재처리큐':1, '_자동화재처리이력':1, '_자동화장애상태':1,
    '_자동화장애알림로그':1, '_자동화유지관리':1, '_백업보존상태':1,
    '_점검일정동기화상태':1, '_점검일정동기화로그':1,
    '_선임완료동기화로그':1, '_정보통신유지보수이식로그':1,
    '_영업지원Discord상태':1,
    '발송파일저장큐_DB':1, '메일발송실패큐_DB':1, '발송파일로그':1,
    '메일발송로그':1, '발주메일발송큐':1, '발주메일발송로그':1,
    'KJ서류분류로그':1, 'KJ서류분류상태':1, 'KJ수행사_업로드동기화_LOG':1,
    'KJ공유복사로그':1, 'KJ공유복사요약':1, '고객사폴더_LOG':1, '고객사파일폴더인덱스':1, '사업자등록증OCR_진행로그':1,
    '사업자등록증_일괄복사_LOG':1, '계약서_일괄복사_LOG':1,
    '계약서_일괄복사_ROLLBACK_LOG':1, '권한테스트로그':1,
    '_중복삭제로그':1, 'TM_7월전메모분리로그':1,
    '__컨택이력_이관로그':1, '__컨택이력_이관미리보기':1,
    '장기미접촉_마스터이관로그':1
  };
  if (explicit[n]) return true;
  if (/(_LOG|_로그|로그|큐|기술|자동화.*상태|자동화.*기록|분류상태|보존상태)$/i.test(n)) return true;
  return false;
}

function SYSTEMLOG_previewMigration() {
  var source = SpreadsheetApp.openById(SYSTEM_LOG_CONFIG.masterSpreadsheetId);
  var target = SYSTEMLOG_getSpreadsheet_();
  var rows = source.getSheets().filter(function(s) {
    return SYSTEMLOG_isSystemSheetName_(s.getName());
  }).map(function(s) {
    var existing = target.getSheetByName(s.getName());
    return {
      sheetName: s.getName(),
      sourceRows: s.getLastRow(),
      sourceColumns: s.getLastColumn(),
      targetExists: !!existing,
      targetRows: existing ? existing.getLastRow() : 0
    };
  });
  var result = { status: 'PREVIEW', count: rows.length, sheets: rows };
  Logger.log(JSON.stringify(result));
  return result;
}

function SYSTEMLOG_executeMigration() {
  TRG_assertAutomationOwner_();
  var source = SpreadsheetApp.openById(SYSTEM_LOG_CONFIG.masterSpreadsheetId);
  var target = SYSTEMLOG_getSpreadsheet_();
  var candidates = source.getSheets().filter(function(s) {
    return SYSTEMLOG_isSystemSheetName_(s.getName());
  });
  var moved = [], merged = [], failed = [];

  candidates.forEach(function(sourceSheet) {
    var name = sourceSheet.getName();
    try {
      var targetSheet = target.getSheetByName(name);
      if (!targetSheet) {
        var copied = sourceSheet.copyTo(target);
        copied.setName(name);
        source.deleteSheet(sourceSheet);
        moved.push(name);
        return;
      }

      var sourceLastRow = sourceSheet.getLastRow();
      var sourceLastCol = sourceSheet.getLastColumn();
      if (sourceLastRow > 0 && sourceLastCol > 0) {
        var sourceValues = sourceSheet.getRange(1, 1, sourceLastRow, sourceLastCol).getValues();
        if (targetSheet.getLastRow() === 0) {
          targetSheet.getRange(1, 1, sourceValues.length, sourceLastCol).setValues(sourceValues);
        } else if (sourceValues.length > 1) {
          var targetWidth = Math.max(targetSheet.getLastColumn(), sourceLastCol);
          if (targetSheet.getMaxColumns() < targetWidth) {
            targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), targetWidth - targetSheet.getMaxColumns());
          }
          var rows = sourceValues.slice(1).map(function(r) {
            var out = r.slice();
            while (out.length < targetWidth) out.push('');
            return out;
          });
          if (rows.length) targetSheet.getRange(targetSheet.getLastRow() + 1, 1, rows.length, targetWidth).setValues(rows);
        }
      }
      source.deleteSheet(sourceSheet);
      merged.push(name);
    } catch (err) {
      failed.push({ sheetName: name, error: String(err && err.message || err) });
    }
  });

  var result = { status: failed.length ? 'PARTIAL' : 'SUCCESS', moved: moved, merged: merged, failed: failed };
  PropertiesService.getScriptProperties().setProperty(
    SYSTEM_LOG_CONFIG.migrationPropertyKey,
    JSON.stringify({ completedAt: new Date().toISOString(), result: result })
  );
  Logger.log(JSON.stringify(result));
  return result;
}

function SYSTEMLOG_archiveExpiredRowsNow() {
  TRG_assertAutomationOwner_();
  return SYSTEMLOG_archiveExpiredRows_({ force: true });
}

function SYSTEMLOG_archiveExpiredRowsDailySafe_() {
  try { return SYSTEMLOG_archiveExpiredRows_({ force: false }); }
  catch (err) { console.error('[SYSTEMLOG_archiveExpiredRowsDailySafe_] ' + String(err && err.stack || err)); return { status:'ERROR', error:String(err && err.message || err) }; }
}

function SYSTEMLOG_archiveExpiredRows_(options) {
  options = options || {};
  var props = PropertiesService.getScriptProperties();
  var today = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd');
  if (!options.force && props.getProperty(SYSTEM_LOG_CONFIG.archiveLastRunPropertyKey) === today) {
    return { status: 'SKIPPED_ALREADY_RAN_TODAY', date: today };
  }

  var ss = SYSTEMLOG_getSpreadsheet_();
  var folder = DriveApp.getFolderById(SYSTEM_LOG_CONFIG.csvArchiveFolderId);
  var cutoffMs = Date.now() - Number(SYSTEM_LOG_CONFIG.retentionDays || 3) * 24 * 60 * 60 * 1000;
  var archivedSheets = [], totalRows = 0, errors = [];

  ss.getSheets().forEach(function(sheet) {
    try {
      var lastRow = sheet.getLastRow(), lastCol = sheet.getLastColumn();
      if (lastRow < 2 || lastCol < 1) return;
      if (SYSTEMLOG_isSnapshotSheet_(sheet.getName())) return;
      var values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
      var timestampCol = SYSTEMLOG_findTimestampColumn_(values[0], values.slice(1, Math.min(values.length, 30)));
      if (timestampCol < 0) return;
      var statusCol = SYSTEMLOG_findStatusColumn_(values[0]);

      var expired = [], keep = [values[0]];
      for (var i = 1; i < values.length; i++) {
        var ms = SYSTEMLOG_toTimeMs_(values[i][timestampCol]);
        var archiveAllowed = SYSTEMLOG_isArchivableQueueRow_(sheet.getName(), values[i], statusCol);
        if (ms && ms < cutoffMs && archiveAllowed) expired.push(values[i]);
        else keep.push(values[i]);
      }
      if (!expired.length) return;

      var csvRows = [values[0]].concat(expired);
      var stamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyyMMdd_HHmmss');
      var safeName = sheet.getName().replace(/[\\/:*?"<>|]/g, '_');
      var blob = Utilities.newBlob('\uFEFF' + SYSTEMLOG_toCsv_(csvRows), 'text/csv', safeName + '_ARCHIVE_' + stamp + '.csv');
      var file = folder.createFile(blob);

      sheet.clearContents();
      sheet.getRange(1, 1, keep.length, lastCol).setValues(keep);
      sheet.setFrozenRows(1);
      archivedSheets.push({ sheetName: sheet.getName(), archivedRows: expired.length, fileId: file.getId(), fileName: file.getName() });
      totalRows += expired.length;
    } catch (err) {
      errors.push({ sheetName: sheet.getName(), error: String(err && err.message || err) });
    }
  });

  props.setProperty(SYSTEM_LOG_CONFIG.archiveLastRunPropertyKey, today);
  var result = { status: errors.length ? 'PARTIAL' : 'SUCCESS', cutoff: new Date(cutoffMs).toISOString(), totalRows: totalRows, sheets: archivedSheets, errors: errors };
  Logger.log(JSON.stringify(result));
  return result;
}


function SYSTEMLOG_isSnapshotSheet_(sheetName) {
  var snapshots = {
    '_트리거현황':1, '_자동화상태':1, '_자동화전환기록':1,
    '_자동화장애상태':1, '_자동화유지관리':1, '_백업보존상태':1,
    '_점검일정동기화상태':1, '_영업지원Discord상태':1,
    'KJ공유복사요약':1
  };
  return !!snapshots[String(sheetName || '')];
}

function SYSTEMLOG_findStatusColumn_(headers) {
  for (var i = 0; i < headers.length; i++) {
    var h = String(headers[i] || '').replace(/\s+/g, '');
    if (h === '상태' || h === '처리상태' || h === '최종상태' || h === '저장상태') return i;
  }
  return -1;
}

function SYSTEMLOG_isArchivableQueueRow_(sheetName, row, statusCol) {
  var name = String(sheetName || '');
  if (name.indexOf('큐') < 0 && name.indexOf('Queue') < 0 && name.indexOf('QUEUE') < 0) return true;
  if (statusCol < 0) return false;
  var status = String(row[statusCol] || '').trim().toUpperCase();
  // 아직 처리 가능하거나 사람이 확인해야 하는 작업은 3일이 지나도 보존한다.
  if (['PENDING','QUEUED','RETRY','RUNNING','WAITING','대기','재시도','실행중'].indexOf(status) >= 0) return false;
  return ['DONE','SUCCESS','COMPLETED','FAIL','FAILED','ERROR','SKIPPED','CANCELLED','완료','성공','실패','오류','취소'].indexOf(status) >= 0;
}

function SYSTEMLOG_findTimestampColumn_(headers, samples) {
  var preferred = /(기록일시|일시|날짜|시간|요청일시|감지|실행|완료|수정|생성|최근시도|최초요청)/;
  for (var i = 0; i < headers.length; i++) if (preferred.test(String(headers[i] || ''))) return i;
  for (var c = 0; c < headers.length; c++) {
    for (var r = 0; r < samples.length; r++) if (SYSTEMLOG_toTimeMs_(samples[r][c])) return c;
  }
  return -1;
}

function SYSTEMLOG_toTimeMs_(value) {
  if (value instanceof Date && !isNaN(value.getTime())) return value.getTime();
  if (typeof value === 'number' && isFinite(value) && value > 100000000000) return value;
  var text = String(value || '').trim();
  if (!text) return 0;
  var parsed = Date.parse(text);
  return isNaN(parsed) ? 0 : parsed;
}

function SYSTEMLOG_toCsv_(rows) {
  return rows.map(function(row) {
    return row.map(function(value) {
      var text;
      if (value instanceof Date && !isNaN(value.getTime())) text = Utilities.formatDate(value, Session.getScriptTimeZone() || 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      else text = String(value == null ? '' : value);
      return '"' + text.replace(/"/g, '""') + '"';
    }).join(',');
  }).join('\r\n');
}

function SYSTEMLOG_previewBackupMigration() {
  var oldFolder = DriveApp.getFolderById(SYSTEM_LOG_CONFIG.previousBackupFolderId);
  var files = oldFolder.getFiles();
  var rows = [];
  while (files.hasNext()) {
    var f = files.next();
    if (f.getName().indexOf(BACKUP_FILE_PREFIX) === 0) rows.push({ id:f.getId(), name:f.getName() });
  }
  var result = { status:'PREVIEW', count:rows.length, files:rows };
  Logger.log(JSON.stringify(result));
  return result;
}

function SYSTEMLOG_moveExistingBackupsNow() {
  TRG_assertAutomationOwner_();
  var oldFolder = DriveApp.getFolderById(SYSTEM_LOG_CONFIG.previousBackupFolderId);
  var newFolder = DriveApp.getFolderById(SYSTEM_LOG_CONFIG.backupFolderId);
  var files = oldFolder.getFiles();
  var moved = [], failed = [];
  while (files.hasNext()) {
    var file = files.next();
    if (file.getName().indexOf(BACKUP_FILE_PREFIX) !== 0) continue;
    try {
      file.moveTo(newFolder);
      moved.push({ id:file.getId(), name:file.getName() });
    } catch (err) {
      failed.push({ id:file.getId(), name:file.getName(), error:String(err && err.message || err) });
    }
  }
  var result = { status:failed.length ? 'PARTIAL':'SUCCESS', moved:moved.length, files:moved, failed:failed };
  PropertiesService.getScriptProperties().setProperty(SYSTEM_LOG_CONFIG.backupMigrationPropertyKey, JSON.stringify({completedAt:new Date().toISOString(), result:result}));
  Logger.log(JSON.stringify(result));
  return result;
}

function SYSTEMLOG_executePhase25MigrationAll() {
  return {
    systemSheets: SYSTEMLOG_executeMigration(),
    backups: SYSTEMLOG_moveExistingBackupsNow(),
    archive: SYSTEMLOG_archiveExpiredRowsNow()
  };
}

function SYSTEMLOG_preflightPhase25() {
  TRG_assertAutomationOwner_();
  var checks = [];
  function check(label, fn) {
    try { checks.push({ label:label, ok:true, detail:String(fn() || '') }); }
    catch (err) { checks.push({ label:label, ok:false, detail:String(err && err.message || err) }); }
  }
  check('영업관리대장', function() { return SpreadsheetApp.openById(SYSTEM_LOG_CONFIG.masterSpreadsheetId).getName(); });
  check('SYSTEM_LOG 스프레드시트', function() { return SYSTEMLOG_getSpreadsheet_().getName(); });
  check('CSV 아카이브 폴더', function() { return DriveApp.getFolderById(SYSTEM_LOG_CONFIG.csvArchiveFolderId).getName(); });
  check('신규 백업 폴더', function() { return DriveApp.getFolderById(SYSTEM_LOG_CONFIG.backupFolderId).getName(); });
  check('기존 백업 폴더', function() { return DriveApp.getFolderById(SYSTEM_LOG_CONFIG.previousBackupFolderId).getName(); });
  var result = { status:checks.every(function(x){return x.ok;})?'SUCCESS':'FAILED', checks:checks };
  Logger.log(JSON.stringify(result));
  if (result.status !== 'SUCCESS') throw new Error('Phase25 권한 사전점검 실패: ' + JSON.stringify(result));
  return result;
}
