// =============================================================================
// 최초 권한 인증 + 권한 테스트 함수
// =============================================================================
// 실행 방법:
// Apps Script 편집기 → 함수 선택: authWarmupAndScopeTest → 실행
//
// 테스트하는 권한:
// - spreadsheets
// - spreadsheets.currentonly
// - drive
// - script.external_request
// - script.scriptapp
// - script.container.ui
// - documents
// - Advanced Drive API v3: Drive.Files.get / Drive.Files.create / Drive.Files.delete
// - Advanced Docs API v1: Docs.Documents.get
// =============================================================================

function authWarmupAndScopeTest() {
  const startedAt = new Date();
  const results = [];

  function ok(name, detail) {
    results.push([new Date(), name, 'OK', detail || '']);
  }

  function fail(name, err) {
    results.push([
      new Date(),
      name,
      'FAIL',
      String(err && err.stack ? err.stack : err)
    ]);
  }

  let activeSs = null;
  let logSheet = null;
  let tempDocId = '';
  let tempDriveFileId = '';
  let tempTrigger = null;

  // ---------------------------------------------------------------------------
  // 1. SpreadsheetApp / current spreadsheet / UI
  // ---------------------------------------------------------------------------
  try {
    activeSs = SpreadsheetApp.getActiveSpreadsheet();
    if (!activeSs) {
      throw new Error('활성 스프레드시트를 찾을 수 없습니다. 컨테이너 바운드 스크립트인지 확인하세요.');
    }

    const name = activeSs.getName();
    const id = activeSs.getId();

    ok('SpreadsheetApp.getActiveSpreadsheet', `name=${name} / id=${id}`);
  } catch (err) {
    fail('SpreadsheetApp.getActiveSpreadsheet', err);
  }

  try {
    if (!activeSs) throw new Error('activeSs 없음');

    logSheet = getOrCreateAuthTestLogSheet_(activeSs);
    logSheet.appendRow([new Date(), '권한 테스트 시작', startedAt]);

    ok('Spreadsheet write/read', `logSheet=${logSheet.getName()}`);
  } catch (err) {
    fail('Spreadsheet write/read', err);
  }

  try {
    // container.ui scope 확인용.
    // 편집기 직접 실행에서는 alert까지 띄우면 불편하므로 getUi 객체 접근만 수행.
    const ui = SpreadsheetApp.getUi();
    ok('script.container.ui', 'SpreadsheetApp.getUi() 접근 성공');
  } catch (err) {
    fail('script.container.ui', err);
  }

  // ---------------------------------------------------------------------------
  // 2. openById 테스트: SOURCE_SS_ID / TARGET_SS_ID
  // ---------------------------------------------------------------------------
  try {
    const sourceId = getConfigIdForAuthTest_('SOURCE_SS_ID');
    if (sourceId) {
      const sourceSs = SpreadsheetApp.openById(sourceId);
      ok('SpreadsheetApp.openById SOURCE_SS_ID', `${sourceSs.getName()} / ${sourceId}`);
    } else {
      ok('SpreadsheetApp.openById SOURCE_SS_ID', 'MAIL_AUTO_CONFIG.SOURCE_SS_ID 없음. 스킵');
    }
  } catch (err) {
    fail('SpreadsheetApp.openById SOURCE_SS_ID', err);
  }

  try {
    const targetId = getConfigIdForAuthTest_('TARGET_SS_ID');
    if (targetId) {
      const targetSs = SpreadsheetApp.openById(targetId);
      ok('SpreadsheetApp.openById TARGET_SS_ID', `${targetSs.getName()} / ${targetId}`);
    } else {
      ok('SpreadsheetApp.openById TARGET_SS_ID', 'MAIL_AUTO_CONFIG.TARGET_SS_ID 없음. 스킵');
    }
  } catch (err) {
    fail('SpreadsheetApp.openById TARGET_SS_ID', err);
  }

  // ---------------------------------------------------------------------------
  // 3. DriveApp 권한 테스트
  // ---------------------------------------------------------------------------
  try {
    const blob = Utilities.newBlob(
      'auth test',
      'text/plain',
      `auth_test_${Date.now()}.txt`
    );

    const file = DriveApp.createFile(blob);
    const fileId = file.getId();

    file.setSharing(DriveApp.Access.PRIVATE, DriveApp.Permission.VIEW);
    file.setTrashed(true);

    ok('DriveApp create/share/trash', `tempFileId=${fileId}`);
  } catch (err) {
    fail('DriveApp create/share/trash', err);
  }

  // ---------------------------------------------------------------------------
  // 4. UrlFetchApp 권한 테스트
  // ---------------------------------------------------------------------------
  try {
    const res = UrlFetchApp.fetch('https://www.google.com/generate_204', {
      muteHttpExceptions: true
    });

    ok('UrlFetchApp external_request', `HTTP ${res.getResponseCode()}`);
  } catch (err) {
    fail('UrlFetchApp external_request', err);
  }

  // ---------------------------------------------------------------------------
  // 5. ScriptApp 권한 테스트: 임시 트리거 생성 후 즉시 삭제
  // ---------------------------------------------------------------------------
  try {
    tempTrigger = ScriptApp.newTrigger('dummyAuthTriggerTarget_')
      .timeBased()
      .after(60 * 60 * 1000)
      .create();

    const triggerId = tempTrigger.getUniqueId();

    ScriptApp.deleteTrigger(tempTrigger);
    tempTrigger = null;

    ok('ScriptApp trigger create/delete', `triggerId=${triggerId}`);
  } catch (err) {
    fail('ScriptApp trigger create/delete', err);
  }

  // ---------------------------------------------------------------------------
  // 6. DocumentApp 권한 테스트
  // ---------------------------------------------------------------------------
  try {
    const doc = DocumentApp.create(`AUTH_TEST_DOC_${Date.now()}`);
    tempDocId = doc.getId();

    doc.getBody().appendParagraph('권한 테스트 문서입니다.');
    doc.saveAndClose();

    ok('DocumentApp create/edit', `docId=${tempDocId}`);
  } catch (err) {
    fail('DocumentApp create/edit', err);
  }

  // ---------------------------------------------------------------------------
  // 7. Advanced Docs API v1 테스트
  // ---------------------------------------------------------------------------
  try {
    if (typeof Docs === 'undefined') {
      throw new Error('Docs is not defined. 고급 Google 서비스에서 Docs API v1을 추가하세요.');
    }

    if (!tempDocId) {
      throw new Error('테스트용 Google Docs 문서가 생성되지 않아 Docs API 테스트를 스킵할 수 없습니다.');
    }

    const docMeta = Docs.Documents.get(tempDocId);
    ok('Advanced Docs API v1', `title=${docMeta.title || ''}`);
  } catch (err) {
    fail('Advanced Docs API v1', err);
  }

  // ---------------------------------------------------------------------------
  // 8. Advanced Drive API v3 테스트
  // ---------------------------------------------------------------------------
  try {
    if (typeof Drive === 'undefined') {
      throw new Error('Drive is not defined. 고급 Google 서비스에서 Drive API v3을 추가하세요.');
    }

    const activeFileMeta = Drive.Files.get(activeSs.getId(), {
      fields: 'id,name,mimeType'
    });

    ok('Advanced Drive API v3 - Files.get', `name=${activeFileMeta.name} / id=${activeFileMeta.id}`);
  } catch (err) {
    fail('Advanced Drive API v3 - Files.get', err);
  }

  try {
    if (typeof Drive === 'undefined') {
      throw new Error('Drive is not defined. 고급 Google 서비스에서 Drive API v3을 추가하세요.');
    }

    const blob = Utilities.newBlob(
      'advanced drive auth test',
      'text/plain',
      `advanced_drive_auth_test_${Date.now()}.txt`
    );

    const metadata = {
      name: blob.getName(),
      mimeType: blob.getContentType()
    };

    const created = Drive.Files.create(metadata, blob, {
      fields: 'id,name'
    });

    tempDriveFileId = created.id;

    ok('Advanced Drive API v3 - Files.create', `name=${created.name} / id=${created.id}`);
  } catch (err) {
    fail('Advanced Drive API v3 - Files.create', err);
  }

  try {
    if (typeof Drive === 'undefined') {
      throw new Error('Drive is not defined. 고급 Google 서비스에서 Drive API v3을 추가하세요.');
    }

    if (tempDriveFileId) {
      Drive.Files.remove(tempDriveFileId);
      ok('Advanced Drive API v3 - Files.remove', `removed=${tempDriveFileId}`);
      tempDriveFileId = '';
    } else {
      ok('Advanced Drive API v3 - Files.remove', '삭제할 임시 파일 없음. 스킵');
    }
  } catch (err) {
    fail('Advanced Drive API v3 - Files.remove', err);
  }

  // ---------------------------------------------------------------------------
  // 9. 정리: 임시 Docs 파일 휴지통 이동
  // ---------------------------------------------------------------------------
  try {
    if (tempDocId) {
      DriveApp.getFileById(tempDocId).setTrashed(true);
      ok('Cleanup temp doc', `trashed=${tempDocId}`);
    }
  } catch (err) {
    fail('Cleanup temp doc', err);
  }

  // ---------------------------------------------------------------------------
  // 10. 결과 기록
  // ---------------------------------------------------------------------------
  try {
    if (!activeSs) activeSs = SpreadsheetApp.getActiveSpreadsheet();
    logSheet = getOrCreateAuthTestLogSheet_(activeSs);

    if (results.length) {
      logSheet
        .getRange(logSheet.getLastRow() + 1, 1, results.length, results[0].length)
        .setValues(results);
    }

    logSheet.autoResizeColumns(1, 4);
    SpreadsheetApp.flush();
  } catch (err) {
    Logger.log('권한 테스트 결과 로그 기록 실패: ' + err);
  }

  const failed = results.filter(r => r[2] === 'FAIL');

  if (failed.length > 0) {
    SpreadsheetApp.getUi().alert(
      `권한 테스트 완료: 실패 ${failed.length}건\n\n` +
      `자세한 내용은 "권한테스트로그" 시트를 확인하세요.`
    );
  } else {
    SpreadsheetApp.getUi().alert(
      '권한 테스트 완료: 전체 OK\n\n' +
      '최초 인증도 완료된 상태입니다.'
    );
  }
}


// =============================================================================
// 권한 테스트 보조 함수
// =============================================================================

function getOrCreateAuthTestLogSheet_(ss) {
  const name = '권한테스트로그';
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
    sheet.getRange(1, 1, 1, 4).setValues([[
      '기록시각',
      '테스트항목',
      '결과',
      '상세'
    ]]);

    sheet.setFrozenRows(1);
    sheet.getRange(1, 1, 1, 4)
      .setFontWeight('bold')
      .setBackground('#1a237e')
      .setFontColor('#ffffff');
  }

  return sheet;
}


function getConfigIdForAuthTest_(key) {
  try {
    if (typeof MAIL_AUTO_CONFIG !== 'undefined' && MAIL_AUTO_CONFIG[key]) {
      return MAIL_AUTO_CONFIG[key];
    }
  } catch (err) {
    // 무시
  }

  return '';
}

// =============================================================================
// 설치형 onEdit 트리거 설치
// =============================================================================

function installMailRequestTrigger() {
  const ss = SpreadsheetApp.openById(MAIL_AUTO_CONFIG.SOURCE_SS_ID);

  const triggers = ScriptApp.getProjectTriggers();

  triggers.forEach(t => {
    try {
      if (t.getHandlerFunction && t.getHandlerFunction() === 'onMailRequestEdit') {
        ScriptApp.deleteTrigger(t);
      }
    } catch (err) {
      Logger.log('기존 onMailRequestEdit 트리거 삭제 실패, 계속 진행: ' + err);
    }
  });

  ScriptApp.newTrigger('onMailRequestEdit')
    .forSpreadsheet(ss)
    .onEdit()
    .create();

  SpreadsheetApp.getUi().alert('메일발송요청 체크박스용 설치형 onEdit 트리거를 설치했습니다.');
}


// ScriptApp 트리거 생성 테스트용 더미 함수
function dummyAuthTriggerTarget_() {
  // 권한 테스트용 빈 함수
}

function forceReauthorize() {
  ScriptApp.invalidateAuth();
}
