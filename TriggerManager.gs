/****************************************************
 * TriggerManager.gs
 * 영업관리대장 설치형 트리거 중앙관리 - 2단계
 *
 * 운영 원칙:
 * - 설치형 트리거 소유 계정은 bang@s1samsung.com 하나로 고정
 * - 현재 단계에서는 현황 조회 / 계획 검증 / 전체 삭제만 제공
 * - 통합 onEdit/onChange 핸들러는 구현 완료
 * - 정식 13개 재설치는 5분 핵심 동기화 파이프라인 구현 후 활성화
 * - 단순 트리거(onOpen/onEdit/onSelectionChange)와 웹앱(doGet/doPost)은
 *   ScriptApp 설치형 트리거가 아니므로 이 관리 대상에서 제외
 ****************************************************/

var TRG_MANAGER_CONFIG = Object.freeze({
  automationOwnerEmail: 'bang@s1samsung.com',
  statusSheetName: '_트리거현황',
  planVersion: '2026-07-19-PHASE2',
  installEnabled: false,
  timezone: 'Asia/Seoul'
});


/****************************************************
 * 공개 실행 함수
 ****************************************************/

/**
 * 정식 트리거 계획과 현재 설치 현황을 로그로 미리 본다.
 * 시트는 수정하지 않는다.
 */
function TRG_previewCanonicalPlan() {
  TRG_assertAutomationOwner_();

  var snapshot = TRG_buildStatusSnapshot_();
  Logger.log(JSON.stringify(snapshot.summary, null, 2));
  Logger.log(JSON.stringify(snapshot.planRows, null, 2));

  return snapshot;
}


/**
 * 현재/예정 트리거 현황을 `_트리거현황` 시트에 기록한다.
 * 기존 업무 데이터 시트와 자동화 로직은 변경하지 않는다.
 */
function TRG_showTriggerStatus() {
  TRG_assertAutomationOwner_();

  var snapshot = TRG_buildStatusSnapshot_();
  var sheet = TRG_writeStatusSheet_(snapshot);

  SpreadsheetApp.getActiveSpreadsheet().setActiveSheet(sheet);
  SpreadsheetApp.getActiveSpreadsheet().toast(
    '트리거 현황을 갱신했습니다. 설치 ' + snapshot.summary.installedTriggerCount +
      '개 / 정식계획 ' + snapshot.summary.canonicalPlannedTriggerCount + '개',
    '트리거 중앙관리',
    7
  );

  return snapshot.summary;
}


/**
 * 정식 계획 대비 현재 설치 상태를 검증하고 상태 시트를 갱신한다.
 */
function TRG_verifyCanonicalTriggers() {
  TRG_assertAutomationOwner_();

  var snapshot = TRG_buildStatusSnapshot_();
  TRG_writeStatusSheet_(snapshot);

  var message = [
    '현재 설치형 트리거: ' + snapshot.summary.installedTriggerCount + '개',
    '정식 계획 트리거: ' + snapshot.summary.canonicalPlannedTriggerCount + '개',
    '계획 일치: ' + snapshot.summary.canonicalMatchedTriggerCount + '개',
    '계획 누락: ' + snapshot.summary.canonicalMissingTriggerCount + '개',
    '계획 초과/중복: ' + snapshot.summary.canonicalExcessTriggerCount + '개',
    '고아 핸들러 트리거: ' + snapshot.summary.orphanTriggerCount + '개',
    '미분류 트리거: ' + snapshot.summary.unknownTriggerCount + '개',
    '',
    '상세 내용은 ' + TRG_MANAGER_CONFIG.statusSheetName + ' 시트를 확인하세요.'
  ].join('\n');

  SpreadsheetApp.getUi().alert('트리거 검증 결과', message, SpreadsheetApp.getUi().ButtonSet.OK);
  return snapshot.summary;
}


/**
 * 현재 계정이 소유한 이 프로젝트의 설치형 트리거를 전부 삭제한다.
 *
 * 주의:
 * - bang@s1samsung.com에서만 실행 가능
 * - 단순 onEdit/onOpen/onSelectionChange 및 doGet/doPost는 삭제되지 않음
 * - 다른 계정 소유 트리거는 Apps Script 제약상 조회/삭제할 수 없음
 */
function TRG_removeAllInstallableTriggers() {
  TRG_assertAutomationOwner_();

  var triggers = ScriptApp.getProjectTriggers();
  var ui = SpreadsheetApp.getUi();

  if (triggers.length === 0) {
    ui.alert(
      '삭제할 트리거 없음',
      '현재 계정 소유의 설치형 트리거가 없습니다.',
      ui.ButtonSet.OK
    );
    TRG_showTriggerStatus();
    return {
      deletedCount: 0,
      failedCount: 0,
      failures: []
    };
  }

  var previewLines = triggers.slice(0, 15).map(function(trigger, index) {
    return (index + 1) + '. ' + TRG_describeTrigger_(trigger);
  });

  if (triggers.length > previewLines.length) {
    previewLines.push('... 외 ' + (triggers.length - previewLines.length) + '개');
  }

  var response = ui.alert(
    '설치형 트리거 전체 삭제',
    [
      '실행 계정: ' + TRG_getEffectiveUserEmail_(),
      '삭제 대상: ' + triggers.length + '개',
      '',
      previewLines.join('\n'),
      '',
      '단순 onEdit/onOpen/onSelectionChange와 웹앱 doGet/doPost는 영향을 받지 않습니다.',
      '정말 현재 계정 소유 설치형 트리거를 전부 삭제하시겠습니까?'
    ].join('\n'),
    ui.ButtonSet.YES_NO
  );

  if (response !== ui.Button.YES) {
    ui.alert('취소됨', '트리거를 삭제하지 않았습니다.', ui.ButtonSet.OK);
    return {
      deletedCount: 0,
      failedCount: 0,
      cancelled: true,
      failures: []
    };
  }

  var result = TRG_deleteTriggers_(triggers);
  TRG_writeStatusSheet_(TRG_buildStatusSnapshot_());

  ui.alert(
    '삭제 완료',
    '삭제 성공: ' + result.deletedCount + '개\n' +
      '삭제 실패: ' + result.failedCount + '개\n\n' +
      '정식 재설치 기능은 아직 잠겨 있습니다. 후속 통합 작업이 완료되기 전에는 설치하지 않습니다.',
    ui.ButtonSet.OK
  );

  return result;
}


/**
 * 2단계 안전장치.
 * 5분 핵심 동기화 파이프라인과 정식 설치 로직을 완성하기 전까지 설치를 차단한다.
 */
function TRG_installCanonicalTriggers() {
  TRG_assertAutomationOwner_();

  if (!TRG_MANAGER_CONFIG.installEnabled) {
    throw new Error(
      '2단계에서는 정식 트리거 설치가 아직 잠겨 있습니다. ' +
      '영업관리대장 통합 onEdit/onChange는 구현됐지만, 5분 핵심 동기화 파이프라인을 구현한 뒤 활성화해야 합니다.'
    );
  }

  throw new Error('정식 트리거 설치 로직은 아직 구현되지 않았습니다.');
}


/****************************************************
 * 정식 13개 트리거 계획
 ****************************************************/

/**
 * 향후 최종 설치할 정식 계획.
 *
 * 정의 행은 12개지만 backupSalesLedger가 하루 2회이므로
 * expectedCount 합계는 13개다.
 *
 * Apps Script Trigger 객체는 시간 트리거의 실제 분/시각 설정을
 * 다시 읽어오는 API를 제공하지 않는다. 따라서 현재 검증에서는
 * 핸들러 + 이벤트 유형 + 대상 파일 + 개수까지만 자동 비교한다.
 */
function TRG_getCanonicalPlan_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  if (!ss) {
    throw new Error('바인딩된 영업관리대장 스프레드시트를 찾을 수 없습니다.');
  }

  var mainSpreadsheetId = ss.getId();
  var vendorFiles = TRG_getVendorSpreadsheetIds_();

  return [
    {
      key: 'MAIN_EDIT_DISPATCHER',
      category: '영업관리대장 이벤트',
      handler: 'AUTOMATION_handleSalesLedgerEdit',
      eventType: 'ON_EDIT',
      sourceType: 'SPREADSHEETS',
      sourceId: mainSpreadsheetId,
      targetLabel: '영업관리대장',
      schedule: '편집 시',
      expectedCount: 1,
      implementationStatus: '2단계 구현 완료 / 정식 설치 대기',
      note: 'sb01/sb02/sb03/고객폴더용 영업관리대장 onEdit 중앙 디스패처 구현 완료'
    },
    {
      key: 'MAIN_CHANGE_DISPATCHER',
      category: '영업관리대장 이벤트',
      handler: 'AUTOMATION_handleSalesLedgerChange',
      eventType: 'ON_CHANGE',
      sourceType: 'SPREADSHEETS',
      sourceId: mainSpreadsheetId,
      targetLabel: '영업관리대장',
      schedule: '구조 변경 시',
      expectedCount: 1,
      implementationStatus: '2단계 구현 완료 / 정식 설치 대기',
      note: '구조 변경 시 직접 전체동기화하지 않고 전체보정 필요 플래그 기록'
    },
    {
      key: 'VENDOR_KJ_EDIT',
      category: '수행사 파일 이벤트',
      handler: 'installedOnEdit',
      eventType: 'ON_EDIT',
      sourceType: 'SPREADSHEETS',
      sourceId: vendorFiles.KJ,
      targetLabel: 'KJ 고객관리',
      schedule: '편집 시',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: 'KJ 수행사 고객관리 → 수주확정 역동기화'
    },
    {
      key: 'VENDOR_ILSHIN_EDIT',
      category: '수행사 파일 이벤트',
      handler: 'installedOnEdit',
      eventType: 'ON_EDIT',
      sourceType: 'SPREADSHEETS',
      sourceId: vendorFiles['일신'],
      targetLabel: '일신 고객관리',
      schedule: '편집 시',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: '일신 수행사 고객관리 → 수주확정 역동기화'
    },
    {
      key: 'CORE_DATA_SYNC',
      category: '핵심 데이터 동기화',
      handler: 'AUTOMATION_runCoreDataSyncPipeline',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '5분',
      expectedCount: 1,
      implementationStatus: '후속 3단계 구현 예정',
      note: '마스터 → 수주확정 → 수행사/정보통신유지보수 순서 고정'
    },
    {
      key: 'MAIL_ARCHIVE_QUEUE',
      category: '메일 자동화',
      handler: 'processDeferredSentFileArchiveQueueV94',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '5분',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: '메일 발송파일 저장 재처리 큐'
    },
    {
      key: 'DISCORD_SUPPORT_ALERT',
      category: '영업지원 알림',
      handler: 'checkSalesSupportNewValues',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '1분',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: 'Discord 영업지원요청 신규값 알림'
    },
    {
      key: 'KJ_RECENT_CLASSIFIER',
      category: 'KJ 서류 분류',
      handler: 'classifyKjDocumentsNow',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '1분',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용 / 후속 경량화 예정',
      note: '최근 파일 중심 분류로 역할 축소 예정'
    },
    {
      key: 'KJ_SAFETY_FULL_SCAN',
      category: 'KJ 서류 분류',
      handler: 'classifyKjDocumentsSafetyFullScan',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '6시간',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: '폴더 보정 및 전체 누락 안전 점검'
    },
    {
      key: 'KJ_VENDOR_UPLOAD_SYNC',
      category: 'KJ 업로드 동기화',
      handler: 'KJUS_runVendorUploadSync',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '30분',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: '수행사 업로드 파일 동기화'
    },
    {
      key: 'SALES_LEDGER_BACKUP',
      category: '백업',
      handler: 'backupSalesLedger',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '매일 12:30대 / 18:00대',
      expectedCount: 2,
      implementationStatus: '기존 핸들러 사용',
      note: '같은 핸들러의 시간 트리거 2개'
    },
    {
      key: 'SENT_FILE_HISTORY',
      category: '메일 자동화',
      handler: 'syncSentFileFolderHistoryDaily',
      eventType: 'CLOCK',
      sourceType: 'CLOCK',
      sourceId: '',
      targetLabel: '프로젝트',
      schedule: '매일 19시대',
      expectedCount: 1,
      implementationStatus: '기존 핸들러 사용',
      note: '발송파일 폴더 이력 일일 동기화'
    }
  ];
}


/**
 * 현재 구조에서 발견된 구형/개별 설치 핸들러 목록.
 * 현황표에서 정식계획 외 트리거를 설명하기 위한 분류용이다.
 */
function TRG_getKnownLegacyHandlers_() {
  return {
    handleContractMasterSyncOnEdit: 'sb01 마스터↔수주확정 개별 onEdit',
    handleContractMasterSyncEvery5Minutes: 'sb01 개별 5분 전체동기화',
    installedOnEdit: 'sb02 공용 onEdit; 수행사 파일 2개는 정식계획 유지, 영업관리대장용 1개는 향후 통합 대상',
    syncAllFromMasterTimeDriven: 'sb02 개별 5분 수행사 동기화',
    ITMAINT_onEditSync_2026: 'sb03 개별 onEdit',
    ITMAINT_onChangeSync_2026: 'sb03 개별 onChange',
    ITMAINT_timeDrivenSync_2026: 'sb03 개별 5분 전체동기화',
    customerFolderInstallableOnEdit: '고객사 폴더 개별 onEdit',
    onMailRequestEdit: 'authWarmup에서 설치하지만 소스에 핸들러가 없는 고아 트리거',
    onEditSync_정보통신유지보수: 'sb03 구버전 onEdit',
    onChangeSync_정보통신유지보수: 'sb03 구버전 onChange',
    timeDrivenSync_정보통신유지보수: 'sb03 구버전 시간 트리거',
    dummyAuthTriggerTarget_: '권한 사전인증용 임시 트리거'
  };
}


/****************************************************
 * 현황 수집 및 비교
 ****************************************************/

function TRG_buildStatusSnapshot_() {
  var plan = TRG_getCanonicalPlan_();
  var installed = ScriptApp.getProjectTriggers();
  var installedRows = installed.map(TRG_triggerToRow_);
  var legacyHandlers = TRG_getKnownLegacyHandlers_();
  var signatureCounts = {};

  installedRows.forEach(function(row) {
    var signature = TRG_makeSignature_(
      row.handler,
      row.eventType,
      row.sourceType,
      row.sourceId
    );
    signatureCounts[signature] = (signatureCounts[signature] || 0) + 1;
  });

  var plannedTotal = 0;
  var matchedTotal = 0;
  var missingTotal = 0;
  var excessTotal = 0;

  var planRows = plan.map(function(item) {
    var signature = TRG_makeSignature_(
      item.handler,
      item.eventType,
      item.sourceType,
      item.sourceId
    );
    var actualCount = signatureCounts[signature] || 0;
    var matchedCount = Math.min(actualCount, item.expectedCount);
    var missingCount = Math.max(0, item.expectedCount - actualCount);
    var excessCount = Math.max(0, actualCount - item.expectedCount);
    var handlerExists = TRG_handlerExists_(item.handler);
    var status;

    plannedTotal += item.expectedCount;
    matchedTotal += matchedCount;
    missingTotal += missingCount;
    excessTotal += excessCount;

    if (!handlerExists) {
      status = '핸들러 미구현';
    } else if (missingCount > 0) {
      status = '트리거 누락';
    } else if (excessCount > 0) {
      status = '중복/초과';
    } else {
      status = '계획 일치';
    }

    return {
      key: item.key,
      category: item.category,
      handler: item.handler,
      handlerExists: handlerExists,
      eventType: item.eventType,
      sourceType: item.sourceType,
      sourceId: item.sourceId,
      targetLabel: item.targetLabel,
      schedule: item.schedule,
      expectedCount: item.expectedCount,
      actualCount: actualCount,
      matchedCount: matchedCount,
      missingCount: missingCount,
      excessCount: excessCount,
      implementationStatus: item.implementationStatus,
      status: status,
      note: item.note
    };
  });

  var canonicalSignatures = {};
  plan.forEach(function(item) {
    canonicalSignatures[TRG_makeSignature_(
      item.handler,
      item.eventType,
      item.sourceType,
      item.sourceId
    )] = true;
  });

  var orphanTriggerCount = 0;
  var unknownTriggerCount = 0;
  var legacyTriggerCount = 0;

  installedRows.forEach(function(row) {
    var signature = TRG_makeSignature_(
      row.handler,
      row.eventType,
      row.sourceType,
      row.sourceId
    );

    row.handlerExists = TRG_handlerExists_(row.handler);
    row.legacyDescription = legacyHandlers[row.handler] || '';

    if (!row.handlerExists) {
      row.classification = '고아 핸들러';
      orphanTriggerCount += 1;
    } else if (canonicalSignatures[signature]) {
      row.classification = '정식 계획 일치 후보';
    } else if (legacyHandlers[row.handler]) {
      row.classification = '구형/개별 트리거';
      legacyTriggerCount += 1;
    } else {
      row.classification = '미분류';
      unknownTriggerCount += 1;
    }
  });

  return {
    generatedAt: new Date(),
    ownerEmail: TRG_MANAGER_CONFIG.automationOwnerEmail,
    effectiveUserEmail: TRG_getEffectiveUserEmail_(),
    planVersion: TRG_MANAGER_CONFIG.planVersion,
    installEnabled: TRG_MANAGER_CONFIG.installEnabled,
    summary: {
      ownerEmail: TRG_MANAGER_CONFIG.automationOwnerEmail,
      effectiveUserEmail: TRG_getEffectiveUserEmail_(),
      installedTriggerCount: installedRows.length,
      canonicalPlannedTriggerCount: plannedTotal,
      canonicalMatchedTriggerCount: matchedTotal,
      canonicalMissingTriggerCount: missingTotal,
      canonicalExcessTriggerCount: excessTotal,
      orphanTriggerCount: orphanTriggerCount,
      legacyTriggerCount: legacyTriggerCount,
      unknownTriggerCount: unknownTriggerCount,
      installEnabled: TRG_MANAGER_CONFIG.installEnabled,
      planVersion: TRG_MANAGER_CONFIG.planVersion
    },
    planRows: planRows,
    installedRows: installedRows
  };
}


function TRG_triggerToRow_(trigger) {
  var sourceId = '';
  var uniqueId = '';

  try {
    sourceId = trigger.getTriggerSourceId() || '';
  } catch (ignoreSourceIdError) {
    sourceId = '';
  }

  try {
    uniqueId = trigger.getUniqueId() || '';
  } catch (ignoreUniqueIdError) {
    uniqueId = '';
  }

  return {
    handler: trigger.getHandlerFunction(),
    eventType: TRG_enumToString_(trigger.getEventType()),
    sourceType: TRG_enumToString_(trigger.getTriggerSource()),
    sourceId: sourceId,
    uniqueId: uniqueId,
    handlerExists: false,
    classification: '',
    legacyDescription: ''
  };
}


/****************************************************
 * 상태 시트 출력
 ****************************************************/

function TRG_writeStatusSheet_(snapshot) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(TRG_MANAGER_CONFIG.statusSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(TRG_MANAGER_CONFIG.statusSheetName);
  }

  sheet.clear();
  sheet.setFrozenRows(1);

  var tz = Session.getScriptTimeZone() || TRG_MANAGER_CONFIG.timezone;
  var generatedAtText = Utilities.formatDate(snapshot.generatedAt, tz, 'yyyy-MM-dd HH:mm:ss');

  var summaryValues = [
    ['트리거 중앙관리 1단계', '값'],
    ['점검 시각', generatedAtText],
    ['자동화 소유 계정', snapshot.ownerEmail],
    ['현재 실행 계정', snapshot.effectiveUserEmail],
    ['계획 버전', snapshot.planVersion],
    ['정식 설치 기능', snapshot.installEnabled ? '활성' : '잠금'],
    ['현재 설치형 트리거', snapshot.summary.installedTriggerCount],
    ['정식 계획 트리거', snapshot.summary.canonicalPlannedTriggerCount],
    ['계획 일치', snapshot.summary.canonicalMatchedTriggerCount],
    ['계획 누락', snapshot.summary.canonicalMissingTriggerCount],
    ['계획 초과/중복', snapshot.summary.canonicalExcessTriggerCount],
    ['고아 핸들러', snapshot.summary.orphanTriggerCount],
    ['구형/개별 트리거', snapshot.summary.legacyTriggerCount],
    ['미분류 트리거', snapshot.summary.unknownTriggerCount],
    ['검증 한계', '시간 트리거의 실제 분/시각은 Apps Script API로 역조회 불가']
  ];

  sheet.getRange(1, 1, summaryValues.length, 2).setValues(summaryValues);
  sheet.getRange(1, 1, 1, 2).setFontWeight('bold').setBackground('#d9ead3');
  sheet.getRange(1, 1, summaryValues.length, 1).setFontWeight('bold');

  var planStartRow = summaryValues.length + 3;
  var planHeaders = [
    '계획키', '구분', '핸들러', '핸들러 존재', '이벤트', '대상유형',
    '대상', '대상 ID', '예상주기', '예정개수', '현재일치개수',
    '누락', '초과', '상태', '구현상태', '비고'
  ];

  sheet.getRange(planStartRow, 1, 1, planHeaders.length).setValues([planHeaders]);
  sheet.getRange(planStartRow, 1, 1, planHeaders.length)
    .setFontWeight('bold')
    .setBackground('#cfe2f3');

  if (snapshot.planRows.length > 0) {
    var planValues = snapshot.planRows.map(function(row) {
      return [
        row.key,
        row.category,
        row.handler,
        row.handlerExists ? '예' : '아니오',
        row.eventType,
        row.sourceType,
        row.targetLabel,
        row.sourceId,
        row.schedule,
        row.expectedCount,
        row.actualCount,
        row.missingCount,
        row.excessCount,
        row.status,
        row.implementationStatus,
        row.note
      ];
    });

    sheet.getRange(planStartRow + 1, 1, planValues.length, planHeaders.length).setValues(planValues);
  }

  var installedStartRow = planStartRow + snapshot.planRows.length + 3;
  var installedHeaders = [
    '번호', '핸들러', '핸들러 존재', '이벤트', '소스유형', '소스 ID',
    '분류', '구형 설명', '트리거 고유 ID'
  ];

  sheet.getRange(installedStartRow, 1, 1, installedHeaders.length).setValues([installedHeaders]);
  sheet.getRange(installedStartRow, 1, 1, installedHeaders.length)
    .setFontWeight('bold')
    .setBackground('#fce5cd');

  if (snapshot.installedRows.length > 0) {
    var installedValues = snapshot.installedRows.map(function(row, index) {
      return [
        index + 1,
        row.handler,
        row.handlerExists ? '예' : '아니오',
        row.eventType,
        row.sourceType,
        row.sourceId,
        row.classification,
        row.legacyDescription,
        row.uniqueId
      ];
    });

    sheet.getRange(installedStartRow + 1, 1, installedValues.length, installedHeaders.length)
      .setValues(installedValues);
  } else {
    sheet.getRange(installedStartRow + 1, 1).setValue('현재 계정 소유 설치형 트리거 없음');
  }

  var lastRow = sheet.getLastRow();
  var lastColumn = Math.max(sheet.getLastColumn(), planHeaders.length);

  if (lastRow > 0 && lastColumn > 0) {
    sheet.getRange(1, 1, lastRow, lastColumn).setVerticalAlignment('middle');
  }

  sheet.setColumnWidth(1, 190);
  sheet.setColumnWidth(2, 180);
  sheet.setColumnWidth(3, 280);
  sheet.setColumnWidth(4, 100);
  sheet.setColumnWidth(5, 110);
  sheet.setColumnWidth(6, 130);
  sheet.setColumnWidth(7, 160);
  sheet.setColumnWidth(8, 300);
  sheet.setColumnWidth(9, 180);
  sheet.setColumnWidth(10, 90);
  sheet.setColumnWidth(11, 110);
  sheet.setColumnWidth(12, 80);
  sheet.setColumnWidth(13, 80);
  sheet.setColumnWidth(14, 130);
  sheet.setColumnWidth(15, 190);
  sheet.setColumnWidth(16, 360);

  return sheet;
}


/****************************************************
 * 삭제 및 안전 검사
 ****************************************************/

function TRG_deleteTriggers_(triggers) {
  var result = {
    deletedCount: 0,
    failedCount: 0,
    failures: []
  };

  triggers.forEach(function(trigger) {
    try {
      var description = TRG_describeTrigger_(trigger);
      ScriptApp.deleteTrigger(trigger);
      result.deletedCount += 1;
      Logger.log('트리거 삭제: ' + description);
    } catch (error) {
      result.failedCount += 1;
      result.failures.push({
        trigger: TRG_safeDescribeTrigger_(trigger),
        error: error && error.message ? error.message : String(error)
      });
    }
  });

  return result;
}


function TRG_assertAutomationOwner_() {
  var actualEmail = TRG_getEffectiveUserEmail_();
  var expectedEmail = String(TRG_MANAGER_CONFIG.automationOwnerEmail || '').toLowerCase();

  if (!actualEmail) {
    throw new Error(
      '현재 실행 계정 이메일을 확인할 수 없습니다. ' +
      expectedEmail + ' 계정으로 Apps Script 편집기에서 직접 실행하세요.'
    );
  }

  if (actualEmail !== expectedEmail) {
    throw new Error(
      '트리거 중앙관리는 자동화 소유 계정에서만 실행할 수 있습니다. ' +
      '허용 계정: ' + expectedEmail + ' / 현재 계정: ' + actualEmail
    );
  }
}


function TRG_getEffectiveUserEmail_() {
  var email = '';

  try {
    email = Session.getEffectiveUser().getEmail() || '';
  } catch (ignoreEffectiveUserError) {
    email = '';
  }

  return String(email).trim().toLowerCase();
}


/****************************************************
 * 보조 함수
 ****************************************************/

function TRG_getVendorSpreadsheetIds_() {
  if (typeof TARGET_FILES !== 'undefined' && TARGET_FILES) {
    return {
      KJ: String(TARGET_FILES.KJ || ''),
      '일신': String(TARGET_FILES['일신'] || '')
    };
  }

  // sb02_completedToVendor.gs의 현재 설정값과 동일한 안전 폴백.
  return {
    KJ: '1uSj0qnAiuelxd1yuDn_7BCB8cHRePaDzJGgih144Boc',
    '일신': '1F_rc7WCrjyMIeKm4N_Kgh004738ZiADTagQG13DuVFw'
  };
}


function TRG_handlerExists_(handlerName) {
  var safeName = String(handlerName || '');

  if (!/^[A-Za-z_$][0-9A-Za-z_$]*$/.test(safeName)) {
    return false;
  }

  try {
    if (typeof globalThis !== 'undefined' && typeof globalThis[safeName] === 'function') {
      return true;
    }
  } catch (ignoreGlobalThisError) {
    // 아래 eval 보조 검사로 계속 진행
  }

  try {
    return eval('typeof ' + safeName + ' === "function"');
  } catch (ignoreEvalError) {
    return false;
  }
}


function TRG_makeSignature_(handler, eventType, sourceType, sourceId) {
  return [
    String(handler || ''),
    String(eventType || '').toUpperCase(),
    String(sourceType || '').toUpperCase(),
    String(sourceId || '')
  ].join('|');
}


function TRG_enumToString_(value) {
  if (value === null || typeof value === 'undefined') {
    return '';
  }

  return String(value).replace(/^.*\./, '').toUpperCase();
}


function TRG_describeTrigger_(trigger) {
  var row = TRG_triggerToRow_(trigger);
  return row.handler + ' / ' + row.eventType + ' / ' + row.sourceType +
    (row.sourceId ? ' / ' + row.sourceId : '');
}


function TRG_safeDescribeTrigger_(trigger) {
  try {
    return TRG_describeTrigger_(trigger);
  } catch (error) {
    return '트리거 정보 조회 실패: ' + (error && error.message ? error.message : String(error));
  }
}
