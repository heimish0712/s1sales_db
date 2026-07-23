/****************************************************
 * TriggerManager.gs
 * 영업관리대장 설치형 트리거 중앙관리 - 13단계
 *
 * 운영 원칙:
 * - 설치형 트리거 소유 계정은 bang@s1samsung.com 하나로 고정
 * - 통합 onEdit/onChange와 5분 핵심 동기화 파이프라인 구현 완료
 * - KJ 1분 분류는 최근파일·매칭폴더만 처리, 6시간 점검은 전체보정 유지
 * - 핵심 자동화는 기능별 lease를 사용하고 편집 실패는 5분 재처리 큐로 이관
 * - 정식 13개 트리거 일괄 재설치 기능 활성화
 * - 구형 개별 설치·부분삭제 공개 함수는 제거하고 중앙관리 진입점만 유지
 * - 재설치는 현재 계정 소유 트리거를 전부 삭제한 뒤 정식 계획만 설치
 * - 단순 트리거(onOpen/onEdit/onSelectionChange)와 웹앱(doGet/doPost)은
 *   ScriptApp 설치형 트리거가 아니므로 이 관리 대상에서 제외
 ****************************************************/

var TRG_MANAGER_CONFIG = Object.freeze({
  automationOwnerEmail: 'bang@s1samsung.com',
  statusSheetName: '_트리거현황',
  planVersion: '2026-07-23-PHASE16',
  installEnabled: true,
  timezone: 'Asia/Seoul',
  reinstallBackupPropertyKey: 'TRG_LAST_REINSTALL_BACKUP_V1',
  repairRequestPropertyKey: 'TRG_CANONICAL_REPAIR_REQUEST_V1'
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

  var managementSs = sheet.getParent();
  managementSs.setActiveSheet(sheet);
  managementSs.toast(
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
    '중앙 복구요청: ' + snapshot.summary.repairRequestCount + '건',
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
      '필요하면 자동화 관리 메뉴의 정식 13개 전환 실행을 사용하세요.',
    ui.ButtonSet.OK
  );

  return result;
}


/**
 * 과거 '빈 상태에서 정식 설치' 공개 함수의 호환 래퍼.
 * 7단계부터는 전환 사전점검과 사후검증을 포함한 공식 전환 함수로 위임한다.
 */
function TRG_installCanonicalTriggers() {
  // 7단계부터 빈 프로젝트 설치도 동일한 사전점검·동기화·사후검증 경로를 사용한다.
  if (typeof AUTOMATION_executeCanonicalCutover === 'function') {
    return AUTOMATION_executeCanonicalCutover();
  }

  throw new Error('AutomationCutover.gs가 로드되지 않았습니다. 전체 소스를 다시 반영하세요.');
}


/**
 * 과거 전체 재설치 공개 함수의 호환 래퍼.
 * 7단계부터는 AUTOMATION_executeCanonicalCutover()가 공식 전환 경로다.
 */
function TRG_reinstallAll() {
  // 7단계부터 공식 재설치는 사전점검·핵심동기화·사후검증을 포함한 전환 함수로 위임한다.
  if (typeof AUTOMATION_executeCanonicalCutover === 'function') {
    return AUTOMATION_executeCanonicalCutover();
  }

  throw new Error('AutomationCutover.gs가 로드되지 않았습니다. 전체 소스를 다시 반영하세요.');
}


/**
 * UI 없이 현재 트리거를 삭제하고 정식 13개를 설치하는 내부 함수.
 * AUTOMATION_executeCanonicalCutover()가 전환 가드를 획득한 상태에서 호출한다.
 */
function TRG_reinstallCanonicalInternal_(options) {
  options = options || {};
  TRG_assertAutomationOwner_();
  TRG_assertCanonicalInstallEnabled_();

  var preflight = options.preflight || TRG_preflightCanonicalInstall_();
  var currentTriggers = options.currentTriggers || ScriptApp.getProjectTriggers();
  var backup = TRG_saveReinstallBackup_(currentTriggers);
  var deleteResult = TRG_deleteTriggers_(currentTriggers);
  var reinstallMeta = {
    source: String(options.source || 'TRG_reinstallCanonicalInternal_'),
    backupSavedAt: String(backup.savedAt || ''),
    previousTriggerCount: Number(backup.count || 0),
    deletedCount: Number(deleteResult.deletedCount || 0),
    deleteFailedCount: Number(deleteResult.failedCount || 0),
    createdCount: 0
  };

  if (deleteResult.failedCount > 0) {
    TRG_writeStatusSheet_(TRG_buildStatusSnapshot_());
    var deleteError = new Error(
      '기존 트리거 일부를 삭제하지 못해 재설치를 중단했습니다. 삭제 실패 ' +
      deleteResult.failedCount + '개'
    );
    deleteError.triggerReinstallResult = reinstallMeta;
    throw deleteError;
  }

  var installResult;

  try {
    installResult = TRG_createCanonicalTriggers_(preflight.plan);
    reinstallMeta.createdCount = Number(installResult.createdCount || 0);
  } catch (installError) {
    installError.triggerReinstallResult = reinstallMeta;
    TRG_writeStatusSheet_(TRG_buildStatusSnapshot_());
    throw installError;
  }

  var verificationSnapshot = TRG_buildStatusSnapshot_();

  if (!TRG_isCanonicalSnapshotHealthy_(verificationSnapshot)) {
    TRG_writeStatusSheet_(verificationSnapshot);
    var verifyError = new Error(
      '재설치는 완료됐지만 정식 계획 검증이 일치하지 않습니다. ' +
      TRG_MANAGER_CONFIG.statusSheetName + ' 시트를 확인하세요.'
    );
    verifyError.triggerReinstallResult = reinstallMeta;
    throw verifyError;
  }

  TRG_clearCanonicalRepairRequest_();
  var snapshot = TRG_buildStatusSnapshot_();
  TRG_writeStatusSheet_(snapshot);

  return {
    cancelled: false,
    source: reinstallMeta.source,
    backup: backup,
    deletedCount: reinstallMeta.deletedCount,
    createdCount: reinstallMeta.createdCount,
    summary: snapshot.summary
  };
}


/**
 * onOpen에서 호출하는 자동화 중앙관리 메뉴.
 * 메뉴는 모든 사용자에게 보일 수 있으나 트리거 변경 함수는
 * bang@s1samsung.com 소유 계정 검증을 통과해야 실행된다.
 */
function TRG_addAutomationManagementMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('자동화 관리')
    .addItem('긴급 장애 복구 미리보기', 'AUTOMATION_previewEmergencyRemediation')
    .addItem('긴급 장애 복구 실행', 'AUTOMATION_executeEmergencyRemediation')
    .addSeparator()
    .addItem('전환 사전점검', 'AUTOMATION_previewCutoverReadiness')
    .addItem('정식 13개 전환 실행', 'AUTOMATION_executeCanonicalCutover')
    .addItem('전환 사후검증', 'AUTOMATION_verifyCutoverNow')
    .addItem('전환 기록 열기', 'AUTOMATION_showCutoverLogSheet')
    .addSeparator()
    .addItem('트리거 현황 열기', 'TRG_showTriggerStatus')
    .addItem('정식 13개 구조 검증', 'TRG_verifyCanonicalTriggers')
    .addItem('백그라운드 파일 바인딩 검증', 'AUTOMATION_verifyBackgroundSpreadsheetBindings')
    .addSeparator()
    .addItem('핵심 동기화 지금 실행', 'AUTOMATION_runCoreDataSyncPipelineNow')
    .addItem('신규 유지보수 이식 미리보기', 'ITMNEW_previewMissingContracts_2026')
    .addItem('신규 유지보수 누락분 지금 이식', 'ITMNEW_syncMissingContractsNow_2026')
    .addItem('신규 유지보수 이식 로그 열기', 'ITMNEW_showTransferLogSheet_2026')
    .addItem('자동화 실행상태 열기', 'AUTOMATION_showAutomationStatusSheet')
    .addSeparator()
    .addItem('장애 상태 미리보기', 'AUTOMATION_previewHealthStatus')
    .addItem('장애 점검·알림 지금 실행', 'AUTOMATION_runHealthMonitorNow')
    .addItem('장애 상태 열기', 'AUTOMATION_showHealthStatusSheet')
    .addItem('장애 알림 로그 열기', 'AUTOMATION_showHealthAlertLogSheet')
    .addSeparator()
    .addItem('재처리 큐 열기', 'AUTOMATION_showRetryQueueSheet')
    .addItem('재처리 큐 지금 처리', 'AUTOMATION_retryQueueNow')
    .addItem('최종 실패 작업 다시 시도', 'AUTOMATION_requeueFailedRetryJobs')
    .addItem('재처리 이력 열기', 'AUTOMATION_showRetryArchiveSheet')
    .addSeparator()
    .addItem('유지관리 미리보기', 'AUTOMATION_previewMaintenance')
    .addItem('유지관리 지금 실행', 'AUTOMATION_runMaintenanceNow')
    .addItem('유지관리 상태 열기', 'AUTOMATION_showMaintenanceStatusSheet')
    .addItem('만료 다운로드 토큰만 정리', 'AUTOMATION_cleanupDownloadTokensNow')
    .addSeparator()
    .addItem('백업 보존정책 미리보기', 'AUTOMATION_previewBackupRetention')
    .addItem('백업 보존정책 지금 정리', 'AUTOMATION_runBackupRetentionNow')
    .addItem('백업 보존상태 열기', 'AUTOMATION_showBackupRetentionStatusSheet')
    .addToUi();
}



/**
 * 정식 트리거 누락을 자동으로 감지한 실행 경로가 중앙관리 복구요청을 기록한다.
 * 트리거를 직접 만들지는 않는다.
 */
function TRG_recordCanonicalRepairRequest_(planKey, sourceName, detail) {
  var props = PropertiesService.getScriptProperties();
  var propertyKey = TRG_MANAGER_CONFIG.repairRequestPropertyKey;
  var nowIso = new Date().toISOString();
  var payload = TRG_readCanonicalRepairRequest_();

  if (!payload || typeof payload !== 'object') {
    payload = {
      version: 'V1',
      firstRequestedAt: nowIso,
      lastRequestedAt: nowIso,
      requestCount: 0,
      requests: {}
    };
  }

  if (!payload.requests || typeof payload.requests !== 'object') {
    payload.requests = {};
  }

  var normalizedKey = String(planKey || 'UNKNOWN').trim() || 'UNKNOWN';
  var current = payload.requests[normalizedKey] || {
    count: 0,
    firstRequestedAt: nowIso
  };

  current.count = Number(current.count || 0) + 1;
  current.lastRequestedAt = nowIso;
  current.sourceName = String(sourceName || '');
  current.detail = String(detail || '').slice(0, 1000);
  payload.requests[normalizedKey] = current;
  payload.lastRequestedAt = nowIso;
  payload.requestCount = Number(payload.requestCount || 0) + 1;

  props.setProperty(propertyKey, JSON.stringify(payload));
  return payload;
}


function TRG_readCanonicalRepairRequest_() {
  var raw = PropertiesService.getScriptProperties().getProperty(
    TRG_MANAGER_CONFIG.repairRequestPropertyKey
  );

  if (!raw) return null;

  try {
    return JSON.parse(raw);
  } catch (error) {
    return {
      version: 'V1',
      firstRequestedAt: '',
      lastRequestedAt: '',
      requestCount: 1,
      requests: {
        CORRUPTED_REPAIR_REQUEST: {
          count: 1,
          detail: '복구요청 JSON 손상: ' + String(error && error.message ? error.message : error)
        }
      }
    };
  }
}


function TRG_clearCanonicalRepairRequest_() {
  PropertiesService.getScriptProperties().deleteProperty(
    TRG_MANAGER_CONFIG.repairRequestPropertyKey
  );
}

/****************************************************
 * 정식 13개 트리거 계획
 ****************************************************/

/**
 * 현재 중앙관리에서 설치하는 정식 계획.
 *
 * 정의 행은 12개지만 backupSalesLedger가 하루 2회이므로
 * expectedCount 합계는 13개다.
 *
 * Apps Script Trigger 객체는 시간 트리거의 실제 분/시각 설정을
 * 다시 읽어오는 API를 제공하지 않는다. 따라서 현재 검증에서는
 * 핸들러 + 이벤트 유형 + 대상 파일 + 개수까지만 자동 비교한다.
 */
function TRG_getCanonicalPlan_() {
  var ss = TRG_getManagementSpreadsheet_();
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
      implementationStatus: '3단계 구현 완료 / 정식 설치 가능',
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
      implementationStatus: '3단계 구현 완료 / 정식 설치 가능',
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
      implementationStatus: '3단계 구현 완료 / 정식 설치 가능',
      note: '1단계 실패 시 후속 중단, 2단계 실패 시 3단계 계속, 전체보정 플래그 안전 소비'
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
      implementationStatus: '경량화 완료',
      note: '최근 원본만 조회하고 매칭된 계약번호 폴더만 필요 시 준비'
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
      implementationStatus: '전체보정 유지',
      note: '고객 전체 폴더 보정 및 모든 원본 누락 안전 점검'
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
    installedOnEdit: 'sb02 공용 onEdit; 수행사 파일 2개는 정식계획, 영업관리대장용은 구형 개별 트리거',
    syncAllFromMasterTimeDriven: 'sb02 개별 5분 수행사 동기화',
    ITMAINT_onEditSync_2026: 'sb03 개별 onEdit',
    ITMAINT_onChangeSync_2026: 'sb03 개별 onChange',
    ITMAINT_timeDrivenSync_2026: 'sb03 개별 5분 전체동기화',
    customerFolderInstallableOnEdit: '고객사 폴더 개별 onEdit',
    onMailRequestEdit: '과거 authWarmup 고아 트리거. 설치 함수는 6단계에서 차단됨',
    onEditSync_정보통신유지보수: 'sb03 구버전 onEdit',
    onChangeSync_정보통신유지보수: 'sb03 구버전 onChange',
    timeDrivenSync_정보통신유지보수: 'sb03 구버전 시간 트리거',
    dummyAuthTriggerTarget_: '권한 사전인증용 임시 트리거'
  };
}



/****************************************************
 * 정식 설치 사전검사·생성
 ****************************************************/

function TRG_assertCanonicalInstallEnabled_() {
  if (!TRG_MANAGER_CONFIG.installEnabled) {
    throw new Error('정식 트리거 설치 기능이 잠겨 있습니다.');
  }
}


function TRG_preflightCanonicalInstall_() {
  var plan = TRG_getCanonicalPlan_();
  var plannedCount = plan.reduce(function(total, item) {
    return total + Number(item.expectedCount || 0);
  }, 0);

  if (plannedCount !== 13) {
    throw new Error('정식 트리거 계획 합계가 13개가 아닙니다: ' + plannedCount);
  }

  if (plannedCount > 20) {
    throw new Error('Apps Script 사용자당 스크립트 트리거 한도 20개를 초과합니다.');
  }

  var missingHandlers = plan.filter(function(item) {
    return !TRG_handlerExists_(item.handler);
  }).map(function(item) {
    return item.handler;
  });

  if (missingHandlers.length > 0) {
    throw new Error('구현되지 않은 정식 핸들러가 있습니다: ' + missingHandlers.join(', '));
  }

  var missingSourceIds = plan.filter(function(item) {
    return item.sourceType === 'SPREADSHEETS' && !String(item.sourceId || '').trim();
  });

  if (missingSourceIds.length > 0) {
    throw new Error(
      '대상 스프레드시트 ID가 비어 있습니다: ' +
      missingSourceIds.map(function(item) { return item.key; }).join(', ')
    );
  }

  // 실제 접근 권한까지 삭제 전에 확인한다.
  var checkedIds = {};
  plan.forEach(function(item) {
    if (item.sourceType !== 'SPREADSHEETS') return;

    var sourceId = String(item.sourceId || '');
    if (checkedIds[sourceId]) return;

    var spreadsheet = SpreadsheetApp.openById(sourceId);
    if (!spreadsheet) {
      throw new Error('대상 스프레드시트를 열 수 없습니다: ' + sourceId);
    }

    checkedIds[sourceId] = true;
  });

  var mainPlan = plan.filter(function(item) { return item.key === 'MAIN_EDIT_DISPATCHER'; })[0];
  var mainSpreadsheetId = String(mainPlan && mainPlan.sourceId || '');
  if (!mainSpreadsheetId) {
    throw new Error('정식 계획에서 영업관리대장 스프레드시트 ID를 확인할 수 없습니다.');
  }
  PropertiesService.getScriptProperties().setProperty(
    typeof PROP_MASTER_SPREADSHEET_ID !== 'undefined'
      ? PROP_MASTER_SPREADSHEET_ID
      : 'MASTER_SPREADSHEET_ID',
    mainSpreadsheetId
  );

  return {
    ok: true,
    plan: plan,
    plannedCount: plannedCount,
    checkedSpreadsheetCount: Object.keys(checkedIds).length
  };
}


function TRG_createCanonicalTriggers_(plan) {
  var createdTriggers = [];

  try {
    plan.forEach(function(item) {
      var createdForItem = TRG_createTriggersForPlanItem_(item);
      createdForItem.forEach(function(trigger) {
        createdTriggers.push(trigger);
      });
    });
  } catch (err) {
    // 이번 설치에서 만든 일부 트리거만 롤백한다.
    createdTriggers.forEach(function(trigger) {
      try {
        ScriptApp.deleteTrigger(trigger);
      } catch (ignoreRollbackError) {
        // 상태 시트 검증에서 잔여 트리거를 확인할 수 있음
      }
    });

    throw new Error(
      '정식 트리거 설치 중 오류가 발생해 이번에 생성한 트리거를 롤백했습니다: ' +
      (err && err.message ? err.message : String(err))
    );
  }

  return {
    createdCount: createdTriggers.length,
    plannedCount: plan.reduce(function(total, item) {
      return total + Number(item.expectedCount || 0);
    }, 0),
    handlers: createdTriggers.map(function(trigger) {
      return trigger.getHandlerFunction();
    })
  };
}


function TRG_createTriggersForPlanItem_(item) {
  var handler = item.handler;
  var sourceId = item.sourceId;
  var created = [];

  try {
    switch (item.key) {
      case 'MAIN_EDIT_DISPATCHER':
        created.push(
          ScriptApp.newTrigger(handler).forSpreadsheet(sourceId).onEdit().create()
        );
        break;

      case 'MAIN_CHANGE_DISPATCHER':
        created.push(
          ScriptApp.newTrigger(handler).forSpreadsheet(sourceId).onChange().create()
        );
        break;

      case 'VENDOR_KJ_EDIT':
      case 'VENDOR_ILSHIN_EDIT':
        created.push(
          ScriptApp.newTrigger(handler).forSpreadsheet(sourceId).onEdit().create()
        );
        break;

      case 'CORE_DATA_SYNC':
      case 'MAIL_ARCHIVE_QUEUE':
        created.push(
          ScriptApp.newTrigger(handler).timeBased().everyMinutes(5).create()
        );
        break;

      case 'DISCORD_SUPPORT_ALERT':
      case 'KJ_RECENT_CLASSIFIER':
        created.push(
          ScriptApp.newTrigger(handler).timeBased().everyMinutes(1).create()
        );
        break;

      case 'KJ_SAFETY_FULL_SCAN':
        created.push(
          ScriptApp.newTrigger(handler).timeBased().everyHours(6).create()
        );
        break;

      case 'KJ_VENDOR_UPLOAD_SYNC':
        created.push(
          ScriptApp.newTrigger(handler).timeBased().everyMinutes(30).create()
        );
        break;

      case 'SALES_LEDGER_BACKUP':
        created.push(
          ScriptApp.newTrigger(handler)
            .timeBased()
            .everyDays(1)
            .atHour(12)
            .nearMinute(30)
            .inTimezone(TRG_MANAGER_CONFIG.timezone)
            .create()
        );
        created.push(
          ScriptApp.newTrigger(handler)
            .timeBased()
            .everyDays(1)
            .atHour(18)
            .nearMinute(0)
            .inTimezone(TRG_MANAGER_CONFIG.timezone)
            .create()
        );
        break;

      case 'SENT_FILE_HISTORY':
        created.push(
          ScriptApp.newTrigger(handler)
            .timeBased()
            .everyDays(1)
            .atHour(TRG_getSentFileHistoryHour_())
            .inTimezone(TRG_MANAGER_CONFIG.timezone)
            .create()
        );
        break;

      default:
        throw new Error('설치 규칙이 정의되지 않은 정식 계획 항목: ' + item.key);
    }

    if (created.length !== Number(item.expectedCount || 0)) {
      throw new Error(
        item.key + ' 생성 개수 불일치: 예정 ' + item.expectedCount +
        ' / 실제 ' + created.length
      );
    }

    return created;
  } catch (err) {
    // 같은 계획 항목 안에서 일부만 생성된 경우도 즉시 정리한다.
    created.forEach(function(trigger) {
      try {
        ScriptApp.deleteTrigger(trigger);
      } catch (ignoreItemRollbackError) {
        // 상위 검증에서 잔여 트리거 확인
      }
    });

    throw err;
  }
}

function TRG_getSentFileHistoryHour_() {
  try {
    if (typeof getSentFileArchiveConfig_ === 'function') {
      var config = getSentFileArchiveConfig_();
      var hour = Number(config && config.DAILY_HISTORY_SYNC_HOUR);
      if (hour >= 0 && hour <= 23) return hour;
    }
  } catch (ignoreConfigError) {
    // 기본 19시 사용
  }

  return 19;
}


function TRG_isCanonicalSnapshotHealthy_(snapshot) {
  return !!(
    snapshot &&
    snapshot.summary &&
    snapshot.summary.installedTriggerCount === snapshot.summary.canonicalPlannedTriggerCount &&
    snapshot.summary.canonicalMissingTriggerCount === 0 &&
    snapshot.summary.canonicalExcessTriggerCount === 0 &&
    snapshot.summary.orphanTriggerCount === 0 &&
    snapshot.summary.unknownTriggerCount === 0 &&
    snapshot.summary.legacyTriggerCount === 0
  );
}


function TRG_saveReinstallBackup_(triggers) {
  var rows = triggers.map(function(trigger) {
    return TRG_triggerToRow_(trigger);
  });

  var backup = {
    savedAt: new Date().toISOString(),
    ownerEmail: TRG_getEffectiveUserEmail_(),
    count: rows.length,
    triggers: rows
  };

  PropertiesService.getScriptProperties().setProperty(
    TRG_MANAGER_CONFIG.reinstallBackupPropertyKey,
    JSON.stringify(backup)
  );

  return backup;
}


/****************************************************
 * 현황 수집 및 비교
 ****************************************************/

function TRG_buildStatusSnapshot_() {
  var plan = TRG_getCanonicalPlan_();
  var installed = ScriptApp.getProjectTriggers();
  var installedRows = installed.map(TRG_triggerToRow_);
  var legacyHandlers = TRG_getKnownLegacyHandlers_();
  var repairRequest = TRG_readCanonicalRepairRequest_();
  var repairRequestKeys = repairRequest && repairRequest.requests
    ? Object.keys(repairRequest.requests)
    : [];
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
      repairRequestCount: repairRequestKeys.length,
      repairRequestLastAt: repairRequest ? String(repairRequest.lastRequestedAt || '') : '',
      installEnabled: TRG_MANAGER_CONFIG.installEnabled,
      planVersion: TRG_MANAGER_CONFIG.planVersion
    },
    repairRequest: repairRequest,
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
  var ss = TRG_getManagementSpreadsheet_();
  var sheet = ss.getSheetByName(TRG_MANAGER_CONFIG.statusSheetName);

  if (!sheet) {
    sheet = ss.insertSheet(TRG_MANAGER_CONFIG.statusSheetName);
  }

  sheet.clear();
  sheet.setFrozenRows(1);

  var tz = Session.getScriptTimeZone() || TRG_MANAGER_CONFIG.timezone;
  var generatedAtText = Utilities.formatDate(snapshot.generatedAt, tz, 'yyyy-MM-dd HH:mm:ss');

  var summaryValues = [
    ['트리거 중앙관리 6단계', '값'],
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
    ['중앙 복구요청', snapshot.summary.repairRequestCount + '건'],
    ['최근 복구요청', snapshot.summary.repairRequestLastAt || '없음'],
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


function TRG_getManagementSpreadsheet_() {
  if (typeof AUTOMATION_getRuntimeMasterSpreadsheet_ === 'function') {
    return AUTOMATION_getRuntimeMasterSpreadsheet_();
  }

  if (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.MASTER_SPREADSHEET_ID) {
    return SpreadsheetApp.openById(String(CONFIG.MASTER_SPREADSHEET_ID));
  }

  throw new Error(
    '영업관리대장 스프레드시트 ID를 확인할 수 없습니다. ' +
    'AutomationRuntime.gs와 CONFIG.MASTER_SPREADSHEET_ID를 확인하세요.'
  );
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
