function onOpen() {
  const ui = SpreadsheetApp.getUi();
  addOrderMailManualMenu_();
  addKjShareCopyMenu_();
  TRG_addAutomationManagementMenu_();
  ui.createMenu('자동 입력')
    .addItem('자동입력 누락 검산(현재 시트)', 'AUTOEDIT_auditActiveSheet')
    .addItem('자동입력 누락 검산·보정(현재 시트)', 'AUTOEDIT_repairActiveSheet')
    .addItem('핵심시트 제한보정 지금 실행', 'AUTOEDIT_runScheduledRepairNow')
    .addSeparator()
    .addItem('연면적 기준 관리등급 일괄 반영', 'fillManagementGradeByAreaOnActiveSheetOnce')
    .addItem('계약단위 기준 기본조건 일괄 반영', 'fillContractDefaultsByUnitOnActiveSheetOnce')
    .addItem('계약조건 기준 최종 견적가 일괄 계산', 'fillFinalQuotePriceByContractConditionsOnActiveSheetOnce')
    .addToUi();

  ui.createMenu('메일자동화')
    .addItem('메일자동발송', 'mailAutoSend')
    .addItem('현재 선택행 저장', 'rememberCurrentSelectedMailRow')
    .addItem('내 선택행 기억 지우기', 'clearMyMailRowHighlight')
    .addItem('진행상태 초기화', 'clearMyMailRunProgress')
    .addSeparator()
    .addItem('작업공간 공유드라이브 폴더ID 저장', 'setMailAutoWorkspaceFolderId')
    .addItem('작업공간 저장 위치 확인', 'checkMailAutoWorkspaceFolder')
    .addSeparator()
    .addItem('발송파일 저장 설정 확인', 'checkSentFileArchiveConfig')
    .addItem('메일발송실패큐_DB 열기', 'openMailSendFailureQueueP523')
    .addItem('발송파일저장큐_DB 열기', 'openSentFileArchiveQueueP523')
    .addItem('메일/발송파일 큐 요약 보기', 'showMailQueueSummaryP524')
    .addItem('발송파일저장큐 즉시 재처리', 'processSentFileArchiveQueueNowP523')
    .addItem('발송이력 일일반영 수동실행', 'syncSentFileFolderHistoryDaily')
    .addSeparator()
    .addItem('도장/로고 캐시 예열', 'warmUpMailAutoPrestampedTemplateCache')
    .addItem('도장/로고 캐시 초기화', 'clearMailAutoPrestampedTemplateCache')
    .addItem('비편집 shortcut 캐시 예열', 'warmUpMailAutoReviewShortcutCache')
    .addItem('비편집 shortcut 캐시 초기화', 'clearMailAutoReviewShortcutCache')
    .addSeparator()
    .addItem('하이웍스 API키 저장', 'setHiworksApiKey')
    .addItem('하이웍스 API키 확인', 'checkHiworksApiKey')
    .addItem('하이웍스 토큰 저장 안내', 'showHiworksTokenGuide')
    .addToUi();

  ui.createMenu('메모 정리')
    .addItem('마스터시트 메모 최종 업데이트본 생성', 'buildFinalMasterMemoUpdateColumn')
    .addToUi();
  KJUS_addVendorUploadSyncMenu_();
}