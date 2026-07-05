/*******************************************************
 * S1 Sales Portal → 기존 자동메일 Worker 브릿지 v76
 * - v44 PROGRESS RETRY SAVEPOINT 기반
 * - v76: 포털 [파일 확인/수정] action 추가
 * - v76: reviewSessionId를 sendPortalMail payload에 그대로 전달
 * - v76: selectedKeys는 체크박스 임시 변경 없이 실행 인스턴스에만 주입
 * - v99: sendPortalMail 단계에서만 [전체문서] 누락 key(수행사정보/샘플보고서)를 복구
 *
 * 붙여넣을 위치:
 * - 시트에서 실제로 정상 발송되는 기존 자동메일 Apps Script 프로젝트
 * - 기존 자동메일 Code.gs 맨 아래의 Worker 브릿지 블록 교체
 *
 * 운영 원칙:
 * - 하이웍스 발송부/첨부 multipart/파일명 로직은 절대 수정하지 않음
 * - 기존 MailAutomationService.sendFromDialog(payload)를 그대로 실행
 * - 기존 MailAutomationService.prepareFilesForReview(payload)를 그대로 실행
 * - 포털에서 넘어온 selectedKeys만 해당 실행 인스턴스에 주입
 * - 마스터 체크박스 임시 변경 없음
 * - getProgress / cancelRun을 Worker에서 처리하여 포털 진행률과 연결
 *******************************************************/

function doPost(e) {
  return mailWorkerDoPostV76_(e);
}

function mailWorkerDoPostV76_(e) {
  try {
    const body = mailWorkerParseJsonBodyV76_(e);
    const expectedSecret = String(PropertiesService.getScriptProperties().getProperty('MAIL_WORKER_SHARED_SECRET') || '').trim();
    const gotSecret = String(body.secret || '').trim();

    if (!expectedSecret) {
      throw new Error('메일 Worker 프로젝트의 Script Properties에 MAIL_WORKER_SHARED_SECRET 값이 없습니다.');
    }
    if (!gotSecret || gotSecret !== expectedSecret) {
      throw new Error('메일 Worker 인증 실패: shared secret이 일치하지 않습니다.');
    }

    const action = String(body.action || '').trim();
    const payload = body.payload || {};

    if (action === 'health') {
      return mailWorkerJsonV76_({
        ok: true,
        message: 'MAIL_WORKER_OK',
        workerScriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })(),
        activeUser: (() => { try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; } })(),
        effectiveUser: (() => { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } })(),
        hasMailAutomationService: typeof MailAutomationService !== 'undefined',
        hasSendMailFromDialog: typeof sendMailFromDialog === 'function',
        hasPrepareFilesForReview: typeof MailAutomationService !== 'undefined' && typeof MailAutomationService.prototype.prepareFilesForReview === 'function',
        hasPreviewPortalMailContent: true,
        hasProgressTracker: typeof ProgressTracker !== 'undefined',
        hasHiworksToken: !!String(PropertiesService.getScriptProperties().getProperty('HIWORKS_API_KEY') || '').trim(),
        runtime: mailWorkerRuntimeInfoV76_()
      });
    }

    if (action === 'sendPortalMail') {
      const result = mailWorkerSendPortalMailV76_(payload);
      return mailWorkerJsonV76_({
        ok: true,
        message: '메일 Worker 발송 완료',
        result: result,
        runtime: mailWorkerRuntimeInfoV76_()
      });
    }

    if (
      action === 'preparePortalMailFilesForReview' ||
      action === 'preparePortalReviewFiles' ||
      action === 'prepareMailFilesForReview'
    ) {
      const result = mailWorkerPreparePortalMailFilesForReviewV76_(payload);
      return mailWorkerJsonV76_({
        ok: true,
        message: '메일 Worker 파일 확인/수정 폴더 생성 완료',
        result: result,
        runtime: mailWorkerRuntimeInfoV76_()
      });
    }

    if (action === 'previewPortalMailContent' || action === 'getPortalMailContentPreview' || action === 'previewMailContent') {
      const result = mailWorkerPreviewPortalMailContentP501_(payload);
      return mailWorkerJsonV76_({
        ok: true,
        message: '메일 Worker 기본 제목/본문 조회 완료',
        result: result,
        runtime: mailWorkerRuntimeInfoV76_()
      });
    }

    if (action === 'getProgress') {
      const runId = String(payload.runId || '').trim();
      const result = new ProgressTracker(runId).get();
      return mailWorkerJsonV76_({
        ok: true,
        message: '진행률 조회 완료',
        result: result
      });
    }

    if (action === 'cancelRun') {
      const runId = String(payload.runId || '').trim();
      const result = new ProgressTracker(runId).requestCancel();
      return mailWorkerJsonV76_({
        ok: true,
        message: '취소 요청 완료',
        result: result
      });
    }

    if (action === 'sendOrderNotificationMail') {
      const result = mailWorkerSendOrderNotificationMailV447_(payload);
      return mailWorkerJsonV76_({
        ok: true,
        message: '발주번호 생성 알림 메일 발송 완료',
        result: result,
        runtime: mailWorkerRuntimeInfoV76_()
      });
    }

    throw new Error('지원하지 않는 Worker action입니다: ' + action);
  } catch (err) {
    return mailWorkerJsonV76_({
      ok: false,
      message: String(err && err.message || err),
      detail: String(err && err.stack || err),
      runtime: mailWorkerRuntimeInfoV76_()
    });
  }
}

function mailWorkerSendPortalMailV76_(payload) {
  payload = payload || {};
  const rowNo = Number(payload.rowNo);
  const mode = String(payload.mode || '').toUpperCase();
  const selectedKeys = mailWorkerResolvePortalSelectedKeysForSendP99_(payload);

  mailWorkerAssertBaseRuntimeV76_();

  if (!rowNo || rowNo < (CONFIG.ROWS && CONFIG.ROWS.MASTER_DATA_START || 3)) {
    throw new Error('발송 행 정보가 올바르지 않습니다: ' + rowNo);
  }
  if (mode !== 'CUSTOMER' && mode !== 'TEST') {
    throw new Error('발송 모드가 올바르지 않습니다: ' + mode);
  }
  if (!selectedKeys.length) {
    throw new Error('selectedKeys가 비어 있습니다. 포털에서 발송자료를 하나 이상 선택해야 합니다.');
  }

  const selectedDefs = mailWorkerResolveFileDefsByKeysV76_(selectedKeys);

  const sendPayload = mailWorkerBuildPortalSendPayloadV76_(payload, {
    rowNo: rowNo,
    mode: mode,
    selectedKeys: selectedKeys
  });

  const result = mailWorkerRunWithInjectedSelectedDefsV76_(selectedDefs, function(service) {
    return service.sendFromDialog(sendPayload);
  });

  return {
    ok: true,
    worker: 'mailWorkerV76',
    action: 'sendPortalMail',
    rowNo: rowNo,
    mode: mode,
    selectedKeys: selectedKeys,
    selectedLabels: selectedDefs.map(function(d) { return d.label || d.key; }),
    reviewSessionId: String(payload.reviewSessionId || '').trim(),
    sendResult: result,

    // 포털 구버전 호환용: sendFromDialog 결과를 자주 쓰는 필드는 상위에도 복사합니다.
    message: result && result.message || '메일 발송 완료',
    requestNo: result && result.requestNo || '',
    to: result && result.to || '',
    attachments: result && result.attachments || []
  };
}

function mailWorkerPreparePortalMailFilesForReviewV76_(payload) {
  payload = payload || {};
  const rowNo = Number(payload.rowNo);
  const selectedKeys = mailWorkerNormalizeSelectedKeysV76_(payload.selectedKeys);

  mailWorkerAssertBaseRuntimeV76_();

  if (typeof MailAutomationService.prototype.prepareFilesForReview !== 'function') {
    throw new Error('MailAutomationService.prepareFilesForReview(payload)를 찾지 못했습니다. 파일 확인/수정 기능이 포함된 최신 자동메일 코드가 필요합니다.');
  }
  if (!rowNo || rowNo < (CONFIG.ROWS && CONFIG.ROWS.MASTER_DATA_START || 3)) {
    throw new Error('파일 확인/수정 대상 행 정보가 올바르지 않습니다: ' + rowNo);
  }
  if (!selectedKeys.length) {
    throw new Error('selectedKeys가 비어 있습니다. 포털에서 확인/수정할 자료를 하나 이상 선택해야 합니다.');
  }

  const selectedDefs = mailWorkerResolveFileDefsByKeysV76_(selectedKeys);

  const reviewPayload = mailWorkerBuildPortalReviewPayloadV76_(payload, {
    rowNo: rowNo,
    selectedKeys: selectedKeys
  });

  const result = mailWorkerRunWithInjectedSelectedDefsV76_(selectedDefs, function(service) {
    return service.prepareFilesForReview(reviewPayload);
  });

  return {
    ok: true,
    worker: 'mailWorkerV76',
    action: 'preparePortalMailFilesForReview',
    rowNo: rowNo,
    selectedKeys: selectedKeys,
    selectedLabels: selectedDefs.map(function(d) { return d.label || d.key; }),
    reviewResult: result,

    // 중요: 포털 11_MailBridgeService.gs가 바로 읽는 필드는 상위에 그대로 둡니다.
    message: result && result.message || '파일 확인/수정용 Drive 폴더 생성 완료',
    reviewSessionId: result && result.reviewSessionId || '',
    requestNo: result && result.requestNo || '',
    folderUrl: result && result.folderUrl || '',
    fileCount: result && result.fileCount || 0,
    files: result && result.files || []
  };
}

function mailWorkerResolvePortalSelectedKeysForPreviewP501_(payload) {
  payload = payload || {};
  let selectedKeys = mailWorkerNormalizeSelectedKeysV76_(payload.selectedKeys);
  if (mailWorkerPayloadLooksLikeAllDocumentsP99_(payload)) {
    selectedKeys = mailWorkerAllFileDefinitionKeysP99_();
  }
  return mailWorkerDedupeKeysP99_(selectedKeys);
}

function mailWorkerPreviewPortalMailContentP501_(payload) {
  payload = payload || {};
  const rowNo = Number(payload.rowNo);
  const selectedKeys = mailWorkerResolvePortalSelectedKeysForPreviewP501_(payload);

  mailWorkerAssertBaseRuntimeV76_();

  if (!rowNo || rowNo < (CONFIG.ROWS && CONFIG.ROWS.MASTER_DATA_START || 3)) {
    throw new Error('메일 본문 기본값 조회 행 정보가 올바르지 않습니다: ' + rowNo);
  }
  if (!selectedKeys.length) {
    throw new Error('selectedKeys가 비어 있습니다. 포털에서 발송자료를 하나 이상 선택해야 합니다.');
  }

  const selectedDefs = mailWorkerResolveFileDefsByKeysV76_(selectedKeys);

  return mailWorkerRunWithInjectedSelectedDefsV76_(selectedDefs, function(service) {
    const master = service.getMasterContext_();
    const rowObj = service.readMasterRow_(master.sheet, master.headerMap, rowNo);
    if (typeof assertPortalPayloadMatchesCurrentMasterRowV88_ === 'function') {
      assertPortalPayloadMatchesCurrentMasterRowV88_(payload, rowObj, '메일 본문 기본값 조회');
    }

    const targetData = rowObj.toPlainObject();
    const finalDefs = typeof service.applyDialogFileSelectionOverrides_ === 'function'
      ? service.applyDialogFileSelectionOverrides_(selectedDefs, payload)
      : selectedDefs.slice();
    if (!finalDefs.length) throw new Error('메일 본문 기본값을 만들 선택 자료가 없습니다.');

    let sender = null;
    try {
      const generatorSs = SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID);
      sender = new SalesRepResolver(generatorSs).resolve(targetData['영업담당자']);
    } catch (err) {
      if (typeof service.buildFallbackSenderForMailPreview_ === 'function') {
        sender = service.buildFallbackSenderForMailPreview_(targetData['영업담당자']);
      } else {
        sender = {
          name: String(targetData['영업담당자'] || '').trim(),
          title: '',
          phone: '',
          email: '',
          sharedEmail: '',
          mainPhone: '',
          division: '',
          region: '',
          displayName: String(targetData['영업담당자'] || '').trim()
        };
      }
    }

    let editor = null;
    if (typeof service.buildMailEditorPreview_ === 'function') {
      editor = service.buildMailEditorPreview_(targetData, sender, finalDefs);
    } else {
      editor = {
        subject: service.buildMailSubject_(targetData, sender, finalDefs),
        bodyHtml: service.buildMailBodyHtml_(targetData, sender, finalDefs, {
          includeInlineImages: false,
          includeLargeAttachmentLinksMarker: true
        })
      };
    }

    const bodyHtml = String(editor && editor.bodyHtml || '');
    return {
      ok: true,
      worker: 'mailWorkerP501',
      action: 'previewPortalMailContent',
      rowNo: rowNo,
      customerNo: String(payload.customerNo || '').trim(),
      selectedKeys: finalDefs.map(function(def) { return def && def.key || ''; }).filter(Boolean),
      selectedLabels: finalDefs.map(function(def) { return def && (def.label || def.key) || ''; }).filter(Boolean),
      subject: String(editor && editor.subject || ''),
      bodyHtml: bodyHtml,
      bodyText: mailWorkerEditableTextFromHtmlP501_(bodyHtml)
    };
  });
}

function mailWorkerEditableTextFromHtmlP501_(html) {
  return String(html || '')
    .replace(/<style[\s\S]*?<\/style>/gi, '')
    .replace(/<script[\s\S]*?<\/script>/gi, '')
    .replace(/<br\s*\/?>/gi, '\n')
    .replace(/<\/p>/gi, '\n\n')
    .replace(/<\/div>/gi, '\n')
    .replace(/<[^>]+>/g, '')
    .replace(/&nbsp;/gi, ' ')
    .replace(/&amp;/gi, '&')
    .replace(/&lt;/gi, '<')
    .replace(/&gt;/gi, '>')
    .replace(/&quot;/gi, '"')
    .replace(/&#39;/gi, "'")
    .replace(/\u00a0/g, ' ')
    .replace(/[ \t]+\n/g, '\n')
    .replace(/\n[ \t]+/g, '\n')
    .replace(/\n{3,}/g, '\n\n')
    .trim();
}


function mailWorkerRunWithInjectedSelectedDefsV76_(selectedDefs, runner) {
  const service = new MailAutomationService();
  const originalGetSelectedFileDefs = service.getSelectedFileDefs_;

  // 중요:
  // prototype을 덮어쓰지 않고 이 service 인스턴스에만 주입합니다.
  // 동시 실행 5~10건에서도 각 실행의 selectedKeys가 서로 섞이지 않습니다.
  service.getSelectedFileDefs_ = function(rowObj) {
    return selectedDefs.slice();
  };

  try {
    return runner(service);
  } finally {
    if (originalGetSelectedFileDefs) service.getSelectedFileDefs_ = originalGetSelectedFileDefs;
  }
}

function mailWorkerBuildPortalSendPayloadV76_(payload, base) {
  const out = {
    rowNo: base.rowNo,
    mode: base.mode,
    selectedKeys: base.selectedKeys.slice(),
    testInput: String(payload.testInput || '').trim(),
    manualTo: payload.manualTo || null,
    manualCc: payload.manualCc || null,
    removedCc: payload.removedCc || [],
    reviewSessionId: String(payload.reviewSessionId || '').trim(),
    runId: String(payload.runId || Utilities.getUuid()),
    customerNo: String(payload.customerNo || '').trim(),

    // 비교견적서 선택/제외 옵션은 기존 MailAutomationService.applyDialogFileSelectionOverrides_가 처리합니다.
    compareQuoteSheets: payload.compareQuoteSheets || payload.selectedCompareQuoteSheets || null,
    selectedCompareQuoteSheets: payload.selectedCompareQuoteSheets || payload.compareQuoteSheets || null,
    excludedCompareQuoteSheets: payload.excludedCompareQuoteSheets || payload.excludedCompareSheets || payload.removedCompareQuoteSheets || null,
    excludedCompareSheets: payload.excludedCompareSheets || payload.excludedCompareQuoteSheets || null,
    removedCompareQuoteSheets: payload.removedCompareQuoteSheets || null
  };

  // 포털에서 제목/본문 수정 기능을 붙일 때 그대로 통과시키는 필드입니다.
  // 비어 있는 값은 전달하지 않아 기존 기본 템플릿 발송 흐름을 보존합니다.
  mailWorkerCopyOptionalStringFieldV100_(payload, out, 'mailSubjectOverride', [
    'mailSubjectOverride',
    'subjectOverride',
    'mailSubject',
    'subject'
  ]);
  mailWorkerCopyOptionalStringFieldV100_(payload, out, 'mailBodyHtmlOverride', [
    'mailBodyHtmlOverride',
    'bodyHtmlOverride',
    'mailBodyHtml',
    'bodyHtml'
  ]);

  return out;
}

function mailWorkerCopyOptionalStringFieldV100_(source, target, targetKey, sourceKeys) {
  source = source || {};
  for (let i = 0; i < sourceKeys.length; i++) {
    const sourceKey = sourceKeys[i];
    if (!Object.prototype.hasOwnProperty.call(source, sourceKey)) continue;

    const value = source[sourceKey];
    if (value === undefined || value === null) continue;

    const text = String(value).trim();
    if (!text) continue;

    target[targetKey] = text;
    return;
  }
}

function mailWorkerBuildPortalReviewPayloadV76_(payload, base) {
  return {
    rowNo: base.rowNo,
    selectedKeys: base.selectedKeys.slice(),
    runId: String(payload.runId || Utilities.getUuid()),
    customerNo: String(payload.customerNo || '').trim(),

    // 비교견적(2) 제외 같은 UI 옵션이 생긴 경우까지 그대로 전달합니다.
    compareQuoteSheets: payload.compareQuoteSheets || payload.selectedCompareQuoteSheets || null,
    selectedCompareQuoteSheets: payload.selectedCompareQuoteSheets || payload.compareQuoteSheets || null,
    excludedCompareQuoteSheets: payload.excludedCompareQuoteSheets || payload.excludedCompareSheets || payload.removedCompareQuoteSheets || null,
    excludedCompareSheets: payload.excludedCompareSheets || payload.excludedCompareQuoteSheets || null,
    removedCompareQuoteSheets: payload.removedCompareQuoteSheets || null
  };
}

function mailWorkerAssertBaseRuntimeV76_() {
  if (typeof MailAutomationService === 'undefined') {
    throw new Error('MailAutomationService를 찾지 못했습니다. 이 파일은 기존 자동메일 코드가 있는 프로젝트에 붙여넣어야 합니다.');
  }
  if (typeof CONFIG === 'undefined' || !CONFIG || !Array.isArray(CONFIG.FILE_DEFINITIONS)) {
    throw new Error('CONFIG.FILE_DEFINITIONS를 찾지 못했습니다. 기존 자동메일 CONFIG가 필요합니다.');
  }
}

function mailWorkerNormalizeSelectedKeysV76_(selectedKeys) {
  const aliases = {
    vendorContract: 'serviceStandardContract',
    standardContract: 'serviceStandardContract',
    serviceContract: 'serviceStandardContract',
    contractorInfoZip: 'contractorInfo',
    terms: 'termsGuide'
  };

  const arr = Array.isArray(selectedKeys) ? selectedKeys : String(selectedKeys || '').split(/[;,\s]+/);
  const out = [];
  const seen = {};

  arr.forEach(function(key) {
    key = String(key || '').trim();
    if (!key) return;
    key = aliases[key] || key;
    if (seen[key]) return;
    seen[key] = true;
    out.push(key);
  });

  return out;
}

function mailWorkerResolvePortalSelectedKeysForSendP99_(payload) {
  payload = payload || {};

  // 기본 호환성은 기존 v76 정규화 함수를 그대로 사용합니다.
  let selectedKeys = mailWorkerNormalizeSelectedKeysV76_(payload.selectedKeys);

  // 포털에서 [전체문서] 같은 프리셋/버튼값만 넘기는 경우에는
  // CONFIG.FILE_DEFINITIONS 기준 전체 key로 복구합니다.
  if (mailWorkerPayloadLooksLikeAllDocumentsP99_(payload)) {
    selectedKeys = mailWorkerAllFileDefinitionKeysP99_();
  }

  // 포털 전체문서 흐름에서 파일 확인/수정 가능한 문서 key만 넘어오면
  // 수행사정보/샘플보고서가 누락되어 본체의 Drive 링크 전환 로직이 실행되지 않습니다.
  // sendPortalMail 최종 발송 단계에서만 두 key를 복구합니다.
  selectedKeys = mailWorkerRestoreSampleAndContractorForFullSendP99_(selectedKeys, payload);
  return mailWorkerDedupeKeysP99_(selectedKeys);
}

function mailWorkerPayloadLooksLikeAllDocumentsP99_(payload) {
  payload = payload || {};

  const boolFields = [
    'allDocuments',
    'allFiles',
    'allSelected',
    'selectAll',
    'sendAll',
    'sendAllDocuments',
    'sendAllFiles',
    'isAllDocuments',
    'isFullPackage',
    'fullPackage'
  ];

  for (let i = 0; i < boolFields.length; i++) {
    if (payload[boolFields[i]] === true) return true;
  }

  const textFields = [
    'preset',
    'sendPreset',
    'filePreset',
    'selectionPreset',
    'selectionMode',
    'fileSelectionMode',
    'documentMode',
    'documentScope',
    'packageType',
    'sendType',
    'mailType',
    'buttonType',
    'buttonLabel',
    'actionLabel'
  ];

  for (let j = 0; j < textFields.length; j++) {
    if (mailWorkerIsAllDocumentTokenP99_(payload[textFields[j]])) return true;
  }

  const rawSelectedKeys = Array.isArray(payload.selectedKeys)
    ? payload.selectedKeys
    : String(payload.selectedKeys || '').split(/[;,\s]+/);

  for (let k = 0; k < rawSelectedKeys.length; k++) {
    if (mailWorkerIsAllDocumentTokenP99_(rawSelectedKeys[k])) return true;
  }

  return false;
}

function mailWorkerIsAllDocumentTokenP99_(value) {
  const compact = String(value || '').trim().replace(/[\s_ ·ㆍ/\-,，]+/g, '').toLowerCase();
  if (!compact) return false;

  const tokens = {
    all: true,
    allfiles: true,
    alldocuments: true,
    allitems: true,
    fullpackage: true,
    fullsend: true,
    entirepackage: true,
    전체: true,
    전체문서: true,
    전체자료: true,
    전체파일: true,
    전체문서발송: true,
    전체자료발송: true,
    전체파일발송: true
  };

  return tokens[compact] === true;
}

function mailWorkerAllFileDefinitionKeysP99_() {
  return (CONFIG.FILE_DEFINITIONS || [])
    .map(function(def) { return def && def.key ? String(def.key).trim() : ''; })
    .filter(Boolean);
}

function mailWorkerRestoreSampleAndContractorForFullSendP99_(selectedKeys, payload) {
  selectedKeys = mailWorkerDedupeKeysP99_(selectedKeys || []);

  // P501: key 모양만 보고 수행사정보/샘플보고서를 자동 추가하지 않습니다.
  // 명시적인 전체문서 프리셋/토큰이 있을 때만 구버전 payload 보정용으로 동작합니다.
  if (!mailWorkerPayloadLooksLikeAllDocumentsP99_(payload || {})) return selectedKeys;

  const set = {};
  selectedKeys.forEach(function(key) {
    key = String(key || '').trim();
    if (key) set[key] = true;
  });

  const looksLikePortalFullDocuments =
    set.quote &&
    set.serviceApplication &&
    set.appointmentDoc &&
    set.termsGuide &&
    set.compareQuote;

  if (!looksLikePortalFullDocuments) return selectedKeys;

  if (!set.contractorInfo) {
    selectedKeys.push('contractorInfo');
    set.contractorInfo = true;
  }
  if (!set.sampleReport) {
    selectedKeys.push('sampleReport');
    set.sampleReport = true;
  }

  return selectedKeys;
}

function mailWorkerDedupeKeysP99_(keys) {
  const out = [];
  const seen = {};
  (keys || []).forEach(function(key) {
    key = String(key || '').trim();
    if (!key || seen[key]) return;
    seen[key] = true;
    out.push(key);
  });
  return out;
}

function mailWorkerResolveFileDefsByKeysV76_(selectedKeys) {
  const defs = CONFIG.FILE_DEFINITIONS || [];
  const byKey = {};
  defs.forEach(function(def) {
    if (def && def.key) byKey[String(def.key)] = def;
  });

  const selected = [];
  const missing = [];

  selectedKeys.forEach(function(key) {
    if (byKey[key]) selected.push(byKey[key]);
    else missing.push(key);
  });

  if (missing.length) {
    throw new Error(
      '자동메일 CONFIG.FILE_DEFINITIONS에 없는 자료 key입니다: ' + missing.join(', ') + '\n' +
      '현재 지원 key: ' + Object.keys(byKey).join(', ')
    );
  }

  return selected;
}

function mailWorkerParseJsonBodyV76_(e) {
  const raw = e && e.postData && e.postData.contents ? String(e.postData.contents) : '';
  if (!raw) throw new Error('POST body가 비어 있습니다.');
  try {
    return JSON.parse(raw);
  } catch (err) {
    throw new Error('POST JSON 파싱 실패: ' + String(err && err.message || err) + '\nbody=' + raw.slice(0, 500));
  }
}

function mailWorkerJsonV76_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj || {}, null, 2))
    .setMimeType(ContentService.MimeType.JSON);
}

function mailWorkerRuntimeInfoV76_() {
  const tokenInfo = (typeof findHiworksTokenInfo_ === 'function')
    ? findHiworksTokenInfo_()
    : { token: '', key: '', scope: '' };
  const token = String(tokenInfo.token || '');
  const secret = String(PropertiesService.getScriptProperties().getProperty('MAIL_WORKER_SHARED_SECRET') || '');

  return {
    worker: 'mailWorkerV76',
    scriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })(),
    effectiveUser: (() => { try { return Session.getEffectiveUser().getEmail(); } catch (e) { return ''; } })(),
    activeUser: (() => { try { return Session.getActiveUser().getEmail(); } catch (e) { return ''; } })(),
    masterSpreadsheetId: typeof CONFIG !== 'undefined' && CONFIG ? CONFIG.MASTER_SPREADSHEET_ID : '',
    generatorSpreadsheetId: typeof CONFIG !== 'undefined' && CONFIG ? CONFIG.GENERATOR_SPREADSHEET_ID : '',
    hasMailAutomationService: typeof MailAutomationService !== 'undefined',
    hasPrepareFilesForReview: typeof MailAutomationService !== 'undefined' && typeof MailAutomationService.prototype.prepareFilesForReview === 'function',
    hasHiworksToken: !!token,
    hiworksTokenScope: tokenInfo.scope || '',
    hiworksTokenKey: tokenInfo.key || '',
    hiworksTokenLength: token.length,
    hiworksTokenHash16: token ? Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, token)).replace(/=+$/g, '').slice(0, 16) : '',
    hasWorkerSecret: !!secret,
    workerSecretLength: secret.length,
    workerSecretHash16: secret ? Utilities.base64EncodeWebSafe(Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, secret)).replace(/=+$/g, '').slice(0, 16) : '',
    fileKeys: (typeof CONFIG !== 'undefined' && CONFIG && Array.isArray(CONFIG.FILE_DEFINITIONS))
      ? CONFIG.FILE_DEFINITIONS.map(function(d) { return d.key; })
      : []
  };
}

function debugMailWorkerRuntimeV76_() {
  const info = mailWorkerRuntimeInfoV76_();
  Logger.log(JSON.stringify(info, null, 2));
  return info;
}


/***************************************
 * P447 포탈 발주번호 생성 알림 메일
 * - 포탈에서 발주번호가 새로 생성되면 background queue가 이 action을 호출합니다.
 * - 발신자: 고객별 영업담당자 이메일(영업담당자 정보 시트에서 이름으로 resolve)
 * - 수신자: master@s1samsung.com
 * - 참조: 없음
 ***************************************/
function mailWorkerSendOrderNotificationMailV447_(payload) {
  payload = payload || {};
  mailWorkerAssertBaseRuntimeV76_();

  const contractNo = String(payload.contractNo || '').trim();
  const company = String(payload.company || '').trim();
  const salesRepName = String(payload.salesRep || payload.salesRepName || payload.contractRep || '').trim();

  if (!contractNo) throw new Error('발주메일 발송 실패: contractNo가 비어 있습니다.');
  if (!company) throw new Error('발주메일 발송 실패: company가 비어 있습니다.');
  if (!salesRepName) throw new Error('발주메일 발송 실패: 영업담당자 값이 비어 있습니다.');

  const generatorSs = SpreadsheetApp.getActiveSpreadsheet();
  const sender = new SalesRepResolver(generatorSs).resolve(salesRepName);
  const subject = contractNo + '. ' + company;
  const bodyText = subject + '\n발주번호 생성 알림';
  const bodyHtml = bodyText
    .split('\n')
    .map(function(line) { return escapeHtmlForOrderNotificationMailV447_(line); })
    .join('<br>');

  const mail = new MailMessage({
    from: sender.email,
    to: ['master@s1samsung.com'],
    cc: [],
    subject: subject,
    bodyHtml: bodyHtml,
    attachments: []
  });

  const result = new HiworksMailer(null).send(mail);

  appendOrderNotificationMailLogV447_({
    requestId: String(payload.requestId || '').trim(),
    status: '성공',
    customerNo: String(payload.customerNo || '').trim(),
    rowNo: Number(payload.rowNo) || '',
    contractNo: contractNo,
    contractRowNo: Number(payload.contractRowNo) || '',
    company: company,
    salesRep: salesRepName,
    from: sender.email,
    to: 'master@s1samsung.com',
    subject: subject,
    hiworksResult: JSON.stringify(result || {}).slice(0, 1000),
    error: ''
  });

  return {
    ok: true,
    requestId: String(payload.requestId || '').trim(),
    customerNo: String(payload.customerNo || '').trim(),
    contractNo: contractNo,
    company: company,
    salesRep: salesRepName,
    from: sender.email,
    to: ['master@s1samsung.com'],
    cc: [],
    subject: subject,
    hiworksResult: result || null
  };
}

function appendOrderNotificationMailLogV447_(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetName = '발주메일발송로그';
    let sheet = ss.getSheetByName(sheetName);
    const headers = [
      '일시', '요청ID', '상태', '고객번호', '마스터행', '계약번호', '계약행',
      '고객사명', '영업담당자', '발신자', '수신자', '메일제목', '하이웍스응답', '오류'
    ];
    if (!sheet) {
      sheet = ss.insertSheet(sheetName);
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      sheet.setFrozenRows(1);
    }
    sheet.appendRow([
      new Date(),
      data.requestId || '',
      data.status || '',
      data.customerNo || '',
      data.rowNo || '',
      data.contractNo || '',
      data.contractRowNo || '',
      data.company || '',
      data.salesRep || '',
      data.from || '',
      data.to || '',
      data.subject || '',
      data.hiworksResult || '',
      data.error || ''
    ]);
  } catch (err) {}
}

function escapeHtmlForOrderNotificationMailV447_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
