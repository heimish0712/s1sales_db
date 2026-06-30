/*******************************************************
 * 발주메일 수동 재발송 / 큐 직접 처리 도구
 *
 * 붙일 위치:
 * - 영업관리대장 Apps Script 또는 메일자동발송 Worker Apps Script
 * - 둘 중 어디에 붙여도 동작하도록 작성함.
 *
 * 동작 방식:
 * 1) Worker 코드 안에 mailWorkerSendOrderNotificationMailV447_ 함수가 있으면 직접 호출
 * 2) 없으면 MailMessage/HiworksMailer 클래스가 있으면 직접 발송
 * 3) 그것도 없으면 스크립트 속성의 MAIL_WORKER_WEBAPP_URL / MAIL_WORKER_SHARED_SECRET로 Worker WebApp 호출
 *
 * 기존 onOpen이 있으면 중복 생성하지 말고 기존 onOpen 안에 아래 한 줄만 추가:
 * addOrderMailManualMenu_();
 *******************************************************/

const ORDER_MAIL_MANUAL_CFG = {
  CONTRACT_SHEET_NAME: '수주확정/계약완료',
  QUEUE_SHEET_NAME: '발주메일발송큐',
  LOG_SHEET_NAME: '발주메일발송로그',
  TO: 'master@s1samsung.com',
  CC: '',
  WORKER_ACTION: 'sendOrderNotificationMail',

  HEADER_SCAN_MAX_ROWS: 5,

  CONTRACT_HEADERS: {
    contractNo: ['계약번호', '발주번호'],
    customerNo: ['고객번호'],
    company: ['고객사명', '회사명', '고객사'],
    salesRep: ['계약담당자', '영업담당자', '담당자']
  },

  QUEUE_HEADERS: {
    requestId: '요청ID',
    status: '상태',
    attempts: '시도횟수',
    createdAt: '생성일시',
    lastTriedAt: '최종시도일시',
    completedAt: '완료일시',
    error: '오류',
    customerNo: '고객번호',
    masterRow: '마스터행',
    contractNo: '계약번호',
    contractRow: '계약행',
    company: '고객사명',
    salesRep: '영업담당자'
  },

  LOG_HEADERS: [
    '일시', '단계', '상태', '요청ID', '고객번호', '마스터행', '계약번호', '계약행',
    '고객사명', '영업담당자', '수신자', '참조', 'WorkerAction', 'WorkerURL',
    '시도횟수', '큐행', '메일제목', '발신자', '요약', '오류', '상세JSON'
  ],

  SALES_REP_INFO_SHEETS: ['영업담당자', '영업담당자 정보'],
  SALES_REP_NAME_HEADERS: ['이름', '담당자', '영업담당자', '계약담당자', '사용자명', '성명'],
  SALES_REP_EMAIL_HEADERS: ['이메일', '메일', 'email', 'Email', '계정', '발신메일', '하이웍스ID']
};


/**
 * 메뉴 추가
 * 기존 onOpen 안에 addOrderMailManualMenu_(); 한 줄만 넣으면 됩니다.
 */
function addOrderMailManualMenu_() {
  SpreadsheetApp.getUi()
    .createMenu('발주메일')
    .addItem('선택 계약행 발주메일 재발송', 'resendOrderMailFromSelectedContractRow')
    .addItem('계약번호로 발주메일 재발송', 'resendOrderMailByContractNoPrompt')
    .addItem('대기/오류 큐 직접 재발송', 'resendPendingOrderMailQueueManual')
    .addSeparator()
    .addItem('발주메일 로그 시트 준비', 'prepareOrderMailManualLogSheet')
    .addToUi();
}


/**
 * 선택된 수주확정/계약완료 행 기준으로 발주메일 재발송
 */
function resendOrderMailFromSelectedContractRow() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const rowNo = sheet.getActiveRange() ? sheet.getActiveRange().getRow() : 0;

  try {
    const payload = buildOrderMailPayloadFromContractRow_(sheet, rowNo, '선택행');
    const result = sendOrderMailManual_(payload);

    appendOrderMailManualLog_({
      step: '수동재발송',
      status: '성공',
      requestId: payload.requestId,
      customerNo: payload.customerNo,
      masterRow: payload.rowNo,
      contractNo: payload.contractNo,
      contractRow: payload.contractRowNo,
      company: payload.company,
      salesRep: payload.salesRep,
      to: ORDER_MAIL_MANUAL_CFG.TO,
      cc: ORDER_MAIL_MANUAL_CFG.CC,
      action: ORDER_MAIL_MANUAL_CFG.WORKER_ACTION,
      workerUrl: result.workerUrl || '',
      attempts: '',
      queueRow: '',
      subject: payload.contractNo + '. ' + payload.company,
      from: result.from || '',
      summary: '선택 계약행 기준 발주메일 재발송 성공',
      detail: result
    });

    ss.toast('발주메일 재발송 완료: ' + payload.contractNo + '. ' + payload.company, '발주메일', 5);
  } catch (err) {
    appendOrderMailManualLog_({
      step: '수동재발송',
      status: '오류',
      summary: '선택 계약행 기준 발주메일 재발송 실패',
      error: err && err.message ? err.message : String(err),
      detail: { stack: err && err.stack ? String(err.stack).slice(0, 3000) : '' }
    });
    ui.alert('발주메일 재발송 실패\n\n' + (err && err.message ? err.message : String(err)));
  }
}


/**
 * 계약번호 입력받아 해당 계약행 발주메일 재발송
 */
function resendOrderMailByContractNoPrompt() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt('발주메일 재발송', '계약번호/발주번호를 입력하세요.', ui.ButtonSet.OK_CANCEL);
  if (res.getSelectedButton() !== ui.Button.OK) return;

  const contractNo = String(res.getResponseText() || '').trim();
  if (!contractNo) {
    ui.alert('계약번호가 비어 있습니다.');
    return;
  }

  try {
    const sheet = getOrderMailContractSheet_();
    const meta = detectOrderMailContractSheetMeta_(sheet);
    const found = findOrderMailContractRowByContractNo_(sheet, meta, contractNo);
    if (!found) throw new Error('수주확정/계약완료 시트에서 계약번호를 찾지 못했습니다: ' + contractNo);

    const payload = buildOrderMailPayloadFromContractRow_(sheet, found.rowNo, '계약번호입력');
    const result = sendOrderMailManual_(payload);

    appendOrderMailManualLog_({
      step: '수동재발송',
      status: '성공',
      requestId: payload.requestId,
      customerNo: payload.customerNo,
      masterRow: payload.rowNo,
      contractNo: payload.contractNo,
      contractRow: payload.contractRowNo,
      company: payload.company,
      salesRep: payload.salesRep,
      to: ORDER_MAIL_MANUAL_CFG.TO,
      cc: ORDER_MAIL_MANUAL_CFG.CC,
      action: ORDER_MAIL_MANUAL_CFG.WORKER_ACTION,
      workerUrl: result.workerUrl || '',
      subject: payload.contractNo + '. ' + payload.company,
      from: result.from || '',
      summary: '계약번호 입력 기준 발주메일 재발송 성공',
      detail: result
    });

    SpreadsheetApp.getActive().toast('발주메일 재발송 완료: ' + payload.contractNo + '. ' + payload.company, '발주메일', 5);
  } catch (err) {
    appendOrderMailManualLog_({
      step: '수동재발송',
      status: '오류',
      contractNo: contractNo,
      summary: '계약번호 입력 기준 발주메일 재발송 실패',
      error: err && err.message ? err.message : String(err),
      detail: { stack: err && err.stack ? String(err.stack).slice(0, 3000) : '' }
    });
    ui.alert('발주메일 재발송 실패\n\n' + (err && err.message ? err.message : String(err)));
  }
}


/**
 * 발주메일발송큐의 대기/오류 건을 직접 재발송
 * - 포탈 트리거가 ScriptLock busy로 스킵되어 큐가 대기 상태로 남았을 때 사용
 * - 완료된 건은 건드리지 않음
 */
function resendPendingOrderMailQueueManual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    const queueSheet = ss.getSheetByName(ORDER_MAIL_MANUAL_CFG.QUEUE_SHEET_NAME);
    if (!queueSheet) throw new Error('발주메일발송큐 시트를 찾지 못했습니다.');

    const meta = getOrderMailQueueHeaderMap_(queueSheet);
    const lastRow = queueSheet.getLastRow();
    if (lastRow < 2) {
      ss.toast('처리할 발주메일 큐가 없습니다.', '발주메일', 4);
      return;
    }

    const values = queueSheet.getRange(2, 1, lastRow - 1, queueSheet.getLastColumn()).getValues();
    let success = 0;
    let failed = 0;
    let skipped = 0;

    for (let i = 0; i < values.length; i++) {
      const rowNo = i + 2;
      const row = values[i];
      const status = String(getByHeaderIndex_(row, meta, 'status') || '').trim();
      const attempts = Number(getByHeaderIndex_(row, meta, 'attempts')) || 0;

      if (status !== '대기' && status !== '오류') {
        skipped++;
        continue;
      }

      const payload = {
        requestId: String(getByHeaderIndex_(row, meta, 'requestId') || '').trim(),
        customerNo: String(getByHeaderIndex_(row, meta, 'customerNo') || '').trim(),
        rowNo: Number(getByHeaderIndex_(row, meta, 'masterRow')) || 0,
        contractNo: String(getByHeaderIndex_(row, meta, 'contractNo') || '').trim(),
        contractRowNo: Number(getByHeaderIndex_(row, meta, 'contractRow')) || 0,
        company: String(getByHeaderIndex_(row, meta, 'company') || '').trim(),
        salesRep: String(getByHeaderIndex_(row, meta, 'salesRep') || '').trim(),
        to: [ORDER_MAIL_MANUAL_CFG.TO],
        cc: []
      };

      try {
        setQueueCellByKey_(queueSheet, meta, rowNo, 'status', '발송중');
        setQueueCellByKey_(queueSheet, meta, rowNo, 'attempts', attempts + 1);
        setQueueCellByKey_(queueSheet, meta, rowNo, 'lastTriedAt', new Date());
        setQueueCellByKey_(queueSheet, meta, rowNo, 'error', '');

        const result = sendOrderMailManual_(payload);

        setQueueCellByKey_(queueSheet, meta, rowNo, 'status', '완료');
        setQueueCellByKey_(queueSheet, meta, rowNo, 'completedAt', new Date());
        setQueueCellByKey_(queueSheet, meta, rowNo, 'error', '');

        appendOrderMailManualLog_({
          step: '큐수동처리',
          status: '성공',
          requestId: payload.requestId,
          customerNo: payload.customerNo,
          masterRow: payload.rowNo,
          contractNo: payload.contractNo,
          contractRow: payload.contractRowNo,
          company: payload.company,
          salesRep: payload.salesRep,
          to: ORDER_MAIL_MANUAL_CFG.TO,
          cc: ORDER_MAIL_MANUAL_CFG.CC,
          action: ORDER_MAIL_MANUAL_CFG.WORKER_ACTION,
          workerUrl: result.workerUrl || '',
          attempts: attempts + 1,
          queueRow: rowNo,
          subject: payload.contractNo + '. ' + payload.company,
          from: result.from || '',
          summary: '발주메일 큐 수동 처리 성공',
          detail: result
        });

        success++;
      } catch (err) {
        const msg = err && err.message ? err.message : String(err);
        setQueueCellByKey_(queueSheet, meta, rowNo, 'status', '오류');
        setQueueCellByKey_(queueSheet, meta, rowNo, 'attempts', attempts + 1);
        setQueueCellByKey_(queueSheet, meta, rowNo, 'lastTriedAt', new Date());
        setQueueCellByKey_(queueSheet, meta, rowNo, 'error', msg.slice(0, 1000));

        appendOrderMailManualLog_({
          step: '큐수동처리',
          status: '오류',
          requestId: payload.requestId,
          customerNo: payload.customerNo,
          masterRow: payload.rowNo,
          contractNo: payload.contractNo,
          contractRow: payload.contractRowNo,
          company: payload.company,
          salesRep: payload.salesRep,
          to: ORDER_MAIL_MANUAL_CFG.TO,
          cc: ORDER_MAIL_MANUAL_CFG.CC,
          action: ORDER_MAIL_MANUAL_CFG.WORKER_ACTION,
          attempts: attempts + 1,
          queueRow: rowNo,
          subject: payload.contractNo + '. ' + payload.company,
          summary: '발주메일 큐 수동 처리 실패',
          error: msg,
          detail: { stack: err && err.stack ? String(err.stack).slice(0, 3000) : '', payload: payload }
        });

        failed++;
      }
    }

    ss.toast('발주메일 큐 수동 처리 완료: 성공 ' + success + ' / 실패 ' + failed + ' / 스킵 ' + skipped, '발주메일', 7);
  } catch (err) {
    ui.alert('발주메일 큐 수동 처리 실패\n\n' + (err && err.message ? err.message : String(err)));
  }
}


/**
 * 발주메일 로그 시트 생성/헤더 보정
 */
function prepareOrderMailManualLogSheet() {
  getOrderMailManualLogSheet_();
  SpreadsheetApp.getActive().toast('발주메일발송로그 시트 준비 완료', '발주메일', 4);
}


function buildOrderMailPayloadFromContractRow_(sheet, rowNo, requestSource) {
  const meta = detectOrderMailContractSheetMeta_(sheet);
  if (!meta) throw new Error('수주확정/계약완료 헤더를 찾지 못했습니다.');
  if (rowNo <= meta.headerRow) throw new Error('계약 데이터 행을 선택해 주세요.');

  const values = sheet.getRange(rowNo, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const contractNo = String(values[meta.cols.contractNo - 1] || '').trim();
  const customerNo = String(values[meta.cols.customerNo - 1] || '').trim();
  const company = String(values[meta.cols.company - 1] || '').trim();
  const salesRep = String(values[meta.cols.salesRep - 1] || '').trim();

  if (!contractNo) throw new Error('계약번호가 비어 있습니다.');
  if (!company) throw new Error('고객사명이 비어 있습니다.');
  if (!salesRep) throw new Error('계약담당자/영업담당자가 비어 있습니다.');

  return {
    requestId: buildOrderMailManualRequestId_(customerNo, contractNo, company),
    customerNo: customerNo,
    rowNo: 0,
    contractNo: contractNo,
    contractRowNo: rowNo,
    company: company,
    salesRep: salesRep,
    to: [ORDER_MAIL_MANUAL_CFG.TO],
    cc: [],
    requestSource: requestSource || ''
  };
}


function sendOrderMailManual_(payload) {
  payload = payload || {};
  payload.to = [ORDER_MAIL_MANUAL_CFG.TO];
  payload.cc = [];

  if (!payload.contractNo) throw new Error('발주메일 발송 실패: 계약번호가 비어 있습니다.');
  if (!payload.company) throw new Error('발주메일 발송 실패: 고객사명이 비어 있습니다.');
  if (!payload.salesRep) throw new Error('발주메일 발송 실패: 영업담당자가 비어 있습니다.');

  // 1) Worker 내부 함수가 있으면 가장 안전하게 기존 Worker 로직 그대로 호출
  if (typeof mailWorkerSendOrderNotificationMailV447_ === 'function') {
    const result = mailWorkerSendOrderNotificationMailV447_(payload);
    result.via = 'directWorkerFunction';
    return result;
  }

  // 2) 메일 클래스가 같은 프로젝트에 있으면 직접 발송
  if (typeof MailMessage !== 'undefined' && typeof HiworksMailer !== 'undefined') {
    return sendOrderMailManualByLocalHiworks_(payload);
  }

  // 3) 포탈/영업관리대장 쪽이면 Worker WebApp 호출
  return callOrderMailWorkerWebAppManual_(payload);
}


function sendOrderMailManualByLocalHiworks_(payload) {
  const senderEmail = resolveOrderMailSalesRepEmailManual_(payload.salesRep);
  const subject = payload.contractNo + '. ' + payload.company;
  const bodyText = subject + '\n발주번호 생성 알림';
  const bodyHtml = bodyText.split('\n').map(escapeHtmlOrderMailManual_).join('<br>');

  const mail = new MailMessage({
    from: senderEmail,
    to: [ORDER_MAIL_MANUAL_CFG.TO],
    cc: [],
    subject: subject,
    bodyHtml: bodyHtml,
    attachments: []
  });

  const hiworksResult = new HiworksMailer(null).send(mail);
  return {
    ok: true,
    via: 'localHiworksClasses',
    from: senderEmail,
    to: [ORDER_MAIL_MANUAL_CFG.TO],
    cc: [],
    subject: subject,
    hiworksResult: hiworksResult || null
  };
}


function callOrderMailWorkerWebAppManual_(payload) {
  const props = PropertiesService.getScriptProperties();
  const workerUrl =
    props.getProperty('MAIL_WORKER_WEBAPP_URL') ||
    props.getProperty('PORTAL_MAIL_WORKER_WEBAPP_URL') ||
    props.getProperty('MAIL_WORKER_WEBAPP_URL_BACKUP') || '';

  const workerSecret =
    props.getProperty('MAIL_WORKER_SHARED_SECRET') ||
    props.getProperty('PORTAL_MAIL_WORKER_SHARED_SECRET') || '';

  if (!workerUrl) throw new Error('Worker WebApp URL이 없습니다. 스크립트 속성 MAIL_WORKER_WEBAPP_URL을 확인하세요.');
  if (!workerSecret) throw new Error('Worker Secret이 없습니다. 스크립트 속성 MAIL_WORKER_SHARED_SECRET을 확인하세요.');

  const response = UrlFetchApp.fetch(workerUrl, {
    method: 'post',
    contentType: 'application/json; charset=utf-8',
    payload: JSON.stringify({
      secret: workerSecret,
      action: ORDER_MAIL_MANUAL_CFG.WORKER_ACTION,
      payload: payload,
      client: {
        manualRetry: true,
        requestedAt: new Date().toISOString(),
        scriptId: (() => { try { return ScriptApp.getScriptId(); } catch (e) { return ''; } })()
      }
    }),
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = response.getResponseCode();
  const text = response.getContentText() || '';
  let data = null;
  try { data = JSON.parse(text); } catch (err) {}

  if (code < 200 || code >= 300) {
    throw new Error('Worker 호출 실패 HTTP ' + code + '\n응답: ' + text.slice(0, 2000));
  }
  if (!data || data.ok !== true) {
    throw new Error('Worker 발주메일 실패: ' + (data && data.message ? data.message : text).slice(0, 2500));
  }

  const result = data.result || data;
  result.via = 'workerWebApp';
  result.workerUrl = workerUrl;
  return result;
}


function resolveOrderMailSalesRepEmailManual_(salesRepName) {
  const name = String(salesRepName || '').trim();
  if (!name) throw new Error('영업담당자 이름이 비어 있어 발신자를 찾을 수 없습니다.');

  // 기존 Worker의 Resolver가 있으면 우선 사용
  if (typeof SalesRepResolver !== 'undefined') {
    try {
      const resolved = new SalesRepResolver(SpreadsheetApp.getActiveSpreadsheet()).resolve(name);
      if (resolved && resolved.email) return resolved.email;
    } catch (err) {
      // 아래 fallback으로 계속 진행
    }
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (let s = 0; s < ORDER_MAIL_MANUAL_CFG.SALES_REP_INFO_SHEETS.length; s++) {
    const sheet = ss.getSheetByName(ORDER_MAIL_MANUAL_CFG.SALES_REP_INFO_SHEETS[s]);
    if (!sheet) continue;

    const meta = detectNameEmailSheetMeta_(sheet);
    if (!meta) continue;

    const lastRow = sheet.getLastRow();
    if (lastRow <= meta.headerRow) continue;

    const values = sheet.getRange(meta.headerRow + 1, 1, lastRow - meta.headerRow, sheet.getLastColumn()).getDisplayValues();
    const normalizedTarget = normalizeOrderMailPersonName_(name);

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      const rowName = String(row[meta.nameCol - 1] || '').trim();
      const rowEmail = String(row[meta.emailCol - 1] || '').trim();
      if (!rowEmail) continue;

      if (normalizeOrderMailPersonName_(rowName) === normalizedTarget) {
        return normalizeOrderMailEmail_(rowEmail);
      }
    }
  }

  throw new Error('영업담당자 이메일을 찾지 못했습니다: ' + name + ' / 영업담당자 정보 시트 확인 필요');
}


function getOrderMailContractSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(ORDER_MAIL_MANUAL_CFG.CONTRACT_SHEET_NAME);
  if (!sheet) throw new Error('수주확정/계약완료 시트를 찾지 못했습니다.');
  return sheet;
}


function detectOrderMailContractSheetMeta_(sheet) {
  const lastCol = sheet.getLastColumn();
  const scanRows = Math.min(ORDER_MAIL_MANUAL_CFG.HEADER_SCAN_MAX_ROWS, sheet.getLastRow());

  for (let r = 1; r <= scanRows; r++) {
    const headers = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const cols = {
      contractNo: findHeaderColOrderMail_(headers, ORDER_MAIL_MANUAL_CFG.CONTRACT_HEADERS.contractNo),
      customerNo: findHeaderColOrderMail_(headers, ORDER_MAIL_MANUAL_CFG.CONTRACT_HEADERS.customerNo),
      company: findHeaderColOrderMail_(headers, ORDER_MAIL_MANUAL_CFG.CONTRACT_HEADERS.company),
      salesRep: findHeaderColOrderMail_(headers, ORDER_MAIL_MANUAL_CFG.CONTRACT_HEADERS.salesRep)
    };

    if (cols.contractNo > 0 && cols.company > 0 && cols.salesRep > 0) {
      return { headerRow: r, cols: cols };
    }
  }

  return null;
}


function findOrderMailContractRowByContractNo_(sheet, meta, contractNo) {
  const lastRow = sheet.getLastRow();
  if (lastRow <= meta.headerRow) return null;
  const values = sheet.getRange(meta.headerRow + 1, meta.cols.contractNo, lastRow - meta.headerRow, 1).getDisplayValues();
  const target = String(contractNo || '').trim();
  for (let i = 0; i < values.length; i++) {
    if (String(values[i][0] || '').trim() === target) return { rowNo: meta.headerRow + 1 + i };
  }
  return null;
}


function getOrderMailQueueHeaderMap_(sheet) {
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getDisplayValues()[0];
  const map = {};
  Object.keys(ORDER_MAIL_MANUAL_CFG.QUEUE_HEADERS).forEach(function(key) {
    const col = findHeaderColOrderMail_(headers, [ORDER_MAIL_MANUAL_CFG.QUEUE_HEADERS[key]]);
    if (col > 0) map[key] = col;
  });
  ['requestId', 'status', 'attempts', 'customerNo', 'contractNo', 'company', 'salesRep'].forEach(function(key) {
    if (!map[key]) throw new Error('발주메일발송큐 헤더를 찾지 못했습니다: ' + ORDER_MAIL_MANUAL_CFG.QUEUE_HEADERS[key]);
  });
  return map;
}


function getByHeaderIndex_(row, map, key) {
  const col = map[key];
  if (!col) return '';
  return row[col - 1];
}


function setQueueCellByKey_(sheet, map, rowNo, key, value) {
  const col = map[key];
  if (!col) return;
  sheet.getRange(rowNo, col).setValue(value);
}


function detectNameEmailSheetMeta_(sheet) {
  const lastCol = sheet.getLastColumn();
  const scanRows = Math.min(ORDER_MAIL_MANUAL_CFG.HEADER_SCAN_MAX_ROWS, sheet.getLastRow());

  for (let r = 1; r <= scanRows; r++) {
    const headers = sheet.getRange(r, 1, 1, lastCol).getDisplayValues()[0];
    const nameCol = findHeaderColOrderMail_(headers, ORDER_MAIL_MANUAL_CFG.SALES_REP_NAME_HEADERS);
    const emailCol = findHeaderColOrderMail_(headers, ORDER_MAIL_MANUAL_CFG.SALES_REP_EMAIL_HEADERS);
    if (nameCol > 0 && emailCol > 0) return { headerRow: r, nameCol: nameCol, emailCol: emailCol };
  }
  return null;
}


function getOrderMailManualLogSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(ORDER_MAIL_MANUAL_CFG.LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(ORDER_MAIL_MANUAL_CFG.LOG_SHEET_NAME);

  const headers = ORDER_MAIL_MANUAL_CFG.LOG_HEADERS;
  const current = sheet.getRange(1, 1, 1, headers.length).getDisplayValues()[0];
  const needHeader = headers.some(function(h, idx) { return String(current[idx] || '') !== h; });
  if (needHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
    try { sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#fce5cd'); } catch (err) {}
  }
  return sheet;
}


function appendOrderMailManualLog_(entry) {
  try {
    entry = entry || {};
    const sheet = getOrderMailManualLogSheet_();
    const detail = entry.detail == null ? '' : (typeof entry.detail === 'string' ? entry.detail : JSON.stringify(entry.detail));
    sheet.appendRow([
      new Date(),
      String(entry.step || '').trim(),
      String(entry.status || '').trim(),
      String(entry.requestId || '').trim(),
      String(entry.customerNo || '').trim(),
      entry.masterRow || '',
      String(entry.contractNo || '').trim(),
      entry.contractRow || '',
      String(entry.company || '').trim(),
      String(entry.salesRep || '').trim(),
      Array.isArray(entry.to) ? entry.to.join(',') : String(entry.to || '').trim(),
      Array.isArray(entry.cc) ? entry.cc.join(',') : String(entry.cc || '').trim(),
      String(entry.action || '').trim(),
      String(entry.workerUrl || '').trim(),
      entry.attempts || '',
      entry.queueRow || '',
      String(entry.subject || '').trim(),
      String(entry.from || '').trim(),
      String(entry.summary || '').trim(),
      String(entry.error || '').slice(0, 4000),
      detail.slice(0, 45000)
    ]);
  } catch (err) {
    try { Logger.log('발주메일 수동 로그 기록 실패: ' + (err && err.stack || err)); } catch (e) {}
  }
}


function buildOrderMailManualRequestId_(customerNo, contractNo, company) {
  const base = [String(customerNo || '').trim(), String(contractNo || '').trim(), String(company || '').trim()].join('|');
  const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, base);
  return 'ORDERMAIL-' + Utilities.base64EncodeWebSafe(digest).replace(/=+$/g, '').slice(0, 20);
}


function findHeaderColOrderMail_(headers, candidates) {
  const normalized = candidates.map(normalizeHeaderOrderMail_);
  for (let i = 0; i < headers.length; i++) {
    if (normalized.indexOf(normalizeHeaderOrderMail_(headers[i])) >= 0) return i + 1;
  }
  return -1;
}


function normalizeHeaderOrderMail_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/[()［］\[\]{}]/g, '')
    .toLowerCase()
    .trim();
}


function normalizeOrderMailPersonName_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/팀|대리|책임|차장|과장|부장|대표|님/g, '')
    .trim();
}


function normalizeOrderMailEmail_(value) {
  let s = String(value || '').trim();
  const m = s.match(/[A-Z0-9._%+-]+@[A-Z0-9.-]+\.[A-Z]{2,}/i);
  if (m) return m[0];
  if (s && s.indexOf('@') < 0) return s + '@s1samsung.com';
  return s;
}


function escapeHtmlOrderMailManual_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}
