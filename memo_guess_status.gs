/****************************************************
 * 메모 기반 "메모상 추측 상태값" 자동 채우기
 *
 * 상태값 기준:
 * - 견적제출완료: 견적서/자료/단가표 제출·발송된 건
 * - 장기 추진건: 올해 대상 아님, 내년 대상, 보류, 예산/추경, 장기 추적 건
 * - 고객 설득 중: 검토중, 비교중, 네고, 할인, 미팅, 재확인, 결정 대기 등 영업 진행 중
 * - 발주완료: 계약 의사 확정, 발주메일, 품의/결재, 나라장터 진행, 용역신청서 송부 등
 * - 계약완료: 계약서/용역신청서/선임신고서/위임장/사업자등록증 등 계약서류 수취·계약체결 완료
 * - 데이터확인필요: 연락처/주소/연면적/담당자/중복/건축물대장/메일오류 등 데이터 확인 필요
 * - 장기미접촉: 장기미접촉 명시 또는 마지막 컨택일이 오래된 건
 * - 수주실패: 타사선정, 타사계약, 직접수행, 기존업체 진행, 거절, 가격 탈락 등
 ****************************************************/

const MEMO_GUESS_DROPDOWN_CONFIG = {
  TARGET_SHEET_NAME: '', // 비워두면 현재 활성 시트 우선. 강제하려면 '마스터시트' 등 입력
  HEADER_SCAN_ROWS: 10,

  SOURCE_HEADER: '메모',
  TARGET_HEADER: '메모상 추측 상태값',

  AUTO_CREATE_TARGET_HEADER: true,
  WRITE_BATCH_SIZE: 5000,

  // 장기미접촉 판단 기준일. 현재 날짜 기준 N일 이상 최신 컨택 없음
  LONG_NO_CONTACT_DAYS: 45,

  // 날짜 기반 장기미접촉 자동 판단 사용 여부
  USE_DATE_BASED_LONG_NO_CONTACT: true,

  STATUS: {
    QUOTE_SENT: '견적제출완료',
    LONG_TERM: '장기 추진건',
    PERSUADING: '고객 설득 중',
    ORDER_DONE: '발주완료',
    CONTRACT_DONE: '계약완료',
    NEED_DATA_CHECK: '데이터확인필요',
    LONG_NO_CONTACT: '장기미접촉',
    LOST: '수주실패',
    NEED_MANUAL: '!!상태지정필요!!'
  }
};


/**
 * 전체 재계산: 기존 값도 전부 덮어씀
 */
function fillMemoGuessStatus_ByDropdownRules_AllRows() {
  fillMemoGuessStatus_ByDropdownRules_({ blankOnly: false });
}


/**
 * 빈칸만 채우기: 기존 값 보존
 */
function fillMemoGuessStatus_ByDropdownRules_BlankOnly() {
  fillMemoGuessStatus_ByDropdownRules_({ blankOnly: true });
}


function fillMemoGuessStatus_ByDropdownRules_(options) {
  const lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('다른 작업이 실행 중입니다. 잠시 후 다시 실행해주세요.');
  }

  try {
    const opt = options || {};
    const blankOnly = opt.blankOnly === true;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = getMemoGuessTargetSheet_(ss);

    const headerInfo = getMemoGuessHeaderInfo_(sheet);
    const headerRow = headerInfo.headerRow;
    const memoCol = headerInfo.memoCol;
    const targetCol = headerInfo.targetCol;

    const lastRow = sheet.getLastRow();
    const dataStartRow = headerRow + 1;

    if (lastRow < dataStartRow) {
      ss.toast('처리할 데이터가 없습니다.', '메모상 추측 상태값', 5);
      return;
    }

    const rowCount = lastRow - headerRow;
    const memoValues = sheet.getRange(dataStartRow, memoCol, rowCount, 1).getDisplayValues();
    const oldValues = sheet.getRange(dataStartRow, targetCol, rowCount, 1).getDisplayValues();

    const output = [];
    const stats = {};
    let changedCount = 0;
    let skippedCount = 0;

    for (let i = 0; i < rowCount; i++) {
      const memo = memoValues[i][0] || '';
      const oldValue = oldValues[i][0] || '';

      if (blankOnly && String(oldValue).trim() !== '') {
        output.push([oldValue]);
        skippedCount++;
        continue;
      }

      const guessed = guessMemoDropdownStatus_(memo);
      output.push([guessed]);

      stats[guessed] = (stats[guessed] || 0) + 1;
      if (oldValue !== guessed) changedCount++;
    }

    writeColumnInBatches_(sheet, dataStartRow, targetCol, output, MEMO_GUESS_DROPDOWN_CONFIG.WRITE_BATCH_SIZE);

    const summary = Object.keys(stats)
      .sort()
      .map(k => `${k} ${stats[k]}건`)
      .join(' / ');

    ss.toast(
      `완료: 변경 ${changedCount}건` +
      (blankOnly ? `, 기존값 보존 ${skippedCount}건` : '') +
      (summary ? ` / ${summary}` : ''),
      '메모상 추측 상태값',
      10
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * 메모 1건 → 상태값 1개 추측
 */
function guessMemoDropdownStatus_(memo) {
  const S = MEMO_GUESS_DROPDOWN_CONFIG.STATUS;

  const raw = String(memo || '').trim();
  if (!raw) return S.NEED_MANUAL;

  const full = normalizeMemoText_(raw);
  const recentRaw = getRecentMemoText_(raw, 22, 2200);
  const recent = normalizeMemoText_(recentRaw);

  if (isTestOnlyMemo_(full)) {
    return S.NEED_MANUAL;
  }

  /****************************************************
   * 1. 수주실패
   * 타사/기존업체/직접수행/거절은 계약 관련 단어보다 우선합니다.
   ****************************************************/
  if (hasAny_(recent, [
    /타\s*사.{0,15}(계약|선정|진행|완료|결정|했|함|됨)/,
    /(다른|타|기존|관내|저렴한|서울|광주|지역|보람정보통신|기계설비|유지보수|관리)\s*(업체|곳|회사|수행사).{0,18}(계약|선정|진행|완료|결정|하기로|함|됨|했다|하셨)/,
    /(계약|선정).{0,10}(완료|함|됨|했다|하셨).{0,12}(타사|다른|기존|관내|저렴한|서울|광주|보람정보통신|기계설비|유지보수|관리)/,
    /(업체|수행사).{0,8}(선정|계약).{0,8}(완료|함|됨|했다|하셨)/,
    /다른\s*곳.{0,12}(계약|선정|진행|완료|결정)/,
    /다른\s*업체.{0,12}(계약|선정|진행|완료|결정)/,
    /기존\s*업체.{0,12}(계약|선정|진행|완료|결정)/,
    /관내\s*업체.{0,12}(계약|선정|진행|완료|결정)/,
    /(직접|자체).{0,5}(수행|진행|관리).{0,8}(함|한다|하기로|결정|예정)?/,
    /(수주\s*실패|거절|탈락|안\s*한다|안\s*할|하지\s*않|관심\s*없|계약\s*못|전화\s*하지\s*마|연락\s*하지\s*마)/,
    /(가격|금액).{0,12}(차이|비싸|안맞|안\s*맞|탈락|어렵)/,
    /타사선정완료|타사계약완료|업체선정완료/,
    /본사에서\s*일괄\s*진행/,
    /소방청에서\s*일괄/,
    /시청에서\s*알아서\s*한다/
  ])) {
    return S.LOST;
  }

  /****************************************************
   * 2. 계약완료
   * 계약서류를 실제로 받았거나 계약체결이 끝난 상태
   ****************************************************/
  if (hasAny_(recent, [
    /계약\s*완료/,
    /계약\s*체결/,
    /계약서.{0,12}(수취|받|받음|보내옴|보내\s*옴|회신|작성\s*완료|작성완료|직인|날인)/,
    /(용역\s*신청서|선임\s*신고서|위임장|사업자\s*등록증|인감|사용인감|통장\s*사본).{0,14}(수취|받|받음|보내옴|보내\s*옴|회신|제출|도착)/,
    /(착수계|착공계|계약보증서|청렴계약|보안서약서|안전보건관리계획서|완납증명서).{0,14}(수취|받|받음|제출|회신|도착)/,
    /(서류|원본).{0,12}(다\s*받|수취|도착|회신\s*완료)/,
    /계약서류.{0,10}(완료|수취|받음)/,
    /계약번호\s*생성/
  ])) {
    return S.CONTRACT_DONE;
  }

  /****************************************************
   * 3. 발주완료
   * 계약 의사 확정, 발주 의사, 결재/품의/나라장터 진행 등
   ****************************************************/
  if (hasAny_(recent, [
    /발주\s*완료/,
    /발주\s*메일|발주메일/,
    /발주\s*의사/,
    /(계약|진행).{0,8}(하신다고|한다고|하기로|예정|의사|확정|진행|올림|올렸|결정)/,
    /(당사|에스원|케이제이|kj|삼구|일신).{0,15}(결정|진행|계약\s*예정|계약\s*진행|계약하기로|계약\s*하신다고)/,
    /(품의|기안|결재|결제).{0,12}(올림|올렸|올려|완료|남|받|진행|예정)/,
    /(나라\s*장터|수의\s*계약).{0,15}(진행|올림|올렸|응답|체결|계약|요청)/,
    /(계약\s*부서|계약\s*담당|구매팀).{0,15}(연락|진행|넘김|넘겼|요청)/,
    /(용역\s*신청서|선임\s*신고서|위임장).{0,12}(요청|발송\s*요청|보내\s*달|작성\s*요청)/,
    /사업자\s*등록증.{0,12}(요청|보내\s*달)/,
    /내방.{0,10}(계약|작성)/,
    /방문계약/,
    /발주메일\s*발송완료/
  ])) {
    return S.ORDER_DONE;
  }

  /****************************************************
   * 4. 장기 추진건
   * 이번년도 대상 아님, 내년, 유예, 보류, 예산 문제, 장기 추적
   ****************************************************/
  if (hasAny_(recent, [
    /장기\s*추진/,
    /(올해|금년).{0,8}(대상\s*아님|해당\s*안됨|안됨|아님)/,
    /(내년|2027|27년|11월|12월|하반기|10월쯤|추후|나중에).{0,16}(연락|진행|계약|검토|예정|준비|다시)/,
    /(보류|중단|유예|관망|홀딩|잠시\s*보류)/,
    /(예산\s*없|예산\s*부족|예산\s*확보\s*안|예산\s*미확보|추경)/,
    /(기간|시간|시기).{0,10}(남|있|여유|아직)/,
    /아직\s*시간/,
    /아직\s*검토\s*전/,
    /검토\s*전/,
    /대상\s*여부.{0,10}(확인|미정)/,
    /1년\s*유보/,
    /과태료.{0,8}유예/,
    /개도기간/
  ])) {
    return S.LONG_TERM;
  }

  /****************************************************
   * 5. 데이터확인필요
   * 정보 자체를 확인해야 하는 건
   ****************************************************/
  if (hasAny_(recent, [
    /데이터\s*확인\s*필요/,
    /확인\s*필요/,
    /주소\s*확인\s*필요/,
    /연면적.{0,15}(확인|수정|다름|상이|재확인|변경|틀림|모름|불명확)/,
    /(주소|상호|회사명|법인명|사업자등록번호|대표자|담당자|전화번호|직통번호|메일|이메일).{0,12}(확인|수정|오류|다름|변경|불명확|모름)/,
    /연락처\s*오류|전화번호\s*오류|없는\s*번호/,
    /메일\s*주소.{0,8}(오류|수정|확인|잘못)/,
    /불량\s*발송자/,
    /중복|시트\s*중복|데이터\s*병합|중복건\s*삭제|병합\s*완료/,
    /건축물대장.{0,12}(확인|발급|안됨|필요)/,
    /대상\s*여부.{0,12}(확인|재확인|문의)/,
    /지자체.{0,12}(확인|문의)/,
    /공개\s*연면적.{0,12}(차이|다름|상이|확인)/,
    /정확한\s*연면적/,
    /담당자.{0,12}(누군지|확인|모름|변경|퇴사|공석)/
  ])) {
    return S.NEED_DATA_CHECK;
  }

  /****************************************************
   * 6. 고객 설득 중
   * 견적 제출 후에도 실제 영업 협의/검토/비교/네고가 진행 중인 건
   ****************************************************/
  if (hasAny_(recent, [
    /고객\s*설득\s*중/,
    /(검토\s*중|비교\s*중|업체\s*선정\s*전|결정\s*전|결정\s*안|미정|취합\s*중)/,
    /(네고|할인|추가\s*할인|금액\s*조정|가격\s*조정|최종\s*견적|재견적)/,
    /(미팅|방문|화상\s*회의|실사|상담).{0,12}(요청|진행|예정|완료|잡|조율)/,
    /(확인\s*후\s*연락|연락\s*준다고|연락\s*주신다고|결정되면\s*연락)/,
    /(긍정|호의|관심|잘\s*부탁|좋은\s*소식)/,
    /(재발송|다시\s*보내|다시\s*발송).{0,12}(요청|완료|예정)?/,
    /(과업\s*지시서|수행사\s*정보|실적|샘플|비교견적).{0,12}(요청|발송|검토)/,
    /(담당자\s*부재|부재|미수신|출장|휴가|연차|회의\s*중|교육\s*중|외근|자리\s*비움|다시\s*전화|재연락|번호\s*전달|연락처\s*남김)/,
    /리마인드/,
    /다음주.{0,8}연락/,
    /월말.{0,8}연락/
  ])) {
    return S.PERSUADING;
  }

  /****************************************************
   * 7. 견적제출완료
   * 단순 견적/자료 제출 완료
   ****************************************************/
  if (hasAny_(recent, [
    /견적\s*제출\s*완료/,
    /견적제출완료/,
    /(견적서|견적|단가표|안내자료|안내장|자료|제안서).{0,14}(발송|송부|제출|보냄|보내드림|메일\s*전송|메일전송|팩스\s*발송|fax)/,
    /(자료\s*발송|자료발송)/,
    /메일\s*발송/,
    /팩스\s*발송/,
    /수기견적.{0,12}(발송|제출|요청)/
  ])) {
    return S.QUOTE_SENT;
  }

  /****************************************************
   * 8. 장기미접촉
   * 명시되어 있거나 마지막 날짜가 오래된 건
   ****************************************************/
  if (hasAny_(full, [
    /장기\s*미접촉/,
    /장기미접촉/
  ])) {
    return S.LONG_NO_CONTACT;
  }

  if (MEMO_GUESS_DROPDOWN_CONFIG.USE_DATE_BASED_LONG_NO_CONTACT) {
    const latestDate = extractLatestMemoDate_(raw);
    if (latestDate) {
      const today = new Date();
      const diffDays = Math.floor((stripTime_(today).getTime() - stripTime_(latestDate).getTime()) / 86400000);

      if (diffDays >= MEMO_GUESS_DROPDOWN_CONFIG.LONG_NO_CONTACT_DAYS) {
        return S.LONG_NO_CONTACT;
      }
    }
  }

  /****************************************************
   * 9. 전체 메모 보정
   * 최근 줄에 없더라도 전체 메모상 명확한 견적 제출 흔적이 있으면 견적제출완료
   ****************************************************/
  if (hasAny_(full, [
    /(견적서|견적|단가표|안내자료|자료|제안서).{0,14}(발송|송부|제출|보냄|보내드림|메일\s*전송|메일전송|팩스\s*발송|fax)/,
    /자료\s*발송|자료발송|견적제출완료/
  ])) {
    return S.QUOTE_SENT;
  }

  return S.NEED_MANUAL;
}


/****************************************************
 * 이하 공통 유틸
 ****************************************************/

function getMemoGuessTargetSheet_(ss) {
  const cfg = MEMO_GUESS_DROPDOWN_CONFIG;

  if (cfg.TARGET_SHEET_NAME) {
    const named = ss.getSheetByName(cfg.TARGET_SHEET_NAME);
    if (!named) throw new Error('대상 시트를 찾을 수 없습니다: ' + cfg.TARGET_SHEET_NAME);
    return named;
  }

  const active = ss.getActiveSheet();
  if (tryFindMemoGuessHeaderInfo_(active)) return active;

  const candidates = ['마스터시트', '마스터시트(신규)', '영업관리대장'];
  for (let i = 0; i < candidates.length; i++) {
    const s = ss.getSheetByName(candidates[i]);
    if (s && tryFindMemoGuessHeaderInfo_(s)) return s;
  }

  const sheets = ss.getSheets();
  for (let j = 0; j < sheets.length; j++) {
    if (tryFindMemoGuessHeaderInfo_(sheets[j])) return sheets[j];
  }

  throw new Error('"메모" 헤더가 있는 시트를 찾지 못했습니다.');
}


function getMemoGuessHeaderInfo_(sheet) {
  const info = tryFindMemoGuessHeaderInfo_(sheet);

  if (!info) {
    throw new Error(
      '헤더를 찾지 못했습니다. 시트명=' + sheet.getName() +
      ', 필요 헤더=' + MEMO_GUESS_DROPDOWN_CONFIG.SOURCE_HEADER
    );
  }

  if (!info.targetCol && MEMO_GUESS_DROPDOWN_CONFIG.AUTO_CREATE_TARGET_HEADER) {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(info.headerRow, newCol).setValue(MEMO_GUESS_DROPDOWN_CONFIG.TARGET_HEADER);
    info.targetCol = newCol;
  }

  if (!info.targetCol) {
    throw new Error('대상 헤더를 찾지 못했습니다: ' + MEMO_GUESS_DROPDOWN_CONFIG.TARGET_HEADER);
  }

  return info;
}


function tryFindMemoGuessHeaderInfo_(sheet) {
  if (!sheet) return null;

  const cfg = MEMO_GUESS_DROPDOWN_CONFIG;
  const scanRows = Math.min(cfg.HEADER_SCAN_ROWS, Math.max(sheet.getLastRow(), 1));
  const lastCol = Math.max(sheet.getLastColumn(), 1);

  const values = sheet.getRange(1, 1, scanRows, lastCol).getDisplayValues();

  for (let r = 0; r < values.length; r++) {
    const row = values[r];

    let memoCol = 0;
    let targetCol = 0;

    for (let c = 0; c < row.length; c++) {
      const h = memoGuessNormalizeHeader_(row[c]);

      if (h === memoGuessNormalizeHeader_(cfg.SOURCE_HEADER)) memoCol = c + 1;
      if (h === memoGuessNormalizeHeader_(cfg.TARGET_HEADER)) targetCol = c + 1;
    }

    if (memoCol) {
      return {
        headerRow: r + 1,
        memoCol: memoCol,
        targetCol: targetCol
      };
    }
  }

  return null;
}


function memoGuessNormalizeHeader_(value) {
  return String(value || '')
    .replace(/\s+/g, '')
    .replace(/[　]/g, '')
    .trim()
    .toLowerCase();
}


function normalizeMemoText_(value) {
  return String(value || '')
    .replace(/\r\n/g, '\n')
    .replace(/\r/g, '\n')
    .replace(/[ \t]+/g, ' ')
    .replace(/[Ａ-Ｚａ-ｚ０-９]/g, ch => String.fromCharCode(ch.charCodeAt(0) - 0xFEE0))
    .toLowerCase()
    .trim();
}


function getRecentMemoText_(memo, maxLines, maxChars) {
  const raw = String(memo || '').replace(/\r\n/g, '\n').replace(/\r/g, '\n');

  const lines = raw
    .split('\n')
    .map(line => line.trim())
    .filter(line => line !== '');

  let recent = lines.slice(-maxLines).join('\n');

  if (recent.length > maxChars) {
    recent = recent.slice(recent.length - maxChars);
  }

  return recent;
}


function hasAny_(text, patterns) {
  const value = String(text || '');

  for (let i = 0; i < patterns.length; i++) {
    if (patterns[i].test(value)) return true;
  }

  return false;
}


function isTestOnlyMemo_(text) {
  const value = String(text || '');
  if (!value) return false;

  const hasTest = /(테스트|\[테스트\]|방수원테스트|수정테스트|동시수정테스트|ㅇㄹㅇ|ㄴㅇ)/.test(value);
  if (!hasTest) return false;

  const hasRealSignal = /(계약|견적|타사|자료|검토|발송|부재|담당자|용역신청서|나라장터|사업자등록증|발주)/.test(value);

  return hasTest && !hasRealSignal;
}


function writeColumnInBatches_(sheet, startRow, col, values, batchSize) {
  const size = Math.max(Number(batchSize) || 5000, 1);

  for (let offset = 0; offset < values.length; offset += size) {
    const chunk = values.slice(offset, offset + size);
    sheet.getRange(startRow + offset, col, chunk.length, 1).setValues(chunk);
  }

  SpreadsheetApp.flush();
}


function extractLatestMemoDate_(memo) {
  const text = String(memo || '');
  const now = new Date();
  const currentYear = now.getFullYear();

  const dates = [];

  // 26.07.03 / 26.07.03. / 26/07/03 / 2026.07.03
  const reFull = /(?:^|[^\d])((?:20)?\d{2})[.\-/]\s*(\d{1,2})[.\-/]\s*(\d{1,2})(?:\s|\.|일|$)/g;
  let m;

  while ((m = reFull.exec(text)) !== null) {
    let y = Number(m[1]);
    const mo = Number(m[2]);
    const d = Number(m[3]);

    if (y < 100) y += 2000;

    const dt = safeDate_(y, mo, d);
    if (dt) dates.push(dt);
  }

  // 7/3, 6.29 같은 연도 없는 날짜
  const reShort = /(?:^|[^\d])(\d{1,2})[./](\d{1,2})(?:\s|\.|일|$)/g;

  while ((m = reShort.exec(text)) !== null) {
    const mo = Number(m[1]);
    const d = Number(m[2]);

    const dt = safeDate_(currentYear, mo, d);
    if (dt) dates.push(dt);
  }

  if (dates.length === 0) return null;

  dates.sort((a, b) => b.getTime() - a.getTime());
  return dates[0];
}


function safeDate_(year, month, day) {
  if (!year || !month || !day) return null;
  if (month < 1 || month > 12) return null;
  if (day < 1 || day > 31) return null;

  const dt = new Date(year, month - 1, day);

  if (
    dt.getFullYear() !== year ||
    dt.getMonth() !== month - 1 ||
    dt.getDate() !== day
  ) {
    return null;
  }

  return dt;
}


function stripTime_(date) {
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}