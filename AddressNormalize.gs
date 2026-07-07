/*******************************************************
 * 주소 정규화 V5 - 최신 법정동코드 기준 + 검증 컬럼 반영
 *
 * 대상 시트: 마스터시트(신규)
 * 헤더 행: 2행
 * 데이터 시작: 3행
 * 원본 주소 헤더: 고객사 상세 주소
 * 결과 헤더: 주소정규화
 * 추가/갱신 헤더: 주소확인필요, 주소정규화상태, 주소정규화비고
 *******************************************************/

var ADDRESS_NORM_V5_CONFIG = {
  MASTER_SHEET_NAME: '마스터시트(신규)',
  LEGAL_DONG_SHEET_NAME: '법정동코드',

  HEADER_ROW: 2,
  DATA_START_ROW: 3,

  SOURCE_ADDRESS_HEADER: '고객사 상세 주소',
  NORMALIZED_ADDRESS_HEADER: '주소정규화',
  ADDRESS_CHECK_HEADER: '주소확인필요',
  NORMALIZE_STATUS_HEADER: '주소정규화상태',
  NORMALIZE_NOTE_HEADER: '주소정규화비고',

  CLEAR_ADDRESS_CHECK_WHEN_NORMAL: false,

  FORCE_PROVINCE_ALIAS_TO_CURRENT: {
    '전남': '전남광주통합특별시',
    '전라남도': '전남광주통합특별시',
    '광주': '전남광주통합특별시',
    '광주시': '전남광주통합특별시',
    '광주광역시': '전남광주통합특별시',

    '서울': '서울특별시',
    '서울시': '서울특별시',
    '서울특별시': '서울특별시',

    '부산': '부산광역시',
    '부산시': '부산광역시',
    '부산직할시': '부산광역시',
    '부산광역시': '부산광역시',

    '대구': '대구광역시',
    '대구시': '대구광역시',
    '대구직할시': '대구광역시',
    '대구광역시': '대구광역시',

    '인천': '인천광역시',
    '인천시': '인천광역시',
    '인천직할시': '인천광역시',
    '인천광역시': '인천광역시',

    '대전': '대전광역시',
    '대전시': '대전광역시',
    '대전직할시': '대전광역시',
    '대전광역시': '대전광역시',

    '울산': '울산광역시',
    '울산시': '울산광역시',
    '울산광역시': '울산광역시',

    '세종': '세종특별자치시',
    '세종시': '세종특별자치시',
    '세종특별시': '세종특별자치시',
    '세종특별자치시': '세종특별자치시',

    '경기': '경기도',
    '경기도': '경기도',

    '강원': '강원특별자치도',
    '강원도': '강원특별자치도',
    '강원특별자치도': '강원특별자치도',

    '충북': '충청북도',
    '충청북도': '충청북도',

    '충남': '충청남도',
    '충청남도': '충청남도',

    '전북': '전북특별자치도',
    '전라북도': '전북특별자치도',
    '전북특별자치도': '전북특별자치도',

    '경북': '경상북도',
    '경상북도': '경상북도',

    '경남': '경상남도',
    '경상남도': '경상남도',

    '제주': '제주특별자치도',
    '제주도': '제주특별자치도',
    '제주특별자치도': '제주특별자치도'
  }
};


/**
 * 일괄 실행 함수
 */
function normalizeMasterAddressColumnOnce() {
  var lock = LockService.getDocumentLock();

  if (!lock.tryLock(10000)) {
    throw new Error('다른 작업이 실행 중입니다. 잠시 후 다시 실행해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = ss.getSheetByName(ADDRESS_NORM_V5_CONFIG.MASTER_SHEET_NAME);

    if (!sheet) {
      throw new Error('마스터시트(신규) 시트를 찾지 못했습니다.');
    }

    var lastRow = sheet.getLastRow();

    if (lastRow < ADDRESS_NORM_V5_CONFIG.DATA_START_ROW) {
      SpreadsheetApp.getUi().alert('정규화할 데이터가 없습니다.');
      return;
    }

    var sourceCol = findHeaderColAddressNormV5_(sheet, ADDRESS_NORM_V5_CONFIG.SOURCE_ADDRESS_HEADER);
    if (sourceCol < 1) {
      throw new Error('"고객사 상세 주소" 헤더를 찾지 못했습니다.');
    }

    var normalizedCol = ensureHeaderColAddressNormV5_(sheet, ADDRESS_NORM_V5_CONFIG.NORMALIZED_ADDRESS_HEADER);
    var checkCol = ensureHeaderColAddressNormV5_(sheet, ADDRESS_NORM_V5_CONFIG.ADDRESS_CHECK_HEADER);
    var statusCol = ensureHeaderColAddressNormV5_(sheet, ADDRESS_NORM_V5_CONFIG.NORMALIZE_STATUS_HEADER);
    var noteCol = ensureHeaderColAddressNormV5_(sheet, ADDRESS_NORM_V5_CONFIG.NORMALIZE_NOTE_HEADER);

    var legalIndex = buildLegalDongIndexAddressNormV5_(ss);
    if (!legalIndex.available) {
      throw new Error('법정동코드 시트를 읽지 못했습니다. 시트명/헤더를 확인하세요.');
    }

    var rowCount = lastRow - ADDRESS_NORM_V5_CONFIG.DATA_START_ROW + 1;

    var sourceValues = sheet
      .getRange(ADDRESS_NORM_V5_CONFIG.DATA_START_ROW, sourceCol, rowCount, 1)
      .getDisplayValues();

    var existingCheckValues = sheet
      .getRange(ADDRESS_NORM_V5_CONFIG.DATA_START_ROW, checkCol, rowCount, 1)
      .getValues();

    var normalizedOutput = [];
    var checkOutput = [];
    var statusOutput = [];
    var noteOutput = [];

    var blankCount = 0;
    var changedCount = 0;
    var legalFixedCount = 0;
    var memoCutCount = 0;
    var needCheckCount = 0;
    var ambiguousCount = 0;

    for (var i = 0; i < rowCount; i++) {
      var raw = sourceValues[i][0];
      var result = normalizeAddressNormV5_(raw, legalIndex);

      normalizedOutput.push([result.normalized]);

      var oldCheck = existingCheckValues[i] ? existingCheckValues[i][0] : false;
      var newCheck;

      if (result.needCheck) {
        newCheck = true;
      } else if (ADDRESS_NORM_V5_CONFIG.CLEAR_ADDRESS_CHECK_WHEN_NORMAL) {
        newCheck = false;
      } else {
        newCheck = oldCheck;
      }

      checkOutput.push([newCheck]);
      statusOutput.push([result.status]);
      noteOutput.push([result.note]);

      if (!String(raw || '').trim()) blankCount++;
      if (cleanSpacesAddressNormV5_(raw) !== result.normalized) changedCount++;
      if (result.legalFixed) legalFixedCount++;
      if (result.memoCut) memoCutCount++;
      if (result.needCheck) needCheckCount++;
      if (result.ambiguous) ambiguousCount++;
    }

    sheet.getRange(ADDRESS_NORM_V5_CONFIG.DATA_START_ROW, normalizedCol, rowCount, 1).setValues(normalizedOutput);
    sheet.getRange(ADDRESS_NORM_V5_CONFIG.DATA_START_ROW, checkCol, rowCount, 1).setValues(checkOutput);
    sheet.getRange(ADDRESS_NORM_V5_CONFIG.DATA_START_ROW, statusCol, rowCount, 1).setValues(statusOutput);
    sheet.getRange(ADDRESS_NORM_V5_CONFIG.DATA_START_ROW, noteCol, rowCount, 1).setValues(noteOutput);

    SpreadsheetApp.getUi().alert(
      [
        '주소정규화 V5 완료',
        '',
        '대상 행 수: ' + rowCount,
        '공란: ' + blankCount,
        '정규화 변경: ' + changedCount,
        '법정동코드 보정: ' + legalFixedCount,
        '메모/복수주소 절단: ' + memoCutCount,
        '주소확인필요 신규 판정: ' + needCheckCount,
        '애매함 판정: ' + ambiguousCount,
        '',
        '법정동코드 존재 행: ' + legalIndex.existCount,
        '법정동 기준 광역명 수: ' + Object.keys(legalIndex.currentProvinceSet).length,
        '',
        ADDRESS_NORM_V5_CONFIG.CLEAR_ADDRESS_CHECK_WHEN_NORMAL
          ? '정상 주소의 주소확인필요 체크 해제: O'
          : '정상 주소의 기존 주소확인필요 체크 보존: O'
      ].join('\n')
    );

  } finally {
    lock.releaseLock();
  }
}


/**
 * 전체 돌리기 전 샘플 테스트
 */
function testAddressNormalizeV5Samples() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var legalIndex = buildLegalDongIndexAddressNormV5_(ss);

  var samples = [
    '전북특별자치도 군산시 외항6길 45',
    '전북특별자치도 군산시 외항로 82',
    '대구광역시 수성구 달구 벌대로541길 36',
    '서울특별시 강남구 압구 정로30길 45',
    '세종특별자치시 국책연구 원3로 12',
    '충청북도 청주시 청원구 오창읍 연구 단지로 53',
    '충청남도 공주시 유구읍 유구 마곡사로 136-23',
    '전라남도 목포시 대양산단로97번길 72',
    '광주광역시 북구 우치로 77',
    '경기도 화성시 향남읍 제약공단2길 45',
    '경기도 화성시 방교동 839-4',
    '경기도 과천시과천대로7다길60',
    '대전광역시 유성구봉명동 대학로 60 매드블럭',
    '흥덕구 직지대로 257',
    '김해시',
    '서울특별시 ㅇㅇ동 523',
    '대전광역시 유성구 몰라동 몰라리',
    '부산광역시 해운대구 해운대해변로 43 / 교육관 부산광역시 해운대구 해운대해변로 47'
  ];

  var msg = samples.map(function (s) {
    var r = normalizeAddressNormV5_(s, legalIndex);
    return s + '\n→ ' + r.normalized + '\n[' + r.status + '] ' + r.note;
  }).join('\n\n');

  Logger.log(msg);
  SpreadsheetApp.getUi().alert(msg);
}


/**
 * 주소 1건 정규화
 */
function normalizeAddressNormV5_(rawValue, legalIndex) {
  var raw = String(rawValue || '').trim();

  if (!raw) {
    return {
      normalized: '',
      needCheck: true,
      status: '공란',
      note: '주소 공란',
      legalFixed: false,
      memoCut: false,
      ambiguous: false
    };
  }

  var picked = pickPrimaryAddressTextAddressNormV5_(raw);
  var s = picked.value;

  if (!s) {
    return {
      normalized: '',
      needCheck: true,
      status: '공란',
      note: '주소 후보를 추출하지 못함',
      legalFixed: false,
      memoCut: picked.memoCut,
      ambiguous: false
    };
  }

  s = normalizeBasicAddressTextAddressNormV5_(s);
  s = normalizeProvinceTokenOnlyAddressNormV5_(s, legalIndex);
  s = splitStickyAddressTokensAddressNormV5_(s, legalIndex);
  s = fixKnownRoadSpacingAddressNormV5_(s);
  s = finalCleanAddressNormV5_(s, legalIndex);

  var parsed = applyLegalDongPrefixAddressNormV5_(s, legalIndex);
  var normalized = finalCleanAddressNormV5_(parsed.value, legalIndex);

  var quality = classifyAddressNormV5_(normalized, raw, parsed.legalFixed, picked.memoCut, parsed.ambiguous, legalIndex);

  return {
    normalized: normalized,
    needCheck: quality.needCheck,
    status: quality.status,
    note: quality.note,
    legalFixed: parsed.legalFixed,
    memoCut: picked.memoCut,
    ambiguous: parsed.ambiguous
  };
}


/**
 * 법정동코드 색인 생성
 */
function buildLegalDongIndexAddressNormV5_(ss) {
  var index = {
    available: false,
    existCount: 0,

    currentProvinceSet: {},
    provinceAliasToCurrent: {},

    fullNameSet: {},
    suffixToFull: {},
    provinceSuffixToFull: {},
    compressedSuffixToFull: {},
    provinceCompressedSuffixToFull: {},

    adminTokenToFullPrefix: {},
    knownLegalTokenSet: {},
    knownLegalTokensSorted: []
  };

  var forceMap = ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT || {};
  Object.keys(forceMap).forEach(function (k) {
    index.provinceAliasToCurrent[k] = forceMap[k];
  });

  var sheet = ss.getSheetByName(ADDRESS_NORM_V5_CONFIG.LEGAL_DONG_SHEET_NAME);
  if (!sheet) return index;

  var values = sheet.getDataRange().getDisplayValues();
  if (!values || values.length < 2) return index;

  var headers = values[0].map(function (v) {
    return String(v || '').trim();
  });

  var nameIdx = headers.indexOf('법정동명');
  var statusIdx = headers.indexOf('폐지여부');

  if (nameIdx < 0) {
    throw new Error('법정동코드 시트에 "법정동명" 헤더가 없습니다.');
  }

  for (var r = 1; r < values.length; r++) {
    var status = statusIdx >= 0 ? String(values[r][statusIdx] || '').trim() : '존재';
    if (status !== '존재') continue;

    var legalName = cleanLegalNameAddressNormV5_(values[r][nameIdx]);
    legalName = canonicalizeLegalNameProvinceAddressNormV5_(legalName);
    if (!legalName) continue;

    var tokens = legalName.split(' ').filter(function (v) { return !!v; });
    if (!tokens.length) continue;

    index.existCount++;
    index.fullNameSet[legalName] = true;

    var province = tokens[0];
    index.currentProvinceSet[province] = true;
    index.provinceAliasToCurrent[province] = province;

    for (var i = 1; i < tokens.length; i++) {
      var tok = tokens[i];
      if (isLegalAdminTokenAddressNormV5_(tok)) {
        index.knownLegalTokenSet[tok] = true;
      }

      if (/(시|군|구|읍|면)$/.test(tok)) {
        addUniqueMapAddressNormV5_(index.adminTokenToFullPrefix, tok, tokens.slice(0, i + 1).join(' '));
      }
    }

    for (var start = 1; start < tokens.length; start++) {
      var suffix = tokens.slice(start).join(' ');
      addUniqueMapAddressNormV5_(index.suffixToFull, suffix, legalName);
      addUniqueMapAddressNormV5_(index.provinceSuffixToFull, province + '|' + suffix, legalName);
    }

    addCompressedLegalSuffixesAddressNormV5_(index, tokens, legalName);
  }

  index.knownLegalTokensSorted = Object.keys(index.knownLegalTokenSet).sort(function (a, b) {
    return b.length - a.length;
  });

  index.available = true;
  return index;
}


function cleanLegalNameAddressNormV5_(value) {
  var s = String(value || '');

  s = s.replace(/\([^)]*\)/g, ' ');
  s = s.replace(/^부산직할시\b/, '부산광역시');
  s = s.replace(/^대구직할시\b/, '대구광역시');
  s = s.replace(/^인천직할시\b/, '인천광역시');
  s = s.replace(/^광주직할시\b/, '광주광역시');
  s = s.replace(/^대전직할시\b/, '대전광역시');

  return cleanSpacesAddressNormV5_(s);
}


function canonicalizeLegalNameProvinceAddressNormV5_(legalName) {
  var s = cleanSpacesAddressNormV5_(legalName);
  if (!s) return '';

  var tokens = s.split(' ').filter(function (v) { return !!v; });
  if (!tokens.length) return '';

  var forceMap = ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT || {};
  if (forceMap[tokens[0]]) {
    tokens[0] = forceMap[tokens[0]];
  }

  return cleanSpacesAddressNormV5_(tokens.join(' '));
}


function isLegalAdminTokenAddressNormV5_(token) {
  return /(.+시|.+군|.+구|.+읍|.+면|.+동|.+리|.+가)$/.test(String(token || ''));
}


function addCompressedLegalSuffixesAddressNormV5_(index, tokens, legalName) {
  if (!tokens || tokens.length < 4) return;

  var province = tokens[0];

  if (/시$/.test(tokens[1]) && /구$/.test(tokens[2])) {
    for (var start = 3; start < tokens.length; start++) {
      var compressed = [tokens[1]].concat(tokens.slice(start)).join(' ');
      addUniqueMapAddressNormV5_(index.compressedSuffixToFull, compressed, legalName);
      addUniqueMapAddressNormV5_(index.provinceCompressedSuffixToFull, province + '|' + compressed, legalName);
    }
  }
}


function addUniqueMapAddressNormV5_(map, key, value) {
  key = cleanSpacesAddressNormV5_(key);
  value = cleanSpacesAddressNormV5_(value);

  if (!key || !value) return;

  if (!map[key]) {
    map[key] = value;
    return;
  }

  if (map[key] !== value) {
    map[key] = '__DUPLICATED__';
  }
}


function pickPrimaryAddressTextAddressNormV5_(value) {
  var original = String(value || '');
  var s = original;

  s = s.replace(/\u00a0/g, ' ');
  s = s.replace(/[\r\n]+/g, ' / ');
  s = cleanSpacesAddressNormV5_(s);

  var before = s;

  var target = s.match(/대상처\s*주소\s*[:：]\s*/);
  if (target) {
    s = s.slice(target.index + target[0].length);
  }

  s = removeLeadingLabelAddressNormV5_(s);

  var parts = s
    .split(/\s*(?:\/|／|\||;|；)\s*/g)
    .map(removeLeadingLabelAddressNormV5_)
    .map(stripTailMemoAddressNormV5_)
    .map(cleanSpacesAddressNormV5_)
    .filter(function (v) { return !!v; });

  var picked = '';

  for (var i = 0; i < parts.length; i++) {
    if (looksLikeAddressTextAddressNormV5_(parts[i])) {
      picked = parts[i];
      break;
    }
  }

  if (!picked && parts.length) picked = parts[0];

  picked = stripTailMemoAddressNormV5_(picked);
  picked = removeLeadingLabelAddressNormV5_(picked);
  picked = cleanSpacesAddressNormV5_(picked);

  return {
    value: picked,
    memoCut: cleanSpacesAddressNormV5_(before) !== picked
  };
}


function removeLeadingLabelAddressNormV5_(value) {
  var s = String(value || '').trim();

  s = s.replace(/^(신청인\s*주소|대상처\s*주소|고객사\s*주소|사업장\s*주소|소재지|주소)\s*[:：]\s*/g, '');

  s = s.replace(/^[^:：]{1,30}\s*[:：]\s*(?=[가-힣]{2,}(특별시|광역시|특별자치시|특별자치도|도|시|군|구)|서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주)/g, '');

  return cleanSpacesAddressNormV5_(s);
}


function stripTailMemoAddressNormV5_(value) {
  var s = String(value || '');

  s = s.replace(/<\s*\d{4}\s*년도\s*대상\s*>/g, ' ');
  s = s.replace(/\([^)]*\)/g, ' ');
  s = s.replace(/\[[^\]]*\]/g, ' ');

  s = s.replace(/\s*(주소\s*확인\s*필요|주소확인필요|확인\s*필요|확인필요).*$/g, '');
  s = s.replace(/\s+별도.*$/g, '');

  s = s.replace(/\s+외\s*\d+\s*필지.*$/g, '');
  s = s.replace(/\s+외\s*필지.*$/g, '');

  s = s.replace(/\s*\([^)]*$/g, ' ');
  s = s.replace(/[.,，、]+$/g, '');

  return cleanSpacesAddressNormV5_(s);
}


function normalizeBasicAddressTextAddressNormV5_(value) {
  var s = String(value || '');

  s = s.replace(/\u00a0/g, ' ');
  s = s.replace(/[\r\n\t]+/g, ' ');
  s = s.replace(/<[^>]*>/g, ' ');

  s = s.replace(/^\s*\(?\d{5}\)?\s+/g, '');
  s = s.replace(/^\s*대한민국\s+/g, '');

  s = s.replace(/[，,]+/g, ' ');
  s = s.replace(/\s*-\s*/g, '-');
  s = s.replace(/(\d+)\s*번지/g, '$1');

  s = s.replace(/^인천광역시시\b/, '인천광역시');
  s = s.replace(/^대전광역시시\b/, '대전광역시');
  s = s.replace(/^부산광역시시\b/, '부산광역시');
  s = s.replace(/^서울시\b/, '서울');
  s = s.replace(/^세종특별시\b/, '세종특별자치시');

  return cleanSpacesAddressNormV5_(s);
}


function normalizeProvinceTokenOnlyAddressNormV5_(value, legalIndex) {
  var s = cleanSpacesAddressNormV5_(value);
  if (!s) return '';

  s = insertSpaceAfterKnownProvinceIfStickyAddressNormV5_(s, legalIndex);

  var tokens = s.split(' ').filter(function (v) { return !!v; });
  if (!tokens.length) return '';

  var first = tokens[0];
  var forceMap = ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT || {};

  if (forceMap[first]) {
    tokens[0] = forceMap[first];
    return cleanSpacesAddressNormV5_(tokens.join(' '));
  }

  if (legalIndex.currentProvinceSet[first]) {
    return cleanSpacesAddressNormV5_(tokens.join(' '));
  }

  if (legalIndex.provinceAliasToCurrent[first]) {
    tokens[0] = legalIndex.provinceAliasToCurrent[first];
    return cleanSpacesAddressNormV5_(tokens.join(' '));
  }

  return cleanSpacesAddressNormV5_(tokens.join(' '));
}


function insertSpaceAfterKnownProvinceIfStickyAddressNormV5_(value, legalIndex) {
  var s = String(value || '').trim();
  if (!s) return '';

  var aliases = Object.keys(ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT || {})
    .concat(Object.keys(legalIndex.currentProvinceSet || {}))
    .filter(function (v, idx, arr) {
      return v && arr.indexOf(v) === idx;
    })
    .sort(function (a, b) {
      return b.length - a.length;
    });

  for (var i = 0; i < aliases.length; i++) {
    var a = aliases[i];

    if (s === a) {
      return ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT[a] || legalIndex.provinceAliasToCurrent[a] || a;
    }

    if (s.indexOf(a) === 0) {
      var next = s.charAt(a.length);
      if (next && next !== ' ') {
        var mapped = ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT[a] || legalIndex.provinceAliasToCurrent[a] || a;
        return cleanSpacesAddressNormV5_(mapped + ' ' + s.slice(a.length));
      }
      break;
    }
  }

  return cleanSpacesAddressNormV5_(s);
}


function splitStickyAddressTokensAddressNormV5_(value, legalIndex) {
  var tokens = cleanSpacesAddressNormV5_(value).split(' ').filter(function (v) {
    return !!v;
  });

  var out = [];
  var known = legalIndex.knownLegalTokensSorted || [];

  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    var split = splitOneStickyTokenByKnownLegalTokenAddressNormV5_(t, known);
    for (var j = 0; j < split.length; j++) {
      out.push(split[j]);
    }
  }

  return cleanSpacesAddressNormV5_(out.join(' '));
}


function splitOneStickyTokenByKnownLegalTokenAddressNormV5_(token, knownLegalTokensSorted) {
  var t = String(token || '');
  if (!t) return [];

  for (var i = 0; i < knownLegalTokensSorted.length; i++) {
    var admin = knownLegalTokensSorted[i];

    if (!admin || admin.length < 2) continue;
    if (t === admin) return [t];

    if (t.indexOf(admin) === 0 && t.length > admin.length) {
      var rest = t.slice(admin.length);

      if (looksLikeRestAfterAdminTokenAddressNormV5_(rest)) {
        return [admin, rest];
      }
    }
  }

  return [t];
}


function looksLikeRestAfterAdminTokenAddressNormV5_(rest) {
  var s = String(rest || '');
  if (!s) return false;

  if (/^[가-힣0-9]+(?:대로|로|길|번길)/.test(s)) return true;
  if (/^[가-힣0-9]+(?:읍|면|동|리|가)/.test(s)) return true;
  if (/^\d/.test(s)) return true;

  return false;
}


function fixKnownRoadSpacingAddressNormV5_(value) {
  var s = cleanSpacesAddressNormV5_(value);

  var replacements = [
    ['달구 벌대로', '달구벌대로'],
    ['압구 정로', '압구정로'],
    ['국책연구 원', '국책연구원'],
    ['연구 단지로', '연구단지로'],
    ['유구 마곡사로', '유구마곡사로'],
    ['옥구 천동로', '옥구천동로'],
    ['가구 단지길', '가구단지길'],
    ['산단구 평길', '산단구평길']
  ];

  for (var i = 0; i < replacements.length; i++) {
    s = s.split(replacements[i][0]).join(replacements[i][1]);
  }

  s = s.replace(/(\d+)\s+번길/g, '$1번길');
  s = s.replace(/(\d+)\s+산단/g, '$1산단');

  return cleanSpacesAddressNormV5_(s);
}


function applyLegalDongPrefixAddressNormV5_(value, legalIndex) {
  var s = cleanSpacesAddressNormV5_(value);
  var tokens = s.split(' ').filter(function (v) { return !!v; });

  if (!tokens.length) {
    return { value: '', legalFixed: false, ambiguous: false };
  }

  if (legalIndex.currentProvinceSet[tokens[0]]) {
    var province = tokens[0];
    var rest = tokens.slice(1);
    var direct = matchProvinceSuffixAddressNormV5_(province, rest, legalIndex);

    if (direct.matched) {
      return {
        value: direct.fullTokens.concat(direct.remaining).join(' '),
        legalFixed: true,
        ambiguous: false
      };
    }

    if (direct.ambiguous) {
      return { value: tokens.join(' '), legalFixed: false, ambiguous: true };
    }

    return { value: tokens.join(' '), legalFixed: false, ambiguous: false };
  }

  var global = matchGlobalSuffixAddressNormV5_(tokens, legalIndex);
  if (global.matched) {
    return {
      value: global.fullTokens.concat(global.remaining).join(' '),
      legalFixed: true,
      ambiguous: false
    };
  }

  if (global.ambiguous) {
    return { value: tokens.join(' '), legalFixed: false, ambiguous: true };
  }

  var first = tokens[0];
  var prefix = legalIndex.adminTokenToFullPrefix[first];
  if (prefix && prefix !== '__DUPLICATED__') {
    return {
      value: prefix.split(' ').concat(tokens.slice(1)).join(' '),
      legalFixed: true,
      ambiguous: false
    };
  }

  if (prefix === '__DUPLICATED__') {
    return { value: tokens.join(' '), legalFixed: false, ambiguous: true };
  }

  return { value: tokens.join(' '), legalFixed: false, ambiguous: false };
}


function matchProvinceSuffixAddressNormV5_(province, restTokens, legalIndex) {
  var maxLen = Math.min(6, restTokens.length);

  for (var len = maxLen; len >= 1; len--) {
    var suffix = restTokens.slice(0, len).join(' ');
    var key = province + '|' + suffix;

    var full = legalIndex.provinceSuffixToFull[key];
    if (full && full !== '__DUPLICATED__') {
      return { matched: true, fullTokens: full.split(' '), remaining: restTokens.slice(len), ambiguous: false };
    }
    if (full === '__DUPLICATED__') {
      return { matched: false, ambiguous: true };
    }

    var compressed = legalIndex.provinceCompressedSuffixToFull[key];
    if (compressed && compressed !== '__DUPLICATED__') {
      return { matched: true, fullTokens: compressed.split(' '), remaining: restTokens.slice(len), ambiguous: false };
    }
    if (compressed === '__DUPLICATED__') {
      return { matched: false, ambiguous: true };
    }
  }

  return { matched: false, ambiguous: false };
}


function matchGlobalSuffixAddressNormV5_(tokens, legalIndex) {
  var maxLen = Math.min(6, tokens.length);

  for (var len = maxLen; len >= 2; len--) {
    var suffix = tokens.slice(0, len).join(' ');

    var full = legalIndex.suffixToFull[suffix];
    if (full && full !== '__DUPLICATED__') {
      return { matched: true, fullTokens: full.split(' '), remaining: tokens.slice(len), ambiguous: false };
    }
    if (full === '__DUPLICATED__') {
      return { matched: false, ambiguous: true };
    }

    var compressed = legalIndex.compressedSuffixToFull[suffix];
    if (compressed && compressed !== '__DUPLICATED__') {
      return { matched: true, fullTokens: compressed.split(' '), remaining: tokens.slice(len), ambiguous: false };
    }
    if (compressed === '__DUPLICATED__') {
      return { matched: false, ambiguous: true };
    }
  }

  return { matched: false, ambiguous: false };
}


function finalCleanAddressNormV5_(value, legalIndex) {
  var s = cleanSpacesAddressNormV5_(value);

  s = fixKnownRoadSpacingAddressNormV5_(s);
  s = stripTailMemoAddressNormV5_(s);
  s = removeRepeatedProvinceTokensAddressNormV5_(s, legalIndex);
  s = cutAtSecondTopLevelTokenAddressNormV5_(s, legalIndex);

  s = s.replace(/\s*-\s*/g, '-');
  s = s.replace(/[.,，、]+$/g, '');
  s = cleanSpacesAddressNormV5_(s);

  return s;
}


function removeRepeatedProvinceTokensAddressNormV5_(value, legalIndex) {
  var tokens = cleanSpacesAddressNormV5_(value).split(' ').filter(function (v) { return !!v; });
  if (!tokens.length) return '';

  var first = tokens[0];
  if (!legalIndex.currentProvinceSet[first]) return tokens.join(' ');

  var removeSecondTokenMap = {
    '서울특별시': { '특별시': true, '시': true },
    '부산광역시': { '광역시': true, '시': true },
    '대구광역시': { '광역시': true, '시': true },
    '인천광역시': { '광역시': true, '시': true },
    '대전광역시': { '광역시': true, '시': true },
    '울산광역시': { '광역시': true, '시': true },
    '세종특별자치시': { '특별자치시': true, '특별시': true, '시': true },
    '경기도': { '도': true },
    '강원특별자치도': { '특별자치도': true, '도': true },
    '충청북도': { '도': true },
    '충청남도': { '도': true },
    '전북특별자치도': { '특별자치도': true, '도': true },
    '경상북도': { '도': true },
    '경상남도': { '도': true },
    '제주특별자치도': { '특별자치도': true, '도': true },
    '전남광주통합특별시': { '통합특별시': true, '특별시': true, '시': true }
  };

  if (tokens.length >= 2 && removeSecondTokenMap[first] && removeSecondTokenMap[first][tokens[1]]) {
    tokens.splice(1, 1);
  }

  return cleanSpacesAddressNormV5_(tokens.join(' '));
}


function cutAtSecondTopLevelTokenAddressNormV5_(value, legalIndex) {
  var tokens = cleanSpacesAddressNormV5_(value).split(' ').filter(function (v) { return !!v; });
  if (tokens.length < 3) return tokens.join(' ');

  var firstTopIdx = -1;
  var secondTopIdx = -1;

  for (var i = 0; i < tokens.length; i++) {
    var t = tokens[i];
    if (legalIndex.currentProvinceSet[t] || ADDRESS_NORM_V5_CONFIG.FORCE_PROVINCE_ALIAS_TO_CURRENT[t]) {
      if (firstTopIdx < 0) {
        firstTopIdx = i;
      } else {
        secondTopIdx = i;
        break;
      }
    }
  }

  if (firstTopIdx >= 0 && secondTopIdx > firstTopIdx) {
    return tokens.slice(0, secondTopIdx).join(' ');
  }

  return tokens.join(' ');
}


function classifyAddressNormV5_(normalized, raw, legalFixed, memoCut, ambiguous, legalIndex) {
  var s = cleanSpacesAddressNormV5_(normalized);
  var rawText = String(raw || '');
  var notes = [];
  var needCheck = false;

  if (!s) {
    return { needCheck: true, status: '공란', note: '주소 공란 또는 주소 후보 없음' };
  }

  if (memoCut) notes.push('메모/복수주소 절단');
  if (legalFixed) notes.push('법정동코드 보정');

  if (ambiguous) {
    needCheck = true;
    notes.push('법정동 후보 중복/애매함');
  }

  if (/ㅇ{2,}|몰라|모름|미상|불명|테스트|주소\s*확인\s*필요|확인\s*필요/.test(rawText + ' ' + s)) {
    needCheck = true;
    notes.push('placeholder/확인필요 문구 포함');
  }

  var tokens = s.split(' ').filter(function (v) { return !!v; });
  var province = tokens[0] || '';

  if (!legalIndex.currentProvinceSet[province]) {
    needCheck = true;
    notes.push('시/도 없음 또는 미인식');
  }

  if (tokens.length <= 1) {
    needCheck = true;
    notes.push('시/도까지만 있음');
  }

  if (tokens.length === 2 && /(시|군|구)$/.test(tokens[1])) {
    needCheck = true;
    notes.push('시군구까지만 있음');
  }

  if (province && province !== '세종특별자치시' && tokens.length >= 2) {
    if (!/(시|군|구)$/.test(tokens[1])) {
      needCheck = true;
      notes.push('시군구 누락 의심');
    }
  }

  if (!/\d/.test(s)) {
    needCheck = true;
    notes.push('상세 지번/건물번호 없음');
  }

  if (/^[가-힣A-Za-z0-9\s]+$/.test(s) && !/(로|길|동|리|읍|면|가|\d)/.test(s)) {
    needCheck = true;
    notes.push('주소 형식 불충분');
  }

  if (s.length < 8) {
    needCheck = true;
    notes.push('주소 길이 과소');
  }

  var status;
  if (ambiguous) {
    status = '애매함';
  } else if (needCheck) {
    status = '확인필요';
  } else {
    status = '정상';
  }

  if (!notes.length) notes.push('정상');

  return {
    needCheck: needCheck,
    status: status,
    note: notes.join(' / ')
  };
}


function looksLikeAddressTextAddressNormV5_(value) {
  var s = String(value || '').trim();
  if (!s) return false;

  if (/(서울|부산|대구|인천|광주|대전|울산|세종|경기|강원|충북|충남|전북|전남|경북|경남|제주|전남광주통합특별시)/.test(s)) return true;
  if (/[가-힣]+(시|군|구)\s*[가-힣0-9]+(읍|면|동|가|리)/.test(s)) return true;
  if (/[가-힣0-9]+(대로|로|길|번길)\s*\d/.test(s)) return true;

  return false;
}


function findHeaderColAddressNormV5_(sheet, headerName) {
  var lastCol = sheet.getLastColumn();
  var headers = sheet
    .getRange(ADDRESS_NORM_V5_CONFIG.HEADER_ROW, 1, 1, lastCol)
    .getDisplayValues()[0]
    .map(function (v) {
      return String(v || '').trim();
    });

  var idx = headers.indexOf(headerName);
  return idx >= 0 ? idx + 1 : -1;
}


function ensureHeaderColAddressNormV5_(sheet, headerName) {
  var found = findHeaderColAddressNormV5_(sheet, headerName);
  if (found > 0) return found;

  var newCol = sheet.getLastColumn() + 1;
  sheet.getRange(ADDRESS_NORM_V5_CONFIG.HEADER_ROW, newCol).setValue(headerName);
  return newCol;
}


function cleanSpacesAddressNormV5_(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, ' ')
    .trim();
}