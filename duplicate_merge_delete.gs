/**
 * 마스터시트(신규) 중복 확정행 병합/삭제
 *
 * 1) 메모에 "책임님 확인 완료. 중복으로 삭제예정.(방수원)"이 있는 행만 대상
 * 2) 보존 행과 다른 지정 필드만 보존 행의 메모에 기록
 * 3) 실행 직전 마스터시트 전체를 복제하여 원상복구 가능한 백업 생성
 * 4) 메모 반영 후 삭제행을 아래쪽부터 삭제
 * 5) 주소 자동매핑 금지: 검토된 고객번호 명시 매핑만 실행
 * 6) 중요 상태/수행사 충돌 및 별도 시설 의심 건은 보류
 * 7) 재실행 시 같은 삭제 고객번호의 병합 블록은 중복 추가하지 않음
 */

var DUPLICATE_CLEANUP_CONFIG = {
  VERSION: '20260711-R2',
  SHEET_NAME: '마스터시트(신규)',
  LOG_SHEET_NAME: '_중복삭제로그',
  BACKUP_SHEET_PREFIX: '_중복삭제전체백업_',

  DELETE_MARKER_REGEX: /책임님\s*확인\s*완료\.?\s*중복으로\s*삭제예정\.?\s*\(방수원\)/,

  COMPLETION_MEMO: '[기타메모] 26.07.11. 12:50 중복 삭제 및 데이터 병합 완료 (방수원)',

  MERGE_FIELDS: [
    '고객사 상세 주소',
    '고객사 담당자',
    '대표전화번호',
    '직통번호',
    '담당자 이메일 주소',
    '연면적',
    '관리등급',
    '건물 유형',
    '할인률(%)',
    '계약시작일',
    '계약종료일',
    '계약단위',
    '관리자 선임 여부',
    '유지점검',
    '성능점검',
    '최종 견적가',
    '부가세'
  ],

  /**
   * 현재 검토본에서 확정된 '삭제 고객번호 -> 보존 고객번호' 매핑입니다.
   * 주소만 같다는 이유로 자동 선택하지 않습니다.
   * 새로운 삭제표시 행이 생기면 이 표에 명시적으로 추가하기 전까지 보류됩니다.
   */
  KEEP_BY_DELETE_CUSTOMER_NO: {
    '16': '2503',
    '48': '49',
    '98': '584',
    '106': '805',
    '141': '590',
    '532': '2315',
    '534': '2206',
    '538': '1225',
    '576': '253',
    '683': '2302',
    '903': '904',
    '1055': '1054',
    '1074': '992',
    '1120': '2715',
    '1349': '1348',
    '1427': '1383',
    '1629': '1627',
    '1727': '751',
    '1734': '909',
    '1814': '1716',
    '1849': '1072',
    '1892': '1907',
    '1931': '450',
    '1938': '297',
    '1969': '1494',
    '2091': '666',
    '2115': '1605',
    '2296': '255',
    '2364': '2347',
    '2396': '1204',
    '2468': '2873',
    '2532': '564',
    '2590': '1181',
    '2684': '2064',
    '2691': '1797',
    '2714': '1130',
    '2720': '659',
    '2727': '917',
    '2805': '128',
    '3035': '3001',
    '3339': '3325',
    '3378': '3789',
    '3409': '3539',
    '3418': '3166',
    '3423': '3771',
    '3542': '3813',
    '3631': '3340',
    '3742': '3419',
    '3754': '3141',
    '3809': '2962'
  },

  /**
   * 검증 과정에서 별도 업체/건물/공장일 가능성이 확인된 항목입니다.
   * 삭제표시가 있어도 자동 병합·삭제하지 않고 로그에 보류로 남깁니다.
   */
  HOLD_BY_DELETE_CUSTOMER_NO: {
    '90':  { keepCustomerNo: '584',  reason: '시립마포 실버케어센터와 서울복지타운은 회사명·담당자·전화·이메일이 달라 별도 기관 가능성' },
    '197': { keepCustomerNo: '38',   reason: '대전테크노파크와 바이오벤처타운·GMP는 메모상 별도 GMP 건물 가능성' },
    '654': { keepCustomerNo: '206',  reason: '마포농수산물시장 전체 데이터와 1동 데이터의 관계가 불명확함' },
    '1431':{ keepCustomerNo: '1384', reason: '코츠 1·2공장 통합 데이터와 1공장 단독 데이터는 별도 시설 가능성' },
    '2309':{ keepCustomerNo: '2162', reason: '황화빌딩과 미림타워는 주소 외 동일 근거가 부족함' },
    '2352':{ keepCustomerNo: '1829', reason: '두선산업 안산공장과 케이레이저는 서로 다른 업체임' },
    '2358':{ keepCustomerNo: '272',  reason: '국제약품중앙연구소와 국제약품은 별도 연구시설 가능성' },
    '2389':{ keepCustomerNo: '1418', reason: '히든힐스 서밋사이트와 프라임사이트는 서로 다른 사이트명임' },
    '2502':{ keepCustomerNo: '2503', reason: '호룡 6공장과 호룡 본사는 별도 주소·별도 공장임' },
    '3095':{ keepCustomerNo: '3094', reason: '연구단지로 162와 161은 번지와 연면적이 다른 별도 건물 가능성' }
  },

  /**
   * 아래 상태가 삭제행에 있는데 보존행 상태와 다르면 자동 삭제하지 않습니다.
   * 진행상태는 요청된 메모 병합 필드가 아니므로 중요한 상태 유실을 차단합니다.
   */
  PROTECTED_DELETE_STATUSES: [
    '계약완료',
    '영업팀 수주 성공',
    '계약서 취합 완료',
    '발주완료'
  ],

  /** 수행사가 양쪽 모두 입력되어 있고 서로 다르면 자동 삭제하지 않습니다. */
  BLOCK_ON_VENDOR_CONFLICT: true
};

/**
 * 실제 변경 없이 매핑 결과와 병합 예정 필드만 _중복삭제로그 시트에 작성합니다.
 */
function previewConfirmedDuplicateCleanup() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = getDuplicateCleanupTargetSheet_(ss);
  var context = loadDuplicateCleanupContext_(sheet);
  var plan = buildDuplicateCleanupPlan_(context);

  writeDuplicateCleanupLog_(ss, plan, '미리보기');
  notifyDuplicateCleanup_(
    ss,
    '중복 정리 미리보기',
    '실행 예정 ' + plan.executable.length + '건 / 보류 ' + plan.unresolved.length + '건\n' +
      DUPLICATE_CLEANUP_CONFIG.LOG_SHEET_NAME + ' 시트를 확인해 주세요.'
  );
}

/**
 * 실제 병합 및 행 삭제를 실행합니다.
 */
function runConfirmedDuplicateCleanup() {
  var lock = LockService.getDocumentLock();
  if (!lock.tryLock(5000)) {
    throw new Error('다른 작업이 마스터시트를 사용 중입니다. 잠시 후 다시 실행해 주세요.');
  }

  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();
    var sheet = getDuplicateCleanupTargetSheet_(ss);
    var context = loadDuplicateCleanupContext_(sheet);
    var plan = buildDuplicateCleanupPlan_(context);

    if (plan.executable.length === 0) {
      writeDuplicateCleanupLog_(ss, plan, '실행');
      notifyDuplicateCleanup_(
        ss,
        '중복 정리',
        '삭제 가능한 대상이 없습니다. 보류 건은 ' + DUPLICATE_CLEANUP_CONFIG.LOG_SHEET_NAME + ' 시트를 확인해 주세요.'
      );
      return;
    }

    var backupSheetName = createDuplicateCleanupBackup_(ss, sheet);
    var memoUpdates = buildMergedMemoUpdates_(context, plan.executable);

    writeMergedMemos_(sheet, context, memoUpdates);
    SpreadsheetApp.flush();

    var deleteRowNumbers = plan.executable.map(function(item) {
      return item.deleteRecord.rowNumber;
    });
    deleteRowsBottomUp_(sheet, deleteRowNumbers);

    plan.executable.forEach(function(item) {
      item.status = '삭제 완료';
    });
    writeDuplicateCleanupLog_(ss, plan, '실행');

    notifyDuplicateCleanup_(
      ss,
      '중복 정리 완료',
      plan.executable.length + '개 행을 병합 후 삭제했습니다.\n' +
        '백업: ' + backupSheetName + '\n' +
        '보류: ' + plan.unresolved.length + '건'
    );
  } finally {
    lock.releaseLock();
  }
}

function getDuplicateCleanupTargetSheet_(ss) {
  var sheet = ss.getSheetByName(DUPLICATE_CLEANUP_CONFIG.SHEET_NAME);
  if (sheet) return sheet;

  sheet = ss.getActiveSheet();
  if (!sheet) {
    throw new Error('대상 시트를 찾을 수 없습니다: ' + DUPLICATE_CLEANUP_CONFIG.SHEET_NAME);
  }
  return sheet;
}

function loadDuplicateCleanupContext_(sheet) {
  var lastRow = sheet.getLastRow();
  var lastColumn = sheet.getLastColumn();

  if (lastRow < 2 || lastColumn < 1) {
    throw new Error('대상 시트에 데이터가 없습니다.');
  }

  var range = sheet.getRange(1, 1, lastRow, lastColumn);
  var displayValues = range.getDisplayValues();
  var rawValues = range.getValues();
  var headerRowIndex = findDuplicateCleanupHeaderRowIndex_(displayValues);
  var headers = displayValues[headerRowIndex];
  var headerMap = buildDuplicateCleanupHeaderMap_(headers);

  var requiredHeaders = [
    '고객번호',
    '회사명',
    '메모',
    '고객사 상세 주소',
    '현재 영업 진행 상황',
    '수행사'
  ].concat(DUPLICATE_CLEANUP_CONFIG.MERGE_FIELDS);

  requiredHeaders.forEach(function(headerName) {
    if (headerMap[normalizeDuplicateCleanupHeader_(headerName)] === undefined) {
      throw new Error('필수 헤더를 찾을 수 없습니다: ' + headerName);
    }
  });

  var records = [];
  for (var i = headerRowIndex + 1; i < displayValues.length; i++) {
    var displayRow = displayValues[i];
    var rawRow = rawValues[i];

    if (isDuplicateCleanupEmptyRow_(displayRow)) continue;

    var customerNo = getDuplicateCleanupCell_(displayRow, headerMap, '고객번호').trim();
    var companyName = getDuplicateCleanupCell_(displayRow, headerMap, '회사명').trim();
    var memo = String(getDuplicateCleanupRawCell_(rawRow, headerMap, '메모') || '');
    var detailAddress = getDuplicateCleanupCell_(displayRow, headerMap, '고객사 상세 주소');
    var normalizedAddressValue = '';

    if (headerMap[normalizeDuplicateCleanupHeader_('주소정규화')] !== undefined) {
      normalizedAddressValue = getDuplicateCleanupCell_(displayRow, headerMap, '주소정규화');
    }

    records.push({
      rowNumber: i + 1,
      displayRow: displayRow,
      rawRow: rawRow,
      customerNo: customerNo,
      companyName: companyName,
      memo: memo,
      isDeleteMarked: isDuplicateCleanupDeleteMarked_(memo),
      addressKey: normalizeDuplicateCleanupAddress_(normalizedAddressValue || detailAddress),
      detailAddress: detailAddress,
      salesStatus: getDuplicateCleanupCell_(displayRow, headerMap, '현재 영업 진행 상황').trim(),
      vendor: getDuplicateCleanupCell_(displayRow, headerMap, '수행사').trim()
    });
  }

  return {
    sheet: sheet,
    displayValues: displayValues,
    rawValues: rawValues,
    headerRowIndex: headerRowIndex,
    headers: headers,
    headerMap: headerMap,
    records: records
  };
}

function buildDuplicateCleanupPlan_(context) {
  var keepRecords = context.records.filter(function(record) {
    return !record.isDeleteMarked;
  });
  var deleteRecords = context.records.filter(function(record) {
    return record.isDeleteMarked;
  });

  var keepByCustomerNo = {};
  keepRecords.forEach(function(record) {
    if (!record.customerNo) return;
    if (!keepByCustomerNo[record.customerNo]) keepByCustomerNo[record.customerNo] = [];
    keepByCustomerNo[record.customerNo].push(record);
  });

  var executable = [];
  var unresolved = [];

  deleteRecords.forEach(function(deleteRecord) {
    var holdRule = DUPLICATE_CLEANUP_CONFIG.HOLD_BY_DELETE_CUSTOMER_NO[deleteRecord.customerNo];
    if (holdRule) {
      var holdKeepRecord = getUniqueDuplicateCleanupKeepRecord_(keepByCustomerNo, holdRule.keepCustomerNo);
      unresolved.push(makeDuplicateCleanupUnresolved_(
        deleteRecord,
        '검증 보류: ' + holdRule.reason,
        holdKeepRecord,
        '명시적 보류'
      ));
      return;
    }

    var keepCustomerNo = DUPLICATE_CLEANUP_CONFIG.KEEP_BY_DELETE_CUSTOMER_NO[deleteRecord.customerNo];
    if (!keepCustomerNo) {
      unresolved.push(makeDuplicateCleanupUnresolved_(
        deleteRecord,
        '명시 매핑이 없는 삭제표시 행입니다. KEEP_BY_DELETE_CUSTOMER_NO에 보존 고객번호를 추가해야 합니다.',
        null,
        '매핑 없음'
      ));
      return;
    }

    var keepCandidates = keepByCustomerNo[keepCustomerNo] || [];
    if (keepCandidates.length === 0) {
      unresolved.push(makeDuplicateCleanupUnresolved_(
        deleteRecord,
        '보존 고객번호 ' + keepCustomerNo + '를 찾을 수 없습니다.',
        null,
        '명시 매핑'
      ));
      return;
    }
    if (keepCandidates.length > 1) {
      unresolved.push(makeDuplicateCleanupUnresolved_(
        deleteRecord,
        '보존 고객번호 ' + keepCustomerNo + '가 ' + keepCandidates.length + '개 행에 존재합니다.',
        null,
        '명시 매핑'
      ));
      return;
    }

    var keepRecord = keepCandidates[0];
    if (!keepRecord || keepRecord.isDeleteMarked) {
      unresolved.push(makeDuplicateCleanupUnresolved_(
        deleteRecord,
        '보존 대상으로 지정된 행도 삭제표시 상태입니다.',
        keepRecord,
        '명시 매핑'
      ));
      return;
    }

    var protectedReason = getDuplicateCleanupProtectedConflictReason_(deleteRecord, keepRecord);
    if (protectedReason) {
      unresolved.push(makeDuplicateCleanupUnresolved_(
        deleteRecord,
        protectedReason,
        keepRecord,
        '보호 규칙'
      ));
      return;
    }

    var mergeData = collectDuplicateCleanupMergeData_(context, deleteRecord, keepRecord);
    executable.push({
      status: '실행 예정',
      reason: '검토 완료 고객번호 명시 매핑',
      deleteRecord: deleteRecord,
      keepRecord: keepRecord,
      differences: mergeData.differences,
      remainingSourceMemo: mergeData.remainingSourceMemo
    });
  });

  return {
    executable: executable,
    unresolved: unresolved
  };
}

function getUniqueDuplicateCleanupKeepRecord_(keepByCustomerNo, customerNo) {
  if (!customerNo) return null;
  var candidates = keepByCustomerNo[String(customerNo)] || [];
  return candidates.length === 1 ? candidates[0] : null;
}

function getDuplicateCleanupProtectedConflictReason_(deleteRecord, keepRecord) {
  var deleteStatus = normalizeDuplicateCleanupText_(deleteRecord.salesStatus);
  var keepStatus = normalizeDuplicateCleanupText_(keepRecord.salesStatus);
  var protectedStatuses = {};

  DUPLICATE_CLEANUP_CONFIG.PROTECTED_DELETE_STATUSES.forEach(function(status) {
    protectedStatuses[normalizeDuplicateCleanupText_(status)] = true;
  });

  if (protectedStatuses[deleteStatus] && deleteStatus !== keepStatus) {
    return '중요 진행상태 충돌: 삭제행=' + (deleteRecord.salesStatus || '(공란)') +
      ', 보존행=' + (keepRecord.salesStatus || '(공란)');
  }

  if (
    DUPLICATE_CLEANUP_CONFIG.BLOCK_ON_VENDOR_CONFLICT &&
    deleteRecord.vendor &&
    keepRecord.vendor &&
    normalizeDuplicateCleanupText_(deleteRecord.vendor) !== normalizeDuplicateCleanupText_(keepRecord.vendor)
  ) {
    return '수행사 충돌: 삭제행=' + deleteRecord.vendor + ', 보존행=' + keepRecord.vendor;
  }

  return '';
}

function collectDuplicateCleanupMergeData_(context, deleteRecord, keepRecord) {
  var differences = [];

  DUPLICATE_CLEANUP_CONFIG.MERGE_FIELDS.forEach(function(fieldName) {
    var columnIndex = context.headerMap[normalizeDuplicateCleanupHeader_(fieldName)];
    var deleteValue = String(deleteRecord.displayRow[columnIndex] || '').trim();
    var keepValue = String(keepRecord.displayRow[columnIndex] || '').trim();

    if (!isDuplicateCleanupUsefulValue_(deleteValue)) return;
    if (isDuplicateCleanupValueAlreadyPreserved_(fieldName, deleteValue, keepValue)) return;

    differences.push({
      fieldName: fieldName,
      deleteValue: deleteValue,
      keepValue: keepValue
    });
  });

  var remainingSourceMemo = extractDuplicateCleanupRemainingMemo_(deleteRecord.memo);
  if (remainingSourceMemo && containsDuplicateCleanupNormalizedText_(keepRecord.memo, remainingSourceMemo)) {
    remainingSourceMemo = '';
  }

  return {
    differences: differences,
    remainingSourceMemo: remainingSourceMemo
  };
}

function buildMergedMemoUpdates_(context, executableItems) {
  var grouped = {};

  executableItems.forEach(function(item) {
    var key = String(item.keepRecord.rowNumber);
    if (!grouped[key]) {
      grouped[key] = {
        keepRecord: item.keepRecord,
        items: []
      };
    }
    grouped[key].items.push(item);
  });

  var memoUpdates = {};

  Object.keys(grouped).forEach(function(key) {
    var group = grouped[key];
    var memo = String(group.keepRecord.memo || '');
    memo = removeDuplicateCleanupCompletionMemo_(memo);

    group.items.forEach(function(item) {
      var sourceToken = '삭제 고객번호 ' + item.deleteRecord.customerNo;
      if (memo.indexOf(sourceToken) !== -1) return;

      var blockLines = [];
      if (item.differences.length > 0 || item.remainingSourceMemo) {
        blockLines.push(
          '[중복 삭제행 데이터 | 삭제 고객번호 ' + item.deleteRecord.customerNo +
          ' | ' + item.deleteRecord.companyName + ']'
        );

        item.differences.forEach(function(diff) {
          blockLines.push('- ' + diff.fieldName + ': ' + diff.deleteValue);
        });

        if (item.remainingSourceMemo) {
          blockLines.push('- 기존 메모:');
          item.remainingSourceMemo.split(/\r?\n/).forEach(function(line) {
            if (line.trim()) blockLines.push('  ' + line.trim());
          });
        }
      }

      if (blockLines.length > 0) {
        memo = appendDuplicateCleanupMemoBlock_(memo, blockLines.join('\n'));
      }
    });

    memo = appendDuplicateCleanupMemoBlock_(memo, DUPLICATE_CLEANUP_CONFIG.COMPLETION_MEMO);
    memoUpdates[key] = memo;
  });

  return memoUpdates;
}

function writeMergedMemos_(sheet, context, memoUpdates) {
  var memoColumnIndex = context.headerMap[normalizeDuplicateCleanupHeader_('메모')] + 1;

  Object.keys(memoUpdates).forEach(function(rowNumberText) {
    var rowNumber = Number(rowNumberText);
    sheet.getRange(rowNumber, memoColumnIndex).setValue(memoUpdates[rowNumberText]);
  });
}

function createDuplicateCleanupBackup_(ss, sourceSheet) {
  var timeZone = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  var suffix = Utilities.formatDate(new Date(), timeZone, 'yyMMdd_HHmmss');
  var baseName = DUPLICATE_CLEANUP_CONFIG.BACKUP_SHEET_PREFIX + suffix;
  var sheetName = makeUniqueDuplicateCleanupSheetName_(ss, baseName);

  // 행 일부를 값으로 옮기는 방식이 아니라 시트 전체를 복제합니다.
  // 수식, 서식, 메모, 데이터 유효성, 숨김 상태 등 원본 복구에 필요한 요소를 최대한 보존합니다.
  var backupSheet = sourceSheet.copyTo(ss);
  backupSheet.setName(sheetName);
  backupSheet.setFrozenRows(sourceSheet.getFrozenRows());
  backupSheet.setFrozenColumns(sourceSheet.getFrozenColumns());

  return sheetName;
}

function deleteRowsBottomUp_(sheet, rowNumbers) {
  var uniqueRows = {};
  rowNumbers.forEach(function(rowNumber) {
    uniqueRows[rowNumber] = true;
  });

  var sorted = Object.keys(uniqueRows).map(Number).sort(function(a, b) {
    return b - a;
  });

  if (sorted.length === 0) return;

  var groupStart = sorted[0];
  var groupEnd = sorted[0];

  for (var i = 1; i <= sorted.length; i++) {
    var current = i < sorted.length ? sorted[i] : null;

    if (current !== null && current === groupEnd - 1) {
      groupEnd = current;
      continue;
    }

    sheet.deleteRows(groupEnd, groupStart - groupEnd + 1);

    if (current !== null) {
      groupStart = current;
      groupEnd = current;
    }
  }
}

function writeDuplicateCleanupLog_(ss, plan, mode) {
  var sheet = ss.getSheetByName(DUPLICATE_CLEANUP_CONFIG.LOG_SHEET_NAME);
  if (!sheet) sheet = ss.insertSheet(DUPLICATE_CLEANUP_CONFIG.LOG_SHEET_NAME);

  sheet.clearContents();

  var timeZone = ss.getSpreadsheetTimeZone() || 'Asia/Seoul';
  var nowText = Utilities.formatDate(new Date(), timeZone, 'yyyy-MM-dd HH:mm:ss');
  var rows = [[
    '기록일시',
    '모드',
    '상태',
    '매핑근거',
    '삭제원본행',
    '삭제고객번호',
    '삭제회사명',
    '삭제주소',
    '보존원본행',
    '보존고객번호',
    '보존회사명',
    '삭제진행상태',
    '보존진행상태',
    '삭제수행사',
    '보존수행사',
    '병합대상',
    '비고'
  ]];

  plan.executable.forEach(function(item) {
    var mergeSummary = item.differences.map(function(diff) {
      return diff.fieldName + '=' + diff.deleteValue;
    });
    if (item.remainingSourceMemo) mergeSummary.push('기존 메모 있음');

    rows.push([
      nowText,
      mode,
      item.status,
      item.reason,
      item.deleteRecord.rowNumber,
      item.deleteRecord.customerNo,
      item.deleteRecord.companyName,
      item.deleteRecord.detailAddress,
      item.keepRecord.rowNumber,
      item.keepRecord.customerNo,
      item.keepRecord.companyName,
      item.deleteRecord.salesStatus,
      item.keepRecord.salesStatus,
      item.deleteRecord.vendor,
      item.keepRecord.vendor,
      mergeSummary.join(' / '),
      ''
    ]);
  });

  plan.unresolved.forEach(function(item) {
    rows.push([
      nowText,
      mode,
      '보류',
      item.mappingReason,
      item.deleteRecord.rowNumber,
      item.deleteRecord.customerNo,
      item.deleteRecord.companyName,
      item.deleteRecord.detailAddress,
      item.keepRecord ? item.keepRecord.rowNumber : '',
      item.keepRecord ? item.keepRecord.customerNo : '',
      item.keepRecord ? item.keepRecord.companyName : '',
      item.deleteRecord.salesStatus,
      item.keepRecord ? item.keepRecord.salesStatus : '',
      item.deleteRecord.vendor,
      item.keepRecord ? item.keepRecord.vendor : '',
      '',
      item.reason
    ]);
  });

  sheet.getRange(1, 1, rows.length, rows[0].length).setValues(rows);
  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, rows[0].length).setFontWeight('bold');
  sheet.autoResizeColumns(1, rows[0].length);
}

function makeDuplicateCleanupUnresolved_(deleteRecord, reason, keepRecord, mappingReason) {
  return {
    deleteRecord: deleteRecord,
    keepRecord: keepRecord || null,
    mappingReason: mappingReason || '',
    reason: reason
  };
}

function findDuplicateCleanupHeaderRowIndex_(displayValues) {
  var maxRows = Math.min(displayValues.length, 10);

  for (var i = 0; i < maxRows; i++) {
    var normalizedHeaders = displayValues[i].map(normalizeDuplicateCleanupHeader_);
    if (
      normalizedHeaders.indexOf(normalizeDuplicateCleanupHeader_('고객번호')) !== -1 &&
      normalizedHeaders.indexOf(normalizeDuplicateCleanupHeader_('회사명')) !== -1 &&
      normalizedHeaders.indexOf(normalizeDuplicateCleanupHeader_('메모')) !== -1 &&
      normalizedHeaders.indexOf(normalizeDuplicateCleanupHeader_('고객사 상세 주소')) !== -1
    ) {
      return i;
    }
  }

  throw new Error('헤더 행을 찾을 수 없습니다. 고객번호/회사명/메모/고객사 상세 주소 헤더를 확인해 주세요.');
}

function buildDuplicateCleanupHeaderMap_(headers) {
  var map = {};
  headers.forEach(function(header, index) {
    var normalized = normalizeDuplicateCleanupHeader_(header);
    if (normalized && map[normalized] === undefined) map[normalized] = index;
  });
  return map;
}

function normalizeDuplicateCleanupHeader_(value) {
  return String(value || '')
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, '')
    .trim();
}

function getDuplicateCleanupCell_(row, headerMap, headerName) {
  var index = headerMap[normalizeDuplicateCleanupHeader_(headerName)];
  return index === undefined ? '' : String(row[index] || '');
}

function getDuplicateCleanupRawCell_(row, headerMap, headerName) {
  var index = headerMap[normalizeDuplicateCleanupHeader_(headerName)];
  return index === undefined ? '' : row[index];
}

function isDuplicateCleanupDeleteMarked_(memo) {
  return DUPLICATE_CLEANUP_CONFIG.DELETE_MARKER_REGEX.test(String(memo || ''));
}

function extractDuplicateCleanupRemainingMemo_(memo) {
  return String(memo || '')
    .split(/\r?\n/)
    .filter(function(line) {
      return !isDuplicateCleanupDeleteMarked_(line);
    })
    .join('\n')
    .trim();
}

function isDuplicateCleanupEmptyRow_(row) {
  for (var i = 0; i < row.length; i++) {
    if (String(row[i] || '').trim() !== '') return false;
  }
  return true;
}

function isDuplicateCleanupUsefulValue_(value) {
  var text = String(value || '').trim();
  if (!text) return false;
  return !/^#(N\/A|REF!|VALUE!|DIV\/0!|NAME\?|NUM!|ERROR!)$/i.test(text);
}

function isDuplicateCleanupValueAlreadyPreserved_(fieldName, deleteValue, keepValue) {
  if (!isDuplicateCleanupUsefulValue_(keepValue)) return false;

  var normalizedField = normalizeDuplicateCleanupHeader_(fieldName);

  if (normalizedField === normalizeDuplicateCleanupHeader_('고객사 상세 주소')) {
    return normalizeDuplicateCleanupAddress_(deleteValue) === normalizeDuplicateCleanupAddress_(keepValue);
  }

  if (
    normalizedField === normalizeDuplicateCleanupHeader_('대표전화번호') ||
    normalizedField === normalizeDuplicateCleanupHeader_('직통번호')
  ) {
    return isDuplicateCleanupSetContained_(
      extractDuplicateCleanupPhoneSet_(deleteValue),
      extractDuplicateCleanupPhoneSet_(keepValue)
    );
  }

  if (normalizedField === normalizeDuplicateCleanupHeader_('담당자 이메일 주소')) {
    return isDuplicateCleanupSetContained_(
      extractDuplicateCleanupEmailSet_(deleteValue),
      extractDuplicateCleanupEmailSet_(keepValue)
    );
  }

  if (
    normalizedField === normalizeDuplicateCleanupHeader_('연면적') ||
    normalizedField === normalizeDuplicateCleanupHeader_('할인률(%)') ||
    normalizedField === normalizeDuplicateCleanupHeader_('최종 견적가')
  ) {
    var deleteNumber = parseDuplicateCleanupNumber_(deleteValue);
    var keepNumber = parseDuplicateCleanupNumber_(keepValue);
    if (deleteNumber !== null && keepNumber !== null) {
      return Math.abs(deleteNumber - keepNumber) <= Math.max(0.01, Math.abs(deleteNumber) * 0.000001);
    }
  }

  if (
    normalizedField === normalizeDuplicateCleanupHeader_('계약시작일') ||
    normalizedField === normalizeDuplicateCleanupHeader_('계약종료일')
  ) {
    return normalizeDuplicateCleanupDate_(deleteValue) === normalizeDuplicateCleanupDate_(keepValue);
  }

  if (
    normalizedField === normalizeDuplicateCleanupHeader_('유지점검') ||
    normalizedField === normalizeDuplicateCleanupHeader_('성능점검')
  ) {
    return normalizeDuplicateCleanupCount_(deleteValue) === normalizeDuplicateCleanupCount_(keepValue);
  }

  if (normalizedField === normalizeDuplicateCleanupHeader_('관리자 선임 여부')) {
    return normalizeDuplicateCleanupAppointment_(deleteValue) === normalizeDuplicateCleanupAppointment_(keepValue);
  }

  if (normalizedField === normalizeDuplicateCleanupHeader_('부가세')) {
    return normalizeDuplicateCleanupVat_(deleteValue) === normalizeDuplicateCleanupVat_(keepValue);
  }

  return normalizeDuplicateCleanupText_(deleteValue) === normalizeDuplicateCleanupText_(keepValue);
}

function normalizeDuplicateCleanupAddress_(value) {
  var text = String(value || '').toLowerCase();
  text = text.replace(/\([^)]*\)/g, ' ');

  var aliases = [
    [/서울특별시|서울시/g, '서울'],
    [/부산광역시|부산시/g, '부산'],
    [/대구광역시|대구시/g, '대구'],
    [/인천광역시|인천시/g, '인천'],
    [/광주광역시|광주시/g, '광주'],
    [/대전광역시|대전시/g, '대전'],
    [/울산광역시|울산시/g, '울산'],
    [/세종특별자치시|세종시/g, '세종'],
    [/경기도|경기/g, '경기'],
    [/강원특별자치도|강원도|강원/g, '강원'],
    [/충청북도|충북/g, '충북'],
    [/충청남도|충남/g, '충남'],
    [/전북특별자치도|전라북도|전북/g, '전북'],
    [/전라남도|전남/g, '전남'],
    [/경상북도|경북/g, '경북'],
    [/경상남도|경남/g, '경남'],
    [/제주특별자치도|제주도/g, '제주']
  ];

  aliases.forEach(function(pair) {
    text = text.replace(pair[0], pair[1]);
  });

  return text.replace(/[^0-9a-z가-힣]/g, '');
}

function normalizeDuplicateCleanupText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/\u00a0/g, ' ')
    .replace(/\s+/g, '')
    .replace(/[.,·ㆍ/\\()\[\]{}_-]/g, '')
    .trim();
}

function extractDuplicateCleanupPhoneSet_(value) {
  var matches = String(value || '').match(/0\d{1,2}[\s\-)]*\d{3,4}[\s\-]*\d{4}/g) || [];
  var set = {};

  matches.forEach(function(match) {
    var digits = match.replace(/\D/g, '');
    if (digits.length >= 9) set[digits] = true;
  });

  return Object.keys(set).sort();
}

function extractDuplicateCleanupEmailSet_(value) {
  var matches = String(value || '').toLowerCase().match(/[a-z0-9._%+-]+@[a-z0-9.-]+\.[a-z]{2,}/g) || [];
  var set = {};
  matches.forEach(function(email) {
    set[email] = true;
  });
  return Object.keys(set).sort();
}

function isDuplicateCleanupSetContained_(deleteItems, keepItems) {
  if (deleteItems.length === 0 || keepItems.length === 0) return false;
  var keepSet = {};
  keepItems.forEach(function(item) {
    keepSet[item] = true;
  });
  return deleteItems.every(function(item) {
    return !!keepSet[item];
  });
}

function parseDuplicateCleanupNumber_(value) {
  var text = String(value || '').replace(/[^0-9.\-]/g, '');
  if (!text || text === '-' || text === '.') return null;
  var number = Number(text);
  return isNaN(number) ? null : number;
}

function normalizeDuplicateCleanupDate_(value) {
  var digits = String(value || '').replace(/\D/g, '');
  if (digits.length === 8) return digits;
  if (digits.length === 6) return '20' + digits;
  return normalizeDuplicateCleanupText_(value);
}

function normalizeDuplicateCleanupCount_(value) {
  var match = String(value || '').match(/-?\d+(?:\.\d+)?/);
  return match ? String(Number(match[0])) : normalizeDuplicateCleanupText_(value);
}

function normalizeDuplicateCleanupAppointment_(value) {
  var text = normalizeDuplicateCleanupText_(value);
  if (!text) return '';
  if (text.indexOf('미선임') !== -1 || text === 'x' || text === 'n' || text === 'no' || text === '아니오') {
    return '미선임';
  }
  if (text.indexOf('선임') !== -1 || text === 'o' || text === 'y' || text === 'yes' || text === '예') {
    return '선임';
  }
  return text;
}

function normalizeDuplicateCleanupVat_(value) {
  var text = normalizeDuplicateCleanupText_(value);
  if (text.indexOf('별도') !== -1) return '별도';
  if (text.indexOf('포함') !== -1) return '포함';
  if (text.indexOf('면세') !== -1) return '면세';
  return text;
}

function containsDuplicateCleanupNormalizedText_(containerText, targetText) {
  var container = normalizeDuplicateCleanupText_(containerText);
  var target = normalizeDuplicateCleanupText_(targetText);
  return !!target && container.indexOf(target) !== -1;
}

function removeDuplicateCleanupCompletionMemo_(memo) {
  var completion = DUPLICATE_CLEANUP_CONFIG.COMPLETION_MEMO;
  return String(memo || '')
    .split(/\r?\n/)
    .filter(function(line) {
      return line.trim() !== completion;
    })
    .join('\n')
    .trim();
}

function appendDuplicateCleanupMemoBlock_(memo, block) {
  var base = String(memo || '').trim();
  var addition = String(block || '').trim();
  if (!addition) return base;
  return base ? base + '\n' + addition : addition;
}

function makeUniqueDuplicateCleanupSheetName_(ss, baseName) {
  var name = baseName.substring(0, 90);
  var counter = 1;

  while (ss.getSheetByName(name)) {
    name = (baseName.substring(0, 85) + '_' + counter).substring(0, 90);
    counter++;
  }

  return name;
}

function notifyDuplicateCleanup_(ss, title, message) {
  ss.toast(message, title, 10);
  try {
    SpreadsheetApp.getUi().alert(title, message, SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    Logger.log(title + ': ' + message);
  }
}
