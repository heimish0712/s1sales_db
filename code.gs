/**
 * 수기견적+메일자동화 / 하이웍스 API 파일 직접첨부 메일 발송 자동화 v23
 * - 컨테이너 바인드 Apps Script 기준
 * - 버튼/그림에는 mailAutoSend 또는 메일자동발송 함수를 연결하세요.
 * - 하이웍스 첨부 발송부는 실제 성공한 multipart/form-data 구조 기준입니다.
 * - v12: 수신/담당자/문의사항 템플릿 및 체크박스별 요청문구 조합 기능 추가
 * - v13: 영업담당자별 명함 및 약관/안내문 체크 시 안내문 이미지 본문 삽입
 * - v14: DOCX 템플릿이 Office 파일일 때 Google Docs 변환 복사 후 편집하도록 수정
 * - v15: 접수번호 등록 LockService 제거. Sheets API append 방식 시도
 * - v16: Sheets API 의존 제거. SpreadsheetApp.appendRow + UUID 마커로 동시 실행 시에도 각자 추가된 행을 찾아 접수번호 확정
 * - v17: 메일 제목 최초발송/재발송 케이스별 자동 분기
 * - v18: 선임신고서 placeholder 치환 강화, DOCX 출력색 검정화, 견적서/용역신청서 이미지 중복 삭제, XLSX #REF 방지
 * - v19: 선임신고서 신고인 정보 우선순위 변경(사업자등록증 법인명/주소 우선), 대표자/사업자번호는 없으면 공란, 날짜 공란 유지
 * - v21: 선임신고서용 사업자등록증상 정보는 마스터 원행 전체와 병합하여 사용. AX 계약 당사자/AY 법인주소 우선 적용
 * - v22: 선임신고서 {{주소}}와 {{대상주소}} 분리. {{주소}}는 사업자등록증상 주소 우선, {{대상주소}}는 I열 고객사 상세 주소 우선 적용
 * - v23: 선임신고서 템플릿 복사 시 Drive API 호출을 최소화. 먼저 DriveApp+DocumentApp 네이티브 복사 시도, 실패 시에만 Drive API 변환 복사 + 지수 백오프 재시도
 * - v24: 견적서/비교견적서 PDF 출력 여백과 스케일을 기준 샘플처럼 조정. A4 한 장에 중앙 정렬, 0.60in 여백, fit-to-page 방식 적용
 * - v25: Google Sheets PDF export에서 오른쪽/아래쪽 외곽선이 잘리는 문제 방지. 출력 범위를 한 칸씩 여유 있게 확장
 * - v26: 견적서/용역신청서 도장·로고 삽입 전 기존 이미지 삭제 강화. anchor가 범위 밖이어도 이미지 사각형이 배치 범위와 겹치면 삭제
 * - v37: 용역신청서 하단 로고/도장 레이어 순서 수정. 로고를 먼저 삽입하고 도장을 나중에 삽입하여 도장이 위에 오도록 처리
 * - v33: AC 안내문/법령요약 4개 파일 일괄 첨부, AD 수행사정보 수행사별 파일 첨부 적용
 * - v36: 선임신고서/위임장 내 모든 고객 상호 표시는 사업자등록증상 법인명 우선, 없으면 기존 회사명 fallback으로 통일
 * - v39: 용역신청서 신청인 정보 보정. 사업자등록증상 법인명/주소 우선, 없으면 회사명/고객사 상세 주소 fallback. 대표자/사업자번호/업태/종목은 값 없으면 공란 유지
 * - v40: 마스터 헤더명 기준 용역신청서 특약사항 반영 및 샘플보고서 폴더 파일 첨부 추가
 * - v41: 샘플보고서 등 대용량 첨부파일을 Apps Script 내부 용량 제한으로 사전 차단하지 않고 하이웍스 API에 그대로 전달하도록 수정
 * - v42: 하이웍스 전송을 UrlFetchApp 네이티브 multipart payload 방식으로 변경. blob.getBytes() 수동 배열 생성으로 인한 메모리 초과 가능성 완화
 * - v43: sendMail API 20MB 제한 대응. 20MB 초과 Drive 폴더 첨부파일은 첨부 제외 후 본문 Drive 링크로 자동 삽입
 * - v44: 샘플보고서와 수행사정보가 동시에 체크된 경우에만 두 자료를 Google Drive 링크로 전달. 샘플보고서 단독/수행사정보 단독은 기존 첨부 방식 유지
 * - v45: 하이웍스 첨부파일명 한글 깨짐 방지. 수동 multipart + UTF-8 filename/filename* 헤더로 전송
 * - v51: 메일자동발송 팝업에서 참조 master@s1samsung.com을 사용자가 삭제 가능하도록 수정
 * - v50: 네이버 수신 차단 대응. 첨부파일 multipart 헤더에서 raw 한글 filename 제거, ASCII fallback filename + UTF-8 filename* 방식으로 변경
 * - v52: 첨부파일명 인코딩 방식 변경. filename에는 RFC2047 encoded-word(=?UTF-8?B?...?=), filename*에는 UTF-8 RFC5987 값을 함께 전달
 * - v54: 하이웍스가 filename-star/RFC2047을 디코딩하지 않아 파일명이 깨지는 문제 대응. filename-star 제거, raw UTF-8 filename 단독 전송 방식으로 변경
 * - v42: 31MB 이상 파일에서 Apps Script 메모리 초과를 피하도록 하이웍스 multipart 전송을 UrlFetchApp 네이티브 payload 방식으로 변경
 * - v56: 팝업 행 선택 우선순위 수정. 현재 활성 선택행이 유효하면 기억된 이전 행보다 우선 사용
 * - v63: 대용량 Google Drive 링크 카드에서 파일명 클릭 시 미리보기가 아니라 직접 다운로드되도록 수정
 * - v62: 대용량 링크 카드 상단의 전체 다운로드 기능 제거. 개별 파일 미리보기/다운로드만 유지
 * - v60: 대용량 링크 자료의 상단 전체 다운로드는 ZIP 생성 없이 웹앱 다운로드 실행 페이지에서 개별 파일을 순차 다운로드하도록 수정
 * - v65: AG 수행사 용역표준계약서 체크 시 수행사별 표준계약서 DOCX를 추가 첨부
 * - v66: 선임신고서/위임장 DOCX 템플릿을 마스터시트 헤더명 기반으로 치환하고, {{A}} or {{B}} fallback 표현을 A 우선/B 대체값으로 처리
 * - v67: 나에게 발송(TEST) 시 참조 전체 제외. 발신자는 영업담당자, 수신자는 입력한 회사계정만 사용
 * - v84: 파일 확인/수정 후 발송은 수정본을 바로 export하여 생성기 복사/도장삽입 재실행을 생략하고, reviewSelectedDefs ReferenceError 수정
 * - v75: 비교견적(1) O4, 비교견적(2) G5 수행사명을 인식해 각 시트 T22에 수행사 도장 삽입
 * - v77: 비교견적(2) 도장 위치를 T29로 변경하고, 파일 확인/수정 시 비교견적서를 편집용 Google Sheets로 생성 후 최종 발송은 PDF로 export. 비교견적별 지정 범위 수식 보존
 * - v78: 팝업에서 비교견적서(2)를 이번 발송에서 제외할 수 있도록 하고, 파일 확인/수정/최종 발송 모두 선택된 비교견적서만 생성·첨부
 * - v76: onOpen 메뉴에 도장/로고 캐시 예열 메뉴 추가
 * - v73: 파일 확인/수정 생성속도 개선. 수정용 Sheets 생성은 Sheets API copyTo 우선 사용, 선임신고서 DOCX 템플릿은 Google Docs 변환본을 캐시하여 재사용
 * - v68: TEST 모드는 서버 초입에서 수동 수신/참조 payload도 폐기하여 입력 회사계정 1건만 수신 처리
 * - v69: [파일 확인/수정] 사전 생성 폴더 기능 추가. 수정용 Sheet/Docs 파일을 Drive 폴더에 만들고, 발송 시 최종 수정본을 첨부로 export
 * - v70: DriveApp.getRootFolder() 접근 거부 회피. 수정용 폴더/임시파일을 원본 생성기 또는 마스터 파일의 부모 폴더 기준으로 생성
 * - v81: 통합 수정용 시트 생성에서 Google Sheets REST API 의존 제거. Sheets API 비활성 프로젝트에서도 SpreadsheetApp 기본 경로로 생성
 * - v82: 파일 확인/수정 UI를 문서별 개별 파일 방식으로 복귀. 비편집 자료는 HTML 링크모음 대신 Drive shortcut으로 표시하고, shortcut 캐시를 적용
 * - v83: 파일 확인/수정은 수정 가능한 문서(견적서/용역신청서/선임신고서 및 위임장/비교견적서)만 생성하도록 단순화
 * - v85: CUSTOMER 발송 성공 시 고객사 공유드라이브 폴더에 [발송] 첨부파일 누적 저장 + 발송파일로그 중앙 기록 추가
 * - v85: 고객사 폴더별 발송이력 Google Sheet 일일 반영 트리거/수동 실행 함수 추가
 * - v87: 자료발송 고객 데이터 캐시 오염 방지. 도장/로고 사전삽입 생성기 캐시의 생성대상 행을 초기화하고, 발송/파일확인 전 실제 고객명 재계산 확인을 추가
 * - v88: 견적서/용역신청서이 자동견적요청 3행을 직접 참조하는 양식까지 방어. 임시 생성기 복사본의 자동견적요청 3행과 생성대상 3행을 현재 고객으로 강제 동기화하고, 선택 시트에 현재 고객명이 반영되지 않으면 발송/파일확인을 중단
 * - v89: 최종 파일생성기 구조 반영. 원본 첫 시트(자동견적요청)는 접수 로그로만 append하고, 발송 계산은 생성대상!3행만 보도록 보정. 임시 생성기에서 자동견적요청!3행 직접참조 수식을 생성대상!3행 참조로 자동 치환
 * - v90: 계약시작일/계약종료일이 비어 있으면 계약시작일 26.07.01 기본값과 계약단위 기준 종료일을 자동 계산하여 생성대상 R/S와 견적서 점검기간에 강제 반영
 * - v91: 메일자동화 작업공간/캐시/수정용 파일 저장 위치를 공유드라이브 폴더로 고정할 수 있도록 MAILAUTO_WORKSPACE_FOLDER_ID 설정 추가
 * - v93: 포탈/마스터 데이터 정합성 반영. 자동견적요청/생성대상 이식 시 날짜·횟수·개월·면적·금액·할인율을 숫자/날짜 데이터와 표시서식으로 유지
 * - v96: 선임신고서/위임장 고객 상호는 사업자등록증상 법인명 우선 유지하되, 위임장 '위임명 {{회사명}} 정보통신유지관리자 선임 신고'의 회사명만 내부 회사명으로 출력
 */

const CONFIG = Object.freeze({
  MASTER_SPREADSHEET_ID: '1ADDJMrej-EJBw4QHkq17xefWxQw5hZ_NgQLf2BuCD8Q',
  GENERATOR_SPREADSHEET_ID: '1K6-JtUb5qjN1nfTE0rRtpzdMMfTGudOL7DGGWAPPi34',
  APPOINTMENT_DOC_TEMPLATE_ID: '1QoNRhmD1u2ttqW7j4TVR1vMZuDKq-Dnh',

  SHEETS: {
    MASTER: '마스터시트(신규)',
    REQUEST_LOG: '자동견적요청',
    TARGET: '생성대상',
    SALES_REP: '영업담당자 정보',
    MAIL_LOG: '메일발송로그',
    QUOTE: '견적서',
    SERVICE_APP: '용역신청서',
    COMPARE_1: '비교견적(1)',
    COMPARE_2: '비교견적(2)'
  },

  ROWS: {
    MASTER_HEADER: 2,
    MASTER_DATA_START: 3,
    REQUEST_HEADER: 2,
    REQUEST_DATA_START: 3,
    TARGET_HEADER: 2,
    TARGET_DATA_ROW: 3,
    SALES_REP_HEADER: 1,
    SALES_REP_DATA_START: 2
  },

  MASTER_STATUS_HEADERS: {
    LAST_SENT: '마지막발송',
    SEND_COUNT: '발송횟수',
    STATUS: '발송상태',
    SENT_AT: '발송일시',
    MESSAGE: '처리메시지',
    TEMP_FOLDER: '저장폴더',
    REQUEST_NO: '접수번호'
  },

  CHECK_HEADERS: [
    '견적서요청',
    '용역신청서요청',
    '선임신고서 및 위임장 요청',
    '약관, 안내문, 법령요약',
    '수행사정보',
    '비교견적서',
    '샘플보고서',
    '수행사 용역표준계약서'
  ],

  // 수행사 추가/삭제는 여기만 수정하면 됩니다.
  VENDORS: {
    '일신': {
      displayName: '일신',
      stampFileId: '1rM5cX4P_rfXOi-fCmUZ9RQhuSjtcI8l5',
      quoteLogoFileId: '1EXgiYccFgl-vkTmN8iWJkYCr3iAEvjq3',
      serviceLogoFileId: '1m4_OTvIDVWf1tRTGzLao0My_IaXqJXvj',
      contractorInfoZipFileId: '1GCiWPuyBJOltB-lMNoVT3bsFZkHgWVcS',
      serviceStandardContractFileId: '1W75TR5fg0ferxAfmBL7KVlKIxtNmqU_Y'
    },
    'KJ': {
      displayName: 'KJ',
      stampFileId: '1tko254QVQQ8fPeIL_dz8hoCF_lx2NObY',
      quoteLogoFileId: '1U2gHWzSxwioybEs_jseCP5uzRRmaRguE',
      serviceLogoFileId: '1eO0IaCTCMj0s-VT6B02Y7UZYdmN6TJoC',
      contractorInfoZipFileId: '1j51-p6fJxdHdtaJyfEDc2VSp3-c0P5Mp',
      serviceStandardContractFileId: '1o2xqRS8FXHIBPRmF_T2cK2CDIi4A1XWK'
    },
    '디엠': {
      displayName: '디엠',
      stampFileId: '1PNQcu9j_CQQvM81q0bHuzGnc-_3r-Gvw',
      quoteLogoFileId: '190JwlssVvNYMjMQjLtkL6qqFnts5vJ-R',
      serviceLogoFileId: '188ljO4ttk-CBurDlz2yjcd0AofIXkrP8',
      contractorInfoZipFileId: '',
      serviceStandardContractFileId: ''
    },
    '삼구': {
      displayName: '삼구',
      stampFileId: '1UYEagKEPe6ILAUO2xSeGVqgGRPC_OYOa',
      quoteLogoFileId: '1mJMyCSjWV_V3cucplSejf8z4eEUvtdfJ',
      serviceLogoFileId: '1GnP6MmwwCHhA_JYOds5vc87etbVrlIxE',
      contractorInfoZipFileId: '10zmg-z_lPi-lyiQGA5Jc3mNzInP2y5zo',
      serviceStandardContractFileId: '1cpxtRXs9j3ezJ84A6hZproZTvI9g9KHz'
    },
    '케이제이': {
      displayName: 'KJ',
      stampFileId: '1tko254QVQQ8fPeIL_dz8hoCF_lx2NObY',
      quoteLogoFileId: '1U2gHWzSxwioybEs_jseCP5uzRRmaRguE',
      serviceLogoFileId: '1eO0IaCTCMj0s-VT6B02Y7UZYdmN6TJoC',
      contractorInfoZipFileId: '1j51-p6fJxdHdtaJyfEDc2VSp3-c0P5Mp',
      serviceStandardContractFileId: '1o2xqRS8FXHIBPRmF_T2cK2CDIi4A1XWK'
    }
  },

  // 나중에 파일 추가 시 여기에 FileDefinition만 추가하면 됩니다.
  FILE_DEFINITIONS: [
    {
      key: 'quote',
      checkHeader: '견적서요청',
      label: '견적서',
      type: 'sheet_pdf',
      sheetName: '견적서',
      // 중요: 실제 내용 범위가 A1:BJ34여도, PDF export가 마지막 행/열 외곽선을 잘라먹는 경우가 있어
      // 오른쪽/아래쪽으로 빈 1칸을 포함해 A1:BK35로 출력합니다.
      exportRange: 'A1:BK36',
      filename: '{company}_견적서.pdf'
    },
    {
      key: 'serviceApplication',
      checkHeader: '용역신청서요청',
      label: '용역신청서',
      type: 'sheet_xlsx_values',
      sheetName: '용역신청서',
      filename: '{company}_용역신청서.xlsx'
    },
    {
      key: 'appointmentDoc',
      checkHeader: '선임신고서 및 위임장 요청',
      label: '선임신고서 및 위임장',
      type: 'docx_template',
      templateFileId: '1QoNRhmD1u2ttqW7j4TVR1vMZuDKq-Dnh',
      filename: '{company}_선임신고서_및_위임장.docx'
    },
    {
      key: 'termsGuide',
      // AC열: 안내문/법령 등 체크 시 아래 4개 파일을 압축 없이 각각 첨부합니다.
      checkHeader: '약관, 안내문, 법령요약',
      checkHeaders: ['약관, 안내문, 법령요약', '안내문, 법령요약', '안내문, 법령', '약관/안내문/법령요약'],
      label: '안내문/법령요약',
      type: 'static_files',
      fileIds: [
        '1n9q_8IfiXz85NUrOXE-U8jQUnd5bn2Wu',
        '1rcVKYvCNRZhd4txH_qe7zSGvCWLRztXm',
        '1qhKGTgSrMThcxWd9etfteJ4UJUZG1jar',
        '1Fbb2rzq8uIWi9y6AHnZV4gifmw8ipIBe'
      ]
    },
    {
      key: 'contractorInfo',
      // AD열: 수행사정보 체크 시 수행사 값에 맞는 자료 1개를 첨부합니다.
      checkHeader: '수행사정보',
      checkHeaders: ['수행사정보', '수행사 정보'],
      label: '수행사정보',
      type: 'vendor_zip',
      filename: '{vendor}_수행사정보.zip'
    },
    {
      key: 'serviceStandardContract',
      // AG열: 수행사 용역표준계약서 체크 시 수행사 값에 맞는 DOCX 템플릿을 고객별로 치환한 뒤 첨부합니다.
      // 템플릿 안의 {마스터시트 헤더명}은 해당 행 값으로 치환됩니다.
      // {A} or {B} 형태는 A 값이 있으면 A, 없으면 B 값으로 치환됩니다.
      checkHeader: '수행사 용역표준계약서',
      checkHeaders: ['수행사 용역표준계약서', '수행사 용역표준계약', '용역표준계약서', '표준계약서'],
      label: '수행사 용역표준계약서',
      type: 'vendor_docx_template',
      vendorFileIdField: 'serviceStandardContractFileId',
      filename: '{company}_{vendor}_용역표준계약서.docx'
    },
    {
      key: 'compareQuote',
      checkHeader: '비교견적서',
      label: '비교견적서',
      type: 'multi_sheet_pdf',
      sheets: [
        // 비교견적도 동일하게 실제 A1:AE40보다 한 칸 여유 있게 출력합니다.
        { sheetName: '비교견적(1)', exportRange: 'A1:AF33', filename: '{company}_비교견적서_1.pdf' },
        { sheetName: '비교견적(2)', exportRange: 'A1:AF33', filename: '{company}_비교견적서_2.pdf' }
      ]
    },
    {
      key: 'sampleReport',
      // 마스터시트(신규)의 헤더명에 '샘플보고서'가 포함된 체크박스가 TRUE이면
      // 아래 Google Drive 폴더 안의 파일을 전부 첨부합니다.
      checkHeader: '샘플보고서',
      checkHeaderContains: true,
      label: '샘플보고서',
      type: 'drive_folder_files',
      folderId: '1ByGWM73gJe0iT3JAcUS6Yet8lKdcvSgi'
    }
  ],

  IMAGE_PLACEMENTS: [
    { sheetName: '견적서', role: 'stampFileId', rangeA1: 'BD4:BJ5', hAlign: 'center', vAlign: 'middle' },
    { sheetName: '견적서', role: 'quoteLogoFileId', rangeA1: 'AG32:BC34', hAlign: 'right', vAlign: 'middle' },

    // 중요 v37:
    // Google Sheets 이미지는 나중에 삽입한 이미지가 위 레이어로 올라옵니다.
    // 용역신청서 하단은 로고(J49:M51)와 도장(M51:N52)이 겹치므로,
    // 로고를 먼저 넣고 도장을 나중에 넣어 도장이 로고 위에 보이게 합니다.
    { sheetName: '용역신청서', role: 'serviceLogoFileId', rangeA1: 'i49:M51', hAlign: 'center', vAlign: 'middle' },
    { sheetName: '용역신청서', role: 'stampFileId', rangeA1: 'H25:H28', hAlign: 'right', vAlign: 'middle' },
    { sheetName: '용역신청서', role: 'stampFileId', rangeA1: 'M51:N52', hAlign: 'left', vAlign: 'middle' }
  ],

  DYNAMIC_STAMP_PLACEMENTS: [
    // v75:
    // 비교견적서는 행/견적 조건에 따라 시트 안의 수행사명이 달라질 수 있으므로
    // 메인 수행사 기준 사전삽입 캐시에 넣지 않고, 실제 생성대상 값이 반영된 뒤 해당 셀을 읽어 도장만 삽입합니다.
    // 비교견적(1): O4 수행사명 → T22 도장
    // 비교견적(2): G5 수행사명 → T29 도장
    { sheetName: '비교견적(1)', vendorCellA1: 'O4', role: 'stampFileId', rangeA1: 'T22:T22', hAlign: 'left', vAlign: 'top' },
    { sheetName: '비교견적(2)', vendorCellA1: 'G5', role: 'stampFileId', rangeA1: 'T29:T29', hAlign: 'left', vAlign: 'top' }
  ],

  IMAGE_PRESTAMP_CACHE: {
    // v74 핵심:
    // 도장/로고 삽입은 Google Sheets에서 가장 느린 작업입니다.
    // 매 발송/파일확인마다 insertImage()를 반복하지 않고, 수행사별로 도장/로고가 이미 들어간
    // 생성기 스프레드시트 캐시를 한 번 만든 뒤 그 파일을 복사해서 사용합니다.
    ENABLED: true,
    PROPERTY_PREFIX: 'MAILAUTO_PRESTAMPED_GENERATOR_',
    CACHE_NAME_PREFIX: '메일자동화_도장로고캐시_',

    // v79 중요:
    // 기존 v74~v78은 생성기 스프레드시트의 최종수정시각을 signature에 넣었습니다.
    // 그런데 메일 발송/파일확인 때 자동견적요청 로그를 생성기 파일에 append하므로,
    // 매 실행마다 생성기 최종수정시각이 바뀌어 도장/로고 캐시가 계속 무효화되었습니다.
    // 따라서 기본값은 stable signature입니다. 생성기 양식/도장 위치를 실제로 바꾼 경우
    // 메뉴의 [도장/로고 캐시 초기화] 또는 아래 CACHE_VERSION 값을 올려서 재생성하세요.
    CACHE_VERSION: 'v93_data_format_sync_20260629',
    REBUILD_IF_GENERATOR_UPDATED: false,
    REBUILD_IF_IMAGE_UPDATED: false,
    STRICT_SIGNATURE_MATCH: true
  },

  REVIEW_FILE_GENERATION: {
    // v82:
    // 사용자가 Drive 폴더에서 바로 파일명을 보고 열어 수정할 수 있게 문서별 개별 파일 방식으로 복귀합니다.
    // 견적서/용역신청서/비교견적서는 각각 별도 Google Sheets 파일로 생성됩니다.
    COMBINE_EDITABLE_SHEETS_IN_REVIEW: false,
    COMBINED_SHEETS_FILE_SUFFIX: '_수정용_통합시트',

    // v82:
    // 비편집 자료도 HTML 링크모음으로 만들지 않습니다. Drive 폴더 안에 shortcut 파일로 보이게 합니다.
    COLLAPSE_NON_EDITABLE_REFERENCES_TO_INDEX: false,
    NON_EDITABLE_INDEX_FILENAME: '비편집_자료_링크모음.html',

    // 현재 Apps Script 프로젝트에서 Google Sheets REST API가 꺼져 있어도 빠르게 실패/재시도하지 않도록
    // 파일 확인/수정용 단일 시트 생성은 SpreadsheetApp 기본 경로를 바로 사용합니다.
    USE_SHEETS_API_FAST_PATH_FOR_REVIEW_COPY: false,

    // 비편집 자료 shortcut은 매번 원본 파일 기준으로 새로 만들지 않고, 공용 캐시 폴더에 만들어둔 shortcut을 복사합니다.
    // 캐시 폴더와 캐시 ID는 ScriptProperties에 저장되므로 같은 Apps Script 프로젝트 사용자끼리 공유됩니다.
    USE_CACHED_SHORTCUTS_FOR_NON_EDITABLE_REFERENCES: true,
    SHORTCUT_CACHE_VERSION: 'v82_shortcut_cache_20260622',
    SHORTCUT_CACHE_FOLDER_NAME: '메일자동화_비편집자료_바로가기캐시'
  },

  SELECTION: {
    HIGHLIGHT_ENABLED: false,
    HIGHLIGHT_COLOR: '#fff2cc',
    // v8: 색상 표시 기능 비활성화. 선택행은 사용자별 UserProperties에만 저장합니다.
    PRESERVE_ORIGINAL_BACKGROUNDS: false,
    HIGHLIGHT_MAX_COLUMNS: 45,
    RESTORE_COLOR: null,
    // v56: 현재 활성 선택행이 유효하면 그 행을 우선 사용합니다.
    // 버튼/그림 클릭 등으로 activeRange가 유효하지 않을 때만 UserProperties의 마지막 선택행을 fallback으로 사용합니다.
    USE_REMEMBERED_ROW_FIRST: false,
    USER_PROP_ROW: 'MAILAUTO_SELECTED_MASTER_ROW',
    USER_PROP_SHEET_ID: 'MAILAUTO_SELECTED_MASTER_SHEET_ID',
    USER_PROP_SS_ID: 'MAILAUTO_SELECTED_MASTER_SS_ID',
    USER_PROP_BG: 'MAILAUTO_SELECTED_MASTER_ROW_BG'
  },

  PROGRESS: {
    CACHE_PREFIX: 'MAILAUTO_PROGRESS_',
    CACHE_SECONDS: 21600,
    POLL_MS: 1200,

    // v16: Google Sheets API를 쓰지 않습니다.
    // SpreadsheetApp.appendRow로 행을 추가하고, 접수번호 칸에 임시 UUID를 넣은 뒤
    // 그 UUID가 들어간 실제 행을 찾아 접수번호를 확정합니다.
    // LockService/Sheets API 둘 다 사용하지 않으므로 동시 실행 시 잠금 시간초과와 API 비활성 오류를 피합니다.
    REQUEST_APPEND_USE_SHEETS_API: false,
    REQUEST_MARKER_PREFIX: 'MAILAUTO_PENDING_',
    REQUEST_MARKER_FIND_RETRY: 12,
    REQUEST_MARKER_FIND_SLEEP_MS: 250,

    // 구버전 호환용. v16 기본 흐름에서는 사용하지 않습니다.
    REQUEST_LOCK_WAIT_MS: 12000,
    REQUEST_COUNTER_PROPERTY: 'MAILAUTO_NEXT_REQUEST_NO'
  },

  HIWORKS: {
    // 실제 첨부 발송 성공 코드 기준.
    // Script Properties에는 성공 코드와 동일하게 HIWORKS_API_KEY를 넣습니다.
    ENDPOINT: 'https://api.hiworks.com/office/v2/webmail/sendMail',
    TOKEN_PROPERTY: 'HIWORKS_API_KEY',
    LEGACY_TOKEN_PROPERTY: 'HIWORKS_ACCESS_TOKEN',
    SAVE_SENT_MAIL: 'Y',
    FILE_FIELD_NAME: 'files[]',

    // v41:
    // 기존에는 Apps Script 코드 내부에서 단일 10MB / 전체 20MB를 사전 차단했습니다.
    // 하이웍스 웹메일의 대용량첨부 전환 가능성을 테스트/활용하기 위해
    // 여기서는 용량 초과를 throw 하지 않고 하이웍스 API에 그대로 전달합니다.
    // 0이면 내부 사전 차단 없음. 실제 허용 여부는 하이웍스 API 응답으로 판단합니다.
    MAX_TOTAL_ATTACHMENT_BYTES: 0,
    MAX_SINGLE_ATTACHMENT_BYTES: 0,
    BLOCK_OVERSIZED_ATTACHMENTS_BEFORE_SEND: false,

    // v45/v52:
    // UrlFetchApp 네이티브 multipart는 한글 파일명이 ???로 깨질 수 있어 사용하지 않습니다.
    // 첨부 총량은 하이웍스 API 20MB 제한에 맞춰 관리하고, 수동 multipart에서 파일명 인코딩 헤더를 직접 구성합니다.
    USE_NATIVE_MULTIPART_PAYLOAD: false,

    // 여러 파일 첨부 시 파일 필드명을 files[0], files[1]... 형태로 보냅니다.
    // 단일 파일이면 기존 호환을 위해 files[] 그대로 보냅니다.
    NATIVE_FILE_FIELD_MODE: 'indexed_brackets',

    // v44:
    // 샘플보고서와 수행사정보가 동시에 체크된 경우에는 두 자료를 첨부에서 제외하고
    // 메일 본문에 Google Drive 링크 카드로 안내합니다.
    // 샘플보고서만 체크하거나 수행사정보만 체크한 경우에는 기존처럼 첨부파일로 발송합니다.
    API_ATTACHMENT_LIMIT_BYTES: 20 * 1024 * 1024,
    SAMPLE_AND_CONTRACTOR_LINK_WHEN_BOTH_SELECTED: true,
    MAKE_DRIVE_LINK_FILES_VIEWABLE_BY_LINK: true,

    // 구버전 호환 플래그. v44에서는 조건부 링크 로직을 별도로 사용하므로 기본 false입니다.
    OVERSIZED_ATTACHMENT_FALLBACK_TO_DRIVE_LINK: false,
    MAKE_OVERSIZED_FILE_VIEWABLE_BY_LINK: true
  },

  MAIL: {
    MASTER_CC: 'master@s1samsung.com',
    TEST_DOMAIN: 's1samsung.com',
    INCLUDE_MASTER_CC_IN_TEST: false,
    INCLUDE_SALES_REP_CC_IN_TEST: false,
    // ===== 메일 제목 수정 위치 v17 =====
    // [규칙]
    // 1) 약관/안내문/법령요약(termsGuide)이 체크되어 있으면 최초 발송 제목 사용
    // 2) 그 외에는 선택 파일 유형에 따라 재발송 제목 사용
    //    - 1개 선택:  ... {subjectFileTypes} 파일 송부 ({company})
    //    - 2개 이상: ... {subjectFileTypes} 등 파일 송부 ({company})
    //
    // 사용 가능 변수:
    //   {company}           회사명
    //   {subjectFileTypes}  제목용 파일 유형 조합. 예: 견적서, 선임신고서, 용역신청서
    SUBJECT: {
      INITIAL_TRIGGER_KEY: 'termsGuide',
      INITIAL_TEMPLATE: '[에스원SECOM특약점] {company} 정보통신 유지보수 견적서 및 계약 필요 서류 보내드립니다.',
      RESEND_TEMPLATE_SINGLE: '[에스원SECOM특약점] 정보통신 유지보수 {subjectFileTypes} 파일 송부 ({company})',
      RESEND_TEMPLATE_MULTI: '[에스원SECOM특약점] 정보통신 유지보수 {subjectFileTypes} 등 파일 송부 ({company})',

      // 제목에 표시할 파일명. 본문용 CASE_BODY_TEXT와 별도로 둡니다.
      FILE_TYPE_TEXT: {
        quote: '견적서',
        appointmentDoc: '선임신고서',
        serviceApplication: '용역신청서',
        compareQuote: '비교견적서',
        contractorInfo: '수행사정보',
        serviceStandardContract: '용역표준계약서',
        termsGuide: '약관/안내문/법령요약',
        sampleReport: '샘플보고서'
      },

      // 제목에서 여러 파일을 나열할 때의 순서입니다.
      FILE_TYPE_ORDER: [
        'quote',
        'appointmentDoc',
        'serviceApplication',
        'compareQuote',
        'contractorInfo',
        'serviceStandardContract',
        'termsGuide',
        'sampleReport'
      ]
    },

    // 구버전 호환용. SUBJECT 설정을 지우면 아래 제목 템플릿을 사용합니다.
    SUBJECT_TEMPLATE: '[에스원SECOM특약점] {company} 정보통신설비 유지보수 관련 자료 송부드립니다',

    // ===== 메일 본문 수정 위치 v12 =====
    // [A] 수신/첫 인사/보내는 사람 소개 문구
    // 사용 가능 변수:
    //   {customerContact}      고객사 담당자명. 예: 김종록 과장
    //   {salesRepDisplay}      영업담당자명+직급. 예: 김경아 차장
    //   {company}              회사명
    BODY_HEADER_HTML:
      '<p style="font-family:Malgun Gothic, 맑은 고딕, Arial, sans-serif; font-size:16px; line-height:2.0; margin:0 0 14px 0;">' +
      '<strong><u>수신: {company} {customerContact} 님</u></strong>' +
      '</p>' +

      '<p style="font-family:Malgun Gothic, 맑은 고딕, Arial, sans-serif; font-size:16px; line-height:2.0; margin:0 0 20px 0;">' +
      '안녕하세요? {customerContact}님,<br>' +
      '에스원 정보통신 특약점 담당자 {salesRepDisplay}입니다.' +
      '</p>',
    // [B] 체크박스별 본문에 들어갈 파일/자료명
    // 체크된 항목들을 조합해서 BODY_REQUEST_HTML의 {selectedRequestText}에 넣습니다.
    // 예: 견적서요청 + 비교견적서 + 수행사정보 체크 시
    //     "견적서 및 비교견적서, 케이제이 수행사 정보"
    CASE_BODY_TEXT: {
      quote: '견적서',
      serviceApplication: '용역신청서',
      appointmentDoc: '선임신고서 및 위임장',
      termsGuide: '약관, 안내문, 법령요약 자료',
      contractorInfo: '{vendorDisplayName} 수행사 정보',
      serviceStandardContract: '{vendorDisplayName} 용역표준계약서',
      compareQuote: '비교견적서',
      sampleReport: '샘플보고서'
    },

    // 본문에서 여러 자료를 나열할 때의 순서입니다. 첨부 생성 순서에는 영향 없습니다.
    BODY_FILE_TYPE_ORDER: [
      'quote',
      'serviceApplication',
      'appointmentDoc',
      'termsGuide',
      'contractorInfo',
      'sampleReport',
      'compareQuote',
      'serviceStandardContract'
    ],

    // [C] 체크박스별 자료명을 한 문단으로 합쳐서 나가는 본문
    BODY_REQUEST_HTML:
      '<p style="font-family:\'Malgun Gothic\', \'맑은 고딕\', Arial, sans-serif; font-size:16px; line-height:2.0; margin:0 0 14px 0;">' +
      '안내드린 정보통신 유지보수 <strong>{selectedRequestText}</strong>를 함께 보내드리오니 첨부파일 확인 부탁드리겠습니다.' +
      '</p>',
    // [D] 모든 케이스에 공통으로 나가는 마무리 문구
    BODY_COMMON_HTML:
      '<p style="font-family:\'Malgun Gothic\', \'맑은 고딕\', Arial, sans-serif; font-size:16px; line-height:2.0; margin:0 0 14px 0;">' +
      '관련하여 추가 문의사항이나 요청사항이 있으시면 언제든지 편하게 연락 주시기 바랍니다.<br>' +
      '감사합니다.' +
      '</p>',

    // [E] 영업담당자별 문의사항 문구
    // 사용 가능 변수:
    //   {salesRepDisplay}      예: 김경아 차장
    //   {salesRepPhone}        예: 010-3225-3151
    //   {salesRepEmail}        예: kim@s1samsung.com
    BODY_SIGNATURE_HTML:
      '<p style="font-family:\'Malgun Gothic\', \'맑은 고딕\', Arial, sans-serif; font-size:16px; line-height:2.0; margin:0 0 10px 0;">' +
      '<strong>※ 문의사항: 에스원 특약점 담당자 {salesRepDisplay} ({salesRepPhone}), 이메일 {salesRepEmail}</strong>' +
      '</p>',

    // [F] 메일 본문 하단 이미지 삽입 설정 v13
    // - 모든 메일: 문의사항 문구 아래 영업담당자별 명함 이미지 삽입
    // - 약관, 안내문, 법령요약 체크 시: 명함 아래 안내문 1/2 이미지 추가 삽입
    // - 새 영업담당자 명함 추가 시 BUSINESS_CARD_FILE_IDS에 '담당자명': 'Drive파일ID' 형식으로 추가하세요.
    // - 안내문 변경 시 TERMS_GUIDE_IMAGE_FILE_IDS 배열의 파일 ID만 바꾸면 됩니다.
    INLINE_IMAGE: {
      ENABLED: true,
      MAKE_VIEWABLE_BY_LINK: true,
      BUSINESS_CARD_WIDTH_PX: 500,
      GUIDE_WIDTH_PX: 800,
      GAP_HTML: '<br>',
      BUSINESS_CARD_FILE_IDS: {
        '김경아': '1IUeMFGM3lIBbg07O8S-mT7A_dTG1Gt8w',
        '이옥희': '1pYDPM4bN3WLB48LZaOOHXB1MDy87WwRI',
        '최보람': '1lAaO5J0urpePkmQQ-pjFO4fT7VIq8jFr'
      },
      TERMS_GUIDE_IMAGE_FILE_IDS: [
        '1qwOCG7rTbc5WRmzH_CqFN8B8Emd2vShG',
        '1IFLhSSfpRkDmPAj5VZ_RXPtw0le354jK'
      ]
    },

    // 구버전 호환용. BODY_REQUEST_HTML/CASE_BODY_TEXT를 지우면 아래 CASE_BODY_HTML 방식으로 fallback 가능합니다.
    CASE_BODY_HTML: {
      quote: '<p>정보통신설비 유지관리 견적서를 첨부드립니다.</p>',
      serviceApplication: '<p>용역신청서를 첨부드립니다. 내용 확인 후 회신 부탁드립니다.</p>',
      appointmentDoc: '<p>선임신고서 및 위임장 양식을 첨부드립니다. 필요 사항 기재 후 회신 부탁드립니다.</p>',
      termsGuide: '<p>약관, 안내문, 법령요약 자료를 첨부드립니다.</p>',
      contractorInfo: '<p>{vendorDisplayName} 수행사 정보를 첨부드립니다.</p>',
      serviceStandardContract: '<p>{vendorDisplayName} 용역표준계약서를 첨부드립니다.</p>',
      compareQuote: '<p>비교견적서를 첨부드립니다.</p>',
      sampleReport: '<p>샘플보고서를 첨부드립니다.</p>'
    },

    BODY_TEMPLATE_HTML:
      '<p>안녕하세요. {company} 담당자님.</p>' +
      '<p>정보통신설비 유지관리 관련 요청 자료를 첨부드립니다.</p>' +
      '<p>확인 부탁드립니다.</p>' +
      '<p>감사합니다.<br>{salesRepName} 드림</p>'
  },


  SENT_FILE_ARCHIVE: {
    // v86:
    // CUSTOMER 모드 발송 성공 후, 실제 발송된 첨부 Blob을 고객사 파일관리 공유드라이브 폴더에 누적 저장합니다.
    // 폴더 결정은 기존 고객사 폴더 자동생성 스크립트와 동일하게
    // 공유드라이브명 + 고객번호 prefix 폴더명 검색을 우선 사용합니다.
    // 마스터시트의 고객사폴더ID는 보조/검증용 캐시로만 사용합니다.
    ENABLED: true,

    // 공유드라이브 자체 이름. Apps Script가 Drive API로 이 이름의 공유드라이브를 찾아 driveId를 캐시합니다.
    SHARED_DRIVE_NAME: 'S1 고객사 파일 관리',
    SHARED_DRIVE_ID: '',
    SHARED_DRIVE_ID_PROPERTY: 'S1_CUSTOMER_SHARED_DRIVE_ID',

    // 마스터시트(신규) 헤더명 기준. 현재 BJ열 '고객사폴더ID'는 보조 검증/캐시로 사용합니다.
    MASTER_FOLDER_ID_HEADERS: ['고객사폴더ID', '고객사 폴더 ID', '고객폴더ID', '고객 폴더 ID', '폴더ID'],

    // 예외 fallback용. 기본 운영에서는 필요 없습니다.
    ROOT_FOLDER_ID: '',
    ROOT_FOLDER_ID_PROPERTY: 'S1_CUSTOMER_FILE_ROOT_FOLDER_ID',

    // TEST 발송은 고객사에 실제 발송한 자료가 아니므로 기본 저장하지 않습니다.
    SAVE_TEST_MAIL: false,

    SENT_PREFIX: '[발송]',
    LOG_SHEET_NAME: '발송파일로그',
    FOLDER_INDEX_SHEET_NAME: '고객사파일폴더인덱스',
    FOLDER_HISTORY_SHEET_NAME: '_메일이력_발송',

    // 고객사 폴더를 못 찾았을 때 새 폴더 생성 허용.
    // 기존 폴더 인덱스가 충분히 잡힌 뒤에는 false로 바꿔도 됩니다.
    CREATE_FOLDER_IF_MISSING: true,

    // 발송자료는 "발송 이력"이므로 같은 파일명/같은 내용이어도 skip하지 않고 _001, _002로 누적 저장합니다.
    ALWAYS_ACCUMULATE: true,

    // 로그용 SHA256. 첨부 Blob 기준으로 계산합니다. 속도가 너무 느리면 false로 바꿔도 파일 저장 자체에는 영향 없습니다.
    CALCULATE_SHA256: true,

    // v94: CUSTOMER 발송 완료 후 고객사 공유드라이브 누적 저장은 사용자 대기시간을 늘리지 않도록
    // 파일 확인/수정 후 발송 건부터 백그라운드 큐로 넘깁니다.
    ASYNC_AFTER_SEND: true,
    ASYNC_REVIEW_ONLY: true,
    QUEUE_SHEET_NAME: '발송파일저장큐',
    ASYNC_TRIGGER_HANDLER: 'processDeferredSentFileArchiveQueueV94',
    ASYNC_TRIGGER_DELAY_MS: 60 * 1000,
    MAX_ASYNC_JOBS_PER_RUN: 3,

    // 중앙 발송파일로그를 고객사 폴더별 _메일이력_발송 Google Sheet로 하루 1회 반영합니다.
    // XLSX를 직접 수정하지 않고, 고객사 폴더 안에 Google Sheet 파일을 생성/갱신합니다.
    DAILY_HISTORY_SYNC_ENABLED: true,
    DAILY_HISTORY_SYNC_HOUR: 19,
    MAX_HISTORY_SYNC_ROWS_PER_RUN: 300
  },

  TEMP: {
    PARENT_FOLDER_ID: '',
    TRASH_TEMP_SPREADSHEET_AFTER_SUCCESS: true,
    TRASH_TEMP_DOC_AFTER_SUCCESS: true,
    KEEP_TEMP_ON_FAILURE: true
  },

  EXPORT: {
    WAIT_MS_AFTER_IMAGE_INSERT: 800,
    PDF_OPTIONS: {
      size: 'A4',
      portrait: 'true',

      // v24 기준 PDF 출력 설정
      // - 사용자가 보내준 기준 PDF처럼 A4 한 장 안에서 적당한 여백을 두고 출력
      // - scale=4는 Google Sheets PDF export의 "fit to page"입니다.
      // - fitw=true만 쓰면 가로만 꽉 차면서 너무 크게 나올 수 있어서 scale=4를 같이 둡니다.
      scale: '4',
      fitw: 'true',

      sheetnames: 'false',
      printtitle: 'false',
      pagenumbers: 'false',
      gridlines: 'false',
      fzr: 'false',
      printnotes: 'false',
      horizontal_alignment: 'CENTER',
      vertical_alignment: 'TOP',

      // Google Sheets export margin 단위: inch
      // 기준 샘플은 기존 0.15보다 훨씬 작게/가운데 출력되어 보여서 0.60으로 맞춤.
      // 더 크게 보이면 0.50, 더 작게 보이면 0.70으로 조정하면 됩니다.
      top_margin: '0.60',
      bottom_margin: '0.60',
      left_margin: '0.60',
      right_margin: '0.60'
    }
  }
});


function doGet(e) {
  const action = String(e && e.parameter && e.parameter.action || '').trim();
  if (action === 'mailAutoMultiDownload') {
    return serveMailAutoMultiDownload_(e);
  }

  return HtmlService
    .createHtmlOutput('<p>메일자동화 다운로드 페이지입니다.</p>')
    .setTitle('메일자동화 다운로드');
}

function serveMailAutoMultiDownload_(e) {
  const token = String(e && e.parameter && e.parameter.token || '').trim();
  if (!token) {
    return HtmlService
      .createHtmlOutput('<p style="font-family:Arial,sans-serif;">다운로드 토큰이 없습니다.</p>')
      .setTitle('다운로드 오류');
  }

  const key = 'MAILAUTO_MULTI_DOWNLOAD_' + token;
  const raw = PropertiesService.getScriptProperties().getProperty(key);
  if (!raw) {
    return HtmlService
      .createHtmlOutput('<p style="font-family:Arial,sans-serif;">다운로드 정보가 만료되었거나 존재하지 않습니다. 발신자에게 다시 요청해 주세요.</p>')
      .setTitle('다운로드 정보 없음');
  }

  let data;
  try {
    data = JSON.parse(raw);
  } catch (err) {
    return HtmlService
      .createHtmlOutput('<p style="font-family:Arial,sans-serif;">다운로드 정보를 읽지 못했습니다. 발신자에게 다시 요청해 주세요.</p>')
      .setTitle('다운로드 오류');
  }

  const files = Array.isArray(data.files) ? data.files.filter(function(file) {
    return file && file.downloadUrl;
  }) : [];

  if (!files.length) {
    return HtmlService
      .createHtmlOutput('<p style="font-family:Arial,sans-serif;">다운로드할 파일이 없습니다.</p>')
      .setTitle('다운로드 파일 없음');
  }

  const safeFilesJson = JSON.stringify(files).replace(/</g, '\\u003c');
  const fallbackLinksHtml = files.map(function(file, idx) {
    const name = escapeHtml_(file.name || ('파일 ' + (idx + 1)));
    const downloadUrl = escapeHtml_(file.downloadUrl || file.url || '');
    return '<li style="margin:7px 0;"><a href="' + downloadUrl + '" target="_blank" rel="noopener" style="color:#1a73e8;text-decoration:none;font-weight:700;">' + name + '</a></li>';
  }).join('');

  const html = '' +
    '<!DOCTYPE html><html lang="ko"><head><meta charset="UTF-8">' +
    '<meta name="viewport" content="width=device-width, initial-scale=1.0">' +
    '<title>파일 다운로드 중</title></head>' +
    '<body style="margin:0;background:#ffffff;font-family:Malgun Gothic, 맑은 고딕, Arial, sans-serif;color:#202124;">' +
      '<div id="box" style="max-width:520px;margin:34px auto;padding:20px 22px;border:1px solid #e5e7eb;border-radius:12px;background:#fff;box-shadow:0 4px 18px rgba(0,0,0,.06);">' +
        '<div id="status" style="font-size:15px;font-weight:700;color:#202124;line-height:1.6;">파일 다운로드를 시작합니다.</div>' +
        '<div id="sub" style="margin-top:8px;font-size:12px;color:#6f7682;line-height:1.6;">다운로드가 시작되면 이 창은 자동으로 닫기를 시도합니다.</div>' +
        '<div id="fallback" style="display:none;margin-top:14px;font-size:13px;color:#3c4043;line-height:1.6;">' +
          '<div style="margin-bottom:6px;">창이 자동으로 닫히지 않거나 다운로드가 막히면 아래 파일명을 눌러 다운로드해 주세요.</div>' +
          '<ol style="margin:0;padding-left:20px;">' + fallbackLinksHtml + '</ol>' +
        '</div>' +
      '</div>' +
      '<script>' +
        'const files=' + safeFilesJson + ';' +
        'let started=false;' +
        'function setText(id,text){const el=document.getElementById(id);if(el)el.textContent=text;}' +
        'function tryCloseLauncher(){' +
          'setText("status","다운로드를 실행했습니다.");' +
          'setText("sub","이 창을 자동으로 닫는 중입니다. 자동으로 닫히지 않으면 그냥 닫아 주세요.");' +
          'try{window.open("","_self");}catch(e){}' +
          'try{window.close();}catch(e){}' +
          'setTimeout(function(){' +
            'const fb=document.getElementById("fallback");if(fb)fb.style.display="block";' +
            'setText("status","다운로드 실행 완료");' +
            'setText("sub","브라우저 정책상 자동 닫기가 차단될 수 있습니다. 이 창은 닫아도 됩니다.");' +
          '},900);' +
        '}' +
        'function startDownloads(){' +
          'if(started)return; started=true;' +
          'setText("status","파일 " + files.length + "개 다운로드 실행 중...");' +
          'files.forEach(function(file,idx){' +
            'setTimeout(function(){' +
              'const a=document.createElement("a");' +
              'a.href=file.downloadUrl;' +
              'a.target="_blank";' +
              'a.rel="noopener";' +
              'a.style.display="none";' +
              'document.body.appendChild(a);' +
              'a.click();' +
              'setTimeout(function(){try{a.remove();}catch(e){}},1200);' +
            '}, idx*700);' +
          '});' +
          'setTimeout(tryCloseLauncher, Math.max(1600, files.length*700 + 900));' +
        '}' +
        'window.addEventListener("load",function(){setTimeout(startDownloads,250);});' +
      '</script>' +
    '</body></html>';

  return HtmlService
    .createHtmlOutput(html)
    .setTitle('파일 다운로드 중')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}


function onSelectionChange(e) {
  // 사용자가 마스터시트의 데이터 셀을 선택할 때마다 해당 행을 유저별로 기억합니다.
  // v8: 색상 표시 없이 UserProperties에만 저장합니다.
  // 버튼/그림 클릭 시 activeRange가 헤더 쪽으로 튀어도 이 저장값의 행을 사용합니다.
  try {
    new MailAutomationService().rememberSelectionFromEvent_(e);
  } catch (err) {
    Logger.log('onSelectionChange 처리 실패: ' + (err && err.stack || err));
  }
}

function rememberCurrentSelectedMailRow() {
  const result = new MailAutomationService().rememberSelectionFromEvent_();
  SpreadsheetApp.getUi().alert(
    '선택행 저장 완료',
    result ? ('마지막 선택 행: ' + result.rowNo + '행') : '저장할 데이터 행을 찾지 못했습니다.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function clearMyMailRowHighlight() {
  new MailAutomationService().clearMySelectionHighlight_();
  SpreadsheetApp.getUi().alert('내 선택행 기억을 지웠습니다.');
}

function 메일자동발송() {
  mailAutoSend();
}

function mailAutoSend() {
  try {
    const preview = new MailAutomationService().previewFromActiveSelection();
    const tpl = HtmlService.createTemplateFromFile('MailConfirmDialog');
    tpl.previewJson = JSON.stringify(preview).replace(/</g, '\\u003c');
    const html = tpl.evaluate().setWidth(760).setHeight(680);
    SpreadsheetApp.getUi().showModalDialog(html, '메일자동발송');
  } catch (err) {
    SpreadsheetApp.getUi().alert('메일 발송 준비 중 오류', String(err.message || err), SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

function sendMailFromDialog(payload) {
  return new MailAutomationService().sendFromDialog(payload);
}

function prepareMailFilesForReview(payload) {
  return new MailAutomationService().prepareFilesForReview(payload);
}

function getMailRunProgress(runId) {
  return new ProgressTracker(runId).get();
}
function cancelMailRun(runId) {
  return new ProgressTracker(runId).requestCancel();
}
function clearMyMailRunProgress() {
  new ProgressTracker('').clearAllForUser_();
  SpreadsheetApp.getUi().alert('내 진행상태 캐시를 초기화했습니다.');
}

function showHiworksTokenGuide() {
  SpreadsheetApp.getUi().alert(
    '하이웍스 토큰 저장',
    '성공했던 기존 코드와 동일하게 Script Properties의 HIWORKS_API_KEY 값을 사용합니다.\n\n' +
    '메뉴에서 [메일자동화 > 하이웍스 API키 저장]을 눌러 OfficeToken/API Key를 저장하세요.\n' +
    '기존에 HIWORKS_ACCESS_TOKEN, HIWORKS_OFFICE_TOKEN, HIWORKS_TOKEN 이름으로 저장한 값도 자동 호환됩니다.',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

function setHiworksApiKey() {
  const ui = SpreadsheetApp.getUi();
  const res = ui.prompt(
    '하이웍스 API 키 / OfficeToken 저장',
    '하이웍스 OfficeToken 또는 API Key 값을 입력하세요.\nBearer 는 빼고 토큰 문자열만 붙여넣으시면 됩니다.',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) {
    ui.alert('저장을 취소했습니다.');
    return;
  }

  const token = String(res.getResponseText() || '').trim().replace(/^Bearer\s+/i, '');
  if (!token) {
    ui.alert('입력값이 비어 있습니다.');
    return;
  }

  // 성공했던 기존 코드의 키명을 1순위로 저장합니다.
  const props = PropertiesService.getScriptProperties();
  props.setProperty('HIWORKS_API_KEY', token);
  props.setProperty(CONFIG.HIWORKS.TOKEN_PROPERTY, token);

  ui.alert('HIWORKS_API_KEY 저장 완료');
}

function checkHiworksApiKey() {
  const info = findHiworksTokenInfo_();
  if (!info.token) {
    SpreadsheetApp.getUi().alert(
      'HIWORKS_API_KEY가 없습니다.\n\n' +
      '메뉴에서 [메일자동화 > 하이웍스 API키 저장]을 먼저 실행하세요.'
    );
    return;
  }

  SpreadsheetApp.getUi().alert(
    '하이웍스 토큰 저장 확인됨\n\n' +
    '저장 위치: ' + info.scope + '\n' +
    '속성명: ' + info.key + '\n' +
    '앞 6자리: ' + info.token.slice(0, 6) + '\n' +
    '길이: ' + info.token.length
  );
}

function findHiworksTokenInfo_() {
  const keys = [
    CONFIG.HIWORKS.TOKEN_PROPERTY,
    'HIWORKS_API_KEY',
    CONFIG.HIWORKS.LEGACY_TOKEN_PROPERTY,
    'HIWORKS_ACCESS_TOKEN',
    'HIWORKS_OFFICE_TOKEN',
    'HIWORKS_TOKEN',
    'OFFICE_TOKEN',
    'API_KEY'
  ].filter(function(v, i, arr) { return v && arr.indexOf(v) === i; });

  const stores = [
    { scope: 'ScriptProperties', store: PropertiesService.getScriptProperties() },
    { scope: 'UserProperties', store: PropertiesService.getUserProperties() },
    { scope: 'DocumentProperties', store: PropertiesService.getDocumentProperties() }
  ];

  for (const item of stores) {
    for (const key of keys) {
      const value = String(item.store.getProperty(key) || '').trim().replace(/^Bearer\s+/i, '');
      if (value) {
        return { token: value, key: key, scope: item.scope };
      }
    }
  }

  return { token: '', key: '', scope: '' };
}

function setHiworksAccessTokenForSetup_() {
  const token = '여기에_하이웍스_API_KEY_또는_TOKEN_붙여넣기';
  if (!token || token.indexOf('여기에_') === 0) {
    throw new Error('token 변수에 실제 하이웍스 API Key/Token을 넣은 뒤 실행하세요. 메뉴의 하이웍스 API키 저장을 써도 됩니다.');
  }
  PropertiesService.getScriptProperties().setProperty('HIWORKS_API_KEY', token.replace(/^Bearer\s+/i, ''));
}



function warmUpMailAutoReviewShortcutCache() {
  const dummyProgress = {
    update: function(percent, message) {
      Logger.log('[비편집 shortcut 캐시 예열] ' + percent + '% / ' + message);
    }
  };

  const vendorNames = Object.keys(CONFIG.VENDORS || {})
    .map(function(name) { return normalizeVendorName_(name); })
    .filter(Boolean)
    .filter(function(name, idx, arr) { return arr.indexOf(name) === idx; });

  const ss = SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID);
  const logs = [];
  const warmedKeys = {};

  function warmFile(builder, fileId, displayName, group) {
    const id = String(fileId || '').trim();
    if (!id) return;
    const key = id + '|' + String(displayName || '');
    if (warmedKeys[key]) return;
    warmedKeys[key] = true;
    try {
      const file = DriveApp.getFileById(id);
      const name = String(displayName || file.getName() || '자료 파일').trim();
      const cached = builder.getOrCreateCachedReviewShortcut_(file, name);
      logs.push((group || '자료') + ': ' + name + ' → ' + cached.getId());
    } catch (err) {
      logs.push((group || '자료') + ': 실패 / ' + id + ' / ' + String(err && err.message || err));
    }
  }

  const baseBuilder = new ReviewFilePackageBuilder(ss, { '회사명': '캐시예열', '수행사': '' }, dummyProgress);

  (CONFIG.FILE_DEFINITIONS || []).forEach(function(def) {
    if (!def) return;
    if (def.type === 'static_files') {
      (def.fileIds || []).forEach(function(fileId) {
        warmFile(baseBuilder, fileId, '', def.label || '고정자료');
      });
    }
    if (def.type === 'drive_folder_files') {
      const folderId = String(def.folderId || '').trim();
      if (!folderId) return;
      try {
        const folder = DriveApp.getFolderById(folderId);
        const files = folder.getFiles();
        while (files.hasNext()) {
          const file = files.next();
          warmFile(baseBuilder, file.getId(), file.getName(), def.label || '폴더자료');
        }
      } catch (err) {
        logs.push((def.label || '폴더자료') + ': 폴더 접근 실패 / ' + folderId + ' / ' + String(err && err.message || err));
      }
    }
  });

  vendorNames.forEach(function(vendorName) {
    const builder = new ReviewFilePackageBuilder(ss, { '회사명': '캐시예열', '수행사': vendorName }, dummyProgress);
    const vendor = getVendorConfig_(vendorName) || {};
    (CONFIG.FILE_DEFINITIONS || []).forEach(function(def) {
      if (!def) return;
      if (def.type === 'vendor_zip') {
        const fileId = String(vendor.contractorInfoZipFileId || '').trim();
        const displayName = def.filename
          ? def.filename.replace(/\{vendor\}/g, vendor.displayName || vendorName)
          : '';
        warmFile(builder, fileId, displayName, (def.label || '수행사정보') + '/' + vendorName);
      }
    });
  });

  SpreadsheetApp.getUi().alert(
    '비편집 자료 shortcut 캐시 예열 완료\n\n' +
    '처리: ' + logs.length + '건\n' +
    logs.slice(0, 25).join('\n') +
    (logs.length > 25 ? '\n...외 ' + (logs.length - 25) + '건' : '')
  );
}

function clearMailAutoReviewShortcutCache() {
  const props = PropertiesService.getScriptProperties();
  const all = props.getProperties();
  Object.keys(all).forEach(function(key) {
    if (key.indexOf('MAILAUTO_REVIEW_SHORTCUT_CACHE_') === 0) {
      props.deleteProperty(key);
    }
  });
  SpreadsheetApp.getUi().alert('비편집 자료 shortcut 캐시 정보를 초기화했습니다. 기존 캐시 폴더/파일은 Drive에 남아 있을 수 있습니다.');
}

class ProgressTracker {
  constructor(runId) {
    this.runId = String(runId || '').trim();
    this.key = this.runId ? (CONFIG.PROGRESS.CACHE_PREFIX + this.runId) : '';
    this.cancelKey = this.key ? (this.key + '_CANCEL') : '';
  }

  start(message, percent) {
    this.put_({
      runId: this.runId,
      status: 'RUNNING',
      percent: Number(percent) || 1,
      message: message || '시작 중',
      startedAt: Date.now(),
      updatedAt: Date.now()
    });
  }

  update(percent, message) {
    const current = this.read_() || {};
    const startedAt = current.startedAt || Date.now();
    const p = Math.max(1, Math.min(99, Number(percent) || current.percent || 1));
    const elapsedSec = Math.max(0, Math.round((Date.now() - startedAt) / 1000));
    const remainingSec = p > 2 ? Math.max(0, Math.round(elapsedSec * (100 - p) / p)) : null;
    this.put_({
      runId: this.runId,
      status: 'RUNNING',
      percent: p,
      message: message || current.message || '처리 중',
      startedAt: startedAt,
      updatedAt: Date.now(),
      elapsedSec: elapsedSec,
      remainingSec: remainingSec
    });
  }

  done(message) {
    const current = this.read_() || {};
    const startedAt = current.startedAt || Date.now();
    this.put_({
      runId: this.runId,
      status: 'DONE',
      percent: 100,
      message: message || '완료',
      startedAt: startedAt,
      updatedAt: Date.now(),
      elapsedSec: Math.max(0, Math.round((Date.now() - startedAt) / 1000)),
      remainingSec: 0
    });
  }

  fail(message) {
    const current = this.read_() || {};
    const startedAt = current.startedAt || Date.now();
    this.put_({
      runId: this.runId,
      status: 'FAIL',
      percent: current.percent || 0,
      message: message || '실패',
      startedAt: startedAt,
      updatedAt: Date.now(),
      elapsedSec: Math.max(0, Math.round((Date.now() - startedAt) / 1000)),
      remainingSec: null
    });
  }
    requestCancel() {
    if (!this.cancelKey) {
      return { ok: false, message: '실행 ID가 없습니다.' };
    }

    const text = JSON.stringify({
      requested: true,
      requestedAt: Date.now()
    });

    const ttl = CONFIG.PROGRESS.CACHE_SECONDS || 21600;

    CacheService.getUserCache().put(this.cancelKey, text, ttl);
    CacheService.getScriptCache().put(this.cancelKey, text, ttl);

    const current = this.read_() || {};
    if (current.status !== 'DONE' && current.status !== 'FAIL') {
      current.cancelRequested = true;
      current.message = '메일 발송 취소 요청됨. 하이웍스 발송 전 단계에서 중단합니다.';
      current.updatedAt = Date.now();
      this.put_(current);
    }

    return { ok: true, message: '취소 요청을 보냈습니다.' };
  }

  isCancelRequested() {
    if (!this.cancelKey) return false;

    return Boolean(
      CacheService.getUserCache().get(this.cancelKey) ||
      CacheService.getScriptCache().get(this.cancelKey)
    );
  }

  throwIfCancelRequested_(stageMessage) {
    if (!this.isCancelRequested()) return;

    this.fail('메일 발송이 취소되었습니다. ' + (stageMessage || ''));

    throw new Error('메일 발송이 취소되었습니다. 하이웍스 API 발송 전 단계에서 중단했습니다.');
  }
  get() {
    const current = this.read_();
    if (!current) {
      return { status: 'PENDING', percent: 0, message: '대기 중', elapsedSec: 0, remainingSec: null };
    }
    if (current.status === 'RUNNING' && current.startedAt) {
      const elapsedSec = Math.max(0, Math.round((Date.now() - current.startedAt) / 1000));
      current.elapsedSec = elapsedSec;
      current.remainingSec = current.percent > 2
        ? Math.max(0, Math.round(elapsedSec * (100 - current.percent) / current.percent))
        : null;
    }
    return current;
  }

  read_() {
    if (!this.key) return null;
    const raw = CacheService.getUserCache().get(this.key) || CacheService.getScriptCache().get(this.key);
    if (!raw) return null;
    try { return JSON.parse(raw); } catch (err) { return null; }
  }

  put_(obj) {
    if (!this.key) return;
    const text = JSON.stringify(obj);
    const ttl = CONFIG.PROGRESS.CACHE_SECONDS || 21600;
    CacheService.getUserCache().put(this.key, text, ttl);
    CacheService.getScriptCache().put(this.key, text, ttl);
  }

  clearAllForUser_() {
    // runId를 모르는 과거 캐시는 TTL 만료로 사라집니다. 현재 실행 중인 건 클라이언트 runId 기준으로 덮어씁니다.
    return true;
  }
}

class MailAutomationService {
  previewFromActiveSelection() {
    const master = this.getMasterContext_();

    // v8 핵심:
    // 색상 표시는 하지 않고, 사용자가 마지막으로 클릭한 데이터 행을 UserProperties에서 읽어 발송합니다.
    // 버튼/그림 클릭 순간 activeRange가 헤더/버튼 영역으로 튀어도 UserProperties의 행을 우선 사용합니다.
    const rowNo = this.resolveSelectedMasterRow_(master);

    const rowObj = this.readMasterRow_(master.sheet, master.headerMap, rowNo);
    const selectedDefs = this.getSelectedFileDefs_(rowObj);
    if (!selectedDefs.length) {
      throw new Error('체크열에서 발송할 파일을 하나 이상 체크하세요.');
    }

    const company = this.valueByHeader_(rowObj, '회사명') || '(회사명 없음)';
    const recipient = this.valueByHeader_(rowObj, '담당자 이메일 주소') || '';
    const salesRep = this.valueByHeader_(rowObj, '영업담당자') || '';
    const contractConditionText = buildContractConditionPreview_(rowObj);

    let defaultTo = splitEmails_(recipient);
    let defaultCc = [];
    try {
      const generatorSs = SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID);
      const sender = new SalesRepResolver(generatorSs).resolve(salesRep);
      defaultCc = uniqueEmails_([sender.email, CONFIG.MAIL.MASTER_CC]);
    } catch (err) {
      // 미리보기 단계에서 영업담당자 정보 확인이 실패해도 팝업 자체는 열리게 합니다.
      // 실제 발송 단계에서는 다시 검증합니다.
      defaultCc = uniqueEmails_([CONFIG.MAIL.MASTER_CC]);
    }

    return {
      rowNo,
      selectedCell: '마지막 선택 ' + rowNo + '행',
      company,
      recipient,
      cc: defaultCc.join(', '),
      contractConditionText,
      defaultTo,
      defaultCc,
      salesRep,
      selectedLabels: selectedDefs.map(d => d.label),
      selectedKeys: selectedDefs.map(d => d.key),
      estimatedSeconds: estimateSecondsForDefinitions_(selectedDefs)
    };
  }


  applyDialogFileSelectionOverrides_(selectedDefs, payload) {
    const defs = selectedDefs || [];
    return defs.map(def => {
      if (!def || def.key !== 'compareQuote' || def.type !== 'multi_sheet_pdf' || !Array.isArray(def.sheets)) {
        return def;
      }

      const selectedSheetNames = this.resolveCompareQuoteSheetNamesFromPayload_(payload, def);
      const filteredSheets = (def.sheets || [])
        .filter(item => item && selectedSheetNames.indexOf(String(item.sheetName || '').trim()) >= 0)
        .map(item => Object.assign({}, item));

      if (!filteredSheets.length) return null;

      const cloned = Object.assign({}, def, {
        sheets: filteredSheets,
        label: buildCompareQuoteLabelFromSheets_(filteredSheets)
      });

      return cloned;
    }).filter(Boolean);
  }


  filterEditableReviewFileDefs_(selectedDefs) {
    const editableKeys = {
      quote: true,
      serviceApplication: true,
      appointmentDoc: true,
      compareQuote: true
    };

    return (selectedDefs || [])
      .filter(function(def) {
        return def && editableKeys[String(def.key || '').trim()] === true;
      })
      .map(function(def) {
        return Object.assign({}, def);
      });
  }

  resolveCompareQuoteSheetNamesFromPayload_(payload, def) {
    const allSheetNames = (def && Array.isArray(def.sheets) ? def.sheets : [])
      .map(item => String(item && item.sheetName || '').trim())
      .filter(Boolean);

    const rawSelected = payload && (payload.compareQuoteSheets || payload.selectedCompareQuoteSheets);
    if (Array.isArray(rawSelected) && rawSelected.length) {
      const selectedSet = buildLowerStringSet_(rawSelected);
      const selected = allSheetNames.filter(name => selectedSet[String(name || '').trim().toLowerCase()]);
      if (selected.length) return selected;
    }

    const rawExcluded = payload && (payload.excludedCompareQuoteSheets || payload.excludedCompareSheets || payload.removedCompareQuoteSheets);
    if (Array.isArray(rawExcluded) && rawExcluded.length) {
      const excludedSet = buildLowerStringSet_(rawExcluded);
      const selected = allSheetNames.filter(name => !excludedSet[String(name || '').trim().toLowerCase()]);
      if (selected.length) return selected;
    }

    return allSheetNames;
  }

  resolveSelectedMasterRow_(master) {
    const remembered = this.getRememberedSelectedRow_(master);
    const activeRow = this.getCurrentUiSelectedRow_(master);

    // v56 핵심:
    // 현재 사용자가 실제로 선택하고 있는 활성 행이 유효하면 그 행을 최우선으로 사용합니다.
    // 기존에는 UserProperties에 저장된 과거 행을 먼저 써서,
    // 화면상 4행을 선택해도 이전에 기억된 3행이 팝업에 뜨는 문제가 있었습니다.
    if (activeRow) {
      this.storeSelectedRow_(master.sheet, activeRow);
      return activeRow;
    }

    // 버튼/그림 클릭 등으로 activeRange가 헤더나 버튼 영역으로 튄 경우에만
    // 마지막으로 기억된 데이터 행을 fallback으로 사용합니다.
    if (remembered) return remembered;

    throw new Error(
      '발송할 데이터 행을 찾지 못했습니다.\n\n' +
      '마스터시트에서 보낼 행의 아무 셀이나 한 번 클릭한 뒤 다시 메일자동발송을 눌러주세요.'
    );
  }

  getCurrentUiSelectedRow_(master) {
    try {
      const activeSs = SpreadsheetApp.getActiveSpreadsheet();
      if (!activeSs || activeSs.getId() !== CONFIG.MASTER_SPREADSHEET_ID) return 0;

      const activeSheet = activeSs.getActiveSheet();
      if (!activeSheet || activeSheet.getName() !== CONFIG.SHEETS.MASTER) return 0;

      const range = activeSheet.getActiveRange();
      if (!range) return 0;

      const rowNo = range.getRow();
      if (rowNo < CONFIG.ROWS.MASTER_DATA_START) return 0;
      return rowNo;
    } catch (err) {
      Logger.log('현재 UI 선택행 확인 실패: ' + (err && err.stack || err));
      return 0;
    }
  }

  getRememberedSelectedRow_(master) {
    try {
      const props = PropertiesService.getUserProperties();
      const ssId = props.getProperty(CONFIG.SELECTION.USER_PROP_SS_ID);
      const sheetId = props.getProperty(CONFIG.SELECTION.USER_PROP_SHEET_ID);
      const rowNo = Number(props.getProperty(CONFIG.SELECTION.USER_PROP_ROW));

      if (ssId !== CONFIG.MASTER_SPREADSHEET_ID) return 0;
      if (String(sheetId) !== String(master.sheet.getSheetId())) return 0;
      if (!rowNo || rowNo < CONFIG.ROWS.MASTER_DATA_START) return 0;
      if (rowNo > master.sheet.getMaxRows()) return 0;

      return rowNo;
    } catch (err) {
      Logger.log('저장 선택행 확인 실패: ' + (err && err.stack || err));
      return 0;
    }
  }

  rememberSelectionFromEvent_(e) {
    const range = e && e.range ? e.range : SpreadsheetApp.getActiveRange();
    if (!range) return null;

    const sheet = range.getSheet();
    if (!sheet || sheet.getName() !== CONFIG.SHEETS.MASTER) return null;
    if (sheet.getParent().getId() !== CONFIG.MASTER_SPREADSHEET_ID) return null;

    const rowNo = range.getRow();
    if (rowNo < CONFIG.ROWS.MASTER_DATA_START) return null;

    this.storeSelectedRow_(sheet, rowNo);

    return { rowNo: rowNo };
  }

  storeSelectedRow_(sheet, rowNo) {
    const props = PropertiesService.getUserProperties();
    props.setProperty(CONFIG.SELECTION.USER_PROP_SS_ID, sheet.getParent().getId());
    props.setProperty(CONFIG.SELECTION.USER_PROP_SHEET_ID, String(sheet.getSheetId()));
    props.setProperty(CONFIG.SELECTION.USER_PROP_ROW, String(rowNo));
  }

  getSelectionHighlightWidth_(sheet) {
    return Math.min(
      Math.max(sheet.getLastColumn(), 1),
      CONFIG.SELECTION.HIGHLIGHT_MAX_COLUMNS || sheet.getLastColumn()
    );
  }

  markSelectedRow_(sheet, rowNo) {
    if (!CONFIG.SELECTION || CONFIG.SELECTION.HIGHLIGHT_ENABLED !== true) return;
    if (!rowNo || rowNo < CONFIG.ROWS.MASTER_DATA_START) return;

    const props = PropertiesService.getUserProperties();
    const width = this.getSelectionHighlightWidth_(sheet);

    // v8에서는 HIGHLIGHT_ENABLED가 false라 보통 실행되지 않습니다.
    // 다시 색 표시가 필요할 때를 대비해 함수는 남겨둡니다.
    // 같은 행을 다시 선택한 경우 불필요한 시트 쓰기 작업을 하지 않습니다.
    const raw = props.getProperty(CONFIG.SELECTION.USER_PROP_BG);
    if (raw) {
      try {
        const saved = JSON.parse(raw);
        if (
          saved &&
          saved.ssId === sheet.getParent().getId() &&
          String(saved.sheetId) === String(sheet.getSheetId()) &&
          Number(saved.rowNo) === Number(rowNo)
        ) {
          return;
        }
      } catch (err) {
        // 저장값 파싱 실패 시 아래에서 새로 표시합니다.
      }
    }

    // 기존 배경 완전 보존 모드는 정확하지만 느립니다.
    // 기본값(false)은 이전 표시 행을 기본 배경으로 빠르게 되돌리는 속도 우선 모드입니다.
    this.restorePreviousSelectionHighlight_(sheet, props);

    const range = sheet.getRange(rowNo, 1, 1, width);
    const meta = {
      ssId: sheet.getParent().getId(),
      sheetId: String(sheet.getSheetId()),
      rowNo: rowNo,
      width: width
    };

    if (CONFIG.SELECTION.PRESERVE_ORIGINAL_BACKGROUNDS === true) {
      meta.backgrounds = range.getBackgrounds();
    }

    props.setProperty(CONFIG.SELECTION.USER_PROP_BG, JSON.stringify(meta));
    range.setBackground(CONFIG.SELECTION.HIGHLIGHT_COLOR || '#fff2cc');
    // 여기서 SpreadsheetApp.flush()를 매번 호출하면 체감이 더 느려질 수 있어 제거했습니다.
  }

  restorePreviousSelectionHighlight_(sheet, props) {
    try {
      const raw = props.getProperty(CONFIG.SELECTION.USER_PROP_BG);
      if (!raw) return;

      const saved = JSON.parse(raw);
      if (!saved) return;
      if (saved.ssId !== sheet.getParent().getId()) return;
      if (String(saved.sheetId) !== String(sheet.getSheetId())) return;
      if (!saved.rowNo || saved.rowNo < CONFIG.ROWS.MASTER_DATA_START) return;
      if (!saved.width) return;
      if (saved.rowNo > sheet.getMaxRows()) return;

      const range = sheet.getRange(saved.rowNo, 1, 1, saved.width);

      if (CONFIG.SELECTION.PRESERVE_ORIGINAL_BACKGROUNDS === true && saved.backgrounds) {
        range.setBackgrounds(saved.backgrounds);
      } else {
        range.setBackground(CONFIG.SELECTION.RESTORE_COLOR || null);
      }
    } catch (err) {
      Logger.log('이전 선택행 배경 복구 실패: ' + (err && err.stack || err));
    } finally {
      props.deleteProperty(CONFIG.SELECTION.USER_PROP_BG);
    }
  }

  clearMySelectionHighlight_() {
    const master = this.getMasterContext_();
    const props = PropertiesService.getUserProperties();
    this.restorePreviousSelectionHighlight_(master.sheet, props);
    props.deleteProperty(CONFIG.SELECTION.USER_PROP_ROW);
    props.deleteProperty(CONFIG.SELECTION.USER_PROP_SHEET_ID);
    props.deleteProperty(CONFIG.SELECTION.USER_PROP_SS_ID);
  }

  sendFromDialog(payload) {
    const rowNo = Number(payload && payload.rowNo);
    const mode = String(payload && payload.mode || '').toUpperCase();
    const testInput = String(payload && payload.testInput || '').trim();

    // v68 핵심:
    // HTML에서 어떤 값이 넘어오더라도 TEST 모드에서는 수동 수신/참조/삭제참조 payload를 서버 초입에서 폐기합니다.
    // 최종 수신자는 testInput 1건, 참조는 []로 buildRecipients_에서 다시 확정합니다.
    const manualTo = mode === 'TEST' ? [] : (payload && payload.manualTo);
    const manualCc = mode === 'TEST' ? [] : (payload && payload.manualCc);
    const removedCc = mode === 'TEST' ? [] : (payload && payload.removedCc);
    const runId = String(payload && payload.runId || Utilities.getUuid());
    const progress = new ProgressTracker(runId);
    const reviewSessionId = String(payload && payload.reviewSessionId || '').trim();
    let reviewPackage = null;

    if (!rowNo || rowNo < CONFIG.ROWS.MASTER_DATA_START) throw new Error('발송 행 정보가 올바르지 않습니다.');
    if (mode !== 'CUSTOMER' && mode !== 'TEST') throw new Error('발송 모드가 올바르지 않습니다.');

    const startedAt = new Date();
    let tempSpreadsheetFile = null;
    let registration = null;
    let targetData = null;
    let selectedDefs = [];

    try {
      progress.start('선택행 데이터 확인 중', 2);
      progress.throwIfCancelRequested_('선택행 확인 전');

      const master = this.getMasterContext_();
      const rowObj = this.readMasterRow_(master.sheet, master.headerMap, rowNo);
      assertPortalPayloadMatchesCurrentMasterRowV88_(payload, rowObj, '메일 발송');
      selectedDefs = this.getSelectedFileDefs_(rowObj);
      selectedDefs = this.applyDialogFileSelectionOverrides_(selectedDefs, payload);
      if (!selectedDefs.length) throw new Error('선택된 파일 체크값이 없습니다. 체크열을 확인하세요.');
      reviewPackage = this.getReviewPackage_(reviewSessionId, rowNo, selectedDefs, rowObj);

      progress.update(6, '마스터시트 상태 기록 중');
      this.updateMasterStatus_(master.sheet, master.headerMap, rowNo, {
        status: '처리중',
        message: '메일 자동발송 시작',
        sentAt: startedAt,
        flush: false
      });

      progress.update(10, '접수번호 확정 중');
      let requestRepo = null;
      if (reviewPackage && reviewPackage.requestNo) {
        registration = {
          requestNo: reviewPackage.requestNo,
          requestRowNo: Number(reviewPackage.requestRowNo) || 0
        };
        progress.update(14, '파일 확인/수정 접수번호 재사용: ' + registration.requestNo);
      } else if (reviewPackage && reviewPackage.requestRowNo) {
        requestRepo = new RequestRepository(CONFIG.GENERATOR_SPREADSHEET_ID, progress);
        registration = requestRepo.readRegistrationByRowNo(reviewPackage.requestRowNo);
        progress.update(14, '파일 확인/수정 접수번호 재사용: ' + registration.requestNo);
      } else {
        requestRepo = new RequestRepository(CONFIG.GENERATOR_SPREADSHEET_ID, progress);
        registration = requestRepo.registerFromMasterRow(rowObj);
      }
      progress.throwIfCancelRequested_('접수번호 확정 후');

      // v84 핵심:
      // 파일 확인/수정 후 발송은 이미 수정용 Drive 파일이 생성되어 있으므로,
      // 생성기 스프레드시트 복사/생성대상 로드/도장 삽입을 다시 하지 않습니다.
      // 수정용 파일을 바로 PDF/XLSX/DOCX로 export하고 메일에 붙여 30초 내 발송을 목표로 합니다.
      if (reviewPackage && reviewPackage.targetData) {
        progress.update(18, '수정본 기준 발송 준비 중');
        targetData = mergeObjectsPreferNonEmpty_(rowObj.toPlainObject(), reviewPackage.targetData || {});
        // v87: 회사명/고객번호/영업담당자/수신자 등 발송 핵심값은 항상 현재 마스터 행을 우선합니다.
        // 파일 확인 이후 마스터가 수정되었거나 세션 캐시가 남아 있어도 다른 고객명이 메일/첨부명에 섞이지 않게 합니다.
        targetData = enforceCurrentMasterIdentityOnTargetDataV87_(targetData, rowObj);
        targetData = applyContractPeriodDefaultsToObjectV90_(targetData);

        progress.update(28, '영업담당자/수신자 확인 중');
        const generatorSsForReviewSend = SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID);
        const sender = new SalesRepResolver(generatorSsForReviewSend).resolve(targetData['영업담당자']);
        const recipient = this.buildRecipients_(mode, targetData, sender, testInput, manualTo, manualCc, removedCc);

        progress.throwIfCancelRequested_('수정본 첨부 생성 전');
        progress.update(42, '수정본 첨부파일 생성 중');
        const attachments = new AttachmentBuilder(generatorSsForReviewSend, targetData, progress, reviewPackage).build(selectedDefs);

        if (!attachments.length) {
          Logger.log('실제 첨부파일 0개. 선택 항목은 메일 본문 Drive 링크 또는 본문 안내로 발송합니다: ' + selectedDefs.map(d => d.label).join(', '));
        }
        progress.throwIfCancelRequested_('수정본 첨부파일 생성 후');

        progress.update(76, '하이웍스 메일 구성 중');
        const mail = new MailMessage({
          from: sender.email,
          to: recipient.to,
          cc: recipient.cc,
          subject: this.buildMailSubject_(targetData, sender, selectedDefs),
          bodyHtml: this.buildMailBodyHtml_(targetData, sender, selectedDefs),
          attachments
        });

        progress.throwIfCancelRequested_('하이웍스 API 호출 직전');

        progress.update(84, '하이웍스 API로 첨부 메일 발송 중');
        const result = new HiworksMailer(progress).send(mail);

        progress.update(90, shouldDeferSentFileArchiveAfterSendV94_(mode, reviewPackage) ? '발송자료 공유드라이브 저장 예약 중' : '발송자료 공유드라이브 저장 중');
        const sentFileArchiveResult = archiveSentFilesOrEnqueueV94_({
          mode: mode,
          source: reviewSessionId ? '포털/파일확인후발송' : '시트/파일확인후발송',
          runId: runId,
          rowNo: rowNo,
          requestNo: registration.requestNo,
          targetData: targetData,
          sender: sender,
          recipient: recipient,
          selectedDefs: selectedDefs,
          mail: mail,
          attachments: attachments,
          hiworksResult: result,
          reviewPackage: reviewPackage
        }, progress);

        progress.update(94, '발송 결과 기록 중');
        this.updateMasterStatus_(master.sheet, master.headerMap, rowNo, {
          lastSent: selectedDefs.map(d => d.label).join(', '),
          incrementCount: true,
          status: mode === 'TEST' ? '테스트발송완료' : '발송완료',
          sentAt: new Date(),
          message: '하이웍스 응답: ' + JSON.stringify(result).slice(0, 420) + ' / 발송파일저장: ' + summarizeSentFileArchiveResult_(sentFileArchiveResult),
          requestNo: registration.requestNo,
          tempFolder: reviewPackage.folderUrl || '',
          flush: false
        });

        this.appendMailLog_([
          new Date(), 'SUCCESS', runId, rowNo, registration.requestNo,
          targetData['회사명'], recipient.to.join(','), recipient.cc.join(','),
          selectedDefs.map(d => d.label).join(', '), JSON.stringify(result)
        ]);

        progress.done('메일 발송 완료');

        return {
          ok: true,
          message: '메일 발송 완료',
          requestNo: registration.requestNo,
          to: recipient.to.join(', '),
          attachments: attachments.map(b => b.getName()),
          sentFileArchive: sentFileArchiveResult
        };
      }

      // 원본 누적 로그는 매우 짧은 lock 안에서 확정하고, 실제 작업은 복사본에서 수행합니다.
      progress.update(18, '임시 작업파일 복사 중');
      const work = new GeneratorWorkspace(registration, rowObj, progress);
      tempSpreadsheetFile = work.createTempCopy();

      progress.update(26, '생성대상 시트에 데이터 로드 중');
      targetData = work.loadTargetRow();

      // v21 핵심:
      // 생성대상/자동견적요청 시트에 없는 마스터 원행의 추가 세부정보(예: AX 계약 당사자, AY 사업자등록증상 법인 주소)를
      // 선임신고서에서 사용할 수 있도록 원본 마스터 행 전체를 targetData에 병합합니다.
      // 단, 생성대상에서 계산/정리된 값이 비어있지 않으면 그 값을 우선합니다.
      targetData = mergeObjectsPreferNonEmpty_(rowObj.toPlainObject(), targetData);
      targetData = applyContractPeriodToGeneratorSheetsV90_(work.ss, targetData, progress, '일반 발송 생성대상 병합 후');
      targetData = applyServiceApplicationApplicantFallbackToTarget_(work.ss, targetData);
      progress.throwIfCancelRequested_('생성대상 데이터 로드 후');

      progress.update(32, '영업담당자/수신자 확인 중');
      const sender = new SalesRepResolver(work.ss).resolve(targetData['영업담당자']);
      const recipient = this.buildRecipients_(mode, targetData, sender, testInput, manualTo, manualCc, removedCc);

      const imagePlan = this.resolveRequiredImagePlan_(selectedDefs);
      this.applyRequiredImages_(work, targetData, imagePlan, progress, {
        cachedMessage: '도장/로고 사전삽입 캐시 적용 완료',
        staticMessage: '도장/로고 삽입 중',
        dynamicMessage: '비교견적서 수행사 도장 삽입 중',
        noneMessage: '도장/로고 삽입 생략',
        waitMs: CONFIG.EXPORT.WAIT_MS_AFTER_IMAGE_INSERT
      });
      progress.throwIfCancelRequested_('이미지 삽입 후');
      progress.update(50, '첨부파일/Drive 링크 생성 중');
      const attachments = new AttachmentBuilder(work.ss, targetData, progress, reviewPackage).build(selectedDefs);

      // v64:
      // 수행사정보+샘플보고서 동시 체크처럼 대용량 자료를 첨부 대신 Google Drive 링크로 전환하는 케이스는
      // 실제 API 첨부파일이 0개가 정상입니다. 선택 파일 정의(selectedDefs)가 이미 1개 이상 확인됐으면
      // 첨부파일 개수만으로 발송을 중단하지 않습니다.
      if (!attachments.length) {
        Logger.log('실제 첨부파일 0개. 선택 항목은 메일 본문 Drive 링크 또는 본문 안내로 발송합니다: ' + selectedDefs.map(d => d.label).join(', '));
      }
      progress.throwIfCancelRequested_('첨부파일/Drive 링크 생성 후');

      progress.update(76, '하이웍스 메일 구성 중');
      const mail = new MailMessage({
        from: sender.email,
        to: recipient.to,
        cc: recipient.cc,
        subject: this.buildMailSubject_(targetData, sender, selectedDefs),
        bodyHtml: this.buildMailBodyHtml_(targetData, sender, selectedDefs),
        attachments
      });

      progress.throwIfCancelRequested_('하이웍스 API 호출 직전');

      progress.update(84, '하이웍스 API로 첨부 메일 발송 중');
      const result = new HiworksMailer(progress).send(mail);

      progress.update(90, shouldDeferSentFileArchiveAfterSendV94_(mode, reviewPackage) ? '발송자료 공유드라이브 저장 예약 중' : '발송자료 공유드라이브 저장 중');
      const sentFileArchiveResult = archiveSentFilesOrEnqueueV94_({
        mode: mode,
        source: reviewSessionId ? '포털/일반발송' : '시트/일반발송',
        runId: runId,
        rowNo: rowNo,
        requestNo: registration.requestNo,
        targetData: targetData,
        sender: sender,
        recipient: recipient,
        selectedDefs: selectedDefs,
        mail: mail,
        attachments: attachments,
        hiworksResult: result,
        reviewPackage: reviewPackage
      }, progress);

      progress.update(94, '발송 결과 기록 중');
      this.updateMasterStatus_(master.sheet, master.headerMap, rowNo, {
        lastSent: selectedDefs.map(d => d.label).join(', '),
        incrementCount: true,
        status: mode === 'TEST' ? '테스트발송완료' : '발송완료',
        sentAt: new Date(),
        message: '하이웍스 응답: ' + JSON.stringify(result).slice(0, 420) + ' / 발송파일저장: ' + summarizeSentFileArchiveResult_(sentFileArchiveResult),
        requestNo: registration.requestNo,
        tempFolder: tempSpreadsheetFile ? tempSpreadsheetFile.getUrl() : '',
        flush: false
      });

      this.appendMailLog_([
        new Date(), 'SUCCESS', runId, rowNo, registration.requestNo,
        targetData['회사명'], recipient.to.join(','), recipient.cc.join(','),
        selectedDefs.map(d => d.label).join(', '), JSON.stringify(result)
      ]);

      if (CONFIG.TEMP.TRASH_TEMP_SPREADSHEET_AFTER_SUCCESS && tempSpreadsheetFile) {
        progress.update(97, '임시 작업파일 정리 중');
        tempSpreadsheetFile.setTrashed(true);
      }

      progress.done('메일 발송 완료');

      return {
        ok: true,
        message: '메일 발송 완료',
        requestNo: registration.requestNo,
        to: recipient.to.join(', '),
        attachments: attachments.map(b => b.getName()),
        sentFileArchive: sentFileArchiveResult
      };
    } catch (err) {
      progress.fail(String(err && err.message || err));
      try {
        const master = this.getMasterContext_();
        const errText = String(err && err.message || err);
        const isCancelled = errText.indexOf('메일 발송이 취소되었습니다') >= 0;

        this.updateMasterStatus_(master.sheet, master.headerMap, rowNo, {
          status: isCancelled ? '발송취소' : '발송실패',
          sentAt: new Date(),
          message: String(err && err.stack || err),
          flush: false
        });
        this.appendMailLog_([new Date(), 'FAIL', runId, rowNo, registration ? registration.requestNo : '', targetData ? targetData['회사명'] : '', '', '', selectedDefs.map(d => d.label).join(', '), String(err && err.stack || err)]);
      } catch (logErr) {
        // 로그 실패는 원래 오류를 덮지 않습니다.
      }
      throw err;
    }
  }



  prepareFilesForReview(payload) {
    const rowNo = Number(payload && payload.rowNo);
    const runId = String(payload && payload.runId || Utilities.getUuid());
    const progress = new ProgressTracker(runId);

    if (!rowNo || rowNo < CONFIG.ROWS.MASTER_DATA_START) throw new Error('파일 확인/수정 대상 행 정보가 올바르지 않습니다.');

    let tempSpreadsheetFile = null;
    let registration = null;
    let selectedDefs = [];

    try {
      progress.start('파일 확인/수정 준비 중', 2);
      const master = this.getMasterContext_();
      const rowObj = this.readMasterRow_(master.sheet, master.headerMap, rowNo);
      assertPortalPayloadMatchesCurrentMasterRowV88_(payload, rowObj, '파일 확인/수정');
      selectedDefs = this.getSelectedFileDefs_(rowObj);
      selectedDefs = this.applyDialogFileSelectionOverrides_(selectedDefs, payload);
      if (!selectedDefs.length) throw new Error('체크열에서 생성할 파일을 하나 이상 체크하세요.');

      // v83:
      // [파일 확인/수정]은 사용자가 실제로 수정해야 하는 문서만 Drive 폴더에 생성합니다.
      // 안내문/수행사정보/샘플보고서/수행사 표준계약서 등 비편집 자료는 여기서 만들지 않고,
      // 최종 발송 시 기존 원본 기준으로 첨부/링크 처리합니다.
      const reviewSelectedDefs = this.filterEditableReviewFileDefs_(selectedDefs);
      if (!reviewSelectedDefs.length) {
        throw new Error('파일 확인/수정 가능한 문서가 선택되어 있지 않습니다. 견적서, 용역신청서, 선임신고서 및 위임장, 비교견적서 중 하나를 체크해 주세요.');
      }

      progress.update(8, '접수번호 확정 중');
      const requestRepo = new RequestRepository(CONFIG.GENERATOR_SPREADSHEET_ID, progress);
      registration = requestRepo.registerFromMasterRow(rowObj);

      progress.update(18, '임시 작업파일 복사 중');
      const work = new GeneratorWorkspace(registration, rowObj, progress);
      tempSpreadsheetFile = work.createTempCopy();

      progress.update(28, '생성대상 시트에 데이터 로드 중');
      let targetData = work.loadTargetRow();
      targetData = mergeObjectsPreferNonEmpty_(rowObj.toPlainObject(), targetData);
      targetData = applyContractPeriodToGeneratorSheetsV90_(work.ss, targetData, progress, '파일 확인/수정 생성대상 병합 후');
      targetData = applyServiceApplicationApplicantFallbackToTarget_(work.ss, targetData);

      const imagePlan = this.resolveRequiredImagePlan_(selectedDefs);
      this.applyRequiredImages_(work, targetData, imagePlan, progress, {
        cachedMessage: '도장/로고 사전삽입 캐시 적용 완료',
        staticMessage: '도장/로고 삽입 중',
        dynamicMessage: '비교견적서 수행사 도장 삽입 중',
        noneMessage: '도장/로고 삽입 생략',
        // v73/v75: 파일 확인/수정 단계에서는 바로 수정용 Google Sheets/Docs/PDF를 만들기 때문에
        // PDF/XLSX export 직전처럼 오래 대기할 필요가 없습니다. 이미지 반영 최소 대기만 둡니다.
        waitMs: Math.min(250, CONFIG.EXPORT.WAIT_MS_AFTER_IMAGE_INSERT || 250)
      });

      progress.update(55, 'Drive 수정용 폴더 생성 중');
      const packageInfo = new ReviewFilePackageBuilder(work.ss, targetData, progress).build(reviewSelectedDefs, registration);
      const sessionId = 'review_' + Utilities.getUuid();
      const reviewPackage = {
        sessionId: sessionId,
        rowNo: rowNo,
        requestNo: registration.requestNo,
        requestRowNo: registration.requestRowNo,
        // selectedKeys는 전체 발송 선택값 기준으로 저장합니다.
        // 그래야 파일 확인/수정 이후 안내문/샘플보고서 같은 비편집 항목이 함께 체크되어 있어도
        // 발송 단계에서 세션 선택값 검증이 어긋나지 않습니다.
        selectedKeys: selectedDefs.map(function(def) { return def.key; }),
        reviewSelectedKeys: reviewSelectedDefs.map(function(def) { return def.key; }),
        // v84: 파일 확인/수정 후 발송 시 생성기 복사/생성대상 재로딩을 생략하기 위한 발송용 데이터 스냅샷입니다.
        targetData: targetData,
        folderId: packageInfo.folderId,
        folderUrl: packageInfo.folderUrl,
        files: packageInfo.files,
        createdAt: new Date().toISOString()
      };

      this.storeReviewPackage_(reviewPackage);

      this.updateMasterStatus_(master.sheet, master.headerMap, rowNo, {
        status: '파일확인중',
        sentAt: new Date(),
        message: '파일 확인/수정용 Drive 폴더 생성: ' + packageInfo.folderUrl,
        requestNo: registration.requestNo,
        tempFolder: packageInfo.folderUrl,
        flush: false
      });

      if (CONFIG.TEMP.TRASH_TEMP_SPREADSHEET_AFTER_SUCCESS && tempSpreadsheetFile) {
        progress.update(94, '임시 작업파일 정리 중');
        tempSpreadsheetFile.setTrashed(true);
      }

      progress.done('파일 확인/수정용 Drive 폴더 생성 완료');

      return {
        ok: true,
        message: '파일 확인/수정용 Drive 폴더 생성 완료',
        reviewSessionId: sessionId,
        requestNo: registration.requestNo,
        folderUrl: packageInfo.folderUrl,
        fileCount: packageInfo.files.length,
        files: packageInfo.files.map(function(file) {
          return { key: file.key, label: file.label, name: file.name, url: file.url, useForSend: file.useForSend !== false };
        })
      };
    } catch (err) {
      progress.fail(String(err && err.message || err));
      try {
        const master = this.getMasterContext_();
        this.updateMasterStatus_(master.sheet, master.headerMap, rowNo, {
          status: '파일확인실패',
          sentAt: new Date(),
          message: String(err && err.stack || err),
          flush: false
        });
      } catch (logErr) {}
      throw err;
    }
  }

  storeReviewPackage_(reviewPackage) {
    if (!reviewPackage || !reviewPackage.sessionId) return;
    const key = 'MAILAUTO_REVIEW_PACKAGE_' + reviewPackage.sessionId;
    const text = JSON.stringify(reviewPackage);
    CacheService.getUserCache().put(key, text, 21600);
    CacheService.getScriptCache().put(key, text, 21600);
  }

  getReviewPackage_(sessionId, rowNo, selectedDefs, rowObj) {
    const id = String(sessionId || '').trim();
    if (!id) return null;

    const key = 'MAILAUTO_REVIEW_PACKAGE_' + id;
    const raw = CacheService.getUserCache().get(key) || CacheService.getScriptCache().get(key);
    if (!raw) throw new Error('파일 확인/수정 세션이 만료되었습니다. [파일 확인/수정]을 다시 실행해 주세요.');

    let pkg;
    try {
      pkg = JSON.parse(raw);
    } catch (err) {
      throw new Error('파일 확인/수정 세션 정보를 읽지 못했습니다. [파일 확인/수정]을 다시 실행해 주세요.');
    }

    if (Number(pkg.rowNo) !== Number(rowNo)) {
      throw new Error('파일 확인/수정 세션의 행과 현재 발송 행이 다릅니다. [파일 확인/수정]을 다시 실행해 주세요.');
    }

    const expected = (selectedDefs || []).map(function(def) { return def && def.key; }).filter(Boolean).sort().join('|');
    const saved = (pkg.selectedKeys || []).filter(Boolean).sort().join('|');
    if (expected !== saved) {
      throw new Error('체크박스 선택값이 파일 확인/수정 당시와 달라졌습니다. [파일 확인/수정]을 다시 실행해 주세요.');
    }

    // v87: 세션 ID가 브라우저/캐시에 남아 있더라도 다른 고객 자료에 재사용되지 않게
    // 현재 마스터 행의 고객번호/회사명과 파일 확인 당시 스냅샷을 한 번 더 검증합니다.
    assertReviewPackageMatchesCurrentRowV87_(pkg, rowObj);

    return pkg;
  }

  buildMailSubject_(targetData, sender, selectedDefs) {
    const subjectCfg = CONFIG.MAIL.SUBJECT || {};
    const extra = this.buildMailTemplateValues_(targetData, sender, selectedDefs);

    const initialKey = subjectCfg.INITIAL_TRIGGER_KEY || 'termsGuide';
    if (this.isSelectedFileKey_(selectedDefs, initialKey)) {
      const template = subjectCfg.INITIAL_TEMPLATE || CONFIG.MAIL.SUBJECT_TEMPLATE || '';
      return this.renderTemplate_(template, targetData, sender, extra);
    }

    const labels = this.buildSubjectFileTypeLabels_(selectedDefs, subjectCfg);
    const subjectFileTypes = labels.length ? labels.join(', ') : '요청자료';
    const template = labels.length <= 1
      ? (subjectCfg.RESEND_TEMPLATE_SINGLE || subjectCfg.RESEND_TEMPLATE || CONFIG.MAIL.SUBJECT_TEMPLATE || '')
      : (subjectCfg.RESEND_TEMPLATE_MULTI || subjectCfg.RESEND_TEMPLATE || CONFIG.MAIL.SUBJECT_TEMPLATE || '');

    return this.renderTemplate_(template, targetData, sender, Object.assign({}, extra, {
      subjectFileTypes: subjectFileTypes,
      selectedSubjectFileTypes: subjectFileTypes
    }));
  }

  buildSubjectFileTypeLabels_(selectedDefs, subjectCfg) {
    const defs = selectedDefs || [];
    const map = (subjectCfg && subjectCfg.FILE_TYPE_TEXT) || {};
    const order = (subjectCfg && subjectCfg.FILE_TYPE_ORDER) || [];
    const byKey = {};

    defs.forEach(def => {
      if (!def || !def.key) return;
      byKey[def.key] = def;
    });

    const orderedKeys = [];
    order.forEach(key => {
      if (byKey[key] && orderedKeys.indexOf(key) < 0) orderedKeys.push(key);
    });
    defs.forEach(def => {
      if (def && def.key && orderedKeys.indexOf(def.key) < 0) orderedKeys.push(def.key);
    });

    return orderedKeys
      .map(key => map[key] || (byKey[key] && byKey[key].label) || '')
      .map(v => String(v || '').trim())
      .filter(Boolean);
  }

  buildMailBodyHtml_(targetData, sender, selectedDefs) {
    const extra = this.buildMailTemplateValues_(targetData, sender, selectedDefs);
    const parts = [];

    if (CONFIG.MAIL.BODY_HEADER_HTML) {
      parts.push(CONFIG.MAIL.BODY_HEADER_HTML);
    } else if (CONFIG.MAIL.BODY_INTRO_HTML) {
      // v11 구버전 호환
      parts.push(CONFIG.MAIL.BODY_INTRO_HTML);
    }

    const requestHtml = this.buildSelectedRequestBodyHtml_(targetData, sender, selectedDefs, extra);
    if (requestHtml) {
      parts.push(requestHtml);
    } else {
      // BODY_REQUEST_HTML을 비워두면 v11처럼 체크박스 key별 HTML을 순서대로 붙입니다.
      const caseMap = CONFIG.MAIL.CASE_BODY_HTML || {};
      const used = {};
      (selectedDefs || []).forEach(def => {
        if (!def || !def.key || used[def.key]) return;
        used[def.key] = true;
        const caseHtml = caseMap[def.key];
        if (caseHtml) parts.push(caseHtml);
      });
    }

    const largeAttachmentLinksHtml = this.buildLargeAttachmentLinksHtml_(targetData);
    if (largeAttachmentLinksHtml) {
      parts.push(largeAttachmentLinksHtml);
    }

    if (CONFIG.MAIL.BODY_COMMON_HTML) {
      parts.push(CONFIG.MAIL.BODY_COMMON_HTML);
    }

    if (CONFIG.MAIL.BODY_SIGNATURE_HTML) {
      parts.push(CONFIG.MAIL.BODY_SIGNATURE_HTML);
    }

    const inlineImageHtml = this.buildMailInlineImagesHtml_(targetData, sender, selectedDefs);
    if (inlineImageHtml) {
      parts.push(inlineImageHtml);
    }

    const html = parts.join('');
    return this.renderTemplate_(html || CONFIG.MAIL.BODY_TEMPLATE_HTML || '', targetData, sender, extra);
  }

  buildLargeAttachmentLinksHtml_(targetData) {
    const links = this.normalizeLargeAttachmentLinks_(targetData && targetData.__MAIL_LARGE_ATTACHMENT_LINKS__);
    if (!links.length) return '';

    // v62:
    // 상단 전체 다운로드 기능은 제거합니다.
    // 대용량 파일은 아래 개별 파일의 미리보기/다운로드 링크로만 제공합니다.

    const rows = links.map(item => {
      const name = escapeHtml_(item.name || '자료 파일');
      const source = escapeHtml_(item.source || '자료');
      const size = item.sizeText ? escapeHtml_(item.sizeText) : '';
      const previewUrl = escapeHtml_(item.url || '');
      const downloadUrl = escapeHtml_(item.downloadUrl || item.url || '');
      const subText = size ? source + ' · ' + size : source;
      const fileIconHtml = this.buildDriveFileIconHtml_();

      return '' +
        '<tr>' +
          '<td style="padding:0;border-top:1px solid #e7eaf0;">' +
            '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:collapse;width:100%;background:#ffffff;">' +
              '<tr>' +
                '<td width="58" valign="middle" style="padding:14px 12px 14px 18px;width:58px;">' +
                  fileIconHtml +
                '</td>' +
                '<td valign="middle" style="padding:14px 8px 14px 0;">' +
                  '<a href="' + downloadUrl + '" target="_blank" style="font-size:15px;font-weight:700;color:#202124;line-height:1.45;word-break:break-all;text-decoration:none;">' + name + '</a>' +
                  '<div style="font-size:12px;color:#6f7682;line-height:1.45;margin-top:2px;">' + subText + '</div>' +
                '</td>' +
                '<td valign="middle" align="right" style="padding:14px 18px 14px 8px;white-space:nowrap;">' +
                  '<a href="' + previewUrl + '" target="_blank" style="color:#5f6368;text-decoration:none;font-size:13px;margin-right:14px;">미리보기</a>' +
                  '<a href="' + downloadUrl + '" target="_blank" style="color:#1a73e8;text-decoration:none;font-size:13px;font-weight:700;">다운로드</a>' +
                '</td>' +
              '</tr>' +
            '</table>' +
          '</td>' +
        '</tr>';
    }).join('');

    return '' +
      '<div style="font-family:\'Malgun Gothic\', \'맑은 고딕\', Arial, sans-serif; font-size:15px; line-height:1.7; margin:0 0 18px 0;">' +
        '<table role="presentation" width="100%" cellspacing="0" cellpadding="0" style="border-collapse:separate;border-spacing:0;width:100%;border:1px solid #d9dde3;border-radius:12px;background:#ffffff;overflow:hidden;">' +
          '<tr>' +
            '<td style="padding:18px 20px 10px 20px;background:#f8fafd;">' +
              '<div style="font-size:17px;font-weight:700;color:#202124;letter-spacing:-0.02em;">※ 수행사 정보 및 샘플보고서 다운로드 안내</div>' +
              '<div style="margin-top:10px;font-size:14px;color:#3c4043;line-height:1.75;">' +
                '수행사 정보와 샘플보고서는 파일 용량이 커서 메일 첨부 대신 아래 Google Drive 링크로 전달드립니다.<br>' +
                '파일명 또는 다운로드 버튼을 클릭하시면 바로 다운로드가 실행됩니다.' +
              '</div>' +
            '</td>' +
          '</tr>' +
          rows +
          '<tr>' +
            '<td style="padding:12px 20px 16px 20px;background:#fbfbfc;border-top:1px solid #e7eaf0;color:#7a7f88;font-size:12px;line-height:1.6;">' +
              '브라우저 설정에 따라 다운로드 위치 선택창이 표시될 수 있습니다. 링크가 열리지 않을 경우 담당자에게 회신 부탁드립니다.' +
            '</td>' +
          '</tr>' +
        '</table>' +
      '</div>';
  }

  buildDriveFileIconHtml_() {
    // 메일 클라이언트 호환성을 위해 외부 이미지 대신 단순 파일 아이콘을 HTML/CSS로 렌더링합니다.
    return '' +
      '<div style="width:40px;height:40px;border:1px solid #d9dde3;border-radius:8px;background:#f7f9fc;text-align:center;line-height:40px;color:#5f6368;font-size:22px;">' +
        '&#128206;' +
      '</div>';
  }


  prepareSequentialDownloadLauncherForLinks_(targetData, links) {
    if (!links || !links.length) return null;

    if (targetData && targetData.__MAIL_MULTI_DOWNLOAD_LAUNCHER__) {
      return targetData.__MAIL_MULTI_DOWNLOAD_LAUNCHER__;
    }

    const webAppUrl = this.getMailAutoDownloadWebAppUrl_();
    if (!webAppUrl) {
      Logger.log(
        '메일자동화 전체 다운로드 링크 생성 실패: 웹앱 URL을 찾지 못했습니다. ' +
        'Apps Script를 웹 앱으로 배포하거나 ScriptProperties의 MAILAUTO_DOWNLOAD_WEBAPP_URL에 웹앱 URL을 저장하세요.'
      );
      return null;
    }

    const files = links
      .map(item => {
        if (!item) return null;
        const downloadUrl = String(item.downloadUrl || item.url || '').trim();
        if (!downloadUrl) return null;

        return {
          name: String(item.name || '자료 파일'),
          source: String(item.source || '자료'),
          sizeText: String(item.sizeText || ''),
          url: String(item.url || downloadUrl),
          downloadUrl: downloadUrl
        };
      })
      .filter(Boolean);

    if (!files.length) return null;

    const token = Utilities.getUuid().replace(/-/g, '') + String(Date.now());
    const expiresAt = Date.now() + 1000 * 60 * 60 * 24 * 30; // 30일
    const payload = {
      token: token,
      createdAt: Date.now(),
      expiresAt: expiresAt,
      files: files
    };

    PropertiesService.getScriptProperties().setProperty(
      this.getMailAutoDownloadTokenKey_(token),
      JSON.stringify(payload)
    );

    const sep = webAppUrl.indexOf('?') >= 0 ? '&' : '?';
    const url = webAppUrl + sep + 'action=mailAutoMultiDownload&token=' + encodeURIComponent(token);
    const info = {
      url: url,
      token: token,
      count: files.length
    };

    if (targetData) targetData.__MAIL_MULTI_DOWNLOAD_LAUNCHER__ = info;
    Logger.log('메일자동화 전체 다운로드 실행 링크 생성: 파일 ' + files.length + '개 / ' + url);
    return info;
  }

  getMailAutoDownloadWebAppUrl_() {
    const props = PropertiesService.getScriptProperties();
    const configured = String(props.getProperty('MAILAUTO_DOWNLOAD_WEBAPP_URL') || '').trim();
    if (configured) return configured;

    try {
      const service = ScriptApp.getService && ScriptApp.getService();
      const url = service && service.getUrl ? String(service.getUrl() || '').trim() : '';
      if (url) return url;
    } catch (err) {
      Logger.log('ScriptApp.getService().getUrl() 확인 실패: ' + (err && err.stack || err));
    }

    return '';
  }

  getMailAutoDownloadTokenKey_(token) {
    return 'MAILAUTO_MULTI_DOWNLOAD_' + String(token || '').trim();
  }

  prepareLargeAttachmentZipForLinks_(targetData, links) {
    if (!links || !links.length) return null;

    if (targetData && targetData.__MAIL_LARGE_ATTACHMENT_ZIP__) {
      return targetData.__MAIL_LARGE_ATTACHMENT_ZIP__;
    }

    try {
      const blobs = [];
      const seenNames = {};

      links.forEach((item, idx) => {
        if (!item || !item.fileId) return;

        const file = DriveApp.getFileById(String(item.fileId));
        const entryName = this.makeUniqueZipEntryName_(item.name || file.getName() || ('자료_' + (idx + 1)), seenNames);
        const blob = file.getBlob();
        blob.setName(entryName);
        blobs.push(blob);
      });

      if (!blobs.length) return null;

      const vendor = sanitizeFileName_((targetData && targetData['수행사']) || '수행사');
      const zipName = sanitizeFileName_(vendor + '_수행사정보_및_샘플보고서_' + getTodayFileDateSuffix_() + '.zip');
      const zipBlob = Utilities.zip(blobs, zipName);
      const parentFolder = CONFIG.TEMP && CONFIG.TEMP.PARENT_FOLDER_ID ? DriveApp.getFolderById(CONFIG.TEMP.PARENT_FOLDER_ID) : null;
      const zipFile = parentFolder ? parentFolder.createFile(zipBlob) : DriveApp.createFile(zipBlob);

      if (CONFIG.HIWORKS && CONFIG.HIWORKS.MAKE_DRIVE_LINK_FILES_VIEWABLE_BY_LINK === true) {
        try {
          zipFile.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
        } catch (shareErr) {
          Logger.log('전체 다운로드 ZIP 링크공유 설정 실패. 기존 권한으로 링크만 삽입합니다: ' + zipName + ' / ' + shareErr);
        }
      }

      const info = {
        name: zipFile.getName(),
        sizeBytes: zipFile.getSize ? Number(zipFile.getSize()) || 0 : 0,
        sizeText: zipFile.getSize ? this.formatLinkBytes_(Number(zipFile.getSize()) || 0) : '',
        url: this.getViewableDriveFileUrl_(zipFile),
        downloadUrl: this.getDriveFileDownloadUrl_(zipFile),
        source: '전체 ZIP'
      };

      if (targetData) targetData.__MAIL_LARGE_ATTACHMENT_ZIP__ = info;
      Logger.log('수행사정보/샘플보고서 전체 다운로드 ZIP 생성: ' + info.name + ' / ' + info.downloadUrl);
      return info;
    } catch (err) {
      Logger.log('수행사정보/샘플보고서 전체 ZIP 생성 실패. 개별 링크만 표시합니다: ' + (err && err.stack || err));
      return null;
    }
  }

  makeUniqueZipEntryName_(name, seenNames) {
    let clean = sanitizeFileName_(String(name || '자료파일').trim()) || '자료파일';
    const key = clean.toLowerCase();
    if (!seenNames[key]) {
      seenNames[key] = 1;
      return clean;
    }

    let dot = clean.lastIndexOf('.');
    let base = dot > 0 ? clean.slice(0, dot) : clean;
    let ext = dot > 0 ? clean.slice(dot) : '';
    let n = ++seenNames[key];
    let candidate = base + '_' + n + ext;

    while (seenNames[candidate.toLowerCase()]) {
      n += 1;
      candidate = base + '_' + n + ext;
    }

    seenNames[candidate.toLowerCase()] = 1;
    return candidate;
  }

  formatLinkBytes_(bytes) {
    const n = Number(bytes) || 0;
    if (n >= 1024 * 1024 * 1024) return (n / 1024 / 1024 / 1024).toFixed(2).replace(/\.00$/, '') + ' GB';
    if (n >= 1024 * 1024) return (n / 1024 / 1024).toFixed(2).replace(/\.00$/, '') + ' MB';
    if (n >= 1024) return (n / 1024).toFixed(1).replace(/\.0$/, '') + ' KB';
    return n + ' B';
  }

  normalizeLargeAttachmentLinks_(value) {
    if (!value) return [];
    if (Array.isArray(value)) return value.filter(Boolean);
    if (typeof value === 'string') {
      try {
        const parsed = JSON.parse(value);
        return Array.isArray(parsed) ? parsed.filter(Boolean) : [];
      } catch (err) {
        return [];
      }
    }
    return [];
  }

  buildMailInlineImagesHtml_(targetData, sender, selectedDefs) {
    const cfg = CONFIG.MAIL.INLINE_IMAGE || {};
    if (!cfg.ENABLED) return '';

    const html = [];
    const cardFileId = this.getBusinessCardFileId_(sender);
    if (cardFileId) {
      html.push(this.buildInlineDriveImageHtml_(cardFileId, cfg.BUSINESS_CARD_WIDTH_PX || 280, '영업담당자 명함'));
    }

    if (this.isSelectedFileKey_(selectedDefs, 'termsGuide')) {
      const guideIds = cfg.TERMS_GUIDE_IMAGE_FILE_IDS || [];
      guideIds.forEach((fileId, idx) => {
        if (!fileId) return;
        html.push(this.buildInlineDriveImageHtml_(fileId, cfg.GUIDE_WIDTH_PX || 620, '안내문 ' + (idx + 1)));
      });
    }

    if (!html.length) return '';
    const gap = cfg.GAP_HTML || '<br>';
    return '<div style="margin-top:8px;line-height:1.4;">' + html.join(gap) + '</div>';
  }

  getBusinessCardFileId_(sender) {
    const cfg = CONFIG.MAIL.INLINE_IMAGE || {};
    const map = cfg.BUSINESS_CARD_FILE_IDS || {};
    const rawName = String(sender && sender.name || '').trim();
    if (!rawName) return '';
    if (map[rawName]) return map[rawName];

    const normalized = normalizeHeader_(rawName);
    const foundKey = Object.keys(map).find(name => normalizeHeader_(name) === normalized);
    return foundKey ? map[foundKey] : '';
  }

  isSelectedFileKey_(selectedDefs, key) {
    return (selectedDefs || []).some(def => def && def.key === key);
  }

  buildInlineDriveImageHtml_(fileId, widthPx, altText) {
    const id = String(fileId || '').trim();
    if (!id) return '';
    this.ensureInlineImageViewable_(id);
    const width = Math.max(120, Number(widthPx) || 300);
    const src = 'https://drive.google.com/thumbnail?id=' + encodeURIComponent(id) + '&sz=w' + Math.ceil(width * 2);
    return '<img src="' + src + '" alt="' + escapeHtml_(altText || '') + '" ' +
      'style="display:block;width:' + width + 'px;max-width:100%;height:auto;border:0;margin:4px 0;">';
  }

  ensureInlineImageViewable_(fileId) {
    const cfg = CONFIG.MAIL.INLINE_IMAGE || {};
    if (!cfg.MAKE_VIEWABLE_BY_LINK) return;
    try {
      const cache = CacheService.getScriptCache();
      const cacheKey = 'MAILAUTO_INLINE_IMG_SHARED_' + fileId;
      if (cache.get(cacheKey)) return;
      DriveApp.getFileById(fileId).setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      cache.put(cacheKey, '1', 21600);
    } catch (err) {
      // 이미지 공유 실패만으로 메일 발송을 막지는 않습니다.
      // 단, 수신자 환경에서 본문 이미지가 보이지 않을 수 있습니다.
      Logger.log('본문 이미지 공유 설정 실패: ' + fileId + ' / ' + err);
    }
  }

  buildSelectedRequestBodyHtml_(targetData, sender, selectedDefs, extra) {
    if (!CONFIG.MAIL.BODY_REQUEST_HTML) return '';
    const labelMap = CONFIG.MAIL.CASE_BODY_TEXT || {};
    const labels = [];
    const used = {};

    const defs = selectedDefs || [];
    const bodyOrder = CONFIG.MAIL.BODY_FILE_TYPE_ORDER || [];
    const byKey = {};

    defs.forEach(def => {
      if (def && def.key) byKey[def.key] = def;
    });

    const orderedKeys = [];
    bodyOrder.forEach(key => {
      if (byKey[key] && orderedKeys.indexOf(key) < 0) orderedKeys.push(key);
    });
    defs.forEach(def => {
      if (def && def.key && orderedKeys.indexOf(def.key) < 0) orderedKeys.push(def.key);
    });

    orderedKeys.forEach(key => {
      if (used[key]) return;
      used[key] = true;
      const def = byKey[key];
      const raw = labelMap[key] || (def && def.label) || '';
      if (!raw) return;
      labels.push(this.renderTemplate_(raw, targetData, sender, extra));
    });

    if (!labels.length) return '';

    const selectedRequestText = joinKoreanList_(labels);
    return this.renderTemplate_(CONFIG.MAIL.BODY_REQUEST_HTML, targetData, sender, Object.assign({}, extra, {
      selectedRequestText: selectedRequestText
    }));
  }

  buildMailTemplateValues_(targetData, sender, selectedDefs) {
    const customerContact = cleanContactText_(targetData['고객사 담당자'] || targetData['담당자 이름'] || targetData['담당자'] || '');
    const salesRepDisplay = buildPersonTitleText_(sender.name || targetData['영업담당자'] || '', sender.title || '');
    const vendorName = String(targetData['수행사'] || '').trim();
    return {
      customerContact: customerContact,
      customerContactNim: customerContact ? customerContact + '님' : '',
      salesRepName: sender.name || targetData['영업담당자'] || '',
      salesRepTitle: sender.title || '',
      salesRepDisplay: salesRepDisplay,
      salesRepPhone: sender.phone || '',
      salesRepEmail: sender.email || '',
      vendorDisplayName: getVendorDisplayName_(vendorName),
      selectedFileCount: selectedDefs ? selectedDefs.length : 0
    };
  }

  resolveRequiredSourceSheets_(selectedDefs) {
    const names = [];
    (selectedDefs || []).forEach(def => {
      if (!def) return;
      if (def.type === 'sheet_pdf' || def.type === 'sheet_xlsx_values') {
        if (def.sheetName) names.push(def.sheetName);
      }
      if (def.type === 'multi_sheet_pdf' && Array.isArray(def.sheets)) {
        def.sheets.forEach(item => item && item.sheetName && names.push(item.sheetName));
      }
    });
    return unique_(names);
  }

  resolveRequiredImageSheets_(selectedDefs) {
    // 구버전 호환용: 정적 도장/로고 대상 시트만 반환합니다.
    return this.resolveRequiredImagePlan_(selectedDefs).staticSheetNames;
  }

  resolveRequiredImagePlan_(selectedDefs) {
    const sourceSheetNames = this.resolveRequiredSourceSheets_(selectedDefs);
    const staticImageSheets = unique_((CONFIG.IMAGE_PLACEMENTS || []).map(p => p.sheetName));
    const dynamicImageSheets = unique_((CONFIG.DYNAMIC_STAMP_PLACEMENTS || []).map(p => p.sheetName));

    return {
      sourceSheetNames: sourceSheetNames,
      staticSheetNames: sourceSheetNames.filter(name => staticImageSheets.indexOf(name) >= 0),
      dynamicSheetNames: sourceSheetNames.filter(name => dynamicImageSheets.indexOf(name) >= 0)
    };
  }

  applyRequiredImages_(work, targetData, imagePlan, progress, options) {
    const plan = imagePlan || { staticSheetNames: [], dynamicSheetNames: [] };
    const opts = options || {};
    const staticSheetNames = unique_(plan.staticSheetNames || []);
    const dynamicSheetNames = unique_(plan.dynamicSheetNames || []);
    let didInsert = false;

    if (staticSheetNames.length && work.isPrestamped) {
      if (progress) progress.update(42, opts.cachedMessage || '도장/로고 사전삽입 캐시 적용 완료');
    } else if (staticSheetNames.length) {
      if (progress) progress.update(40, opts.staticMessage || '도장/로고 삽입 중');
      new ImagePlacer(work.ss, targetData['수행사']).placeForSheets(staticSheetNames);
      didInsert = true;
    }

    if (dynamicSheetNames.length) {
      if (progress) progress.update(staticSheetNames.length ? 44 : 42, opts.dynamicMessage || '동적 도장 삽입 중');
      new ImagePlacer(work.ss, targetData['수행사']).placeDynamicStampsForSheets(dynamicSheetNames);
      didInsert = true;
    }

    if (!staticSheetNames.length && !dynamicSheetNames.length && progress) {
      progress.update(45, opts.noneMessage || '도장/로고 삽입 생략');
    }

    if (didInsert) {
      SpreadsheetApp.flush();
      Utilities.sleep(Number(opts.waitMs) || 0);
    }

    return didInsert;
  }

  getMasterContext_() {
    // 컨테이너 바인드/버튼 실행에서는 현재 UI 스프레드시트의 선택 영역을 기준으로 해야 합니다.
    // openById로 연 Spreadsheet 객체는 UI의 activeRange를 안정적으로 보장하지 않습니다.
    const activeSs = SpreadsheetApp.getActiveSpreadsheet();
    let ss = activeSs;

    if (!ss || ss.getId() !== CONFIG.MASTER_SPREADSHEET_ID) {
      ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    }

    let sheet = null;
    const activeSheet = ss.getActiveSheet ? ss.getActiveSheet() : null;

    if (activeSheet && activeSheet.getName() === CONFIG.SHEETS.MASTER) {
      sheet = activeSheet;
    } else {
      sheet = ss.getSheetByName(CONFIG.SHEETS.MASTER);
    }

    if (!sheet) throw new Error('마스터시트가 없습니다: ' + CONFIG.SHEETS.MASTER);
    const headerMap = HeaderMapper.fromSheet(sheet, CONFIG.ROWS.MASTER_HEADER);
    return { ss, sheet, headerMap };
  }

  readMasterRow_(sheet, headerMap, rowNo) {
    const width = sheet.getLastColumn();
    const values = sheet.getRange(rowNo, 1, 1, width).getValues()[0];
    return RowObject.fromValues(headerMap, values, rowNo);
  }

  getSelectedFileDefs_(rowObj) {
    return CONFIG.FILE_DEFINITIONS.filter(def => {
      const headers = def.checkHeaders || [def.checkHeader];
      return headers.some(header => {
        if (def.checkHeaderContains === true) {
          return isChecked_(firstRowValueByHeaderContains_(rowObj, header));
        }
        return isChecked_(this.valueByHeader_(rowObj, header));
      });
    });
  }

  valueByHeader_(rowObj, header) {
    return rowObj.get(header);
  }

  buildRecipients_(mode, targetData, sender, testInput, manualTo, manualCc, removedCc) {
    const customerEmail = String(targetData['담당자 이메일 주소'] || '').trim();

    if (mode === 'CUSTOMER') {
      const baseTo = splitEmails_(customerEmail);
      const extraTo = normalizeAdditionalEmailList_(manualTo, CONFIG.MAIL.TEST_DOMAIN);
      const extraCc = normalizeAdditionalEmailList_(manualCc, CONFIG.MAIL.TEST_DOMAIN);
      const removedCcSet = buildLowerEmailSet_(normalizeAdditionalEmailList_(removedCc, CONFIG.MAIL.TEST_DOMAIN));

      const to = uniqueEmails_(baseTo.concat(extraTo));
      const baseCc = [sender.email, CONFIG.MAIL.MASTER_CC]
        .filter(email => !removedCcSet[String(email || '').trim().toLowerCase()]);
      const cc = uniqueEmails_(baseCc.concat(extraCc));

      if (!to.length) {
        throw new Error('고객 담당자 이메일 주소가 비어 있습니다. 팝업에서 수신자를 추가하거나 마스터시트의 담당자 이메일 주소를 입력하세요.');
      }

      return { to, cc };
    }

    const testEmail = normalizeTestEmail_(testInput, CONFIG.MAIL.TEST_DOMAIN);

    // v67 핵심:
    // 팝업의 [나에게 발송] TEST 모드는 실제 고객에게 보내지 않고 내부 확인용으로만 발송합니다.
    // 이때 참조(CC)는 영업담당자/마스터/수동 입력값을 모두 제외하고,
    // 발신자(from)는 위에서 resolve한 영업담당자 이메일, 수신자(to)는 사용자가 입력한 회사계정 1개만 사용합니다.
    return { to: [testEmail], cc: [] };
  }

  renderTemplate_(template, targetData, sender, extraValues) {
    const base = {
      company: targetData['회사명'] || '',
      vendor: targetData['수행사'] || '',
      vendorDisplayName: getVendorDisplayName_(targetData['수행사'] || ''),
      salesRepName: sender.name || targetData['영업담당자'] || '',
      salesRepTitle: sender.title || '',
      salesRepDisplay: buildPersonTitleText_(sender.name || targetData['영업담당자'] || '', sender.title || ''),
      salesRepPhone: sender.phone || '',
      salesRepEmail: sender.email || '',
      customerContact: cleanContactText_(targetData['고객사 담당자'] || targetData['담당자 이름'] || targetData['담당자'] || '')
    };
    const values = Object.assign({}, targetData, base, extraValues || {});
    return String(template).replace(/\{([^}]+)\}/g, (_, key) => {
      const raw = values[key] != null ? values[key] : '';
      return escapeHtml_(formatTemplateValueForMailAutoV434_(key, raw));
    });
  }

  updateMasterStatus_(sheet, headerMap, rowNo, status) {
    const setByHeader = (header, value) => {
      const col = headerMap.findCol(header);
      if (col) sheet.getRange(rowNo, col).setValue(value);
    };
    if (status.lastSent != null) setByHeader(CONFIG.MASTER_STATUS_HEADERS.LAST_SENT, status.lastSent);
    if (status.status != null) setByHeader(CONFIG.MASTER_STATUS_HEADERS.STATUS, status.status);
    if (status.sentAt != null) setByHeader(CONFIG.MASTER_STATUS_HEADERS.SENT_AT, status.sentAt);
    if (status.message != null) setByHeader(CONFIG.MASTER_STATUS_HEADERS.MESSAGE, status.message);
    if (status.requestNo != null) setByHeader(CONFIG.MASTER_STATUS_HEADERS.REQUEST_NO, status.requestNo);
    if (status.tempFolder != null) setByHeader(CONFIG.MASTER_STATUS_HEADERS.TEMP_FOLDER, status.tempFolder);

    if (status.incrementCount) {
      const countCol = headerMap.findCol(CONFIG.MASTER_STATUS_HEADERS.SEND_COUNT);
      if (countCol) {
        const cell = sheet.getRange(rowNo, countCol);
        const prev = Number(cell.getValue()) || 0;
        cell.setValue(prev + 1);
      }
    }
    if (status.flush !== false) SpreadsheetApp.flush();
  }

  appendMailLog_(row) {
    const ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    let sheet = ss.getSheetByName(CONFIG.SHEETS.MAIL_LOG);
    if (!sheet) sheet = ss.insertSheet(CONFIG.SHEETS.MAIL_LOG);
    sheet.appendRow(row);
  }
}

class RequestRepository {
  constructor(spreadsheetId, progress) {
    this.spreadsheetId = spreadsheetId;
    this.ss = SpreadsheetApp.openById(spreadsheetId);
    this.sheet = mustGetSheet_(this.ss, CONFIG.SHEETS.REQUEST_LOG);
    this.headerMap = HeaderMapper.fromSheet(this.sheet, CONFIG.ROWS.REQUEST_HEADER);
    this.progress = progress || null;
  }

  registerFromMasterRow(masterRowObj) {
    if (this.progress) this.progress.update(10, '접수번호 등록 준비 중');

    const width = this.sheet.getLastColumn();
    const targetHeaders = this.sheet
      .getRange(CONFIG.ROWS.REQUEST_HEADER, 1, 1, width)
      .getValues()[0];

    const receiptCol = this.findReceiptColumn_(targetHeaders);

    // v16 핵심:
    // - LockService 사용 안 함
    // - Google Sheets API 사용 안 함
    // - 접수번호 칸에 실행별 UUID 마커를 넣어 appendRow
    // - append 후 UUID가 들어간 실제 행을 찾아 row 기준 접수번호 확정
    const marker = this.buildPendingMarker_();

    let valuesForAppend = targetHeaders.map((header, idx) => {
      if (idx + 1 === receiptCol) return marker;
      return masterRowObj.getFlexible(header);
    });

    // v93: 자동견적요청 접수 로그로 이식할 때도 마스터시트와 같은 데이터 타입/표시서식을 유지합니다.
    // - 날짜: 실제 Date 값 + yyyy.MM.dd.
    // - 횟수/개월/금액/면적/할인율: 숫자 값 + 전용 표시서식
    valuesForAppend = normalizeDataFormatRowValuesV434_(targetHeaders, valuesForAppend, masterRowObj.toPlainObject());

    return this.appendBySpreadsheetAppWithMarker_(targetHeaders, valuesForAppend, receiptCol, marker);
  }

  readRegistrationByRowNo(rowNo) {
    const row = Number(rowNo);
    if (!row || row < CONFIG.ROWS.REQUEST_DATA_START) {
      throw new Error('파일 확인/수정 접수 행 번호가 올바르지 않습니다: ' + rowNo);
    }

    const width = this.sheet.getLastColumn();
    const targetHeaders = this.sheet
      .getRange(CONFIG.ROWS.REQUEST_HEADER, 1, 1, width)
      .getValues()[0];
    const values = this.sheet.getRange(row, 1, 1, width).getValues()[0];
    const receiptCol = this.findReceiptColumn_(targetHeaders);
    const requestNo = values[receiptCol - 1] || this.requestNoFromRowNo_(row);

    return {
      requestNo: requestNo,
      requestRowNo: row,
      headers: targetHeaders,
      values: values
    };
  }

  buildPendingMarker_() {
    const prefix = CONFIG.PROGRESS.REQUEST_MARKER_PREFIX || 'MAILAUTO_PENDING_';
    return prefix + Utilities.getUuid();
  }

  appendBySpreadsheetAppWithMarker_(targetHeaders, valuesForAppend, receiptCol, marker) {
    if (this.progress) this.progress.update(12, '자동견적요청 행 추가 중');

    // appendRow는 Apps Script 내장 Spreadsheet 서비스라 별도 Google Sheets API 활성화가 필요 없습니다.
    // 여러 사용자가 동시에 실행해도 각 실행은 자기 UUID 마커가 들어간 행을 다시 찾아 처리합니다.
    this.sheet.appendRow(valuesForAppend);
    SpreadsheetApp.flush();

    const rowNo = this.findMarkerRow_(receiptCol, marker);
    const requestNo = this.requestNoFromRowNo_(rowNo);

    this.sheet.getRange(rowNo, receiptCol).setValue(requestNo);
    // v93: appendRow 이후 자동견적요청 실제 행에도 표시서식을 강제 적용합니다.
    normalizeAndApplyDataFormatsToRowV434_(this.sheet, CONFIG.ROWS.REQUEST_HEADER, rowNo, targetHeaders.length);
    SpreadsheetApp.flush();

    const values = valuesForAppend.slice();
    values[receiptCol - 1] = requestNo;

    if (this.progress) this.progress.update(14, '접수번호 확정: ' + requestNo);

    return {
      requestNo: requestNo,
      requestRowNo: rowNo,
      headers: targetHeaders,
      values: values
    };
  }

  findMarkerRow_(receiptCol, marker) {
    const retry = CONFIG.PROGRESS.REQUEST_MARKER_FIND_RETRY || 12;
    const sleepMs = CONFIG.PROGRESS.REQUEST_MARKER_FIND_SLEEP_MS || 250;

    for (let i = 0; i < retry; i++) {
      const lastRow = Math.max(this.sheet.getLastRow(), CONFIG.ROWS.REQUEST_DATA_START);
      const numRows = Math.max(lastRow - CONFIG.ROWS.REQUEST_DATA_START + 1, 1);
      const range = this.sheet.getRange(CONFIG.ROWS.REQUEST_DATA_START, receiptCol, numRows, 1);
      const cell = range
        .createTextFinder(marker)
        .matchEntireCell(true)
        .findNext();

      if (cell) {
        return cell.getRow();
      }

      SpreadsheetApp.flush();
      Utilities.sleep(sleepMs);
    }

    throw new Error(
      '자동견적요청 append 후 방금 추가한 행을 찾지 못했습니다.\n' +
      '임시마커: ' + marker + '\n' +
      '자동견적요청 시트의 접수번호 열에 위 임시마커가 남아 있는지 확인하세요.'
    );
  }

  requestNoFromRowNo_(rowNo) {
    return rowNo - CONFIG.ROWS.REQUEST_DATA_START + 1;
  }

  findReceiptColumn_(headers) {
    const idx = headers.findIndex(header => normalizeHeader_(header) === normalizeHeader_('접수번호'));
    return idx >= 0 ? idx + 1 : 1;
  }

  // v15에서 쓰던 Google Sheets API 방식은 API 비활성 프로젝트에서 403이 나므로 v16 기본 흐름에서는 사용하지 않습니다.
  // 필요 시 디버깅용으로만 남겨둡니다.
  appendValuesBySheetsApi_(values) {
    const range = this.buildAppendRangeA1_();
    const url =
      'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(this.spreadsheetId) +
      '/values/' + encodeURIComponent(range) +
      ':append?valueInputOption=USER_ENTERED&insertDataOption=INSERT_ROWS&includeValuesInResponse=false';

    const res = UrlFetchApp.fetch(url, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      payload: JSON.stringify({ values: [values] }),
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      },
      muteHttpExceptions: true
    });

    const code = res.getResponseCode();
    const body = res.getContentText();
    if (code < 200 || code >= 300) {
      throw new Error('자동견적요청 append 실패: HTTP ' + code + '\n' + body);
    }

    try {
      return JSON.parse(body);
    } catch (err) {
      throw new Error('자동견적요청 append 응답 파싱 실패: ' + body);
    }
  }

  buildAppendRangeA1_() {
    const sheetName = String(CONFIG.SHEETS.REQUEST_LOG).replace(/'/g, "''");
    const lastCol = Math.max(this.sheet.getLastColumn(), 1);
    return "'" + sheetName + "'!A" + CONFIG.ROWS.REQUEST_DATA_START + ':' + columnToLetter_(lastCol);
  }

  parseUpdatedRowNo_(appendResult) {
    const updatedRange = String(
      appendResult && appendResult.updates && appendResult.updates.updatedRange || ''
    );

    const match = updatedRange.match(/![A-Z]+(\d+)(?::[A-Z]+\d+)?$/i);
    if (!match) {
      throw new Error('자동견적요청 append 결과에서 row 번호를 확인하지 못했습니다: ' + updatedRange);
    }

    const rowNo = Number(match[1]);
    if (!Number.isFinite(rowNo) || rowNo < CONFIG.ROWS.REQUEST_DATA_START) {
      throw new Error('자동견적요청 append row 번호가 비정상입니다: ' + updatedRange);
    }
    return rowNo;
  }

  appendWithShortLockFallback_(targetHeaders, valuesForAppend, receiptCol) {
    const lock = LockService.getScriptLock();
    const waitMs = CONFIG.PROGRESS.REQUEST_LOCK_WAIT_MS || 3000;
    const gotLock = lock.tryLock(waitMs);
    if (!gotLock) {
      throw new Error('접수번호 확정 잠금 시간초과: 잠시 후 다시 누르세요.');
    }

    try {
      const rowNo = Math.max(this.sheet.getLastRow() + 1, CONFIG.ROWS.REQUEST_DATA_START);
      const requestNo = this.requestNoFromRowNo_(rowNo);
      const values = valuesForAppend.slice();
      values[receiptCol - 1] = requestNo;
      this.sheet.getRange(rowNo, 1, 1, values.length).setValues([values]);
      return { requestNo, requestRowNo: rowNo, headers: targetHeaders, values };
    } finally {
      lock.releaseLock();
    }
  }
}


function getMailAutoWorkspaceFolderId_() {
  const propValue = String(PropertiesService.getScriptProperties().getProperty('MAILAUTO_WORKSPACE_FOLDER_ID') || '').trim();
  if (propValue) return extractDriveId_(propValue);

  const configValue = String(CONFIG.TEMP && CONFIG.TEMP.PARENT_FOLDER_ID || '').trim();
  if (configValue) return extractDriveId_(configValue);

  return '';
}

function getMailAutoWorkspaceFolder_() {
  const folderId = getMailAutoWorkspaceFolderId_();
  if (!folderId) return null;
  return DriveApp.getFolderById(folderId);
}

function setMailAutoWorkspaceFolderId() {
  const ui = SpreadsheetApp.getUi();
  const current = getMailAutoWorkspaceFolderId_();
  const res = ui.prompt(
    '메일자동화 작업공간 공유드라이브 폴더ID 저장',
    '도장/로고 캐시, 임시 작업파일, 파일 확인/수정 폴더, 비편집 shortcut 캐시, DOCX 변환 캐시를 저장할 공유드라이브 폴더 ID를 입력하세요.\n\n' +
    '폴더 URL에서 /folders/ 뒤의 ID만 붙여넣어도 됩니다.\n\n' +
    '현재값: ' + (current || '(없음)'),
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) {
    ui.alert('저장을 취소했습니다.');
    return;
  }

  const folderId = extractDriveId_(res.getResponseText());
  if (!folderId) {
    ui.alert('폴더 ID가 비어 있습니다.');
    return;
  }

  const folder = DriveApp.getFolderById(folderId);
  folder.getName(); // 접근 권한 확인
  PropertiesService.getScriptProperties().setProperty('MAILAUTO_WORKSPACE_FOLDER_ID', folderId);

  ui.alert(
    '메일자동화 작업공간 저장 완료\n\n' +
    '폴더명: ' + folder.getName() + '\n' +
    '폴더ID: ' + folderId + '\n\n' +
    '이제 캐시/임시파일/파일확인 폴더는 이 폴더 아래에서 생성됩니다.'
  );
}

function checkMailAutoWorkspaceFolder() {
  const ui = SpreadsheetApp.getUi();
  const folderId = getMailAutoWorkspaceFolderId_();
  if (!folderId) {
    ui.alert(
      '메일자동화 작업공간 폴더ID가 없습니다.\n\n' +
      '현재는 fallback 순서로 저장됩니다.\n' +
      '1) 생성기 파일 부모 폴더\n' +
      '2) 마스터 파일 부모 폴더\n' +
      '3) 현재 활성 파일 부모 폴더\n\n' +
      '공유드라이브에 고정하려면 [작업공간 공유드라이브 폴더ID 저장]을 먼저 실행하세요.'
    );
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    ui.alert(
      '메일자동화 작업공간 확인 완료\n\n' +
      '폴더명: ' + folder.getName() + '\n' +
      '폴더ID: ' + folderId + '\n' +
      '폴더URL: ' + folder.getUrl()
    );
  } catch (err) {
    ui.alert('작업공간 폴더 접근 실패\n\n폴더ID: ' + folderId + '\n오류: ' + String(err && err.message || err));
  }
}

function resolveMailAutoWritableParentFolder_(preferredSourceFile) {
  const errors = [];

  function tryFolder(label, fn) {
    try {
      const folder = fn();
      if (folder) return folder;
    } catch (err) {
      errors.push(label + ': ' + String(err && err.message || err));
    }
    return null;
  }

  // v91 핵심:
  // MAILAUTO_WORKSPACE_FOLDER_ID 또는 CONFIG.TEMP.PARENT_FOLDER_ID가 있으면 무조건 최우선 사용합니다.
  // 이 값을 공유드라이브 폴더로 지정하면 도장/로고 캐시, 임시 작업파일, 파일확인 폴더,
  // shortcut 캐시, DOCX 변환 캐시가 모두 같은 공유드라이브 작업공간 아래에 생성됩니다.
  const configuredWorkspace = tryFolder('MAILAUTO_WORKSPACE_FOLDER_ID/CONFIG.TEMP.PARENT_FOLDER_ID', function() {
    return getMailAutoWorkspaceFolder_();
  });
  if (configuredWorkspace) return configuredWorkspace;

  if (preferredSourceFile) {
    const preferredParent = tryFolder('preferredSourceFile parent', function() {
      return getFirstParentFolderFromDriveFile_(preferredSourceFile);
    });
    if (preferredParent) return preferredParent;
  }

  const generatorParent = tryFolder('GENERATOR_SPREADSHEET parent', function() {
    return getFirstParentFolderFromDriveFile_(DriveApp.getFileById(CONFIG.GENERATOR_SPREADSHEET_ID));
  });
  if (generatorParent) return generatorParent;

  const masterParent = tryFolder('MASTER_SPREADSHEET parent', function() {
    return getFirstParentFolderFromDriveFile_(DriveApp.getFileById(CONFIG.MASTER_SPREADSHEET_ID));
  });
  if (masterParent) return masterParent;

  const activeParent = tryFolder('active spreadsheet parent', function() {
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (!active) return null;
    return getFirstParentFolderFromDriveFile_(DriveApp.getFileById(active.getId()));
  });
  if (activeParent) return activeParent;

  Logger.log('메일자동화 부모 폴더 탐색 실패: ' + errors.join(' / '));
  return null;
}

function createBlankSpreadsheetFileInMailAutoWorkspace_(fileName, folder) {
  const name = sanitizeFileName_(fileName || ('메일자동화_임시시트_' + Utilities.getUuid()));
  const parent = folder || resolveMailAutoWritableParentFolder_(null);
  const parentId = parent && parent.getId ? parent.getId() : '';

  if (parentId) {
    try {
      const res = UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/files?supportsAllDrives=true&fields=id,name,webViewLink,parents', {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({
          name: name,
          mimeType: 'application/vnd.google-apps.spreadsheet',
          parents: [parentId]
        }),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code >= 200 && code < 300) {
        const data = JSON.parse(res.getContentText() || '{}');
        if (data && data.id) return DriveApp.getFileById(data.id);
      }
      Logger.log('공유드라이브 작업공간 스프레드시트 직접 생성 실패. SpreadsheetApp fallback: HTTP ' + code + ' / ' + res.getContentText().slice(0, 500));
    } catch (err) {
      Logger.log('공유드라이브 작업공간 스프레드시트 직접 생성 예외. SpreadsheetApp fallback: ' + (err && err.stack || err));
    }
  }

  const temp = SpreadsheetApp.create(name);
  const file = DriveApp.getFileById(temp.getId());
  if (parent) {
    try {
      file.moveTo(parent);
    } catch (moveErr) {
      try { parent.addFile(file); } catch (addErr) {
        Logger.log('임시 스프레드시트 작업공간 이동/추가 실패: ' + moveErr + ' / ' + addErr);
      }
    }
  }
  return file;
}

function getFirstParentFolderFromDriveFile_(file) {
  if (!file) return null;
  const parents = file.getParents();
  if (parents && parents.hasNext()) return parents.next();
  return null;
}

function authorizeMailAutomationDriveOnce() {
  // 권한 승인용 수동 실행 함수입니다.
  // Apps Script 편집기에서 한 번 실행하면 DriveApp/SpreadsheetApp 권한 재승인 창을 띄울 수 있습니다.
  const masterName = DriveApp.getFileById(CONFIG.MASTER_SPREADSHEET_ID).getName();
  const generatorName = DriveApp.getFileById(CONFIG.GENERATOR_SPREADSHEET_ID).getName();
  SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID).getName();
  SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID).getName();
  return '권한 확인 완료: ' + masterName + ' / ' + generatorName;
}


function warmUpMailAutoPrestampedTemplateCache() {
  const vendorNames = Object.keys(CONFIG.VENDORS || {})
    .map(function(name) { return normalizeVendorName_(name); })
    .filter(Boolean)
    .filter(function(name, idx, arr) { return arr.indexOf(name) === idx; });

  if (!vendorNames.length) {
    try { SpreadsheetApp.getUi().alert('CONFIG.VENDORS에 수행사 정보가 없습니다.'); } catch (err) {}
    return { ok: false, message: 'CONFIG.VENDORS에 수행사 정보가 없습니다.' };
  }

  const logs = [];

  vendorNames.forEach(function(vendorName) {
    const fakeRowObj = {
      get: function(header) {
        const h = String(header || '').trim();
        if (
          h === '수행사' ||
          h === '최종수행사' ||
          h === '최종 수행사' ||
          h === '수행사명' ||
          h === '수행사 선택' ||
          h === '수행업체' ||
          h === '협력사' ||
          h === '최종협력사'
        ) {
          return vendorName;
        }
        return '';
      },
      toPlainObject: function() {
        return {
          '수행사': vendorName,
          '최종수행사': vendorName
        };
      }
    };

    const dummyProgress = {
      update: function(percent, message) {
        Logger.log('[도장/로고 캐시 예열] ' + vendorName + ' / ' + percent + '% / ' + message);
      }
    };

    const fileId = resolvePrestampedGeneratorTemplateId_(fakeRowObj, dummyProgress);
    logs.push(vendorName + ': ' + fileId);
  });

  const message = '도장/로고 캐시 예열 완료\n\n' + logs.join('\n');
  try { SpreadsheetApp.getUi().alert(message); } catch (err) {}
  return { ok: true, message: message, logs: logs };
}

function clearMailAutoPrestampedTemplateCache() {
  const props = PropertiesService.getScriptProperties();
  const prefix = (CONFIG.IMAGE_PRESTAMP_CACHE && CONFIG.IMAGE_PRESTAMP_CACHE.PROPERTY_PREFIX) || 'MAILAUTO_PRESTAMPED_GENERATOR_';
  const all = props.getProperties();
  let count = 0;

  Object.keys(all).forEach(function(key) {
    if (key.indexOf(prefix) !== 0) return;
    try {
      const info = JSON.parse(all[key] || '{}');
      if (info && info.fileId) {
        try { DriveApp.getFileById(info.fileId).setTrashed(true); } catch (e) {}
      }
    } catch (err) {}
    props.deleteProperty(key);
    count++;
  });

  try {
    SpreadsheetApp.getUi().alert('도장/로고 캐시를 초기화했습니다. 삭제 대상: ' + count + '개');
  } catch (err) {}

  return { ok: true, count: count };
}

function resolvePrestampedGeneratorTemplateId_(masterRowObj, progress) {
  const cfg = CONFIG.IMAGE_PRESTAMP_CACHE || {};
  if (cfg.ENABLED !== true) return CONFIG.GENERATOR_SPREADSHEET_ID;

  const vendorName = resolveVendorNameFromMasterRowObj_(masterRowObj);
  if (!vendorName) return CONFIG.GENERATOR_SPREADSHEET_ID;

  const vendor = getVendorConfig_(vendorName);
  if (!vendor) return CONFIG.GENERATOR_SPREADSHEET_ID;

  const vendorKey = normalizeVendorName_(vendorName) || String(vendorName || '').trim();
  const propKey = String(cfg.PROPERTY_PREFIX || 'MAILAUTO_PRESTAMPED_GENERATOR_') + normalizeHeader_(vendorKey);
  const props = PropertiesService.getScriptProperties();
  const source = DriveApp.getFileById(CONFIG.GENERATOR_SPREADSHEET_ID);
  const signature = buildPrestampedGeneratorSignature_(source, vendorKey, vendor);
  const existingRaw = props.getProperty(propKey);

  if (existingRaw) {
    try {
      const existing = JSON.parse(existingRaw);
      if (existing && existing.fileId) {
        const existingFile = DriveApp.getFileById(existing.fileId);
        if (!existingFile.isTrashed()) {
          // v79:
          // 성능 우선. 생성기 로그 append 때문에 signature가 자주 바뀌면 안 됩니다.
          // STRICT_SIGNATURE_MATCH=true일 때만 signature 불일치로 재생성합니다.
          // v87 핵심:
          // 기존에는 STRICT_SIGNATURE_MATCH=false이면 signature가 달라도 기존 캐시를 계속 재사용했습니다.
          // 그 상태에서 생성기 양식/생성대상 값이 오염된 캐시가 남으면 다른 고객 발송에도
          // 예전 고객 회사명이 찍힐 수 있으므로, 이제 signature가 정확히 맞는 캐시만 사용합니다.
          if (existing.signature === signature) {
            return existing.fileId;
          }
          Logger.log('도장/로고 캐시 signature 불일치로 재생성합니다. vendor=' + vendorKey);
        }
      }
    } catch (err) {
      // 캐시 정보가 깨졌으면 아래에서 새로 만듭니다.
    }
  }

  if (progress) {
    progress.update(16, '도장/로고 캐시 생성 중: ' + vendorKey + ' (최초 1회만 느립니다)');
  }

  const cacheFile = createPrestampedGeneratorTemplate_(source, vendorKey, vendorName, signature);
  props.setProperty(propKey, JSON.stringify({
    fileId: cacheFile.getId(),
    vendorKey: vendorKey,
    signature: signature,
    cacheVersion: String((CONFIG.IMAGE_PRESTAMP_CACHE && CONFIG.IMAGE_PRESTAMP_CACHE.CACHE_VERSION) || ''),
    createdAt: new Date().toISOString()
  }));

  return cacheFile.getId();
}

function resolveVendorNameFromMasterRowObj_(rowObj) {
  if (!rowObj) return '';
  const candidates = [
    '수행사', '최종수행사', '최종 수행사', '수행사명', '수행사 선택', '수행업체', '협력사', '최종협력사'
  ];

  for (let i = 0; i < candidates.length; i++) {
    let v = '';
    try { v = rowObj.get(candidates[i]); } catch (err) { v = ''; }
    v = String(v || '').trim();
    if (v) return v;
  }

  const obj = typeof rowObj.toPlainObject === 'function' ? rowObj.toPlainObject() : {};
  for (const key in obj) {
    if (!Object.prototype.hasOwnProperty.call(obj, key)) continue;
    const normalized = normalizeHeader_(key);
    if (normalized.indexOf(normalizeHeader_('수행사')) >= 0 || normalized.indexOf(normalizeHeader_('협력사')) >= 0) {
      const v = String(obj[key] || '').trim();
      if (v) return v;
    }
  }

  return '';
}

function buildPrestampedGeneratorSignature_(sourceFile, vendorKey, vendor) {
  const cfg = CONFIG.IMAGE_PRESTAMP_CACHE || {};
  const cacheVersion = String(cfg.CACHE_VERSION || 'v79_stable').trim();

  // v79:
  // sourceFile.getLastUpdated()는 사용하지 않습니다.
  // 생성기 파일에는 자동견적요청 로그가 계속 append되어 최종수정시각이 매번 바뀌기 때문입니다.
  // 도장 이미지 최종수정시각도 기본적으로 조회하지 않습니다. DriveApp 호출 수를 줄이기 위함입니다.
  // 도장/로고 파일을 실제로 교체했으면 메뉴에서 캐시 초기화 또는 CACHE_VERSION 변경으로 재생성하세요.
  const imageUpdatedMap = {};
  if (cfg.REBUILD_IF_IMAGE_UPDATED === true) {
    (CONFIG.IMAGE_PLACEMENTS || []).forEach(function(p) {
      const role = p.role || '';
      const fileId = String(vendor && vendor[role] || '').trim();
      if (!fileId || Object.prototype.hasOwnProperty.call(imageUpdatedMap, fileId)) return;
      try {
        const imageFile = DriveApp.getFileById(fileId);
        imageUpdatedMap[fileId] = imageFile.getLastUpdated ? imageFile.getLastUpdated().getTime() : 0;
      } catch (err) {
        imageUpdatedMap[fileId] = 'UNKNOWN';
      }
    });
  }

  const placementSignature = (CONFIG.IMAGE_PLACEMENTS || []).map(function(p) {
    const role = p.role || '';
    const fileId = String(vendor && vendor[role] || '').trim();
    return [
      p.sheetName || '',
      role,
      p.rangeA1 || '',
      p.hAlign || '',
      p.vAlign || '',
      fileId,
      imageUpdatedMap[fileId] || ''
    ].join('|');
  }).join(';;');

  return Utilities.base64EncodeWebSafe(Utilities.computeDigest(
    Utilities.DigestAlgorithm.SHA_256,
    [CONFIG.GENERATOR_SPREADSHEET_ID, cacheVersion, vendorKey, placementSignature].join('###'),
    Utilities.Charset.UTF_8
  ));
}

function createPrestampedGeneratorTemplate_(sourceFile, vendorKey, vendorName, signature) {
  const cfg = CONFIG.IMAGE_PRESTAMP_CACHE || {};
  const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
  const cacheName = String(cfg.CACHE_NAME_PREFIX || '메일자동화_도장로고캐시_') + sanitizeFileName_(vendorKey) + '_' + timestamp;
  const parent = resolveMailAutoWritableParentFolder_(sourceFile);
  const cacheFile = parent ? sourceFile.makeCopy(cacheName, parent) : sourceFile.makeCopy(cacheName);

  try {
    const cacheSs = SpreadsheetApp.openById(cacheFile.getId());
    new ImagePlacer(cacheSs, vendorName).placeAll();

    // v87 핵심:
    // 도장/로고 캐시는 수행사별 공용 템플릿입니다.
    // 여기에 직전 테스트 고객의 생성대상 행이 남아 있으면, 이후 다른 사용자의 발송자료에도
    // 예전 회사명이 표시될 수 있으므로 캐시 생성 직후 고객별 런타임 데이터를 반드시 비웁니다.
    resetGeneratorRuntimeDataForCacheV87_(cacheSs);

    SpreadsheetApp.flush();
    // 캐시 파일은 이후 복사본의 원본이므로 첫 생성 때만 이미지 반영 대기합니다.
    Utilities.sleep(Math.min(1000, CONFIG.EXPORT.WAIT_MS_AFTER_IMAGE_INSERT || 800));
    try {
      cacheFile.setDescription('메일자동화 도장/로고 사전삽입 캐시 / vendor=' + vendorKey + ' / sig=' + signature);
    } catch (descErr) {}
    return cacheFile;
  } catch (err) {
    try { cacheFile.setTrashed(true); } catch (trashErr) {}
    throw err;
  }
}

function resetGeneratorRuntimeDataForCacheV87_(ss) {
  if (!ss) return;
  try {
    const targetSheet = ss.getSheetByName(CONFIG.SHEETS.TARGET);
    if (targetSheet) {
      const dataRow = CONFIG.ROWS.TARGET_DATA_ROW;
      const width = Math.max(1, targetSheet.getLastColumn());
      targetSheet.getRange(dataRow, 1, 1, width).clearContent();
      try { targetSheet.getRange(dataRow, 1, 1, width).setNote(''); } catch (noteErr) {}
    }
  } catch (err) {
    Logger.log('v87 생성대상 캐시 초기화 실패: ' + (err && err.stack || err));
  }

  // v89 핵심:
  // 최종 파일생성기는 첫 번째 시트(자동견적요청)를 접수 로그로 사용하고,
  // 실제 견적서/용역신청서 계산은 생성대상!3행만 보게 하는 구조입니다.
  // 도장/로고 캐시/임시 복사본 안에 예전 테스트 고객 로그가 남아 있으면 혼동될 수 있으므로,
  // 캐시/임시 복사본의 자동견적요청 데이터행은 런타임 데이터로 보고 비웁니다.
  try {
    const requestSheet = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
    if (requestSheet) {
      const startRow = CONFIG.ROWS.REQUEST_DATA_START;
      const lastRow = Math.max(requestSheet.getLastRow(), startRow);
      const width = Math.max(1, requestSheet.getLastColumn());
      if (lastRow >= startRow) {
        requestSheet.getRange(startRow, 1, lastRow - startRow + 1, width).clearContent();
      }
    }
  } catch (err) {
    Logger.log('v88 자동견적요청 캐시 초기화 실패: ' + (err && err.stack || err));
  }
}

function firstNonEmptyObjectValueV87_(obj, keys) {
  obj = obj || {};
  for (let i = 0; i < (keys || []).length; i++) {
    const key = keys[i];
    if (Object.prototype.hasOwnProperty.call(obj, key)) {
      const v = obj[key];
      if (v !== null && v !== undefined && String(v).trim() !== '') return v;
    }
  }

  const normalizedKeys = (keys || []).map(function(k) { return normalizeHeader_(k); });
  const actualKeys = Object.keys(obj || {});
  for (let a = 0; a < actualKeys.length; a++) {
    const actual = actualKeys[a];
    const normalized = normalizeHeader_(actual);
    if (normalizedKeys.indexOf(normalized) >= 0) {
      const v = obj[actual];
      if (v !== null && v !== undefined && String(v).trim() !== '') return v;
    }
  }
  return '';
}

function normalizeGeneratorSyncTextV87_(value) {
  return String(value == null ? '' : value)
    .toLowerCase()
    .replace(/주식회사|\(주\)|㈜|유한회사|재단법인|사단법인/g, '')
    .replace(/[\s_\-()\[\]{}.,·ㆍ]/g, '')
    .trim();
}

function sheetDisplayContainsTextV87_(sheet, normalizedNeedle, maxRows, maxCols) {
  if (!sheet || !normalizedNeedle) return false;
  const rows = Math.min(Math.max(1, Number(maxRows || 60)), Math.max(1, sheet.getMaxRows ? sheet.getMaxRows() : sheet.getLastRow()));
  const cols = Math.min(Math.max(1, Number(maxCols || 80)), Math.max(1, sheet.getMaxColumns ? sheet.getMaxColumns() : sheet.getLastColumn()));
  const values = sheet.getRange(1, 1, rows, cols).getDisplayValues();
  for (let r = 0; r < values.length; r++) {
    for (let c = 0; c < values[r].length; c++) {
      const text = normalizeGeneratorSyncTextV87_(values[r][c]);
      if (text && text.indexOf(normalizedNeedle) >= 0) return true;
    }
  }
  return false;
}

function ensureGeneratorSheetsCurrentForTargetV87_(ss, targetData, progress, context) {
  if (!ss) return true;
  const company = String(firstNonEmptyObjectValueV87_(targetData || {}, [
    '회사명', '고객사명', '건물명', '상호', '계약 당사자(사업자등록증상 법인명)'
  ]) || '').trim();
  const normalizedCompany = normalizeGeneratorSyncTextV87_(company);

  // 회사명을 확인할 수 없는 경우에도 최소 flush/sleep로 수식 재계산 시간을 확보합니다.
  if (!normalizedCompany) {
    SpreadsheetApp.flush();
    Utilities.sleep(180);
    return true;
  }

  const sheetNames = unique_([
    CONFIG.SHEETS.QUOTE,
    CONFIG.SHEETS.SERVICE_APP,
    CONFIG.SHEETS.COMPARE_1,
    CONFIG.SHEETS.COMPARE_2,
    '견적서',
    '용역신청서',
    '비교견적(1)',
    '비교견적(2)'
  ].filter(Boolean));

  for (let attempt = 0; attempt < 5; attempt++) {
    SpreadsheetApp.flush();
    Utilities.sleep(attempt === 0 ? 120 : 320);

    for (let i = 0; i < sheetNames.length; i++) {
      const sheet = ss.getSheetByName(sheetNames[i]);
      if (!sheet) continue;
      if (sheetDisplayContainsTextV87_(sheet, normalizedCompany, 70, 90)) {
        return true;
      }
    }
  }

  // 회사명이 견적서에 직접 노출되지 않는 예외 양식도 있을 수 있으므로 발송 자체를 막지는 않습니다.
  // 대신 로그를 남겨 캐시/수식 재계산 문제를 추적할 수 있게 합니다.
  Logger.log('v87 생성기 재계산 확인 경고: ' + (context || '') + ' / 회사명=' + company + ' / fileId=' + (ss.getId ? ss.getId() : ''));
  if (progress && progress.update) {
    try { progress.update(49, '생성기 계산값 확인 중: ' + company); } catch (e) {}
  }
  return false;
}

function enforceCurrentMasterIdentityOnTargetDataV87_(targetData, rowObj) {
  const result = Object.assign({}, targetData || {});
  if (!rowObj || typeof rowObj.toPlainObject !== 'function') return result;
  const current = rowObj.toPlainObject();

  ['고객번호', '회사명', '고객사명', '건물명', '영업담당자', '담당자 이메일 주소', '고객사 담당자'].forEach(function(key) {
    const v = firstNonEmptyObjectValueV87_(current, [key]);
    if (v !== null && v !== undefined && String(v).trim() !== '') result[key] = v;
  });

  return result;
}

function assertReviewPackageMatchesCurrentRowV87_(pkg, rowObj) {
  if (!pkg || !rowObj || !pkg.targetData) return;

  const current = typeof rowObj.toPlainObject === 'function' ? rowObj.toPlainObject() : {};
  const currentCustomerNo = String(firstNonEmptyObjectValueV87_(current, ['고객번호', 'customerNo']) || '').trim().replace(/\.0$/, '');
  const pkgCustomerNo = String(firstNonEmptyObjectValueV87_(pkg.targetData, ['고객번호', 'customerNo']) || '').trim().replace(/\.0$/, '');
  if (currentCustomerNo && pkgCustomerNo && currentCustomerNo !== pkgCustomerNo) {
    throw new Error('파일 확인/수정 세션의 고객번호가 현재 발송 고객과 다릅니다. [파일 확인/수정]을 다시 실행해 주세요. 현재=' + currentCustomerNo + ', 세션=' + pkgCustomerNo);
  }

  const currentCompany = normalizeGeneratorSyncTextV87_(firstNonEmptyObjectValueV87_(current, ['회사명', '고객사명', '건물명']));
  const pkgCompany = normalizeGeneratorSyncTextV87_(firstNonEmptyObjectValueV87_(pkg.targetData, ['회사명', '고객사명', '건물명']));
  if (currentCompany && pkgCompany && currentCompany !== pkgCompany) {
    throw new Error('파일 확인/수정 세션의 회사명이 현재 발송 고객과 다릅니다. [파일 확인/수정]을 다시 실행해 주세요.');
  }
}

function assertPortalPayloadMatchesCurrentMasterRowV88_(payload, rowObj, actionLabel) {
  payload = payload || {};
  if (!rowObj || typeof rowObj.toPlainObject !== 'function') return;

  const current = rowObj.toPlainObject();
  const payloadCustomerNo = String(payload.customerNo || payload.customerNumber || payload['고객번호'] || '').trim().replace(/\.0$/, '');
  const currentCustomerNo = String(firstNonEmptyObjectValueV87_(current, ['고객번호', 'customerNo']) || '').trim().replace(/\.0$/, '');

  if (payloadCustomerNo && currentCustomerNo && payloadCustomerNo !== currentCustomerNo) {
    throw new Error(
      (actionLabel || '자료발송') + ' 대상 불일치: 포털이 보낸 고객번호(' + payloadCustomerNo +
      ')와 Worker가 읽은 마스터 행 고객번호(' + currentCustomerNo + ')가 다릅니다. ' +
      '현재 선택 행/발송탭이 꼬였을 수 있으니 고객을 다시 선택한 뒤 [파일 확인/수정]을 다시 실행하세요.'
    );
  }
}

function syncCurrentRegistrationToTempRequestLogRowV88_(ss, registration) {
  if (!ss || !registration) return;
  const sheet = ss.getSheetByName(CONFIG.SHEETS.REQUEST_LOG);
  if (!sheet) return;

  const headerRow = CONFIG.ROWS.REQUEST_HEADER;
  const dataRow = CONFIG.ROWS.REQUEST_DATA_START;
  const width = Math.max(1, sheet.getLastColumn());
  const headers = sheet.getRange(headerRow, 1, 1, width).getValues()[0];
  const registrationObj = RowObject.fromHeadersAndValues(targetHeadersFromRegistration_(registration), registration.values || [], registration.requestRowNo || dataRow);

  const values = headers.map(function(header, idx) {
    const normalized = normalizeHeader_(header);
    if (normalized === normalizeHeader_('접수번호')) return registration.requestNo || registrationObj.getFlexible(header);
    return registrationObj.getFlexible(header);
  });

  // 자동견적요청 3행을 직접 참조하는 구버전/수정 양식을 위해 현재 접수 데이터를 3행에도 넣습니다.
  sheet.getRange(dataRow, 1, 1, width).clearContent();
  sheet.getRange(dataRow, 1, 1, values.length).setValues([values]);
}


function rewriteRequestLogRow3FormulaToTargetV89_(formula) {
  let text = String(formula || '');
  if (!text || text.indexOf(CONFIG.SHEETS.REQUEST_LOG) < 0) return text;

  // 최종 파일생성기 운영 원칙:
  // - 자동견적요청: 접수 로그 시트
  // - 생성대상!3행: 이번 발송/파일확인 계산용 단일 런타임 행
  //
  // 따라서 견적서/용역신청서/비교견적서 안에서 자동견적요청!3행을 직접 보는 수식은
  // 임시 생성기 복사본 안에서만 생성대상!3행으로 치환합니다.
  const requestSheetNameEscaped = String(CONFIG.SHEETS.REQUEST_LOG || '자동견적요청').replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const quotedRe = new RegExp("'" + requestSheetNameEscaped + "'!\\$?([A-Z]{1,3})\\$?3\\b", 'g');
  const plainRe = new RegExp("\\b" + requestSheetNameEscaped + "!\\$?([A-Z]{1,3})\\$?3\\b", 'g');

  text = text.replace(quotedRe, function(match, col) {
    return "'" + CONFIG.SHEETS.TARGET + "'!" + col + "3";
  });
  text = text.replace(plainRe, function(match, col) {
    return "'" + CONFIG.SHEETS.TARGET + "'!" + col + "3";
  });

  return text;
}

function rewriteGeneratorRequestLogFixedRowRefsToTargetV89_(ss, progress, context) {
  if (!ss) return { changed: 0, scanned: 0 };

  const requestName = String(CONFIG.SHEETS.REQUEST_LOG || '자동견적요청');
  const targetName = String(CONFIG.SHEETS.TARGET || '생성대상');
  let changed = 0;
  let scanned = 0;

  ss.getSheets().forEach(function(sheet) {
    const name = sheet.getName();
    if (name === requestName || name === targetName) return;

    let cells = [];
    try {
      cells = sheet.createTextFinder(requestName)
        .matchFormulaText(true)
        .findAll() || [];
    } catch (err) {
      Logger.log('v89 자동견적요청 직접참조 검색 실패: ' + name + ' / ' + (err && err.stack || err));
      return;
    }

    cells.forEach(function(cell) {
      scanned++;
      const formula = String(cell.getFormula ? cell.getFormula() : '');
      if (!formula) return;

      const rewritten = rewriteRequestLogRow3FormulaToTargetV89_(formula);
      if (rewritten && rewritten !== formula) {
        cell.setFormula(rewritten);
        changed++;
        Logger.log('v89 자동견적요청!3행 직접참조 치환: ' + name + '!' + cell.getA1Notation() + ' / ' + formula + ' -> ' + rewritten);
      }
    });
  });

  if (changed > 0) {
    SpreadsheetApp.flush();
    if (progress && progress.update) {
      try { progress.update(31, '파일생성기 수식 참조 보정 완료: ' + changed + '개'); } catch (e) {}
    }
  }

  Logger.log('v89 파일생성기 수식 참조 보정 결과: context=' + (context || '') + ', scanned=' + scanned + ', changed=' + changed);
  return { changed: changed, scanned: scanned };
}

function assertNoSelectedGeneratorRequestLogRow3RefsV89_(ss, definitions) {
  if (!ss) return true;
  const sheetNames = getSelectedGeneratorSheetNamesV88_(definitions);
  if (!sheetNames.length) return true;

  const requestName = String(CONFIG.SHEETS.REQUEST_LOG || '자동견적요청');
  const requestSheetNameEscaped = requestName.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  const quotedRe = new RegExp("'" + requestSheetNameEscaped + "'!\\$?[A-Z]{1,3}\\$?3\\b");
  const plainRe = new RegExp("\\b" + requestSheetNameEscaped + "!\\$?[A-Z]{1,3}\\$?3\\b");

  const hits = [];
  sheetNames.forEach(function(sheetName) {
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) return;

    let cells = [];
    try {
      cells = sheet.createTextFinder(requestName).matchFormulaText(true).findAll() || [];
    } catch (err) {
      return;
    }

    cells.forEach(function(cell) {
      const formula = String(cell.getFormula ? cell.getFormula() : '');
      if (quotedRe.test(formula) || plainRe.test(formula)) {
        hits.push(sheetName + '!' + cell.getA1Notation() + '=' + formula);
      }
    });
  });

  if (hits.length) {
    throw new Error(
      '파일생성기 수식에 자동견적요청 3행 직접참조가 남아 있습니다. 발송을 중단합니다.\n' +
      '최종 파일생성기에서는 자동견적요청은 접수 로그, 생성대상!3행은 현재 발송 고객 계산행이어야 합니다.\n\n' +
      hits.slice(0, 20).join('\n')
    );
  }

  return true;
}

function getSelectedGeneratorSheetNamesV88_(definitions) {
  const names = [];
  (definitions || []).forEach(function(def) {
    if (!def) return;
    if ((def.type === 'sheet_pdf' || def.type === 'sheet_xlsx_values') && def.sheetName) names.push(def.sheetName);
    if (def.type === 'multi_sheet_pdf' && Array.isArray(def.sheets)) {
      def.sheets.forEach(function(item) { if (item && item.sheetName) names.push(item.sheetName); });
    }
  });
  return unique_(names).filter(Boolean);
}

function ensureSelectedGeneratorSheetsShowCurrentCustomerV88_(ss, targetData, definitions, progress, context) {
  if (!ss) return true;

  const company = String(firstNonEmptyObjectValueV87_(targetData || {}, [
    '회사명', '고객사명', '건물명', '상호', '계약 당사자(사업자등록증상 법인명)'
  ]) || '').trim();
  const normalizedCompany = normalizeGeneratorSyncTextV87_(company);

  if (!normalizedCompany) {
    SpreadsheetApp.flush();
    Utilities.sleep(180);
    return true;
  }

  const sheetNames = getSelectedGeneratorSheetNamesV88_(definitions);
  if (!sheetNames.length) return true;

  for (let attempt = 0; attempt < 6; attempt++) {
    SpreadsheetApp.flush();
    Utilities.sleep(attempt === 0 ? 150 : 350);

    for (let i = 0; i < sheetNames.length; i++) {
      const sheet = ss.getSheetByName(sheetNames[i]);
      if (!sheet) continue;
      if (sheetDisplayContainsTextV87_(sheet, normalizedCompany, 80, 100)) return true;
    }
  }

  const message =
    '생성기 양식에 현재 고객명이 반영되지 않았습니다. 발송을 중단합니다.\n' +
    '현재 고객: ' + company + '\n' +
    '확인 시트: ' + sheetNames.join(', ') + '\n\n' +
    '가능 원인: 파일생성기 수식이 생성대상!3행이 아닌 과거 테스트행 또는 고정값을 직접 참조하고 있음.\n' +
    'Worker v89는 자동견적요청!3행 직접참조를 생성대상!3행으로 자동 치환하지만, 그래도 실패하면 파일생성기 수식/값 고정을 확인해야 합니다.';

  if (progress && progress.fail) {
    try { progress.fail(message); } catch (e) {}
  }
  throw new Error(message);
}


function getMailAutoDataFormatRulesV434_() {
  return [
    {
      type: 'date',
      format: 'yyyy.MM.dd.',
      aliases: ['계약시작일', '계약 시작일', '계약개시일', '계약 개시일', '계약시작일자', '계약 시작일자', '용역시작일', '용역 시작일']
    },
    {
      type: 'date',
      format: 'yyyy.MM.dd.',
      aliases: ['계약종료일', '계약 종료일', '계약만료일', '계약 만료일', '계약종료일자', '계약 종료일자', '용역종료일', '용역 종료일']
    },
    {
      type: 'date',
      format: 'yyyy.MM.dd.',
      aliases: ['마스터시트 최초등록일', '최초등록일', '최초 등록일', '등록일']
    },
    {
      type: 'number',
      format: '0"개월"',
      aliases: ['계약단위', '계약 단위', '계약기간', '계약 기간', '계약개월', '계약 개월', '점검기간']
    },
    {
      type: 'number',
      format: '0"회"',
      aliases: ['유지점검', '유지점검횟수', '유지 횟수', '유지횟수', '유지보수·관리 점검', '유지보수관리점검']
    },
    {
      type: 'number',
      format: '0"회"',
      aliases: ['성능점검', '성능점검횟수', '성능 횟수', '성능횟수', '정보통신설비 성능점검']
    },
    {
      type: 'number',
      format: '#,##0.##',
      aliases: ['연면적', '연면적(㎡)', '연면적㎡']
    },
    {
      type: 'number',
      format: '"₩"#,##0',
      aliases: ['최종 견적가', '최종견적가', '최종 견적가\n(부가세 별도 기준)', '견적가', '견적금액', '최종금액', '계약금액']
    },
    {
      type: 'number',
      format: '0.##',
      aliases: ['할인율', '할인률', '할인', '적용할인율']
    }
  ];
}

function findDataFormatRuleByHeaderV434_(header) {
  const normalized = normalizeHeader_(header);
  if (!normalized) return null;
  const rules = getMailAutoDataFormatRulesV434_();
  for (let i = 0; i < rules.length; i++) {
    const aliases = (rules[i].aliases || []).map(normalizeHeader_).filter(Boolean);
    if (aliases.indexOf(normalized) >= 0) return rules[i];
  }
  return null;
}

function formatYyyyMmDdV434_(date) {
  const d = dateOnlyV90_(date);
  if (!d) return '';
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy.MM.dd.');
}

function parsePlainNumberV434_(value) {
  if (value === null || value === undefined || value === '') return '';
  if (typeof value === 'number' && Number.isFinite(value)) return value;
  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) return '';

  const raw = String(value || '').trim();
  if (!raw) return '';

  const cleaned = raw
    .replace(/[₩￦원,\s]/g, '')
    .replace(/개월|회|㎡|m2|M2|퍼센트|%/g, '')
    .replace(/[^0-9.\-]/g, '');

  if (!cleaned || cleaned === '-' || cleaned === '.' || cleaned === '-.') return '';
  const n = Number(cleaned);
  return Number.isFinite(n) ? n : '';
}

function roundTwoIfNumberV434_(value) {
  if (value === '' || value === null || typeof value === 'undefined') return value;
  const n = Number(value);
  if (!Number.isFinite(n)) return value;
  return Math.round(n * 100) / 100;
}

function normalizeValueByDataFormatRuleV434_(header, value, sourceObj) {
  const rule = findDataFormatRuleByHeaderV434_(header);
  if (!rule) return value;

  const normalizedHeader = normalizeHeader_(header);

  if (rule.type === 'date') {
    let v = value;
    if ((v === null || v === undefined || v === '') && sourceObj) {
      const aliases = rule.aliases || [];
      v = firstNonEmptyObjectValueV87_(sourceObj, aliases);
    }

    // 계약시작일/계약종료일은 비어 있으면 v90 기본값 보정값을 사용합니다.
    if ((v === null || v === undefined || v === '') && sourceObj) {
      const fixed = applyContractPeriodDefaultsToObjectV90_(sourceObj || {});
      if (getContractStartAliasesV90_().map(normalizeHeader_).indexOf(normalizedHeader) >= 0) v = fixed['계약시작일'];
      if (getContractEndAliasesV90_().map(normalizeHeader_).indexOf(normalizedHeader) >= 0) v = fixed['계약종료일'];
    }

    return parseContractDateV90_(v) || '';
  }

  const n = parsePlainNumberV434_(value);
  if (n === '') return '';

  if (normalizedHeader === normalizeHeader_('할인율') || normalizedHeader === normalizeHeader_('할인률') || normalizedHeader === normalizeHeader_('할인') || normalizedHeader === normalizeHeader_('적용할인율')) {
    return roundTwoIfNumberV434_(n);
  }

  if (rule.format === '0"회"' || rule.format === '0"개월"') {
    return Math.round(n);
  }

  if (rule.format === '#,##0.##') {
    return roundTwoIfNumberV434_(n);
  }

  if (rule.format === '"₩"#,##0') {
    return Math.round(n);
  }

  return n;
}

function normalizeDataFormatRowValuesV434_(headers, rowValues, sourceObj) {
  const out = (rowValues || []).slice();
  const source = sourceObj || {};
  (headers || []).forEach(function(header, idx) {
    const rule = findDataFormatRuleByHeaderV434_(header);
    if (!rule) return;
    out[idx] = normalizeValueByDataFormatRuleV434_(header, out[idx], source);
  });
  return out;
}

function applyDataFormatsToRowV434_(sheet, headers, rowNo, width) {
  if (!sheet || !headers || !headers.length) return;
  const max = Math.min(Number(width) || headers.length, headers.length);
  for (let i = 0; i < max; i++) {
    const rule = findDataFormatRuleByHeaderV434_(headers[i]);
    if (!rule || !rule.format) continue;
    try {
      sheet.getRange(rowNo, i + 1).setNumberFormat(rule.format);
    } catch (err) {
      Logger.log('v434 데이터 표시서식 적용 실패: row=' + rowNo + ', col=' + (i + 1) + ', header=' + headers[i] + ' / ' + (err && err.stack || err));
    }
  }
}

function normalizeAndApplyDataFormatsToRowV434_(sheet, headerRow, dataRow, width) {
  if (!sheet) return;
  const colCount = Math.max(1, Number(width) || sheet.getLastColumn());
  const headers = sheet.getRange(headerRow, 1, 1, colCount).getValues()[0];
  const range = sheet.getRange(dataRow, 1, 1, colCount);
  const values = range.getValues()[0];
  const normalized = normalizeDataFormatRowValuesV434_(headers, values, null);
  range.setValues([normalized]);
  applyDataFormatsToRowV434_(sheet, headers, dataRow, colCount);
}

function getContractStartAliasesV90_() {
  return ['계약시작일', '계약 시작일', '계약개시일', '계약 개시일', '계약시작일자', '계약 시작일자', '용역시작일', '용역 시작일'];
}

function getContractEndAliasesV90_() {
  return ['계약종료일', '계약 종료일', '계약만료일', '계약 만료일', '계약종료일자', '계약 종료일자', '용역종료일', '용역 종료일'];
}

function getContractUnitAliasesV90_() {
  return ['계약단위', '계약 단위', '계약기간', '계약 기간', '계약개월', '계약 개월', '점검기간'];
}

function dateOnlyV90_(date) {
  if (!date || Object.prototype.toString.call(date) !== '[object Date]' || isNaN(date.getTime())) return null;
  return new Date(date.getFullYear(), date.getMonth(), date.getDate());
}

function parseContractDateV90_(value) {
  if (value === null || value === undefined || value === '') return null;

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return dateOnlyV90_(value);
  }

  if (typeof value === 'number' && Number.isFinite(value)) {
    // Google Sheets/Excel serial date fallback. 2026-07-01 is around 46204.
    if (value > 30000 && value < 70000) {
      const utc = new Date(Math.round((value - 25569) * 86400 * 1000));
      return new Date(utc.getUTCFullYear(), utc.getUTCMonth(), utc.getUTCDate());
    }
  }

  const raw = String(value || '').trim();
  if (!raw) return null;

  const numeric = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(numeric) && numeric > 30000 && numeric < 70000) {
    return parseContractDateV90_(numeric);
  }

  const match = raw.match(/(\d{2,4})\s*[.\-\/년]\s*(\d{1,2})\s*[.\-\/월]\s*(\d{1,2})/);
  if (!match) return null;

  let year = Number(match[1]);
  if (year < 100) year += 2000;
  const month = Number(match[2]);
  const day = Number(match[3]);
  if (!year || !month || !day) return null;

  const parsed = new Date(year, month - 1, day);
  if (isNaN(parsed.getTime())) return null;
  return parsed;
}

function parseContractMonthsV90_(value) {
  if (value === null || value === undefined || value === '') return 12;

  if (typeof value === 'number' && Number.isFinite(value) && value > 0) {
    return Math.max(1, Math.round(value));
  }

  const raw = String(value || '').trim();
  if (!raw) return 12;

  const yearMatch = raw.match(/(\d+(?:\.\d+)?)\s*년/);
  if (yearMatch) {
    return Math.max(1, Math.round(Number(yearMatch[1]) * 12));
  }

  const monthMatch = raw.match(/(\d+(?:\.\d+)?)\s*개?월/);
  if (monthMatch) {
    return Math.max(1, Math.round(Number(monthMatch[1])));
  }

  const numberMatch = raw.replace(/,/g, '').match(/\d+(?:\.\d+)?/);
  if (numberMatch) {
    return Math.max(1, Math.round(Number(numberMatch[0])));
  }

  return 12;
}

function addMonthsMinusOneDayV90_(startDate, months) {
  const start = dateOnlyV90_(startDate) || new Date(2026, 6, 1);
  const m = Math.max(1, Number(months) || 12);
  return new Date(start.getFullYear(), start.getMonth() + m, start.getDate() - 1);
}

function formatYyMmDdV90_(date) {
  // v93: 기존 함수명은 호환용으로 유지하되, 계약/점검기간 표시는 yyyy.MM.dd. 기준으로 통일합니다.
  return formatYyyyMmDdV434_(date);
}

function applyContractPeriodDefaultsToObjectV90_(obj) {
  const result = Object.assign({}, obj || {});

  const rawStart = firstNonEmptyObjectValueV87_(result, getContractStartAliasesV90_());
  const rawEnd = firstNonEmptyObjectValueV87_(result, getContractEndAliasesV90_());
  const rawUnit = firstNonEmptyObjectValueV87_(result, getContractUnitAliasesV90_());

  const startDate = parseContractDateV90_(rawStart) || new Date(2026, 6, 1);
  const months = parseContractMonthsV90_(rawUnit);
  const endDate = parseContractDateV90_(rawEnd) || addMonthsMinusOneDayV90_(startDate, months);

  const startText = formatYyMmDdV90_(startDate);
  const endText = formatYyMmDdV90_(endDate);
  const periodText = startText && endText ? (startText + ' ~ ' + endText) : '';

  result['계약시작일'] = startText;
  result['계약종료일'] = endText;
  result['점검기간'] = periodText;
  result['계약기간표시'] = periodText;

  return result;
}

function setRowValueByHeaderAliasesV90_(headers, rowValues, aliases, value) {
  const wanted = (aliases || []).map(function(alias) { return normalizeHeader_(alias); }).filter(Boolean);
  for (let i = 0; i < (headers || []).length; i++) {
    const normalized = normalizeHeader_(headers[i]);
    if (wanted.indexOf(normalized) >= 0) {
      rowValues[i] = value;
    }
  }
}

function normalizeContractPeriodRowValuesV90_(headers, rowValues, sourceObj) {
  const obj = {};
  (headers || []).forEach(function(header, idx) {
    if (header !== null && header !== undefined && String(header).trim() !== '') {
      obj[String(header)] = rowValues[idx];
    }
  });

  const merged = mergeObjectsPreferNonEmpty_(sourceObj || {}, obj);
  const fixed = applyContractPeriodDefaultsToObjectV90_(merged);

  const startText = String(fixed['계약시작일'] || '').trim();
  const endText = String(fixed['계약종료일'] || '').trim();

  const out = (rowValues || []).slice();
  setRowValueByHeaderAliasesV90_(headers, out, getContractStartAliasesV90_(), startText);
  setRowValueByHeaderAliasesV90_(headers, out, getContractEndAliasesV90_(), endText);
  return out;
}

function writeContractPeriodToTargetSheetV90_(targetSheet, targetData) {
  if (!targetSheet) return;
  const headerRow = CONFIG.ROWS.TARGET_HEADER;
  const dataRow = CONFIG.ROWS.TARGET_DATA_ROW;
  const width = Math.max(1, targetSheet.getLastColumn());
  const headers = targetSheet.getRange(headerRow, 1, 1, width).getValues()[0];
  const fixed = applyContractPeriodDefaultsToObjectV90_(targetData || {});

  const startDate = parseContractDateV90_(fixed['계약시작일']);
  const endDate = parseContractDateV90_(fixed['계약종료일']);

  headers.forEach(function(header, idx) {
    const normalized = normalizeHeader_(header);
    const cell = targetSheet.getRange(dataRow, idx + 1);
    if (getContractStartAliasesV90_().map(normalizeHeader_).indexOf(normalized) >= 0) {
      cell.setNumberFormat('yyyy.MM.dd.').setValue(startDate || '');
    }
    if (getContractEndAliasesV90_().map(normalizeHeader_).indexOf(normalized) >= 0) {
      cell.setNumberFormat('yyyy.MM.dd.').setValue(endDate || '');
    }
  });

  // v93: 계약기간만 별도 보정해도 다른 주요 계약조건 서식은 같이 유지합니다.
  applyDataFormatsToRowV434_(targetSheet, headers, dataRow, width);
}

function buildContractPeriodTextV90_(targetData) {
  const fixed = applyContractPeriodDefaultsToObjectV90_(targetData || {});
  return String(fixed['점검기간'] || '').trim();
}

function applyQuoteInspectionPeriodToSheetV90_(sheet, targetData) {
  if (!sheet) return;
  const periodText = buildContractPeriodTextV90_(targetData || {});
  if (!periodText) return;

  // 최종 파일생성기 기준 견적서 점검기간 값 셀. 원본 양식에서는 P12가 '생성대상'!R3&" ~ "&'생성대상'!S3 입니다.
  // 일부 환경에서 날짜 수식 재계산이 늦거나 R/S가 빈값으로 넘어오면 " ~ "만 출력되므로,
  // 발송/파일확인용 임시 복사본에서는 표시 텍스트를 직접 고정합니다.
  try {
    sheet.getRange('P12').setNumberFormat('@').setValue(periodText);
  } catch (err) {
    Logger.log('v90 견적서 점검기간 직접 반영 실패: ' + (err && err.stack || err));
  }
}

function applyContractPeriodToGeneratorSheetsV90_(ss, targetData, progress, context) {
  if (!ss) return targetData || {};
  const fixed = applyContractPeriodDefaultsToObjectV90_(targetData || {});

  try {
    const targetSheet = ss.getSheetByName(CONFIG.SHEETS.TARGET);
    if (targetSheet) writeContractPeriodToTargetSheetV90_(targetSheet, fixed);
  } catch (err) {
    Logger.log('v90 생성대상 계약기간 반영 실패: ' + (err && err.stack || err));
  }

  try {
    const quoteSheet = ss.getSheetByName(CONFIG.SHEETS.QUOTE) || ss.getSheetByName('견적서');
    if (quoteSheet) applyQuoteInspectionPeriodToSheetV90_(quoteSheet, fixed);
  } catch (err) {
    Logger.log('v90 견적서 계약기간 반영 실패: ' + (err && err.stack || err));
  }

  if (progress && progress.update) {
    try { progress.update(33, '계약/점검기간 반영: ' + buildContractPeriodTextV90_(fixed)); } catch (e) {}
  }

  SpreadsheetApp.flush();
  Utilities.sleep(120);
  return fixed;
}


class GeneratorWorkspace {
  constructor(registration, masterRowObj, progress) {
    this.registration = registration;
    this.masterRowObj = masterRowObj;
    this.progress = progress || null;
    this.file = null;
    this.ss = null;
    this.isPrestamped = false;
  }

  createTempCopy() {
    const sourceFileId = resolvePrestampedGeneratorTemplateId_(this.masterRowObj, this.progress);
    this.isPrestamped = sourceFileId && sourceFileId !== CONFIG.GENERATOR_SPREADSHEET_ID;
    const source = DriveApp.getFileById(sourceFileId || CONFIG.GENERATOR_SPREADSHEET_ID);
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const company = sanitizeFileName_(this.masterRowObj.get('회사명') || '무명고객');
    const name = `메일자동발송_작업_${this.registration.requestNo}_${company}_${timestamp}`;

    // v70: source.makeCopy(name)만 쓰면 My Drive 루트에 접근하려다
    // 일부 Workspace/공유드라이브 환경에서 "액세스가 거부됨: DriveApp"이 날 수 있습니다.
    // 그래서 우선 CONFIG.TEMP.PARENT_FOLDER_ID, 그다음 생성기 파일의 부모 폴더,
    // 그다음 마스터 파일의 부모 폴더를 사용하고, 마지막에만 루트 복사를 시도합니다.
    const parentFolder = resolveMailAutoWritableParentFolder_(source);
    try {
      this.file = parentFolder ? source.makeCopy(name, parentFolder) : source.makeCopy(name);
    } catch (err) {
      throw new Error(
        '임시 작업파일 복사 중 Drive 접근이 거부되었습니다.\n' +
        '생성기 파일 또는 설정된 TEMP.PARENT_FOLDER_ID 폴더에 현재 실행 계정의 편집 권한이 있는지 확인하세요.\n\n' +
        String(err && err.message || err)
      );
    }

    this.ss = SpreadsheetApp.openById(this.file.getId());
    // v87: 캐시 또는 원본 생성기에서 복사된 직전 고객 데이터가 남아 있지 않도록
    // 실제 생성대상 로드 전 복사본의 런타임 데이터도 한 번 비웁니다.
    resetGeneratorRuntimeDataForCacheV87_(this.ss);
    SpreadsheetApp.flush();
    return this.file;
  }

  loadTargetRow() {
    if (!this.ss) throw new Error('임시 생성기 파일이 아직 생성되지 않았습니다.');
    const targetSheet = mustGetSheet_(this.ss, CONFIG.SHEETS.TARGET);
    const targetHeaderMap = HeaderMapper.fromSheet(targetSheet, CONFIG.ROWS.TARGET_HEADER);
    const targetHeaders = targetSheet.getRange(CONFIG.ROWS.TARGET_HEADER, 1, 1, targetSheet.getLastColumn()).getValues()[0];

    const registrationObj = RowObject.fromHeadersAndValues(targetHeadersFromRegistration_(this.registration), this.registration.values, this.registration.requestRowNo);
    let rowValues = targetHeaders.map(header => registrationObj.getFlexible(header));

    // v90 핵심:
    // 계약시작일/계약종료일이 마스터 또는 접수 로그에서 비어 있으면
    // 계약시작일 26.07.01 기본값 + 계약단위 기준 종료일을 계산해 생성대상 R/S에 먼저 넣습니다.
    const sourceForContractPeriodV90 = this.masterRowObj && typeof this.masterRowObj.toPlainObject === 'function'
      ? mergeObjectsPreferNonEmpty_(this.masterRowObj.toPlainObject(), registrationObj.toPlainObject())
      : registrationObj.toPlainObject();
    rowValues = normalizeContractPeriodRowValuesV90_(targetHeaders, rowValues, sourceForContractPeriodV90);
    // v93: 생성대상!3행으로 이식할 때도 마스터와 동일한 데이터 타입/표시서식을 유지합니다.
    rowValues = normalizeDataFormatRowValuesV434_(targetHeaders, rowValues, sourceForContractPeriodV90);

    // v89 핵심:
    // 최종 파일생성기는 첫 번째 시트(자동견적요청)에 접수 로그를 쌓고,
    // 실제 계산은 생성대상!3행 하나만 보도록 운영합니다.
    // 따라서 임시 생성기 안의 자동견적요청!3행에 현재 고객을 억지로 덮어쓰지 않고,
    // 잘못 남은 자동견적요청!3행 직접참조 수식만 생성대상!3행 참조로 보정합니다.
    rewriteGeneratorRequestLogFixedRowRefsToTargetV89_(this.ss, this.progress, '생성대상 로드 전');

    const row = CONFIG.ROWS.TARGET_DATA_ROW;
    const width = rowValues.length;
    targetSheet.getRange(row, 1, 1, width).clearContent();
    SpreadsheetApp.flush();
    Utilities.sleep(80);
    targetSheet.getRange(row, 1, 1, width).setValues([rowValues]);
    applyDataFormatsToRowV434_(targetSheet, targetHeaders, row, width);
    SpreadsheetApp.flush();

    const rowObj = RowObject.fromValues(targetHeaderMap, rowValues, row);
    let targetData = rowObj.toPlainObject();
    if (this.masterRowObj && typeof this.masterRowObj.toPlainObject === 'function') {
      targetData = mergeObjectsPreferNonEmpty_(this.masterRowObj.toPlainObject(), targetData);
    }
    targetData = applyContractPeriodToGeneratorSheetsV90_(this.ss, targetData, this.progress, '생성대상 로드 직후');
    ensureGeneratorSheetsCurrentForTargetV87_(this.ss, targetData, this.progress, '생성대상 로드 직후');
    return targetData;
  }
}

class SalesRepResolver {
  constructor(generatorSs) {
    this.sheet = this.getSalesRepSheet_(generatorSs);
    this.headerMap = HeaderMapper.fromSheet(this.sheet, CONFIG.ROWS.SALES_REP_HEADER);
  }

  getSalesRepSheet_(ss) {
    const sheetNames = [
      CONFIG.SHEETS.SALES_REP, // 기본: 파일생성기 기준 '영업담당자 정보'
      '영업담당자 정보',
      '영업담당자'
    ];

    const uniqueNames = sheetNames
      .map(name => String(name || '').trim())
      .filter(Boolean)
      .filter((name, idx, arr) => arr.indexOf(name) === idx);

    for (const name of uniqueNames) {
      const sheet = ss.getSheetByName(name);
      if (sheet) return sheet;
    }

    throw new Error('영업담당자 정보 시트를 찾지 못했습니다. 확인한 시트명: ' + uniqueNames.join(', '));
  }

  resolve(name) {
    const cleanName = String(name || '').trim();
    if (!cleanName) throw new Error('영업담당자 값이 비어 있습니다.');

    const lastRow = this.sheet.getLastRow();
    const lastCol = this.sheet.getLastColumn();
    const values = this.sheet.getRange(CONFIG.ROWS.SALES_REP_DATA_START, 1, Math.max(0, lastRow - 1), lastCol).getValues();

    const nameCol = this.headerMap.findCol('이름');
    const emailCol = this.headerMap.findCol('이메일');
    const phoneCol = this.headerMap.findCol('전화번호');
    const titleCol = this.headerMap.findCol('직급');

    if (!nameCol || !emailCol) {
      throw new Error(this.sheet.getName() + ' 시트에 이름/이메일 헤더가 필요합니다.');
    }

    for (const row of values) {
      if (String(row[nameCol - 1] || '').trim() === cleanName) {
        const email = String(row[emailCol - 1] || '').trim();
        if (!email) throw new Error(cleanName + ' 담당자의 이메일이 비어 있습니다.');

        return {
          name: cleanName,
          email,
          phone: phoneCol ? row[phoneCol - 1] : '',
          title: titleCol ? row[titleCol - 1] : ''
        };
      }
    }

    throw new Error(this.sheet.getName() + ' 시트에서 담당자를 찾지 못했습니다: ' + cleanName);
  }
}

class ImagePlacer {
  constructor(ss, vendorName) {
    this.ss = ss;
    this.vendorName = String(vendorName || '').trim();
    this.vendor = getVendorConfig_(this.vendorName);
    this.blobCache = {};
  }

  placeAll() {
    this.placeForSheets(unique_(CONFIG.IMAGE_PLACEMENTS.map(p => p.sheetName)));
  }

  placeForSheets(sheetNames) {
    const targets = unique_(sheetNames || []);
    if (!targets.length) return;
    if (!this.vendor) {
      throw new Error('수행사 정보가 CONFIG.VENDORS에 없습니다: ' + this.vendorName);
    }

    targets.forEach(name => this.removeAutoImages_(mustGetSheet_(this.ss, name)));

    for (const placement of CONFIG.IMAGE_PLACEMENTS) {
      if (targets.indexOf(placement.sheetName) < 0) continue;
      const fileId = this.vendor[placement.role];
      if (!fileId) throw new Error(`${this.vendorName} 수행사의 이미지 파일 ID가 비어 있습니다: ${placement.role}`);
      const sheet = mustGetSheet_(this.ss, placement.sheetName);
      this.insertFittedImage_(sheet, fileId, placement, this.vendorName);
    }
  }

  placeDynamicStampsForSheets(sheetNames) {
    const targets = unique_(sheetNames || []);
    if (!targets.length) return;

    const placements = (CONFIG.DYNAMIC_STAMP_PLACEMENTS || [])
      .filter(function(placement) { return targets.indexOf(placement.sheetName) >= 0; });

    unique_(placements.map(function(p) { return p.sheetName; }))
      .forEach(name => this.removeAutoImages_(mustGetSheet_(this.ss, name)));

    placements.forEach(placement => {
      const sheet = mustGetSheet_(this.ss, placement.sheetName);
      const vendorName = this.readDynamicVendorName_(sheet, placement);
      const vendor = getVendorConfig_(vendorName);
      if (!vendor) {
        throw new Error(
          placement.sheetName + ' ' + placement.vendorCellA1 + ' 수행사 값을 CONFIG.VENDORS에서 찾지 못했습니다: ' + vendorName
        );
      }

      const role = String(placement.role || 'stampFileId').trim();
      const fileId = String(vendor[role] || '').trim();
      if (!fileId) {
        throw new Error(placement.sheetName + ' 수행사 도장 파일 ID가 비어 있습니다: ' + vendorName + ' / ' + role);
      }

      this.insertFittedImage_(sheet, fileId, placement, vendorName);
    });
  }

  readDynamicVendorName_(sheet, placement) {
    const a1 = String(placement.vendorCellA1 || '').trim();
    if (!a1) throw new Error('동적 도장 설정에 vendorCellA1이 없습니다: ' + placement.sheetName);

    const range = sheet.getRange(a1);
    const display = String(range.getDisplayValue ? range.getDisplayValue() : '').trim();
    const value = String(range.getValue ? range.getValue() : '').trim();
    const vendorName = display || value;
    if (!vendorName) {
      throw new Error(placement.sheetName + ' ' + a1 + ' 수행사 값이 비어 있어 도장을 선택할 수 없습니다.');
    }
    return vendorName;
  }

  getImageBlob_(fileId) {
    const key = String(fileId || '').trim();
    if (!key) throw new Error('이미지 파일 ID가 비어 있습니다.');
    if (!this.blobCache[key]) {
      this.blobCache[key] = DriveApp.getFileById(key).getBlob();
    }
    const blob = this.blobCache[key];
    return blob.copyBlob ? blob.copyBlob() : blob;
  }

  removeAutoImages_(sheet) {
    const prefix = 'AUTO_MAIL_ASSET';

    sheet.getImages().forEach(img => {
      const title = img.getAltTextTitle ? String(img.getAltTextTitle() || '') : '';
      const desc = img.getAltTextDescription ? String(img.getAltTextDescription() || '') : '';

      // 중요 v28:
      // 템플릿에 원래 들어있는 에스원/SECOM 로고는 절대 위치 기준으로 삭제하면 안 됩니다.
      // 기존 v27의 overlap/near-anchor 삭제는 견적서 하단의 기본 에스원 로고까지 지워버렸습니다.
      // 그래서 이제 자동화가 삽입하면서 AUTO_MAIL_ASSET 태그를 단 이미지 만 삭제합니다.
      // 즉, 수동/기본 템플릿 로고는 보존하고, 이전 자동화 실행분만 정리합니다.
      const isAutoInserted = title.indexOf(prefix) === 0 || desc.indexOf(prefix) >= 0;

      if (isAutoInserted) {
        img.remove();
      }
    });
  }

  insertFittedImage_(sheet, fileId, placement, vendorNameForAlt) {
    const range = sheet.getRange(placement.rangeA1);
    const box = rangePixelBox_(sheet, range);

    // v38/v75 핵심:
    // 도장/로고 원본 파일을 미리 원하는 출력 크기로 줄여둔다는 전제입니다.
    // 따라서 코드에서는 setWidth/setHeight로 사이즈를 다시 조정하지 않습니다.
    // 원본 Blob을 그대로 삽입하고, 지정 range 안에서 위치 offset만 맞춥니다.
    // 같은 실행 안에서 같은 이미지 파일을 여러 번 읽지 않도록 blobCache를 사용합니다.
    const blob = this.getImageBlob_(fileId);
    const img = sheet.insertImage(blob, range.getColumn(), range.getRow());

    const width = img.getWidth ? img.getWidth() : 0;
    const height = img.getHeight ? img.getHeight() : 0;

    let xOffset = 0;
    if (placement.hAlign === 'center') xOffset = Math.floor((box.width - width) / 2);
    if (placement.hAlign === 'right') xOffset = Math.floor(box.width - width);

    let yOffset = 0;
    if (placement.vAlign === 'middle') yOffset = Math.floor((box.height - height) / 2);
    if (placement.vAlign === 'bottom') yOffset = Math.floor(box.height - height);

    img
      .setAnchorCell(range.getCell(1, 1))
      .setAnchorCellXOffset(Math.max(0, xOffset))
      .setAnchorCellYOffset(Math.max(0, yOffset));

    const altVendorName = String(vendorNameForAlt || this.vendorName || '').trim();
    if (img.setAltTextTitle) img.setAltTextTitle(`AUTO_MAIL_ASSET:${placement.sheetName}:${placement.rangeA1}`);
    if (img.setAltTextDescription) img.setAltTextDescription(`${altVendorName} ${placement.role || 'stampFileId'}`);
  }
}



function stripKnownExtension_(filename) {
  return String(filename || '')
    .replace(/\.(pdf|xlsx|xls|docx|doc|zip)$/i, '')
    .trim() || '파일';
}

class ReviewFilePackageBuilder {
  constructor(generatorSs, targetData, progress) {
    this.ss = generatorSs;
    this.targetData = targetData || {};
    this.progress = progress || null;
    this.exporter = new ExportService(progress);
  }

  build(definitions, registration) {
    const defs = definitions || [];
    rewriteGeneratorRequestLogFixedRowRefsToTargetV89_(this.ss, this.progress, '파일 확인/수정 생성 전');
    assertNoSelectedGeneratorRequestLogRow3RefsV89_(this.ss, defs);
    ensureSelectedGeneratorSheetsShowCurrentCustomerV88_(this.ss, this.targetData, defs, this.progress, '파일 확인/수정 생성 전');
    const folder = this.createReviewFolder_(registration);
    const files = [];
    const perfCfg = CONFIG.REVIEW_FILE_GENERATION || {};

    const editableSheetDefs = [];
    const nonEditableReferenceDefs = [];
    const normalDefs = [];

    defs.forEach(def => {
      if (!def) return;
      if (perfCfg.COMBINE_EDITABLE_SHEETS_IN_REVIEW === true && this.isCombinedEditableSheetDefinition_(def)) {
        editableSheetDefs.push(def);
        return;
      }
      if (perfCfg.COLLAPSE_NON_EDITABLE_REFERENCES_TO_INDEX === true && this.isNonEditableReferenceDefinition_(def)) {
        nonEditableReferenceDefs.push(def);
        return;
      }
      normalDefs.push(def);
    });

    if (editableSheetDefs.length) {
      if (this.progress) this.progress.update(58, '수정용 통합시트 생성 중');
      this.buildCombinedEditableSheetFiles_(editableSheetDefs, folder, registration).forEach(function(fileInfo) {
        files.push(fileInfo);
      });
    }

    if (nonEditableReferenceDefs.length) {
      if (this.progress) this.progress.update(72, '비편집 자료 링크모음 생성 중');
      const indexInfo = this.buildNonEditableReferenceIndex_(nonEditableReferenceDefs, folder);
      if (indexInfo) files.push(indexInfo);
    }

    for (let i = 0; i < normalDefs.length; i++) {
      const def = normalDefs[i];
      if (this.progress) this.progress.update(76 + Math.floor((i / Math.max(normalDefs.length, 1)) * 12), '수정용 파일 생성 중: ' + def.label);
      this.buildDefinitionFiles_(def, folder).forEach(function(fileInfo) {
        files.push(fileInfo);
      });
    }

    return {
      folderId: folder.getId(),
      folderUrl: folder.getUrl(),
      files: files
    };
  }

  createReviewFolder_(registration) {
    const timestamp = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyyMMdd_HHmmss');
    const company = sanitizeFileName_(this.targetData['회사명'] || '고객사');
    const requestNo = registration && registration.requestNo ? registration.requestNo : '미접수';
    const name = '메일발송_파일확인수정_' + requestNo + '_' + company + '_' + timestamp;

    // v70: DriveApp.getRootFolder()는 일부 계정/공유드라이브/권한 정책에서 접근 거부가 납니다.
    // 수정용 폴더는 루트가 아니라, 설정 폴더 또는 생성기/마스터 파일이 들어있는 폴더 아래에 생성합니다.
    const parent = resolveMailAutoWritableParentFolder_(null);
    if (!parent) {
      throw new Error(
        '수정용 Drive 폴더를 만들 부모 폴더를 찾지 못했습니다.\n' +
        'CONFIG.TEMP.PARENT_FOLDER_ID에 현재 실행 계정이 편집 가능한 Google Drive 폴더 ID를 넣어주세요.'
      );
    }

    try {
      return parent.createFolder(name);
    } catch (err) {
      throw new Error(
        '수정용 Drive 폴더 생성 중 Drive 접근이 거부되었습니다.\n' +
        '현재 실행 계정이 아래 우선순위의 폴더에 파일 생성 권한을 갖고 있는지 확인하세요.\n' +
        '1) CONFIG.TEMP.PARENT_FOLDER_ID 폴더, 2) 생성기 스프레드시트의 부모 폴더, 3) 마스터 스프레드시트의 부모 폴더\n\n' +
        String(err && err.message || err)
      );
    }
  }

  buildDefinitionFiles_(def, folder) {
    const result = [];

    if (def.type === 'sheet_pdf') {
      const sheet = mustGetSheet_(this.ss, def.sheetName);
      if (def.key === 'quote' || def.sheetName === CONFIG.SHEETS.QUOTE || def.sheetName === '견적서') {
        this.targetData = applyContractPeriodToGeneratorSheetsV90_(this.ss, this.targetData, this.progress, '파일 확인/수정 견적서 생성 전');
        applyQuoteInspectionPeriodToSheetV90_(sheet, this.targetData);
        fixQuoteKoreanAmountErrors_(sheet);
        SpreadsheetApp.flush();
      }

      const sendFilename = this.renderFilename_(def.filename);
      const editName = stripKnownExtension_(sendFilename) + '_수정용';
      const file = this.createEditableSingleSheetSpreadsheet_(sheet, editName, folder);
      result.push({
        key: def.key,
        label: def.label,
        name: file.getName(),
        fileId: file.getId(),
        url: file.getUrl(),
        mimeType: MimeType.GOOGLE_SHEETS,
        exportAs: 'pdf_from_sheet',
        exportRange: def.exportRange,
        sheetName: def.sheetName,
        sendFilename: sendFilename,
        useForSend: true
      });
      return result;
    }

    if (def.type === 'sheet_xlsx_values') {
      const sheet = mustGetSheet_(this.ss, def.sheetName);
      if (def.key === 'serviceApplication' || def.sheetName === CONFIG.SHEETS.SERVICE_APP || def.sheetName === '용역신청서') {
        applyServiceApplicationApplicantFallbackToSheet_(sheet, this.targetData);
        applyServiceApplicationSpecialTermsToSheet_(sheet, this.targetData);
        SpreadsheetApp.flush();
      }

      const sendFilename = this.renderFilename_(def.filename);
      const editName = stripKnownExtension_(sendFilename) + '_수정용';
      const file = this.createEditableSingleSheetSpreadsheet_(sheet, editName, folder);
      result.push({
        key: def.key,
        label: def.label,
        name: file.getName(),
        fileId: file.getId(),
        url: file.getUrl(),
        mimeType: MimeType.GOOGLE_SHEETS,
        exportAs: 'xlsx_from_sheet',
        sheetName: def.sheetName,
        sendFilename: sendFilename,
        useForSend: true
      });
      return result;
    }

    if (def.type === 'docx_template') {
      const sendFilename = this.renderFilename_(def.filename);
      const editName = stripKnownExtension_(sendFilename) + '_수정용';
      const file = new DocxTemplateBuilder(this.targetData, this.ss).buildEditableGoogleDoc(def.templateFileId, editName, folder);
      result.push({
        key: def.key,
        label: def.label,
        name: file.getName(),
        fileId: file.getId(),
        url: file.getUrl(),
        mimeType: MimeType.GOOGLE_DOCS,
        exportAs: 'docx_from_doc',
        sendFilename: sendFilename,
        useForSend: true
      });
      return result;
    }

    if (def.type === 'multi_sheet_pdf') {
      // v77:
      // 파일 확인/수정 단계에서는 비교견적서를 PDF가 아니라 편집 가능한 Google Sheets로 생성합니다.
      // 단, 사용자가 수정 후 최종 발송할 때는 수정본 Sheets를 다시 PDF로 export합니다.
      (def.sheets || []).forEach(item => {
        const sheet = mustGetSheet_(this.ss, item.sheetName);
        const sendFilename = this.renderFilename_(item.filename);
        const editName = stripKnownExtension_(sendFilename) + '_수정용';
        const preserveFormulaRanges = this.getPreserveFormulaRangesForReviewSheet_(item.sheetName);
        const file = this.createEditableSingleSheetSpreadsheet_(sheet, editName, folder, preserveFormulaRanges);
        result.push({
          key: def.key,
          label: def.label,
          name: file.getName(),
          fileId: file.getId(),
          url: file.getUrl(),
          mimeType: MimeType.GOOGLE_SHEETS,
          exportAs: 'pdf_from_sheet',
          exportRange: item.exportRange,
          sheetName: item.sheetName,
          sendFilename: sendFilename,
          useForSend: true
        });
      });
      return result;
    }

    if (def.type === 'static_files') {
      // v79:
      // 안내문/법령요약 같은 비편집 고정 파일은 수정용 폴더에 실제 복사하지 않고 shortcut을 우선 생성합니다.
      // 발송 시에는 reviewPackage의 shortcut을 쓰지 않고 원본 파일을 그대로 첨부하므로 속도와 안정성을 모두 챙깁니다.
      (def.fileIds || []).forEach(fileId => {
        const src = DriveApp.getFileById(fileId);
        const file = this.createReviewShortcutOrCopy_(src, src.getName(), folder);
        result.push({
          key: def.key,
          label: def.label,
          name: file.getName(),
          fileId: file.getId(),
          url: file.getUrl(),
          mimeType: file.getMimeType(),
          exportAs: 'raw_file',
          sendFilename: src.getName(),
          useForSend: false
        });
      });
      return result;
    }

    if (def.type === 'vendor_docx_template') {
      const vendorName = String(this.targetData['수행사'] || '').trim();
      const vendor = getVendorConfig_(vendorName);
      const fieldName = String(def.vendorFileIdField || '').trim();
      const fileId = vendor && fieldName ? String(vendor[fieldName] || '').trim() : '';
      if (!vendor || !fileId) {
        throw new Error((def.label || '수행사별 DOCX 템플릿') + ' 파일 ID가 없습니다. 수행사=' + vendorName + ' / CONFIG.VENDORS의 ' + fieldName + '를 확인하세요.');
      }
      const sendFilename = def.filename
        ? this.renderFilename_(def.filename)
        : appendDateSuffixBeforeExtension_(sanitizeFileName_(vendorName || '수행사') + '_용역표준계약서.docx', getTodayFileDateSuffix_());
      const blob = new ServiceStandardContractDocxBuilder(this.targetData, this.ss).build(fileId, sendFilename);
      const file = folder.createFile(blob);
      result.push({
        key: def.key,
        label: def.label,
        name: file.getName(),
        fileId: file.getId(),
        url: file.getUrl(),
        mimeType: file.getMimeType(),
        exportAs: 'raw_file',
        sendFilename: sendFilename,
        useForSend: true
      });
      return result;
    }

    if (def.type === 'vendor_zip') {
      const vendorName = String(this.targetData['수행사'] || '').trim();
      const vendor = getVendorConfig_(vendorName);
      if (!vendor || !vendor.contractorInfoZipFileId) {
        throw new Error('수행사정보 파일 ID가 없습니다. 수행사=' + vendorName + ' / CONFIG.VENDORS의 contractorInfoZipFileId를 확인하세요.');
      }
      const src = DriveApp.getFileById(vendor.contractorInfoZipFileId);
      const copyName = def.filename ? this.renderFilename_(def.filename) : src.getName();
      const file = this.createReviewShortcutOrCopy_(src, copyName, folder);
      result.push({
        key: def.key,
        label: def.label,
        name: file.getName(),
        fileId: file.getId(),
        url: file.getUrl(),
        mimeType: file.getMimeType(),
        exportAs: 'raw_file',
        sendFilename: copyName,
        useForSend: false
      });
      return result;
    }

    if (def.type === 'drive_folder_files') {
      const folderId = String(def.folderId || '').trim();
      if (!folderId) throw new Error((def.label || '폴더 첨부') + ' 폴더 ID가 비어 있습니다.');
      const srcFolder = DriveApp.getFolderById(folderId);
      const iter = srcFolder.getFiles();
      while (iter.hasNext()) {
        const src = iter.next();
        const file = this.createReviewShortcutOrCopy_(src, src.getName(), folder);
        result.push({
          key: def.key,
          label: def.label,
          name: file.getName(),
          fileId: file.getId(),
          url: file.getUrl(),
          mimeType: file.getMimeType(),
          exportAs: 'raw_file',
          sendFilename: src.getName(),
          useForSend: false
        });
      }
      return result;
    }

    return result;
  }



  isCombinedEditableSheetDefinition_(def) {
    if (!def) return false;
    return def.type === 'sheet_pdf' || def.type === 'sheet_xlsx_values' || def.type === 'multi_sheet_pdf';
  }

  isNonEditableReferenceDefinition_(def) {
    if (!def) return false;
    // 발송 시 사용자가 편집한 수정본을 다시 읽을 필요가 없는 자료만 링크모음으로 축약합니다.
    // vendor_docx_template은 생성된 DOCX 자체가 발송 대상이므로 여기서 축약하지 않습니다.
    return def.type === 'static_files' || def.type === 'vendor_zip' || def.type === 'drive_folder_files';
  }

  buildCombinedEditableSheetFiles_(definitions, folder, registration) {
    const records = this.collectCombinedEditableSheetRecords_(definitions);
    if (!records.length) return [];

    // 서비스신청서 특약사항/견적서 한글 금액 보정처럼, 시트 값을 읽기 전에 반드시 반영되어야 하는 보정만 선처리합니다.
    records.forEach(record => {
      this.prepareReviewSourceSheet_(record);
    });
    SpreadsheetApp.flush();

    const company = sanitizeFileName_(this.targetData['회사명'] || '고객사');
    const requestNo = registration && registration.requestNo ? String(registration.requestNo) : '미접수';
    const suffix = String((CONFIG.REVIEW_FILE_GENERATION && CONFIG.REVIEW_FILE_GENERATION.COMBINED_SHEETS_FILE_SUFFIX) || '_수정용_통합시트');
    const fileName = company + '_' + requestNo + suffix;
    const file = this.createCombinedEditableSpreadsheet_(records, fileName, folder);

    return records.map(record => {
      return {
        key: record.key,
        label: record.label,
        name: file.getName() + ' - ' + record.editSheetTitle,
        fileId: file.getId(),
        url: file.getUrl(),
        mimeType: MimeType.GOOGLE_SHEETS,
        exportAs: record.exportAs,
        exportRange: record.exportRange,
        sheetName: record.editSheetTitle,
        originalSheetName: record.sourceSheetName,
        sendFilename: record.sendFilename,
        useForSend: true,
        combinedReviewWorkbook: true
      };
    });
  }

  collectCombinedEditableSheetRecords_(definitions) {
    const records = [];
    const titleSeen = {};

    const makeTitle = function(base) {
      let clean = String(base || '수정용').trim().replace(/[\\/\?\*\[\]:]/g, ' ').replace(/\s+/g, ' ');
      if (!clean) clean = '수정용';
      if (clean.length > 80) clean = clean.slice(0, 80).trim();
      let title = clean;
      let idx = 2;
      while (titleSeen[title.toLowerCase()]) {
        const suffix = '_' + idx;
        title = clean.slice(0, Math.max(1, 80 - suffix.length)) + suffix;
        idx++;
      }
      titleSeen[title.toLowerCase()] = true;
      return title;
    };

    (definitions || []).forEach(def => {
      if (!def) return;

      if (def.type === 'sheet_pdf') {
        const sheet = mustGetSheet_(this.ss, def.sheetName);
        const sendFilename = this.renderFilename_(def.filename);
        records.push({
          key: def.key,
          label: def.label,
          def: def,
          sourceSheet: sheet,
          sourceSheetName: def.sheetName,
          editSheetTitle: makeTitle(def.label || def.sheetName || '견적서'),
          exportAs: 'pdf_from_sheet',
          exportRange: def.exportRange,
          sendFilename: sendFilename,
          preserveFormulaRanges: []
        });
        return;
      }

      if (def.type === 'sheet_xlsx_values') {
        const sheet = mustGetSheet_(this.ss, def.sheetName);
        const sendFilename = this.renderFilename_(def.filename);
        records.push({
          key: def.key,
          label: def.label,
          def: def,
          sourceSheet: sheet,
          sourceSheetName: def.sheetName,
          editSheetTitle: makeTitle(def.label || def.sheetName || '용역신청서'),
          exportAs: 'xlsx_from_sheet_single',
          exportRange: def.exportRange || '',
          sendFilename: sendFilename,
          preserveFormulaRanges: []
        });
        return;
      }

      if (def.type === 'multi_sheet_pdf') {
        (def.sheets || []).forEach(item => {
          const sheet = mustGetSheet_(this.ss, item.sheetName);
          const sendFilename = this.renderFilename_(item.filename);
          const label = compareQuoteSheetNameToLabel_(item.sheetName);
          records.push({
            key: def.key,
            label: label,
            def: def,
            sourceSheet: sheet,
            sourceSheetName: item.sheetName,
            editSheetTitle: makeTitle(label || item.sheetName || '비교견적서'),
            exportAs: 'pdf_from_sheet',
            exportRange: item.exportRange,
            sendFilename: sendFilename,
            preserveFormulaRanges: this.getPreserveFormulaRangesForReviewSheet_(item.sheetName)
          });
        });
      }
    });

    return records;
  }

  prepareReviewSourceSheet_(record) {
    if (!record || !record.sourceSheet) return;
    const def = record.def || {};
    const sheet = record.sourceSheet;

    if (def.key === 'quote' || record.sourceSheetName === CONFIG.SHEETS.QUOTE || record.sourceSheetName === '견적서') {
      this.targetData = applyContractPeriodToGeneratorSheetsV90_(this.ss, this.targetData, this.progress, '통합 수정용 견적서 생성 전');
      applyQuoteInspectionPeriodToSheetV90_(sheet, this.targetData);
      fixQuoteKoreanAmountErrors_(sheet);
    }

    if (def.key === 'serviceApplication' || record.sourceSheetName === CONFIG.SHEETS.SERVICE_APP || record.sourceSheetName === '용역신청서') {
      applyServiceApplicationApplicantFallbackToSheet_(sheet, this.targetData);
      applyServiceApplicationSpecialTermsToSheet_(sheet, this.targetData);
    }
  }

  createCombinedEditableSpreadsheet_(records, fileName, folder) {
    // v81:
    // v80의 통합 수정용 시트 생성은 Sheets REST API copyTo/batchUpdate를 사용했는데,
    // 일부 Apps Script 프로젝트에서 Google Sheets API가 비활성화되어 있으면 403으로 실패합니다.
    // 이 프로젝트는 별도 API 활성화 없이 바로 동작해야 하므로, 통합 시트 생성은 Apps Script 내장 SpreadsheetApp 경로를 기본값으로 사용합니다.
    // 파일 수를 여러 개 만들지 않는 v80 구조는 유지하되, Sheets API 호출 자체를 제거했습니다.
    return this.createCombinedEditableSpreadsheetBySpreadsheetApp_(records, fileName, folder);
  }

  createCombinedEditableSpreadsheetBySpreadsheetApp_(records, fileName, folder) {
    SpreadsheetApp.flush();

    const file = this.createBlankSpreadsheetFileInFolder_(fileName, folder);
    const tempId = file.getId();

    try {
      const tempSs = SpreadsheetApp.openById(tempId);
      const defaultSheets = tempSs.getSheets();
      const copiedSheets = [];

      (records || []).forEach(record => {
        if (!record || !record.sourceSheet) return;

        const sourceSheet = record.sourceSheet;
        const lastRow = Math.max(sourceSheet.getLastRow(), 1);
        const lastCol = Math.max(sourceSheet.getLastColumn(), 1);

        // copyTo는 형식/열너비/병합/이미지/도장을 보존합니다.
        // 이후 표시값만 덮어써서 수식 대부분을 값으로 고정하고, 지정 범위 수식만 다시 복원합니다.
        const copied = sourceSheet.copyTo(tempSs);
        copied.setName(record.editSheetTitle || sourceSheet.getName());
        copiedSheets.push(copied);

        const values = sourceSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();
        copied.getRange(1, 1, lastRow, lastCol).setValues(values);

        if (record.preserveFormulaRanges && record.preserveFormulaRanges.length) {
          this.restoreSelectedFormulas_(sourceSheet, copied, record.preserveFormulaRanges);
        }
      });

      // 복사 시트가 하나라도 있으면 기본 빈 시트 삭제.
      // 처음부터 삭제하면 마지막 시트 삭제 오류가 날 수 있으므로 복사 후 제거합니다.
      if (copiedSheets.length) {
        tempSs.getSheets().forEach(function(sh) {
          const keep = copiedSheets.some(function(copied) {
            return copied.getSheetId() === sh.getSheetId();
          });
          if (!keep && tempSs.getSheets().length > 1) {
            tempSs.deleteSheet(sh);
          }
        });
        tempSs.setActiveSheet(copiedSheets[0]);
      }

      SpreadsheetApp.flush();
      this.moveFileToReviewFolder_(file, folder);
      return file;
    } catch (err) {
      try { file.setTrashed(true); } catch (e) {}
      throw new Error('통합 수정용 시트 생성 실패: ' + (err && err.message || err));
    }
  }

  createBlankSpreadsheetFileInFolder_(fileName, folder) {
    return createBlankSpreadsheetFileInMailAutoWorkspace_(fileName, folder);
  }

  copySheetToSpreadsheetByApi_(sourceSheet, destinationSpreadsheetId) {
    const sourceSsId = sourceSheet.getParent().getId();
    const sourceSheetId = sourceSheet.getSheetId();
    const copyUrl =
      'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(sourceSsId) +
      '/sheets/' + encodeURIComponent(sourceSheetId) + ':copyTo';

    const copyRes = this.fetchSheetsApiWithRetry_(copyUrl, {
      method: 'post',
      contentType: 'application/json; charset=utf-8',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify({ destinationSpreadsheetId: destinationSpreadsheetId }),
      muteHttpExceptions: true
    }, '통합 수정용 시트 복사');

    const copiedData = JSON.parse(copyRes.getContentText() || '{}');
    const copiedSheetId = Number(copiedData.sheetId || (copiedData.properties && copiedData.properties.sheetId));
    if (!copiedSheetId) {
      throw new Error('Sheets API copyTo 응답에서 복사된 sheetId를 확인하지 못했습니다: ' + copyRes.getContentText().slice(0, 500));
    }
    return copiedSheetId;
  }

  collectFormulaRestoreData_(sourceSheet, targetSheetName, preserveFormulaRanges) {
    const data = [];
    const ranges = unique_((preserveFormulaRanges || [])
      .map(function(a1) { return String(a1 || '').trim(); })
      .filter(Boolean));

    ranges.forEach(function(a1) {
      const srcRange = sourceSheet.getRange(a1);
      const startRow = srcRange.getRow();
      const startCol = srcRange.getColumn();
      const formulas = srcRange.getFormulas();

      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const formula = String(formulas[r][c] || '').trim();
          if (!formula) continue;
          const cellA1 = columnToLetter_(startCol + c) + (startRow + r);
          data.push({
            range: quoteSheetNameForA1_(targetSheetName) + '!' + cellA1,
            values: [[formula]]
          });
        }
      }
    });

    return data;
  }

  buildNonEditableReferenceIndex_(definitions, folder) {
    const links = [];
    (definitions || []).forEach(def => {
      this.collectNonEditableReferenceLinks_(def).forEach(function(item) {
        links.push(item);
      });
    });

    if (!links.length) return null;

    const filename = String((CONFIG.REVIEW_FILE_GENERATION && CONFIG.REVIEW_FILE_GENERATION.NON_EDITABLE_INDEX_FILENAME) || '비편집_자료_링크모음.html');
    const company = escapeHtml_(this.targetData['회사명'] || '고객사');
    const rows = links.map(function(item, idx) {
      return '<tr>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;color:#6b7280;white-space:nowrap;">' + (idx + 1) + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;white-space:nowrap;">' + escapeHtml_(item.group || '') + '</td>' +
        '<td style="padding:8px 10px;border-bottom:1px solid #e5e7eb;"><a href="' + escapeHtml_(item.url || '') + '" target="_blank" rel="noopener" style="color:#1a73e8;text-decoration:none;font-weight:700;">' + escapeHtml_(item.name || '파일') + '</a></td>' +
      '</tr>';
    }).join('');

    const html = '<!doctype html><html><head><meta charset="UTF-8"><title>비편집 자료 링크모음</title></head>' +
      '<body style="font-family:Malgun Gothic, Arial, sans-serif;margin:24px;color:#111827;">' +
      '<h2 style="margin:0 0 8px;font-size:20px;">비편집 자료 링크모음</h2>' +
      '<p style="margin:0 0 18px;color:#4b5563;font-size:13px;line-height:1.6;">' + company + ' 발송 패키지 중 수정 대상이 아닌 자료입니다. 발송 시에는 원본 파일 기준으로 첨부 또는 Drive 링크 정책이 적용됩니다.</p>' +
      '<table style="border-collapse:collapse;width:100%;font-size:13px;">' +
      '<thead><tr style="background:#f3f4f6;text-align:left;">' +
      '<th style="padding:8px 10px;border-bottom:1px solid #d1d5db;width:46px;">No</th>' +
      '<th style="padding:8px 10px;border-bottom:1px solid #d1d5db;width:140px;">구분</th>' +
      '<th style="padding:8px 10px;border-bottom:1px solid #d1d5db;">파일명</th>' +
      '</tr></thead><tbody>' + rows + '</tbody></table></body></html>';

    const file = folder.createFile(Utilities.newBlob(html, 'text/html', filename));
    return {
      key: '__nonEditableReferenceIndex',
      label: '비편집 자료 링크모음',
      name: file.getName(),
      fileId: file.getId(),
      url: file.getUrl(),
      mimeType: file.getMimeType(),
      exportAs: 'reference_index',
      sendFilename: filename,
      useForSend: false
    };
  }

  collectNonEditableReferenceLinks_(def) {
    const result = [];
    if (!def) return result;

    if (def.type === 'static_files') {
      (def.fileIds || []).forEach(fileId => {
        try {
          const file = DriveApp.getFileById(fileId);
          result.push({ group: def.label || '고정자료', name: file.getName(), url: file.getUrl(), fileId: file.getId() });
        } catch (err) {
          Logger.log('비편집 자료 링크 수집 실패: ' + fileId + ' / ' + err);
        }
      });
      return result;
    }

    if (def.type === 'vendor_zip') {
      const vendorName = String(this.targetData['수행사'] || '').trim();
      const vendor = getVendorConfig_(vendorName);
      const fileId = vendor && vendor.contractorInfoZipFileId ? String(vendor.contractorInfoZipFileId).trim() : '';
      if (fileId) {
        try {
          const file = DriveApp.getFileById(fileId);
          const displayName = def.filename ? this.renderFilename_(def.filename) : file.getName();
          result.push({ group: def.label || '수행사정보', name: displayName, url: file.getUrl(), fileId: file.getId() });
        } catch (err) {
          Logger.log('수행사정보 링크 수집 실패: ' + fileId + ' / ' + err);
        }
      }
      return result;
    }

    if (def.type === 'drive_folder_files') {
      const folderId = String(def.folderId || '').trim();
      if (!folderId) return result;
      try {
        const srcFolder = DriveApp.getFolderById(folderId);
        const files = [];
        const iter = srcFolder.getFiles();
        while (iter.hasNext()) files.push(iter.next());
        files.sort((a, b) => String(a.getName() || '').localeCompare(String(b.getName() || ''), 'ko'));
        files.forEach(file => {
          result.push({ group: def.label || '폴더자료', name: file.getName(), url: file.getUrl(), fileId: file.getId() });
        });
      } catch (err) {
        Logger.log('폴더자료 링크 수집 실패: ' + folderId + ' / ' + err);
      }
      return result;
    }

    return result;
  }

  createReviewShortcutOrCopy_(srcFile, displayName, folder) {
    const name = String(displayName || (srcFile && srcFile.getName ? srcFile.getName() : '') || '자료 파일').trim();
    const reviewCfg = CONFIG.REVIEW_FILE_GENERATION || {};

    // v82:
    // 안내문/수행사정보/샘플보고서 등 비편집 자료는 사용자가 수정할 파일이 아닙니다.
    // 매번 원본 파일 기준 shortcut을 새로 만들면 Drive API 호출이 파일 수만큼 반복됩니다.
    // 공용 캐시 폴더에 원본별 shortcut을 1회 만들어두고, 이후에는 그 shortcut을 현재 수정용 폴더로 복사합니다.
    if (reviewCfg.USE_CACHED_SHORTCUTS_FOR_NON_EDITABLE_REFERENCES === true) {
      try {
        const cachedShortcut = this.getOrCreateCachedReviewShortcut_(srcFile, name);
        if (cachedShortcut) {
          return cachedShortcut.makeCopy(name, folder);
        }
      } catch (cacheErr) {
        Logger.log('비편집 자료 shortcut 캐시 사용 실패. 직접 shortcut 생성으로 fallback: ' + (cacheErr && cacheErr.stack || cacheErr));
      }
    }

    return this.createDriveShortcutOrCopy_(srcFile, name, folder);
  }

  getOrCreateCachedReviewShortcut_(srcFile, displayName) {
    const folder = this.getReviewShortcutCacheFolder_();
    if (!folder) return null;

    const props = PropertiesService.getScriptProperties();
    const cfg = CONFIG.REVIEW_FILE_GENERATION || {};
    const cacheVersion = String(cfg.SHORTCUT_CACHE_VERSION || 'v1');
    const targetId = srcFile.getId();
    const name = String(displayName || srcFile.getName() || '자료 파일').trim();
    const key = 'MAILAUTO_REVIEW_SHORTCUT_CACHE_' + sha1Hex_(cacheVersion + '|' + targetId + '|' + name);
    const cachedId = String(props.getProperty(key) || '').trim();

    if (cachedId) {
      try {
        const cachedFile = DriveApp.getFileById(cachedId);
        // 접근 가능한 캐시 shortcut이면 그대로 재사용합니다.
        cachedFile.getName();
        return cachedFile;
      } catch (err) {
        props.deleteProperty(key);
      }
    }

    const cachedShortcut = this.createDriveShortcutOrCopy_(srcFile, name, folder);
    props.setProperty(key, cachedShortcut.getId());
    return cachedShortcut;
  }

  getReviewShortcutCacheFolder_() {
    const props = PropertiesService.getScriptProperties();
    const cfg = CONFIG.REVIEW_FILE_GENERATION || {};
    const cacheVersion = String(cfg.SHORTCUT_CACHE_VERSION || 'v1');
    const key = 'MAILAUTO_REVIEW_SHORTCUT_CACHE_FOLDER_' + sha1Hex_(cacheVersion + '|' + String(cfg.SHORTCUT_CACHE_FOLDER_NAME || ''));
    const cachedFolderId = String(props.getProperty(key) || '').trim();

    if (cachedFolderId) {
      try {
        const cachedFolder = DriveApp.getFolderById(cachedFolderId);
        cachedFolder.getName();
        return cachedFolder;
      } catch (err) {
        props.deleteProperty(key);
      }
    }

    const parent = resolveMailAutoWritableParentFolder_(null);
    if (!parent) return null;

    const folderName = String(cfg.SHORTCUT_CACHE_FOLDER_NAME || '메일자동화_비편집자료_바로가기캐시').trim();
    const folder = parent.createFolder(folderName);
    props.setProperty(key, folder.getId());
    return folder;
  }

  createDriveShortcutOrCopy_(srcFile, displayName, folder) {
    const name = String(displayName || (srcFile && srcFile.getName ? srcFile.getName() : '') || '자료 파일').trim();
    const targetId = srcFile.getId();
    const parentId = folder.getId();

    // DriveApp에는 환경별로 shortcut 생성 메서드 지원 차이가 있어 Drive REST API를 우선 사용합니다.
    // 실패하면 기존 방식처럼 실제 복사본을 생성합니다.
    try {
      const url = 'https://www.googleapis.com/drive/v3/files?supportsAllDrives=true';
      const res = UrlFetchApp.fetch(url, {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({
          name: name,
          mimeType: 'application/vnd.google-apps.shortcut',
          parents: [parentId],
          shortcutDetails: { targetId: targetId }
        }),
        muteHttpExceptions: true
      });
      const code = res.getResponseCode();
      if (code >= 200 && code < 300) {
        const data = JSON.parse(res.getContentText() || '{}');
        if (data && data.id) return DriveApp.getFileById(data.id);
      }
      Logger.log('shortcut 생성 실패. 복사로 fallback: HTTP ' + code + ' / ' + res.getContentText().slice(0, 500));
    } catch (err) {
      Logger.log('shortcut 생성 예외. 복사로 fallback: ' + (err && err.stack || err));
    }

    return srcFile.makeCopy(name, folder);
  }

  createEditableSingleSheetSpreadsheet_(sourceSheet, fileName, folder, preserveFormulaRanges) {
    SpreadsheetApp.flush();
    const lastRow = Math.max(sourceSheet.getLastRow(), 1);
    const lastCol = Math.max(sourceSheet.getLastColumn(), 1);
    const displayValues = sourceSheet.getRange(1, 1, lastRow, lastCol).getDisplayValues();

    const reviewCfg = CONFIG.REVIEW_FILE_GENERATION || {};
    if (reviewCfg.USE_SHEETS_API_FAST_PATH_FOR_REVIEW_COPY !== true) {
      return this.createEditableSingleSheetSpreadsheetBySpreadsheetApp_(sourceSheet, fileName, folder, displayValues, lastRow, lastCol, preserveFormulaRanges);
    }

    // 선택적으로만 Sheets API 고속 경로를 사용합니다.
    // API 비활성 프로젝트에서는 403 실패를 매 파일마다 기다리는 것 자체가 병목이므로 기본값은 false입니다.
    try {
      return this.createEditableSingleSheetSpreadsheetBySheetsApi_(sourceSheet, fileName, folder, displayValues, lastRow, lastCol, preserveFormulaRanges);
    } catch (apiErr) {
      Logger.log('수정용 Sheets API 고속 생성 실패. 기존 SpreadsheetApp 방식으로 fallback: ' + (apiErr && apiErr.stack || apiErr));
      return this.createEditableSingleSheetSpreadsheetBySpreadsheetApp_(sourceSheet, fileName, folder, displayValues, lastRow, lastCol, preserveFormulaRanges);
    }
  }

  createEditableSingleSheetSpreadsheetBySheetsApi_(sourceSheet, fileName, folder, displayValues, lastRow, lastCol, preserveFormulaRanges) {
    const file = createBlankSpreadsheetFileInMailAutoWorkspace_(fileName, folder);
    const tempId = file.getId();
    const temp = SpreadsheetApp.openById(tempId);

    try {
      const defaultSheets = temp.getSheets();
      const defaultSheetId = defaultSheets && defaultSheets.length ? defaultSheets[0].getSheetId() : 0;
      const sourceSsId = sourceSheet.getParent().getId();
      const sourceSheetId = sourceSheet.getSheetId();

      const copyUrl =
        'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(sourceSsId) +
        '/sheets/' + encodeURIComponent(sourceSheetId) + ':copyTo';

      const copyRes = this.fetchSheetsApiWithRetry_(copyUrl, {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ destinationSpreadsheetId: tempId }),
        muteHttpExceptions: true
      }, '수정용 시트 고속 복사');

      const copiedData = JSON.parse(copyRes.getContentText() || '{}');
      const copiedSheetId = Number(copiedData.sheetId || (copiedData.properties && copiedData.properties.sheetId));
      if (!copiedSheetId) {
        throw new Error('Sheets API copyTo 응답에서 복사된 sheetId를 확인하지 못했습니다: ' + copyRes.getContentText().slice(0, 500));
      }

      const requests = [];
      if (defaultSheetId && defaultSheetId !== copiedSheetId) {
        requests.push({ deleteSheet: { sheetId: defaultSheetId } });
      }
      requests.push({
        updateSheetProperties: {
          properties: { sheetId: copiedSheetId, title: sourceSheet.getName() },
          fields: 'title'
        }
      });

      const batchUrl =
        'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(tempId) + ':batchUpdate';
      this.fetchSheetsApiWithRetry_(batchUrl, {
        method: 'post',
        contentType: 'application/json; charset=utf-8',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ requests: requests }),
        muteHttpExceptions: true
      }, '수정용 시트 정리');

      const rangeA1 = quoteSheetNameForA1_(sourceSheet.getName()) + '!A1:' + columnToLetter_(lastCol) + lastRow;
      const valuesUrl =
        'https://sheets.googleapis.com/v4/spreadsheets/' + encodeURIComponent(tempId) +
        '/values/' + encodeURIComponent(rangeA1) + '?valueInputOption=RAW';
      this.fetchSheetsApiWithRetry_(valuesUrl, {
        method: 'put',
        contentType: 'application/json; charset=utf-8',
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        payload: JSON.stringify({ values: displayValues }),
        muteHttpExceptions: true
      }, '수정용 시트 값 고정');

      if (preserveFormulaRanges && preserveFormulaRanges.length) {
        const tempSsForFormula = SpreadsheetApp.openById(tempId);
        const targetSheet = tempSsForFormula.getSheetByName(sourceSheet.getName()) || tempSsForFormula.getSheets()[0];
        this.restoreSelectedFormulas_(sourceSheet, targetSheet, preserveFormulaRanges);
        SpreadsheetApp.flush();
      }

      this.moveFileToReviewFolder_(file, folder);
      return file;
    } catch (err) {
      try { file.setTrashed(true); } catch (e) {}
      throw err;
    }
  }

  createEditableSingleSheetSpreadsheetBySpreadsheetApp_(sourceSheet, fileName, folder, displayValues, lastRow, lastCol, preserveFormulaRanges) {
    const file = createBlankSpreadsheetFileInMailAutoWorkspace_(fileName, folder);
    const tempId = file.getId();
    const temp = SpreadsheetApp.openById(tempId);

    try {
      const copied = sourceSheet.copyTo(temp).setName(sourceSheet.getName());
      temp.setActiveSheet(copied);
      temp.getSheets().forEach(function(sh) {
        if (sh.getSheetId() !== copied.getSheetId()) temp.deleteSheet(sh);
      });
      copied.getRange(1, 1, lastRow, lastCol).setValues(displayValues);
      if (preserveFormulaRanges && preserveFormulaRanges.length) {
        this.restoreSelectedFormulas_(sourceSheet, copied, preserveFormulaRanges);
      }
      SpreadsheetApp.flush();
      this.moveFileToReviewFolder_(file, folder);
      return file;
    } catch (err) {
      try { file.setTrashed(true); } catch (e) {}
      throw err;
    }
  }

  getPreserveFormulaRangesForReviewSheet_(sheetName) {
    const name = String(sheetName || '').trim();

    // v77:
    // 비교견적 파일 확인/수정용 시트는 대부분 값을 고정하되,
    // 사용자가 수정하면서 자동 재계산해야 하는 지정 구간의 수식만 보존합니다.
    if (name === '비교견적(1)') {
      return ['S12:W17', 'K18:AD18', 'H7:I7'];
    }

    if (name === '비교견적(2)') {
      return ['S12:W17', 'L18:AD19'];
    }

    return [];
  }

  restoreSelectedFormulas_(sourceSheet, targetSheet, preserveFormulaRanges) {
    const ranges = unique_((preserveFormulaRanges || [])
      .map(function(a1) { return String(a1 || '').trim(); })
      .filter(Boolean));

    ranges.forEach(function(a1) {
      const srcRange = sourceSheet.getRange(a1);
      const startRow = srcRange.getRow();
      const startCol = srcRange.getColumn();
      const formulas = srcRange.getFormulas();

      for (let r = 0; r < formulas.length; r++) {
        for (let c = 0; c < formulas[r].length; c++) {
          const formula = String(formulas[r][c] || '').trim();
          if (!formula) continue;
          targetSheet.getRange(startRow + r, startCol + c).setFormula(formula);
        }
      }
    });
  }

  moveFileToReviewFolder_(file, folder) {
    if (!file || !folder) return;
    try {
      file.moveTo(folder);
    } catch (moveErr) {
      // 일부 공유드라이브/Workspace 환경에서는 moveTo가 막힐 수 있습니다.
      // 그 경우 폴더에 추가만 시도하고, 실패 시 원래 Drive 위치의 파일 URL이라도 반환합니다.
      try {
        folder.addFile(file);
      } catch (addErr) {
        Logger.log('수정용 스프레드시트 폴더 이동/추가 실패: ' + moveErr + ' / ' + addErr);
      }
    }
  }

  fetchSheetsApiWithRetry_(url, options, label) {
    const waits = [300, 700, 1200, 2500, 5000];
    let lastCode = '';
    let lastText = '';

    for (let i = 0; i < waits.length; i++) {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const text = res.getContentText();

      if (code >= 200 && code < 300) {
        return res;
      }

      lastCode = code;
      lastText = text;

      if (![429, 500, 502, 503, 504].includes(code)) break;
      Utilities.sleep(waits[i]);
    }

    throw new Error((label || 'Sheets API 호출') + ' 실패 HTTP ' + lastCode + ': ' + String(lastText || '').slice(0, 800));
  }

  renderFilename_(template) {
    const company = sanitizeFileName_(this.targetData['회사명'] || '고객사');
    const vendor = sanitizeFileName_(this.targetData['수행사'] || '수행사');
    const rendered = String(template || 'attachment')
      .replace(/\{company\}/g, company)
      .replace(/\{vendor\}/g, vendor)
      .replace(/\{date\}/g, getTodayFileDateSuffix_());
    return appendDateSuffixBeforeExtension_(rendered, getTodayFileDateSuffix_());
  }
}

class AttachmentBuilder {
  constructor(generatorSs, targetData, progress, reviewPackage) {
    this.ss = generatorSs;
    this.targetData = targetData;
    this.progress = progress || null;
    this.exporter = new ExportService(progress);
    this.reviewPackage = reviewPackage || null;
    this.reviewFilesByKey = this.buildReviewFilesByKey_(reviewPackage);
  }
  hasReviewFileForKey_(key) {
    const k = String(key || '').trim();
    if (!k) return false;

    const list = this.reviewFilesByKey && this.reviewFilesByKey[k];
    return Array.isArray(list) && list.length > 0;
  }

  requiresLiveGeneratorSheetCheck_(def) {
    if (!def) return false;

    // 파일 확인/수정 후 발송에서는 수정본 Drive 파일을 바로 export해야 하므로
    // 해당 key의 review 파일이 있으면 원본 생성기 시트 고객명 검사를 하지 않습니다.
    if (this.hasReviewFileForKey_(def.key)) return false;

    return (
      def.type === 'sheet_pdf' ||
      def.type === 'sheet_xlsx_values' ||
      def.type === 'multi_sheet_pdf'
    );
  }
  build(definitions) {
    const blobs = [];
    const defs = definitions || [];

    // 파일 확인/수정 후 발송에서는 견적서/용역신청서/비교견적서 등이
    // reviewPackage 안의 수정본 Drive 파일로 이미 존재합니다.
    // 이 경우 원본 파일생성기 시트에 현재 고객명이 반영되어 있는지 검사하면
    // 정상 수정본 발송도 "생성기 양식에 현재 고객명이 반영되지 않았습니다"로 막힙니다.
    const liveGeneratorCheckDefs = defs.filter(def => this.requiresLiveGeneratorSheetCheck_(def));

    if (liveGeneratorCheckDefs.length) {
      rewriteGeneratorRequestLogFixedRowRefsToTargetV89_(this.ss, this.progress, '첨부파일 생성 전');
      assertNoSelectedGeneratorRequestLogRow3RefsV89_(this.ss, liveGeneratorCheckDefs);
      ensureSelectedGeneratorSheetsShowCurrentCustomerV88_(
        this.ss,
        this.targetData,
        liveGeneratorCheckDefs,
        this.progress,
        '첨부파일 생성 전'
      );
    }

    const linkSampleAndContractor = this.shouldSendSampleAndContractorAsDriveLinks_(defs);

    if (linkSampleAndContractor) {
      Logger.log('샘플보고서와 수행사정보가 동시에 선택되어 두 자료를 첨부 대신 Google Drive 링크로 전달합니다.');
    }

    for (let i = 0; i < defs.length; i++) {
      const def = defs[i];
      if (this.progress) this.progress.update(52 + Math.floor((i / Math.max(defs.length, 1)) * 20), '첨부 생성 중: ' + def.label);

      const reviewFile = this.popReviewFile_(def.key);
      if (reviewFile) {
        const reviewBlobs = this.buildBlobFromReviewFile_(def, reviewFile);
        reviewBlobs.forEach(function(blob) { blobs.push(blob); });
        continue;
      }

      if (def.type === 'sheet_pdf') {
        const sheet = mustGetSheet_(this.ss, def.sheetName);

        if (def.key === 'quote' || def.sheetName === CONFIG.SHEETS.QUOTE || def.sheetName === '견적서') {
          this.targetData = applyContractPeriodToGeneratorSheetsV90_(this.ss, this.targetData, this.progress, '첨부 견적서 생성 전');
          applyQuoteInspectionPeriodToSheetV90_(sheet, this.targetData);
          fixQuoteKoreanAmountErrors_(sheet);
          SpreadsheetApp.flush();
        }

        blobs.push(this.exporter.exportSheetPdf(this.ss, sheet, def.exportRange, this.renderFilename_(def.filename)));
      } else if (def.type === 'multi_sheet_pdf') {
        for (const item of def.sheets) {
          const sheet = mustGetSheet_(this.ss, item.sheetName);
          blobs.push(this.exporter.exportSheetPdf(this.ss, sheet, item.exportRange, this.renderFilename_(item.filename)));
        }
      } else if (def.type === 'sheet_xlsx_values') {
        const sheet = mustGetSheet_(this.ss, def.sheetName);

        if (def.key === 'serviceApplication' || def.sheetName === CONFIG.SHEETS.SERVICE_APP || def.sheetName === '용역신청서') {
          applyServiceApplicationApplicantFallbackToSheet_(sheet, this.targetData);
          applyServiceApplicationSpecialTermsToSheet_(sheet, this.targetData);
          SpreadsheetApp.flush();
        }

        blobs.push(this.exporter.exportSingleSheetXlsxValues(sheet, this.renderFilename_(def.filename)));
      } else if (def.type === 'docx_template') {
        blobs.push(new DocxTemplateBuilder(this.targetData, this.ss).build(def.templateFileId, this.renderFilename_(def.filename)));
      } else if (def.type === 'static_files') {
        for (const fileId of (def.fileIds || [])) {
          blobs.push(DriveApp.getFileById(fileId).getBlob());
        }
      } else if (def.type === 'drive_folder_files') {
        const folderId = String(def.folderId || '').trim();
        if (!folderId) {
          throw new Error((def.label || '폴더 첨부') + ' 폴더 ID가 비어 있습니다.');
        }

        const folder = DriveApp.getFolderById(folderId);
        const files = [];
        const iter = folder.getFiles();
        while (iter.hasNext()) files.push(iter.next());

        files.sort((a, b) => String(a.getName() || '').localeCompare(String(b.getName() || ''), 'ko'));

        if (!files.length) {
          throw new Error((def.label || '폴더 첨부') + ' 폴더 안에 첨부할 파일이 없습니다. folderId=' + folderId);
        }

        files.forEach(file => {
          if (linkSampleAndContractor && def.key === 'sampleReport') {
            this.addDriveDownloadLink_(file, def, { source: '샘플보고서' });
            return;
          }

          if (this.shouldLinkInsteadOfAttach_(file, def)) {
            this.addDriveDownloadLink_(file, def, { source: (def && def.label) || '대용량 첨부' });
            return;
          }
          blobs.push(file.getBlob());
        });
      } else if (def.type === 'vendor_zip') {
        const vendorName = String(this.targetData['수행사'] || '').trim();
        const vendor = getVendorConfig_(vendorName);
        if (!vendor || !vendor.contractorInfoZipFileId) {
          throw new Error('수행사정보 파일 ID가 없습니다. 수행사=' + vendorName + ' / CONFIG.VENDORS의 contractorInfoZipFileId를 확인하세요.');
        }
        const file = DriveApp.getFileById(vendor.contractorInfoZipFileId);
        if (linkSampleAndContractor && def.key === 'contractorInfo') {
          this.addDriveDownloadLink_(file, def, {
            source: '수행사정보',
            displayName: def.filename ? this.renderFilename_(def.filename) : file.getName()
          });
          continue;
        }

        const blob = file.getBlob();
        if (def.filename) blob.setName(this.renderFilename_(def.filename));
        blobs.push(blob);
      } else if (def.type === 'vendor_docx_template') {
        const vendorName = String(this.targetData['수행사'] || '').trim();
        const vendor = getVendorConfig_(vendorName);
        const fieldName = String(def.vendorFileIdField || '').trim();
        const fileId = vendor && fieldName ? String(vendor[fieldName] || '').trim() : '';
        if (!vendor || !fileId) {
          throw new Error((def.label || '수행사별 DOCX 템플릿') + ' 파일 ID가 없습니다. 수행사=' + vendorName + ' / CONFIG.VENDORS의 ' + fieldName + '를 확인하세요.');
        }

        const filename = def.filename
          ? this.renderFilename_(def.filename)
          : appendDateSuffixBeforeExtension_(sanitizeFileName_(vendorName || '수행사') + '_용역표준계약서.docx', getTodayFileDateSuffix_());
        blobs.push(new ServiceStandardContractDocxBuilder(this.targetData, this.ss).build(fileId, filename));
      } else if (def.type === 'vendor_file') {
        const vendorName = String(this.targetData['수행사'] || '').trim();
        const vendor = getVendorConfig_(vendorName);
        const fieldName = String(def.vendorFileIdField || '').trim();
        const fileId = vendor && fieldName ? String(vendor[fieldName] || '').trim() : '';
        if (!vendor || !fileId) {
          throw new Error((def.label || '수행사별 파일') + ' 파일 ID가 없습니다. 수행사=' + vendorName + ' / CONFIG.VENDORS의 ' + fieldName + '를 확인하세요.');
        }

        const file = DriveApp.getFileById(fileId);
        const blob = file.getBlob();
        if (def.filename) blob.setName(this.renderFilename_(def.filename));
        blobs.push(blob);
      } else {
        throw new Error('지원하지 않는 파일 생성 타입입니다: ' + def.type);
      }
    }
    return blobs;
  }

  buildReviewFilesByKey_(reviewPackage) {
    const map = {};
    const files = reviewPackage && Array.isArray(reviewPackage.files) ? reviewPackage.files : [];

    files.forEach(function(file) {
      if (!file || file.useForSend === false) return;
      const key = String(file.key || '').trim();
      if (!key) return;
      if (!map[key]) map[key] = [];
      map[key].push(file);
    });

    return map;
  }

  popReviewFile_(key) {
    const k = String(key || '').trim();
    if (!k || !this.reviewFilesByKey || !this.reviewFilesByKey[k] || !this.reviewFilesByKey[k].length) {
      return null;
    }

    const files = this.reviewFilesByKey[k].slice();
    delete this.reviewFilesByKey[k];
    return files;
  }

  buildBlobFromReviewFile_(def, reviewFile) {
    let items = Array.isArray(reviewFile) ? reviewFile : [reviewFile];

    if (def && def.type === 'multi_sheet_pdf' && Array.isArray(def.sheets)) {
      const allowedSheetSet = buildLowerStringSet_(def.sheets.map(function(item) {
        return item && item.sheetName;
      }).filter(Boolean));

      items = items.filter(function(item) {
        const sheetName = String(item && item.sheetName || '').trim().toLowerCase();
        return !sheetName || allowedSheetSet[sheetName];
      });
    }

    const blobs = [];

    items.forEach(item => {
      if (!item || !item.fileId) return;

      const exportAs = String(item.exportAs || '').trim();
      const filename = String(item.sendFilename || item.name || (def && def.filename ? this.renderFilename_(def.filename) : 'attachment')).trim();

      if (exportAs === 'pdf_from_sheet') {
        const editedSs = SpreadsheetApp.openById(item.fileId);
        let sheet = item.sheetName ? editedSs.getSheetByName(item.sheetName) : null;
        if (!sheet) sheet = editedSs.getSheets()[0];
        if (!sheet) throw new Error('수정용 견적서 시트를 찾지 못했습니다: ' + (item.name || item.fileId));

        const rangeA1 = item.exportRange || (def && def.exportRange) || sheet.getDataRange().getA1Notation();
        blobs.push(this.exporter.exportSheetPdf(editedSs, sheet, rangeA1, filename));
        return;
      }

      if (exportAs === 'xlsx_from_sheet_single') {
        const editedSs = SpreadsheetApp.openById(item.fileId);
        let sheet = item.sheetName ? editedSs.getSheetByName(item.sheetName) : null;
        if (!sheet && item.originalSheetName) sheet = editedSs.getSheetByName(item.originalSheetName);
        if (!sheet) sheet = editedSs.getSheets()[0];
        if (!sheet) throw new Error('수정용 XLSX 시트를 찾지 못했습니다: ' + (item.name || item.fileId));
        blobs.push(this.exporter.exportSingleSheetXlsxValues(sheet, filename));
        return;
      }

      if (exportAs === 'xlsx_from_sheet') {
        blobs.push(this.exporter.exportSpreadsheetXlsxById(item.fileId, filename));
        return;
      }

      if (exportAs === 'docx_from_doc') {
        blobs.push(this.exporter.exportGoogleDocDocxById(item.fileId, filename));
        return;
      }

      if (exportAs === 'raw_file' || !exportAs) {
        const file = DriveApp.getFileById(item.fileId);
        const blob = file.getBlob();
        blob.setName(filename || file.getName());
        blobs.push(blob);
        return;
      }

      throw new Error('지원하지 않는 수정용 파일 exportAs 값입니다: ' + exportAs + ' / ' + (item.name || item.fileId));
    });

    return blobs;
  }

  shouldSendSampleAndContractorAsDriveLinks_(defs) {
    if (!CONFIG.HIWORKS || CONFIG.HIWORKS.SAMPLE_AND_CONTRACTOR_LINK_WHEN_BOTH_SELECTED !== true) return false;
    return this.hasDefinitionKey_(defs, 'sampleReport') && this.hasDefinitionKey_(defs, 'contractorInfo');
  }

  hasDefinitionKey_(defs, key) {
    return (defs || []).some(def => def && def.key === key);
  }

  shouldLinkInsteadOfAttach_(file, def) {
    const cfg = CONFIG.HIWORKS || {};
    if (cfg.OVERSIZED_ATTACHMENT_FALLBACK_TO_DRIVE_LINK !== true) return false;

    const limit = Number(def && def.linkInsteadOfAttachOverBytes || cfg.API_ATTACHMENT_LIMIT_BYTES || 0);
    if (!limit) return false;

    let size = 0;
    try {
      size = Number(file.getSize && file.getSize()) || 0;
    } catch (err) {
      Logger.log('Drive 파일 크기 확인 실패: ' + (file && file.getName ? file.getName() : '') + ' / ' + err);
      return false;
    }

    return size > limit;
  }

  addDriveDownloadLink_(file, def, options) {
    const opts = options || {};
    const name = String(opts.displayName || (file.getName ? file.getName() : '') || '자료 파일');
    const size = file.getSize ? Number(file.getSize()) || 0 : 0;
    const url = this.getViewableDriveFileUrl_(file);
    const downloadUrl = this.getDriveFileDownloadUrl_(file);

    if (!this.targetData.__MAIL_LARGE_ATTACHMENT_LINKS__) {
      this.targetData.__MAIL_LARGE_ATTACHMENT_LINKS__ = [];
    }

    this.targetData.__MAIL_LARGE_ATTACHMENT_LINKS__.push({
      name: name,
      fileId: file.getId ? file.getId() : '',
      mimeType: file.getMimeType ? file.getMimeType() : '',
      sizeBytes: size,
      sizeText: this.formatBytes_(size),
      url: url,
      downloadUrl: downloadUrl,
      source: opts.source || (def && def.label) || '자료'
    });

    Logger.log('첨부 대신 Drive 링크로 전달: ' + name + ' / ' + this.formatBytes_(size) + ' / ' + url);
  }

  getViewableDriveFileUrl_(file) {
    const cfg = CONFIG.HIWORKS || {};
    if (cfg.MAKE_DRIVE_LINK_FILES_VIEWABLE_BY_LINK === true || cfg.MAKE_OVERSIZED_FILE_VIEWABLE_BY_LINK === true) {
      try {
        file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
      } catch (err) {
        Logger.log('Drive 링크공유 설정 실패. 기존 권한으로 링크만 삽입합니다: ' + (file.getName ? file.getName() : '') + ' / ' + err);
      }
    }

    try {
      return file.getUrl();
    } catch (err) {
      return 'https://drive.google.com/file/d/' + file.getId() + '/view?usp=drive_link';
    }
  }

  getDriveFileDownloadUrl_(file) {
    try {
      return 'https://drive.google.com/uc?export=download&id=' + encodeURIComponent(file.getId());
    } catch (err) {
      return this.getViewableDriveFileUrl_(file);
    }
  }

  formatBytes_(bytes) {
    const n = Number(bytes) || 0;
    if (n >= 1024 * 1024) return (n / 1024 / 1024).toFixed(2) + ' MB';
    if (n >= 1024) return (n / 1024).toFixed(2) + ' KB';
    return n + ' B';
  }

  renderFilename_(template) {
    const company = sanitizeFileName_(this.targetData['회사명'] || '고객사');
    const vendor = sanitizeFileName_(this.targetData['수행사'] || '수행사');
    const rendered = String(template || 'attachment')
      .replace(/\{company\}/g, company)
      .replace(/\{vendor\}/g, vendor)
      .replace(/\{date\}/g, getTodayFileDateSuffix_());
    return appendDateSuffixBeforeExtension_(rendered, getTodayFileDateSuffix_());
  }
}

class ExportService {
  constructor(progress) {
    this.progress = progress || null;
  }

  exportSheetPdf(ss, sheet, rangeA1, filename) {
    const params = Object.assign({}, CONFIG.EXPORT.PDF_OPTIONS, {
      format: 'pdf',
      gid: sheet.getSheetId(),
      range: rangeA1
    });
    const url = this.buildSpreadsheetExportUrl_(ss.getId(), params);
    return this.fetchExport_(url, filename, MimeType.PDF);
  }

  exportSingleSheetXlsxValues(sourceSheet, filename) {
    SpreadsheetApp.flush();
    if (CONFIG.EXPORT.WAIT_MS_AFTER_IMAGE_INSERT) {
      Utilities.sleep(Math.min(800, CONFIG.EXPORT.WAIT_MS_AFTER_IMAGE_INSERT));
    }

    const lastRow = Math.max(sourceSheet.getLastRow(), 1);
    const lastCol = Math.max(sourceSheet.getLastColumn(), 1);

    // 중요: sourceSheet.copyTo(temp) 후에는 다른 시트 참조 수식이 #REF!로 깨질 수 있습니다.
    // 따라서 복사 전에 원본 시트의 표시값을 먼저 읽고, 복사본에는 그 표시값을 덮어씁니다.
    const sourceDisplayValues = sourceSheet
      .getRange(1, 1, lastRow, lastCol)
      .getDisplayValues();

    const tempFile = createBlankSpreadsheetFileInMailAutoWorkspace_('값전용_' + sanitizeFileName_(sourceSheet.getName()) + '_' + Utilities.getUuid(), null);
    const tempId = tempFile.getId();
    const temp = SpreadsheetApp.openById(tempId);
    let copied = null;
    try {
      copied = sourceSheet.copyTo(temp).setName(sourceSheet.getName());
      temp.setActiveSheet(copied);

      temp.getSheets().forEach(sh => {
        if (sh.getSheetId() !== copied.getSheetId()) temp.deleteSheet(sh);
      });

      copied
        .getRange(1, 1, lastRow, lastCol)
        .setValues(sourceDisplayValues);

      SpreadsheetApp.flush();

      const url = `https://docs.google.com/spreadsheets/d/${tempId}/export?format=xlsx`;
      return this.fetchExport_(url, filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
    } finally {
      try {
        DriveApp.getFileById(tempId).setTrashed(true);
      } catch (err) {
        Logger.log('임시 XLSX 파일 삭제 실패: ' + err);
      }
    }
  }

  exportSpreadsheetXlsxById(spreadsheetId, filename) {
    SpreadsheetApp.flush();
    const url = 'https://docs.google.com/spreadsheets/d/' + encodeURIComponent(spreadsheetId) + '/export?format=xlsx';
    return this.fetchExport_(url, filename, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
  }

  exportGoogleDocDocxById(docId, filename) {
    DocumentApp.openById(docId).saveAndClose();
    Utilities.sleep(300);
    const url = 'https://docs.google.com/document/d/' + encodeURIComponent(docId) + '/export?format=docx';
    return this.fetchExport_(url, filename, 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
  }

  buildSpreadsheetExportUrl_(spreadsheetId, params) {
    const query = Object.keys(params)
      .filter(k => params[k] !== '' && params[k] != null)
      .map(k => `${encodeURIComponent(k)}=${encodeURIComponent(params[k])}`)
      .join('&');
    return `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${query}`;
  }

  fetchExport_(url, filename, contentType) {
    const res = UrlFetchApp.fetch(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    });
    const code = res.getResponseCode();
    if (code < 200 || code >= 300) {
      throw new Error(`Google export 실패 HTTP ${code}: ${res.getContentText().slice(0, 500)}`);
    }
    return res.getBlob().setName(filename).setContentType(contentType);
  }
}

class DocxTemplateBuilder {
  constructor(targetData, generatorSs) {
    this.targetData = targetData || {};
    this.generatorSs = generatorSs || null;
  }

  build(templateFileId, filename) {
    const sourceId = templateFileId || CONFIG.APPOINTMENT_DOC_TEMPLATE_ID;
    const tempName = 'DOCX작업_' + sanitizeFileName_(this.value_(['회사명', '고객사명', '고객명'], '고객사')) + '_' + Utilities.getUuid();

    // 선임신고서 원본이 Google Docs가 아니라 .docx Office 파일이면
    // Drive API copy + mimeType=Google Docs로 변환 복사한 뒤 DocumentApp으로 치환합니다.
    const tempDocId = this.copyTemplateAsGoogleDoc_(sourceId, tempName);
    const tempFile = DriveApp.getFileById(tempDocId);

    try {
      const doc = this.openGoogleDocWithRetry_(tempDocId);
      const body = doc.getBody();
      const docxValues = this.buildDocxValues_();
      const templateValues = this.buildAppointmentTemplateValues_(docxValues);

      // 1) 새 표준양식 fallback 표현 치환
      //    {{A}} or {{B}}는 A 값이 있으면 A, A가 비어 있으면 B 값으로 치환합니다.
      //    A/B는 기본적으로 마스터시트 헤더명이며, 수행사상호/수행사사업자등록번호/수행사주소/수행사전화번호만 수행사 정보에서 가져옵니다.
      // v95: 선임신고서 표준양식에는 {{계약 당사자(사업자등록증상 법인명)}}{{회사명}}처럼
      //      or 없이 두 placeholder가 붙어 있는 fallback 표현이 있으므로 먼저 한 덩어리로 치환합니다.
      // v96 예외: 위임장 '위임명 {{회사명}} 정보통신유지관리자 선임 신고'의 {{회사명}}은 내부 회사명 그대로 사용합니다.
      this.replaceAppointmentDelegationTitleCompanyV96_(body, templateValues);
      this.replaceAdjacentDoubleBraceFallbackExpressions_(body, templateValues);
      this.replaceDoubleBraceFallbackExpressions_(body, templateValues);

      // 2) {{마스터시트 헤더명}} 및 {{수행사...}} 단일 placeholder 치환
      this.replaceDoubleBracePlaceholders_(body, templateValues);

      // 3) 구버전 템플릿의 시트 참조 문구도 직접 치환합니다.
      //    이 부분 때문에 기존 테스트 출력물에 ‘수기견적계산’!G6 같은 문구가 남았습니다.
      const legacyMap = this.buildLegacyPlaceholderMap_(docxValues);
      Object.keys(legacyMap).forEach(oldText => {
        body.replaceText(escapeRegex_(oldText), this.escapeReplacementText_(legacyMap[oldText] == null ? '' : String(legacyMap[oldText])));
      });

      // 4) 혹시 남은 {{...}} placeholder는 빈칸으로 정리합니다.
      //    발송 문서에 빨간 템플릿 문구가 그대로 남는 것을 방지합니다.
      this.cleanupRemainingDoubleBracePlaceholders_(body);

      // 5) 출력용 DOCX에서는 placeholder의 빨간색/서식이 남지 않게 전체 글자색을 검정으로 통일합니다.
      this.setAllTextColorBlack_(body);

      doc.saveAndClose();

      const url = `https://docs.google.com/document/d/${tempDocId}/export?format=docx`;
      const res = this.fetchGoogleApiWithRetry_(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      }, '선임신고서 DOCX export');
      return res.getBlob()
        .setName(filename)
        .setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } finally {
      if (CONFIG.TEMP.TRASH_TEMP_DOC_AFTER_SUCCESS) {
        try {
          tempFile.setTrashed(true);
        } catch (err) {
          Logger.log('임시 Google Docs 삭제 실패: ' + err);
        }
      }
    }
  }

  buildEditableGoogleDoc(templateFileId, docName, folder) {
    const sourceId = templateFileId || CONFIG.APPOINTMENT_DOC_TEMPLATE_ID;
    const name = sanitizeFileName_(docName || ('선임신고서_및_위임장_' + Utilities.getUuid()));
    const tempDocId = this.copyTemplateAsGoogleDoc_(sourceId, name, folder);
    const tempFile = DriveApp.getFileById(tempDocId);

    try {
      const doc = this.openGoogleDocWithRetry_(tempDocId);
      const body = doc.getBody();
      const docxValues = this.buildDocxValues_();
      const templateValues = this.buildAppointmentTemplateValues_(docxValues);

      // v95: or 없이 붙어 있는 {{A}}{{B}} fallback 표현 선처리
      // v96 예외: 위임장 '위임명 {{회사명}} 정보통신유지관리자 선임 신고'의 {{회사명}}은 내부 회사명 그대로 사용합니다.
      this.replaceAppointmentDelegationTitleCompanyV96_(body, templateValues);
      this.replaceAdjacentDoubleBraceFallbackExpressions_(body, templateValues);
      this.replaceDoubleBraceFallbackExpressions_(body, templateValues);
      this.replaceDoubleBracePlaceholders_(body, templateValues);

      const legacyMap = this.buildLegacyPlaceholderMap_(docxValues);
      Object.keys(legacyMap).forEach(oldText => {
        body.replaceText(escapeRegex_(oldText), this.escapeReplacementText_(legacyMap[oldText] == null ? '' : String(legacyMap[oldText])));
      });

      this.cleanupRemainingDoubleBracePlaceholders_(body);
      this.setAllTextColorBlack_(body);
      doc.saveAndClose();

      return tempFile;
    } catch (err) {
      try { tempFile.setTrashed(true); } catch (e) {}
      throw err;
    }
  }

  buildAppointmentTemplateValues_(docxValues) {
    const values = {};

    // 새 선임신고서 표준양식의 {{...}}는 기본적으로 마스터시트 헤더명 그대로 사용합니다.
    // targetData는 마스터 원행 + 생성대상 계산값이 병합된 객체입니다.
    Object.keys(this.targetData || {}).forEach(key => {
      values[key] = this.stringifyTemplateValue_(this.targetData[key]);
    });

    // v96: 선임신고서 대부분의 고객상호는 사업자등록증상 법인명 우선이지만,
    // 위임장 '위임명 {{회사명}} 정보통신유지관리자 선임 신고'의 {{회사명}}만 내부용 회사명을 써야 하므로 원래 회사명을 별도로 보존합니다.
    values.__INTERNAL_COMPANY_NAME_V96 = this.stringifyTemplateValue_(this.firstNonBlankTemplateValueV95_(this.targetData || {}, [
      '회사명', '고객사명', '고객명', '건물명', '건물명정규화', '고객명/건물명'
    ]));

    // 아래 4개는 사용자가 지정한 예외입니다.
    // 마스터 헤더가 아니라 기존 수행사 정보 시트/수행사 설정에서 가져옵니다.
    ['수행사상호', '수행사사업자등록번호', '수행사주소', '수행사전화번호'].forEach(key => {
      values[key] = this.stringifyTemplateValue_(docxValues[key]);
    });

    // 구버전 템플릿/별칭 호환값은 마스터 원본 헤더가 없거나 공란일 때 보조로 넣습니다.
    // v95: 생성대상/마스터 병합 과정에서 같은 헤더가 공란으로 들어오면
    //      실제 사업자등록증상 법인명 값이 있어도 {{계약 당사자...}}{{회사명}} 표현이 회사명으로만 나가는 문제가 있었습니다.
    Object.keys(docxValues || {}).forEach(key => {
      const cur = Object.prototype.hasOwnProperty.call(values, key) ? String(values[key] || '').trim() : '';
      if (cur) return;
      values[key] = this.stringifyTemplateValue_(docxValues[key]);
    });

    // v95: 선임신고서 표준양식 핵심 placeholder는 정확한 헤더명으로도 보강합니다.
    //      템플릿에 {{계약 당사자(사업자등록증상 법인명)}}{{회사명}}처럼 붙어 있는 경우
    //      첫 번째 값이 반드시 사업자등록증상 법인명 우선으로 잡혀야 합니다.
    this.applyAppointmentCriticalTemplateAliasesV95_(values, docxValues);

    return values;
  }

  applyAppointmentCriticalTemplateAliasesV95_(values, docxValues) {
    const setIfBlank = (key, value) => {
      const cur = Object.prototype.hasOwnProperty.call(values || {}, key) ? String(values[key] || '').trim() : '';
      const next = this.stringifyTemplateValue_(value);
      if (!cur && next) values[key] = next;
    };

    const applicantName = this.firstNonBlankTemplateValueV95_(docxValues, [
      '계약 당사자(사업자등록증상 법인명)',
      '계약당사자(사업자등록증상법인명)',
      '사업자등록증상 법인명',
      '사업자등록증상법인명',
      '신고인상호',
      '상호(명칭)',
      '상호명칭',
      '법인명',
      '회사명'
    ]);
    // 선임신고서/위임장에서는 고객 상호 표기는 사업자등록증상 법인명 우선으로 통일합니다.
    // targetData의 내부 회사명이 이미 values.회사명에 들어와 있어도, 이 문서 안에서는 applicantName이 있으면 덮습니다.
    if (applicantName) {
      ['계약 당사자(사업자등록증상 법인명)', '계약당사자(사업자등록증상법인명)', '사업자등록증상 법인명', '사업자등록증상법인명', '신고인상호', '상호(명칭)', '상호명칭', '법인명', '회사명', '고객사명', '고객명', '건물명', '업체명', '신고인', '위임자'].forEach(key => {
        values[key] = this.stringifyTemplateValue_(applicantName);
      });
    } else {
      ['계약 당사자(사업자등록증상 법인명)', '계약당사자(사업자등록증상법인명)', '사업자등록증상 법인명', '사업자등록증상법인명', '신고인상호', '상호(명칭)', '상호명칭', '법인명'].forEach(key => setIfBlank(key, applicantName));
    }

    const applicantAddress = this.firstNonBlankTemplateValueV95_(docxValues, [
      '사업자등록증상 법인 주소',
      '사업자등록증상법인주소',
      '사업자등록증상 주소',
      '신고인주소',
      '주소'
    ]);
    ['사업자등록증상 법인 주소', '사업자등록증상법인주소', '사업자등록증상 주소', '신고인주소'].forEach(key => setIfBlank(key, applicantAddress));

    const applicantPhone = this.firstNonBlankTemplateValueV95_(docxValues, [
      '대표전화번호',
      '대표 전화번호',
      '신고인전화번호',
      '전화번호',
      '신전화번호'
    ]);
    ['대표전화번호', '대표 전화번호', '신고인전화번호', '전화번호'].forEach(key => setIfBlank(key, applicantPhone));

    const targetAddress = this.firstNonBlankTemplateValueV95_(docxValues, [
      '고객사 상세 주소',
      '고객사 상세주소',
      '대상주소',
      '대상건축물주소',
      '건축물주소',
      '고객사주소'
    ]);
    ['고객사 상세 주소', '고객사 상세주소', '대상주소', '대상건축물주소', '건축물주소'].forEach(key => setIfBlank(key, targetAddress));
  }

  firstNonBlankTemplateValueV95_(obj, keys) {
    for (let i = 0; i < (keys || []).length; i++) {
      const key = keys[i];
      const actualKey = this.findTemplateValueKey_(key, obj || {});
      if (!actualKey) continue;
      const value = obj[actualKey];
      if (value !== null && value !== undefined && String(value).trim() !== '') return value;
    }
    return '';
  }

  replaceAppointmentDelegationTitleCompanyV96_(body, values) {
    if (!body) return;
    const internalCompany = this.getTemplateValue_('__INTERNAL_COMPANY_NAME_V96', values) || this.getTemplateValue_('회사명', values);
    if (!internalCompany) return;

    // 위임장 '위임명' 칸만 예외 처리합니다.
    // 템플릿의 다른 {{회사명}}은 기존 v95 정책대로 사업자등록증상 법인명 우선으로 치환됩니다.
    const replacement = this.escapeReplacementText_(String(internalCompany).trim() + ' 정보통신유지관리자 선임 신고');
    const companyPlaceholder = '\{\{\s*' + this.templateKeyRegex_('회사명') + '\s*\}\}';
    const trailingText = '\s*정보통신유지관리자\s*선임\s*신고';

    // 같은 셀: 위임명 {{회사명}} 정보통신유지관리자 선임 신고
    body.replaceText('위임명\s*' + companyPlaceholder + trailingText, this.escapeReplacementText_('위임명 ' + String(internalCompany).trim() + ' 정보통신유지관리자 선임 신고'));

    // 다른 셀: {{회사명}} 정보통신유지관리자 선임 신고
    // 이 패턴은 위임명 구간 외에는 사실상 쓰이지 않으므로, 안전하게 내부 회사명으로 먼저 치환합니다.
    body.replaceText(companyPlaceholder + trailingText, replacement);
  }

  replaceAdjacentDoubleBraceFallbackExpressions_(body, values) {
    const pairs = [
      ['계약 당사자(사업자등록증상 법인명)', '회사명'],
      ['계약당사자(사업자등록증상법인명)', '회사명'],
      ['사업자등록증상 법인명', '회사명'],
      ['사업자등록증상 법인 주소', '고객사 상세 주소'],
      ['사업자등록증상법인주소', '고객사 상세주소'],
      ['대표전화번호', '직통번호']
    ];

    pairs.forEach(pair => {
      const primaryKey = this.normalizeTemplateKey_(pair[0]);
      const fallbackKey = this.normalizeTemplateKey_(pair[1]);
      const primaryValue = this.getTemplateValue_(primaryKey, values);
      const fallbackValue = this.getTemplateValue_(fallbackKey, values);
      const chosen = primaryValue !== '' ? primaryValue : fallbackValue;
      this.replaceAdjacentFallbackExpression_(body, primaryKey, fallbackKey, chosen);
    });
  }

  replaceAdjacentFallbackExpression_(body, primaryKey, fallbackKey, value) {
    const pattern =
      '\\{\\{\\s*' + this.templateKeyRegex_(primaryKey) + '\\s*\\}\\}' +
      '\\s*' +
      '\\{\\{\\s*' + this.templateKeyRegex_(fallbackKey) + '\\s*\\}\\}';
    body.replaceText(pattern, this.escapeReplacementText_(value));
  }

  replaceDoubleBraceFallbackExpressions_(body, values) {
    const text = body.getText ? String(body.getText() || '') : '';
    const regex = /\{\{\s*([^{}]+?)\s*\}\}\s*(?:or|OR|또는)\s*\{\{\s*([^{}]+?)\s*\}\}/g;
    const seen = {};
    let match;

    while ((match = regex.exec(text)) !== null) {
      const primaryKey = this.normalizeTemplateKey_(match[1]);
      const fallbackKey = this.normalizeTemplateKey_(match[2]);
      if (!primaryKey && !fallbackKey) continue;

      const seenKey = normalizeHeader_(primaryKey) + '||' + normalizeHeader_(fallbackKey);
      if (seen[seenKey]) continue;
      seen[seenKey] = true;

      const primaryValue = this.getTemplateValue_(primaryKey, values);
      const fallbackValue = this.getTemplateValue_(fallbackKey, values);
      const chosen = primaryValue !== '' ? primaryValue : fallbackValue;

      this.replaceFallbackExpression_(body, primaryKey, fallbackKey, chosen);
    }
  }

  replaceFallbackExpression_(body, primaryKey, fallbackKey, value) {
    const pattern =
      '\\{\\{\\s*' + this.templateKeyRegex_(primaryKey) + '\\s*\\}\\}' +
      '\\s*(?:or|OR|또는)\\s*' +
      '\\{\\{\\s*' + this.templateKeyRegex_(fallbackKey) + '\\s*\\}\\}';

    body.replaceText(pattern, this.escapeReplacementText_(value));
  }

  replaceDoubleBracePlaceholders_(body, values) {
    const keys = Object.keys(values || {})
      .filter(key => String(key || '').trim() !== '')
      .sort((a, b) => String(b).length - String(a).length);

    keys.forEach(key => {
      const value = this.getTemplateValue_(key, values);
      const pattern = '\\{\\{\\s*' + this.templateKeyRegex_(key) + '\\s*\\}\\}';
      body.replaceText(pattern, this.escapeReplacementText_(value));
    });
  }

  cleanupRemainingDoubleBracePlaceholders_(body) {
    body.replaceText('\\{\\{[^{}]*\\}\\}', '');
  }

  getTemplateValue_(key, values) {
    const actualKey = this.findTemplateValueKey_(key, values);
    if (!actualKey) return '';

    const value = values[actualKey];
    if (value === null || value === undefined) return '';
    return String(value).trim();
  }

  findTemplateValueKey_(key, values) {
    const rawKey = String(key == null ? '' : key).trim();
    if (!rawKey) return '';

    if (Object.prototype.hasOwnProperty.call(values || {}, rawKey)) return rawKey;

    const normalizedKey = normalizeHeader_(this.normalizeTemplateKey_(rawKey));
    const actualKeys = Object.keys(values || {});
    for (const actualKey of actualKeys) {
      if (normalizeHeader_(this.normalizeTemplateKey_(actualKey)) === normalizedKey) return actualKey;
    }

    return '';
  }

  normalizeTemplateKey_(key) {
    return String(key == null ? '' : key)
      .replace(/[\r\n]+/g, ' ')
      .replace(/\s+/g, ' ')
      .trim();
  }

  templateKeyRegex_(key) {
    // Word/Google Docs 변환 과정에서 placeholder 내부에 줄바꿈이나 공백이 섞여도 매칭되게 처리합니다.
    const compact = normalizeHeader_(this.normalizeTemplateKey_(key));
    if (!compact) return '';
    return compact.split('').map(ch => escapeRegex_(ch)).join('\\s*');
  }

  stringifyTemplateValue_(value) {
    if (value === null || value === undefined) return '';

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, Session.getScriptTimeZone(), 'yyyy. M. d.');
    }

    return String(value).trim();
  }

  escapeReplacementText_(value) {
    // DocumentApp.replaceText의 replacement에서 역슬래시/달러 문자가 특수 처리되는 것을 방지합니다.
    return String(value == null ? '' : value).replace(/\\/g, '\\\\').replace(/\$/g, '\\$');
  }

  buildDocxValues_() {
    const vendorInfo = this.resolveVendorInfo_();

    // v19 선임신고서 신고인 정보 우선순위
    // 1) 상호(명칭): AX 계약 당사자(사업자등록증상 법인명) 우선, 없으면 회사명/고객명/건물명
    // 2) 대표자 성명: 대표자 값이 있을 때만 입력, 없으면 공란
    // 3) 사업자등록번호: 있을 때만 입력, 없으면 공란
    // 4) 주소: AY 사업자등록증상 법인 주소 우선, 없으면 I열 고객사 상세주소/주소 계열
    // 5) 전화번호: 대표전화번호만 사용
    // 6) 날짜: 실제 작성일을 모를 수 있으므로 원문처럼 공란 날짜 유지
    const fallbackCompany = this.value_([
      '회사명', '고객사명', '고객명', '건물명', '건물명정규화', '고객명/건물명'
    ]);

    const applicantName = this.value_([
      // 마스터시트 AX열: 계약 당사자(사업자등록증상 법인명)
      '계약 당사자(사업자등록증상 법인명)',
      '계약당사자(사업자등록증상법인명)',
      '계약 당사자',
      '계약당사자',
      '사업자등록증상 법인명',
      '사업자등록증상법인명',
      '사업자등록증상 상호',
      '사업자등록증상 회사명',
      '사업자등록증상 인명(단체명)',
      '사업자등록증상 인명',
      '등록증상 법인명',
      '등록증상 상호',
      '등록증상 회사명',
      '상호(명칭)',
      '상호',
      '법인명',
      '인명(단체명)'
    ], fallbackCompany);

    const representativeName = this.value_([
      '사업자등록증상 대표자',
      '사업자등록증상 대표자명',
      '사업자등록증상 대표자 성명',
      '등록증상 대표자',
      '등록증상 대표자명',
      '대표자', '대표자명', '대표자 성명', '대표자성명'
    ]);

    const businessNo = this.value_([
      '사업자등록증상 사업자등록번호',
      '등록증상 사업자등록번호',
      '사업자등록번호', '사업자 등록번호', '사업자번호', '등록번호'
    ]);

    const businessCertAddress = this.value_([
      // 마스터시트 AY열: 사업자등록증상 법인 주소
      '사업자등록증상 법인 주소',
      '사업자등록증상법인주소',
      '사업자등록증상 주소',
      '사업자등록증상 사업장 주소',
      '사업자등록증 주소',
      '사업자등록증주소',
      '등록증상 주소',
      '법인 주소',
      '사업장 주소',
      '사업장소재지',
      '소재지'
    ]);

    const customerAddress = this.value_([
      // I열: 고객사 상세 주소. 사업자등록증 주소가 비어 있을 때 신고인 주소 fallback으로도 사용합니다.
      '고객사 상세 주소',
      '고객사 상세주소',
      '고객사상세주소',
      '고객사 주소',
      '고객사주소',
      '상세주소',
      '상세 주소',
      '주소',
      '도로명주소',
      '선택도로명주소',
      '지번주소',
      '주소수동입력(도로명)',
      '주소수동입력도로명'
    ]);

    // v22: 신고인 주소와 대상 건축물 주소를 분리합니다.
    // - {{주소}}: 사업자등록증상 법인 주소 우선, 없으면 I열 고객사 상세 주소 fallback
    // - {{대상주소}}: 사업자등록증 주소가 아니라 실제 대상 건축물 주소, 즉 I열 고객사 상세 주소 우선
    const targetBuildingAddress = this.value_([
      '고객사 상세 주소',
      '고객사 상세주소',
      '고객사상세주소',
      '대상주소',
      '대상 주소',
      '대상건축물주소',
      '대상 건축물 주소',
      '건축물주소',
      '건축물 주소',
      '용역대상 주소',
      '용역 대상 주소',
      '관리대상 주소',
      '관리 대상 주소',
      '고객사 주소',
      '고객사주소',
      '도로명주소',
      '선택도로명주소',
      '지번주소',
      '주소수동입력(도로명)',
      '주소수동입력도로명'
    ], customerAddress);

    const applicantAddress = businessCertAddress || customerAddress;
    const applicantPhone = this.value_([
      '대표전화번호', '대표 전화번호', '대표번호', '회사 대표전화번호'
    ]);
    const blankDate = '        년        월        일';

    return {
      // v95: 새 선임신고서 표준양식의 정확한 placeholder명도 명시적으로 제공합니다.
      //      {{계약 당사자(사업자등록증상 법인명)}}{{회사명}}처럼 or 없이 붙은 fallback 표현을 안전하게 처리하기 위함입니다.
      '계약 당사자(사업자등록증상 법인명)': applicantName,
      '계약당사자(사업자등록증상법인명)': applicantName,
      '사업자등록증상 법인명': applicantName,
      '사업자등록증상법인명': applicantName,
      '사업자등록증상 법인 주소': applicantAddress,
      '사업자등록증상법인주소': applicantAddress,
      '사업자등록증상 주소': applicantAddress,
      '고객사 상세 주소': targetBuildingAddress,
      '고객사 상세주소': targetBuildingAddress,
      '대표전화번호': applicantPhone,
      '직통번호': this.value_(['직통번호', '직통 번호', '직통', '휴대폰번호', '핸드폰번호']),

      // 구 템플릿 호환용 placeholder. 실제로는 신고인 상호(명칭) 값입니다.
      // 선임신고서/위임장 안의 모든 고객 상호·회사명·업체명·건물명 표기는
      // 사업자등록증상 법인명(applicantName)을 우선 사용합니다.
      '회사명': applicantName,
      '고객사명': applicantName,
      '고객명': applicantName,
      '건물명': applicantName,
      '건물명정규화': applicantName,
      '업체명': applicantName,
      '법인명': applicantName,
      '상호': applicantName,
      '상호명': applicantName,
      '상호(명칭)': applicantName,
      '신청인': applicantName,
      '신청인상호': applicantName,
      '신고인': applicantName,
      '신고인명': applicantName,
      '위임자': applicantName,
      '위임자상호': applicantName,
      '위임자명': applicantName,

      // 새 템플릿용 명시적 placeholder
      '신고인상호': applicantName,
      '상호명칭': applicantName,
      '신고인대표자성명': representativeName,
      '신고인사업자등록번호': businessNo,
      '신고인주소': applicantAddress,
      '신고인전화번호': applicantPhone,
      '신고인연면적': this.cleanArea_(this.value_(['연면적', '연면적(㎡)', '연면적㎡'])),
      '대상주소': targetBuildingAddress,
      '대상건축물주소': targetBuildingAddress,
      '건축물주소': targetBuildingAddress,
      '문서날짜': blankDate,

      // 구 템플릿 placeholder 호환
      '대표자성명': representativeName,
      '대표자명': representativeName,
      '사업자등록번호': businessNo,
      '주소': applicantAddress,
      '고객사주소': customerAddress,
      '대상주소': targetBuildingAddress,
      '대상 건축물 주소': targetBuildingAddress,
      '대상건축물주소': targetBuildingAddress,
      '건축물주소': targetBuildingAddress,
      '신전화번호': applicantPhone,
      '전화번호': applicantPhone,
      '연면적': this.cleanArea_(this.value_(['연면적', '연면적(㎡)', '연면적㎡'])),
      '오늘날짜': blankDate,

      '수행사': this.value_(['수행사']),
      '수행사상호': vendorInfo.name,
      '수행사명': vendorInfo.name,
      '수행사사업자등록번호': vendorInfo.businessNo,
      '수행사주소': vendorInfo.address,
      '수행사전화번호': vendorInfo.phone
    };
  }

  buildLegacyPlaceholderMap_(v) {
    return {
      '‘수기견적계산’!G6': v['회사명'],
      '‘수기견적계산’!F13': v['대표자성명'],
      '‘수기견적계산’!G13': v['사업자등록번호'],
      // 구 템플릿은 신고인 주소/대상주소가 둘 다 수기견적계산!H6으로 되어 있어
      // 완벽히 분리할 수 없습니다. 새 표준양식에서는 {{주소}} / {{대상주소}}를 사용하세요.
      '‘수기견적계산’!H6': v['주소'],
      '‘수기견적계산’!I13': v['전화번호'],
      '‘수기견적계산’!J6': v['연면적'],
      '‘수행사 정보’시트에서 J3, J19, J29 중 하나': v['수행사상호'],
      '‘수행사 정보’시트에서 AL13, AL26, AL34 중 하나': v['수행사사업자등록번호'],
      '‘수행사 정보’시트에서 AL3, AL19, AL29 중 하나': v['수행사주소'],
      '‘수행사 정보’시트에서 AL14, AL21, AL30 중 하나': v['수행사전화번호'],
      '2026년    00월     00일': v['오늘날짜'],
      '2026년    00월    00일': v['오늘날짜']
    };
  }

  resolveVendorInfo_() {
    const vendorKey = normalizeVendorName_(this.value_(['수행사']));
    const fallbackName = getVendorDisplayName_(vendorKey);
    const result = { name: fallbackName, businessNo: '', address: '', phone: '' };

    if (!this.generatorSs) return result;

    const sheet = this.generatorSs.getSheetByName('수행사 정보');
    if (!sheet) return result;

    const values = sheet.getDataRange().getDisplayValues();
    let start = -1;
    for (let r = 0; r < values.length; r++) {
      if (normalizeVendorName_(values[r][0]) === vendorKey) {
        start = r;
        break;
      }
    }
    if (start < 0) return result;

    for (let r = start; r < values.length; r++) {
      const row = values[r];
      if (r > start && String(row[0] || '').trim()) break;

      const labelB = normalizeHeader_(row[1]);
      const valueJ = String(row[9] || '').trim();
      if (labelB.indexOf(normalizeHeader_('상호')) >= 0 && valueJ) result.name = valueJ;
      if (labelB.indexOf(normalizeHeader_('주소')) >= 0 && valueJ) result.address = valueJ;
      if (labelB.indexOf(normalizeHeader_('등록번호')) >= 0 && valueJ) result.businessNo = valueJ;

      for (let c = 0; c < row.length - 1; c++) {
        const label = normalizeHeader_(row[c]);
        const nextValue = String(row[c + 1] || '').trim();
        if (!nextValue) continue;
        if (!result.phone && (label.indexOf(normalizeHeader_('대표번호')) >= 0 || label.indexOf(normalizeHeader_('전화번호')) >= 0)) {
          result.phone = nextValue;
        }
        if (!result.businessNo && label.indexOf(normalizeHeader_('사업자등록번호')) >= 0) {
          result.businessNo = nextValue;
        }
      }
    }

    return result;
  }

  value_(keys, fallback) {
    const list = Array.isArray(keys) ? keys : [keys];

    // 1차: 정확한 헤더명 매칭
    for (const key of list) {
      const v = this.targetData[key];
      if (v !== null && v !== undefined && String(v).trim() !== '') return String(v).trim();
    }

    // 2차: 공백/괄호/특수문자 차이를 흡수하는 정규화 매칭
    // 예: 사업자등록증상 법인명 / 사업자등록증상법인명 / 사업자등록증상_법인명
    const actualKeys = Object.keys(this.targetData || {});
    for (const key of list) {
      const target = normalizeHeader_(key);
      for (const actualKey of actualKeys) {
        if (normalizeHeader_(actualKey) !== target) continue;
        const v = this.targetData[actualKey];
        if (v !== null && v !== undefined && String(v).trim() !== '') return String(v).trim();
      }
    }

    return fallback || '';
  }

  cleanArea_(value) {
    return String(value || '').replace(/㎡/g, '').trim();
  }

  setAllTextColorBlack_(body) {
    this.walkElement_(body, element => {
      try {
        if (element.editAsText) {
          const text = element.editAsText();
          const len = text.getText().length;
          if (len > 0) text.setForegroundColor(0, len - 1, '#000000');
        }
      } catch (err) {
        // 일부 요소는 editAsText 범위 설정이 안 될 수 있으므로 무시합니다.
      }
    });
  }

  walkElement_(element, fn) {
    fn(element);
    if (!element.getNumChildren) return;
    const n = element.getNumChildren();
    for (let i = 0; i < n; i++) {
      this.walkElement_(element.getChild(i), fn);
    }
  }

  copyTemplateAsGoogleDoc_(fileId, name, folder) {
    // v73:
    // 1) 원본이 네이티브 Google Docs면 바로 지정 폴더에 복사합니다.
    // 2) 원본이 DOCX 등 Office 파일이면 매번 변환하지 않고, 원본 수정시각 기준으로 변환된 Google Docs 템플릿을 캐시합니다.
    //    이후 실제 수정용 파일은 캐시된 Google Docs 템플릿을 복사하므로 선임신고서 생성 시간이 크게 줄어듭니다.
    const sourceFile = DriveApp.getFileById(fileId);
    const sourceMime = String(sourceFile.getMimeType ? sourceFile.getMimeType() : '');

    if (sourceMime === 'application/vnd.google-apps.document' || sourceMime === MimeType.GOOGLE_DOCS) {
      return this.copyGoogleDocFile_(sourceFile, name, folder).getId();
    }

    const cachedTemplateId = this.getOrCreateConvertedGoogleDocTemplateCache_(fileId, sourceFile);
    const cachedFile = DriveApp.getFileById(cachedTemplateId);
    return this.copyGoogleDocFile_(cachedFile, name, folder).getId();
  }

  copyGoogleDocFile_(sourceFile, name, folder) {
    const targetFolder = folder || resolveMailAutoWritableParentFolder_(sourceFile);
    try {
      return targetFolder ? sourceFile.makeCopy(name, targetFolder) : sourceFile.makeCopy(name);
    } catch (copyErr) {
      // 일부 공유드라이브/권한 환경에서는 folder 지정 복사가 실패할 수 있으므로 기본 복사 후 이동을 시도합니다.
      const copied = sourceFile.makeCopy(name);
      if (targetFolder) {
        try {
          copied.moveTo(targetFolder);
        } catch (moveErr) {
          try { targetFolder.addFile(copied); } catch (addErr) {
            Logger.log('Google Docs 복사본 작업공간 이동/추가 실패: ' + moveErr + ' / ' + addErr);
          }
        }
      }
      return copied;
    }
  }

  getOrCreateConvertedGoogleDocTemplateCache_(sourceId, sourceFile) {
    const props = PropertiesService.getScriptProperties();
    const lastUpdated = sourceFile && sourceFile.getLastUpdated ? sourceFile.getLastUpdated().getTime() : 0;
    const sourceName = sourceFile && sourceFile.getName ? sourceFile.getName() : '선임신고서_템플릿';
    const key = 'MAILAUTO_APPOINTMENT_DOC_TEMPLATE_CACHE_' + sourceId;

    const raw = props.getProperty(key);
    if (raw) {
      try {
        const cached = JSON.parse(raw);
        if (cached && cached.templateId && Number(cached.lastUpdated || 0) === Number(lastUpdated || 0)) {
          const cachedFile = DriveApp.getFileById(cached.templateId);
          const cachedMime = String(cachedFile.getMimeType ? cachedFile.getMimeType() : '');
          if (cachedMime === 'application/vnd.google-apps.document' || cachedMime === MimeType.GOOGLE_DOCS) {
            return cached.templateId;
          }
        }
      } catch (err) {
        Logger.log('선임신고서 Google Docs 템플릿 캐시 확인 실패. 새로 생성합니다: ' + err);
      }
    }

    const parent = resolveMailAutoWritableParentFolder_(sourceFile);
    const cacheName = '_메일자동화_선임신고서_GoogleDocs변환캐시_' + sanitizeFileName_(sourceName) + '_' + String(lastUpdated || Date.now());
    const convertedId = this.copyTemplateAsGoogleDocByDriveApi_(sourceId, cacheName, parent);

    props.setProperty(key, JSON.stringify({
      sourceId: sourceId,
      templateId: convertedId,
      lastUpdated: lastUpdated,
      sourceName: sourceName,
      createdAt: new Date().toISOString()
    }));

    return convertedId;
  }

  copyTemplateAsGoogleDocByDriveApi_(fileId, name, folder) {
    // .docx 등 Office 파일은 Google Docs로 변환 복사해야 DocumentApp에서 수정할 수 있습니다.
    // 사용자 rate limit이 걸릴 수 있으므로 지수 백오프 재시도를 적용합니다.
    const url =
      'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(fileId) +
      '/copy?supportsAllDrives=true&fields=id,name,mimeType';

    const payload = {
      name: name,
      mimeType: 'application/vnd.google-apps.document'
    };
    if (folder && folder.getId) {
      payload.parents = [folder.getId()];
    }

    const res = this.fetchGoogleApiWithRetry_(url, {
      method: 'post',
      contentType: 'application/json',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    }, '선임신고서 템플릿 Google Docs 변환 복사');

    const text = res.getContentText();
    const data = JSON.parse(text || '{}');
    if (!data.id) {
      throw new Error('선임신고서 템플릿 변환 복사 응답에 파일 ID가 없습니다: ' + text.slice(0, 800));
    }

    return data.id;
  }

  getDriveFileMeta_(fileId) {
    const url =
      'https://www.googleapis.com/drive/v3/files/' +
      encodeURIComponent(fileId) +
      '?supportsAllDrives=true&fields=id,name,mimeType';

    const res = this.fetchGoogleApiWithRetry_(url, {
      method: 'get',
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true
    }, '선임신고서 템플릿 파일 정보 조회');

    return JSON.parse(res.getContentText() || '{}');
  }

  fetchGoogleApiWithRetry_(url, options, label) {
    const waits = [800, 1600, 3200, 6400, 10000, 15000];
    let lastCode = '';
    let lastText = '';

    for (let i = 0; i < waits.length; i++) {
      const res = UrlFetchApp.fetch(url, options);
      const code = res.getResponseCode();
      const text = res.getContentText();

      if (code >= 200 && code < 300) {
        return res;
      }

      lastCode = code;
      lastText = text;

      if (!this.isRetryableGoogleApiError_(code, text) || i === waits.length - 1) {
        break;
      }

      // 여러 사용자가 동시에 선임신고서를 만들 때 같은 타이밍에 재시도하지 않도록 약간 흔듭니다.
      Utilities.sleep(waits[i] + Math.floor(Math.random() * 500));
    }

    throw new Error(
      label + ' 실패\n' +
      'HTTP ' + lastCode + '\n' +
      String(lastText || '').slice(0, 1200) + '\n\n' +
      '확인할 것: 템플릿 파일 접근 권한, Drive API 권한, 파일 ID, 또는 잠시 후 재시도'
    );
  }

  isRetryableGoogleApiError_(code, text) {
    const body = String(text || '');
    return (
      code === 429 ||
      code === 500 ||
      code === 502 ||
      code === 503 ||
      code === 504 ||
      (code === 403 && (
        body.indexOf('userRateLimitExceeded') >= 0 ||
        body.indexOf('rateLimitExceeded') >= 0 ||
        body.indexOf('User rate limit exceeded') >= 0
      ))
    );
  }

  openGoogleDocWithRetry_(docId) {
    let lastErr = null;

    for (let i = 0; i < 6; i++) {
      try {
        return DocumentApp.openById(docId);
      } catch (err) {
        lastErr = err;
        Utilities.sleep(700 + i * 500);
      }
    }

    throw new Error(
      '변환된 선임신고서 Google Docs를 열지 못했습니다.\n' +
      '임시문서ID: ' + docId + '\n' +
      '원인: ' + String(lastErr && lastErr.stack ? lastErr.stack : lastErr)
    );
  }
}


class ServiceStandardContractDocxBuilder extends DocxTemplateBuilder {
  constructor(targetData, generatorSs) {
    super(targetData, generatorSs);
  }

  build(templateFileId, filename) {
    const sourceId = templateFileId;
    if (!sourceId) throw new Error('용역표준계약서 템플릿 파일 ID가 비어 있습니다.');

    const tempName = '용역표준계약서작업_' + sanitizeFileName_(this.value_(['회사명', '고객사명', '고객명'], '고객사')) + '_' + Utilities.getUuid();
    const tempDocId = this.copyTemplateAsGoogleDoc_(sourceId, tempName);
    const tempFile = DriveApp.getFileById(tempDocId);

    try {
      const doc = this.openGoogleDocWithRetry_(tempDocId);
      const body = doc.getBody();
      const values = this.buildStandardContractValues_();

      // 1) {A} or {B} 형태는 먼저 치환합니다.
      //    A가 비어 있지 않으면 A, 비어 있으면 B를 사용합니다.
      this.replaceKnownFallbackExpressions_(body, values);

      // 2) 금액 한글 표기처럼 뒤에 "원整"이 붙어 있는 특수 케이스를 먼저 치환합니다.
      this.replaceStandardContractSpecialPlaceholders_(body, values);

      // 3) {마스터시트 헤더명} 단일 placeholder 치환.
      //    업로드된 표준계약서 placeholder는 마스터시트 헤더명과 동일하므로 기본적으로 그대로 매칭합니다.
      this.replaceSingleBracePlaceholders_(body, values);

      // 4) 템플릿에 박혀 있는 날짜는 실제 발송일로 교체합니다.
      //    예: 2026. 06. 16. -> 2026. 06. 17.
      this.replaceContractDateLiteral_(body, values['오늘날짜']);

      // 5) 최종 출력 전 전체 글자색 검정 통일.
      this.setAllTextColorBlack_(body);

      doc.saveAndClose();

      const url = 'https://docs.google.com/document/d/' + encodeURIComponent(tempDocId) + '/export?format=docx';
      const res = this.fetchGoogleApiWithRetry_(url, {
        headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
        muteHttpExceptions: true
      }, '용역표준계약서 DOCX export');

      return res.getBlob()
        .setName(filename || '용역표준계약서.docx')
        .setContentType('application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    } finally {
      if (CONFIG.TEMP.TRASH_TEMP_DOC_AFTER_SUCCESS) {
        try {
          tempFile.setTrashed(true);
        } catch (err) {
          Logger.log('임시 용역표준계약서 Google Docs 삭제 실패: ' + err);
        }
      }
    }
  }

  buildStandardContractValues_() {
    const values = {};
    const data = this.targetData || {};

    Object.keys(data).forEach(key => {
      values[key] = this.standardContractValueToText_(key, data[key]);
    });

    // 자주 쓰는 헤더는 정규화 매칭으로도 한 번 더 보강합니다.
    const ensure = (name, aliases, fallback) => {
      if (!this.isBlank_(values[name])) return;
      const found = this.value_(aliases || [name], fallback || '');
      values[name] = this.standardContractValueToText_(name, found);
    };

    ensure('회사명', ['회사명', '고객사명', '고객명', '건물명', '건물명정규화']);
    ensure('계약 당사자(사업자등록증상 법인명)', [
      '계약 당사자(사업자등록증상 법인명)',
      '계약당사자(사업자등록증상법인명)',
      '계약 당사자',
      '계약당사자',
      '사업자등록증상 법인명',
      '사업자등록증상법인명',
      '사업자등록증상 상호',
      '사업자등록증상 회사명',
      '사업자등록증상 인명(단체명)',
      '등록증상 법인명',
      '등록증상 상호',
      '상호(명칭)',
      '상호',
      '법인명'
    ]);
    ensure('사업자등록증상 법인 주소', [
      '사업자등록증상 법인 주소',
      '사업자등록증상법인주소',
      '사업자등록증상 주소',
      '사업자등록증상 사업장 주소',
      '사업자등록증 주소',
      '사업자등록증주소',
      '등록증상 주소',
      '법인 주소',
      '사업장 주소',
      '사업장소재지',
      '소재지'
    ]);
    ensure('고객사 상세 주소', [
      '고객사 상세 주소',
      '고객사 상세주소',
      '고객사상세주소',
      '고객사 주소',
      '고객사주소',
      '상세주소',
      '상세 주소',
      '주소',
      '도로명주소',
      '선택도로명주소',
      '지번주소'
    ]);
    ensure('대표전화번호', ['대표전화번호', '대표 전화번호', '대표번호', '회사 대표전화번호', '전화번호']);
    ensure('직통번호', ['직통번호', '직통 번호', '직통번호or휴대폰번호', '휴대폰번호', '핸드폰번호']);
    ensure('사업자등록번호', ['사업자등록번호', '사업자 등록번호', '사업자번호', '등록번호']);
    ensure('계약단위', ['계약단위', '계약 단위', '계약기간', '계약 기간']);
    ensure('최종 견적가', ['최종 견적가', '최종견적가', '최종가', '최종 금액', '계약금액']);
    ensure('부가세', ['부가세', '부가세 포함\n(별도시 체크X)', 'VAT']);
    ensure('용역신청서특약사항', ['용역신청서특약사항', '용역신청서 특약사항', '특약사항', '기타사항', '기타 사항']);

    const amount = parseMoneyNumber_(values['최종 견적가'] || this.value_(['최종 견적가', '최종견적가', '최종가', '계약금액']));
    const koreanWonJeong = amount > 0 ? numberToKoreanWonJeong_(amount) : '';
    values['최종 견적가'] = amount > 0 ? this.formatWonCurrency_(amount) : String(values['최종 견적가'] || '').trim();
    values['numberToKoreanWonJeong_(최종 견적가)'] = koreanWonJeong;
    values['numberToKoreanWonJeong_(최종견적가)'] = koreanWonJeong;
    values['한글최종견적가'] = koreanWonJeong;

    const today = this.getTodayContractDateText_();
    values['오늘날짜'] = today;
    values['문서날짜'] = today;
    values['계약일'] = today;
    values['발송일'] = today;

    return values;
  }

  replaceKnownFallbackExpressions_(body, values) {
    const pairs = [
      ['계약 당사자(사업자등록증상 법인명)', '회사명'],
      ['사업자등록증상 법인 주소', '고객사 상세 주소'],
      ['대표전화번호', '직통번호']
    ];

    pairs.forEach(pair => {
      this.replaceFallbackExpression_(body, pair[0], pair[1], values);
    });
  }

  replaceFallbackExpression_(body, primaryKey, fallbackKey, values) {
    const primary = this.valueFromMap_(values, primaryKey);
    const fallback = this.valueFromMap_(values, fallbackKey);
    const chosen = !this.isBlank_(primary) ? primary : fallback;

    const pattern = '\\{' + escapeRegex_(primaryKey) + '\\}\\s*or\\s*\\{' + escapeRegex_(fallbackKey) + '\\}';
    body.replaceText(pattern, chosen == null ? '' : String(chosen));
  }

  replaceStandardContractSpecialPlaceholders_(body, values) {
    const koreanWonJeong = this.valueFromMap_(values, 'numberToKoreanWonJeong_(최종 견적가)');
    const koreanWonChineseJeong = String(koreanWonJeong || '').replace(/정$/, '整');

    // 계약서 1면 금액 칸은 템플릿 자체가 "{numberToKoreanWonJeong_(최종 견적가)}원整" 형태입니다.
    // 함수 결과가 "삼백오십만원정"인 상태에서 뒤의 "원整"이 또 붙지 않도록 전체 구문을 먼저 치환합니다.
    if (koreanWonChineseJeong) {
      body.replaceText('\\{numberToKoreanWonJeong_\\(최종 견적가\\)\\}원整', koreanWonChineseJeong);
      body.replaceText('\\{numberToKoreanWonJeong_\\(최종견적가\\)\\}원整', koreanWonChineseJeong);
    }
  }

  replaceSingleBracePlaceholders_(body, values) {
    const keys = Object.keys(values || {}).sort((a, b) => b.length - a.length);
    keys.forEach(key => {
      const value = values[key] == null ? '' : String(values[key]);
      body.replaceText('\\{' + escapeRegex_(key) + '\\}', value);
    });

    // 정규화하면 같은 헤더인데 실제 띄어쓰기/괄호 차이가 있는 경우를 마지막으로 한 번 더 흡수합니다.
    const actualKeys = Object.keys(this.targetData || {});
    actualKeys.forEach(actualKey => {
      const normalized = normalizeHeader_(actualKey);
      if (!normalized) return;
      const value = this.standardContractValueToText_(actualKey, this.targetData[actualKey]);
      body.replaceText('\\{' + escapeRegex_(actualKey) + '\\}', value);
    });
  }

  replaceContractDateLiteral_(body, todayText) {
    const today = String(todayText || this.getTodayContractDateText_()).trim();
    if (!today) return;

    // 표준계약서 상단/하단의 고정 날짜를 발송일로 교체합니다.
    // 2026. 06. 16 / 2026. 06. 16. / 2026.06.16. 모두 대응합니다.
    body.replaceText('20\\d{2}\\.\\s*\\d{1,2}\\.\\s*\\d{1,2}\\.?', today);
  }

  valueFromMap_(values, key) {
    if (!values) return '';
    if (Object.prototype.hasOwnProperty.call(values, key)) return values[key];

    const wanted = normalizeHeader_(key);
    const keys = Object.keys(values);
    for (let i = 0; i < keys.length; i++) {
      if (normalizeHeader_(keys[i]) === wanted) return values[keys[i]];
    }
    return '';
  }

  standardContractValueToText_(key, value) {
    if (value === null || typeof value === 'undefined') return '';

    if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
      return Utilities.formatDate(value, this.getScriptTimezone_(), 'yyyy. MM. dd.');
    }

    const text = String(value).trim();
    if (!text) return '';

    const normalizedKey = normalizeHeader_(key);
    if (normalizedKey === normalizeHeader_('최종 견적가') || normalizedKey === normalizeHeader_('최종견적가') || normalizedKey === normalizeHeader_('최종가') || normalizedKey.indexOf(normalizeHeader_('계약금액')) >= 0) {
      const amount = parseMoneyNumber_(text);
      return amount > 0 ? this.formatWonCurrency_(amount) : text;
    }

    return text;
  }

  getTodayContractDateText_() {
    return Utilities.formatDate(new Date(), this.getScriptTimezone_(), 'yyyy. MM. dd.');
  }

  getScriptTimezone_() {
    try {
      return Session.getScriptTimeZone() || 'Asia/Seoul';
    } catch (err) {
      return 'Asia/Seoul';
    }
  }

  formatWonCurrency_(value) {
  const formatted = this.formatThousands_(value);
  return formatted ? '￦ ' + formatted : '';
  }
  
  formatThousands_(value) {
    const n = Math.floor(Number(value) || 0);
    if (!n) return '';
    return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ',');
  }

  isBlank_(value) {
    return value === null || typeof value === 'undefined' || String(value).trim() === '';
  }
}


function formatTemplateValueForMailAutoV434_(key, value) {
  if (value === null || value === undefined) return '';

  if (Object.prototype.toString.call(value) === '[object Date]' && !isNaN(value.getTime())) {
    return formatYyyyMmDdV434_(value);
  }

  const rule = findDataFormatRuleByHeaderV434_(key);
  if (rule && rule.type === 'date') {
    const d = parseContractDateV90_(value);
    return d ? formatYyyyMmDdV434_(d) : String(value || '').trim();
  }

  return String(value).trim();
}

class MailMessage {
  constructor({ from, to, cc, subject, bodyHtml, attachments }) {
    this.from = from;
    this.to = to || [];
    this.cc = cc || [];
    this.subject = subject;
    this.bodyHtml = bodyHtml;
    this.attachments = attachments || [];
  }
}

class HiworksMailer {
  constructor(progress) {
    this.progress = progress || null;
  }

  send(message) {
    const token = this.getToken_();
    const fields = {
      to: this.normalizeAddressList_(message.to),
      user_id: this.normalizeUserId_(message.from),
      subject: String(message.subject || '').trim(),
      content: String(message.bodyHtml || ''),
      save_sent_mail: CONFIG.HIWORKS.SAVE_SENT_MAIL || 'Y'
    };

    const cc = this.normalizeAddressList_(message.cc);
    if (cc) fields.cc = cc;

    this.validateFields_(fields);
    this.validateAttachments_(message.attachments || []);

    // v42 핵심:
    // 기존 방식은 blob.getBytes()를 직접 payload 배열에 push해서 multipart를 만들었기 때문에,
    // 31MB 샘플보고서 같은 파일에서 Apps Script 메모리 초과가 날 수 있습니다.
    // 기본값은 UrlFetchApp 네이티브 multipart payload 방식입니다.
    if (CONFIG.HIWORKS.USE_NATIVE_MULTIPART_PAYLOAD !== false) {
      return this.sendNativeMultipart_(token, fields, message.attachments || []);
    }

    // 혹시 하이웍스 API가 네이티브 필드명을 거부할 때를 대비해 구방식 fallback 함수는 남겨둡니다.
    // 단, 대용량 파일에서는 이 경로를 쓰면 메모리 초과가 날 수 있으므로 기본값으로 사용하지 않습니다.
    return this.sendManualMultipart_(token, fields, message.attachments || []);
  }

  sendNativeMultipart_(token, fields, attachments) {
    if (this.progress) this.progress.update(86, '하이웍스 네이티브 multipart 구성 중');

    const payload = {};
    Object.keys(fields).forEach(name => {
      const value = fields[name];
      if (value === null || typeof value === 'undefined' || value === '') return;
      payload[name] = String(value);
    });

    const fileNames = [];
    const fileFields = [];
    const multiple = (attachments || []).length > 1;

    (attachments || []).forEach((blob, idx) => {
      if (!blob) return;
      if (this.progress) {
        this.progress.update(88 + Math.min(5, idx), '하이웍스 첨부 준비 중: ' + (blob.getName() || 'attachment'));
      }

      const fileName = this.sanitizeAttachmentFileName_(blob.getName() || ('attachment_' + (idx + 1)));
      const fieldName = this.getNativeFileFieldName_(idx, multiple);

      // setName은 Blob 객체 자신을 반환합니다. 파일명 정리만 하고 바이너리는 직접 getBytes()로 풀지 않습니다.
      payload[fieldName] = blob.setName(fileName);
      fileNames.push(fileName);
      fileFields.push(fieldName);
    });

    if (this.progress) this.progress.update(93, '하이웍스 서버 응답 대기 중');

    const res = UrlFetchApp.fetch(CONFIG.HIWORKS.ENDPOINT, {
      method: 'post',
      headers: { Authorization: 'Bearer ' + token },
      payload: payload,
      muteHttpExceptions: true
      // contentType을 일부러 지정하지 않습니다.
      // payload에 Blob이 있으면 UrlFetchApp이 multipart/form-data와 boundary를 자동 생성합니다.
    });

    return this.handleHiworksResponse_(res, fields, fileNames, {
      mode: 'native_multipart',
      fileFields: fileFields,
      note: 'UrlFetchApp native payload. Blob을 JS byte[]로 직접 풀지 않음.'
    });
  }

  getNativeFileFieldName_(idx, multiple) {
    const base = CONFIG.HIWORKS.FILE_FIELD_NAME || 'files[]';

    if (!multiple) return base;

    const mode = String(CONFIG.HIWORKS.NATIVE_FILE_FIELD_MODE || 'indexed_brackets');
    if (mode === 'same_name_first_only') return base;

    // files[] -> files[0], files[1] ...
    // PHP/일반 multipart 파서에서 배열 파일 필드로 해석될 가능성이 가장 높습니다.
    if (mode === 'indexed_brackets') {
      const stem = base.replace(/\[\]$/, '');
      return stem + '[' + idx + ']';
    }

    // files[]_0, files[]_1 형태. 거의 쓰지 않지만 긴급 테스트용으로 남겨둡니다.
    if (mode === 'suffix') return base + '_' + idx;

    return idx === 0 ? base : base.replace(/\[\]$/, '') + '[' + idx + ']';
  }

  sendManualMultipart_(token, fields, attachments) {
    if (this.progress) this.progress.update(86, '하이웍스 수동 multipart 구성 중');

    const boundary = '----HiworksMail' + Utilities.getUuid().replace(/-/g, '');
    const payload = [];
    const fileNames = [];
    const fileFields = [];

    const appendBytes = bytes => {
      for (let i = 0; i < bytes.length; i++) payload.push(bytes[i]);
    };

    // v55 핵심:
    // 네이버 수신이 정상 확인된 기존 개발자 코드와 동일한 multipart 구성으로 복귀합니다.
    // - 텍스트 파트에는 별도 Content-Type을 넣지 않음
    // - 첨부 filename에는 raw 파일명을 그대로 넣음
    // - filename* / RFC2047 encoded-word는 하이웍스에서 파일명으로 그대로 노출되어 제거
    // - Content-Transfer-Encoding: binary 포함
    const appendText = text => appendBytes(Utilities.newBlob(String(text)).getBytes());

    Object.keys(fields).forEach(name => {
      const value = fields[name];
      if (value === null || typeof value === 'undefined' || value === '') return;
      appendText(
        '--' + boundary + '\r\n' +
        'Content-Disposition: form-data; name="' + name + '"\r\n\r\n' +
        String(value) + '\r\n'
      );
    });

    (attachments || []).forEach((blob, idx) => {
      if (!blob) return;
      if (this.progress) this.progress.update(88 + Math.min(5, idx), '하이웍스 전송 데이터 구성 중: ' + (blob.getName() || 'attachment'));

      const fileName = this.sanitizeAttachmentFileName_(blob.getName() || ('attachment_' + (idx + 1)));
      const contentType = blob.getContentType() || 'application/octet-stream';
      const fieldName = CONFIG.HIWORKS.FILE_FIELD_NAME || 'files[]';

      appendText(
        '--' + boundary + '\r\n' +
        'Content-Disposition: form-data; name="' + fieldName + '"; filename="' + fileName + '"\r\n' +
        'Content-Type: ' + contentType + '\r\n' +
        'Content-Transfer-Encoding: binary\r\n\r\n'
      );
      appendBytes(blob.getBytes());
      appendText('\r\n');
      fileNames.push(fileName);
      fileFields.push(fieldName);
    });

    appendText('--' + boundary + '--\r\n');

    if (this.progress) this.progress.update(93, '하이웍스 서버 응답 대기 중');
    const res = UrlFetchApp.fetch(CONFIG.HIWORKS.ENDPOINT, {
      method: 'post',
      contentType: 'multipart/form-data; boundary=' + boundary,
      headers: { Authorization: 'Bearer ' + token },
      payload: payload,
      muteHttpExceptions: true
    });

    return this.handleHiworksResponse_(res, fields, fileNames, {
      mode: 'manual_multipart_legacy_raw_filename_v55',
      payloadBytes: payload.length,
      fileFields: fileFields,
      note: 'v55: 네이버 수신 정상 확인된 기존 개발자 방식. filename* / RFC2047 제거, raw filename + Content-Transfer-Encoding: binary 사용.'
    });
  }

  buildUtf8FileContentDisposition_(fieldName, fileName) {
    const safeFieldName = String(fieldName || 'files[]').replace(/[\r\n"]/g, '');
    const safeFileName = this.sanitizeAttachmentFileName_(fileName || 'attachment');

    // v55 보조함수:
    // 하이웍스 기존 성공 코드와 동일하게 filename 파라미터 하나만 사용합니다.
    // filename-star / RFC2047 인코딩값은 하이웍스에서 실제 파일명으로 그대로 노출되어 사용하지 않습니다.
    return 'Content-Disposition: form-data; name="' + safeFieldName + '"; filename="' + safeFileName + '"';
  }

  buildAsciiAttachmentFallbackFileName_(fileName) {
    const original = String(fileName || 'attachment').trim() || 'attachment';
    const ext = this.extractSafeAsciiExtension_(original);
    const baseWithoutExt = ext ? original.slice(0, original.length - ext.length) : original;

    // 원래 파일명에 영문/숫자가 있으면 일부만 살리고, 한글/특수문자는 제거합니다.
    let asciiBase = baseWithoutExt
      .normalize('NFKD')
      .replace(/[^A-Za-z0-9._-]+/g, '_')
      .replace(/_+/g, '_')
      .replace(/^[_\.\-]+|[_\.\-]+$/g, '')
      .slice(0, 40);

    if (!asciiBase) asciiBase = 'attachment';

    // 한글 파일들이 모두 attachment.pdf로 겹치지 않도록 원본 파일명 기반 짧은 해시를 붙입니다.
    const hash = this.shortHashForHeader_(original);
    return asciiBase + '_' + hash + (ext || '');
  }

  extractSafeAsciiExtension_(fileName) {
    const match = String(fileName || '').match(/(\.[A-Za-z0-9]{1,10})$/);
    if (!match) return '';
    return match[1].toLowerCase();
  }

  shortHashForHeader_(value) {
    const text = String(value || '');
    let hash = 2166136261;
    for (let i = 0; i < text.length; i++) {
      hash ^= text.charCodeAt(i);
      hash += (hash << 1) + (hash << 4) + (hash << 7) + (hash << 8) + (hash << 24);
    }
    return ('00000000' + (hash >>> 0).toString(16)).slice(-8);
  }


  encodeRfc2047FileName_(value) {
    const text = String(value || 'attachment');

    // RFC 2047 encoded-word 형식입니다.
    // JavaMail의 MimeUtility.encodeText(value, "UTF-8", "B")와 같은 계열입니다.
    // Apps Script에는 MimeUtility가 없으므로 UTF-8 byte[]를 직접 Base64 인코딩합니다.
    const bytes = Utilities.newBlob(text, 'text/plain; charset=UTF-8').getBytes();
    const encoded = Utilities.base64Encode(bytes);
    return '=?UTF-8?B?' + encoded + '?=';
  }

  escapeMultipartHeaderValue_(value) {
    return String(value || '')
      .replace(/[\r\n]/g, ' ')
      .replace(/\\/g, '\\\\')
      .replace(/"/g, '\\"');
  }

  encodeRfc5987Value_(value) {
    return encodeURIComponent(String(value || ''))
      .replace(/['()]/g, function(ch) {
        return '%' + ch.charCodeAt(0).toString(16).toUpperCase();
      })
      .replace(/\*/g, '%2A');
  }

  handleHiworksResponse_(res, fields, fileNames, extraDebug) {
    const code = res.getResponseCode();
    const text = res.getContentText();
    let parsed = null;
    try { parsed = JSON.parse(text); } catch (e) { parsed = null; }

    const ok = this.isSuccess_(code, parsed, fields.to);
    const result = {
      ok,
      code,
      body: text,
      parsed,
      attachmentCount: fileNames.length,
      payloadDebug: Object.assign({
        to: fields.to,
        cc: fields.cc || '',
        user_id: fields.user_id,
        subject: fields.subject,
        save_sent_mail: fields.save_sent_mail,
        attachmentCount: fileNames.length,
        attachmentNames: fileNames
      }, extraDebug || {})
    };

    if (!ok) {
      throw new Error(
        '하이웍스 첨부 발송 실패\n' +
        'HTTP: ' + code + '\n' +
        '응답본문: ' + text.slice(0, 1500) + '\n' +
        '전송값: ' + JSON.stringify(result.payloadDebug) + '\n' +
        '참고: v42는 Apps Script 내부에서 단일/전체 첨부 용량 초과로 사전 차단하지 않습니다. ' +
        'native_multipart 모드에서도 하이웍스가 용량 오류를 반환하면 sendMail API 자체가 일반첨부만 받고 별도 대용량첨부 API가 필요한 구조입니다.'
      );
    }

    return result;
  }

  getToken_() {
    const info = findHiworksTokenInfo_();
    if (!info.token) {
      throw new Error(
        '하이웍스 토큰이 저장되어 있지 않습니다.\n' +
        '성공했던 기존 코드와 동일하게 HIWORKS_API_KEY를 사용합니다.\n' +
        '시트 메뉴 [메일자동화 > 하이웍스 API키 저장]에서 OfficeToken/API Key를 1회 저장하세요.\n' +
        '호환 검색 키: HIWORKS_API_KEY, HIWORKS_ACCESS_TOKEN, HIWORKS_OFFICE_TOKEN, HIWORKS_TOKEN'
      );
    }
    return info.token;
  }

  validateFields_(fields) {
    if (!fields.to) throw new Error('하이웍스 첨부 발송 실패: 받는 사람(to)이 비어 있습니다.');
    if (!fields.user_id) throw new Error('하이웍스 첨부 발송 실패: 보내는 사람 user_id가 비어 있습니다.');
    if (!fields.subject) throw new Error('하이웍스 첨부 발송 실패: 제목(subject)이 비어 있습니다.');
    if (!fields.content) throw new Error('하이웍스 첨부 발송 실패: 본문(content)이 비어 있습니다.');
  }

  validateAttachments_(attachments) {
    // v43:
    // 20MB 초과 파일이 Drive 링크로 전환되면 실제 API 첨부파일이 0개일 수도 있습니다.
    // sendMail API는 본문만으로도 발송 가능하므로 여기서 0개를 실패 처리하지 않습니다.
    const list = attachments || [];
    Logger.log(
      '하이웍스 첨부파일 확인 완료: ' +
      list.length + '개 / 파일명: ' +
      list.map(b => (b && b.getName ? b.getName() : 'attachment')).join(', ')
    );
  }

  normalizeUserId_(value) {
    const v = String(value || '').trim();
    if (!v) return '';
    return v.indexOf('@') >= 0 ? v.split('@')[0].trim() : v;
  }

  normalizeAddressList_(value) {
    const values = Array.isArray(value) ? value : String(value || '').split(/[,\n;]/);
    return values
      .map(v => String(v || '').trim())
      .filter(Boolean)
      .join(',');
  }

  sanitizeAttachmentFileName_(name) {
    return String(name || 'attachment')
      .replace(/[\\/:*?"<>|]/g, '_')
      .replace(/\r|\n/g, '')
      .trim();
  }

  isSuccess_(httpCode, parsed, to) {
    if (httpCode < 200 || httpCode >= 300) return false;
    if (!parsed) return false;
    if (String(parsed.code || '').toUpperCase() !== 'SUC') return false;

    const result = parsed.result || {};
    const wrongList = Array.isArray(result.wrongList) ? result.wrongList : [];
    if (wrongList.length > 0) return false;

    return true;
  }

  formatBytes_(bytes) {
    const n = Number(bytes) || 0;
    if (n >= 1024 * 1024) return (n / 1024 / 1024).toFixed(2) + ' MB';
    if (n >= 1024) return (n / 1024).toFixed(2) + ' KB';
    return n + ' B';
  }
}

class HeaderMapper {
  constructor(map, rawHeaders) {
    this.map = map;
    this.rawHeaders = rawHeaders || [];
  }

  static fromSheet(sheet, headerRow) {
    const headers = sheet.getRange(headerRow, 1, 1, sheet.getLastColumn()).getValues()[0];
    return HeaderMapper.fromHeaders(headers);
  }

  static fromHeaders(headers) {
    const map = {};
    headers.forEach((h, idx) => {
      const key = normalizeHeader_(h);
      if (key && !map[key]) map[key] = idx + 1;
    });
    return new HeaderMapper(map, headers);
  }

  findCol(header) {
    const key = normalizeHeader_(header);
    if (this.map[key]) return this.map[key];
    const aliasKey = this.findAliasKey_(key);
    if (aliasKey && this.map[aliasKey]) return this.map[aliasKey];

    // 예: target '최종견적가' vs source '최종견적가(부가세별도기준)'
    const keys = Object.keys(this.map);
    const starts = keys.find(k => k.indexOf(key) === 0 || key.indexOf(k) === 0);
    return starts ? this.map[starts] : null;
  }

  findAliasKey_(key) {
    const aliases = {
      [normalizeHeader_('최종 견적가')]: [normalizeHeader_('최종 견적가\n(부가세 별도 기준)')],
      [normalizeHeader_('부가세')]: [normalizeHeader_('부가세 포함\n(별도시 체크X)')]
    };
    const list = aliases[key] || [];
    return list.find(k => this.map[k]);
  }
}


function mergeObjectsPreferNonEmpty_(baseObj, overrideObj) {
  const result = Object.assign({}, baseObj || {});
  Object.keys(overrideObj || {}).forEach(key => {
    const v = overrideObj[key];
    if (v !== null && v !== undefined && String(v).trim() !== '') {
      result[key] = v;
    } else if (!(key in result)) {
      result[key] = v;
    }
  });
  return result;
}


function applyServiceApplicationApplicantFallbackToTarget_(ss, targetData) {
  const data = Object.assign({}, targetData || {});
  const normalized = buildServiceApplicationApplicantValues_(data);

  Object.assign(data, normalized.targetFields);

  try {
    if (ss) {
      const targetSheet = mustGetSheet_(ss, CONFIG.SHEETS.TARGET);
      writeValuesToTargetRowByAliases_(targetSheet, CONFIG.ROWS.TARGET_HEADER, CONFIG.ROWS.TARGET_DATA_ROW, normalized.targetFieldsByAliases);
      SpreadsheetApp.flush();
    }
  } catch (err) {
    // 생성대상 보정 실패만으로 발송 전체를 막지는 않습니다.
    // 뒤의 용역신청서 시트 직접 보정에서 다시 반영합니다.
    Logger.log('생성대상 신청인 정보 fallback 반영 실패: ' + (err && err.stack || err));
  }

  return data;
}

function applyServiceApplicationApplicantFallbackToSheet_(sheet, targetData) {
  if (!sheet) return;

  const normalized = buildServiceApplicationApplicantValues_(targetData || {});

  // 신청인 영역만 대상으로 합니다. 아래 계약 수행사 영역의 대표자/사업자번호/주소는 건드리지 않습니다.
  const applicantRangeA1 = 'A14:M21';

  // 상호명(기관 및 법인명): 사업자등록증상 법인명 우선, 없으면 회사명 fallback
  setValueNearLabelInRange_(sheet, applicantRangeA1, [
    '상호명(기관 및 법인명)', '상호명\n(기관 및 법인명)', '상호명', '기관 및 법인명', '기관및법인명', '상호'
  ], normalized.applicantName, 0, 2);

  // 신청인 주소: 사업자등록증상 법인 주소 우선, 없으면 고객사 상세 주소 fallback
  setValueNearLabelInRange_(sheet, applicantRangeA1, [
    '주소', '소재지', '사업장 주소', '사업장소재지'
  ], normalized.applicantAddress, 0, 2);

  // 사업자등록증 세부정보는 fallback하지 않습니다. 실제 값이 없으면 공란입니다.
  setValueNearLabelInRange_(sheet, applicantRangeA1, [
    '사업자등록번호', '사업자 등록번호', '사업자번호', '등록번호'
  ], normalized.businessNo, 0, 2);

  setValueNearLabelInRange_(sheet, applicantRangeA1, [
    '대표자명', '대표자', '대표자 성명', '대표자성명'
  ], normalized.representativeName, 0, 2);

  setValueNearLabelInRange_(sheet, applicantRangeA1, [
    '업태'
  ], normalized.businessType, 0, 2);

  setValueNearLabelInRange_(sheet, applicantRangeA1, [
    '종목'
  ], normalized.businessItem, 0, 2);

  // 현재 용역신청서 표준 양식 기준 보정. 라벨 탐색 실패 대비용입니다.
  safeSetSheetValue_(sheet, 'F14', normalized.applicantName);
  safeSetSheetValue_(sheet, 'K16', normalized.applicantAddress);
  safeSetSheetValue_(sheet, 'F20', normalized.businessNo);
}

function applyServiceApplicationSpecialTermsToSheet_(sheet, targetData) {
  if (!sheet) return;

  const specialTerms = firstObjectValueByHeaderContains_(targetData || {}, '용역신청서특약사항');
  if (specialTerms === null || specialTerms === undefined || String(specialTerms).trim() === '') return;

  // 용역신청서 표준 양식의 특약사항 병합셀은 F38:K40이며, 병합셀 입력은 좌상단 F38에 넣으면 됩니다.
  safeSetSheetValue_(sheet, 'F38', specialTerms);
}

function buildServiceApplicationApplicantValues_(targetData) {
  const data = targetData || {};

  const fallbackCompany = firstObjectValueByHeaders_(data, [
    '회사명', '고객사명', '고객명', '건물명', '건물명정규화', '고객명/건물명'
  ]);

  const fallbackAddress = firstObjectValueByHeaders_(data, [
    '고객사 상세 주소', '고객사 상세주소', '고객사상세주소',
    '고객사 주소', '고객사주소', '상세주소', '상세 주소',
    '주소', '도로명주소', '선택도로명주소', '지번주소',
    '주소수동입력(도로명)', '주소수동입력도로명'
  ]);

  const certCompany = firstObjectValueByHeaders_(data, [
    '계약 당사자(사업자등록증상 법인명)', '계약당사자(사업자등록증상법인명)',
    '계약 당사자', '계약당사자',
    '사업자등록증상 법인명', '사업자등록증상법인명',
    '사업자등록증상 상호', '사업자등록증상 회사명',
    '사업자등록증상 인명(단체명)', '등록증상 법인명', '등록증상 상호',
    '상호(명칭)', '상호', '법인명', '인명(단체명)'
  ]);

  const certAddress = firstObjectValueByHeaders_(data, [
    '사업자등록증상 법인 주소', '사업자등록증상법인주소',
    '사업자등록증상 주소', '사업자등록증상 사업장 주소',
    '사업자등록증 주소', '사업자등록증주소', '등록증상 주소',
    '법인 주소', '사업장 주소', '사업장소재지', '소재지'
  ]);

  const businessNoRaw = firstObjectValueByHeaders_(data, [
    '사업자등록증상 사업자등록번호', '사업자등록증상사업자등록번호',
    '등록증상 사업자등록번호', '사업자등록번호', '사업자 등록번호', '사업자번호', '등록번호'
  ]);

  const representativeName = firstObjectValueByHeaders_(data, [
    '사업자등록증상 대표자', '사업자등록증상 대표자명', '사업자등록증상 대표자 성명',
    '등록증상 대표자', '등록증상 대표자명', '대표자명', '대표자 성명', '대표자성명', '대표자'
  ]);

  const businessType = firstObjectValueByHeaders_(data, [
    '사업자등록증상 업태', '등록증상 업태', '업태'
  ]);

  const businessItem = firstObjectValueByHeaders_(data, [
    '사업자등록증상 종목', '등록증상 종목', '종목'
  ]);

  const applicantName = certCompany || fallbackCompany;
  const applicantAddress = certAddress || fallbackAddress;
  const businessNo = normalizeKoreanBusinessNoForDisplay_(businessNoRaw);

  const targetFields = {
    '계약 당사자(사업자등록증상 법인명)': applicantName,
    '사업자등록증상 법인 주소': applicantAddress,
    '사업자등록번호': businessNo,
    '대표자명': representativeName,
    '업태': businessType,
    '종목': businessItem
  };

  const targetFieldsByAliases = [
    { aliases: ['계약 당사자(사업자등록증상 법인명)', '계약당사자(사업자등록증상법인명)', '계약 당사자', '계약당사자', '사업자등록증상 법인명', '사업자등록증상법인명', '사업자등록증상 상호', '사업자등록증상 회사명'], value: applicantName },
    { aliases: ['사업자등록증상 법인 주소', '사업자등록증상법인주소', '사업자등록증상 주소', '사업자등록증 주소', '사업자등록증주소', '등록증상 주소', '법인 주소'], value: applicantAddress },
    { aliases: ['사업자등록증상 사업자등록번호', '사업자등록증상사업자등록번호', '등록증상 사업자등록번호', '사업자등록번호', '사업자 등록번호', '사업자번호', '등록번호'], value: businessNo },
    { aliases: ['사업자등록증상 대표자', '사업자등록증상 대표자명', '사업자등록증상 대표자 성명', '등록증상 대표자', '등록증상 대표자명', '대표자명', '대표자 성명', '대표자성명', '대표자'], value: representativeName },
    { aliases: ['사업자등록증상 업태', '등록증상 업태', '업태'], value: businessType },
    { aliases: ['사업자등록증상 종목', '등록증상 종목', '종목'], value: businessItem }
  ];

  return {
    applicantName: applicantName,
    applicantAddress: applicantAddress,
    businessNo: businessNo,
    representativeName: representativeName,
    businessType: businessType,
    businessItem: businessItem,
    targetFields: targetFields,
    targetFieldsByAliases: targetFieldsByAliases
  };
}

function writeValuesToTargetRowByAliases_(sheet, headerRow, dataRow, items) {
  if (!sheet || !items || !items.length) return;
  const headerMap = HeaderMapper.fromSheet(sheet, headerRow);
  items.forEach(item => {
    const aliases = item.aliases || [];
    for (const alias of aliases) {
      const col = headerMap.findCol(alias);
      if (col) {
        sheet.getRange(dataRow, col).setValue(item.value == null ? '' : item.value);
        return;
      }
    }
  });
}

function setValueNearLabelInRange_(sheet, rangeA1, labelAliases, value, rowOffset, colOffset) {
  try {
    const range = sheet.getRange(rangeA1);
    const values = range.getDisplayValues();
    const wanted = (labelAliases || []).map(normalizeHeader_).filter(Boolean);
    if (!wanted.length) return false;

    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        const cellText = normalizeHeader_(values[r][c]);
        if (!cellText) continue;
        const matched = wanted.some(label => cellText === label || cellText.indexOf(label) >= 0 || label.indexOf(cellText) >= 0);
        if (!matched) continue;

        const targetRow = range.getRow() + r + (Number(rowOffset) || 0);
        const targetCol = range.getColumn() + c + (Number(colOffset) || 0);
        sheet.getRange(targetRow, targetCol).setValue(value == null ? '' : value);
        return true;
      }
    }
  } catch (err) {
    Logger.log('용역신청서 라벨 기반 값 반영 실패: ' + rangeA1 + ' / ' + (err && err.stack || err));
  }

  return false;
}

function safeSetSheetValue_(sheet, a1, value) {
  try {
    sheet.getRange(a1).setValue(value == null ? '' : value);
  } catch (err) {
    Logger.log('용역신청서 셀 직접 값 반영 실패: ' + a1 + ' / ' + (err && err.stack || err));
  }
}

function firstObjectValueByHeaders_(obj, headers) {
  const keys = Object.keys(obj || {});
  const normalizedKeys = keys.map(key => ({ raw: key, normalized: normalizeHeader_(key) }));

  for (const header of headers || []) {
    const wanted = normalizeHeader_(header);
    if (!wanted) continue;

    for (const item of normalizedKeys) {
      if (item.normalized === wanted || item.normalized.indexOf(wanted) === 0 || wanted.indexOf(item.normalized) === 0) {
        const value = obj[item.raw];
        if (!isBlankPreviewValue_(value)) return String(value).trim();
      }
    }
  }

  return '';
}

function normalizeKoreanBusinessNoForDisplay_(value) {
  if (isBlankPreviewValue_(value)) return '';
  const raw = String(value).trim();
  const digits = raw.replace(/\D/g, '');

  // 사업자등록번호는 10자리입니다. 고객번호/생년월일/임의 숫자 6자리 등이 들어오면 공란 처리합니다.
  if (digits.length !== 10) return '';
  return digits.slice(0, 3) + '-' + digits.slice(3, 5) + '-' + digits.slice(5);
}

class RowObject {
  constructor(headerMap, values, rowNo) {
    this.headerMap = headerMap;
    this.values = values;
    this.rowNo = rowNo;
  }

  static fromValues(headerMap, values, rowNo) {
    return new RowObject(headerMap, values, rowNo);
  }

  static fromHeadersAndValues(headers, values, rowNo) {
    return new RowObject(HeaderMapper.fromHeaders(headers), values, rowNo);
  }

  get(header) {
    const col = this.headerMap.findCol(header);
    if (!col) return '';
    return this.values[col - 1];
  }

  getFlexible(header) {
    const v = this.get(header);
    return v == null ? '' : v;
  }

  toPlainObject() {
    const obj = {};
    this.headerMap.rawHeaders.forEach((h, idx) => {
      if (h !== '' && h != null) obj[String(h)] = this.values[idx];
    });
    return obj;
  }
}


function buildLowerStringSet_(items) {
  const set = {};
  (items || []).forEach(function(item) {
    const key = String(item || '').trim().toLowerCase();
    if (key) set[key] = true;
  });
  return set;
}

function compareQuoteSheetNameToLabel_(sheetName) {
  const name = String(sheetName || '').trim();
  if (name.indexOf('(1)') >= 0 || name === '비교견적1' || name === '비교견적서1') return '비교견적서(1)';
  if (name.indexOf('(2)') >= 0 || name === '비교견적2' || name === '비교견적서2') return '비교견적서(2)';
  return name ? ('비교견적서 - ' + name) : '비교견적서';
}

function buildCompareQuoteLabelFromSheets_(sheets) {
  const labels = unique_((sheets || [])
    .map(function(item) { return compareQuoteSheetNameToLabel_(item && item.sheetName); })
    .filter(Boolean));
  return labels.length ? labels.join(', ') : '비교견적서';
}

function estimateSecondsForDefinitions_(defs) {
  let sec = 18; // 접수번호/복사/데이터로드/메일 API 기본 비용
  (defs || []).forEach(def => {
    if (def.type === 'sheet_pdf') sec += 8;
    else if (def.type === 'multi_sheet_pdf') sec += 14;
    else if (def.type === 'sheet_xlsx_values') sec += 12;
    else if (def.type === 'docx_template') sec += 10;
    else if (def.type === 'static_files') sec += 2;
    else if (def.type === 'drive_folder_files') sec += 4;
    else if (def.type === 'vendor_zip') sec += 2;
    else if (def.type === 'vendor_docx_template') sec += 10;
    else if (def.type === 'vendor_file') sec += 2;
    else sec += 6;
  });
  return sec;
}

function targetHeadersFromRegistration_(registration) {
  return registration.headers || [];
}

function buildContractConditionPreview_(rowObj) {
  const grade = formatPlainPreviewValue_(firstRowValueByHeaders_(rowObj, [
    '관리등급', '등급', '점검등급', '유지관리등급'
  ]));

  const discount = formatDiscountPreview_(firstRowValueByHeaders_(rowObj, [
    '할인율', '할인', '적용할인율'
  ]));

  const contractUnit = formatPlainPreviewValue_(firstRowValueByHeaders_(rowObj, [
    '계약단위', '계약기간', '점검기간', '계약 개월', '계약개월'
  ]));

  const appointment = formatAppointmentPreview_(firstRowValueByHeaders_(rowObj, [
    '관리자 선임여부', '관리자선임여부', '관리자 선임 여부', '선임여부', '관리자선임', '비상주 선임', '비상주선임'
  ]));

  const maintenance = formatCountPreview_(firstRowValueByHeaders_(rowObj, [
    '유지점검', '유지점검횟수', '유지 횟수', '유지횟수', '유지보수·관리 점검', '유지보수관리점검'
  ]), '유지');

  const performance = formatCountPreview_(firstRowValueByHeaders_(rowObj, [
    '성능점검', '성능점검횟수', '성능 횟수', '성능횟수', '정보통신설비 성능점검'
  ]), '성능');

  const vat = formatVatPreview_(firstRowValueByHeaders_(rowObj, [
    '부가세', '부가세 포함', '부가세포함', '부가세 포함\n(별도시 체크X)', 'VAT', '부가세여부'
  ]));

  const finalPrice = formatMoneyWonPreview_(firstRowValueByHeaders_(rowObj, [
    '최종견적가', '최종 견적가', '최종 견적가\n(부가세 별도 기준)', '견적가', '견적금액', '최종금액'
  ]));

  return [grade, discount, contractUnit, appointment, maintenance, performance, vat, finalPrice]
    .map(v => String(v == null ? '' : v).trim())
    .filter(Boolean)
    .join(' / ');
}

function firstRowValueByHeaders_(rowObj, headers) {
  for (const header of headers || []) {
    const value = rowObj.getFlexible(header);
    if (!isBlankPreviewValue_(value)) return value;
  }
  return '';
}

function firstRowValueByHeaderContains_(rowObj, headerNeedle) {
  if (!rowObj || !rowObj.headerMap || !Array.isArray(rowObj.headerMap.rawHeaders)) return '';

  const needle = normalizeHeader_(headerNeedle);
  if (!needle) return '';

  for (let i = 0; i < rowObj.headerMap.rawHeaders.length; i++) {
    const header = rowObj.headerMap.rawHeaders[i];
    if (!header) continue;

    const key = normalizeHeader_(header);
    if (key.indexOf(needle) < 0) continue;

    const value = rowObj.values ? rowObj.values[i] : '';
    if (!isBlankPreviewValue_(value)) return value;
  }

  return '';
}

function firstObjectValueByHeaderContains_(obj, headerNeedle) {
  const needle = normalizeHeader_(headerNeedle);
  if (!needle) return '';

  const data = obj || {};
  const keys = Object.keys(data);
  for (const key of keys) {
    if (normalizeHeader_(key).indexOf(needle) < 0) continue;

    const value = data[key];
    if (!isBlankPreviewValue_(value)) return value;
  }

  return '';
}

function isBlankPreviewValue_(value) {
  return value === null || value === undefined || String(value).trim() === '';
}

function formatPlainPreviewValue_(value) {
  if (isBlankPreviewValue_(value)) return '';
  return String(value).trim();
}

function formatDiscountPreview_(value) {
  if (isBlankPreviewValue_(value)) return '';

  if (typeof value === 'number') {
    const percent = value > 0 && value <= 1 ? value * 100 : value;
    return trimTrailingZero_(percent) + '% 할인';
  }

  const raw = String(value).trim();
  if (!raw) return '';
  if (raw.indexOf('할인') >= 0) return raw;

  const numText = raw.replace(/,/g, '').replace(/%/g, '').trim();
  const num = Number(numText);
  if (Number.isFinite(num)) {
    const percent = num > 0 && num <= 1 ? num * 100 : num;
    return trimTrailingZero_(percent) + '% 할인';
  }

  return raw;
}

function formatAppointmentPreview_(value) {
  if (isBlankPreviewValue_(value)) return '';
  if (value === true) return '선임 O';
  if (value === false) return '선임 X';

  const raw = String(value).trim();
  const normalized = raw.toLowerCase().replace(/\s+/g, '');

  if (/^(o|ok|y|yes|true|1)$/i.test(raw) || normalized.indexOf('선임') >= 0 && normalized.indexOf('미선임') < 0) {
    return '선임 O';
  }

  if (/^(x|n|no|false|0)$/i.test(raw) || normalized.indexOf('미선임') >= 0 || normalized.indexOf('없') >= 0 || normalized.indexOf('아니') >= 0) {
    return '선임 X';
  }

  return raw.indexOf('선임') >= 0 ? raw : '선임 ' + raw;
}

function formatCountPreview_(value, label) {
  if (isBlankPreviewValue_(value)) return '';
  if (typeof value === 'number') return label + ' ' + trimTrailingZero_(value) + '회';

  const raw = String(value).trim();
  if (!raw) return '';
  if (raw.indexOf(label) === 0) return raw;
  if (raw.indexOf('회') >= 0) return label + ' ' + raw;

  const num = Number(raw.replace(/,/g, ''));
  if (Number.isFinite(num)) return label + ' ' + trimTrailingZero_(num) + '회';

  return label + ' ' + raw;
}

function formatVatPreview_(value) {
  if (isBlankPreviewValue_(value)) return '';
  if (value === true) return '부포';
  if (value === false) return '부별';

  const raw = String(value).trim();
  const normalized = raw.toLowerCase().replace(/\s+/g, '');

  if (normalized.indexOf('포함') >= 0 || /^(o|ok|y|yes|true|1)$/i.test(raw)) return '부포';
  if (normalized.indexOf('별도') >= 0 || /^(x|n|no|false|0)$/i.test(raw)) return '부별';

  return raw;
}

function formatMoneyWonPreview_(value) {
  if (isBlankPreviewValue_(value)) return '';

  const n = parseMoneyNumber_(value);
  if (n > 0) return Math.round(n).toLocaleString('ko-KR') + '원';

  const raw = String(value).trim();
  return raw;
}

function trimTrailingZero_(value) {
  const n = Number(value);
  if (!Number.isFinite(n)) return String(value || '');
  return String(Math.round(n * 100) / 100).replace(/\.0+$/, '').replace(/(\.\d*?)0+$/, '$1');
}

function mustGetSheet_(ss, name) {
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('시트를 찾지 못했습니다: ' + name);
  return sheet;
}

function quoteSheetNameForA1_(sheetName) {
  return "'" + String(sheetName || '').replace(/'/g, "''") + "'";
}

function columnToLetter_(column) {
  let col = Math.max(1, Number(column) || 1);
  let letter = '';
  while (col > 0) {
    const mod = (col - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    col = Math.floor((col - mod) / 26);
  }
  return letter;
}

function normalizeHeader_(header) {
  return String(header == null ? '' : header)
    .replace(/\r?\n/g, '')
    .replace(/\s+/g, '')
    .trim();
}

function isChecked_(value) {
  if (value === true) return true;
  if (Number(value) === 1) return true;
  const s = String(value || '').trim().toLowerCase();
  return ['true', 'y', 'yes', 'o', 'ok', '체크', '1'].indexOf(s) >= 0;
}

function splitEmails_(text) {
  return uniqueEmails_(String(text || '')
    .split(/[\n,;\s]+/)
    .map(s => s.trim())
    .filter(Boolean));
}

function uniqueEmails_(emails) {
  const seen = {};
  return (emails || [])
    .map(e => String(e || '').trim())
    .filter(Boolean)
    .filter(e => {
      const key = e.toLowerCase();
      if (seen[key]) return false;
      seen[key] = true;
      return true;
    });
}

function buildLowerEmailSet_(emails) {
  const set = {};
  (emails || []).forEach(function(email) {
    const key = String(email || '').trim().toLowerCase();
    if (key) set[key] = true;
  });
  return set;
}

function normalizeAdditionalEmailList_(input, defaultDomain) {
  let list = [];

  if (Array.isArray(input)) {
    list = input;
  } else {
    list = String(input || '').split(/[\n,;\s]+/);
  }

  return uniqueEmails_(list
    .map(item => normalizeLooseEmail_(item, defaultDomain))
    .filter(Boolean));
}

function normalizeLooseEmail_(input, defaultDomain) {
  let s = String(input || '').trim();
  if (!s) return '';

  if (s.indexOf('@') < 0) {
    s = s + '@' + defaultDomain;
  }

  if (!/^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(s)) {
    throw new Error('이메일 주소 형식이 올바르지 않습니다: ' + input);
  }

  return s;
}

function normalizeTestEmail_(input, domain) {
  let s = String(input || '').trim();
  if (!s) throw new Error('테스트 발송 이메일 아이디를 입력하세요.');
  if (s.indexOf('@') < 0) s = s + '@' + domain;
  const re = new RegExp('^[A-Za-z0-9._%+\\-]+@' + domain.replace('.', '\\.') + '$');
  if (!re.test(s)) throw new Error('테스트 발송 주소는 @' + domain + ' 형식이어야 합니다.');
  return s;
}


function rangePixelRect_(sheet, range, paddingPx) {
  const topLeft = cellTopLeftPixel_(sheet, range.getRow(), range.getColumn());
  const box = rangePixelBox_(sheet, range);
  const pad = Number(paddingPx) || 0;

  return {
    x: topLeft.x - pad,
    y: topLeft.y - pad,
    width: box.width + pad * 2,
    height: box.height + pad * 2
  };
}

function imagePixelRect_(sheet, img) {
  const anchor = img.getAnchorCell ? img.getAnchorCell() : null;
  if (!anchor) {
    throw new Error('이미지 anchor cell을 확인할 수 없습니다.');
  }

  const topLeft = cellTopLeftPixel_(sheet, anchor.getRow(), anchor.getColumn());
  const xOffset = img.getAnchorCellXOffset ? Number(img.getAnchorCellXOffset()) || 0 : 0;
  const yOffset = img.getAnchorCellYOffset ? Number(img.getAnchorCellYOffset()) || 0 : 0;
  const width = img.getWidth ? Number(img.getWidth()) || 0 : 0;
  const height = img.getHeight ? Number(img.getHeight()) || 0 : 0;

  return {
    x: topLeft.x + xOffset,
    y: topLeft.y + yOffset,
    width: width,
    height: height
  };
}

function cellTopLeftPixel_(sheet, row, col) {
  let x = 0;
  for (let c = 1; c < col; c++) {
    x += sheet.getColumnWidth(c);
  }

  let y = 0;
  for (let r = 1; r < row; r++) {
    y += sheet.getRowHeight(r);
  }

  return { x: x, y: y };
}

function rectsOverlap_(a, b) {
  if (!a || !b) return false;
  if (a.width <= 0 || a.height <= 0 || b.width <= 0 || b.height <= 0) return false;

  return !(
    a.x + a.width < b.x ||
    b.x + b.width < a.x ||
    a.y + a.height < b.y ||
    b.y + b.height < a.y
  );
}

function isCellInsideExpandedRange_(cell, range, rowPadding, colPadding) {
  const row = cell.getRow();
  const col = cell.getColumn();
  const rp = Number(rowPadding) || 0;
  const cp = Number(colPadding) || 0;

  return row >= range.getRow() - rp &&
    row < range.getRow() + range.getNumRows() + rp &&
    col >= range.getColumn() - cp &&
    col < range.getColumn() + range.getNumColumns() + cp;
}

function rangePixelBox_(sheet, range) {
  let width = 0;
  for (let c = range.getColumn(); c < range.getColumn() + range.getNumColumns(); c++) {
    width += sheet.getColumnWidth(c);
  }
  let height = 0;
  for (let r = range.getRow(); r < range.getRow() + range.getNumRows(); r++) {
    height += sheet.getRowHeight(r);
  }
  return { width, height };
}

function isCellInsideRange_(cell, range) {
  const row = cell.getRow();
  const col = cell.getColumn();
  return row >= range.getRow() &&
    row < range.getRow() + range.getNumRows() &&
    col >= range.getColumn() &&
    col < range.getColumn() + range.getNumColumns();
}

function getDriveImageThumbnailBlob_(fileId, maxWidthPx) {
  const cache = CacheService.getScriptCache();
  const cacheKey = 'IMG_THUMB_OK_' + fileId + '_' + maxWidthPx;
  const cached = cache.get(cacheKey);

  // thumbnail URL 자체는 캐시에 오래 저장하지 않고, 성공 여부만 캐시합니다.
  // 실제 blob은 매번 가져오는 쪽이 Drive 권한/만료 URL 문제를 피하기 쉽습니다.
  const url = 'https://drive.google.com/thumbnail?id=' +
    encodeURIComponent(fileId) +
    '&sz=w' +
    encodeURIComponent(String(maxWidthPx || 600));

  const res = UrlFetchApp.fetch(url, {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    followRedirects: true,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const blob = res.getBlob();
  const contentType = String(blob.getContentType() || '').toLowerCase();

  if (code >= 200 && code < 300 && contentType.indexOf('image/') === 0 && blob.getBytes().length <= 2 * 1024 * 1024) {
    cache.put(cacheKey, '1', 21600);
    return blob.setName('asset_' + fileId + '.png');
  }

  // thumbnail fetch가 막히는 파일이면 원본으로 한 번만 시도합니다.
  // 원본도 크면 아래에서 명확한 안내를 내보냅니다.
  const originalBlob = DriveApp.getFileById(fileId).getBlob();
  const originalSize = originalBlob.getBytes().length;

  if (originalSize > 2 * 1024 * 1024) {
    throw new Error(
      '도장/로고 이미지가 너무 큽니다. Drive thumbnail 생성도 실패했습니다. ' +
      '파일ID=' + fileId + ', 원본크기=' + Math.round(originalSize / 1024) + 'KB. ' +
      '이미지를 1000px 이하 PNG/JPG로 줄여 다시 올리거나, 해당 파일 공유 권한을 확인하세요.'
    );
  }

  return originalBlob;
}

function getTodayFileDateSuffix_() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyMMdd');
}

function appendDateSuffixBeforeExtension_(filename, dateSuffix) {
  const raw = String(filename || 'attachment').trim() || 'attachment';
  const suffix = String(dateSuffix || '').trim();
  if (!suffix) return raw;

  // 이미 _260613 같은 날짜 suffix가 붙어 있으면 중복으로 붙이지 않습니다.
  const already = new RegExp('_' + suffix + '(\\.[^\\./]+)?$');
  if (already.test(raw)) return raw;

  const m = raw.match(/^(.*?)(\.[^\./]+)$/);
  if (!m) return raw + '_' + suffix;
  return m[1] + '_' + suffix + m[2];
}


function sha1Hex_(text) {
  const bytes = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_1, String(text || ''), Utilities.Charset.UTF_8);
  return bytes.map(function(b) {
    const v = (b < 0 ? b + 256 : b).toString(16);
    return v.length === 1 ? '0' + v : v;
  }).join('');
}

function sanitizeFileName_(name) {
  return String(name || '')
    .replace(/[\\/:*?"<>|#%{}~&]/g, '_')
    .replace(/\s+/g, ' ')
    .trim()
    .slice(0, 120) || 'file';
}

function unique_(arr) {
  const seen = {};
  return arr.filter(v => {
    if (seen[v]) return false;
    seen[v] = true;
    return true;
  });
}


function cleanContactText_(value) {
  return String(value || '')
    .replace(/\s+/g, ' ')
    .replace(/\s*님\s*$/, '')
    .trim();
}

function buildPersonTitleText_(name, title) {
  const n = String(name || '').trim();
  const t = String(title || '').trim();
  return [n, t].filter(Boolean).join(' ');
}

function normalizeVendorName_(vendorName) {
  const raw = String(vendorName || '').trim();
  const compact = normalizeHeader_(raw);

  if (!compact) return '';
  if (compact === normalizeHeader_('KJ') || compact.indexOf(normalizeHeader_('케이제이')) >= 0) return 'KJ';
  if (compact.indexOf(normalizeHeader_('일신')) >= 0) return '일신';
  if (compact.indexOf(normalizeHeader_('디엠')) >= 0 || compact === normalizeHeader_('DM')) return '디엠';
  if (compact.indexOf(normalizeHeader_('삼구')) >= 0) return '삼구';

  return raw;
}

function getVendorConfig_(vendorName) {
  const key = normalizeVendorName_(vendorName);
  return CONFIG.VENDORS[key] || CONFIG.VENDORS[String(vendorName || '').trim()] || null;
}

function getVendorDisplayName_(vendorName) {
  const key = normalizeVendorName_(vendorName);
  const vendor = getVendorConfig_(vendorName);
  if (vendor && vendor.displayName) return vendor.displayName;
  return key || String(vendorName || '').trim();
}

function joinKoreanList_(items) {
  const arr = (items || [])
    .map(v => String(v || '').trim())
    .filter(Boolean);
  if (arr.length <= 1) return arr.join('');
  if (arr.length === 2) return arr[0] + ' 및 ' + arr[1];
  return arr.slice(0, -1).join(', ') + ' 및 ' + arr[arr.length - 1];
}

function escapeHtml_(value) {
  return String(value == null ? '' : value)
    .replace(/&/g, '&amp;')
    .replace(/</g, '&lt;')
    .replace(/>/g, '&gt;')
    .replace(/"/g, '&quot;')
    .replace(/'/g, '&#39;');
}

function escapeRegex_(text) {
  return String(text).replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
}

function fixQuoteKoreanAmountErrors_(sheet) {
  const targets = [];

  // 1) 수식 자체에서 커스텀 함수 찾기
  ['numberToKorean', 'NUMBERSTRING'].forEach(function(keyword) {
    try {
      const found = sheet.createTextFinder(keyword).matchFormulaText(true).findAll();
      found.forEach(cell => targets.push(cell));
    } catch (e) {}
  });

  // 2) 표시값으로 에러 찾기
  ['#ERROR!', '#ERROR', '#NAME?', '#NAME'].forEach(function(keyword) {
    try {
      const found = sheet.createTextFinder(keyword).findAll();
      found.forEach(cell => targets.push(cell));
    } catch (e) {}
  });

  const unique = {};
  targets.forEach(function(cell) {
    const key = cell.getRow() + ':' + cell.getColumn();
    if (unique[key]) return;
    unique[key] = true;

    const amount = findAmountForKoreanTextCell_(sheet, cell);
    if (amount > 0) {
      cell.setNumberFormat('@');
      cell.setValue(numberToKoreanWonJeong_(amount));
    }
  });
}

function findAmountForKoreanTextCell_(sheet, cell) {
  const row = cell.getRow();
  const lastCol = sheet.getLastColumn();
  let maxAmount = 0;

  // 같은 행에서 가장 큰 금액 숫자 찾기
  // 예: "일금 #ERROR! ₩2,450,000" 행에서는 2,450,000이 잡힘
  for (let col = 1; col <= lastCol; col++) {
    const value = sheet.getRange(row, col).getValue();
    const display = sheet.getRange(row, col).getDisplayValue();

    const candidates = [value, display];
    candidates.forEach(function(v) {
      const n = parseMoneyNumber_(v);
      if (n > maxAmount) maxAmount = n;
    });
  }

  // 같은 행에서 못 찾으면 근처 아래/위 행도 보수적으로 검색
  if (maxAmount <= 0) {
    const startRow = Math.max(1, row - 2);
    const endRow = Math.min(sheet.getLastRow(), row + 2);

    for (let r = startRow; r <= endRow; r++) {
      for (let c = 1; c <= lastCol; c++) {
        const n = parseMoneyNumber_(sheet.getRange(r, c).getDisplayValue());
        if (n > maxAmount) maxAmount = n;
      }
    }
  }

  return maxAmount;
}

function parseMoneyNumber_(value) {
  if (typeof value === 'number') return value > 0 ? value : 0;

  const s = String(value || '')
    .replace(/₩/g, '')
    .replace(/\\/g, '')
    .replace(/원/g, '')
    .replace(/,/g, '')
    .replace(/\s/g, '')
    .trim();

  const m = s.match(/\d+/);
  if (!m) return 0;

  const n = Number(m[0]);
  return Number.isFinite(n) ? n : 0;
}

function numberToKoreanWonJeong_(amount) {
  let n = Math.floor(Number(amount) || 0);
  if (n <= 0) return '';

  const digits = ['', '일', '이', '삼', '사', '오', '육', '칠', '팔', '구'];
  const smallUnits = ['', '십', '백', '천'];
  const bigUnits = ['', '만', '억', '조'];

  let result = '';
  let bigIdx = 0;

  while (n > 0) {
    const chunk = n % 10000;

    if (chunk > 0) {
      let chunkText = '';
      let c = chunk;

      for (let i = 0; i < 4; i++) {
        const d = c % 10;
        if (d > 0) {
          const digitText = (d === 1 && i > 0) ? '' : digits[d];
          chunkText = digitText + smallUnits[i] + chunkText;
        }
        c = Math.floor(c / 10);
      }

      result = chunkText + bigUnits[bigIdx] + result;
    }

    n = Math.floor(n / 10000);
    bigIdx++;
  }

  return result + '원정';
}

/*******************************************************
 * v86 발송파일 고객사 폴더 누적 저장 + 발송파일로그
 * - 기존 하이웍스 발송/첨부 생성/Worker 브릿지 로직은 수정하지 않고,
 *   HiworksMailer.send(mail) 성공 직후에만 호출됩니다.
 * - 고객사 폴더는 공유드라이브명 + 고객번호 prefix 폴더명 검색을 우선 사용합니다.
 * - 마스터시트 BJ열 고객사폴더ID는 보조 캐시로만 사용합니다.
 * - TEST 모드는 기본 skip합니다.
 * - 같은 파일명/같은 내용이어도 발송 이력 보존을 위해 항상 누적 저장합니다.
 *******************************************************/

function setSentFileArchiveRootFolderId() {
  const ui = SpreadsheetApp.getUi();
  const current = getSentFileArchiveRootFolderId_();
  const res = ui.prompt(
    '발송파일 저장 fallback 루트 폴더 ID 저장',
    '기본 운영은 공유드라이브명과 고객번호_ 폴더명으로 찾습니다.\n고객사 폴더를 못 찾는 예외 상황에서만 쓸 fallback 루트 폴더 ID를 입력하세요.\n\n' +
    '현재값: ' + (current || '(없음)') + '\n\n' +
    '폴더 URL에서 /folders/ 뒤의 ID만 붙여넣으면 됩니다.',
    ui.ButtonSet.OK_CANCEL
  );

  if (res.getSelectedButton() !== ui.Button.OK) {
    ui.alert('저장을 취소했습니다.');
    return;
  }

  const id = extractDriveId_(res.getResponseText());
  if (!id) {
    ui.alert('폴더 ID가 비어 있습니다.');
    return;
  }

  PropertiesService.getScriptProperties().setProperty(
    getSentFileArchiveConfig_().ROOT_FOLDER_ID_PROPERTY || 'S1_CUSTOMER_FILE_ROOT_FOLDER_ID',
    id
  );

  let msg = '발송파일 저장 fallback 루트 폴더 ID 저장 완료\n\n' + id;
  try {
    const folder = DriveApp.getFolderById(id);
    msg += '\n\n폴더명: ' + folder.getName();
  } catch (err) {
    msg += '\n\n주의: 폴더 접근 확인 실패 - ' + String(err && err.message || err);
  }
  ui.alert(msg);
}

function checkSentFileArchiveConfig() {
  const cfg = getSentFileArchiveConfig_();
  let driveId = '';
  let driveStatus = '미확인';

  try {
    driveId = sfaGetSharedDriveId_(cfg);
    driveStatus = '공유드라이브 이름 조회 성공';
  } catch (err) {
    driveStatus = '공유드라이브 이름 조회 실패: ' + String(err && err.message || err);
  }

  const fallbackId = getSentFileArchiveRootFolderId_();
  let fallbackStatus = fallbackId ? 'fallback ID 있음' : 'fallback 미사용';
  if (fallbackId) {
    try {
      const folder = DriveApp.getFolderById(fallbackId);
      fallbackStatus = 'fallback 접근 가능: ' + folder.getName();
    } catch (err) {
      fallbackStatus = 'fallback 접근 실패: ' + String(err && err.message || err);
    }
  }

  SpreadsheetApp.getUi().alert(
    '발송파일 저장 설정\n\n' +
    'ENABLED: ' + cfg.ENABLED + '\n' +
    'SAVE_TEST_MAIL: ' + cfg.SAVE_TEST_MAIL + '\n' +
    'SENT_PREFIX: ' + cfg.SENT_PREFIX + '\n' +
    '폴더 결정 방식: 공유드라이브명 + 고객번호_ 폴더명 검색 우선\n' +
    '공유드라이브명: ' + cfg.SHARED_DRIVE_NAME + '\n' +
    '공유드라이브ID: ' + (driveId || '(없음)') + '\n' +
    '공유드라이브 상태: ' + driveStatus + '\n' +
    '마스터 폴더ID 헤더: ' + (cfg.MASTER_FOLDER_ID_HEADERS || []).join(', ') + '\n' +
    'fallback ROOT_FOLDER_ID: ' + (fallbackId || '(없음)') + '\n' +
    'fallback 상태: ' + fallbackStatus + '\n\n' +
    '중앙 로그 시트: ' + cfg.LOG_SHEET_NAME + '\n' +
    '폴더 인덱스 시트: ' + cfg.FOLDER_INDEX_SHEET_NAME + '\n' +
    '일일 이력 반영 시간: ' + cfg.DAILY_HISTORY_SYNC_HOUR + '시'
  );
}

function installSentFileHistoryDailyTrigger() {
  const cfg = getSentFileArchiveConfig_();
  const handler = 'syncSentFileFolderHistoryDaily';

  ScriptApp.getProjectTriggers().forEach(function(trigger) {
    if (trigger.getHandlerFunction && trigger.getHandlerFunction() === handler) {
      ScriptApp.deleteTrigger(trigger);
    }
  });

  ScriptApp.newTrigger(handler)
    .timeBased()
    .everyDays(1)
    .atHour(Number(cfg.DAILY_HISTORY_SYNC_HOUR) || 19)
    .create();

  SpreadsheetApp.getUi().alert(
    '발송이력 일일반영 트리거 설치 완료\n\n' +
    '실행 함수: ' + handler + '\n' +
    '실행 시간: 매일 ' + (Number(cfg.DAILY_HISTORY_SYNC_HOUR) || 19) + '시대'
  );
}

function syncSentFileFolderHistoryDaily() {
  return new SentFileArchiveService(null).syncFolderHistoryFromCentralLog_();
}

function summarizeSentFileArchiveResult_(result) {
  if (!result) return '결과없음';
  if (result.skipped) return 'SKIP(' + (result.message || result.reason || '') + ')';
  if (result.ok) return String(result.savedCount || 0) + '개 저장';
  return '저장실패(' + String(result.message || result.error || '').slice(0, 120) + ')';
}

function getSentFileArchiveConfig_() {
  const defaults = {
    ENABLED: false,
    SHARED_DRIVE_NAME: 'S1 고객사 파일 관리',
    SHARED_DRIVE_ID: '',
    SHARED_DRIVE_ID_PROPERTY: 'S1_CUSTOMER_SHARED_DRIVE_ID',
    MASTER_FOLDER_ID_HEADERS: ['고객사폴더ID', '고객사 폴더 ID', '고객폴더ID', '고객 폴더 ID', '폴더ID'],
    ROOT_FOLDER_ID: '',
    ROOT_FOLDER_ID_PROPERTY: 'S1_CUSTOMER_FILE_ROOT_FOLDER_ID',
    SAVE_TEST_MAIL: false,
    SENT_PREFIX: '[발송]',
    LOG_SHEET_NAME: '발송파일로그',
    FOLDER_INDEX_SHEET_NAME: '고객사파일폴더인덱스',
    FOLDER_HISTORY_SHEET_NAME: '_메일이력_발송',
    CREATE_FOLDER_IF_MISSING: true,
    ALWAYS_ACCUMULATE: true,
    CALCULATE_SHA256: true,
    ASYNC_AFTER_SEND: true,
    ASYNC_REVIEW_ONLY: true,
    QUEUE_SHEET_NAME: '발송파일저장큐',
    ASYNC_TRIGGER_HANDLER: 'processDeferredSentFileArchiveQueueV94',
    ASYNC_TRIGGER_DELAY_MS: 60 * 1000,
    MAX_ASYNC_JOBS_PER_RUN: 3,
    DAILY_HISTORY_SYNC_ENABLED: true,
    DAILY_HISTORY_SYNC_HOUR: 19,
    MAX_HISTORY_SYNC_ROWS_PER_RUN: 300
  };
  const cfg = (typeof CONFIG !== 'undefined' && CONFIG && CONFIG.SENT_FILE_ARCHIVE) ? CONFIG.SENT_FILE_ARCHIVE : {};
  return Object.assign({}, defaults, cfg);
}

function getSentFileArchiveRootFolderId_() {
  const cfg = getSentFileArchiveConfig_();
  const fromConfig = String(cfg.ROOT_FOLDER_ID || '').trim();
  if (fromConfig) return extractDriveId_(fromConfig);

  const propKey = String(cfg.ROOT_FOLDER_ID_PROPERTY || 'S1_CUSTOMER_FILE_ROOT_FOLDER_ID').trim();
  const fromProp = String(PropertiesService.getScriptProperties().getProperty(propKey) || '').trim();
  return extractDriveId_(fromProp);
}

function extractDriveId_(value) {
  const s = String(value || '').trim();
  if (!s) return '';
  const folderMatch = s.match(/\/folders\/([A-Za-z0-9_-]+)/);
  if (folderMatch) return folderMatch[1];
  const idMatch = s.match(/[?&]id=([A-Za-z0-9_-]+)/);
  if (idMatch) return idMatch[1];
  return s.replace(/[<>'"\s]/g, '');
}


function sfaGetSharedDriveId_(cfg) {
  cfg = cfg || getSentFileArchiveConfig_();

  if (String(cfg.SHARED_DRIVE_ID || '').trim()) {
    return String(cfg.SHARED_DRIVE_ID || '').trim();
  }

  const propKey = String(cfg.SHARED_DRIVE_ID_PROPERTY || 'S1_CUSTOMER_SHARED_DRIVE_ID').trim();
  const props = PropertiesService.getScriptProperties();
  const cached = String(props.getProperty(propKey) || '').trim();
  if (cached) return cached;

  const driveName = String(cfg.SHARED_DRIVE_NAME || 'S1 고객사 파일 관리').trim();
  const q = 'name = ' + sfaDriveQueryString_(driveName);
  const data = sfaDriveFetch_(
    'drives?pageSize=10&q=' + encodeURIComponent(q) + '&fields=drives(id,name)',
    { method: 'get' }
  );

  const drives = data.drives || [];
  if (!drives.length) {
    throw new Error('공유드라이브를 찾지 못했습니다: ' + driveName);
  }

  const driveId = drives[0].id;
  props.setProperty(propKey, driveId);
  return driveId;
}

function sfaFindChildFolder_(parentId, driveId, folderName) {
  if (!folderName) return null;

  const q = [
    sfaDriveQueryString_(parentId) + ' in parents',
    "mimeType = 'application/vnd.google-apps.folder'",
    'name = ' + sfaDriveQueryString_(folderName),
    'trashed = false'
  ].join(' and ');

  const path =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=drive' +
    '&driveId=' + encodeURIComponent(driveId) +
    '&pageSize=10' +
    '&q=' + encodeURIComponent(q) +
    '&fields=files(id,name,webViewLink,trashed,parents)';

  const data = sfaDriveFetch_(path, { method: 'get' });
  const files = data.files || [];
  return files.length ? files[0] : null;
}

function sfaFindCustomerFolderByPrefixInParent_(driveId, parentId, customerNo) {
  const no = sfaNormalizeCustomerNo_(customerNo);
  if (!no) return null;

  const prefix = no + '_';
  const q = [
    sfaDriveQueryString_(parentId) + ' in parents',
    "mimeType = 'application/vnd.google-apps.folder'",
    'name contains ' + sfaDriveQueryString_(prefix),
    'trashed = false'
  ].join(' and ');

  const path =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=drive' +
    '&driveId=' + encodeURIComponent(driveId) +
    '&pageSize=50' +
    '&q=' + encodeURIComponent(q) +
    '&fields=files(id,name,webViewLink,trashed,parents)';

  const data = sfaDriveFetch_(path, { method: 'get' });
  const files = (data.files || []).filter(function(file) {
    return String(file.name || '').indexOf(prefix) === 0;
  });

  if (!files.length) return null;
  files.sort(function(a, b) {
    return String(a.name || '').length - String(b.name || '').length;
  });
  return files[0];
}

function sfaCreateDriveFolder_(folderName, parentId) {
  return sfaDriveFetch_(
    'files?supportsAllDrives=true&fields=id,name,webViewLink,trashed,parents',
    {
      method: 'post',
      payload: {
        name: folderName,
        mimeType: 'application/vnd.google-apps.folder',
        parents: [parentId]
      }
    }
  );
}

function sfaDriveFetch_(path, options) {
  const url = 'https://www.googleapis.com/drive/v3/' + path;
  const params = Object.assign(
    {
      method: 'get',
      muteHttpExceptions: true,
      headers: {
        Authorization: 'Bearer ' + ScriptApp.getOAuthToken()
      }
    },
    options || {}
  );

  if (params.payload && typeof params.payload !== 'string') {
    params.contentType = 'application/json';
    params.payload = JSON.stringify(params.payload);
  }

  const res = UrlFetchApp.fetch(url, params);
  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('Drive API 오류 ' + code + ': ' + body);
  }
  return body ? JSON.parse(body) : {};
}

function sfaDriveQueryString_(value) {
  const s = String(value == null ? '' : value).trim()
    .replace(/\\/g, '\\\\')
    .replace(/'/g, "\\'");
  return "'" + s + "'";
}

function sfaNormalizeCustomerNo_(value) {
  const s = String(value == null ? '' : value).trim();
  if (!s) return '';
  const numeric = Number(s);
  if (Number.isFinite(numeric) && Math.floor(numeric) === numeric) return String(numeric);
  return s.replace(/\.0$/, '').trim();
}

function sfaNormalizeFolderText_(value) {
  return String(value || '')
    .toLowerCase()
    .replace(/주식회사|\(주\)|㈜|유한회사|재단법인|사단법인/g, '')
    .replace(/[\s_\-()\[\]{}.,·ㆍ]/g, '')
    .trim();
}

function sfaIsFolderNameSafeForCustomer_(folderName, customerNo) {
  const no = sfaNormalizeCustomerNo_(customerNo);
  if (!no) return true;

  const raw = String(folderName || '').trim();
  if (!raw) return false;

  // 기존 고객사 폴더명이 1079_회사명, 1079 회사명, 1079-회사명, [1079] 회사명 등으로 섞여 있어도 허용합니다.
  const escapedNo = escapeRegex_(no);
  const prefixRe = new RegExp('^(?:\\[|\\(|#)?\\s*' + escapedNo + '(?:\\]|\\))?\\s*[_\\- .)]');
  if (prefixRe.test(raw)) return true;

  // 폴더명 전체를 정규화했을 때 맨 앞이 고객번호면 허용합니다.
  const compact = sfaNormalizeFolderText_(raw);
  return compact.indexOf(no) === 0;
}

function sfaFolderCandidateScore_(file, customerNo, company) {
  const no = sfaNormalizeCustomerNo_(customerNo);
  const companyKey = sfaNormalizeFolderText_(company);
  const name = String(file && file.name || '').trim();
  const nameKey = sfaNormalizeFolderText_(name);
  if (!name) return 0;

  let score = 0;
  if (no) {
    if (sfaIsFolderNameSafeForCustomer_(name, no)) score += 300;
    else if (name.indexOf(no) >= 0 || nameKey.indexOf(no) >= 0) score += 120;
  }
  if (companyKey) {
    if (nameKey === companyKey) score += 160;
    else if (nameKey.indexOf(companyKey) >= 0 || companyKey.indexOf(nameKey) >= 0) score += 90;
  }

  // 캐시/작업공간/로그 폴더는 고객 폴더 후보에서 제외합니다.
  const bad = ['메일자동화', '작업공간', '캐시', 'shortcut', '바로가기', '발송파일로그', '고객사파일폴더인덱스'];
  const badHit = bad.some(function(word) { return name.indexOf(word) >= 0; });
  if (badHit) score -= 500;

  return Math.max(0, score);
}

function sfaFindCustomerFolderAnywhereInDrive_(driveId, customerNo, company) {
  const no = sfaNormalizeCustomerNo_(customerNo);
  const companyText = String(company || '').trim();
  const qParts = [
    "mimeType = 'application/vnd.google-apps.folder'",
    'trashed = false'
  ];

  const needles = [];
  if (no) needles.push(no);
  if (companyText) needles.push(companyText);
  const uniqueNeedles = needles.filter(function(v, i, arr) { return v && arr.indexOf(v) === i; });

  if (uniqueNeedles.length) {
    qParts.push('(' + uniqueNeedles.map(function(v) {
      return 'name contains ' + sfaDriveQueryString_(v);
    }).join(' or ') + ')');
  }

  const pathBase =
    'files' +
    '?supportsAllDrives=true' +
    '&includeItemsFromAllDrives=true' +
    '&corpora=drive' +
    '&driveId=' + encodeURIComponent(driveId) +
    '&pageSize=100' +
    '&q=' + encodeURIComponent(qParts.join(' and ')) +
    '&fields=nextPageToken,files(id,name,webViewLink,trashed,parents,mimeType)';

  const files = sfaDriveFetchAllFiles_(pathBase);
  if (!files.length) return null;

  const scored = files.map(function(file) {
    const score = sfaFolderCandidateScore_(file, no, companyText);
    return Object.assign({}, file, { __score: score });
  }).filter(function(file) {
    return file.__score > 0;
  });

  if (!scored.length) return null;
  scored.sort(function(a, b) {
    if (b.__score !== a.__score) return b.__score - a.__score;
    return String(a.name || '').length - String(b.name || '').length;
  });

  const picked = scored[0];
  picked.note = '공유드라이브 전체 검색 매칭(score=' + picked.__score + ')';
  return picked;
}

function sfaDriveFetchAllFiles_(pathBase) {
  const out = [];
  let pageToken = '';

  for (let guard = 0; guard < 20; guard++) {
    const sep = pathBase.indexOf('?') >= 0 ? '&' : '?';
    const path = pathBase + (pageToken ? sep + 'pageToken=' + encodeURIComponent(pageToken) : '');
    const data = sfaDriveFetch_(path, { method: 'get' });
    (data.files || []).forEach(function(file) { out.push(file); });
    pageToken = String(data.nextPageToken || '').trim();
    if (!pageToken) break;
  }

  return out;
}

function sfaFolderHasFileName_(folderId, fileName) {
  const id = String(folderId || '').trim();
  const name = String(fileName || '').trim();
  if (!id || !name) return false;

  try {
    const q = [
      sfaDriveQueryString_(id) + ' in parents',
      'name = ' + sfaDriveQueryString_(name),
      'trashed = false'
    ].join(' and ');
    const path =
      'files?supportsAllDrives=true' +
      '&includeItemsFromAllDrives=true' +
      '&pageSize=1' +
      '&q=' + encodeURIComponent(q) +
      '&fields=files(id,name)';
    const data = sfaDriveFetch_(path, { method: 'get' });
    return Boolean(data.files && data.files.length);
  } catch (err) {
    Logger.log('발송파일명 중복 확인 실패. 중복 없음으로 처리: ' + name + ' / ' + (err && err.stack || err));
    return false;
  }
}

function sfaCreateFileInFolderFromBlob_(folderId, blob, fileName) {
  const parentId = String(folderId || '').trim();
  if (!parentId) throw new Error('발송파일 저장 폴더ID가 비어 있습니다.');
  if (!blob) throw new Error('저장할 첨부 Blob이 없습니다.');

  const name = String(fileName || (blob.getName ? blob.getName() : '') || 'attachment').trim();
  const contentType = String(blob.getContentType ? blob.getContentType() : '' || 'application/octet-stream') || 'application/octet-stream';
  const boundary = '----S1SentFileArchive' + Utilities.getUuid().replace(/-/g, '');
  const delimiter = '\r\n--' + boundary + '\r\n';
  const closeDelimiter = '\r\n--' + boundary + '--';

  const metadata = {
    name: name,
    parents: [parentId]
  };

  const payload = [];
  function addText(text) {
    const bytes = Utilities.newBlob(String(text), 'text/plain; charset=UTF-8').getBytes();
    for (let i = 0; i < bytes.length; i++) payload.push(bytes[i]);
  }
  function addBytes(bytes) {
    for (let i = 0; i < bytes.length; i++) payload.push(bytes[i]);
  }

  addText(delimiter);
  addText('Content-Type: application/json; charset=UTF-8\r\n\r\n');
  addText(JSON.stringify(metadata));
  addText(delimiter);
  addText('Content-Type: ' + contentType + '\r\n');
  addText('Content-Transfer-Encoding: binary\r\n\r\n');
  addBytes(blob.getBytes());
  addText(closeDelimiter);

  const res = UrlFetchApp.fetch('https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&supportsAllDrives=true&fields=id,name,webViewLink,size,mimeType', {
    method: 'post',
    contentType: 'multipart/related; boundary=' + boundary,
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    payload: payload,
    muteHttpExceptions: true
  });

  const code = res.getResponseCode();
  const body = res.getContentText();
  if (code < 200 || code >= 300) {
    throw new Error('공유드라이브 발송파일 업로드 실패 HTTP ' + code + ': ' + String(body || '').slice(0, 1000));
  }

  const data = JSON.parse(body || '{}');
  if (!data.id) throw new Error('공유드라이브 발송파일 업로드 응답에 파일ID가 없습니다: ' + body);
  return DriveApp.getFileById(data.id);
}

/**
 * v94: CUSTOMER 발송 후 고객사 공유드라이브 누적 저장을 사용자 응답 경로에서 분리합니다.
 * - 하이웍스 발송 성공까지는 동기 처리합니다.
 * - 발송파일 누적 저장은 큐에 넣고 시간 기반 트리거가 뒤에서 처리합니다.
 * - 현재는 파일 확인/수정 후 발송(reviewPackage 보유) 건을 안전하게 비동기화합니다.
 */
function shouldDeferSentFileArchiveAfterSendV94_(mode, reviewPackage) {
  const cfg = getSentFileArchiveConfig_();
  if (!cfg.ENABLED || cfg.ASYNC_AFTER_SEND !== true) return false;
  if (String(mode || '').toUpperCase() !== 'CUSTOMER') return false;
  if (cfg.ASYNC_REVIEW_ONLY !== false && !(reviewPackage && reviewPackage.files && reviewPackage.files.length)) return false;
  return true;
}

function archiveSentFilesOrEnqueueV94_(ctx, progress) {
  ctx = ctx || {};
  if (!shouldDeferSentFileArchiveAfterSendV94_(ctx.mode, ctx.reviewPackage)) {
    return new SentFileArchiveService(progress).archiveAfterSend(ctx);
  }
  return enqueueDeferredSentFileArchiveJobV94_(ctx);
}

function enqueueDeferredSentFileArchiveJobV94_(ctx) {
  const sheet = getOrCreateDeferredSentFileArchiveQueueSheetV94_();
  const jobId = 'archive_' + Utilities.getUuid();
  const job = buildDeferredSentFileArchiveJobV94_(jobId, ctx);
  const json = JSON.stringify(job);
  if (json.length > 48000) {
    // 큐 셀 한도 초과 위험이 있으면 안전하게 동기 저장으로 fallback합니다.
    const result = new SentFileArchiveService(null).archiveAfterSend(ctx);
    result.fallbackSync = true;
    result.fallbackReason = 'QUEUE_JSON_TOO_LARGE';
    return result;
  }

  const headers = getDeferredSentFileArchiveQueueHeadersV94_();
  sheet.appendRow([
    new Date(), '', '', jobId, 'PENDING', Number(ctx.rowNo || 0) || '', String(ctx.requestNo || ''), String(ctx.runId || ''),
    String(ctx.targetData && (ctx.targetData['고객번호'] || ctx.targetData['customerNo']) || ''),
    String(ctx.targetData && (ctx.targetData['회사명'] || ctx.targetData['고객사명'] || ctx.targetData['건물명']) || ''),
    json, '', ''
  ]);
  ensureDeferredSentFileArchiveTriggerV94_();
  return {
    ok: true,
    deferred: true,
    skipped: false,
    savedCount: 0,
    jobId: jobId,
    message: '발송파일 공유드라이브 저장은 백그라운드 큐로 예약됨'
  };
}

function buildDeferredSentFileArchiveJobV94_(jobId, ctx) {
  const mail = ctx.mail || {};
  return {
    version: 'v94',
    jobId: jobId,
    queuedAt: new Date().toISOString(),
    mode: String(ctx.mode || ''),
    source: String(ctx.source || ''),
    runId: String(ctx.runId || ''),
    rowNo: Number(ctx.rowNo || 0) || 0,
    requestNo: String(ctx.requestNo || ''),
    targetData: sanitizeJsonObjectV94_(ctx.targetData || {}),
    sender: sanitizeJsonObjectV94_(ctx.sender || {}),
    recipient: sanitizeJsonObjectV94_(ctx.recipient || {}),
    selectedDefs: sanitizeJsonObjectV94_(ctx.selectedDefs || []),
    mail: {
      from: String(mail.from || ''),
      to: Array.isArray(mail.to) ? mail.to.slice() : [],
      cc: Array.isArray(mail.cc) ? mail.cc.slice() : [],
      subject: String(mail.subject || '')
    },
    hiworksResult: sanitizeJsonObjectV94_(ctx.hiworksResult || {}),
    reviewPackage: sanitizeJsonObjectV94_(ctx.reviewPackage || null)
  };
}

function sanitizeJsonObjectV94_(value) {
  return JSON.parse(JSON.stringify(value == null ? null : value, function(key, val) {
    if (key === 'attachments') return undefined;
    if (val && typeof val.getBytes === 'function') return undefined;
    if (val instanceof Date) return val.toISOString();
    return val;
  }));
}

function getDeferredSentFileArchiveQueueHeadersV94_() {
  return ['등록일시', '시작일시', '완료일시', '작업ID', '상태', '행번호', '접수번호', 'runId', '고객번호', '회사명', '작업JSON', '결과JSON', '오류'];
}

function getOrCreateDeferredSentFileArchiveQueueSheetV94_() {
  const cfg = getSentFileArchiveConfig_();
  const ss = SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID || CONFIG.MASTER_SPREADSHEET_ID);
  const name = String(cfg.QUEUE_SHEET_NAME || '발송파일저장큐');
  let sheet = ss.getSheetByName(name);
  if (!sheet) {
    sheet = ss.insertSheet(name);
    try { sheet.hideSheet(); } catch (e) {}
  }
  const headers = getDeferredSentFileArchiveQueueHeadersV94_();
  const currentWidth = Math.max(sheet.getLastColumn(), headers.length);
  const current = sheet.getRange(1, 1, 1, currentWidth).getValues()[0].map(function(v) { return String(v || '').trim(); });
  let needsHeader = sheet.getLastRow() < 1 || current.filter(Boolean).length === 0;
  headers.forEach(function(h, idx) {
    if (current[idx] !== h) needsHeader = true;
  });
  if (needsHeader) {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function ensureDeferredSentFileArchiveTriggerV94_() {
  const cfg = getSentFileArchiveConfig_();
  const handler = String(cfg.ASYNC_TRIGGER_HANDLER || 'processDeferredSentFileArchiveQueueV94');
  const existing = ScriptApp.getProjectTriggers().some(function(t) {
    return t && t.getHandlerFunction && t.getHandlerFunction() === handler;
  });
  if (existing) return;
  ScriptApp.newTrigger(handler)
    .timeBased()
    .after(Number(cfg.ASYNC_TRIGGER_DELAY_MS || 60000) || 60000)
    .create();
}

function processDeferredSentFileArchiveQueueV94() {
  const cfg = getSentFileArchiveConfig_();
  const sheet = getOrCreateDeferredSentFileArchiveQueueSheetV94_();
  const headers = getDeferredSentFileArchiveQueueHeadersV94_();
  const idx = {};
  headers.forEach(function(h, i) { idx[h] = i + 1; });
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return { ok: true, processed: 0, message: '큐 비어 있음' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(3000)) return { ok: false, processed: 0, message: '큐 처리 lock 획득 실패' };
  let processed = 0;
  try {
    const max = Number(cfg.MAX_ASYNC_JOBS_PER_RUN || 3) || 3;
    const values = sheet.getRange(2, 1, lastRow - 1, headers.length).getValues();
    for (let r = 0; r < values.length && processed < max; r++) {
      const row = values[r];
      const status = String(row[idx['상태'] - 1] || '').trim();
      if (status !== 'PENDING') continue;
      const rowNo = r + 2;
      const jobJson = String(row[idx['작업JSON'] - 1] || '').trim();
      if (!jobJson) continue;

      sheet.getRange(rowNo, idx['상태']).setValue('RUNNING');
      sheet.getRange(rowNo, idx['시작일시']).setValue(new Date());
      SpreadsheetApp.flush();

      try {
        const job = JSON.parse(jobJson);
        const result = runDeferredSentFileArchiveJobV94_(job);
        sheet.getRange(rowNo, idx['상태']).setValue('DONE');
        sheet.getRange(rowNo, idx['완료일시']).setValue(new Date());
        sheet.getRange(rowNo, idx['결과JSON']).setValue(JSON.stringify(result).slice(0, 45000));
        processed++;
      } catch (err) {
        sheet.getRange(rowNo, idx['상태']).setValue('FAIL');
        sheet.getRange(rowNo, idx['완료일시']).setValue(new Date());
        sheet.getRange(rowNo, idx['오류']).setValue(String(err && err.stack || err).slice(0, 45000));
        processed++;
      }
    }
  } finally {
    try { lock.releaseLock(); } catch (e) {}
  }
  return { ok: true, processed: processed };
}

function runDeferredSentFileArchiveJobV94_(job) {
  job = job || {};
  const progress = new ProgressTracker(String(job.runId || Utilities.getUuid()) + '_archive');
  const targetData = applyContractPeriodDefaultsToObjectV90_(job.targetData || {});
  const reviewPackage = job.reviewPackage || null;
  if (!reviewPackage || !Array.isArray(reviewPackage.files) || !reviewPackage.files.length) {
    throw new Error('백그라운드 발송파일 저장은 현재 파일확인/수정 세션 자료만 처리합니다. reviewPackage가 없습니다.');
  }
  const selectedDefs = Array.isArray(job.selectedDefs) ? job.selectedDefs : [];
  if (!selectedDefs.length) throw new Error('백그라운드 발송파일 저장용 selectedDefs가 없습니다.');

  progress.start('백그라운드 발송파일 저장 시작', 5);
  const generatorSs = SpreadsheetApp.openById(CONFIG.GENERATOR_SPREADSHEET_ID);
  const attachments = new AttachmentBuilder(generatorSs, targetData, progress, reviewPackage).build(selectedDefs);
  const mailSnapshot = job.mail || {};
  const mail = new MailMessage({
    from: mailSnapshot.from || (job.sender && job.sender.email) || '',
    to: Array.isArray(mailSnapshot.to) ? mailSnapshot.to : [],
    cc: Array.isArray(mailSnapshot.cc) ? mailSnapshot.cc : [],
    subject: mailSnapshot.subject || '',
    bodyHtml: '',
    attachments: attachments
  });
  const result = new SentFileArchiveService(progress).archiveAfterSend({
    mode: job.mode,
    source: String(job.source || '') + '/백그라운드저장',
    runId: job.runId,
    rowNo: job.rowNo,
    requestNo: job.requestNo,
    targetData: targetData,
    sender: job.sender || {},
    recipient: job.recipient || {},
    selectedDefs: selectedDefs,
    mail: mail,
    attachments: attachments,
    hiworksResult: job.hiworksResult || {}
  });
  progress.done('백그라운드 발송파일 저장 완료');
  return result;
}

class SentFileArchiveService {
  constructor(progress) {
    this.progress = progress || null;
    this.cfg = getSentFileArchiveConfig_();
    this.ss = null;
    this.folderIndexCache_ = null;
  }

  archiveAfterSend(ctx) {
    ctx = ctx || {};

    if (!this.cfg.ENABLED) {
      return { ok: false, skipped: true, reason: 'DISABLED', message: '발송파일 저장 비활성화' };
    }

    const mode = String(ctx.mode || '').toUpperCase();
    if (mode === 'TEST' && !this.cfg.SAVE_TEST_MAIL) {
      return { ok: true, skipped: true, reason: 'TEST_MODE', message: 'TEST 발송은 저장하지 않음' };
    }

    const attachments = Array.isArray(ctx.attachments) ? ctx.attachments : [];
    if (!attachments.length) {
      return { ok: true, skipped: true, reason: 'NO_ATTACHMENTS', message: '저장할 실제 첨부파일 없음' };
    }

    const lock = LockService.getScriptLock();
    let locked = false;

    try {
      locked = lock.tryLock(12000);
      if (!locked) {
        return { ok: false, skipped: true, reason: 'LOCK_TIMEOUT', message: '발송파일 저장 lock 획득 실패' };
      }

      this.ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
      const logSheet = this.getOrCreateLogSheet_();
      const targetData = ctx.targetData || {};
      const folderInfo = this.resolveCustomerFolder_(targetData);
      const folder = DriveApp.getFolderById(folderInfo.folderId);
      const sentAt = new Date();
      const rows = [];
      const files = [];
      const selectedText = (ctx.selectedDefs || []).map(function(def) { return def && (def.label || def.key) || ''; }).filter(Boolean).join(', ');
      const hiworksText = JSON.stringify(ctx.hiworksResult || {}).slice(0, 500);

      attachments.forEach((blob, idx) => {
        try {
          const originalName = this.normalizeFileName_(blob && blob.getName ? blob.getName() : ('attachment_' + (idx + 1)));
          const savedName = this.buildUniqueSentFileName_(folder, originalName);
          const archiveBlob = blob.copyBlob ? blob.copyBlob() : blob;
          archiveBlob.setName(savedName);
          const file = sfaCreateFileInFolderFromBlob_(folderInfo.folderId, archiveBlob, savedName);
          const fileSize = this.safeBlobSize_(blob);
          const sha256 = this.cfg.CALCULATE_SHA256 ? this.sha256Blob_(blob) : '';
          const mimeType = blob.getContentType ? String(blob.getContentType() || '') : '';

          rows.push(this.buildLogRow_({
            status: '저장완료',
            message: '',
            sentAt: sentAt,
            ctx: ctx,
            targetData: targetData,
            folderInfo: folderInfo,
            selectedText: selectedText,
            originalName: originalName,
            savedName: savedName,
            file: file,
            fileSize: fileSize,
            mimeType: mimeType,
            sha256: sha256,
            hiworksText: hiworksText
          }));

          files.push({ name: savedName, id: file.getId(), url: file.getUrl(), size: fileSize, sha256: sha256 });
        } catch (fileErr) {
          rows.push(this.buildLogRow_({
            status: '저장실패',
            message: String(fileErr && fileErr.stack || fileErr),
            sentAt: sentAt,
            ctx: ctx,
            targetData: targetData,
            folderInfo: folderInfo,
            selectedText: selectedText,
            originalName: blob && blob.getName ? this.normalizeFileName_(blob.getName()) : ('attachment_' + (idx + 1)),
            savedName: '',
            file: null,
            fileSize: this.safeBlobSize_(blob),
            mimeType: blob && blob.getContentType ? String(blob.getContentType() || '') : '',
            sha256: '',
            hiworksText: hiworksText
          }));
        }
      });

      this.appendRows_(logSheet, rows);

      const failedRows = rows.filter(function(row) { return row && row['저장상태'] === '저장실패'; });
      return {
        ok: files.length > 0,
        skipped: false,
        savedCount: files.length,
        failedCount: failedRows.length,
        logRows: rows.length,
        folderId: folderInfo.folderId,
        folderUrl: folderInfo.folderUrl,
        folderName: folderInfo.folderName,
        files: files,
        message: files.length > 0
          ? ('발송파일 ' + files.length + '개 저장 완료' + (failedRows.length ? ' / 일부 실패 ' + failedRows.length + '건' : ''))
          : ('발송파일 저장 실패: ' + (failedRows[0] ? failedRows[0]['저장메시지'] : '저장된 파일 없음'))
      };
    } catch (err) {
      try {
        this.appendArchiveErrorLog_(ctx, err);
      } catch (logErr) {}
      return { ok: false, skipped: false, error: String(err && err.stack || err), message: String(err && err.message || err) };
    } finally {
      if (locked) {
        try { lock.releaseLock(); } catch (releaseErr) {}
      }
    }
  }

  appendArchiveErrorLog_(ctx, err) {
    this.ss = this.ss || SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
    const logSheet = this.getOrCreateLogSheet_();
    const targetData = ctx && ctx.targetData || {};
    const folderInfo = { folderId: '', folderUrl: '', folderName: '' };
    const row = this.buildLogRow_({
      status: '저장실패',
      message: String(err && err.stack || err),
      sentAt: new Date(),
      ctx: ctx || {},
      targetData: targetData,
      folderInfo: folderInfo,
      selectedText: (ctx && ctx.selectedDefs || []).map(function(def) { return def && (def.label || def.key) || ''; }).filter(Boolean).join(', '),
      originalName: '',
      savedName: '',
      file: null,
      fileSize: '',
      mimeType: '',
      sha256: '',
      hiworksText: JSON.stringify(ctx && ctx.hiworksResult || {}).slice(0, 500)
    });
    this.appendRows_(logSheet, [row]);
  }

  getOrCreateLogSheet_() {
    const headers = this.getLogHeaders_();
    let sheet = this.ss.getSheetByName(this.cfg.LOG_SHEET_NAME);
    if (!sheet) sheet = this.ss.insertSheet(this.cfg.LOG_SHEET_NAME);
    this.ensureHeaders_(sheet, headers);
    return sheet;
  }

  getLogHeaders_() {
    return [
      '기록일시', '발송일시', 'runId', '접수번호', '고객번호', '회사명', '수행사', '영업담당자',
      '발신자', '수신자', '참조', '메일제목', '선택자료',
      '파일명_원본', '파일명_저장', '파일ID', '파일URL', '파일크기', 'MIME타입', 'SHA256',
      '자료구분', '저장상태', '저장메시지', '하이웍스응답요약', '포털/시트발송구분',
      '고객폴더ID', '고객폴더URL', '고객폴더명',
      '폴더이력반영상태', '폴더이력반영일시', '폴더이력메시지'
    ];
  }

  getFolderIndexHeaders_() {
    return ['고객번호', '회사명', '수행사', '폴더명', '폴더ID', '폴더URL', '최종확인일시', '비고'];
  }

  ensureHeaders_(sheet, headers) {
    const width = Math.max(sheet.getLastColumn(), headers.length);
    const current = sheet.getRange(1, 1, 1, width).getValues()[0].map(function(v) { return String(v || '').trim(); });
    const existingSet = {};
    current.forEach(function(h) { if (h) existingSet[h] = true; });

    let changed = false;
    headers.forEach(function(h) {
      if (!existingSet[h]) {
        current.push(h);
        existingSet[h] = true;
        changed = true;
      }
    });

    if (sheet.getLastRow() === 0 || current.filter(Boolean).length === 0) changed = true;

    if (changed) {
      sheet.getRange(1, 1, 1, current.length).setValues([current]);
      sheet.setFrozenRows(1);
    }
  }

  appendRows_(sheet, rows) {
    if (!rows || !rows.length) return;
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v) { return String(v || '').trim(); });
    const values = rows.map(function(rowObj) {
      return headers.map(function(h) { return rowObj[h] != null ? rowObj[h] : ''; });
    });
    const startRow = Math.max(sheet.getLastRow() + 1, 2);
    sheet.getRange(startRow, 1, values.length, headers.length).setValues(values);
  }

  buildLogRow_(p) {
    const ctx = p.ctx || {};
    const targetData = p.targetData || {};
    const folderInfo = p.folderInfo || {};
    const sender = ctx.sender || {};
    const recipient = ctx.recipient || {};
    const mail = ctx.mail || {};
    return {
      '기록일시': new Date(),
      '발송일시': p.sentAt || new Date(),
      'runId': ctx.runId || '',
      '접수번호': ctx.requestNo || '',
      '고객번호': this.getCustomerNo_(targetData),
      '회사명': this.getValue_(targetData, ['회사명', '건물명', '고객사명']),
      '수행사': this.getVendorValue_(targetData),
      '영업담당자': this.getValue_(targetData, ['영업담당자', '견적담당', '담당자']),
      '발신자': sender.email || mail.from || '',
      '수신자': (recipient.to || mail.to || []).join ? (recipient.to || mail.to || []).join(',') : String(recipient.to || mail.to || ''),
      '참조': (recipient.cc || mail.cc || []).join ? (recipient.cc || mail.cc || []).join(',') : String(recipient.cc || mail.cc || ''),
      '메일제목': mail.subject || '',
      '선택자료': p.selectedText || '',
      '파일명_원본': p.originalName || '',
      '파일명_저장': p.savedName || '',
      '파일ID': p.file ? p.file.getId() : '',
      '파일URL': p.file ? p.file.getUrl() : '',
      '파일크기': p.fileSize || '',
      'MIME타입': p.mimeType || '',
      'SHA256': p.sha256 || '',
      '자료구분': '발송',
      '저장상태': p.status || '',
      '저장메시지': String(p.message || '').slice(0, 1500),
      '하이웍스응답요약': p.hiworksText || '',
      '포털/시트발송구분': ctx.source || '',
      '고객폴더ID': folderInfo.folderId || '',
      '고객폴더URL': folderInfo.folderUrl || '',
      '고객폴더명': folderInfo.folderName || '',
      '폴더이력반영상태': '',
      '폴더이력반영일시': '',
      '폴더이력메시지': ''
    };
  }

  resolveCustomerFolder_(targetData) {
    const customerNo = this.getCustomerNo_(targetData);
    const company = this.getValue_(targetData, ['회사명', '건물명', '고객사명']);
    const vendor = this.getVendorValue_(targetData);

    if (!customerNo && !company) {
      throw new Error('고객번호/회사명을 찾지 못해 고객사 폴더를 결정할 수 없습니다.');
    }

    const driveId = sfaGetSharedDriveId_(this.cfg);

    // 1순위: 공유드라이브에서 고객번호_ prefix 폴더명을 직접 검색합니다.
    // BJ 고객사폴더ID가 사람이 잘못 수정되어도 이 경로가 우선이라 오저장 위험이 낮습니다.
    const foundByName = this.findFolderBySharedDriveName_(driveId, customerNo, company, vendor);
    if (foundByName) {
      this.upsertFolderIndex_(customerNo, company, vendor, foundByName.folderName, foundByName.folderId, foundByName.folderUrl, foundByName.note || '공유드라이브명/고객번호 prefix 매칭');
      return foundByName;
    }

    // 2순위: 마스터시트 BJ열 고객사폴더ID를 보조로 사용합니다.
    // 단, 폴더명이 고객번호_로 시작하는 경우에만 신뢰합니다.
    const fromMasterId = this.findFolderFromMasterFolderId_(targetData, customerNo);
    if (fromMasterId) {
      this.upsertFolderIndex_(customerNo, company, vendor, fromMasterId.folderName, fromMasterId.folderId, fromMasterId.folderUrl, '마스터 고객사폴더ID 보조매칭');
      return fromMasterId;
    }

    // 3순위: 발송파일용 내부 인덱스. 이것도 고객번호 prefix 검증 후 사용합니다.
    const indexed = this.findFolderFromIndex_(customerNo, company);
    if (indexed && this.isFolderNameSafeForCustomer_(indexed.folderName, customerNo)) {
      return indexed;
    }

    if (!this.cfg.CREATE_FOLDER_IF_MISSING) {
      throw new Error('고객사 폴더를 찾지 못했습니다: ' + customerNo + ' / ' + company);
    }

    // 4순위: 없으면 공유드라이브 루트에 신규 생성.
    const newFolderName = this.buildCustomerFolderName_(customerNo, company, vendor);
    const newFolder = sfaCreateDriveFolder_(newFolderName, driveId);
    const created = {
      folderId: newFolder.id,
      folderUrl: newFolder.webViewLink || ('https://drive.google.com/drive/folders/' + newFolder.id),
      folderName: newFolder.name || newFolderName,
      note: '공유드라이브 루트 신규 생성'
    };
    this.upsertFolderIndex_(customerNo, company, vendor, created.folderName, created.folderId, created.folderUrl, created.note);
    return created;
  }

  findFolderBySharedDriveName_(driveId, customerNo, company, vendor) {
    const expectedFolderName = this.buildCustomerFolderName_(customerNo, company, vendor);

    // 1) 루트에서 정확한 이름 우선.
    const exactRoot = sfaFindChildFolder_(driveId, driveId, expectedFolderName);
    if (exactRoot) return this.rawFolderToInfo_(exactRoot, '공유드라이브 루트 정확한 폴더명 매칭');

    // 2) 루트에서 고객번호 prefix 검색.
    const prefixRoot = sfaFindCustomerFolderByPrefixInParent_(driveId, driveId, customerNo);
    if (prefixRoot) return this.rawFolderToInfo_(prefixRoot, '공유드라이브 루트 고객번호 prefix 매칭');

    // 3) 수주실패 폴더 안 검색.
    const failedParentName = '수주실패';
    const failedParent = sfaFindChildFolder_(driveId, driveId, failedParentName);
    if (failedParent && failedParent.id) {
      const exactFailed = sfaFindChildFolder_(failedParent.id, driveId, expectedFolderName);
      if (exactFailed) return this.rawFolderToInfo_(exactFailed, '수주실패 폴더 정확한 폴더명 매칭');

      const prefixFailed = sfaFindCustomerFolderByPrefixInParent_(driveId, failedParent.id, customerNo);
      if (prefixFailed) return this.rawFolderToInfo_(prefixFailed, '수주실패 폴더 고객번호 prefix 매칭');
    }

    // v92 핵심:
    // 고객사 폴더가 공유드라이브 루트 바로 아래가 아니라 수행사/계약상태/연도 폴더 밑에 있어도 찾습니다.
    // 기존 v91은 루트와 수주실패 1단계만 검색해서, 전산 자동발송 파일이 실제 고객사별 폴더를 못 찾고 저장을 건너뛰거나 루트 신규폴더로 빠질 수 있었습니다.
    const recursive = sfaFindCustomerFolderAnywhereInDrive_(driveId, customerNo, company);
    if (recursive) return this.rawFolderToInfo_(recursive, recursive.note || '공유드라이브 전체 고객폴더 매칭');

    return null;
  }

  rawFolderToInfo_(folder, note) {
    return {
      folderId: folder.id,
      folderUrl: folder.webViewLink || ('https://drive.google.com/drive/folders/' + folder.id),
      folderName: folder.name || '',
      note: note || ''
    };
  }

  findFolderFromMasterFolderId_(targetData, customerNo) {
    const headers = this.cfg.MASTER_FOLDER_ID_HEADERS || ['고객사폴더ID'];
    const folderId = this.getValue_(targetData, headers);
    if (!folderId) return null;

    try {
      const folder = DriveApp.getFolderById(String(folderId).trim());
      const folderName = folder.getName();
      if (!this.isFolderNameSafeForCustomer_(folderName, customerNo)) {
        return null;
      }
      return {
        folderId: folder.getId(),
        folderUrl: folder.getUrl(),
        folderName: folderName
      };
    } catch (err) {
      return null;
    }
  }

  isFolderNameSafeForCustomer_(folderName, customerNo) {
    return sfaIsFolderNameSafeForCustomer_(folderName, customerNo);
  }

  getFolderIndexSheet_() {
    const headers = this.getFolderIndexHeaders_();
    let sheet = this.ss.getSheetByName(this.cfg.FOLDER_INDEX_SHEET_NAME);
    if (!sheet) sheet = this.ss.insertSheet(this.cfg.FOLDER_INDEX_SHEET_NAME);
    this.ensureHeaders_(sheet, headers);
    return sheet;
  }

  loadFolderIndex_() {
    if (this.folderIndexCache_) return this.folderIndexCache_;
    const sheet = this.getFolderIndexSheet_();
    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    const out = [];
    if (lastRow >= 2) {
      const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || '').trim(); });
      const values = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      values.forEach(function(row, idx) {
        const obj = { rowNo: idx + 2 };
        headers.forEach(function(h, i) { obj[h] = row[i]; });
        out.push(obj);
      });
    }
    this.folderIndexCache_ = out;
    return out;
  }

  findFolderFromIndex_(customerNo, company) {
    const rows = this.loadFolderIndex_();
    const no = this.normalizeCustomerNo_(customerNo);
    const companyKey = this.normalizeText_(company);

    let hit = null;
    if (no) {
      hit = rows.find(row => this.normalizeCustomerNo_(row['고객번호']) === no && row['폴더ID']);
    }
    if (!hit && companyKey) {
      hit = rows.find(row => this.normalizeText_(row['회사명']) === companyKey && row['폴더ID']);
    }
    if (!hit) return null;

    try {
      const folder = DriveApp.getFolderById(String(hit['폴더ID'] || '').trim());
      return { folderId: folder.getId(), folderUrl: folder.getUrl(), folderName: folder.getName() };
    } catch (err) {
      return null;
    }
  }

  findFolderInRoot_(root, customerNo, company) {
    const no = this.normalizeCustomerNo_(customerNo);
    const companyKey = this.normalizeText_(company);
    const candidates = [];
    const folders = root.getFolders();

    while (folders.hasNext()) {
      const folder = folders.next();
      const name = String(folder.getName() || '');
      const nameKey = this.normalizeText_(name);
      let score = 0;

      if (no && name.indexOf(no + '_') === 0) score += 100;
      if (companyKey && nameKey.indexOf(companyKey) >= 0) score += 50;
      if (score > 0) {
        candidates.push({ score: score, folder: folder, name: name });
      }
    }

    if (!candidates.length) return null;
    candidates.sort(function(a, b) {
      if (b.score !== a.score) return b.score - a.score;
      return a.name.length - b.name.length;
    });

    const picked = candidates[0].folder;
    return { folderId: picked.getId(), folderUrl: picked.getUrl(), folderName: picked.getName() };
  }

  upsertFolderIndex_(customerNo, company, vendor, folderName, folderId, folderUrl, note) {
    const sheet = this.getFolderIndexSheet_();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v) { return String(v || '').trim(); });
    const no = this.normalizeCustomerNo_(customerNo);
    const rows = this.loadFolderIndex_();
    let rowNo = 0;

    if (no) {
      const hit = rows.find(row => this.normalizeCustomerNo_(row['고객번호']) === no);
      if (hit) rowNo = hit.rowNo;
    }

    const obj = {
      '고객번호': no,
      '회사명': company || '',
      '수행사': vendor || '',
      '폴더명': folderName || '',
      '폴더ID': folderId || '',
      '폴더URL': folderUrl || '',
      '최종확인일시': new Date(),
      '비고': note || ''
    };

    const values = headers.map(function(h) { return obj[h] != null ? obj[h] : ''; });
    if (rowNo) sheet.getRange(rowNo, 1, 1, headers.length).setValues([values]);
    else sheet.appendRow(values);

    this.folderIndexCache_ = null;
  }

  buildCustomerFolderName_(customerNo, company, vendor) {
    const parts = [];
    const no = this.normalizeCustomerNo_(customerNo);
    if (no) parts.push(no);
    parts.push(this.cleanFileName_(company || '회사명없음'));
    const vendorCode = this.getVendorFolderCode_(vendor);
    if (vendorCode) parts.push(vendorCode);
    return parts.join('_');
  }

  buildUniqueSentFileName_(folder, originalName) {
    const prefix = this.cfg.SENT_PREFIX || '[발송]';
    let baseName = this.normalizeFileName_(originalName);
    if (baseName.indexOf(prefix) !== 0) baseName = prefix + baseName;

    const folderId = folder && folder.getId ? folder.getId() : '';
    if (!sfaFolderHasFileName_(folderId, baseName)) return baseName;

    const dot = baseName.lastIndexOf('.');
    const stem = dot > 0 ? baseName.slice(0, dot) : baseName;
    const ext = dot > 0 ? baseName.slice(dot) : '';

    for (let i = 1; i < 10000; i++) {
      const suffix = '_' + ('000' + i).slice(-3);
      const candidate = stem + suffix + ext;
      if (!sfaFolderHasFileName_(folderId, candidate)) return candidate;
    }

    throw new Error('파일명 충돌이 너무 많습니다: ' + baseName);
  }

  syncFolderHistoryFromCentralLog_() {
    const cfg = this.cfg;
    if (!cfg.DAILY_HISTORY_SYNC_ENABLED) {
      return { ok: true, skipped: true, message: '일일 이력 반영 비활성화' };
    }

    const lock = LockService.getScriptLock();
    let locked = false;
    try {
      locked = lock.tryLock(12000);
      if (!locked) throw new Error('일일 이력 반영 lock 획득 실패');

      this.ss = SpreadsheetApp.openById(CONFIG.MASTER_SPREADSHEET_ID);
      const logSheet = this.getOrCreateLogSheet_();
      const lastRow = logSheet.getLastRow();
      const lastCol = logSheet.getLastColumn();
      if (lastRow < 2) return { ok: true, processed: 0, message: '발송파일로그 데이터 없음' };

      const headers = logSheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function(v) { return String(v || '').trim(); });
      const idx = this.indexHeaders_(headers);
      const values = logSheet.getRange(2, 1, lastRow - 1, lastCol).getValues();
      const limit = Number(cfg.MAX_HISTORY_SYNC_ROWS_PER_RUN) || 300;
      let processed = 0;
      let errorCount = 0;

      for (let i = 0; i < values.length; i++) {
        if (processed >= limit) break;
        const row = values[i];
        const rowNo = i + 2;
        const saveStatus = this.getByHeaderIndex_(row, idx, '저장상태');
        const historyStatus = this.getByHeaderIndex_(row, idx, '폴더이력반영상태');
        const folderId = this.getByHeaderIndex_(row, idx, '고객폴더ID');

        if (saveStatus !== '저장완료') continue;
        if (historyStatus === '완료') continue;
        if (!folderId) continue;

        try {
          this.appendOneFolderHistory_(folderId, headers, row);
          this.setLogCellByHeader_(logSheet, rowNo, idx, '폴더이력반영상태', '완료');
          this.setLogCellByHeader_(logSheet, rowNo, idx, '폴더이력반영일시', new Date());
          this.setLogCellByHeader_(logSheet, rowNo, idx, '폴더이력메시지', '일일반영 완료');
          processed++;
        } catch (err) {
          errorCount++;
          this.setLogCellByHeader_(logSheet, rowNo, idx, '폴더이력반영상태', '오류');
          this.setLogCellByHeader_(logSheet, rowNo, idx, '폴더이력반영일시', new Date());
          this.setLogCellByHeader_(logSheet, rowNo, idx, '폴더이력메시지', String(err && err.message || err).slice(0, 500));
        }
      }

      return { ok: true, processed: processed, errors: errorCount, limit: limit };
    } finally {
      if (locked) {
        try { lock.releaseLock(); } catch (releaseErr) {}
      }
    }
  }

  appendOneFolderHistory_(folderId, logHeaders, logRow) {
    const folder = DriveApp.getFolderById(String(folderId));
    const historySs = this.getOrCreateFolderHistorySpreadsheet_(folder);
    const sheet = historySs.getSheets()[0];
    const historyHeaders = [
      '반영일시', '발송일시', '접수번호', '고객번호', '회사명', '수행사', '영업담당자',
      '발신자', '수신자', '참조', '메일제목', '선택자료',
      '파일명_저장', '파일URL', '파일크기', '자료구분', 'runId'
    ];
    this.ensureHeaders_(sheet, historyHeaders);

    const logIdx = this.indexHeaders_(logHeaders);
    const obj = {
      '반영일시': new Date(),
      '발송일시': this.getByHeaderIndex_(logRow, logIdx, '발송일시'),
      '접수번호': this.getByHeaderIndex_(logRow, logIdx, '접수번호'),
      '고객번호': this.getByHeaderIndex_(logRow, logIdx, '고객번호'),
      '회사명': this.getByHeaderIndex_(logRow, logIdx, '회사명'),
      '수행사': this.getByHeaderIndex_(logRow, logIdx, '수행사'),
      '영업담당자': this.getByHeaderIndex_(logRow, logIdx, '영업담당자'),
      '발신자': this.getByHeaderIndex_(logRow, logIdx, '발신자'),
      '수신자': this.getByHeaderIndex_(logRow, logIdx, '수신자'),
      '참조': this.getByHeaderIndex_(logRow, logIdx, '참조'),
      '메일제목': this.getByHeaderIndex_(logRow, logIdx, '메일제목'),
      '선택자료': this.getByHeaderIndex_(logRow, logIdx, '선택자료'),
      '파일명_저장': this.getByHeaderIndex_(logRow, logIdx, '파일명_저장'),
      '파일URL': this.getByHeaderIndex_(logRow, logIdx, '파일URL'),
      '파일크기': this.getByHeaderIndex_(logRow, logIdx, '파일크기'),
      '자료구분': this.getByHeaderIndex_(logRow, logIdx, '자료구분'),
      'runId': this.getByHeaderIndex_(logRow, logIdx, 'runId')
    };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(function(v) { return String(v || '').trim(); });
    sheet.appendRow(headers.map(function(h) { return obj[h] != null ? obj[h] : ''; }));
  }

  getOrCreateFolderHistorySpreadsheet_(folder) {
    const name = this.cfg.FOLDER_HISTORY_SHEET_NAME || '_메일이력_발송';
    const files = folder.getFilesByName(name);
    while (files.hasNext()) {
      const file = files.next();
      if (file.getMimeType && file.getMimeType() === MimeType.GOOGLE_SHEETS) {
        return SpreadsheetApp.openById(file.getId());
      }
    }

    const ss = SpreadsheetApp.create(name);
    const file = DriveApp.getFileById(ss.getId());
    folder.addFile(file);
    try { DriveApp.getRootFolder().removeFile(file); } catch (err) {}
    ss.getSheets()[0].setName('발송이력');
    return ss;
  }

  indexHeaders_(headers) {
    const idx = {};
    headers.forEach(function(h, i) { if (h) idx[h] = i; });
    return idx;
  }

  getByHeaderIndex_(row, idx, header) {
    const i = idx[header];
    if (i == null || i < 0) return '';
    return row[i] != null ? row[i] : '';
  }

  setLogCellByHeader_(sheet, rowNo, idx, header, value) {
    const i = idx[header];
    if (i == null || i < 0) return;
    sheet.getRange(rowNo, i + 1).setValue(value);
  }

  normalizeFileName_(name) {
    let s = String(name || 'attachment').trim() || 'attachment';
    s = s.replace(/[\\/:*?"<>|]/g, '_').replace(/\s+/g, ' ').trim();
    if (s.length > 180) {
      const dot = s.lastIndexOf('.');
      const ext = dot > 0 ? s.slice(dot) : '';
      const stem = dot > 0 ? s.slice(0, dot) : s;
      s = stem.slice(0, Math.max(20, 180 - ext.length)) + ext;
    }
    return s;
  }

  cleanFileName_(value) {
    return this.normalizeFileName_(String(value || '').replace(/\s+/g, ' ').trim() || '값없음');
  }

  safeBlobSize_(blob) {
    try { return blob.getBytes().length; } catch (err) { return ''; }
  }

  sha256Blob_(blob) {
    const bytes = blob.getBytes();
    const digest = Utilities.computeDigest(Utilities.DigestAlgorithm.SHA_256, bytes);
    return digest.map(function(b) {
      const v = (b < 0 ? b + 256 : b).toString(16);
      return v.length === 1 ? '0' + v : v;
    }).join('');
  }

  getValue_(obj, keys) {
    for (let i = 0; i < keys.length; i++) {
      const key = keys[i];
      const value = obj[key];
      if (value != null && String(value).trim() !== '') return String(value).trim();
    }
    return '';
  }

  getCustomerNo_(obj) {
    return this.normalizeCustomerNo_(this.getValue_(obj, ['고객번호', '고객 No', '고객NO', '고객No', '고객 no', '고객ID', '고객 id']));
  }

  getVendorValue_(obj) {
    return this.getValue_(obj, ['최종수행사', '수행사', '수행사명', '계약수행사']);
  }

  normalizeCustomerNo_(value) {
    const s = String(value == null ? '' : value).trim();
    if (!s) return '';
    const numeric = Number(s);
    if (Number.isFinite(numeric) && Math.floor(numeric) === numeric) return String(numeric);
    return s.replace(/\.0$/, '');
  }

  normalizeText_(value) {
    return String(value || '')
      .toLowerCase()
      .replace(/주식회사|\(주\)|㈜|유한회사|재단법인|사단법인/g, '')
      .replace(/[\s_\-()\[\]{}.,·ㆍ]/g, '')
      .trim();
  }

  getVendorFolderCode_(vendor) {
    const normalized = (typeof normalizeVendorName_ === 'function') ? normalizeVendorName_(vendor) : String(vendor || '').trim();
    if (normalized === '케이제이') return 'KJ';
    if (normalized === '일신') return '일신';
    if (normalized === '삼구') return '삼구';
    if (normalized === '디엠') return '디엠';
    return this.cleanFileName_(vendor || '');
  }
}

