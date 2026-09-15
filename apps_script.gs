/* ════════════════════════════════════════════════════════════
   오헤이오 캠페인 신청서 수집 스크립트
   - 캠페인(월)별로 기록할 탭을 자동 분리
   - 타입/고료 구분은 컬럼 + 헤더 필터로 (행 색상 없음)
   ════════════════════════════════════════════════════════════ */

/* ── 1. 대상 스프레드시트 ──────────────────────────────────
   이 스크립트가 대상 시트에 바인딩(시트 > 확장 프로그램 > Apps Script)돼 있으면
   '' 로 비워두면 됩니다. 아래처럼 ID를 넣으면 그 파일에 강제로 기록합니다. */
const SPREADSHEET_ID = '18TXjNFsO4GjW5ye_OiTAMYR-LHca-19GG2CaqnX3Bc0';

/* ── 2. 캠페인별 탭 분리 규칙 ──────────────────────────────
   가이드라인이 보내는 type 값(예: 'A Type 9월')에 match 문자열이 있으면
   해당 탭에 기록합니다. 위에서부터 순서대로 검사합니다.
   → 10월 캠페인이 생기면 { match: '10월', sheetName: '[응답] 10월 OOO' } 한 줄만 추가 */
const SHEET_ROUTES = [
  { match: '9월', sheetName: '[응답] 9월 블러쉬' }
];

/* 위 규칙에 걸리지 않는 신청(= 8월 하이라이터 등)이 들어갈 기본 탭 */
const DEFAULT_SHEET_NAME = '시트1';

/* ── 3. 컬럼 구성 ──────────────────────────────────────────
   순서를 바꾸면 아래 PHONE_COL / ZIP_COL 번호도 같이 바꿔야 합니다. */
const HEADERS = [
  '제출일시', '타입', '고료', '이름', '인스타그램', '휴대폰',
  '이메일', '제공 제품 조합', '우편번호', '배송지 주소', '요청사항'
];
const PHONE_COL = 6;   // 휴대폰
const ZIP_COL   = 9;   // 우편번호


/* ════════════════════════════════════════════════════════════
   신청서 수신
   ════════════════════════════════════════════════════════════ */
function doPost(e) {
  // 동시에 여러 명이 제출해도 행이 겹치지 않도록 잠금
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);

  try {
    const data = JSON.parse(e.postData.contents);
    const type = String(data.type || '');

    const route = resolveSheet_(type);
    const sheet = getOrCreateSheet_(route.name, route.routed);

    const values = [
      new Date(),
      type,
      data.fee || '',
      data.name || '',
      data.instagram || '',
      data.phone || '',
      data.email || '',
      data.shades || '',
      data.zipcode || '',
      data.address || '',
      data.note || ''
    ];

    const row = sheet.getLastRow() + 1;

    // 앞자리 0이 날아가지 않도록 숫자처럼 보이는 칸을 먼저 텍스트 서식으로 고정
    sheet.getRange(row, PHONE_COL).setNumberFormat('@');
    sheet.getRange(row, ZIP_COL).setNumberFormat('@');

    const range = sheet.getRange(row, 1, 1, values.length);
    range.setValues([values]);

    // 타입/고료 구분은 색이 아니라 '타입'·'고료' 컬럼과 필터로 합니다.
    // 행 배경은 전부 흰색(채우기 없음)으로 고정.
    range.setBackground(null);
    range.setVerticalAlignment('middle');

    return json_({ result: 'success', sheet: sheet.getName(), row: row });

  } catch (err) {
    // 실패해도 신청 내용이 사라지지 않도록 로그에 원문을 남김
    console.error('제출 처리 실패: ' + err + ' | payload: ' +
      (e && e.postData ? e.postData.contents : '(없음)'));
    return json_({ result: 'error', message: String(err) });

  } finally {
    lock.releaseLock();
  }
}

function doGet() {
  return json_({ status: 'ok' });
}


/* ════════════════════════════════════════════════════════════
   내부 함수
   ════════════════════════════════════════════════════════════ */

/**
 * type 값에 맞는 탭을 찾는다.
 * routed=true 는 이 코드가 만든 캠페인 탭(위 SHEET_ROUTES)이라는 뜻으로,
 * 이 탭만 헤더를 자동으로 맞춥니다. 기본 탭(8월 등)은 컬럼 구성이 달라 건드리지 않습니다.
 */
function resolveSheet_(type) {
  for (let i = 0; i < SHEET_ROUTES.length; i++) {
    if (type.indexOf(SHEET_ROUTES[i].match) !== -1) {
      return { name: SHEET_ROUTES[i].sheetName, routed: true };
    }
  }
  return { name: DEFAULT_SHEET_NAME, routed: false };
}

function getSpreadsheet_() {
  return SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();
}

/**
 * 탭이 없으면 만들고, 헤더가 없으면 헤더를 세팅한다.
 * syncHeader=true 면 이미 있는 탭도 1행 제목이 코드와 다를 때 자동으로 맞춥니다.
 * (컬럼명을 바꿔도 다음 신청이 들어오는 순간 기존 탭에 반영됨. 데이터 행은 건드리지 않음)
 */
function getOrCreateSheet_(name, syncHeader) {
  const ss = getSpreadsheet_();
  let sheet = ss.getSheetByName(name);

  if (!sheet) {
    sheet = ss.insertSheet(name);
  }

  if (syncHeader && sheet.getLastRow() > 0) {
    const head = sheet.getRange(1, 1, 1, HEADERS.length);
    if (head.getValues()[0].join('||') !== HEADERS.join('||')) {
      head.setValues([HEADERS]);
    }
  }

  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1, 1, HEADERS.length)
      .setValues([HEADERS])
      .setFontWeight('bold')
      .setBackground('#3E2C22')
      .setFontColor('#FFFFFF')
      .setVerticalAlignment('middle');

    sheet.setFrozenRows(1);
    sheet.setColumnWidth(1, 150);   // 제출일시
    sheet.setColumnWidth(2, 100);   // 타입
    sheet.setColumnWidth(3, 90);    // 고료
    sheet.setColumnWidth(5, 230);   // 인스타그램
    sheet.setColumnWidth(8, 220);   // 제공 제품 조합
    sheet.setColumnWidth(10, 300);  // 배송지 주소
    sheet.setColumnWidth(11, 220);  // 요청사항

    if (!sheet.getFilter()) {
      sheet.getRange(1, 1, 1, HEADERS.length).createFilter();
    }
  }

  return sheet;
}

function json_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}


/* ════════════════════════════════════════════════════════════
   시트에서 직접 쓰는 메뉴 (스크립트가 시트에 바인딩된 경우에만 표시)
   ════════════════════════════════════════════════════════════ */
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('오헤이오')
    .addItem('현재 탭 타입별로 정렬', 'sortActiveSheetByType')
    .addItem('현재 탭 행 색상 모두 지우기', 'clearRowColors')
    .addItem('현재 탭 헤더 이름 맞추기', 'syncActiveSheetHeader')
    .addToUi();
}

/** 현재 탭을 타입 → 제출일시 순으로 정렬 */
function sortActiveSheetByType() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const last = sheet.getLastRow();
  if (last < 3) return;

  sheet.getRange(2, 1, last - 1, HEADERS.length).sort([
    { column: 2, ascending: true },   // 타입
    { column: 1, ascending: true }    // 제출일시
  ]);
}

/**
 * 현재 탭 1행의 제목만 위 HEADERS로 덮어씁니다. 데이터 행은 건드리지 않습니다.
 * 컬럼명을 바꿨을 때(예: '제품 구성' → '제공 제품 조합') 기존 탭에 반영하는 용도.
 * 8월 탭처럼 컬럼 구성이 다른 시트에서는 실행하지 마세요.
 */
function syncActiveSheetHeader() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const ui = SpreadsheetApp.getUi();

  const answer = ui.alert(
    '헤더 이름 맞추기',
    '"' + sheet.getName() + '" 탭의 1행 제목을 현재 코드 기준으로 덮어씁니다.\n' +
    '데이터는 그대로입니다. 계속할까요?',
    ui.ButtonSet.OK_CANCEL);
  if (answer !== ui.Button.OK) return;

  sheet.getRange(1, 1, 1, HEADERS.length).setValues([HEADERS]);
}

/** 이미 색이 칠해진 기존 행까지 전부 흰색으로 되돌리기 */
function clearRowColors() {
  const sheet = SpreadsheetApp.getActiveSheet();
  const last = sheet.getLastRow();
  if (last < 2) return;

  sheet.getRange(2, 1, last - 1, HEADERS.length).setBackground(null);
}


/* ════════════════════════════════════════════════════════════
   테스트용 — 편집기에서 실행하면 9월 탭에 A~G 더미 행 7개가 들어갑니다.
   확인 후 해당 행들은 지우세요.
   ════════════════════════════════════════════════════════════ */
function 테스트_더미행_넣기() {
  const cases = [
    ['A Type 9월', '50,000원'],
    ['B Type 9월', '100,000원'],
    ['C Type 9월', '150,000원'],
    ['D Type 9월', '200,000원'],
    ['E Type 9월', '250,000원'],
    ['F Type 9월', '300,000원'],
    ['G Type 9월', '고료 조정']
  ];

  cases.forEach(function (c, i) {
    doPost({
      postData: {
        contents: JSON.stringify({
          type: c[0],
          fee: c[1],
          name: '테스트' + (i + 1),
          instagram: 'https://www.instagram.com/test' + (i + 1),
          phone: '010-0000-000' + (i + 1),
          email: 'test@example.com',
          shades: i % 2 === 0
            ? '[A조합] 블러셔 06. 라떼핀드 + 하이라이터 01. 로우시에나'
            : '[B조합] 블러셔 07. 로우키인러브 + 하이라이터 02. 로즈헤이즈',
          zipcode: '06234',
          address: '서울시 강남구 테헤란로 1 101동 101호',
          note: '테스트 행입니다'
        })
      }
    });
  });
}
