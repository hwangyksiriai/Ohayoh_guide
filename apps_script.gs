function doPost(e) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('시트1')
    || SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(['제출일시', '타입', '고료', '이름', '인스타그램', '휴대폰', '이메일', '색상 선택', '우편번호', '배송지 주소', '요청사항']);
  }

  const data = JSON.parse(e.postData.contents);

  const nextRow = sheet.getLastRow() + 1;
  // 숫자로 보이는 열(휴대폰, 우편번호)을 텍스트 서식으로 고정 (앞자리 0 소실 방지)
  sheet.getRange(nextRow, 6).setNumberFormat('@');
  sheet.getRange(nextRow, 9).setNumberFormat('@');

  sheet.appendRow([
    new Date(),
    data.type || '',
    data.fee || '',
    data.name || '',
    data.instagram || '',
    data.phone || '',
    data.email || '',
    data.shades || '',
    data.zipcode || '',
    data.address || '',
    data.note || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ result: 'success' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({ status: 'ok' }))
    .setMimeType(ContentService.MimeType.JSON);
}
