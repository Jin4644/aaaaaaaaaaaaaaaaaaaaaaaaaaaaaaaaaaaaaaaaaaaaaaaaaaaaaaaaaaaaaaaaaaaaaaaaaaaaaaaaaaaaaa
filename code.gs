const SPREADSHEET_ID = '19McvSbnc5uNJHklLMNh1e87QTUmmV77gmr0m9fJAHwQ';

function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('온라인 건의함')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function submitSuggestion(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0]; // 첫 번째 시트에 데이터 저장

    // 시트가 비어 있으면 헤더(제목 행) 자동 생성
    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출일시', '분류', '작성자', '건의 내용']);
      sheet.getRange(1, 1, 1, 4).setFontWeight('bold').setBackground('#f3f4f6');
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const category = data.category || '기타';
    const author = data.author.trim() === '' ? '익명' : data.author.trim();
    const content = data.content.trim();

    sheet.appendRow([timestamp, category, author, content]);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
