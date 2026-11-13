const SPREADSHEET_ID = '19McvSbnc5uNJHklLMNh1e87QTUmmV77gmr0m9fJAHwQ';
const ADMIN_PASSWORD = '1234'; // 어드민 접속 비밀번호

function doGet(e) {
  const page = e && e.parameter && e.parameter.page;

  // 주소 뒤에 ?page=admin 이 붙으면 어드민 페이지(admin.html) 열기
  if (page === 'admin') {
    return HtmlService.createHtmlOutputFromFile('admin')
      .setTitle('온라인 건의함 - 어드민')
      .addMetaTag('viewport', 'width=device-width, initial-scale=1')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
  }

  // 기본 접속 시 일반 건의함(index.html) 열기
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('온라인 건의함')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

// 건의사항 저장 (index.html 용)
function submitSuggestion(data) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];

    if (sheet.getLastRow() === 0) {
      sheet.appendRow(['제출일시', '분류', '작성자', '건의 내용', '처리상태']);
      sheet.getRange(1, 1, 1, 5).setFontWeight('bold').setBackground('#f3f4f6');
    }

    const timestamp = Utilities.formatDate(new Date(), 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
    const category = data.category || '기타';
    const author = data.author.trim() === '' ? '익명' : data.author.trim();
    const content = data.content.trim();

    sheet.appendRow([timestamp, category, author, content, '접수']);

    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 어드민 목록 불러오기 (admin.html 용)
function getSuggestions(password) {
  if (String(password).trim() !== ADMIN_PASSWORD) {
    return { success: false, message: '비밀번호가 일치하지 않습니다.' };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];
    const lastRow = sheet.getLastRow();

    if (lastRow <= 1) {
      return { success: true, data: [] };
    }

    const numCols = Math.max(sheet.getLastColumn(), 5);
    const values = sheet.getRange(2, 1, lastRow - 1, numCols).getValues();
    const list = [];

    for (let i = 0; i < values.length; i++) {
      const row = values[i];
      let timestamp = row[0];
      if (timestamp instanceof Date) {
        timestamp = Utilities.formatDate(timestamp, 'Asia/Seoul', 'yyyy-MM-dd HH:mm:ss');
      } else {
        timestamp = String(timestamp || '');
      }

      list.push({
        rowIndex: i + 2,
        timestamp: timestamp,
        category: String(row[1] || '기타'),
        author: String(row[2] || '익명'),
        content: String(row[3] || ''),
        status: String(row[4] || '접수')
      });
    }

    return { success: true, data: list.reverse() };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}

// 처리상태 변경 (admin.html 용)
function updateStatus(password, rowIndex, newStatus) {
  if (String(password).trim() !== ADMIN_PASSWORD) {
    return { success: false, message: '권한이 없습니다.' };
  }

  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheets()[0];

    if (!sheet.getRange(1, 5).getValue()) {
      sheet.getRange(1, 5).setValue('처리상태').setFontWeight('bold').setBackground('#f3f4f6');
    }

    sheet.getRange(rowIndex, 5).setValue(newStatus);
    return { success: true };
  } catch (error) {
    return { success: false, message: error.toString() };
  }
}
