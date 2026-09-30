// ===== 메가커피 발주 체크리스트 - Google Apps Script 백엔드 =====
// 이 코드를 그대로 Apps Script 편집기에 붙여넣으세요.

const SHEET_NAME = 'Checklist';

function getSheet_() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
    sheet.appendRow(['date', 'itemId', 'label', 'value', 'updatedAt']);
  }
  // 날짜/itemId 열이 구글시트에 의해 자동으로 날짜형으로 바뀌지 않도록 텍스트로 고정
  sheet.getRange('A:B').setNumberFormat('@');
  return sheet;
}

// 셀 값이 Date 객체로 저장되어 있어도 'YYYY-MM-DD' 문자열로 통일
function normalizeDate_(val) {
  if (Object.prototype.toString.call(val) === '[object Date]') {
    return Utilities.formatDate(val, Session.getScriptTimeZone(), 'yyyy-MM-dd');
  }
  return String(val);
}

function findRow_(sheet, date, itemId) {
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (normalizeDate_(data[i][0]) === date && String(data[i][1]) === itemId) {
      return i + 1; // 1-indexed row number
    }
  }
  return -1;
}

// 14일보다 오래된 기록을 매일 한 번만 정리 (같은 날 반복 호출 시 스킵)
const RETENTION_DAYS = 14;

function cleanupOldRows_(sheet) {
  const props = PropertiesService.getScriptProperties();
  const today = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
  if (props.getProperty('lastCleanup') === today) return;

  const cutoffDate = new Date();
  cutoffDate.setDate(cutoffDate.getDate() - RETENTION_DAYS);
  const cutoff = Utilities.formatDate(cutoffDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');

  const data = sheet.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    if (normalizeDate_(data[i][0]) < cutoff) {
      sheet.deleteRow(i + 1);
    }
  }

  props.setProperty('lastCleanup', today);
}

// 프론트엔드가 값을 저장할 때 호출 (POST)
function doPost(e) {
  const body = JSON.parse(e.postData.contents);
  const sheet = getSheet_();
  cleanupOldRows_(sheet);
  const now = new Date().toISOString();

  const date = body.date;
  const itemId = body.itemId;
  const label = body.label || '';
  const value = body.value;

  const rowNum = findRow_(sheet, date, itemId);
  if (rowNum > 0) {
    sheet.getRange(rowNum, 3).setValue(label);
    sheet.getRange(rowNum, 4).setValue(value);
    sheet.getRange(rowNum, 5).setValue(now);
  } else {
    sheet.appendRow([date, itemId, label, value, now]);
  }

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

// 프론트엔드가 특정 날짜의 저장된 값을 불러올 때 호출 (GET ?date=YYYY-MM-DD)
// 날짜 없이 호출하면 저장된 전체 날짜 목록을 반환 (GET ?dates=1)
function doGet(e) {
  const sheet = getSheet_();
  const data = sheet.getDataRange().getValues();

  if (e.parameter.dates) {
    const dateSet = {};
    for (let i = 1; i < data.length; i++) {
      dateSet[normalizeDate_(data[i][0])] = true;
    }
    const dates = Object.keys(dateSet).sort().reverse();
    return ContentService
      .createTextOutput(JSON.stringify({ dates: dates }))
      .setMimeType(ContentService.MimeType.JSON);
  }

  const date = e.parameter.date;
  const result = {};
  for (let i = 1; i < data.length; i++) {
    if (normalizeDate_(data[i][0]) === date) {
      result[data[i][1]] = data[i][3];
    }
  }
  return ContentService
    .createTextOutput(JSON.stringify({ date: date, values: result }))
    .setMimeType(ContentService.MimeType.JSON);
}
