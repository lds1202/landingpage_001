var SHEET_ID = '1QRf2_lzRankZ19PLWVlcAYWdhThAYzqJBM1lXFmPoiI';
var SHEET_NAME = '상담신청';

function doPost(e) {
  var sheet = getOrCreateSheet_();
  var params = e.parameter || {};
  var parentPhone = formatPhoneNumber_(params.parentPhone || '');
  var grade = resolveGrade_(params.grade, params.gradeChoice, params.gradeOther);

  ensureHeader_(sheet);

  var row = new Array();
  row.push(new Date());
  row.push(params.submittedAt || '');
  row.push(params.source || '');
  row.push(params.landingPage || '');
  row.push(params.pageUrl || '');
  row.push(parentPhone);
  row.push(grade);
  row.push(params.school || '');
  row.push(params.mathLevel || '');
  row.push(params.score || '');
  row.push(params.classType || '');
  row.push(params.concern || '');
  row.push(params.privacyAgree || '');
  sheet.appendRow(row);

  return ContentService.createTextOutput('{"ok":true}').setMimeType(ContentService.MimeType.JSON);
}

function formatPhoneNumber_(value) {
  var numbers = String(value || '').replace(/\D/g, '').slice(0, 11);

  if (numbers.length <= 3) return numbers;
  if (numbers.length <= 7) return numbers.replace(/(\d{3})(\d+)/, '$1-$2');

  return numbers.replace(/(\d{3})(\d{4})(\d+)/, '$1-$2-$3');
}

function resolveGrade_(grade, gradeChoice, gradeOther) {
  var selectedGrade = String(grade || '').trim();
  if (selectedGrade) return selectedGrade;

  var selectedChoice = String(gradeChoice || '').trim();
  if (selectedChoice === '기타') return String(gradeOther || '').trim();

  return selectedChoice;
}

function doGet() {
  return ContentService.createTextOutput('Dunsan landing page form endpoint is running.').setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet_() {
  var spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeader_(sheet) {
  var headers = new Array();
  headers.push('접수일시');
  headers.push('브라우저 제출시각');
  headers.push('유입페이지');
  headers.push('랜딩페이지');
  headers.push('페이지URL');
  headers.push('학부모 연락처');
  headers.push('학생 학년');
  headers.push('학교명');
  headers.push('현재 수학 상태');
  headers.push('현재 점수/등급');
  headers.push('희망 수업 방식');
  headers.push('고민 내용');
  headers.push('개인정보 동의');

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues(new Array(headers));
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#EFF4FF');
}
