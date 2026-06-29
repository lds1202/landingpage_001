const SHEET_ID = '1QRf2_lzRankZ19PLWVlcAYWdhThAYzqJBM1lXFmPoiI';
const SHEET_NAME = '상담신청';

function doPost(e) {
  const sheet = getOrCreateSheet_();
  const params = e.parameter || {};

  ensureHeader_(sheet);

  sheet.appendRow([
    new Date(),
    params.submittedAt || '',
    params.source || '',
    params.pageUrl || '',
    params.parentPhone || '',
    params.grade || '',
    params.school || '',
    params.mathLevel || '',
    params.score || '',
    params.classType || '',
    params.concern || '',
    params.privacyAgree || ''
  ]);

  return ContentService
    .createTextOutput(JSON.stringify({ ok: true }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doGet() {
  return ContentService
    .createTextOutput('Dunsan landing page form endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getOrCreateSheet_() {
  const spreadsheet = SpreadsheetApp.openById(SHEET_ID);
  return spreadsheet.getSheetByName(SHEET_NAME) || spreadsheet.insertSheet(SHEET_NAME);
}

function ensureHeader_(sheet) {
  if (sheet.getLastRow() > 0) return;

  sheet.appendRow([
    '접수일시',
    '브라우저 제출시각',
    '유입페이지',
    '페이지URL',
    '학부모 연락처',
    '학생 학년',
    '학교명',
    '현재 수학 상태',
    '현재 점수/등급',
    '희망 수업 방식',
    '고민 내용',
    '개인정보 동의'
  ]);

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, 12).setFontWeight('bold').setBackground('#EFF4FF');
}
