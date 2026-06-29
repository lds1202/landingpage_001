const SHEET_ID = '1QRf2_lzRankZ19PLWVlcAYWdhThAYzqJBM1lXFmPoiI';
const SHEET_NAME = '상담신청';

function doPost(e) {
  const sheet = getOrCreateSheet_();
  const params = e.parameter || {};
  const parentPhone = formatPhoneNumber_(params.parentPhone || '');
  const grade = resolveGrade_(params.grade, params.gradeChoice, params.gradeOther);

  ensureHeader_(sheet);

  sheet.appendRow([
    new Date(),
    params.submittedAt || '',
    params.source || '',
    params.landingPage || '',
    params.pageUrl || '',
    parentPhone,
    grade,
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

function formatPhoneNumber_(value) {
  const numbers = String(value || '').replace(/\D/g, '').slice(0, 11);

  if (numbers.length <= 3) return numbers;
  if (numbers.length <= 7) return numbers.replace(/(\d{3})(\d+)/, '$1-$2');

  return numbers.replace(/(\d{3})(\d{4})(\d+)/, '$1-$2-$3');
}

function resolveGrade_(grade, gradeChoice, gradeOther) {
  const selectedGrade = String(grade || '').trim();
  if (selectedGrade) return selectedGrade;

  const selectedChoice = String(gradeChoice || '').trim();
  if (selectedChoice === '기타') return String(gradeOther || '').trim();

  return selectedChoice;
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
  const headers = [
    '접수일시',
    '브라우저 제출시각',
    '유입페이지',
    '랜딩페이지',
    '페이지URL',
    '학부모 연락처',
    '학생 학년',
    '학교명',
    '현재 수학 상태',
    '현재 점수/등급',
    '희망 수업 방식',
    '고민 내용',
    '개인정보 동의'
  ];

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length).setFontWeight('bold').setBackground('#EFF4FF');
}
