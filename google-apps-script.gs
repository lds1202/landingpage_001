var SHEET_ID = '1QRf2_lzRankZ19PLWVlcAYWdhThAYzqJBM1lXFmPoiI';
var LOGISTICS_SHEET_NAME = '물류대행신청';

function doPost(e) {
  var params = e.parameter || {};
  var formType = String(params.formType || '').toLowerCase();

  if (formType === '3pl' || formType === 'logistics') {
    return handleLogisticsPost_(params);
  }

  return handleEducationPost_(params);
}

function handleLogisticsPost_(params) {
  var sheet = getSheetByName_(LOGISTICS_SHEET_NAME);
  ensureLogisticsHeader_(sheet);

  var row = [];
  row.push(new Date());
  row.push(params.submittedAt || '');
  row.push(params.source || '');
  row.push(params.landingPage || '');
  row.push(params.pageUrl || '');
  row.push(params.companyName || '');
  row.push(params.managerName || '');
  row.push(formatPhoneNumber_(params.phone || ''));
  row.push(params.salesChannels || '');
  row.push(params.productCategory || '');
  row.push(params.monthlyShipments || '');
  row.push(params.storageNeed || '');
  row.push(params.largeCargo || '');
  row.push(params.coupangMilkrun || '');
  row.push(params.rocketGrowth || '');
  row.push(params.currentLogistics || '');
  row.push(params.location || '');
  row.push(params.message || '');
  row.push(params.privacyAgree || '');

  sheet.appendRow(row);
  return jsonResponse_({ ok: true, type: '3pl' });
}

function handleEducationPost_(params) {
  var sheet = getFirstSheet_();
  var parentPhone = formatPhoneNumber_(params.parentPhone || '');
  var grade = resolveGrade_(params.grade, params.gradeChoice, params.gradeOther);

  ensureEducationHeader_(sheet);

  var row = [];
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
  return jsonResponse_({ ok: true, type: 'education' });
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
  return ContentService
    .createTextOutput('Landing page form endpoint is running.')
    .setMimeType(ContentService.MimeType.TEXT);
}

function getSpreadsheet_() {
  return SpreadsheetApp.openById(SHEET_ID);
}

function getFirstSheet_() {
  var sheets = getSpreadsheet_().getSheets();
  return sheets[0];
}

function getSheetByName_(sheetName) {
  var spreadsheet = getSpreadsheet_();
  var sheet = spreadsheet.getSheetByName(sheetName);
  if (!sheet) {
    sheet = spreadsheet.insertSheet(sheetName);
  }
  return sheet;
}

function ensureEducationHeader_(sheet) {
  var headers = [
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

  setHeader_(sheet, headers, '#EFF4FF');
}

function ensureLogisticsHeader_(sheet) {
  var headers = [
    '접수일시',
    '브라우저 제출시각',
    '유입페이지',
    '랜딩페이지',
    '페이지URL',
    '업체명',
    '담당자명',
    '연락처',
    '판매 채널',
    '주요 상품군',
    '월 평균 출고량',
    '보관 필요 여부',
    '대형화물 여부',
    '쿠팡 밀크런 필요 여부',
    '로켓그로스 필요 여부',
    '현재 물류 방식',
    '위치/지역',
    '상담 희망 내용',
    '개인정보 동의'
  ];

  setHeader_(sheet, headers, '#F1ECFF');
}

function setHeader_(sheet, headers, background) {
  if (sheet.getLastRow() === 0) {
    sheet.appendRow(headers);
  } else {
    sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  }

  sheet.setFrozenRows(1);
  sheet.getRange(1, 1, 1, headers.length)
    .setFontWeight('bold')
    .setBackground(background);
}

function jsonResponse_(payload) {
  return ContentService
    .createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}
