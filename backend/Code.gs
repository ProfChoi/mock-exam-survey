/**
 * 모의고사집 개정 의견 조사 - 응답 수집 백엔드 (v2)
 * 최규하상담심리연구소(2CPlab.)
 */

var SHEET_NAME = '응답';
var ACCESS_KEY = 'cpl2026';   // index.html 의 ACCESS_KEY 와 동일해야 함

var FIELDS = [
  'timestamp',
  'q1', 'q2',
  'q3', 'q4', 'q5', 'q6', 'q7', 'q7_txt',
  'q8', 'q9', 'q10', 'q10_txt', 'q11', 'q12', 'q13', 'q14',
  'q15', 'q15_etc', 'q16', 'q17', 'q18',
  'q19', 'q19_etc', 'q20', 'q21', 'q22', 'q23'
];

var HEADERS = [
  '제출시각',
  '구매시점', '푼문항수',
  '가채점체감', '난이도(1-5)', '난이도불일치과목', '영역일치도(1-5)', '누락내용유무', '[서술] 누락내용',
  '도움도(1-5)', '해설상세도', '문항오류', '[서술] 오류내용', '테스트와이즈니스', '단권형태', '가장도움된점', '[서술] 아쉬운점',
  '인지경로', '인지경로_기타', '[서술] 망설임요인', '가격평가(49000원)', '추천의향(0-10)',
  '구매처', '구매처_기타', '주문번호', '인용동의', '개인정보동의', '연락처'
];


function doPost(e) {
  var lock = LockService.getScriptLock();
  lock.waitLock(20000);

  try {
    var p = (e && e.parameter) ? e.parameter : {};

    if (p.k !== ACCESS_KEY) {
      return jsonOut_({ result: 'denied' });
    }

    var sheet = getSheet_();
    var row = FIELDS.map(function (key) {
      return p[key] !== undefined ? p[key] : '';
    });

    sheet.appendRow(row);
    return jsonOut_({ result: 'ok' });

  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });

  } finally {
    lock.releaseLock();
  }
}


function doGet() {
  return jsonOut_({ result: 'alive' });
}


function getSheet_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sheet = ss.getSheetByName(SHEET_NAME);

  if (!sheet) {
    sheet = ss.insertSheet(SHEET_NAME);
  }

  if (sheet.getLastRow() === 0) {
    sheet.appendRow(HEADERS);
    var head = sheet.getRange(1, 1, 1, HEADERS.length);
    head.setFontWeight('bold');
    head.setBackground('#44546A');
    head.setFontColor('#FFFFFF');
    sheet.setFrozenRows(1);

    // 서술형 열을 넓게 (누락내용 9, 오류내용 13, 아쉬운점 17, 망설임 20)
    [9, 13, 17, 20].forEach(function (col) {
      sheet.setColumnWidth(col, 400);
    });
  }

  return sheet;
}


function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
