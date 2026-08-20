/**
 * 모의고사집 개정 의견 조사 - 응답 수집 백엔드 (v3)
 * 최규하상담심리연구소(2CPlab.)
 *
 * 총 22문항. 가채점 문항 제거 후 번호 재정렬.
 */

var SHEET_NAME = '응답';
var ACCESS_KEY = 'cpl2026';   // index.html 의 ACCESS_KEY 와 동일해야 함

var FIELDS = [
  'timestamp',
  'q1', 'q2',
  'q3', 'q4', 'q5', 'q6', 'q6_txt',
  'q7', 'q8', 'q9', 'q9_txt', 'q10', 'q11', 'q12', 'q13',
  'q14', 'q14_etc', 'q15', 'q16', 'q17',
  'q18', 'q18_etc', 'q19', 'q20', 'q21', 'q22'
];

var HEADERS = [
  '제출시각',
  '구매시점', '푼문항수',
  '난이도(1-5)', '난이도불일치과목', '영역일치도(1-5)', '누락내용유무', '[서술] 누락내용',
  '도움도(1-5)', '해설상세도', '문항오류', '[서술] 오류내용', '테스트와이즈니스', '전자책불편사항', '가장도움된점', '[서술] 아쉬운점',
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

    // 서술형 열 확장: 누락내용(8), 오류내용(12), 아쉬운점(16), 망설임(19)
    [8, 12, 16, 19].forEach(function (col) {
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

