/**
 * 모의고사집 개정 의견 조사 - 응답 수집 백엔드 (v4 · 진단 기능 포함)
 * 2CPlab.
 *
 * 배포 후 /exec 주소를 브라우저에서 열면 연결 상태를 확인할 수 있습니다.
 */

var SHEET_NAME = '응답';
var ACCESS_KEY = 'cpl2026';

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
  '도움도(1-5)', '해설상세도', '문항오류', '[서술] 오류내용', '테스트와이즈니스', '사용방식', '가장도움된점', '[서술] 아쉬운점',
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

    return jsonOut_({
      result: 'ok',
      rows: sheet.getLastRow(),
      sheet: sheet.getName(),
      file: sheet.getParent().getName()
    });

  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });

  } finally {
    lock.releaseLock();
  }
}


/**
 * 브라우저에서 /exec 주소를 열면 현재 연결 상태를 보여줍니다.
 * 어느 스프레드시트에 기록되고 있는지 URL로 직접 확인할 수 있습니다.
 */
function doGet() {
  try {
    var ss = SpreadsheetApp.getActiveSpreadsheet();

    if (!ss) {
      return jsonOut_({
        result: 'not_bound',
        message: '이 스크립트가 스프레드시트에 연결되어 있지 않습니다. 구글 시트에서 확장 프로그램 > Apps Script 로 다시 만들어 주십시오.'
      });
    }

    var sheet = ss.getSheetByName(SHEET_NAME);

    return jsonOut_({
      result: 'alive',
      file: ss.getName(),
      fileUrl: ss.getUrl(),
      allTabs: ss.getSheets().map(function (s) { return s.getName(); }),
      targetTab: SHEET_NAME,
      targetTabExists: !!sheet,
      rows: sheet ? sheet.getLastRow() : 0
    });

  } catch (err) {
    return jsonOut_({ result: 'error', message: String(err) });
  }
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
