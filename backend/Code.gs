/**
 * 모의고사집 개정 의견 조사 - 응답 수집 백엔드
 * 최규하상담심리연구소(2CPlab.)
 *
 * 설치:
 *  1) 구글 시트 생성 -> 확장 프로그램 -> Apps Script
 *  2) 이 코드 전체를 붙여넣고 저장
 *  3) 배포 -> 새 배포 -> 웹 앱
 *     - 실행 계정: 나
 *     - 액세스 권한: 모든 사용자
 *  4) 발급된 /exec URL 을 index.html 의 ENDPOINT 에 입력
 */

var SHEET_NAME  = '응답';
var ACCESS_KEY  = 'cpl2026';   // index.html 의 ACCESS_KEY 와 동일해야 함

var FIELDS = [
  'timestamp',
  'q1', 'q1_etc', 'q2',
  'q3', 'q4', 'q5', 'q6',
  'q7', 'q8', 'q9', 'q10',
  'q11', 'q12', 'q13', 'q14', 'q15',
  'q16', 'q16_etc', 'q17', 'q18', 'q19',
  'q20', 'q21', 'q22'
];

var HEADERS = [
  '제출시각',
  '구매처', '구매처_기타', '주문번호',
  '구매시점', '푼문항수', '회독수', '병행교재',
  '난이도(1-5)', '난이도불일치과목', '영역일치도(1-5)', '[서술] 시험엔나왔으나 교재에없던내용',
  '해설상세도', '문항오류', '테스트와이즈니스', '분량판형', '[서술] 아쉬운점',
  '인지경로', '인지경로_기타', '[서술] 망설임요인', '적정가격', '추천의향(0-10)',
  '인용동의', '개인정보동의', '연락처'
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
    [12, 17, 20].forEach(function (col) {
      sheet.setColumnWidth(col, 420);
    });
  }

  return sheet;
}


function jsonOut_(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
