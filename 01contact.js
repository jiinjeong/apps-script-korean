/* 작성자: Jiin Jeong
 * 작성일: 2018년 7월 15일
 * 설명서: 구글 스프레드시트에 있는 연락처 기반으로 이메일 보내기
 */

// 스프레드시트를 열면 자동으로 실행되는 함수.
function onOpen() {
  var ui = SpreadsheetApp.getUi();  // 구글 스프레드시트 앱의 유저 인터페이스.
  ui.createMenu('연락')
    .addItem('이메일', 'sendEmail')  // 새로운 메뉴를 만듦. (메뉴 제목, 기능)
    .addToUi();  // 유저 인터페이스에 추가.
}

// 정보를 스프레드시트에서 가져오기.
function getData() {
  var sheet = SpreadsheetApp.getActiveSheet();  // 현재 사용 중인 시트
  var startRow = 2;  // 시작 행 (참고로 1부터 인덱스 시작.)
  var startCol = 1;  // 시작 열
  var numRows = 1;  // 행 수
  var numCols = 3;  // 열 수
  
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)  // 범위
  var data = dataRange.getValues();  // 데이터 값만 가져오기. 2D-배열 형태.
  Logger.log(data)  // 로그에 데이터 출력 (디버깅 목적).
  return data;  // 값 반환.
}

// 이메일 보내기. 함수 이름은 위에 만든 메뉴 기능 이름과 동일해야 합니다.
function sendEmail() {
  var data = getData();  // 정보 불러오기.

  for (var i = 0; i < data.length; i ++) {
    var row = data[i];  // 각 열
    var address = row[1];  // 두번째 열: 이메일 주소
    var body = row[2];  // 세번째 열: 이메일 내용
    var subject = "안녕하세요";  // 이메일 제목
    
    MailApp.sendEmail(address, subject, body);  // 메일 앱을 통해 보내기.
  }
}
