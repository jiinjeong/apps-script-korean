/* 작성자: Jiin Jeong
 * 작성일: 2018년 7월 16일
 * 설명서: 구글 스프레드시트에 있는 연락처 기반으로 이메일 보내기 - 내용 맞춤형
 */

// 첫번째 코드와 동일.
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('연락')
    .addItem('맞춤형 이메일', 'sendEmail2')
    .addItem('구글문서 활용한 이메일', 'sendEmail3')
    .addToUi();
}

// ***행수, 열수 바꾸기.
function getData() {
  var sheet = SpreadsheetApp.getActiveSheet();
  var startRow = 2;
  var startCol = 1;
  var numRows = 2;
  var numCols = 4;
  
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
  var data = dataRange.getValues();
  Logger.log(data)
  return data;
}

// 맞춤형 이메일 보내기. ***열 내용 바꾸기.
function sendEmail2() {
  var data = getData();

  for (var i = 0; i < data.length; i ++) {
    var row = data[i];
    if (row[1] == "수신허용") {  // 두번째 열: 수신동의 여부 (허용일 때만 실행)
      var name = row[0];  // 첫번째 열: 이름
      var address = row[2];  // 세번째 열: 이메일 주소
      var body = row[3];  // 네번째 열: 이메일 내용
      body = "%s님,\n \n" + body;  // \n는 새줄 캐릭터
      body = body.replace("%s", name);  // 이름으로 대체
      var subject = "[테스팅] 안녕하세요";  // 이메일 제목
      MailApp.sendEmail(address, subject, body);  // 메일 앱을 통해 보내기.
    }
  }
}

// 구글 문서 내용 가져오기.
function Content() {
  var doc = DocumentApp.openById("문서ID");  // 문서 열기. ***문서ID 바꾸기.
  var body = doc.getBody();  // 문서 내용
  var text = body.getText();  // 문서 글감 가져오기.
  return text;
}

// 구글 문서를 활용한 맞춤형 이메일 보내기.
function sendEmail3() {
  var data = getData();

  for (var i = 0; i < data.length; i ++) {
    var row = data[i];
    if (row[1] == "수신허용") {  // 두번째 열: 수신동의 여부 (허용일 때만 실행)
      var name = row[0];  // 첫번째 열: 이름
      var address = row[2];  // 세번째 열: 이메일 주소
      var body = "%s님,\n \n" + Content();  // \n는 새줄 캐릭터
      body = body.replace("%s", name);  // 이름으로 대체
      var subject = "[테스팅] 안녕하세요";  // 이메일 제목
      MailApp.sendEmail(address, subject, body);  // 메일 앱을 통해 보내기.
    }
  }
}
