/* 작성자: Jiin Jeong
 * 작성일: 2018년 7월 19일
 * 설명서: 구글 스프레드시트에 있는 연락처 기반으로 이메일 보내기 - 팝업생성
 */

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('연락')
    .addItem('이메일', 'rangePopUp')
    .addToUi();
}

// 팝업 생성.
function popUp(message){
  var ui = SpreadsheetApp.getUi();
  var result = ui.prompt(
      '내용을 입력해주시길 바랍니다.',
      message,
      ui.ButtonSet.OK_CANCEL);

  // 결과 가져오기.
  var button = result.getSelectedButton();
  var content = result.getResponseText();

  if (button == ui.Button.OK) {
    return content;
  }
  else if (button == ui.Button.CANCEL) {
    return false;
  }
  else if (button == ui.Button.CLOSE) {
    ui.alert('창을 닫았습니다.');
    return false;
  }
}

// 범위 설정하는 팝업 생성.
function rangePopUp(){
  var sheetName = popUp("분석하고 싶은 시트 이름을 입력하세요. \n"+
                    "*철자와 띄어쓰기 주의");
  var startRow = parseInt(popUp("시작할 행 수를 입력하세요. \n"), 10);  // 시작 행 숫자로
  var endRow = parseInt(popUp("마지막 행 수를 입력하세요. \n"), 10);  // 마지막 행 숫자로

  sendEmail4(sheetName, startRow, endRow);
}

// 정보 가져오기.
function getData(sheetName, startRow, endRow) {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = spreadsheet.getSheetByName(sheetName);  // 이름으로 불러오기.
  var startCol = 1;
  var numRows = endRow - startRow + 1;
  var numCols = 4;
  
  var dataRange = sheet.getRange(startRow, startCol, numRows, numCols)
  var data = dataRange.getValues();
  Logger.log(data)
  return data;
}

// 맞춤형 이메일 보내기.
function sendEmail4(sheetName, startRow, endRow) {
  var data = getData(sheetName, startRow, endRow);

  // 두번째 코드와 동일.
  for (var i = 0; i < data.length; i ++) {
    var row = data[i];
    if (row[1] == "수신허용") {
      var name = row[0];
      var address = row[2];
      var body = row[3];
      body = "%s님,\n \n" + body;
      body = body.replace("%s", name);
      var subject = "[테스팅] 안녕하세요";
      MailApp.sendEmail(address, subject, body);
    }
  }
}
