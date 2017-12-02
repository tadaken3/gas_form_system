var sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("来館情報");
var statusColumn = 16;

function onFormSubmit(e) {
try{
  var row = e.range.getRow();
  var formData = e.namedValues;
  var inputterMail   = formData['入力者メールアドレス'][0];
  var password       = formData['来訪ステータス'][0];
  Logger.log(formData)
  const corretPassword = 'password';
  
  //IDをセット
  var Properties = PropertiesService.getScriptProperties();
  uid = Number(Properties.getProperty('UniqueID')) + Number(1);
  var today = Utilities.formatDate(new Date() , 'JST' , 'YYYYMMdd');
  var visitId = today + '-' +　uid;
  sheet.getRange(row, 1).setValue(visitId);
  Properties.setProperty('UniqueID',uid);
  
  notifyChatwork(inputterMail);
  
  if (password == corretPassword){
     var title = "登録が完了しました。";
     var body  = createMailBody(formData);
     updateStatus("",row);
     createCalendarEvent(visitId,formData)
  }else{ 
     var title = "パスワードが違います";
     var body  = "パスワードをご確認下さい"
     updateStatus('error',row);
  }
  
  GmailApp.sendEmail(inputterMail,title,body);
 
  }catch (e) {
    notifyChatwork(e.message);
  }
  
}

function notifyChatwork(body) {
  var roomId = "66017872" //テスト用
  var client = ChatWorkClient.factory({token: "e37e39c36c00d282825fc520aef02782"});
  client.sendMessage({room_id: roomId, body: body}); 
}

function createMailBody(formData){
  //必須項目のみ
  var attendDate      = formData['来訪日'][0];
  var attendTime      = formData['来訪時間'][0];
  
  var inputterCompany = formData['入力者社名'][0];
  var inputterName    = formData['入力者氏名'][0];
  var inputterMail    = formData['入力者メールアドレス'][0];
  var inputterPhone   = formData['入力者電話番号'][0];

  var visitorName     = formData['来訪予定者氏名カナ'][0];
 
  
  var body
   = "以下の内容で、登録が完了致しました。\n\n"
   + "--------------------------------------\n"
   + "■入力者情報\n"
   + "入力者社名：" + inputterCompany + "\n"
   + "入力者氏名：" + inputterName + "\n"
   + "入力者アドレス：" + inputterMail + "\n"
   + "入力者電話番号：" + inputterPhone + "\n"
   + "\n"
   + "■入力者情報\n"
   + "来訪日：" + attendDate + "\n"
   + "来訪時間：" + attendTime + "\n"
   + "来訪者氏名：" + visitorName + "様\n";
   return body;
}

function updateStatus(status,row){
  sheet.getRange(row, statusColumn).setValue(status);
}

function createCalendarEvent(visitId,formData){
  var attendDate      = formData['来訪日'][0];
  var attendTime      = formData['来訪時間'][0];
  var visitorName     = formData['来訪予定者氏名カナ'][0];
  var inputterName    = formData['入力者氏名'][0];
  
  var title = visitId + ":" + visitorName + ":" + inputterName;
  var timeStamp = new Date(attendDate + " " + attendTime);
  
  var calendar = CalendarApp.getCalendarById('mvd0drud68381rc5b04fga8pj4@group.calendar.google.com'); 
  calendar.createEvent(title ,timeStamp, timeStamp);
}

function changeStatusToDone(){
try{
  var statusCell = sheet.getActiveCell();
  var statusRow　　　 = statusCell.getRow();
  
  var inputterMail       = sheet.getRange(statusRow, 7).getValue();
  var visitorName        = sheet.getRange(statusRow, 10).getValue();
  var visitorActcualNum  = sheet.getRange(statusRow, 11).getValue();
  
  if(statusCell.getColumn()==statusColumn　&& statusCell.getValue()=='done'){ 
  　var message =　'来訪者が来たことお知らせします。メールを配信してもよろしいですか？';
    var button = Browser.msgBox(message, Browser.Buttons.YES_NO);
    if (button == "yes"){    
      var title = visitorName +　"様がいらっしゃいました。"
      var body  = visitorName +　"様が" + visitorActcualNum + "名でいっらしゃいました。" ;
      GmailApp.sendEmail(inputterMail,title,body);
    }else{
      //init Status
      statusCell.setValue("");
    }
 
  }
 }catch(e){
   notifyChatwork(e.message);
 }
}



function archiveData(){
    var row = sheet.getLastRow();
    var column = sheet.getLastColumn();

    var archiveSpreedsheet = SpreadsheetApp.openById('1n9eeh0HpZrxmhd9Segx7g9zb3L6ckyiK68TSWh7I8Js');
    
    var archiveSheet = archiveSpreedsheet.getSheetByName('archive');
    var arrData = sheet.getDataRange().getValues();
    var maxRow = arrData.length;
    var moveDataToAchive =[];
    var notMoveData =[];
    
    //0はヘッダーのためスキップ
    for(var i=1; i<maxRow; i++){
      var　ｓecurityCardStatus = arrData[i][13];
      if(ｓecurityCardStatus == '返却完了'){
        moveDataToAchive.push(arrData[i]);
      }else{
        notMoveData.push(arrData[i]);
      }
    }
    
    var moveDataCnt    = moveDataToAchive.length;
    var notMoveDataCnt = notMoveData.length
    var dataCnt = moveDataCnt + notMoveDataCnt;
    
    if(moveDataCnt==0){
      throw new Error("移行するデータはありませんでした。");//exitの代わり
    }
    
    //念のためbackup作成
    //backupシート初期化
    var sheets = archiveSpreedsheet.getSheets();
    for(var cntSheet = 0; cntSheet < sheets.length; cntSheet++){
    if(sheets[cntSheet].getName()=='backup'){
        var backupSheet = archiveSpreedsheet.getSheetByName('backup');
        archiveSpreedsheet.deleteSheet(backupSheet);
		break;
      }
    }
    
    sheet.copyTo(archiveSpreedsheet).setName('backup');
    var backupSheet = archiveSpreedsheet.getSheetByName('backup');
    var backupDataCnt = backupSheet.getLastRow() - 1;
     
    if (dataCnt == backupDataCnt){
      //init
      var row = sheet.getLastRow();
      var column = sheet.getLastColumn();
      sheet.deleteRows(2, row)
      
      //set archive sheet
      var archiveSheetRow = archiveSheet.getLastRow();
      var moveDataRow     = moveDataToAchive.length;
      var moveDatacolumn  = moveDataToAchive[0].length;
      archiveSheet.getRange(archiveSheetRow+1 ,1 ,moveDataRow ,moveDatacolumn).setValues(moveDataToAchive);
     
     //set row sheet
      var notMoveDataRow     = notMoveData.length;
      var notMoveDatacolumn  = notMoveData[0].length;
      sheet.getRange(2 ,1 ,notMoveDataRow ,notMoveDatacolumn).setValues(notMoveData);
     
   }else if(dataCnt != backupDataCnt) {
     //どこかにbackupシートと元データを確認するように通知投げる
     Browser.msgBox("クリーニング処理がうまくいきませんでした。ご確認下さい")
   }
   archiveSpreedsheet.deleteSheet(backupSheet);
}

