function getPictureIdList() {
  getPicsId();
  removeUsedPicsId();
}


function getPicsId() {
  const sheetId = '';
  const folderId = '';

  var sheets = SpreadsheetApp.openById(sheetId);
  var sheet = sheets.getSheetByName('current');


  var folder = DriveApp.getFolderById(folderId);
  var images = folder.getFilesByType(MimeType.JPEG);

  var imageIds = []

  while(images.hasNext()){
    var image = images.next();
    imageIds.push([image.getId()])
    //console.log(image.getId());
  }

  //console.log(imageIds);
  //console.log(imageIds.length);

  sheet.getRange("A:A").clear(); //現在シートに残っているIDを削除する
  var dataRange = sheet.getRange(1,1,imageIds.length);
  dataRange.setValues(imageIds); 

  // ランダムに並び替え
  dataRange.randomize();
}


function removeUsedPicsId() {
  const sheetId = '';
  var sheets = SpreadsheetApp.openById(sheetId);
  var sheet = sheets.getSheetByName('current');
  var lastRow = sheet.getLastRow();

  //console.log(lastRow);

  var fomulaRange = sheet.getRange(1,2,lastRow);

  fomulaRange.setFormula("=COUNTIF(used!A:A,A1)");


  for(var i=lastRow; i>0; i--){
    //console.log(i);
    var checkRange = sheet.getRange(i, 2, 1);
    //console.log(i + '--' + checkRange.getValue());

    if(!checkRange.getValue()==0){
      checkRange.setBackground('green');  
      sheet.deleteRow(i)
    }
  }

  // 重複データ削除後のデータ数を取得
  var lastRow = sheet.getLastRow();
  // データ数が0の場合はusedシートのデータをすべて削除して、再取得する
  if(lastRow==0){
    var sheetUsed = sheets.getSheetByName('used');
    sheetUsed.getRange("A:A").clear(); //現在シートに残っているIDを削除する
    console.log("再実行");
    getPicsId();
  }
}