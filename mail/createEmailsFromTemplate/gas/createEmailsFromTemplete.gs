function saveDraftsFromTemplate() {
  // スプレッドシートのデータを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  var data = sheet.getDataRange().getValues();
  
  // Google ドライブ内のテンプレートファイルIDを指定
  var templateFileId = 'YOUR_TEMPLATE_FILE_ID';  // テンプレートのGoogleドキュメントのファイルIDを入力
  
  // テンプレートをGoogle ドキュメントから取得
  var templateDoc = DocumentApp.openById(templateFileId);
  var templateBody = templateDoc.getBody().getText();
  
  // 各行のデータに対して下書きを保存
  for (var i = 1; i < data.length; i++) { // 1行目はヘッダーなのでスキップ
    var row = data[i];
    var emailTo = row[0];  // A列 (To)
    var emailCc = row[1];  // B列 (CC)
    var emailBcc = row[2]; // C列 (BCC)
    var subject = row[3];  // D列 (件名)
    
    // プレースホルダのデータを取得 (E列以降)
    var placeholders = row.slice(4); // E列以降を取得 (param1~param10)
    
    // テンプレートのコピーを作成
    var message = templateBody;
    
    // プレースホルダを埋め込む
    for (var j = 0; j < placeholders.length; j++) {
      var param = placeholders[j];
      
      if (param === "") {
        // プレースホルダが空の場合、置換を停止して次の行に進む
        break;
      }
      
      // プレースホルダを置換
      var placeholder = "{param" + (j + 1) + "}";
      message = message.replace(placeholder, param);
    }
    
    // 置換が完了したメッセージを下書きとして保存
    GmailApp.createDraft(emailTo, subject, message, {
      cc: emailCc,
      bcc: emailBcc
    });
  }
}

function insertHeaders() {
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // 1行目にタイトルを挿入
  var headers = ['To', 'CC', 'BCC', '件名', 'param1', 'param2', 'param3', 'param4', 'param5', 'param6', 'param7', 'param8', 'param9', 'param10'];
  
  // 1行目にヘッダーを設定
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
}