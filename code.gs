function sendEmails() {
  // 開始行と終了行をユーザーに入力させる
  var startRow = parseInt(Browser.inputBox('開始行を入力してください（ヘッダー行を除く数字）'));
  var endRow = parseInt(Browser.inputBox('終了行を入力してください'));
  
  // 送信予約時間を指定（フォーマット: 'YYYY-MM-DD HH:MM:SS'）
  var scheduledTimeInput = Browser.inputBox('送信予約日時を入力してください（例: 2023-12-31 15:00:00）');
  var scheduledTime = new Date(scheduledTimeInput);
  
  // 現在のスクリプトに時間ベースのトリガーを設定
  ScriptApp.newTrigger('sendScheduledEmails')
    .timeBased()
    .at(scheduledTime)
    .create();
  
  // ユーザーの入力値をプロパティに保存
  var scriptProperties = PropertiesService.getScriptProperties();
  scriptProperties.setProperty('START_ROW', startRow);
  scriptProperties.setProperty('END_ROW', endRow);
  
  // メール本文のテンプレートIDを保存
  var templateFileId = '1T1Scc_l_TLSuMgukWTBwFahxFxROCsvX9v9NCNGq6jE'; // あなたのファイルID
  scriptProperties.setProperty('TEMPLATE_FILE_ID', templateFileId);
}

function sendScheduledEmails() {
  var scriptProperties = PropertiesService.getScriptProperties();
  
  var startRow = parseInt(scriptProperties.getProperty('START_ROW'));
  var endRow = parseInt(scriptProperties.getProperty('END_ROW'));
  var templateFileId = scriptProperties.getProperty('TEMPLATE_FILE_ID');
  
  // テンプレートファイルを取得（Googleドキュメントの場合）
  var doc = DocumentApp.openById(templateFileId);
  
  // ドキュメントの本文を取得
  var body = doc.getBody();
  var numChildren = body.getNumChildren();
  var messageHtmlTemplate = '';

  // 各子要素をループ
  for (var i = 0; i < numChildren; i++) {
    var element = body.getChild(i);
    
    switch (element.getType()) {
      case DocumentApp.ElementType.PARAGRAPH:
        var text = element.asParagraph().getText();
        messageHtmlTemplate += '<p>' + text + '</p>';
        break;
      case DocumentApp.ElementType.TABLE:
        var table = element.asTable();
        var numRows = table.getNumRows();
        messageHtmlTemplate += '<table border="1" cellpadding="5" cellspacing="0">';
        
        for (var r = 0; r < numRows; r++) {
          var row = table.getRow(r);
          var numCells = row.getNumCells();
          messageHtmlTemplate += '<tr>';
          
          for (var c = 0; c < numCells; c++) {
            var cell = row.getCell(c);
            var cellText = cell.getText();
            messageHtmlTemplate += '<td>' + cellText + '</td>';
          }
          
          messageHtmlTemplate += '</tr>';
        }
        
        messageHtmlTemplate += '</table>';
        break;
      case DocumentApp.ElementType.LIST_ITEM:
        var listItem = element.asListItem();
        var nestingLevel = listItem.getNestingLevel();
        var glyphType = listItem.getGlyphType();
        var tag = (glyphType == DocumentApp.GlyphType.NUMBER) ? 'ol' : 'ul';

        // リストの開始タグ
        if (i == 0 || body.getChild(i - 1).getType() != DocumentApp.ElementType.LIST_ITEM) {
          messageHtmlTemplate += '<' + tag + '>';
        }

        // リスト項目
        messageHtmlTemplate += '<li>' + listItem.getText() + '</li>';

        // リストの終了タグ
        if (i == numChildren - 1 || body.getChild(i + 1).getType() != DocumentApp.ElementType.LIST_ITEM) {
          messageHtmlTemplate += '</' + tag + '>';
        }
        break;
      default:
        // 他の要素タイプは必要に応じて追加
        break;
    }
  }
  
  // アクティブなシートを取得
  var sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  
  // データ範囲を取得
  var numRows = endRow - startRow + 1;
  var dataRange = sheet.getRange(startRow, 1, numRows, sheet.getLastColumn());
  var data = dataRange.getValues();
  
  // 各行をループ処理
  for (var i = 0; i < data.length; i++) {
    var row = data[i];
    
    var companyName = row[0]; // A列（会社名）
    var recipientName = row[1]; // B列（氏名）
    var emailAddress = row[7]; // H列（メールアドレス）
    
    // メールアドレスがない場合はスキップ
    if (!emailAddress) {
      continue;
    }
    
    // 件名を設定
    var subject = recipientName + "様に独立のご連絡・ご挨拶";
    
    // プレースホルダーを置換
    var messageHtml = messageHtmlTemplate.replace(/{会社名}/g, companyName).replace(/{氏名}/g, recipientName);
    
    // メールを送信（HTML形式）
    MailApp.sendEmail({
      to: emailAddress,
      subject: subject,
      htmlBody: messageHtml
    });
  }
  
  // トリガーを削除
  deleteTriggers();
}

function deleteTriggers(){
  // このスクリプトの全てのトリガーを削除
  var triggers = ScriptApp.getProjectTriggers();
  for(var i=0; i<triggers.length; i++){
    ScriptApp.deleteTrigger(triggers[i]);
  }
}
