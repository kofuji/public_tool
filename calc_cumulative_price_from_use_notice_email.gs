function myFunction() {
  // 検索条件に該当するスレッド一覧を取得
  var threads = GmailApp.search('subject:ご利用のお知らせ【三井住友カード】 -label:累計済み');
  
  // スレッドを一つずつ取り出す
  threads.forEach(function(thread) {
    // スレッド内のメール一覧を取得
    var messages = thread.getMessages();
    
    // メールを一つずつ取り出す
    messages.forEach(function(message) {
      //メール受領日の取得
      var date = message.getDate();
      // メール本文
      var plainBody = message.getPlainBody();
      // 店舗
      var store = plainBody.match(/◇利用先：.+/);
      // 日時
      var price = plainBody.match(/◇利用金額：(.+)(円|JPY)/);
      var price2 = price[1];
      //お問い合わせ概要
      var outline = plainBody.match(/お問い合わせ概要: (.*)/);
      
      // 書き込むシートを取得
      var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  
      // 最終行を取得
      var lastRow = sheet.getLastRow() + 1;

      // セルを取得して値を転記
      sheet.getRange(lastRow, 1).setValue(date);
      sheet.getRange(lastRow, 2).setValue(store);
      sheet.getRange(lastRow, 3).setValue(price);
      sheet.getRange(lastRow, 4).setValue(price2);

      // 受領日の昇順に並べ替え
      var narabekae = sheet.getRange('A2:J');
      narabekae.sort({column: 1, ascending: true})
    });
    
    // スレッドに転記済みラベルを付ける
    var label = GmailApp.getUserLabelByName('累計済み');
    thread.addLabel(label);
  });

  // メールを送信
  var sheet = SpreadsheetApp.getActive().getSheetByName('シート1');
  var amount = sheet.getRange('F1').getValue();
  MailApp.sendEmail({
    to: 'xxxxxxxx@gmail.com, xxxxxxxx@gmail.com',
    subject: '今月の食費使用状況',
    body: '今月の食費は、現在' + amount + '円使用しています。'
  });  
}
