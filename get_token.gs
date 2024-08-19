function getApiToken() {
  var url = 'https://api.m2msystems.cloud/login';
  
  const mail = PropertiesService.getScriptProperties().getProperty("mail_address");
  const pass = PropertiesService.getScriptProperties().getProperty("pass");
  
  // APIリクエストのためのログイン情報
  var payload = {
    email: mail, // 実際のメールアドレスに置き換えてください
    password: pass // 実際のパスワードに置き換えてください
  };
  
  // POSTリクエストのオプションを設定
  var options = {
    method: 'post',
    contentType: 'application/json',
    // ペイロードを文字列に変換
    payload: JSON.stringify(payload),
    muteHttpExceptions: true // HTTP例外をミュートに設定
  };
  
  // APIにリクエストを送信
  var response = UrlFetchApp.fetch(url, options);
  
  // レスポンスからトークンを取得
  if (response.getResponseCode() == 200) {
    var json = JSON.parse(response.getContentText());
    var token = json.accessToken; // レスポンスの形式に応じて適宜調整してください
    Logger.log("取得したトークン: " + token);
    return token; // トークンを返す
  } else {
    Logger.log("エラーが発生しました。ステータスコード: " + response.getResponseCode());
    return null; // エラーが発生した場合はnullを返す
  }
}
