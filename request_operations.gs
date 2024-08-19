function request_operation() { 
  try {
    // トークンを取得
    var token = getApiToken();
    var tokenFailed = false;

    if (token === null) {
      Logger.log("トークンを取得できませんでした。仮のペイロードを記録します。");
      tokenFailed = true; // トークン取得失敗フラグを立てる
    }

    const sheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';
    const ss = SpreadsheetApp.openById(sheetId);
    const operationSheet = ss.getSheetByName('main');  // 'main'シートを指定

    operationSheet.getRange("F3:F").clearContent();

    // 実行者のメールアドレスを取得してC1に出力
    const userEmail = Session.getActiveUser().getEmail();
    
    // タイムスタンプの取得
    const timestamp = new Date().toLocaleString(); 
    
    // C1にメールアドレスとタイムスタンプをセット
    operationSheet.getRange("C1").setValue("実行時間 : " + timestamp);

    // 'main'シートのA1セルの日付を取得
    var originCleaningDate = operationSheet.getRange("A1").getValue();
    const cleaningDate = convertDate(originCleaningDate);
    Logger.log('cleaningDate: ' + cleaningDate);

    const startRow = 3;
    var nameColumn = operationSheet.getRange("B:B").getValues(); 
    var lastRowWithData = 1;
    for (var i = 0; i < nameColumn.length; i++) {
        if (nameColumn[i][0] !== "" && nameColumn[i][0] !== null) {
            lastRowWithData = i + 1; 
        }
    }
    var operationData = operationSheet.getRange(startRow, 1, lastRowWithData - startRow + 1, 5).getValues(); // A列からD列まで取得
    Logger.log(operationData);

    var requestValues = [];
    for (var i = 0; i < operationData.length; i++) {
        requestValues.push(operationData[i]);
    }
    Logger.log(requestValues);

    // エラーメッセージを蓄積するための配列
    var errorMessages = [];

    // APIリクエスト（仮想的な実行）
    for (var i = 0; i < requestValues.length; i++) {
      const api_url = 'https://api-cleaning.m2msystems.cloud/v3/cleanings/create_with_placement';
      Logger.log(api_url);

      var payload = {
      "placement": "commonArea", 
      "commonAreaId": requestValues[i][3],  // D列から取得
      "listingId": "", 
      "cleaningDate": cleaningDate,  // 'main'シートのA1セルの日付を使用
      "note": "",  // 空の文字列
      "cleaners": [requestValues[i][2]],  // 配列として設定
      "photoTourId": requestValues[i][4],  // D列から取得
      };

      Logger.log('payload: ' + JSON.stringify(payload));

      var options = {
        'method': 'post',
        'contentType': 'application/json',
        'headers': {
          'Authorization': 'Bearer ' + token  // Bearerトークンの正しい設定
        },
        'payload': JSON.stringify(payload),
        'muteHttpExceptions': true
      };

      Logger.log(options);

      var response;
      var result;

      if (tokenFailed) {
        response = "トークン取得失敗のためリクエスト未送信";
        result = "error";  // 仮にエラーとして記録
      } else {
        try {
          response = UrlFetchApp.fetch(api_url, options).getContentText();
          result = response.includes('error') ? 'error' : 'ok';
        } catch (e) {
          Logger.log("APIリクエストエラー: " + e.message);
          response = "リクエストエラー: " + e.message;
          result = "error";
        }
      }

      Logger.log('response: ' + response);
      Logger.log('result: ' + result);

      // 結果が 'ok' でない場合にエラーメッセージを配列に追加
      if (result !== 'ok') {
        errorMessages.push("Error for row " + (i + startRow) + ": " + response);
      }

      // 'ツアーpayload'タブにペイロードと結果を記録
      operationSheet.getRange(i + 3, 6, 1, 1).setValues([[result]]);
    }

    // すべての処理が完了後、エラーメッセージをまとめて表示
    if (errorMessages.length > 0) {
      Browser.msgBox("以下のエラーが発生しました:\n" + errorMessages.join("\n"));
    }

  } catch (e) {
    Logger.log("request_operation関数内でエラーが発生しました: " + e.message);
    Browser.msgBox("request_operation関数内でエラーが発生しました: " + e.message);
  }
}


function convertDate(dateString) {
  if (!dateString) {
    return "";
  }

  const date = new Date(dateString);

  if (isNaN(date)) {
    return "";
  }

  const year = date.getFullYear();
  let month = date.getMonth() + 1;
  let day = date.getDate();

  month = (month < 10) ? '0' + month : month;
  day = (day < 10) ? '0' + day : day;

  return year + '-' + month + '-' + day;
}
