function moveAllFunctions() {
  try {
    displayTomorrowDate();
    getBigQueryData();
    // request_operation()
  } catch (error) {
    notifySlack(error.message);
  }
}


function notifySlack(errorMessage) {
  var slackWebhookUrl = "https://hooks.slack.com/services/your/webhook/url"; // Webhook URLをここに設定
  var payload = {
    "username": "トラブル報告",  // ユーザー名を設定
    "text": "ゴミチェッカーツアーの作成に失敗しました: " + errorMessage,
    "icon_emoji": ":warning:"  // Slackに表示するアイコンを設定
  };

  var options = {
    "method": "post",
    "contentType": "application/json",
    "payload": JSON.stringify(payload)
  };

  UrlFetchApp.fetch(slackWebhookUrl, options);
}

//変更
function displayTomorrowDate() {
  const sheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';  // 使用するスプレッドシートのID
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = ss.getSheetByName('main');  // 'main'シートを指定

  // 明日の日付を取得
  const today = new Date();
  const tomorrow = new Date(today);
  tomorrow.setDate(today.getDate() + 1);

  // 日付をフォーマット
  const timeZone = "Asia/Tokyo"; // UTC+9のタイムゾーン
  const formattedDate = Utilities.formatDate(tomorrow, timeZone, 'yyyy-MM-dd');

  // A1セルに明日の日付を表示
  sheet.getRange("A1").setValue(formattedDate);
}


function getBigQueryData() {
  // 読み込みたいプロジェクトの名前
  const project_id = "m2m-core";
  
  // 出力先スプレッドシートのIDとシート名
  const ss = SpreadsheetApp.openById("1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs");
  const outputSheet = ss.getSheetByName("出力");
  const mainSheet = ss.getSheetByName("main");

  // C1セルに「実行中」と表示
  mainSheet.getRange("B1").setValue("実行中");

  // A1セルから日付を取得
  const a1Date = mainSheet.getRange("A1").getValue();
  if (!a1Date || !(a1Date instanceof Date)) {
    Logger.log("A1セルには有効な日付が含まれていません。");
    mainSheet.getRange("A1").setValue("エラー");
    return;
  }

  // 日付をUTC+9のタイムゾーンでフォーマット
  const timeZone = "Asia/Tokyo"; // UTC+9のタイムゾーン
  const workDate = Utilities.formatDate(a1Date, timeZone, 'yyyy-MM-dd');

  // 実行するクエリ
  const query_execute = `
    SELECT
      building_name,
      work_name,
      worker_name,
      cleaning_id,
      room_name_common_area_name,
      status,
      work_date,
      prefecture
    FROM
      \`m2m-core.su_wo.wo_cleaning_tour\`
    WHERE
      work_date = '${workDate}'
      AND (work_name = '通常清掃' OR work_name = '重点清掃')
      AND worker_name IS NOT NULL
      AND worker_name != ''
      AND prefecture = '東京都'
  `;

  // 実行するクエリーの確認
  Logger.log(query_execute);

  // BigQuery APIを使ってクエリを実行
  const result = BigQuery.Jobs.query(
    {
      useLegacySql: false,
      query: query_execute,
    },
    project_id
  );

  // クエリ結果の取得
  const rows = result.rows.map(row => {
    return row.f.map(cell => cell.v);
  });
  Logger.log(rows);

  outputSheet.clear();
  mainSheet.getRange("B3:B").clear();
  mainSheet.getRange("F3:F").clear();
  outputSheet.setFrozenRows(1);

  // ヘッダーを追加
  const headers = ['building_name', 'work_name', 'worker_name', 'cleaning_id', 'room_name_common_area_name', 'status', 'work_date', 'prefecture'];
  outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

  // 結果を出力シートに書き込み
  if (rows.length > 0) {
    outputSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);
  }

  customIndexSortFilter();

  // C1セルに「読み込み完了」と表示
  mainSheet.getRange("B1").setValue("読み込み完了");
}
