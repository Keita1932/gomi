function testSpreadsheetManager() {
  const sheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';
  const sheetManager = new SpreadsheetManager(sheetId);

  // Test: 書き込みテスト
  sheetManager.setValue('main', 'A1', '2024-08-23');
  Logger.log("SpreadsheetManager: A1セルに書き込みテスト完了");

  // Test: 読み取りテスト
  const value = sheetManager.getValues('main', 'A1');
  Logger.log("SpreadsheetManager: A1セルの値 = " + value);
}

function testBigQueryManager() {
  const bigQueryManager = new BigQueryManager("m2m-core");

  // テスト用のクエリ（サンプルデータ取得）
  const query = `
    SELECT
      building_name,
      work_name
    FROM
      \`m2m-core.su_wo.wo_cleaning_tour\`
    LIMIT 5
  `;

  const result = bigQueryManager.query(query);
  const rows = bigQueryManager.extractRows(result);

  Logger.log("BigQueryManager: クエリ結果 = " + JSON.stringify(rows));
}

function testSlackNotifier() {
  const slackNotifier = new SlackNotifier([
    "https://hooks.slack.com/services/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"  
  ]);

  // テストメッセージの送信
  slackNotifier.notify("これはテストメッセージです。");
  Logger.log("SlackNotifier: テストメッセージを送信しました。");
}


function testOperationManager() {
  const slackNotifier = new SlackNotifier([
    "https://hooks.slack.com/services/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"
  ]);

  try {
    const sheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';
    const sheetManager = new SpreadsheetManager(sheetId);

    // Operationの実行
    const operationManager = new OperationManager(sheetManager, ApiTokenManager, slackNotifier);
    operationManager.performOperations();

    let assignedBuildings = sheetManager.getValues("該当ありなし", "A2:A")
      .flat() // 二次元配列を一次元配列に変換
      .filter(value => value); // 空の値を除外

    let withoutAssignBuildings = sheetManager.getValues("該当ありなし", "C2:C")
      .flat() // 二次元配列を一次元配列に変換
      .filter(value => value); // 空の値を除外

    // 建物名をSlack通知用のフォーマットに変換
    let assignedBuildingsMessage = assignedBuildings.length > 0 ? `\n【該当ありの建物】\n${assignedBuildings.join("\n")}` : '';
    let withoutAssignBuildingsMessage = withoutAssignBuildings.length > 0 ? `\n【該当なしの建物】\n${withoutAssignBuildings.join("\n")}` : '';

    // 処理が成功したときの通知
    slackNotifier.notify("ゴミ庫ツアーを作成しました。" +
      "\n https://docs.google.com/spreadsheets/d/1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs/edit?gid=0#gid=0" +
      assignedBuildingsMessage +
      withoutAssignBuildingsMessage
    );
  } catch (error) {
    // エラーが発生したときの通知
    slackNotifier.notify(`エラーが発生しました: ${error.message}`);
  } finally {
    // 成功・失敗に関わらずトリガーを設定
    TriggerManager.resetTriggers();
    TriggerManager.createTriggerForTomorrow();
  }
}


function testBigQueryDataHandler() {
  const sheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';  // テスト用のスプレッドシートIDを指定
  const sheetManager = new SpreadsheetManager(sheetId);
  const bigQueryManager = new BigQueryManager("m2m-core");  // プロジェクトIDを指定

  const bigQueryDataHandler = new BigQueryDataHandler(sheetManager, bigQueryManager);

  // BigQueryからデータを取得して、Google Sheetsに書き込むテスト
  bigQueryDataHandler.fetchAndWriteData();

  Logger.log("BigQueryDataHandler: テスト実行完了");
}

function testTriggerManager() {
  // 1. トリガーのリセットをテスト
  Logger.log("----- トリガーのリセットをテスト開始 -----");
  TriggerManager.resetTriggers();
  var triggersAfterReset = ScriptApp.getProjectTriggers();
  if (triggersAfterReset.length === 0) {
    Logger.log("すべてのトリガーが正常に削除されました。");
  } else {
    Logger.log("トリガー削除に問題があります。残っているトリガー数: " + triggersAfterReset.length);
  }
  Logger.log("----- トリガーのリセットをテスト終了 -----");

  // 2. 翌日のトリガー作成をテスト
  Logger.log("----- トリガー作成をテスト開始 -----");
  TriggerManager.createTriggerForTomorrow();
  var triggersAfterCreation = ScriptApp.getProjectTriggers();
  if (triggersAfterCreation.length > 0) {
    Logger.log("トリガーが正常に作成されました。現在のトリガー数: " + triggersAfterCreation.length);
    triggersAfterCreation.forEach(trigger => {
      Logger.log("作成されたトリガー: " + trigger.getHandlerFunction() + ", ID: " + trigger.getUniqueId());
    });
  } else {
    Logger.log("トリガー作成に失敗しました。トリガーが存在しません。");
  }
  Logger.log("----- トリガー作成をテスト終了 -----");

  // // 3. 最後にトリガーを再度リセットしてクリーンアップ
  // TriggerManager.resetTriggers();
  // Logger.log("テスト終了後にトリガーをリセットしました。");
}


function runAllTests() {
  Logger.log("=== テスト開始 ===");
  
  // 各テストを順番に実行
  testSpreadsheetManager();
  testBigQueryManager();
  testSlackNotifier();
  testOperationManager();
  
  Logger.log("=== テスト終了 ===");
}
