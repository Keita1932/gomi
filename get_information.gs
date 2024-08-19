function get_information(){
  worker_name();
  getCommonAreaIds();
}

function worker_name() {
  // 読み込みたいプロジェクトの名前
  const project_id = "m2m-core";
  
  // 実行するクエリ
  const query_execute = "SELECT id,name FROM `m2m_users_prod.user` where activation_status = 2 ORDER BY name";
  
  // 実行するクエリの確認
  Logger.log(query_execute);
  
  // 出力先シート名
  const ss = SpreadsheetApp.openById("1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs");
  const output_sheet = ss.getSheetByName("information");  // 変更箇所
  
  output_sheet.setFrozenRows(1);
  
  const result = BigQuery.Jobs.query(
    {
      useLegacySql: false,
      query: query_execute,
    },
    project_id
  );
  
  const rows = result.rows.map(row => {
    return row.f.map(cell => cell.v);
  });
  
  Logger.log(rows);
  
  output_sheet
    .getRange(
      2, 
      1, 
      rows.length, 
      rows[0].length
    )
    .setValues(rows);
}

function getCommonAreaIds() {
  // スプレッドシートとシートの指定
  const spreadsheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';
  const sheet = SpreadsheetApp.openById(spreadsheetId).getSheetByName('information');
  
  // BigQuery クエリを作成
  const projectId = 'm2m-core'; // プロジェクトID
  const datasetId = 'su_wo'; // データセット名
  const tableId = 'room_list_operation'; // テーブル名
  
  let query = `
    SELECT building_name, commonarea_id
    FROM \`${projectId}.${datasetId}.${tableId}\`
  `;
  Logger.log('Query Created: ' + query);
  
  // BigQuery でクエリを実行
  const queryResults = BigQuery.Jobs.query({
    query: query,
    useLegacySql: false
  }, projectId);
  Logger.log('Query Executed. Job Complete: ' + queryResults.jobComplete);
  
  // クエリ結果を取得
  if (queryResults.jobComplete && queryResults.rows) {
    const results = queryResults.rows.map(row => {
      return row.f.map(cell => cell.v);
    });
    Logger.log('Number of Rows Retrieved: ' + results.length);

    // D列以降にデータを書き込む準備
    const outputValues = results.map(row => [row[0], row[1]]); // building_name と commonarea_id を取得
    
    sheet.getRange(3, 4, outputValues.length, 2).setValues(outputValues); // D列に building_name, E列に commonarea_id を書き込む
    Logger.log('All data written to Columns D and E.');
    
  } else {
    Logger.log('Query did not complete successfully or no rows returned.');
  }
}



