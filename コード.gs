function moveAllFunctions() {
  const slackNotifier = new SlackNotifier([
    "https://hooks.slack.com/services/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"
  ]);

  try {
    const sheetId = '1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs';
    const sheetManager = new SpreadsheetManager(sheetId);
    const bigQueryManager = new BigQueryManager("m2m-core");

    // 日付の設定
    sheetManager.setValue('main', 'A1', DateManager.getTomorrowFormatted());

    // BigQueryデータの処理とフィルタリング
    const bigQueryDataHandler = new BigQueryDataHandler(sheetManager, bigQueryManager);
    bigQueryDataHandler.fetchAndWriteData();

    // Operationの実行
    const operationManager = new OperationManager(sheetManager, ApiTokenManager, slackNotifier);
    operationManager.performOperations();

    // 処理が成功したときの通知
    slackNotifier.notify("ゴミ庫ツアーを作成しました。");
  } catch (error) {
    // エラーが発生したときの通知
    slackNotifier.notify(`エラーが発生しました: ${error.message}`);
  } finally {
    // 成功・失敗に関わらずトリガーを設定
    TriggerManager.resetTriggers();
    TriggerManager.createTriggerForTomorrow();
  }
}


class SpreadsheetManager {
  constructor(sheetId) {
    this.sheet = SpreadsheetApp.openById(sheetId);
  }

  getSheetByName(sheetName) {
    return this.sheet.getSheetByName(sheetName);
  }

  setValue(sheetName, range, value) {
    this.getSheetByName(sheetName).getRange(range).setValue(value);
  }

  clearRange(sheetName, range) {
    this.getSheetByName(sheetName).getRange(range).clearContent();
  }

  getValues(sheetName, range) {
    return this.getSheetByName(sheetName).getRange(range).getValues();
  }
}

class BigQueryManager {
  constructor(projectId) {
    this.projectId = projectId;
  }

  query(query) {
    return BigQuery.Jobs.query({
      useLegacySql: false,
      query: query,
    }, this.projectId);
  }

  extractRows(queryResult) {
    return queryResult.rows ? queryResult.rows.map(row => row.f.map(cell => cell.v)) : [];
  }
}

class SlackNotifier {
  constructor(webhookUrls) {
    this.webhookUrls = webhookUrls;
  }

  notify(errorMessage) {
    const payload = {
      "username": "トラブル報告",
      "text":  errorMessage,
      "icon_emoji": ":warning:"
    };
    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload)
    };
    this.webhookUrls.forEach(url => UrlFetchApp.fetch(url, options));
  }
}

class TriggerManager {
  static resetTriggers() {
    ScriptApp.getProjectTriggers().forEach(trigger => ScriptApp.deleteTrigger(trigger));
  }

  static createTriggerForTomorrow() {
    const now = new Date();
    const tomorrow = new Date(now.getFullYear(), now.getMonth(), now.getDate() + 1, 20, 5, 0);
    ScriptApp.newTrigger("moveAllFunctions")
      .timeBased()
      .at(tomorrow)
      .create();
  }
}

class DateManager {
  static getTomorrowFormatted() {
    const today = new Date();
    const tomorrow = new Date(today);
    tomorrow.setDate(today.getDate() + 1);
    const timeZone = "Asia/Tokyo";
    return Utilities.formatDate(tomorrow, timeZone, 'yyyy-MM-dd');
  }

  static convertDate(dateString) {
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
    return `${year}-${month}-${day}`;
  }
}

class ApiTokenManager {
  static getToken() {
    const url = 'https://api.m2msystems.cloud/login';
    const mail = PropertiesService.getScriptProperties().getProperty("mail_address");
    const pass = PropertiesService.getScriptProperties().getProperty("pass");

    const payload = {
      email: mail,
      password: pass
    };
    
    const options = {
      method: 'post',
      contentType: 'application/json',
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };
    
    const response = UrlFetchApp.fetch(url, options);
    if (response.getResponseCode() === 200) {
      const json = JSON.parse(response.getContentText());
      return json.accessToken;
    } else {
      Logger.log("トークン取得に失敗しました。ステータスコード: " + response.getResponseCode());
      return null;
    }
  }
}

class OperationManager {
  constructor(sheetManager, tokenManager, slackNotifier) {
    this.sheetManager = sheetManager;
    this.tokenManager = tokenManager;
    this.slackNotifier = slackNotifier;
  }

  performOperations() {
    const operationSheetName = 'main';
    this.sheetManager.clearRange(operationSheetName, "F3:F");

    const originCleaningDate = this.sheetManager.getValues(operationSheetName, "A1")[0][0];
    const cleaningDate = DateManager.convertDate(originCleaningDate);
    const nameColumn = this.sheetManager.getValues(operationSheetName, "B:B");

    let lastRowWithData = 1;
    for (let i = 0; i < nameColumn.length; i++) {
      if (nameColumn[i][0]) {
        lastRowWithData = i + 1;
      }
    }

    const operationData = this.sheetManager.getValues(operationSheetName, `A3:E${lastRowWithData}`);
    const token = this.tokenManager.getToken();
    const tokenFailed = !token;

    const errorMessages = [];
    const errorProperties = []; // 失敗した物件名を収集する配列
    
    operationData.forEach((rowData, index) => {
      const propertyName = rowData[0];  // A列が物件名

      let payload;

      if(rowData[2] !== ''){
         payload= {
          "placement": "commonArea",
          "commonAreaId": rowData[3],
          "listingId": "",
          "cleaningDate": cleaningDate,
          "note": "",
          "cleaners": [rowData[2]],
          "photoTourId": rowData[4],
        };
      }else{
         payload= {
          "placement": "commonArea",
          "commonAreaId": rowData[3],
          "listingId": "",
          "cleaningDate": cleaningDate,
          "note": "",
          "cleaners": [],
          "photoTourId": rowData[4],
        };
      }


      let response;
      let result;

      if (tokenFailed) {
        response = "トークン取得失敗のためリクエスト未送信";
        result = "error";
      } else {
        try {
          response = UrlFetchApp.fetch('https://api-cleaning.m2msystems.cloud/v3/cleanings/create_with_placement', {
            'method': 'post',
            'contentType': 'application/json',
            'headers': { 'Authorization': 'Bearer ' + token },
            'payload': JSON.stringify(payload),
            'muteHttpExceptions': true
          }).getContentText();
          result = response.includes('error') ? 'error' : 'ok';
        } catch (e) {
          response = "リクエストエラー: " + e.message;
          result = "error";
        }
      }

      if (result !== 'ok') {
        errorMessages.push(`Error for row ${index + 3}: ${response}`);
        errorProperties.push(propertyName); // 失敗した物件名を追加
      }

      this.sheetManager.setValue(operationSheetName, `F${index + 3}`, result);
    });

    if (errorMessages.length > 0) {
      const errorPropertiesMessage = errorProperties.length > 0
        ? `【失敗した物件名】\n${errorProperties.join("\n")}` // 改行で区切る
        : "";
      this.slackNotifier.notify(`ツアー作成に失敗しました。\n${errorPropertiesMessage}\n\n https://docs.google.com/spreadsheets/d/1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs/edit?gid=0#gid=0`);
    }
  }
}


class BigQueryDataHandler {
  constructor(sheetManager, bigQueryManager) {
    this.sheetManager = sheetManager;
    this.bigQueryManager = bigQueryManager;
  }

  fetchAndWriteData() {
    const mainSheet = this.sheetManager.getSheetByName("main");
    const outputSheet = this.sheetManager.getSheetByName("出力");

    const a1Date = mainSheet.getRange("A1").getValue();
    if (!a1Date || !(a1Date instanceof Date)) {
      Logger.log("A1セルには有効な日付が含まれていません。");
      mainSheet.getRange("A1").setValue("エラー");
      return;
    }

    const timeZone = "Asia/Tokyo";
    const workDate = Utilities.formatDate(a1Date, timeZone, 'yyyy-MM-dd');

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
        AND prefecture IN ('東京都', '大阪府')
    `;

    Logger.log(query_execute);

    const result = this.bigQueryManager.query(query_execute);
    const rows = this.bigQueryManager.extractRows(result);

    if (rows.length > 0) {
      outputSheet.clear();
      mainSheet.getRange("B3:B").clear();
      mainSheet.getRange("F3:F").clear();
      outputSheet.setFrozenRows(1);

      const headers = ['building_name', 'work_name', 'worker_name', 'cleaning_id', 'room_name_common_area_name', 'status', 'work_date', 'prefecture'];
      outputSheet.getRange(1, 1, 1, headers.length).setValues([headers]);

      outputSheet.getRange(2, 1, rows.length, rows[0].length).setValues(rows);

      this.customIndexSortFilter(mainSheet, outputSheet);
      mainSheet.getRange("B1").setValue("一覧読み込み完了");
    } else {
      Logger.log("BigQueryからデータが返されませんでした");
      mainSheet.getRange("B1").setValue("エラー");
    }
  }

  customIndexSortFilter(mainSheet, outputSheet) {
    const exclusionSheet = this.sheetManager.getSheetByName("除外");

    var dataA = outputSheet.getRange("A:A").getValues();
    var dataC = outputSheet.getRange("C:C").getValues();
    var exclusionValues = exclusionSheet.getRange("C:C").getValues().flat();

    var rangeValues = mainSheet.getRange(3, 1, mainSheet.getLastRow() - 2, 1).getValues();

    var result = [];

    for (var i = 0; i < rangeValues.length; i++) {
      var jValue = rangeValues[i][0];

      if (jValue === "") {
        result.push([""]);
        continue;
      }

      var filteredValues = [];
      for (var j = 0; j < dataA.length; j++) {
        if (dataA[j][0] == jValue) {
          filteredValues.push(dataC[j][0]);
        }
      }

      if (filteredValues.length == 0) {
        result.push(["該当なし"]);
        continue;
      }

      var counts = {};
      for (var k = 0; k < filteredValues.length; k++) {
        counts[filteredValues[k]] = (counts[filteredValues[k]] || 0) + 1;
      }

      filteredValues.sort(function(a, b) {
        return counts[b] - counts[a];
      });

      var validValue = "";
      for (var m = 0; m < filteredValues.length; m++) {
        if (!exclusionValues.includes(filteredValues[m])) {
          validValue = filteredValues[m];
          break;
        }
      }

      if (validValue !== "") {
        result.push([validValue]);
      } else {
        result.push(["該当なし"]);
      }
    }

    mainSheet.getRange(3, 2, result.length, 1).setValues(result);
  }
}

