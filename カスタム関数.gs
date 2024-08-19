function customIndexSortFilter() {
  // 1. 指定されたシートからデータを取得
  const ss = SpreadsheetApp.openById("1YNmoXwDNvuNJ1aC1azTCQQ_Fy-YfFk4LNe35OlZDqfs");
  const mainSheet = ss.getSheetByName("main");
  const outputSheet = ss.getSheetByName("出力");
  const exclusionSheet = ss.getSheetByName("除外");

  var dataA = outputSheet.getRange("A:A").getValues(); // 列Aの全データを取得
  var dataC = outputSheet.getRange("C:C").getValues(); // 列Cの全データを取得
  var exclusionValues = exclusionSheet.getRange("C:C").getValues().flat(); // 除外シートのC列データをフラットな配列に変換

  // 2. mainシートのA3からのデータを取得
  var rangeValues = mainSheet.getRange(3, 1, mainSheet.getLastRow() - 2, 1).getValues(); // A3:Aのデータ

  // 3. 結果を格納する配列を初期化
  var result = [];

  // 4. 指定された範囲の各セルを処理
  for (var i = 0; i < rangeValues.length; i++) {
    var jValue = rangeValues[i][0]; // 処理中のセルの値を取得

    // 5. 空セルの処理
    if (jValue === "") {
      result.push([""]);
      continue;
    }

    // 6. 列Aの値と一致する列Cの値を抽出
    var filteredValues = [];
    for (var j = 0; j < dataA.length; j++) {
      if (dataA[j][0] == jValue) {
        filteredValues.push(dataC[j][0]);
      }
    }

    // 7. 抽出された値がなければ空の配列を追加
    if (filteredValues.length == 0) {
      result.push(["該当なし"]);
      continue;
    }

    // 8. 抽出された値の出現回数をカウント
    var counts = {};
    for (var k = 0; k < filteredValues.length; k++) {
      counts[filteredValues[k]] = (counts[filteredValues[k]] || 0) + 1;
    }

    // 9. 出現回数に基づいて降順にソート
    filteredValues.sort(function(a, b) {
      return counts[b] - counts[a];
    });

    // 10. 除外リストに含まれない値を探す
    var validValue = ""; // 有効な値を格納
    for (var m = 0; m < filteredValues.length; m++) {
      if (!exclusionValues.includes(filteredValues[m])) {
        validValue = filteredValues[m]; // 除外リストに含まれない最初の値を格納
        break;
      }
    }

    // 11. 有効な値が見つかった場合は結果に追加、見つからなかった場合は"該当なし"
    if (validValue !== "") {
      result.push([validValue]);
    } else {
      result.push(["該当なし"]);
    }
  }

  // 12. 結果を出力シートの適切な範囲に書き込む
  mainSheet.getRange(3, 2, result.length, 1).setValues(result); // B列に書き込む
}
