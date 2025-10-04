function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('series:Dashboard_Data_Update')
    .addItem('Step1: Sales_Order_Update', 'transformCsvAndApplyOperations')
    // .addItem('Step2: Order_Picking更新', 'copyOnlyNewRowsWithKey')
    // .addItem('Step2: Integrated_Data_Update', 'createSummarySheet')
    .addItem('Step2: Integrated_Data_Update', 'createSummarySheetAndProcessData')
    .addItem('Step3: Prediction_Result_Update', 'generateOperationForecastTable')
    .addItem('Step4: Trouble_Report_Update', 'troubleReportConvertAndOutput')
    .addToUi();

  SpreadsheetApp.getUi()
    .createMenu('one-time:Reflect_Picker_Checker_Packer_Prize_list')
    // .addItem('Step1: Integrated_Data_Update', 'createSummarySheet')
    .addItem('Step1: Integrated_Data_Update', 'createSummarySheetAndProcessData')
    .addItem('Step2: Prediction_Result_Update', 'generateOperationForecastTable')
    .addToUi();
}

// ▼ 日付を yyyymmdd 形式に統一して返す関数（全体で共通利用）
function toDateKey(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = ('0' + (val.getMonth() + 1)).slice(-2);
    var d = ('0' + val.getDate()).slice(-2);
    return '' + y + m + d;
  }
  if (typeof val === "string") {
    var m1 = val.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})$/); // mm/dd/yyyy
    var m2 = val.match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/); // yyyy/mm/dd
    if (m1) {
      var y = m1[3];
      var m = ('0' + m1[1]).slice(-2);
      var d = ('0' + m1[2]).slice(-2);
      return '' + y + m + d;
    } else if (m2) {
      var y = m2[1];
      var m = ('0' + m2[2]).slice(-2);
      var d = ('0' + m2[3]).slice(-2);
      return '' + y + m + d;
    }
    var m3 = val.match(/^(\d{4})(\d{2})(\d{2})$/);
    if (m3) {
      return '' + m3[1] + m3[2] + m3[3];
    }
  }
  return String(val);
}

// ▼ カラム名→列番号インデックス（全体で共通利用）
function getIndexMap(header) {
  var idx = {};
  header.forEach((h, i) => idx[h] = i);
  return idx;
}

// ▼ 任意カラム名配列を使ってキー生成（toDateKey自動適用、全体で共通利用）
function createRowKeyWithColumns(row, idx, keyColumns) {
  return keyColumns.map(function(col) {
    if (col.toLowerCase().indexOf('date') >= 0) {
      return toDateKey(row[idx[col]]);
    }
    return row[idx[col]];
  }).join('|');
}

// ▼ その他ユーティリティ（createSummarySheetで使用）
function toSlashDateStr(str) {
  if (!str || str.length !== 8) return str;
  return str.substr(0,4) + '/' + str.substr(4,2) + '/' + str.substr(6,2);
}
function formatDateTime(val) {
  if (!val) return "";
  var d = val;
  if (!(val instanceof Date)) d = new Date(val);
  if (d && !isNaN(d.getTime())) {
    return formatDate(d, "yyyy/MM/dd HH:mm:ss");
  }
  return String(val);
}
function formatDate(date, fmt) {
  var d = date instanceof Date ? date : new Date(date);
  if (isNaN(d.getTime())) return "";
  if (!fmt) fmt = "yyyy/MM/dd";
  return fmt.replace("yyyy", d.getFullYear())
    .replace("MM", ('0' + (d.getMonth() + 1)).slice(-2))
    .replace("dd", ('0' + d.getDate()).slice(-2))
    .replace("HH", ('0' + d.getHours()).slice(-2))
    .replace("mm", ('0' + d.getMinutes()).slice(-2))
    .replace("ss", ('0' + d.getSeconds()).slice(-2));
}
function combineDateTimeStr(dateStr, timeStr) {
  var d = "";
  if (!dateStr || !timeStr) return "";
  var dateKey = toDateKey(dateStr);
  d = toSlashDateStr(dateKey);

  var h = "00", m = "00", s = "00";
  if (timeStr instanceof Date) {
    h = ('0' + timeStr.getHours()).slice(-2);
    m = ('0' + timeStr.getMinutes()).slice(-2);
    s = ('0' + timeStr.getSeconds()).slice(-2);
  } else if (typeof timeStr === "string") {
    var t = timeStr.trim();
    var tt = t.split(":");
    if (tt.length >= 2) {
      h = ('0' + tt[0]).slice(-2);
      m = ('0' + tt[1]).slice(-2);
      s = ('0' + (tt[2] ? tt[2] : "00")).slice(-2);
    }
  }
  return d + " " + h + ":" + m + ":" + s;
}
function pad(num, n) {
  return num.toString().padStart(n, '0');
}
function timeDiff(start, end) {
  if (!start || !end) return "";
  try {
    var t1 = typeof start === 'string' ? new Date('2000/01/01 ' + start) : start;
    var t2 = typeof end === 'string' ? new Date('2000/01/01 ' + end) : end;
    return (t2 - t1) / 60000;
  } catch(e) { return ""; }
}
function lastIndexOfHeader(headerRow, colName) {
  var idx = -1;
  for (var i = 0; i < headerRow.length; i++) {
    if (headerRow[i] === colName) idx = i;
  }
  return idx;
}

function uploadToFirestoreBatch() {
  const API_KEY = 'AIzaSyDZOvO0Qswnc-8D_OIXCk7YLR7woj0zHFA';
  const PROJECT_ID = 'linklusion-dashboard-5a255';
  const COLLECTION_NAME = '統合データ';

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName('Rvised_統合データ');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rows = data.slice(1);

  const BATCH_SIZE = 500;
  const totalRows = rows.length;
  let batchRequests = [];
  let url = `https://firestore.googleapis.com/v1/projects/${PROJECT_ID}/databases/(default)/documents:batchWrite?key=${API_KEY}`;

  for (let i = 0; i < totalRows; i++) {
    const row = rows[i];
    const docData = {};
    headers.forEach((h, j) => {
      docData[h] = row[j];
    });

    const docId = Utilities.getUuid();  // 自動生成のドキュメントID

    const docFields = Object.fromEntries(
      Object.entries(docData).map(([k, v]) => [k, { stringValue: String(v) }])
    );

    batchRequests.push({
      update: {
        name: `projects/${PROJECT_ID}/databases/(default)/documents/${COLLECTION_NAME}/${docId}`,
        fields: docFields
      }
    });

    if (batchRequests.length === BATCH_SIZE || i === totalRows - 1) {
      const options = {
        method: 'post',
        contentType: 'application/json',
        payload: JSON.stringify({ writes: batchRequests }),
        muteHttpExceptions: true
      };

      UrlFetchApp.fetch(url, options);
      batchRequests = [];  // バッチをクリアして次のグループへ
      Utilities.sleep(500);  // Firestoreへの過負荷防止で0.5秒待つ（適宜調整）
    }
  }
}

// ▼ 指定スプレッドシートへ出力（指定キー列で重複するのは上書き更新）
// sheetName は省略可能。渡された場合はそのシートを使用。
function outputToExternalSheetWithHeaderAndUpsert(sheetId, targetHeaders, newData, keyColumns, sheetName) {
  const externalSS = SpreadsheetApp.openById(sheetId);
  const sheet = sheetName
    ? externalSS.getSheetByName(sheetName)
    : externalSS.getSheets()[0];

  if (!sheet) {
    throw new Error(`指定されたシートが存在しません: ${sheetName || '(1枚目のシート)'}`);
  }

  // 既存ヘッダーの状態を「上書きせずに」判定
  const headerRowValues = sheet.getRange(1, 1, 1, targetHeaders.length).getValues()[0];
  const isHeaderEmpty = headerRowValues.every(cell => cell === "" || cell === null);

  let lastRow = sheet.getLastRow();

  // ヘッダーが空の時だけ targetHeaders をセット
  if (isHeaderEmpty) {
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    lastRow = 1;
  }

  // 既存データ取得
  const existingData = (lastRow > 1)
    ? sheet.getRange(2, 1, lastRow - 1, targetHeaders.length).getValues()
    : [];

  // キー作成
  const makeKey = row => keyColumns.map(idx => row[idx]).join('__');

  // 既存のキー → 行番号マップ
  const existingMap = new Map();
  for (let i = 0; i < existingData.length; i++) {
    const key = makeKey(existingData[i]);
    existingMap.set(key, 2 + i);
  }

  // newDataを先頭から処理
  const seenNewKeysForAppend = new Set();
  const seenKeysForUpdate = new Set();
  const updates = [];
  const appends = [];

  for (let i = 0; i < newData.length; i++) {
    const row = newData[i];
    const key = makeKey(row);

    if (existingMap.has(key)) {
      if (!seenKeysForUpdate.has(key)) {
        updates.push({ rIndex: existingMap.get(key), row });
        seenKeysForUpdate.add(key);
      }
    } else {
      if (!seenNewKeysForAppend.has(key)) {
        appends.push(row);
        seenNewKeysForAppend.add(key);
      }
    }
  }

  // 更新
  updates
    .sort((a, b) => a.rIndex - b.rIndex)
    .forEach(u => sheet.getRange(u.rIndex, 1, 1, targetHeaders.length).setValues([u.row]));

  // 追記
  if (appends.length > 0) {
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, appends.length, targetHeaders.length).setValues(appends);
  }
}

// ▼ 指定スプレッドシートへ出力（指定キー列で重複排除して追記）
function outputToExternalSheetWithHeaderAndDedup(sheetId, targetHeaders, newData, keyColumns) {
  const externalSS = SpreadsheetApp.openById(sheetId);
  const sheet = externalSS.getSheets()[0];

  const headerRow = sheet.getRange(1, 1, 1, targetHeaders.length).getValues()[0];
  const isHeaderEmpty = headerRow.every(cell => cell === "" || cell === null);

  let lastRow = sheet.getLastRow();

  if (isHeaderEmpty) {
    sheet.getRange(1, 1, 1, targetHeaders.length).setValues([targetHeaders]);
    lastRow = 1;
  }

  // ヘッダーがあるならデータ取得、ないなら空配列
  const existingData = (lastRow > 1)
    ? sheet.getRange(2, 1, lastRow - 1, targetHeaders.length).getValues()
    : [];

  const existingKeySet = new Set(existingData.map(row =>
    keyColumns.map(idx => row[idx]).join('__')
  ));

  const filteredNewData = newData.filter(row => {
    const key = keyColumns.map(idx => row[idx]).join('__');
    if (existingKeySet.has(key)) {
      return false;
    }
    existingKeySet.add(key);
    return true;
  });

  if (filteredNewData.length > 0) {
    sheet.getRange(lastRow + 1, 1, filteredNewData.length, targetHeaders.length).setValues(filteredNewData);
  }
}





