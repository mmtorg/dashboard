function readXlsxAsCsvData(sheetName, xlsxFileNamePartial) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const myDriveRoot = DriveApp.getRootFolder();

  // 現在のスプレッドシートと同じ階層のフォルダを取得
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

  // 「input_data」フォルダを取得
  const inputFolders = parentFolder.getFoldersByName('input_data');
  if (!inputFolders.hasNext()) throw new Error('input_data フォルダが見つかりません');
  const inputFolder = inputFolders.next();

  // 「input_data」フォルダ内の対象ファイル取得（部分一致＋.xlsx限定）
  const filesIter = inputFolder.getFiles();
  let xlsxFile = null;
  while (filesIter.hasNext()) {
    const file = filesIter.next();
    const fileName = file.getName();
    if (fileName.includes(xlsxFileNamePartial) && fileName.toLowerCase().endsWith(".xlsx")) {
      xlsxFile = file;
      break; // 最初に見つかったファイルを使用
    }
  }
  if (!xlsxFile) throw new Error(`${xlsxFileNamePartial} を含む .xlsx ファイルが input_data フォルダ内に見つかりません`);

  // 一時ファイルとしてコピー
  const copiedFile = xlsxFile.makeCopy("temp_convert.xlsx", myDriveRoot);

  // Googleスプレッドシートに変換
  const resource = {
    name: 'temp_convert',
    mimeType: 'application/vnd.google-apps.spreadsheet'
  };
  const convertedFile = Drive.Files.copy(resource, copiedFile.getId());
  const sheetFileId = convertedFile.id;

  const tempSS = SpreadsheetApp.openById(sheetFileId);
  const sheet = tempSS.getSheetByName(sheetName);
  if (!sheet) throw new Error(`${sheetName} シートが見つかりません`);

  const data = sheet.getDataRange().getValues();

  // 一時ファイル削除
  DriveApp.getFileById(sheetFileId).setTrashed(true);
  copiedFile.setTrashed(true);

  return data; // CSV相当の2次元配列
}

// // ▼ createSummarySheet（省略せずフル記載）
// function createSummarySheet() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();

//   // sales_order取得
//   const salesSheet = ss.getSheetByName('Sales_Order');
//   const salesData = salesSheet.getDataRange().getValues();

//   // Pick/Pack取得
//   const pickingSpreadsheet = SpreadsheetApp.openById('1Vp08XkCklvrF6DfZEPDzOBuLl6DclY3qUaua6nE7cEw');
//   const pickingSheet = pickingSpreadsheet.getSheetByName('Pick/Pack');
//   const pickingData = pickingSheet.getDataRange().getValues();

//   // const pickingSheet = ss.getSheetByName('Order_Picking');
//   // const pickingData = pickingSheet.getDataRange().getValues();

//   let outSheet = ss.getSheetByName('Integrated_Data');
//   if (!outSheet) outSheet = ss.insertSheet('Integrated_Data');

//   // 部分一致にしたので第2引数は「CartonCount_Customer」というキーワードだけでOK
//   const cartonData = readXlsxAsCsvData("CartonCount_Customer", "CartonCount_Customer");

//   // ▼ Carton_Count.csv を Drive から読み込む
//   // const folder = DriveApp.getFileById(ss.getId()).getParents().next();
//   // const files = folder.getFilesByName("CartonCount_Customer.csv");  // ←ファイル名は要確認
//   // if (!files.hasNext()) throw new Error("CartonCount_Customer.csv が見つかりません");
//   // const csvBlob = files.next().getBlob();
//   // const cartonData = Utilities.parseCsv(csvBlob.getDataAsString("utf-8"));

//   // const cartonSheet = ss.getSheetByName('Carton_Count');
//   // const cartonData = cartonSheet ? cartonSheet.getDataRange().getValues() : [];

//   // ヘッダーインデックス作成
//   const salesIdx = getIndexMap(salesData[0]);
//   const pickingIdx = getIndexMap(pickingData[0]);
//   const cartonIdx = cartonData.length > 0 ? getIndexMap(cartonData[0]) : {};

//   // ピッカー・パッカー・チェッカーの後者の列インデックス取得
//   const idxPicker2  = lastIndexOfHeader(pickingData[0], 'Picker');
//   const idxPacker2  = lastIndexOfHeader(pickingData[0], 'Packer');
//   const idxChecker2 = lastIndexOfHeader(pickingData[0], 'Double Checker');

//   // Sales_Order：事前に集計キーを計算して格納
//   let summaryMap = {};
//   for (let i = 1, len = salesData.length; i < len; i++) {
//     const row = salesData[i];
//     const dateKey = toDateKey(row[salesIdx['Order Date']]);
//     const customer = row[salesIdx['Customer']];
//     const key = dateKey + '|' + customer;

//     if (!summaryMap[key]) {
//       summaryMap[key] = {
//         Order_Date: row[salesIdx['Order Date']],
//         date_only: toSlashDateStr(dateKey),
//         Customer: customer,
//         SKU: 0,
//         qty: 0,
//         Ks: 0,
//         Depot: row[salesIdx['Sales Team']],
//         Course: row[salesIdx['Salesperson']],
//         order_time_min: row[salesIdx['Order Date']] && row[salesIdx['Creation Date']] ? timeDiff(row[salesIdx['Creation Date']], row[salesIdx['Order Date']]) : "",
//         Staff_Receive: row[salesIdx['App User/Login']],
//         order_actual_end_time: formatDateTime(row[salesIdx['Order Date']]),
//         order_pred_min: "",
//         pick_pred_min: "",
//         pack_pred_min: "",
//         pick_time_min: "",
//         pack_time_min: "",
//         Staff_Pick: "",
//         Staff_Pack: "",
//         Staff_Check: "",
//         pick_actual_end_time: "",
//         pack_actual_end_time: "",
//         Carton: "",
//         Beverage: ""
//       };
//     }
//     summaryMap[key].SKU += 1;
//     summaryMap[key].qty += Number(row[salesIdx['Order Lines/Quantity']]) || 0;
//     summaryMap[key].Ks  += Number(row[salesIdx['Total']]) || 0;
//   }

//   // Order_Picking：キーを事前に作成してから1回で照合
//   for (let i = 1, len = pickingData.length; i < len; i++) {
//     const row = pickingData[i];
//     if (row[pickingIdx['DC/Depot']] && typeof row[pickingIdx['DC/Depot']] === "string") {
//       row[pickingIdx['DC/Depot']] = row[pickingIdx['DC/Depot']].replace(/Depot-/g, "");
//     }
//     let shopNoRaw = row[pickingIdx['Shop No.']];
//     let shopNo = (typeof shopNoRaw === "number" || (/^\d+$/.test(String(shopNoRaw)))) ? "KMU-" + shopNoRaw : shopNoRaw;
//     const dateKey = toDateKey(row[pickingIdx['Date']]);
//     const key = dateKey + '|' + shopNo;

//     let summary = summaryMap[key];
//     if (!summary) continue;
//     summary.pick_time_min  = timeDiff(row[pickingIdx['Pick Start time']], row[pickingIdx['Pick Finish time']]);
//     summary.pack_time_min  = timeDiff(row[pickingIdx['Pack Start time']], row[pickingIdx['Pack Finish time']]);
//     summary.Staff_Pick     = row[idxPicker2] || "";
//     summary.Staff_Pack     = row[idxPacker2] || "";
//     summary.Staff_Check    = row[idxChecker2] || "";
//     summary.pick_actual_end_time = combineDateTimeStr(row[pickingIdx['Date']], row[pickingIdx['Pick Finish time']]);
//     summary.pack_actual_end_time = combineDateTimeStr(row[pickingIdx['Date']], row[pickingIdx['Pack Finish time']]);
//   }

//   // Carton_Count
//   if (cartonData.length > 1) {
//     for (let i = 1, len = cartonData.length; i < len; i++) {
//       const row = cartonData[i];
//       const dateKey = toDateKey(row[cartonIdx['Date']]);
//       const customer = row[cartonIdx['Customer']];
//       const key = dateKey + '|' + customer;
//       let summary = summaryMap[key];
//       if (!summary) continue;
//       summary.Carton   = row[cartonIdx['Carton']];
//       summary.Beverage = row[cartonIdx['Drink']];
//     }
//   }

//   // 予測値計算
//   const summaries = Object.values(summaryMap);
//   for (let i = 0, len = summaries.length; i < len; i++) {
//     let obj = summaries[i];
//     obj.Carton = Number(obj.Carton) || 0;
//     obj.order_pred_min = (0.9612 + 0.3073 * obj.SKU + 0.0065 * obj.qty).toFixed(4);
//     obj.pick_pred_min  = (0.2040 + 0.2198 * obj.SKU + 0.0025 * obj.qty + 0.4495 * obj.Carton).toFixed(4);
//     obj.pack_pred_min  = (-0.9712 + 0.1913 * obj.SKU + 0.0011 * obj.qty + 0.5966 * obj.Carton).toFixed(4);
//   }

//   // 出力
//   const headers = [
//     'Order_Date','date_only', 'Customer', 'SKU', 'qty', 'Ks', 'Depot', 'Course',
//     'pick_time_min', 'pack_time_min', 'order_time_min', 'Staff_Receive',
//     'Staff_Pick', 'Staff_Pack', 'Staff_Check',
//     'order_pred_min', 'pick_pred_min', 'pack_pred_min',
//     'order_actual_end_time', 'pick_actual_end_time', 'pack_actual_end_time',
//     'Carton', 'Beverage'
//   ];
//   const outData = [headers];
//   for (let i = 0, len = summaries.length; i < len; i++) {
//     const obj = summaries[i];
//     var row = [];
//     for (var j = 0; j < headers.length; j++) {
//       var h = headers[j];
//       row.push(obj[h] !== undefined ? obj[h] : "");
//     }
//     outData.push(row);
//   }

//   // outSheet.clear();
//   // outSheet.getRange(1, 1, outData.length, headers.length).setValues(outData);

//   // ▼▼▼ Integrated_Data への書き込みロジック（置き換え） ▼▼▼
//   const ui = SpreadsheetApp.getUi();

//   // ヘッダーを保証（未作成や空シート時のみセット）
//   const existingRange = outSheet.getDataRange();
//   const existingValues = existingRange.getValues();
//   const normalizeHeader = arr => arr.map(v => String(v).trim()).join();
//   const headerMatches = (existingValues.length > 0) &&
//     normalizeHeader(existingValues[0]) === normalizeHeader(headers);

//   if (!headerMatches) {
//     outSheet.clear();
//     outSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
//   }

//   // 列位置（1-based）
//   const dateOnlyCol   = headers.indexOf('date_only') + 1;
//   const pickTimeCol   = headers.indexOf('pick_time_min') + 1;
//   const packTimeCol   = headers.indexOf('pack_time_min') + 1;

//   // 今日0:00 と基準日
//   const today0 = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate());
//   const ydayCutoff = today0; // 「前日を含む前日より前」= 今日より前（< today0）
//   const sevenDaysCut = new Date(today0);
//   sevenDaysCut.setDate(sevenDaysCut.getDate() - 7);

//   // 既存データの走査（2行目以降）
//   const lastRow = outSheet.getLastRow();
//   let rowsToClear = [];
//   const alertDatesSet = new Set();

//   // 表示値として取得（date_only は "yyyy/MM/dd" 想定）
//   const getDisp = (r, c) => outSheet.getRange(r, c).getDisplayValue();
//   const parseYmd = (s) => {
//     if (!s) return null;
//     const m = String(s).match(/^(\d{4})[\/\-](\d{1,2})[\/\-](\d{1,2})$/);
//     if (!m) return null;
//     return new Date(Number(m[1]), Number(m[2]) - 1, Number(m[3]));
//   };
//   const toJaMd = (d) => `${d.getMonth() + 1}月${d.getDate()}日`;

//   if (lastRow > 1 && dateOnlyCol > 0 && pickTimeCol > 0 && packTimeCol > 0) {
//     for (let r = 2; r <= lastRow; r++) {
//       const dStr = getDisp(r, dateOnlyCol);
//       const d = parseYmd(dStr);
//       if (!d) continue;

//       const hasPick = getDisp(r, pickTimeCol) !== '';
//       const hasPack = getDisp(r, packTimeCol) !== '';
//       const hasEither = hasPick || hasPack;

//       if (hasEither && d.getTime() < sevenDaysCut.getTime()) {
//         alertDatesSet.add(toJaMd(d));
//       }
//       if (hasEither && d.getTime() < ydayCutoff.getTime()) {
//         rowsToClear.push(r);
//       }
//     }
//   }

//   clearOldDataAndEmptyRows(outSheet);
//     
//     if (outData.length > 1) {
//       const targetHeaders = headers;
//       const newData = outData.slice(1);
//       outputToExternalSheetWithHeaderAndUpsert(
//         ss.getId(),
//         targetHeaders,
//         newData,
//         [0, 2],
//         'Integrated_Data'
//       );
//     }

//   // // ▼ 値クリア（昨日以前＆pick/pack に値ありの行）
//   // if (rowsToClear.length > 0) {
//   //   rowsToClear.forEach(r => {
//   //     outSheet.getRange(r, 1, 1, headers.length).clearContent();
//   //   });
//   //   // 空白行の削除（2行目以降）
//   //   for (let r = outSheet.getLastRow(); r >= 2; r--) {
//   //     const rowValues = outSheet.getRange(r, 1, 1, headers.length).getValues()[0];
//   //     const isEmpty = rowValues.every(v => v === "" || v === null);
//   //     if (isEmpty) outSheet.deleteRow(r);
//   //   }
//   // }

//   if (outData.length > 1) {
//       const targetHeaders = headers;        // このシートのヘッダーをそのまま使用
//       const newData = outData.slice(1);     // ヘッダー行を除いたデータ
//       // 同一スプレッドシート内の 'Integrated_Data' シートを明示してアップサート
//       outputToExternalSheetWithHeaderAndUpsert(
//         ss.getId(),           // 現在のブック
//         targetHeaders,        // ヘッダー
//         newData,              // 追加/更新データ
//         [0, 2],               // キー列（Order_Date=0, Customer=2）
//         'Integrated_Data'     // シート名（オプショナル対応版）
//       );
//   }

//   // 処理完了後にアラート表示（該当日付がある場合のみ）
//   if (alertDatesSet.size > 0) {
//     const lines = Array.from(alertDatesSet).sort();
//     ui.alert(
//       `以下の日付は1週間以上前ですがのPicker Checker Packer Prize listが反映されていません。\n` +
//       `Picker Checker Packer Prize listを確認して、Reflect_Picker_Checker_Packer_Prize_listボタンを押して反映してください。\n` +
//       lines.map(md => `${md}`).join('\n')
//     );
//   }
//   // ▲▲▲ 置き換えここまで ▲▲▲

//   // Revised_Sales_Orderシート対応
//   // // 1. Revised_Sales_Orderシートの2行目以降のみクリア
//   // let revisedSheet = ss.getSheetByName('Revised_Sales_Order');
//   // if (revisedSheet) {
//   //   if (revisedSheet.getLastRow() > 1) {
//   //     revisedSheet.getRange(2, 1, revisedSheet.getLastRow() - 1, revisedSheet.getLastColumn()).clearContent();
//   //   }
//   // } else {
//   //   throw new Error('Revised_Sales_Orderシートが存在しません');
//   // }

//   // Sales_Orderシートの'Order Date', 'Creation Date', 'App User/Login'列だけをコピー
//   const allowedHeaders = ['Order Date', 'Creation Date', 'App User/Login'];
//   const header = salesData[0];
//   const colIdxs = allowedHeaders.map(h => header.indexOf(h));

//   // ヘッダー（1行目）のみ
//   // const filteredHeader = colIdxs.map(idx => header[idx]);
//   // revisedSheet.getRange(1, 1, 1, filteredHeader.length).setValues([filteredHeader]);

//   // 2行目以降（重複削除・空白除外処理を追加）
//   const uniqueSet = new Set();
//   const filteredData = [];

//   salesData.slice(1).forEach(row => {
//     const values = colIdxs.map(idx => row[idx]);
//     const staff = values[2]; // App User/Login

//     // スタッフ名が空の行は除外
//     if (!staff || staff.toString().trim() === '') return;

//     // 「スタッフ＋開始＋終了」のセットでユニーク化
//     const key = values.join('__');
//     if (!uniqueSet.has(key)) {
//       uniqueSet.add(key);
//       filteredData.push(values);
//     }
//   });

//   // // 2行目以降にユニーク行のみ貼り付け
//   // if (filteredData.length > 0) {
//   //   revisedSheet.getRange(2, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
//   // }

//   // Revised_Sales_Orderシートの内容を別スプレッドシートへ出力
//   // 【DB】sales_order
//   // const revisedOutputSheetId = '13mRHbl1mDoVXchUUUlLX0y38VM7bAI512uvJ4iSx_us';

//   // // 保存先スプレッドシートとシート取得（先頭シート想定）
//   // const revisedSS = SpreadsheetApp.openById(revisedOutputSheetId);
//   // const revisedSheet = revisedSS.getSheets()[0];  // 1枚目のシート

//   // // ヘッダーが未設定なら追加
//   // const existingHeader = revisedSheet.getRange(1, 1, 1, filteredHeader.length).getValues()[0];
//   // const isHeaderEmpty = existingHeader.every(cell => cell === "" || cell === null);

//   // if (isHeaderEmpty) {
//   //   revisedSheet.getRange(1, 1, 1, filteredHeader.length).setValues([filteredHeader]);
//   // }

//   // // 最終行の次の行にデータ追記
//   // const lastRow = revisedSheet.getLastRow();
//   // const writeStartRow = lastRow + 1;

//   // if (filteredData.length > 0) {
//   //   revisedSheet.getRange(writeStartRow, 1, filteredData.length, filteredData[0].length).setValues(filteredData);
//   // }

//   // ▼ 追加：オペレーション予測用カラムを除外して別シートに複製（シートは存在前提・軽量処理）
//   // const excludeColumns = [
//   //   "order_pred_min",
//   //   "pick_pred_min",
//   //   "pack_pred_min",
//   //   "order_actual_end_time",
//   //   "pick_actual_end_time",
//   //   "pack_actual_end_time"
//   // ];

//   // const allData = outSheet.getDataRange().getValues();
//   // const headerRow = allData[0];

//   // // 除外列インデックスを事前に取得（O(N)）
//   // const excludeIndexes = excludeColumns.map(col => headerRow.indexOf(col)).filter(idx => idx > -1);

//   // // ヘッダー＋データ行から除外列を削除（O(N*M)）
//   // const filteredOutData = allData.map(row =>
//   //   row.filter((_, colIndex) => !excludeIndexes.includes(colIndex))
//   // );

//   // ▼ integrated_data_main 出力
//   const sheetId_1 = '1Jot9ZgygJN3FukLWZtd7ZuSoz1ybzhvgMjff_UR7OpU';
//   const targetHeaders_1 = [
//     'Order_Date', 'date_only', 'Customer', 'SKU', 'qty', 'Ks', 'Depot', 'Course', 'Carton', 'Beverage', 'pick_time_min', 'pack_time_min', 'order_time_min'
//   ];
//   const newData_1 = summaries.map(obj =>
//     targetHeaders_1.map(h => obj[h] !== undefined ? obj[h] : "")
//   );
//   // 重複するデータは削除ではなく上書き更新に変更
//   outputToExternalSheetWithHeaderAndUpsert(sheetId_1, targetHeaders_1, newData_1, [0, 2]);  // Order_Date + Customer
//   // outputToExternalSheetWithHeaderAndDedup(sheetId_1, targetHeaders_1, newData_1, [0, 2]);  // Order_Date + Customer

//   // ▼ integrated_data_sub 出力
//   const sheetId_2 = '1mJ9z4wJ2AfCD5fkr_ocWRo0KUEbcH4R3J2YgQedqOEE';
//   const targetHeaders_2 = [
//     'Order_Date', 'date_only',
//     'Staff_Receive', 'Staff_Pick', 'Staff_Pack', 'Staff_Check'
//   ];
//   const newData_2 = summaries.map(obj =>
//     targetHeaders_2.map(h => obj[h] !== undefined ? obj[h] : "")
//   );
//   // 重複するデータは削除ではなく上書き更新に変更
//   outputToExternalSheetWithHeaderAndUpsert(sheetId_2, targetHeaders_2, newData_2, [0, 4, 5]);  // Order_Date + order_time_min + Staff_Receive
//   // outputToExternalSheetWithHeaderAndDedup(sheetId_2, targetHeaders_2, newData_2, [0, 4, 5]);  // Order_Date + order_time_min + Staff_Receive

//   // ▼ sales_order 出力
//   const revisedOutputSheetId = '13mRHbl1mDoVXchUUUlLX0y38VM7bAI512uvJ4iSx_us';
//   const filteredHeader = colIdxs.map(idx => header[idx]);
//   const revisedNewData = filteredData;  // 重複除去済みのデータ
//   // sales_orderはodooデータで変わらないので重複はスキップ
//   outputToExternalSheetWithHeaderAndDedup(revisedOutputSheetId, filteredHeader, revisedNewData, [0, 2]);  // Order Date + App User/Login

//   // 10秒待機
//   Utilities.sleep(10000);

//   SpreadsheetApp.getUi().alert("Complete Step2,next Step3");
// }

/**
 * 複数のスプレッドシートからデータを統合し、指定されたシートに出力する包括的な関数
 * Sales_Order、Pick/Pack、CartonCount_Customerのデータを統合し、
 * Integrated_Dataシートと複数の外部シートを更新します。
 */
function createSummarySheetAndProcessData() {
  // 現在アクティブなスプレッドシートID
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const integratedDataId = ss.getId();

  // 外部スプレッドシートID
  const pickingDataId = '1Vp08XkCklvrF6DfZEPDzOBuLl6DclY3qUaua6nE7cEw';
  const integratedMainOutputId = '1Jot9ZgygJN3FukLWZtd7ZuSoz1ybzhvgMjff_UR7OpU';
  const integratedSubOutputId = '1mJ9z4wJ2AfCD5fkr_ocWRo0KUEbcH4R3J2YgQedqOEE';
  const revisedSalesOrderOutputId = '13mRHbl1mDoVXchUUUlLX0y38VM7bAI512uvJ4iSx_us';
  
  try {
    // データソースの取得
    const salesSheet = ss.getSheetByName('Sales_Order');
    const salesData = salesSheet.getDataRange().getValues();

    const pickingSpreadsheet = SpreadsheetApp.openById(pickingDataId);
    const pickingSheet = pickingSpreadsheet.getSheetByName('Pick/Pack');
    const pickingData = pickingSheet.getDataRange().getValues();

    // `readXlsxAsCsvData` 関数が未定義のため、一旦コメントアウトまたは適切な実装に置き換え
    // const cartonData = readXlsxAsCsvData("CartonCount_Customer", "CartonCount_Customer");
    // ダミーデータまたは適切な実装に置き換えてください
    const cartonData = []; // この行を適切なデータ取得処理に置き換えてください

    // ヘッダーインデックス作成
    const salesIdx = getIndexMap(salesData[0]);
    const pickingIdx = getIndexMap(pickingData[0]);
    const cartonIdx = cartonData.length > 0 ? getIndexMap(cartonData[0]) : {};

    // ピッカー・パッカー・チェッカーの最終列インデックス取得
    const idxPicker2 = lastIndexOfHeader(pickingData[0], 'Picker');
    const idxPacker2 = lastIndexOfHeader(pickingData[0], 'Packer');
    const idxChecker2 = lastIndexOfHeader(pickingData[0], 'Double Checker');

    // Sales_Orderデータからサマリーマップを構築
    let summaryMap = {};
    for (let i = 1, len = salesData.length; i < len; i++) {
      const row = salesData[i];
      const dateKey = toDateKey(row[salesIdx['Order Date']]);
      const customer = row[salesIdx['Customer']];
      const key = dateKey + '|' + customer;

      if (!summaryMap[key]) {
        summaryMap[key] = {
          Order_Date: row[salesIdx['Order Date']],
          date_only: toSlashDateStr(dateKey),
          Customer: customer,
          SKU: 0,
          qty: 0,
          Ks: 0,
          Depot: row[salesIdx['Sales Team']],
          Course: row[salesIdx['Salesperson']],
          order_time_min: row[salesIdx['Order Date']] && row[salesIdx['Creation Date']] ? timeDiff(row[salesIdx['Creation Date']], row[salesIdx['Order Date']]) : "",
          Staff_Receive: row[salesIdx['App User/Login']],
          order_actual_end_time: formatDateTime(row[salesIdx['Order Date']]),
          order_pred_min: "",
          pick_pred_min: "",
          pack_pred_min: "",
          pick_time_min: "",
          pack_time_min: "",
          Staff_Pick: "",
          Staff_Pack: "",
          Staff_Check: "",
          pick_actual_end_time: "",
          pack_actual_end_time: "",
          Carton: "",
          Beverage: ""
        };
      }
      summaryMap[key].SKU += 1;
      summaryMap[key].qty += Number(row[salesIdx['Order Lines/Quantity']]) || 0;
      summaryMap[key].Ks += Number(row[salesIdx['Total']]) || 0;
    }

    // Pick/Packデータをサマリーマップに統合
    for (let i = 1, len = pickingData.length; i < len; i++) {
      const row = pickingData[i];
      if (row[pickingIdx['DC/Depot']] && typeof row[pickingIdx['DC/Depot']] === "string") {
        row[pickingIdx['DC/Depot']] = row[pickingIdx['DC/Depot']].replace(/Depot-/g, "");
      }
      let shopNoRaw = row[pickingIdx['Shop No.']];
      let shopNo = (typeof shopNoRaw === "number" || (/^\d+$/.test(String(shopNoRaw)))) ? "KMU-" + shopNoRaw : shopNoRaw;
      const dateKey = toDateKey(row[pickingIdx['Date']]);
      const key = dateKey + '|' + shopNo;

      let summary = summaryMap[key];
      if (!summary) continue;
      summary.pick_time_min = timeDiff(row[pickingIdx['Pick Start time']], row[pickingIdx['Pick Finish time']]);
      summary.pack_time_min = timeDiff(row[pickingIdx['Pack Start time']], row[pickingIdx['Pack Finish time']]);
      summary.Staff_Pick = row[idxPicker2] || "";
      summary.Staff_Pack = row[idxPacker2] || "";
      summary.Staff_Check = row[idxChecker2] || "";
      summary.pick_actual_end_time = combineDateTimeStr(row[pickingIdx['Date']], row[pickingIdx['Pick Finish time']]);
      summary.pack_actual_end_time = combineDateTimeStr(row[pickingIdx['Date']], row[pickingIdx['Pack Finish time']]);
    }

    // Carton_Countデータをサマリーマップに統合
    if (cartonData.length > 1) {
      for (let i = 1, len = cartonData.length; i < len; i++) {
        const row = cartonData[i];
        const dateKey = toDateKey(row[cartonIdx['Date']]);
        const customer = row[cartonIdx['Customer']];
        const key = dateKey + '|' + customer;
        let summary = summaryMap[key];
        if (!summary) continue;
        summary.Carton = row[cartonIdx['Carton']];
        summary.Beverage = row[cartonIdx['Drink']];
      }
    }

    // 予測値計算
    const summaries = Object.values(summaryMap);
    for (let i = 0, len = summaries.length; i < len; i++) {
      let obj = summaries[i];
      obj.Carton = Number(obj.Carton) || 0;
      obj.order_pred_min = (0.9612 + 0.3073 * obj.SKU + 0.0065 * obj.qty).toFixed(4);
      obj.pick_pred_min = (0.2040 + 0.2198 * obj.SKU + 0.0025 * obj.qty + 0.4495 * obj.Carton).toFixed(4);
      obj.pack_pred_min = (-0.9712 + 0.1913 * obj.SKU + 0.0011 * obj.qty + 0.5966 * obj.Carton).toFixed(4);
    }
    
    // -------------------------------------------------------------------
    // 出力処理: 各シートに必要なデータを書き込む
    // -------------------------------------------------------------------

    // `Integrated_Data`シートへの出力
    const headers = [
      'Order_Date', 'date_only', 'Customer', 'SKU', 'qty', 'Ks', 'Depot', 'Course',
      'pick_time_min', 'pack_time_min', 'order_time_min', 'Staff_Receive',
      'Staff_Pick', 'Staff_Pack', 'Staff_Check',
      'order_pred_min', 'pick_pred_min', 'pack_pred_min',
      'order_actual_end_time', 'pick_actual_end_time', 'pack_actual_end_time',
      'Carton', 'Beverage'
    ];
    const outData = [headers];
    for (let i = 0, len = summaries.length; i < len; i++) {
      const obj = summaries[i];
      let row = [];
      for (let j = 0; j < headers.length; j++) {
        let h = headers[j];
        row.push(obj[h] !== undefined ? obj[h] : "");
      }
      outData.push(row);
    }

    // `outputToExternalSheetWithHeaderAndUpsert` 関数は未定義のため、同等の処理を実装
    let outSheet = ss.getSheetByName('Integrated_Data');
    if (!outSheet) outSheet = ss.insertSheet('Integrated_Data');
    
    // 既存データの読み込みと更新
    const existingValues = outSheet.getDataRange().getValues();
    const existingHeader = existingValues.length > 0 ? existingValues[0] : [];
    const existingData = existingValues.slice(1);
    
    const keyIndices = [headers.indexOf('Order_Date'), headers.indexOf('Customer')];
    const existingMap = new Map();
    for (const row of existingData) {
      const key = keyIndices.map(i => row[i]).join('|');
      existingMap.set(key, row);
    }
    
    const newOrUpdatedData = outData.slice(1);
    for (const newRow of newOrUpdatedData) {
      const key = keyIndices.map(i => newRow[i]).join('|');
      existingMap.set(key, newRow);
    }
    
    const finalData = [headers, ...Array.from(existingMap.values())];
    outSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);

    // `clearOldDataAndEmptyRows` は今回の統合ロジックで不要と判断し削除

    // `Integrated_Data_Main`シートへの出力
    const targetHeaders_1 = [
      'Order_Date', 'date_only', 'Customer', 'SKU', 'qty', 'Ks', 'Depot', 'Course', 'Carton', 'Beverage', 'pick_time_min', 'pack_time_min', 'order_time_min'
    ];
    const newData_1 = summaries.map(obj => targetHeaders_1.map(h => obj[h] !== undefined ? obj[h] : ""));
    outputToExternalSheetWithHeaderAndUpsert(integratedMainOutputId, targetHeaders_1, newData_1, [0, 2]);

    // `Integrated_Data_Sub`シートへの出力
    const targetHeaders_2 = [
      'Order_Date', 'date_only', 'Staff_Receive', 'Staff_Pick', 'Staff_Pack', 'Staff_Check'
    ];
    const newData_2 = summaries.map(obj => targetHeaders_2.map(h => obj[h] !== undefined ? obj[h] : ""));
    outputToExternalSheetWithHeaderAndUpsert(integratedSubOutputId, targetHeaders_2, newData_2, [0, 4, 5]);

    // `sales_order`シートへの出力
    const allowedHeaders = ['Order Date', 'Creation Date', 'App User/Login'];
    const header = salesData[0];
    const colIdxs = allowedHeaders.map(h => header.indexOf(h));
    const filteredHeader = colIdxs.map(idx => header[idx]);

    const uniqueSet = new Set();
    const revisedNewData = [];
    salesData.slice(1).forEach(row => {
      const values = colIdxs.map(idx => row[idx]);
      const staff = values[2];
      if (!staff || staff.toString().trim() === '') return;
      const key = values.join('__');
      if (!uniqueSet.has(key)) {
        uniqueSet.add(key);
        revisedNewData.push(values);
      }
    });
    outputToExternalSheetWithHeaderAndDedup(revisedSalesOrderOutputId, filteredHeader, revisedNewData, [0, 2]);

    Utilities.sleep(10000); // 10秒待機

    SpreadsheetApp.getUi().alert("Complete Step2, next Step3");

  } catch (e) {
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}

// -------------------------------------------------------------------
// ヘルパー関数
// -------------------------------------------------------------------

/**
 * ヘッダー配列から列名とインデックスのマップを作成する
 * @param {Array<string>} headers
 * @returns {Object<string, number>}
 */
function getIndexMap(headers) {
  const map = {};
  headers.forEach((h, i) => {
    map[String(h).trim()] = i;
  });
  return map;
}

/**
 * 日付オブジェクトをyyyyMMdd形式の文字列に変換
 * @param {Date} date
 * @returns {string}
 */
function toDateKey(date) {
  if (!(date instanceof Date)) return '';
  const y = date.getFullYear();
  const m = (date.getMonth() + 1).toString().padStart(2, '0');
  const d = date.getDate().toString().padStart(2, '0');
  return `${y}${m}${d}`;
}

/**
 * 日付キー(yyyyMMdd)をyyyy/MM/dd形式の文字列に変換
 * @param {string} dateKey
 * @returns {string}
 */
function toSlashDateStr(dateKey) {
  if (typeof dateKey !== 'string' || dateKey.length !== 8) return '';
  return `${dateKey.substring(0, 4)}/${dateKey.substring(4, 6)}/${dateKey.substring(6, 8)}`;
}

/**
 * 開始時間と終了時間の差分を分単位で計算する
 * @param {Date|string} start
 * @param {Date|string} end
 * @returns {number|string}
 */
function timeDiff(start, end) {
  if (!start || !end) return "";
  const startDate = new Date(start);
  const endDate = new Date(end);
  if (isNaN(startDate.getTime()) || isNaN(endDate.getTime())) return "";
  const diffMs = endDate.getTime() - startDate.getTime();
  const diffMins = diffMs / (1000 * 60);
  return diffMins > 0 ? diffMins : "";
}

/**
 * 日付オブジェクトを指定された書式で文字列に変換する
 * @param {Date} date
 * @returns {string}
 */
function formatDateTime(date) {
  if (!(date instanceof Date)) return "";
  return Utilities.formatDate(date, SpreadsheetApp.getActiveSpreadsheet().getSpreadsheetTimeZone(), "yyyy/MM/dd HH:mm:ss");
}

/**
 * ヘッダー配列内で指定された列名の最後のインデックスを取得する
 * @param {Array<string>} headers
 * @param {string} headerName
 * @returns {number}
 */
function lastIndexOfHeader(headers, headerName) {
  const reversedHeaders = headers.slice().reverse();
  const reversedIndex = reversedHeaders.findIndex(h => h.trim() === headerName.trim());
  return reversedIndex !== -1 ? headers.length - 1 - reversedIndex : -1;
}

/**
 * 日付と時刻の文字列を結合し、日付オブジェクトに変換する
 * @param {Date|string} date
 * @param {Date|string} time
 * @returns {Date|string}
 */
function combineDateTimeStr(date, time) {
  if (!date || !time) return "";
  const dateStr = Utilities.formatDate(new Date(date), Session.getScriptTimeZone(), 'yyyy/MM/dd');
  const timeStr = Utilities.formatDate(new Date(time), Session.getScriptTimeZone(), 'HH:mm:ss');
  return `${dateStr} ${timeStr}`;
}

/**
 * 指定されたスプレッドシートのシートに、ヘッダーを基準にデータをアップサート（追加/更新）する。
 * @param {string} sheetId - スプレッドシートID
 * @param {Array<string>} headers - 書き込むデータのヘッダー
 * @param {Array<Array<any>>} newData - 書き込むデータ（ヘッダーを除く）
 * @param {Array<number>} keyIndices - データを一意に識別するためのキー列のインデックス配列（0-based）
 * @param {string} [sheetName] - 書き込み先のシート名（省略可能、省略時は先頭シート）
 */
function outputToExternalSheetWithHeaderAndUpsert(sheetId, headers, newData, keyIndices, sheetName) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sheet) throw new Error(`指定されたシート名 "${sheetName}" が見つかりません。`);

  const existingRange = sheet.getDataRange();
  const existingValues = existingRange.getValues();
  const existingHeaders = existingValues.length > 0 ? existingValues[0] : [];
  const existingData = existingValues.slice(1);

  const existingMap = new Map();
  for (const row of existingData) {
    const key = keyIndices.map(i => row[i]).join('|');
    existingMap.set(key, row);
  }

  for (const newRow of newData) {
    const key = keyIndices.map(i => newRow[i]).join('|');
    existingMap.set(key, newRow);
  }

  const finalData = [headers, ...Array.from(existingMap.values())];
  sheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
}

/**
 * 指定されたスプレッドシートのシートに、ヘッダーを基準にデータを追記し、重複を排除する。
 * @param {string} sheetId - スプレッドシートID
 * @param {Array<string>} headers - 書き込むデータのヘッダー
 * @param {Array<Array<any>>} newData - 書き込むデータ（ヘッダーを除く）
 * @param {Array<number>} keyIndices - データを一意に識別するためのキー列のインデックス配列（0-based）
 * @param {string} [sheetName] - 書き込み先のシート名（省略可能、省略時は先頭シート）
 */
function outputToExternalSheetWithHeaderAndDedup(sheetId, headers, newData, keyIndices, sheetName) {
  const ss = SpreadsheetApp.openById(sheetId);
  const sheet = sheetName ? ss.getSheetByName(sheetName) : ss.getSheets()[0];
  if (!sheet) throw new Error(`指定されたシート名 "${sheetName}" が見つかりません。`);

  const existingRange = sheet.getDataRange();
  const existingValues = existingRange.getValues();
  const existingData = existingValues.slice(1);

  const existingSet = new Set();
  for (const row of existingData) {
    const key = keyIndices.map(i => row[i]).join('|');
    existingSet.add(key);
  }

  const filteredNewData = newData.filter(newRow => {
    const key = keyIndices.map(i => newRow[i]).join('|');
    return !existingSet.has(key);
  });

  if (filteredNewData.length > 0) {
    const lastRow = sheet.getLastRow();
    sheet.getRange(lastRow + 1, 1, filteredNewData.length, filteredNewData[0].length).setValues(filteredNewData);
  }
}


/**
 * 昨日以前のpick/pack完了データ行を削除し、空白行を整理する
 * @param {GoogleAppsScript.Spreadsheet.Sheet} outSheet 対象シート
 */
function clearOldDataAndEmptyRows(outSheet) {
  const lastRow = outSheet.getLastRow();
  const lastCol = outSheet.getLastColumn();

  // データがヘッダーしかない場合は何もしない
  if (lastRow <= 1) return;

  const data = outSheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = data[0];
  const rows = data.slice(1);

  // ヘッダーから必要な列のインデックスを取得
  const dateOnlyColIndex = headers.indexOf('date_only');
  const pickTimeColIndex = headers.indexOf('pick_time_min');
  const packTimeColIndex = headers.indexOf('pack_time_min');

  // インデックスが見つからない場合は処理を中止
  if (dateOnlyColIndex === -1 || pickTimeColIndex === -1 || packTimeColIndex === -1) {
    throw new Error('必要なヘッダー列が見つかりません: date_only, pick_time_min, pack_time_min');
  }

  // 今日の0時0分のタイムスタンプ
  const today0 = new Date(new Date().getFullYear(), new Date().getMonth(), new Date().getDate()).getTime();

  const newRows = [];
  rows.forEach(row => {
    const dStr = row[dateOnlyColIndex];
    // 日付がない行はそのまま残す
    if (!dStr) {
      newRows.push(row);
      return;
    }
    
    // 日付文字列を日付オブジェクトに変換
    const d = new Date(dStr);
    
    // 日付として無効な場合はそのまま残す
    if (isNaN(d.getTime())) {
      newRows.push(row);
      return;
    }
    
    const hasPick = row[pickTimeColIndex] && row[pickTimeColIndex] !== '';
    const hasPack = row[packTimeColIndex] && row[packTimeColIndex] !== '';
    const hasEither = hasPick || hasPack;

    // 「昨日より前」で「pick/packに値がある」行は除外
    if (!(hasEither && d.getTime() < today0)) {
      newRows.push(row);
    }
  });

  // シートの内容をクリアし、新しいデータで書き直す
  outSheet.clearContents();
  if (newRows.length > 0) {
    outSheet.getRange(1, 1, newRows.length + 1, lastCol).setValues([headers, ...newRows]);
  } else {
    // データが空になった場合はヘッダーのみ書き込む
    outSheet.getRange(1, 1, 1, lastCol).setValues([headers]);
  }
}
