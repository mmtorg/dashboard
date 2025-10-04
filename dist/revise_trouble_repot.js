function troubleReportConvertAndOutput() {
  const xlsxFileId = "1A3hksVbw4i9WCpAuyIzBmezy_YKEhhyu";
  const file = DriveApp.getFileById(xlsxFileId);
  const myDriveRoot = DriveApp.getRootFolder();
  const copiedFile = file.makeCopy('temp_convert.xlsx', myDriveRoot);

  // const resource = {
  //   name: 'temp_convert',
  //   mimeType: 'application/vnd.google-apps.spreadsheet'
  // };

  const sheetFileId = Drive.Files.copy({
    name: 'temp_convert',
    mimeType: 'application/vnd.google-apps.spreadsheet'
  }, copiedFile.getId()).id;

  // 変換されたスプレッドシートを開く
  const ss = SpreadsheetApp.openById(sheetFileId);
  const sheets = ss.getSheets();

  // 残したいシート名（例："Trouble"）だけ残して他は削除
  const keepSheetName = 'Trouble';
  sheets.forEach(sheet => {
    if (sheet.getName() !== keepSheetName) {
      ss.deleteSheet(sheet);
    }
  });

  const sourceSpreadsheet = SpreadsheetApp.openById(sheetFileId);
  const sourceSheet = sourceSpreadsheet.getSheetByName("Trouble");
  if (!sourceSheet) throw new Error('「Trouble」シートが見つかりません');

  const startRow = 5;
  const lastRow = sourceSheet.getLastRow();
  if (lastRow < startRow) return;

  const rawData = sourceSheet.getRange(startRow, 1, lastRow - startRow + 1, sourceSheet.getLastColumn()).getValues();

  // 必要な列のインデックス（1始まり）
  const columns = [1, 3, 5, 16, 43, 44, 45];

  const outputRows = rawData.map((row, i) => {
    const extracted = columns.map(index => row[index - 1]);
    const cValue = extracted[2]; // Shop no.
    const dValue = extracted[3]; // キャンセル理由

    // Revised Shop no.
    const revised = (typeof cValue === 'number') ? `KMU-${("000" + cValue).slice(-3)}` : cValue;

    // Broken Egg 判定
    const broken = (dValue && dValue.toString().toLowerCase().includes('broken') && dValue.toString().toLowerCase().includes('egg')) ? 'Broken Egg' : '';

    return [
      extracted[1],     // Date Order
      extracted[4],     // Responsibility
      extracted[5],     // Error
      revised,          // Revised Shop no.
      broken            // Broken egg
    ];
  });

  const outputHeaders = ['Date Order', 'Responsibility', 'Error', 'Revised Shop no.', 'Broken egg'];
  const revisedOutputSheetId = '1AUBSOigFfcHxeDhsyX3zAstG05URtNGpXovBVrI7DTI';
  const outputSheet = SpreadsheetApp.openById(revisedOutputSheetId).getSheets()[0];

  // 出力
  outputSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

  if (outputSheet.getLastRow() > 1) {
    outputSheet.getRange(2, 1, outputSheet.getLastRow() - 1, outputSheet.getLastColumn()).clearContent();
  }

  if (outputRows.length > 0) {
    outputSheet.getRange(2, 1, outputRows.length, outputRows[0].length).setValues(outputRows);
  }

  // 一時ファイル削除
  DriveApp.getFileById(sheetFileId).setTrashed(true);
  copiedFile.setTrashed(true);

  // 10秒待機
  Utilities.sleep(10000);

  SpreadsheetApp.getUi().alert("Complete Step4,done");
}



// // ▼ Trouble_Report 新規追加ではなく全上書き・A2:Iクリア・数式&Broken Egg判定復元
// function troubleReportConvertAndAppend() {
//   var xlsxFileId = "1A3hksVbw4i9WCpAuyIzBmezy_YKEhhyu";
//   var file = DriveApp.getFileById(xlsxFileId);  // ← file を取得
//   // const sheetFileId = convertExcelToSpreadsheet(xlsxFileId); 

//   // var file = DriveApp.getFileById(xlsxFileId);
//   // var blob = file.getBlob();

//   // var resource = {
//   //   title: 'temp_convert',
//   //   mimeType: MimeType.GOOGLE_SHEETS
//   // };
//   // var convertedFile = Drive.Files.insert(resource, blob);
//   // var sheetFileId = convertedFile.id;

//   // 一時ファイルをマイドライブ直下に作成
//   var myDriveRoot = DriveApp.getRootFolder();
//   var copiedFile = file.makeCopy('temp_convert.xlsx', myDriveRoot);

//   var copiedFile = file.makeCopy('temp_convert.xlsx');
//   var resource = {
//     name: 'temp_convert',
//     mimeType: 'application/vnd.google-apps.spreadsheet'
//   };
//   var convertedFile = Drive.Files.copy(resource, copiedFile.getId());
//   var sheetFileId = convertedFile.id;

//   var sourceSpreadsheet = SpreadsheetApp.openById(sheetFileId);
//   var sourceSheet = sourceSpreadsheet.getSheetByName("Trouble");
//   var targetSpreadsheet = SpreadsheetApp.openById("1W-ZRUO797GhBE600KU4Q8WqGgoSE_jjTC8vBEo0uupU");
//   var targetSheet = targetSpreadsheet.getSheetByName("Trouble_Report");

//   var header = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
//   var idx = getIndexMap(header);
//   var keyColumns = ['No.', 'Date Order', 'Shop no.'];

//   var lastRow = sourceSheet.getLastRow();
//   var startRow = 5;
//   if (lastRow < startRow) return;
//   var data = sourceSheet.getRange(startRow, 1, lastRow - startRow + 1, sourceSheet.getLastColumn()).getValues();

//   // 必要な列（A〜G）
//   var columns = [1, 3, 5, 16, 43, 44, 45];
//   var filteredData = data.map(function(row) {
//     return columns.map(function(colIndex) {
//       return row[colIndex - 1];
//     });
//   });

//   // A2:Iをクリアしてから全上書き
//   targetSheet.getRange("A2:I").clearContent();

//   if (filteredData.length > 0) {
//     var outputRows = [];

//     for (var i = 0; i < filteredData.length; i++) {
//       var row = filteredData[i].slice(); // A〜G列

//       var rowNum = i + 2;
//       var cValue = row[2]; // C列（Shop no.）
//       var dValue = row[3]; // D列（キャンセル理由）

//       // H列（Revised Shop no. 数式）
//       var hFormula = `=IF(ISNUMBER(C${rowNum}),"KMU-"&(REPT(0,3-LEN(C${rowNum}))&C${rowNum}),C${rowNum})`;

//       // I列（Broken Egg 判定）
//       var iText = (dValue && dValue.toString().toLowerCase().includes('broken') && dValue.toString().toLowerCase().includes('egg')) ? 'Broken Egg' : '';

//       row.push(hFormula); // H列
//       row.push(iText);    // I列

//       outputRows.push(row);
//     }

//     // A〜I列の値を貼り付け（H列はあとで setFormulas で置き換える）
//     targetSheet.getRange(2, 1, outputRows.length, 9).setValues(outputRows);

//     // H列のみ setFormulas で再設定
//     var hFormulas = outputRows.map(row => [row[7]]);
//     targetSheet.getRange(2, 8, hFormulas.length, 1).setFormulas(hFormulas);
//   }

//   // 一時ファイル削除
//   DriveApp.getFileById(sheetFileId).setTrashed(true);
//   copiedFile.setTrashed(true);

//   // ▼ 以前のロジック（新規のみ追加）は不要になったがコメントとして残す
//   /*
//   // 既存キー抽出（新規のみ追加）
//   var dstLastRow = targetSheet.getLastRow();
//   var dstHeader = targetSheet.getRange(1, 1, 1, columns.length).getValues()[0];
//   var dstIdx = getIndexMap(dstHeader);
//   var dstValues = dstLastRow >= 2 ? targetSheet.getRange(2, 1, dstLastRow - 1, columns.length).getValues() : [];
//   var dstKeys = new Set(dstValues.map(row => createRowKeyWithColumns(row, dstIdx, keyColumns)));

//   // 新規だけ抽出
//   var newRows = [];
//   for (var i = 0; i < filteredData.length; i++) {
//     var key = createRowKeyWithColumns(data[i], idx, keyColumns);
//     if (!dstKeys.has(key)) {
//       newRows.push(filteredData[i]);
//     }
//   }
//   */

//   // 2. Trouble_Reportシートの指定列の2行目以降をコピー
//   var colNames = ['Date Order', 'Responsibility', 'Error', 'Revised Shop no.', 'Broken egg'];
//   var reportHeader = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
//   var colIndexes = colNames.map(function(h) {
//     var idx = reportHeader.map(hdr => hdr.trim()).indexOf(h.trim());
//     if (idx === -1) throw new Error(`ヘッダー「${h}」が見つかりません`);
//     return idx + 1;
//   });

//   var reportRowCount = targetSheet.getLastRow() - 1;
//   var result = [];

//   if (reportRowCount > 0) {
//     var colsData = colIndexes.map(col => targetSheet.getRange(2, col, reportRowCount, 1).getValues());
//     for (var i = 0; i < reportRowCount; i++) {
//       result.push(colsData.map(col => col[i][0]));
//     }
//   }

//   // 【DB】trouble_report 出力
//   const revisedOutputSheetId = '1AUBSOigFfcHxeDhsyX3zAstG05URtNGpXovBVrI7DTI';
//   const outputHeaders = colNames;
//   const outputData    = result;

//   // 出力先スプレッドシートとシート名
//   const outputSpreadsheet = SpreadsheetApp.openById(revisedOutputSheetId);
//   const outputSheet = outputSpreadsheet.getSheets()[0];  // もしくは getSheetByName("trouble_report")

//   // ヘッダー設定（1行目）
//   outputSheet.getRange(1, 1, 1, outputHeaders.length).setValues([outputHeaders]);

//   // A2以降をクリア
//   if (outputSheet.getLastRow() > 1) {
//     outputSheet.getRange(2, 1, outputSheet.getLastRow() - 1, outputSheet.getLastColumn()).clearContent();
//   }

//   // A2以降に新データを貼り付け
//   if (outputData.length > 0) {
//     outputSheet.getRange(2, 1, outputData.length, outputData[0].length).setValues(outputData);
//   }

//   // Date Order(0), Responsibility(1), Revised Shop no.(3), Broken egg(4)
//   // outputToExternalSheetWithHeaderAndDedup(
//   //   revisedOutputSheetId,
//   //   outputHeaders,
//   //   outputData,
//   //   [0, 1, 3, 4]
//   // );

//   // // Revised_Trouble_Reportシート対応（使用していないが残す）
//   // var revisedSheet = targetSpreadsheet.getSheetByName('Revised_Trouble_Report');
//   // if (!revisedSheet) throw new Error('Revised_Trouble_Reportシートが存在しません');
//   // if (revisedSheet.getLastRow() > 1) {
//   //   revisedSheet.getRange(2, 1, revisedSheet.getLastRow() - 1, revisedSheet.getLastColumn()).clearContent();
//   // }
//   // if (reportRowCount > 0) {
//   //   revisedSheet.getRange(2, 1, result.length, result[0].length).setValues(result);
//   // }
// }

// function convertExcelToSpreadsheet(xlsxFileId) {
//   const file = DriveApp.getFileById(xlsxFileId);
//   const blob = file.getBlob();

//   const metadata = {
//     name: 'temp_convert',
//     mimeType: 'application/vnd.google-apps.spreadsheet'
//   };

//   const options = {
//     method: "post",
//     contentType: "application/json",
//     payload: JSON.stringify(metadata),
//     headers: {
//       Authorization: "Bearer " + ScriptApp.getOAuthToken()
//     },
//     muteHttpExceptions: true
//   };

//   const uploadUrl = "https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&convert=true";

//   const boundary = "-------314159265358979323846";
//   const delimiter = "\r\n--" + boundary + "\r\n";
//   const close_delim = "\r\n--" + boundary + "--";

//   const multipartRequestBody =
//     delimiter +
//     'Content-Type: application/json\r\n\r\n' +
//     JSON.stringify(metadata) +
//     delimiter +
//     'Content-Type: ' + blob.getContentType() + '\r\n\r\n' +
//     blob.getBytes() +
//     close_delim;

//   const uploadOptions = {
//     method: "post",
//     contentType: "multipart/related; boundary=" + boundary,
//     payload: multipartRequestBody,
//     headers: {
//       Authorization: "Bearer " + ScriptApp.getOAuthToken()
//     }
//   };

//   const response = UrlFetchApp.fetch(uploadUrl, uploadOptions);
//   const result = JSON.parse(response.getContentText());

//   return result.id; // ← 新しく作られたスプレッドシートのID
// }


// // // ▼ Trouble_Report 新規追加のみ・A2:Iクリア・数式&Broken Egg判定復元
// // function troubleReportConvertAndAppend() {
// //   var xlsxFileId = "1UikDbDllSS2arEedUaM6B4A8IHZQveLX";
// //   var file = DriveApp.getFileById(xlsxFileId);
// //   var copiedFile = file.makeCopy('temp_convert.xlsx');
// //   var resource = {
// //     name: 'temp_convert',
// //     mimeType: 'application/vnd.google-apps.spreadsheet'
// //   };
// //   var convertedFile = Drive.Files.copy(resource, copiedFile.getId());
// //   var sheetFileId = convertedFile.id;

// //   var sourceSpreadsheet = SpreadsheetApp.openById(sheetFileId);
// //   var sourceSheet = sourceSpreadsheet.getSheetByName("Trouble");
// //   var targetSpreadsheet = SpreadsheetApp.openById("1W-ZRUO797GhBE600KU4Q8WqGgoSE_jjTC8vBEo0uupU");
// //   var targetSheet = targetSpreadsheet.getSheetByName("Trouble_Report");

// //   var header = sourceSheet.getRange(1, 1, 1, sourceSheet.getLastColumn()).getValues()[0];
// //   var idx = getIndexMap(header);
// //   var keyColumns = ['No.', 'Date Order', 'Shop no.'];

// //   var lastRow = sourceSheet.getLastRow();
// //   var startRow = 5;
// //   if (lastRow < startRow) return;
// //   var data = sourceSheet.getRange(startRow, 1, lastRow - startRow + 1, sourceSheet.getLastColumn()).getValues();

// //   // 必要な列
// //   var columns = [1, 3, 5, 16, 43, 44, 45];
// //   var filteredData = data.map(function(row) {
// //     return columns.map(function(colIndex) {
// //       return row[colIndex - 1];
// //     });
// //   });

// //   // 既存キー抽出（新規のみ追加）
// //   var dstLastRow = targetSheet.getLastRow();
// //   var dstHeader = targetSheet.getRange(1, 1, 1, columns.length).getValues()[0];
// //   var dstIdx = getIndexMap(dstHeader);
// //   var dstValues = dstLastRow >= 2 ? targetSheet.getRange(2, 1, dstLastRow - 1, columns.length).getValues() : [];
// //   var dstKeys = new Set(dstValues.map(row => createRowKeyWithColumns(row, dstIdx, keyColumns)));

// //   // 新規だけ抽出
// //   var newRows = [];
// //   for (var i = 0; i < filteredData.length; i++) {
// //     var key = createRowKeyWithColumns(data[i], idx, keyColumns);
// //     if (!dstKeys.has(key)) {
// //       newRows.push(filteredData[i]);
// //     }
// //   }

// //   // A2:Iをクリアしてから貼り付け、数式・判定列も必ず実施
// //   targetSheet.getRange("A2:I").clearContent();

// //   if (newRows.length > 0) {
// //     targetSheet.getRange(2, 1, newRows.length, columns.length).setValues(newRows);

// //     // H列数式
// //     var formulaRowCount = newRows.length;
// //     if (formulaRowCount > 0) {
// //       var formulas = [];
// //       for (var i = 0; i < formulaRowCount; i++) {
// //         var rowNum = i + 2;
// //         formulas.push([`=IF(ISNUMBER(C${rowNum}),"KMU-"&(REPT(0,3-LEN(C${rowNum}))&C${rowNum}),C${rowNum})`]);
// //       }
// //       targetSheet.getRange(2, 8, formulaRowCount, 1).setFormulas(formulas);
// //     }

// //     // I列判定
// //     var iValues = [];
// //     for (var i = 0; i < newRows.length; i++) {
// //       var dValue = newRows[i][3];
// //       if (dValue) {
// //         var text = dValue.toString().toLowerCase();
// //         if (text.includes('broken') && text.includes('egg')) {
// //           iValues.push(['Broken Egg']);
// //         } else {
// //           iValues.push(['']);
// //         }
// //       } else {
// //         iValues.push(['']);
// //       }
// //     }
// //     targetSheet.getRange(2, 9, iValues.length, 1).setValues(iValues);
// //   }

// //   // 一時ファイル削除
// //   DriveApp.getFileById(sheetFileId).setTrashed(true);
// //   copiedFile.setTrashed(true);

// //   // 2. Trouble_Reportシートの指定列の2行目以降をコピー
// //   // 列名
// //   var colNames = ['Date Order', 'Responsibility', 'Error', 'Revised Shop no.', 'Broken egg'];
// //   // Trouble_Reportのヘッダー取得
// //   var reportHeader = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
// //   // 必要な列のインデックス（1始まりに変換）
// //   var colIndexes = colNames.map(h => reportHeader.indexOf(h) + 1);

// //   // データ取得（2行目以降）
// //   var reportRowCount = targetSheet.getLastRow() - 1;
// //   if (reportRowCount > 0) {
// //     // 指定列だけ抽出
// //     var colsData = colIndexes.map(col => targetSheet.getRange(2, col, reportRowCount, 1).getValues());
// //     // 行列変換（transpose）
// //     var result = [];
// //     for (var i = 0; i < reportRowCount; i++) {
// //       result.push(colsData.map(col => col[i][0]));
// //     }
// //   }

// //     // ヘッダー
// //   var colNames = ['Date Order', 'Responsibility', 'Error', 'Revised Shop no.', 'Broken egg'];

// //   // 指定列だけ抽出
// //   var colsData = colIndexes.map(col => targetSheet.getRange(2, col, reportRowCount, 1).getValues());

// //   // 行列変換
// //   var result = [];
// //   for (var i = 0; i < reportRowCount; i++) {
// //     result.push(colsData.map(col => col[i][0]));
// //   }

// //   // 【DB】trouble_report
// //   const revisedOutputSheetId = '1AUBSOigFfcHxeDhsyX3zAstG05URtNGpXovBVrI7DTI';

// //   const outputHeaders = colNames;
// //   const outputData    = result;

// //   // Date Order(0), Responsibility(1), Revised Shop no.(3), Broken egg(4)
// //   outputToExternalSheetWithHeaderAndDedup(
// //     revisedOutputSheetId,
// //     outputHeaders,
// //     outputData,
// //     [0, 1, 3, 4]
// //   );

// //   // // Revised_Trouble_Reportシート対応
// //   // // 1. Revised_Trouble_Reportシートの2行目以降をクリア
// //   // var revisedSheet = targetSpreadsheet.getSheetByName('Revised_Trouble_Report');
// //   // if (!revisedSheet) throw new Error('Revised_Trouble_Reportシートが存在しません');
// //   // if (revisedSheet.getLastRow() > 1) {
// //   //   revisedSheet.getRange(2, 1, revisedSheet.getLastRow() - 1, revisedSheet.getLastColumn()).clearContent();
// //   // }

// //   // // 2. Trouble_Reportシートの指定列の2行目以降をコピー
// //   // // 列名
// //   // var colNames = ['Date Order', 'Responsibility', 'Error', 'Revised Shop no.', 'Broken egg'];
// //   // // Trouble_Reportのヘッダー取得
// //   // var reportHeader = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0];
// //   // // 必要な列のインデックス（1始まりに変換）
// //   // var colIndexes = colNames.map(h => reportHeader.indexOf(h) + 1);

// //   // // データ取得（2行目以降）
// //   // var reportRowCount = targetSheet.getLastRow() - 1;
// //   // if (reportRowCount > 0) {
// //   //   // 指定列だけ抽出
// //   //   var colsData = colIndexes.map(col => targetSheet.getRange(2, col, reportRowCount, 1).getValues());
// //   //   // 行列変換（transpose）
// //   //   var result = [];
// //   //   for (var i = 0; i < reportRowCount; i++) {
// //   //     result.push(colsData.map(col => col[i][0]));
// //   //   }

// //   //   // 3. Revised_Trouble_ReportのA2に貼り付け
// //   //   revisedSheet.getRange(2, 1, result.length, result[0].length).setValues(result);
// //   // }
// // }