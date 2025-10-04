function transformCsvAndApplyOperations() {
  const csvFileName = "Sales Order (sale.order)"; // 部分一致の基準（拡張子は不要）
  const targetSheetName = "Sales_Order"; // 貼り付け先のシート名
  // const csvFileName = "Sales Order (sale.order).csv"; // CSVファイル名
  // const targetSheetName = "Sales_Order"; // 貼り付け先のシート名

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(targetSheetName);
  if (!sheet) throw new Error("指定されたシートが見つかりません");

  // === 1. CSVファイル読み込み ===

  // 現在のスプレッドシートと同じ階層のフォルダを取得
  const parentFolder = DriveApp.getFileById(ss.getId()).getParents().next();

  // 「input_data」フォルダを取得
  const inputFolders = parentFolder.getFoldersByName('input_data');
  if (!inputFolders.hasNext()) throw new Error('input_data フォルダが見つかりません');
  const inputFolder = inputFolders.next();

  // 「input_data」フォルダ内の対象ファイル取得（部分一致＋CSV限定）
  const filesIter = inputFolder.getFiles();
  let targetFile = null;
  while (filesIter.hasNext()) {
    const file = filesIter.next();
    const fileName = file.getName();
    if (fileName.includes(csvFileName) && fileName.toLowerCase().endsWith(".csv")) {
      targetFile = file;
      break; // 最初に見つかったファイルを使用
    }
  }
  if (!targetFile) throw new Error(`${csvFileName} を含むCSVファイルが input_data フォルダ内に見つかりません`);

  // const files = inputFolder.getFilesByName(csvFileName);
  // if (!files.hasNext()) throw new Error(`${csvFileName} が input_data フォルダ内に見つかりません`);

  const csvBlob = targetFile.getBlob();
  // const csvBlob = files.next().getBlob();
  const csvData = Utilities.parseCsv(csvBlob.getDataAsString("utf-8"));

  // const folder = DriveApp.getFileById(ss.getId()).getParents().next();
  // const files = folder.getFilesByName(csvFileName);
  // if (!files.hasNext()) throw new Error("CSVファイルが見つかりません");

  // === 2. 貼り付け ===
  sheet.clearContents().clearFormats();
  sheet.getRange(1, 1, csvData.length, csvData[0].length).setValues(csvData);

  const data = sheet.getDataRange().getValues(); // 全体を配列で取得（高速）
  const numRows = data.length;
  const numCols = data[0].length;

  // === 3. 空白セルを「直近の上の値」で埋める（式ではなく値として）===
  for (let col = 0; col < numCols; col++) {
    for (let row = 1; row < numRows; row++) {
      const val = data[row][col];
      if (val === "" || val === null || (typeof val === "string" && val.trim() === "")) {
        data[row][col] = data[row - 1][col];
      }
    }
  }

  // === Customer列の列番号を特定 ===
  const headerRow = data[0];
  const customerColIndex = headerRow.indexOf("Customer");
  if (customerColIndex === -1) {
    throw new Error("Customer列が見つかりません");
  }

  // 書き戻し
  sheet.getRange(1, 1, numRows, numCols).setValues(data);

  // === 4. 書式コピー ===
  sheet.getRange("A2").copyTo(sheet.getRange("A3:A" + numRows), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);
  sheet.getRange("B2").copyTo(sheet.getRange("B3:B" + numRows), SpreadsheetApp.CopyPasteType.PASTE_FORMAT, false);

  // === 5. Customer列 "YGN-HQ" 削除（customerColIndex を使う） ===
  const filteredData = [data[0]];
  for (let i = 1; i < data.length; i++) {
    if (data[i][customerColIndex] !== "YGN-HQ") {
      filteredData.push(data[i]);
    }
  }

  sheet.clearContents();
  sheet.getRange(1, 1, filteredData.length, filteredData[0].length).setValues(filteredData);

  // === 6. フィルター解除 ===
  if (sheet.getFilter()) {
    sheet.getFilter().remove();
  }

  // 10秒待機
  Utilities.sleep(10000);

  SpreadsheetApp.getUi().alert("Complete Step1, next Step2");
}
