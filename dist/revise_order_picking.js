// ▼ Order_Picking 新規追加のみ
function copyOnlyNewRowsWithKey() {
  var srcSpreadsheet = SpreadsheetApp.openById('1Vp08XkCklvrF6DfZEPDzOBuLl6DclY3qUaua6nE7cEw');
  var srcSheet = srcSpreadsheet.getSheetByName('Pick/Pack');
  var srcAcol = srcSheet.getRange('A:A').getValues();
  var srcLastRow = 0;
  for (var i = srcAcol.length - 1; i >= 0; i--) {
    if (srcAcol[i][0] !== "" && srcAcol[i][0] !== null) {
      srcLastRow = i + 1;
      break;
    }
  }
  if (srcLastRow < 2) return;

  var srcValues = srcSheet.getRange(2, 1, srcLastRow - 1, 19).getValues(); // A～S列
  var header = srcSheet.getRange(1, 1, 1, 19).getValues()[0];
  var idx = getIndexMap(header);
  var keyColumns = ['Date', 'DC/Depot', 'Car', 'Shop No.'];

  var dstSheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName('Order_Picking');
  var dstAcol = dstSheet.getRange('A:A').getValues();
  var dstLastRow = 0;
  for (var i = dstAcol.length - 1; i >= 0; i--) {
    if (dstAcol[i][0] !== "" && dstAcol[i][0] !== null) {
      dstLastRow = i + 1;
      break;
    }
  }
  var dstHeader = dstSheet.getRange(1, 1, 1, 19).getValues()[0];
  var dstIdx = getIndexMap(dstHeader);
  var dstValues = dstLastRow >= 2 ? dstSheet.getRange(2, 1, dstLastRow - 1, 19).getValues() : [];
  var dstKeys = new Set(dstValues.map(row => createRowKeyWithColumns(row, dstIdx, keyColumns)));

  var newRows = [];
  for (var i = 0; i < srcValues.length; i++) {
    if (!srcValues[i][0]) continue;
    var key = createRowKeyWithColumns(srcValues[i], idx, keyColumns);
    if (!dstKeys.has(key)) {
      newRows.push(srcValues[i]);
    }
  }
  if (newRows.length > 0) {
    var pasteRow = dstLastRow < 2 ? 2 : dstLastRow + 1;
    dstSheet.getRange(pasteRow, 1, newRows.length, 19).setValues(newRows);
  }
}