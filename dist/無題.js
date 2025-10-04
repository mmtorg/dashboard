/**
 * Pick/PackシートのデータをIntegrated_Dataシートに反映する関数
 */
function updateIntegratedDataFromPicking() {
  const integratedDataId = '1W-ZRUO797GhBE600KU4Q8WqGgoSE_jjTC8vBEo0uupU';
  const pickingDataId = '1Vp08XkCklvrF6DfZEPDzOBuLl6DclY3qUaua6nE7cEw';
  const integratedSheetName = 'Integrated_Data';
  const pickingSheetName = 'Pick/Pack';

  try {
    const ssIntegrated = SpreadsheetApp.openById(integratedDataId);
    const sheetIntegrated = ssIntegrated.getSheetByName(integratedSheetName);
    if (!sheetIntegrated) {
      throw new Error(`Integrated_Dataシートが見つかりません: ${integratedDataId}`);
    }

    const ssPicking = SpreadsheetApp.openById(pickingDataId);
    const sheetPicking = ssPicking.getSheetByName(pickingSheetName);
    if (!sheetPicking) {
      throw new Error(`Pick/Packシートが見つかりません: ${pickingDataId}`);
    }

    // 既存のデータをメモリに読み込み
    const integratedData = sheetIntegrated.getDataRange().getValues();
    const pickingData = sheetPicking.getDataRange().getValues();

    // ヘッダーインデックスマップを作成
    const integratedIdx = getIndexMap(integratedData[0]);
    const pickingIdx = getIndexMap(pickingData[0]);

    // 集計キーを元に、統合データをマップに変換
    const integratedMap = {};
    for (let i = 1; i < integratedData.length; i++) {
      const row = integratedData[i];
      const dateOnly = toDateKey(row[integratedIdx['Order_Date']]);
      const customer = row[integratedIdx['Customer']];
      const key = `${dateOnly}|${customer}`;
      integratedMap[key] = row;
    }

    // Pick/PackデータからIntegrated_Dataを更新
    for (let i = 1; i < pickingData.length; i++) {
      const row = pickingData[i];
      const dateKey = toDateKey(row[pickingIdx['Date']]);
      
      // Shop No.の形式を調整 (例: "MUB-162")
      let shopNoRaw = row[pickingIdx['Shop No.']];
      let shopNo = (typeof shopNoRaw === "number" || (/^\d+$/.test(String(shopNoRaw)))) ? "KMU-" + shopNoRaw : shopNoRaw;
      
      // DC/Depotを調整 (例: "Depot-MUB" -> "MUB")
      if (row[pickingIdx['DC/Depot']] && typeof row[pickingIdx['DC/Depot']] === "string") {
        row[pickingIdx['DC/Depot']] = row[pickingIdx['DC/Depot']].replace(/Depot-/g, "");
      }
      
      // 複合キーを作成
      const key = `${dateKey}|${shopNo}`;

      // 統合データのマップから一致する行を検索
      let integratedRow = integratedMap[key];
      if (integratedRow) {
        // 一致する行が見つかったらデータを更新
        integratedRow[integratedIdx['pick_time_min']] = timeDiff(row[pickingIdx['Pick Start time']], row[pickingIdx['Pick Finish time']]);
        integratedRow[integratedIdx['pack_time_min']] = timeDiff(row[pickingIdx['Pack Start time']], row[pickingIdx['Pack Finish time']]);
        
        // Picker/Packer/Checkerの列名が重複している可能性があるため、最後の列のインデックスを取得
        const idxPicker2 = lastIndexOfHeader(pickingData[0], 'Picker');
        const idxPacker2 = lastIndexOfHeader(pickingData[0], 'Packer');
        const idxChecker2 = lastIndexOfHeader(pickingData[0], 'Double Checker');
        
        integratedRow[integratedIdx['Staff_Pick']] = row[idxPicker2] || "";
        integratedRow[integratedIdx['Staff_Pack']] = row[idxPacker2] || "";
        integratedRow[integratedIdx['Staff_Check']] = row[idxChecker2] || "";
        
        integratedRow[integratedIdx['pick_actual_end_time']] = combineDateTimeStr(row[pickingIdx['Date']], row[pickingIdx['Pick Finish time']]);
        integratedRow[integratedIdx['pack_actual_end_time']] = combineDateTimeStr(row[pickingIdx['Date']], row[pickingIdx['Pack Finish time']]);
      }
    }

    // 更新されたデータを元の2次元配列に戻す
    const updatedData = [integratedData[0]];
    Object.values(integratedMap).forEach(row => updatedData.push(row));

    // シートに一括で書き戻す
    if (updatedData.length > 1) {
      sheetIntegrated.getRange(1, 1, updatedData.length, updatedData[0].length).setValues(updatedData);
    }

    SpreadsheetApp.getUi().alert('Pick/PackデータがIntegrated_Dataに正常に反映されました。');
    
  } catch (e) {
    SpreadsheetApp.getUi().alert(`エラーが発生しました: ${e.message}`);
  }
}

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