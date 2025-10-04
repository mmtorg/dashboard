/**
 * Integrated_Dataシートからデータを集計し、
 * 予測テーブルを生成して外部シートに書き込む。
 * 処理完了後、元のシートから処理済みの行を効率的に削除する。
 */
function generateOperationForecastTable() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inputSheet = ss.getSheetByName('Integrated_Data');

  const inputData = inputSheet.getDataRange().getValues();
  const header = inputData[0];
  const rows = inputData.slice(1);

  const idxDepot = header.indexOf('Depot');
  const idxDate = header.indexOf('date_only');
  const idxCourse = header.indexOf('Course');
  const idxOrderPred = header.indexOf('order_pred_min');
  const idxPickPred = header.indexOf('pick_pred_min');
  const idxPackPred = header.indexOf('pack_pred_min');
  const idxOrderActual = header.indexOf('order_actual_end_time');
  const idxPickActual = header.indexOf('pick_actual_end_time');
  const idxPackActual = header.indexOf('pack_actual_end_time');

  const START_TIMES = {
    'KGG': { Order: 485, Pick: 600, Pack: 630 },
    'MUB': { Order: 485, Pick: 600, Pack: 630 },
    'MGD': { Order: 630, Pick: 780, Pack: 810 },
  };

  const grouped = {};

  rows.forEach(row => {
    const depot = row[idxDepot];
    const date = row[idxDate];
    const course = row[idxCourse];
    const key = depot + '___' + date;

    if (!grouped[key]) {
      grouped[key] = {
        depot,
        date,
        courses: new Set(),
        orderPredTotal: 0,
        pickPredTotal: 0,
        packPredTotal: 0,
        orderActualList: [],
        pickActualList: [],
        packActualList: [],
      };
    }

    const group = grouped[key];
    group.courses.add(course);

    group.orderPredTotal += Number(row[idxOrderPred] || 0);
    group.pickPredTotal += Number(row[idxPickPred] || 0);
    group.packPredTotal += Number(row[idxPackPred] || 0);

    if (row[idxOrderActual]) group.orderActualList.push(new Date(row[idxOrderActual]));
    if (row[idxPickActual]) group.pickActualList.push(new Date(row[idxPickActual]));
    if (row[idxPackActual]) group.packActualList.push(new Date(row[idxPackActual]));
  });

  const result = [];

  const headers = [
    'Depot','date_only','Course_count',
    'staff_count_order','staff_count_pick','staff_count_pack',
    'order_pred_duration_min','pick_pred_duration_min','pack_pred_duration_min',
    'order_actual_duration_min','pick_actual_duration_min','pack_actual_duration_min',
    'order_pred_end_hhmm','pick_pred_end_hhmm','pack_pred_end_hhmm',
    'order_actual_end_hhmm','pick_actual_end_hhmm','pack_actual_end_hhmm',
  ];
  result.push(headers);

  function calcPredictedDuration(depot, task, totalPredMin, staffCount) {
    const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;
    const durationMin = staffCount > 0 ? totalPredMin / staffCount : 0;
    return Math.round(durationMin);
  }

  function calcActualDuration(list, depot, task) {
    list = list.filter(d => d instanceof Date && !isNaN(d.getTime()));
    if (list.length === 0) return '';
    const maxDate = new Date(Math.max.apply(null, list));
    const timezone = 'Asia/Yangon';
    const hh = parseInt(Utilities.formatDate(maxDate, timezone, 'HH'), 10);
    const mm = parseInt(Utilities.formatDate(maxDate, timezone, 'mm'), 10);
    const actualEndMin = hh * 60 + mm;
    const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;
    const durationMin = actualEndMin - startMin;
    return durationMin >= 0 ? durationMin : '';
  }

  function calcEndTime(depot, task, totalPredMin, staffCount) {
    const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;
    const endMin = startMin + (staffCount > 0 ? totalPredMin / staffCount : 0);
    const hh = Math.floor(endMin / 60);
    const mm = Math.floor(endMin % 60);
    return { hhmm: ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2), min: Math.round(endMin) };
  }

  function maxActualTime(list) {
    list = list.filter(d => d instanceof Date && !isNaN(d.getTime()));
    list.forEach(d => { d.setSeconds(0); d.setMilliseconds(0); });
    if (list.length === 0) return { hhmm: '', min: '' };
    const maxDate = new Date(Math.max.apply(null, list));
    maxDate.setSeconds(0);
    maxDate.setMilliseconds(0);
    const timezone = 'Asia/Yangon';
    const hh = parseInt(Utilities.formatDate(maxDate, timezone, 'HH'), 10);
    const mm = parseInt(Utilities.formatDate(maxDate, timezone, 'mm'), 10);
    return {
      hhmm: ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2),
      min: hh * 60 + mm
    };
  }

  for (let key in grouped) {
    const g = grouped[key];
    const courseCount = g.courses.size;
    const staffOrder = courseCount * 2;
    const staffPick = courseCount;
    const staffPack = courseCount;
    const orderPredDuration = calcPredictedDuration(g.depot, 'Order', g.orderPredTotal, staffOrder);
    const pickPredDuration = calcPredictedDuration(g.depot, 'Pick', g.pickPredTotal, staffPick);
    const packPredDuration = calcPredictedDuration(g.depot, 'Pack', g.packPredTotal, staffPack);
    const orderActualDuration = calcActualDuration(g.orderActualList, g.depot, 'Order');
    const pickActualDuration = calcActualDuration(g.pickActualList, g.depot, 'Pick');
    const packActualDuration = calcActualDuration(g.packActualList, g.depot, 'Pack');
    const orderPred = calcEndTime(g.depot, 'Order', g.orderPredTotal, staffOrder);
    const pickPred = calcEndTime(g.depot, 'Pick', g.pickPredTotal, staffPick);
    const packPred = calcEndTime(g.depot, 'Pack', g.packPredTotal, staffPack);
    const orderActual = maxActualTime(g.orderActualList);
    const pickActual = maxActualTime(g.pickActualList);
    const packActual = maxActualTime(g.packActualList);

    result.push([
      g.depot,
      g.date,
      courseCount,
      staffOrder,
      staffPick,
      staffPack,
      orderPredDuration,
      pickPredDuration,
      packPredDuration,
      orderActualDuration,
      pickActualDuration,
      packActualDuration,
      orderPred.hhmm,
      pickPred.hhmm,
      packPred.hhmm,
      orderActual.hhmm,
      pickActual.hhmm,
      packActual.hhmm,
    ]);
  }

  const outputSheetId = '1djSGvGrIKJdbPcrxByYqB33SADzlLvwZz40sFhCxZkg';
  const headersRow = result[0];
  const dataRows = result.slice(1);
  outputToExternalSheetWithHeaderAndUpsert(outputSheetId, headersRow, dataRows, [0, 1, 2]);

  Utilities.sleep(10000);
  SpreadsheetApp.getUi().alert("Complete Step3,next Step4");

  // ここから効率化された削除処理、削除処理コメントアウト
  // const targetSpreadsheetId = '1W-ZRUO797GhBE600KU4Q8WqGgoSE_jjTC8vBEo0uupU';
  // const targetSheetName = 'Integrated_Data';
  
  // try {
  //   const targetSs = SpreadsheetApp.openById(targetSpreadsheetId);
  //   const targetSheet = targetSs.getSheetByName(targetSheetName);
    
  //   if (!targetSheet) {
  //     throw new Error(`指定されたスプレッドシートに ${targetSheetName} シートが見つかりません。`);
  //   }

  //   const allData = targetSheet.getDataRange().getValues();
  //   const headerRow = allData[0];
  //   const dataRowsToProcess = allData.slice(1);
    
  //   const pickTimeColIndex = headerRow.indexOf('pick_time_min');
  //   const packTimeColIndex = headerRow.indexOf('pack_time_min');
    
  //   if (pickTimeColIndex === -1 || packTimeColIndex === -1) {
  //     throw new Error('必要な列(pick_time_min, pack_time_min)が見つかりません。');
  //   }
    
  //   // 削除しない行だけをフィルタリング
  //   const filteredData = dataRowsToProcess.filter(row => {
  //     const pickTime = row[pickTimeColIndex];
  //     const packTime = row[packTimeColIndex];
  //     // pick_time_min か pack_time_min のいずれかが空の場合、残す
  //     return (pickTime === '' || pickTime === null || pickTime === undefined) ||
  //            (packTime === '' || packTime === null || packTime === undefined);
  //   });

  //   // フィルタリングしたデータを一括で書き込み
  //   // ヘッダー行を再度追加
  //   const finalData = [headerRow, ...filteredData];
    
  //   // シートの全内容をクリアしてから書き込むことで、行の削除/挿入を避ける
  //   targetSheet.clearContents();
  //   if (finalData.length > 0) {
  //     targetSheet.getRange(1, 1, finalData.length, finalData[0].length).setValues(finalData);
  //   }
    
  //   Logger.log(`不要な行を効率的に削除しました。`);

  // } catch (e) {
  //   Logger.log(`削除処理中にエラーが発生しました: ${e.message}`);
  //   SpreadsheetApp.getUi().alert(`データ削除中にエラーが発生しました: ${e.message}`);
  // }
}

// -------------------------------------------------------------------
// ヘルパー関数 (変更なし)
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
  const existingHeader = existingValues.length > 0 ? existingValues[0] : [];
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

// function generateOperationForecastTable() {
//   const ss = SpreadsheetApp.getActiveSpreadsheet();
//   const inputSheet = ss.getSheetByName('Integrated_Data');
//   // const outputSheet = ss.getSheetByName('Prediction_Result');

//   const inputData = inputSheet.getDataRange().getValues();
//   const header = inputData[0];
//   const rows = inputData.slice(1);

//   const idxDepot = header.indexOf('Depot');
//   const idxDate = header.indexOf('date_only');
//   const idxCourse = header.indexOf('Course');
//   const idxOrderPred = header.indexOf('order_pred_min');
//   const idxPickPred = header.indexOf('pick_pred_min');
//   const idxPackPred = header.indexOf('pack_pred_min');
//   const idxOrderActual = header.indexOf('order_actual_end_time');
//   const idxPickActual = header.indexOf('pick_actual_end_time');
//   const idxPackActual = header.indexOf('pack_actual_end_time');

//   const START_TIMES = {
//     'KGG': { Order: 485, Pick: 600, Pack: 630 },
//     'MUB': { Order: 485, Pick: 600, Pack: 630 },
//     'MGD': { Order: 630, Pick: 780, Pack: 810 },
//   };

//   const grouped = {};

//   rows.forEach(row => {
//     const depot = row[idxDepot];
//     const date = row[idxDate];
//     const course = row[idxCourse];
//     const key = depot + '___' + date;

//     if (!grouped[key]) {
//       grouped[key] = {
//         depot,
//         date,
//         courses: new Set(),
//         orderPredTotal: 0,
//         pickPredTotal: 0,
//         packPredTotal: 0,
//         orderActualList: [],
//         pickActualList: [],
//         packActualList: [],
//       };
//     }

//     const group = grouped[key];
//     group.courses.add(course);

//     group.orderPredTotal += Number(row[idxOrderPred] || 0);
//     group.pickPredTotal  += Number(row[idxPickPred] || 0);
//     group.packPredTotal  += Number(row[idxPackPred] || 0);

//     if (row[idxOrderActual]) group.orderActualList.push(new Date(row[idxOrderActual]));
//     if (row[idxPickActual])  group.pickActualList.push(new Date(row[idxPickActual]));
//     if (row[idxPackActual])  group.packActualList.push(new Date(row[idxPackActual]));
//   });

//   const result = [];

//   // ヘッダー行
// const headers = [
//   'Depot','date_only','Course_count',
//   'staff_count_order','staff_count_pick','staff_count_pack',
//   'order_pred_duration_min','pick_pred_duration_min','pack_pred_duration_min',
//   'order_actual_duration_min','pick_actual_duration_min','pack_actual_duration_min',

//   // ▼ 予測・実績終了時間算出処理現状使わないのでコメントアウト
//   'order_pred_end_hhmm','pick_pred_end_hhmm','pack_pred_end_hhmm',
//   'order_actual_end_hhmm','pick_actual_end_hhmm','pack_actual_end_hhmm',
//   /*
//   'order_pred_end_hhmm','pick_pred_end_hhmm','pack_pred_end_hhmm',
//   'order_actual_end_hhmm','pick_actual_end_hhmm','pack_actual_end_hhmm',
//   'order_pred_end_min','pick_pred_end_min','pack_pred_end_min',
//   'order_actual_end_min','pick_actual_end_min','pack_actual_end_min'
//   */
// ];
//   result.push(headers);

// function calcPredictedDuration(depot, task, totalPredMin, staffCount) {
//   const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;  // 使用しないが明示
//   const durationMin = staffCount > 0 ? totalPredMin / staffCount : 0;
//   return Math.round(durationMin);  // 経過時間そのもの
// }

// function calcActualDuration(list, depot, task) {
//   list = list.filter(d => d instanceof Date && !isNaN(d.getTime()));
//   if (list.length === 0) return '';

//   const maxDate = new Date(Math.max.apply(null, list));
//   const timezone = 'Asia/Yangon';

//   const hh = parseInt(Utilities.formatDate(maxDate, timezone, 'HH'), 10);
//   const mm = parseInt(Utilities.formatDate(maxDate, timezone, 'mm'), 10);
//   const actualEndMin = hh * 60 + mm;

//   const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;
//   const durationMin = actualEndMin - startMin;

//   return durationMin >= 0 ? durationMin : '';  // 負値なら不正なので空欄
// }

//   // ▼ 予測・実績終了時間算出処理現状使わないのでコメントアウト
//   function calcEndTime(depot, task, totalPredMin, staffCount) {
//     const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;
//     const endMin = startMin + (staffCount > 0 ? totalPredMin / staffCount : 0);
//     const hh = Math.floor(endMin / 60);
//     const mm = Math.floor(endMin % 60);
//     return { hhmm: ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2), min: Math.round(endMin) };
//   }

//   function maxActualTime(list) {
//     list = list.filter(d => d instanceof Date && !isNaN(d.getTime()));
//     list.forEach(d => { d.setSeconds(0); d.setMilliseconds(0); });
//     if (list.length === 0) return { hhmm: '', min: '' };

//     const maxDate = new Date(Math.max.apply(null, list));
//     maxDate.setSeconds(0);
//     maxDate.setMilliseconds(0);

//     const timezone = 'Asia/Yangon';

//     const hh = parseInt(Utilities.formatDate(maxDate, timezone, 'HH'), 10);
//     const mm = parseInt(Utilities.formatDate(maxDate, timezone, 'mm'), 10);

//     return {
//       hhmm: ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2),
//       min: hh * 60 + mm
//     };
//   }
//   /*
//   function calcEndTime(depot, task, totalPredMin, staffCount) {
//     const startMin = START_TIMES[depot] ? START_TIMES[depot][task] : 0;
//     const endMin = startMin + (staffCount > 0 ? totalPredMin / staffCount : 0);
//     const hh = Math.floor(endMin / 60);
//     const mm = Math.floor(endMin % 60);
//     return { hhmm: ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2), min: Math.round(endMin) };
//   }

//   function maxActualTime(list) {
//     list = list.filter(d => d instanceof Date && !isNaN(d.getTime()));
//     list.forEach(d => { d.setSeconds(0); d.setMilliseconds(0); });
//     if (list.length === 0) return { hhmm: '', min: '' };

//     const maxDate = new Date(Math.max.apply(null, list));
//     maxDate.setSeconds(0);
//     maxDate.setMilliseconds(0);

//     const timezone = 'Asia/Yangon';

//     const hh = parseInt(Utilities.formatDate(maxDate, timezone, 'HH'), 10);
//     const mm = parseInt(Utilities.formatDate(maxDate, timezone, 'mm'), 10);

//     return {
//       hhmm: ('0' + hh).slice(-2) + ':' + ('0' + mm).slice(-2),
//       min: hh * 60 + mm
//     };
//   }
//   */

// for (let key in grouped) {
//   const g = grouped[key];
//   const courseCount = g.courses.size;

//   const staffOrder = courseCount * 2;
//   const staffPick = courseCount;
//   const staffPack = courseCount;

// const orderPredDuration = calcPredictedDuration(g.depot, 'Order', g.orderPredTotal, staffOrder);
// const pickPredDuration  = calcPredictedDuration(g.depot, 'Pick', g.pickPredTotal, staffPick);
// const packPredDuration  = calcPredictedDuration(g.depot, 'Pack', g.packPredTotal, staffPack);

// const orderActualDuration = calcActualDuration(g.orderActualList, g.depot, 'Order');
// const pickActualDuration  = calcActualDuration(g.pickActualList,  g.depot, 'Pick');
// const packActualDuration  = calcActualDuration(g.packActualList,  g.depot, 'Pack');

// // ▼ 予測・実績終了時間算出処理現状使わないのでコメントアウト
// const orderPred   = calcEndTime(g.depot, 'Order', g.orderPredTotal, staffOrder);
// const pickPred    = calcEndTime(g.depot, 'Pick',  g.pickPredTotal,  staffPick);
// const packPred    = calcEndTime(g.depot, 'Pack',  g.packPredTotal,  staffPack);
// const orderActual = maxActualTime(g.orderActualList);
// const pickActual  = maxActualTime(g.pickActualList);
// const packActual  = maxActualTime(g.packActualList);


//   result.push([
//     g.depot,
//     g.date,
//     courseCount,

//     staffOrder,
//     staffPick,
//     staffPack,

//     orderPredDuration,
//     pickPredDuration,
//     packPredDuration,

//     orderActualDuration,
//     pickActualDuration,
//     packActualDuration,

//     // ▼ 予測・実績終了時間算出処理現状使わないのでコメントアウト
//     orderPred.hhmm,
//     pickPred.hhmm,
//     packPred.hhmm,

//     orderActual.hhmm,
//     pickActual.hhmm,
//     packActual.hhmm,
//     /*
//     orderPred.hhmm,
//     pickPred.hhmm,
//     packPred.hhmm,

//     orderActual.hhmm,
//     pickActual.hhmm,
//     packActual.hhmm,

//     orderPred.min,
//     pickPred.min,
//     packPred.min,

//     orderActual.min,
//     pickActual.min,
//     packActual.min
//     */
//   ]);
// }

// // 別スプレッドシートへ出力
// const outputSheetId = '1djSGvGrIKJdbPcrxByYqB33SADzlLvwZz40sFhCxZkg'; // 実際のスプレッドシートIDをここに入力

// const headersRow = result[0];
// const dataRows   = result.slice(1);

// // Depot + date_only + Course_count で重複排除して追記→重複データは上書き更新に変更
// // 上書きしたいシート名に置き換えてください
// outputToExternalSheetWithHeaderAndUpsert(outputSheetId, headersRow, dataRows, [0, 1, 2]);
// // outputToExternalSheetWithHeaderAndDedup(outputSheetId, headersRow, dataRows, [0, 1, 2]);

// // 10秒待機
// Utilities.sleep(10000);

// SpreadsheetApp.getUi().alert("Complete Step3,next Step4");

//   // outputSheet.clear();
//   // outputSheet.getRange(1, 1, result.length, result[0].length).setValues(result);

//   // hh:mm, 分数両方のデータを出したので必要なデータだけDBに移すようにテーブル2つにしたがそれが無くなったので不要、コメントアウト
//   // // ▼ 指定処理追加（Revised_Prediction_Result シート用）
//   // const revisedSheet = ss.getSheetByName('Revised_Prediction_Result');
//   // if (!revisedSheet) throw new Error('Revised_Prediction_Resultシートが存在しません');

//   // // 1. データクリア（ヘッダー含む全行クリア）
//   // revisedSheet.clear();

//   // // 2. Prediction_Resultシートから全列コピー
//   // const predData = outputSheet.getDataRange().getValues();
//   // revisedSheet.getRange(1, 1, predData.length, predData[0].length).setValues(predData);
// }
