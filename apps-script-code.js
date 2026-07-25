// ===== 瑞安街電費 API =====
// 貼到 Google Sheet > 擴充功能 > Apps Script

var SS = SpreadsheetApp.getActiveSpreadsheet();

function fmtMonth(val) {
  if (val instanceof Date) {
    var y = val.getFullYear();
    var m = val.getMonth() + 1;
    return y + '-' + (m < 10 ? '0' : '') + m;
  }
  return String(val);
}

function doGet(e) {
  var action = (e && e.parameter && e.parameter.action) || 'getAll';
  var result;

  if (action === 'getAll') {
    result = getAllData();
  } else if (action === 'getMonth') {
    result = getMonthData(e.parameter.month);
  } else {
    result = { error: 'Unknown action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  var body = JSON.parse(e.postData.contents);
  var action = body.action;
  var result;

  if (action === 'saveMonth') {
    result = saveMonth(body.month, body.readings);
  } else if (action === 'saveSettings') {
    result = saveSettings(body.rooms);
  } else if (action === 'saveRate') {
    result = saveRate(body.rate);
  } else {
    result = { error: 'Unknown action' };
  }

  return ContentService
    .createTextOutput(JSON.stringify(result))
    .setMimeType(ContentService.MimeType.JSON);
}

function getAllData() {
  var settings = getSettingsSheet();
  var records = getRecordsSheet();
  var config = getConfigSheet();

  // 讀取房間設定
  var rooms = [];
  var sData = settings.getDataRange().getValues();
  for (var i = 1; i < sData.length; i++) {
    rooms.push({
      id: String(sData[i][0]),
      tenant: sData[i][1] || '',
      start: sData[i][2] || 0,
      start2: sData[i][3] || 0,
      dual: sData[i][4] === true || sData[i][4] === 'TRUE'
    });
  }

  // 讀取所有月度紀錄
  var months = {};
  var rData = records.getDataRange().getValues();
  for (var i = 1; i < rData.length; i++) {
    var mo = fmtMonth(rData[i][0]);
    var rid = String(rData[i][1]);
    if (!mo || !rid) continue;
    if (!months[mo]) months[mo] = {};
    months[mo][rid] = { r1: rData[i][2] || 0 };
    if (rData[i][3]) months[mo][rid].r2 = rData[i][3];
  }

  // 讀取設定值
  var rate = 6;
  var cData = config.getDataRange().getValues();
  for (var i = 1; i < cData.length; i++) {
    if (cData[i][0] === '每度電費') rate = cData[i][1];
  }

  return { rooms: rooms, months: months, rate: rate };
}

function getMonthData(month) {
  var records = getRecordsSheet();
  var rData = records.getDataRange().getValues();
  var result = {};
  for (var i = 1; i < rData.length; i++) {
    if (fmtMonth(rData[i][0]) === month) {
      var rid = String(rData[i][1]);
      result[rid] = { r1: rData[i][2] || 0 };
      if (rData[i][3]) result[rid].r2 = rData[i][3];
    }
  }
  return result;
}

function saveMonth(month, readings) {
  var records = getRecordsSheet();
  var rData = records.getDataRange().getValues();

  // 刪除該月份的舊資料
  var rowsToDelete = [];
  for (var i = rData.length - 1; i >= 1; i--) {
    if (fmtMonth(rData[i][0]) === month) rowsToDelete.push(i + 1);
  }
  for (var i = 0; i < rowsToDelete.length; i++) {
    records.deleteRow(rowsToDelete[i]);
  }

  // 寫入新資料（月份欄設為文字格式，避免被解析成日期）
  for (var i = 0; i < readings.length; i++) {
    var r = readings[i];
    var newRow = records.getLastRow() + 1;
    records.getRange(newRow, 1).setNumberFormat('@').setValue(month);
    records.getRange(newRow, 2).setValue(r.id);
    records.getRange(newRow, 3).setValue(r.r1 || '');
    records.getRange(newRow, 4).setValue(r.r2 || '');
  }

  // 排序
  var lastRow = records.getLastRow();
  if (lastRow > 1) {
    records.getRange(2, 1, lastRow - 1, 4).sort([
      { column: 1, ascending: true },
      { column: 2, ascending: true }
    ]);
  }

  return { success: true, month: month };
}

function saveSettings(rooms) {
  var settings = getSettingsSheet();

  for (var i = 0; i < rooms.length; i++) {
    var r = rooms[i];
    var row = i + 2; // 跳過標題行
    settings.getRange(row, 2).setValue(r.tenant || '');
    settings.getRange(row, 3).setValue(r.start || 0);
    if (r.start2) settings.getRange(row, 4).setValue(r.start2);
  }

  return { success: true };
}

function saveRate(rate) {
  var config = getConfigSheet();
  config.getRange(2, 2).setValue(rate);
  return { success: true, rate: rate };
}

function getSettingsSheet() { return SS.getSheetByName('設定'); }
function getRecordsSheet() { return SS.getSheetByName('月度紀錄'); }
function getConfigSheet() { return SS.getSheetByName('設定值'); }
