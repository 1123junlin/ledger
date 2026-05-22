// =============================================================
//  日常帳本記錄器 — Google Apps Script 後端
//  貼到 Google Apps Script 後，部署為 Web App（Anyone can access）
//  每次修改都要重新部署才會生效
// =============================================================

var SPREADSHEET_ID = ''; // ← 留空白：腳本自動抓綁定的試算表
                         //   或填你的 Spreadsheet ID（網址中的那串）

function doPost(e) {
  try {
    var data = JSON.parse(e.postData.contents);
    appendRow(data);
    return jsonOk({ status: 'ok', message: '記帳成功！' });
  } catch (err) {
    return jsonErr({ status: 'error', message: err.toString() });
  }
}

function doGet(e) {
  return jsonOk({ status: 'ok', message: '日常帳本 GAS 後端運作中' });
}

function appendRow(d) {
  var ss = SPREADSHEET_ID
    ? SpreadsheetApp.openById(SPREADSHEET_ID)
    : SpreadsheetApp.getActiveSpreadsheet();

  var sheetName = d.sheet || '2026';
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
    writeHeader(sheet);
  }

  // 把日期字串轉成 Excel 序號（對應你帳本的第一欄格式）
  var dateSerial = dateToSerial(d.date);

  var row = [
    dateSerial,          // A  日期序號
    d.memo || '',        // B  備忘

    // 飲食
    d.breakfast || 0,    // C  早餐
    d.breakfast_d || '', // D  早餐明細
    d.lunch || 0,        // E  午餐
    d.lunch_d || '',     // F  午餐明細
    d.dinner || 0,       // G  晚餐
    d.dinner_d || '',    // H  晚餐明細
    d.afternoon || 0,    // I  下午茶
    d.afternoon_d || '', // J  下午茶明細
    d.drink || 0,        // K  飲料
    d.drink_d || '',     // L  飲料明細
    d.night || 0,        // M  消夜
    d.night_d || '',     // N  消夜明細

    // 交通
    d.car || 0,          // O  汽車油錢
    d.car_d || '',       // P
    d.moto || 0,         // Q  機車油錢
    d.moto_d || '',      // R
    d.t1 || 0,           // S  交通1
    d.t1_d || '',        // T
    d.t2 || 0,           // U  交通2
    d.t2_d || '',        // V
    d.park1 || 0,        // W  停車費1
    d.park1_d || '',     // X
    d.park2 || 0,        // Y  停車費2
    d.park2_d || '',     // Z

    // 美安消費
    d.amway1 || 0,       // AA
    d.amway1_d || '',    // AB
    d.amway2 || 0,       // AC
    d.amway2_d || '',    // AD
    d.amway3 || 0,       // AE
    d.amway3_d || '',    // AF

    // 雜項
    d.misc1 || 0,        // AG
    d.misc1_d || '',     // AH
    d.misc2 || 0,        // AI
    d.misc2_d || '',     // AJ
    d.misc3 || 0,        // AK
    d.misc3_d || '',     // AL

    // 錢包餘額
    d.wallet || '',      // AM
  ];

  sheet.appendRow(row);

  // 把日期欄格式化為數字（避免顯示成字串）
  var lastRow = sheet.getLastRow();
  sheet.getRange(lastRow, 1).setNumberFormat('0');
}

// 日期字串 "2026-05-22" → Excel 序號（台灣帳本慣用）
function dateToSerial(dateStr) {
  if (!dateStr) return '';
  var parts = dateStr.split('-');
  var d = new Date(parseInt(parts[0]), parseInt(parts[1]) - 1, parseInt(parts[2]));
  // Excel epoch: 1900-01-00（實際上 1900-01-01 = 1）
  var epoch = new Date(1899, 11, 30);
  var diff = Math.round((d - epoch) / 86400000);
  return diff;
}

function writeHeader(sheet) {
  var headers = [
    '日期序號','備忘',
    '早餐','早餐明細','午餐','午餐明細','晚餐','晚餐明細',
    '下午茶','下午茶明細','飲料','飲料明細','消夜','消夜明細',
    '汽車油錢','汽車油錢明細','機車油錢','機車油錢明細',
    '交通1','交通1明細','交通2','交通2明細',
    '停車費1','停車費1明細','停車費2','停車費2明細',
    '美安消費1','美安消費1明細','美安消費2','美安消費2明細','美安消費3','美安消費3明細',
    '雜項1','雜項1明細','雜項2','雜項2明細','雜項3','雜項3明細',
    '錢包餘額'
  ];
  sheet.appendRow(headers);
}

function jsonOk(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

function jsonErr(obj) {
  return ContentService
    .createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}
