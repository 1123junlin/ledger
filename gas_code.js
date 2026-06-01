const SPREADSHEET_ID = '1i1vX2YZVXcOJ2tbm9DELEWjtAliObX37PzPxtlvPq8c';

function doGet(e) {
  const action = e.parameter.action || '';
  const sheetName = e.parameter.sheet || '2026';

  if (action === 'month') {
    return handleMonth(e, sheetName);
  } else if (action === 'balance') {
    return handleBalance(sheetName);
  }

  return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: 'unknown action' }))
    .setMimeType(ContentService.MimeType.JSON);
}

function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const sheetName = data.sheet || '2026';
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(sheetName);

    // 找下一筆空白列（從第6列開始）
    let targetRow = 6;
    const lastRow = ws.getLastRow();
    for (let r = 6; r <= lastRow + 1; r++) {
      if (!ws.getRange(r, 1).getValue()) {
        targetRow = r;
        break;
      }
    }

    // 寫入資料
    ws.getRange(targetRow, 1).setValue(data.date);         // A: 日期
    ws.getRange(targetRow, 2).setValue(data.memo || '');   // B: 備忘
    ws.getRange(targetRow, 3).setValue(data.wallet || 0);  // C: 錢包餘額

    // 飲食 (D~S)  ← D欄刪除後全部前移一格
    ws.getRange(targetRow, 4).setValue(data.breakfast || 0);    // D: 早餐
    ws.getRange(targetRow, 5).setValue(data.breakfast_d || ''); // E: 早餐明細
    ws.getRange(targetRow, 6).setValue(data.lunch || 0);        // F: 午餐
    ws.getRange(targetRow, 7).setValue(data.lunch_d || '');     // G: 午餐明細
    ws.getRange(targetRow, 8).setValue(data.dinner || 0);       // H: 晚餐
    ws.getRange(targetRow, 9).setValue(data.dinner_d || '');    // I: 晚餐明細
    ws.getRange(targetRow, 10).setValue(data.afternoon || 0);   // J: 下午茶
    ws.getRange(targetRow, 11).setValue(data.afternoon_d || '');// K: 下午茶明細
    ws.getRange(targetRow, 12).setValue(data.night || 0);       // L: 消夜
    ws.getRange(targetRow, 13).setValue(data.night_d || '');    // M: 消夜明細
    ws.getRange(targetRow, 14).setValue(data.drink || 0);       // N: 飲料
    ws.getRange(targetRow, 15).setValue(data.drink_d || '');    // O: 飲料明細

    // 交通 (T~AK)
    ws.getRange(targetRow, 20).setValue(data.car || 0);         // T: 汽車油錢
    ws.getRange(targetRow, 21).setValue(data.moto || 0);        // U: 機車油錢
    ws.getRange(targetRow, 22).setValue(data.t1 || 0);          // V: 交通1
    ws.getRange(targetRow, 23).setValue(data.t1_d || '');       // W: 交通明細1
    ws.getRange(targetRow, 24).setValue(data.t2 || 0);          // X: 交通2
    ws.getRange(targetRow, 25).setValue(data.t2_d || '');       // Y: 交通明細2
    ws.getRange(targetRow, 32).setValue(data.park1 || 0);       // AF: 停車費1
    ws.getRange(targetRow, 33).setValue(data.park1_d || '');    // AG: 停車費1明細
    ws.getRange(targetRow, 34).setValue(data.park2 || 0);       // AH: 停車費2
    ws.getRange(targetRow, 35).setValue(data.park2_d || '');    // AI: 停車費2明細

    // 美安消費 (AL~AQ)
    ws.getRange(targetRow, 38).setValue(data.amway1 || 0);      // AL: 美安消費1
    ws.getRange(targetRow, 39).setValue(data.amway1_d || '');   // AM: 美安消費1明細
    ws.getRange(targetRow, 40).setValue(data.amway2 || 0);      // AN: 美安消費2
    ws.getRange(targetRow, 41).setValue(data.amway2_d || '');   // AO: 美安消費2明細
    ws.getRange(targetRow, 42).setValue(data.amway3 || 0);      // AP: 美安消費3
    ws.getRange(targetRow, 43).setValue(data.amway3_d || '');   // AQ: 美安消費3明細

    // 雜項 (AR~BA)
    ws.getRange(targetRow, 44).setValue(data.misc1 || 0);       // AR: 雜項1
    ws.getRange(targetRow, 45).setValue(data.misc1_d || '');    // AS: 雜項1明細
    ws.getRange(targetRow, 46).setValue(data.misc2 || 0);       // AT: 雜項2
    ws.getRange(targetRow, 47).setValue(data.misc2_d || '');    // AU: 雜項2明細
    ws.getRange(targetRow, 48).setValue(data.misc3 || 0);       // AV: 雜項3
    ws.getRange(targetRow, 49).setValue(data.misc3_d || '');    // AW: 雜項3明細

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', row: targetRow }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 讀取當月資料（給圖表用）────────────────────────────
function handleMonth(e, sheetName) {
  try {
    const year  = parseInt(e.parameter.year);
    const month = parseInt(e.parameter.month);
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(sheetName);
    const lastRow = ws.getLastRow();

    const rows = [];
    for (let r = 6; r <= lastRow; r++) {
      const dateVal = ws.getRange(r, 1).getValue();
      if (!dateVal) continue;
      const d = new Date(dateVal);
      if (d.getFullYear() !== year || d.getMonth() + 1 !== month) continue;

      const fmt = Utilities.formatDate(d, 'Asia/Taipei', 'MM/dd');
      rows.push({
        date:      fmt,
        breakfast: ws.getRange(r, 4).getValue()  || 0,  // D
        lunch:     ws.getRange(r, 6).getValue()  || 0,  // F
        dinner:    ws.getRange(r, 8).getValue()  || 0,  // H
        afternoon: ws.getRange(r, 10).getValue() || 0,  // J
        night:     ws.getRange(r, 12).getValue() || 0,  // L
        drink:     ws.getRange(r, 14).getValue() || 0,  // N
        car:       ws.getRange(r, 20).getValue() || 0,  // T
        moto:      ws.getRange(r, 21).getValue() || 0,  // U
        t1:        ws.getRange(r, 22).getValue() || 0,  // V
        t2:        ws.getRange(r, 24).getValue() || 0,  // X
        park1:     ws.getRange(r, 32).getValue() || 0,  // AF
        park2:     ws.getRange(r, 34).getValue() || 0,  // AH
        amway1:    ws.getRange(r, 38).getValue() || 0,  // AL
        amway2:    ws.getRange(r, 40).getValue() || 0,  // AN
        amway3:    ws.getRange(r, 42).getValue() || 0,  // AP
        misc1:     ws.getRange(r, 44).getValue() || 0,  // AR
        misc2:     ws.getRange(r, 46).getValue() || 0,  // AT
        misc3:     ws.getRange(r, 48).getValue() || 0,  // AV
      });
    }

    return ContentService.createTextOutput(JSON.stringify({ status: 'ok', rows }))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

// ── 讀取最新一筆帳戶餘額（給 Dashboard 用）─────────────
function handleBalance(sheetName) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const ws = ss.getSheetByName(sheetName);
    const lastRow = ws.getLastRow();

    // 從最後一列往上找，找到 A 欄有日期的那列
    let targetRow = null;
    for (let r = lastRow; r >= 6; r--) {
      if (ws.getRange(r, 1).getValue()) {
        targetRow = r;
        break;
      }
    }

    if (!targetRow) {
      return ContentService.createTextOutput(JSON.stringify({ status: 'ok', found: false }))
        .setMimeType(ContentService.MimeType.JSON);
    }

    const dateVal = ws.getRange(targetRow, 1).getValue();
    const dateStr = Utilities.formatDate(new Date(dateVal), 'Asia/Taipei', 'MM/dd');

    const result = {
      status:  'ok',
      found:   true,
      date:    dateStr,
      wallet:  ws.getRange(targetRow, 3).getValue(),   // C: 錢包
      land:    ws.getRange(targetRow, 81).getValue(),  // CD→CC: 土地銀行
      esun:    ws.getRange(targetRow, 82).getValue(),  // CE→CD: 玉山銀行
      taishin: ws.getRange(targetRow, 83).getValue(),  // CF→CE: 台新銀行
      sinopac: ws.getRange(targetRow, 84).getValue(),  // CG→CF: 永豐銀行
      ctbc:    ws.getRange(targetRow, 85).getValue(),  // CH→CG: 中信
      linepay: ws.getRange(targetRow, 86).getValue(),  // CI→CH: LINE PAY
      pxmart:  ws.getRange(targetRow, 87).getValue(),  // CJ→CI: 全聯禮券
      seven:   ws.getRange(targetRow, 88).getValue(),  // CK→CJ: 7-11禮券
      jkopay:  ws.getRange(targetRow, 89).getValue(),  // CL→CK: 街口
    };

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);

  } catch(err) {
    return ContentService.createTextOutput(JSON.stringify({ status: 'error', message: err.toString() }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
