// =========================================================
// 1. 【一括保存】DBへ全現場のデータを保存・更新する関数
// =========================================================
function transferData() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const destSheet = ss.getSheetByName("DB"); 
  const sheetTransform = ss.getSheetByName('変換用シート'); 
  
  if (!sheetTransform || !destSheet) {
    Browser.msgBox("エラー", "「変換用シート」または「DB」シートが見つかりません。", Browser.Buttons.OK);
    return;
  }

  const targetSheetNames = [
    "受払及び稼働状況A3_現場A",
    "受払及び稼働状況A3_現場B",
    "受払及び稼働状況A3_現場C"
  ]; 

  let savedSites = [];
  let targetYmDisplay = "";
  let errorMessages = [];

  for (let s = 0; s < targetSheetNames.length; s++) {
    let currentSheet = ss.getSheetByName(targetSheetNames[s]);
    
    if (!currentSheet) {
      errorMessages.push("【" + targetSheetNames[s] + "】 シートが見つかりません。");
      continue; 
    }

    let sheetName = currentSheet.getName();
    let checkYm = currentSheet.getRange("BE6").getValue();
    let checkMachine = currentSheet.getRange("D8").getValue();
    
    if (!checkYm || !checkMachine) {
      errorMessages.push("【" + sheetName + "】 対象年月(BE6)または機械No(D8)が未入力のためスキップ。");
      continue; 
    }

    let convertResult = convertToDB(currentSheet, sheetTransform, sheetName);
    if (convertResult !== "success") {
      errorMessages.push("【" + sheetName + "】 " + convertResult);
      continue;
    }

    let rawTargetYm = sheetTransform.getRange("A4").getValue(); 
    if (!rawTargetYm) continue;
    
    let targetYm = "";
    if (rawTargetYm instanceof Date) {
      targetYm = rawTargetYm.getFullYear() + "/" + ("0" + (rawTargetYm.getMonth() + 1)).slice(-2);
    } else {
      targetYm = String(rawTargetYm).replace(/-/g, "/").substring(0, 7);
    }

    if (targetYmDisplay === "") targetYmDisplay = targetYm;

    removeOldMonthAndSiteData(destSheet, targetYm, sheetName, 12); 

    const lastRow = sheetTransform.getRange("B" + sheetTransform.getMaxRows()).getNextDataCell(SpreadsheetApp.Direction.UP).getRow();
    
    if (lastRow >= 4) {
      const data = sheetTransform.getRange(4, 1, lastRow - 3, 12).getValues();
      const filteredData = data.filter(row => row[1] !== "");
      
      if (filteredData.length > 0) {
        destSheet.getRange(destSheet.getLastRow() + 1, 1, filteredData.length, 12).setValues(filteredData);
        savedSites.push(sheetName);
      }
    }
  }
  
  if (savedSites.length > 0) {
    let msg = targetYmDisplay + " のデータ\\nをデータベースに一括保存・更新しました！";
    if (errorMessages.length > 0) {
      msg += "\\n\\n⚠️ 一部スキップされたシートがあります\\n" + errorMessages.join("\\n");
    }
    Browser.msgBox("一括保存完了", msg, Browser.Buttons.OK);
  } else {
    let msg = "保存するデータがありませんでした。";
    if (errorMessages.length > 0) {
      msg += "\\n\\n【詳細】\\n" + errorMessages.join("\\n");
    }
    Browser.msgBox("確認", msg, Browser.Buttons.OK);
  }
}

// --- 同年月・同シートの古いデータを削除するための補助関数 ---
function removeOldMonthAndSiteData(sheet, targetYm, targetSheetName, colCount) {
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return; 
  
  const data = sheet.getRange(2, 1, lastRow - 1, colCount).getValues();
  const newData = [];
  
  for (let i = 0; i < data.length; i++) {
    let val = data[i][0]; 
    let rowYm = "";
    
    if (val) {
      if (val instanceof Date) {
        let y = val.getFullYear();
        let m = ("0" + (val.getMonth() + 1)).slice(-2);
        rowYm = y + "/" + m; 
      } else {
        rowYm = String(val).replace(/-/g, "/").substring(0, 7);
      }
    }
    
    let rowSheetName = String(data[i][11]).trim(); 
    let isRowEmpty = String(data[i].join("")).trim() === "";
    
    let isTargetRow = (rowYm === targetYm) && (rowSheetName === String(targetSheetName).trim());
    
    if (!isTargetRow && !isRowEmpty) {
      newData.push(data[i]);
    }
  }
  
  sheet.getRange(2, 1, lastRow - 1, colCount).clearContent();
  if (newData.length > 0) {
    sheet.getRange(2, 1, newData.length, colCount).setValues(newData);
  }
}

// --- 変換用シート自動生成プログラム（補助関数） ---
function convertToDB(sheetA3, sheetTransform, sheetName) {
  const data = sheetA3.getDataRange().getValues();
  
  if (data.length < 12) return "入力データがありません。";

  const targetYearMonth = data[5][56]; 
  const baseDateObj = data[7][64];     
  
  if (!(baseDateObj instanceof Date)) {
    return "BM8セルが正しい日付形式ではありません。（数式エラーの可能性があります）";
  }
  
  const baseDate = new Date(baseDateObj);
  const baseMonth = baseDate.getMonth();

  let outputData = [];
  const maxMachines = 7; 

  for (let i = 0; i < maxMachines; i++) {
    let offset = i * 8; 
    let machineNum  = data[7][3 + offset]; 
    let machineName = data[7][5 + offset]; 
    let location    = data[8][3 + offset]; 

    if (!machineNum) continue;

    for (let day = 0; day < 31; day++) {
      let rowIdx = (day === 0) ? 11 : 12 + day;
      
      if (rowIdx >= data.length) break;

      let currentDate = new Date(baseDate.getFullYear(), baseMonth, baseDate.getDate() + day);

      if (currentDate.getMonth() !== baseMonth) break;

      let dateStr = Utilities.formatDate(currentDate, Session.getScriptTimeZone(), "yyyy/MM/dd");

      let transactionVol = data[rowIdx][1];  
      let taxFreeVol     = data[rowIdx][59]; 
      let stockVol       = data[rowIdx][61]; 

      let usageVol  = data[rowIdx][3 + offset]; 
      let status    = data[rowIdx][7 + offset]; 
      let hourMeter = data[rowIdx][8 + offset]; 

      outputData.push([
        targetYearMonth, dateStr, machineNum, machineName, location,
        transactionVol, usageVol, status, hourMeter, stockVol, taxFreeVol, sheetName
      ]);
    }
  }

  const maxRows = sheetTransform.getMaxRows();
  if (maxRows > 3) {
    // ★修正: A〜M列（13列分）をクリアするように変更
    sheetTransform.getRange(4, 1, maxRows - 3, 13).clearContent();
  }

  if (outputData.length > 0) {
    sheetTransform.getRange(4, 1, outputData.length, 12).setValues(outputData);
  } else {
    return "機械番号などが入力されていないため、データを作成できませんでした。";
  }

  // ★大修正：L列(シート名)と被らないように、計算式を【M列】に移動しました！
  sheetTransform.getRange("M3").setValue("前月末在庫");
  sheetTransform.getRange("M4").setFormula('=IFERROR(INDEX(SORT(FILTER(DB!J:J, DB!B:B < DATE(YEAR(B4), MONTH(B4), 1)), FILTER(DB!B:B, DB!B:B < DATE(YEAR(B4), MONTH(B4), 1)), FALSE), 1, 1), 0)');
  
  return "success";
}

// =========================================================
// 2. 【一括クリア】シートの入力内容をリセットする関数
// =========================================================
function clearA3Inputs() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  
  const targetSheetNames = [
    "受払及び稼働状況A3_現場A",
    "受払及び稼働状況A3_現場B",
    "受払及び稼働状況A3_現場C"
  ]; 
  
  const response = Browser.msgBox(
    "一括クリアの確認", 
    "現場A・B・Cすべての対象年月（BE6）と、日々の入力データ（B12〜BG43）を【一括でクリア】しますか？\\n\\n※重機名や現場名（8,9行目）は残ります。\\n※この操作は取り消しできません。", 
    Browser.Buttons.OK_CANCEL
  );
  
  if (response !== "ok") {
    return; 
  }

  let clearedSheets = [];
  
  for (let i = 0; i < targetSheetNames.length; i++) {
    let sheet = ss.getSheetByName(targetSheetNames[i]);
    if (sheet) {
      sheet.getRange("BE6").clearContent();
      sheet.getRange("B12:BG43").clearContent();
      clearedSheets.push(targetSheetNames[i]);
    }
  }
  
  if (clearedSheets.length > 0) {
    Browser.msgBox(
      "完了", 
      "以下のシートの入力をクリアしました！\\n\\n・" + clearedSheets.join("\\n・") + "\\n\\n新しい月の作業を始めてください。", 
      Browser.Buttons.OK
    );
  } else {
    Browser.msgBox("確認", "クリア対象のシートが見つかりませんでした。", Browser.Buttons.OK);
  }
}

// =========================================================
// 3. 【一括読込】データ無しは白紙にリセットする完全同期版関数
// =========================================================
function loadFromDB() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const activeSheet = ss.getActiveSheet(); 
  const dbSheet = ss.getSheetByName("DB");

  if (!dbSheet) {
    Browser.msgBox("エラー", "「DB」シートが見つかりません。", Browser.Buttons.OK);
    return;
  }

  const masterYmStr = activeSheet.getRange("BE6").getValue();
  if (!masterYmStr) {
    Browser.msgBox("確認", "対象年月（BE6）を選択してから実行してください。", Browser.Buttons.OK);
    return;
  }

  let checkTargetYm = "";
  if (masterYmStr instanceof Date) {
    checkTargetYm = masterYmStr.getFullYear() + "/" + ("0" + (masterYmStr.getMonth() + 1)).slice(-2);
  } else {
    checkTargetYm = String(masterYmStr).replace(/-/g, "/").substring(0, 7);
  }

  const targetSheetNames = [
    "受払及び稼働状況A3_現場A",
    "受払及び稼働状況A3_現場B",
    "受払及び稼働状況A3_現場C"
  ];
  
  const dbData = dbSheet.getDataRange().getValues();
  if (dbData.length < 2) {
    Browser.msgBox("確認", "DBにデータがありません。", Browser.Buttons.OK);
    return;
  }

  let restoreTasks = []; 
  let messageLines = []; 

  for (let s = 0; s < targetSheetNames.length; s++) {
    let sheetName = targetSheetNames[s];
    let sheet = ss.getSheetByName(sheetName);
    
    if (!sheet) continue;

    let foundData = [];
    let uniqueMachines = []; 
    let actualSiteName = ""; 

    for (let i = 1; i < dbData.length; i++) {
      let row = dbData[i];
      let rowYm = "";
      if (row[0] instanceof Date) {
        rowYm = row[0].getFullYear() + "/" + ("0" + (row[0].getMonth() + 1)).slice(-2);
      } else {
        rowYm = String(row[0]).replace(/-/g, "/").substring(0, 7);
      }
      
      let rowSheetName = String(row[11]).trim();

      if (rowYm === checkTargetYm && rowSheetName === sheetName) {
        foundData.push(row);
        
        if (actualSiteName === "") {
          actualSiteName = String(row[4]).trim();
        }
        
        let mNum = row[2];
        if (mNum !== "" && uniqueMachines.indexOf(mNum) === -1) {
          uniqueMachines.push(mNum);
        }
      }
    }

    if (foundData.length > 0) {
      restoreTasks.push({
        sheet: sheet,
        sheetName: sheetName,
        siteName: actualSiteName, 
        foundData: foundData,
        uniqueMachines: uniqueMachines 
      });
      messageLines.push("・" + sheetName + "：" + foundData.length + "件 見つかりました");
    } else {
      restoreTasks.push({
        sheet: sheet,
        sheetName: sheetName,
        siteName: "", 
        foundData: [],
        uniqueMachines: [] 
      });
      messageLines.push("・" + sheetName + "：データなし（白紙にリセットされます）");
    }
  }

  const response = Browser.msgBox(
    "一括復元の確認",
    "DBの検索結果は以下の通りです。\\n\\n" + messageLines.join("\\n") + "\\n\\n画面を当時の状態に完全同期します。\\n※データ無しのシートは現在の入力内容が【すべて消去】され白紙になります。\\n※この操作は取り消しできません。実行しますか？",
    Browser.Buttons.OK_CANCEL
  );

  if (response !== "ok") {
    return;
  }

  let successCount = 0;
  let errorMsgs = [];

  for (let t = 0; t < restoreTasks.length; t++) {
    let task = restoreTasks[t];
    let sheet = task.sheet;
    let foundData = task.foundData;
    let uniqueMachines = task.uniqueMachines;
    let taskSiteName = task.siteName;

    let sheetMachines = [];
    
    for (let m = 0; m < 7; m++) {
      let mNum = (m < uniqueMachines.length) ? uniqueMachines[m] : "";
      let sName = (m < uniqueMachines.length) ? taskSiteName : "";
      
      let mNumValue = (mNum !== "" && !isNaN(mNum)) ? Number(mNum) : mNum;
      
      sheet.getRange(8, 4 + m * 8).setValue(mNumValue);  
      sheet.getRange(9, 4 + m * 8).setValue(sName); 
      
      sheetMachines.push(mNum);
    }

    let restoreData = new Array(32).fill(null).map(() => new Array(58).fill(""));
    
    for (let r = 0; r < 32; r++) {
      if (r === 1) continue; 
      for (let m = 0; m < 7; m++) {
        restoreData[r][6 + m * 8] = false; 
      }
    }

    for (let i = 0; i < foundData.length; i++) {
      let row = foundData[i];
      let dateVal = row[1];
      let day = 1;
      
      if (dateVal instanceof Date) {
        day = dateVal.getDate();
      } else {
        let parts = String(dateVal).split("/");
        if (parts.length >= 3) day = parseInt(parts[2], 10);
      }

      let arrRowIdx = (day === 1) ? 0 : day;
      if (arrRowIdx < 0 || arrRowIdx > 31) continue;

      let machineNum = row[2];
      let transactionVol = row[5];
      let usageVol = row[6];
      let status = row[7];
      let hourMeter = row[8];

      if (transactionVol !== "") {
        restoreData[arrRowIdx][0] = transactionVol;
      }

      let mIndex = sheetMachines.indexOf(machineNum);
      if (mIndex !== -1) {
        restoreData[arrRowIdx][2 + mIndex * 8] = usageVol; 
        restoreData[arrRowIdx][6 + mIndex * 8] = (status === true || status === "true" || status === "〇") ? true : false; 
        restoreData[arrRowIdx][7 + mIndex * 8] = hourMeter;
      }
    }

    try {
      sheet.getRange(12, 2, 32, 58).setValues(restoreData);
      sheet.getRange("BE6").setValue(masterYmStr); 
      successCount++;
    } catch (e) {
      errorMsgs.push("【" + task.sheetName + "】 " + e.message);
    }
  }

  if (errorMsgs.length > 0) {
    Browser.msgBox("完了（一部エラー）", successCount + "シートの同期が完了しましたが、一部でエラーが発生しました。\\n\\n" + errorMsgs.join("\\n"), Browser.Buttons.OK);
  } else {
    Browser.msgBox("同期完了", successCount + "シートの完全同期が完了しました！\\n（データが無かったシートは白紙にリセットされました）", Browser.Buttons.OK);
  }
}