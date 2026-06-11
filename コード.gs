/**
 * バックエンド（GAS）のコード
 * このコードはスプレッドシートの「拡張機能」＞「Apps Script」を開き、
 * デフォルトで用意されている「コード.gs」に上書きで貼り付けます。
 */

// アプリの初期画面（index.html）を表示するおまじない
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('免税軽油 稼働・給油入力')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

// 画面を開いた時に、スプレッドシートからマスタデータを取得して画面に送る関数
function getInitialData() {
  // マスタデータがある別のスプレッドシートをIDで直接開く
  const masterSs = SpreadsheetApp.openById('1-HV-cb7tPiOvTD4nzKnlhfmPnfr59Kq2qxNNsSlblWU');
  
  // 1. 重機マスタの取得
  const machineSheet = masterSs.getSheetByName('重機マスタ');
  let machineList = [];
  if (machineSheet) {
    const data = machineSheet.getDataRange().getValues();
    // 1行目はヘッダーなので i=1 からスタート
    for (let i = 1; i < data.length; i++) {
      if (data[i][6] === '稼働中') { // G列（インデックス6）が「稼働中」のものだけ抽出
        machineList.push({ id: data[i][0], name: data[i][1] }); // A列(id), B列(name)
      }
    }
  }

  // 2. 現場マスタの取得
  const siteSheet = masterSs.getSheetByName('現場マスタ');
  let siteList = [];
  if (siteSheet) {
    const data = siteSheet.getDataRange().getValues();
    // 1行目はヘッダーなので i=1 からスタート
    for (let i = 1; i < data.length; i++) {
      const cellB = data[i][1]; // B列（インデックス1）
      const cellC = data[i][2]; // C列（インデックス2）
      
      // B列とC列の両方に値がある場合は「B列 C列」の形にする（カッコをなくす）
      if (cellB && cellC) {
        siteList.push(`${cellB} ${cellC}`);
      } else if (cellB) {
        // B列のみの場合はB列だけを表示
        siteList.push(cellB);
      }
    }
  }

  return { machines: machineList, sites: siteList };
}

// 対象月の入力済み重機リストを取得する関数
function getSubmittedMachines(targetMonth) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meterSheet = ss.getSheetByName('月次アワメータ入力');
  const recordSheet = ss.getSheetByName('稼働・給油記録');
  
  const submittedIds = new Set();
  
  // 1. 月次アワメータ入力から検索
  if (meterSheet) {
    const data = meterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      // 対象月が日付型(Date)に自動変換されている場合の対策
      let cellMonth = data[i][0];
      if (cellMonth instanceof Date) {
        cellMonth = Utilities.formatDate(cellMonth, Session.getScriptTimeZone(), 'yyyy-MM');
      } else {
        cellMonth = String(cellMonth).replace(/^'/, ''); // 先頭のシングルクォートがあれば除去
      }

      if (cellMonth === targetMonth) { 
        submittedIds.add(data[i][1]);   // B列(重機ID)を追加
      }
    }
  }

  // 2. 稼働・給油記録から検索
  if (recordSheet) {
    const data = recordSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      const dateVal = data[i][0];
      if (dateVal instanceof Date) {
        const dateStr = Utilities.formatDate(dateVal, Session.getScriptTimeZone(), 'yyyy-MM');
        if (dateStr === targetMonth) {  // 年月が一致するか
          submittedIds.add(data[i][1]); // B列(重機ID)を追加
        }
      }
    }
  }
  
  // Set（重複のない集合）を配列に変換して画面に返す
  return Array.from(submittedIds);
}

// ★追加：対象月・重機IDの「既存データ」をスプレッドシートから取得して画面に返す関数
function getExistingData(targetMonth, machineId) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const meterSheet = ss.getSheetByName('月次アワメータ入力');
  const recordSheet = ss.getSheetByName('稼働・給油記録');
  
  let result = {
    startMeter: '',
    endMeter: '',
    site: '',
    records: []
  };
  
  // 1. アワメータの取得
  if (meterSheet) {
    const data = meterSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let cellMonth = data[i][0];
      if (cellMonth instanceof Date) {
        cellMonth = Utilities.formatDate(cellMonth, Session.getScriptTimeZone(), 'yyyy-MM');
      } else {
        cellMonth = String(cellMonth).replace(/^'/, '');
      }
      
      if (cellMonth === targetMonth && data[i][1] === machineId) {
        result.startMeter = data[i][3] !== undefined ? data[i][3] : ''; // D列
        result.endMeter = data[i][4] !== undefined ? data[i][4] : '';   // E列
        break; // 同じ月・重機は1行のはずなので見つけたら終了
      }
    }
  }
  
  // 2. 稼働・給油記録の取得
  if (recordSheet) {
    const data = recordSheet.getDataRange().getValues();
    for (let i = 1; i < data.length; i++) {
      let cellDate = data[i][0];
      let cellDateStr = '';
      let cellMonth = '';
      
      if (cellDate instanceof Date) {
        cellDateStr = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM-dd');
        cellMonth = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM');
      } else {
        cellDateStr = String(cellDate);
        cellMonth = cellDateStr.substring(0, 7);
      }
      
      // 対象月と重機が一致する行を配列に追加
      if (cellMonth === targetMonth && data[i][1] === machineId) {
        result.records.push({
          date: cellDateStr,
          isWorking: data[i][3] === '〇',
          fuelAmount: data[i][4] !== undefined ? data[i][4] : ''
        });
        
        // 現場名（1ヶ月で基本同じ想定なので最初に見つけたものをセット）
        if (data[i][5] && !result.site) {
          result.site = data[i][5]; 
        }
      }
    }
  }
  
  return result;
}


// ★変更：重複エラーを無くし「洗い替え方式（既存データを消して新しく書き込む）」に変更
function saveData(payload) {
  const ss = SpreadsheetApp.getActiveSpreadsheet(); 
  const recordSheet = ss.getSheetByName('稼働・給油記録');
  const meterSheet = ss.getSheetByName('月次アワメータ入力');
  
  if (!recordSheet || !meterSheet) {
    return { success: false, message: '保存先のシートが見つかりません。「稼働・給油記録」と「月次アワメータ入力」シートを作成してください。' };
  }

  const records = payload.records;
  const hm = payload.hourMeter;
  const targetMonth = hm.targetMonth;
  const machineId = hm.machineId;

  // ============================================
  // 1. 洗い替え処理：既存データの削除
  // ============================================
  
  // アワメータの既存データ削除（行がずれないように下から上へループして削除）
  const meterData = meterSheet.getDataRange().getValues();
  for (let i = meterData.length - 1; i >= 1; i--) {
    let cellMonth = meterData[i][0];
    if (cellMonth instanceof Date) {
      cellMonth = Utilities.formatDate(cellMonth, Session.getScriptTimeZone(), 'yyyy-MM');
    } else {
      cellMonth = String(cellMonth).replace(/^'/, ''); 
    }

    if (cellMonth === targetMonth && meterData[i][1] === machineId) {
      meterSheet.deleteRow(i + 1); // deleteRowは1始まりのため i+1
    }
  }

  // 稼働記録の既存データ削除（行がずれないように下から上へループして削除）
  const recordData = recordSheet.getDataRange().getValues();
  for (let i = recordData.length - 1; i >= 1; i--) {
    let cellDate = recordData[i][0];
    let cellMonth = '';
    if (cellDate instanceof Date) {
      cellMonth = Utilities.formatDate(cellDate, Session.getScriptTimeZone(), 'yyyy-MM');
    } else {
      cellMonth = String(cellDate).substring(0, 7);
    }

    if (cellMonth === targetMonth && recordData[i][1] === machineId) {
      recordSheet.deleteRow(i + 1);
    }
  }

  // ============================================
  // 2. 新規データの追加（上書き保存）
  // ============================================
  
  let savedMsg = '完了しました！\n（既存データがある場合は上書きされました）\n';
  
  // アワメータの書き込み
  if (hm.startMeter || hm.endMeter) {
    const meterRowToInsert = [`'${hm.targetMonth}`, hm.machineId, hm.machineName, hm.startMeter, hm.endMeter];
    const meterStartRow = meterSheet.getLastRow() + 1;
    meterSheet.getRange(meterStartRow, 1, 1, meterRowToInsert.length).setValues([meterRowToInsert]);
    savedMsg += `・アワメータを保存しました。\n`;
  }
  
  // カレンダーデータの配列化
  const rowsToInsert = [];
  records.forEach(record => {
    rowsToInsert.push([
      record.date,          // A列: 日付
      record.machineId,     // B列: 重機ID
      record.machineName,   // C列: 機械名
      record.isWorking ? '〇' : '', // D列: 稼働
      record.fuelAmount,    // E列: 給油量
      record.site           // F列: 現場
    ]);
  });

  // カレンダーデータの書き込み
  if (rowsToInsert.length > 0) {
    const startRow = recordSheet.getLastRow() + 1;
    recordSheet.getRange(startRow, 1, rowsToInsert.length, rowsToInsert[0].length).setValues(rowsToInsert);
    savedMsg += `・${rowsToInsert.length}件の稼働/給油記録を保存しました。`;
  }

  return { success: true, message: savedMsg };
}