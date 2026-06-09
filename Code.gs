// #### 檔案：Code.gs

function doGet() {
  return HtmlService.createTemplateFromFile('Index')
    .evaluate()
    .setTitle('四年甲班班務系統')
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

function include(filename) {
  return HtmlService.createHtmlOutputFromFile(filename).getContent();
}

function forceDateString(val) {
  if (!val) return '';
  if (val instanceof Date) {
    const offset = val.getTimezoneOffset() * 60000;
    return new Date(val.getTime() - offset).toISOString().split('T')[0];
  }
  return String(val).replace(/'/g, "").trim();
}

// --- 通用：儲存功能 ---
function saveDataGeneric(data, sheetName) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName(sheetName);
    if (!ws) {
      ws = ss.insertSheet(sheetName);
      ws.appendRow(['時間戳記', '日期', '項目名稱', '缺交座號', '備註', '狀態']);
    }

    const dateStr = "'" + data.dateStr; 
    const timestamp = new Date();
    const status = (data.missingSeats && data.missingSeats.length > 0) ? '未完成' : '全班完成';

    if (data.rowIndex) {
      ws.getRange(data.rowIndex, 2).setValue(dateStr);
      ws.getRange(data.rowIndex, 3).setValue(data.taskName);
      ws.getRange(data.rowIndex, 4).setValue(data.missingSeats);
      ws.getRange(data.rowIndex, 6).setValue(status);
      const msg = (status === '全班完成') ? '🎉 太棒了！全班已完成，該項目已結案移除。' : '✅ 狀態已更新！';
      return { success: true, message: msg };
    } else {
      ws.appendRow([timestamp, dateStr, data.taskName, data.missingSeats, data.note, status]);
      return { success: true, message: '✅ 新項目已建立！' };
    }
  } catch (e) {
    return { success: false, message: "後端錯誤：" + e.toString() };
  }
}

function saveHomeworkData(data) { return saveDataGeneric(data, '作業追蹤'); }
function saveAdminData(data) { return saveDataGeneric(data, '行政事務'); }

// --- 通用：讀取功能 ---
function getListGeneric(sheetName) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName(sheetName);
  if (!ws) return { today: [], history: [] };

  const data = ws.getDataRange().getValues();
  const d = new Date();
  const offset = d.getTimezoneOffset() * 60000;
  const todayStr = new Date(d.getTime() - offset).toISOString().split('T')[0];
  
  let todayList = [];
  let historyList = [];
  
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const rowDate = forceDateString(row[1]);
    const missing = String(row[3]);
    
    if (missing && missing.trim() !== '') {
      const item = {
        rowIndex: i + 1, 
        date: rowDate,
        taskName: row[2],
        missingSeats: missing,
        count: missing.split(',').length
      };
      
      if (rowDate === todayStr) {
        todayList.push(item);
      } else {
        historyList.push(item);
      }
    }
  }
  historyList.reverse();
  todayList.reverse();
  return { today: todayList, history: historyList };
}

function getHomeworkList() { return getListGeneric('作業追蹤'); }
function getAdminList() { return getListGeneric('行政事務'); }

// --- 取得欠債大戶名單 ---
function getAtRiskStudents() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('作業追蹤'); 
  if (!ws) return [];

  const data = ws.getDataRange().getValues();
  let studentMap = {}; 

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const date = forceDateString(row[1]); 
    const taskName = row[2];
    const missingStr = String(row[3]); 

    if (missingStr && missingStr.trim() !== '') {
      const seats = missingStr.split(',');
      seats.forEach(s => {
        const seatNum = s.trim();
        if (!seatNum) return;
        if (!studentMap[seatNum]) studentMap[seatNum] = [];
        studentMap[seatNum].push(`[${date}] ${taskName}`);
      });
    }
  }

  let result = [];
  for (const seat in studentMap) {
    if (studentMap[seat].length >= 2) {
      result.push({ seat: seat, count: studentMap[seat].length, tasks: studentMap[seat] });
    }
  }
  result.sort((a, b) => parseInt(a.seat) - parseInt(b.seat));
  return result;
}

// --- 🌟 新增功能：清理 30 天前的舊資料 ---
function clearOldData() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheetsToClean = ['作業追蹤', '行政事務'];
    let totalDeleted = 0;
    
    // 計算 30 天前的日期
    const thirtyDaysAgo = new Date();
    thirtyDaysAgo.setDate(thirtyDaysAgo.getDate() - 30);
    
    sheetsToClean.forEach(sheetName => {
      const ws = ss.getSheetByName(sheetName);
      if (!ws) return;
      
      const data = ws.getDataRange().getValues();
      if (data.length <= 1) return; // 只有標題列
      
      let newData = [data[0]]; // 保留標題列
      let deletedCount = 0;
      
      for (let i = 1; i < data.length; i++) {
        let rowDate;
        if (data[i][1] instanceof Date) {
          rowDate = data[i][1];
        } else {
          rowDate = new Date(String(data[i][1]).replace(/'/g, "").trim());
        }
        
        // 如果日期大於等於 30 天前，或者是無效日期(防呆)，就保留
        if (rowDate >= thirtyDaysAgo || isNaN(rowDate.getTime())) {
          newData.push(data[i]);
        } else {
          // 否則就是過期，準備刪除
          deletedCount++;
        }
      }
      
      if (deletedCount > 0) {
        ws.clearContents(); // 清空內容
        ws.getRange(1, 1, newData.length, newData[0].length).setValues(newData); // 貼回保留的資料
        totalDeleted += deletedCount;
      }
    });
    
    if (totalDeleted > 0) {
      return { success: true, message: `✅ 大掃除完成！共清除了 ${totalDeleted} 筆超過一個月的舊紀錄，系統速度已恢復。` };
    } else {
      return { success: true, message: `✨ 目前系統很乾淨，沒有超過一個月的舊紀錄需要清理。` };
    }
  } catch (e) {
    return { success: false, message: "清理失敗：" + e.toString() };
  }
}

// --- 聯絡簿 ---
function saveCommData(data) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let ws = ss.getSheetByName('聯絡簿');
    if (!ws) {
      ws = ss.insertSheet('聯絡簿');
      ws.appendRow(['時間戳記', '日期', '聯絡事項', '早自修事項']);
    }
    const dateStr = "'" + data.dateStr;
    const timestamp = new Date();
    const allData = ws.getDataRange().getValues();
    let rowIndexToUpdate = -1;
    for (let i = 1; i < allData.length; i++) {
      let dbDate = forceDateString(allData[i][1]); 
      if (dbDate === data.dateStr) { rowIndexToUpdate = i + 1; break; }
    }
    const rowData = [timestamp, dateStr, data.contactItems.join('\n'), data.studyItems.join('\n')];
    if (rowIndexToUpdate > 0) ws.getRange(rowIndexToUpdate, 1, 1, 4).setValues([rowData]);
    else ws.appendRow(rowData);
    return { success: true, message: '✅ 黑板已更新！' };
  } catch (e) {
    return { success: false, message: "後端錯誤：" + e.toString() };
  }
}

function getCommData(dateStr) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const ws = ss.getSheetByName('聯絡簿');
  if (!ws) return { contact: '', study: '' };
  const data = ws.getDataRange().getValues();
  for (let i = data.length - 1; i >= 1; i--) {
    let dbDate = forceDateString(data[i][1]); 
    if (dbDate === dateStr) { return { contact: data[i][2], study: data[i][3] }; }
  }
  return { contact: '', study: '' };
}
