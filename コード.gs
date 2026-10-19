// ==========================================
// 1. 定数・設定
// ==========================================
const SPREADSHEET_IDS = {
  KYOTEN: "14RbsDcv0nXfEwweki8-9cK3lQUg1XUuhozLNF9u2qAs",
  AGENCY: "1Fd8uOCE1SKvLCIPjZZ7sE2rFsQFOhjVjHqoaqs-pXqE",
  MONTHLY_1: "1JhLJuxkp6T5hEkfultUE57pmRXBI_lkN6ROp2-CuhCc",
  MONTHLY_2: "10Z8jg4o7Ri9Ggf7u_jauyDDZa6bVnc-iuuuqmkNDafw",
  SHIFT_NEW: "1LFVmqwJU-WQbNOuSai8k72bSK790Eq_lBZeNKmYu8co",
  CALENDAR: "1WlmirSDOPnIcV2cwY5ClXkMWBFMM-4zrAw4XFDNUPGw"
};

// ==========================================
// 2. メニューとUI連携
// ==========================================
function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('勤務くんシステム')
    .addItem('集計UIを開く (シート出力)', 'showDialog')
    .addItem('医師検索', 'showSearchDialog')
    .addToUi();
}

function showDialog() {
  const html = HtmlService.createTemplateFromFile('Index').evaluate()
      .setTitle('勤務実績集計 (勤務くん)')
      .setWidth(450)
      .setHeight(450); 
  SpreadsheetApp.getUi().showSidebar(html);
}

function cancelProcess() {
  CacheService.getUserCache().put('cancel_flag', 'true', 300);
  return "中止命令を送信しました...";
}

function getKyotenList() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.KYOTEN);
    const data = ss.getSheetByName('拠点名').getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    const iName = headers.indexOf("正規記載");
    const iGroup = headers.indexOf("拠点グループ");
    const iOpenDate = headers.indexOf("開院日");

    if (iName === -1 || iGroup === -1 || iOpenDate === -1) {
      throw new Error("必要なヘッダー（正規記載, 拠点グループ, 開院日）が見つかりません。");
    }

    const list = [];
    const today = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[iName];
      const group = row[iGroup];
      const openDateStr = row[iOpenDate];
      if (!openDateStr) continue;
      
      const openDate = new Date(openDateStr);
      if (group !== "MQC" && !isNaN(openDate.getTime()) && openDate <= today) {
        if (name === "北葛西" || name === "亀有") {
          list.push({ id: `${name}_内科`, label: `${name}（内科）` });
          list.push({ id: `${name}_小児科`, label: `${name}（小児科）` });
        } else {
          list.push({ id: name, label: name });
        }
      }
    }
    
    // 重複を排除して返す
    const uniqueList = [];
    const seen = new Set();
    for (const item of list) {
      if (!seen.has(item.id)) {
        seen.add(item.id);
        uniqueList.push(item);
      }
    }
    return uniqueList;

  } catch(e) {
    throw new Error("拠点マスタ読込エラー: " + e.message);
  }
}

// ==========================================
// 3. 最終勤務日の動的算出ロジック
// ==========================================
function getMaxDateForKyoten(targetKyotenId) {
  if (!targetKyotenId) return null;
  CacheService.getUserCache().remove('cancel_flag'); 

  const [targetKyotenName, targetShinryoka] = targetKyotenId.split('_');
  let maxDate = null;
  
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.SHIFT_NEW);
    const sheets = ss.getSheets();
    
    sheets.forEach(sheet => {
      const sName = sheet.getName();
      if (sName !== "確定シフト" && sName !== "応募シフト") return;
      
      const isObo = (sName === "応募シフト");
      const data = sheet.getDataRange().getValues();
      if (data.length < 2) return;
      
      const headers = data[0];
      const iKyoten = headers.indexOf("クリニック名");
      const iShinryo = headers.indexOf("診療科");
      const iDate = headers.indexOf("勤務日");
      const iWage = headers.indexOf("時給合計");
      const iRemarks = headers.indexOf("業務内容（備考）");

      if (iKyoten === -1 || iDate === -1) return;

      for (let i = 1; i < data.length; i++) {
        if (i % 100 === 0 && CacheService.getUserCache().get('cancel_flag') === 'true') {
          throw new Error("処理が中止されました。");
        }

        const row = data[i];
        if (row[iKyoten] !== targetKyotenName) continue;
        if (targetShinryoka && row[iShinryo] !== targetShinryoka) continue;

        const totalWage = iWage !== -1 ? Number(row[iWage]) : 0;
        const remarks = iRemarks !== -1 ? row[iRemarks] : "";
        if (totalWage === 0) continue; 
        if (isTeikiHijokin(remarks, isObo)) continue;

        const workDate = new Date(row[iDate]);
        if (!isNaN(workDate.getTime())) {
          if (!maxDate || workDate > maxDate) maxDate = workDate;
        }
      }
    });
  } catch(e) {
    if (e.message.includes("中止")) throw e;
    Logger.log("最終日取得エラー: " + e.message);
  }
  
  return maxDate ? Utilities.formatDate(maxDate, "JST", "yyyy-MM-dd") : null;
}

// ==========================================
// 4. メイン集計処理
// ==========================================
function runAggregation(targetKyotenId, inputStartDate, inputEndDate) {
  if (!targetKyotenId || !inputStartDate || !inputEndDate) {
    throw new Error("拠点が正しく選択されていないか、日付が未入力です。");
  }
  
  CacheService.getUserCache().remove('cancel_flag'); 

  const [targetKyotenName, targetShinryoka] = targetKyotenId.split('_');
  const startDate = new Date(inputStartDate + " 00:00:00");
  const endDate = new Date(inputEndDate + " 23:59:59");

  const todayZero = new Date();
  todayZero.setHours(0, 0, 0, 0);

  const agencyIds = getAgencyIds();
  const holidaysList = getHolidays(); 
  
  const allData = fetchAllShiftData(startDate, endDate, targetKyotenName, targetShinryoka);
  const resultData = new Map();
  let actualMaxDate = new Date(1900, 0, 1); 

  let loopCount = 0;
  allData.forEach(row => {
    if (++loopCount % 100 === 0 && CacheService.getUserCache().get('cancel_flag') === 'true') {
      throw new Error("処理が中止されました。");
    }

    const docId = String(row.id).trim(); 
    const workDate = new Date(row.date);
    const totalWage = Number(row.wage);
    
    if (totalWage === 0) return; 
    if (isTeikiHijokin(row.remarks, row.isObo)) return; 

    if (!resultData.has(docId)) {
      resultData.set(docId, {
        name: row.name,
        attribute: agencyIds.includes(docId) ? "紹介会社" : "スポット",
        totalCount: 0,
        holidayCount: 0,
        lastWorkDate: null, 
        totalAm: 0, 
        totalPm1: 0,
        totalPm2: 0, 
        frames: {
          1: { am:0, pm1:0, pm2:0 }, 2: { am:0, pm1:0, pm2:0 }, 3: { am:0, pm1:0, pm2:0 },
          4: { am:0, pm1:0, pm2:0 }, 5: { am:0, pm1:0, pm2:0 }, 6: { am:0, pm1:0, pm2:0 },
          0: { am:0, pm1:0, pm2:0 }
        }
      });
    }

    const docData = resultData.get(docId);
    
    if (workDate < todayZero) {
      if (!docData.lastWorkDate || workDate > docData.lastWorkDate) {
        docData.lastWorkDate = workDate;
      }
    }

    if (workDate > actualMaxDate) actualMaxDate = workDate; 

    const dateString = Utilities.formatDate(workDate, "JST", "yyyy/MM/dd");
    if (holidaysList.includes(dateString)) docData.holidayCount += 1;

    const isKitakasai = (targetKyotenName === "北葛西");
    const frameCounts = calcShiftFrames(row.start, row.end, isKitakasai);
    const dayOfWeek = workDate.getDay();

    docData.frames[dayOfWeek].am += frameCounts.am;
    docData.frames[dayOfWeek].pm1 += frameCounts.pm1;
    docData.frames[dayOfWeek].pm2 += frameCounts.pm2;
    
    docData.totalAm += frameCounts.am;
    docData.totalPm1 += frameCounts.pm1;
    docData.totalPm2 += frameCounts.pm2;
    
    const shiftSum = frameCounts.am + frameCounts.pm1 + frameCounts.pm2;
    if (shiftSum > 0) docData.totalCount += shiftSum;
  });

  const filteredArr = Array.from(resultData.values()).filter(doc => doc.totalCount >= 3);

  if (filteredArr.length === 0) return "対象期間内に該当データ（勤務3回以上）がありません。";

  writeToKinmukunSheet(targetKyotenId, filteredArr, inputStartDate, actualMaxDate);
  return `集計完了！ 対象: ${filteredArr.length}名`;
}

// ==========================================
// 5. データ収集・ヘルパー関数群
// ==========================================
function getHolidays() {
  const holidays = [];
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.CALENDAR);
    const sheet = ss.getSheetByName("祝日登録");
    if (!sheet) return holidays;

    const data = sheet.getDataRange().getValues();
    for (let i = 2; i < data.length; i++) {
      const dateVal = data[i][0];
      if (dateVal && dateVal instanceof Date) {
        holidays.push(Utilities.formatDate(dateVal, "JST", "yyyy/MM/dd"));
      } else if (typeof dateVal === 'string' && dateVal.trim() !== '') {
        try {
          const d = new Date(dateVal);
          if (!isNaN(d.getTime())) holidays.push(Utilities.formatDate(d, "JST", "yyyy/MM/dd"));
        } catch(e) {}
      }
    }
  } catch(e) {}
  return holidays;
}

function isTeikiHijokin(remarks, isObo) {
  if (isObo) return false;
  if (!remarks) return false;
  return /\d{4}\/\d{2}\/\d{2}/.test(remarks) && remarks.includes("【") && remarks.includes("曜日");
}

function calcShiftFrames(startStr, endStr, isKitakasai) {
  const frames = { am: 0, pm1: 0, pm2: 0 };
  if (!startStr || !endStr) return frames;
  const parseTime = (t) => {
    try { const d = new Date(t); return isNaN(d.getTime()) ? null : d.getHours() + d.getMinutes() / 60; } catch(e) { return null; }
  };
  const start = parseTime(startStr), end = parseTime(endStr);
  if (start === null || end === null) return frames;
  if (start < 13 && end > 9) frames.am = 1;
  if (start < 18 && end > 15) frames.pm1 = 1;
  if (start < (isKitakasai ? 20 : 21) && end > 18) frames.pm2 = 1;
  return frames;
}

function getAgencyIds() {
  try {
    const sheet = SpreadsheetApp.openById(SPREADSHEET_IDS.AGENCY).getSheetByName("紹介会社応募表") || SpreadsheetApp.openById(SPREADSHEET_IDS.AGENCY).getSheets()[0];
    const data = sheet.getDataRange().getValues();
    if (data.length < 2) return [];

    const headers = data[0];
    let idIndex = headers.indexOf("医籍番号");
    
    if (idIndex === -1) idIndex = 1; 

    const ids = [];
    for (let i = 1; i < data.length; i++) {
      if (data[i][idIndex]) {
        ids.push(String(data[i][idIndex]).trim());
      }
    }
    return ids;
  } catch(e) { return []; }
}

function fetchAllShiftData(startDate, endDate, targetName, targetShinryoka) {
  const aggregatedData = [];
  const startMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const endMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  
  const processSheet = (ssId, isNewShiftBook) => {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const sName = sheet.getName();
        if (!isNewShiftBook) {
          if (!/^\d{4}\/\d{2}$/.test(sName)) return; 
          const sheetDate = new Date(sName + "/01");
          if (sheetDate < startMonth || sheetDate > endMonth) return; 
        } else {
          if (sName !== "確定シフト" && sName !== "応募シフト") return;
        }
        
        const isObo = (sName === "応募シフト");
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;
        
        const headers = data[0];
        const iName = headers.indexOf("名前"), iId = headers.indexOf("医籍番号"), iKyoten = headers.indexOf("クリニック名");
        const iShinryo = headers.indexOf("診療科"), iDate = headers.indexOf("勤務日");
        let iStart = headers.indexOf("(元のシフト）勤務開始時間"), iEnd = headers.indexOf("(元のシフト)勤務終了時間");
        if (iStart === -1) iStart = headers.indexOf("勤務開始時間");
        if (iEnd === -1) iEnd = headers.indexOf("勤務終了時間");
        const iWage = headers.indexOf("時給合計"), iRemarks = headers.indexOf("業務内容（備考）");

        if (iName===-1 || iId===-1 || iKyoten===-1 || iDate===-1) return;

        for (let i = 1; i < data.length; i++) {
          if (i % 100 === 0 && CacheService.getUserCache().get('cancel_flag') === 'true') {
            throw new Error("処理が中止されました。");
          }

          const row = data[i];
          const rDate = new Date(row[iDate]);
          
          if (row[iKyoten] !== targetName || (targetShinryoka && row[iShinryo] !== targetShinryoka)) continue;
          if (rDate < startDate || rDate > endDate) continue;

          aggregatedData.push({
            name: row[iName], id: String(row[iId]).trim(), date: row[iDate], start: row[iStart], end: row[iEnd],
            wage: iWage !== -1 ? row[iWage] : 0, remarks: iRemarks !== -1 ? row[iRemarks] : "", isObo: isObo
          });
        }
      });
    } catch(e) { 
      if(e.message.includes("中止")) throw e;
    }
  };

  processSheet(SPREADSHEET_IDS.MONTHLY_1, false);
  processSheet(SPREADSHEET_IDS.MONTHLY_2, false);
  processSheet(SPREADSHEET_IDS.SHIFT_NEW, true);
  return aggregatedData;
}

// ==========================================
// 6. シート書き出し
// ==========================================
function writeToKinmukunSheet(targetKyotenId, filteredArr, startDateStr, maxDate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("勤務くん");
  
  const top30 = filteredArr.sort((a, b) => b.totalCount - a.totalCount).slice(0, 30);
  const displayTitle = targetKyotenId.replace('_', '（') + (targetKyotenId.includes('_') ? '）' : '');
  sheet.getRange("A1").setValue(displayTitle);
  
  const formatDateStr = (d) => Utilities.formatDate(d, "JST", "yyyy/MM/dd");
  
  sheet.getRange("I1").setValue(startDateStr.replace(/-/g, '/')); 
  sheet.getRange("M1").setValue(formatDateStr(maxDate));

  const startRow = 5;
  const numRows = sheet.getMaxRows() - startRow + 1;
  if(numRows > 0) sheet.getRange(startRow, 1, numRows, 29).clearContent();
  if (top30.length === 0) return;

  const outputValues = [];
  const outputBackgrounds = [];

  top30.forEach(doc => {
    const row = new Array(29).fill("");
    const bgRow = new Array(29).fill("#ffffff"); 

    row[0] = doc.name; 
    row[1] = doc.attribute; 
    row[2] = doc.totalCount; 
    row[3] = doc.lastWorkDate ? formatDateStr(doc.lastWorkDate) : "-"; 
    row[4] = doc.holidayCount; 
    
    if (doc.totalAm) {
      row[5] = doc.totalAm;
    } else {
      bgRow[5] = "#f3f3f3";
    }
    
    if (doc.totalPm1) {
      row[6] = doc.totalPm1;
    } else {
      bgRow[6] = "#f3f3f3";
    }
    
    if (doc.totalPm2) {
      row[7] = doc.totalPm2;
    } else {
      bgRow[7] = "#f3f3f3";
    }
    
    let colIdx = 8; 
    [1, 2, 3, 4, 5, 6, 0].forEach(day => {
      // 9-13
      if(doc.frames[day].am) row[colIdx] = doc.frames[day].am;
      else bgRow[colIdx] = "#f3f3f3"; 
      colIdx++;
      
      // 15-18
      if(doc.frames[day].pm1) row[colIdx] = doc.frames[day].pm1;
      else bgRow[colIdx] = "#f3f3f3";
      colIdx++;
      
      // 18-21
      if(doc.frames[day].pm2) row[colIdx] = doc.frames[day].pm2;
      else bgRow[colIdx] = "#f3f3f3";
      colIdx++;
    });
    
    outputValues.push(row);
    outputBackgrounds.push(bgRow);
  });

  const targetRange = sheet.getRange(startRow, 1, outputValues.length, 29);
  targetRange.setValues(outputValues);
  targetRange.setBackgrounds(outputBackgrounds);
  targetRange.setHorizontalAlignment("left"); 
}


// ==========================================
// 7. 医師検索機能 バックエンド処理
// ==========================================

function showSearchDialog() {
  const html = HtmlService.createTemplateFromFile('SearchIndex').evaluate()
      .setTitle('医師検索パネル')
      .setWidth(900)  
      .setHeight(600);
  SpreadsheetApp.getUi().showModalDialog(html, '医師検索・分析パネル');
}

function getKyotenTree() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_IDS.KYOTEN);
    const data = ss.getSheetByName('拠点名').getDataRange().getValues();
    if (data.length < 2) return {};

    const headers = data[0];
    const iName = headers.indexOf("正規記載");
    const iGroup = headers.indexOf("拠点グループ");
    const iArea = headers.indexOf("エリア");
    const iOpenDate = headers.indexOf("開院日");

    if (iName === -1 || iGroup === -1 || iArea === -1 || iOpenDate === -1) {
      throw new Error("必要なヘッダー（正規記載, 拠点グループ, エリア, 開院日）が見つかりません。");
    }

    const tree = {};
    const today = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[iName];
      const group = row[iGroup];
      const area = row[iArea];
      const openDateStr = row[iOpenDate];
      
      if (!openDateStr || !area || !group) continue;
      
      const openDate = new Date(openDateStr);
      if (group !== "MQC" && !isNaN(openDate.getTime()) && openDate <= today) {
        if (!tree[area]) tree[area] = {};
        if (!tree[area][group]) tree[area][group] = [];
        
        if (name === "北葛西" || name === "亀有") {
          if (!tree[area][group].includes(`${name}（内科）`)) {
            tree[area][group].push(`${name}（内科）`, `${name}（小児科）`);
          }
        } else {
          if (!tree[area][group].includes(name)) {
            tree[area][group].push(name);
          }
        }
      }
    }
    return tree;
  } catch (e) {
    throw new Error("拠点マスタの読み込みに失敗しました: " + e.message);
  }
}

function getUniqueDoctorNames() {
  const names = new Set();
  
  try {
    const agencySs = SpreadsheetApp.openById(SPREADSHEET_IDS.AGENCY);
    const agencySheet = agencySs.getSheets()[0];
    const agencyData = agencySheet.getDataRange().getValues();
    if (agencyData.length > 1) {
      const headers = agencyData[0];
      let nameIdx = headers.indexOf("名前");
      if (nameIdx === -1) nameIdx = 0; 
      for (let i = 1; i < agencyData.length; i++) {
        if (agencyData[i][nameIdx]) names.add(String(agencyData[i][nameIdx]).trim());
      }
    }
  } catch(e) {}

  try {
    const shiftSs = SpreadsheetApp.openById(SPREADSHEET_IDS.SHIFT_NEW);
    ["確定シフト", "応募シフト"].forEach(sName => {
      const sheet = shiftSs.getSheetByName(sName);
      if (!sheet) return;
      const data = sheet.getDataRange().getValues();
      if (data.length > 1) {
        const headers = data[0];
        const nameIdx = headers.indexOf("名前");
        if (nameIdx !== -1) {
          for (let i = 1; i < data.length; i++) {
            if (data[i][nameIdx]) names.add(String(data[i][nameIdx]).trim());
          }
        }
      }
    });
  } catch(e) {}

  return Array.from(names).sort();
}

function searchDoctorData(params) {
  const { doctorName, selectedKyotens, includeTeiki } = params;
  if (!doctorName) throw new Error("医師名を入力してください。");
  if (!selectedKyotens || selectedKyotens.length === 0) throw new Error("拠点を1つ以上選択してください。");

  const today = new Date();
  const startDate = new Date(today.getFullYear() - 1, today.getMonth(), 1);
  const endDate = new Date(today.getFullYear() + 1, today.getMonth(), 1);

  const holidaysList = getHolidays();
  const allData = fetchAllShiftDataForSearch(startDate, endDate, selectedKyotens);
  
  const resultData = new Map();

  allData.forEach(row => {
    if (!row.name.includes(doctorName)) return;

    const totalWage = Number(row.wage);
    if (totalWage === 0) return; 
    
    if (!includeTeiki) {
      if (isTeikiHijokin(row.remarks, row.isObo)) return; 
    }

    const key = `${row.name}_${row.kyoten}_${row.shinryoka}`;
    const displayKyoten = row.shinryoka && (row.kyoten === "北葛西" || row.kyoten === "亀有") 
                          ? `${row.kyoten}（${row.shinryoka}）` : row.kyoten;

    if (!resultData.has(key)) {
      resultData.set(key, {
        name: row.name,
        kyoten: displayKyoten,
        totalCount: 0,
        holidayCount: 0,
        frames: { am: 0, pm1: 0, pm2: 0 },
        days: { 0:0, 1:0, 2:0, 3:0, 4:0, 5:0, 6:0 } 
      });
    }

    const docData = resultData.get(key);
    const workDate = new Date(row.date);
    
    const dateString = Utilities.formatDate(workDate, "JST", "yyyy/MM/dd");
    if (holidaysList.includes(dateString)) docData.holidayCount += 1;

    const isKitakasai = (row.kyoten === "北葛西");
    const frameCounts = calcShiftFrames(row.start, row.end, isKitakasai);
    const dayOfWeek = workDate.getDay();

    docData.frames.am += frameCounts.am;
    docData.frames.pm1 += frameCounts.pm1;
    docData.frames.pm2 += frameCounts.pm2;
    
    const shiftSum = frameCounts.am + frameCounts.pm1 + frameCounts.pm2;
    if (shiftSum > 0) {
      docData.totalCount += shiftSum;
      docData.days[dayOfWeek] += shiftSum;
    }
  });

  return Array.from(resultData.values()).sort((a, b) => b.totalCount - a.totalCount);
}

function fetchAllShiftDataForSearch(startDate, endDate, selectedKyotens) {
  const aggregatedData = [];
  const startMonth = new Date(startDate.getFullYear(), startDate.getMonth(), 1);
  const endMonth = new Date(endDate.getFullYear(), endDate.getMonth(), 1);
  
  const processSheet = (ssId, isNewShiftBook) => {
    try {
      const ss = SpreadsheetApp.openById(ssId);
      const sheets = ss.getSheets();
      
      sheets.forEach(sheet => {
        const sName = sheet.getName();
        if (!isNewShiftBook) {
          if (!/^\d{4}\/\d{2}$/.test(sName)) return; 
          const sheetDate = new Date(sName + "/01");
          if (sheetDate < startMonth || sheetDate > endMonth) return; 
        } else {
          if (sName !== "確定シフト" && sName !== "応募シフト") return;
        }
        
        const isObo = (sName === "応募シフト");
        const data = sheet.getDataRange().getValues();
        if (data.length < 2) return;
        
        const headers = data[0];
        const iName = headers.indexOf("名前"), iKyoten = headers.indexOf("クリニック名");
        const iShinryo = headers.indexOf("診療科"), iDate = headers.indexOf("勤務日");
        let iStart = headers.indexOf("(元のシフト）勤務開始時間"), iEnd = headers.indexOf("(元のシフト)勤務終了時間");
        if (iStart === -1) iStart = headers.indexOf("勤務開始時間");
        if (iEnd === -1) iEnd = headers.indexOf("勤務終了時間");
        const iWage = headers.indexOf("時給合計"), iRemarks = headers.indexOf("業務内容（備考）");

        if (iName===-1 || iKyoten===-1 || iDate===-1) return;

        for (let i = 1; i < data.length; i++) {
          const row = data[i];
          const rDate = new Date(row[iDate]);
          
          if (rDate < startDate || rDate > endDate) continue;
          
          const kyotenName = row[iKyoten];
          const shinryo = row[iShinryo];
          const displayKyoten = shinryo && (kyotenName === "北葛西" || kyotenName === "亀有") 
                                ? `${kyotenName}（${shinryo}）` : kyotenName;
          
          if (!selectedKyotens.includes(displayKyoten)) continue;

          aggregatedData.push({
            name: row[iName], kyoten: kyotenName, shinryoka: shinryo,
            date: row[iDate], start: row[iStart], end: row[iEnd],
            wage: iWage !== -1 ? row[iWage] : 0, remarks: iRemarks !== -1 ? row[iRemarks] : "", isObo: isObo
          });
        }
      });
    } catch(e) {}
  };

  processSheet(SPREADSHEET_IDS.MONTHLY_1, false);
  processSheet(SPREADSHEET_IDS.MONTHLY_2, false);
  processSheet(SPREADSHEET_IDS.SHIFT_NEW, true);
  return aggregatedData;
}