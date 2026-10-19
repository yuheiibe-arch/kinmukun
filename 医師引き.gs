// ==========================================
// 医師検索機能 バックエンド処理
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
    const tree = {};
    const today = new Date();
    
    for (let i = 1; i < data.length; i++) {
      const row = data[i];
      const name = row[0];
      const group = row[5];
      const area = row[6];
      const openDateStr = row[7];
      if (!openDateStr || !area || !group) continue;
      
      const openDate = new Date(openDateStr);
      if (group !== "MQC" && openDate <= today) {
        if (!tree[area]) tree[area] = {};
        if (!tree[area][group]) tree[area][group] = [];
        
        if (name === "北葛西" || name === "亀有") {
          tree[area][group].push(`${name}（内科）`, `${name}（小児科）`);
        } else {
          tree[area][group].push(name);
        }
      }
    }
    return tree;
  } catch (e) {
    throw new Error("拠点マスタの読み込みに失敗しました。");
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

// ▼ 検索実行関数 (includeTeiki の判定を追加)
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
    
    // ▼ スイッチがOFF(false)の場合のみ、定期非常勤を除外する
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