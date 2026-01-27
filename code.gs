// --- Configuration ---
var CONFIG = {
  SOURCE_SHEET_NAME: "Shotship", // ชื่อ Sheet ข้อมูลหลัก
  // ชื่อคอลัมน์ใน Sheet Shotship (ต้องตรงกับ Header เป๊ะๆ)
  COL_DATE: "Date",
  COL_PLANT: "plant",
  COL_DOC_NO: "เลขที่ใบเบิก",
  COL_MAT_CODE: "Material Code",
  COL_MAT_NAME: "Material Name",
  COL_TYPE: "ประเภทอะไหล่",
  COL_REQ: "จำนวนที่ขอเบิก",
  COL_APPR: "จำนวนอนุมัติ",
  COL_DIFF: "ผลต่าง",
  COL_NOTE: "Note" // ถ้าใน Shotship ไม่มีคอลัมน์นี้ อาจต้องดึงจาก ShotShipRemarks แต่เบื้องต้นสมมติว่าเอาจากที่นี่
};

function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('ShotShip Reports')
      .addItem('Generate All Reports', 'generateReports')
      .addToUi();
}

function generateReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("ไม่พบ Sheet ชื่อ: " + CONFIG.SOURCE_SHEET_NAME);
    return;
  }
  
  var data = sourceSheet.getDataRange().getValues();
  if (data.length < 2) return;
  
  var headers = data[0];
  var colMap = {};
  headers.forEach(function(h, i) { colMap[h] = i; });
  
  // Helper to get value securely
  var getVal = function(row, colName) {
    var idx = colMap[CONFIG[colName]];
    return idx !== undefined ? row[idx] : "";
  };
  
  // Data Containers
  var diffItems = [];
  var dailyMap = {};
  var monthlyMap = {};
  var quarterlyMap = {};
  
  var today = new Date();
  today.setHours(0,0,0,0);
  var fifteenDaysAgo = new Date(today);
  fifteenDaysAgo.setDate(today.getDate() - 15);
  
  // Process Data (Skip Header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    var dateStr = getVal(row, 'COL_DATE'); // e.g., "1/10/2023 10:00:00"
    if (!dateStr) continue;
    
    var dateObj = parseDate(dateStr);
    if (!dateObj) continue; // Skip invalid dates
    
    var plant = String(getVal(row, 'COL_PLANT') || "Unknown").trim();
    if (plant.length === 3) plant = "0" + plant;
    
    var req = parseFloat(getVal(row, 'COL_REQ') || 0);
    var diff = parseFloat(getVal(row, 'COL_DIFF') || 0);
    var note = getVal(row, 'COL_NOTE');
    
    // Logic: Real Approved = Requested - Diff
    // (User said: "หักจาก จำนวนจาก Sheet ผลต่างเพื่อให้ข้อมูลสรุปถูกต้อง")
    // If Diff > 0, Real Appr = Req - Diff. (This handles partial delivery failure)
    var realAppr = req - diff;
    if (realAppr < 0) realAppr = 0;
    
    // --- 1. Collect Diff Items ---
    if (diff > 0 || (note && String(note).trim() !== "")) {
      // Add row to diff items
      diffItems.push([
        formatDate(dateObj), plant, 
        getVal(row, 'COL_DOC_NO'), 
        getVal(row, 'COL_MAT_CODE'), 
        getVal(row, 'COL_MAT_NAME'),
        getVal(row, 'COL_TYPE'),
        req, realAppr, diff, note
      ]);
    }
    
    // --- Aggregation Keys ---
    // Daily (Date + Plant)
    var dayKey = formatDate(dateObj); // "dd/mm/yyyy"
    
    // Monthly (Month/Year + Plant)
    var monthKey = (dateObj.getMonth() + 1) + "/" + dateObj.getFullYear();
    
    // Quarterly (Qx/Year + Plant)
    var q = Math.ceil((dateObj.getMonth() + 1) / 3);
    var quarterKey = "Q" + q + "/" + dateObj.getFullYear();
    
    // --- 2. Aggregate Daily (Only if within 15 days) ---
    if (dateObj >= fifteenDaysAgo) {
      if (!dailyMap[dayKey]) dailyMap[dayKey] = {};
      if (!dailyMap[dayKey][plant]) dailyMap[dayKey][plant] = initStats();
      updateStats(dailyMap[dayKey][plant], req, realAppr);
    }
    
    // --- 3. Aggregate Monthly ---
    if (!monthlyMap[monthKey]) monthlyMap[monthKey] = {};
    if (!monthlyMap[monthKey][plant]) monthlyMap[monthKey][plant] = initStats();
    updateStats(monthlyMap[monthKey][plant], req, realAppr);
    
    // --- 4. Aggregate Quarterly ---
    if (!quarterlyMap[quarterKey]) quarterlyMap[quarterKey] = {};
    if (!quarterlyMap[quarterKey][plant]) quarterlyMap[quarterKey][plant] = initStats();
    updateStats(quarterlyMap[quarterKey][plant], req, realAppr);
  }
  
  // Write Sheets
  writeSheet(ss, "Report_Diff_Items", ["Date", "Plant", "DocNo", "Material", "Name", "Type", "Req", "RealAppr", "Diff", "Note"], diffItems);
  writeAggregatedSheet(ss, "Report_Daily_15Days", "Date", dailyMap);
  writeAggregatedSheet(ss, "Report_Monthly", "Month", monthlyMap);
  writeAggregatedSheet(ss, "Report_Quarterly", "Quarter", quarterlyMap);
  
  SpreadsheetApp.getUi().alert("สร้างรายงานเสร็จสิ้น!");
}

function initStats() {
  return { reqItems: 0, apprItems: 0, totalReq: 0, totalAppr: 0, docSet: {} };
}

function updateStats(stats, req, realAppr, docNo) {
  stats.totalReq += req;
  stats.totalAppr += realAppr;
  if (req > 0) stats.reqItems++;
  if (realAppr > 0) stats.apprItems++;
  if (docNo) stats.docSet[docNo] = true;
}

// Helper: Parse "d/m/yyyy H:M:S" or similar
function parseDate(str) {
  if (str instanceof Date) return str;
  var parts = String(str).split(' ')[0].split('/');
  if (parts.length < 3) return null;
  // Note: new Date(y, m-1, d) creates date in script's timezone (usually UTC or America/Pacific)
  // Better to handle string directly if possible, or correct the zone.
  return new Date(parts[2], parts[1]-1, parts[0]);
}

function formatDate(date) {
  // Use Utilities.formatDate to strictly use Thai Timezone (GMT+7)
  // This prevents 19:47 becoming the next day due to timezone shifts
  return Utilities.formatDate(date, "GMT+7", "dd/MM/yyyy");
}

// Output Buffer
// Output Buffer
function writeSheet(ss, sheetName, headers, rowData) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }
  
  var allData = [headers];
  if (rowData.length > 0) {
    allData = allData.concat(rowData);
  }
  
  if (allData.length > 0) {
    sheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);
  }
  SpreadsheetApp.flush();
}

function writeAggregatedSheet(ss, sheetName, keyLabel, mapData) {
  var sheet = ss.getSheetByName(sheetName);
  if (!sheet) {
    sheet = ss.insertSheet(sheetName);
  } else {
    sheet.clearContents();
  }
  
  // Update Headers to include "Part Type"
  var headers = [keyLabel, "Plant", "Part Type", "Doc Count", "Req Items", "Appr Items", "Req Qty", "Real Appr Qty", "% Efficiency"];
  var rows = [];
  
  var sortedKeys = Object.keys(mapData).sort(function(a,b) {
     var pA = a.replace("Q","").split('/'); 
     var pB = b.replace("Q","").split('/');
     var yA = parseInt(pA[pA.length-1]), yB = parseInt(pB[pB.length-1]);
     var mA = parseInt(pA[0]), mB = parseInt(pB[0]);
     if (pA.length > 2) { 
        mA = parseInt(pA[1]); var dA = parseInt(pA[0]);
        mB = parseInt(pB[1]); var dB = parseInt(pB[0]);
        return (yA - yB) || (mA - mB) || (dA - dB);
     }
     return (yA - yB) || (mA - mB);
  });

  for (var k = 0; k < sortedKeys.length; k++) {
    var mainKey = sortedKeys[k];
    var plants = mapData[mainKey];
    for (var p in plants) {
      var types = plants[p];
      for (var t in types) {
        var s = types[t];
        var eff = s.totalReq > 0 ? (s.totalAppr / s.totalReq) * 100 : 0;
        // Calculate unique docs size
        var docCount = 0;
        if (s.docSet) {
           // Apps Script limitation: Set is not fully supported in older V8?? No, V8 is fine.
           // However, to be safe and simple:
           docCount = Object.keys(s.docSet).length;
        }
        
        rows.push([
          mainKey, p, t, 
          docCount, // Doc Count
          s.reqItems, s.apprItems, 
          s.totalReq, s.totalAppr, 
          eff.toFixed(2) + "%"
        ]);
      }
    }
  }
  
  var allData = [headers].concat(rows);
  
  if (allData.length > 0) {
    sheet.getRange(1, 1, allData.length, allData[0].length).setValues(allData);
  }
  SpreadsheetApp.flush();
}

function generateReports() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var sourceSheet = ss.getSheetByName(CONFIG.SOURCE_SHEET_NAME);
  
  if (!sourceSheet) {
    SpreadsheetApp.getUi().alert("ไม่พบ Sheet ชื่อ: " + CONFIG.SOURCE_SHEET_NAME);
    return;
  }
  
  var data = sourceSheet.getDataRange().getValues();
  // Fetch Display Values to get exact date strings as shown in sheet (avoids timezone issues)
  var displayData = sourceSheet.getDataRange().getDisplayValues();
  
  if (data.length < 2) return;
  
  var headers = data[0];
  var colMap = {};
  headers.forEach(function(h, i) { colMap[h] = i; });
  
  // Helper to get value
  var getVal = function(row, colName) {
    var idx = colMap[CONFIG[colName]];
    return idx !== undefined ? row[idx] : "";
  };
  
  // Helper to get display value (for Dates)
  var getDisplayVal = function(rowIndex, colName) {
    var idx = colMap[CONFIG[colName]];
    return idx !== undefined ? displayData[rowIndex][idx] : "";
  };
  
  // Data Containers
  var diffItems = [];
  var dailyMap = {};
  var monthlyMap = {};
  var quarterlyMap = {};
  
  var today = new Date();
  today.setHours(0,0,0,0);
  var thirtyDaysAgo = new Date(today);
  thirtyDaysAgo.setDate(today.getDate() - 30);
  
  // Process Data (Skip Header)
  for (var i = 1; i < data.length; i++) {
    var row = data[i];
    // Use Display Value for Date to get literal "dd/mm/yyyy"
    var dateStr = getDisplayVal(i, 'COL_DATE'); 
    
    // Fallback: If display val is empty/weird but raw val is Date Obj, use raw?
    if (!dateStr) continue;
    
    // Ensure dateStr is treated as String
    dateStr = String(dateStr);
    
    var dateObj = parseDate(dateStr);
    if (!dateObj) continue; // Skip invalid dates 
    
    var plant = String(getVal(row, 'COL_PLANT') || "Unknown").trim();
    if (plant.length === 3) plant = "0" + plant;

    var type = String(getVal(row, 'COL_TYPE') || "Other").trim(); 
    var docNo = String(getVal(row, 'COL_DOC_NO') || "").trim();
    
    var req = parseFloat(getVal(row, 'COL_REQ') || 0);
    var diff = parseFloat(getVal(row, 'COL_DIFF') || 0);
    var note = getVal(row, 'COL_NOTE');
    
    // Logic: Real Approved = Requested - Diff (Only for aggregated stats)
    var realAppr = req - diff;
    if (realAppr < 0) realAppr = 0;
    
    // --- 1. Diff Items (Filter > 0 or Note) ---
    // User requested: "All columns of Sheet: Shotship where Diff column > 0"
    if (diff > 0 || (String(note).trim() !== "")) {
      // Push the ENTIRE source row
      diffItems.push(row);
    }
    
    // Generate Keys using Timezone-Safe String (dayKey is already GMT+7 "dd/MM/yyyy")
    var dayKey = formatDate(dateObj); 
    
    // Parse the SAFE dayKey back to ensure consistent Month/Year
    // dayKey is "26/01/2026", so parts: [26, 01, 2026]
    var parts = dayKey.split('/');
    var d = parseInt(parts[0]);
    var m = parseInt(parts[1]); 
    var y = parseInt(parts[2]);
    
    var monthKey = m + "/" + y;
    
    var q = Math.ceil(m / 3);
    var quarterKey = "Q" + q + "/" + y;
    
    var ensureStats = function(map, key, plant, type) {
        if (!map[key]) map[key] = {};
        if (!map[key][plant]) map[key][plant] = {};
        if (!map[key][plant][type]) map[key][plant][type] = initStats();
        return map[key][plant][type];
    };

    if (dateObj >= thirtyDaysAgo) {
      updateStats(ensureStats(dailyMap, dayKey, plant, type), req, realAppr, docNo);
    }
    
    updateStats(ensureStats(monthlyMap, monthKey, plant, type), req, realAppr, docNo);
    updateStats(ensureStats(quarterlyMap, quarterKey, plant, type), req, realAppr, docNo);
  }
  
  // Write Sheets
  // Report_Diff_Items uses raw headers from source
  // writeSheet(ss, "Report_Diff_Items", headers, diffItems);
  
  writeAggregatedSheet(ss, "Report_Daily_30Days", "Date", dailyMap); 
  writeAggregatedSheet(ss, "Report_Monthly", "Month", monthlyMap);
  writeAggregatedSheet(ss, "Report_Quarterly", "Quarter", quarterlyMap);
  
  SpreadsheetApp.getUi().alert("สร้างรายงานเสร็จสิ้น!");
}


function doPost(e) {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  var sheet = ss.getSheetByName("ShotShipRemarks");
  
  // ถ้ายังไม่มี Sheet ให้สร้างใหม่และใส่ Header
  if (!sheet) {
    sheet = ss.insertSheet("ShotShipRemarks");
    sheet.appendRow(["Key", "DocNo", "MaterialCode", "MaterialName", "PartType", "Note", "User", "Timestamp"]);
  }
  
  var data = JSON.parse(e.postData.contents);
  var action = data.action;
  var key = data.docNo + "-" + data.matCode; // Key: เลขที่ใบเบิก + Material Code
  
  var rows = sheet.getDataRange().getValues();
  var rowIndex = -1;
  
  // ค้นหา Key ที่มีอยู่แล้ว (เริ่มค้นหาจากแถวที่ 2 เพื่อข้าม Header)
  for (var i = 1; i < rows.length; i++) {
    if (rows[i][0] == key) {
      rowIndex = i + 1; // แปลงเป็น index ของ Sheet (เริ่มที่ 1)
      break;
    }
  }
  
  if (action == "save") {
    var timestamp = new Date();
    if (rowIndex > 0) {
      // ถ้ามีข้อมูลอยู่แล้ว ให้ Update หมายเหตุ, ผู้บันทึก และเวลา
      sheet.getRange(rowIndex, 6).setValue(data.note);
      sheet.getRange(rowIndex, 7).setValue(data.user);
      sheet.getRange(rowIndex, 8).setValue(timestamp);
    } else {
      // ถ้าไม่มี ให้เพิ่มแถวใหม่
      sheet.appendRow([key, data.docNo, data.matCode, data.matName, data.partType, data.note, data.user, timestamp]);
    }
    return ContentService.createTextOutput("Success").setMimeType(ContentService.MimeType.TEXT);
  } 
  else if (action == "delete") {
    if (rowIndex > 0) {
      // ถ้าเจอ Key ให้ลบแถวนั้น
      sheet.deleteRow(rowIndex);
      return ContentService.createTextOutput("Deleted").setMimeType(ContentService.MimeType.TEXT);
    } else {
      return ContentService.createTextOutput("Not Found").setMimeType(ContentService.MimeType.TEXT);
    }
  }
  
  return ContentService.createTextOutput("Invalid Action").setMimeType(ContentService.MimeType.TEXT);
}
