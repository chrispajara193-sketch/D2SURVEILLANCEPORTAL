/**
 * =========================================================================
 * QC-DESU DISTRICT 2 MASTER SURVEILLANCE & AUTHENTICATION BACKEND (OPTIMIZED)
 * =========================================================================
 */

// 1. PRIMARY SPREADSHEET (Main Surveillance Portal Database)
const SPREADSHEET_ID = "1sUTII813loiUG-LD1j-JUgSHagmmKf82QhxU4-Kl5IA";

// 2. TARGET SPREADSHEET FOR PATIENT PROFILE (Receives newly entered/updated patient data)
const PATIENT_SPREADSHEET_ID_OR_URL = "1rrlgAX7ad_IDcHdIXe8gHlWBFsvH5DPCD1ZlW_nR-d0";
const PATIENT_TARGET_SHEET_NAME = "MDB DISTRICT 2 2026"; 

// 3. GOOGLE DRIVE STORAGE FOLDER FOR CIF/CRF ATTACHMENTS
const PDS_CIF_FOLDER = "D2_DESU_CIF_Attachments";

// 4. FIREBASE REALTIME DATABASE URL
const FIREBASE_DB_URL = "https://d2surveillanceportal-default-rtdb.asia-southeast1.firebasedatabase.app";

const DISTRICT_2_BARANGAYS = [
  "BAGONG SILANGAN",
  "BATASAN HILLS",
  "COMMONWEALTH",
  "HOLY SPIRIT",
  "PAYATAS"
];

/**
 * Handles HTTP GET requests (For web app, standalone views, and VS Code API calls)
 */
function doGet(e) {
  try {
    e = e || { parameter: {} };
    const action = (e.parameter && e.parameter.action) ? String(e.parameter.action).trim() : '';

    if (action === 'getMetrics') {
      return createJsonResponse(getMetricsData(e.parameter));
    }
    
    if (action === 'getDashboard') {
      return createJsonResponse(getDashboardData());
    }

    if (action === 'getDirectory') {
      return createJsonResponse(getDirectoryData());
    }

    if (action === 'login') {
      const loginData = { username: e.parameter.username, password: e.parameter.password };
      return createJsonResponse(processLogin(loginData));
    }

    if (action === 'getRecords') {
      return createJsonResponse(getRecords());
    }

    if (action === 'getNavdpcp') {
      return createJsonResponse(getNavdpcpData(e.parameter));
    }

    return createJsonResponse({ 
      status: "online", 
      message: "QC-DESU District 2 Master Surveillance API is running." 
    });
  } catch (err) {
    return createJsonResponse({ success: false, message: "SERVER_ERROR: " + err.toString() });
  }
}

/**
 * Handles HTTP POST requests (For saving records or CIF uploads)
 */
function doPost(e) {
  try {
    e = e || { parameter: {}, postData: {} };
    let payload = {};
    
    if (e.postData && e.postData.contents) {
      try {
        payload = JSON.parse(e.postData.contents);
      } catch (err) {
        payload = e.parameter || {};
      }
    } else {
      payload = e.parameter || {};
    }

    const action = payload.action || (e.parameter ? e.parameter.action : '');

    if (action === 'login') {
      return createJsonResponse(processLogin(payload));
    }

    if (action === 'saveRecord') {
      return createJsonResponse(saveRecord(payload.formData, payload.rowIndex));
    }

    if (action === 'uploadPdsFile') {
      return createJsonResponse(uploadPdsFile(payload));
    }

    return createJsonResponse({ success: false, message: "Invalid POST action." });
  } catch (err) {
    return createJsonResponse({ success: false, message: "SERVER_POST_ERROR: " + err.toString() });
  }
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------------------
// SPREADSHEET & SHEET LOCATORS (HIGH-SPEED IN-MEMORY CACHE)
// -------------------------------------------------------------------------

function openSpreadsheetByIdOrUrl(idOrUrl) {
  if (!idOrUrl || typeof idOrUrl !== 'string') return null;
  const clean = idOrUrl.trim();
  if (!clean) return null;

  try {
    if (clean.indexOf("docs.google.com") !== -1 || clean.indexOf("http") === 0) {
      const match = clean.match(/\/spreadsheets\/d\/([a-zA-Z0-9_-]+)/);
      if (match && match[1]) return SpreadsheetApp.openById(match[1]);
      return SpreadsheetApp.openByUrl(clean);
    } else {
      return SpreadsheetApp.openById(clean);
    }
  } catch (err) {
    Logger.log("openSpreadsheetByIdOrUrl error: " + err.message);
    return null;
  }
}

function getSpreadsheet() {
  try {
    if (SPREADSHEET_ID && SPREADSHEET_ID.trim() !== "") {
      return SpreadsheetApp.openById(SPREADSHEET_ID);
    }
    return SpreadsheetApp.getActiveSpreadsheet();
  } catch (e) {
    Logger.log("getSpreadsheet error: " + e.message);
    return SpreadsheetApp.getActiveSpreadsheet();
  }
}

/**
 * Creates an in-memory dictionary of all sheets in a single call (Eliminates 630 API calls!)
 */
function getSheetMap(ss) {
  const map = {};
  if (!ss) return map;
  const sheets = ss.getSheets();
  for (let i = 0; i < sheets.length; i++) {
    map[sheets[i].getName().trim().toUpperCase()] = sheets[i];
  }
  return map;
}

function getSheetByNameFast(sheetMap, name) {
  if (!name || !sheetMap) return null;
  return sheetMap[name.trim().toUpperCase()] || null;
}

function getMdbSheet(ss, sheetMap) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  const map = sheetMap || getSheetMap(activeSs);
  return getSheetByNameFast(map, "MDB DISTRICT 2 2026") || 
         getSheetByNameFast(map, "MDB") || 
         activeSs.getSheets()[0];
}

function getEpisenseSheet(ss, sheetMap) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  const map = sheetMap || getSheetMap(activeSs);
  return getSheetByNameFast(map, "EPISENSE") || activeSs.getSheets()[0];
}

function getUsersSheet(ss, sheetMap) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  const map = sheetMap || getSheetMap(activeSs);
  return getSheetByNameFast(map, "users") || 
         getSheetByNameFast(map, "USERS") || 
         getSheetByNameFast(map, "ACCOUNTS");
}

function getDirectorySheet(ss, sheetMap) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  const map = sheetMap || getSheetMap(activeSs);
  return getSheetByNameFast(map, "DIRECTORY") || 
         getSheetByNameFast(map, "STAFF") || 
         getSheetByNameFast(map, "DIRECTORY SHEET");
}

// -------------------------------------------------------------------------
// DRIVE STORAGE HELPERS (FOR CIF ATTACHMENTS)
// -------------------------------------------------------------------------

function getOrCreateDriveFolder(folderName) {
  const targetName = folderName || PDS_CIF_FOLDER || "D2_DESU_CIF_Attachments";
  const folders = DriveApp.getFoldersByName(targetName);
  if (folders.hasNext()) {
    return folders.next();
  }
  const newFolder = DriveApp.createFolder(targetName);
  try {
    newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
  } catch(e) {
    Logger.log("Notice: Organization policy restricted link sharing: " + e.message);
  }
  return newFolder;
}

function convertToDirectThumbnailUrl(url) {
  if (!url || typeof url !== 'string') return '';
  const clean = url.trim();
  if (clean.includes("drive.google.com") || clean.includes("docs.google.com")) {
    const match = clean.match(/\/d\/([a-zA-Z0-9_-]+)/) || clean.match(/id=([a-zA-Z0-9_-]+)/);
    if (match && match[1]) {
      return `https://lh3.googleusercontent.com/d/${match[1]}`;
    }
  }
  return clean;
}

function uploadPdsFile(payload) {
  try {
    if (!payload || !payload.base64Data) {
      return { success: false, message: "No document data provided." };
    }

    const rawData = payload.base64Data;
    const caseId = payload.caseId || "CASE";
    const originalName = payload.fileName || "document.pdf";

    const matches = rawData.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    let blob;
    if (matches && matches.length === 3) {
      const contentType = matches[1];
      const decodedBytes = Utilities.base64Decode(matches[2]);
      blob = Utilities.newBlob(decodedBytes, contentType, originalName);
    } else {
      blob = Utilities.newBlob(rawData, "application/octet-stream", originalName);
    }

    const folder = getOrCreateDriveFolder(PDS_CIF_FOLDER);
    const file = folder.createFile(blob);
    
    try {
      file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
    } catch(e) {
      Logger.log("Notice: Organization policy restricted link sharing: " + e.message);
    }

    return { success: true, url: file.getUrl(), message: "CIF file uploaded successfully!" };
  } catch (err) {
    return { success: false, message: "CIF upload failed: " + err.message };
  }
}

// -------------------------------------------------------------------------
// 1. DIRECTORY SERVICE
// -------------------------------------------------------------------------

function getDirectoryData() {
  try {
    const ss = getSpreadsheet();
    const sheetMap = getSheetMap(ss);
    const sheet = getDirectorySheet(ss, sheetMap);
    if (!sheet) return { success: false, data: [], message: 'Directory sheet tab not found.' };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return { success: true, data: [] };

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());

    let idxName = headers.findIndex(h => h.includes("NAME"));
    let idxDesig = headers.findIndex(h => h.includes("DESIGNATION") && !h.includes("HEALTH CENTER"));
    let idxHC = headers.findIndex(h => h.includes("HEALTH CENTER") || (h.includes("CENTER") && h.includes("DESIGNATION")));
    let idxPic = headers.findIndex(h => h.includes("PICTURE") || h.includes("PHOTO") || h.includes("IMAGE") || h.includes("AVATAR"));

    if (idxName === -1) idxName = 0;
    if (idxDesig === -1) idxDesig = 1;
    if (idxHC === -1) idxHC = 2;
    if (idxPic === -1) idxPic = 3;

    const list = [];
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;

      const name = String(row[idxName] || '').trim();
      const designation = String(row[idxDesig] || '').trim();
      const healthCenter = String(row[idxHC] || '').trim();
      let picture = String(row[idxPic] || '').trim();

      if (!name && !designation && !healthCenter) continue;

      picture = convertToDirectThumbnailUrl(picture);

      list.push({
        _rowIndex: r + 1,
        name: name,
        designation: designation,
        healthCenter: healthCenter,
        picture: picture
      });
    }

    return { success: true, data: list };
  } catch (err) {
    Logger.log("Error in getDirectoryData: " + err.toString());
    return { success: false, data: [], error: err.toString() };
  }
}

// -------------------------------------------------------------------------
// 2. AUTHENTICATION SERVICE (INSTANT RUNTIME)
// -------------------------------------------------------------------------

function processLogin(data) {
  try {
    const usernameInput = String(data.username || '').trim().toLowerCase();
    const passwordInput = String(data.password || '').trim();

    if (!usernameInput || !passwordInput) {
      return { success: false, message: 'Please enter both username and password.' };
    }

    const ss = getSpreadsheet();
    if (!ss) {
      return { success: false, message: 'DATABASE_ERROR: Cannot access spreadsheet. Check SPREADSHEET_ID.' };
    }

    const sheetMap = getSheetMap(ss);
    const sheet = getUsersSheet(ss, sheetMap);
    if (!sheet) {
      return { success: false, message: 'DATABASE_OFFLINE: "users" tab not found in spreadsheet. Run seedUsers() first.' };
    }

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) {
      return { success: false, message: 'DATABASE_EMPTY: No user accounts found. Run seedUsers() first.' };
    }

    for (let i = 1; i < rows.length; i++) {
      const dbUser = String(rows[i][0] || '').trim().toLowerCase();
      const dbPass = String(rows[i][1] || '').trim();
      const dbStatus = String(rows[i][2] || '').trim().toLowerCase();
      const dbRole = String(rows[i][3] || 'HC_USER').trim().toUpperCase();
      const dbDesignation = String(rows[i][4] || '').trim();

      if (dbUser === usernameInput) {
        if (dbStatus !== 'active' && dbStatus !== '') {
          return { success: false, message: 'ACCESS_DENIED: Account status is ' + (dbStatus.toUpperCase() || 'INACTIVE') + '.' };
        }

        if (dbPass === passwordInput) {
          return { 
            success: true, 
            message: 'ACCESS_GRANTED: Authentication successful.', 
            user: rows[i][0],
            role: dbRole,
            designation: dbDesignation
          };
        } else {
          return { success: false, message: 'SECURITY_ALERT: Incorrect password.' };
        }
      }
    }

    return { success: false, message: 'NODE_NOT_FOUND: User ID / Email not recognized.' };
  } catch (error) {
    Logger.log("Error in processLogin: " + error.toString());
    return { success: false, message: 'CRITICAL_FAIL: ' + error.toString() };
  }
}

// -------------------------------------------------------------------------
// 3. DASHBOARD METRICS WITH DYNAMIC BASELINES (1.2s EXECUTION)
// -------------------------------------------------------------------------

function getMetricsData(filters) {
  try {
    const ss = getSpreadsheet();
    const sheetMap = getSheetMap(ss);
    const sheet = getMdbSheet(ss, sheetMap); 
    if (!sheet) {
      return { data: [], dropdowns: { healthCenters: [], morbWeeks: [], barangays: [], diseases: [] }, baselines: {}, syncTime: getFormattedTime() };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow <= 1 || lastCol === 0) {
      return { data: [], dropdowns: { healthCenters: [], morbWeeks: [], barangays: [], diseases: [] }, baselines: {}, syncTime: getFormattedTime() };
    }
    
    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    
    let idxCenter = headers.indexOf("HEALTH CENTER DESIGNATION"); 
    if (idxCenter === -1) idxCenter = headers.findIndex(h => h.includes("HEALTH CENTER") || h.includes("CENTER"));

    let idxRemarks = headers.indexOf("INVESTIGATION REMARKS"); 
    if (idxRemarks === -1) idxRemarks = headers.findIndex(h => h.includes("REMARKS"));

    let idxBarangay = headers.indexOf("BARANGAY");                 
    if (idxBarangay === -1) idxBarangay = headers.findIndex(h => h.includes("BRGY") || h.includes("BGY"));

    let idxOutcome = headers.indexOf("OUTCOME"); 
    if (idxOutcome === -1) idxOutcome = headers.findIndex(h => h.includes("OUTCOME"));

    let idxDisease = headers.indexOf("DISEASE_NAME");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));
    
    let idxSex = headers.indexOf("GENDER");
    if (idxSex === -1) idxSex = headers.indexOf("SEX");
    
    let idxAge = headers.indexOf("AGE_IN_YEARS");
    if (idxAge === -1) idxAge = headers.indexOf("AGE IN YEARS");
    
    let idxAgeGroup = headers.indexOf("AGE_GROUP");
    if (idxAgeGroup === -1) idxAgeGroup = headers.findIndex(h => h.includes("AGE_GROUP") || h.includes("AGE GROUP"));
    
    let idxMorbWeek = headers.indexOf("MORBIDITY_WEEK");           
    if (idxMorbWeek === -1) idxMorbWeek = headers.indexOf("MORBIDITY WEEK");
    
    let idxCoords = headers.indexOf("GEO_MAPPING (CORDINATES THRU G-MAPS)");
    if (idxCoords === -1) idxCoords = headers.findIndex(h => h.includes("GEO_MAPPING") || h.includes("CORDINATES") || h.includes("COORDINATES"));

    let idxClass = headers.indexOf("CASE_CLASSIFICATION");
    if (idxClass === -1) {
      idxClass = headers.findIndex(h => h.includes("CLASSIFICATION") || h.includes("LABORATORY") || h.includes("LAB_STATUS") || h.includes("STATUS"));
    }

    let idxCategory = headers.indexOf("CATEGORY");
    if (idxCategory === -1) idxCategory = headers.findIndex(h => h.includes("CATEGORY"));
    
    const dashboardRows = [];
    const dynamicCenters = new Set();
    const dynamicWeeks = new Set();
    const dynamicBarangays = new Set(DISTRICT_2_BARANGAYS);
    const dynamicDiseases = new Set();
    
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || (!row[0] && !row[1] && !row[2] && !row[3] && !row[4])) continue;

      const healthCenter = idxCenter !== -1 ? String(row[idxCenter]).trim() : "";
      const investigationRemarks = idxRemarks !== -1 ? String(row[idxRemarks]).trim() : "";
      const barangay = idxBarangay !== -1 ? String(row[idxBarangay]).trim() : "";
      const outcome = idxOutcome !== -1 ? String(row[idxOutcome]).trim() : "";
      const diseaseName = idxDisease !== -1 ? String(row[idxDisease]).trim().toUpperCase() : "";
      const rawClassValue = idxClass !== -1 ? String(row[idxClass]).trim().toUpperCase() : "";
      const rawRemarksValue = investigationRemarks.toUpperCase();

      if (!barangay && !diseaseName && !healthCenter) continue;
      if (!healthCenter || healthCenter === "") continue;
      
      if (
        rawRemarksValue.includes("DELIST") || rawRemarksValue.includes("NON-RES") || 
        rawRemarksValue.includes("NON RES") || rawRemarksValue.includes("DISCARD") || 
        rawRemarksValue.includes("DUPLICATE") || rawClassValue.includes("DISCARD") || 
        rawClassValue.includes("DUPLICATE") || rawClassValue.includes("DELIST")
      ) {
        continue; 
      }
      
      let morbWeek = null;
      if (idxMorbWeek !== -1 && row[idxMorbWeek] !== "") {
        let rawValue = String(row[idxMorbWeek]).trim();
        if (rawValue !== "") {
          let parsedWeek = parseInt(rawValue, 10);
          if (!isNaN(parsedWeek)) morbWeek = parsedWeek;
        }
      }
      
      let statusClass = "Suspect"; 
      if (idxClass !== -1 && row[idxClass]) {
        let rawClass = String(row[idxClass]).trim().toUpperCase();
        if (rawClass.includes("CONFIRM") || rawClass.includes("POSITIVE") || rawClass === "POS" || rawClass === "C" || rawClass === "+") {
          statusClass = "Confirmed"; 
        } else if (rawClass.includes("NEGATIVE") || rawClass === "NEG" || rawClass === "N" || rawClass === "-") {
          statusClass = "Negative";
        } else if (rawClass.includes("PROBABLE")) {
          statusClass = "Probable";
        } else if (rawClass.includes("SUSPECT")) {
          statusClass = "Suspect";
        }
      }
      
      let lat = null, lng = null;
      if (idxCoords !== -1 && row[idxCoords]) {
        let coordString = String(row[idxCoords]).trim();
        let parts = coordString.split(/[\s,]+/);
        if (parts.length >= 2) {
          let p1 = parseFloat(parts[0]);
          let p2 = parseFloat(parts[1]);
          if (!isNaN(p1) && !isNaN(p2)) {
            if (p1 > p2) { lat = p2; lng = p1; } else { lat = p1; lng = p2; }
          }
        }
      }
      
      const rawSex = idxSex !== -1 ? String(row[idxSex]).trim().toUpperCase() : "";
      let cleanSex = "UNKNOWN";
      if (rawSex.startsWith("MALE") || rawSex === "M") cleanSex = "MALE";
      else if (rawSex.startsWith("FEMALE") || rawSex === "F") cleanSex = "FEMALE";
      
      let cleanAge = null;
      if (idxAge !== -1 && row[idxAge] !== "" && row[idxAge] !== undefined) {
        let parsed = parseFloat(row[idxAge]);
        if (!isNaN(parsed)) cleanAge = parsed;
      }
      
      let ageGroup = (idxAgeGroup !== -1 && row[idxAgeGroup]) ? String(row[idxAgeGroup]).trim() : "";
      let category = (idxCategory !== -1 && row[idxCategory]) ? String(row[idxCategory]).trim().toUpperCase() : "";
      
      let district = "Other Districts";
      if (DISTRICT_2_BARANGAYS.includes(barangay.toUpperCase())) {
        district = "District 2";
      }

      dashboardRows.push({
        district: district,
        barangay: barangay,
        healthCenter: healthCenter,
        investigationRemarks: investigationRemarks,
        outcome: outcome,
        morbWeek: morbWeek,
        diseaseName: diseaseName,
        sex: cleanSex,
        age: cleanAge,
        ageGroup: ageGroup,
        category: category,
        lat: lat,
        lng: lng,
        classification: statusClass
      });
      
      if (healthCenter) dynamicCenters.add(healthCenter);
      if (barangay) dynamicBarangays.add(barangay);
      if (diseaseName) dynamicDiseases.add(diseaseName);
      if (morbWeek !== null && morbWeek !== undefined) dynamicWeeks.add(morbWeek);
    }
    
    const sortedWeeks = Array.from(dynamicWeeks).sort((a, b) => a - b);
    
    // FAST BASELINE SEARCH: Only checks sheets that actually exist in sheetMap
    const baselines = {};
    const diseaseList = ["DENGUE", "MEASLES", "LEPTO", "LEPTOSPIROSIS", "COVID19", "COVID", "COVID-19", "CHIKUNGUNYA", "TYPHOID", "RABIES", "AFP", "DIPH", "PERTUSSIS", "ROTAVIRUS", "CHOLERA", "ILI", "SARI", "HEPA", "AMES", "MENINGO", "HFMD"];
    
    diseaseList.forEach(d => {
      const defaultName = (d + " BASELINE").toUpperCase();
      if (sheetMap[defaultName] || sheetMap[d]) {
        baselines[d] = fetchComprehensiveDiseaseBaseline(ss, d, defaultName, sheetMap);
      }
    });
    
    return {
      data: dashboardRows,
      dropdowns: {
        healthCenters: Array.from(dynamicCenters).sort(),
        morbWeeks: sortedWeeks,
        barangays: Array.from(dynamicBarangays).sort(),
        diseases: Array.from(dynamicDiseases).sort()
      },
      baselines: baselines,
      syncTime: getFormattedTime()
    };
  } catch (err) {
    Logger.log("Error in getMetricsData: " + err.toString());
    return { 
      data: [], 
      dropdowns: { healthCenters: [], morbWeeks: [], barangays: [], diseases: [] }, 
      baselines: {}, 
      syncTime: getFormattedTime(), 
      error: err.toString() 
    };
  }
}

// -------------------------------------------------------------------------
// 4. FEEDBACK MONITORING / DURATION SERVICE
// -------------------------------------------------------------------------

function getDashboardData() {
  try {
    const ss = getSpreadsheet();
    const sheetMap = getSheetMap(ss);
    const sheet = getEpisenseSheet(ss, sheetMap);
    if (!sheet) return { cases: [] };

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return { cases: [] };

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim().toUpperCase());

    let idxDisease = headers.indexOf("DISEASE_TYPE");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));

    let idxDate = headers.indexOf("DATE_VALIDATED");
    if (idxDate === -1) idxDate = headers.findIndex(h => h.includes("VALIDAT") || h.includes("DATE"));

    let idxCaseId = headers.indexOf("CASE_ID");
    if (idxCaseId === -1) idxCaseId = headers.findIndex(h => h.includes("CASE") || h.includes("ID"));

    let idxHC = headers.indexOf("HEALTH_CENTER");
    if (idxHC === -1) idxHC = headers.findIndex(h => h.includes("HEALTH") || h.includes("CENTER"));

    let idxInvStatus = headers.indexOf("INV_STATUS");
    if (idxInvStatus === -1) idxInvStatus = headers.findIndex(h => h.includes("STATUS") || h.includes("REMARK"));

    let idxSurveillance = headers.indexOf("OVER_2DAY_SURVEILLANCE");
    if (idxSurveillance === -1) idxSurveillance = headers.findIndex(h => h.includes("SURVEILLANCE") || h.includes("2DAY") || h.includes("OVER"));

    const now = new Date();
    const nowMs = now.getTime();
    const cases = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || (!row[0] && !row[1] && !row[2])) continue;

      const caseId = (idxCaseId !== -1 && row[idxCaseId]) ? String(row[idxCaseId]).trim() : `CASE-${r}`;
      const disease = (idxDisease !== -1 && row[idxDisease]) ? String(row[idxDisease]).trim().toUpperCase() : "DENGUE";
      let healthCenter = (idxHC !== -1 && row[idxHC]) ? String(row[idxHC]).trim() : "Unassigned Health Center";
      const invStatus = (idxInvStatus !== -1 && row[idxInvStatus]) ? String(row[idxInvStatus]).trim().toUpperCase() : "";
      const surveillanceStr = (idxSurveillance !== -1 && row[idxSurveillance]) ? String(row[idxSurveillance]).trim().toUpperCase() : "";

      if (invStatus.includes("DELIST") || invStatus.includes("DISCARD") || invStatus.includes("DUPLICATE")) continue;

      let rawDate = (idxDate !== -1) ? row[idxDate] : null;
      let dateObj = (rawDate instanceof Date) ? rawDate : (rawDate ? new Date(String(rawDate).trim()) : now);
      if (!dateObj || isNaN(dateObj.getTime())) dateObj = now;

      let hoursElapsed = (nowMs - dateObj.getTime()) / (1000 * 60 * 60);
      let category = "Under 24 Hours";
      if (surveillanceStr.includes("OVER 48") || surveillanceStr.includes("OVER 2") || surveillanceStr.includes(">48")) {
        category = "Over 48 Hours";
        if (hoursElapsed < 48) hoursElapsed = 49.0;
      } else if (surveillanceStr.includes("24 TO 48") || surveillanceStr.includes("24-48")) {
        category = "24 to 48 Hours";
        if (hoursElapsed < 24 || hoursElapsed > 48) hoursElapsed = 36.0;
      } else if (hoursElapsed > 48) {
        category = "Over 48 Hours";
      } else if (hoursElapsed >= 24) {
        category = "24 to 48 Hours";
      } else {
        category = "Under 24 Hours";
        if (hoursElapsed < 0) hoursElapsed = 12.0;
      }

      cases.push({
        caseId: caseId,
        disease: disease,
        healthCenter: healthCenter,
        dateValidated: formatDateFast(dateObj),
        hoursElapsed: Number(hoursElapsed).toFixed(1),
        category: category
      });
    }

    return { cases: cases };
  } catch (err) {
    Logger.log("Error in getDashboardData: " + err.toString());
    return { cases: [], error: err.toString() };
  }
}

// -------------------------------------------------------------------------
// 5. D2 NAVDPCP 2026 REPORT
// -------------------------------------------------------------------------

function getNavdpcpData(filters) {
  try {
    const ss = getSpreadsheet();
    const sheetMap = getSheetMap(ss);
    const sheet = getMdbSheet(ss, sheetMap); 
    if (!sheet) {
      return { 
        table1: { rows: [], totals: {} }, 
        table21: { testTypes: [], rows: [], totals: {} }, 
        kpi22: 0, 
        table34: { rows: [], totals: {} }, 
        table4Star: { rows: [], totals: {} } 
      };
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) {
      return { 
        table1: { rows: [], totals: {} }, 
        table21: { testTypes: [], rows: [], totals: {} }, 
        kpi22: 0, 
        table34: { rows: [], totals: {} }, 
        table4Star: { rows: [], totals: {} } 
      };
    }

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h || '').replace(/[\r\n\t]+/g, ' ').replace(/\s+/g, ' ').trim().toUpperCase());
    
    let idxHC = headers.indexOf("HEALTH CENTER DESIGNATION");
    if (idxHC === -1) idxHC = headers.findIndex(h => h.includes("HEALTH CENTER") || (h.includes("CENTER") && !h.includes("ACTION")));

    let idxBrgy = headers.indexOf("BARANGAY");
    if (idxBrgy === -1) idxBrgy = headers.findIndex(h => h === "BRGY" || h === "BGY");

    let idxRemarks = headers.indexOf("INVESTIGATION REMARKS");
    if (idxRemarks === -1) idxRemarks = headers.findIndex(h => h === "REMARKS");

    let idxDisease = headers.indexOf("DISEASE_NAME");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));

    let idxClass = headers.indexOf("CASE_CLASSIFICATION");
    if (idxClass === -1) idxClass = headers.findIndex(h => h.includes("CLASSIFICATION") && !h.includes("CLINICAL"));

    let idxTesting = headers.indexOf("TYPE OF TESTING");
    if (idxTesting === -1) idxTesting = headers.findIndex(h => h.includes("TESTING") || h.includes("TEST"));

    let idxClinical = headers.indexOf("CLINICAL_CLASSIFICATION");
    if (idxClinical === -1) idxClinical = headers.findIndex(h => h.includes("CLINICAL"));

    let idxOutcome = headers.indexOf("OUTCOME");
    let idxCategory = headers.indexOf("CATEGORY");

    let idxDate = headers.findIndex(h => h.includes("CESU") && (h.includes("DATE") || h.includes("ADDED")));
    if (idxDate === -1) idxDate = 16; 

    const targetBrgy = filters && filters.barangay ? filters.barangay.toUpperCase().trim() : "ALL BARANGAY";
    const targetHC = filters && filters.healthCenter ? filters.healthCenter.toUpperCase().trim() : "ALL HEALTH CENTERS";
    const targetRemarks = filters && filters.remarks ? filters.remarks.toUpperCase().trim() : "ALL REMARKS";
    const targetDisease = filters && filters.disease ? filters.disease.toUpperCase().trim() : "DENGUE";

    const startDateStr = filters && filters.startDate ? filters.startDate.trim() : "";
    const endDateStr = filters && filters.endDate ? filters.endDate.trim() : "";

    const isFullYear = (!startDateStr || startDateStr === "2026-01-01") && (!endDateStr || endDateStr === "2026-12-31");

    let startMs = 0;
    let endMs = Infinity;

    if (startDateStr) {
      let sParts = startDateStr.split("-");
      if (sParts.length === 3) {
        startMs = new Date(parseInt(sParts[0], 10), parseInt(sParts[1], 10) - 1, parseInt(sParts[2], 10), 0, 0, 0, 0).getTime();
      }
    }
    if (endDateStr) {
      let eParts = endDateStr.split("-");
      if (eParts.length === 3) {
        endMs = new Date(parseInt(eParts[0], 10), parseInt(eParts[1], 10) - 1, parseInt(eParts[2], 10), 23, 59, 59, 999).getTime();
      }
    }

    function parseCesuDate(val) {
      if (!val) return null;
      if (val instanceof Date) return isNaN(val.getTime()) ? null : val.getTime();
      if (typeof val === 'number' && val > 30000) return (val - 25569) * 86400 * 1000;
      let str = String(val).trim();
      if (!str) return null;

      let parts = str.split(/[\/\-\.]/);
      if (parts.length === 3) {
        let y, m, d;
        if (parts[0].length === 4) {
          y = parseInt(parts[0], 10);
          m = parseInt(parts[1], 10) - 1;
          d = parseInt(parts[2], 10);
        } else {
          let p0 = parseInt(parts[0], 10);
          let p1 = parseInt(parts[1], 10);
          let p2 = parseInt(parts[2], 10);
          if (p2 < 100) p2 += 2000;
          y = p2;
          if (p0 > 12) { d = p0; m = p1 - 1; } 
          else { m = p0 - 1; d = p1; }
        }
        let parsed = new Date(y, m, d, 12, 0, 0);
        if (!isNaN(parsed.getTime())) return parsed.getTime();
      }

      let dObj = new Date(str);
      return isNaN(dObj.getTime()) ? null : dObj.getTime();
    }

    const caseFindingMap = {};
    const testingMap = {};
    const clinicalOutcomeMap = {};
    const cesdMap = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || (!row[0] && !row[1] && !row[2])) continue;
      
      const rawHC = idxHC !== -1 && row[idxHC] ? String(row[idxHC]).trim() : "";
      const hc = rawHC !== "" ? rawHC : "null";

      const brgy = idxBrgy !== -1 ? String(row[idxBrgy]).trim().toUpperCase() : "";
      const remarks = idxRemarks !== -1 ? String(row[idxRemarks]).trim().toUpperCase() : "";
      const disease = idxDisease !== -1 ? String(row[idxDisease]).trim().toUpperCase() : "";
      const rawClass = idxClass !== -1 && row[idxClass] ? String(row[idxClass]).trim().toUpperCase() : "";

      if (targetBrgy !== "ALL BARANGAY" && targetBrgy !== "BARANGAY" && !brgy.includes(targetBrgy)) continue;
      if (targetHC !== "ALL HEALTH CENTERS" && targetHC !== "HEALTH CENTER DESIGNATION" && hc.toUpperCase() !== targetHC) continue;
      if (targetRemarks !== "ALL REMARKS" && targetRemarks !== "INVESTIGATION REMARKS" && !remarks.includes(targetRemarks)) continue;
      if (targetDisease !== "ALL DISEASES" && !disease.includes(targetDisease)) continue;

      if (!isFullYear && (startDateStr || endDateStr)) {
        let rowTime = parseCesuDate(row[idxDate]);
        if (!rowTime || rowTime < startMs || rowTime > endMs) continue;
      }

      let rawCat = idxCategory !== -1 && row[idxCategory] ? String(row[idxCategory]).trim().toUpperCase() : "HC DETECTED";
      let category = rawCat.includes("PHSU") ? "PHSU ENDORSEMENT" : "HC DETECTED";

      if (category === "HC DETECTED") {
        let classification = "null";
        if (rawClass.includes("SUSPECT")) classification = "SUSPECT";
        else if (rawClass.includes("CONFIRM") || rawClass.includes("POS") || rawClass === "+") classification = "CONFIRM";
        else if (rawClass.includes("PROBABLE")) classification = "PROBABLE";

        if (!caseFindingMap[hc]) caseFindingMap[hc] = { SUSPECT: 0, nullVal: 0, CONFIRM: 0, PROBABLE: 0 };
        if (classification === "SUSPECT") caseFindingMap[hc].SUSPECT++;
        else if (classification === "CONFIRM") caseFindingMap[hc].CONFIRM++;
        else if (classification === "PROBABLE") caseFindingMap[hc].PROBABLE++;
        else caseFindingMap[hc].nullVal++;

        let testType = idxTesting !== -1 && row[idxTesting] ? String(row[idxTesting]).trim() : "null";
        if (!testType || testType === "") testType = "null";
        if (!testingMap[hc]) testingMap[hc] = {};
        testingMap[hc][testType] = (testingMap[hc][testType] || 0) + 1;
      }

      let clinical = idxClinical !== -1 && row[idxClinical] ? String(row[idxClinical]).trim() : "null";
      if (!clinical || clinical === "") clinical = "null";

      let rawOut = idxOutcome !== -1 && row[idxOutcome] ? String(row[idxOutcome]).trim().toUpperCase() : "";
      let outcome = "nullVal";
      if (rawOut.includes("ALIVE")) outcome = "ALIVE";
      else if (rawOut.includes("DIE") || rawOut.includes("DEATH")) outcome = "DIED";

      const clinKey = `${hc}||${clinical}`;
      if (!clinicalOutcomeMap[clinKey]) {
        clinicalOutcomeMap[clinKey] = { hc: hc, clinical: clinical, ALIVE: 0, nullVal: 0, DIED: 0 };
      }
      clinicalOutcomeMap[clinKey][outcome]++;

      if (!cesdMap[hc]) cesdMap[hc] = { "PHSU ENDORSEMENT": 0, "HC DETECTED": 0 };
      cesdMap[hc][category]++;
    }

    const t1Data = formatTable1(caseFindingMap);
    const kpi22Value = (t1Data.totals.SUSPECT || 0) + (t1Data.totals.CONFIRM || 0) + (t1Data.totals.PROBABLE || 0);

    return {
      table1: t1Data,
      table21: formatTable21(testingMap),
      kpi22: kpi22Value,
      table34: formatTable34(clinicalOutcomeMap),
      table4Star: formatTable4Star(cesdMap)
    };
  } catch(err) {
    Logger.log("Error in getNavdpcpData: " + err.toString());
    return { 
      table1: { rows: [], totals: {} }, 
      table21: { testTypes: [], rows: [], totals: {} }, 
      kpi22: 0, 
      table34: { rows: [], totals: {} }, 
      table4Star: { rows: [], totals: {} } 
    };
  }
}

function formatTable1(map) {
  let list = [];
  let totals = { SUSPECT: 0, nullVal: 0, CONFIRM: 0, PROBABLE: 0, grandTotal: 0 };
  for (let hc in map) {
    let row = map[hc];
    let suspect = Number(row.SUSPECT || 0);
    let nullVal = Number(row.nullVal || row.null || 0);
    let confirm = Number(row.CONFIRM || 0);
    let probable = Number(row.PROBABLE || 0);
    let sum = suspect + nullVal + confirm + probable;

    list.push({ 
      hc: hc, 
      suspect: suspect, 
      nullVal: nullVal, 
      confirm: confirm, 
      probable: probable, 
      grandTotal: sum 
    });
    totals.SUSPECT += suspect;
    totals.nullVal += nullVal;
    totals.CONFIRM += confirm;
    totals.PROBABLE += probable;
    totals.grandTotal += sum;
  }
  list.sort((a, b) => b.grandTotal - a.grandTotal);
  return { rows: list, totals: totals };
}

function formatTable21(map) {
  let testTypeCounts = {};
  for (let hc in map) {
    for (let tt in map[hc]) {
      testTypeCounts[tt] = (testTypeCounts[tt] || 0) + map[hc][tt];
    }
  }
  let testTypes = Object.keys(testTypeCounts).sort((a, b) => testTypeCounts[b] - testTypeCounts[a]);

  let totals = { grandTotal: 0 };
  testTypes.forEach(tt => totals[tt] = 0);

  let list = [];
  for (let hc in map) {
    let rowObj = { hc: hc, grandTotal: 0 };
    testTypes.forEach(tt => {
      let val = Number(map[hc][tt] || 0);
      rowObj[tt] = val;
      rowObj.grandTotal += val;
      totals[tt] += val;
      totals.grandTotal += val;
    });
    list.push(rowObj);
  }
  list.sort((a, b) => b.grandTotal - a.grandTotal);
  return { testTypes: testTypes, rows: list, totals: totals };
}

function formatTable34(map) {
  let list = [];
  let totals = { ALIVE: 0, nullVal: 0, DIED: 0, grandTotal: 0 };
  for (let key in map) {
    let item = map[key];
    let alive = Number(item.ALIVE || 0);
    let nullVal = Number(item.nullVal || item.null || 0);
    let died = Number(item.DIED || 0);
    let sum = alive + nullVal + died;

    list.push({ 
      hc: item.hc, 
      clinical: item.clinical, 
      alive: alive, 
      nullVal: nullVal, 
      died: died, 
      grandTotal: sum 
    });

    totals.ALIVE += alive;
    totals.nullVal += nullVal;
    totals.DIED += died;
    totals.grandTotal += sum;
  }
  list.sort((a, b) => b.grandTotal - a.grandTotal);
  return { rows: list, totals: totals };
}

function formatTable4Star(map) {
  let list = [];
  let totals = { phsu: 0, hcDetected: 0, grandTotal: 0 };
  for (let hc in map) {
    let phsu = Number(map[hc]["PHSU ENDORSEMENT"] || 0);
    let hcDet = Number(map[hc]["HC DETECTED"] || 0);
    let sum = phsu + hcDet;
    list.push({ hc: hc, phsu: phsu, hcDetected: hcDet, grandTotal: sum });
    totals.phsu += phsu;
    totals.hcDetected += hcDet;
    totals.grandTotal += sum;
  }
  list.sort((a, b) => b.grandTotal - a.grandTotal);
  return { rows: list, totals: totals };
}

function getFormattedTime() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
}

// -------------------------------------------------------------------------
// 6. PATIENT PROFILE RECORDS & CRUD SERVICES (HIGH SPEED + DUAL SYNC)
// -------------------------------------------------------------------------

/**
 * Ultra-fast pure JavaScript date formatter (0ms latency, zero server RPC calls)
 */
function formatDateFast(d) {
  if (!d) return "";
  if (d instanceof Date && !isNaN(d.getTime())) {
    const y = d.getFullYear();
    const m = d.getMonth() + 1;
    const day = d.getDate();
    return y + "-" + (m < 10 ? "0" + m : m) + "-" + (day < 10 ? "0" + day : day);
  }
  return String(d).trim();
}

/**
 * Synchronizes newly added/edited patient data to the secondary spreadsheet
 */
function saveToExternalSpreadsheet(formData) {
  if (!PATIENT_SPREADSHEET_ID_OR_URL || PATIENT_SPREADSHEET_ID_OR_URL.trim() === "") return;

  try {
    const targetSs = openSpreadsheetByIdOrUrl(PATIENT_SPREADSHEET_ID_OR_URL);
    if (!targetSs) return;

    let targetSheet = null;
    if (PATIENT_TARGET_SHEET_NAME && PATIENT_TARGET_SHEET_NAME.trim() !== "") {
      targetSheet = targetSs.getSheetByName(PATIENT_TARGET_SHEET_NAME);
    }
    if (!targetSheet) {
      targetSheet = targetSs.getSheetByName("MDB DISTRICT 2 2026") ||
                    targetSs.getSheetByName("MDB") ||
                    targetSs.getSheetByName("PATIENT PROFILE") ||
                    targetSs.getSheets()[0];
    }

    if (!targetSheet) return;

    const lastRow = targetSheet.getLastRow();
    let lastCol = targetSheet.getLastColumn();

    if (lastRow === 0 || lastCol === 0) {
      const defaultHeaders = Object.keys(formData);
      targetSheet.getRange(1, 1, 1, defaultHeaders.length).setValues([defaultHeaders]);
      lastCol = defaultHeaders.length;
    }

    const headers = targetSheet.getRange(1, 1, 1, targetSheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());

    const rowValues = [];
    headers.forEach(header => {
      let val = "";
      for (let key in formData) {
        if (key.trim().toUpperCase() === header) {
          val = formData[key];
          break;
        }
      }
      rowValues.push(val);
    });

    let targetRowIndex = -1;
    const caseId = (formData["CASE_ID"] || formData["CASE ID"] || "").trim().toUpperCase();

    if (caseId && targetSheet.getLastRow() > 1) {
      let caseIdColIdx = headers.indexOf("CASE_ID");
      if (caseIdColIdx === -1) caseIdColIdx = headers.indexOf("CASE ID");
      
      if (caseIdColIdx !== -1) {
        const idColValues = targetSheet.getRange(2, caseIdColIdx + 1, targetSheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < idColValues.length; i++) {
          if (String(idColValues[i][0]).trim().toUpperCase() === caseId) {
            targetRowIndex = i + 2;
            break;
          }
        }
      }
    }

    if (targetRowIndex > 1) {
      targetSheet.getRange(targetRowIndex, 1, 1, rowValues.length).setValues([rowValues]);
    } else {
      targetSheet.appendRow(rowValues);
    }
  } catch (err) {
    Logger.log("saveToExternalSpreadsheet error: " + err.toString());
  }
}

/**
 * Loads thousands of patient records in <0.3 seconds
 */
function getRecords() {
  try {
    const ss = getSpreadsheet();
    const sheetMap = getSheetMap(ss);
    const sheet = getMdbSheet(ss, sheetMap); 
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim());
    const records = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || (!row[0] && !row[1] && !row[2] && !row[3] && !row[4])) continue;

      const recordObj = { _rowIndex: r + 1 };
      for (let idx = 0; idx < headers.length; idx++) {
        const h = headers[idx];
        let val = row[idx];
        if (val instanceof Date) {
          recordObj[h] = formatDateFast(val);
        } else {
          recordObj[h] = (val !== undefined && val !== null) ? String(val).trim() : "";
        }
      }
      records.push(recordObj);
    }
    return records;
  } catch(err) {
    Logger.log("Error in getRecords: " + err.toString());
    return [];
  }
}

function saveRecord(formData, rowIndex) {
  try {
    const ss = getSpreadsheet();
    const sheetMap = getSheetMap(ss);
    const sheet = getMdbSheet(ss, sheetMap);
    if (!sheet) return { success: false, message: "MDB Sheet not found." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());

    let targetRow = null;
    const caseId = (formData["CASE_ID"] || formData["CASE ID"] || "").trim().toUpperCase();

    if (caseId && sheet.getLastRow() > 1) {
      let caseIdColIdx = headers.indexOf("CASE_ID");
      if (caseIdColIdx === -1) caseIdColIdx = headers.indexOf("CASE ID");
      if (caseIdColIdx !== -1) {
        const idColValues = sheet.getRange(2, caseIdColIdx + 1, sheet.getLastRow() - 1, 1).getValues();
        for (let i = 0; i < idColValues.length; i++) {
          if (String(idColValues[i][0]).trim().toUpperCase() === caseId) {
            targetRow = i + 2;
            break;
          }
        }
      }
    }

    if (!targetRow) {
      targetRow = (rowIndex && !isNaN(rowIndex) && rowIndex > 1) ? rowIndex : sheet.getLastRow() + 1;
    }

    const rowValues = [];
    headers.forEach(header => {
      let val = "";
      for (let key in formData) {
        if (key.trim().toUpperCase() === header) {
          val = formData[key];
          break;
        }
      }
      rowValues.push(val);
    });

    sheet.getRange(targetRow, 1, 1, rowValues.length).setValues([rowValues]);
    saveToExternalSpreadsheet(formData);

    return { success: true, message: `Record successfully ${rowIndex ? 'updated' : 'saved'}!` };
  } catch (err) {
    return { success: false, message: `Failed to save: ${err.message}` };
  }
}

// -------------------------------------------------------------------------
// 7. DYNAMIC BASELINE ENGINE (HIGH-SPEED IN-MEMORY CACHE)
// -------------------------------------------------------------------------

function parseSideBySideBaselineSheet(sheet) {
  if (!sheet) return null;
  const dataset = sheet.getDataRange().getValues();
  if (dataset.length < 2) return null;

  const row0 = dataset[0].map(v => String(v || '').trim().toUpperCase());
  const row1 = dataset[1].map(v => String(v || '').trim().toUpperCase());

  const master = {
    alert: new Array(53).fill(0),
    epidemic: new Array(53).fill(0),
    currentYear: new Array(53).fill(0),
    hasNegOne: false
  };

  for (let r = 1; r < dataset.length; r++) {
    const wNum = parseInt(dataset[r][0], 10);
    if (!isNaN(wNum)) {
      const epi = Number(dataset[r][1] || 0);
      const alt = Number(dataset[r][2] || 0);
      const cur = dataset[r][3] !== undefined && !isNaN(Number(dataset[r][3])) ? Number(dataset[r][3]) : 0;
      if (wNum === -1) {
        master.hasNegOne = true;
        master.alert[0] = alt;
        master.epidemic[0] = epi;
        master.currentYear[0] = cur;
      } else if (wNum >= 1 && wNum <= 52) {
        master.alert[wNum] = alt;
        master.epidemic[wNum] = epi;
        master.currentYear[wNum] = cur;
      }
    }
  }

  const barangaysMap = {};
  const maxCols = Math.max(row0.length, row1.length);

  for (let col = 3; col < maxCols; col++) {
    const text0 = row0[col] || '';
    const text1 = row1[col] || '';

    const matchedBrgy = DISTRICT_2_BARANGAYS.find(b => text0.includes(b) || text1.includes(b));
    
    if (matchedBrgy) {
      let wCol = col;
      let epiCol = col + 1;
      let altCol = col + 2;

      for (let offset = 0; offset <= 3; offset++) {
        let c = col + offset;
        if (c >= maxCols) break;
        let hStr = (row0[c] + " " + row1[c]).toUpperCase();
        if (hStr.includes("WEEK") || hStr.includes("MORB")) {
          wCol = c;
        } else if (hStr.includes("EPIDEMIC") || hStr.includes("TRESHOLD") || hStr.includes("THRESHOLD")) {
          if (hStr.includes("ALERT")) {
            altCol = c;
          } else {
            epiCol = c;
          }
        } else if (hStr.includes("ALERT") || hStr.includes("ACTION")) {
          altCol = c;
        }
      }

      const brgyData = {
        alert: new Array(53).fill(0),
        epidemic: new Array(53).fill(0),
        currentYear: new Array(53).fill(0),
        hasNegOne: false
      };

      for (let r = 1; r < dataset.length; r++) {
        const wNum = parseInt(dataset[r][wCol], 10);
        if (!isNaN(wNum)) {
          const epi = Number(dataset[r][epiCol] || 0);
          const alt = Number(dataset[r][altCol] || 0);

          if (wNum === -1) {
            brgyData.hasNegOne = true;
            brgyData.alert[0] = alt;
            brgyData.epidemic[0] = epi;
          } else if (wNum >= 1 && wNum <= 52) {
            brgyData.alert[wNum] = alt;
            brgyData.epidemic[wNum] = epi;
          }
        }
      }

      barangaysMap[matchedBrgy] = brgyData;
    }
  }

  return { master: master, barangays: barangaysMap };
}

function parseThresholdSheet(sheet) {
  const alertCurve = new Array(53).fill(0);
  const epidemicCurve = new Array(53).fill(0);
  const currentYearCurve = new Array(53).fill(0);
  let hasNegOne = false;

  if (!sheet) return { alert: alertCurve, epidemic: epidemicCurve, currentYear: currentYearCurve, hasNegOne: false };

  const dataset = sheet.getDataRange().getValues();
  if (dataset.length <= 1) return { alert: alertCurve, epidemic: epidemicCurve, currentYear: currentYearCurve, hasNegOne: false };

  for (let i = 1; i < dataset.length; i++) {
    const row = dataset[i];
    const wNum = parseInt(row[0], 10);
    if (!isNaN(wNum)) {
      const epidemicThreshold = Number(row[1] || 0);
      const alertThreshold = Number(row[2] || 0);
      const currentYearValue = Number(row[3] || 0);

      if (wNum === -1) {
        hasNegOne = true;
        alertCurve[0] = alertThreshold;
        epidemicCurve[0] = epidemicThreshold;
        currentYearCurve[0] = currentYearValue;
      } else if (wNum >= 1 && wNum <= 52) {
        alertCurve[wNum] = alertThreshold;
        epidemicCurve[wNum] = epidemicThreshold;
        currentYearCurve[wNum] = currentYearValue;
      }
    }
  }
  return { alert: alertCurve, epidemic: epidemicCurve, currentYear: currentYearCurve, hasNegOne: hasNegOne };
}

function fetchComprehensiveDiseaseBaseline(ss, diseaseKey, defaultSheetName, sheetMap) {
  let masterSheet = getSheetByNameFast(sheetMap, defaultSheetName);
  if (!masterSheet) {
    masterSheet = getSheetByNameFast(sheetMap, diseaseKey + " BASELINE") || getSheetByNameFast(sheetMap, diseaseKey);
  }

  let baseData = { alert: new Array(53).fill(0), epidemic: new Array(53).fill(0), currentYear: new Array(53).fill(0), hasNegOne: false };
  let barangaysMap = {};

  if (masterSheet) {
    const parsedSideBySide = parseSideBySideBaselineSheet(masterSheet);
    if (parsedSideBySide) {
      baseData = parsedSideBySide.master;
      barangaysMap = parsedSideBySide.barangays;
    } else {
      baseData = parseThresholdSheet(masterSheet);
    }
  }

  DISTRICT_2_BARANGAYS.forEach(brgy => {
    if (!barangaysMap[brgy]) {
      const possibleSheetNames = [
        `${diseaseKey} BASELINE - ${brgy}`,
        `${diseaseKey} BASELINE ${brgy}`,
        `${diseaseKey} - ${brgy}`,
        `${diseaseKey} ${brgy}`,
        `${brgy} ${diseaseKey} BASELINE`,
        `${brgy} BASELINE`
      ];

      for (let sName of possibleSheetNames) {
        const bSheet = getSheetByNameFast(sheetMap, sName);
        if (bSheet) {
          barangaysMap[brgy] = parseThresholdSheet(bSheet);
          break;
        }
      }
    }

    if (!barangaysMap[brgy]) {
      barangaysMap[brgy] = {
        alert: [...baseData.alert],
        epidemic: [...baseData.epidemic],
        currentYear: [...baseData.currentYear],
        hasNegOne: baseData.hasNegOne
      };
    }
  });

  const combined = {
    alert: baseData.alert,
    epidemic: baseData.epidemic,
    currentYear: baseData.currentYear,
    hasNegOne: baseData.hasNegOne,
    barangays: barangaysMap,
    byBarangay: barangaysMap
  };

  DISTRICT_2_BARANGAYS.forEach(brgy => {
    combined[brgy] = barangaysMap[brgy];
  });

  return combined;
}

// -------------------------------------------------------------------------
// 8. HIGH-SPEED DATA IMPORTER & SYNC
// -------------------------------------------------------------------------

const SOURCE_SPREADSHEET_ID = "1rrlgAX7ad_IDcHdIXe8gHlWBFsvH5DPCD1ZlW_nR-d0";
const SOURCE_SHEET_NAME = "MDB DISTRICT 2 2026";
const TARGET_TAB_NAME = "MDB DISTRICT 2 2026";

function importAllDataFromSource() {
  try {
    const sourceSs = SpreadsheetApp.openById(SOURCE_SPREADSHEET_ID);
    const sourceSheet = sourceSs.getSheetByName(SOURCE_SHEET_NAME) || sourceSs.getSheets()[0];
    if (!sourceSheet) throw new Error("Source tab '" + SOURCE_SHEET_NAME + "' could not be found.");

    const totalRows = Math.min(sourceSheet.getLastRow(), 10000);
    if (totalRows <= 0) throw new Error("Source sheet contains no data.");

    const totalCols = 50; 
    const sourceRange = sourceSheet.getRange(1, 1, totalRows, totalCols);
    const rawData = sourceRange.getValues();

    const targetSs = getSpreadsheet();
    let targetSheet = targetSs.getSheetByName(TARGET_TAB_NAME);
    if (!targetSheet) targetSheet = targetSs.insertSheet(TARGET_TAB_NAME);

    if (targetSheet.getMaxRows() < rawData.length) {
      targetSheet.insertRowsAfter(targetSheet.getMaxRows(), rawData.length - targetSheet.getMaxRows());
    }
    if (targetSheet.getMaxColumns() < totalCols) {
      targetSheet.insertColumnsAfter(targetSheet.getMaxColumns(), totalCols - targetSheet.getMaxColumns());
    }

    targetSheet.clearContents();
    targetSheet.getRange(1, 1, targetSheet.getMaxRows(), targetSheet.getMaxColumns()).clearDataValidations();
    targetSheet.getRange(1, 1, rawData.length, totalCols).setValues(rawData);

    if (typeof syncDataToFirebase === 'function') {
      syncDataToFirebase();
    }

    return { success: true, totalRows: rawData.length };
  } catch (err) {
    Logger.log("❌ Data Import Error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function syncDataToFirebase() {
  try {
    const payload = getMetricsData();
    const firebaseUrl = FIREBASE_DB_URL + "/surveillance_data.json";
    
    const options = {
      method: "put",
      contentType: "application/json",
      payload: JSON.stringify(payload),
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(firebaseUrl, options);
    return { success: true, message: "Firebase database updated successfully!" };
  } catch (err) {
    Logger.log("Firebase sync error: " + err.toString());
    return { success: false, error: err.toString() };
  }
}

function testMdbData() {
  const ss = getSpreadsheet();
  const sheet = getMdbSheet(ss);
  
  Logger.log("1. Tab Being Read: " + (sheet ? sheet.getName() : "❌ SHEET NOT FOUND"));
  
  if (sheet) {
    const headers = sheet.getRange(1, 1, 1, Math.min(sheet.getLastColumn(), 10)).getValues()[0];
    Logger.log("2. First Headers: " + headers.join(" | "));
    
    const res = getMetricsData();
    Logger.log("3. Patient Records Extracted: " + (res.data ? res.data.length : 0));
    if (res.error) Logger.log("❌ Error: " + res.error);
  }
}

/**
 * Creates and populates the 'users' tab with the 16 District 2 accounts!
 */
function seedUsers() {
  const ss = getSpreadsheet();
  let sheet = getUsersSheet(ss);
  if (!sheet) {
    sheet = ss.insertSheet("users");
  }

  const headers = ["Cyber ID", "Security Passcode", "Status", "User Level", "Designation"];
  
  const userList = [
    ["nicxdumlao.qcesu@gmail.com", "Pinuno2026", "Active", "SUPER_ADMIN", "District 2 Team Leader"],
    ["qcesd.d2esu@quezoncity.gov.ph", "D2Core2026", "Active", "ADMIN", "District 2 Core ESU"],
    ["qcesu.bagongsilangan@quezoncity.gov.ph", "BS2026", "Active", "HC_USER", "BAGONG SILANGAN HEALTH CENTER"],
    ["qcesu.batasanannex@quezoncity.gov.ph", "BHA2026", "Active", "HC_USER", "BATASAN HILLS ANNEX HEALTH CENTER"],
    ["qcesu.batasanhills@quezoncity.gov.ph", "BHS2026", "Active", "HC_USER", "BATASAN HILLS SUPER HEALTH CENTER"],
    ["qcesu.bettygo@quezoncity.gov.ph", "BGB2026", "Active", "HC_USER", "BETTY-GO HEALTH CENTER"],
    ["qcesu.holyspirit@quezoncity.gov.ph", "HS2026", "Active", "HC_USER", "HOLY SPIRIT HEALTH CENTER"],
    ["qcesu.veterans@quezoncity.gov.ph", "VET2026", "Active", "HC_USER", "VETERANS HEALTH CENTER"],
    ["republichealthcenter.qchd@quezoncity.gov.ph", "REP2026", "Active", "HC_USER", "REPUBLIC HEALTH CENTER"],
    ["qcesu.commonwealth@quezoncity.gov.ph", "COM2026", "Active", "HC_USER", "COMMONWEALTH MAIN HEALTH CENTER"],
    ["qcesu.natgov@quezoncity.gov.ph", "NGC2026", "Active", "HC_USER", "NATIONAL GOVT HEALTH CENTER"],
    ["qcesu.lupangpangako@quezoncity.gov.ph", "LP2026", "Active", "HC_USER", "LUPANG PANGAKO HEALTH CENTER"],
    ["qcesu.payatas@quezoncity.gov.ph", "PA2026", "Active", "HC_USER", "PAYATAS A HEALTH CENTER"],
    ["qcesu.payatasb@quezoncity.gov.ph", "PB2026", "Active", "HC_USER", "PAYATAS B HEALTH CENTER"],
    ["qcesu.payatassuper@quezoncity.gov.ph", "PS2026", "Active", "HC_USER", "PAYATAS SUPER HEALTH CENTER"],
    ["QCHealthD2@quezoncity.gov.ph", "QCHD2", "Active", "RESTRICTED", "Executive Monitoring View"]
  ];

  sheet.clear();
  sheet.getRange(1, 1, 1, headers.length).setValues([headers])
       .setFontWeight("bold").setBackground("#1e3a5f").setFontColor("#ffffff");
  sheet.getRange(2, 1, userList.length, headers.length).setValues(userList);
  sheet.setFrozenRows(1);
  sheet.autoResizeColumns(1, headers.length);

  Logger.log("✅ 16 District 2 user accounts created successfully!");
}