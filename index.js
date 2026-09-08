/**
 * =========================================================================
 * QC-DESU DISTRICT 2 MASTER SURVEILLANCE & AUTHENTICATION BACKEND
 * =========================================================================
 */

// PASTE YOUR ACTUAL SPREADSHEET ID HERE
const SPREADSHEET_ID = "1sUTII813loiUG-LD1j-JUgSHagmmKf82QhxU4-Kl5IA";

const DISTRICT_2_BARANGAYS = [
  "BAGONG SILANGAN",
  "BATASAN HILLS",
  "COMMONWEALTH",
  "HOLY SPIRIT",
  "PAYATAS"
];

/**
 * Handles HTTP GET requests (Safe for standalone web app and external API consumption)
 */
function doGet(e) {
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

  // Default: Serve standalone Web App HTML portal
  return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('QC-DESU District 2 Master Surveillance Portal')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1.0, maximum-scale=5.0');
}

/**
 * Handles HTTP POST requests (For saving records, logins, or uploads)
 */
function doPost(e) {
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

  if (action === 'uploadDirectoryPhoto') {
    return createJsonResponse(uploadDirectoryPhoto(payload));
  }

  if (action === 'uploadPdsFile') {
    return createJsonResponse(uploadPdsFile(payload));
  }

  return createJsonResponse({ success: false, message: "Invalid POST action." });
}

function createJsonResponse(data) {
  return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
}

// -------------------------------------------------------------------------
// SPREADSHEET & SHEET LOCATORS (CASE-INSENSITIVE)
// -------------------------------------------------------------------------

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

function getSheetByNameInsensitive(ss, name) {
  if (!ss || !name) return null;
  const sheets = ss.getSheets();
  const cleanTarget = name.trim().toUpperCase();
  for (let i = 0; i < sheets.length; i++) {
    if (sheets[i].getName().trim().toUpperCase() === cleanTarget) {
      return sheets[i];
    }
  }
  return null;
}

function getMdbSheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "MDB DISTRICT 2 2026") || 
         getSheetByNameInsensitive(activeSs, "MDB") || 
         activeSs.getSheets()[0];
}

function getEpisenseSheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "EPISENSE") || activeSs.getSheets()[0];
}

function getUsersSheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "users") || 
         getSheetByNameInsensitive(activeSs, "USERS") || 
         getSheetByNameInsensitive(activeSs, "ACCOUNTS");
}

function getDirectorySheet(ss) {
  const activeSs = ss || getSpreadsheet();
  if (!activeSs) return null;
  return getSheetByNameInsensitive(activeSs, "DIRECTORY") || 
         getSheetByNameInsensitive(activeSs, "STAFF") || 
         getSheetByNameInsensitive(activeSs, "DIRECTORY SHEET");
}

// -------------------------------------------------------------------------
// DRIVE STORAGE HELPERS (FOR DIRECTORY PHOTOS & PDS/CIF ATTACHMENTS)
// -------------------------------------------------------------------------

function getOrCreateDriveFolder(folderName) {
  const folders = DriveApp.getFoldersByName(folderName);
  if (folders.hasNext()) {
    return folders.next();
  }
  const newFolder = DriveApp.createFolder(folderName);
  newFolder.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);
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

// -------------------------------------------------------------------------
// 1. DIRECTORY SERVICE & PHOTO UPLOADS
// -------------------------------------------------------------------------

function getDirectoryData() {
  try {
    const ss = getSpreadsheet();
    const sheet = getDirectorySheet(ss);
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

/**
 * Uploads a staff avatar to Google Drive and writes the direct public URL into DIRECTORY sheet
 */
function uploadDirectoryPhoto(payload) {
  try {
    if (!payload || !payload.base64Data) {
      return { success: false, message: "No image file payload provided." };
    }

    const ss = getSpreadsheet();
    const sheet = getDirectorySheet(ss);
    if (!sheet) return { success: false, message: "Directory sheet not found." };

    const rawData = payload.base64Data;
    const staffName = payload.staffName || "Staff";
    const rowIndex = parseInt(payload.rowIndex, 10);

    const matches = rawData.match(/^data:([A-Za-z-+\/]+);base64,(.+)$/);
    if (!matches || matches.length !== 3) {
      return { success: false, message: "Invalid image base64 format." };
    }

    const contentType = matches[1];
    const base64Content = matches[2];
    const decodedBytes = Utilities.base64Decode(base64Content);
    const extension = contentType.split('/')[1] || 'jpg';
    const fileName = `Avatar_${staffName.replace(/[^a-zA-Z0-9]/g, '_')}_${new Date().getTime()}.${extension}`;

    const blob = Utilities.newBlob(decodedBytes, contentType, fileName);
    const folder = getOrCreateDriveFolder(DIRECTORY_PHOTO_FOLDER);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    const directUrl = `https://lh3.googleusercontent.com/d/${file.getId()}`;

    // Find or create PICTURE column in DIRECTORY
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());
    let picCol = headers.findIndex(h => h.includes("PICTURE") || h.includes("PHOTO") || h.includes("IMAGE") || h.includes("AVATAR"));

    if (picCol === -1) {
      picCol = headers.length;
      sheet.getRange(1, picCol + 1).setValue("PICTURE");
    }

    if (rowIndex && rowIndex > 1 && rowIndex <= sheet.getLastRow()) {
      sheet.getRange(rowIndex, picCol + 1).setValue(directUrl);
    } else {
      // Fallback find row by name
      const allRows = sheet.getDataRange().getValues();
      for (let r = 1; r < allRows.length; r++) {
        if (String(allRows[r][0] || '').trim().toUpperCase() === staffName.toUpperCase()) {
          sheet.getRange(r + 1, picCol + 1).setValue(directUrl);
          break;
        }
      }
    }

    return { success: true, url: directUrl, message: "Staff photo updated successfully!" };
  } catch (err) {
    Logger.log("uploadDirectoryPhoto error: " + err.toString());
    return { success: false, message: "Photo upload failed: " + err.message };
  }
}

/**
 * Uploads patient PDS / CIF / CRF documentation to Google Drive
 */
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
      blob = Utilities.newBlob(decodedBytes, contentType, `CIF_${caseId}_${originalName}`);
    } else {
      blob = Utilities.newBlob(rawData, "application/octet-stream", `CIF_${caseId}_${originalName}`);
    }

    const folder = getOrCreateDriveFolder(PDS_CIF_FOLDER);
    const file = folder.createFile(blob);
    file.setSharing(DriveApp.Access.ANYONE_WITH_LINK, DriveApp.Permission.VIEW);

    return { success: true, url: file.getUrl(), message: "CIF file uploaded successfully!" };
  } catch (err) {
    return { success: false, message: "CIF upload failed: " + err.message };
  }
}

// -------------------------------------------------------------------------
// 2. AUTHENTICATION SERVICE
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

    const sheet = getUsersSheet(ss);
    if (!sheet) {
      return { success: false, message: 'DATABASE_OFFLINE: "users" tab not found in the spreadsheet.' };
    }

    const rows = sheet.getDataRange().getValues();
    if (rows.length <= 1) {
      return { success: false, message: 'DATABASE_EMPTY: No user accounts registered.' };
    }

    for (let i = 1; i < rows.length; i++) {
      const dbUser = String(rows[i][0] || '').trim().toLowerCase();
      const dbPass = String(rows[i][1] || '').trim();
      const dbStatus = String(rows[i][2] || '').trim().toLowerCase();
      
      // Reads Column D (User Level) and Column E (Designation)
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
// 3. DASHBOARD METRICS WITH DYNAMIC BASELINE ARRAYS
// -------------------------------------------------------------------------

function getMetricsData(filters) {
  try {
    const ss = getSpreadsheet();
    const sheet = getMdbSheet(ss); 
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
      if (!row || row.join("").trim() === "") continue;

      const healthCenter = idxCenter !== -1 ? String(row[idxCenter]).trim() : "";
      const investigationRemarks = idxRemarks !== -1 ? String(row[idxRemarks]).trim() : "";
      const barangay = idxBarangay !== -1 ? String(row[idxBarangay]).trim() : "";
      const outcome = idxOutcome !== -1 ? String(row[idxOutcome]).trim() : "";
      const diseaseName = idxDisease !== -1 ? String(row[idxDisease]).trim().toUpperCase() : "";
      const rawClassValue = idxClass !== -1 ? String(row[idxClass]).trim().toUpperCase() : "";
      const rawRemarksValue = investigationRemarks.toUpperCase();

      if (!barangay && !diseaseName && !healthCenter) continue;
      if (!healthCenter || healthCenter === "") continue;
      
      // Filter out non-surveillance / discarded logs
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
    
    // Baseline threshold dictionary for all monitored disease vectors
    const baselines = {
      DENGUE: fetchComprehensiveDiseaseBaseline(ss, "DENGUE", "DENGUE BASELINE"),
      MEASLES: fetchComprehensiveDiseaseBaseline(ss, "MEASLES", "MEASLES BASELINE"),
      LEPTO: fetchComprehensiveDiseaseBaseline(ss, "LEPTO", "LEPTO BASELINE"),
      LEPTOSPIROSIS: fetchComprehensiveDiseaseBaseline(ss, "LEPTO", "LEPTO BASELINE"),
      COVID19: fetchComprehensiveDiseaseBaseline(ss, "COVID", "COVID BASELINE"),
      COVID: fetchComprehensiveDiseaseBaseline(ss, "COVID", "COVID BASELINE"),
      "COVID-19": fetchComprehensiveDiseaseBaseline(ss, "COVID", "COVID BASELINE"),
      CHIKUNGUNYA: fetchComprehensiveDiseaseBaseline(ss, "CHIKUNGUNYA", "CHIKUNGUNYA BASELINE"),
      TYPHOID: fetchComprehensiveDiseaseBaseline(ss, "TYPHOID", "TYPHOID BASELINE"),
      RABIES: fetchComprehensiveDiseaseBaseline(ss, "RABIES", "RABIES BASELINE"),
      AFP: fetchComprehensiveDiseaseBaseline(ss, "AFP", "AFP BASELINE"),
      DIPH: fetchComprehensiveDiseaseBaseline(ss, "DIPH", "DIPH BASELINE"),
      PERTUSSIS: fetchComprehensiveDiseaseBaseline(ss, "PERTUSSIS", "PERTUSSIS BASELINE"),
      ROTAVIRUS: fetchComprehensiveDiseaseBaseline(ss, "ROTAVIRUS", "ROTAVIRUS BASELINE"),
      CHOLERA: fetchComprehensiveDiseaseBaseline(ss, "CHOLERA", "CHOLERA BASELINE"),
      ILI: fetchComprehensiveDiseaseBaseline(ss, "ILI", "ILI BASELINE"),
      SARI: fetchComprehensiveDiseaseBaseline(ss, "SARI", "SARI BASELINE"),
      HEPA: fetchComprehensiveDiseaseBaseline(ss, "HEPA", "HEPA BASELINE"),
      AMES: fetchComprehensiveDiseaseBaseline(ss, "AMES", "AMES BASELINE"),
      MENINGO: fetchComprehensiveDiseaseBaseline(ss, "MENINGO", "MENINGO BASELINE"),
      HFMD: fetchComprehensiveDiseaseBaseline(ss, "HFMD", "HFMD BASELINE")
    };
    
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
    const sheet = getEpisenseSheet(ss);
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
      if (!row || row.join("").trim() === "") continue;

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

      const yyyy = dateObj.getFullYear();
      const mm = String(dateObj.getMonth() + 1).padStart(2, '0');
      const dd = String(dateObj.getDate()).padStart(2, '0');

      cases.push({
        caseId: caseId,
        disease: disease,
        healthCenter: healthCenter,
        dateValidated: `${yyyy}-${mm}-${dd}`,
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
    const sheet = getMdbSheet(ss); 
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
    const headers = values[0].map(h => String(h).trim().toUpperCase());
    
    let idxHC = headers.indexOf("HEALTH CENTER DESIGNATION");
    if (idxHC === -1) idxHC = headers.findIndex(h => h.includes("HEALTH CENTER") || h.includes("CENTER"));

    let idxBrgy = headers.indexOf("BARANGAY");
    if (idxBrgy === -1) idxBrgy = headers.findIndex(h => h.includes("BRGY") || h.includes("BGY"));

    let idxRemarks = headers.indexOf("INVESTIGATION REMARKS");
    if (idxRemarks === -1) idxRemarks = headers.findIndex(h => h.includes("REMARKS"));

    let idxDisease = headers.indexOf("DISEASE_NAME");
    if (idxDisease === -1) idxDisease = headers.findIndex(h => h.includes("DISEASE"));

    let idxClass = headers.indexOf("CASE_CLASSIFICATION");
    if (idxClass === -1) idxClass = headers.findIndex(h => h.includes("CLASSIFICATION") || h.includes("LABORATORY") || h.includes("STATUS"));

    let idxTesting = headers.indexOf("TYPE OF TESTING");
    if (idxTesting === -1) idxTesting = headers.findIndex(h => h.includes("TESTING") || h.includes("TEST"));

    let idxClinical = headers.indexOf("CLINICAL_CLASSIFICATION");
    if (idxClinical === -1) idxClinical = headers.findIndex(h => h.includes("CLINICAL"));

    let idxOutcome = headers.indexOf("OUTCOME");
    if (idxOutcome === -1) idxOutcome = headers.findIndex(h => h.includes("OUTCOME"));

    let idxCategory = headers.indexOf("CATEGORY");
    if (idxCategory === -1) idxCategory = headers.findIndex(h => h.includes("CATEGORY"));

    let idxDate = headers.indexOf("CESU DATE ADDED");
    if (idxDate === -1) idxDate = headers.findIndex(h => h.includes("CESU") || h.includes("DATE") || h.includes("ONSET"));

    const targetBrgy = filters && filters.barangay ? filters.barangay.toUpperCase().trim() : "ALL BARANGAY";
    const targetHC = filters && filters.healthCenter ? filters.healthCenter.toUpperCase().trim() : "ALL HEALTH CENTERS";
    const targetRemarks = filters && filters.remarks ? filters.remarks.toUpperCase().trim() : "ALL REMARKS";
    const targetDisease = filters && filters.disease ? filters.disease.toUpperCase().trim() : "DENGUE";

    const startDateStr = filters && filters.startDate ? filters.startDate.trim() : "";
    const endDateStr = filters && filters.endDate ? filters.endDate.trim() : "";

    let startMs = startDateStr ? new Date(startDateStr).getTime() : 0;
    let endMs = endDateStr ? new Date(endDateStr + "T23:59:59").getTime() : Infinity;

    const caseFindingMap = {};
    const testingMap = {};
    let confirmatoryCount = 0;
    const clinicalOutcomeMap = {};
    const cesdMap = {};

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;
      
      const hc = idxHC !== -1 && row[idxHC] ? String(row[idxHC]).trim() : "Unassigned Health Center";
      const brgy = idxBrgy !== -1 ? String(row[idxBrgy]).trim().toUpperCase() : "";
      const remarks = idxRemarks !== -1 ? String(row[idxRemarks]).trim().toUpperCase() : "";
      const disease = idxDisease !== -1 ? String(row[idxDisease]).trim().toUpperCase() : "";
      const rawClass = idxClass !== -1 ? String(row[idxClass]).trim().toUpperCase() : "";

      if (remarks.includes("DELIST") || remarks.includes("NON-RES") || remarks.includes("DISCARD") || remarks.includes("DUPLICATE")) continue;

      if (targetBrgy !== "ALL BARANGAY" && !brgy.includes(targetBrgy)) continue;
      if (targetHC !== "ALL HEALTH CENTERS" && hc.toUpperCase() !== targetHC) continue;
      if (targetRemarks !== "ALL REMARKS" && !remarks.includes(targetRemarks)) continue;
      if (targetDisease !== "ALL DISEASES" && !disease.includes(targetDisease)) continue;

      if (startDateStr || endDateStr) {
        let rawD = idxDate !== -1 ? row[idxDate] : null;
        let dObj = rawD instanceof Date ? rawD : new Date(rawD);
        let rowTime = dObj.getTime();
        if (isNaN(rowTime) || rowTime < startMs || rowTime > endMs) continue;
      }

      let classification = "null";
      if (rawClass.includes("SUSPECT")) classification = "SUSPECT";
      else if (rawClass.includes("CONFIRM") || rawClass.includes("POS")) classification = "CONFIRM";
      else if (rawClass.includes("PROBABLE")) classification = "PROBABLE";

      if (!caseFindingMap[hc]) caseFindingMap[hc] = { SUSPECT: 0, CONFIRM: 0, PROBABLE: 0, null: 0 };
      caseFindingMap[hc][classification] = (caseFindingMap[hc][classification] || 0) + 1;

      let testType = idxTesting !== -1 && row[idxTesting] ? String(row[idxTesting]).trim() : "null";
      if (!testType || testType === "") testType = "null";
      if (!testingMap[hc]) testingMap[hc] = {};
      testingMap[hc][testType] = (testingMap[hc][testType] || 0) + 1;

      if (testType.toUpperCase().includes("PCR") || testType.toUpperCase().includes("RTPCR") || testType.toUpperCase().includes("POLYMERASE") || testType.toUpperCase().includes("SERUM") || testType.toUpperCase().includes("SWAB")) {
        confirmatoryCount++;
      }

      let clinical = idxClinical !== -1 && row[idxClinical] ? String(row[idxClinical]).trim().toUpperCase() : "null";
      let outcome = idxOutcome !== -1 && row[idxOutcome] ? String(row[idxOutcome]).trim().toUpperCase() : "null";
      if (outcome.includes("ALIVE")) outcome = "ALIVE";
      else if (outcome.includes("DIE") || outcome.includes("DEATH")) outcome = "DIED";
      else outcome = "null";

      const clinKey = `${hc}||${clinical}`;
      if (!clinicalOutcomeMap[clinKey]) clinicalOutcomeMap[clinKey] = { hc: hc, clinical: clinical, ALIVE: 0, DIED: 0, null: 0 };
      clinicalOutcomeMap[clinKey][outcome]++;

      let category = idxCategory !== -1 && row[idxCategory] ? String(row[idxCategory]).trim().toUpperCase() : "HC DETECTED";
      if (category.includes("PHSU")) category = "PHSU ENDORSEMENT";
      else category = "HC DETECTED";

      if (!cesdMap[hc]) cesdMap[hc] = { "PHSU ENDORSEMENT": 0, "HC DETECTED": 0 };
      cesdMap[hc][category]++;
    }

    return {
      table1: formatTable1(caseFindingMap),
      table21: formatTable21(testingMap),
      kpi22: confirmatoryCount,
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

// -------------------------------------------------------------------------
// 6. PATIENT PROFILE RECORDS & CRUD SERVICES
// -------------------------------------------------------------------------

function getRecords() {
  try {
    const ss = getSpreadsheet();
    const sheet = getMdbSheet(ss); 
    if (!sheet) return [];

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();
    if (lastRow <= 1 || lastCol === 0) return [];

    const values = sheet.getRange(1, 1, lastRow, lastCol).getValues();
    const headers = values[0].map(h => String(h).trim());
    const records = [];

    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.join("").trim() === "") continue;

      const recordObj = { _rowIndex: r + 1 };
      headers.forEach((h, idx) => {
        let val = row[idx];
        if (val instanceof Date) {
          val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
        }
        recordObj[h] = (val !== undefined && val !== null) ? String(val).trim() : "";
      });
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
    const sheet = getMdbSheet(ss);
    if (!sheet) return { success: false, message: "MDB Sheet not found." };

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0].map(h => String(h).trim().toUpperCase());

    let targetRow = rowIndex;
    if (!targetRow || isNaN(targetRow) || targetRow <= 1) {
      targetRow = sheet.getLastRow() + 1;
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
    return { success: true, message: `Record successfully ${rowIndex ? 'updated' : 'saved'}!` };
  } catch (err) {
    return { success: false, message: `Failed to save: ${err.message}` };
  }
}

// -------------------------------------------------------------------------
// 7. DYNAMIC BASELINE ENGINE (MULTI-BARANGAY SIDE-BY-SIDE COLUMNS)
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

function fetchComprehensiveDiseaseBaseline(ss, diseaseKey, defaultSheetName) {
  let masterSheet = getSheetByNameInsensitive(ss, defaultSheetName);
  if (!masterSheet) {
    masterSheet = getSheetByNameInsensitive(ss, diseaseKey + " BASELINE") || getSheetByNameInsensitive(ss, diseaseKey);
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
        const bSheet = getSheetByNameInsensitive(ss, sName);
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
// 8. TABLE FORMATTERS & UTILITIES
// -------------------------------------------------------------------------

function formatTable1(map) {
  let list = [];
  let totals = { SUSPECT: 0, CONFIRM: 0, null: 0, PROBABLE: 0, grandTotal: 0 };
  for (let hc in map) {
    let row = map[hc];
    let sum = (row.SUSPECT || 0) + (row.CONFIRM || 0) + (row.null || 0) + (row.PROBABLE || 0);
    list.push({ hc: hc, suspect: row.SUSPECT || 0, confirm: row.CONFIRM || 0, nullVal: row.null || 0, probable: row.PROBABLE || 0, grandTotal: sum });
    totals.SUSPECT += (row.SUSPECT || 0);
    totals.CONFIRM += (row.CONFIRM || 0);
    totals.null += (row.null || 0);
    totals.PROBABLE += (row.PROBABLE || 0);
    totals.grandTotal += sum;
  }
  return { rows: list, totals: totals };
}

function formatTable21(map) {
  let testTypesSet = new Set();
  let list = [];
  for (let hc in map) {
    for (let tt in map[hc]) testTypesSet.add(tt);
  }
  let testTypes = Array.from(testTypesSet).sort();
  let totals = { grandTotal: 0 };
  testTypes.forEach(tt => totals[tt] = 0);

  for (let hc in map) {
    let rowObj = { hc: hc, grandTotal: 0 };
    testTypes.forEach(tt => {
      let val = map[hc][tt] || 0;
      rowObj[tt] = val;
      rowObj.grandTotal += val;
      totals[tt] += val;
      totals.grandTotal += val;
    });
    list.push(rowObj);
  }
  return { testTypes: testTypes, rows: list, totals: totals };
}

function formatTable34(map) {
  let list = [];
  let totals = { ALIVE: 0, nullVal: 0, DIED: 0, grandTotal: 0 };
  for (let key in map) {
    let item = map[key];
    let sum = item.ALIVE + item.null + item.DIED;
    list.push({ hc: item.hc, clinical: item.clinical, alive: item.ALIVE, nullVal: item.null, died: item.DIED, grandTotal: sum });
    totals.ALIVE += item.ALIVE;
    totals.nullVal += item.null;
    totals.DIED += item.DIED;
    totals.grandTotal += sum;
  }
  return { rows: list, totals: totals };
}

function formatTable4Star(map) {
  let list = [];
  let totals = { phsu: 0, hcDetected: 0, grandTotal: 0 };
  for (let hc in map) {
    let phsu = map[hc]["PHSU ENDORSEMENT"] || 0;
    let hcDet = map[hc]["HC DETECTED"] || 0;
    let sum = phsu + hcDet;
    list.push({ hc: hc, phsu: phsu, hcDetected: hcDet, grandTotal: sum });
    totals.phsu += phsu;
    totals.hcDetected += hcDet;
    totals.grandTotal += sum;
  }
  return { rows: list, totals: totals };
}

function getFormattedTime() {
  const now = new Date();
  return Utilities.formatDate(now, Session.getScriptTimeZone(), "hh:mm a");
}