// ==========================================
// ENVIRONMENT CONFIGURATION
// ==========================================
const ACTIVE_ENV = 'PROD';

const ENV_CONFIG = {
  PROD: {
    SPREADSHEET_ID: '1IbxJY59urIChYaLrwURWrOEBBrLNDiOqDO-w6LwP1xI',
    PHOTO_FOLDER_ID: '1nMFek_9bTttYPVW_vlV1eOfawDz3RGy-'
  },
  DEV: {
    SPREADSHEET_ID: '1E1GPV36RLHn7p4gHmdHB2zedtZJF2_Zb3ncGXZZTy5Y',
    PHOTO_FOLDER_ID: '1nMFek_9bTttYPVW_vlV1eOfawDz3RGy-'
  }
};

const CONSTANTS = {
  SPREADSHEET_ID: ENV_CONFIG[ACTIVE_ENV].SPREADSHEET_ID,
  PHOTO_FOLDER_ID: ENV_CONFIG[ACTIVE_ENV].PHOTO_FOLDER_ID,
  SHEETS: {
    TRAINEE_INFO: 'TRAINEE INFO',
    SETTINGS: 'Settings'
  },
  INFO_COLUMN: 'Trainee Info Card Text',
  CACHE_EXPIRATION: 21600 // 6 Hours
};

// ==========================================
// CACHE SERVICE WITH DYNAMIC CHUNKING (PILLAR 2)
// ==========================================
const CacheHelper = {
  put: function(key, valueString) {
    const cache = CacheService.getScriptCache();
    const chunkSize = 90000; // Safe limit under 100KB
    let chunks = [];
    let i = 0;
    while (i < valueString.length) {
      chunks.push(valueString.substring(i, i + chunkSize));
      i += chunkSize;
    }
    cache.put(key + '_chunks', chunks.length.toString(), CONSTANTS.CACHE_EXPIRATION);
    let cacheObj = {};
    chunks.forEach((chunk, idx) => cacheObj[`${key}_${idx}`] = chunk);
    cache.putAll(cacheObj, CONSTANTS.CACHE_EXPIRATION);
  },
  get: function(key) {
    const cache = CacheService.getScriptCache();
    const chunkCountStr = cache.get(key + '_chunks');
    if (!chunkCountStr) return null;
    const chunkCount = parseInt(chunkCountStr, 10);
    let result = '';
    for (let i = 0; i < chunkCount; i++) {
      const chunk = cache.get(`${key}_${i}`);
      if (!chunk) return null;
      result += chunk;
    }
    return result;
  },
  remove: function(key) {
    const cache = CacheService.getScriptCache();
    const chunkCountStr = cache.get(key + '_chunks');
    if (chunkCountStr) {
      const chunkCount = parseInt(chunkCountStr, 10);
      let keys = [key + '_chunks'];
      for (let i = 0; i < chunkCount; i++) keys.push(`${key}_${i}`);
      cache.removeAll(keys);
    } else {
      cache.remove(key);
    }
  }
};

/**
 * Creates Custom Menus in the Google Sheet for Administrators
 */
function onOpen() {
  SpreadsheetApp.getUi().createMenu('Trainee Info Cards')
    .addItem('Generate Info Text (Active Row)', 'menuGenerateActiveRow')
    .addItem('Generate Info Text (All Trainees)', 'menuGenerateAll')
    .addToUi();
}

function menuGenerateActiveRow() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  if (sheet.getName() !== CONSTANTS.SHEETS.TRAINEE_INFO) {
    SpreadsheetApp.getUi().alert("Please run this from the TRAINEE INFO sheet.");
    return;
  }
  const rowIdx = sheet.getActiveRange().getRow();
  if (rowIdx < 2) {
    SpreadsheetApp.getUi().alert("Please select a valid trainee row.");
    return;
  }
  const data = sheet.getDataRange().getValues();
  const traineeName = String(data[rowIdx - 1][2]).trim();

  if (!traineeName) {
    SpreadsheetApp.getUi().alert("No short name found on this row (Column C).");
    return;
  }

  SpreadsheetApp.getActiveSpreadsheet().toast(`Generating for ${traineeName}...`, "AI Generation", 10);
  try {
    adminGenerateCardText(traineeName);
    SpreadsheetApp.getActiveSpreadsheet().toast(`Successfully generated info card text for ${traineeName}!`, "Success", 5);
  } catch(e) {
    SpreadsheetApp.getUi().alert(`Error: ${e.message}`);
  }
}

function menuGenerateAll() {
  const ui = SpreadsheetApp.getUi();
  const response = ui.alert('Generate All', 'This may take several minutes. Proceed?', ui.ButtonSet.YES_NO);
  if (response !== ui.Button.YES) return;

  const trainees = getInitialData().allTrainees;
  SpreadsheetApp.getActiveSpreadsheet().toast(`Starting generation for ${trainees.length} trainees...`, "AI Generation");
  let count = 0;
  for (let t of trainees) {
    try {
      adminGenerateCardText(t);
      count++;
    } catch(e) {
      Logger.log(`Failed for ${t}: ${e.message}`);
    }
  }
  SpreadsheetApp.getActiveSpreadsheet().toast(`Completed generation for ${count} trainees.`, "Success", 10);
}

function FORCE_AUTHORIZATION() {
  try {
    SpreadsheetApp.openById(CONSTANTS.SPREADSHEET_ID);
    DriveApp.getFolderById(CONSTANTS.PHOTO_FOLDER_ID);
    Logger.log(`SUCCESS: Permissions granted.`);
  } catch(e) {
    Logger.log(`ERROR: Check access to ID: ${CONSTANTS.SPREADSHEET_ID}`);
    throw e;
  }
}

function doGet() {
  return ContentService.createTextOutput(`Trainee Info Backend (${ACTIVE_ENV}) is running properly.`);
}

function doPost(e) {
  try {
    if (!e || !e.postData || !e.postData.contents) throw new Error("No data received.");
    const request = JSON.parse(e.postData.contents);
    const action = request.action;
    const payload = request.payload || {};
    let result;

    switch (action) {
      case 'getInitialData': result = getInitialData(); break;
      case 'login': result = loginUser(payload.password); break;
      case 'verifySettingsPassword': result = verifyPassword('Settings', payload.password); break;
      case 'saveAppSettings': result = saveAppSettings(payload.newMappings); break;
      case 'updatePasswords': result = updatePasswords(payload.passwords); break;
      case 'getTraineeCardData': result = getTraineeCardData(payload.traineeName, payload.profile); break;
      case 'adminGenerateCardText': result = adminGenerateCardText(payload.traineeName); break;
      case 'adminRebuildSystemCache': result = adminRebuildSystemCache(); break;
      default: throw new Error(`Unknown action requested: ${action}`);
    }

    return ContentService.createTextOutput(JSON.stringify({ success: true, data: result })).setMimeType(ContentService.MimeType.JSON);
  } catch (error) {
    Logger.log(`API Error: ${error.message}`);
    return ContentService.createTextOutput(JSON.stringify({ success: false, error: error.message })).setMimeType(ContentService.MimeType.JSON);
  }
}

function getInitialData() {
  const cacheKey = 'INITIAL_DATA';
  const cached = CacheHelper.get(cacheKey);
  if (cached) return JSON.parse(cached);

  const traineeData = getTraineeData_();
  const result = { appSettings: getAppSettings_(), allTrainees: getTraineeList_(traineeData), availableFields: getAvailableFields_(traineeData) };
  CacheHelper.put(cacheKey, JSON.stringify(result));
  return result;
}

function getTraineeData_() {
  const ss = SpreadsheetApp.openById(CONSTANTS.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONSTANTS.SHEETS.TRAINEE_INFO);
  if (!sheet) throw new Error(`Sheet not found.`);
  return sheet.getDataRange().getValues();
}

function getAppSettings_() {
  const ss = SpreadsheetApp.openById(CONSTANTS.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONSTANTS.SHEETS.SETTINGS);
  const data = sheet.getDataRange().getValues();
  const settings = { profileDefaultFieldMappings: {} };
  data.forEach(row => {
    const key = row[0], value = row[1];
    if (key.endsWith(' Fields')) settings.profileDefaultFieldMappings[key.replace(' Fields', '').trim()] = value.split(',').map(field => field.trim());
  });
  return settings;
}

function getTraineeList_(data) {
  const names = new Set();
  const PROJECT_COL_INDEX = 1, SHORT_NAME_COL_INDEX = 2;
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    if (row.length > SHORT_NAME_COL_INDEX) {
      const projectStatus = String(row[PROJECT_COL_INDEX]).trim(), name = String(row[SHORT_NAME_COL_INDEX]).trim();
      if (name && !/exited/i.test(projectStatus)) names.add(name);
    }
  }
  return [...names].sort();
}

function getAvailableFields_(data) {
  if (!data || data.length === 0) throw new Error(`No data`);
  const headers = data[0], SHORT_NAME_COL_INDEX = 2;
  const fieldDisplayLabels = { "Trainee’s Full Name": "Full Name", "Age": "Age", "Gender": "Gender", "Address": "Address", "Contact 1\nRelation (Name)": "Contact 1 Relation (Name)", "Contact 1\nNumber": "Contact 1 Number", "Spoken Language / Dialect": "Spoken Language / Dialect", "Current Medication （经期有没有服药物）": "Current Medication", "Past Medical Conditions (e.g. any Major Operation etc.) （前期病历表）": "Past Medical Conditions", "Dietary Restriction(s) (if any) （食物限制）": "Dietary Restrictions", "Functioning": "Functioning", "Verbal": "Verbal", "Mobility": "Mobility", "Travelling": "Travelling", "Engagement Tips and Fun Facts": "Engagement Tips and Fun Facts", "Current Employment / Weekday Activities": "Current Employment / Weekday Activities", "General Comments Issues and Goals": "General Comments Issues and Goals" };
  const fieldsForSelection =[];
  headers.forEach((originalSheetHeader, index) => {
    const cleanedHeader = String(originalSheetHeader).trim();
    if (!cleanedHeader || index === SHORT_NAME_COL_INDEX || cleanedHeader === "Trainee’s Full Name") return;
    fieldsForSelection.push({ value: originalSheetHeader, label: fieldDisplayLabels[originalSheetHeader] || cleanedHeader.replace(/\r?\n/g, ' ') });
  });
  return fieldsForSelection;
}

function saveAppSettings(newMappings) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error("System is currently busy. Please try again.");
  try {
    const ss = SpreadsheetApp.openById(CONSTANTS.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONSTANTS.SHEETS.SETTINGS);
    const data = sheet.getRange("A:B").getValues();
    for (let i = 0; i < data.length; i++) {
      const profileNameRow = String(data[i][0] || '');
      if (!profileNameRow) continue;
      
      const profileName = profileNameRow.replace(' Fields', '').trim();
      
      // Handle dynamic Adhoc naming fallback
      let updateData = newMappings[profileName];
      if (!updateData && profileName.toLowerCase().includes('adhoc')) {
         const adhocKey = Object.keys(newMappings).find(k => k.toLowerCase().includes('adhoc'));
         if (adhocKey) updateData = newMappings[adhocKey];
      }
      
      if (updateData) {
        sheet.getRange(i + 1, 2).setValue(updateData.join(', '));
      }
    }
    CacheHelper.remove('INITIAL_DATA');
    return "Settings saved successfully!";
  } finally {
    lock.releaseLock();
  }
}

function loginUser(password) {
  const props = PropertiesService.getScriptProperties();
  if (password === props.getProperty('Regular-Volunteer')) return 'Regular Volunteer';
  if (password === props.getProperty('Adhoc-Volunteer')) return 'Adhoc Volunteer';
  if (password === props.getProperty('Settings')) return 'Settings';
  return false;
}

function verifyPassword(profile, password) {
  const expected = PropertiesService.getScriptProperties().getProperty({ 'Regular Volunteer': 'Regular-Volunteer', 'Adhoc Volunteer': 'Adhoc-Volunteer', 'Settings': 'Settings' }[profile]);
  return expected !== null && expected === password;
}

function updatePasswords(passwords) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error("System is busy.");
  try {
    const props = {}, map = { 'Regular Volunteer': 'Regular-Volunteer', 'Adhoc Volunteer': 'Adhoc-Volunteer', 'Settings': 'Settings' };
    let updated = false;
    for (const p in passwords) if (passwords[p] && map[p]) { props[map[p]] = passwords[p]; updated = true; }
    if (!updated) throw new Error("No passwords provided.");
    PropertiesService.getScriptProperties().setProperties(props, false);
    return "Passwords updated successfully!";
  } finally {
    lock.releaseLock();
  }
}

// Write-Through Cached Raw Data Fetcher
function getRawTraineeData_(traineeName) {
  const cacheKey = `RAW_TRAINEE_${traineeName.replace(/\s+/g, '')}`;
  const cached = CacheHelper.get(cacheKey);
  if (cached) return JSON.parse(cached);

  // Removed LockService here. Read operations must run concurrently to survive multi-user spikes.
  const ss = SpreadsheetApp.openById(CONSTANTS.SPREADSHEET_ID);
  const sheet = ss.getSheetByName(CONSTANTS.SHEETS.TRAINEE_INFO);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const rowIndex = data.findIndex(r => String(r[2]).trim() === traineeName);
  if (rowIndex === -1) throw new Error("Trainee not found in database.");
  const row = data[rowIndex];

  let infoData = {};
  let infoColIdx = headers.indexOf(CONSTANTS.INFO_COLUMN);
  if (infoColIdx !== -1 && row[infoColIdx]) {
    try { infoData = JSON.parse(row[infoColIdx]); } catch(e) {}
  }
  headers.forEach((h, i) => { if (infoData[h] === undefined) infoData[h] = row[i] || ''; });

  let photoBase64 = null, photoMime = null;
  try {
    const folder = DriveApp.getFolderById(CONSTANTS.PHOTO_FOLDER_ID);
    const files = folder.searchFiles(`title contains '${traineeName}'`);
    if (files.hasNext()) {
      const file = files.next();
      photoBase64 = Utilities.base64Encode(file.getBlob().getBytes());
      photoMime = file.getBlob().getContentType();
    }
  } catch (e) { Logger.log("Photo error: " + e.message); }

  const rawData = { infoData, photoBase64, photoMime };
  CacheHelper.put(cacheKey, JSON.stringify(rawData));
  return rawData;
}

function getTraineeCardData(traineeName, profile) {
  const rawData = getRawTraineeData_(traineeName);
  const infoData = rawData.infoData;
  
  let caregiverLines = [];[ {rel: 'Contact 1\nRelation (Name)', num: 'Contact 1\nNumber'},
    {rel: 'Contact 2\nRelation (Name)', num: 'Contact 2\nNumber'},
    {rel: 'Contact 3\nRelation / Name', num: 'Contact 3\nNumber'} ].forEach(c => {
      let r = infoData[c.rel], n = infoData[c.num];
      if (r || n) caregiverLines.push(`${r || ''}${(r && n) ? ' : ' : ''}${n || ''}`);
  });
  if (caregiverLines.length > 0) infoData['caregiverContactInfo'] = caregiverLines.join('\n');

  const settings = getInitialData().appSettings;
  
  // Robustly handle the Adhoc naming link
  const adhocKey = Object.keys(settings.profileDefaultFieldMappings).find(k => k.toLowerCase().includes('adhoc')) || 'Adhoc Volunteer';
  
  let allowedFields = settings.profileDefaultFieldMappings[profile];
  if (!allowedFields && profile.toLowerCase().includes('adhoc')) {
    allowedFields = settings.profileDefaultFieldMappings[adhocKey];
  }
  allowedFields = allowedFields || [];

  const labels = { "Age": "Age", "Gender": "Gender", "Address": "Address", "Spoken Language / Dialect": "Spoken Language / Dialect", "Current Medication （经期有没有服药物）": "Current Medication", "Past Medical Conditions (e.g. any Major Operation etc.) （前期病历表）": "Past Medical Conditions", "Dietary Restriction(s) (if any) （食物限制）": "Dietary Restrictions", "Functioning": "Functioning", "Verbal": "Verbal", "Mobility": "Mobility", "Travelling": "Travelling", "Engagement Tips and Fun Facts": "Engagement Tips and Fun Facts", "Current Employment / Weekday Activities": "Current Employment / Weekday Activities", "General Comments Issues and Goals": "General Comments Issues and Goals", "caregiverContactInfo": "Caregiver Contact Info" };

  const CATEGORY_MAP =[
    { title: "Basic Information", icon: "ph-user-circle", color: "text-blue-500", bgClass: "section-basic", keys:["Age", "Gender", "Address", "Spoken Language / Dialect", "caregiverContactInfo"] },
    { title: "Medical & Dietary", icon: "ph-heartbeat", color: "text-red-500", bgClass: "section-medical", keys:["Current Medication （经期有没有服药物）", "Past Medical Conditions (e.g. any Major Operation etc.) （前期病历表）", "Dietary Restriction(s) (if any) （食物限制）"] },
    { title: "Abilities & Support", icon: "ph-wheelchair", color: "text-purple-500", bgClass: "section-abilities", keys:["Functioning", "Verbal", "Mobility", "Travelling"] },
    { title: "Employment & Activities", icon: "ph-briefcase", color: "text-green-500", bgClass: "section-employment", keys:["Current Employment / Weekday Activities"] },
    { title: "Engagement & Notes", icon: "ph-star", color: "text-yellow-500", bgClass: "section-engagement", keys:["Engagement Tips and Fun Facts", "General Comments Issues and Goals"] }
  ];

  let processedKeys = new Set();
  let categories = [];
  
  CATEGORY_MAP.forEach(cat => {
    let fields = [];
    cat.keys.forEach(k => {
      processedKeys.add(k);
      let val = String(infoData[k] || '').trim();
      let label = labels[k] || String(k).trim().replace(/\r?\n/g, ' ');
      let isAllowed = k === 'caregiverContactInfo' ? allowedFields.some(a => a.includes("Contact")) : allowedFields.includes(label);
      if (isAllowed && val && val !== 'N/A') fields.push({ label: label, value: val });
    });
    if (fields.length > 0) categories.push({ ...cat, fields });
  });

  ['Contact 1\nRelation (Name)', 'Contact 1\nNumber', 'Contact 2\nRelation (Name)', 'Contact 2\nNumber', 'Contact 3\nRelation / Name', 'Contact 3\nNumber'].forEach(k => processedKeys.add(k));

  let additionalFields = [];
  Object.keys(infoData).forEach(h => {
    let cleanH = String(h).trim();
    if (!processedKeys.has(h) && cleanH && cleanH !== CONSTANTS.INFO_COLUMN && cleanH !== "Trainee’s Full Name") {
      let label = labels[h] || cleanH.replace(/\r?\n/g, ' ');
      if (allowedFields.includes(label)) {
        let val = String(infoData[h] || '').trim();
        if (val && val !== 'N/A') additionalFields.push({ label: label, value: val });
      }
    }
  });

  if (additionalFields.length > 0) {
    categories.push({
      title: "Additional Information",
      icon: "ph-list-plus",
      color: "text-gray-600",
      bgClass: "section-additional",
      fields: additionalFields
    });
  }

  return {
    fullName: infoData["Trainee’s Full Name"] || traineeName,
    photoBase64: rawData.photoBase64,
    photoMime: rawData.photoMime,
    categories: categories
  };
}

function adminGenerateCardText(traineeName) {
  const lock = LockService.getScriptLock();
  // Max tryLock is legally 30000ms. Anything higher throws a fatal error in Google Apps Script.
  if (!lock.tryLock(30000)) throw new Error("System is busy handling another generation request.");
  try {
    const ss = SpreadsheetApp.openById(CONSTANTS.SPREADSHEET_ID);
    const sheet = ss.getSheetByName(CONSTANTS.SHEETS.TRAINEE_INFO);
    const data = sheet.getDataRange().getValues();
    const headers = data[0];

    let targetColIndex = headers.indexOf(CONSTANTS.INFO_COLUMN);
    if (targetColIndex === -1) {
      targetColIndex = headers.length;
      sheet.getRange(1, targetColIndex + 1).setValue(CONSTANTS.INFO_COLUMN);
    }

    const rowIndex = data.findIndex((r, idx) => idx > 0 && String(r[2]).trim() === traineeName);
    if (rowIndex === -1) throw new Error(`Trainee "${traineeName}" not found.`);

    const row = data[rowIndex];
    let infoJSON = {};
    const API_KEY = PropertiesService.getScriptProperties().getProperty('Gemini_API_Key');

    headers.forEach((h, i) => {
      let headerName = String(h).trim();
      let val = row[i] ? String(row[i]).trim() : '';

      if (headerName === "Engagement Tips and Fun Facts" && val && val !== 'N/A' && API_KEY) {
        let prompt = `Reformat the following text by organizing it into logical sections with clear subheaders. Ensure the content under each subheader is concise. Do not add any introductory or concluding statements.\n\nText to reformat:\n${val}`;
        infoJSON[headerName] = callGeminiAPI_(API_KEY, prompt) || val;
      } else if (headerName === "General Comments Issues and Goals" && val && val !== 'N/A' && API_KEY) {
        let prompt = `Analyze the following comments and organize them into distinct categories: "Issues", "Goals", and "General Comments". Summarize the key points for each category using bullet points. Ensure the output is concise and does not include any introductory or concluding remarks.\n\nText to analyze:\n${val}`;
        infoJSON[headerName] = callGeminiAPI_(API_KEY, prompt) || val;
      } else {
        infoJSON[headerName] = val;
      }
    });

    sheet.getRange(rowIndex + 1, targetColIndex + 1).setValue(JSON.stringify(infoJSON));
    
    CacheHelper.remove(`RAW_TRAINEE_${traineeName.replace(/\s+/g, '')}`);
    getRawTraineeData_(traineeName);
    
    return true;
  } finally {
    lock.releaseLock();
  }
}

// Background CRON pre-computation task
function adminRebuildSystemCache() {
  const lock = LockService.getScriptLock();
  // Max tryLock is legally 30000ms.
  if (!lock.tryLock(30000)) throw new Error("System is already rebuilding cache.");
  try {
    CacheHelper.remove('INITIAL_DATA');
    const initData = getInitialData();
    let rebuiltCount = 0;
    
    for (const t of initData.allTrainees) {
      CacheHelper.remove(`RAW_TRAINEE_${t.replace(/\s+/g, '')}`);
      getRawTraineeData_(t);
      rebuiltCount++;
    }
    return `Successfully pre-computed memory cache for ${rebuiltCount} trainees!`;
  } finally {
    lock.releaseLock();
  }
}

function callGeminiAPI_(apiKey, prompt) {
  try {
    const response = UrlFetchApp.fetch(`https://generativelanguage.googleapis.com/v1beta/models/gemini-flash-latest:generateContent?key=${apiKey}`, { method: "post", contentType: "application/json", payload: JSON.stringify({ contents:[{ role: "user", parts: [{ text: prompt }] }] }), muteHttpExceptions: true });
    const json = JSON.parse(response.getContentText());
    if (json.candidates) return json.candidates[0].content.parts[0].text;
  } catch (e) { Logger.log("Gemini Err: " + e.message); }
  return null;
}