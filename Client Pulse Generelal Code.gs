/**
 * ============================================================
 * GGE / JCC DUES TRACKER - Apps Script Backend (ENHANCED)
 * With per-advisor client preference toggles
 * ============================================================
 */

const SHEET_NAME = 'Dues Tracker';
const HEADERS = ['Policy Number','Client Name','Email','Product','Premium Mode','Premium Amount','Due Date','Policy Status','Last Reminder Sent','Send Dues?'];

const BIRTHDAY_SHEET_NAME = 'Birthday Tracker';
const BIRTHDAY_HEADERS = ['Full Name','Email','Contact Number','Location','Date of Birth','Last Greeting Sent (Year)','Send Birthday?'];

const CONFIG_DEFAULTS = {
  senderName: '',
  contactEmail: '',
  headerImageFileId: '',
  footerImageFileId: '',
  connectLink: '',
  payLink: ''
};
const CONFIG_KEYS = Object.keys(CONFIG_DEFAULTS);

/* ============================================================
   ADVISOR ACTIVITY / HARD-STOP CHECK
   ------------------------------------------------------------
   Every real upload (pushDuesRows or pushBirthdayRows) stamps
   LAST_UPLOAD_TIMESTAMP. If more than INACTIVITY_LIMIT_DAYS pass
   without a fresh upload, the advisor is considered inactive and
   the script hard-stops: doGet/doPost refuse every action except
   the ones needed to let them reactivate (uploading new data, or
   checking their own status), and the time-triggered senders
   (daily dues reminders, birthday greetings) skip themselves
   entirely so nothing sends on a stale account.
   ============================================================ */
const INACTIVITY_LIMIT_DAYS = 90;

// Actions allowed to run even while hard-stopped. Uploading is
// deliberately included — it's the only way to self-reactivate — and
// the status/branding reads are included so the app can render a clear
// "you're locked out, upload to continue" screen instead of a raw error.
const ACTIONS_EXEMPT_FROM_HARD_STOP = [
  'pushDues', 'pushBirthdays', 'getAdvisorActiveStatus', 'getConfig', 'getAdvisorProfile', 'getProfileImagePreview'
];

function recordUploadActivity(){
  PropertiesService.getScriptProperties().setProperty('LAST_UPLOAD_TIMESTAMP', new Date().toISOString());
}

// Returns the activity status. If no upload has ever been recorded,
// falls back to a stamped first-run install date (set on first call)
// so a brand-new deployment gets a full grace period rather than
// looking instantly stale.
function getAdvisorActiveStatus(){
  const props = PropertiesService.getScriptProperties();
  let lastUpload = props.getProperty('LAST_UPLOAD_TIMESTAMP');
  if (!lastUpload){
    let installedAt = props.getProperty('FIRST_INSTALL_TIMESTAMP');
    if (!installedAt){
      installedAt = new Date().toISOString();
      props.setProperty('FIRST_INSTALL_TIMESTAMP', installedAt);
    }
    lastUpload = installedAt;
  }
  const daysSince = (Date.now() - new Date(lastUpload).getTime()) / (24 * 60 * 60 * 1000);
  return {
    active: daysSince <= INACTIVITY_LIMIT_DAYS,
    lastUploadTimestamp: lastUpload,
    daysSinceLastUpload: Math.floor(daysSince),
    daysUntilLock: Math.max(0, Math.ceil(INACTIVITY_LIMIT_DAYS - daysSince))
  };
}

function isAdvisorActive(){
  return getAdvisorActiveStatus().active;
}

/* ============================================================
   AUTO-CLEAR CLIENT DATA AFTER 90 DAYS OF INACTIVITY
   ------------------------------------------------------------
   Runs once a day (installed alongside the other daily triggers —
   see createInactivityPurgeTrigger). If it's been exactly
   INACTIVITY_LIMIT_DAYS or more since the last real upload, every
   stored client record (Dues Tracker, Birthday Tracker) is wiped —
   but a CSV snapshot is saved to Drive first, and the advisor is
   emailed a notice, so this is never a silent, unrecoverable
   surprise. Guarded so it only fires once per inactivity stretch: a
   fresh upload resets LAST_UPLOAD_TIMESTAMP, which automatically
   re-arms it for the next 90-day window.
   ============================================================ */
function purgeInactiveClientData(){
  const status = getAdvisorActiveStatus();
  if (status.daysSinceLastUpload < INACTIVITY_LIMIT_DAYS) return; // not due yet

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_PURGE_BASELINE') === status.lastUploadTimestamp){
    return; // already purged for this inactivity stretch
  }

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const summary = { duesCleared: 0, birthdaysCleared: 0, backupFileUrl: null };

  try{
    summary.backupFileUrl = backupClientDataToDrive(ss);
  }catch(e){
    // Backup failing should never block the purge itself — 90 days of
    // inactivity means data minimization takes priority.
  }

  summary.duesCleared      = clearSheetDataRows(ss, SHEET_NAME, HEADERS.length);
  summary.birthdaysCleared = clearSheetDataRows(ss, BIRTHDAY_SHEET_NAME, BIRTHDAY_HEADERS.length);

  props.setProperty('LAST_PURGE_BASELINE', status.lastUploadTimestamp);
  props.setProperty('LAST_PURGE_TIMESTAMP', new Date().toISOString());

  notifyAdvisorOfPurge(summary);
  return summary;
}

// Clears every data row below the header in a sheet (leaves the header
// and the sheet itself intact). Returns how many rows were cleared.
function clearSheetDataRows(ss, sheetName, numCols){
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return 0;
  const lastRow = sheet.getLastRow();
  if (lastRow <= 1) return 0; // header only (or empty) — nothing to clear
  const clearedCount = lastRow - 1;
  sheet.getRange(2, 1, clearedCount, numCols).clearContent();
  return clearedCount;
}

// Snapshots the Dues Tracker and Birthday Tracker sheets to a single
// timestamped CSV in a "Client Pulse Purge Backups" Drive folder,
// right before they get wiped. Returns the backup file's URL, or null
// if there was nothing worth backing up.
function backupClientDataToDrive(ss){
  const folderName = 'Client Pulse Purge Backups';
  const existingFolders = DriveApp.getFoldersByName(folderName);
  const folder = existingFolders.hasNext() ? existingFolders.next() : DriveApp.createFolder(folderName);
  const tz = Session.getScriptTimeZone();
  const stamp = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd_HHmm');

  const duesSheet = ss.getSheetByName(SHEET_NAME);
  const bdaySheet = ss.getSheetByName(BIRTHDAY_SHEET_NAME);
  let csv = '';
  if (duesSheet && duesSheet.getLastRow() > 1){
    csv += 'DUES TRACKER\n' + sheetToCsv(duesSheet) + '\n\n';
  }
  if (bdaySheet && bdaySheet.getLastRow() > 1){
    csv += 'BIRTHDAY TRACKER\n' + sheetToCsv(bdaySheet) + '\n';
  }
  if (!csv) return null;
  const file = folder.createFile('client_pulse_backup_' + stamp + '.csv', csv, MimeType.CSV);
  return file.getUrl();
}

function sheetToCsv(sheet){
  const data = sheet.getDataRange().getValues();
  const tz = Session.getScriptTimeZone();
  return data.map(row => row.map(cell => {
    if (cell instanceof Date) return Utilities.formatDate(cell, tz, 'yyyy-MM-dd');
    const s = String(cell == null ? '' : cell);
    return '"' + s.replace(/"/g, '""') + '"';
  }).join(',')).join('\n');
}

// contactEmail (not getActiveUser/getEffectiveUser) is used deliberately
// here — this runs from a time-driven trigger, and some Workspace
// policies (e.g. Sun Life's) block the scope those need. contactEmail
// sidesteps that entirely.
function notifyAdvisorOfPurge(summary){
  try{
    const config = getBrandConfig();
    const recipient = config.contactEmail;
    if (!recipient) return; // nowhere configured to send it
    const body =
      'This is an automated notice from Client Pulse.\n\n' +
      'No client data has been uploaded in ' + INACTIVITY_LIMIT_DAYS + ' days, so all stored client information has been automatically cleared:\n\n' +
      '- Dues Tracker: ' + summary.duesCleared + ' row(s) cleared\n' +
      '- Birthday Tracker: ' + summary.birthdaysCleared + ' row(s) cleared\n\n' +
      (summary.backupFileUrl ? ('A backup was saved before deletion, in case you need it back:\n' + summary.backupFileUrl + '\n\n') : '') +
      'Upload a new client list anytime to reactivate the app and start fresh.';
    GmailApp.sendEmail(recipient, 'Client Pulse: client data auto-cleared (90 days inactive)', body);
  }catch(e){
    // A failed notification should never surface as a broken purge —
    // the data clearing above has already completed successfully.
  }
}

function getBrandConfig(){
  const props = PropertiesService.getScriptProperties();
  const config = {};
  CONFIG_KEYS.forEach(key => {
    config[key] = props.getProperty('CFG_' + key) || CONFIG_DEFAULTS[key];
  });
  return config;
}

function saveBrandConfig(partialConfig){
  const props = PropertiesService.getScriptProperties();
  CONFIG_KEYS.forEach(key => {
    if (partialConfig[key] !== undefined){
      props.setProperty('CFG_' + key, String(partialConfig[key]));
    }
  });
}

function assertConfigured(config){
  const required = ['senderName','headerImageFileId','footerImageFileId'];
  const missing = required.filter(key => !config[key]);
  if (missing.length > 0){
    throw new Error(
      'Branding not set up yet. Open the app, tap "Setup", fill in ' +
      '"Your branding" (missing: ' + missing.join(', ') + '), and tap ' +
      'SAVE BRANDING before reminders can be sent.'
    );
  }
}

function assertConfiguredForBirthday(config){
  const required = ['senderName','headerImageFileId','footerImageFileId'];
  const missing = required.filter(key => !config[key]);
  if (missing.length > 0){
    throw new Error(
      'Branding not set up yet. Open the app, tap "Setup", fill in ' +
      '"Your branding" (missing: ' + missing.join(', ') + '), and tap ' +
      'SAVE BRANDING before birthday greetings can be sent.'
    );
  }
}

function uploadBrandImage(target, base64, mimeType){
  if (target !== 'header' && target !== 'footer'){
    throw new Error('Invalid image target: ' + target);
  }
  const configKey = target === 'header' ? 'headerImageFileId' : 'footerImageFileId';
  const propKey = 'CFG_' + configKey;
  const props = PropertiesService.getScriptProperties();

  const oldFileId = props.getProperty(propKey);
  if (oldFileId){
    try{ DriveApp.getFileById(oldFileId).setTrashed(true); }catch(e){ /* already gone */ }
  }

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/png', target + '.png');
  const file = DriveApp.createFile(blob);

  props.setProperty(propKey, file.getId());
  return file.getId();
}

function getAdvisorProfile(){
  const props = PropertiesService.getScriptProperties();
  return {
    advisorName: props.getProperty('ADVISOR_NAME') || '',
    profileImageFileId: props.getProperty('ADVISOR_PROFILE_IMAGE_FILE_ID') || ''
  };
}

function saveAdvisorName(name){
  PropertiesService.getScriptProperties().setProperty('ADVISOR_NAME', name || '');
}

function uploadProfileImage(base64, mimeType){
  const props = PropertiesService.getScriptProperties();
  const propKey = 'ADVISOR_PROFILE_IMAGE_FILE_ID';

  const oldFileId = props.getProperty(propKey);
  if (oldFileId){
    try{ DriveApp.getFileById(oldFileId).setTrashed(true); }catch(e){ }
  }

  const bytes = Utilities.base64Decode(base64);
  const blob = Utilities.newBlob(bytes, mimeType || 'image/png', 'profile.png');
  const file = DriveApp.createFile(blob);

  props.setProperty(propKey, file.getId());
  return file.getId();
}

function getProfileImagePreviewData(){
  const fileId = getAdvisorProfile().profileImageFileId;
  if (!fileId) return { base64: null };
  try{
    const blob = DriveApp.getFileById(fileId).getBlob();
    return { base64: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType() };
  }catch(e){
    return { base64: null };
  }
}

function setupSheet(){
  createInactivityPurgeTrigger(); // self-installs once; cheap no-op after that
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SHEET_NAME);
  if (!sheet){
    sheet = ss.insertSheet(SHEET_NAME);
  }
  if (sheet.getLastRow() === 0){
    sheet.appendRow(HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() > 0){
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Send Dues?')){
      const lastCol = HEADERS.length;
      sheet.getRange(1, lastCol).setValue('Send Dues?');
      for (let i = 2; i <= sheet.getLastRow(); i++){
        sheet.getRange(i, lastCol).setValue(true);
      }
    }
  }
  const policyColIndex = HEADERS.indexOf('Policy Number') + 1;
  sheet.getRange(1, policyColIndex, sheet.getMaxRows(), 1).setNumberFormat('@');
}

function setupBirthdaySheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(BIRTHDAY_SHEET_NAME);
  if (!sheet){
    sheet = ss.insertSheet(BIRTHDAY_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0){
    sheet.appendRow(BIRTHDAY_HEADERS);
    sheet.setFrozenRows(1);
  } else if (sheet.getLastRow() > 0){
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headers.includes('Send Birthday?')){
      const lastCol = BIRTHDAY_HEADERS.length;
      sheet.getRange(1, lastCol).setValue('Send Birthday?');
      for (let i = 2; i <= sheet.getLastRow(); i++){
        sheet.getRange(i, lastCol).setValue(true);
      }
    }
  }
}

function getAutoSendStatus(){
  const val = PropertiesService.getScriptProperties().getProperty('AUTO_SEND_ENABLED');
  return { enabled: val === null ? true : val === '1' };
}

function setAutoSendStatus(enabled){
  PropertiesService.getScriptProperties().setProperty('AUTO_SEND_ENABLED', enabled ? '1' : '0');
}

function getBirthdayAutoSendStatus(){
  const val = PropertiesService.getScriptProperties().getProperty('BDAY_AUTO_SEND_ENABLED');
  return { enabled: val === null ? true : val === '1' };
}

function setBirthdayAutoSendStatus(enabled){
  PropertiesService.getScriptProperties().setProperty('BDAY_AUTO_SEND_ENABLED', enabled ? '1' : '0');
}

function getSendHour(){
  const val = PropertiesService.getScriptProperties().getProperty('SEND_HOUR');
  return { hour: val === null ? 6 : Number(val) };
}

function setSendHour(hour){
  hour = Number(hour);
  if (!(hour >= 6 && hour <= 16)){
    throw new Error('Send hour must be between 6 (6-7AM) and 16 (4-5PM).');
  }
  PropertiesService.getScriptProperties().setProperty('SEND_HOUR', String(hour));
  createDailyTrigger(hour);
  createBirthdayDailyTrigger(hour);
  createInactivityPurgeTrigger();
  return { hour: hour };
}

function createDailyTrigger(hour){
  hour = hour !== undefined ? hour : getSendHour().hour;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyReminders') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyReminders')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
}

function createBirthdayDailyTrigger(hour){
  hour = hour !== undefined ? hour : getSendHour().hour;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyBirthdayGreetings') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyBirthdayGreetings')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
}

// Runs independently of the advisor's chosen send hour — fixed at 3AM
// so it never competes with, or gets skipped alongside, the reminder/
// birthday sends if the advisor later disables those. Checks for an
// existing trigger first rather than delete-then-recreate every time,
// since this is also called from setupSheet() on every upload — that
// keeps it cheap and lets it self-install for advisors who were
// already using the app before this feature existed, without needing
// them to touch Send Hour settings.
function createInactivityPurgeTrigger(){
  const alreadyInstalled = ScriptApp.getProjectTriggers()
    .some(t => t.getHandlerFunction() === 'purgeInactiveClientData');
  if (alreadyInstalled) return;
  ScriptApp.newTrigger('purgeInactiveClientData')
    .timeBased()
    .everyDays(1)
    .atHour(3)
    .create();
}

/* ============================================================
   CLIENT PREFERENCE MANAGEMENT
   ============================================================ */

function getDuesClientList(){
  setupSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const policyNum = row[col('Policy Number')];
    if (!policyNum) continue;
    let parsedAmount = 0;
    const premiumAmount = row[col('Premium Amount')];
    if (premiumAmount){
      if (typeof premiumAmount === 'number'){
        parsedAmount = premiumAmount;
      } else {
        parsedAmount = parseFloat(String(premiumAmount).replace(/[^\d.-]/g, '').trim()) || 0;
      }
    }
    const dueDate = row[col('Due Date')];
    result.push({
      policyNumber: policyNum,
      clientName: row[col('Client Name')],
      email: row[col('Email')],
      product: row[col('Product')],
      premiumMode: row[col('Premium Mode')],
      premiumAmount: parsedAmount,
      dueDate: dueDate instanceof Date ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      policyStatus: row[col('Policy Status')],
      sendDues: row[col('Send Dues?')] === true || row[col('Send Dues?')] === 'TRUE' || row[col('Send Dues?')] === 1 || row[col('Send Dues?')] === '1'
    });
  }
  return result;
}

function getBirthdayClientList(){
  setupBirthdaySheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const fullName = row[col('Full Name')];
    const email = row[col('Email')];
    if (!fullName || !email) continue;
    const dob = row[col('Date of Birth')];
    result.push({
      fullName: fullName,
      email: email,
      contactNumber: row[col('Contact Number')],
      location: row[col('Location')],
      dateOfBirth: dob instanceof Date ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      sendBirthday: row[col('Send Birthday?')] === true || row[col('Send Birthday?')] === 'TRUE' || row[col('Send Birthday?')] === 1 || row[col('Send Birthday?')] === '1'
    });
  }
  return result;
}

function setDuesPreference(policyNumber, enabled){
  setupSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const policyCol = headers.indexOf('Policy Number');
  const sendCol = headers.indexOf('Send Dues?');
  for (let i = 1; i < data.length; i++){
    if (String(data[i][policyCol]) === String(policyNumber)){
      sheet.getRange(i + 1, sendCol + 1).setValue(enabled);
      return { success: true };
    }
  }
  return { success: false, error: 'Policy not found' };
}

function setBirthdayPreference(email, enabled){
  setupBirthdaySheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const emailCol = headers.indexOf('Email');
  const sendCol = headers.indexOf('Send Birthday?');
  for (let i = 1; i < data.length; i++){
    if (String(data[i][emailCol]).toLowerCase() === String(email).toLowerCase()){
      sheet.getRange(i + 1, sendCol + 1).setValue(enabled);
      return { success: true };
    }
  }
  return { success: false, error: 'Email not found' };
}

/* ============================================================
   WEB APP ENTRY POINTS
   ============================================================ */
function doGet(e){
  const action = e.parameter.action;
  if (action === 'getAdvisorActiveStatus')    return jsonResponse(getAdvisorActiveStatus());
  if (!ACTIONS_EXEMPT_FROM_HARD_STOP.includes(action) && !isAdvisorActive()){
    return jsonResponse(Object.assign({ error: 'ADVISOR_INACTIVE' }, getAdvisorActiveStatus()));
  }
  if (action === 'getDuesClientList')         return jsonResponse({ clients: getDuesClientList() });
  if (action === 'getBirthdayClientList')     return jsonResponse({ clients: getBirthdayClientList() });
  if (action === 'getDueToday')               return jsonResponse({ rows: getDueTodayRows() });
  if (action === 'getConfig')                 return jsonResponse({ config: getBrandConfig() });
  if (action === 'getImagePreview')           return jsonResponse(getImagePreviewData(e.parameter.target));
  if (action === 'getDailyStats')             return jsonResponse(getDailyStats());
  if (action === 'getAdvisorProfile')         return jsonResponse(getAdvisorProfile());
  if (action === 'getProfileImagePreview')    return jsonResponse(getProfileImagePreviewData());
  if (action === 'getAutoSendStatus')         return jsonResponse(getAutoSendStatus());
  if (action === 'getBirthdaysToday')         return jsonResponse({ rows: getBirthdaysTodayRows() });
  if (action === 'getBirthdayDailyStats')     return jsonResponse(getBirthdayDailyStats());
  if (action === 'getBirthdayAutoSendStatus') return jsonResponse(getBirthdayAutoSendStatus());
  if (action === 'getSendHour')               return jsonResponse(getSendHour());
  return jsonResponse({ error: 'Unknown action' });
}

function getImagePreviewData(target){
  const config = getBrandConfig();
  const fileId = target === 'header' ? config.headerImageFileId : config.footerImageFileId;
  if (!fileId) return { base64: null };
  try{
    const blob = DriveApp.getFileById(fileId).getBlob();
    return { base64: Utilities.base64Encode(blob.getBytes()), mimeType: blob.getContentType() };
  }catch(e){
    return { base64: null };
  }
}

function getDueTodayRows(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const todayFormatted = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const dueDate = row[col('Due Date')];
    const lastReminderSent = row[col('Last Reminder Sent')] || '';
    const lastSentStr = lastReminderSent ? String(lastReminderSent) : '';
    const wasSentToday = lastSentStr === todayStr;
    const dueDateStr = (dueDate instanceof Date) ? Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd') : '';
    const isDueToday = dueDateStr === todayStr;
    if (!isDueToday && !wasSentToday) continue;
    result.push({
      policyNumber: row[col('Policy Number')],
      clientName: row[col('Client Name')],
      product: row[col('Product')],
      premiumAmount: row[col('Premium Amount')],
      premiumMode: row[col('Premium Mode')],
      dueDateFormatted: isDueToday ? Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy') : todayFormatted,
      lastReminderSent: lastSentStr
    });
  }
  return result;
}

function doPost(e){
  let body;
  try{ body = JSON.parse(e.postData.contents); }
  catch(err){ return jsonResponse({ error: 'Invalid request body' }); }

  if (!ACTIONS_EXEMPT_FROM_HARD_STOP.includes(body.action) && !isAdvisorActive()){
    return jsonResponse(Object.assign({ error: 'ADVISOR_INACTIVE' }, getAdvisorActiveStatus()));
  }

  if (body.action === 'setDuesPreference')      return jsonResponse(setDuesPreference(body.policyNumber, body.enabled));
  if (body.action === 'setBirthdayPreference')  return jsonResponse(setBirthdayPreference(body.email, body.enabled));
  if (body.action === 'saveConfig')             { saveBrandConfig(body.config || {}); return jsonResponse({ success: true }); }
  if (body.action === 'uploadImage')            { const fileId = uploadBrandImage(body.target, body.base64, body.mimeType); return jsonResponse({ success: true, fileId: fileId }); }
  if (body.action === 'saveAdvisorName')        { saveAdvisorName(body.name || ''); return jsonResponse({ success: true }); }
  if (body.action === 'uploadProfileImage')     { const fileId = uploadProfileImage(body.base64, body.mimeType); return jsonResponse({ success: true, fileId: fileId }); }
  if (body.action === 'setAutoSendStatus')      { setAutoSendStatus(!!body.enabled); return jsonResponse({ success: true }); }
  if (body.action === 'pushDues')               { const result = pushDuesRows(body.rows || []); return jsonResponse(Object.assign({ success: true }, result)); }
  if (body.action === 'pushBirthdays')          { const result = pushBirthdayRows(body.rows || []); return jsonResponse(Object.assign({ success: true }, result)); }
  if (body.action === 'setBirthdayAutoSendStatus') { setBirthdayAutoSendStatus(!!body.enabled); return jsonResponse({ success: true }); }
  if (body.action === 'setSendHour')            { const result = setSendHour(body.hour); return jsonResponse(Object.assign({ success: true }, result)); }
  if (body.action === 'sendDuesTestEmail')      { try{ return jsonResponse(sendDuesTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: err.message }); } }
  if (body.action === 'sendBirthdayTestEmail')  { try{ return jsonResponse(sendBirthdayTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: err.message }); } }

  return jsonResponse({ error: 'Unknown action' });
}

function jsonResponse(obj){
  return ContentService.createTextOutput(JSON.stringify(obj)).setMimeType(ContentService.MimeType.JSON);
}

/* ============================================================
   PUSH PARSED ROWS FROM THE PARSER TOOL
   ============================================================ */
function pushDuesRows(rows){
  setupSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const policyCol = HEADERS.indexOf('Policy Number');
  const lastReminderCol = HEADERS.indexOf('Last Reminder Sent');
  const existingRowByPolicy = {};
  for (let i = 1; i < data.length; i++){
    existingRowByPolicy[String(data[i][policyCol])] = i;
  }
  let added = 0, updated = 0;
  const newRows = [];
  rows.forEach(r => {
    const dueDateValue = r.dueDate ? new Date(r.dueDate) : '';
    const rowValues = [r.policyNumber, r.clientName, r.email, r.product, r.premiumMode, r.premiumAmount, dueDateValue, r.policyStatus, '', true];
    const idx = existingRowByPolicy[String(r.policyNumber)];
    if (idx !== undefined){
      const lastReminderSent = data[idx][lastReminderCol];
      const sendDues = data[idx][HEADERS.indexOf('Send Dues?')];
      data[idx] = rowValues;
      data[idx][lastReminderCol] = lastReminderSent;
      data[idx][HEADERS.indexOf('Send Dues?')] = sendDues;
      updated++;
    } else {
      newRows.push(rowValues);
      added++;
    }
  });
  const fullData = data.concat(newRows);
  sheet.getRange(1, 1, fullData.length, HEADERS.length).setValues(fullData);
  if (rows.length > 0) recordUploadActivity();
  return { added: added, updated: updated, total: rows.length };
}

function pushBirthdayRows(rows){
  setupBirthdaySheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const emailCol = BIRTHDAY_HEADERS.indexOf('Email');
  const lastSentCol = BIRTHDAY_HEADERS.indexOf('Last Greeting Sent (Year)');
  const existingRowByEmail = {};
  for (let i = 1; i < data.length; i++){
    existingRowByEmail[String(data[i][emailCol]).toLowerCase()] = i;
  }
  let added = 0, updated = 0;
  const newRows = [];
  rows.forEach(r => {
    const dobValue = r.dateOfBirth ? new Date(r.dateOfBirth) : '';
    const rowValues = [r.fullName, r.email, r.contactNumber, r.location, dobValue, '', true];
    const idx = existingRowByEmail[String(r.email).toLowerCase()];
    if (idx !== undefined){
      const lastSent = data[idx][lastSentCol];
      const sendBday = data[idx][BIRTHDAY_HEADERS.indexOf('Send Birthday?')];
      data[idx] = rowValues;
      data[idx][lastSentCol] = lastSent;
      data[idx][BIRTHDAY_HEADERS.indexOf('Send Birthday?')] = sendBday;
      updated++;
    } else {
      newRows.push(rowValues);
      added++;
    }
  });
  const fullData = data.concat(newRows);
  sheet.getRange(1, 1, fullData.length, BIRTHDAY_HEADERS.length).setValues(fullData);
  if (rows.length > 0) recordUploadActivity();
  return { added: added, updated: updated, total: rows.length };
}

/* ============================================================
   DAILY REMINDER CHECK
   ============================================================ */
function getTodayDateStr(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

function bumpDailyStat(statKey){
  const props = PropertiesService.getScriptProperties();
  const todayStr = getTodayDateStr();
  const storedDate = props.getProperty(statKey + '_DATE');
  let count = (storedDate === todayStr) ? (Number(props.getProperty(statKey + '_COUNT')) || 0) : 0;
  count++;
  props.setProperty(statKey + '_DATE', todayStr);
  props.setProperty(statKey + '_COUNT', String(count));
}

function getDailyStat(statKey){
  const props = PropertiesService.getScriptProperties();
  const todayStr = getTodayDateStr();
  const storedDate = props.getProperty(statKey + '_DATE');
  if (storedDate !== todayStr) return 0;
  return Number(props.getProperty(statKey + '_COUNT')) || 0;
}

function countDueOnOffset(offsetDays){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const targetDate = new Date();
  targetDate.setDate(targetDate.getDate() + offsetDays);
  const targetStr = Utilities.formatDate(targetDate, tz, 'yyyy-MM-dd');
  let count = 0;
  for (let i = 1; i < data.length; i++){
    const dueDate = data[i][col('Due Date')];
    if (!(dueDate instanceof Date)) continue;
    if (Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd') === targetStr) count++;
  }
  return count;
}

function getDailyStats(){
  return {
    sent: getDailyStat('STAT_SENT'),
    failed: getDailyStat('STAT_FAILED'),
    dueToday: countDueOnOffset(0),
    dueTomorrow: countDueOnOffset(1)
  };
}

function sendDailyReminders(){
  if (!getAutoSendStatus().enabled) return;
  if (!isAdvisorActive()) return; // hard stop: no upload in 90+ days
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const sendDues = row[col('Send Dues?')];
    if (sendDues === false || sendDues === 'FALSE' || sendDues === 0 || sendDues === '0') continue;
    const dueDate = row[col('Due Date')];
    if (!(dueDate instanceof Date)) continue;
    const dueDateStr = Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd');
    const lastSent = row[col('Last Reminder Sent')];
    const lastSentStr = lastSent ? String(lastSent) : '';
    if (dueDateStr === todayStr && lastSentStr !== todayStr){
      let sent = false;
      try{ sent = sendReminderEmail(row, col); }
      catch(err){ bumpDailyStat('STAT_FAILED'); continue; }
      if (sent){
        bumpDailyStat('STAT_SENT');
        sheet.getRange(i + 1, col('Last Reminder Sent') + 1).setValue(todayStr);
        advanceDueDate(sheet, i + 1, col, dueDate, row[col('Premium Mode')]);
      }
    }
  }
}

function sendReminderEmail(row, col){
  const email = row[col('Email')];
  if (!email) return false;
  const config = getBrandConfig();
  assertConfigured(config);

  const clientName = row[col('Client Name')];
  const product = row[col('Product')];
  const amount = row[col('Premium Amount')];
  const dueDate = row[col('Due Date')];
  const policyNumber = row[col('Policy Number')];
  const tz = Session.getScriptTimeZone();
  const subjectDate = Utilities.formatDate(dueDate, tz, 'MMMM d');
  const subject = 'PREMIUM DUE REMINDER - ' + subjectDate.toUpperCase();
  const htmlBody = buildReminderEmailHtml(clientName, policyNumber, product, amount, dueDate, config);

  const options = {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  };
  if (config.contactEmail){
    options.cc = config.contactEmail;
    options.replyTo = config.contactEmail;
  }
  sendWithOptionalFromAlias(email, subject, options, config.contactEmail);
  return true;
}

function sendWithOptionalFromAlias(to, subject, options, fromAlias){
  if (fromAlias){
    try{
      const aliasOptions = Object.assign({}, options, { from: fromAlias });
      GmailApp.sendEmail(to, subject, '', aliasOptions);
      return;
    }catch(err){
      // Alias not verified — fall through and send normally.
    }
  }
  GmailApp.sendEmail(to, subject, '', options);
}

function getEmailImages(config){
  return {
    headerImg: DriveApp.getFileById(config.headerImageFileId).getBlob().setName('header.png'),
    footerImg: DriveApp.getFileById(config.footerImageFileId).getBlob().setName('footer.png')
  };
}

function formatClientName(rawName){
  const name = String(rawName || '').trim();
  if (!name) return '';
  const commaIdx = name.indexOf(',');
  if (commaIdx === -1) return name;
  const lastName = name.slice(0, commaIdx).trim();
  const rest = name.slice(commaIdx + 1).trim();
  const firstName = rest.split(/\s+/)[0] || '';
  return (firstName + ' ' + lastName).trim();
}

function buildReminderEmailHtml(clientName, policyNumber, product, amount, dueDate, config){
  const tz = Session.getScriptTimeZone();
  const formattedAmount = 'PHP ' + Number(amount).toLocaleString('en-PH', { minimumFractionDigits: 2 });
  const formattedDate = Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy');
  const greetingName = formatClientName(clientName);
  return ''
    + '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #E7DFCF;border-radius:10px;overflow:hidden;">'
    + '  <img src="cid:headerImg" alt="Header" style="width:100%;display:block;">'
    + '  <div style="padding:24px;background:#FDF8F0;color:#1C2A38;">'
    + '    <p>Hi ' + greetingName + ',</p>'
    + '    <p>This is a friendly reminder that your premium payment is due <strong>today</strong>.</p>'
    + '    <table style="width:100%;margin:16px 0;border-collapse:collapse;font-size:14px;">'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Policy Number</td><td style="text-align:right;font-weight:700;">' + policyNumber + '</td></tr>'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Product</td><td style="text-align:right;font-weight:700;">' + product + '</td></tr>'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Amount Due</td><td style="text-align:right;font-weight:700;color:#0C447C;">' + formattedAmount + '</td></tr>'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Due Date</td><td style="text-align:right;font-weight:700;">' + formattedDate + '</td></tr>'
    + '    </table>'
    + '    <p>Please settle this at your earliest convenience to keep your policy in force. If you have already made this payment, kindly disregard this reminder.</p>'
    + '    <div style="text-align:center;margin:22px 0;">'
    + '      <a href="' + config.payLink + '" style="display:inline-block;background:#0C447C;color:#FFFFFF;text-decoration:none;padding:14px 30px;border-radius:8px;font-weight:700;font-size:14px;letter-spacing:.5px;">PAY ONLINE NOW</a>'
    + '    </div>'
    + '    <p style="text-align:center;font-size:14px;margin:20px 0 0;">Would you like to have a 15-Minutes policy review with me online?</p>'
    + '    <div style="text-align:center;margin:14px 0 6px;">'
    + '      <a href="' + config.connectLink + '" style="display:inline-block;background:#C99A3B;color:#FFFFFF;text-decoration:none;padding:14px 30px;border-radius:8px;font-weight:700;font-size:14px;letter-spacing:.5px;">CONNECT WITH ME</a>'
    + '    </div>'
    + '    <p style="margin-top:20px;">Thank you,</p>'
    + '  </div>'
    + '  <img src="cid:footerImg" alt="Footer" style="width:100%;display:block;">'
    + '</div>';
}

function advanceDueDate(sheet, rowNum, col, currentDueDate, premiumMode){
  const newDate = new Date(currentDueDate.getTime());
  const mode = (premiumMode || '').trim();
  if (mode === 'Monthly')      newDate.setMonth(newDate.getMonth() + 1);
  else if (mode === 'Quarterly')   newDate.setMonth(newDate.getMonth() + 3);
  else if (mode === 'Half-Yearly') newDate.setMonth(newDate.getMonth() + 6);
  else if (mode === 'Yearly')      newDate.setFullYear(newDate.getFullYear() + 1);
  else return;
  sheet.getRange(rowNum, col('Due Date') + 1).setValue(newDate);
}

function previewReminderEmail(){
  const myEmail = Session.getActiveUser().getEmail();
  const config = getBrandConfig();
  assertConfigured(config);
  const sampleDueDate = new Date();
  const htmlBody = buildReminderEmailHtml('Dela Cruz, Juan Miguel', '0123456789', 'Sample Insurance Plan', 50000, sampleDueDate, config);
  const tz = Session.getScriptTimeZone();
  const subjectDate = Utilities.formatDate(sampleDueDate, tz, 'MMMM d');
  GmailApp.sendEmail(myEmail, 'PREVIEW, PREMIUM DUE REMINDER - ' + subjectDate.toUpperCase(), '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
}

function sendDuesTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfigured(config);
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first, then try the test email again.');
  const sampleDueDate = new Date();
  const htmlBody = buildReminderEmailHtml('Dela Cruz, Juan Miguel', '0123456789', 'Sample Insurance Plan', 50000, sampleDueDate, config);
  const tz = Session.getScriptTimeZone();
  const subjectDate = Utilities.formatDate(sampleDueDate, tz, 'MMMM d');
  sendWithOptionalFromAlias(recipient, 'TEST, PREMIUM DUE REMINDER - ' + subjectDate.toUpperCase(), {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  }, config.contactEmail);
  return { success: true, sentTo: recipient };
}

/* ============================================================
   BIRTHDAY GREETINGS
   ============================================================ */

function getBirthdaysTodayRows(){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayMonth = today.getMonth(), todayDay = today.getDate();
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const dob = row[col('Date of Birth')];
    if (!(dob instanceof Date)) continue;
    if (dob.getMonth() === todayMonth && dob.getDate() === todayDay){
      result.push({
        fullName: row[col('Full Name')],
        email: row[col('Email')],
        location: row[col('Location')],
        dobFormatted: Utilities.formatDate(dob, tz, 'MMMM d'),
        lastGreetingSent: row[col('Last Greeting Sent (Year)')] || ''
      });
    }
  }
  return result;
}

function countBirthdaysOnOffset(offsetDays){
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const target = new Date();
  target.setDate(target.getDate() + offsetDays);
  const targetMonth = target.getMonth(), targetDay = target.getDate();
  let count = 0;
  for (let i = 1; i < data.length; i++){
    const dob = data[i][col('Date of Birth')];
    if (!(dob instanceof Date)) continue;
    if (dob.getMonth() === targetMonth && dob.getDate() === targetDay) count++;
  }
  return count;
}

function getBirthdayDailyStats(){
  return {
    sent: getDailyStat('BDAY_STAT_SENT'),
    failed: getDailyStat('BDAY_STAT_FAILED'),
    birthdaysToday: countBirthdaysOnOffset(0),
    birthdaysTomorrow: countBirthdaysOnOffset(1)
  };
}

function sendDailyBirthdayGreetings(){
  if (!getBirthdayAutoSendStatus().enabled) return;
  if (!isAdvisorActive()) return; // hard stop: no upload in 90+ days
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  if (!sheet) return;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const today = new Date();
  const todayMonth = today.getMonth(), todayDay = today.getDate();
  const currentYearStr = String(today.getFullYear());
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const sendBday = row[col('Send Birthday?')];
    if (sendBday === false || sendBday === 'FALSE' || sendBday === 0 || sendBday === '0') continue;
    const dob = row[col('Date of Birth')];
    if (!(dob instanceof Date)) continue;
    if (dob.getMonth() !== todayMonth || dob.getDate() !== todayDay) continue;
    const lastSentYear = String(row[col('Last Greeting Sent (Year)')] || '');
    if (lastSentYear === currentYearStr) continue;
    let sent = false;
    try{ sent = sendBirthdayEmail(row, col); }
    catch(err){ bumpDailyStat('BDAY_STAT_FAILED'); continue; }
    if (sent){
      bumpDailyStat('BDAY_STAT_SENT');
      sheet.getRange(i + 1, col('Last Greeting Sent (Year)') + 1).setValue(currentYearStr);
    }
  }
}

function sendBirthdayEmail(row, col){
  const email = row[col('Email')];
  if (!email) return false;
  const config = getBrandConfig();
  assertConfiguredForBirthday(config);
  const fullName = row[col('Full Name')];
  // Subject dynamically uses the client's first name
  const subject = 'HAPPY BIRTHDAY, ' + firstNameOnly(fullName).toUpperCase() + '!';
  const htmlBody = buildBirthdayEmailHtml(fullName, config);

  const options = {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  };
  if (config.contactEmail){
    options.cc = config.contactEmail;
    options.replyTo = config.contactEmail;
  }
  sendWithOptionalFromAlias(email, subject, options, config.contactEmail);
  return true;
}

function firstNameOnly(rawName){
  const name = String(rawName || '').trim();
  if (!name) return '';
  const commaIdx = name.indexOf(',');
  if (commaIdx !== -1){
    const rest = name.slice(commaIdx + 1).trim();
    return rest.split(/\s+/)[0] || '';
  }
  return name.split(/\s+/)[0] || '';
}

function buildBirthdayEmailHtml(fullName, config){
  const greetingName = firstNameOnly(fullName);
  const connectBlock = config.connectLink
    ? ('    <p style="text-align:center;font-size:14px;margin:20px 0 0;">If there\u2019s ever anything you need, or you\u2019d simply like to catch up, I\u2019m always just a message away.</p>'
      + '    <div style="text-align:center;margin:14px 0 6px;">'
      + '      <a href="' + config.connectLink + '" style="display:inline-block;background:#C99A3B;color:#FFFFFF;text-decoration:none;padding:14px 30px;border-radius:8px;font-weight:700;font-size:14px;letter-spacing:.5px;">CONNECT WITH ME</a>'
      + '    </div>')
    : '';
  return ''
    + '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #E7DFCF;border-radius:10px;overflow:hidden;">'
    + '  <img src="cid:headerImg" alt="Header" style="width:100%;display:block;">'
    + '  <div style="padding:24px;background:#FDF8F0;color:#1C2A38;text-align:center;">'
    + '    <p style="font-size:18px;font-weight:700;color:#0C447C;margin:0 0 10px;">Happy Birthday, ' + greetingName + '! &#127881;</p>'
    + '    <p style="font-size:14px;">On your special day, I just want you to know how much you\u2019re valued, not only as a client, but as someone I genuinely enjoy staying connected with. Wishing you good health, happiness, and a year ahead filled with everything you\u2019ve been hoping for.</p>'
    + connectBlock
    + '    <p style="margin-top:20px;text-align:left;">Warm regards,</p>'
    + '  </div>'
    + '  <img src="cid:footerImg" alt="Footer" style="width:100%;display:block;">'
    + '</div>';
}

function previewBirthdayEmail(){
  const myEmail = Session.getActiveUser().getEmail();
  const config = getBrandConfig();
  assertConfiguredForBirthday(config);
  const htmlBody = buildBirthdayEmailHtml('Juan Miguel Dela Cruz', config);
  GmailApp.sendEmail(myEmail, 'HAPPY BIRTHDAY, ' + firstNameOnly('Juan Miguel Dela Cruz').toUpperCase() + '!', '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
}

function sendBirthdayTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfiguredForBirthday(config);
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first, then try the test email again.');
  const htmlBody = buildBirthdayEmailHtml('Juan Miguel Dela Cruz', config);
  sendWithOptionalFromAlias(recipient, 'HAPPY BIRTHDAY, ' + firstNameOnly('Juan Miguel Dela Cruz').toUpperCase() + '!', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  }, config.contactEmail);
  return { success: true, sentTo: recipient };
}
