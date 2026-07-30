/**
 * ============================================================
 * GGE / JCC DUES TRACKER - Apps Script Backend (ENHANCED)
 * With per-advisor client preference toggles
 * ============================================================
 */

const SHEET_NAME = 'Dues Tracker';
const HEADERS = ['Policy Number','Client Name','Email','Product','Premium Mode','Premium Amount','Fund Value','Due Date','Policy Status','Last Reminder Sent','Send Dues?','Lapse Date','Issued Date'];

const BIRTHDAY_SHEET_NAME = 'Birthday Tracker';
const BIRTHDAY_HEADERS = ['Full Name','Email','Contact Number','Location','Date of Birth','Last Greeting Sent (Year)','Send Birthday?'];

// Scheduled broadcasts: each row is one queued send. The full payload
// (subject, body, recipients, attachments, template flag) is stored as
// a JSON string in PayloadJSON, since PropertiesService's 9KB-per-value
// limit is too small once inline images/attachments are included —
// Sheet cells comfortably hold much larger text. TriggerId lets a
// scheduled send be cancelled later by deleting its specific trigger.
const SCHEDULE_SHEET_NAME = 'Scheduled Broadcasts';
const SCHEDULE_HEADERS = ['Schedule ID','Scheduled For','Subject','PayloadJSON','TriggerId','Status','Created At','Sent At','Error','SentCount','FailedCount'];

// Broadcast Drafts: saved (not sent, not scheduled) messages an advisor
// can come back to and finish later, or reuse as a starting point for
// a future broadcast. Same PayloadJSON-in-a-sheet-cell pattern as
// Scheduled Broadcasts, since the same size constraints apply.
const DRAFT_SHEET_NAME = 'Broadcast Drafts';
const DRAFT_HEADERS = ['Draft ID','Subject','PayloadJSON','Created At','Updated At'];

const CONFIG_DEFAULTS = {
  senderName: '',
  contactEmail: '',
  headerImageFileId: '',
  footerImageFileId: '',
  connectLink: '',
  payLink: ''
};
const CONFIG_KEYS = Object.keys(CONFIG_DEFAULTS);

function getBrandConfig(){
  const props = PropertiesService.getScriptProperties();
  const config = {};
  CONFIG_KEYS.forEach(key => {
    config[key] = props.getProperty('CFG_' + key) || CONFIG_DEFAULTS[key];
  });
  return config;
}

// DIAGNOSTIC ONLY \u2014 checks the real, current size of the saved header
// and footer images in Drive, plus what the broadcast template size
// guard would calculate. Run this directly from the Apps Script editor
// (select this function in the dropdown, click Run) to see the actual
// numbers, or call it via ?action=diagnoseTemplateSize in a browser.
function diagnoseTemplateSize(){
  const config = getBrandConfig();
  const result = { headerImageFileId: config.headerImageFileId, footerImageFileId: config.footerImageFileId };

  if (!config.headerImageFileId){
    result.error = 'No headerImageFileId saved in config \u2014 header photo was never uploaded, or config is empty.';
    return result;
  }
  if (!config.footerImageFileId){
    result.error = 'No footerImageFileId saved in config \u2014 footer photo was never uploaded, or config is empty.';
    return result;
  }

  try{
    const headerBlob = DriveApp.getFileById(config.headerImageFileId).getBlob();
    result.headerBytes = headerBlob.getBytes().length;
    result.headerMB = (result.headerBytes / (1024 * 1024)).toFixed(2);
    result.headerMimeType = headerBlob.getContentType();
  }catch(e){
    result.headerError = 'Could not read header file: ' + e.message;
  }

  try{
    const footerBlob = DriveApp.getFileById(config.footerImageFileId).getBlob();
    result.footerBytes = footerBlob.getBytes().length;
    result.footerMB = (result.footerBytes / (1024 * 1024)).toFixed(2);
    result.footerMimeType = footerBlob.getContentType();
  }catch(e){
    result.footerError = 'Could not read footer file: ' + e.message;
  }

  if (result.headerBytes !== undefined && result.footerBytes !== undefined){
    const combinedMB = (result.headerBytes + result.footerBytes) / (1024 * 1024);
    result.combinedMB = combinedMB.toFixed(2);
    result.wouldBeBlockedByOurCheck = combinedMB > 2;
  }

  return result;
}

function saveBrandConfig(partialConfig){
  const props = PropertiesService.getScriptProperties();
  CONFIG_KEYS.forEach(key => {
    if (partialConfig[key] !== undefined){
      props.setProperty('CFG_' + key, String(partialConfig[key]));
    }
  });
}

// Only the truly essential fields are required to send emails.
// contactEmail is optional \u2014 used as cc/replyTo if present, skipped if blank.
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

// Birthday uses the same required fields (contactEmail is optional here too).
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
    // Fund Value sits between Premium Amount and Due Date in HEADERS, so
    // for sheets created before this field existed, it needs to be
    // inserted at that exact position (not appended at the end) or every
    // column after it would end up misaligned with its header.
    if (!headers.includes('Fund Value')){
      const premiumAmtCol = headers.indexOf('Premium Amount');
      const insertAfterCol = premiumAmtCol !== -1 ? premiumAmtCol + 2 : headers.length + 1;
      sheet.insertColumnAfter(insertAfterCol - 1);
      sheet.getRange(1, insertAfterCol).setValue('Fund Value');
    }
    const refreshedHeaders = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!refreshedHeaders.includes('Send Dues?')){
      // Uses this sheet's own current column count (not HEADERS.length,
      // which no longer reflects "Send Dues?" being the last column now
      // that Lapse Date was added after it) — appending after whatever
      // this specific sheet currently has is always correct regardless
      // of how HEADERS itself has grown since.
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('Send Dues?');
      for (let i = 2; i <= sheet.getLastRow(); i++){
        sheet.getRange(i, lastCol).setValue(true);
      }
    }
    const headersAfterSendDues = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headersAfterSendDues.includes('Lapse Date')){
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('Lapse Date');
    }
    const headersAfterLapseDate = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headersAfterLapseDate.includes('Issued Date')){
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('Issued Date');
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

function setupScheduleSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet){
    sheet = ss.insertSheet(SCHEDULE_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0){
    sheet.appendRow(SCHEDULE_HEADERS);
    sheet.setFrozenRows(1);
  } else {
    // Existing sheets created before SentCount/FailedCount existed —
    // append any missing columns at the end rather than inserting them
    // mid-row, so nothing already in the sheet shifts position.
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    SCHEDULE_HEADERS.forEach(h => {
      if (!headers.includes(h)){
        sheet.getRange(1, sheet.getLastColumn() + 1).setValue(h);
      }
    });
  }
  return sheet;
}

function setupDraftSheet(){
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DRAFT_SHEET_NAME);
  if (!sheet){
    sheet = ss.insertSheet(DRAFT_SHEET_NAME);
  }
  if (sheet.getLastRow() === 0){
    sheet.appendRow(DRAFT_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
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
    let parsedFundValue = 0;
    const fundValueCol = col('Fund Value');
    if (fundValueCol !== -1){
      const fundValueRaw = row[fundValueCol];
      if (fundValueRaw){
        parsedFundValue = typeof fundValueRaw === 'number'
          ? fundValueRaw
          : parseFloat(String(fundValueRaw).replace(/[^\d.-]/g, '').trim()) || 0;
      }
    }
    const dueDate = row[col('Due Date')];
    const lapseDateCol = col('Lapse Date');
    const lapseDateRaw = lapseDateCol !== -1 ? row[lapseDateCol] : null;
    const issuedDateCol = col('Issued Date');
    const issuedDateRaw = issuedDateCol !== -1 ? row[issuedDateCol] : null;
    result.push({
      policyNumber: policyNum,
      clientName: row[col('Client Name')],
      email: row[col('Email')],
      product: row[col('Product')],
      premiumMode: row[col('Premium Mode')],
      premiumAmount: parsedAmount,
      fundValue: parsedFundValue,
      dueDate: dueDate instanceof Date ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      lapseDate: lapseDateRaw instanceof Date ? Utilities.formatDate(lapseDateRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      issuedDate: issuedDateRaw instanceof Date ? Utilities.formatDate(issuedDateRaw, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
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
  if (action === 'getDuesClientList')         return jsonResponse({ clients: getDuesClientList() });
  if (action === 'getBirthdayClientList')     return jsonResponse({ clients: getBirthdayClientList() });
  if (action === 'getRemainingEmailQuota')    return jsonResponse(getRemainingEmailQuota());
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
  if (action === 'diagnoseTemplateSize')      return jsonResponse(diagnoseTemplateSize());
  if (action === 'getScheduledBroadcasts')    return jsonResponse({ schedules: getScheduledBroadcasts() });
  if (action === 'getDrafts')                 return jsonResponse({ drafts: getDrafts() });
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
  if (body.action === 'sendDuesTestEmail')      { try{ return jsonResponse(sendDuesTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'sendBirthdayTestEmail')  { try{ return jsonResponse(sendBirthdayTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'sendBroadcastBatch')     { try{ return jsonResponse(Object.assign({ success: true }, sendBroadcastEmailBatch(body.rows || [], body.subject || '', body.htmlBody || '', body.attachments || [], body.useTemplate))); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'scheduleBroadcast')      { try{ return jsonResponse(scheduleBroadcast(body.scheduledFor, body.payload)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'cancelScheduledBroadcast') { try{ return jsonResponse(cancelScheduledBroadcast(body.scheduleId)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'deleteCompletedBroadcast') { try{ return jsonResponse(deleteCompletedBroadcast(body.scheduleId)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'getScheduledBroadcasts') { try{ return jsonResponse({ schedules: getScheduledBroadcasts() }); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'editScheduledBroadcast') { try{ return jsonResponse(editScheduledBroadcast(body.scheduleId, body.scheduledFor, body.payload)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'getScheduledBroadcastPayload') { try{ return jsonResponse(getScheduledBroadcastPayload(body.scheduleId)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'saveDraft')              { try{ return jsonResponse(saveDraft(body.draftId, body.payload)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'getDrafts')              { try{ return jsonResponse({ drafts: getDrafts() }); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'deleteDraft')            { try{ return jsonResponse(deleteDraft(body.draftId)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }

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
  const sendDuesCol = HEADERS.indexOf('Send Dues?');
  const existingRowByPolicy = {};
  for (let i = 1; i < data.length; i++){
    existingRowByPolicy[String(data[i][policyCol])] = i;
  }
  let added = 0, updated = 0;
  const newRows = [];
  rows.forEach(r => {
    const dueDateValue = r.dueDate ? new Date(r.dueDate) : '';
    const lapseDateValue = r.lapseDate ? new Date(r.lapseDate) : '';
    const issuedDateValue = r.issuedDate ? new Date(r.issuedDate) : '';
    // Row order must exactly match HEADERS: Policy Number, Client Name,
    // Email, Product, Premium Mode, Premium Amount, Fund Value, Due Date,
    // Policy Status, Last Reminder Sent, Send Dues?, Lapse Date, Issued Date
    const rowValues = [r.policyNumber, r.clientName, r.email, r.product, r.premiumMode, r.premiumAmount, (r.fundValue || 0), dueDateValue, r.policyStatus, '', true, lapseDateValue, issuedDateValue];
    const idx = existingRowByPolicy[String(r.policyNumber)];
    if (idx !== undefined){
      const lastReminderSent = data[idx][lastReminderCol];
      const sendDues = data[idx][sendDuesCol];
      data[idx] = rowValues;
      data[idx][lastReminderCol] = lastReminderSent;
      data[idx][sendDuesCol] = sendDues;
      updated++;
    } else {
      newRows.push(rowValues);
      added++;
    }
  });
  const fullData = data.concat(newRows);
  sheet.getRange(1, 1, fullData.length, HEADERS.length).setValues(fullData);
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

  // Build options \u2014 cc, replyTo, and from-alias are all optional
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

// Sends via GmailApp, using contactEmail as the visible "From" address
// if it's set up as a verified "Send mail as" alias in the deploying
// account's Gmail settings (Settings > Accounts and Import > Send mail
// as). This is what hides a personal Gmail address from clients \u2014
// without a verified alias, GmailApp.sendEmail() throws "Invalid from
// address", so this always falls back to the account's own address if
// the alias send fails for any reason (not yet verified, removed, etc.),
// rather than letting the whole reminder/test silently fail.
function sendWithOptionalFromAlias(to, subject, options, fromAlias){
  if (fromAlias){
    try{
      const aliasOptions = Object.assign({}, options, { from: fromAlias });
      GmailApp.sendEmail(to, subject, '', aliasOptions);
      return;
    }catch(err){
      // Alias not verified (or any other from-related failure) \u2014
      // fall through and send normally below instead of failing outright.
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

// Test email: sends to contactEmail if set, otherwise falls back to
// the active user's email (which works fine in the Apps Script editor context).
function sendDuesTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfigured(config);
  // Session.getEffectiveUser().getEmail() needs the userinfo.email scope,
  // which some Workspace policies (e.g. Sun Life's) block outright \u2014
  // even though it works fine for personal Gmail deployments. Using
  // contactEmail directly sidesteps that scope entirely.
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
  const subject = 'HAPPY BIRTHDAY FROM ' + (config.senderName || 'YOUR ADVISOR').toUpperCase() + '!';
  const htmlBody = buildBirthdayEmailHtml(fullName, config);

  // Build options \u2014 cc and replyTo are optional
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

/* ============================================================
   BROADCAST EMAIL — custom message sent to all Dues Tracker
   clients, with optional image/PDF attachments and {FirstName}
   personalization. Sent in batches from the front-end (same
   pattern as pushDuesRows) to stay under the 30-second Web App
   execution limit regardless of client list size.
   ============================================================ */

// attachments: array of { base64, mimeType, filename }
// rows: array of { email, clientName } for one batch
// htmlBody: may contain the literal text "{FirstName}" which gets
// replaced per-recipient before sending.
// useTemplate: when true, wraps htmlBody with the same header/footer
// images used by dues reminders and birthday greetings, via cid:
// references + inlineImages — this is the reliable method real email
// clients render correctly, unlike data-URL images which many inboxes
// (including Gmail's own web client in some cases) strip or block.
// Personal Gmail accounts (the common case for individual advisors) are
// capped at 100 email recipients per rolling 24-hour period; Google
// Workspace accounts get up to 1,500/day. Checking this before a large
// broadcast starts sending lets the advisor see the real ceiling and
// decide how to proceed, rather than discovering it partway through a
// send when a batch fails with a generic quota error and no context
// for how much of the list actually has a realistic chance of going out
// today.
function getRemainingEmailQuota(){
  return { remaining: MailApp.getRemainingDailyQuota() };
}

function sendBroadcastEmailBatch(rows, subject, htmlBody, attachments, useTemplate){
  const config = getBrandConfig();
  // Broadcast Email only strictly needs a sender name — header/footer
  // photos are optional here (unlike dues reminders and birthday
  // greetings, which always embed them). Requiring them unconditionally
  // was blocking every broadcast for any advisor who hadn't set up a
  // header/footer yet, even when the "Use header & footer template"
  // toggle was off and no template was going to be embedded at all —
  // this was the actual cause of broadcasts failing to send regardless
  // of message size.
  if (!config.senderName){
    throw new Error(
      'Branding not set up yet. Open the app, tap "Setup", fill in ' +
      '"Your branding" (missing: senderName), and tap SAVE BRANDING before broadcasts can be sent.'
    );
  }
  if (useTemplate && (!config.headerImageFileId || !config.footerImageFileId)){
    throw new Error(
      'You turned on "Use header & footer template" but haven\u2019t saved both a header and footer photo yet. ' +
      'Go to Settings \u2192 Branding Studio to add them, or turn the template toggle off to send without it.'
    );
  }

  // Check the ACTUAL stored header/footer file sizes before attempting
  // anything. Client-side compression only affects newly-uploaded
  // photos going forward — an advisor's existing header/footer (saved
  // before this fix, or re-saved via an older code path) can still be
  // large. Without this check, every single recipient in the batch
  // fails one-by-one with the same size error, which wastes the whole
  // send attempt; this fails once, immediately, with one clear fix.
  let templateMB = 0;
  if (useTemplate){
    const headerBytes = DriveApp.getFileById(config.headerImageFileId).getBlob().getBytes().length;
    const footerBytes = DriveApp.getFileById(config.footerImageFileId).getBlob().getBytes().length;
    templateMB = (headerBytes + footerBytes) / (1024 * 1024);
  }

  // The size check above only ever covered the header/footer template —
  // it completely missed images inserted directly into the message body
  // via the editor's own Image button, which get embedded as base64
  // data URLs right inside htmlBody itself. That gap is exactly what
  // The size check below measures the real total as it will actually
  // be transmitted: the processed HTML body (after converting any
  // embedded data-URL images to cid: references, see below), those
  // extracted body images, the header/footer template, and attachments.
  const attachmentBytesTotal = (attachments || []).reduce((sum, a) => sum + Math.ceil((a.base64 || '').length * 0.75), 0);

  // Any image inserted via the composer's own Image button lands in
  // htmlBody as a literal <img src="data:..."> data URL — this is
  // simple to build client-side, but it means the image's full base64
  // text sits directly inside the HTML body string. GmailApp's
  // body/header size quota (which is separate from, and apparently far
  // stricter than, the 25MB attachment limit) counts that inflated HTML
  // text directly, which is why a broadcast that measured well under
  // our 6MB threshold could still be rejected by Gmail for every single
  // recipient — the real quota-counted size was dominated by base64
  // text sitting in the body, not by attachments or the header/footer
  // template at all. Converting each data-URL image into a proper
  // cid: reference (exactly like the header/footer template already
  // does) moves that image out of the quota-counted HTML text and into
  // a separate inline resource, which is what actually brings the real
  // sent size down to something Gmail's stricter body quota accepts.
  const dataUrlImagePattern = /<img[^>]+src="data:([^;]+);base64,([^"]+)"[^>]*>/g;
  const bodyInlineImages = {};
  let bodyImageCounter = 0;
  let processedHtmlBody = String(htmlBody || '').replace(dataUrlImagePattern, (fullMatch, mimeType, base64Data) => {
    const cid = 'bcBodyImg' + (bodyImageCounter++);
    try{
      const bytes = Utilities.base64Decode(base64Data);
      bodyInlineImages[cid] = Utilities.newBlob(bytes, mimeType, cid);
    }catch(e){
      return fullMatch; // if decoding fails for any reason, leave this one exactly as-is rather than breaking the whole send
    }
    return fullMatch.replace(/src="data:[^"]+"/, 'src="cid:' + cid + '"');
  });
  // Recalculate the real body size using the now-much-smaller HTML (data
  // URLs replaced with short cid: references) plus the actual decoded
  // image bytes, which reflects what Gmail's body quota will actually
  // see much more accurately than measuring the original data-URL-laden
  // htmlBody ever could.
  const processedHtmlBodyBytes = Utilities.newBlob(processedHtmlBody).getBytes().length;
  const bodyImageBytesTotal = Object.values(bodyInlineImages).reduce((sum, blob) => sum + blob.getBytes().length, 0);
  // Every one of these pieces gets MIME/base64-encoded before Gmail
  // actually transmits it, inflating the real wire size over these raw
  // byte counts.
  const MIME_ENCODING_OVERHEAD = 1.37;
  const totalMB = ((processedHtmlBodyBytes / (1024 * 1024)) + (bodyImageBytesTotal / (1024 * 1024)) + templateMB + (attachmentBytesTotal / (1024 * 1024))) * MIME_ENCODING_OVERHEAD;
  // A REAL broadcast that measured ~0.88MB under the OLD calculation
  // (which counted body images as inflated base64 text sitting directly
  // in the HTML, rather than as separate inline resources) was rejected
  // by Gmail for every single recipient. That means Gmail's actual body
  // quota is far stricter than 6MB — this threshold is set with a wide
  // safety margin below that real failure point now that data-URL
  // images are correctly converted to cid: references before this
  // check runs, rather than trying to guess a new raw-byte number that
  // might still be wrong in the same direction.
  if (totalMB > 2){
    throw new Error(
      'This message is too large (~' + totalMB.toFixed(2) + 'MB total once encoded for sending, including any inserted photos, the header/footer template, and attachments) to send reliably. ' +
      'Remove an inline image or attachment, or turn off "Use header & footer template", then try again.'
    );
  }
  if (useTemplate && (templateMB * MIME_ENCODING_OVERHEAD) > 1){
    throw new Error(
      'Your saved header/footer photos are too large (~' + (templateMB * MIME_ENCODING_OVERHEAD).toFixed(2) + 'MB combined once encoded for sending) to embed in every email of this broadcast. ' +
      'Go to Settings \u2192 Branding Studio and re-upload your header and footer photos \u2014 they\u2019ll now be compressed automatically to a safe size. ' +
      'Or turn off "Use header & footer template" for this broadcast and send without it.'
    );
  }

  const blobs = (attachments || []).map(a => {
    const bytes = Utilities.base64Decode(a.base64);
    return Utilities.newBlob(bytes, a.mimeType || 'application/octet-stream', a.filename || 'attachment');
  });

  const wrappedBody = useTemplate
    ? '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #E7DFCF;border-radius:10px;overflow:hidden;">'
      + '<img src="cid:headerImg" alt="Header" style="width:100%;display:block;">'
      + '<div style="padding:24px;background:#FDF8F0;color:#1C2A38;">' + processedHtmlBody + '</div>'
      + '<img src="cid:footerImg" alt="Footer" style="width:100%;display:block;">'
      + '</div>'
    : processedHtmlBody;

  let sent = 0, failed = 0;
  const failedEmails = [];
  const failureReasons = []; // { email, reason } — surfaced to the frontend so
                              // "1 failed" isn't a dead end with no explanation
  let quotaExhausted = false; // once true, every remaining row in this batch is marked failed without attempting to send

  rows.forEach(r => {
    if (!r.email) {
      failed++;
      failureReasons.push({ email: '(blank)', reason: 'No email address on file for this client.' });
      return;
    }
    if (quotaExhausted){
      // The daily sending quota was already confirmed exhausted earlier
      // in this same batch — attempting the remaining recipients would
      // only produce the identical rejection for every one of them
      // (exactly what happened in a real broadcast: 44 and 17 failures,
      // all the same "quota" error). Marking them failed immediately
      // avoids wasting Apps Script execution time on calls guaranteed
      // to fail, and keeps the failure list to one clear explanation
      // instead of dozens of duplicates of the same message.
      failed++;
      failedEmails.push(r.email);
      failureReasons.push({ email: r.email, reason: 'Not attempted \u2014 daily sending quota was already exhausted earlier in this send.' });
      return;
    }
    try{
      const firstName = firstNameOnly(r.clientName) || 'there';
      // Bold + black so the personalized greeting stands out regardless
      // of whatever color the surrounding message text uses.
      const styledFirstName = '<span style="font-weight:700;color:#000000;">' + firstName + '</span>';
      // Case-insensitive match: {FirstName}, {firstName}, {FIRSTNAME},
      // {firstname} all work \u2014 typing the exact capitalization
      // correctly shouldn't be a requirement for personalization to
      // actually apply.
      const personalizedBody = wrappedBody.replace(/\{firstname\}/gi, styledFirstName);

      const options = {
        htmlBody: personalizedBody,
        name: config.senderName,
      };
      // Merges the header/footer template images (when enabled) with
      // any body images that were converted from data URLs to cid:
      // references above \u2014 both need to be present in inlineImages
      // for their respective cid: references in the HTML to resolve.
      const combinedInlineImages = Object.assign({}, useTemplate ? getEmailImages(config) : {}, bodyInlineImages);
      if (Object.keys(combinedInlineImages).length > 0) options.inlineImages = combinedInlineImages;
      if (blobs.length > 0) options.attachments = blobs;
      if (config.contactEmail){
        options.cc = config.contactEmail;
        options.replyTo = config.contactEmail;
      }
      sendWithOptionalFromAlias(r.email, subject, options, config.contactEmail);
      sent++;
    }catch(err){
      failed++;
      failedEmails.push(r.email);
      const translatedReason = toEnglishErrorMessage(err.message || String(err));
      // If this is a size-related failure, attach the exact numbers our
      // own pre-check calculated for this send \u2014 the processed body
      // size, body image total, template size, and attachment total \u2014
      // so if Gmail rejects a send that our pre-check considered safe,
      // the actual breakdown is visible immediately instead of having
      // to guess again which component the pre-check is undercounting.
      const isSizeRelated = /too large|limit exceeded|laki ng body/i.test(translatedReason);
      const diagnosticSuffix = isSizeRelated
        ? ' [diagnostic: htmlBody=' + (processedHtmlBodyBytes/1024).toFixed(0) + 'KB, bodyImages=' + (bodyImageBytesTotal/1024/1024).toFixed(2) + 'MB, template=' + templateMB.toFixed(2) + 'MB, attachments=' + (attachmentBytesTotal/1024/1024).toFixed(2) + 'MB, precheck total=' + totalMB.toFixed(2) + 'MB]'
        : '';
      failureReasons.push({ email: r.email, reason: translatedReason + diagnosticSuffix });
      // A quota-exhausted error is a hard stop for the rest of this
      // batch — unlike a bad email address or an isolated failure,
      // every subsequent send would fail for the exact same reason
      // until the quota resets, which the daily limit does NOT do
      // mid-batch.
      if (/quota|invoked too many times|masyadong madaming beses.*araw/i.test(err.message || '')){
        quotaExhausted = true;
      }
    }
  });

  return { sent: sent, failed: failed, failedEmails: failedEmails, failureReasons: failureReasons, total: rows.length, quotaExhausted: quotaExhausted };
}

/* ============================================================
   SCHEDULED BROADCASTS — lets an advisor queue a broadcast to
   send at a specific future date/time instead of immediately.
   Each schedule gets its own one-time Apps Script trigger
   (ScriptApp...at(specificDateTime)) that fires
   sendScheduledBroadcast() at exactly that moment. The full
   payload (subject, body, recipients, attachments, template
   flag) is stored as JSON in the Scheduled Broadcasts sheet —
   PropertiesService's 9KB-per-value limit is too small once
   inline images/attachments are included, but a sheet cell
   comfortably holds far more. Multiple schedules can be queued
   at once, each with its own row and its own trigger, entirely
   independent of one another.
   ============================================================ */

// scheduledFor: ISO datetime string for when this should send.
// payload: { rows, subject, htmlBody, attachments, useTemplate } —
// exactly the same shape sendBroadcastEmailBatch already accepts,
// just captured now and replayed later at the scheduled time.
function scheduleBroadcast(scheduledFor, payload){
  const scheduledDate = new Date(scheduledFor);
  if (isNaN(scheduledDate.getTime())){
    throw new Error('Invalid scheduled date/time.');
  }
  if (scheduledDate.getTime() <= Date.now()){
    throw new Error('Scheduled time must be in the future.');
  }

  const sheet = setupScheduleSheet();
  const scheduleId = Utilities.getUuid();
  const payloadJson = JSON.stringify(payload || {});

  // One-time trigger, distinct from the recurring daily triggers used
  // elsewhere — this fires exactly once, at exactly this timestamp.
  const trigger = ScriptApp.newTrigger('runScheduledBroadcastTrigger')
    .timeBased()
    .at(scheduledDate)
    .create();
  const triggerId = trigger.getUniqueId();

  // The trigger only knows to call runScheduledBroadcastTrigger() with
  // no arguments (Apps Script time triggers can't carry custom
  // parameters), so the scheduleId has to be recoverable some other
  // way — storing triggerId alongside the row lets the trigger handler
  // look up "which row was I created for" when it fires.
  sheet.appendRow([
    scheduleId,
    scheduledDate,
    payload && payload.subject || '',
    payloadJson,
    triggerId,
    'scheduled',
    new Date(),
    '',
    ''
  ]);

  return { success: true, scheduleId: scheduleId, scheduledFor: scheduledDate.toISOString() };
}

// The actual function every scheduled trigger calls. Since Apps Script
// time-based triggers can't pass custom data, this looks itself up by
// matching the trigger's own unique ID against the TriggerId column —
// whichever row matches is the schedule that just came due.
// Apps Script kills a running script outright once it hits its 6-minute
// execution limit — this does NOT throw a catchable exception, so a
// broadcast to 100+ recipients (each GmailApp.sendEmail() call taking a
// meaningful fraction of a second) could previously get partway through
// sending, hit the time limit, and simply stop — leaving the Error
// column blank and Status stuck wherever it happened to be, since the
// script never reached the line that would have recorded a failure.
// This processes recipients in small time-boxed chunks instead of all
// at once, checking elapsed time frequently, and — if a chunk finishes
// but recipients remain — creates a new one-time trigger to pick up
// exactly where this run left off, chaining across as many trigger
// firings as needed rather than risking one long, uninterruptible run.
const SCHEDULED_BROADCAST_MAX_RUNTIME_MS = 4.5 * 60 * 1000; // 4.5 min, safely under the 6-minute hard limit

function runScheduledBroadcastTrigger(e){
  const triggerId = e && e.triggerUid ? e.triggerUid : null;
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++){
    if (triggerId && String(data[i][col('TriggerId')]) === String(triggerId)){
      rowIndex = i;
      break;
    }
  }

  // Always clean up the one-time trigger regardless of outcome below —
  // it has already fired and will never fire again, so leaving it
  // registered only clutters the project's trigger list.
  if (triggerId){
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
    });
  }

  if (rowIndex === -1) return; // no matching row found — nothing to do

  const rowNum = rowIndex + 1;
  const status = data[rowIndex][col('Status')];
  if (status === 'cancelled'){
    return; // person cancelled it before it fired — do nothing
  }

  const startTime = Date.now();
  const sentCountCol = col('SentCount');
  const failedCountCol = col('FailedCount');
  const alreadySent = Number(data[rowIndex][sentCountCol]) || 0;
  const alreadyFailed = Number(data[rowIndex][failedCountCol]) || 0;

  try{
    const payload = JSON.parse(data[rowIndex][col('PayloadJSON')] || '{}');
    const allRows = payload.rows || [];
    const remainingRows = allRows.slice(alreadySent + alreadyFailed);

    if (remainingRows.length === 0){
      // Nothing left to send (shouldn't normally happen, but covers the
      // edge case of a schedule with zero recipients).
      sheet.getRange(rowNum, col('Status') + 1).setValue('sent');
      sheet.getRange(rowNum, col('Sent At') + 1).setValue(new Date());
      return;
    }

    // Small batches (10 recipients per sendBroadcastEmailBatch call)
    // rather than one giant call for everyone, or one call per single
    // recipient — a single call for 100+ recipients is exactly what
    // silently hit the 6-minute execution wall with no error logged,
    // while calling once per individual recipient would repeat the
    // header/footer size validation (a Drive lookup) 100+ times,
    // adding real overhead against the same time budget this is meant
    // to protect. Checking elapsed time between each small batch is
    // frequent enough to stop safely with time to spare.
    const CHUNK_SIZE = 10;
    let chunkSent = 0, chunkFailed = 0;
    const chunkFailureReasons = [];

    for (let i = 0; i < remainingRows.length; i += CHUNK_SIZE){
      if (Date.now() - startTime > SCHEDULED_BROADCAST_MAX_RUNTIME_MS){
        break; // time's up for this run — whatever's left continues in the next chained trigger
      }
      const batchSlice = remainingRows.slice(i, i + CHUNK_SIZE);
      const batchResult = sendBroadcastEmailBatch(
        batchSlice,
        payload.subject || '',
        payload.htmlBody || '',
        payload.attachments || [],
        payload.useTemplate
      );
      chunkSent += batchResult.sent;
      chunkFailed += batchResult.failed;
      if (batchResult.failureReasons) chunkFailureReasons.push(...batchResult.failureReasons);
    }

    const newSentCount = alreadySent + chunkSent;
    const newFailedCount = alreadyFailed + chunkFailed;
    const totalProcessed = newSentCount + newFailedCount;
    const stillRemaining = allRows.length - totalProcessed;

    sheet.getRange(rowNum, sentCountCol + 1).setValue(newSentCount);
    sheet.getRange(rowNum, failedCountCol + 1).setValue(newFailedCount);

    if (chunkFailureReasons.length > 0){
      const existingError = String(data[rowIndex][col('Error')] || '');
      const newErrorText = chunkFailureReasons.map(fr => fr.email + ': ' + fr.reason).join(' | ');
      sheet.getRange(rowNum, col('Error') + 1).setValue(
        existingError ? existingError + ' | ' + newErrorText : newErrorText
      );
    }

    if (stillRemaining > 0){
      // Not finished — chain a new trigger to continue almost
      // immediately (a few seconds out is the minimum Apps Script
      // allows), rather than leaving the rest unsent.
      const continuationTrigger = ScriptApp.newTrigger('runScheduledBroadcastTrigger')
        .timeBased()
        .after(5000)
        .create();
      sheet.getRange(rowNum, col('TriggerId') + 1).setValue(continuationTrigger.getUniqueId());
      sheet.getRange(rowNum, col('Status') + 1).setValue('sending');
    } else {
      const allFailed = newSentCount === 0 && newFailedCount > 0;
      sheet.getRange(rowNum, col('Status') + 1).setValue(allFailed ? 'failed' : 'sent');
      sheet.getRange(rowNum, col('Sent At') + 1).setValue(new Date());
    }
  }catch(err){
    // Per the advisor's own instruction: if anything is missing or
    // broken by the time this fires (recipient list changed, Setup URL
    // gone, branding incomplete), skip silently and just log the
    // error in the sheet — no additional notification.
    sheet.getRange(rowNum, col('Status') + 1).setValue('failed');
    sheet.getRange(rowNum, col('Error') + 1).setValue(toEnglishErrorMessage(err.message || String(err)));
  }
}

// Lists all non-cancelled schedules, most recently created first, for
// display in the Broadcast Email UI so the advisor can see what's
// queued and cancel anything before it fires.
function getScheduledBroadcasts(){
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    result.push({
      scheduleId: row[col('Schedule ID')],
      scheduledFor: row[col('Scheduled For')] instanceof Date ? row[col('Scheduled For')].toISOString() : String(row[col('Scheduled For')]),
      subject: row[col('Subject')],
      status: row[col('Status')],
      createdAt: row[col('Created At')] instanceof Date ? row[col('Created At')].toISOString() : String(row[col('Created At')]),
      sentAt: row[col('Sent At')] instanceof Date ? row[col('Sent At')].toISOString() : String(row[col('Sent At')] || ''),
      error: row[col('Error')] || '',
      sentCount: Number(row[col('SentCount')]) || 0,
      failedCount: Number(row[col('FailedCount')]) || 0
    });
  }
  result.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return result;
}

// getScheduledBroadcasts() intentionally omits PayloadJSON (recipient
// lists and attachments can be large, and the list view never needs
// them) — this fetches the full payload for exactly one schedule,
// used only when the advisor opens a queued broadcast to edit it.
function getScheduledBroadcastPayload(scheduleId){
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Schedule ID')]) === String(scheduleId)){
      return { payload: JSON.parse(data[i][col('PayloadJSON')] || '{}') };
    }
  }
  return { payload: null };
}

// Deletes the underlying trigger (so it can never fire) and marks the
// row cancelled rather than deleting it outright, so there's still a
// record of what was scheduled and cancelled.
function cancelScheduledBroadcast(scheduleId){
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Schedule ID')]) === String(scheduleId)){
      const triggerId = data[i][col('TriggerId')];
      if (triggerId){
        ScriptApp.getProjectTriggers().forEach(t => {
          if (t.getUniqueId() === triggerId) ScriptApp.deleteTrigger(t);
        });
      }
      sheet.getRange(i + 1, col('Status') + 1).setValue('cancelled');
      return { success: true };
    }
  }
  return { success: false, error: 'Schedule not found.' };
}

// Permanently removes a finished broadcast's row so it stops cluttering
// the Completed Broadcasts list on the dashboard. Only allowed for
// broadcasts that are actually done (sent/failed) — a scheduled or
// currently-sending broadcast still has a live trigger and recipients
// waiting, so those must go through cancelScheduledBroadcast instead,
// which also cleans up that trigger before removing anything.
function deleteCompletedBroadcast(scheduleId){
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Schedule ID')]) === String(scheduleId)){
      const status = data[i][col('Status')];
      if (status !== 'sent' && status !== 'failed'){
        return { success: false, error: 'Only completed (sent or failed) broadcasts can be removed this way \u2014 cancel it instead if it\u2019s still scheduled or sending.' };
      }
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Broadcast not found.' };
}

// Edits an already-queued scheduled broadcast — the message, recipients,
// attachments, template flag, and/or the send time itself, all before
// it fires. Only works while status is still 'scheduled' (a broadcast
// that already sent, failed, or was cancelled can't be edited back into
// existence — start a new one instead).
//
// If scheduledFor is provided and differs from the stored time, the old
// one-time trigger is deleted and a new one created at the new time —
// Apps Script triggers have no "reschedule" operation, only delete and
// recreate. If scheduledFor is omitted, only the payload changes and
// the existing trigger/time are left untouched.
function editScheduledBroadcast(scheduleId, scheduledFor, payload){
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  let rowIndex = -1;
  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Schedule ID')]) === String(scheduleId)){
      rowIndex = i;
      break;
    }
  }
  if (rowIndex === -1){
    return { success: false, error: 'Schedule not found.' };
  }
  const rowNum = rowIndex + 1;
  const currentStatus = data[rowIndex][col('Status')];
  if (currentStatus === 'sending'){
    return { success: false, error: 'This broadcast is currently sending and can\u2019t be edited right now \u2014 wait for it to finish, or cancel it to stop the remaining recipients.' };
  }
  if (currentStatus !== 'scheduled'){
    return { success: false, error: 'This broadcast already ' + currentStatus + ' and can no longer be edited.' };
  }

  let newTriggerId = data[rowIndex][col('TriggerId')];
  if (scheduledFor){
    const newDate = new Date(scheduledFor);
    if (isNaN(newDate.getTime())){
      return { success: false, error: 'Invalid scheduled date/time.' };
    }
    if (newDate.getTime() <= Date.now()){
      return { success: false, error: 'Scheduled time must be in the future.' };
    }
    // Create the new trigger FIRST, before deleting the old one — if
    // trigger creation is going to fail (e.g. the project has hit
    // Apps Script's per-user trigger limit), it's better to fail with
    // the old schedule still intact than to have already deleted the
    // working trigger and then fail to replace it, which would leave
    // this broadcast silently unscheduled with no trigger at all.
    const newTrigger = ScriptApp.newTrigger('runScheduledBroadcastTrigger')
      .timeBased()
      .at(newDate)
      .create();
    newTriggerId = newTrigger.getUniqueId();

    const oldTriggerId = data[rowIndex][col('TriggerId')];
    if (oldTriggerId && oldTriggerId !== newTriggerId){
      // getProjectTriggers() scans every trigger across every feature
      // in this project (daily reminders, birthday greetings, every
      // other scheduled broadcast) — with many broadcasts queued this
      // can be slow enough to push total execution time toward the
      // 30-second Web App limit, which surfaces to the browser as a
      // raw network failure ("Load failed") rather than a clean error.
      // This still does the scan, but only once, after the new trigger
      // already exists, so a slow scan here never risks leaving the
      // broadcast without any working trigger.
      ScriptApp.getProjectTriggers().forEach(t => {
        if (t.getUniqueId() === oldTriggerId) ScriptApp.deleteTrigger(t);
      });
    }
    sheet.getRange(rowNum, col('Scheduled For') + 1).setValue(newDate);
    sheet.getRange(rowNum, col('TriggerId') + 1).setValue(newTriggerId);
  }

  if (payload){
    sheet.getRange(rowNum, col('Subject') + 1).setValue(payload.subject || '');
    sheet.getRange(rowNum, col('PayloadJSON') + 1).setValue(JSON.stringify(payload));
  }

  return { success: true };
}

/* ============================================================
   BROADCAST DRAFTS — lets an advisor save a message (subject,
   body, recipients, attachments, template flag) without sending
   or scheduling it, to finish later or reuse as a starting
   point. Unlike Scheduled Broadcasts, drafts never trigger
   anything on their own — they just sit until explicitly opened,
   edited, or deleted. Same PayloadJSON-in-a-cell pattern, for
   the same reason (attachments/images are too big for
   PropertiesService's 9KB-per-value limit).
   ============================================================ */

// Creates a new draft if draftId is omitted, or overwrites an existing
// one if provided — lets "Save Draft" double as "update this draft"
// once the advisor has saved it once and keeps editing.
function saveDraft(draftId, payload){
  const sheet = setupDraftSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const payloadJson = JSON.stringify(payload || {});
  const now = new Date();

  if (draftId){
    for (let i = 1; i < data.length; i++){
      if (String(data[i][col('Draft ID')]) === String(draftId)){
        sheet.getRange(i + 1, col('Subject') + 1).setValue(payload && payload.subject || '');
        sheet.getRange(i + 1, col('PayloadJSON') + 1).setValue(payloadJson);
        sheet.getRange(i + 1, col('Updated At') + 1).setValue(now);
        return { success: true, draftId: draftId };
      }
    }
    // draftId was provided but not found (e.g. it was deleted elsewhere) —
    // fall through and create a fresh one instead of silently failing.
  }

  const newDraftId = Utilities.getUuid();
  sheet.appendRow([newDraftId, payload && payload.subject || '', payloadJson, now, now]);
  return { success: true, draftId: newDraftId };
}

// Lists all saved drafts, most recently updated first, for display in
// the Broadcast Email UI so the advisor can pick one up where they
// left off.
function getDrafts(){
  const sheet = setupDraftSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    result.push({
      draftId: row[col('Draft ID')],
      subject: row[col('Subject')],
      payload: JSON.parse(row[col('PayloadJSON')] || '{}'),
      createdAt: row[col('Created At')] instanceof Date ? row[col('Created At')].toISOString() : String(row[col('Created At')]),
      updatedAt: row[col('Updated At')] instanceof Date ? row[col('Updated At')].toISOString() : String(row[col('Updated At')])
    });
  }
  result.sort((a, b) => new Date(b.updatedAt) - new Date(a.updatedAt));
  return result;
}

function deleteDraft(draftId){
  const sheet = setupDraftSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Draft ID')]) === String(draftId)){
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Draft not found.' };
}

// GmailApp.sendEmail() (and other Google services) return error messages
// in whatever language the deploying Google account's locale is set to —
// not necessarily English, e.g. Tagalog: "Nalagpasan ang Limitasyon: Laki
// ng Body ng E-mail". Since every message shown to the advisor must be
// English, this recognizes the known Gmail error patterns we've actually
// seen and maps them to a clean English equivalent. Anything unrecognized
// falls through unchanged rather than being silently hidden — better to
// show an unfamiliar-but-honest message than to guess wrong.
function toEnglishErrorMessage(rawMessage){
  const msg = String(rawMessage || '');
  const patterns = [
    { match: /Limitasyon.*Laki ng Body|Body.*[Ll]imit exceeded|Nalagpasan.*[Ll]imitasyon/i,
      english: 'Email is too large to send \u2014 remove an inline image or attachment and try again.' },
    { match: /Invalid email|[Mm]ali ang email|hindi wasto ang email/i,
      english: 'Invalid email address.' },
    // Matches both the English quota message ("Service invoked too
    // many times for one day") and its actual Tagalog form as returned
    // by Google for this account ("...masyadong madaming beses
    // pinagana sa isang araw...") — the previous pattern only matched
    // the word "limitasyon" near "araw", which this specific message
    // never contains, so it fell through untranslated and displayed as
    // raw Tagalog error text instead of a clear, actionable English
    // explanation.
    { match: /quota|limitasyon.*araw|daily.*limit|invoked too many times|masyadong madaming beses.*araw/i,
      english: 'Daily sending limit reached for this Google account (personal Gmail accounts get 100 emails/day; Google Workspace accounts get up to 1,500/day) \u2014 try again after the quota resets, or send the rest tomorrow.' },
    { match: /Recipient address required|kinakailangan ang address/i,
      english: 'Recipient address is missing or invalid.' },
    { match: /rate limit|masyadong marami/i,
      english: 'Sending too fast \u2014 please wait a moment and try again.' },
  ];
  for (const p of patterns){
    if (p.match.test(msg)) return p.english;
  }
  // Unrecognized message: return as-is rather than hide it, so nothing
  // gets silently swallowed if a new/unseen Gmail error shows up.
  return msg;
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
  GmailApp.sendEmail(myEmail, 'PREVIEW \u2013 HAPPY BIRTHDAY GREETING', '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
}

// Test email: sends to contactEmail if set, otherwise falls back to
// the active user's email (works in the Apps Script editor context).
function sendBirthdayTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfiguredForBirthday(config);
  // Same scope restriction as the dues test email \u2014 use contactEmail
  // directly instead of Session.getEffectiveUser().
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first, then try the test email again.');
  const htmlBody = buildBirthdayEmailHtml('Juan Miguel Dela Cruz', config);
  sendWithOptionalFromAlias(recipient, 'TEST \u2013 HAPPY BIRTHDAY GREETING', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  }, config.contactEmail);
  return { success: true, sentTo: recipient };
}
