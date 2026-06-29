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
  const missing = CONFIG_KEYS.filter(key => !config[key]);
  if (missing.length > 0){
    throw new Error(
      'Branding not set up yet. Open the app, tap "Setup", fill in ' +
      '"Your branding" (missing: ' + missing.join(', ') + '), and tap ' +
      'SAVE BRANDING before reminders can be sent.'
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
   NEW: CLIENT PREFERENCE MANAGEMENT
   ============================================================ */

// Fetch all clients from Dues Tracker with their current Send Dues? preference
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
    const clientName = row[col('Client Name')];
    const email = row[col('Email')];
    const product = row[col('Product')];
    const premiumMode = row[col('Premium Mode')];
    let premiumAmount = row[col('Premium Amount')];
    const dueDate = row[col('Due Date')];
    const policyStatus = row[col('Policy Status')];
    const sendDues = row[col('Send Dues?')];
    
    if (!policyNum) continue; // skip empty rows
    
    // Parse premium amount - handle currency symbols, spaces, commas
    let parsedAmount = 0;
    if (premiumAmount) {
      if (typeof premiumAmount === 'number') {
        parsedAmount = premiumAmount;
      } else {
        // Remove common currency symbols and spaces, parse as number
        const cleaned = String(premiumAmount).replace(/[^\d.-]/g, '').trim();
        parsedAmount = parseFloat(cleaned) || 0;
      }
    }
    
    result.push({
      policyNumber: policyNum,
      clientName: clientName,
      email: email,
      product: product,
      premiumMode: premiumMode,
      premiumAmount: parsedAmount,
      dueDate: dueDate instanceof Date ? Utilities.formatDate(dueDate, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      policyStatus: policyStatus,
      sendDues: sendDues === true || sendDues === 'TRUE' || sendDues === 1 || sendDues === '1'
    });
  }
  return result;
}

// Fetch all clients from Birthday Tracker with their current Send Birthday? preference
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
    const contactNumber = row[col('Contact Number')];
    const location = row[col('Location')];
    const dob = row[col('Date of Birth')];
    const sendBday = row[col('Send Birthday?')];
    
    if (!fullName || !email) continue; // skip empty rows
    
    result.push({
      fullName: fullName,
      email: email,
      contactNumber: contactNumber,
      location: location,
      dateOfBirth: dob instanceof Date ? Utilities.formatDate(dob, Session.getScriptTimeZone(), 'yyyy-MM-dd') : '',
      sendBirthday: sendBday === true || sendBday === 'TRUE' || sendBday === 1 || sendBday === '1'
    });
  }
  return result;
}

// Save preference toggle for a specific policy (Dues)
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

// Save preference toggle for a specific client (Birthday)
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
  if (action === 'getDuesClientList'){
    return jsonResponse({ clients: getDuesClientList() });
  }
  if (action === 'getBirthdayClientList'){
    return jsonResponse({ clients: getBirthdayClientList() });
  }
  if (action === 'getDueToday'){
    return jsonResponse({ rows: getDueTodayRows() });
  }
  if (action === 'getConfig'){
    return jsonResponse({ config: getBrandConfig() });
  }
  if (action === 'getImagePreview'){
    return jsonResponse(getImagePreviewData(e.parameter.target));
  }
  if (action === 'getDailyStats'){
    return jsonResponse(getDailyStats());
  }
  if (action === 'getAdvisorProfile'){
    return jsonResponse(getAdvisorProfile());
  }
  if (action === 'getProfileImagePreview'){
    return jsonResponse(getProfileImagePreviewData());
  }
  if (action === 'getAutoSendStatus'){
    return jsonResponse(getAutoSendStatus());
  }
  if (action === 'getBirthdaysToday'){
    return jsonResponse({ rows: getBirthdaysTodayRows() });
  }
  if (action === 'getBirthdayDailyStats'){
    return jsonResponse(getBirthdayDailyStats());
  }
  if (action === 'getBirthdayAutoSendStatus'){
    return jsonResponse(getBirthdayAutoSendStatus());
  }
  if (action === 'getSendHour'){
    return jsonResponse(getSendHour());
  }
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

    // Keep the row if it's still due today OR it was already reminded today.
    // (sendDailyReminders() advances Due Date to the next cycle right after
    // sending, so isDueToday alone would make a just-sent policy vanish from
    // this list within the same day — wasSentToday keeps it visible, tagged,
    // the same way the Birthdays Today list keeps showing greeted clients.)
    if (!isDueToday && !wasSentToday) continue;

    result.push({
      policyNumber: row[col('Policy Number')],
      clientName: row[col('Client Name')],
      product: row[col('Product')],
      premiumAmount: row[col('Premium Amount')],
      premiumMode: row[col('Premium Mode')],
      // Always show "today" here, since this is the Due Today list — once
      // sent, the cell itself has already moved to the next cycle's date.
      dueDateFormatted: isDueToday ? Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy') : todayFormatted,
      lastReminderSent: lastSentStr
    });
  }
  return result;
}

function doPost(e){
  let body;
  try{
    body = JSON.parse(e.postData.contents);
  }catch(err){
    return jsonResponse({ error: 'Invalid request body' });
  }

  if (body.action === 'setDuesPreference'){
    return jsonResponse(setDuesPreference(body.policyNumber, body.enabled));
  }
  if (body.action === 'setBirthdayPreference'){
    return jsonResponse(setBirthdayPreference(body.email, body.enabled));
  }
  if (body.action === 'saveConfig'){
    saveBrandConfig(body.config || {});
    return jsonResponse({ success: true });
  }
  if (body.action === 'uploadImage'){
    const fileId = uploadBrandImage(body.target, body.base64, body.mimeType);
    return jsonResponse({ success: true, fileId: fileId });
  }
  if (body.action === 'saveAdvisorName'){
    saveAdvisorName(body.name || '');
    return jsonResponse({ success: true });
  }
  if (body.action === 'uploadProfileImage'){
    const fileId = uploadProfileImage(body.base64, body.mimeType);
    return jsonResponse({ success: true, fileId: fileId });
  }
  if (body.action === 'setAutoSendStatus'){
    setAutoSendStatus(!!body.enabled);
    return jsonResponse({ success: true });
  }
  if (body.action === 'pushDues'){
    const result = pushDuesRows(body.rows || []);
    return jsonResponse(Object.assign({ success: true }, result));
  }
  if (body.action === 'pushBirthdays'){
    const result = pushBirthdayRows(body.rows || []);
    return jsonResponse(Object.assign({ success: true }, result));
  }
  if (body.action === 'setBirthdayAutoSendStatus'){
    setBirthdayAutoSendStatus(!!body.enabled);
    return jsonResponse({ success: true });
  }
  if (body.action === 'setSendHour'){
    const result = setSendHour(body.hour);
    return jsonResponse(Object.assign({ success: true }, result));
  }
  if (body.action === 'sendDuesTestEmail'){
    try{
      return jsonResponse(sendDuesTestEmailToSelf());
    }catch(err){
      return jsonResponse({ success: false, error: err.message });
    }
  }
  if (body.action === 'sendBirthdayTestEmail'){
    try{
      return jsonResponse(sendBirthdayTestEmailToSelf());
    }catch(err){
      return jsonResponse({ success: false, error: err.message });
    }
  }

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
    const rowValues = [
      r.policyNumber, r.clientName, r.email, r.product,
      r.premiumMode, r.premiumAmount, dueDateValue, r.policyStatus, '', true // empty Last Reminder Sent, default Send Dues? to true
    ];
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
    const rowValues = [
      r.fullName, r.email, r.contactNumber, r.location, dobValue, '', true // empty Last Greeting Sent, default Send Birthday? to true
    ];
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
    
    // SKIP if Send Dues? is FALSE
    if (sendDues === false || sendDues === 'FALSE' || sendDues === 0 || sendDues === '0') continue;
    
    const dueDate = row[col('Due Date')];
    if (!(dueDate instanceof Date)) continue;

    const dueDateStr = Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd');
    const lastSent = row[col('Last Reminder Sent')];
    const lastSentStr = lastSent ? String(lastSent) : '';

    if (dueDateStr === todayStr && lastSentStr !== todayStr){
      let sent = false;
      try{
        sent = sendReminderEmail(row, col);
      }catch(err){
        bumpDailyStat('STAT_FAILED');
        continue;
      }
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

  GmailApp.sendEmail(email, subject, '', {
    htmlBody: htmlBody,
    name: config.senderName,
    cc: config.contactEmail,
    replyTo: config.contactEmail,
    inlineImages: getEmailImages(config)
  });
  return true;
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

  if (mode === 'Monthly') newDate.setMonth(newDate.getMonth() + 1);
  else if (mode === 'Quarterly') newDate.setMonth(newDate.getMonth() + 3);
  else if (mode === 'Half-Yearly') newDate.setMonth(newDate.getMonth() + 6);
  else if (mode === 'Yearly') newDate.setFullYear(newDate.getFullYear() + 1);
  else return;

  sheet.getRange(rowNum, col('Due Date') + 1).setValue(newDate);
}

function previewReminderEmail(){
  const myEmail = Session.getActiveUser().getEmail();
  const config = getBrandConfig();
  assertConfigured(config);
  const sampleDueDate = new Date();
  const htmlBody = buildReminderEmailHtml(
    'Dela Cruz, Juan Miguel',
    '0123456789',
    'Sample Insurance Plan',
    50000,
    sampleDueDate,
    config
  );
  const tz = Session.getScriptTimeZone();
  const subjectDate = Utilities.formatDate(sampleDueDate, tz, 'MMMM d');
  GmailApp.sendEmail(myEmail, 'PREVIEW, PREMIUM DUE REMINDER - ' + subjectDate.toUpperCase(), '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
}

// Called from the "CLIENT DUE REMINDER EMAIL TEST" button in the app.
// Sends to config.contactEmail (not Session.getActiveUser()) because the
// active user is unreliable from a Web App request context — contactEmail
// is the advisor's own address, already required by assertConfigured().
function sendDuesTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfigured(config);
  const sampleDueDate = new Date();
  const htmlBody = buildReminderEmailHtml(
    'Dela Cruz, Juan Miguel',
    '0123456789',
    'Sample Insurance Plan',
    50000,
    sampleDueDate,
    config
  );
  const tz2 = Session.getScriptTimeZone();
  const subjectDate2 = Utilities.formatDate(sampleDueDate, tz2, 'MMMM d');
  GmailApp.sendEmail(config.contactEmail, 'TEST, PREMIUM DUE REMINDER - ' + subjectDate2.toUpperCase(), '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
  return { success: true, sentTo: config.contactEmail };
}

/* ============================================================
   BIRTHDAY GREETINGS
   ============================================================ */

function assertConfiguredForBirthday(config){
  const required = ['senderName','contactEmail','headerImageFileId','footerImageFileId'];
  const missing = required.filter(key => !config[key]);
  if (missing.length > 0){
    throw new Error(
      'Branding not set up yet. Open the app, tap "Setup", fill in ' +
      '"Your branding" (missing: ' + missing.join(', ') + '), and tap ' +
      'SAVE BRANDING before birthday greetings can be sent.'
    );
  }
}

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
    
    // SKIP if Send Birthday? is FALSE
    if (sendBday === false || sendBday === 'FALSE' || sendBday === 0 || sendBday === '0') continue;
    
    const dob = row[col('Date of Birth')];
    if (!(dob instanceof Date)) continue;
    if (dob.getMonth() !== todayMonth || dob.getDate() !== todayDay) continue;

    const lastSentYear = String(row[col('Last Greeting Sent (Year)')] || '');
    if (lastSentYear === currentYearStr) continue;

    let sent = false;
    try{
      sent = sendBirthdayEmail(row, col);
    }catch(err){
      bumpDailyStat('BDAY_STAT_FAILED');
      continue;
    }
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

  GmailApp.sendEmail(email, subject, '', {
    htmlBody: htmlBody,
    name: config.senderName,
    cc: config.contactEmail,
    replyTo: config.contactEmail,
    inlineImages: getEmailImages(config)
  });
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
    + '    <p style="font-size:18px;font-weight:700;color:#0C447C;margin:0 0 10px;">Happy Birthday, ' + greetingName + '! \uD83C\uDF89</p>'
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

// Called from the "CLIENT BIRTHDAY EMAIL TEST" button in the app.
// Sends to config.contactEmail for the same reason as the dues test above.
function sendBirthdayTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfiguredForBirthday(config);
  const htmlBody = buildBirthdayEmailHtml('Juan Miguel Dela Cruz', config);
  GmailApp.sendEmail(config.contactEmail, 'TEST \u2013 HAPPY BIRTHDAY GREETING', '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
  return { success: true, sentTo: config.contactEmail };
}
