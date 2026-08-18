/**
 * ============================================================
 * GGE / JCC DUES TRACKER - Apps Script Backend (ENHANCED)
 * With per-advisor client preference toggles
 * ============================================================
 */

// The Sheet ID is stored in Script Properties so this same script
// works for any buyer — each one runs setSpreadsheetId() once with
// their own Sheet ID, and the script remembers it from then on.
// To set it: open Apps Script editor, run setSpreadsheetId('YOUR_SHEET_ID')
// or call it from the browser: ?action=setSpreadsheetId&id=YOUR_SHEET_ID
function getSpreadsheet(){
  // Works in both contexts:
  // 1. Bound to a sheet (Extensions → Apps Script) → getActiveSpreadsheet() works
  // 2. Standalone script with Sheet ID set → openById() works
  try{
    const active = SpreadsheetApp.getActiveSpreadsheet();
    if (active) return active;
  }catch(e){}
  const id = PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID');
  if (!id) throw new Error('Sheet ID not set. Run ?action=setSpreadsheetId&id=YOUR_SHEET_ID first.');
  return SpreadsheetApp.openById(id);
}

function setSpreadsheetId(id){
  PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id);
  return 'Connected to sheet: ' + id;
}

const SHEET_NAME = 'Dues Tracker';
const HEADERS = ['Policy Number','Client Name','Email','Product','Premium Mode','Premium Amount','Fund Value','Due Date','Policy Status','Last Reminder Sent','Send Dues?','Lapse Date','Issued Date','Last Anniversary Sent (Year)','Send Anniversary?','Advance Reminder Sent'];

// ============================================================
// TRIGGER VERSION MARKER
// ------------------------------------------------------------
// Bump this string every time you push an update that touches
// sendDailyReminders, sendDailyAnniversaryGreetings,
// dailyInactivityCheck, or watchdogEnsureTriggers — anything that
// changes what the automatic triggers actually DO. A time-driven
// trigger stays permanently bound to whatever deployment version was
// active the moment it was created; simply deploying a new version
// does NOT make existing triggers pick up the new code, only a fresh
// trigger does. Without this marker, every buyer would need the
// exact same manual "delete every trigger, reopen the app" cleanup
// after every single future update, forever.
//
// With it: the self-heal functions below compare this value against
// what's stored in each buyer's own Script Properties. A mismatch
// means their triggers were built under older code, so they get
// automatically deleted and recreated — picking up whatever's
// current — the next time anything touches the app (or within an
// hour via the watchdog). No manual steps needed for buyers ever
// again after this point; you only ever touch this one line, here,
// in the one master copy you distribute.
// ============================================================
const TRIGGER_CODE_VERSION = '2026-08-18-v1';

// Shared by every trigger self-heal below. Deletes any existing
// trigger(s) for the given handler function if the stored version
// marker doesn't match TRIGGER_CODE_VERSION (stale — rebuild), or if
// no trigger exists at all yet (first-time install). Returns true if
// a (re)create actually happened, false if the existing trigger was
// already current and nothing needed to change.
function _rebuildTriggerIfStale(handlerFunctionName, createFn){
  const props = PropertiesService.getScriptProperties();
  const versionKey = 'TRIGGER_VERSION_' + handlerFunctionName;
  const storedVersion = props.getProperty(versionKey);
  const existing = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === handlerFunctionName);

  if (existing.length > 0 && storedVersion === TRIGGER_CODE_VERSION){
    return false; // already current, nothing to do
  }

  // Either stale (version mismatch) or missing entirely — clear out
  // any existing trigger(s) for this handler and build fresh.
  existing.forEach(t => ScriptApp.deleteTrigger(t));
  createFn();
  props.setProperty(versionKey, TRIGGER_CODE_VERSION);
  return true;
}

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
  payLink: '',
  reviewLink: ''
};
const CONFIG_KEYS = Object.keys(CONFIG_DEFAULTS);

/* ============================================================
   ADVISOR ACTIVITY / HARD-STOP CHECK
   ------------------------------------------------------------
   Every real upload (pushDuesRows or pushBirthdayRows) stamps
   LAST_UPLOAD_TIMESTAMP. The deadline is calendar-anchored rather
   than a flat rolling count: it's the 1st day of the month AFTER
   the last upload, plus INACTIVITY_GRACE_DAYS. E.g. an upload any
   time in June anchors to July 1, giving a deadline of ~July 31 —
   a predictable "you have through the end of next month" cadence
   tied to calendar months, instead of an arbitrary window measured
   from the exact upload timestamp. Once that deadline passes, the
   advisor is considered inactive and the script hard-stops:
   doGet/doPost refuse every action except the ones needed to let
   them reactivate (uploading new data, or checking their own
   status), and the time-triggered senders (daily dues reminders,
   birthday greetings, scheduled broadcasts) skip themselves
   entirely so nothing sends on a stale account.
   ============================================================ */
const INACTIVITY_GRACE_DAYS = 30;

// Actions allowed to run even while hard-stopped. Uploading is
// deliberately included — it's the only way to self-reactivate — and
// the status/branding reads are included so the app can render a clear
// "you're locked out, upload to continue" screen instead of a raw error.
const ACTIONS_EXEMPT_FROM_HARD_STOP = [
  'pushDues', 'pushBirthdays', 'getAdvisorActiveStatus', 'getConfig', 'getAdvisorProfile', 'getProfileImagePreview', 'setSpreadsheetId'
];

function recordUploadActivity(){
  PropertiesService.getScriptProperties().setProperty('LAST_UPLOAD_TIMESTAMP', new Date().toISOString());
}

// 1st day of the month following the given timestamp, plus the grace
// window — e.g. an upload on June 25 anchors to July 1, deadline
// July 31. Computed in UTC purely for clean month-boundary math; a
// few hours of timezone drift doesn't matter for a 30-day business
// rule like this one.
function computeInactivityDeadline(lastUploadIso){
  const lastUpload = new Date(lastUploadIso);
  const anchor = new Date(Date.UTC(lastUpload.getUTCFullYear(), lastUpload.getUTCMonth() + 1, 1));
  return new Date(anchor.getTime() + INACTIVITY_GRACE_DAYS * 24 * 60 * 60 * 1000);
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
  const deadline = computeInactivityDeadline(lastUpload);
  const msRemaining = deadline.getTime() - Date.now();
  const daysSince = (Date.now() - new Date(lastUpload).getTime()) / (24 * 60 * 60 * 1000);
  return {
    active: msRemaining > 0,
    lastUploadTimestamp: lastUpload,
    daysSinceLastUpload: Math.floor(daysSince),
    daysUntilLock: Math.max(0, Math.ceil(msRemaining / (24 * 60 * 60 * 1000))),
    cycleDeadline: deadline.toISOString()
  };
}

function isAdvisorActive(){
  return getAdvisorActiveStatus().active;
}

/* ============================================================
   AUTO-CLEAR CLIENT DATA — 30 DAYS FROM THE 1ST OF THE MONTH
   ------------------------------------------------------------
   Runs once a day (installed alongside the other daily triggers —
   see createInactivityPurgeTrigger). Once the calendar-anchored
   deadline computed in getAdvisorActiveStatus() has passed, every
   stored client record (Dues Tracker, Birthday Tracker, Scheduled
   Broadcasts, Broadcast Drafts) is wiped — but a CSV snapshot is
   saved to Drive first, and the advisor is emailed a notice, so
   this is never a silent, unrecoverable surprise. Guarded so it
   only fires once per inactivity stretch: a fresh upload resets
   LAST_UPLOAD_TIMESTAMP, which automatically re-arms it for the
   next cycle.
   ============================================================ */
function purgeInactiveClientData(){
  const status = getAdvisorActiveStatus();
  if (status.active) return; // deadline hasn't passed yet — not due

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_PURGE_BASELINE') === status.lastUploadTimestamp){
    return; // already purged for this inactivity stretch
  }

  const ss = getSpreadsheet();
  const summary = { duesCleared: 0, birthdaysCleared: 0, scheduledCleared: 0, draftsCleared: 0, backupFileUrl: null };

  try{
    summary.backupFileUrl = backupClientDataToDrive(ss);
  }catch(e){
    // Backup failing should never block the purge itself — being past
    // the inactivity deadline means data minimization takes priority.
  }

  summary.duesCleared      = clearSheetDataRows(ss, SHEET_NAME, HEADERS.length);
  summary.birthdaysCleared = clearSheetDataRows(ss, BIRTHDAY_SHEET_NAME, BIRTHDAY_HEADERS.length);
  summary.scheduledCleared = clearScheduledBroadcastsAndTriggers(ss);
  summary.draftsCleared    = clearSheetDataRows(ss, DRAFT_SHEET_NAME, DRAFT_HEADERS.length);

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

// Scheduled Broadcasts rows can carry a live one-time trigger — delete
// those before clearing the row so nothing fires against data that's
// about to disappear.
function clearScheduledBroadcastsAndTriggers(ss){
  const sheet = ss.getSheetByName(SCHEDULE_SHEET_NAME);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  if (data.length <= 1) return 0;
  const triggerIdCol = data[0].indexOf('TriggerId');
  const projectTriggers = ScriptApp.getProjectTriggers();
  for (let i = 1; i < data.length; i++){
    const triggerId = triggerIdCol !== -1 ? data[i][triggerIdCol] : null;
    if (triggerId){
      projectTriggers.forEach(t => {
        if (t.getUniqueId() === String(triggerId)) ScriptApp.deleteTrigger(t);
      });
    }
  }
  const clearedCount = data.length - 1;
  sheet.getRange(2, 1, clearedCount, data[0].length).clearContent();
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
// sidesteps that entirely, same as the rest of this file.
function notifyAdvisorOfPurge(summary){
  try{
    const config = getBrandConfig();
    const recipient = config.contactEmail;
    if (!recipient) return; // nowhere configured to send it
    const body =
      'This is an automated notice from Client Pulse.\n\n' +
      'No client data has been uploaded in time, so all stored client information has been automatically cleared:\n\n' +
      '- Dues Tracker: ' + summary.duesCleared + ' row(s) cleared\n' +
      '- Birthday Tracker: ' + summary.birthdaysCleared + ' row(s) cleared\n' +
      '- Scheduled Broadcasts: ' + summary.scheduledCleared + ' row(s) cleared\n' +
      '- Broadcast Drafts: ' + summary.draftsCleared + ' row(s) cleared\n\n' +
      (summary.backupFileUrl ? ('A backup was saved before deletion, in case you need it back:\n' + summary.backupFileUrl + '\n\n') : '') +
      'Upload a new client list anytime to reactivate the app and start fresh.';
    GmailApp.sendEmail(recipient, 'Client Pulse: client data auto-cleared (inactive too long)', body);
  }catch(e){
    // A failed notification should never surface as a broken purge —
    // the data clearing above has already completed successfully.
  }
}

/* ============================================================
   5-DAY ADVANCE REMINDER
   ------------------------------------------------------------
   Runs daily alongside the purge check (see dailyInactivityCheck).
   Once the advisor is within REMINDER_DAYS_BEFORE_LOCK days of the
   hard-stop/purge deadline, sends one heads-up email so they get a
   chance to re-upload before everything gets wiped — not just find
   out after the fact. Fires once per inactivity cycle, same
   once-only guard pattern as the purge itself; a fresh upload
   automatically re-arms it for the next cycle.
   ============================================================ */
const REMINDER_DAYS_BEFORE_LOCK = 5;

function sendInactivityReminderIfNeeded(){
  const status = getAdvisorActiveStatus();
  if (!status.active) return; // already past the deadline — the purge handles this case
  if (status.daysUntilLock > REMINDER_DAYS_BEFORE_LOCK) return; // not close enough yet

  const props = PropertiesService.getScriptProperties();
  if (props.getProperty('LAST_REMINDER_BASELINE') === status.lastUploadTimestamp){
    return; // already reminded for this cycle
  }

  try{
    const config = getBrandConfig();
    const recipient = config.contactEmail;
    if (recipient){
      const days = status.daysUntilLock;
      const body =
        'You have ' + days + ' day' + (days === 1 ? '' : 's') + ' remaining to reupload your life policy list and client list.\n\n' +
        'After which the app will clear your data and stop auto reminder.\n\n' +
        'Upload your latest Dues or Birthday list anytime before then to keep Client Pulse running without interruption.';
      GmailApp.sendEmail(recipient, 'Client Pulse: ' + days + ' day' + (days === 1 ? '' : 's') + ' left before your data is cleared', body);
    }
  }catch(e){
    // A failed reminder should never block the purge check from running.
  }

  props.setProperty('LAST_REMINDER_BASELINE', status.lastUploadTimestamp);
}

// Installed on the daily 3AM trigger — sends the 5-day advance
// reminder first, then runs the purge check. Order matters only in
// that both should run every day; the reminder's own "already past
// deadline" guard means it naturally stops firing once the purge has
// taken over for that cycle.
function dailyInactivityCheck(){
  sendInactivityReminderIfNeeded();
  purgeInactiveClientData();
}

function getBrandConfig(){
  const props = PropertiesService.getScriptProperties();
  const config = {};
  CONFIG_KEYS.forEach(key => {
    config[key] = props.getProperty('CFG_' + key) || CONFIG_DEFAULTS[key];
  });
  return config;
}

// DIAGNOSTIC ONLY — checks the real, current size of the saved header
// and footer images in Drive, plus what the broadcast template size
// guard would calculate. Run this directly from the Apps Script editor
// (select this function in the dropdown, click Run) to see the actual
// numbers, or call it via ?action=diagnoseTemplateSize in a browser.
function diagnoseTemplateSize(){
  const config = getBrandConfig();
  const result = { headerImageFileId: config.headerImageFileId, footerImageFileId: config.footerImageFileId };

  if (!config.headerImageFileId){
    result.error = 'No headerImageFileId saved in config — header photo was never uploaded, or config is empty.';
    return result;
  }
  if (!config.footerImageFileId){
    result.error = 'No footerImageFileId saved in config — footer photo was never uploaded, or config is empty.';
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
// contactEmail is optional — used as cc/replyTo if present, skipped if blank.
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
  createInactivityPurgeTrigger(); // self-installs once; cheap no-op after that
  ensureAnniversaryDailyTriggerExists(); // same reasoning — advisors from before this feature existed need this too
  ensureDailyReminderTriggerExists(); // same fix, now for dues (due-today + advance reminders)
  const ss = getSpreadsheet();
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
    const headersAfterIssuedDate = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headersAfterIssuedDate.includes('Last Anniversary Sent (Year)')){
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('Last Anniversary Sent (Year)');
    }
    const headersAfterAnniversarySent = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headersAfterAnniversarySent.includes('Send Anniversary?')){
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('Send Anniversary?');
      // Default existing rows to included, same convention as Send Dues?
      for (let i = 2; i <= sheet.getLastRow(); i++){
        sheet.getRange(i, lastCol).setValue(true);
      }
    }
    const headersAfterSendAnniversary = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    if (!headersAfterSendAnniversary.includes('Advance Reminder Sent')){
      const lastCol = sheet.getLastColumn() + 1;
      sheet.getRange(1, lastCol).setValue('Advance Reminder Sent');
      // Left blank for existing rows — blank means "never sent", same
      // convention as Last Reminder Sent. Tracks the advance touchpoint
      // completely separately from the due-today touchpoint, so neither
      // can ever affect the other's "already sent today" check.
    }
  }
  const policyColIndex = HEADERS.indexOf('Policy Number') + 1;
  sheet.getRange(1, policyColIndex, sheet.getMaxRows(), 1).setNumberFormat('@');
}

function setupBirthdaySheet(){
  const ss = getSpreadsheet();
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
  const ss = getSpreadsheet();
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
  const ss = getSpreadsheet();
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

// How many days before the actual due date to start sending advance
// reminders (1 = only the day before, up to 7 = a full week of daily
// nudges leading up to the due date). Default 1 preserves the original
// "day-before" behavior for advisors who never touch this setting.
function getAdvanceDays(){
  const val = PropertiesService.getScriptProperties().getProperty('ADVANCE_DAYS');
  return { advanceDays: val === null ? 1 : Number(val) };
}

function setAdvanceDays(n){
  n = Number(n);
  if (!(n >= 1 && n <= 7)){
    throw new Error('Advance days must be between 1 and 7.');
  }
  PropertiesService.getScriptProperties().setProperty('ADVANCE_DAYS', String(n));
  return { advanceDays: n };
}

function setSendHour(hour){
  hour = Number(hour);
  if (!(hour >= 6 && hour <= 16)){
    throw new Error('Send hour must be between 6 (6-7AM) and 16 (4-5PM).');
  }
  PropertiesService.getScriptProperties().setProperty('SEND_HOUR', String(hour));
  createDailyTrigger(hour);
  createBirthdayDailyTrigger(hour);
  createAnniversaryDailyTrigger(hour);
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

// Cheap, idempotent self-install for advisors whose dues trigger never
// got created — createDailyTrigger() was previously only ever called
// from setSendHour(), meaning any advisor who never explicitly touched
// the Send Hour dropdown had NO daily trigger at all, and
// sendDailyReminders() simply never ran — including both the due-today
// AND advance-reminder emails, since both fire from that same function.
// This only ensures the TRIGGER exists (same idempotent pattern as
// ensureAnniversaryDailyTriggerExists) — it does NOT enable auto-send
// for anyone who hasn't already turned it on; getAutoSendStatus().enabled
// still gates whether anything actually sends.
//
// LockService-protected: this runs on EVERY request now (see doGet/
// doPost), and a single app page load fires several requests nearly
// simultaneously. Without a lock, multiple concurrent calls could each
// see "not installed yet" before any of them finished creating one,
// and each go on to create their own — which is exactly what produced
// a pile of duplicate triggers in practice. The lock fully serializes
// the check-then-create sequence so only one can ever win.
//
// Also now version-aware via _rebuildTriggerIfStale — createDailyTrigger()
// is passed in as the builder, so a version mismatch (or a missing
// trigger) both result in a correct, fresh trigger bound to whatever
// is currently deployed, with no manual cleanup ever needed again.
function ensureDailyReminderTriggerExists(){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(3000);
  }catch(e){
    return; // couldn't get the lock in time — safe to skip this cycle, the next request (or the hourly watchdog) will try again
  }
  try{
    _rebuildTriggerIfStale('sendDailyReminders', createDailyTrigger);
  }finally{
    lock.releaseLock();
  }
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
// Same lock protection as ensureDailyReminderTriggerExists and the
// other self-heal functions — this one runs from setupSheet() (called
// during uploads) and was missed in the first pass, which is exactly
// why dailyInactivityCheck showed up duplicated when the others didn't.
// Also now version-aware via _rebuildTriggerIfStale, same as the others.
function createInactivityPurgeTrigger(){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(3000);
  }catch(e){
    return;
  }
  try{
    // Clean up the old direct-purge trigger from before the 5-day
    // reminder existed, so it isn't left running alongside the new
    // combined check.
    ScriptApp.getProjectTriggers().forEach(t => {
      if (t.getHandlerFunction() === 'purgeInactiveClientData') ScriptApp.deleteTrigger(t);
    });
    _rebuildTriggerIfStale('dailyInactivityCheck', () => {
      ScriptApp.newTrigger('dailyInactivityCheck')
        .timeBased()
        .everyDays(1)
        .atHour(3)
        .create();
    });
  }finally{
    lock.releaseLock();
  }
}

/* ============================================================
   TRIGGER WATCHDOG
   ------------------------------------------------------------
   The self-heal in doGet/doPost (ensureDailyReminderTriggerExists,
   ensureAnniversaryDailyTriggerExists) only ever runs when someone
   actually interacts with the app. Real evidence (Aug 16) showed a
   trigger genuinely missing at its scheduled morning fire time, then
   getting self-healed later that same day once the app was opened —
   correctly fixed for the NEXT day, but too late to catch that day's
   window. This hourly watchdog is a second, independent layer that
   doesn't depend on anyone opening anything: even on a day nobody
   touches the app at all, a missing trigger gets caught and repaired
   within an hour instead of silently persisting indefinitely.
   ============================================================ */
function watchdogEnsureTriggers(){
  ensureDailyReminderTriggerExists();
  ensureAnniversaryDailyTriggerExists();
}

// Self-installs once, same idempotent pattern as every other trigger
// installer in this file — safe to call repeatedly, cheap no-op once
// the watchdog itself already exists and is current. Same lock
// protection and version-awareness as the others.
function ensureWatchdogTriggerExists(){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(3000);
  }catch(e){
    return;
  }
  try{
    _rebuildTriggerIfStale('watchdogEnsureTriggers', () => {
      ScriptApp.newTrigger('watchdogEnsureTriggers')
        .timeBased()
        .everyHours(1)
        .create();
    });
  }finally{
    lock.releaseLock();
  }
}

/* ============================================================
   CLIENT PREFERENCE MANAGEMENT
   ============================================================ */

function getDuesClientList(){
  setupSheet();
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
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
    const advanceReminderCol = col('Advance Reminder Sent');
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
      sendDues: row[col('Send Dues?')] === true || row[col('Send Dues?')] === 'TRUE' || row[col('Send Dues?')] === 1 || row[col('Send Dues?')] === '1',
      lastReminderSent: normalizeDateCellToYmd(row[col('Last Reminder Sent')], Session.getScriptTimeZone()),
      advanceReminderSent: advanceReminderCol !== -1 ? normalizeDateCellToYmd(row[advanceReminderCol], Session.getScriptTimeZone()) : ''
    });
  }
  return result;
}

function getBirthdayClientList(){
  setupBirthdaySheet();
  const sheet = getSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
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
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
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
  const sheet = getSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
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
  if (action === 'setSpreadsheetId')          { const id = e.parameter.id || ''; if (!id) return jsonResponse({ error: 'Missing id parameter' }); PropertiesService.getScriptProperties().setProperty('SPREADSHEET_ID', id); return jsonResponse({ success: true, message: 'Connected to sheet: ' + id }); }
  // All other actions need the sheet to be configured first
  if (!PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')){
    return jsonResponse({ error: 'SHEET_NOT_CONFIGURED', message: 'Run ?action=setSpreadsheetId&id=YOUR_SHEET_ID first.' });
  }
  // Self-heals the dues + anniversary triggers on EVERY single request,
  // not just ones that happen to call setupSheet() — previously a
  // missing/silently-deleted trigger could sit broken for days until
  // something specific (like an upload) happened to touch it. This has
  // been the recurring root cause behind several days of "advance
  // reminders didn't fire" investigations, none of which ever found a
  // logic bug — because the trigger simply wasn't there to run. This
  // check is cheap (just enumerating triggers, no Sheet access), so
  // running it unconditionally here is worth the guarantee.
  ensureDailyReminderTriggerExists();
  ensureAnniversaryDailyTriggerExists();
  ensureWatchdogTriggerExists(); // second, independent layer — catches a missing trigger within an hour even on a day nobody opens the app
  if (!ACTIONS_EXEMPT_FROM_HARD_STOP.includes(action) && !isAdvisorActive()){
    return jsonResponse(Object.assign({ error: 'ADVISOR_INACTIVE' }, getAdvisorActiveStatus()));
  }
  if (action === 'getDuesClientList')         return jsonResponse({ clients: getDuesClientList() });
  if (action === 'getBirthdayClientList')     return jsonResponse({ clients: getBirthdayClientList() });
  if (action === 'getRemainingEmailQuota')    return jsonResponse(getRemainingEmailQuota());
  if (action === 'getDueToday')               return jsonResponse({ rows: getDueTodayRows() });
  if (action === 'getConfig')                 return jsonResponse({ config: getBrandConfig() });
  if (action === 'getImagePreview')           return jsonResponse(getImagePreviewData(e.parameter.target));
  // GET-based push handlers — POST from custom domains is blocked by
  // Google's CORS redirect flow. Sending data as base64-encoded JSON
  // in a GET parameter bypasses this entirely since GET requests are
  // never redirected by Apps Script's authorization layer.
  if (action === 'pushDuesGet'){
    try{
      const arrays = JSON.parse(decodeURIComponent(e.parameter.data || '[]'));
      // Expand positional arrays back to named objects that pushDuesRows expects.
      // Field order matches what callBackendGet sends:
      // [policyNumber, clientName, email, product, premiumMode, premiumAmount,
      //  fundValue, dueDate, policyStatus, lapseDate, issuedDate]
      const rows = arrays.map(function(a){
        if (Array.isArray(a)) {
          return { policyNumber:a[0], clientName:a[1], email:a[2], product:a[3],
            premiumMode:a[4], premiumAmount:a[5], fundValue:a[6],
            dueDate:a[7], policyStatus:a[8], lapseDate:a[9], issuedDate:a[10] };
        }
        // Fallback: handle both slim {pn,cn...} and full {policyNumber,...} objects
        return { policyNumber:a.pn||a.policyNumber, clientName:a.cn||a.clientName,
          email:a.em||a.email, product:a.pr||a.product, premiumMode:a.pm||a.premiumMode,
          premiumAmount:a.pa||a.premiumAmount, fundValue:a.fv||a.fundValue,
          dueDate:a.dd||a.dueDate, policyStatus:a.ps||a.policyStatus,
          lapseDate:a.ld||a.lapseDate, issuedDate:a.id||a.issuedDate };
      });
      return jsonResponse(Object.assign({ success: true }, pushDuesRows(rows)));
    }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); }
  }
  if (action === 'pushBirthdaysGet'){
    try{
      const arrays = JSON.parse(decodeURIComponent(e.parameter.data || '[]'));
      // Field order: [fullName, email, dateOfBirth, contactNumber, location]
      const rows = arrays.map(function(a){
        if (Array.isArray(a)) {
          return { fullName:a[0], email:a[1], dateOfBirth:a[2], contactNumber:a[3], location:a[4] };
        }
        return { fullName:a.fn||a.fullName, email:a.em||a.email,
          dateOfBirth:a.dob||a.dateOfBirth, contactNumber:a.ct||a.contactNumber,
          location:a.lo||a.location };
      });
      return jsonResponse(Object.assign({ success: true }, pushBirthdayRows(rows)));
    }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); }
  }
  if (action === 'getDailyStats')             return jsonResponse(getDailyStats());
  if (action === 'getLastReminderRunLog')     return jsonResponse(getLastReminderRunLog());
  if (action === 'getLastAnniversaryRunLog')  return jsonResponse(getLastAnniversaryRunLog());
  if (action === 'getAdvisorProfile')         return jsonResponse(getAdvisorProfile());
  if (action === 'getProfileImagePreview')    return jsonResponse(getProfileImagePreviewData());
  if (action === 'getAutoSendStatus')         return jsonResponse(getAutoSendStatus());
  if (action === 'getBirthdaysToday')         return jsonResponse({ rows: getBirthdaysTodayRows() });
  if (action === 'getBirthdayDailyStats')     return jsonResponse(getBirthdayDailyStats());
  if (action === 'getBirthdayAutoSendStatus') return jsonResponse(getBirthdayAutoSendStatus());
  if (action === 'getAnniversariesToday')     return jsonResponse({ rows: getPolicyAnniversariesTodayRows() });
  if (action === 'getAnniversaryDailyStats')  return jsonResponse(getAnniversaryDailyStats());
  if (action === 'getAnniversaryAutoSendStatus') return jsonResponse(getAnniversaryAutoSendStatus());
  if (action === 'getSendHour')               return jsonResponse(getSendHour());
  if (action === 'getAdvanceDays')            return jsonResponse(getAdvanceDays());
  if (action === 'diagnoseTemplateSize')      return jsonResponse(diagnoseTemplateSize());
  if (action === 'getScheduledBroadcasts')    return jsonResponse({ schedules: getScheduledBroadcasts() });
  if (action === 'getSentEmailsForSubject')   return getSentEmailsForSubject(e.parameter.subject || '');
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

// Google Sheets often silently auto-converts a date-like string (e.g.
// "2026-07-13") into a real Date value once it's written to a cell —
// even though the cell still visibly displays "the date is there".
// Comparing that against a plain string with === then fails, since
// String(dateObject) produces something like "Mon Jul 13 2026
// 00:00:00 GMT+0800..." rather than "2026-07-13". This normalizes
// either shape (Date object or string) back to a clean yyyy-MM-dd, so
// "was this already sent today" comparisons work regardless of which
// form Sheets actually stored.
function normalizeDateCellToYmd(value, tz){
  if (!value) return '';
  if (value instanceof Date) return Utilities.formatDate(value, tz, 'yyyy-MM-dd');
  const str = String(value).trim();
  // A stray value that looks like a full date-and-time string (Sheets
  // occasionally hands back a string in this shape too) still needs
  // parsing down to just the date part, not a raw string compare.
  const parsed = new Date(str);
  if (!isNaN(parsed.getTime()) && /\d{4}/.test(str)) {
    return Utilities.formatDate(parsed, tz, 'yyyy-MM-dd');
  }
  return str;
}

function getDueTodayRows(){
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const todayFormatted = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');

  // Combines BOTH reminder touchpoints into one list — due today, AND
  // the single advance-notice day (today + Advance Days) — matching
  // exactly what sendDailyReminders() itself checks. Previously this
  // only ever showed the due-today half, so the dashboard never
  // reflected the advance reminders that were actually about to go out.
  // advanceDays is always >=1 (clamped in setAdvanceDays), so the two
  // target dates can never land on the same day.
  const advanceDays = getAdvanceDays().advanceDays;
  const advanceTargetDate = new Date();
  advanceTargetDate.setDate(advanceTargetDate.getDate() + advanceDays);
  const advanceTargetStr = Utilities.formatDate(advanceTargetDate, tz, 'yyyy-MM-dd');
  const advanceTargetFormatted = Utilities.formatDate(advanceTargetDate, tz, 'MMMM d, yyyy');

  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const dueDate = row[col('Due Date')];
    const dueDateStr = (dueDate instanceof Date) ? Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd') : '';
    const isDueToday = dueDateStr === todayStr;
    const isAdvanceDay = dueDateStr === advanceTargetStr;

    // Bug fix: this used to check ONLY "Last Reminder Sent" regardless of
    // which touchpoint applied — but sendDailyReminders() now writes
    // advance sends to a completely separate "Advance Reminder Sent"
    // column. Checking the wrong column here meant an advance-day row
    // could show a stale/misleading "sent" badge from leftover due-today
    // data, or fail to show a real advance send that actually happened —
    // this now checks whichever column actually matches the touchpoint.
    const advanceReminderColIdx = col('Advance Reminder Sent');
    const relevantSentStr = isAdvanceDay && advanceReminderColIdx !== -1
      ? normalizeDateCellToYmd(row[advanceReminderColIdx], tz)
      : normalizeDateCellToYmd(row[col('Last Reminder Sent')], tz);
    const wasSentToday = relevantSentStr === todayStr;

    if (!isDueToday && !isAdvanceDay && !wasSentToday) continue;

    let daysAhead, dueDateFormatted;
    if (isDueToday){
      daysAhead = 0;
      dueDateFormatted = Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy');
    } else if (isAdvanceDay){
      daysAhead = advanceDays;
      dueDateFormatted = advanceTargetFormatted;
    } else {
      // Neither touchpoint matches anymore (e.g. already sent today and
      // the due date has since advanced) — kept in the list so a just-sent
      // reminder doesn't vanish mid-session, shown against today's date.
      daysAhead = 0;
      dueDateFormatted = todayFormatted;
    }

    result.push({
      policyNumber: row[col('Policy Number')],
      clientName: row[col('Client Name')],
      product: row[col('Product')],
      premiumAmount: row[col('Premium Amount')],
      premiumMode: row[col('Premium Mode')],
      dueDateFormatted: dueDateFormatted,
      daysAhead: daysAhead,
      lastReminderSent: relevantSentStr
    });
  }
  result.sort((a, b) => a.daysAhead - b.daysAhead); // due-today first, then the advance-day items
  return result;
}

function doPost(e){
  let body;
  try{ body = JSON.parse(e.postData.contents); }
  catch(err){ return jsonResponse({ error: 'Invalid request body' }); }

  if (!ACTIONS_EXEMPT_FROM_HARD_STOP.includes(body.action) && !PropertiesService.getScriptProperties().getProperty('SPREADSHEET_ID')){
    return jsonResponse({ error: 'SHEET_NOT_CONFIGURED', message: 'Run ?action=setSpreadsheetId&id=YOUR_SHEET_ID first.' });
  }

  // Same self-heal as doGet — every POST request also re-verifies these
  // triggers exist, not just ones that happen to call setupSheet().
  ensureDailyReminderTriggerExists();
  ensureAnniversaryDailyTriggerExists();
  ensureWatchdogTriggerExists();

  if (!ACTIONS_EXEMPT_FROM_HARD_STOP.includes(body.action) && !isAdvisorActive()){
    return jsonResponse(Object.assign({ error: 'ADVISOR_INACTIVE' }, getAdvisorActiveStatus()));
  }

  if (body.action === 'setDuesPreference')      return jsonResponse(setDuesPreference(body.policyNumber, body.enabled));
  if (body.action === 'setBirthdayPreference')  return jsonResponse(setBirthdayPreference(body.email, body.enabled));
  if (body.action === 'setAnniversaryPreference') return jsonResponse(setAnniversaryPreference(body.policyNumber, body.enabled));
  if (body.action === 'saveConfig')             { saveBrandConfig(body.config || {}); return jsonResponse({ success: true }); }
  if (body.action === 'uploadImage')            { const fileId = uploadBrandImage(body.target, body.base64, body.mimeType); return jsonResponse({ success: true, fileId: fileId }); }
  if (body.action === 'saveAdvisorName')        { saveAdvisorName(body.name || ''); return jsonResponse({ success: true }); }
  if (body.action === 'uploadProfileImage')     { const fileId = uploadProfileImage(body.base64, body.mimeType); return jsonResponse({ success: true, fileId: fileId }); }
  if (body.action === 'setAutoSendStatus')      { setAutoSendStatus(!!body.enabled); return jsonResponse({ success: true }); }
  if (body.action === 'pushDues')               { const result = pushDuesRows(body.rows || []); return jsonResponse(Object.assign({ success: true }, result)); }
  if (body.action === 'pushBirthdays')          { const result = pushBirthdayRows(body.rows || []); return jsonResponse(Object.assign({ success: true }, result)); }
  if (body.action === 'setBirthdayAutoSendStatus') { setBirthdayAutoSendStatus(!!body.enabled); return jsonResponse({ success: true }); }
  if (body.action === 'setAnniversaryAutoSendStatus') { setAnniversaryAutoSendStatus(!!body.enabled); return jsonResponse({ success: true }); }
  if (body.action === 'setSendHour')            { const result = setSendHour(body.hour); return jsonResponse(Object.assign({ success: true }, result)); }
  if (body.action === 'setAdvanceDays')         { try{ const result = setAdvanceDays(body.days); return jsonResponse(Object.assign({ success: true }, result)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'sendDuesTestEmail')      { try{ return jsonResponse(sendDuesTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'sendBirthdayTestEmail')  { try{ return jsonResponse(sendBirthdayTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'sendAnniversaryTestEmail') { try{ return jsonResponse(sendAnniversaryTestEmailToSelf()); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'sendBroadcastBatch')     { try{ return jsonResponse(Object.assign({ success: true }, sendBroadcastEmailBatch(body.rows || [], body.subject || '', body.htmlBody || '', body.attachments || [], body.useTemplate))); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'scheduleBroadcast')      { try{ return jsonResponse(scheduleBroadcast(body.scheduledFor, body.payload)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'cancelScheduledBroadcast') { try{ return jsonResponse(cancelScheduledBroadcast(body.scheduleId)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'deleteCompletedBroadcast') { try{ return jsonResponse(deleteCompletedBroadcast(body.scheduleId)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'annotateResendOnSchedule') { try{ return jsonResponse(annotateResendOnSchedule(body.scheduleId, body.sentCount, body.failedCount)); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
  if (body.action === 'getScheduledBroadcasts') { try{ return jsonResponse({ schedules: getScheduledBroadcasts() }); }catch(err){ return jsonResponse({ success: false, error: toEnglishErrorMessage(err.message) }); } }
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
   ------------------------------------------------------------
   OPTIMIZED: previously this rewrote the ENTIRE sheet (every
   existing row + new rows) on every single batch call, via
   data.concat(newRows) + one big setValues() over the whole
   table. That write cost scaled with TOTAL sheet size, so as
   the client list grew, every batch got slower — even though
   each batch only actually changes a small number of rows.
   Now: existing data is still read once (needed to detect
   duplicates and preserve Last Reminder Sent / Send Dues?),
   but only the rows that actually changed get written —
   matched rows are updated in place one small range at a time,
   and brand-new rows are appended in a single bulk write right
   after the last existing row. Write cost now scales with
   batch size, not total sheet size.
   ============================================================ */
function pushDuesRows(rows){
  setupSheet();
  const sheet = SpreadsheetApp.getActiveSpreadsheet()
  .getSheetByName("Dues Tracker");

if (!sheet) {
  throw new Error("Sheet 'Dues Tracker' not found.");
}
  const data = sheet.getDataRange().getValues();
  const policyCol = HEADERS.indexOf('Policy Number');
  const lastReminderCol = HEADERS.indexOf('Last Reminder Sent');
  const sendDuesCol = HEADERS.indexOf('Send Dues?');
  const advanceReminderCol = HEADERS.indexOf('Advance Reminder Sent');
  const existingRowByPolicy = {};
  for (let i = 1; i < data.length; i++){
    existingRowByPolicy[String(data[i][policyCol])] = i;
  }
  // Use the sheet's actual current column count (its header row), same
  // fallback the original full-rewrite logic relied on — keeps this
  // working the same way on older sheets with fewer columns than
  // HEADERS, since setupSheet() above has already backfilled any
  // missing ones by this point.
  const numCols = data[0] ? data[0].length : HEADERS.length;

  let added = 0, updated = 0;
  const newRows = [];
  const updatedRowWrites = []; // { rowNumber, values } — only rows that actually changed

  rows.forEach(r => {
    const dueDateValue = r.dueDate ? new Date(r.dueDate) : '';
    const lapseDateValue = r.lapseDate ? new Date(r.lapseDate) : '';
    const issuedDateValue = r.issuedDate ? new Date(r.issuedDate) : '';
    // Row order must exactly match HEADERS: Policy Number, Client Name,
    // Email, Product, Premium Mode, Premium Amount, Fund Value, Due Date,
    // Policy Status, Last Reminder Sent, Send Dues?, Lapse Date, Issued Date,
    // Last Anniversary Sent (Year), Send Anniversary?, Advance Reminder Sent
    const rowValues = [r.policyNumber, r.clientName, r.email, r.product, r.premiumMode, r.premiumAmount, (r.fundValue || 0), dueDateValue, r.policyStatus, '', true, lapseDateValue, issuedDateValue, '', true, ''];
    const idx = existingRowByPolicy[String(r.policyNumber)];
    if (idx !== undefined){
      const lastReminderSent = data[idx][lastReminderCol];
      const sendDues = data[idx][sendDuesCol];
      const advanceReminderSent = advanceReminderCol !== -1 ? data[idx][advanceReminderCol] : '';
      rowValues[lastReminderCol] = lastReminderSent;
      rowValues[sendDuesCol] = sendDues;
      rowValues[advanceReminderCol] = advanceReminderSent;
      updatedRowWrites.push({ rowNumber: idx + 1, values: rowValues });
      updated++;
    } else {
      newRows.push(rowValues);
      added++;
    }
  });

  // Matched rows are scattered across the sheet, so each still needs
  // its own range write — but each one is now 1 row x numCols cells,
  // not the whole table.
  updatedRowWrites.forEach(w => {
    sheet.getRange(w.rowNumber, 1, 1, numCols).setValues([w.values]);
  });

  // Brand-new rows go in one bulk write right after the last existing row.
  if (newRows.length > 0){
    const startRow = data.length + 1;
    sheet.getRange(startRow, 1, newRows.length, numCols).setValues(newRows);
  }

  if (rows.length > 0) recordUploadActivity();
  return { added: added, updated: updated, total: rows.length };
}

function pushBirthdayRows(rows){
  setupBirthdaySheet();
  const sheet = getSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const emailCol = BIRTHDAY_HEADERS.indexOf('Email');
  const lastSentCol = BIRTHDAY_HEADERS.indexOf('Last Greeting Sent (Year)');
  const sendBdayCol = BIRTHDAY_HEADERS.indexOf('Send Birthday?');
  const existingRowByEmail = {};
  for (let i = 1; i < data.length; i++){
    existingRowByEmail[String(data[i][emailCol]).toLowerCase()] = i;
  }
  const numCols = BIRTHDAY_HEADERS.length;

  let added = 0, updated = 0;
  const newRows = [];
  const updatedRowWrites = [];

  rows.forEach(r => {
    const dobValue = r.dateOfBirth ? new Date(r.dateOfBirth) : '';
    const rowValues = [r.fullName, r.email, r.contactNumber, r.location, dobValue, '', true];
    const idx = existingRowByEmail[String(r.email).toLowerCase()];
    if (idx !== undefined){
      const lastSent = data[idx][lastSentCol];
      const sendBday = data[idx][sendBdayCol];
      rowValues[lastSentCol] = lastSent;
      rowValues[sendBdayCol] = sendBday;
      updatedRowWrites.push({ rowNumber: idx + 1, values: rowValues });
      updated++;
    } else {
      newRows.push(rowValues);
      added++;
    }
  });

  updatedRowWrites.forEach(w => {
    sheet.getRange(w.rowNumber, 1, 1, numCols).setValues([w.values]);
  });

  if (newRows.length > 0){
    const startRow = data.length + 1;
    sheet.getRange(startRow, 1, newRows.length, numCols).setValues(newRows);
  }

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
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
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
  // Two real gaps fixed here, not just more visibility:
  // 1. The log write was only ever called at specific points inside the
  //    function — if ANYTHING threw an uncaught error anywhere else
  //    (e.g. the sheet write below, or advanceDueDate()), the whole
  //    function would crash and the log would never get saved at all,
  //    even though some rows may have already sent successfully before
  //    the crash. Wrapping the entire body in try/finally guarantees
  //    the log is written no matter what happens or where it happens.
  // 2. Each row's post-send work (writing Last Reminder Sent, advancing
  //    the due date) is now its OWN try/catch — previously an error
  //    there would crash the entire loop, silently ending the run and
  //    leaving every row positioned after it in the sheet completely
  //    unprocessed, with no record of why. Isolating it means one bad
  //    row can never take down everyone who comes after it.
  const runLog = [];
  let matchedCount = 0, skippedAlreadySent = 0, sentCount = 0, failedCount = 0;

  try{
    if (!getAutoSendStatus().enabled){ runLog.push('EXIT: auto-send is OFF'); return; }
    if (!isAdvisorActive()){ runLog.push('EXIT: advisor inactive (hard-stop)'); return; }
    setupSheet(); // ensures the Advance Reminder Sent column exists before the loop below reads/writes it, same convention sendDailyAnniversaryGreetings already uses
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet){ runLog.push('EXIT: Dues Tracker sheet not found'); return; }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);
    const tz = Session.getScriptTimeZone();
    const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

    // Sends exactly TWO reminders per policy — one early heads-up on the
    // single day that's exactly advanceDays before the due date, and one
    // more on the actual due date itself. Nothing in between: a policy due
    // in 3 days with advanceDays=3 gets one email today (daysAhead=3) and
    // stays silent until the due date itself, when it gets the second and
    // final email (daysAhead=0) — no daily countdown.
    const advanceDays = getAdvanceDays().advanceDays;
    const advanceTargetDate = new Date();
    advanceTargetDate.setDate(advanceTargetDate.getDate() + advanceDays);
    const advanceTargetStr = Utilities.formatDate(advanceTargetDate, tz, 'yyyy-MM-dd');

    const startMsg = 'START — today=' + todayStr + ' advanceDays=' + advanceDays + ' advanceTarget=' + advanceTargetStr + ' totalRows=' + (data.length - 1);
    Logger.log('[sendDailyReminders] ' + startMsg);
    runLog.push(startMsg);
    saveReminderRunLog(runLog); // saved immediately, not just at the end — see note below on why

    for (let i = 1; i < data.length; i++){
      const row = data[i];
      const sendDues = row[col('Send Dues?')];
      if (sendDues === false || sendDues === 'FALSE' || sendDues === 0 || sendDues === '0') continue;
      // Second layer of defense: the frontend now nulls Due Date for
      // lapsed rows before pushing, but this check protects against a
      // lapsed policy ever getting a due-date reminder even if it slips
      // through with a real Due Date some other way (a manual sheet edit,
      // a future upload-path change, etc.) — same isLapsedStatus() guard
      // the Policy Anniversary Greeter already uses for the same reason.
      if (isLapsedStatus(row[col('Policy Status')])) continue;
      const dueDate = row[col('Due Date')];
      if (!(dueDate instanceof Date)) continue;
      const dueDateStr = Utilities.formatDate(dueDate, tz, 'yyyy-MM-dd');

      // advanceDays is always >=1 (clamped in setAdvanceDays), so these
      // two target dates can never collide — a policy is never matched
      // by both conditions on the same day.
      let daysAhead;
      if (dueDateStr === advanceTargetStr) daysAhead = advanceDays;
      else if (dueDateStr === todayStr) daysAhead = 0;
      else continue; // neither touchpoint applies today for this policy

      matchedCount++;
      const policyNumber = row[col('Policy Number')];

      // Due-today and advance are tracked in two completely separate
      // columns now — 'Last Reminder Sent' for the due-today touchpoint,
      // 'Advance Reminder Sent' for the advance one. Previously both
      // shared the same column, which worked correctly in practice (the
      // two dates can never land on the same calendar day), but kept
      // them tangled together with no independent record of which
      // touchpoint actually fired. Splitting them removes any chance of
      // one touchpoint's write ever being misread by the other's check,
      // and gives a clean, separate audit trail for each in the sheet.
      const trackingColName = daysAhead === 0 ? 'Last Reminder Sent' : 'Advance Reminder Sent';
      const trackingCol = col(trackingColName);
      const lastSentStr = normalizeDateCellToYmd(row[trackingCol], tz);
      if (lastSentStr === todayStr){
        // guards against the trigger firing twice same day
        skippedAlreadySent++;
        const msg = 'SKIP (already sent today, ' + trackingColName + ') — policy=' + policyNumber + ' daysAhead=' + daysAhead + ' lastSent=' + lastSentStr;
        Logger.log('[sendDailyReminders] ' + msg);
        runLog.push(msg);
        continue;
      }

      // Everything for this one row — the send itself, marking it sent,
      // and advancing its due date — is now inside a single try/catch,
      // so a failure at ANY step here only affects this one row and
      // logs exactly what broke, instead of silently ending the entire
      // run and leaving every subsequent row unprocessed.
      try{
        const sent = sendReminderEmail(row, col, daysAhead);
        if (sent){
          sentCount++;
          bumpDailyStat('STAT_SENT');
          sheet.getRange(i + 1, trackingCol + 1).setValue(todayStr);
          const msg = 'SENT (' + trackingColName + ') — policy=' + policyNumber + ' daysAhead=' + daysAhead;
          Logger.log('[sendDailyReminders] ' + msg);
          runLog.push(msg);
          // Cycling the due date forward to the next billing period only
          // makes sense once the policy has actually reached its due
          // date — the advance reminder (daysAhead > 0) is just a
          // heads-up and must never move the due date early.
          if (daysAhead === 0){
            advanceDueDate(sheet, i + 1, col, dueDate, row[col('Premium Mode')]);
          }
        } else {
          const msg = 'NOT SENT (sendReminderEmail returned false, no email on file?) — policy=' + policyNumber + ' daysAhead=' + daysAhead;
          Logger.log('[sendDailyReminders] ' + msg);
          runLog.push(msg);
        }
      }catch(err){
        failedCount++;
        // Previously this class of error was either swallowed with only
        // a counter bump, or — worse, for errors outside the narrow
        // original try/catch — could crash the entire remaining loop.
        // Now it's fully contained to this one row and fully logged.
        const msg = 'ERROR — policy=' + policyNumber + ' daysAhead=' + daysAhead + ' error=' + (err && err.message ? err.message : String(err));
        Logger.log('[sendDailyReminders] ' + msg);
        runLog.push(msg);
        bumpDailyStat('STAT_FAILED');
      }
      // Saved after EVERY matched row, not just once at the very end.
      // A plain try/finally only protects against a catchable JS
      // exception — it does NOT protect against Apps Script's own
      // platform-level hard kill when a script exceeds its maximum
      // execution time, which terminates the script outright without
      // running any cleanup code, finally block included. If that's
      // what happened on a slow run, saving progressively here means
      // whatever got through before the kill is still on record,
      // instead of the whole run vanishing without a trace.
      saveReminderRunLog(runLog);
    }

    const endMsg = 'END — matched=' + matchedCount + ' sent=' + sentCount + ' failed=' + failedCount + ' skippedAlreadySent=' + skippedAlreadySent;
    Logger.log('[sendDailyReminders] ' + endMsg);
    runLog.push(endMsg);

  }catch(err){
    // Catches anything unexpected outside the per-row protection above
    // (e.g. a problem reading the sheet itself, or in getAdvanceDays())
    // — still guarantees a log entry instead of a silent, traceless crash.
    const msg = 'FATAL ERROR — run did not complete: ' + (err && err.message ? err.message : String(err));
    Logger.log('[sendDailyReminders] ' + msg);
    runLog.push(msg);
  }finally{
    saveReminderRunLog(runLog);
  }
}

// Persists the run log to Script Properties (survives regardless of
// whether Cloud Logs captured this particular run), capped well under
// the 9KB-per-property limit by keeping only the most recent lines if
// an unusually large day's run would otherwise overflow it.
function saveReminderRunLog(runLog){
  try{
    const tz = Session.getScriptTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
    let entries = runLog;
    // Rough safety margin: each line averages well under 200 bytes, so
    // 40 lines keeps this comfortably inside the property size limit
    // even for a verbose run — keeping the START/END summary lines plus
    // the most recent detail lines if trimming is ever needed.
    if (entries.length > 40){
      entries = [entries[0]].concat(entries.slice(-39));
    }
    PropertiesService.getScriptProperties().setProperty('LAST_REMINDER_RUN_LOG', JSON.stringify({ timestamp: timestamp, entries: entries }));
  }catch(e){
    // Logging failing should never surface as if the actual send run failed.
  }
}

// Retrieves whatever sendDailyReminders() last recorded — the one
// reliable way to see what an AUTOMATIC (time-driven trigger) run
// actually did, since Cloud Logs have repeatedly not been available
// for those specifically, only for manual runs from the editor.
function getLastReminderRunLog(){
  const raw = PropertiesService.getScriptProperties().getProperty('LAST_REMINDER_RUN_LOG');
  if (!raw) return { timestamp: null, entries: [] };
  try{
    return JSON.parse(raw);
  }catch(e){
    return { timestamp: null, entries: [] };
  }
}

// daysAhead: 0 = due today, 1+ = that many days before the due date.
function sendReminderEmail(row, col, daysAhead){
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
  let subject;
  if (daysAhead === 0){
    subject = 'PREMIUM DUE REMINDER - ' + subjectDate.toUpperCase();
  } else if (daysAhead === 1){
    subject = 'PREMIUM DUE REMINDER IN 1 DAY - ' + subjectDate.toUpperCase();
  } else {
    subject = 'PREMIUM DUE REMINDER IN ' + daysAhead + ' DAYS - ' + subjectDate.toUpperCase();
  }
  const htmlBody = buildReminderEmailHtml(clientName, policyNumber, product, amount, dueDate, config, daysAhead);

  // Build options — cc, replyTo, and from-alias are all optional
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
// as). This is what hides a personal Gmail address from clients —
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
      // Alias not verified (or any other from-related failure) —
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

// Some Filipino names lead with "Ma." (short for Maria) or the spelled-
// out "Maria" as a formal prefix before the name someone actually goes
// by — e.g. "Dela Cruz, Ma. Teresa" or "Dela Cruz, Maria Cristina".
// People go by "Teresa" or "Cristina", not "Ma." or "Maria", so when
// the first word is one of these, skip it and use the next word
// instead. Falls back to the prefix itself if there's nothing after it
// (e.g. the name really is just "Maria" with no middle name).
const MARIA_STYLE_PREFIXES = ['ma', 'maria'];

function firstNameOnly(rawName){
  const name = String(rawName || '').trim();
  if (!name) return '';
  const commaIdx = name.indexOf(',');
  const rest = commaIdx !== -1 ? name.slice(commaIdx + 1).trim() : name;
  const words = rest.split(/\s+/).filter(Boolean);
  if (words.length === 0) return '';
  const firstWordNormalized = words[0].toLowerCase().replace(/\.$/, '');
  if (MARIA_STYLE_PREFIXES.includes(firstWordNormalized) && words.length > 1){
    return words[1];
  }
  return words[0];
}

function buildReminderEmailHtml(clientName, policyNumber, product, amount, dueDate, config, daysAhead){
  const tz = Session.getScriptTimeZone();
  const formattedAmount = 'PHP ' + Number(amount).toLocaleString('en-PH', { minimumFractionDigits: 2 });
  const formattedDate = Utilities.formatDate(dueDate, tz, 'MMMM d, yyyy');
  const greetingName = firstNameOnly(clientName);

  let openingLine;
  if (daysAhead === 0){
    openingLine = 'This is a friendly reminder that your premium payment is due <strong>today</strong>.';
  } else if (daysAhead === 1){
    openingLine = 'This is a friendly reminder that your premium payment is due <strong>tomorrow, ' + formattedDate + '</strong>.';
  } else {
    openingLine = 'This is a friendly reminder that your premium payment is due in <strong>' + daysAhead + ' days, on ' + formattedDate + '</strong>.';
  }
  const closingLine = daysAhead === 0
    ? 'Please settle this at your earliest convenience to keep your policy in force. If you have already made this payment, kindly disregard this reminder.'
    : 'Please prepare your payment in advance to keep your policy in force. If you have already made this payment, kindly disregard this reminder.';

  return ''
    + '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #E7DFCF;border-radius:10px;overflow:hidden;">'
    + '  <img src="cid:headerImg" alt="Header" style="width:100%;display:block;">'
    + '  <div style="padding:24px;background:#FDF8F0;color:#1C2A38;">'
    + '    <p>Hi ' + greetingName + ',</p>'
    + '    <p>' + openingLine + '</p>'
    + '    <table style="width:100%;margin:16px 0;border-collapse:collapse;font-size:14px;">'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Policy Number</td><td style="text-align:right;font-weight:700;">' + policyNumber + '</td></tr>'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Product</td><td style="text-align:right;font-weight:700;">' + product + '</td></tr>'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Amount Due</td><td style="text-align:right;font-weight:700;color:#0C447C;">' + formattedAmount + '</td></tr>'
    + '      <tr><td style="padding:8px 0;color:#6B7280;">Due Date</td><td style="text-align:right;font-weight:700;">' + formattedDate + '</td></tr>'
    + '    </table>'
    + '    <p>' + closingLine + '</p>'
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
  const daysAhead = getAdvanceDays().advanceDays;
  const tz = Session.getScriptTimeZone();
  const sampleDueDate = new Date();
  sampleDueDate.setDate(sampleDueDate.getDate() + daysAhead); // shows the due date the way an advance reminder actually would — X days out, matching your current setting
  const htmlBody = buildReminderEmailHtml('Dela Cruz, Juan Miguel', '0123456789', 'Sample Insurance Plan', 50000, sampleDueDate, config, daysAhead);
  const subjectDate = Utilities.formatDate(sampleDueDate, tz, 'MMMM d');
  const daysLabel = daysAhead === 0 ? '' : (daysAhead === 1 ? ' IN 1 DAY' : ' IN ' + daysAhead + ' DAYS');
  GmailApp.sendEmail(myEmail, 'PREVIEW, PREMIUM DUE REMINDER' + daysLabel + ' - ' + subjectDate.toUpperCase(), '', {
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
  // which some Workspace policies (e.g. Sun Life's) block outright —
  // even though it works fine for personal Gmail deployments. Using
  // contactEmail directly sidesteps that scope entirely.
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first, then try the test email again.');
  const daysAhead = getAdvanceDays().advanceDays;
  const tz = Session.getScriptTimeZone();
  const sampleDueDate = new Date();
  sampleDueDate.setDate(sampleDueDate.getDate() + daysAhead); // shows the due date the way an advance reminder actually would — X days out, matching your current setting
  const htmlBody = buildReminderEmailHtml('Dela Cruz, Juan Miguel', '0123456789', 'Sample Insurance Plan', 50000, sampleDueDate, config, daysAhead);
  const subjectDate = Utilities.formatDate(sampleDueDate, tz, 'MMMM d');
  const daysLabel = daysAhead === 0 ? '' : (daysAhead === 1 ? ' IN 1 DAY' : ' IN ' + daysAhead + ' DAYS');
  sendWithOptionalFromAlias(recipient, 'TEST, PREMIUM DUE REMINDER' + daysLabel + ' - ' + subjectDate.toUpperCase(), {
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
  const sheet = getSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayMonth = today.getMonth(), todayDay = today.getDate();
  const currentYearStr = String(today.getFullYear());
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const dob = row[col('Date of Birth')];
    if (!(dob instanceof Date)) continue;
    if (dob.getMonth() === todayMonth && dob.getDate() === todayDay){
      const lastGreetingYear = String(row[col('Last Greeting Sent (Year)')] || '');
      // Only counts as "sent" if it matches THIS year specifically —
      // a value left over from last year's birthday (or any earlier
      // one) shouldn't make today's greeting look like it already
      // went out.
      const wasSentThisYear = lastGreetingYear === currentYearStr;
      result.push({
        fullName: row[col('Full Name')],
        email: row[col('Email')],
        location: row[col('Location')],
        dobFormatted: Utilities.formatDate(dob, tz, 'MMMM d'),
        lastGreetingSent: wasSentThisYear ? lastGreetingYear : ''
      });
    }
  }
  return result;
}

function countBirthdaysOnOffset(offsetDays){
  const sheet = getSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
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
  if (!isAdvisorActive()) return; // hard stop: past the inactivity deadline
  const sheet = getSpreadsheet().getSheetByName(BIRTHDAY_SHEET_NAME);
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

  // Build options — cc and replyTo are optional
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
   POLICY ANNIVERSARY GREETER
   ------------------------------------------------------------
   Reuses the existing Dues Tracker sheet (Issued Date is already
   a column there) rather than a separate tracker — a policy
   anniversary is just "today's month/day matches Issued Date",
   the same shape as a birthday but keyed off the policy instead
   of the person. "Send Anniversary?" is its own independent
   toggle per row (like "Send Dues?"), so an advisor can turn
   anniversary greetings off for a client without touching their
   due reminders.
   ============================================================ */

// Mirrors isLapsedStatus() in the frontend exactly — that one only ever
// existed client-side, but the Policy Anniversary Greeter needs the
// same check server-side to skip lapsed policies when sending. Kept
// identical on purpose (same substring match, same case-folding) so
// "lapsed" means the same thing everywhere in this app, not two
// slightly different rules that could disagree on an edge case.
function isLapsedStatus(policyStatus){
  const s = String(policyStatus || '').toLowerCase();
  return s.includes('lapsed') || s.includes('lapse');
}

function assertConfiguredForAnniversary(config){
  const required = ['senderName','headerImageFileId','footerImageFileId'];
  const missing = required.filter(key => !config[key]);
  if (missing.length > 0){
    throw new Error(
      'Branding not set up yet. Open the app, tap "Setup", fill in ' +
      '"Your branding" (missing: ' + missing.join(', ') + '), and tap ' +
      'SAVE BRANDING before anniversary greetings can be sent.'
    );
  }
}

// e.g. 1 -> "1st", 2 -> "2nd", 3 -> "3rd", 4 -> "4th", 11-13 -> "th"
function ordinalSuffix(n){
  const rem100 = n % 100;
  if (rem100 >= 11 && rem100 <= 13) return 'th';
  switch (n % 10){
    case 1: return 'st';
    case 2: return 'nd';
    case 3: return 'rd';
    default: return 'th';
  }
}

function getPolicyAnniversariesTodayRows(){
  setupSheet();
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return [];
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const today = new Date();
  const todayMonth = today.getMonth(), todayDay = today.getDate();
  const currentYear = today.getFullYear();
  const currentYearStr = String(currentYear);
  const result = [];
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const issuedDate = row[col('Issued Date')];
    if (!(issuedDate instanceof Date)) continue;
    if (issuedDate.getMonth() !== todayMonth || issuedDate.getDate() !== todayDay) continue;
    // A lapsed policy isn't really something to celebrate — skip it
    // entirely rather than greeting someone for a policy that's no
    // longer active. Uses the same isLapsedStatus() check the rest of
    // the app already relies on as the source of truth for "lapsed",
    // not the Lapse Date column alone (a policy can carry an old Lapse
    // Date value while its current Policy Status says otherwise, e.g.
    // reinstated).
    if (isLapsedStatus(row[col('Policy Status')])) continue;
    const yearsCount = currentYear - issuedDate.getFullYear();
    if (yearsCount <= 0) continue; // issued today this same year — not an anniversary yet
    const lastSentYear = String(row[col('Last Anniversary Sent (Year)')] || '');
    const wasSentThisYear = lastSentYear === currentYearStr;
    result.push({
      policyNumber: row[col('Policy Number')],
      clientName: row[col('Client Name')],
      email: row[col('Email')],
      product: row[col('Product')],
      issuedDateFormatted: Utilities.formatDate(issuedDate, tz, 'MMMM d, yyyy'),
      years: yearsCount,
      yearsLabel: yearsCount + ordinalSuffix(yearsCount) + ' Anniversary',
      lastAnniversarySent: wasSentThisYear ? lastSentYear : ''
    });
  }
  return result;
}

function countAnniversariesOnOffset(offsetDays){
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) return 0;
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const target = new Date();
  target.setDate(target.getDate() + offsetDays);
  const targetMonth = target.getMonth(), targetDay = target.getDate();
  const targetYear = target.getFullYear();
  let count = 0;
  const statusCol = col('Policy Status');
  for (let i = 1; i < data.length; i++){
    const issuedDate = data[i][col('Issued Date')];
    if (!(issuedDate instanceof Date)) continue;
    if (isLapsedStatus(data[i][statusCol])) continue; // keep this count consistent with the actual list/send logic
    if (issuedDate.getMonth() === targetMonth && issuedDate.getDate() === targetDay && issuedDate.getFullYear() < targetYear) count++;
  }
  return count;
}

function getAnniversaryDailyStats(){
  return {
    sent: getDailyStat('ANNIV_STAT_SENT'),
    failed: getDailyStat('ANNIV_STAT_FAILED'),
    anniversariesToday: countAnniversariesOnOffset(0),
    anniversariesTomorrow: countAnniversariesOnOffset(1)
  };
}

function sendDailyAnniversaryGreetings(){
  // Same reliability hardening as sendDailyReminders — per-row isolation
  // (one bad row can't crash the rest of the loop) and a progressive,
  // persisted run log (survives even if Apps Script's own execution
  // timeout kills the script mid-run, which try/finally alone can't
  // protect against). Anniversary greetings never had either of these
  // until now, despite being exposed to the exact same failure classes.
  const runLog = [];
  let sentCount = 0, failedCount = 0, matchedCount = 0;

  try{
    if (!getAnniversaryAutoSendStatus().enabled){ runLog.push('EXIT: auto-send is OFF'); return; }
    if (!isAdvisorActive()){ runLog.push('EXIT: advisor inactive (hard-stop)'); return; }
    setupSheet();
    const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
    if (!sheet){ runLog.push('EXIT: Dues Tracker sheet not found'); return; }
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);
    const today = new Date();
    const todayMonth = today.getMonth(), todayDay = today.getDate();
    const currentYear = today.getFullYear();
    const currentYearStr = String(currentYear);

    const startMsg = 'START — today=' + Utilities.formatDate(today, Session.getScriptTimeZone(), 'yyyy-MM-dd') + ' totalRows=' + (data.length - 1);
    Logger.log('[sendDailyAnniversaryGreetings] ' + startMsg);
    runLog.push(startMsg);
    saveAnniversaryRunLog(runLog);

    for (let i = 1; i < data.length; i++){
      const row = data[i];
      const sendAnniv = row[col('Send Anniversary?')];
      if (sendAnniv === false || sendAnniv === 'FALSE' || sendAnniv === 0 || sendAnniv === '0') continue;
      const issuedDate = row[col('Issued Date')];
      if (!(issuedDate instanceof Date)) continue;
      if (issuedDate.getMonth() !== todayMonth || issuedDate.getDate() !== todayDay) continue;
      if (isLapsedStatus(row[col('Policy Status')])) continue; // same rule as the display list — no greeting for a lapsed policy
      const yearsCount = currentYear - issuedDate.getFullYear();
      if (yearsCount <= 0) continue;
      const lastSentYear = String(row[col('Last Anniversary Sent (Year)')] || '');
      if (lastSentYear === currentYearStr) continue;

      matchedCount++;
      const policyNumber = row[col('Policy Number')];

      // Sending AND marking it sent are now both inside one try/catch —
      // previously only the send itself was protected, so a failure
      // writing the sheet afterward could crash every remaining row
      // in the loop, exactly the bug class that caused the dues
      // reminder investigation this whole week.
      try{
        const sent = sendAnniversaryEmail(row, col, yearsCount);
        if (sent){
          sentCount++;
          bumpDailyStat('ANNIV_STAT_SENT');
          sheet.getRange(i + 1, col('Last Anniversary Sent (Year)') + 1).setValue(currentYearStr);
          const msg = 'SENT — policy=' + policyNumber + ' years=' + yearsCount;
          Logger.log('[sendDailyAnniversaryGreetings] ' + msg);
          runLog.push(msg);
        } else {
          const msg = 'NOT SENT (no email on file?) — policy=' + policyNumber;
          Logger.log('[sendDailyAnniversaryGreetings] ' + msg);
          runLog.push(msg);
        }
      }catch(err){
        failedCount++;
        bumpDailyStat('ANNIV_STAT_FAILED');
        const msg = 'ERROR — policy=' + policyNumber + ' error=' + (err && err.message ? err.message : String(err));
        Logger.log('[sendDailyAnniversaryGreetings] ' + msg);
        runLog.push(msg);
      }
      // Saved after every matched row, not just once at the end — same
      // reason as sendDailyReminders: a hard platform timeout kill
      // bypasses try/finally entirely, so whatever ran before the kill
      // needs to already be on record.
      saveAnniversaryRunLog(runLog);
    }

    const endMsg = 'END — matched=' + matchedCount + ' sent=' + sentCount + ' failed=' + failedCount;
    Logger.log('[sendDailyAnniversaryGreetings] ' + endMsg);
    runLog.push(endMsg);

  }catch(err){
    const msg = 'FATAL ERROR — run did not complete: ' + (err && err.message ? err.message : String(err));
    Logger.log('[sendDailyAnniversaryGreetings] ' + msg);
    runLog.push(msg);
  }finally{
    saveAnniversaryRunLog(runLog);
  }
}

// Same persistence pattern as saveReminderRunLog/getLastReminderRunLog,
// under a separate Script Property so the two features' diagnostics
// never overwrite each other.
function saveAnniversaryRunLog(runLog){
  try{
    const tz = Session.getScriptTimeZone();
    const timestamp = Utilities.formatDate(new Date(), tz, "yyyy-MM-dd'T'HH:mm:ss");
    let entries = runLog;
    if (entries.length > 40){
      entries = [entries[0]].concat(entries.slice(-39));
    }
    PropertiesService.getScriptProperties().setProperty('LAST_ANNIVERSARY_RUN_LOG', JSON.stringify({ timestamp: timestamp, entries: entries }));
  }catch(e){
    // Logging failing should never surface as if the actual send run failed.
  }
}

function getLastAnniversaryRunLog(){
  const raw = PropertiesService.getScriptProperties().getProperty('LAST_ANNIVERSARY_RUN_LOG');
  if (!raw) return { timestamp: null, entries: [] };
  try{
    return JSON.parse(raw);
  }catch(e){
    return { timestamp: null, entries: [] };
  }
}

function sendAnniversaryEmail(row, col, years){
  const email = row[col('Email')];
  if (!email) return false;
  const config = getBrandConfig();
  assertConfiguredForAnniversary(config);
  const clientName = row[col('Client Name')];
  const subject = 'HAPPY POLICY ANNIVERSARY!';
  const htmlBody = buildAnniversaryEmailHtml(clientName, years, config);

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

// Deliberately no "Pay Online" button here — an anniversary is a
// relationship touchpoint, not a billing moment. The two CTAs instead
// invite a review conversation (reviewLink) and general contact
// (connectLink), each shown only if that link is actually configured.
function buildAnniversaryEmailHtml(clientName, years, config){
  const greetingName = firstNameOnly(clientName);
  const yearsLabel = years + ordinalSuffix(years);

  const reviewBlock = config.reviewLink
    ? ('    <div style="text-align:center;margin:18px 0 6px;">'
      + '      <a href="' + config.reviewLink + '" style="display:inline-block;background:#0C447C;color:#FFFFFF;text-decoration:none;padding:14px 30px;border-radius:8px;font-weight:700;font-size:14px;letter-spacing:.5px;">SCHEDULE A 15-MIN POLICY REVIEW</a>'
      + '    </div>')
    : '';
  const connectBlock = config.connectLink
    ? ('    <div style="text-align:center;margin:10px 0 6px;">'
      + '      <a href="' + config.connectLink + '" style="display:inline-block;background:#C99A3B;color:#FFFFFF;text-decoration:none;padding:14px 30px;border-radius:8px;font-weight:700;font-size:14px;letter-spacing:.5px;">CONNECT WITH ME</a>'
      + '    </div>')
    : '';

  return ''
    + '<div style="font-family:Arial,sans-serif;max-width:480px;margin:0 auto;border:1px solid #E7DFCF;border-radius:10px;overflow:hidden;">'
    + '  <img src="cid:headerImg" alt="Header" style="width:100%;display:block;">'
    + '  <div style="padding:24px;background:#FDF8F0;color:#1C2A38;text-align:center;">'
    + '    <p style="font-size:18px;font-weight:700;color:#0C447C;margin:0 0 10px;">Happy ' + yearsLabel + ' Policy Anniversary, ' + greetingName + '! &#127881;</p>'
    + '    <p style="font-size:14px;">' + years + ' year' + (years === 1 ? '' : 's') + ' ago, you took a step toward protecting what matters most. I\u2019ve been honored to walk that journey with you since, and I\u2019m grateful for your continued trust.</p>'
    + '    <p style="font-size:14px;">A policy anniversary is also a great moment to check that your coverage still fits your life today. If you\u2019d like, let\u2019s take 15 minutes to go over it together.</p>'
    + reviewBlock
    + connectBlock
    + '    <p style="margin-top:20px;text-align:left;">Warm regards,</p>'
    + '  </div>'
    + '  <img src="cid:footerImg" alt="Footer" style="width:100%;display:block;">'
    + '</div>';
}

function previewAnniversaryEmail(){
  const myEmail = Session.getActiveUser().getEmail();
  const config = getBrandConfig();
  assertConfiguredForAnniversary(config);
  const htmlBody = buildAnniversaryEmailHtml('Dela Cruz, Juan Miguel', 5, config);
  GmailApp.sendEmail(myEmail, 'PREVIEW – HAPPY POLICY ANNIVERSARY', '', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  });
}

// Test email: sends to contactEmail if set, otherwise falls back to
// the active user's email (works in the Apps Script editor context).
function sendAnniversaryTestEmailToSelf(){
  const config = getBrandConfig();
  assertConfiguredForAnniversary(config);
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first, then try the test email again.');
  const htmlBody = buildAnniversaryEmailHtml('Dela Cruz, Juan Miguel', 5, config);
  sendWithOptionalFromAlias(recipient, 'TEST – HAPPY POLICY ANNIVERSARY', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  }, config.contactEmail);
  return { success: true, sentTo: recipient };
}

function getAnniversaryAutoSendStatus(){
  const val = PropertiesService.getScriptProperties().getProperty('ANNIV_AUTO_SEND_ENABLED');
  return { enabled: val === null ? true : val === '1' };
}

function setAnniversaryAutoSendStatus(enabled){
  PropertiesService.getScriptProperties().setProperty('ANNIV_AUTO_SEND_ENABLED', enabled ? '1' : '0');
}

function createAnniversaryDailyTrigger(hour){
  hour = hour !== undefined ? hour : getSendHour().hour;
  ScriptApp.getProjectTriggers().forEach(t => {
    if (t.getHandlerFunction() === 'sendDailyAnniversaryGreetings') ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('sendDailyAnniversaryGreetings')
    .timeBased()
    .everyDays(1)
    .atHour(hour)
    .create();
}

// Cheap, idempotent self-install for advisors who were already using
// the app before the Policy Anniversary Greeter existed and haven't
// touched their Send Hour setting since — createAnniversaryDailyTrigger()
// above always deletes-and-recreates, which is fine for the rare
// "advisor changed their send hour" case it's built for, but far too
// expensive to call on every single upload. Same lock protection as
// ensureDailyReminderTriggerExists — this also runs on every request
// now, and was exposed to the identical concurrent-duplicate race.
// Also now version-aware via _rebuildTriggerIfStale.
function ensureAnniversaryDailyTriggerExists(){
  const lock = LockService.getScriptLock();
  try{
    lock.waitLock(3000);
  }catch(e){
    return;
  }
  try{
    _rebuildTriggerIfStale('sendDailyAnniversaryGreetings', createAnniversaryDailyTrigger);
  }finally{
    lock.releaseLock();
  }
}

function setAnniversaryPreference(policyNumber, enabled){
  setupSheet();
  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const policyCol = headers.indexOf('Policy Number');
  const sendCol = headers.indexOf('Send Anniversary?');
  for (let i = 1; i < data.length; i++){
    if (String(data[i][policyCol]) === String(policyNumber)){
      sheet.getRange(i + 1, sendCol + 1).setValue(enabled);
      return { success: true };
    }
  }
  return { success: false, error: 'Policy not found' };
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
      'Go to Settings → Branding Studio to add them, or turn the template toggle off to send without it.'
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
      'Go to Settings → Branding Studio and re-upload your header and footer photos — they\u2019ll now be compressed automatically to a safe size. ' +
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
  const sentEmails = []; // used to log this send for the "already sent this subject" exclusion check
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
      failureReasons.push({ email: r.email, reason: 'Not attempted — daily sending quota was already exhausted earlier in this send.' });
      return;
    }
    try{
      const firstName = firstNameOnly(r.clientName) || 'there';
      // Bold + black so the personalized greeting stands out regardless
      // of whatever color the surrounding message text uses.
      const styledFirstName = '<span style="font-weight:700;color:#000000;">' + firstName + '</span>';
      // Case-insensitive match: {FirstName}, {firstName}, {FIRSTNAME},
      // {firstname} all work — typing the exact capitalization
      // correctly shouldn't be a requirement for personalization to
      // actually apply.
      const personalizedBody = wrappedBody.replace(/\{firstname\}/gi, styledFirstName);

      const options = {
        htmlBody: personalizedBody,
        name: config.senderName,
      };
      // Merges the header/footer template images (when enabled) with
      // any body images that were converted from data URLs to cid:
      // references above — both need to be present in inlineImages
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
      sentEmails.push(r.email);
    }catch(err){
      failed++;
      failedEmails.push(r.email);
      const translatedReason = toEnglishErrorMessage(err.message || String(err));
      // If this is a size-related failure, attach the exact numbers our
      // own pre-check calculated for this send — the processed body
      // size, body image total, template size, and attachment total —
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
      // mid-batch. Checked against the RAW message (not the translated
      // one) since this only needs to detect the pattern, not display
      // it — but that raw message's language depends on the deploying
      // Google account's locale, which can be Chinese (調用次數過多 =
      // "too many invocations") just as easily as English or Tagalog.
      if (/quota|invoked too many times|masyadong madaming beses.*araw|調用次數過多|调用次数过多/i.test(err.message || '')){
        quotaExhausted = true;
      }
    }
  });

  if (sentEmails.length > 0){
    try{ recordBroadcastSentTo(subject, sentEmails); }
    catch(e){ /* logging failure should never block the actual send result from being returned */ }
  }

  return { sent: sent, failed: failed, failedEmails: failedEmails, failureReasons: failureReasons, total: rows.length, quotaExhausted: quotaExhausted };
}

/* ============================================================
   BROADCAST LOG — "already sent this subject" exclusion
   ------------------------------------------------------------
   Separate "Broadcast Log" tab, auto-created on first successful
   send. Columns: Timestamp | Subject | Email — one row per
   recipient per successful send (both immediate and scheduled
   broadcasts funnel through sendBroadcastEmailBatch, so both are
   covered by this same log automatically). Used by the recipient
   picker to grey out anyone who already received THIS exact
   subject, so re-running a campaign can't accidentally double-send
   to someone who already got it — they still show in the list
   (so it's obvious who's excluded and why), just can't be
   selected.

   Matching is by exact subject text after trimming and
   case-folding — "Q3 Newsletter" and "q3 newsletter " are treated
   as the same campaign, but a genuinely different subject is
   never treated as a duplicate, even for the same recipients.
   ============================================================ */
const BROADCAST_LOG_TAB_NAME = 'Broadcast Log';
const BROADCAST_LOG_HEADERS = ['Timestamp', 'Subject', 'Email'];

function getBroadcastLogTab(){
  const ss = getSpreadsheet();
  let sheet = ss.getSheetByName(BROADCAST_LOG_TAB_NAME);
  if (!sheet){
    sheet = ss.insertSheet(BROADCAST_LOG_TAB_NAME);
    sheet.appendRow(BROADCAST_LOG_HEADERS);
    sheet.setFrozenRows(1);
  }
  return sheet;
}

function normalizeSubjectForMatch(subject){
  return String(subject || '').trim().toLowerCase();
}

function recordBroadcastSentTo(subject, emails){
  if (!emails || emails.length === 0) return;
  const sheet = getBroadcastLogTab();
  const now = new Date();
  // One writeup of all rows at once (setValues) rather than an
  // appendRow() per recipient — for a broadcast to hundreds of
  // people, hundreds of individual writes would be needlessly slow
  // and risks pushing this past the Web App's execution time limit.
  const startRow = sheet.getLastRow() + 1;
  const rows = emails.map(email => [now, subject, email]);
  sheet.getRange(startRow, 1, rows.length, 3).setValues(rows);
}

/**
 * Returns every distinct email that has already received a broadcast
 * with this exact subject (trimmed, case-insensitive match) — used by
 * the recipient picker to grey those rows out. Scans the whole log
 * rather than filtering as it's read since Sheets doesn't support a
 * server-side query here; for the realistic size of this log (one row
 * per recipient per send) this stays fast well beyond normal usage.
 */
function getSentEmailsForSubject(subject){
  const target = normalizeSubjectForMatch(subject);
  if (!target) return jsonResponse({ emails: [] });

  const sheet = getBroadcastLogTab();
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return jsonResponse({ emails: [] });

  const data = sheet.getRange(2, 1, lastRow - 1, 3).getValues();
  const matched = new Set();
  for (let i = 0; i < data.length; i++){
    if (normalizeSubjectForMatch(data[i][1]) === target){
      matched.add(String(data[i][2]).trim().toLowerCase());
    }
  }
  return jsonResponse({ emails: Array.from(matched) });
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

  if (!isAdvisorActive()){
    // hard stop: past the inactivity deadline. Mark the row instead of
    // silently dropping it so it's visible in the sheet why nothing went out.
    sheet.getRange(rowNum, col('Status') + 1).setValue('skipped');
    sheet.getRange(rowNum, col('Error') + 1).setValue('Advisor inactive (no recent upload) — broadcast skipped');
    return;
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
      failedCount: Number(row[col('FailedCount')]) || 0,
      // Lets a broadcast with failures be reloaded straight back into the
      // composer (subject, body, attachments, template flag, original
      // recipient list) instead of the advisor having to rebuild it from
      // scratch — the frontend cross-checks the recipient list against
      // the Broadcast Log to automatically exclude anyone who already
      // received it, same mechanism drafts already use.
      payload: JSON.parse(row[col('PayloadJSON')] || '{}')
    });
  }
  result.sort((a, b) => new Date(b.createdAt) - new Date(a.createdAt));
  return result;
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
        return { success: false, error: 'Only completed (sent or failed) broadcasts can be removed this way — cancel it instead if it\u2019s still scheduled or sending.' };
      }
      sheet.deleteRow(i + 1);
      return { success: true };
    }
  }
  return { success: false, error: 'Broadcast not found.' };
}

// Called after an "Edit & Resend" completes — leaves a note on the
// ORIGINAL broadcast row saying the failed recipients were successfully
// resent, without touching its original SentCount/FailedCount. Those
// numbers stay exactly as they were as an accurate historical record of
// that specific attempt; this just appends context so a glance at
// Completed Broadcasts doesn't wrongly suggest 21 people never actually
// got the message, when the Broadcast Log shows they did (just via a
// second, separate send rather than a correction to the first).
function annotateResendOnSchedule(scheduleId, resendSentCount, resendFailedCount){
  const sheet = setupScheduleSheet();
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Schedule ID')]) === String(scheduleId)){
      const tz = Session.getScriptTimeZone();
      const todayStr = Utilities.formatDate(new Date(), tz, 'MMMM d, yyyy');
      const note = 'Resent on ' + todayStr + ' via Edit & Resend \u2014 ' + resendSentCount + ' sent' +
        (resendFailedCount > 0 ? ', ' + resendFailedCount + ' still failed' : '') +
        ' (see Broadcast Log for exact recipients).';
      const existingError = String(data[i][col('Error')] || '');
      const newError = existingError ? existingError + ' | ' + note : note;
      sheet.getRange(i + 1, col('Error') + 1).setValue(newError);
      return { success: true };
    }
  }
  return { success: false, error: 'Original schedule not found (it may have been deleted).' };
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
    // ── Email size ──────────────────────────────────────────────────────
    { match: /Limitasyon.*Laki ng Body|Body.*[Ll]imit exceeded|Nalagpasan.*[Ll]imitasyon|message too large|attachment.*too large|sumasobra.*laki/i,
      english: 'Email is too large to send — remove an inline image or attachment and try again.' },

    // ── Invalid email address ───────────────────────────────────────────
    { match: /Invalid email|[Mm]ali ang email|hindi wasto ang email|invalid.*recipient|di.*wastong.*email/i,
      english: 'Invalid email address.' },

    // ── Daily quota / sending limit ─────────────────────────────────────
    { match: /quota|limitasyon.*araw|daily.*limit|invoked too many times|masyadong madaming beses.*araw|naabot.*limitasyon|Service.*invoked.*many|調用次數過多|调用次数过多|單日.*過多|单日.*过多/i,
      english: 'Daily sending limit reached for this Google account (personal Gmail accounts get 100 emails/day; Google Workspace accounts get up to 1,500/day) — try again after the quota resets, or send the rest tomorrow.' },

    // ── Rate limiting ───────────────────────────────────────────────────
    { match: /rate limit|masyadong marami|too many.*request|too fast|mabilis.*nang/i,
      english: 'Sending too fast — please wait a moment and try again.' },

    // ── Recipient address ───────────────────────────────────────────────
    { match: /Recipient address required|kinakailangan ang address|walang.*tatanggap|tatanggap.*wala/i,
      english: 'Recipient address is missing or invalid.' },

    // ── Permission / authorization ──────────────────────────────────────
    { match: /permiso|pahintulot|walang.*pahintulot|You do not have permission|hindi.*pinahintulutan|access.*denied|hindi.*ma-access/i,
      english: 'Permission denied — make sure the script is authorized and deployed with Execute as: Me.' },

    // ── Spreadsheet not found ───────────────────────────────────────────
    { match: /Spreadsheet.*not found|hindi.*mahanap.*spreadsheet|walang.*spreadsheet|sheet.*not found|hindi.*sheet/i,
      english: 'Spreadsheet not found — check that the Sheet ID is correct and the script has access to it.' },

    // ── Script not authorized ───────────────────────────────────────────
    { match: /Script.*not authorized|hindi.*awtorisado|awtorisasyon.*kailangan|Authorization.*required|kailangang.*payagan/i,
      english: 'Script not authorized — open the Apps Script editor and run any function once to complete authorization.' },

    // ── Drive storage full ──────────────────────────────────────────────
    { match: /Drive.*storage|storage.*full|puno.*na.*storage|Drive.*puno/i,
      english: 'Google Drive storage is full — free up space in Google Drive and try again.' },

    // ── GmailApp disabled / not enabled ────────────────────────────────
    { match: /GmailApp.*disabled|hindi.*pinagana.*Gmail|Gmail.*not enabled/i,
      english: 'Gmail service is not enabled for this script — add it under Services in the Apps Script editor.' },

    // ── Execution time limit ────────────────────────────────────────────
    { match: /time.*limit|execution.*exceeded|naabot.*oras|lumampas.*oras|napatagal/i,
      english: 'Script execution timed out — the batch size may be too large. Try again; it will pick up from where it stopped.' },

    // ── Network / connection errors ─────────────────────────────────────
    { match: /network.*error|koneksyon.*error|walang.*koneksyon|connection.*failed|hindi.*kumokonekta/i,
      english: 'Network connection error — check your internet connection and try again.' },

    // ── Blocked recipient (corporate mail gateway) ──────────────────────
    { match: /blocked|na-block|message.*rejected|tinanggihan.*mensahe|550|554/i,
      english: 'Message blocked — the recipient\'s email server rejected it. This is usually a corporate mail filter. Nothing you can do on your end.' },

    // ── Catch-all: any remaining Tagalog text ───────────────────────────
    // Detects common Tagalog words and replaces the whole message so
    // raw untranslated Tagalog never reaches the advisor's screen.
    { match: /\b(ang|ng|na|sa|ay|hindi|wala|para|nang|ito|kami|mo|niya|sila|kayo|ako|mga|kung|pero|at|o|may|magpadala|tatanggap|mensahe|error|problema)\b/i,
      english: 'An error occurred while processing your request. Please try again. If the problem persists, check your Apps Script deployment settings.' },
  ];
  for (const p of patterns){
    if (p.match.test(msg)) return p.english;
  }
  // Unrecognized message in any language: return as-is so nothing
  // gets silently swallowed if a new/unseen error shows up.
  return msg;
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
  GmailApp.sendEmail(myEmail, 'PREVIEW – HAPPY BIRTHDAY GREETING', '', {
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
  // Same scope restriction as the dues test email — use contactEmail
  // directly instead of Session.getEffectiveUser().
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first, then try the test email again.');
  const htmlBody = buildBirthdayEmailHtml('Juan Miguel Dela Cruz', config);
  sendWithOptionalFromAlias(recipient, 'TEST – HAPPY BIRTHDAY GREETING', {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  }, config.contactEmail);
  return { success: true, sentTo: recipient };
}
function checkEmailQuota() {
  const remaining = MailApp.getRemainingDailyQuota();

  Logger.log("Remaining recipient quota today: " + remaining);
}
// TEMPORARY DIAGNOSTIC — paste this into Code.gs, run it directly from
// the Apps Script editor (select this function in the dropdown, click
// Run), then check View -> Logs (or the Execution log) for the output.
// Shows exactly what sendDailyReminders() sees for a specific policy —
// the raw Due Date cell, its computed date string, today's advance
// target string, and whether they actually match — so we can see the
// real values instead of guessing. Safe to delete after we're done.
function diagnoseAdvanceDayMatch(){
  const policyNumbersToCheck = ['0817593993', '0850328560']; // Lobo, Villena

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');

  const advanceDays = getAdvanceDays().advanceDays;
  const advanceTargetDate = new Date();
  advanceTargetDate.setDate(advanceTargetDate.getDate() + advanceDays);
  const advanceTargetStr = Utilities.formatDate(advanceTargetDate, tz, 'yyyy-MM-dd');

  Logger.log('=== DIAGNOSTIC ===');
  Logger.log('Today: ' + todayStr + ' | Advance Days setting: ' + advanceDays + ' | Advance target date: ' + advanceTargetStr);
  Logger.log('');

  let foundAny = false;
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const policyNumber = String(row[col('Policy Number')] || '').trim();
    if (!policyNumbersToCheck.includes(policyNumber)) continue;
    foundAny = true;

    const sendDues = row[col('Send Dues?')];
    const policyStatus = row[col('Policy Status')];
    const dueDateRaw = row[col('Due Date')];
    const lastReminderRaw = row[col('Last Reminder Sent')];
    const advanceReminderColIdx = col('Advance Reminder Sent');
    const advanceReminderRaw = advanceReminderColIdx !== -1 ? row[advanceReminderColIdx] : '(column not found)';

    Logger.log('--- Policy ' + policyNumber + ' (row ' + (i + 1) + ') ---');
    Logger.log('Client Name: ' + row[col('Client Name')]);
    Logger.log('Send Dues? raw value: ' + JSON.stringify(sendDues) + ' (type: ' + typeof sendDues + ')');
    Logger.log('Policy Status: ' + JSON.stringify(policyStatus));
    Logger.log('isLapsedStatus() result: ' + isLapsedStatus(policyStatus));
    Logger.log('Due Date raw value: ' + JSON.stringify(dueDateRaw) + ' (type: ' + typeof dueDateRaw + ', instanceof Date: ' + (dueDateRaw instanceof Date) + ')');
    if (dueDateRaw instanceof Date){
      const dueDateStr = Utilities.formatDate(dueDateRaw, tz, 'yyyy-MM-dd');
      Logger.log('Due Date formatted: ' + dueDateStr);
      Logger.log('Matches advance target (' + advanceTargetStr + ')? ' + (dueDateStr === advanceTargetStr));
      Logger.log('Matches today (' + todayStr + ')? ' + (dueDateStr === todayStr));
    }
    Logger.log('Last Reminder Sent (due-today column) raw value: ' + JSON.stringify(lastReminderRaw));
    Logger.log('Last Reminder Sent normalized: ' + normalizeDateCellToYmd(lastReminderRaw, tz));
    Logger.log('Advance Reminder Sent (advance-day column) raw value: ' + JSON.stringify(advanceReminderRaw));
    Logger.log('Advance Reminder Sent normalized: ' + (advanceReminderColIdx !== -1 ? normalizeDateCellToYmd(advanceReminderRaw, tz) : '(column not found)'));
    Logger.log('');
  }

  if (!foundAny){
    Logger.log('NONE of the specified policy numbers were found in the sheet at all — check for typos or extra whitespace in the Policy Number column.');
  }
  Logger.log('=== END DIAGNOSTIC ===');
}

/* ============================================================
   COMPREHENSIVE ADVANCE REMINDER DIAGNOSTIC SUITE
   ------------------------------------------------------------
   Four new, purely diagnostic/test functions — none of these send
   real reminder emails or modify Advance Reminder Sent / Last Reminder
   Sent / Due Date. Built to make it possible to conclusively answer
   "why didn't this policy get its advance reminder" without guessing,
   run directly from the Apps Script editor (select the function,
   click Run, then View -> Logs).
   ============================================================ */

// Run directly from the editor. Analyzes every row in the Dues
// Tracker against the current Advance Days setting and prints a full
// eligibility breakdown for every policy matching today's advance
// target date — no emails sent, nothing written to the sheet.
function diagnoseAdvanceReminders(){
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const advanceDays = getAdvanceDays().advanceDays;
  const advanceTargetDate = new Date();
  advanceTargetDate.setDate(advanceTargetDate.getDate() + advanceDays);
  const advanceTargetStr = Utilities.formatDate(advanceTargetDate, tz, 'yyyy-MM-dd');
  const autoSend = getAutoSendStatus().enabled;
  const advisorActive = isAdvisorActive();
  const sendHour = getSendHour().hour;
  const triggerFound = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sendDailyReminders');

  Logger.log('========================================');
  Logger.log('CLIENT PULSE ADVANCE REMINDER DIAGNOSTIC');
  Logger.log('========================================');
  Logger.log('');
  Logger.log('Today:');
  Logger.log(todayStr);
  Logger.log('');
  Logger.log('Script Timezone:');
  Logger.log(tz);
  Logger.log('');
  Logger.log('Advance Days:');
  Logger.log(String(advanceDays));
  Logger.log('');
  Logger.log('Advance Target:');
  Logger.log(advanceTargetStr);
  Logger.log('');
  Logger.log('Auto Send:');
  Logger.log(autoSend ? 'ON' : 'OFF');
  Logger.log('');
  Logger.log('Advisor Active:');
  Logger.log(advisorActive ? 'YES' : 'NO (hard-stopped)');
  Logger.log('');
  Logger.log('Daily Trigger:');
  Logger.log(triggerFound ? 'FOUND' : 'MISSING');
  Logger.log('');
  Logger.log('Send Hour:');
  Logger.log(String(sendHour));
  Logger.log('');
  Logger.log('Remaining Email Quota:');
  Logger.log(String(MailApp.getRemainingDailyQuota()));
  Logger.log('========================================');
  Logger.log('');

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet){
    Logger.log('SHEET NOT FOUND — cannot continue diagnostic.');
    return;
  }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);
  const advanceReminderColIdx = col('Advance Reminder Sent');

  let totalRows = 0, advanceTargetMatches = 0, eligibleToSend = 0, alreadySent = 0,
      missingEmail = 0, sendDuesDisabled = 0, lapsed = 0, invalidDueDate = 0;

  for (let i = 1; i < data.length; i++){
    const row = data[i];
    const policyNumber = row[col('Policy Number')];
    if (!policyNumber) continue;
    totalRows++;

    const dueDateRaw = row[col('Due Date')];
    if (!(dueDateRaw instanceof Date)){
      invalidDueDate++;
      continue;
    }
    const dueDateStr = Utilities.formatDate(dueDateRaw, tz, 'yyyy-MM-dd');
    if (dueDateStr !== advanceTargetStr) continue; // only reporting advance-target matches in detail, per spec

    advanceTargetMatches++;
    const clientName = row[col('Client Name')];
    const sendDues = row[col('Send Dues?')];
    const sendDuesBool = !(sendDues === false || sendDues === 'FALSE' || sendDues === 0 || sendDues === '0');
    const policyStatus = row[col('Policy Status')];
    const isLapsed = isLapsedStatus(policyStatus);
    const email = row[col('Email')];
    const advanceReminderRaw = advanceReminderColIdx !== -1 ? row[advanceReminderColIdx] : '';
    const advanceReminderSentStr = normalizeDateCellToYmd(advanceReminderRaw, tz);
    const wasAlreadySent = advanceReminderSentStr === todayStr;

    let expectedAction = 'SHOULD SEND';
    if (!sendDuesBool){ expectedAction = 'SKIP (Send Dues? is off)'; sendDuesDisabled++; }
    else if (isLapsed){ expectedAction = 'SKIP (lapsed)'; lapsed++; }
    else if (!email){ expectedAction = 'SKIP (no email on file)'; missingEmail++; }
    else if (wasAlreadySent){ expectedAction = 'SKIP (already sent today)'; alreadySent++; }
    else { eligibleToSend++; }

    Logger.log('Policy:');
    Logger.log(String(policyNumber));
    Logger.log('');
    Logger.log('Client:');
    Logger.log(String(clientName));
    Logger.log('');
    Logger.log('Due Date:');
    Logger.log(dueDateStr);
    Logger.log('');
    Logger.log('Send Dues?:');
    Logger.log(sendDuesBool ? 'TRUE' : 'FALSE');
    Logger.log('');
    Logger.log('Policy Status:');
    Logger.log(String(policyStatus) + (isLapsed ? ' (LAPSED)' : ''));
    Logger.log('');
    Logger.log('Email:');
    Logger.log(email ? String(email) : '(BLANK)');
    Logger.log('');
    Logger.log('Advance Reminder Sent:');
    Logger.log(advanceReminderSentStr || 'BLANK');
    Logger.log('');
    Logger.log('Matches Advance Target:');
    Logger.log('YES');
    Logger.log('');
    Logger.log('Expected Action:');
    Logger.log(expectedAction);
    Logger.log('----------------------------------------');
  }

  Logger.log('');
  Logger.log('========================================');
  Logger.log('Total Rows:');
  Logger.log(String(totalRows));
  Logger.log('');
  Logger.log('Advance Target Matches:');
  Logger.log(String(advanceTargetMatches));
  Logger.log('');
  Logger.log('Eligible To Send:');
  Logger.log(String(eligibleToSend));
  Logger.log('');
  Logger.log('Already Sent:');
  Logger.log(String(alreadySent));
  Logger.log('');
  Logger.log('Missing Email:');
  Logger.log(String(missingEmail));
  Logger.log('');
  Logger.log('Send Dues Disabled:');
  Logger.log(String(sendDuesDisabled));
  Logger.log('');
  Logger.log('Lapsed:');
  Logger.log(String(lapsed));
  Logger.log('');
  Logger.log('Invalid Due Date:');
  Logger.log(String(invalidDueDate));
  Logger.log('========================================');
}

// Diagnostic-only check for ONE specific policy — does NOT send any
// email and does NOT touch Due Date, Advance Reminder Sent, or Last
// Reminder Sent. Run from the editor after setting policyNumber below,
// or call it programmatically with a policy number string.
function testAdvanceReminderForPolicy(policyNumber){
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const advanceDays = getAdvanceDays().advanceDays;
  const advanceTargetDate = new Date();
  advanceTargetDate.setDate(advanceTargetDate.getDate() + advanceDays);
  const advanceTargetStr = Utilities.formatDate(advanceTargetDate, tz, 'yyyy-MM-dd');

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet){ Logger.log('Dues Tracker sheet not found.'); return; }
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  let found = false;
  for (let i = 1; i < data.length; i++){
    const row = data[i];
    if (String(row[col('Policy Number')]).trim() !== String(policyNumber).trim()) continue;
    found = true;

    const dueDateRaw = row[col('Due Date')];
    const dueDateStr = (dueDateRaw instanceof Date) ? Utilities.formatDate(dueDateRaw, tz, 'yyyy-MM-dd') : null;
    const isAdvanceMatch = dueDateStr === advanceTargetStr;
    const isDueTodayMatch = dueDateStr === todayStr;
    const sendDues = row[col('Send Dues?')];
    const sendDuesBool = !(sendDues === false || sendDues === 'FALSE' || sendDues === 0 || sendDues === '0');
    const policyStatus = row[col('Policy Status')];
    const isLapsed = isLapsedStatus(policyStatus);
    const email = row[col('Email')];
    const advanceReminderColIdx = col('Advance Reminder Sent');
    const advanceReminderSentStr = advanceReminderColIdx !== -1 ? normalizeDateCellToYmd(row[advanceReminderColIdx], tz) : '';
    const lastReminderSentStr = normalizeDateCellToYmd(row[col('Last Reminder Sent')], tz);

    Logger.log('=== testAdvanceReminderForPolicy: ' + policyNumber + ' ===');
    Logger.log('Today: ' + todayStr + ' | Advance Days: ' + advanceDays + ' | Advance Target: ' + advanceTargetStr);
    Logger.log('Client Name: ' + row[col('Client Name')]);
    Logger.log('Due Date: ' + (dueDateStr || 'INVALID/NOT A DATE'));
    Logger.log('Is Advance-Target Match: ' + isAdvanceMatch);
    Logger.log('Is Due-Today Match: ' + isDueTodayMatch);
    Logger.log('Send Dues?: ' + sendDuesBool);
    Logger.log('Policy Status: ' + policyStatus + (isLapsed ? ' (LAPSED)' : ''));
    Logger.log('Email: ' + (email || '(BLANK)'));
    Logger.log('Advance Reminder Sent: ' + (advanceReminderSentStr || 'BLANK'));
    Logger.log('Last Reminder Sent: ' + (lastReminderSentStr || 'BLANK'));

    let verdict;
    if (!isAdvanceMatch && !isDueTodayMatch) verdict = 'NOT ELIGIBLE TODAY — due date matches neither touchpoint';
    else if (!sendDuesBool) verdict = 'WOULD SKIP — Send Dues? is off';
    else if (isLapsed) verdict = 'WOULD SKIP — policy is lapsed';
    else if (!email) verdict = 'WOULD SKIP — no email on file';
    else if (isAdvanceMatch && advanceReminderSentStr === todayStr) verdict = 'WOULD SKIP — advance reminder already sent today';
    else if (isDueTodayMatch && lastReminderSentStr === todayStr) verdict = 'WOULD SKIP — due-today reminder already sent today';
    else verdict = 'WOULD SEND (daysAhead=' + (isDueTodayMatch ? 0 : advanceDays) + ')';

    Logger.log('VERDICT: ' + verdict);
    break;
  }
  if (!found) Logger.log('Policy ' + policyNumber + ' not found in Dues Tracker.');
}

// Sends a REAL advance-style email using this policy's actual client
// name, due date, and daysAhead — but to the advisor's own Contact
// Email, not the client, and clearly marked as a TEST in the subject.
// Deliberately does NOT touch Advance Reminder Sent, matching the
// same "test emails never affect real tracking" rule the existing
// due-today test email already follows.
function sendAdvanceReminderTestEmail(policyNumber){
  const tz = Session.getScriptTimeZone();
  const advanceDays = getAdvanceDays().advanceDays;

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  if (!sheet) throw new Error('Dues Tracker sheet not found.');
  const data = sheet.getDataRange().getValues();
  const headers = data[0];
  const col = name => headers.indexOf(name);

  let targetRow = null;
  for (let i = 1; i < data.length; i++){
    if (String(data[i][col('Policy Number')]).trim() === String(policyNumber).trim()){
      targetRow = data[i];
      break;
    }
  }
  if (!targetRow) throw new Error('Policy ' + policyNumber + ' not found in Dues Tracker.');

  const config = getBrandConfig();
  assertConfigured(config);
  const recipient = config.contactEmail;
  if (!recipient) throw new Error('Please set a Contact Email in Your Branding first.');

  const clientName = targetRow[col('Client Name')];
  const product = targetRow[col('Product')];
  const amount = targetRow[col('Premium Amount')];
  const dueDateRaw = targetRow[col('Due Date')];
  const dueDate = (dueDateRaw instanceof Date) ? dueDateRaw : new Date();

  const subjectDate = Utilities.formatDate(dueDate, tz, 'MMMM d').toUpperCase();
  const daysLabel = advanceDays === 1 ? ' IN 1 DAY' : ' IN ' + advanceDays + ' DAYS';
  const subject = 'TEST - PREMIUM DUE REMINDER' + daysLabel + ' - ' + subjectDate;
  const htmlBody = buildReminderEmailHtml(clientName, policyNumber, product, amount, dueDate, config, advanceDays);

  sendWithOptionalFromAlias(recipient, subject, {
    htmlBody: htmlBody,
    name: config.senderName,
    inlineImages: getEmailImages(config)
  }, config.contactEmail);

  return { success: true, sentTo: recipient, policyNumber: policyNumber, daysAhead: advanceDays };
}

// High-level troubleshooting summary — no emails sent, nothing written
// to the sheet. Run directly from the editor.
function runReminderDiagnosticNow(){
  const tz = Session.getScriptTimeZone();
  const todayStr = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const advanceDays = getAdvanceDays().advanceDays;
  const advanceTargetDate = new Date();
  advanceTargetDate.setDate(advanceTargetDate.getDate() + advanceDays);
  const advanceTargetStr = Utilities.formatDate(advanceTargetDate, tz, 'yyyy-MM-dd');
  const autoSend = getAutoSendStatus().enabled;
  const advisorActive = isAdvisorActive();
  const triggerFound = ScriptApp.getProjectTriggers().some(t => t.getHandlerFunction() === 'sendDailyReminders');
  const quota = MailApp.getRemainingDailyQuota();

  const sheet = getSpreadsheet().getSheetByName(SHEET_NAME);
  let dueTodayMatches = 0, advanceMatches = 0, alreadySent = 0, eligible = 0, skipped = 0;

  if (sheet){
    const data = sheet.getDataRange().getValues();
    const headers = data[0];
    const col = name => headers.indexOf(name);
    const advanceReminderColIdx = col('Advance Reminder Sent');

    for (let i = 1; i < data.length; i++){
      const row = data[i];
      const dueDateRaw = row[col('Due Date')];
      if (!(dueDateRaw instanceof Date)) continue;
      const dueDateStr = Utilities.formatDate(dueDateRaw, tz, 'yyyy-MM-dd');
      const isDueToday = dueDateStr === todayStr;
      const isAdvanceMatch = dueDateStr === advanceTargetStr;
      if (!isDueToday && !isAdvanceMatch) continue;
      if (isDueToday) dueTodayMatches++;
      if (isAdvanceMatch) advanceMatches++;

      const sendDues = row[col('Send Dues?')];
      const sendDuesBool = !(sendDues === false || sendDues === 'FALSE' || sendDues === 0 || sendDues === '0');
      const isLapsed = isLapsedStatus(row[col('Policy Status')]);
      const email = row[col('Email')];

      if (!sendDuesBool || isLapsed || !email){ skipped++; continue; }

      const trackingColIdx = isDueToday ? col('Last Reminder Sent') : advanceReminderColIdx;
      const sentStr = trackingColIdx !== -1 ? normalizeDateCellToYmd(row[trackingColIdx], tz) : '';
      if (sentStr === todayStr){ alreadySent++; } else { eligible++; }
    }
  }

  Logger.log('=== runReminderDiagnosticNow ===');
  Logger.log('Current Date: ' + todayStr);
  Logger.log('Advance Days Setting: ' + advanceDays);
  Logger.log('Advance Target Date: ' + advanceTargetStr);
  Logger.log('Due-Today Matches: ' + dueTodayMatches);
  Logger.log('Advance-Day Matches: ' + advanceMatches);
  Logger.log('Already Sent Today: ' + alreadySent);
  Logger.log('Eligible To Send: ' + eligible);
  Logger.log('Skipped (Send Dues off / lapsed / no email): ' + skipped);
  Logger.log('Remaining Email Quota: ' + quota);
  Logger.log('Auto-Send Status: ' + (autoSend ? 'ON' : 'OFF'));
  Logger.log('Daily Trigger Status: ' + (triggerFound ? 'FOUND' : 'MISSING'));
  Logger.log('Advisor Active: ' + (advisorActive ? 'YES' : 'NO (hard-stopped)'));
}
