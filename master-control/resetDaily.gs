// ============================================================================
// resetDaily.gs — Cold Email Automation System
// Location: Master Control Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Resets daily counters and writes the previous day's summary.
// Runs at 11:55 PM daily — last function to run each day.
//
// What it does:
//  1. Reads sentToday from each active sender's senderAccounts row
//  2. Writes a daily summary row to the dailySummary sheet
//     (date, totalSent, initialSent, followupsSent, errors, activesSenders)
//  3. Resets sentToday = 0 for all active senders in senderAccounts
//  4. Clears all rows from each sender's dailyInbox and followupInbox
//     that have status = 'Sent' or status = 'Error' (permanent)
//     Rows with status = 'Queued' or 'Sending' are left in place
//     (Sending rows are treated as potential crashes — left for retry)
//
// Key design rules:
//  - Runs at 11:55 PM — after all sends are done for the day
//  - sentToday reset happens AFTER summary is written — never before
//  - Inbox clearing only removes terminal rows (Sent, permanent Error)
//  - Queued rows that were never sent remain for next day's run
//  - All sheet reads/writes in batch (getValues / setValues)
//  - Lock acquired before any shared write
//  - Each sender's inbox is cleared via webhook POST to sender webApp
//    (sender script clears its own sheets — Master Control cannot write
//     directly to sender spreadsheets efficiently at scale)
//
// Dependencies:
//   utils.gs, distributor.gs must be present in Master Control project.
// ============================================================================


// ============================================================================
// SECTION 1 — CONSTANTS
// ============================================================================

const RESET_LOCK_WAIT      = 10;            // HARDCODED — verify this
const DAILY_SUMMARY_SHEET  = 'dailySummary'; // HARDCODED — verify this
const ACTION_RESET_INBOX   = 'resetInbox';   // HARDCODED — verify this

// Status values considered terminal (safe to clear from inbox after day ends)
const TERMINAL_STATUSES = ['Sent', 'Error']; // HARDCODED — verify this


// ============================================================================
// SECTION 2 — runDailyReset()
// Main function. Triggered at 11:55 PM daily from Master Control.
// ============================================================================

/**
 * Runs the end-of-day reset: writes summary, resets counters, clears inboxes.
 * Triggered at 11:55 PM daily from Master Control spreadsheet.
 */
function runDailyReset() {
  log('runDailyReset: started', 'INFO');

  if (!acquireLock(RESET_LOCK_WAIT)) {
    log('runDailyReset: could not acquire lock — exiting', 'WARN');
    return;
  }

  try {
    // --- Step 1: Read all active senders and their daily counters ---
    const { senders, sheet: senderSheet, data: senderData, headers: senderHeaders } =
      readSenderAccountsForReset();

    if (senders.length === 0) {
      log('runDailyReset: no active senders found — exiting', 'INFO');
      return;
    }

    // --- Step 2: Aggregate daily totals ---
    let totalInitialSent  = 0;
    let totalFollowupSent = 0;
    let totalErrors       = 0;

    let totalSent = 0;
    for (let i = 0; i < senders.length; i++) {
      totalSent += senders[i].sentToday || 0;
    }

    // --- Step 3: Write daily summary row ---
    writeDailySummaryRow({
      date:           formatDate(new Date()),
      totalSent:      totalSent,
      initialSent:    totalInitialSent,
      followupsSent:  totalFollowupSent,
      errors:         totalErrors,
      activeSenders:  senders.length
    });

    // --- Step 4: Reset daily counters for all active senders ---
    resetSenderDailyCounters(senderSheet, senderData, senderHeaders, senders);
    log('runDailyReset: daily counters reset for all senders', 'INFO');

    log('runDailyReset: summary written — totalSent=' + totalSent +
        ' initialSent=' + totalInitialSent +
        ' followupsSent=' + totalFollowupSent +
        ' errors=' + totalErrors, 'INFO');

    // --- Step 5: Tell all senders to clear inboxes in parallel batches ---
    const resetRequests = [];
    const resetMeta     = [];

    for (let i = 0; i < senders.length; i++) {
      const sender = senders[i];
      if (!sender.webAppUrl) { continue; }

      resetRequests.push({
        url:                sender.webAppUrl,
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({
          action:           ACTION_RESET_INBOX,
          terminalStatuses: TERMINAL_STATUSES
        }),
        muteHttpExceptions: true
      });

      resetMeta.push({ emailID: sender.emailID });
    }

    // Fire in parallel batches of 50
    for (let b = 0; b < resetRequests.length; b += 50) {
      const batchRequests = resetRequests.slice(b, b + 50);
      const batchMeta     = resetMeta.slice(b, b + 50);

      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(batchRequests);
      } catch (e) {
        log('runDailyReset: fetchAll error on batch ' + Math.floor(b / 50 + 1) + ' — ' + e.message, 'ERROR');
        continue;
      }

      for (let r = 0; r < responses.length; r++) {
        const code = responses[r].getResponseCode();
        if (code === 200) {
          log('runDailyReset: sender ' + batchMeta[r].emailID + ' inbox cleared', 'INFO');
        } else {
          log('runDailyReset: could not clear inbox for sender ' + batchMeta[r].emailID + ' (' + code + ')', 'ERROR');
        }
      }
    }

    log('runDailyReset: completed successfully', 'INFO');

  } catch (e) {
    log('runDailyReset: unexpected error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}



/**
 * Resets sentToday = 0 for all senders in senderAccounts.
 * Also POSTs reset to each sender's Config sheet via webhook.
 * Triggered at 6:00 AM daily — BEFORE distributeCold() at 7:00 AM.
 */
function resetSentToday() {
  log('resetSentToday: started', 'INFO');
  if (!acquireLock(RESET_LOCK_WAIT)) {
    log('resetSentToday: could not acquire lock — exiting', 'WARN');
    return;
  }
  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    if (!sheet || sheet.getLastRow() < 2) {
      log('resetSentToday: no sender rows found — exiting', 'WARN');
      return;
    }
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Reset sentToday to 0 for ALL rows in one pass
    for (let i = 0; i < data.length; i++) {
      data[i][col['sentToday']] = 0;
    }
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    SpreadsheetApp.flush();
    log('resetSentToday: reset sentToday for ' + data.length + ' senders in senderAccounts', 'INFO');

    // Reset sentToday in each active sender's Config sheet via webhook — batched fetchAll
    const resetRequests = [];
    for (let i = 0; i < data.length; i++) {
      const isActive  = String(data[i][col['isActive']] || '').trim();
      const webAppUrl = String(data[i][col['webAppUrl']] || '').trim();
      if (isActive !== 'Active' || !webAppUrl) { continue; }
      resetRequests.push({
        url:                webAppUrl,
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({ action: 'resetSentToday' }),
        muteHttpExceptions: true
      });
    }
    // Fire in batches of 50
    for (let i = 0; i < resetRequests.length; i += 50) {
      const batch = resetRequests.slice(i, i + 50);
      try {
        UrlFetchApp.fetchAll(batch);
        log('resetSentToday: batch ' + (Math.floor(i / 50) + 1) + ' of ' + Math.ceil(resetRequests.length / 50) + ' sent', 'INFO');
      } catch (e) {
        log('resetSentToday: fetchAll batch error — ' + e.message, 'ERROR');
      }
    }
    log('resetSentToday: completed', 'INFO');
  } catch (e) {
    log('resetSentToday: unexpected error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}


function setupResetSentTodayTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'resetSentToday') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('resetSentToday')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(0)
    .inTimezone('Africa/Cairo')
    .create();
  log('setupResetSentTodayTrigger: created 6:00 AM trigger', 'INFO');
}


// ============================================================================
// SECTION 3 — readSenderAccountsForReset()
// Reads senderAccounts sheet and returns active senders with daily counters.
// Returns the raw sheet reference and data array so the caller can write
// resets back without opening the spreadsheet a second time.
// ============================================================================

/**
 * Reads senderAccounts and returns active senders with their daily counters.
 *
 * @returns {{
 *   senders: Array,
 *   sheet:   GoogleAppsScript.Spreadsheet.Sheet,
 *   data:    Array,
 *   headers: Array
 * }}
 *
 * Sender object includes:
 *   emailID, emailAddress, webAppUrl, dataIndex (0-based in data array),
 *   initialSentToday, followupSentToday, errorsToday, sentToday
 */
function readSenderAccountsForReset() {
  const result = { senders: [], sheet: null, data: [], headers: [] };

  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);

  if (!sheet) {
    log('readSenderAccountsForReset: senderAccounts sheet not found', 'ERROR');
    return result;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { return result; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  const col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  const senders = [];

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const isActive = String(row[col['isActive']] || '').trim();
    if (isActive !== 'Active') { continue; }

    senders.push({
      emailID:          String(row[col['emailID']]          || '').trim(),
      emailAddress:     String(row[col['emailAddress']]     || '').trim(),
      webAppUrl:        String(row[col['webAppUrl']]        || '').trim(),
      dataIndex:        i, // 0-based index in data array
      sentToday:        parseInt(row[col['sentToday']]        || 0, 10),
      initialSentToday: parseInt(row[col['initialSentToday']] || 0, 10), // HARDCODED — verify this
      followupSentToday:parseInt(row[col['followupSentToday']]|| 0, 10), // HARDCODED — verify this
      errorsToday:      parseInt(row[col['errorsToday']]      || 0, 10)  // HARDCODED — verify this
    });
  }

  result.senders = senders;
  result.sheet   = sheet;
  result.data    = data;
  result.headers = headers;

  return result;
}


// ============================================================================
// SECTION 4 — writeDailySummaryRow()
// Appends one row to the dailySummary sheet in Master Control.
// Creates the sheet with headers if it does not exist yet.
// ============================================================================

/**
 * Appends a daily summary row to the dailySummary sheet.
 *
 * @param {Object} summary - Summary object with totals for the day.
 *   { date, totalSent, initialSent, followupsSent, errors, activeSenders }
 * @returns {void}
 *
 * dailySummary sheet columns:
 *   date | totalSent | initialSent | followupsSent | errors | activeSenders
 */
function writeDailySummaryRow(summary) {
  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  let   sheet = ss.getSheetByName(DAILY_SUMMARY_SHEET);

  // Create sheet with headers if it does not exist
  if (!sheet) {
    sheet = ss.insertSheet(DAILY_SUMMARY_SHEET);
    const summaryHeaders = [
      'date', 'emailsSent', 'initialSent', 'followupsSent',
      'repliesReceived', 'pendingLeads', 'activeSenders', 'activeCampaigns', 'bouncedEmails', 'notes'
    ]; // HARDCODED — verify this
    sheet.getRange(1, 1, 1, summaryHeaders.length).setValues([summaryHeaders]);

    // Style header
    const headerRange = sheet.getRange(1, 1, 1, summaryHeaders.length);
    headerRange.setBackground('#1a73e8');
    headerRange.setFontColor('#ffffff');
    headerRange.setFontWeight('bold');
    sheet.setFrozenRows(1);

    log('writeDailySummaryRow: created dailySummary sheet with headers', 'INFO');
  }

  // Append summary row
  sheet.appendRow([
    summary.date,
    summary.totalSent,
    summary.initialSent,
    summary.followupsSent,
    summary.repliesReceived  || 0,
    summary.pendingLeads     || 0,
    summary.activeSenders,
    summary.activeCampaigns  || 0,
    summary.bouncedEmails    || 0,
    summary.notes            || ''
  ]);

  SpreadsheetApp.flush();
  log('writeDailySummaryRow: appended summary for ' + summary.date, 'INFO');
}


// ============================================================================
// SECTION 5 — resetSenderDailyCounters()
// Resets sentToday, initialSentToday, followupSentToday, errorsToday to 0
// for all active senders in ONE setValues() call.
// ============================================================================

/**
 * Resets daily counter columns to 0 for all active senders.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet   - senderAccounts sheet.
 * @param {Array}  data    - Full data array read from senderAccounts.
 * @param {Array}  headers - Header row array.
 * @param {Array}  senders - Active sender objects with dataIndex references.
 * @returns {void}
 *
 * Columns reset: sentToday, initialSentToday, followupSentToday, errorsToday.
 * Uses the dataIndex from each sender to update the correct row in data array.
 * Writes entire data array back in ONE setValues() call.
 */
function resetSenderDailyCounters(sheet, data, headers, senders) {
  if (!sheet || !data || data.length === 0) {
    log('resetSenderDailyCounters: no data to reset', 'WARN');
    return;
  }

  const col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  const counterCols = ['sentToday', 'initialSentToday', 'followupSentToday', 'errorsToday'];

  // Update in-memory data array for each active sender
  for (let i = 0; i < senders.length; i++) {
    const sender   = senders[i];
    const dataIdx  = sender.dataIndex;

    for (let c = 0; c < counterCols.length; c++) {
      const colName = counterCols[c];
      if (colName in col) {
        data[dataIdx][col[colName]] = 0;
      }
    }
  }

  // Write entire data array back in ONE setValues() call
  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();
  log('resetSenderDailyCounters: reset counters for ' + senders.length + ' senders', 'INFO');
}


// ============================================================================
// SECTION 6 — notifySenderToReset()
// POSTs a resetInbox action to a sender's webApp.
// The sender's webApp.gs handles clearing terminal rows from its own
// dailyInbox and followupInbox sheets.
// ============================================================================

/**
 * Sends a resetInbox POST to a sender's webApp.
 *
 * @param {string} webAppUrl - Sender's deployed webApp URL.
 * @returns {boolean} true if HTTP 200, false otherwise.
 *
 * Payload: { action: 'resetInbox', terminalStatuses: ['Sent', 'Error'] }
 * The sender's doPost router handles clearing those rows from both inboxes.
 */
function notifySenderToReset(webAppUrl) {
  if (!webAppUrl) {
    log('notifySenderToReset: missing webAppUrl', 'ERROR');
    return false;
  }

  const payload = JSON.stringify({
    action:           ACTION_RESET_INBOX,
    terminalStatuses: TERMINAL_STATUSES
  });

  const options = {
    method:             'post',
    contentType:        'application/json',
    payload:            payload,
    muteHttpExceptions: true
  };

  try {
    const response     = UrlFetchApp.fetch(webAppUrl, options);
    const responseCode = response.getResponseCode();

    if (responseCode !== 200) {
      log('notifySenderToReset: non-200 response ' + responseCode +
          ' from ' + webAppUrl, 'ERROR');
      return false;
    }

    return true;

  } catch (e) {
    log('notifySenderToReset: fetch error — ' + e.message, 'ERROR');
    return false;
  }
}


// ============================================================================
// SECTION 7 — setupDailyResetTrigger()
// One-time setup. Creates the 11:55 PM daily trigger.
// Run manually once from Master Control spreadsheet.
// ============================================================================

/**
 * Creates the daily 11:55 PM trigger for runDailyReset().
 * Run ONCE manually from Master Control Apps Script editor.
 */
function setupDailyResetTrigger() {
  // Delete existing triggers for this function
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runDailyReset') {
      ScriptApp.deleteTrigger(triggers[i]);
      log('setupDailyResetTrigger: deleted existing trigger', 'INFO');
    }
  }

  // Apps Script atHour() does not support minutes directly for daily triggers.
  // We use atHour(23) which fires between 23:00-24:00. The nearMinute(55)
  // pushes it as close to 11:55 PM as the scheduler allows. // HARDCODED — verify this
  ScriptApp.newTrigger('runDailyReset')
    .timeBased()
    .everyDays(1)
    .atHour(23)         // HARDCODED — verify this
    .nearMinute(55)     // HARDCODED — verify this
    .inTimezone('Africa/Cairo') // HARDCODED — verify this
    .create();

  log('setupDailyResetTrigger: created 11:55 PM daily trigger', 'INFO');
  Logger.log('✅ Daily reset trigger created — runs at ~11:55 PM Cairo time.');
}

function runSyncStats() {
  const settings = getSettings(MASTER_SHEET_ID);
  syncCampaignStats(MASTER_SHEET_ID, settings.leadsSpreadsheetId);
}

function setupSyncStatsTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runSyncStats') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('runSyncStats')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(30)
    .inTimezone('Africa/Cairo')
    .create();
  log('setupSyncStatsTrigger: created 11:30 PM trigger', 'INFO');
}
