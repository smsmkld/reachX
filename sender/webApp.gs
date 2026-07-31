// ============================================================================
// webApp.gs — Cold Email Automation System
// Location: BOTH Master Control AND each Sender Spreadsheet
// ----------------------------------------------------------------------------
// Single file deployed in two contexts with different doPost routers.
// Same file is copied to both — the CONTEXT determines which actions apply.
//
// MASTER CONTROL webApp handles:
//   - updateColdSent       → updates masterLeads after cold send
//   - updateFollowupSent   → updates masterLeads after followup send
//   - updateLeadReplied    → marks lead as Replied in masterLeads
//
// SENDER webApp handles:
//   - writeColdInbox       → writes rows to sender's dailyInbox
//   - writeFollowupInbox   → writes rows to sender's followupInbox
//   - syncConfig           → writes campaign templates to sender's Config sheet
//   - resetInbox           → clears terminal rows from sender's inboxes
//   - scanReplies          → scans sender's Gmail for replies, returns results
//
// Deployment rules (IMPORTANT):
//   - Execute as: Me (the spreadsheet owner)
//   - Who has access: Anyone (no login required)
//   - doGet() is NOT implemented — only doPost()
//   - All responses include CORS headers
//   - All requests must include correct action string in JSON body
//   - No authentication token in this version — add if exposing to internet
//
// Key design rules:
//   - doPost() router reads action from JSON body — never from URL params
//   - Every handler wraps in try/catch — never let doPost() throw uncaught
//   - All sheet writes use batch getValues/setValues — never cell by cell
//   - Lock acquired per handler that writes to sheets
//   - Returns { success: true/false, error: '...' } JSON on all responses
//
// Dependencies (Master Control context):
//   utils.gs, distributor.gs must be present
// Dependencies (Sender context):
//   utils.gs, sender.gs, followup.gs must be present
// ============================================================================


// ============================================================================
// SECTION 1 — doPost() ROUTER
// Single entry point for all incoming POST requests.
// Reads action from JSON body and routes to correct handler.
// Returns JSON response with CORS headers on every path.
// ============================================================================

/**
 * Main HTTP POST entry point. Routes requests by action field.
 * Deployed as a Web App — Execute as Me, Anyone can access.
 *
 * @param {Object} e - Apps Script event object with postData.
 * @returns {GoogleAppsScript.Content.TextOutput} JSON response.
 */
const WEBAPP_REPLY_SCAN_DAYS      = 30;  // HARDCODED — verify this
const WEBAPP_MAX_THREADS_PER_SENDER = 500; // HARDCODED — verify this
const WEBAPP_BOUNCE_SCAN_DAYS       = 30;

function doPost(e) {
  // Always return CORS headers — required for Cloudflare Pages frontend
  const corsHeaders = {
    'Access-Control-Allow-Origin':  '*', // HARDCODED — verify this (restrict to your domain in production)
    'Access-Control-Allow-Methods': 'POST',
    'Access-Control-Allow-Headers': 'Content-Type'
  };

  try {
    // Parse request body
    if (!e || !e.postData || !e.postData.contents) {
      return buildResponse({ success: false, error: 'Empty request body' }, corsHeaders);
    }

    let params;
    try {
      params = JSON.parse(e.postData.contents);
    } catch (parseError) {
      return buildResponse({ success: false, error: 'Invalid JSON in request body' }, corsHeaders);
    }

    const action = String(params.action || '').trim();

    if (!action) {
      return buildResponse({ success: false, error: 'Missing action field' }, corsHeaders);
    }

    log('doPost: received action=' + action, 'INFO');

    // ---- MASTER CONTROL ACTIONS ----
    // These actions only make sense when deployed in Master Control spreadsheet.
    // If this webApp is deployed in a sender spreadsheet, these will still route
    // but will fail gracefully because the required sheets won't exist.

    if (action === 'updateColdSent') {
      return buildResponse(handleUpdateColdSent(params), corsHeaders);
    }

    if (action === 'updateFollowupSent') {
      return buildResponse(handleUpdateFollowupSent(params), corsHeaders);
    }

    if (action === 'updateLeadReplied') {
      return buildResponse(handleUpdateLeadReplied(params), corsHeaders);
    }

    // ---- SENDER ACTIONS ----
    // These actions only make sense when deployed in a sender spreadsheet.

    if (action === 'writeColdInbox') {
      return buildResponse(handleWriteColdInbox(params), corsHeaders);
    }

    if (action === 'writeFollowupInbox') {
      return buildResponse(handleWriteFollowupInbox(params), corsHeaders);
    }

    if (action === 'syncConfig') {
      return buildResponse(handleSyncConfig(params), corsHeaders);
    }

    if (action === 'resetInbox') {
      return buildResponse(handleResetInbox(params), corsHeaders);
    }

    if (action === 'scanReplies') {
      return buildResponse(handleScanReplies(params), corsHeaders);
    }

    if (action === 'scanBounces') {
      return buildResponse(handleScanBounces(params), corsHeaders);
    }

    if (action === 'kickoffAllSenders') {
      return buildResponse(handleKickoffAllSenders(params), corsHeaders);
    }

    if (action === 'kickoffAllFollowups') {
      return buildResponse(handleKickoffAllFollowups(params), corsHeaders);
    }

    if (action === 'emergencyStop') {
      return buildResponse(handleEmergencyStop(params), corsHeaders);
    }

    if (action === 'resumeAll') {
      return buildResponse(handleResumeAll(params), corsHeaders);
    }

    if (action === 'setupAllTriggers') {
      return buildResponse(handleSetupAllTriggers(params), corsHeaders);
    }

    if (action === 'cleanupAllTriggers') {
      return buildResponse(handleCleanupAllTriggers(params), corsHeaders);
    }
    if (action === 'purgeAllInboxes') {
      return buildResponse(handlePurgeAllInboxes(params), corsHeaders);
    }
    if (action === 'sendTestEmail') {
      return buildResponse(handleSendTestEmailDirect(params), corsHeaders);
    }
    if (action === 'kickoffCold') {
      return buildResponse(handleKickoffCold(params), corsHeaders);
    }

    if (action === 'kickoffFollowup') {
      return buildResponse(handleKickoffFollowup(params), corsHeaders);
    }

    if (action === 'setupTriggers') {
      return buildResponse(handleSetupTriggers(params), corsHeaders);
    }

    if (action === 'cleanupTriggers') {
      return buildResponse(handleCleanupTriggers(params), corsHeaders);
    }

    if (action === 'getSenderStatus') {
      return buildResponse(handleGetSenderStatus(params), corsHeaders);
    }

    if (action === 'resetSentToday') {
      return buildResponse(handleResetSentToday(params), corsHeaders);
    }

    if (action === 'purgeInbox') {
      return buildResponse(handlePurgeInbox(params), corsHeaders);
    }
    // Unknown action
    return buildResponse({ success: false, error: 'Unknown action: ' + action }, corsHeaders);

  } catch (e) {
    log('doPost: uncaught error — ' + e.message, 'ERROR');
    return buildResponse({ success: false, error: 'Internal server error: ' + e.message }, corsHeaders);
  }
}


// ============================================================================
// SECTION 2 — buildResponse()
// Builds a JSON TextOutput with CORS headers.
// Called by doPost() on every response path.
// ============================================================================

/**
 * Builds a JSON response with CORS headers.
 *
 * @param {Object} data        - Response data object.
 * @param {Object} corsHeaders - CORS header key-value pairs.
 * @returns {GoogleAppsScript.Content.TextOutput}
 */
function buildResponse(data, corsHeaders) {
  const output = ContentService
    .createTextOutput(JSON.stringify(data))
    .setMimeType(ContentService.MimeType.JSON);

  // Note: Apps Script Web Apps do not support custom response headers directly.
  // CORS is handled by the Cloudflare Pages frontend making requests — the
  // '*' origin in doPost is for documentation purposes. Apps Script deployed
  // with "Anyone" access automatically allows cross-origin POST requests.
  // If you proxy through Cloudflare Workers, add CORS headers there instead.

  return output;
}


// ============================================================================
// SECTION 3 — MASTER CONTROL HANDLERS
// handleUpdateColdSent(), handleUpdateFollowupSent(), handleUpdateLeadReplied()
// These run in the Master Control spreadsheet context.
// They update masterLeads DB and senderAccounts counters.
// ============================================================================

/**
 * Handles updateColdSent action from sender scripts.
 * Updates masterLeads: status=Sent, threadId, nextSentDate, sequenceStep.
 * Increments initialSentToday and sentToday in senderAccounts.
 *
 * @param {Object} params - Parsed request body.
 *   params.updates: Array of {
 *     leadId, campaignId, status, sentTime,
 *     threadId, sequenceStep, nextSentDate
 *   }
 * @returns {{ success: boolean, updated: number, error?: string }}
 */
function handleUpdateColdSent(params) {
  const updates = params.updates || [];
  if (updates.length === 0) {
    return { success: false, error: 'No updates provided' };
  }

  try {
    const ss           = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const stagingSheet = ss.getSheetByName('stagingUpdates');
    if (!stagingSheet) {
      return { success: false, error: 'stagingUpdates sheet not found' };
    }

    const rows = updates.map(function(upd) {
      return [
        new Date().toISOString(), // timestamp
        params.senderId || '',
        'coldSent',
        upd.leadId        || '',
        upd.status        || 'Sent',
        upd.threadId      || '',
        upd.sentTime      || '',
        upd.nextSentDate  || '',
        upd.sequenceStep  || 1,
        '',                        // sequenceStatus — not used for cold
        ''                         // lastSentTime — not used for cold
      ];
    });

    const startRow = stagingSheet.getLastRow() + 1;
    stagingSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();

    // Still increment sender counters — this is fast, no masterLeads touch
    incrementSenderCounters(updates.length, 0, params.senderId || '');

    log('handleUpdateColdSent: staged ' + rows.length + ' updates', 'INFO');
    return { success: true, staged: rows.length };

  } catch (e) {
    log('handleUpdateColdSent: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}


/**
 * Handles updateFollowupSent action from sender scripts.
 * Updates masterLeads: status, lastSentDate, nextSentDate, sequenceStep,
 * and sequenceStatus=Completed if final step.
 *
 * @param {Object} params - Parsed request body.
 *   params.updates: Array of {
 *     leadId, campaignId, status, lastSentTime,
 *     sequenceStep, nextSentDate, sequenceStatus
 *   }
 * @returns {{ success: boolean, updated: number, error?: string }}
 */
function handleUpdateFollowupSent(params) {
  const updates = params.updates || [];
  if (updates.length === 0) {
    return { success: false, error: 'No updates provided' };
  }

  try {
    const ss           = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const stagingSheet = ss.getSheetByName('stagingUpdates');
    if (!stagingSheet) {
      return { success: false, error: 'stagingUpdates sheet not found' };
    }

    const rows = updates.map(function(upd) {
      return [
        new Date().toISOString(), // timestamp
        params.senderId    || '',
        'followupSent',
        upd.leadId         || '',
        upd.status         || 'Sent',
        '',                        // threadId — not used for followups
        '',                        // sentTime — not used for followups
        upd.nextSentDate   || '',
        upd.sequenceStep   || 1,
        upd.sequenceStatus || '',
        upd.lastSentTime   || ''
      ];
    });

    const startRow = stagingSheet.getLastRow() + 1;
    stagingSheet.getRange(startRow, 1, rows.length, rows[0].length).setValues(rows);
    SpreadsheetApp.flush();

    // Increment sender counters — fast, no masterLeads touch
    incrementSenderCounters(0, updates.length, params.senderId || '');

    log('handleUpdateFollowupSent: staged ' + rows.length + ' updates', 'INFO');
    return { success: true, staged: rows.length };

  } catch (e) {
    log('handleUpdateFollowupSent: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}


/**
 * Handles updateLeadReplied action.
 * Marks a lead as Replied in masterLeads.
 * Called by sender's scanReplies handler after detecting a reply.
 *
 * @param {Object} params - { leadId, threadId, replyDate }
 * @returns {{ success: boolean, error?: string }}
 */
function handleUpdateLeadReplied(params) {
  const leadId  = String(params.leadId  || '').trim();
  const threadId = String(params.threadId || '').trim();

  if (!leadId) {
    return { success: false, error: 'Missing leadId' };
  }

  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;

    const lss    = SpreadsheetApp.openById(leadsSheetId);
    const lSheet = lss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow = lSheet.getLastRow();
    const headers = lSheet.getRange(1, 1, 1, lSheet.getLastColumn()).getValues()[0];
    const data    = lSheet.getRange(2, 1, lastRow - 1, lSheet.getLastColumn()).getValues();

    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    let found = false;
    const today = formatDate(new Date());

    for (let i = 0; i < data.length; i++) {
      const rowLeadId = String(data[i][col['leadId']] || '').trim();
      if (rowLeadId !== leadId) { continue; }

      // Already replied — skip
      if (String(data[i][col['replyStatus']] || '').trim() === 'Replied') {
        log('handleUpdateLeadReplied: ' + leadId + ' already marked Replied', 'INFO');
        found = true;
        break;
      }

      if ('replyStatus' in col) { data[i][col['replyStatus']] = 'Replied'; }
      if ('replyDate'   in col) { data[i][col['replyDate']]   = params.replyDate || today; }

      found = true;
      break;
    }

    if (found) {
      lSheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }

    return { success: true, found: found };

  } catch (e) {
    log('handleUpdateLeadReplied: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}


// ============================================================================
// SECTION 4 — SENDER HANDLERS
// handleWriteColdInbox(), handleWriteFollowupInbox(), handleSyncConfig(),
// handleResetInbox(), handleScanReplies()
// These run in the SENDER spreadsheet context.
// ============================================================================

/**
 * Handles writeColdInbox action from distributor.gs.
 * Appends rows to sender's dailyInbox sheet.
 *
 * @param {Object} params - { rows: Array of inbox row objects }
 * @returns {{ success: boolean, written: number, error?: string }}
 */
function handleWriteColdInbox(params) {
  const rows = params.rows || [];
  if (rows.length === 0) {
    return { success: false, error: 'No rows provided' };
  }

  if (!acquireLock(SENDER_LOCK_WAIT)) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(DAILY_INBOX_SHEET);

    if (!sheet) {
      return { success: false, error: 'dailyInbox sheet not found' };
    }

    // Read existing headers to ensure column order matches
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Build 2D array of rows to append
    const rowsToAppend = rows.map(function(row) {
      const newRow = new Array(headers.length).fill('');
      if ('leadId'       in col) { newRow[col['leadId']]       = row.leadId       || ''; }
      if ('email'        in col) { newRow[col['email']]        = row.email        || ''; }
      if ('campaignId'   in col) { newRow[col['campaignId']]   = row.campaignId   || ''; }
      if ('sequenceStep' in col) { newRow[col['sequenceStep']] = row.sequenceStep || 1; }
      if ('leadData'     in col) { newRow[col['leadData']]     = row.leadData     || '{}'; }
      if ('status'       in col) { newRow[col['status']]       = row.status       || 'Queued'; }
      if ('sentTime'     in col) { newRow[col['sentTime']]     = row.sentTime     || ''; }
      if ('retryCount'   in col) { newRow[col['retryCount']]   = row.retryCount   || 0; }
      if ('windowStart'  in col) { newRow[col['windowStart']]  = row.windowStart  || ''; }
      if ('windowEnd'    in col) { newRow[col['windowEnd']]    = row.windowEnd    || ''; }
      if ('sendingDays'  in col) { newRow[col['sendingDays']]  = row.sendingDays  || ''; }
      if ('timezone'     in col) { newRow[col['timezone']]     = row.timezone     || ''; }
      return newRow;
    });

    // Append all rows in ONE setValues() call
    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, headers.length)
         .setValues(rowsToAppend);
    SpreadsheetApp.flush();

    log('handleWriteColdInbox: wrote ' + rowsToAppend.length + ' rows to dailyInbox', 'INFO');
    return { success: true, written: rowsToAppend.length };

  } catch (e) {
    log('handleWriteColdInbox: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}


/**
 * Handles writeFollowupInbox action from distributor.gs.
 * Appends rows to sender's followupInbox sheet.
 *
 * @param {Object} params - { rows: Array of followup inbox row objects }
 * @returns {{ success: boolean, written: number, error?: string }}
 */
function handleWriteFollowupInbox(params) {
  const rows = params.rows || [];
  if (rows.length === 0) {
    return { success: false, error: 'No rows provided' };
  }

  if (!acquireLock(SENDER_LOCK_WAIT)) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    const ss    = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);

    if (!sheet) {
      return { success: false, error: 'followupInbox sheet not found' };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const rowsToAppend = rows.map(function(row) {
      const newRow = new Array(headers.length).fill('');
      if ('leadId'             in col) { newRow[col['leadId']]             = row.leadId             || ''; }
      if ('email'              in col) { newRow[col['email']]              = row.email              || ''; }
      if ('campaignId'         in col) { newRow[col['campaignId']]         = row.campaignId         || ''; }
      if ('sequenceStep'       in col) { newRow[col['sequenceStep']]       = row.sequenceStep       || 1; }
      if ('totalSequenceSteps' in col) { newRow[col['totalSequenceSteps']] = row.totalSequenceSteps || 1; }
      if ('leadData'           in col) { newRow[col['leadData']]           = row.leadData           || '{}'; }
      if ('threadId'           in col) { newRow[col['threadId']]           = row.threadId           || ''; }
      if ('status'             in col) { newRow[col['status']]             = row.status             || 'Queued'; }
      if ('lastSentTime'       in col) { newRow[col['lastSentTime']]       = row.lastSentTime       || ''; }
      if ('retryCount'   in col) { newRow[col['retryCount']]   = row.retryCount   || 0; }
      if ('windowStart'  in col) { newRow[col['windowStart']]  = row.windowStart  || ''; }
      if ('windowEnd'    in col) { newRow[col['windowEnd']]    = row.windowEnd    || ''; }
      if ('sendingDays'  in col) { newRow[col['sendingDays']]  = row.sendingDays  || ''; }
      if ('timezone'     in col) { newRow[col['timezone']]     = row.timezone     || ''; }
      return newRow;
    });

    const startRow = sheet.getLastRow() + 1;
    sheet.getRange(startRow, 1, rowsToAppend.length, headers.length)
         .setValues(rowsToAppend);
    SpreadsheetApp.flush();

    log('handleWriteFollowupInbox: wrote ' + rowsToAppend.length + ' rows to followupInbox', 'INFO');
    return { success: true, written: rowsToAppend.length };

  } catch (e) {
    log('handleWriteFollowupInbox: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}


/**
 * Handles syncConfig action from campaignSync.gs.
 * Writes campaign templates and settings to sender's Config sheet.
 *
 * @param {Object} params - { campaignId, config: { key: value, ... } }
 * @returns {{ success: boolean, written: number, error?: string }}
 */
function handleSyncConfig(params) {
  const configData = params.config || {};
  if (Object.keys(configData).length === 0) {
    return { success: false, error: 'No config data provided' };
  }

  if (!acquireLock(SENDER_LOCK_WAIT)) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET);

    if (!configSheet) {
      return { success: false, error: 'Config sheet not found' };
    }

    const keys    = Object.keys(configData);
    const lastRow = configSheet.getLastRow();
    const data    = configSheet.getRange(1, 1, lastRow, 2).getValues();
    const keyMap  = {};
    data.forEach(function(row, i) { keyMap[String(row[0]).trim()] = i; });

    keys.forEach(function(key) {
      const value = String(configData[key] || '');
      if (key in keyMap) {
        data[keyMap[key]][1] = value;
      } else {
        data.push([key, value]);
      }
    });

    configSheet.getRange(1, 1, data.length, 2).setValues(data);
    SpreadsheetApp.flush();

    log('handleSyncConfig: wrote ' + keys.length + ' config values for campaign ' +
        (params.campaignId || ''), 'INFO');
    return { success: true, written: keys.length };

  } catch (e) {
    log('handleSyncConfig: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}


/**
 * Handles resetInbox action from resetDaily.gs.
 * Removes rows with terminal statuses from dailyInbox and followupInbox.
 * Leaves Queued and Sending rows in place.
 *
 * @param {Object} params - { terminalStatuses: ['Sent', 'Error'] }
 * @returns {{ success: boolean, removedCold: number, removedFollowup: number }}
 */
function handleResetInbox(params) {
  const terminalStatuses = params.terminalStatuses || TERMINAL_STATUSES;

  if (!acquireLock(SENDER_LOCK_WAIT)) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const removedCold    = clearTerminalRows(ss.getSheetByName(DAILY_INBOX_SHEET),    terminalStatuses);
    const removedFollowup = clearTerminalRows(ss.getSheetByName(FOLLOWUP_INBOX_SHEET), terminalStatuses);

    log('handleResetInbox: removed ' + removedCold + ' cold rows, ' +
        removedFollowup + ' followup rows', 'INFO');
    return { success: true, removedCold: removedCold, removedFollowup: removedFollowup };

  } catch (e) {
    log('handleResetInbox: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}


/**
 * Handles scanReplies action from replyScanner.gs.
 * Runs GmailApp.search() in this sender's Gmail context.
 * Returns array of { threadId, replyDate } for threads with inbound replies.
 *
 * @param {Object} params - { threadIds: ['id1','id2',...], senderEmail: '...' }
 * @returns {{ success: boolean, replies: Array }}
 */
function handleScanReplies(params) {
  const threadIdsToCheck = params.threadIds  || [];
  const senderEmail      = String(params.senderEmail || '').toLowerCase();

  if (threadIdsToCheck.length === 0) {
    return { success: true, replies: [] };
  }

  try {
    // Build a Set for O(1) lookups
    const threadIdSet = {};
    threadIdsToCheck.forEach(function(id) { threadIdSet[id] = true; });

    // Search sender's Gmail inbox for threads
    const scanFromDate = new Date();
    scanFromDate.setDate(scanFromDate.getDate() - WEBAPP_REPLY_SCAN_DAYS);
    const dateFilter = formatDate(scanFromDate).replace(/-/g, '/');
    const query      = 'in:inbox after:' + dateFilter; // HARDCODED — verify this

    const threads = GmailApp.search(query, 0, WEBAPP_MAX_THREADS_PER_SENDER);
    const replies = [];

    for (let t = 0; t < threads.length; t++) {
      const thread   = threads[t];
      const threadId = thread.getId();

      if (!(threadId in threadIdSet)) { continue; }

      const messages = thread.getMessages();
      let   replyFound = false;
      let   replyDate  = '';

      for (let m = 0; m < messages.length; m++) {
        const msg         = messages[m];
        const fromAddress = msg.getFrom().toLowerCase();

        if (fromAddress.indexOf(senderEmail) === -1) {
          replyFound = true;
          const msgDate = msg.getDate();
          replyDate     = msgDate instanceof Date ? formatDate(msgDate) : formatDate(new Date());
          break;
        }
      }

      if (replyFound) {
        replies.push({ threadId: threadId, replyDate: replyDate });
      }
    }

    log('handleScanReplies: found ' + replies.length + ' replies out of ' +
        threadIdsToCheck.length + ' threads checked', 'INFO');
    if (replies.length > 0) {
          const ss          = SpreadsheetApp.getActiveSpreadsheet();
          const configSheet = ss.getSheetByName(CONFIG_SHEET);
          const config      = getConfig(configSheet);
          const current     = parseInt(config.repliesReceived || 0, 10);
          updateConfigValue(configSheet, 'repliesReceived', String(current + replies.length));
        }
    return { success: true, replies: replies };

  } catch (e) {
    log('handleScanReplies: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message, replies: [] };
  }
}


// ============================================================================
// SECTION 5 — HELPER FUNCTIONS
// ============================================================================

/**
 * Removes rows with terminal statuses from an inbox sheet.
 * Reads all rows, filters out terminal ones, writes survivors back.
 * Uses deleteRow() working from bottom up to avoid index shifting.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} sheet    - Inbox sheet to clear.
 * @param {Array}                              statuses - Status values to remove.
 * @returns {number} Number of rows removed.
 */
function clearTerminalRows(sheet, statuses) {
  if (!sheet) { return 0; }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { return 0; }

  const headers   = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data      = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const statusSet = {};
  statuses.forEach(function(s) { statusSet[s] = true; });

  const col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  // Filter: keep rows that are NOT terminal
  const survivingRows = data.filter(function(row) {
    const status = String(row[col['status']] || '').trim();
    return !(status in statusSet);
  });

  const removedCount = data.length - survivingRows.length;

  if (removedCount === 0) { return 0; }

  // Clear all data rows and rewrite survivors
  sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();

  if (survivingRows.length > 0) {
    sheet.getRange(2, 1, survivingRows.length, headers.length).setValues(survivingRows);
  }

  SpreadsheetApp.flush();
  return removedCount;
}


/**
 * Increments sentToday, initialSentToday, and/or followupSentToday in
 * senderAccounts for the sender identified by senderId.
 * Called after cold or followup send updates.
 *
 * @param {number} coldCount    - Number of cold emails to add to counters.
 * @param {number} followupCount - Number of followup emails to add to counters.
 * @param {string} senderId     - emailID of the sender to update.
 * @returns {void}
 */
function incrementSenderCounters(coldCount, followupCount, senderId) {
  if (!senderId) {
    log('incrementSenderCounters: no senderId provided — skipping counter update', 'WARN');
    return;
  }

  if (coldCount === 0 && followupCount === 0) { return; }

  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);

    if (!sheet) { return; }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    for (let i = 0; i < data.length; i++) {
      const rowSenderId = String(data[i][col['emailID']] || '').trim();
      if (rowSenderId !== senderId) { continue; }

      if ('sentToday'          in col) { data[i][col['sentToday']]          = (parseInt(data[i][col['sentToday']]          || 0, 10)) + coldCount + followupCount; }
      if ('totalSentCount'     in col) { data[i][col['totalSentCount']]     = (parseInt(data[i][col['totalSentCount']]     || 0, 10)) + coldCount + followupCount; }
      if ('totalLeadsContacted' in col && coldCount > 0) { data[i][col['totalLeadsContacted']] = (parseInt(data[i][col['totalLeadsContacted']] || 0, 10)) + coldCount; }
      if ('lastRunTime'        in col) { data[i][col['lastRunTime']]        = new Date().toISOString(); }
      break;
    }

    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    SpreadsheetApp.flush();

  } catch (e) {
    log('incrementSenderCounters: error — ' + e.message, 'ERROR');
  }
}

/**
 * Handles scanBounces action from bounceScanner.gs.
 * Searches sender's Gmail for Mail Delivery Subsystem bounce notifications.
 * Returns array of { email, bounceDate } for bounced addresses.
 *
 * @param {Object} params - { senderEmail: '...' }
 * @returns {{ success: boolean, bounces: Array }}
 */
function handleScanBounces(params) {
  const senderEmail = String(params.senderEmail || '').toLowerCase();

  try {
    const scanFromDate = new Date();
    scanFromDate.setDate(scanFromDate.getDate() - WEBAPP_BOUNCE_SCAN_DAYS);
    const dateFilter   = formatDate(scanFromDate).replace(/-/g, '/');

    // Exact query matching what Google actually sends
    const query   = 'from:mailer-daemon@googlemail.com subject:"Delivery Status Notification" after:' + dateFilter;
    const threads = GmailApp.search(query, 0, 500);

    log('handleScanBounces: found ' + threads.length + ' bounce notification threads', 'INFO');

    const bounces = [];
    const seen    = {};

    for (let t = 0; t < threads.length; t++) {
      const messages = threads[t].getMessages();

      for (let m = 0; m < messages.length; m++) {
        const msg  = messages[m];
        const from = msg.getFrom().toLowerCase();

        // Only process messages from mailer-daemon
        if (from.indexOf('mailer-daemon') === -1) { continue; }

        // Get the full raw body to find the bounced email
        const body = msg.getPlainBody().toLowerCase();

        // Extract all email addresses from body
        const emailMatches = body.match(/[a-zA-Z0-9._%+\-]+@[a-zA-Z0-9.\-]+\.[a-zA-Z]{2,}/g);
        if (!emailMatches) { continue; }

        for (let e = 0; e < emailMatches.length; e++) {
          const foundEmail = emailMatches[e].toLowerCase().trim();

          // Skip system and sender emails
          if (foundEmail === senderEmail)                   { continue; }
          if (foundEmail.indexOf('mailer-daemon') !== -1)   { continue; }
          if (foundEmail.indexOf('googlemail.com') !== -1)  { continue; }
          if (foundEmail.indexOf('google.com')    !== -1)   { continue; }
          if (foundEmail.indexOf('postmaster')    !== -1)   { continue; }

          // Deduplicate
          if (seen[foundEmail]) { continue; }
          seen[foundEmail] = true;

          const msgDate    = msg.getDate();
          const bounceDate = msgDate instanceof Date
            ? formatDate(msgDate)
            : formatDate(new Date());

          bounces.push({ email: foundEmail, bounceDate: bounceDate });
        }
      }
    }

    log('handleScanBounces: extracted ' + bounces.length + ' unique bounced emails', 'INFO');
    if (bounces.length > 0) {
      const ss          = SpreadsheetApp.getActiveSpreadsheet();
      const configSheet = ss.getSheetByName(CONFIG_SHEET);
      const config      = getConfig(configSheet);
      const current     = parseInt(config.bouncedEmails || 0, 10);
      updateConfigValue(configSheet, 'bouncedEmails', String(current + bounces.length));
    }
    return { success: true, bounces: bounces };

  } catch (e) {
    log('handleScanBounces: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message, bounces: [] };
  }
}

/**
 * Fires kickoffCold on all active senders in parallel.
 * Called from Cloudflare UI "Start All Senders" button.
 */
function handleKickoffAllSenders(params) {
  return fanOutToSenders_('kickoffCold', 'handleKickoffAllSenders');
}

/**
 * Fires kickoffFollowup on all active senders in parallel.
 */
function handleKickoffAllFollowups(params) {
  return fanOutToSenders_('kickoffFollowup', 'handleKickoffAllFollowups');
}

/**
 * Pauses ALL senders immediately by setting isActive = Paused.
 * One sheet write — takes under 1 second.
 */
function handleEmergencyStop(params) {
  try {
    const count = setAllSendersStatus_('Paused');
    log('handleEmergencyStop: paused ' + count + ' senders', 'INFO');
    return { success: true, paused: count };
  } catch (e) {
    log('handleEmergencyStop: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Resumes ALL senders by setting isActive = Active.
 */
function handleResumeAll(params) {
  try {
    const count = setAllSendersStatus_('Active');
    log('handleResumeAll: resumed ' + count + ' senders', 'INFO');
    return { success: true, resumed: count };
  } catch (e) {
    log('handleResumeAll: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Fires setupTriggers on all active senders in parallel.
 * Run once when setting up new sender spreadsheets.
 */
function handleSetupAllTriggers(params) {
  return fanOutToSenders_('setupTriggers', 'handleSetupAllTriggers');
}

/**
 * Fires cleanupTriggers on all active senders in parallel.
 * Use when something is stuck — resets Sending rows and deletes orphan triggers.
 */
function handleCleanupAllTriggers(params) {
  return fanOutToSenders_('cleanupTriggers', 'handleCleanupAllTriggers');
}

/**
 * Core fan-out helper. POSTs an action to all active senders in parallel.
 * Batches in groups of 50. Returns success/fail counts and per-sender results.
 *
 * @param {string} senderAction  - Action string to send to each sender webApp.
 * @param {string} callerName    - For logging only.
 * @returns {{ success: boolean, total: number, succeeded: number,
 *             failed: number, results: Array }}
 */
function fanOutToSenders_(senderAction, callerName) {
  try {
    const senders = getActiveSenders();
    if (senders.length === 0) {
      return { success: true, total: 0, succeeded: 0, failed: 0, results: [] };
    }

    const fetchRequests = senders
      .filter(function(s) { return !!s.webAppUrl; })
      .map(function(s) {
        return {
          url:                s.webAppUrl,
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({ action: senderAction }),
          muteHttpExceptions: true
        };
      });

    const validSenders = senders.filter(function(s) { return !!s.webAppUrl; });
    const results      = [];
    let succeeded      = 0;
    let failed         = 0;

    for (let b = 0; b < fetchRequests.length; b += 50) {
      const batchReqs    = fetchRequests.slice(b, b + 50);
      const batchSenders = validSenders.slice(b, b + 50);

      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(batchReqs);
      } catch (e) {
        log(callerName + ': fetchAll error — ' + e.message, 'ERROR');
        // Mark all in this batch as failed
        batchSenders.forEach(function(s) {
          results.push({ emailID: s.emailID, success: false, error: 'fetchAll failed' });
          failed++;
        });
        continue;
      }

      for (let r = 0; r < responses.length; r++) {
        const sender = batchSenders[r];
        const code   = responses[r].getResponseCode();
        let   result = { emailID: sender.emailID, success: false };

        if (code === 200) {
          try {
            const body     = JSON.parse(responses[r].getContentText());
            result.success = body.success !== false;
            result.message = body.message || '';
            result.data    = body;
          } catch (e) {
            result.error = 'Invalid JSON response';
          }
        } else {
          result.error = 'HTTP ' + code;
        }

        if (result.success) { succeeded++; } else { failed++; }
        results.push(result);
        log(callerName + ': sender ' + sender.emailID + ' — ' +
          (result.success ? 'OK' : result.error), result.success ? 'INFO' : 'ERROR');
      }
    }

    log(callerName + ': done — ' + succeeded + ' OK, ' + failed + ' failed', 'INFO');

    return {
      success:   true,
      total:     validSenders.length,
      succeeded: succeeded,
      failed:    failed,
      results:   results
    };

  } catch (e) {
    log(callerName + ': unexpected error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Sets isActive for ALL rows in senderAccounts in one setValues() call.
 *
 * @param {string} status - 'Active' or 'Paused'
 * @returns {number} Number of rows updated.
 */
function setAllSendersStatus_(status) {
  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
  if (!sheet || sheet.getLastRow() < 2) { return 0; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const col     = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  for (let i = 0; i < data.length; i++) {
    data[i][col['isActive']] = status;
  }

  sheet.getRange(2, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();
  return data.length;
}

/**
 * Handles purgeAllInboxes action from Cloudflare UI or manual call.
 * Fans out purgeInbox to ALL active senders in parallel via fetchAll.
 * Each sender clears its own dailyInbox and followupInbox and kills orphan triggers.
 *
 * @param {Object} params - (no params needed)
 * @returns {{ success: boolean, total: number, succeeded: number,
 *             failed: number, results: Array }}
 */
function handlePurgeAllInboxes(params) {
  return fanOutToSenders_('purgeInbox', 'handlePurgeAllInboxes');
}

/**
 * Handles sendTestEmail action from Cloudflare UI or manual call.
 * Picks one sender assigned to the campaign, grabs a real lead row
 * from masterLeads for that campaign to use as variable data,
 * then POSTs to that sender's webApp to fire the actual test email.
 *
 * Flow:
 *  1. Read campaign settings — get assignedSenders, sequenceStep
 *  2. Pick first active assigned sender
 *  3. Read one real lead row from masterLeads for this campaign
 *     (uses it for variable replacement — does NOT mark it as anything)
 *  4. POST sendTestEmail to that sender's webApp
 *  5. Return result
 *
 * @param {Object} params - {
 *   campaignId:   string  — required
 *   sequenceStep: number  — optional, default 1
 *   testEmail:    string  — required, where to send the test
 *   senderId:     string  — optional, force a specific sender instead of auto-pick
 * }
 * @returns {{ success: boolean, sentTo: string, sentFrom: string,
 *             subject: string, usedLeadId: string, error?: string }}
 */
function handleSendTestEmail(params) {
  const campaignId   = String(params.campaignId   || '').trim();
  const sequenceStep = parseInt(params.sequenceStep || 1, 10);
  const testEmail    = String(params.testEmail     || '').trim();
  const forceSender  = String(params.senderId      || '').trim();

  // ── Validate ─────────────────────────────────────────────────────────────
  if (!campaignId) {
    return { success: false, error: 'Missing campaignId' };
  }
  if (!testEmail) {
    return { success: false, error: 'Missing testEmail' };
  }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;

    if (!leadsSheetId) {
      return { success: false, error: 'leadsSpreadsheetId not in Settings' };
    }

    // ── Step 1: Get campaign settings ────────────────────────────────────
    const campaignSettings = getCampaignSettings(MASTER_SHEET_ID);
    const campSettings     = campaignSettings[campaignId];

    if (!campSettings) {
      return { success: false, error: 'Campaign not found: ' + campaignId };
    }

    // ── Step 2: Pick a sender ─────────────────────────────────────────────
    // If forceSender provided — use it. Otherwise pick first active assigned sender.
    const allSenders    = getActiveSenders();
    const allSendersMap = {};
    allSenders.forEach(function(s) { allSendersMap[s.emailID] = s; });

    let chosenSender = null;

    if (forceSender) {
      chosenSender = allSendersMap[forceSender] || null;
      if (!chosenSender) {
        return { success: false, error: 'Sender not found or not active: ' + forceSender };
      }
    } else {
      // Pick first active sender assigned to this campaign
      const assignedIds = campSettings.assignedSenders || [];
      for (let i = 0; i < assignedIds.length; i++) {
        const candidate = allSendersMap[assignedIds[i]];
        if (candidate && candidate.webAppUrl) {
          chosenSender = candidate;
          break;
        }
      }
    }

    if (!chosenSender) {
      return {
        success: false,
        error:   'No active sender with webAppUrl found for campaign ' + campaignId
      };
    }

    // ── Step 3: Grab one real lead row for variable data ──────────────────
    // We read masterLeads and find the first lead belonging to this campaign.
    // We do NOT change its status — purely reading for variable data.
    const leadData   = getSampleLeadData_(leadsSheetId, campaignId);
    const usedLeadId = leadData.leadId || 'none — using empty variables';

    log('handleSendTestEmail: using leadId=' + usedLeadId +
        ' sender=' + chosenSender.emailID +
        ' campaign=' + campaignId +
        ' step=' + sequenceStep, 'INFO');

    // ── Step 4: POST to sender webApp ─────────────────────────────────────
    const payload = JSON.stringify({
      action:              'sendTestEmail',
      campaignId:          campaignId,
      sequenceStep:        sequenceStep,
      testEmail:           testEmail,
      leadData:            leadData,
      masterSpreadsheetId: MASTER_SHEET_ID
    });

    const options = {
      method:             'post',
      contentType:        'application/json',
      payload:            payload,
      muteHttpExceptions: true
    };

    const response = UrlFetchApp.fetch(chosenSender.webAppUrl, options);
    const code     = response.getResponseCode();

    if (code !== 200) {
      return {
        success: false,
        error:   'Sender webApp returned HTTP ' + code
      };
    }

    let result;
    try {
      result = JSON.parse(response.getContentText());
    } catch (e) {
      return { success: false, error: 'Invalid JSON response from sender webApp' };
    }

    if (!result.success) {
      return { success: false, error: result.error || 'Sender reported failure' };
    }

    return {
      success:     true,
      sentTo:      result.sentTo,
      sentFrom:    result.sentFrom,
      subject:     result.subject,
      campaignId:  campaignId,
      step:        sequenceStep,
      usedLeadId:  usedLeadId,
      usedSender:  chosenSender.emailID
    };

  } catch (e) {
    log('handleSendTestEmail: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}


/**
 * Reads masterLeads and returns the first lead row for a given campaignId
 * as a plain object ready for variable replacement.
 * Returns empty object with placeholder values if no lead found.
 * Never modifies any row — read only.
 *
 * @param {string} leadsSheetId
 * @param {string} campaignId
 * @returns {Object} Lead fields as { firstName, lastName, email, ... }
 */
function getSampleLeadData_(leadsSheetId, campaignId) {
  try {
    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName(LEADS_SHEET_NAME);

    if (!sheet || sheet.getLastRow() < 2) {
      return buildPlaceholderLead_();
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1,
                                   sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Find first lead belonging to this campaign — any status is fine
    // We just need variable data, not a sendable lead
    for (let i = 0; i < data.length; i++) {
      const rowCampaignId = String(data[i][col['campaignId']] || '').trim();
      if (rowCampaignId !== campaignId) { continue; }

      // Build plain object from this row
      const leadObj = {};
      headers.forEach(function(h, idx) {
        const key = String(h).trim();
        if (key) { leadObj[key] = sanitize(data[i][idx]); }
      });
      return leadObj;
    }

    // No lead found for this campaign — return placeholders
    log('getSampleLeadData_: no lead found for campaign ' + campaignId +
        ' — using placeholder data', 'WARN');
    return buildPlaceholderLead_();

  } catch (e) {
    log('getSampleLeadData_: error — ' + e.message + ' — using placeholder data', 'ERROR');
    return buildPlaceholderLead_();
  }
}


/**
 * Returns a placeholder lead object when no real lead is available.
 * Values are obviously fake so you know variables are not resolving.
 *
 * @returns {Object}
 */
function buildPlaceholderLead_() {
  return {
    leadId:          'TEST-000',
    firstName:       'John',
    lastName:        'Doe',
    email:           'john.doe@example.com',
    companyName:     'Acme Corp',
    companyWebsite:  'acmecorp.com',
    jobTitle:        'CEO',
    industry:        'Technology',
    customVar1:      '',
    customVar2:      '',
    customVar3:      '',
    customVar4:      '',
    customVar5:      ''
  };
}


function processStagingUpdates() {
  if (!acquireLock(30)) {
    log('processStagingUpdates: could not acquire lock — exiting', 'WARN');
    return;
  }

  try {
    const ss           = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const stagingSheet = ss.getSheetByName('stagingUpdates');
    if (!stagingSheet || stagingSheet.getLastRow() < 2) {
      log('processStagingUpdates: nothing to process', 'INFO');
      return;
    }

    const headers  = stagingSheet.getRange(1, 1, 1, stagingSheet.getLastColumn()).getValues()[0];
    const data     = stagingSheet.getRange(2, 1, stagingSheet.getLastRow() - 1, stagingSheet.getLastColumn()).getValues();
    const col      = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Read masterLeads ONCE
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    const lss          = SpreadsheetApp.openById(leadsSheetId);
    const lSheet       = lss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow      = lSheet.getLastRow();
    const lHeaders     = lSheet.getRange(1, 1, 1, lSheet.getLastColumn()).getValues()[0];
    const lData        = lSheet.getRange(2, 1, lastRow - 1, lSheet.getLastColumn()).getValues();
    const lCol         = {};
    lHeaders.forEach(function(h, i) { lCol[String(h).trim()] = i; });

    // Build leadId map ONCE
    const leadIdMap = {};
    for (let i = 0; i < lData.length; i++) {
      const leadId = String(lData[i][lCol['leadId']] || '').trim();
      if (leadId) { leadIdMap[leadId] = i; }
    }

    // Apply ALL staging rows to masterLeads in memory
    let updateCount = 0;
    for (let i = 0; i < data.length; i++) {
      const action   = String(data[i][col['action']]   || '').trim();
      const leadId   = String(data[i][col['leadId']]   || '').trim();
      if (!leadId || !(leadId in leadIdMap))            { continue; }

      const idx = leadIdMap[leadId];

      if (action === 'coldSent') {
        if ('status'       in lCol) { lData[idx][lCol['status']]       = data[i][col['status']]       || 'Sent'; }
        if ('threadId'     in lCol) { lData[idx][lCol['threadId']]     = data[i][col['threadId']]     || ''; }
        if ('lastSentDate' in lCol) { lData[idx][lCol['lastSentDate']] = data[i][col['sentTime']]     || ''; }
        if ('nextSentDate' in lCol) { lData[idx][lCol['nextSentDate']] = data[i][col['nextSentDate']] || ''; }
        if ('sequenceStep' in lCol) { lData[idx][lCol['sequenceStep']] = data[i][col['sequenceStep']] || 1; }
        updateCount++;
      }

      if (action === 'followupSent') {
        if ('status'         in lCol) { lData[idx][lCol['status']]         = data[i][col['status']]         || 'Sent'; }
        if ('lastSentDate'   in lCol) { lData[idx][lCol['lastSentDate']]   = data[i][col['lastSentTime']]   || ''; }
        if ('nextSentDate'   in lCol) { lData[idx][lCol['nextSentDate']]   = data[i][col['nextSentDate']]   || ''; }
        if ('sequenceStep'   in lCol) { lData[idx][lCol['sequenceStep']]   = data[i][col['sequenceStep']]   || 1; }
        if ('sequenceStatus' in lCol) { lData[idx][lCol['sequenceStatus']] = data[i][col['sequenceStatus']] || ''; }
        updateCount++;
      }
    }

    // Write masterLeads ONCE
    if (updateCount > 0) {
      lSheet.getRange(2, 1, lData.length, lHeaders.length).setValues(lData);
      SpreadsheetApp.flush();
      log('processStagingUpdates: wrote ' + updateCount + ' updates to masterLeads', 'INFO');
    }

    // Clear staging sheet — keep header only
    if (stagingSheet.getLastRow() > 1) {
      stagingSheet.getRange(2, 1, stagingSheet.getLastRow() - 1, stagingSheet.getLastColumn()).clearContent();
      SpreadsheetApp.flush();
      log('processStagingUpdates: cleared staging sheet', 'INFO');
    }

  } catch (e) {
    log('processStagingUpdates: error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}


function setupStagingProcessorTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processStagingUpdates') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  ScriptApp.newTrigger('processStagingUpdates')
    .timeBased()
    .everyMinutes(20)
    .create();
  log('setupStagingProcessorTrigger: created every-5-min trigger', 'INFO');
}

function runTestEmail() {
  const result = handleSendTestEmail({
    campaignId: 'C001',
    sequenceStep: 1,
    testEmail: 'almtsmkhaled@gmail.com'
  });
  Logger.log(JSON.stringify(result));
}


/**
 * Runs scheduleDailySends() in this sender's context.
 * Called remotely from Master Control — replaces manual trigger.
 */
function handleKickoffCold(params) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const alreadyRunning = triggers.some(function(t) {
      return t.getHandlerFunction() === 'processSingleSend';
    });
    if (alreadyRunning) {
      log('handleKickoffCold: chain already running — skipped', 'INFO');
      return { success: true, message: 'Chain already running — skipped' };
    }
    scheduleDailySends();
    return { success: true, message: 'Cold sending chain started' };
  } catch (e) {
    log('handleKickoffCold: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Runs scheduleFollowupSends() in this sender's context.
 */
function handleKickoffFollowup(params) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const alreadyRunning = triggers.some(function(t) {
      return t.getHandlerFunction() === 'processSingleFollowup';
    });
    if (alreadyRunning) {
      log('handleKickoffFollowup: chain already running — skipped', 'INFO');
      return { success: true, message: 'Chain already running — skipped' };
    }
    scheduleFollowupSends();
    return { success: true, message: 'Followup sending chain started' };
  } catch (e) {
    log('handleKickoffFollowup: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Creates scheduleDailySends and scheduleFollowupSends daily triggers.
 */
function handleSetupTriggers(params) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = 0;
    for (let i = 0; i < triggers.length; i++) {
      const fn = triggers[i].getHandlerFunction();
      if (fn === 'scheduleDailySends' || fn === 'scheduleFollowupSends') {
        ScriptApp.deleteTrigger(triggers[i]);
        deleted++;
      }
    }

    ScriptApp.newTrigger('scheduleDailySends')
      .timeBased()
      .atHour(8)
      .nearMinute(10)
      .everyDays(1)
      .create();

    ScriptApp.newTrigger('scheduleFollowupSends')
      .timeBased()
      .atHour(8)
      .nearMinute(45)
      .everyDays(1)
      .create();

    log('handleSetupTriggers: created daily triggers, deleted ' + deleted + ' old', 'INFO');
    return { success: true, message: 'Triggers created', deleted: deleted };

  } catch (e) {
    log('handleSetupTriggers: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Deletes all processSingleSend and processSingleFollowup orphan triggers.
 * Resets any Sending rows back to Queued.
 */
function handleCleanupTriggers(params) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    let deleted = 0;

    for (let i = 0; i < triggers.length; i++) {
      const fn = triggers[i].getHandlerFunction();
      if (fn === 'processSingleSend' || fn === 'processSingleFollowup') {
        ScriptApp.deleteTrigger(triggers[i]);
        deleted++;
      }
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    let resetCount = 0;

    [DAILY_INBOX_SHEET, FOLLOWUP_INBOX_SHEET].forEach(function(sheetName) {
      const sheet = ss.getSheetByName(sheetName);
      if (!sheet || sheet.getLastRow() < 2) { return; }

      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
      const col     = {};
      headers.forEach(function(h, i) { col[String(h).trim()] = i; });

      let changed = false;
      for (let i = 0; i < data.length; i++) {
        const status = String(data[i][col['status']] || '').trim();
        if (status === 'Sending') {
          data[i][col['status']]    = 'Queued';
          data[i][col['triggerId']] = '';
          changed = true;
          resetCount++;
        }
        if (data[i][col['triggerId']]) {
          data[i][col['triggerId']] = '';
          changed = true;
        }
      }

      if (changed) {
        sheet.getRange(2, 1, data.length, headers.length).setValues(data);
        SpreadsheetApp.flush();
      }
    });

    log('handleCleanupTriggers: deleted ' + deleted + ' triggers, reset ' + resetCount + ' rows', 'INFO');
    return { success: true, deleted: deleted, resetRows: resetCount };

  } catch (e) {
    log('handleCleanupTriggers: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Returns live status of this sender for the dashboard.
 */
function handleGetSenderStatus(params) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET);
    const config      = getConfig(configSheet);

    const coldInbox     = ss.getSheetByName(DAILY_INBOX_SHEET);
    const followupInbox = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);

    const coldCounts     = countRowsByStatus_(coldInbox);
    const followupCounts = countRowsByStatus_(followupInbox);

    const triggers = ScriptApp.getProjectTriggers();
    const activeTriggers = triggers
      .filter(function(t) {
        const fn = t.getHandlerFunction();
        return fn === 'processSingleSend' || fn === 'processSingleFollowup';
      })
      .map(function(t) { return t.getHandlerFunction(); });

    return {
      success:        true,
      sentToday:      parseInt(config.sentToday  || 0, 10),
      dailyLimit:     parseInt(config.dailyLimit || 25, 10),
      coldQueued:     coldCounts.Queued    || 0,
      coldSent:       coldCounts.Sent      || 0,
      coldError:      coldCounts.Error     || 0,
      followupQueued: followupCounts.Queued || 0,
      followupSent:   followupCounts.Sent   || 0,
      activeTriggers: activeTriggers.length,
      triggerNames:   activeTriggers
    };

  } catch (e) {
    log('handleGetSenderStatus: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Resets sentToday to 0 in this sender's Config sheet.
 */
function handleResetSentToday(params) {
  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET);
    updateConfigValue(configSheet, 'sentToday', '0');
    return { success: true };
  } catch (e) {
    log('handleResetSentToday: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Clears ALL data rows from dailyInbox and followupInbox. Kills orphan triggers.
 */
function handlePurgeInbox(params) {
  if (!acquireLock(SENDER_LOCK_WAIT)) {
    return { success: false, error: 'Could not acquire lock' };
  }

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const coldSheet       = ss.getSheetByName(DAILY_INBOX_SHEET);
    const clearedCold     = purgeSheetData_(coldSheet);

    const followupSheet   = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);
    const clearedFollowup = purgeSheetData_(followupSheet);

    const triggers = ScriptApp.getProjectTriggers();
    let deletedTriggers = 0;

    for (let i = 0; i < triggers.length; i++) {
      const fn = triggers[i].getHandlerFunction();
      if (fn === 'processSingleSend' || fn === 'processSingleFollowup') {
        ScriptApp.deleteTrigger(triggers[i]);
        deletedTriggers++;
      }
    }

    log('handlePurgeInbox: cleared ' + clearedCold + ' cold rows, ' +
    clearedFollowup + ' followup rows, deleted ' +
    deletedTriggers + ' triggers', 'INFO');

    // Notify Master Control to reset orphaned leads for this sender
    const purgeConfigSS    = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet      = purgeConfigSS.getSheetByName(CONFIG_SHEET);
    const config           = getConfig(configSheet);

    if (config.masterWebhookUrl && config.senderId) {
      try {
        UrlFetchApp.fetch(config.masterWebhookUrl, {
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({
            action:   'resetOrphanedLeads',
            senderId: config.senderId
          }),
          muteHttpExceptions: true
        });
        log('handlePurgeInbox: notified master to reset orphaned leads', 'INFO');
      } catch (e) {
        log('handlePurgeInbox: could not notify master — ' + e.message, 'WARN');
      }
    }

    return {
      success:         true,
      clearedCold:     clearedCold,
      clearedFollowup: clearedFollowup,
      deletedTriggers: deletedTriggers
    };

  } catch (e) {
    log('handlePurgeInbox: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

/**
 * Clears all data rows from a sheet while preserving the header row.
 */
function purgeSheetData_(sheet) {
  if (!sheet) {
    log('purgeSheetData_: sheet not found — skipping', 'WARN');
    return 0;
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    log('purgeSheetData_: ' + sheet.getName() + ' already empty — skipping', 'INFO');
    return 0;
  }

  const lastCol    = sheet.getLastColumn();
  const rowsToWipe = lastRow - 1;

  sheet.getRange(2, 1, rowsToWipe, lastCol).clearContent();
  SpreadsheetApp.flush();

  log('purgeSheetData_: cleared ' + rowsToWipe + ' rows from ' + sheet.getName(), 'INFO');
  return rowsToWipe;
}

/**
 * Counts rows in an inbox sheet grouped by status.
 */
function countRowsByStatus_(sheet) {
  const counts = {};
  if (!sheet || sheet.getLastRow() < 2) { return counts; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
  const col     = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  for (let i = 0; i < data.length; i++) {
    const status = String(data[i][col['status']] || '').trim();
    counts[status] = (counts[status] || 0) + 1;
  }
  return counts;
}

/**
 * Sender-side sendTestEmail — actually fires the email via GmailApp.
 * Called by Master Control's handleSendTestEmail after it picks a sender.
 */
function handleSendTestEmailDirect(params) {
  const campaignId          = String(params.campaignId    || '').trim();
  const sequenceStep        = parseInt(params.sequenceStep || 1, 10);
  const testEmail           = String(params.testEmail      || '').trim();
  const leadData            = params.leadData              || {};
  const masterSpreadsheetId = String(params.masterSpreadsheetId || '').trim();

  if (!campaignId)          { return { success: false, error: 'Missing campaignId' }; }
  if (!testEmail)           { return { success: false, error: 'Missing testEmail' }; }
  if (!masterSpreadsheetId) { return { success: false, error: 'Missing masterSpreadsheetId' }; }

  try {
    const ss          = SpreadsheetApp.getActiveSpreadsheet();
    const configSheet = ss.getSheetByName(CONFIG_SHEET);
    const config      = getConfig(configSheet);

    const templateCache = getCampaignTemplates(masterSpreadsheetId);
    const cacheKey      = campaignId + '_' + sequenceStep;
    const template      = templateCache[cacheKey] || null;

    if (!template) {
      return {
        success: false,
        error:   'No template found for campaignId=' + campaignId +
                 ' sequenceStep=' + sequenceStep
      };
    }

    const senderObj = {
      senderFirstName: config.senderFirstName || '',
      senderLastName:  config.senderLastName  || ''
    };

    const resolvedSubject = replaceVariables(template.subject,  leadData, senderObj);
    const resolvedBody    = replaceVariables(template.htmlBody, leadData, senderObj);

    const finalSubject = 'TEST ' + resolvedSubject;

    GmailApp.sendEmail(testEmail, finalSubject, resolvedBody, {
      name:    (config.senderFirstName + ' ' + (config.senderLastName || '')).trim(),
      replyTo: config.replyToEmail || config.senderEmail || ''
    });

    log('handleSendTestEmailDirect: sent to ' + testEmail +
        ' campaign=' + campaignId + ' step=' + sequenceStep, 'INFO');

    return {
      success:    true,
      sentTo:     testEmail,
      sentFrom:   config.senderEmail || 'unknown',
      subject:    finalSubject,
      campaignId: campaignId,
      step:       sequenceStep
    };

  } catch (e) {
    log('handleSendTestEmailDirect: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}