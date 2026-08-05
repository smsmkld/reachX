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
    if (action === 'systemBoot') {
      return buildResponse(handleSystemBoot(params), corsHeaders);
    }

    if (action === 'updateSender') {
     return buildResponse(handleUpdateSender(params), corsHeaders); 
    }

    if (action === 'archiveCampaignLeads') {
     return buildResponse(handleArchiveCampaignLeads(params), corsHeaders); 
    }
    if (action === 'getSettings') {
      return buildResponse(handleGetSettings(params), corsHeaders);
    }

    if (action === 'updateSettings') {
      return buildResponse(handleUpdateSettings(params), corsHeaders);
    }

    if (action === 'createCampaign') {
      return buildResponse(handleCreateCampaign(params), corsHeaders);
    }

    if (action === 'distributeCold') {
      return buildResponse(handleRunDistributeCold(params), corsHeaders);
    }

    if (action === 'distributeFollowups') {
      return buildResponse(handleRunDistributeFollowups(params), corsHeaders);
    }

    if (action === 'getCampaigns') { 
      return buildResponse(handleGetCampaigns(params), corsHeaders); 
    }

    if (action === 'updateCampaign') { 
      return buildResponse(handleUpdateCampaign(params), corsHeaders); 
    }

    if (action === 'deleteCampaign') { 
      return buildResponse(handleDeleteCampaign(params), corsHeaders); 
    }

    if (action === 'getSenders') {
      return buildResponse(handleGetSenders(params), corsHeaders); 
    }

    if (action === 'getDailySummary') { 
      return buildResponse(handleGetDailySummary(params), corsHeaders); 
    }

    if (action === 'getLeads') { 
      return buildResponse(handleGetLeads(params), corsHeaders); 
    }

    if (action === 'getSystemStatus') { 
      return buildResponse(handleGetSystemStatus(params), corsHeaders); 
    }

    if (action === 'healthCheck') {
      return buildResponse(handleHealthCheck(params), corsHeaders);
    }

    if (action === 'getCampaignStats') {
      return buildResponse(handleGetCampaignStats(params), corsHeaders);
    }

    if (action === 'getMasterTriggers') {
      return buildResponse(handleGetMasterTriggers(params), corsHeaders);
    }
    if (action === 'updateSenderConfig') {
      return buildResponse(handleUpdateSenderConfig(params), corsHeaders);
    }
    if (action === 'runWarmupIncrease') {
      return buildResponse(handleRunWarmupIncrease(params), corsHeaders);
    }
    if (action === 'getSendersByTag') { 
      return buildResponse(handleGetSendersByTag(params), corsHeaders);
    }
    if (action === 'getAllTags') { 
      return buildResponse(handleGetAllTags(params), corsHeaders);
    }
    if (action === 'assignTag') { 
      return buildResponse(handleAssignTag(params), corsHeaders);
      }
    if (action === 'removeTag') { 
      return buildResponse(handleRemoveTag(params), corsHeaders);
      }
    if (action === 'updateSendersByTag') { return buildResponse(handleUpdateSendersByTag(params), corsHeaders); 
    }
    if (action === 'addToBlocklist') { 
      return buildResponse(handleAddToBlocklist(params), corsHeaders);
    }
    if (action === 'removeFromBlocklist') {     return buildResponse(handleRemoveFromBlocklist(params), corsHeaders);
    }
    if (action === 'getBlocklist') { 
      return buildResponse(handleGetBlocklist(params), corsHeaders); 
    }
    if (action === 'getFullSystemReport') {
      return buildResponse(handleGetFullSystemReport(params), corsHeaders);
    }
    if (action === 'validateSystem') {
      return buildResponse(handleValidateSystem(params), corsHeaders);
    }
    if (action === 'checkSenderChains') {
      return buildResponse(handleCheckSenderChains(params), corsHeaders);
    }
    if (action === 'restartSenderChain') {
      return buildResponse(handleRestartSenderChain(params), corsHeaders);
    }
    if (action === 'syncStats') {
      return buildResponse(handleSyncStats(params), corsHeaders);
    }
    if (action === 'sendDailyDigest') {
      return buildResponse(handleSendDailyDigest(params), corsHeaders);
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
      return buildResponse(handleSendTestEmail(params), corsHeaders);
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
    if (action === 'resetOrphanedLeads') {
      return buildResponse(handleResetOrphanedLeads(params), corsHeaders);
    }

    if (action === 'uploadLeads') {
      return buildResponse(handleUploadLeads(params), corsHeaders);
    }

    // Unknown action
    return buildResponse({ success: false, error: 'Unknown action: ' + action }, corsHeaders);

  } catch (e) {
    log('doPost: uncaught error — ' + e.message, 'ERROR');
    return buildResponse({ success: false, error: 'Internal server error: ' + e.message }, corsHeaders);
  }
}

/**
 * Answers GET requests with JSON instead of failing.
 *
 * Without it, every browser visit to /exec is a failed execution in the log.
 * It deliberately exposes nothing: no spreadsheet IDs, no settings, no data —
 * only enough to confirm the deployment is reachable and serving JSON.
 */
function doGet(e) {
  return ContentService
    .createTextOutput(JSON.stringify({
      success: true,
      service: 'reachx-master-control',
      message: 'This endpoint is alive. Actions are POST-only.',
      time:    new Date().toISOString()
    }))
    .setMimeType(ContentService.MimeType.JSON);
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

  // MUST hold the lock: getLastRow() + setValues() is a read-modify-write, and at
  // 200-300 senders several senders POST here at the same time. Without the lock
  // two executions compute the same startRow and one silently overwrites the
  // other, and processStagingUpdates() can clear rows appended after it read.
  // A false here makes the sender cache the batch and retry (savePendingUpdates_).
  if (!acquireLock(30)) {
    log('handleUpdateColdSent: could not acquire lock — sender will retry', 'WARN');
    return { success: false, error: 'Busy — could not acquire lock' };
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
  } finally {
    releaseLock();
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

  // Same concurrent read-modify-write hazard as handleUpdateColdSent — see note there.
  if (!acquireLock(30)) {
    log('handleUpdateFollowupSent: could not acquire lock — sender will retry', 'WARN');
    return { success: false, error: 'Busy — could not acquire lock' };
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
  } finally {
    releaseLock();
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

    let matchedIndex = -1;

    for (let i = 0; i < data.length; i++) {
      const rowSenderId = String(data[i][col['emailID']] || '').trim();
      if (rowSenderId !== senderId) { continue; }

      if ('sentToday'          in col) { data[i][col['sentToday']]          = (parseInt(data[i][col['sentToday']]          || 0, 10)) + coldCount + followupCount; }
      if ('totalSentCount'     in col) { data[i][col['totalSentCount']]     = (parseInt(data[i][col['totalSentCount']]     || 0, 10)) + coldCount + followupCount; }
      if ('totalLeadsContacted' in col && coldCount > 0) { data[i][col['totalLeadsContacted']] = (parseInt(data[i][col['totalLeadsContacted']] || 0, 10)) + coldCount; }
      if ('lastRunTime'        in col) { data[i][col['lastRunTime']]        = new Date().toISOString(); }
      matchedIndex = i;
      break;
    }

    if (matchedIndex === -1) {
      log('incrementSenderCounters: sender ' + senderId + ' not found in senderAccounts', 'WARN');
      return;
    }

    // Write ONLY this sender's row. Writing back the whole array would push a
    // snapshot taken at read time over the top of every other sender's row —
    // at 200-300 senders that reverts other senders' counters and any UI edit
    // (dailyLimit, isActive, tags) made while this request was in flight.
    sheet.getRange(matchedIndex + 2, 1, 1, headers.length).setValues([data[matchedIndex]]);
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
  // Step 1: Purge all sender inboxes
  const result = fanOutToSenders_('purgeInbox', 'handlePurgeAllInboxes');

    // Step 2: Reset all Queued/Sending leads back to Pending in masterLeads.
  // Needs the lock — this rewrites the whole sheet, and processStagingUpdates
  // runs every 10 minutes, so without it one silently overwrites the other.
  if (!acquireLock(30)) {
    log('handlePurgeAllInboxes: could not acquire lock — inboxes were purged but ' +
        'masterLeads was NOT reset', 'ERROR');
    result.resetLeads = false;
    result.resetError = 'Could not acquire lock — run purgeAllInboxes again in a minute';
    return result;
  }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    const ss           = SpreadsheetApp.openById(leadsSheetId);
    const sheet        = ss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow      = sheet.getLastRow();

    if (lastRow > 1) {
      const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
      const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
      const col     = {};
      headers.forEach(function(h, i) { col[String(h).trim()] = i; });

      let resetCount = 0;
      for (let i = 0; i < data.length; i++) {
        const status = String(data[i][col['status']] || '').trim();
        if (status === 'Queued' || status === 'Sending') {
          const currentStep  = parseInt(data[i][col['sequenceStep']]       || 0, 10);
          const totalSteps   = parseInt(data[i][col['totalSequenceSteps']] || 1, 10);
          const lastSentDate = String(data[i][col['lastSentDate']]         || '').trim();

          if (!lastSentDate || currentStep === 0) {
            // Never sent — full reset, clear owner so distributor reassigns
            data[i][col['status']]       = 'Pending';
            data[i][col['ownerSender']]  = '';
            data[i][col['sequenceStep']] = 0;
          } else if (currentStep >= totalSteps) {
            // Was queued for final step — step back, KEEP owner
            data[i][col['status']]       = 'Sent';
            data[i][col['sequenceStep']] = currentStep - 1;
            // ownerSender stays — followup must go to same sender
          } else {
            // Mid sequence — the queued step was never actually sent, so step
            // back to the last step that WAS sent. Without this the next
            // distribution run queues the FOLLOWING step and that email is
            // skipped for this lead entirely.
            data[i][col['status']]       = 'Sent';
            data[i][col['sequenceStep']] = currentStep - 1;
            // ownerSender stays — followup must go to same sender
          }
          resetCount++;
        }
      }

      if (resetCount > 0) {
        sheet.getRange(2, 1, data.length, headers.length).setValues(data);
        SpreadsheetApp.flush();
        log('handlePurgeAllInboxes: reset ' + resetCount + ' leads to Pending in masterLeads', 'INFO');
      }
    }

    result.resetLeads = true;
  } catch (e) {
    log('handlePurgeAllInboxes: error resetting masterLeads — ' + e.message, 'ERROR');
    result.resetLeads = false;
    result.resetError = e.message;
  } finally {
    releaseLock();
  }

  return result;
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
    .everyMinutes(10)
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

    const finalSubject = '[TEST] ' + resolvedSubject;

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

function handleResetOrphanedLeads(params) {
  const senderId = String(params.senderId || '').trim();
    if (!senderId) {
    return { success: false, error: 'Missing senderId' };
  }

  // Rewrites the whole masterLeads sheet — must hold the lock so it cannot
  // collide with processStagingUpdates or a distribution run.
  if (!acquireLock(30)) {
    return { success: false, error: 'Busy — could not acquire lock, try again' };
  }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    const ss           = SpreadsheetApp.openById(leadsSheetId);
    const sheet        = ss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow      = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, reset: 0 }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    let resetCount = 0;
    for (let i = 0; i < data.length; i++) {
      const rowSender = String(data[i][col['ownerSender']] || '').trim();
      const status    = String(data[i][col['status']]      || '').trim();

      if (rowSender === senderId && 
        (status === 'Queued' || status === 'Sending')) {
        const currentStep  = parseInt(data[i][col['sequenceStep']]       || 0, 10);
        const totalSteps   = parseInt(data[i][col['totalSequenceSteps']] || 1, 10);
        const lastSentDate = String(data[i][col['lastSentDate']]         || '').trim();

        if (!lastSentDate || currentStep === 0) {
          // Never sent — full reset, clear owner so distributor reassigns
          data[i][col['status']]       = 'Pending';
          data[i][col['ownerSender']]  = '';
          data[i][col['sequenceStep']] = 0;
        } else if (currentStep >= totalSteps) {
          // Was queued for final step — step back, KEEP owner
          data[i][col['status']]       = 'Sent';
          data[i][col['sequenceStep']] = currentStep - 1;
          // ownerSender stays — followup must go to same sender
        } else {
          // Mid sequence — the queued step was never actually sent, so step back
          // to the last step that WAS sent, or that email gets skipped.
          data[i][col['status']]       = 'Sent';
          data[i][col['sequenceStep']] = currentStep - 1;
          // ownerSender stays — followup must go to same sender
        }
        resetCount++;
      }
    }

    if (resetCount > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }

    log('handleResetOrphanedLeads: reset ' + resetCount + 
        ' leads for sender ' + senderId, 'INFO');
    return { success: true, reset: resetCount };

    } catch (e) {
    log('handleResetOrphanedLeads: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

function handleSystemBoot(params) {
  const results = {};

  try {
    // Step 1 — Setup all master triggers
    log('systemBoot: step 1 — setting up master triggers', 'INFO');
    setupAllMasterTriggers();
    results.masterTriggers = true;

    // Step 2 — Setup triggers on all senders
    log('systemBoot: step 2 — setting up sender triggers', 'INFO');
    const triggerResult = fanOutToSenders_('setupTriggers', 'systemBoot');
    results.senderTriggers = triggerResult;

    log('systemBoot: all triggers set up successfully', 'INFO');
    return { success: true, results: results };

  } catch(e) {
    log('systemBoot: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message, results: results };
  }
}

function handleUploadLeads(params) {
  const campaignId     = String(params.campaignId     || '').trim();
  const leads          = params.leads                 || [];
  const skipExisting   = params.skipExisting   !== false;
  const skipOtherCamps = params.skipOtherCamps !== false;
  // Load blocklist once
  const blocklist      = getBlocklist_();
  const blockedEmails  = blocklist.blockedEmails;
  const blockedDomains = blocklist.blockedDomains;
  const skippedBlocklist = [];

  if (!campaignId) { return { success: false, error: 'Missing campaignId' }; }
  if (leads.length === 0) { return { success: false, error: 'No leads provided' }; }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    const ss           = SpreadsheetApp.openById(leadsSheetId);
    const sheet        = ss.getSheetByName(LEADS_SHEET_NAME);
    const headers      = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col          = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const lastRow = sheet.getLastRow();

    // ── Build existing email map + find last leadId ───────────────────────
    const existingEmails = {};
    let   lastIdNum      = 0;

    if (lastRow > 1) {
      const existingData = sheet
        .getRange(2, 1, lastRow - 1, sheet.getLastColumn())
        .getValues();

      for (let i = 0; i < existingData.length; i++) {
        const email        = String(existingData[i][col['email']]      || '').trim().toLowerCase();
        const existingCamp = String(existingData[i][col['campaignId']] || '').trim();
        const idStr        = String(existingData[i][col['leadId']]     || '').trim();

        if (email) { existingEmails[email] = existingCamp; }

        if (idStr.startsWith('L')) {
          const num = parseInt(idStr.substring(1), 10);
          if (!isNaN(num) && num > lastIdNum) { lastIdNum = num; }
        }
      }
    }

    // ── Get totalSequenceSteps from campaign ─────────────────────────────
    const campSettings = getCampaignSettings(MASTER_SHEET_ID);
    const camp         = campSettings[campaignId];
    const totalSteps   = camp ? camp.totalSequenceSteps : 1;

    // ── Filter leads ─────────────────────────────────────────────────────
    const skippedDuplicate = [];
    const skippedOtherCamp = [];
    const skippedNoEmail   = [];
    const toInsert         = [];

    for (let i = 0; i < leads.length; i++) {
      const lead  = leads[i];
      const email = String(lead.email || '').trim().toLowerCase();

      if (!email) {
        skippedNoEmail.push(lead);
        continue;
      }

      // ── Blocklist check ──────────────────────────────────────────────────
      const emailDomain = email.split('@')[1] || '';
      if (blockedEmails[email]) {
        skippedBlocklist.push(email);
        continue;
      }
      if (blockedDomains[emailDomain]) {
        skippedBlocklist.push(email);
        continue;
      }
      // ────────────────────────────────────────────────────────────────────

      if (email in existingEmails) {
        const existingCamp = existingEmails[email];
        if (skipExisting) {
          skippedDuplicate.push(email);
          continue;
        }
        if (skipOtherCamps && existingCamp !== campaignId) {
          skippedOtherCamp.push(email);
          continue;
        }
      }

      toInsert.push(lead);
    }

    // ── Build rows ────────────────────────────────────────────────────────
    const rows = toInsert.map(function(lead, index) {
      const row    = new Array(headers.length).fill('');
      const leadId = 'L' + String(lastIdNum + index + 1).padStart(4, '0');

      if ('leadId'             in col) { row[col['leadId']]             = leadId; }
      if ('campaignId'         in col) { row[col['campaignId']]         = campaignId; }
      if ('campaignStatus'     in col) { row[col['campaignStatus']]     = 'Active'; }
      if ('status'             in col) { row[col['status']]             = 'Pending'; }
      if ('sequenceStep'       in col) { row[col['sequenceStep']]       = 0; }
      if ('totalSequenceSteps' in col) { row[col['totalSequenceSteps']] = totalSteps; }
      if ('createdDate'        in col) { row[col['createdDate']]        = formatDate(new Date()); }
      if ('email'              in col) { row[col['email']]              = sanitize(lead.email          || ''); }
      if ('firstName'          in col) { row[col['firstName']]          = sanitize(lead.firstName      || ''); }
      if ('lastName'           in col) { row[col['lastName']]           = sanitize(lead.lastName       || ''); }
      if ('companyName'        in col) { row[col['companyName']]        = sanitize(lead.companyName    || ''); }
      if ('companyWebsite'     in col) { row[col['companyWebsite']]     = sanitize(lead.companyWebsite || ''); }
      if ('jobTitle'           in col) { row[col['jobTitle']]           = sanitize(lead.jobTitle       || ''); }
      if ('industry'           in col) { row[col['industry']]           = sanitize(lead.industry       || ''); }
      if ('department'         in col) { row[col['department']]         = sanitize(lead.department     || ''); }
      if ('phoneNumber'        in col) { row[col['phoneNumber']]        = sanitize(lead.phoneNumber    || ''); }
      if ('city'               in col) { row[col['city']]               = sanitize(lead.city           || ''); }
      if ('country'            in col) { row[col['country']]            = sanitize(lead.country        || ''); }
      if ('addressLine'        in col) { row[col['addressLine']]        = sanitize(lead.addressLine        || ''); }
      if ('state'              in col) { row[col['state']]              = sanitize(lead.state              || ''); }
      if ('countryCode'        in col) { row[col['countryCode']]        = sanitize(lead.countryCode        || ''); }
      if ('linkedinPersonUrl'  in col) { row[col['linkedinPersonUrl']]  = sanitize(lead.linkedinPersonUrl  || ''); }
      if ('linkedinCompanyUrl' in col) { row[col['linkedinCompanyUrl']] = sanitize(lead.linkedinCompanyUrl || ''); }
      if ('customVar1'         in col) { row[col['customVar1']]         = sanitize(lead.customVar1     || ''); }
      if ('customVar2'         in col) { row[col['customVar2']]         = sanitize(lead.customVar2     || ''); }
      if ('customVar3'         in col) { row[col['customVar3']]         = sanitize(lead.customVar3     || ''); }
      if ('customVar4'         in col) { row[col['customVar4']]         = sanitize(lead.customVar4     || ''); }
      if ('customVar5'         in col) { row[col['customVar5']]         = sanitize(lead.customVar5     || ''); }

      return row;
    });

    // ── Write all in ONE call ─────────────────────────────────────────────
    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, headers.length).setValues(rows);
      SpreadsheetApp.flush();
    }

    log('handleUploadLeads: inserted=' + rows.length +
        ' skippedDuplicate=' + skippedDuplicate.length +
        ' skippedOtherCamp=' + skippedOtherCamp.length +
        ' skippedNoEmail='   + skippedNoEmail.length, 'INFO');

  // Reflect the upload now rather than at 11:30 PM.
    try {
      syncCampaignStats(MASTER_SHEET_ID, leadsSheetId);
    } catch (e) {
      log('handleUploadLeads: stats resync failed — ' + e.message, 'WARN');
    }


    return {
      success:          true,
      campaignId:       campaignId,
      inserted:         rows.length,
      skippedDuplicate: skippedDuplicate.length,
      skippedOtherCamp: skippedOtherCamp.length,
      skippedNoEmail:   skippedNoEmail.length,
      skippedBlocklist: skippedBlocklist.length,
      totalReceived:    leads.length
    };

  } catch(e) {
    log('handleUploadLeads: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleGetCampaigns(params) {
  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName('Campaigns');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, campaigns: [] }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Group by campaignId — one object per campaign not per step
    const campaignMap = {};
    for (let i = 0; i < data.length; i++) {
      const campaignId = String(data[i][col['campaignId']] || '').trim();
      if (!campaignId) { continue; }

      if (!campaignMap[campaignId]) {
        campaignMap[campaignId] = {
          campaignId:          campaignId,
          campaignName:        String(data[i][col['campaignName']]        || '').trim(),
          campaignStatus:      String(data[i][col['campaignStatus']]      || '').trim(),
          totalSequenceSteps:  parseInt(data[i][col['totalSequenceSteps']] || 1, 10),
          assignedSenders:     String(data[i][col['assignedSenders']]     || '').trim(),
          windowStart:         formatTimeValue_(data[i][col['windowStart']]),
          windowEnd:           formatTimeValue_(data[i][col['windowEnd']]),
          sendingDays:         String(data[i][col['sendingDays']]         || '').trim(),
          timezone:            String(data[i][col['timezone']]            || '').trim(),
          totalLeads:          parseInt(data[i][col['totalLeads']]        || 0, 10),
          sentCount:           parseInt(data[i][col['sentCount']]         || 0, 10),
          replyCount:          parseInt(data[i][col['replyCount']]        || 0, 10),
          pendingCount:        parseInt(data[i][col['pendingCount']]      || 0, 10),
          bouncedEmails:       parseInt(data[i][col['bouncedEmails']]     || 0, 10),
          stepsSent:           parseInt(data[i][col['stepsSent']]         || 0, 10),
          contactedCount:      parseInt(data[i][col['contactedCount']]    || 0, 10),
          createdDate:         String(data[i][col['createdDate']]         || '').trim(),
          notes:                 String(data[i][col['notes']]                || '').trim(),
          coldPriorityLimit:     parseFloat(data[i][col['coldPriorityLimit']])     || null,
          followupPriorityLimit: parseFloat(data[i][col['followupPriorityLimit']]) || null,
          notes:               String(data[i][col['notes']]               || '').trim(),
          steps:               []
        };
      }

      // Add this step's copy
      campaignMap[campaignId].steps.push({
        sequenceStep: parseInt(data[i][col['sequenceStep']] || 1, 10),
        subjectLine:  String(data[i][col['subjectLine']]    || '').trim(),
        emailBody:    String(data[i][col['emailBody']]      || '').trim(),
        awaitDays:    parseInt(data[i][col['awaitDays']]    || 0, 10)
      });
    }

    const campaigns = Object.values(campaignMap);
    campaigns.sort(function(a, b) {
      return a.campaignId.localeCompare(b.campaignId);
    });

    return { success: true, campaigns: campaigns };

  } catch(e) {
    log('handleGetCampaigns: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleCreateCampaign(params) {
  const steps = params.steps || [];
  if (!Array.isArray(steps) || steps.length === 0) {
    return { success: false,
             error: 'steps is required — [{ subjectLine, emailBody, awaitDays }, ...]' };
  }

  if (!acquireLock(30)) {
    return { success: false, error: 'Busy — could not acquire lock' };
  }

  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName('Campaigns');
    if (!sheet) { return { success: false, error: 'Campaigns sheet not found' }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    if (!('campaignId' in col) || !('sequenceStep' in col)) {
      return { success: false,
               error: 'Campaigns sheet is missing a campaignId or sequenceStep column' };
    }

    // Generate the next id: C001, C002, ...
    let maxNum    = 0;
    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const existing = sheet.getRange(2, col['campaignId'] + 1, lastRow - 1, 1).getValues();
      for (let i = 0; i < existing.length; i++) {
        const m = String(existing[i][0] || '').trim().match(/^C(\d+)$/i);
        if (m) { maxNum = Math.max(maxNum, parseInt(m[1], 10)); }
      }
    }
    const campaignId = 'C' + String(maxNum + 1).padStart(3, '0');

    // One row per step
    const today = formatDate(new Date());
    const rows  = [];

    for (let s = 0; s < steps.length; s++) {
      const row = new Array(headers.length).fill('');
      const set = function(name, value) {
        if (name in col) { row[col[name]] = value; }
      };

      set('campaignId',         campaignId);
      set('campaignName',       params.campaignName || campaignId);
      set('sequenceStep',       s + 1);
      set('totalSequenceSteps', steps.length);
      set('campaignStatus',     params.campaignStatus || 'Paused');
      set('assignedSenders',    params.assignedSenders || '');
      set('windowStart',        params.windowStart     || '');
      set('windowEnd',          params.windowEnd       || '');
      set('sendingDays',        params.sendingDays     || '');
      set('timezone',           params.timezone        || '');
      set('subjectLine',        steps[s].subjectLine   || '');
      set('emailBody',          steps[s].emailBody     || '');
      set('awaitDays',          parseInt(steps[s].awaitDays || 0, 10));
      set('createdDate',        today);
      set('lastModified',       today);
      set('totalLeads',    0);
      set('sentCount',     0);
      set('replyCount',    0);
      set('pendingCount',  0);
      set('bouncedEmails', 0);
      if (params.coldPriorityLimit     !== undefined) { set('coldPriorityLimit',     params.coldPriorityLimit); }
      if (params.followupPriorityLimit !== undefined) { set('followupPriorityLimit', params.followupPriorityLimit); }

      rows.push(row);
    }

    sheet.getRange(sheet.getLastRow() + 1, 1, rows.length, headers.length).setValues(rows);
    SpreadsheetApp.flush();

    log('handleCreateCampaign: created ' + campaignId + ' with ' + rows.length + ' steps', 'INFO');
    return { success: true, campaignId: campaignId, rowsCreated: rows.length };

  } catch (e) {
    log('handleCreateCampaign: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

function handleRunDistributeCold(params) {
  try {
    ScriptApp.newTrigger('runDistributeColdNow').timeBased().after(5000).create();
    log('handleRunDistributeCold: queued', 'INFO');
    return { success: true, message: 'Cold distribution queued — starts within about a minute' };
  } catch (e) {
    log('handleRunDistributeCold: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleRunDistributeFollowups(params) {
  try {
    ScriptApp.newTrigger('runDistributeFollowupsNow').timeBased().after(5000).create();
    log('handleRunDistributeFollowups: queued', 'INFO');
    return { success: true, message: 'Followup distribution queued — starts within about a minute' };
  } catch (e) {
    log('handleRunDistributeFollowups: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleUpdateCampaign(params) {
  const campaignId = String(params.campaignId || '').trim();
  if (!campaignId) { return { success: false, error: 'Missing campaignId' }; }

  if (!acquireLock(30)) {
    return { success: false, error: 'Busy — could not acquire lock' };
  }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName('Campaigns');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: false, error: 'No campaigns found' }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Split the sheet into this campaign's rows and everyone else's.
    const mine   = [];
    const others = [];
    for (let i = 0; i < data.length; i++) {
      const rowId = String(data[i][col['campaignId']] || '').trim();
      (rowId === campaignId ? mine : others).push(data[i]);
    }
    if (mine.length === 0) {
      return { success: false, error: 'Campaign not found: ' + campaignId };
    }

    mine.sort(function(a, b) {
      return (parseInt(a[col['sequenceStep']] || 1, 10)) -
             (parseInt(b[col['sequenceStep']] || 1, 10));
    });

    const today = formatDate(new Date());

    // ── Decide the final step count ──────────────────────────────────────
    // TWO DIFFERENT CONTRACTS, deliberately:
    //
    //   params.steps    — MERGE. Edits the steps you name, matched by
    //                     sequenceStep. Never changes how many steps exist.
    //                     Sending one step edits that step and nothing else.
    //
    //   params.setSteps — REPLACE. The array is the complete desired
    //                     sequence: extra entries add rows, missing ones
    //                     DELETE rows. Only send this when the user has
    //                     explicitly asked to add or remove a step.
    //
    // Keeping these separate matters: an editor that saves one step at a
    // time would silently destroy the rest of the sequence under a single
    // authoritative contract.
    let finalRows;

    if (Array.isArray(params.setSteps) && params.setSteps.length > 0) {
      const wanted = params.setSteps.slice().sort(function(a, b) {
        return (a.sequenceStep || 0) - (b.sequenceStep || 0);
      });

      finalRows = wanted.map(function(stepData, idx) {
        // Reuse the existing row for this step so its copy is preserved
        // when the caller sends only some fields.
        const existing = mine[idx];
        const row = existing
          ? existing.slice()
          : (mine[0] ? mine[0].slice() : new Array(headers.length).fill(''));

        const set = function(name, value) {
          if (name in col) { row[col[name]] = value; }
        };

        set('sequenceStep', idx + 1);
        if (!existing) {
          // Brand new step — start empty rather than cloning step 1's copy.
          set('subjectLine', '');
          set('emailBody',   '');
          set('awaitDays',   3);
          set('createdDate', today);
        }
        if (stepData.subjectLine !== undefined) { set('subjectLine', stepData.subjectLine); }
        if (stepData.emailBody   !== undefined) { set('emailBody',   stepData.emailBody); }
        if (stepData.awaitDays   !== undefined) { set('awaitDays',   parseInt(stepData.awaitDays, 10) || 0); }

        return row;
      });

    } else {
      // Merge path: keep every existing row, edit only the ones named in
      // params.steps. A step the caller did not mention is left untouched.
      finalRows = mine.map(function(r) { return r.slice(); });

      if (Array.isArray(params.steps)) {
        finalRows.forEach(function(row) {
          const rowStep = parseInt(row[col['sequenceStep']] || 1, 10);
          const stepData = params.steps.find(function(s) {
            return parseInt(s.sequenceStep, 10) === rowStep;
          });
          if (!stepData) { return; }

          const set = function(name, value) {
            if (name in col) { row[col[name]] = value; }
          };
          if (stepData.subjectLine !== undefined) { set('subjectLine', stepData.subjectLine); }
          if (stepData.emailBody   !== undefined) { set('emailBody',   stepData.emailBody); }
          if (stepData.awaitDays   !== undefined) { set('awaitDays',   parseInt(stepData.awaitDays, 10) || 0); }
        });
      }
    }

    // ── Apply campaign-level fields to every row ─────────────────────────
    const totalSteps = finalRows.length;

    finalRows.forEach(function(row) {
      const set = function(name, value) {
        if (name in col) { row[col[name]] = value; }
      };
      if (params.campaignName          !== undefined) { set('campaignName',          params.campaignName); }
      if (params.campaignStatus        !== undefined) { set('campaignStatus',        params.campaignStatus); }
      if (params.windowStart           !== undefined) { set('windowStart',           params.windowStart); }
      if (params.windowEnd             !== undefined) { set('windowEnd',             params.windowEnd); }
      if (params.sendingDays           !== undefined) { set('sendingDays',           params.sendingDays); }
      if (params.timezone              !== undefined) { set('timezone',              params.timezone); }
      if (params.assignedSenders       !== undefined) { set('assignedSenders',       params.assignedSenders); }
      if (params.notes                 !== undefined) { set('notes',                 params.notes); }
      if (params.coldPriorityLimit     !== undefined) { set('coldPriorityLimit',     params.coldPriorityLimit); }
      if (params.followupPriorityLimit !== undefined) { set('followupPriorityLimit', params.followupPriorityLimit); }
      set('totalSequenceSteps', totalSteps);
      set('lastModified', today);
    });

    // ── Rewrite the sheet ────────────────────────────────────────────────
    const out = others.concat(finalRows);
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    if (out.length > 0) {
      sheet.getRange(2, 1, out.length, headers.length).setValues(out);
    }
    SpreadsheetApp.flush();

    // ── Keep existing leads in step with the campaign ────────────────────
    // Without this, leads keep the old totalSequenceSteps and either stop
    // early or chase a step that no longer exists.
    let leadsUpdated = 0;
    if (mine.length !== totalSteps) {
      leadsUpdated = syncLeadTotalSteps_(campaignId, totalSteps);
    }

    if (params.campaignStatus !== undefined) {
      cascadeCampaignStatus(campaignId, params.campaignStatus);
    }

    log('handleUpdateCampaign: ' + campaignId + ' now has ' + totalSteps +
        ' step(s), ' + leadsUpdated + ' lead(s) resynced', 'INFO');

    return {
      success: true,
      campaignId: campaignId,
      updated: finalRows.length,
      totalSequenceSteps: totalSteps,
      stepsAdded:   Math.max(0, totalSteps - mine.length),
      stepsRemoved: Math.max(0, mine.length - totalSteps),
      leadsResynced: leadsUpdated
    };

  } catch (e) {
    log('handleUpdateCampaign: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

/**
 * Rewrites totalSequenceSteps on every lead in a campaign after its step
 * count changes. Also releases leads that were marked Completed only because
 * they had reached the old final step.
 *
 * @param {string} campaignId
 * @param {number} totalSteps
 * @returns {number} rows updated
 */
function syncLeadTotalSteps_(campaignId, totalSteps) {
  try {
    const settings = getSettings(MASTER_SHEET_ID);
    const ss       = SpreadsheetApp.openById(settings.leadsSpreadsheetId);
    const sheet    = ss.getSheetByName(LEADS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) { return 0; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    if (!('totalSequenceSteps' in col)) { return 0; }

    let changed = 0;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][col['campaignId']] || '').trim() !== campaignId) { continue; }

      data[i][col['totalSequenceSteps']] = totalSteps;
      changed++;

      // A lead parked at the old last step should resume when steps are added.
      if ('sequenceStatus' in col) {
        const step = parseInt(data[i][col['sequenceStep']] || 0, 10);
        const done = String(data[i][col['sequenceStatus']] || '').trim();
        if (done === 'Completed' && step < totalSteps) {
          data[i][col['sequenceStatus']] = '';
        }
      }
    }

    if (changed > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }
    return changed;

  } catch (e) {
    log('syncLeadTotalSteps_: error — ' + e.message, 'ERROR');
    return 0;
  }
}

/**
 * Rewrites totalSequenceSteps on every lead in a campaign after its step
 * count changes. Also releases leads that were marked Completed only because
 * they had reached the old final step.
 *
 * @param {string} campaignId
 * @param {number} totalSteps
 * @returns {number} rows updated
 */
function syncLeadTotalSteps_(campaignId, totalSteps) {
  try {
    const settings = getSettings(MASTER_SHEET_ID);
    const ss       = SpreadsheetApp.openById(settings.leadsSpreadsheetId);
    const sheet    = ss.getSheetByName(LEADS_SHEET_NAME);
    if (!sheet || sheet.getLastRow() < 2) { return 0; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    if (!('totalSequenceSteps' in col)) { return 0; }

    let changed = 0;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][col['campaignId']] || '').trim() !== campaignId) { continue; }

      data[i][col['totalSequenceSteps']] = totalSteps;
      changed++;

      // A lead parked at the old last step should resume when steps are added.
      if ('sequenceStatus' in col) {
        const step = parseInt(data[i][col['sequenceStep']] || 0, 10);
        const done = String(data[i][col['sequenceStatus']] || '').trim();
        if (done === 'Completed' && step < totalSteps) {
          data[i][col['sequenceStatus']] = '';
        }
      }
    }

    if (changed > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }
    return changed;

  } catch (e) {
    log('syncLeadTotalSteps_: error — ' + e.message, 'ERROR');
    return 0;
  }
}

function handleGetDailySummary(params) {
  const days = parseInt(params.days || 14, 10); // default last 14 days

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName('dailySummary');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, summary: [] };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const lastRow = sheet.getLastRow();
    const rowsToRead = Math.min(lastRow - 1, days);
    const startRow   = Math.max(2, lastRow - rowsToRead + 1);
    const data    = sheet.getRange(startRow, 1, rowsToRead, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const summary = data.map(function(row) {
      return {
        date:           String(row[col['date']]           || '').trim(),
        emailsSent:     parseInt(row[col['emailsSent']]   || 0, 10),
        repliesReceived:parseInt(row[col['repliesReceived']]||0,10),
        bouncedEmails:  parseInt(row[col['bouncedEmails']] || 0, 10),
        activeSenders:  parseInt(row[col['activeSenders']] || 0, 10)
      };
    });

    return { success: true, summary: summary };

  } catch(e) {
    log('handleGetDailySummary: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleDeleteCampaign(params) {
  const campaignId = String(params.campaignId || '').trim();
  if (!campaignId) { return { success: false, error: 'Missing campaignId' }; }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName('Campaigns');
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, deleted: 0 }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Keep rows that don't belong to this campaign
    const surviving = data.filter(function(row) {
      return String(row[col['campaignId']] || '').trim() !== campaignId;
    });

    const deletedCount = data.length - surviving.length;

    if (deletedCount > 0) {
      sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
      if (surviving.length > 0) {
        sheet.getRange(2, 1, surviving.length, headers.length).setValues(surviving);
      }
      SpreadsheetApp.flush();

      // Also pause leads in masterLeads for this campaign
      cascadeCampaignStatus(campaignId, 'Deleted');
    }

    log('handleDeleteCampaign: deleted ' + deletedCount + ' rows for ' + campaignId, 'INFO');
    return { success: true, deleted: deletedCount };

  } catch(e) {
    log('handleDeleteCampaign: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleGetLeads(params) {
  const campaignId = String(params.campaignId || '').trim();
  const page       = parseInt(params.page     || 1,    10);
  const pageSize   = parseInt(params.pageSize || 50,   10);
  const statusFilter = String(params.status   || '').trim();
  const search = String(params.search || '').trim().toLowerCase();

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    const ss           = SpreadsheetApp.openById(leadsSheetId);
    const sheet        = ss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow      = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, leads: [], total: 0 }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Filter
    let filtered = data.filter(function(row) {
      const rowCampaign = String(row[col['campaignId']] || '').trim();
      const rowStatus   = String(row[col['status']]     || '').trim();
      if (campaignId && rowCampaign !== campaignId) { return false; }
      if (statusFilter && rowStatus !== statusFilter)  { return false; }

      if (search) {
        // Match the fields a person would actually search by.
        const haystack = [
          row[col['email']], row[col['firstName']], row[col['lastName']],
          row[col['companyName']], row[col['leadId']]
        ].map(function(v) { return String(v || '').toLowerCase(); }).join(' ');
        if (haystack.indexOf(search) === -1) { return false; }
      }
      return true;
    });

    const total  = filtered.length;
    const start  = (page - 1) * pageSize;
    const paged  = filtered.slice(start, start + pageSize);

        const leads = paged.map(function(row) {
      // Every column, read only when present, so this is safe whatever shape
      // the sheet is in.
      const get = function(name) {
        return (name in col) ? String(row[col[name]] || '').trim() : '';
      };

      return {
        leadId:             get('leadId'),
        email:              get('email'),
        firstName:          get('firstName'),
        lastName:           get('lastName'),
        companyName:        get('companyName'),
        companyWebsite:     get('companyWebsite'),
        jobTitle:           get('jobTitle'),
        department:         get('department'),
        industry:           get('industry'),
        addressLine:        get('addressLine'),
        city:               get('city'),
        state:              get('state'),
        country:            get('country'),
        countryCode:        get('countryCode'),
        phoneNumber:        get('phoneNumber'),
        linkedinPersonUrl:  get('linkedinPersonUrl'),
        linkedinCompanyUrl: get('linkedinCompanyUrl'),
        customVar1:         get('customVar1'),
        customVar2:         get('customVar2'),
        customVar3:         get('customVar3'),
        customVar4:         get('customVar4'),
        customVar5:         get('customVar5'),
        campaignId:         get('campaignId'),
        status:             get('status'),
        sequenceStep:       parseInt(row[col['sequenceStep']] || 0, 10),
        sequenceStatus:     get('sequenceStatus'),
        replyStatus:        get('replyStatus'),
        bounceStatus:       get('bounceStatus'),
        lastSentDate:       get('lastSentDate'),
        nextSentDate:       get('nextSentDate'),
        ownerSender:        get('ownerSender')
      };
    });

    return {
      success:  true,
      searched: !!search,
      leads:    leads,
      total:    total,
      page:     page,
      pageSize: pageSize,
      pages:    Math.ceil(total / pageSize)
    };

  } catch(e) {
    log('handleGetLeads: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleUpdateSender(params) {
  const emailID = String(params.emailID || '').trim();
  if (!emailID) { return { success: false, error: 'Missing emailID' }; }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    let senderRow    = null;
    let senderWebApp = null;

    for (let i = 0; i < data.length; i++) {
      const rowEmailID = String(data[i][col['emailID']] || '').trim();
      if (rowEmailID !== emailID) { continue; }

      // Fields anyone can update
      if (params.emailFirstName !== undefined) { data[i][col['emailFirstName']] = params.emailFirstName; }
      if (params.emailLastName  !== undefined) { data[i][col['emailLastName']]  = params.emailLastName; }
      if (params.isActive       !== undefined) { data[i][col['isActive']]       = params.isActive; }
      if (params.notes          !== undefined) { data[i][col['notes']]          = params.notes; }

      // Admin only fields
      if (params.dailyLimit     !== undefined) { data[i][col['dailyLimit']]     = parseInt(params.dailyLimit, 10); }
      if (params.emailAddress   !== undefined) { data[i][col['emailAddress']]   = params.emailAddress; }
      if (params.webAppUrl      !== undefined) { data[i][col['webAppUrl']]      = params.webAppUrl; }
      if (params.spreadsheetId  !== undefined) { data[i][col['spreadsheetId']]  = params.spreadsheetId; }

      senderWebApp = String(data[i][col['webAppUrl']] || '').trim();
      senderRow    = data[i];
      break;
    }

    if (!senderRow) {
      return { success: false, error: 'Sender not found: ' + emailID };
    }

    // Write back to senderAccounts
    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    SpreadsheetApp.flush();

    // Sync changed values to sender's own Config sheet
    if (senderWebApp) {
      const configToSync = {};
      if (params.dailyLimit     !== undefined) { configToSync['dailyLimit']       = String(params.dailyLimit); }
      if (params.emailFirstName !== undefined) { configToSync['senderFirstName']  = params.emailFirstName; }
      if (params.emailLastName  !== undefined) { configToSync['senderLastName']   = params.emailLastName; }
      if (params.isActive       !== undefined) { configToSync['isActive']         = params.isActive; }

      if (Object.keys(configToSync).length > 0) {
        try {
          UrlFetchApp.fetch(senderWebApp, {
            method:             'post',
            contentType:        'application/json',
            payload:            JSON.stringify({
              action: 'syncConfig',
              config: configToSync
            }),
            muteHttpExceptions: true
          });
          log('handleUpdateSender: synced config to sender ' + emailID, 'INFO');
        } catch(e) {
          log('handleUpdateSender: could not sync to sender — ' + e.message, 'WARN');
        }
      }
    }

    log('handleUpdateSender: updated sender ' + emailID, 'INFO');
    return { success: true, emailID: emailID };

  } catch(e) {
    log('handleUpdateSender: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleArchiveCampaignLeads(params) {
  const campaignId = String(params.campaignId || '').trim();
  if (!campaignId) { return { success: false, error: 'Missing campaignId' }; }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    const alertsEmail  = String(settings.alertsEmail || '').trim();

    const ss      = SpreadsheetApp.openById(leadsSheetId);
    const sheet   = ss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, archived: 0 }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Split leads
    const toArchive = [];
    const toKeep    = [];
    for (let i = 0; i < data.length; i++) {
      const rowCampaign = String(data[i][col['campaignId']] || '').trim();
      if (rowCampaign === campaignId) { toArchive.push(data[i]); }
      else                            { toKeep.push(data[i]); }
    }

    if (toArchive.length === 0) {
      return { success: true, archived: 0, message: 'No leads found for campaign' };
    }

    // ── Create new spreadsheet ───────────────────────────────────────────
    const archiveName   = 'Archive_' + campaignId + '_' + formatDate(new Date());
    const archiveSS     = SpreadsheetApp.create(archiveName);
    const archiveSheet  = archiveSS.getActiveSheet();
    archiveSheet.setName('leads');

    // Write headers + data
    archiveSheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    archiveSheet.getRange(2, 1, toArchive.length, headers.length).setValues(toArchive);
    SpreadsheetApp.flush();
    // ── metrics sheet ─────────────────────────────────────────────────────
    // Computed from the rows being archived, so it stays true even after the
    // leads leave masterLeads and the campaign row is deleted.
    try {
      writeArchiveMetrics_(archiveSS, campaignId, toArchive, col);
    } catch (e) {
      log('handleArchiveCampaignLeads: metrics sheet failed — ' + e.message, 'WARN');
    }

    // ── Set sharing to anyone with link can edit ─────────────────────────
    const archiveFile = DriveApp.getFileById(archiveSS.getId());
    archiveFile.setSharing(
      DriveApp.Access.ANYONE_WITH_LINK,
      DriveApp.Permission.EDIT
    );
    const archiveUrl = archiveFile.getUrl();

    // ── Send email notification ──────────────────────────────────────────
    if (alertsEmail) {
      GmailApp.sendEmail(
        alertsEmail,
        'Archive Ready — ' + campaignId + ' — ' + formatDate(new Date()),
        'Campaign ' + campaignId + ' leads have been archived.\n\n' +
        'Total archived: ' + toArchive.length + ' leads\n' +
        'Archive spreadsheet: ' + archiveUrl + '\n\n' +
        'This archive was created by Reach X.'
      );
      log('handleArchiveCampaignLeads: notification sent to ' + alertsEmail, 'INFO');
    } else {
      log('handleArchiveCampaignLeads: no alertsEmail in Settings — skipping notification', 'WARN');
    }

    // ── Rewrite masterLeads without archived leads ───────────────────────
    sheet.getRange(2, 1, lastRow - 1, headers.length).clearContent();
    if (toKeep.length > 0) {
      sheet.getRange(2, 1, toKeep.length, headers.length).setValues(toKeep);
    }
    SpreadsheetApp.flush();

    log('handleArchiveCampaignLeads: archived=' + toArchive.length +
        ' kept=' + toKeep.length +
        ' url=' + archiveUrl, 'INFO');

    // Reflect the archive immediately rather than waiting for the nightly sync.
    try {
      const settings = getSettings(MASTER_SHEET_ID);
      syncCampaignStats(MASTER_SHEET_ID, settings.leadsSpreadsheetId);
    } catch (e) {
      log('handleArchiveCampaignLeads: stats resync failed — ' + e.message, 'WARN');
    }

    return {
      success:    true,
      archived:   toArchive.length,
      kept:       toKeep.length,
      archiveUrl: archiveUrl,
      archiveName: archiveName
    };
    

  } catch(e) {
    log('handleArchiveCampaignLeads: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleGetSystemStatus(params) {
  try {
    const senders  = getActiveSenders();
    const requests = senders
      .filter(function(s) { return !!s.webAppUrl; })
      .map(function(s) {
        return {
          url:                s.webAppUrl,
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({ action: 'getSenderStatus' }),
          muteHttpExceptions: true
        };
      });

    let totalSentToday    = 0;
    let totalColdQueued   = 0;
    let totalFollowQueued = 0;
    let totalErrors       = 0;
    let activeTriggers    = 0;
    const senderDetails   = [];

    for (let b = 0; b < requests.length; b += 50) {
      const batch     = requests.slice(b, b + 50);
      const batchMeta = senders.slice(b, b + 50);
      let   responses = [];
      try { responses = UrlFetchApp.fetchAll(batch); } catch(e) { continue; }

      for (let r = 0; r < responses.length; r++) {
        if (responses[r].getResponseCode() !== 200) { continue; }
        try {
          const s = JSON.parse(responses[r].getContentText());
          totalSentToday    += s.sentToday      || 0;
          totalColdQueued   += s.coldQueued      || 0;
          totalFollowQueued += s.followupQueued  || 0;
          totalErrors       += s.coldError       || 0;
          activeTriggers    += s.activeTriggers  || 0;
          senderDetails.push({
            emailID:        batchMeta[r].emailID,
            sentToday:      s.sentToday     || 0,
            dailyLimit:     s.dailyLimit    || 25,
            coldQueued:     s.coldQueued    || 0,
            followupQueued: s.followupQueued|| 0,
            errors:         s.coldError     || 0,
            activeTriggers: s.activeTriggers|| 0
          });
        } catch(e) { continue; }
      }
    }

    return {
      success:          true,
      totalSentToday:   totalSentToday,
      totalColdQueued:  totalColdQueued,
      totalFollowQueued:totalFollowQueued,
      totalErrors:      totalErrors,
      activeTriggers:   activeTriggers,
      senderCount:      senders.length,
      senders:          senderDetails
    };

  } catch(e) {
    return { success: false, error: e.message };
  }
}
/**
 * Returns the Settings sheet as a flat object so a UI can render it.
 */
function handleGetSettings(params) {
  try {
    const settings = getSettings(MASTER_SHEET_ID);
    // leadsSpreadsheetId is the one value that must never be edited from a UI —
    // a typo there points the whole system at an empty sheet.
    const out = {};
    Object.keys(settings).forEach(function(k) {
      if (k === 'leadsSpreadsheetId') { return; }
      out[k] = settings[k];
    });
    out.leadsSpreadsheetConfigured = !!settings.leadsSpreadsheetId;
    return { success: true, settings: out };
  } catch (e) {
    log('handleGetSettings: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

var SETTINGS_WRITABLE = [
  'coldPriorityLimit', 'followupPriorityLimit', 'flexibleAllocation',
  'distributionMethod', 'alertsEmail',
  'sendingWindowStart', 'sendingWindowEnd', 'sendingDays', 'timeZone',
  'warmupEnabled', 'warmupStartLimit', 'warmupIncreaseBy',
  'warmupEveryDays', 'warmupMaxLimit', 'warmupNextIncreaseDate'
];

/**
 * Updates key/value pairs in the Settings sheet.
 * Only keys in SETTINGS_WRITABLE may change. leadsSpreadsheetId is deliberately
 * not writable, and warmupCurrentLimit is managed by runWarmupIncrease.
 */
function handleUpdateSettings(params) {
  const incoming = params.settings || {};
  const keys = Object.keys(incoming);
  if (keys.length === 0) {
    return { success: false, error: 'No settings provided' };
  }

  if (!acquireLock(30)) {
    return { success: false, error: 'Busy — could not acquire lock, try again' };
  }

  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName('Settings');
    if (!sheet) { return { success: false, error: 'Settings sheet not found' }; }

    const lastRow = Math.max(1, sheet.getLastRow());
    const data    = sheet.getRange(1, 1, lastRow, 2).getValues();
    const rowOf   = {};
    data.forEach(function(r, i) {
      const k = String(r[0]).trim();
      if (k) { rowOf[k] = i; }
    });

    const updated = [], rejected = [];

    keys.forEach(function(key) {
      if (SETTINGS_WRITABLE.indexOf(key) === -1) { rejected.push(key); return; }

      let value = incoming[key];
      // Dates must land as text or Sheets reformats them and formatDate()
      // comparisons stop matching — same trick runWarmupIncrease uses.
      if (key === 'warmupNextIncreaseDate' && value) { value = "'" + String(value).trim(); }
      else { value = String(value); }

      if (key in rowOf) { data[rowOf[key]][1] = value; }
      else              { data.push([key, value]); }
      updated.push(key);
    });

    if (updated.length > 0) {
      sheet.getRange(1, 1, data.length, 2).setValues(data);
      SpreadsheetApp.flush();
    }

    log('handleUpdateSettings: updated ' + updated.join(', ') +
        (rejected.length ? ' — rejected ' + rejected.join(', ') : ''), 'INFO');
    return { success: true, updated: updated, rejected: rejected };

  } catch (e) {
    log('handleUpdateSettings: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

function handleGetSenders(params) {
  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, senders: [] }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const senders = data.map(function(row) {
      return {
        emailID:        String(row[col['emailID']]        || '').trim(),
        emailAddress:   String(row[col['emailAddress']]   || '').trim(),
        emailFirstName: String(row[col['emailFirstName']] || '').trim(),
        emailLastName:  String(row[col['emailLastName']]  || '').trim(),
        isActive:       String(row[col['isActive']]       || '').trim(),
        dailyLimit:     parseInt(row[col['dailyLimit']]   || 25, 10),
        sentToday:      parseInt(row[col['sentToday']]    || 0,  10),
        totalSentCount: parseInt(row[col['totalSentCount']]|| 0, 10),
        // Dashboard reply/bounce rates divide by this, not by totalSentCount —
        // followups inflate totalSentCount so it is the wrong denominator.
        totalLeadsContacted: parseInt(row[col['totalLeadsContacted']] || 0, 10),
        repliesReceived:parseInt(row[col['repliesReceived']]||0, 10),
        bouncedEmails:  parseInt(row[col['bouncedEmails']] || 0, 10),
        lastRunTime:    String(row[col['lastRunTime']]    || '').trim(),
        notes:          String(row[col['notes']]          || '').trim(),
        tagName:        String(row[col['tagName']]        || '').trim(),
        tagColor:       String(row[col['tagColor']]       || '').trim()
      };
    }).filter(function(s) { return !!s.emailID; });

    return { success: true, senders: senders };

  } catch(e) {
    log('handleGetSenders: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleHealthCheck(params) {
  try {
    const senders  = getActiveSenders();
    const settings = getSettings(MASTER_SHEET_ID);

    return {
      success:           true,
      timestamp:         new Date().toISOString(),
      activeSenders:     senders.length,
      masterSheetId:     MASTER_SHEET_ID,
      leadsSheetId:      settings.leadsSpreadsheetId || 'NOT SET',
      alertsEmail:       settings.alertsEmail        || 'NOT SET'
    };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function handleGetCampaignStats(params) {
  const campaignId = String(params.campaignId || '').trim();
  if (!campaignId) { return { success: false, error: 'Missing campaignId' }; }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const ss           = SpreadsheetApp.openById(settings.leadsSpreadsheetId);
    const sheet        = ss.getSheetByName(LEADS_SHEET_NAME);
    const lastRow      = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, stats: {} }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const stats = {
      total:     0,
      pending:   0,
      queued:    0,
      sent:      0,
      replied:   0,
      bounced:   0,
      completed: 0,
      error:     0
    };

    for (let i = 0; i < data.length; i++) {
      const rowCamp = String(data[i][col['campaignId']] || '').trim();
      if (rowCamp !== campaignId) { continue; }

      stats.total++;
      const status         = String(data[i][col['status']]         || '').trim().toLowerCase();
      const replyStatus    = String(data[i][col['replyStatus']]    || '').trim().toLowerCase();
      const bounceStatus   = String(data[i][col['bounceStatus']]   || '').trim().toLowerCase();
      const sequenceStatus = String(data[i][col['sequenceStatus']] || '').trim().toLowerCase();

      if (bounceStatus   === 'bounced')   { stats.bounced++;   continue; }
      if (replyStatus    === 'replied')   { stats.replied++;   continue; }
      if (sequenceStatus === 'completed') { stats.completed++; continue; }
      if (status === 'pending')           { stats.pending++;   }
      else if (status === 'queued')       { stats.queued++;    }
      else if (status === 'sent')         { stats.sent++;      }
      else if (status === 'error')        { stats.error++;     }
    }
    // Recount from the rows just walked, so the Overview tab does not depend
    // on when the sync last ran.
    stats.stepsSent      = 0;
    stats.contactedCount = 0;
    for (let i = 0; i < data.length; i++) {
      if (String(data[i][col['campaignId']] || '').trim() !== campaignId) { continue; }
      const step = parseInt(data[i][col['sequenceStep']] || 0, 10);
      if (step > 0) {
        stats.stepsSent += step;
        stats.contactedCount++;
      }
    }

    // Calculate rates
    stats.replyRate  = stats.total > 0 ? ((stats.replied  / stats.total) * 100).toFixed(1) + '%' : '0%';
    stats.bounceRate = stats.total > 0 ? ((stats.bounced  / stats.total) * 100).toFixed(1) + '%' : '0%';

    return { success: true, campaignId: campaignId, stats: stats };

  } catch(e) {
    log('handleGetCampaignStats: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleGetMasterTriggers(params) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const list     = triggers.map(function(t) {
      return {
        functionName: t.getHandlerFunction(),
        triggerId:    t.getUniqueId(),
        type:         t.getEventType().toString()
      };
    });
    return { success: true, count: list.length, triggers: list };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

function handleUpdateSenderConfig(params) {
  // params.senderIds = ['S001','S002'] OR 'all' for everyone
  // params.fields = { dailyLimit: 50, senderFirstName: 'John' } etc.

  const senderIds = params.senderIds || 'all';
  const fields    = params.fields    || {};

  if (Object.keys(fields).length === 0) {
    return { success: false, error: 'No fields provided' };
  }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: false, error: 'No senders found' }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Build target sender list
    const targetAll = (senderIds === 'all');
    const targetIds = targetAll ? [] : 
      (Array.isArray(senderIds) ? senderIds : [senderIds]);

    // Map config field names to senderAccounts column names
    const colMapping = {
      'senderFirstName':    'emailFirstName',
      'senderLastName':     'emailLastName',
      'senderId':           'emailID',
      'senderEmail':        'emailAddress',
      'dailyLimit':         'dailyLimit',
      'sentToday':          'sentToday',
      'totalSentCount':     'totalSentCount',
      'repliesReceived':    'repliesReceived',
      'bouncedEmails':      'bouncedEmails',
      'lastRunTime':        'lastRunTime'
    };

    const fetchRequests = [];
    const fetchMeta     = [];
    let   updatedCount  = 0;

    for (let i = 0; i < data.length; i++) {
      const rowEmailID  = String(data[i][col['emailID']]  || '').trim();
      const rowWebApp   = String(data[i][col['webAppUrl']]|| '').trim();
      const rowIsActive = String(data[i][col['isActive']] || '').trim();

      // Skip if not in target list
      if (!targetAll && targetIds.indexOf(rowEmailID) === -1) { continue; }

      // Update senderAccounts row in memory
      Object.keys(fields).forEach(function(fieldKey) {
        const accountCol = colMapping[fieldKey] || fieldKey;
        if (accountCol in col) {
          data[i][col[accountCol]] = fields[fieldKey];
        }
      });

      // Build config payload to push to sender's Config sheet
      // Only include fields that are relevant to sender Config
      const configPayload = {};
      Object.keys(fields).forEach(function(fieldKey) {
        configPayload[fieldKey] = String(fields[fieldKey]);
      });

      if (rowWebApp) {
        fetchRequests.push({
          url:                rowWebApp,
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({
            action: 'syncConfig',
            config: configPayload
          }),
          muteHttpExceptions: true
        });
        fetchMeta.push({ emailID: rowEmailID });
      }

      updatedCount++;
    }

    // Write updated senderAccounts in one call
    if (updatedCount > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }

    // Push to all sender Config sheets in parallel batches
    let synced = 0;
    let failed = 0;

    for (let b = 0; b < fetchRequests.length; b += 50) {
      const batch     = fetchRequests.slice(b, b + 50);
      const batchMeta = fetchMeta.slice(b, b + 50);
      let   responses = [];

      try {
        responses = UrlFetchApp.fetchAll(batch);
      } catch(e) {
        log('handleUpdateSenderConfig: fetchAll error — ' + e.message, 'ERROR');
        failed += batch.length;
        continue;
      }

      for (let r = 0; r < responses.length; r++) {
        if (responses[r].getResponseCode() === 200) {
          synced++;
          log('handleUpdateSenderConfig: synced ' + batchMeta[r].emailID, 'INFO');
        } else {
          failed++;
          log('handleUpdateSenderConfig: failed ' + batchMeta[r].emailID + 
              ' (' + responses[r].getResponseCode() + ')', 'ERROR');
        }
      }
    }

    log('handleUpdateSenderConfig: updated=' + updatedCount + 
        ' synced=' + synced + ' failed=' + failed, 'INFO');

    return {
      success:      true,
      updated:      updatedCount,
      synced:       synced,
      failed:       failed,
      targetedAll:  targetAll,
      fieldsChanged: Object.keys(fields)
    };

  } catch(e) {
    log('handleUpdateSenderConfig: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleRunWarmupIncrease(params) {
  try {
    runWarmupIncrease();
    return { success: true, message: 'Warmup increase ran successfully' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

// ============================================================================
// SENDER TAGS
// ============================================================================

/**
 * Get all senders with a specific tag
 */
function handleGetSendersByTag(params) {
  const tag = String(params.tag || '').trim();
  if (!tag) { return { success: false, error: 'Missing tag' }; }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, senders: [] }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const senders = [];
    for (let i = 0; i < data.length; i++) {
        const rowTags = String(data[i][col['tagName']] || '').trim().toLowerCase()
                .split(',').map(function(t) { return t.trim(); });
        if (rowTags.indexOf(tag.toLowerCase()) !== -1) {
        senders.push({
          emailID:      String(data[i][col['emailID']]      || '').trim(),
          emailAddress: String(data[i][col['emailAddress']] || '').trim(),
          tagName:      String(data[i][col['tagName']]      || '').trim(),
          tagColor:     String(data[i][col['tagColor']]     || '').trim(),
          isActive:     String(data[i][col['isActive']]     || '').trim(),
          dailyLimit:   parseInt(data[i][col['dailyLimit']] || 25, 10)
        });
      }
    }

    log('handleGetSendersByTag: found ' + senders.length + ' senders with tag=' + tag, 'INFO');
    return { success: true, tag: tag, senders: senders };

  } catch(e) {
    log('handleGetSendersByTag: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/** Splits a comma list into trimmed members. */
function splitList_(value) {
  return String(value || '').split(',')
    .map(function(x) { return x.trim(); })
    .filter(function(x) { return x !== ''; });
}

/**
 * Reads a row's tags as [{ name, color }], tolerating a colour list that is
 * shorter than the tag list (rows written before this patch have exactly one).
 */
function readRowTags_(tagsRaw, colorsRaw) {
  const names  = splitList_(tagsRaw);
  const colors = splitList_(colorsRaw);
  return names.map(function(name, i) {
    return { name: name, color: colors[i] || '' };
  });
}

function handleAssignTag(params) {
  // No default of 'all' — tagging every sender in the system must be asked
  // for explicitly, never fallen into by omitting a field.
  const senderIds = params.senderIds;
  const tagName   = String(params.tagName  || '').trim();
  const tagColor  = String(params.tagColor || '#6b7280').trim();

  if (!tagName)   { return { success: false, error: 'Missing tagName' }; }
  if (!senderIds) { return { success: false, error: 'Missing senderIds — pass an array, or the string "all"' }; }

  if (!acquireLock(30)) {
    return { success: false, error: 'Busy — could not acquire lock' };
  }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: false, error: 'No senders found' }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    if (!('tagName' in col) || !('tagColor' in col)) {
      return { success: false, error: 'tagName or tagColor column not found in senderAccounts' };
    }

    const targetAll = (senderIds === 'all');
    const targetIds = targetAll ? [] : (Array.isArray(senderIds) ? senderIds : [senderIds]);

    let updated = 0;
    for (let i = 0; i < data.length; i++) {
      const rowEmailID = String(data[i][col['emailID']] || '').trim();
      if (!targetAll && targetIds.indexOf(rowEmailID) === -1) { continue; }

      const tags = readRowTags_(data[i][col['tagName']], data[i][col['tagColor']]);
      const hit  = tags.filter(function(t) { return t.name === tagName; })[0];

      if (hit) { hit.color = tagColor; }        // re-colouring an existing tag
      else     { tags.push({ name: tagName, color: tagColor }); }

      data[i][col['tagName']]  = tags.map(function(t) { return t.name; }).join(',');
      data[i][col['tagColor']] = tags.map(function(t) { return t.color || '#6b7280'; }).join(',');
      updated++;
    }

    if (updated > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }

    log('handleAssignTag: ' + tagName + ' → ' + updated + ' sender(s)', 'INFO');
    return { success: true, tag: tagName, color: tagColor, updated: updated };

  } catch (e) {
    log('handleAssignTag: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

function handleRemoveTag(params) {
  const senderIds = params.senderIds;
  const tagName   = String(params.tagName || '').trim();

  if (!tagName)   { return { success: false, error: 'Missing tagName' }; }
  if (!senderIds) { return { success: false, error: 'Missing senderIds — pass an array, or the string "all"' }; }

  if (!acquireLock(30)) {
    return { success: false, error: 'Busy — could not acquire lock' };
  }

  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, updated: 0 }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const targetAll = (senderIds === 'all');
    const targetIds = targetAll ? [] : (Array.isArray(senderIds) ? senderIds : [senderIds]);

    let updated = 0;
    for (let i = 0; i < data.length; i++) {
      const rowEmailID = String(data[i][col['emailID']] || '').trim();
      if (!targetAll && targetIds.indexOf(rowEmailID) === -1) { continue; }

      const tags = readRowTags_(data[i][col['tagName']], data[i][col['tagColor']]);
      // Compare against each member. The original compared the whole joined
      // string, so "warm,uk" never matched "uk" and removal silently no-opped.
      const kept = tags.filter(function(t) { return t.name !== tagName; });
      if (kept.length === tags.length) { continue; }

      data[i][col['tagName']]  = kept.map(function(t) { return t.name; }).join(',');
      data[i][col['tagColor']] = kept.map(function(t) { return t.color || '#6b7280'; }).join(',');
      updated++;
    }

    if (updated > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
    }

    log('handleRemoveTag: ' + tagName + ' removed from ' + updated + ' sender(s)', 'INFO');
    return { success: true, tag: tagName, updated: updated };

  } catch (e) {
    log('handleRemoveTag: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  } finally {
    releaseLock();
  }
}

function handleGetAllTags(params) {
  try {
    const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet   = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return { success: true, tags: [] }; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const tagMap = {};
    for (let i = 0; i < data.length; i++) {
      const tags = readRowTags_(data[i][col['tagName']], data[i][col['tagColor']]);
      tags.forEach(function(t) {
        if (!tagMap[t.name]) {
          tagMap[t.name] = { tagName: t.name, tagColor: t.color || '', senderCount: 0 };
        }
        // First non-empty colour wins, so one tag reads the same everywhere
        // even if two rows disagree.
        if (!tagMap[t.name].tagColor && t.color) { tagMap[t.name].tagColor = t.color; }
        tagMap[t.name].senderCount++;
      });
    }

    const tags = Object.keys(tagMap).map(function(k) {
      if (!tagMap[k].tagColor) { tagMap[k].tagColor = '#6b7280'; }
      return tagMap[k];
    }).sort(function(a, b) { return a.tagName.localeCompare(b.tagName); });

    return { success: true, tags: tags };

  } catch (e) {
    log('handleGetAllTags: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Update sender config for all senders with a specific tag
 */
function handleUpdateSendersByTag(params) {
  const tag    = String(params.tag    || '').trim();
  const fields = params.fields        || {};

  if (!tag)                             { return { success: false, error: 'Missing tag' }; }
  if (Object.keys(fields).length === 0) { return { success: false, error: 'No fields provided' }; }

  try {
    // Get sender IDs with this tag
    const tagResult = handleGetSendersByTag({ tag: tag });
    if (!tagResult.success || tagResult.senders.length === 0) {
      return { success: false, error: 'No senders found with tag: ' + tag };
    }

    const senderIds = tagResult.senders.map(function(s) { return s.emailID; });

    // Use existing updateSenderConfig with specific IDs
    const result = handleUpdateSenderConfig({
      senderIds: senderIds,
      fields:    fields
    });

    log('handleUpdateSendersByTag: tag=' + tag + ' senders=' + senderIds.length +
        ' synced=' + result.synced, 'INFO');

    return {
      success:   result.success,
      tag:       tag,
      senders:   senderIds.length,
      synced:    result.synced,
      failed:    result.failed
    };

  } catch(e) {
    log('handleUpdateSendersByTag: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}


// ============================================================================
// GLOBAL BLOCKLIST
// ============================================================================

/**
 * Reads the Blocklist sheet and returns { blockedEmails, blockedDomains }
 * Used internally by uploadLeads to filter before inserting
 */
function getBlocklist_() {
  const blockedEmails  = {};
  const blockedDomains = {};

  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName('Blocklist');
    if (!sheet || sheet.getLastRow() < 2) { 
      return { blockedEmails: blockedEmails, blockedDomains: blockedDomains }; 
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    for (let i = 0; i < data.length; i++) {
      const email  = String(data[i][col['blacklistedEmail']]  || '').trim().toLowerCase();
      const domain = String(data[i][col['blacklistedDomain']] || '').trim().toLowerCase();
      if (email)  { blockedEmails[email]   = true; }
      if (domain) { blockedDomains[domain] = true; }
    }

    log('getBlocklist_: loaded ' + Object.keys(blockedEmails).length + 
        ' emails, ' + Object.keys(blockedDomains).length + ' domains', 'INFO');

  } catch(e) {
    log('getBlocklist_: error — ' + e.message, 'ERROR');
  }

  return { blockedEmails: blockedEmails, blockedDomains: blockedDomains };
}

/**
 * Add emails/domains to blocklist
 */
function handleAddToBlocklist(params) {
  const emails  = params.emails  || [];
  const domains = params.domains || [];
  const reason  = String(params.reason || 'manual').trim();

  if (emails.length === 0 && domains.length === 0) {
    return { success: false, error: 'No emails or domains provided' };
  }

  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    let   sheet = ss.getSheetByName('Blocklist');

    // Create sheet if it doesn't exist
    if (!sheet) {
      sheet = ss.insertSheet('Blocklist');
      const headers = ['blacklistedEmail', 'blacklistedDomain', 'reason', 'dateAdded'];
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
      const headerRange = sheet.getRange(1, 1, 1, headers.length);
      headerRange.setBackground('#dc2626');
      headerRange.setFontColor('#ffffff');
      headerRange.setFontWeight('bold');
      sheet.setFrozenRows(1);
    }

    const today   = formatDate(new Date());
    const rows    = [];

    emails.forEach(function(email) {
      const clean = String(email || '').trim().toLowerCase();
      if (clean) { rows.push([clean, '', reason, today]); }
    });

    domains.forEach(function(domain) {
      const clean = String(domain || '').trim().toLowerCase();
      if (clean) { rows.push(['', clean, reason, today]); }
    });

    if (rows.length > 0) {
      const startRow = sheet.getLastRow() + 1;
      sheet.getRange(startRow, 1, rows.length, 4).setValues(rows);
      SpreadsheetApp.flush();
    }

    log('handleAddToBlocklist: added ' + emails.length + ' emails, ' + 
        domains.length + ' domains', 'INFO');
    return { 
      success: true, 
      addedEmails:  emails.length,
      addedDomains: domains.length
    };

  } catch(e) {
    log('handleAddToBlocklist: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Remove emails/domains from blocklist
 */
function handleRemoveFromBlocklist(params) {
  const emails  = params.emails  || [];
  const domains = params.domains || [];

  if (emails.length === 0 && domains.length === 0) {
    return { success: false, error: 'No emails or domains provided' };
  }

  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName('Blocklist');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, removed: 0 };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const emailSet  = {};
    const domainSet = {};
    emails.forEach(function(e)  { emailSet[String(e  || '').toLowerCase()] = true; });
    domains.forEach(function(d) { domainSet[String(d || '').toLowerCase()] = true; });

    const surviving = data.filter(function(row) {
      const rowEmail  = String(row[col['blacklistedEmail']]  || '').trim().toLowerCase();
      const rowDomain = String(row[col['blacklistedDomain']] || '').trim().toLowerCase();
      return !emailSet[rowEmail] && !domainSet[rowDomain];
    });

    const removedCount = data.length - surviving.length;

    if (removedCount > 0) {
      sheet.getRange(2, 1, data.length, headers.length).clearContent();
      if (surviving.length > 0) {
        sheet.getRange(2, 1, surviving.length, headers.length).setValues(surviving);
      }
      SpreadsheetApp.flush();
    }

    log('handleRemoveFromBlocklist: removed ' + removedCount + ' entries', 'INFO');
    return { success: true, removed: removedCount };

  } catch(e) {
    log('handleRemoveFromBlocklist: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Get full blocklist
 */
function handleGetBlocklist(params) {
  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName('Blocklist');
    if (!sheet || sheet.getLastRow() < 2) {
      return { success: true, emails: [], domains: [] };
    }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const emails  = [];
    const domains = [];

    for (let i = 0; i < data.length; i++) {
      const email  = String(data[i][col['blacklistedEmail']]  || '').trim();
      const domain = String(data[i][col['blacklistedDomain']] || '').trim();
      const reason = String(data[i][col['reason']]            || '').trim();
      const date   = String(data[i][col['dateAdded']]         || '').trim();

      if (email)  { emails.push({ email:  email,  reason: reason, dateAdded: date }); }
      if (domain) { domains.push({ domain: domain, reason: reason, dateAdded: date }); }
    }

    return { success: true, emails: emails, domains: domains };

  } catch(e) {
    log('handleGetBlocklist: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleGetFullSystemReport(params) {
  const report = {
    generatedAt:    new Date().toISOString(),
    senders:        {},
    campaigns:      {},
    dailySummary:   {},
    masterTriggers: {},
    totals:         {}
  };

  try {
    // ── 1. Master Control triggers ────────────────────────────────────────
    try {
      const triggers     = ScriptApp.getProjectTriggers();
      report.masterTriggers = {
        count: triggers.length,
        list:  triggers.map(function(t) {
          return { name: t.getHandlerFunction(), id: t.getUniqueId() };
        })
      };
    } catch(e) {
      report.masterTriggers = { error: e.message };
    }

    // ── 2. Campaign stats ────────────────────────────────────────────────
    try {
      const campResult   = handleGetCampaigns({});
      report.campaigns   = {
        count:     campResult.campaigns ? campResult.campaigns.length : 0,
        campaigns: campResult.campaigns || []
      };
    } catch(e) {
      report.campaigns = { error: e.message };
    }

    // ── 3. Daily summary (last 7 days) ───────────────────────────────────
    try {
      const summaryResult  = handleGetDailySummary({ days: 7 });
      report.dailySummary  = {
        last7Days: summaryResult.summary || []
      };
    } catch(e) {
      report.dailySummary = { error: e.message };
    }

    // ── 4. All sender statuses in parallel ────────────────────────────────
    try {
      const senders  = getActiveSenders();
      const requests = senders
        .filter(function(s) { return !!s.webAppUrl; })
        .map(function(s) {
          return {
            url:                s.webAppUrl,
            method:             'post',
            contentType:        'application/json',
            payload:            JSON.stringify({ action: 'getSenderStatus' }),
            muteHttpExceptions: true
          };
        });

      const senderDetails  = [];
      let   totalSentToday = 0;
      let   totalQueued    = 0;
      let   totalErrors    = 0;
      let   totalTriggers  = 0;
      let   activeSenders  = 0;
      let   failedSenders  = 0;

      for (let b = 0; b < requests.length; b += 50) {
        const batch     = requests.slice(b, b + 50);
        const batchMeta = senders.slice(b, b + 50);
        let   responses = [];

        try { responses = UrlFetchApp.fetchAll(batch); } catch(e) { continue; }

        for (let r = 0; r < responses.length; r++) {
          const sender = batchMeta[r];
          if (responses[r].getResponseCode() !== 200) {
            failedSenders++;
            senderDetails.push({
              emailID: sender.emailID,
              status:  'unreachable',
              error:   'HTTP ' + responses[r].getResponseCode()
            });
            continue;
          }

          try {
            const s = JSON.parse(responses[r].getContentText());
            totalSentToday += s.sentToday      || 0;
            totalQueued    += s.coldQueued      || 0;
            totalQueued    += s.followupQueued  || 0;
            totalErrors    += s.coldError       || 0;
            totalTriggers  += s.activeTriggers  || 0;
            activeSenders++;

            senderDetails.push({
              emailID:        sender.emailID,
              emailAddress:   sender.emailAddress,
              status:         'ok',
              sentToday:      s.sentToday       || 0,
              dailyLimit:     s.dailyLimit      || 25,
              coldQueued:     s.coldQueued      || 0,
              followupQueued: s.followupQueued  || 0,
              errors:         s.coldError       || 0,
              activeTriggers: s.activeTriggers  || 0
            });
          } catch(e) {
            failedSenders++;
            senderDetails.push({
              emailID: sender.emailID,
              status:  'parse_error',
              error:   e.message
            });
          }
        }
      }

      report.senders = {
        total:         senders.length,
        active:        activeSenders,
        unreachable:   failedSenders,
        details:       senderDetails
      };

      report.totals = {
        sentToday:      totalSentToday,
        totalQueued:    totalQueued,
        totalErrors:    totalErrors,
        activeTriggers: totalTriggers
      };

    } catch(e) {
      report.senders = { error: e.message };
    }

    return { success: true, report: report };

  } catch(e) {
    log('handleGetFullSystemReport: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleValidateSystem(params) {
  const results = {
    passed:  [],
    failed:  [],
    warnings: []
  };

  // ── Helper functions ──────────────────────────────────────────────────────
  function pass(check)    { results.passed.push(check); }
  function fail(check)    { results.failed.push(check); }
  function warn(check)    { results.warnings.push(check); }

  try {
    const ss = SpreadsheetApp.openById(MASTER_SHEET_ID);

    // ── Check 1: Required Settings keys ────────────────────────────────────
    const requiredSettings = [
      'leadsSpreadsheetId', 'coldPriorityLimit', 'followupPriorityLimit',
      'flexibleAllocation', 'distributionMethod'
    ];
    const settings = getSettings(MASTER_SHEET_ID);
    requiredSettings.forEach(function(key) {
      if (settings[key]) { pass('Settings: ' + key + ' = ' + settings[key]); }
      else               { fail('Settings: missing key "' + key + '"'); }
    });

    // ── Check 2: masterLeads sheet exists and has required columns ──────────
    const requiredLeadCols = [
      'leadId', 'email', 'campaignId', 'campaignStatus', 'status',
      'sequenceStep', 'totalSequenceSteps', 'threadId', 'ownerSender',
      'lastSentDate', 'nextSentDate', 'replyStatus', 'bounceStatus'
    ];
    try {
      const leadsSS    = SpreadsheetApp.openById(settings.leadsSpreadsheetId);
      const leadsSheet = leadsSS.getSheetByName('masterLeads');
      if (!leadsSheet) {
        fail('masterLeads: sheet not found in leads spreadsheet');
      } else {
        pass('masterLeads: sheet exists');
        const leadHeaders = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
        const leadColSet  = {};
        leadHeaders.forEach(function(h) { leadColSet[String(h).trim()] = true; });
        requiredLeadCols.forEach(function(col) {
          if (leadColSet[col]) { pass('masterLeads column: ' + col); }
          else                  { fail('masterLeads column: missing "' + col + '"'); }
        });
        const leadCount = leadsSheet.getLastRow() - 1;
        pass('masterLeads: ' + leadCount + ' leads loaded');
      }
    } catch(e) {
      fail('masterLeads: could not open leads spreadsheet — ' + e.message);
    }

    // ── Check 3: Campaigns sheet exists and has active campaigns ───────────
    const campSheet = ss.getSheetByName('Campaigns');
    if (!campSheet) {
      fail('Campaigns: sheet not found');
    } else {
      pass('Campaigns: sheet exists');
      const campData    = campSheet.getRange(2, 1, Math.max(1, campSheet.getLastRow() - 1),
                           campSheet.getLastColumn()).getValues();
      const campHeaders = campSheet.getRange(1, 1, 1, campSheet.getLastColumn()).getValues()[0];
      const campCol     = {};
      campHeaders.forEach(function(h, i) { campCol[String(h).trim()] = i; });

      let activeCamps = 0;
      campData.forEach(function(row) {
        if (String(row[campCol['campaignStatus']] || '').trim() === 'Active') { activeCamps++; }
      });
      if (activeCamps > 0) { pass('Campaigns: ' + activeCamps + ' active campaign rows'); }
      else                  { warn('Campaigns: no Active campaigns found'); }
    }

    // ── Check 4: senderAccounts exists and has active senders ─────────────
    const saSheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    if (!saSheet) {
      fail('senderAccounts: sheet not found');
    } else {
      pass('senderAccounts: sheet exists');
      const senders = getActiveSenders();
      if (senders.length > 0) { pass('senderAccounts: ' + senders.length + ' active senders'); }
      else                     { fail('senderAccounts: no active senders found'); }

      // Check required columns
      const requiredSACols = ['emailID', 'emailAddress', 'webAppUrl', 
                               'spreadsheetId', 'isActive', 'dailyLimit'];
      const saHeaders = saSheet.getRange(1, 1, 1, saSheet.getLastColumn()).getValues()[0];
      const saColSet  = {};
      saHeaders.forEach(function(h) { saColSet[String(h).trim()] = true; });
      requiredSACols.forEach(function(col) {
        if (saColSet[col]) { pass('senderAccounts column: ' + col); }
        else                { fail('senderAccounts column: missing "' + col + '"'); }
      });
    }

    // ── Check 5: stagingUpdates sheet exists ───────────────────────────────
    const stagingSheet = ss.getSheetByName('stagingUpdates');
    if (stagingSheet) { pass('stagingUpdates: sheet exists'); }
    else               { fail('stagingUpdates: sheet not found — processStagingUpdates will fail'); }

    // ── Check 6: Ping all active sender webApps ───────────────────────────
    const senders  = getActiveSenders();
    const requests = senders
      .filter(function(s) { return !!s.webAppUrl; })
      .map(function(s) {
        return {
          url:                s.webAppUrl,
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({ action: 'getSenderStatus' }),
          muteHttpExceptions: true
        };
      });

    let reachable   = 0;
    let unreachable = 0;

    for (let b = 0; b < requests.length; b += 50) {
      const batch     = requests.slice(b, b + 50);
      const batchMeta = senders.slice(b, b + 50);
      let   responses = [];
      try { responses = UrlFetchApp.fetchAll(batch); } catch(e) { continue; }

      for (let r = 0; r < responses.length; r++) {
        if (responses[r].getResponseCode() === 200) {
          reachable++;

          // Check sender has required Config keys
          try {
            const status = JSON.parse(responses[r].getContentText());
            if (status.dailyLimit) {
              pass('Sender ' + batchMeta[r].emailID + ': reachable, limit=' + status.dailyLimit);
            } else {
              warn('Sender ' + batchMeta[r].emailID + ': reachable but missing dailyLimit in Config');
            }
          } catch(e) {
            warn('Sender ' + batchMeta[r].emailID + ': reachable but response parse error');
          }
        } else {
          unreachable++;
          fail('Sender ' + batchMeta[r].emailID + ': UNREACHABLE — HTTP ' + 
               responses[r].getResponseCode());
        }
      }
    }

    // ── Check 7: Master Control triggers ──────────────────────────────────
    const requiredTriggers = [
      'resetSentToday', 'scanAllSendersForReplies', 'distributeCold',
      'distributeFollowups', 'processStagingUpdates', 'scanAllSendersForBounces',
      'runSyncStats', 'runDailyReset'
    ];
    const projectTriggers  = ScriptApp.getProjectTriggers();
    const triggerNames     = projectTriggers.map(function(t) { return t.getHandlerFunction(); });

    requiredTriggers.forEach(function(name) {
      if (triggerNames.indexOf(name) !== -1) { pass('Trigger: ' + name + ' exists'); }
      else                                    { fail('Trigger: ' + name + ' MISSING'); }
    });

    // ── Summary ───────────────────────────────────────────────────────────
    const allGood = results.failed.length === 0;

    log('handleValidateSystem: passed=' + results.passed.length +
        ' failed=' + results.failed.length +
        ' warnings=' + results.warnings.length, 'INFO');

    return {
      success:      true,
      healthy:      allGood,
      passedCount:  results.passed.length,
      failedCount:  results.failed.length,
      warningCount: results.warnings.length,
      passed:       results.passed,
      failed:       results.failed,
      warnings:     results.warnings,
      summary:      allGood
        ? '✅ System is healthy — ' + results.passed.length + ' checks passed'
        : '❌ ' + results.failed.length + ' checks failed — fix before launching'
    };

  } catch(e) {
    log('handleValidateSystem: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

function handleSendDailyDigest(params) {
  try {
    sendDailyDigest();
    return { success: true, message: 'Daily digest sent' };
  } catch(e) {
    return { success: false, error: e.message };
  }
}

/**
 * Writes a permanent performance snapshot into an archive spreadsheet.
 *
 * Counted the same way getCampaignStats does — bounced, then replied, then
 * completed, then status — so the archive agrees with what the app showed
 * while the campaign was live.
 *
 * @param {Spreadsheet} archiveSS  the newly created archive
 * @param {string}      campaignId
 * @param {Array}       rows       the lead rows being archived
 * @param {Object}      col        header name → index, from masterLeads
 */
function writeArchiveMetrics_(archiveSS, campaignId, rows, col) {
  const s = {
    total: 0, pending: 0, queued: 0, sent: 0,
    replied: 0, bounced: 0, completed: 0, error: 0
  };
  const senders = {};
  let firstSent = '', lastSent = '';

  for (let i = 0; i < rows.length; i++) {
    const status    = String(rows[i][col['status']]         || '').trim().toLowerCase();
    const reply     = String(rows[i][col['replyStatus']]    || '').trim().toLowerCase();
    const bounce    = String(rows[i][col['bounceStatus']]   || '').trim().toLowerCase();
    const seqStatus = String(rows[i][col['sequenceStatus']] || '').trim().toLowerCase();
    const owner     = String(rows[i][col['ownerSender']]    || '').trim();
    const sentDate  = String(rows[i][col['lastSentDate']]   || '').trim();

    s.total++;
    if (owner) { senders[owner] = (senders[owner] || 0) + 1; }

    if (sentDate) {
      if (!firstSent || sentDate < firstSent) { firstSent = sentDate; }
      if (!lastSent  || sentDate > lastSent)  { lastSent  = sentDate; }
    }

    if      (bounce    === 'bounced')   { s.bounced++; }
    else if (reply     === 'replied')   { s.replied++; }
    else if (seqStatus === 'completed') { s.completed++; }
    else if (status === 'pending')      { s.pending++; }
    else if (status === 'queued')       { s.queued++; }
    else if (status === 'sent')         { s.sent++; }
    else if (status === 'error')        { s.error++; }
  }

  // Anyone who received at least one email. Rates against total leads would
  // count people who were never contacted and read far too low.
  const contacted = s.sent + s.replied + s.completed + s.bounced;
  const rate = function(part) {
    return contacted > 0 ? ((part / contacted) * 100).toFixed(2) + '%' : '0%';
  };

  const sheet = archiveSS.insertSheet('metrics');
  const out = [
    ['Metric', 'Value'],
    ['Campaign ID',        campaignId],
    ['Archived on',        formatDate(new Date())],
    ['First send',         firstSent || 'never sent'],
    ['Last send',          lastSent  || 'never sent'],
    ['', ''],
    ['Total leads',        s.total],
    ['Contacted',          contacted],
    ['Never contacted',    s.pending + s.queued],
    ['', ''],
    ['Replied',            s.replied],
    ['Reply rate',         rate(s.replied)],
    ['Bounced',            s.bounced],
    ['Bounce rate',        rate(s.bounced)],
    ['', ''],
    ['Completed sequence', s.completed],
    ['Awaiting reply',     s.sent],
    ['Pending',            s.pending],
    ['Queued',             s.queued],
    ['Errors',             s.error],
    ['', ''],
    ['Senders used',       Object.keys(senders).length]
  ];

  Object.keys(senders).sort().forEach(function(id) {
    out.push(['   ' + id, senders[id]]);
  });

  sheet.getRange(1, 1, out.length, 2).setValues(out);

  const header = sheet.getRange(1, 1, 1, 2);
  header.setBackground('#0f1117').setFontColor('#ffffff').setFontWeight('bold');
  sheet.setFrozenRows(1);
  sheet.setColumnWidth(1, 200);
  sheet.setColumnWidth(2, 160);
  sheet.getRange(1, 1, out.length, 1).setFontWeight('bold');

  SpreadsheetApp.flush();
}
/**
 * Recomputes every campaign's counters from masterLeads, on demand.
 *
 * Same work the nightly runSyncStats trigger does. Exposed as an action so the
 * app does not have to wait until 11:30 PM to show what was sent today.
 */
function handleSyncStats(params) {
  try {
    const settings = getSettings(MASTER_SHEET_ID);
    if (!settings.leadsSpreadsheetId) {
      return { success: false, error: 'leadsSpreadsheetId is not set in Settings' };
    }

    syncCampaignStats(MASTER_SHEET_ID, settings.leadsSpreadsheetId);

    log('handleSyncStats: campaign counters resynced on demand', 'INFO');
    return { success: true, message: 'Campaign counters are up to date.' };

  } catch (e) {
    log('handleSyncStats: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}
/**
 * Finds senders that have work queued but no chain running, and optionally
 * restarts them.
 *
 * A sender is "stalled" when its inbox holds Queued rows while no
 * processSingleSend / processSingleFollowup trigger exists. That combination
 * means the chain died and nothing will restart it on its own.
 *
 * @param {Object} params
 *   params.autoFix    - restart what is found. Default true.
 *   params.senderIds  - optional array to check only those senders.
 * @returns {{ success, checked, healthy, stalled, unreachable, restarted, senders }}
 */
function handleCheckSenderChains(params) {
  const autoFix = params.autoFix !== false;
  const only    = Array.isArray(params.senderIds) ? params.senderIds : null;

  try {
    let senders = getActiveSenders().filter(function(s) { return !!s.webAppUrl; });
    if (only) {
      senders = senders.filter(function(s) { return only.indexOf(s.emailID) !== -1; });
    }
    if (senders.length === 0) {
      return { success: true, checked: 0, healthy: 0, stalled: 0,
               unreachable: 0, restarted: 0, senders: [] };
    }

    const report = [];
    let healthy = 0, stalled = 0, unreachable = 0, restarted = 0;

    for (let b = 0; b < senders.length; b += 50) {
      const batch = senders.slice(b, b + 50);

      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(batch.map(function(s) {
          return {
            url: s.webAppUrl, method: 'post', contentType: 'application/json',
            payload: JSON.stringify({ action: 'getSenderStatus' }),
            muteHttpExceptions: true
          };
        }));
      } catch (e) {
        batch.forEach(function(s) {
          unreachable++;
          report.push({ emailID: s.emailID, state: 'unreachable', error: e.message });
        });
        continue;
      }

      // Collect the restarts for this batch and fire them together.
      const toRestart = [];

      for (let r = 0; r < responses.length; r++) {
        const s = batch[r];

        if (responses[r].getResponseCode() !== 200) {
          unreachable++;
          report.push({ emailID: s.emailID, state: 'unreachable',
                        error: 'HTTP ' + responses[r].getResponseCode() });
          continue;
        }

        let st;
        try {
          st = JSON.parse(responses[r].getContentText());
        } catch (e) {
          unreachable++;
          report.push({ emailID: s.emailID, state: 'unreadable', error: e.message });
          continue;
        }

        const names       = st.triggerNames || [];
        const coldQueued  = st.coldQueued     || 0;
        const followQueued= st.followupQueued || 0;
        const coldRunning = names.indexOf('processSingleSend') !== -1;
        const followRunning = names.indexOf('processSingleFollowup') !== -1;
        const atLimit     = (st.sentToday || 0) >= (st.dailyLimit || 25);

        const coldStalled   = coldQueued   > 0 && !coldRunning   && !atLimit;
        const followStalled = followQueued > 0 && !followRunning && !atLimit;

        if (!coldStalled && !followStalled) {
          healthy++;
          report.push({
            emailID: s.emailID, state: atLimit ? 'at daily limit' : 'ok',
            coldQueued: coldQueued, followupQueued: followQueued,
            sentToday: st.sentToday || 0, dailyLimit: st.dailyLimit || 0
          });
          continue;
        }

        stalled++;
        report.push({
          emailID: s.emailID, state: 'stalled',
          coldQueued: coldQueued, followupQueued: followQueued,
          coldStalled: coldStalled, followupStalled: followStalled,
          sentToday: st.sentToday || 0, dailyLimit: st.dailyLimit || 0
        });

        if (autoFix) {
          if (coldStalled)   { toRestart.push({ url: s.webAppUrl, act: 'kickoffCold' }); }
          if (followStalled) { toRestart.push({ url: s.webAppUrl, act: 'kickoffFollowup' }); }
        }
      }

      if (toRestart.length > 0) {
        try {
          UrlFetchApp.fetchAll(toRestart.map(function(t) {
            return {
              url: t.url, method: 'post', contentType: 'application/json',
              payload: JSON.stringify({ action: t.act }),
              muteHttpExceptions: true
            };
          }));
          restarted += toRestart.length;
        } catch (e) {
          log('handleCheckSenderChains: restart batch failed — ' + e.message, 'WARN');
        }
      }
    }

    log('handleCheckSenderChains: checked=' + senders.length +
        ' healthy=' + healthy + ' stalled=' + stalled +
        ' unreachable=' + unreachable + ' restarted=' + restarted, 'INFO');

    return {
      success: true,
      checked: senders.length,
      healthy: healthy, stalled: stalled,
      unreachable: unreachable, restarted: restarted,
      autoFixed: autoFix,
      senders: report
    };

  } catch (e) {
    log('handleCheckSenderChains: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}

/**
 * Restarts one sender's chains, whether or not it looks stalled.
 * Backs the per-row restart button on the Email Accounts page.
 */
function handleRestartSenderChain(params) {
  const emailID = String(params.emailID || '').trim();
  if (!emailID) { return { success: false, error: 'Missing emailID' }; }

  try {
    const sender = getActiveSenders().filter(function(s) {
      return s.emailID === emailID;
    })[0];

    if (!sender)            { return { success: false, error: 'Sender not found or not active: ' + emailID }; }
    if (!sender.webAppUrl)  { return { success: false, error: 'Sender has no webAppUrl: ' + emailID }; }

    const results = ['kickoffCold', 'kickoffFollowup'].map(function(act) {
      try {
        const res = UrlFetchApp.fetch(sender.webAppUrl, {
          method: 'post', contentType: 'application/json',
          payload: JSON.stringify({ action: act }),
          muteHttpExceptions: true
        });
        return { action: act, ok: res.getResponseCode() === 200 };
      } catch (e) {
        return { action: act, ok: false, error: e.message };
      }
    });

    const ok = results.every(function(r) { return r.ok; });
    log('handleRestartSenderChain: ' + emailID + ' — ' + (ok ? 'restarted' : 'partial'), 'INFO');

    return { success: true, emailID: emailID, restarted: ok, results: results };

  } catch (e) {
    log('handleRestartSenderChain: error — ' + e.message, 'ERROR');
    return { success: false, error: e.message };
  }
}
