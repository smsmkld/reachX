// ============================================================================
// replyScanner.gs — Cold Email Automation System
// Location: Master Control Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Scans Gmail inboxes of all active senders for replies to cold emails.
// Runs at 6:30 AM daily — BEFORE distributeCold() at 7:00 AM.
// This order is critical: replied leads must be marked before distribution
// so they are never queued for another send.
//
// PARALLEL FETCH ARCHITECTURE:
//  - All sender webhook requests are built into an array first
//  - UrlFetchApp.fetchAll() fires all requests simultaneously
//  - Responses are collected and processed in one pass
//  - For 200 senders, fetchAll() runs them in parallel batches of 50
//    (UrlFetchApp limit) via a chained trigger pattern so we never
//    exceed the 6-minute execution limit or the fetchAll cap.
//
// Flow:
//  1. Read active senders from senderAccounts sheet
//  2. Read masterLeads — build threadId → rowIndex lookup map in memory
//  3. Build all fetch requests at once (fetchRequests + fetchMeta arrays)
//  4. UrlFetchApp.fetchAll() fires all requests in parallel (batches of 50)
//  5. Collect all replied threadIds from responses
//  6. Batch update masterLeads: replyStatus='Replied', replyDate=today
//  7. Write all updates in ONE setValues() call
//
// Key design rules:
//  - Runs at 6:30 AM — before ANY distribution trigger
//  - Never marks a lead as Replied unless we find a genuine inbound message
//  - All sheet reads/writes in batch (getValues / setValues)
//  - Lock acquired before any shared write
//  - Date objects from Sheets always handled via formatDate()
//
// Dependencies:
//   utils.gs must be present in Master Control's Apps Script project.
// ============================================================================

// ============================================================================
// SECTION 1 — CONSTANTS
// ============================================================================
const REPLY_SCAN_DAYS     = 30;          // HARDCODED — verify this
const MAX_THREADS_PER_SENDER = 500;      // HARDCODED — verify this
const REPLY_LEADS_SHEET   = 'masterLeads';    // HARDCODED — verify this
const REPLY_SENDER_SHEET  = 'senderAccounts'; // HARDCODED — verify this
const REPLY_MASTER_SHEET_ID = MASTER_SHEET_ID; // Inherited from distributor.gs

// Max senders per fetchAll() batch — UrlFetchApp.fetchAll() processes up to
// 50 parallel requests reliably before response times degrade.
const FETCH_BATCH_SIZE = 50; // HARDCODED — verify this

// CacheService key used to pass batch progress between chained trigger runs
const SCAN_PROGRESS_KEY = 'replyScanProgress';

// ============================================================================
// SECTION 2 — scanAllSendersForReplies()
// Main entry point. Triggered at 6:30 AM daily from Master Control.
// Initialises the scan state and kicks off the first batch via
// processReplyBatch(). If more batches remain, chains a one-time trigger.
//
// Flow:
//  1. Acquire lock
//  2. Read active senders + build threadId map
//  3. Store full sender list + threadIds in CacheService (shared between batches)
//  4. Call processReplyBatch(0) for the first batch of up to 50 senders
//  5. If senders.length > 50, chain a trigger for the next batch
// ============================================================================
/**
 * Scans all active sender Gmail accounts for replies to cold emails.
 * Triggered at 6:30 AM daily from Master Control spreadsheet.
 * Batches 200 senders into groups of 50 via chained one-time triggers.
 */
function scanAllSendersForReplies() {
  log('scanAllSendersForReplies: started', 'INFO');

  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    log('scanAllSendersForReplies: could not acquire lock — exiting', 'WARN');
    return;
  }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    if (!leadsSheetId) {
      log('scanAllSendersForReplies: leadsSpreadsheetId not in Settings — exiting', 'ERROR');
      return;
    }

    // ── Step 1: Read active senders ─────────────────────────────────────────
    const senders = getActiveSendersForScan();
    if (senders.length === 0) {
      log('scanAllSendersForReplies: no active senders — exiting', 'INFO');
      return;
    }
    log('scanAllSendersForReplies: ' + senders.length + ' active senders found', 'INFO');

    // ── Step 2: Read masterLeads and build threadId map ──────────────────────
    const { threadMap } = buildThreadIdMap(leadsSheetId);
    if (Object.keys(threadMap).length === 0) {
      log('scanAllSendersForReplies: no threadIds in masterLeads — nothing to scan', 'INFO');
      return;
    }
    const knownThreadIds = Object.keys(threadMap);
    // Build ownerSender map for per-sender filtering
    const threadOwnerMap = {};
    Object.keys(threadMap).forEach(function(id) {
      threadOwnerMap[id] = threadMap[id].ownerSender || '';
    });
    log('scanAllSendersForReplies: threadMap has ' + knownThreadIds.length + ' entries', 'INFO');

    // ── Step 3: Store scan state in CacheService ─────────────────────────────
    // CacheService passes data between this run and any chained batch runs.
    // Max value size is 100KB — for large threadId lists, we pass the
    // leadsSheetId and re-read the map in each batch instead of caching it.
    // The sender list is deliberately NOT cached: at 200-300 senders the array
    // alone is 50-80KB of JSON and would push this value past the 100KB
    // CacheService limit once repliedThreadIds start accumulating. Each batch
    // re-reads it from senderAccounts instead.
    const scanState = {
      leadsSheetId:    leadsSheetId,
      repliedThreadIds: [],
      totalBatches:    Math.ceil(senders.length / FETCH_BATCH_SIZE),
      startedAt:       new Date().toISOString()
    };

    const cache = CacheService.getScriptCache();
    // Split into two keys if needed — CacheService max value = 100KB per key
    cache.put(SCAN_PROGRESS_KEY, JSON.stringify(scanState), 600); // 10 min TTL

    // ── Step 4: Process first batch (senders 0 to FETCH_BATCH_SIZE-1) ────────
    processReplyBatch_(0);

  } catch (e) {
    log('scanAllSendersForReplies: unexpected error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}

// ============================================================================
// SECTION 3 — processReplyBatch_()
// Processes one batch of up to FETCH_BATCH_SIZE senders in parallel.
// Called directly by scanAllSendersForReplies() for batch 0, then
// called by the chained trigger handler for subsequent batches.
//
// Flow:
//  1. Load scan state from CacheService
//  2. Slice senders for this batch
//  3. Build fetchRequests + fetchMeta arrays
//  4. UrlFetchApp.fetchAll() — fires all requests in parallel
//  5. Collect replies from responses
//  6. Append to repliedThreadIds in state
//  7. If more batches remain: chain a one-time trigger for next batch
//  8. If this was the last batch: write all updates to masterLeads
//
// @param {number} batchIndex - 0-based index of the batch to process.
// ============================================================================
/**
 * Processes one parallel batch of sender webhook calls.
 * Internal function — called by scanAllSendersForReplies() and chained triggers.
 *
 * @param {number} batchIndex - Which batch to process (0-based).
 */
function processReplyBatch_(batchIndex) {
  log('processReplyBatch_: processing batch ' + batchIndex, 'INFO');

  const cache     = CacheService.getScriptCache();
  const stateJson = cache.get(SCAN_PROGRESS_KEY);
  if (!stateJson) {
    log('processReplyBatch_: scan state expired from cache — cannot continue', 'ERROR');
    return;
  }

  let state;
  try {
    state = JSON.parse(stateJson);
  } catch (e) {
    log('processReplyBatch_: failed to parse scan state — ' + e.message, 'ERROR');
    return;
  }

  const { leadsSheetId, totalBatches } = state;
  const repliedThreadIds = state.repliedThreadIds || [];

  // Re-read the sender list every batch — it is not carried in the cached state
  const senders = getActiveSendersForScan();
  if (senders.length === 0) {
    log('processReplyBatch_: no active senders found — exiting', 'WARN');
    return;
  }

  // Re-read masterLeads fresh every batch — avoids CacheService size limit
  const { threadMap } = buildThreadIdMap(leadsSheetId);
  const knownThreadIds = Object.keys(threadMap);
  const threadOwnerMap = {};
  Object.keys(threadMap).forEach(function(id) {
    threadOwnerMap[id] = threadMap[id].ownerSender || '';
  });

  if (knownThreadIds.length === 0) {
    log('processReplyBatch_: no threadIds found in masterLeads — exiting', 'INFO');
    return;
  }

  // ── Slice senders for this batch ───────────────────────────────────────────
  const batchStart   = batchIndex * FETCH_BATCH_SIZE;
  const batchEnd     = Math.min(batchStart + FETCH_BATCH_SIZE, senders.length);
  const batchSenders = senders.slice(batchStart, batchEnd);

  log('processReplyBatch_: batch ' + batchIndex + ' — senders ' + batchStart + ' to ' + (batchEnd - 1), 'INFO');

  // ── Build fetchRequests + fetchMeta ───────────────────────────────────────
  const fetchRequests = [];
  const fetchMeta     = [];

  for (let i = 0; i < batchSenders.length; i++) {
    const sender = batchSenders[i];
    if (!sender.webAppUrl) {
      log('processReplyBatch_: sender ' + sender.emailID + ' has no webAppUrl — skipping', 'WARN');
      continue;
    }

    const senderThreadIds = knownThreadIds.filter(function(id) {
      return threadOwnerMap[id] === sender.emailID;
    });

    if (senderThreadIds.length === 0) {
      log('processReplyBatch_: sender ' + sender.emailID + ' has no threadIds — skipping', 'INFO');
      continue;
    }

    fetchRequests.push({
      url:         sender.webAppUrl,
      method:      'post',
      contentType: 'application/json',
      payload:     JSON.stringify({
        action:      'scanReplies',
        threadIds:   senderThreadIds,
        senderEmail: sender.emailAddress
      }),
      muteHttpExceptions: true
    });

    fetchMeta.push({
      emailID:      sender.emailID,
      emailAddress: sender.emailAddress
    });
  }

  // ── Fire all requests in parallel ─────────────────────────────────────────
  let batchReplies = [];

  if (fetchRequests.length > 0) {
    let responses;
    try {
      responses = UrlFetchApp.fetchAll(fetchRequests);
    } catch (e) {
      log('processReplyBatch_: fetchAll error on batch ' + batchIndex + ' — ' + e.message, 'ERROR');
      responses = [];
    }

    for (let r = 0; r < responses.length; r++) {
      const response = responses[r];
      const meta     = fetchMeta[r];

      try {
        const code = response.getResponseCode();
        if (code !== 200) {
          log('processReplyBatch_: non-200 (' + code + ') from sender ' + meta.emailID, 'ERROR');
          continue;
        }
        const result = JSON.parse(response.getContentText());
        // Stamp the owning sender onto every reply. The sender's handleScanReplies
        // returns only { threadId, replyDate }, but updateSenderAccountsReplyCounts_
        // groups by senderEmail — without this it groups nothing, matches no rows,
        // and still reports success.
        const senderReplies = (result.replies || []).map(function(rep) {
          return {
            threadId:    rep.threadId,
            replyDate:   rep.replyDate,
            senderEmail: rep.senderEmail || meta.emailAddress
          };
        });
        batchReplies = batchReplies.concat(senderReplies);
        log('processReplyBatch_: sender ' + meta.emailID + ' — ' + senderReplies.length + ' replies', 'INFO');
      } catch (e) {
        log('processReplyBatch_: error parsing response from ' + meta.emailID + ' — ' + e.message, 'ERROR');
      }
    }
  }

  // Accumulate across batches
  const allReplies = repliedThreadIds.concat(batchReplies);
  log('processReplyBatch_: batch ' + batchIndex + ' collected ' + batchReplies.length + ' replies (total so far: ' + allReplies.length + ')', 'INFO');

  // ── More batches remaining? Chain a trigger ────────────────────────────────
  const nextBatchIndex = batchIndex + 1;
  if (nextBatchIndex < totalBatches) {
    state.repliedThreadIds = allReplies;
    state.nextBatchIndex   = nextBatchIndex;  // the batch runNextReplyBatch_ must run next
    try {
      cache.put(SCAN_PROGRESS_KEY, JSON.stringify(state), 600);
    } catch (e) {
      // CacheService rejects values over 100KB — write what we have rather than
      // dying and leaving replied leads unmarked.
      log('processReplyBatch_: could not cache scan state — ' + e.message +
          ' — writing partial results', 'ERROR');
      writeReplyUpdates_(allReplies, leadsSheetId);
      return;
    }

    try {
      ScriptApp.newTrigger('runNextReplyBatch_')
        .timeBased()
        .after(30 * 1000)
        .create();
      log('processReplyBatch_: chained trigger created for batch ' + nextBatchIndex, 'INFO');
    } catch (e) {
      log('processReplyBatch_: failed to chain next trigger — ' + e.message + ' — attempting to write partial results', 'ERROR');
      writeReplyUpdates_(allReplies, leadsSheetId);
    }
    return;
  }

  // ── Last batch — write all updates to masterLeads ─────────────────────────
  log('processReplyBatch_: all batches complete — writing ' + allReplies.length + ' replies to masterLeads', 'INFO');
  writeReplyUpdates_(allReplies, leadsSheetId);

  cache.remove(SCAN_PROGRESS_KEY);
}

// ============================================================================
// SECTION 4 — runNextReplyBatch_()
// Chained trigger handler. Called by the one-time trigger created in
// processReplyBatch_() when more sender batches remain.
// Reads the current batch index from CacheService, increments it,
// deletes its own trigger, then calls processReplyBatch_().
// ============================================================================
/**
 * Chained trigger handler for multi-batch reply scanning.
 * Reads scan state, determines next batch index, then calls processReplyBatch_().
 * Deletes its own trigger on entry to avoid orphan buildup.
 */
function runNextReplyBatch_() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runNextReplyBatch_') {
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }

  const cache     = CacheService.getScriptCache();
  const stateJson = cache.get(SCAN_PROGRESS_KEY);
  if (!stateJson) {
    log('runNextReplyBatch_: scan state not found in cache — scan may have already completed', 'WARN');
    return;
  }

  let state;
  try {
    state = JSON.parse(stateJson);
  } catch (e) {
    log('runNextReplyBatch_: failed to parse scan state — ' + e.message, 'ERROR');
    return;
  }

  // Read nextBatchIndex from state — processReplyBatch_ sets it when chaining
  // Do NOT increment here — processReplyBatch_ owns the index progression
  const nextBatchIndex = (state.nextBatchIndex !== undefined) ? state.nextBatchIndex : 1;

  log('runNextReplyBatch_: running batch ' + nextBatchIndex, 'INFO');
  processReplyBatch_(nextBatchIndex);
}
// ============================================================================
// SECTION 5 — writeReplyUpdates_()
// Takes the full collected array of replied threadIds and writes all updates
// to masterLeads in ONE setValues() call. Called only by the final batch.
// ============================================================================
/**
 * Writes all reply updates to masterLeads in a single batch setValues() call.
 *
 * @param {Array}  repliedThreadIds - Array of { threadId, replyDate, senderEmail }
 * @param {string} leadsSheetId     - Google Sheets ID of the Leads Database.
 */
function writeReplyUpdates_(repliedThreadIds, leadsSheetId) {
  if (!repliedThreadIds || repliedThreadIds.length === 0) {
    log('writeReplyUpdates_: no replies to write', 'INFO');
    return;
  }

  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    log('writeReplyUpdates_: could not acquire lock — updates not written', 'ERROR');
    return;
  }

  try {
    const { threadMap, headers, data, sheet } = buildThreadIdMap(leadsSheetId);
    if (!sheet) {
      log('writeReplyUpdates_: could not open masterLeads — updates not written', 'ERROR');
      return;
    }

    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });
    const today = formatDate(new Date());
    let updateCount = 0;

    for (let r = 0; r < repliedThreadIds.length; r++) {
      const match    = repliedThreadIds[r];
      const threadId = match.threadId;
      const mapEntry = threadMap[threadId];
      if (!mapEntry) { continue; }

      const dataIdx           = mapEntry.rowIndex - 2; // 1-based → 0-based
      const currentReplyStatus = String(data[dataIdx][col['replyStatus']] || '').trim();

      // Skip if already marked — buildThreadIdMap should have filtered these
      // but double-check here in case of timing edge case
      if (currentReplyStatus === 'Replied') { continue; }

      data[dataIdx][col['replyStatus']] = 'Replied';
      data[dataIdx][col['replyDate']]   = match.replyDate || today;
      updateCount++;
    }

    // Write ALL changes in ONE setValues() call
    if (updateCount > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
      log('writeReplyUpdates_: wrote ' + updateCount + ' Replied updates to masterLeads', 'INFO');
      updateSenderAccountsReplyCounts_(repliedThreadIds);
    } else {
      log('writeReplyUpdates_: all matched threads were already marked Replied', 'INFO');
    }

  } catch (e) {
    log('writeReplyUpdates_: error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}

// ============================================================================
// SECTION 6 — buildThreadIdMap()
// Reads masterLeads once and builds an in-memory lookup map:
//   threadId → { rowIndex, leadId, replyStatus }
// Only includes rows where threadId is non-empty and replyStatus ≠ 'Replied'.
// Returns the raw data array and sheet reference for the caller to write back.
// ============================================================================
/**
 * Reads masterLeads and builds a threadId → row lookup map.
 *
 * @param {string} leadsSheetId
 * @returns {{ threadMap: Object, headers: Array, data: Array, sheet: Sheet }}
 */
function buildThreadIdMap(leadsSheetId) {
  const result = { threadMap: {}, headers: [], data: [], sheet: null };
  try {
    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName(REPLY_LEADS_SHEET);
    if (!sheet) {
      log('buildThreadIdMap: masterLeads sheet not found', 'ERROR');
      return result;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      log('buildThreadIdMap: no lead rows found', 'INFO');
      return result;
    }

    // Read all data in ONE call
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    if (!('threadId' in col)) {
      log('buildThreadIdMap: threadId column not found in masterLeads', 'ERROR');
      return result;
    }

    const threadMap = {};
    for (let i = 0; i < data.length; i++) {
      const threadId    = String(data[i][col['threadId']]    || '').trim();
      const replyStatus = String(data[i][col['replyStatus']] || '').trim();
      const leadId      = String(data[i][col['leadId']]      || '').trim();
      if (!threadId)               { continue; } // never sent
      if (replyStatus === 'Replied') { continue; } // already handled
      threadMap[threadId] = { 
        rowIndex:    i + 2, 
        leadId:      leadId, 
        replyStatus: replyStatus,
        ownerSender: String(data[i][col['ownerSender']] || '').trim()
      };
    }

    result.threadMap = threadMap;
    result.headers   = headers;
    result.data      = data;
    result.sheet     = sheet;
  } catch (e) {
    log('buildThreadIdMap: error — ' + e.message, 'ERROR');
  }
  return result;
}

// ============================================================================
// SECTION 7 — scanSenderGmailForReplies()
// Scans one sender's Gmail for threads that have inbound replies.
// This runs on the SENDER spreadsheet (not master) via the webApp
// 'scanReplies' action. Kept here for reference and local testing.
// ============================================================================
/**
 * Scans a single sender's Gmail for inbound replies to their cold emails.
 * Runs in the sender's Apps Script context (authenticated as that sender).
 *
 * @param {Object} sender    - { emailID, emailAddress }
 * @param {Object} threadMap - threadId → row lookup map
 * @returns {Array} Array of { threadId, replyDate, senderEmail }
 */
function scanSenderGmailForReplies(sender, threadMap) {
  const replies = [];
  try {
    const scanFromDate = new Date();
    scanFromDate.setDate(scanFromDate.getDate() - REPLY_SCAN_DAYS);
    const dateFilter = formatDate(scanFromDate).replace(/-/g, '/');
    const query      = 'in:inbox after:' + dateFilter;
    const threads    = GmailApp.search(query, 0, MAX_THREADS_PER_SENDER);

    log('scanSenderGmailForReplies: ' + threads.length + ' threads found for ' + sender.emailID, 'INFO');

    for (let t = 0; t < threads.length; t++) {
      const thread   = threads[t];
      const threadId = thread.getId();
      if (!(threadId in threadMap)) { continue; }

      const messages    = thread.getMessages();
      const senderEmail = sender.emailAddress.toLowerCase();
      let replyFound    = false;
      let replyDate     = '';

      for (let m = 0; m < messages.length; m++) {
        const msg  = messages[m];
        const from = msg.getFrom().toLowerCase();
        if (from.indexOf(senderEmail) === -1) {
          replyFound = true;
          const msgDate = msg.getDate();
          replyDate     = msgDate instanceof Date ? formatDate(msgDate) : formatDate(new Date());
          break;
        }
      }

      if (replyFound) {
        replies.push({ threadId: threadId, replyDate: replyDate, senderEmail: sender.emailAddress });
      }
    }
  } catch (e) {
    log('scanSenderGmailForReplies: error for ' + sender.emailID + ' — ' + e.message, 'ERROR');
  }
  return replies;
}

// ============================================================================
// SECTION 8 — getActiveSendersForScan()
// Reads senderAccounts and returns active senders for reply scanning.
// ============================================================================
/**
 * @returns {Array} Array of { emailID, emailAddress, webAppUrl, spreadsheetId }
 */
function getActiveSendersForScan() {
  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(REPLY_SENDER_SHEET);
  if (!sheet) {
    log('getActiveSendersForScan: senderAccounts sheet not found', 'ERROR');
    return [];
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) { return []; }

  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
  const col     = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  const senders = [];
  for (let i = 0; i < data.length; i++) {
    const isActive = String(data[i][col['isActive']] || '').trim();
    if (isActive !== 'Active') { continue; }

    const emailID       = String(data[i][col['emailID']]       || '').trim();
    const emailAddress  = String(data[i][col['emailAddress']]  || '').trim();
    const webAppUrl     = String(data[i][col['webAppUrl']]     || '').trim();
    const spreadsheetId = String(data[i][col['spreadsheetId']] || '').trim();

    if (!emailID || !emailAddress) { continue; }
    senders.push({ emailID, emailAddress, webAppUrl, spreadsheetId });
  }
  return senders;
}

// ============================================================================
// SECTION 9 — setupReplyScannerTrigger()
// One-time setup. Creates the 6:30 AM daily trigger for scanAllSendersForReplies().
// Also cleans up any leftover runNextReplyBatch_ triggers from prior runs.
// ============================================================================
/**
 * Creates the daily 6:30 AM trigger for scanAllSendersForReplies().
 * Run ONCE manually from Master Control Apps Script editor.
 * Safe to re-run — deletes existing triggers first.
 */
function setupReplyScannerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'scanAllSendersForReplies' || fn === 'runNextReplyBatch_') {
      ScriptApp.deleteTrigger(triggers[i]);
      log('setupReplyScannerTrigger: deleted trigger for ' + fn, 'INFO');
    }
  }

  ScriptApp.newTrigger('scanAllSendersForReplies')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(30)
    .inTimezone('Africa/Cairo') // HARDCODED — verify this
    .create();

  log('setupReplyScannerTrigger: created 6:30 AM Cairo daily trigger', 'INFO');
  Logger.log('✅ Reply scanner trigger created — runs at 6:30 AM Cairo time daily.');
}

// ============================================================================
// SECTION 10 — cleanupReplyScannerTriggers()
// Safety utility. Run manually if a chained batch gets stuck.
// Deletes all runNextReplyBatch_ triggers and clears the CacheService state.
// ============================================================================
/**
 * Emergency cleanup. Deletes all runNextReplyBatch_ one-time triggers
 * and removes the scan progress key from CacheService.
 * Run manually from the Apps Script editor if the batch chain gets stuck.
 */
function cleanupReplyScannerTriggers() {
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'runNextReplyBatch_') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }

  CacheService.getScriptCache().remove(SCAN_PROGRESS_KEY);

  log('cleanupReplyScannerTriggers: deleted ' + deleted + ' triggers, cleared cache state', 'INFO');
  Logger.log('✅ Cleanup done. Deleted ' + deleted + ' runNextReplyBatch_ triggers.\nCache state cleared. Safe to re-run scanAllSendersForReplies() manually.');
}

function updateSenderAccountsReplyCounts_(repliedThreadIds) {
  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(REPLY_SENDER_SHEET);
    if (!sheet || sheet.getLastRow() < 2) { return; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1,
                                   sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const repliesPerSender = {};
    let unattributed = 0;
    repliedThreadIds.forEach(function(match) {
      const sender = String(match.senderEmail || '').toLowerCase();
      if (!sender) { unattributed++; return; }
      repliesPerSender[sender] = (repliesPerSender[sender] || 0) + 1;
    });

    if (unattributed > 0) {
      log('updateSenderAccountsReplyCounts_: ' + unattributed + ' replies had no senderEmail — ' +
          'they cannot be counted against a sender', 'WARN');
    }

    let matchedRows = 0, added = 0;
    for (let i = 0; i < data.length; i++) {
      const email = String(data[i][col['emailAddress']] || '').trim().toLowerCase();
      if (!repliesPerSender[email]) { continue; }
      if ('repliesReceived' in col) {
        data[i][col['repliesReceived']] =
          (parseInt(data[i][col['repliesReceived']] || 0, 10)) +
          repliesPerSender[email];
        matchedRows++;
        added += repliesPerSender[email];
      }
    }

    if (matchedRows === 0) {
      log('updateSenderAccountsReplyCounts_: no senderAccounts rows matched — ' +
          'nothing written. Replies seen for: ' + (Object.keys(repliesPerSender).join(', ') || 'nobody') +
          ('repliesReceived' in col ? '' : ' (senderAccounts has no repliesReceived column)'), 'WARN');
      return;
    }

    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    SpreadsheetApp.flush();
    log('updateSenderAccountsReplyCounts_: added ' + added + ' replies across ' +
        matchedRows + ' sender row(s)', 'INFO');
  } catch (e) {
    log('updateSenderAccountsReplyCounts_: error — ' + e.message, 'ERROR');
  }
}
