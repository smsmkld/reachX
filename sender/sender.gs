// ============================================================================
// sender.gs — Cold Email Automation System
// Location: EACH Sender Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Handles cold (initial) email sending for one Gmail sender account.
//
// TRIGGER ARCHITECTURE (chained one-time triggers):
//  - scheduleDailySends() runs once at 7:30 AM via a fixed daily trigger.
//    It reads the first Queued row, creates ONE one-time trigger for it,
//    writes the triggerId to that row, and exits.
//  - processSingleSend() is called by each one-time trigger.
//    It sends one email, deletes itself, then creates the next one-time
//    trigger for the next Queued row (chained). Random delay 7–18 min.
//  - This means at most 1–2 active triggers exist at any time.
//    Never hits the Apps Script 20-trigger limit.
//
// Key design rules:
//  - status = 'Sending' set BEFORE Gmail call (crash guard, no duplicates)
//  - retryCount incremented on Error, stopped at MAX_RETRY_COUNT
//  - ALL sheet reads/writes done in batch (getValues / setValues)
//  - Templates cached via getCampaignTemplates() — never re-fetched mid-run
//  - Date columns from Sheets are Date objects — always use formatDate()
//  - reply() not replyAll() for all follow-up sends (enforced in followup.gs)
//  - Timezone: Africa/Cairo // HARDCODED — verify this
//
// Dependencies:
//   utils.gs must be in this sender spreadsheet's Apps Script project.
//   Functions used: sanitize(), replaceVariables(), getConfig(), formatDate(),
//   addDays(), acquireLock(), releaseLock(), log()
// ============================================================================

// ============================================================================
// SECTION 1 — CONSTANTS
// ============================================================================
const DAILY_INBOX_SHEET    = 'dailyInbox';
const FOLLOWUP_INBOX_SHEET = 'followupInbox';
const CONFIG_SHEET         = 'Config';
const MIN_GAP_MINUTES         = 6;

const COLD_COLS = {
  leadId:       'leadId',
  email:        'email',
  campaignId:   'campaignId',
  sequenceStep: 'sequenceStep',
  leadData:     'leadData',
  status:       'status',
  sentTime:     'sentTime',
  retryCount:   'retryCount',
  triggerId:    'triggerId'   // NEW — stores the one-time trigger ID for this row
};

const MAX_RETRY_COUNT  = 3;  // HARDCODED — verify this
const SENDER_LOCK_WAIT = 10; // seconds

// Random delay range between chained sends (milliseconds)
const MIN_CHAIN_DELAY_MS = 7  * 60 * 1000; //  7 minutes


// ============================================================================
// SECTION 2 — scheduleDailySends()
// Runs once at 7:30 AM via a fixed daily trigger.
// Kicks off the chain by scheduling the first send of the day.
//
// Flow:
//  1. Read Config — get dailyLimit, sentToday, masterSpreadsheetId
//  2. Check sending window — exit if not open yet
//  3. Check sentToday vs dailyLimit — exit if already at limit
//  4. Read dailyInbox — find first Queued row
//  5. Create ONE one-time trigger for processSingleSend
//  6. Write triggerId to that row
//  7. Exit
// ============================================================================
/**
 * Kicks off the daily send chain. Called once at 7:30 AM by a fixed trigger.
 * Schedules the first processSingleSend() one-time trigger of the day.
 */
function scheduleDailySends() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const config      = getConfig(configSheet);
  const inboxSheet  = ss.getSheetByName(DAILY_INBOX_SHEET);

  if (!inboxSheet) {
    log('scheduleDailySends: dailyInbox sheet not found', 'ERROR');
    return;
  }

  // ── Duplicate chain check ──────────────────────────────────────────────
  const existingTriggers    = ScriptApp.getProjectTriggers();
  const chainAlreadyRunning = existingTriggers.some(function(t) {
    return t.getHandlerFunction() === 'processSingleSend';
  });
  if (chainAlreadyRunning) {
    log('scheduleDailySends: chain already running — skipping duplicate', 'INFO');
    return;
  }
  // ──────────────────────────────────────────────────────────────────────

  // ── Daily limit check ─────────────────────────────────────────────────
  const dailyLimit = parseInt(config.dailyLimit || '25', 10);
  const sentToday  = parseInt(config.sentToday  || '0',  10);
  if (sentToday >= dailyLimit) {
    log('scheduleDailySends: daily limit reached (' + sentToday + '/' + dailyLimit + ') — exiting', 'INFO');
    return;
  }

  // ── Read all Queued rows ───────────────────────────────────────────────────
  const lastRow = inboxSheet.getLastRow();
  if (lastRow < 2) {
    log('scheduleDailySends: no rows in dailyInbox', 'INFO');
    return;
  }

  const headers = inboxSheet.getRange(1, 1, 1, inboxSheet.getLastColumn()).getValues()[0];
  const data    = inboxSheet.getRange(2, 1, lastRow - 1, inboxSheet.getLastColumn()).getValues();
  const colMap  = {};
  headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

  if (colMap['scheduledTime'] === undefined) {
    log('scheduleDailySends: scheduledTime column not found — please add it to dailyInbox', 'ERROR');
    return;
  }

  // ── Group Queued rows by campaignId ───────────────────────────────────────
  const campaignGroups = {};
  for (let i = 0; i < data.length; i++) {
    const status    = String(data[i][colMap['status']]     || '').trim();
    const campaignId = String(data[i][colMap['campaignId']] || '').trim();
    if (status !== 'Queued') { continue; }

    if (!campaignGroups[campaignId]) { campaignGroups[campaignId] = []; }
    campaignGroups[campaignId].push({ dataIndex: i, row: data[i] });
  }

  if (Object.keys(campaignGroups).length === 0) {
    log('scheduleDailySends: no Queued rows found', 'INFO');
    return;
  }

  // ── Generate schedule per campaign then merge ─────────────────────────────
  // allScheduledItems will be: [ { dataIndex, scheduledMinute }, ... ]
  const allScheduledItems = [];

  const campaignIds = Object.keys(campaignGroups);
  for (let c = 0; c < campaignIds.length; c++) {
    const campaignId = campaignIds[c];
    const group      = campaignGroups[campaignId];
    const firstRow   = group[0].row;

    // Read window from the row (embedded by distributor)
    const wsRaw       = firstRow[colMap['windowStart']];
    const weRaw       = firstRow[colMap['windowEnd']];
    const windowStart = (wsRaw instanceof Date)
      ? String(wsRaw.getHours()).padStart(2, '0') + ':' + String(wsRaw.getMinutes()).padStart(2, '0')
      : String(wsRaw || '').trim();
    const windowEnd   = (weRaw instanceof Date)
      ? String(weRaw.getHours()).padStart(2, '0') + ':' + String(weRaw.getMinutes()).padStart(2, '0')
      : String(weRaw || '').trim();

    if (!windowStart || !windowEnd) {
      log('scheduleDailySends: campaign ' + campaignId + ' has no windowStart/windowEnd — skipping', 'WARN');
      continue;
    }

    // Convert window to minutes from midnight
    const startMinutes = timeStringToMinutes_(windowStart);
    const endMinutes   = timeStringToMinutes_(windowEnd);
    const windowSize   = endMinutes - startMinutes;

    if (windowSize <= 0) {
      log('scheduleDailySends: campaign ' + campaignId + ' has invalid window — skipping', 'WARN');
      continue;
    }

    const count = group.length;

    // Generate random timestamps within the window
    // We generate more than we need then pick the best spread
    const randomMinutes = [];
    for (let i = 0; i < count; i++) {
      // Pure random within window
      randomMinutes.push(startMinutes + Math.floor(Math.random() * windowSize));
    }

    // Sort ascending
    randomMinutes.sort(function(a, b) { return a - b; });

    // Enforce minimum gap of MIN_GAP_MINUTES
    for (let i = 1; i < randomMinutes.length; i++) {
      if (randomMinutes[i] < randomMinutes[i - 1] + MIN_GAP_MINUTES) {
        randomMinutes[i] = randomMinutes[i - 1] + MIN_GAP_MINUTES;
      }
    }

    // Check if last send still fits in window — if not, compress backwards
    if (randomMinutes[randomMinutes.length - 1] > endMinutes) {
      // Redistribute: space evenly with min gap as fallback
      const interval = Math.max(MIN_GAP_MINUTES, Math.floor(windowSize / count));
      for (let i = 0; i < count; i++) {
        // Add small random jitter of up to 2 minutes per slot
        const jitter = Math.floor(Math.random() * 3) - 1; // -1, 0, or +1 minute
        randomMinutes[i] = startMinutes + (i * interval) + jitter;
      }
    }

    // Assign scheduled minutes to each row in this campaign group
    for (let i = 0; i < group.length; i++) {
      allScheduledItems.push({
        dataIndex:       group[i].dataIndex,
        scheduledMinute: randomMinutes[i] || startMinutes
      });
    }
  }

  if (allScheduledItems.length === 0) {
    log('scheduleDailySends: no items to schedule', 'INFO');
    return;
  }

  // ── Sort ALL items across ALL campaigns by scheduled time ─────────────────
  allScheduledItems.sort(function(a, b) { return a.scheduledMinute - b.scheduledMinute; });

  // ── Re-enforce 7 min gap across the merged full schedule ──────────────────
  // (a Campaign A send and Campaign B send on same Gmail account need gap too)
  for (let i = 1; i < allScheduledItems.length; i++) {
    if (allScheduledItems[i].scheduledMinute < allScheduledItems[i - 1].scheduledMinute + MIN_GAP_MINUTES) {
      allScheduledItems[i].scheduledMinute = allScheduledItems[i - 1].scheduledMinute + MIN_GAP_MINUTES;
    }
  }

  // ── Write scheduledTime back to sheet ────────────────────────────────────
  const campaignTimezone = data[Object.values(campaignGroups)[0][0].dataIndex][colMap['timezone']] || 'America/New_York';
  const today = Utilities.formatDate(new Date(), campaignTimezone, 'yyyy-MM-dd');
  log('DEBUG schedule: campaignTimezone=' + campaignTimezone + ' today=' + today + ' sampleTime=' + allScheduledItems[0].scheduledMinute, 'INFO');
  for (let i = 0; i < allScheduledItems.length; i++) {
    const item       = allScheduledItems[i];
    const hh         = String(Math.floor(item.scheduledMinute / 60)).padStart(2, '0');
    const mm         = String(item.scheduledMinute % 60).padStart(2, '0');
    const timeStr = today + 'T' + hh + ':' + mm + ':00'; // full datetime string
    data[item.dataIndex][colMap['scheduledTime']] = "'" + timeStr;
    // Also clear any old triggerId so chain starts fresh
    data[item.dataIndex][colMap['triggerId']] = '';
  }

  // Write entire data array back in ONE setValues() call
  inboxSheet.getRange(2, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();
  log('scheduleDailySends: wrote schedule for ' + allScheduledItems.length + ' rows', 'INFO');

  // ── Schedule first trigger ────────────────────────────────────────────────
  const firstItem    = allScheduledItems[0];
  const firstTimeStr = data[firstItem.dataIndex][colMap['scheduledTime']];
  const firstRow     = data[firstItem.dataIndex];
  const weRawCheck   = firstRow[colMap['windowEnd']];
  const windowEndStr = (weRawCheck instanceof Date)
    ? String(weRawCheck.getHours()).padStart(2, '0') + ':' + String(weRawCheck.getMinutes()).padStart(2, '0')
    : String(weRawCheck || '').trim();

  // Check window-end using the same timeapi.io call — no double fetch
  if (windowEndStr) {
    try {
      const url      = 'https://timeapi.io/api/time/current/zone?timeZone=' + encodeURIComponent(campaignTimezone);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json     = JSON.parse(response.getContentText());
      log('DEBUG timeapi response: ' + response.getContentText(), 'INFO');
      const nowHHmm  = String(json.hour).padStart(2,'0') + ':' + String(json.minute).padStart(2,'0');
      if (nowHHmm > windowEndStr) {
        log('scheduleDailySends: sending window already closed (' + nowHHmm + ' > ' + windowEndStr + ') — not scheduling', 'INFO');
        return;
      }
    } catch (e) {
      log('scheduleDailySends: could not verify window end — ' + e.message + ' — proceeding anyway', 'WARN');
    }
  }

  const firstMs  = localTimeToUtcMs_(firstTimeStr, campaignTimezone);
  const nowMs    = new Date().getTime();
  const rawDelay = firstMs - nowMs;

  // Get current time in campaign timezone to check window boundaries
  const tzNowStr   = Utilities.formatDate(new Date(), campaignTimezone, 'HH:mm');
  const tzNowParts = tzNowStr.split(':');
  const tzNowMin   = parseInt(tzNowParts[0], 10) * 60 + parseInt(tzNowParts[1], 10);

  // Get window boundaries from the first row
  const wsRawFirst = data[firstItem.dataIndex][colMap['windowStart']];
  const weRawFirst = data[firstItem.dataIndex][colMap['windowEnd']];
  const winStartStr = (wsRawFirst instanceof Date)
    ? String(wsRawFirst.getHours()).padStart(2,'0') + ':' + String(wsRawFirst.getMinutes()).padStart(2,'0')
    : String(wsRawFirst || '').trim();
  const winEndStr = (weRawFirst instanceof Date)
    ? String(weRawFirst.getHours()).padStart(2,'0') + ':' + String(weRawFirst.getMinutes()).padStart(2,'0')
    : String(weRawFirst || '').trim();
  const winStartMin = timeStringToMinutes_(winStartStr);
  const winEndMin   = timeStringToMinutes_(winEndStr);

  let delayMs;
  if (tzNowMin >= winStartMin && tzNowMin <= winEndMin) {
    // We are currently INSIDE the window — fire in 1 min
    log('scheduleDailySends: currently inside window (' + tzNowStr + ') — firing in 1 min', 'INFO');
    delayMs = 60 * 1000;
  } else if (rawDelay <= 0) {
    log('scheduleDailySends: scheduled time already passed — firing in 1 min', 'WARN');
    delayMs = 60 * 1000;
  } else if (rawDelay > 23 * 60 * 60 * 1000) {
    // More than 23h — target time already passed today, next is tomorrow
    if (tzNowMin <= winEndMin) {
      // Still inside today's window somehow
      log('scheduleDailySends: large delay but still in window — firing in 1 min', 'INFO');
      delayMs = 60 * 1000;
    } else {
      // Window closed for today — wait until tomorrow
      log('scheduleDailySends: window closed for today — waiting for tomorrow', 'INFO');
      delayMs = rawDelay;
    }
  } else {
    delayMs = rawDelay;
    log('scheduleDailySends: scheduling first send in ' + Math.round(delayMs/60000) + ' min', 'INFO');
  }

  const triggerId = createChainedTrigger_(delayMs);
  if (!triggerId) {
    log('scheduleDailySends: failed to create first trigger', 'ERROR');
    return;
  }

  // Write triggerId to first row
  const firstRowSheetIndex = firstItem.dataIndex + 2; // +2 for header and 0-based
  inboxSheet.getRange(firstRowSheetIndex, colMap['triggerId'] + 1).setValue(triggerId);
  SpreadsheetApp.flush();

  log('scheduleDailySends: first trigger created, fires in ' + Math.round(delayMs / 60000) + ' min', 'INFO');
}

// ============================================================================
// SECTION 3 — processSingleSend()
// Called by each one-time trigger. Sends ONE email then chains the next.
//
// Flow:
//  1. Identify own trigger ID → find matching row in dailyInbox
//  2. Delete own trigger (cleanup)
//  3. Check sending window — if closed, reset row to Queued and exit (no chain)
//  4. Check sentToday vs dailyLimit — exit if at limit (no chain)
//  5. Acquire lock
//  6. Set row status = 'Sending' (crash guard)
//  7. Build template from cache
//  8. Send email
//  9. On success: mark Sent, update sentToday, notify Master Control
// 10. On failure: increment retryCount, set Error or back to Queued
// 11. Find NEXT Queued row
// 12. If found: create next one-time trigger, write triggerId — chain continues
// 13. If not found: log done, exit
// ============================================================================
/**
 * Sends one email and chains the next trigger. Called by one-time triggers.
 */
function processSingleSend() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const config      = getConfig(configSheet);

    // ── Pre-flight window check ───────────────────────────────────────────────
    try {
      const preFlightInbox = SpreadsheetApp.getActiveSpreadsheet()
                              .getSheetByName(DAILY_INBOX_SHEET);
      if (preFlightInbox && preFlightInbox.getLastRow() > 1) {
        const firstQueued = findFirstQueuedRow_(preFlightInbox);
        if (firstQueued) {
          const pfWs = firstQueued.row[firstQueued.colMap['windowStart']];
          const pfWe = firstQueued.row[firstQueued.colMap['windowEnd']];
          const pfTz = String(firstQueued.row[firstQueued.colMap['timezone']] 
                      || 'Africa/Cairo').trim();
          const pfWindowSettings = {
            windowStart: (pfWs instanceof Date)
              ? String(pfWs.getHours()).padStart(2,'0') + ':' + 
                String(pfWs.getMinutes()).padStart(2,'0')
              : String(pfWs || '').trim(),
            windowEnd: (pfWe instanceof Date)
              ? String(pfWe.getHours()).padStart(2,'0') + ':' + 
                String(pfWe.getMinutes()).padStart(2,'0')
              : String(pfWe || '').trim(),
            sendingDays: String(
              firstQueued.row[firstQueued.colMap['sendingDays']] || '').trim(),
            timezone: pfTz
          };

          if (!isInsideSendingWindow_(config, pfWindowSettings)) {
            log('processSingleSend: PRE-FLIGHT — outside window — ' +
                'this is an orphan trigger — deleting ONLY myself and stopping', 'WARN');

            // Find and delete ONLY the trigger that matches our row's triggerId
            // so we don't kill the correctly scheduled 5PM trigger
            const myTriggerId = firstQueued.row[firstQueued.colMap['triggerId']];
            const allTriggers = ScriptApp.getProjectTriggers();
            allTriggers.forEach(function(t) {
              if (t.getHandlerFunction() === 'processSingleSend') {
                // Only delete if it does NOT match the stored triggerId
                // meaning it's an orphan not the correct one
                if (String(t.getUniqueId()) !== String(myTriggerId || '')) {
                  ScriptApp.deleteTrigger(t);
                  log('processSingleSend: deleted orphan trigger ' + 
                      t.getUniqueId(), 'WARN');
                }
              }
            });
            return;
          }
        }
      }
    } catch(e) {
      log('processSingleSend: pre-flight error — ' + e.message + 
          ' — continuing', 'WARN');
    }
    // ─────────────────────────────────────────────────────────────────────────

  // ── Identify which trigger just fired ────────────────────────────────────
  // We find our own trigger ID by matching against what's stored in the sheet.
  const inboxSheet = ss.getSheetByName(DAILY_INBOX_SHEET);
  if (!inboxSheet) {
    log('processSingleSend: dailyInbox sheet not found', 'ERROR');
    deleteOwnTrigger_(null); // attempt cleanup even without a known ID
    return;
  }

  const { rows, headers, colMap } = readInbox_(inboxSheet);
  if (!rows) {
    log('processSingleSend: could not read dailyInbox', 'ERROR');
    deleteOwnTrigger_(null);
    return;
  }

  // Find the row whose triggerId matches a currently-firing trigger
  const myRow = findMyRow_(rows, colMap);
  if (!myRow) {
    log('processSingleSend: could not find matching row for this trigger — cleaning up orphan triggers', 'WARN');
    deleteOwnTrigger_(null);
    return;
  }

  // ── Delete own trigger immediately ────────────────────────────────────────
  deleteOwnTrigger_(myRow.triggerId);

  // ── Sending window check ──────────────────────────────────────────────────
  // Extract embedded window settings from this row
  const rawWs = myRow.row[colMap['windowStart']];
  const rawWe = myRow.row[colMap['windowEnd']];
  const rowWindowSettings = {
    windowStart: (rawWs instanceof Date)
      ? (String(rawWs.getHours()).padStart(2,'0') + ':' + String(rawWs.getMinutes()).padStart(2,'0'))
      : String(rawWs || '').trim(),
    windowEnd: (rawWe instanceof Date)
      ? (String(rawWe.getHours()).padStart(2,'0') + ':' + String(rawWe.getMinutes()).padStart(2,'0'))
      : String(rawWe || '').trim(),
    sendingDays: String(myRow.row[colMap['sendingDays']] || '').trim(),
    timezone:    String(myRow.row[colMap['timezone']]    || '').trim()
  };
  if (!isInsideSendingWindow_(config, rowWindowSettings)) {
    log('processSingleSend: outside sending window — resetting row to Queued, stopping chain', 'INFO');
    inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.status] + 1).setValue('Queued');
    inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.triggerId] + 1).setValue('');
    return; // no chain — window closed
  }

  // ── Daily limit check ─────────────────────────────────────────────────────
  const dailyLimit = parseInt(config.dailyLimit || '25', 10);
  const sentToday  = parseInt(config.sentToday  || '0',  10);
  if (sentToday >= dailyLimit) {
    log('processSingleSend: daily limit reached (' + sentToday + '/' + dailyLimit + ') — stopping chain', 'INFO');
    inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.triggerId] + 1).setValue('');
    return; // no chain — limit hit
  }

  // ── Acquire lock ──────────────────────────────────────────────────────────
  const lock = acquireLock(SENDER_LOCK_WAIT);
  if (!lock) {
    log('processSingleSend: could not acquire lock — will retry on next chain', 'WARN');
    // Re-schedule self with a short delay so it retries
    const retryTriggerId = createChainedTrigger_(60 * 1000); // retry in 1 min
    if (retryTriggerId) {
      writeTriggerId_(inboxSheet, myRow.rowIndex, colMap, retryTriggerId);
    }
    return;
  }

  const updates = [];
  let   sendSucceeded = false;

  try {
    const leadId       = String(myRow.row[colMap[COLD_COLS.leadId]]       || '').trim();
    const toEmail      = String(myRow.row[colMap[COLD_COLS.email]]         || '').trim();
    const campaignId   = String(myRow.row[colMap[COLD_COLS.campaignId]]    || '').trim();
    const sequenceStep = parseInt(myRow.row[colMap[COLD_COLS.sequenceStep]] || 1, 10);
    const retryCount   = parseInt(myRow.row[colMap[COLD_COLS.retryCount]]   || 0, 10);
    const leadDataRaw  = String(myRow.row[colMap[COLD_COLS.leadData]]      || '{}');

    if (!toEmail || !campaignId) {
      log('processSingleSend: row ' + myRow.rowIndex + ' missing email or campaignId — skipping', 'WARN');
      inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.status]    + 1).setValue('Error');
      inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.triggerId] + 1).setValue('');
    } else {
      // ── Build template cache ────────────────────────────────────────────
      const templateCache = getCampaignTemplates(config.masterSpreadsheetId);
      const cacheKey      = campaignId + '_' + sequenceStep;
      const template      = templateCache[cacheKey] || null;
      const nextTemplate  = templateCache[campaignId + '_' + (sequenceStep + 1)] || null;

      if (!template) {
        log('processSingleSend: no template for ' + cacheKey + ' — skipping lead ' + leadId, 'WARN');
        inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.status]    + 1).setValue('Error');
        inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.triggerId] + 1).setValue('');
      } else {
        let leadObj = {};
        try { leadObj = JSON.parse(leadDataRaw); } catch (e) { leadObj = {}; }

        // ── Crash guard: set Sending BEFORE Gmail call ──────────────────
        inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.status]    + 1).setValue('Sending');
        inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.triggerId] + 1).setValue('');
        SpreadsheetApp.flush();

        try {
          const senderObj = {
            senderFirstName: config.senderFirstName || '',
            senderLastName:  config.senderLastName  || ''
          };
          const result = sendColdEmail(toEmail, template, leadObj, senderObj, config);

          // ── Success ────────────────────────────────────────────────────
          const nowStr    = formatDate(new Date());
          const awaitDays = (nextTemplate && nextTemplate.awaitDays) ? nextTemplate.awaitDays : 3;
          const nextSentDate = addDays(formatDate(new Date()), awaitDays);

          inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.status]   + 1).setValue('Sent');
          inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.sentTime] + 1).setValue(nowStr);

          // Increment sentToday in Config
          updateConfigValue(configSheet, 'sentToday', String(sentToday + 1));
          updateConfigValue(configSheet, 'totalSentCount', String(parseInt(config.totalSentCount || 0, 10) + 1));
          updateConfigValue(configSheet, 'lastRunTime', new Date().toISOString());

          updates.push({
            leadId:       leadId,
            campaignId:   campaignId,
            status:       'Sent',
            sentTime:     nowStr,
            threadId:     result.threadId || '',
            sequenceStep: sequenceStep,
            nextSentDate: nextSentDate
          });

          sendSucceeded = true;
          log('processSingleSend: sent to ' + toEmail + ' leadId=' + leadId, 'INFO');

        } catch (e) {
          // ── Send failure ───────────────────────────────────────────────
          log('processSingleSend: Gmail error for lead ' + leadId + ' — ' + e.message, 'ERROR');
          const newRetry  = retryCount + 1;
          const newStatus = newRetry >= MAX_RETRY_COUNT ? 'Error' : 'Queued';
          inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.status]     + 1).setValue(newStatus);
          inboxSheet.getRange(myRow.rowIndex, colMap[COLD_COLS.retryCount] + 1).setValue(newRetry);
        }
      }
    }

    // ── Notify Master Control ───────────────────────────────────────────────
    if (updates.length > 0 && config.masterWebhookUrl) {
      notifyMasterControl(config.masterWebhookUrl, updates, 'updateColdSent', config.senderId || '');
    }

    if (sendSucceeded) {
      clearTerminalRows(ss.getSheetByName(DAILY_INBOX_SHEET), ['Sent']);
    }

  } finally {
    releaseLock(lock);
  }

  // ── Chain: schedule next send based on scheduledTime column ──────────────
  const refreshed = readInbox_(inboxSheet);
  if (refreshed && refreshed.rows) {
    const nextQueued = findFirstQueuedRow_(inboxSheet);
    if (nextQueued) {
      const rawNextWs = nextQueued.row[nextQueued.colMap['windowStart']];
      const rawNextWe = nextQueued.row[nextQueued.colMap['windowEnd']];
      const nextRowWindowSettings = {
        windowStart: (rawNextWs instanceof Date)
          ? (String(rawNextWs.getHours()).padStart(2,'0') + ':' + String(rawNextWs.getMinutes()).padStart(2,'0'))
          : String(rawNextWs || '').trim(),
        windowEnd: (rawNextWe instanceof Date)
          ? (String(rawNextWe.getHours()).padStart(2,'0') + ':' + String(rawNextWe.getMinutes()).padStart(2,'0'))
          : String(rawNextWe || '').trim(),
        sendingDays: String(nextQueued.row[nextQueued.colMap['sendingDays']] || '').trim(),
        timezone:    String(nextQueued.row[nextQueued.colMap['timezone']]    || '').trim()
      };
      if (!isInsideSendingWindow_(config, nextRowWindowSettings)) {
        log('processSingleSend: window closed — stopping chain, no more sends today', 'INFO');
      } else {
        const rawScheduled = nextQueued.row[nextQueued.colMap['scheduledTime']];
        log('DEBUG rawScheduled: ' + rawScheduled + ' | type=' + typeof rawScheduled + ' | isDate=' + (rawScheduled instanceof Date), 'INFO');
        const rowTz = String(nextQueued.row[nextQueued.colMap['timezone']] || 'America/New_York').trim();
        let delayMs;
        try {
          let scheduledMs;
          if (rawScheduled instanceof Date) {
            scheduledMs = rawScheduled.getTime();
          } else {
            const scheduledTimeStr = String(rawScheduled || '').replace(/^'/, '').trim();
            scheduledMs = scheduledTimeStr ? localTimeToUtcMs_(scheduledTimeStr, rowTz) : NaN;
          }
          const nowMs = new Date().getTime();
          const rawDelay = scheduledMs - nowMs;
          const jitterMs = Math.floor(Math.random() * 7 * 60 * 1000);
          if (!isNaN(rawDelay) && rawDelay > MIN_GAP_MINUTES * 60 * 1000) {
            delayMs = Math.min(rawDelay + Math.floor(Math.random() * 3 * 60 * 1000), 6 * 60 * 60 * 1000);
          } else if (!isNaN(rawDelay) && rawDelay > 0) {
            delayMs = MIN_GAP_MINUTES * 60 * 1000 + jitterMs;
          } else {
            log('processSingleSend: scheduled time already passed — using min gap + jitter', 'WARN');
            delayMs = MIN_GAP_MINUTES * 60 * 1000 + jitterMs;
          }
        } catch (e) {
          log('processSingleSend: error calculating delay — ' + e.message + ' — using min gap', 'WARN');
          delayMs = MIN_GAP_MINUTES * 60 * 1000;
        }
        if (delayMs > 30 * 60 * 1000) {
          log('processSingleSend: delayMs ' + Math.round(delayMs/60000) + ' min exceeds 30 min cap — clamping', 'WARN');
          delayMs = MIN_GAP_MINUTES * 60 * 1000 + Math.floor(Math.random() * 7 * 60 * 1000);
        }
        log('processSingleSend: delayMs calculated = ' + Math.round(delayMs/60000) + ' min', 'INFO');
        const nextTrigId = createChainedTrigger_(delayMs);
        if (nextTrigId) {
          writeTriggerId_(inboxSheet, nextQueued.rowIndex, nextQueued.colMap, nextTrigId);
          log('processSingleSend: next send scheduled in ~' + Math.round(delayMs / 60000) + ' min, row=' + nextQueued.rowIndex, 'INFO');
        } else {
          log('processSingleSend: failed to create next trigger', 'ERROR');
        }
      }
    } else {
      log('processSingleSend: no more Queued rows — chain complete for today', 'INFO');
    }
  }
}

// ============================================================================
// SECTION 4 — PRIVATE HELPERS
// ============================================================================

/**
 * Checks whether current time is inside the sending window for a given inbox row.
 * Reads windowStart/windowEnd/sendingDays/timezone embedded in the inbox row first.
 * Falls back to global Settings if inbox row values are empty.
 *
 * @param {Object} config - Sender Config values.
 * @param {Object} rowSettings - Optional { windowStart, windowEnd, sendingDays, timezone }
 *                               embedded in the inbox row by distributor.
 * @returns {boolean}
 */
function isInsideSendingWindow_(config, rowSettings) {
  const masterSpreadsheetId = config.masterSpreadsheetId || '';
  rowSettings = rowSettings || {};

  try {
    // Try row-level settings first, fall back to global Settings
    let windowStart = rowSettings.windowStart || '';
    let windowEnd   = rowSettings.windowEnd   || '';
    let sendingDays = rowSettings.sendingDays || '';
    let timezone    = rowSettings.timezone    || 'America/New_York';

    // If any are missing, read global Settings as fallback
    if (!windowStart || !windowEnd) {
      if (!masterSpreadsheetId) { return true; }
      const masterSs      = SpreadsheetApp.openById(masterSpreadsheetId);
      const settingsSheet = masterSs.getSheetByName('Settings');
      if (!settingsSheet) { return true; }
      const settingsData = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
      const settingsMap  = {};
      settingsData.forEach(function(row) { settingsMap[String(row[0]).trim()] = String(row[1]).trim(); });
      windowStart = windowStart || settingsMap['sendingWindowStart'] || '';
      windowEnd   = windowEnd   || settingsMap['sendingWindowEnd']   || '';
      sendingDays = sendingDays || settingsMap['sendingDays']        || '';
      timezone    = timezone    || settingsMap['timeZone']           || 'America/New_York';
    }

    if (!windowStart || !windowEnd) { return true; } // can't check — allow

    const nowHHmm      = Utilities.formatDate(new Date(), timezone, 'HH:mm');
    const nowMinutes   = parseInt(nowHHmm.split(':')[0], 10) * 60 + parseInt(nowHHmm.split(':')[1], 10);
    const startMinutes = timeStringToMinutes_(windowStart);
    const endMinutes   = timeStringToMinutes_(windowEnd);

    log('isInsideSendingWindow_: now=' + nowHHmm + ' start=' + windowStart + ' end=' + windowEnd + ' nowMin=' + nowMinutes + ' startMin=' + startMinutes + ' endMin=' + endMinutes, 'INFO');

    if (nowMinutes < startMinutes || nowMinutes > endMinutes) { return false; }

    // Check sending days if set (e.g. "Mon,Tue,Wed,Thu,Fri")
    if (sendingDays) {
      const dayNames  = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
      const todayDay = Utilities.formatDate(new Date(), timezone, 'EEE');
      const allowedDays = sendingDays.split(',').map(function(d) { return d.trim(); });
      if (allowedDays.indexOf(todayDay) === -1) {
        log('isInsideSendingWindow_: today (' + todayDay + ') not in sendingDays (' + sendingDays + ')', 'INFO');
        return false;
      }
    }

    return true;

  } catch (e) {
    log('isInsideSendingWindow_: error — ' + e.message + ' — allowing send', 'WARN');
    return true;
  }
}

/**
 * Reads all rows from dailyInbox into memory.
 * Returns { rows, headers, colMap } or { rows: null } on failure.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inboxSheet
 * @returns {{ rows: Array[]|null, headers: string[], colMap: Object }}
 */
function readInbox_(inboxSheet) {
  try {
    const lastRow = inboxSheet.getLastRow();
    if (lastRow < 2) { return { rows: [], headers: [], colMap: {} }; }

    const headers = inboxSheet.getRange(1, 1, 1, inboxSheet.getLastColumn()).getValues()[0];
    const rows    = inboxSheet.getRange(2, 1, lastRow - 1, inboxSheet.getLastColumn()).getValues();
    const colMap  = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });
    return { rows: rows, headers: headers, colMap: colMap };
  } catch (e) {
    log('readInbox_: error — ' + e.message, 'ERROR');
    return { rows: null, headers: [], colMap: {} };
  }
}

/**
 * Finds the first row in dailyInbox with status = 'Queued' and no triggerId set.
 * Returns { rowIndex (1-based, sheet row), row, colMap } or null if none found.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inboxSheet
 * @returns {{ rowIndex: number, row: Array, colMap: Object } | null}
 */
function findFirstQueuedRow_(inboxSheet) {
  const { rows, colMap } = readInbox_(inboxSheet);
  if (!rows || rows.length === 0) { return null; }

  for (let i = 0; i < rows.length; i++) {
    const status     = String(rows[i][colMap[COLD_COLS.status]]     || '').trim();
    const triggerId  = String(rows[i][colMap[COLD_COLS.triggerId]]  || '').trim();
    const retryCount = parseInt(rows[i][colMap[COLD_COLS.retryCount]] || 0, 10);

    if (status === 'Queued' && !triggerId && retryCount < MAX_RETRY_COUNT) {
      return { rowIndex: i + 2, row: rows[i], colMap: colMap }; // +2 for header + 0-index
    }
  }
  return null;
}

/**
 * Matches the currently-firing trigger to a row in dailyInbox by triggerId column.
 * Iterates all project triggers, checks which one matches a triggerId in the sheet.
 * Returns { rowIndex, row, colMap, triggerId } or null.
 *
 * @param {Array[]} rows   - Sheet rows (from readInbox_).
 * @param {Object}  colMap - Column index map.
 * @returns {{ rowIndex: number, row: Array, colMap: Object, triggerId: string } | null}
 */
function findMyRow_(rows, colMap) {
  if (!rows || rows.length === 0) { return null; }
  if (colMap[COLD_COLS.triggerId] === undefined) { return null; }

  const activeTriggers = ScriptApp.getProjectTriggers();
  const activeTriggerIds = {};
  activeTriggers.forEach(function(t) { 
    activeTriggerIds[t.getUniqueId()] = true; 
  });

  // Also check properties — covers race condition where trigger already fired
  // and was removed from active list before we could read it
  const props = PropertiesService.getScriptProperties().getProperties();
  const pendingTriggerIds = {};
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('pendingTrigger_') === 0) {
      const id = key.replace('pendingTrigger_', '');
      pendingTriggerIds[id] = true;
    }
  });

  for (let i = 0; i < rows.length; i++) {
    const rowTriggerId = String(rows[i][colMap[COLD_COLS.triggerId]] || '').trim();
    if (!rowTriggerId) { continue; }
    if (activeTriggerIds[rowTriggerId] || pendingTriggerIds[rowTriggerId]) {
      // Clean up property now that we found it
      PropertiesService.getScriptProperties()
        .deleteProperty('pendingTrigger_' + rowTriggerId);
      return {
        rowIndex:  i + 2,
        row:       rows[i],
        colMap:    colMap,
        triggerId: rowTriggerId
      };
    }
  }
  return null;
}

/**
 * Creates a one-time time-based trigger for processSingleSend().
 * Fires after `delayMs` milliseconds from now.
 * Returns the trigger's unique ID string, or null on failure.
 *
 * @param {number} delayMs - Milliseconds from now to fire.
 * @returns {string|null}
 */
function createChainedTrigger_(delayMs) {
  try {
    const safeDelay = Math.max(delayMs, 60 * 1000);
    const clampedDelay = Math.min(safeDelay, 23 * 60 * 60 * 1000);
    const trigger = ScriptApp.newTrigger('processSingleSend')
      .timeBased()
      .after(clampedDelay)
      .create();
    const id = trigger.getUniqueId();
    // Store in properties so findMyRow_ can still find it after trigger fires
    PropertiesService.getScriptProperties()
      .setProperty('pendingTrigger_' + id, 'processSingleSend');
    return id;
  } catch (e) {
    log('createChainedTrigger_: error — ' + e.message, 'ERROR');
    return null;
  }
}

/**
 * Deletes the trigger matching the given unique ID.
 * If triggerId is null, attempts to delete any orphaned processSingleSend triggers.
 *
 * @param {string|null} triggerId
 */
function deleteOwnTrigger_(triggerId) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      const t = triggers[i];
      if (t.getHandlerFunction() !== 'processSingleSend') { continue; }
        if (triggerId === null) {
          log('deleteOwnTrigger_: called with null — skipping to avoid chain kill', 'WARN');
          return;
        }
        if (t.getUniqueId() === triggerId) {
          ScriptApp.deleteTrigger(t);
          return;
        }
    }
  } catch (e) {
    log('deleteOwnTrigger_: error — ' + e.message, 'WARN');
  }
}

/**
 * Writes a triggerId value into the triggerId column of a specific sheet row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inboxSheet
 * @param {number} rowIndex  - 1-based sheet row number.
 * @param {Object} colMap    - Column index map.
 * @param {string} triggerId - Trigger unique ID to write.
 */
function writeTriggerId_(inboxSheet, rowIndex, colMap, triggerId) {
  if (colMap[COLD_COLS.triggerId] === undefined) {
    log('writeTriggerId_: triggerId column not found in dailyInbox — please add it', 'ERROR');
    return;
  }
  inboxSheet.getRange(rowIndex, colMap[COLD_COLS.triggerId] + 1).setValue(triggerId);
}

// ============================================================================
// SECTION 5 — sendColdEmail()
// Builds and sends one cold email via GmailApp.
// Returns { threadId } on success. Throws on failure.
// ============================================================================
/**
 * Builds and sends a cold email.
 *
 * @param {string} toEmail
 * @param {Object} template  - { subject, htmlBody, plainBody }
 * @param {Object} leadObj   - Parsed lead data for variable replacement.
 * @param {Object} senderObj - { senderFirstName, senderLastName }
 * @param {Object} config    - Sender Config values.
 * @returns {{ threadId: string }}
 * @throws {Error} On any Gmail failure.
 */
function sendColdEmail(toEmail, template, leadObj, senderObj, config) {
  const subject = replaceVariables(template.subject,  leadObj, senderObj);
  const body    = replaceVariables(template.htmlBody, leadObj, senderObj);
  const name    = ((config.senderFirstName || '') + ' ' + (config.senderLastName || '')).trim();

  GmailApp.sendEmail(toEmail, subject, body, {
    name:    name,
    replyTo: config.replyToEmail || config.senderEmail || ''
  });

  const sent     = GmailApp.search('to:' + toEmail + ' subject:"' + subject + '" in:sent', 0, 1);
  const threadId = sent.length > 0 ? sent[0].getId() : '';
  return { threadId: threadId };
}

function toQuotedPrintable_(text) {
  // Split into lines, add QP soft breaks so Gmail doesn't hard-wrap
  return text.split('\n').map(function(line) {
    var out = '';
    while (line.length > 75) {
      out += line.substring(0, 75) + '=\n';
      line = line.substring(75);
    }
    return out + line;
  }).join('\n');
}

// ============================================================================
// SECTION 6 — getCampaignTemplates()
// Reads entire Campaigns sheet ONCE into a cache object.
// Keys: 'C001_1', 'C001_2' etc. (campaignId + '_' + sequenceStep)
// Values: { subject, htmlBody, plainBody, awaitDays }
// ============================================================================
/**
 * @param {string} masterSpreadsheetId
 * @returns {Object} Cache of templates keyed by 'campaignId_sequenceStep'
 */
function getCampaignTemplates(masterSpreadsheetId) {
  const cache = {};
  if (!masterSpreadsheetId) {
    log('getCampaignTemplates: masterSpreadsheetId is empty', 'ERROR');
    return cache;
  }
  try {
    const ss    = SpreadsheetApp.openById(masterSpreadsheetId);
    const sheet = ss.getSheetByName('Campaigns');
    if (!sheet) {
      log('getCampaignTemplates: Campaigns sheet not found', 'ERROR');
      return cache;
    }
    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return cache; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    for (let i = 0; i < data.length; i++) {
      const rowCampaignId = String(data[i][col['campaignId']]    || '').trim();
      const rowStep       = parseInt(data[i][col['sequenceStep']] || 0, 10);
      if (!rowCampaignId || !rowStep) { continue; }

      cache[rowCampaignId + '_' + rowStep] = {
        subject:   String(data[i][col['subjectLine']] || ''),
        htmlBody:  String(data[i][col['emailBody']]   || ''),
        plainBody: String(data[i][col['emailBody']]   || ''),
        awaitDays: parseInt(data[i][col['awaitDays']] || 0, 10)
      };
    }
    log('getCampaignTemplates: loaded ' + Object.keys(cache).length + ' templates', 'INFO');
  } catch (e) {
    log('getCampaignTemplates: error — ' + e.message, 'ERROR');
  }
  return cache;
}

// ============================================================================
// SECTION 7 — getTemplate()
// Kept for backward compatibility with followup.gs.
// Wraps getCampaignTemplates() — do not call inside a loop.
// ============================================================================
/**
 * @deprecated Use getCampaignTemplates() for batch sends.
 */
function getTemplate(masterSpreadsheetId, campaignId, sequenceStep) {
  const cache = getCampaignTemplates(masterSpreadsheetId);
  const key   = campaignId + '_' + sequenceStep;
  if (cache[key]) { return cache[key]; }
  log('getTemplate: no template for campaignId=' + campaignId + ' step=' + sequenceStep, 'WARN');
  return null;
}

// ============================================================================
// SECTION 8 — updateConfigValue()
// Writes a single key-value pair to the Config sheet.
// ============================================================================
/**
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet
 * @param {string} key
 * @param {string} value
 */
function updateConfigValue(configSheet, key, value) {
  if (!configSheet || !key) {
    log('updateConfigValue: missing configSheet or key', 'WARN');
    return;
  }
  const lastRow = configSheet.getLastRow();
  if (lastRow < 1) {
    configSheet.appendRow([key, value]);
    return;
  }
  const data = configSheet.getRange(1, 1, lastRow, 2).getValues();
  for (let i = 0; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      configSheet.getRange(i + 1, 2).setValue(value);
      return;
    }
  }
  configSheet.appendRow([key, value]);
  log('updateConfigValue: key "' + key + '" not found — appended new row', 'INFO');
}

// ============================================================================
// SECTION 9 — notifyMasterControl()
// POSTs send results to Master Control webApp in batches of 50.
// ============================================================================
/**
 * @param {string} masterWebhookUrl
 * @param {Array}  updates
 * @param {string} action
 * @param {string} senderId
 */
function notifyMasterControl(masterWebhookUrl, updates, action, senderId) {
  if (!masterWebhookUrl) {
    log('notifyMasterControl: masterWebhookUrl is empty', 'ERROR');
    return;
  }
  const BATCH_SIZE = 50;
  for (let i = 0; i < updates.length; i += BATCH_SIZE) {
    const batch   = updates.slice(i, i + BATCH_SIZE);
    const payload = JSON.stringify({ action: action, updates: batch, senderId: senderId || '' });
    const options = {
      method:             'post',
      contentType:        'application/json',
      payload:            payload,
      muteHttpExceptions: true
    };
    try {
      const response = UrlFetchApp.fetch(masterWebhookUrl, options);
      if (response.getResponseCode() !== 200) {
        log('notifyMasterControl: non-200 response ' + response.getResponseCode() + ' — saving to cache for retry', 'ERROR');
        savePendingUpdates_(masterWebhookUrl, batch, action, senderId);
      } else {
        log('notifyMasterControl: batch ' + (Math.floor(i / BATCH_SIZE) + 1) + ' sent OK', 'INFO');
      }
    } catch (e) {
      log('notifyMasterControl: fetch error — ' + e.message + ' — saving to cache for retry', 'ERROR');
      savePendingUpdates_(masterWebhookUrl, batch, action, senderId);
    }
  }
}

function savePendingUpdates_(masterWebhookUrl, updates, action, senderId) {
  try {
    const cache     = CacheService.getScriptCache();
    const key       = 'pendingUpdates_' + senderId + '_' + new Date().getTime();
    const value     = JSON.stringify({
      masterWebhookUrl: masterWebhookUrl,
      updates:          updates,
      action:           action,
      senderId:         senderId
    });
    // CacheService max value = 100KB, TTL = 2 hours
    cache.put(key, value, 7200);
    // Update the key index so retryPendingUpdates() can find all pending keys
    const indexKey   = 'pendingUpdatesIndex_' + (senderId || '');
    const existingIndex = cache.get(indexKey);
    const keyList    = existingIndex ? JSON.parse(existingIndex) : [];
    keyList.push(key);
    cache.put(indexKey, JSON.stringify(keyList), 7200);
    log('savePendingUpdates_: saved ' + updates.length + ' updates to cache key ' + key, 'INFO');

    // Schedule a retry trigger in 5-8 minutes
    const delayMs = (5 * 60 * 1000) + Math.floor(Math.random() * 3 * 60 * 1000);
    ScriptApp.newTrigger('retryPendingUpdates')
      .timeBased()
      .after(delayMs)
      .create();
    log('savePendingUpdates_: retry trigger scheduled in ' + Math.round(delayMs / 60000) + ' min', 'INFO');
  } catch (e) {
    log('savePendingUpdates_: error saving to cache — ' + e.message, 'ERROR');
  }
}

function retryPendingUpdates() {
  // Delete own trigger first
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'retryPendingUpdates') {
      ScriptApp.deleteTrigger(triggers[i]);
      break;
    }
  }

  const cache = CacheService.getScriptCache();
  const ss    = SpreadsheetApp.getActiveSpreadsheet();
  const config = getConfig(ss.getSheetByName(CONFIG_SHEET));

  // List all pending keys — we stored them with prefix 'pendingUpdates_'
  // CacheService has no list() — we store a key index separately
  const indexKey  = 'pendingUpdatesIndex_' + (config.senderId || '');
  const indexJson = cache.get(indexKey);
  const keys      = indexJson ? JSON.parse(indexJson) : [];

  if (keys.length === 0) {
    log('retryPendingUpdates: no pending updates found in cache', 'INFO');
    return;
  }

  const stillFailing = [];

  for (let k = 0; k < keys.length; k++) {
    const key      = keys[k];
    const valueStr = cache.get(key);
    if (!valueStr) { continue; } // expired

    let pending;
    try { pending = JSON.parse(valueStr); } catch (e) { continue; }

    const options = {
      method:             'post',
      contentType:        'application/json',
      payload:            JSON.stringify({
        action:   pending.action,
        updates:  pending.updates,
        senderId: pending.senderId
      }),
      muteHttpExceptions: true
    };

    try {
      const response = UrlFetchApp.fetch(pending.masterWebhookUrl, options);
      if (response.getResponseCode() === 200) {
        cache.remove(key);
        log('retryPendingUpdates: key ' + key + ' succeeded — removed from cache', 'INFO');
      } else {
        log('retryPendingUpdates: key ' + key + ' still failing (' + response.getResponseCode() + ') — will retry again', 'WARN');
        stillFailing.push(key);
      }
    } catch (e) {
      log('retryPendingUpdates: fetch error for key ' + key + ' — ' + e.message, 'WARN');
      stillFailing.push(key);
    }
  }

  if (stillFailing.length > 0) {
    // Update index with only the still-failing keys
    cache.put(indexKey, JSON.stringify(stillFailing), 7200);
    // Schedule another retry
    const delayMs = (5 * 60 * 1000) + Math.floor(Math.random() * 3 * 60 * 1000);
    ScriptApp.newTrigger('retryPendingUpdates')
      .timeBased()
      .after(delayMs)
      .create();
    log('retryPendingUpdates: ' + stillFailing.length + ' still failing — retry scheduled in ' + Math.round(delayMs / 60000) + ' min', 'WARN');
  } else {
    cache.remove(indexKey);
    log('retryPendingUpdates: all pending updates resolved', 'INFO');
  }
}

// ============================================================================
// SECTION 10 — setupSenderTrigger()
// One-time setup. Creates the 7:30 AM daily trigger for scheduleDailySends().
// Also cleans up any leftover processSingleSend triggers from previous runs.
// Run manually once after deploying to a sender spreadsheet.
// ============================================================================
/**
 * Creates the daily 7:30 AM trigger for scheduleDailySends().
 * Safe to re-run — deletes existing triggers for both functions first.
 */
function setupSenderTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'scheduleDailySends' || fn === 'processSingleSend') {
      ScriptApp.deleteTrigger(triggers[i]);
      log('setupSenderTrigger: deleted trigger for ' + fn, 'INFO');
    }
  }

  // Fixed daily trigger at 7:30 AM for scheduleDailySends
  ScriptApp.newTrigger('scheduleDailySends')
    .timeBased()
    .atHour(14)       // 2 PM Cairo = 7 AM New York (EDT)
    .nearMinute(0)
    .everyDays(1)
    .create();

  log('setupSenderTrigger: created daily 7:30 AM trigger for scheduleDailySends', 'INFO');
  Logger.log('✅ Trigger created — scheduleDailySends will run daily at 7:30 AM.');
}

// ============================================================================
// SECTION 11 — setupDailyInboxSheet()
// One-time setup. Creates dailyInbox sheet with correct headers.
// NOW INCLUDES triggerId column.
// ============================================================================
/**
 * Creates the dailyInbox sheet with headers including triggerId.
 * Safe to re-run — skips if sheet already exists.
 */
function setupDailyInboxSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(DAILY_INBOX_SHEET);
  if (sheet) {
    Logger.log('dailyInbox sheet already exists — no changes made.\n\nIf upgrading an existing sheet, manually add a "triggerId" column header after "retryCount".');
    return;
  }
  sheet = ss.insertSheet(DAILY_INBOX_SHEET);
  const headers = [
    'leadId', 'email', 'campaignId', 'sequenceStep',
    'leadData', 'windowStart', 'windowEnd', 'sendingDays', 'timezone', 'status', 'sentTime', 'retryCount', 'triggerId', 'scheduledTime'
  ];
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#1a73e8');
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  log('setupDailyInboxSheet: created!', 'INFO');
  Logger.log('✅ dailyInbox sheet created successfully.');
}

// ============================================================================
// SECTION 12 — cleanupOrphanTriggers()
// Safety utility. Run manually if something goes wrong mid-day.
// Deletes all processSingleSend triggers and clears all triggerId cells.
// ============================================================================
/**
 * Emergency cleanup. Deletes all processSingleSend triggers and clears
 * triggerId column in dailyInbox. Resets any 'Sending' rows back to 'Queued'.
 * Run manually from the Apps Script editor if the chain gets stuck.
 */
function cleanupOrphanTriggers() {
  // Delete all processSingleSend triggers
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processSingleSend') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }

  // Clear triggerId column and reset Sending → Queued in dailyInbox
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const inboxSheet = ss.getSheetByName(DAILY_INBOX_SHEET);
  if (inboxSheet && inboxSheet.getLastRow() > 1) {
    const { rows, colMap } = readInbox_(inboxSheet);
    if (rows && rows.length > 0) {
      for (let i = 0; i < rows.length; i++) {
        const sheetRow = i + 2;
        // Clear triggerId
        if (colMap[COLD_COLS.triggerId] !== undefined) {
          inboxSheet.getRange(sheetRow, colMap[COLD_COLS.triggerId] + 1).setValue('');
        }
        // Reset Sending → Queued
        const status = String(rows[i][colMap[COLD_COLS.status]] || '').trim();
        if (status === 'Sending') {
          inboxSheet.getRange(sheetRow, colMap[COLD_COLS.status] + 1).setValue('Queued');
        }
      }
    }
  }

  log('cleanupOrphanTriggers: deleted ' + deleted + ' triggers, cleared triggerId column', 'INFO');
}

// ============================================================================
// SECTION 13 — scanRepliesLocally()
// Scans Gmail inbox for replies to known thread IDs. Unchanged.
// ============================================================================
function scanRepliesLocally(knownThreadIds) {
  const replies     = [];
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const config      = getConfig(ss.getSheetByName('Config'));
  const senderEmail = (config.senderEmail || '').toLowerCase();
  const scanFromDate = new Date();
  scanFromDate.setDate(scanFromDate.getDate() - 30);
  const dateFilter = formatDate(scanFromDate).replace(/-/g, '/');
  const threads    = GmailApp.search('in:inbox after:' + dateFilter, 0, 500);
  const threadIdSet = {};
  knownThreadIds.forEach(function(id) { threadIdSet[id] = true; });

  for (let t = 0; t < threads.length; t++) {
    const thread   = threads[t];
    const threadId = thread.getId();
    if (!threadIdSet[threadId]) { continue; }
    const messages = thread.getMessages();
    for (let m = 0; m < messages.length; m++) {
      const msg = messages[m];
      if (msg.getFrom().toLowerCase().indexOf(senderEmail) === -1) {
        replies.push({ threadId: threadId, replyDate: formatDate(msg.getDate()) });
        break;
      }
    }
  }
  return replies;
}

/**
 * Converts a 'HH:mm' time string to total minutes from midnight.
 * e.g. '08:30' → 510
 *
 * @param {string} timeStr - 'HH:mm' format
 * @returns {number} minutes from midnight
 */
function timeStringToMinutes_(value) {
  if (!value) { return 0; }
  // Handle Date objects (Sheets auto-converts HH:mm to Date)
  if (value instanceof Date) {
    return value.getHours() * 60 + value.getMinutes();
  }
  // Handle string like '08:30'
  const str   = String(value).trim();
  const parts = str.split(':');
  if (parts.length < 2) { return 0; }
  return parseInt(parts[0], 10) * 60 + parseInt(parts[1], 10);
}

/**
 * Converts a local datetime string to a UTC millisecond timestamp.
 * Handles DST automatically by deriving the current UTC offset
 * from the target timezone using Utilities.formatDate().
 *
 * Works correctly regardless of the Apps Script project timezone setting.
 *
 * @param {string} localTimeStr - 'yyyy-MM-ddTHH:mm:ss' or 'yyyy-MM-dd HH:mm:ss'
 * @param {string} timezone     - IANA timezone, e.g. 'America/New_York'
 * @returns {number}            - UTC milliseconds, or NaN on parse failure
 */
function localTimeToUtcMs_(localTimeStr, timezone) {
  if (!localTimeStr || !timezone) { return NaN; }

  const normalised = String(localTimeStr).trim().replace(' ', 'T');
  const match = normalised.match(/T(\d{2}):(\d{2})/);
  if (!match) {
    log('localTimeToUtcMs_: could not parse time from "' + localTimeStr + '"', 'ERROR');
    return NaN;
  }
  const targetHour   = parseInt(match[1], 10);
  const targetMinute = parseInt(match[2], 10);

  const MAX_RETRIES = 5;
  const RETRY_DELAY_MS = 2000; // 2 seconds between retries

  for (let attempt = 1; attempt <= MAX_RETRIES; attempt++) {
    try {
      const url      = 'https://timeapi.io/api/time/current/zone?timeZone=' + encodeURIComponent(timezone);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json     = JSON.parse(response.getContentText());

      const nowHour   = parseInt(json.hour,    10);
      const nowMinute = parseInt(json.minute,  10);
      const nowSecond = parseInt(json.seconds, 10);

      const nowTotalMinutes    = nowHour * 60 + nowMinute;
      const targetTotalMinutes = targetHour * 60 + targetMinute;
      let diffMinutes          = targetTotalMinutes - nowTotalMinutes;
      if (diffMinutes < 0) { diffMinutes += 24 * 60; }

      const diffMs = (diffMinutes * 60 - nowSecond) * 1000;
      log('localTimeToUtcMs_: tz=' + timezone + ' now=' + nowHour + ':' + String(nowMinute).padStart(2,'0') + ' target=' + targetHour + ':' + String(targetMinute).padStart(2,'0') + ' diffMs=' + diffMs + ' (attempt ' + attempt + ')', 'INFO');
      return new Date().getTime() + diffMs;

    } catch (e) {
      log('localTimeToUtcMs_: attempt ' + attempt + '/' + MAX_RETRIES + ' failed — ' + e.message, 'WARN');
      if (attempt < MAX_RETRIES) {
        Utilities.sleep(RETRY_DELAY_MS);
      }
    }
  }

  log('localTimeToUtcMs_: all ' + MAX_RETRIES + ' attempts failed — returning NaN', 'ERROR');
  return NaN;
}



function testLocalTimeToUtcMs() {
  const tz     = 'America/New_York';
  const now    = new Date();
  // Build a target 30 minutes from now in HH:mm
  const target = new Date(now.getTime() + 30 * 60 * 1000);
  const hh     = String(target.getUTCHours()).padStart(2, '0');   // rough, just for test
  const mm     = String(target.getUTCMinutes()).padStart(2, '0');
  const fakeStr = '2026-01-01T' + hh + ':' + mm + ':00';

  const result = localTimeToUtcMs_(fakeStr, tz);
  const diffMin = Math.round((result - now.getTime()) / 60000);
  Logger.log('diffMin should be ~30: ' + diffMin);
}





function testDelayCalc() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const inboxSheet = ss.getSheetByName(DAILY_INBOX_SHEET);
  const nextQueued = findFirstQueuedRow_(inboxSheet);
  
  if (!nextQueued) {
    Logger.log('No Queued rows found');
    return;
  }

  const rawScheduled = nextQueued.row[nextQueued.colMap['scheduledTime']];
  Logger.log('rawScheduled value: ' + rawScheduled);
  Logger.log('is Date object: ' + (rawScheduled instanceof Date));

  // ADD THIS:
  const rowTz = String(nextQueued.row[nextQueued.colMap['timezone']] || 'America/New_York').trim();
  Logger.log('timezone: ' + rowTz);
  const result = localTimeToUtcMs_(rawScheduled, rowTz);
  Logger.log('localTimeToUtcMs_ result: ' + result);
  Logger.log('nowMs: ' + new Date().getTime());
  Logger.log('Delay in minutes: ' + Math.round((result - new Date().getTime()) / 60000));
}