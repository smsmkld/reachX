// ============================================================================
// followup.gs — Cold Email Automation System
// Location: EACH Sender Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Handles follow-up email sending for one Gmail sender account.
//
// TRIGGER ARCHITECTURE (chained one-time triggers — mirrors sender.gs):
//  - scheduleFollowupSends() runs once at 8:00 AM via a fixed daily trigger.
//    It reads the first Queued row in followupInbox, creates ONE one-time
//    trigger for it, writes the triggerId to that row, and exits.
//  - processSingleFollowup() is called by each one-time trigger.
//    It sends one follow-up, deletes itself, then creates the next one-time
//    trigger for the next Queued row (chained). Random delay 7–18 min.
//  - At most 1–2 active triggers exist at any time.
//    Never hits the Apps Script 20-trigger limit.
//
// Key design rules:
//  - status = 'Sending' set BEFORE Gmail call (crash guard, no duplicates)
//  - Uses thread.reply() NOT replyAll() — threads into original cold email
//  - threadId is REQUIRED — rows without threadId are skipped with WARN
//  - retryCount incremented on Error, stopped at MAX_RETRY_COUNT (3)
//  - ALL sheet reads/writes done in batch (getValues / setValues)
//  - Templates cached via getCampaignTemplates() — never re-fetched mid-run
//  - Date objects from Sheets always converted via formatDate()
//  - sequenceStatus set to 'Completed' when sequenceStep = totalSequenceSteps
//  - Timezone: Africa/Cairo // HARDCODED — verify this
//
// Dependencies:
//   utils.gs must be present in this sender spreadsheet's Apps Script project.
//   sender.gs must also be present — shares: isInsideSendingWindow_(),
//   readInbox_(), findMyRow_(), createChainedTrigger_(), deleteOwnTrigger_(),
//   writeTriggerId_(), getCampaignTemplates(), updateConfigValue(),
//   notifyMasterControl(), MAX_RETRY_COUNT, SENDER_LOCK_WAIT,
//   MIN_CHAIN_DELAY_MS, MAX_CHAIN_DELAY_MS, CONFIG_SHEET,
//   FOLLOWUP_INBOX_SHEET
// ============================================================================

// ============================================================================
// SECTION 1 — CONSTANTS
// followupInbox column names. Must match sheet headers exactly.
// NOTE: Sheet-name constants (CONFIG_SHEET, FOLLOWUP_INBOX_SHEET) and
// execution constants (MAX_RETRY_COUNT, SENDER_LOCK_WAIT, MIN_CHAIN_DELAY_MS,
// MAX_CHAIN_DELAY_MS) are declared in sender.gs — do not redeclare here.
// ============================================================================
const FOLLOWUP_COLS = {
  leadId:             'leadId',
  email:              'email',
  campaignId:         'campaignId',
  sequenceStep:       'sequenceStep',
  totalSequenceSteps: 'totalSequenceSteps',
  leadData:           'leadData',
  threadId:           'threadId',
  status:             'status',
  lastSentTime:       'lastSentTime',
  retryCount:         'retryCount',
  triggerId:          'triggerId',
  scheduledTime:      'scheduledTime'  // NEW — stores one-time trigger ID for this row
}; // HARDCODED — verify this

// ============================================================================
// SECTION 2 — scheduleFollowupSends()
// Runs once at 8:00 AM via a fixed daily trigger (30 min after cold scheduler).
// Kicks off the follow-up chain by scheduling the first send of the day.
//
// Flow:
//  1. Read Config — get dailyLimit, sentToday, masterSpreadsheetId
//  2. Check sending window — exit if not open yet
//  3. Check sentToday vs dailyLimit — exit if already at limit
//  4. Read followupInbox — find first Queued row
//  5. Create ONE one-time trigger for processSingleFollowup
//  6. Write triggerId to that row
//  7. Exit
// ============================================================================
/**
 * Kicks off the daily follow-up send chain. Called once at 8:00 AM.
 * Schedules the first processSingleFollowup() one-time trigger of the day.
 */
function scheduleFollowupSends() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const config      = getConfig(configSheet);
  const inboxSheet  = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);

  if (!inboxSheet) {
    log('scheduleFollowupSends: followupInbox sheet not found', 'ERROR');
    return;
  }

  // ── Duplicate chain check ──────────────────────────────────────────────
  const existingTriggers  = ScriptApp.getProjectTriggers();
  const followupTriggers  = existingTriggers.filter(function(t) {
    return t.getHandlerFunction() === 'processSingleFollowup';
  });

  if (followupTriggers.length > 1) {
    followupTriggers.forEach(function(t) { ScriptApp.deleteTrigger(t); });
    log('scheduleFollowupSends: deleted ' + followupTriggers.length + 
        ' orphan triggers — starting fresh', 'WARN');
  } else if (followupTriggers.length === 1) {
    log('scheduleFollowupSends: chain already running — skipping duplicate', 'INFO');
    return;
  }
  // ──────────────────────────────────────────────────────────────────────

  // ── Daily limit check ─────────────────────────────────────────────────────
  const dailyLimit = parseInt(config.dailyLimit || '25', 10);
  const sentToday  = parseInt(config.sentToday  || '0',  10);
  if (sentToday >= dailyLimit) {
    log('scheduleFollowupSends: daily limit reached (' + sentToday + '/' + dailyLimit + ') — exiting', 'INFO');
    return;
  }

  // ── Read all Queued rows ───────────────────────────────────────────────────
  const lastRow = inboxSheet.getLastRow();
  if (lastRow < 2) {
    log('scheduleFollowupSends: no rows in followupInbox', 'INFO');
    return;
  }

  const headers = inboxSheet.getRange(1, 1, 1, inboxSheet.getLastColumn()).getValues()[0];
  const data    = inboxSheet.getRange(2, 1, lastRow - 1, inboxSheet.getLastColumn()).getValues();
  const colMap  = {};
  headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });

  if (colMap['scheduledTime'] === undefined) {
    log('scheduleFollowupSends: scheduledTime column not found — please add it to followupInbox', 'ERROR');
    return;
  }

  // ── Group Queued rows by campaignId ───────────────────────────────────────
  const campaignGroups = {};
  for (let i = 0; i < data.length; i++) {
    const status     = String(data[i][colMap['status']]      || '').trim();
    const campaignId = String(data[i][colMap['campaignId']]  || '').trim();
    if (status !== 'Queued') { continue; }

    if (!campaignGroups[campaignId]) { campaignGroups[campaignId] = []; }
    campaignGroups[campaignId].push({ dataIndex: i, row: data[i] });
  }

  if (Object.keys(campaignGroups).length === 0) {
    log('scheduleFollowupSends: no Queued rows found', 'INFO');
    return;
  }

  // ── Generate schedule per campaign then merge ─────────────────────────────
  const allScheduledItems = [];

  const campaignIds = Object.keys(campaignGroups);
  for (let c = 0; c < campaignIds.length; c++) {
    const campaignId = campaignIds[c];
    const group      = campaignGroups[campaignId];
    const firstRow   = group[0].row;

    const wsRaw       = firstRow[colMap['windowStart']];
    const weRaw       = firstRow[colMap['windowEnd']];
    const windowStart = (wsRaw instanceof Date)
      ? String(wsRaw.getHours()).padStart(2, '0') + ':' + String(wsRaw.getMinutes()).padStart(2, '0')
      : String(wsRaw || '').trim();
    const windowEnd   = (weRaw instanceof Date)
      ? String(weRaw.getHours()).padStart(2, '0') + ':' + String(weRaw.getMinutes()).padStart(2, '0')
      : String(weRaw || '').trim();

    if (!windowStart || !windowEnd) {
      log('scheduleFollowupSends: campaign ' + campaignId + ' has no window — skipping', 'WARN');
      continue;
    }

    const startMinutes = timeStringToMinutes_(windowStart);
    const endMinutes   = timeStringToMinutes_(windowEnd);
    const windowSize   = endMinutes - startMinutes;

    if (windowSize <= 0) {
      log('scheduleFollowupSends: campaign ' + campaignId + ' invalid window — skipping', 'WARN');
      continue;
    }

    const count = group.length;

    const randomMinutes = [];
    for (let i = 0; i < count; i++) {
      randomMinutes.push(startMinutes + Math.floor(Math.random() * windowSize));
    }

    randomMinutes.sort(function(a, b) { return a - b; });

    for (let i = 1; i < randomMinutes.length; i++) {
      if (randomMinutes[i] < randomMinutes[i - 1] + MIN_GAP_MINUTES) {
        randomMinutes[i] = randomMinutes[i - 1] + MIN_GAP_MINUTES;
      }
    }

    if (randomMinutes[randomMinutes.length - 1] > endMinutes) {
      const interval = Math.max(MIN_GAP_MINUTES, Math.floor(windowSize / count));
      for (let i = 0; i < count; i++) {
        const jitter = Math.floor(Math.random() * 3) - 1;
        randomMinutes[i] = startMinutes + (i * interval) + jitter;
      }
    }

    for (let i = 0; i < group.length; i++) {
      allScheduledItems.push({
        dataIndex:       group[i].dataIndex,
        scheduledMinute: randomMinutes[i] || startMinutes
      });
    }
  }

  if (allScheduledItems.length === 0) {
    log('scheduleFollowupSends: no items to schedule', 'INFO');
    return;
  }

  // ── Sort and re-enforce gap across merged schedule ────────────────────────
  allScheduledItems.sort(function(a, b) { return a.scheduledMinute - b.scheduledMinute; });

  for (let i = 1; i < allScheduledItems.length; i++) {
    if (allScheduledItems[i].scheduledMinute < allScheduledItems[i - 1].scheduledMinute + MIN_GAP_MINUTES) {
      allScheduledItems[i].scheduledMinute = allScheduledItems[i - 1].scheduledMinute + MIN_GAP_MINUTES;
    }
  }

  // ── Write scheduledTime back to sheet ─────────────────────────────────────
  const campaignTimezone = data[Object.values(campaignGroups)[0][0].dataIndex][colMap['timezone']] || 'America/New_York';
const today = Utilities.formatDate(new Date(), campaignTimezone, 'yyyy-MM-dd');
  for (let i = 0; i < allScheduledItems.length; i++) {
    const item    = allScheduledItems[i];
    const hh      = String(Math.floor(item.scheduledMinute / 60)).padStart(2, '0');
    const mm      = String(item.scheduledMinute % 60).padStart(2, '0');
    const timeStr = today + 'T' + hh + ':' + mm + ':00';
    data[item.dataIndex][colMap['scheduledTime']] = "'" + timeStr;
    data[item.dataIndex][colMap['triggerId']]     = '';
  }

  inboxSheet.getRange(2, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();
  log('scheduleFollowupSends: wrote schedule for ' + allScheduledItems.length + ' rows', 'INFO');

  // ── Schedule first trigger ─────────────────────────────────────────────────
  const firstItem    = allScheduledItems[0];
  const firstTimeStr = data[firstItem.dataIndex][colMap['scheduledTime']];
  const firstRow     = data[firstItem.dataIndex];
  const weRawCheck   = firstRow[colMap['windowEnd']];
  const windowEndStr = (weRawCheck instanceof Date)
    ? String(weRawCheck.getHours()).padStart(2, '0') + ':' + String(weRawCheck.getMinutes()).padStart(2, '0')
    : String(weRawCheck || '').trim();

  // Check window-end using timeapi.io — simple string comparison
  if (windowEndStr) {
    try {
      const url      = 'https://timeapi.io/api/time/current/zone?timeZone=' + encodeURIComponent(campaignTimezone);
      const response = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
      const json     = JSON.parse(response.getContentText());
      const nowHHmm  = String(json.hour).padStart(2,'0') + ':' + String(json.minute).padStart(2,'0');
      if (nowHHmm > windowEndStr) {
        log('scheduleFollowupSends: sending window already closed (' + nowHHmm + ' > ' + windowEndStr + ') — not scheduling', 'INFO');
        return;
      }
    } catch (e) {
      log('scheduleFollowupSends: could not verify window end — ' + e.message + ' — proceeding anyway', 'WARN');
    }
  }

  const firstMs  = localTimeToUtcMs_(firstTimeStr, campaignTimezone);
  const nowMs    = new Date().getTime();
  const rawDelay = firstMs - nowMs;

  // Get current time in campaign timezone
  const fuTzNowStr   = Utilities.formatDate(new Date(), campaignTimezone, 'HH:mm');
  const fuTzNowParts = fuTzNowStr.split(':');
  const fuTzNowMin   = parseInt(fuTzNowParts[0], 10) * 60 + parseInt(fuTzNowParts[1], 10);

  // Get window boundaries from first row
  const fuWsRaw    = data[firstItem.dataIndex][colMap['windowStart']];
  const fuWeRaw    = data[firstItem.dataIndex][colMap['windowEnd']];
  const fuWinStart = (fuWsRaw instanceof Date)
    ? String(fuWsRaw.getHours()).padStart(2,'0') + ':' + String(fuWsRaw.getMinutes()).padStart(2,'0')
    : String(fuWsRaw || '').trim();
  const fuWinEnd   = (fuWeRaw instanceof Date)
    ? String(fuWeRaw.getHours()).padStart(2,'0') + ':' + String(fuWeRaw.getMinutes()).padStart(2,'0')
    : String(fuWeRaw || '').trim();
  const fuWinStartMin = timeStringToMinutes_(fuWinStart);
  const fuWinEndMin   = timeStringToMinutes_(fuWinEnd);

  let delayMs;
  if (fuTzNowMin >= fuWinStartMin && fuTzNowMin <= fuWinEndMin) {
    log('scheduleFollowupSends: currently inside window (' + fuTzNowStr + ') — firing in 1 min', 'INFO');
    delayMs = 60 * 1000;
  } else if (rawDelay <= 0) {
    log('scheduleFollowupSends: scheduled time already passed — firing in 1 min', 'WARN');
    delayMs = 60 * 1000;
  } else if (rawDelay > 23 * 60 * 60 * 1000) {
    if (fuTzNowMin <= fuWinEndMin) {
      log('scheduleFollowupSends: large delay but still in window — firing in 1 min', 'INFO');
      delayMs = 60 * 1000;
    } else {
      log('scheduleFollowupSends: window closed for today — waiting for tomorrow', 'INFO');
      delayMs = rawDelay;
    }
  } else {
    delayMs = rawDelay;
    log('scheduleFollowupSends: scheduling first send in ' + Math.round(delayMs/60000) + ' min', 'INFO');
  }

  const triggerId = createChainedFollowupTrigger_(delayMs);
  if (!triggerId) {
    log('scheduleFollowupSends: failed to create first trigger', 'ERROR');
    return;
  }

  const firstRowSheetIndex = firstItem.dataIndex + 2;
  inboxSheet.getRange(firstRowSheetIndex, colMap['triggerId'] + 1).setValue(triggerId);
  SpreadsheetApp.flush();

  log('scheduleFollowupSends: first trigger created, fires in ' + Math.round(delayMs / 60000) + ' min', 'INFO');
}

// ============================================================================
// SECTION 3 — processSingleFollowup()
// Called by each one-time trigger. Sends ONE follow-up then chains the next.
//
// Flow:
//  1. Identify own trigger ID → find matching row in followupInbox
//  2. Delete own trigger (cleanup)
//  3. Check sending window — if closed, reset row to Queued, stop chain
//  4. Check sentToday vs dailyLimit — stop chain if at limit
//  5. Acquire lock
//  6. Validate threadId present — skip row if missing
//  7. Set row status = 'Sending' (crash guard)
//  8. Look up template from cache
//  9. Send follow-up via thread.reply()
// 10. On success: mark Sent, update sentToday, notify Master Control
//     If final step: set sequenceStatus = Completed in update payload
// 11. On failure: increment retryCount, set Error or back to Queued
// 12. Find NEXT Queued row in followupInbox
// 13. If found: create next trigger (7–18 min random delay), write triggerId
// 14. If not found: log chain complete, exit
// ==================================================
/**
 * Sends one follow-up email and chains the next trigger.
 * Called by one-time triggers — never called directly.
 */
function processSingleFollowup() {
  const ss          = SpreadsheetApp.getActiveSpreadsheet();
  const configSheet = ss.getSheetByName(CONFIG_SHEET);
  const config      = getConfig(configSheet);

  // ── Pre-flight window check ───────────────────────────────────────────────
  try {
    const preFlightInbox = SpreadsheetApp.getActiveSpreadsheet()
                            .getSheetByName(FOLLOWUP_INBOX_SHEET);
    if (preFlightInbox && preFlightInbox.getLastRow() > 1) {
      const firstQueued = findFirstQueuedFollowupRow_(preFlightInbox);
      if (firstQueued) {
        const pfWs = firstQueued.row[firstQueued.colMap['windowStart']];
        const pfWe = firstQueued.row[firstQueued.colMap['windowEnd']];
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
          timezone: String(
            firstQueued.row[firstQueued.colMap['timezone']] || 
            'Africa/Cairo').trim()
        };

        if (!isInsideSendingWindow_(config, pfWindowSettings)) {
          log('processSingleFollowup: PRE-FLIGHT — outside window — ' +
              'orphan trigger — deleting ONLY myself and stopping', 'WARN');

          const myTriggerId = firstQueued.row[firstQueued.colMap['triggerId']];
          const allTriggers = ScriptApp.getProjectTriggers();
          allTriggers.forEach(function(t) {
            if (t.getHandlerFunction() === 'processSingleFollowup') {
              if (String(t.getUniqueId()) !== String(myTriggerId || '')) {
                ScriptApp.deleteTrigger(t);
                log('processSingleFollowup: deleted orphan trigger ' + 
                    t.getUniqueId(), 'WARN');
              }
            }
          });
          return;
        }
      }
    }
  } catch(e) {
    log('processSingleFollowup: pre-flight error — ' + e.message + 
        ' — continuing', 'WARN');
  }
  // ─────────────────────────────────────────────────────────────────────────

  // ── Identify which trigger just fired ────────────────────────────────────
  const inboxSheet = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);
  if (!inboxSheet) {
    log('processSingleFollowup: followupInbox sheet not found', 'ERROR');
    deleteOwnFollowupTrigger_(null);
    return;
  }

  const { rows, colMap } = readFollowupInbox_(inboxSheet);
  if (!rows) {
    log('processSingleFollowup: could not read followupInbox', 'ERROR');
    deleteOwnFollowupTrigger_(null);
    return;
  }

  // Find the row whose triggerId matches an active project trigger
  const myRow = findMyFollowupRow_(rows, colMap);
  if (!myRow) {
    log('processSingleFollowup: could not find matching row for this trigger — cleaning up', 'WARN');
    deleteOwnFollowupTrigger_(null);
    return;
  }

  // ── Delete own trigger immediately ────────────────────────────────────────
  deleteOwnFollowupTrigger_(myRow.triggerId);

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
    log('processSingleFollowup: outside sending window — resetting row to Queued, stopping chain', 'INFO');
    inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]    + 1).setValue('Queued');
    inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
    return; // no chain — window closed
  }

  // ── Daily limit check ─────────────────────────────────────────────────────
  const dailyLimit = parseInt(config.dailyLimit || '25', 10);
  const sentToday  = parseInt(config.sentToday  || '0',  10);
  if (sentToday >= dailyLimit) {
    log('processSingleFollowup: daily limit reached (' + sentToday + '/' + dailyLimit + ') — stopping chain', 'INFO');
    inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
    return; // no chain — limit hit
  }

  // ── Acquire lock ──────────────────────────────────────────────────────────
  const lock = acquireLock(SENDER_LOCK_WAIT);
  if (!lock) {
    log('processSingleFollowup: could not acquire lock — retrying in 1 min', 'WARN');
    const retryTriggerId = createChainedFollowupTrigger_(60 * 1000);
    if (retryTriggerId) {
      writeFollowupTriggerId_(inboxSheet, myRow.rowIndex, colMap, retryTriggerId);
    }
    return;
  }

  const updates = [];

  try {
    const leadId             = String(myRow.row[colMap[FOLLOWUP_COLS.leadId]]             || '').trim();
    const toEmail            = String(myRow.row[colMap[FOLLOWUP_COLS.email]]              || '').trim();
    const campaignId         = String(myRow.row[colMap[FOLLOWUP_COLS.campaignId]]         || '').trim();
    const sequenceStep       = parseInt(myRow.row[colMap[FOLLOWUP_COLS.sequenceStep]]     || 0,  10);
    const totalSequenceSteps = parseInt(myRow.row[colMap[FOLLOWUP_COLS.totalSequenceSteps]] || 0, 10);
    const threadId           = String(myRow.row[colMap[FOLLOWUP_COLS.threadId]]           || '').trim();
    const retryCount         = parseInt(myRow.row[colMap[FOLLOWUP_COLS.retryCount]]       || 0,  10);
    const leadDataRaw        = String(myRow.row[colMap[FOLLOWUP_COLS.leadData]]           || '{}');

    // ── Validate required fields ──────────────────────────────────────────
    if (!toEmail || !campaignId) {
      log('processSingleFollowup: row ' + myRow.rowIndex + ' missing email or campaignId — skipping', 'WARN');
      inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]    + 1).setValue('Error');
      inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
      // Still chain — one bad row shouldn't kill the rest
    } else if (!threadId) {
      // threadId is REQUIRED — a follow-up without one would create a new thread
      log('processSingleFollowup: row ' + myRow.rowIndex + ' missing threadId for ' + toEmail + ' — skipping', 'WARN');
      inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]    + 1).setValue('Error');
      inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
      // Still chain
    } else {
      // ── Build template cache ──────────────────────────────────────────
      const templateCache = getCampaignTemplates(config.masterSpreadsheetId);
      const cacheKey      = campaignId + '_' + sequenceStep;
      const template      = templateCache[cacheKey] || null;
      const nextTemplate  = templateCache[campaignId + '_' + (sequenceStep + 1)] || null;

      if (!template) {
        log('processSingleFollowup: no template for ' + cacheKey + ' — skipping lead ' + leadId, 'WARN');
        inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]    + 1).setValue('Error');
        inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
        // Still chain
      } else {
        let leadObj = {};
        try { leadObj = JSON.parse(leadDataRaw); } catch (e) { leadObj = {}; }

        const senderObj = {
          senderFirstName: config.senderFirstName || '',
          senderLastName:  config.senderLastName  || ''
        };

        // ── Crash guard: set Sending BEFORE Gmail call ────────────────
        inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]    + 1).setValue('Sending');
        inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
        SpreadsheetApp.flush();

        try {
          sendFollowupEmail(threadId, toEmail, template, leadObj, senderObj, config);

          // ── Success ──────────────────────────────────────────────────
          const nowStr         = formatDate(new Date());
          const isFinalStep    = (sequenceStep >= totalSequenceSteps);
          const awaitDays      = (nextTemplate && nextTemplate.awaitDays) ? nextTemplate.awaitDays : 3;
          const nextSentDate   = isFinalStep ? '' : addDays(formatDate(new Date()), awaitDays);

          inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]      + 1).setValue('Sent');
          inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.lastSentTime] + 1).setValue(nowStr);

          // Increment sentToday in Config
          updateConfigValue(configSheet, 'sentToday', String(sentToday + 1));
          updateConfigValue(configSheet, 'totalSentCount', String(parseInt(config.totalSentCount || 0, 10) + 1));
          updateConfigValue(configSheet, 'lastRunTime', new Date().toISOString());

          updates.push({
            leadId:         leadId,
            campaignId:     campaignId,
            status:         isFinalStep ? 'Completed' : 'Sent',
            lastSentTime:   nowStr,
            sequenceStep:   sequenceStep,
            nextSentDate:   nextSentDate,
            sequenceStatus: isFinalStep ? 'Completed' : ''
          });

          log('processSingleFollowup: sent step ' + sequenceStep + ' to ' + toEmail +
              (isFinalStep ? ' — FINAL STEP, sequence complete' : ''), 'INFO');

        } catch (e) {
          // ── Send failure ──────────────────────────────────────────────
          log('processSingleFollowup: Gmail error for lead ' + leadId + ' — ' + e.message, 'ERROR');
          const newRetry  = retryCount + 1;
          const newStatus = newRetry >= MAX_RETRY_COUNT ? 'Error' : 'Queued';
          inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.status]     + 1).setValue(newStatus);
          inboxSheet.getRange(myRow.rowIndex, colMap[FOLLOWUP_COLS.retryCount] + 1).setValue(newRetry);
        }
      }
    }

    // ── Notify Master Control ─────────────────────────────────────────────
    if (updates.length > 0 && config.masterWebhookUrl) {
      notifyMasterControl(config.masterWebhookUrl, updates, 'updateFollowupSent', config.senderId || '');
    }
    
    clearTerminalRows(ss.getSheetByName(FOLLOWUP_INBOX_SHEET), ['Sent']);

  } finally {
    releaseLock(lock);
  }

  // ── Chain: schedule next followup based on scheduledTime column ───────────
  const nextQueued = findFirstQueuedFollowupRow_(inboxSheet);
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
      log('processSingleFollowup: window closed — stopping chain, no more followups today', 'INFO');
    } else {
      const rawScheduled = nextQueued.row[nextQueued.colMap['scheduledTime']];
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
          log('processSingleFollowup: scheduled time already passed — using min gap + jitter', 'WARN');
          delayMs = MIN_GAP_MINUTES * 60 * 1000 + jitterMs;
        }
      } catch (e) {
        log('processSingleFollowup: error calculating delay — ' + e.message + ' — using min gap', 'WARN');
        delayMs = MIN_GAP_MINUTES * 60 * 1000;
      }
      if (delayMs > 30 * 60 * 1000) {
        log('processSingleFollowup: delayMs ' + Math.round(delayMs/60000) + ' min exceeds 30 min cap — clamping', 'WARN');
        delayMs = MIN_GAP_MINUTES * 60 * 1000 + Math.floor(Math.random() * 7 * 60 * 1000);
      }
      log('processSingleFollowup: delayMs calculated = ' + Math.round(delayMs/60000) + ' min', 'INFO');
      const nextTrigId = createChainedFollowupTrigger_(delayMs);
      if (nextTrigId) {
        writeFollowupTriggerId_(inboxSheet, nextQueued.rowIndex, nextQueued.colMap, nextTrigId);
        log('processSingleFollowup: next followup scheduled in ~' + Math.round(delayMs / 60000) + ' min, row=' + nextQueued.rowIndex, 'INFO');
      } else {
        log('processSingleFollowup: failed to create next trigger', 'ERROR');
      }
    }
  } else {
    log('processSingleFollowup: no more Queued rows — followup chain complete for today', 'INFO');
  }
}

// ============================================================================
// SECTION 4 — sendFollowupEmail()
// Sends one follow-up by replying to the original Gmail thread.
// Uses thread.reply() NOT replyAll() — system design decision.
// Throws on any failure so processSingleFollowup() can handle retries.
// ============================================================================
/**
 * Sends a follow-up email as a reply in the original Gmail thread.
 *
 * @param {string} threadId  - Gmail thread ID from the original cold send.
 * @param {string} toEmail   - Recipient email (used for logging only).
 * @param {Object} template  - { subject, htmlBody, plainBody } from cache.
 * @param {Object} leadObj   - Parsed lead data for variable replacement.
 * @param {Object} senderObj - { senderFirstName, senderLastName }
 * @param {Object} config    - Sender Config values.
 * @returns {void}
 * @throws {Error} On any Gmail API failure or invalid threadId.
 *
 * Notes:
 *  - thread.reply() preserves the original subject — subject variable
 *    replacement still runs here for logging consistency.
 *  - replyTo set to config.replyToEmail if present, else senderEmail.
 */
function sendFollowupEmail(threadId, toEmail, template, leadObj, senderObj, config) {
  // ── Step 1: Open thread and validate ─────────────────────────────────────
  const cleanThreadId = String(threadId).trim();
  const thread = GmailApp.getThreadById(cleanThreadId);
  if (!thread) {
    throw new Error('Thread not found for threadId: ' + cleanThreadId);
  }

  // ── Step 2: Get subject from thread ──────────────────────────────────────
  const originalSubject = thread.getFirstMessageSubject();
  const replySubject    = originalSubject.startsWith('Re: ')
    ? originalSubject
    : 'Re: ' + originalSubject;

  // ── Step 3: Replace variables in body ────────────────────────────────────
  const plainBody   = replaceVariables(template.plainBody, leadObj, senderObj);
  const senderName  = ((config.senderFirstName || '') + ' ' + (config.senderLastName || '')).trim();
  const senderEmail = config.senderEmail || '';

  // ── Step 4: Fetch headers from LAST message in thread ────────────────────
  // In-Reply-To = last message's Message-ID (the one we're directly replying to)
  // References  = existing References chain + last message's Message-ID
  const messages    = thread.getMessages();
  const lastMessage = messages[messages.length - 1];
  const lastGmailId = lastMessage.getId();

  let messageIdHeader  = '';
  let referencesHeader = '';

  try {
    const rawMessage = Gmail.Users.Messages.get('me', lastGmailId, { format: 'full' });
    if (rawMessage.payload && rawMessage.payload.headers) {
      for (const header of rawMessage.payload.headers) {
        const name = header.name.toLowerCase();
        if (name === 'message-id') { messageIdHeader  = header.value; }
        if (name === 'references') { referencesHeader = header.value; }
      }
    }
  } catch (e) {
    log('sendFollowupEmail: could not fetch RFC headers — ' + e.message + ' — using fallback', 'WARN');
  }

  if (!messageIdHeader) {
    messageIdHeader = '<' + lastGmailId + '@mail.gmail.com>';
    log('sendFollowupEmail: using fallback Message-ID ' + messageIdHeader, 'WARN');
  }

  const references = referencesHeader
    ? referencesHeader + ' ' + messageIdHeader
    : messageIdHeader;

  // ── Step 5: Build and send raw RFC 2822 plain text email ─────────────────
  const rawEmail = [
    'MIME-Version: 1.0',
    'From: ' + senderName + ' <' + senderEmail + '>',
    'To: ' + toEmail,
    'Subject: ' + replySubject,
    'In-Reply-To: ' + messageIdHeader,
    'References: ' + references,
    'Content-Type: text/plain; charset=UTF-8',
    '',
    plainBody
  ].join('\r\n');

  Gmail.Users.Messages.send({
    raw:      Utilities.base64EncodeWebSafe(rawEmail),
    threadId: cleanThreadId
  }, 'me');

  log('sendFollowupEmail: threaded reply sent to ' + toEmail + ' in thread ' + cleanThreadId, 'INFO');
}

// ============================================================================
// SECTION 5 — PRIVATE HELPERS (followup-specific)
// Mirrors the pattern of sender.gs private helpers but scoped to followupInbox
// and processSingleFollowup triggers so the two chains never collide.
// ============================================================================

/**
 * Reads all rows from followupInbox into memory.
 * Returns { rows, headers, colMap } or { rows: null } on failure.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inboxSheet
 * @returns {{ rows: Array[]|null, headers: string[], colMap: Object }}
 */
function readFollowupInbox_(inboxSheet) {
  try {
    const lastRow = inboxSheet.getLastRow();
    if (lastRow < 2) { return { rows: [], headers: [], colMap: {} }; }

    const headers = inboxSheet.getRange(1, 1, 1, inboxSheet.getLastColumn()).getValues()[0];
    const rows    = inboxSheet.getRange(2, 1, lastRow - 1, inboxSheet.getLastColumn()).getValues();
    const colMap  = {};
    headers.forEach(function(h, i) { colMap[String(h).trim()] = i; });
    return { rows: rows, headers: headers, colMap: colMap };
  } catch (e) {
    log('readFollowupInbox_: error — ' + e.message, 'ERROR');
    return { rows: null, headers: [], colMap: {} };
  }
}

/**
 * Finds the first row in followupInbox with status = 'Queued',
 * no triggerId set, and retryCount < MAX_RETRY_COUNT.
 * Returns { rowIndex, row, colMap } or null.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inboxSheet
 * @returns {{ rowIndex: number, row: Array, colMap: Object } | null}
 */
function findFirstQueuedFollowupRow_(inboxSheet) {
  const { rows, colMap } = readFollowupInbox_(inboxSheet);
  if (!rows || rows.length === 0) { return null; }

  for (let i = 0; i < rows.length; i++) {
    const status     = String(rows[i][colMap[FOLLOWUP_COLS.status]]    || '').trim();
    const triggerId  = String(rows[i][colMap[FOLLOWUP_COLS.triggerId]] || '').trim();
    const retryCount = parseInt(rows[i][colMap[FOLLOWUP_COLS.retryCount]] || 0, 10);

    if (status === 'Queued' && !triggerId && retryCount < MAX_RETRY_COUNT) {
      return { rowIndex: i + 2, row: rows[i], colMap: colMap };
    }
  }
  return null;
}

/**
 * Matches the currently-firing trigger to a row in followupInbox by triggerId.
 * Returns { rowIndex, row, colMap, triggerId } or null.
 *
 * @param {Array[]} rows
 * @param {Object}  colMap
 * @returns {{ rowIndex: number, row: Array, colMap: Object, triggerId: string } | null}
 */
function findMyFollowupRow_(rows, colMap) {
  if (!rows || rows.length === 0) { return null; }
  if (colMap[FOLLOWUP_COLS.triggerId] === undefined) { return null; }

  const activeTriggers = ScriptApp.getProjectTriggers();
  const activeTriggerIds = {};
  activeTriggers.forEach(function(t) { 
    activeTriggerIds[t.getUniqueId()] = true; 
  });

  const props = PropertiesService.getScriptProperties().getProperties();
  const pendingTriggerIds = {};
  Object.keys(props).forEach(function(key) {
    if (key.indexOf('pendingTrigger_') === 0) {
      const id = key.replace('pendingTrigger_', '');
      pendingTriggerIds[id] = true;
    }
  });

  for (let i = 0; i < rows.length; i++) {
    const rowTriggerId = String(rows[i][colMap[FOLLOWUP_COLS.triggerId]] || '').trim();
    if (!rowTriggerId) { continue; }
    if (activeTriggerIds[rowTriggerId] || pendingTriggerIds[rowTriggerId]) {
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
 * Creates a one-time trigger for processSingleFollowup() firing after delayMs.
 * Returns trigger unique ID or null on failure.
 *
 * @param {number} delayMs
 * @returns {string|null}
 */
function createChainedFollowupTrigger_(delayMs) {
  try {
    const safeDelay = Math.max(delayMs, 60 * 1000);
    const clampedDelay = Math.min(safeDelay, 23 * 60 * 60 * 1000);
    const trigger = ScriptApp.newTrigger('processSingleFollowup')
      .timeBased()
      .after(clampedDelay)
      .create();
    const id = trigger.getUniqueId();
    PropertiesService.getScriptProperties()
      .setProperty('pendingTrigger_' + id, 'processSingleFollowup');
    return id;
  } catch (e) {
    log('createChainedFollowupTrigger_: error — ' + e.message, 'ERROR');
    return null;
  }
}

/**
 * Deletes the processSingleFollowup trigger matching the given unique ID.
 * If triggerId is null, deletes all orphaned processSingleFollowup triggers.
 *
 * @param {string|null} triggerId
 */
function deleteOwnFollowupTrigger_(triggerId) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      const t = triggers[i];
      if (t.getHandlerFunction() !== 'processSingleFollowup') { continue; }
      if (triggerId === null) {
        log('deleteOwnFollowupTrigger_: called with null — skipping to avoid chain kill', 'WARN');
        return;
      }
      if (t.getUniqueId() === triggerId) {
        ScriptApp.deleteTrigger(t);
        return;
      }
    }
  } catch (e) {
    log('deleteOwnFollowupTrigger_: error — ' + e.message, 'WARN');
  }
}

/**
 * Writes a triggerId into the triggerId column of a followupInbox row.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} inboxSheet
 * @param {number} rowIndex  - 1-based sheet row.
 * @param {Object} colMap
 * @param {string} triggerId
 */
function writeFollowupTriggerId_(inboxSheet, rowIndex, colMap, triggerId) {
  if (colMap[FOLLOWUP_COLS.triggerId] === undefined) {
    log('writeFollowupTriggerId_: triggerId column not found in followupInbox — please add it', 'ERROR');
    return;
  }
  inboxSheet.getRange(rowIndex, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue(triggerId);
}

// ============================================================================
// SECTION 6 — setupFollowupTrigger()
// One-time setup. Creates the 8:00 AM daily trigger for scheduleFollowupSends().
// Also cleans up any leftover processSingleFollowup triggers from prior runs.
// Run manually once after deploying to a sender spreadsheet.
// ============================================================================
/**
 * Creates the daily 8:00 AM trigger for scheduleFollowupSends().
 * Safe to re-run — deletes existing triggers for both functions first.
 */
function setupFollowupTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    const fn = triggers[i].getHandlerFunction();
    if (fn === 'scheduleFollowupSends' || fn === 'processSingleFollowup') {
      ScriptApp.deleteTrigger(triggers[i]);
      log('setupFollowupTrigger: deleted trigger for ' + fn, 'INFO');
    }
  }

  // Fixed daily trigger at 8:00 AM — 30 min after cold scheduler at 7:30 AM
  ScriptApp.newTrigger('scheduleFollowupSends')
    .timeBased()
    .atHour(8)
    .nearMinute(45)
    .everyDays(1)
    .create();

  log('setupFollowupTrigger: created daily 8:00 AM trigger for scheduleFollowupSends', 'INFO');
  Logger.log('✅ Trigger created — scheduleFollowupSends will run daily at 8:00 AM.');
}

// ============================================================================
// SECTION 7 — setupFollowupInboxSheet()
// One-time setup. Creates the followupInbox sheet with correct headers.
// NOW INCLUDES triggerId column.
// ============================================================================
/**
 * Creates the followupInbox sheet with headers including triggerId.
 * Safe to re-run — skips if sheet already exists.
 */
function setupFollowupInboxSheet() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let sheet = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);
  if (sheet) {
    Logger.log('followupInbox sheet already exists — no changes made.\n\nIf upgrading an existing sheet, manually add a "triggerId" column header after "retryCount".');
    return;
  }
  sheet = ss.insertSheet(FOLLOWUP_INBOX_SHEET);
  const headers = [
    'leadId', 'email', 'campaignId', 'sequenceStep',
    'leadData', 'windowStart', 'windowEnd', 'sendingDays', 'timezone', 'status', 'sentTime', 'retryCount', 'triggerId', 'scheduledTime'
  ];// HARDCODED — verify this
  sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
  const headerRange = sheet.getRange(1, 1, 1, headers.length);
  headerRange.setBackground('#0f9d58'); // Green — distinguishes from dailyInbox (blue)
  headerRange.setFontColor('#ffffff');
  headerRange.setFontWeight('bold');
  sheet.setFrozenRows(1);
  log('setupFollowupInboxSheet: created with headers including triggerId', 'INFO');
  Logger.log('✅ followupInbox sheet created successfully.');
}

// ============================================================================
// SECTION 8 — cleanupOrphanFollowupTriggers()
// Safety utility. Run manually if the followup chain gets stuck.
// Mirrors cleanupOrphanTriggers() in sender.gs but targets followupInbox
// and processSingleFollowup triggers only.
// ============================================================================
/**
 * Emergency cleanup for the followup chain.
 * Deletes all processSingleFollowup triggers, clears triggerId column in
 * followupInbox, and resets any 'Sending' rows back to 'Queued'.
 * Run manually from the Apps Script editor if the chain gets stuck.
 */
function cleanupOrphanFollowupTriggers() {
  // Delete all processSingleFollowup triggers
  const triggers = ScriptApp.getProjectTriggers();
  let deleted = 0;
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'processSingleFollowup') {
      ScriptApp.deleteTrigger(triggers[i]);
      deleted++;
    }
  }

  // Clear triggerId column and reset Sending → Queued in followupInbox
  const ss         = SpreadsheetApp.getActiveSpreadsheet();
  const inboxSheet = ss.getSheetByName(FOLLOWUP_INBOX_SHEET);
  if (inboxSheet && inboxSheet.getLastRow() > 1) {
    const { rows, colMap } = readFollowupInbox_(inboxSheet);
    if (rows && rows.length > 0) {
      for (let i = 0; i < rows.length; i++) {
        const sheetRow = i + 2;
        if (colMap[FOLLOWUP_COLS.triggerId] !== undefined) {
          inboxSheet.getRange(sheetRow, colMap[FOLLOWUP_COLS.triggerId] + 1).setValue('');
        }
        const status = String(rows[i][colMap[FOLLOWUP_COLS.status]] || '').trim();
        if (status === 'Sending') {
          inboxSheet.getRange(sheetRow, colMap[FOLLOWUP_COLS.status] + 1).setValue('Queued');
        }
      }
    }
  }

  log('cleanupOrphanFollowupTriggers: deleted ' + deleted + ' triggers, cleared triggerId column', 'INFO');
  Logger.log('✅ Cleanup done. Deleted ' + deleted + ' processSingleFollowup triggers.\nSending rows reset to Queued. You can now re-run scheduleFollowupSends() manually.');
}