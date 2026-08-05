// ============================================================================
// distributor.gs — Cold Email Automation System
// Location: Master Control Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Runs twice daily via time-based triggers:
//   7:00 AM → distributeCold()
//   7:30 AM → distributeFollowups()
//
// Reads the Leads DB, calculates sender capacity, writes lead queues into
// each sender's dailyInbox or followupInbox via webhook POST.
//
// Dependencies: utils.gs must be present in the same Apps Script project.
// ============================================================================


// ============================================================================
// SECTION 1 — CONSTANTS
// All IDs and sheet names used in this file.
// ============================================================================

// The Google Sheets ID of the Master Control spreadsheet.
// Found in the sheet URL: docs.google.com/spreadsheets/d/THIS_PART/edit
const MASTER_SHEET_ID = '1TZ9oJd4m3OAHbM-_ZftI20p9GYAJKITxJPVRnbPOC3s'; // HARDCODED — verify this

// Sheet names inside Master Control
const SHEET_SENDER_ACCOUNTS = 'senderAccounts'; // HARDCODED — verify this
const SHEET_CAMPAIGNS       = 'Campaigns';       // HARDCODED — verify this
const SHEET_SETTINGS        = 'Settings';        // HARDCODED — verify this

// The Google Sheets ID of the Leads Database spreadsheet.
// This is also stored in Settings under 'leadsSpreadsheetId' —
// we read it from Settings at runtime so it can be changed without redeploying.
const LEADS_SHEET_NAME = 'masterLeads'; // HARDCODED — verify this

// Webhook actions sent to each sender's webApp
const ACTION_WRITE_COLD_INBOX    = 'writeColdInbox';    // HARDCODED — verify this
const ACTION_WRITE_FOLLOWUP_INBOX = 'writeFollowupInbox'; // HARDCODED — verify this

// Lock wait time in seconds
const LOCK_WAIT_SECONDS = 10; // HARDCODED — verify this

// Max leads to check per distribution run (safety cap for 6-min execution limit)
const MAX_LEADS_TO_SCAN = 60000; // HARDCODED — verify this

// Webhook batch size — never POST one lead at a time
const WEBHOOK_BATCH_SIZE = 50; // HARDCODED — verify this


// ============================================================================
// SECTION 2 — distributeCold()
// Main function for cold email distribution. Triggered at 7:00 AM.
//
// Flow:
//  1. Acquire lock
//  2. Read settings, active senders, eligible cold leads
//  3. Assign leads to senders by capacity
//  4. Write to each sender's dailyInbox via webhook
//  5. Batch update masterLeads (status=Queued, ownerSender)
//  6. Release lock
// ============================================================================

/**
 * Distributes cold (initial) email leads to sender dailyInbox sheets.
 * Triggered at 7:00 AM daily from Master Control spreadsheet.
 * Reads leads where: status=Pending, sequenceStep=0, campaignStatus=Active,
 * replyStatus is empty.
 */
function distributeCold() {
  log('distributeCold: started', 'INFO');

  startExecutionTimer_();

  // Check if this is a continuation run.
  // Resume by campaignId, NOT by index: leads queued in the previous pass are no
  // longer Pending, so fully-processed campaigns drop out of the eligible set and
  // every positional index shifts down. Resuming on a stale index skips campaigns.
  const coldCheckpoint     = loadCheckpoint_('distributeCold');
  const coldResumeCampaign = coldCheckpoint ? String(coldCheckpoint.campaignId || '') : '';
  if (coldResumeCampaign) {
    log('distributeCold: resuming from campaign ' + coldResumeCampaign, 'INFO');
  }

  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    log('distributeCold: could not acquire lock — exiting', 'WARN');
    return;
  }

  try {
    // --- Step 1: Read global settings + campaign settings ---
    const settings        = getSettings(MASTER_SHEET_ID);
    const leadsSheetId    = settings.leadsSpreadsheetId;
    const flexibleAllocation  = String(settings.flexibleAllocation || '').trim().toUpperCase() === 'TRUE';
    const distributionMethod  = settings.distributionMethod || 'roundRobin';
    const globalColdLimit     = parseFloat(settings.coldPriorityLimit)     || 0.75;
    const globalFollowupLimit = parseFloat(settings.followupPriorityLimit) || 0.25;
    const campaignSettings    = getCampaignSettings(MASTER_SHEET_ID);

    if (!leadsSheetId) {
      log('distributeCold: leadsSpreadsheetId not found in Settings — exiting', 'ERROR');
      return;
    }

    // --- Step 2: Read ALL active senders once ---
    const allSenders = getActiveSenders();
    if (allSenders.length === 0) {
      log('distributeCold: no active senders found — exiting', 'INFO');
      return;
    }

    // Build O(1) lookup map — avoids find() looping through 140 senders every time
    const allSendersMap = {};
    allSenders.forEach(function(s) { allSendersMap[s.emailID] = s; });
    // Read real queued counts from all senders in parallel
    const statusRequests = [];
    const statusMeta     = [];

    for (let i = 0; i < allSenders.length; i++) {
      const s = allSenders[i];
      if (!s.webAppUrl) { continue; }
      statusRequests.push({
        url:                s.webAppUrl,
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({ action: 'getSenderStatus' }),
        muteHttpExceptions: true
      });
      statusMeta.push({ emailID: s.emailID });
    }

    for (let b = 0; b < statusRequests.length; b += 50) {
      const batch     = statusRequests.slice(b, b + 50);
      const batchMeta = statusMeta.slice(b, b + 50);
      let   responses = [];
      try { responses = UrlFetchApp.fetchAll(batch); } catch(e) {
        log('distributeCold: getSenderStatus fetchAll error — ' + e.message, 'WARN');
        continue;
      }
      for (let r = 0; r < responses.length; r++) {
        if (responses[r].getResponseCode() !== 200) { continue; }
        try {
          const status = JSON.parse(responses[r].getContentText());
          const sender = allSendersMap[batchMeta[r].emailID];
          if (sender) {
            sender.existingColdQueued     = status.coldQueued     || 0;
            sender.existingFollowupQueued = status.followupQueued || 0;
          }
        } catch(e) { continue; }
      }
    }
    log('distributeCold: sender status loaded for capacity checks', 'INFO');

    // --- Step 3: Read eligible cold leads ---
    const { leads, headers } = getEligibleColdLeads(leadsSheetId);
    if (leads.length === 0) {
      log('distributeCold: no eligible cold leads found — exiting', 'INFO');
      return;
    }
    log('distributeCold: found ' + leads.length + ' eligible cold leads', 'INFO');

    // --- Step 4: Group leads by campaignId ---
    const leadsByCampaign = {};
    for (let i = 0; i < leads.length; i++) {
      const cId = leads[i].campaignId;
      if (!leadsByCampaign[cId]) { leadsByCampaign[cId] = []; }
      leadsByCampaign[cId].push(leads[i]);
    }

    const dbUpdates        = [];
    const allFetchRequests = [];
    const allFetchMeta     = [];

    // --- Step 5: Distribute per campaign ---
    const campaignIds = Object.keys(leadsByCampaign).sort();

    // Map the checkpointed campaignId back onto the freshly-read campaign list
    let coldStartIndex = 0;
    if (coldResumeCampaign) {
      const coldResumeAt = campaignIds.indexOf(coldResumeCampaign);
      coldStartIndex = coldResumeAt >= 0 ? coldResumeAt : 0;
      log('distributeCold: resume campaign ' + coldResumeCampaign +
          ' resolved to index ' + coldStartIndex + ' of ' + campaignIds.length, 'INFO');
    }

      for (let c = coldStartIndex; c < campaignIds.length; c++) {
      const campaignId    = campaignIds[c];

      // Time limit check
      if (isTimeLimitApproaching_()) {
        log('distributeCold: time limit approaching at campaign index ' + c + ' — saving checkpoint', 'WARN');

        // Fire whatever requests we built so far before stopping
        for (let b = 0; b < allFetchRequests.length; b += 50) {
          const bReqs = allFetchRequests.slice(b, b + 50);
          const bMeta = allFetchMeta.slice(b, b + 50);
          let responses = [];
          try { responses = UrlFetchApp.fetchAll(bReqs); } catch(e) { continue; }
          for (let r = 0; r < responses.length; r++) {
            if (responses[r].getResponseCode() === 200) {
              bMeta[r].assignedLeads.forEach(function(lead) {
                dbUpdates.push({
                  rowIndex: lead.rowIndex,
                  fields: {
                    status:             'Queued',
                    ownerSender:        bMeta[r].senderId,
                    sequenceStep:       1,
                    totalSequenceSteps: bMeta[r].totalSequenceSteps || 0
                  }
                });
              });
            }
          }
        }

        // These requests have now been fired. Drop them so that if the checkpoint
        // save fails below and execution continues, the final fan-out cannot POST
        // them a second time and write duplicate rows into sender inboxes.
        allFetchRequests.length = 0;
        allFetchMeta.length     = 0;

        // Write DB updates so far
        if (dbUpdates.length > 0) {
          batchUpdateLeadsDB(leadsSheetId, dbUpdates);
          log('distributeCold: partial DB update — wrote ' + dbUpdates.length + ' leads before checkpoint', 'INFO');
        }

        // Save only small checkpoint — campaignId is what we resume on
        const saved = saveCheckpointAndChain_('distributeCold', {
          campaignIndex: c,
          campaignId:    campaignId
        });
        if (saved) { return; }
        log('distributeCold: checkpoint save failed — continuing anyway', 'WARN');
      }
      const campaignLeads = leadsByCampaign[campaignId];
      const campSettings  = campaignSettings[campaignId];

      if (!campSettings) {
        log('distributeCold: no settings found for campaign ' + campaignId + ' — skipping', 'WARN');
        continue;
      }

      // The campaign's OWN status is authoritative. handleUploadLeads stamps every
      // new lead with campaignStatus='Active' no matter what the campaign says, so
      // without this a Paused or Deleted campaign would still send.
      // Blank is treated as Active so existing campaigns keep working.
      const campStatus = String(campSettings.campaignStatus || '').trim();
      if (campStatus && campStatus !== 'Active') {
        log('distributeCold: campaign ' + campaignId + ' is ' + campStatus + ' — skipping', 'INFO');
        continue;
      }
      // Skip distribution on non-sending days
      if (campSettings.sendingDays) {
        const dayNames    = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
        const tz          = campSettings.timezone || 'Africa/Cairo';
        const todayStr    = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
        const todayDay    = dayNames[new Date(todayStr + 'T12:00:00').getDay()];
        const allowedDays = campSettings.sendingDays.split(',').map(function(d) { return d.trim(); });
        if (allowedDays.indexOf(todayDay) === -1) {
          log('distributeCold: today (' + todayDay + ') not a sending day for campaign ' + campaignId + ' — skipping', 'INFO');
          continue;
        }
      }

      // Filter senders to only those assigned to this campaign
      const assignedIds = campSettings.assignedSenders; // e.g. ['S001','S002']
      if (!assignedIds || assignedIds.length === 0) {
        log('distributeCold: campaign ' + campaignId + ' has no assignedSenders — skipping', 'WARN');
        continue;
      }

      const campaignSenders = assignedIds
        .map(function(id) { return allSendersMap[id]; })
        .filter(Boolean);

      if (campaignSenders.length === 0) {
        log('distributeCold: no active senders match assignedSenders for campaign ' + campaignId, 'WARN');
        continue;
      }

      // Use campaign coldPriorityLimit or fall back to global
      const coldPriorityLimit     = campSettings.coldPriorityLimit     || globalColdLimit;
      const followupPriorityLimit = campSettings.followupPriorityLimit || globalFollowupLimit;

      // Calculate cold slots per sender for this campaign
      const sendersWithSlots = campaignSenders.map(function(sender) {
      const rawColdSlots       = Math.floor(sender.dailyLimit * coldPriorityLimit);
      const rawFollowupSlots   = Math.floor(sender.dailyLimit * followupPriorityLimit);
      const existingColdQueued = sender.existingColdQueued || 0;
      const usedSlots          = sender.usedColdSlots || 0;

      // Calculate wasted slots and give to priority holder
      const totalRaw   = rawColdSlots + rawFollowupSlots;
      const wasted     = sender.dailyLimit - totalRaw;
      const coldIsHigherPriority = coldPriorityLimit >= followupPriorityLimit;

      let finalColdSlots = rawColdSlots;
      if (wasted > 0) {
        if (coldIsHigherPriority) {
          finalColdSlots = rawColdSlots + wasted;
        }
        // if followup is priority wasted goes to followup (handled in distributeFollowups)
      }

      // flexibleAllocation bonus
      let bonusColdSlots = 0;
      if (flexibleAllocation) {
        const followupCeiling        = Math.floor(sender.dailyLimit * globalFollowupLimit);
        const existingFollowupQueued = sender.existingFollowupQueued || 0;
        bonusColdSlots               = Math.max(0, followupCeiling - existingFollowupQueued);
      }

      // Hard ceiling — never hand out more than the sender can still send today.
      // dailyLimit covers cold + followups together.
      const roomLeft = Math.max(0, sender.availableCapacity
                                   - existingColdQueued
                                   - (sender.existingFollowupQueued || 0)
                                   - usedSlots);
      const wantSlots      = Math.max(0, finalColdSlots + bonusColdSlots - existingColdQueued - usedSlots);
      const availableSlots = Math.min(roomLeft, wantSlots);

      log('distributeCold: sender=' + sender.emailID +
          ' dailyLimit=' + sender.dailyLimit +
          ' rawCold=' + rawColdSlots +
          ' wasted=' + wasted +
          ' coldPriority=' + coldIsHigherPriority +
          ' finalCold=' + finalColdSlots +
          ' bonus=' + bonusColdSlots +
          ' available=' + availableSlots, 'INFO');

      return Object.assign({}, sender, { coldSlots: availableSlots });
    });

      // Assign leads to senders
      const assignments = assignLeadsToSenders(campaignLeads, sendersWithSlots, 'coldSlots', distributionMethod);
      if (assignments.size === 0) {
        log('distributeCold: no capacity for campaign ' + campaignId, 'INFO');
        continue;
      }

      // Write to sender inboxes — embed campaign window settings in each row
      // Build all requests for this campaign first — fire them all at once

      assignments.forEach(function(assignedLeads, senderId) {
        const sender = sendersWithSlots.find(function(s) { return s.emailID === senderId; });
        if (!sender) { return; }

        const inboxRows = assignedLeads.map(function(lead) {
          return {
            leadId:       lead.leadId,
            email:        lead.email,
            campaignId:   lead.campaignId,
            sequenceStep: 1,
            leadData:     buildLeadDataJson(lead.rawRow, headers),
            status:       'Queued',
            sentTime:     '',
            retryCount:   0,
            windowStart:  campSettings.windowStart  || '',
            windowEnd:    campSettings.windowEnd    || '',
            sendingDays:  campSettings.sendingDays  || '',
            timezone:     campSettings.timezone     || ''
          };
        });

        allFetchRequests.push({
          url:                sender.webAppUrl,
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({ action: ACTION_WRITE_COLD_INBOX, rows: inboxRows }),
          muteHttpExceptions: true
        });

        allFetchMeta.push({
          senderId:           senderId,
          assignedLeads:      assignedLeads,
          inboxRows:          inboxRows,
          totalSequenceSteps: campSettings.totalSequenceSteps || 0
        });
      });
    }

    // Fire ALL requests across ALL campaigns in one go
    for (let b = 0; b < allFetchRequests.length; b += 50) {
      const batchRequests = allFetchRequests.slice(b, b + 50);
      const batchMeta     = allFetchMeta.slice(b, b + 50);

      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(batchRequests);
      } catch (e) {
        log('distributeCold: fetchAll error — ' + e.message, 'ERROR');
        continue;
      }

      for (let r = 0; r < responses.length; r++) {
        const meta = batchMeta[r];
        const code = responses[r].getResponseCode();

        if (code === 200) {
          meta.assignedLeads.forEach(function(lead) {
            dbUpdates.push({
              rowIndex: lead.rowIndex,
              fields: {
                status:             'Queued',
                ownerSender:        meta.senderId,
                sequenceStep:       1,
                totalSequenceSteps: meta.totalSequenceSteps || 0
              }
            });
          });
          log('distributeCold: wrote ' + meta.inboxRows.length + ' rows to sender ' + meta.senderId, 'INFO');
        } else {
          log('distributeCold: webhook failed (' + code + ') for sender ' + meta.senderId, 'ERROR');
        }
      }
    }

    // --- Step 6: Batch update masterLeads ---
    if (dbUpdates.length > 0) {
      batchUpdateLeadsDB(leadsSheetId, dbUpdates);
      log('distributeCold: updated ' + dbUpdates.length + ' leads in DB', 'INFO');
    }

    log('distributeCold: completed successfully', 'INFO');

    // Kick off sending chain on all assigned senders.
    // Each sender runs scheduleDailySends() synchronously inside this POST, so at
    // 200-300 senders this fan-out costs 60-120s+ on its own. There is no
    // checkpoint past this point — if we run out of budget here the execution is
    // killed and the remaining batches never get kicked off at all. Defer instead.
    if (isTimeLimitApproaching_()) {
      log('distributeCold: no execution budget left for kickoff — deferring to a one-time trigger', 'WARN');
      scheduleDeferredKickoff_('kickoffColdChainsDeferred');
    } else {
      log('distributeCold: kicking off send chains on all senders', 'INFO');
      fanOutToSenders_('kickoffCold', 'distributeCold');
    }

  } catch (e) {
    log('distributeCold: unexpected error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}


// NEW ASIM

function backfillTotalSequenceSteps() {
  const settings     = getSettings(MASTER_SHEET_ID);
  const leadsSheetId = settings.leadsSpreadsheetId;

  // Read campaigns into a map: campaignId → totalSequenceSteps
  const masterSS      = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const campaignSheet = masterSS.getSheetByName('Campaigns');
  const campHeaders   = campaignSheet.getRange(1, 1, 1, campaignSheet.getLastColumn()).getValues()[0];
  const campData      = campaignSheet.getRange(2, 1, campaignSheet.getLastRow() - 1, campaignSheet.getLastColumn()).getValues();
  const campCol       = {};
  campHeaders.forEach(function(h, i) { campCol[String(h).trim()] = i; });

  const campaignMap = {};
  for (let i = 0; i < campData.length; i++) {
    const cId    = String(campData[i][campCol['campaignId']]        || '').trim();
    const cSteps = parseInt(campData[i][campCol['totalSequenceSteps']] || 0, 10);
    if (cId && cSteps) { campaignMap[cId] = cSteps; }
  }

  // Read masterLeads and fill totalSequenceSteps
  const leadsSS    = SpreadsheetApp.openById(leadsSheetId);
  const leadsSheet = leadsSS.getSheetByName('masterLeads');
  const headers    = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
  const data       = leadsSheet.getRange(2, 1, leadsSheet.getLastRow() - 1, leadsSheet.getLastColumn()).getValues();
  const col        = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  for (let i = 0; i < data.length; i++) {
    const cId = String(data[i][col['campaignId']] || '').trim();
    if (cId && campaignMap[cId]) {
      data[i][col['totalSequenceSteps']] = campaignMap[cId];
    }
  }

  leadsSheet.getRange(2, 1, data.length, headers.length).setValues(data);
  SpreadsheetApp.flush();
  log('backfillTotalSequenceSteps: done', 'INFO');
}


// ============================================================================
// SECTION 3 — distributeFollowups()
// Main function for follow-up distribution. Triggered at 7:30 AM.
//
// Eligible leads: status=Sent, sequenceStep>0, sequenceStep<totalSequenceSteps,
//                 campaignStatus=Active, replyStatus=empty, nextSentDate<=today.
//
// flexibleAllocation: if TRUE, unused cold slots are added to followup slots.
// ============================================================================

/**
 * Distributes follow-up email leads to sender followupInbox sheets.
 * Triggered at 7:30 AM daily from Master Control spreadsheet.
 */
function distributeFollowups() {
  log('distributeFollowups: started', 'INFO');

  startExecutionTimer_();

  // Resume by campaignId, NOT by index — see the note in distributeCold().
  const followupCheckpoint     = loadCheckpoint_('distributeFollowups');
  const followupResumeCampaign = followupCheckpoint ? String(followupCheckpoint.campaignId || '') : '';
  if (followupResumeCampaign) {
    log('distributeFollowups: resuming from campaign ' + followupResumeCampaign, 'INFO');
  }

  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    log('distributeFollowups: could not acquire lock — exiting', 'WARN');
    return;
  }

  try {
    // --- Step 1: Read global settings + campaign settings ---
    const settings              = getSettings(MASTER_SHEET_ID);
    const leadsSheetId          = settings.leadsSpreadsheetId;
    const flexibleAllocation    = String(settings.flexibleAllocation || '').trim().toUpperCase() === 'TRUE';
    const globalFollowupLimit   = parseFloat(settings.followupPriorityLimit) || 0.25;
    const globalColdLimit       = parseFloat(settings.coldPriorityLimit)     || 0.75;
    const campaignSettings      = getCampaignSettings(MASTER_SHEET_ID); // NEW

    if (!leadsSheetId) {
      log('distributeFollowups: leadsSpreadsheetId not found in Settings — exiting', 'ERROR');
      return;
    }

    // --- Step 2: Read ALL active senders once ---
    const allSenders = getActiveSenders();
    if (allSenders.length === 0) {
      log('distributeFollowups: no active senders found — exiting', 'INFO');
      return;
    }

    // Build O(1) lookup map
    const allSendersMap = {};
    allSenders.forEach(function(s) { allSendersMap[s.emailID] = s; });
    // Read real queued counts from all senders in parallel
    const statusRequests = [];
    const statusMeta     = [];

    for (let i = 0; i < allSenders.length; i++) {
      const s = allSenders[i];
      if (!s.webAppUrl) { continue; }
      statusRequests.push({
        url:                s.webAppUrl,
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({ action: 'getSenderStatus' }),
        muteHttpExceptions: true
      });
      statusMeta.push({ emailID: s.emailID });
    }

    for (let b = 0; b < statusRequests.length; b += 50) {
      const batch     = statusRequests.slice(b, b + 50);
      const batchMeta = statusMeta.slice(b, b + 50);
      let   responses = [];
      try { responses = UrlFetchApp.fetchAll(batch); } catch(e) {
        log('distributeFollowups: getSenderStatus fetchAll error — ' + e.message, 'WARN');
        continue;
      }
      for (let r = 0; r < responses.length; r++) {
        if (responses[r].getResponseCode() !== 200) { continue; }
        try {
          const status = JSON.parse(responses[r].getContentText());
          const sender = allSendersMap[batchMeta[r].emailID];
          if (sender) {
            sender.existingColdQueued     = status.coldQueued     || 0;
            sender.existingFollowupQueued = status.followupQueued || 0;
          }
        } catch(e) { continue; }
      }
    }
    log('distributeFollowups: sender status loaded for capacity checks', 'INFO');

    // --- Step 3: Read eligible follow-up leads ---
    const { leads, headers } = getEligibleFollowupLeads(leadsSheetId);
    if (leads.length === 0) {
      log('distributeFollowups: no eligible follow-up leads found — exiting', 'INFO');
      return;
    }
    log('distributeFollowups: found ' + leads.length + ' eligible follow-up leads', 'INFO');

    // --- Step 4: Group leads by campaignId ---
    const leadsByCampaign = {};
    for (let i = 0; i < leads.length; i++) {
      const cId = leads[i].campaignId;
      if (!leadsByCampaign[cId]) { leadsByCampaign[cId] = []; }
      leadsByCampaign[cId].push(leads[i]);
    }

    const dbUpdates        = [];
    const allFetchRequests = [];
    const allFetchMeta     = [];

    // --- Step 5: Distribute per campaign ---
    const campaignIds = Object.keys(leadsByCampaign).sort();

    // Map the checkpointed campaignId back onto the freshly-read campaign list
    let followupStartIndex = 0;
    if (followupResumeCampaign) {
      const followupResumeAt = campaignIds.indexOf(followupResumeCampaign);
      followupStartIndex = followupResumeAt >= 0 ? followupResumeAt : 0;
      log('distributeFollowups: resume campaign ' + followupResumeCampaign +
          ' resolved to index ' + followupStartIndex + ' of ' + campaignIds.length, 'INFO');
    }

      for (let c = followupStartIndex; c < campaignIds.length; c++) {
      const campaignId    = campaignIds[c];

      // Time limit check
      if (isTimeLimitApproaching_()) {
        log('distributeFollowups: time limit approaching at campaign index ' + c + ' — saving checkpoint', 'WARN');

        // Fire whatever requests built so far before stopping
        for (let b = 0; b < allFetchRequests.length; b += 50) {
          const bReqs = allFetchRequests.slice(b, b + 50);
          const bMeta = allFetchMeta.slice(b, b + 50);
          let responses = [];
          try { responses = UrlFetchApp.fetchAll(bReqs); } catch(e) { continue; }
          for (let r = 0; r < responses.length; r++) {
            if (responses[r].getResponseCode() === 200) {
              bMeta[r].assignedLeads.forEach(function(lead) {
                dbUpdates.push({
                  rowIndex: lead.rowIndex,
                  fields: {
                    status:       'Queued',
                    sequenceStep: lead.sequenceStep + 1
                  }
                });
              });
            }
          }
        }

        // Already fired — drop them so a failed checkpoint save cannot cause the
        // final fan-out to POST the same rows twice.
        allFetchRequests.length = 0;
        allFetchMeta.length     = 0;

        // Write DB updates so far
        if (dbUpdates.length > 0) {
          batchUpdateLeadsDB(leadsSheetId, dbUpdates);
          log('distributeFollowups: partial DB update — wrote ' + dbUpdates.length + ' leads before checkpoint', 'INFO');
        }

        // Save only small checkpoint — campaignId is what we resume on
        const saved = saveCheckpointAndChain_('distributeFollowups', {
          campaignIndex: c,
          campaignId:    campaignId
        });
        if (saved) { return; }
        log('distributeFollowups: checkpoint save failed — continuing anyway', 'WARN');
      }
      const campaignLeads = leadsByCampaign[campaignId];
      const campSettings  = campaignSettings[campaignId];

      if (!campSettings) {
        log('distributeFollowups: no settings for campaign ' + campaignId + ' — skipping', 'WARN');
        continue;
      }
      // Same authoritative campaign-status gate as distributeCold — see note there.
      const campStatus = String(campSettings.campaignStatus || '').trim();
      if (campStatus && campStatus !== 'Active') {
        log('distributeFollowups: campaign ' + campaignId + ' is ' + campStatus + ' — skipping', 'INFO');
        continue;
      }
      // Skip distribution on non-sending days
      if (campSettings.sendingDays) {
        const dayNames    = ['Sun','Mon','Tue','Wed','Thu','Fri','Sat'];
        const tz          = campSettings.timezone || 'Africa/Cairo';
        const todayStr    = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
        const todayDay    = dayNames[new Date(todayStr + 'T12:00:00').getDay()];
        const allowedDays = campSettings.sendingDays.split(',').map(function(d) { return d.trim(); });
        if (allowedDays.indexOf(todayDay) === -1) {
          log('distributeFollowups: today (' + todayDay + ') not a sending day for campaign ' + campaignId + ' — skipping', 'INFO');
          continue;
        }
      }

      const assignedIds = campSettings.assignedSenders;
      if (!assignedIds || assignedIds.length === 0) {
        log('distributeFollowups: campaign ' + campaignId + ' has no assignedSenders — skipping', 'WARN');
        continue;
      }

      // O(1) map lookup instead of filter + indexOf across all 140 senders
      const campaignSenders = assignedIds
        .map(function(id) { return allSendersMap[id]; })
        .filter(Boolean); // remove any undefined (sender not active or not found)

      if (campaignSenders.length === 0) {
        log('distributeFollowups: no active senders match for campaign ' + campaignId, 'WARN');
        continue;
      }

      // Use campaign limits or fall back to global
      const followupPriorityLimit = campSettings.followupPriorityLimit || globalFollowupLimit;
      const coldPriorityLimit     = campSettings.coldPriorityLimit     || globalColdLimit;

      const sendersWithSlots = campaignSenders.map(function(sender) {
      const rawColdSlots      = Math.floor(sender.dailyLimit * coldPriorityLimit);
      const rawFollowupSlots  = Math.floor(sender.dailyLimit * followupPriorityLimit);
      const existingQueued    = sender.existingFollowupQueued || 0;

      // Wasted slots go to followup if followup has higher priority
      const totalRaw              = rawColdSlots + rawFollowupSlots;
      const wasted                = sender.dailyLimit - totalRaw;
      const followupIsHigherPriority = followupPriorityLimit > coldPriorityLimit;

      let finalFollowupSlots = rawFollowupSlots;
      if (wasted > 0 && followupIsHigherPriority) {
        finalFollowupSlots = rawFollowupSlots + wasted;
      }

      // flexibleAllocation bonus
      let bonusSlots = 0;
      if (flexibleAllocation) {
        const coldCeiling   = Math.floor(sender.dailyLimit * coldPriorityLimit);
        const usedColdSlots = sender.existingColdQueued || 0;
        bonusSlots          = Math.max(0, coldCeiling - usedColdSlots);
      }

      const roomLeft = Math.max(0, sender.availableCapacity
                                   - (sender.existingColdQueued || 0)
                                   - existingQueued);
      const wantSlots     = Math.max(0, finalFollowupSlots + bonusSlots - existingQueued);
      const followupSlots = Math.min(roomLeft, wantSlots);

      log('distributeFollowups: sender=' + sender.emailID +
          ' dailyLimit=' + sender.dailyLimit +
          ' rawFollowup=' + rawFollowupSlots +
          ' wasted=' + wasted +
          ' followupPriority=' + followupIsHigherPriority +
          ' finalFollowup=' + finalFollowupSlots +
          ' bonus=' + bonusSlots +
          ' total=' + followupSlots, 'INFO');

      return Object.assign({}, sender, { followupSlots: followupSlots });
    });

      const assignments = assignLeadsToSenders(campaignLeads, sendersWithSlots, 'followupSlots');
      if (assignments.size === 0) {
        log('distributeFollowups: no followup capacity for campaign ' + campaignId, 'INFO');
        continue;
      }

      // Build all requests for this campaign first — fire them all at once

      assignments.forEach(function(assignedLeads, senderId) {
        const sender = sendersWithSlots.find(function(s) { return s.emailID === senderId; });
        if (!sender) { return; }

        const inboxRows = assignedLeads.map(function(lead) {
          return {
            leadId:             lead.leadId,
            email:              lead.email,
            campaignId:         lead.campaignId,
            sequenceStep:       lead.sequenceStep + 1,
            totalSequenceSteps: lead.totalSequenceSteps,
            leadData:           buildLeadDataJson(lead.rawRow, headers),
            threadId:           lead.threadId,
            messageId:          lead.messageId,
            status:             'Queued',
            lastSentTime:       '',
            retryCount:         0,
            windowStart:        campSettings.windowStart || '',
            windowEnd:          campSettings.windowEnd   || '',
            sendingDays:        campSettings.sendingDays || '',
            timezone:           campSettings.timezone    || ''
          };
        });

        allFetchRequests.push({
          url:                sender.webAppUrl,
          method:             'post',
          contentType:        'application/json',
          payload:            JSON.stringify({ action: ACTION_WRITE_FOLLOWUP_INBOX, rows: inboxRows }),
          muteHttpExceptions: true
        });

        allFetchMeta.push({
          senderId:      senderId,
          assignedLeads: assignedLeads,
          inboxRows:     inboxRows
        });
      });
    }
    // Fire ALL requests across ALL campaigns in one go
    for (let b = 0; b < allFetchRequests.length; b += 50) {
      const batchRequests = allFetchRequests.slice(b, b + 50);
      const batchMeta     = allFetchMeta.slice(b, b + 50);

      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(batchRequests);
      } catch (e) {
        log('distributeFollowups: fetchAll error — ' + e.message, 'ERROR');
        continue;
      }

      for (let r = 0; r < responses.length; r++) {
        const meta = batchMeta[r];
        const code = responses[r].getResponseCode();

        if (code === 200) {
          meta.assignedLeads.forEach(function(lead) {
            dbUpdates.push({
              rowIndex: lead.rowIndex,
              fields: {
                status:       'Queued',
                sequenceStep: lead.sequenceStep + 1
              }
            });
          });
          log('distributeFollowups: wrote ' + meta.inboxRows.length + ' rows to sender ' + meta.senderId, 'INFO');
        } else {
          log('distributeFollowups: webhook failed (' + code + ') for sender ' + meta.senderId, 'ERROR');
        }
      }
    }
    // --- Step 6: Batch update masterLeads ---
    if (dbUpdates.length > 0) {
      batchUpdateLeadsDB(leadsSheetId, dbUpdates);
      log('distributeFollowups: updated ' + dbUpdates.length + ' leads in DB', 'INFO');
    }

    log('distributeFollowups: completed successfully', 'INFO');

    // Same budget guard as distributeCold — see the note there.
    if (isTimeLimitApproaching_()) {
      log('distributeFollowups: no execution budget left for kickoff — deferring to a one-time trigger', 'WARN');
      scheduleDeferredKickoff_('kickoffFollowupChainsDeferred');
    } else {
      log('distributeFollowups: kicking off followup chains on all senders', 'INFO');
      fanOutToSenders_('kickoffFollowup', 'distributeFollowups');
    }

  } catch (e) {
    log('distributeFollowups: unexpected error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}


// ============================================================================
// SECTION 4 — getActiveSenders()
// Reads senderAccounts sheet and returns only Active senders with capacity info.
// Also reads each sender's existing Queued row counts to avoid overfilling.
// ============================================================================

/**
 * Reads senderAccounts sheet, returns array of active sender objects.
 *
 * @returns {Array} Array of sender objects:
 *   {
 *     emailID, emailAddress, webAppUrl, spreadsheetId,
 *     dailyLimit, sentToday, availableCapacity,
 *     existingColdQueued, existingFollowupQueued
 *   }
 *
 * Notes:
 *  - Reads entire senderAccounts sheet in ONE getValues() call.
 *  - Filters to isActive = 'Active' only.
 *  - availableCapacity = dailyLimit - sentToday, clamped to 0 minimum.
 *  - existingColdQueued and existingFollowupQueued are read from each
 *    sender's spreadsheet to prevent overfilling inboxes that still have
 *    Queued rows from a previous run.
 */
function getActiveSenders() {
  const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
  const sheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);

  if (!sheet) {
    log('getActiveSenders: senderAccounts sheet not found', 'ERROR');
    return [];
  }

  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    log('getActiveSenders: no sender rows found', 'WARN');
    return [];
  }

  // Read headers + all data in ONE call
  const headers  = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const data     = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

  // Build column index map from headers
  const col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  const activeSenders = [];

  for (let i = 0; i < data.length; i++) {
    const row      = data[i];
    const isActive = String(row[col['isActive']] || '').trim();

    if (isActive !== 'Active') { continue; }

    const dailyLimit   = parseInt(row[col['dailyLimit']]  || 25, 10); // HARDCODED — verify this
    const sentToday    = parseInt(row[col['sentToday']]   || 0,  10);
    const spreadsheetId = String(row[col['spreadsheetId']] || '').trim();
    const webAppUrl     = String(row[col['webAppUrl']]     || '').trim();
    const emailID       = String(row[col['emailID']]       || '').trim();

    if (!spreadsheetId || !webAppUrl || !emailID) {
      log('getActiveSenders: skipping sender with missing spreadsheetId/webAppUrl/emailID at row ' + (i + 2), 'WARN');
      continue;
    }

    const availableCapacity = Math.max(0, dailyLimit - sentToday);

    // Read existing Queued counts from sender's inbox sheets
    // This prevents overfilling when previous runs left Queued rows
    let existingColdQueued    = 0;
    let existingFollowupQueued = 0;

    activeSenders.push({
      emailID:               emailID,
      emailAddress:          String(row[col['emailAddress']]   || '').trim(),
      emailFirstName:        String(row[col['emailFirstName']] || '').trim(),
      emailLastName:         String(row[col['emailLastName']]  || '').trim(),
      webAppUrl:             webAppUrl,
      spreadsheetId:         spreadsheetId,
      dailyLimit:            dailyLimit,
      sentToday:             sentToday,
      availableCapacity:     availableCapacity,
      existingColdQueued:    existingColdQueued,
      existingFollowupQueued: existingFollowupQueued
    });
  }

  log('getActiveSenders: found ' + activeSenders.length + ' active senders', 'INFO');
  return activeSenders;
}


// ============================================================================
// SECTION 5 — getEligibleColdLeads()
// Reads masterLeads and returns leads eligible for initial cold send.
// Filter: status=Pending, sequenceStep=0, campaignStatus=Active, replyStatus=''
// ============================================================================

/**
 * Reads masterLeads and returns cold-eligible leads.
 *
 * @param {string} leadsSheetId - Google Sheets ID of the Leads Database.
 * @returns {{ leads: Array, headers: Array }}
 *   leads:   Array of lead objects with all fields plus rowIndex and rawRow.
 *   headers: Array of column header strings (used by buildLeadDataJson).
 *
 * Reads the ENTIRE masterLeads sheet in ONE getValues() call.
 * All filtering is done in-memory — never reads row by row.
 */
function getEligibleColdLeads(leadsSheetId) {
  const result = { leads: [], headers: [] };

  try {
    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName(LEADS_SHEET_NAME);

    if (!sheet) {
      log('getEligibleColdLeads: masterLeads sheet not found', 'ERROR');
      return result;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      log('getEligibleColdLeads: no lead rows found', 'INFO');
      return result;
    }

    // Read entire sheet in ONE call (capped at MAX_LEADS_TO_SCAN for safety)
    const rowsToRead = Math.min(lastRow - 1, MAX_LEADS_TO_SCAN);
    const headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data       = sheet.getRange(2, 1, rowsToRead, sheet.getLastColumn()).getValues();

    result.headers = headers;

    // Build column index map
    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Filter in-memory — no sheet reads inside the loop
    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const status         = String(row[col['status']]         || '').trim();
      const sequenceStep   = parseInt(row[col['sequenceStep']] || 0, 10);
      const campaignStatus = String(row[col['campaignStatus']] || '').trim();
      const replyStatus    = String(row[col['replyStatus']]    || '').trim();

      // Cold eligible: never contacted, active campaign, no reply
      const bounceStatus = String(row[col['bounceStatus']] || '').trim();

      if (status !== 'Pending')        { continue; }
      if (sequenceStep !== 0)          { continue; }
      if (campaignStatus !== 'Active') { continue; }
      if (replyStatus !== '')          { continue; }
      if (bounceStatus === 'Bounced')  { continue; }

      result.leads.push({
        rowIndex:     i + 2, // +2: 1 for header row, 1 for 0-based index
        leadId:       String(row[col['leadId']]       || '').trim(),
        email:        String(row[col['email']]         || '').trim(),
        campaignId:   String(row[col['campaignId']]   || '').trim(),
        sequenceStep: sequenceStep,
        threadId:     '',
        rawRow:       row  // Full row array for buildLeadDataJson
      });
    }

  } catch (e) {
    log('getEligibleColdLeads: error — ' + e.message, 'ERROR');
  }

  return result;
}


// ============================================================================
// SECTION 6 — getEligibleFollowupLeads()
// Reads masterLeads and returns leads eligible for follow-up send.
// Filter: status=Sent, sequenceStep>0, sequenceStep<totalSequenceSteps,
//         campaignStatus=Active, replyStatus='', nextSentDate<=today.
// nextSentDate is the HARD gate — never queue before this date.
// ============================================================================

/**
 * Reads masterLeads and returns follow-up eligible leads.
 *
 * @param {string} leadsSheetId - Google Sheets ID of the Leads Database.
 * @returns {{ leads: Array, headers: Array }}
 */
function getEligibleFollowupLeads(leadsSheetId) {
  const result  = { leads: [], headers: [] };
  const today   = formatDate(new Date()); // 'YYYY-MM-DD'

  try {
    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName(LEADS_SHEET_NAME);

    if (!sheet) {
      log('getEligibleFollowupLeads: masterLeads sheet not found', 'ERROR');
      return result;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      log('getEligibleFollowupLeads: no lead rows found', 'INFO');
      return result;
    }

    const rowsToRead = Math.min(lastRow - 1, MAX_LEADS_TO_SCAN);
    const headers    = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data       = sheet.getRange(2, 1, rowsToRead, sheet.getLastColumn()).getValues();

    result.headers = headers;

    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    for (let i = 0; i < data.length; i++) {
      const row = data[i];

      const status             = String(row[col['status']]             || '').trim();
      const sequenceStep       = parseInt(row[col['sequenceStep']]     || 0,  10);
      const totalSequenceSteps = parseInt(row[col['totalSequenceSteps']]|| 0, 10);
      const campaignStatus     = String(row[col['campaignStatus']]     || '').trim();
      const replyStatus        = String(row[col['replyStatus']]        || '').trim();
      const nextSentDateRaw = row[col['nextSentDate']];
      const nextSentDate = nextSentDateRaw instanceof Date
        ? formatDate(nextSentDateRaw)
        : String(nextSentDateRaw || '').trim();
      const sequenceStatusVal  = String(row[col['sequenceStatus']]     || '').trim();

      const bounceStatus = String(row[col['bounceStatus']] || '').trim();

      if (status !== 'Sent')                 { continue; }
      if (sequenceStep <= 0)                 { continue; }
      if (sequenceStep >= totalSequenceSteps) { continue; }
      if (campaignStatus !== 'Active')       { continue; }
      if (replyStatus !== '')                { continue; }
      if (bounceStatus === 'Bounced')        { continue; }
      if (sequenceStatusVal === 'Completed') { continue; }
      if (nextSentDate === '')               { continue; }
      if (nextSentDate > today)              { continue; }

      result.leads.push({
        rowIndex:         i + 2,
        leadId:           String(row[col['leadId']]       || '').trim(),
        email:            String(row[col['email']]         || '').trim(),
        campaignId:       String(row[col['campaignId']]   || '').trim(),
        sequenceStep:     sequenceStep,
        totalSequenceSteps: totalSequenceSteps,
        threadId:         String(row[col['threadId']]     || '').trim(),
        ownerSender:      String(row[col['ownerSender']]  || '').trim(),
        rawRow:           row
      });
    }

  } catch (e) {
    log('getEligibleFollowupLeads: error — ' + e.message, 'ERROR');
  }

  return result;
}


// ============================================================================
// SECTION 7 — assignLeadsToSenders()
// Capacity-based assignment. Highest available capacity gets leads first.
// Returns a Map: senderId → [array of lead objects].
// The slotField parameter is either 'coldSlots' or 'followupSlots'.
// ============================================================================

/**
 * Assigns leads to senders by available capacity (highest capacity first).
 *
 * @param {Array}  leads     - Array of lead objects from getEligibleColdLeads or getEligibleFollowupLeads.
 * @param {Array}  senders   - Array of sender objects with slot counts attached.
 * @param {string} slotField - Which slot field to use: 'coldSlots' or 'followupSlots'.
 * @returns {Map} senderId (string) → Array of assigned lead objects.
 *
 * Rules:
 *  - Sort senders by slotField descending (highest capacity gets leads first).
 *  - Assign leads sequentially until each sender's slots are filled.
 *  - Never assign more leads than a sender's slots.
 *  - If total leads < total slots, some senders get fewer than their max.
 *  - If total leads > total slots, leftover leads are not assigned (logged).
 *
 * For follow-ups: leads are filtered to ownerSender before assignment
 * so follow-ups always go to the sender who owns that lead.
 */
function assignLeadsToSenders(leads, senders, slotField, distributionMethod) {
  const assignments = new Map();

  const eligibleSenders = senders.filter(function(s) {
    return (s[slotField] || 0) > 0;
  });

  if (eligibleSenders.length === 0) { return assignments; }

  // For follow-ups: leads must go to their ownerSender only
  if (slotField === 'followupSlots') {
    const eligibleSendersMap = {};
    eligibleSenders.forEach(function(s) { eligibleSendersMap[s.emailID] = s; });

    for (let i = 0; i < leads.length; i++) {
      const lead   = leads[i];
      const sender = eligibleSendersMap[lead.ownerSender];
      if (!sender) { continue; }
      if (!assignments.has(sender.emailID)) {
        assignments.set(sender.emailID, []);
      }
      const assigned = assignments.get(sender.emailID);
      if (assigned.length < (sender[slotField] || 0)) {
        assigned.push(lead);
      }
    }
    return assignments;
  }

  // For cold: round-robin across eligible senders
  const senderQueues = eligibleSenders.map(function(s) {
    return {
      emailID:  s.emailID,
      slots:    s[slotField] || 0,
      assigned: 0
    };
  });

  let leadIndex = 0;

  if (distributionMethod === 'capacityBased') {
    const totalSlots = senderQueues.reduce(function(sum, s) { return sum + s.slots; }, 0);
    if (totalSlots === 0) { return assignments; }

    for (let s = 0; s < senderQueues.length; s++) {
      const sq    = senderQueues[s];
      const share = Math.round((sq.slots / totalSlots) * leads.length);
      const count = Math.min(share, sq.slots, leads.length - leadIndex);
      if (count <= 0) { continue; }
      if (!assignments.has(sq.emailID)) { assignments.set(sq.emailID, []); }
      for (let i = 0; i < count; i++) {
        assignments.get(sq.emailID).push(leads[leadIndex++]);
      }
      sq.assigned += count;
    }

    while (leadIndex < leads.length) {
      const remaining = senderQueues.filter(function(s) { return s.assigned < s.slots; });
      if (remaining.length === 0) { break; }
      remaining.sort(function(a, b) { return (b.slots - b.assigned) - (a.slots - a.assigned); });
      const sq = remaining[0];
      if (!assignments.has(sq.emailID)) { assignments.set(sq.emailID, []); }
      assignments.get(sq.emailID).push(leads[leadIndex++]);
      sq.assigned++;
    }

  } else {
    let madeProgress = true;
    while (leadIndex < leads.length && madeProgress) {
      madeProgress = false;
      for (let s = 0; s < senderQueues.length; s++) {
        const sq = senderQueues[s];
        if (sq.assigned >= sq.slots)   { continue; }
        if (leadIndex >= leads.length) { break; }
        if (!assignments.has(sq.emailID)) { assignments.set(sq.emailID, []); }
        assignments.get(sq.emailID).push(leads[leadIndex]);
        sq.assigned++;
        leadIndex++;
        madeProgress = true;
      }
    }
  }

  assignments.forEach(function(assignedLeads, senderId) {
    log('assignLeadsToSenders: assigned ' + assignedLeads.length + ' leads to ' + senderId, 'INFO');
  });

  if (leadIndex < leads.length) {
    log('assignLeadsToSenders: ' + (leads.length - leadIndex) + ' leads could not be assigned — insufficient capacity', 'INFO');
  }

  return assignments;
}


// ============================================================================
// SECTION 8 — writeToSenderInbox()
// Makes a single HTTP POST to a sender's webApp with all inbox rows.
// Uses UrlFetchApp.fetchAll() for batched calls across multiple senders.
// Never called in a loop lead-by-lead — always passes the full row array.
// ============================================================================

/**
 * POSTs inbox rows to a single sender's webApp.
 *
 * @param {string} webhookUrl - The sender's deployed webApp URL.
 * @param {Array}  rows       - Array of inbox row objects to write.
 * @param {string} action     - 'writeColdInbox' or 'writeFollowupInbox'.
 * @returns {boolean} true if HTTP 200 received, false on any failure.
 *
 * Payload structure:
 *   { action: 'writeColdInbox', rows: [...] }
 *
 * Notes:
 *  - Rows are sent in batches of WEBHOOK_BATCH_SIZE (50) if the array is large.
 *  - Each batch is a separate POST call.
 *  - Non-200 responses are logged but do not throw — returns false.
 */
function writeToSenderInbox(webhookUrl, rows, action) {
  if (!webhookUrl) {
    log('writeToSenderInbox: missing webhookUrl for action ' + action, 'ERROR');
    return false;
  }

  if (!rows || rows.length === 0) {
    log('writeToSenderInbox: no rows to write for action ' + action, 'WARN');
    return false;
  }

  let allSuccess = true;

  // Send in batches of WEBHOOK_BATCH_SIZE to stay within UrlFetchApp limits
  for (let i = 0; i < rows.length; i += WEBHOOK_BATCH_SIZE) {
    const batch   = rows.slice(i, i + WEBHOOK_BATCH_SIZE);
    const payload = JSON.stringify({ action: action, rows: batch });

    const options = {
      method:      'post',
      contentType: 'application/json',
      payload:     payload,
      muteHttpExceptions: true // Don't throw on 4xx/5xx — we handle response codes
    };

    try {
      const response     = UrlFetchApp.fetch(webhookUrl, options);
      const responseCode = response.getResponseCode();

      if (responseCode !== 200) {
        log('writeToSenderInbox: non-200 response ' + responseCode + ' from ' + webhookUrl +
            ' (batch ' + Math.floor(i / WEBHOOK_BATCH_SIZE + 1) + ')', 'ERROR');
        allSuccess = false;
      } else {
        log('writeToSenderInbox: batch ' + Math.floor(i / WEBHOOK_BATCH_SIZE + 1) +
            ' of ' + Math.ceil(rows.length / WEBHOOK_BATCH_SIZE) + ' sent to ' + webhookUrl, 'INFO');
      }
    } catch (e) {
      log('writeToSenderInbox: fetch error — ' + e.message, 'ERROR');
      allSuccess = false;
    }
  }

  return allSuccess;
}


// ============================================================================
// SECTION 9 — batchUpdateLeadsDB()
// Reads masterLeads once, applies all updates in-memory, writes back in ONE
// setValues() call. Never writes row by row. Never calls openById() twice.
// ============================================================================

/**
 * Batch-updates fields in masterLeads for multiple rows in a single write.
 *
 * @param {string} leadsSheetId - Google Sheets ID of the Leads Database.
 * @param {Array}  updates      - Array of { rowIndex, fields } objects.
 *   rowIndex: 1-based sheet row number (as returned by getEligibleLeads).
 *   fields:   Plain object of { columnName: newValue } pairs.
 *
 * Example:
 *   batchUpdateLeadsDB(id, [
 *     { rowIndex: 5, fields: { status: 'Queued', ownerSender: 'S001' } },
 *     { rowIndex: 8, fields: { status: 'Queued', ownerSender: 'S002' } }
 *   ]);
 *
 * Notes:
 *  - Opens masterLeads ONCE. Reads all data into memory ONCE.
 *  - Applies ALL updates to the in-memory 2D array.
 *  - Writes the entire array back in ONE setValues() call.
 *  - Column name → index mapping is built once from the header row.
 *  - Skips updates for unknown column names (logs a WARN).
 */
function batchUpdateLeadsDB(leadsSheetId, updates) {
  if (!updates || updates.length === 0) {
    log('batchUpdateLeadsDB: no updates to apply', 'WARN');
    return;
  }

  try {
    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName(LEADS_SHEET_NAME);

    if (!sheet) {
      log('batchUpdateLeadsDB: masterLeads sheet not found', 'ERROR');
      return;
    }

    const lastRow = sheet.getLastRow();
    const lastCol = sheet.getLastColumn();

    if (lastRow < 2) {
      log('batchUpdateLeadsDB: no data rows to update', 'WARN');
      return;
    }

    // Read headers and ALL data in ONE call each
    const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, lastCol).getValues();

    // Build column name → 0-based index map
    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Apply all updates to the in-memory array
    for (let u = 0; u < updates.length; u++) {
      const update   = updates[u];
      const rowIndex = update.rowIndex; // 1-based sheet row
      const dataIdx  = rowIndex - 2;   // Convert to 0-based data array index (-1 header, -1 0-based)

      if (dataIdx < 0 || dataIdx >= data.length) {
        log('batchUpdateLeadsDB: rowIndex ' + rowIndex + ' out of range — skipping', 'WARN');
        continue;
      }

      const fields = update.fields || {};
      const fieldKeys = Object.keys(fields);

      for (let f = 0; f < fieldKeys.length; f++) {
        const fieldName = fieldKeys[f];
        if (!(fieldName in col)) {
          log('batchUpdateLeadsDB: unknown column "' + fieldName + '" — skipping', 'WARN');
          continue;
        }
        data[dataIdx][col[fieldName]] = fields[fieldName];
      }
    }

    // Write entire data array back in ONE setValues() call
    sheet.getRange(2, 1, data.length, lastCol).setValues(data);
    SpreadsheetApp.flush(); // Confirm write before function returns

    log('batchUpdateLeadsDB: wrote ' + updates.length + ' updates to masterLeads', 'INFO');

  } catch (e) {
    log('batchUpdateLeadsDB: error — ' + e.message, 'ERROR');
  }
}


// ============================================================================
// SECTION 10 — Execution Time Guard
// ============================================================================

const MAX_EXECUTION_MS = 4 * 60 * 1000; // 4 minutes
let EXECUTION_START_TIME = null;

function startExecutionTimer_() {
  EXECUTION_START_TIME = new Date().getTime();
}

function isTimeLimitApproaching_() {
  if (!EXECUTION_START_TIME) { return false; }
  return (new Date().getTime() - EXECUTION_START_TIME) >= MAX_EXECUTION_MS;
}

/**
 * Saves a SMALL checkpoint — only campaign index and IDs, never lead data.
 * Re-reads leads fresh on continuation to stay under 100KB cache limit.
 */
function saveCheckpointAndChain_(functionName, checkpoint) {
  try {
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'checkpoint_' + functionName;
    const payload  = JSON.stringify(checkpoint);

    if (payload.length > 90000) {
      log('saveCheckpointAndChain_: payload too large (' + payload.length + ' bytes) — cannot checkpoint', 'ERROR');
      return false;
    }

    cache.put(cacheKey, payload, 600); // 10 min TTL

    const handlerName = functionName === 'distributeCold'
      ? 'continueColdDistribution'
      : 'continueFollowupDistribution';

    ScriptApp.newTrigger(handlerName)
      .timeBased()
      .after(60 * 1000) // 1 minute
      .create();

    log('saveCheckpointAndChain_: checkpoint saved for ' + functionName + ' — resuming in 1 min', 'INFO');
    return true;

  } catch (e) {
    log('saveCheckpointAndChain_: error — ' + e.message, 'ERROR');
    return false;
  }
}

function loadCheckpoint_(functionName) {
  try {
    const cache    = CacheService.getScriptCache();
    const cacheKey = 'checkpoint_' + functionName;
    const payload  = cache.get(cacheKey);
    if (!payload) { return null; }
    cache.remove(cacheKey);
    return JSON.parse(payload);
  } catch (e) {
    log('loadCheckpoint_: error — ' + e.message, 'ERROR');
    return null;
  }
}

function deleteOwnContinuationTrigger_(handlerName) {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    for (let i = 0; i < triggers.length; i++) {
      if (triggers[i].getHandlerFunction() === handlerName) {
        ScriptApp.deleteTrigger(triggers[i]);
      }
    }
  } catch (e) {
    log('deleteOwnContinuationTrigger_: error — ' + e.message, 'WARN');
  }
}


// new function 

/**
 * Creates a one-time trigger for a deferred sender fan-out.
 * Used when a distribution run has no execution budget left to fan out inline.
 *
 * @param {string} handlerName - Function to run in a fresh execution.
 * @returns {boolean} true if the trigger was created.
 */
function scheduleDeferredKickoff_(handlerName) {
  try {
    ScriptApp.newTrigger(handlerName)
      .timeBased()
      .after(60 * 1000)
      .create();
    log('scheduleDeferredKickoff_: ' + handlerName + ' scheduled in 1 min', 'INFO');
    return true;
  } catch (e) {
    log('scheduleDeferredKickoff_: could not schedule ' + handlerName + ' — ' + e.message, 'ERROR');
    return false;
  }
}

/**
 * Deferred cold kickoff. Runs in its own execution with a full 6-minute budget.
 */
function kickoffColdChainsDeferred() {
  deleteOwnContinuationTrigger_('kickoffColdChainsDeferred');
  fanOutToSenders_('kickoffCold', 'kickoffColdChainsDeferred');
}

/**
 * Deferred followup kickoff. Runs in its own execution with a full 6-minute budget.
 */
function kickoffFollowupChainsDeferred() {
  deleteOwnContinuationTrigger_('kickoffFollowupChainsDeferred');
  fanOutToSenders_('kickoffFollowup', 'kickoffFollowupChainsDeferred');
}

function runDistributeColdNow() {
  deleteOwnContinuationTrigger_('runDistributeColdNow');
  log('runDistributeColdNow: manual cold distribution requested via API', 'INFO');
  distributeCold();
}

function runDistributeFollowupsNow() {
  deleteOwnContinuationTrigger_('runDistributeFollowupsNow');
  log('runDistributeFollowupsNow: manual followup distribution requested via API', 'INFO');
  distributeFollowups();
}

/**
 * Continuation trigger for distributeCold.
 * Deletes itself then re-runs distributeCold which detects the checkpoint.
 */
function continueColdDistribution() {
  deleteOwnContinuationTrigger_('continueColdDistribution');
  log('continueColdDistribution: resuming distributeCold', 'INFO');
  distributeCold();
}

/**
 * Continuation trigger for distributeFollowups.
 * Deletes itself then re-runs distributeFollowups which detects the checkpoint.
 */
function continueFollowupDistribution() {
  deleteOwnContinuationTrigger_('continueFollowupDistribution');
  log('continueFollowupDistribution: resuming distributeFollowups', 'INFO');
  distributeFollowups();
}


function setupAllMasterTriggers() {
  // Delete ALL existing time-based triggers first
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }

  // 6:00 AM — Reset sent today counters
  ScriptApp.newTrigger('resetSentToday')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(0)
    .inTimezone('Africa/Cairo')
    .create();

    ScriptApp.newTrigger('runWarmupIncrease')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(5)
    .inTimezone('Africa/Cairo')
    .create();

  // 6:30 AM — Scan for replies
  ScriptApp.newTrigger('scanAllSendersForReplies')
    .timeBased()
    .everyDays(1)
    .atHour(6)
    .nearMinute(30)
    .inTimezone('Africa/Cairo')
    .create();

  // 7:00 AM — Distribute cold emails
  ScriptApp.newTrigger('distributeCold')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(0)
    .inTimezone('Africa/Cairo')
    .create();

    ScriptApp.newTrigger('sendDailyDigest')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(5)
    .inTimezone('Africa/Cairo')
    .create();

  // 7:30 AM — Distribute followups
  ScriptApp.newTrigger('distributeFollowups')
    .timeBased()
    .everyDays(1)
    .atHour(7)
    .nearMinute(30)
    .inTimezone('Africa/Cairo')
    .create();

  // Every 20 minutes — Process staging updates
  ScriptApp.newTrigger('processStagingUpdates')
    .timeBased()
    .everyMinutes(10)
    .create();

  // 11:00 PM — Scan for bounces
  ScriptApp.newTrigger('scanAllSendersForBounces')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(0)
    .inTimezone('Africa/Cairo')
    .create();

  // 11:30 PM — Sync campaign stats
  ScriptApp.newTrigger('runSyncStats')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(30)
    .inTimezone('Africa/Cairo')
    .create();

  // 11:55 PM — Daily reset
  ScriptApp.newTrigger('runDailyReset')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(55)
    .inTimezone('Africa/Cairo')
    .create();

  log('setupAllMasterTriggers: all 8 triggers created successfully', 'INFO');
  Logger.log('✅ All 8 Master Control triggers created successfully.');
}

// ============================================================================
// RAMP UP — Daily sender limit warmup function
// Runs daily at 6:05 AM Cairo — after resetSentToday, before distributeCold
// Settings sheet keys required:
//   warmupEnabled        → TRUE/FALSE
//   warmupStartLimit     → starting dailyLimit (e.g. 1)
//   warmupIncreaseBy     → how much to add each time (e.g. 1)
//   warmupEveryDays      → days between each increase (e.g. 3)
//   warmupMaxLimit       → stop increasing at this number (e.g. 25)
//   warmupNextIncreaseDate → YYYY-MM-DD of next scheduled increase
//   warmupStartDate      → YYYY-MM-DD when warmup began
// ============================================================================

function runWarmupIncrease() {
  log('runWarmupIncrease: started', 'INFO');

  try {
    const settings = getSettings(MASTER_SHEET_ID);

    // ── Check if warmup is enabled ────────────────────────────────────────
    const warmupEnabled = String(settings.warmupEnabled || '').trim().toUpperCase() === 'TRUE';
    if (!warmupEnabled) {
      log('runWarmupIncrease: warmupEnabled=FALSE — exiting', 'INFO');
      return;
    }

    // ── Read warmup settings ──────────────────────────────────────────────
    const warmupIncreaseBy  = parseInt(settings.warmupIncreaseBy  || 1,  10);
    const warmupEveryDays   = parseInt(settings.warmupEveryDays   || 3,  10);
    const warmupMaxLimit    = parseInt(settings.warmupMaxLimit     || 25, 10);
    const today = formatDate(new Date());

    // Handle both Date objects and strings from Sheets
    const rawNextDate    = settings.warmupNextIncreaseDate;
    const warmupNextDate = (rawNextDate instanceof Date)
      ? formatDate(rawNextDate)
      : String(rawNextDate || '').trim();

    log('runWarmupIncrease: today=' + today + ' nextIncreaseDate=' + warmupNextDate, 'INFO');

    // ── Check if today is increase day ────────────────────────────────────
    if (!warmupNextDate || today < warmupNextDate) {
      log('runWarmupIncrease: not increase day yet — next increase on ' + warmupNextDate, 'INFO');
      return;
    }

    // ── Read current dailyLimit from senderAccounts ───────────────────────
    const ss          = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const saSheet     = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);
    const saLastRow   = saSheet.getLastRow();
    if (saLastRow < 2) {
      log('runWarmupIncrease: no senders found — exiting', 'WARN');
      return;
    }

    const saHeaders = saSheet.getRange(1, 1, 1, saSheet.getLastColumn()).getValues()[0];
    const saData    = saSheet.getRange(2, 1, saLastRow - 1, saSheet.getLastColumn()).getValues();
    const saCol     = {};
    saHeaders.forEach(function(h, i) { saCol[String(h).trim()] = i; });

        // ── Ramp each opted-in sender from its own current limit ──────────────
    // Reading one sender's limit and applying it to everyone dragged
    // established accounts back down whenever new ones were added.
    const hasRampCol = ('rampUp' in saCol);
    const ramping    = [];   // { rowIndex, emailID, from, to, done }

    for (let i = 0; i < saData.length; i++) {
      if (String(saData[i][saCol['isActive']] || '').trim() !== 'Active') { continue; }

      // Without the column, fall back to ramping everyone — the old behaviour,
      // so this patch is safe before the column is added.
      const flag = hasRampCol
        ? String(saData[i][saCol['rampUp']] || '').trim().toLowerCase()
        : 'active';
      if (flag !== 'active' && flag !== 'true' && flag !== 'yes') { continue; }

      const from = parseInt(saData[i][saCol['dailyLimit']] || 1, 10);
      if (from >= warmupMaxLimit) {
        // Already there: mark it done so it stops being considered.
        if (hasRampCol) { saData[i][saCol['rampUp']] = 'Done'; }
        continue;
      }

      let to   = from + warmupIncreaseBy;
      let done = false;
      if (to >= warmupMaxLimit) { to = warmupMaxLimit; done = true; }

      if (hasRampCol && done) { saData[i][saCol['rampUp']] = 'Done'; }

      saData[i][saCol['dailyLimit']] = to;
      ramping.push({
        emailID: String(saData[i][saCol['emailID']] || '').trim(),
        from: from, to: to, done: done
      });
    }

    if (ramping.length === 0) {
      log('runWarmupIncrease: no senders are flagged to ramp — nothing to do', 'INFO');
      return;
    }

    // Write the sheet once, then push each new limit to that sender's Config.
    saSheet.getRange(2, 1, saData.length, saHeaders.length).setValues(saData);
    SpreadsheetApp.flush();

    let synced = 0, failed = 0;
    ramping.forEach(function(r) {
      const res = handleUpdateSenderConfig({
        senderIds: [r.emailID],
        fields:    { dailyLimit: r.to }
      });
      synced += (res.synced || 0);
      failed += (res.failed || 0);
    });

    // Warmup is complete only when nothing is left ramping.
    const stillRamping = hasRampCol
      ? saData.some(function(row) {
          return String(row[saCol['rampUp']] || '').trim().toLowerCase() === 'active';
        })
      : ramping.some(function(r) { return !r.done; });

    const warmupComplete = !stillRamping;
    const newLimit = Math.max.apply(null, ramping.map(function(r) { return r.to; }));

    log('runWarmupIncrease: ramped ' + ramping.length + ' sender(s), ' +
        'synced=' + synced + ' failed=' + failed +
        ' complete=' + warmupComplete, 'INFO');

    // ── Update Settings sheet ─────────────────────────────────────────────
    const settingsSheet = ss.getSheetByName('Settings');
    const settingsData  = settingsSheet.getRange(1, 1, settingsSheet.getLastRow(), 2).getValues();
    const settingsMap   = {};
    settingsData.forEach(function(row, i) {
      settingsMap[String(row[0]).trim()] = i;
    });

    const nextIncreaseDate = warmupComplete
      ? ''
      : addDays(today, warmupEveryDays);

    // Update keys in Settings
    const updates = {
      'warmupNextIncreaseDate': "'" + nextIncreaseDate,
      'warmupLastIncreaseDate': "'" + today,
      'warmupCurrentLimit':     String(newLimit),
      'warmupEnabled':          warmupComplete ? 'FALSE' : 'TRUE'
    };

    Object.keys(updates).forEach(function(key) {
      if (key in settingsMap) {
        settingsData[settingsMap[key]][1] = updates[key];
      } else {
        settingsData.push([key, updates[key]]);
      }
    });

    settingsSheet.getRange(1, 1, settingsData.length, 2).setValues(settingsData);
    SpreadsheetApp.flush();

    log('runWarmupIncrease: Settings updated — nextIncreaseDate=' + nextIncreaseDate +
        ' warmupComplete=' + warmupComplete, 'INFO');

    // ── Per-sender warmup tracking in senderAccounts ──────────────────────
    // Update warmupCurrentLimit column if it exists
    if ('warmupCurrentLimit' in saCol) {
      for (let i = 0; i < saData.length; i++) {
        saData[i][saCol['warmupCurrentLimit']] = newLimit;
      }
      saSheet.getRange(2, 1, saData.length, saHeaders.length).setValues(saData);
      SpreadsheetApp.flush();
    }

    log('runWarmupIncrease: completed — all senders now at ' + newLimit + '/day', 'INFO');

  } catch(e) {
    log('runWarmupIncrease: error — ' + e.message, 'ERROR');
  }
}

// ============================================================================
// DAILY DIGEST EMAIL
// Runs at 7:05 AM Cairo — after distributeCold and distributeFollowups
// Sends a summary email to alertsEmail in Settings
// ============================================================================

function sendDailyDigest() {
  log('sendDailyDigest: started', 'INFO');

  try {
    const settings    = getSettings(MASTER_SHEET_ID);
    const alertsEmail = String(settings.alertsEmail || '').trim();

    if (!alertsEmail) {
      log('sendDailyDigest: no alertsEmail in Settings — skipping', 'WARN');
      return;
    }

    const today = formatDate(new Date());

    // ── 1. Read masterLeads stats ─────────────────────────────────────────
    const leadsStats = { total: 0, pending: 0, queued: 0, sent: 0,
                         replied: 0, bounced: 0, completed: 0, error: 0 };
    const campStats  = {};

    try {
      const leadsSS    = SpreadsheetApp.openById(settings.leadsSpreadsheetId);
      const leadsSheet = leadsSS.getSheetByName('masterLeads');

      if (leadsSheet && leadsSheet.getLastRow() > 1) {
        const headers = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
        const data    = leadsSheet.getRange(2, 1, leadsSheet.getLastRow() - 1,
                                            leadsSheet.getLastColumn()).getValues();
        const col     = {};
        headers.forEach(function(h, i) { col[String(h).trim()] = i; });

        for (let i = 0; i < data.length; i++) {
          const status         = String(data[i][col['status']]         || '').trim().toLowerCase();
          const replyStatus    = String(data[i][col['replyStatus']]    || '').trim().toLowerCase();
          const bounceStatus   = String(data[i][col['bounceStatus']]   || '').trim().toLowerCase();
          const sequenceStatus = String(data[i][col['sequenceStatus']] || '').trim().toLowerCase();
          const campaignId     = String(data[i][col['campaignId']]     || '').trim();

          leadsStats.total++;

          if (!campStats[campaignId]) {
            campStats[campaignId] = { pending: 0, queued: 0, sent: 0,
                                      replied: 0, bounced: 0, completed: 0 };
          }

          if (bounceStatus === 'bounced') {
            leadsStats.bounced++;
            campStats[campaignId].bounced++;
          } else if (replyStatus === 'replied') {
            leadsStats.replied++;
            campStats[campaignId].replied++;
          } else if (sequenceStatus === 'completed') {
            leadsStats.completed++;
            campStats[campaignId].completed++;
          } else if (status === 'pending') {
            leadsStats.pending++;
            campStats[campaignId].pending++;
          } else if (status === 'queued') {
            leadsStats.queued++;
            campStats[campaignId].queued++;
          } else if (status === 'sent') {
            leadsStats.sent++;
            campStats[campaignId].sent++;
          } else if (status === 'error') {
            leadsStats.error++;
          }
        }
      }
    } catch(e) {
      log('sendDailyDigest: error reading masterLeads — ' + e.message, 'WARN');
    }

        // ── 1b. Seed every campaign, including ones with no leads ─────────────
    try {
      const campSheet = SpreadsheetApp.openById(MASTER_SHEET_ID).getSheetByName('Campaigns');

      if (campSheet && campSheet.getLastRow() > 1) {
        const cHeaders = campSheet.getRange(1, 1, 1, campSheet.getLastColumn()).getValues()[0];
        const cData    = campSheet.getRange(2, 1, campSheet.getLastRow() - 1,
                                            campSheet.getLastColumn()).getValues();
        const cCol = {};
        cHeaders.forEach(function(h, i) { cCol[String(h).trim()] = i; });

        for (let i = 0; i < cData.length; i++) {
          const cId = String(cData[i][cCol['campaignId']] || '').trim();
          if (!cId) { continue; }

          const name   = String(cData[i][cCol['campaignName']]   || '').trim() || cId;
          const status = String(cData[i][cCol['campaignStatus']] || '').trim();

          if (!campStats[cId]) {
            campStats[cId] = { pending: 0, queued: 0, sent: 0,
                               replied: 0, bounced: 0, completed: 0 };
          }
          campStats[cId].name   = name;
          campStats[cId].status = status;
        }
      }
    } catch (e) {
      log('sendDailyDigest: error reading Campaigns — ' + e.message, 'WARN');
    }

    // ── 2. Read senderAccounts stats ─────────────────────────────────────
    const senderStats = { total: 0, active: 0, paused: 0,
                          sentToday: 0, totalSentCount: 0,
                          repliesReceived: 0, bouncedEmails: 0,
                          warmupEnabled: false, warmupCurrentLimit: 0 };

    try {
      const ss      = SpreadsheetApp.openById(MASTER_SHEET_ID);
      const saSheet = ss.getSheetByName(SHEET_SENDER_ACCOUNTS);

      if (saSheet && saSheet.getLastRow() > 1) {
        const headers = saSheet.getRange(1, 1, 1, saSheet.getLastColumn()).getValues()[0];
        const data    = saSheet.getRange(2, 1, saSheet.getLastRow() - 1,
                                         saSheet.getLastColumn()).getValues();
        const col     = {};
        headers.forEach(function(h, i) { col[String(h).trim()] = i; });

        for (let i = 0; i < data.length; i++) {
          senderStats.total++;
          const isActive = String(data[i][col['isActive']] || '').trim();
          if (isActive === 'Active') { senderStats.active++; }
          else                       { senderStats.paused++; }

          senderStats.sentToday      += parseInt(data[i][col['sentToday']]       || 0, 10);
          senderStats.totalSentCount += parseInt(data[i][col['totalSentCount']]  || 0, 10);
          senderStats.repliesReceived+= parseInt(data[i][col['repliesReceived']] || 0, 10);
          senderStats.bouncedEmails  += parseInt(data[i][col['bouncedEmails']]   || 0, 10);
        }
      }

      // Warmup status from Settings
      senderStats.warmupEnabled      = String(settings.warmupEnabled || '').toUpperCase() === 'TRUE';
      senderStats.warmupCurrentLimit = parseInt(settings.warmupCurrentLimit || 0, 10);
      senderStats.warmupMaxLimit     = parseInt(settings.warmupMaxLimit     || 25, 10);
      senderStats.warmupNextDate     = String(settings.warmupNextIncreaseDate || '').trim();

    } catch(e) {
      log('sendDailyDigest: error reading senderAccounts — ' + e.message, 'WARN');
    }

    // ── 3. Read daily summary (yesterday's numbers) ───────────────────────
    let yesterdaySent     = 0;
    let yesterdayReplies  = 0;
    let yesterdayBounces  = 0;

    try {
      const ss           = SpreadsheetApp.openById(MASTER_SHEET_ID);
      const summarySheet = ss.getSheetByName('dailySummary');

      if (summarySheet && summarySheet.getLastRow() > 1) {
        const lastDataRow = summarySheet.getLastRow();
        const lastRow     = summarySheet.getRange(lastDataRow, 1, 1,
                             summarySheet.getLastColumn()).getValues()[0];
        const headers     = summarySheet.getRange(1, 1, 1,
                             summarySheet.getLastColumn()).getValues()[0];
        const col         = {};
        headers.forEach(function(h, i) { col[String(h).trim()] = i; });

        yesterdaySent    = parseInt(lastRow[col['emailsSent']]      || 0, 10);
        yesterdayReplies = parseInt(lastRow[col['repliesReceived']] || 0, 10);
        yesterdayBounces = parseInt(lastRow[col['bouncedEmails']]   || 0, 10);
      }
    } catch(e) {
      log('sendDailyDigest: error reading dailySummary — ' + e.message, 'WARN');
    }

    // ── 4. Build email body ───────────────────────────────────────────────
    const contacted = leadsStats.sent + leadsStats.replied +
                      leadsStats.completed + leadsStats.bounced;

    const replyRate  = contacted > 0
      ? ((leadsStats.replied / contacted) * 100).toFixed(1) + '%'
      : '0%';
    const bounceRate = contacted > 0
      ? ((leadsStats.bounced / contacted) * 100).toFixed(1) + '%'
      : '0%';

    // Campaign breakdown
    let campLines = '';
    Object.keys(campStats).sort().forEach(function(cId) {
      const c = campStats[cId];
      campLines += cId + ': ' +
        c.pending   + ' pending, ' +
        c.queued    + ' queued, ' +
        c.sent      + ' sent, ' +
        c.replied   + ' replied, ' +
        c.completed + ' completed\n';
    });

    // Warmup line
    let warmupLine = 'Warmup: disabled';
    if (senderStats.warmupEnabled) {
      warmupLine = 'Warmup: ON — current limit ' + senderStats.warmupCurrentLimit +
                   '/' + senderStats.warmupMaxLimit +
                   ' — next increase ' + (senderStats.warmupNextDate || 'not set');
    } else if (senderStats.warmupCurrentLimit >= senderStats.warmupMaxLimit) {
      warmupLine = 'Warmup: COMPLETE — all senders at max limit ' +
                   senderStats.warmupMaxLimit + '/day';
    }

    const body = [
      '=== DAILY SEND REPORT — ' + today + ' ===',
      '',
      '--- YESTERDAY ---',
      'Emails sent:       ' + yesterdaySent,
      'Replies received:  ' + yesterdayReplies,
      'Bounces detected:  ' + yesterdayBounces,
      '',
      '--- TODAY DISTRIBUTED ---',
      'Senders active:    ' + senderStats.active + ' / ' + senderStats.total,
      'Sent today so far: ' + senderStats.sentToday,
      warmupLine,
      '',
      '--- LEAD DATABASE ---',
      'Total leads:       ' + leadsStats.total,
      'Pending:           ' + leadsStats.pending,
      'Queued today:      ' + leadsStats.queued,
      'Sent (all time):   ' + leadsStats.sent,
      'Replied:           ' + leadsStats.replied + ' (' + replyRate + ')',
      'Bounced:           ' + leadsStats.bounced + ' (' + bounceRate + ')',
      'Completed:         ' + leadsStats.completed,
      'Errors:            ' + leadsStats.error,
      '',
      '--- CAMPAIGNS ---',
      campLines,
      '--- ALL TIME ---',
      'Total sent ever:   ' + senderStats.totalSentCount,
      'Total replies:     ' + senderStats.repliesReceived,
      'Total bounces:     ' + senderStats.bouncedEmails,
      '',
      '=== END OF REPORT ==='
    ].join('\n');

        // ── 5. Send the email ─────────────────────────────────────────────────
      // htmlBody renders in every mail client; `body` stays as the plain-text
      // fallback for clients that refuse HTML.
      GmailApp.sendEmail(
        alertsEmail,
        'Daily Send Report — ' + today + ' — ' + senderStats.sentToday + ' sent today',
        body,
        {
          name:     'Reach X',
          htmlBody: buildDigestHtml_({
            today:            today,
            leadsStats:       leadsStats,
            senderStats:      senderStats,
            campStats:        campStats,
            contacted:        contacted,
            yesterdaySent:    yesterdaySent,
            yesterdayReplies: yesterdayReplies,
            yesterdayBounces: yesterdayBounces,
            replyRate:        replyRate,
            bounceRate:       bounceRate,
            warmupLine:       warmupLine
          })
        }
      );

    log('sendDailyDigest: sent to ' + alertsEmail, 'INFO');

  } catch(e) {
    log('sendDailyDigest: error — ' + e.message, 'ERROR');
  }
}

/**
 * Renders the daily digest as HTML.
 *
 * Written for mail clients, not browsers: tables for layout, every style
 * inline, no external images or fonts. Gmail strips <style> blocks and most
 * clients ignore flexbox, so neither is used.
 */
function buildDigestHtml_(d) {
  const esc = function(v) {
    return String(v === null || v === undefined ? '' : v)
      .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;');
  };
  const n = function(v) {
    return Number(v || 0).toLocaleString('en-US');
  };

  // One headline figure.
  const stat = function(label, value, sub, color) {
    return '' +
      '<td width="25%" style="padding:0 6px;" valign="top">' +
        '<table role="presentation" width="100%" cellpadding="0" cellspacing="0" ' +
               'style="background:#f6f8fc;border:1px solid #e3e8f0;border-radius:10px;">' +
          '<tr><td style="padding:14px 16px;">' +
            '<div style="font:500 11px/1.4 Arial,sans-serif;color:#64748b;' +
                 'text-transform:uppercase;letter-spacing:.5px;">' + esc(label) + '</div>' +
            '<div style="font:600 24px/1.3 Arial,sans-serif;color:' + (color || '#0f172a') + ';' +
                 'margin-top:4px;">' + esc(value) + '</div>' +
            (sub ? '<div style="font:400 12px/1.4 Arial,sans-serif;color:#64748b;">' +
                   esc(sub) + '</div>' : '') +
          '</td></tr>' +
        '</table>' +
      '</td>';
  };

  const sectionTitle = function(text) {
    return '<tr><td style="padding:26px 24px 10px;">' +
             '<div style="font:600 13px/1.4 Arial,sans-serif;color:#0f172a;' +
                  'text-transform:uppercase;letter-spacing:.6px;">' + esc(text) + '</div>' +
             '<div style="height:2px;width:34px;background:#4f8ef7;margin-top:6px;"></div>' +
           '</td></tr>';
  };

  const row = function(label, value, color) {
    return '<tr>' +
      '<td style="padding:9px 0;border-bottom:1px solid #eef1f6;' +
           'font:400 14px/1.4 Arial,sans-serif;color:#475569;">' + esc(label) + '</td>' +
      '<td align="right" style="padding:9px 0;border-bottom:1px solid #eef1f6;' +
           'font:600 14px/1.4 Arial,sans-serif;color:' + (color || '#0f172a') + ';">' +
           esc(value) + '</td>' +
    '</tr>';
  };

  // Campaign table rows
  let campRows = '';
  Object.keys(d.campStats).sort().forEach(function(cId) {
    const c    = d.campStats[cId];
    const done = c.sent + c.replied + c.completed + c.bounced;
    const rate = done > 0 ? ((c.replied / done) * 100).toFixed(1) + '%' : '—';

    const paused    = String(c.status || '').toLowerCase() === 'paused';
    const nameColor = paused ? '#94a3b8' : '#0f172a';
    const cell = 'padding:9px 8px;border-bottom:1px solid #eef1f6;' +
                 'font:400 13px/1.4 Arial,sans-serif;color:#475569;';

    campRows +=
      '<tr>' +
        '<td style="padding:9px 8px;border-bottom:1px solid #eef1f6;">' +
          '<div style="font:600 13px/1.4 Arial,sans-serif;color:' + nameColor + ';">' +
            esc(c.name || cId) + '</div>' +
          '<div style="font:400 11px/1.4 Arial,sans-serif;color:#94a3b8;">' +
            esc(cId) + (c.status ? ' &middot; ' + esc(c.status) : '') + '</div>' +
        '</td>' +
        '<td align="right" style="' + cell + '">' + n(c.pending)   + '</td>' +
        '<td align="right" style="' + cell + '">' + n(c.queued)    + '</td>' +
        '<td align="right" style="' + cell + '">' + n(c.sent)      + '</td>' +
        '<td align="right" style="padding:9px 8px;border-bottom:1px solid #eef1f6;' +
             'font:600 13px/1.4 Arial,sans-serif;color:#16a34a;">' + n(c.replied) + '</td>' +
        '<td align="right" style="' + cell + '">' + esc(rate)      + '</td>' +
        '<td align="right" style="' + cell + '">' + n(c.completed) + '</td>' +
      '</tr>';
  });
  if (!campRows) {
    campRows = '<tr><td colspan="7" style="padding:16px 8px;font:400 13px Arial,sans-serif;' +
               'color:#94a3b8;text-align:center;">No campaigns yet</td></tr>';
  }

  const th = function(text, align) {
    return '<td align="' + (align || 'left') + '" style="padding:0 8px 8px;' +
           'font:500 11px/1.4 Arial,sans-serif;color:#64748b;text-transform:uppercase;' +
           'letter-spacing:.4px;">' + esc(text) + '</td>';
  };

  return '' +
'<!DOCTYPE html><html><body style="margin:0;padding:0;background:#eef2f7;">' +
'<table role="presentation" width="100%" cellpadding="0" cellspacing="0" style="background:#eef2f7;padding:24px 12px;">' +
'<tr><td align="center">' +

  '<table role="presentation" width="640" cellpadding="0" cellspacing="0" ' +
         'style="width:640px;max-width:100%;background:#ffffff;border-radius:14px;overflow:hidden;' +
                'box-shadow:0 1px 3px rgba(15,23,42,.08);">' +

    // Header
    '<tr><td style="background:#0f1117;padding:22px 24px;">' +
      '<div style="font:600 18px/1.3 Arial,sans-serif;color:#ffffff;">Reach X</div>' +
      '<div style="font:400 13px/1.4 Arial,sans-serif;color:#8892a4;margin-top:2px;">' +
        'Daily send report &middot; ' + esc(d.today) + '</div>' +
    '</td></tr>' +

    // Yesterday
    sectionTitle('Yesterday') +
    '<tr><td style="padding:0 18px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0"><tr>' +
        stat('Emails sent', n(d.yesterdaySent)) +
        stat('Replies', n(d.yesterdayReplies), null, '#16a34a') +
        stat('Bounces', n(d.yesterdayBounces), null, d.yesterdayBounces > 0 ? '#dc2626' : null) +
        stat('Sent today', n(d.senderStats.sentToday)) +
      '</tr></table>' +
    '</td></tr>' +

    // Senders
    sectionTitle('Senders') +
    '<tr><td style="padding:0 24px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
        row('Active', n(d.senderStats.active) + ' of ' + n(d.senderStats.total),
            d.senderStats.active === 0 ? '#dc2626' : '#16a34a') +
        row('Sent today', n(d.senderStats.sentToday)) +
        row('Ramp up', d.warmupLine.replace(/^Warmup:\\s*/, '')) +
      '</table>' +
    '</td></tr>' +

    // Lead database
    sectionTitle('Lead database') +
    '<tr><td style="padding:0 24px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
        row('Total leads', n(d.leadsStats.total)) +
        row('Pending', n(d.leadsStats.pending)) +
        row('Queued', n(d.leadsStats.queued)) +
        row('Contacted', n(d.contacted)) +
        row('Awaiting reply', n(d.leadsStats.sent)) +
        row('Replied', n(d.leadsStats.replied) + '  (' + esc(d.replyRate) + ')', '#16a34a') +
        row('Bounced', n(d.leadsStats.bounced) + '  (' + esc(d.bounceRate) + ')',
            d.leadsStats.bounced > 0 ? '#dc2626' : null) +
        row('Completed', n(d.leadsStats.completed)) +
        (d.leadsStats.error > 0 ? row('Errors', n(d.leadsStats.error), '#dc2626') : '') +
      '</table>' +
    '</td></tr>' +

    // Campaigns
    sectionTitle('Campaigns') +
    '<tr><td style="padding:0 24px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
      '<tr>' + th('Campaign') + th('Pending', 'right') + th('Queued', 'right') +
                 th('Sent', 'right') + th('Replied', 'right') + th('Reply %', 'right') +
                 th('Done', 'right') + '</tr>' +
        campRows +
      '</table>' +
    '</td></tr>' +

    // All time
    sectionTitle('All time') +
    '<tr><td style="padding:0 24px;">' +
      '<table role="presentation" width="100%" cellpadding="0" cellspacing="0">' +
        row('Emails sent', n(d.senderStats.totalSentCount)) +
        row('Replies', n(d.senderStats.repliesReceived), '#16a34a') +
        row('Bounces', n(d.senderStats.bouncedEmails),
            d.senderStats.bouncedEmails > 0 ? '#dc2626' : null) +
      '</table>' +
    '</td></tr>' +

    // Footer
    '<tr><td style="padding:24px;">' +
      '<div style="border-top:1px solid #eef1f6;padding-top:16px;' +
           'font:400 12px/1.6 Arial,sans-serif;color:#94a3b8;">' +
        'Generated automatically by Reach X. Reply rate and bounce rate are ' +
        'measured against total leads.' +
      '</div>' +
    '</td></tr>' +

  '</table>' +
'</td></tr></table></body></html>';
}
