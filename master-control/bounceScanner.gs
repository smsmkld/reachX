// ============================================================================
// bounceScanner.gs — Cold Email Automation System
// Location: Master Control Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Scans Gmail inboxes of all active senders for bounce notifications.
// Runs at 11:00 PM daily — before runDailyReset() at 11:55 PM.
//
// Flow:
//  1. Read active senders from senderAccounts sheet
//  2. Fire parallel webhook requests to all sender webApps (scanBounces action)
//  3. Each sender scans their own Gmail for Mail Delivery Subsystem emails
//  4. Collect all bounced email addresses from responses
//  5. Match against masterLeads by email address
//  6. Batch update masterLeads: bounceStatus='Bounced', bounceDate=today
//  7. Write all updates in ONE setValues() call
//
// Dependencies:
//   utils.gs must be present in Master Control's Apps Script project.
// ============================================================================


// ============================================================================
// SECTION 1 — CONSTANTS
// ============================================================================
const BOUNCE_SCAN_DAYS    = 30;           // HARDCODED — verify this
const BOUNCE_LEADS_SHEET  = 'masterLeads'; // HARDCODED — verify this
const BOUNCE_SENDER_SHEET = 'senderAccounts'; // HARDCODED — verify this
const BOUNCE_FETCH_BATCH  = 50;           // HARDCODED — verify this


// ============================================================================
// SECTION 2 — scanAllSendersForBounces()
// Main entry point. Triggered at 11:00 PM daily from Master Control.
// ============================================================================
/**
 * Scans all active sender Gmail accounts for bounce notifications.
 * Triggered at 11:00 PM daily from Master Control spreadsheet.
 */
function scanAllSendersForBounces() {
  log('scanAllSendersForBounces: started', 'INFO');

  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    log('scanAllSendersForBounces: could not acquire lock — exiting', 'WARN');
    return;
  }

  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    if (!leadsSheetId) {
      log('scanAllSendersForBounces: leadsSpreadsheetId not in Settings — exiting', 'ERROR');
      return;
    }

    // ── Step 1: Read active senders ─────────────────────────────────────────
    const senders = getActiveSendersForScan(); // reuses function from replyScanner.gs
    if (senders.length === 0) {
      log('scanAllSendersForBounces: no active senders — exiting', 'INFO');
      return;
    }
    log('scanAllSendersForBounces: ' + senders.length + ' active senders found', 'INFO');

    // ── Step 2: Build all fetch requests ────────────────────────────────────
    const fetchRequests = [];
    const fetchMeta     = [];

    for (let i = 0; i < senders.length; i++) {
      const sender = senders[i];
      if (!sender.webAppUrl) { continue; }

      fetchRequests.push({
        url:                sender.webAppUrl,
        method:             'post',
        contentType:        'application/json',
        payload:            JSON.stringify({
          action:      'scanBounces',
          senderEmail: sender.emailAddress
        }),
        muteHttpExceptions: true
      });

      fetchMeta.push({
        emailID:      sender.emailID,
        emailAddress: sender.emailAddress
      });
    }

    // ── Step 3: Fire all requests in parallel batches of 50 ─────────────────
    const allBouncedEmails = []; // will collect { email, bounceDate } objects

    for (let b = 0; b < fetchRequests.length; b += BOUNCE_FETCH_BATCH) {
      const batchRequests = fetchRequests.slice(b, b + BOUNCE_FETCH_BATCH);
      const batchMeta     = fetchMeta.slice(b, b + BOUNCE_FETCH_BATCH);

      let responses = [];
      try {
        responses = UrlFetchApp.fetchAll(batchRequests);
      } catch (e) {
        log('scanAllSendersForBounces: fetchAll error on batch ' + Math.floor(b / BOUNCE_FETCH_BATCH + 1) + ' — ' + e.message, 'ERROR');
        continue;
      }

      for (let r = 0; r < responses.length; r++) {
        const meta = batchMeta[r];
        try {
          const code = responses[r].getResponseCode();
          if (code !== 200) {
            log('scanAllSendersForBounces: non-200 (' + code + ') from sender ' + meta.emailID, 'ERROR');
            continue;
          }
          const result         = JSON.parse(responses[r].getContentText());
          const senderBounces  = result.bounces || [];
          log('scanAllSendersForBounces: sender ' + meta.emailID + ' — ' + senderBounces.length + ' bounces found', 'INFO');

          for (let j = 0; j < senderBounces.length; j++) {
            allBouncedEmails.push({
              email:       senderBounces[j].email,
              bounceDate:  senderBounces[j].bounceDate,
              senderEmail: meta.emailAddress
            });
          }
        } catch (e) {
          log('scanAllSendersForBounces: error parsing response from ' + meta.emailID + ' — ' + e.message, 'ERROR');
        }
      }
    }

    log('scanAllSendersForBounces: total bounced emails collected: ' + allBouncedEmails.length, 'INFO');

    // ── Step 4: Write bounce updates to masterLeads ──────────────────────────
    if (allBouncedEmails.length > 0) {
      writeBounceUpdates_(allBouncedEmails, leadsSheetId);
    } else {
      log('scanAllSendersForBounces: no bounces found — nothing to write', 'INFO');
    }

    log('scanAllSendersForBounces: completed successfully', 'INFO');

  } catch (e) {
    log('scanAllSendersForBounces: unexpected error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}


// ============================================================================
// SECTION 3 — writeBounceUpdates_()
// Reads masterLeads once, matches bounced emails, writes all updates in ONE
// setValues() call.
// ============================================================================
/**
 * Matches bounced email addresses against masterLeads and marks them.
 *
 * @param {Array}  bouncedEmails - Array of { email, bounceDate }
 * @param {string} leadsSheetId  - Google Sheets ID of the Leads Database.
 */
function writeBounceUpdates_(bouncedEmails, leadsSheetId) {
  if (!acquireLock(LOCK_WAIT_SECONDS)) {
    log('writeBounceUpdates_: could not acquire lock — updates not written', 'ERROR');
    return;
  }

  try {
    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName(BOUNCE_LEADS_SHEET);
    if (!sheet) {
      log('writeBounceUpdates_: masterLeads sheet not found', 'ERROR');
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) {
      log('writeBounceUpdates_: no lead rows found', 'INFO');
      return;
    }

    // Read entire sheet in ONE call
    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    // Check required columns exist
    if (!('bounceStatus' in col)) {
      log('writeBounceUpdates_: bounceStatus column not found in masterLeads — please add it', 'ERROR');
      return;
    }
    if (!('bounceDate' in col)) {
      log('writeBounceUpdates_: bounceDate column not found in masterLeads — please add it', 'ERROR');
      return;
    }

    // Build email → dataIndex map for O(1) lookups
    // One email could appear multiple times (multiple sequence steps) — mark all
    const emailMap = {};
    for (let i = 0; i < data.length; i++) {
      const email = String(data[i][col['email']] || '').trim().toLowerCase();
      if (!email) { continue; }
      if (!emailMap[email]) { emailMap[email] = []; }
      emailMap[email].push(i);
    }

    // Build set of bounced emails for O(1) lookup
    const bouncedMap = {};
    const today      = formatDate(new Date());
    for (let b = 0; b < bouncedEmails.length; b++) {
      const email = String(bouncedEmails[b].email || '').trim().toLowerCase();
      if (email) {
        bouncedMap[email] = bouncedEmails[b].bounceDate || today;
      }
    }

    let updateCount = 0;

    // Match and update
    const bouncedEmailList = Object.keys(bouncedMap);
    for (let b = 0; b < bouncedEmailList.length; b++) {
      const email      = bouncedEmailList[b];
      const bounceDate = bouncedMap[email];
      const indices    = emailMap[email];
      if (!indices) { continue; } // email not in masterLeads

      for (let i = 0; i < indices.length; i++) {
        const idx = indices[i];

        // Skip if already marked
        const existing = String(data[idx][col['bounceStatus']] || '').trim();
        if (existing === 'Bounced') { continue; }

        data[idx][col['bounceStatus']] = 'Bounced';
        data[idx][col['bounceDate']]   = bounceDate;

        // Also mark status as Bounced so distributor skips it
        if ('status' in col) {
          data[idx][col['status']] = 'Bounced';
        }

        updateCount++;
      }
    }

    // Write ALL changes in ONE setValues() call
    if (updateCount > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
      log('writeBounceUpdates_: marked ' + updateCount + ' leads as Bounced in masterLeads', 'INFO');
      updateSenderAccountsBounceCounts_(bouncedEmails);
    } else {
      log('writeBounceUpdates_: all matched emails already marked Bounced', 'INFO');
    }

  } catch (e) {
    log('writeBounceUpdates_: error — ' + e.message, 'ERROR');
  } finally {
    releaseLock();
  }
}


// ============================================================================
// SECTION 4 — setupBounceScannerTrigger()
// One-time setup. Creates the 11:00 PM daily trigger.
// Run manually once from Master Control Apps Script editor.
// ============================================================================
/**
 * Creates the daily 11:00 PM trigger for scanAllSendersForBounces().
 * Run ONCE manually from Master Control Apps Script editor.
 * Safe to re-run — deletes existing trigger first.
 */
function setupBounceScannerTrigger() {
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'scanAllSendersForBounces') {
      ScriptApp.deleteTrigger(triggers[i]);
      log('setupBounceScannerTrigger: deleted existing trigger', 'INFO');
    }
  }

  ScriptApp.newTrigger('scanAllSendersForBounces')
    .timeBased()
    .everyDays(1)
    .atHour(23)
    .nearMinute(0)
    .inTimezone('Africa/Cairo') // HARDCODED — verify this
    .create();

  log('setupBounceScannerTrigger: created 11:00 PM Cairo daily trigger', 'INFO');
  Logger.log('✅ Bounce scanner trigger created — runs at 11:00 PM Cairo time daily.');
}

function updateSenderAccountsBounceCounts_(bouncedEmails) {
  try {
    const ss    = SpreadsheetApp.openById(MASTER_SHEET_ID);
    const sheet = ss.getSheetByName(BOUNCE_SENDER_SHEET);
    if (!sheet || sheet.getLastRow() < 2) { return; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1,
                                   sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    const bouncesPerSender = {};
    bouncedEmails.forEach(function(bounce) {
      const sender = String(bounce.senderEmail || '').toLowerCase();
      if (!sender) { return; }
      bouncesPerSender[sender] = (bouncesPerSender[sender] || 0) + 1;
    });

    for (let i = 0; i < data.length; i++) {
      const email = String(data[i][col['emailAddress']] || '').trim().toLowerCase();
      if (!bouncesPerSender[email]) { continue; }
      if ('bouncedEmails' in col) {
        data[i][col['bouncedEmails']] =
          (parseInt(data[i][col['bouncedEmails']] || 0, 10)) +
          bouncesPerSender[email];
      }
    }

    sheet.getRange(2, 1, data.length, headers.length).setValues(data);
    SpreadsheetApp.flush();
    log('updateSenderAccountsBounceCounts_: updated bouncedEmails in senderAccounts', 'INFO');
  } catch (e) {
    log('updateSenderAccountsBounceCounts_: error — ' + e.message, 'ERROR');
  }
}
