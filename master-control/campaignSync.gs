// ============================================================================
// campaignSync.gs — Cold Email Automation System
// Location: Master Control Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Handles campaign status changes and cascades them to masterLeads.
// ============================================================================

/**
 * onEdit trigger. Detects when campaignStatus column is changed in Campaigns sheet.
 * Enforces a 3-minute cooldown to prevent rapid duplicate cascades.
 */
function onCampaignEdit(e) {
  if (!e || !e.range) { return; }

  const sheet = e.range.getSheet();
  if (sheet.getName() !== 'Campaigns') { return; }

  // Read headers to find which column was edited
  const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  const col = {};
  headers.forEach(function(h, i) { col[String(h).trim()] = i; });

  const editedCol = e.range.getColumn() - 1; // Convert to 0-based
  if (editedCol !== col['campaignStatus']) { return; }

  const editedRow  = e.range.getRow();
  if (editedRow < 2) { return; } // Ignore header row

  // Read the full row to get campaignId and lastModified
  const rowData    = sheet.getRange(editedRow, 1, 1, sheet.getLastColumn()).getValues()[0];
  const campaignId = String(rowData[col['campaignId']]   || '').trim();
  const newStatus  = String(rowData[col['campaignStatus']] || '').trim();

  if (!campaignId) {
    log('onCampaignEdit: no campaignId found on edited row — exiting', 'WARN');
    return;
  }

  // --- 3-minute cooldown check ---
  const lastModifiedRaw = rowData[col['lastModified']];
  if (lastModifiedRaw) {
    const lastModified = new Date(lastModifiedRaw);
    const gapMinutes   = (new Date().getTime() - lastModified.getTime()) / (1000 * 60);
    if (gapMinutes < 3) {
      log('onCampaignEdit: cooldown active for campaign ' + campaignId + ' — skipping', 'INFO');
      return;
    }
  }

  // Update lastModified timestamp on the edited row
  sheet.getRange(editedRow, col['lastModified'] + 1).setValue(new Date());

  log('onCampaignEdit: cascading status "' + newStatus + '" for campaign ' + campaignId, 'INFO');
  cascadeCampaignStatus(campaignId, newStatus);
}

/**
 * Updates campaignStatus for all leads in masterLeads that belong to this campaign.
 * Single batch read + write — never row by row.
 *
 * @param {string} campaignId - The campaign ID to update.
 * @param {string} newStatus  - The new status value to set.
 */
function cascadeCampaignStatus(campaignId, newStatus) {
  try {
    const settings     = getSettings(MASTER_SHEET_ID);
    const leadsSheetId = settings.leadsSpreadsheetId;
    if (!leadsSheetId) {
      log('cascadeCampaignStatus: leadsSpreadsheetId not in Settings — exiting', 'ERROR');
      return;
    }

    const ss    = SpreadsheetApp.openById(leadsSheetId);
    const sheet = ss.getSheetByName('masterLeads');
    if (!sheet) {
      log('cascadeCampaignStatus: masterLeads sheet not found', 'ERROR');
      return;
    }

    const lastRow = sheet.getLastRow();
    if (lastRow < 2) { return; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, lastRow - 1, sheet.getLastColumn()).getValues();

    const col = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    let updateCount = 0;
    for (let i = 0; i < data.length; i++) {
      const rowCampaignId = String(data[i][col['campaignId']] || '').trim();
      if (rowCampaignId !== campaignId) { continue; }
      data[i][col['campaignStatus']] = newStatus;
      updateCount++;
    }

    if (updateCount > 0) {
      sheet.getRange(2, 1, data.length, headers.length).setValues(data);
      SpreadsheetApp.flush();
      log('cascadeCampaignStatus: updated ' + updateCount + ' leads to status "' + newStatus + '"', 'INFO');
    } else {
      log('cascadeCampaignStatus: no leads found for campaignId ' + campaignId, 'INFO');
    }
  } catch (e) {
    log('cascadeCampaignStatus: error — ' + e.message, 'ERROR');
  }
}

function syncCampaignStats(masterSheetId, leadsSheetId) {
  try {
    // Read masterLeads once
    const leadsSS    = SpreadsheetApp.openById(leadsSheetId);
    const leadsSheet = leadsSS.getSheetByName('masterLeads');
    if (!leadsSheet || leadsSheet.getLastRow() < 2) { return; }

    const leadsHeaders = leadsSheet.getRange(1, 1, 1, leadsSheet.getLastColumn()).getValues()[0];
    const leadsData    = leadsSheet.getRange(2, 1, leadsSheet.getLastRow() - 1, leadsSheet.getLastColumn()).getValues();
    const lCol = {};
    leadsHeaders.forEach(function(h, i) { lCol[String(h).trim()] = i; });

    // Aggregate counts per campaignId
    const stats = {};
    for (let i = 0; i < leadsData.length; i++) {
      const cId         = String(leadsData[i][lCol['campaignId']]   || '').trim();
      const status      = String(leadsData[i][lCol['status']]       || '').trim();
      const replyStatus = String(leadsData[i][lCol['replyStatus']]  || '').trim();
      const bounceStatus = String(leadsData[i][lCol['bounceStatus']] || '').trim();

      if (!stats[cId]) {
        stats[cId] = { totalLeads: 0, sentCount: 0, replyCount: 0, pendingCount: 0, bouncedEmails: 0 };
      }
      stats[cId].totalLeads++;
      if (status === 'Sent' || status === 'Completed')  { stats[cId].sentCount++;    }
      if (replyStatus === 'Replied')                    { stats[cId].replyCount++;   }
      if (status === 'Pending')                         { stats[cId].pendingCount++; }
      if (bounceStatus === 'Bounced')                   { stats[cId].bouncedEmails++;  }
    }

    // Read Campaigns sheet once
    const masterSS      = SpreadsheetApp.openById(masterSheetId);
    const campSheet     = masterSS.getSheetByName('Campaigns');
    if (!campSheet || campSheet.getLastRow() < 2) { return; }

    const campHeaders = campSheet.getRange(1, 1, 1, campSheet.getLastColumn()).getValues()[0];
    const campData    = campSheet.getRange(2, 1, campSheet.getLastRow() - 1, campSheet.getLastColumn()).getValues();
    const cCol = {};
    campHeaders.forEach(function(h, i) { cCol[String(h).trim()] = i; });

    const today = formatDate(new Date());

    for (let i = 0; i < campData.length; i++) {
      const cId = String(campData[i][cCol['campaignId']] || '').trim();
      if (!cId || !stats[cId]) { continue; }

      const s = stats[cId];
      if ('totalLeads'   in cCol) { campData[i][cCol['totalLeads']]   = s.totalLeads;   }
      if ('sentCount'    in cCol) { campData[i][cCol['sentCount']]    = s.sentCount;    }
      if ('replyCount'   in cCol) { campData[i][cCol['replyCount']]   = s.replyCount;   }
      if ('pendingCount' in cCol) { campData[i][cCol['pendingCount']] = s.pendingCount; }
      if ('bouncedEmails' in cCol) { campData[i][cCol['bouncedEmails']] = s.bouncedEmails; }
      if ('lastModified' in cCol) { campData[i][cCol['lastModified']] = today;          }

      // Set createdDate only if it is currently empty
      if ('createdDate' in cCol) {
        const existing = String(campData[i][cCol['createdDate']] || '').trim();
        if (!existing) { campData[i][cCol['createdDate']] = today; }
      }
    }

    // Write back in ONE call
    campSheet.getRange(2, 1, campData.length, campHeaders.length).setValues(campData);
    SpreadsheetApp.flush();
    log('syncCampaignStats: updated stats for ' + Object.keys(stats).length + ' campaigns', 'INFO');

  } catch (e) {
    log('syncCampaignStats: error — ' + e.message, 'ERROR');
  }
}