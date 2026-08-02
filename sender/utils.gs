// ============================================================================
// utils.gs — Cold Email Automation System
// Location: Master Control Spreadsheet → Apps Script editor
// ----------------------------------------------------------------------------
// Pure helper functions. No triggers. No sheet-specific logic.
// Every other file in this system imports or copies from here.
// Build order: utils → distributor → sender → followup → replyScanner
//              → campaignSync → resetDaily → webApp → index.html
// ============================================================================

// ============================================================================
// SECTION 1 — sanitize()
// Converts any value to a clean string safe for JSON storage and email sending.
// Removes newlines, trims whitespace, converts double-quotes to single-quotes.
// Called on every lead field before it is written to the inbox or DB.
// ============================================================================
/**
 * Sanitizes a value for safe storage in JSON and email sending.
 *
 * @param {*} value - Any value (string, number, null, undefined, etc.)
 * @returns {string} Cleaned string. Empty string for null/undefined/empty input.
 *
 * Rules applied in order:
 *  1. null / undefined → return ''
 *  2. Convert to string
 *  3. Replace \r and \n (and \r\n) with a single space
 *  4. Replace all double-quote characters with single-quote
 *     (prevents JSON.parse failures when value is stored in a JSON string)
 *  5. Trim leading/trailing whitespace
 */
function sanitize(value) {
  if (value === null || value === undefined || value === '') {
    return '';
  }
  let str = String(value);
  str = str.replace(/\r\n/g, ' ').replace(/[\r\n]/g, ' ');
  str = str.replace(/"/g, "'");
  str = str.trim();
  return str;
}

// ============================================================================
// SECTION 2 — processSpintax()
// Resolves all {{random|option1|option2|...}} blocks in a string.
// Each block is replaced with one randomly chosen option.
// Must be called BEFORE replaceVariables() so that any {{variable}}
// placeholders inside a chosen spintax option are resolved in the
// second pass.
//
// Format:  {{random|option one|option two|option three}}
//  - The word 'random' is case-insensitive ({{RANDOM|...}} works too)
//  - Options are separated by | (pipe)
//  - Any number of options is supported (minimum 1)
//  - Whitespace around each option is trimmed
//  - Multiple spintax blocks in one string are each resolved independently
//  - {{variable}} placeholders inside options are left intact here —
//    they will be replaced by replaceVariables() in the next pass
//
// Example:
//   Input:  "Hi {{firstName}}, {{random|hope you're well|trust this finds you well}}"
//   After processSpintax():  "Hi {{firstName}}, trust this finds you well"
//   After replaceVariables(): "Hi John, trust this finds you well"
// ============================================================================
/**
 * Resolves all spintax blocks in a string, picking one option at random.
 *
 * @param {string} text - Raw template string that may contain spintax blocks.
 * @returns {string} String with all spintax blocks replaced by chosen options.
 *
 * Notes:
 *  - Safe to call on strings with no spintax — returns the string unchanged.
 *  - Does NOT resolve {{variable}} placeholders — call replaceVariables() after.
 *  - Each spintax block is resolved independently via its own Math.random() call.
 */
function processSpintax(text) {
  if (!text || typeof text !== 'string') {
    return text || '';
  }

  // Regex explanation:
  //   {{          — literal opening braces
  //   \s*         — optional whitespace before 'random'
  //   random      — the keyword (case-insensitive via /i flag)
  //   \s*         — optional whitespace after 'random'
  //   \|          — required pipe separator after 'random'
  //   (.*?)       — captured options string (non-greedy)
  //   }}          — literal closing braces
  // The /gi flags make it global (replace all) and case-insensitive on 'random'.
  return text.replace(/\[\[\s*random\s*\|(.*?)\]\]/gi, function(fullMatch, optionsStr) {
    // Split options by pipe and trim each one
    const options = optionsStr.split('|').map(function(opt) { return opt.trim(); });

    // Filter out any empty options that result from trailing pipes
    const validOptions = options.filter(function(opt) { return opt.length > 0; });

    if (validOptions.length === 0) {
      // No valid options — replace with empty string rather than leaving raw block
      log('processSpintax: spintax block had no valid options — replacing with empty string', 'WARN');
      return '';
    }

    // Pick one option at random
    // Use a more distributed random by hashing time + option count
    const chosenIndex = Math.floor(Math.random() * validOptions.length);
    return validOptions[chosenIndex];
  });
}

// ============================================================================
// SECTION 3 — replaceVariables()
// Replaces all {{key}} placeholders in an email template string.
// Calls processSpintax() FIRST to resolve any spintax blocks before
// variable replacement — this ensures variables inside chosen spintax
// options are correctly substituted in the same pass.
//
// Lookup order: leadObj first → senderObj second → '' if not found.
// Missing keys are replaced with empty string — never left as {{placeholder}}.
// ============================================================================
/**
 * Resolves spintax then replaces all {{key}} placeholders in a template string.
 *
 * Processing order (important — do not change):
 *  1. processSpintax(template) — resolves {{random|...}} blocks first
 *  2. Replace all remaining {{key}} placeholders using leadObj then senderObj
 *
 * Supported variables (case-sensitive):
 *   Lead fields:   {{firstName}}, {{lastName}}, {{companyName}},
 *                  {{companyWebsite}}, {{jobTitle}}, {{industry}},
 *                  {{customVar1}} through {{customVar5}}
 *   Sender fields: {{senderFirstName}}, {{senderLastName}}
 *
 * @param {string} template   - The email subject or body containing {{placeholders}}.
 * @param {Object} leadObj    - Plain JS object of lead fields (parsed leadData JSON).
 * @param {Object} senderObj  - Plain JS object of sender fields from Config sheet.
 * @returns {string} Fully resolved string with all spintax and variables replaced.
 *
 * Example:
 *   template:  "Hi {{firstName}}, {{random|hope you're well|trust this finds you well}}"
 *   leadObj:   { firstName: 'John' }
 *   returns:   "Hi John, trust this finds you well"   (or the other option)
 */
function replaceVariables(template, leadObj, senderObj) {
  if (!template || typeof template !== 'string') {
    return '';
  }

  const lead   = leadObj   || {};
  const sender = senderObj || {};

  // ── Step 1: Resolve all spintax blocks FIRST ─────────────────────────────
  // This must happen before variable replacement so that any {{variable}}
  // placeholders inside a chosen spintax option are caught in step 2.
  let resolved = processSpintax(template);

  // ── Step 2: Replace all remaining {{key}} placeholders ───────────────────
  // /{{(.*?)}}/g — non-greedy match, global, captures key name
  resolved = resolved.replace(/{{(.*?)}}/g, function(fullMatch, key) {
    const trimmedKey = key.trim();

    // Check lead object first
    if (trimmedKey in lead && lead[trimmedKey] !== undefined && lead[trimmedKey] !== null) {
      return String(lead[trimmedKey]);
    }
    // Fall back to sender object
    if (trimmedKey in sender && sender[trimmedKey] !== undefined && sender[trimmedKey] !== null) {
      return String(sender[trimmedKey]);
    }
    // Key not found — replace with empty string
    // A sent email must NEVER contain a literal {{placeholder}}
    return '';
  });

  return resolved;
}

// ============================================================================
// SECTION 4 — getSettings()
// Reads the Settings sheet from Master Control into a plain JS object.
// Always called with the Master Control spreadsheet ID as argument.
// Returns { settingName: value } — all strings, caller casts if needed.
// ============================================================================
/**
 * Reads global settings from the Settings sheet in Master Control.
 *
 * @param {string} masterSheetId - Google Sheets ID of the Master Control spreadsheet.
 * @returns {Object} Plain JS object: { settingName: value, ... }
 */
function getSettings(masterSheetId) {
  const ss    = SpreadsheetApp.openById(masterSheetId);
  const sheet = ss.getSheetByName('Settings');
  if (!sheet) {
    log('getSettings: Settings sheet not found in spreadsheet ' + masterSheetId, 'ERROR');
    return {};
  }
  const lastRow = sheet.getLastRow();
  if (lastRow < 1) {
    log('getSettings: Settings sheet is empty', 'WARN');
    return {};
  }
  const rawData  = sheet.getRange(1, 1, lastRow, 2).getValues();
  const settings = {};
  for (let i = 0; i < rawData.length; i++) {
    const key   = String(rawData[i][0]).trim();
    const value = String(rawData[i][1]).trim();
    if (key === '') { continue; }
    settings[key] = value;
  }
  return settings;
}

// ============================================================================
// SECTION 5 — getConfig()
// Reads a sender's Config sheet into a plain JS object.
// Takes an already-opened Sheet object (NOT a spreadsheet ID).
// ============================================================================
/**
 * Reads a sender's Config sheet into a plain JS object.
 *
 * @param {GoogleAppsScript.Spreadsheet.Sheet} configSheet
 * @returns {Object} Plain JS object: { configKey: value, ... }
 */
function getConfig(configSheet) {
  if (!configSheet) {
    log('getConfig: configSheet argument is null or undefined', 'ERROR');
    return {};
  }
  const lastRow = configSheet.getLastRow();
  if (lastRow < 1) {
    log('getConfig: Config sheet is empty', 'WARN');
    return {};
  }
  const rawData = configSheet.getRange(1, 1, lastRow, 2).getValues();
  const config  = {};
  for (let i = 0; i < rawData.length; i++) {
    const key   = String(rawData[i][0]).trim();
    const value = String(rawData[i][1]).trim();
    if (key === '') { continue; }
    config[key] = value;
  }
  return config;
}

// ============================================================================
// SECTION 6 — buildLeadDataJson()
// Zips a lead row array with a header row array into a sanitized JSON string.
// Stored in the leadData column of dailyInbox and followupInbox.
// ============================================================================
/**
 * Builds a sanitized JSON string from a single lead row and its header row.
 *
 * @param {Array} leadRow   - One row from masterLeads sheet.
 * @param {Array} headerRow - The header row from masterLeads.
 * @returns {string} JSON string of all lead fields, all values sanitized.
 */
function buildLeadDataJson(leadRow, headerRow) {
  if (!leadRow || !headerRow || leadRow.length === 0 || headerRow.length === 0) {
    log('buildLeadDataJson: empty leadRow or headerRow received', 'WARN');
    return '{}';
  }
  const leadObj = {};
  const len = Math.min(leadRow.length, headerRow.length);
  for (let i = 0; i < len; i++) {
    const key   = String(headerRow[i]).trim();
    const value = sanitize(leadRow[i]);
    if (key === '') { continue; }
    leadObj[key] = value;
  }
  return JSON.stringify(leadObj);
}

// ============================================================================
// SECTION 7 — formatDate() and addDays()
// Date utility functions. All dates stored as 'YYYY-MM-DD' strings.
// ============================================================================
/**
 * Formats a JavaScript Date object as a 'YYYY-MM-DD' string.
 *
 * @param {Date} date
 * @returns {string} 'YYYY-MM-DD' or '' on invalid input.
 */
function formatDate(date) {
  if (!(date instanceof Date) || isNaN(date.getTime())) {
    log('formatDate: invalid Date object received', 'WARN');
    return '';
  }
  const year  = date.getFullYear();
  const month = String(date.getMonth() + 1).padStart(2, '0');
  const day   = String(date.getDate()).padStart(2, '0');
  return year + '-' + month + '-' + day;
}

/**
 * Adds N days to a 'YYYY-MM-DD' string and returns a new 'YYYY-MM-DD' string.
 *
 * @param {string} dateStr - Date in 'YYYY-MM-DD' format.
 * @param {number} days    - Integer days to add (can be 0 or negative).
 * @returns {string} New date in 'YYYY-MM-DD' format.
 *
 * Uses noon (12:00:00) to avoid DST boundary bugs where midnight on a
 * changeover day can roll back to the previous date.
 */
function addDays(dateStr, days) {
  if (!dateStr || typeof dateStr !== 'string') {
    log('addDays: invalid dateStr received: ' + dateStr, 'WARN');
    return '';
  }
  const parts = dateStr.split('-');
  if (parts.length !== 3) {
    log('addDays: dateStr is not in YYYY-MM-DD format: ' + dateStr, 'WARN');
    return '';
  }
  const year  = parseInt(parts[0], 10);
  const month = parseInt(parts[1], 10) - 1;
  const day   = parseInt(parts[2], 10);
  const date  = new Date(year, month, day, 12, 0, 0);
  date.setDate(date.getDate() + days);
  return formatDate(date);
}

// ============================================================================
// SECTION 8 — acquireLock() and releaseLock()
// Wrapper around Apps Script LockService.
// Every function that writes to a shared sheet MUST call acquireLock() first.
// ============================================================================
/**
 * Acquires a script-level lock, waiting up to waitSeconds for it.
 *
 * @param {number} waitSeconds - Max seconds to wait. Default: 10.
 * @returns {boolean} true if acquired, false if timeout (caller must exit).
 */
function acquireLock(waitSeconds) {
  const wait = (waitSeconds || 10) * 1000;
  try {
    const lock = LockService.getScriptLock();
    lock.waitLock(wait);
    return true;
  } catch (e) {
    log('acquireLock: could not acquire lock within ' + waitSeconds + 's — ' + e.message, 'WARN');
    return false;
  }
}

/**
 * Releases the script-level lock acquired by acquireLock().
 * Always wrap in try/catch — release failures are non-fatal.
 */
function releaseLock() {
  try {
    LockService.getScriptLock().releaseLock();
  } catch (e) {
    log('releaseLock: error releasing lock: ' + e.message, 'WARN');
  }
}

// ============================================================================
// SECTION 9 — log()
// Lightweight logger wrapper around Apps Script Logger.log().
// Adds ISO timestamp and level prefix to every message.
// ============================================================================
/**
 * Logs a message with timestamp and level to the Apps Script Execution Log.
 *
 * @param {string} message - The message to log.
 * @param {string} level   - 'INFO', 'WARN', or 'ERROR'. Default: 'INFO'.
 *
 * Output format:
 *   [INFO]  2026-04-26T07:00:01.234Z — distributeCold: started
 *   [ERROR] 2026-04-26T07:01:10.512Z — sendColdEmail: GmailApp quota exceeded
 */
function log(message, level) {
  const lvl       = (level || 'INFO').toUpperCase();
  const timestamp = new Date().toISOString();
  Logger.log('[' + lvl + '] ' + timestamp + ' — ' + message);
}

/**
 * Reads the Campaigns sheet and returns a map of campaignId → campaign settings.
 * Only reads from rows where sequenceStep = 1 (settings live on step 1 row only).
 *
 * @param {string} masterSheetId
 * @returns {Object} { campaignId: { assignedSenders, windowStart, windowEnd,
 *                    sendingDays, timezone, coldPriorityLimit,
 *                    followupPriorityLimit, campaignStatus } }
 */
function getCampaignSettings(masterSheetId) {
  const result = {};
  try {
    const ss    = SpreadsheetApp.openById(masterSheetId);
    const sheet = ss.getSheetByName('Campaigns');
    if (!sheet || sheet.getLastRow() < 2) { return result; }

    const headers = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
    const data    = sheet.getRange(2, 1, sheet.getLastRow() - 1, sheet.getLastColumn()).getValues();
    const col     = {};
    headers.forEach(function(h, i) { col[String(h).trim()] = i; });

    for (let i = 0; i < data.length; i++) {
      const step       = parseInt(data[i][col['sequenceStep']] || 0, 10);
      const campaignId = String(data[i][col['campaignId']]    || '').trim();
      if (step !== 1 || !campaignId) { continue; } // only step 1 rows

      // Parse assignedSenders — comma-separated like "S001,S002,S003"
      const assignedRaw = String(data[i][col['assignedSenders']] || '').trim();
      const assignedSenders = assignedRaw
        ? assignedRaw.split(',').map(function(s) { return s.trim(); }).filter(Boolean)
        : [];

      result[campaignId] = {
        assignedSenders:      assignedSenders,
        windowStart: formatTimeValue_(data[i][col['windowStart']]),
        totalSequenceSteps: parseInt(data[i][col['totalSequenceSteps']] || 0, 10),
        windowEnd:   formatTimeValue_(data[i][col['windowEnd']]),
        sendingDays:          String(data[i][col['sendingDays']]          || '').trim(),
        timezone:             String(data[i][col['timezone']]             || '').trim(),
        coldPriorityLimit:    parseFloat(data[i][col['coldPriorityLimit']]    || 0) || null,
        followupPriorityLimit:parseFloat(data[i][col['followupPriorityLimit']]|| 0) || null,
        campaignStatus:       String(data[i][col['campaignStatus']]       || '').trim()
      };
    }
    log('getCampaignSettings: loaded settings for ' + Object.keys(result).length + ' campaigns', 'INFO');
  } catch (e) {
    log('getCampaignSettings: error — ' + e.message, 'ERROR');
  }
  return result;
}

/**
 * Converts a Sheets time value to HH:mm string.
 * Handles both Date objects (Sheets auto-converts HH:mm to 1899 dates)
 * and plain strings like '08:00'.
 *
 * @param {*} value - Raw cell value from Sheets.
 * @returns {string} 'HH:mm' string or empty string.
 */
function formatTimeValue_(value) {
  if (!value && value !== 0) { return ''; }
  if (value instanceof Date) {
    const h = String(value.getHours()).padStart(2, '0');
    const m = String(value.getMinutes()).padStart(2, '0');
    return h + ':' + m;
  }
  return String(value).trim();
}
