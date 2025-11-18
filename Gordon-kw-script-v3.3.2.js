// ============================================================================
// RANKING FUNCTIONALITY
// ============================================================================

/**
 * Get DataForSEO API configuration
 */
function getDataForSEOConfig() {
  const basicAuth = PropertiesService.getScriptProperties().getProperty('basic');
  
  if (!basicAuth) {
    throw new Error('DataForSEO API key not found! Please add "basic" property in Project Settings > Script Properties');
  }
  
  return {
    Authorization: `Basic ${basicAuth}`,
    baseUrl: 'https://api.dataforseo.com/v3/serp/google/organic',
    taskPostUrl: 'https://api.dataforseo.com/v3/serp/google/organic/task_post',
    taskGetUrl: 'https://api.dataforseo.com/v3/serp/google/organic/task_get/regular'
  };
}

/**
 * Submit a single keyword ranking job to DataForSEO
 */
function submitKeywordRankingJob(keyword, lat, lng) {
  try {
    const config = getDataForSEOConfig();
    
    // Prepare POST data for desktop ranking
    const postData = [{
      "keyword": keyword,
      "location_coordinate": `${lat},${lng}`,
      "language_code": "en",
      "device": "desktop",
      "os": "windows",
      "depth": 30
    }];
    
    console.log(`📤 Submitting ranking job for: "${keyword}" in Location (3 pages)`);
    
    // Make API call
    const response = UrlFetchApp.fetch(config.taskPostUrl, {
      method: 'POST',
      headers: {
        'Authorization': config.Authorization,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(postData)
    });
    
    const responseData = JSON.parse(response.getContentText());
    const taskId = responseData.tasks[0].id;
    
    console.log(`✅ Ranking job submitted successfully: ${taskId}`);
    return taskId;
    
  } catch (error) {
    console.error(`❌ Failed to submit ranking job for "${keyword}":`, error);
    return null;
  }
}

/**
 * Submit a batch of keyword ranking jobs to DataForSEO (up to 500 keywords)
 */
function submitBatchRankingJobs(keywordsData) {
  try {
    const config = getDataForSEOConfig();
    
    // Prepare POST data for batch ranking
    const postData = keywordsData.map(item => ({
      "keyword": item.keyword,
      "location_coordinate": `${item.lat},${item.lng}`,
      "language_code": "en",
      "device": "desktop",
      "os": "windows",
      "depth": 30
    }));
    
    console.log(`📤 Submitting batch of ${keywordsData.length} keywords to DataForSEO (3 pages per keyword)`);
    
    // Make API call
    const response = UrlFetchApp.fetch(config.taskPostUrl, {
      method: 'POST',
      headers: {
        'Authorization': config.Authorization,
        'Content-Type': 'application/json'
      },
      payload: JSON.stringify(postData)
    });
    
    const responseData = JSON.parse(response.getContentText());
    
    if (responseData.tasks && responseData.tasks.length > 0) {
      const taskIds = responseData.tasks.map(task => task.id);
      console.log(`✅ Batch submitted successfully: ${taskIds.length} task IDs generated`);
      return taskIds;
    } else {
      throw new Error('No tasks returned from DataForSEO');
    }
    
  } catch (error) {
    console.error(`❌ Failed to submit batch ranking jobs:`, error);
    return [];
  }
}

/**
 * Fetch ranking results from DataForSEO
 */
function fetchKeywordRankingResults(taskId, keyword) {
  try {
    const config = getDataForSEOConfig();
    
    console.log(`📥 Fetching ranking results for task: ${keyword ? keyword + ' - ' : ''}${taskId}`);
    
    // Fetch results for this task
    const response = UrlFetchApp.fetch(
      `${config.taskGetUrl}/${taskId}`,
      {
        headers: { 'Authorization': config.Authorization }
      }
    );
    
    const responseData = JSON.parse(response.getContentText());
    const taskResult = responseData.tasks[0];
    
    if (taskResult.result && taskResult.result.length > 0) {
      // Extract ranking data for aaacwildliferemoval.com domain
      const rankings = extractKeywordRankingData(taskResult.result);
      
      console.log(`✅ Ranking results fetched: Found ${rankings.length} domain matches`);
      return {
        rankings: rankings,
        status: 'completed',
        rawData: taskResult
      };
    } else {
      console.log(`⚠️ No ranking results found for task: ${taskId}`);
      return {
        rankings: [],
        status: 'no_results',
        rawData: taskResult
      };
    }
    
  } catch (error) {
    console.error(`❌ Failed to fetch ranking results for task ${keyword ? keyword + ' - ' : ''}${taskId}:`, error);
    return {
      rankings: [],
      status: 'error',
      error: error.message
    };
  }
}

/**
 * Extract ranking data from SERP results, looking for aaacwildliferemoval.com domain
 */
function extractKeywordRankingData(serpResults) {
  const rankings = [];
  const targetDomain = 'aaacwildliferemoval.com';
  
  console.log(`🔍 Searching for ${targetDomain} in SERP results`);
  
  for (const result of serpResults) {
    if (result.items && Array.isArray(result.items)) {
      for (let i = 0; i < result.items.length; i++) {
        const item = result.items[i];
        
        if (item.type === "organic" && item.rank_group && item.url) {
          // Look for any URL containing the target domain
          if (item.url.includes(targetDomain)) {
            console.log(`✅ MATCH FOUND! Rank ${item.rank_group}: ${item.url}`);
            rankings.push({
              rank: item.rank_group,
              url: item.url
            });
          }
        }
      }
    }
  }
  
  console.log(`📈 Total domain matches found: ${rankings.length}`);
  return rankings;
}

/**
 * Check if a row already has ranking data
 */
function hasRankingData(sheet, row) {
  // Deprecated for v3.1 mapping (outputs are on ranking_results).
  // Keep signature for compatibility; determine keyword from kw_variants row
  // and check if it exists in ranking_results!D (Keyword column).
  const kwSheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || sheet;
  const keyword = String(kwSheet.getRange(row, 11).getValue() || '').trim(); // K
  if (!keyword) return false;
  return keywordAlreadyProcessed(keyword);
}

/**
 * Set up ranking headers
 */
function setupRankingHeaders(sheet) {
  // v3.2 writes to a dedicated results sheet. This function now ensures
  // headers exist in ranking_results A:C.
  const rs = getOrCreateRankingResultsSheet();
  rs.getRange(1, 1).setValue('Ranking Position'); // A
  rs.getRange(1, 2).setValue('Ranking URL');      // B
  rs.getRange(1, 3).setValue('Keyword'); // C (used for idempotent skip)
  // Basic styling for header row (optional)
  rs.getRange(1, 1, 1, 3).setFontWeight('bold');
}

/**
 * Write single ranking result to sheet
 * @param {Sheet} sheet - The kw_variants sheet (for compatibility)
 * @param {number} row - Original row number (used for inferring keyword if not provided)
 * @param {Object} result - Result object with ranking, url, keyword
 * @param {number} targetRow - Optional: specific row in ranking_results to write to (maintains sequence)
 */
function writeSingleRankingResult(sheet, row, result, targetRow) {
  // v3.2: Write to ranking_results with columns:
  // A: Position, B: URL, C: Keyword
  const rs = getOrCreateRankingResultsSheet();
  const position = result.ranking || result.position || 'Not Found';
  const url = result.url || 'Not Found';
  const kw = result.keyword || inferKeywordFromRow(row);
  
  // If targetRow is provided, write to that specific row (maintains sequence)
  // Otherwise, append to bottom (backward compatibility)
  let writeRow;
  if (targetRow && targetRow >= 2) {
    writeRow = targetRow;
    // Ensure ranking_results has enough rows (pad if needed)
    const rsLastRow = rs.getLastRow();
    if (writeRow > rsLastRow) {
      const rowsToAdd = writeRow - rsLastRow;
      if (rowsToAdd > 0) {
        rs.insertRowsAfter(rsLastRow, rowsToAdd);
      }
    }
  } else {
    // Backward compatibility: append to bottom
    writeRow = Math.max(2, rs.getLastRow() + 1);
  }
  
  rs.getRange(writeRow, 1, 1, 3).setValues([[position, url, kw]]);
}

// v3.2 helpers for new sheets and skip logic
function getOrCreateRankingResultsSheet() {
  const ss = SpreadsheetApp.getActive();
  let rs = ss.getSheetByName('ranking_results');
  if (!rs) {
    rs = ss.insertSheet('ranking_results');
    rs.getRange(1, 1).setValue('Ranking Position');
    rs.getRange(1, 2).setValue('Ranking URL');
    rs.getRange(1, 3).setValue('Keyword');
    rs.getRange(1, 1, 1, 3).setFontWeight('bold');
  }
  return rs;
}

function inferKeywordFromRow(row) {
  // Read keyword from kw_variants K for the given row
  const ss = SpreadsheetApp.getActive();
  const kwSheet = ss.getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  return String(kwSheet.getRange(row, 11).getValue() || '').trim();
}

function keywordAlreadyProcessed(keyword) {
  if (!keyword) return false;
  const rs = getOrCreateRankingResultsSheet();
  const lastRow = rs.getLastRow();
  if (lastRow < 2) return false;
  const keys = rs.getRange(2, 3, lastRow - 1, 1).getValues(); // C: Keyword
  for (let i = 0; i < keys.length; i++) {
    const v = String(keys[i][0] || '').trim();
    if (v && v === keyword) return true;
  }
  return false;
}

/**
 * Check if a specific row in ranking_results has actual results (not "Pending" or empty)
 * @param {number} targetRow - The row number to check
 * @returns {boolean} - True if row has actual results, false if empty or "Pending"
 */
function hasActualResultsAtRow(targetRow) {
  if (!targetRow || targetRow < 2) return false;
  const rs = getOrCreateRankingResultsSheet();
  const lastRow = rs.getLastRow();
  if (targetRow > lastRow) return false;
  
  const position = String(rs.getRange(targetRow, 1).getValue() || '').trim();
  // If position is "Pending" or empty, it's not a real result
  return position && position !== 'Pending' && position !== '';
}

/**
 * Test DataForSEO connection
 */
function testDataForSEOConnection() {
  try {
    const config = getDataForSEOConfig();
    console.log('✅ DataForSEO configuration loaded successfully');
    console.log(`🌐 Base URL: ${config.baseUrl}`);
    
    SpreadsheetApp.getUi().alert('✅ Connection Test', 'DataForSEO connection test successful!', SpreadsheetApp.getUi().ButtonSet.OK);
  } catch (error) {
    console.error('❌ Connection Error:', error);
    SpreadsheetApp.getUi().alert('❌ Connection Error', `Failed to connect: ${error.message}`, SpreadsheetApp.getUi().ButtonSet.OK);
  }
}


/**
 * Stop all scheduled triggers (submit and fetch)
 */
function stopAllSchedulers() {
  removeTriggerByHandler('submitScheduled');
  removeTriggerByHandler('fetchScheduled');
  SpreadsheetApp.getUi().alert('✅ Schedulers Stopped', 'All submit and fetch schedulers have been stopped.', SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast('🛑 All schedulers stopped');
}

/**
 * Reset the submit cursor to start from row 2
 */
function resetSubmitCursor() {
  const props = getScriptProps();
  props.deleteProperty('v3_submit_cursor');
  SpreadsheetApp.getUi().alert('✅ Cursor Reset', 'Submit cursor has been reset. Next run will start from row 2.', SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast('🔄 Submit cursor reset to row 2');
}

/**
 * Clear logs and ranking_results sheets (keeps row 1 headers)
 * Note: config sheet is not cleared as it contains configuration settings
 */
function clearResultsSheets() {
  const ss = SpreadsheetApp.getActive();
  let cleared = [];
  
  // Clear logs sheet
  const logsSheet = ss.getSheetByName('logs');
  if (logsSheet) {
    const lastRow = logsSheet.getLastRow();
    if (lastRow >= 2) {
      logsSheet.getRange(2, 1, lastRow - 1, logsSheet.getLastColumn()).clearContent();
      cleared.push('logs');
    }
  }
  
  // Clear ranking_results sheet
  const rankingSheet = ss.getSheetByName('ranking_results');
  if (rankingSheet) {
    const lastRow = rankingSheet.getLastRow();
    if (lastRow >= 2) {
      rankingSheet.getRange(2, 1, lastRow - 1, rankingSheet.getLastColumn()).clearContent();
      cleared.push('ranking_results');
    }
  }
  
  if (cleared.length > 0) {
    SpreadsheetApp.getUi().alert('✅ Sheets Cleared', `Cleared data from: ${cleared.join(', ')} (headers preserved)`, SpreadsheetApp.getUi().ButtonSet.OK);
    SpreadsheetApp.getActive().toast(`✅ Cleared: ${cleared.join(', ')}`);
  } else {
    SpreadsheetApp.getUi().alert('ℹ️ No Data', 'No data found to clear in logs or ranking_results sheets.', SpreadsheetApp.getUi().ButtonSet.OK);
  }
}

/**
 * Reset all ranking data: stop schedulers, reset cursor, and clear sheets
 * This is a destructive operation - shows confirmation dialog first
 */
function resetAllRankingData() {
  const ui = SpreadsheetApp.getUi();
  
  // Confirmation dialog
  const response = ui.alert(
    '⚠️ Reset All Ranking Data',
    'This will:\n' +
    '• Stop all schedulers (submit & fetch triggers)\n' +
    '• Reset submit cursor to row 2\n' +
    '• Clear logs and ranking_results sheets\n\n' +
    'Config sheet will NOT be cleared.\n\n' +
    'Continue?',
    ui.ButtonSet.YES_NO
  );
  
  if (response !== ui.Button.YES) {
    SpreadsheetApp.getActive().toast('❌ Reset cancelled');
    return;
  }
  
  const results = {
    schedulers: false,
    cursor: false,
    sheets: false
  };
  
  // 1. Stop all schedulers
  try {
    removeTriggerByHandler('submitScheduled');
    removeTriggerByHandler('fetchScheduled');
    results.schedulers = true;
  } catch (e) {
    console.error('Error stopping schedulers:', e);
  }
  
  // 2. Reset submit cursor
  try {
    const props = getScriptProps();
    props.deleteProperty('v3_submit_cursor');
    results.cursor = true;
  } catch (e) {
    console.error('Error resetting cursor:', e);
  }
  
  // 3. Clear sheets (suppress individual alerts, we'll show summary)
  try {
    const ss = SpreadsheetApp.getActive();
    let cleared = [];
    
    const logsSheet = ss.getSheetByName('logs');
    if (logsSheet) {
      const lastRow = logsSheet.getLastRow();
      if (lastRow >= 2) {
        logsSheet.getRange(2, 1, lastRow - 1, logsSheet.getLastColumn()).clearContent();
        cleared.push('logs');
      }
    }
    
    const rankingSheet = ss.getSheetByName('ranking_results');
    if (rankingSheet) {
      const lastRow = rankingSheet.getLastRow();
      if (lastRow >= 2) {
        rankingSheet.getRange(2, 1, lastRow - 1, rankingSheet.getLastColumn()).clearContent();
        cleared.push('ranking_results');
      }
    }
    
    results.sheets = true;
  } catch (e) {
    console.error('Error clearing sheets:', e);
  }
  
  // Summary message
  const successCount = Object.values(results).filter(v => v).length;
  const totalCount = Object.keys(results).length;
  
  if (successCount === totalCount) {
    ui.alert('✅ Reset Complete', 'All operations completed successfully:\n• Schedulers stopped\n• Cursor reset\n• Sheets cleared', ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast('✅ Reset complete: All operations successful');
  } else {
    const failed = [];
    if (!results.schedulers) failed.push('schedulers');
    if (!results.cursor) failed.push('cursor');
    if (!results.sheets) failed.push('sheets');
    ui.alert('⚠️ Partial Reset', `Some operations failed:\n• Failed: ${failed.join(', ')}\n\nPlease try individual operations if needed.`, ui.ButtonSet.OK);
    SpreadsheetApp.getActive().toast(`⚠️ Partial reset: ${failed.join(', ')} failed`);
  }
}
  
// ============================================================================
// CONFIG SHEET HELPERS (Sheet name: "config")
// ============================================================================

function getConfigSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('config');
  if (!sheet) {
    sheet = ss.insertSheet('config');
  }
  // Ensure headers (config only - no logs)
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1).setValue('Batch Size');
    sheet.getRange(1, 2).setValue('Submit Interval (minutes)');
    sheet.getRange(1, 3).setValue('Fetch Interval (minutes)');
  }
  // Default batch size at A2 if empty
  const batchSizeCell = sheet.getRange(2, 1);
  if (!batchSizeCell.getValue()) {
    batchSizeCell.setValue(50); // default batch size
  }
  // Default submit interval at B2 if empty
  const submitIntervalCell = sheet.getRange(2, 2);
  if (!submitIntervalCell.getValue()) {
    submitIntervalCell.setValue(5); // default 5 minutes
  }
  // Default fetch interval at C2 if empty
  const fetchIntervalCell = sheet.getRange(2, 3);
  if (!fetchIntervalCell.getValue()) {
    fetchIntervalCell.setValue(5); // default 5 minutes
  }
  return sheet;
}

// ============================================================================
// LOGS SHEET HELPERS (Sheet name: "logs")
// ============================================================================

function getLogsSheet() {
  const ss = SpreadsheetApp.getActive();
  let sheet = ss.getSheetByName('logs');
  if (!sheet) {
    sheet = ss.insertSheet('logs');
  }
  // Ensure headers
  if (sheet.getLastRow() === 0) {
    sheet.getRange(1, 1).setValue('Job IDs');
    sheet.getRange(1, 2).setValue('Keyword');
    sheet.getRange(1, 3).setValue('Source Row');
    sheet.getRange(1, 4).setValue('status');
    sheet.getRange(1, 5).setValue('submitted timestamp');
    sheet.getRange(1, 6).setValue('completed timestamp');
    // Basic styling for header row (optional)
    sheet.getRange(1, 1, 1, 6).setFontWeight('bold');
  }
  return sheet;
}

function readBatchSizeFromConfig() {
  const cfg = getConfigSheet();
  const val = Number(cfg.getRange(2, 1).getValue());
  return Number.isFinite(val) && val > 0 ? Math.floor(val) : 50;
}

function readSubmitIntervalFromConfig() {
  const cfg = getConfigSheet();
  const val = Number(cfg.getRange(2, 2).getValue()); // B2
  // Validate: must be integer, between 1 and 60 minutes
  if (!Number.isFinite(val) || val < 1 || val > 60) {
    return 5; // default
  }
  return Math.floor(val);
}

function readFetchIntervalFromConfig() {
  const cfg = getConfigSheet();
  const val = Number(cfg.getRange(2, 3).getValue()); // C2
  // Validate: must be integer, between 1 and 60 minutes
  if (!Number.isFinite(val) || val < 1 || val > 60) {
    return 5; // default
  }
  return Math.floor(val);
}

function getScriptProps() {
  return PropertiesService.getScriptProperties();
}

function findTriggersByHandler(handler) {
  return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === handler);
}

function ensureSubmitTrigger() {
  if (findTriggersByHandler('submitScheduled').length === 0) {
    const interval = readSubmitIntervalFromConfig();
    ScriptApp.newTrigger('submitScheduled').timeBased().everyMinutes(interval).create();
  }
}

function ensureFetchTrigger() {
  if (findTriggersByHandler('fetchScheduled').length === 0) {
    const interval = readFetchIntervalFromConfig();
    ScriptApp.newTrigger('fetchScheduled').timeBased().everyMinutes(interval).create();
  }
}

function removeTriggerByHandler(handler) {
  const triggers = findTriggersByHandler(handler);
  for (let i = 0; i < triggers.length; i++) {
    ScriptApp.deleteTrigger(triggers[i]);
  }
}

// ============================================================================
// SUBMIT PHASE (writes to logs: Job ID, status=submitted, submitted timestamp)
// ============================================================================

function submitScheduled() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  // Diagnostic: show immediate activity
  try { sheet.getRange(1, 14).setValue('🔎 Scanning for submit-ready rows...'); } catch (e) {}

  // Read input data from Sheet1 (K,L,M)
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActive().toast('❌ No data found in kw_variants');
    sheet.getRange(1, 14).setValue('❌ No data found');
    return;
  }

  const dataRange = sheet.getRange(2, 11, lastRow - 1, 3); // K,L,M
  const data = dataRange.getValues();

  // Use a moving cursor to avoid resubmitting the same rows across runs
  const props = getScriptProps();
  let cursor = Number(props.getProperty('v3_submit_cursor') || '2');
  if (!Number.isFinite(cursor) || cursor < 2) cursor = 2;

  const items = [];
  // Optimization: Read ranking_results data once to check which rows have actual results
  // This avoids thousands of individual sheet reads in the loop
  const rowsWithResults = new Set();
  try {
    const rs = getOrCreateRankingResultsSheet();
    const rsLastRow = rs.getLastRow();
    if (rsLastRow >= 2) {
      // Read all positions (column A) from ranking_results
      const positions = rs.getRange(2, 1, rsLastRow - 1, 1).getValues();
      for (let i = 0; i < positions.length; i++) {
        const position = String(positions[i][0] || '').trim();
        // If position has actual results (not "Pending" or empty), mark that row as processed
        if (position && position !== 'Pending' && position !== '') {
          rowsWithResults.add(i + 2); // i+2 because we started reading from row 2
        }
      }
    }
  } catch (e) {
    // ignore; fall back to empty set
  }
  
  // Check row-by-row if results already exist (row-specific, not keyword-only)
  // This allows duplicate keywords at different rows to be processed
  let scanRow = cursor;
  while (scanRow <= lastRow && items.length < 2000) { // hard cap safety
    const idx = scanRow - 2;
    const keyword = data[idx] ? data[idx][0] : '';
    const lat = data[idx] ? data[idx][1] : '';
    const lng = data[idx] ? data[idx][2] : '';
    const kwStr = String(keyword || '').trim();
    
    // Only skip if this specific row already has actual results (not "Pending")
    // This allows same keyword at different rows to be processed
    if (kwStr && lat && lng && !rowsWithResults.has(scanRow)) {
      items.push({ keyword: kwStr, lat: Number(lat), lng: Number(lng), row: scanRow });
    }
    scanRow++;
  }
  // Diagnostic: report how many found
  try { sheet.getRange(1, 14).setValue(`🧮 Found ${items.length} items to submit`); } catch (e) {}
  if (items.length === 0) {
    // No more rows to submit → remove submit trigger if present
    removeTriggerByHandler('submitScheduled');
    props.deleteProperty('v3_submit_cursor');
    SpreadsheetApp.getActive().toast('✅ Nothing to submit');
    sheet.getRange(1, 14).setValue('✅ Nothing to submit');
    return;
  }

  const batchSize = readBatchSizeFromConfig();
  const toSubmit = items.slice(0, batchSize);
  console.log(`📤 Submitting ${toSubmit.length} keywords (batch size ${batchSize})`);

  let taskIds = submitBatchRankingJobs(toSubmit);
  if (!taskIds || taskIds.length === 0) {
    SpreadsheetApp.getActive().toast('❌ Submit failed: no task IDs');
    sheet.getRange(1, 14).setValue('❌ Submit failed: no task IDs');
    return;
  }

  // Write to logs sheet (A..F)
  const logs = getLogsSheet();
  const now = new Date();
  const rows = taskIds.length;
  const startWriteRow = logs.getLastRow() >= 2 ? logs.getLastRow() + 1 : 2;
  const out = [];
  for (let i = 0; i < rows; i++) {
    out.push([taskIds[i], toSubmit[i].keyword, toSubmit[i].row, 'submitted', now, '']);
  }
  // A..F where A=Job IDs, B=Keyword, C=Source Row, D=status, E=submitted timestamp, F=completed timestamp
  logs.getRange(startWriteRow, 1, rows, 6).setValues(out);
  // Format Source Row column (C) as number to prevent auto-formatting as dates
  if (rows > 0) {
    logs.getRange(startWriteRow, 3, rows, 1).setNumberFormat('0'); // Column C = Source Row
  }
  
  // Write "Pending" placeholders to ranking_results at source row positions
  const rs = getOrCreateRankingResultsSheet();
  for (let i = 0; i < rows; i++) {
    const sourceRow = toSubmit[i].row;
    const keyword = toSubmit[i].keyword;
    // Ensure ranking_results has enough rows (pad if needed)
    const rsLastRow = rs.getLastRow();
    if (sourceRow > rsLastRow) {
      // Pad with empty rows, then write "Pending" at source row
      const rowsToAdd = sourceRow - rsLastRow;
      if (rowsToAdd > 0) {
        rs.insertRowsAfter(rsLastRow, rowsToAdd);
      }
    }
    // Write "Pending" placeholder (overwrite old data if present)
    // This ensures we start fresh for this keyword at this row
    rs.getRange(sourceRow, 1).setValue('Pending'); // Position
    rs.getRange(sourceRow, 2).setValue(''); // URL (empty)
    rs.getRange(sourceRow, 3).setValue(keyword); // Keyword
  }
  // Force flush to make "Pending" writes immediately visible
  SpreadsheetApp.flush();

  // Advance cursor
  const lastSubmittedRow = toSubmit[toSubmit.length - 1].row;
  const nextCursor = lastSubmittedRow + 1;
  if (nextCursor > lastRow) {
    props.deleteProperty('v3_submit_cursor');
    // No more rows to submit; remove submit trigger
    removeTriggerByHandler('submitScheduled');
  } else {
    props.setProperty('v3_submit_cursor', String(nextCursor));
    ensureSubmitTrigger(); // keep it running until we reach the end
  }
  
  // Ensure fetch trigger is active when new tasks are submitted
  ensureFetchTrigger();

  SpreadsheetApp.getActive().toast(`✅ Submitted ${rows} tasks`);
  sheet.getRange(1, 14).setValue(`✅ Submitted ${rows} tasks`);
}

// ============================================================================
// FETCH PHASE (reads logs, fetches ready tasks, writes to ranking_results)
// ============================================================================

function fetchScheduled() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  const logs = getLogsSheet();

  const lastRow = logs.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActive().toast('❌ No submitted tasks found in logs');
    sheet.getRange(1, 14).setValue('❌ No submitted tasks');
    return;
  }

  const rows = logs.getRange(2, 1, lastRow - 1, 6).getValues(); // A..F
  let fetched = 0;
  let pending = 0;
  let totalConsidered = 0;
  for (let i = 0; i < rows.length; i++) {
    const taskId = rows[i][0]; // A: Job IDs
    const keyword = rows[i][1]; // B: Keyword
    const sourceRow = rows[i][2]; // C: Source Row
    const status = rows[i][3]; // D: status
    const completedAt = rows[i][5]; // F: completed timestamp
    if (!taskId) continue;
    if (status === 'fetched' && completedAt) continue; // already done
    totalConsidered++;

    const res = fetchKeywordRankingResults(taskId, keyword);
    if (res && res.rankings && res.rankings.length > 0) {
      const best = res.rankings.reduce((b, c) => (c.rank < b.rank ? c : b));
      
      // Use keyword from config sheet (more reliable than extracting from raw data)
      const kw = String(keyword || '').trim();
      // Use sourceRow if available, otherwise append (backward compatibility)
      // Handle case where Google Sheets auto-formats numbers as dates (e.g., "1/1/1900" = row 2)
      let targetRow = null;
      if (sourceRow) {
        if (sourceRow instanceof Date) {
          // If stored as date, get the day number from the date
          // In Sheets, day 1 = 1/1/1900, day 2 = 1/2/1900, etc.
          // We need to calculate days since 12/30/1899 (Excel epoch)
          const excelEpoch = new Date(1899, 11, 30); // Dec 30, 1899
          const daysDiff = Math.floor((sourceRow.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
          targetRow = daysDiff >= 1 ? daysDiff + 1 : null; // +1 because row 2 = day 1
        } else {
          const numRow = Number(sourceRow);
          if (numRow >= 2 && Number.isFinite(numRow)) {
            targetRow = Math.floor(numRow);
          }
        }
      }
      
      // Only skip if target row has actual results (not "Pending" or old data we want to overwrite)
      if (targetRow && hasActualResultsAtRow(targetRow)) {
        // Check if the keyword matches - if so, skip (already processed)
        const rs = getOrCreateRankingResultsSheet();
        const existingKeyword = String(rs.getRange(targetRow, 3).getValue() || '').trim();
        if (existingKeyword === kw) {
          logs.getRange(i + 2, 4).setValue('fetched'); // D: status
          logs.getRange(i + 2, 6).setValue(new Date()); // F: completed timestamp
          continue;
        }
      } else if (!targetRow && kw && keywordAlreadyProcessed(kw)) {
        // For backward compatibility (no targetRow), check if keyword exists anywhere
        logs.getRange(i + 2, 4).setValue('fetched'); // D: status
        logs.getRange(i + 2, 6).setValue(new Date()); // F: completed timestamp
        continue;
      }
      
      // Write result (will overwrite "Pending" or old data at target row)
      writeSingleRankingResult(sheet, 2, { // row is unused; we pass keyword explicitly
        ranking: best.rank,
        url: best.url,
        keyword: kw
      }, targetRow);
      logs.getRange(i + 2, 4).setValue('fetched'); // D: status
      logs.getRange(i + 2, 6).setValue(new Date()); // F: completed timestamp
      fetched++;
    } else {
      logs.getRange(i + 2, 4).setValue('pending'); // D: status - keep waiting
      pending++;
    }
  }

  SpreadsheetApp.getActive().toast(`ℹ️ Fetch complete | Fetched: ${fetched} | Pending: ${pending}`);
  sheet.getRange(1, 14).setValue(`ℹ️ Fetch complete | Fetched: ${fetched} | Pending: ${pending}`);

  // If there is nothing left to fetch, remove the fetch trigger
  if (pending === 0 && totalConsidered === 0) {
    removeTriggerByHandler('fetchScheduled');
  } else {
    ensureFetchTrigger();
  }
}

// ============================================================================
// TRIGGER HELPERS (manual creation, every 30 minutes)
// ============================================================================

function createFetchTrigger() {
  // Remove existing fetch triggers
  const triggers = ScriptApp.getProjectTriggers();
  for (let i = 0; i < triggers.length; i++) {
    if (triggers[i].getHandlerFunction() === 'fetchScheduled') {
      ScriptApp.deleteTrigger(triggers[i]);
    }
  }
  const interval = readFetchIntervalFromConfig();
  ScriptApp.newTrigger('fetchScheduled').timeBased().everyMinutes(interval).create();
  SpreadsheetApp.getUi().alert('✅ Trigger Created', `Fetch trigger set to run every ${interval} minute${interval !== 1 ? 's' : ''}.`, SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// ORCHESTRATOR ENTRY POINT
// ============================================================================

function runRankingCheckOnSheet() {
  const submitInterval = readSubmitIntervalFromConfig();
  const fetchInterval = readFetchIntervalFromConfig();
  SpreadsheetApp.getUi().alert('🚀 Scheduler Starting', `Setting up submit (${submitInterval}-min) and fetch (${fetchInterval}-min) schedulers now.`, SpreadsheetApp.getUi().ButtonSet.OK);
  // Kick off first batch immediately
  submitScheduled();
  // Remove existing triggers to ensure they're recreated with new intervals from config
  removeTriggerByHandler('submitScheduled');
  removeTriggerByHandler('fetchScheduled');
  // Create triggers with current config values
  ensureSubmitTrigger();
  ensureFetchTrigger();
  SpreadsheetApp.getUi().alert('🚀 Scheduler Running', `Submission (${submitInterval}-min) and fetch (${fetchInterval}-min) schedulers are active. You can close the sheet; they will auto-stop when finished.`, SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(`🚀 Scheduler running: submit (${submitInterval}-min) + fetch (${fetchInterval}-min) active`);
  const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  sheet.getRange(1, 14).setValue(`🚀 Scheduler running: submit (${submitInterval}-min) + fetch (${fetchInterval}-min) active`);
}