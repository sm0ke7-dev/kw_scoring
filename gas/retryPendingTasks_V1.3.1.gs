/**
 * retryPendingTasks_V1.3.1.gs
 * Enhanced fetch functionality with manual retry, better logging, timeout handling, and improved status handling
 * V1.2: Added proper handling for "no_results" status - marks as "no_results" instead of "pending"
 * V1.3.1: Logs consolidated into ranking_results columns J-N (no longer using separate logs sheet)
 */

// ranking_results column indices (1-based for Google Sheets)
const RANKING_COL = {
  POSITION: 1,      // A
  URL: 2,           // B
  KEYWORD: 3,       // C
  RANK_SCORE: 4,    // D
  KEYWORD_SCORE: 5, // E
  GEO: 6,           // F
  FINAL_SCORE: 7,   // G
  IDEAL: 8,         // H
  GAP: 9,           // I
  JOB_ID: 10,       // J - from old logs A (Job IDs)
  SOURCE_ROW: 11,   // K - from old logs C (Source Row)
  STATUS: 12,       // L - from old logs D (status)
  SUBMITTED_AT: 13, // M - from old logs E (submitted timestamp)
  COMPLETED_AT: 14  // N - from old logs F (completed timestamp)
};

/**
 * Manual retry function for pending tasks
 * Resubmits all pending tasks and then attempts to fetch results
 */
function retryPendingTasks() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  const rs = getOrCreateRankingResultsSheet();
  
  const lastRow = rs.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('ℹ️ No Tasks', 'No tasks found in ranking_results sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // v1.3.1: Read logs data from ranking_results columns J-N
  const logsData = rs.getRange(2, RANKING_COL.JOB_ID, lastRow - 1, 5).getValues(); // J-N (5 columns)
  const keywordData = rs.getRange(2, RANKING_COL.KEYWORD, lastRow - 1, 1).getValues(); // C (1 column)
  const pendingTasks = [];
  
  // Step 1: Collect all pending tasks and get their coordinates from kw_variants (OPTIMIZED: batch read)
  console.log(`🔍 Finding pending tasks to resubmit...`);
  
  // First pass: Collect all unique row numbers that need to be read
  const rowNumbersToRead = new Set();
  const logEntries = [];
  
  for (let i = 0; i < logsData.length; i++) {
    const taskId = logsData[i][0]; // J: Job ID (index 0 in logsData)
    const keyword = keywordData[i][0]; // C: Keyword
    const sourceRow = logsData[i][1]; // K: Source Row (index 1 in logsData)
    const status = logsData[i][2]; // L: status (index 2 in logsData)
    const completedAt = logsData[i][4]; // N: completed timestamp (index 4 in logsData)
    
    // Only process pending/submitted tasks that haven't been completed
    // Skip "no_results" - task is complete, no rankings found
    if (!taskId) continue;
    if (status === 'fetched' && completedAt) continue;
    if (status === 'no_results') continue; // Skip - already marked as no results
    
    // Convert sourceRow to row number
    let rowNum = null;
    if (sourceRow) {
      if (sourceRow instanceof Date) {
        const excelEpoch = new Date(1899, 11, 30);
        const daysDiff = Math.floor((sourceRow.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
        rowNum = daysDiff >= 1 ? daysDiff + 1 : null;
      } else {
        const numRow = Number(sourceRow);
        if (numRow >= 2 && Number.isFinite(numRow)) {
          rowNum = Math.floor(numRow);
        }
      }
    }
    
    if (rowNum) {
      rowNumbersToRead.add(rowNum);
    }
    
    logEntries.push({
      taskId: taskId,
      keyword: keyword,
      sourceRow: sourceRow,
      rowNum: rowNum,
      logIndex: i + 2
    });
  }
  
  if (logEntries.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ No Pending Tasks', 'No pending tasks found to resubmit.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Batch read all needed rows from kw_variants in one call
  console.log(`📖 Batch reading ${rowNumbersToRead.size} rows from kw_variants...`);
  const rowMap = new Map();
  
  if (rowNumbersToRead.size > 0) {
    const rowArray = Array.from(rowNumbersToRead).sort((a, b) => a - b);
    const minRow = Math.min(...rowArray);
    const maxRow = Math.max(...rowArray);
    const totalRows = maxRow - minRow + 1;
    
    // Read all rows in one batch (columns K, L, M)
    try {
      const batchData = sheet.getRange(minRow, 11, totalRows, 3).getValues(); // K, L, M
      
      // Create lookup map: row number -> [keyword, lat, lng]
      for (let idx = 0; idx < batchData.length; idx++) {
        const actualRow = minRow + idx;
        if (rowNumbersToRead.has(actualRow)) {
          rowMap.set(actualRow, batchData[idx]);
        }
      }
      
      console.log(`✅ Batch read complete: ${rowMap.size} rows loaded`);
    } catch (e) {
      console.error(`❌ Error batch reading from kw_variants:`, e);
      SpreadsheetApp.getUi().alert('❌ Error', 'Failed to read data from kw_variants sheet. Check console logs.', SpreadsheetApp.getUi().ButtonSet.OK);
      return;
    }
  }
  
  // Second pass: Use batch-read data to build pendingTasks array
  for (const entry of logEntries) {
    let kw = String(entry.keyword || '').trim();
    let lat = null;
    let lng = null;
    
    if (entry.rowNum && rowMap.has(entry.rowNum)) {
      const kwData = rowMap.get(entry.rowNum);
      kw = String(kwData[0] || entry.keyword || '').trim();
      lat = Number(kwData[1]);
      lng = Number(kwData[2]);
    }
    
    if (kw && lat && lng) {
      pendingTasks.push({
        keyword: kw,
        lat: lat,
        lng: lng,
        sourceRow: entry.sourceRow,
        logIndex: entry.logIndex,
        oldTaskId: entry.taskId
      });
      
      if (pendingTasks.length % 100 === 0) {
        console.log(`📝 Collected ${pendingTasks.length} tasks so far...`);
      }
    } else {
      console.warn(`⚠️ Skipping task - missing data: keyword=${kw}, lat=${lat}, lng=${lng} (row ${entry.rowNum})`);
    }
  }
  
  console.log(`✅ Collected ${pendingTasks.length} pending tasks to resubmit`);
  
  if (pendingTasks.length === 0) {
    SpreadsheetApp.getUi().alert('ℹ️ No Pending Tasks', 'No pending tasks found to resubmit.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  console.log(`📤 Resubmitting ${pendingTasks.length} pending tasks...`);
  SpreadsheetApp.getActive().toast(`📤 Resubmitting ${pendingTasks.length} tasks...`);
  
  // Step 2: Resubmit all pending tasks
  const toResubmit = pendingTasks.map(t => ({
    keyword: t.keyword,
    lat: t.lat,
    lng: t.lng,
    row: t.sourceRow
  }));
  
  const newTaskIds = submitBatchRankingJobs(toResubmit);
  
  if (!newTaskIds || newTaskIds.length === 0) {
    SpreadsheetApp.getUi().alert('❌ Resubmit Failed', 'Failed to resubmit tasks. Check console logs for details.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  if (newTaskIds.length !== pendingTasks.length) {
    console.warn(`⚠️ Warning: Submitted ${newTaskIds.length} tasks but expected ${pendingTasks.length}`);
  }
  
  // Step 3: Update ranking_results with new task IDs (columns J, L, M)
  const now = new Date();
  for (let i = 0; i < Math.min(newTaskIds.length, pendingTasks.length); i++) {
    const logRow = pendingTasks[i].logIndex;
    rs.getRange(logRow, RANKING_COL.JOB_ID).setValue(newTaskIds[i]); // J: Job ID
    rs.getRange(logRow, RANKING_COL.STATUS).setValue('RE-SUBMITTED'); // L: status
    rs.getRange(logRow, RANKING_COL.SUBMITTED_AT).setValue(now); // M: submitted timestamp
  }
  
  console.log(`✅ Resubmitted ${newTaskIds.length} tasks with new task IDs (status: RE-SUBMITTED)`);
  
  // Step 4: Ensure fetch scheduler is running (5 minutes hardcoded)
  // Skip immediate fetch - tasks need 5+ minutes to process, scheduler will handle it
  if (newTaskIds.length > 0) {
    console.log(`⏰ Creating fetch scheduler (5 minutes) to fetch ${newTaskIds.length} resubmitted tasks...`);
    
    // Remove any existing fetch triggers first
    const existingTriggers = ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === 'fetchScheduled');
    for (let i = 0; i < existingTriggers.length; i++) {
      ScriptApp.deleteTrigger(existingTriggers[i]);
    }
    
    // Create new fetch trigger with 5 minutes hardcoded
    ScriptApp.newTrigger('fetchScheduled').timeBased().everyMinutes(5).create();
    console.log(`✅ Fetch scheduler created (runs every 5 minutes)`);
  }
  
  // Summary
  const summary = `Retry Complete:\n` +
    `📤 Resubmitted: ${newTaskIds.length} tasks\n` +
    `📝 Status: RE-SUBMITTED\n` +
    `⏰ Fetch scheduler created (5 minutes)\n\n` +
    `Tasks will be fetched automatically by the scheduler.`;
  
  SpreadsheetApp.getUi().alert('🔄 Retry Complete', summary, SpreadsheetApp.getUi().ButtonSet.OK);
  SpreadsheetApp.getActive().toast(`✅ Retry: ${newTaskIds.length} resubmitted, scheduler active (5 min)`);
  sheet.getRange(1, 14).setValue(`✅ Retry: ${newTaskIds.length} resubmitted, scheduler active (5 min)`);
  
  console.log(summary);
}

/**
 * Mark tasks as failed if they've been pending too long
 * Reads timeout hours from config sheet (default: 24 hours)
 * @returns {number} Number of tasks marked as failed
 */
function markTimedOutTasks() {
  const rs = getOrCreateRankingResultsSheet();
  const lastRow = rs.getLastRow();
  if (lastRow < 2) return 0;
  
  // Read timeout from config (default 24 hours)
  const timeoutHours = readTimeoutHoursFromConfig();
  const timeoutMs = timeoutHours * 60 * 60 * 1000;
  const now = new Date().getTime();
  
  // v1.3.1: Read logs data from ranking_results columns J-N
  const logsData = rs.getRange(2, RANKING_COL.JOB_ID, lastRow - 1, 5).getValues(); // J-N (5 columns)
  let markedFailed = 0;
  
  for (let i = 0; i < logsData.length; i++) {
    const status = logsData[i][2]; // L: status (index 2 in logsData)
    const submittedAt = logsData[i][3]; // M: submitted timestamp (index 3 in logsData)
    const completedAt = logsData[i][4]; // N: completed timestamp (index 4 in logsData)
    
    // Only check pending/submitted tasks that haven't been completed
    // Skip "no_results" - task is complete, just no rankings found
    if (status !== 'fetched' && status !== 'failed' && status !== 'no_results' && submittedAt) {
      const submittedTime = submittedAt instanceof Date ? submittedAt.getTime() : new Date(submittedAt).getTime();
      const timeSinceSubmission = now - submittedTime;
      
      if (timeSinceSubmission > timeoutMs) {
        rs.getRange(i + 2, RANKING_COL.STATUS).setValue('failed'); // L: status
        markedFailed++;
        console.log(`⏰ Marked task as failed (timeout): Row ${i + 2}, submitted ${Math.round(timeSinceSubmission / (60 * 60 * 1000))} hours ago`);
      }
    }
  }
  
  if (markedFailed > 0) {
    console.log(`⏰ Marked ${markedFailed} tasks as failed due to timeout (${timeoutHours} hours)`);
  }
  
  return markedFailed;
}

/**
 * Read timeout hours from config sheet
 * Default: 24 hours
 * @returns {number} Timeout in hours
 */
function readTimeoutHoursFromConfig() {
  try {
    const cfg = getConfigSheet();
    // Check column D (row 2) for timeout hours
    const val = Number(cfg.getRange(2, 4).getValue());
    if (Number.isFinite(val) && val > 0 && val <= 168) { // Max 1 week
      return Math.floor(val);
    }
  } catch (e) {
    console.log('Using default timeout: 24 hours');
  }
  return 24; // Default: 24 hours
}

/**
 * Enhanced fetch with better logging and proper status handling
 * V1.2: Handles "no_results" status - marks as "no_results" and writes "Not Found" to ranking_results
 * Can be called from fetchScheduled or used independently
 */
function fetchWithEnhancedLogging() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  const rs = getOrCreateRankingResultsSheet();
  
  const lastRow = rs.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getActive().toast('❌ No submitted tasks found in ranking_results');
    sheet.getRange(1, 14).setValue('❌ No submitted tasks');
    return;
  }
  
  // v1.3.1: Read logs data from ranking_results columns J-N
  const logsData = rs.getRange(2, RANKING_COL.JOB_ID, lastRow - 1, 5).getValues(); // J-N (5 columns)
  const keywordData = rs.getRange(2, RANKING_COL.KEYWORD, lastRow - 1, 1).getValues(); // C (1 column)
  
  let fetched = 0;
  let pending = 0;
  let noResults = 0;
  let totalConsidered = 0;
  const logDetails = [];
  
  for (let i = 0; i < logsData.length; i++) {
    const taskId = logsData[i][0]; // J: Job ID (index 0 in logsData)
    const keyword = keywordData[i][0]; // C: Keyword
    const sourceRow = logsData[i][1]; // K: Source Row (index 1 in logsData)
    const status = logsData[i][2]; // L: status (index 2 in logsData)
    const completedAt = logsData[i][4]; // N: completed timestamp (index 4 in logsData)
    if (!taskId) continue;
    if (status === 'fetched' && completedAt) continue; // already done
    if (status === 'no_results') continue; // already marked as no results, skip
    totalConsidered++;
    
    const res = fetchKeywordRankingResults(taskId, keyword);
    
    // Check status first to handle "no_results" properly
    if (res && res.status === 'no_results') {
      // Task is complete but no rankings found - mark as "no_results" and write "Not Found"
      const kw = String(keyword || '').trim();
      
      // Handle targetRow
      let targetRow = null;
      if (sourceRow) {
        if (sourceRow instanceof Date) {
          const excelEpoch = new Date(1899, 11, 30);
          const daysDiff = Math.floor((sourceRow.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
          targetRow = daysDiff >= 1 ? daysDiff + 1 : null;
        } else {
          const numRow = Number(sourceRow);
          if (numRow >= 2 && Number.isFinite(numRow)) {
            targetRow = Math.floor(numRow);
          }
        }
      }
      
      // Write "Not Found" to ranking_results
      writeSingleRankingResult(sheet, 2, {
        ranking: 'Not Found',
        url: 'Not Found',
        keyword: kw
      }, targetRow);
      
      rs.getRange(i + 2, RANKING_COL.STATUS).setValue('no_results'); // L: status
      rs.getRange(i + 2, RANKING_COL.COMPLETED_AT).setValue(new Date()); // N: completed timestamp
      noResults++;
      logDetails.push(`❌ ${kw}: No results found`);
      console.log(`❌ No results found for: ${kw}`);
      
    } else if (res && res.rankings && res.rankings.length > 0) {
      // Rankings found - write result
      const best = res.rankings.reduce((b, c) => (c.rank < b.rank ? c : b));
      const kw = String(keyword || '').trim();
      
      // Handle targetRow (same logic as fetchScheduled)
      let targetRow = null;
      if (sourceRow) {
        if (sourceRow instanceof Date) {
          const excelEpoch = new Date(1899, 11, 30);
          const daysDiff = Math.floor((sourceRow.getTime() - excelEpoch.getTime()) / (1000 * 60 * 60 * 24));
          targetRow = daysDiff >= 1 ? daysDiff + 1 : null;
        } else {
          const numRow = Number(sourceRow);
          if (numRow >= 2 && Number.isFinite(numRow)) {
            targetRow = Math.floor(numRow);
          }
        }
      }
      
      // Check if already processed
      if (targetRow && hasActualResultsAtRow(targetRow)) {
        const existingKeyword = String(rs.getRange(targetRow, RANKING_COL.KEYWORD).getValue() || '').trim();
        if (existingKeyword === kw) {
          rs.getRange(i + 2, RANKING_COL.STATUS).setValue('fetched');
          rs.getRange(i + 2, RANKING_COL.COMPLETED_AT).setValue(new Date());
          continue;
        }
      } else if (!targetRow && kw && keywordAlreadyProcessed(kw)) {
        rs.getRange(i + 2, RANKING_COL.STATUS).setValue('fetched');
        rs.getRange(i + 2, RANKING_COL.COMPLETED_AT).setValue(new Date());
        continue;
      }
      
      // Write result
      writeSingleRankingResult(sheet, 2, {
        ranking: best.rank,
        url: best.url,
        keyword: kw
      }, targetRow);
      
      rs.getRange(i + 2, RANKING_COL.STATUS).setValue('fetched');
      rs.getRange(i + 2, RANKING_COL.COMPLETED_AT).setValue(new Date());
      fetched++;
      logDetails.push(`✅ ${kw}: Rank ${best.rank}`);
      
    } else {
      // Still processing or error - keep as pending
      let reason = 'No results yet';
      if (res && res.status === 'error') {
        reason = `API Error: ${res.error || 'Unknown'}`;
      }
      
      rs.getRange(i + 2, RANKING_COL.STATUS).setValue('pending'); // L: status
      pending++;
      logDetails.push(`⏳ ${keyword || taskId}: ${reason}`);
      console.log(`⏳ Pending - ${keyword || taskId}: ${reason}`);
    }
  }
  
  // Mark timed out tasks
  const timedOut = markTimedOutTasks();
  
  SpreadsheetApp.getActive().toast(`ℹ️ Fetch complete | Fetched: ${fetched} | No Results: ${noResults} | Pending: ${pending} | Failed: ${timedOut}`);
  sheet.getRange(1, 14).setValue(`ℹ️ Fetch: ${fetched} fetched, ${noResults} no results, ${pending} pending, ${timedOut} failed`);
  
  // Log details to console
  console.log(`📊 Fetch Summary: ${fetched} fetched, ${noResults} no results, ${pending} pending, ${timedOut} failed`);
  if (logDetails.length > 0) {
    console.log('Details:', logDetails.slice(0, 10).join(', ')); // Log first 10
  }
}

