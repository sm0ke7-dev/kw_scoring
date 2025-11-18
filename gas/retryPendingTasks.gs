/**
 * retryPendingTasks.gs
 * Enhanced fetch functionality with manual retry, better logging, and timeout handling
 */

/**
 * Manual retry function for pending tasks
 * Resubmits all pending tasks and then attempts to fetch results
 */
function retryPendingTasks() {
  const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
  const logs = getLogsSheet();
  
  const lastRow = logs.getLastRow();
  if (lastRow < 2) {
    SpreadsheetApp.getUi().alert('ℹ️ No Tasks', 'No tasks found in logs sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }
  
  // Read all log rows
  const logRows = logs.getRange(2, 1, lastRow - 1, 6).getValues(); // A..F
  const pendingTasks = [];
  
  // Step 1: Collect all pending tasks and get their coordinates from kw_variants
  console.log(`🔍 Finding pending tasks to resubmit...`);
  
  for (let i = 0; i < logRows.length; i++) {
    const taskId = logRows[i][0]; // A: Job IDs
    const keyword = logRows[i][1]; // B: Keyword
    const sourceRow = logRows[i][2]; // C: Source Row
    const status = logRows[i][3]; // D: status
    const completedAt = logRows[i][5]; // F: completed timestamp
    
    // Only process pending/submitted tasks that haven't been completed
    if (!taskId) continue;
    if (status === 'fetched' && completedAt) continue;
    
    // Get coordinates from kw_variants using sourceRow
    let lat = null;
    let lng = null;
    let kw = String(keyword || '').trim();
    
    if (sourceRow) {
      let rowNum = null;
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
      
      if (rowNum) {
        // Read keyword, lat, lng from kw_variants (columns K, L, M)
        try {
          const kwData = sheet.getRange(rowNum, 11, 1, 3).getValues()[0]; // K, L, M
          kw = String(kwData[0] || keyword || '').trim();
          lat = Number(kwData[1]);
          lng = Number(kwData[2]);
        } catch (e) {
          console.error(`⚠️ Could not read row ${rowNum} from kw_variants:`, e);
          continue;
        }
      }
    }
    
    if (kw && lat && lng) {
      // Calculate targetRow for logging
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
      
      pendingTasks.push({
        keyword: kw,
        lat: lat,
        lng: lng,
        sourceRow: sourceRow,
        logIndex: i + 2, // Row number in logs sheet
        oldTaskId: taskId
      });
      
      console.log(`📝 Task ${pendingTasks.length}: "${kw}" -> Will write to ranking_results row ${rowNum || 'end'}`);
    } else {
      console.warn(`⚠️ Skipping task - missing data: keyword=${kw}, lat=${lat}, lng=${lng}`);
    }
  }
  
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
  
  // Step 3: Update logs with new task IDs
  const now = new Date();
  for (let i = 0; i < Math.min(newTaskIds.length, pendingTasks.length); i++) {
    const logRow = pendingTasks[i].logIndex;
    logs.getRange(logRow, 1).setValue(newTaskIds[i]); // A: Job ID
    logs.getRange(logRow, 4).setValue('RE-SUBMITTED'); // D: status
    logs.getRange(logRow, 5).setValue(now); // E: submitted timestamp
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
  const logs = getLogsSheet();
  const lastRow = logs.getLastRow();
  if (lastRow < 2) return 0;
  
  // Read timeout from config (default 24 hours)
  const timeoutHours = readTimeoutHoursFromConfig();
  const timeoutMs = timeoutHours * 60 * 60 * 1000;
  const now = new Date().getTime();
  
  const rows = logs.getRange(2, 1, lastRow - 1, 6).getValues(); // A..F
  let markedFailed = 0;
  
  for (let i = 0; i < rows.length; i++) {
    const status = rows[i][3]; // D: status
    const submittedAt = rows[i][4]; // E: submitted timestamp
    const completedAt = rows[i][5]; // F: completed timestamp
    
    // Only check pending/submitted tasks that haven't been completed
    if (status !== 'fetched' && status !== 'failed' && submittedAt) {
      const submittedTime = submittedAt instanceof Date ? submittedAt.getTime() : new Date(submittedAt).getTime();
      const timeSinceSubmission = now - submittedTime;
      
      if (timeSinceSubmission > timeoutMs) {
        logs.getRange(i + 2, 4).setValue('failed'); // D: status
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
 * Enhanced fetch with better logging
 * Can be called from fetchScheduled or used independently
 */
function fetchWithEnhancedLogging() {
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
  const logDetails = [];
  
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
        const rs = getOrCreateRankingResultsSheet();
        const existingKeyword = String(rs.getRange(targetRow, 3).getValue() || '').trim();
        if (existingKeyword === kw) {
          logs.getRange(i + 2, 4).setValue('fetched');
          logs.getRange(i + 2, 6).setValue(new Date());
          continue;
        }
      } else if (!targetRow && kw && keywordAlreadyProcessed(kw)) {
        logs.getRange(i + 2, 4).setValue('fetched');
        logs.getRange(i + 2, 6).setValue(new Date());
        continue;
      }
      
      // Write result
      writeSingleRankingResult(sheet, 2, {
        ranking: best.rank,
        url: best.url,
        keyword: kw
      }, targetRow);
      
      logs.getRange(i + 2, 4).setValue('fetched');
      logs.getRange(i + 2, 6).setValue(new Date());
      fetched++;
      logDetails.push(`✅ ${kw}: Rank ${best.rank}`);
      
    } else {
      // Enhanced logging for pending tasks
      let reason = 'No results yet';
      if (res && res.status === 'error') {
        reason = `API Error: ${res.error || 'Unknown'}`;
      } else if (res && res.status === 'no_results') {
        reason = 'No results found in SERP';
      }
      
      logs.getRange(i + 2, 4).setValue('pending');
      pending++;
      logDetails.push(`⏳ ${keyword || taskId}: ${reason}`);
      console.log(`⏳ Pending - ${keyword || taskId}: ${reason}`);
    }
  }
  
  // Mark timed out tasks
  const timedOut = markTimedOutTasks();
  
  SpreadsheetApp.getActive().toast(`ℹ️ Fetch complete | Fetched: ${fetched} | Pending: ${pending} | Failed: ${timedOut}`);
  sheet.getRange(1, 14).setValue(`ℹ️ Fetch: ${fetched} fetched, ${pending} pending, ${timedOut} failed`);
  
  // Log details to console
  console.log(`📊 Fetch Summary: ${fetched} fetched, ${pending} pending, ${timedOut} failed`);
  if (logDetails.length > 0) {
    console.log('Details:', logDetails.slice(0, 10).join(', ')); // Log first 10
  }
}


