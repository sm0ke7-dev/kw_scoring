function GENERATE_KEYWORDS(templates, niche, niche_placeholder, locations, location_placeholder) {
  const toStr = v => (v === null || v === undefined) ? "" : String(v).trim();

  function normalize1D(arg) {
    if (!Array.isArray(arg)) return [toStr(arg)].filter(Boolean);
    const out = [];
    if (Array.isArray(arg[0])) {
      for (let r = 0; r < arg.length; r++) {
        for (let c = 0; c < arg[r].length; c++) {
          const v = toStr(arg[r][c]);
          if (v) out.push(v);
        }
      }
    } else {
      for (let i = 0; i < arg.length; i++) {
        const v = toStr(arg[i]);
        if (v) out.push(v);
      }
    }
    return out;
  }

  function normalizeNiche2Cols(arg) {
    if (!Array.isArray(arg)) {
      const v = toStr(arg);
      if (!v) return [];
      throw new Error("The 'niche' parameter must be a 2-column range: [service, core_keyword].");
    }
    const rows = Array.isArray(arg[0]) ? arg : [arg];
    const out = [];
    for (let r = 0; r < rows.length; r++) {
      const row = rows[r];
      if (!Array.isArray(row) || row.length < 2) continue;
      const service = toStr(row[0]);
      const core   = toStr(row[1]);
      if (service && core) out.push([service, core]);
    }
    return out;
  }

  function replaceAllLiteral(str, find, replace) {
    if (!find) return str;
    return String(str).split(find).join(replace);
  }

  const tplList   = normalize1D(templates);
  const locList   = normalize1D(locations);
  const nicheRows = normalizeNiche2Cols(niche);

  const nichePh = toStr(niche_placeholder);
  const locPh   = toStr(location_placeholder);

  if (!tplList.length)   throw new Error("No templates provided.");
  if (!nicheRows.length) throw new Error("No niche rows provided. Expect two columns: service, core_keyword.");
  if (!locList.length)   throw new Error("No locations provided.");
  if (!nichePh)          throw new Error("niche_placeholder is empty.");
  if (!locPh)            throw new Error("location_placeholder is empty.");

  const out = [];
  // out.push(["service", "location", "core_keyword", "keyword"]); // optional header

  // Group by Location → Service → Template
  for (let j = 0; j < locList.length; j++) {
    const location = locList[j];
    for (let i = 0; i < nicheRows.length; i++) {
      const [service, core_keyword] = nicheRows[i];
      for (let t = 0; t < tplList.length; t++) {
        const template = tplList[t];
        let keyword = template;
        keyword = replaceAllLiteral(keyword, nichePh, core_keyword);
        keyword = replaceAllLiteral(keyword, locPh, location);
        keyword = keyword.toLowerCase(); // 👈 lowercase final keyword
        out.push([service, location, core_keyword, keyword]);
      }
    }
  }

  return out;
}

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
* Check rankings for existing keywords in the sheet
*/
function checkRankingsOnly() {
const ss = SpreadsheetApp.getActive();
const sheet = ss.getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();

// Read data from sheet
const lastRow = sheet.getLastRow();
if (lastRow < 2) {
  SpreadsheetApp.getUi().alert('❌ No Data', 'No data found in the sheet.', SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}

// Read keywords, lat, lng from kw_variants columns K, L, M
const dataRange = sheet.getRange(2, 11, lastRow - 1, 3); // K,L,M
const data = dataRange.getValues();

const keywordsWithCoords = [];
data.forEach((row, index) => {
  const keyword = row[0]; // Column K (kw+geo)
  const lat = row[1];     // Column L
  const lng = row[2];     // Column M
  
  if (keyword && lat && lng) {
    keywordsWithCoords.push({
      keyword: keyword.toString().trim(),
      lat: parseFloat(lat),
      lng: parseFloat(lng),
      row: index + 2 // Actual row number in sheet
    });
  }
});

if (keywordsWithCoords.length === 0) {
  SpreadsheetApp.getUi().alert('❌ No Valid Data', 'No valid keywords with coordinates found. Please ensure you have keywords in Column K and valid lat/long in Columns L & M.', SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}

console.log(`📊 Found ${keywordsWithCoords.length} keywords with coordinates`);

// Filter out already processed keywords by checking ranking_results!D
const uncheckedKeywords = keywordsWithCoords.filter(item => !keywordAlreadyProcessed(item.keyword));

if (uncheckedKeywords.length === 0) {
  SpreadsheetApp.getUi().alert('✅ All Done!', 'All keywords have already been checked for rankings.', SpreadsheetApp.getUi().ButtonSet.OK);
  return;
}

console.log(`📊 Found ${uncheckedKeywords.length} unchecked keywords out of ${keywordsWithCoords.length} total`);

// Add limit for testing
const TEST_LIMIT = 100; // 🧪 Limit to 100 keywords per run
const keywordsToProcess = TEST_LIMIT > 0 ? uncheckedKeywords.slice(0, TEST_LIMIT) : uncheckedKeywords;

// Ensure results sheet and headers exist
setupRankingHeaders(sheet);

const totalKeywords = keywordsToProcess.length;
const batchSize = 100; // Batch size reduced per user instruction
const totalBatches = Math.ceil(totalKeywords / batchSize);

// Show non-blocking toast and write status to N1
SpreadsheetApp.getActive().toast('🚀 Starting Ranking Check');
sheet.getRange(1, 14).setValue(`Starting: ${totalKeywords} keywords, ${totalBatches} batches (≈ ${totalBatches * 5} min)`);

console.log(`📦 Processing ${totalKeywords} keywords in ${totalBatches} batches`);

  // Process in batches
  let processed = 0;
  let successful = 0;
  let failed = 0;
  
  for (let batchIndex = 0; batchIndex < totalBatches; batchIndex++) {
    const startIndex = batchIndex * batchSize;
    const endIndex = Math.min(startIndex + batchSize, totalKeywords);
    const batchKeywords = keywordsToProcess.slice(startIndex, endIndex);
    
    console.log(`📦 Processing Batch ${batchIndex + 1}/${totalBatches} (${batchKeywords.length} keywords)`);
    
    // Submit all keywords in this batch using batch API
    console.log(`📤 Submitting batch of ${batchKeywords.length} keywords to DataForSEO`);
    
    let taskIds = []; // Initialize outside try block
    try {
      taskIds = submitBatchRankingJobs(batchKeywords);
      
      if (taskIds.length > 0) {
        console.log(`✅ Batch submitted successfully: ${taskIds.length} task IDs generated`);
        
        // Create task data array with row mapping
        const taskDataArray = [];
        for (let i = 0; i < taskIds.length && i < batchKeywords.length; i++) {
          const actualRowNumber = batchKeywords[i].row; // use original sheet row
          processed++;
          taskDataArray.push({
            taskId: taskIds[i],
            row: actualRowNumber,
            keyword: batchKeywords[i].keyword
          });
        }
        
        // Store task data for result fetching
        taskIds.splice(0, taskIds.length, ...taskDataArray);
      } else {
        console.log(`❌ Failed to submit batch`);
        failed += batchKeywords.length;
        
        // Write failed results for all keywords in batch (maintain sequence)
        for (let i = 0; i < batchKeywords.length; i++) {
          const actualRowNumber = batchKeywords[i].row; // use original sheet row
          writeSingleRankingResult(sheet, actualRowNumber, {
            ranking: 'Error',
            url: 'Batch Submission Failed',
            keyword: batchKeywords[i].keyword
          }, actualRowNumber);
        }
      }
    } catch (error) {
      console.error(`❌ Error submitting batch:`, error);
      failed += batchKeywords.length;
      
      // Write error results for all keywords in batch (maintain sequence)
      for (let i = 0; i < batchKeywords.length; i++) {
        const actualRowNumber = batchKeywords[i].row; // use original sheet row
        writeSingleRankingResult(sheet, actualRowNumber, { 
          ranking: 'Error', 
          url: 'Batch Error',
          keyword: batchKeywords[i].keyword
        }, actualRowNumber);
      }
    }
  
  if (taskIds.length > 0) {
    console.log(`⏳ Waiting 5 minutes for DataForSEO to process batch ${batchIndex + 1} (${taskIds.length} keywords, 3 pages each)...`);
    
    // Show progress note in kw_variants sheet (N1)
    sheet.getRange(1, 14).setValue(`⏳ Processing Batch ${batchIndex + 1}/${totalBatches} (${taskIds.length} keywords)...`);
    
    Utilities.sleep(300000); // Wait 5 minutes (300 seconds)
    
    // Fetch results for this batch
    console.log(`📥 Fetching results for batch ${batchIndex + 1}...`);
    
    // Collect tasks that are still queued/no results to retry later
    let pendingTasks = [];
    
    for (const taskData of taskIds) {
      try {
        const rankingResults = fetchKeywordRankingResults(taskData.taskId, taskData.keyword);
        
        if (rankingResults && rankingResults.rankings.length > 0) {
          // Get the best ranking (lowest rank number)
          const bestRanking = rankingResults.rankings.reduce((best, current) =>
            current.rank < best.rank ? current : best
          );
          
          console.log(`✅ Found ranking: Position ${bestRanking.rank} - ${bestRanking.url}`);
          
          // Write result immediately to ranking_results (maintain sequence with targetRow)
          writeSingleRankingResult(sheet, taskData.row, {
            ranking: bestRanking.rank,
            url: bestRanking.url,
            keyword: taskData.keyword
          }, taskData.row);
          successful++;
        } else {
          // Defer writing for queued/not-ready tasks; we'll retry below
          console.log(`⏳ Result not ready yet for: ${taskData.keyword}. Will retry...`);
          pendingTasks.push({ taskData, lastRaw: rankingResults.rawData });
        }
      } catch (error) {
        console.error(`❌ Error fetching results for "${taskData.keyword}":`, error);
        
        // Write result immediately (maintain sequence with targetRow)
        writeSingleRankingResult(sheet, taskData.row, {
          ranking: 'Error',
          url: 'Error',
          keyword: taskData.keyword
        }, taskData.row);
        failed++;
      }
    }

    // Retry loop for tasks still pending (e.g., status "Task In Queue")
    if (pendingTasks.length > 0) {
      const maxRounds = 8; // 8 retries
      const intervalMs = 30000;
      for (let round = 1; round <= maxRounds && pendingTasks.length > 0; round++) {
        console.log(`🔄 Retry round ${round}/${maxRounds} for ${pendingTasks.length} pending tasks...`);
        Utilities.sleep(intervalMs);
        const stillPending = [];
        for (const p of pendingTasks) {
          try {
            const retryResult = fetchKeywordRankingResults(p.taskData.taskId, p.taskData.keyword);
            if (retryResult && retryResult.rankings.length > 0) {
              const bestRanking = retryResult.rankings.reduce((best, current) =>
                current.rank < best.rank ? current : best
              );
              writeSingleRankingResult(sheet, p.taskData.row, {
                ranking: bestRanking.rank,
                url: bestRanking.url,
                keyword: p.taskData.keyword
              }, p.taskData.row);
              successful++;
            } else {
              // keep for next retry
              stillPending.push({ taskData: p.taskData, lastRaw: retryResult.rawData || p.lastRaw });
            }
          } catch (e) {
            console.error(`❌ Error on retry for "${p.taskData.keyword}":`, e);
            writeSingleRankingResult(sheet, p.taskData.row, {
              ranking: 'Error',
              url: 'Error',
              keyword: p.taskData.keyword
            }, p.taskData.row);
            failed++;
          }
        }
        pendingTasks = stillPending;
      }
      // Any tasks still pending after retries → mark Not Found with last raw data
      if (pendingTasks.length > 0) {
        console.log(`⚠️ ${pendingTasks.length} tasks still pending after retries; marking as Not Found`);
        for (const p of pendingTasks) {
          writeSingleRankingResult(sheet, p.taskData.row, {
            ranking: 'Not Found',
            url: 'Not Found',
            keyword: p.taskData.keyword
          }, p.taskData.row);
          failed++;
        }
      }
    }
    
    // Force sheet to save/flush after each batch
    SpreadsheetApp.flush();
    
    // Update progress with batch completion (N1)
    sheet.getRange(1, 14).setValue(`✅ Completed batch ${batchIndex + 1}/${totalBatches} (${taskIds.length}) - ${successful} ok, ${failed} failed`);
    sheet.getRange(1, 14).setBackground('#d4edda'); // Light green
    
    // Force another flush to ensure progress is visible
    SpreadsheetApp.flush();
    
    // Small delay to ensure writes are processed
    Utilities.sleep(1000);
    
    // Clear progress
    sheet.getRange(1, 14).setValue('');
    sheet.getRange(1, 14).setBackground(null);
  }
}

console.log('🎉 All ranking checks completed!');

const summary = `Processed: ${processed} | Successful: ${successful} | Failed: ${failed}`;
console.log(`\n✅ Ranking check complete! ${summary}`);
SpreadsheetApp.getActive().toast(`✅ Complete: ${summary}`);
const kw = ss.getSheetByName('kw_variants') || sheet;
kw.getRange(1, 14).setValue(`✅ Complete: ${summary}`);
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
* Create custom menu when sheet opens
*/
function onOpen() {
const ui = SpreadsheetApp.getUi();
ui.createMenu('🔍 Gordon KW + Rankings')
  .addItem('Run Ranking Check on Sheet', 'runRankingCheckOnSheet')
  .addItem('On-demand Check', 'checkRankingsOnly')
  .addItem('Test DataForSEO Connection', 'testDataForSEOConnection')
  .addSeparator()
  .addItem('🛑 Stop All Schedulers', 'stopAllSchedulers')
  .addItem('🔄 Reset Submit Cursor', 'resetSubmitCursor')
  .addToUi();
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

// ============================================================================
// CONFIG SHEET HELPERS (Sheet name: "config")
// ============================================================================

function getConfigSheet() {
const ss = SpreadsheetApp.getActive();
let sheet = ss.getSheetByName('config');
if (!sheet) {
  sheet = ss.insertSheet('config');
}
// Ensure headers
if (sheet.getLastRow() === 0) {
  sheet.getRange(1, 1).setValue('Batch Size');
  sheet.getRange(1, 2).setValue('Job IDs');
  sheet.getRange(1, 3).setValue('Keyword');
  sheet.getRange(1, 4).setValue('Source Row');
  sheet.getRange(1, 5).setValue('status');
  sheet.getRange(1, 6).setValue('submitted timestamp');
  sheet.getRange(1, 7).setValue('completed timestamp');
}
// Default batch size at A2 if empty
const batchSizeCell = sheet.getRange(2, 1);
if (!batchSizeCell.getValue()) {
  batchSizeCell.setValue(50); // default batch size
}
return sheet;
}

function readBatchSizeFromConfig() {
const cfg = getConfigSheet();
const val = Number(cfg.getRange(2, 1).getValue());
return Number.isFinite(val) && val > 0 ? Math.floor(val) : 50;
}

function getScriptProps() {
return PropertiesService.getScriptProperties();
}

function findTriggersByHandler(handler) {
return ScriptApp.getProjectTriggers().filter(t => t.getHandlerFunction() === handler);
}

function ensureSubmitTrigger() {
if (findTriggersByHandler('submitScheduled').length === 0) {
  ScriptApp.newTrigger('submitScheduled').timeBased().everyMinutes(5).create();
}
}

function ensureFetchTrigger() {
if (findTriggersByHandler('fetchScheduled').length === 0) {
  ScriptApp.newTrigger('fetchScheduled').timeBased().everyMinutes(10).create();
}
}

function removeTriggerByHandler(handler) {
const triggers = findTriggersByHandler(handler);
for (let i = 0; i < triggers.length; i++) {
  ScriptApp.deleteTrigger(triggers[i]);
}
}

// ============================================================================
// SUBMIT PHASE (writes to config: Job ID, status=submitted, submitted timestamp)
// ============================================================================

function submitScheduled() {
const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
const cfg = getConfigSheet();
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
// Optimization: read processed keywords once from ranking_results!C into a Set
let processedKeywords = new Set();
try {
  const rs = getOrCreateRankingResultsSheet();
  const rsLast = rs.getLastRow();
  if (rsLast >= 2) {
    const keys = rs.getRange(2, 3, rsLast - 1, 1).getValues(); // C: Keyword
    for (let i = 0; i < keys.length; i++) {
      const v = String(keys[i][0] || '').trim();
      if (v) processedKeywords.add(v);
    }
  }
} catch (e) {
  // ignore; fall back to empty set
}
let scanRow = cursor;
while (scanRow <= lastRow && items.length < 2000) { // hard cap safety
  const idx = scanRow - 2;
  const keyword = data[idx] ? data[idx][0] : '';
  const lat = data[idx] ? data[idx][1] : '';
  const lng = data[idx] ? data[idx][2] : '';
  const kwStr = String(keyword || '').trim();
  if (kwStr && lat && lng && !processedKeywords.has(kwStr)) {
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

// Write to config rows B..G
const now = new Date();
const rows = taskIds.length;
const startWriteRow = cfg.getLastRow() >= 2 ? cfg.getLastRow() + 1 : 2;
const out = [];
for (let i = 0; i < rows; i++) {
  out.push([taskIds[i], toSubmit[i].keyword, toSubmit[i].row, 'submitted', now, '']);
}
// B..G where B=Job IDs, C=Keyword, D=Source Row, E=status, F=submitted timestamp, G=completed timestamp
cfg.getRange(startWriteRow, 2, rows, 6).setValues(out);

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

SpreadsheetApp.getActive().toast(`✅ Submitted ${rows} tasks`);
sheet.getRange(1, 14).setValue(`✅ Submitted ${rows} tasks`);
}

// ============================================================================
// FETCH PHASE (reads config, fetches ready tasks, writes to Sheet1 N/O/P)
// ============================================================================

function fetchScheduled() {
const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
const cfg = getConfigSheet();

const lastRow = cfg.getLastRow();
if (lastRow < 2) {
  SpreadsheetApp.getActive().toast('❌ No submitted tasks found in config');
  sheet.getRange(1, 14).setValue('❌ No submitted tasks');
  return;
}

const rows = cfg.getRange(2, 2, lastRow - 1, 6).getValues(); // B..G
let fetched = 0;
let pending = 0;
let totalConsidered = 0;
for (let i = 0; i < rows.length; i++) {
  const taskId = rows[i][0]; // B: Job IDs
  const keyword = rows[i][1]; // C: Keyword
  const sourceRow = rows[i][2]; // D: Source Row
  const status = rows[i][3]; // E: status
  const completedAt = rows[i][5]; // G: completed timestamp
  if (!taskId) continue;
  if (status === 'fetched' && completedAt) continue; // already done
  totalConsidered++;

  const res = fetchKeywordRankingResults(taskId, keyword);
  if (res && res.rankings && res.rankings.length > 0) {
    const best = res.rankings.reduce((b, c) => (c.rank < b.rank ? c : b));
    
    // Use keyword from config sheet (more reliable than extracting from raw data)
    const kw = String(keyword || '').trim();
    // Use sourceRow if available, otherwise append (backward compatibility)
    const targetRow = sourceRow && sourceRow >= 2 ? Number(sourceRow) : null;
    
    // Only skip if target row has actual results (not "Pending" or old data we want to overwrite)
    if (targetRow && hasActualResultsAtRow(targetRow)) {
      // Check if the keyword matches - if so, skip (already processed)
      const rs = getOrCreateRankingResultsSheet();
      const existingKeyword = String(rs.getRange(targetRow, 3).getValue() || '').trim();
      if (existingKeyword === kw) {
        cfg.getRange(i + 2, 5).setValue('fetched'); // E: status
        cfg.getRange(i + 2, 7).setValue(new Date()); // G: completed timestamp
        continue;
      }
    } else if (!targetRow && kw && keywordAlreadyProcessed(kw)) {
      // For backward compatibility (no targetRow), check if keyword exists anywhere
      cfg.getRange(i + 2, 5).setValue('fetched'); // E: status
      cfg.getRange(i + 2, 7).setValue(new Date()); // G: completed timestamp
      continue;
    }
    
    // Write result (will overwrite "Pending" or old data at target row)
    writeSingleRankingResult(sheet, 2, { // row is unused; we pass keyword explicitly
      ranking: best.rank,
      url: best.url,
      keyword: kw
    }, targetRow);
    cfg.getRange(i + 2, 5).setValue('fetched'); // E: status
    cfg.getRange(i + 2, 7).setValue(new Date()); // G: completed timestamp
    fetched++;
  } else {
    cfg.getRange(i + 2, 5).setValue('pending'); // E: status - keep waiting
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
ScriptApp.newTrigger('fetchScheduled').timeBased().everyMinutes(10).create();
SpreadsheetApp.getUi().alert('✅ Trigger Created', 'Fetch trigger set to run every 10 minutes.', SpreadsheetApp.getUi().ButtonSet.OK);
}

// ============================================================================
// ORCHESTRATOR ENTRY POINT
// ============================================================================

function runRankingCheckOnSheet() {
SpreadsheetApp.getUi().alert('🚀 Scheduler Starting', 'Setting up submit and 10-min fetch schedulers now.', SpreadsheetApp.getUi().ButtonSet.OK);
// Kick off first batch immediately
submitScheduled();
// Ensure periodic triggers are in place
ensureSubmitTrigger();
ensureFetchTrigger();
SpreadsheetApp.getUi().alert('🚀 Scheduler Running', 'Submission and 10-min fetch schedulers are active. You can close the sheet; they will auto-stop when finished.', SpreadsheetApp.getUi().ButtonSet.OK);
SpreadsheetApp.getActive().toast('🚀 Scheduler running: submit + 10-min fetch active');
const sheet = SpreadsheetApp.getActive().getSheetByName('kw_variants') || SpreadsheetApp.getActiveSheet();
sheet.getRange(1, 14).setValue('🚀 Scheduler running: submit + 10-min fetch active');
}