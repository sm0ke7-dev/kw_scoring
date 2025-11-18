function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Rank Scoring menu
  ui.createMenu('Rank Scoring')
    .addItem('Run All', 'runAll')
    .addSeparator()
    .addItem('Location Score', 'runLocationScore')
    .addItem('Niche Score', 'runNicheScore')
    .addItem('Keyword Score', 'runKeywordScore')
    .addItem('Ranking Score', 'runGRankingScore')
    .addItem('Final Merge', 'runFinalMerge')
    .addSeparator()
    .addItem('Clear Log', 'clearLog')
    .addToUi();
  
  // Gordon KW + Rankings menu
  ui.createMenu('🔍 Gordon KW + Rankings')
    .addItem('Run Ranking Check on Sheet', 'runRankingCheckOnSheet')
    .addItem('Test DataForSEO Connection', 'testDataForSEOConnection')
    .addSeparator()
    .addItem('🔄 Retry Pending Tasks', 'retryPendingTasks')
    .addSeparator()
    .addItem('🔄 Reset All Ranking Data', 'resetAllRankingData')
    .addSeparator()
    .addItem('🛑 Stop All Schedulers', 'stopAllSchedulers')
    .addItem('🔄 Reset Submit Cursor', 'resetSubmitCursor')
    .addItem('🗑️ Clear Results Sheets', 'clearResultsSheets')
    .addToUi();
}

function runAll() {
  runLocationScore();
  runNicheScore();
  runKeywordScore();
  runGRankingScore();
  runFinalMerge();
}

function clearLog() {
  const ss = getSpreadsheet();
  let logSheet = ss.getSheetByName('Log');
  if (!logSheet) return;
  logSheet.clearContents();
}


