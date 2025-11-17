/**
 * GMenu.gs
 * Unified menu for Gordon's complete system (Rankings + Scoring)
 * NOTE: This is the ONLY onOpen() function - the one in Gordon_KW_pop_rank.js has been removed
 */

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  // Gordon KW + Rankings menu (from Gordon_KW_pop_rank.js)
  ui.createMenu('🔍 Gordon KW + Rankings')
    .addItem('Run Ranking Check on Sheet', 'runRankingCheckOnSheet')
    .addItem('On-demand Check', 'checkRankingsOnly')
    .addItem('Test DataForSEO Connection', 'testDataForSEOConnection')
    .addSeparator()
    .addItem('🛑 Stop All Schedulers', 'stopAllSchedulers')
    .addItem('🔄 Reset Submit Cursor', 'resetSubmitCursor')
    .addToUi();
  
  // Gordon Scoring menu (new scoring system)
  ui.createMenu('📊 Gordon Scoring')
    .addItem('▶️ Run All Scores', 'runGAllScores')
    .addSeparator()
    .addItem('Location Score', 'runGLocationScore')
    .addItem('Niche Score', 'runGNicheScore')
    .addItem('Keyword Score', 'runGKeywordScore')
    .addItem('Ranking Score', 'runGRankingScore')
    .addItem('Final Merge (Coming Soon)', 'comingSoonG')
    .addToUi();
}

function runGAllScores() {
  try {
    SpreadsheetApp.getActive().toast('🚀 Starting scoring calculations...', '📊 Gordon Scoring', 3);
    
    runGLocationScore();
    runGNicheScore();
    runGKeywordScore();
    runGRankingScore();
    // runGFinalMerge();     // Coming in Module 5
    
    SpreadsheetApp.getActive().toast('✅ All scores calculated!', '📊 Gordon Scoring', 5);
  } catch (error) {
    console.error('❌ Error in runGAllScores:', error);
    SpreadsheetApp.getUi().alert(
      '❌ Scoring Error',
      'Error: ' + error.message,
      SpreadsheetApp.getUi().ButtonSet.OK
    );
  }
}


function comingSoonG() {
  SpreadsheetApp.getUi().alert(
    'ℹ️ Coming Soon',
    'This module will be available in the next phase. Currently available:\n\n✅ Location Score\n✅ Niche Score\n✅ Keyword Score\n✅ Ranking Score\n\nStay tuned!',
    SpreadsheetApp.getUi().ButtonSet.OK
  );
}

