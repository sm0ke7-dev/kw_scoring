function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Rank Scoring')
    .addItem('Run All', 'runAll')
    .addSeparator()
    .addItem('Location Score', 'runLocationScore')
    .addItem('Niche Score', 'runNicheScore')
    .addItem('Keyword Score', 'runKeywordScore')
    .addItem('Ranking Score', 'runRankingScore')
    .addItem('Final Merge', 'runFinalMerge')
    .addSeparator()
    .addItem('Plot Decay Curve', 'plotDecayCurve')
    .addSeparator()
    .addItem('Clear Log', 'clearLog')
    .addToUi();
}

function runAll() {
  runLocationScore();
  runNicheScore();
  runKeywordScore();
  runRankingScore();
  runFinalMerge();
}

function clearLog() {
  const ss = getSpreadsheet();
  let logSheet = ss.getSheetByName('Log');
  if (!logSheet) return;
  logSheet.clearContents();
}


