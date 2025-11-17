/**
 * GNicheScore.gs
 * Niche/Service scoring for Gordon's ranking system
 * 
 * Scores services based on keyword planner data
 * Normalizes across all services to sum = 1.0
 */

function runGNicheScore() {
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('niche');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      '❌ Sheet Not Found', 
      'Sheet "niche" not found. Please ensure it exists with columns: Service, Keyword Planner Score',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw new Error('Sheet "niche" not found');
  }
  
  const values = getGDataRangeValues(sheet);
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Data',
      'No data found in niche sheet',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeGKey(h));
  
  // Find required columns
  const idxService = normHeader.indexOf('service');
  const idxScore = normHeader.indexOf('keyword planner score');
  
  if (idxService === -1) {
    throw new Error('niche sheet missing "Service" column');
  }
  if (idxScore === -1) {
    throw new Error('niche sheet missing "Keyword Planner Score" column');
  }
  
  console.log(`📊 Niche Scoring Started`);
  console.log(`📊 Processing ${values.length - 1} services...`);
  
  // Read raw keyword planner scores
  const rawScores = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const score = Number(row[idxScore]) || 0;
    rawScores.push(score);
  }
  
  // Normalize to sum = 1.0
  const normalized = normalizeGVector(rawScores);
  
  // Write output
  writeGColumnByHeader(sheet, 'niche_score', normalized);
  
  // Format column as decimal (6 places)
  const scoreCol = findGHeaderIndex(sheet, 'niche_score') + 1;
  if (scoreCol > 0 && normalized.length > 0) {
    sheet.getRange(2, scoreCol, normalized.length, 1).setNumberFormat('0.000000');
  }
  
  // Validation log (console only)
  const sumCheck = sumG(normalized);
  console.log(`✅ Niche scores calculated: ${normalized.length} services`);
  console.log(`✅ Normalized sum check: ${sumCheck.toFixed(6)} (should be ~1.0)`);
  
  // Find top 5 services
  const ranked = [];
  for (let i = 0; i < values.length - 1; i++) {
    ranked.push({
      name: values[i + 1][idxService],
      score: normalized[i],
      rawScore: rawScores[i]
    });
  }
  ranked.sort((a, b) => b.score - a.score);
  
  console.log('🏆 Top 5 Services by Niche Score:');
  for (let i = 0; i < Math.min(5, ranked.length); i++) {
    const pct = (ranked[i].score * 100).toFixed(2);
    console.log(`  ${i + 1}. ${ranked[i].name} = ${pct}% (raw: ${ranked[i].rawScore})`);
  }
  
  SpreadsheetApp.getActive().toast(
    `✅ Niche scores calculated! Top: ${ranked[0].name} (${(ranked[0].score * 100).toFixed(1)}%)`, 
    '📊 Niche Scoring', 
    5
  );
}

