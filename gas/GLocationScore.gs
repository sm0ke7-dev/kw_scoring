/**
 * GLocationScore.gs
 * Location scoring for Gordon's ranking system
 * 
 * Calculates economic value of locations based on:
 * Formula: (income - poverty_line) × population
 * Normalizes across all locations to sum = 1.0
 */

function runGLocationScore() {
  const cfg = getGConfig();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('locations');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      '❌ Sheet Not Found', 
      'Sheet "locations" not found. Please ensure it exists with columns: target, population, income',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw new Error('Sheet "locations" not found');
  }
  
  const values = getGDataRangeValues(sheet);
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Data',
      'No data found in locations sheet',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeGKey(h));
  
  // Find required columns
  const idxTarget = normHeader.indexOf('target');
  const idxPopulation = normHeader.indexOf('population');
  const idxIncome = normHeader.indexOf('income');
  
  if (idxTarget === -1) {
    throw new Error('locations sheet missing "target" column');
  }
  if (idxPopulation === -1) {
    throw new Error('locations sheet missing "population" column');
  }
  if (idxIncome === -1) {
    throw new Error('locations sheet missing "income" column');
  }
  
  const povertyLine = cfg.povertyLine || 32000;
  console.log(`📊 Location Scoring Started`);
  console.log(`📊 Using poverty line: $${povertyLine.toLocaleString()}`);
  console.log(`📊 Processing ${values.length - 1} locations...`);
  
  // Calculate numerical scores
  const numerical = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const pop = Number(row[idxPopulation]) || 0;
    const inc = Number(row[idxIncome]) || 0;
    const val = (inc - povertyLine) * pop;
    numerical.push(val);
  }
  
  // Normalize to sum = 1.0
  const totalSum = sumG(numerical);
  const normalized = [];
  
  if (totalSum > 0) {
    for (let i = 0; i < numerical.length; i++) {
      normalized.push((Number(numerical[i]) || 0) / totalSum);
    }
  } else {
    // All zeros - equal distribution
    console.warn('⚠️ Total sum is zero, using equal distribution');
    for (let i = 0; i < numerical.length; i++) {
      normalized.push(numerical.length > 0 ? 1 / numerical.length : 0);
    }
  }
  
  // Write output
  writeGColumnByHeader(sheet, 'location_score', normalized);
  
  // Format column as decimal (6 places)
  const scoreCol = findGHeaderIndex(sheet, 'location_score') + 1;
  if (scoreCol > 0 && normalized.length > 0) {
    sheet.getRange(2, scoreCol, normalized.length, 1).setNumberFormat('0.000000');
  }
  
  // Validation log (console only)
  const sumCheck = sumG(normalized);
  console.log(`✅ Location scores calculated: ${numerical.length} locations`);
  console.log(`✅ Normalized sum check: ${sumCheck.toFixed(6)} (should be ~1.0)`);
  
  // Find top 5 locations
  const ranked = [];
  for (let i = 0; i < values.length - 1; i++) {
    ranked.push({
      name: values[i + 1][idxTarget],
      score: normalized[i],
      population: values[i + 1][idxPopulation],
      income: values[i + 1][idxIncome]
    });
  }
  ranked.sort((a, b) => b.score - a.score);
  
  console.log('🏆 Top 5 Locations by Economic Score:');
  for (let i = 0; i < Math.min(5, ranked.length); i++) {
    const pct = (ranked[i].score * 100).toFixed(2);
    console.log(`  ${i + 1}. ${ranked[i].name} = ${pct}% (pop: ${ranked[i].population.toLocaleString()}, income: $${ranked[i].income.toLocaleString()})`);
  }
  
  SpreadsheetApp.getActive().toast(`✅ Location scores calculated! Top: ${ranked[0].name} (${(ranked[0].score * 100).toFixed(1)}%)`, '📊 Location Scoring', 5);
}

