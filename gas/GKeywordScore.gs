/**
 * GKeywordScore.gs
 * Keyword scoring for Gordon's ranking system
 * 
 * Applies decay curve to core keywords, grouped by service
 * Each service group gets 0.66 budget distributed with logarithmic decay
 * First keyword in each service gets highest score, last gets lowest
 */

function runGKeywordScore() {
  const cfg = getGConfig();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('keyword');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      '❌ Sheet Not Found', 
      'Sheet "keyword" not found. Please ensure it exists with columns: Services, Keywords',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw new Error('Sheet "keyword" not found');
  }
  
  const values = getGDataRangeValues(sheet);
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Data',
      'No data found in keyword sheet',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeGKey(h));
  
  // Find required columns
  const idxService = normHeader.indexOf('services');
  const idxKeyword = normHeader.indexOf('keywords');
  
  if (idxService === -1) {
    throw new Error('keyword sheet missing "Services" column');
  }
  if (idxKeyword === -1) {
    throw new Error('keyword sheet missing "Keywords" column');
  }
  
  const coreBudget = Number(cfg.partitionSplit) || 0.66;
  const kappa = Number(cfg.kappa) || 1;
  
  console.log(`📊 Keyword Scoring Started`);
  console.log(`📊 Core budget per service: ${coreBudget}`);
  console.log(`📊 Decay curve kappa: ${kappa}`);
  console.log(`📊 Processing ${values.length - 1} keywords...`);
  
  // Group keywords by service
  const serviceGroups = new Map(); // service -> array of {keyword, rowIndex}
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const service = String(row[idxService] || '').trim();
    const keyword = String(row[idxKeyword] || '').trim();
    
    if (!service || !keyword) continue;
    
    if (!serviceGroups.has(service)) {
      serviceGroups.set(service, []);
    }
    serviceGroups.get(service).push({
      keyword: keyword,
      rowIndex: r // 1-based row in sheet (row 2 = index 1 in data array)
    });
  }
  
  console.log(`📊 Found ${serviceGroups.size} service groups`);
  
  // Calculate scores for each service group
  const allScores = new Array(values.length - 1).fill(0); // scores array aligned with data rows
  const serviceValidation = [];
  
  for (const [service, keywords] of serviceGroups.entries()) {
    const N = keywords.length;
    if (N === 0) continue;
    
    console.log(`\n🔹 ${service}: ${N} keywords`);
    
    // Calculate decay weights for this group
    const weights = [];
    for (let i = 0; i < N; i++) {
      const x = ((i + 1) / (N + 1)) * coreBudget; // sample in [0, split] - ensures last keyword never reaches zero
      const weight = 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
      weights.push(weight);
    }
    
    // Normalize weights so they sum to coreBudget (0.66)
    const sumWeights = weights.reduce((a, b) => a + b, 0);
    const scores = [];
    for (let i = 0; i < N; i++) {
      const score = sumWeights > 0 ? (weights[i] / sumWeights) * coreBudget : 0;
      scores.push(score);
      
      // Assign score to correct position in allScores array
      const dataIndex = keywords[i].rowIndex - 1; // Convert to 0-based data array index
      allScores[dataIndex] = score;
      
      // Log first and last keyword in each group
      if (i === 0 || i === N - 1) {
        console.log(`  ${i === 0 ? '(highest)' : '(lowest) '} ${keywords[i].keyword} = ${score.toFixed(6)}`);
      }
    }
    
    // Validate group sum
    const groupSum = scores.reduce((a, b) => a + b, 0);
    serviceValidation.push({
      service: service,
      count: N,
      sum: groupSum
    });
    console.log(`  ✅ Group sum: ${groupSum.toFixed(6)} (target: ${coreBudget.toFixed(2)})`);
  }
  
  // Write output
  writeGColumnByHeader(sheet, 'keyword core score', allScores);
  
  // Format column as decimal (6 places)
  const scoreCol = findGHeaderIndex(sheet, 'keyword core score') + 1;
  if (scoreCol > 0 && allScores.length > 0) {
    sheet.getRange(2, scoreCol, allScores.length, 1).setNumberFormat('0.000000');
  }
  
  // Validation summary
  const grandTotal = allScores.reduce((a, b) => a + b, 0);
  const expectedTotal = coreBudget * serviceGroups.size;
  
  console.log(`\n✅ Keyword scores calculated!`);
  console.log(`✅ Total services: ${serviceGroups.size}`);
  console.log(`✅ Total keywords: ${allScores.length}`);
  console.log(`✅ Grand total: ${grandTotal.toFixed(6)} (expected: ${expectedTotal.toFixed(2)})`);
  console.log(`✅ Difference: ${Math.abs(grandTotal - expectedTotal).toFixed(6)}`);
  
  // Show summary per service
  console.log(`\n📊 Service Summary:`);
  for (const val of serviceValidation) {
    const pct = (val.sum / grandTotal * 100).toFixed(1);
    console.log(`  ${val.service}: ${val.count} keywords, sum=${val.sum.toFixed(6)} (${pct}% of total)`);
  }
  
  SpreadsheetApp.getActive().toast(
    `✅ Keyword scores calculated! ${allScores.length} keywords across ${serviceGroups.size} services`, 
    '📊 Keyword Scoring', 
    5
  );
}

