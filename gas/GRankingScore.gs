/**
 * GRankingScore.gs
 * Ranking and Keyword score calculation for Gordon's system
 * 
 * Adds THREE columns to ranking_results:
 * 1. geo: Location name (joined from kw_variants)
 * 2. rank_score: Converts ranking positions (1-10) to scores using logarithmic decay
 * 3. keyword_score: Core keywords lookup from keyword sheet, geo keywords calculated per location+service
 */

function runGRankingScore() {
  const cfg = getGConfig();
  const ss = SpreadsheetApp.getActive();
  const sheet = ss.getSheetByName('ranking_results');
  
  if (!sheet) {
    SpreadsheetApp.getUi().alert(
      '❌ Sheet Not Found', 
      'Sheet "ranking_results" not found. Please run ranking check first to generate ranking data.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    throw new Error('Sheet "ranking_results" not found');
  }
  
  const values = getGDataRangeValues(sheet);
  if (values.length < 2) {
    SpreadsheetApp.getUi().alert(
      '⚠️ No Data',
      'No ranking data found. Please run ranking check first.',
      SpreadsheetApp.getUi().ButtonSet.OK
    );
    return;
  }
  
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeGKey(h));
  
  // Find required columns
  const idxRank = normHeader.indexOf('ranking position');
  const idxKeyword = normHeader.indexOf('keyword');
  
  if (idxRank === -1) {
    throw new Error('ranking_results sheet missing "Ranking Position" column');
  }
  if (idxKeyword === -1) {
    throw new Error('ranking_results sheet missing "Keyword" column');
  }
  
  const kappa = Number(cfg.rankModificationKappa) || 10;
  
  console.log(`📊 Ranking Score Started`);
  console.log(`📊 Rank modification kappa: ${kappa}`);
  console.log(`📊 Processing ${values.length - 1} ranking entries...`);
  
  if (kappa <= 0) {
    throw new Error('Rank modification kappa must be greater than 0');
  }
  
  // Calculate logarithmic decay scores
  // Normalized score between 0 and 1: rank 1 = 1.0, rank 10 ≈ 0
  const scores = [];
  let rank1Count = 0;
  let rank2_10Count = 0;
  let notRankedCount = 0;
  
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rank = Number(row[idxRank]) || 0;
    let score = 0;
    
    if (rank >= 1 && rank <= 10) {
      // Normalize rank to [0, 1] where 1 -> 0, 10 -> 1
      const x = (rank - 1) / 9;
      // Logarithmic decay: y = 1 - log(1 + kappa * x) / log(1 + kappa)
      score = 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
      // Round to 6 decimals
      score = Number(score.toFixed(6));
      
      if (rank === 1) rank1Count++;
      else rank2_10Count++;
    } else {
      // Ranks outside 1-10 get zero
      score = 0;
      notRankedCount++;
    }
    
    scores.push(score);
  }
  
  // Write output
  writeGColumnByHeader(sheet, 'rank_score', scores);
  
  // Format column as decimal (6 places)
  const scoreCol = findGHeaderIndex(sheet, 'rank_score') + 1;
  if (scoreCol > 0 && scores.length > 0) {
    sheet.getRange(2, scoreCol, scores.length, 1).setNumberFormat('0.000000');
  }
  
  // Calculate stats
  const avgScore = scores.reduce((a, b) => a + b, 0) / scores.length;
  const maxScore = Math.max(...scores);
  const minNonZero = Math.min(...scores.filter(s => s > 0));
  
  console.log(`\n✅ Ranking scores calculated!`);
  console.log(`✅ Total entries: ${scores.length}`);
  console.log(`✅ Rank 1 (score=1.0): ${rank1Count} entries`);
  console.log(`✅ Rank 2-10: ${rank2_10Count} entries`);
  console.log(`✅ Not ranked (>10 or 0): ${notRankedCount} entries`);
  console.log(`\n📊 Score Statistics:`);
  console.log(`   Average: ${avgScore.toFixed(6)}`);
  console.log(`   Max: ${maxScore.toFixed(6)}`);
  console.log(`   Min (non-zero): ${minNonZero === Infinity ? 'N/A' : minNonZero.toFixed(6)}`);
  
  // Show sample rankings
  console.log(`\n📊 Sample Rank → Score Mapping (kappa=${kappa}):`);
  for (let rank = 1; rank <= 10; rank++) {
    const x = (rank - 1) / 9;
    const s = 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
    console.log(`   Rank ${rank}: ${s.toFixed(6)}`);
  }
  
  const pctRanked = ((rank1Count + rank2_10Count) / scores.length * 100).toFixed(1);
  
  // ============================================================================
  // KEYWORD SCORE CALCULATION
  // ============================================================================
  
  console.log(`\n📊 Starting Keyword Score calculation...`);
  
  // Read kw_variants to get service and location for each keyword
  const kwVariantsSheet = ss.getSheetByName('kw_variants');
  if (!kwVariantsSheet) {
    console.warn('⚠️ kw_variants sheet not found, skipping keyword score');
    SpreadsheetApp.getActive().toast(`✅ Ranking scores calculated! ${pctRanked}% ranked`, '📊 Ranking Scoring', 5);
    return;
  }
  
  const kvValues = getGDataRangeValues(kwVariantsSheet);
  if (kvValues.length < 2) {
    console.warn('⚠️ No data in kw_variants, skipping keyword score');
    SpreadsheetApp.getActive().toast(`✅ Ranking scores calculated! ${pctRanked}% ranked`, '📊 Ranking Scoring', 5);
    return;
  }
  
  const kvHeader = kvValues[0].map(h => String(h));
  const kvNormHeader = kvHeader.map(h => normalizeGKey(h));
  const idxKvService = kvNormHeader.indexOf('service');
  const idxKvGeo = kvNormHeader.indexOf('geo'); // Column H: geo (location name)
  const idxKvKeyword = kvNormHeader.indexOf('kw+geo'); // Column J: kw+geo (final keyword with location)
  
  if (idxKvService === -1 || idxKvGeo === -1 || idxKvKeyword === -1) {
    console.warn('⚠️ kw_variants missing required columns (service, geo, kw+geo), skipping keyword score');
    console.warn(`   Found: service=${idxKvService}, geo=${idxKvGeo}, kw+geo=${idxKvKeyword}`);
    SpreadsheetApp.getActive().toast(`✅ Ranking scores calculated! ${pctRanked}% ranked`, '📊 Ranking Scoring', 5);
    return;
  }
  
  // Note: We'll match ranking_results rows to kw_variants rows by ROW INDEX (1-to-1 correspondence)
  // This ensures correct geo assignment even for core keywords that appear in multiple locations
  console.log(`   kw_variants has ${kvValues.length - 1} data rows`);
  console.log(`   Using row-by-row matching (ranking_results row N = kw_variants row N)`);
  
  // Read keyword sheet for core keywords and scores
  const keywordSheet = ss.getSheetByName('keyword');
  const keywordCoreMap = new Map(); // keyword → score
  const coreKeywordsSet = new Set(); // Set of all core keywords (for fast lookup)
  
  if (!keywordSheet) {
    console.warn('⚠️ keyword sheet not found, all keywords will be treated as geo');
  } else {
    const kwsValues = getGDataRangeValues(keywordSheet);
    if (kwsValues.length >= 2) {
      const kwsHeader = kwsValues[0].map(h => String(h));
      const kwsNormHeader = kwsHeader.map(h => normalizeGKey(h));
      const idxKwKeyword = kwsNormHeader.indexOf('keywords');
      const idxKwCore = kwsNormHeader.indexOf('keyword core score');
      if (idxKwKeyword !== -1 && idxKwCore !== -1) {
        for (let r = 1; r < kwsValues.length; r++) {
          const row = kwsValues[r];
          const kw = String(row[idxKwKeyword] || '').trim();
          const key = normalizeGKey(kw);
          const coreVal = Number(row[idxKwCore]) || 0;
          if (key) {
            keywordCoreMap.set(key, coreVal);
            coreKeywordsSet.add(key); // Add to Set for fast detection
          }
        }
        console.log(`   Found ${keywordCoreMap.size} core keywords in keyword sheet`);
      }
    }
  }
  
  // Build location tokens for geo detection
  const locationTokens = new Set();
  const locationsSheet = ss.getSheetByName('locations');
  if (locationsSheet) {
    const locValues = getGDataRangeValues(locationsSheet);
    if (locValues.length >= 2) {
      const locHeader = locValues[0].map(h => String(h));
      const locNormHeader = locHeader.map(h => normalizeGKey(h));
      const idxTarget = locNormHeader.indexOf('target');
      if (idxTarget !== -1) {
        for (let r = 1; r < locValues.length; r++) {
          const row = locValues[r];
          const t = String(row[idxTarget] || '');
          if (t) locationTokens.add(normalizeGKey(t));
        }
        console.log(`   Found ${locationTokens.size} location tokens for geo detection`);
      }
    }
  }
  
  function isGeoString(str) {
    const s = normalizeGKey(str || '');
    if (!s) return false;
    for (const tok of locationTokens) {
      if (tok && s.indexOf(tok) !== -1) return true;
    }
    return false;
  }
  
  // Count unique geo keywords per location+service to calculate geo scores
  const geoBudget = 1 - (Number(cfg.partitionSplit) || 0.66); // 0.34
  const locationServiceGeoScoreMap = new Map();
  const locationServiceKwGeoMap = new Map();
  
  // Use row-by-row matching to count geo keywords per location+service
  for (let r = 1; r < values.length; r++) {
    const keyword = String(values[r][idxKeyword] || '').trim();
    if (!keyword) continue;
    
    // Check if this row exists in kw_variants
    if (r >= kvValues.length) {
      console.warn(`⚠️ Row ${r + 1} in ranking_results but no matching row in kw_variants`);
      continue;
    }
    
    const kvRow = kvValues[r]; // Get matching row from kw_variants
    const service = String(kvRow[idxKvService] || '').trim();
    const location = String(kvRow[idxKvGeo] || '').trim();
    
    if (!service || !location) continue;
    
    // Only count GEO keywords (NOT in core keyword sheet)
    const kwNorm = normalizeGKey(keyword);
    const isCoreKeyword = coreKeywordsSet.has(kwNorm);
    if (isCoreKeyword) continue; // Skip core keywords
    
    const serviceNorm = normalizeGKey(service);
    const locationNorm = normalizeGKey(location);
    const compositeKey = locationNorm + '|' + serviceNorm;
    
    if (!locationServiceKwGeoMap.has(compositeKey)) {
      locationServiceKwGeoMap.set(compositeKey, new Set());
    }
    locationServiceKwGeoMap.get(compositeKey).add(kwNorm);
  }
  
  // Calculate geo score per location+service
  for (const [compositeKey, kwGeoSet] of locationServiceKwGeoMap.entries()) {
    const count = kwGeoSet.size;
    const geoScore = count > 0 ? geoBudget / count : 0;
    locationServiceGeoScoreMap.set(compositeKey, geoScore);
    const [location, service] = compositeKey.split('|');
    console.log(`   ${location} + ${service}: ${count} geo keywords, score per keyword = ${geoScore.toFixed(6)}`);
  }
  
  console.log(`   Calculated geo scores for ${locationServiceGeoScoreMap.size} location+service combinations`);
  
  // Build geo column and keyword scores for each row
  const geoColumn = [];
  const keywordScores = [];
  let coreCount = 0;
  let geoCount = 0;
  let missingCount = 0;
  let geoMatchesLocation = 0; // Validation: geo keywords that contain location name
  
  for (let r = 1; r < values.length; r++) {
    const keyword = String(values[r][idxKeyword] || '').trim();
    
    if (!keyword) {
      geoColumn.push('');
      keywordScores.push(0);
      missingCount++;
      continue;
    }
    
    // ROW-BY-ROW MATCHING: Get geo and service from SAME row number in kw_variants
    if (r >= kvValues.length) {
      console.warn(`⚠️ Row ${r + 1} in ranking_results but no matching row in kw_variants`);
      geoColumn.push('');
      keywordScores.push(0);
      missingCount++;
      continue;
    }
    
    const kvRow = kvValues[r]; // Get matching row from kw_variants
    const service = String(kvRow[idxKvService] || '').trim();
    const location = String(kvRow[idxKvGeo] || '').trim();
    
    if (!service || !location) {
      geoColumn.push('');
      keywordScores.push(0);
      missingCount++;
      continue;
    }
    
    // Populate geo for ALL keywords (from matched kw_variants row)
    geoColumn.push(location);
    
    // PRIMARY DETECTION: Is this keyword in the keyword sheet?
    const kwNorm = normalizeGKey(keyword);
    const isCoreKeyword = coreKeywordsSet.has(kwNorm);
    
    let keywordScore = 0;
    if (isCoreKeyword) {
      // CORE keyword: lookup score from keyword sheet
      keywordScore = keywordCoreMap.get(kwNorm) || 0;
      coreCount++;
    } else {
      // GEO keyword: calculate per-location-service geo score
      const serviceNorm = normalizeGKey(service);
      const locationNorm = normalizeGKey(location);
      const compositeKey = locationNorm + '|' + serviceNorm;
      keywordScore = locationServiceGeoScoreMap.get(compositeKey) || 0;
      geoCount++;
      
      // Validation: Check if geo keyword actually contains location name
      if (isGeoString(keyword)) {
        geoMatchesLocation++;
      }
    }
    
    keywordScores.push(keywordScore);
  }
  
  // Write geo column to ranking_results (Column D)
  writeGColumnByHeader(sheet, 'geo', geoColumn);
  console.log(`   Added geo column (${geoColumn.filter(g => g).length} geo keywords with location)`);
  
  // Write keyword scores
  writeGColumnByHeader(sheet, 'keyword_score', keywordScores);
  
  // Format column as decimal (6 places)
  const kwScoreCol = findGHeaderIndex(sheet, 'keyword_score') + 1;
  if (kwScoreCol > 0 && keywordScores.length > 0) {
    sheet.getRange(2, kwScoreCol, keywordScores.length, 1).setNumberFormat('0.000000');
  }
  
  console.log(`\n✅ Keyword scores calculated!`);
  console.log(`   Core keywords: ${coreCount} (found in keyword sheet)`);
  console.log(`   Geo keywords: ${geoCount} (NOT in keyword sheet)`);
  console.log(`   Missing/Not found: ${missingCount}`);
  console.log(`\n📊 Validation:`);
  console.log(`   Geo keywords containing location name: ${geoMatchesLocation}/${geoCount} (${geoCount > 0 ? (geoMatchesLocation/geoCount*100).toFixed(1) : 0}%)`);
  
  // ============================================================================
  // FINAL SCORE & OPPORTUNITY GAP
  // ============================================================================
  console.log(`\n🎯 Calculating Final Scores & Opportunity Gaps...`);
  
  // Read location scores from 'locations' sheet
  const locationSheet = ss.getSheetByName('locations');
  if (!locationSheet) {
    console.warn('⚠️ locations sheet not found - skipping final score calculation');
    SpreadsheetApp.getActive().toast(
      `✅ Scores calculated! Ranking: ${pctRanked}% ranked | Keywords: ${coreCount} core, ${geoCount} geo`, 
      '📊 Ranking + Keyword Scoring', 
      5
    );
    return;
  }
  
  const locValues = getGDataRangeValues(locationSheet);
  const idxLocTarget = findGHeaderIndex(locationSheet, 'target');
  const idxLocScore = findGHeaderIndex(locationSheet, 'location_score');
  
  if (idxLocTarget === -1 || idxLocScore === -1) {
    console.warn('⚠️ locations sheet missing required columns - skipping final score calculation');
    SpreadsheetApp.getActive().toast(
      `✅ Scores calculated! Ranking: ${pctRanked}% ranked | Keywords: ${coreCount} core, ${geoCount} geo`, 
      '📊 Ranking + Keyword Scoring', 
      5
    );
    return;
  }
  
  // Build location → location_score map
  const locationScoreMap = new Map();
  for (let r = 1; r < locValues.length; r++) {
    const loc = String(locValues[r][idxLocTarget] || '').trim();
    const score = Number(locValues[r][idxLocScore]) || 0;
    if (loc) {
      locationScoreMap.set(normalizeGKey(loc), score);
    }
  }
  console.log(`   Loaded ${locationScoreMap.size} location scores`);
  
  // Read niche scores from 'niche' sheet
  const nicheSheet = ss.getSheetByName('niche');
  if (!nicheSheet) {
    console.warn('⚠️ niche sheet not found - skipping final score calculation');
    SpreadsheetApp.getActive().toast(
      `✅ Scores calculated! Ranking: ${pctRanked}% ranked | Keywords: ${coreCount} core, ${geoCount} geo`, 
      '📊 Ranking + Keyword Scoring', 
      5
    );
    return;
  }
  
  const nicheValues = getGDataRangeValues(nicheSheet);
  const idxNicheService = findGHeaderIndex(nicheSheet, 'Service');
  const idxNicheScore = findGHeaderIndex(nicheSheet, 'niche_score');
  
  if (idxNicheService === -1 || idxNicheScore === -1) {
    console.warn('⚠️ niche sheet missing required columns - skipping final score calculation');
    SpreadsheetApp.getActive().toast(
      `✅ Scores calculated! Ranking: ${pctRanked}% ranked | Keywords: ${coreCount} core, ${geoCount} geo`, 
      '📊 Ranking + Keyword Scoring', 
      5
    );
    return;
  }
  
  // Build service → niche_score map
  const nicheScoreMap = new Map();
  for (let r = 1; r < nicheValues.length; r++) {
    const service = String(nicheValues[r][idxNicheService] || '').trim();
    const score = Number(nicheValues[r][idxNicheScore]) || 0;
    if (service) {
      nicheScoreMap.set(normalizeGKey(service), score);
    }
  }
  console.log(`   Loaded ${nicheScoreMap.size} niche scores`);
  
  // Calculate FINAL_SCORE, Ideal, and GAP for each row
  const finalScores = [];
  const idealScores = [];
  const gapScores = [];
  let finalScoreSum = 0;
  let idealScoreSum = 0;
  let gapSum = 0;
  
  for (let r = 1; r < values.length; r++) {
    // Skip if missing data in row
    if (r >= geoColumn.length || r >= keywordScores.length || r >= scores.length) {
      finalScores.push(0);
      idealScores.push(0);
      gapScores.push(0);
      continue;
    }
    
    // Get geo (location) from Column D
    const location = geoColumn[r - 1]; // Adjusted index
    const locationNorm = normalizeGKey(location);
    const locationScore = locationScoreMap.get(locationNorm) || 0;
    
    // Get service from kw_variants (same row)
    let service = '';
    if (r < kvValues.length) {
      service = String(kvValues[r][idxKvService] || '').trim();
    }
    const serviceNorm = normalizeGKey(service);
    const nicheScore = nicheScoreMap.get(serviceNorm) || 0;
    
    // Get keyword_score and rank_score (already calculated)
    const keywordScore = keywordScores[r - 1] || 0;
    const rankScore = scores[r - 1] || 0; // 'scores' contains the rank scores
    
    // Calculate Ideal Score = Location × Niche × Keyword × 1.0 (perfect rank)
    const idealScore = locationScore * nicheScore * keywordScore * 1.0;
    idealScores.push(idealScore);
    idealScoreSum += idealScore;
    
    // Calculate FINAL_SCORE = Location × Niche × Keyword × Ranking (actual)
    const finalScore = locationScore * nicheScore * keywordScore * rankScore;
    finalScores.push(finalScore);
    finalScoreSum += finalScore;
    
    // Calculate GAP = Ideal - FINAL_SCORE
    const gap = idealScore - finalScore;
    gapScores.push(gap);
    gapSum += gap;
  }
  
  // Write FINAL_SCORE column (first)
  writeGColumnByHeader(sheet, 'FINAL_SCORE', finalScores);
  const finalScoreColIdx = findGHeaderIndex(sheet, 'FINAL_SCORE') + 1;
  if (finalScoreColIdx > 0 && finalScores.length > 0) {
    sheet.getRange(2, finalScoreColIdx, finalScores.length, 1)
      .setNumberFormat('0.000000')
      .setBackground('#d9ead3'); // Light green
  }
  
  // Write Ideal column (second)
  writeGColumnByHeader(sheet, 'Ideal', idealScores);
  const idealColIdx = findGHeaderIndex(sheet, 'Ideal') + 1;
  if (idealColIdx > 0 && idealScores.length > 0) {
    sheet.getRange(2, idealColIdx, idealScores.length, 1)
      .setNumberFormat('0.000000')
      .setBackground('#e0e0e0'); // Light gray
  }
  
  // Write GAP column (third)
  writeGColumnByHeader(sheet, 'GAP', gapScores);
  const gapColIdx = findGHeaderIndex(sheet, 'GAP') + 1;
  if (gapColIdx > 0 && gapScores.length > 0) {
    sheet.getRange(2, gapColIdx, gapScores.length, 1)
      .setNumberFormat('0.000000')
      .setBackground('#fff2cc'); // Light yellow
  }
  
  console.log(`\n✅ Final scores calculated!`);
  console.log(`   Sum of FINAL_SCORE: ${finalScoreSum.toFixed(6)}`);
  console.log(`   Sum of Ideal: ${idealScoreSum.toFixed(6)}`);
  console.log(`   Sum of GAP: ${gapSum.toFixed(6)}`);
  console.log(`   Top 5 opportunities by GAP:`);
  
  // Show top 5 opportunities
  const topOpportunities = finalScores
    .map((score, idx) => ({
      keyword: String(values[idx + 1][idxKeyword] || ''),
      finalScore: score,
      ideal: idealScores[idx],
      gap: gapScores[idx]
    }))
    .filter(item => item.keyword)
    .sort((a, b) => b.gap - a.gap)
    .slice(0, 5);
  
  topOpportunities.forEach((item, i) => {
    console.log(`   ${i + 1}. ${item.keyword} (GAP: ${item.gap.toFixed(6)}, Final: ${item.finalScore.toFixed(6)}, Ideal: ${item.ideal.toFixed(6)})`);
  });
  
  SpreadsheetApp.getActive().toast(
    `✅ Complete! Ranking: ${pctRanked}% | Keywords: ${coreCount} core, ${geoCount} geo | Final scores calculated!`, 
    '📊 Full Scoring Complete', 
    5
  );
}

