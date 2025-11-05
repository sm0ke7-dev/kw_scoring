function runFinalMerge() {
  const ss = getSpreadsheet();
  const rankings = getSheet('Rankings');
  const rValues = getDataRangeValues(rankings);
  if (rValues.length < 2) return;
  const rHeader = rValues[0].map(String);
  const rNorm = rHeader.map(h => normalizeKey(h));

  const idxService = rNorm.indexOf('service');
  const idxGeo = rNorm.indexOf('geo');
  const idxKeyword = rNorm.indexOf('keyword');
  const idxKwGeo = rNorm.indexOf('kw+geo');
  const idxRankMod = rNorm.indexOf('rank score');
  if ([idxService, idxGeo, idxKeyword].some(i => i === -1)) {
    throw new Error('Rankings sheet missing service/geo/keyword headers');
  }

  // Build lookup maps
  const locationSheet = getSheet('Location');
  const lValues = getDataRangeValues(locationSheet);
  const lHeader = lValues[0].map(String);
  const lNorm = lHeader.map(h => normalizeKey(h));
  const idxTarget = lNorm.indexOf('target');
  const idxLocNorm = lNorm.indexOf('normalized score');
  if (idxTarget === -1 || idxLocNorm === -1) throw new Error('Location sheet missing target/normalized score');
  const locMap = new Map();
  for (let r = 1; r < lValues.length; r++) {
    const row = lValues[r];
    locMap.set(normalizeKey(row[idxTarget]), Number(row[idxLocNorm]) || 0);
  }

  const nicheSheet = getSheet('Niche');
  const nValues = getDataRangeValues(nicheSheet);
  const nHeader = nValues[0].map(String);
  const nNorm = nHeader.map(h => normalizeKey(h));
  const idxNService = nNorm.indexOf('service');
  const idxNNorm = nNorm.indexOf('normalization');
  if (idxNService === -1 || idxNNorm === -1) throw new Error('Niche sheet missing Service/Normalization');
  const nicheMap = new Map();
  for (let r = 1; r < nValues.length; r++) {
    const row = nValues[r];
    nicheMap.set(normalizeKey(row[idxNService]), Number(row[idxNNorm]) || 0);
  }
  
  // Diagnostic: log niche map contents
  console.log('Debug - Niche Map:');
  console.log('  Total entries: ' + nicheMap.size);
  const nicheKeys = Array.from(nicheMap.keys());
  console.log('  First 3 service keys in Niche sheet:');
  for (let i = 0; i < Math.min(3, nicheKeys.length); i++) {
    console.log('    "' + nicheKeys[i] + '" = ' + nicheMap.get(nicheKeys[i]));
  }

  const keywordSheet = getSheet('Keyword');
  const kValues = getDataRangeValues(keywordSheet);
  const kHeader = kValues[0].map(String);
  const kNorm = kHeader.map(h => normalizeKey(h));
  const idxKKeyword = kNorm.indexOf('keywords');
  const idxKCore = kNorm.indexOf('keyword core score');
  const idxKGeo = kNorm.indexOf('keyword geo score');
  if (idxKKeyword === -1) throw new Error('Keyword sheet missing Keywords');
  if (idxKCore === -1 || idxKGeo === -1) throw new Error('Keyword sheet missing "keyword core score" and/or "keyword geo score" headers');
  const keywordCoreMap = new Map();
  const keywordGeoMap = new Map();
  for (let r = 1; r < kValues.length; r++) {
    const row = kValues[r];
    const key = normalizeKey(row[idxKKeyword]);
    const coreVal = Number(row[idxKCore]) || 0;
    const geoVal = Number(row[idxKGeo]) || 0;
    keywordCoreMap.set(key, coreVal);
    keywordGeoMap.set(key, geoVal);
  }

  // Build a set of location tokens for geo detection
  const locationTokens = new Set();
  for (let r = 1; r < lValues.length; r++) {
    const row = lValues[r];
    const t = String(row[idxTarget] || '');
    if (t) locationTokens.add(normalizeKey(t));
  }

  function isGeoString(str) {
    const s = normalizeKey(str || '');
    if (!s) return false;
    for (const tok of locationTokens) {
      if (tok && s.indexOf(tok) !== -1) return true;
    }
    return false;
  }

  // Build final rows
  const outRows = [];
  const finalRaw = [];
  const rankingKeywordScores = [];
  const rankingFinalScores = [];
  const sampleLookups = []; // For diagnostic logging
  for (let r = 1; r < rValues.length; r++) {
    const row = rValues[r];
    const service = row[idxService];
    const geo = row[idxGeo];
    const keyword = row[idxKeyword];
    const kwPlusGeo = idxKwGeo !== -1 ? row[idxKwGeo] : '';
    const locationScore = locMap.get(normalizeKey(geo)) || 0;
    const serviceNorm = normalizeKey(service);
    const nicheScore = nicheMap.get(serviceNorm) || 0;
    
    // Diagnostic: collect first 3 lookups
    if (sampleLookups.length < 3) {
      sampleLookups.push({
        raw: String(service || ''),
        normalized: serviceNorm,
        found: nicheMap.has(serviceNorm),
        score: nicheScore
      });
    }
    const useGeo = isGeoString(kwPlusGeo);
    const kKey = normalizeKey(keyword);
    let keywordScore = 0;
    if (useGeo) {
      if (!keywordGeoMap.has(kKey)) {
        console.warn('Keyword geo score missing for "' + String(keyword || '') + '"; defaulting to 0');
        keywordScore = 0;
      } else {
        keywordScore = Number(keywordGeoMap.get(kKey)) || 0;
      }
    } else {
      if (!keywordCoreMap.has(kKey)) {
        console.warn('Keyword core score missing for "' + String(keyword || '') + '"; defaulting to 0');
        keywordScore = 0;
      } else {
        keywordScore = Number(keywordCoreMap.get(kKey)) || 0;
      }
    }
    const rankingScore = idxRankMod !== -1 ? (Number(row[idxRankMod]) || 0) : 0;
    // For Rankings tab audit columns
    rankingKeywordScores.push(keywordScore);
    rankingFinalScores.push(keywordScore * rankingScore);
    const raw = locationScore * nicheScore * keywordScore * rankingScore;
    finalRaw.push(raw);
    outRows.push([
      service,
      geo,
      keyword,
      locationScore,
      nicheScore,
      keywordScore,
      rankingScore,
      raw,
      0 // placeholder for final_normalized
    ]);
  }
  
  // Diagnostic: log sample lookups
  console.log('Debug - Sample Niche Lookups from Rankings:');
  for (let i = 0; i < sampleLookups.length; i++) {
    const lookup = sampleLookups[i];
    console.log('  Row ' + (i + 1) + ':');
    console.log('    Raw service: "' + lookup.raw + '"');
    console.log('    Normalized: "' + lookup.normalized + '"');
    console.log('    Found in map: ' + (lookup.found ? 'YES' : 'NO'));
    console.log('    Niche score: ' + lookup.score);
  }
  
  const finalNorm = normalizeVector(finalRaw);
  logSumCheck('Final normalized', finalNorm);
  for (let i = 0; i < outRows.length; i++) outRows[i][8] = finalNorm[i];

  // Write Rankings audit columns: keyword score and final score (keyword * ranking modifier)
  if (rankingKeywordScores.length) {
    writeColumnByHeader(rankings, 'keyword score', rankingKeywordScores);
  }
  if (rankingFinalScores.length) {
    writeColumnByHeader(rankings, 'final score', rankingFinalScores);
  }

  // Write to Final sheet
  const finalHeader = ['service','geo','keyword','location_score','niche_score','keyword_score','ranking_score','final_raw','final_normalized'];
  let finalSheet = ss.getSheetByName('Final');
  if (!finalSheet) finalSheet = ss.insertSheet('Final');
  finalSheet.clearContents();
  finalSheet.getRange(1, 1, 1, finalHeader.length).setValues([finalHeader]);
  if (outRows.length > 0) {
    finalSheet.getRange(2, 1, outRows.length, finalHeader.length).setValues(outRows);
    // shade entire Final output table (all columns are outputs)
    try {
      var bg = OUTPUT_BG_COLOR || '#FFF2CC';
      finalSheet.getRange(1, 1, outRows.length + 1, finalHeader.length).setBackground(bg);
    } catch (e) {
      // ignore styling errors
    }
  }
}

/**
 * Visualize final scores - shows top keywords by final_normalized value
 * Creates sortable data for charting most important keywords to target
 */
function visualizeFinalScores() {
  const ss = getSpreadsheet();
  const finalSheet = getSheet('Final');
  const values = getDataRangeValues(finalSheet);
  if (values.length < 2) {
    console.log('No Final data to visualize. Run Final Merge first.');
    return;
  }

  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeKey(h));
  
  const idxService = normHeader.indexOf('service');
  const idxGeo = normHeader.indexOf('geo');
  const idxKeyword = normHeader.indexOf('keyword');
  const idxFinalNorm = normHeader.indexOf('final_normalized');
  
  if ([idxService, idxGeo, idxKeyword, idxFinalNorm].some(i => i === -1)) {
    throw new Error('Final sheet missing required columns');
  }

  // Build array of rows with final_normalized scores
  const rows = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const service = String(row[idxService] || '');
    const geo = String(row[idxGeo] || '');
    const keyword = String(row[idxKeyword] || '');
    const finalNorm = Number(row[idxFinalNorm]) || 0;
    
    if (service || geo || keyword) {
      rows.push({
        service: service,
        geo: geo,
        keyword: keyword,
        final_normalized: finalNorm,
        label: service + ' | ' + geo + ' | ' + keyword
      });
    }
  }

  // Sort by final_normalized (descending)
  rows.sort((a, b) => b.final_normalized - a.final_normalized);

  // Write visualization columns to Final sheet
  const positions = [];
  const scores = [];
  const labels = [];
  
  for (let i = 0; i < rows.length; i++) {
    positions.push(i + 1);
    scores.push(rows[i].final_normalized);
    labels.push(rows[i].label);
  }

  // Create or get visualization sheet
  let vizSheet = ss.getSheetByName('Final Visualization');
  if (!vizSheet) {
    vizSheet = ss.insertSheet('Final Visualization');
  } else {
    vizSheet.clearContents();
  }
  
  // Write visualization data to new sheet
  const vizHeader = ['Position', 'Final Score', 'Service | Geo | Keyword'];
  vizSheet.getRange(1, 1, 1, 3).setValues([vizHeader]);
  
  const vizData = [];
  for (let i = 0; i < rows.length; i++) {
    vizData.push([positions[i], scores[i], labels[i]]);
  }
  
  if (vizData.length > 0) {
    vizSheet.getRange(2, 1, vizData.length, 3).setValues(vizData);
    
    // Format header
    vizSheet.getRange(1, 1, 1, 3).setFontWeight('bold').setBackground('#E8F0FE');
    
    // Automatically create a chart on the visualization sheet
    try {
      const numRows = Math.min(rows.length, 100); // Limit to top 100 for readability
      const chartRange = vizSheet.getRange(1, 1, numRows + 1, 2); // Position and score columns
      
      const chartBuilder = vizSheet.newChart()
        .setChartType(Charts.ChartType.LINE)
        .addRange(chartRange)
        .setPosition(2, 5, 0, 0) // Position chart next to data
        .setOption('title', 'Top Keywords by Final Score (Most Important to Target)')
        .setOption('hAxis.title', 'Position (Rank)')
        .setOption('vAxis.title', 'Final Normalized Score')
        .setOption('legend.position', 'none')
        .setOption('pointSize', 3)
        .setOption('lineWidth', 2)
        .setOption('curveType', 'function'); // Smooth curve
      
      vizSheet.insertChart(chartBuilder.build());
      console.log('Chart created on Final Visualization sheet');
    } catch (e) {
      console.warn('Could not create chart automatically:', e.message);
    }
  }
  
  console.log('Top 5 keywords to target:');
  for (let i = 0; i < Math.min(5, rows.length); i++) {
    console.log((i + 1) + '. ' + rows[i].label + ' = ' + rows[i].final_normalized.toFixed(6));
  }
}


