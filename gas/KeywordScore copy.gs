function runKeywordScore() {
  const cfg = getConfig();
  const keywordSheet = getSheet('Keyword');
  const keywordValues = getDataRangeValues(keywordSheet);
  if (keywordValues.length < 2) return;
  const keywordHeader = keywordValues[0].map(h => String(h));
  const keywordNormHeader = keywordHeader.map(h => normalizeKey(h));
  const idxKeywords = keywordNormHeader.indexOf('keywords');
  if (idxKeywords === -1) throw new Error('Keyword sheet missing Keywords header');

  // Read all keywords from Keyword sheet
  const keywords = [];
  for (let r = 1; r < keywordValues.length; r++) {
    const row = keywordValues[r];
    const kw = row[idxKeywords];
    if (kw !== '' && kw !== null) keywords.push(String(kw));
    else keywords.push('');
  }
  const n = keywords.length;
  if (n === 0) return;

  // Read Rankings sheet to identify core vs geo keywords
  const rankingsSheet = getSheet('Rankings');
  const rankingsValues = getDataRangeValues(rankingsSheet);
  if (rankingsValues.length < 2) {
    // No rankings data - assign all keywords to core bucket
    const coreScore = Number(cfg.partitionSplit) || 0.66;
    const geoScore = 1 - coreScore;
    const coreScores = keywords.map(_ => coreScore / n);
    const geoScores = keywords.map(_ => geoScore / n);
    writeColumnByHeader(keywordSheet, 'keyword core score', coreScores);
    writeColumnByHeader(keywordSheet, 'keyword geo score', geoScores);
    writeColumnByHeader(keywordSheet, 'keyword score', coreScores);
    return;
  }

  const rankingsHeader = rankingsValues[0].map(h => String(h));
  const rankingsNormHeader = rankingsHeader.map(h => normalizeKey(h));
  const idxRankKeyword = rankingsNormHeader.indexOf('keyword');
  const idxRankKwGeo = rankingsNormHeader.indexOf('kw+geo');
  
  if (idxRankKeyword === -1 || idxRankKwGeo === -1) {
    throw new Error('Rankings sheet missing keyword or kw+geo headers');
  }

  // Build sets of keywords that appear as core vs geo in Rankings
  // A keyword can appear as BOTH core and geo (e.g., "raccoon removal" as core, "raccoon removal riverview" as geo)
  const coreKeywordSet = new Set();
  const geoKeywordSet = new Set();
  
  for (let r = 1; r < rankingsValues.length; r++) {
    const row = rankingsValues[r];
    const keyword = String(row[idxRankKeyword] || '').trim();
    const kwGeo = String(row[idxRankKwGeo] || '').trim();
    if (!keyword) continue;
    
    const kwNorm = normalizeKey(keyword);
    if (kwGeo && normalizeKey(keyword) === normalizeKey(kwGeo)) {
      // Core keyword: kw+geo exactly matches keyword
      coreKeywordSet.add(kwNorm);
    } else if (kwGeo) {
      // Geo keyword: kw+geo exists but doesn't match keyword exactly
      geoKeywordSet.add(kwNorm);
    }
  }

  // Count unique keywords in each bucket
  const coreCount = Math.max(coreKeywordSet.size, 1); // Avoid division by zero
  const geoCount = Math.max(geoKeywordSet.size, 1);

  // Calculate equal shares within each bucket
  const partitionSplit = Number(cfg.partitionSplit) || 0.66;
  const coreShare = partitionSplit;
  const geoShare = 1 - partitionSplit;
  
  const coreScorePerKeyword = coreShare / coreCount;
  const geoScorePerKeyword = geoShare / geoCount;

  // Build score arrays
  // Each keyword gets BOTH core and geo scores (FinalMerge will pick which one to use based on kw+geo)
  const coreScores = [];
  const geoScores = [];
  for (let i = 0; i < keywords.length; i++) {
    const kw = keywords[i];
    if (!kw) {
      // Empty keyword - both scores 0
      coreScores.push(0);
      geoScores.push(0);
      continue;
    }
    
    const kwNorm = normalizeKey(kw);
    const isCore = coreKeywordSet.has(kwNorm);
    const isGeo = geoKeywordSet.has(kwNorm);
    
    // Assign scores: keyword gets core score if it appears as core, geo score if it appears as geo
    coreScores.push(isCore ? coreScorePerKeyword : 0);
    geoScores.push(isGeo ? geoScorePerKeyword : 0);
  }

  logSumCheck('Keyword core normalized', coreScores);
  logSumCheck('Keyword geo normalized', geoScores);

  // Write both columns for audit; keep legacy 'keyword score' as core
  writeColumnByHeader(keywordSheet, 'keyword core score', coreScores);
  writeColumnByHeader(keywordSheet, 'keyword geo score', geoScores);
  writeColumnByHeader(keywordSheet, 'keyword score', coreScores);
}

/**
 * Plot decay curve visualization on Keyword sheet
 * Writes curve values to columns so you can create a chart
 */
function plotDecayCurve() {
  const cfg = getConfig();
  const sheet = getSheet('Keyword');
  const values = getDataRangeValues(sheet);
  if (values.length < 2) return;
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeKey(h));
  const idxKeywords = normHeader.indexOf('keywords');
  if (idxKeywords === -1) throw new Error('Keyword sheet missing Keywords header');

  const keywords = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const kw = row[idxKeywords];
    if (kw !== '' && kw !== null) keywords.push(String(kw));
    else keywords.push('');
  }
  const n = keywords.length;
  if (n === 0) return;
  const kappa = Number(cfg.kappa) || 1;

  // Calculate decay curve values (same as in runKeywordScore)
  const positions = [];
  const rawWeights = [];
  const normalizedWeights = [];
  
  for (let i = 0; i < n; i++) {
    const x = n > 1 ? i / (n - 1) : 0;
    const w = 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
    positions.push(i + 1); // Position number (1, 2, 3, ...)
    rawWeights.push(w);
  }
  
  // Normalize the weights
  const total = rawWeights.reduce((a, b) => a + b, 0);
  for (let i = 0; i < rawWeights.length; i++) {
    normalizedWeights.push(total > 0 ? rawWeights[i] / total : 0);
  }

  // Write curve visualization data
  writeColumnByHeader(sheet, 'curve_position', positions);
  writeColumnByHeader(sheet, 'curve_raw_weight', rawWeights);
  writeColumnByHeader(sheet, 'curve_normalized', normalizedWeights);
  
  console.log('Decay curve plotted. Use Insert > Chart to visualize curve_raw_weight or curve_normalized vs curve_position');
}


