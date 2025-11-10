function runKeywordScore() {
  const cfg = getConfig();
  const keywordSheet = getSheet('Keyword');
  const keywordValues = getDataRangeValues(keywordSheet);
  if (keywordValues.length < 2) return;
  const keywordHeader = keywordValues[0].map(h => String(h));
  const keywordNormHeader = keywordHeader.map(h => normalizeKey(h));
  const idxKeywords = keywordNormHeader.indexOf('keywords');
  const idxServiceCol = (function() {
    const a = keywordNormHeader.indexOf('services');
    const b = keywordNormHeader.indexOf('service');
    return a !== -1 ? a : b;
  })();
  if (idxKeywords === -1) throw new Error('Keyword sheet missing Keywords header');

  // Read all keywords (and services) from Keyword sheet
  const keywords = [];
  const services = [];
  for (let r = 1; r < keywordValues.length; r++) {
    const row = keywordValues[r];
    const kw = row[idxKeywords];
    if (kw !== '' && kw !== null) keywords.push(String(kw));
    else keywords.push('');
    const svc = idxServiceCol !== -1 ? row[idxServiceCol] : '';
    services.push(String(svc || ''));
  }
  const n = keywords.length;
  if (n === 0) return;

  // Read Rankings sheet to identify core vs geo keywords
  const rankingsSheet = getSheet('Rankings');
  const rankingsValues = getDataRangeValues(rankingsSheet);
  if (rankingsValues.length < 2) {
    // No rankings data - assign equal shares across all keywords for core only
    const split = Number(cfg.partitionSplit) || 0.66;
    const coreBudget = split;
    const coreScores = keywords.map(_ => (n > 0 ? coreBudget / n : 0));
    writeColumnByHeader(keywordSheet, 'keyword core score', coreScores);
    return;
  }

  const rankingsHeader = rankingsValues[0].map(h => String(h));
  const rankingsNormHeader = rankingsHeader.map(h => normalizeKey(h));
  const idxRankKeyword = rankingsNormHeader.indexOf('keyword');
  const idxRankKwGeo = rankingsNormHeader.indexOf('kw+geo');
  
  if (idxRankKeyword === -1 || idxRankKwGeo === -1) {
    throw new Error('Rankings sheet missing keyword or kw+geo headers');
  }

  // Build set of keywords that appear as core in Rankings
  // Core keywords: kw+geo exactly matches keyword (no location modifier)
  const coreKeywordSet = new Set();
  
  for (let r = 1; r < rankingsValues.length; r++) {
    const row = rankingsValues[r];
    const keyword = String(row[idxRankKeyword] || '').trim();
    const kwGeo = String(row[idxRankKwGeo] || '').trim();
    if (!keyword) continue;
    
    const kwNorm = normalizeKey(keyword);
    if (kwGeo && normalizeKey(keyword) === normalizeKey(kwGeo)) {
      // Core keyword: kw+geo exactly matches keyword
      coreKeywordSet.add(kwNorm);
    }
  }

  // Budgets and parameters
  const partitionSplit = Number(cfg.partitionSplit) || 0.66;
  const coreBudget = partitionSplit;
  const kappa = Number(cfg.kappa) || 1;

  // Build score arrays (only core scores - geo scores calculated per-location in FinalMerge)
  const coreScores = new Array(n).fill(0);

  // Collect distinct services present in Keyword sheet
  const serviceList = [];
  const seen = new Set();
  for (let i = 0; i < services.length; i++) {
    const s = normalizeKey(services[i] || '');
    if (!s) continue;
    if (!seen.has(s)) {
      seen.add(s);
      serviceList.push(s);
    }
  }

  // Assign core scores per service with full partitionSplit budget per service and decay inside each service group
  for (let si = 0; si < serviceList.length; si++) {
    const serviceKey = serviceList[si];
    const serviceCoreBudget = coreBudget;

    // indices for this service that are core keywords
    const indices = [];
    for (let i = 0; i < n; i++) {
      if (normalizeKey(services[i] || '') !== serviceKey) continue;
      const kw = keywords[i];
      if (!kw) continue;
      const k = normalizeKey(kw);
      if (coreKeywordSet.has(k)) indices.push(i);
    }

    const m = indices.length;
    if (m === 0 || serviceCoreBudget === 0) continue;
    // decay weights within this service
    let wSum = 0;
    const wArr = [];
    for (let j = 0; j < m; j++) {
      const x = ((j + 1) / (m + 1)) * partitionSplit; // sample in [0, split]
      const w = 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
      wArr.push(w);
      wSum += w;
    }
    if (wSum > 0) {
      for (let j = 0; j < m; j++) {
        const idx = indices[j];
        coreScores[idx] = serviceCoreBudget * (wArr[j] / wSum);
      }
    }
  }

  logSumCheck('Keyword core normalized', coreScores);

  // Write core scores (geo scores calculated per-location in FinalMerge)
  writeColumnByHeader(keywordSheet, 'keyword core score', coreScores);
  
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


