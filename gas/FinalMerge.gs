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
  if (idxKKeyword === -1) throw new Error('Keyword sheet missing Keywords');
  if (idxKCore === -1) throw new Error('Keyword sheet missing "keyword core score" header');
  const keywordCoreMap = new Map();
  for (let r = 1; r < kValues.length; r++) {
    const row = kValues[r];
    const key = normalizeKey(row[idxKKeyword]);
    const coreVal = Number(row[idxKCore]) || 0;
    keywordCoreMap.set(key, coreVal);
  }

  // Pre-calculate geo scores per location-service combination (0.34 divided by unique kw+geo count per location-service)
  const cfg = getConfig();
  const geoBudget = 1 - (Number(cfg.partitionSplit) || 0.66);
  const locationServiceGeoScoreMap = new Map(); // "location|service" -> geo score per keyword
  
  // Group Rankings rows by location AND service, count unique kw+geo per location-service combination
  const locationServiceKwGeoMap = new Map(); // "location|service" -> Set of unique kw+geo values
  for (let r = 1; r < rValues.length; r++) {
    const row = rValues[r];
    const geo = String(row[idxGeo] || '').trim();
    const service = String(row[idxService] || '').trim();
    const kwPlusGeo = idxKwGeo !== -1 ? String(row[idxKwGeo] || '').trim() : '';
    if (!geo || !service || !kwPlusGeo) continue;
    
    const geoNorm = normalizeKey(geo);
    const serviceNorm = normalizeKey(service);
    const kwGeoNorm = normalizeKey(kwPlusGeo);
    
    // Only count if this is a geo keyword (kw+geo doesn't match base keyword)
    const keyword = String(row[idxKeyword] || '').trim();
    if (keyword && normalizeKey(keyword) === kwGeoNorm) continue; // Skip core keywords
    
    // Use composite key: location|service
    const compositeKey = geoNorm + '|' + serviceNorm;
    if (!locationServiceKwGeoMap.has(compositeKey)) {
      locationServiceKwGeoMap.set(compositeKey, new Set());
    }
    locationServiceKwGeoMap.get(compositeKey).add(kwGeoNorm);
  }
  
  // Calculate geo score per location-service: 0.34 / unique kw+geo count for this location-service combination
  for (const [compositeKey, kwGeoSet] of locationServiceKwGeoMap.entries()) {
    const count = kwGeoSet.size;
    const geoScore = count > 0 ? geoBudget / count : 0;
    locationServiceGeoScoreMap.set(compositeKey, geoScore);
    const [location, service] = compositeKey.split('|');
    console.log('Location "' + location + '" + Service "' + service + '": ' + count + ' unique geo keywords, score per keyword = ' + geoScore.toFixed(6));
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

  // Build final scores
  const finalRaw = [];
  const rankingKeywordScores = [];
  const rankingIdealScores = [];
  const rankingGapScores = [];
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
      // Geo keywords: use per-location-service geo score (0.34 / unique kw+geo count for this location-service combination)
      const geoNorm = normalizeKey(geo);
      const serviceNorm = normalizeKey(service);
      const compositeKey = geoNorm + '|' + serviceNorm;
      keywordScore = locationServiceGeoScoreMap.get(compositeKey) || 0;
      if (keywordScore === 0) {
        console.warn('Geo keyword score missing for location "' + String(geo || '') + '" and service "' + String(service || '') + '"; defaulting to 0');
      }
    } else {
      // Core keywords: lookup from Keyword sheet (pre-calculated with decay)
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
    const ideal = locationScore * nicheScore * keywordScore;
    rankingIdealScores.push(ideal);
    const raw = locationScore * nicheScore * keywordScore * rankingScore;
    const gap = ideal - raw; // GAP = Ideal - Final Score
    rankingGapScores.push(gap);
    finalRaw.push(raw);
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
  
  // Aggregate Ideal and Actual scores by Location
  const locationAggregates = new Map(); // normalized location -> {ideal, actual}
  for (let r = 1; r < rValues.length; r++) {
    const row = rValues[r];
    const geoRaw = row[idxGeo];
    const geoNorm = normalizeKey(geoRaw);
    const i = r - 1;
    const ideal = rankingIdealScores[i];
    const actual = finalRaw[i];
    
    if (!locationAggregates.has(geoNorm)) {
      locationAggregates.set(geoNorm, {ideal: 0, actual: 0});
    }
    const agg = locationAggregates.get(geoNorm);
    agg.ideal += ideal;
    agg.actual += actual;
  }
  
  // Aggregate Ideal and Actual scores by Service
  const serviceAggregates = new Map(); // normalized service -> {ideal, actual}
  for (let r = 1; r < rValues.length; r++) {
    const row = rValues[r];
    const serviceRaw = row[idxService];
    const serviceNorm = normalizeKey(serviceRaw);
    const i = r - 1;
    const ideal = rankingIdealScores[i];
    const actual = finalRaw[i];
    
    if (!serviceAggregates.has(serviceNorm)) {
      serviceAggregates.set(serviceNorm, {ideal: 0, actual: 0});
    }
    const agg = serviceAggregates.get(serviceNorm);
    agg.ideal += ideal;
    agg.actual += actual;
  }
  
  // Write Rankings audit columns: keyword score, Ideal, GAP, and FINAL_SCORE
  if (rankingKeywordScores.length) {
    writeColumnByHeader(rankings, 'keyword score', rankingKeywordScores);
  }
  if (rankingIdealScores.length) {
    writeColumnByHeader(rankings, 'Ideal', rankingIdealScores);
  }
  if (rankingGapScores.length) {
    writeColumnByHeader(rankings, 'GAP', rankingGapScores);
  }
  if (finalRaw.length) {
    writeColumnByHeader(rankings, 'FINAL_SCORE', finalRaw);
  }

  // Write aggregated Ideal, Actual, and GAP to Location sheet
  const locIdealCol = [];
  const locActualCol = [];
  const locGapCol = [];
  for (let r = 1; r < lValues.length; r++) {
    const row = lValues[r];
    const targetNorm = normalizeKey(row[idxTarget]);
    const agg = locationAggregates.get(targetNorm);
    if (agg) {
      locIdealCol.push(agg.ideal);
      locActualCol.push(agg.actual);
      locGapCol.push(agg.ideal - agg.actual);
    } else {
      locIdealCol.push(0);
      locActualCol.push(0);
      locGapCol.push(0);
    }
  }
  if (locIdealCol.length) {
    writeColumnByHeader(locationSheet, 'Ideal Score (Sum: Location×Service×Keyword)', locIdealCol);
    writeColumnByHeader(locationSheet, 'Actual Score (Sum: Location×Service×Keyword×Ranking)', locActualCol);
    writeColumnByHeader(locationSheet, 'Opportunity GAP (Ideal - Actual)', locGapCol);
  }
  
  // Log top 3 locations by GAP
  const locationGapList = [];
  for (let r = 1; r < lValues.length; r++) {
    const row = lValues[r];
    const targetName = String(row[idxTarget] || '');
    const gap = locGapCol[r - 1] || 0;
    locationGapList.push({name: targetName, gap: gap});
  }
  locationGapList.sort((a, b) => b.gap - a.gap);
  console.log('Top 3 Locations by Opportunity GAP:');
  for (let i = 0; i < Math.min(3, locationGapList.length); i++) {
    console.log('  ' + (i + 1) + '. ' + locationGapList[i].name + ' = ' + locationGapList[i].gap.toFixed(6));
  }

  // Write aggregated Ideal, Actual, and GAP to Niche sheet
  const nicheIdealCol = [];
  const nicheActualCol = [];
  const nicheGapCol = [];
  for (let r = 1; r < nValues.length; r++) {
    const row = nValues[r];
    const serviceNorm = normalizeKey(row[idxNService]);
    const agg = serviceAggregates.get(serviceNorm);
    if (agg) {
      nicheIdealCol.push(agg.ideal);
      nicheActualCol.push(agg.actual);
      nicheGapCol.push(agg.ideal - agg.actual);
    } else {
      nicheIdealCol.push(0);
      nicheActualCol.push(0);
      nicheGapCol.push(0);
    }
  }
  if (nicheIdealCol.length) {
    writeColumnByHeader(nicheSheet, 'Ideal Score (Sum: Location×Service×Keyword)', nicheIdealCol);
    writeColumnByHeader(nicheSheet, 'Actual Score (Sum: Location×Service×Keyword×Ranking)', nicheActualCol);
    writeColumnByHeader(nicheSheet, 'Opportunity GAP (Ideal - Actual)', nicheGapCol);
  }
  
  // Log top 3 services by GAP
  const serviceGapList = [];
  for (let r = 1; r < nValues.length; r++) {
    const row = nValues[r];
    const serviceName = String(row[idxNService] || '');
    const gap = nicheGapCol[r - 1] || 0;
    serviceGapList.push({name: serviceName, gap: gap});
  }
  serviceGapList.sort((a, b) => b.gap - a.gap);
  console.log('Top 3 Services by Opportunity GAP:');
  for (let i = 0; i < Math.min(3, serviceGapList.length); i++) {
    console.log('  ' + (i + 1) + '. ' + serviceGapList[i].name + ' = ' + serviceGapList[i].gap.toFixed(6));
  }
}


