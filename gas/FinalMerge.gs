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
  const idxRankMod = rNorm.indexOf('ranking modifier score');
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

  const keywordSheet = getSheet('Keyword');
  const kValues = getDataRangeValues(keywordSheet);
  const kHeader = kValues[0].map(String);
  const kNorm = kHeader.map(h => normalizeKey(h));
  const idxKKeyword = kNorm.indexOf('keywords');
  const idxKCore = kNorm.indexOf('keyword core score');
  const idxKGeo = kNorm.indexOf('keyword geo score');
  const idxKLegacy = kNorm.indexOf('keyword score');
  if (idxKKeyword === -1) throw new Error('Keyword sheet missing Keywords');
  const keywordCoreMap = new Map();
  const keywordGeoMap = new Map();
  for (let r = 1; r < kValues.length; r++) {
    const row = kValues[r];
    const key = normalizeKey(row[idxKKeyword]);
    const coreVal = idxKCore !== -1 ? (Number(row[idxKCore]) || 0) : (idxKLegacy !== -1 ? (Number(row[idxKLegacy]) || 0) : 0);
    const geoVal = idxKGeo !== -1 ? (Number(row[idxKGeo]) || 0) : (idxKLegacy !== -1 ? (Number(row[idxKLegacy]) || 0) : 0);
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
  for (let r = 1; r < rValues.length; r++) {
    const row = rValues[r];
    const service = row[idxService];
    const geo = row[idxGeo];
    const keyword = row[idxKeyword];
    const kwPlusGeo = idxKwGeo !== -1 ? row[idxKwGeo] : '';
    const locationScore = locMap.get(normalizeKey(geo)) || 0;
    const nicheScore = nicheMap.get(normalizeKey(service)) || 0;
    const useGeo = isGeoString(kwPlusGeo);
    const kKey = normalizeKey(keyword);
    const keywordScore = useGeo ? (keywordGeoMap.get(kKey) || 0) : (keywordCoreMap.get(kKey) || 0);
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


