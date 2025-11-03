function runKeywordScore() {
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

  // Build decay weights across keyword list
  const weights = [];
  for (let i = 0; i < n; i++) {
    const x = n > 1 ? i / (n - 1) : 0;
    const w = 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
    weights.push(w);
  }
  const base = normalizeVector(weights);

  // Partition into core vs geo buckets
  const coreShare = Number(cfg.partitionSplit) || 0;
  const geoShare = 1 - coreShare;
  const coreScores = base.map(v => v * coreShare);
  const geoScores = base.map(v => v * geoShare);

  logSumCheck('Keyword core normalized', coreScores);
  logSumCheck('Keyword geo normalized', geoScores);

  // Write both columns for audit; keep legacy 'keyword score' as core
  writeColumnByHeader(sheet, 'keyword core score', coreScores);
  writeColumnByHeader(sheet, 'keyword geo score', geoScores);
  writeColumnByHeader(sheet, 'keyword score', coreScores);
}


