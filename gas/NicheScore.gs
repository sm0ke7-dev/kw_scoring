function runNicheScore() {
  const sheet = getSheet('Niche');
  const values = getDataRangeValues(sheet);
  if (values.length < 2) return;
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeKey(h));
  const idxScore = normHeader.indexOf('keyword planner score');
  if (idxScore === -1) throw new Error('Niche sheet missing Keyword Planner Score header');

  const raw = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    raw.push(Number(row[idxScore]) || 0);
  }
  const normalized = normalizeVector(raw);
  logSumCheck('Niche normalized', normalized);
  writeColumnByHeader(sheet, 'Normalization', normalized);
}


