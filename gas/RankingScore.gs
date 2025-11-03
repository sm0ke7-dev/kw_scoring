function runRankingScore() {
  const cfg = getConfig();
  const sheet = getSheet('Rankings');
  const values = getDataRangeValues(sheet);
  if (values.length < 2) return;
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeKey(h));
  const idxRank = normHeader.indexOf('ranking position');
  if (idxRank === -1) throw new Error('Rankings sheet missing Ranking Position header');
  const kappa = Number(cfg.rankModificationKappa) || 10;

  const raw = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rank = Number(row[idxRank]) || 0;
    let score = 0;
    if (rank >= 1 && rank <= 10) {
      score = Math.exp(-kappa * (rank - 1));
    } else {
      score = 0;
    }
    raw.push(score);
  }
  const normalized = normalizeVector(raw);
  logSumCheck('Ranking normalized', normalized);
  writeColumnByHeader(sheet, 'ranking modifier score', normalized);
}


