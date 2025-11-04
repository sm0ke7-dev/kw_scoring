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

  // Calculate raw decay scores (no normalization - rank 1 should be 1.0 as multiplier)
  const scores = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const rank = Number(row[idxRank]) || 0;
    let score = 0;
    if (rank >= 1 && rank <= 10) {
      score = Math.exp(-kappa * (rank - 1));
    } else {
      score = 0; // Ranks > 10 get zero
    }
    scores.push(score);
  }
  // Log check: rank 1 should be 1.0 (or close to it)
  const rank1Count = scores.filter(s => s === 1.0 || Math.abs(s - 1.0) < 0.0001).length;
  console.log('Ranking modifier scores: rank 1 = 1.0 (count: ' + rank1Count + '), raw scores (no normalization)');
  writeColumnByHeader(sheet, 'ranking modifier score', scores);
}


