function runRankingScore() {
  const cfg = getConfig();
  const sheet = getSheet('Rankings');
  const values = getDataRangeValues(sheet);
  if (values.length < 2) return;
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeKey(h));
  const idxRank = normHeader.indexOf('ranking position');
  if (idxRank === -1) throw new Error('Rankings sheet missing Ranking Position header');
  const kappa = Number(cfg.rankModificationKappa) || 1000;
  
  if (kappa <= 0) {
    throw new Error('Kappa must be greater than 0');
  }

  // Calculate logarithmic decay scores using simple formula
  // Normalized score between 0 and 1: rank 1 = 1.0, rank 10 = 0
  const scores = [];
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
    } else {
      score = 0; // Ranks outside 1-10 get zero
    }
    
    scores.push(score);
  }
  
  writeColumnByHeader(sheet, 'rank score', scores);
}


