function runLocationScore() {
  const cfg = getConfig();
  const sheet = getSheet('Location');
  const values = getDataRangeValues(sheet);
  if (values.length < 2) return;
  const header = values[0].map(h => String(h));
  const normHeader = header.map(h => normalizeKey(h));

  const idxPopulation = normHeader.indexOf('population');
  const idxIncome = normHeader.indexOf('income');
  if (idxPopulation === -1 || idxIncome === -1) throw new Error('Location sheet missing population or income headers');

  const numerical = [];
  for (let r = 1; r < values.length; r++) {
    const row = values[r];
    const pop = Number(row[idxPopulation]) || 0;
    const inc = Number(row[idxIncome]) || 0;
    const val = (inc - cfg.povertyLine) * pop;
    numerical.push(val);
  }
  const normalized = normalizeVector(numerical);
  logSumCheck('Location normalized', normalized);

  writeColumnByHeader(sheet, 'numerical', numerical);
  writeColumnByHeader(sheet, 'normalized score', normalized);
}


