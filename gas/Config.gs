const DEFAULT_CONFIG = {
  povertyLine: 32000,
  partitionSplit: 0.66, // fraction allocated to core items; geo gets (1 - split)
  kappa: 1,
  rankModificationKappa: 10,
  spreadsheetId: '' // optional; if blank, use script property or prompt
};

/**
 * Returns normalized, trimmed lowercase key used for joins.
 */
function normalizeKey(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toLowerCase();
}

/**
 * Load configuration from the `Config` sheet if present. Falls back to defaults.
 * Expected headers: partition split, kappa, rank modification kappa, poverty_line, spreadsheet_id
 */
function getConfig() {
  const cfg = Object.assign({}, DEFAULT_CONFIG);
  try {
    const ss = SpreadsheetApp.openById(getSpreadsheetId());
    const sheet = ss.getSheetByName('Config');
    if (!sheet) return cfg;
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 2) return cfg;
    const header = values[0].map(h => normalizeKey(h));
    // Collect first non-empty data row
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.every(v => v === '' || v === null)) continue;
      for (let c = 0; c < header.length; c++) {
        const h = header[c];
        const v = row[c];
        if (!h) continue;
        if (h === 'partition split') cfg.partitionSplit = Number(v) || cfg.partitionSplit;
        if (h === 'kappa') cfg.kappa = Number(v) || cfg.kappa;
        if (h === 'rank modification kappa') cfg.rankModificationKappa = Number(v) || cfg.rankModificationKappa;
        if (h === 'poverty_line' || h === 'poverty line') cfg.povertyLine = Number(v) || cfg.povertyLine;
        if (h === 'spreadsheet_id' || h === 'spreadsheet id') cfg.spreadsheetId = String(v || '').trim() || cfg.spreadsheetId;
      }
      break;
    }
  } catch (e) {
    console.warn('getConfig fallback to defaults:', e && e.message);
  }
  return cfg;
}

/**
 * Resolve spreadsheet ID in priority order:
 * 1) Config sheet `spreadsheet_id`
 * 2) Script Properties `SPREADSHEET_ID`
 * 3) DEFAULT_CONFIG.spreadsheetId
 */
function getSpreadsheetId() {
  const props = PropertiesService.getScriptProperties();
  const propId = (props && props.getProperty('SPREADSHEET_ID')) || '';
  if (DEFAULT_CONFIG.spreadsheetId) return DEFAULT_CONFIG.spreadsheetId;
  if (propId) return propId;
  // If a Config sheet exists with spreadsheet_id set, use it
  try {
    const active = SpreadsheetApp.getActive();
    if (active) return active.getId();
  } catch (e) {
    // ignore
  }
  return DEFAULT_CONFIG.spreadsheetId || propId;
}

function getSpreadsheet() {
  const id = getSpreadsheetId();
  if (!id) throw new Error('Missing spreadsheetId. Set Config!Config!spreadsheet_id or Script Property SPREADSHEET_ID.');
  return SpreadsheetApp.openById(id);
}