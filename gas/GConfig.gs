/**
 * GConfig.gs
 * Configuration for Gordon's Scoring System
 */

const G_DEFAULT_CONFIG = {
  povertyLine: 32000,
  partitionSplit: 0.66, // fraction allocated to core keywords; geo gets (1 - split)
  kappa: 1,
  rankModificationKappa: 10,
  spreadsheetId: '' // optional; if blank, use active spreadsheet
};

/**
 * Returns normalized, trimmed lowercase key used for joins.
 */
function normalizeGKey(value) {
  if (value === null || value === undefined) return '';
  return String(value).trim().toLowerCase();
}

/**
 * Load configuration from the `config` sheet if present. Falls back to defaults.
 * 
 * Expected headers in Row 1 (suggested columns I-L to avoid DataForSEO tracking columns A-G):
 * - Column I: "partition split" (default: 0.66)
 * - Column J: "kappa" (default: 1)
 * - Column K: "rank modification kappa" (default: 10)
 * - Column L: "poverty_line" (default: 32000)
 * 
 * Values should be in Row 2
 */
function getGConfig() {
  const cfg = Object.assign({}, G_DEFAULT_CONFIG);
  try {
    const ss = SpreadsheetApp.getActive();
    const sheet = ss.getSheetByName('config');
    if (!sheet) {
      console.log('📊 Config sheet not found, using defaults');
      return cfg;
    }
    const range = sheet.getDataRange();
    const values = range.getValues();
    if (values.length < 2) {
      console.log('📊 Config sheet has no data, using defaults');
      return cfg;
    }
    const header = values[0].map(h => normalizeGKey(h));
    
    // Track what we find
    let foundAny = false;
    
    // Read first non-empty data row (usually row 2)
    for (let r = 1; r < values.length; r++) {
      const row = values[r];
      if (!row || row.every(v => v === '' || v === null)) continue;
      
      for (let c = 0; c < header.length; c++) {
        const h = header[c];
        const v = row[c];
        if (!h) continue;
        
        if (h === 'partition split') {
          const val = Number(v);
          if (val > 0 && val <= 1) {
            cfg.partitionSplit = val;
            foundAny = true;
          }
        }
        if (h === 'kappa') {
          const val = Number(v);
          if (val > 0) {
            cfg.kappa = val;
            foundAny = true;
          }
        }
        if (h === 'rank modification kappa') {
          const val = Number(v);
          if (val > 0) {
            cfg.rankModificationKappa = val;
            foundAny = true;
          }
        }
        if (h === 'poverty_line' || h === 'poverty line') {
          const val = Number(v);
          if (val > 0) {
            cfg.povertyLine = val;
            foundAny = true;
          }
        }
        if (h === 'spreadsheet_id' || h === 'spreadsheet id') {
          const val = String(v || '').trim();
          if (val) {
            cfg.spreadsheetId = val;
            foundAny = true;
          }
        }
      }
      break;
    }
    
    if (foundAny) {
      console.log('📊 Config loaded from sheet:');
      console.log(`   Partition Split: ${cfg.partitionSplit}`);
      console.log(`   Kappa: ${cfg.kappa}`);
      console.log(`   Rank Modification Kappa: ${cfg.rankModificationKappa}`);
      console.log(`   Poverty Line: $${cfg.povertyLine.toLocaleString()}`);
    } else {
      console.log('📊 No config headers found in sheet, using defaults');
    }
  } catch (e) {
    console.warn('⚠️ getGConfig error, falling back to defaults:', e && e.message);
  }
  return cfg;
}

/**
 * Get the active spreadsheet (Gordon's version always uses active)
 */
function getGSpreadsheet() {
  const ss = SpreadsheetApp.getActive();
  if (!ss) throw new Error('No active spreadsheet found. Please run from within Google Sheets.');
  return ss;
}

