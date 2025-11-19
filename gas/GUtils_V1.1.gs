/**
 * GUtils.gs
 * Utilities for sheet IO, normalization, joins, and logging
 * Gordon's Scoring System
 */

// Light highlight for computed output columns
var G_OUTPUT_BG_COLOR = '#C6EFCE'; // light green

function getGSheet(name) {
  const ss = getGSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function getGHeaderRow(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  return range.getValues()[0];
}

function ensureGHeader(sheet, header) {
  const current = getGHeaderRow(sheet);
  const maxLen = Math.max(current.length, header.length);
  const next = [];
  for (let i = 0; i < maxLen; i++) next.push(header[i] || current[i] || '');
  sheet.getRange(1, 1, 1, next.length).setValues([next]);
}

function getGDataRangeValues(sheet) {
  const range = sheet.getDataRange();
  return range.getValues();
}

/**
 * Write values to a column by header name
 * Creates column if it doesn't exist
 * Applies light green background to output
 * NOTE: Scoring columns always start at column E (index 4) to preserve A-D for ranking data
 */
function writeGColumnByHeader(sheet, headerName, values, headerRowIndex = 1) {
  const header = getGHeaderRow(sheet);
  let colIndex = header.findIndex(h => normalizeGKey(h) === normalizeGKey(headerName));
  if (colIndex === -1) {
    // append new column, but ensure it's at least column E (index 4)
    // Columns A-D are reserved for: Ranking Position, Ranking URL, Keyword, Service/Niche
    colIndex = Math.max(header.length, 4);
    
    // Pad header array if needed to reach column E
    while (header.length < colIndex) {
      header.push('');
    }
    
    header.push(headerName);
    sheet.getRange(1, 1, 1, header.length).setValues([header]);
  }
  const startRow = headerRowIndex + 1;
  const numRows = Math.max(values.length, 1);
  const range = sheet.getRange(startRow, colIndex + 1, numRows, 1);
  const colValues = values.map(v => [v]);
  range.clearContent();
  range.setValues(colValues);
  // shade output cells and header
  try {
    range.setBackground(G_OUTPUT_BG_COLOR);
    sheet.getRange(headerRowIndex, colIndex + 1).setBackground(G_OUTPUT_BG_COLOR);
  } catch (e) {
    // ignore styling errors
  }
}

function findGHeaderIndex(sheet, name) {
  const header = getGHeaderRow(sheet);
  const idx = header.findIndex(h => normalizeGKey(h) === normalizeGKey(name));
  return idx; // 0-based, -1 if not found
}

function getGColumnValuesByHeader(sheet, name) {
  const idx = findGHeaderIndex(sheet, name);
  if (idx === -1) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
  return range.getValues().map(r => r[0]);
}

function sumG(arr) {
  return arr.reduce((a, b) => a + (Number(b) || 0), 0);
}

function normalizeGVector(values) {
  const total = sumG(values);
  if (!total) return values.map(_ => 0);
  return values.map(v => (Number(v) || 0) / total);
}

function buildGIndex(values, keySelector) {
  const map = new Map();
  for (let i = 0; i < values.length; i++) {
    const key = normalizeGKey(keySelector(values[i], i));
    if (!map.has(key)) map.set(key, i);
  }
  return map;
}

function logGSumCheck(label, values) {
  const s = sumG(values);
  const msg = label + ' SUM=' + s.toFixed(6);
  console.log(msg);
  // No sheet logging - console only for debugging
  return s;
}

