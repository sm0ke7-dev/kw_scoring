/** Utilities for sheet IO, normalization, joins, and logging */

// Light highlight for computed output columns
var OUTPUT_BG_COLOR = '#FFF2CC'; // pale yellow

function getSheet(name) {
  const ss = getSpreadsheet();
  const sheet = ss.getSheetByName(name);
  if (!sheet) throw new Error('Missing sheet: ' + name);
  return sheet;
}

function getHeaderRow(sheet) {
  const range = sheet.getRange(1, 1, 1, sheet.getLastColumn());
  return range.getValues()[0];
}

function ensureHeader(sheet, header) {
  const current = getHeaderRow(sheet);
  const maxLen = Math.max(current.length, header.length);
  const next = [];
  for (let i = 0; i < maxLen; i++) next.push(header[i] || current[i] || '');
  sheet.getRange(1, 1, 1, next.length).setValues([next]);
}

function getDataRangeValues(sheet) {
  const range = sheet.getDataRange();
  return range.getValues();
}

function writeColumnByHeader(sheet, headerName, values, headerRowIndex = 1) {
  const header = getHeaderRow(sheet);
  let colIndex = header.findIndex(h => normalizeKey(h) === normalizeKey(headerName));
  if (colIndex === -1) {
    // append new column
    colIndex = header.length;
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
    range.setBackground(OUTPUT_BG_COLOR);
    sheet.getRange(headerRowIndex, colIndex + 1).setBackground(OUTPUT_BG_COLOR);
  } catch (e) {
    // ignore styling errors
  }
}

function findHeaderIndex(sheet, name) {
  const header = getHeaderRow(sheet);
  const idx = header.findIndex(h => normalizeKey(h) === normalizeKey(name));
  return idx; // 0-based, -1 if not found
}

function getColumnValuesByHeader(sheet, name) {
  const idx = findHeaderIndex(sheet, name);
  if (idx === -1) return [];
  const lastRow = sheet.getLastRow();
  if (lastRow < 2) return [];
  const range = sheet.getRange(2, idx + 1, lastRow - 1, 1);
  return range.getValues().map(r => r[0]);
}

function sum(arr) {
  return arr.reduce((a, b) => a + (Number(b) || 0), 0);
}

function normalizeVector(values) {
  const total = sum(values);
  if (!total) return values.map(_ => 0);
  return values.map(v => (Number(v) || 0) / total);
}

function buildIndex(values, keySelector) {
  const map = new Map();
  for (let i = 0; i < values.length; i++) {
    const key = normalizeKey(keySelector(values[i], i));
    if (!map.has(key)) map.set(key, i);
  }
  return map;
}

function logSumCheck(label, values) {
  const s = sum(values);
  const msg = label + ' SUM=' + s.toFixed(6);
  console.log(msg);
  try {
    const ss = getSpreadsheet();
    let logSheet = ss.getSheetByName('Log');
    if (!logSheet) logSheet = ss.insertSheet('Log');
    if (logSheet.getLastRow() === 0) {
      logSheet.getRange(1, 1, 1, 3).setValues([["timestamp","label","sum"]]);
    }
    const nextRow = logSheet.getLastRow() + 1;
    logSheet.getRange(nextRow, 1, 1, 3).setValues([[new Date(), label, s]]);
  } catch (e) {
    // ignore log sheet failures
  }
  return s;
}


