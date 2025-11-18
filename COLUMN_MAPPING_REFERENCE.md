# 📊 Quick Reference: Column Mapping Changes

## Before (v3.3) → After (v3.4)

### **Old Structure: TWO SHEETS**

#### **logs sheet (DELETED):**
```
A: Job IDs
B: Keyword (duplicate)
C: Source Row
D: status
E: submitted timestamp
F: completed timestamp
```

#### **ranking_results sheet (OLD):**
```
A: Ranking Position
B: Ranking URL
C: Keyword
D-I: (added by GRankingScore)
```

---

### **New Structure: ONE SHEET**

#### **ranking_results sheet (NEW - v3.4):**
```
A: Ranking Position          [UNCHANGED]
B: Ranking URL               [UNCHANGED]
C: Keyword                   [UNCHANGED]
D: rank_score               [UNCHANGED - added by GRankingScore]
E: keyword_score            [UNCHANGED - added by GRankingScore]
F: geo                      [UNCHANGED - added by GRankingScore]
G: FINAL_SCORE              [UNCHANGED - added by GRankingScore]
H: Ideal                    [UNCHANGED - added by GRankingScore]
I: GAP                      [UNCHANGED - added by GRankingScore]
J: Job_ID                   [NEW - from logs A]
K: Source_Row               [NEW - from logs C]
L: status                   [NEW - from logs D]
M: submitted_timestamp      [NEW - from logs E]
N: completed_timestamp      [NEW - from logs F]
```

**Note:** Old logs column B (Keyword) was intentionally skipped as it duplicated column C.

---

## Code Reference Changes

### **Reading Logs Data**

#### **OLD (v3.3):**
```javascript
const logs = getLogsSheet();
const rows = logs.getRange(2, 1, lastRow - 1, 6).getValues(); // A..F
const taskId = rows[i][0];        // A
const keyword = rows[i][1];       // B
const sourceRow = rows[i][2];     // C
const status = rows[i][3];        // D
const submittedAt = rows[i][4];   // E
const completedAt = rows[i][5];   // F
```

#### **NEW (v3.4):**
```javascript
const rs = getOrCreateRankingResultsSheet();
const logsData = rs.getRange(2, RANKING_COL.JOB_ID, lastRow - 1, 5).getValues(); // J-N
const keywordData = rs.getRange(2, RANKING_COL.KEYWORD, lastRow - 1, 1).getValues(); // C
const taskId = logsData[i][0];        // J (index 0)
const sourceRow = logsData[i][1];     // K (index 1)
const status = logsData[i][2];        // L (index 2)
const submittedAt = logsData[i][3];   // M (index 3)
const completedAt = logsData[i][4];   // N (index 4)
const keyword = keywordData[i][0];    // C
```

---

### **Writing Logs Data**

#### **OLD (v3.3):**
```javascript
const logs = getLogsSheet();
logs.getRange(row, 1).setValue(taskId);        // A: Job ID
logs.getRange(row, 2).setValue(keyword);       // B: Keyword
logs.getRange(row, 3).setValue(sourceRow);     // C: Source Row
logs.getRange(row, 4).setValue('submitted');   // D: status
logs.getRange(row, 5).setValue(new Date());    // E: submitted timestamp
logs.getRange(row, 6).setValue('');            // F: completed timestamp
```

#### **NEW (v3.4):**
```javascript
const rs = getOrCreateRankingResultsSheet();
rs.getRange(row, RANKING_COL.JOB_ID).setValue(taskId);              // J
rs.getRange(row, RANKING_COL.SOURCE_ROW).setValue(sourceRow);       // K
rs.getRange(row, RANKING_COL.STATUS).setValue('submitted');         // L
rs.getRange(row, RANKING_COL.SUBMITTED_AT).setValue(new Date());    // M
rs.getRange(row, RANKING_COL.COMPLETED_AT).setValue('');            // N
// Keyword already written to column C
```

---

## Array Index Reference

When reading data via `getValues()`:

### **OLD logs sheet (A-F = 6 columns):**
```javascript
const rows = logs.getRange(2, 1, lastRow - 1, 6).getValues();
rows[i][0] = Job ID        (A)
rows[i][1] = Keyword       (B)
rows[i][2] = Source Row    (C)
rows[i][3] = status        (D)
rows[i][4] = submitted     (E)
rows[i][5] = completed     (F)
```

### **NEW ranking_results J-N (5 columns):**
```javascript
const logsData = rs.getRange(2, RANKING_COL.JOB_ID, lastRow - 1, 5).getValues();
logsData[i][0] = Job ID        (J)
logsData[i][1] = Source Row    (K)
logsData[i][2] = status        (L)
logsData[i][3] = submitted     (M)
logsData[i][4] = completed     (N)
// Keyword read separately from column C
```

---

## Status Values

Column L (status) can contain:

| Status | Meaning |
|--------|---------|
| `submitted` | Just submitted to DataForSEO API |
| `pending` | Waiting for results from API |
| `fetched` | Results successfully retrieved |
| `no_results` | Task complete but no rankings found for our domain |
| `failed` | Task timed out or encountered error |
| `RE-SUBMITTED` | Task was retried via `retryPendingTasks()` |

---

## Constant Reference

Use these constants instead of hardcoded numbers:

```javascript
// Column numbers (1-based for Google Sheets API)
RANKING_COL.POSITION       // 1  (A)
RANKING_COL.URL            // 2  (B)
RANKING_COL.KEYWORD        // 3  (C)
RANKING_COL.RANK_SCORE     // 4  (D)
RANKING_COL.KEYWORD_SCORE  // 5  (E)
RANKING_COL.GEO            // 6  (F)
RANKING_COL.FINAL_SCORE    // 7  (G)
RANKING_COL.IDEAL          // 8  (H)
RANKING_COL.GAP            // 9  (I)
RANKING_COL.JOB_ID         // 10 (J) ← NEW
RANKING_COL.SOURCE_ROW     // 11 (K) ← NEW
RANKING_COL.STATUS         // 12 (L) ← NEW
RANKING_COL.SUBMITTED_AT   // 13 (M) ← NEW
RANKING_COL.COMPLETED_AT   // 14 (N) ← NEW
```

---

## Quick Comparison

| Operation | OLD (v3.3) | NEW (v3.4) |
|-----------|-----------|-----------|
| Get logs sheet | `getLogsSheet()` | `getOrCreateRankingResultsSheet()` |
| Read task ID | `logs.getRange(row, 1)` | `rs.getRange(row, RANKING_COL.JOB_ID)` |
| Read status | `logs.getRange(row, 4)` | `rs.getRange(row, RANKING_COL.STATUS)` |
| Update status | `logs.getRange(row, 4).setValue('fetched')` | `rs.getRange(row, RANKING_COL.STATUS).setValue('fetched')` |
| Clear sheets | Clear logs + ranking_results | Clear ranking_results only |
| Total sheets | 2 sheets (logs + ranking_results) | 1 sheet (ranking_results) |

---

## Migration Path

### **Step 1: Update Code**
- Deploy `Gordon-kw-script-v3.4.js`
- Deploy `retryPendingTasks_V1.3.1.gs`

### **Step 2: Initialize Headers**
```javascript
// Run once to create new headers
getOrCreateRankingResultsSheet();
```

### **Step 3: Test**
```javascript
// Submit a test batch
submitScheduled();

// Wait 5-10 minutes
Utilities.sleep(300000);

// Fetch results
fetchScheduled();
```

### **Step 4: Verify**
- Check that columns J-N have data
- Check that columns A-C show ranking results
- Verify Source_Row (K) is formatted as numbers

### **Step 5: Archive Old Logs**
- Rename `logs` sheet to `logs_archived`
- Or delete it (after confirming no data needed)

---

**Print this page for quick reference during migration!**

