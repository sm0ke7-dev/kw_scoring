## Rank Score Calculator (Standalone Apps Script) — Implementation Plan

### Environment
- Standalone Google Apps Script (V8). Reads a target Google Sheet via `SPREADSHEET_ID` (in `Config.gs` or read from `Config` tab).
- Case-insensitive, trimmed joins. Non-matches resolve to score 0.
- Each module is idempotent, recomputes in memory, writes full audit to its sheet.

### Files
- `Config.gs`: Read config from `Config` tab; fallback constants (poverty line, partition split, `kappa`, `rank modification kappa`).
- `Utils.gs`: Sheet access helpers, case-insensitive key builders, normalization, number formatting, logging, safe write to ranges, checksum logs.
- `LocationScore.gs`: Compute `numerical = (income - poverty_line) * population` and normalized score; write to `Location` tab columns incl. `numerical`, `normalized score`.
- `NicheScore.gs`: Read `Niche` tab, compute normalization of `Keyword Planner Score`, write normalized column.
- `KeywordScore.gs`: Build weights for `core` vs `geo` lists using split and `kappa` with logarithmic decay; normalize so sum=1; write back per keyword on `Keyword` tab.
- `RankingScore.gs`: For ranks 1–10 compute `e^(−kappa*(rank−1))`, else 0; normalize within `Rankings` tab; write `ranking modifier score` column.
- `FinalMerge.gs`: Join `Rankings` rows to normalized layers: `geo→target`, `service→Service`, `keyword→Keywords`; multiply components, produce `final_raw` then re-normalize to `final_normalized`; write to `Final` tab with specified columns.
- `Menu.gs`: Add a custom menu: Run All, Run: Location/Niche/Keyword/Ranking/Final, Clear outputs.

### Data Contracts (Sheets)
- `Config`: columns `partition split`, `kappa`, `rank modification kappa`, optional `poverty_line`, optional `spreadsheet_id`.
- `Keyword`: columns `Services`, `Keywords`; module writes `keyword score` column.
- `Location`: `target`, `state`, `lat`, `long`, `population`, `income`, plus `numerical`, `normalized score` (module writes last two).
- `Niche`: `Service`, `Keyword Planner Score`, module writes `Normalization`.
- `Rankings`: `service`, `geo`, `keyword`, `kw+geo`, `lat`, `long`, `Ranking Position`, `Ranking URL`, module writes `ranking modifier score`.
- `Final`: output table per brief with columns: `service`, `geo`, `keyword`, `location_score`, `niche_score`, `keyword_score`, `ranking_score`, `final_raw`, `final_normalized`.

### Key Logic Details
- Joins: `normalizeKey(s) => trim(toLowerCase(s))`; use maps keyed by normalized strings for `geo/target`, `service/Service`, and `keyword/Keywords`.
- Normalization: divide by sum; guard zero-sum case -> all zeros persist, log warning.
- Keyword curve: `y(x) = 1 - log(1 + κx) / log(1 + κ)`; allocate totals per partition split to core vs geo sets; rescale per area-under-curve to sum=1.
- Ranking decay: zero for rank>10; include `rank modification kappa`.
- Idempotency: Each run overwrites computed columns, not appending.
- Validation: After each module and after final merge, log `SUM≈1.000000` checks to console and optionally to a `Log` tab.

### Runbook
- Use menu to run modules independently or `Run All` in order: Location → Niche → Keyword → Ranking → Final Merge.
- Config can be edited in `Config` tab without code changes.

### Acceptance
- All normalized columns individually sum ≈ 1.0; final normalized sum = 1.0.
- Case-insensitive joins; non-matches yield 0 scores.
- Re-running yields identical outputs given unchanged inputs.

### Implementation Todos
- setup-config: Create Config.gs with config loader and SPREADSHEET_ID support
- implement-utils: Implement Utils.gs helpers for IO, joins, normalization, logging
- location-score: Implement LocationScore.gs to compute numerical and normalized columns
- niche-score: Implement NicheScore.gs to normalize Keyword Planner Score
- keyword-score: Implement KeywordScore.gs with logarithmic decay and split normalization
- ranking-score: Implement RankingScore.gs with exponential decay and normalization
- final-merge: Implement FinalMerge.gs to join layers and compute final normalized
- menu-and-commands: Add Menu.gs with Run All and per-module entries
- validation: Add SUM≈1 checks and optional Log tab outputs

