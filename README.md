# Rank Score Calculator

A modular Google Apps Script scoring system that quantifies the value of local SEO keywords based on **location**, **niche (service)**, **keyword type**, and **ranking performance**.

## Overview

This system calculates normalized scores across multiple dimensions to identify keyword opportunities across markets. Each scoring component is normalized independently and combined multiplicatively to produce a final normalized value.

## Key Discoveries: Implementation Fixes

### Discovery 1: Keyword Scoring Implementation

**Initial Implementation (INCORRECT):**
- Applied a logarithmic decay curve based on keyword position in the Keyword sheet
- Last keyword in the list would get a score of 0
- This caused final scores to be zeroed out even for high-value locations/niches/rankings

**Correct Implementation (as per client instructions):**
- Keywords should be scored using **flat/equal distribution** within each bucket (core vs geo)
- No decay curve should be applied to keyword list position
- All core keywords get equal share of `partitionSplit`
- All geo keywords get equal share of `(1 - partitionSplit)`

**Why This Matters:**

From the transcript: *"take an average so that you get the chunk and then you just say we got five keywords, they're all worth oneif of whatever that first partition is worth"*

This approach is:
- More robust for smaller markets
- Doesn't rely on sparse keyword planner data
- Keeps keyword scores "brain dead" - simple and auditable
- Ensures keywords have the same score regardless of location (location is factored separately)

### Discovery 2: Ranking Modifier Score Normalization

**The Problem:**
- Ranking modifier scores were being normalized (sum = 1)
- This turned rank 1 from 1.0 into ~0.0347
- Client requirement: rank 1 should be 1.0

**Why This Matters:**
- Ranking modifier acts as a **performance multiplier** in the final score calculation
- If rank 1 = 0.0347 instead of 1.0, it unfairly penalizes rank 1 performance
- Rank 1 should preserve full value (multiply by 1.0), not be penalized

**The Fix:**
- Removed normalization from ranking modifier scores
- Now using raw decay curve values directly (no normalization)
- Rank 1 = 1.0 (no penalty, preserves full value)
- Rank 2 = exp(-kappa × 1) (small multiplier)
- Rank 3 = exp(-kappa × 2) (smaller multiplier)
- Ranks 4-10 = continue decaying (very small but visible)
- Ranks > 10 = 0 (as expected)

**Result:**
- Rank 1 correctly multiplies by 1.0 (preserves full value)
- Lower ranks get appropriate decay multipliers
- Decay curve steepness controlled by `rank modification kappa` in Config
- Ranking modifier is now a true performance multiplier, not a normalized score

### Discovery 3: Niche Score Join Issue

**The Problem:**
- Final scores were all zero because `niche_score` was zero
- Diagnostic logging revealed: Niche sheet had "raccoon", "squirrel", "bat"
- Rankings sheet was looking for "wildlife removal"
- Join failed because service values didn't match

**Why This Matters:**
- The final score calculation is multiplicative: `final_raw = location × niche × keyword × ranking`
- If any component is zero, the final score becomes zero
- Service names must match exactly (case-insensitive, trimmed) between Niche and Rankings sheets

**The Fix:**
- Added "wildlife removal" service to Niche sheet with Keyword Planner Score
- Re-ran Niche Score script to recalculate normalized scores
- Now Niche sheet contains all services that appear in Rankings sheet

**Result:**
- Niche scores now properly match between Niche and Rankings sheets
- Final scores calculated correctly with all components non-zero
- Diagnostic logging added to FinalMerge to help identify future join issues

## Project Structure

### Scripts

1. **LocationScore.gs** - Calculates market potential using population and income data
2. **NicheScore.gs** - Normalizes Keyword Planner service values
3. **KeywordScore.gs** - Assigns keyword scores (core vs geo buckets with equal distribution) ✅ **FIXED**
4. **RankingScore.gs** - Applies exponential decay to ranking positions (ranks 1-10, raw scores as multipliers) ✅ **FIXED**
5. **FinalMerge.gs** - Multiplies all normalized scores and re-normalizes

### Sheets

- **Config** - Configuration parameters (partition split, kappa, rank modification kappa, poverty line)
- **Location** - Location data with calculated scores
- **Niche** - Service data with normalized scores
- **Keyword** - Keyword list with core/geo scores
- **Rankings** - Ranking data with modifier scores
- **Final** - Final merged scores per keyword row
- **Log** - Audit log of normalization checks

## Calculation Flow

1. **Location Score**: `(income - poverty_line) × population` → normalized
2. **Niche Score**: `Keyword Planner Score` → normalized
3. **Keyword Score**: 
   - Identify core vs geo keywords
   - Core keywords: equal share of `partitionSplit`
   - Geo keywords: equal share of `(1 - partitionSplit)`
4. **Ranking Score**: `exp(-kappa × (rank - 1))` for ranks 1-10, else 0 → **raw scores (no normalization, acts as multiplier)**
5. **Final Score**: `location × niche × keyword × ranking` → re-normalized

## Configuration

| Parameter | Default | Description |
|-----------|---------|-------------|
| `partition split` | 0.66 | Fraction of total value allocated to core keywords |
| `kappa` | 1 | Curve steepness (for ranking decay, not keyword scoring) |
| `rank modification kappa` | 10 | Decay steepness for ranking positions |
| `poverty_line` | 32000 | Poverty line for location calculation |

## Usage

1. Open your Google Sheet
2. Go to **Extensions → Apps Script**
3. Copy all `.gs` files into the script editor
4. Set `SPREADSHEET_ID` in Config or Script Properties
5. Use **Rank Scoring** menu to run calculations
6. Run **Plot Decay Curve** to visualize (for ranking, not keywords)

## Current Status

✅ **Keyword scoring fixed** - Now uses flat/equal distribution within core and geo buckets.

✅ **Ranking modifier score fixed** - Now uses raw decay scores (no normalization) so rank 1 = 1.0 as multiplier.

✅ **Niche score join fixed** - Added diagnostic logging and identified service name mismatch between Niche and Rankings sheets.

See `docs/rank_score_calculator_plan.md` and `docs/keyword_scoring_fix_plan.md` for implementation details.

## Per‑Service Core Decay (optional)

An alternative keyword scoring variant is available in `gas/KeywordScore1.gs`:

- Split the core budget (`partition split`) across services using weights from `Niche!Normalization` (falls back to equal per service if all weights are zero/missing).
- Within each service group, apply the logarithmic decay controlled by `kappa` to that service’s core keywords (order as listed on the `Keyword` sheet).
- Geo budget is unchanged: equal‑share across unique geo keywords globally.
- Outputs remain the same columns: `keyword core score`, `keyword geo score`.

To use this variant, wire your menu to call `runKeywordScore` from `KeywordScore1.gs` instead of `KeywordScore.gs`.

## References

- Project Brief: `rank_score_calculator_brief.md`
- Instructions Summary: `instructions_summary_transcript.ms`
- Meeting Transcript: `captions (16).srt`

## Authors

- Gordon Ligon (project architecture)
- Daniel Capistrano (implementation)

