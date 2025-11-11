# Rank Score Calculator

A modular Google Apps Script scoring system that quantifies the value of local SEO keywords based on **location**, **niche (service)**, **keyword type**, and **ranking performance**.

## Overview

This system calculates normalized scores across multiple dimensions to identify keyword opportunities across markets. Each scoring component is normalized independently and combined multiplicatively to produce a final normalized value.

## Key Discoveries: Implementation Fixes

### Discovery 1: Per-Service Budget Allocation

**Current Implementation (KeywordScore2.gs):**
- Each service gets the FULL budget allocation (not divided across services)
- Core keywords: Each service gets `partitionSplit` (0.66) divided among its core keywords with decay
- Geo keywords: Each service gets `(1 - partitionSplit)` (0.34) per location divided among its geo keywords in that location
- Logarithmic decay applied within each service's keyword group

**Why This Matters:**
- Each service (wildlife, raccoon, squirrel, bat) is treated independently
- A service with 4 core keywords gets 0.66, each keyword weighted by decay position
- Same service's geo keywords in Atlanta get 0.34 / (count of geo keywords for that service in Atlanta)
- Allows proper scaling across multiple service lines without dilution

**Example:**
- Wildlife removal in Riverview with 10 geo keywords: each gets 0.34/10 = 0.034
- Raccoon removal in Riverview with 3 geo keywords: each gets 0.34/3 = 0.113
- Each service maintains its full budget allocation independently

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
3. **KeywordScore2.gs** - Assigns keyword scores with per-service budget allocation and decay ✅ **ACTIVE**
4. **RankingScore.gs** - Applies logarithmic decay to ranking positions (ranks 1-10, raw scores as multipliers)
5. **FinalMerge.gs** - Combines all scores, calculates GAPs, writes to Rankings/Location/Niche sheets
6. **Menu.gs** - Defines custom menu for running scoring functions
7. **Utils.gs** - Shared utility functions for sheet operations and normalization
8. **Config.gs** - Configuration management and spreadsheet ID resolution

### Sheets

- **Config** - Configuration parameters (partition split, kappa, rank modification kappa, poverty line)
- **Location** - Location data with calculated scores + Ideal/Actual/GAP aggregates
- **Niche** - Service data with normalized scores + Ideal/Actual/GAP aggregates
- **Keyword** - Keyword list with core scores (geo scores calculated per-location in FinalMerge)
- **Rankings** - Ranking data with all computed scores (keyword score, Ideal, GAP, FINAL_SCORE)
- **Log** - Audit log of normalization checks

**Note:** The "Final" sheet has been removed. All final scores now appear in the Rankings sheet.

## Calculation Flow

1. **Location Score**: `(income - poverty_line) × population` → normalized
2. **Niche Score**: `Keyword Planner Score` → normalized
3. **Keyword Score** (per service): 
   - Core keywords: Each service gets full `partitionSplit` (0.66), distributed with logarithmic decay within service
   - Geo keywords: Each service gets full `(1 - partitionSplit)` (0.34) per location, divided equally among geo keywords for that service in that location
4. **Ranking Score**: Logarithmic decay for ranks 1-10, else 0 → **raw scores (no normalization, acts as multiplier)**
5. **Final Score**: `location × niche × keyword × ranking` → written to Rankings sheet
6. **GAP Analysis**: `Ideal - Actual` calculated at keyword, location, and service levels

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
5. Use **Rank Scoring** menu:
   - **Run All** - Executes complete scoring pipeline
   - **Location Score** - Calculate location weights
   - **Niche Score** - Calculate service weights
   - **Keyword Score** - Calculate keyword weights (uses KeywordScore2.gs)
   - **Ranking Score** - Calculate ranking modifiers
   - **Final Merge** - Combine all scores, write to Rankings sheet, calculate GAPs
   - **Clear Log** - Clear the Log sheet

## Current Status

✅ **Per-service budget allocation** - Each service gets full budget (0.66 core + 0.34 geo per location), not divided across services.

✅ **Ranking modifier score fixed** - Uses raw decay scores (no normalization) so rank 1 = 1.0 as multiplier.

✅ **Niche score join fixed** - Added diagnostic logging and identified service name mismatch between Niche and Rankings sheets.

✅ **Sheet consolidation** - Removed redundant Final sheet; all scores now in Rankings sheet.

✅ **GAP analysis** - Opportunity GAP calculated at keyword, location, and service levels.

✅ **Code cleanup** - Removed unused visualization functions, redundant columns, and obsolete menu items.

## Active Implementation: KeywordScore2.gs

The system uses `KeywordScore2.gs` which implements per-service budget allocation:
- Each service gets the full `partitionSplit` (0.66) for core keywords
- Each service gets the full `(1 - partitionSplit)` (0.34) for geo keywords per location
- Logarithmic decay applied within each service group
- Outputs: `keyword core score` column only (geo scores calculated dynamically in FinalMerge)

**Note:** Historical implementations (KeywordScore.gs, KeywordScore1.gs) are preserved for reference but not actively used.

## References

- Project Brief: `rank_score_calculator_brief.md`
- Instructions Summary: `instructions_summary_transcript.ms`
- Meeting Transcript: `captions (16).srt`

## Authors

- Gordon Ligon (project architecture)
- Daniel Capistrano (implementation)

