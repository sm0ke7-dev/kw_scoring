# Rank Score Calculator

A modular Google Apps Script scoring system that quantifies the value of local SEO keywords based on **location**, **niche (service)**, **keyword type**, and **ranking performance**.

## Overview

This system calculates normalized scores across multiple dimensions to identify keyword opportunities across markets. Each scoring component is normalized independently and combined multiplicatively to produce a final priority score with GAP analysis at keyword, location, and service levels.

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

**Note:** Historical implementations (KeywordScore.gs, KeywordScore1.gs) are preserved for reference but not actively used.

### Sheets

- **Config** - Configuration parameters (partition split, kappa, rank modification kappa, poverty line)
- **Location** - Location data with calculated scores + Ideal/Actual/GAP aggregates
- **Niche** - Service data with normalized scores + Ideal/Actual/GAP aggregates
- **Keyword** - Keyword list with core scores (geo scores calculated per-location in FinalMerge)
- **Rankings** - Ranking data with all computed scores (keyword score, Ideal, GAP, FINAL_SCORE)
- **Log** - Audit log of normalization checks

**Note:** The Final sheet has been removed to eliminate redundancy. All final scores and analysis now appear in the Rankings sheet.

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
| `partition split` | 0.66 | Fraction of total value allocated to core keywords per service |
| `kappa` | 1 | Curve steepness for keyword decay within service groups |
| `rank modification kappa` | 10 | Decay steepness for ranking positions |
| `poverty_line` | 32000 | Poverty line threshold for location calculation |

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

## Implementation Notes

### Per-Service Budget Allocation

Each service (wildlife, raccoon, squirrel, bat) gets the **full budget allocation** independently:
- Core keywords: Full `partitionSplit` (0.66) per service, distributed with logarithmic decay
- Geo keywords: Full `(1 - partitionSplit)` (0.34) per location per service, divided equally

**Example:**
- Wildlife removal in Riverview with 10 geo keywords: each gets 0.34/10 = 0.034
- Raccoon removal in Riverview with 3 geo keywords: each gets 0.34/3 = 0.113

This allows proper scaling across multiple service lines without budget dilution.

### Ranking as Performance Multiplier

Ranking scores use **raw logarithmic decay values** (no normalization):
- Rank 1 = 1.0 (preserves full value)
- Rank 2-10 = decaying multipliers based on `rank modification kappa`
- Ranks > 10 = 0

This treats ranking as a performance multiplier rather than a normalized score component.

### Service Name Matching

Service names must match exactly (case-insensitive, trimmed) between Niche and Rankings sheets. Diagnostic logging in FinalMerge helps identify mismatches that would cause zero scores.

### GAP Analysis

Opportunity GAP = `Ideal - Actual` where:
- **Ideal** = location × niche × keyword (assumes perfect ranking = 1.0)
- **Actual** = location × niche × keyword × ranking
- **GAP** = lost opportunity due to ranking position

Higher GAP values indicate keywords/locations/services where ranking improvements would have the most impact.

## Keyword Ranking System

The project also includes a keyword ranking tracking system (`Gordon-kw-script-v3.2.js`) that:
- Generates keyword variations from templates, niches, and locations
- Submits ranking jobs to DataForSEO API
- Tracks ranking positions for target domain (aaacwildliferemoval.com)
- Writes results to `ranking_results` sheet with position, URL, and keyword

### Current Sheets

- **kw_variants** - Generated keyword variations with coordinates
- **config** - Currently holds both configuration (batch size) and job logs (Job IDs, timestamps, status)
- **ranking_results** - Ranking positions, URLs, and keywords (columns A-C)

### Planned Improvements

> **Note:** This section should be removed after the improvements are implemented.

#### 1. Split Config Sheet

**Current State:** The `config` sheet serves dual purposes:
- Configuration parameters (batch size)
- Job execution logs (Job IDs, keywords, source rows, status, timestamps)

**Planned Change:** Split into two separate sheets:
- **config** - Configuration only (batch size, future config parameters)
- **logs** - Job execution logs (Job IDs, keywords, source rows, status, submitted/completed timestamps)

This separation will improve maintainability and make it easier to add additional configuration parameters without cluttering the logs.

#### 2. Additional Audit Columns for ranking_results

**Planned Addition:** Add two new columns to `ranking_results` sheet for auditing purposes:
- **Column J** - Location (e.g., "Riverview", "Port Richey")
- **Column K** - Service (e.g., "wildlife removal", "raccoon removal")

These columns will help with data analysis and filtering of ranking results by location and service type.

## References

- Project Brief: `rank_score_calculator_brief.md`
- Instructions Summary: `instructions_summary_transcript.ms`
- Meeting Transcript: `captions (16).srt`
- Implementation Plans: `docs/rank_score_calculator_plan.md`, `docs/keyword_scoring_fix_plan.md`

## Authors

- Gordon Ligon (project architecture)
- Daniel Capistrano (implementation)
