# Rank Score Calculator

A modular Google Apps Script scoring system that quantifies the value of local SEO keywords based on **location**, **niche (service)**, **keyword type**, and **ranking performance**.

## Overview

This system calculates normalized scores across multiple dimensions to identify keyword opportunities across markets. Each scoring component is normalized independently and combined multiplicatively to produce a final normalized value.

## Key Discovery: Keyword Scoring Implementation

### The Problem We Found

**Initial Implementation (INCORRECT):**
- Applied a logarithmic decay curve based on keyword position in the Keyword sheet
- Last keyword in the list would get a score of 0
- This caused final scores to be zeroed out even for high-value locations/niches/rankings

**Correct Implementation (as per client instructions):**
- Keywords should be scored using **flat/equal distribution** within each bucket (core vs geo)
- No decay curve should be applied to keyword list position
- All core keywords get equal share of `partitionSplit`
- All geo keywords get equal share of `(1 - partitionSplit)`

### Why This Matters

From the transcript: *"take an average so that you get the chunk and then you just say we got five keywords, they're all worth oneif of whatever that first partition is worth"*

This approach is:
- More robust for smaller markets
- Doesn't rely on sparse keyword planner data
- Keeps keyword scores "brain dead" - simple and auditable
- Ensures keywords have the same score regardless of location (location is factored separately)

## Project Structure

### Scripts

1. **LocationScore.gs** - Calculates market potential using population and income data
2. **NicheScore.gs** - Normalizes Keyword Planner service values
3. **KeywordScore.gs** - Assigns keyword scores (core vs geo buckets with equal distribution) ⚠️ **NEEDS FIX**
4. **RankingScore.gs** - Applies exponential decay to ranking positions (ranks 1-10)
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
4. **Ranking Score**: `exp(-kappa × (rank - 1))` for ranks 1-10, else 0 → normalized
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

⚠️ **Keyword scoring needs to be fixed** - Currently uses decay curve instead of flat/equal distribution.

See `docs/rank_score_calculator_plan.md` for implementation details.

## References

- Project Brief: `rank_score_calculator_brief.md`
- Instructions Summary: `instructions_summary_transcript.ms`
- Meeting Transcript: `captions (16).srt`

## Authors

- Gordon Ligon (project architecture)
- Daniel Capistrano (implementation)

