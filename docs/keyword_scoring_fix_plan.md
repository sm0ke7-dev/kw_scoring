# Keyword Scoring Fix Plan

## Problem Identified

The current `KeywordScore.gs` implementation incorrectly uses a logarithmic decay curve based on keyword position in the list. This causes:
- Last keyword to get a score of 0
- Final scores to be zeroed out even for high-value locations/niches/rankings
- Misalignment with client instructions for flat/equal distribution

## Correct Implementation (Per Transcript)

**From transcript:** *"take an average so that you get the chunk and then you just say we got five keywords, they're all worth oneif of whatever that first partition is worth"*

### Expected Behavior

1. **Identify Core vs Geo Keywords**
   - Core: Pure keywords (e.g., "raccoon removal")
   - Geo: Location-modified keywords (e.g., "raccoon removal riverview")
   - Detection: Check if `kw+geo` equals `keyword` (case-insensitive) → core, else geo

2. **Equal Distribution Within Each Bucket**
   - Core bucket: All core keywords get equal share of `partitionSplit`
     - Example: 5 core keywords, partitionSplit = 0.66 → each gets 0.66/5 = 0.132
   - Geo bucket: All geo keywords get equal share of `(1 - partitionSplit)`
     - Example: 10 geo keywords, partitionSplit = 0.66 → each gets 0.34/10 = 0.034

3. **No Decay Curve for Keywords**
   - Position in list doesn't matter for keyword scoring
   - All keywords in the same bucket get the same score
   - Decay curve is only used for ranking positions (RankingScore.gs)

## Implementation Steps

### Step 1: Update KeywordScore.gs

**Remove:**
- Decay curve calculation based on position
- Normalization of decay weights

**Add:**
- Core vs geo keyword identification
  - Read `Rankings` sheet to see which keywords are core vs geo
  - Or check if keyword appears in `kw+geo` column (exact match = core)
- Count core keywords and geo keywords separately
- Calculate equal shares:
  - `coreScore = partitionSplit / coreKeywordCount`
  - `geoScore = (1 - partitionSplit) / geoKeywordCount`
- Assign scores based on core/geo classification

### Step 2: Update Keyword Sheet Structure

**Current:**
- Keywords listed in order (doesn't matter anymore)
- Scores calculated from position

**New:**
- Keywords can be in any order
- Scores based on core/geo classification
- Need to identify core vs geo for each keyword

### Step 3: Update FinalMerge.gs

**Current:**
- Uses `kw+geo` to detect if geo keyword
- Looks up keyword score from Keyword sheet

**New:**
- Should still use same lookup logic
- But scores will now be flat/equal within buckets

### Step 4: Testing

- Verify core keywords all have same score
- Verify geo keywords all have same score
- Verify core scores sum to `partitionSplit`
- Verify geo scores sum to `(1 - partitionSplit)`
- Verify no keywords get zero scores
- Verify final scores aren't zeroed out

## Questions to Clarify

1. **How to identify core vs geo?**
   - Option A: Check `Rankings` sheet `kw+geo` column - if `kw+geo` equals `keyword` (exact match) = core
   - Option B: Have a separate classification column in Keyword sheet
   - **Recommendation:** Option A (check against Rankings data)

2. **What if a keyword appears in both core and geo forms?**
   - Should it get both scores? Or just one?
   - **Recommendation:** Keyword appears once in Keyword sheet, gets one score based on its primary form

3. **What if no core keywords exist? Or no geo keywords?**
   - Handle edge case: if no core → all value goes to geo
   - If no geo → all value goes to core

## Files to Modify

1. `gas/KeywordScore.gs` - Complete rewrite of scoring logic
2. `gas/FinalMerge.gs` - May need minor updates for lookup
3. `gas/Utils.gs` - No changes needed
4. `README.md` - Already updated

## Acceptance Criteria

- ✅ All core keywords have identical scores (sum to partitionSplit)
- ✅ All geo keywords have identical scores (sum to 1 - partitionSplit)
- ✅ No keywords get zero scores (unless partitionSplit = 0 or 1)
- ✅ Final scores properly reflect location/niche/ranking value
- ✅ Keyword scores are location-independent (same keyword, same score)

