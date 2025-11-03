# Project Brief: Rank Score Calculator (Local SEO Keyword Scoring System)

## Objective
Develop a modular scoring system that quantifies the value of local SEO keywords based on **location**, **niche (service)**, **keyword type**, and **ranking performance**. Each scoring component is normalized independently and combined multiplicatively to produce a final normalized value that reflects keyword opportunity across markets.

---

## Deliverables
Five primary scripts will be implemented, each responsible for computing and logging a specific score layer:

| # | Script | Purpose | Output |
|---|---------|----------|--------|
| 1 | **Location Score Script** | Calculates each location's market potential using population and income data. | Normalized location scores (sum = 1) |
| 2 | **Niche Score Script** | Normalizes Keyword Planner service values (e.g., raccoon, bat, squirrel removal). | Normalized niche scores (sum = 1) |
| 3 | **Keyword Score Script** | Assigns weighted keyword values via a logarithmic decay function between core and geo keywords. | Normalized keyword scores (sum = 1) |
| 4 | **Ranking Score Script** | Applies exponential decay to ranking positions (ranks 1–10). | Normalized rank modifier scores (sum = 1) |
| 5 | **Final Merge Script** | Multiplies all normalized scores and re-normalizes to 1. | Final normalized score per keyword row |

---

## Calculation Rules and Formulas

### 1. Location Score
```text
numerical = (income - poverty_line) * population
normalized = numerical / SUM(all numerical)
```
**Parameters:**
- Poverty line = 32,000 (configurable)
- Output columns: `numerical`, `normalized score`

### 2. Niche Score
```text
normalized = keyword_planner_score / SUM(all keyword_planner_scores)
```
**Sheet:** `Niche`

### 3. Keyword Score
Based on a logarithmic decay curve:
```javascript
const y = x => 1 - Math.log(1 + kappa * x) / Math.log(1 + kappa);
```
**Parameters:**
- `split` – fraction of total value allocated to core vs geo keywords
- `kappa` – curve steepness constant

**Normalization:** All keyword scores are divided by the total area under the curve to ensure their sum equals 1.

### 4. Ranking Score
```text
rank_score = e^(-kappa * (rank - 1))
if rank > 10 → rank_score = 0
normalized = rank_score / SUM(all rank_scores)
```
**Parameter:** `kappa` from config (`rank modification kappa`)

### 5. Final Score
```text
final_raw = location_score × niche_score × keyword_score × ranking_score
final_normalized = final_raw / SUM(all final_raw)
```
**Output Columns:**
`service`, `geo`, `keyword`, `location_score`, `niche_score`, `keyword_score`, `ranking_score`, `final_raw`, `final_normalized`

---

## Sheet Structures

### Config
| partition split | kappa | rank modification kappa |
|-----------------|--------|--------------------------|
| 0.66 | 1 | 10 |

### Keyword
| Services | Keywords |
|-----------|-----------|
| wildlife removal | Wildlife Pest Control |
| raccoon removal | Raccoon Removal |

### Location
| target | state | lat | long | population | income | numerical | normalized score |
|--------|--------|-----|------|-------------|---------|------------|------------------|

### Niche
| Service | Keyword Planner Score | Normalization |
|----------|----------------------|----------------|

### Rankings
| service | geo | keyword | kw+geo | lat | long | Ranking Position | Ranking URL | keyword score | ranking modifier score | final score |

---

## Coding Conditions
- **Platform:** Google Apps Script (V8, Spreadsheet backend)
- **Modularity:** Each script executes independently, computes results in memory, and writes an audit log to its respective sheet.
- **Idempotency:** Rerunning any script from the start regenerates identical normalized scores.
- **Normalization Check:** Every script logs to console ensuring `SUM(normalized) ≈ 1.000000`.
- **Merge Logic:**
  - `geo` (Rankings) → `target` (Location)
  - `service` → `Service` (Niche)
  - `keyword` → `Keywords` (Keyword)
  - Missing matches default to 0.

---

## Expected Output (Final Sheet Example)
| service | geo | keyword | location_score | niche_score | keyword_score | ranking_score | final_raw | final_normalized |
|----------|-----|----------|----------------|--------------|----------------|----------------|-------------|------------------|
| wildlife removal | Riverview | Wildlife Pest Control | 0.0881 | 0.4285 | 0.0152 | 0.231 | 0.00013 | 0.0238 |
| raccoon removal | Riverview | Raccoon Removal | 0.0881 | 0.4285 | 0.0179 | 0.263 | 0.00018 | 0.0329 |

---

## Validation Criteria
- ✅ Each normalized column sums ≈ 1.000000.
- ✅ Final normalized sum = 1.
- ✅ Configurable parameters adjust without changing code.
- ✅ Outputs auditable and regenerable.

---

## Optional Extensions
- **Dashboard:** Chart top `final_normalized` scores by service and location.
- **Market Share Analysis:** Aggregate normalized scores by competitor URL.
- **Data Export:** Generate JSON export of `Final` tab for API or reporting use.

---

## Ownership & Context
- **Original Designers:** Gordon Ligon (project architecture), Daniel Capistrano (implementation)
- **Goal:** Build an auditable, repeatable SEO scoring engine for identifying keyword opportunities across multiple service areas.