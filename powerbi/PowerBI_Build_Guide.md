# Power BI Build Guide — Horse Scoring & Bloodline Analytics

This takes about 45–60 minutes in Power BI Desktop. Everything you need is in this folder:

| File | What it is |
|---|---|
| `data/*.csv` | 7 model-ready tables (star schema) exported by the Python notebook |
| `DAX_measures.dax` | All measures, ready to paste |
| `horse_theme.json` | Report theme (colours, fonts, card style) |
| `PowerQuery_load.m` | Optional M script for loading CSVs from a folder parameter |

---

## 1. Load the data

1. **Home → Get Data → Text/CSV** and load each file in `data/`:
   `dim_horse`, `dim_parent`, `dim_judge`, `dim_trait`, `fact_judge_scores`, `fact_trait_scores`, `horse_score_summary`.
2. In **Transform data**, check types: score columns = *Decimal number*, `BirthYear`/`Age`/`Rank`/`SortOrder` = *Whole number*, IDs = *Text*. Close & Apply.
3. **View → Themes → Browse for themes** → `horse_theme.json`.

## 2. Build the model (Model view)

```
                 dim_judge (Judge)
                    │1
          ┌─────────┴──────────┐
          *                    *
 fact_judge_scores      fact_trait_scores ──*──1── dim_trait (Trait)
          *                    *
          └─────────┬──────────┘
                    │1
               dim_horse (HorseID) ──1──1── horse_score_summary
                    ┊
               dim_parent  (no relationship — linked in DAX with TREATAS)
```

| From (many) | To (one) | Key | Cross-filter |
|---|---|---|---|
| fact_judge_scores | dim_horse | HorseID | Single |
| fact_judge_scores | dim_judge | Judge | Single |
| fact_trait_scores | dim_horse | HorseID | Single |
| fact_trait_scores | dim_judge | Judge | Single |
| fact_trait_scores | dim_trait | Trait | Single |
| horse_score_summary | dim_horse | HorseID | 1:1, Both |

`dim_parent` stays unrelated because a horse links to it twice (sire and dam sire). The `Sire Show Score` and `Dam Sire Show Score` measures look it up with `TREATAS`.

**Sort orders:** select `dim_trait[TraitLabel]` → *Sort by column* → `SortOrder`. For age groups, add a calculated column to `dim_horse`:

```DAX
AgeSort = SWITCH ( dim_horse[AgeGroup], "Junior", 1, "Young Horse", 2, "Mature", 3 )
```

then sort `AgeGroup` by `AgeSort`.

## 3. Add measures

Create a measures table (**Home → Enter data** → name it `_Measures` → Load), then paste each measure from `DAX_measures.dax` with **New measure**. Format `Elite %` as a percentage and the score measures to 1–2 decimals.

## 4. Report pages

Canvas 1280 × 720. Put a slicer panel on the left of every page (**Sire**, **AgeGroup**, **Sex**, **FarmGroup**) and sync it with **View → Sync slicers**.

### Page 1 · Overview
| Visual | Fields |
|---|---|
| 5 cards | `Horse Count`, `Avg Horse Total`, `Elite Horses`, `Avg Judge Spread`, `Top Horse` |
| Clustered bar: score by sire | Y: `dim_horse[Sire]` · X: `Avg Horse Total` · sort descending · X-axis start at 100 |
| Histogram-style column | X: `horse_score_summary[AvgTotal]` (binned, bin size 1) · Y: count of HorseID |
| Table: Top 10 | `RegisteredName`, `Sire`, `Sex`, `AgeGroup`, `Avg Horse Total`, `Score Tier` · Top N filter = 10 by `Avg Horse Total` |
| Donut or bar: tiers | Legend: `horse_score_summary[ScoreTier]` · Values: count of HorseID |

### Page 2 · Bloodline comparison
| Visual | Fields |
|---|---|
| Matrix heatmap | Rows: `dim_horse[Sire]` · Columns: `dim_trait[TraitLabel]` · Values: `Trait vs Herd` · conditional formatting → background colour, diverging red / grey / blue, centre 0 |
| Bar chart | Axis: `dim_horse[Sire]` · Values: `Avg Horse Total`, `Sire Show Score` (clustered) |
| Table | `Sire`, `PaternalLine`, `Horse Count`, `Avg Horse Total`, `Score Std Dev`, `Elite %`, `Offspring vs Sire`, `Strongest Trait`, `Weakest Trait` |
| Bar chart | Axis: `dim_horse[DamSire]` · Values: `Avg Horse Total` |
| Bar chart | Axis: `dim_horse[DamLine]` · Values: `Avg Horse Total` |

### Page 3 · Parent performance
| Visual | Fields |
|---|---|
| Scatter | X: `horse_score_summary[MidParentScore]` · Y: `horse_score_summary[AvgTotal]` · Details: `RegisteredName` · Legend: `Sire` · Analytics pane → Trend line |
| Scatter | X: `Sire Show Score` · Y: `Avg Horse Total` · Details: `dim_horse[Sire]` |
| Table | `Sire`, `Sire Show Score`, `Avg Horse Total`, `Offspring vs Sire` (conditional colour on the last column) |

### Page 4 · Age & sex
| Visual | Fields |
|---|---|
| Matrix | Rows: `AgeGroup` · Columns: `Sex` · Values: `Avg Horse Total`, `Horse Count` · background gradient on the score |
| Clustered column | Axis: `AgeGroup` · Legend: `Sex` · Values: `Avg Horse Total` |
| Line or column | Axis: `dim_trait[TraitLabel]` · Legend: `AgeGroup` · Values: `Avg Trait Score` |
| Card + note | `Sex/Age Mismatches`, with a text box: "Sex labels conflict with age for these horses, so read sex-class comparisons with caution." |

### Page 5 · Conformation & judges
| Visual | Fields |
|---|---|
| Diverging bar | Axis: `dim_trait[TraitLabel]` · Values: `Horses Strongest In Trait`, `Horses Weakest In Trait` |
| Column | Axis: `TraitLabel` · Values: `Avg Trait Score`, `Herd Avg Trait Score` |
| Matrix heatmap | Rows: `dim_judge[Judge]` · Columns: `TraitLabel` · Values: `Judge vs Panel (Trait)` (diverging colours) |
| Table | `RegisteredName`, `horse_score_summary[JudgeSpread]`, `Avg Horse Total` · sorted by spread descending |

### Page 6 · Horse profile (drill-through)
Add `dim_horse[RegisteredName]` to the **Drill-through** well. Then add:
- Cards: `Avg Horse Total`, `Horse Rank`, `Score Tier`, `Sire`, `DamSire`, `Strongest Trait`, `Weakest Trait`
- Clustered column: Axis `TraitLabel` · Values `Avg Trait Score`, `Herd Avg Trait Score`
- Table: `fact_judge_scores` rows for the horse (Judge, the six traits, TotalScore, JudgeComments)

Right-click any horse on another page → **Drill through → Horse profile**.

## 5. Check the numbers

Your visuals should show these values from the Python and Excel analysis:

| Check | Expected |
|---|---|
| Horse Count | 40 |
| Avg Horse Total | 109.84 |
| Elite Horses | 11 |
| Pogrom / Cavali / Abha Qatar / Van Gogh AM / Titan AS | 112.31 / 112.15 / 108.95 / 108.90 / 106.10 |
| Top Horse | HA Rami Royal Flame (114.33) |
| Herd weakest trait | Legs (18.12) |
| Sex/Age Mismatches | 17 |

## 6. Publish

**File → Save** as `Horse_Scoring_Bloodline.pbix`, then **Publish** to Power BI Service if you want to share it. For a portfolio, export each page with **File → Export → PDF** or take screenshots.
