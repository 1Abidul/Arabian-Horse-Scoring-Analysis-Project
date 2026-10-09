# Arabian-Horse-Scoring-Analysis-Project
End-to-end analysis of judge scores and pedigrees for 40 Arabian horses: Excel dashboard, Python statistics, SQL database and Power BI model, all giving the same results.

Key findings
#	Finding	Evidence
1	Sire line is the biggest driver of score. Pogrom (112.3) and Cavali by Da Valentio (112.1) lead; Titan AS trails (106.1).	
ANOVA p ≈ 2×10⁻⁷; sire line explains 64% of score variation
2	Cavali by Da Valentio is the most consistent line. 6 of its 7 horses are Elite.	Std dev 0.69 vs 2.76 for Titan AS
3	Parent show scores predict offspring scores only weakly. The highest-scoring sire (Abha Qatar, 112.9) has mid-table offspring.	Mid-parent r = +0.23 (not significant); dam sire r = +0.06
4	Age matters a little; sex doesn't. Mature horses average 110.8 vs 108.3 for Juniors.	Age p ≈ 0.05; sex p = 0.45
5	Legs is the herd's weakest trait (avg 18.12, the worst trait for 12 of 40 horses). Type is the strongest.	Trait averages and best/worst counts
6	Judges agree only moderately and score at different levels (Judge A +1.05, Judge C −0.85 vs panel).	ICC = 0.46; Friedman p = 0.017
7	17 horses have a sex label that conflicts with their age (e.g. 6-year-old "colts").	Flagged in every tool
<p> <img src="charts/03_trait_heatmap_by_sire.png" width="49%" alt="Trait strengths by sire"> <img src="charts/06_strengths_weaknesses.png" width="49%" alt="Strongest vs weakest traits"> </p>
The data
Sheet	Rows	Contents
Horse_Profile	40	Name, sex, birth year, age group, colour, country, sire, dam, dam sire, paternal line, dam line, farm group
Judge_Scores	120	3 judges × 40 horses; 6 categories scored 0–20 (Type, Head/Neck, Body/Topline, Legs, Movement, Balance), total out of 120
Parent_Reference	15	5 sires + 10 dam sires with show scores and influence ratings

A horse's score is the average of its three judges' totals. Tiers: Elite ≥ 112 · High ≥ 106 · Developing ≥ 100 · Needs Review < 100.

What's in each tool
Excel: excel/Horse_Scoring_Bloodline_Analytics.xlsx

An 11-sheet workbook with 2,525 live formulas and no hard-coded results. Change a score or a tier cut-off and everything recalculates.

Dashboard: KPI tiles, a dynamic Top 10, a sire summary, a dropdown horse lookup and charts
Analysis sheets: Bloodline, Parent Performance, Age × Sex, Conformation, Judge Analysis
Settings: editable reference year, tier cut-offs and age rules
Formulas used: AVERAGEIF(S), COUNTIFS, MAXIFS/MINIFS, INDEX/MATCH, RANK, SUMPRODUCT, CORREL, RSQ, SLOPE, data validation, conditional formatting
Python: python/horse_scoring_analysis.ipynb

A pandas/SciPy notebook in 12 steps: validate → summarise → test → visualise → export.

Data-quality checks (missing values, duplicates, score ranges, referential integrity, sex/age logic)
Statistics: one-way ANOVA, Kruskal-Wallis, pairwise Welch t-tests with Bonferroni correction, Pearson correlation and linear regression, ICC(2,1) inter-rater agreement, Friedman test
8 charts (charts/), 10 result tables (tables/), and a star schema for Power BI (powerbi/data/)
SQL: sql/

A normalised relational database (tested on SQLite; standard SQL that ports to PostgreSQL, MySQL and SQL Server).

7 tables with primary keys, foreign keys, CHECK constraints and indexes, including a long-format trait_scores table (720 rows)
2 views: v_horse_scores (per-horse summary) and v_sire_line_summary
11 analysis queries using window functions (RANK, ROW_NUMBER … PARTITION BY), CTEs, conditional aggregation pivots and data-quality checks
Power BI: powerbi/
7 star-schema CSVs (4 dimension tables, 2 fact tables, 1 summary)
30+ DAX measures (DAX_measures.dax): averages, ranking, tiers, trait vs herd, judge vs panel, parent comparisons
Report theme, Power Query loader, and a step-by-step build guide for a 6-page report with drill-through
dim_judge ─┐                        ┌─ dim_trait
           ├─ fact_judge_scores     │
           └─ fact_trait_scores ────┘
                    │
               dim_horse ── horse_score_summary
               (dim_parent linked via DAX TREATAS)
Repository structure
arabian-horse-scoring-analysis/
├── data/          arabian_horse_score_dataset.xlsx   source data
├── excel/         Horse_Scoring_Bloodline_Analytics.xlsx
├── python/        horse_scoring_analysis.ipynb (+ .py script version)
├── charts/        8 PNG charts from the notebook
├── tables/        10 CSV result tables + key_findings.md
├── sql/           01_schema · 02_load_data · 03_views · 04_analysis_queries · horse_scoring.db
├── powerbi/       data/ (7 CSVs) · DAX_measures.dax · horse_theme.json · PowerQuery_load.m · PowerBI_Build_Guide.md
├── scripts/       build scripts that regenerate the outputs
└── requirements.txt
How to run it
bash
git clone https://github.com/<your-username>/arabian-horse-scoring-analysis.git
cd arabian-horse-scoring-analysis
pip install -r requirements.txt
To…	Run (from the repo root)
Re-run the analysis notebook	jupyter nbconvert --to notebook --execute --inplace python/horse_scoring_analysis.ipynb
Rebuild the Excel workbook	python scripts/build_excel_workbook.py data/arabian_horse_score_dataset.xlsx excel/Horse_Scoring_Bloodline_Analytics.xlsx
Regenerate the SQL data load	python scripts/generate_sql_inserts.py
Rebuild the SQLite database and run all queries	python scripts/build_sqlite_db.py
Explore the database	Open sql/horse_scoring.db in DB Browser for SQLite
Build the Power BI report	Follow powerbi/PowerBI_Build_Guide.md

The Excel build script writes formulas without cached values. Open the file in Excel (or recalculate in LibreOffice) to see the numbers.

Results cross-check
Figure	Excel	Python	SQL
Herd average total	109.84	109.84	109.84
Elite horses	11	11	11
Top horse	HA Rami Royal Flame (114.33)	✓	✓
Pogrom sire-line average	112.31	112.31	112.31
Sex/age flags	17	17	17
Limitations
Synthetic data; findings are illustrative.
Small groups: 7–10 horses per sire line, 1–5 per age × sex cell. The sire-level parent correlation uses only 5 points.
Ages assume a 2026 reference year (editable in the Excel Settings sheet).
Tools

Excel · Python (pandas, NumPy, SciPy, Matplotlib, openpyxl) · Jupyter · SQL (SQLite) · Power BI (DAX, Power Query)
