-- =====================================================================
-- Analysis queries (window functions need SQLite 3.25+, PostgreSQL, MySQL 8+, SQL Server)
-- =====================================================================

-- Q1. Overall ranking with score tier
SELECT
    RANK() OVER (ORDER BY avg_total DESC)  AS rank_no,
    registered_name, sire, sex, age_group,
    ROUND(avg_total, 2)                    AS avg_total,
    CASE WHEN avg_total >= 112 THEN 'Elite'
         WHEN avg_total >= 106 THEN 'High'
         WHEN avg_total >= 100 THEN 'Developing'
         ELSE 'Needs Review' END           AS score_tier
FROM v_horse_scores
ORDER BY rank_no
LIMIT 10;

-- Q2. Sire line comparison
SELECT * FROM v_sire_line_summary ORDER BY avg_total DESC;

-- Q3. Best horse in each sire line
SELECT sire, registered_name, ROUND(avg_total, 2) AS avg_total
FROM (
    SELECT v.*, ROW_NUMBER() OVER (PARTITION BY sire ORDER BY avg_total DESC) AS rn
    FROM v_horse_scores v
) t
WHERE rn = 1
ORDER BY avg_total DESC;

-- Q4. Trait profile by sire vs herd average (positive = strength)
WITH herd AS (
    SELECT t.trait_code, AVG(t.score) AS herd_avg
    FROM trait_scores t GROUP BY t.trait_code
)
SELECT h.sire, ts.trait_code,
       ROUND(AVG(ts.score), 2)                AS sire_avg,
       ROUND(AVG(ts.score) - herd.herd_avg, 2) AS vs_herd
FROM trait_scores ts
JOIN horses h  ON h.horse_id = ts.horse_id
JOIN herd      ON herd.trait_code = ts.trait_code
JOIN traits tr ON tr.trait_code = ts.trait_code
GROUP BY h.sire, ts.trait_code, herd.herd_avg, tr.sort_order
ORDER BY h.sire, tr.sort_order;

-- Q5. Strongest and weakest trait for each sire line
WITH st AS (
    SELECT h.sire, ts.trait_code, AVG(ts.score) AS avg_score
    FROM trait_scores ts JOIN horses h ON h.horse_id = ts.horse_id
    GROUP BY h.sire, ts.trait_code
), ranked AS (
    SELECT st.*,
           ROW_NUMBER() OVER (PARTITION BY sire ORDER BY avg_score DESC) AS best_rn,
           ROW_NUMBER() OVER (PARTITION BY sire ORDER BY avg_score ASC)  AS worst_rn
    FROM st
)
SELECT sire,
       MAX(CASE WHEN best_rn  = 1 THEN trait_code END) AS strongest_trait,
       MAX(CASE WHEN worst_rn = 1 THEN trait_code END) AS weakest_trait
FROM ranked GROUP BY sire ORDER BY sire;

-- Q6. Dam-sire influence: offspring average vs dam-sire show score
SELECT dam_sire,
       COUNT(*)                         AS horses,
       ROUND(AVG(avg_total), 2)         AS offspring_avg,
       MAX(dam_sire_score)              AS dam_sire_score,
       ROUND(AVG(avg_total) - MAX(dam_sire_score), 2) AS offspring_vs_score
FROM v_horse_scores
GROUP BY dam_sire
ORDER BY offspring_avg DESC;

-- Q7. Age group x sex matrix
SELECT age_group,
       ROUND(AVG(CASE WHEN sex = 'Colt'     THEN avg_total END), 2) AS colt,
       ROUND(AVG(CASE WHEN sex = 'Filly'    THEN avg_total END), 2) AS filly,
       ROUND(AVG(CASE WHEN sex = 'Gelding'  THEN avg_total END), 2) AS gelding,
       ROUND(AVG(CASE WHEN sex = 'Mare'     THEN avg_total END), 2) AS mare,
       ROUND(AVG(CASE WHEN sex = 'Stallion' THEN avg_total END), 2) AS stallion,
       ROUND(AVG(avg_total), 2)                                      AS all_sexes,
       COUNT(*)                                                      AS horses
FROM v_horse_scores
GROUP BY age_group
ORDER BY CASE age_group WHEN 'Junior' THEN 1 WHEN 'Young Horse' THEN 2 ELSE 3 END;

-- Q8. Herd-wide conformation: trait averages and number of horses below 17.5
SELECT tr.trait_label,
       ROUND(AVG(x.avg_score), 2)                                 AS herd_avg,
       SUM(CASE WHEN x.avg_score < 17.5 THEN 1 ELSE 0 END)        AS horses_below_17_5
FROM (SELECT horse_id, trait_code, AVG(score) AS avg_score
      FROM trait_scores GROUP BY horse_id, trait_code) x
JOIN traits tr ON tr.trait_code = x.trait_code
GROUP BY tr.trait_label, tr.sort_order
ORDER BY herd_avg DESC;

-- Q9. Judge bias: each judge's average vs panel, per trait
WITH panel AS (SELECT trait_code, AVG(score) AS panel_avg FROM trait_scores GROUP BY trait_code)
SELECT ts.judge_name, ts.trait_code,
       ROUND(AVG(ts.score) - p.panel_avg, 2) AS vs_panel
FROM trait_scores ts JOIN panel p ON p.trait_code = ts.trait_code
GROUP BY ts.judge_name, ts.trait_code, p.panel_avg
ORDER BY ts.judge_name, ts.trait_code;

-- Q10. Horses where judges disagree most (spread of total scores)
SELECT registered_name, sire, ROUND(avg_total, 2) AS avg_total, ROUND(judge_spread, 1) AS judge_spread
FROM v_horse_scores
ORDER BY judge_spread DESC
LIMIT 5;

-- Q11. Data quality: sex label conflicts with age (colt/filly < 4, stallion/mare 4+), reference year 2026
SELECT horse_id, registered_name, sex, birth_year, 2026 - birth_year AS age, age_group
FROM horses
WHERE (sex IN ('Colt', 'Filly')    AND 2026 - birth_year >= 4)
   OR (sex IN ('Stallion', 'Mare') AND 2026 - birth_year <  4)
ORDER BY horse_id;
