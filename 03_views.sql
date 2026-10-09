-- =====================================================================
-- Reusable views
-- =====================================================================

-- One row per horse: trait averages across judges, total, judge spread, parent scores
CREATE VIEW v_horse_scores AS
SELECT
    h.horse_id,
    h.registered_name,
    h.sex,
    h.age_group,
    h.sire,
    h.dam_sire,
    b.bloodline_name                                   AS paternal_line,
    h.dam_line,
    h.farm_group,
    ROUND(AVG(s.type_score), 2)                        AS avg_type,
    ROUND(AVG(s.head_neck), 2)                         AS avg_head_neck,
    ROUND(AVG(s.body_topline), 2)                      AS avg_body_topline,
    ROUND(AVG(s.legs), 2)                              AS avg_legs,
    ROUND(AVG(s.movement), 2)                          AS avg_movement,
    ROUND(AVG(s.balance), 2)                           AS avg_balance,
    AVG(s.type_score + s.head_neck + s.body_topline + s.legs + s.movement + s.balance) AS avg_total,
    MAX(s.type_score + s.head_neck + s.body_topline + s.legs + s.movement + s.balance)
      - MIN(s.type_score + s.head_neck + s.body_topline + s.legs + s.movement + s.balance) AS judge_spread,
    ps.parent_show_score                               AS sire_score,
    pd.parent_show_score                               AS dam_sire_score,
    (ps.parent_show_score + pd.parent_show_score) / 2.0 AS mid_parent_score
FROM horses h
JOIN bloodline_groups b  ON b.bloodline_id = h.bloodline_id
JOIN judge_scores s      ON s.horse_id = h.horse_id
JOIN parent_reference ps ON ps.parent_name = h.sire
JOIN parent_reference pd ON pd.parent_name = h.dam_sire
GROUP BY h.horse_id, h.registered_name, h.sex, h.age_group, h.sire, h.dam_sire, b.bloodline_name,
         h.dam_line, h.farm_group, ps.parent_show_score, pd.parent_show_score;

-- One row per sire line
CREATE VIEW v_sire_line_summary AS
SELECT
    v.sire,
    v.paternal_line,
    COUNT(*)                                            AS horses,
    ROUND(AVG(v.avg_total), 2)                          AS avg_total,
    ROUND(MIN(v.avg_total), 2)                          AS min_total,
    ROUND(MAX(v.avg_total), 2)                          AS max_total,
    SUM(CASE WHEN v.avg_total >= 112 THEN 1 ELSE 0 END) AS elite_horses,
    MAX(v.sire_score)                                   AS sire_show_score,
    ROUND(AVG(v.avg_total) - MAX(v.sire_score), 2)      AS offspring_vs_sire
FROM v_horse_scores v
GROUP BY v.sire, v.paternal_line;
