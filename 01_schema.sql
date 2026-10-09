-- =====================================================================
-- Horse Scoring & Bloodline Analytics — SQL schema
-- Standard SQL; tested on SQLite 3.45. Works on PostgreSQL / MySQL /
-- SQL Server with minor type tweaks (DECIMAL, VARCHAR lengths).
-- Run order: 01_schema.sql -> 02_load_data.sql -> 03_views.sql -> 04_analysis_queries.sql
-- =====================================================================

DROP VIEW IF EXISTS v_sire_line_summary;
DROP VIEW IF EXISTS v_horse_scores;
DROP TABLE IF EXISTS trait_scores;
DROP TABLE IF EXISTS judge_scores;
DROP TABLE IF EXISTS horses;
DROP TABLE IF EXISTS parent_reference;
DROP TABLE IF EXISTS bloodline_groups;
DROP TABLE IF EXISTS traits;
DROP TABLE IF EXISTS judges;

-- Paternal bloodline groups (one per sire line)
CREATE TABLE bloodline_groups (
    bloodline_id     INTEGER PRIMARY KEY,
    bloodline_name   VARCHAR(40) NOT NULL UNIQUE
);

-- Sires and dam sires with their (hypothetical) show scores
CREATE TABLE parent_reference (
    parent_name         VARCHAR(40) PRIMARY KEY,
    parent_type         VARCHAR(10) NOT NULL CHECK (parent_type IN ('Sire', 'Dam Sire')),
    primary_bloodline   VARCHAR(40) NOT NULL,
    parent_show_score   DECIMAL(5,1) NOT NULL,
    type_influence      VARCHAR(15),
    movement_influence  VARCHAR(15)
);

CREATE TABLE horses (
    horse_id            VARCHAR(10) PRIMARY KEY,
    registered_name     VARCHAR(60) NOT NULL,
    barn_name           VARCHAR(30),
    sex                 VARCHAR(10) NOT NULL CHECK (sex IN ('Colt','Filly','Gelding','Mare','Stallion')),
    birth_year          INTEGER NOT NULL,
    age_group           VARCHAR(15) NOT NULL,
    color               VARCHAR(15),
    country_of_foaling  VARCHAR(30),
    sire                VARCHAR(40) NOT NULL REFERENCES parent_reference(parent_name),
    dam                 VARCHAR(40),
    dam_sire            VARCHAR(40) NOT NULL REFERENCES parent_reference(parent_name),
    bloodline_id        INTEGER NOT NULL REFERENCES bloodline_groups(bloodline_id),
    dam_line            VARCHAR(40),
    farm_group          VARCHAR(30),
    notes               VARCHAR(100)
);

CREATE TABLE judges (
    judge_name  VARCHAR(20) PRIMARY KEY
);

-- Trait-level scoring categories (AHA-style halter)
CREATE TABLE traits (
    trait_code   VARCHAR(20) PRIMARY KEY,
    trait_label  VARCHAR(30) NOT NULL,
    sort_order   INTEGER NOT NULL,
    max_score    DECIMAL(4,1) NOT NULL DEFAULT 20
);

-- Wide score card: one row per horse per judge (as entered)
CREATE TABLE judge_scores (
    score_id        VARCHAR(10) PRIMARY KEY,
    horse_id        VARCHAR(10) NOT NULL REFERENCES horses(horse_id),
    judge_name      VARCHAR(20) NOT NULL REFERENCES judges(judge_name),
    type_score      DECIMAL(4,1) NOT NULL CHECK (type_score BETWEEN 0 AND 20),
    head_neck       DECIMAL(4,1) NOT NULL CHECK (head_neck BETWEEN 0 AND 20),
    body_topline    DECIMAL(4,1) NOT NULL CHECK (body_topline BETWEEN 0 AND 20),
    legs            DECIMAL(4,1) NOT NULL CHECK (legs BETWEEN 0 AND 20),
    movement        DECIMAL(4,1) NOT NULL CHECK (movement BETWEEN 0 AND 20),
    balance         DECIMAL(4,1) NOT NULL CHECK (balance BETWEEN 0 AND 20),
    judge_comments  VARCHAR(200),
    UNIQUE (horse_id, judge_name)
);

-- Long (unpivoted) trait scores: one row per horse x judge x trait.
-- Easier for trait-level analysis and Power BI slicing.
CREATE TABLE trait_scores (
    score_id    VARCHAR(10) NOT NULL REFERENCES judge_scores(score_id),
    horse_id    VARCHAR(10) NOT NULL REFERENCES horses(horse_id),
    judge_name  VARCHAR(20) NOT NULL REFERENCES judges(judge_name),
    trait_code  VARCHAR(20) NOT NULL REFERENCES traits(trait_code),
    score       DECIMAL(4,1) NOT NULL CHECK (score BETWEEN 0 AND 20),
    PRIMARY KEY (score_id, trait_code)
);

CREATE INDEX ix_horses_sire     ON horses(sire);
CREATE INDEX ix_horses_damsire  ON horses(dam_sire);
CREATE INDEX ix_scores_horse    ON judge_scores(horse_id);
CREATE INDEX ix_traits_horse    ON trait_scores(horse_id, trait_code);
