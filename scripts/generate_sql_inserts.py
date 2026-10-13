"""Writes sql/02_load_data.sql from the source workbook.
Usage (from the repo root): python scripts/generate_sql_inserts.py
"""
import pandas as pd

SRC = 'data/arabian_horse_score_dataset.xlsx'
OUT = 'sql/02_load_data.sql'
TRAITS = ['Type', 'Head_Neck', 'Body_Topline', 'Legs', 'Movement', 'Balance']
LABELS = {'Head_Neck': 'Head / Neck', 'Body_Topline': 'Body / Topline'}

p = pd.read_excel(SRC, 'Horse_Profile')
s = pd.read_excel(SRC, 'Judge_Scores')
r = pd.read_excel(SRC, 'Parent_Reference')


def q(v):
    """Format a Python value as a SQL literal."""
    if pd.isna(v):
        return 'NULL'
    if isinstance(v, (int, float)) and not isinstance(v, bool):
        return str(v)
    return "'" + str(v).replace("'", "''") + "'"


bid = {b: i + 1 for i, b in enumerate(sorted(p.PaternalLine.unique()))}
out = ['-- Data load generated from arabian_horse_score_dataset.xlsx', 'BEGIN TRANSACTION;', '']
out += [f"INSERT INTO bloodline_groups VALUES ({i}, {q(b)});" for b, i in bid.items()] + ['']
out += [f"INSERT INTO parent_reference VALUES ({q(x.ParentName)}, {q(x.ParentType)}, {q(x.PrimaryBloodline)}, "
        f"{x.HypotheticalShowScore}, {q(x.TypeInfluence)}, {q(x.MovementInfluence)});" for x in r.itertuples()] + ['']
out += [f"INSERT INTO horses VALUES ({q(x.HorseID)}, {q(x.RegisteredName)}, {q(x.BarnName)}, {q(x.Sex)}, {x.BirthYear}, "
        f"{q(x.AgeGroup)}, {q(x.Color)}, {q(x.CountryOfFoaling)}, {q(x.Sire)}, {q(x.Dam)}, {q(x.DamSire)}, "
        f"{bid[x.PaternalLine]}, {q(x.DamLine)}, {q(x.FarmGroup)}, {q(x.Notes)});" for x in p.itertuples()] + ['']
out += [f"INSERT INTO judges VALUES ({q(j)});" for j in sorted(s.Judge.unique())] + ['']
out += [f"INSERT INTO traits VALUES ({q(t)}, {q(LABELS.get(t, t))}, {i + 1}, 20);" for i, t in enumerate(TRAITS)] + ['']
out += [f"INSERT INTO judge_scores VALUES ({q(x.ScoreID)}, {q(x.HorseID)}, {q(x.Judge)}, {x.Type}, {x.Head_Neck}, "
        f"{x.Body_Topline}, {x.Legs}, {x.Movement}, {x.Balance}, {q(x.JudgeComments)});" for x in s.itertuples()] + ['']
for x in s.itertuples():
    for t in TRAITS:
        out.append(f"INSERT INTO trait_scores VALUES ({q(x.ScoreID)}, {q(x.HorseID)}, {q(x.Judge)}, {q(t)}, {getattr(x, t)});")
out += ['', 'COMMIT;']
open(OUT, 'w').write('\n'.join(out) + '\n')
print(f'wrote {OUT}')
