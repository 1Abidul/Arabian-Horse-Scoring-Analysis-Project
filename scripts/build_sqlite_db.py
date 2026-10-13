"""Builds sql/horse_scoring.db by running the SQL files in order, then runs the analysis queries.
Usage (from the repo root): python scripts/build_sqlite_db.py
"""
import re
import sqlite3
from pathlib import Path

db = Path('sql/horse_scoring.db')
db.unlink(missing_ok=True)
con = sqlite3.connect(db)
con.execute('PRAGMA foreign_keys = ON')
for name in ['01_schema', '02_load_data', '03_views']:
    con.executescript(Path(f'sql/{name}.sql').read_text())
assert not con.execute('PRAGMA foreign_key_check').fetchall(), 'foreign key violations found'

queries = Path('sql/04_analysis_queries.sql').read_text()
for block in [b.strip() for b in re.split(r';\s*\n', queries) if 'SELECT' in b]:
    title = [line for line in block.splitlines() if line.startswith('-- Q')]
    rows = con.execute(block).fetchall()
    print(f"{title[-1] if title else ''}  ->  {len(rows)} rows")
con.commit()
print(f'built {db}')
