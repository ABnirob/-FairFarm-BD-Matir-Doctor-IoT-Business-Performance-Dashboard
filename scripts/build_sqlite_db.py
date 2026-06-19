"""
build_sqlite_db.py — loads the clean star-schema CSVs into a local SQLite
database so the SQL queries in sql/fairfarm_analysis_queries.sql can be
executed and validated (proof the SQL actually runs, not just sample code).
"""
import sqlite3
from pathlib import Path
import pandas as pd

ROOT = Path(__file__).resolve().parents[1]
CSV_DIR = ROOT / "data" / "csv"
DB_PATH = ROOT / "data" / "fairfarm_bd.db"

DB_PATH.unlink(missing_ok=True)
conn = sqlite3.connect(DB_PATH)

tables = ["dim_date", "dim_division", "dim_district", "dim_product", "dim_customer",
          "fact_web_traffic", "fact_sales", "fact_iot_readings", "fact_support_tickets",
          "stakeholder_gap_analysis"]

for t in tables:
    df = pd.read_csv(CSV_DIR / f"{t}.csv")
    df.to_sql(t, conn, if_exists="replace", index=False)
    print(f"Loaded {t}: {len(df):,} rows")

conn.commit()
conn.close()
print("\nSQLite DB created at", DB_PATH)
