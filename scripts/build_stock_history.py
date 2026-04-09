#!/usr/bin/env python3
"""
build-stock-history.py
======================
Construiește tabela `stock_history` în elefant-erp.db.

Schema:
    stock_history (
        stock_date    TEXT    NOT NULL,   -- YYYY-MM-DD
        article_code  TEXT    NOT NULL,
        stock_online  INTEGER NOT NULL,
        PRIMARY KEY (stock_date, article_code)
    )

Logică SQL:
    1. delta zilnic per (data, articol):
         NIR (+)  ← purchases
         NAC (−)  ← purchase_returns
         vânzări (−) / retururi clienți (+)  ← sales
    2. cumsum per articol (ORDER BY stock_date) = cum_sum[D]
    3. ref_cum = cum_sum la ultima zi ≤ REF_DATE per articol
    4. stock[D] = ref_stock + cum_sum[D] − ref_cum
       unde ref_stock vine din procurement_erp.stock_online (snapshot 2026-03-22)

    Rânduri generate: câte o zi per articol, NUMAI pentru zilele cu mișcări.
    Artcolele fără nicio mișcare primesc un singur rând la REF_DATE.
    Interogare stoc la data D arbitrară:
        SELECT stock_online FROM stock_history
        WHERE article_code=X AND stock_date<=D
        ORDER BY stock_date DESC LIMIT 1;

Surse (exclusiv DB):
    procurement_erp, purchases, purchase_returns, sales
"""

import sqlite3
import time
from pathlib import Path

DB_FILE  = Path(__file__).parent.parent / 'elefant-erp.db'
REF_DATE = '2026-03-22'

SQL_BUILD = f"""
CREATE TABLE stock_history AS
WITH
-- ── 1. Toate mișcările, normalizate ca delta ─────────────────────────────────
movements AS (
    SELECT SUBSTR(date, 1, 10)       AS stock_date,
           article_code,
           quantity                  AS delta
    FROM purchases
    WHERE article_code IS NOT NULL

    UNION ALL

    SELECT SUBSTR(date, 1, 10),
           article_code,
           -quantity
    FROM purchase_returns
    WHERE article_code IS NOT NULL

    UNION ALL

    SELECT SUBSTR(order_date, 1, 10),
           article_code,
           -order_quantity
    FROM sales
    WHERE article_code IS NOT NULL
      AND order_line_status IN ('Facturat', 'Comanda in picking')

    UNION ALL

    SELECT SUBSTR(order_date, 1, 10),
           article_code,
           order_quantity
    FROM sales
    WHERE article_code IS NOT NULL
      AND order_line_status = 'Returnat'
),
-- ── 2. Delta zilnic agregat per (data, articol) ───────────────────────────────
daily AS (
    SELECT stock_date, article_code,
           CAST(ROUND(SUM(delta)) AS INTEGER) AS delta
    FROM movements
    GROUP BY stock_date, article_code
),
-- ── 3. Cumsum cumulat per articol (cronologic) ────────────────────────────────
cumulative AS (
    SELECT stock_date,
           article_code,
           delta,
           SUM(delta) OVER (
               PARTITION BY article_code
               ORDER BY stock_date
               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
           ) AS cum_sum
    FROM daily
),
-- ── 4. Valoarea cumulată la REF_DATE (ultima zi ≤ REF_DATE per articol) ───────
ref_cum AS (
    SELECT article_code,
           MAX(CASE WHEN stock_date <= '{REF_DATE}' THEN cum_sum END) AS ref_cum
    FROM cumulative
    GROUP BY article_code
),
-- ── 5. Stoc de referință din BRUT (procurement_erp) ──────────────────────────
ref_stock AS (
    SELECT article_code,
           CAST(stock_online AS INTEGER) AS ref_stock
    FROM procurement_erp
    WHERE article_code IS NOT NULL
)
-- ── 6. Stoc reconstruit la fiecare dată cu mișcări ───────────────────────────
SELECT
    c.stock_date,
    c.article_code,
    CAST(
        COALESCE(r.ref_stock, 0) + c.cum_sum - COALESCE(rc.ref_cum, 0)
    AS INTEGER) AS stock_online
FROM cumulative  c
LEFT JOIN ref_stock r  ON r.article_code  = c.article_code
LEFT JOIN ref_cum   rc ON rc.article_code = c.article_code;
"""

SQL_STATIC = f"""
INSERT OR IGNORE INTO stock_history (stock_date, article_code, stock_online)
SELECT '{REF_DATE}', article_code, CAST(stock_online AS INTEGER)
FROM procurement_erp
WHERE article_code IS NOT NULL
  AND article_code NOT IN (
      SELECT DISTINCT article_code FROM stock_history
  );
"""


def main():
    t0 = time.time()
    con = sqlite3.connect(DB_FILE)
    con.execute('PRAGMA journal_mode=WAL')
    con.execute('PRAGMA synchronous=NORMAL')
    con.execute('PRAGMA temp_store=MEMORY')
    con.execute('PRAGMA cache_size=-131072')   # 128 MB cache

    print(f'[1/3] Construiesc stock_history (SQL CTE + window functions)...')
    con.execute('DROP TABLE IF EXISTS stock_history')
    con.executescript(SQL_BUILD)
    con.commit()
    (n1,) = con.execute('SELECT COUNT(*) FROM stock_history').fetchone()
    print(f'      {n1:,} rânduri din mișcări')

    print(f'[2/3] Adaug articolele fără mișcări (rând REF_DATE)...')
    con.executescript(SQL_STATIC)
    con.commit()
    (n2,) = con.execute('SELECT COUNT(*) FROM stock_history').fetchone()
    print(f'      +{n2-n1:,} articole statice → total {n2:,} rânduri')

    print('[3/3] Indexuri...')
    con.execute('CREATE INDEX IF NOT EXISTS idx_sh_article ON stock_history(article_code)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_sh_date    ON stock_history(stock_date)')
    con.commit()

    (n_rows,)  = con.execute('SELECT COUNT(*) FROM stock_history').fetchone()
    (n_art,)   = con.execute('SELECT COUNT(DISTINCT article_code) FROM stock_history').fetchone()
    (n_days,)  = con.execute('SELECT COUNT(DISTINCT stock_date) FROM stock_history').fetchone()
    (min_d,)   = con.execute('SELECT MIN(stock_date) FROM stock_history').fetchone()
    (max_d,)   = con.execute('SELECT MAX(stock_date) FROM stock_history').fetchone()
    con.close()

    print(f'\n✓ Gata în {time.time()-t0:.1f}s → {DB_FILE}')
    print(f'  {n_rows:,} rânduri | {n_art:,} articole | {n_days} date distincte')
    print(f'  Interval: {min_d} → {max_d}')
    print()
    print('Exemplu query stoc la o dată oarecare:')
    print("  SELECT stock_online FROM stock_history")
    print("  WHERE article_code='PF0002568485' AND stock_date<='2026-03-15'")
    print("  ORDER BY stock_date DESC LIMIT 1;")


if __name__ == '__main__':
    main()
