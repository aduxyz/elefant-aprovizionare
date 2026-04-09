#!/usr/bin/env python3
"""
rebuild_db.py
=============
Reconstruiește elefant-erp.db complet din fișierele sursă XLS.

Ordinea pașilor:
  1. purchases, purchase_returns, sales    ← arhivă/export-erp/*.xlsx
  2. procurement_erp                       ← arhivă/EBS D.O.I. (BRUT).xlsx
  3. stock_history                         ← derivat (SQL window functions)
  4. daily_sales, daily_purchases, ...     ← agregate derivate

Referință stoc: REF_DATE = 2026-03-22 (snapshot din EBS BRUT de pe 2026-03-23).
Dacă actualizezi sursele cu date mai recente, actualizează REF_DATE și EBS_FILE.

Usage:
  python data/rebuild_db.py            # reconstruiește din zero
  python data/rebuild_db.py --dry-run  # validează că fișierele sursă există, nu scrie nimic
"""

import argparse
import sqlite3
import sys
import time
from pathlib import Path
from datetime import datetime, timedelta

# ── Paths ────────────────────────────────────────────────────────────────────
ROOT        = Path(__file__).parent.parent   # repo root
DB_FILE     = ROOT / 'elefant-erp.db'
EXPORT_DIR  = Path(__file__).parent / 'export-erp'
EBS_FILE    = Path(__file__).parent / 'erp-export-DOI.xlsx'
REF_DATE    = '2026-03-22'   # data snapshot-ului EBS BRUT

# ── Sursele XLS (structura GDrive: 01_achizitii/, 02_vanzari/) ───────────────
PURCHASES_GLOB        = '01_achizitii/011_Achizitii_DOI_*.xlsx'
PURCHASE_RETURNS_GLOB = '01_achizitii/012_Achizitii_RETUR_DOI_*.xlsx'
SALES_GLOB            = '02_vanzari/021_Vanzari_DOI_*.xlsx'

EXCEL_EPOCH = datetime(1899, 12, 30)


# ─────────────────────────────────────────────────────────────────────────────
# Helpers
# ─────────────────────────────────────────────────────────────────────────────

def log(msg):
    print(msg, flush=True)


def excel_date_to_str(val):
    try:
        import pandas as pd
        if pd.isna(val):
            return None
    except Exception:
        if val is None:
            return None
    try:
        return (EXCEL_EPOCH + timedelta(days=int(val))).strftime('%Y-%m-%d')
    except Exception:
        return str(val) if val is not None else None


def check_sources():
    missing = []
    if not EBS_FILE.exists():
        missing.append(str(EBS_FILE))
    for pat in (PURCHASES_GLOB, PURCHASE_RETURNS_GLOB, SALES_GLOB):
        if not list(EXPORT_DIR.glob(pat)):
            missing.append(f'{EXPORT_DIR}/{pat} (no files matched)')
    return missing


# ─────────────────────────────────────────────────────────────────────────────
# Pasul 1: imports XLS → purchases, purchase_returns, sales
# ─────────────────────────────────────────────────────────────────────────────

def create_raw_tables(con):
    con.executescript("""
    CREATE TABLE IF NOT EXISTS purchases (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        date          TEXT,
        doc_type      TEXT,
        article_code  TEXT,
        article_desc  TEXT,
        partner       TEXT,
        document      TEXT,
        quantity      REAL,
        price         REAL,
        net_value     REAL,
        total_value   REAL,
        warehouse     TEXT,
        pkl_code      TEXT
    );
    CREATE TABLE IF NOT EXISTS purchase_returns (
        id            INTEGER PRIMARY KEY AUTOINCREMENT,
        date          TEXT,
        doc_type      TEXT,
        title         TEXT,
        article_code  TEXT,
        article_desc  TEXT,
        partner       TEXT,
        document      TEXT,
        quantity      REAL,
        price         REAL,
        net_value     REAL,
        total_value   REAL,
        warehouse     TEXT,
        pkl_code      TEXT
    );
    CREATE TABLE IF NOT EXISTS sales (
        id                INTEGER PRIMARY KEY AUTOINCREMENT,
        order_ref         TEXT,
        order_date        TEXT,
        buyer             TEXT,
        order_quantity    REAL,
        order_net_value   REAL,
        order_type        TEXT,
        order_id          TEXT,
        article_code      TEXT,
        article_desc      TEXT,
        manufacturer      TEXT,
        selling_price     REAL,
        order_line_status TEXT,
        sku               TEXT
    );
    """)
    con.commit()


def import_purchases(con, path):
    import pandas as pd
    df = pd.read_excel(path)
    rows = [
        (excel_date_to_str(r.iloc[0]),
         str(r.iloc[1]) if pd.notna(r.iloc[1]) else None,
         str(r.iloc[2]) if pd.notna(r.iloc[2]) else None,
         str(r.iloc[3]) if pd.notna(r.iloc[3]) else None,
         str(r.iloc[4]) if pd.notna(r.iloc[4]) else None,
         str(r.iloc[5]) if pd.notna(r.iloc[5]) else None,
         float(r.iloc[6]) if pd.notna(r.iloc[6]) else None,
         float(r.iloc[7]) if pd.notna(r.iloc[7]) else None,
         float(r.iloc[8]) if pd.notna(r.iloc[8]) else None,
         float(r.iloc[9]) if pd.notna(r.iloc[9]) else None,
         str(r.iloc[10]) if pd.notna(r.iloc[10]) else None,
         str(r.iloc[11]) if pd.notna(r.iloc[11]) else None)
        for _, r in df.iterrows()
    ]
    con.executemany("""
        INSERT INTO purchases (date,doc_type,article_code,article_desc,partner,document,
            quantity,price,net_value,total_value,warehouse,pkl_code)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?)
    """, rows)
    con.commit()
    return len(rows)


def import_purchase_returns(con, path):
    import pandas as pd
    df = pd.read_excel(path)
    rows = [
        (excel_date_to_str(r.iloc[0]),
         str(r.iloc[1]) if pd.notna(r.iloc[1]) else None,
         str(r.iloc[2]) if pd.notna(r.iloc[2]) else None,
         str(r.iloc[3]) if pd.notna(r.iloc[3]) else None,
         str(r.iloc[4]) if pd.notna(r.iloc[4]) else None,
         str(r.iloc[5]) if pd.notna(r.iloc[5]) else None,
         str(r.iloc[6]) if pd.notna(r.iloc[6]) else None,
         float(r.iloc[7]) if pd.notna(r.iloc[7]) else None,
         float(r.iloc[8]) if pd.notna(r.iloc[8]) else None,
         float(r.iloc[9]) if pd.notna(r.iloc[9]) else None,
         float(r.iloc[10]) if pd.notna(r.iloc[10]) else None,
         str(r.iloc[11]) if pd.notna(r.iloc[11]) else None,
         str(r.iloc[12]) if pd.notna(r.iloc[12]) else None)
        for _, r in df.iterrows()
    ]
    con.executemany("""
        INSERT INTO purchase_returns (date,doc_type,title,article_code,article_desc,partner,
            document,quantity,price,net_value,total_value,warehouse,pkl_code)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, rows)
    con.commit()
    return len(rows)


def import_sales(con, path):
    import pandas as pd
    df = pd.read_excel(path)
    rows = [
        (str(r.iloc[0]) if pd.notna(r.iloc[0]) else None,
         excel_date_to_str(r.iloc[1]),
         str(r.iloc[2]) if pd.notna(r.iloc[2]) else None,
         float(r.iloc[3]) if pd.notna(r.iloc[3]) else None,
         float(r.iloc[4]) if pd.notna(r.iloc[4]) else None,
         str(r.iloc[5]) if pd.notna(r.iloc[5]) else None,
         str(r.iloc[6]) if pd.notna(r.iloc[6]) else None,
         str(r.iloc[7]) if pd.notna(r.iloc[7]) else None,
         str(r.iloc[8]) if pd.notna(r.iloc[8]) else None,
         str(r.iloc[9]) if pd.notna(r.iloc[9]) else None,
         float(r.iloc[10]) if pd.notna(r.iloc[10]) else None,
         str(r.iloc[11]) if pd.notna(r.iloc[11]) else None,
         str(r.iloc[12]) if pd.notna(r.iloc[12]) else None)
        for _, r in df.iterrows()
    ]
    con.executemany("""
        INSERT INTO sales (order_ref,order_date,buyer,order_quantity,order_net_value,
            order_type,order_id,article_code,article_desc,manufacturer,
            selling_price,order_line_status,sku)
        VALUES (?,?,?,?,?,?,?,?,?,?,?,?,?)
    """, rows)
    con.commit()
    return len(rows)


def step1_import_xls(con):
    log('\n[1/4] Import XLS → purchases, purchase_returns, sales')
    create_raw_tables(con)

    total = {'purchases': 0, 'purchase_returns': 0, 'sales': 0}

    for path in sorted(EXPORT_DIR.glob(PURCHASE_RETURNS_GLOB)):
        n = import_purchase_returns(con, path)
        total['purchase_returns'] += n
        log(f'  purchase_returns  {path.name}: {n:,}')

    for path in sorted(EXPORT_DIR.glob(PURCHASES_GLOB)):
        n = import_purchases(con, path)
        total['purchases'] += n
        log(f'  purchases         {path.name}: {n:,}')

    for path in sorted(EXPORT_DIR.glob(SALES_GLOB)):
        n = import_sales(con, path)
        total['sales'] += n
        log(f'  sales             {path.name}: {n:,}')

    log(f'  TOTAL purchases={total["purchases"]:,}  '
        f'purchase_returns={total["purchase_returns"]:,}  '
        f'sales={total["sales"]:,}')


# ─────────────────────────────────────────────────────────────────────────────
# Pasul 2: EBS BRUT → procurement_erp
# ─────────────────────────────────────────────────────────────────────────────

def step2_import_ebs(con):
    log(f'\n[2/4] Import EBS BRUT → procurement_erp  ({EBS_FILE.name})')
    import openpyxl
    wb = openpyxl.load_workbook(EBS_FILE, read_only=True, data_only=True)
    ws = wb.active
    rows = list(ws.iter_rows(min_row=2, values_only=True))
    wb.close()
    log(f'  {len(rows):,} rânduri citite')

    con.execute('DROP TABLE IF EXISTS procurement_erp')
    con.execute("""
        CREATE TABLE procurement_erp (
            sku TEXT, ean TEXT, article_code TEXT, article_desc TEXT,
            author TEXT, supplier TEXT, rrp REAL, discount REAL,
            stock_online REAL, avail_rec REAL,
            sales_l1w REAL, sales_lm REAL, sales_l2m REAL, sales_ls REAL, sales_ly REAL,
            publish_date TEXT, category TEXT, subcategory TEXT,
            manufacturer TEXT, supplier_code TEXT, total_reserved REAL,
            avail_mm_auchan REAL, collection TEXT, format TEXT, age_group TEXT,
            sku_abon TEXT, purchases_all REAL, sales_all REAL,
            out_of_print TEXT, unavailable TEXT,
            avail_custf REAL, avail_stpr REAL, purchase_price REAL
        )
    """)
    con.executemany(
        'INSERT INTO procurement_erp VALUES (' + ','.join(['?'] * 33) + ')',
        [r[:33] for r in rows]
    )
    con.commit()
    log(f'  {len(rows):,} rânduri importate în procurement_erp')


# ─────────────────────────────────────────────────────────────────────────────
# Pasul 3: stock_history (din build-stock-history.py, inline)
# ─────────────────────────────────────────────────────────────────────────────

def step3_build_stock_history(con):
    log(f'\n[3/4] Construiesc stock_history (REF_DATE={REF_DATE})')

    sql_build = f"""
CREATE TABLE stock_history AS
WITH
movements AS (
    SELECT SUBSTR(date, 1, 10) AS stock_date, article_code,  quantity AS delta
    FROM purchases WHERE article_code IS NOT NULL
    UNION ALL
    SELECT SUBSTR(date, 1, 10), article_code, -quantity
    FROM purchase_returns WHERE article_code IS NOT NULL
    UNION ALL
    SELECT SUBSTR(order_date, 1, 10), article_code, -order_quantity
    FROM sales WHERE article_code IS NOT NULL
      AND order_line_status IN ('Facturat', 'Comanda in picking')
    UNION ALL
    SELECT SUBSTR(order_date, 1, 10), article_code,  order_quantity
    FROM sales WHERE article_code IS NOT NULL
      AND order_line_status = 'Returnat'
),
daily AS (
    SELECT stock_date, article_code,
           CAST(ROUND(SUM(delta)) AS INTEGER) AS delta
    FROM movements GROUP BY stock_date, article_code
),
cumulative AS (
    SELECT stock_date, article_code, delta,
           SUM(delta) OVER (
               PARTITION BY article_code ORDER BY stock_date
               ROWS BETWEEN UNBOUNDED PRECEDING AND CURRENT ROW
           ) AS cum_sum
    FROM daily
),
ref_cum AS (
    SELECT article_code,
           MAX(CASE WHEN stock_date <= '{REF_DATE}' THEN cum_sum END) AS ref_cum
    FROM cumulative GROUP BY article_code
),
ref_stock AS (
    SELECT article_code, CAST(stock_online AS INTEGER) AS ref_stock
    FROM procurement_erp WHERE article_code IS NOT NULL
)
SELECT c.stock_date, c.article_code,
       CAST(COALESCE(r.ref_stock,0) + c.cum_sum - COALESCE(rc.ref_cum,0) AS INTEGER) AS stock_online
FROM cumulative c
LEFT JOIN ref_stock r  ON r.article_code  = c.article_code
LEFT JOIN ref_cum   rc ON rc.article_code = c.article_code;
"""
    sql_static = f"""
INSERT OR IGNORE INTO stock_history (stock_date, article_code, stock_online)
SELECT '{REF_DATE}', article_code, CAST(stock_online AS INTEGER)
FROM procurement_erp
WHERE article_code IS NOT NULL
  AND article_code NOT IN (SELECT DISTINCT article_code FROM stock_history);
"""
    con.execute('DROP TABLE IF EXISTS stock_history')
    con.executescript(sql_build)
    con.commit()
    (n1,) = con.execute('SELECT COUNT(*) FROM stock_history').fetchone()
    log(f'  {n1:,} rânduri din mișcări')

    con.executescript(sql_static)
    con.commit()
    (n2,) = con.execute('SELECT COUNT(*) FROM stock_history').fetchone()
    log(f'  +{n2-n1:,} articole statice → total {n2:,} rânduri')

    con.execute('CREATE INDEX IF NOT EXISTS idx_sh_article ON stock_history(article_code)')
    con.execute('CREATE INDEX IF NOT EXISTS idx_sh_date    ON stock_history(stock_date)')
    con.commit()


# ─────────────────────────────────────────────────────────────────────────────
# Pasul 4: agregate zilnice
# ─────────────────────────────────────────────────────────────────────────────

def step4_normalize(con):
    log('\n[4/4] Agregate zilnice (daily_sales, daily_purchases, daily_purchase_returns)')
    con.executescript("""
    DROP TABLE IF EXISTS daily_sales;
    CREATE TABLE daily_sales AS
        SELECT SUBSTR(order_date,1,10) AS date, article_code,
               SUM(CASE WHEN order_line_status IN ('Facturat','Comanda in picking')
                        THEN order_quantity ELSE 0 END) AS quantity_sold,
               SUM(CASE WHEN order_line_status = 'Returnat'
                        THEN order_quantity ELSE 0 END) AS quantity_returned
        FROM sales WHERE article_code IS NOT NULL
        GROUP BY SUBSTR(order_date,1,10), article_code;
    CREATE UNIQUE INDEX IF NOT EXISTS idx_ds ON daily_sales(article_code, date);

    DROP TABLE IF EXISTS daily_purchases;
    CREATE TABLE daily_purchases AS
        SELECT SUBSTR(date,1,10) AS date, article_code, SUM(quantity) AS quantity
        FROM purchases WHERE article_code IS NOT NULL
        GROUP BY SUBSTR(date,1,10), article_code;
    CREATE UNIQUE INDEX IF NOT EXISTS idx_dp ON daily_purchases(article_code, date);

    DROP TABLE IF EXISTS daily_purchase_returns;
    CREATE TABLE daily_purchase_returns AS
        SELECT SUBSTR(date,1,10) AS date, article_code, SUM(quantity) AS quantity
        FROM purchase_returns WHERE article_code IS NOT NULL
        GROUP BY SUBSTR(date,1,10), article_code;
    CREATE UNIQUE INDEX IF NOT EXISTS idx_dpr ON daily_purchase_returns(article_code, date);
    """)
    con.commit()
    (ds,)  = con.execute('SELECT COUNT(*) FROM daily_sales').fetchone()
    (dp,)  = con.execute('SELECT COUNT(*) FROM daily_purchases').fetchone()
    (dpr,) = con.execute('SELECT COUNT(*) FROM daily_purchase_returns').fetchone()
    log(f'  daily_sales={ds:,}  daily_purchases={dp:,}  daily_purchase_returns={dpr:,}')


# ─────────────────────────────────────────────────────────────────────────────
# Validare finală
# ─────────────────────────────────────────────────────────────────────────────

def validate(con):
    log('\n── Validare ─────────────────────────────────────────────')
    checks = [
        ('purchases',              'SELECT COUNT(*) FROM purchases',              194_000, 210_000),
        ('purchase_returns',       'SELECT COUNT(*) FROM purchase_returns',        10_000,  20_000),
        ('sales',                  'SELECT COUNT(*) FROM sales',                  900_000,1_200_000),
        ('procurement_erp',        'SELECT COUNT(*) FROM procurement_erp',         10_000,  20_000),
        ('stock_history',          'SELECT COUNT(*) FROM stock_history',        1_000_000,5_000_000),
        ('daily_sales',            'SELECT COUNT(*) FROM daily_sales',             500_000,2_000_000),
    ]
    ok = True
    for name, sql, lo, hi in checks:
        (n,) = con.execute(sql).fetchone()
        status = '✓' if lo <= n <= hi else '⚠'
        if status == '⚠':
            ok = False
        log(f'  {status} {name}: {n:,}  (așteptat {lo:,}–{hi:,})')

    (min_s, max_s) = con.execute('SELECT MIN(stock_date), MAX(stock_date) FROM stock_history').fetchone()
    log(f'  stock_history interval: {min_s} → {max_s}')
    return ok


# ─────────────────────────────────────────────────────────────────────────────
# Main
# ─────────────────────────────────────────────────────────────────────────────

def main():
    parser = argparse.ArgumentParser(description='Reconstruiește elefant-erp.db din XLS-uri sursă')
    parser.add_argument('--dry-run', action='store_true', help='Verifică sursele fără să scrie nimic')
    args = parser.parse_args()

    log(f'DB target: {DB_FILE}')
    log(f'REF_DATE:  {REF_DATE}  (snapshot EBS BRUT)')
    log(f'EBS file:  {EBS_FILE.name}')

    missing = check_sources()
    if missing:
        log('\n❌ Fișiere sursă lipsă:')
        for m in missing:
            log(f'   {m}')
        sys.exit(1)
    log('✓ Toate fișierele sursă găsite')

    if args.dry_run:
        log('\n--dry-run: nimic scris.')
        return

    if DB_FILE.exists():
        log(f'\nȘterg {DB_FILE.name} existent...')
        DB_FILE.unlink()
        for ext in ('-shm', '-wal'):
            p = DB_FILE.with_suffix('.db' + ext)
            if p.exists():
                p.unlink()

    t0 = time.time()
    con = sqlite3.connect(DB_FILE)
    con.execute('PRAGMA journal_mode=WAL')
    con.execute('PRAGMA synchronous=NORMAL')
    con.execute('PRAGMA temp_store=MEMORY')
    con.execute('PRAGMA cache_size=-131072')  # 128 MB

    step1_import_xls(con)
    step2_import_ebs(con)
    step3_build_stock_history(con)
    step4_normalize(con)

    ok = validate(con)
    con.close()

    elapsed = time.time() - t0
    size_mb = DB_FILE.stat().st_size / 1024 / 1024
    log(f'\n{"✓" if ok else "⚠"} Gata în {elapsed:.0f}s  →  {DB_FILE.name}  ({size_mb:.0f} MB)')


if __name__ == '__main__':
    main()
