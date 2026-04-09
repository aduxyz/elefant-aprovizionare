"""
Create aggregate tables in elefant-erp.db for fast analytical queries.
Original tables remain untouched.

Usage: python -m procurement.normalize_db
"""
import sqlite3
import time
from procurement.config import DB_PATH, PRICE_THRESHOLD


def normalize(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.execute("PRAGMA journal_mode=WAL")
    cur = conn.cursor()

    t0 = time.time()

    # --- daily_sales: net qty per article_code per day ---
    print("Creating daily_sales...")
    cur.execute("DROP TABLE IF EXISTS daily_sales")
    cur.execute("""
        CREATE TABLE daily_sales (
            article_code TEXT NOT NULL,
            date TEXT NOT NULL,
            qty REAL NOT NULL,
            PRIMARY KEY (article_code, date)
        )
    """)
    cur.execute(f"""
        INSERT INTO daily_sales (article_code, date, qty)
        SELECT article_code, order_date,
            SUM(CASE WHEN order_line_status = 'Facturat' THEN order_quantity ELSE 0 END)
            - SUM(CASE WHEN order_line_status = 'Returnat' THEN order_quantity ELSE 0 END)
        FROM sales
        WHERE selling_price >= {PRICE_THRESHOLD}
          AND order_line_status IN ('Facturat', 'Returnat')
        GROUP BY article_code, order_date
    """)
    n_sales = cur.rowcount
    print(f"  → {n_sales:,} rows ({time.time() - t0:.1f}s)")

    # --- daily_purchases ---
    print("Creating daily_purchases...")
    cur.execute("DROP TABLE IF EXISTS daily_purchases")
    cur.execute("""
        CREATE TABLE daily_purchases (
            article_code TEXT NOT NULL,
            date TEXT NOT NULL,
            qty REAL NOT NULL,
            PRIMARY KEY (article_code, date)
        )
    """)
    cur.execute("""
        INSERT INTO daily_purchases (article_code, date, qty)
        SELECT article_code, date, SUM(quantity)
        FROM purchases
        GROUP BY article_code, date
    """)
    n_purch = cur.rowcount
    print(f"  → {n_purch:,} rows ({time.time() - t0:.1f}s)")

    # --- daily_purchase_returns ---
    print("Creating daily_purchase_returns...")
    cur.execute("DROP TABLE IF EXISTS daily_purchase_returns")
    cur.execute("""
        CREATE TABLE daily_purchase_returns (
            article_code TEXT NOT NULL,
            date TEXT NOT NULL,
            qty REAL NOT NULL,
            PRIMARY KEY (article_code, date)
        )
    """)
    cur.execute("""
        INSERT INTO daily_purchase_returns (article_code, date, qty)
        SELECT article_code, date, SUM(quantity)
        FROM purchase_returns
        GROUP BY article_code, date
    """)
    n_ret = cur.rowcount
    print(f"  → {n_ret:,} rows ({time.time() - t0:.1f}s)")

    conn.commit()

    # Verify
    for tbl in ['daily_sales', 'daily_purchases', 'daily_purchase_returns']:
        cnt = cur.execute(f"SELECT count(*) FROM {tbl}").fetchone()[0]
        print(f"  {tbl}: {cnt:,} rows")

    conn.close()
    print(f"Done in {time.time() - t0:.1f}s")


if __name__ == '__main__':
    normalize()
