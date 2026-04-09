import sqlite3
from datetime import date, timedelta
import statistics
from collections import defaultdict

conn = sqlite3.connect('elefant-erp.db')

print("Pre-loading stock_history into memory...", flush=True)
# Load all stock_history sorted by article + date
stock_all = conn.execute("""
    SELECT article_code, stock_date, stock_online
    FROM stock_history
    ORDER BY article_code, stock_date
""").fetchall()

# Build dict: article_code -> sorted list of (stock_date, stock_online)
stock_by_art = defaultdict(list)
for art, dt, stoc in stock_all:
    stock_by_art[art].append((dt, stoc))
print(f"  Loaded {len(stock_all)} rows for {len(stock_by_art)} articles", flush=True)

print("Pre-loading first sale dates...", flush=True)
first_sale_rows = conn.execute("""
    SELECT article_code, MIN(SUBSTR(order_date,1,10))
    FROM sales GROUP BY article_code
""").fetchall()
first_sale_dict = {r[0]: r[1] for r in first_sale_rows}
print(f"  Loaded {len(first_sale_dict)} articles", flush=True)


def get_stock_at(art, date_str):
    """Latest stock_online for article on or before date_str."""
    entries = stock_by_art.get(art)
    if not entries:
        return 0
    # Binary search for last date <= date_str
    lo, hi = 0, len(entries) - 1
    result = None
    while lo <= hi:
        mid = (lo + hi) // 2
        if entries[mid][0] <= date_str:
            result = entries[mid][1]
            lo = mid + 1
        else:
            hi = mid - 1
    return result if result is not None else 0


def backtest(test_date_str, formula_fn, label="v1"):
    t = test_date_str
    t_minus30 = str(date.fromisoformat(t) - timedelta(days=30))
    t_plus30   = str(date.fromisoformat(t) + timedelta(days=30))
    cutoff_new = str(date.fromisoformat(t) - timedelta(days=60))

    print(f"\nLoading sales for {t}...", flush=True)

    sales_lm_rows = conn.execute("""
        SELECT article_code,
               SUM(CASE WHEN order_line_status IN ('Facturat','Comanda in picking') THEN order_quantity
                        WHEN order_line_status = 'Returnat' THEN -order_quantity
                        ELSE 0 END) AS qty
        FROM sales
        WHERE SUBSTR(order_date,1,10) >= ? AND SUBSTR(order_date,1,10) < ?
        GROUP BY article_code
        HAVING qty > 0
    """, (t_minus30, t)).fetchall()
    sales_lm_dict = {r[0]: r[1] for r in sales_lm_rows}
    print(f"  SalesLM articles: {len(sales_lm_dict)}", flush=True)

    real_rows = conn.execute("""
        SELECT article_code,
               SUM(CASE WHEN order_line_status IN ('Facturat','Comanda in picking') THEN order_quantity
                        WHEN order_line_status = 'Returnat' THEN -order_quantity
                        ELSE 0 END) AS qty
        FROM sales
        WHERE SUBSTR(order_date,1,10) >= ? AND SUBSTR(order_date,1,10) <= ?
        GROUP BY article_code
    """, (t, t_plus30)).fetchall()
    real_dict = {r[0]: max(0, r[1]) for r in real_rows}

    results = []
    for art, slm in sales_lm_dict.items():
        if art not in first_sale_dict or first_sale_dict[art] > cutoff_new:
            continue
        stoc = max(0, get_stock_at(art, t))
        cantitate = formula_fn(slm, stoc)
        real = real_dict.get(art, 0)
        if real == 0:
            continue
        total_avail = stoc + cantitate
        diff_pct = abs(total_avail - real) / real * 100
        direction = 'SUPRA' if total_avail > real else 'SUB'
        results.append({
            'art': art, 'slm': slm, 'stoc': stoc,
            'cantitate': cantitate, 'real': real,
            'total_avail': total_avail, 'diff_pct': diff_pct,
            'direction': direction,
            'real_vs_slm': real / slm if slm > 0 else 0,
        })

    total = len(results)
    if total == 0:
        print(f"  No eligible articles for {test_date_str}"); return []

    over30 = [r for r in results if r['diff_pct'] > 30]
    over30_supra = [r for r in over30 if r['direction'] == 'SUPRA']
    over30_sub   = [r for r in over30 if r['direction'] == 'SUB']
    avg_diff     = statistics.mean([r['diff_pct'] for r in results])
    median_diff  = statistics.median([r['diff_pct'] for r in results])

    print(f"\n{'='*65}")
    print(f"  Perioadă: {t}  |  Formula: {label}")
    print(f"{'='*65}")
    print(f"  Articole analizate:  {total}")
    print(f"  Avg diff%:           {avg_diff:.1f}%")
    print(f"  Median diff%:        {median_diff:.1f}%")
    print(f"  Abateri > 30%:       {len(over30)} ({len(over30)/total*100:.1f}%)")
    print(f"    din care SUPRA:    {len(over30_supra)} ({len(over30_supra)/total*100:.1f}%)")
    print(f"    din care SUB:      {len(over30_sub)} ({len(over30_sub)/total*100:.1f}%)")

    print(f"\n  Analiza SUPRA:")
    supra_stoc_acoperit = [r for r in over30_supra if r['stoc'] >= r['real']]
    supra_slm_exag      = [r for r in over30_supra if r['real_vs_slm'] < 0.5]
    print(f"    Stoc >= vânzări reale (comandat inutil):        {len(supra_stoc_acoperit)}")
    print(f"    Real < 50% din SalesLM (luna prev. spike/sezon):{len(supra_slm_exag)}")

    print(f"\n  Analiza SUB:")
    sub_real_gt_slm   = [r for r in over30_sub if r['real'] > r['slm']]
    sub_cantitate_0   = [r for r in over30_sub if r['cantitate'] == 0]
    print(f"    Real > SalesLM (vânzări au crescut):           {len(sub_real_gt_slm)}")
    print(f"    Cantitate=0 (stoc>=SalesLM dar s-a vândut mai mult): {len(sub_cantitate_0)}")

    worst = sorted(over30, key=lambda x: x['diff_pct'], reverse=True)[:10]
    print(f"\n  Top 10 abateri:")
    print(f"  {'Stoc':>6} {'SLM':>6} {'Cant':>6} {'Real':>6} {'Avail':>6} {'Diff%':>7} {'Dir':>6}")
    for r in worst:
        print(f"  {r['stoc']:>6.0f} {r['slm']:>6.0f} {r['cantitate']:>6.0f} {r['real']:>6.0f} {r['total_avail']:>6.0f} {r['diff_pct']:>7.1f}% {r['direction']:>6}")

    return results


def formula_v1(slm, stoc):
    return max(0, slm - stoc)


for dt in ['2025-04-27', '2025-09-15', '2026-01-22']:
    backtest(dt, formula_v1, label="v1: max(0, SalesLM - Stoc)")

conn.close()
print("\nDone.")
