"""
Backtest formula v2 vs v1.

Formula v2: MAX(0, VZ_medie_zi*30 * spike_factor - Stoc)
  - VZ_medie_zi = SalesL1W/7*0.6 + max(SalesL2M-SalesLM,0)/30*0.4
  - spike_factor = 0.7 dacă SalesLM > 2*SalesLM_prev (și prev>=5)
                   1.0 altfel

Excel: =MAX(0, S2*30 * IF(AND(3*O2>2*P2, P2-O2>=5), 0.7, 1) - I2)
"""
import sqlite3
from datetime import date, timedelta
import statistics
from collections import defaultdict

conn = sqlite3.connect('elefant-erp.db')

print("Pre-loading stock_history...", flush=True)
stock_by_art = defaultdict(list)
for art, dt, stoc in conn.execute(
    'SELECT article_code, stock_date, stock_online FROM stock_history ORDER BY article_code, stock_date'
).fetchall():
    stock_by_art[art].append((dt, stoc))
print(f"  {sum(len(v) for v in stock_by_art.values())} rows, {len(stock_by_art)} articles", flush=True)

first_sale_dict = dict(conn.execute(
    'SELECT article_code, MIN(SUBSTR(order_date,1,10)) FROM sales GROUP BY article_code'
).fetchall())


def get_stock_at(art, date_str):
    entries = stock_by_art.get(art)
    if not entries:
        return 0
    lo, hi, result = 0, len(entries) - 1, None
    while lo <= hi:
        mid = (lo + hi) // 2
        if entries[mid][0] <= date_str:
            result = entries[mid][1]
            lo = mid + 1
        else:
            hi = mid - 1
    return result if result is not None else 0


def load_sales_window(date_from, date_to):
    """Net sales per article in [date_from, date_to]."""
    rows = conn.execute("""
        SELECT article_code,
               SUM(CASE WHEN order_line_status IN ('Facturat','Comanda in picking') THEN order_quantity
                        WHEN order_line_status = 'Returnat' THEN -order_quantity
                        ELSE 0 END)
        FROM sales
        WHERE SUBSTR(order_date,1,10) >= ? AND SUBSTR(order_date,1,10) <= ?
        GROUP BY article_code
    """, (date_from, date_to)).fetchall()
    return {r[0]: max(0, r[1]) for r in rows if r[1] is not None}


def vz_medie(s1w, slm, s2m):
    trend = max(s2m - slm, 0) / 30 * 0.4
    return s1w / 7 * 0.6 + trend


def formula_v1(slm, stoc, **kw):
    return max(0, slm - stoc)


def formula_v2(slm, stoc, s1w=0, s2m=0, **kw):
    vz = vz_medie(s1w, slm, s2m)
    slm_prev = s2m - slm
    spike = slm > 2 * slm_prev and slm_prev >= 5
    factor = 0.7 if spike else 1.0
    return max(0, vz * 30 * factor - stoc)


def backtest(test_date_str, formula_fn, label):
    t = test_date_str
    d = date.fromisoformat(t)
    t_m7   = str(d - timedelta(days=7))
    t_m30  = str(d - timedelta(days=30))
    t_m60  = str(d - timedelta(days=60))
    t_p30  = str(d + timedelta(days=30))
    cutoff_new = str(d - timedelta(days=60))

    print(f"\n  Loading {t}...", flush=True)
    s1w_dict  = load_sales_window(t_m7,  str(d - timedelta(days=1)))
    slm_dict  = load_sales_window(t_m30, str(d - timedelta(days=1)))
    s2m_dict  = load_sales_window(t_m60, str(d - timedelta(days=1)))
    real_dict = load_sales_window(t,     t_p30)

    # Only articles with SalesLM > 0
    universe = {a: v for a, v in slm_dict.items() if v > 0}

    results = []
    for art, slm in universe.items():
        if first_sale_dict.get(art, '9999') > cutoff_new:
            continue
        stoc = max(0, get_stock_at(art, t))
        s1w  = s1w_dict.get(art, 0)
        s2m  = s2m_dict.get(art, 0)
        cant = formula_fn(slm, stoc, s1w=s1w, s2m=s2m)
        real = real_dict.get(art, 0)
        if real == 0:
            continue
        avail = stoc + cant
        dp = abs(avail - real) / real * 100
        results.append({
            'slm': slm, 'stoc': stoc, 'cant': cant, 'real': real,
            'avail': avail, 'dp': dp,
            'dir': 'SUPRA' if avail > real else 'SUB',
            'rvs': real / slm if slm > 0 else 0,
            's1w': s1w, 's2m': s2m,
        })

    n = len(results)
    if not n:
        print(f"    no articles"); return []

    over   = [x for x in results if x['dp'] > 30]
    sup    = [x for x in over if x['dir'] == 'SUPRA']
    sub    = [x for x in over if x['dir'] == 'SUB']
    avg_d  = statistics.mean([x['dp'] for x in results])
    med_d  = statistics.median([x['dp'] for x in results])

    print(f"\n{'='*65}")
    print(f"  {t}  |  {label}")
    print(f"{'='*65}")
    print(f"  Articole:      {n}")
    print(f"  Avg diff%:     {avg_d:.1f}%")
    print(f"  Median diff%:  {med_d:.1f}%")
    print(f"  >30%:          {len(over)} ({len(over)/n*100:.1f}%)")
    print(f"    SUPRA:       {len(sup)} ({len(sup)/n*100:.1f}%)")
    print(f"    SUB:         {len(sub)} ({len(sub)/n*100:.1f}%)")
    return results


print("\n--- FORMULA v1: max(0, SalesLM - Stoc) ---")
r1 = {}
for dt in ['2025-07-15', '2025-09-15', '2026-01-22']:
    r1[dt] = backtest(dt, formula_v1, "v1: max(0, SalesLM - Stoc)")

print("\n\n--- FORMULA v2: max(0, VZ*30*spike_factor - Stoc) ---")
r2 = {}
for dt in ['2025-07-15', '2025-09-15', '2026-01-22']:
    r2[dt] = backtest(dt, formula_v2, "v2: max(0, VZ*30*spike_factor - Stoc)")

# Comparison summary
print("\n\n" + "="*65)
print("  COMPARATIE v1 vs v2")
print("="*65)
print(f"  {'Perioadă':<12} {'v1 Med%':>8} {'v2 Med%':>8} {'v1 >30%':>8} {'v2 >30%':>8}")
for dt in ['2025-07-15', '2025-09-15', '2026-01-22']:
    if not r1[dt] or not r2[dt]:
        continue
    m1 = statistics.median([x['dp'] for x in r1[dt]])
    m2 = statistics.median([x['dp'] for x in r2[dt]])
    o1 = sum(1 for x in r1[dt] if x['dp'] > 30) / len(r1[dt]) * 100
    o2 = sum(1 for x in r2[dt] if x['dp'] > 30) / len(r2[dt]) * 100
    print(f"  {dt:<12} {m1:>8.1f}% {m2:>8.1f}% {o1:>7.1f}% {o2:>7.1f}%")

conn.close()
print("\nDone.")
