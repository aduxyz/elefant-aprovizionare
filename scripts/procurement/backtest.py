"""
Backtesting framework for procurement formulas.

For each simulation date T, computes predicted vs actual 30-day sales
for all active products, using multiple formula variants.

Usage: python -m procurement.backtest
"""
import sqlite3
import time
from datetime import datetime, timedelta
from procurement.config import DB_PATH, SIMULATION_DATES, TARGET_DAYS


def get_connection(db_path=DB_PATH):
    conn = sqlite3.connect(db_path)
    conn.row_factory = sqlite3.Row
    return conn


def get_sales_windows(conn, article_code, ref_date):
    """Get cumulative sales in each window ending at ref_date."""
    windows = {}
    for name, days in [('L1W', 7), ('LM', 30), ('L2M', 60), ('LS', 180), ('LY', 365)]:
        start = (datetime.strptime(ref_date, '%Y-%m-%d') - timedelta(days=days)).strftime('%Y-%m-%d')
        row = conn.execute(
            "SELECT COALESCE(SUM(qty), 0) as total FROM daily_sales "
            "WHERE article_code = ? AND date >= ? AND date < ?",
            (article_code, start, ref_date)
        ).fetchone()
        windows[name] = row['total']
    return windows


def get_all_windows_bulk(conn, ref_date):
    """Get sales windows for ALL products at once. Much faster than per-product queries."""
    results = {}
    window_defs = [('L1W', 7), ('LM', 30), ('L2M', 60), ('LS', 180), ('LY', 365)]

    for name, days in window_defs:
        start = (datetime.strptime(ref_date, '%Y-%m-%d') - timedelta(days=days)).strftime('%Y-%m-%d')
        rows = conn.execute(
            "SELECT article_code, SUM(qty) as total FROM daily_sales "
            "WHERE date >= ? AND date < ? GROUP BY article_code",
            (start, ref_date)
        ).fetchall()
        for r in rows:
            ac = r['article_code']
            if ac not in results:
                results[ac] = {'L1W': 0, 'LM': 0, 'L2M': 0, 'LS': 0, 'LY': 0}
            results[ac][name] = r['total']

    return results


def get_actual_sales_bulk(conn, ref_date, days=30):
    """Get actual sales in the [ref_date, ref_date + days) window for all products."""
    end = (datetime.strptime(ref_date, '%Y-%m-%d') + timedelta(days=days)).strftime('%Y-%m-%d')
    rows = conn.execute(
        "SELECT article_code, SUM(qty) as total FROM daily_sales "
        "WHERE date >= ? AND date < ? GROUP BY article_code",
        (ref_date, end)
    ).fetchall()
    return {r['article_code']: r['total'] for r in rows}


# ── Formulas ──────────────────────────────────────────────────────

def v0_simple_avg(w):
    """Baseline: average daily sales over last 30 days."""
    return w['LM'] / 30.0


def v1_weighted_blend(w):
    """60% last week rate + 40% previous month rate."""
    return w['L1W'] / 7.0 * 0.6 + max(0, (w['L2M'] - w['LM']) / 30.0) * 0.4


def v2_trend(w):
    """Trend detection with 5 branches (ported from GAS evalQuantity_)."""
    L1W, LM, L2M, LS, LY = w['L1W'], w['LM'], w['L2M'], w['LS'], w['LY']

    if LM > 50 and L1W < LM * 0.1 and L2M > 0 and LM > (L2M - LM) * 3:
        return L1W / 7.0                                                    # POST-SPIKE
    elif LM > 5 and L1W / 7 > LM / 30 * 2:
        return L1W / 7.0 * 0.7 + LM / 30.0 * 0.3                          # ACCELERATING
    elif L1W == 0 and LM == 0:
        return 0                                                             # STAGNANT
    elif LY > 20 and LM / 30 < LY / 365 * 0.5:
        return LM / 30.0 * 0.6 + LY / 365.0 * 0.4                         # DECLINING
    else:
        return L1W/7*0.35 + LM/30*0.35 + LS/180*0.15 + LY/365*0.15        # NORMAL


def v3_trend_buffered(w):
    """v2_trend + 15% safety buffer to reduce stockout risk, capped for post-spike."""
    L1W, LM, L2M, LS, LY = w['L1W'], w['LM'], w['L2M'], w['LS'], w['LY']

    if LM > 50 and L1W < LM * 0.1 and L2M > 0 and LM > (L2M - LM) * 3:
        return L1W / 7.0                                                    # POST-SPIKE: no buffer
    elif LM > 5 and L1W / 7 > LM / 30 * 2:
        return (L1W / 7.0 * 0.7 + LM / 30.0 * 0.3) * 1.1                 # ACCELERATING: 10% buffer
    elif L1W == 0 and LM == 0:
        return 0                                                             # STAGNANT
    elif LY > 20 and LM / 30 < LY / 365 * 0.5:
        return (LM / 30.0 * 0.6 + LY / 365.0 * 0.4) * 1.15               # DECLINING: 15% buffer
    else:
        return (L1W/7*0.35 + LM/30*0.35 + LS/180*0.15 + LY/365*0.15) * 1.15  # NORMAL: 15% buffer


def v4_adaptive(w):
    """Adaptive weights based on trend strength + momentum detection."""
    L1W, LM, L2M, LS, LY = w['L1W'], w['LM'], w['L2M'], w['LS'], w['LY']

    w_rate = L1W / 7.0 if L1W > 0 else 0
    m_rate = LM / 30.0 if LM > 0 else 0
    prev_m_rate = max(0, (L2M - LM)) / 30.0
    s_rate = LS / 180.0 if LS > 0 else 0
    y_rate = LY / 365.0 if LY > 0 else 0

    # Post-spike: week dropped >90% vs month, and month was unusually high
    if LM > 50 and w_rate < m_rate * 0.1 and m_rate > prev_m_rate * 3:
        return w_rate  # trust only current week

    # Stagnant
    if L1W == 0 and LM == 0:
        return y_rate * 0.15 if LY > 10 else 0

    # Compute momentum: how much is the trend changing?
    if m_rate > 0:
        momentum = w_rate / m_rate  # >1 = accelerating, <1 = decelerating
    else:
        momentum = 2.0 if w_rate > 0 else 0

    # Adaptive weights: trust recent data more when momentum is high
    if momentum > 1.5:  # strong acceleration
        daily = w_rate * 0.6 + m_rate * 0.3 + s_rate * 0.1
    elif momentum < 0.5 and LY > 20:  # strong deceleration
        daily = m_rate * 0.4 + s_rate * 0.3 + y_rate * 0.3
    else:  # stable
        daily = w_rate * 0.3 + m_rate * 0.35 + s_rate * 0.2 + y_rate * 0.15

    # Small safety buffer
    return daily * 1.1


FORMULAS = {
    'v0_avg30d': v0_simple_avg,
    'v1_weighted': v1_weighted_blend,
    'v2_trend': v2_trend,
    'v3_buffered': v3_trend_buffered,
    'v4_adaptive': v4_adaptive,
}


def run_backtest(conn, sim_date, formulas=FORMULAS):
    """Run all formulas for a single simulation date. Returns per-formula metrics."""
    windows_all = get_all_windows_bulk(conn, sim_date)
    actual_all = get_actual_sales_bulk(conn, sim_date, TARGET_DAYS)

    # Only test products that had SOME sales in 90d before sim_date AND some actual sales after
    start_90d = (datetime.strptime(sim_date, '%Y-%m-%d') - timedelta(days=90)).strftime('%Y-%m-%d')
    active_rows = conn.execute(
        "SELECT DISTINCT article_code FROM daily_sales WHERE date >= ? AND date < ?",
        (start_90d, sim_date)
    ).fetchall()
    active_articles = {r['article_code'] for r in active_rows}

    results = {}
    for fname, fn in formulas.items():
        total_pred = 0
        total_actual = 0
        abs_errors = []
        overstock_count = 0
        stockout_count = 0
        n = 0
        big_errors = []

        for ac in active_articles:
            w = windows_all.get(ac)
            if w is None:
                continue
            actual_30d = actual_all.get(ac, 0)
            predicted_daily = fn(w)
            predicted_30d = max(0, round(predicted_daily * TARGET_DAYS))

            delta = predicted_30d - actual_30d
            abs_err = abs(delta)

            total_pred += predicted_30d
            total_actual += actual_30d
            abs_errors.append(abs_err)
            n += 1

            if actual_30d > 0 and predicted_30d > actual_30d * 1.3:
                overstock_count += 1
            if actual_30d > 0 and predicted_30d < actual_30d * 0.7:
                stockout_count += 1

            if abs_err > 20 or (actual_30d > 0 and abs_err / max(actual_30d, 1) > 1.0):
                big_errors.append((ac, predicted_30d, actual_30d, delta))

        mae = sum(abs_errors) / n if n else 0
        big_errors.sort(key=lambda x: abs(x[3]), reverse=True)

        results[fname] = {
            'n_products': n,
            'total_pred': total_pred,
            'total_actual': total_actual,
            'mae': mae,
            'overstock_pct': overstock_count / n * 100 if n else 0,
            'stockout_pct': stockout_count / n * 100 if n else 0,
            'top_errors': big_errors[:30],
        }

    return results


def print_results(all_results):
    """Print comparison table across all simulation dates."""
    formula_names = list(FORMULAS.keys())

    # Header
    print(f"\n{'Date':<12}", end='')
    for fn in formula_names:
        print(f" │ {fn:>12} MAE  Over%  Out%  Pred/Act", end='')
    print()
    print("─" * (12 + len(formula_names) * 48))

    # Per-date rows
    totals = {fn: {'mae': 0, 'over': 0, 'out': 0, 'pred': 0, 'act': 0} for fn in formula_names}
    n_dates = 0

    for sim_date, results in sorted(all_results.items()):
        n_dates += 1
        print(f"{sim_date:<12}", end='')
        for fn in formula_names:
            r = results[fn]
            ratio = r['total_pred'] / r['total_actual'] if r['total_actual'] else 0
            print(f" │ {r['mae']:>12.1f} {r['overstock_pct']:5.1f} {r['stockout_pct']:5.1f} {ratio:8.2f}", end='')
            totals[fn]['mae'] += r['mae']
            totals[fn]['over'] += r['overstock_pct']
            totals[fn]['out'] += r['stockout_pct']
            totals[fn]['pred'] += r['total_pred']
            totals[fn]['act'] += r['total_actual']
        print()

    # Average row
    print("─" * (12 + len(formula_names) * 48))
    print(f"{'AVG':<12}", end='')
    for fn in formula_names:
        t = totals[fn]
        ratio = t['pred'] / t['act'] if t['act'] else 0
        print(f" │ {t['mae']/n_dates:>12.1f} {t['over']/n_dates:5.1f} {t['out']/n_dates:5.1f} {ratio:8.2f}", end='')
    print()

    # Find winner (lowest avg MAE)
    best = min(formula_names, key=lambda fn: totals[fn]['mae'])
    print(f"\n★ Best formula by MAE: {best}")
    print(f"  Avg MAE: {totals[best]['mae']/n_dates:.1f}")
    print(f"  Avg Overstock: {totals[best]['over']/n_dates:.1f}%")
    print(f"  Avg Stockout risk: {totals[best]['out']/n_dates:.1f}%")
    print(f"  Total Predicted/Actual ratio: {totals[best]['pred']/totals[best]['act']:.3f}")

    # Print top errors for best formula (last sim date)
    last_date = sorted(all_results.keys())[-1]
    top_err = all_results[last_date][best]['top_errors'][:20]
    if top_err:
        print(f"\nTop errors for {best} on {last_date}:")
        print(f"  {'article_code':<20} {'predicted':>10} {'actual':>10} {'delta':>10}")
        print(f"  {'─'*20} {'─'*10} {'─'*10} {'─'*10}")
        for ac, pred, act, delta in top_err:
            print(f"  {ac:<20} {pred:>10.0f} {act:>10.0f} {delta:>+10.0f}")


def main():
    conn = get_connection()
    all_results = {}

    for sim_date in SIMULATION_DATES:
        t0 = time.time()
        print(f"Simulating {sim_date}...", end='', flush=True)
        results = run_backtest(conn, sim_date)
        elapsed = time.time() - t0
        n = results[list(results.keys())[0]]['n_products']
        print(f" {n:,} products, {elapsed:.1f}s")
        all_results[sim_date] = results

    print_results(all_results)
    conn.close()


if __name__ == '__main__':
    main()
