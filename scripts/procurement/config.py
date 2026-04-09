import os

DB_PATH = os.path.join(os.path.dirname(__file__), '..', '..', '..', 'elefant-erp.db')

# Minimum selling price to count as a real sale (excludes gifts/promos)
PRICE_THRESHOLD = 5.0

# Target stock coverage in days
TARGET_DAYS = 30

# Backtesting simulation dates
SIMULATION_DATES = [
    '2025-07-01', '2025-08-01', '2025-09-01', '2025-10-01',
    '2025-11-01', '2025-12-01', '2026-01-01', '2026-02-01',
]

# Sales window definitions (days back from reference date)
WINDOWS = {
    'L1W': 7,
    'LM': 30,
    'L2M': 60,
    'LS': 180,
    'LY': 365,
}
