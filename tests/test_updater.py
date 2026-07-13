"""Unit tests for updater.py pure helpers.

Run from the repo root:
    pip install -r requirements.txt pytest
    pytest -q

Note: importing updater.py requires the full requirements installed — it
fail-fasts (sys.exit) on missing yfinance/pandas/openpyxl/certifi.

These cover the two classes of silent failure that have bitten this project:
  1. Date sorting across year boundaries (dates must be ISO strings).
  2. Spread estimators returning None vs a value on known inputs.
"""
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).parent.parent))
import updater as u


# ── _push: ISO dates must sort correctly across New Year ─────────────────────

def test_push_sorts_across_year_boundary():
    hist = [{'date': '2026-12-30', 'illiq': 1.0},
            {'date': '2026-12-31', 'illiq': 2.0}]
    hist = u._push(hist, {'date': '2027-01-02', 'illiq': 3.0})
    assert [e['date'] for e in hist] == ['2026-12-30', '2026-12-31', '2027-01-02']


def test_push_trims_oldest_not_newest():
    # Regression: with day-month labels ('01 Dec' … '31 Dec' vs '04 Jan'),
    # January sorted BEFORE December and the trim deleted the newest entry.
    hist = [{'date': f'2026-12-{d:02d}', 'illiq': 1.0} for d in range(1, 32)]
    hist = u._push(hist, {'date': '2027-01-04', 'illiq': 9.9}, maxn=10)
    assert len(hist) == 10
    assert hist[-1]['date'] == '2027-01-04'   # newest entry survives the trim


def test_push_dedupes_same_date():
    hist = [{'date': '2026-07-10', 'illiq': 1.0}]
    hist = u._push(hist, {'date': '2026-07-10', 'illiq': 2.0})
    assert len(hist) == 1 and hist[0]['illiq'] == 2.0


# ── Roll (1984) implied spread ────────────────────────────────────────────────

def test_roll_spread_detects_bidask_bounce():
    # Alternating closes around 100 → strong negative autocovariance
    closes = [99 if i % 2 else 101 for i in range(30)]
    s = u._roll_spread(closes)
    assert s is not None and 0 < s <= 5


def test_roll_spread_none_on_trend():
    # Smooth uptrend → constant log returns → cov = 0 → no estimate
    closes = [100 * (1.01 ** i) for i in range(30)]
    assert u._roll_spread(closes) is None


def test_roll_spread_none_on_short_series():
    assert u._roll_spread([100, 101, 99]) is None


# ── Abdi-Ranaldo (2017) implied spread ────────────────────────────────────────

def test_ar_spread_positive_on_bounce():
    n = 30
    closes = [101 if i % 2 else 99 for i in range(n)]
    highs = [102] * n
    lows = [98] * n
    s = u._ar_spread(highs, lows, closes)
    assert s is not None and 0 < s <= 5


def test_ar_spread_none_on_short_series():
    assert u._ar_spread([101] * 5, [99] * 5, [100] * 5) is None


# ── _rank6 ────────────────────────────────────────────────────────────────────

def test_rank6_ascending_and_none_excluded():
    ranks = u._rank6({'a': 1.0, 'b': 3.0, 'c': 2.0, 'd': None})
    assert ranks == {'a': 1, 'c': 2, 'b': 3}


def test_rank6_descending():
    ranks = u._rank6({'a': 1.0, 'b': 3.0}, ascending=False)
    assert ranks == {'b': 1, 'a': 2}
