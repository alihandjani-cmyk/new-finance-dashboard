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


# ── Universe / ADV filter helpers ─────────────────────────────────────────────

def test_clean_ticker_strips_footnotes():
    assert u._clean_ticker('AAPL[1]') == 'AAPL'
    assert u._clean_ticker('MC(a)') == 'MC'
    assert u._clean_ticker('') == ''
    assert u._clean_ticker('nan') == ''


def test_clean_ticker_rejects_long_strings():
    # Strings > 15 chars are not tickers (probably company names)
    assert u._clean_ticker('ThisIsTooLongToBeATicker') == ''


def test_find_col_case_insensitive():
    import pandas as pd
    df = pd.DataFrame({'Ticker': ['A'], 'Company': ['Foo']})
    assert u._find_col(df, u._TICKER_ALIASES) == 'Ticker'
    assert u._find_col(df, u._NAME_ALIASES)   == 'Company'


def test_find_col_returns_none_when_missing():
    import pandas as pd
    df = pd.DataFrame({'Price': [1.0], 'Volume': [1000]})
    assert u._find_col(df, u._TICKER_ALIASES) is None


def test_get_universe_falls_back_to_hardcoded():
    # With empty universes dict, should return fallback tickers
    dash = {'universes': {}}
    tickers, names = u._get_universe(dash, 'six')
    assert len(tickers) > 0
    assert 'NESN.SW' in tickers   # known fallback ticker


def test_get_universe_uses_dynamic_when_present():
    dash = {'universes': {'six': {
        'tickers': ['NESN.SW', 'NOVN.SW'],
        'names':   {'NESN.SW': 'Nestlé', 'NOVN.SW': 'Novartis'},
        'n_universe': 2, 'n_constituents': 50, 'refreshed': '2026-07-13',
    }}}
    tickers, names = u._get_universe(dash, 'six')
    assert tickers == ['NESN.SW', 'NOVN.SW']
    assert names['NESN.SW'] == 'Nestlé'


def test_build_universe_returns_none_on_empty_constituents(monkeypatch):
    # If _fetch_constituents returns empty, build_universe should return None
    monkeypatch.setattr(u, '_fetch_constituents', lambda ex_key: {})
    result = u.build_universe('six', {'GBP': 1.27, 'EUR': 1.08, 'CHF': 1.11, 'USD': 1.0})
    assert result is None
