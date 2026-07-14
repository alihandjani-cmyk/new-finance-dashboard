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
from datetime import date, timedelta
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


# ── _trim_stale_dates: drop leftover dates orphaned by a fetch gap ───────────

def test_trim_stale_dates_drops_isolated_old_entry():
    # Regression: a stale '2026-07-03' lingered in ENX volume history for days
    # after a mid-week fetch gap, instead of being pushed out once only 4 new
    # dates (not 5) had accumulated since.
    dates = ['2026-07-03', '2026-07-08', '2026-07-09', '2026-07-10', '2026-07-13']
    assert u._trim_stale_dates(dates) == ['2026-07-08', '2026-07-09', '2026-07-10', '2026-07-13']


def test_trim_stale_dates_keeps_full_contiguous_window():
    dates = ['2026-07-07', '2026-07-08', '2026-07-09', '2026-07-10', '2026-07-13']
    assert u._trim_stale_dates(dates) == dates


def test_trim_stale_dates_caps_at_n():
    dates = ['2026-07-06', '2026-07-07', '2026-07-08', '2026-07-09', '2026-07-10', '2026-07-13']
    assert u._trim_stale_dates(dates) == ['2026-07-07', '2026-07-08', '2026-07-09', '2026-07-10', '2026-07-13']


def test_trim_stale_dates_empty_input():
    assert u._trim_stale_dates([]) == []


# ── _clean_ticker: numeric-code exchanges (TSE) vs letter-code exchanges ─────

def test_clean_ticker_rejects_digit_led_by_default():
    # Western index pages: a stray numeric cell (footnote, etc.) must not be
    # treated as a ticker.
    assert u._clean_ticker('7203') == ''


def test_clean_ticker_allow_numeric_plain_code():
    assert u._clean_ticker('7203', allow_numeric=True) == '7203'


def test_clean_ticker_allow_numeric_alphanumeric_code():
    # TSE's newer 4-char codes end in a letter once numeric codes run out.
    assert u._clean_ticker('543A', allow_numeric=True) == '543A'


def test_clean_ticker_allow_numeric_extracts_from_prefixed_text():
    # Regression: naive '(...)' stripping deletes the code itself when the
    # whole 'TYO: 7203' is wrapped in one set of parens.
    assert u._clean_ticker('TYO: 7203', allow_numeric=True) == '7203'
    assert u._clean_ticker('(TYO: 543A)', allow_numeric=True) == '543A'


def test_clean_ticker_allow_numeric_rejects_junk():
    assert u._clean_ticker('N/A', allow_numeric=True) == ''
    assert u._clean_ticker('nan', allow_numeric=True) == ''


# ── _clean_ticker: zero-padded numeric exchanges (HKEX) ──────────────────────

def test_clean_ticker_numeric_pad_extracts_and_pads():
    # HKEX's Wikipedia table lists bare codes like 'SEHK: 5', but Yahoo
    # requires the 4-digit zero-padded form ('0005.HK').
    assert u._clean_ticker('SEHK: 5', numeric_pad=4) == '0005'
    assert u._clean_ticker('SEHK: 388', numeric_pad=4) == '0388'


def test_clean_ticker_numeric_pad_full_width_unchanged():
    assert u._clean_ticker('SEHK: 9999', numeric_pad=4) == '9999'


def test_clean_ticker_numeric_pad_rejects_junk():
    assert u._clean_ticker('N/A', numeric_pad=4) == ''
    assert u._clean_ticker('nan', numeric_pad=4) == ''


# ── _backfill_new_exchanges: adding an exchange to a pre-existing dash ───────

def _old_dash(keys):
    """Minimal dashboard_data.json shape as it would exist before a new
    exchange was added to EXCHANGES."""
    return {
        'universes':  {k: {'tickers': []} for k in keys},
        'gainers':    {k: [] for k in keys},
        'market_cap': {k: {'date': None, 'currency': 'GBP', 'top10': []} for k in keys},
        'vol':        {k: {'currency': 'GBP B', 'dates': ['2026-07-13'], 'value': [5.5]} for k in keys},
        'vol_comparable': {k: {
            'currency': 'GBP B', 'currency_usd': 'USD B',
            'dates': [], 'value': [], 'value_usd': [], 'shares': [], 'n_universe': 0,
        } for k in keys},
        'amihud':         {k: {'history': [], 'marketAvg': None} for k in keys},
        'ar_spread':      {k: {'history': [], 'marketAvg': None} for k in keys},
        'turnover_ratio': {k: {'history': []} for k in keys},
        'current_ranking': {'date': None, 'ranks': {
            k: {'vol': None, 'illiq': None, 'tr': None, 'ar': None, 'composite': None} for k in keys
        }},
    }


def test_backfill_new_exchanges_reproduces_the_keyerror_without_the_fix():
    # Regression: adding 'tse' to EXCHANGES without backfilling an existing
    # dashboard_data.json crashed the very first run with KeyError('tse') the
    # moment main() tried dash['vol'][ex_key].get(...) for the new exchange.
    old_keys = [k for k in u.EXCHANGES if k != 'tse']
    dash = _old_dash(old_keys)
    try:
        dash['vol']['tse'].get('dates', [])
        assert False, 'expected KeyError before backfill'
    except KeyError:
        pass


def test_backfill_new_exchanges_adds_missing_keys():
    old_keys = [k for k in u.EXCHANGES if k != 'tse']
    dash = _old_dash(old_keys)
    u._backfill_new_exchanges(dash)
    for section in ('universes', 'gainers', 'market_cap', 'vol', 'vol_comparable',
                     'amihud', 'ar_spread', 'turnover_ratio'):
        assert 'tse' in dash[section], f'tse missing from {section}'
    assert 'tse' in dash['current_ranking']['ranks']
    assert dash['vol']['tse']['currency'] == 'JPY B'
    assert dash['market_cap']['tse']['currency'] == 'JPY'


def test_backfill_new_exchanges_does_not_clobber_existing_data():
    old_keys = [k for k in u.EXCHANGES if k != 'tse']
    dash = _old_dash(old_keys)
    u._backfill_new_exchanges(dash)
    assert dash['vol']['lse']['dates'] == ['2026-07-13']
    assert dash['vol']['lse']['value'] == [5.5]


def test_backfill_new_exchanges_is_idempotent():
    old_keys = [k for k in u.EXCHANGES if k != 'tse']
    dash = _old_dash(old_keys)
    u._backfill_new_exchanges(dash)
    dash['vol']['tse']['dates'] = ['2026-07-14']   # simulate a real run having populated it
    u._backfill_new_exchanges(dash)
    assert dash['vol']['tse']['dates'] == ['2026-07-14'], 'second backfill call clobbered live data'


# ── _needs_universe_refresh: a new exchange forces refresh early ─────────────

def test_needs_universe_refresh_forces_refresh_for_new_exchange():
    # Regression: TSE's universe was silently never built because the other
    # 6 exchanges' 30-day timer hadn't expired yet — a global timer skipped
    # the refresh entirely, leaving TSE stuck on its hardcoded fallback list.
    old_keys = [k for k in u.EXCHANGES if k != 'tse']
    dash = {'universes': {k: {'tickers': ['X.L']} for k in old_keys}}
    dash['universes']['tse'] = {}
    needs, missing = u._needs_universe_refresh(dash, date.today().isoformat())
    assert needs is True
    assert missing == ['tse']


def test_needs_universe_refresh_skips_when_all_present_and_recent():
    dash = {'universes': {k: {'tickers': ['X.L']} for k in u.EXCHANGES}}
    needs, missing = u._needs_universe_refresh(dash, date.today().isoformat())
    assert needs is False
    assert missing == []


def test_needs_universe_refresh_triggers_on_monthly_expiry():
    dash = {'universes': {k: {'tickers': ['X.L']} for k in u.EXCHANGES}}
    stale_date = (date.today() - timedelta(days=31)).isoformat()
    needs, missing = u._needs_universe_refresh(dash, stale_date)
    assert needs is True
    assert missing == []


def test_needs_universe_refresh_triggers_on_first_ever_run():
    dash = {'universes': {k: {'tickers': ['X.L']} for k in u.EXCHANGES}}
    needs, _ = u._needs_universe_refresh(dash, '')
    assert needs is True


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
