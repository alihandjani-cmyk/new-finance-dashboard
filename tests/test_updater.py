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


# ── _clean_ticker: variable-width numeric exchanges (KRX) ────────────────────

def test_clean_ticker_numeric_width_plain_6digit_code():
    assert u._clean_ticker('005930', allow_numeric=True, numeric_width=6) == '005930'
    assert u._clean_ticker('000660', allow_numeric=True, numeric_width=6) == '000660'


def test_clean_ticker_numeric_width_mid_string_letter():
    # Regression: KRX has at least one oddball code with a letter in the
    # middle rather than trailing (Samsung Epis: '0126Z0'), unlike TSE where
    # the letter is always the last character ('543A'). The extraction must
    # not assume the letter is trailing.
    assert u._clean_ticker('0126Z0', allow_numeric=True, numeric_width=6) == '0126Z0'


def test_clean_ticker_numeric_width_default_unchanged_for_tse():
    # numeric_width defaults to 4 — existing TSE behavior must be untouched.
    assert u._clean_ticker('7203', allow_numeric=True) == '7203'
    assert u._clean_ticker('543A', allow_numeric=True) == '543A'


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


# ── _clean_ticker: share-class dot notation and dual-class cells ─────────────

def test_clean_ticker_converts_share_class_dot_to_dash():
    # Regression: LSE/NYSE Wikipedia tables list share classes with a dot
    # ('BT.A', 'BRK.B', 'BF.B'), but Yahoo requires a dash ('BT-A', 'BRK-B',
    # 'BF-B'). Left unconverted, the presence of '.' also wrongly signals to
    # _parse_wiki_table that an exchange suffix is already present, so the
    # raw broken ticker gets sent to yfinance as-is (observed live: '$BT.A:
    # possibly delisted').
    assert u._clean_ticker('BT.A') == 'BT-A'
    assert u._clean_ticker('BRK.B') == 'BRK-B'
    assert u._clean_ticker('BF.B') == 'BF-B'


def test_clean_ticker_splits_dual_class_cell():
    # Regression: some Wikipedia cells list two share classes together
    # ('SCHN / SCHP' for Schindler, 'LISN / LISP' for Lindt & Sprüngli) —
    # take the first/primary one instead of passing the garbled string
    # through (observed live: '$SCHN / SCHP.SW: possibly delisted').
    assert u._clean_ticker('SCHN / SCHP') == 'SCHN'
    assert u._clean_ticker('LISN / LISP') == 'LISN'


def test_find_col_case_insensitive():
    import pandas as pd
    df = pd.DataFrame({'Ticker': ['A'], 'Company': ['Foo']})
    assert u._find_col(df, u._TICKER_ALIASES) == 'Ticker'
    assert u._find_col(df, u._NAME_ALIASES)   == 'Company'


def test_find_col_returns_none_when_missing():
    import pandas as pd
    df = pd.DataFrame({'Price': [1.0], 'Volume': [1000]})
    assert u._find_col(df, u._TICKER_ALIASES) is None


def test_find_col_strips_footnote_marker():
    # Wikipedia headers sometimes carry a footnote ref, e.g. 'Ticker[a]' —
    # this alone was enough to make a real, well-formed table invisible to
    # _find_col (root cause of Nasdaq-100/Nikkei 225 universe builds failing).
    import pandas as pd
    df = pd.DataFrame({'Ticker[a]': ['AAPL'], 'Company[b]': ['Apple Inc.']})
    assert u._find_col(df, u._TICKER_ALIASES) == 'Ticker[a]'
    assert u._find_col(df, u._NAME_ALIASES)   == 'Company[b]'


def test_find_col_matches_multiindex_column():
    # pandas returns tuple column names for tables with a rowspan/merged
    # header — a flat string comparison against str(column) never matches.
    import pandas as pd
    df = pd.DataFrame({('Ticker', 'Unnamed: 0_level_1'): ['AAPL'],
                        ('Company', 'Unnamed: 1_level_1'): ['Apple Inc.']})
    assert u._find_col(df, u._TICKER_ALIASES) == ('Ticker', 'Unnamed: 0_level_1')
    assert u._find_col(df, u._NAME_ALIASES)   == ('Company', 'Unnamed: 1_level_1')


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


# ── _chunked_download: rate-limit mitigation for large ticker universes ──────

def _fake_ohlcv(tickers):
    """Build a minimal MultiIndex-columned frame shaped like yf.download()'s
    real output (field, then ticker) for the given tickers."""
    import pandas as pd
    idx = pd.date_range('2026-07-01', periods=3, freq='D')
    cols = pd.MultiIndex.from_product([['Close', 'Volume'], tickers])
    data = {}
    for t in tickers:
        data[('Close', t)]  = [10.0, 11.0, 12.0]
        data[('Volume', t)] = [100, 200, 300]
    return pd.DataFrame(data, index=idx, columns=cols)


def test_chunked_download_single_chunk_passthrough(monkeypatch):
    # A ticker list at or under chunk_size should go through as one call,
    # no sleeping.
    calls = []
    def fake_download(tickers, progress=False, threads=True, **kwargs):
        calls.append(list(tickers))
        return _fake_ohlcv(tickers)
    monkeypatch.setattr(u.yf, 'download', fake_download)
    monkeypatch.setattr(u.time, 'sleep', lambda s: (_ for _ in ()).throw(AssertionError('should not sleep')))

    tickers = [f'T{i}' for i in range(10)]
    raw = u._chunked_download(tickers, chunk_size=75, period='5d', interval='1d')
    assert len(calls) == 1
    assert raw is not None
    assert set(raw['Close'].columns) == set(tickers)


def test_chunked_download_splits_and_recombines(monkeypatch):
    # A ticker list over chunk_size should split into multiple yf.download()
    # calls, pause between them, and the combined frame should carry every
    # ticker from every chunk.
    calls, sleeps = [], []
    def fake_download(tickers, progress=False, threads=True, **kwargs):
        calls.append(list(tickers))
        return _fake_ohlcv(tickers)
    monkeypatch.setattr(u.yf, 'download', fake_download)
    monkeypatch.setattr(u.time, 'sleep', lambda s: sleeps.append(s))

    tickers = [f'T{i}' for i in range(10)]
    raw = u._chunked_download(tickers, chunk_size=4, pause=0.1, period='5d', interval='1d')
    assert len(calls) == 3          # 4 + 4 + 2
    assert len(sleeps) == 2         # pause between chunks, not after the last
    assert raw is not None
    assert set(raw['Close'].columns) == set(tickers)


def test_chunked_download_skips_failed_chunk(monkeypatch):
    # One chunk raising shouldn't take down the tickers in other chunks.
    def fake_download(tickers, progress=False, threads=True, **kwargs):
        if tickers[0] == 'T4':
            raise RuntimeError('rate limited')
        return _fake_ohlcv(tickers)
    monkeypatch.setattr(u.yf, 'download', fake_download)
    monkeypatch.setattr(u.time, 'sleep', lambda s: None)

    tickers = [f'T{i}' for i in range(10)]
    raw = u._chunked_download(tickers, chunk_size=4, pause=0.1, period='5d', interval='1d')
    assert raw is not None
    got = set(raw['Close'].columns)
    assert got == {'T0', 'T1', 'T2', 'T3', 'T8', 'T9'}   # T4-T7 chunk dropped


def test_chunked_download_returns_none_if_all_chunks_fail(monkeypatch):
    def fake_download(tickers, progress=False, threads=True, **kwargs):
        raise RuntimeError('rate limited')
    monkeypatch.setattr(u.yf, 'download', fake_download)
    monkeypatch.setattr(u.time, 'sleep', lambda s: None)

    tickers = [f'T{i}' for i in range(10)]
    raw = u._chunked_download(tickers, chunk_size=4, pause=0.1, period='5d', interval='1d')
    assert raw is None


def test_chunked_download_empty_tickers():
    assert u._chunked_download([], period='5d') is None


# ── _parse_wiki_table: html5lib fallback when lxml misses every table ────────

def test_parse_wiki_table_falls_back_to_html5lib(monkeypatch):
    # Confirmed live on Nasdaq-100: lxml (pandas' default parser) can find
    # zero usable tables on a page that genuinely has a clean Ticker/Company
    # table — html5lib parses some malformed/nested markup lxml chokes on.
    # This checks the fallback actually kicks in and returns the table
    # html5lib finds when the first (lxml) pass comes back empty.
    import pandas as pd

    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: b'<html>fake</html>')

    good_table = pd.DataFrame({
        'Ticker':  ['AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL'],
        'Company': ['Apple', 'Microsoft', 'Nvidia', 'Amazon', 'Alphabet'],
    })
    unrelated_table = pd.DataFrame({0: range(10), 1: range(10)})  # e.g. an infobox

    def fake_read_html(html, flavor=None):
        if flavor == 'html5lib':
            return [good_table]
        return [unrelated_table]   # lxml (default) pass: nothing usable

    monkeypatch.setattr(u.pd, 'read_html', fake_read_html)

    result = u._parse_wiki_table('https://en.wikipedia.org/wiki/Nasdaq-100', None)
    assert result == {
        'AAPL': 'Apple', 'MSFT': 'Microsoft', 'NVDA': 'Nvidia',
        'AMZN': 'Amazon', 'GOOGL': 'Alphabet',
    }


def test_parse_wiki_table_returns_empty_when_both_parsers_miss(monkeypatch):
    import pandas as pd
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: b'<html>fake</html>')
    unrelated_table = pd.DataFrame({0: range(10), 1: range(10)})
    monkeypatch.setattr(u.pd, 'read_html', lambda html, flavor=None: [unrelated_table])

    result = u._parse_wiki_table('https://en.wikipedia.org/wiki/Nikkei_225', None)
    assert result == {}


def test_parse_wiki_table_skips_html5lib_retry_when_lxml_succeeds(monkeypatch):
    # If lxml already found a usable table, the html5lib fallback should
    # never even be attempted — no wasted work on the exchanges that
    # already work fine.
    import pandas as pd
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: b'<html>fake</html>')

    good_table = pd.DataFrame({
        'Ticker':  ['AAPL', 'MSFT', 'NVDA', 'AMZN', 'GOOGL'],
        'Company': ['Apple', 'Microsoft', 'Nvidia', 'Amazon', 'Alphabet'],
    })
    calls = []
    def fake_read_html(html, flavor=None):
        calls.append(flavor)
        return [good_table]
    monkeypatch.setattr(u.pd, 'read_html', fake_read_html)

    result = u._parse_wiki_table('https://en.wikipedia.org/wiki/FTSE_100', None)
    assert len(result) == 5
    assert calls == [None]   # only the default (lxml) call was made


_NIKKEI_FAKE_HTML = b"""
<html><body>
<p>As of April 2026, the Nikkei 225 consists of the following companies
(Japanese securities identification code in parentheses):</p>
<ul><li>
<a href="/wiki/All_Nippon_Airways" title="All Nippon Airways">ANA Holdings</a>Inc.
(<a href="/wiki/Tokyo_Stock_Exchange" title="Tokyo Stock Exchange">TYO</a>:
<a rel="nofollow" class="external text"
   href="https://www2.jpx.co.jp/tseHpFront/StockSearch.do?callJorEFlg=1&amp;method=topsearch&amp;topSearchStr=9202">9202</a>)
<a href="/wiki/Archion" title="Archion">Archion</a>Corp.
(<a href="/wiki/Tokyo_Stock_Exchange" title="Tokyo Stock Exchange">TYO</a>:
<a rel="nofollow" class="external text"
   href="https://www2.jpx.co.jp/tseHpFront/StockSearch.do?callJorEFlg=1&amp;method=topsearch&amp;topSearchStr=543A">543A</a>)
</li></ul>
<h2>See also</h2>
<p>Unrelated section mentioning topSearchStr=9999">9999</a> that must not be picked up.</p>
</body></html>
"""


def test_parse_nikkei225_prose_extracts_tickers_and_names(monkeypatch):
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: _NIKKEI_FAKE_HTML)
    result = u._parse_nikkei225_prose('https://en.wikipedia.org/wiki/Nikkei_225')
    assert result == {
        '9202.T': 'ANA Holdings',
        '543A.T': 'Archion',   # alphanumeric JPX code (post-2024 format)
    }


def test_parse_nikkei225_prose_stops_before_next_h2(monkeypatch):
    # The bogus 9999 code lives after the <h2>See also</h2> marker and must
    # not be picked up as a 226th constituent.
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: _NIKKEI_FAKE_HTML)
    result = u._parse_nikkei225_prose('https://en.wikipedia.org/wiki/Nikkei_225')
    assert '9999.T' not in result


def test_parse_nikkei225_prose_missing_marker_returns_empty(monkeypatch):
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: b'<html>no marker here</html>')
    result = u._parse_nikkei225_prose('https://en.wikipedia.org/wiki/Nikkei_225')
    assert result == {}


def test_parse_nikkei225_prose_no_data_returns_empty(monkeypatch):
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: None)
    result = u._parse_nikkei225_prose('https://en.wikipedia.org/wiki/Nikkei_225')
    assert result == {}


_NIKKEI_TITLE_FIRST_HTML = b"""
<html><body>
<p>the following companies (Japanese securities identification code in parentheses):</p>
<ul><li>
<a title="Mitsubishi UFJ Financial Group" href="/wiki/Mitsubishi_UFJ_Financial_Group">Mitsubishi UFJ Financial Group</a>, Inc.
(<a title="Tokyo Stock Exchange" href="/wiki/Tokyo_Stock_Exchange">TYO</a>:
<a rel="nofollow" class="external text"
   href="https://www2.jpx.co.jp/tseHpFront/StockSearch.do?callJorEFlg=1&amp;method=topsearch&amp;topSearchStr=8306">8306</a>)
</li></ul>
</body></html>
"""


def test_parse_nikkei225_prose_handles_title_before_href(monkeypatch):
    # Confirmed live: an earlier version of this parser assumed href always
    # precedes title within an <a> tag's attributes, which isn't guaranteed
    # — Mitsubishi UFJ Financial Group's link came back title-first and
    # silently fell back to its bare code '8306' as the display name.
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: _NIKKEI_TITLE_FIRST_HTML)
    result = u._parse_nikkei225_prose('https://en.wikipedia.org/wiki/Nikkei_225')
    assert result == {'8306.T': 'Mitsubishi UFJ Financial Group'}


_NIKKEI_NO_LINK_HTML = b"""
<html><body>
<p>the following companies (Japanese securities identification code in parentheses):</p>
<ul><li>
Sompo Holdings Inc.
(<a href="/wiki/Tokyo_Stock_Exchange" title="Tokyo Stock Exchange">TYO</a>:
<a rel="nofollow" class="external text"
   href="https://www2.jpx.co.jp/tseHpFront/StockSearch.do?callJorEFlg=1&amp;method=topsearch&amp;topSearchStr=8630">8630</a>)
</li></ul>
</body></html>
"""


def test_parse_nikkei225_prose_falls_back_to_plain_text_name(monkeypatch):
    # A company with no wiki-article link at all (plain text before the
    # opening paren) should still recover a readable name rather than
    # falling all the way back to the bare ticker code.
    monkeypatch.setattr(u, '_get', lambda url, xlsx=False: _NIKKEI_NO_LINK_HTML)
    result = u._parse_nikkei225_prose('https://en.wikipedia.org/wiki/Nikkei_225')
    assert result == {'8630.T': 'Sompo Holdings Inc.'}


def test_fetch_constituents_uses_prose_parser_for_tse(monkeypatch):
    # _fetch_constituents must special-case 'tse' to call the prose parser
    # instead of the generic pd.read_html-based _parse_wiki_table path.
    calls = []
    monkeypatch.setattr(u, '_parse_nikkei225_prose', lambda url: (calls.append(url), {'9202.T': 'ANA Holdings'})[1])

    def fail_if_called(*args, **kwargs):
        raise AssertionError('_parse_wiki_table should not be called for tse')
    monkeypatch.setattr(u, '_parse_wiki_table', fail_if_called)

    result = u._fetch_constituents('tse')
    assert result == {'9202.T': 'ANA Holdings'}
    assert calls == ['https://en.wikipedia.org/wiki/Nikkei_225']
