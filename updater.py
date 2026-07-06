#!/usr/bin/env python3
"""
Finance Dashboard Updater
=========================
Fetches data for all 6 exchanges and writes dashboard_data.json.
The HTML reads this file on load — no HTML patching required.

Usage:
  python updater.py          # Full run — news, vol, ILLIQ, spreads, rankings, commentary
  python updater.py --fast   # Fast run — ticker bar, top-10 gainers, market cap only

Scheduled via two GitHub Actions workflows:
  update.yml       — Full run 4× daily (07:30, 10:30, 16:30, 22:00 UTC Mon-Fri)
  update-fast.yml  — Fast run every 30 min during market hours (08:00–21:30 UTC Mon-Fri)

Requirements (requirements.txt):
    openpyxl>=3.1  yfinance>=0.2  pandas>=2.0  anthropic>=0.25
    requests>=2.31  certifi>=2024.2
"""

import argparse, io, json, math, os, ssl, sys, urllib.request, urllib.error, urllib.parse
from datetime import datetime, date, timedelta
from pathlib import Path

# ─── Fail fast on missing dependencies ───────────────────────────────────────
try:
    import certifi
except ImportError:
    sys.exit('ERROR: certifi not installed. Run: pip install -r requirements.txt')
try:
    import yfinance as yf
except ImportError:
    sys.exit('ERROR: yfinance not installed. Run: pip install -r requirements.txt')
try:
    import pandas as pd
except ImportError:
    sys.exit('ERROR: pandas not installed. Run: pip install -r requirements.txt')
try:
    import openpyxl
except ImportError:
    sys.exit('ERROR: openpyxl not installed. Run: pip install -r requirements.txt')
try:
    import anthropic as _ant
except ImportError:
    _ant = None  # AI commentary disabled if not installed

# ─── SSL / HTTP ───────────────────────────────────────────────────────────────
_SSL = ssl.create_default_context(cafile=certifi.where())
_HDR = {'User-Agent': 'Mozilla/5.0 (compatible; FinancePulse/1.0)'}

def _get(url, xlsx=True):
    """HTTP GET → bytes, or None on failure."""
    try:
        req = urllib.request.Request(url, headers=_HDR)
        with urllib.request.urlopen(req, context=_SSL, timeout=25) as r:
            if r.status == 200:
                data = r.read()
                if xlsx and data[:2] != b'PK':
                    print(f'  ⚠  non-xlsx response from {url.split("/")[-1]}')
                    return None
                print(f'  ✓  {url.split("?")[0].split("/")[-1]} ({len(data)//1024} KB)')
                return data
    except Exception as exc:
        print(f'  ✗  {url.split("?")[0].split("/")[-1]}: {type(exc).__name__}: {exc}')
    return None

# ─── Paths ────────────────────────────────────────────────────────────────────
_ROOT      = Path(__file__).parent
DATA_FILE  = _ROOT / 'dashboard_data.json'
STATE_FILE = _ROOT / 'updater_state.json'  # gitignored — LSE file number etc.
DAYS       = 90    # rolling history window (business days)

# ══════════════════════════════════════════════════════════════════════════════
#  EXCHANGE UNIVERSE — tickers and display names
# ══════════════════════════════════════════════════════════════════════════════

FTSE_TICKERS = [
    'SHEL.L','AZN.L','HSBA.L','ULVR.L','BP.L','RIO.L','GSK.L','DGE.L','BATS.L','LSEG.L',
    'VOD.L','LLOY.L','NWG.L','PRU.L','NG.L','REL.L','EXPN.L','WPP.L','IMB.L','STAN.L',
    'BARC.L','AAL.L','CNA.L','IHG.L','JD.L','MKS.L','TSCO.L','SBRY.L','RKT.L','HLN.L',
    'ABF.L','ANTO.L','AUTO.L','RR.L','BA.L','CPG.L','BT-A.L','SSE.L','LAND.L','SGRO.L',
    'FLTR.L','MNDI.L','FRES.L','OCDO.L','PSON.L','WEIR.L','IMI.L','GLEN.L',
    'BHP.L','BNZL.L','III.L','SMWH.L','KGF.L','INF.L','ADM.L','LGEN.L','BTRW.L','TW.L',
]
LSE_NAMES = {
    'AZN.L':'AstraZeneca','SHEL.L':'Shell','HSBA.L':'HSBC Holdings','ULVR.L':'Unilever',
    'BP.L':'BP','RIO.L':'Rio Tinto','GSK.L':'GSK','DGE.L':'Diageo','BATS.L':'British American Tobacco',
    'LSEG.L':'London Stock Exchange Group','VOD.L':'Vodafone','LLOY.L':'Lloyds Banking Group',
    'NWG.L':'NatWest Group','PRU.L':'Prudential','NG.L':'National Grid','REL.L':'RELX',
    'EXPN.L':'Experian','WPP.L':'WPP','IMB.L':'Imperial Brands','STAN.L':'Standard Chartered',
    'BARC.L':'Barclays','AAL.L':'Anglo American','CNA.L':'Centrica','IHG.L':'IHG Hotels & Resorts',
    'JD.L':'JD Sports Fashion','MKS.L':'Marks & Spencer','TSCO.L':'Tesco','SBRY.L':"Sainsbury's",
    'RKT.L':'Reckitt','HLN.L':'Haleon','ABF.L':'Associated British Foods','ANTO.L':'Antofagasta',
    'AUTO.L':'Auto Trader','RR.L':'Rolls-Royce Holdings','BA.L':'BAE Systems','CPG.L':'Compass Group',
    'BT-A.L':'BT Group','SSE.L':'SSE','LAND.L':'Land Securities','SGRO.L':'Segro',
    'FLTR.L':'Flutter Entertainment','MNDI.L':'Mondi','FRES.L':'Fresnillo',
    'OCDO.L':'Ocado Group','PSON.L':'Pearson','WEIR.L':'Weir Group','IMI.L':'IMI',
    'GLEN.L':'Glencore','BHP.L':'BHP Group','BNZL.L':'Bunzl','III.L':'3i Group',
    'SMWH.L':'WH Smith','KGF.L':'Kingfisher','INF.L':'Informa','ADM.L':'Admiral Group',
    'LGEN.L':'Legal & General','BTRW.L':'Barratt Redrow','TW.L':'Taylor Wimpey',
    'MNDI.L':'Mondi',
}

ENX_TICKERS = [
    'MC.PA','OR.PA','RMS.PA','SAN.PA','SU.PA','TTE.PA','AIR.PA','AI.PA','SAF.PA','BNP.PA',
    'KER.PA','DSY.PA','DG.PA','AXA.PA','ENGI.PA','GLE.PA','ORA.PA','SGO.PA','RI.PA','CAP.PA',
    'STM.PA','HO.PA','VIE.PA','RNO.PA','ML.PA','EL.PA','PUB.PA','SW.PA',
    'ASML.AS','INGA.AS','PHIA.AS','ADYEN.AS','UNA.AS','WKL.AS','RAND.AS','ABN.AS','NN.AS','HEIA.AS',
    'ENEL.MI','ENI.MI','ISP.MI','UCG.MI','G.MI','LDO.MI','PRY.MI',
    'UCB.BR','KBC.BR','ABI.BR',
    'EQNR.OL','DNB.OL',
]
ENX_NAMES = {
    'MC.PA':'LVMH','OR.PA':"L'Oréal",'RMS.PA':'Hermès','SAN.PA':'Sanofi',
    'SU.PA':'Schneider Electric','TTE.PA':'TotalEnergies','AIR.PA':'Airbus',
    'AI.PA':'Air Liquide','SAF.PA':'Safran','BNP.PA':'BNP Paribas','KER.PA':'Kering',
    'DSY.PA':'Dassault Systèmes','DG.PA':'Vinci','AXA.PA':'AXA','ENGI.PA':'Engie',
    'GLE.PA':'Société Générale','ORA.PA':'Orange','SGO.PA':'Saint-Gobain',
    'RI.PA':'Pernod Ricard','CAP.PA':'Capgemini','STM.PA':'STMicroelectronics',
    'HO.PA':'Thales','VIE.PA':'Veolia','RNO.PA':'Renault','ML.PA':'Michelin',
    'EL.PA':'EssilorLuxottica','PUB.PA':'Publicis','SW.PA':'Sodexo',
    'ASML.AS':'ASML Holding','INGA.AS':'ING Group','PHIA.AS':'Philips','ADYEN.AS':'Adyen',
    'UNA.AS':'Unilever NV','WKL.AS':'Wolters Kluwer','RAND.AS':'Randstad',
    'ABN.AS':'ABN AMRO','NN.AS':'NN Group','HEIA.AS':'Heineken',
    'ENEL.MI':'Enel','ENI.MI':'ENI','ISP.MI':'Intesa Sanpaolo','UCG.MI':'UniCredit',
    'G.MI':'Generali','LDO.MI':'Leonardo','PRY.MI':'Prysmian',
    'UCB.BR':'UCB','KBC.BR':'KBC Group','ABI.BR':'AB InBev',
    'EQNR.OL':'Equinor','DNB.OL':'DNB Bank',
}

NDX_TICKERS = [
    'AAPL','MSFT','NVDA','AMZN','META','GOOGL','TSLA','AVGO','COST',
    'NFLX','AMD','TMUS','QCOM','INTU','AMAT','BKNG','ISRG','AMGN','CMCSA',
    'TXN','HON','PEP','VRTX','SBUX','GILD','MDLZ','ADI','REGN','KDP',
    'PANW','KLAC','MELI','LRCX','SNPS','CDNS','CTAS','MU','MAR','ORLY',
    'CRWD','CSX','MRVL','PCAR','WDAY','AEP','FTNT','ABNB','MNST','PYPL',
    'CHTR','ODFL','DXCM','EXC','FAST','CEG','ROST','VRSK','CPRT','IDXX',
    'PAYX','TTD','KHC','BIIB','XEL','ZS','ANSS','ON','TEAM','DDOG',
    'WBD','DLTR','CDW','ILMN','MCHP','NXPI','ADSK','GEHC','INTC','CSCO',
    'ADBE','APP','SMCI','ARM','LIN','ANET','FANG','EBAY','PLTR','ASML',
]
NDX_VOL_TICKERS = {
    'AAPL','MSFT','NVDA','AMZN','GOOGL','META','TSLA','AVGO','COST',
    'NFLX','ASML','TMUS','PLTR','AMD','APP','INTU','QCOM','AMAT','BKNG',
    'ISRG','AMGN','TXN','VRTX','CMCSA','ADI','REGN','HON','PEP','PANW',
    'KLAC','MELI','LRCX','SNPS','CDNS','CTAS','MU','ORLY','CRWD','MRVL',
    'CSCO','ADBE','ARM','ANET','CEG','FTNT','ADSK','WDAY','INTC','SBUX',
}
NDX_NAMES = {
    'AAPL':'Apple','MSFT':'Microsoft','NVDA':'NVIDIA','AMZN':'Amazon',
    'META':'Meta Platforms','GOOGL':'Alphabet','TSLA':'Tesla','AVGO':'Broadcom',
    'COST':'Costco','NFLX':'Netflix','AMD':'AMD','TMUS':'T-Mobile US',
    'QCOM':'Qualcomm','INTU':'Intuit','AMAT':'Applied Materials','BKNG':'Booking Holdings',
    'ISRG':'Intuitive Surgical','AMGN':'Amgen','CMCSA':'Comcast','TXN':'Texas Instruments',
    'HON':'Honeywell','PEP':'PepsiCo','VRTX':'Vertex Pharma','SBUX':'Starbucks',
    'GILD':'Gilead Sciences','MDLZ':'Mondelēz Intl','ADI':'Analog Devices',
    'REGN':'Regeneron','KDP':'Keurig Dr Pepper','PANW':'Palo Alto Networks',
    'KLAC':'KLA Corp','MELI':'MercadoLibre','LRCX':'Lam Research','SNPS':'Synopsys',
    'CDNS':'Cadence Design','CTAS':'Cintas','MU':'Micron Technology','MAR':'Marriott',
    'ORLY':"O'Reilly Auto",'CRWD':'CrowdStrike','CSX':'CSX Corp',
    'MRVL':'Marvell Technology','PCAR':'PACCAR','WDAY':'Workday',
    'AEP':'American Electric','FTNT':'Fortinet','ABNB':'Airbnb',
    'MNST':'Monster Beverage','PYPL':'PayPal','CHTR':'Charter Comms',
    'ODFL':'Old Dominion Freight','DXCM':'DexCom','EXC':'Exelon','FAST':'Fastenal',
    'CEG':'Constellation Energy','ROST':'Ross Stores','VRSK':'Verisk Analytics',
    'CPRT':'Copart','IDXX':'IDEXX Labs','PAYX':'Paychex','TTD':'The Trade Desk',
    'KHC':'Kraft Heinz','BIIB':'Biogen','XEL':'Xcel Energy','ZS':'Zscaler',
    'ANSS':'ANSYS','ON':'ON Semiconductor','TEAM':'Atlassian','DDOG':'Datadog',
    'WBD':'Warner Bros. Discovery','DLTR':'Dollar Tree','CDW':'CDW Corp',
    'ILMN':'Illumina','MCHP':'Microchip Tech','NXPI':'NXP Semiconductors',
    'ADSK':'Autodesk','GEHC':'GE HealthCare','INTC':'Intel','CSCO':'Cisco',
    'ADBE':'Adobe','APP':'AppLovin','SMCI':'Super Micro Computer','ARM':'Arm Holdings',
    'LIN':'Linde','ANET':'Arista Networks','FANG':'Diamondback Energy',
    'EBAY':'eBay','PLTR':'Palantir','ASML':'ASML Holding',
}

NYSE_TICKERS = [
    'BRK-B','JPM','V','JNJ','WMT','PG','XOM','MA','CVX','BAC',
    'LLY','UNH','KO','MRK','ABBV','DIS','GS','MS','CAT','IBM',
    'AXP','DE','C','WFC','RTX','LMT','PFE','BA','T','VZ',
    'COP','GE','NKE','MCD','ACN','PM','CRM','MMM','BMY','NEE',
    'AMT','TGT','HD','LOW','SYK','CI','CB','PLD','SO','DUK',
]
NYSE_NAMES = {
    'BRK-B':'Berkshire Hathaway','JPM':'JPMorgan Chase','V':'Visa',
    'JNJ':'Johnson & Johnson','WMT':'Walmart','PG':'Procter & Gamble',
    'XOM':'ExxonMobil','MA':'Mastercard','CVX':'Chevron','BAC':'Bank of America',
    'LLY':'Eli Lilly','UNH':'UnitedHealth','KO':'Coca-Cola','MRK':'Merck',
    'ABBV':'AbbVie','DIS':'Walt Disney','GS':'Goldman Sachs','MS':'Morgan Stanley',
    'CAT':'Caterpillar','IBM':'IBM','AXP':'American Express','DE':'Deere & Co',
    'C':'Citigroup','WFC':'Wells Fargo','RTX':'RTX Corp','LMT':'Lockheed Martin',
    'PFE':'Pfizer','BA':'Boeing','T':'AT&T','VZ':'Verizon',
    'COP':'ConocoPhillips','GE':'GE Aerospace','NKE':'Nike','MCD':"McDonald's",
    'ACN':'Accenture','PM':'Philip Morris Intl','CRM':'Salesforce','MMM':'3M',
    'BMY':'Bristol-Myers Squibb','NEE':'NextEra Energy','AMT':'American Tower',
    'TGT':'Target','HD':'Home Depot','LOW':"Lowe's",'SYK':'Stryker',
    'CI':'Cigna','CB':'Chubb','PLD':'Prologis','SO':'Southern Company','DUK':'Duke Energy',
}

XETRA_TICKERS = [
    'SAP.DE','SIE.DE','ALV.DE','DTE.DE','BMW.DE','MBG.DE','BAS.DE','MUV2.DE',
    'ADS.DE','RWE.DE','DBK.DE','BAYN.DE','IFX.DE','EOAN.DE','VOW3.DE','MTX.DE',
    'DTG.DE','VNA.DE','DB1.DE','RHM.DE','P911.DE','BEI.DE','HEI.DE','BNR.DE',
    'ENR.DE','DHL.DE','CBK.DE','PAH3.DE','MRK.DE','HNR1.DE','PUM.DE','SRT3.DE',
    'ZAL.DE','FME.DE','QIA.DE','SY1.DE','CON.DE','HLAG.DE',
]
XETRA_NAMES = {
    'SAP.DE':'SAP','SIE.DE':'Siemens','ALV.DE':'Allianz','DTE.DE':'Deutsche Telekom',
    'BMW.DE':'BMW','MBG.DE':'Mercedes-Benz','BAS.DE':'BASF','MUV2.DE':'Munich Re',
    'ADS.DE':'Adidas','RWE.DE':'RWE','DBK.DE':'Deutsche Bank','BAYN.DE':'Bayer',
    'IFX.DE':'Infineon','EOAN.DE':'E.ON','VOW3.DE':'Volkswagen','MTX.DE':'MTU Aero Engines',
    'DTG.DE':'Daimler Truck','VNA.DE':'Vonovia','DB1.DE':'Deutsche Boerse',
    'RHM.DE':'Rheinmetall','P911.DE':'Porsche AG','BEI.DE':'Beiersdorf',
    'HEI.DE':'Heidelberg Materials','BNR.DE':'Brenntag','ENR.DE':'Siemens Energy',
    'DHL.DE':'DHL Group','CBK.DE':'Commerzbank','PAH3.DE':'Porsche Holding',
    'MRK.DE':'Merck KGaA','HNR1.DE':'Hannover Re','PUM.DE':'Puma','SRT3.DE':'Sartorius',
    'ZAL.DE':'Zalando','FME.DE':'Fresenius Medical','QIA.DE':'Qiagen',
    'SY1.DE':'Symrise','CON.DE':'Continental','HLAG.DE':'Hapag-Lloyd',
}

SIX_TICKERS = [
    'NESN.SW','NOVN.SW','ROG.SW','UBSG.SW','ALC.SW','ABBN.SW','SREN.SW',
    'GEBN.SW','GIVN.SW','LOGN.SW','PGHN.SW','SCMN.SW','SIKA.SW','SLHN.SW',
    'SOON.SW','CFR.SW','ZURN.SW','HOLN.SW','STMN.SW','VACN.SW','KNIN.SW',
    'LISN.SW','BALN.SW',
]
SIX_NAMES = {
    'NESN.SW':'Nestlé','NOVN.SW':'Novartis','ROG.SW':'Roche','UBSG.SW':'UBS Group',
    'ALC.SW':'Alcon','ABBN.SW':'ABB','SREN.SW':'Swiss Re','GEBN.SW':'Geberit',
    'GIVN.SW':'Givaudan','LOGN.SW':'Lonza Group','PGHN.SW':'Partners Group',
    'SCMN.SW':'Swisscom','SIKA.SW':'Sika','SLHN.SW':'Swiss Life','SOON.SW':'Sonova',
    'CFR.SW':'Richemont','ZURN.SW':'Zurich Insurance','HOLN.SW':'Holcim',
    'STMN.SW':'Straumann','VACN.SW':'VAT Group','KNIN.SW':'Kuehne+Nagel',
    'LISN.SW':'Lindt & Spruengli','BALN.SW':'Baloise',
}

# ─── Exchange config ──────────────────────────────────────────────────────────
EXCHANGES = {
    'lse':   {'name':'London Stock Exchange','currency':'GBP','vol_currency':'GBP B',
               'vol_method':'file','pence':True, 'nok_eur':False,
               'tickers':FTSE_TICKERS,'names':LSE_NAMES},
    'enx':   {'name':'Euronext',            'currency':'EUR','vol_currency':'EUR B',
               'vol_method':'file','pence':False,'nok_eur':True,
               'tickers':ENX_TICKERS,'names':ENX_NAMES},
    'ndx':   {'name':'Nasdaq 100',          'currency':'USD','vol_currency':'USD B',
               'vol_method':'yf',  'pence':False,'nok_eur':False,
               'tickers':NDX_TICKERS,'names':NDX_NAMES,'vol_tickers':NDX_VOL_TICKERS},
    'nyse':  {'name':'NYSE',                'currency':'USD','vol_currency':'USD B',
               'vol_method':'yf',  'pence':False,'nok_eur':False,
               'tickers':NYSE_TICKERS,'names':NYSE_NAMES,'vol_tickers':set(NYSE_TICKERS)},
    'xetra': {'name':'Xetra',               'currency':'EUR','vol_currency':'EUR B',
               'vol_method':'yf',  'pence':False,'nok_eur':False,
               'tickers':XETRA_TICKERS,'names':XETRA_NAMES,'vol_tickers':set(XETRA_TICKERS)},
    'six':   {'name':'SIX Swiss Exchange',  'currency':'CHF','vol_currency':'CHF B',
               'vol_method':'yf',  'pence':False,'nok_eur':False,
               'tickers':SIX_TICKERS,'names':SIX_NAMES,'vol_tickers':set(SIX_TICKERS)},
}

# ── Ticker-bar symbols (live prices shown in the top bar) ────────────────────
TICKER_SYMBOLS = {
    'FTSE':  '^FTSE',    'DAX':    '^GDAXI',   'CAC40': '^FCHI',   'SMI': '^SSMI',
    'SP500': '^GSPC',    'NDX':    '^NDX',      'DJI':   '^DJI',
    'GBPUSD':'GBPUSD=X', 'EURUSD': 'EURUSD=X', 'EURGBP':'EURGBP=X','USDJPY':'JPY=X',
    'GOLD':  'GC=F',     'OIL':    'CL=F',
}

# ══════════════════════════════════════════════════════════════════════════════
#  DATA FILE HELPERS
# ══════════════════════════════════════════════════════════════════════════════

def _load_dashboard():
    """Load existing dashboard_data.json (preserves 90-day histories)."""
    try:
        if DATA_FILE.exists():
            return json.loads(DATA_FILE.read_text())
    except Exception as exc:
        print(f'  ⚠  Could not load existing data: {exc}')
    # Return the empty skeleton
    return {
        'generated_at': None, 'stale_fallback': True, 'commentary': None,
        'commentary_date': None,
        'news': [], 'tickers': {},
        'gainers':    {k: [] for k in EXCHANGES},
        'market_cap': {k: {'date': None, 'currency': v['currency'], 'top10': []} for k, v in EXCHANGES.items()},
        'vol':        {k: {'currency': v['vol_currency'], 'dates': [], 'value': []} for k, v in EXCHANGES.items()},
        'amihud':     {k: {'history': [], 'marketAvg': None} for k in EXCHANGES},
        'spread':     {k: {'history': [], 'marketAvg': None} for k in EXCHANGES},
        'ar_spread':  {k: {'history': [], 'marketAvg': None} for k in EXCHANGES},
        'current_ranking': {'date': None, 'ranks': {k: {'vol':None,'illiq':None,'spread':None,'ar':None,'composite':None} for k in EXCHANGES}},
        'ranking_history': [],
    }

def _save_dashboard(data):
    data['generated_at'] = datetime.utcnow().strftime('%Y-%m-%dT%H:%M:%SZ')
    data['stale_fallback'] = False
    DATA_FILE.write_text(json.dumps(data, indent=2, ensure_ascii=False))
    print(f'\n  ✓  Saved {DATA_FILE} ({DATA_FILE.stat().st_size // 1024} KB)')

# ─── Date / history utilities ─────────────────────────────────────────────────

def _date_key(label):
    """'02 Jul' → int for chronological sorting."""
    mon = {'Jan':1,'Feb':2,'Mar':3,'Apr':4,'May':5,'Jun':6,
           'Jul':7,'Aug':8,'Sep':9,'Oct':10,'Nov':11,'Dec':12}
    try:
        d, m = label.strip().split()
        return mon.get(m, 0) * 100 + int(d)
    except Exception:
        return 0

def _push(lst, entry, key='date', maxn=DAYS):
    """Append/update entry (matched by key), sort, trim to maxn."""
    out = [e for e in lst if e.get(key) != entry.get(key)]
    out.append(entry)
    out.sort(key=lambda e: _date_key(e.get(key, '')))
    return out[-maxn:]

def _today():
    return date.today().strftime('%d %b')

def _prev_bday(dt=None, n=1):
    d = dt or date.today()
    for _ in range(n):
        d -= timedelta(days=1)
        while d.weekday() >= 5:
            d -= timedelta(days=1)
    return d

# ══════════════════════════════════════════════════════════════════════════════
#  NEWS
# ══════════════════════════════════════════════════════════════════════════════

def fetch_news():
    """Fetch top 10 financial news headlines from Reuters RSS."""
    import xml.etree.ElementTree as ET
    feeds = [
        'https://feeds.reuters.com/reuters/businessNews',
        'https://feeds.reuters.com/news/wealth',
    ]
    items = []
    for url in feeds:
        raw = _get(url, xlsx=False)
        if not raw:
            continue
        try:
            root = ET.fromstring(raw)
            ns = {'atom': 'http://www.w3.org/2005/Atom'}
            for item in root.iter('item'):
                title = (item.findtext('title') or '').strip()
                link  = (item.findtext('link') or '').strip()
                pub   = (item.findtext('pubDate') or '').strip()
                if title and link and link.startswith('https://'):
                    items.append({'title': title, 'link': link, 'pubDate': pub, 'source': 'Reuters'})
        except Exception as exc:
            print(f'  ⚠  RSS parse error: {exc}')
    seen, out = set(), []
    for it in items:
        if it['title'] not in seen:
            seen.add(it['title'])
            out.append(it)
        if len(out) >= 10:
            break
    print(f'  {len(out)} news items')
    return out or None

# ══════════════════════════════════════════════════════════════════════════════
#  TICKER BAR
# ══════════════════════════════════════════════════════════════════════════════

def fetch_tickers():
    """Fetch latest price + % change for ticker bar symbols."""
    syms = list(TICKER_SYMBOLS.values())
    try:
        raw = yf.download(syms, period='2d', interval='1d', progress=False,
                          auto_adjust=True, threads=True)
    except Exception as exc:
        print(f'  ✗  Ticker download failed: {exc}')
        return None
    result = {}
    for key, sym in TICKER_SYMBOLS.items():
        try:
            closes = raw['Close'][sym].dropna() if sym in raw['Close'].columns else raw['Close'].dropna()
            if len(closes) < 1:
                continue
            price = float(closes.iloc[-1])
            prev  = float(closes.iloc[-2]) if len(closes) >= 2 else price
            pct   = (price - prev) / prev * 100 if prev else 0.0
            result[key] = {'price': round(price, 4), 'pct': round(pct, 4)}
        except Exception:
            continue
    print(f'  {len(result)} tickers')
    return result or None

# ══════════════════════════════════════════════════════════════════════════════
#  GAINERS  (generic — works for all 6 exchanges)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_gainers(ex_key):
    """Return top-10 % gainers for the given exchange."""
    cfg      = EXCHANGES[ex_key]
    tickers  = cfg['tickers']
    names    = cfg['names']
    is_pence = cfg['pence']
    currency = cfg['currency']
    try:
        raw = yf.download(tickers, period='2d', interval='1d', progress=False,
                          auto_adjust=True, threads=True)
        # Default (no group_by): MultiIndex is (Price, Ticker) so raw['Close'] works.
    except Exception as exc:
        print(f'  ✗  {ex_key} gainers download: {exc}')
        return None

    is_multi = isinstance(raw.columns, pd.MultiIndex)
    if is_multi:
        if 'Close' not in raw.columns.get_level_values(0):
            print(f'  ✗  {ex_key}: no Close column in yfinance result (all tickers failed?)')
            return None
        close_df = raw['Close']   # DataFrame: index=date, columns=tickers
    else:
        if 'Close' not in raw.columns:
            print(f'  ✗  {ex_key}: no Close column in yfinance result')
            return None
        close_df = raw[['Close']].rename(columns={'Close': tickers[0]})
    results = []
    for sym in tickers:
        try:
            if is_multi:
                if sym not in close_df.columns:
                    continue
                closes = close_df[sym].dropna()
            else:
                closes = close_df.dropna()
            if len(closes) < 2:
                continue
            curr = float(closes.iloc[-1])
            prev = float(closes.iloc[-2])
            if prev <= 0:
                continue
            pct = (curr - prev) / prev * 100
            chg = curr - prev
            results.append({'sym': sym, 'price': curr, 'chg': chg, 'pct': pct})
        except Exception:
            continue

    if not results:
        print(f'  ✗  {ex_key}: no gainers found')
        return None

    results.sort(key=lambda x: x['pct'], reverse=True)
    suffix_map = {'lse': '.L', 'enx': None, 'ndx': None, 'nyse': None,
                  'xetra': '.DE', 'six': '.SW'}
    suffix = suffix_map.get(ex_key)

    output = []
    for r in results[:10]:
        sym = r['sym']
        p, chg, pct = r['price'], r['chg'], r['pct']
        display = names.get(sym, sym.split('.')[0] if suffix else sym)
        ticker_clean = sym.replace(suffix, '') if suffix and sym.endswith(suffix) else sym.split('.')[0]

        if is_pence:
            price_str = f"{int(p)}p" if p >= 100 else f"{p:.2f}p"
            chg_str   = f"{'+' if chg>=0 else ''}{int(chg)}p" if abs(chg) >= 1 else f"{'+' if chg>=0 else ''}{chg:.2f}p"
        else:
            price_str = f"{currency}{p:.2f}"
            chg_str   = f"{'+' if chg>=0 else ''}{currency}{chg:.2f}"

        output.append({'name': display, 'ticker': ticker_clean,
                       'price': price_str, 'change': chg_str,
                       'pct': f"{pct:.2f}%"})

    print(f'  {len(output)} gainers · top: {output[0]["name"]} +{output[0]["pct"]}')
    return output

# ══════════════════════════════════════════════════════════════════════════════
#  MARKET CAP  (generic — works for all 6 exchanges)
# ══════════════════════════════════════════════════════════════════════════════

def fetch_market_cap(ex_key):
    """Return top-10 stocks by market cap for the given exchange."""
    cfg      = EXCHANGES[ex_key]
    tickers  = cfg['tickers']
    names    = cfg['names']
    is_pence = cfg['pence']
    currency = cfg['currency']
    mc_divisor = 1e11 if is_pence else 1e9   # pence→£B needs /1e11; EUR/USD/CHF→B needs /1e9

    try:
        tickers_obj = yf.Tickers(' '.join(tickers))
    except Exception as exc:
        print(f'  ✗  {ex_key} market cap: {exc}')
        return None

    results = []
    for sym in tickers:
        try:
            fi = tickers_obj.tickers[sym].fast_info
            mc    = getattr(fi, 'market_cap', None) or getattr(fi, 'marketCap', None)
            price = getattr(fi, 'last_price',  None) or getattr(fi, 'lastPrice',  0) or 0
            if mc and float(mc) > 0:
                results.append({'sym': sym, 'mc': float(mc), 'price': float(price)})
        except Exception:
            continue

    if not results:
        print(f'  ✗  {ex_key}: no market cap data')
        return None

    results.sort(key=lambda x: x['mc'], reverse=True)
    suffix_map = {'lse': '.L', 'enx': None, 'ndx': None, 'nyse': None,
                  'xetra': '.DE', 'six': '.SW'}
    suffix = suffix_map.get(ex_key)

    top10 = []
    for r in results[:10]:
        sym = r['sym']
        mc_b = r['mc'] / mc_divisor
        p    = r['price']
        display = names.get(sym, sym.split('.')[0] if suffix else sym)
        ticker_clean = sym.replace(suffix, '') if suffix and sym.endswith(suffix) else sym.split('.')[0]

        if is_pence:
            price_str = f'{int(p)}p' if p >= 100 else f'{p:.2f}p'
        else:
            price_str = f'{currency}{p:.2f}'

        top10.append({'name': display, 'ticker': ticker_clean,
                      'mcap_b': round(mc_b, 1), 'price': price_str})

    today_label = _today()
    print(f'  {len(top10)} stocks · largest: {top10[0]["name"]} {top10[0]["mcap_b"]:.1f}B')
    return {'date': today_label, 'currency': currency, 'top10': top10}

# ══════════════════════════════════════════════════════════════════════════════
#  VOLUME — file-based (LSE and Euronext)
# ══════════════════════════════════════════════════════════════════════════════

LSE_URL = ('https://docs.londonstockexchange.com/sites/default/files'
           '/reports/Order%20book%20trading_{num}.xlsx')

def _lse_fetch(state):
    last_num = state.get('lse_file_num', 1513)
    latest   = last_num
    for n in range(last_num + 1, last_num + 40):
        try:
            req = urllib.request.Request(LSE_URL.format(num=n), method='HEAD', headers=_HDR)
            with urllib.request.urlopen(req, context=_SSL, timeout=10) as r:
                if r.status == 200:
                    latest = n
                else:
                    break
        except Exception:
            break
    raw = _get(LSE_URL.format(num=latest))
    if raw:
        state['lse_file_num'] = latest
    return raw

def _lse_parse(raw):
    """Extract last 5 complete trading days from LSE xlsx."""
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb['Daily Order Book Trading']
    rows, header = [], False
    for row in ws.iter_rows(values_only=True):
        if row[1] == 'Trade Date':
            header = True
            continue
        if header and isinstance(row[1], datetime) and row[4] is not None:
            rows.append({
                'label':  row[1].strftime('%d %b'),
                'value':  round(float(row[5]) / 1e9, 2),        # £B turnover
                'trades': round(int(row[4]) / 1000, 1),         # thousands of trades
            })
        if len(rows) >= 5:
            break
    return list(reversed(rows[:5]))  # oldest → newest

def _enx_url(dt):
    d  = dt.strftime('%Y%m%d')
    yr = dt.strftime('%Y')
    base = 'https://live.euronext.com/sites/default/files/statistics/cash/nextday'
    return f'{base}/{yr}/Cash%20{d}.xlsx', f'{base}/2017/Cash%20{d}.xlsx'

def _enx_fetch(dt):
    url_yr, url_17 = _enx_url(dt)
    raw = _get(url_yr)
    if raw is None:
        raw = _get(url_17)
    return raw

def _enx_parse(raw):
    """Extract T, T-1, T-2 from Euronext xlsx. Returns list of up to 3 dicts."""
    wb = openpyxl.load_workbook(io.BytesIO(raw), data_only=True)
    ws = wb['All Markets']
    rows = list(ws.iter_rows(values_only=True))
    date_row     = rows[4]    # row 5
    turnover_row = rows[6]    # row 7 — TOTAL TURNOVER (mln €)
    trades_row   = rows[22] if len(rows) > 22 else None  # row 23 — TOTAL TRADES
    pts = []
    for col in [4, 3, 2]:    # E→D→C (oldest→newest)
        d = date_row[col]
        v = turnover_row[col]
        if isinstance(d, datetime) and v is not None:
            pt = {'label': d.strftime('%d %b'),
                  'value': round(float(v) / 1000, 2)}   # mln€ → B€
            if trades_row is not None:
                try:
                    t = trades_row[col]
                    if t is not None:
                        pt['trades'] = round(float(t) / 1_000_000, 2)  # → millions
                except Exception:
                    pass
            pts.append(pt)
    return pts

def fetch_vol_file(ex_key, state):
    """Download and parse LSE or Euronext official volume file.
    Returns list of {'label', 'value'} dicts (oldest→newest, up to 5 days).
    Returns None if download fails.
    """
    if ex_key == 'lse':
        raw = _lse_fetch(state)
        if raw is None:
            return None
        return _lse_parse(raw)

    # Euronext — build 5 days from 2 files
    today_dt, today_raw = date.today(), None
    for _ in range(4):
        today_raw = _enx_fetch(today_dt)
        if today_raw:
            break
        today_dt = _prev_bday(today_dt)
    if today_raw is None:
        print('  ✗  Euronext: could not fetch any file')
        return None

    today_pts = _enx_parse(today_raw)
    if len(today_pts) < 2:
        return today_pts or None

    # Try to fetch the file from T-2 to get older points
    from datetime import datetime as dt_cls
    try:
        t2_dt = dt_cls.strptime(today_pts[0]['label'] + f' {date.today().year}', '%d %b %Y').date()
    except Exception:
        return today_pts

    t2_raw = _enx_fetch(t2_dt)
    if t2_raw is None:
        return today_pts

    t2_pts = _enx_parse(t2_raw)
    combined = t2_pts[:2] + today_pts     # up to 5 days, oldest→newest
    return combined or None

# ──────────────────────────────────────────────────────────────────────────────
#  VOLUME — yfinance (NDX / NYSE / Xetra / SIX)
# ──────────────────────────────────────────────────────────────────────────────

def fetch_vol_yf(ex_key):
    """Compute aggregate daily turnover (£/€/$/CHF billions) from yfinance.
    Returns list of {'label', 'value'} dicts (oldest→newest, last 5 trading days).
    """
    cfg         = EXCHANGES[ex_key]
    vol_tickers = list(cfg.get('vol_tickers', cfg['tickers']))
    try:
        raw = yf.download(vol_tickers, period='15d', interval='1d',
                          progress=False, auto_adjust=True, threads=True)
        if raw.empty:
            print(f'  ✗  {ex_key} vol: no data')
            return None
    except Exception as exc:
        print(f'  ✗  {ex_key} vol download: {exc}')
        return None

    try:
        close_df  = raw['Close']  if isinstance(raw.columns, pd.MultiIndex) else raw
        volume_df = raw['Volume'] if isinstance(raw.columns, pd.MultiIndex) else None
        if volume_df is None:
            return None

        # Daily turnover = Σ(Close × Volume) / 1e9 for each trading day
        daily_turnover = (close_df * volume_df).sum(axis=1, min_count=1) / 1e9
        daily_turnover = daily_turnover.dropna()
        # Daily shares traded = Σ(Volume) / 1e6 (millions of shares)
        daily_shares = volume_df.sum(axis=1, min_count=1) / 1e6
        # Exclude today if market not closed yet
        today_str = _today()
        pts = []
        for ts, val in daily_turnover.items():
            label = ts.strftime('%d %b')
            if label == today_str:
                continue  # skip today's incomplete intraday
            if val > 0:
                sh = float(daily_shares.get(ts, 0)) if ts in daily_shares.index else 0.0
                pts.append({'label': label, 'value': round(float(val), 2),
                            'shares_m': round(sh, 1) if sh > 0 else None})
        pts = pts[-5:]  # last 5 complete days
        print(f'  {len(pts)} vol days · last: {pts[-1]["value"]:.1f}B ({ex_key})' if pts else f'  ✗  {ex_key}: no vol points')
        return pts or None
    except Exception as exc:
        print(f'  ✗  {ex_key} vol parse: {exc}')
        return None

# ══════════════════════════════════════════════════════════════════════════════
#  AMIHUD ILLIQUIDITY  (generic — all 6 exchanges)
#  ILLIQ = |R| / DVOL_M  (% price move per 1M currency unit traded)
# ══════════════════════════════════════════════════════════════════════════════

def _fetch_nok_eur():
    try:
        raw = yf.download('NOKEUR=X', period='2d', interval='1d',
                          progress=False, auto_adjust=True)
        return float(raw['Close'].dropna().iloc[-1])
    except Exception:
        return 0.086   # fallback ~0.086 EUR per NOK

def fetch_amihud(ex_key, existing_history):
    """Compute per-day Amihud ILLIQ for the exchange universe.
    Appends new days to existing_history (90-day rolling window).
    Returns (updated_history, marketAvg, top10_liquid, top10_illiquid).
    """
    cfg      = EXCHANGES[ex_key]
    tickers  = cfg['tickers']
    names    = cfg['names']
    is_pence = cfg['pence']
    nok_eur  = _fetch_nok_eur() if cfg['nok_eur'] else None
    period   = '40d'  # fetch 40 days to cover 20+ business days

    try:
        raw = yf.download(tickers, period=period, interval='1d', progress=False,
                          auto_adjust=True, threads=True)
        if raw.empty:
            print(f'  ✗  {ex_key} amihud: no data')
            return existing_history, None
    except Exception as exc:
        print(f'  ✗  {ex_key} amihud download: {exc}')
        return existing_history, None

    daily_buckets = {}   # date_label → [per-stock ILLIQ values]
    stock_avgs    = {}   # sym → 20-day avg ILLIQ (for current marketAvg)

    for sym in tickers:
        try:
            df = raw[[('Close', sym), ('Volume', sym)]].copy()
            df.columns = ['Close', 'Volume']
            df = df.dropna()
            if len(df) < 6:
                continue

            if is_pence:
                # Close in pence; dvol_m = Close × Volume / 100 / 1e6 = £M
                dvol_m = df['Close'] * df['Volume'] / 100 / 1_000_000
            else:
                # Close in EUR/USD/CHF; dvol_m = Close × Volume / 1e6 = currency M
                dvol_m = df['Close'] * df['Volume'] / 1_000_000
                if cfg['nok_eur'] and sym.endswith('.OL') and nok_eur:
                    dvol_m = dvol_m * nok_eur   # NOK M → EUR M

            ret_pct = df['Close'].pct_change().abs() * 100
            illiq   = (ret_pct / dvol_m).replace([float('inf'), float('-inf')], float('nan')).dropna()
            illiq   = illiq.iloc[-20:]
            if len(illiq) < 5:
                continue

            stock_avgs[sym] = float(illiq.mean())
            for ts, val in illiq.items():
                if pd.isna(val):
                    continue
                label = ts.strftime('%d %b')
                daily_buckets.setdefault(label, []).append(float(val))
        except Exception:
            continue

    if not stock_avgs:
        print(f'  ✗  {ex_key}: no ILLIQ computed')
        return existing_history, None, [], []

    market_avg = sum(stock_avgs.values()) / len(stock_avgs)
    today_label = _today()

    # Determine last completed trading date from yfinance data
    try:
        last_trade_date = raw.dropna(how='all').index[-1].strftime('%d %b')
    except Exception:
        last_trade_date = None

    # Debug: show last 3 dates found in daily_buckets
    bucket_dates = sorted(daily_buckets.keys(), key=_date_key)
    print(f'  bucket dates (last 3): {bucket_dates[-3:]} · last_trade: {last_trade_date} ({ex_key})')

    new_history = list(existing_history)
    for label in sorted(daily_buckets, key=_date_key):
        if label == today_label:
            continue  # exclude today (may be incomplete)
        vals = daily_buckets[label]
        new_history = _push(new_history,
                            {'date': label, 'illiq': round(sum(vals)/len(vals), 6)})

    # Fallback: guarantee the last trading date is always in history.
    # If high-ILLIQ filtering knocked it out of daily_buckets entirely, add it
    # using the 20-day market_avg as a proxy so the ranking date can still advance.
    if (last_trade_date and last_trade_date != today_label and
            not any(e.get('date') == last_trade_date for e in new_history)):
        print(f'  ⚠  {ex_key}: {last_trade_date} absent from daily_buckets — injecting market_avg as fallback')
        new_history = _push(new_history,
                            {'date': last_trade_date, 'illiq': round(market_avg, 6)})

    # Per-stock rankings (most → least liquid)
    sorted_stocks = sorted(stock_avgs.items(), key=lambda x: x[1])
    def _stock_entry(sym, val):
        suffix_map = {'lse':'.L','enx':None,'ndx':None,'nyse':None,'xetra':'.DE','six':'.SW'}
        suf = suffix_map.get(ex_key)
        ticker_clean = sym.replace(suf, '') if suf and sym.endswith(suf) else sym.split('.')[0]
        return {
            'ticker': ticker_clean,
            'name': names.get(sym, ticker_clean),
            'illiq': round(val, 6)
        }
    top10_liquid   = [_stock_entry(s, v) for s, v in sorted_stocks[:10]]
    top10_illiquid = [_stock_entry(s, v) for s, v in reversed(sorted_stocks[-10:])]

    n = len(stock_avgs)
    print(f'  {n} stocks · market avg ILLIQ: {market_avg:.4f} ({ex_key})')
    return new_history, round(market_avg, 6), top10_liquid, top10_illiquid

# ══════════════════════════════════════════════════════════════════════════════
#  ROLL (1984) IMPLIED SPREAD  (generic — all 6 exchanges)
#  S% = 2 × √(−Cov(r_t, r_{t+1})) × 100
# ══════════════════════════════════════════════════════════════════════════════

def _roll_spread(closes):
    """Roll (1984) implied bid-ask spread from daily close prices.
    S% = 2 × √(−Cov(r_t, r_{t+1})) × 100
    Uses the full 60-day window supplied by fetch_spread (~42 trading days).
    Bid-ask bounce dominates over this horizon, consistent with the AR period.
    Returns spread as % of price, or None if cov ≥ 0.
    """
    prices = [float(p) for p in closes if p and float(p) > 0]
    if len(prices) < 12:
        return None
    log_ret = [math.log(prices[i] / prices[i-1]) for i in range(1, len(prices))]
    if len(log_ret) < 8:
        return None
    r1, r2 = log_ret[:-1], log_ret[1:]
    m1 = sum(r1) / len(r1)
    m2 = sum(r2) / len(r2)
    cov = sum((a - m1) * (b - m2) for a, b in zip(r1, r2)) / (len(r1) - 1)
    if cov >= 0:
        return None
    s = round(2 * ((-cov) ** 0.5) * 100, 4)
    return s if 0 < s <= 5 else None   # cap at 5 % (outlier filter)

def fetch_spread(ex_key, existing_history):
    """Compute Roll implied spread for exchange.
    Returns (updated_history, marketAvg, top10_tight, top10_wide)."""
    cfg     = EXCHANGES[ex_key]
    tickers = cfg['tickers']
    names   = cfg['names']
    try:
        # 60 calendar days (~42 trading days) — consistent with AR period;
        # bid-ask bounce still dominates over this horizon.
        raw = yf.download(tickers, period='60d', interval='1d', progress=False,
                          auto_adjust=True)
        if raw.empty:
            return existing_history, None, [], []
        close_df = raw['Close'] if isinstance(raw.columns, pd.MultiIndex) else raw
    except Exception as exc:
        print(f'  ✗  {ex_key} spread download: {exc}')
        return existing_history, None, [], []

    try:
        last_trade_date = close_df.dropna(how='all').index[-1].strftime('%d %b')
    except Exception:
        last_trade_date = _today()

    stock_spreads = {}
    for sym in tickers:
        try:
            col = sym if sym in close_df.columns else None
            if col is None:
                continue
            closes = close_df[col].dropna().tolist()
            sp = _roll_spread(closes)
            if sp is not None:
                stock_spreads[sym] = sp
        except Exception:
            continue

    if not stock_spreads:
        print(f'  ✗  {ex_key}: no Roll spread computed')
        return existing_history, None, [], []

    market_avg = round(sum(stock_spreads.values()) / len(stock_spreads), 4)
    new_history = _push(existing_history,
                        {'date': last_trade_date, 'avgSpread': market_avg})

    suffix_map = {'lse':'.L','enx':None,'ndx':None,'nyse':None,'xetra':'.DE','six':'.SW'}
    suf = suffix_map.get(ex_key)
    def _sp_entry(sym, val):
        tk = sym.replace(suf, '') if suf and sym.endswith(suf) else sym.split('.')[0]
        return {'ticker': tk, 'name': names.get(sym, tk), 'spread': round(val, 4)}

    sorted_s = sorted(stock_spreads.items(), key=lambda x: x[1])
    top10_tight = [_sp_entry(s, v) for s, v in sorted_s[:10]]
    top10_wide  = [_sp_entry(s, v) for s, v in reversed(sorted_s[-10:])]

    print(f'  {len(stock_spreads)} stocks · Roll spread: {market_avg:.4f}% ({ex_key}, {last_trade_date})')
    return new_history, market_avg, top10_tight, top10_wide

# ══════════════════════════════════════════════════════════════════════════════
#  ABDI-RANALDO (2017) IMPLIED SPREAD  (generic — all 6 exchanges)
#  c_t = ln(C_t) − 0.5*(ln(H_t)+ln(L_t))
#  S% = 2 × √(−Cov(c_t, c_{t-1})) × 100
# ══════════════════════════════════════════════════════════════════════════════

def _ar_spread(highs, lows, closes):
    """Abdi-Ranaldo (2017) range-based implied bid-ask spread.
    c_t = ln(C_t) − 0.5×(ln(H_t)+ln(L_t))
    S% = 2 × √(−Cov(c_t, c_{t-1})) × 100
    c_t measures where the close lands within the day's H/L range — a pure
    microstructure signal that is insensitive to trend direction, so 60 days
    of data (≈42 trading days) reliably gives cov < 0 even in trending markets.
    Returns spread as % of price, or None if cov ≥ 0.
    """
    data = [(float(c), float(h), float(l))
            for c, h, l in zip(closes, highs, lows)
            if c and h and l and float(c) > 0 and float(h) > 0 and float(l) > 0]
    if len(data) < 12:
        return None
    ct = [math.log(c) - 0.5 * (math.log(h) + math.log(l)) for c, h, l in data]
    if len(ct) < 8:
        return None
    c1, c2 = ct[:-1], ct[1:]
    m1 = sum(c1) / len(c1)
    m2 = sum(c2) / len(c2)
    cov = sum((a - m1) * (b - m2) for a, b in zip(c1, c2)) / (len(c1) - 1)
    if cov >= 0:
        return None
    s = round(2 * ((-cov) ** 0.5) * 100, 4)
    return s if 0 < s <= 5 else None   # cap at 5 % (outlier filter)

def fetch_ar_spread(ex_key, existing_history):
    """Compute Abdi-Ranaldo spread for exchange.
    Returns (updated_history, marketAvg, top10_tight, top10_wide)."""
    cfg     = EXCHANGES[ex_key]
    tickers = cfg['tickers']
    names   = cfg['names']
    try:
        # 60 calendar days (~42 trading days) — c_t metric is trend-insensitive
        # so 42 days reliably gives cov < 0 even in trending markets.
        raw = yf.download(tickers, period='60d', interval='1d', progress=False,
                          auto_adjust=True)
        if raw.empty:
            return existing_history, None, [], []
    except Exception as exc:
        print(f'  ✗  {ex_key} AR download: {exc}')
        return existing_history, None, [], []

    is_multi = isinstance(raw.columns, pd.MultiIndex)
    try:
        close_df = raw['Close']  if is_multi else raw
        high_df  = raw['High']   if is_multi else None
        low_df   = raw['Low']    if is_multi else None
        if high_df is None or low_df is None:
            return existing_history, None, [], []
        last_trade_date = close_df.dropna(how='all').index[-1].strftime('%d %b')
    except Exception:
        last_trade_date = _today()

    stock_spreads = {}
    for sym in tickers:
        try:
            if sym not in close_df.columns:
                continue
            cl = close_df[sym].dropna()
            hi = high_df[sym].reindex(cl.index).dropna()
            lo = low_df[sym].reindex(cl.index).dropna()
            idx = cl.index.intersection(hi.index).intersection(lo.index)
            sp = _ar_spread(hi.loc[idx].tolist(), lo.loc[idx].tolist(), cl.loc[idx].tolist())
            if sp is not None:
                stock_spreads[sym] = sp
        except Exception:
            continue

    if not stock_spreads:
        print(f'  ✗  {ex_key}: no AR spread computed')
        return existing_history, None, [], []

    market_avg = round(sum(stock_spreads.values()) / len(stock_spreads), 4)
    new_history = _push(existing_history,
                        {'date': last_trade_date, 'avgSpread': market_avg})

    suffix_map = {'lse':'.L','enx':None,'ndx':None,'nyse':None,'xetra':'.DE','six':'.SW'}
    suf = suffix_map.get(ex_key)
    def _ar_entry(sym, val):
        tk = sym.replace(suf, '') if suf and sym.endswith(suf) else sym.split('.')[0]
        return {'ticker': tk, 'name': names.get(sym, tk), 'spread': round(val, 4)}

    sorted_s = sorted(stock_spreads.items(), key=lambda x: x[1])
    top10_tight = [_ar_entry(s, v) for s, v in sorted_s[:10]]
    top10_wide  = [_ar_entry(s, v) for s, v in reversed(sorted_s[-10:])]

    print(f'  {len(stock_spreads)} stocks · AR spread: {market_avg:.4f}% ({ex_key}, {last_trade_date})')
    return new_history, market_avg, top10_tight, top10_wide

# ══════════════════════════════════════════════════════════════════════════════
#  RANKING COMPUTATION  (Python — pre-computed, stored in JSON)
#  Lower rank number = more liquid (rank 1 = best)
# ══════════════════════════════════════════════════════════════════════════════

def _rank6(vals_dict, ascending=True):
    """Rank 6 exchange values 1–6. ascending=True → lowest value = rank 1.
    Exchanges with None values are excluded."""
    items = [(k, v) for k, v in vals_dict.items() if v is not None]
    items.sort(key=lambda x: x[1], reverse=not ascending)
    return {k: i+1 for i, (k, _) in enumerate(items)}

def compute_rankings(dash):
    """Build current_ranking and ranking_history from stored metric histories.
    Uses dates where at least 4/6 exchanges have ILLIQ data, excluding today.
    Exchanges missing a date's data (e.g. closed for a market holiday) fall back
    to their most recent prior value so the ranking date always advances."""
    ex_keys     = list(EXCHANGES.keys())
    today_label = _today()

    # Build per-exchange ILLIQ lookup: date → illiq (exclude today's partial data)
    illiq_by_ex = {}
    for k in ex_keys:
        illiq_by_ex[k] = {e['date']: e['illiq']
                          for e in dash['amihud'][k].get('history', [])
                          if e.get('date') != today_label}

    # Vol lookup: latest available value per exchange (not date-matched).
    vol_latest = {}
    for k in ex_keys:
        vd = dash['vol'][k]
        vals = vd.get('value', [])
        vol_latest[k] = vals[-1] if vals else None

    # Spread and AR lookups (exclude today — spread/AR use last_trade_date which
    # could be today's partial bar when US markets are open during the run)
    spread_by_ex = {}
    for k in ex_keys:
        spread_by_ex[k] = {e['date']: e['avgSpread']
                           for e in dash['spread'][k].get('history', [])
                           if e.get('date') != today_label}

    ar_by_ex = {}
    for k in ex_keys:
        ar_by_ex[k] = {e['date']: e['avgSpread']
                       for e in dash['ar_spread'][k].get('history', [])
                       if e.get('date') != today_label}

    # Debug: show latest ILLIQ dates per exchange to identify gaps
    for k in ex_keys:
        dates = sorted(illiq_by_ex[k].keys(), key=_date_key)
        print(f'  ILLIQ {k}: {len(dates)} dates · latest 3: {dates[-3:] if len(dates) >= 3 else dates}')

    # Date spine: dates where at least 4 of 6 exchanges have ILLIQ.
    # This tolerates single-market holidays (e.g. US observed Independence Day,
    # UK bank holidays) where 1–2 exchanges have no data for that date.
    all_dates = set()
    for k in ex_keys:
        all_dates.update(illiq_by_ex[k].keys())

    min_open = max(3, len(ex_keys) - 2)   # at least 4 of 6
    complete_dates = sorted(
        [d for d in all_dates
         if sum(1 for k in ex_keys if d in illiq_by_ex[k]) >= min_open],
        key=_date_key
    )

    if not complete_dates:
        print('  ⚠  Rankings: no complete dates found')
        return

    def _best_on_or_before(lookup, d):
        """Return lookup[d] if present; else the most recent entry with date ≤ d."""
        v = lookup.get(d)
        if v is not None:
            return v
        prior = sorted([dt for dt in lookup if _date_key(dt) <= _date_key(d)],
                       key=_date_key)
        return lookup[prior[-1]] if prior else None

    history = []
    for d in complete_dates:
        # Use exact date if available; fall back to most recent prior value
        # so a closed exchange (holiday) is ranked on its last trading day.
        illiq_vals  = {k: _best_on_or_before(illiq_by_ex[k],  d) for k in ex_keys}
        vol_vals    = {k: vol_latest[k]                            for k in ex_keys}
        spread_vals = {k: _best_on_or_before(spread_by_ex[k], d) for k in ex_keys}
        ar_vals     = {k: _best_on_or_before(ar_by_ex[k],     d) for k in ex_keys}

        r_illiq  = _rank6(illiq_vals,  ascending=True)   # lower ILLIQ = rank 1
        r_spread = _rank6(spread_vals, ascending=True)   # lower spread = rank 1
        r_ar     = _rank6(ar_vals,     ascending=True)   # lower AR = rank 1
        r_vol    = _rank6(vol_vals,    ascending=False)  # higher vol = rank 1

        # Composite: average of available sub-ranks, then re-rank 1-6
        composites = {}
        for k in ex_keys:
            sub = [r_illiq.get(k), r_spread.get(k), r_ar.get(k), r_vol.get(k)]
            avail = [x for x in sub if x is not None]
            if avail:
                composites[k] = sum(avail) / len(avail)
        r_comp = _rank6(composites, ascending=True)

        ranks = {}
        for k in ex_keys:
            ranks[k] = {
                'vol':      r_vol.get(k),
                'illiq':    r_illiq.get(k),
                'spread':   r_spread.get(k),
                'ar':       r_ar.get(k),
                'composite': r_comp.get(k),
            }
        history.append({'date': d, 'ranks': ranks})

    dash['ranking_history'] = history[-DAYS:]
    last = history[-1]
    dash['current_ranking'] = {'date': last['date'], 'ranks': last['ranks']}
    print(f'  Ranking computed · {len(complete_dates)} complete days · current: {last["date"]}')
    # Print ranking summary
    sorted_ex = sorted(ex_keys, key=lambda k: last['ranks'][k].get('composite') or 99)
    for rank, k in enumerate(sorted_ex, 1):
        print(f'    #{rank} {EXCHANGES[k]["name"]} (composite rank {last["ranks"][k].get("composite")})')

# ══════════════════════════════════════════════════════════════════════════════
#  AI COMMENTARY  (Anthropic Claude — runs once per day at morning trigger)
# ══════════════════════════════════════════════════════════════════════════════

def generate_commentary(dash):
    """Generate a short AI ranking commentary. Returns string or None."""
    if _ant is None:
        print('  ⚠  anthropic not installed — commentary skipped')
        return None

    api_key = os.environ.get('ANTHROPIC_API_KEY')
    if not api_key:
        print('  ⚠  ANTHROPIC_API_KEY not set — commentary skipped')
        return None

    cr = dash.get('current_ranking', {})
    ranking_date = cr.get('date', 'unknown date')
    ranks = cr.get('ranks', {})

    if not ranks:
        print('  ⚠  No ranking data for commentary')
        return None

    sorted_ex = sorted(EXCHANGES.keys(),
                       key=lambda k: ranks.get(k, {}).get('composite') or 99)

    def amihud_val(k):
        avg = dash['amihud'][k].get('marketAvg')
        return f'{avg:.4f}' if avg is not None else 'N/A'
    def spread_val(k):
        avg = dash['spread'][k].get('marketAvg')
        return f'{avg:.4f}%' if avg is not None else 'N/A'
    def ar_val(k):
        avg = dash['ar_spread'][k].get('marketAvg')
        return f'{avg:.4f}%' if avg is not None else 'N/A'

    ranking_summary = '\n'.join(
        f'  #{i+1} {EXCHANGES[k]["name"]}: composite rank {ranks[k].get("composite")}, '
        f'ILLIQ={amihud_val(k)}, Roll spread={spread_val(k)}, AR spread={ar_val(k)}'
        for i, k in enumerate(sorted_ex)
    )

    prompt = f"""You are a financial markets analyst writing a brief commentary on exchange liquidity rankings.

Today's data is for {ranking_date}.

Exchange Liquidity Ranking (1 = most liquid):
{ranking_summary}

Metrics explained:
- ILLIQ (Amihud): % price move per 1M currency unit traded — lower = more liquid
- Roll spread: implied bid-ask spread (Roll 1984) — lower = tighter spreads
- AR spread: Abdi-Ranaldo (2017) implied spread — lower = tighter spreads

Write a concise 3-4 sentence commentary (max 80 words) for a professional audience. Focus on the top and bottom ranked exchanges and any notable differences in the metrics. Be specific and analytical. Do not use bullet points."""

    try:
        client = _ant.Anthropic(api_key=api_key)
        msg = client.messages.create(
            model='claude-haiku-4-5-20251001',
            max_tokens=200,
            messages=[{'role': 'user', 'content': prompt}]
        )
        text = msg.content[0].text.strip()
        print(f'  ✓  Commentary generated ({len(text)} chars)')
        return text
    except Exception as exc:
        print(f'  ✗  Commentary generation failed: {exc}')
        return None

# ══════════════════════════════════════════════════════════════════════════════
#  MAIN
# ══════════════════════════════════════════════════════════════════════════════

def main():
    # ── CLI args ─────────────────────────────────────────────────────────────
    parser = argparse.ArgumentParser(description='Finance Dashboard Updater')
    parser.add_argument('--fast', action='store_true',
                        help='Fast mode: update ticker bar, gainers, and market cap only')
    args = parser.parse_args()
    FAST_MODE = args.fast

    today_iso = date.today().isoformat()
    state     = {}
    if STATE_FILE.exists():
        try:
            state = json.loads(STATE_FILE.read_text())
        except Exception:
            pass

    mode_label = 'FAST (news / ticker / gainers / market cap)' if FAST_MODE else 'FULL'
    print('\n═══ Finance Dashboard Updater ═══')
    print(f'Date: {date.today()} UTC   Mode: {mode_label}\n')

    dash = _load_dashboard()

    # ── News (every run — headlines rarely change within an hour) ────────────
    print('── News')
    news = fetch_news()
    if news:
        dash['news'] = news

    # ── Ticker bar ──────────────────────────────────────────────────────────
    print('\n── Tickers')
    tickers = fetch_tickers()
    if tickers:
        dash['tickers'] = tickers

    # ── Per-exchange data ────────────────────────────────────────────────────
    for ex_key, cfg in EXCHANGES.items():
        print(f'\n── {cfg["name"]} ({ex_key})')

        # Gainers
        print('   Gainers...')
        gainers = fetch_gainers(ex_key)
        if gainers:
            dash['gainers'][ex_key] = gainers

        # Market Cap
        print('   Market Cap...')
        mcap = fetch_market_cap(ex_key)
        if mcap:
            dash['market_cap'][ex_key] = mcap

        if FAST_MODE:
            continue   # skip all analytics in fast mode

        # Volume
        print('   Volume...')
        if cfg['vol_method'] == 'file':
            vol_pts = fetch_vol_file(ex_key, state)
        else:
            vol_pts = fetch_vol_yf(ex_key)
        if vol_pts:
            # Merge into existing vol history (keep last 5 for display)
            existing_dates = list(dash['vol'][ex_key].get('dates', []))
            existing_vals  = list(dash['vol'][ex_key].get('value', []))
            vol_map = dict(zip(existing_dates, existing_vals))
            for pt in vol_pts:
                vol_map[pt['label']] = pt['value']
            sorted_labels = sorted(vol_map.keys(), key=_date_key)[-5:]
            dash['vol'][ex_key]['dates'] = sorted_labels
            dash['vol'][ex_key]['value'] = [vol_map[l] for l in sorted_labels]
            # Shares traded (yfinance exchanges only — file-based pts have no shares_m)
            valid_sh = [pt['shares_m'] for pt in vol_pts if pt.get('shares_m') is not None]
            if valid_sh:
                dash['vol'][ex_key]['shares_latest'] = valid_sh[-1]
                dash['vol'][ex_key]['shares_avg']    = round(sum(valid_sh) / len(valid_sh), 1)
            # Trades count (file-based exchanges: LSE in thousands, ENX in millions)
            valid_tr = [pt['trades'] for pt in vol_pts if pt.get('trades') is not None]
            if valid_tr:
                dash['vol'][ex_key]['trades_latest'] = valid_tr[-1]
                dash['vol'][ex_key]['trades_avg']    = round(sum(valid_tr) / len(valid_tr), 1)

        # Amihud ILLIQ
        print('   Amihud ILLIQ...')
        new_hist, mkt_avg, top_liq, top_illiq = fetch_amihud(
            ex_key, dash['amihud'][ex_key].get('history', []))
        dash['amihud'][ex_key]['history'] = new_hist
        if mkt_avg is not None:
            dash['amihud'][ex_key]['marketAvg']      = mkt_avg
            dash['amihud'][ex_key]['top10_liquid']   = top_liq
            dash['amihud'][ex_key]['top10_illiquid'] = top_illiq

        # Roll Spread
        print('   Roll spread...')
        new_hist, mkt_avg, top_tight, top_wide = fetch_spread(
            ex_key, dash['spread'][ex_key].get('history', []))
        dash['spread'][ex_key]['history'] = new_hist
        if mkt_avg is not None:
            dash['spread'][ex_key]['marketAvg']   = mkt_avg
            dash['spread'][ex_key]['top10_tight'] = top_tight
            dash['spread'][ex_key]['top10_wide']  = top_wide

        # AR Spread
        print('   AR spread...')
        new_hist, mkt_avg, top_tight, top_wide = fetch_ar_spread(
            ex_key, dash['ar_spread'][ex_key].get('history', []))
        dash['ar_spread'][ex_key]['history'] = new_hist
        if mkt_avg is not None:
            dash['ar_spread'][ex_key]['marketAvg']   = mkt_avg
            dash['ar_spread'][ex_key]['top10_tight'] = top_tight
            dash['ar_spread'][ex_key]['top10_wide']  = top_wide

    if not FAST_MODE:
        # ── Rankings ─────────────────────────────────────────────────────────
        print('\n── Rankings')
        compute_rankings(dash)

        # ── AI Commentary ────────────────────────────────────────────────────
        print('\n── Commentary')
        # Regenerate if: no commentary yet, new day, or ranking date changed
        ranking_date = dash.get('current_ranking', {}).get('date')
        stale = (state.get('commentary_date') != today_iso or
                 state.get('commentary_ranking_date') != ranking_date or
                 not dash.get('commentary'))
        if not stale:
            print('  ℹ  Reusing today\'s commentary')
        else:
            commentary = generate_commentary(dash)
            if commentary:
                dash['commentary']               = commentary
                dash['commentary_date']          = today_iso
                state['commentary_date']         = today_iso
                state['commentary_ranking_date'] = ranking_date

    # ── Save ──────────────────────────────────────────────────────────────────
    _save_dashboard(dash)
    STATE_FILE.write_text(json.dumps(state, indent=2))
    print('  ✓  State saved')
    print('\n═══ Done ═══\n')

if __name__ == '__main__':
    main()
