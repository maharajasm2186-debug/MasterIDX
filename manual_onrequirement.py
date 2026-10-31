import os, sys, re, smtplib, requests, pandas as pd
from io import StringIO
from datetime import datetime, date, timedelta
from email.message import EmailMessage
from pathlib import Path
import configparser, time, threading
from concurrent.futures import ThreadPoolExecutor
import holidays as us_holidays

# =============================================================================
# MANUAL SEC FILING REPORT  (branch version — NOT scheduled)
# -----------------------------------------------------------------------------
# • Runs ONLY when you invoke:   python manual_onrequirement.py
# • All requirements live in manual_config.ini (form types + date period).
# • Independent of the daily cron script — never modifies it.
# • NEW: form-type variants are grouped (10-K, 10-Q, 8-K, DEF, 20-F, 40-F ...)
# • NEW: Russell / S&P / NASDAQ membership flags in Excel + tabulated email
# =============================================================================

CONFIG_FILE = Path(__file__).resolve().with_name("manual_config.ini")
if not CONFIG_FILE.exists():
    raise SystemExit(f"Config not found: {CONFIG_FILE}")

cfg = configparser.ConfigParser(interpolation=None)
cfg.read(CONFIG_FILE)
if "report" not in cfg:
    raise SystemExit("[report] section missing in manual_config.ini")
section = cfg["report"]

# ---------- Form types ----------
raw_forms = section.get("form_types", "").strip()
if not raw_forms:
    FORM_PREFIXES = ("8-K","10-K","10-Q","DEF 14","DEFC14","DEFM14",
                     "DEFR14","DEFN14","20-F","40-F")
    print(f"form_types blank → defaults: {FORM_PREFIXES}")
elif raw_forms.upper() == "ALL":
    FORM_PREFIXES = None
    print("form_types = ALL → no form filter")
else:
    FORM_PREFIXES = tuple(f.strip().upper() for f in raw_forms.split(",") if f.strip())
    print(f"form_types: {FORM_PREFIXES}")

# ---------- Date period ----------
def _pdate(s): return datetime.strptime(s.strip(), "%Y-%m-%d").date()

single  = section.get("date", "").strip()
start_s = section.get("start_date", "").strip()
end_s   = section.get("end_date", "").strip()

if single:
    start = end = _pdate(single)
elif start_s and end_s:
    start, end = _pdate(start_s), _pdate(end_s)
elif start_s:
    start = end = _pdate(start_s)
elif end_s:
    start = end = _pdate(end_s)
else:
    start = end = date.today() - timedelta(days=section.getint("days_back", 1))

if end < start:
    raise SystemExit(f"end_date ({end}) is before start_date ({start})")

all_dates = [start + timedelta(days=i) for i in range((end - start).days + 1)]
print(f"Date period: {start} → {end}   ({len(all_dates)} calendar days)")

# ---------- Recipients ----------
RECIPIENTS = [r.strip() for r in section.get("recipients", "").split(",") if r.strip()] or [
    "maharaja@secanalyzer.net",
]

# =============================================================================
# FORM GROUPING  (variants → one group)
# =============================================================================
GROUP_RULES = [
    (re.compile(r"^10-K"),              "10-K"),   # 10-K, 10-KT, 10-K/A, 10-KT/A, 10-K405
    (re.compile(r"^10-Q"),              "10-Q"),   # 10-Q, 10-QT, 10-Q/A, 10-QT/A
    (re.compile(r"^8-K"),               "8-K"),    # 8-K, 8-K/A, 8-K12B, 8-K12G3, 8-K15D5
    (re.compile(r"^DEF[A-Z]? ?14[AC]"), "DEF"),    # DEF 14A, DEFC14A, DEFR14A, DEFM14A,
                                                   # DEFN14A, DEFC14C, DEFR14C, DEF 14C ...
    (re.compile(r"^20-F"),              "20-F"),
    (re.compile(r"^40-F"),              "40-F"),
]
GROUP_ORDER = ["8-K", "10-K", "10-Q", "DEF", "20-F", "40-F"]

def form_group(ft):
    s = str(ft).strip().upper()
    for rx, g in GROUP_RULES:
        if rx.match(s):
            return g
    return s          # anything else keeps its own form type as group

# =============================================================================
# WEEKEND / HOLIDAY GUARD
# =============================================================================
_hol_cache = {}
def non_trading_reason(d):
    if d.weekday() == 5: return "Saturday"
    if d.weekday() == 6: return "Sunday"
    _hol_cache.setdefault(d.year, us_holidays.UnitedStates(years=d.year))
    if d in _hol_cache[d.year]:
        return f"US Federal Holiday — {_hol_cache[d.year][d]}"
    return None

# =============================================================================
# HTTP
# =============================================================================
session = requests.Session()
session.headers.update({"User-Agent": "SEC Research your_email@example.com"})
MAX_WORKERS, REQ_PER_SEC, TIMEOUT = 10, 9, 30

class RateLimiter:
    def __init__(self, mps):
        self.min_interval = 1.0 / mps
        self.lock = threading.Lock(); self.next = time.monotonic()
    def wait(self):
        with self.lock:
            now = time.monotonic()
            s = max(0.0, self.next - now)
            self.next = max(now, self.next) + self.min_interval
        if s > 0: time.sleep(s)
limiter = RateLimiter(REQ_PER_SEC)

# =============================================================================
# INDEX MEMBERSHIP  (Russell, S&P, NASDAQ)
# -----------------------------------------------------------------------------
# Auto-loaded from public sources and cached in index_cache.csv (refreshed
# every `index_refresh_days`, default 7).  You can ALWAYS add / override with
# a local file  index_constituents.csv  having columns:  Index,Ticker,CIK
#   (Index must be exactly one of INDEX_NAMES above; Ticker or CIK may be blank)
# =============================================================================
INDEX_NAMES = ["Russell 1000", "Russell 2000", "Russell 3000",
               "S&P 100", "S&P 500", "S&P 400", "S&P 600", "NASDAQ 100"]
FLAG_COLS   = ["In_" + re.sub(r"\W+", "_", n) for n in INDEX_NAMES]
BASE_DIR    = Path(__file__).resolve().parent
INDEX_CACHE = BASE_DIR / "index_cache_v2.csv"
INDEX_LOCAL = BASE_DIR / "index_constituents.csv"
REFRESH_DAYS = section.getint("index_refresh_days", 7)
WEB_HDR = {"User-Agent": "SECReportBot/1.0 (contact: your_email@example.com)",
           "Accept": "text/html,text/csv,*/*"}
LOAD_ERR = {}          # index name → why it failed to load

def _nt(t):   # normalise ticker
    return str(t).strip().upper().replace(".", "-").replace("/", "-")

def _web(url, referer=None):
    h = dict(WEB_HDR)
    if referer: h["Referer"] = referer
    r = requests.get(url, headers=h, timeout=TIMEOUT)
    r.raise_for_status()
    return r.text

def _wiki(page, label):
    rows = []
    try:
        html = _web(f"https://en.wikipedia.org/wiki/{page}")
        for t in pd.read_html(StringIO(html)):
            col = next((c for c in t.columns if str(c).strip() in
                        ("Symbol", "Ticker symbol", "Ticker")), None)
            if col is None: continue
            cikc = next((c for c in t.columns if str(c).strip() == "CIK"), None)
            for _, r in t.iterrows():
                if pd.isna(r[col]): continue
                rows.append((label, _nt(r[col]),
                             str(int(r[cikc])) if cikc is not None and pd.notna(r[cikc]) else ""))
            break
        if not rows: LOAD_ERR[label] = "page loaded but no constituents table found"
    except Exception as e:
        LOAD_ERR[label] = f"{type(e).__name__}: {str(e)[:140]}"
    return rows

def _ishares(pid, slug, fund, label):
    rows = []
    base = f"https://www.ishares.com/us/products/{pid}/{slug}"
    url = f"{base}/1467271812596.ajax?fileType=csv&fileName={fund}_holdings&dataType=fund"
    try:
        lines = _web(url, referer=base).lstrip("\ufeff").splitlines()
        h = next(i for i, l in enumerate(lines) if l.strip('"\ufeff ').startswith("Ticker"))
        df = pd.read_csv(StringIO("\n".join(lines[h:])), dtype=str)
        if "Asset Class" in df.columns: df = df[df["Asset Class"] == "Equity"]
        rows = [(label, _nt(t), "") for t in df["Ticker"].dropna()]
        if not rows: LOAD_ERR[label] = "file loaded but no tickers parsed"
    except StopIteration:
        LOAD_ERR[label] = "response was not a holdings CSV (site blocked or URL changed)"
    except Exception as e:
        LOAD_ERR[label] = f"{type(e).__name__}: {str(e)[:140]}"
    return rows

LOADERS = {
    "S&P 500":      lambda: _wiki("List_of_S%26P_500_companies", "S&P 500"),
    "S&P 400":      lambda: _wiki("List_of_S%26P_400_companies", "S&P 400"),
    "S&P 600":      lambda: _wiki("List_of_S%26P_600_companies", "S&P 600"),
    "S&P 100":      lambda: _wiki("S%26P_100", "S&P 100"),
    "NASDAQ 100":   lambda: _wiki("Nasdaq-100", "NASDAQ 100"),
    "Russell 1000": lambda: _ishares(239707, "ishares-russell-1000-etf", "IWB", "Russell 1000"),
    "Russell 2000": lambda: _ishares(239710, "ishares-russell-2000-etf", "IWM", "Russell 2000"),
    "Russell 3000": lambda: _ishares(239714, "ishares-russell-3000-etf", "IWV", "Russell 3000"),
}

def load_index_members():
    cached = pd.DataFrame(columns=["Index", "Ticker", "CIK"])
    fresh = False
    if INDEX_CACHE.exists():
        try:
            cached = pd.read_csv(INDEX_CACHE, dtype=str).fillna("")
            fresh = (time.time() - INDEX_CACHE.stat().st_mtime) < REFRESH_DAYS*86400
        except Exception:
            pass
    have = set(cached["Index"]) if fresh else set()
    df = cached if fresh else pd.DataFrame(columns=["Index", "Ticker", "CIK"])
    missing = [n for n in INDEX_NAMES if n not in have]
    fetched = []
    for n in missing:
        print(f"  Loading {n} ...")
        rows = LOADERS[n]()
        print(f"    → {len(rows)} rows" + (f"   ! {LOAD_ERR[n]}" if n in LOAD_ERR else ""))
        if rows:
            fetched += rows
        else:                                   # fetch failed → fall back to old committed/cached copy
            old = cached[cached["Index"] == n]
            if not fresh and not old.empty:
                df = pd.concat([df, old], ignore_index=True)
                LOAD_ERR[n] = LOAD_ERR.get(n, "no rows") + "  (using OLD cached list)"
                print(f"    → using {len(old)} rows from old cache")
    if fetched:
        df = pd.concat([df, pd.DataFrame(fetched, columns=["Index", "Ticker", "CIK"])],
                       ignore_index=True)
        df.to_csv(INDEX_CACHE, index=False)       # only indices that actually loaded are cached
    if INDEX_LOCAL.exists():
        loc = pd.read_csv(INDEX_LOCAL, dtype=str).fillna("")
        for c in ("Index", "Ticker", "CIK"):
            if c not in loc.columns: loc[c] = ""
        loc["Ticker"] = loc["Ticker"].apply(lambda x: _nt(x) if x else "")
        df = pd.concat([df, loc[["Index", "Ticker", "CIK"]]], ignore_index=True)
        print(f"Local overrides merged: {INDEX_LOCAL.name} ({len(loc)} rows)")
    df["CIK"] = df["CIK"].astype(str).str.lstrip("0")
    t_map, c_map = {}, {}
    for ix, tk, ck in df[["Index", "Ticker", "CIK"]].itertuples(index=False):
        if tk: t_map.setdefault(tk, set()).add(ix)
        if ck: c_map.setdefault(ck, set()).add(ix)
    counts = df.groupby("Index").size().to_dict() if not df.empty else {}
    print("Index constituents available:")
    for n in INDEX_NAMES:
        print(f"  {n:13s}: {counts.get(n, 0)}")
    warns = [n for n in INDEX_NAMES if counts.get(n, 0) == 0]
    if warns:
        print("\n" + "!"*70 +
              f"\n!! NO constituents for: {', '.join(warns)}\n"
              "!! Index columns will be 0 for these. Fix network access or supply\n"
              "!! index_constituents.csv (Index,Ticker,CIK).\n" + "!"*70 + "\n")
    return t_map, c_map, warns, counts

TICKER_IDX, CIK_IDX, INDEX_WARN, INDEX_COUNTS = load_index_members()

def membership(cik, tickers):
    s = set(CIK_IDX.get(str(cik).lstrip("0"), set()))
    for t in str(tickers or "").split(","):
        t = _nt(t)
        if t: s |= TICKER_IDX.get(t, set())
    return s

# =============================================================================
# EDGAR INDEX
# =============================================================================
IDX_COLS = ["CIK","Company","Form_Type","Date_Filed","File_Name"]

def _parse_idx(text):
    lines = [l for l in text.split("\n") if l.count("|") >= 4]
    if not lines: return pd.DataFrame(columns=IDX_COLS)
    df = pd.read_csv(StringIO("\n".join(lines)), sep="|",
                     names=IDX_COLS, dtype=str)
    for c in IDX_COLS: df[c] = df[c].astype(str).str.strip()
    return df

def fetch_idx_for_date(d):
    y, q = d.year, (d.month - 1)//3 + 1
    nod = d.strftime("%Y%m%d"); ds = d.strftime("%Y-%m-%d")
    daily = f"https://www.sec.gov/Archives/edgar/daily-index/{y}/QTR{q}/master{nod}.idx"
    qtr   = f"https://www.sec.gov/Archives/edgar/full-index/{y}/QTR{q}/master.idx"

    frames = []
    for label, url in (("Daily", daily), ("Quarterly", qtr)):
        try:
            limiter.wait()
            r = session.get(url, timeout=TIMEOUT)
            if r.status_code == 200 and "|" in r.text:
                dfd = _parse_idx(r.text)
                dfd = dfd[dfd["Date_Filed"] == ds]
                frames.append(dfd)
                print(f"  ✓ {label} IDX → {len(dfd)} filings")
        except Exception as e:
            print(f"  ! {label} IDX fetch failed: {e}")

    if not frames: return pd.DataFrame(columns=IDX_COLS)
    return (pd.concat(frames, ignore_index=True)
              .drop_duplicates(subset=["File_Name"])
              .reset_index(drop=True))

def matches(ft):
    if FORM_PREFIXES is None: return True
    return any(str(ft).upper().startswith(p) for p in FORM_PREFIXES)

# =============================================================================
# LOOP OVER DATE RANGE
# =============================================================================
all_frames, skipped, empty_days = [], [], []
idx_totals = {}          # IDX date → total filings released that day (ALL forms)
for d in all_dates:
    ds = d.strftime("%Y-%m-%d")
    reason = non_trading_reason(d)
    if reason:
        print(f"\n[{ds}] SKIP — {reason}")
        skipped.append((ds, reason)); continue

    print(f"\n[{ds}] Fetching EDGAR index ...")
    df_day = fetch_idx_for_date(d)
    if df_day.empty:
        print("  No index entries."); empty_days.append(ds); continue

    idx_totals[ds] = len(df_day)
    df_filt = df_day[df_day["Form_Type"].apply(matches)].copy()
    print(f"  → {len(df_filt)} filings match form filter")
    if not df_filt.empty: all_frames.append(df_filt)

if not all_frames:
    print("\nNo filings matched. Nothing to report."); sys.exit(0)

df_filings = pd.concat(all_frames, ignore_index=True)
print(f"\nTotal matched: {len(df_filings)}")

df_filings["Accession_No"] = df_filings["File_Name"].apply(
    lambda x: x.split("/")[-1].replace(".txt",""))
df_filings["Filing_URL"] = "https://www.sec.gov/Archives/" + df_filings["File_Name"]
df_filings["_acc_key"]   = df_filings["Accession_No"].str.replace("-","",regex=False)
df_filings["Form_Group"] = df_filings["Form_Type"].apply(form_group)

# =============================================================================
# ENRICH — every SEC property per filing
# =============================================================================
RENAME = {
    "reportDate":"Report_Date","acceptanceDateTime":"Acceptance_DateTime",
    "act":"Act","fileNumber":"File_Number","filmNumber":"Film_Number",
    "items":"Items","size":"Size","isXBRL":"Is_XBRL",
    "isInlineXBRL":"Is_Inline_XBRL","primaryDocument":"Primary_Document",
    "primaryDocDescription":"Primary_Doc_Description",
}

def fetch_cik(cik_raw):
    cik10 = str(cik_raw).zfill(10)
    api = f"https://data.sec.gov/submissions/CIK{cik10}.json"
    limiter.wait()
    try:
        r = session.get(api, timeout=TIMEOUT); data = r.json()
    except Exception as e:
        print(f"  ! submissions failed for {cik_raw}: {e}"); return {}, {}

    a = (data.get("addresses") or {}).get("business") or {}
    m = (data.get("addresses") or {}).get("mailing")  or {}

    company = {
        "Company_Name_API": data.get("name"),
        "Tickers": ", ".join(data.get("tickers") or []),
        "SIC": data.get("sic"),
        "Industry": data.get("sicDescription"),
        "State_of_Incorporation": data.get("stateOfIncorporation"),
        "State_of_Incorporation_Desc": data.get("stateOfIncorporationDescription"),
        "Fiscal_Year_End": data.get("fiscalYearEnd"),
        "Entity_Type": data.get("entityType"),
        "Phone": data.get("phone"),
        "Business_Street": a.get("street1"), "Business_Street2": a.get("street2"),
        "Business_City": a.get("city"), "Business_State": a.get("stateOrCountry"),
        "Business_State_Desc": a.get("stateOrCountryDescription"),
        "Business_Zip": a.get("zipCode"),
        "Mailing_Street": m.get("street1"), "Mailing_Street2": m.get("street2"),
        "Mailing_City": m.get("city"), "Mailing_State": m.get("stateOrCountry"),
        "Mailing_Zip": m.get("zipCode"),
        "Former_Names": "; ".join(fn.get("name","") for fn in (data.get("formerNames") or [])),
    }

    recent = (data.get("filings") or {}).get("recent") or {}
    accs = recent.get("accessionNumber") or []
    n = len(accs)
    fmap = {}
    for i, acc in enumerate(accs):
        key = acc.replace("-","")
        fmap[key] = {new:(recent.get(old) or [None]*n)[i]
                     for old,new in RENAME.items()}
    return company, fmap

unique_ciks = df_filings["CIK"].unique().tolist()
print(f"\nFetching properties for {len(unique_ciks)} unique companies ...")
cik_company, cik_filings = {}, {}
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
    for cik, (c, f) in zip(unique_ciks, ex.map(fetch_cik, unique_ciks)):
        cik_company[cik] = c; cik_filings[cik] = f

def _enrich(row):
    out = {}
    out.update(cik_company.get(row["CIK"], {}))
    out.update(cik_filings.get(row["CIK"], {}).get(row["_acc_key"], {}))
    return out

enriched = df_filings.apply(_enrich, axis=1, result_type="expand")
if enriched.empty:
    for c in list(RENAME.values()) + [
        "Tickers","SIC","Industry","State_of_Incorporation","Fiscal_Year_End",
        "Entity_Type","Phone","Former_Names","Business_Street","Business_City",
        "Business_State","Business_Zip","Mailing_Street","Mailing_City",
        "Mailing_State","Mailing_Zip"]:
        enriched[c] = pd.Series(dtype=object)

df_full = pd.concat([df_filings.reset_index(drop=True), enriched], axis=1).drop(columns=["_acc_key"])

def _has_item(x, code):
    if not x: return 0
    return 1 if code in [p.strip() for p in str(x).split(",")] else 0
df_full["Item_5_02"] = df_full["Items"].apply(lambda x: _has_item(x, "5.02"))
df_full["Item_4_01"] = df_full["Items"].apply(lambda x: _has_item(x, "4.01"))
df_full["Item_5_07"] = df_full["Items"].apply(lambda x: _has_item(x, "5.07"))

# ---------- Index membership (Russell / S&P / NASDAQ) ----------
_mem = df_full.apply(lambda r: membership(r["CIK"], r.get("Tickers")), axis=1)
for _n, _c in zip(INDEX_NAMES, FLAG_COLS):
    df_full[_c] = _mem.apply(lambda x, n=_n: 1 if n in x else 0)
df_full["In_Any_Index"] = (df_full[FLAG_COLS].sum(axis=1) > 0).astype(int)
df_full["Index_Membership"] = _mem.apply(
    lambda s: "; ".join(n for n in INDEX_NAMES if n in s) or "No Index")
df_full["Industry"] = df_full["Industry"].fillna("Unknown")

# ---------- Index load / match diagnostics ----------
index_status = pd.DataFrame([{
    "Index": n,
    "Constituents Loaded": INDEX_COUNTS.get(n, 0),
    "Filings Matched": int(df_full[c].sum()),
    "Companies Matched": df_full.loc[df_full[c] == 1, "CIK"].nunique(),
    "Load Problem": LOAD_ERR.get(n, ""),
} for n, c in zip(INDEX_NAMES, FLAG_COLS)])
print("\nIndex match summary:\n" + index_status.to_string(index=False))

ORDER = [
    "Date_Filed","CIK","Company","Company_Name_API","Tickers",
    "Form_Type","Form_Group","Index_Membership",*FLAG_COLS,
    "In_Any_Index","Report_Date",
    "Accession_No","Filing_URL","Acceptance_DateTime","Act","File_Number",
    "Film_Number","Size","Is_XBRL","Is_Inline_XBRL","Primary_Document",
    "Primary_Doc_Description","Items","Item_5_02","Item_4_01","Item_5_07",
    "SIC","Industry","State_of_Incorporation","State_of_Incorporation_Desc",
    "Fiscal_Year_End","Entity_Type","Phone","Former_Names",
    "Business_Street","Business_Street2","Business_City","Business_State",
    "Business_State_Desc","Business_Zip","Mailing_Street","Mailing_Street2",
    "Mailing_City","Mailing_State","Mailing_Zip",
]
for c in ORDER:
    if c not in df_full.columns: df_full[c] = None
df_full = df_full[ORDER]

# =============================================================================
# SUMMARY TABLE  (one row per IDX date × form group, then TOTAL row)
# =============================================================================
COLS = [
    "IDX Date",
    "Form Group",
    "Companies Released",
    "Industries",
    "Companies NOT in Russell/S&P/NASDAQ",
    "Companies in Russell/S&P/NASDAQ",
    "Total IDX Date Releases",
]

def _gkey(g):
    return (GROUP_ORDER.index(g) if g in GROUP_ORDER else len(GROUP_ORDER), g)

def build_summary(df, totals):
    rows = []
    for dt in sorted(df["Date_Filed"].unique()):
        g_day = df[df["Date_Filed"] == dt]
        for fg in sorted(g_day["Form_Group"].unique(), key=_gkey):
            h = g_day[g_day["Form_Group"] == fg]
            n_in  = h.loc[h["In_Any_Index"] == 1, "CIK"].nunique()
            n_out = h.loc[h["In_Any_Index"] == 0, "CIK"].nunique()
            rows.append([dt, fg, h["CIK"].nunique(), h["Industry"].nunique(),
                         n_out, n_in, totals.get(dt, 0)])
    out = pd.DataFrame(rows, columns=COLS)
    dates_present = out["IDX Date"].unique()
    total = ["TOTAL", "",
             int(out[COLS[2]].sum()), int(out[COLS[3]].sum()),
             int(out[COLS[4]].sum()), int(out[COLS[5]].sum()),
             int(sum(totals.get(d, 0) for d in dates_present))]
    out.loc[len(out)] = total
    return out

summary = build_summary(df_full, idx_totals)

# ---------- Mail table 1: Form Group over the whole IDX date range ----------
PERIOD_DESC = (start.strftime("%Y-%m-%d") if start == end else f"{start} → {end}")
T1_COLS = ["Form Group", "Range of IDX Date", "Companies Released",
           "Companies in Indices", "Companies Not in Indices"]
T2_COLS = ["IDX Date", "Companies Released",
           "Companies in Indices", "Companies Not in Indices"]

def _counts(h):
    return [h["CIK"].nunique(),
            h.loc[h["In_Any_Index"] == 1, "CIK"].nunique(),
            h.loc[h["In_Any_Index"] == 0, "CIK"].nunique()]

def build_formgroup_table(df, period):
    rows = [[fg, period] + _counts(df[df["Form_Group"] == fg])
            for fg in sorted(df["Form_Group"].unique(), key=_gkey)]
    out = pd.DataFrame(rows, columns=T1_COLS)
    out.loc[len(out)] = ["TOTAL", period] + [int(out[c].sum()) for c in T1_COLS[2:]]
    return out

# ---------- Mail table 2: date wise (all form groups together) ----------
def build_date_table(df):
    rows = [[dt] + _counts(df[df["Date_Filed"] == dt])
            for dt in sorted(df["Date_Filed"].unique())]
    out = pd.DataFrame(rows, columns=T2_COLS)
    out.loc[len(out)] = ["TOTAL"] + [int(out[c].sum()) for c in T2_COLS[1:]]
    return out

table_formgroup = build_formgroup_table(df_full, PERIOD_DESC)
table_datewise  = build_date_table(df_full)

industry_breakdown = (
    df_full.groupby(["Date_Filed", "Form_Group", "Industry"])
    .agg(Companies=("CIK", "nunique"),
         In_Index=("In_Any_Index", lambda s: df_full.loc[s.index][s == 1]["CIK"].nunique()),
         Not_In_Index=("In_Any_Index", lambda s: df_full.loc[s.index][s == 0]["CIK"].nunique()))
    .reset_index()
    .sort_values(["Date_Filed", "Form_Group", "Companies"], ascending=[True, True, False])
)

_ic = []
for (_dt, _fg), _h in sorted(df_full.groupby(["Date_Filed", "Form_Group"]),
                             key=lambda kv: (kv[0][0], _gkey(kv[0][1]))):
    _r = {"IDX Date": _dt, "Form Group": _fg, "Companies": _h["CIK"].nunique()}
    for _n, _c in zip(INDEX_NAMES, FLAG_COLS):
        _r[_n] = _h.loc[_h[_c] == 1, "CIK"].nunique()
    _ic.append(_r)
index_counts = pd.DataFrame(_ic)

# =============================================================================
# EXPORT
# =============================================================================
tag = "ALL" if FORM_PREFIXES is None else "-".join(
    p.replace(" ","").replace("/","") for p in FORM_PREFIXES)[:60]
period_tag = (start.strftime("%Y-%m-%d") if start == end
              else f"{start.strftime('%Y%m%d')}_to_{end.strftime('%Y%m%d')}")
OUTPUT_FILE = f"manual_report_{period_tag}_{tag}.xlsx"

with pd.ExcelWriter(OUTPUT_FILE) as w:
    table_formgroup.to_excel(w, sheet_name="FormGroup_Range", index=False)
    table_datewise.to_excel(w, sheet_name="Date_Wise", index=False)
    summary.to_excel(w, sheet_name="Summary", index=False)
    industry_breakdown.to_excel(w, sheet_name="Industry_Breakdown", index=False)
    index_counts.to_excel(w, sheet_name="Index_Counts", index=False)
    index_status.to_excel(w, sheet_name="Index_Status", index=False)
    df_full.to_excel(w, sheet_name="All_Properties", index=False)
    for ws in w.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0 for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 45)
print(f"\nReport Generated: {OUTPUT_FILE}   ({len(df_full)} rows)")

# =============================================================================
# EMAIL  (tabulated, grouped)
# =============================================================================
def summary_html(df, left=2):
    th = ("style='border:1px solid #999;padding:6px 10px;background:#1f3864;"
          "color:#fff;text-align:center;font-size:13px'")
    html = ["<table style='border-collapse:collapse;font-family:Arial,sans-serif'>",
            "<tr>" + "".join(f"<th {th}>{c}</th>" for c in df.columns) + "</tr>"]
    for i, row in df.iterrows():
        is_tot = row.iloc[0] == "TOTAL"
        bg = "background:#d9e1f2;font-weight:bold;" if is_tot else \
             ("background:#f7f7f7;" if i % 2 else "")
        cells = []
        for j, v in enumerate(row):
            al = "left" if j < left else "right"
            cells.append(f"<td style='border:1px solid #bbb;padding:5px 10px;"
                         f"text-align:{al};font-size:13px;{bg}'>{v}</td>")
        html.append("<tr>" + "".join(cells) + "</tr>")
    html.append("</table>")
    return "".join(html)

def send_report(path, recipients, ps, pe, form_prefixes,
                skipped, empty_days, n_total, n_matched):
    sender   = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")
    if not sender or not password:
        print("WARNING: SENDER_EMAIL / SENDER_APP_PASSWORD not set — "
              "report saved but NOT emailed."); return

    n502 = int(df_full["Item_5_02"].sum())
    n401 = int(df_full["Item_4_01"].sum())
    n507 = int(df_full["Item_5_07"].sum())

    form_desc = "ALL" if form_prefixes is None else ", ".join(form_prefixes)
    period_desc = ps.strftime("%Y-%m-%d") if ps == pe else f"{ps} → {pe}"

    skipped_txt = ""
    if skipped:
        skipped_txt = "\nSkipped (weekend/holiday):\n" + "\n".join(
            f"  {d} — {r}" for d, r in skipped) + "\n"
    empty_txt = ""
    if empty_days:
        empty_txt = "\nDates with no index entries:\n  " + ", ".join(empty_days) + "\n"
    warn_txt = ""
    if INDEX_WARN:
        warn_txt = ("\nNOTE: no constituents loaded for: " + ", ".join(INDEX_WARN) +
                    " — those companies are counted as 'not in index'.\n")

    group_note = ("Form groups: 10-K = 10-K, 10-KT, 10-K/A, 10-KT/A | "
                  "10-Q = 10-Q, 10-QT, 10-Q/A, 10-QT/A | 8-K = 8-K, 8-K/A | "
                  "DEF = DEF 14A, DEFC14A, DEFR14A, DEFM14A, DEFN14A, DEFC14C, "
                  "DEFR14C, DEF 14C")

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"Manual SEC Filing Report ({period_desc}) — {form_desc}"

    plain = (
        f"Manual SEC filing report.\n"
        f"Date period       : {period_desc}\n"
        f"Form filter       : {form_desc}\n"
        f"Filings scanned   : {n_total}\n"
        f"Filings in report : {n_matched}\n\n"
        f"Table 1 — Form Group (IDX range {period_desc})\n{table_formgroup.to_string(index=False)}\n\n"
        f"Table 2 — Date wise\n{table_datewise.to_string(index=False)}\n\n"
        f"{group_note}\n\n"
        f"8-K items flagged:  5.02: {n502} | 4.01: {n401} | 5.07: {n507}\n"
        f"{skipped_txt}{empty_txt}{warn_txt}\n"
        f"Per-index flags and all SEC properties are in the attached workbook."
    )
    html = (
        f"<p><b>Manual SEC filing report</b><br>"
        f"Date period: {period_desc}<br>Form filter: {form_desc}<br>"
        f"Filings scanned: {n_total} &nbsp;|&nbsp; Filings in report: {n_matched}</p>"
        f"<p><b>Table 1 — Form Group (IDX range {period_desc})</b></p>"
        f"{summary_html(table_formgroup, left=2)}"
        f"<p><b>Table 2 — Date wise</b></p>"
        f"{summary_html(table_datewise, left=1)}"
        f"<p style='font-size:12px;color:#555'>{group_note}<br>"
        f"Indices checked: {', '.join(INDEX_NAMES)}. Per-index counts are on the 'Index_Counts' sheet; the detailed date × form group table is on 'Summary'. "
        f"Counts are unique companies (CIK). 'Industries' = distinct SIC industries. "
        f"'In Indices' = member of any of the 8 indices; TOTAL rows are column sums.</p>"
        f"<p>8-K items flagged: Item 5.02: {n502} | Item 4.01: {n401} | Item 5.07: {n507}</p>"
        + (f"<p>Skipped (weekend/holiday): {'; '.join(f'{d} — {r}' for d, r in skipped)}</p>" if skipped else "")
        + (f"<p>Dates with no index entries: {', '.join(empty_days)}</p>" if empty_days else "")
        + (f"<p style='color:#b00'>NOTE: no constituents loaded for {', '.join(INDEX_WARN)}; "
           f"those companies count as 'not in index'.</p>" if INDEX_WARN else "")
        + "<p>Index flags and all SEC properties are on the attached workbook.</p>"
    )
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")
    with open(path, "rb") as f:
        msg.add_attachment(
            f.read(), maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(path))
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls(); s.login(sender, password); s.send_message(msg)
    print(f"Email sent to: {', '.join(recipients)}")

send_report(OUTPUT_FILE, RECIPIENTS, start, end, FORM_PREFIXES,
            skipped, empty_days, sum(idx_totals.values()), len(df_full))
