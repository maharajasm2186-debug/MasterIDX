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
# • Fund companies are flagged (Is_Fund) and counted in a separate column
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
# FUND COMPANY DETECTION   (Is_Fund = 1)
# -----------------------------------------------------------------------------
# A filer is a fund when ANY of these is true (checked in this order):
#   1. SEC entity type is "investment"  (registered investment companies)
#   2. SIC is a pure fund code (default 6726, 6722)
#   3. Name contains FUND / FUNDS / ETF / COMMODITY POOL
#   4. Name has a fund sponsor word (iShares, SPDR, ProShares, WisdomTree, VanEck,
#      Grayscale, Teucrium) AND the word TRUST   -> "WisdomTree, Inc." is NOT a fund
#   5. SIC is an AMBIGUOUS code (default 6221, commodity brokers) AND the name also
#      looks like a fund (FUND/ETF/TRUST/POOL/FUTURES/PORTFOLIO/L.P.)
#      -> an operating "xyz Inc." with SIC 6221 is NOT a fund
# Overridable in manual_config.ini: fund_sic_codes, fund_sic_codes_need_name,
# fund_name_regex
# =============================================================================
FUND_SICS = {x.strip() for x in section.get("fund_sic_codes", "6726,6722").split(",") if x.strip()}
FUND_SICS_NEED_NAME = {x.strip() for x in section.get("fund_sic_codes_need_name", "6221").split(",") if x.strip()}
FUND_NAME_RX = re.compile(section.get("fund_name_regex", r"\b(FUNDS?|ETFS?|COMMODITY POOL)\b"), re.I)
FUND_SPONSOR_RX = re.compile(
    r"\b(ISHARES|SPDR|PROSHARES|WISDOMTREE|VANECK|GRAYSCALE|TEUCRIUM)\b.*\bTRUST\b", re.I)
FUND_HINT_RX = re.compile(r"\b(FUNDS?|ETFS?|TRUST|POOL|FUTURES|PORTFOLIO)\b|L\.?P\.?$", re.I)

def is_fund(row):
    sic  = str(row.get("SIC") or "").split(".")[0].strip()
    name = f"{row.get('Company') or ''} {row.get('Company_Name_API') or ''}".strip()
    if str(row.get("Entity_Type") or "").strip().lower() == "investment": return 1
    if sic in FUND_SICS: return 1
    if FUND_NAME_RX.search(name): return 1
    if FUND_SPONSOR_RX.search(name): return 1
    if sic in FUND_SICS_NEED_NAME and FUND_HINT_RX.search(name): return 1
    return 0

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

# ---------- Fund company flag ----------
df_full["Is_Fund"] = df_full.apply(is_fund, axis=1)
df_full["Industry"] = df_full["Industry"].fillna("Unknown")

ORDER = [
    "Date_Filed","CIK","Company","Company_Name_API","Tickers",
    "Form_Type","Form_Group","Is_Fund","Report_Date",
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
    "Fund Companies",
    "Industries",
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
            rows.append([dt, fg, h["CIK"].nunique(),
                         h.loc[h["Is_Fund"] == 1, "CIK"].nunique(),
                         h["Industry"].nunique(), totals.get(dt, 0)])
    out = pd.DataFrame(rows, columns=COLS)
    dates_present = out["IDX Date"].unique()
    total = ["TOTAL", "",
             int(out[COLS[2]].sum()), int(out[COLS[3]].sum()),
             int(out[COLS[4]].sum()),
             int(sum(totals.get(d, 0) for d in dates_present))]
    out.loc[len(out)] = total
    return out

summary = build_summary(df_full, idx_totals)

# ---------- Mail table 1: Form Group over the whole IDX date range ----------
PERIOD_DESC = (start.strftime("%Y-%m-%d") if start == end else f"{start} → {end}")
T1_COLS = ["Form Group", "Range of IDX Date", "Companies Released", "Fund Companies"]
T2_COLS = ["IDX Date", "Companies Released", "Fund Companies"]

def _counts(h):
    return [h["CIK"].nunique(),
            h.loc[h["Is_Fund"] == 1, "CIK"].nunique()]

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
         Fund_Companies=("Is_Fund", lambda s: df_full.loc[s.index][s == 1]["CIK"].nunique()))
    .reset_index()
    .sort_values(["Date_Filed", "Form_Group", "Companies"], ascending=[True, True, False])
)

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
        f"{skipped_txt}{empty_txt}\n"
        f"All SEC properties (incl. Is_Fund flag) are in the attached workbook."
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
        f"The detailed date × form group table is on the 'Summary' sheet. "
        f"Counts are unique companies (CIK). 'Industries' = distinct SIC industries. "
        f"'Fund Companies' = funds/ETFs/commodity pools (SIC or fund-name match), already included in 'Companies Released'; TOTAL rows are column sums.</p>"
        f"<p>8-K items flagged: Item 5.02: {n502} | Item 4.01: {n401} | Item 5.07: {n507}</p>"
        + (f"<p>Skipped (weekend/holiday): {'; '.join(f'{d} — {r}' for d, r in skipped)}</p>" if skipped else "")
        + (f"<p>Dates with no index entries: {', '.join(empty_days)}</p>" if empty_days else "")
        + "<p>All SEC properties (incl. Is_Fund flag) are on the attached workbook.</p>"
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
