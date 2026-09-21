import os, sys, smtplib, requests, pandas as pd
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
# • Runs ONLY when you invoke:   python manual_report.py
# • All requirements live in manual_config.ini (form types + date period).
# • Completely independent of the daily cron script — the daily job is
#   untouched and this file never modifies it.
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
    "auditfeeteam@gmail.com", "santhakumarcu@gmail.com",
    "maharaja@secanalyzer.net", "chandru@secanalyzer.net",
]

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
        "SIC","Industry","State_of_Incorporation","Fiscal_Year_End",
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

ORDER = [
    "Date_Filed","CIK","Company","Company_Name_API","Form_Type","Report_Date",
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
# SUMMARY + EXPORT
# =============================================================================
summary = (df_full.groupby(["Date_Filed","Form_Type","Industry",
                            "Item_5_02","Item_4_01","Item_5_07"])
                  .size().reset_index(name="Count")
                  .sort_values(["Date_Filed","Form_Type","Count"],
                               ascending=[True,True,False]))

tag = "ALL" if FORM_PREFIXES is None else "-".join(
    p.replace(" ","").replace("/","") for p in FORM_PREFIXES)[:60]
period_tag = (start.strftime("%Y-%m-%d") if start == end
              else f"{start.strftime('%Y%m%d')}_to_{end.strftime('%Y%m%d')}")
OUTPUT_FILE = f"manual_report_{period_tag}_{tag}.xlsx"

with pd.ExcelWriter(OUTPUT_FILE) as w:
    summary.to_excel(w, sheet_name="Summary", index=False)
    df_full.to_excel(w, sheet_name="All_Properties", index=False)
print(f"\nReport Generated: {OUTPUT_FILE}   ({len(df_full)} rows)")

# =============================================================================
# EMAIL (same recipients from config)
# =============================================================================
def send_report(path, recipients, ps, pe, form_prefixes,
                skipped, empty_days, n_total, n_matched):
    sender   = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")
    if not sender or not password:
        print("WARNING: SENDER_EMAIL / SENDER_APP_PASSWORD not set — "
              "report saved but NOT emailed."); return

    ft = df_full["Form_Type"].astype(str)
    by_form = ft.value_counts().head(25).to_string()
    by_date = df_full.groupby("Date_Filed").size().sort_index().to_string()

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

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"Manual SEC Filing Report ({period_desc}) — {form_desc}"
    msg.set_content(
        f"Manual SEC filing report.\n"
        f"Date period       : {period_desc}\n"
        f"Form filter       : {form_desc}\n\n"
        f"Filings scanned   : {n_total}\n"
        f"Filings in report : {n_matched}\n\n"
        f"Filings per date:\n{by_date}\n\n"
        f"Top form types:\n{by_form}\n\n"
        f"8-K items flagged:\n"
        f"  Item 5.02 : {n502}\n  Item 4.01 : {n401}\n  Item 5.07 : {n507}\n"
        f"{skipped_txt}{empty_txt}\n"
        f"All SEC properties are on the 'All_Properties' sheet."
    )
    with open(path, "rb") as f:
        msg.add_attachment(
            f.read(), maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(path))
    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls(); s.login(sender, password); s.send_message(msg)
    print(f"Email sent to: {', '.join(recipients)}")

send_report(OUTPUT_FILE, RECIPIENTS, start, end, FORM_PREFIXES,
            skipped, empty_days, len(df_filings), len(df_full))