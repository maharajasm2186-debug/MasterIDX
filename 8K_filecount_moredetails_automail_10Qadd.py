import os
import sys
import smtplib
import requests
import pandas as pd
from io import StringIO
from datetime import datetime, date, timedelta
from email.message import EmailMessage
import time
import re
import threading
from concurrent.futures import ThreadPoolExecutor
import holidays as us_holidays

# =============================================================================
# DAILY AUTOMATION + EMAIL  (format aligned with manual_onrequirement.py)
# =============================================================================

DAYS_BACK = 1

if len(sys.argv) > 1:
    input_date = sys.argv[1]
else:
    input_date = (date.today() - timedelta(days=DAYS_BACK)).strftime("%Y-%m-%d")

dt = datetime.strptime(input_date, "%Y-%m-%d")

RECIPIENTS = ["maharaja@secanalyzer.net"]
RECIPIENTS = ["lawrenceamal@gmail.com"]
RECIPIENTS = ["auditfeeteam@gmail.com"]
RECIPIENTS = ["chandru@secanalyzer.net"]
RECIPIENTS = ["santhakumarcu@gmail.com"]

# =============================================================================
# WEEKEND / SEC HOLIDAY GUARD
# =============================================================================
def send_skip_notification(report_date, reason):
    sender = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")
    if not sender or not password:
        print(f"WARNING: SENDER_EMAIL / SENDER_APP_PASSWORD not set. "
              f"Skipping notification email for: {reason}")
        return
    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(RECIPIENTS)
    msg["Subject"] = f"Daily Filing Report ({report_date}) — No Report Today"
    msg.set_content(
        f"No EDGAR filing report was generated for {report_date}.\n\n"
        f"Reason: {reason}\n\n"
        f"The SEC does not process filings on weekends or federal holidays.\n"
        f"The next report will cover the next regular trading day."
    )
    try:
        with smtplib.SMTP("smtp.gmail.com", 587) as server:
            server.starttls()
            server.login(sender, password)
            server.send_message(msg)
        print(f"Skip notification sent to: {', '.join(RECIPIENTS)}")
    except Exception as e:
        print(f"WARNING: Failed to send skip notification email: {e}")


if dt.weekday() == 5:
    reason = "Saturday — EDGAR is closed on weekends"
    print(f"Skipping {input_date}: {reason}")
    send_skip_notification(input_date, reason); sys.exit(0)
if dt.weekday() == 6:
    reason = "Sunday — EDGAR is closed on weekends"
    print(f"Skipping {input_date}: {reason}")
    send_skip_notification(input_date, reason); sys.exit(0)

fed_holidays = us_holidays.UnitedStates(years=dt.year)
if dt.date() in fed_holidays:
    reason = f"US Federal Holiday — {fed_holidays[dt.date()]}"
    print(f"Skipping {input_date}: {reason}")
    send_skip_notification(input_date, reason); sys.exit(0)

# =============================================================================
# NORMAL PROCESSING
# =============================================================================
OUTPUT_FILE = f"filing_report_{input_date}.xlsx"
headers = {"User-Agent": "SEC Research your_email@example.com"}

MAX_WORKERS = 10
MAX_BYTES   = 250_000
REQ_PER_SEC = 9
TIMEOUT     = 30

session = requests.Session()
session.headers.update(headers)


class RateLimiter:
    def __init__(self, max_per_sec):
        self.min_interval = 1.0 / max_per_sec
        self.lock = threading.Lock()
        self.next_time = time.monotonic()

    def wait(self):
        with self.lock:
            now = time.monotonic()
            sleep_for = max(0.0, self.next_time - now)
            self.next_time = max(now, self.next_time) + self.min_interval
        if sleep_for > 0:
            time.sleep(sleep_for)


limiter = RateLimiter(REQ_PER_SEC)

year    = dt.year
quarter = (dt.month - 1) // 3 + 1


def _parse_idx(text):
    lines = text.split("\n")
    data_lines = [l for l in lines if l.count("|") >= 4]
    df_raw = pd.read_csv(
        StringIO("\n".join(data_lines)), sep="|",
        names=["CIK", "Company", "Form_Type", "Date_Filed", "File_Name"],
        dtype=str,
    )
    for col in ["CIK", "Company", "Form_Type", "Date_Filed", "File_Name"]:
        df_raw[col] = df_raw[col].astype(str).str.strip()
    return df_raw


# -----------------------------------------------------------------------------
# INDEX FETCH (daily first, quarterly fallback, merge both)
# -----------------------------------------------------------------------------
date_nodash = input_date.replace("-", "")
daily_url = (f"https://www.sec.gov/Archives/edgar/daily-index/"
             f"{year}/QTR{quarter}/master{date_nodash}.idx")
quarterly_url = (f"https://www.sec.gov/Archives/edgar/full-index/"
                 f"{year}/QTR{quarter}/master.idx")

daily_resp = session.get(daily_url, timeout=TIMEOUT)
if daily_resp.status_code == 200 and "|" in daily_resp.text:
    df_daily = _parse_idx(daily_resp.text)
    df_daily = df_daily[df_daily["Date_Filed"] == input_date]
    print(f"✓ Daily index loaded  → {len(df_daily)} filings for {input_date}")

    quarterly_resp = session.get(quarterly_url, timeout=TIMEOUT)
    df_quarterly = _parse_idx(quarterly_resp.text)
    df_quarterly = df_quarterly[df_quarterly["Date_Filed"] == input_date]
    print(f"✓ Quarterly idx loaded → {len(df_quarterly)} filings for {input_date}")

    df = (pd.concat([df_daily, df_quarterly], ignore_index=True)
            .drop_duplicates(subset=["File_Name"])
            .reset_index(drop=True))
    print(f"✓ Combined total       → {len(df)} unique filings")
else:
    print(f"Daily index not available for {input_date}, using quarterly master.idx")
    quarterly_resp = session.get(quarterly_url, timeout=TIMEOUT)
    df = _parse_idx(quarterly_resp.text)
    df = df[df["Date_Filed"] == input_date].reset_index(drop=True)
    print(f"✓ Quarterly idx loaded → {len(df)} filings for {input_date}")

# =============================================================================
# FORM FILTER  (8-K / 10-K / 10-Q / DEF* / 20-F / 40-F)
# =============================================================================
FORM_PREFIXES = (
    "8-K", "10-K", "10-Q",
    "DEF 14", "DEFC14", "DEFM14", "DEFR14", "DEFN14",
    "20-F", "40-F",
)

def _matches_prefix(form_type):
    return any(form_type.startswith(p) for p in FORM_PREFIXES)

df_filings = df[df["Form_Type"].apply(_matches_prefix)].copy()

print(f"\nFilings found for {input_date}:")
print(df_filings["Form_Type"].value_counts().to_string())
print(f"Total: {len(df_filings)}\n")

df_filings["Accession_No"] = df_filings["File_Name"].apply(
    lambda x: x.split("/")[-1].replace(".txt", "")
)

# =============================================================================
# CAPPED DOWNLOAD (item detection for 8-Ks only)
# =============================================================================
def fetch_capped(full_url, max_bytes=MAX_BYTES):
    limiter.wait()
    with session.get(full_url, stream=True, timeout=TIMEOUT) as r:
        chunks, total = [], 0
        for chunk in r.iter_content(chunk_size=16384):
            if not chunk: continue
            chunks.append(chunk); total += len(chunk)
            if total >= max_bytes: break
    return b"".join(chunks).decode("utf-8", errors="ignore")


def detect_items(text):
    """8-K item flags from the SEC header or body captions."""
    low = text.lower().replace("\xa0", " ")
    header_titles = " ".join(re.findall(r'item information:\s*(.+)', low))
    i502 = 1 if "departure of directors" in header_titles else 0
    i401 = 1 if "certifying accountant" in header_titles else 0
    i507 = 1 if "submission of matters to a vote of security holders" in header_titles else 0
    if not i502 and re.search(r'item\s*5\.02', low): i502 = 1
    if not i401 and re.search(r'item\s*4\.01', low): i401 = 1
    if not i507 and re.search(r'item\s*5\.07', low): i507 = 1
    return i502, i401, i507


_progress = {"done": 0}
_progress_lock = threading.Lock()
_total = len(df_filings)


def process_filing(item):
    idx, file = item
    form_type = str(df_filings.at[idx, "Form_Type"])

    # Non-8-K forms don't carry 8-K items → skip the fetch entirely.
    if not form_type.startswith("8-K"):
        with _progress_lock:
            _progress["done"] += 1
            if _progress["done"] % 25 == 0 or _progress["done"] == _total:
                print(f"  processed {_progress['done']}/{_total} filings")
        return idx, 0, 0, 0

    full_url = "https://www.sec.gov/Archives/" + file
    try:
        text = fetch_capped(full_url)
        i502, i401, i507 = detect_items(text)
    except Exception:
        i502 = i401 = i507 = 0

    with _progress_lock:
        _progress["done"] += 1
        if _progress["done"] % 25 == 0 or _progress["done"] == _total:
            print(f"  processed {_progress['done']}/{_total} filings")
    return idx, i502, i401, i507


print(f"Scanning {_total} filings for {input_date} ...")
results = {}
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
    for idx, i502, i401, i507 in ex.map(process_filing,
                                        list(df_filings["File_Name"].items())):
        results[idx] = (i502, i401, i507)

df_filings["Item_5_02"] = [results[i][0] for i in df_filings.index]
df_filings["Item_4_01"] = [results[i][1] for i in df_filings.index]
df_filings["Item_5_07"] = [results[i][2] for i in df_filings.index]

# =============================================================================
# FILTER RELEVANT
# =============================================================================
is_8k = df_filings["Form_Type"].str.startswith("8-K")
has_flag = ((df_filings["Item_5_02"] == 1) |
            (df_filings["Item_4_01"] == 1) |
            (df_filings["Item_5_07"] == 1))
filtered_df = df_filings[(is_8k & has_flag) | (~is_8k)].copy()
print(f"Filings in report: {len(filtered_df)}\n")

# =============================================================================
# FORM GROUPING  (variants → one group)
# =============================================================================
GROUP_RULES = [
    (re.compile(r"^10-K"),              "10-K"),
    (re.compile(r"^10-Q"),              "10-Q"),
    (re.compile(r"^8-K"),               "8-K"),
    (re.compile(r"^DEF[A-Z]? ?14[AC]"), "DEF"),
    (re.compile(r"^20-F"),              "20-F"),
    (re.compile(r"^40-F"),              "40-F"),
]
GROUP_ORDER = ["8-K", "10-K", "10-Q", "DEF", "20-F", "40-F"]

def form_group(ft):
    s = str(ft).strip().upper()
    for rx, g in GROUP_RULES:
        if rx.match(s):
            return g
    return s

# =============================================================================
# FUND DETECTION
# =============================================================================
FUND_SICS            = {"6726", "6722"}
FUND_SICS_NEED_NAME  = {"6221"}
FUND_NAME_RX    = re.compile(r"\b(FUNDS?|ETFS?|COMMODITY POOL)\b", re.I)
FUND_SPONSOR_RX = re.compile(
    r"\b(ISHARES|SPDR|PROSHARES|WISDOMTREE|VANECK|GRAYSCALE|TEUCRIUM)\b.*\bTRUST\b", re.I)
FUND_HINT_RX = re.compile(
    r"\b(FUNDS?|ETFS?|TRUST|POOL|FUTURES|PORTFOLIO)\b|L\.?P\.?$", re.I)

def is_fund(row):
    sic  = str(row.get("SIC") or "").split(".")[0].strip()
    name = f"{row.get('Company') or ''} {row.get('Company_Name_API') or ''}".strip()
    if str(row.get("Entity_Type") or "").strip().lower() == "investment": return 1
    if sic in FUND_SICS:            return 1
    if FUND_NAME_RX.search(name):   return 1
    if FUND_SPONSOR_RX.search(name): return 1
    if sic in FUND_SICS_NEED_NAME and FUND_HINT_RX.search(name): return 1
    return 0

# =============================================================================
# FULL SEC ENRICHMENT  (matches manual_onrequirement.py All_Properties)
# =============================================================================
RENAME = {
    "reportDate":          "Report_Date",
    "acceptanceDateTime":  "Acceptance_DateTime",
    "act":                 "Act",
    "fileNumber":          "File_Number",
    "filmNumber":          "Film_Number",
    "items":               "Items",
    "size":                "Size",
    "isXBRL":              "Is_XBRL",
    "isInlineXBRL":        "Is_Inline_XBRL",
    "primaryDocument":     "Primary_Document",
    "primaryDocDescription":"Primary_Doc_Description",
}

def fetch_cik(cik_raw):
    cik10 = str(cik_raw).zfill(10)
    api = f"https://data.sec.gov/submissions/CIK{cik10}.json"
    limiter.wait()
    try:
        r = session.get(api, timeout=TIMEOUT)
        data = r.json()
    except Exception as e:
        print(f"  ! submissions failed for {cik_raw}: {e}")
        return {}, {}

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
        "Business_City": a.get("city"),       "Business_State": a.get("stateOrCountry"),
        "Business_State_Desc": a.get("stateOrCountryDescription"),
        "Business_Zip": a.get("zipCode"),
        "Mailing_Street": m.get("street1"),   "Mailing_Street2": m.get("street2"),
        "Mailing_City": m.get("city"),        "Mailing_State": m.get("stateOrCountry"),
        "Mailing_Zip": m.get("zipCode"),
        "Former_Names": "; ".join(
            fn.get("name", "") for fn in (data.get("formerNames") or [])
        ),
    }

    recent = (data.get("filings") or {}).get("recent") or {}
    accs = recent.get("accessionNumber") or []
    n = len(accs)
    fmap = {}
    for i, acc in enumerate(accs):
        key = acc.replace("-", "")
        fmap[key] = {new: (recent.get(old) or [None]*n)[i]
                     for old, new in RENAME.items()}
    return company, fmap


filtered_df["_acc_key"]    = filtered_df["Accession_No"].str.replace("-", "", regex=False)
filtered_df["Form_Group"]  = filtered_df["Form_Type"].apply(form_group)
filtered_df["Filing_URL"]  = "https://www.sec.gov/Archives/" + filtered_df["File_Name"]

unique_ciks = filtered_df["CIK"].unique().tolist()
print(f"Fetching SEC properties for {len(unique_ciks)} unique companies ...")

cik_company, cik_filings = {}, {}
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
    for cik, (c, f) in zip(unique_ciks, ex.map(fetch_cik, unique_ciks)):
        cik_company[cik] = c
        cik_filings[cik] = f

def _enrich(row):
    out = {}
    out.update(cik_company.get(row["CIK"], {}))
    out.update(cik_filings.get(row["CIK"], {}).get(row["_acc_key"], {}))
    return out

filtered_df = filtered_df.reset_index(drop=True)
enriched = filtered_df.apply(_enrich, axis=1, result_type="expand")
df_full  = pd.concat([filtered_df, enriched], axis=1).drop(columns=["_acc_key"])

# --- ensure every expected column exists ---
for c in list(RENAME.values()) + [
    "Tickers","SIC","Industry","State_of_Incorporation",
    "State_of_Incorporation_Desc","Fiscal_Year_End","Entity_Type","Phone",
    "Former_Names","Business_Street","Business_Street2","Business_City",
    "Business_State","Business_State_Desc","Business_Zip",
    "Mailing_Street","Mailing_Street2","Mailing_City","Mailing_State","Mailing_Zip",
]:
    if c not in df_full.columns:
        df_full[c] = None

df_full["Industry"] = df_full["Industry"].fillna("Unknown")

# --- item flag safety net: force int, OR with API Items if present ---
for col, code in [("Item_5_02","5.02"), ("Item_4_01","4.01"), ("Item_5_07","5.07")]:
    df_full[col] = pd.to_numeric(df_full[col], errors="coerce").fillna(0).astype(int)
    if "Items" in df_full.columns:
        # Pre-build the regex pattern to avoid a backslash inside the f-string
        pattern = r"\b" + re.escape(code) + r"\b"
        
        api_hit = (df_full["Items"].fillna("").astype(str)
                     .str.contains(pattern, regex=True)
                     .astype(int))
        mismatch = int(((df_full[col] == 0) & (api_hit == 1)).sum())
        if mismatch:
            print(f"  note: {mismatch} filing(s) have {code} in API 'Items' but not in text")
        df_full[col] = ((df_full[col] == 1) | (api_hit == 1)).astype(int)

# Non-8-K → all 8-K item flags 0 (mirrors manual report)
non_8k = df_full["Form_Group"] != "8-K"
df_full.loc[non_8k, ["Item_5_02", "Item_4_01", "Item_5_07"]] = 0

# --- fund flag ---
df_full["Is_Fund"] = (
    df_full.apply(lambda r: is_fund(r), axis=1).astype(int)
    if not df_full.empty else pd.Series(dtype=int)
)

# --- accepted dates ---
_acc    = pd.to_datetime(df_full["Acceptance_DateTime"], utc=True, errors="coerce")
_acc_et = _acc.dt.tz_convert("America/New_York")
df_full["Accepted_UTC"]  = _acc.dt.strftime("%Y-%m-%d %H:%M:%S")
df_full["Accepted_ET"]   = _acc_et.dt.strftime("%Y-%m-%d %H:%M:%S")
df_full["Accepted_Date"] = _acc.dt.strftime("%Y-%m-%d")
df_full["Filed_vs_Accepted_Diff"] = (
    (df_full["Accepted_Date"].notna()) &
    (df_full["Accepted_Date"] != df_full["Date_Filed"])
).astype(int)

# --- final column order (same as manual report) ---
ORDER = [
    "Date_Filed","CIK","Company","Company_Name_API","Tickers",
    "Form_Type","Form_Group","Is_Fund","Report_Date",
    "Accession_No","Filing_URL","Acceptance_DateTime","Accepted_UTC","Accepted_ET",
    "Accepted_Date","Filed_vs_Accepted_Diff",
    "Act","File_Number","Film_Number","Size","Is_XBRL","Is_Inline_XBRL",
    "Primary_Document","Primary_Doc_Description","Items",
    "Item_5_02","Item_4_01","Item_5_07",
    "SIC","Industry","State_of_Incorporation","State_of_Incorporation_Desc",
    "Fiscal_Year_End","Entity_Type","Phone","Former_Names",
    "Business_Street","Business_Street2","Business_City","Business_State",
    "Business_State_Desc","Business_Zip",
    "Mailing_Street","Mailing_Street2","Mailing_City","Mailing_State","Mailing_Zip",
]
for c in ORDER:
    if c not in df_full.columns:
        df_full[c] = None
df_full = df_full[ORDER]

# =============================================================================
# SUMMARY TABLES  (identical to manual format)
# =============================================================================
def _gkey(g):
    return (GROUP_ORDER.index(g) if g in GROUP_ORDER else len(GROUP_ORDER), g)

ITEM_FLAGS = [("Item 5.02", "Item_5_02"),
              ("Item 5.07", "Item_5_07"),
              ("Item 4.01", "Item_4_01")]

def build_formgroup_table(df, period):
    rows, base = [], []
    for fg in sorted(df["Form_Group"].unique(), key=_gkey):
        h = df[df["Form_Group"] == fg]
        rows.append([fg, period, h["CIK"].nunique(),
                     h.loc[h["Is_Fund"] == 1, "CIK"].nunique(),
                     h["Industry"].nunique(), len(h)])
        base.append(len(rows) - 1)
        if fg == "8-K":
            for name, flag in ITEM_FLAGS:
                hh = h[h[flag] == 1]
                rows.append([f"8-K {name}", period, hh["CIK"].nunique(),
                             hh.loc[hh["Is_Fund"] == 1, "CIK"].nunique(),
                             hh["Industry"].nunique(), len(hh)])
    out = pd.DataFrame(rows, columns=[
        "Form Group", "IDX Date", "Companies Released",
        "Fund Companies", "Industries", "Filings",
    ])
    if not out.empty:
        out.loc[len(out)] = ["TOTAL", period,
                             int(out.loc[base, "Companies Released"].sum()),
                             int(out.loc[base, "Fund Companies"].sum()),
                             int(out.loc[base, "Industries"].sum()),
                             int(out.loc[base, "Filings"].sum())]
    return out

table_formgroup = build_formgroup_table(df_full, input_date)

if df_full.empty:
    industry_breakdown = pd.DataFrame(
        columns=["Form_Group","Industry","Companies","Fund_Companies","Filings"])
else:
    industry_breakdown = (
        df_full.groupby(["Form_Group", "Industry"])
        .agg(Companies=("CIK", "nunique"),
             Fund_Companies=("Is_Fund",
                              lambda s: df_full.loc[s.index][s == 1]["CIK"].nunique()),
             Filings=("Accession_No", "nunique"))
        .reset_index()
        .sort_values(["Form_Group", "Companies"], ascending=[True, False])
    )

# =============================================================================
# EXPORT TO EXCEL
# =============================================================================
with pd.ExcelWriter(OUTPUT_FILE) as w:
    table_formgroup.to_excel(w, sheet_name="FormGroup_Summary", index=False)
    industry_breakdown.to_excel(w, sheet_name="Industry_Breakdown", index=False)
    df_full.to_excel(w, sheet_name="All_Properties", index=False)
    for ws in w.book.worksheets:
        ws.freeze_panes = "A2"
        for col in ws.columns:
            width = max(len(str(c.value)) if c.value is not None else 0
                        for c in col[:200])
            ws.column_dimensions[col[0].column_letter].width = min(max(10, width + 2), 45)

print(f"Report Generated: {OUTPUT_FILE}   ({len(df_full)} rows)")

# =============================================================================
# EMAIL  (HTML tables matching manual format)
# =============================================================================
def summary_html(df, left=2):
    if df.empty:
        return "<p>No filings found.</p>"
    th = ("style='border:1px solid #999;padding:6px 10px;background:#1f3864;"
          "color:#fff;text-align:center;font-size:13px'")
    html = ["<table style='border-collapse:collapse;font-family:Arial,sans-serif'>",
            "<tr>" + "".join(f"<th {th}>{c}</th>" for c in df.columns) + "</tr>"]
    for i, row in df.iterrows():
        is_tot = row.iloc[0] == "TOTAL"
        is_sub = str(row.iloc[0]).startswith("8-K Item")
        bg = ("background:#d9e1f2;font-weight:bold;" if is_tot
              else ("background:#fff8e1;font-style:italic;" if is_sub
                    else ("background:#f7f7f7;" if i % 2 else "")))
        cells = []
        for j, v in enumerate(row):
            al = "left" if j < left else "right"
            cells.append(f"<td style='border:1px solid #bbb;padding:5px 10px;"
                         f"text-align:{al};font-size:13px;{bg}'>{v}</td>")
        html.append("<tr>" + "".join(cells) + "</tr>")
    html.append("</table>")
    return "".join(html)


def send_report(path, recipients, report_date, n_total, n_matched):
    sender   = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")
    if not sender or not password:
        print("WARNING: SENDER_EMAIL / SENDER_APP_PASSWORD not set. "
              "Report saved but NOT emailed."); return

    n502 = int(df_full["Item_5_02"].sum()) if len(df_full) else 0
    n401 = int(df_full["Item_4_01"].sum()) if len(df_full) else 0
    n507 = int(df_full["Item_5_07"].sum()) if len(df_full) else 0

    group_note = ("Form groups: 10-K = 10-K, 10-KT, 10-K/A, 10-KT/A | "
                  "10-Q = 10-Q, 10-QT, 10-Q/A, 10-QT/A | 8-K = 8-K, 8-K/A | "
                  "DEF = DEF 14A, DEFC14A, DEFR14A, DEFM14A, DEFN14A, "
                  "DEFC14C, DEFR14C, DEF 14C")

    plain = (
        f"Daily SEC filing report.\n"
        f"Report date       : {report_date}\n"
        f"Filings scanned   : {n_total}\n"
        f"Filings in report : {n_matched}\n\n"
        f"Form group summary:\n{table_formgroup.to_string(index=False)}\n\n"
        f"8-K items flagged:  5.02: {n502} | 4.01: {n401} | 5.07: {n507}\n\n"
        f"All SEC properties are in the attached workbook "
        f"(sheet 'All_Properties')."
    )

    html = (
        f"<p><b>Daily SEC filing report</b><br>"
        f"Report date: {report_date}<br>"
        f"Filings scanned: {n_total} &nbsp;|&nbsp; "
        f"Filings in report: {n_matched}</p>"

        f"<p><b>Form Group Summary ({report_date})</b></p>"
        f"{summary_html(table_formgroup, left=2)}"

        f"<p style='font-size:12px;color:#555'>{group_note}<br>"
        f"Counts are unique companies (CIK). 'Industries' = distinct SIC industries. "
        f"'8-K Item …' rows = companies whose 8-K reported that item "
        f"(a filing can report several, so they overlap and are not added into TOTAL). "
        f"'Fund Companies' = funds/ETFs/commodity pools (SIC or fund-name match), "
        f"already included in 'Companies Released'. "
        f"TOTAL row is a column sum across main form groups.</p>"

        f"<p>8-K items flagged: Item 5.02: {n502} | Item 4.01: {n401} | Item 5.07: {n507}</p>"

        f"<p>All SEC properties (incl. Is_Fund flag) are on the attached workbook "
        f"(sheet 'All_Properties'). The 'Filings' column in this email equals the "
        f"row count in All_Properties for that form group — the counts are consistent.</p>"
    )

    msg = EmailMessage()
    msg["From"]    = sender
    msg["To"]      = ", ".join(recipients)
    msg["Subject"] = f"Daily Filing Report ({report_date})"
    msg.set_content(plain)
    msg.add_alternative(html, subtype="html")

    with open(path, "rb") as f:
        msg.add_attachment(
            f.read(),
            maintype="application",
            subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            filename=os.path.basename(path),
        )

    with smtplib.SMTP("smtp.gmail.com", 587) as s:
        s.starttls()
        s.login(sender, password)
        s.send_message(msg)
    print(f"Email sent to: {', '.join(recipients)}")


send_report(OUTPUT_FILE, RECIPIENTS, input_date,
            n_total=len(df_filings), n_matched=len(df_full))
