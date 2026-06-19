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

# =============================================================================
# DAILY AUTOMATION + EMAIL
# =============================================================================
# This script is meant to be scheduled (cron / Task Scheduler) to run once a
# day. It picks the date automatically, builds the report, and emails the
# xlsx to the recipients below.
#
# DATE SELECTION (no manual editing needed):
#   - default: YESTERDAY's date (today usually isn't fully indexed on EDGAR yet)
#   - or pass a date explicitly:   python script.py 2026-03-23
#   - DAYS_BACK controls the offset: 1 = yesterday (default), 0 = today,
#     2 = two days back, etc.
#
# EMAIL CREDENTIALS (never hardcoded):
#   The script reads the sender account from environment variables so no
#   password is stored in this file:
#       SENDER_EMAIL          your Gmail address
#       SENDER_APP_PASSWORD   a Gmail *App Password* (16 chars, requires 2FA)
#   Create an App Password at: https://myaccount.google.com/apppasswords
#   Set them once in your shell profile / scheduler environment, e.g.:
#       export SENDER_EMAIL="you@gmail.com"
#       export SENDER_APP_PASSWORD="your_app_password_here"
# =============================================================================

DAYS_BACK = 1  # process YESTERDAY by default (EDGAR isn't fully indexed for today yet)

if len(sys.argv) > 1:
    input_date = sys.argv[1]                      # explicit override: YYYY-MM-DD
else:
    input_date = (date.today() - timedelta(days=DAYS_BACK)).strftime("%Y-%m-%d")

# Validate format early so a bad argument fails clearly
datetime.strptime(input_date, "%Y-%m-%d")

# Recipients for the daily report
RECIPIENTS = [
    "auditfeeteam@gmail.com",
    "santhakumarcu@gmail.com",
    "maharajasm2186@gmail.com",
]

# Date-stamped output file so each day's report is kept separately
OUTPUT_FILE = f"filing_report_{input_date}.xlsx"

headers = {
    "User-Agent": "SEC Research your_email@example.com"
}

# =========================
# PERFORMANCE SETTINGS
# =========================
MAX_WORKERS = 10        # concurrent downloads
MAX_BYTES = 250_000     # only read the first ~250 KB of each submission
REQ_PER_SEC = 9         # stay safely under SEC's 10 req/sec limit
TIMEOUT = 30

# Shared session = reused TCP connections (faster than a fresh connect each time)
session = requests.Session()
session.headers.update(headers)


# =========================
# GLOBAL RATE LIMITER
# =========================
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


# =========================
# YEAR & QUARTER
# =========================
dt = datetime.strptime(input_date, "%Y-%m-%d")
year = dt.year
quarter = (dt.month - 1) // 3 + 1

url = f"https://www.sec.gov/Archives/edgar/full-index/{year}/QTR{quarter}/master.idx"

response = session.get(url, timeout=TIMEOUT)
content = response.text

# Remove header
data = "\n".join(content.split("\n")[11:])

df = pd.read_csv(
    StringIO(data),
    sep="|",
    names=["CIK","Company","Form_Type","Date_Filed","File_Name"]
)

# =========================
# FILTER TARGET FORM TYPES
# =========================
# Match whole form *families* by prefix (the trailing "*" means variants are
# included automatically):
#   8-K*    -> 8-K, 8-K/A, 8-K12B, ...
#   10-K*   -> 10-K, 10-K/A, 10-KSB, 10-K405, ...
#   10-Q*   -> 10-Q, 10-Q/A, 10-QSB, ...          ← NEW
#   DEF 14* -> DEF 14A, DEF 14C            (note the SPACE)
#   DEFC14* / DEFM14* / DEFR14* / DEFN14*  (proxy variants, NO space)
#   20-F*   -> 20-F, 20-F/A
#   40-F*   -> 40-F, 40-F/A
FORM_PREFIXES = (
    "8-K",
    "10-K",
    "10-Q",       # NEW: quarterly reports + amendments (10-Q/A, 10-QSB, etc.)
    "DEF 14",
    "DEFC14",
    "DEFM14",
    "DEFR14",
    "DEFN14",
    "20-F",
    "40-F",
)

form_pattern = r'^(?:' + '|'.join(re.escape(p) for p in FORM_PREFIXES) + r')'

df_filings = df[
    df["Form_Type"].astype(str).str.match(form_pattern, na=False) &
    (df["Date_Filed"] == input_date)
].copy()

# =========================
# ACCESSION NUMBER
# =========================
df_filings["Accession_No"] = df_filings["File_Name"].apply(
    lambda x: x.split("/")[-1].replace(".txt","")
)

# =========================
# CAPPED DOWNLOAD
# =========================
def fetch_capped(full_url, max_bytes=MAX_BYTES):
    limiter.wait()
    with session.get(full_url, stream=True, timeout=TIMEOUT) as r:
        chunks = []
        total = 0
        for chunk in r.iter_content(chunk_size=16384):
            if not chunk:
                continue
            chunks.append(chunk)
            total += len(chunk)
            if total >= max_bytes:
                break
    return b"".join(chunks).decode("utf-8", errors="ignore")


# =========================
# ITEM DETECTION FUNCTIONS
# =========================

def detect_items(text):
    """Determine which 8-K items a filing reports.

    Uses EDGAR's own `ITEM INFORMATION:` header metadata first, then falls
    back to explicit "Item X.YZ" captions in the body. Returns (i502, i401,
    i507) flags — always 0/0/0 for non-8-K forms (10-K, 10-Q, DEF 14, etc.).
    """
    low = text.lower().replace("\xa0", " ")

    # 1) Authoritative: EDGAR's declared item titles in the SEC header.
    header_titles = " ".join(re.findall(r'item information:\s*(.+)', low))

    i502 = 1 if "departure of directors" in header_titles else 0
    i401 = 1 if "certifying accountant" in header_titles else 0
    i507 = 1 if "submission of matters to a vote of security holders" in header_titles else 0

    # 2) Fallback: explicit item-number caption in the body.
    if not i502 and re.search(r'item\s*5\.02', low):
        i502 = 1
    if not i401 and re.search(r'item\s*4\.01', low):
        i401 = 1
    if not i507 and re.search(r'item\s*5\.07', low):
        i507 = 1

    return i502, i401, i507


# =========================
# PROCESS FILINGS (THREADED)
# =========================
_progress = {"done": 0}
_progress_lock = threading.Lock()
_total = len(df_filings)


def process_filing(item):
    idx, file = item
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
    for idx, i502, i401, i507 in ex.map(process_filing, list(df_filings["File_Name"].items())):
        results[idx] = (i502, i401, i507)

df_filings["Item_5_02"] = [results[i][0] for i in df_filings.index]
df_filings["Item_4_01"] = [results[i][1] for i in df_filings.index]
df_filings["Item_5_07"] = [results[i][2] for i in df_filings.index]

# =========================
# FILTER RELEVANT FILINGS
# =========================
# Item flags apply only to 8-Ks:
#   - 8-K rows kept ONLY when an item (5.02/4.01/5.07) is flagged
#   - All other form types (10-K*, 10-Q*, DEF 14*, 20-F*, 40-F*, proxies)
#     are listed in full — no item gate.
is_8k = df_filings["Form_Type"].str.startswith("8-K")
has_flag = (
    (df_filings["Item_5_02"] == 1) |
    (df_filings["Item_4_01"] == 1) |
    (df_filings["Item_5_07"] == 1)
)

filtered_df = df_filings[(is_8k & has_flag) | (~is_8k)].copy()

# =========================
# GET INDUSTRY (SIC) — DEDUPED + THREADED
# =========================
def get_industry(cik):
    cik = str(cik).zfill(10)
    api = f"https://data.sec.gov/submissions/CIK{cik}.json"
    limiter.wait()
    try:
        r = session.get(api, timeout=TIMEOUT)
        return r.json().get("sicDescription", "Unknown")
    except Exception:
        return "Unknown"


unique_ciks = filtered_df["CIK"].unique()
print(f"Fetching industry for {len(unique_ciks)} unique companies ...")

industry_map = {}
with ThreadPoolExecutor(max_workers=MAX_WORKERS) as ex:
    for cik, industry in zip(
        unique_ciks,
        ex.map(get_industry, unique_ciks)
    ):
        industry_map[cik] = industry

filtered_df["Industry"] = filtered_df["CIK"].map(industry_map)

# =========================
# SUMMARY
# =========================
summary = (
    filtered_df.groupby(["Form_Type","Industry","Item_5_02","Item_4_01","Item_5_07"])
    .size()
    .reset_index(name="Count")
)

# =========================
# DETAILS
# =========================
details = filtered_df[[
    "CIK",
    "Company",
    "Form_Type",
    "Industry",
    "Date_Filed",
    "Accession_No",
    "Item_5_02",
    "Item_4_01",
    "Item_5_07"
]]

# =========================
# EXPORT EXCEL
# =========================
with pd.ExcelWriter(OUTPUT_FILE) as writer:
    summary.to_excel(writer, sheet_name="Summary", index=False)
    details.to_excel(writer, sheet_name="Company_List", index=False)

print(f"Report Generated: {OUTPUT_FILE}")


# =========================
# EMAIL THE REPORT
# =========================
def send_report(attachment_path, recipients, report_date, n_total, n_matched):
    sender = os.environ.get("SENDER_EMAIL")
    password = os.environ.get("SENDER_APP_PASSWORD")

    if not sender or not password:
        print("WARNING: SENDER_EMAIL / SENDER_APP_PASSWORD not set in the "
              "environment. Report saved but NOT emailed.")
        return

    n502 = int(details["Item_5_02"].sum()) if len(details) else 0
    n401 = int(details["Item_4_01"].sum()) if len(details) else 0
    n507 = int(details["Item_5_07"].sum()) if len(details) else 0

    # Per-family counts across filings in the report
    if len(details):
        ft = details["Form_Type"].astype(str)
        n_8k  = int(ft.str.startswith("8-K").sum())
        n_10k = int(ft.str.startswith("10-K").sum())
        n_10q = int(ft.str.startswith("10-Q").sum())   # NEW
        n_20f = int(ft.str.startswith("20-F").sum())
        n_40f = int(ft.str.startswith("40-F").sum())
        n_def = int(ft.str.startswith("DEF").sum())
    else:
        n_8k = n_10k = n_10q = n_20f = n_40f = n_def = 0

    msg = EmailMessage()
    msg["From"] = sender
    msg["To"] = ", ".join(recipients)
    msg["Subject"] = f"Daily Filing Report ({report_date})"
    msg.set_content(
        f"Automated filing report for {report_date}.\n\n"
        f"Filings scanned   : {n_total}\n"
        f"Filings in report : {n_matched}\n\n"
        f"By form type:\n"
        f"  8-K*   : {n_8k}\n"
        f"  10-K*  : {n_10k}\n"
        f"  10-Q*  : {n_10q}\n"
        f"  20-F*  : {n_20f}\n"
        f"  40-F*  : {n_40f}\n"
        f"  DEF*   : {n_def}\n\n"
        f"8-K items flagged:\n"
        f"  Item 5.02 : {n502}\n"
        f"  Item 4.01 : {n401}\n"
        f"  Item 5.07 : {n507}\n\n"
        f"The full breakdown is attached ({os.path.basename(attachment_path)})."
    )

    with open(attachment_path, "rb") as f:
        data = f.read()
    msg.add_attachment(
        data,
        maintype="application",
        subtype="vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=os.path.basename(attachment_path),
    )

    with smtplib.SMTP("smtp.gmail.com", 587) as server:
        server.starttls()
        server.login(sender, password)
        server.send_message(msg)

    print(f"Email sent to: {', '.join(recipients)}")


send_report(
    OUTPUT_FILE,
    RECIPIENTS,
    input_date,
    n_total=len(df_filings),
    n_matched=len(filtered_df),
)
