#!/usr/bin/env python3
"""
send_birthdays.py

- Downloads an Excel workbook from SHAREPOINT_URL (if provided) and reads staff birthdays.
  Falls back to ./birthdays.csv if download or parsing fails.
- Sends a single plain-text email reminder to HR for birthdays in these windows:
    10 days, 5 days, 2 days, same day (0)
- Environment variables:
    GMAIL_USER         (required) Gmail address to send from
    GMAIL_APP_PASSWORD (required) App password for the Gmail account
    HR_EMAIL           (optional) defaults to hr@riskcontrolnigeria.com
    TZ                 (optional) timezone name (IANA), defaults to Africa/Lagos
    SHAREPOINT_URL     (optional) direct-download link to the Excel workbook
"""
import csv
import datetime
import os
import re
import smtplib
import sys
import tempfile
import requests
import shutil
from collections import defaultdict
from email.message import EmailMessage

try:
    # Python 3.9+
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

# Optional dependency: pandas + openpyxl
try:
    import pandas as pd
except Exception:
    pd = None

CSV_PATH = "birthdays.csv"
DEFAULT_HR_EMAIL = "hr@riskcontrolnigeria.com"
SMTP_HOST = "smtp.gmail.com"
SMTP_PORT_SSL = 465

ORDINAL_RE = re.compile(r"(\d+)(st|nd|rd|th)", flags=re.IGNORECASE)
VALID_WINDOWS = {10, 5, 2, 0}


def normalize_date_str(s: str) -> str:
    if not s:
        return ""
    s = s.strip()
    s = ORDINAL_RE.sub(r"\1", s)          # remove ordinal suffixes
    s = s.replace(",", " ")
    s = re.sub(r"\s+of\s+", " ", s, flags=re.IGNORECASE)
    s = " ".join(s.split())
    return s.title()


def parse_day_month(s: str):
    """Returns (day:int, month:int) or None"""
    s = normalize_date_str(s)
    for fmt in ("%d %B", "%d %b", "%B %d", "%b %d", "%d-%B", "%d/%B"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.day, dt.month
        except Exception:
            continue
    m = re.search(r"(\d{1,2})\s+([A-Za-z]+)", s)
    if m:
        day = int(m.group(1))
        mon = m.group(2)
        try:
            dt = datetime.datetime.strptime(mon.title(), "%B")
            return day, dt.month
        except Exception:
            try:
                dt = datetime.datetime.strptime(mon.title(), "%b")
                return day, dt.month
            except Exception:
                return None
    return None


def today_date():
    tz_name = os.environ.get("TZ", "Africa/Lagos")
    if ZoneInfo:
        try:
            now = datetime.datetime.now(ZoneInfo(tz_name))
            return now.date()
        except Exception:
            return datetime.datetime.utcnow().date()
    else:
        return datetime.datetime.utcnow().date()


def download_sharepoint_excel(url: str, timeout: int = 30) -> str:
    """
    Download the file at `url` (streaming) into a temporary file and return the local path.
    Raises requests.HTTPError on failure.
    """
    if not url:
        raise ValueError("No URL provided")
    resp = requests.get(url, stream=True, timeout=timeout)
    resp.raise_for_status()
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx")
    tmp_path = tmp.name
    tmp.close()
    with requests.get(url, stream=True, timeout=timeout) as r:
        r.raise_for_status()
        with open(tmp_path, "wb") as f:
            shutil.copyfileobj(r.raw, f)
    return tmp_path


def read_birthdays_from_excel(path: str):
    """
    Read the first sheet of an Excel file and map columns to expected keys.
    Returns list of dicts with keys: sn, name, designation, dob
    """
    if pd is None:
        raise RuntimeError("pandas is required to read Excel files. Install pandas and openpyxl.")
    df = pd.read_excel(path, engine="openpyxl")
    # Normalize columns
    cols = {c.strip().lower(): c for c in df.columns}
    def col_by_candidates(candidates):
        for cand in candidates:
            cand = cand.strip().lower()
            if cand in cols:
                return cols[cand]
        return None
    sn_col = col_by_candidates(["s/n", "sn", "s n", "serial"])
    name_col = col_by_candidates(["names", "name", "employee", "full name"])
    desig_col = col_by_candidates(["designation", "role", "job title", "position"])
    dob_col = col_by_candidates(["date of birth", "dob", "birthdate", "date"])
    rows = []
    for _, row in df.iterrows():
        rows.append({
            "sn": _ if sn_col is None else row.get(sn_col),
            "name": "" if name_col is None else str(row.get(name_col)).strip(),
            "designation": "" if desig_col is None else str(row.get(desig_col)).strip(),
            "dob": "" if dob_col is None else str(row.get(dob_col)).strip()
        })
    return rows


def read_birthdays_from_csv(csv_path=CSV_PATH):
    rows = []
    with open(csv_path, newline="", encoding="utf-8") as f:
        reader = csv.DictReader(f)
        for r in reader:
            rows.append({
                "sn": r.get("S/N") or r.get("S/N ") or r.get("SN"),
                "name": (r.get("NAMES") or "").strip(),
                "designation": (r.get("DESIGNATION") or "").strip(),
                "dob": (r.get("DATE OF BIRTH") or "").strip()
            })
    return rows


def days_until_next_birthday(day: int, month: int, from_date: datetime.date):
    year = from_date.year
    try:
        next_bd = datetime.date(year, month, day)
    except ValueError:
        return None
    if next_bd < from_date:
        try:
            next_bd = datetime.date(year + 1, month, day)
        except ValueError:
            return None
    return (next_bd - from_date).days


def build_email_body(grouped: dict, for_date: datetime.date):
    lines = []
    lines.append("Hello HR team,")
    lines.append("")
    lines.append(f"This is an automated birthday reminder for {for_date.strftime('%d %B %Y')}.")
    lines.append("")
    total = sum(len(v) for v in grouped.values())
    if total == 0:
        lines.append("There are no upcoming birthdays in the configured windows (10, 5, 2, 0 days).")
        lines.append("")
        lines.append("Regards,")
        lines.append("Birthday Automation")
        return "\n".join(lines)
    lines.append(f"There are {total} staff with birthdays in the next configured windows:")
    lines.append("")
    for window in sorted(VALID_WINDOWS, reverse=True):
        items = grouped.get(window, [])
        if not items:
            continue
        header = "Same day" if window == 0 else f"{window} day{'s' if window != 1 else ''} remaining"
        lines.append(header + ":")
        for m in items:
            parsed = parse_day_month(m["dob"])
            if parsed:
                d, mo = parsed
                days = days_until_next_birthday(d, mo, for_date)
                if days is None:
                    nb_str = m["dob"]
                else:
                    nb_date = for_date + datetime.timedelta(days=days)
                    nb_str = nb_date.strftime("%d %B %Y")
            else:
                nb_str = m["dob"]
            lines.append(f"- {m['name']} — {m['designation']} (DOB: {m['dob']}) — Next: {nb_str}")
        lines.append("")
    lines.append("Please join in wishing them well.")
    lines.append("")
    lines.append("Regards,")
    lines.append("Birthday Automation")
    return "\n".join(lines)


def send_email(subject, body, from_addr, to_addrs, smtp_user, smtp_pass):
    msg = EmailMessage()
    msg["Subject"] = subject
    msg["From"] = from_addr
    msg["To"] = ", ".join(to_addrs) if isinstance(to_addrs, (list, tuple)) else to_addrs
    msg.set_content(body)
    with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT_SSL, timeout=30) as server:
        server.login(smtp_user, smtp_pass)
        server.send_message(msg)


def main():
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASSWORD")
    hr_email = os.environ.get("HR_EMAIL", DEFAULT_HR_EMAIL)
    sharepoint_url = os.environ.get("SHAREPOINT_URL")

    if not gmail_user or not gmail_pass:
        print("Missing GMAIL_USER or GMAIL_APP_PASSWORD environment variables. Exiting.")
        sys.exit(0)

    rows = []
    # Try Excel from SharePoint first (if provided)
    if sharepoint_url:
        print("SHAREPOINT_URL provided — attempting to download Excel...")
        try:
            path = download_sharepoint_excel(sharepoint_url)
            print(f"Downloaded Excel to {path}; parsing...")
            rows = read_birthdays_from_excel(path)
            print(f"Parsed {len(rows)} rows from Excel.")
            try:
                os.remove(path)
            except Exception:
                pass
        except Exception as e:
            print("Failed to download or parse Excel from SHAREPOINT_URL:", str(e))
            print("Falling back to local CSV if available.")
            try:
                rows = read_birthdays_from_csv(CSV_PATH)
                print(f"Parsed {len(rows)} rows from CSV.")
            except FileNotFoundError:
                print(f"{CSV_PATH} not found. Exiting.")
                sys.exit(0)
    else:
        # No sharepoint URL — use CSV
        try:
            rows = read_birthdays_from_csv(CSV_PATH)
            print(f"Parsed {len(rows)} rows from CSV.")
        except FileNotFoundError:
            print(f"{CSV_PATH} not found and SHAREPOINT_URL not set. Exiting.")
            sys.exit(0)

    today = today_date()
    grouped = defaultdict(list)
    for r in rows:
        parsed = parse_day_month(r.get("dob") or "")
        if not parsed:
            continue
        day, month = parsed
        delta = days_until_next_birthday(day, month, today)
        if delta is None:
            continue
        if delta in VALID_WINDOWS:
            grouped[delta].append(r)

    if not any(grouped.values()):
        print(f"No birthdays within {sorted(VALID_WINDOWS)} days of {today.isoformat()}. Nothing to send.")
        return

    subject_parts = []
    for w in sorted(grouped.keys()):
        label = "Same day" if w == 0 else f"{w}d"
        subject_parts.append(f"{len(grouped[w])} x {label}")
    subject = f"Birthday Reminder — {' | '.join(subject_parts)} — {today.strftime('%d %b %Y')}"
    body = build_email_body(grouped, today)

    try:
        send_email(subject, body, gmail_user, [hr_email], gmail_user, gmail_pass)
        print(f"Sent birthday reminder to {hr_email} for windows: {', '.join(str(k) for k in sorted(grouped.keys()))}.")
    except Exception as e:
        print("Failed to send email:", str(e))


if __name__ == "__main__":
    main()
