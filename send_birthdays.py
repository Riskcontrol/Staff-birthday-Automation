#!/usr/bin/env python3
"""
send_birthdays.py

- Downloads the staff birthday workbook from SHAREPOINT_URL if set
- Falls back to birthdays.csv if download fails
- Sends a plain-text birthday reminder email for the 10, 5, 2, and same-day windows
"""
import csv
import datetime
import io
import os
import re
import smtplib
import sys
import tempfile
from collections import defaultdict
from email.message import EmailMessage

import requests

try:
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

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
    s = ORDINAL_RE.sub(r"\1", s)
    s = s.replace(",", " ")
    s = re.sub(r"\s+of\s+", " ", s, flags=re.IGNORECASE)
    s = " ".join(s.split())
    return s.title()


def parse_day_month(s: str):
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
            return datetime.datetime.now(ZoneInfo(tz_name)).date()
        except Exception:
            return datetime.datetime.utcnow().date()
    return datetime.datetime.utcnow().date()


def download_sharepoint_excel(url: str, timeout: int = 30):
    if not url:
        raise ValueError("SHAREPOINT_URL is empty")

    # Try direct download
    response = requests.get(url, allow_redirects=True, timeout=timeout)
    print("SharePoint HTTP status:", response.status_code)
    print("SharePoint Content-Type:", response.headers.get("Content-Type"))
    print("SharePoint final URL:", response.url)

    if response.status_code != 200:
        raise RuntimeError(f"SharePoint URL returned status {response.status_code}")

    content_type = response.headers.get("Content-Type", "").lower()
    if "html" in content_type or "text/plain" in content_type:
        raise RuntimeError(
            "URL did not return an Excel file. It likely returned an HTML login page or SharePoint landing page."
        )

    if "excel" not in content_type and "octet-stream" not in content_type and "xlsx" not in content_type and "zip" not in content_type:
        # still allow it in case the URL is returning a valid file with a generic content type
        print("Warning: unexpected content type for an Excel file. Continuing anyway.")

    # Save binary content to temp file
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmp:
        tmp.write(response.content)
        tmp_path = tmp.name

    print(f"Saved workbook to: {tmp_path}")
    return tmp_path


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


def read_birthdays_from_excel(path: str):
    if pd is None:
        raise RuntimeError("pandas/openpyxl not installed")

    # Parse from bytes directly to avoid file issues
    with open(path, "rb") as f:
        data = f.read()

    df = pd.read_excel(io.BytesIO(data), engine="openpyxl", dtype=str, keep_default_na=False)
    print("Excel columns detected:", list(df.columns))

    # normalize column names
    normalized = {str(c).strip().lower(): c for c in df.columns}
    def find_col(*names):
        for n in names:
            key = str(n).strip().lower()
            if key in normalized:
                return normalized[key]
        return None

    sn_col = find_col("s/n", "sn", "s n", "serial")
    name_col = find_col("names", "name", "full name", "employee name")
    desig_col = find_col("designation", "role", "job title", "position")
    dob_col = find_col("date of birth", "dob", "date", "birth date")

    rows = []
    for _, row in df.iterrows():
        rows.append({
            "sn": "" if sn_col is None else str(row.get(sn_col, "")).strip(),
            "name": "" if name_col is None else str(row.get(name_col, "")).strip(),
            "designation": "" if desig_col is None else str(row.get(desig_col, "")).strip(),
            "dob": "" if dob_col is None else str(row.get(dob_col, "")).strip()
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

    if sharepoint_url:
        print("Attempting to use SHAREPOINT_URL")
        try:
            workbook_path = download_sharepoint_excel(sharepoint_url)
            rows = read_birthdays_from_excel(workbook_path)
            print(f"Parsed {len(rows)} rows from SharePoint workbook.")
            os.remove(workbook_path)
        except Exception as e:
            print("Excel download/parse failed:", str(e))
            print("Falling back to local CSV...")
            try:
                rows = read_birthdays_from_csv(CSV_PATH)
                print(f"Parsed {len(rows)} rows from CSV fallback.")
            except FileNotFoundError:
                print(f"{CSV_PATH} not found. Exiting.")
                sys.exit(1)
    else:
        try:
            rows = read_birthdays_from_csv(CSV_PATH)
            print(f"Parsed {len(rows)} rows from CSV.")
        except FileNotFoundError:
            print(f"{CSV_PATH} not found and SHAREPOINT_URL not set. Exiting.")
            sys.exit(1)

    today = today_date()
    grouped = defaultdict(list)

    for r in rows:
        dob = (r.get("dob") or "").strip()
        parsed = parse_day_month(dob)
        if not parsed:
            continue
        day, month = parsed
        delta = days_until_next_birthday(day, month, today)
        if delta in VALID_WINDOWS:
            grouped[delta].append(r)

    if not any(grouped.values()):
        print(f"No birthdays in the configured windows for {today.isoformat()}.")
        return

    subject_parts = []
    for w in sorted(grouped.keys()):
        label = "Same day" if w == 0 else f"{w}d"
        subject_parts.append(f"{len(grouped[w])} x {label}")

    subject = f"Birthday Reminder — {' | '.join(subject_parts)} — {today.strftime('%d %b %Y')}"
    body = build_email_body(grouped, today)

    try:
        send_email(subject, body, gmail_user, [hr_email], gmail_user, gmail_pass)
        print(f"Sent reminder to {hr_email}")
    except Exception as e:
        print("Failed to send email:", str(e))
        raise


if __name__ == "__main__":
    main()
