#!/usr/bin/env python3
"""
send_birthdays.py

- Reads ./birthdays.csv (S/N,NAMES,DESIGNATION,DATE OF BIRTH)
- Sends a single plain-text email reminder to HR when any staff have birthdays same-day or in 2, 5, or 10 days.
- Groups reminders by days-remaining: 10, 5, 2, 0
- Environment variables:
    GMAIL_USER             (required) Gmail address to send from (e.g. example@gmail.com)
    GMAIL_APP_PASSWORD     (required) App password for the Gmail account (recommended) or SMTP password
    HR_EMAIL               (optional) recipient; defaults to hr@riskcontrolnigeria.com
    TZ                     (optional) timezone name (IANA), defaults to Africa/Lagos
"""
import csv
import datetime
import os
import re
import smtplib
from collections import defaultdict
from email.message import EmailMessage

try:
    # Python 3.9+
    from zoneinfo import ZoneInfo
except Exception:
    ZoneInfo = None

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
    # Normalize spacing and casing (e.g., "28TH OF AUGUST" -> "28 of August")
    s = re.sub(r"\s+of\s+", " ", s, flags=re.IGNORECASE)
    s = " ".join(s.split())
    return s.title()                      # makes month name Title case


def parse_day_month(s: str):
    """Returns (day:int, month:int) or None"""
    s = normalize_date_str(s)
    for fmt in ("%d %B", "%d %b", "%B %d", "%b %d", "%d-%B", "%d/%B"):
        try:
            dt = datetime.datetime.strptime(s, fmt)
            return dt.day, dt.month
        except Exception:
            continue
    # Try extracting numbers and month words manually
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
        # fallback: no zoneinfo available
        return datetime.datetime.utcnow().date()


def read_birthdays(csv_path=CSV_PATH):
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
    """Return number of days from from_date until next birthday (0 if today)."""
    year = from_date.year
    try:
        next_bd = datetime.date(year, month, day)
    except ValueError:
        # Invalid date
        return None
    if next_bd < from_date:
        # next year's birthday
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

    # Order windows descending (10,5,2,0) so that soonest appear last or choose 10->5->2->0
    for window in sorted(VALID_WINDOWS, reverse=True):
        items = grouped.get(window, [])
        if not items:
            continue
        header = "Same day" if window == 0 else f"{window} day{'s' if window != 1 else ''} remaining"
        lines.append(header + ":")
        for m in items:
            # Show next birthday date for clarity
            parsed = parse_day_month(m['dob'])
            if parsed:
                d, mo = parsed
                # compute the year for next birthday
                days = days_until_next_birthday(d, mo, for_date)
                if days is None:
                    nb_str = m['dob']
                else:
                    nb_date = for_date + datetime.timedelta(days=days)
                    nb_str = nb_date.strftime('%d %B %Y')
            else:
                nb_str = m['dob']
            lines.append(f"- {m['name']} — {m['designation']} (DOB: {m['dob']}) — Next: {nb_str}")
        lines.append("")

    lines.append("Please join in wishing them well."
                 )
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

    try:
        with smtplib.SMTP_SSL(SMTP_HOST, SMTP_PORT_SSL, timeout=30) as server:
            server.login(smtp_user, smtp_pass)
            server.send_message(msg)
    except Exception:
        raise


def main():
    gmail_user = os.environ.get("GMAIL_USER")
    gmail_pass = os.environ.get("GMAIL_APP_PASSWORD")
    hr_email = os.environ.get("HR_EMAIL", DEFAULT_HR_EMAIL)

    if not gmail_user or not gmail_pass:
        print("Missing GMAIL_USER or GMAIL_APP_PASSWORD environment variables. Exiting.")
        return

    try:
        rows = read_birthdays(CSV_PATH)
    except FileNotFoundError:
        print(f"{CSV_PATH} not found. Place the CSV in the repository root and name it 'birthdays.csv'.")
        return

    today = today_date()
    grouped = defaultdict(list)

    for r in rows:
        parsed = parse_day_month(r["dob"])
        if not parsed:
            continue
        day, month = parsed
        delta = days_until_next_birthday(day, month, today)
        if delta is None:
            continue
        if delta in VALID_WINDOWS:
            grouped[delta].append(r)

    if not any(grouped.values()):
        print(f"No birthdays within {sorted(VALID_WINDOWS)} days of {today.isoformat()}.")
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
