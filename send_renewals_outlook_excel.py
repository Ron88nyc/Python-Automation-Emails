"""
Automated Lease Renewal Reminder System
---------------------------------------
This script reads lease data from an Excel sheet and sends personalized renewal reminder emails
via Outlook's SMTP server (Office365). It automatically detects which tenants are
60 days or 30 days away from their lease end and emails them accordingly.

✅ Sends reminders weekly (every Saturday at 10am) using APScheduler
✅ Avoids duplicates (won’t re-send within 21 days)
✅ Updates Excel with last sent date
✅ Logs all actions to send_log.csv

Folder layout suggestion:
lease_renewal_automation/
  ├─ send_renewals_outlook_excel.py   <-- this script
  ├─ .env                             <-- environment variables (SMTP settings)
  └─ prospects.xlsx                   <-- your data
"""

import os
import smtplib
import traceback
from datetime import date, datetime
from email.message import EmailMessage

import pandas as pd
from apscheduler.schedulers.blocking import BlockingScheduler  # for weekly scheduling
from dotenv import load_dotenv  # loads environment variables from a .env file if present

# Load .env values (optional but convenient for local dev)
# If a .env file exists next to the script, its variables will be loaded into the environment.
load_dotenv()

# ============ SETTINGS YOU CAN CUSTOMIZE ============

EXCEL_PATH = "prospects.xlsx"        # Excel file containing tenant/prospect data
SHEET_NAME = "Prospects"             # Sheet name inside the Excel file
LOG_PATH = "send_log.csv"            # Log file to track actions

# Outlook SMTP settings (use environment variables for security)
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.office365.com")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER")
SMTP_PASS = os.getenv("SMTP_PASS")

# Milestones: how many days before lease end to send reminders
# "tolerance" gives flexibility (e.g., run Saturday but catch Thursday or Friday expirations too)
MILESTONES = [
    {"days": 60, "tolerance": 2, "col": "last_sent_60"},
    {"days": 30, "tolerance": 2, "col": "last_sent_30"},
]

COOLDOWN_DAYS = 21  # don't re-send within this period
DRY_RUN = False     # if True, prints but doesn't actually send emails

# Email templates (use {placeholders} to personalize content)
SUBJECT_60 = "Let's lock in renewal options for {property} — Unit {unit}"
SUBJECT_30 = "30 days left for {property} — Renewal options for Unit {unit}"

HTML_60 = """
<p>Hi {first_name},</p>
<p>Your lease for <b>{property} – {unit}</b> ends on <b>{lease_end_fmt}</b> (about 60 days out).</p>
<p>We’d love to help you renew early and secure the best terms.</p>
<ul>
  <li>Flexible term lengths</li>
  <li>Priority maintenance scheduling</li>
  <li>Fast digital paperwork</li>
</ul>
<p>Reply to this email to get started or schedule a quick call.</p>
<p>Thanks!<br>Resident Relations</p>
<hr><p style="font-size:12px;color:#666">To stop renewal emails, reply “STOP”.</p>
"""

HTML_30 = """
<p>Hi {first_name},</p>
<p>Quick reminder: your lease for <b>{property} – {unit}</b> ends on <b>{lease_end_fmt}</b> (about 30 days).</p>
<p>We can finalize your renewal in minutes—reply here and we’ll prepare options.</p>
<p>Thanks!<br>Resident Relations</p>
<hr><p style="font-size:12px;color:#666">To stop renewal emails, reply “STOP”.</p>
"""

# ============ HELPER FUNCTIONS ============

def parse_date(cell):
    """Safely convert a cell value to a Python date (handles Excel, string, or empty)."""
    if pd.isna(cell) or cell == "":
        return None
    if isinstance(cell, (datetime, pd.Timestamp)):
        return cell.date()
    if isinstance(cell, date):
        return cell
    try:
        return pd.to_datetime(cell).date()
    except Exception:
        return None

def within_window(days_until, target, tol):
    """Returns True if today is within ±tol days of the target window."""
    return (target - tol) <= days_until <= (target + tol)

def send_outlook(to_addr, subject, html):
    """Send an HTML email using Outlook's SMTP service."""
    msg = EmailMessage()
    msg["From"] = SMTP_USER
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.set_content("This message contains HTML content.")
    msg.add_alternative(html, subtype="html")

    if DRY_RUN:
        print(f"[DRY-RUN] Would send to {to_addr}: {subject}")
        return

    # Connect to Outlook SMTP, start TLS encryption, log in, and send
    with smtplib.SMTP(SMTP_HOST, SMTP_PORT) as s:
        s.starttls()
        s.login(SMTP_USER, SMTP_PASS)
        s.send_message(msg)

def log(status, email, note=""):
    """Append an entry to the log CSV file."""
    ts = datetime.now().isoformat(timespec="seconds")
    line = f'{ts},{email},{status},"{note.replace(",", ";")}"\n'
    new = not os.path.exists(LOG_PATH)
    with open(LOG_PATH, "a", encoding="utf-8") as f:
        if new:
            f.write("ts,email,status,note\n")
        f.write(line)

# ============ CORE FUNCTION ============

def process_reminders():
    """Main workflow: reads Excel, filters who to email, sends reminders, updates sheet."""
    print(f"\n[RUNNING] Lease renewal reminder check — {datetime.now()}\n")

    if not (SMTP_USER and SMTP_PASS):
        raise SystemExit("Missing SMTP_USER/SMTP_PASS environment variables. "
                         "Create a .env file or set them in your OS environment.")

    # Read Excel data
    df = pd.read_excel(EXCEL_PATH, sheet_name=SHEET_NAME, engine="openpyxl")

    # Ensure the tracking columns exist
    for m in MILESTONES:
        if m["col"] not in df.columns:
            df[m["col"]] = ""

    today = date.today()
    sent_counter = {m["col"]: 0 for m in MILESTONES}

    # Loop through every tenant row
    for idx, row in df.iterrows():
        email_addr = str(row.get("email") or "").strip()
        if not email_addr:
            continue  # skip empty emails

        if str(row.get("opted_out") or "").lower() in ("true", "yes", "1"):
            log("skip", email_addr, "opted_out")
            continue

        lease_end = parse_date(row.get("lease_end_date"))
        if not lease_end:
            log("skip", email_addr, "invalid lease_end_date")
            continue

        days_until = (lease_end - today).days
        lease_end_fmt = lease_end.strftime("%b %d, %Y")

        # Check both 60-day and 30-day windows
        for m in MILESTONES:
            last_sent = parse_date(row.get(m["col"]))
            # skip if recently emailed for this milestone
            if last_sent and (today - last_sent).days < COOLDOWN_DAYS:
                continue

            if within_window(days_until, m["days"], m["tolerance"]):
                ctx = {
                    "first_name": str(row.get("first_name") or "there").strip(),
                    "property": str(row.get("property") or "").strip(),
                    "unit": str(row.get("unit") or "").strip(),
                    "lease_end_fmt": lease_end_fmt,
                }

                # Select the right template based on milestone
                if m["days"] == 60:
                    subject = SUBJECT_60.format(**ctx)
                    html = HTML_60.format(**ctx)
                else:
                    subject = SUBJECT_30.format(**ctx)
                    html = HTML_30.format(**ctx)

                try:
                    send_outlook(email_addr, subject, html)
                    df.at[idx, m["col"]] = today.strftime("%Y-%m-%d")  # stamp that we emailed
                    sent_counter[m["col"]] += 1
                    log("sent", email_addr, f"milestone={m['days']}")
                except Exception:
                    log("error", email_addr, traceback.format_exc()[:300])

    # Save updated Excel back
    with pd.ExcelWriter(EXCEL_PATH, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, index=False, sheet_name=SHEET_NAME)

    print(f"Done. Sent summary: {sent_counter}\n")

# ============ SCHEDULER (RUN EVERY SATURDAY) ============

def start_weekly_scheduler():
    """Schedules the reminder process to run automatically every Saturday at 10am."""
    scheduler = BlockingScheduler()

    # Cron-style schedule: run at hour 10, day_of_week 'sat'
    @scheduler.scheduled_job('cron', day_of_week='sat', hour=10)
    def weekly_job():
        process_reminders()

    print("Scheduler started — reminders will run every Saturday at 10:00 AM.")
    scheduler.start()

# ============ RUN ============
if __name__ == "__main__":
    # Choose one of the two lines below:
    # 1) Run once (good for testing)
    # process_reminders()

    # 2) Run forever and send every Saturday at 10am (production)
    start_weekly_scheduler()
