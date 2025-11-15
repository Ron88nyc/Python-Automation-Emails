"""
Lease Renewal Automation Script
Author: Ronald Li

Features:
  • Reads lease data from Excel
  • Finds tenants with leases ending in 60 or 30 days
  • Sends HTML renewal reminders via SendGrid SMTP
  • Logs every action to renewal_log.txt
  • (Optional) Runs automatically every Saturday at 10:00 AM with APScheduler
"""

import os
import smtplib
from datetime import datetime
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

import pandas as pd
from apscheduler.schedulers.blocking import BlockingScheduler
from dotenv import load_dotenv

# ---------------------------------------------------------------------------
# 1️⃣ Config & Environment
# ---------------------------------------------------------------------------

load_dotenv()

# SendGrid SMTP (from your .env)
SMTP_HOST = os.getenv("SMTP_HOST", "smtp.sendgrid.net")
SMTP_PORT = int(os.getenv("SMTP_PORT", "587"))
SMTP_USER = os.getenv("SMTP_USER", "apikey")      # SendGrid uses literal "apikey"
SMTP_PASS = os.getenv("SMTP_PASS", "")            # Your SendGrid API key
FROM_EMAIL = os.getenv("FROM_EMAIL", SMTP_USER)   # Verified sender in SendGrid

# Excel file with lease data
EXCEL_FILE = "leases.xlsx"

# Log file
LOG_FILE = "renewal_log.txt"

# Toggle to avoid sending real emails while testing
DRY_RUN = False  # True = only print + log "DRY_RUN", no real sends


# ---------------------------------------------------------------------------
# 2️⃣ Logging helper
# ---------------------------------------------------------------------------

def log_event(tenant: str, email: str, days_left: int, status: str, message: str = ""):
    """
    Append a line to renewal_log.txt

    status: "SENT", "DRY_RUN", "ERROR"
    """
    timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    line = f"{timestamp} | {status:<7} | {days_left:>3} days | {tenant} <{email}> | {message}\n"
    with open(LOG_FILE, "a", encoding="utf-8") as f:
        f.write(line)


# ---------------------------------------------------------------------------
# 3️⃣ Email sending via SendGrid SMTP
# ---------------------------------------------------------------------------

def send_smtp(to_addr: str, subject: str, html_content: str, tenant: str, days_left: int):
    """
    Send an HTML email using SendGrid SMTP.

    Uses:
      SMTP_HOST, SMTP_PORT, SMTP_USER ('apikey'), SMTP_PASS (API key), FROM_EMAIL
    """

    if DRY_RUN:
        print(f"[DRY-RUN] Would send to {to_addr}: {subject}")
        log_event(tenant, to_addr, days_left, "DRY_RUN", "No email sent (DRY_RUN=True)")
        return

    msg = MIMEMultipart("alternative")
    msg["From"] = FROM_EMAIL
    msg["To"] = to_addr
    msg["Subject"] = subject
    msg.attach(MIMEText(html_content, "html"))

    try:
        with smtplib.SMTP(SMTP_HOST, SMTP_PORT, timeout=30) as server:
            server.starttls()
            server.login(SMTP_USER, SMTP_PASS)
            server.send_message(msg)

        print(f"✅ Sent email to {to_addr} via SendGrid SMTP")
        log_event(tenant, to_addr, days_left, "SENT", "OK")

    except Exception as e:
        err = str(e)
        print(f"❌ Error sending to {to_addr}: {err}")
        log_event(tenant, to_addr, days_left, "ERROR", err)


# ---------------------------------------------------------------------------
# 4️⃣ Lease processing & reminder logic
# ---------------------------------------------------------------------------

def process_lease_data():
    """
    Load leases.xlsx, find tenants exactly 60 or 30 days from lease end,
    and send them reminders.

    Expected columns in Excel:
      Tenant, Email, Lease_End_Date
    """

    if not os.path.exists(EXCEL_FILE):
        print(f"⚠️ Excel file not found: {EXCEL_FILE}")
        return

    df = pd.read_excel(EXCEL_FILE)

    # Normalize columns (adjust these if your headers differ)
    required_cols = ["Tenant", "Email", "Lease_End_Date"]
    for col in required_cols:
        if col not in df.columns:
            raise ValueError(f"Missing required column in Excel: {col}")

    # Parse dates
    df["Lease_End_Date"] = pd.to_datetime(df["Lease_End_Date"], errors="coerce")

    today = datetime.today().date()
    df["Days_Remaining"] = (df["Lease_End_Date"].dt.date - today).dt.days

    # Filter for exactly 60 or 30 days remaining
    targets = df[df["Days_Remaining"].isin([60, 30])]

    if targets.empty:
        print("No leases at 60 or 30 days today.")
        return

    print(f"Found {len(targets)} leases at 60 or 30 days. Processing...\n")

    for _, row in targets.iterrows():
        tenant = str(row["Tenant"]).strip()
        email = str(row["Email"]).strip()
        days_left = int(row["Days_Remaining"])

        if not email or email.lower() == "nan":
            print(f"Skipping {tenant}: no email.")
            log_event(tenant, "N/A", days_left, "ERROR", "Missing email")
            continue

        subject = f"Lease Renewal Reminder - {days_left} Days Remaining"

        html = f"""
        <html>
          <body style="font-family: -apple-system, BlinkMacSystemFont, 'Segoe UI', sans-serif;">
            <p>Hi {tenant},</p>
            <p>This is a friendly reminder that your lease is ending in <b>{days_left} days</b>.</p>
            <p>If you’d like to renew or discuss options, please reply to this email and our team will assist.</p>
            <p>Thank you,<br>Rose Property Management</p>
          </body>
        </html>
        """

        send_smtp(email, subject, html, tenant, days_left)


# ---------------------------------------------------------------------------
# 5️⃣ Weekly scheduler (every Saturday 10:00 AM)
# ---------------------------------------------------------------------------

def schedule_weekly_run():
    """
    Run process_lease_data() automatically every Saturday at 10:00 AM.
    """
    scheduler = BlockingScheduler()

    # day_of_week='sat' (0=mon .. 6=sun), hour=10 (10:00)
    scheduler.add_job(process_lease_data, "cron", day_of_week="sat", hour=10)

    print("🕒 Scheduler started. Reminders will run every Saturday at 10:00 AM.")
    scheduler.start()


# ---------------------------------------------------------------------------
# 6️⃣ Entry point
# ---------------------------------------------------------------------------

if __name__ == "__main__":
    # For development:
    #   - Set DRY_RUN = True above
    #   - Run once to see which emails would go out
    #
    # For production:
    #   - Set DRY_RUN = False
    #   - Use schedule_weekly_run() OR a system scheduler calling this script

    print("🚀 Running Lease Renewal Automation...\n")

    # Option A: run once and exit
    process_lease_data()

    # Option B: enable weekly scheduler (uncomment when ready)
    # schedule_weekly_run()
