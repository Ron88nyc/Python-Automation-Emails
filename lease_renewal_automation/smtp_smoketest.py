
import smtplib
from email.mime.text import MIMEText

HOST = "smtp.sendgrid.net"
PORT = 587
USER = "apikey"
PASS = "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"

FROM = "rlipmg@outlook.com"  # must match verified sender in SendGrid
TO = "rlipmg@outlook.com"    # send to yourself first to test

msg = MIMEText("<p>Test email from Lease Renewal Automation via SendGrid.</p>", "html")
msg["From"] = FROM
msg["To"] = TO
msg["Subject"] = "SendGrid SMTP smoke test"

print(f"Connecting to {HOST}:{PORT} as {USER}...")
with smtplib.SMTP(HOST, PORT, timeout=30) as s:
    s.set_debuglevel(1)
    s.starttls()
    s.login(USER, PASS)
    s.send_message(msg)
    print("✅ Sent test email successfully.")
