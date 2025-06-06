import smtplib
import os
from email.message import EmailMessage
from datetime import datetime
from dotenv import load_dotenv
import os

load_dotenv()

def send_email():
    # Get current date and time formatted like: '2025-06-06 12:30'
    now = datetime.now().strftime("%Y-%m-%d %H:%M")
    # CONFIGURE THESE
    sender_email = os.getenv('user_email_send')
    receiver_email = os.getenv('user_email_recv')
    app_password = os.getenv('app_pass')  # Not your Gmail password!
    xlsx_file = "status_report.xlsx"  # Path to the file you already saved
    
    # Split emails into list for SMTP
    to_addrs = [email.strip() for email in receiver_email.split(",")]

    # Compose email
    msg = EmailMessage()
    msg['From'] = sender_email
    msg['To'] = receiver_email
    msg['Subject'] = f"Daily Phonesuite Report - {now}"
    msg.set_content("Please find the attached devices report.")

    # Attach Excel file
    with open(xlsx_file, 'rb') as f:
        file_data = f.read()
        file_name = os.path.basename(xlsx_file)
        msg.add_attachment(file_data, maintype='application', subtype='vnd.openxmlformats-officedocument.spreadsheetml.sheet', filename=file_name)

    # Send email using Gmail SMTP
    with smtplib.SMTP_SSL('smtp.gmail.com', 465) as smtp:
        smtp.login(sender_email, app_password)
        smtp.send_message(msg, from_addr=sender_email, to_addrs=to_addrs)

    print("Email sent!")