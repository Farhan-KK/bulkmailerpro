import smtplib 
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import time

# Function to send email with attachment
def send_email(to_email, subject, body, from_email, from_password, attachments=None):
    try:
        # Create message container
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Attach the email body to the message (HTML format)
        msg.attach(MIMEText(body, 'html'))

        # Attach files (if any)
        if attachments:
            for attachment_path in attachments:
                filename = os.path.basename(attachment_path)
                try:
                    with open(attachment_path, "rb") as attachment:
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(attachment.read())
                    encoders.encode_base64(part)
                    part.add_header('Content-Disposition', f"attachment; filename={filename}")
                    msg.attach(part)
                except Exception as file_error:
                    print(f"Failed to attach file {attachment_path}: {str(file_error)}")

        # Connect to the server
        server = smtplib.SMTP('smtp.gmail.com', 587)
        server.starttls()
        server.login(from_email, from_password)

        # Send email
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()

        print(f"Email successfully sent to {to_email}")
        return True

    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}. Failed to send email to {to_email}.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}. Failed to send email to {to_email}.")
    return False

# Function to send email with retry logic
def send_email_with_retry(to_email, subject, body, from_email, from_password, attachments=None, retries=3):
    for attempt in range(retries):
        if send_email(to_email, subject, body, from_email, from_password, attachments):
            return True
        print(f"Retrying... ({attempt + 1}/{retries})")
        time.sleep(5)  # Wait 5 seconds before retrying
    return False

# Read the CSV file
file_path = r""
attendees_data = pd.read_excel(file_path, engine='openpyxl')

# Email credentials
from_email = ''
from_password = ''

# Email content template with placeholders
subject = 'Seating and Food Arrangement Details for the International Conclave on Next-Gen Higher Education'
body_template = '''
Dear Researchers, Students and Faculty from CUSAT,
,<br><br>

Greetings from The Kerala State Higher Education Council (KSHEC)!<br><br>

Considering the exceptional number of registrations, we have arranged for multiple venues to ensure a smooth and comfortable experience for all participants.
Seating Arrangements.<br><br>

The sessions will be held at the following venues:<br>

For Researchers and Faculty: Seminar Hall, Lab Block, SOE CUSAT<br>
For Students: Placement Cell Auditorium, SOE CUSAT<br>

The registration desk is arranged at Software Block. The detailed seating and food arrangements for each venue are provided in the attachment. Please refer to it to confirm your assigned seat and locate your venue using the included map links.<br><br>

We kindly request your cooperation in proceeding directly to your assigned venue upon arrival to facilitate smooth organization and proceedings.<br>

We look forward to welcoming you to this enriching event!<br>

Warm regards,<br>
Event Team<br>
International Conclave on Next-Gen Higher Education<br>
The Kerala State Higher Education Council (KSHEC)
'''

# List of common attachments
common_attachments = [r""]

# Initialize counters for total emails and successful emails
total_emails = 0
successful_emails = 0
batch_size = 10
batch_delay = 10 * 60  # 10 minutes delay

# Iterate through the email list
for index, row in attendees_data.iterrows():
    to_email = row['Email']
    name = row['name']
    

    # Generate the email body
    body = body_template.format(name=name)

    # You can add personalized attachments here if needed
    personalized_attachments = common_attachments.copy()

    total_emails += 1

    if send_email_with_retry(to_email, subject, body, from_email, from_password, personalized_attachments):
        successful_emails += 1

    time.sleep(5)  # Delay between emails

    # Batch processing
    if total_emails % batch_size == 0:
        print(f"Batch limit reached. Pausing for {batch_delay / 60} minutes...")
        time.sleep(batch_delay)

# Final report
print(f"Total emails attempted: {total_emails}")
print(f"Total emails successfully sent: {successful_emails}")
