import smtplib 
import pandas as pd
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import os
import time

# Function to send email with attachment
def send_email(to_email, subject, body, from_email, from_password, attachment_path=None):
    try:
        # Create message container
        msg = MIMEMultipart()
        msg['From'] = from_email
        msg['To'] = to_email
        msg['Subject'] = subject

        # Attach the email body to the message (HTML format)
        msg.attach(MIMEText(body, 'html'))

        # Attach file (if any)
        if attachment_path:
            filename = os.path.basename(attachment_path)
            try:
                with open(attachment_path, "rb") as attachment:
                    # Instance of MIMEBase and named as part
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(attachment.read())

                # Encode into base64
                encoders.encode_base64(part)

                part.add_header('Content-Disposition', f"attachment; filename= {filename}")

                # Attach the instance 'part' to the message
                msg.attach(part)
            except Exception as file_error:
                print(f"Failed to attach file {attachment_path}: {str(file_error)}")

        # Connect to the server
        server = smtplib.SMTP('smtp.gmail.com', 587)  # Use Gmail's SMTP server and port
        server.starttls()

        # Login to your email account
        server.login(from_email, from_password)

        # Send email
        server.sendmail(from_email, to_email, msg.as_string())
        server.quit()

        # Success message
        print(f"Email successfully sent to {to_email}")
        return True  # Return success status

    except smtplib.SMTPException as e:
        print(f"SMTP error occurred: {e}. Failed to send email to {to_email}.")
    except Exception as e:
        print(f"An unexpected error occurred: {e}. Failed to send email to {to_email}.")
    return False  # Return failure status

# Function to send email with retry logic
def send_email_with_retry(to_email, subject, body, from_email, from_password, attachment_path=None, retries=3):
    for attempt in range(retries):
        if send_email(to_email, subject, body, from_email, from_password, attachment_path):
            return True
        print(f"Retrying... ({attempt + 1}/{retries})")
        time.sleep(5)  # Wait 5 seconds before retrying
    return False

# Read the CSV file
file_path = r'D:\Conclave.xlsx'  # Specify the correct path to your Excel file
attendees_data = pd.read_excel(file_path, engine='openpyxl')  # Use read_excel for .xlsx files


# Email credentials
from_email = ''  # Replace with your email
from_password = ''  # Replace with your app password

# Email content template with placeholders
subject = 'International Conclave on Next-Gen Higher Education - Ticket'
body_template = '''
Dear {name},<br><br>

Warm greetings from the Department of Higher Education, Government of Kerala!<br><br>

We are delighted to confirm your registration for the International Conclave on Next-Gen Higher Education. Thank you for joining us in this prestigious event dedicated to shaping the future of education.<br><br>

<b>Event Details</b><br>
Date: January 14–15, 2025<br>
Time: 9:00 am<br>
Venue: Cochin University of Science and Technology, Kochi<br>
Location Link: <a href="https://maps.app.goo.gl/pqdb3ymzjdaxNQ7R8">Seminar Complex, CUSAT</a><br><br>

Click here to access your ticket:{url}<br>This ticket contains a unique QR code that must be presented at the registration desk upon your arrival.<br><br>


Our team looks forward to welcoming you and ensuring your participation is seamless and enjoyable.<br><br>

Warm regards,<br>
Event Team<br>
Department of Higher Education, Government of Kerala
'''


# List of attachments (if any)
attachments = []  # Add paths to attachments if needed

# Initialize counters for total emails and successful emails
total_emails = 0
successful_emails = 0
batch_size = 10  # Number of emails per batch
batch_delay = 10 * 60  # 10 minutes delay between batches in seconds

# Iterate through the email list and send personalized emails in batches
for index, row in attendees_data.iterrows():
    to_email = row['email']
    name = row['name']
    url = row['ticket url']

    # Generate the email body with placeholders filled in
    body = body_template.format(name=name, url=url)
    total_emails += 1

    # Send email with all attachments
    for attachment_path in attachments:
        if send_email_with_retry(to_email, subject, body, from_email, from_password, attachment_path):
            successful_emails += 1

    # If no attachments, send email without any attachments
    if not attachments:
        if send_email_with_retry(to_email, subject, body, from_email, from_password):
            successful_emails += 1

    # Pause between each email
    time.sleep(5)  # Delay of 5 seconds between each email

    # Check if the batch size has been reached
    if total_emails % batch_size == 0:
        print(f"Batch limit reached. Pausing for {batch_delay / 60} minutes...")
        time.sleep(batch_delay)  # Pause before sending the next batch

# Print the final count of emails sent
print(f"Total emails attempted: {total_emails}")
print(f"Total emails successfully sent: {successful_emails}")
