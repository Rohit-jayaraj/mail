import smtplib
import ssl
import os
import csv
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import COMMASPACE
from email import encoders

port = 587  # For starttls
smtp_server = "smtp.office365.com"
sender_email = "<email>"  # Enter your sender email address
password = "<password>"  # Enter your sender email password
data_file = "mail_att.csv"  # Enter the path to your data source file

# Read the data from the CSV file
with open(data_file, "r") as f:
    reader = csv.reader(f)
    next(reader)  # Skip the header row
    for row in reader:
        recipient_email = row[0]  # Get the recipient email address
        attachment_file = row[1]  # Get the filename of the attachment
        subject = f"Email with attachment: {os.path.basename(attachment_file)}"
        body = f"Hello,\n\nPlease see the attached file: {os.path.basename(attachment_file)}"

        # Create a multipart message object and add the email content
        message = MIMEMultipart()
        message['From'] = sender_email
        message['To'] = recipient_email
        message['Subject'] = subject
        message.attach(MIMEText(body, 'plain'))

        # Add the attachment to the message
        with open(attachment_file, "rb") as f:
            attachment = MIMEBase("application", "octet-stream")
            attachment.set_payload(f.read())
            encoders.encode_base64(attachment)
            attachment.add_header(
                "Content-Disposition",
                f"attachment; filename={os.path.basename(attachment_file)}",
            )

        # Add the attachment to the message
        message.attach(attachment)

        # Send the email
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, port) as server:
            server.starttls(context=context)
            server.login(sender_email, password)
            server.sendmail(sender_email, recipient_email, message.as_string())

        print(f"Email with attachment sent to {recipient_email} successfully!")
