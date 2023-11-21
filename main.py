from dotenv import load_dotenv
import os
import imaplib
import email
import re
from datetime import datetime
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import configparser
import paramiko

# Load environment variables from the .env file
load_dotenv()

# Construct the full path for the configuration file using environment variables
config_path = os.getenv("CONFIG_PATH")
config_full_path = os.path.join(os.getcwd(), config_path, "config.ini")

# Parse the configuration file for settings
config = configparser.ConfigParser()
config.read(config_full_path)

# Retrieve and process email recipients from the configuration file
try:
    EMAIL_RECIPIENTS = [email.strip() for email in config['EMAIL']['recipients'].split(',')]
except KeyError as e:
    print(f"Error: {e}. Check if 'EMAIL' section exists in the config file.")

# Load email server configuration from environment variables
hostname = os.getenv('EMAIL_HOSTNAME')
username = os.getenv('EMAIL_USERNAME')
password = os.getenv('EMAIL_PASSWORD')


# Function to upload files via SFTP
def sftp_upload_file(hostname, port, username, password, local_file, remote_path):
    """
    Uploads a file to a remote server via SFTP.

    :param hostname: SFTP server hostname
    :param port: SFTP server port
    :param username: SFTP username
    :param password: SFTP password
    :param local_file: Path to the local file to be uploaded
    :param remote_path: Remote path where the file should be uploaded
    """
    try:
        # Initialize and connect to the SSH client
        ssh_client = paramiko.SSHClient()
        ssh_client.set_missing_host_key_policy(paramiko.AutoAddPolicy())
        ssh_client.connect(hostname, port=port, username=username, password=password)

        # Open an SFTP session and upload the file
        with ssh_client.open_sftp() as sftp:
            sftp.put(local_file, remote_path + os.path.basename(local_file))
            print("File uploaded successfully")
    except Exception as e:
        print(f"Error during SFTP upload: {e}")
    finally:
        ssh_client.close()


# Function to send emails
def send_email(subject, body, to_emails):
    """
    Sends an email to specified recipients.

    :param subject: Email subject
    :param body: Email body content
    :param to_emails: List of email recipients
    """
    # Set up email details
    sender_email = os.getenv("SMTP_SENDER_EMAIL")
    sender_password = os.getenv("SMTP_SENDER_PASSWORD")
    message = MIMEMultipart()
    message["From"] = sender_email
    message["To"] = ", ".join(to_emails)
    message["Subject"] = subject
    message.attach(MIMEText(body, 'plain'))

    # Send the email
    try:
        server = smtplib.SMTP(os.getenv("SMTP_SERVER"), 587)
        server.starttls()
        server.login(sender_email, sender_password)
        server.sendmail(sender_email, to_emails, message.as_string())
        server.close()
        print("Email sent successfully")
    except Exception as e:
        print(f"Error sending email: {e}")


# Extract the body content from an email message
def get_email_body(msg):
    """
    Extracts the body content from an email message.

    The function handles both multipart and singlepart emails.
    :param msg: Email message object
    :return: Body content of the email
    """
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get('Content-Disposition'))
            if "attachment" not in content_disposition and content_type in ["text/plain", "text/html"]:
                return part.get_payload(decode=True).decode()
    return msg.get_payload(decode=True).decode()


# Check if an email message is a reply
def is_reply(email_message):
    """
    Determines if an email message is a reply.

    :param email_message: Email message object
    :return: True if it is a reply, False otherwise
    """
    if email_message['subject'].startswith('Re:'):
        return True
    return email_message.get('In-Reply-To') is not None or email_message.get('References') is not None


# Extract filename from the email subject
def extract_filename(subject):
    """
    Extracts a filename embedded within square brackets in the subject line.

    :param subject: Email subject line
    :return: Extracted filename or None if not found
    """
    match = re.search(r'\[([^\]]+)\]', subject)
    return match.group(1) if match else None


# Process the textual content of an email to extract relevant data
def process_email_text(text):
    """
    Processes the text content of an email to extract relevant data.

    The function splits the text into lines and searches each line for specific patterns.
    :param text: Raw text content of an email
    :return: A list of extracted data
    """
    lines = text.strip().split('\n')
    extracted_data = []
    for line in lines:
        match = re.search(r'^\s*(\d\s+\d{5}\s+\d{5}\s+\d)\s+P', line)
        if match:
            sku_upc = match.group(1).replace(' ', '')
            extracted_data.append(sku_upc)
    return extracted_data


# Write data to a CSV file
def write_to_csv(data, filename):
    """
    Writes provided data into a CSV file.

    Each row of data is written to the file in a specific format.
    :param data: List of data to be written
    :param filename: Name of the CSV file to be created
    """
    with open(filename, 'w') as file:
        file.write('SKU,UPC,EAN,QUANTITY_AVAILABLE,STATUS,\n')
        for item in data:
            file.write(f'{item},{item},,0,discontinued,\n')


# Fetch and process emails from an IMAP server
def fetch_emails():
    """
    Connects to an IMAP server, fetches unread emails, and processes them.

    The function extracts email content, processes it, writes data to a CSV file,
    and if required, uploads the file via SFTP.
    """
    try:
        mail = imaplib.IMAP4_SSL(hostname)
        mail.login(username, password)
        mail.select('inbox')
        print("trying email inbox")

        status, messages = mail.uid('search', None, 'UNSEEN')
        messages = messages[0].split()

        for uid in messages:
            status, data = mail.uid('fetch', uid, '(RFC822)')
            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    subject = msg['subject']

                    if is_reply(msg):
                        filename = extract_filename(subject)
                        if filename:
                            print(f"Reply successful and found filename: {filename}")
                            # Load SFTP details from environment variables
                            sftp_hostname = os.getenv("SFTP_HOSTNAME")
                            sftp_port = os.getenv("SFTP_PORT")
                            sftp_username = os.getenv("SFTP_USERNAME")
                            sftp_password = os.getenv("SFTP_PWORD")
                            local_file_path = os.getenv("LOCAL_FILE_PATH") + filename
                            remote_path = os.getenv("REMOTE_PATH")

                            # Upload the file via SFTP
                            sftp_upload_file(sftp_hostname, sftp_port, sftp_username, sftp_password, local_file_path, remote_path)
                    else:
                        body = get_email_body(msg)
                        extracted_data = process_email_text(body)

                        # Format the extracted data as CSV for email body
                        csv_data = 'SKU,UPC,EAN,QUANTITY_AVAILABLE,STATUS\n'
                        for item in extracted_data:
                            csv_data += f'{item},{item},,0,discontinued\n'

                        # Generate a unique filename and write data to a CSV file
                        timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
                        unique_filename = f'nordsvcp_{timestamp}.csv'
                        write_to_csv(extracted_data, unique_filename)
                        mail.uid('store', uid, '+FLAGS', '\\Seen')

                        # Prepare and send an email with the CSV data
                        outgoing_subject = f"DSVNord Closeout Approval: [{unique_filename}]"
                        outgoing_body = f"Filename: {unique_filename}\nhas been received.\nDoes this look correct?\n\n{csv_data}"
                        send_email(outgoing_subject, outgoing_body, EMAIL_RECIPIENTS)

        mail.close()
        mail.logout()

    except Exception as e:
        print(f'An error occurred: {e}')


fetch_emails()
