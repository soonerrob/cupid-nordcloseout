from dotenv import load_dotenv
import os
import imaplib
import email
import re
from datetime import datetime

# Load environment variables
load_dotenv()

hostname = os.getenv('EMAIL_HOSTNAME')
username = os.getenv('EMAIL_USERNAME')
password = os.getenv('EMAIL_PASSWORD')


def process_email_text(text):
    """
    Processes the text extracted from an email.

    This function takes the raw text of an email, splits it into lines, and extracts
    the relevant data from each line. It specifically looks for data before the 'P' character,
    formats it by removing unnecessary spaces, and compiles a list of this data.
    """
    lines = text.strip().split('\n')
    extracted_data = []

    for line in lines:
        match = re.search(r'^\s*(\d\s+\d{5}\s+\d{5}\s+\d)\s+P', line)
        if match:
            sku_upc = match.group(1).replace(' ', '')
            extracted_data.append(sku_upc)

    return extracted_data


def write_to_csv(data, filename='nordsvcp.csv'):
    """
    Writes extracted data to a CSV file.

    This function takes the processed data and writes it into a CSV file.
    Each entry is written in a predefined format, including setting the quantity
    to 0 and status to 'discontinued'.
    """

    # Generate a unique filename based on the current date and time
    timestamp = datetime.now().strftime('%Y%m%d%H%M%S')
    unique_filename = f'nordsvcp_{timestamp}.csv'

    with open(unique_filename, 'w') as file:
        file.write('SKU,UPC,EAN,QUANTITY_AVAILABLE,STATUS,\n')
        for item in data:
            file.write(f'{item},{item},,0,discontinued,\n')


def get_email_body(msg):
    """
    Extracts the body of an email.

    Given an email message, this function parses it to extract the body.
    It handles both multipart and singlepart emails, and returns the
    text/plain or text/html part of the email.
    """
    if msg.is_multipart():
        for part in msg.walk():
            content_type = part.get_content_type()
            content_disposition = str(part.get('Content-Disposition'))

            if "attachment" not in content_disposition and content_type in ["text/plain", "text/html"]:
                return part.get_payload(decode=True).decode()
    else:
        return msg.get_payload(decode=True).decode()


def fetch_emails():
    """
    Fetches and processes emails from an IMAP server.

    This function connects to an IMAP server using the provided configuration,
    searches for unread emails, and processes each one. The processing involves
    extracting the email body, processing the text, writing the relevant data to a CSV file,
    and marking the email as read.
    """
    try:
        mail = imaplib.IMAP4_SSL(hostname)
        mail.login(username, password)
        mail.select('inbox')

        status, messages = mail.uid('search', None, 'UNSEEN')
        messages = messages[0].split()

        for uid in messages:
            status, data = mail.uid('fetch', uid, '(RFC822)')
            for response_part in data:
                if isinstance(response_part, tuple):
                    msg = email.message_from_bytes(response_part[1])
                    body = get_email_body(msg)
                    subject = msg['subject']
                    print(f'Subject: {subject}')
                    # print(f'Body: {body}')
                    extracted_data = process_email_text(body)
                    write_to_csv(extracted_data)
                    mail.uid('store', uid, '+FLAGS', '\\Seen')

        mail.close()
        mail.logout()

    except Exception as e:
        print(f'An error occurred: {e}')


fetch_emails()

