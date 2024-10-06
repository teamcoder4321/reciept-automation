import imaplib
import email
import os
import PyPDF2
import pytesseract
from PIL import Image
import re
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from google.oauth2.credentials import Credentials
from googleapiclient.discovery import build

# Email Fetching
def fetch_emails(email_address, password, imap_server):
    mail = imaplib.IMAP4_SSL(imap_server)
    mail.login(email_address, password)
    mail.select('inbox')
    
    _, search_data = mail.search(None, 'UNSEEN SUBJECT "receipt"')
    
    attachments = []
    for num in search_data[0].split():
        _, data = mail.fetch(num, '(RFC822)')
        _, bytes_data = data[0]
        
        email_message = email.message_from_bytes(bytes_data)
        for part in email_message.walk():
            if part.get_content_maintype() == 'multipart':
                continue
            if part.get('Content-Disposition') is None:
                continue
            
            file_name = part.get_filename()
            if file_name:
                if file_name.endswith('.pdf') or file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
                    file_data = part.get_payload(decode=True)
                    attachments.append((file_name, file_data))
    
    mail.close()
    mail.logout()
    return attachments

# Attachment Processing and Data Extraction
def process_attachment(file_name, file_data):
    if file_name.endswith('.pdf'):
        return extract_from_pdf(file_data)
    elif file_name.lower().endswith(('.png', '.jpg', '.jpeg')):
        return extract_from_image(file_data)
    else:
        raise ValueError(f"Unsupported file type: {file_name}")

def extract_from_pdf(file_data):
    pdf_reader = PyPDF2.PdfReader(file_data)
    text = ""
    for page in pdf_reader.pages:
        text += page.extract_text()
    return extract_details(text)

def extract_from_image(file_data):
    image = Image.open(file_data)
    text = pytesseract.image_to_string(image)
    return extract_details(text)

def extract_details(text):
    date_pattern = r'\d{2}/\d{2}/\d{4}'
    receipt_number_pattern = r'Receipt #:\s*(\d+)'
    vendor_pattern = r'Vendor:\s*(.+)'
    total_pattern = r'Total:\s*\$(\d+\.\d{2})'
    items_pattern = r'Items:(.*?)(?=Total:)'
    
    date = re.search(date_pattern, text)
    receipt_number = re.search(receipt_number_pattern, text)
    vendor = re.search(vendor_pattern, text)
    total = re.search(total_pattern, text)
    items = re.search(items_pattern, text, re.DOTALL)
    
    return {
        'date': date.group() if date else None,
        'receipt_number': receipt_number.group(1) if receipt_number else None,
        'vendor': vendor.group(1) if vendor else None,
        'total': total.group(1) if total else None,
        'items': items.group(1).strip() if items else None
    }

# Google Form Submission
def submit_to_google_form(form_id, details):
    creds = Credentials.from_authorized_user_file('path/to/token.json', ['https://www.googleapis.com/auth/forms.responses.write'])
    service = build('forms', 'v1', credentials=creds)
    
    body = {
        'responses': [
            {'textAnswers': {'answers': [{'value': str(value)}]}}
            for value in details.values()
        ]
    }
    
    service.forms().responses().create(formId=form_id, body=body).execute()

# Email Notification
def send_notification(sender_email, sender_password, recipient_email, details):
    message = MIMEMultipart()
    message['From'] = sender_email
    message['To'] = recipient_email
    message['Subject'] = 'New Receipt Processed'
    
    body = f"""
    A new receipt has been processed:
    
    Vendor: {details['vendor']}
    Total Amount: ${details['total']}
    Date: {details['date']}
    Receipt Number: {details['receipt_number']}
    
    The full details have been submitted to the Google Form.
    """
    
    message.attach(MIMEText(body, 'plain'))
    
    with smtplib.SMTP('smtp.gmail.com', 587) as server:
        server.starttls()
        server.login(sender_email, sender_password)
        server.send_message(message)

# Main Execution
def main():
    email_address = 'your_email@gmail.com'
    email_password = 'your_email_password'
    imap_server = 'imap.gmail.com'
    google_form_id = 'your_google_form_id'
    finance_team_email = 'finance@yourcompany.com'
    
    attachments = fetch_emails(email_address, email_password, imap_server)
    
    for file_name, file_data in attachments:
        try:
            details = process_attachment(file_name, file_data)
            submit_to_google_form(google_form_id, details)
            send_notification(email_address, email_password, finance_team_email, details)
            print(f"Processed: {file_name}")
        except Exception as e:
            print(f"Error processing {file_name}: {str(e)}")

if __name__ == "__main__":
    main()