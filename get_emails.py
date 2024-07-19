import os
import base64
import re
import pickle
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError

# Define the scope
SCOPES = ['https://www.googleapis.com/auth/gmail.readonly']

# Path to the token.pickle file
TOKEN_PATH = 'token.pickle'
CREDENTIALS_PATH = '/path/to/credentials.json'
EXCEL_PATH = '/path/to/excel/file.xlsx'


# Authenticate and initialize the Gmail API client
creds = None
if os.path.exists(TOKEN_PATH):
    os.remove(TOKEN_PATH)

if os.path.exists(TOKEN_PATH):
    with open(TOKEN_PATH, 'rb') as token:
        creds = pickle.load(token)

if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file(CREDENTIALS_PATH, SCOPES)
        creds = flow.run_local_server(port=0)
    with open(TOKEN_PATH, 'wb') as token:
        pickle.dump(creds, token)

service = build('gmail', 'v1', credentials=creds)

def get_label_id(label_name):
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])
    for label in labels:
        if label['name'] == label_name:
            return label['id']
    return None

def extract_emails_from_text(text):
    email_regex = r'[a-zA-Z0-9_.+-]+@[a-zA-Z0-9-]+\.[a-zA-Z0-9-.]+'
    return [email for email in re.findall(email_regex, text) if email != 'user email']

def process_message_part(part, valid_emails_set, failed_emails):
    if 'body' in part and 'data' in part['body']:
        body_data = part['body']['data']
        email_body = base64.urlsafe_b64decode(body_data).decode('utf-8')
        found_emails = extract_emails_from_text(email_body)
        for email in found_emails:
            if email in valid_emails_set:
                failed_emails.add(email)

        # Check for nested "Forwarded message" sections
        forwarded_emails = extract_emails_from_forwarded(email_body)
        for email in forwarded_emails:
            if email in valid_emails_set:
                failed_emails.add(email)

    if 'parts' in part:
        for sub_part in part['parts']:
            process_message_part(sub_part, valid_emails_set, failed_emails)

def extract_emails_from_forwarded(body):
    forwarded_regex = r'Forwarded message -+[\s\S]+?From: [^\n]+?[\s\S]+?To: ([^\n]+)'
    matches = re.findall(forwarded_regex, body, re.IGNORECASE)
    extracted_emails = []
    for match in matches:
        emails = extract_emails_from_text(match)
        extracted_emails.extend(emails)
    return extracted_emails

def get_failed_emails(returned_label_name, retired_label_name, valid_emails_set):
    query = f'label:{returned_label_name} -label:{retired_label_name}'
    print(f"Constructed query: {query}")

    failed_emails = set()
    page_token = None

    while True:
        try:
            results = service.users().messages().list(userId='me', q=query, pageToken=page_token).execute()
        except HttpError as error:
            print(f"An error occurred: {error}")
            break
        
        messages = results.get('messages', [])
        print(f"Number of emails found with the query: {len(messages)}")

        if not messages:
            print("No messages found with the given query.")
            break

        for message in messages:
            try:
                msg = service.users().messages().get(userId='me', id=message['id'], format='full').execute()
            except HttpError as error:
                print(f"An error occurred fetching the message: {error}")
                continue

            payload = msg.get('payload', {})
            headers = payload.get('headers', [])
            from_address = None
            to_address = None

            for header in headers:
                if header['name'] == 'From':
                    from_address = header['value']
                if header['name'] == 'To':
                    to_address = header['value']

            if from_address and 'cwhsiao@smartecstore.com' in from_address and to_address:
                if to_address in valid_emails_set:
                    failed_emails.add(to_address)

            parts = payload.get('parts', [])
            if not parts:
                parts = [payload]

            for part in parts:
                process_message_part(part, valid_emails_set, failed_emails)

        page_token = results.get('nextPageToken')
        if not page_token:
            break

    return list(failed_emails)

# Load the list of valid emails from the Excel file
df = pd.read_excel(EXCEL_PATH)
valid_emails_set = set(df['Email'].dropna().unique())

# Define the label names
returned_label_name = "Returned Mails"
retired_label_name = "Retired"

# Get the label IDs for logging purposes (optional)
returned_label_id = get_label_id(returned_label_name)
retired_label_id = get_label_id(retired_label_name)

if returned_label_id and retired_label_id:
    print(f"Returned Mails Label ID: {returned_label_id}")
    print(f"Retired Label ID: {retired_label_id}")

    # Get failed emails
    failed_emails = get_failed_emails(returned_label_name, retired_label_name, valid_emails_set)

    # Print the failed email addresses
    if failed_emails:
        for email in failed_emails:
            print(email)
    else:
        print("No failed emails collected.")

    # Optionally, save the failed email addresses to a file
    with open('failed_email_set_3.txt', 'w') as f:
        for email in failed_emails:
            f.write(email + '\n')

    print(f'Total failed emails collected: {len(failed_emails)}')
else:
    if not returned_label_id:
        print("Label 'Returned Mails' not found.")
    if not retired_label_id:
        print("Label 'retired' not found.")
