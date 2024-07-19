import os
import base64
import pandas as pd
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
import pickle

# Define the scope
SCOPES = ['https://www.googleapis.com/auth/gmail.send']

# Load the Excel file
df = pd.read_excel('path to excel/all_emails.xlsx')

# Authenticate and initialize the Gmail API client
creds = None
if os.path.exists('token.pickle'):
    with open('token.pickle', 'rb') as token:
        creds = pickle.load(token)

# If there are no valid credentials available, let the user log in.
if not creds or not creds.valid:
    if creds and creds.expired and creds.refresh_token:
        creds.refresh(Request())
    else:
        flow = InstalledAppFlow.from_client_secrets_file('path to credentials/credentials.json', SCOPES)
        creds = flow.run_local_server(port=0)
    # Save the credentials for the next run
    with open('token.pickle', 'wb') as token:
        pickle.dump(creds, token)

service = build('gmail', 'v1', credentials=creds)

# Email content with CSS styling
subject = 'Subject Line'
html_content = """
<html>
Insert Content for Email
</html>
"""


# Paths to the images
image_paths = {
    "image1": "picture1.jpeg",
    "image2": "picture2.jpeg",
    "image3": "picture3.jpeg",
    "image4": "picture4.jpeg",
    "image5": "picture5.jpeg",
    "image6": "picture6.jpeg"
}

# Function to attach an image to the email

def attach_image(msg, image_cid, image_path):
    with open(image_path, 'rb') as img:
        mime = MIMEBase('image', 'jpeg', filename=os.path.basename(image_path))
        mime.add_header('Content-Disposition', 'attachment', filename=os.path.basename(image_path))
        mime.add_header('X-Attachment-Id', image_cid)
        mime.add_header('Content-ID', f'<{image_cid}>')
        mime.set_payload(img.read())
        encoders.encode_base64(mime)
        msg.attach(mime)

# Iterate through the DataFrame rows
batch_size = 100
email_count = 0
max_emails = 1950 # to avoid typical max email limit of 2000
email_start = ___ # set starting point

for start in range(email_start, min(email_start + max_emails, len(df)), batch_size):
    end = start + batch_size
    batch_df = df[start:end]

    for index, row in batch_df.iterrows():
        first_name = row['First Name']
        middle_name = row['Middle Name']
        last_name = row['Last Name']
        email = row['Email']

        # Customize the email body with placeholders
        email_body = html_content.format(
            FirstName=first_name,
            MiddleName=middle_name if pd.notna(middle_name) else '',
            LastName=last_name
        )

        # Create the email
        message = MIMEMultipart()
        message['to'] = email
        message['from'] = 'cwhsiao@smartecstore.com'
        message['subject'] = subject 

        msg = MIMEText(email_body, 'html')
        message.attach(msg)

        # Attach images
        for image_cid, image_path in image_paths.items():
            attach_image(message, image_cid, image_path)

        raw = base64.urlsafe_b64encode(message.as_bytes()).decode()

        # Send the email
        try:
            message = (service.users().messages().send(userId='me', body={'raw': raw}).execute())
            print(f'Email sent to {email}')
        except Exception as e:
            print(f'An error occurred: {e}')
        email_count += 1    
        if email_count >= max_emails:
          email_start += max_emails
          break
    
    
    print(email_count)
    # Wait to avoid hitting the sending limit

print('Emails sent successfully!')

