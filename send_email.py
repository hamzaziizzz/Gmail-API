'''
This module sends emails with attachments to the participants
Reference - https://developers.google.com/gmail/api/quickstart/python
'''

from email.mime.application import MIMEApplication
from email.mime.multipart import MIMEMultipart
import os
from google.auth.transport.requests import Request
from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from email.mime.text import MIMEText
import base64

# If modifying these scopes, delete the file token.json.
SCOPES = ['https://www.googleapis.com/auth/gmail.send']
def aunthentication():
    creds = None
    # The file token.json stores the user's access and refresh tokens, and is created automatically when the authorization flow completes for the first time.
    if os.path.exists('token.json'):
        creds = Credentials.from_authorized_user_file('token.json', SCOPES)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file('client_secrets.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open('token.json', 'w') as token:
            token.write(creds.to_json())
    return creds

def prepare_and_send_email(recipient, subject, message_text):
    """Prepares and send email with attachment to the participants"""
    creds = aunthentication()
    try:
        # Call the Gmail API
        service = build('gmail', 'v1', credentials=creds)
        # create message
        # TODO: Put your Email-ID from which you have created your Google Cloud account
        msg = create_message('example@domain.com', recipient, subject, message_text) 
        send_message(service, 'me', msg)
    except HttpError as error:
        # TODO (developer) - Handle errors from gmail API.
        print(f'An error occurred: {error}')

def create_message(sender, to, subject, message_text):
    """
    Create a message for an email.

    Args:
        sender: Email address of the sender.
        to: Email address of the receiver.
        subject: The subject of the email message.
        message_text: The text of the email message.

    Returns:
        An object containing a base64url encoded email object.
    """
    message = MIMEMultipart()
    message['from'] = sender
    message['to'] = to
    message['subject'] = subject

    message.attach(MIMEText(message_text))

    # attachment
    attachment = ""     # TODO: Put your attachment's path in quotes
    with open(attachment, 'rb') as fp:
        attachment_data = MIMEApplication(fp.read(), Name="EmailAttachment")
    message.attach(attachment_data)

    return {'raw':base64.urlsafe_b64encode(message.as_string().encode()).decode()}

def send_message(service, user_id, message):
    """
    Send an email message.

    Args:
        service: Authorized Gmail API service instance.
        user_id: User's email address. The special value "me"
        can be used to indicate the authenticated user.
        message: Message to be sent.

    Returns:
        Sent Message.
    """
    try:
        message = (service.users().messages().send(userId=user_id, body=message).execute())
        print('Message Id: %s' % message['id'])
        return message
    except HttpError as error:
        print('An error occurred: %s' % error)


if __name__ == '__main__':
    # TODO: Put the recepient email in place of 'example@email.com' and change the subject and message text as per your requirement
    prepare_and_send_email('example@email.com', 'Greeting from ABESIT', 'This is a test email for our upcoming app')
