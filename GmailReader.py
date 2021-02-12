from __future__ import print_function
import pickle
import os.path
import base64
from googleapiclient.discovery import build
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request


# If modifying these scopes, delete the file token.pickle.
SCOPES = ['https://www.googleapis.com/auth/gmail.modify']

# Get the ID of a specific label
def get_label_id(label_name, labels):
    for label in labels:
        if label['name'].lower() == label_name.lower():
            return label['id']

# Get unread messages with specific params
def get_unread_messages(service, label):
    try:
        return service.users().messages().list(userId='me', labelIds=[label, 'INBOX']).execute()
    except Exception as error:
        msg = "Unable to get messages with the label ID %s. Error: %s" % (label, error)
        print(msg)

# Retrieve ids for all unread messages
def get_unread_message_ids(service, label_id):
    try:
        unread_messages = str(get_unread_messages(service, label_id))
        if unread_messages == "{'resultSizeEstimate': 0}":
            return []
        else:
            unread_converted = unread_messages.split(
                '[')[1].lstrip().split(']')[0]
            unread_converted = unread_converted.split()
            unread_ids = []
            count = 0
            for item in unread_converted:
                item = item.split("'")[1].lstrip().split("'")[0]
                if count == 1:
                    unread_ids.append(item)
                if count == 3:
                    count = 0
                else:
                    count += 1

            return unread_ids
    except Exception as error:
        msg = "Error getting new message IDs: %s" % error
        print(msg)

# Look for attachments in unread emails. If an attachment found, - download and marked email as read.
def get_msg_attachments(service, unread_msg_ids):
    try:
        for i in range(len(unread_msg_ids)):
            message = service.users().messages().get(userId='me', id=unread_msg_ids[i]).execute()
            parts = [message['payload']]
            while parts:
                part = parts.pop()
                if part.get('parts'):
                    parts.extend(part['parts'])
                if 'attachmentId' in part['body']:
                    attachment = service.users().messages().attachments().get(
                        userId='me', messageId=message['id'], id=part['body']['attachmentId']
                    ).execute()
                    file_data = base64.urlsafe_b64decode(
                        attachment['data'].encode('UTF-8'))
                    if file_data:
                        print("There is an attachment. Will download shortly..")
                        filename = ""
                        if part['filename'] == '':
                            filename = "noname"
                        else:
                            filename = part['filename']
                        current_dir = os.getcwd() + "//"
                        path = ''.join([current_dir, filename])
                        with open(path, 'wb') as f:
                            f.write(file_data)
                        print("An attachment with %s name was saved." % (filename))
                        service.users().messages().modify(userId='me', id=unread_msg_ids[i], body={'removeLabelIds': ['UNREAD']}).execute()
    except Exception as error:
        print("Can't get an attachment. Error: %s" % (error))

def main():
    """Shows basic usage of the Gmail API.
    Lists the user's Gmail labels.
    """
    creds = None
    # The file token.pickle stores the user's access and refresh tokens, and is
    # created automatically when the authorization flow completes for the first
    # time.
    token_name = 'token.pickle'
    if os.path.exists(token_name):
        with open(token_name, 'rb') as token:
            creds = pickle.load(token)
    # If there are no (valid) credentials available, let the user log in.
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(
                'credentials.json', SCOPES)
            creds = flow.run_local_server(port=0)
        # Save the credentials for the next run
        with open(token_name, 'wb') as token:
            pickle.dump(creds, token)

    service = build('gmail', 'v1', credentials=creds)

    # Call the Gmail API
    results = service.users().labels().list(userId='me').execute()
    labels = results.get('labels', [])

    if not labels:
        print('There are no labels found.')
    else:
        unread = get_label_id('UNREAD', labels)
        message_ids = get_unread_message_ids(service, unread)
        if message_ids:
            get_msg_attachments(service, message_ids)
        else:
            print('Was unable to find undread messages')

if __name__ == '__main__':
    main()