import os
import pickle
import globals
import base64
import pyzipper
import helper
import csv

# Google API utilities
from googleapiclient.discovery import build
from google.auth.transport.requests import Request
from google_auth_oauthlib.flow import InstalledAppFlow
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart

# Google API helper functions

def authenticate_gmail_api(GMAIL_TOKEN_PATH, GMAIL_CREDENTIALS_PATH):
    GMAIL_SCOPES = ['https://www.googleapis.com/auth/gmail.readonly', 'https://www.googleapis.com/auth/gmail.modify']
    creds = None
    if os.path.exists(GMAIL_TOKEN_PATH):
        with open(GMAIL_TOKEN_PATH, "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(GMAIL_CREDENTIALS_PATH, GMAIL_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(GMAIL_TOKEN_PATH, "wb") as token:
            pickle.dump(creds, token)
    return build('gmail', 'v1', credentials=creds)

def authenticate_sheets_api(SHEETS_TOKEN_PATH, SHEETS_CREDENTIALS_PATH):
    SHEETS_SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
    creds = None
    if os.path.exists(SHEETS_TOKEN_PATH):
        with open(SHEETS_TOKEN_PATH, "rb") as token:
            creds = pickle.load(token)
    if not creds or not creds.valid:
        if creds and creds.expired and creds.refresh_token:
            creds.refresh(Request())
        else:
            flow = InstalledAppFlow.from_client_secrets_file(SHEETS_CREDENTIALS_PATH, SHEETS_SCOPES)
            creds = flow.run_local_server(port=0)
        with open(SHEETS_TOKEN_PATH, "wb") as token:
            pickle.dump(creds, token)
    return build('sheets', 'v4', credentials=creds)

# Authenticate Gmail API, Google Drive API, and Google Sheets API
gmail_service = authenticate_gmail_api(globals.GMAIL_TOKEN_PATH, globals.CREDENTIALS_PATH)
sheets_service = authenticate_sheets_api(globals.SHEETS_TOKEN_PATH, globals.CREDENTIALS_PATH)

# General helper functions

# Creates a folder if it does the folder path provided is not located 
def createFolderIfDoesNotExist(path):
    if not os.path.exists(path):
        os.makedirs(path)
        print(f"Folder created: {path}")

# Sends error messages to the email given if an error occers
def sendErrorEmail(EXCEPTION_EMAIL_ADDRESS, organisation):
    message = MIMEMultipart()
    message['to'] = EXCEPTION_EMAIL_ADDRESS
    message['subject'] = 'Error occurred in ' + organisation + ' Standard Report Automation.'
    body_message = 'An error has occurred for ' + organisation + ' Standard Report Automation, and the code did not manage to finish running. Please fix the issue, thank you.'
    body = MIMEText(body_message, 'plain')
    message.attach(body)
    raw_message = base64.urlsafe_b64encode(message.as_bytes()).decode('utf-8')

    # Send the message using the Gmail API
    response = gmail_service.users().messages().send(userId='me', body={'raw': raw_message}).execute()
    print('Message sent. Message ID:', response['id'])

# Append new rows of data onto spreadsheet, dataArray should be in 2d-array
def appendToSpreadsheet(dataArray, spreadsheet_id, sheetname):
    sheets_service.spreadsheets().values().append(
        spreadsheetId=spreadsheet_id,
        range=sheetname,
        body={'values':dataArray},
        valueInputOption='USER_ENTERED'
        # USER_ENTERED so that it can detect formats such as number automatically
    ).execute()
    print("Appended data to sheet.")

# Get messages of that label by searching for the label of that specified name
def get_messages_with_label(label_name):        
    labels = helper.gmail_service.users().labels().list(userId='me').execute()
    label_id = None
    for label in labels['labels']:
        if label['name'] == label_name:
            label_id = label['id']
            break

    if label_id:
        messages = helper.gmail_service.users().messages().list(userId='me', labelIds=[label_id]).execute()
        if 'messages' in messages:
            return messages['messages']
        else:
            print(f"No messages found in the label '{label_name}'.")
            return []
    else:
        print(f"Label '{label_name}' not found.")
        return []

# Download attachments for email located using email id, into the folder path given, and file path of the attachment will be returned
def download_attachments(email_id, download_folder_path):
    email = helper.gmail_service.users().messages().get(userId='me', id=email_id).execute()
    parts = email['payload']['parts']  # Get email parts
    downloaded_files = []

    for part in parts:
        filename = part['filename']
        print(filename)
        if filename:
            att_id = part['body']['attachmentId']
            att = helper.gmail_service.users().messages().attachments().get(userId='me', messageId=email_id, id=att_id).execute()
            data = att['data']

            file_data = base64.urlsafe_b64decode(data.encode('UTF-8'))
            file_path = os.path.join(download_folder_path, filename)

            with open(file_path, 'wb') as f:
                f.write(file_data)
                print("File downloaded.")
                downloaded_files.append(file_path)

    return downloaded_files

# Unzip password protected file
def unzip_password_protected_zip(zip_file_path, password, output_dir):
    try:
        with pyzipper.AESZipFile(zip_file_path, 'r', compression=pyzipper.ZIP_LZMA) as z:
            z.pwd = password.encode('utf-8')
            z.extractall(output_dir)
            file_list = z.namelist()
        if len(file_list) == 1:
            extracted_file_path = os.path.join(output_dir, file_list[0])
            print("Extraction successful.")
            return extracted_file_path
        
    except pyzipper.BadZipFile:
        print("Invalid zip file.")
    except RuntimeError as e:
        print(f"Error: {e}")
    except Exception as e:
        print(f"An error occurred: {e}")

# Returns 2-d array of the data
def read_csv_file(file_path):
    try:
        with open(file_path, 'r', newline='') as csv_file:
            csv_reader = csv.reader(csv_file)
            csv_data = [row for row in csv_reader]
        return csv_data
    except Exception as e:
        print(f"Error reading CSV file: {e}")
        return []
    
# Filter only wanted information, and returns a 2-d array
def extractDataFrom2dArray(wholeSheetData):
    data = []
    count = 0
    number_of_rows = 0
    start = 0

    for row in wholeSheetData:
        if start == 0:
            if row[0] == 'Total records counts':
                number_of_rows = row[1]
                if number_of_rows == 0:
                    return None
            if row[0] == 'Product':
                start = 1
        else:
            if count == number_of_rows:
                return data
            else:
                row_data = []
                for element in row:
                    row_data.append(element)
                data.append(row_data)
                count += 1
    return data

# Obtain 2-d array  of data in spreadsheet
def read_data_from_spreadsheet(spreadsheet_id, range_name):
    result = sheets_service.spreadsheets().values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    values = result.get('values', [])
    return values

# Get recipient of email by its id
def get_recipient_of_email(message_id):
    message = helper.gmail_service.users().messages().get(userId='me', id=message_id).execute()
    headers = message['payload']['headers']
    for header in headers:
        if header['name'] == 'To':
            # Assume that the recipient is only one
            recipient = header['value']
            return recipient

# Get label id by its name
def get_label_id_by_name(label_name):
    try:
        labels_result = gmail_service.users().labels().list(userId='me').execute()
        labels = labels_result.get('labels', [])
        for label in labels:
            if label['name'] == label_name:
                return label['id']
        return None  # Return None if the label name is not found
    except Exception as e:
        print(f"Error retrieving label ID: {e}")
        return None
    
# Adds a label to a message
def add_label_to_message(message_id, label_name):
    label = {'addLabelIds': [get_label_id_by_name(label_name)]}
    message = gmail_service.users().messages().modify(userId='me', id=message_id, body=label).execute()
    return message

# Removes a label from the message
def remove_label_from_message(message_id, label_name):
    label = {'removeLabelIds': [get_label_id_by_name(label_name)]}
    message = gmail_service.users().messages().modify(userId='me', id=message_id, body=label).execute()
    return message