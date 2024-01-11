import os

# Load the Environment Variables
from dotenv import load_dotenv
load_dotenv()

# Access the environment variables for them to be usable
SPREADSHEET_ID = os.getenv('SPREADSHEET_ID')
PASSWORD_SHEET = os.getenv('PASSWORD_SHEET')

CREDENTIALS_PATH = os.getenv('CREDENTIALS_PATH')
GMAIL_TOKEN_PATH = os.getenv('GMAIL_TOKEN_PATH')
SHEETS_TOKEN_PATH = os.getenv('SHEETS_TOKEN_PATH')

EXCEPTION_EMAIL_ADDRESS = os.getenv('EXCEPTION_EMAIL_ADDRESS')
DOWNLOADS_FOLDER = os.getenv('DOWNLOADS_FOLDER')