from googleapiclient.discovery import build  # Import the Google API client library to build service objects
from google.oauth2 import service_account  # Import Google OAuth2 library to handle authentication
import json


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'  # Path to the service account credentials JSON file
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets'
]
SPREADSHEET_ID_SOURCE = '1wiAQXkSvcOS8QdLeST2AmjsaV03_bS-1dIM3XpiNNq0'  # ID of the Google Sheet
SPREADSHEET_ID_TARGET = ''
TEMPLATE_DOC_ID = 'your-template-doc-id'
RANGE_NAME = 'Quotes!A1:Z'  # Range of data to read from the sheet


# === AUTHENTICATION ===
def authenticate_gsheet(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('sheets', 'v4', credentials=credentials).spreadsheets()


def authenticate_gdoc(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('docs', 'v1', credentials=credentials).documents()


def authenticate_drive(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('drive', 'v3', credentials=credentials).files()

# === STEP 1 - Group the services per Quote-ID in GroupedQuotes ===
