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


# === COPY TEMPLATE ===
def copy_template(drive_service, template_id, new_name):
    copied = drive_service.files().copy(fileId=template_id, body={'name': new_name}).execute()
    print(f"Copied template: {copied['name']}")
    return copied['id']


# === STEP 1 - Group the services per Quote-ID in GroupedQuotes ===
def group_rows_by_quote_id(data, header):
    """Group rows by Quote ID and return a structured dictionary."""

    grouped = {}

    for row in data[1:]:  # Skip header
        row += [''] * (len(header) - len(row))  # Pad short rows
        row_data = dict(zip(header, row))

        quote_id = row_data.get("Quote ID", "").strip()
        if not quote_id:
            continue

        if quote_id not in grouped:
            grouped[quote_id] = {
                'Quote ID': quote_id,
                'Date': row_data.get('Date', ''),
                'Client Name': row_data.get('Client Name', ''),
                'Email': row_data.get('Email', ''),
                'Organization': row_data.get('Organization', ''),
                'Notes': row_data.get('Notes', ''),
                'rows': [],
                'Grand Total': 0.0
            }

        # Normalize keys (underscored) for compatibility with Document Studio
        service_data = {
            'Service_Type': row_data.get('Service Type', ''),
            'Language_Pair': row_data.get('Language Pair', ''),
            'Modality': row_data.get('Modality', ''),
            'Word_Count': row_data.get('Word Count', ''),
            'Duration_hrs': row_data.get('Duration (hrs)', ''),
            'Rate': row_data.get('Rate', ''),
            'Details': row_data.get('Details', ''),
            'Total': row_data.get('Total', '')
        }

        grouped[quote_id]['rows'].append(service_data)

        # Accumulate total
        try:
            grouped[quote_id]['Grand Total'] += float(row_data.get('Total', '0'))
        except ValueError:
            pass

    return list(grouped.values())
