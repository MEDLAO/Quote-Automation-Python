from google.oauth2 import service_account
from googleapiclient.discovery import build
import json


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'
SCOPES = [
    'https://www.googleapis.com/auth/documents',
    'https://www.googleapis.com/auth/drive',
    'https://www.googleapis.com/auth/spreadsheets.readonly'
]

TEMPLATE_DOC_ID = '18OVjzAQnTZKqhFmiaemIaHaw0QV7G8fPgMDGnwh-Wpg'
SPREADSHEET_ID = '1wiAQXkSvcOS8QdLeST2AmjsaV03_bS-1dIM3XpiNNq'
SHEET_NAME = 'GroupedQuotes'


# === AUTH FUNCTIONS ===
def authenticate_gsheet(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('sheets', 'v4', credentials=credentials).spreadsheets()


def authenticate_gdoc(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('docs', 'v1', credentials=credentials).documents()


def authenticate_drive(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('drive', 'v3', credentials=credentials).files()


# === BUILD SERVICES BLOCK ===
def build_services_block(services_json):
    try:
        services = json.loads(services_json)
    except json.JSONDecodeError:
        return "Invalid service data."

    lines = []
    for i, service in enumerate(services, start=1):
        lines.append(
            f"{i}. {service.get('Service_Type', '')} | "
            f"{service.get('Language_Pair', '')} | "
            f"{service.get('Modality', '')}"
        )
    return '\n'.join(lines)


# === MAIN ===
def main():
    # Authenticate
    gsheet = authenticate_gsheet(SERVICE_ACCOUNT_FILE, SCOPES)
    gdoc = authenticate_gdoc(SERVICE_ACCOUNT_FILE, SCOPES)
    gdrive = authenticate_drive(SERVICE_ACCOUNT_FILE, SCOPES)

    # Read first row of data from spreadsheet (after header)
    result = gsheet.values().get(
        spreadsheetId=SPREADSHEET_ID,
        range=f"{SHEET_NAME}!A1:Z2"
    ).execute()
    values = result.get('values', [])

    if len(values) < 2:
        print("No data found.")
        return

    header = values[0]
    row = values[1] + [''] * (len(header) - len(values[1]))
    quote_data = dict(zip(header, row))

    # Copy the template
    copied_file = gdrive.copy(
        fileId=TEMPLATE_DOC_ID,
        body={'name': f"Quote_{quote_data['Quote ID']}"}
    ).execute()
    new_doc_id = copied_file.get('id')

    # Prepare replacements
    replacements = {
        '{{Quote ID}}': quote_data.get('Quote ID', ''),
        '{{Date}}': quote_data.get('Date', ''),
        '{{Client Name}}': quote_data.get('Client Name', ''),
        '{{Email}}': quote_data.get('Email', ''),
        '{{Organization}}': quote_data.get('Organization', ''),
        '{{Notes}}': quote_data.get('Notes', ''),
        '{{Grand Total}}': quote_data.get('Grand Total', ''),
        '{{Services}}': build_services_block(quote_data.get('Services', '[]'))
    }

    # Build replacement requests
    requests = []
    for placeholder, value in replacements.items():
        requests.append({
            'replaceAllText': {
                'containsText': {'text': placeholder, 'matchCase': True},
                'replaceText': value
            }
        })

    # Apply to new doc
    gdoc.batchUpdate(
        documentId=new_doc_id,
        body={'requests': requests}
    ).execute()

    print(f"Document created: https://docs.google.com/document/d/{new_doc_id}")


if __name__ == '__main__':
    main()
