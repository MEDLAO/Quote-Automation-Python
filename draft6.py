from googleapiclient.discovery import build
from google.oauth2 import service_account


# Replace with your service account JSON key file and desired document ID
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'
DOCUMENT_ID = '151-qGyh7xoTk6PWnIo-aTKHJHdJ6ZH6xfENr0F3Wrkk'

# Authenticate with Google Drive API
SCOPES = ['https://www.googleapis.com/auth/drive']
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES
)
gdrive = build('drive', 'v3', credentials=creds)


# Function to share the document
def share_document(gdrive, doc_id, anyone=True, email=None):
    body = {
        'type': 'anyone' if anyone else 'user',
        'role': 'writer',
    }
    if email:
        body['type'] = 'user'
        body['emailAddress'] = email

    permission = gdrive.permissions().create(
        fileId=doc_id,
        body=body,
        fields='id'
    ).execute()
    print(f"Permission granted: {permission}")


# Call the function to test
share_document(gdrive, DOCUMENT_ID)


def share_documents_from_sheet(service_account_file, spreadsheet_id, range_name, anyone=True, email=None):
    # Authenticate Sheets and Drive
    creds = service_account.Credentials.from_service_account_file(service_account_file)
    sheet_service = build('sheets', 'v4', credentials=creds)
    drive_service = build('drive', 'v3', credentials=creds)

    # Read data from the specified column
    result = sheet_service.spreadsheets().values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()
    values = result.get('values', [])

    if not values:
        print("No document links found.")
        return

    for row in values:
        if not row:
            continue
        doc_url = row[0]
        if "id=" in doc_url:
            doc_id = doc_url.split("id=")[-1]
        elif "document/d/" in doc_url:
            doc_id = doc_url.split("document/d/")[1].split("/")[0]
        else:
            print(f"Invalid URL format: {doc_url}")
            continue

        # Build permission body
        body = {
            'type': 'anyone' if anyone else 'user',
            'role': 'writer'
        }
        if email:
            body['type'] = 'user'
            body['emailAddress'] = email

        # Share the doc
        try:
            drive_service.permissions().create(
                fileId=doc_id,
                body=body,
                fields='id'
            ).execute()
            print(f"Shared: https://docs.google.com/document/d/{doc_id}")
        except Exception as e:
            print(f"Error sharing {doc_id}: {e}")
