from google.oauth2 import service_account
from googleapiclient.discovery import build


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


def authenticate_gdoc(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('docs', 'v1', credentials=credentials).documents()


def authenticate_drive(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('drive', 'v3', credentials=credentials).files()


def insert_text_into_template_copy(template_id, new_title, text, docs_service, drive_service):
    """
    Makes a copy of the template and inserts text at the top of the copy.
    """
    copied_file = drive_service.copy(
        fileId=template_id,
        body={'name': new_title}
    ).execute()
    new_doc_id = copied_file['id']

    requests = [
        {
            'insertText': {
                'location': {'index': 1},
                'text': text + '\n'
            }
        }
    ]

    docs_service.batchUpdate(
        documentId=new_doc_id,
        body={'requests': requests}
    ).execute()

    print(f"New document created: https://docs.google.com/document/d/{new_doc_id}")


if __name__ == '__main__':
    docs_service = authenticate_gdoc(SERVICE_ACCOUNT_FILE, SCOPES)
    drive_service = authenticate_drive(SERVICE_ACCOUNT_FILE, SCOPES)
    insert_text_into_template_copy(
        template_id=TEMPLATE_DOC_ID,
        new_title='Test Quote Doc',
        text='Hello from Python!',
        docs_service=docs_service,
        drive_service=drive_service
    )
