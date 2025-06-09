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
