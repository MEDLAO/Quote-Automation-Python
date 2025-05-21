from gdoctableapppy import gdoctableapp
from google.oauth2 import service_account


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'
DOCUMENT_ID = '18OVjzAQnTZKqhFmiaemIaHaw0QV7G8fPgMDGnwh-Wpg'
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']

# === AUTHENTICATE ===
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)


print(f"Using document: https://docs.google.com/document/d/{DOCUMENT_ID}")

resource = {
    "service_account": creds,
    "documentId": DOCUMENT_ID,
}
res = gdoctableapp.GetTables(resource)
print(res)


