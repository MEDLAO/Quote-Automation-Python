from gdoctableapppy import gdoctableapp
from google.oauth2 import service_account


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'
DOCUMENT_ID = '1tp8b09JT4Q23sq4FJAsCDjXz0udlK9cfcKb-m6vwt7Y'
              # '18OVjzAQnTZKqhFmiaemIaHaw0QV7G8fPgMDGnwh-Wpg'
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
DUMMY_ROW = ['Translation', 'English <> French', 'Remote', '1000', '2', '0.12', 'Standard', '180']

# === AUTHENTICATE ===
creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)


print(f"Using document: https://docs.google.com/document/d/{DOCUMENT_ID}")

# resource = {
#     "service_account": creds,
#     "documentId": DOCUMENT_ID,
# }
# res = gdoctableapp.GetTables(resource)
# print(res)


# === Append a new row ===
dummy_row = [["Translation", "English <> French", "Remote", "1500", "2", "0.12", "Standard rate", "180"]]

# === Set row at index 1 (after header) ===
resource = {
    "service_account": creds,
    "documentId": DOCUMENT_ID,
    "tableIndex": 0,  # First table
    "values": [
        {
            "values": dummy_row,
            "range": {
                "startRowIndex": 1,  # Row just after header
                "startColumnIndex": 0
            }
        }
    ]
}

res = gdoctableapp.SetValues(resource)
print("Inserted row just after the header.")
