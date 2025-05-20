from gdoctableapppy import gdoctableapp
from google.oauth2 import service_account


# === Setup ===
SCOPES = ['https://www.googleapis.com/auth/documents']
SERVICE_ACCOUNT_FILE = 'your-service-account.json'
DOCUMENT_ID = 'your-doc-id'

creds = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

gdoc = gdoctableapp.DocsTable(creds)

# === Read all tables ===
tables = gdoc.GetTables(DOCUMENT_ID)
print(f"Found {len(tables)} tables.")

# === Get and print first table ===
values = gdoc.GetValues(DOCUMENT_ID, tableIndex=0)
for row in values:
    print(row)

# === Set new values ===
new_values = [
    ['Service Type', 'Language Pair', 'Modality'],
    ['Translation', 'English <> French', 'Remote'],
    ['Interpretation', 'Arabic <> English', 'On-site']
]
gdoc.SetValues(DOCUMENT_ID, tableIndex=0, values=new_values)
