from google.oauth2 import service_account
from googleapiclient.discovery import build


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
TEMPLATE_DOC_ID = '18OVjzAQnTZKqhFmiaemIaHaw0QV7G8fPgMDGnwh-Wpg'
NEW_DOC_NAME = 'Replicated Table Document'

# Table data from the original template
REPLICATED_TABLE = [
    ['Service Type', 'Language Pair', 'Modality', 'Word Count', 'Duration (hrs)', 'Rate', 'Details', 'Total (USD)'],
    ['', '', '', '', '', '', '', ''],
    ['{{#if Grand Total}}\nGrand Total: {{Grand Total}} USD\n{{/if}}', '', '', '', '', '', '', '']
]


# === AUTHENTICATION ===
def authenticate_docs_and_drive(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    docs_service = build('docs', 'v1', credentials=credentials)
    drive_service = build('drive', 'v3', credentials=credentials)
    return docs_service, drive_service


# === COPY TEMPLATE ===
def copy_template(drive_service, template_id, new_name):
    copied = drive_service.files().copy(fileId=template_id, body={'name': new_name}).execute()
    print(f"Copied template: {copied['name']}")
    return copied['id']


# === Insert empty row after first data row ===
def insert_empty_row_after(doc_id, docs_service, table_index=0, after_row=1):
    doc = docs_service.documents().get(documentId=doc_id).execute()
    content = doc.get('body', {}).get('content', [])

    table_counter = 0
    for element in content:
        if 'table' in element:
            if table_counter == table_index:
                table_start_index = element['startIndex']
                break
            table_counter += 1
    else:
        print("Table not found.")
        return

    # Insert a new row below the first data row (after_row=1)
    requests = [{
        'insertTableRow': {
            'tableCellLocation': {
                'tableStartLocation': {'index': table_start_index},
                'rowIndex': after_row
            },
            'insertBelow': True
        }
    }]

    docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
    print(f"Inserted a new empty row after row {after_row + 1}.")


# === MAIN ===
if __name__ == '__main__':
    docs_service, drive_service = authenticate_docs_and_drive(SERVICE_ACCOUNT_FILE, SCOPES)
    new_doc_id = copy_template(drive_service, TEMPLATE_DOC_ID, NEW_DOC_NAME)
    insert_empty_row_after(new_doc_id, docs_service, table_index=0, after_row=1)

    print(f"Document ready: https://docs.google.com/document/d/{new_doc_id}")
