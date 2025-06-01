from google.oauth2 import service_account
from googleapiclient.discovery import build


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'
SCOPES = ['https://www.googleapis.com/auth/documents', 'https://www.googleapis.com/auth/drive']
TEMPLATE_DOC_ID = '1IOjo6NyH7Q6W0Sgs_GoMX_QdhSlb7snFtOhnTynNFJs'
# '18OVjzAQnTZKqhFmiaemIaHaw0QV7G8fPgMDGnwh-Wpg'
NEW_DOC_NAME = 'Replicated Table Document'

DUMMY_ROW = ['X'] * 8


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


# === FILL THE SECOND ROW (AFTER HEADER) ===
def insert_row_with_same_value(doc_id, docs_service, value='X', table_index=0):
    # Fetch the doc and locate the first table
    doc = docs_service.documents().get(documentId=doc_id).execute()
    content = doc.get('body', {}).get('content', [])

    table = None
    table_start_index = None
    for el in content:
        if 'table' in el:
            table = el['table']
            table_start_index = el['startIndex']
            break
    if not table:
        print("Table not found.")
        return

    # Step 1: Insert a new row below the header (row 0)
    docs_service.documents().batchUpdate(documentId=doc_id, body={
        'requests': [{
            'insertTableRow': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': table_start_index},
                    'rowIndex': 0
                },
                'insertBelow': True
            }
        }]
    }).execute()

    # Step 2: Refresh doc to get the updated row
    doc = docs_service.documents().get(documentId=doc_id).execute()
    table = doc['body']['content'][[i for i, el in enumerate(doc['body']['content']) if 'table' in el][table_index]]['table']
    new_row_cells = table['tableRows'][1]['tableCells']  # second row now

    # Step 3: Build insert requests with safety
    requests = []
    for col_index, cell in enumerate(new_row_cells):
        if col_index >= 8:
            break

        try:
            content = cell.get('content', [])
            paragraph = content[0].get('paragraph') if content else None
            elements = paragraph.get('elements') if paragraph else None

            if not elements:
                # Insert a newline to force paragraph creation
                requests.append({
                    'insertText': {
                        'location': {'index': cell['startIndex'] + 1},
                        'text': '\n'
                    }
                })
                # Then insert actual value
                requests.append({
                    'insertText': {
                        'location': {'index': cell['startIndex'] + 1},
                        'text': value
                    }
                })
            else:
                start_index = elements[0]['startIndex']
                requests.append({
                    'insertText': {
                        'location': {'index': start_index},
                        'text': value
                    }
                })
        except Exception as e:
            print(f"Skipping column {col_index}: {e}")

    # Step 4: Apply insertions
    if requests:
        docs_service.documents().batchUpdate(documentId=doc_id, body={'requests': requests}).execute()
        print("Inserted 'X' in all 8 cells safely.")


# === MAIN ===
if __name__ == '__main__':
    docs_service, drive_service = authenticate_docs_and_drive(SERVICE_ACCOUNT_FILE, SCOPES)
    new_doc_id = copy_template(drive_service, TEMPLATE_DOC_ID, NEW_DOC_NAME)
    insert_empty_row_after(new_doc_id, docs_service, table_index=0, after_row=1)
    # insert_row_with_same_value(new_doc_id, docs_service, value='X')

    print(f"Document ready: https://docs.google.com/document/d/{new_doc_id}")
