from googleapiclient.discovery import build  # Import the Google API client library to build service objects
from google.oauth2 import service_account  # Import Google OAuth2 library to handle authentication
from gdoctableapppy import gdoctableapp
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
TEMPLATE_DOC_ID = '18OVjzAQnTZKqhFmiaemIaHaw0QV7G8fPgMDGnwh-Wpg'
RANGE_NAME = 'Quotes!A1:Z'  # Range of data to read from the sheet


# === AUTHENTICATION ===
def authenticate_gsheet(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('sheets', 'v4', credentials=credentials).spreadsheets()


def authenticate_gdoc(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('docs', 'v1', credentials=credentials)


def authenticate_drive(service_account_file, scopes):
    credentials = service_account.Credentials.from_service_account_file(service_account_file, scopes=scopes)
    return build('drive', 'v3', credentials=credentials).files()


# === COPY TEMPLATE ===
def copy_template(drive_service, template_id, new_name):
    copied = drive_service.files().copy(fileId=template_id, body={'name': new_name}).execute()
    print(f"Copied template: {copied['name']}")
    return copied['id']


# === STEP 1 - Fetch/Read the data from Quotes Spreadsheet ===
def read_sheet_data(sheet, spreadsheet_id: str, range_name: str):
    """Read data from the specified Google Sheet range."""
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get('values', [])


# === STEP 2 - Group the fetched data per Quote-ID ===
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
                'Grand Total': 0.0,
                'Num Services': 0
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
        grouped[quote_id]['Num Services'] += 1  # INCREMENT

        try:
            grouped[quote_id]['Grand Total'] += float(row_data.get('Total', '0'))
        except ValueError:
            pass

    return list(grouped.values())


# === STEP 3 - Write the grouped data to GroupedQuotes Spreadsheet ===
def write_grouped_data(sheet, spreadsheet_id, target_sheet_name, grouped_data):
    """Write grouped quotes to an existing sheet starting at cell A1."""

    # Define the header
    header = ['Quote ID', 'Date', 'Client Name', 'Email', 'Organization', 'Notes', 'Services',
              'Grand Total', 'Num Services']
    rows_to_write = [header]

    for entry in grouped_data:
        # Convert rows to a compact JSON string (no escaping)
        rows_json = json.dumps(entry['rows'], ensure_ascii=False, separators=(',', ':'))

        # Prepare row without extra escaping
        rows_to_write.append([
            entry['Quote ID'],
            entry['Date'],
            entry['Client Name'],
            entry['Email'],
            entry['Organization'],
            entry['Notes'],
            rows_json,
            f"{entry['Grand Total']:.2f}",
            entry['Num Services']
        ])

    # Write the data to the target sheet
    sheet.values().update(
        spreadsheetId=spreadsheet_id,
        range=f"{target_sheet_name}!A1",
        valueInputOption='RAW',
        body={"values": rows_to_write}
    ).execute()

    print(f"Grouped data written to existing sheet '{target_sheet_name}'.")


# === STEP 4 - Generate the Quotes documents and add the right number of empty rows ===
def insert_empty_row_after(doc_id, docs_service, entry, table_index=0, after_row=1):
    """
    Inserts (number of services - 1) empty rows into the first table of a Google Doc.
    """
    num_services = len(entry.get('rows', []))

    if num_services <= 1:
        return  # Only 1 row needed, already in template

    # Get the document content
    doc = docs_service.documents().get(documentId=doc_id).execute()
    content = doc.get('body', {}).get('content', [])

    # Find the specified table
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

    # Build requests to insert (num_services - 1) rows
    requests = []
    for i in range(num_services - 1):
        requests.append({
            'insertTableRow': {
                'tableCellLocation': {
                    'tableStartLocation': {'index': table_start_index},
                    'rowIndex': after_row + i
                },
                'insertBelow': True
            }
        })

    if requests:
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={'requests': requests}
        ).execute()
        print(f"{len(requests)} empty row(s) inserted for Quote ID {entry['Quote ID']}.")
    else:
        print(f"No insert requests created for Quote ID {entry['Quote ID']}.")

    print(f"{len(requests)} empty row(s) inserted for Quote ID {entry['Quote ID']}.")


def fill_services_table(doc_id, creds, services, table_index=0, start_row=1, start_col=0):
    """
    Fills the services table in a Google Doc using gdoctableapp.

    Args:
        doc_id (str): ID of the target Google Doc.
        creds: Authenticated service account credentials.
        services (list): A list of lists representing service rows.
        table_index (int): Index of the table in the document (default: 0).
        start_row (int): Row to start filling from (default: 1 = after header).
        start_col (int): Column to start filling from (default: 0).
    """
    resource = {
        "service_account": creds,
        "documentId": doc_id,
        "tableIndex": table_index,
        "values": [
            {
                "values": services,
                "range": {
                    "startRowIndex": start_row,
                    "startColumnIndex": start_col
                }
            }
        ]
    }

    res = gdoctableapp.SetValues(resource)
    print(f"Filled {len(services)} service row(s) in document {doc_id}.")
    return res


def generate_docs_for_grouped_quotes(grouped_data, gdoc, gdrive, template_id, creds):
    """
    Copies the template for each quote, inserts empty rows, and fills the table with service data.

    Args:
        grouped_data (list): List of grouped quote entries (each with 'Quote ID' and 'Services').
        gdoc: Authenticated Google Docs API client.
        gdrive: Authenticated Google Drive API client.
        template_id (str): Google Doc template ID.
        creds: Authenticated service account credentials for gdoctableapp.
    """
    for entry in grouped_data:
        # Step 1: Copy the template
        copied_file = gdrive.copy(
            fileId=template_id,
            body={'name': f"Quote_{entry['Quote ID']}"}
        ).execute()
        new_doc_id = copied_file['id']

        # Step 2: Insert correct number of empty rows
        insert_empty_row_after(new_doc_id, gdoc, entry)

        # Step 3: Transform entry["rows"] into a list of lists
        services = [
            [
                s.get('Service_Type', ''),
                s.get('Language_Pair', ''),
                s.get('Modality', ''),
                s.get('Word_Count', ''),
                s.get('Duration_hrs', ''),
                s.get('Rate', ''),
                s.get('Details', ''),
                s.get('Total', '')
            ]
            for s in entry["rows"]
        ]

        # Step 4: Fill table with service data
        fill_services_table(
            doc_id=new_doc_id,
            creds=creds,
            services=services
        )

        # Final log
        print(f"Document created and filled: https://docs.google.com/document/d/{new_doc_id}")


def main():
    # Step 1: Authenticate all Google services
    sheet = authenticate_gsheet(SERVICE_ACCOUNT_FILE, SCOPES)
    gdoc = authenticate_gdoc(SERVICE_ACCOUNT_FILE, SCOPES)
    gdrive = authenticate_drive(SERVICE_ACCOUNT_FILE, SCOPES)

    # Step 2: Create credentials object for gdoctableapp
    creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES
    )

    # Step 3: Read data from the source spreadsheet
    values = read_sheet_data(sheet, SPREADSHEET_ID_SOURCE, RANGE_NAME)
    if not values:
        print("No data found in the source sheet.")
        return

    # Step 4: Group the rows by Quote ID
    header = values[0]
    grouped_data = group_rows_by_quote_id(values, header)

    # Step 5: Generate quote documents from grouped data
    generate_docs_for_grouped_quotes(
        grouped_data=grouped_data,
        gdoc=gdoc,
        gdrive=gdrive,
        template_id=TEMPLATE_DOC_ID,
        creds=creds
    )


if __name__ == '__main__':
    main()
