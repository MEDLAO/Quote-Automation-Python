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


def extract_drive_file_id(url):
    """Extracts the file ID from a Google Drive 'open?id=' style URL."""
    if "open?id=" in url:
        return url.split("open?id=")[1].split("&")[0]
    return None


def share_document(gdrive, doc_id, anyone=True, email=None):
    """
    Shares a Google Doc either publicly (anyone with the link) or with a specific email.
    """
    body = {
        'type': 'anyone' if anyone else 'user',
        'role': 'writer',
    }

    if email:
        body['type'] = 'user'
        body['emailAddress'] = email

    gdrive.permissions().create(
        fileId=doc_id,
        body=body,
        fields='id'
    ).execute()


# === STEP 2 - Group the fetched data per Quote-ID ===
def group_rows_by_quote_id(data, header):
    """Group rows by Quote ID, skipping rows with Total == 'Manual' and removing empty groups."""
    grouped = {}

    for row in data[1:]:  # Skip header row
        row += [''] * (len(header) - len(row))  # Pad short rows with empty strings
        row_data = dict(zip(header, row))

        quote_id = row_data.get("Quote ID", "").strip()
        if not quote_id:
            continue  # Skip rows without a Quote ID

        # Extract document link and ID
        doc_url = row_data.get("[Document Studio] File Link #4zzo1e", "").strip()
        doc_id = extract_drive_file_id(doc_url)

        # Initialize a new group if this Quote ID is new
        if quote_id not in grouped:
            grouped[quote_id] = {
                'Quote ID': quote_id,
                'Date': row_data.get('Date', ''),
                'Client Name': row_data.get('Client Name', ''),
                'Email': row_data.get('Email', ''),
                'Organization': row_data.get('Organization', ''),
                'Notes': row_data.get('Notes', ''),
                'Document ID': doc_id,
                'rows': [],  # Will hold valid service rows
                'Grand Total': 0.0,
                'Num Services': 0
            }

        # Skip this row if the Total field is 'Manual'
        if row_data.get('Total', '').strip().lower() == 'manual':
            continue

        # Collect and normalize service data
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

        # Skip service rows where all fields are empty
        if not any(str(value).strip() for value in service_data.values()):
            continue

        # Add valid row to the group
        grouped[quote_id]['rows'].append(service_data)
        grouped[quote_id]['Num Services'] += 1

        # Try adding the Total to the Grand Total
        try:
            grouped[quote_id]['Grand Total'] += float(row_data.get('Total', '0'))
        except ValueError:
            pass  # Ignore totals that are not numeric

    # Remove groups that have no valid rows (i.e., only "Manual" rows were skipped)
    return [entry for entry in grouped.values() if entry['rows']]


# === STEP 3 - Write the grouped data to GroupedQuotes Spreadsheet ===
def write_grouped_data(sheet, spreadsheet_id, target_sheet_name, grouped_data):
    """
    Write grouped quotes to an existing sheet starting at cell A1,
    ensuring that rows with 'Total' == 'Manual' are excluded.
    """

    # Define the header for the GroupedQuotes sheet
    header = ['Quote ID', 'Date', 'Client Name', 'Email', 'Organization', 'Notes', 'Services',
              'Grand Total', 'Num Services']
    rows_to_write = [header]

    for entry in grouped_data:
        # Filter out any service rows where Total is 'Manual' (as a double safety check)
        filtered_rows = [
            row for row in entry['rows']
            if str(row.get('Total', '')).strip().lower() != 'manual'
        ]

        if not filtered_rows:
            continue  # Skip this entry if it contains no valid rows

        # Convert the filtered service rows into a compact JSON string for storage
        rows_json = json.dumps(filtered_rows, ensure_ascii=False, separators=(',', ':'))

        # Add the processed entry to the output
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

    # Write all collected rows to the target sheet
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
    Inserts (number of services - 1) empty rows into the specified table of a Google Doc.
    """
    num_services = len(entry.get('rows', []))

    # No need to insert if only one service (the template already includes one row)
    if num_services <= 1:
        return

    # Retrieve the document structure
    doc = docs_service.documents().get(documentId=doc_id).execute()
    content = doc.get('body', {}).get('content', [])

    # Locate the start index of the desired table
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

    # Prepare insertion requests for (num_services - 1) rows
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

    # Send requests to Google Docs API
    if requests:
        docs_service.documents().batchUpdate(
            documentId=doc_id,
            body={'requests': requests}
        ).execute()
        print(f"{len(requests)} empty row(s) inserted for Quote ID {entry['Quote ID']}.")
    else:
        print(f"No insert requests created for Quote ID {entry['Quote ID']}.")


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
    Opens each Document Studio-generated doc by its ID, inserts rows, and fills service table.
    """
    for entry in grouped_data:
        doc_id = entry.get("Document ID")
        if not doc_id:
            print(f"Skipping Quote ID {entry['Quote ID']} (no doc ID found).")
            continue

        # Share the document before modifying
        share_document(gdrive, doc_id)

        # Step 1: Insert correct number of empty rows
        insert_empty_row_after(doc_id, gdoc, entry)

        # Step 2: Transform entry["rows"] into a list of lists
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

        # Step 3: Fill the service table
        fill_services_table(
            doc_id=doc_id,
            creds=creds,
            services=services
        )

        print(f"Document filled: https://docs.google.com/document/d/{doc_id}")


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

    write_grouped_data(
        sheet=sheet,
        spreadsheet_id=SPREADSHEET_ID_SOURCE,
        target_sheet_name="GroupedQuotes",
        grouped_data=grouped_data
    )

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
