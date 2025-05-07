from googleapiclient.discovery import build  # Import the Google API client library to build service objects
from google.oauth2 import service_account  # Import Google OAuth2 library to handle authentication


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'  # Path to the service account credentials JSON file
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
SPREADSHEET_ID = '1wiAQXkSvcOS8QdLeST2AmjsaV03_bS-1dIM3XpiNNq0'  # ID of the Google Sheet
RANGE_NAME = 'Quotes!A1:Z2'  # Range of data to read from the sheet


def authenticate_gsheet(service_account_file: str, scopes: list):
    """Authenticate with Google Sheets API and return a service object."""
    credentials = service_account.Credentials.from_service_account_file(
        service_account_file, scopes=scopes)
    service = build('sheets', 'v4', credentials=credentials)
    return service.spreadsheets()


def read_sheet_data(sheet, spreadsheet_id: str, range_name: str):
    """Read data from the specified Google Sheet range."""
    result = sheet.values().get(spreadsheetId=spreadsheet_id, range=range_name).execute()
    return result.get('values', [])


def display_data(values: list):
    """Print data rows or show a 'No data found' message."""
    if not values:
        print('No data found.')
    else:
        print('Data from sheet:')
        for row in values:
            print(row)


def print_sheet_header(sheet, spreadsheet_id: str, range_name: str):
    """Fetch and print the header row from a Google Sheet."""

    # Request the specified range (e.g., first row only)
    result = sheet.values().get(
        spreadsheetId=spreadsheet_id,
        range=range_name
    ).execute()

    # Get all rows returned (will be a list of lists)
    values = result.get('values', [])

    if values:
        # header = values[0]  # First row is the header
        # print("Header row:")
        # print(header)
        for row in values:
            print(row)
    else:
        print("No data found in the range.")


def append_quote_row(sheet, spreadsheet_id: str, tab_name: str, new_row: list):
    """
    Appends a new quote row to the specified tab in the Google Sheet.

    Args:
        sheet: Authenticated Google Sheets API client.
        spreadsheet_id: ID of the spreadsheet.
        tab_name: Name of the tab (sheet) to append to.
        new_row: List of values in the same order as your header.
    """
    # Define the target range — doesn't matter where exactly, just use A1 to anchor
    target_range = f"{tab_name}!A1"

    # Format the request body
    body = {
        'values': [new_row]
    }

    # Append the row
    response = sheet.values().append(
        spreadsheetId=spreadsheet_id,
        range=target_range,
        valueInputOption='USER_ENTERED',  # allows formulas, auto-formatting
        insertDataOption='INSERT_ROWS',  # always append as a new row
        body=body
    ).execute()

    print(f"{response.get('updates', {}).get('updatedRows', 0)} row(s) appended.")


def main():
    # Authenticate and connect to the Google Sheet
    sheet = authenticate_gsheet(SERVICE_ACCOUNT_FILE, SCOPES)
    #
    # # Read data from the Google Sheet
    # values = read_sheet_data(sheet, SPREADSHEET_ID, RANGE_NAME)
    #
    # # Display the data
    # display_data(values)

    # new_row = [
    #     '',  # Leave Quote ID empty — sheet will auto-generate
    #     'Quote Block 15',  # Quote Block
    #     '2025-05-11',  # Date
    #     'Bob Example',  # Client Name
    #     'bob@example.org',  # Email
    #     'Example Org',  # Organization
    #     'Interpretation',  # Service Type
    #     'English <> French',  # Language Pair
    #     '',  # Word Count (not used for interpretation)
    #     '2',  # Duration (hrs)
    #     'In-person',  # Modality
    #     '60',  # Rate
    #     '120',  # Total
    #     '',  # (blank column)
    #     'Interpretation: 2 hrs × 60',  # Details
    #     'Client prefers Monday',  # Notes
    #     'TarjimlyQuote_20250511_Bob',  # Output Filename
    #     '120',  # Grand Total
    #     '', '', '', ''  # Document Studio columns
    # ]

    # append_quote_row(sheet, SPREADSHEET_ID, 'Quotes', new_row)


def group_rows_by_quote_id(data, header):
    """Group rows by Quote ID and return a structured dictionary."""
    idx = {key: header.index(key) for key in [
        'Quote ID', 'Date', 'Client Name', 'Organization', 'Email', 'Notes',
        'Service Type', 'Language Pair', 'Modality', 'Word Count', 'Duration (hrs)',
        'Rate', 'Details', 'Total'
    ] if key in header}

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
                'Grand Total': 0.0
            }

        # Add the service details to the quote's 'rows' list
        service_data = {
            'Service Type': row_data.get('Service Type', ''),
            'Language Pair': row_data.get('Language Pair', ''),
            'Modality': row_data.get('Modality', ''),
            'Word Count': row_data.get('Word Count', ''),
            'Duration (hrs)': row_data.get('Duration (hrs)', ''),
            'Rate': row_data.get('Rate', ''),
            'Details': row_data.get('Details', ''),
            'Total': row_data.get('Total', '')
        }

        grouped[quote_id]['rows'].append(service_data)

        # Accumulate total
        try:
            grouped[quote_id]['Grand Total'] += float(row_data.get('Total', '0'))
        except ValueError:
            pass

    return list(grouped.values())


if __name__ == '__main__':
    main()
