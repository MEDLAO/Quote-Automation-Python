from googleapiclient.discovery import build  # Import the Google API client library to build service objects
from google.oauth2 import service_account  # Import Google OAuth2 library to handle authentication


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'qas-credentials.json'  # Path to the service account credentials JSON file
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']  # Scope: only read access to Google Sheets
SPREADSHEET_ID = ''  # ID of the Google Sheet
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


def main():
    # Authenticate and connect to the Google Sheet
    sheet = authenticate_gsheet(SERVICE_ACCOUNT_FILE, SCOPES)
    #
    # # Read data from the Google Sheet
    # values = read_sheet_data(sheet, SPREADSHEET_ID, RANGE_NAME)
    #
    # # Display the data
    # display_data(values)

    # print_sheet_header(sheet, SPREADSHEET_ID, RANGE_NAME)
    choices = get_column_dropdown_choices(sheet, SPREADSHEET_ID, "Quotes",
                                          "G")  # Column G = 'Service Type'
    print("Dropdown options:", choices)


if __name__ == '__main__':
    main()
