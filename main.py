from googleapiclient.discovery import build  # Import the Google API client library to build service objects
from google.oauth2 import service_account  # Import Google OAuth2 library to handle authentication


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'path/to/your/credentials.json'  # Path to your service account credentials JSON file
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']  # Scope: only read access to Google Sheets
SPREADSHEET_ID = 'your-google-sheet-id-here'  # ID of the Google Sheet
RANGE_NAME = 'Sheet1!A1:E'  # Range of data to read from the sheet


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


def main():
    # Authenticate and connect to the Google Sheet
    sheet = authenticate_gsheet(SERVICE_ACCOUNT_FILE, SCOPES)

    # Read data from the Google Sheet
    values = read_sheet_data(sheet, SPREADSHEET_ID, RANGE_NAME)

    # Display the data
    display_data(values)


if __name__ == '__main__':
    main()
