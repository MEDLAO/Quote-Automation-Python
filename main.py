from googleapiclient.discovery import build  # Import the Google API client library to build service objects
from google.oauth2 import service_account  # Import Google OAuth2 library to handle authentication


# === CONFIGURATION ===
SERVICE_ACCOUNT_FILE = 'path/to/your/credentials.json'  # Path to your service account credentials JSON file
SCOPES = ['https://www.googleapis.com/auth/spreadsheets.readonly']  # Scope: only read access to Google Sheets

# ID of the Google Sheet you want to access (found in the sheet URL)
SPREADSHEET_ID = 'your-google-sheet-id-here'
# Range of data to read from the sheet (e.g., first 5 columns)
RANGE_NAME = 'Sheet1!A1:E'

# === AUTHENTICATION ===
# Load service account credentials with the specified scopes
credentials = service_account.Credentials.from_service_account_file(
    SERVICE_ACCOUNT_FILE, scopes=SCOPES)

# Build the Sheets API service using the credentials
service = build('sheets', 'v4', credentials=credentials)
# Access the 'spreadsheets' resource of the Sheets API
sheet = service.spreadsheets()

# === READ DATA ===
# Make an API call to read values from the specified sheet and range
result = sheet.values().get(spreadsheetId=SPREADSHEET_ID, range=RANGE_NAME).execute()
# Extract the actual data values from the API response
values = result.get('values', [])

# Check if any data was returned
if not values:
    print('No data found.')  # If no data, print a message
else:
    print('Data from sheet:')  # Otherwise, print the retrieved data
    for row in values:
        print(row)  # Print each row individually
