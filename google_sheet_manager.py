from google.oauth2.credentials import Credentials
from google_auth_oauthlib.flow import InstalledAppFlow
from google.auth.transport.requests import Request
from googleapiclient.discovery import build
import os.path
import pickle

class GoogleSheetsManager:
    def __init__(self, spreadsheet_id=None):
        """Initialize Google Sheets Manager."""
        self.SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
        self.spreadsheet_id = spreadsheet_id
        self.service = None

    def authenticate(self):
        """Authenticate with Google Sheets API."""
        creds = None

        # Load existing credentials
        if os.path.exists('token.pickle'):
            with open('token.pickle', 'rb') as token:
                creds = pickle.load(token)

        # If no valid credentials, get new ones
        if not creds or not creds.valid:
            if creds and creds.expired and creds.refresh_token:
                creds.refresh(Request())
            else:
                flow = InstalledAppFlow.from_client_secrets_file(
                    'client_secret_832531750728-t97pdct75n3v5bprctkg81l0tbdg5k92.apps.googleusercontent.com.json',
                    self.SCOPES)
                creds = flow.run_local_server(port=0)

            # Save credentials
            with open('token.pickle', 'wb') as token:
                pickle.dump(creds, token)

        self.service = build('sheets', 'v4', credentials=creds)

    def create_spreadsheet(self, title="My Data"):
        """Create a new Google Spreadsheet."""
        if not self.service:
            self.authenticate()

        spreadsheet = {
            'properties': {
                'title': title
            }
        }

        spreadsheet = self.service.spreadsheets().create(body=spreadsheet).execute()
        self.spreadsheet_id = spreadsheet['spreadsheetId']
        print(f"Created spreadsheet: {spreadsheet['properties']['title']}")
        print(f"Spreadsheet ID: {self.spreadsheet_id}")
        return self.spreadsheet_id

    def add_data(self, data, range_name="Sheet1!A1"):
        """Add data to spreadsheet."""
        if not self.service:
            self.authenticate()

        body = {
            'values': [data]
        }

        result = self.service.spreadsheets().values().append(
            spreadsheetId=self.spreadsheet_id,
            range=range_name,
            valueInputOption='RAW',
            insertDataOption='INSERT_ROWS',
            body=body
        ).execute()

        print(f"Added {len(data)} cells to {result.get('updates').get('updatedRange')}")

    def read_data(self, range_name="Sheet1!A:Z"):
        """Read data from spreadsheet."""
        if not self.service:
            self.authenticate()

        result = self.service.spreadsheets().values().get(
            spreadsheetId=self.spreadsheet_id,
            range=range_name
        ).execute()

        return result.get('values', [])