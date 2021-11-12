
from googleapiclient.discovery import build

import json
from csv import reader

from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'key.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
READ_SPREADSHEET_ID = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'
service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()
result = sheet.values().get(spreadsheetID=READ_SPREADSHEET_ID, range='Evaluator View!A1:E22').execute()

print(result)

#############################
