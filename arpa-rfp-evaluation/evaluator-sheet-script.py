
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd


SERVICE_ACCOUNT_FILE = 'arpa-processing-25528ff0b6f2.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
ASSIGN_SPREADSHEET_ID = '1oy7i8HOhDbxsvsXcwRnrBUV75o3giiS1tWeJQEhP-is'
FORMAT_SPREADSHEET_ID = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

result = sheet.values().get(spreadsheetId=ASSIGN_SPREADSHEET_ID,range="Eligible Proposals and Assignments!A1:L4").execute()

values = result.get('values', [])

#############################

def create_sheet(values):
        #Creates a dictionary where each evaluator name is a key, and the corresponding value is a list, which contains 2 lists.
        #The first is a list of the proposals assigned, the second is a list of the links to the proposals 
        # !!! Links must be typed out !!!
        values.pop(0)
        evals = {}
        for j in values:
                for item in j[5:]:
                        if item in evals:
                                evals[item][0].append(j[1])
                                evals[item][1].append(j[2])
                        else:
                                evals[item]=[[j[1]], [j[2]]]
        print(evals)

         
create_sheet(values)
