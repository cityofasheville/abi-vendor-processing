#from __future__ import print_function
#import os.path
from googleapiclient.discovery import build
#from google_auth_oauthlib.flow import InstalledAppFlow
#from google.auth.transport.requests import Request
#from google.oauth2.credentials import Credentials
import json
from csv import reader

from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'key.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
SAMPLE_SPREADSHEET_ID = '1jwbcFVDx3X34gWXcj77C-PFMgI85QgHMlYoGjNY36hU'
service = build('sheets', 'v4', credentials=creds)


def writeCSV(asset, nm):
        sections = asset['sections']
        for secNumber in range(sections):
                if sections ==1:
                        name = nm
                else:
                        name = nm + "." + str(secNumber +1)
                print(name)
                sheet = service.spreadsheets()
        
                with open("./localdata/" + name + ".csv", 'r') as read_obj:
                        csv_reader=reader(read_obj)
                        data = list(csv_reader)
                        request = sheet.values().update(spreadsheetId=SAMPLE_SPREADSHEET_ID, range= name + "!A1", valueInputOption='USER_ENTERED', body={"values":data}).execute()



#############################

inputs = None
with open('inputs.json', 'r') as inputsFile:
        inputs = json.load(inputsFile)
for nm in inputs['files']:
        asset = inputs['files'][nm]
        data = None
        csv_path = "./localdata/" + nm + ".csv"
        data = writeCSV(asset, nm)


        