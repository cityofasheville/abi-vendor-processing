from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd
from os.path import exists
import time

from sqlalchemy import column


SERVICE_ACCOUNT_FILE = 'budget-data-processing-2e6c08029402.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
OUTPUTS_SPREADSHEET_ID = None
INPUTS_SPREADSHEET_ID = None

sheetService = None


#########################################################


def setUpServices():
  global sheetService
  creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )
  sheetService = build('sheets', 'v4', credentials=creds)
  #driveService = build('drive', 'v3', credentials=creds)


def build_list(INPUTS_SPREADSHEET_ID):
    finalEntryList = []
    sheet = sheetService.spreadsheets()
    
    badFormatList = [4, 11, 12, 32, 34, 41, 44, 53, 61, 62, 82, 83, 94, 100, 116, 118, 165, 234, 235, 236, 237, 238, 127, 134, 14, 119, 166, 63, 65, 79, 90, 111, 117, 120, 139, 140, 141, 144, 146, 147, 148, 149, 150, 151, 152, 153, 154, 155, 156, 157, 158, 159, 160, 161, 162, 163, 164, 167, ]


    # Reads spreadsheet to get list from the spreadsheet
    total_list = []
    results = sheet.values().get(spreadsheetId=INPUTS_SPREADSHEET_ID,range='Project List!A2:L147').execute()
    link_ss_values = results.get('values', [])

    y=0   
    sheetCount = 0

    for entry in link_ss_values:
      print(link_ss_values.index(entry)+2)
      sheetCount += 1
      y += 1

      if link_ss_values.index(entry) in badFormatList:
        finalEntryList.append([' '])
      
      else:
        url = entry[1]
        notImportant, important = url.split('/d/')
        id, sheetTabId = important.split('/edit#gid=')
        sheets= sheet.get(spreadsheetId=id, fields='sheets/properties').execute()
        nameIdDictionary = {}
        for item in sheets['sheets']:
          nameIdDictionary[str(item['properties']['sheetId'])] = item['properties']['title']
        
        sheetTitle = nameIdDictionary[sheetTabId]
 
        # Read the project spreadsheets

        resultsProject = sheet.values().get(spreadsheetId=id,range=sheetTitle + '!A1:A42').execute()
        values = resultsProject.get('values', [])

        #print(len(values))
        finalEntry = values[22] + values[24] + values[26] + values[28] + values[30] + values[32] + values[34]

        valuesLength = len(values)
        otherinfo = ""
        i=0

        if valuesLength > 36:
          finalEntry.append(values[36][0])
        
        if valuesLength > 36:
          difference = valuesLength - 36
          while i < difference-1:
            print("this is round "+ str(i))
            otherinfo  = otherinfo + " " + values[37+i][0]
            i+=1
          finalEntry.append(otherinfo)
        #print(finalEntry)


        finalEntryList.append(finalEntry)
        #print(finalEntryList)
        #print(nameIdDictionary)

      if y == 300:
        break

        # We have the sheet ID, but not the title, which we need to specify the range
        # The code below makes a dictionary that we can use to get the title
      if sheetCount >20:
        print('Pausing for 30 seconds ... ')
        time.sleep(50)
        sheetCount = 0



 
    return(finalEntryList)

# Main Program Star-------------------------------------------------


#Open Json
inputs = None
if exists('./inputs.json'):
    with open('inputs.json', 'r') as file:
        inputs = json.load(file)
else:
    print('You must create an inputs.json file')
    sys.exit()

# # Set const values
INPUTS_SPREADSHEET_ID = inputs["INPUTS_SPREADSHEET_ID"]
OUTPUTS_SPREADSHEET_ID = inputs["OUTPUTS_SPREADSHEET_ID"]

setUpServices()
sheet = sheetService.spreadsheets()


list_to_append = build_list(INPUTS_SPREADSHEET_ID)


 # Update sheet
resource = {
  "majorDimension": "ROWS",
   "values": list_to_append
}

sheetService.spreadsheets().values().update(
  spreadsheetId=OUTPUTS_SPREADSHEET_ID,
  range="Sheet1!M2:AA1000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()



