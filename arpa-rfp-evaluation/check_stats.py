from googleapiclient.discovery import build
import json
import sys
import time
from csv import reader
from google.oauth2 import service_account
import numpy as np
import statistics as stat
from os.path import exists



SERVICE_ACCOUNT_FILE = None
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
INPUTS_SPREADSHEET_ID = None
OUTPUTS_MASTER_ID = None
INPUTS_EVAL_MAPPING_ID = None

sheetService = None


#########################################################


def setUpServices():
  global sheetService
  creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )
  sheetService = build('sheets', 'v4', credentials=creds)
  #driveService = build('drive', 'v3', credentials=creds)

def stripLower(lst):
    return list(map(lambda itm: itm.strip().lower() if itm else None, lst))

def getSheetTitles(sheet, spreadsheetId):
    sheets = sheet.get(spreadsheetId=spreadsheetId, fields='sheets/properties').execute()
    return [{"title": sheet['properties']['title'], "id": sheet['properties']['sheetId']} for sheet in sheets['sheets']]

evaluationStatus = []

# Read the list of categories, strip whitespace and lowercase them
def readCategories(OUTPUTS_MASTER_ID):
  values = sheetService.spreadsheets().values().get(spreadsheetId=OUTPUTS_MASTER_ID,range='All Data!A1:R1').execute().get('values', [])
  return stripLower(values[0][8:18])

def getQuestionMaxPoints():
    global INPUTS_SPREADSHEET_ID
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=INPUTS_SPREADSHEET_ID,range='Score Weighting!D8:D27').execute()
    values = results.get('values', [])
    del values[13]
    del values[6]
    maxPoints = np.array(values)[0:len(values),0].astype(float).tolist()
    print('Max Points: ', maxPoints)
    return(maxPoints)


def readAllData(maxPoints):
    global OUTPUTS_MASTER_ID
    answerFactors = {
        'high': 1.0, 'medium': 2.0/3.0, 'low': 1.0/3.0, 'none': 0.0
    }

    allData = {}
    values = sheetService.spreadsheets().values().get(spreadsheetId=OUTPUTS_MASTER_ID,range='All Data!A2:H').execute().get('values', [])
    # Drop the ones that aren't complete
    result = list(filter(lambda row: row[7] == 'yes', values))
    for row in result:
        if row[2] not in allData:
            allData[row[2]] = {}
        if row[0] not in allData[row[2]]:
            allData[row[2]][row[0]] = 0.0
        qNum = int(row[4]) - 1
        allData[row[2]][row[0]] += answerFactors[row[6].strip().lower()] * maxPoints[qNum]
    out = [[
      'Proposal', 'Median', 'Mean', 'Evaluator1', 'Evaluator2', 'Evaluator3', 'Evaluator4', 'Evaluator5'
    ]]
    for proposal in allData:
        scores = []
        for review in allData[proposal]:
            scores.append(round(allData[proposal][review], 2))
        med = round(stat.median(scores), 2)
        mean = round(stat.mean(scores), 2)
        row = [proposal, med, mean]
        row.extend(scores)
        out.append(row)
    return out


############################ Main Program Start

#Open Json
inputs = None
if exists('./inputs.json'):
    with open('inputs.json', 'r') as file:
        inputs = json.load(file)
else:
    print('You must create an inputs.json file')
    sys.exit()

# Set global variables
INPUTS_SPREADSHEET_ID = inputs['INPUTS_SPREADSHEET_ID']
OUTPUTS_MASTER_ID = inputs["OUTPUTS_MASTER_ID"]
INPUTS_EVAL_MAPPING_ID = inputs["INPUTS_EVAL_MAPPING_ID"]
SERVICE_ACCOUNT_FILE = inputs["SERVICE_ACCOUNT_FILE"]

setUpServices()

maxPoints = getQuestionMaxPoints()
# Read in the detailed data
allData = readAllData(maxPoints)

# Update sheet
resource = {
  "majorDimension": "ROWS",
  "values": allData
}
sheetService.spreadsheets().values().update(
   spreadsheetId=OUTPUTS_MASTER_ID,
   range="StatCheck!A1:AA10000",
   body=resource,
   valueInputOption="USER_ENTERED").execute()

# # Update sheet
# resource = {
#   "majorDimension": "ROWS",
#   "values": evaluationStatus
# }

# sheetService.spreadsheets().values().update(
#   spreadsheetId=OUTPUTS_MASTER_ID,
#   range="Evaluation Status!A2:AA10000",
#   body=resource,
#   valueInputOption="USER_ENTERED").execute()



