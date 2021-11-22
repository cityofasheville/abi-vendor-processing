from googleapiclient.discovery import build
import json
import sys
from csv import reader
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'arpa-processing-25528ff0b6f2.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )


# IDs of the various spreadsheets, sheets and folders
INPUTS_SPREADSHEET_ID = '1xrEqDmNd0jBAh_vth5ReC0DaUxFhYfKT0asQw-pF4kI' 
INPUTS_README_TAB_ID = 1285419200
INPUTS_EVAL_TEMPLATE_TAB_ID = 1023300661

TARGET_FOLDER_ID = '14_2ov-PiOeSAeFPYuxPuOzayMdx7L4yx'

sheetService = build('sheets', 'v4', credentials=creds)
driveService = build('drive', 'v3', credentials=creds)

evaluatorCount = 0
evaluatorIndices = {}
proposalIndices = {}
matrixMap = []
#############################

def getEvaluatorIndices():
    global evaluatorCount
    global evaluatorIndices
    global matrixMap
    result = sheetService.spreadsheets().values().get(spreadsheetId=INPUTS_SPREADSHEET_ID,range="Evaluators!A1:A100").execute()
    tmp = result.get('values', [])
    mmList = ['Proposal']
    evaluatorList = []
    for i in range(len(tmp)):
        evaluatorList.append(tmp[i][0])
        mmList.append(tmp[i][0])
    matrixMap.append(mmList)
    evaluatorCount = len(evaluatorList)

    for i in range(len(evaluatorList)):
        evaluatorIndices[evaluatorList[i]] = i

def process_assignments(values):
    global evaluatorCount
    global matrixMap
    values.pop(0)
    evals = {}
    proposalIndex = 0
    for row in values:
        if len(row[1]) == 0:
            continue
        proposalName = str(row[0]) + ' ' + row[1]
        lst = [-1] * (evaluatorCount+1)
        lst[0] = proposalName
        matrixMap.append(lst)
        proposalIndices[proposalName] = proposalIndex
        proposalIndex += 1
        for name in row[5:]:
            if name not in evals.keys():
                evals[name] = []
                evals[name].append ({'name': proposalName, 'link': row[2], 'categories': row[3]})
            else:
                evals[name].append({'name': proposalName, 'link': row[2], 'categories': row[3]})
    return evals

def copyAndRenameSheet(fromSpreadsheetId, fromSheetId, toSpreadsheetId, sheetName):

    response = sheetService.spreadsheets().sheets().copyTo(spreadsheetId=fromSpreadsheetId, sheetId=fromSheetId, body={
      'destination_spreadsheet_id': toSpreadsheetId
    }).execute()
    newSheetId = response['sheetId']

    # Rename the first sheet to README
    request = sheetService.spreadsheets().batchUpdate(spreadsheetId=toSpreadsheetId, body={
        'requests': [
          { "updateSheetProperties": {
            "properties": {
              "sheetId": newSheetId,
              "title": sheetName
            },
            "fields": 'title'
          }}
        ]
    })
    response = request.execute()
    return newSheetId

def createSpreadsheet(spreadsheetName, folderId):
    file_metadata = {
        'name': spreadsheetName,
        'parents': [folderId],
        'mimeType': 'application/vnd.google-apps.spreadsheet',
    }
    res = driveService.files().create(body=file_metadata).execute()
    return res['id']

def create_one_sheet(evaluator, proposals):
    print('Creating a spreadsheet for ', evaluator)
    # Create the spreadsheet with the name of the evaluator
    evaluatorSheetId = createSpreadsheet(evaluator, TARGET_FOLDER_ID)

    # Copy over the README sheet
    response = copyAndRenameSheet(INPUTS_SPREADSHEET_ID,
    INPUTS_README_TAB_ID,
    evaluatorSheetId, 'README')

    # Delete the original sheet1
    response = sheetService.spreadsheets().batchUpdate(spreadsheetId=evaluatorSheetId, body={
        'requests': [
          { "deleteSheet": {"sheetId": 0}}
        ]
    }).execute()

    # Now copy over the evaluation sheet for each of the assigned evaluations
    for proposal in proposals:
        print('   Adding proposal: ', proposal['name'])
        newSheetId = copyAndRenameSheet(INPUTS_SPREADSHEET_ID,INPUTS_EVAL_TEMPLATE_TAB_ID, evaluatorSheetId, proposal['name'])
        matrixMap[proposalIndices[proposal['name']]+1][evaluatorIndices[evaluator]+1] = newSheetId
        # Now update the cells at top
        hyperlink = '=HYPERLINK("'+proposal['link'] + '","Link to Proposal")'
        sheetService.spreadsheets().values().update(spreadsheetId=evaluatorSheetId, 
        range=proposal['name']+"!B1:B4",
        valueInputOption='USER_ENTERED', body={'values': [
          ['Evaluator: ' + evaluator],
          ['Project Name: ' + proposal['name']],
          ['Categories: ' + proposal['categories']],
          [hyperlink]
        ]}).execute()
    return evaluatorSheetId

##
## Main program
##

# Read the assignments spreadsheet (both evaluators and assignments)
getEvaluatorIndices()

result = sheetService.spreadsheets().values().get(spreadsheetId=INPUTS_SPREADSHEET_ID,range="Eligible Proposals and Assignments!A1:L100").execute()
values = result.get('values', [])

# Create a dictionary mapping evaluators to an array of proposals
# (each proposal is a dictionary with proposal name, categories, and link)
evaluators = process_assignments(values)

print(matrixMap)

# Loop through evaluators, creating a spreadsheet for each with
# a README sheet plus one sheet per assigned proposal. As we create 
# them, collect data (evaluator name, spreadsheet ID, and spreadsheet URL)
# to write out to a spreadsheet mapping evaluators to their individual 
# spreadsheets
mapping = [["Name", "Sheet ID", "Sheet Link"]]
for e in evaluators.keys():
  eId = create_one_sheet(e, evaluators[e])
  eUrl = "https://docs.google.com/spreadsheets/d/" + eId + "/edit"
  mapping.append([e, eId, eUrl])

print(matrixMap)

# Now create the mapping spreadsheet
mappingFileId = createSpreadsheet("Evaluator Mappings", TARGET_FOLDER_ID)
# Write out the spreadsheet mapping tab
rangeValue = "Sheet1!A1:C"+str(len(mapping))
sheetService.spreadsheets().values().update(
  spreadsheetId=mappingFileId,
  range=rangeValue,
  valueInputOption='USER_ENTERED',
  body={'values': mapping}
).execute()
# Rename the tab
request = sheetService.spreadsheets().batchUpdate(spreadsheetId=mappingFileId, body={
    'requests': [
      { "updateSheetProperties": {
        "properties": {
          "sheetId": 0,
          "title": 'Sheet Mapping'
        },
        "fields": 'title'
      }}
    ]
})
response = request.execute()
# Create a new tab for tab mapping
request = sheetService.spreadsheets().batchUpdate(spreadsheetId=mappingFileId, body={
    'requests': [{
        'addSheet': {
            'properties': {
                'title': 'Tab Mapping'
            }
        }
    }]
})
response = request.execute()

# Now write out the values
cnt = evaluatorCount
if cnt > 25: # 25 because the first column is the proposal name
  cnt -= 26
  letter = 'A'+ chr(ord('A') + cnt)
else:
  letter = chr(ord('A') + cnt)
rangeValue = "Tab Mapping!A1:"+letter + str(1+len(proposalIndices))
print('rangeValue = ', rangeValue)

sheetService.spreadsheets().values().update(
  spreadsheetId=mappingFileId,
  range=rangeValue,
  valueInputOption='USER_ENTERED',
  body={'values': matrixMap}
).execute()


