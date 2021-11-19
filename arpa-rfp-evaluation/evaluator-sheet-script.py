
from googleapiclient.discovery import build
import json
import sys
from csv import reader
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'arpa-processing-25528ff0b6f2.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )


# IDs of the various spreadsheets, sheets and folders
assignmentsSpreadsheetId = '1xrEqDmNd0jBAh_vth5ReC0DaUxFhYfKT0asQw-pF4kI' 
baseEvaluatorSpreadsheetId = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'
baseEvaluatorSpreadsheetReadmeId = 2068704255
baseEvaluatorSpreadsheetEvaluationId = 159904982

targetEvaluatorsFolderId = '14_2ov-PiOeSAeFPYuxPuOzayMdx7L4yx'

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
    result = sheetService.spreadsheets().values().get(spreadsheetId=assignmentsSpreadsheetId,range="Evaluators!A1:A100").execute()
    tmp = result.get('values', [])
    evaluatorList = []
    for i in range(len(tmp)):
        evaluatorList.append(tmp[i][0])
    evaluatorCount = len(evaluatorList)

    for i in range(len(evaluatorList)):
        evaluatorIndices[evaluatorList[i]] = i
    print('Count of evaluators: ', evaluatorCount)
    print(evaluatorIndices)

def process_assignments(values):
    global evaluatorCount
    global matrixMap
    values.pop(0)
    evals = {}
    proposalIndex = 0
    for row in values:
        if len(row[1]) == 0:
            continue
        lst = [-1] * (evaluatorCount+1)
        lst[0] = row[1]
        matrixMap.append(lst)
        proposalIndices[row[1]] = proposalIndex
        proposalIndex += 1
        for name in row[5:]:
            if name not in evals.keys():
                evals[name] = []
                evals[name].append ({'name': row[1], 'link': row[2], 'categories': row[3]})
            else:
                evals[name].append({'name': row[1], 'link': row[2], 'categories': row[3]})
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
    evaluatorSheetId = createSpreadsheet(evaluator, targetEvaluatorsFolderId)

    # Copy over the README sheet
    response = copyAndRenameSheet(baseEvaluatorSpreadsheetId,
    baseEvaluatorSpreadsheetReadmeId,
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
        newSheetId = copyAndRenameSheet(baseEvaluatorSpreadsheetId,baseEvaluatorSpreadsheetEvaluationId, evaluatorSheetId, proposal['name'])
        matrixMap[proposalIndices[proposal['name']]][evaluatorIndices[evaluator]+1] = newSheetId
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

result = sheetService.spreadsheets().values().get(spreadsheetId=assignmentsSpreadsheetId,range="Eligible Proposals and Assignments!A1:L100").execute()
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
mappingFileId = createSpreadsheet("Evaluator Mappings", targetEvaluatorsFolderId)
rangeValue = "Sheet1!A1:C"+str(len(mapping))
sheetService.spreadsheets().values().update(
  spreadsheetId=mappingFileId,
  range=rangeValue,
  valueInputOption='USER_ENTERED',
  body={'values': mapping}
).execute()
