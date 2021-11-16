
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'arpa-processing-25528ff0b6f2.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )


# IDs of the various spreadsheets, sheets and folders
assignmentsSpreadsheetId = '1oy7i8HOhDbxsvsXcwRnrBUV75o3giiS1tWeJQEhP-is'
baseEvaluatorSpreadsheetId = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'
baseEvaluatorSpreadsheetReadmeId = 1990079504
baseEvaluatorSpreadsheetEvaluationId = 1231958865

targetEvaluatorsFolderId = '103-m5FGsrgm7P5o4RUylB95LVzBWl8ht'

sheetService = build('sheets', 'v4', credentials=creds)
driveService = build('drive', 'v3', credentials=creds)

#############################

def process_assignments(values):
    values.pop(0)
    evals = {}
    for row in values:
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

def create_one_sheet(evaluator, proposals):
    print('Creating a spreadsheet for ', evaluator)
    # Create the spreadsheet with the name of the evaluator
    file_metadata = {
        'name': evaluator,
        'parents': [targetEvaluatorsFolderId],
        'mimeType': 'application/vnd.google-apps.spreadsheet',
    }
    res = driveService.files().create(body=file_metadata).execute()
    evaluatorSheetId = res['id']

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
        copyAndRenameSheet(baseEvaluatorSpreadsheetId,baseEvaluatorSpreadsheetEvaluationId, evaluatorSheetId, proposal['name'])

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

##
## Main program
##

result = sheetService.spreadsheets().values().get(spreadsheetId=assignmentsSpreadsheetId,range="Eligible Proposals and Assignments!A1:L4").execute()
values = result.get('values', [])
evaluators = process_assignments(values)
for e in evaluators.keys():
  create_one_sheet(e, evaluators[e])
