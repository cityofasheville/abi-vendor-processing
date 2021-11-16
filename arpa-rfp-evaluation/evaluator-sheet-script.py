
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'arpa-processing-25528ff0b6f2.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
assignmentsSpreadsheetId = '1oy7i8HOhDbxsvsXcwRnrBUV75o3giiS1tWeJQEhP-is'
baseEvaluatorSpreadsheetId = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'
baseEvaluatorSpreadsheetReadmeId = 1990079504
baseEvaluatorSpreadsheetEvaluationId = 1231958865

targetEvaluatorsFolder = '103-m5FGsrgm7P5o4RUylB95LVzBWl8ht'


sheetService = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = sheetService.spreadsheets()

result = sheet.values().get(spreadsheetId=assignmentsSpreadsheetId,range="Eligible Proposals and Assignments!A1:L4").execute()

values = result.get('values', [])

#############################

def read_assignments(values):
    values.pop(0)
    evals = {}
    for row in values:
        for name in row[5:]:
            if name not in evals.keys():
                evals[name] = []
                evals[name].append ({'name': row[1], 'link': row[2]})
            else:
                evals[name].append({'name': row[1], 'link': row[2]})
    return evals

def create_one_sheet(evaluator, evaluatees):
    print('Creating a sheet for ', evaluator, evaluatees)
    # Create the spreadsheet with the name of the evaluator
    # and get the ID
    drive = build('drive', 'v3', credentials=creds)
    file_metadata = {
        'name': evaluator,
        'parents': [targetEvaluatorsFolder],
        'mimeType': 'application/vnd.google-apps.spreadsheet',
    }
    res = drive.files().create(body=file_metadata).execute()
    evaluatorSheetId = res['id']
    # evaluatorSheet = sheetService.spreadsheets().get(spreadsheetId=evaluatorSheetId).execute()
    # print('EVALUATOR SHEET: ')
    # print(evaluatorSheet)

    copy_sheet_to_another_spreadsheet_request_body = {
        # The ID of the spreadsheet to copy the sheet to.
        'destination_spreadsheet_id': evaluatorSheetId
    }

    # Copy over the README sheet
    response = sheetService.spreadsheets().sheets().copyTo(spreadsheetId=baseEvaluatorSpreadsheetId, sheetId=baseEvaluatorSpreadsheetReadmeId, body={
      'destination_spreadsheet_id': evaluatorSheetId
    }).execute()
    print(response)
    readmeSheetId = response['sheetId']

    # Delete the original sheet1
    request = sheetService.spreadsheets().batchUpdate(spreadsheetId=evaluatorSheetId, body={
        'requests': [
          { "deleteSheet": {"sheetId": 0}}
        ]
    })
    response = request.execute()

    # Rename the first sheet to README
    request = sheetService.spreadsheets().batchUpdate(spreadsheetId=evaluatorSheetId, body={
        'requests': [
          { "updateSheetProperties": {
            "properties": {
              "sheetId": readmeSheetId,
              "title": 'README'
            },
            "fields": 'title'
          }}
        ]
    })
    response = request.execute()

    # batch_update_spreadsheet_request_body = {
    #     'requests': [
    #       { "updateSheetProperties": {
    #         "properties": {
    #           "sheetId": response['sheetId'],
    #           "title": 'README'
    #         },
    #         "fields": 'title'
    #       }}
    #     ]
    # }

    # request = sheetService.spreadsheets().batchUpdate(spreadsheetId=evaluatorSheetId, body=batch_update_spreadsheet_request_body)
    # response = request.execute()

    # Now copy over the evaluation sheet for each of the assigned evaluations

evaluators = read_assignments(values)
print(list(evaluators.keys()))
e = list(evaluators.keys())[0]
create_one_sheet(e, evaluators[e])
