from googleapiclient.discovery import build
import json
import sys
import time
from csv import reader
from google.oauth2 import service_account
import pandas as pd
from os.path import exists
import numpy as np



SERVICE_ACCOUNT_FILE = None
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

# The ID and range of a sample spreadsheet.
OUTPUTS_MASTER_ID = None
INPUTS_EVAL_MAPPING_ID = None

sheetService = None


#########################################################


def setUpServices():
  global sheetService
  creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )
  sheetService = build('sheets', 'v4', credentials=creds)

def stripLower(lst):
    return list(map(lambda itm: itm.strip().lower() if itm else None, lst))

def getSheetTitles(sheet, spreadsheetId):
    sheets = sheet.get(spreadsheetId=spreadsheetId, fields='sheets/properties').execute()
    return [{"title": sheet['properties']['title'], "id": sheet['properties']['sheetId']} for sheet in sheets['sheets']]

evaluationStatus = []

def build_list(INPUTS_EVAL_MAPPING_ID):
    allQuestions = []
    sheet = sheetService.spreadsheets()

    # Read the mapping of evaluator to spreadsheet ID/URL. We just use the spreadsheet ID
    total_list = []
    results = sheet.values().get(spreadsheetId=INPUTS_EVAL_MAPPING_ID,range='Sheet Mapping!A2:C51').execute()
    evaluatorMap = results.get('values', [])
    # For each evaluator, iterate through each evaluation to get categories and answers
    for evaluatorEntry in evaluatorMap:
        evaluator = evaluatorEntry[0]
        print('Reading evaluator ' + evaluator)
        id = evaluatorEntry[1] # ID of this evaluator's spreadsheet
        link = evaluatorEntry[2]
        tabs = getSheetTitles(sheet, id)
        for tab in tabs[1:]: # Each tab is one evaluation
            print(' Working on tab ' + tab['title'] + ' (' + str(tab['id']) + ')')
            tabLink = link + '#gid=' + str(tab['id'])
            values = sheet.values().get(spreadsheetId=id,range=tab['title'] +'!A1:C51').execute().get('values', [])
            projectName = values[4][2]
            applicantName = values[3][2]

            countResponses = 0
            holdQuestions = []
            # Goes through each row. For each, builds a list of the needed values
            noQuestionRows = [11, 12, 15, 20, 24, 29, 34, 35, 41, 43]
            scoreTotal = 0
            for row in range(9,46):
                if row in noQuestionRows:
                    pass
                elif len(values[row])<3:
                    pass
                else:
                    qNumber = values[row][1]
                    #qCategory = values[row][4]
                    response = values[row][2]
                    if response == "High":
                        responseScore = 3
                    elif response == "Medium":
                        responseScore = 2
                    elif response == "Low":
                        responseScore = 1
                    elif response == "None":
                        responseScore = 0

                    if response:
                        countResponses += 1
                    short_list = [projectName, link, qNumber, response, responseScore]
                    allQuestions.append(short_list)
                    scoreTotal += responseScore

            status = None
            if countResponses == 27:
                status = 'Complete'
            else:
                status = 'Incomplete'
            evaluationStatus.append([applicantName, evaluator, projectName, countResponses, scoreTotal, status])
            time.sleep(1) # To deal with Google API quotas

    return allQuestions, evaluationStatus

def build_score_list(evaluationStatus):
    # Filtering so that incomplete scores are not counted in the aggregation
    df = pd.DataFrame(evaluationStatus, columns = ['applicantName', 'evaluator', 'projectName', 'countResponses', 'scoreTotal', 'status'])
    filterDf = df[~df['status'].str.contains("Incomplete")]
    groupedByScores = pd.DataFrame(filterDf.groupby(['applicantName', 'projectName']).agg(values=('scoreTotal',lambda x: list(x)), mean=('scoreTotal','mean'), median=('scoreTotal','median'), range=('scoreTotal', np.ptp)))
    groupedByScores.reset_index(inplace=True)
    groupedByScores['values'] = groupedByScores['values'].apply(lambda x: ', '.join(map(str, x)))

    # Aggregating to count the number of complete and incomplete values
    aggregateScoreStatus = df.groupby(['applicantName', 'projectName'])['status'].value_counts().unstack(level=2)
    aggregateScoreStatus.fillna(0, inplace=True)
    mergedDf = pd.merge(aggregateScoreStatus, groupedByScores, how='left', left_on=['applicantName', 'projectName'], right_on = ['applicantName', 'projectName'])

    mergedDf.fillna('not available', inplace=True)
    groupedList = mergedDf.values.tolist()
    return(groupedList)

def build_incomplete_list(evaluationStatus):
        df = pd.DataFrame(evaluationStatus, columns = ['applicantName', 'evaluator', 'projectName', 'countResponses', 'scoreTotal', 'status'])
        filterDf = df[df['status'].str.contains("Incomplete")]
        grouped = pd.DataFrame(filterDf.groupby(['evaluator']).agg(incompleteEvaluations=('applicantName',lambda x: list(x))))
        print(grouped)
        grouped.reset_index(inplace=True)
        grouped['incompleteEvaluations'] = grouped['incompleteEvaluations'].apply(lambda x: ', '.join(map(str, x)))
        groupedList = grouped.values.tolist()
        print(groupedList)
        return(groupedList)



############################ Main Program Start

#Open Json
inputs = None
if exists('spf-evaluation-processing/inputs.json'):
    with open('spf-evaluation-processing/inputs.json', 'r') as file:
        inputs = json.load(file)
else:
    print('You must create an inputs.json file')
    if exists('./inputs.json'):
        print('idk man')
    sys.exit()

# Set global variables
OUTPUTS_MASTER_ID = inputs["OUTPUTS_MASTER_ID"]
INPUTS_EVAL_MAPPING_ID = inputs["INPUTS_EVAL_MAPPING_ID"]
SERVICE_ACCOUNT_FILE = inputs["SERVICE_ACCOUNT_FILE"]

setUpServices()


# Calls list building function, creates the list to append to the spreadsheet
questionList, scoreList = build_list(INPUTS_EVAL_MAPPING_ID)
groupedList = build_score_list(scoreList)
incompleteList = build_incomplete_list(scoreList)

# Update sheet
resource = {
  "majorDimension": "ROWS",
  "values": questionList
}

sheetService.spreadsheets().values().update(
  spreadsheetId=OUTPUTS_MASTER_ID,
  range="All Responses!A2:AA10000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()

# Update sheet
resource = {
  "majorDimension": "ROWS",
  "values": groupedList
}

sheetService.spreadsheets().values().update(
  spreadsheetId=OUTPUTS_MASTER_ID,
  range="Totals!A2:AA10000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()

resource = {
  "majorDimension": "ROWS",
  "values": incompleteList
}

sheetService.spreadsheets().values().update(
  spreadsheetId=OUTPUTS_MASTER_ID,
  range="Incomplete list!A2:AA10000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()