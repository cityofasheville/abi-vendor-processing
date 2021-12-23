from os import link
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd
from os.path import exists
import numpy as np
from functools import reduce




SERVICE_ACCOUNT_FILE = None
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']



INPUTS_EVAL_MAPPING_ID =None
OUTPUTS_MASTER_ID = None
INPUTS_SPREADSHEET_ID = None

sheetService = None

def setUpServices():
  global sheetService
  creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )
  sheetService = build('sheets', 'v4', credentials=creds)
  #driveService = build('drive', 'v3', credentials=creds)


def list_tab_links(INPUTS_EVAL_MAPPING_ID):
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=INPUTS_EVAL_MAPPING_ID,range='Tab Mapping!A1:AB').execute()
    tabs = results.get('values', [])
    tab_links_df = pd.DataFrame(tabs)
    tab_links_df.iloc[0,0] = 'Project'
    tab_links_df.columns = tab_links_df.iloc[0]
    tab_links_df.drop(tab_links_df.index[0], inplace=True)
    tab_links_df.reset_index(inplace=True)
    return(tab_links_df)


def build_evaluator_status_list(INPUTS_EVAL_MAPPING_ID):

    tab_links_df = list_tab_links(INPUTS_EVAL_MAPPING_ID)


    # Get spreadsheet links/ids from the spreadsheet
    total_list = []
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=INPUTS_EVAL_MAPPING_ID,range='Sheet Mapping!A2:C').execute()
    link_ss_values = results.get('values', [])

    for thing in link_ss_values:
        id = thing[1]

        sheet = sheetService.spreadsheets()
        sheets = sheet.get(spreadsheetId=id, fields='sheets/properties/title').execute()
        ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
        
        format_list = []

        # Goes through each tab and gets values
        for tab in ranges[1:]:
            results = sheet.values().get(spreadsheetId=id,range=tab +'!A1:E24').execute()
            values = results.get('values', [])
            data = values[6:]

            # two paths, if else
            #if df contains no nulls, calculate stats
            #if df contains nulls, skip that step and add "Not Complete"


            #Make a dataframe, then change the rating values to numbers
            df = pd.DataFrame(data, columns = ["question_num", 'question', 'rating', 'guidance', 'scoring_category'])
            df = df.replace(r'^\s*$', np.nan, regex=True)
            if df['rating'].isnull().values.any():
                numUnanswered = df['rating'].isnull().sum().astype(str)
                print(numUnanswered)
                status = 'Incomplete'
            else:    
                numUnanswered = 0
                status =  'Finished'

            #Grabbing info from list to put into the right output format
            project_name = values[1][1].split(": ",1)[1]
            project_number = project_name[0]
            evaluator = values[0][1].split(": ",1)[1]
            evaluator=evaluator.strip()
            link = thing[2]

            # Using the df from the beginning of this function to look up the links
            # to individual tabs on evaluator sheets. Appending that to the end of the list.
            eval_link = tab_links_df[evaluator].iloc[int(project_number)-1]


            format_list = [evaluator, project_number, project_name, eval_link, numUnanswered, status]
            total_list.append(format_list)
    return(total_list)

def updateSheet(my_list, spreadSheetID, range):
    resource = {
    "majorDimension": "ROWS",
    "values": my_list
}

    sheetService.spreadsheets().values().update(
    spreadsheetId=spreadSheetID,
    range=range,
    body=resource,
    valueInputOption="USER_ENTERED").execute()




inputs = None
if exists('./inputs.json'):
    with open('inputs.json', 'r') as file:
        inputs = json.load(file)
else:
    print('You must create an inputs.json file')
    sys.exit()


INPUTS_EVAL_MAPPING_ID = inputs["INPUTS_EVAL_MAPPING_ID"]
OUTPUTS_MASTER_ID = inputs["OUTPUTS_MASTER_ID"]
INPUTS_SPREADSHEET_ID = inputs['INPUTS_SPREADSHEET_ID']
SERVICE_ACCOUNT_FILE = inputs['SERVICE_ACCOUNT_FILE']


setUpServices()
sheet = sheetService.spreadsheets()


my_list = build_evaluator_status_list(INPUTS_EVAL_MAPPING_ID)

updateSheet(my_list, OUTPUTS_MASTER_ID, "Evaluator Status!A2:AA1000")
