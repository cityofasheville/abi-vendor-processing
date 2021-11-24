from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd
from os.path import exists



SERVICE_ACCOUNT_FILE = 'arpa-processing-202b3d5190f8.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
OUTPUTS_MASTER_ID = None
INPUTS_EVAL_MAPPING_ID = None

sheetService = None


#########################################################


def setUpServices():
  global sheetService
  creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )
  sheetService = build('sheets', 'v4', credentials=creds)
  #driveService = build('drive', 'v3', credentials=creds)


def build_list(all_cat_list, INPUTS_EVAL_MAPPING_ID):
    format_list = []
    sheet = sheetService.spreadsheets()

    # Reads spreadsheet to get list from the spreadsheet
    total_list = []
    results = sheet.values().get(spreadsheetId=INPUTS_EVAL_MAPPING_ID,range='Sheet Mapping!A2:C').execute()
    link_ss_values = results.get('values', [])

    for entry in link_ss_values:
        id = entry[1]
        sheets = sheet.get(spreadsheetId=id, fields='sheets/properties/title').execute()
        ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
        
        for tab in ranges[1:]:
            results = sheet.values().get(spreadsheetId=id,range=tab +'!A1:R24').execute()
            values = results.get('values', [])
            # Goes through each row. For each, builds a list of the needed values
            for number in range(6,24):
                evaluator = values[0][1].split(": ",1)[1] 
                q_number = values[number][0]
                project_name = values[1][1].split(": ",1)[1] 
                project_number = project_name[0]
                link = entry[2]
                question_cat = values[number][4]
                response = values[number][2]
                short_list = [evaluator, project_number, project_name, link, q_number, question_cat, response]
                
                # Adds "no" responses for the categories
                for number in range(0,10):
                    short_list.append('no')
                
                #Generates list of categories for the specific project
                project_cat_fixed = []
                project_cat= values[1][1].split(": ")[1]
                project_cat_list = project_cat.split(', ')
                for item in project_cat_list:
                  j = item.strip().lower()
                  project_cat_fixed.append(j)
                project_cat_list = project_cat_fixed

                # Goes through the the list of this project's categories to see if there are any matches with
                # the project category columsn. If there is, the value for that column is replaced with
                # 'yes' in the short list
                for category in all_cat_list:
                    if category in project_cat_list:
                        index = all_cat_list.index(category)
                        short_list[6 + index] = 'yes'

                format_list.append(short_list)
    return(format_list)


def create_category_list(OUTPUTS_MASTER_ID):
  results = sheet.values().get(spreadsheetId=OUTPUTS_MASTER_ID,range='All Data!A1:R1').execute()
  values = results.get('values', [])

  all_cat_list = values[0][7:17] #This range gets a list of categories from the spreadsheet
  all_cat_list_fixed = []

  for item in all_cat_list:
    j = item.strip().lower()
    all_cat_list_fixed.append(j)

  all_cat_list = all_cat_list_fixed
  return(all_cat_list)

############################ Main Program Start

#Open Json
inputs = None
if exists('./inputs.json'):
    with open('inputs.json', 'r') as file:
        inputs = json.load(file)
else:
    print('You must create an inputs.json file')
    sys.exit()

# Set const values
OUTPUTS_MASTER_ID = inputs["OUTPUTS_MASTER_ID"]
INPUTS_EVAL_MAPPING_ID = inputs["INPUTS_EVAL_MAPPING_ID"]

setUpServices()
sheet = sheetService.spreadsheets()

# Gets the list of categories from the master sheet, which is passed
# into the function

all_cat_list = create_category_list(OUTPUTS_MASTER_ID)

# Calls list building function, creates the list to append to the spreadsheet
list_to_append = build_list(all_cat_list, INPUTS_EVAL_MAPPING_ID)


# Update sheet
resource = {
  "majorDimension": "ROWS",
  "values": list_to_append
}

sheetService.spreadsheets().values().update(
  spreadsheetId=OUTPUTS_MASTER_ID,
  range="All Data!A2:AA1000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()



