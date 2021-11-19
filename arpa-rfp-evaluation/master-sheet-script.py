from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd


SERVICE_ACCOUNT_FILE = 'arpa-processing-202b3d5190f8.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
WRITE_SPREADSHEET_ID = '1nFyGhmkFGMdq1HzzTIp6aRtTpfIxgWeKN-veUfef_Eg'
EVALUATOR_LINKS_SPREADSHEET_ID = '1tAJXTaol2iNrT-ieLYyFPeaX5sUZDRoN3tz6IGR3XNc'

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

#########################################################

# Script assumes project number is in the project name in each evaluator sheet.
# May need to change this


def build_list(all_cat_list, EVALUATOR_LINKS_SPREADSHEET_ID):
    format_list = []
    sheet = service.spreadsheets()


    # Reads spreadsheet to get list from the spreadsheet
    total_list = []
    results = sheet.values().get(spreadsheetId=EVALUATOR_LINKS_SPREADSHEET_ID,range='Copy of Sheet1!A2:B').execute()
    link_ss_values = results.get('values', [])

    for thing in link_ss_values:
        item = thing[1]

        # Gets ids and range of tabs for each sheet
        item = item.split('/d/')
        item = item[1].split('/edit')
        id = item[0]

        sheets = sheet.get(spreadsheetId=id, fields='sheets/properties/title').execute()
        ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
        
       
        
        for tab in ranges[1:]:
            results = sheet.values().get(spreadsheetId=id,range=tab +'!A1:R24').execute()
            values = results.get('values', [])
            # Goes through each row. For each, builds a list of the needed values
            for number in range(6,24):
                evaluator = values[3][1].split(": ",1)[1] 
                q_number = values[number][0]
                project_name = values[0][1].split(": ",1)[1] 
                project_number = project_name[1]
                link = values[2][1]
                question_cat = values[number][4]
                response = values[number][2]
                short_list = [evaluator, project_number, link, q_number, question_cat, response]
                
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


# This bit of code gets the list of categories from the master sheet, which is passed
# into the function
results = sheet.values().get(spreadsheetId=WRITE_SPREADSHEET_ID,range='All Data Test!A1:R1').execute()
values = results.get('values', [])

all_cat_list = values[0][6:16] #This range gets a list of categories from the spreadsheet
all_cat_list_fixed = []

for item in all_cat_list:
  j = item.strip().lower()
  all_cat_list_fixed.append(j)

all_cat_list = all_cat_list_fixed


# Calls list building function, creates the list to append to the spreadsheet
list_to_append = build_list(all_cat_list, EVALUATOR_LINKS_SPREADSHEET_ID)


# Update sheet
resource = {
  "majorDimension": "ROWS",
  "values": list_to_append
}

service.spreadsheets().values().update(
  spreadsheetId=WRITE_SPREADSHEET_ID,
  range="All Data Test!A2:AA1000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()



