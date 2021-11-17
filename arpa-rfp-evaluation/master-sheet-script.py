
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account


SERVICE_ACCOUNT_FILE = 'arpa-processing-202b3d5190f8.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
READ_SPREADSHEET_ID = '1HUMDVAopkvbJDjc1-Cs-ahCFAnl8-FpToPNQ6n38C-I'
WRITE_SPREADSHEET_ID = '1nFyGhmkFGMdq1HzzTIp6aRtTpfIxgWeKN-veUfef_Eg'

service = build('sheets', 'v4', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

#########################################################
# Change tab range when in use!
# Make sure to check category info- this is only taking into account 4 (dummy) categories currently

# I still need to figure out how to iterate through all the sheets. 
# I could either iterate through the folder, or pass a list of the individual sheets. 

# We need to put the project number at the beginning of the title of each sheet
# Do we want the Low/medium/high score to show up on the master sheet, or rather the numbers?


def build_list(READ_SPREADSHEET_ID, all_cat_list):
    # Gets list of tab names
    sheet = service.spreadsheets()
    sheets = sheet.get(spreadsheetId=READ_SPREADSHEET_ID, fields='sheets/properties/title').execute()
    ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
    format_list = []
    # Goes through each tab and gets values
    for tab in ranges[1:3]:
        results = sheet.values().get(spreadsheetId=READ_SPREADSHEET_ID,range=tab +'!A1:E24').execute()
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
            for number in range(0,4):
                short_list.append('no')
            
            #Generates list of categories for the specific project
            project_cat= values[1][1].split(": ")[1]
            project_cat_list = project_cat.split(', ')

            # Goes through the the list of this project's categories to see if there are any matches with
            # the project category columsn. If there is, the value for that column is replaced with
            # 'yes' in the short list
            for category in all_cat_list:
                if category in project_cat_list:
                    index = all_cat_list.index(category)
                    short_list[6 + index] = 'yes'

            format_list.append(short_list)
    return(format_list)


# This bit of code gets the list of categories from the master sheet
results = sheet.values().get(spreadsheetId=WRITE_SPREADSHEET_ID,range='All Data Test!A1:L1').execute()
values = results.get('values', [])

all_cat_list = values[0][6:10] #NEED TO CHANGE THIS TO MORE CATEGORIES WHEN ADDED

# Calls list building function, appends list to master spreadsheet
list_to_append = build_list(READ_SPREADSHEET_ID, all_cat_list)

resource = {
  "majorDimension": "ROWS",
  "values": list_to_append
}

service.spreadsheets().values().append(
  spreadsheetId=WRITE_SPREADSHEET_ID,
  range="All Data Test!A:A",
  body=resource,
  valueInputOption="USER_ENTERED").execute()



#####################


###################
#Take all info and add to one doc

#########################################################
# Take scores and add to summary sheet