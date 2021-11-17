
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd


SERVICE_ACCOUNT_FILE = 'arpa-processing-202b3d5190f8.json'
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']

creds = service_account.Credentials.from_service_account_file(
        SERVICE_ACCOUNT_FILE, scopes=SCOPES)


# The ID and range of a sample spreadsheet.
READ_SPREADSHEET_ID = '1HUMDVAopkvbJDjc1-Cs-ahCFAnl8-FpToPNQ6n38C-I'
WRITE_SPREADSHEET_ID = '1nFyGhmkFGMdq1HzzTIp6aRtTpfIxgWeKN-veUfef_Eg'
SCORE_SPREADSHEET_ID = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'

service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

#########################################################
# Change tab range when in use!

# We need to put the project number at the beginning of the title of each sheet
# Do we want the Low/medium/high score to show up on the master sheet, or rather the numbers?

# The function builds a df from one evaluation tab. It calculate the score for that one project, for that
# one evaluator. When it is returned, it's appended to another df, which then calculates the mean/median/other score for
# all the evaluations for that project.


def build_project_summary_list(weight_df):
    total_list = []
    #Gets list of files in google drive folder, grabs IDs
    results = drive_service.files().list(q = "'1iE_EuNlJqSA5fTFXIGI1WKtdIbT7IbdU' in parents", pageSize=10, fields="nextPageToken, files(id, name)").execute()
    items = results.get('files', [])
    for f in range(0, len(items)):
        fId = items[f].get('id')

    # Gets range of tabs for each sheet
        sheet = service.spreadsheets()

        sheets = sheet.get(spreadsheetId=fId, fields='sheets/properties/title').execute()
        ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
        format_list = []

        # Goes through each tab and gets values
        for tab in ranges[1:]:
            results = sheet.values().get(spreadsheetId=fId,range=tab +'!A1:E24').execute()
            values = results.get('values', [])
            data = values[6:]

            #Make a dataframe, then change the rating values to numbers
            df = pd.DataFrame(data, columns = ["question_num", 'question', 'rating', 'guidance', 'scoring_category'])
            df["rating"] = df["rating"].str.lower()
            df["rating"].replace({"none": 0, "low": 1/3, "medium": 2/3, 'high':1}, inplace=True)

            #add more columns
            df['total_category_weight'] = df['scoring_category']
            df["total_category_weight"].replace({"Equitable Community Impact": 40, "Project Plan and Evaluation": 40, "Organizational Qualification": 20}, inplace=True)

            # Adding df of scoring weights to the df I just created
            df = pd.concat([df, weight_df], axis=1)
            df['category_rating'] = df['rating'].astype(float) * df['weight_in_cat'].astype(float)
            
            #Calc category value by global weight
            cat_weights_global = df.groupby(['scoring_category']).category_rating.sum()

            #Formatting output
            ECI_score = (cat_weights_global['Equitable Community Impact']/11) * 100
            PPE_score = (cat_weights_global['Project Plan and Evaluation']/8) * 100
            OQ_score = (cat_weights_global['Organizational Qualification']/9) * 100
            total_score = (ECI_score * .4) + (PPE_score * .4) + (OQ_score * .2)

            project_name = values[0][1].split(": ",1)[1]
            project_number = project_name[1]
            evaluator = values[3][1].split(": ",1)[1]
            link = values[2][1]

            format_list = [project_number, project_name, evaluator, link, total_score, ECI_score, PPE_score, OQ_score]
            total_list.append(format_list)
            #print(cat_weights_global)

    print(total_list)
    return(total_list)


def summarize_all_project(my_list):
    my_df = pd.DataFrame(my_list, columns=['project_number', 'project_name', 'evaluator', 'link_to_proposal', 
                    'total_score', 'ECI_score', 'PPE_score', 'OQ_score'])

    summary_df = pd.DataFrame(my_df.groupby(['project_number', 'project_name', 'link_to_proposal'])['total_score', 'ECI_score', 'PPE_score', 'OQ_score'].apply(lambda x : x.astype(int).mean()))
    median_df = pd.DataFrame(my_df.groupby(['project_name'])['total_score'].median())
    median_df = median_df.rename({'total_score':'median_score'}, axis=1)
    individual_score_list_df = pd.DataFrame(my_df.groupby(['project_name'])['total_score'].apply(list))
    individual_score_list_df= individual_score_list_df.rename({'total_score':'score_list'}, axis=1)

    summary_df=summary_df.merge(median_df, on='project_name')
    summary_df=summary_df.merge(individual_score_list_df, on='project_name')


    final_list = summary_df.values.tolist()
    print(summary_df)


    #print(pd.DataFrame(meh))



# Gets the weights for each question. Global_weight isn't currently used
sheet = service.spreadsheets()
results = sheet.values().get(spreadsheetId=SCORE_SPREADSHEET_ID,range='Score Weighting!C8:D27').execute()
values = results.get('values', [])
del values[13]
del values[6]

weight_df = pd.DataFrame(values, columns=['weight_in_cat', 'global_weight'])
weight_df['weight_in_cat'] = weight_df['weight_in_cat'].astype(float)
weight_df['global_weight'] = weight_df['global_weight'].astype(float)



# Calls list building function, appends list to master spreadsheet
all_project_scores = build_project_summary_list(weight_df)
voila = summarize_all_project(all_project_scores)


#resource = {
#  "majorDimension": "ROWS",
#  "values": list_to_append
#}

#service.spreadsheets().values().append(
#  spreadsheetId=WRITE_SPREADSHEET_ID,
#  range="All Data Test!A:A",
#  body=resource,
#  valueInputOption="USER_ENTERED").execute()
