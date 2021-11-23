from os import link
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
# Links to spreadsheets generated for each of the evaluators
EVALUATOR_LINKS_SPREADSHEET_ID = '1tAJXTaol2iNrT-ieLYyFPeaX5sUZDRoN3tz6IGR3XNc'
# Master Output Spreadsheet
WRITE_SPREADSHEET_ID = '1nFyGhmkFGMdq1HzzTIp6aRtTpfIxgWeKN-veUfef_Eg'
# Old evaluation sheet with scoring info on it (Master scoring Rubric sheet)
SCORE_SPREADSHEET_ID = '1-AemNS14zBpFWeWCKqqhnbk_34FCEKTh3IfncyIYUxU'
# Spreadsheet with just a list of links to project proposals (Assignment sheet)
PROJECT_LINKS_SPREADSHEET_ID = '1oy7i8HOhDbxsvsXcwRnrBUV75o3giiS1tWeJQEhP-is'
# Spreadsheet with links to individual tabs
EACH_TAB_SPREADSHEET_ID = '1FssUUapqP8cgZBU0t2hMOdHmagCqRBksB-dhOWjQt2k'

# Do you think we have enough spreadsheets yet

service = build('sheets', 'v4', credentials=creds)
drive_service = build('drive', 'v3', credentials=creds)

# Call the Sheets API
sheet = service.spreadsheets()

#########################################################

# The first function creates two dfs, one with links to projects, one with scoring weight info
# Those dfs are passed to the second function, which goes through each sheet and calculates
# the individual score 1 evaluator gives 1 project.
# A list with that info is passed to the third function, which creates the summary list using
# all of the evaluator's scores for each project.

def grab_weights_and_links(link_sheet_id, weight_sheet_id):
    # Gets score weights from the evaluation sheet, and project links, and puts these things into 2
    # dfs to merge with the main summary df later
    sheet = service.spreadsheets()
    results = sheet.values().get(spreadsheetId=weight_sheet_id,range='Score Weighting!C8:D27').execute()
    values = results.get('values', [])
    del values[13]
    del values[6]

    weight_df = pd.DataFrame(values, columns=['weight_in_cat', 'global_weight'])
    weight_df['weight_in_cat'] = weight_df['weight_in_cat'].astype(float)
    weight_df['global_weight'] = weight_df['global_weight'].astype(float)

    # Gets project links from the evaluation assignment sheet
    sheet = service.spreadsheets()
    results = sheet.values().get(spreadsheetId=link_sheet_id,range='Eligible Proposals and Assignments!A2:C').execute()
    values = results.get('values', [])
    link_df = pd.DataFrame(values, columns=['project_number', 'project_name', 'project_link'])
    return(link_df, weight_df)




def build_project_summary_list(weight_df, EVALUATOR_LINKS_SPREADSHEET_ID, EACH_TAB_SPREADSHEET_ID):
    # Creating a dataframe with links to individual tabs to use later
    sheet = service.spreadsheets()
    results = sheet.values().get(spreadsheetId=EACH_TAB_SPREADSHEET_ID,range='Sheet1!A1:C').execute()
    tabs = results.get('values', [])
    tab_links_df = pd.DataFrame(tabs)
    tab_links_df.iloc[0,0] = 'Project'
    tab_links_df.columns = tab_links_df.iloc[0]
    tab_links_df.drop(tab_links_df.index[0], inplace=True)
    tab_links_df.reset_index(inplace=True)



    # Get spreadsheet links/ids from the spreadsheet
    total_list = []
    results = sheet.values().get(spreadsheetId=EVALUATOR_LINKS_SPREADSHEET_ID,range='Copy of Sheet1!A2:B').execute()
    link_ss_values = results.get('values', [])

    for thing in link_ss_values:
        item = thing[1]

        # Gets id and range of tabs for each sheet
        item = item.split('/d/')
        item = item[1].split('/edit')
        id = item[0]

        sheet = service.spreadsheets()
        sheets = sheet.get(spreadsheetId=id, fields='sheets/properties/title').execute()
        ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
        

        format_list = []

        # Goes through each tab and gets values
        for tab in ranges[1:]:
            results = sheet.values().get(spreadsheetId=id,range=tab +'!A1:E24').execute()
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
            # What are the 11, 8, and 9? They're the total of the "weight within category".
            # That makes more sense if you take a look at the scoring sheet- 
            # "Evaluation Score" Score Weighting tab, column C
        
            ECI_score = (cat_weights_global['Equitable Community Impact']/11) * 40
            PPE_score = (cat_weights_global['Project Plan and Evaluation']/8) * 40
            OQ_score = (cat_weights_global['Organizational Qualification']/9) * 20
            total_score = ECI_score + PPE_score + OQ_score 

            #Grabbing info from list to put into the right output format
            project_name = values[0][1].split(": ",1)[1]
            project_number = project_name[1]
            evaluator = values[3][1].split(": ",1)[1]
            evaluator=evaluator.strip()
            link = values[2][1]

            # Using the df from the beginning of this function to look up the links
            # to individual tabs on evaluator sheets. Appending that to the end of the list.

            eval_link = tab_links_df[evaluator].iloc[int(project_number)-1]


            format_list = [project_number, project_name, evaluator, link, total_score, ECI_score, PPE_score, OQ_score, eval_link]
            total_list.append(format_list)

    return(total_list)




def summarize_all_project(my_list, link_df):
    # Creating initial df
    my_df = pd.DataFrame(my_list, columns=['project_number', 'project_name', 'evaluator', 'link_to_proposal', 
                    'total_score', 'ECI_score', 'PPE_score', 'OQ_score', 'eval_link'])
    my_df = my_df.round(2)


    #Calculating mean and median, renaming columsn and resetting index (so that project #s show up when converted to list)
    summary_df = pd.DataFrame(my_df.groupby(['project_number', 'project_name', 'link_to_proposal'])['total_score', 'ECI_score', 'PPE_score', 'OQ_score'].mean())
    summary_df = summary_df.reset_index()

    median_df = pd.DataFrame(my_df.groupby(['project_name'])['total_score'].median())
    median_df = median_df.rename({'total_score':'median_score'}, axis=1)

    # Creating string of all scores per project
    my_df['total_score'] = my_df['total_score'].astype(str)
    individual_score_list_df = pd.DataFrame(my_df.groupby(['project_name'])['total_score'].apply(', '.join).reset_index())
    individual_score_list_df= individual_score_list_df.rename({'total_score':'score_list'}, axis=1)

    # Creating string of all links
    eval_links_df = pd.DataFrame(my_df.groupby(['project_name'])['eval_link'].apply(', '.join).reset_index())
  

    # Merging the various dfs to create 1 summary
    summary_df=summary_df.merge(median_df, on='project_name')
    summary_df=summary_df.merge(individual_score_list_df, on='project_name')
    summary_df = summary_df.merge(link_df[['project_number', 'project_link']], on='project_number', how='left')
    summary_df=summary_df.merge(eval_links_df, on='project_name')



    # Reordering columns so the info is in the correct order in the list 
    summary_df = summary_df[['project_number', 'project_name', 'project_link', 'median_score', 'total_score',
       'ECI_score', 'PPE_score', 'OQ_score', 'score_list', 'link_to_proposal', 'eval_link']]
    summary_df = summary_df.drop('link_to_proposal', 1)
    summary_df = summary_df.round(2)

    final_list = summary_df.values.tolist()

    # evals is string of the links to evaluation tabs
    # I'm making it a list and appending it to the final_list, so that each link
    # will end up in a separate column
    for entry in final_list:
        evals = entry.pop()
        evals = list(evals.split(', '))
        entry.extend(evals)
        
    return(final_list)



###########################################################################

link_df, weight_df = grab_weights_and_links(PROJECT_LINKS_SPREADSHEET_ID, SCORE_SPREADSHEET_ID)


# Calls list building function
all_project_scores = build_project_summary_list(weight_df, EVALUATOR_LINKS_SPREADSHEET_ID, EACH_TAB_SPREADSHEET_ID)
list_to_append = summarize_all_project(all_project_scores, link_df)


#Update Spreadsheet
resource = {
  "majorDimension": "ROWS",
  "values": list_to_append
}

service.spreadsheets().values().update(
  spreadsheetId=WRITE_SPREADSHEET_ID,
  range="Summary!A2:AA1000",
  body=resource,
  valueInputOption="USER_ENTERED").execute()

print('boop')


#############################################


    # This bit of code gets list of files in google drive folder, grabs IDs
    # We do not need it currently
    # but it took me awhile to figure out
    # and I can't bring myself to delete it yet


    #results = drive_service.files().list(q = "'1iE_EuNlJqSA5fTFXIGI1WKtdIbT7IbdU' in parents", pageSize=10, fields="nextPageToken, files(id, name)").execute()
    #items = results.get('files', [])
    #for f in range(0, len(items)):
    #    fId = items[f].get('id')
