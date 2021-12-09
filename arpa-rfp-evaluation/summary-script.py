from os import link
from googleapiclient.discovery import build
import json
from csv import reader
from google.oauth2 import service_account
import pandas as pd
from os.path import exists
import numpy as np
from functools import reduce
import time

SERVICE_ACCOUNT_FILE = None
SCOPES = ['https://www.googleapis.com/auth/spreadsheets', 'https://www.googleapis.com/auth/drive']


INPUTS_EVAL_MAPPING_ID =None
OUTPUTS_MASTER_ID = None
INPUTS_SPREADSHEET_ID = None

sheetService = None

#########################################################



def setUpServices():
  global sheetService
  creds = service_account.Credentials.from_service_account_file( SERVICE_ACCOUNT_FILE, scopes=SCOPES )
  sheetService = build('sheets', 'v4', credentials=creds)


def grab_weights_and_links(inputSpreadsheetId):
    # Gets score weights from the evaluation sheet, and project links, and puts these things into 2
    # dfs to merge with the main summary df later
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=inputSpreadsheetId,range='Score Weighting!C8:D27').execute()
    values = results.get('values', [])
    del values[13]
    del values[6]

    weight_df = pd.DataFrame(values, columns=['weight_in_cat', 'global_weight'])
    weight_df['weight_in_cat'] = weight_df['weight_in_cat'].astype(float)
    weight_df['global_weight'] = weight_df['global_weight'].astype(float)

    # Gets project links from the evaluation assignment sheet
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=inputSpreadsheetId,range='Eligible Proposals and Assignments!A2:C').execute()
    values = results.get('values', [])
    links_df = pd.DataFrame(values, columns=['project_number', 'project_name', 'project_link'])

    return(links_df, weight_df)


def list_tab_links(evaluationMappingSheetId):
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=evaluationMappingSheetId,range='Tab Mapping!A1:AB').execute()
    tabs = results.get('values', [])
    tab_links_df = pd.DataFrame(tabs)
    tab_links_df.iloc[0,0] = 'Project'
    tab_links_df.columns = tab_links_df.iloc[0]
    tab_links_df.drop(tab_links_df.index[0], inplace=True)
    tab_links_df.reset_index(inplace=True)
    return(tab_links_df)


def build_project_summary_list(links_df, weight_df, evaluationMappingSheetId):
    tab_links_df = list_tab_links(evaluationMappingSheetId)
    # Get spreadsheet links/ids from the spreadsheet
    total_list = []
    sheet = sheetService.spreadsheets()
    results = sheet.values().get(spreadsheetId=evaluationMappingSheetId,range='Sheet Mapping!A2:C').execute()
    link_ss_values = results.get('values', [])

    for thing in link_ss_values:
        id = thing[1]
        print('   Sheet ' + thing[0])
        sheet = sheetService.spreadsheets()
        sheets = sheet.get(spreadsheetId=id, fields='sheets/properties/title').execute()
        ranges = [sheet['properties']['title'] for sheet in sheets['sheets']]
        
        format_list = []

        # Goes through each tab and gets values
        for tab in ranges[1:]:
            print ('       Tab ' + tab)
            results = sheet.values().get(spreadsheetId=id,range=tab +'!A1:E24').execute()
            values = results.get('values', [])
            data = values[6:]

            #Make a dataframe, then change the rating values to numbers
            df = pd.DataFrame(data, columns = ["question_num", 'question', 'rating', 'guidance', 'scoring_category'])
            df = df.replace(r'^\s*$', np.nan, regex=True)
            if df['rating'].isnull().values.any():
                ECI_score = "Not Complete"
                PPE_score = "Not Complete"
                OQ_score = "Not Complete"
                total_score = "Not Complete"
            else:    
                #df["rating"] = df[~df["rating"].isnull()].str.lower()
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
                #What are the 11, 8, and 9? They're the total of the "weight within category".
                #That makes more sense if you take a look at the scoring sheet- 
                #"Evaluation Score" Score Weighting tab, column 
                
                ECI_score = (cat_weights_global['Equitable Community Impact']/11) * 40
                PPE_score = (cat_weights_global['Project Plan and Evaluation']/8) * 40
                OQ_score = (cat_weights_global['Organizational Qualification']/9) * 20
                total_score = round(ECI_score + PPE_score + OQ_score, 2)

            #Grabbing info from list to put into the right output format
            project_name = values[1][1].split(": ",1)[1]
            project_number = project_name.split(' ')[0]
            evaluator = values[0][1].split(": ",1)[1]
            evaluator=evaluator.strip()
            link = thing[2]

            # Using the df from the beginning of this function to look up the links
            # to individual tabs on evaluator sheets. Appending that to the end of the list.
            eval_link = tab_links_df[evaluator].iloc[int(project_number)-1]

            format_list = [project_number, project_name, evaluator, link, total_score, ECI_score, PPE_score, OQ_score, eval_link]
            total_list.append(format_list)
            time.sleep(1)
        time.sleep(3)
    return(total_list)


def maxMinDifference(df):
    #Get the links dataframe, merge the two together to be able to output links for each project
    df.merge(links_df['project_number'], on='project_number', how='left')

    #Calculate the difference between the min and max score for each column
    maxMinDF = pd.DataFrame(df.groupby(['project_number', 'project_name'])['total_score'].agg(['max','min']))
    maxMinDF['totalScoreVaries'] = maxMinDF['max'] - maxMinDF['min']
    
    ECIMaxMinDF = pd.DataFrame(df.groupby(['project_number', 'project_name'])['ECI_score'].agg(['max','min']))
    ECIMaxMinDF['ECIScoreVaries'] = maxMinDF['max'] - maxMinDF['min']

    PPEMaxMinDF = pd.DataFrame(df.groupby(['project_number', 'project_name'])['PPE_score'].agg(['max','min']))
    PPEMaxMinDF['PPEScoreVaries'] = maxMinDF['max'] - maxMinDF['min']

    OQMaxMinDF = pd.DataFrame(df.groupby(['project_number', 'project_name'])['OQ_score'].agg(['max','min']))
    OQMaxMinDF['OQScoreVaries'] = maxMinDF['max'] - maxMinDF['min']

    #Merge all these calculations together into one dataframe
    maxMinDF = maxMinDF.merge(ECIMaxMinDF['ECIScoreVaries'], on=['project_number', 'project_name'])
    maxMinDF = maxMinDF.merge(PPEMaxMinDF['PPEScoreVaries'], on=['project_number', 'project_name'])
    maxMinDF = maxMinDF.merge(OQMaxMinDF['OQScoreVaries'], on=['project_number', 'project_name'])
    maxMinDF.drop(['max', 'min'], axis=1, inplace=True)
    columnList = ['totalScoreVaries', 'ECIScoreVaries', 'PPEScoreVaries', 'OQScoreVaries']
    #If the different is greater than 50, "True" is assigned, otherwise np.nan. This is so we can use .dropna to
    #drop the rows which have all np.nans in them.
    for entry in columnList:
        maxMinDF[maxMinDF[entry] >= 50] = True
        maxMinDF[maxMinDF[entry] !=True] = np.nan
    maxMinDF = maxMinDF.dropna( how='all', subset=['totalScoreVaries', 'ECIScoreVaries', 'PPEScoreVaries', 'OQScoreVaries'])
    maxMinDF = maxMinDF.replace(np.nan, '')
    maxMinDF = maxMinDF.reset_index()
    print(maxMinDF)
    maxMinList = maxMinDF.values.tolist()
    return(maxMinList)


def summarize_all_project(my_list, links_df):
    # Creating initial df
    my_df = pd.DataFrame(my_list, columns=['project_number', 'project_name', 'evaluator', 'link_to_proposal', 
                    'total_score', 'ECI_score', 'PPE_score', 'OQ_score', 'eval_link'])
    my_df = my_df.round(2)

    #Calculating mean and median, renaming columsn and resetting index (so that project #s show up when converted to list)
    numericScoreDF = my_df[pd.to_numeric(my_df['total_score'], errors='coerce').notnull()]
    numericScoreDF['total_score'] = numericScoreDF['total_score'].astype(float)
    numericScoreDF['ECI_score'] = numericScoreDF['ECI_score'].astype(float)
    numericScoreDF['PPE_score'] = numericScoreDF['PPE_score'].astype(float)
    numericScoreDF['OQ_score'] = numericScoreDF['OQ_score'].astype(float)
        
    maxMinList = maxMinDifference(numericScoreDF)

    summary_df = pd.DataFrame(numericScoreDF.groupby(['project_number', 'project_name'])['total_score', 'ECI_score', 'PPE_score', 'OQ_score'].mean())
    summary_df = summary_df.reset_index()

    median_df = pd.DataFrame(numericScoreDF.groupby(['project_name'])['total_score'].median())
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
    summary_df = summary_df.merge(links_df[['project_number', 'project_link']], on='project_number', how='left')
    summary_df=summary_df.merge(eval_links_df, on='project_name')

    # Reordering columns so the info is in the correct order in the list 
    summary_df = summary_df[['project_number', 'project_name', 'project_link', 'median_score', 'total_score',
       'ECI_score', 'PPE_score', 'OQ_score', 'score_list', 'eval_link']]
    summary_df = summary_df.round(2)

    final_list = summary_df.values.tolist()

    # evals is string of the links to evaluation tabs
    # I'm making it a list and appending it to the final_list, so that each link
    # will end up in a separate column
    for entry in final_list:
        evals = entry.pop()
        evals = list(evals.split(', '))
        entry.extend(evals)
        
    return(final_list, maxMinList)

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

###########################################################################


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

print('Set up services')
setUpServices()
sheet = sheetService.spreadsheets()

print('Load weights')
links_df, weight_df = grab_weights_and_links(INPUTS_SPREADSHEET_ID)


# Calls list building function
print('Build project summary list')
all_project_scores = build_project_summary_list(links_df, weight_df, INPUTS_EVAL_MAPPING_ID)
print('Summarize all the projects')
list_to_append, maxMinList = summarize_all_project(all_project_scores, links_df)


updateSheet(list_to_append, OUTPUTS_MASTER_ID, "Summary!A2:AA1000")
updateSheet(maxMinList, OUTPUTS_MASTER_ID, "Potential Issues!A3:AA1000")

print('Finished, Party time')

