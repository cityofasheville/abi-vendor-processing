
import requests
import pandas as pd
import json
from os.path import exists
from csv import reader

secrets = None
inputs = None

# API Key
if exists('./keyfile.json'):
    with open('keyfile.json', 'r') as file:
        secrets = json.load(file)

key = secrets['API_KEY']

# Getting inputs for api for the inputs file
if exists('./inputs.json'):
    with open('inputs.json', 'r') as file:
        inputs = json.load(file)

year = inputs['YEAR']
geography = inputs['GEOGRAPHY']
stateid = inputs['STATEID']
countyid = inputs['COUNTYID']
censustype = inputs['CENSUSTYPE']
numberofyears = inputs['NUMBEROFYEARS']

# Some of the chars in the url differ slightly between the two census types, so this if/else fixes that
if censustype == 'dec':
    secondpart= 'sf'
    numberofyears = '1'
elif censustype == 'acs':
    secondpart= 'acs'
else:
    print("Error: Did not recognize census type from the inputs file. Try 'dec' or 'acs'")
    exit()


#Getting the list of codes for which census tables we want from the csv
with open('CensusCodes.csv', 'r') as read_obj:
    csv_reader = reader(read_obj)
    tablecodes = list(csv_reader)
    tablecodes.pop(0)

#------------------------------
dfTotal = None

#This loop goes through the list of codes and calls the api for each one, then appends the data in a dataframe
for entry in tablecodes:
    code = entry[1]
    #url  = f'https://api.census.gov/data/{year}/{censustype}/{secondpart}{numberofyears}?get={code},NAME&for={geography}:*&in=state:{stateid}%20county:{countyid}&key={key}'
    url  = f'https://api.census.gov/data/{year}/{censustype}/{secondpart}{numberofyears}?get=NAME,{code}&for={geography}:*&in=state:{stateid}%20county:{countyid}&key={key}'    
    response=requests.get(url)
    data = response.json()
    df = pd.DataFrame(data[1:], columns=data[0])

    df['label'] = entry[2]
    df['race'] = entry[3]
    df = df.iloc[: , 1:]
    df.rename(columns={df.columns[0]: "value" }, inplace = True)

    if dfTotal is None:
        dfTotal = df.copy()
    else:
        dfTotal = dfTotal.append(df, ignore_index=True)
    
print('done')    
dfTotal.to_csv('data.csv')

#print(response.text)
#https://api.census.gov/data/2010/dec/sf1?get=PCT012A015,PCT012A119&for=state:01&key=[user key]