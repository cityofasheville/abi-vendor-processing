#from email.mime import base
from typing import final
import requests
import pandas as pd
import urllib.parse
import re
import numpy as np

#Open file, read data
data = pd.read_csv(r'C:\Users\cameronhenshaw\data-processing-scripts\general-data-processing\good_filtered_data.csv')
baseLocator = 'https://arcgis.ashevillenc.gov/arcgis/rest/services/Geolocators/BC_address_unit/GeocodeServer/findAddressCandidates'

#-----------------------------

def get_coords(address, baseLocator):
    #Get data from API
    geolocatorUrl = baseLocator + '?Street=&City=&ZIP=&Single+Line+Input=' + urllib.parse.quote(address) + '&category=&outFields=&maxLocations=&outSR=&searchExtent=&location=&distance=&magicKey=&f=html'
    textResponse = requests.get(geolocatorUrl).text
    
    #Search for the coordinates
    xCoord = re.findall(r'(?<=X: </i> )[^\s]*',textResponse)
    yCoord = re.findall(r'(?<=Y: </i> )[^\s]*',textResponse)

    #Filter through coords to find missing or multiple responses
    if (len(xCoord) ==0):
        return(np.NaN)
    elif len(xCoord) == 1:
        return (xCoord[0] + ', ' + yCoord[0])
    else:
        scores = re.findall(r'(?<=Score: </i> )[^\s]*',textResponse)
        result = all(elem == scores[0] for elem in scores)
        return ('Multiple Matches')
        

#-----------------------------


finalDf = pd.DataFrame(columns=['Id', 'Title', 'Department', 'Address', 'Street', 'City', 'Zip', 'Coords'])

#This loop lets you set how many lines you want to process in onebatch, and how many overall. 

dfList = []
for x in range(0,10):
    data100 = data.tail(100)
    data = data.iloc[:-100 , :] # This line drops the last n rows
    data100['Coords'] = data100.apply(lambda x: get_coords(x.Address, baseLocator), axis=1)
    print('Locating ' + str(x*10))
    finalDf = pd.concat([finalDf, data100])


finalDf.to_csv(r'C:\Users\cameronhenshaw\data-processing-scripts\general-data-processing\data_with_coords.csv')

print('Success')
