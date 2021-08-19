import csv
import sys
import os

# DOT DATA FILE INPUT COLUMNS
# 0 Company Name	
# 1 Mailing Address
# 2 Mailing City
# 3 Mailing State
# 4 Mailing Zip
# 5 Physical Address
# 6 Physical City
# 7 Physical State
# 8 Physical Zip
# 9 Home County & Division
# 10 Contact Name
# 11 Phone
# 12 Fax
# 13 Email
# 14 Reporting Number
# 15 Firm Type
# 16 Certifications
# 17 Prequalification Status
# 18 Construction Work Codes
# 19 SBE Work codes
# 20 Consulting Disciplines
# 21 NAICS
# 22 Work Locations
#
# OUTPUT COLUMNS
#  vendor_number
#  vendor_name
#  commodity_code
#  commodity_code_description
#  coa_certified
#  bbe
#  wbe
#  dbe
#  hbe
#  aibe
#  dobe
#  aabe
#  sbe
#  hub
#  ncdot
#  vendor_city
#  vendor_state
#  vendor_zip
#  contact_name
#  contact_phone
#  contact_email
#  email_addresses
#  vendor_county

includedCounties = ['BUNCOMBE', 'MADISON', 'HENDERSON', 'HAYWOOD', 'JACKSON', 'TRANSYLVANIA', 'POLK', 'RUTHERFORD', 'MCDOWELL', 'YANCEY']

# The -s is to suppress annoying error messages from xlrd
argc = len(sys.argv)
if (argc < 3 or argc > 4):
  print('Usage: dot_csv_to_csv inputfilename outputfilename')
  sys.exit()

inputFileName = sys.argv[1]
outputFileName = sys.argv[2]

rows = []
skipped = 1
total = 0

with open(inputFileName, 'r') as inputFile:
    csvReader = csv.reader(inputFile)
    rowIndex = 0
    for row in csvReader:
      rowIndex = rowIndex + 1
      if (rowIndex < 3  or len(row) < 23 or row[9] is None):
        continue
      endCountyIndex = row[9].find(' DIVISION')
      county = 'no-county'
      if (endCountyIndex > 0):
        county = row[9][:endCountyIndex].strip()
      if (county not in includedCounties):
        continue
      total = total + 1

      rec = {
        'vendor_number': None,
        'vendor_name': None,
        'commodity_code': None,
        'commodity_code_description': None,
        'coa_certified': 'FALSE',
        'bbe': 'FALSE',
        'wbe': 'FALSE',
        'hbe': 'FALSE',
        'aibe': 'FALSE',
        'aabe': 'FALSE',
        'sbe': 'FALSE',
        'dobe': 'FALSE',
        'dbe': 'FALSE',
        'hub': 'FALSE',
        'ncdot': 'TRUE',
        'vendor_city': None,
        'vendor_state': None,
        'vendor_zip': None,
        'contact_name': None,
        'contact_phone': None,
        'contact_email': None,
        'email_addresses': None,
        'notes': None,
        'vendor_county': None
      }

      rec['vendor_name'] = row[0].strip()
      rec['contact_name'] = row[10].strip()
      rec['contact_phone'] = row[11].strip()
      rec['contact_email'] = rec['email_addresses'] = row[13].strip()
      if (row[6] is None or len(row[6].strip()) <= 0):
        rec['vendor_city'] = row[2].strip()
      else:
        rec['vendor_city'] = row[6].strip()

      if (row[7] is None or len(row[7].strip()) <= 0):
        rec['vendor_state'] = row[3].strip()
      else:
        rec['vendor_state'] = row[7].strip()

      if (row[8] is None or len(row[8].strip()) <= 0):
        rec['vendor_zip'] = row[4].strip()
      else:
        rec['vendor_zip'] = row[8].strip()
      
      rec['vendor_county'] = county
      rec['commodity_code_description'] = row[21].strip()
      rec['notes'] = 'Certifications: ' + row[16].strip()
      rows.append(rec)

print('Total rows output ' + str(len(rows)))

with open(outputFileName, 'w', newline='') as data_file:
 
  # create the csv writer object
  csv_writer = csv.writer(data_file)
  
  # Counter variable used for writing
  # headers to the CSV file
  count = 0
  for emp in rows:
      if count == 0:
  
          # Writing headers of CSV file
          header = emp.keys()
          csv_writer.writerow(header)
          count += 1
  
      # Writing data of CSV file
      csv_writer.writerow(emp.values())
 

			

