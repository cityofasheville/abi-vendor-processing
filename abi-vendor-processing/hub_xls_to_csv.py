import xlrd
import csv
import sys
import os

# HUB DATA FILE INPUT COLUMNS
# 0  Company Name
# 1  Contact Name
# 2  Address
# 3  Address2
# 4  City, State
# 5  Zip
# 6  County
# 7  Phone
# 8  Toll Free
# 9  Fax
# 10 Email
# 11 Hub Certification (B, W, HA, AA, AI accepted; D, SE skipped)
# 12 NCSBE
# 13 Certify Date
# 14 Certify End Date
# 15 Commodity In Search Criteria
# 16 Construction Codes In Search Criteria
# 17 Construction License
# 18 Construction License/Limitation
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
#  notes
#  vendor_county

includedCounties = ['BUNCOMBE', 'MADISON', 'HENDERSON', 'HAYWOOD', 'JACKSON', 'TRANSYLVANIA', 'POLK', 'RUTHERFORD', 'MCDOWELL', 'YANCEY']

# The -s is to suppress annoying error messages from xlrd
argc = len(sys.argv)
if (argc < 3 or argc > 4):
  print('Usage: xls_to_csv inputfilename outputfilename [-s]')
  sys.exit()

inputFileName = sys.argv[1]
outputFileName = sys.argv[2]
suppressXlrdMessages = (len(sys.argv) == 4)

if suppressXlrdMessages:
  book = xlrd.open_workbook(inputFileName, logfile=(open(os.devnull, 'w')))
else:
  book = xlrd.open_workbook(inputFileName)

sh = book.sheet_by_index(0)

fmt = 'Reading worksheet {0} with {1} rows, {2} columns'
print(fmt.format(sh.name, sh.nrows, sh.ncols))

rows = []
skipped = 1
with open(outputFileName, 'w', newline='') as csvfile:
    ww = csv.writer(csvfile)
    for rx in range(sh.nrows):
      if rx == 0:
        continue
      cols = []
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
        'hub': 'TRUE',
        'ncdot': 'FALSE',
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
      for i in range(sh.ncols):
        cols.append(sh.cell_value(rowx=rx, colx=i))
      rec['vendor_name'] = cols[0].strip()
      rec['contact_name'] = cols[1].strip()
      rec['contact_phone'] = cols[7].strip()
      rec['contact_email'] = rec['email_addresses'] = cols[10].strip()
      citystate = cols[4].split()
      rec['vendor_city'] = citystate[0].strip()
      rec['vendor_state'] = citystate[1].strip()
      rec['vendor_zip'] = cols[5].strip()
      rec['vendor_county'] = cols[6].strip()
      rec['commodity_code_description'] = cols[15].strip()
      hc = cols[11]
      skip = False
      if rec['vendor_state'] != 'NC' or rec['vendor_county'] not in includedCounties:
        skip = True
        skipped = skipped + 1
      if hc == 'B':
        rec['bbe'] = 'TRUE'
      elif hc == 'W':
        rec['wbe'] = 'TRUE'
      elif hc == 'HA':
        rec['hbe'] = 'TRUE'
      elif hc == 'AA':
        rec['aabe'] = 'TRUE'
      elif hc == 'AI':
        rec['aibe'] = 'TRUE'
      elif hc == 'D' or hc == 'SE':
        skip = True
        skipped = skipped + 1
      else:
        print('Unknown certification ' + hc)
      if not skip:
        rows.append(rec)
#      ww.writerow(cols)
print('Total rows output ' + str(len(rows)))

# now we will open a file for writing
data_file = open(outputFileName, 'w')
 
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
 
data_file.close()

			

