import xlrd
import csv
import sys
import os

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
with open(outputFileName, 'w', newline='') as csvfile:
    ww = csv.writer(csvfile)
    for rx in range(sh.nrows):
      cols = []
      rec = {
        'vendor_number': None,
        'vendor_name': '',
        'commodity_code': '',
        'commodity_code_description': '',
        'bbe': 'FALSE',
        'wbe': 'FALSE',
        'dbe': 'FALSE',
        'hbe': 'FALSE',
        'aibe': 'FALSE',
        'dobe': 'FALSE',
        'aabe': 'FALSE',
        'sbe': 'FALSE',
        'hub': 'FALSE',
        'ncdot': 'FALSE',
        'vendor_city': '',
        'vendor_state': '',
        'vendor_zip': '',
        'contact_name': '',
        'contact_phone': '',
        'contact_email': '',
        'email_addresses': ''
      }
      for i in range(sh.ncols):
        cols.append(sh.cell_value(rowx=rx, colx=i))
      rec['vendor_name'] = cols[0]
      rec['contact_name'] = cols[1]
      rows.append(cols)
#      ww.writerow(cols)

# HUB:
# Company Name	Contact Name	Address	Address2	City, State	Zip	County	Phone	Toll Free	Fax	Email	Hub Certification	NCSBE	Certify Date	Certify End Date	Commodity In Search Criteria	Construction Codes In Search Criteria	Construction License	Construction License/Limitation

# vendor_number	vendor_name	commodity_code	commodity_code_description	coa_certified	bbe	wbe	dbe	hbe	aibe	dobe	aabe	sbe	hub	ncdot	vendor_city	vendor_state	vendor_zip	contact_name	contact_phone	contact_email	email_addresses				

print(rows[0])

# Verify that the number of rows matches
# rowCount = 0
# with open(outputFileName, newline='') as csvfile:
#   rr = csv.reader(csvfile)
#   for row in rr:
#     rowCount = rowCount + 1

# if rowCount != sh.nrows:
#   fmt = 'Output row count ({0}) does not match input row count ({1})'
#   print(fmt.format(rowCount, sh.nrows))
