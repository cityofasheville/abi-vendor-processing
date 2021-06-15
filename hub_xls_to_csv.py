import xlrd
import csv
import sys
import os

# vendor_number	vendor_name	commodity_code	commodity_code_description	coa_certified	bbe	wbe	dbe	hbe	aibe	dobe	aabe	sbe	hub	ncdot	vendor_city	vendor_state	vendor_zip	contact_name	contact_phone	contact_email	email_addresses				

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

with open(outputFileName, 'w', newline='') as csvfile:
    ww = csv.writer(csvfile)
    for rx in range(sh.nrows):
      cols = []
      for i in range(sh.ncols):
        cols.append(sh.cell_value(rowx=rx, colx=i))
      ww.writerow(cols)

# Verify that the number of rows matches
rowCount = 0
with open(outputFileName, newline='') as csvfile:
  rr = csv.reader(csvfile)
  for row in rr:
    rowCount = rowCount + 1

if rowCount != sh.nrows:
  fmt = 'Output row count ({0}) does not match input row count ({1})'
  print(fmt.format(rowCount, sh.nrows))
