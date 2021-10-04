import xlrd
import csv
import json
import sys
import os

def readXlsFile(asset, prefix, nm):
  skipRows = asset["skipRows"]
  skipCols = asset["skipCols"]
  rowCount = asset["rowCount"]
  sheetName = asset["sheetName"]
  skipBetweenRows = asset['skipBetweenRows']
  sections = asset['sections']
  path = prefix+asset['filename']
  book = xlrd.open_workbook(path, logfile=(open(os.devnull, 'w')))
  sheet = book.sheet_by_name(sheetName)

  fmt = '  Reading worksheet {0} with {1} rows, {2} columns'
  print(fmt.format(sheet.name, sheet.nrows, sheet.ncols))


  for secNumber in range(sections):
    rows = []
    for rx in range(sheet.nrows):
      sectionStart = skipRows + ((rowCount + skipBetweenRows) * (secNumber))
      if rx < sectionStart:
        continue
      if rx >= sectionStart + rowCount:
        break
      cols = []
      for i in range(sheet.ncols):
        if i < skipCols:
          continue
        cols.append(sheet.cell_value(rowx=rx, colx=i))
      rows.append(cols)
    if sections == 1:
      outputFileName = prefix + nm + '.csv'
    else:
      outputFileName = prefix + nm + "." + str(secNumber +1) + '.csv'
    print('\nHere is the data we read:\n')
    print(rows)
    with open(outputFileName, 'w', newline='') as csvfile:
      csv_writer = csv.writer(csvfile)
      for row in rows:
        csv_writer.writerow(row)
  return rows


# Main script
inputs = None
with open('inputs.json', 'r') as inputsFile:
  inputs = json.load(inputsFile)
prefix = './' + inputs['prefix'] + '/'
for nm in inputs['files']:
  asset = inputs['files'][nm]
  data = None
  if asset['active']:
    print('Process ' + nm)
    data = readXlsFile(asset, prefix, nm)
  else:
    print('\nSkip ' + nm)
