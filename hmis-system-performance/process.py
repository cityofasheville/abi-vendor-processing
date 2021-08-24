import xlrd
import csv
import json
import sys
import os

def extractXlsData(sheet, skipRows, skipCols, rowCount):
  rows = []
  for rx in range(sheet.nrows):
    if rx < skipRows:
      continue
    if rx > skipRows + rowCount:
      break
    cols = []
    for i in range(sheet.ncols):
      if i < skipCols:
        continue
      cols.append(sheet.cell_value(rowx=rx, colx=i))
    rows.append(cols)
  return rows

def processM2(asset, prefix):
  skipRows = asset["skipRows"]
  skipCols = asset["skipCols"]
  rowCount = asset["rowCount"]
  sheetName = asset["sheetName"]
  path = prefix+asset['filename']
  book = xlrd.open_workbook(path, logfile=(open(os.devnull, 'w')))
  sh = book.sheet_by_name(sheetName)

  fmt = 'Reading worksheet {0} with {1} rows, {2} columns'
  print(fmt.format(sh.name, sh.nrows, sh.ncols))
  return extractXlsData(sh, skipRows, skipCols, rowCount)

inputs = None
with open('inputs.json', 'r') as inputsFile:
  inputs = json.load(inputsFile)
prefix = './' + inputs['prefix'] + '/'
for nm in inputs['files']:
  asset = inputs['files'][nm]
  outputFileName = prefix + nm + '.csv'
  data = None
  if asset['active']:
    print('Process ' + nm)
    if nm == 'm2':
      data = processM2(asset, prefix)
    print(data)

    with open(outputFileName, 'w', newline='') as csvfile:
      csv_writer = csv.writer(csvfile)
      for row in data:
        csv_writer.writerow(row)
  else:
    print('Skip ' + nm)

#print(json.dumps(inputs,indent=' '))




