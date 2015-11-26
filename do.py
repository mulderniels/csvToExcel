#made by n.mulder1@uu.nl 2015

#converts csv to xlsx for all file in arbitrary deep folder structure
#converts float and int to numbers
#file extension is case sensitive (.CSV != .csv)

#inti
import os
import csv
import fnmatch
from xlsxwriter.workbook import Workbook


#settings
csvDelimiter = ','

#do stuff
print('hallo')

def isFloat(value):
	try:
		float(value)
		return True
	except:
		return False

def isInt(value):
	try:
		float(value)
		return True
	except:
		return False

for dirpath, dirs, files in os.walk('.'):
	for filename in fnmatch.filter(files, '*.CSV'):
		csvfile = os.path.join(dirpath, filename)
		print(csvfile)
		workbook = Workbook(csvfile + '.xlsx')
		worksheet = workbook.add_worksheet()
		with open(csvfile, 'rb') as f:
			reader = csv.reader(f, delimiter=csvDelimiter)
			for r, row in enumerate(reader):
				for c, cellValue in enumerate(row):
					if isFloat(cellValue):
						cellValue = float(cellValue)
					elif isInt(cellValue):
						cellValue = int(cellValue)
					worksheet.write(r, c, cellValue)
		workbook.close()

print('doei')
