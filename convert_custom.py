import xlrd
import csv
import os
from openpyxl import load_workbook


def exceltocsv():
	x= os.listdir("Specify Directory Here")
	print x
	for filenames in x:
		m = filenames.rstrip()
		z = "Specify Directory Here" + m
		wb = load_workbook(z)
		wb = xlrd.open_workbook(z)
		sh = wb.sheet_by_name('Sheet1')
		your_csv_file = open(m +'.csv', 'wb')
		wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)
    
    #for deleting first two rows
		i = 3
		for rownum in range(sh.nrows):
			if i > 1:
				
				i = i-1
			else:
				wr.writerow(sh.row_values(rownum))

		your_csv_file.close()
		sh = wb.sheet_by_name('Sheet1')
		your_csv_file = open('ALL_Combined' +'.csv', 'ab')
		wr = csv.writer(your_csv_file, quoting=csv.QUOTE_ALL)

		i = 3
		for rownum in range(sh.nrows):
			if i > 1:
				
				i = i-1
			else:
				wr.writerow(sh.row_values(rownum))

		your_csv_file.close()

exceltocsv()
