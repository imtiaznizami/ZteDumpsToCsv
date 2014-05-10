from xlrd import *
import csv
import glob
import codecs
import time
import os
import logging

# logging configuration
#logging.basicConfig(filename='log.txt',level=logging.DEBUG)
logging.basicConfig(level=logging.DEBUG)

start_time = time.time()
count = 0

logging.info("=" * 70)
logging.info("Starting script at: " + time.strftime("%Y-%m-%d %H:%M:%S"))

for fn in glob.iglob('dumps/*.xls*'):
	count = count + 1
	logging.info("Opening file: " + fn)
	wb = open_workbook(fn)

	for ws in wb.sheets():
		if not (ws.name == "Index" or ws.name == "TemplateInfo"):
			csv_fn = 'dumps/' + ws.name + '.csv'
			logging.info("Working with sheet: " + ws.name)

			if os.path.isfile(csv_fn):
				if count == 1:
					logging.error("The file, %s, already exist. Please delete and re-run." % (csv_fn,))
					input("Press Enter to continue...")
					exit(1)
				logging.info("Opening existing csv file: " + ws.name)
				rbc = open(csv_fn, 'a', newline='')
			else:
				logging.info("Creating and opening new csv file: " + ws.name)
				rbc = open(csv_fn, 'w', newline='')

			bcw = csv.writer(rbc,csv.excel)

			for row in range(ws.nrows):
				this_row = []
				for col in range(ws.ncols):
					val = ws.cell_value(row, col)

					this_row.append(val)
				if(count == 1 and (this_row[0] == "Modification Indication" or
						this_row[0] == "A,D,M,P" or
						this_row[0] == "A:Add, D:Delete, M:Modify, P:Pass" or
						this_row[0] == "")):
					continue
				if(count != 1 and (this_row[0] == "Modification Indication" or
						this_row[0] == "A,D,M,P" or
						this_row[0] == "A:Add, D:Delete, M:Modify, P:Pass" or
						this_row[0] == "" or this_row[0] == "MODIND")):
					continue
				else:
					bcw.writerow(this_row)

			logging.info("Closing file: " + ws.name)
			rbc.close()

logging.info("Ending script at: " + time.strftime("%Y-%m-%d %H:%M:%S"))
logging.info("Elapsed time: " + repr(time.time() - start_time) + " seconds.")
input("Press Enter to continue...")
