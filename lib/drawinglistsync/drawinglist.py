import csv
import shutil
from drawinglistsync.collections import DrawingList
from os import system
from os.path import dirname, join
from tempfile import mkdtemp

PARAM_MAX_COLS = 1000


def createCsvFile(xls, worksheet):
	tmp = mkdtemp(prefix='drawinglistsync')
	copy = join(tmp, 'sheets.xls')
	csv = join(tmp, 'sheets.csv')
	convert = join(dirname(__file__), 'convert.bat')
	shutil.copyfile(xls, copy)
	system('{} "{}" "{}" "{}"'.format(convert, copy, worksheet, csv))
	return csv


def getParameterCols(rows, parameterRow):
	row = rows[parameterRow - 1]
	return [(value, row[value]) for value in row if row[value]]


def getDrawinglistFromCsv(file, parameterRow, sheetIdParameter):
	drawingList = DrawingList()
	rows = []
	with open(file) as f:
		reader = csv.DictReader(f, range(1, PARAM_MAX_COLS))
		for row in reader:
			rows.append(row)
	parameterCols = getParameterCols(rows, parameterRow)
	sheetNumberCol = [item[0] for item in parameterCols if item[1] == sheetIdParameter][0]
	for n in range(parameterRow, len(rows)):
		row = rows[n]
		nr = row[sheetNumberCol]
		if nr:
			data = {}
			for item in parameterCols:
				try:
					col = item[0]
					name = item[1]
					data[name] = row[col]
				except:
					pass
			drawingList.add(nr, data)
	return drawingList, sheetNumberCol