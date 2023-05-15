import csv
import ctypes
from ctypes import wintypes
import re
from drawinglistsync.date import DATE_REGEX, normalizeDateString
from drawinglistsync.collections import DrawingList
from os import system, environ
from os.path import dirname, join
from tempfile import mkdtemp

PARAM_MAX_COLS = 1000
# Define the CopyFileEx function from the Windows API
copy_file_ex = ctypes.windll.kernel32.CopyFileExW
copy_file_ex.argtypes = [wintypes.LPCWSTR, wintypes.LPCWSTR, wintypes.LPVOID, wintypes.LPVOID, wintypes.LPBOOL, wintypes.DWORD]
copy_file_ex.restype = wintypes.BOOL

def copy_file(source, dest):
    """Copy a file from source to dest using the Windows CopyFileEx function."""
    if not copy_file_ex(source, dest, None, None, None, 0):
        raise ctypes.WinError()

def createCsvFile(xls, worksheet):
	tmp = mkdtemp(prefix='drawinglistsync')
	copy = join(tmp, 'sheets.xls')
	csv = join(tmp, 'sheets.csv')
	convert = join(dirname(__file__), 'convert.bat')
	if '%USERPROFILE%' in xls:
		user_profile = environ.get('USERPROFILE')
		xls = xls.replace('%USERPROFILE%',user_profile)
	copy_file(xls, copy)
	system('{} "{}" "{}" "{}"'.format(convert, copy, worksheet, csv))
	return csv

def getParameterCols(rows, parameterRow):
	row = rows[parameterRow - 1]
	return [(value, row[value]) for value in row if row[value]]


def getDrawinglistFromCsv(file, parameterRow, sheetIdParameter, dateFormat):
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
					value = row[col]
					match = re.match('^' + DATE_REGEX + '$', value)
					if match:
						value = normalizeDateString(value, dateFormat)
					data[name] = value
				except:
					pass
			drawingList.add(nr, data)
	return drawingList, sheetNumberCol