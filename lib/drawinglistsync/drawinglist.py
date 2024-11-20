import csv
import ctypes
from ctypes import wintypes
import re
from drawinglistsync.date import DATE_REGEX, normalizeDateString
from drawinglistsync.collections import DrawingList
from os import system, environ, getenv
from os.path import dirname, join, isfile
from tempfile import mkdtemp
import clr
clr.AddReference('RevitAPIUI') 
from Autodesk.Revit.UI import TaskDialog

PARAM_MAX_COLS = 1000
# Define the necessary types
LPWSTR = ctypes.c_wchar_p
LPVOID = ctypes.c_void_p
LPBOOL = ctypes.POINTER(ctypes.c_int)
DWORD = ctypes.c_uint32

# Define the CopyFileEx function from the Windows API
copy_file_ex = ctypes.windll.kernel32.CopyFileExW
copy_file_ex.argtypes = [LPWSTR, LPWSTR, LPVOID, LPVOID, LPBOOL, DWORD]
copy_file_ex.restype = ctypes.c_int

def copy_file(source, dest):
	"""Copy a file from source to dest using the Windows CopyFileEx function."""
	user_profile_path = getenv('USERPROFILE')
	source = source.replace('%USERPROFILE%', user_profile_path)
	if not isfile(source):
		error_msg = TaskDialog.Show("HdM DT - Error","Source file can not be found at path.")
		return None
	copy_file_ex(source, dest, None, None, None, 0)
	if not isfile(dest):
		error_msg = TaskDialog.Show("HdM DT - Error","Could not copy source file.")
		return None
	return True

def createCsvFile(xls, worksheet):
	tmp = mkdtemp(prefix='drawinglistsync')
	copy = join(tmp, 'sheets.xls')
	csv = join(tmp, 'sheets.csv')
	convert = join(dirname(__file__), 'convert.bat')
	copy_process = copy_file(xls, copy)
	if copy_process == None:
		return None
	system('{} "{}" "{}" "{}"'.format(convert, copy, worksheet, csv))
	return csv

def getParameterCols(rows, parameterRow):
	row = rows[parameterRow - 1]
	return [(value, row[value]) for value in row if row[value]]

def getDrawinglistFromCsv(file, parameterRow, sheetIdParameter, dateFormat):
    if not isfile(file):
        error_msg = TaskDialog.Show("HdM DT - Error", "Drawing-List file can not be converted to csv.")
        return None, None
        
    drawingList = DrawingList()
    rows = []
    with open(file) as f:
        reader = csv.DictReader(f, range(1, PARAM_MAX_COLS))
        for row in reader:
            rows.append(row)
            
    parameterCols = getParameterCols(rows, parameterRow)
    try:
        sheetNumberCol = [item[0] for item in parameterCols if item[1] == sheetIdParameter][0]
    except:
        TaskDialog.Show("HdM DT - Warning", "Please check the Row Number configuration and remove empty rows from the top of your Drawing-List Excel")
        return None, None
        
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