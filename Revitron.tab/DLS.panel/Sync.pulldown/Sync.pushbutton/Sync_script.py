from revitron import Filter, Transaction, _
from drawinglistsync import (
    Config,
    createCsvFile,
    getDrawinglistFromCsv,
    getRevisionsFromCsv,
    createOrUpdateSheets,
    ModelSheetCollection,
    RevisionsList,
    RevisionFormat,
    getParams
)
from pyrevit import forms

config = Config()

if not config.xlsFile or not config.parameterRow or not config.revisionsRow:
	forms.alert('Please first configure the synching!', exitscript=True)

modelSheetCollection = ModelSheetCollection()
for item in Filter().byCategory('Sheets').noTypes().getElements():
	modelSheetCollection.add(_(item).get(config.sheetIdParameter), item)

paramNames = getParams(modelSheetCollection, config)
if len(paramNames) < 1:
	forms.alert('There is no Parameters match the given prefix.', exitscript=True)
	
config.paramNames = paramNames
config.maxRevisionLines = len(paramNames)

csvSheets = createCsvFile(config.xlsFile, config.worksheet)
drawingList, sheetNumberCol = getDrawinglistFromCsv(csvSheets, config.parameterRow, config.sheetIdParameter, config.dateFormat)
revisionList = RevisionsList()

if (config.createRevisionList):
	revisionList = getRevisionsFromCsv(
	    csvSheets, config.revisionsRow, sheetNumberCol, RevisionFormat(config)
	)

with Transaction():
	createOrUpdateSheets(drawingList, revisionList, modelSheetCollection, config)