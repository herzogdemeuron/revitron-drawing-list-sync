from revitron import _, DB, DOC

def createOrUpdateSheets(drawingList, revisionList, modelSheetCollection, config):
	for item in drawingList.all():
		number = item.Key
		data = item.Value
		sheet = modelSheetCollection.get(number)
		if not sheet and config.createMissingSheets:
			sheet = createSheet()
		if not sheet:
			continue
		if _(sheet).isNotOwned():
			for key, value in data.items():
				if _(sheet).getParameter(key).exists() or config.createMissingParameters:
					_(sheet).set(key, value)
			sheetRevisions = revisionList.get(number)
			if sheetRevisions:
				lines = sheetRevisions.getLines()
				for i in range(0, len(lines)):
					paramName = config.paramNames[i]
					_(sheet).set(paramName, lines[i])


def createSheet():
	invalid = DB.ElementId.InvalidElementId
	return DB.ViewSheet.Create(DOC, invalid)


def getParams(modelSheetCollection, config):
	"""Get all parameter names with the prefix from the first sheet in the collection"""
	params = []
	enum = modelSheetCollection.all()
	enum.MoveNext()
	viewSheet =  enum.Current.Value
	prefix = config.revisoinPrefix
	for p in viewSheet.Parameters:
		pName = p.Definition.Name
		if pName.startswith(prefix) and (pName != prefix):
			params.append(pName)
	params.sort()
	return params