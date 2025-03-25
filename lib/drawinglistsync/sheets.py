from revitron import _, DB, DOC
from drawinglistsync.config import Config
from pyrevit import forms

# Add new variable to store selected titleblock for all sheets
_selected_titleblock = None

def createSheet():
	global _selected_titleblock
	invalid = DB.ElementId.InvalidElementId
	
	# Get all titleblock types
	collector = DB.FilteredElementCollector(DOC).OfCategory(DB.BuiltInCategory.OST_TitleBlocks)
	titleblock_types = collector.WhereElementIsElementType().ToElements()
	
	if titleblock_types.Count > 0:
		# If we already have a selected titleblock, use it
		if _selected_titleblock:
			if _selected_titleblock == "None":
				return DB.ViewSheet.Create(DOC, invalid)
			sheet = DB.ViewSheet.Create(DOC, _selected_titleblock.Id)
			if sheet:
				return sheet
				
		# Create dictionary of titleblock names for selection
		titleblock_dict = {"{}: {}".format(t.FamilyName, t.get_Parameter(DB.BuiltInParameter.SYMBOL_NAME_PARAM).AsString()): t 
						  for t in titleblock_types}
		# Add None option
		titleblock_dict["None"] = "None"
		
		# Create message
		message = 'Selected Titleblock will be applied to all created sheets'
		
		# Prompt user to select titleblock type
		selected = forms.CommandSwitchWindow.show(
			sorted(titleblock_dict.keys()),
			message=message
		)
		
		if selected:
			# Store selection
			_selected_titleblock = titleblock_dict[selected]
			
			# Create new sheet with selected titleblock
			if _selected_titleblock == "None":
				return DB.ViewSheet.Create(DOC, invalid)
			sheet = DB.ViewSheet.Create(DOC, _selected_titleblock.Id)
			if sheet:
				return sheet
	
	# If no titleblocks available or selection cancelled, create sheet without titleblock
	return DB.ViewSheet.Create(DOC, invalid)


def createOrUpdateSheets(drawingList, revisionList, modelSheetCollection, config):
	global _selected_titleblock
	_selected_titleblock = None  # Reset the selection at start
	
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
				for i in range(config.maxRevisionLines):
					if i <len(lines):
						paramName = config.paramNames[i]
						_(sheet).set(paramName, lines[i])
					else:
						_(sheet).set(config.paramNames[i], '')


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