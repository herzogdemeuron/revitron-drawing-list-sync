from revitron import DocumentConfigStorage

CONFIG_KEY = 'revitron.drawing-list-sync'


class Config:

	xlsFile = ''
	parameterRow = ''
	sheetIdParameter = ''
	createMissingParameters = False
	createMissingSheets = False
	createRevisionList = False
	revisionsRow = ''
	revisoinPrefix = ''
	worksheet = ''
	maxCharsIndex = None
	maxCharsDate = None
	maxCharsTitle = None
	maxRevisionLines = None
	dateFormat = None
	paramNames = None

	def __init__(self):
		config = DocumentConfigStorage().get(CONFIG_KEY, dict())
		self.xlsFile = config.get('xlsFile', '')
		self.parameterRow = int(config.get('parameterRow', '1'))
		self.sheetIdParameter = config.get('sheetIdParameter', 'Sheet Number')
		self.createMissingParameters = config.get('createMissingParameters', False)
		self.createMissingSheets = config.get('createMissingSheets', False)
		self.createRevisionList = config.get('createRevisionList', False)
		self.revisionsRow = int(config.get('revisionsRow', '2'))
		self.revisoinPrefix = config.get('revisoinPrefix', 'Revisions')
		self.worksheet = config.get('worksheet', 'Sheets')
		self.maxCharsIndex = int(config.get('maxCharsIndex', '4'))
		self.maxCharsDate = int(config.get('maxCharsDate', '12'))
		self.maxCharsTitle = int(config.get('maxCharsTitle', '36'))
		self.dateFormat = config.get('dateFormat', r'%d.%m.%Y')