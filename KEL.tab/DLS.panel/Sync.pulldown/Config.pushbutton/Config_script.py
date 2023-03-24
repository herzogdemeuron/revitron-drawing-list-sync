import revitron
from drawinglistsync import Config, CONFIG_KEY
from revitron.ui import TabWindow, TextBox, CheckBox

config = Config()
window = TabWindow(
    'Drawing List Sync Configuration', ['Sheets', 'Revisions'], width=550, height=650
)

CheckBox.create(
    window,
    'Sheets',
    'createMissingSheets',
    config.createMissingSheets,
    title='Create Missing Sheets'
)

CheckBox.create(
    window,
    'Sheets',
    'createMissingParameters',
    config.createMissingParameters,
    title='Create Missing Parameters'
)

TextBox.create(window, 'Sheets', 'xlsFile', config.xlsFile, title='Excel File Path')

TextBox.create(
    window, 'Sheets', 'worksheet', str(config.worksheet), title='Worksheet Name'
)

TextBox.create(
    window,
    'Sheets',
    'parameterRow',
    str(config.parameterRow),
    title='Parameter Name Row Number'
)

TextBox.create(
    window,
    'Sheets',
    'sheetIdParameter',
    str(config.sheetIdParameter),
    title='Sheet Number Parameter'
)

CheckBox.create(
    window,
    'Revisions',
    'createRevisionList',
    config.createRevisionList,
    title='Enable Parsing Revisions'
)

TextBox.create(
    window,
    'Revisions',
    'revisionsRow',
    str(config.revisionsRow),
    title='Revisions Row Number'
)

TextBox.create(
    window,
    'Revisions',
    'revisoinPrefix',
    str(config.revisoinPrefix),
    title='Prefix of Revision Parameters'
)

TextBox.create(
    window,
    'Revisions',
    'maxCharsIndex',
    str(config.maxCharsIndex),
    title='Maximum Index Length'
)

TextBox.create(
    window,
    'Revisions',
    'maxCharsDate',
    str(config.maxCharsDate),
    title='Maximum Date Length'
)

TextBox.create(
    window,
    'Revisions',
    'maxCharsTitle',
    str(config.maxCharsTitle),
    title='Maximum Title Length'
)

TextBox.create(
    window, 'Revisions', 'dateFormat', str(config.dateFormat), title='Date Format'
)

window.show()

if window.ok:
	revitron.DocumentConfigStorage().set(CONFIG_KEY, window.values)
