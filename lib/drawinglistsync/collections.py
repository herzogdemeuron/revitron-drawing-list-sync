import re
from revitron import DB
from drawinglistsync.date import DATE_REGEX, normalizeDateString
from System.Collections.Generic import Dictionary
from Autodesk.Revit.UI import TaskDialog
import sys

class GenericCollection(object):

    valueType = dict

    def __init__(self):
        self._collection = Dictionary[str, self.valueType]()

    def add(self, key, data):
        if not key:
            TaskDialog.Show(
                "HdM DT - Error",
                "Parameter issue detected. Please check your sheet number parameter configuration and try again."
            )
            sys.exit()
            
        str_key = str(key)
        if self._collection.ContainsKey(str_key):
            TaskDialog.Show(
                "HdM DT - Error",
                "Duplicate sheet number '{}' found in drawing list. Please remove duplicates and try again.".format(str_key)
            )
            sys.exit()
            
        self._collection.Add(str_key, data)

    def get(self, key):
        success, value = self._collection.TryGetValue(str(key))
        return value

    def all(self):
        return self._collection.GetEnumerator()

class DrawingList(GenericCollection):
	valueType = dict


class ModelSheetCollection(GenericCollection):
	valueType = DB.ViewSheet


class Revision(object):

	index = ''
	date = ''
	title = ''
	author = ''
	format = None

	def __init__(self, index, text, format):
		matches = re.match('^' + DATE_REGEX + '(.*)$', text, re.MULTILINE)
		indexAuthorRegex = '^([\w-]+)\s*(?:\(([^)]+)\))?'
		matchesIndexAuthor = re.match(indexAuthorRegex, index)
		self.index = matchesIndexAuthor.group(1)
		self.author = matchesIndexAuthor.group(2)
		self.author = self.author if self.author else ''
		self.format = format
		try:
			date = normalizeDateString(matches.group(1), format.dateFormat)
			self.date = date
			title = matches.group(2).lstrip()
			if len(title) > format.maxCharsTitle:
				title = title[:format.maxCharsTitle] + ' ...'
			self.title = title
		except:
			pass

	def __str__(self):
		lineList = [self.index, self.date, self.title]
		if self.format.showAuthor:
			lineList.insert(2, self.author)
		return '\t'.join(lineList)


class Revisions(GenericCollection):

	valueType = Revision
	maxLines = None

	def __init__(self, maxLines):
		self.maxLines = maxLines
		super(Revisions, self).__init__()

	def getLines(self):
		text = ''
		for rev in self._collection:
			text += str(rev.Value) + '\r\n'
		lines = text.splitlines()[-self.maxLines:]
		for n in range(len(lines)):
			if lines[n].startswith(' '):
				del (lines[n])
			else:
				break
		return lines

	def add(self, revision):
		key = '{}_{}'.format(revision.index, revision.date)
		self._collection.Add(key, revision)


class RevisionsList(GenericCollection):

	valueType = Revisions