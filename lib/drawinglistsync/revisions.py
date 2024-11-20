import csv
import datetime
from drawinglistsync.collections import Revision, Revisions, RevisionsList
from drawinglistsync.date import getDateFromString

REVISIONS_MAX_COLS = 1000


def getRevisionCols(rows, revisionsRow):
	row = rows[revisionsRow - 1]
	return [(value, row[value]) for value in row if row[value]]


def getRevisionsFromCsv(file, revisionsRow, sheetNumberCol, format):
	revisionsList = RevisionsList()
	rows = []
	with open(file) as f:
		reader = csv.DictReader(f, range(1, REVISIONS_MAX_COLS))
		for row in reader:
			rows.append(row)
	revisionCols = getRevisionCols(rows, revisionsRow)
	for n in range(revisionsRow, len(rows)):
		row = rows[n]
		nr = row[sheetNumberCol]
		if not nr:
			continue
		sheetRevisions = Revisions(format.maxLines)
		for item in revisionCols:
			col = item[0]
			if not row[col]:
				continue
			try:
				rev = Revision(row[col], item[1], format)
				date = getDateFromString(rev.date)
				sheetRevisions.add(rev)
				if date.date() >= datetime.datetime.now().date():
					break
			except:
				pass
		revisionsList.add(nr, sheetRevisions)
	return revisionsList


class RevisionFormat(object):

    maxCharsIndex = None
    maxCharsDate = None
    maxCharsTitle = None
    maxLines = None
    dateFormat = None
    showAuthor = None
    tabSeparatorCount = None

    def __init__(self, config):
        self.maxCharsIndex = config.maxCharsIndex
        self.maxCharsDate = config.maxCharsDate
        self.maxCharsTitle = config.maxCharsTitle
        self.maxLines = config.maxRevisionLines
        self.dateFormat = config.dateFormat
        self.showAuthor = config.showAuthor
        self.tabSeparatorCount = config.tabSeparatorCount