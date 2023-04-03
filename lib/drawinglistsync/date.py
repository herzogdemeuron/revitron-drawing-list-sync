import datetime

DATE_REGEX = '(\d{1,2}\.\d{1,2}\.\d{2,4}|\d{6}|\d{4}-\d{2}-\d{2}|\d{1,2}\/\d{1,2}\/\d{2,4}|\d{1,2}-\d{1,2}-\d{4})'
DATE_FORMATS = [
    r'%d.%m.%Y',  # 31.01.2023
	r'%d.%m.%y',  # 31.01.23
    r'%y%m%d',  # 230131
    r'%Y-%m-%d',  # 2023-01-31
    r'%m/%d/%Y',  # 01/31/2023
    r'%m/%d/%y',  # 01/31/23
    r'%m-%d-%Y',  # 01-31-2023
]


def getDateFromString(dateString):
	for format in DATE_FORMATS:
		try:
			date = datetime.datetime.strptime(dateString, format)
			if date:
				return date
		except:
			pass


def normalizeDateString(dateString, outputFormat):
	date = getDateFromString(dateString)
	if date:
		return date.strftime(outputFormat)
	return ''