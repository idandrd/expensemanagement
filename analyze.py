# -*- coding: utf-8 -*-

import codecs
import gspread

# Find this value in the url with 'key=XXX' and copy XXX below
spreadsheet_key = '0AqDyqi_aCyu4dHI2VHhQTnpSSE5MXzV2eFRBQTRTbFE'

gc = gspread.login('limited@cohenyuval.com', 'aBcD1@3$')
expenses_ws = gc.open('Personal Expenses').sheet1
provider_ws = gc.open('Personal Expenses').worksheet('Providers')
cells = expenses_ws.get_all_values()
provider_cells = provider_ws.get_all_values()

providers = dict(zip([cell[0] for cell in provider_cells[1:]], range(1, len(provider_cells))))

column_translation = {u'שם בית עסק': 1, u'סכום לחיוב': 2, u'תאריך רכישה': 3}

f = open('/Users/yuval/Downloads/output.txt')

raw_data = unicode(f.read().decode('1255'))

data = [row.split('\t') for row in raw_data.split('\n')]

expenses = [dict(zip(data[2], row)) for row in data[4:-2]]

next_row = len(cells) + 1

for expense in expenses:
	for v, k in expense.iteritems():
		if column_translation.has_key(v):
			expenses_ws.update_cell(next_row, column_translation[v.strip()], k)
        		if (column_translation[v.strip()] == 1) and not providers.has_key(k):
                		provider_ws.update_cell(len(providers)+2, 1, k)
                		providers[k] = len(providers) + 2

	next_row = next_row + 1

