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

csv_headers = data[2]
csv_body = data[4:-2]

expenses = [dict(zip(data[2], row)) for row in data[4:-2]]

next_row = len(cells) + 1

cells_for_dups = [row[:3] for row in cells]

for expense in expenses:
	new_expense = {}
	for k, v in expense.iteritems():
		if column_translation.has_key(k):
			new_expense[column_translation[k]] = v
	if not [new_expense[1],new_expense[2],new_expense[3]] in cells_for_dups:
		for k,v in new_expense.iteritems:
			expenses_ws.update_cell(next_row, k, v)
        		if (k == 1) and not providers.has_key(v):
                		provider_ws.update_cell(len(providers)+2, 1, v)
                		providers[v] = len(providers) + 2

	next_row = next_row + 1
