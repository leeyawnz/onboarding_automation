# Importing statements
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# Input statements
wb_name = input('Insert excelsheet file name (.xlsx): ')
appname = input('Insert app team name: ')
persg = input('Insert population: ')
werks = input('Insert service (with " "): ')

# Loading workbook
wb = load_workbook(wb_name)
ws = wb.active

# Functions
def get_length(a, b):
    for col in range(a, b):
        char = get_column_letter(col)
        i = 1
        while True:
            cell = ws[char + str(i)].value
            if cell == None:
                return i
            else:
                i += 1
    return i

def get_array(a, b, end_of_row):
    array = []
    for col in range(a, b):
        char = get_column_letter(col)
        for row in range(1, end_of_row):
            cell = ws[char + str(row)].value
            array.append(cell)
    return array

# MDS YAML script
fields_length = get_length(2, 3)
fields_array = get_array(2, 3, fields_length)
table_array = get_array(1, 2, fields_length)

full_split_fields_array = []
ind_fields_array = []
i = 0
while i < fields_length - 1:
    if i == 0:
        ind_fields_array.append(fields_array[i].lower())
        i += 1
    elif table_array[i] == None and table_array[i - 1] != None:
        ind_fields_array.append(fields_array[i].lower())
        i += 1
    elif table_array[i] == None and table_array[i - 1] == None:
        ind_fields_array.append(fields_array[i].lower())
        i += 1
    else:
        full_split_fields_array.append(ind_fields_array)
        ind_fields_array = []
        ind_fields_array.append(fields_array[i].lower())
        i += 1
full_split_fields_array.append(ind_fields_array)

table_title_array = list(filter(lambda item: item is not None, table_array))
table_title_array = list(map(lambda item: item.lower(), table_title_array))

yaml_header = f'''---
role: [{appname}]
filter: {{
    it0001_persg: [{persg}],
    it0001_werks: [{werks}]
}}

table:'''

j = 0
while j < len(table_title_array):
    full_split_fields_array[j] = str(full_split_fields_array[j]).replace("'", "")
    yaml_entry = f'''
-   tablename: {table_title_array[j]}_main
    columns: {full_split_fields_array[j]}
    limit: null
    allow_aggregations: false
'''
    yaml_header += yaml_entry
    j += 1

print(yaml_header)
with open(f'{appname}.txt', 'w') as file:
    file.write(yaml_header)

txt_file = os.path.join('.', f'{appname}.txt')
yaml_file = txt_file.replace('.txt', '.yaml')
os.rename(txt_file, yaml_file)
