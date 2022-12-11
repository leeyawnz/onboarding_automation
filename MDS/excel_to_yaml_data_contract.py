import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter
import yaml

# input statements
workbook_name = input('Enter Excelsheet Name (.xlsx): ')
data_table_length = input('Enter length of data (Last cell + 1): ')
data_table_length = int(data_table_length)
appname = input('Enter Application Team Name: ')

# loading workbook
wb = load_workbook(workbook_name)
ws = wb.active

# Tablename related
# Full list of tablenames including None
full_tablename_list = []
for col in range(1, 2):
    char = get_column_letter(col)
    for row in range(2, data_table_length):
        cell = ws[char + str(row)].value
        full_tablename_list.append(cell)

# Tablenames List for YAML
tablename_list = []
for col in range(1, 2):
    char = get_column_letter(col)
    for row in range(2, data_table_length):
        if ws[char + str(row)].value != None:
            table_value = ws[char + str(row)].value
            table_value = table_value.lower()
            tablename_list.append(table_value)

# Splitting tables in individual arrays
sliced_full_tablename_list = []
individual_tablename_list = []
i = 0
while i < len(full_tablename_list):
    if i == 0:
        individual_tablename_list.append(full_tablename_list[i])
        i += 1
    elif full_tablename_list[i] == None and full_tablename_list[i-1] != None:
        individual_tablename_list.append(full_tablename_list[i])
        i += 1    
    elif full_tablename_list[i] == None and full_tablename_list[i-1] == None:
        individual_tablename_list.append(full_tablename_list[i])
        i += 1    
    else:
        sliced_full_tablename_list.append(individual_tablename_list)
        individual_tablename_list = []
        individual_tablename_list.append(full_tablename_list[i])
        i += 1
sliced_full_tablename_list.append(individual_tablename_list)

# Getting length of each individual array in sliced table name list array
table_name_number_list = []
j = 0
while j < len(sliced_full_tablename_list):
    length = len(sliced_full_tablename_list[j])
    table_name_number_list.append(length)
    j += 1
# print(table_name_number_list)

# Fields related
# Full List for fields
full_field_list = []
for col in range(2, 3):
    char = get_column_letter(col)
    for row in range(2, data_table_length):
        cell = ws[char + str(row)].value
        cell = cell.lower()
        full_field_list.append(cell)

# Splitting fields array to match tablename array
k = 0
sliced_full_field_list = []
individual_field_list = []
for i in range(len(table_name_number_list)):
    k = 0
    while k < table_name_number_list[i]:
        individual_field_list.append(full_field_list[k])
        k += 1
    sliced_full_field_list.append(individual_field_list)
    individual_field_list = []
    k += table_name_number_list[i]

# Getting length of each individual array in sliced field list array
field_number_list = []
j = 0
while j < len(sliced_full_field_list):
    length = len(sliced_full_field_list[j])
    field_number_list.append(length)
    j += 1
# print(field_number_list)

# YAML formatting
with open(f'{appname}.txt', 'a') as file:
    yaml_header = '''
---
role: []
filter: {
    it0001_persg: [],
    it0001_werks: []
}

table:
    '''
    file.write(yaml_header)


l = 0
while l < len(field_number_list):
    sliced_full_field_list[l] = str(sliced_full_field_list[l]).replace("'", "")
    with open(f'{appname}.txt', 'a') as file:
        yaml_template = f'''
- tablename: {tablename_list[l]}
  columns: {sliced_full_field_list[l]}
  limit: null
  allow_aggregations: false
'''
        file.write(yaml_template)
        l += 1

old_file_name = os.path.join('.', f'{appname}.txt')
new_file_name = old_file_name.replace('.txt', '.yaml')
os.rename(old_file_name, new_file_name)
