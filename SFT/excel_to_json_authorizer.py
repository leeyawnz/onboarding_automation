# import statements
import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# constants
wb_name = input('Insert excelsheet file name (.xlsx): ')
env = input('what SFT environment is this?: ')
appname = input('Insert app team name: ')

# loading workbook
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

def sft_template(appname, array, i):
    sft_entry = f'''
        {{
            "PutRequest": {{
                "Item": {{
                    "pk": {{
                        "s": "{appname}#{array[i]}"
                    }},
                    "pkstatus": {{
                        "s": true
                    }}
                }}
            }}
        }}'''
    return sft_entry

# SFT JSON script
sft_array_length = get_length(1, 2)
sft_array = get_array(1, 2, sft_array_length)
sft_body = sft_template(appname, sft_array, 1) + ','
sft_header = f'''{{
    "sft-authorizer-{env}": ['''
sft_footer = f'''
    ]
}}'''

i = 0
while i < sft_array_length - 1:
    if i != sft_array_length - 2:
        sft_header += sft_template(appname, sft_array, i) + ','
        i += 1
    else:
        sft_header += sft_template(appname, sft_array, i)
        i += 1
sft_header += sft_footer

with open(f'{appname}.txt', 'a') as file:
    file.write(sft_header)

txt_file = os.path.join('.', f'{appname}.txt')
json_file = txt_file.replace('.txt', '.json')
os.rename(txt_file, json_file)
