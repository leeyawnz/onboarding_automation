import os
from openpyxl import Workbook, load_workbook
from openpyxl.utils import get_column_letter

# input statements
wb_name = input('Enter Excelsheet Name (.xlsx): ')
data_tb_len = input('Enter length of data (Last cell + 1): ')
data_tb_len = int(data_tb_len)
appname = input('Enter Application Team Name: ')

# loading workbook
wb = load_workbook(wb_name)
ws = wb.active

# full file type list
full_ft_list = []
for col in range(1, 2):
    char = get_column_letter(col)
    for row in range(1, data_tb_len):
        cell = ws[char + str(row)].value
        full_ft_list.append(cell)


with open(f'{appname}.txt', 'a') as file:
    json_file_header = '''
{
    "sft-authorizer-sit": [
'''
    file.write(json_file_header)

i = 0
while i < len(full_ft_list):
    if i != len(full_ft_list) - 1:
        with open(f'{appname}.txt', 'a') as file:
            json_template = f"""
        {{
            "PutRequest": {{
                "Item": {{
                    "pk": {{
                        "s": "{appname}#{full_ft_list[i]}"
                    }},
                    "pkstatus": {{
                        "s": true
                    }}
                }}
            }}
        }},
"""
            file.write(json_template)
            i += 1
    else:
        with open(f'{appname}.txt', 'a') as file:
            json_template = f"""
        {{
            "PutRequest": {{
                "Item": {{
                    "pk": {{
                        "s": "{appname}#{full_ft_list[i]}"
                    }},
                    "pkstatus": {{
                        "s": true
                    }}
                }}
            }}
        }}
"""
            file.write(json_template)
            i += 1


with open(f'{appname}.txt', 'a') as file:
    json_file_footer = '''
    ]
}
'''
    file.write(json_file_footer)

txt_file = os.path.join('.', f'{appname}.txt')
yaml_file = txt_file.replace('.txt', '.json')
os.rename(txt_file, yaml_file)
