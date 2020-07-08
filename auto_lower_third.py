import subprocess
import sys
from docx import Document
import os


path = os.getcwd()
document = Document(f'{path}\\source_file.docx')
table = document.tables[0]
data = []
keys = {}
c = 0
for i, row in enumerate(table.rows):
    c += 1
    print(f'Делаем {c} титр')

    text = (cell.text for cell in row.cells)
    if i == 0:
        keys = tuple(text)
        continue
    row_data = dict(zip(keys, text))

    name = row_data['name'].strip()
    regalii = row_data['reg'].strip()


    with open(f"{path}\\data.txt", 'w') as file:
        i = 1
        file.write(f'var titr = ["{name}", "{regalii}"];' + '\n')


    with open("bat_runner.bat", 'w') as file:
        name_for_filename = name.replace(' ', '_')
        file.write(f'chcp 1251\n"C:\Program Files\Adobe\Adobe After Effects CC 2018\Support Files\\aerender.exe" -project {path}\lowerthird.aep -comp titr -OMtemplate title -output {path}\\render\{name_for_filename}.mov')

    program = f"{path}\\bat_runner.bat"
    process = subprocess.Popen(program)
    exit_code = process.wait()

    if exit_code == 0:
        print("Success!")
    else:
        print("Error!")