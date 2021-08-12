import win32com.client as win32
import os
import re
from win32com.client import constants
import pandas as pd
from docx import Document
from openpyxl import Workbook
from openpyxl import load_workbook
import numpy as np


def convert_to_docx(path):
    word = win32.gencache.EnsureDispatch('Word.Application')
    document = word.Documents.Open(path)
    document.Activate()

    new_file = os.path.abspath(path)
    new_file = re.sub(r'\.\w+$', '.docx', new_file)

    word.ActiveDocument.SaveAs(
        new_file,
        FileFormat=constants.wdFormatXMLDocument
    )
    document.Close(False)
    print(path)
    os.remove(path)


def is_numeric(s):
    try:
        float(s)
        return True
    except ValueError:
        return False


def read_table(document, table=9):
    # Choose table in document
    table = document.tables[table-1]
    # Put data from table in a array
    data = [[cell.text for cell in row.cells] for row in table.rows]
    # Delete first line and select just the columns 2 and 5
    data = np.delete(data, 0, axis=0)
    data = np.array(data[:, [2, 5]])
    # Delete items 7 and 10 | flatten (axis=None) the array
    data = np.delete(data, [7, 10], axis=None)
    # Swap positions 5 and 6
    data[5], data[6] = data[6], data[5]
    for i in range(14):
        data[i] = data[i].replace(',', '.')
        if not is_numeric(data[i]):
            data[i] = '0'

    data = data.astype(float)
    print(data)
    return data


def write_excel(data):
    wb = load_workbook('Preços de  03 a 09  de agosto de 2021 (1).xlsm', keep_vba=True)
    ws = wb['Preços']
    for i in range(1, 15):
        if data[i-1] != 0:
            ws.cell(row=16, column=i+1).value = data[i-1]
    wb.save('Preços de  03 a 09  de agosto de 2021 (1).xlsm')


#convert_to_docx('C:\\Users\\Cliente\\PycharmProjects\\excelGenerator\\Informativoconjuntural 09-08-2021.doc')
doc = Document("Informativoconjuntural 09-08-2021.docx")
result = read_table(doc)
write_excel(result)
#result.to_excel('exemplo.xlsx', sheet_name='sheet_exemplo')
