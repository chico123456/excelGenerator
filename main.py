import win32com.client as win32
import os
import re
from win32com.client import constants
import pandas as pd
from docx import Document
import openpyxl


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
    os.remove(path)


def read_table(document, table=9, nheader=1):
    table = document.tables[table-1]
    data = [[cell.text for cell in row.cells] for row in table.rows]
    df = pd.DataFrame(data)
   # print(df.iloc[0])
    #if nheader == 1:
        #df = df.iloc[0].drop(df.index[0]).reset_index(drop=True)
        #df = df.rename(columns=df.iloc[0].drop(df.index[0]).reset_index(drop=True))

    return df


convert_to_docx('C:\\Users\\Cliente\\PycharmProjects\\excelGenerator\\saopedrodosul.doc')
doc = Document("formigueiro.docx")
result = read_table(doc)
print(result)
#result.to_excel('exemplo.xlsx', sheet_name='sheet_exemplo')
