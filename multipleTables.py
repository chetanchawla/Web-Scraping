from tablepyxl.tablepyxl import get_Tables, write_rows
from premailer import Premailer

#Writes multiple tables to a single sheet
def document_to_one_sheet_xl(doc, filename):
    wb = document_to_one_sheet_workbook(doc)
    wb.save(filename)

def document_to_one_sheet_workbook(doc):
    #Initiating the workbook and the worksheet
    wb = tablepyxl.Workbook()
    sheet = wb.active

    inline_styles_doc = Premailer(doc, remove_classes=False).transform()
    
    tables = get_Tables(inline_styles_doc)
    
    #Appending tables one below the other without leaving extra rows.
    #To append column wise, use write_columns
    
    row = 1
    for table in tables:
        if table.head:
            row = write_rows(sheet, table.head, row)
        if table.body:
            row = write_rows(sheet, table.body, row)
    
    return wb