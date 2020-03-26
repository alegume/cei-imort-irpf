#!/usr/bin/python3

'''
    Import xls files in import-file-xls and create the files:
        bens-direitos.csv,
'''

import xlrd

# some consts for easy configuring
DIR = 'import-file-xls'
NEGOTIATION_STR = 'INFORMAÇÕES DE NEGOCIAÇÃO DE ATIVOS'

def search(sheet, str):
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) == str:
                return (row, col)

def read_table(sheet, header, row, col):
    lines = []
    for r in range(row, sheet.nrows):
        # Abort if it the end of table
        if (sheet.cell(r, col).value == xlrd.empty_cell.value):
            break
        # Remove empty columns
        line = [x for x in sheet.row_values(r) if x is not '']
        cod = line[0].strip() # remove spaces
        # Remove F after the cod
        if cod.endswith('F'):
            line[0] = cod[0:-1]
        lines.append(line)
    # Create the dict
    negotiation = []
    for value in lines:
        negotiation.append(dict(zip(header, value)))

    return negotiation

def monthly_negotiations(sheet):
    # Find negotiation table
    row, col = search(sheet, NEGOTIATION_STR)
    # Jump empty rows and header
    row += 3
    # header =  [x for x in sheet.row_values(row) if x is not '']
    header = ['cod', 'data', 'qtd_compra', 'qtd_venda', 'pm_compra', 'pm_venda', 'qtd_liquida', 'posicao']
    negotiations = read_table(sheet, header, row, col)
    return negotiations


#### Begin
if __name__ == "__main__":
    workbook = xlrd.open_workbook('import-file-xls/jan.xls')
    sheet = workbook.sheet_by_index(0)

    print(monthly_negotiations(sheet))
