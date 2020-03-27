#!/usr/bin/python3

'''
    Import xls files in import-file-xls and create the files:
        bens-direitos.csv,
'''

# TODO: desconsiderar vendas, ver formula



from os import listdir
from os.path import isfile, join
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
        else:
            line[0] = cod
        # Converting
        line[0] = str(line[0])
        line[2] = float(str(line[2]).replace(',', '.'))
        line[3] = float(str(line[3]).replace(',', '.'))
        line[4] = float(str(line[4]).replace(',', '.'))
        line[5] = float(str(line[5]).replace(',', '.'))
        line[6] = float(str(line[6]).replace(',', '.'))
        line[7] = str(line[7])
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
    files = [f for f in listdir(DIR) if isfile(join(DIR, f)) and f.endswith('.xls')]
    negotiations = []
    for file in files:
        workbook = xlrd.open_workbook(join(DIR, file))
        sheet = workbook.sheet_by_index(0)
        negotiations += monthly_negotiations(sheet)

    print(negotiations)
    dict_pm_qtd = {}
    for nego in negotiations:
        if nego['posicao'].strip() == 'VENDIDA':
            continue
        else:
            cod = nego['cod']
            pm, qtd = dict_pm_qtd.get(cod, [0, 0])
            # Update pm and qtd
            pm += nego['pm_compra'] * nego['qtd_compra']
            qtd += nego['qtd_compra']
            dict_pm_qtd[cod] = [pm, qtd]

    for key, value in dict_pm_qtd.items():
        pm, qtd = value
        dict_pm_qtd[key][0] = pm / qtd

    print(dict_pm_qtd)