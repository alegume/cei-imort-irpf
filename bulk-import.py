#!/usr/bin/python3

'''
    Import xls files in import-file-xls and create the files:
        bens-direitos.csv,
'''

# TODO: salvar bens e vendas
# DONE: desconsiderar vendas, ver formula

import sys
import csv
from os import listdir
from os.path import isfile, join, getsize
import xlrd

# some consts for easy configuring
DIR = 'import-file-xls'
NEGOTIATION_STR = 'INFORMAÇÕES DE NEGOCIAÇÃO DE ATIVOS'
FILE_SELLS = 'vendas.csv'
FILE_PM = 'preco-medio-acoes.csv'

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
        # Verify format
        if len(line) != 8:
            print('\t\t Format error in table')
            sys.exit(1)
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

def record_sells(cod, info):
    '''
        Record sells in a csv file
    '''
    with open(FILE_SELLS, mode='a+') as file:
        writer = csv.writer(file)
        if getsize(FILE_SELLS) == 0:
            writer.writerow(['COD', 'Vendas', 'Compras', 'Lucro', 'Total Acc'])
        writer.writerow([
            cod,
            info['n_sell'],
            info['n_buy'],
            info['profit'],
            info['total_price']
        ])
    # print(cod, info)

def record_pms(pms):
    '''
        Record PMs in a csv file
    '''
    with open(FILE_PM, mode='w') as file:
        writer = csv.writer(file)
        if getsize(FILE_PM) == 0:
            writer.writerow(['Cod', 'Qtd vendas', 'Qtd compras', 'PM', 'Total Acc'])
        for cod, dict in pms.items():
            writer.writerow([
                cod,
                dict.get('n_sell', 0),
                dict.get('n_buy', 0),
                dict.get('pm', 0),
                dict.get('total_price', 0)
            ])
    # print(cod, info)


def median_prices(negotiations):
    '''
        Calculation of median prices by the formula:
            pm = sum(qtd * pm_compra) - sum(qtd * pm_venda) / (sum(qtd_compra) - sum(qtd_venda))
    '''
    dicts_cod_sums = {}
    for nego in negotiations:
        cod = nego['cod']
        cod_sum = dicts_cod_sums.get(cod, {
            'total_price': 0,
            'n_buy': 0,
            'n_sell': 0
        })
        if nego['posicao'].strip() == 'VENDIDA':
            # Update total_price, n_sell
            cod_sum['total_price'] = cod_sum.get('total_price', 0) - nego['pm_venda'] * nego['qtd_venda']
            cod_sum['n_sell'] = cod_sum.get('n_sell', 0) + nego['qtd_venda']
            # print('\tVEnda', cod)
            # print(nego['pm_venda'], '*', nego['qtd_venda'], '=',cod_sum['total_price'])
        elif nego['posicao'].strip() == 'COMPRADA':
            # Update total_price, n_buy
            cod_sum['total_price'] = cod_sum.get('total_price', 0) + nego['pm_compra'] * nego['qtd_compra']
            cod_sum['n_buy'] = cod_sum.get('n_buy', 0) + nego['qtd_compra']
            # print('\tCompra', cod)
            # print(nego['pm_compra'], '*', nego['qtd_compra'], '=',cod_sum['total_price'])
        else:
            print('\t\t Error in column Posição')

        dicts_cod_sums[cod] = cod_sum
    for cod in dicts_cod_sums:
        # print(cod,
        #     dicts_cod_sums[cod]['total_price'],
        #     dicts_cod_sums[cod]['n_buy'],
        #     dicts_cod_sums[cod]['n_sell']
        # )
        try:
            dicts_cod_sums[cod]['pm'] = dicts_cod_sums[cod]['total_price'] / (dicts_cod_sums[cod].get('n_buy', 0) - dicts_cod_sums[cod].get('n_sell', 0))
        except ZeroDivisionError:
            # TODO: record sells
            # there are a profit if it's negative
            dicts_cod_sums[cod]['profit'] = - dicts_cod_sums[cod]['total_price']
            record_sells(cod, dicts_cod_sums[cod])
            print('\t\tStock', cod, 'fully selled!!!')

    return dicts_cod_sums

#### Begin
if __name__ == "__main__":
    files = [f for f in listdir(DIR) if isfile(join(DIR, f)) and f.endswith('.xls')]
    negotiations = []
    for file in files:
        workbook = xlrd.open_workbook(join(DIR, file))
        sheet = workbook.sheet_by_index(0)
        negotiations += monthly_negotiations(sheet)

    pms = median_prices(negotiations)
    record_pms(pms)
    # print(negotiations)
    # print(pms)