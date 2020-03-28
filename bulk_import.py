#!/usr/bin/python3

'''
    Import xls files in import-file-xls and create the files:
        preco-medio-acoes.csv, vendas.csv and one csv file per stock in negotiations.
'''

# Lembrar: Verificar a venda de TAEE em maio. Compra e vanda mesmo dia. Prejuizo.
# Calcular e colacar em vendas.csv
# Escrever a venda no ato da venda considerando o PM atual e registrar lucro/preju


# DONE: Order by date
# TODO: Readme.md
# DONE: VErificar se sell >= buy pra vendas, etc
# DONE: Criar planilha de negociacao por ativo, salvar bens e vendas, desconsiderar vendas, ver formula

import sys
from os import listdir
from os.path import isfile, join

from sheets_manipulation import *

# some consts for easy configuring
DIR = 'import-file-xls'
NEGOTIATION_STR = 'INFORMAÇÕES DE NEGOCIAÇÃO DE ATIVOS'


def monthly_negotiations(sheet, datemode):
    # Find negotiation table
    row, col = search(sheet, NEGOTIATION_STR)
    # Jump empty rows and header
    row += 3
    header = ['cod', 'data', 'qtd_compra', 'qtd_venda', 'pm_compra', 'pm_venda', 'qtd_liquida', 'posicao']
    negotiations = read_table(sheet, header, row, col, datemode)
    return negotiations


def median_prices(negotiations):
    '''
        Calculation of median prices by the formula:
            pm = sum(qtd * pm_compra) - sum(qtd * pm_venda) / (sum(qtd_compra) - sum(qtd_venda))
    '''
    dicts_cod_sums = {}
    # Populate the dict with infos and sums
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
            cod_sum['data'] = nego['data']
            # print('\tVEnda', cod)
            # print(nego['pm_venda'], '*', nego['qtd_venda'], '=',cod_sum['total_price'])
        elif nego['posicao'].strip() == 'COMPRADA':
            # Update total_price, n_buy
            cod_sum['total_price'] = cod_sum.get('total_price', 0) + \
                nego['pm_compra'] * nego['qtd_compra']
            cod_sum['n_buy'] = cod_sum.get('n_buy', 0) + nego['qtd_compra']
            # print('\tCompra', cod)
            # print(nego['pm_compra'], '*', nego['qtd_compra'], '=',cod_sum['total_price'])
        else:
            print('\t\t Error in column Posição. Unnable to proceed, verify import file.')
        dicts_cod_sums[cod] = cod_sum

    # Calculate PM
    for cod in dicts_cod_sums:
        try:
            dicts_cod_sums[cod]['pm'] = dicts_cod_sums[cod]['total_price'] \
                / (dicts_cod_sums[cod].get('n_buy', 0) - dicts_cod_sums[cod].get('n_sell', 0))
        except ZeroDivisionError:
            print('\t\tStock', cod, 'fully selled!!!')

    return dicts_cod_sums

#### Begin
if __name__ == "__main__":
    files = [f for f in listdir(DIR) if isfile(join(DIR, f)) and f.endswith('.xls')]
    negotiations = []
    for file in files:
        workbook = xlrd.open_workbook(join(DIR, file))
        sheet = workbook.sheet_by_index(0)
        negotiations += monthly_negotiations(sheet, workbook.datemode)

    # Sort negotiations by date
    negotiations = sorted(negotiations, key=lambda k: k['data'])
    # Create csv files in negotiations dir and create delss file
    record_negotiations(negotiations)
    pms = median_prices(negotiations)
    record_pms(pms)
    # print(pms)