#!/usr/bin/python3

'''
    Import xls files in import-file-xls and create the files:
        preco-medio-acoes.csv, vendas.csv and one csv file per stock in negotiations.
'''

# Lembrar: Verificar a venda de TAEE em maio. Compra e vanda mesmo dia. Prejuizo.
# Calcular e colacar em vendas.csv
# Explicar README: Escrever a venda no ato da venda considerando o PM atual e registrar lucro/preju


# DONE: Order by date
# TODO: Considerar custos operacionais
# TODO: Readme.md
# DONE: VErificar se sell >= buy pra vendas, etc
# DONE: Criar planilha de negociacao por ativo, salvar bens e vendas, desconsiderar vendas, ver formula

from os import listdir
from os.path import isfile, join

from sheets_manipulation import *

# some consts for easy configuring
COST_OF_OPERATION = 3.10
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
        Calculation of median prices and record sells
    '''
    dicts_cod_sums = {}
    sells = []
    # Populate the dict with infos and sums
    for nego in negotiations:
        cod = nego['cod']
        cod_sum = dicts_cod_sums.get(cod, {
            'total_price': 0,
            'n_buy': 0,
            'n_sell': 0,
            'total_stocks': 0,
            'pm': 0
        })

        # Add the cost of operation on each negotiation
        cod_sum['total_price'] += COST_OF_OPERATION
        cod_sum['total_price'] += nego['pm_compra'] * nego['qtd_compra']
        cod_sum['total_price'] -= nego['pm_venda'] * nego['qtd_venda']

        cod_sum['n_sell'] += nego['qtd_venda']
        cod_sum['n_buy'] += nego['qtd_compra']

        cod_sum['total_stocks'] += nego['qtd_compra'] - nego['qtd_venda']
        cod_sum['data'] = nego['data']

        if nego['posicao'].strip() == 'VENDIDA':
            # Needs to update total price, cause total_stocks changed but not PM!!!
            cod_sum['total_price'] = cod_sum['pm'] * cod_sum['total_stocks']
            # profit Calculation
            cod_sum['profit'] = (nego['pm_venda'] - cod_sum['pm']) * nego['qtd_venda']
            # Append cod to cod_sum for easy handling
            cod_sum['cod'] = cod
            # We should prevent a pointer like behavior
            sells.append(cod_sum.copy())
        elif nego['posicao'].strip() == 'COMPRADA':
            # Pm is only calculated in buying
            # Calculate PM
            try:
                cod_sum['pm'] = cod_sum['total_price'] / cod_sum['total_stocks']
            except ZeroDivisionError:
                print('\t\tStock', cod, 'fully selled!!!')
        else:
            print('\t\t Error in column Posição. Unnable to proceed, verify import file.')

        print(nego)
        print(
            cod,
            cod_sum['n_buy'],
            cod_sum['n_sell'],
            cod_sum['total_stocks'],
            cod_sum['total_price'],
            cod_sum['pm'],
            cod_sum.get('profit', 0),
        )

        # Update to dict of dicts
        dicts_cod_sums[cod] = cod_sum

    # Record sells before return
    record_sells(sells)
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

    # print(negotiations)

    # Create csv files in negotiations dir and create selss file
    record_negotiations(negotiations)
    pms = median_prices(negotiations)
    record_pms(pms)
    # print(pms)