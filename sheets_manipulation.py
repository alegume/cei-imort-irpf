import csv
from os.path import getsize, join, isfile, dirname, abspath
from os import remove, listdir
import xlrd
import datetime

# some consts for easy configuring
FILE_SELLS = 'vendas.csv'
FILE_PM = 'preco-medio-acoes.csv'
NEGOTIATIONS_DIR = 'negotiations'
MSG_TO_MANY_SELLS = 'ATENÇÃO! Mais vendas do que o possível. É provável que aconteceu algum split, transferência de ativos ou outro evento que tenha aumentado sua quantidade de ações; porém, não entrou na planilha de negociações da CEI e não foi contabilizada. VERIFIQUE!'
BASE_DIR = dirname(abspath(__file__))


def remove_old_files_endswith(dir, ends):
    files = [f for f in listdir(dir) if isfile(join(dir, f)) and f.endswith(ends)]
    for file in files:
        remove(join(dir, file))

def search(sheet, str):
    for row in range(sheet.nrows):
        for col in range(sheet.ncols):
            if sheet.cell_value(row, col) == str:
                return (row, col)

def read_table(sheet, header, row, col, datemode):
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
        line[1] = datetime.datetime(*xlrd.xldate_as_tuple(line[1], datemode)).date()
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

def record_negotiations(negotiations):
    '''
        Record ALL negotiations in diferents csv files
    '''
    remove_old_files_endswith(NEGOTIATIONS_DIR, '.csv')
    total = {}
    for nego in negotiations:
        cod = nego['cod']
        total[cod] = total.get(cod, 0) + nego['qtd_compra']
        total[cod] = total.get(cod, 0) - nego['qtd_venda']
        file_name = join(NEGOTIATIONS_DIR, nego['cod'] + '.csv')
        obs = ''
        # If total stocks < 0 than something are missing
        obs = MSG_TO_MANY_SELLS if total[cod] < 0  else ''
        with open(file_name, mode='a+') as file:
            writer = csv.writer(file)
            if getsize(file_name) == 0:
                writer.writerow([
                    'COD', 'Data', 'Qtd compras', 'Qtd vendas', 'Posição',
                    'Qtd ações', 'Observações'
                ])
            writer.writerow([
                nego['cod'],
                nego['data'],
                nego['qtd_compra'],
                nego['qtd_venda'],
                nego['posicao'],
                total[cod],
                obs
            ])

def record_pms(pms):
    '''
        Record PMs and Sells in a csv file
    '''
    with open(FILE_PM, mode='w') as file:
        writer = csv.writer(file)
        if getsize(FILE_PM) == 0:
            writer.writerow(['Cod', 'Qtd vendas', 'Qtd compras', 'PM', 'Observações'])
        for cod, dict in pms.items():
            obs = MSG_TO_MANY_SELLS if dict.get('n_sell', 0) > dict.get('n_buy', 0) else ''
            writer.writerow([
                cod,
                dict.get('n_sell', 0),
                dict.get('n_buy', 0),
                dict.get('pm', 0),
                obs
            ])

    # Record Sells
    with open(FILE_SELLS, mode='w') as file:
        writer = csv.writer(file)
        if getsize(FILE_SELLS) == 0:
            writer.writerow([
                'Cod', 'Data', 'Qtd vendas', 'Qtd compras',
                'PM', 'Lucro', 'Observações'
            ])
        for cod, dict in pms.items():
            # Discart only buys
            if (dict.get('n_sell', 0) <= 0):
                continue

            # If total stocks < 0 than something are missing
            obs = MSG_TO_MANY_SELLS if dict.get('n_sell', 0) > dict.get('n_buy', 0) else ''
            profit =
            writer.writerow([dict.get('n_sell', 0), 
                cod,
                dict.get('data', 0),
                dict.get('n_sell', 0),
                dict.get('n_buy', 0),
                dict.get('pm', 0),
                # if its total_price < 0 it's a proffit
                profit,
                obs
            ])