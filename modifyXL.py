import openpyxl as xl
import logging

DebugMessages = False


def print_debug(message):
    if DebugMessages is True:
        print(message)


# - %(levelname)s
logging.basicConfig(filename='app.log', filemode='w', format='%(asctime)s : %(message)s\n',
                    datefmt='%d.%m.%Y, %H:%M:%S',
                    level=logging.DEBUG)

start = ('*' * 20 + ' Start ' + '*' * 20)
end = ('*' * 20 + ' End ' + '*' * 20)

logging.info(start)


def antet_function(name):
    return '*' * 10 + ' ' + name + ' ' + '*' * 10


def connectXL(file_name, sheet_name):
    """
    This function is used to connect to excel document
    :param file_name: name of .xlsx file
    :param sheet_name: sheet name or sheet number
    :return: wb and wb_sh
    """

    logging.info(antet_function('connectXL({}, {})'.format(file_name, sheet_name)))

    logging.info('Connect to file {}'.format(file_name))
    # wb = xl.load_workbook(file_name, read_only=False, keep_links=True, keep_vba=True) # for .xlsm files
    wb = xl.load_workbook(file_name, read_only=False) # for non .xlsm files

    logging.info('Open sheet {} from file {}'.format(sheet_name, file_name))
    if isinstance(sheet_name, int):  # sheet_name is a nomber // 0 - first sheet
        wb_sh = wb.worksheets[sheet_name]
    elif isinstance(sheet_name, str):  # sheet_name is a string // name of sheet
        wb_sh = wb[sheet_name]
    return wb, wb_sh


def modifyXL(wb_sh, cells, values):
    """
    This function is used to write [values] in [cells] from [file_name]
    :param wb_sh: sheet name or range
    :param cells: list of cells // A1, A2, B15
    :param values: new values for cells // 7, 12, 90 => A1 = 7, A2 = 12, B15 = 90
    :return:
    """
    logging.info(antet_function('modifyXL({}, {}, {})'.format(wb_sh, cells, values)))

    if len(cells) == len(values):
        for i in range(0, len(cells)):
            print_debug('cell {} contains value {}'.format(cells[i], values[i]))
            logging.info('start modifying value of {} to {}'.format(cells[i], values[i]))
            wb_sh[cells[i]].value = values[i]
            logging.info('value of {} was modified to {}'.format(cells[i], values[i]))
    elif len(cells) > len(values):
        logging.info('len(cells) > len(values)')
        print_debug('lungimea cells > lungime values')
    elif len(values) > len(cells):
        logging.info('lungimea values > lungime cells')
        print_debug('lungimea values > lungime cells')


# cells to be modified
cells = [
    'A1',
    'B2',
    'C3',
    'D4',
    'E5'
]

# new values
values = [
    'A1_value',
    'B2_value',
    'C3_value',
    'D4_value',
    'E5_value'
]

file_name = 'test_1.xlsx'

# connect to file
wb, wb_sh = connectXL(file_name=file_name, sheet_name='Sheet1')

# modifyXL cell values
modifyXL(wb_sh=wb_sh, cells=cells, values=values)

# save modified file
wb.save('file_2.xlsx')

# end of story
logging.info(end)
