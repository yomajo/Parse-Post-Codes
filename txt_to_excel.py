import os
import sys
import logging
from textwrap import wrap
import openpyxl

#LOGGING SETUP
log_path = os.path.join(os.path.dirname(__file__), 'txt_to_excel.log') 
logging.basicConfig(level='DEBUG', filename=log_path, filemode='w')

# GLOBAL VARIABLES
txt_file = 'post_codes.txt' #Default value
split_at_chars = 13
xlsx_file = 'Postal_Codes_Manager.xlsx'
sheet_names = ['Free Codes', 'Expired Codes']

# PATH BUILDING
wb_path = os.path.join(os.path.dirname(__file__), xlsx_file)
# If sys argument with path has been provided - use it as source abs path of text file, if not, default to hardcoded name.
if len(sys.argv) > 1:
    txt_path = sys.argv[1]
    logging.debug(f'This path was given through system arguments: {txt_path}')
else:
    txt_path = os.path.join(os.path.dirname(__file__), txt_file)
    logging.debug(f'No system arguments were passed; defaulting to hardcoded path: {txt_path}')


def convert_txt_to_list(txt_path):
    '''takes argument of txt file with a block of codes without spaces like 'UA982420356LTUA522331316LTUA675574744LT'
    and returns a list'''
    logging.debug(f'Opening file to read: {txt_path}')
    with open(txt_path, 'r') as f:
        codes_block = f.read()
    # Code block to list:
    return wrap(codes_block, split_at_chars)

def edit_wb(xlsx_path):
    '''takes xlsx path, edits and saves workbook'''
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb[sheet_names[0]]
    # some_value = ws.cell(row=3, column=1).value
    last_row = ws.max_row
    logging.debug(f'Filling workbook with codes')
    add_codes(ws, last_row, 1)
    # NEW LINE BELOW
    logging.debug(f'Currently workbook contains {ws.max_row} codes')
    # NEW LINE ABOVE
    logging.debug(f'Workbook {xlsx_file} updated with new codes')
    wb.save(xlsx_path)

def add_codes(ws, last_row, fill_col):
    '''iterates simultaneously through list and cells to transfer postal codes to xlsx'''
    codes_list = convert_txt_to_list(txt_path)
    logging.debug(f'Read {txt_file} and formed a list codes_list containing {len(codes_list)} codes')
    # Handling the case of completely empty FreeCodes sheet, so filling would start at 1row instead of 2
    if last_row == 1:
        last_row = 0
    for idx in range(1, len(codes_list)+1):
        # Setting values:
        ws.cell(row=last_row + idx, column=fill_col).value = codes_list[idx-1]
    logging.debug('Finished looping through codes')

def print_last_row_number(xlsx_path):
    '''takes xlsx file, opens, reads last row number and prints out'''
    wb = openpyxl.load_workbook(xlsx_path)
    ws = wb['Free Codes']
    last_row = ws.max_row
    first_cell_data = ws.cell(row=1, column=1).value
    last_row_data = ws.cell(row=last_row, column=1).value
    logging.debug(f'Reading A1 data: {first_cell_data}')
    logging.debug(f'Reading Last Row Data: {last_row_data} that is found in row number: {last_row}')
    wb.save(xlsx_path)

def run():
    print(f'Logging to file: {log_path}')
    logging.debug('Run() start')
    edit_wb(wb_path)
    logging.debug(f'SCRIPT FINISHED WITHOUT ERRORS')
    print(f'SCRIPT FINISHED WITHOUT ERRORS')

if __name__ == "__main__":
    run()