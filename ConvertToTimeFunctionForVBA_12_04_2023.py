import string
import pandas as pd  # to the excel
# windowmessage
import ctypes  # An included library with Python install.

# ------------------------------------------------------------------------------------- #
# import warnings
# from openpyxl import *  # library to the excel - GNU license
# from openpyxl import Workbook, load_workbook  # library to the excel - GNU license


# to the test
# _input_included_time_wrong_symbol = ["abc#@!?efg;:*$**?***08:00:00"]  # test

# wb = load_workbook('czas_pracy_1.3.24_Pawel_Bednarczyk_IV_2023_04-04.xlsm', keep_vba=True)
# ws = wb.active

# print(ws)
# warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')

d = pd.read_excel('czas_pracy_1.4.5_Pawel_Bednarczyk_IV_2023_04-12.xlsm')  # load xls file ->from pandas
_input_included_time_wrong_symbol = d.at[12, 12]  # .at[numer wiersza, numer kolumny z excela]
_input_included_time_wrong_symbol.append(d.at[12, 13])
_input_included_time_wrong_symbol.append(d.at[12, 14])
_input_included_time_wrong_symbol.append(d.at[12, 15])

print(_input_included_time_wrong_symbol)

def convert_to_time(_input_included_time_wrong_symbol):
    # convert string to Time 00:00:00 [hours:minutes:seconds]'
    # remove chars(words) from times
    # string_with_time.remove()  # remove chars(words) from times
    # string_with_time = string_with_time.translate({ord(c): None for c in '***'})

    # ASCII: 33 to 47 - !,.../
    str_lower_word = list(string.ascii_lowercase)
    str_upper_word = list(string.ascii_uppercase)

    # special owns:
    _prohibited_characters = ['*', '**', '***', '****', '!', '@', '#',
                              '$', '%', '^', '&', '(', ')', '-', '+',
                              '_', '=', '[', ']', '{' '}', ';', "'",
                              '"', '\\', '|', '<', '>', ',', '.',
                              '/', '?', '|', '`', '~'] + str_lower_word + str_upper_word
    #  unecessary marks going to remove
    _fixed_string: str = ''.join(c for c in _input_included_time_wrong_symbol[0] if c not in _prohibited_characters)

    # wsadzenie z powrotem do excela
    d.at[12, 12] = _fixed_string[0]
    d.at[12, 13] = _fixed_string[1]
    d.at[12, 14] = _fixed_string[2]
    d.at[12, 15] = _fixed_string[3]
    # -----------------to the debug ---------------------------------------
    # windowd message
    ctypes.windll.user32.MessageBoxW(0, "Your text  "+d, "Your title", 1)

    # -------------------------------------------------------------------

    return


# ToDo issue when case ":" ":08:00:00    08:00:00:, 08:0:0:0:0 etc"
# if c == ":"

# if __name__ == '__main__':
#    print(convert_to_time())
# warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
