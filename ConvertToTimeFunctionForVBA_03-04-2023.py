import string
import warnings
# from openpyxl import *  # library to the excel - GNU license
from openpyxl import Workbook, load_workbook  # library to the excel - GNU license

_input_included_time_wrong_symbol = ["abc#@!?efg;:*$**?***"]  # test
wb = load_workbook('czas_pracy_1.3.24_Pawel_Bednarczyk_IV_2023_04-03.xlsm', keep_vba=True)
ws = wb.active

warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
print(ws)
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')


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
                              '_', '=', '[', ']', '{', '}', ';', "'",
                              ':', '"', '\\', '|', '<', '>', ',', '.',
                              '/', '?', '|', '`', '~',
                              str_lower_word, str_upper_word]

    return ''.join(c for c in _input_included_time_wrong_symbol if c not in _prohibited_characters)


#if __name__ == '__main__':
#    print(convert_to_time())
warnings.filterwarnings('ignore', category=UserWarning, module='openpyxl')
