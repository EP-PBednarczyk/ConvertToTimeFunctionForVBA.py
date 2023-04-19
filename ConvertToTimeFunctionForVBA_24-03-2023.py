# string_with_time = ["abc#@!?efg;:*$**?***"] #test
import string


def convert_to_time(in_included_time_prohibited_sign):

    # convert string to Time 00:00:00 [hours:minutes:seconds]'
    # remove chars(words) from times
    # string_with_time.remove()  # remove chars(words) from times
    # string_with_time = string_with_time.translate({ord(c): None for c in '***'})

    # ASCII: 33 to 47 - !,.../
    str_lower_word = list(string.ascii_lowercase)
    str_upper_word = list(string.ascii_uppercase)

    #special owns:
    prohibited_characters = ['*', '**', '***', '****', '!', '@', '#',
                                 '$', '%', '^', '&', '(', ')', '-', '+',
                                 '_', '=', '[', ']', '{', '}', ';', "'",
                                 ':', '"', '\\', '|', '<', '>', ',', '.',
                                 '/', '?', '|', '`', '~',
                                 str_lower_word, str_upper_word]

    return ''.join(c for c in in_included_time_prohibited_sign if c not in prohibited_characters)

