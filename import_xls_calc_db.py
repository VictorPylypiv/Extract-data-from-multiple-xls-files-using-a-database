"""
Extracting data from multiple xls files and writing to a new one
using database SQLite3.

This module collects data from files (folder "calculations")
according to the list in the file "det_list _ *. xlsx"
and combines this data into one file.
"""
from def_xls_db import *
from hf_xls_db import get_data_calc_hf
from calc_to_xls_db import calculation_data


def create_calc_file(date1):
    create_tables(DB)

    imp_det_list(DB, date1)
    empty_list_1 = get_files_name(DB, PATH_C)
    empty_list = get_data_calc(DB, PATH_C, SHEET_NAME, empty_list_1)
    get_data_calc_hf(DB, PATH_HF)
    calculation_data(date, empty_list)


if __name__ == "__main__":
    date = input("Enter the date of the detail group: ")
    # date = '2022-04-01'
    create_calc_file(date)
