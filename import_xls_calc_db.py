"""
Extracting data from multiple xls files and writing to a new one
using database SQLite3.

This module collects data from files (folder "calculations")
according to the list in the file "det_list _ *. xlsx"
and combines this data into one file.
"""
import sqlite3

from def_xls_db import *
from hf_xls_db import get_data_calc_hf
from calc_to_xls_db import calculation_data


def create_calc_file(date1):
    conn = sqlite3.connect(DB)
    cur = conn.cursor()
    create_tables(cur)

    imp_det_list(cur, date1)
    empty_list_1 = get_files_name(cur, PATH_C)
    empty_list = get_data_calc(cur, PATH_C, SHEET_NAME, empty_list_1)
    get_data_calc_hf(cur, PATH_HF)
    calculation_data(cur, date, empty_list)

    conn.commit()
    conn.close()


if __name__ == "__main__":
    # date = input("Enter the date of the detail group: ")
    date = '2022-04-01'
    create_calc_file(date)
