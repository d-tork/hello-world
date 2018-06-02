# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
import xlwings as xw
from shutil import copy2
import os


def generate_data():
    user_id = np.random.randint(123456, 999999, size=3000)
    
    names = pd.read_excel('names.xlsx', squeeze=True)
    countries = pd.read_excel('names.xlsx', 1, squeeze=True)
    
    rand_len = np.random.randint(100, 3000)
    
    df = pd.DataFrame(np.random.rand(rand_len, 4), columns=list('ABCD'))
    df['name'] = np.random.choice(names, size=rand_len)
    df['user_id'] = np.random.choice(user_id, size=rand_len)
    df['A_text'] = np.random.choice(countries, size=rand_len)
    cols = ['name', 'user_id', 'A_text'] + list('ABCD')
    df = df[cols]
    
    # Add up numbers
    sumlist = list('ABCD')
    df['sum'] = df[sumlist].sum(axis=1)
    
    # Round the totals
    df['sum'] = df['sum'].round(1)
    return df


# print(df.head())

# define additional lists
sheet_list = 'APRC1 APRC2 APRC3 JTFOA APTF MOAR'.split()
file_list = 'East1 East2 East3 West1 West2 West3'.split()

# app1 = xw.App(visible=False)
app1 = xw.apps[0]
fnames = file_list[3:]
app1.screen_updating = False

for fname in fnames:
    fname = copy2('my_template.xlsx', '{}.xlsx'.format(fname))
    sheets = np.random.choice(sheet_list, 4, replace=False).tolist()
    # fname = os.path.abspath(fname)

    # write to excel    
    wb = xw.Book(fname)
#    wbshts = [x for x in wb.sheets if 'WFA' in str(x)]
    wbshts = [x for x in wb.sheets][1:-1]
    for sheet, sht in zip(sheets, wbshts):
        df = generate_data()
        sht.name = sheet
        sht.range('A4').options(index=False, header=False).value = df
        print('Wrote: {} {} rows'.format(sht.name, len(df)))
    summary = wb.sheets('Player Summary')
    summary.range('Q2').options(transpose=True).value = sheets
    wb.save()
    print('Saved:', wb.fullname)
    wb.close()

# app1.quit()
