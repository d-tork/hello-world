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
    
    names = pd.read_csv('names.csv', squeeze=True)
    countries = pd.read_csv('countries.csv', squeeze=True)
    cats = ['cat1', 'cat2', 'cat3']
    
    rand_len = np.random.randint(100, 3000)
    
    df = pd.DataFrame(np.random.rand(rand_len, 4), columns=list('ABCD'))
    df['name'] = np.random.choice(names, size=rand_len)
    df['user_id'] = np.random.choice(user_id, size=rand_len)
    df['category'] = np.random.choice(cats, size=rand_len)
    df['A_text'] = np.random.choice(countries, size=rand_len)
    cols = ['name', 'user_id', 'category', 'A_text'] + list('ABCD')
    df = df[cols]
    
    # Add up numbers
    sumlist = list('ABCD')
    df['sum'] = df[sumlist].sum(axis=1)
    
    # Round the totals
    df['sum'] = df['sum'].round(1)
    return df

def fill_readme(sheet):
    txtlist = []
    with open('poc_info.txt') as f:
        for row in f:
            txtlist.append(row)
    sheet.range('H10').options(transpose=True).value = txtlist
    sheet.range('H:H').autofit()
    print('POC info written to README')
    return


# print(df.head())

# define additional lists
sheet_list = 'NCAA1 NCAA2 NCAA3 GSAC-3 FIVB APFT MOAR'.split()
file_list = 'East1 East2 East3 West1 West2 West3 NORTH SOUTH SOC PAC DICTIC DRY'.split()

# Open Excel instance
try:
    app1 = xw.apps[0]
except IndexError:
    app1 = xw.App(visible=False)
app1.screen_updating = False  # faster; don't show what it's doing

# set file names to loop through (may or may not be created)
fnames = [file_list[0]]  # single file name, for testing
fnames = np.random.choice(file_list, 5, replace=False).tolist()

for fname in fnames:
    fhand = copy2('my_template.xlsx', '{}.xlsx'.format(fname))
    fhand = os.path.abspath(fhand)
    rand_num = np.random.randint(1, 8)
    sheets = np.random.choice(sheet_list, rand_num, replace=False).tolist()

    wb = xw.Book(fhand)
    print('Opening:', wb.fullname)
    print('# of sheets: {}'.format(len(sheets)))
    existing_sheets = list(wb.sheets)
    # wfa_sheets = [x for x in wb.sheets if 'WFA' in str(x)]
    wfa_sheets = existing_sheets[1:-1]

    # write to excel
    for sheet, sht in zip(sheets, wfa_sheets):
        df = generate_data()
        sht.name = sheet
        sht.range('A5').options(index=False, header=False).value = df
        print('\tWrote: {} {} rows'.format(sht.name, len(df)))
    summary = wb.sheets('Player Summary')
    summary.range('Q2').options(transpose=True).value = sheets
    fill_readme(wb.sheets['README'])

    # delete unused sheets
    wfa_sheets = [x for x in wb.sheets if 'WFA' in str(x)]
    for sht in wfa_sheets:
        sht.delete()

    # Reset active
    wb.sheets[1].activate()

    wb.save()
    print('Saved:', wb.fullname)
    wb.close()

app1.quit()
input('Press enter to close this window.')
