import pandas as pd
import numpy as np
import xlwings as xw
from shutil import copy2
import os

wfa = pd.read_csv('18q3_full.csv')
wfa_org_list = wfa.org.unique()

if not os.path.exists('output'):
    os.mkdir('output')


def fill_readme(sheet):
    txtlist = []
    with open('poc_info.txt') as f:
        for row in f:
            txtlist.append(row)
    sheet.range('H10').options(transpose=True).value = txtlist
    sheet.range('H:H').autofit()
    print('POC info written to README')
    return


def write_workbook(org):
    outfile = os.path.join('output', '18q3 Roster - {}.xlsx'.format(org))
    fhand = copy2('my_template.xlsx', outfile)
    fhand = os.path.abspath(fhand)
    wb = xw.Book(fhand)
    print('Opening:', wb.fullname)

    df1 = wfa.loc[wfa.org == org]
    sht_list = df1.sheet.unique()
    print('# of sheets: {}'.format(len(sht_list)))
    existing_sheets = list(wb.sheets)[1:-1]
    for sheet, sht in zip(sht_list, existing_sheets):
        sht.name = sheet
        df2 = df1.loc[df1.sheet == sheet]
        sht.range('A5').options(index=False, header=False).value = df2
        print('\tWrote: {} {} rows'.format(sht.name, len(df2)))
    summary = wb.sheets('Player Summary')
    summary.range('Q2').options(transpose=True).value = sht_list
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


# Open Excel instance
try:
    app1 = xw.apps[0](visible=False)
except IndexError:
    app1 = xw.App(visible=False)
app1.screen_updating = False  # faster; don't show what it's doing

for org in wfa_org_list[:3]:
    write_workbook(org)

app1.quit()
