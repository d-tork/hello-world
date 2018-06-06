import pandas as pd
import numpy as np
import xlwings as xw
from shutil import copy2
import os
from datetime import datetime


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
    with open(logfile, append_write) as writer:
        writer.write(org + '\n')
    outfile = os.path.join('output', '18q3 Roster - {}.xlsx'.format(org))
    copy2('my_template.xlsx', outfile)
    wb = xw.Book(outfile)
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
        with open(logfile, append_write) as writer:
            writer.write('\t' + sheet + len(df2) + '\n')
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

# BEGIN RUNNING
# Create timestamp log
logfile = 'log.txt'
if os.path.exists(logfile):
    append_write = 'a'
else:
    append_write = 'w'
start = datetime.now()
startline = 'START: ' + str(start)
print(startline)
with open(logfile, append_write) as writer:
    writer.write(startline + '\n')

# Read in source file
wfa = pd.read_csv('18q3_full.csv')
wfa_org_list = wfa.org.unique()

# Check for output folder
if not os.path.exists('output'):
    os.mkdir('output')

# Open Excel instance
try:
    app1 = xw.apps[0]
except IndexError:
    app1 = xw.App(visible=False)
app1.screen_updating = True  # faster; don't show what it's doing

# choice = int(input('Which index for org? : '))

for org in wfa_org_list:
    write_workbook(org)
app1.quit()

end = datetime.now()
endline = 'FINISH:' + str(end)
durline = '\tin {:.10}'.format(str(end-start))
with open(logfile, append_write) as writer:
    writer.write(endline + '\n')
    writer.write(durline + '\n' + '-'*50 + '\n')
print(endline)
print(durline)
