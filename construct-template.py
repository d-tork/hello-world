"""
Builds the framework for the test template.
"""

import xlwings as xw
import pandas.DataFrame as DF
import os

# Open Excel instance
try:
    app1 = xw.apps[0]
except IndexError:
    app1 = xw.App(visible=False)

wb = xw.Book()
sht = wb.sheets[0]
sht.name = 'README'
for i in range(1,8):
    shtname = 'WFA({})'.format(i)
    wb.sheets.add(name=shtname, after=sht)
    sht = wb.sheets[shtname]
xw.sheets.add(name='Player Summary' after=sht)


def fill_wfa(sheet):
    head1 = 'A1:Z1'
    head2 = 'A2:Z2'
    col_rng = 'A3:I3'

    head1_txt = 'Beautiful header!'
    head2_txt = 'Lighter sub-header'
    cols = ['name', 'user_id', 'category',
            'A_text', 'A_1', 'B_1',
            'C_1', 'D_1', 'Total Points']

    head1_color = (174, 170, 170)
    head2_color = (231, 230, 230)

    sheet.range('A1').value = head1_txt
    sheet.range('A2').value = head2_txt
    sheet.range(head1).color = head1_color
    sheet.range(head2).color = head2_color
    sheet.range(col_rng).value = cols
    return


def fill_summary(sheet):
    head = ['Sheets', 'cat1', 'cat2', 'cat3']
    rng1 = 'Q1:T1'

    sheet.range(rng1).value = head
    return

# Fill worksheets
fill_summary(wb.sheets['Player Summary'])
for sht in list(xw.sheets)[1:-1]:
    fill_wfa(sht)

if not os.path.exists('export'):
    os.mkdir('export')
wb.save(os.path.normpath('export/my_template.xlsx'))
print('Saved:', wb.fullname)
wb.close()

app1.quit()