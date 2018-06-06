import pandas as pd
import numpy as np
import os
import sys
import string

cols = ['player_id', 'player_name', 'category', 'org',
        'country_01', 'description_01', 'amount_01',
        'country_02', 'description_02', 'amount_02',
        'country_03', 'description_03', 'amount_03',
        'country_04', 'description_04', 'amount_04',
        'country_05', 'description_05', 'amount_05',
        'country_06', 'description_06', 'amount_06',
        'country_07', 'description_07', 'amount_07',
        'country_08', 'description_08', 'amount_08',
        'country_09', 'description_09', 'amount_09',
        'country_10', 'description_10', 'amount_10', 'total_amount', 'comments',
        'sheet']
sheet_list = 'NCAA1 NCAA2 NCAA3 GSAC-3 FIVB APFT MOAR'.split()
org_list = """East1 East2 East3 West1 West2 West3 
           NORTH SOUTH SOC PAC DICTIC DRY
           ONECOM TWOCOM THREECOM FOURCOM FIVECOM SIXCOM
           DODIC PALIC CHIPOTLIC PAPAJOHN DOMINO CHICFIL
           HEIN BUDWEIS JDAN TITO GREYG BERRY""".split()


def gen_word(length=12):
    allchar = list(string.ascii_letters + string.punctuation)
    word = ''.join(np.random.choice(allchar, size=None) for x in range(np.random.randint(5, length)))
    return word


def fill_cols(df, cols, row_ct, org):
    names = pd.read_csv('names.csv', squeeze=True)
    countries = pd.read_csv('countries.csv', squeeze=True)
    cats = ['cat1', 'cat2', 'cat3']

    for col in cols:
        if 'country' in col:
            df[col] = np.random.choice(countries, size=row_ct)
        if 'description' in col:
            df[col] = gen_word(25)
        if 'amount' in col:
            df[col] = np.random.rand(row_ct, 1)
    df['player_id'] = np.random.randint(111111, 999999, size=row_ct)
    df['player_name'] = np.random.choice(names, size=row_ct)
    df['category'] = np.random.choice(cats, size=row_ct)
    df['org'] = org
    df['sheet'] = np.random.choice(sheet_list, size=row_ct)
    sumlist = cols[7:-3]
    df['total_amount'] = df[sumlist].sum(axis=1).round(1)
    return df


def gen_df(org, row_ct=200):
    df = pd.DataFrame(columns=cols)
    df = fill_cols(df, cols, row_ct, org)
    return df


def gen_file(sht_ct=1):
    rand_org = np.random.choice(org_list)
    writer = pd.ExcelWriter('18Q3 Roster-{}.xlsx'.format(rand_org))
    for i in range(sht_ct):
        rand_ct = np.random.randint(50, 1000)
        df = gen_df(org=rand_org, row_ct=rand_ct)
        sheet = np.random.choice(sheet_list)
        df.to_excel(writer, sheet_name=sheet)
    writer.save()


gen_file(sht_ct=4)

df1 = pd.DataFrame()
for i in range(29):
    org_name = np.random.choice(org_list)
    rand_ct = np.random.randint(50, 1000)
    df1 = df1.append(gen_df(org_name, rand_ct))
df1.to_csv('18q3_full.csv', index=False)


