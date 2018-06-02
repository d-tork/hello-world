# -*- coding: utf-8 -*-
"""
Spyder Editor

This is a temporary script file.
"""

import pandas as pd
import numpy as np
import matplotlib.pyplot as plt

s = pd.Series([1,3,5,np.nan,6,8])

user_id = np.random.randint(123456, 999999, size=300)

names = pd.read_excel('names.xlsx', squeeze=True)
countries = pd.read_excel('names.xlsx', 1, squeeze=True)

df = pd.DataFrame(np.random.rand(len(names), 4), columns=list('ABCD'))
df['name'] = names
df['user_id'] = np.random.choice(user_id, size=len(names))
df['A_text'] = np.random.choice(countries, size=len(names))
cols = ['name', 'user_id', 'A_text'] + list('ABCD')
df = df[cols]

# Add up numbers
sumlist = list('ABCD')
df['sum'] = df[sumlist].sum(axis=1)



print(df.head())