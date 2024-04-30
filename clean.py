import pandas as pd
from datetime import datetime


def clean(df):

    today = str(datetime.today().date()).replace('-', '_')

    df_list = pd.read_csv('data\\hhidfile.csv')
    df_list_old = pd.read_excel('data\\hhidfile_old.xlsx')

    df['address'] = df['1.1.1. Ҳудуд:'] + df['1.1.2. Туман:'] + df['1.1.3. Маҳалла:']
    df_list['address'] = df_list['region'] + df_list['district'] + df_list['mahalla']
    df = pd.merge(df, df_list[['address', 'name', 'pinfl']], how='left', on=['address', 'name'])
    print(df['pinfl'])



    for i, row in df.iterrows():
        if pd.isnull(row['pinfl']):
            mask = (df_list_old['name'] == row['name']) 
            df.at[i, 'pinfl'] = df_list_old[mask]['pinfl'].values[0]


    # print(df['pinfl'])



    df.drop_duplicates(subset=['pinfl'], keep='last', inplace=True)
    df.sort_values(by=['_submission_time'], ascending=True, inplace=True)
    df.dropna(how='all', axis=1, inplace=True)

    df.to_excel(f'data\\database_{today}.xlsx', index=False)

    return df