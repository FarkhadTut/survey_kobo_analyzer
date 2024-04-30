
from datetime import datetime
import pandas as pd
import numpy as np
import sys


def report():
    ############################################################################ Туманлар
    today = str(datetime.today().date()).replace('-', '_')
    report_filename = f'________report_analysis_{today}.xlsx'.format(today=today)
    df=pd.read_excel(f'out\\db_{today}.xlsx')
    
    
    df_orig = df.copy()
    df.rename(columns={'HP1. Ҳудуд/вилоятни танланг:': 'region', 'HP2. Туманни кўрсатинг:': 'district', 'gen_uuid': 'count'}, inplace=True)
    df = df[['region', 'district', 'count', 'cur_date']]
    dates = df_orig['cur_date'].unique()
    df_out = pd.DataFrame()
    for i, date in enumerate(sorted(dates)):
        mask = df['cur_date'] == date
        df_temp = df[mask]
        df_temp.drop('cur_date', axis=1, inplace=True)
        df_temp = df_temp.groupby(by=['district'], as_index=False).count()
        df_temp.rename(columns={'region':'Ҳудуд',  'district': 'Туман', 'count': date}, inplace=True)
        total = df_temp[date].sum()
        df_temp.drop('Ҳудуд', axis=1, inplace=True)
        
        
        if i == 0:
            df_out = df_temp
        else:
            df_out = pd.merge(df_out, df_temp, how='outer', on=['Туман'])
    
    
    total_col = df_out[dates.tolist()].sum(axis=1)
    df_out['Жами']= total_col

    # sys.exit()



    # df_total = pd.DataFrame(columns=df_out.columns, data=[['Жами:', np.nan, np.nan, np.nan, np.nan, np.nan, np.nan , total]])
    df_total = df_out[dates].sum()
    total_all = sum(df_total.values.tolist())
    df_total = pd.DataFrame(columns=['Туман'] +df_total.index.values.tolist(), data=[['Жами:']+df_total[dates].values.tolist()])
    
    df_out = pd.concat([df_out, df_total], axis=0)
    df_out.reset_index(drop=True, inplace=True)
    df_out.at[df_out.index.values.tolist()[-1], 'Жами'] = total_all
    with pd.ExcelWriter(f'out/report/{report_filename}', mode='w', engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='districts')

    # print(df_out)
    ############################################################################ Вилоятлар    
    df_orig.rename(columns={'HP1. Ҳудуд/вилоятни танланг:': 'region', 'HP2. Туманни кўрсатинг:': 'district', 'gen_uuid': 'count'}, inplace=True)
    df_orig = df_orig[['region', 'district', 'count', 'cur_date']]
    # df_orig = df_orig.groupby(by=['region'], as_index=False).count()[['region', 'count']]
    # df_orig.rename(columns={'region':'Ҳудуд', 'count': 'Уй хўжаликлар сони'}, inplace=True)
    df_out = pd.DataFrame()
    for k, date in enumerate(dates):
        mask = df_orig['cur_date'] == date
        df_temp = df_orig[mask]
        df_temp = df_temp.groupby(by=['region'], as_index=False).count()[['region', 'count']]
        total = df_temp['count'].sum()
        df_temp.rename(columns={'region':'Ҳудуд', 'count': date}, inplace=True)
        df_total = pd.DataFrame(columns=df_temp.columns, data=[['Жами:', total]])
        df_temp = pd.concat([df_temp, df_total], axis=0)
        df_temp.rename(columns={'region':'Ҳудуд', 'count': date}, inplace=True)
        if k == 0:
            df_out = df_temp
        else:
            df_out = pd.merge(df_out, df_temp, how='left', on=['Ҳудуд'])
        
    df_out['Жами'] = df_out[dates].sum(axis=1)
    
        
    with pd.ExcelWriter(f'out/report/{report_filename}', mode='a', engine='openpyxl') as writer:
        df_out.to_excel(writer, sheet_name='regions')
    print(f'{datetime.today()} - Starting generating REPORT file...')
    return report_filename



report()
