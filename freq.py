from datetime import datetime
import pandas as pd
import os 
import movecolumn as mc
import numpy as np
import sys
import math 
import movecolumn as mc
root = os.getcwd()

def rec_age(df):
    mask = df['b_age'] >= 500
    for i, row in df[mask].iterrows():
        df.at[i, 'b_age'] = math.floor(df.at[i, 'b_age']/365.25)
    return df


def freq_table(df_filename):
    
    if 'out' not in df_filename:
        df_filename = os.path.join(root, 'out', df_filename)
    print(df_filename)
    df = pd.read_excel(df_filename)
    # print(len(df['gender_18'].dropna()))
    # sys.exit()

    # print(ages.max())
    # sys.exit()
    df_orig = df.copy()
    start_col = 'GOV1. Мамлакатдаги умумий вазиятдан (ҳолатдан) қониқиш даражангизни 7 баллик шкалада баҳоланг:'
    end_col = 'GOV9.19. Олий Мажлис депутатлари?'

    today = str(datetime.today().date()).replace('-', '_')

    columns = df.columns.values.tolist()
    start_idx = columns.index(start_col)
    end_idx = columns.index(end_col)
    columns = columns[start_idx:end_idx] + ['HP3. Жой турини кўрсатинг:']

    DROP = [ 'GOV9. Қуйидаги институтларнинг ҳаётингиздаги аҳамияти ва ролини баҳоланг? -->']
    df = df[columns]
    df.drop(DROP, axis=1, inplace=True)
    df.reset_index(drop=True, inplace=True)

    columns = df.columns
    df_out = pd.DataFrame()

    filename_out = f'out/freq/___mdp_freq_table_{today}.xlsx'
    df_big = pd.DataFrame()
    town_vil = ['Шаҳар', 'Қишлоқ', 'Аёл', 'Эркак', 'all', '18-35', '36-49', '50-999']
    for i, c in enumerate(list(reversed(df.columns.values))):
        if c in ['HP3. Жой турини кўрсатинг:']:
                continue
        
        df_temp = pd.DataFrame()
        for j, tv in enumerate(town_vil):
            if tv in ['Шаҳар', 'Қишлоқ']:
                mask = df_orig['HP3. Жой турини кўрсатинг:'] == tv
                df = df_orig[mask]
            elif tv in ['18-35', '36-49', '50-999']:
                btm_age, top_age = tv.split('-')
                
                btm_age, top_age = int(btm_age), int(top_age)
                if btm_age == 18:
                    btm_age = 17
                btm_mask = df_orig['b_age'] >= btm_age
                top_mask = df_orig['b_age'] <= top_age
                # mask = ~(btm_mask & top_mask)
                df = df_orig[btm_mask][top_mask]
                df = df.dropna(subset='b_age', axis=0)
                
            elif tv in ['Аёл', 'Эркак']:
                mask = df_orig['gender_18'] == tv
                df = df_orig[mask]
            else:
                df = df_orig.copy()
            
            ctab = pd.crosstab(index = df.index, columns = df[c])
            ctab = ctab.sum()
           
            ctab_pct = (ctab/ctab.values.sum()).round(decimals=4)
            
            tab = pd.concat([ctab, ctab_pct], axis=1)

            cols = ['Frequency', 'Percentage (%)', c]
            tab[c] = tab.index
            df_out = pd.DataFrame(columns=cols, data=tab.values)
            df_out = mc.MoveTo1(df_out, c)

            if not df[c].dtype == 'object':
                df_out['Average score'] = df_out['Percentage (%)']/100 * df_out[c]
                print(df_out['Percentage (%)'], df_out[c])
                df_out['Average score'] = df_out['Average score'].round(decimals=3)
            else:
                df_out['Average score'] = [np.nan] * len(df_out['Percentage (%)'])
            
            df_out[['Percentage (%)', 'Average score']] = df_out[['Percentage (%)', 'Average score']] * 100

            
            freq_sum = df_out['Frequency'].sum()
            pct_sum =df_out['Percentage (%)'].sum()
            avg_sum = df_out['Average score'].sum()
            if avg_sum == 0:
                avg_sum = np.nan
            df_out['Frequency'] = df_out['Frequency'].astype(float).astype('Int64')
            df_total = pd.DataFrame(columns=df_out.columns, data=[['Total:', freq_sum,\
                                                                pct_sum,
                                                                 avg_sum]])
            df_out = pd.concat([df_out, df_total], axis=0)


            
            # Add space after tables
            df_out.rename(columns={'Frequency': f'{tv}_Frequency', 'Percentage (%)': f'{tv}_Percentage (%)', 'Average score': f'{tv}_Average score'}, inplace=True)
            if j == 0:
                df_temp = df_out
            else:
                if tv == 'all' or tv in ['Аёл'] or tv == '18-35':
                    space = pd.DataFrame(np.nan, index=list(range(len(df_out))), columns=['A', 'B'])
                    # space = df_out[df_out.columns.values[:2]].fill(0)
                    df_out.reset_index(drop=True, inplace=True)
                    df_out = pd.concat([space, df_out], axis=1)
                
         
                df_temp = pd.merge(df_temp, df_out, how='left', on=c)
                # if tv == 'all':
                #     print(df_out)
                #     print(df_temp)
                #     sys.exit()

        
        
        df_space = pd.DataFrame(columns=df_temp.columns.values.tolist(), data=[[np.nan]*len(df_temp.columns),
                                                                                [np.nan]*len(df_temp.columns)])
        
        df_temp = pd.concat([df_temp, df_space], axis=0)
    

        # columns = [c] + ['Frequency_x', 'Frequency_y', 'Percentage (%)_x',  'Percentage (%)_y', 'Average score_x', 'Average score_y', 'A', 'B', 'Frequency', 'Percentage (%)', 'Average score']
        # print(df_temp.columns.values)
        # sys.exit()
        columns = [c] + ['Шаҳар_Frequency', 'Қишлоқ_Frequency', 'Шаҳар_Percentage (%)',  'Қишлоқ_Percentage (%)', 'Шаҳар_Average score', 'Қишлоқ_Average score', 'A_x', 'B_x', \
                          'Аёл_Frequency', 'Эркак_Frequency', 'Аёл_Percentage (%)', 'Эркак_Percentage (%)', 'Аёл_Average score', 'Эркак_Average score', 'A_y', 'B_y',
                          '18-35_Frequency', '36-49_Frequency', '50-999_Frequency', '18-35_Percentage (%)', '36-49_Percentage (%)', '50-999_Percentage (%)', '18-35_Average score', '36-49_Average score', '50-999_Average score', 'A', 'B',
                          'all_Frequency', 'all_Percentage (%)', 'all_Average score']

   
        micolumns = pd.MultiIndex.from_tuples(
                                            [(None, c),
                                             ("Frequency", "Шаҳар"), (None, "Қишлоқ"), 
                                             ("Percentage (%)", "Шаҳар"), (None, "Қишлоқ"),
                                             ("Average score", "Шаҳар"), (None, "Қишлоқ"),
                                             (None, None),
                                             (None, None),
                                             ("Frequency", "Аёл"), (None, "Эркак"), 
                                             ("Percentage (%)", "Аёл"), (None, "Эркак"),
                                             ("Average score", "Аёл"), (None, "Эркак"),
                                             (None, None),
                                             (None, None),
                                             ("Frequency", '18-35'),(None, '36-49'),(None, '50+'), 
                                             ("Percentage (%)", '18-35'),(None, '36-49'),(None, '50+'),
                                             ("Average score", "18-35"),(None, "36-49"),(None, "50+"),
                                             (None, None),
                                             (None, None),
                                             ('Across all data', "Frequency"), 
                                             (None, "Percentage (%)"),
                                             (None, "Average score")],)

        
        df_temp =df_temp[columns]
        df_temp = pd.DataFrame(columns=micolumns, data=df_temp.values)

    
        
 
        if i == 0:
            df_big = df_temp
        else:
            
            df_temp = df_temp.reset_index(drop=True).T.reset_index().T
            df_temp.rename(columns=dict(zip(df_temp.columns.values.tolist(), df_big.columns.values.tolist())), inplace=True)
            if i == 1:
                columns = df_big.columns.values.tolist()
                df_big = df_big.reset_index(drop=True).T.reset_index().T
                df_big.rename(columns=dict(zip(df_big.columns.values.tolist(), columns)), inplace=True)
            
            df_big = pd.concat([df_temp, df_big], axis=0)

    
    df_big.rename(columns=dict(zip(df_big.columns.values.tolist(), ['MDP Frequency table', np.nan, np.nan, np.nan, np.nan, np.nan, today.replace('_', '-')])), inplace=True)
    
    with pd.ExcelWriter(filename_out, mode='w', engine='openpyxl') as writer:
        df_big.to_excel(writer, index=False)



    print(f'{datetime.today()} - Starting generating FREQ TABLE file...')
    return filename_out

file = 'out\\db_2023_03_17.xlsx'
freq_table(file)

print()