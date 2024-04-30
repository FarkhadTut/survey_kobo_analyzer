import pandas as pd



df = pd.read_excel('data/db_2023_06_15')
df_list = pd.read_csv('data/hhidfile.csv')[['name', 'label']]


mask = pd.isnull(df['1.1.5. Ер эгасини танланг:'])
df['1.1.5. Ер эгасини танланг:'] = df['1.1.5. Ер эгасини танланг:'].fillna(df['1.1.5. Ер эгасини танланг:.1'])
df.drop(columns=['1.1.5. Ер эгасини танланг:.1'], inplace=True)

df = pd.merge(df, df_list, left_on='1.1.5. Ер эгасини танланг:', right_on='name')

print(df['label'])
