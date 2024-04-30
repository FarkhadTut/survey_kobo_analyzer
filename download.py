import requests
import pandas as pd 
import os
import sys
import numpy as np
import json
from default import TOKEN
from utils.utils import preprocess_data, edit_db




# URL = 'https://kobo.humanitarianresponse.info/api/v2/assets/aEeYjDP398CtEcfHrE4giK/data/bulk/'
# PARAMS = {
#     'format': 'json'
# }


def excel_to_pandas(URL, local_path):
    # URL = 'https://kobo.humanitarianresponse.info/api/v2/assets/aEeYjDP398CtEcfHrE4giK/data/bulk/'

    raw_local_path = local_path.replace('db_', 'raw_db_')
    HEADERS = {
        'Authorization': f'Token {TOKEN}'
    }
    
    while True:
        try:
            resp = requests.get(
                url=URL,
                headers=HEADERS
            )
            if not resp.ok:
                print(f'Error downloading database: {resp}')
                return pd.DataFrame() 
            with open(raw_local_path, 'wb') as output:
                output.write(resp.content)
            df = pd.read_excel(raw_local_path)
            # df = edit_db(df)
            df.to_excel(raw_local_path, index=False)
            print('Preprossecing data...')
            df = preprocess_data(df)
            print('...done!')
            print('Local path:', local_path)
            df.to_excel(local_path, index=False)
            print('Data downloaded successfuly!')
            break
        except Exception as e:
            if 'zip' in str(e).lower():
                print('Bad zip file error skipped...')
                pass
            else:
                raise e
    print(df.empty, df.shape)
    if df.empty or df.shape[0] == 0:
        return pd.DataFrame()
    new_columns = []
    for c in df.columns:
        c = c.replace('\n', '').replace('\n', '').replace('\r', '').strip()
        new_columns.append(c)


    z_columns = dict(zip(df.columns.values.tolist(), new_columns))
    df = df.rename(z_columns, axis=1)
    print(local_path)
    return df

