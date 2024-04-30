






from data import generate_pdf, freq_table, report_by_region, report, report_all
import pandas as pd
import time
import sys
from utils.utils import preprocess_data
def start():
    while True:
        # df = generate_pdf()
        df = pd.read_excel('test.xlsx')
        # df = preprocess_data(df)
        
        if df.empty:
            continue
        
        # freq_table(df.copy())
        # out_filename = report(df.copy())
        # report_all(df, out_filename=out_filename)
        # df = pd.read_excel('data\\db_2023_06_15 - Copy.xlsx')
        # df = preprocess_data(df)
        report_by_region(df)
        # time.sleep(120)
        break

start() 

