






from data import generate_pdf, freq_table, report_by_region, report, report_all, analyze_suspic
import pandas as pd
import time
import sys
from clean import clean

def start():
    while True:
        df = generate_pdf()
        # df = pd.read_excel('data\\db_2024_04_30.xlsx')
        # if df.empty:
        #     time.sleep(180)
        #     continue
        
        freq_table(df.copy())
        out_filename = report(df.copy())
        # report_all(df.copy(), out_filename=out_filename)
        # report_by_region((df.copy()))
        # analyze_suspic(df.copy())
        time.sleep(180)

start() 

