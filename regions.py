import pandas as pd 
import sys
import movecolumn as mc
from datetime import datetime


def report_by_region():
    today = str(datetime.today().date()).replace('-', '_')
    df_filename = 'out\\db_2023_03_19_1.xlsx'
    df = pd.read_excel(df_filename)
    cols = ['GOV5. Мамлакатимизда ўтказилиши режалаштирилаётган Референдум ҳақида хабардормисиз?',
           'GOV7. Мазкур Референдумда овоз беришда қатнашасизми?',
           'GOV6. Конституцияга киритилаётган ўзгаришларни қўллаб-қувватлайсизми?',
           'GOV8.1. Мамлакатнинг умумий иқтисодий ривожланиши?',
           'GOV8.13. Коррупция?', 'GOV8.14. Озиқ-овқат ва ёқилғи нархлари?']
    region_col = 'HP1. Ҳудуд/вилоятни танланг:'
    df_out = pd.DataFrame()
    for i, col in enumerate(cols):
        ctab = pd.crosstab(index = df[region_col], columns = df[col])
        ctab_pct = (pd.crosstab(index = df[region_col], columns = df[col], normalize='index') * 100).round(decimals=1)
        total_by_region = ctab.sum(axis=1)
        total_all = total_by_region.sum()
        total_by_region.name='Жами'
        # ctab = pd.concat([ctab, total_by_answers], axis=0)
        total_by_region= pd.DataFrame(total_by_region)
        ctab = pd.concat([total_by_region['Жами'], ctab], axis=1)
        
        total_by_answers = ctab.sum(axis=0)
        total_by_answers = total_by_answers.reset_index().T
        total_by_answers.rename(dict(zip(total_by_answers.columns.values.tolist(), total_by_answers.iloc[0].values.tolist())), inplace=True, axis=1)
       
        total_by_answers = total_by_answers.tail(-1)
        total_by_answers.index = ['Жами:']
        ctab = pd.concat([ctab, total_by_answers], axis=0)
        
        
        ctab = pd.concat([ctab['Жами'], ctab_pct], axis=1)
        ctab.drop(index='Жами:', inplace=True)
        total_by_answers_pct = (total_by_answers.div(total_all)).astype(float).multiply(100).round(decimals=1)
        # total_by_answers_pct['Жами'] = None
        total_by_answers_pct.index=['Жами, (%):']
        ctab = pd.concat([ctab, total_by_answers_pct], axis=0)
        ctab = pd.concat([ctab, total_by_answers], axis=0)
 
        # ctab = (pd.crosstab(index = df[region_col], columns = df[col], normalize='index') * 100).round(decimals=1)
        ctab[col] = ctab.index 
        ctab = mc.MoveTo1(ctab, col)
        ctab = ctab.reset_index(drop=True).T.reset_index().T
        space = pd.DataFrame(columns=ctab.columns, data=[[None]*len(ctab.columns), [None]*len(ctab.columns)])
        ctab = pd.concat([ctab, space], axis=0)
        ctab.reset_index(drop=True, inplace=True)
        
        # ctab.style.background_gradient(cmap='coolwarm').set_precision(2)
        # ctab.show()
        if i ==0:
            df_out = ctab
        else:
            df_out = pd.concat([df_out, ctab], axis=0)

    filename_out = f'out\\regional\\___regional_{today}.xlsx'
    sheet_name='regional'
    writer = pd.ExcelWriter(filename_out, engine='xlsxwriter', mode='w')
    
    df_out.rename(columns=dict(zip(df_out.columns.values.tolist(), [None]*len(df_out.columns.values.tolist()))), inplace=True)
    print(filename_out)
    df_out.to_excel(writer, sheet_name=sheet_name, index=False)
    

    workbook = writer.book
    wrap_format = workbook.add_format({'text_wrap': True})
    worksheet = writer.sheets[sheet_name]


    worksheet.conditional_format('C3:D17', {'type': '3_color_scale'})  
    worksheet.conditional_format('C22:F35', {'type': '3_color_scale'})  
    worksheet.conditional_format('C41:F54', {'type': '3_color_scale'})   
    worksheet.conditional_format('C60:M73', {'type': '3_color_scale'}) 
    worksheet.conditional_format('C79:M92', {'type': '3_color_scale'})   
    worksheet.conditional_format('C98:M111', {'type': '3_color_scale'})  
    
    worksheet.set_column('A:A', 60, wrap_format) 
    worksheet.set_column('B:D', 10, wrap_format) 
    worksheet.set_column('E:E', 15, wrap_format) 

    # worksheet.write(i,0,"My documentation text")
     
    merge_format = workbook.add_format({'align': 'center'}).set_font_size(14)
    worksheet.merge_range('B1:L1', 'Percentage crosstab, (%)', merge_format)


    writer.save()     

report_by_region()