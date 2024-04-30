import pandas as pd
from openpyxl.cell.text import InlineFont
from openpyxl.styles import Font
from io import StringIO
import movecolumn as mc 
from datetime import datetime
from datetime import date as create_date 

def getIndexes(df, value):
     
    # Empty list
    listOfPos = []

    # isin() method will return a dataframe with
    # boolean values, True at the positions   
    # where element exists
    result = df.isin([value])

    # any() method will return
    # a boolean series
    seriesObj = result.any()

    # Get list of column names where
    # element exists
    columnNames = list(seriesObj[seriesObj == True].index)

    # Iterate over the list of columns and
    # extract the row index where element exists
    for col in columnNames:
        rows = list(result[col][result[col] == True].index)

        for row in rows:
            listOfPos.append((row, col))

    # This list contains a list tuples with
    # the index of element in the dataframe
    return listOfPos


def crosstab(index, columns, rename=None, remove_names=True):
    ctab = (pd.crosstab(index = index, columns = columns, normalize='index', margins=True, margins_name='Республика бўйича') * 100).round(decimals=1)
    if remove_names:
        ctab.columns.name = None
        ctab.index.name = None
        
    ctab = mc.MoveTo1(ctab.T, 'Республика бўйича').T

    ctab_count = (pd.crosstab(index = index, columns = columns, margins=True, margins_name='Республика бўйича'))
    if remove_names:
        ctab_count.columns.name = None
        ctab_count.index.name = None
    ctab_count = mc.MoveTo1(ctab_count.T, 'Республика бўйича').T
    ctab_count = ctab_count[ctab_count.columns.values[:-1]]

    if not rename is None:
        ctab_count.rename(columns=rename, inplace=True)
        ctab.rename(columns=rename, inplace=True)

    tab2 = ctab_count.join(ctab, lsuffix='_Сони', rsuffix='_%')

    tab2.columns = tab2.columns.map(lambda x: tuple(x.split('_')))

    tab2 = tab2.sort_index(ascending=[True, False] , axis=1)

    tab2.columns.name = None
    tab2.index.name = None
    ctab = tab2

    # if not remove_names:
    #     print(ctab)

    return ctab

def change_inn(df):
    df_change = pd.read_excel('edits\\change.xlsx')
    for i, row in df_change.iterrows():
        _id = row['_id']
        idx = df[df['_id'] == _id].index.values[0]
        new_inn = row['new_inn']
        df.at[idx, '####**6. Респондентни танланг:**'] = new_inn
        df.at[idx, '####**5. ТАРМОҚНИ ТАНЛАНГ:**'] = row['####**5. ТАРМОҚНИ ТАНЛАНГ:**']
        df.at[idx, '####**4. Рўхат турини танланг:**'] = row['####**4. Рўхат турини танланг:**']

    return df

def date_filter(df):
    mask = df['_submission_time'] >= pd.to_datetime(create_date(2023,8,9))
    return df[mask]


def drop_by_id(df):
    df_dups = pd.read_excel('edits\\duplicates.xlsx')
    drop_ids = df_dups[df_dups['good'] == 0]['_submission_id'].values.tolist()  
    mask = ~df['_id'].isin(drop_ids)
    return df[mask]


def change_values(df):
    sheet_names = ['Сугориш буйича', 'Стаж (йил)']
    for sheet_name in sheet_names:
        df_edits = pd.read_excel('edits\\edits.xlsx', sheet_name=sheet_name)
        for i, row in df_edits.iterrows():
            if sheet_name == 'Сугориш буйича':
                COL_CHANGE = '####**45. 1 га ерни суғориш учун қанча миқдорда сув ишлатасиз?**'
                inn = row['ИНН']
                value = row['Тўғри жавоб']
                mask = df['####**6. Респондентни танланг:**'] == inn
                if len(df[mask].values) != 1:
                    raise Exception(f'INN not unique: {inn}')
                else:
                    idx = df[mask].head(1).index[0]
                    df.at[idx, COL_CHANGE] = value
                    
            elif sheet_name == 'Стаж (йил)':
                COL_CHANGE = '####**8. Неча йилдан бери фаолият юритасиз?**'
                inn = row['ИНН']
                value = row['Тўғри жавоб']
                mask = df['####**6. Респондентни танланг:**'] == inn
                if len(df[mask].values) != 1:
                    raise Exception(f'INN not unique: {inn}')
                else:
                    idx = df[mask].head(1).index[0]
                    df.at[idx, COL_CHANGE] = value
                    
    return df


def edit_db(df):
    # df = change_inn(df)
    df = date_filter(df)
    # df = drop_by_id(df)
    # df = change_values(df)
    return df

def preprocess_data(df):
    # df.replace({'Бекобод шаҳри': 'Бекобод тумани'}, regex=True, inplace=True)
    #### only leave the successful ones #######################

    columns = df.columns.values
    # new_columns = [strip_tags(c) for c in columns]
    # z_columns = dict(zip(columns, new_columns))
    # df.rename(columns=z_columns, inplace=True)
    old_columns = []
    new_columns = []
    drop_columns = []
    for i, c in enumerate(columns):
        if c.startswith("#####"):
            if i+1 < len(columns) and not '<span' in c:
                old_c = columns[i+1]
                old_columns.append(old_c)
                new_c = strip_tags(old_c)
                new_c = new_c.strip(' ') + '/' + c.replace('#', '').strip(' ')
                new_columns.append(new_c)
                drop_columns.append(c)

    
    z_columns = dict(zip(old_columns, new_columns))
    df.rename(columns=z_columns, inplace=True)
    # df.drop(drop_columns, axis=1, inplace=True)
    today = str(datetime.today().date()).replace('-', '_')
    db_filename = 'db_{today}.xlsx'.format(today=today)

    new_columns = [c.replace('*', '').replace('#', '') for c in columns]
    z_columns = dict(zip(columns, new_columns))
    df.rename(columns=z_columns, inplace=True)


    # DF_LIST.rename(columns={'name': '_name_'}, inplace=True)
    # df = pd.merge(df, DF_LIST, left_on='6. Респондентни танланг:', right_on='_name_', how='left')
    # mask_date = df['_submission_time'] >= pd.to_datetime(create_date(2023, 7, 22))
    # df = df[mask_date]


    # df['5. ТАРМОҚНИ ТАНЛАНГ:'] = df['5. ТАРМОҚНИ ТАНЛАНГ:'].replace({'АСОСИЙ':'',
    #                                     'ЗАХИРА':""}, regex=True)
    
    # df['5. ТАРМОҚНИ ТАНЛАНГ:'] = df['5. ТАРМОҚНИ ТАНЛАНГ:'].map(lambda x: ''.join([c for c in x if not c.isnumeric()]))
    
    df.rename(columns={'1. Ҳудудни танланг?': 'region', '2. Туманингизни танланг?': 'district', '3. Маҳаллангизни танланг?': 'mahalla_code'}, inplace=True)
    
    df_mahalla = pd.read_csv('mahalla_data.csv')[['mahalla_code', 'mahalla']]
    df = pd.merge(df, df_mahalla, how='left', on='mahalla_code')
    
    df.drop_duplicates(subset=['region', 'district', 'mahalla_code'], keep='last', inplace=True)


    df.to_excel(f'data\\{db_filename}')
    return df


SIZE = 16
HEADER_SIZE = 20
FONT_NAME = 'Arial'
COLS = {'Давлат томонидан қуйида келтирилган соҳалардан қай бирида амалга оширилаётган ишларни маъқуллайсиз?': ['ишларни',
                                                                                                                'маъқуллайсиз'],
        'Шавкат Мирзиёевнинг Президентлик даврида эришган асосий ютуқларини номлаб беринг?':                   ['асосий',
                                                                                                                'ютуқларини'],
        '2017-2022 йилларда давлат сиёсатида асосий камчиликларни номлаб беринг?':                             ['асосий',
                                                                                                                'камчиликларни'],
        'Сиз яшаб турган ҳудуддаги энг асосий муаммони кўрсатинг:':                                             ['энг',
                                                                                                                'асосий',
                                                                                                                'муаммони'],
        'Қониқиш даражаси':                                                                                     ['Қониқиш',
                                                                                                                ],
        'Умумий вазият яхшиланмоқдами ёки деярли ўзгармаяптими?':                                             ['Умумий',
                                                                                                               'вазият',
                                                                                                                ],
        'Мамлакат ва ислоҳотларга ишонч':                                                                       ['Мамлакат',
                                                                                                               'ва',
                                                                                                               'ислоҳотларга',
                                                                                                               'ишонч'
                                                                                                                ],
        'Амалдаги Президентга ишонч':                                                                           ['Амалдаги',
                                                                                                                 'Президентга',
                                                                                                                 'ишонч',
                                                                                                                ],
        'Ҳукумат ва Парламент раҳбарларига муносабат':                                                          ['Ҳукумат',
                                                                                                                 'ва',
                                                                                                                 'Парламент',
                                                                                                                ],
        'Сиёсий партиялар номзодларига муносабат':                                                          ['Сиёсий',
                                                                                                                 'партиялар',
                                                                                                                 'номзодларига',
                                                                                                                ],
        'Вилоят ҳокимларига муносабат':                                                                         ['Вилоят',
                                                                                                                 'ҳокимларига',
                                                                                                                ],
        'Туман/шаҳар ҳокимларига муносабат':                                                                    ['Туман/шаҳар',
                                                                                                                 'ҳокимларига',
                                                                                                                ],
        'Парламент сайловларига муносабат':                                                                    ['Парламент',
                                                                                                                 'сайловларига',
                                                                                                                ],
        'Маълумотларнинг асосий манбаси':                                                                       ['асосий',
                                                                                                                 'манбаси',
                                                                                                                ],
        'Президент сайловига муносабат':                                                                       ['Президент',
                                                                                                                 'сайловига',
                                                                                                                 'муносабат',
                                                                                                                ],
        'Президент Шавкат Мирзиёев ўз лавозимида фикрингиз бўйича қандай фаолият кўрсатаётганлигини 7 баллик шкалада баҳоланг?':             ['Президент',
                                                                                                                                            'сайловига',
                                                                                                                                            'муносабат',
                                                                                                                                            ],   
        'Вилоят хокими қандай ишлаяпти?':                                                                       ['Вилоят',
                                                                                                                 'хокими',
                                                                                                                ],                                                                                                 
        'Туман ҳокими ўз лавозимида қандай ишламоқда?':                                                         ['Туман', "ҳокими"],                               
        'Олий Мажлис Қонунчилик палатасининг амалдаги таркибининг фаолиятини қандай баҳолайсиз?':               ['Олий', "Мажлис", "Қонунчилик", "палатасининг"], 
        'Мамлакатдаги умумий вазиятдан қониқиш даражангиз?':                                                    ['Мамлакатдаги'],
        'Вилоятингиздаги умумий вазиятдан қониқиш даражангиз?':                                                 ['Вилоятингиздаги'], 
        'Маҳаллангиздаги умумий вазиятдан қониқиш даражангиз?':                                                 ['Маҳаллангиздаги'],                               
        'Мамлакатдаги умумий вазият':                                                                           ['Мамлакатдаги'],                               
        'Вилоятдаги умумий вазият':                                                                             ['Вилоятдаги'],                               
        'Маҳалладаги умумий вазият':                                                                            ['Маҳалладаги'],                               
        '“Мен Ўзбекистон иқтисодиёти ривожланишига ишонаман”,  мазкур фикрга:':                                 ['“Мен', 'Ўзбекистон', 'иқтисодиёти', 'ривожланишига', 'ишонаман”'],                               
        '“Мамлакатимизда олиб борилаётган ислоҳотлар тўғри йўлда кетмоқда”, мазкур фикрга:':                    ['“Мамлакатимизда', 'олиб', 'борилаётган', 'ислоҳотлар', 'тўғри',
                                                                                                                 'йўлда', 'кетмоқда”'], 
        '“Ҳукумат фуқаролар билан очиқ мулоқотда бўлмоқда ва уларнинг муаммоларига ўз вақтида жавоб қайтармоқда" Мазкур фикрга:':    ['“Ҳукумат', 'фуқаролар', 'билан', 'очиқ', 'мулоқотда',
                                                                                                                                        'бўлмоқда', 'ва', 'уларнинг', 'муаммоларига', 'ўз',
                                                                                                                                        'вақтида', 'жавоб', 'қайтармоқда"'],
        'Абдулла Арипов, Ўзбекистон Республикаси Бош вазири?':                                                  ['Абдулла', 'Арипов,'],
        'Танзила Норбоева, Ўзбекистон Республикаси Олий Мажлиси Сенати раиси?':                                 ['Танзила', 'Норбоева,'],
        'Нуриддин Исмоилов, Ўзбекистон Республикаси Олий Мажлиси Қонунчилик палатасининг спикери?':             ['Нуриддин', 'Исмоилов,'],
        'Шавкат Мирзиёев, Ўзбекистон Либерал-демократик партиясидан президентликка номзод?':                    ['Шавкат', 'Мирзиёев,'],
        'Роба Маҳмудова, Ўзбекистон «Адолат»  демократик партиясидан президентликка номзод?':                   ['Роба', 'Маҳмудова,'],
        'Улуғбек Иноятов, Ўзбекистон Халқ демократик партиясидан президентликка номзод?':                       ['Улуғбек', 'Иноятов,'],
        'Абдушукур Ҳамзаев, Ўзбекистон Экологик партиядан президентликка номзод?':                              ['Абдушукур', 'Ҳамзаев,'],
                                       


                                                                                                                }

def get_rich_text(s, size=False):
    
    if isinstance(s, str):
        s = s.strip()
    if not size:
        size = HEADER_SIZE
    red = InlineFont(color="C00000", b=True, sz=size, rFont=FONT_NAME)
    blue = InlineFont(color="1f4e78", b=True, sz=size, rFont=FONT_NAME)
    rich = CellRichText()
    if s in list(COLS.keys()):
        words = COLS[s]
        split_header = s.split(' ')
        
        if '?' in split_header[-1]:
            split_header[-1] = split_header[-1][:-1]
            split_header.append('?')
        for i, w in enumerate(split_header):
            if i != len(split_header) - 1 and split_header[i+1] != '?':
                w += ' '
            if w.strip() in words:
                t = TextBlock(red, w)
            else:
                t = TextBlock(blue, w)
            rich.append(t)
    else:
        return s
    
    return rich
    
from html.parser import HTMLParser

class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.text = StringIO()
    def handle_data(self, d):
        self.text.write(d)
    def get_data(self):
        return self.text.getvalue()



def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()