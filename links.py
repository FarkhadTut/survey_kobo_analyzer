from download import excel_to_pandas

# r = requests.get('https://kobo.humanitarianresponse.info/api/v2/assets/aQNE7tNyRVaRyvTxaPcVHS/data.json')

url = 'https://kobo.humanitarianresponse.info/api/v2/assets/aEeYjDP398CtEcfHrE4giK/export-settings/esCfUNpZFAfbiD4S75EnSxH/data.xlsx'




excel_to_pandas(url, 'data.xlsx')