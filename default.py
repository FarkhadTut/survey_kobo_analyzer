import os 

root = os.getcwd()
datafolder = 'data'
outfolder = 'out'
CSS = 'css/bootstrap.css'
outpath = os.path.join(root, outfolder)

URL = 'https://kf.cerrsurvey.uz/api/v2/assets/aJpTcaNZESVv3dYd2r3gdo/export-settings/es8ypw9wATVCkG8HJKUTUoh/data.xlsx'
HHPASSPORT = 'HR1. Сизнинг уй хўжалигингизда нечта одам истиқомат қилмоқда?'
TOKEN = '3bc477d2d6c4f572389685ee92229d4b87cd4c86'




SAME_COLUMNS = [('M12. Интервью натижалари:', 'M12. Интервью натижалари:.1')]


ASK = { HHPASSPORT:'superw_name',
        'superw_name':'HS1. Уй-жой тури?',
        'hr_closest_adult':'B3. Сиз айни дамда турмуш қурганмисиз ёки ким биландир бирга яшаябсизми?',
        'hr_closest_child04': 'B1. (Боланинг исми) қайси кун, ой ва йилда туғилган?',
        'hr_closest_child_male514': 'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?',
        'hr_closest_child_male1517':'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?.2',
        'hr_closest_child_female514':'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?.1',
        'hr_closest_child_female1517':'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?.3',}

CHANGE_YESNO = {
                'HS1. Уй-жой тури?': '"Уй хўжалиги саволномаси" бўлимидан ўтиш (агар ушбу уй хўжалиги саволноманинг бу бўлимидан олдин ўтган бўлса "ЙЎҚ" тугмасини босинг)',
                'B3. Сиз айни дамда турмуш қурганмисиз ёки ким биландир бирга яшаябсизми?': 232,
                'ub_age': 408,
                'cb_age': 543,
                'tn_age': 867,
                'cb_age_001': 699,
                'tn_age_001': 1040,
                'HR1. Сизнинг уй хўжалигингизда нечта одам истиқомат қилмоқда?':'"Уй хўжалиги рўйхати (HR)" бўлимидан ўтиш (агар ушбу уй хўжалиги саволноманинг бу бўлимидан олдин ўтган бўлса "ЙЎҚ" тугмасини босинг)',
                'passport_hh_group_conveyed':'HP8. Ассалому алайкум. Менинг исмим (исмингиз). Мен Ўзбекистон Республикаси Президенти ҳузуридаги Давлат статистика агентлигиданман. Биз камбағалликка таъсир этувчи турли омиллар, жумладан, турмуш шароити, соғлиқни сақлаш, таълим, ижтимоий ҳимоя ва ҳоказоларни яхшироқ тушуниш учун камбағаллик бўйича тадқиқот олиб бормоқдамиз. Ушбу маълумотлар ҳукуматга камбағалликни қисқартириш сиёсатини такомиллаштириш бўйича кейинги чора-тадбирларни режалаштиришда ёрдам беради. Мен сиз билан ана шу мавзулар юзасидан суҳбатлашмоқчи эдим. Ушбу интервью одатда 30 дақиқага яқин вақт олади. Уни кетидан яна сиз билан ёки уй хўжалигизнинг бошқа аъзолари билан алоҳида интервью ўтказишни сўрашим мумкин. Биз йиғадиган барча маълумотлар қатъий равишда маҳфий ва аноним тарзда сақланади. Агар сиз бирон саволга жавоб беришни истамасангиз ёки интервьюни тўхтатишни истасангиз, марҳамат қилиб менга мурожаат қилинг. Сиз сўровномада қатнашишга розимисиз?'}


REPORT_COLS = { HHPASSPORT:'Уй хўжалиги паспорти',
                'HS1. Уй-жой тури?':'Уй хўжалиги сўровномаси',
                'B3. Сиз айни дамда турмуш қурганмисиз ёки ким биландир бирга яшаябсизми?':'Вояга етганлар',
                'B1. (Боланинг исми) қайси кун, ой ва йилда туғилган?': '0 дан 4 ёш гача болалар',
                'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?': '(ЎҒИЛ) 5 дан 14 ёш гача болалар',
                'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?.2':'(ЎҒИЛ) 15 дан 17 ёш гача болалар',
                'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?.1':'(ҚИЗ) 5 дан 14 ёш гача болалар',
                'B1. (Боланинг исми) қайси йил/ой/кунда туғилган?.3':'(ҚИЗ) 15 дан 17 ёш гача болалар',}

STATIC_COLS = ['gen_uuid', 'hh_id_unique','cur_datetime', 'HP5. Уй хўжалиги IDси (ID HH):', 'HP1. Ҳудуд/вилоятни танланг:', 'HP2. Туманни кўрсатинг:', 'HP6. Интервью олувчининг исми ва IDси:']

COLUMNS = {'interviewers': 'Интервьюер коди:',
           'hhid': 'hh_id_unique',
           'extra_hhid': 'HP5. Уй хўжалиги IDси (ID HH):'}
