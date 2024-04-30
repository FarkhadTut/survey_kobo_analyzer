from typing import Optional, Union
import time
from telegram import (
    Bot, Update
)

from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler ,
    MessageHandler,
    BaseHandler,
    filters
)

from telegram.ext.filters import  MessageFilter

from telegram import ReplyKeyboardMarkup, ReplyKeyboardRemove, Update
from telegram.constants import ParseMode

import logging
from config.config import TOKEN
import os
from datetime import datetime
from data import generate_pdf
import pandas as pd 
import random

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)

hhid_folder = 'hhid'
hhid_filename = 'hhid.xlsx'
logger = logging.getLogger(__name__)
# df_locations = pd.read_excel('data/locations.xlsx')

def check_file(filename):
    folder = 'out'
    today = str(datetime.today().date()).replace('-', '_')
    filepath_xl = os.path.join(folder, filename)
    if 'report' in filename:
        filepath_xl = os.path.join(folder, 'report', filename)
        if os.path.isfile(filepath_xl):
            return filepath_xl
    
    if 'freq_table' in filename:
        filepath_xl = os.path.join(folder, 'freq', os.path.basename(filename))
        if os.path.isfile(filepath_xl):
            return filepath_xl
    if 'regional' in filename:
        filepath_xl = os.path.join(folder, 'regional', os.path.basename(filename))
        if os.path.isfile(filepath_xl):
            return filepath_xl
    
    if 'database' in filename:
        filepath_xl = os.path.join('data', os.path.basename(filename))
        if os.path.isfile(filepath_xl):
            return filepath_xl
    if 'db_' in filename:
        filepath_xl = os.path.join('data', os.path.basename(filename))
        if os.path.isfile(filepath_xl):
            return filepath_xl
    if os.path.isfile(filepath_xl):
        df = pd.read_excel(filepath_xl)
        if df.empty:
            return False
        else:
            folder = 'pdf'
            print(filename)
            if 'state' in filename:
                filepath_pdf = os.path.join(folder, f'ecopol_status_pdf_{today}.pdf')
            elif 'success' in filename:
                filepath_pdf = os.path.join(folder, f'ecopol_success_pdf_{today}.pdf')
            if os.path.isfile(filepath_pdf):
                return filepath_pdf
    return False

async def send_document(filename, update, context):
    print('Analyzing...')
    chat_id = update.effective_chat.id
    message = await bot.send_message(text='Loading. Please, wait...', chat_id=chat_id)
    
    print(filename)
    filepath = check_file(filename)
    print(filepath)
    if not filepath:
        text = 'Муаммоли сўровнома топилмади.'
        if 'succes' in filename:
            text = 'Тугалланган сўровномалар хозирча мавжуд эмас.'
        if 'report' in filename:
            text = '.'

        await bot.send_message(text=text, chat_id=chat_id)
        await context.bot.deleteMessage(message_id = message.id,
                                    chat_id = update.message.chat_id)
        return 
    document = open(filepath, 'rb')
    await context.bot.send_document(chat_id, document)
    await context.bot.deleteMessage(message_id = message.id,
                                    chat_id = update.message.chat_id)
    username = update.message.from_user.username
    print(f'File sent to {username} at {datetime.now().date()}!')
    
async def get_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'state_analysis_{today}.xlsx'
    await send_document(filename, update, context)


def is_supervisor(username, chat_id):
    if is_mega(username, chat_id):
        return True
    if username is None:
        return False
    username = str(username).lower()
    chat_id = str(chat_id)
    root = os.getcwd()
    with open(os.path.join(root, 'config', 'supervisors.txt'), 'r') as f:
        sups = f.readlines()
    
    sups = [str(s).replace('\n', '').strip().lower() for s in sups]
    
    if username in sups or chat_id in sups:
        return True
    return False


def is_mega(username, chat_id):
    if username is None:
        return False
    username = str(username).lower()
    chat_id = str(chat_id)
    root = os.getcwd()
    with open(os.path.join(root, 'config', 'megausers.txt'), 'r') as f:
        sups = f.readlines()
    
    sups = [str(s).replace('\n', '').strip().lower() for s in sups]
    
    if username in sups or chat_id in sups:
        return True
    return False



async def generate_hhid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if not is_supervisor(username, chat_id):
        text = 'Уй хўжалиги IDсини фақат супервайзерлар ярата олади.'
        await bot.send_message(text=text, chat_id=chat_id)
        return 
    hhid_file = os.path.join(hhid_folder, hhid_filename)
    used_hhid = []
    while True:
        hhid = random.randint(998000, 998999)
        if os.path.isfile(hhid_file):
            df = pd.read_excel(hhid_file)
            used_hhid = df['used_id'].values.tolist()
        else:
            df = pd.DataFrame(columns=['used_id'])
        if not hhid in used_hhid:
            df = pd.DataFrame(columns=['used_id'], data=df['used_id'].values.tolist() + [hhid])
            df.to_excel(hhid_file, index=False)
            break

    await bot.send_message(text=f'New HHID:\n*{hhid}*', chat_id=chat_id, parse_mode=ParseMode.MARKDOWN)


async def get_successful(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'success_analysis_{today}.xlsx'
    await send_document(filename, update, context)


async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    report_filename = 'report_analysis_{today}.xlsx'.format(today=today)
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if os.path.isfile(os.path.join(root, 'out', 'report', report_filename)):
        if is_supervisor(username=username, chat_id=chat_id):
            await send_document(report_filename, update, context)
            return
        else:
            text = 'Маълумот топилмади.' #'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    else:
        text = 'Маълумот топилмади.'
    await bot.send_message(text=text, chat_id=chat_id)




async def all_report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    report_filename = 'all_report_analysis_{today}.xlsx'.format(today=today)
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    print(report_filename)
    if os.path.isfile(os.path.join(root, 'out', 'report', report_filename)):
        if is_mega(username=username, chat_id=chat_id):
            await send_document(report_filename, update, context)
            return
        else:
            text = 'Маълумот топилмади.' #'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    else:
        text = 'Маълумот топилмади.'
    await bot.send_message(text=text, chat_id=chat_id)


async def not_done(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    report_filename = 'not_done_{today}.xlsx'.format(today=today)
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if os.path.isfile(os.path.join(root, 'out', 'report', report_filename)):
        if is_supervisor(username=username, chat_id=chat_id):
            await send_document(os.path.join(root, 'out', 'report', report_filename), update, context)
            return
        else:
            text = 'Маълумот топилмади' #'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    else:
        text = 'Маълумот топилмади.'
    await bot.send_message(text=text, chat_id=chat_id)


async def get_freq(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'out\\freq\\hy_freq_table_{today}.xlsx'
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if is_mega(username, chat_id):
        print(os.path.join(root, filename))
        if os.path.isfile(os.path.join(root, filename)):
            await send_document(filename, update, context)
            return
        else:
            text = 'Маълумот топилмади'

    else:
        text = 'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    await bot.send_message(text=text, chat_id=chat_id)



async def get_regional(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'regional/regional_{today}.xlsx'
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if is_mega(username, chat_id):
        if os.path.isfile(os.path.join(root, 'out', filename)):
            await send_document(filename, update, context)
            return
        else:
            text = 'Маълумот топилмади'

    else:
        text = 'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    await bot.send_message(text=text, chat_id=chat_id)



async def get_clean_database(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'data\\db_{today}.xlsx'
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if is_mega(username, chat_id):
        if os.path.isfile(os.path.join(root, filename)):
            await send_document(filename, update, context)
            return
        else:
            text = 'Маълумот топилмади'

    else:
        text = 'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    await bot.send_message(text=text, chat_id=chat_id)


async def get_raw_database(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'data\\raw_db_{today}.xlsx'
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if is_mega(username, chat_id):
        if os.path.isfile(os.path.join(root, filename)):
            await send_document(filename, update, context)
            return
        else:
            text = 'Маълумот топилмади'

    else:
        text = 'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    await bot.send_message(text=text, chat_id=chat_id)


async def get_suspicious(update: Update, context: ContextTypes.DEFAULT_TYPE):
    root = os.getcwd()
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'out\\suspicious\\suspicious_{today}.xlsx'
    chat_id = update.effective_chat.id
    username = update.message.from_user.username
    if is_mega(username, chat_id):
        if os.path.isfile(os.path.join(root, filename)):
            await send_document(filename, update, context)
            return
        else:
            text = 'Маълумот топилмади'
    else:
        text = 'Бу маълумотни кўриш учун сизда етарлича рухсат йўқ.'
    await bot.send_message(text=text, chat_id=chat_id)










bot = Bot(token=TOKEN)

def run():
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler(command='ftbl', callback=get_freq))
    application.add_handler(CommandHandler(command='report', callback=report))
    # application.add_handler(CommandHandler(command='not_done', callback=not_done))
    application.add_handler(CommandHandler(command='clean', callback=get_clean_database))
    application.add_handler(CommandHandler(command='raw_db', callback=get_raw_database))
    # application.add_handler(CommandHandler(command='regional', callback=get_regional))
    # application.add_handler(CommandHandler(command='suspiciuos', callback=get_suspicious))







    application.run_polling()

