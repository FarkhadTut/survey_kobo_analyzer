from telegram import (
    Bot, Update
)

import telegram
from telegram.ext import (
    Application,
    CommandHandler,
    ContextTypes,
    ConversationHandler,
    MessageHandler,
    Filters
)

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
interviewers = pd.read_excel('config\\interviwers.xlsx')

def check_file(filename):
    folder = 'out'
    today = str(datetime.today().date()).replace('-', '_')
    filepath_xl = os.path.join(folder, filename)
    if os.path.isfile(filepath_xl):
        df = pd.read_excel(filepath_xl)
        if df.empty:
            return False
        else:
            folder = 'pdf'
            print(filename)
            if 'state' in filename:
                filepath_pdf = os.path.join(folder, f'mdp_status_pdf_{today}.pdf')
            elif 'success' in filename:
                filepath_pdf = os.path.join(folder, f'mdp_success_pdf_{today}.pdf')
            if os.path.isfile(filepath_pdf):
                return filepath_pdf
    return False

async def send_document(filename, update, context):
    print('Analyzing...')
    chat_id = update.effective_chat.id
    message = await bot.send_message(text='Loading. Please, wait...', chat_id=chat_id)
    
    filepath = check_file(filename)
    if not filepath:
        await bot.send_message(text='No problems detected', chat_id=chat_id)
        await context.bot.deleteMessage(message_id = message.id,
                                    chat_id = update.message.chat_id)
        print('No problems detected')
        return 
    document = open(filepath, 'rb')
    await context.bot.send_document(chat_id, document)
    await context.bot.deleteMessage(message_id = message.id,
                                    chat_id = update.message.chat_id)
    print('File sent!')
    
async def get_status(update: Update, context: ContextTypes.DEFAULT_TYPE):
    today = str(datetime.today().date()).replace('-', '_')
    filename = f'state_analysis_{today}.xlsx'
    await send_document(filename, update, context)


async def generate_hhid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
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


async def ask_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    chat_id = update.effective_chat.id
    text = update.message.text
    interviewers['codes'] = 


bot = Bot(token=TOKEN)

def run():
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler(command='getstatus', callback=get_status))
    application.add_handler(CommandHandler(command='success', callback=get_successful))
    application.add_handler(CommandHandler(command='hhid', callback=generate_hhid))
    application.add_handler(ConversationHandler(
                                            entry_points=[CommandHandler('success', ask_id)],
                                            states={'GENERATE': [MessageHandler(Filters.text, get_successful)]}
                                                ))
    application.run_polling()


