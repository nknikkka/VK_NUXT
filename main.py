import telebot
from telebot import types
import pandas as pd
import threading
import schedule
import time

TOKEN = '7285439243:AAFDtqicqounDbLwzq_ULXLuoSdE_PZ2snk'
bot = telebot.TeleBot(TOKEN)

teacher_subscriptions = {}
student_subscriptions = {}

def get_schedule_from_excel():
    xls = pd.ExcelFile('groups.xlsx', engine='openpyxl')
    if 'schedule' in xls.sheet_names:
        schedule_df = pd.read_excel(xls, sheet_name='schedule')
    else:
        raise ValueError("Worksheet named 'schedule' not found")

    if 'substitutions' in xls.sheet_names:
        substitutions_df = pd.read_excel(xls, sheet_name='substitutions')
    else:
        substitutions_df = pd.DataFrame()

    return schedule_df, substitutions_df

@bot.message_handler(commands=['start'])
def start(message):
    chat_id = message.chat.id
    greeting_message = ("Привіт! 👋\n"
                        "Я — твоя права рука під час навчального року, в мене ти завжди можеш дізнатись, "
                        "які в тебе пари протягом тижня 😈\n\n"
                        "Хочеш більше дізнатися про бота?"
                        "Знаєш, як його покращити?"
                        "Щось пішло не так? \n"
                        "Пиши @Kyluk_Veronika, не соромся ")
    bot.send_message(chat_id, greeting_message)

    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    register_button = types.KeyboardButton('Зареєструватися')
    news_button = types.KeyboardButton('Новини 🌐')
    social_media_button = types.KeyboardButton('Соц-Мережі 📲')
    markup.add(register_button, news_button, social_media_button)
    bot.send_message(chat_id, "Вибери кнопку щоб продовжити:", reply_markup=markup)

@bot.message_handler(func=lambda message: message.text == 'Новини 🌐')
def news(message):
    chat_id = message.chat.id
    news_message = "Ось останні новини:\n" \
                   "Новини 1\n" \
                   "Новини 2\n" \
                   "Новини 3"
    bot.send_message(chat_id, news_message)

@bot.message_handler(func=lambda message: message.text == 'Соц-Мережі 📲')
def social_media(message):
    chat_id = message.chat.id
    social_message = "Наші соціальні мережі🇺🇦 :\n" \
                     "Instagram:https://www.instagram.com/vifk_nukht/\n" \
                     "Facebook: https://www.facebook.com/groups/385020108575468\n" \
                     "Telegram: https://t.me/vcnuftvn"
    bot.send_message(chat_id, social_message)

@bot.message_handler(func=lambda message: message.text == 'Зареєструватися')
def register(message):
    chat_id = message.chat.id
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    student_button = types.KeyboardButton('Студент')
    teacher_button = types.KeyboardButton('Викладач')
    back_button = types.KeyboardButton('Назад')
    markup.add(student_button, teacher_button, back_button)
    bot.send_message(chat_id, "Для початку, скажи мені, хто ти 😼:", reply_markup=markup)

@bot.message_handler(func=lambda message: True)
def handle_message(message):
    chat_id = message.chat.id
    text = message.text

    if text == 'Викладач':
        bot.send_message(chat_id, "Введіть своє прізвище:", reply_markup=types.ReplyKeyboardRemove())
        bot.register_next_step_handler(message, get_teacher_schedule)
        reset_to_main_menu(chat_id)
    elif text == 'Студент':
        bot.send_message(chat_id,
                         "Введи свою групу у форматі 'курс-група-номер', наприклад, '1-АМ-1' або '3-ОК-1'.", reply_markup=types.ReplyKeyboardRemove())
        bot.register_next_step_handler(message, get_group)
        reset_to_main_menu(chat_id)
    elif text == 'Назад':
        start(message)
    else:
        bot.send_message(chat_id, "Не розумію тебе! Спробуйте ще раз, будь ласка.")

def reset_to_main_menu(chat_id):
    markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    register_button = types.KeyboardButton('Зареєструватися')
    news_button = types.KeyboardButton('Новини 🌐')
    social_media_button = types.KeyboardButton('Соц-Мережі 📲')
    markup.add(register_button, news_button, social_media_button)

def get_teacher_schedule(message):
    chat_id = message.chat.id
    surname = message.text.strip()

    xls = pd.ExcelFile('groups.xlsx', engine='openpyxl')
    schedule_df = pd.read_excel(xls, sheet_name='schedule')

    print(f"Teacher surname entered: {surname}")
    # Check if the surname exists in any of the teacher columns
    teacher_schedule = schedule_df[
        (schedule_df['Викладач1'] == surname) |
        (schedule_df['Викладач1_1'] == surname) |
        (schedule_df['Викладач2'] == surname) |
        (schedule_df['Викладач2_2'] == surname) |
        (schedule_df['Викладач3'] == surname) |
        (schedule_df['Викладач3_3'] == surname) |
        (schedule_df['Викладач4'] == surname) |
        (schedule_df['Викладач4_4'] == surname)
    ]

    if teacher_schedule.empty:
        bot.send_message(chat_id, f"Вибачте, розклад для викладача {surname} не знайдено.")
    else:
        send_teacher_schedule(chat_id, teacher_schedule, surname)
        teacher_subscriptions[chat_id] = surname  # додаємо ідентифікатор чату до списку підписок викладачів

def send_teacher_schedule(chat_id, teacher_schedule, surname):
    schedule_message = f"Розклад для викладача {surname}:\n\n"
    days = teacher_schedule['День тижня'].unique()
    for day in days:
        day_schedule = teacher_schedule[teacher_schedule['День тижня'] == day]
        schedule_message += f"{day}:\n"
        for index, row in day_schedule.iterrows():
            for i in range(1, 5):
                if f'Викладач{i}' in row and row[f'Викладач{i}'] == surname:
                    pair = row[f'Пара {i}']
                    group = row['Група']
                    schedule_message += f"Пара {i} - {pair} ({group})\n"
                elif f'Викладач{i}_1' in row and row[f'Викладач{i}_1'] == surname:
                    pair = row[f'Пара {i}_1']
                    group = row['Група']
                    schedule_message += f"Пара {i} - {pair} ({group})\n"
        schedule_message += "\n"

    sent_message = bot.send_message(chat_id, schedule_message)
    bot.pin_chat_message(chat_id, sent_message.message_id)

def get_group(message):
    chat_id = message.chat.id
    text = message.text.strip()
    schedule_df, _ = get_schedule_from_excel()
    if text in schedule_df['Група'].values:
        student_subscriptions[chat_id] = text
        bot.send_message(chat_id, f"Ти підписався на групу '{text}'. Ось твій розклад на тиждень:")
        send_weekly_schedule(chat_id, text)
    else:
        bot.send_message(chat_id,
                         f"Група '{text}' не знайдена. Спробуйте ще раз, будь ласка, або зверніться до адміністратора для отримання додаткової інформації.")

def send_weekly_schedule(chat_id, group):
    schedule_df, _ = get_schedule_from_excel()
    schedule = schedule_df[schedule_df['Група'] == group]
    schedule_message = f"Розклад на тиждень для групи {group}:\n\n"

    for day in schedule['День тижня'].unique():
        day_schedule = schedule[schedule['День тижня'] == day]
        schedule_message += f"{day}:\n"

        for para_num in range(1, 5):
            subjects = []
            teachers = []

            for i in range(5):
                subject_col = f'Пара {para_num}' if i == 0 else f'Пара {para_num}_{i}'
                teacher_col = f'Викладач{para_num}' if i == 0 else f'Викладач{para_num}_{i}'

                if subject_col in day_schedule.columns and teacher_col in day_schedule.columns:
                    subject = day_schedule[subject_col].values[0]
                    teacher = day_schedule[teacher_col].values[0]

                    if pd.notna(subject) and subject != '-':
                        subjects.append(str(subject))
                    if pd.notna(teacher) and teacher != '-':
                        teachers.append(str(teacher))

            if subjects:
                schedule_message += f"Пара {para_num} - {'/'.join(subjects)} ({'/'.join(teachers)})\n"
            else:
                schedule_message += f"Пара {para_num} - -\n"

        schedule_message += "\n"

    sent_message = bot.send_message(chat_id, schedule_message)
    bot.pin_chat_message(chat_id, sent_message.message_id)

def send_teacher_substitution_updates():
    schedule_data, substitutions_df = get_schedule_from_excel()
    if substitutions_df is not None and not substitutions_df.empty:
        for chat_id, surname in teacher_subscriptions.items():
            teacher_substitutions = substitutions_df[
                (substitutions_df['Викладач1'] == surname) |
                (substitutions_df['Викладач2'] == surname) |
                (substitutions_df['Викладач3'] == surname) |
                (substitutions_df['Викладач4'] == surname)
            ]

            if not teacher_substitutions.empty:
                for index, substitution in teacher_substitutions.iterrows():
                    date_str = pd.to_datetime(substitution['Дата']).strftime('%Y-%м-%d')
                    substitution_message = f"Оновлення для викладача {surname} на {date_str}:\n"

                    for i in range(1, 5):
                        teacher_col = f'Викладач{i}'
                        pair_col = f'Пара {i}'

                        if teacher_col in substitution and substitution[teacher_col] == surname:
                            pair = substitution[pair_col]
                            if pair != '-':
                                substitution_message += f"Пара {i}: {pair}\n"

                    if substitution_message:
                        bot.send_message(chat_id, substitution_message)

def send_student_substitution_updates():
    schedule_data, substitutions_df = get_schedule_from_excel()
    if substitutions_df is not None and not substitutions_df.empty:
        for chat_id, group in student_subscriptions.items():
            group_substitutions = substitutions_df[substitutions_df['Група'] == group]
            if not group_substitutions.empty:
                for index, substitution in group_substitutions.iterrows():
                    date_str = pd.to_datetime(substitution['Дата']).strftime('%Y-%м-%d')
                    substitution_message = f"Оновлення для групи {group} на {date_str}:\n"
                    for i in range(1, 5):
                        pair = substitution[f'Пара {i}']
                        if pair != '-':
                            substitution_message += f"Пара {i}: {pair}\n"
                    if substitution_message:
                        bot.send_message(chat_id, substitution_message)

def schedule_daily_updates():
    schedule.every().day.at("21:03").do(send_teacher_substitution_updates)
    schedule.every().day.at("21:03").do(send_student_substitution_updates)
    while True:
        schedule.run_pending()
        time.sleep(1)

threading.Thread(target=schedule_daily_updates, daemon=True).start()

bot.remove_webhook()
bot.infinity_polling()
