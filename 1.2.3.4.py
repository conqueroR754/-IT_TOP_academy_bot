import telebot
import pandas as pd
import re
from io import BytesIO

# Ваш токен для бота
TOKEN = '7785888241:AAGBdDryyw1MiQFlfwwChByROxmWZP-4jm8'
bot = telebot.TeleBot(TOKEN)

# Функция для нормализации столбцов
def normalize_columns(df):
    column_mapping = {
        'ФИО преподавателя': 'ФИО преподавателя',
        'Средняя посещаемость': 'Средняя посещаемость',
        'Всего пар': 'Всего пар',
        'Всего групп': 'Всего групп',
        'Unnamed: 5': 'Проверено',
        'Unnamed: 6': 'Выдано',
        'Тема урока': 'Тема урока'
    }

    # Переименовываем столбцы по маппингу
    for col in df.columns:
        for key, value in column_mapping.items():
            if key in col:
                df.rename(columns={col: value}, inplace=True)

    return df

# Проверка корректности темы урока
def check_lesson_topic(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        df = normalize_columns(df)

        # Ищем столбцы, содержащие "Тема" или "Тема урока"
        lesson_topic_columns = [col for col in df.columns if 'Тема' in col]

        if not lesson_topic_columns:
            return "Ошибка: В файле отсутствует столбец, содержащий 'Тема' или 'Тема урока'."

        # Регулярное выражение для проверки формата "Урок №. Тема:"
        topic_pattern = r"^Урок \d+\. .+"

        # Проверяем корректность тем уроков
        incorrect_topics = df[~df[lesson_topic_columns[0]].str.match(topic_pattern, na=False)]

        messages = [
            f"Некорректная тема урока: {row[lesson_topic_columns[0]]}. Пожалуйста, исправьте, формат должен быть: 'Урок №. Тема:'"
            for _, row in incorrect_topics.iterrows()
        ]
        return messages or ["Все темы уроков заполнены корректно."]
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Функция для расчета процента выполнения домашних заданий
def calculate_homework_status(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        df = normalize_columns(df)

        # Проверяем наличие нужных столбцов
        if 'ФИО преподавателя' not in df.columns or 'Проверено' not in df.columns:
            return f"Ошибка: В файле отсутствуют нужные столбцы. Доступные столбцы: {df.columns.tolist()}"

        df['Проверено'] = pd.to_numeric(df['Проверено'], errors='coerce').fillna(0)
        max_homework = df['Проверено'].max()
        df['Процент проверки'] = (df['Проверено'] / max_homework) * 100
        low_check = df[(df['Процент проверки'] < 75) & (df['ФИО преподавателя'].notna())]

        alerts = [
            f"Уважаемый(ая) {row['ФИО преподавателя']}, ваш процент проверки домашних заданий составляет {row['Процент проверки']:.2f}%. Пожалуйста, уделите внимание проверке."
            for _, row in low_check.iterrows()
        ]
        return alerts or ["Все преподаватели проверили домашние задания в достаточном объёме."]
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Функция для расчета процента выданных домашних заданий
def calculate_homework_given(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        df = normalize_columns(df)

        if 'ФИО преподавателя' not in df.columns or 'Выдано' not in df.columns:
            return f"Ошибка: В файле отсутствуют нужные столбцы для расчета выданного д/з. Доступные столбцы: {df.columns.tolist()}"

        df['Выдано'] = pd.to_numeric(df['Выдано'], errors='coerce').fillna(0)
        max_given = df['Выдано'].max()
        df['Процент выданного'] = (df['Выдано'] / max_given) * 100
        low_given = df[(df['Процент выданного'] < 70) & (df['ФИО преподавателя'].notna())]

        alerts = [
            f"Уважаемый(ая) {row['ФИО преподавателя']}, ваш процент выданного домашнего задания составляет {row['Процент выданного']:.2f}%. Пожалуйста, обратите внимание."
            for _, row in low_given.iterrows()
        ]
        return alerts or ["Все преподаватели выдали достаточно домашнего задания."]
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Функция для проверки посещаемости преподавателей
def check_attendance(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        df = normalize_columns(df)

        if not all(
                col in df.columns for col in ['ФИО преподавателя', 'Средняя посещаемость', 'Всего пар', 'Всего групп']):
            return f"Ошибка: В файле отсутствуют нужные столбцы. Доступные столбцы: {df.columns.tolist()}"

        df[['Средняя посещаемость', 'Всего пар', 'Всего групп']] = df[[
            'Средняя посещаемость', 'Всего пар', 'Всего групп']].apply(pd.to_numeric, errors='coerce').fillna(0)

        low_attendance = df[(df['Средняя посещаемость'] < 65) & (df['Всего пар'] > 0) & (df['Всего групп'] > 0)]
        messages = [
            f"У преподавателя {row['ФИО преподавателя']} средняя посещаемость {row['Средняя посещаемость']}%. Всего пар: {row['Всего пар']}, всего групп: {row['Всего групп']}. Пожалуйста, обратите внимание."
            for _, row in low_attendance.iterrows()
        ]
        return messages or ["Посещаемость всех преподавателей соответствует норме."]
    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Отправь мне Excel файл с данными для анализа.")

# Обработчик получения файлов
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)

        # Проверяем статус домашних заданий (проверено)
        result_check = calculate_homework_status(BytesIO(file))
        send_result(message, result_check, "Результаты проверки домашних заданий")

        # Проверяем статус домашних заданий (выдано)
        result_given = calculate_homework_given(BytesIO(file))
        send_result(message, result_given, "Результаты выданного домашнего задания")

        # Проверяем правильность темы уроков
        result_topics = check_lesson_topic(BytesIO(file))
        send_result(message, result_topics, "Результаты проверки тем уроков")

        # Проверяем посещаемость
        result_attendance = check_attendance(BytesIO(file))
        send_result(message, result_attendance, "Результаты проверки посещаемости")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")

# Функция отправки результата
def send_result(message, result, caption):
    if isinstance(result, str):
        bot.send_message(message.chat.id, result)
    else:
        with BytesIO() as output:
            output.write("\n".join(result).encode('utf-8'))
            output.seek(0)
            bot.send_document(message.chat.id, output, caption=caption)

# Запуск бота
if __name__ == '__main__':
    bot.polling(none_stop=True)
