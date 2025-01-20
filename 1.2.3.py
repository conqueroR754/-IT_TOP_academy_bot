import telebot
import pandas as pd
from io import BytesIO

# Ваш токен для бота
TOKEN = 'YOUR_TELEGRAM_API_TOKEN'

bot = telebot.TeleBot('7785888241:AAGBdDryyw1MiQFlfwwChByROxmWZP-4jm8')
# Функция для расчета процента проверки
def calculate_homework_status(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        required_columns = ['ФИО преподавателя', 'Unnamed: 5']
        if not all(col in df.columns for col in required_columns):
            return f"Ошибка: В файле отсутствуют нужные столбцы. Доступные столбцы: {df.columns.tolist()}"

        df.rename(columns={'Unnamed: 5': 'Проверено'}, inplace=True)
        df['Проверено'] = pd.to_numeric(df['Проверено'], errors='coerce').fillna(0)
        max_homework = df['Проверено'].max()
        df['Процент проверки'] = (df['Проверено'] / max_homework) * 100
        low_check = df[(df['Процент проверки'] < 75) & (df['ФИО преподавателя'].notna())]

        alerts = []
        for _, row in low_check.iterrows():
            alerts.append(f"Уважаемый(ая) {row['ФИО преподавателя']}, ваш процент проверки домашних заданий составляет {row['Процент проверки']:.2f}%. Пожалуйста, уделите внимание проверке.")

        return alerts

    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Функция для расчета процента выданного д/з
def calculate_homework_given(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        required_columns = ['ФИО преподавателя', 'Unnamed: 6']
        if not all(col in df.columns for col in required_columns):
            return f"Ошибка: В файле отсутствуют нужные столбцы для расчета выданного д/з. Доступные столбцы: {df.columns.tolist()}"

        df.rename(columns={'Unnamed: 6': 'Выдано'}, inplace=True)
        df['Выдано'] = pd.to_numeric(df['Выдано'], errors='coerce').fillna(0)
        max_given = df['Выдано'].max()
        df['Процент выданного'] = (df['Выдано'] / max_given) * 100
        low_given = df[(df['Процент выданного'] < 70) & (df['ФИО преподавателя'].notna())]

        alerts = []
        for _, row in low_given.iterrows():
            alerts.append(f"Уважаемый(ая) {row['ФИО преподавателя']}, ваш процент выданного домашнего задания составляет {row['Процент выданного']:.2f}%. Пожалуйста, обратите внимание.")

        return alerts

    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Проверка корректности темы урока
def check_lesson_topic(file):
    try:
        df = pd.read_excel(file, sheet_name=0)
        if 'Тема урока' not in df.columns:
            return "Ошибка: В файле отсутствует столбец 'Тема урока'."

        incorrect_topics = df[~df['Тема урока'].str.startswith('Урок ', na=False)]

        if incorrect_topics.empty:
            return "Все темы уроков заполнены корректно."

        messages = []
        for _, row in incorrect_topics.iterrows():
            messages.append(f"Некорректная тема урока: {row['Тема урока']}. Пожалуйста, исправьте.")

        return messages

    except Exception as e:
        return f"Ошибка при обработке файла: {e}"

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Отправь мне Excel файл с данными, и я посчитаю проценты проверенных и выданных заданий. Если у педагогов процент ниже порога, они получат уведомление. Также могу проверить корректность заполнения темы урока.")

# Обработчик получения файлов
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)

        # Проверяем статус домашних заданий (проверено)
        result_check = calculate_homework_status(BytesIO(file))

        if isinstance(result_check, str):
            bot.send_message(message.chat.id, result_check)
        else:
            with BytesIO() as output:
                output.write("\n".join(result_check).encode('utf-8'))
                output.seek(0)
                bot.send_document(message.chat.id, output, caption="Результаты проверки домашних заданий")

        # Проверяем статус домашних заданий (выдано)
        result_given = calculate_homework_given(BytesIO(file))

        if isinstance(result_given, str):
            bot.send_message(message.chat.id, result_given)
        else:
            with BytesIO() as output:
                output.write("\n".join(result_given).encode('utf-8'))
                output.seek(0)
                bot.send_document(message.chat.id, output, caption="Результаты выданного домашнего задания")

        # Проверяем корректность темы уроков
        result_topics = check_lesson_topic(BytesIO(file))

        if isinstance(result_topics, str):
            bot.send_message(message.chat.id, result_topics)
        else:
            with BytesIO() as output:
                output.write("\n".join(result_topics).encode('utf-8'))
                output.seek(0)
                bot.send_document(message.chat.id, output, caption="Результаты проверки тем уроков")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")

# Запуск бота
if __name__ == '__main__':
    bot.polling(none_stop=True)
