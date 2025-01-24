import telebot
import pandas as pd
from io import BytesIO

# Токен вашего бота
TOKEN = '7785888241:AAGBdDryyw1MiQFlfwwChByROxmWZP-4jm8'
bot = telebot.TeleBot(TOKEN)

# Функция для анализа данных
def analyze_student_data(file):
    try:
        df = pd.read_excel(file, sheet_name=0)

        # Проверка наличия нужных столбцов
        required_columns = ['FIO', 'Homework', 'Classroom', 'Average score']
        if not all(col in df.columns for col in required_columns):
            return [f"Ошибка: В файле отсутствуют нужные столбцы: {', '.join(required_columns)}"]

        # Преобразуем нужные столбцы в числовой формат
        df[['Homework', 'Classroom', 'Average score']] = df[
            ['Homework', 'Classroom', 'Average score']].apply(pd.to_numeric, errors='coerce')

        # Фильтруем студентов, у которых средний балл меньше 3
        low_score_students = df[df['Average score'] < 3]

        # Формируем отчет
        if low_score_students.empty:
            return ["Все студенты имеют средний балл 3 и выше."]
        else:
            report_lines = [
                f"Студент: {row['FIO']}, Средний балл: {row['Average score']:.2f}, Домашняя работа: {row['Homework']}, Аудиторная работа: {row['Classroom']}"
                for _, row in low_score_students.iterrows()
            ]
            return report_lines
    except Exception as e:
        return [f"Ошибка при анализе файла: {e}"]

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start(message):
    bot.send_message(message.chat.id, "Привет! Отправь мне Excel файл с данными, и я проверю успеваемость студентов.")

# Обработчик получения файлов
@bot.message_handler(content_types=['document'])
def handle_document(message):
    try:
        # Получаем файл от пользователя
        file_id = message.document.file_id
        file_info = bot.get_file(file_id)
        file = bot.download_file(file_info.file_path)

        # Анализируем данные
        results = analyze_student_data(BytesIO(file))

        # Создаем текстовый файл с результатами
        with BytesIO() as output:
            output.write("\n".join(results).encode('utf-8'))
            output.seek(0)
            bot.send_document(message.chat.id, output, caption="Результаты анализа студентов")

    except Exception as e:
        bot.send_message(message.chat.id, f"Произошла ошибка: {e}")

# Запуск бота
if __name__ == '__main__':
    bot.polling(none_stop=True)
