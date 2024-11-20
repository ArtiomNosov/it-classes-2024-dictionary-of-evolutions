import time
import telebot
import requests
from bs4 import BeautifulSoup
import sqlite3
import nltk
from nltk.stem import WordNetLemmatizer
from telebot.apihelper import ApiTelegramException

# Загрузка ресурса для лемматизации
nltk.download('wordnet')
lemmatizer = WordNetLemmatizer()

# Настройки бота и базы данных
API_TOKEN = '7932886518:AAFblUZTowAFxOiZ42dqEBEPmI4Za3p-7-0'
bot = telebot.TeleBot(API_TOKEN)
conn = sqlite3.connect("items.db", check_same_thread=False)
cursor = conn.cursor()

cursor.execute("""
    CREATE TABLE IF NOT EXISTS items (
        id INTEGER PRIMARY KEY,
        century INTEGER,
        item_name TEXT,
        image BLOB
    )
""")
conn.commit()

# Глобальная структура для отслеживания данных пользователя
user_data = {}

# Функция безопасной отправки сообщения
def safe_send_message(user_id, text):
    while True:
        try:
            bot.send_message(user_id, text)
            break  # Если успешно, выходим из цикла
        except ApiTelegramException as e:
            if "retry after" in str(e):
                wait_time = int(str(e).split("retry after")[1].split()[0])  # Извлекаем время ожидания
                time.sleep(wait_time + 1)  # Ждем указанный интервал
            else:
                raise  # Повторно бросаем исключение, если ошибка не 429

# Функция поиска изображений
def search_images(query):
    headers = {"User-Agent": "Mozilla/5.0"}
    url = f"https://www.google.com/search?tbm=isch&q={query}"
    response = requests.get(url, headers=headers)
    soup = BeautifulSoup(response.text, 'html.parser')
    images = [img['src'] for img in soup.find_all('img') if 'src' in img.attrs]
    return images[:20]  # Возвращаем до 20 результатов

# Функция для загрузки изображения и преобразования в бинарный формат
def download_image_as_blob(url):
    try:
        response = requests.get(url)
        response.raise_for_status()
        return response.content
    except requests.RequestException:
        return None

# Обработчик команды /start
@bot.message_handler(commands=['start'])
def start_message(message):
    user_id = message.chat.id
    user_data[user_id] = {'item_name': None, 'century': 1, 'images': []}
    safe_send_message(user_id, "Привет! Напишите название предмета.")

# Обработка названия предмета
@bot.message_handler(func=lambda message: user_data.get(message.chat.id, {}).get('item_name') is None)
def get_item_name(message):
    user_id = message.chat.id
    item_name = message.text.strip()

    if item_name == ".":
        safe_send_message(user_id, "Пропускаем предмет. Введите следующий предмет или нажмите /start для перезапуска.")
        return

    lemmatized_item = lemmatizer.lemmatize(item_name.lower())
    user_data[user_id]['item_name'] = lemmatized_item
    century = user_data[user_id]['century']
    query = f"{lemmatized_item} в {century} веке фото"
    images = search_images(query)

    user_data[user_id]['images'] = images
    if not images:
        safe_send_message(user_id, "Не удалось найти изображения. Введите следующий предмет.")
        user_data[user_id] = {'item_name': None, 'century': 1, 'images': []}
        return

    for i, img in enumerate(images):
        safe_send_message(user_id, f"Фото #{i + 1}:\n{img}")

    safe_send_message(user_id, "Выберите номер фото, которое подходит, или отправьте '.' для пропуска.")

# Обработка выбора изображения или пропуска
@bot.message_handler(func=lambda message: user_data.get(message.chat.id, {}).get('item_name') is not None)
def choose_image(message):
    user_id = message.chat.id
    choice = message.text.strip()
    data = user_data[user_id]

    if choice == ".":
        safe_send_message(user_id, f"Пропущено для {data['century']} века.")
    else:
        try:
            choice_index = int(choice) - 1
            selected_image_url = data['images'][choice_index]
            image_blob = download_image_as_blob(selected_image_url)

            if image_blob:
                cursor.execute("INSERT INTO items (century, item_name, image) VALUES (?, ?, ?)",
                               (data['century'], data['item_name'], image_blob))
                conn.commit()
                safe_send_message(user_id, f"Сохранено: {data['item_name']} для {data['century']} века.")
            else:
                safe_send_message(user_id, "Не удалось загрузить изображение. Попробуйте снова.")
                return
        except (IndexError, ValueError):
            safe_send_message(user_id, "Некорректный ввод. Попробуйте снова.")
            return

    data['century'] += 1
    if data['century'] > 21:
        safe_send_message(user_id, "Все века обработаны. Введите следующий предмет или нажмите /start.")
        user_data[user_id] = {'item_name': None, 'century': 1, 'images': []}
    else:
        query = f"{data['item_name']} в {data['century']} веке фото"
        images = search_images(query)
        data['images'] = images
        if not images:
            safe_send_message(user_id, f"Не удалось найти изображения для {data['century']} века. Пропускаем.")
            return choose_image(message)

        for i, img in enumerate(images):
            safe_send_message(user_id, f"Фото #{i + 1}:\n{img}")

        safe_send_message(user_id, "Выберите номер фото, которое подходит, или отправьте '.' для пропуска.")

# Запуск бота
bot.polling()
