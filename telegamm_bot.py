import telebot
import sqlite3
from telebot.types import Message
from rembg import remove
from PIL import Image
from io import BytesIO

TOKEN = '7847326259:AAEF9lmyqgYT2rwwtxpkwMuR1dJQV_IdMhA'
bot = telebot.TeleBot(TOKEN)

db_path = '../history_/items.db'
MAX_SIZE = (200, 200)  # Set maximum width and height in pixels


def init_db():
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute('''CREATE TABLE IF NOT EXISTS items (
                        id INTEGER PRIMARY KEY AUTOINCREMENT,
                        century TEXT,
                        item_name TEXT,
                        image BLOB
                    )''')
    conn.commit()
    conn.close()


def save_item(century, item_name, image_data):
    conn = sqlite3.connect(db_path)
    cursor = conn.cursor()
    cursor.execute("INSERT INTO items (century, item_name, image) VALUES (?, ?, ?)",
                   (century, item_name, image_data))
    conn.commit()
    conn.close()


@bot.message_handler(commands=['start'])
def start_message(message: Message):
    bot.send_message(message.chat.id, "Привет! Отправь мне изображение с подписью в формате: 'век предмет'.")


@bot.message_handler(content_types=['photo'])
def handle_image(message: Message):
    if message.caption:
        try:
            century, item_name = message.caption.split(" ", 1)
        except ValueError:
            bot.reply_to(message, "Пожалуйста, используйте формат подписи: 'век предмет'.")
            return

        file_info = bot.get_file(message.photo[-1].file_id)
        downloaded_file = bot.download_file(file_info.file_path)

        # Remove background and get bytes output
        image_data = downloaded_file

        # Open the processed image with Pillow
        image = Image.open(BytesIO(image_data))

        # Resize image if larger than MAX_SIZE
        if image.size[0] > MAX_SIZE[0] or image.size[1] > MAX_SIZE[1]:
            image.thumbnail(MAX_SIZE)  # Proportionally resize to fit within MAX_SIZE

        # Convert the image back to bytes for database storage
        output_buffer = BytesIO()
        image.save(output_buffer, format="PNG")
        save_item(century, item_name, output_buffer.getvalue())

        bot.reply_to(message, f"Данные сохранены: {century} {item_name}")
    else:
        bot.reply_to(message, "Пожалуйста, добавьте подпись к изображению в формате 'век предмет'.")


init_db()
bot.polling(none_stop=True)
