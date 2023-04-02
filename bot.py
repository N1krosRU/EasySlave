import telebot  # pip install pytelegrambotapi
from main import xlsx_to_xml
import os
import dotenv

dotenv.load_dotenv()
BOT_API_KEY = os.getenv("BOT_API_KEY")  # Считываем ключ из переменной окружения или файла .env если он имеется
bot = telebot.TeleBot(BOT_API_KEY)
help_message = """Приветствую тебя, я Бот для создания шаблонов Modbus Slave устройств в программе Эмулятор Modbus устройств
(https://www.ardsoft.ru/mEmulator.html)

Как со мной работать:
Заполни эксель таблицу из примера и отправь её мне.
Взамен я пришлю тебе файл шаблон для импорта в программу Эмулятор Modbus TCP/RTU устройств.

1) Открой программу Эмулятор Modbus TCP/RTU устройств
2) Сконфигурируй сервер по своему выбору
3) ЛКМ по сконфигурированному серверу
4) Правка > Добавить устройство из шаблона
5) Выбрать полученный файл .tmpl

Список команд:
/example - получить пример .xlsx файла
/help - повторно вывести данное информационное сообщение

По всем вопросам и предложениям: @nenakhovaa
"""
if not os.path.exists("_downloads/"):
    os.mkdir("_downloads/")
path_to_example_doc = os.path.join(os.path.dirname(__file__), "_example/example.xlsx")
path_to_downloads = os.path.join(os.path.dirname(__file__), "_downloads/")


# Message_handlers
@bot.message_handler(commands=['start', 'help'])
def command_start(message):
    bot.send_message(message.from_user.id, help_message)


@bot.message_handler(commands=['example'])
def command_example(message):
    # doc = open(path_to_example_doc, 'rb')
    # bot.send_document(message.from_user.id, doc)
    with open(path_to_example_doc, 'rb') as doc:
        bot.send_document(message.from_user.id, doc)


@bot.message_handler(content_types=['document'])
def get_document(message):
    if message.document.mime_type != "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet":
        bot.send_message(message.from_user.id, "Неверное расширение файла, ожидается .xlsx!")
    else:
        file_name = message.document.file_name
        suf_file_name = file_name.split(".")[1]
        pref_file_name = file_name.split(".")[0]
        file_id = bot.get_file(message.document.file_id)
        downloaded_file = bot.download_file(file_id.file_path)
        # Save xlsx File
        with open(f"{path_to_downloads}{file_name}", "wb") as file_obj:
            file_obj.write(downloaded_file)
        # Create xml
        if xlsx_to_xml(f"{path_to_downloads}{file_name}"):
            # Send xml File
            with open(f"{path_to_downloads}{pref_file_name}.tmpl", 'rb') as send_file:
                bot.send_document(message.from_user.id, send_file)
        else:
            bot.send_message(message.from_user.id, "Ошибка, проверьте правильность заполнения таблицы!")


@bot.message_handler(content_types=["text"])
def weather_place(message):
    bot.send_message(message.from_user.id, 'Я не обрабатываю текстовые сообщения')


bot.infinity_polling()
# bot.polling(none_stop=True, interval=0)
