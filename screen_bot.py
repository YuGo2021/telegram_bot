import telebot
import os
from os import getenv
from selenium import webdriver
from telebot import types
import random
from sys import exit
import time
import openpyxl
from datetime import datetime
import os

homepath = os.getenv('USERPROFILE')
#print(homepath)  # C:\Users\MyUser

chrome_path = os.path.normpath(homepath + '/AppData/Local/Google/Chrome/User Data/Defaul')
#print(chrome_path)  # C:\Users\MyUser\Music



bot_token = getenv("BOT_TOKEN")
#bot_token = "########################################"
if not bot_token:
    exit("Error: no token provided")

bot = telebot.TeleBot(bot_token, threaded = False)


#настраиваем браузер для корректной работы в headless режиме

options = webdriver.ChromeOptions()
#options.add_argument("user-data-dir=C:\\Users\\yuri.golubev\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
#options.add_argument("user-data-dir=C:\\Users\\golub\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
options.add_argument(f"user-data-dir={chrome_path}")

options.add_argument('--headless')
options.add_argument('--disable-gpu')
options.add_argument('--disable-dev-shm-usage')
options.add_argument('--no-sandbox')

options1 = webdriver.ChromeOptions()
#options1.add_argument("user-data-dir=C:\\Users\\yuri.golubev\\AppData\\Local\\Google\\Chrome\\User Data\\Default")
options1.add_argument(f"user-data-dir={chrome_path}")


admin = [112533702]
users = {}

#users.append(admin[0])
wb_users = openpyxl.load_workbook("Пользователи.xlsx")
sheet_users = wb_users.active
for i in range(2, sheet_users.max_row + 1):
    users.update({(sheet_users.cell(row=i, column=1)).value: {"name": (sheet_users.cell(row=i, column=2)).value, "status": (sheet_users.cell(row=i, column=3)).value}})

users.update({admin[0]:{"name": "Golubev", "status": 1}})

# only used for console output now
def listener(messages):
    """
    When new messages arrive TeleBot will call this function.
    """
    for m in messages:
        if m.content_type == 'text':
            # print the sent message to the console
            print(str(m.chat.first_name) + " [" + str(m.chat.id) + "]: " + m.text)

@bot.message_handler(func=lambda message: message.chat.id not in users)
def some(message):
    # wb_users = openpyxl.load_workbook("Пользователи.xlsx")
    # sheet_users = wb_users.active
    # for i in range(1, sheet_users.max_row + 1):
    #     if sheet_users.cell(row=i, column=1) not in users:
    #         users.append(sheet_users.cell(row=i, column=1))


    users.update({message.chat.id: {"name": message.chat.first_name, "status": 2}})
    for user in users:
        print(user, users[user]["name"], users[user]["status"])
    line = 0
    for user in users:
        line += 1
        sheet_users.cell(row=line, column=1).value = user
        sheet_users.cell(row=line, column=2).value = users[user]["name"]
        sheet_users.cell(row=line, column=3).value = users[user]["status"]
    wb_users.save("Пользователи.xlsx")
    bot.send_message(message.chat.id, 'Это закртый бот. Запросите доступ у администратора')
    bot.send_message(admin[0], 'Есть новые пользователи для подключения. Нажмите /start')
    #
    # markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
    # item0 = types.KeyboardButton("Запросить доступ")
    # markup.add(item0)
    # bot.send_message(message.chat.id, 'Запросить доступ?', reply_markup=markup)
#    bot.send_message(admin[0], 'Запросить доступ', reply_markup=markup)

# Команда start
@bot.message_handler(commands=["start"])
def start(m, res=False):
    # обновляем спсиок users
    # wb_users = openpyxl.load_workbook("Пользователи.xlsx")
    # sheet_users = wb_users.active
    # for i in range(1, sheet_users.max_row + 1):
    #     if sheet_users.cell(row=i, column=1) not in users:
    #         users.append(sheet_users.cell(row=i, column=1))

    if m.chat.id == admin[0]:
        print(users)
        for user in users:
            if users[user]["status"] == 2:
                # markup = types.ReplyKeyboardMarkup(resize_keyboard=True)
                # item_yes = types.KeyboardButton(f"AddUser:{user}")
                # item_no = types.KeyboardButton(f"Decline:{user}")
                # markup.row(item_yes, item_no)
                # bot.send_message(m.chat.id, f'{users[user]["name"]} [{user}] запрашивает доступ к боту', reply_markup=markup)
                keyboard = telebot.types.InlineKeyboardMarkup()
                button1 = telebot.types.InlineKeyboardButton(text="Предоставить доступ", callback_data=f"AddUser:{user}")
                button2 = telebot.types.InlineKeyboardButton(text="Отказать в доуступе", callback_data=f"Decline:{user}")
                keyboard.row(button1, button2)
                bot.send_message(m.chat.id, f'{users[user]["name"]} [{user}] запрашивает доступ к боту', reply_markup=keyboard)
    if m.chat.id in users and users[m.chat.id]["status"] == 1:
        #bot.send_message(m.chat.id, "Добро пожаловать в бот! Для начала работы нажмите /start")

        # Добавляем кнопки
        markup1 = types.ReplyKeyboardMarkup(resize_keyboard=True)
        item1 = types.KeyboardButton("Общий борд КЦ")
        item2 = types.KeyboardButton("Врачи")
        item3 = types.KeyboardButton("ВО")
        item4 = types.KeyboardButton("ОО")
        # markup.add(item1)
        # markup.add(item2)
        # markup.add(item3)
        # markup.add(item4)
        markup1.row(item1, item2)
        markup1.row(item3, item4)
        bot.send_message(m.chat.id, 'Для получения нужного скриншота нажми на кнопку (время загрузки около 10 сек)',  reply_markup=markup1)
    elif m.chat.id in users and users[m.chat.id]["status"] == 2:
        bot.send_message(m.chat.id, 'Ваш запрос на доступ к боту на рассмотрении. Для ускорения обратитесь к администратору')


# Получение сообщений от юзера
@bot.message_handler(content_types=["text"])
def handle_text(message):
    # wb_users = openpyxl.load_workbook("Пользователи.xlsx")
    # sheet_users = wb_users.active
    # for i in range(1, sheet_users.max_row + 1):
    #     if sheet_users.cell(row=i, column=1) not in users:
    #         users.append(sheet_users.cell(row=i, column=1))
    if message.chat.id in users and users[message.chat.id]["status"] == 1:

        # отправляем скрин борда по запросу
        if message.text.strip() == 'Общий борд КЦ' :
            #answer = random.choice(facts)
            uid = message.chat.id
            driver = webdriver.Chrome(chrome_options=options)
            driver.set_window_size(1920, 1080)
            now = datetime.now()
            if 'send_time1' not in globals():
                global send_time1
                send_time1 = datetime(2017, 7, 18, 4, 52, 33, 51204)
            delta = now - send_time1
            if delta.seconds > 60:
                global photo_path1
                photo_path1 = "Image\\" + str(random.randint(10000000, 99999999)) + '.png'

                driver.get("https://grafana2.gemotest.ru:3000/d/i_tq1wY7z/glavnaia-2?orgId=1&refresh=5m")
                #driver.get("https://mail.ru/")
                time.sleep(10)
                driver.save_screenshot(photo_path1)
                #global send_time1
                send_time1 = datetime.now()

            bot.send_photo(uid, photo=open(photo_path1, 'rb'))

            #driver.quit()
            #os.remove(photo_path)

        elif message.text.strip() == 'Врачи':
            uid = message.chat.id
            driver = webdriver.Chrome(chrome_options=options)
            driver.set_window_size(1920, 1080)
            now = datetime.now()
            if 'send_time2' not in globals():
                global send_time2
                send_time2 = datetime(2017, 7, 18, 4, 52, 33, 51204)

            delta = now - send_time2
            if delta.seconds > 60:
                global photo_path2
                photo_path2 = "Image\\" + str(random.randint(10000000, 99999999)) + '.png'

                driver.get("https://grafana2.gemotest.ru:3000/d/WCQ6wTsnz/vrachi?orgId=1&refresh=1m")
                #driver.get("https://yandex.ru/")
                time.sleep(10)
                driver.save_screenshot(photo_path2)
                send_time2 = datetime.now()
            bot.send_photo(uid, photo=open(photo_path2, 'rb'))
            #driver.quit()
            #os.remove(photo_path)
                #answer = random.choice(thinks)
        elif message.text.strip() == 'ВО':
            uid = message.chat.id
            driver = webdriver.Chrome(chrome_options=options)
            driver.set_window_size(1920, 1080)
            now = datetime.now()
            if 'send_time3' not in globals():
                global send_time3
                send_time3 = datetime(2017, 7, 18, 4, 52, 33, 51204)
            delta = now - send_time3
            if delta.seconds > 60:
                global photo_path3
                photo_path3 = "Image\\" + str(random.randint(10000000, 99999999)) + '.png'

                driver.get("https://grafana2.gemotest.ru:3000/d/n_7NiLrnk/bvo?orgId=1&refresh=15m")
                #driver.get("https://yandex.ru/")
                time.sleep(10)
                driver.save_screenshot(photo_path3)
                send_time3 = datetime.now()
            bot.send_photo(uid, photo=open(photo_path3, 'rb'))
            #driver.quit()
        elif message.text.strip() == 'ОО':
            uid = message.chat.id
            driver = webdriver.Chrome(chrome_options=options)
            driver.set_window_size(1920, 1080)
            now = datetime.now()
            if 'send_time4' not in globals():
                global send_time4
                send_time4 = datetime(2017, 7, 18, 4, 52, 33, 51204)
            delta = now - send_time4
            if delta.seconds > 60:
                global photo_path4
                photo_path4 = "Image\\" + str(random.randint(10000000, 99999999)) + '.png'

                driver.get("https://grafana2.gemotest.ru:3000/d/-zNr5Zq7z/oo?orgId=1&refresh=1m")
                #driver.get("https://yandex.ru/")
                time.sleep(10)
                driver.save_screenshot(photo_path4)
                send_time4 = datetime.now()
            bot.send_photo(uid, photo=open(photo_path4, 'rb'))
            #driver.quit()
    if message.chat.id == admin[0]:
        if message.text.strip() == 'отладка':
            uid = message.chat.id
            photo_path = "Image\\" + str(random.randint(10000000, 99999999)) + '.png'
            driver = webdriver.Chrome(chrome_options=options1)
            driver.set_window_size(1920, 1080)
            driver.get("https://grafana2.gemotest.ru:3000/d/i_tq1wY7z/glavnaia-2?orgId=1&refresh=5m")
            #driver.get("https://yandex.ru/")
            time.sleep(30)
            driver.save_screenshot(photo_path)
            bot.send_photo(uid, photo=open(photo_path, 'rb'))
        # if message.text.strip()[:7] == 'AddUser:':
        #     print(message.text.strip[8:])
        #     if message.chat.id == admin[0]:
        #         for user in users:
        #             if user == message.text.strip[8:]:
        #                 bot.send_message(message.chat.id, 'Пользовтель уже добавлен')
        #                 break
        #         else:
        #             users[message.text.strip[8:]]["status"] = 1
        #             for user in users:
        #                 print(user, users[user]["name"], users[user]["status"])
        #             line = 0
        #             for i in users:
        #                 line += 1
        #                 sheet_users.cell(row=line, column=1).value = i
        #                 sheet_users.cell(row=line, column=2).value = users[i]["name"]
        #                 sheet_users.cell(row=line, column=3).value = users[i]["status"]
        #             wb_users.save("Пользователи.xlsx")
        #             bot.send_message(message.chat.id, 'Пользовтель добавлен')
        #             bot.send_message(message.text.strip[8:], 'Вам предоставлен доступ к боту')
    # if message.chat.id not in users:
    #     if message.text.strip() == 'Запросить доступ':
    #
    #         uid = message.chat.id
    #         users.update({uid:{"name":message.chat.first_name, "status":2}})
    #         for user in users:
    #             print(user, user["name"], user["status"])
    #         line = 0
    #         for user in users:
    #             line += 1
    #             sheet_users.cell(row=line, column=1).value = user
    #             sheet_users.cell(row=line, column=2).value = user["name"]
    #             sheet_users.cell(row=line, column=3).value = user["status"]
    #         wb_users.save("Пользователи.xlsx")
    #
    #
    #
    #         bot.send_message(admin[0], f'{message.chat.first_name} [{uid}] запрашивает доступ к боту')
    #         markup2 = types.ReplyKeyboardMarkup(resize_keyboard=True)
    #         item5 = types.KeyboardButton(f'AddUser:{uid}')
    #         item6 = types.KeyboardButton(f"Decline:{uid}")
    #
    #         markup2.row(item5, item6)
    #         bot.send_message(admin[0],
    #                          f'{message.chat.first_name} [{uid}] запрашивает доступ к боту',  reply_markup=markup2)

@bot.callback_query_handler(func=lambda call: True)
def callback_function1(callback_obj):
    if callback_obj.data[:8] == "AddUser:":
        for user in users:
            if user == callback_obj.data[8:]  and users[user]["status"] == 1:
                #bot.send_message(message.chat.id, 'Пользовтель уже добавлен')
                bot.send_message(callback_obj.from_user.id, f"Пользовтель {callback_obj.data[8:]} уже добавлен")

                break
        else:
            users[int(callback_obj.data[8:])]["status"] = 1
            for user in users:
                print(user, users[user]["name"], users[user]["status"])
            line = 0
            for i in users:
                line += 1
                sheet_users.cell(row=line, column=1).value = i
                sheet_users.cell(row=line, column=2).value = users[i]["name"]
                sheet_users.cell(row=line, column=3).value = users[i]["status"]
            wb_users.save("Пользователи.xlsx")
            #bot.send_message(message.chat.id, 'Пользовтель добавлен')
            bot.send_message(callback_obj.from_user.id, f"Пользовтель {callback_obj.data[8:]} добавлен")
            bot.send_message(int(callback_obj.data[8:]), 'Вам предоставлен доступ к боту. Нажмите /start')

    elif callback_obj.data[:8] == "Decline:":
        users[int(callback_obj.data[8:])]["status"] = 0
        # print(users, users["status"])
        line = 0
        for i in users:
            line += 1
            sheet_users.cell(row=line, column=1).value = i
            sheet_users.cell(row=line, column=2).value = users[i]["name"]
            sheet_users.cell(row=line, column=3).value = users[i]["status"]
        wb_users.save("Пользователи.xlsx")
        bot.send_message(int(callback_obj.data[8:]), 'Вам отказано в достпуе к боту')

        bot.send_message(callback_obj.from_user.id, f"Пользователь {callback_obj.data[8:]} заблокирован")


    bot.answer_callback_query(callback_query_id=callback_obj.id)

#Handles all text messages that match the regular expression
# @bot.message_handler(regexp=r"((AddUser:)(\d*))")
# def handle_message(message):
#     print(message.text.strip[8:])
#     if message.chat.id == admin[0]:
#         for user in users:
#             if user == message.text.strip[8:]:
#                 bot.send_message(message.chat.id,'Пользовтель уже добавлен')
#                 break
#         else:
#             users[message.text.strip[8:]]["status"] = 1
#             for user in users:
#                 print(user, users[user]["name"], users[user]["status"])
#             line = 0
#             for i in users:
#                 line += 1
#                 sheet_users.cell(row=line, column=1).value = i
#                 sheet_users.cell(row=line, column=2).value = users[i]["name"]
#                 sheet_users.cell(row=line, column=3).value = users[i]["status"]
#             wb_users.save("Пользователи.xlsx")
#             bot.send_message(message.chat.id, 'Пользовтель добавлен')
#             bot.send_message(message.text.strip[8:], 'Вам предоставлен доступ к боту')

# @bot.message_handler(regexp=r"((Decline:)(\d*)")
# def handle_message(message):
#     if message.chat.id in admin:
#         users[message.text.strip[8:]]["status"] = 0
#         #print(users, users["status"])
#         line = 0
#         for i in users:
#             line += 1
#             sheet_users.cell(row=line, column=1).value = i
#             sheet_users.cell(row=line, column=2).value = users[i]["name"]
#             sheet_users.cell(row=line, column=3).value = users[i]["status"]
#         wb_users.save("Пользователи.xlsx")
#         bot.send_message(message.text.strip[8:], 'Вам отказано в достпуе к боту')

# Запускаем бота
bot.polling(none_stop=True, interval=0)


# отправка картинки пользователю
# uid = message.chat.id
# photo_path = str(uid) + '.png'
# driver = webdriver.Chrome(chrome_options = options)
# driver.set_window_size(1280, 720)
# driver.get(url)
# driver.save_screenshot(photo_path)
# bot.send_photo(uid, photo = open(photo_path, 'rb'))
# driver.quit()
# os.remove(photo_path)