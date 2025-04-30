import asyncio
import logging
import os
import signal
import sys
import time
import random
from datetime import datetime, timedelta
from pytz import timezone
from main import schedule_picker, pdf_picker
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup, ChatMemberUpdated, Bot
from telegram.constants import ParseMode
from telegram.ext import Application, CommandHandler, MessageHandler, CallbackQueryHandler, filters, CallbackContext, ContextTypes, ChatMemberHandler
import aiohttp


async def shutdown():
    print('Завершение работы бота... Сохранение Логов')
    sys.stderr.close()
    tasks = [t for t in asyncio.all_tasks() if t is not asyncio.current_task()]
    for task in tasks:
        task.cancel()
        try:
            await task
        except asyncio.CancelledError:
            pass

def signal_handler(sig, frame):
    asyncio.create_task(shutdown())
    sys.exit(0)

# Регистрация обработчика сигнала для SIGINT (Ctrl+C)
signal.signal(signal.SIGINT, signal_handler)


TOKEN = ''
CHAT_ID = '-4628489295'

async def start(update: Update, context: CallbackContext) -> None:
    print("start")
    if 'mode' not in context.user_data:
        context.user_data['mode'] = 'menu'
        await show_menu(update, context)

async def show_menu(update: Update, context: CallbackContext) -> None:
    print("menu")
    keyboard = [
        [InlineKeyboardButton("Подбор расписания", callback_data='mode1')],
        [InlineKeyboardButton("Создание pdf служебки", callback_data='mode2')],
    ]
    reply_markup = InlineKeyboardMarkup(keyboard)
    if update.message:
        await update.message.reply_text(text=u'Памятка, как пользоваться ботом разделена на части и представлена для каждого режима отдельно\n\n'
                                        'Для получения более подробной информации о принципах работы алгоритма введите /info', parse_mode=ParseMode.HTML, reply_markup=reply_markup)


    elif update.callback_query:
        query = update.callback_query
        await query.answer()
        await query.message.chat.send_message(text=u'Памятка, как пользоваться ботом разделена на части и представлена для каждого режима отдельно\n\n'
                                        'Для получения более подробной информации о принципах работы алгоритма введите /info',parse_mode=ParseMode.HTML , reply_markup=reply_markup)
    context.user_data['mode'] = 'menu'

async def mode1(update: Update, context: CallbackContext) -> None:
    print("mode1")
    query = update.callback_query
    await query.message.reply_text(text='Памятка, как пользоваться ботом: \n\n'
                                        'Выберите режим "Подбор расписания" и отправьте ему запрос в формате:\n\n'
                                        '<b>Фамилия Имя Цель_для_чего_нужна_аудитория день.месяц час_начала:минуты_начала час_конца:минуты_конца тип_аудитории</b>\n\n'
                                        'Каждый элемент запроса важен и должен быть отделен одним пробелом. '
                                        'Даты пишутся через точку, время через двоеточие. '
                                        'В запросе может быть неограниченное количество аудиторий, которые необходимо выписать. Каждая отдельная '
                                        'должна быть в вышеописанном формате. Каждая аудитория пишется с новой строки\n'
                                        'тип аудитории может быть:\n\n'
                                        '<b>any</b> - любая\n\n'
                                        '<b>any2</b> - любая во 2 корпусе\n\n'
                                        '<b>big</b> - большая (такими считаются 502, 513, 314, 105, 328, 307, 309, 324)\n\n'
                                        '<b>big2</b> -большая во 2 корпусе\n\n'
                                        '<b>105</b>(любой доступный номер аудитории (пока что работает только для 2 корпуса)) - если вы уверены в том, какая аудитория вам необходима, можно ввести ее номер\n\n'
                                        'В качестве результата работы программа вернет список запроса с пометками, какие аудитории удалось найти, а какие нет\n\n'
                                        'Пример:\n'
                                        '<b>Иванов Иван танцы 03.02 14:15 18:20 big2\n'
                                        'Петров Петр муз 04.02 16:20 21:00 105</b>\n\n'
                                        'После обработки запроса функция выведет запрос с пометками справа. Справа может быть записана <b>аудитория+корпус</b> '
                                        'или <b>no\u00A0free\u00A0room</b> если в расписании нет свободного места для этого запроса (В этом случае можно попытаться изменить '
                                        'временные рамки этого запроса или тип аудитории) или <b>no\u00A0such\u00A0day</b> в случае, если в расписании скачанном с портала '
                                        'нет нужного дня. Это может возникать если вы ввели несуществующий день по типу 32 января или если вы пытаетесь выписать'
                                        ' аудиторию более чем через месяц от сегодняшнего дня. <b>Wrong\u00A0postfix</b> - если ошибка в типе аудитории.\n\n'
                                        'Введите текст:',parse_mode=ParseMode.HTML ,)

    context.user_data['mode'] = 'mode1'

async def mode2(update: Update, context: CallbackContext) -> None:
    print("mode2")
    query = update.callback_query
    await query.message.reply_text(text='Для получения pdf служебки необходимо выбрать 2 режим и отправить ему правильный запрос. в ответ функция пришлет pdf служебку\n\n'
                                        'Если фунция прислала в ответ ошибку, вероятно не удалось найти аудитории для каждой строки запроса.\n'
                                        'Сгенерировать служебку возможно только если для всех запросов удалось найти аудиторию'
                                        'Введите текст:',parse_mode=ParseMode.HTML ,)

    context.user_data['mode'] = 'mode2'

async def info(update: Update, context: CallbackContext) -> None:
    print("info")
    keyboard = [[InlineKeyboardButton("Вернуться в меню", callback_data='menu')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text('БОЛЕЕ ПОДРОБНО О РАБОТЕ БОТА\n\n'
                                    'Боту для работы требуется четко структурированный запрос (ведутся работы по устранению этого недостатка)\n'
                                    'При получении запроса, вызывается соответствущая функция (pdf или schedule). '
                                    'Функция обращается к расписанию, полученному с портала в формате .xlsx. '
                                    'Сначала перебирается столбец A:A для того чтобы определить в каких строчках начинаются и заканчиваются дни'
                                    ' (расписание с портала не создает всегда 8 строк на день, их количество может варьироваться от 0 до 8 взависимости от того, какая пара в этот день последняя и первая)'
                                    ' Далее перебирается строка 1 для того чтобы записать координаты нужных нам аудиторий (раньше адреса указывались абсолбютно'
                                    ' но порталоделам вдруг захотелось все перемешать и с работой программы возникли проблемы) '
                                    'Далее на основании известных координат начала и конца дня (номера строк) и координат соответствующих аудиториям (буквы столбцов)'
                                    ' информация о расписании для каждой аудитории копируется из таблицы в массивы. Структура выгдялит так:\n'
                                    'hash_map с днями (кодовыми словами служит день и месяц. это будет удобно в дальнейшем)\n'
                                    'в каждой ее ячейке хранится кортеж с информацией (день недели, [массив])\n'
                                    'В массиве лежат кортежи, хранящие информацию о всех аудиториях в этот день\n'
                                    'Каждый кортеж хранит в себе номер аудитории, корпус, тип(big/small) и массив с данными о расписании этой конкретной аудитории в этот день\n'
                                    'Для корректной работы программы, требуется чтобы на всех свободных временных ячейках была записана пустая строка '
                                    'Но портал не вписывает в таблицу пары, если в это время не запланировано ничего '
                                    '',



                                    reply_markup=reply_markup)

async def text_received(update: Update, context: CallbackContext) -> None:

    user_text = update.message.text
    mode = context.user_data.get('mode')
    print("text_received: "+ user_text)
    if mode == 'mode1':
        msg = await update.message.reply_text('Вы в режиме 1. Ожидайте...')
        keyboard = [[InlineKeyboardButton("Вернуться в меню", callback_data='menu')]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        try:
            replay = await schedule_picker(user_text)  # Симуляция длительной работы
            print(replay)
            await msg.delete()
            await update.message.reply_text(f'Функция завершила работу:\n' + replay,
                        reply_markup=reply_markup)
        except Exception:
            await update.message.reply_text(f'Ошибка ввода',
                                            reply_markup=reply_markup)
            print("Exept: mode1")
    elif mode == 'mode2':
        msg = await update.message.reply_text('Вы в режиме 2. Ожидайте...')
        keyboard = [[InlineKeyboardButton("Вернуться в меню", callback_data='menu')]]
        reply_markup = InlineKeyboardMarkup(keyboard)
        try:
            replay, pdf_name = await pdf_picker(user_text)  # Симуляция длительной работы
            await msg.delete()
            await update.message.reply_text(f'Сообщение для постановщиков: \n'
                                            f''+replay,
                                        reply_markup=reply_markup)
            await context.bot.send_document(chat_id=update.message.chat_id, document=open(pdf_name, 'rb'))
        except Exception:
            await update.message.reply_text(f'Попробуйте в более стабильном режиме без pdf',
                                            reply_markup=reply_markup)
            print("Exept: mode2 no room")
    else:
        await update.message.reply_text('Пожалуйста, выберите режим в главном меню. Используйте команду /menu.')

async def menu_command(update: Update, context: CallbackContext) -> None:
    print("menu_command")
    await show_menu(update,context)

async def version(update: Update, context: CallbackContext)->None:
    print("version")
    keyboard = [[InlineKeyboardButton("Вернуться в меню", callback_data='menu')]]
    reply_markup = InlineKeyboardMarkup(keyboard)
    await update.message.reply_text(f'Текщая версия бота: v.0.1.4\n\n'
                                    f'Последние изменения:\n'
                                    f'-Упрощение обновления расписания (пока что не автоматизация)\n'
                                    f'-Создание системы автоматической проверки, онлайн ли бот\n'
                                    f'v.0.1.3:\n'
                                    f'-Добавление сообщения для постановщиков о том, какие аудитории выписаны\n'
                                    f'-Исправление форматирования сообщений\n'
                                    f'-Добавление логгирования ошибок\n'
                                    f'\n'
                                    f'v.0.1.2:\n'
                                    f'-Полная переработка основного алгоритма поиска в расписании всвязи со злыми порталоделами\n'
                                    f'-Исправление памяток\n'
                                    f'-Добавление объяснений о принципах работы бота\n'
                                    f'-Добавление информации о версии\nИзменение названий файлов служебок\n'
                                    f'-Добавление дня недели в служебку\n\n'
                                    f'Планы:\n'
                                    
                                    f'-Добавление возможности взаимодействия с гугл таблицами\n'
                                    f'-Обучение и добавление нейросети, которая позволит обрабатывать запросы в произвольной форме\n'
                                    f'-Создание системы автоматического обновления расписания',
                                            reply_markup=reply_markup)
    raise Exception("test_exept")

async def error_handler(update: object, context: ContextTypes.DEFAULT_TYPE) -> None:
    logging.exception("Error:")

async def send_message(session, token, chat_id, text, offset=None):
    try:
        url1 = f'https://api.telegram.org/bot{token}/sendMessage'
        payload = {
            'chat_id': chat_id,
            'text': text
        }
        async with session.post(url1, json=payload) as response:
            response.raise_for_status()  # Поднимает исключение для неверных запросов
            #return response
    except aiohttp.ClientError as e:
        print(f'Ошибка при отправке сообщения: {e}')
        return None

async def monitor_bot_loop(session):
    offset= None
    while True:
        try:
            # Отправляем тестовое сообщение основному боту
            test_message = [
    "Всё отлично.",
    "Всё в порядке.",
    "Никаких проблем.",
    "Всё хорошо.",
    "Всё нормально.",
    "Всё замечательно.",
    "Всё прекрасно.",
    "Всё чудесно.",
    "Всё великолепно.",
    "Всё гладко.",
    "Всё спокойно.",
    "Всё функционирует.",
    "Всё стабильно.",
    "Всё безупречно.",
    "Всё нормально идёт.",
    "Всё как надо.",
    "Всё чётко.",
    "Всё идёт по плану.",
    "Всё по расписанию.",
    "Всё по графику.",
    "Никаких забот.",
    "Никаких тревог.",
    "Никаких волнений.",
    "Никаких неполадок.",
    "Всё под контролем.",
    "Всё идёт как по маслу.",
    "Всё в норме.",
    "Всё как часики.",
    "Всё ровно.",
    "Всё окей.",
    "Всё ништяк.",
    "Всё тип-топ.",
    "Всё без изменений.",
    "Всё стабильно.",
    "Всё супер.",
    "Всё отлично работает.",
    "Всё как всегда.",
    "Всё как обычно.",
    "Никаких вопросов.",
    "Всё чики-пики.",
    "Всё как по маслу.",
    "Никаких осложнений.",
    "Всё тихо.",
    "Всё гладко проходит.",
    "Всё на высоте.",
    "Всё путём.",
    "Всё по старому.",
    "Всё на месте.",
    "Всё замечательно работает.",
    "Всё под контролем.",
    "Всё без сбоев.",
    "Всё ладно.",
    "Всё по фэншую.",
    "Всё на ура.",
    "Всё по правилам.",
    "Всё чудненько.",
    "Всё на уровне.",
    "Всё как задумано.",
    "Всё чётко работает.",
    "Всё без проблем.",
    "Всё спокойно и тихо.",
    "Всё гладко проходит.",
    "Всё на месте.",
    "Всё в рабочем состоянии.",
    "Всё чин-чинарём.",
    "Всё цветёт и пахнет.",
    "Всё пучком.",
    "Всё по плану.",
    "Всё как часы.",
    "Всё по стандарту.",
    "Всё в норме вещей.",
    "Всё как обычно.",
    "Всё благополучно.",
    "Всё чинно.",
    "Всё в ажуре.",
    "Всё комфортно.",
    "Всё корректно.",
    "Всё в шоколаде.",
    "Всё образцово.",
    "Всё классно.",
    "Всё утроено.",
    "Всё исправно.",
    "Всё без заморочек.",
    "Всё без волнений.",
    "Всё достойно.",
    "Всё без изменений.",
    "Всё как в аптеке.",
    "Всё на своём месте.",
    "Всё в своём темпе.",
    "Всё как по учебнику.",
    "Всё как по книжке.",
    "Всё по уставу.",
    "Всё по предписанию.",
    "Всё без перебоев.",
    "Всё работает.",
    "Всё функционирует.",
    "Всё в порядке вещей.",
    "Всё отлично функционирует."
]
            test_message1 = ['Все пиздец', 'Мы все умрем']
            ra = random.randint(0, 97)
            await send_message(session,TOKEN, CHAT_ID, test_message[ra], offset)
        except Exception as e:
            # Обработка ошибок
            print(f'Ошибка: {e}')

        # Интервал времени между проверками (1 минута)
        await asyncio.sleep(60)

async def run_polling_loop(application):
    print("Polling запущен, выполняем дополнительный код")
    await application.updater.start_polling()

async def main1() -> None:
    #table_update()

    print("Начало работы")
    application = Application.builder().token(TOKEN).build()
    application.add_handler(CommandHandler('start', start))
    application.add_handler(CommandHandler('version', version))
    application.add_handler(CommandHandler('menu', menu_command))
    application.add_handler(CommandHandler('info', info))
    application.add_handler(CallbackQueryHandler(show_menu, pattern='^menu$'))
    application.add_handler(CallbackQueryHandler(mode1, pattern='^mode1$'))
    application.add_handler(CallbackQueryHandler(mode2, pattern='^mode2$'))
    application.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, text_received))
    application.add_error_handler(error_handler)

    await application.initialize()
    await application.start()

    async with aiohttp.ClientSession() as session:
        await asyncio.gather(
            run_polling_loop(application),
            monitor_bot_loop(session)
        )
    await application.stop()
    #application.run_polling()





'''async def table_update():
    browser = await launch(headless=True)
    page = await browser.newPage()
    await asyncio.sleep(10)
    # Откройте сайт и выполните авторизацию
    await page.goto('https://portal.unn.ru/')
    await page.type('input[name="USER_LOGIN"]', 's23380293')
    await page.type('input[name="USER_PASSWORD"]', 'kodt2005')
    await asyncio.sleep(10)
    await page.click('button[type="submit"]')
    await asyncio.sleep(10)
    # Подождите, пока страница загрузится
    await page.waitForNavigation()
    await asyncio.sleep(10)
    # Перейдите на нужную страницу
    await page.goto('https://portal.unn.ru/ruz/main')
    await asyncio.sleep(10)
    # Подождите, пока кнопка станет доступной и нажмите на нее
    await page.waitForSelector('a.nav-link.ng-star-inserted')  # Ожидание появления ссылки с классом nav-link ng-star-inserted
    links = await page.querySelectorAll('a.nav-link.ng-star-inserted')
    await asyncio.sleep(10)
    # Найдите кнопку с текстом "Аудитория" и нажмите на нее
    for link in links:
        text = await page.evaluate('(link) => link.textContent', link)
        if text.strip() == 'Загруженность аудиторий':
            await link.click()
            break
    await asyncio.sleep(10)
    # Получите текущую дату и дату через 14 дней
    today = datetime.now().strftime('%d.%m.%Y')
    date_plus_14 = (datetime.now() + timedelta(days=14)).strftime('%d.%m.%Y')

    # Введите сегодняшнюю дату в поле с id="start"
    await page.waitForSelector('input[name="start"]')
    await page.type('input[name="start"]', today)
    await asyncio.sleep(10)
    # Введите дату через 14 дней в поле с id="end"
    await page.waitForSelector('input[name="end"]')
    await page.type('input[name="end"]', date_plus_14)
    await asyncio.sleep(10)
    await page.waitForSelector('input#autocomplete-building')
    await page.type('input#autocomplete-building', 'корпус')
    await asyncio.sleep(10)
    # Подождите, пока элемент с текстом "Корпус № 2" станет доступным и нажмите на него
    await page.waitForXPath('//div[text()="Корпус № 2"]')
    element = await page.xpath('//div[text()="Корпус № 2"]')
    await element[0].click()
    await asyncio.sleep(10)
    # Подождите, пока кнопка "Экспорт" станет доступной и нажмите на нее
    await page.waitForSelector('button.btn.btn-outline-secondary')
    await page.click('button.btn.btn-outline-secondary')

    # Подождите некоторое время, чтобы файл успел скачаться
    await asyncio.sleep(10)

    await browser.close()'''

log_dir = 'logs'
if not os.path.exists(log_dir):
    os.makedirs(log_dir)
    print("папка для логов создана")

log_filename = os.path.join(log_dir, 'LOG' + datetime.now().strftime(" %m.%d_%H-%M-%S") + ".log")
logging.basicConfig(level=logging.ERROR, filename=log_filename, filemode='w', format='%(asctime)s - %(levelname)s - %(message)s')
sys.stderr = open(log_filename, 'a')
if __name__ == '__main__':
    try:
        asyncio.run(main1())
    except KeyboardInterrupt:
        print("Получен сигнал прерывания (Ctrl+C). Завершение работы...")

