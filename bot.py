import os
import logging
from datetime import datetime
import re
from io import BytesIO
import signal
import sys

import telebot
from telebot import types
import gspread
from oauth2client.service_account import ServiceAccountCredentials
from PIL import Image, ImageDraw, ImageFont
import qrcode
from reportlab.lib.pagesizes import A4, landscape
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib.utils import ImageReader
from dotenv import load_dotenv
import yadisk

# Настройка корректного завершения
def signal_handler(sig, frame):
    logger.info("Bot shutdown gracefully")
    sys.exit(0)

signal.signal(signal.SIGINT, signal_handler)
signal.signal(signal.SIGTERM, signal_handler)

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[logging.FileHandler('bot.log'), logging.StreamHandler()]
)
logger = logging.getLogger(__name__)

# Загрузка переменных окружения
load_dotenv()
BOT_TOKEN = os.getenv('BOT_TOKEN')
SPREADSHEET_NAME = os.getenv('SPREADSHEET_NAME')
MASTER_ID = int(os.getenv('MASTER_ID'))
MASTER_PHONE = os.getenv('MASTER_PHONE')
YANDEX_DISK_TOKEN = os.getenv('YANDEX_DISK_TOKEN')
YANDEX_DISK_FOLDER = os.getenv('YANDEX_DISK_FOLDER')

# Инициализация бота
bot = telebot.TeleBot(BOT_TOKEN, threaded=False)

# Настройка Google Sheets
try:
    gs_scope = ['https://spreadsheets.google.com/feeds', 'https://www.googleapis.com/auth/drive']
    creds = ServiceAccountCredentials.from_json_keyfile_name('credentials.json', gs_scope)
    client = gspread.authorize(creds)
    spreadsheet = client.open(SPREADSHEET_NAME)
    sheet = spreadsheet.sheet1
    SPREADSHEET_URL = spreadsheet.url
except Exception as e:
    logger.error(f"Ошибка при подключении к Google Sheets: {e}")
    sys.exit(1)

# Инициализация Яндекс.Диска
try:
    y = yadisk.YaDisk(token=YANDEX_DISK_TOKEN)
    if not y.exists(YANDEX_DISK_FOLDER):
        y.mkdir(YANDEX_DISK_FOLDER)
except Exception as e:
    logger.error(f"Ошибка при подключении к Яндекс.Диску: {e}")
    sys.exit(1)

def upload_to_yadisk(local_path, remote_path):
    try:
        full_remote_path = f"{YANDEX_DISK_FOLDER}/{remote_path}"
        y.upload(local_path, full_remote_path)
        return True
    except Exception as e:
        logger.error(f"Ошибка загрузки на Яндекс.Диск: {e}")
        return False

# Регистрация шрифта для PDF
try:
    pdfmetrics.registerFont(TTFont('Arial', 'Arial.ttf'))
except:
    logger.warning("Arial font not found, using default")

# Хранение состояний и данных
user_states = {}
application_data = {}
user_chat_ids = {}

class Application:
    def __init__(self):
        self.device_type = None
        self.device_model = None
        self.problem = None
        self.comment = None
        self.name = None
        self.phone = None
        self.photo = None
        self.id = None
        self.chat_id = None
        self.date = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        self.status = 'Новая'

# Клавиатуры
def create_main_menu():
    kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
    kb.add('📝 Оставить заявку', '📊 Узнать статус')
    kb.add('📞 Связаться с мастером', '📢 Наш Telegram-канал')
    return kb

@bot.message_handler(commands=['start'])
def send_welcome(message):
    try:
        bot.send_message(message.chat.id, 'Добро пожаловать в RoboFix! Выберите действие:', reply_markup=create_main_menu())
    except Exception as e:
        logger.error(f"Ошибка в send_welcome: {e}")

# Обработка пунктов меню
@bot.message_handler(func=lambda m: m.text in ['📝 Оставить заявку', '📊 Узнать статус', '📞 Связаться с мастером', '📢 Наш Telegram-канал'])
def handle_menu(message):
    try:
        if message.text == '📝 Оставить заявку': 
            start_application(message)
        elif message.text == '📊 Узнать статус': 
            check_status(message)
        elif message.text == '📞 Связаться с мастером': 
            contact_master(message)
        else: 
            show_channel(message)
    except Exception as e:
        logger.error(f"Ошибка в handle_menu: {e}")
        bot.send_message(message.chat.id, "Произошла ошибка. Пожалуйста, попробуйте позже.")

# Начало заявки
def start_application(message):
    try:
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        kb.add('🔙 Назад')
        
        application_data[message.chat.id] = Application()
        application_data[message.chat.id].chat_id = message.chat.id
        user_states[message.chat.id] = 'device_type'
        bot.send_message(message.chat.id, 'Укажите тип устройства:', reply_markup=kb)
    except Exception as e:
        logger.error(f"Ошибка в start_application: {e}")
        bot.send_message(message.chat.id, "Ошибка при создании заявки. Пожалуйста, попробуйте позже.", reply_markup=create_main_menu())

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='device_type')
def handle_device_type(message):
    if message.text == '🔙 Назад':
        bot.send_message(message.chat.id, 'Возвращаемся в главное меню', reply_markup=create_main_menu())
        user_states.pop(message.chat.id, None)
        return
    try:
        application_data[message.chat.id].device_type = message.text
        user_states[message.chat.id] = 'device_model'
        bot.send_message(message.chat.id, 'Укажите модель устройства:')
    except Exception as e:
        logger.error(f"Ошибка в handle_device_type: {e}")

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='device_model')
def handle_device_model(message):
    try:
        application_data[message.chat.id].device_model = message.text
        user_states[message.chat.id] = 'problem'
        bot.send_message(message.chat.id, 'Опишите неисправность:')
    except Exception as e:
        logger.error(f"Ошибка в handle_device_model: {e}")

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='problem')
def handle_problem(message):
    try:
        application_data[message.chat.id].problem = message.text
        user_states[message.chat.id] = 'comment'
        bot.send_message(message.chat.id, 'Комментарий (или \'-\' если нет):')
    except Exception as e:
        logger.error(f"Ошибка в handle_problem: {e}")

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='comment')
def handle_comment(message):
    try:
        application_data[message.chat.id].comment = '' if message.text=='-' else message.text
        user_states[message.chat.id] = 'name'
        bot.send_message(message.chat.id, 'Ваше имя:')
    except Exception as e:
        logger.error(f"Ошибка в handle_comment: {e}")

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='name')
def handle_name(message):
    try:
        application_data[message.chat.id].name = message.text
        user_states[message.chat.id] = 'phone'
        bot.send_message(message.chat.id, 'Телефон (+7XXXXXXXXXX):')
    except Exception as e:
        logger.error(f"Ошибка в handle_name: {e}")

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='phone')
def handle_phone(message):
    try:
        if not re.match(r'^\+7\d{10}$', message.text):
            return bot.send_message(message.chat.id, 'Неверный формат. Введите +7XXXXXXXXXX')
        application_data[message.chat.id].phone = message.text
        user_states[message.chat.id] = 'photo'
        bot.send_message(message.chat.id, 'Пришлите фото устройства или /skip:')
    except Exception as e:
        logger.error(f"Ошибка в handle_phone: {e}")

@bot.message_handler(commands=['skip'], func=lambda m: user_states.get(m.chat.id)=='photo')
def skip_photo(message):
    try:
        user_states[message.chat.id] = 'preview'
        show_preview(message)
    except Exception as e:
        logger.error(f"Ошибка в skip_photo: {e}")

@bot.message_handler(content_types=['photo'], func=lambda m: user_states.get(m.chat.id)=='photo')
def handle_photo(message):
    try:
        file_info = bot.get_file(message.photo[-1].file_id)
        downloaded = bot.download_file(file_info.file_path)
        path = f"photos/{message.chat.id}_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
        os.makedirs(os.path.dirname(path), exist_ok=True)
        with open(path,'wb') as f: 
            f.write(downloaded)
        application_data[message.chat.id].photo = path
        
        # Загрузка фото на Яндекс.Диск
        remote_path = f"photos/{os.path.basename(path)}"
        if upload_to_yadisk(path, remote_path):
            logger.info(f"Фото загружено на Яндекс.Диск: {remote_path}")
        
        user_states[message.chat.id] = 'preview'
        show_preview(message)
    except Exception as e:
        logger.error(f"Ошибка в handle_photo: {e}")
        bot.send_message(message.chat.id, "Не удалось загрузить фото. Попробуйте еще раз или используйте /skip")

# Превью и подтверждение
def show_preview(message):
    try:
        app = application_data[message.chat.id]
        text = (
            f"Проверьте данные:\n"
            f"🔌 Устройство: {app.device_type}\n"
            f"ℹ️ Модель: {app.device_model}\n"
            f"⚙️ Неисправность: {app.problem}\n"
            f"📝 Комментарий: {app.comment or '-'}\n"
            f"👤 Имя: {app.name}\n"
            f"📞 Телефон: {app.phone}\n"
            f"📅 Дата: {app.date}\n"
            f"Отправить заявку? (да/нет)"
        )
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        kb.add('да','нет')
        
        if app.photo:
            with open(app.photo,'rb') as photo:
                bot.send_photo(message.chat.id, photo, caption=text, reply_markup=kb)
        else:
            bot.send_message(message.chat.id, text, reply_markup=kb)
    except Exception as e:
        logger.error(f"Ошибка в show_preview: {e}")
        bot.send_message(message.chat.id, "Ошибка при отображении предпросмотра. Пожалуйста, начните заново.", reply_markup=create_main_menu())

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='preview' and m.text.lower() in ['да','нет'])
def handle_preview_confirm(message):
    try:
        if message.text.lower()=='нет':
            bot.send_message(message.chat.id, 'Отмена.', reply_markup=create_main_menu())
            user_states.pop(message.chat.id, None)
            return
        
        app = application_data[message.chat.id]
        
        try:
            vals = sheet.get_all_values()
            if len(vals) > 1:
                last_id = vals[-1][0].strip()
                app.id = int(last_id) + 1 if last_id else 1
            else:
                app.id = 1
        except Exception as e:
            logger.error(f"Ошибка при генерации ID: {e}")
            app.id = 1
        
        user_chat_ids[app.id] = app.chat_id
        
        row = [
            app.id, app.date, app.name, app.phone,
            app.device_type, app.device_model, app.problem,
            app.comment or '-', app.photo or '', app.status, '', '', ''
        ]
        
        try:
            sheet.append_row(row)
            send_to_master(app)
            bot.send_message(
                app.chat_id, 
                '✅ Ваша заявка отправлена!\n\n'
                '📲 Свяжитесь с мастером и договоритесь о встрече:\n'
                '🛵 Передайте устройство мастеру, он подтвердит заявку\n'
                '📄 Вы получите квитанцию о приёме после осмотра', 
                reply_markup=create_main_menu()
            )
        except Exception as e:
            logger.error(f"Ошибка при сохранении в Google Sheets: {e}")
            bot.send_message(
                app.chat_id,
                "Ошибка при сохранении заявки. Пожалуйста, попробуйте позже.",
                reply_markup=create_main_menu()
            )
        
        user_states.pop(message.chat.id, None)
    except Exception as e:
        logger.error(f"Ошибка в handle_preview_confirm: {e}")
        bot.send_message(message.chat.id, "Ошибка при подтверждении заявки.", reply_markup=create_main_menu())

# Уведомление мастеру и стикер
def send_to_master(app):
    try:
        kb = types.InlineKeyboardMarkup()
        kb.add(
            types.InlineKeyboardButton('✅ Принять', callback_data=f'accept_{app.id}'),
            types.InlineKeyboardButton('❌ Отказать', callback_data=f'reject_{app.id}')
        )
        msg = (
            f"🔔 Новая заявка #{app.id}\n"
            f"🔌 Устройство: {app.device_type}\n"
            f"⚙️ Проблема: {app.problem}\n"
            f"👤 Имя: {app.name}\n"
            f"📞 Телефон: {app.phone}"
        )

        if app.photo:
            with open(app.photo,'rb') as photo:
                bot.send_photo(MASTER_ID, photo, caption=msg, reply_markup=kb)
        else:
            bot.send_message(MASTER_ID, msg, reply_markup=kb)
        
        sticker_path = generate_sticker_pdf(app)
        with open(sticker_path,'rb') as sticker:
            bot.send_document(MASTER_ID, sticker, caption=f"Стикер #{app.id}")
    except Exception as e:
        logger.error(f"Ошибка в send_to_master: {e}")

# Генерация стикера в PDF
def generate_sticker_pdf(app):
    try:
        width_mm, height_mm = 40, 30
        dpi = 300
        
        mm_to_px = dpi / 25.4
        width_px = int(width_mm * mm_to_px)
        height_px = int(height_mm * mm_to_px)

        img = Image.new('RGB', (width_px, height_px), 'white')
        draw = ImageDraw.Draw(img)

        font_size_mm = 2.2
        font_size_px = int(font_size_mm * mm_to_px)
        try:
            font = ImageFont.truetype('arial.ttf', font_size_px)
        except:
            font = ImageFont.load_default()
            font.size = font_size_px

        date_str = datetime.now().strftime("%d-%m")
        lines = [
            f"ID: {app.id}",
            f"Дата: {date_str}",
            f"Клиент: {app.name[:14]}" if len(app.name) > 14 else f"Клиент: {app.name}",
            f"Тел: {app.phone}",
            f"Пробл: {app.problem[:20] + '...' if len(app.problem) > 20 else app.problem}"
        ]

        text_margin_mm = 2.3
        text_margin_px = int(text_margin_mm * mm_to_px)
        line_spacing_mm = 4.5
        line_spacing_px = int(line_spacing_mm * mm_to_px)
        
        y_position = text_margin_px
        for line in lines:
            draw.text((text_margin_px, y_position), line, font=font, fill='black')
            y_position += line_spacing_px

        qr_size_mm = 16
        qr_size_px = int(qr_size_mm * mm_to_px)
        
        qr = qrcode.QRCode(
            version=3,
            error_correction=qrcode.constants.ERROR_CORRECT_H,
            box_size=4,
            border=1
        )
        qr.add_data(SPREADSHEET_URL)
        qr.make(fit=True)
        
        qr_img = qr.make_image(fill_color="black", back_color="white")
        qr_img = qr_img.resize((qr_size_px, qr_size_px))

        qr_margin_mm = 2
        qr_margin_px = int(qr_margin_mm * mm_to_px)
        qr_x = width_px - qr_size_px - qr_margin_px
        qr_y = qr_margin_px
        
        img.paste(qr_img, (qr_x, qr_y))

        pdf_path = f"stickers/стикер ({app.id}).pdf"
        os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
        
        from reportlab.lib.pagesizes import mm
        c = canvas.Canvas(pdf_path, pagesize=(width_mm*mm, height_mm*mm))
        
        img_io = BytesIO()
        img.save(img_io, format='PNG', dpi=(dpi, dpi))
        img_io.seek(0)
        
        c.drawImage(ImageReader(img_io), 0, 0, width_mm*mm, height_mm*mm)
        c.save()
        
        # Загрузка стикера на Яндекс.Диск
        remote_path = f"stickers/{os.path.basename(pdf_path)}"
        if upload_to_yadisk(pdf_path, remote_path):
            logger.info(f"Стикер загружен на Яндекс.Диск: {remote_path}")
        
        return pdf_path

    except Exception as e:
        logger.error(f"Ошибка генерации стикера: {str(e)}")
        raise RuntimeError("Не удалось создать стикер")

# Обработка действий мастера
@bot.callback_query_handler(func=lambda c: c.data.startswith(('accept_','reject_')))
def handle_master_action(call):
    try:
        action, sid = call.data.split('_')
        aid = int(sid)
        
        try:
            cell = sheet.find(str(aid))
            row = cell.row
        except Exception as e:
            logger.error(f"Не найдена заявка #{aid}: {e}")
            bot.answer_callback_query(call.id, "Заявка не найдена")
            return
        
        if action == 'accept':
            try:
                sheet.update_cell(row, 10, 'Принято')
                
                pdf_filename = f"Квитанция_№{aid}.pdf"
                pdf_path = f"pdf_receipts/{pdf_filename}"
                os.makedirs(os.path.dirname(pdf_path), exist_ok=True)
                create_pdf(aid, pdf_path)
                
                # Загрузка квитанции на Яндекс.Диск
                remote_path = f"pdf_receipts/{pdf_filename}"
                if upload_to_yadisk(pdf_path, remote_path):
                    logger.info(f"Квитанция загружена на Яндекс.Диск: {remote_path}")
                
                message_text = (
                    f"✅ Ваша заявка принята!\n"
                    f"📄 Вот квитанция о приеме устройства в ремонт.\n"
                    f"🆔 Номер вашей заявки: {aid}\n\n"
                    f"После диагностики мастер свяжется с вами для согласования стоимости ремонта."
                )
                
                if aid in user_chat_ids:
                    with open(pdf_path, 'rb') as f:
                        bot.send_document(
                            user_chat_ids[aid],
                            (pdf_filename, f),
                            caption=message_text
                        )
                
                bot.answer_callback_query(call.id, '✅ Заявка принята')
                
            except Exception as e:
                logger.error(f"Ошибка при обработке заявки: {e}")
                bot.answer_callback_query(call.id, '⚠️ Ошибка при обработке')
                
        else:
            sheet.update_cell(row, 10, 'Отклонено')
            if aid in user_chat_ids:
                bot.send_message(
                    user_chat_ids[aid],
                    f"❌ Ваша заявка №{aid} отклонена.\n\n"
                    f"По всем вопросам обращайтесь к мастеру."
                )
            bot.answer_callback_query(call.id, '❌ Заявка отклонена')
    except Exception as e:
        logger.error(f"Ошибка в handle_master_action: {e}")
        bot.answer_callback_query(call.id, '⚠️ Ошибка при обработке')

# Генерация PDF квитанции
def create_pdf(aid, output_path):
    try:
        cell = sheet.find(str(aid))
        data = sheet.row_values(cell.row)
        
        c = canvas.Canvas(output_path, pagesize=A4)
        
        try:
            c.setFont('Arial', 12)
        except:
            c.setFont('Helvetica', 12)
        
        date_parts = data[1].split()[0].split('-')
        formatted_date = f"{date_parts[2]}-{date_parts[1]}-{date_parts[0]}"
        
        c.drawString(50, 800, f"Квитанция №{aid}")
        labels = ['Имя','Телефон','Устройство','Модель','Неисправность','Комментарий','Дата']
        keys = [2,3,4,5,6,7,1]
        
        y = 770
        for lbl, idx in zip(labels, keys):
            value = data[idx] if idx < len(data) else ''
            if lbl == 'Дата':
                value = formatted_date
            c.drawString(50, y, f"{lbl}: {value}")
            y -= 20
        
        c.drawString(50, y, "Срок диагностики: 1-3 дня")
        y -= 30
        c.drawString(50, y, "Спасибо что обратились в наш сервис")
        c.save()
    except Exception as e:
        logger.error(f"Ошибка при создании PDF: {e}")
        raise

# Проверка статуса клиентом
@bot.message_handler(func=lambda m: m.text=='📊 Узнать статус')
def check_status(message):
    try:
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        kb.add('🔙 Назад')
        user_states[message.chat.id] = 'check'
        bot.send_message(message.chat.id, 'Введите ID заявки:', reply_markup=kb)
    except Exception as e:
        logger.error(f"Ошибка в check_status: {e}")

@bot.message_handler(func=lambda m: user_states.get(m.chat.id)=='check')
def handle_check(message):
    if message.text == '🔙 Назад':
        bot.send_message(message.chat.id, 'Возвращаемся в главное меню', reply_markup=create_main_menu())
        user_states.pop(message.chat.id, None)
        return
    try:
        aid = int(message.text)
        try:
            cell = sheet.find(str(aid))
            row = cell.row
            status = sheet.cell(row,10).value
            cost = sheet.cell(row,11).value
            icons = {'Новая':'🟡','Принято':'🟡','В работе':'🟠','Готово':'🟢','Отклонено':'🔴'}
            text = f"{icons.get(status, '')}{status}"
            if status=='Готово' and cost and cost!='':
                text += f"\nК оплате: {cost} руб. Свяжитесь с мастером."
            bot.send_message(message.chat.id, text, reply_markup=create_main_menu())
        except:
            bot.send_message(message.chat.id, 'ID не найден.', reply_markup=create_main_menu())
    except ValueError:
        bot.send_message(message.chat.id, 'ID должен быть числом. Попробуйте еще раз или нажмите "🔙 Назад"')
    except Exception as e:
        logger.error(f"Ошибка в handle_check: {e}")
        bot.send_message(message.chat.id, 'Произошла ошибка. Попробуйте позже.', reply_markup=create_main_menu())
    finally:
        user_states.pop(message.chat.id, None)

# Контакты мастера
def contact_master(message):
    try:
        formatted_phone = format_phone(MASTER_PHONE)
        whatsapp_link = f"https://wa.me/{MASTER_PHONE[1:]}"
        
        kb = types.InlineKeyboardMarkup()
        kb.row(
            types.InlineKeyboardButton(
                "📞 Позвонить", 
                callback_data=f'call_{MASTER_PHONE}'
            )
        )
        kb.row(
            types.InlineKeyboardButton(
                "✉️ Telegram", 
                url=f'tg://user?id={MASTER_ID}'
            ),
            types.InlineKeyboardButton(
                "💬 WhatsApp", 
                url=whatsapp_link
            )
        )
        
        contact_text = (
            "⌛ Режим работы: Пн-Сб 10:00-20:00"
        )
        
        bot.send_message(
            message.chat.id,
            contact_text,
            reply_markup=kb
        )
        
    except Exception as e:
        logger.error(f"Ошибка в contact_master: {e}")
        bot.send_message(
            message.chat.id,
            "Связь с мастером: @username_мастера",
            reply_markup=create_main_menu()
        )

def format_phone(phone):
    if not phone or len(phone) != 12 or not phone.startswith('+7'):
        return phone
    return f"{phone[:2]} ({phone[2:5]}) {phone[5:8]}-{phone[8:10]}-{phone[10:]}"

@bot.callback_query_handler(func=lambda call: call.data.startswith('call_'))
def handle_call(call):
    try:
        phone = call.data.split('_')[1]
        formatted_phone = format_phone(phone)
        bot.answer_callback_query(
            call.id,
            f"Телефон мастера: {formatted_phone}",
            show_alert=True
        )
    except Exception as e:
        logger.error(f"Ошибка в handle_call: {e}")
        bot.answer_callback_query(
            call.id,
            "Не удалось отобразить номер",
            show_alert=True
        )

def show_channel(message):
    bot.send_message(message.chat.id, "Наш канал: t.me/robotfixservice")

# Команды для мастера
@bot.message_handler(commands=['setstatus'])
def set_status(message):
    if message.from_user.id != MASTER_ID: 
        return
    
    try:
        parts = message.text.split()
        if len(parts) < 2: 
            return bot.send_message(message.chat.id, 'Используйте: /setstatus [ID]')
        
        aid = int(parts[1])
        user_states[message.chat.id] = f'set_{aid}'
        
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        kb.add('Принято','В работе','Готово','Выдано','Отклонено')
        
        bot.send_message(message.chat.id, 'Выберите статус:', reply_markup=kb)
    except ValueError:
        bot.send_message(message.chat.id, 'ID должен быть числом')
    except Exception as e:
        logger.error(f"Ошибка в set_status: {e}")
        bot.send_message(message.chat.id, 'Произошла ошибка')

@bot.message_handler(func=lambda m: isinstance(user_states.get(m.chat.id), str) and user_states[m.chat.id].startswith('set_'))
def handle_set_status(message):
    if message.from_user.id != MASTER_ID: 
        user_states.pop(message.chat.id, None)
        return
    
    try:
        aid = int(user_states[message.chat.id].split('_')[1])
        new_status = message.text
        
        try:
            cell = sheet.find(str(aid))
            row = cell.row
            sheet.update_cell(row, 10, new_status)
            
            bot.send_message(
                message.chat.id, 
                f"#{aid} => {new_status}", 
                reply_markup=create_main_menu()
            )
            
            if new_status == 'Готово' and aid in user_chat_ids:
                cost = sheet.cell(row, 11).value
                note = '🟢 Ваше устройство готово.'
                if cost and cost != '': 
                    note += f"\nК оплате: {cost} руб. Свяжитесь с мастером чтоб забрать устройство."
                bot.send_message(user_chat_ids[aid], note)
                
        except Exception as e:
            logger.error(f"Не найдена заявка #{aid}: {e}")
            bot.send_message(
                message.chat.id,
                f"Заявка #{aid} не найдена",
                reply_markup=create_main_menu()
            )
            
    except Exception as e:
        logger.error(f"Ошибка в handle_set_status: {e}")
        bot.send_message(
            message.chat.id,
            "Произошла ошибка",
            reply_markup=create_main_menu()
        )
    finally:
        user_states.pop(message.chat.id, None)

@bot.message_handler(commands=['mystat'])
def mystat(message):
    if message.from_user.id != MASTER_ID: 
        return
    
    try:
        # Создаем клавиатуру для выбора периода
        kb = types.ReplyKeyboardMarkup(resize_keyboard=True)
        kb.row('📊 Общая статистика')
        kb.row('📅 За текущий месяц', '📅 За прошлый месяц')
        kb.row('📆 За все время', '🔙 Назад')
        
        bot.send_message(
            message.chat.id,
            "Выберите период для статистики:",
            reply_markup=kb
        )
        user_states[message.chat.id] = 'stat_period'
    except Exception as e:
        logger.error(f"Ошибка в mystat: {e}")
        bot.send_message(message.chat.id, "Не удалось загрузить статистику", reply_markup=create_main_menu())

@bot.message_handler(func=lambda m: user_states.get(m.chat.id) == 'stat_period')
def handle_stat_period(message):
    if message.text == '🔙 Назад':
        bot.send_message(message.chat.id, "Возвращаемся в главное меню", reply_markup=create_main_menu())
        user_states.pop(message.chat.id, None)
        return
    
    try:
        records = sheet.get_all_records()
        
        if message.text == '📊 Общая статистика':
            # Статистика по статусам
            status_stats = {}
            total_cost = 0
            completed_cost = 0
            cost_values = []
            
            for r in records:
                status = r.get('Статус', 'Нет статуса')
                status_stats[status] = status_stats.get(status, 0) + 1
                
                # Считаем общую стоимость (берем 11-ю колонку, индексация с 0)
                cost_str = str(r.get('Стоимость', '0')).strip()
                if cost_str.isdigit():
                    cost = int(cost_str)
                    total_cost += cost
                    cost_values.append(cost)
                    if status == 'Готово':
                        completed_cost += cost
            
            avg_cost = round(total_cost/len(records)) if len(records) > 0 else 0
            if cost_values:
                avg_cost = round(sum(cost_values)/len(cost_values))
            
            text = (
                "📊 Общая статистика:\n"
                f"Всего заявок: {len(records)}\n"
                "\n📌 По статусам:\n"
            )
            text += '\n'.join(f"{k}: {v}" for k, v in sorted(status_stats.items()))
            
            text += (
                f"\n\n💰 Общая стоимость всех заявок: {total_cost} руб.\n"
                f"💰 Стоимость выполненных заявок: {completed_cost} руб.\n"
                f"💵 Средняя стоимость заявки: {avg_cost} руб."
            )
            
        elif message.text in ['📅 За текущий месяц', '📅 За прошлый месяц', '📆 За все время']:
            now = datetime.now()
            
            if message.text == '📅 За текущий месяц':
                month_records = [
                    r for r in records 
                    if parse_date(r.get('Дата', '')).month == now.month
                    and parse_date(r.get('Дата', '')).year == now.year
                ]
                text = generate_monthly_report(month_records, now.month, now.year)
                
            elif message.text == '📅 За прошлый месяц':
                last_month = now.month - 1 if now.month > 1 else 12
                last_year = now.year if now.month > 1 else now.year - 1
                month_records = [
                    r for r in records 
                    if parse_date(r.get('Дата', '')).month == last_month
                    and parse_date(r.get('Дата', '')).year == last_year
                ]
                text = generate_monthly_report(month_records, last_month, last_year)
                
            elif message.text == '📆 За все время':
                text = generate_full_report(records)
                
        else:
            return
            
        bot.send_message(
            message.chat.id,
            text,
            reply_markup=create_main_menu()
        )
        user_states.pop(message.chat.id, None)
        
    except Exception as e:
        logger.error(f"Ошибка в handle_stat_period: {str(e)}", exc_info=True)
        bot.send_message(
            message.chat.id,
            "Произошла ошибка при формировании отчета. Убедитесь, что данные в таблице корректны.",
            reply_markup=create_main_menu()
        )
        user_states.pop(message.chat.id, None)

def parse_date(date_str):
    """Парсит дату из строки в объект datetime"""
    try:
        # Пробуем разные форматы даты
        for fmt in ("%Y-%m-%d %H:%M:%S", "%d.%m.%Y %H:%M", "%Y-%m-%d"):
            try:
                return datetime.strptime(date_str, fmt)
            except ValueError:
                continue
        return datetime.min  # Возвращаем минимальную дату если не распарсилось
    except:
        return datetime.min

def generate_monthly_report(records, month, year):
    """Генерация отчета за месяц"""
    if not records:
        month_names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                     'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
        return f"📅 За {month_names[month-1]} {year} нет данных"
    
    status_stats = {}
    total_cost = 0
    completed_cost = 0
    cost_values = []
    
    for r in records:
        status = r.get('Статус', 'Нет статуса')
        status_stats[status] = status_stats.get(status, 0) + 1
        
        cost_str = str(r.get('Стоимость', '0')).strip()
        if cost_str.isdigit():
            cost = int(cost_str)
            total_cost += cost
            cost_values.append(cost)
            if status == 'Готово':
                completed_cost += cost
    
    avg_cost = round(total_cost/len(records)) if len(records) > 0 else 0
    if cost_values:
        avg_cost = round(sum(cost_values)/len(cost_values))
    
    month_names = ['Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                 'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
    
    completed = status_stats.get('Готово', 0)
    completion_rate = round(completed/len(records)*100) if len(records) > 0 else 0
    
    text = (
        f"📅 Отчет за {month_names[month-1]} {year}:\n"
        f"Всего заявок: {len(records)}\n"
        "\n📌 По статусам:\n"
    )
    text += '\n'.join(f"{k}: {v}" for k, v in sorted(status_stats.items()))
    
    text += (
        f"\n\n💰 Общая стоимость: {total_cost} руб.\n"
        f"💰 Стоимость выполненных: {completed_cost} руб.\n"
        f"💵 Средняя стоимость: {avg_cost} руб.\n"
        f"📈 Выполнено: {completed} из {len(records)} ({completion_rate}%)"
    )
    
    return text

def generate_full_report(records):
    """Генерация полного отчета за все время"""
    if not records:
        return "📆 Нет данных за все время"
    
    status_stats = {}
    monthly_stats = {}
    total_cost = 0
    completed_cost = 0
    cost_values = []
    
    for r in records:
        status = r.get('Статус', 'Нет статуса')
        status_stats[status] = status_stats.get(status, 0) + 1
        
        date = parse_date(r.get('Дата', ''))
        month_key = f"{date.year}-{date.month:02d}"
        monthly_stats[month_key] = monthly_stats.get(month_key, 0) + 1
        
        cost_str = str(r.get('Стоимость', '0')).strip()
        if cost_str.isdigit():
            cost = int(cost_str)
            total_cost += cost
            cost_values.append(cost)
            if status == 'Готово':
                completed_cost += cost
    
    avg_cost = round(total_cost/len(records)) if len(records) > 0 else 0
    if cost_values:
        avg_cost = round(sum(cost_values)/len(cost_values))
    
    completed = status_stats.get('Готово', 0)
    completion_rate = round(completed/len(records)*100) if len(records) > 0 else 0
    
    text = (
        "📆 Полная статистика:\n"
        f"Всего заявок: {len(records)}\n"
        "\n📌 По статусам:\n"
    )
    text += '\n'.join(f"{k}: {v}" for k, v in sorted(status_stats.items()))
    
    text += "\n\n📅 По месяцам:\n"
    for month in sorted(monthly_stats.keys()):
        year, month_num = map(int, month.split('-'))
        month_names = ['Янв', 'Фев', 'Мар', 'Апр', 'Май', 'Июн',
                      'Июл', 'Авг', 'Сен', 'Окт', 'Ноя', 'Дек']
        text += f"{month_names[month_num-1]} {year}: {monthly_stats[month]}\n"
    
    text += (
        f"\n💰 Общая стоимость: {total_cost} руб.\n"
        f"💰 Стоимость выполненных: {completed_cost} руб.\n"
        f"💵 Средняя стоимость: {avg_cost} руб.\n"
        f"📈 Выполнено: {completed} из {len(records)} ({completion_rate}%)"
    )
    
    return text

@bot.message_handler(commands=['money'])
def set_money(message):
    if message.from_user.id != MASTER_ID: 
        return
    
    try:
        parts = message.text.split()
        if len(parts) < 3: 
            return bot.send_message(message.chat.id, 'Используйте: /money [ID] [стоимость]')
        
        aid = int(parts[1])
        cost = parts[2]
        
        try:
            cell = sheet.find(str(aid))
            row = cell.row
            sheet.update_cell(row, 11, cost)
            bot.send_message(message.chat.id, f"Стоимость #{aid} установлена: {cost}")
        except:
            bot.send_message(message.chat.id, f"Заявка #{aid} не найдена")
            
    except ValueError:
        bot.send_message(message.chat.id, 'ID должен быть числом')
    except Exception as e:
        logger.error(f"Ошибка в set_money: {e}")
        bot.send_message(message.chat.id, "Произошла ошибка")

# Фоллбэк
@bot.message_handler(func=lambda _: True)
def fallback(message):
    bot.send_message(message.chat.id, 'Пожалуйста, выберите действие:', reply_markup=create_main_menu())

# Создание директорий
for d in ['photos','stickers', 'pdf_receipts']:
    os.makedirs(d, exist_ok=True)

# Запуск
if __name__ == '__main__':
    logger.info('Bot starting...')
    try:
        bot.infinity_polling()
    except Exception as e:
        logger.error(f"Bot stopped with error: {e}")
        sys.exit(1)