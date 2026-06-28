import os
import re
import logging
import json
from typing import Dict, List
import pandas as pd
import openpyxl
from flask import Flask, request, jsonify
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ContextTypes, CallbackQueryHandler
import tempfile

# Настройка логирования
logging.basicConfig(
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    level=logging.INFO
)
logger = logging.getLogger(__name__)

BOT_TOKEN = "8735478967:AAHn5bcYleXLqA-pfgUmItFDBqo4zyBczwk"

# Создаем приложение
application = Application.builder().token(BOT_TOKEN).build()
app = Flask(__name__)

class ExcelParser:
    # ... (весь код парсера из предыдущей версии)
    @staticmethod
    def detect_columns(df: pd.DataFrame) -> Dict[str, str]:
        columns_map = {'article': None, 'size': None, 'quantity': None}
        article_keywords = ['артикул', 'код', 'article', 'code', 'id', 'номер', '№', 'sku', 'арт', 'товар']
        size_keywords = ['размер', 'size', 'р-р', 's/m/l', 'рост', 'объем', 'разм', 'р']
        quantity_keywords = ['количество', 'кол-во', 'quantity', 'qty', 'шт', 'к-во', 'кол', 'остаток']
        
        for col in df.columns:
            col_lower = str(col).lower().strip()
            if not columns_map['article']:
                for keyword in article_keywords:
                    if keyword in col_lower:
                        columns_map['article'] = col
                        break
            if not columns_map['size']:
                for keyword in size_keywords:
                    if keyword in col_lower:
                        columns_map['size'] = col
                        break
            if not columns_map['quantity']:
                for keyword in quantity_keywords:
                    if keyword in col_lower:
                        columns_map['quantity'] = col
                        break
        
        if not columns_map['article']:
            for col in df.columns:
                sample = df[col].dropna().head(10)
                if len(sample) > 0:
                    if all(isinstance(x, str) for x in sample):
                        if any(re.search(r'[A-Za-zА-Яа-я]', str(x)) and re.search(r'\d', str(x)) for x in sample):
                            columns_map['article'] = col
                            break
        
        if not columns_map['quantity']:
            for col in df.columns:
                sample = df[col].dropna().head(10)
                if len(sample) > 0 and all(isinstance(x, (int, float)) for x in sample):
                    if any(x > 0 for x in sample):
                        columns_map['quantity'] = col
                        break
        
        return columns_map
    
    @staticmethod
    def parse_excel(file_path: str) -> List[Dict]:
        try:
            try:
                df = pd.read_excel(file_path, engine='openpyxl')
            except:
                df = pd.read_excel(file_path)
            
            columns = ExcelParser.detect_columns(df)
            
            if not any(columns.values()):
                try:
                    df = pd.read_excel(file_path, engine='openpyxl', header=None)
                    for i in range(min(5, len(df))):
                        row = df.iloc[i].astype(str).str.lower()
                        if any('артикул' in x or 'код' in x for x in row):
                            df = pd.read_excel(file_path, engine='openpyxl', header=i)
                            columns = ExcelParser.detect_columns(df)
                            break
                except:
                    pass
            
            products = []
            
            if not any(columns.values()):
                return ExcelParser.parse_simple_excel(df)
            
            article_col = columns.get('article')
            size_col = columns.get('size')
            quantity_col = columns.get('quantity')
            
            for idx, row in df.iterrows():
                try:
                    product = {}
                    if article_col and pd.notna(row[article_col]):
                        product['article'] = str(row[article_col]).strip()
                    else:
                        continue
                    if size_col and pd.notna(row[size_col]):
                        product['size'] = str(row[size_col]).strip()
                    else:
                        product['size'] = 'не указан'
                    if quantity_col and pd.notna(row[quantity_col]):
                        try:
                            product['quantity'] = int(float(row[quantity_col]))
                        except:
                            product['quantity'] = 0
                    else:
                        product['quantity'] = 1
                    if product['article'] and product['quantity'] > 0:
                        products.append(product)
                except:
                    continue
            return products
        except Exception as e:
            logger.error(f"Ошибка при чтении Excel: {e}")
            return []
    
    @staticmethod
    def parse_simple_excel(df: pd.DataFrame) -> List[Dict]:
        products = []
        for idx, row in df.iterrows():
            try:
                row_values = [str(x).strip() for x in row if pd.notna(x)]
                if len(row_values) < 2:
                    continue
                article = None
                size = 'не указан'
                quantity = 1
                for val in row_values:
                    if re.search(r'[A-Za-zА-Яа-я]{2,}[\d\-]+', val) or re.search(r'[\d\-]+[A-Za-zА-Яа-я]{2,}', val):
                        article = val
                    elif re.match(r'^[SMLXL\d]+$', val.upper()):
                        size = val.upper()
                    elif val.isdigit() and int(val) > 0:
                        quantity = int(val)
                if article:
                    products.append({'article': article, 'size': size, 'quantity': quantity})
            except:
                continue
        return products
    
    @staticmethod
    def format_products(products: List[Dict]) -> str:
        if not products:
            return "❌ В файле не найдено товаров."
        result = "📦 **Найденные товары:**\n\n"
        display_products = products[:50]
        result += "```\n"
        result += f"{'№':<4} {'Артикул':<25} {'Размер':<10} {'Кол-во':<6}\n"
        result += "-" * 50 + "\n"
        for i, product in enumerate(display_products, 1):
            article = product['article'][:23] + '..' if len(product['article']) > 25 else product['article']
            result += f"{i:<4} {article:<25} {product['size']:<10} {product['quantity']:<6}\n"
        result += "```\n"
        if len(products) > 50:
            result += f"\n... и еще {len(products) - 50} товаров"
        total_quantity = sum(p['quantity'] for p in products)
        unique_articles = len(set(p['article'] for p in products))
        result += f"\n\n📊 **Итого:** {len(products)} позиций, {total_quantity} шт."
        result += f"\n📌 **Уникальных артикулов:** {unique_articles}"
        size_stats = {}
        for p in products:
            size = p['size'] if p['size'] else 'не указан'
            size_stats[size] = size_stats.get(size, 0) + p['quantity']
        if size_stats:
            result += "\n\n📏 **По размерам:**\n"
            for size, count in sorted(size_stats.items()):
                result += f"  • {size}: {count} шт.\n"
        return result

# Регистрируем обработчики команд
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user = update.effective_user
    welcome_text = f"👋 Привет, {user.first_name}!\n\nЯ бот для парсинга Excel файлов с товарами.\nОтправь мне Excel файл (.xlsx, .xls), и я извлеку из него артикулы, размеры и количество."
    await update.message.reply_text(welcome_text)

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = "🤖 **Помощь по боту**\n\nОтправьте Excel файл с колонками: Артикул, Размер, Количество\n\nКоманды:\n/start - Приветствие\n/help - Эта справка"
    await update.message.reply_text(help_text, parse_mode='Markdown')

async def handle_excel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    document = update.message.document
    file_name = document.file_name.lower()
    if not (file_name.endswith('.xlsx') or file_name.endswith('.xls')):
        await update.message.reply_text("❌ Пожалуйста, отправьте файл в формате Excel (.xlsx или .xls)")
        return
    if document.file_size > 10 * 1024 * 1024:
        await update.message.reply_text("❌ Файл слишком большой. Максимум 10MB")
        return
    processing_msg = await update.message.reply_text("⏳ Обрабатываю Excel файл...")
    try:
        file = await context.bot.get_file(document.file_id)
        with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as tmp_file:
            await file.download_to_drive(tmp_file.name)
            tmp_file_path = tmp_file.name
        products = ExcelParser.parse_excel(tmp_file_path)
        os.unlink(tmp_file_path)
        if not products:
            await processing_msg.edit_text("❌ Не найдено товаров в файле.")
            return
        result_text = ExcelParser.format_products(products)
        await processing_msg.edit_text(result_text, parse_mode='Markdown')
    except Exception as e:
        logger.error(f"Ошибка: {e}")
        await processing_msg.edit_text(f"❌ Ошибка: {str(e)}")

# Настройка webhook
@app.route('/webhook', methods=['POST'])
async def webhook():
    try:
        data = request.get_json()
        update = Update.de_json(data, application.bot)
        await application.process_update(update)
        return jsonify({'status': 'ok'})
    except Exception as e:
        logger.error(f"Webhook error: {e}")
        return jsonify({'status': 'error', 'message': str(e)}), 500

@app.route('/')
def index():
    return "🤖 Excel Telegram Bot is running!"

@app.route('/ping')
def ping():
    return "pong"

# Инициализация
application.add_handler(CommandHandler("start", start))
application.add_handler(CommandHandler("help", help_command))
application.add_handler(MessageHandler(filters.Document.ALL, handle_excel))

if __name__ == '__main__':
    # Устанавливаем webhook
    import requests
    webhook_url = "https://ваш-сайт.com/webhook"  # Замените на URL вашего сервиса
    requests.post(f"https://api.telegram.org/bot{BOT_TOKEN}/setWebhook", json={'url': webhook_url})
    
    # Запускаем Flask
    app.run(host='0.0.0.0', port=8080)
