import os
import logging
import asyncio
import requests
import edge_tts
import sqlite3
import re
from telegram import Update
from telegram.ext import Application, MessageHandler, CommandHandler, filters, ContextTypes
from faster_whisper import WhisperModel
from bs4 import BeautifulSoup
import glob
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

logging.basicConfig(format='%(asctime)s - %(name)s - %(levelname)s - %(message)s', level=logging.INFO)
VOICE_ENABLED = True
LAST_SITE_CODE = ""  # Память для хранения последнего сгенерированного сайта

# --- Whisper (ВРЕМЕННО ОТКЛЮЧЕН ДЛЯ ОБЛАКА) ---
# model_path = r"D:\Jarvis_Release\models\models--Systran--faster-whisper-small\snapshots"
# snapshot_dir = glob.glob(os.path.join(model_path, "*"))[0]
# model = WhisperModel(snapshot_dir, device="cpu", compute_type="int8")
# print("Модель загружена.")

def create_file(filename, content):
    # Универсальный путь (работает и на Windows, и на Linux)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    path = os.path.join(base_dir, "projects", filename)
    os.makedirs(os.path.dirname(path), exist_ok=True)
    with open(path, "w", encoding="utf-8") as f: 
        f.write(content)
    return path

def ask_deepseek(prompt, system_prompt=None, model_name="anthropic/claude-haiku-5.5"):
    # ВСТАВЬ СЮДА СВОЙ КЛЮЧ ОТ OFOXAI
    api_key = "sk-of-CetcdLRVvTnJxNxHvF1ITmrYZbQLEMMBJzJtmtXoTjCzzsJUCbXwmPUTJNkjA1BFY" 
    
    url = "https://api.ofox.ai/anthropic/v1/messages"
    
    headers = {
        "x-api-key": api_key,
        "anthropic-version": "2023-06-01",
        "Content-Type": "application/json"
    }
    
    data = {
        "model": model_name,
        "max_tokens": 8192,
        "messages": [{"role": "user", "content": prompt}],
        "thinking": {"type": "disabled"}
    }
    
    if system_prompt:
        data["system"] = system_prompt
    else:
        data["system"] = "Ты — Жорик. Отвечай кратко на РУССКОМ."
    
    try:
        r = requests.post(url, headers=headers, json=data, timeout=300)
        r.raise_for_status()
        response_data = r.json()
        
        content_blocks = response_data.get("content", [])
        for block in content_blocks:
            if block.get("type") == "text":
                return block.get("text", "")
        
        return f"⚠️ Ответ API (нет текста): {response_data}"
            
    except Exception as e:
        return f"❌ Ошибка сети: {e}"

def audit_site(url: str) -> str:
    try:
        headers = {'User-Agent': 'Mozilla/5.0'}
        resp = requests.get(url, headers=headers, timeout=10)
        resp.raise_for_status()
        html = resp.text.lower()
        
        data = {
            "Мобильная версия": "✅ Есть" if 'viewport' in html else "❌ Нет",
            "Онлайн-запись": "✅ Есть" if any(x in html for x in ['yclients', 'dikidi', 'записаться', 'забронировать']) else "❌ Нет",
            "Отзывы": "✅ Есть" if any(x in html for x in ['отзыв', 'review', 'яндекс']) else "❌ Нет",
            "Рекомендации": "✅ Есть" if any(x in html for x in ['приведи друга', 'поделиться', 'share']) else "❌ Нет",
            "Цены": "✅ Есть" if any(x in html for x in ['цена', 'стоимость', 'руб']) else "❌ Нет",
            "Скорость": "✅ Быстро" if len(resp.content) < 100000 else "⚠️ Медленно",
            "HTTPS": "✅ Есть" if url.startswith('https') else "❌ Нет"
        }

        wb = Workbook()
        ws = wb.active
        ws.title = "Аудит сайта"

        ws['A1'] = "Параметр"
        ws['B1'] = "Статус"
        
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
        ws['A1'].font = header_font
        ws['B1'].font = header_font
        ws['A1'].fill = header_fill
        ws['B1'].fill = header_fill
        ws.column_dimensions['A'].width = 25
        ws.column_dimensions['B'].width = 20

        for row_idx, (key, value) in enumerate(data.items(), start=2):
            ws.cell(row=row_idx, column=1, value=key)
            ws.cell(row=row_idx, column=2, value=value)
            if "✅" in value:
                ws.cell(row=row_idx, column=2).font = Font(color="008000")
            elif "❌" in value:
                ws.cell(row=row_idx, column=2).font = Font(color="FF0000")
            elif "⚠️" in value:
                ws.cell(row=row_idx, column=2).font = Font(color="FFA500")

        filename = "audit_report.xlsx"
        base_dir = os.path.dirname(os.path.abspath(__file__))
        path = os.path.join(base_dir, "projects", filename)
        os.makedirs(os.path.dirname(path), exist_ok=True)
        wb.save(path)

        return path

    except Exception as e:
        return f"❌ Ошибка: {e}"

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Привет. Отправь ссылку (http...) для аудита. Или напиши 'Сделай сайт' и описание.")

async def handle_text(update: Update, context: ContextTypes.DEFAULT_TYPE):
    global LAST_SITE_CODE
    text = update.message.text
    
    if text.startswith("http://") or text.startswith("https://"):
        status_msg = await update.message.reply_text("🔍 Анализирую и создаю таблицу...")
        result = audit_site(text)
        
        if result.endswith(".xlsx"):
            await status_msg.edit_text("✅ Аудит готов! Отправляю файл Excel...")
            with open(result, "rb") as f:
                await update.message.reply_document(document=f, filename="audit_report.xlsx")
        else:
            await status_msg.edit_text(result)
        return

    is_new_site = "сделай сайт" in text.lower()
    
    trigger_words = [
        "добавь", "измени", "исправь", "сделай фон", "переделай", "убери", 
        "поменяй", "сделай кнопку", "изменить", "поменять", "верни", "вставь", 
        "напиши в", "измени цвет", "перекрась", "замени"
    ]
    is_correction = any(word in text.lower() for word in trigger_words) and LAST_SITE_CODE != ""

    if is_new_site or is_correction:
        if is_new_site and len(text.split()) <= 2:
            await update.message.reply_text("Пришли описание сайта, и я сгенерирую файл.")
            return
            
        status_msg = await update.message.reply_text("⏳ Генерирую сайт... Это займет 1-2 минуты.")
        
        if is_correction:
            prompt = (
                f"Вот текущий HTML код сайта:\n```html\n{LAST_SITE_CODE}\n```\n"
                f"Твоя задача: внести в этот код ТОЛЬКО следующие изменения: {text}.\n"
                f"КРИТИЧЕСКИ ВАЖНО: НЕ меняй остальной дизайн, тексты, структуру и цвета, если об этом не просят. "
                f"Верни ПОЛНЫЙ обновленный код HTML с CSS внутри <style>."
            )
            await status_msg.edit_text("🛠 Вношу правки в сайт...")
        else:
            prompt = f"Напиши код одностраничного сайта. Задача: {text}. Выведи ТОЛЬКО код HTML с CSS внутри <style>."
        
        code = ask_deepseek(prompt, "Ты — веб-разработчик. Пиши только чистый HTML/CSS.")
        
        code = re.sub(r'^.*?(?=<!DOCTYPE|<html)', '', code, flags=re.DOTALL|re.IGNORECASE)
        code = re.sub(r'</html>.*$', '</html>', code, flags=re.DOTALL|re.IGNORECASE)
        
        LAST_SITE_CODE = code
        
        path = create_file("index_generated.html", code)
        await status_msg.edit_text("✅ Сайт готов! Отправляю файл...")
        with open(path, "rb") as f:
            await update.message.reply_document(document=f, filename="index_generated.html")
        return

    await update.message.reply_text(f"🤖 Жорик: {ask_deepseek(text)}")

def main():
    TOKEN = "9173136903:AA... (вставь сюда токен от Jarik_Cloud_bot)"
    app = Application.builder().token(TOKEN).build()
    app.add_handler(CommandHandler("start", start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_text))
    print("Жорик запущен.")
    app.run_polling()

if __name__ == "__main__":
    main()
