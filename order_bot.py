#!/usr/bin/env python3
import logging
import os
import json
import re
from datetime import datetime
from telegram import Update
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    filters,
    ConversationHandler,
    ContextTypes,
)
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---------- CONFIGURATION ----------
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "Order_Data_TelBot")

CLOUDFLARE_WORKER_URL = os.environ.get("CLOUDFLARE_WORKER_URL", "")

WAITING_FOR_ORDER_ID = 1

# ---------- LOGGING ----------
logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# ---------- GOOGLE SHEETS SETUP ----------
def init_google_sheet():
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client.open(SPREADSHEET_NAME).sheet1

sheet = init_google_sheet()

# ---------- HELPER FUNCTIONS ----------
def clean_text(s):
    """Remove invisible characters and trim spaces"""
    if not s:
        return ""
    s = re.sub(r'[\u200b\u00a0\u200c\u200d]', '', str(s))
    return s.strip()

def find_order_details(order_id: str):
    """Case‑insensitive lookup returning all required fields"""
    clean_input = clean_text(order_id).lower()
    
    records = sheet.get_all_records()
    for record in records:
        sheet_value = clean_text(record.get("Order ID", "")).lower()
        if sheet_value == clean_input:
            return {
                "order_status": record.get("Status Order", "N/A"),
                "channel": record.get("Channel Name", "N/A"),
                "fallout": record.get("Fallout Reason", "") or "(Blank)",
                "salesforce": record.get("SalesForce", "N/A"),
                "tanggal_complete": record.get("Tanggal Complete", "-"),
                "tanggal_input": record.get("Tanggal Input", "-"),
                "sub_error": record.get("Sub Error Code", "-"),
                "technician_notes": record.get("Technician Notes", "-"),
            }
    return None

def get_last_update_time():
    return datetime.now().strftime("%d/%m/%Y")

# ---------- BOT HANDLERS ----------
async def ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("pong")

async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    text = "Semangat Pagi, Masukan Order ID\nContoh: AOs326032509275620607db90"
    await update.message.reply_text(text)
    return WAITING_FOR_ORDER_ID

async def receive_order_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    raw_input = update.message.text
    order_id = clean_text(raw_input)

    data = find_order_details(order_id)

    if data is None:
        last_update = get_last_update_time()
        error_msg = (
            f"❌ Maaf Order ID Tidak Ditemukan atau Belum Terupdate\n"
            f"📅 Last Update Data: {last_update}\n\n"
            f"Silahkan Coba Lagi dengan memasukan Order ID Lain atau Perbaiki formatnya."
        )
        await update.message.reply_text(error_msg)
        return WAITING_FOR_ORDER_ID

    reply = (
        f"Order ID: {order_id}\n"
        f"Order Status: {data['order_status']}\n"
        f"Channel Name: {data['channel']}\n"
        f"Fallout Reason: {data['fallout']}\n"
        f"Salesforce: {data['salesforce']}\n"
        f"Tanggal Complete: {data['tanggal_complete']}\n"
        f"Tanggal Input: {data['tanggal_input']}\n"
        f"Sub Error Code: {data['sub_error']}\n"
        f"Technician Notes: {data['technician_notes']}\n\n"
        f"Jika ingin mengecek lagi, ketik /start"
    )
    await update.message.reply_text(reply)
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Perintah dibatalkan. Ketik /start untuk memulai lagi.")
    return ConversationHandler.END

# ---------- MAIN ----------
def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN environment variable not set")
    if not GOOGLE_CREDENTIALS_JSON:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")

    builder = Application.builder().token(TELEGRAM_BOT_TOKEN)
    if CLOUDFLARE_WORKER_URL:
        base_url = f"{CLOUDFLARE_WORKER_URL.rstrip('/')}/bot"
        builder = builder.base_url(base_url)
        logger.info(f"Using proxy base URL: {base_url}")
    else:
        logger.info("No proxy URL set, connecting directly to Telegram API")

    app = builder.build()

    app.add_handler(CommandHandler("ping", ping))

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={
            WAITING_FOR_ORDER_ID: [
                MessageHandler(filters.TEXT & ~filters.COMMAND, receive_order_id)
            ]
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(conv_handler)

    logger.info("Bot is polling...")
    app.run_polling()

if __name__ == "__main__":
    main()
