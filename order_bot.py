#!/usr/bin/env python3
import logging
import os
import json
from datetime import datetime
from telegram import Update
from telegram.ext import Application, CommandHandler, MessageHandler, filters, ConversationHandler, ContextTypes
import gspread
from oauth2client.service_account import ServiceAccountCredentials

# ---------- CONFIGURATION ----------
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON")
SPREADSHEET_NAME = os.environ.get("SPREADSHEET_NAME", "YOUR_SPREADSHEET_NAME")

WAITING_FOR_ORDER_ID = 1
TIMEOUT_SECONDS = 300

# ---------- LOGGING ----------
logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
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
def find_order_details(order_id: str):
    records = sheet.get_all_records()
    for record in records:
        if str(record.get("Order ID", "")).strip() == order_id.strip():
            return {
                "channel": record.get("Channel Name", "N/A"),
                "salesforce": record.get("SalesForce", "N/A"),
                "submit_date": record.get("Tanggal Submit", "N/A"),
                "status": record.get("Status Order", "N/A"),
                "fallout": record.get("Fallout Reason", "") or "(Blank)"
            }
    return None

def get_last_update_time():
    return datetime.now().strftime("%d/%m/%Y")

# ---------- BOT HANDLERS ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    hour = datetime.now().hour
    if hour < 12: greeting = "Good Morning"
    elif hour < 18: greeting = "Good Afternoon"
    else: greeting = "Good Night"

    text = f"{greeting}, Masukan Order ID\nContoh: AOs326032509275620607db90"
    await update.message.reply_text(text)

    if "order_timeout" in context.chat_data:
        context.chat_data["order_timeout"].schedule_removal()
    job = context.job_queue.run_once(timeout_callback, TIMEOUT_SECONDS, data={"chat_id": update.effective_chat.id})
    context.chat_data["order_timeout"] = job
    return WAITING_FOR_ORDER_ID

async def timeout_callback(context: ContextTypes.DEFAULT_TYPE):
    await context.bot.send_message(chat_id=context.job.chat_id, text="⏰ Waktu habis. Silakan ketik /start untuk memulai lagi.")

async def receive_order_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    order_id = update.message.text.strip()

    if "order_timeout" in context.chat_data:
        context.chat_data["order_timeout"].schedule_removal()
        del context.chat_data["order_timeout"]

    data = find_order_details(order_id)

    if data is None:
        last_update = get_last_update_time()
        error_msg = (f"❌ *Maaf Order ID Tidak Ditemukan atau Belum Terupdate*\n"
                     f"📅 *Last Update Data:* {last_update}\n\n"
                     f"Silahkan Coba Lagi dengan memasukan Order ID Lain atau Perbaiki formatnya.")
        await update.message.reply_text(error_msg, parse_mode="Markdown")
        return WAITING_FOR_ORDER_ID

    reply = (f"✅ *Order ID:* {order_id}\n"
             f"📢 *Channel Name:* {data['channel']}\n"
             f"👤 *SalesForce:* {data['salesforce']}\n"
             f"📅 *Tanggal Submit:* {data['submit_date']}\n"
             f"⚙️ *Status Order:* {data['status']}\n"
             f"⚠️ *Fallout Reason:* {data['fallout']}\n\n"
             f"Jika ingin mengecek lagi, ketik /start")
    await update.message.reply_text(reply, parse_mode="Markdown")
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if "order_timeout" in context.chat_data:
        context.chat_data["order_timeout"].schedule_removal()
        del context.chat_data["order_timeout"]
    await update.message.reply_text("Perintah dibatalkan. Ketik /start untuk memulai lagi.")
    return ConversationHandler.END

# ---------- MAIN ----------
def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN environment variable not set")
    if not GOOGLE_CREDENTIALS_JSON:
        raise ValueError("GOOGLE_CREDENTIALS_JSON environment variable not set")

    app = Application.builder().token(TELEGRAM_BOT_TOKEN).build()

    conv_handler = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAITING_FOR_ORDER_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_order_id)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(conv_handler)

    logger.info("Bot is polling...")
    app.run_polling()

if __name__ == "__main__":
    main()