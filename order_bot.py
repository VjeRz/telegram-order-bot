#!/usr/bin/env python3
import logging
import os
import json
import re
from datetime import datetime, timedelta
from telegram import Update, InlineKeyboardButton, InlineKeyboardMarkup
from telegram.ext import (
    Application,
    CommandHandler,
    MessageHandler,
    CallbackQueryHandler,
    filters,
    ConversationHandler,
    ContextTypes,
)
import gspread
from oauth2client.service_account import ServiceAccountCredentials
import csv
import io

# ---------- CONFIGURATION ----------
TELEGRAM_BOT_TOKEN = os.environ.get("TELEGRAM_BOT_TOKEN")
GOOGLE_CREDENTIALS_JSON = os.environ.get("GOOGLE_CREDENTIALS_JSON")
ORDER_SHEET_NAME = os.environ.get("SPREADSHEET_NAME", "Order_Data_TelBot")
BOT_DATA_SHEET_NAME = os.environ.get("BOT_DATA_SHEET_NAME", "Bot_Data")
CLOUDFLARE_WORKER_URL = os.environ.get("CLOUDFLARE_WORKER_URL", "")
IT_TELEGRAM_ID = int(os.environ.get("IT_TELEGRAM_ID", "0"))

WAITING_FOR_ORDER_ID = 1
REG_NAME = 2
REG_EMAIL = 3
REG_ROLE_GROUP = 4
REG_SUBROLE = 5
REG_WOK = 6
REG_SFID = 7

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
logger = logging.getLogger(__name__)

# ---------- GOOGLE SHEETS SETUP ----------
def init_sheet(sheet_name):
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client.open(sheet_name).sheet1

order_sheet = init_sheet(ORDER_SHEET_NAME)
bot_data = init_sheet(BOT_DATA_SHEET_NAME)

# Users and UsageLog tabs (same as before)
try:
    users_sheet = bot_data.worksheet("Users")
except gspread.WorksheetNotFound:
    users_sheet = bot_data.add_worksheet("Users", 100, 20)
    users_sheet.append_row(["TelegramID", "Name", "Email", "RoleGroup", "SubRole", "ApprovalStatus", "RegistrationDate", "ApprovedBy", "ApprovedAt", "WOK", "SFID"])

try:
    usage_sheet = bot_data.worksheet("UsageLog")
except gspread.WorksheetNotFound:
    usage_sheet = bot_data.add_worksheet("UsageLog", 1000, 10)
    usage_sheet.append_row(["Timestamp", "TelegramID", "UserName", "RoleGroup", "SubRole", "OrderID"])

# ---------- HELPER FUNCTIONS ----------
def clean_text(s):
    if not s:
        return ""
    s = re.sub(r'[\u200b\u00a0\u200c\u200d]', '', str(s))
    return s.strip()

def find_order_details(order_id: str):
    clean_input = clean_text(order_id).lower()
    records = order_sheet.get_all_records()
    for record in records:
        sheet_value = clean_text(record.get("Order ID", "")).lower()
        if sheet_value == clean_input:
            return {
                "sto": record.get("STO", "-"),
                "wok": record.get("WOK", "-"),
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

def is_user_approved(telegram_id):
    records = users_sheet.get_all_records()
    for r in records:
        if str(r.get("TelegramID", "")) == str(telegram_id):
            return r.get("ApprovalStatus", "") == "approved"
    return False

def get_user_role(telegram_id):
    records = users_sheet.get_all_records()
    for r in records:
        if str(r.get("TelegramID", "")) == str(telegram_id):
            return r.get("RoleGroup", ""), r.get("SubRole", "")
    return None, None

def can_view_reports(telegram_id):
    _, subrole = get_user_role(telegram_id)
    return subrole in ["Manager", "Supervisor", "HSA", "IT"]

def log_usage(telegram_id, order_id):
    name, role_group, subrole = "", "", ""
    records = users_sheet.get_all_records()
    for r in records:
        if str(r.get("TelegramID", "")) == str(telegram_id):
            name = r.get("Name", "")
            role_group = r.get("RoleGroup", "")
            subrole = r.get("SubRole", "")
            break
    usage_sheet.append_row([
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        str(telegram_id),
        name,
        role_group,
        subrole,
        order_id
    ])

def notify_approver(bot, user_id, name, role_group, subrole, wok, sfid):
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("✅ Approve", callback_data=f"approve_{user_id}"),
         InlineKeyboardButton("❌ Reject", callback_data=f"reject_{user_id}")]
    ])
    bot.send_message(
        chat_id=IT_TELEGRAM_ID,
        text=f"New registration pending:\nName: {name}\nRole: {role_group} - {subrole}\nWOK: {wok}\nSFID: {sfid}\nUser ID: {user_id}",
        reply_markup=keyboard
    )

# ---------- REGISTRATION FLOW (unchanged) ----------
async def register_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if is_user_approved(user_id):
        await update.message.reply_text("You are already registered and approved.")
        return ConversationHandler.END
    records = users_sheet.get_all_records()
    for r in records:
        if str(r.get("TelegramID", "")) == str(user_id) and r.get("ApprovalStatus") == "pending":
            await update.message.reply_text("You already have a pending registration. Please wait for approval.")
            return ConversationHandler.END
    await update.message.reply_text("Welcome! Let's register.\n\nPlease enter your full name:")
    return REG_NAME

async def reg_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["reg_name"] = update.message.text
    await update.message.reply_text("Please enter your email address:")
    return REG_EMAIL

async def reg_email(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["reg_email"] = update.message.text
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Agency", callback_data="group_Agency")],
        [InlineKeyboardButton("Branch", callback_data="group_Branch")],
        [InlineKeyboardButton("Technician", callback_data="group_Technician")]
    ])
    await update.message.reply_text("Select your role group:", reply_markup=keyboard)
    return REG_ROLE_GROUP

async def reg_role_group(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    group = query.data.split("_")[1]
    context.user_data["reg_role_group"] = group
    if group == "Agency":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Team Leader", callback_data="sub_TL")],
            [InlineKeyboardButton("Salesforce", callback_data="sub_SF")]
        ])
    elif group == "Branch":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Inputters", callback_data="sub_IN")],
            [InlineKeyboardButton("Supervisor", callback_data="sub_SPV")],
            [InlineKeyboardButton("Manager", callback_data="sub_MGR")],
            [InlineKeyboardButton("IT", callback_data="sub_IT")]
        ])
    else:  # Technician
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Technician", callback_data="sub_TECH")],
            [InlineKeyboardButton("HSA", callback_data="sub_HSA")]
        ])
    await query.edit_message_text("Select your sub-role:", reply_markup=keyboard)
    return REG_SUBROLE

async def reg_subrole(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    sub = query.data.split("_")[1]
    sub_map = {
        "TL": "Team Leader", "SF": "Salesforce",
        "IN": "Inputters", "SPV": "Supervisor", "MGR": "Manager", "IT": "IT",
        "TECH": "Technician", "HSA": "HSA"
    }
    subrole = sub_map.get(sub, sub)
    context.user_data["reg_subrole"] = subrole
    wok_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Manado-Talaud", callback_data="wok_Manado-Talaud")],
        [InlineKeyboardButton("Bitung-Minahasa", callback_data="wok_Bitung-Minahasa")],
        [InlineKeyboardButton("Bolaang Mongondow", callback_data="wok_Bolaang Mongondow")],
        [InlineKeyboardButton("Gorontalo", callback_data="wok_Gorontalo")]
    ])
    await query.edit_message_text("Pilih WOK (Working Area):", reply_markup=wok_keyboard)
    return REG_WOK

async def reg_wok(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    wok = query.data.split("_")[1]
    context.user_data["reg_wok"] = wok
    await query.edit_message_text("Masukkan SF ID:")
    return REG_SFID

async def reg_sfid(update: Update, context: ContextTypes.DEFAULT_TYPE):
    sfid = update.message.text.strip()
    context.user_data["reg_sfid"] = sfid
    user_id = update.effective_user.id
    name = context.user_data["reg_name"]
    email = context.user_data["reg_email"]
    role_group = context.user_data["reg_role_group"]
    subrole = context.user_data["reg_subrole"]
    wok = context.user_data["reg_wok"]
    if role_group == "Branch":
        status = "approved"
        approved_by = "auto"
        approved_at = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        reply_text = "Registration complete! You are approved. You can now use /start to check orders."
    else:
        status = "pending"
        approved_by = ""
        approved_at = ""
        reply_text = "Registration submitted. You will be notified once approved by IT."
        notify_approver(update.get_bot(), user_id, name, role_group, subrole, wok, sfid)
    users_sheet.append_row([
        user_id, name, email, role_group, subrole, status,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), approved_by, approved_at,
        wok, sfid
    ])
    await update.message.reply_text(reply_text)
    return ConversationHandler.END

# ---------- APPROVAL HANDLERS ----------
async def approval_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    action, target_id = data.split("_")
    target_id = int(target_id)
    if action == "approve":
        records = users_sheet.get_all_records()
        for idx, row in enumerate(records, start=2):
            if str(row.get("TelegramID", "")) == str(target_id):
                users_sheet.update(f"F{idx}", "approved")
                users_sheet.update(f"H{idx}", str(update.effective_user.id))
                users_sheet.update(f"I{idx}", datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
                await context.bot.send_message(chat_id=target_id, text="Your registration has been approved! You can now use /start to check orders.")
                await query.edit_message_text(f"✅ User {target_id} approved.")
                break
    elif action == "reject":
        records = users_sheet.get_all_records()
        for idx, row in enumerate(records, start=2):
            if str(row.get("TelegramID", "")) == str(target_id):
                users_sheet.update(f"F{idx}", "rejected")
                await context.bot.send_message(chat_id=target_id, text="Your registration was rejected. You can try /register again.")
                await query.edit_message_text(f"❌ User {target_id} rejected.")
                break

async def pending(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if user_id != IT_TELEGRAM_ID:
        await update.message.reply_text("Only IT can view pending registrations.")
        return
    records = users_sheet.get_all_records()
    pending_list = []
    for row in records:
        if row.get("ApprovalStatus") == "pending" and row.get("RoleGroup") in ["Agency", "Technician"]:
            pending_list.append(f"{row.get('Name')} - {row.get('RoleGroup')}/{row.get('SubRole')} - WOK: {row.get('WOK', 'N/A')} - ID: {row.get('TelegramID')}")
    if not pending_list:
        await update.message.reply_text("No pending registrations.")
    else:
        await update.message.reply_text("Pending registrations:\n" + "\n".join(pending_list))

# ---------- REPORT COMMAND ----------
async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not can_view_reports(user_id):
        await update.message.reply_text("You are not authorized to view reports.")
        return
    args = context.args
    if not args:
        await update.message.reply_text("Usage: /report [day|week|month|from_date to_date]")
        return
    now = datetime.now()
    if args[0] == "day":
        start = now.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=1)
    elif args[0] == "week":
        start = now - timedelta(days=now.weekday())
        start = start.replace(hour=0, minute=0, second=0, microsecond=0)
        end = start + timedelta(days=7)
    elif args[0] == "month":
        start = now.replace(day=1, hour=0, minute=0, second=0, microsecond=0)
        next_month = start.replace(day=28) + timedelta(days=4)
        end = next_month - timedelta(days=next_month.day)
        end = end.replace(hour=23, minute=59, second=59)
    else:
        if len(args) < 2:
            await update.message.reply_text("For custom range: /report YYYY-MM-DD YYYY-MM-DD")
            return
        try:
            start = datetime.strptime(args[0], "%Y-%m-%d")
            end = datetime.strptime(args[1], "%Y-%m-%d") + timedelta(days=1)
        except ValueError:
            await update.message.reply_text("Invalid date format. Use YYYY-MM-DD.")
            return
    logs = usage_sheet.get_all_records()
    filtered = []
    for log in logs:
        try:
            ts = datetime.strptime(log.get("Timestamp"), "%Y-%m-%d %H:%M:%S")
            if start <= ts < end:
                filtered.append(log)
        except:
            continue
    if not filtered:
        await update.message.reply_text("No usage data in this period.")
        return
    summary = {}
    for log in filtered:
        uid = log.get("TelegramID")
        if uid not in summary:
            summary[uid] = {"name": log.get("UserName"), "role": log.get("SubRole"), "count": 0, "first": None, "last": None}
        summary[uid]["count"] += 1
        ts = datetime.strptime(log.get("Timestamp"), "%Y-%m-%d %H:%M:%S")
        if not summary[uid]["first"] or ts < summary[uid]["first"]:
            summary[uid]["first"] = ts
        if not summary[uid]["last"] or ts > summary[uid]["last"]:
            summary[uid]["last"] = ts
    if args[0] == "month" or len(args) > 1:
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(["TelegramID", "Name", "SubRole", "Number of lookups", "First lookup", "Last lookup", "Duration (minutes)"])
        for uid, data in summary.items():
            duration = (data["last"] - data["first"]).total_seconds() / 60 if data["first"] and data["last"] else 0
            writer.writerow([uid, data["name"], data["role"], data["count"], data["first"].strftime("%Y-%m-%d %H:%M:%S") if data["first"] else "", data["last"].strftime("%Y-%m-%d %H:%M:%S") if data["last"] else "", round(duration, 2)])
        output.seek(0)
        await update.message.reply_document(document=output, filename=f"report_{args[0]}.csv", caption=f"Report for {args[0]}")
    else:
        lines = [f"📊 Report for {args[0]}\n"]
        for uid, data in summary.items():
            duration = (data["last"] - data["first"]).total_seconds() / 60 if data["first"] and data["last"] else 0
            lines.append(f"👤 {data['name']} ({data['role']}) - {data['count']} lookups - Duration: {duration:.0f} min")
        await update.message.reply_text("\n".join(lines))

# ---------- ORDER LOOKUP (MODIFIED WITH STO & WOK) ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_user_approved(user_id):
        await update.message.reply_text("You are not registered or not approved. Please use /register to register.")
        return
    text = "Semangat Pagi, Masukan Order ID\nContoh: AOs326032509275620607db90"
    await update.message.reply_text(text)
    return WAITING_FOR_ORDER_ID

async def receive_order_id(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
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
    log_usage(user_id, order_id)
    reply = (
        f"✅ Order ID: {order_id}\n"
        f"📠 STO: {data['sto']}\n"
        f"🪪 WOK: {data['wok']}\n"
        f"⚙️ Order Status: {data['order_status']}\n"
        f"📢 Channel Name: {data['channel']}\n"
        f"⚠️ Fallout Reason: {data['fallout']}\n"
        f"👤 Salesforce: {data['salesforce']}\n"
        f"📅 Tanggal Complete: {data['tanggal_complete']}\n"
        f"📅 Tanggal Input: {data['tanggal_input']}\n"
        f"🧠 Sub Error Code: {data['sub_error']}\n"
        f"👨🏼‍🔧 Technician Notes: {data['technician_notes']}\n\n"
        f"Jika ingin mengecek lagi, ketik /start"
    )
    await update.message.reply_text(reply)
    return ConversationHandler.END

async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Perintah dibatalkan. Ketik /start untuk memulai lagi.")
    return ConversationHandler.END

async def ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("pong")

# ---------- MAIN ----------
def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN not set")
    if not GOOGLE_CREDENTIALS_JSON:
        raise ValueError("GOOGLE_CREDENTIALS_JSON not set")
    if not BOT_DATA_SHEET_NAME:
        raise ValueError("BOT_DATA_SHEET_NAME not set")
    if IT_TELEGRAM_ID == 0:
        logger.warning("IT_TELEGRAM_ID not set, approval notifications will not work")

    builder = Application.builder().token(TELEGRAM_BOT_TOKEN)
    if CLOUDFLARE_WORKER_URL:
        base_url = f"{CLOUDFLARE_WORKER_URL.rstrip('/')}/bot"
        builder = builder.base_url(base_url)
        logger.info(f"Using proxy base URL: {base_url}")
    else:
        logger.info("No proxy URL set")
    app = builder.build()

    # Registration conversation
    reg_conv = ConversationHandler(
        entry_points=[CommandHandler("register", register_start)],
        states={
            REG_NAME: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_name)],
            REG_EMAIL: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_email)],
            REG_ROLE_GROUP: [CallbackQueryHandler(reg_role_group, pattern="^group_")],
            REG_SUBROLE: [CallbackQueryHandler(reg_subrole, pattern="^sub_")],
            REG_WOK: [CallbackQueryHandler(reg_wok, pattern="^wok_")],
            REG_SFID: [MessageHandler(filters.TEXT & ~filters.COMMAND, reg_sfid)],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(reg_conv)
    app.add_handler(CallbackQueryHandler(approval_callback, pattern="^(approve|reject)_"))
    app.add_handler(CommandHandler("pending", pending))
    app.add_handler(CommandHandler("report", report))
    app.add_handler(CommandHandler("ping", ping))

    # Order lookup conversation
    order_conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAITING_FOR_ORDER_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_order_id)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(order_conv)

    logger.info("Bot is polling...")
    app.run_polling()

if __name__ == "__main__":
    main()
