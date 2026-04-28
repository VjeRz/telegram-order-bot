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

# Conversation states
WAITING_FOR_SINGLE_ORDER = 1
REG_NAME = 2
REG_EMAIL = 3
REG_ROLE_GROUP = 4
REG_SUBROLE = 5
REG_WOK = 6
REG_SFID = 7
BULK_AWAITING_IDS = 10
SALES_WOK = 20
DETAIL_WOK = 21          # new state for detailed report WOK selection
SALES_CHANNEL = 22
SALES_YEAR = 23
SALES_MONTH = 24
SUMMARY_YEAR = 30
SUMMARY_MONTH = 31

ALL_STATUSES = [
    "PENDING_CUSTOMER_VERIFICATION", "PROVISION_START", "TECH_ASSIGNED",
    "PENDING_APPOINTMENT_CREATION", "PENDING_CONTRACT_APPROVAL", "PROVISION_ISSUED",
    "COMPLETED", "OSS_TESTING_SERVICE", "RE", "FALLOUT", "ODP_AVAILABLE",
    "CANCELLED", "PENDING_PAYMENT_FOLLOWUP", "PAYMENT_INPROGRESS", "CANCEL_OSM_COMPLETED",
    "TSEL_ACTIVATION_FALLOUT", "CANCEL_ORDER_INPROGRESS", "TECH_ARRIVED",
    "CANCELLED_SLA", "PENDING_DUNNING_PAYMENT_FOLLOWUP", "PENDING_PAYMENT",
    "TECH_PICKED_UP", "TECH_ON_THE_WAY", "CONTRACT_APPROVED",
    "WIRELESS_FULFILMENT_INPROGRESS", "PROVISION_DESIGN"
]

logging.basicConfig(format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO)
logger = logging.getLogger(__name__)

# ---------- GOOGLE SHEETS SETUP ----------
def init_order_sheet(sheet_name):
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client.open(sheet_name).sheet1

def get_bot_data_spreadsheet(spreadsheet_name):
    creds_dict = json.loads(GOOGLE_CREDENTIALS_JSON)
    scope = ["https://spreadsheets.google.com/feeds", "https://www.googleapis.com/auth/drive"]
    creds = ServiceAccountCredentials.from_json_keyfile_dict(creds_dict, scope)
    client = gspread.authorize(creds)
    return client.open(spreadsheet_name)

order_sheet = init_order_sheet(ORDER_SHEET_NAME)
bot_data_spreadsheet = get_bot_data_spreadsheet(BOT_DATA_SHEET_NAME)

try:
    users_sheet = bot_data_spreadsheet.worksheet("Users")
except gspread.WorksheetNotFound:
    users_sheet = bot_data_spreadsheet.add_worksheet("Users", 100, 20)
    users_sheet.append_row(["TelegramID", "Name", "Email", "RoleGroup", "SubRole", "ApprovalStatus", "RegistrationDate", "ApprovedBy", "ApprovedAt", "WOK", "SFID"])

try:
    usage_sheet = bot_data_spreadsheet.worksheet("UsageLog")
except gspread.WorksheetNotFound:
    usage_sheet = bot_data_spreadsheet.add_worksheet("UsageLog", 1000, 10)
    usage_sheet.append_row(["Timestamp", "TelegramID", "UserName", "RoleGroup", "SubRole", "OrderID"])

# ---------- HELPER FUNCTIONS ----------
def clean_text(s):
    if not s:
        return ""
    s = re.sub(r'[\u200b\u00a0\u200c\u200d]', '', str(s))
    return s.strip()

def get_order_sheet_records():
    if not hasattr(get_order_sheet_records, "cache"):
        get_order_sheet_records.cache = None
        get_order_sheet_records.cache_time = None
    now = datetime.now()
    if get_order_sheet_records.cache is None or (now - get_order_sheet_records.cache_time).total_seconds() > 300:
        get_order_sheet_records.cache = order_sheet.get_all_records()
        get_order_sheet_records.cache_time = now
    return get_order_sheet_records.cache

def get_raw_records():
    return order_sheet.get_all_records()

def find_order_details(order_id: str):
    clean_input = clean_text(order_id)
    try:
        cell = order_sheet.find(clean_input, in_column=1)
        if cell is None:
            return None
        row = order_sheet.row_values(cell.row)
        return {
            "sto": row[1] if len(row) > 1 else "-",
            "wok": row[2] if len(row) > 2 else "-",
            "order_status": row[3] if len(row) > 3 else "N/A",
            "channel": row[4] if len(row) > 4 else "N/A",
            "fallout": row[5] if len(row) > 5 else "(Blank)",
            "salesforce": row[6] if len(row) > 6 else "N/A",
            "tanggal_complete": row[7] if len(row) > 7 else "-",
            "tanggal_input": row[8] if len(row) > 8 else "-",
            "sub_error": row[9] if len(row) > 9 else "-",
            "technician_notes": row[10] if len(row) > 10 else "-",
        }
    except Exception as e:
        logger.error(f"Error finding order {clean_input}: {e}")
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

def can_view_sales_report(telegram_id):
    _, subrole = get_user_role(telegram_id)
    return subrole in ["Supervisor", "Team Leader", "IT", "Manager"]

def can_view_summary(telegram_id):
    _, subrole = get_user_role(telegram_id)
    return subrole in ["Manager", "Supervisor", "IT"]

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

async def notify_approver(bot, user_id, name, role_group, subrole, wok="", sfid=""):
    if IT_TELEGRAM_ID:
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Approve", callback_data=f"approve_{user_id}_"),
             InlineKeyboardButton("❌ Reject", callback_data=f"reject_{user_id}_")]
        ])
        text = f"New registration pending:\nName: {name}\nRole: {role_group} - {subrole}"
        if wok:
            text += f"\nWOK: {wok}"
        if sfid:
            text += f"\nSFID: {sfid}"
        text += f"\nUser ID: {user_id}"
        await bot.send_message(chat_id=IT_TELEGRAM_ID, text=text, reply_markup=keyboard)

# ---------- MAIN MENU ----------
def get_main_menu_keyboard(telegram_id):
    _, subrole = get_user_role(telegram_id)
    buttons = [
        [InlineKeyboardButton("🔍 Cek Order", callback_data="menu_single_order")],
        [InlineKeyboardButton("📦 Cek Banyak Order", callback_data="menu_bulk_order")],
    ]
    if can_view_sales_report(telegram_id):
        buttons.insert(1, [InlineKeyboardButton("📊 Cek Laporan Sales", callback_data="menu_sales_report")])
    buttons.append([InlineKeyboardButton("📖 Panduan Pengguna", callback_data="menu_guide")])
    return InlineKeyboardMarkup(buttons)

async def show_main_menu(update: Update, user_id):
    if is_user_approved(user_id):
        text = "Selamat datang! Silakan pilih menu:"
        keyboard = get_main_menu_keyboard(user_id)
        await update.message.reply_text(text, reply_markup=keyboard)
    else:
        text = "Anda belum terdaftar. Silakan daftar terlebih dahulu."
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📝 Daftar", callback_data="menu_register")],
            [InlineKeyboardButton("📖 Panduan Pengguna", callback_data="menu_guide")]
        ])
        await update.message.reply_text(text, reply_markup=keyboard)

async def menu_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    user_id = update.effective_user.id
    if data == "menu_register":
        await query.message.delete()
        await register_start(update, context)
    elif data == "menu_single_order":
        await query.message.delete()
        await single_order_start(update, context)
    elif data == "menu_bulk_order":
        await query.message.delete()
        await bulk_order_start(update, context)
    elif data == "menu_sales_report":
        await query.message.delete()
        await sales_report_main(update, context)
    elif data == "menu_guide":
        await query.message.delete()
        await send_guide(update, user_id)
        await show_main_menu(update, user_id)

# ---------- COMBINED GUIDE ----------
async def send_guide(update: Update, user_id):
    if update.callback_query:
        msg = update.callback_query.message
        reply = msg.reply_text
    else:
        reply = update.message.reply_text

    role_group, subrole = get_user_role(user_id)
    if not role_group:
        text = (
            "📖 *Panduan Pengguna*\n\n"
            "Anda belum terdaftar. Silakan gunakan tombol *Daftar* di bawah untuk memulai registrasi.\n\n"
            "Setelah mendaftar, Anda harus menunggu persetujuan dari IT. Anda akan diberi tahu setelah disetujui.\n\n"
            "Jika Anda sudah terdaftar dan disetujui, gunakan /start untuk memeriksa Order ID.\n\n"
            "Untuk bantuan lebih lanjut, ketik /help."
        )
        await reply(text, parse_mode="Markdown")
        return
    if can_view_sales_report(user_id) or subrole in ["Manager", "Supervisor", "HSA", "IT"]:
        text = (
            f"📋 *Panduan Pengguna*\n\n"
            f"✅ Anda terdaftar sebagai {role_group} - {subrole}.\n\n"
            "🔍 *Cek Order (satu)*:\n"
            "Klik tombol 'Cek Order', lalu masukkan satu Order ID.\n\n"
            "📦 *Cek Banyak Order*:\n"
            "Klik tombol 'Cek Banyak Order', lalu masukkan 2 hingga 10 Order ID (dipisah spasi).\n\n"
            "📊 *Laporan Performa Sales*:\n"
            "Klik tombol 'Cek Laporan Sales' dan ikuti menu interaktif (tersedia untuk Supervisor, Team Leader, IT, Manager).\n\n"
            "📎 Untuk laporan penggunaan bot, gunakan perintah /report day|week|month|from to.\n\n"
            "Untuk daftar perintah lengkap, ketik /help."
        )
    else:
        text = (
            f"📋 *Panduan Pengguna*\n\n"
            f"✅ Anda terdaftar sebagai {role_group} - {subrole}.\n\n"
            "🔍 *Cek Order (satu)*:\n"
            "Klik tombol 'Cek Order', lalu masukkan satu Order ID.\n\n"
            "📦 *Cek Banyak Order*:\n"
            "Klik tombol 'Cek Banyak Order', lalu masukkan 2 hingga 10 Order ID (dipisah spasi).\n\n"
            "📊 *Laporan*:\n"
            "Laporan hanya tersedia untuk Supervisor, Manager, HSA, IT, dan Team Leader.\n\n"
            "Untuk daftar perintah lengkap, ketik /help."
        )
    await reply(text, parse_mode="Markdown")

# ---------- REGISTRATION FLOW ----------
async def register_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    if update.callback_query:
        query = update.callback_query
        await query.answer()
        send_func = query.message.reply_text
        user_id = query.from_user.id
    else:
        send_func = update.message.reply_text
        user_id = update.effective_user.id

    if is_user_approved(user_id):
        await send_func("Anda sudah terdaftar dan disetujui.")
        return ConversationHandler.END
    records = users_sheet.get_all_records()
    for r in records:
        if str(r.get("TelegramID", "")) == str(user_id) and r.get("ApprovalStatus") == "pending":
            await send_func("Anda sudah memiliki pendaftaran yang menunggu persetujuan. Mohon tunggu.")
            return ConversationHandler.END
    await send_func("Selamat datang! Mari daftar.\n\nMasukkan nama lengkap Anda:")
    return REG_NAME

async def reg_name(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["reg_name"] = update.message.text
    await update.message.reply_text("Masukkan alamat email Anda:")
    return REG_EMAIL

async def reg_email(update: Update, context: ContextTypes.DEFAULT_TYPE):
    context.user_data["reg_email"] = update.message.text
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Agency", callback_data="group_Agency")],
        [InlineKeyboardButton("Branch", callback_data="group_Branch")],
        [InlineKeyboardButton("Technician", callback_data="group_Technician")]
    ])
    await update.message.reply_text("Pilih grup peran Anda:", reply_markup=keyboard)
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
    else:
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Technician", callback_data="sub_TECH")],
            [InlineKeyboardButton("HSA", callback_data="sub_HSA")]
        ])
    await query.edit_message_text("Pilih sub-peran Anda:", reply_markup=keyboard)
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
    group = context.user_data["reg_role_group"]
    if group == "Agency":
        wok_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Manado-Talaud", callback_data="wok_Manado-Talaud")],
            [InlineKeyboardButton("Bitung-Minahasa", callback_data="wok_Bitung-Minahasa")],
            [InlineKeyboardButton("Bolaang Mongondow", callback_data="wok_Bolaang Mongondow")],
            [InlineKeyboardButton("Gorontalo", callback_data="wok_Gorontalo")]
        ])
        await query.edit_message_text("Pilih WOK (Working Area):", reply_markup=wok_keyboard)
        return REG_WOK
    else:
        await query.edit_message_text("Pendaftaran selesai. Menyimpan data...")
        return await save_registration(update, context, skip_wok_sfid=True)

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
    return await save_registration(update, context, skip_wok_sfid=False)

async def save_registration(update: Update, context: ContextTypes.DEFAULT_TYPE, skip_wok_sfid=False):
    if isinstance(update, Update) and update.callback_query:
        await update.callback_query.edit_message_text("Menyimpan data registrasi...")
        effective_user = update.callback_query.from_user
        send_message = update.callback_query.message.reply_text
    else:
        effective_user = update.effective_user
        send_message = update.message.reply_text

    user_id = effective_user.id
    name = context.user_data.get("reg_name")
    email = context.user_data.get("reg_email")
    role_group = context.user_data.get("reg_role_group")
    subrole = context.user_data.get("reg_subrole")
    wok = context.user_data.get("reg_wok", "") if not skip_wok_sfid else ""
    sfid = context.user_data.get("reg_sfid", "") if not skip_wok_sfid else ""

    status = "pending"
    approved_by = ""
    approved_at = ""
    reply_text = "Pendaftaran dikirim. Anda akan diberi tahu setelah disetujui oleh IT."

    await notify_approver(update.get_bot(), user_id, name, role_group, subrole, wok, sfid)

    users_sheet.append_row([
        user_id, name, email, role_group, subrole, status,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S"), approved_by, approved_at,
        wok, sfid
    ])

    await send_message(reply_text)
    return ConversationHandler.END

# ---------- APPROVAL HANDLERS ----------
async def pending(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    _, subrole = get_user_role(user_id)
    if not (is_user_approved(user_id) and subrole == "IT"):
        await update.message.reply_text("Hanya pengguna IT yang dapat melihat registrasi tertunda.")
        return

    records = users_sheet.get_all_records()
    pending_users = []
    for idx, row in enumerate(records, start=2):
        if row.get("ApprovalStatus") == "pending":
            pending_users.append((idx, row))
    if not pending_users:
        await update.message.reply_text("Tidak ada registrasi tertunda.")
        return

    for idx, row in pending_users:
        name = row.get("Name")
        role_group = row.get("RoleGroup")
        subrole = row.get("SubRole")
        wok = row.get("WOK", "N/A") if role_group == "Agency" else "N/A"
        telegram_id = row.get("TelegramID")
        text = (
            f"Registrasi tertunda:\n"
            f"Nama: {name}\n"
            f"Peran: {role_group} - {subrole}\n"
            f"WOK: {wok}\n"
            f"ID Pengguna: {telegram_id}"
        )
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Setujui", callback_data=f"approve_{telegram_id}_{idx}"),
             InlineKeyboardButton("❌ Tolak", callback_data=f"reject_{telegram_id}_{idx}")]
        ])
        await update.message.reply_text(text, reply_markup=keyboard)

async def approval_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    parts = data.split("_", 2)
    if len(parts) != 3:
        logger.error(f"Invalid callback data: {data}")
        await query.edit_message_text("Terjadi kesalahan. Silakan coba lagi.")
        return
    action = parts[0]
    try:
        target_id = int(parts[1])
        row_index = int(parts[2])
    except ValueError:
        logger.error(f"Invalid numbers in callback data: {parts}")
        await query.edit_message_text("Terjadi kesalahan. Silakan coba lagi.")
        return

    if action == "approve":
        users_sheet.update_cell(row_index, 6, "approved")
        users_sheet.update_cell(row_index, 8, str(update.effective_user.id))
        users_sheet.update_cell(row_index, 9, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        await context.bot.send_message(chat_id=target_id, text="Pendaftaran Anda telah disetujui! Anda sekarang dapat menggunakan tombol menu.")
        await query.edit_message_text(f"✅ Pengguna {target_id} disetujui.")
    elif action == "reject":
        users_sheet.update_cell(row_index, 6, "rejected")
        await context.bot.send_message(chat_id=target_id, text="Pendaftaran Anda ditolak. Silakan coba /register lagi.")
        await query.edit_message_text(f"❌ Pengguna {target_id} ditolak.")

# ---------- SINGLE ORDER ----------
async def single_order_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    if not is_user_approved(user_id):
        await query.message.reply_text("Anda belum terdaftar atau belum disetujui. Silakan gunakan tombol Daftar.")
        return ConversationHandler.END
    await query.message.reply_text("Masukkan Order ID (contoh: AOs326032509275620607db90):")
    return WAITING_FOR_SINGLE_ORDER

async def receive_single_order(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    raw_input = update.message.text
    order_id = clean_text(raw_input)
    data = find_order_details(order_id)
    if data is None:
        last_update = get_last_update_time()
        error_msg = (
            f"❌ Maaf Order ID Tidak Ditemukan atau Belum Terupdate\n"
            f"📅 Last Update Data: {last_update}\n\n"
            f"Silakan coba lagi dengan Order ID lain."
        )
        await update.message.reply_text(error_msg)
        return WAITING_FOR_SINGLE_ORDER
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
        f"👨🏼‍🔧 Technician Notes: {data['technician_notes']}"
    )
    await update.message.reply_text(reply)
    return ConversationHandler.END

# ---------- BULK ORDER ----------
async def bulk_order_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    if not is_user_approved(user_id):
        await query.message.reply_text("Anda belum terdaftar atau belum disetujui. Silakan gunakan tombol Daftar.")
        return ConversationHandler.END
    await query.message.reply_text(
        "Masukkan 2 hingga 10 Order ID dalam satu pesan.\n"
        "Pisahkan dengan spasi.\n\n"
        "Contoh: AOi123 AOi456 AOi789"
    )
    return BULK_AWAITING_IDS

async def process_bulk_input(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    raw_text = update.message.text.strip()
    order_ids = re.split(r'\s+', raw_text)
    order_ids = [clean_text(oid) for oid in order_ids if oid]
    if len(order_ids) < 2 or len(order_ids) > 10:
        await update.message.reply_text("Jumlah Order ID harus antara 2 dan 10. Silakan coba lagi.")
        return BULK_AWAITING_IDS
    found = []
    not_found = []
    for oid in order_ids:
        data = find_order_details(oid)
        if data:
            found.append((oid, data))
            log_usage(user_id, oid)
        else:
            not_found.append(oid)
    for oid, data in found:
        reply = (
            f"✅ Order ID: {oid}\n"
            f"📠 STO: {data['sto']}\n"
            f"🪪 WOK: {data['wok']}\n"
            f"⚙️ Order Status: {data['order_status']}\n"
            f"📢 Channel Name: {data['channel']}\n"
            f"⚠️ Fallout Reason: {data['fallout']}\n"
            f"👤 Salesforce: {data['salesforce']}\n"
            f"📅 Tanggal Complete: {data['tanggal_complete']}\n"
            f"📅 Tanggal Input: {data['tanggal_input']}\n"
            f"🧠 Sub Error Code: {data['sub_error']}\n"
            f"👨🏼‍🔧 Technician Notes: {data['technician_notes']}"
        )
        await update.message.reply_text(reply)
    if not_found:
        await update.message.reply_text("❌ Order ID tidak ditemukan:\n" + "\n".join(not_found))
    await update.message.reply_text(f"✅ Selesai. {len(found)} dari {len(order_ids)} Order ID berhasil ditemukan.")
    return ConversationHandler.END

# ---------- SALES REPORT MAIN MENU ----------
async def sales_report_main(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    if not (is_user_approved(user_id) and can_view_sales_report(user_id)):
        await query.message.reply_text("Anda tidak memiliki akses ke laporan sales.")
        return ConversationHandler.END

    keyboard = [
        [InlineKeyboardButton("📊 Laporan Detail (per WOK)", callback_data="sales_detail")],
    ]
    if can_view_summary(user_id):
        keyboard.append([InlineKeyboardButton("📈 Ringkasan (per WOK & Channel)", callback_data="sales_summary")])
    keyboard.append([InlineKeyboardButton("🔙 Kembali ke Menu Utama", callback_data="sales_back")])

    await query.message.reply_text("Pilih jenis laporan:", reply_markup=InlineKeyboardMarkup(keyboard))
    return SALES_WOK

async def sales_choose(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "sales_detail":
        await query.message.delete()
        wok_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("MANADO TALAUD", callback_data="wok_MANADO TALAUD")],
            [InlineKeyboardButton("BOLAANG MONGONDOW", callback_data="wok_BOLAANG MONGONDOW")],
            [InlineKeyboardButton("GORONTALO - PAHUWATO", callback_data="wok_GORONTALO - PAHUWATO")],
            [InlineKeyboardButton("BITUNG MINAHASA", callback_data="wok_BITUNG MINAHASA")]
        ])
        await query.message.reply_text("Pilih WOK:", reply_markup=wok_keyboard)
        return DETAIL_WOK   # new state for WOK selection
    elif data == "sales_summary":
        await query.message.delete()
        current_year = datetime.now().year
        year_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton(str(current_year - 1), callback_data=f"sumyear_{current_year-1}"),
             InlineKeyboardButton(str(current_year), callback_data=f"sumyear_{current_year}")]
        ])
        await query.message.reply_text("Pilih tahun:", reply_markup=year_keyboard)
        return SUMMARY_YEAR
    elif data == "sales_back":
        await query.message.delete()
        await show_main_menu(update, update.effective_user.id)
        return ConversationHandler.END
    return ConversationHandler.END

# ---------- DETAILED REPORT WOK HANDLER ----------
async def detail_wok_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    wok = query.data.split("_", 1)[1]
    context.user_data["report_wok"] = wok

    user_id = update.effective_user.id
    _, subrole = get_user_role(user_id)
    context.user_data["report_subrole"] = subrole

    if subrole == "Team Leader":
        context.user_data["report_channel"] = "AGENCY"
        current_year = datetime.now().year
        year_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton(str(current_year - 1), callback_data=f"year_{current_year-1}"),
             InlineKeyboardButton(str(current_year), callback_data=f"year_{current_year}")]
        ])
        await query.edit_message_text("Pilih tahun:", reply_markup=year_keyboard)
        return SALES_YEAR
    else:
        channel_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("B2B2C&OTHERS", callback_data="chan_B2B2C&OTHERS")],
            [InlineKeyboardButton("AGENCY", callback_data="chan_AGENCY")],
            [InlineKeyboardButton("GRAPARI", callback_data="chan_GRAPARI")],
            [InlineKeyboardButton("SOBI AFFILIATE", callback_data="chan_SOBI AFFILIATE")],
            [InlineKeyboardButton("WEB&APP", callback_data="chan_WEB&APP")]
        ])
        await query.edit_message_text("Pilih Channel:", reply_markup=channel_keyboard)
        return SALES_CHANNEL

# ---------- CHANNEL, YEAR, MONTH HANDLERS (same as before) ----------
async def sales_channel_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    channel = query.data.split("_", 1)[1]
    context.user_data["report_channel"] = channel
    current_year = datetime.now().year
    year_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton(str(current_year - 1), callback_data=f"year_{current_year-1}"),
         InlineKeyboardButton(str(current_year), callback_data=f"year_{current_year}")]
    ])
    await query.edit_message_text("Pilih tahun:", reply_markup=year_keyboard)
    return SALES_YEAR

async def sales_year_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year = int(query.data.split("_")[1])
    context.user_data["report_year"] = year
    month_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Januari", callback_data="month_1"),
         InlineKeyboardButton("Februari", callback_data="month_2"),
         InlineKeyboardButton("Maret", callback_data="month_3")],
        [InlineKeyboardButton("April", callback_data="month_4"),
         InlineKeyboardButton("Mei", callback_data="month_5"),
         InlineKeyboardButton("Juni", callback_data="month_6")],
        [InlineKeyboardButton("Juli", callback_data="month_7"),
         InlineKeyboardButton("Agustus", callback_data="month_8"),
         InlineKeyboardButton("September", callback_data="month_9")],
        [InlineKeyboardButton("Oktober", callback_data="month_10"),
         InlineKeyboardButton("November", callback_data="month_11"),
         InlineKeyboardButton("Desember", callback_data="month_12")]
    ])
    await query.edit_message_text("Pilih bulan:", reply_markup=month_keyboard)
    return SALES_MONTH

async def sales_month_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    # Detailed report logic (same as before)
    query = update.callback_query
    await query.answer()
    month_num = int(query.data.split("_")[1])
    month_names = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                   "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    month_name = month_names[month_num - 1]
    year = context.user_data.get("report_year", datetime.now().year)
    wok = context.user_data["report_wok"]
    channel = context.user_data["report_channel"]
    subrole = context.user_data["report_subrole"]

    records = get_order_sheet_records()
    filtered = []
    for rec in records:
        rec_wok = rec.get("WOK", "")
        rec_channel = rec.get("Channel Name", "")
        rec_tanggal_input = rec.get("Tanggal Input", "")
        if not rec_tanggal_input:
            continue
        try:
            if '-' in rec_tanggal_input:
                date_part = rec_tanggal_input.split()[0]
                y, m, d = map(int, date_part.split('-'))
            else:
                date_part = rec_tanggal_input.split()[0]
                d, m, y = map(int, date_part.split('/'))
            if m == month_num and y == year and rec_wok == wok and rec_channel == channel:
                filtered.append(rec)
        except:
            continue

    if not filtered:
        await query.edit_message_text(f"Tidak ada data untuk WOK: {wok}, Channel: {channel}, Tahun: {year}, Bulan: {month_name}.")
        return ConversationHandler.END

    await query.delete_message()
    processing_msg = await query.message.reply_text("⏳ Sedang memproses data...")

    aggregate_only_roles = ["Manager", "Supervisor", "Inputters", "IT"]

    if channel == "AGENCY" and subrole in aggregate_only_roles:
        status_counts = {s: 0 for s in ALL_STATUSES}
        status_counts["OTHER"] = 0
        total = 0
        for rec in filtered:
            total += 1
            status = rec.get("Status Order", "").upper().strip()
            if status in status_counts:
                status_counts[status] += 1
            else:
                status_counts["OTHER"] += 1
        lines = [f"📢 Channel: {channel}", f"📅 Tahun: {year}", f"📅 Bulan: {month_name.upper()}", f"📍 WOK: {wok}", ""]
        lines.append(f"📊 Total Order: {total}")
        lines.append("🔍 Rincian Status:")
        for s in ALL_STATUSES:
            if status_counts[s] > 0:
                if s == "COMPLETED":
                    lines.append(f"   ✅ {s}: {status_counts[s]}")
                elif s in ["FALLOUT", "CANCELLED", "CANCELLED_SLA", "CANCEL_OSM_COMPLETED", "CANCEL_ORDER_INPROGRESS"]:
                    lines.append(f"   ❌ {s}: {status_counts[s]}")
                else:
                    lines.append(f"   🔄 {s}: {status_counts[s]}")
        if status_counts["OTHER"] > 0:
            lines.append(f"   ❓ OTHER: {status_counts['OTHER']}")
        await processing_msg.edit_text("\n".join(lines))
        return ConversationHandler.END

    if channel == "AGENCY":
        sales_dict = {}
        for rec in filtered:
            sf = rec.get("SalesForce", "").strip()
            if not sf:
                continue
            status = rec.get("Status Order", "").upper().strip()
            if sf not in sales_dict:
                sales_dict[sf] = {"total": 0}
                for s in ALL_STATUSES:
                    sales_dict[sf][s] = 0
                sales_dict[sf]["orders"] = []
            sales_dict[sf]["total"] += 1
            if status in sales_dict[sf]:
                sales_dict[sf][status] += 1
            else:
                if "OTHER" not in sales_dict[sf]:
                    sales_dict[sf]["OTHER"] = 0
                sales_dict[sf]["OTHER"] += 1
            if subrole == "Team Leader":
                order_id = rec.get("Order ID", "")
                sales_dict[sf]["orders"].append((order_id, status))

        sorted_sf = sorted(sales_dict.items(), key=lambda x: x[1]["total"], reverse=True)
        header = f"📢 Channel: {channel}\n📅 Tahun: {year}\n📅 Bulan: {month_name.upper()}\n📍 WOK: {wok}"
        await processing_msg.edit_text(header)

        for sf, data in sorted_sf:
            if subrole == "Team Leader":
                lines = [f"👤 {sf}"]
                lines.append(f"   📦 Total Order: {data['total']}")
                status_summary = []
                for s in ALL_STATUSES:
                    count = data.get(s, 0)
                    if count > 0:
                        if s == "COMPLETED":
                            status_summary.append(f"✅ {s}: {count}")
                        elif s in ["FALLOUT", "CANCELLED", "CANCELLED_SLA", "CANCEL_OSM_COMPLETED", "CANCEL_ORDER_INPROGRESS"]:
                            status_summary.append(f"❌ {s}: {count}")
                        else:
                            status_summary.append(f"🔄 {s}: {count}")
                if data.get("OTHER", 0) > 0:
                    status_summary.append(f"❓ OTHER: {data['OTHER']}")
                if status_summary:
                    lines.append("   📊 Status:")
                    for stat_line in status_summary:
                        lines.append(f"      {stat_line}")
                lines.append("   🧾 Daftar Order:")
                order_lines = [f"      • {oid} ({stat})" for oid, stat in data["orders"]]
                base_text = "\n".join(lines[:-1]) + "\n   🧾 Daftar Order:\n"
                current_chunk = []
                current_len = len(base_text) + 10
                for order_line in order_lines:
                    if current_len + len(order_line) + 1 > 4000:
                        chunk_text = base_text + "\n".join(current_chunk)
                        await query.message.reply_text(chunk_text)
                        current_chunk = []
                        current_len = len(base_text) + 10
                    current_chunk.append(order_line)
                    current_len += len(order_line) + 1
                if current_chunk:
                    chunk_text = base_text + "\n".join(current_chunk)
                    await query.message.reply_text(chunk_text)
            else:
                lines = [f"👤 {sf}", f"   🔢 Total Order: {data['total']}"]
                status_summary = []
                for s in ALL_STATUSES:
                    count = data.get(s, 0)
                    if count > 0:
                        if s == "COMPLETED":
                            status_summary.append(f"✅ {s}: {count}")
                        elif s in ["FALLOUT", "CANCELLED", "CANCELLED_SLA", "CANCEL_OSM_COMPLETED", "CANCEL_ORDER_INPROGRESS"]:
                            status_summary.append(f"❌ {s}: {count}")
                        else:
                            status_summary.append(f"🔄 {s}: {count}")
                if data.get("OTHER", 0) > 0:
                    status_summary.append(f"❓ OTHER: {data['OTHER']}")
                if status_summary:
                    lines.append("   📊 Rincian Status:")
                    for stat_line in status_summary:
                        lines.append(f"      {stat_line}")
                await query.message.reply_text("\n".join(lines))
        await query.message.reply_text("✅ Selesai. Terima kasih.")
    else:
        status_counts = {s: 0 for s in ALL_STATUSES}
        status_counts["OTHER"] = 0
        total = 0
        for rec in filtered:
            total += 1
            status = rec.get("Status Order", "").upper().strip()
            if status in status_counts:
                status_counts[status] += 1
            else:
                status_counts["OTHER"] += 1
        lines = [f"📢 Channel: {channel}", f"📅 Tahun: {year}", f"📅 Bulan: {month_name.upper()}", f"📍 WOK: {wok}", ""]
        lines.append(f"📊 Total Order: {total}")
        lines.append("🔍 Rincian Status:")
        for s in ALL_STATUSES:
            if status_counts[s] > 0:
                if s == "COMPLETED":
                    lines.append(f"   ✅ {s}: {status_counts[s]}")
                elif s in ["FALLOUT", "CANCELLED", "CANCELLED_SLA", "CANCEL_OSM_COMPLETED", "CANCEL_ORDER_INPROGRESS"]:
                    lines.append(f"   ❌ {s}: {status_counts[s]}")
                else:
                    lines.append(f"   🔄 {s}: {status_counts[s]}")
        if status_counts["OTHER"] > 0:
            lines.append(f"   ❓ OTHER: {status_counts['OTHER']}")
        await processing_msg.edit_text("\n".join(lines))
    return ConversationHandler.END

# ---------- SUMMARY REPORT HANDLERS (unchanged from last working version) ----------
async def summary_year_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year = int(query.data.split("_")[1])
    context.user_data["summary_year"] = year
    month_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Januari", callback_data="summonth_1"),
         InlineKeyboardButton("Februari", callback_data="summonth_2"),
         InlineKeyboardButton("Maret", callback_data="summonth_3")],
        [InlineKeyboardButton("April", callback_data="summonth_4"),
         InlineKeyboardButton("Mei", callback_data="summonth_5"),
         InlineKeyboardButton("Juni", callback_data="summonth_6")],
        [InlineKeyboardButton("Juli", callback_data="summonth_7"),
         InlineKeyboardButton("Agustus", callback_data="summonth_8"),
         InlineKeyboardButton("September", callback_data="summonth_9")],
        [InlineKeyboardButton("Oktober", callback_data="summonth_10"),
         InlineKeyboardButton("November", callback_data="summonth_11"),
         InlineKeyboardButton("Desember", callback_data="summonth_12")]
    ])
    await query.edit_message_text("Pilih bulan:", reply_markup=month_keyboard)
    return SUMMARY_MONTH

async def summary_month_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    month_num = int(query.data.split("_")[1])
    month_names = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                   "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    month_name = month_names[month_num - 1]
    year = context.user_data.get("summary_year", datetime.now().year)

    records = get_raw_records()
    prev_month = month_num - 1 if month_num > 1 else 12
    prev_year = year if month_num > 1 else year - 1

    def extract_date(date_str):
        if not date_str:
            return None, None
        try:
            if '-' in date_str:
                parts = date_str.split()[0].split('-')
                if len(parts) == 3:
                    return int(parts[0]), int(parts[1])
            parts = date_str.split()[0].split('/')
            if len(parts) == 3:
                return int(parts[2]), int(parts[1])
        except:
            pass
        return None, None

    wok_list = ["MANADO TALAUD", "BOLAANG MONGONDOW", "GORONTALO - PAHUWATO", "BITUNG MINAHASA"]
    wok_stats = {}
    total_input = 0
    total_completed = 0
    total_fallout = 0

    for rec in records:
        wok = rec.get("WOK", "")
        if wok not in wok_list:
            continue
        tgl = rec.get("Tanggal Input", "")
        y, m = extract_date(tgl)
        if y is None or m is None:
            continue
        status = (rec.get("Status Order", "") or "").strip().upper()
        if y == year and m == month_num:
            total_input += 1
            if status == "COMPLETED":
                total_completed += 1
            elif status == "FALLOUT":
                total_fallout += 1
            if wok not in wok_stats:
                wok_stats[wok] = {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0}
            wok_stats[wok]["input"] += 1
            if status == "COMPLETED":
                wok_stats[wok]["completed"] += 1
            elif status == "FALLOUT":
                wok_stats[wok]["fallout"] += 1
        if y == prev_year and m == prev_month:
            if wok not in wok_stats:
                wok_stats[wok] = {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0}
            wok_stats[wok]["prev_input"] += 1

    table_wok = []
    for wok in wok_list:
        stats = wok_stats.get(wok, {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0})
        inp = stats["input"]
        comp = stats["completed"]
        flt = stats["fallout"]
        prev = stats["prev_input"]
        completion = (comp / inp * 100) if inp > 0 else 0
        contrib = (inp / total_input * 100) if total_input > 0 else 0
        mom = ((inp - prev) / prev * 100) if prev > 0 else (100 if inp > 0 else 0)
        table_wok.append({
            "wok": wok,
            "input": inp,
            "completed": comp,
            "fallout": flt,
            "completion": completion,
            "contrib": contrib,
            "mom": mom
        })

    total_completion = (total_completed / total_input * 100) if total_input > 0 else 0
    total_prev = sum(stats.get("prev_input", 0) for stats in wok_stats.values())
    total_mom = ((total_input - total_prev) / total_prev * 100) if total_prev > 0 else (100 if total_input > 0 else 0)

    lines1 = ["📊 *RINGKASAN PER WOK*", f"📅 {month_name.upper()} {year}", ""]
    lines1.append("```")
    lines1.append(f"{'WOK':<22} {'Input':>7} {'Cmpl':>6} {'Flt':>5} {'IO/PS':>6} {'Kontrib':>8} {'MoM':>6}")
    lines1.append("-" * 70)
    for row in table_wok:
        lines1.append(f"{row['wok'][:20]:<20} {row['input']:>7} {row['completed']:>6} {row['fallout']:>5} {row['completion']:>5.1f}% {row['contrib']:>7.1f}% {row['mom']:>5.1f}%")
    lines1.append("-" * 70)
    lines1.append(f"{'TOTAL':<20} {total_input:>7} {total_completed:>6} {total_fallout:>5} {total_completion:>5.1f}% {'100.0':>7}% {total_mom:>5.1f}%")
    lines1.append("```")
    await query.message.reply_text("\n".join(lines1), parse_mode="Markdown")

    # Per channel summary (same as before)
    channels = ["B2B2C&OTHERS", "AGENCY", "GRAPARI", "SOBI AFFILIATE", "WEB&APP"]
    chan_stats = {}
    total_chan_input = 0
    total_chan_completed = 0
    total_chan_fallout = 0

    for rec in records:
        channel = rec.get("Channel Name", "")
        if channel not in channels:
            continue
        tgl = rec.get("Tanggal Input", "")
        y, m = extract_date(tgl)
        if y is None or m is None:
            continue
        status = (rec.get("Status Order", "") or "").strip().upper()
        if y == year and m == month_num:
            total_chan_input += 1
            if status == "COMPLETED":
                total_chan_completed += 1
            elif status == "FALLOUT":
                total_chan_fallout += 1
            if channel not in chan_stats:
                chan_stats[channel] = {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0}
            chan_stats[channel]["input"] += 1
            if status == "COMPLETED":
                chan_stats[channel]["completed"] += 1
            elif status == "FALLOUT":
                chan_stats[channel]["fallout"] += 1
        if y == prev_year and m == prev_month:
            if channel not in chan_stats:
                chan_stats[channel] = {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0}
            chan_stats[channel]["prev_input"] += 1

    table_chan = []
    for ch in channels:
        stats = chan_stats.get(ch, {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0})
        inp = stats["input"]
        comp = stats["completed"]
        flt = stats["fallout"]
        prev = stats["prev_input"]
        completion = (comp / inp * 100) if inp > 0 else 0
        contrib = (inp / total_chan_input * 100) if total_chan_input > 0 else 0
        mom = ((inp - prev) / prev * 100) if prev > 0 else (100 if inp > 0 else 0)
        table_chan.append({
            "channel": ch,
            "input": inp,
            "completed": comp,
            "fallout": flt,
            "completion": completion,
            "contrib": contrib,
            "mom": mom
        })

    total_chan_completion = (total_chan_completed / total_chan_input * 100) if total_chan_input > 0 else 0
    total_chan_prev = sum(stats.get("prev_input", 0) for stats in chan_stats.values())
    total_chan_mom = ((total_chan_input - total_chan_prev) / total_chan_prev * 100) if total_chan_prev > 0 else (100 if total_chan_input > 0 else 0)

    lines2 = ["📊 *RINGKASAN PER CHANNEL*", f"📅 {month_name.upper()} {year}", ""]
    lines2.append("```")
    lines2.append(f"{'Channel':<18} {'Input':>7} {'Cmpl':>6} {'Flt':>5} {'IO/PS':>6} {'Kontrib':>8} {'MoM':>6}")
    lines2.append("-" * 70)
    for row in table_chan:
        lines2.append(f"{row['channel'][:16]:<16} {row['input']:>7} {row['completed']:>6} {row['fallout']:>5} {row['completion']:>5.1f}% {row['contrib']:>7.1f}% {row['mom']:>5.1f}%")
    lines2.append("-" * 70)
    lines2.append(f"{'TOTAL':<16} {total_chan_input:>7} {total_chan_completed:>6} {total_chan_fallout:>5} {total_chan_completion:>5.1f}% {'100.0':>7}% {total_chan_mom:>5.1f}%")
    lines2.append("```")
    await query.message.reply_text("\n".join(lines2), parse_mode="Markdown")
    return ConversationHandler.END

# ---------- USAGE REPORT ----------
async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not can_view_reports(user_id):
        await update.message.reply_text("Anda tidak memiliki izin untuk melihat laporan.")
        return
    args = context.args
    if not args:
        await update.message.reply_text("Penggunaan: /report [day|week|month|YYYY-MM-DD YYYY-MM-DD]")
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
            await update.message.reply_text("Untuk rentang kustom: /report YYYY-MM-DD YYYY-MM-DD")
            return
        try:
            start = datetime.strptime(args[0], "%Y-%m-%d")
            end = datetime.strptime(args[1], "%Y-%m-%d") + timedelta(days=1)
        except ValueError:
            await update.message.reply_text("Format tanggal salah. Gunakan YYYY-MM-DD.")
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
        await update.message.reply_text("Tidak ada data penggunaan dalam periode ini.")
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
        await update.message.reply_document(document=output, filename=f"report_{args[0]}.csv", caption=f"Laporan untuk {args[0]}")
    else:
        lines = [f"📊 Laporan untuk {args[0]}\n"]
        for uid, data in summary.items():
            duration = (data["last"] - data["first"]).total_seconds() / 60 if data["first"] and data["last"] else 0
            lines.append(f"👤 {data['name']} ({data['role']}) - {data['count']} pencarian - Durasi: {duration:.0f} menit")
        await update.message.reply_text("\n".join(lines))

# ---------- GENERAL COMMANDS ----------
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Perintah dibatalkan. Kembali ke menu utama.")
    await show_main_menu(update, update.effective_user.id)
    return ConversationHandler.END

async def ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("pong")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_guide(update, update.effective_user.id)

async def guide_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_guide(update, update.effective_user.id)

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_main_menu(update, update.effective_user.id)

# ---------- MAIN ----------
def main():
    if not TELEGRAM_BOT_TOKEN:
        raise ValueError("TELEGRAM_BOT_TOKEN not set")
    if not GOOGLE_CREDENTIALS_JSON:
        raise ValueError("GOOGLE_CREDENTIALS_JSON not set")
    if not BOT_DATA_SHEET_NAME:
        raise ValueError("BOT_DATA_SHEET_NAME not set")

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
        entry_points=[CommandHandler("register", register_start), CallbackQueryHandler(register_start, pattern="^menu_register$")],
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

    # Sales report conversation
    sales_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(sales_report_main, pattern="^menu_sales_report$")],
        states={
            SALES_WOK: [CallbackQueryHandler(sales_choose, pattern="^(sales_detail|sales_summary|sales_back)$")],
            DETAIL_WOK: [CallbackQueryHandler(detail_wok_selected, pattern="^wok_")],
            SALES_CHANNEL: [CallbackQueryHandler(sales_channel_selected, pattern="^chan_")],
            SALES_YEAR: [CallbackQueryHandler(sales_year_selected, pattern="^year_")],
            SALES_MONTH: [CallbackQueryHandler(sales_month_selected, pattern="^month_")],
            SUMMARY_YEAR: [CallbackQueryHandler(summary_year_selected, pattern="^sumyear_")],
            SUMMARY_MONTH: [CallbackQueryHandler(summary_month_selected, pattern="^summonth_")],
        },
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(sales_conv)

    # Single order conversation
    single_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(single_order_start, pattern="^menu_single_order$")],
        states={WAITING_FOR_SINGLE_ORDER: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_single_order)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(single_conv)

    # Bulk order conversation
    bulk_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(bulk_order_start, pattern="^menu_bulk_order$")],
        states={BULK_AWAITING_IDS: [MessageHandler(filters.TEXT & ~filters.COMMAND, process_bulk_input)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(bulk_conv)

    # Main menu callback
    app.add_handler(CallbackQueryHandler(menu_callback, pattern="^menu_"))

    # Fallback slash commands
    app.add_handler(CommandHandler("start", start_command))
    app.add_handler(CommandHandler("guide", guide_command))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("ping", ping))

    logger.info("Bot is polling...")
    app.run_polling()

if __name__ == "__main__":
    main()
