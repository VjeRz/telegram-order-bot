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
DETAIL_WOK = 21
SALES_CHANNEL = 22
SALES_YEAR = 23
SALES_MONTH = 24
SUMMARY_YEAR = 30
SUMMARY_MONTH = 31
TEAM_LEADER_OPTION = 32
REPORT_OPTION = 40
GRAPARI_STO_YEAR = 50
GRAPARI_STO_MONTH = 51

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

def get_last_order_date():
    """Return the latest Tanggal Input from the order sheet as DD/MM/YYYY HH:MM,
       or current datetime if none. Cached for 5 minutes."""
    if not hasattr(get_last_order_date, "cache"):
        get_last_order_date.cache = None
        get_last_order_date.cache_time = None
    now = datetime.now()
    if get_last_order_date.cache is None or (now - get_last_order_date.cache_time).total_seconds() > 300:
        try:
            records = get_raw_records()
            latest_dt = None
            for rec in records:
                tgl = rec.get("Tanggal Input", "")
                if not tgl:
                    continue
                tgl_clean = clean_text(tgl)
                for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
                    try:
                        dt = datetime.strptime(tgl_clean, fmt)
                        if latest_dt is None or dt > latest_dt:
                            latest_dt = dt
                        break
                    except:
                        continue
            if latest_dt:
                get_last_order_date.cache = latest_dt.strftime("%d/%m/%Y %H:%M")
            else:
                get_last_order_date.cache = datetime.now().strftime("%d/%m/%Y %H:%M")
            get_last_order_date.cache_time = now
        except Exception as e:
            logger.error(f"Error getting last order datetime: {e}")
            get_last_order_date.cache = datetime.now().strftime("%d/%m/%Y %H:%M")
            get_last_order_date.cache_time = now
    return get_last_order_date.cache

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
            "paket": row[11] if len(row) > 11 else "-",
        }
    except Exception as e:
        logger.error(f"Error finding order {clean_input}: {e}")
        return None

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
    return subrole in ["Supervisor", "Team Leader", "IT", "Manager", "Team Leader Grapari", "CS Grapari"]

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
    if can_view_reports(telegram_id):
        buttons.append([InlineKeyboardButton("📊 Laporan Pengguna Bot", callback_data="menu_usage_report")])
    buttons.append([InlineKeyboardButton("📖 Panduan Pengguna", callback_data="menu_guide")])
    return InlineKeyboardMarkup(buttons)

async def show_main_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    """Safely send main menu using context.bot.send_message."""
    user_id = update.effective_user.id
    bot = context.bot

    if is_user_approved(user_id):
        text = "Selamat datang! Silakan pilih menu:"
        keyboard = get_main_menu_keyboard(user_id)
        await bot.send_message(chat_id=user_id, text=text, reply_markup=keyboard)
    else:
        text = "Anda belum terdaftar. Silakan daftar terlebih dahulu."
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📝 Daftar", callback_data="menu_register")],
            [InlineKeyboardButton("📖 Panduan Pengguna", callback_data="menu_guide")]
        ])
        await bot.send_message(chat_id=user_id, text=text, reply_markup=keyboard)

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
    elif data == "menu_usage_report":
        await query.message.delete()
        await usage_report_menu(update, context)
    elif data == "menu_guide":
        await query.message.delete()
        await send_guide(update, user_id, context)
        await show_main_menu(update, context)

# ---------- USAGE REPORT MENU ----------
async def usage_report_menu(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    if not can_view_reports(user_id):
        await query.message.reply_text("Anda tidak memiliki izin untuk melihat laporan penggunaan.")
        return ConversationHandler.END
    keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("📆 Harian", callback_data="report_day")],
        [InlineKeyboardButton("📆 Mingguan", callback_data="report_week")],
        [InlineKeyboardButton("📆 Bulanan", callback_data="report_month")],
        [InlineKeyboardButton("🔙 Kembali ke Menu Utama", callback_data="report_back")]
    ])
    await query.message.reply_text("Pilih periode laporan:", reply_markup=keyboard)
    return REPORT_OPTION

async def generate_report(update: Update, message, user_id, args):
    if not can_view_reports(user_id):
        await message.reply_text("Anda tidak memiliki izin untuk melihat laporan.")
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
        await message.reply_text("Periode tidak dikenal.")
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
        await message.reply_text(f"Tidak ada data penggunaan untuk periode {args[0]}.")
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
    if args[0] == "month":
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(["TelegramID", "Name", "SubRole", "Number of lookups", "First lookup", "Last lookup", "Duration (minutes)"])
        for uid, data in summary.items():
            duration = (data["last"] - data["first"]).total_seconds() / 60 if data["first"] and data["last"] else 0
            writer.writerow([uid, data["name"], data["role"], data["count"], data["first"].strftime("%Y-%m-%d %H:%M:%S") if data["first"] else "", data["last"].strftime("%Y-%m-%d %H:%M:%S") if data["last"] else "", round(duration, 2)])
        output.seek(0)
        await message.reply_document(document=output, filename=f"report_{args[0]}.csv", caption=f"Laporan untuk {args[0]}")
    else:
        lines = [f"📊 Laporan untuk {args[0]}\n"]
        for uid, data in summary.items():
            duration = (data["last"] - data["first"]).total_seconds() / 60 if data["first"] and data["last"] else 0
            lines.append(f"👤 {data['name']} ({data['role']}) - {data['count']} pencarian - Durasi: {duration:.0f} menit")
        await message.reply_text("\n".join(lines))

async def report_option_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    if data == "report_back":
        await query.message.delete()
        await show_main_menu(update, context)
        return ConversationHandler.END
    period = data.split("_")[1]
    user_id = update.effective_user.id
    await generate_report(update, query.message, user_id, [period])
    await query.delete_message()
    return ConversationHandler.END

# ---------- COMBINED GUIDE ----------
async def send_guide(update: Update, user_id, context: ContextTypes.DEFAULT_TYPE = None):
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
            "Klik tombol 'Cek Laporan Sales' dan ikuti menu interaktif (tersedia untuk Supervisor, Team Leader, IT, Manager, Team Leader Grapari, CS Grapari).\n\n"
            "📊 *Laporan Pengguna Bot*:\n"
            "Klik tombol 'Laporan Pengguna Bot' lalu pilih Harian, Mingguan, atau Bulanan.\n\n"
            "📎 Untuk laporan penggunaan bot, alternatifnya bisa menggunakan perintah /report.\n\n"
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
            "Laporan hanya tersedia untuk Supervisor, Manager, HSA, IT, Team Leader, Team Leader Grapari, dan CS Grapari.\n\n"
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
        [InlineKeyboardButton("Technician", callback_data="group_Technician")],
        [InlineKeyboardButton("Grapari", callback_data="group_Grapari")]
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
    elif group == "Grapari":
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("Team Leader Grapari", callback_data="sub_TLG")],
            [InlineKeyboardButton("CS Grapari", callback_data="sub_CSG")]
        ])
        await query.edit_message_text("Pilih sub-peran Anda:", reply_markup=keyboard)
        return REG_SUBROLE
    else:  # Technician
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
        "TECH": "Technician", "HSA": "HSA",
        "TLG": "Team Leader Grapari",
        "CSG": "CS Grapari"
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
        last_update = get_last_order_date()
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
        f"👨🏼‍🔧 Technician Notes: {data['technician_notes']}\n"
        f"📦 Paket: {data['paket']}"
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
            f"👨🏼‍🔧 Technician Notes: {data['technician_notes']}\n"
            f"📦 Paket: {data['paket']}"
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

    _, subrole = get_user_role(user_id)
    keyboard = [
        [InlineKeyboardButton("📊 Laporan Detail (per WOK)", callback_data="sales_detail")],
    ]
    if can_view_summary(user_id):
        keyboard.append([InlineKeyboardButton("📈 Ringkasan (per WOK & Channel)", callback_data="sales_summary")])
    if subrole in ["Team Leader Grapari", "CS Grapari"]:
        keyboard.append([InlineKeyboardButton("📊 Grapari Performance (per STO)", callback_data="grapari_sto")])
    keyboard.append([InlineKeyboardButton("🔙 Kembali ke Menu Utama", callback_data="sales_back")])

    await query.message.reply_text("Pilih jenis laporan:", reply_markup=InlineKeyboardMarkup(keyboard))
    return SALES_WOK

# ---------- GRAPARI PERFORMANCE (per STO) ----------
async def grapari_sto_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    user_id = update.effective_user.id
    _, subrole = get_user_role(user_id)
    if subrole not in ["Team Leader Grapari", "CS Grapari"]:
        await query.message.reply_text("Anda tidak memiliki akses ke laporan ini.")
        return ConversationHandler.END

    # Only show 2025 and 2026 as requested
    year_buttons = [
        [InlineKeyboardButton("2025", callback_data="stoyear_2025")],
        [InlineKeyboardButton("2026", callback_data="stoyear_2026")]
    ]
    year_keyboard = InlineKeyboardMarkup(year_buttons)
    await query.message.reply_text("Pilih tahun:", reply_markup=year_keyboard)
    return GRAPARI_STO_YEAR

async def grapari_sto_year_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    year = int(query.data.split("_")[1])
    context.user_data["grapari_sto_year"] = year
    month_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("Januari", callback_data="stomonth_1"),
         InlineKeyboardButton("Februari", callback_data="stomonth_2"),
         InlineKeyboardButton("Maret", callback_data="stomonth_3")],
        [InlineKeyboardButton("April", callback_data="stomonth_4"),
         InlineKeyboardButton("Mei", callback_data="stomonth_5"),
         InlineKeyboardButton("Juni", callback_data="stomonth_6")],
        [InlineKeyboardButton("Juli", callback_data="stomonth_7"),
         InlineKeyboardButton("Agustus", callback_data="stomonth_8"),
         InlineKeyboardButton("September", callback_data="stomonth_9")],
        [InlineKeyboardButton("Oktober", callback_data="stomonth_10"),
         InlineKeyboardButton("November", callback_data="stomonth_11"),
         InlineKeyboardButton("Desember", callback_data="stomonth_12")]
    ])
    await query.edit_message_text("Pilih bulan:", reply_markup=month_keyboard)
    return GRAPARI_STO_MONTH

async def grapari_sto_month_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    month_num = int(query.data.split("_")[1])
    month_names = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                   "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    month_name = month_names[month_num - 1]
    year = context.user_data.get("grapari_sto_year", datetime.now().year)

    records = get_raw_records()
    prev_month = month_num - 1 if month_num > 1 else 12
    prev_year = year if month_num > 1 else year - 1

    def extract_date(date_str):
        if not date_str:
            return None, None
        date_str = clean_text(date_str)
        for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
            try:
                dt = datetime.strptime(date_str, fmt)
                return dt.year, dt.month
            except:
                continue
        return None, None

    sto_stats = {}
    total_input = 0
    total_completed = 0
    total_fallout = 0

    for rec in records:
        channel = rec.get("Channel Name", "").strip().upper()
        if channel != "GRAPARI":
            continue
        sto = rec.get("STO", "").strip()
        if not sto:
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
            if sto not in sto_stats:
                sto_stats[sto] = {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0}
            sto_stats[sto]["input"] += 1
            if status == "COMPLETED":
                sto_stats[sto]["completed"] += 1
            elif status == "FALLOUT":
                sto_stats[sto]["fallout"] += 1
        if y == prev_year and m == prev_month:
            if sto not in sto_stats:
                sto_stats[sto] = {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0}
            sto_stats[sto]["prev_input"] += 1

    sorted_stos = sorted(sto_stats.items(), key=lambda x: x[1]["input"], reverse=True)
    lines = [f"📊 *RINGKASAN PER STO (GRAPARI)*", f"📅 {month_name.upper()} {year}", ""]
    for sto, stats in sorted_stos:
        inp = stats["input"]
        comp = stats["completed"]
        flt = stats["fallout"]
        prev = stats["prev_input"]
        completion = (comp / inp * 100) if inp > 0 else 0
        contrib = (inp / total_input * 100) if total_input > 0 else 0
        mom = ((inp - prev) / prev * 100) if prev > 0 else (100 if inp > 0 else 0)
        lines.append(f"• *{sto}*: Input {inp} | Comp {comp} | FO {flt} | IO/PS {completion:.1f}% | Kontrib {contrib:.1f}% | MoM {mom:.1f}%")
    total_completion = (total_completed / total_input * 100) if total_input > 0 else 0
    total_prev = sum(stats.get("prev_input", 0) for stats in sto_stats.values())
    total_mom = ((total_input - total_prev) / total_prev * 100) if total_prev > 0 else (100 if total_input > 0 else 0)
    lines.append(f"\n• *TOTAL*: Input {total_input} | Comp {total_completed} | FO {total_fallout} | IO/PS {total_completion:.1f}% | Kontrib 100.0% | MoM {total_mom:.1f}%")
    lines.append(f"\n📅 Last Update Data: {get_last_order_date()}")

    await query.message.reply_text("\n".join(lines), parse_mode="Markdown")
    return ConversationHandler.END

# ---------- SALES CHOOSE (handles all main sales report options) ----------
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
        return DETAIL_WOK
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
        await show_main_menu(update, context)
        return ConversationHandler.END
    elif data == "grapari_sto":
        await query.message.delete()
        await grapari_sto_start(update, context)
        return ConversationHandler.END
    return ConversationHandler.END

# ---------- DETAILED WOK HANDLER ----------
async def detail_wok_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    wok = query.data.split("_", 1)[1]
    context.user_data["report_wok"] = wok

    user_id = update.effective_user.id
    _, subrole = get_user_role(user_id)
    context.user_data["report_subrole"] = subrole

    aggregate_only_roles = ["Manager", "Supervisor", "Inputters", "IT"]
    if subrole in aggregate_only_roles:
        current_year = datetime.now().year
        year_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton(str(current_year - 1), callback_data=f"year_{current_year-1}"),
             InlineKeyboardButton(str(current_year), callback_data=f"year_{current_year}")]
        ])
        await query.edit_message_text("Pilih tahun:", reply_markup=year_keyboard)
        return SALES_YEAR
    else:
        context.user_data["report_channel"] = "AGENCY"
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
    query = update.callback_query
    await query.answer()
    month_num = int(query.data.split("_")[1])
    month_names = ["Januari", "Februari", "Maret", "April", "Mei", "Juni",
                   "Juli", "Agustus", "September", "Oktober", "November", "Desember"]
    month_name = month_names[month_num - 1]
    year = context.user_data.get("report_year", datetime.now().year)
    wok = context.user_data["report_wok"]
    subrole = context.user_data.get("report_subrole")

    aggregate_only_roles = ["Manager", "Supervisor", "Inputters", "IT"]

    if subrole in aggregate_only_roles:
        records = get_raw_records()
        filtered = []
        for rec in records:
            rec_wok = rec.get("WOK", "")
            if rec_wok != wok:
                continue
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
                if m == month_num and y == year:
                    filtered.append(rec)
            except:
                continue

        if not filtered:
            await query.edit_message_text(f"Tidak ada data untuk WOK: {wok}, Tahun: {year}, Bulan: {month_name}.")
            return ConversationHandler.END

        await query.delete_message()
        processing_msg = await query.message.reply_text("⏳ Sedang memproses data...")

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

        lines = [f"📍 WOK: {wok}", f"📅 Tahun: {year}", f"📅 Bulan: {month_name.upper()}", ""]
        lines.append(f"📊 Total Order: {total}")
        lines.append("🔍 Rincian Status:")
        for s in ALL_STATUSES:
            count = status_counts[s]
            if count > 0:
                perc = (count / total) * 100 if total > 0 else 0
                if s == "COMPLETED":
                    lines.append(f"   ✅ {s}: {count} ({perc:.1f}%)")
                elif s in ["FALLOUT", "CANCELLED", "CANCELLED_SLA", "CANCEL_OSM_COMPLETED", "CANCEL_ORDER_INPROGRESS"]:
                    lines.append(f"   ❌ {s}: {count} ({perc:.1f}%)")
                else:
                    lines.append(f"   🔄 {s}: {count} ({perc:.1f}%)")
        if status_counts["OTHER"] > 0:
            perc_other = (status_counts["OTHER"] / total) * 100 if total > 0 else 0
            lines.append(f"   ❓ OTHER: {status_counts['OTHER']} ({perc_other:.1f}%)")
        lines.append(f"\n\n📅 Last Update Data: {get_last_order_date()}")
        await processing_msg.edit_text("\n".join(lines))
        return ConversationHandler.END

    if subrole == "Team Leader":
        context.user_data["tl_wok"] = wok
        context.user_data["tl_year"] = year
        context.user_data["tl_month_num"] = month_num
        context.user_data["tl_month_name"] = month_name
        context.user_data["tl_channel"] = "AGENCY"
        option_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📥 Download CSV", callback_data="tl_csv")],
            [InlineKeyboardButton("📋 Tampilkan Data", callback_data="tl_show")],
            [InlineKeyboardButton("📦 Popular Paket (Agency)", callback_data="tl_paket")]
        ])
        await query.edit_message_text(
            f"WOK: {wok}\n📅 {month_name.upper()} {year}\n\nPilih format laporan:",
            reply_markup=option_keyboard
        )
        return TEAM_LEADER_OPTION

    elif subrole in ["Team Leader Grapari", "CS Grapari"]:
        context.user_data["tl_wok"] = wok
        context.user_data["tl_year"] = year
        context.user_data["tl_month_num"] = month_num
        context.user_data["tl_month_name"] = month_name
        context.user_data["tl_channel"] = "GRAPARI"
        option_keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("📥 Download CSV", callback_data="tl_csv")],
            [InlineKeyboardButton("📋 Tampilkan Data", callback_data="tl_show")],
            [InlineKeyboardButton("📦 Popular Paket (Grapari)", callback_data="tl_paket")]
        ])
        await query.edit_message_text(
            f"📍 WOK: {wok} | Channel: GRAPARI\n📅 {month_name.upper()} {year}\n\nPilih format laporan:",
            reply_markup=option_keyboard
        )
        return TEAM_LEADER_OPTION

    return ConversationHandler.END

# ---------- TEAM LEADER OPTION CALLBACK (CSV, SHOW, PAKET) ----------
async def team_leader_option_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    action = query.data

    wok = context.user_data.get("tl_wok")
    year = context.user_data.get("tl_year")
    month_num = context.user_data.get("tl_month_num")
    month_name = context.user_data.get("tl_month_name")
    target_channel = context.user_data.get("tl_channel", "AGENCY")

    if not wok:
        await query.edit_message_text("Terjadi kesalahan. Silakan coba lagi.")
        return ConversationHandler.END

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
            if m == month_num and y == year and rec_wok == wok and rec_channel == target_channel:
                filtered.append(rec)
        except:
            continue

    if action == "tl_csv":
        if not filtered:
            await query.edit_message_text(f"Tidak ada data untuk WOK: {wok}, Bulan: {month_name} {year}.")
            return ConversationHandler.END
        output = io.StringIO()
        writer = csv.writer(output)
        writer.writerow(["Order ID", "STO", "WOK", "Status Order", "Channel Name", "Fallout Reason", "SalesForce", "Tanggal Complete", "Tanggal Input", "Sub Error Code", "Technician Notes", "Paket"])
        for rec in filtered:
            writer.writerow([
                rec.get("Order ID", ""),
                rec.get("STO", ""),
                rec.get("WOK", ""),
                rec.get("Status Order", ""),
                rec.get("Channel Name", ""),
                rec.get("Fallout Reason", ""),
                rec.get("SalesForce", ""),
                rec.get("Tanggal Complete", ""),
                rec.get("Tanggal Input", ""),
                rec.get("Sub Error Code", ""),
                rec.get("Technician Notes", ""),
                rec.get("Paket", "")
            ])
        output.seek(0)
        filename = f"orders_{wok}_{year}_{month_num}_{target_channel}.csv"
        await query.message.reply_document(document=output, filename=filename, caption=f"📊 Data Order {wok} - {month_name} {year} ({target_channel})")
        await query.delete_message()
        return ConversationHandler.END

    elif action == "tl_show":
        if not filtered:
            await query.edit_message_text(f"Tidak ada data untuk WOK: {wok}, Bulan: {month_name} {year}.")
            return ConversationHandler.END
        await query.delete_message()
        processing_msg = await query.message.reply_text("⏳ Sedang memproses data...")

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
            order_id = rec.get("Order ID", "")
            sales_dict[sf]["orders"].append((order_id, status))

        sorted_sf = sorted(sales_dict.items(), key=lambda x: x[1]["total"], reverse=True)
        header = f"📢 Channel: {target_channel}\n📅 Tahun: {year}\n📅 Bulan: {month_name.upper()}\n📍 WOK: {wok}"
        await processing_msg.edit_text(header)

        for sf, data in sorted_sf:
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
        await query.message.reply_text("✅ Selesai. Terima kasih.")
        return ConversationHandler.END

    elif action == "tl_paket":
        await query.delete_message()
        processing_msg = await query.message.reply_text(f"⏳ Menghitung semua paket COMPLETED ({target_channel} only)...")

        paket_records = []
        for rec in records:
            rec_wok = rec.get("WOK", "")
            rec_channel = rec.get("Channel Name", "").strip().upper()
            rec_tanggal = rec.get("Tanggal Input", "")
            if rec_wok != wok or rec_channel != target_channel:
                continue
            if not rec_tanggal:
                continue
            try:
                dt = None
                for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
                    try:
                        dt = datetime.strptime(rec_tanggal.strip(), fmt)
                        break
                    except:
                        continue
                if dt and dt.year == year and dt.month == month_num:
                    paket_records.append(rec)
            except:
                continue

        if not paket_records:
            await processing_msg.edit_text(f"Tidak ada data {target_channel} untuk WOK {wok} pada {month_name} {year}.")
            return ConversationHandler.END

        total_completed = 0
        paket_counts = {}
        for rec in paket_records:
            status = (rec.get("Status Order", "") or "").strip().upper()
            if status != "COMPLETED":
                continue
            total_completed += 1
            paket = rec.get("Paket", "").strip()
            if not paket:
                continue
            paket_counts[paket] = paket_counts.get(paket, 0) + 1

        if total_completed == 0:
            await processing_msg.edit_text(f"Tidak ada order COMPLETED untuk {target_channel} di {wok} pada {month_name} {year}.")
            return ConversationHandler.END

        sorted_paket = sorted(paket_counts.items(), key=lambda x: x[1], reverse=True)
        lines = [f"📦 *SEMUA PAKET COMPLETED – {target_channel} ONLY*",
                 f"📍 WOK: {wok}",
                 f"📅 {month_name.upper()} {year}",
                 ""]
        for pkg, cnt in sorted_paket:
            pct = (cnt / total_completed) * 100
            lines.append(f"• {pkg}: {cnt} ({pct:.1f}%)")
        lines.append(f"\n📅 Last Update Data: {get_last_order_date()}")
        await processing_msg.edit_text("\n".join(lines), parse_mode="Markdown")
        return ConversationHandler.END

    else:
        await query.edit_message_text("Perintah tidak dikenal.")
        return ConversationHandler.END

# ---------- SUMMARY REPORT (RINGKASAN) ----------
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

    # ----- 1. per WOK summary -----
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

    lines1 = [f"📊 *RINGKASAN PER WOK*", f"📅 {month_name.upper()} {year}", ""]
    for wok in wok_list:
        stats = wok_stats.get(wok, {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0})
        inp = stats["input"]
        comp = stats["completed"]
        flt = stats["fallout"]
        prev = stats["prev_input"]
        completion = (comp / inp * 100) if inp > 0 else 0
        contrib = (inp / total_input * 100) if total_input > 0 else 0
        mom = ((inp - prev) / prev * 100) if prev > 0 else (100 if inp > 0 else 0)
        lines1.append(f"• *{wok}*: Input {inp} | Comp {comp} | FO {flt} | IO/PS {completion:.1f}% | Kontrib {contrib:.1f}% | MoM {mom:.1f}%")
    total_completion = (total_completed / total_input * 100) if total_input > 0 else 0
    total_prev = sum(stats.get("prev_input", 0) for stats in wok_stats.values())
    total_mom = ((total_input - total_prev) / total_prev * 100) if total_prev > 0 else (100 if total_input > 0 else 0)
    lines1.append(f"\n• *TOTAL*: Input {total_input} | Comp {total_completed} | FO {total_fallout} | IO/PS {total_completion:.1f}% | Kontrib 100.0% | MoM {total_mom:.1f}%")
    lines1.append(f"\n\n📅 Last Update Data: {get_last_order_date()}")
    await query.message.reply_text("\n".join(lines1), parse_mode="Markdown")

    # ----- 2. overall per channel summary -----
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

    lines2 = [f"📊 *RINGKASAN PER CHANNEL*", f"📅 {month_name.upper()} {year}", ""]
    for ch in channels:
        stats = chan_stats.get(ch, {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0})
        inp = stats["input"]
        comp = stats["completed"]
        flt = stats["fallout"]
        prev = stats["prev_input"]
        completion = (comp / inp * 100) if inp > 0 else 0
        contrib = (inp / total_chan_input * 100) if total_chan_input > 0 else 0
        mom = ((inp - prev) / prev * 100) if prev > 0 else (100 if inp > 0 else 0)
        lines2.append(f"• *{ch}*: Input {inp} | Comp {comp} | FO {flt} | IO/PS {completion:.1f}% | Kontrib {contrib:.1f}% | MoM {mom:.1f}%")
    total_chan_completion = (total_chan_completed / total_chan_input * 100) if total_chan_input > 0 else 0
    total_chan_prev = sum(stats.get("prev_input", 0) for stats in chan_stats.values())
    total_chan_mom = ((total_chan_input - total_chan_prev) / total_chan_prev * 100) if total_chan_prev > 0 else (100 if total_chan_input > 0 else 0)
    lines2.append(f"\n• *TOTAL*: Input {total_chan_input} | Comp {total_chan_completed} | FO {total_chan_fallout} | IO/PS {total_chan_completion:.1f}% | Kontrib 100.0% | MoM {total_chan_mom:.1f}%")
    lines2.append(f"\n\n📅 Last Update Data: {get_last_order_date()}")
    await query.message.reply_text("\n".join(lines2), parse_mode="Markdown")

    # ----- 3. per WOK channel summaries -----
    for wok in wok_list:
        wok_chan_stats = {ch: {"input": 0, "completed": 0, "fallout": 0, "prev_input": 0} for ch in channels}
        total_wok_input = 0
        total_wok_completed = 0
        total_wok_fallout = 0
        for rec in records:
            rec_wok = rec.get("WOK", "")
            if rec_wok != wok:
                continue
            rec_channel = rec.get("Channel Name", "")
            if rec_channel not in channels:
                continue
            tgl = rec.get("Tanggal Input", "")
            y, m = extract_date(tgl)
            if y is None or m is None:
                continue
            status = (rec.get("Status Order", "") or "").strip().upper()
            if y == year and m == month_num:
                total_wok_input += 1
                if status == "COMPLETED":
                    total_wok_completed += 1
                elif status == "FALLOUT":
                    total_wok_fallout += 1
                wok_chan_stats[rec_channel]["input"] += 1
                if status == "COMPLETED":
                    wok_chan_stats[rec_channel]["completed"] += 1
                elif status == "FALLOUT":
                    wok_chan_stats[rec_channel]["fallout"] += 1
            if y == prev_year and m == prev_month:
                wok_chan_stats[rec_channel]["prev_input"] += 1

        lines_wok = [f"📊 *RINGKASAN PER CHANNEL UNTUK {wok}*", f"📅 {month_name.upper()} {year}", ""]
        for ch in channels:
            stats = wok_chan_stats[ch]
            inp = stats["input"]
            comp = stats["completed"]
            flt = stats["fallout"]
            prev = stats["prev_input"]
            completion = (comp / inp * 100) if inp > 0 else 0
            contrib = (inp / total_wok_input * 100) if total_wok_input > 0 else 0
            mom = ((inp - prev) / prev * 100) if prev > 0 else (100 if inp > 0 else 0)
            lines_wok.append(f"• *{ch}*: Input {inp} | Comp {comp} | FO {flt} | IO/PS {completion:.1f}% | Kontrib {contrib:.1f}% | MoM {mom:.1f}%")
        total_wok_completion = (total_wok_completed / total_wok_input * 100) if total_wok_input > 0 else 0
        total_wok_prev = sum(wok_chan_stats[ch]["prev_input"] for ch in channels)
        total_wok_mom = ((total_wok_input - total_wok_prev) / total_wok_prev * 100) if total_wok_prev > 0 else (100 if total_wok_input > 0 else 0)
        lines_wok.append(f"\n• *TOTAL*: Input {total_wok_input} | Comp {total_wok_completed} | FO {total_wok_fallout} | IO/PS {total_wok_completion:.1f}% | Kontrib 100.0% | MoM {total_wok_mom:.1f}%")
        lines_wok.append(f"\n\n📅 Last Update Data: {get_last_order_date()}")
        await query.message.reply_text("\n".join(lines_wok), parse_mode="Markdown")

    # --- Popular Paket Summary – ALL COMPLETED packages (no limit) ---
    await query.message.reply_text("📦 Menghitung semua paket dengan status COMPLETED...")

    def clean_field(s):
        if not s:
            return ""
        s = re.sub(r'[\u200b\u00a0\u200c\u200d]', '', str(s))
        return s.strip()

    def robust_extract_date(date_str):
        if not date_str:
            return None, None
        date_str = clean_field(date_str)
        for fmt in ["%d/%m/%Y %H:%M", "%d/%m/%Y", "%Y-%m-%d %H:%M:%S", "%Y-%m-%d"]:
            try:
                dt = datetime.strptime(date_str, fmt)
                return dt.year, dt.month
            except ValueError:
                continue
        try:
            parts = date_str.split()
            date_part = parts[0]
            if '-' in date_part:
                y, m, d = map(int, date_part.split('-'))
                return y, m
            else:
                d, m, y = map(int, date_part.split('/'))
                return y, m
        except:
            return None, None

    def get_all_paket_with_pct(records_filter_func, total_completed_orders):
        paket_counts = {}
        for rec in records_filter_func:
            status = clean_field(rec.get("Status Order", "")).upper()
            if status != "COMPLETED":
                continue
            paket = clean_field(rec.get("Paket", ""))
            if not paket:
                continue
            paket_counts[paket] = paket_counts.get(paket, 0) + 1
        sorted_paket = sorted(paket_counts.items(), key=lambda x: x[1], reverse=True)
        result = []
        for paket, count in sorted_paket:
            pct = (count / total_completed_orders * 100) if total_completed_orders > 0 else 0
            result.append((paket, count, pct))
        return result

    # Global (all WOKs, all channels)
    global_completed_records = []
    for rec in records:
        tgl = rec.get("Tanggal Input", "")
        y, m = robust_extract_date(tgl)
        if y == year and m == month_num:
            status = clean_field(rec.get("Status Order", "")).upper()
            if status == "COMPLETED":
                global_completed_records.append(rec)

    global_total = len(global_completed_records)
    logger.info(f"Global completed orders for {month_name} {year}: {global_total}")

    global_pakets = get_all_paket_with_pct(global_completed_records, global_total)
    if global_pakets:
        lines = [f"📊 *SEMUA PAKET COMPLETED (GLOBAL)*", f"📅 {month_name.upper()} {year}", ""]
        for paket, count, pct in global_pakets:
            lines.append(f"• {paket}: {count} ({pct:.1f}%)")
        lines.append(f"\n📅 Last Update Data: {get_last_order_date()}")
        await query.message.reply_text("\n".join(lines), parse_mode="Markdown")
    else:
        await query.message.reply_text(f"Tidak ada order COMPLETED untuk periode {month_name} {year}.")

    # Per WOK
    for wok in wok_list:
        wok_completed = []
        for rec in records:
            if rec.get("WOK", "") != wok:
                continue
            tgl = rec.get("Tanggal Input", "")
            y, m = robust_extract_date(tgl)
            if y == year and m == month_num:
                status = clean_field(rec.get("Status Order", "")).upper()
                if status == "COMPLETED":
                    wok_completed.append(rec)
        wok_total = len(wok_completed)
        wok_pakets = get_all_paket_with_pct(wok_completed, wok_total)
        if wok_pakets:
            lines = [f"📊 *SEMUA PAKET COMPLETED – {wok}*", f"📅 {month_name.upper()} {year}", ""]
            for paket, count, pct in wok_pakets:
                lines.append(f"• {paket}: {count} ({pct:.1f}%)")
            lines.append(f"\n📅 Last Update Data: {get_last_order_date()}")
            await query.message.reply_text("\n".join(lines), parse_mode="Markdown")
        else:
            await query.message.reply_text(f"Tidak ada order COMPLETED di {wok} pada {month_name} {year}.")

    return ConversationHandler.END

# ---------- USAGE REPORT SLASH COMMAND FALLBACK ----------
async def report(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not can_view_reports(user_id):
        await update.message.reply_text("Anda tidak memiliki izin untuk melihat laporan.")
        return
    args = context.args
    if not args:
        await update.message.reply_text("Penggunaan: /report [day|week|month]")
        return
    await generate_report(update, update.message, user_id, args)

# ---------- GENERAL COMMANDS ----------
async def cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Perintah dibatalkan. Kembali ke menu utama.")
    await show_main_menu(update, context)
    return ConversationHandler.END

async def ping(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("pong")

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_guide(update, update.effective_user.id, context)

async def guide_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_guide(update, update.effective_user.id, context)

async def start_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await show_main_menu(update, context)

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
            SALES_WOK: [CallbackQueryHandler(sales_choose, pattern="^(sales_detail|sales_summary|sales_back|grapari_sto)$")],
            DETAIL_WOK: [CallbackQueryHandler(detail_wok_selected, pattern="^wok_")],
            SALES_YEAR: [CallbackQueryHandler(sales_year_selected, pattern="^year_")],
            SALES_MONTH: [CallbackQueryHandler(sales_month_selected, pattern="^month_")],
            TEAM_LEADER_OPTION: [CallbackQueryHandler(team_leader_option_callback, pattern="^(tl_csv|tl_show|tl_paket)$")],
            SUMMARY_YEAR: [CallbackQueryHandler(summary_year_selected, pattern="^sumyear_")],
            SUMMARY_MONTH: [CallbackQueryHandler(summary_month_selected, pattern="^summonth_")],
            GRAPARI_STO_YEAR: [CallbackQueryHandler(grapari_sto_year_selected, pattern="^stoyear_")],
            GRAPARI_STO_MONTH: [CallbackQueryHandler(grapari_sto_month_selected, pattern="^stomonth_")],
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

    # Usage report conversation
    report_conv = ConversationHandler(
        entry_points=[CallbackQueryHandler(usage_report_menu, pattern="^menu_usage_report$")],
        states={REPORT_OPTION: [CallbackQueryHandler(report_option_callback, pattern="^report_")]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(report_conv)

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
