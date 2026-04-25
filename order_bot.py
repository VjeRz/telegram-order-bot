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

SALES_WOK = 10
SALES_CHANNEL = 11
SALES_MONTH = 12

ALL_STATUSES = [
    "PENDING_CUSTOMER_VERIFICATION",
    "PROVISION_START",
    "TECH_ASSIGNED",
    "PENDING_APPOINTMENT_CREATION",
    "PENDING_CONTRACT_APPROVAL",
    "PROVISION_ISSUED",
    "COMPLETED",
    "OSS_TESTING_SERVICE",
    "RE",
    "FALLOUT",
    "ODP_AVAILABLE",
    "CANCELLED",
    "PENDING_PAYMENT_FOLLOWUP",
    "PAYMENT_INPROGRESS",
    "CANCEL_OSM_COMPLETED",
    "TSEL_ACTIVATION_FALLOUT",
    "CANCEL_ORDER_INPROGRESS",
    "TECH_ARRIVED",
    "CANCELLED_SLA",
    "PENDING_DUNNING_PAYMENT_FOLLOWUP",
    "PENDING_PAYMENT",
    "TECH_PICKED_UP",
    "TECH_ON_THE_WAY",
    "CONTRACT_APPROVED",
    "WIRELESS_FULFILMENT_INPROGRESS",
    "PROVISION_DESIGN"
]

logging.basicConfig(
    format="%(asctime)s - %(name)s - %(levelname)s - %(message)s", level=logging.INFO
)
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

def notify_approver(bot, user_id, name, role_group, subrole, wok="", sfid=""):
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
        bot.send_message(chat_id=IT_TELEGRAM_ID, text=text, reply_markup=keyboard)

# ---------- WELCOME/GUIDE MESSAGES ----------
async def send_welcome_message(update_or_context, user_id, is_new_approval=False):
    role_group, subrole = get_user_role(user_id)
    if not role_group:
        text = (
            "📖 Selamat Datang di Bot Cek Order SF Branch Manado\n\n"
            "Anda belum terdaftar. Silakan gunakan perintah /register untuk memulai pendaftaran.\n\n"
            "Setelah mendaftar, Anda harus menunggu persetujuan dari IT. Anda akan diberi tahu setelah disetujui.\n\n"
            "Jika Anda sudah terdaftar dan disetujui, gunakan /start untuk memeriksa Order ID.\n\n"
            "Untuk bantuan lebih lanjut, ketik /help."
        )
        if isinstance(update_or_context, Update):
            await update_or_context.message.reply_text(text)
        else:
            await update_or_context.send_message(chat_id=user_id, text=text)
        return

    if subrole in ["Manager", "Supervisor", "HSA", "IT"]:
        text = (
            f"📋 Panduan Penggunaan Bot\n\n"
            f"✅ Anda terdaftar sebagai {role_group} - {subrole}.\n\n"
            "🔍 Cek Order:\n"
            "Gunakan /start lalu masukkan Order ID.\n\n"
            "📦 Cek Bulk Order:\n"
            "Gunakan /bulk lalu masukkan beberapa Order ID dipisah spasi atau baris baru.\n\n"
            "📊 Laporan Penggunaan:\n"
            "• /report day → laporan hari ini (teks)\n"
            "• /report week → laporan minggu ini (teks)\n"
            "• /report month → laporan bulan ini (file CSV)\n"
            "• /report 2026-04-01 2026-04-10 → laporan rentang tanggal (file CSV)\n\n"
            "📈 Laporan Performa Sales:\n"
            "• /salesreport → ikuti menu interaktif untuk memilih WOK, Channel, dan Bulan\n\n"
            "📎 File CSV dapat diunduh dan dibuka di Excel atau Google Sheets.\n\n"
            "Untuk daftar perintah lengkap, ketik /help."
        )
    else:
        text = (
            f"📋 Panduan Penggunaan Bot\n\n"
            f"✅ Anda terdaftar sebagai {role_group} - {subrole}.\n\n"
            "🔍 Cek Order:\n"
            "Gunakan /start lalu masukkan Order ID.\n\n"
            "📦 Cek Bulk Order:\n"
            "Gunakan /bulk lalu masukkan beberapa Order ID dipisah spasi atau baris baru.\n\n"
            "📊 Laporan:\n"
            "Laporan hanya tersedia untuk Supervisor, Manager, HSA, dan IT.\n\n"
            "Untuk daftar perintah lengkap, ketik /help."
        )
    if isinstance(update_or_context, Update):
        await update_or_context.message.reply_text(text)
    else:
        await update_or_context.send_message(chat_id=user_id, text=text)

# ---------- REGISTRATION FLOW ----------
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
    else:
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
        await query.edit_message_text("Registration completed. Saving your data...")
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
        await update.callback_query.edit_message_text("Saving registration data...")
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
    reply_text = "Registration submitted. You will be notified once approved by IT."

    notify_approver(update.get_bot(), user_id, name, role_group, subrole, wok, sfid)

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
        await update.message.reply_text("Only IT users can view pending registrations.")
        return

    records = users_sheet.get_all_records()
    pending_users = []
    for idx, row in enumerate(records, start=2):
        if row.get("ApprovalStatus") == "pending":
            pending_users.append((idx, row))
    if not pending_users:
        await update.message.reply_text("No pending registrations.")
        return

    for idx, row in pending_users:
        name = row.get("Name")
        role_group = row.get("RoleGroup")
        subrole = row.get("SubRole")
        wok = row.get("WOK", "N/A") if role_group == "Agency" else "N/A"
        telegram_id = row.get("TelegramID")
        text = (
            f"Pending registration:\n"
            f"Name: {name}\n"
            f"Role: {role_group} - {subrole}\n"
            f"WOK: {wok}\n"
            f"User ID: {telegram_id}"
        )
        keyboard = InlineKeyboardMarkup([
            [InlineKeyboardButton("✅ Approve", callback_data=f"approve_{telegram_id}_{idx}"),
             InlineKeyboardButton("❌ Reject", callback_data=f"reject_{telegram_id}_{idx}")]
        ])
        await update.message.reply_text(text, reply_markup=keyboard)

async def approval_callback(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    data = query.data
    parts = data.split("_")
    action = parts[0]
    target_id = int(parts[1])
    row_index = int(parts[2])

    if action == "approve":
        users_sheet.update_cell(row_index, 6, "approved")
        users_sheet.update_cell(row_index, 8, str(update.effective_user.id))
        users_sheet.update_cell(row_index, 9, datetime.now().strftime("%Y-%m-%d %H:%M:%S"))
        await context.bot.send_message(chat_id=target_id, text="Your registration has been approved! You can now use /start to check orders.")
        await send_welcome_message(context.bot, target_id, is_new_approval=True)
        await query.edit_message_text(f"✅ User {target_id} approved.")
    elif action == "reject":
        users_sheet.update_cell(row_index, 6, "rejected")
        await context.bot.send_message(chat_id=target_id, text="Your registration was rejected. You can try /register again.")
        await query.edit_message_text(f"❌ User {target_id} rejected.")

# ---------- BULK ORDER CHECK ----------
async def bulk_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_user_approved(user_id):
        await update.message.reply_text("Anda belum terdaftar atau belum disetujui. Silakan gunakan /register terlebih dahulu.")
        return
    await update.message.reply_text(
        "📦 *Cek Bulk Order*\n\n"
        "Kirimkan beberapa Order ID dalam satu pesan.\n"
        "Pisahkan dengan spasi atau baris baru.\n\n"
        "Contoh:\n"
        "`AOi426042509434427179f980 AOi4260425091936300715b70`\n\n"
        "atau:\n"
        "`AOi426042509434427179f980`\n`AOi4260425091936300715b70`\n\n"
        "Bot akan membalas detail masing-masing Order ID secara terpisah.",
        parse_mode="Markdown"
    )
    return

async def bulk_process(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    raw_text = update.message.text
    # Split by whitespace (spaces, newlines, tabs)
    order_ids = re.split(r'\s+', raw_text.strip())
    order_ids = [oid for oid in order_ids if oid]  # remove empty strings

    if not order_ids:
        await update.message.reply_text("Tidak ada Order ID yang ditemukan. Kirim ulang dengan format yang benar.")
        return

    await update.message.reply_text(f"⏳ Memproses {len(order_ids)} Order ID...")

    found_count = 0
    not_found = []
    for oid in order_ids:
        clean_oid = clean_text(oid)
        data = find_order_details(clean_oid)
        if data is None:
            not_found.append(clean_oid)
            continue
        found_count += 1
        log_usage(user_id, clean_oid)
        reply = (
            f"✅ Order ID: {clean_oid}\n"
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
        not_found_msg = "❌ Order ID tidak ditemukan:\n" + "\n".join(not_found)
        await update.message.reply_text(not_found_msg)

    await update.message.reply_text(f"✅ Selesai. {found_count} dari {len(order_ids)} Order ID berhasil ditemukan.")

# ---------- SALES REPORT INTERACTIVE MENU ----------
async def sales_report_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not (is_user_approved(user_id) and can_view_sales_report(user_id)):
        await update.message.reply_text("Hanya Supervisor, Team Leader, IT, dan Manager yang dapat melihat laporan performa sales.")
        return ConversationHandler.END

    wok_keyboard = InlineKeyboardMarkup([
        [InlineKeyboardButton("MANADO TALAUD", callback_data="wok_MANADO TALAUD")],
        [InlineKeyboardButton("BOLAANG MONGONDOW", callback_data="wok_BOLAANG MONGONDOW")],
        [InlineKeyboardButton("GORONTALO - PAHUWATO", callback_data="wok_GORONTALO - PAHUWATO")],
        [InlineKeyboardButton("BITUNG MINAHASA", callback_data="wok_BITUNG MINAHASA")]
    ])
    await update.message.reply_text("Pilih WOK:", reply_markup=wok_keyboard)
    return SALES_WOK

async def sales_wok_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    wok = query.data.split("_", 1)[1]
    context.user_data["report_wok"] = wok

    user_id = update.effective_user.id
    _, subrole = get_user_role(user_id)
    context.user_data["report_subrole"] = subrole

    if subrole == "Team Leader":
        context.user_data["report_channel"] = "AGENCY"
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

async def sales_channel_selected(update: Update, context: ContextTypes.DEFAULT_TYPE):
    query = update.callback_query
    await query.answer()
    channel = query.data.split("_", 1)[1]
    context.user_data["report_channel"] = channel
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
    context.user_data["report_month"] = month_name
    context.user_data["report_month_num"] = month_num

    records = get_order_sheet_records()
    wok = context.user_data["report_wok"]
    channel = context.user_data["report_channel"]
    subrole = context.user_data["report_subrole"]

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
                year, month, day = map(int, date_part.split('-'))
            else:
                date_part = rec_tanggal_input.split()[0]
                day, month, year = map(int, date_part.split('/'))
            if month == month_num and rec_wok == wok and rec_channel == channel:
                filtered.append(rec)
        except:
            continue

    if not filtered:
        await query.edit_message_text(f"Tidak ada data untuk WOK: {wok}, Channel: {channel}, Bulan: {month_name}.")
        return ConversationHandler.END

    await query.delete_message()
    processing_msg = await query.message.reply_text("⏳ Sedang memproses data...")

    if channel == "AGENCY":
        sales_dict = {}
        for rec in filtered:
            sf = rec.get("SalesForce", "").strip()
            if not sf:
                continue
            status = rec.get("Status Order", "").upper()
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

        header = f"📢 Channel: {channel}\n📅 Bulan: {month_name.upper()}\n📍 WOK: {wok}"
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
                order_lines = []
                for oid, stat in data["orders"]:
                    order_lines.append(f"      • {oid} ({stat})")
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
                lines = [f"👤 {sf}"]
                lines.append(f"   🔢 Total Order: {data['total']}")
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
                msg_text = "\n".join(lines)
                await query.message.reply_text(msg_text)
        await query.message.reply_text("✅ Selesai. Terima kasih.")
    else:
        status_counts = {s: 0 for s in ALL_STATUSES}
        status_counts["OTHER"] = 0
        total = 0
        for rec in filtered:
            total += 1
            status = rec.get("Status Order", "").upper()
            if status in status_counts:
                status_counts[status] += 1
            else:
                status_counts["OTHER"] += 1
        lines = [f"📢 Channel: {channel}", f"📅 Bulan: {month_name.upper()}", f"📍 WOK: {wok}", ""]
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

async def sales_cancel(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text("Laporan dibatalkan.")
    return ConversationHandler.END

# ---------- USAGE REPORT ----------
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

# ---------- ORDER LOOKUP (SINGLE) ----------
async def start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_id = update.effective_user.id
    if not is_user_approved(user_id):
        guide = (
            "📖 Selamat Datang di Bot Cek Order SF Branch Manado\n\n"
            "Anda belum terdaftar. Silakan gunakan perintah /register untuk memulai pendaftaran.\n\n"
            "Setelah mendaftar, Anda harus menunggu persetujuan dari IT. Anda akan diberi tahu setelah disetujui.\n\n"
            "Untuk bantuan lebih lanjut, ketik /help."
        )
        await update.message.reply_text(guide)
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

async def help_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    help_text = (
        "📖 Daftar Perintah Bot\n\n"
        "/register - Memulai proses registrasi pengguna baru (semua role)\n"
        "/start - Memeriksa Order ID (satu per satu)\n"
        "/bulk - Memeriksa beberapa Order ID sekaligus (dipisah spasi atau baris baru)\n"
        "/guide - Menampilkan panduan penggunaan sesuai role Anda\n"
        "/pending - Melihat dan menyetujui/menolak registrasi (khusus IT)\n"
        "/report [day|week|month|YYYY-MM-DD YYYY-MM-DD] - Laporan penggunaan bot (untuk Manager, SPV, HSA, IT)\n"
        "/salesreport - Laporan performa sales interaktif (untuk Supervisor, Team Leader, IT, Manager)\n"
        "/ping - Tes koneksi bot\n"
        "/help - Menampilkan pesan bantuan ini\n\n"
        "📌 Catatan:\n"
        "- Semua registrasi memerlukan persetujuan IT.\n"
        "- Hanya Agency yang diminta WOK dan SF ID.\n"
        "- Laporan harian/mingguan ditampilkan sebagai teks, laporan bulanan/custom sebagai file CSV."
    )
    await update.message.reply_text(help_text)

async def guide_command(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await send_welcome_message(update, update.effective_user.id)

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

    sales_conv = ConversationHandler(
        entry_points=[CommandHandler("salesreport", sales_report_start)],
        states={
            SALES_WOK: [CallbackQueryHandler(sales_wok_selected, pattern="^wok_")],
            SALES_CHANNEL: [CallbackQueryHandler(sales_channel_selected, pattern="^chan_")],
            SALES_MONTH: [CallbackQueryHandler(sales_month_selected, pattern="^month_")],
        },
        fallbacks=[CommandHandler("cancel", sales_cancel)],
        allow_reentry=True,
    )
    app.add_handler(sales_conv)

    app.add_handler(CommandHandler("ping", ping))
    app.add_handler(CommandHandler("help", help_command))
    app.add_handler(CommandHandler("guide", guide_command))

    # Single order lookup conversation
    order_conv = ConversationHandler(
        entry_points=[CommandHandler("start", start)],
        states={WAITING_FOR_ORDER_ID: [MessageHandler(filters.TEXT & ~filters.COMMAND, receive_order_id)]},
        fallbacks=[CommandHandler("cancel", cancel)],
        allow_reentry=True,
    )
    app.add_handler(order_conv)

    # Bulk order command (no conversation, just process and reply)
    app.add_handler(CommandHandler("bulk", bulk_start))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, bulk_process))

    logger.info("Bot is polling...")
    app.run_polling()

if __name__ == "__main__":
    main()
