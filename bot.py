import asyncio
import logging
import os
import json
import aiosqlite
from datetime import datetime
from aiogram import Bot, Dispatcher, F
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, CallbackQuery, FSInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from aiogram.fsm.storage.memory import MemoryStorage
from aiogram.types import InlineKeyboardMarkup, InlineKeyboardButton
import httpx
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

logging.basicConfig(level=logging.INFO)

BOT_TOKEN = os.getenv("BOT_TOKEN")
OPENAI_API_KEY = os.getenv("OPENAI_API_KEY", "")
ADMIN_IDS = list(map(int, os.getenv("ADMIN_IDS", "0").split(",")))
DB = "bot.db"
DOCS_DIR = "docs"
os.makedirs(DOCS_DIR, exist_ok=True)
PRICE = 10000

# ─── DIZAYN FUNKSIYALARI ─────────────────────────────────────────────────────

def set_border(table):
    tbl = table._tbl
    tblPr = tbl.tblPr
    tblBorders = OxmlElement('w:tblBorders')
    for border_name in ['top', 'left', 'bottom', 'right', 'insideH', 'insideV']:
        border = OxmlElement(f'w:{border_name}')
        border.set(qn('w:val'), 'single')
        border.set(qn('w:sz'), '4')
        border.set(qn('w:space'), '0')
        border.set(qn('w:color'), '1B3A6B')
        tblBorders.append(border)
    tblPr.append(tblBorders)

def add_header(doc, title):
    header_table = doc.add_table(rows=1, cols=1)
    header_table.alignment = WD_TABLE_ALIGNMENT.CENTER
    cell = header_table.cell(0, 0)
    tcPr = cell._tc.get_or_add_tcPr()
    shd = OxmlElement('w:shd')
    shd.set(qn('w:val'), 'clear')
    shd.set(qn('w:color'), 'auto')
    shd.set(qn('w:fill'), '1B3A6B')
    tcPr.append(shd)
    p = cell.paragraphs[0]
    p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    run = p.add_run("⚖️  YURIDIK YORDAMCHI  ⚖️")
    run.font.color.rgb = RGBColor(255, 255, 255)
    run.font.size = Pt(13)
    run.font.bold = True
    run.font.name = 'Times New Roman'
    doc.add_paragraph()
    title_p = doc.add_paragraph()
    title_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    title_run = title_p.add_run(title)
    title_run.font.size = Pt(16)
    title_run.font.bold = True
    title_run.font.name = 'Times New Roman'
    title_run.font.color.rgb = RGBColor(27, 58, 107)
    sep = doc.add_paragraph()
    sep.add_run("─" * 65).font.size = Pt(10)
    doc.add_paragraph()

def add_info_table(doc, rows_data):
    table = doc.add_table(rows=len(rows_data), cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    set_border(table)
    for i, (label, value) in enumerate(rows_data):
        lc = table.cell(i, 0)
        tcPr = lc._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), 'E8EEF7')
        tcPr.append(shd)
        lp = lc.paragraphs[0]
        lr = lp.add_run(f"  {label}")
        lr.font.bold = True
        lr.font.size = Pt(11)
        lr.font.name = 'Times New Roman'
        lr.font.color.rgb = RGBColor(27, 58, 107)
        vc = table.cell(i, 1)
        vp = vc.paragraphs[0]
        vr = vp.add_run(f"  {value}")
        vr.font.size = Pt(11)
        vr.font.name = 'Times New Roman'

def add_signature_table(doc, left_title, right_title, left_name, right_name):
    doc.add_paragraph()
    sig_p = doc.add_paragraph()
    sig_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    sig_r = sig_p.add_run("TOMONLAR IMZOLARI")
    sig_r.font.bold = True
    sig_r.font.size = Pt(12)
    sig_r.font.color.rgb = RGBColor(27, 58, 107)
    sig_r.font.name = 'Times New Roman'
    doc.add_paragraph()
    table = doc.add_table(rows=4, cols=2)
    table.alignment = WD_TABLE_ALIGNMENT.CENTER
    sana = datetime.now().strftime('%d.%m.%Y')
    for col, title in enumerate([left_title, right_title]):
        cell = table.cell(0, col)
        tcPr = cell._tc.get_or_add_tcPr()
        shd = OxmlElement('w:shd')
        shd.set(qn('w:val'), 'clear')
        shd.set(qn('w:color'), 'auto')
        shd.set(qn('w:fill'), '1B3A6B')
        tcPr.append(shd)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(title)
        r.font.bold = True
        r.font.color.rgb = RGBColor(255, 255, 255)
        r.font.size = Pt(11)
        r.font.name = 'Times New Roman'
    for col, name in enumerate([left_name, right_name]):
        cell = table.cell(1, col)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        r = p.add_run(name)
        r.font.size = Pt(11)
        r.font.name = 'Times New Roman'
    for col in range(2):
        cell = table.cell(2, col)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run("Imzo: ___________________").font.size = Pt(11)
    for col in range(2):
        cell = table.cell(3, col)
        p = cell.paragraphs[0]
        p.alignment = WD_ALIGN_PARAGRAPH.CENTER
        p.add_run(f"Sana: {sana}").font.size = Pt(11)
    set_border(table)

def add_footer(doc):
    doc.add_paragraph()
    footer_p = doc.add_paragraph()
    footer_p.alignment = WD_ALIGN_PARAGRAPH.CENTER
    footer_r = footer_p.add_run("© Yuridik Yordamchi Bot")
    footer_r.font.size = Pt(8)
    footer_r.font.italic = True
    footer_r.font.color.rgb = RGBColor(150, 150, 150)

def make_doc(doc_type, data, order_id):
    doc = Document()
    section = doc.sections[0]
    section.left_margin = Cm(2.5)
    section.right_margin = Cm(2)
    section.top_margin = Cm(2)
    section.bottom_margin = Cm(2)
    sana = datetime.now().strftime('%d.%m.%Y')

    if doc_type == "ijara":
        add_header(doc, "IJARA SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Ijara beruvchi:", data.get('beruvchi_fio','___')),
            ("Pasport:", data.get('beruvchi_pasport','___')),
            ("Ijara oluvchi:", data.get('oluvchi_fio','___')),
            ("Pasport:", data.get('oluvchi_pasport','___')),
            ("Manzil:", data.get('manzil','___')),
            ("Oylik ijara haqi:", f"{data.get('narx','___')} so'm"),
            ("Muddat:", f"{data.get('muddat','___')} oy"),
        ])
        add_signature_table(doc, "IJARA BERUVCHI", "IJARA OLUVCHI",
            data.get('beruvchi_fio','___'), data.get('oluvchi_fio','___'))

    elif doc_type == "oldi_sotdi":
        add_header(doc, "OLDI-SOTDI SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Sotuvchi:", data.get('sotuvchi_fio','___')),
            ("Pasport:", data.get('sotuvchi_pasport','___')),
            ("Xaridor:", data.get('xaridor_fio','___')),
            ("Pasport:", data.get('xaridor_pasport','___')),
            ("Mulk:", data.get('tovar','___')),
            ("Narxi:", f"{data.get('narx','___')} so'm"),
        ])
        add_signature_table(doc, "SOTUVCHI", "XARIDOR",
            data.get('sotuvchi_fio','___'), data.get('xaridor_fio','___'))

    elif doc_type == "tilxat":
        add_header(doc, "TILXAT")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Muallif:", data.get('fio','___')),
            ("Pasport:", data.get('pasport','___')),
            ("Manzil:", data.get('manzil','___')),
            ("Mazmun:", data.get('mazmun','___')),
            ("Miqdor:", f"{data.get('miqdor','___')} so'm"),
        ])
        add_footer(doc)

    elif doc_type == "ishonchnoma":
        add_header(doc, "ISHONCHNOMA")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Beruvchi:", data.get('beruvchi_fio','___')),
            ("Pasport:", data.get('beruvchi_pasport','___')),
            ("Oluvchi:", data.get('oluvchi_fio','___')),
            ("Pasport:", data.get('oluvchi_pasport','___')),
            ("Vakolat:", data.get('vakolat','___')),
            ("Muddat:", data.get('muddat','___')),
        ])
        add_signature_table(doc, "BERUVCHI", "OLUVCHI",
            data.get('beruvchi_fio','___'), data.get('oluvchi_fio','___'))

    elif doc_type == "nikoh":
        add_header(doc, "NIKOH SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Er:", data.get('er_fio','___')),
            ("Pasport:", data.get('er_pasport','___')),
            ("Xotin:", data.get('xotin_fio','___')),
            ("Pasport:", data.get('xotin_pasport','___')),
            ("Mulkiy shartlar:", data.get('shartlar','___')),
        ])
        add_signature_table(doc, "ER", "XOTIN",
            data.get('er_fio','___'), data.get('xotin_fio','___'))

    elif doc_type == "qarz":
        add_header(doc, "QARZ SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Qarz beruvchi:", data.get('beruvchi_fio','___')),
            ("Pasport:", data.get('beruvchi_pasport','___')),
            ("Qarz oluvchi:", data.get('oluvchi_fio','___')),
            ("Pasport:", data.get('oluvchi_pasport','___')),
            ("Miqdor:", f"{data.get('miqdor','___')} so'm"),
            ("Muddat:", data.get('muddat','___')),
            ("Foiz:", f"{data.get('foiz','0')}%"),
        ])
        add_signature_table(doc, "QARZ BERUVCHI", "QARZ OLUVCHI",
            data.get('beruvchi_fio','___'), data.get('oluvchi_fio','___'))

    elif doc_type == "mehnat":
        add_header(doc, "MEHNAT SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Ish beruvchi:", data.get('ish_beruvchi','___')),
            ("Xodim:", data.get('xodim_fio','___')),
            ("Pasport:", data.get('xodim_pasport','___')),
            ("Lavozim:", data.get('lavozim','___')),
            ("Oylik maosh:", f"{data.get('oylik','___')} so'm"),
            ("Muddat:", data.get('muddat','___')),
        ])
        add_signature_table(doc, "ISH BERUVCHI", "XODIM",
            data.get('ish_beruvchi','___'), data.get('xodim_fio','___'))

    elif doc_type == "pudrat":
        add_header(doc, "PUDRAT SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("Buyurtmachi:", data.get('buyurtmachi','___')),
            ("Pudratchi:", data.get('pudratchi','___')),
            ("Ish tavsifi:", data.get('ish_tavsifi','___')),
            ("Narxi:", f"{data.get('narx','___')} so'm"),
            ("Muddat:", data.get('muddat','___')),
        ])
        add_signature_table(doc, "BUYURTMACHI", "PUDRATCHI",
            data.get('buyurtmachi','___'), data.get('pudratchi','___'))

    elif doc_type == "aliment":
        add_header(doc, "ALIMENT SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("To'lovchi:", data.get('tolovchi','___')),
            ("Pasport:", data.get('tolovchi_pasport','___')),
            ("Oluvchi:", data.get('oluvchi','___')),
            ("Farzand:", data.get('bola_ismi','___')),
            ("Oylik miqdor:", f"{data.get('miqdor','___')} so'm"),
            ("Muddat:", data.get('muddat','___')),
        ])
        add_signature_table(doc, "TO'LOVCHI", "OLUVCHI",
            data.get('tolovchi','___'), data.get('oluvchi','___'))

    elif doc_type == "hamkorlik":
        add_header(doc, "HAMKORLIK SHARTNOMASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Shahar:", data.get('shahar','')),
            ("1-tomon:", data.get('tomon1','___')),
            ("2-tomon:", data.get('tomon2','___')),
            ("Loyiha:", data.get('loyiha','___')),
            ("Foyda:", data.get('foyda_taqsim','___')),
            ("Muddat:", data.get('muddat','___')),
        ])
        add_signature_table(doc, "1-TOMON", "2-TOMON",
            data.get('tomon1','___'), data.get('tomon2','___'))

    elif doc_type == "davo":
        add_header(doc, "DA'VO ARIZASI")
        add_info_table(doc, [
            ("Sana:", sana), ("Sud:", f"{data.get('shahar','___')} tumani"),
            ("Ariza beruvchi:", data.get('ariza_beruvchi','___')),
            ("Pasport:", data.get('pasport','___')),
            ("Javobgar:", data.get('javobgar','___')),
            ("Muammo:", data.get('muammo','___')),
            ("Talab:", data.get('talab','___')),
        ])
        add_footer(doc)

    add_footer(doc)
    path = f"{DOCS_DIR}/{doc_type}_{order_id}.docx"
    doc.save(path)
    return path
    # ─── TILLAR ──────────────────────────────────────────────────────────────────

TEXTS = {
    "uz": {
        "welcome": "⚖️ <b>Yuridik Yordamchi Botiga Xush Kelibsiz!</b>\n\n📄 Hujjatlar konstruktori\n🤖 AI huquqiy maslahat\n\nXizmatni tanlang 👇",
        "info": "ℹ️ <b>Xizmatlar narxi:</b>\n\n• Barcha hujjatlar — 10,000 so'm\n• AI Maslahat — 10,000 so'm",
        "ai_start": "🤖 <b>AI Huquqiy Maslahat</b>\n\nHuquqiy muammoingizni yozing:",
        "thinking": "🤔 Tahlil qilinmoqda...",
        "ai_error": "❌ AI xizmati hozircha mavjud emas.",
        "paid_ok": "✅ <b>To'lov tasdiqlandi! Hujjatingiz tayyor.</b>\n\nWord (.docx) formatida yuborildi.",
        "preparing": "⏳ Hujjat tayyorlanmoqda...",
        "cancel": "❌ Bekor qilish",
        "back": "⬅️ Orqaga",
        "pay_btn": "💳 To'lash (Demo)",
        "confirm_btn": "✅ Tasdiqlash (Demo)",
        "no_access": "❌ Ruxsat yo'q.",
        "choose_lang": "🌐 Tilni tanlang:",
        "lang_saved": "✅ Til saqlandi!",
        "broadcast_ask": "📢 Xabarni yozing:",
        "broadcast_done": "✅ {n} ta foydalanuvchiga yuborildi.",
    },
    "ru": {
        "welcome": "⚖️ <b>Добро пожаловать!</b>\n\n📄 Конструктор документов\n🤖 AI консультация\n\nВыберите услугу 👇",
        "info": "ℹ️ <b>Цены:</b>\n\n• Все документы — 10,000 сум\n• AI консультация — 10,000 сум",
        "ai_start": "🤖 <b>AI Юридическая Консультация</b>\n\nОпишите проблему:",
        "thinking": "🤔 Анализируется...",
        "ai_error": "❌ AI временно недоступен.",
        "paid_ok": "✅ <b>Оплата подтверждена! Документ готов.</b>",
        "preparing": "⏳ Готовится...",
        "cancel": "❌ Отмена",
        "back": "⬅️ Назад",
        "pay_btn": "💳 Оплатить (Демо)",
        "confirm_btn": "✅ Подтвердить (Демо)",
        "no_access": "❌ Нет доступа.",
        "choose_lang": "🌐 Выберите язык:",
        "lang_saved": "✅ Язык сохранён!",
        "broadcast_ask": "📢 Напишите сообщение:",
        "broadcast_done": "✅ Отправлено {n} пользователям.",
    },
    "en": {
        "welcome": "⚖️ <b>Welcome to Legal Assistant!</b>\n\n📄 Document constructor\n🤖 AI consultation\n\nChoose a service 👇",
        "info": "ℹ️ <b>Prices:</b>\n\n• All documents — 10,000 UZS\n• AI consultation — 10,000 UZS",
        "ai_start": "🤖 <b>AI Legal Consultation</b>\n\nDescribe your issue:",
        "thinking": "🤔 Analyzing...",
        "ai_error": "❌ AI temporarily unavailable.",
        "paid_ok": "✅ <b>Payment confirmed! Document ready.</b>",
        "preparing": "⏳ Preparing...",
        "cancel": "❌ Cancel",
        "back": "⬅️ Back",
        "pay_btn": "💳 Pay (Demo)",
        "confirm_btn": "✅ Confirm (Demo)",
        "no_access": "❌ Access denied.",
        "choose_lang": "🌐 Choose language:",
        "lang_saved": "✅ Language saved!",
        "broadcast_ask": "📢 Write message:",
        "broadcast_done": "✅ Sent to {n} users.",
    }
}

def t(lang, key):
    return TEXTS.get(lang, TEXTS["uz"]).get(key, key)

# ─── KLAVIATURALAR ────────────────────────────────────────────────────────────

def main_menu(lang="uz"):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📄 Hujjat olish", callback_data="menu_docs")],
        [InlineKeyboardButton(text="🤖 AI Maslahat", callback_data="menu_ai")],
        [InlineKeyboardButton(text="ℹ️ Ma'lumot", callback_data="menu_info")],
        [InlineKeyboardButton(text="🌐 Til", callback_data="menu_lang")],
    ])

def cancel_menu(lang="uz"):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=t(lang, "cancel"), callback_data="back_main")],
    ])

def pay_menu(order_id, lang="uz"):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=t(lang, "pay_btn"), callback_data=f"pay_{order_id}")],
        [InlineKeyboardButton(text=t(lang, "back"), callback_data="back_main")],
    ])

def confirm_menu(order_id, lang="uz"):
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=t(lang, "confirm_btn"), callback_data=f"confirm_{order_id}")],
        [InlineKeyboardButton(text=t(lang, "back"), callback_data="back_main")],
    ])

def lang_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇺🇿 O'zbek", callback_data="setlang_uz")],
        [InlineKeyboardButton(text="🇷🇺 Русский", callback_data="setlang_ru")],
        [InlineKeyboardButton(text="🇬🇧 English", callback_data="setlang_en")],
    ])

def admin_menu():
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📊 Statistika", callback_data="admin_stats")],
        [InlineKeyboardButton(text="👥 Foydalanuvchilar", callback_data="admin_users")],
        [InlineKeyboardButton(text="📋 Buyurtmalar", callback_data="admin_orders")],
        [InlineKeyboardButton(text="📢 Xabar tarqatish", callback_data="admin_broadcast")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="back_main")],
    ])

# ─── DATABASE ─────────────────────────────────────────────────────────────────

async def init_db():
    async with aiosqlite.connect(DB) as db:
        await db.execute("""CREATE TABLE IF NOT EXISTS users (
            user_id INTEGER PRIMARY KEY, username TEXT, full_name TEXT,
            language TEXT DEFAULT 'uz', joined_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        await db.execute("""CREATE TABLE IF NOT EXISTS orders (
            id INTEGER PRIMARY KEY AUTOINCREMENT, user_id INTEGER,
            doc_type TEXT, amount INTEGER, status TEXT DEFAULT 'pending',
            data TEXT, file_path TEXT, created_at TEXT DEFAULT CURRENT_TIMESTAMP)""")
        await db.commit()

async def save_user(uid, username, full_name):
    async with aiosqlite.connect(DB) as db:
        await db.execute("INSERT OR IGNORE INTO users (user_id,username,full_name) VALUES (?,?,?)", (uid,username,full_name))
        await db.commit()

async def get_lang(uid):
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT language FROM users WHERE user_id=?", (uid,)) as cur:
            row = await cur.fetchone()
            return row[0] if row else "uz"

async def set_lang(uid, lang):
    async with aiosqlite.connect(DB) as db:
        await db.execute("UPDATE users SET language=? WHERE user_id=?", (lang,uid))
        await db.commit()

async def create_order(uid, doc_type, amount, data):
    async with aiosqlite.connect(DB) as db:
        cur = await db.execute("INSERT INTO orders (user_id,doc_type,amount,data) VALUES (?,?,?,?)",
            (uid, doc_type, amount, json.dumps(data, ensure_ascii=False)))
        await db.commit()
        return cur.lastrowid

async def get_order(order_id):
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT * FROM orders WHERE id=?", (order_id,)) as cur:
            row = await cur.fetchone()
            if row:
                return {"id":row[0],"user_id":row[1],"doc_type":row[2],
                        "amount":row[3],"status":row[4],
                        "data":json.loads(row[5] or "{}"),"file_path":row[6]}

async def update_order(order_id, status, file_path=None):
    async with aiosqlite.connect(DB) as db:
        await db.execute("UPDATE orders SET status=?,file_path=? WHERE id=?", (status,file_path,order_id))
        await db.commit()

async def get_stats():
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT COUNT(*) FROM users") as c:
            users = (await c.fetchone())[0]
        async with db.execute("SELECT COUNT(*),SUM(amount) FROM orders WHERE status='paid'") as c:
            row = await c.fetchone(); orders,revenue = row[0],row[1] or 0
        async with db.execute("SELECT COUNT(*) FROM users WHERE joined_at>=date('now','-1 day')") as c:
            new_today = (await c.fetchone())[0]
    return users,orders,revenue,new_today

async def get_recent_users(limit=10):
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT user_id,username,full_name,joined_at FROM users ORDER BY joined_at DESC LIMIT ?", (limit,)) as cur:
            return await cur.fetchall()

async def get_recent_orders(limit=10):
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT id,user_id,doc_type,amount,status,created_at FROM orders ORDER BY created_at DESC LIMIT ?", (limit,)) as cur:
            return await cur.fetchall()

async def get_all_users():
    async with aiosqlite.connect(DB) as db:
        async with db.execute("SELECT user_id FROM users") as cur:
            return [r[0] for r in await cur.fetchall()]
            # ─── FSM STATELARI ───────────────────────────────────────────────────────────

class Ijara(StatesGroup):
    beruvchi_fio = State(); beruvchi_pasport = State()
    oluvchi_fio = State(); oluvchi_pasport = State()
    manzil = State(); narx = State(); muddat = State(); shahar = State()

class OldiSotdi(StatesGroup):
    sotuvchi_fio = State(); sotuvchi_pasport = State()
    xaridor_fio = State(); xaridor_pasport = State()
    tovar = State(); narx = State(); shahar = State()

class Tilxat(StatesGroup):
    fio = State(); pasport = State(); manzil = State()
    mazmun = State(); miqdor = State(); shahar = State()

class Ishonchnoma(StatesGroup):
    beruvchi_fio = State(); beruvchi_pasport = State()
    oluvchi_fio = State(); oluvchi_pasport = State()
    vakolat = State(); muddat = State(); shahar = State()

class Nikoh(StatesGroup):
    er_fio = State(); er_pasport = State()
    xotin_fio = State(); xotin_pasport = State()
    shartlar = State(); shahar = State()

class Qarz(StatesGroup):
    beruvchi_fio = State(); beruvchi_pasport = State()
    oluvchi_fio = State(); oluvchi_pasport = State()
    miqdor = State(); muddat = State(); foiz = State(); shahar = State()

class Mehnat(StatesGroup):
    ish_beruvchi = State(); xodim_fio = State(); xodim_pasport = State()
    lavozim = State(); oylik = State(); muddat = State(); shahar = State()

class Pudrat(StatesGroup):
    buyurtmachi = State(); pudratchi = State(); pudratchi_pasport = State()
    ish_tavsifi = State(); narx = State(); muddat = State(); shahar = State()

class Aliment(StatesGroup):
    tolovchi = State(); tolovchi_pasport = State()
    oluvchi = State(); bola_ismi = State()
    miqdor = State(); muddat = State(); shahar = State()

class Hamkorlik(StatesGroup):
    tomon1 = State(); tomon1_pasport = State()
    tomon2 = State(); tomon2_pasport = State()
    loyiha = State(); foyda_taqsim = State(); muddat = State(); shahar = State()

class Davo(StatesGroup):
    ariza_beruvchi = State(); pasport = State()
    javobgar = State(); muammo = State(); talab = State(); shahar = State()

class AI(StatesGroup):
    savol = State()

class Broadcast(StatesGroup):
    xabar = State()

# ─── BOT ─────────────────────────────────────────────────────────────────────

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher(storage=MemoryStorage())

@dp.message(CommandStart())
async def start(message: Message):
    await save_user(message.from_user.id, message.from_user.username or "", message.from_user.full_name or "")
    lang = await get_lang(message.from_user.id)
    await message.answer(t(lang, "welcome"), parse_mode="HTML", reply_markup=main_menu(lang))

@dp.callback_query(F.data == "back_main")
async def back_main(cb: CallbackQuery, state: FSMContext):
    await state.clear()
    lang = await get_lang(cb.from_user.id)
    await cb.message.edit_text(t(lang, "welcome"), parse_mode="HTML", reply_markup=main_menu(lang))

@dp.callback_query(F.data == "menu_info")
async def info(cb: CallbackQuery):
    lang = await get_lang(cb.from_user.id)
    await cb.message.edit_text(t(lang, "info"), parse_mode="HTML", reply_markup=main_menu(lang))

@dp.callback_query(F.data == "menu_lang")
async def choose_lang(cb: CallbackQuery):
    lang = await get_lang(cb.from_user.id)
    await cb.message.edit_text(t(lang, "choose_lang"), reply_markup=lang_menu())

@dp.callback_query(F.data.startswith("setlang_"))
async def set_language(cb: CallbackQuery):
    new_lang = cb.data.split("_")[1]
    await set_lang(cb.from_user.id, new_lang)
    await cb.message.edit_text(t(new_lang, "lang_saved"), reply_markup=main_menu(new_lang))

@dp.callback_query(F.data == "menu_docs")
async def docs(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📁 Asosiy hujjatlar", callback_data="cat_asosiy")],
        [InlineKeyboardButton(text="📝 Ish hujjatlari", callback_data="cat_ish")],
        [InlineKeyboardButton(text="🏗️ Qurilish", callback_data="cat_qurilish")],
        [InlineKeyboardButton(text="👨‍👩‍👧 Oila", callback_data="cat_oila")],
        [InlineKeyboardButton(text="🏢 Biznes", callback_data="cat_biznes")],
        [InlineKeyboardButton(text="⚖️ Yuridik", callback_data="cat_yuridik")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="back_main")],
    ])
    await cb.message.edit_text("📂 <b>Kategoriyani tanlang:</b>", parse_mode="HTML", reply_markup=keyboard)

@dp.callback_query(F.data == "cat_asosiy")
async def cat_asosiy(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏠 Ijara shartnomasi", callback_data="doc_ijara")],
        [InlineKeyboardButton(text="🤝 Oldi-sotdi shartnomasi", callback_data="doc_oldi_sotdi")],
        [InlineKeyboardButton(text="✍️ Tilxat", callback_data="doc_tilxat")],
        [InlineKeyboardButton(text="📋 Ishonchnoma", callback_data="doc_ishonchnoma")],
        [InlineKeyboardButton(text="💍 Nikoh shartnomasi", callback_data="doc_nikoh")],
        [InlineKeyboardButton(text="💰 Qarz shartnomasi", callback_data="doc_qarz")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu_docs")],
    ])
    await cb.message.edit_text("📁 <b>Asosiy hujjatlar:</b>", parse_mode="HTML", reply_markup=keyboard)

@dp.callback_query(F.data == "cat_ish")
async def cat_ish(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="📝 Mehnat shartnomasi", callback_data="doc_mehnat")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu_docs")],
    ])
    await cb.message.edit_text("📝 <b>Ish hujjatlari:</b>", parse_mode="HTML", reply_markup=keyboard)

@dp.callback_query(F.data == "cat_qurilish")
async def cat_qurilish(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏗️ Pudrat shartnomasi", callback_data="doc_pudrat")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu_docs")],
    ])
    await cb.message.edit_text("🏗️ <b>Qurilish hujjatlari:</b>", parse_mode="HTML", reply_markup=keyboard)

@dp.callback_query(F.data == "cat_oila")
async def cat_oila(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="👨‍👩‍👧 Aliment shartnomasi", callback_data="doc_aliment")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu_docs")],
    ])
    await cb.message.edit_text("👨‍👩‍👧 <b>Oila hujjatlari:</b>", parse_mode="HTML", reply_markup=keyboard)

@dp.callback_query(F.data == "cat_biznes")
async def cat_biznes(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🏢 Hamkorlik shartnomasi", callback_data="doc_hamkorlik")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu_docs")],
    ])
    await cb.message.edit_text("🏢 <b>Biznes hujjatlari:</b>", parse_mode="HTML", reply_markup=keyboard)

@dp.callback_query(F.data == "cat_yuridik")
async def cat_yuridik(cb: CallbackQuery):
    keyboard = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="⚖️ Da'vo arizasi", callback_data="doc_davo")],
        [InlineKeyboardButton(text="⬅️ Orqaga", callback_data="menu_docs")],
    ])
    await cb.message.edit_text("⚖️ <b>Yuridik hujjatlar:</b>", parse_mode="HTML", reply_markup=keyboard)
    # ─── HUJJAT HANDLERLARI ──────────────────────────────────────────────────────

@dp.callback_query(F.data == "doc_ijara")
async def ijara_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Ijara.beruvchi_fio)
    await cb.message.edit_text("🏠 <b>Ijara Shartnomasi</b>\n\n1️⃣ Ijara beruvchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Ijara.beruvchi_fio)
async def i1(m: Message, state: FSMContext):
    await state.update_data(beruvchi_fio=m.text); await state.set_state(Ijara.beruvchi_pasport)
    await m.answer("2️⃣ Ijara beruvchining pasporti:")

@dp.message(Ijara.beruvchi_pasport)
async def i2(m: Message, state: FSMContext):
    await state.update_data(beruvchi_pasport=m.text); await state.set_state(Ijara.oluvchi_fio)
    await m.answer("3️⃣ Ijara oluvchining to'liq ismi:")

@dp.message(Ijara.oluvchi_fio)
async def i3(m: Message, state: FSMContext):
    await state.update_data(oluvchi_fio=m.text); await state.set_state(Ijara.oluvchi_pasport)
    await m.answer("4️⃣ Ijara oluvchining pasporti:")

@dp.message(Ijara.oluvchi_pasport)
async def i4(m: Message, state: FSMContext):
    await state.update_data(oluvchi_pasport=m.text); await state.set_state(Ijara.manzil)
    await m.answer("5️⃣ Mulkning to'liq manzili:")

@dp.message(Ijara.manzil)
async def i5(m: Message, state: FSMContext):
    await state.update_data(manzil=m.text); await state.set_state(Ijara.narx)
    await m.answer("6️⃣ Oylik ijara haqi (so'mda):")

@dp.message(Ijara.narx)
async def i6(m: Message, state: FSMContext):
    await state.update_data(narx=m.text); await state.set_state(Ijara.muddat)
    await m.answer("7️⃣ Ijara muddati (oyda):")

@dp.message(Ijara.muddat)
async def i7(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text); await state.set_state(Ijara.shahar)
    await m.answer("8️⃣ Shahar:")

@dp.message(Ijara.shahar)
async def i8(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "ijara", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Beruvchi: {data['beruvchi_fio']}\n• Oluvchi: {data['oluvchi_fio']}\n• Manzil: {data['manzil']}\n• Narx: {data['narx']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_oldi_sotdi")
async def os_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(OldiSotdi.sotuvchi_fio)
    await cb.message.edit_text("🤝 <b>Oldi-Sotdi Shartnomasi</b>\n\n1️⃣ Sotuvchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(OldiSotdi.sotuvchi_fio)
async def os1(m: Message, state: FSMContext):
    await state.update_data(sotuvchi_fio=m.text); await state.set_state(OldiSotdi.sotuvchi_pasport)
    await m.answer("2️⃣ Sotuvchining pasporti:")

@dp.message(OldiSotdi.sotuvchi_pasport)
async def os2(m: Message, state: FSMContext):
    await state.update_data(sotuvchi_pasport=m.text); await state.set_state(OldiSotdi.xaridor_fio)
    await m.answer("3️⃣ Xaridorning to'liq ismi:")

@dp.message(OldiSotdi.xaridor_fio)
async def os3(m: Message, state: FSMContext):
    await state.update_data(xaridor_fio=m.text); await state.set_state(OldiSotdi.xaridor_pasport)
    await m.answer("4️⃣ Xaridorning pasporti:")

@dp.message(OldiSotdi.xaridor_pasport)
async def os4(m: Message, state: FSMContext):
    await state.update_data(xaridor_pasport=m.text); await state.set_state(OldiSotdi.tovar)
    await m.answer("5️⃣ Sotilayotgan mulk tavsifi:")

@dp.message(OldiSotdi.tovar)
async def os5(m: Message, state: FSMContext):
    await state.update_data(tovar=m.text); await state.set_state(OldiSotdi.narx)
    await m.answer("6️⃣ Mulk narxi (so'mda):")

@dp.message(OldiSotdi.narx)
async def os6(m: Message, state: FSMContext):
    await state.update_data(narx=m.text); await state.set_state(OldiSotdi.shahar)
    await m.answer("7️⃣ Shahar:")

@dp.message(OldiSotdi.shahar)
async def os7(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "oldi_sotdi", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Sotuvchi: {data['sotuvchi_fio']}\n• Xaridor: {data['xaridor_fio']}\n• Mulk: {data['tovar']}\n• Narx: {data['narx']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_tilxat")
async def tx_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Tilxat.fio)
    await cb.message.edit_text("✍️ <b>Tilxat</b>\n\n1️⃣ To'liq ismi familiyangiz:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Tilxat.fio)
async def tx1(m: Message, state: FSMContext):
    await state.update_data(fio=m.text); await state.set_state(Tilxat.pasport)
    await m.answer("2️⃣ Pasport seriyasi va raqami:")

@dp.message(Tilxat.pasport)
async def tx2(m: Message, state: FSMContext):
    await state.update_data(pasport=m.text); await state.set_state(Tilxat.manzil)
    await m.answer("3️⃣ Yashash manzili:")

@dp.message(Tilxat.manzil)
async def tx3(m: Message, state: FSMContext):
    await state.update_data(manzil=m.text); await state.set_state(Tilxat.mazmun)
    await m.answer("4️⃣ Tilxat mazmuni:")

@dp.message(Tilxat.mazmun)
async def tx4(m: Message, state: FSMContext):
    await state.update_data(mazmun=m.text); await state.set_state(Tilxat.miqdor)
    await m.answer("5️⃣ Miqdori (so'mda):")

@dp.message(Tilxat.miqdor)
async def tx5(m: Message, state: FSMContext):
    await state.update_data(miqdor=m.text); await state.set_state(Tilxat.shahar)
    await m.answer("6️⃣ Shahar:")

@dp.message(Tilxat.shahar)
async def tx6(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "tilxat", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Muallif: {data['fio']}\n• Miqdor: {data['miqdor']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_ishonchnoma")
async def ish_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Ishonchnoma.beruvchi_fio)
    await cb.message.edit_text("📋 <b>Ishonchnoma</b>\n\n1️⃣ Beruvchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Ishonchnoma.beruvchi_fio)
async def ish1(m: Message, state: FSMContext):
    await state.update_data(beruvchi_fio=m.text); await state.set_state(Ishonchnoma.beruvchi_pasport)
    await m.answer("2️⃣ Beruvchining pasporti:")

@dp.message(Ishonchnoma.beruvchi_pasport)
async def ish2(m: Message, state: FSMContext):
    await state.update_data(beruvchi_pasport=m.text); await state.set_state(Ishonchnoma.oluvchi_fio)
    await m.answer("3️⃣ Oluvchining to'liq ismi:")

@dp.message(Ishonchnoma.oluvchi_fio)
async def ish3(m: Message, state: FSMContext):
    await state.update_data(oluvchi_fio=m.text); await state.set_state(Ishonchnoma.oluvchi_pasport)
    await m.answer("4️⃣ Oluvchining pasporti:")

@dp.message(Ishonchnoma.oluvchi_pasport)
async def ish4(m: Message, state: FSMContext):
    await state.update_data(oluvchi_pasport=m.text); await state.set_state(Ishonchnoma.vakolat)
    await m.answer("5️⃣ Beriladigan vakolat:")

@dp.message(Ishonchnoma.vakolat)
async def ish5(m: Message, state: FSMContext):
    await state.update_data(vakolat=m.text); await state.set_state(Ishonchnoma.muddat)
    await m.answer("6️⃣ Muddati:")

@dp.message(Ishonchnoma.muddat)
async def ish6(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text); await state.set_state(Ishonchnoma.shahar)
    await m.answer("7️⃣ Shahar:")

@dp.message(Ishonchnoma.shahar)
async def ish7(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "ishonchnoma", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Beruvchi: {data['beruvchi_fio']}\n• Oluvchi: {data['oluvchi_fio']}\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_nikoh")
async def nikoh_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Nikoh.er_fio)
    await cb.message.edit_text("💍 <b>Nikoh Shartnomasi</b>\n\n1️⃣ Erning to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Nikoh.er_fio)
async def n1(m: Message, state: FSMContext):
    await state.update_data(er_fio=m.text); await state.set_state(Nikoh.er_pasport)
    await m.answer("2️⃣ Erning pasporti:")

@dp.message(Nikoh.er_pasport)
async def n2(m: Message, state: FSMContext):
    await state.update_data(er_pasport=m.text); await state.set_state(Nikoh.xotin_fio)
    await m.answer("3️⃣ Xotinning to'liq ismi:")

@dp.message(Nikoh.xotin_fio)
async def n3(m: Message, state: FSMContext):
    await state.update_data(xotin_fio=m.text); await state.set_state(Nikoh.xotin_pasport)
    await m.answer("4️⃣ Xotinning pasporti:")

@dp.message(Nikoh.xotin_pasport)
async def n4(m: Message, state: FSMContext):
    await state.update_data(xotin_pasport=m.text); await state.set_state(Nikoh.shartlar)
    await m.answer("5️⃣ Mulkiy shartlar:")

@dp.message(Nikoh.shartlar)
async def n5(m: Message, state: FSMContext):
    await state.update_data(shartlar=m.text); await state.set_state(Nikoh.shahar)
    await m.answer("6️⃣ Shahar:")

@dp.message(Nikoh.shahar)
async def n6(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "nikoh", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Er: {data['er_fio']}\n• Xotin: {data['xotin_fio']}\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_qarz")
async def qarz_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Qarz.beruvchi_fio)
    await cb.message.edit_text("💰 <b>Qarz Shartnomasi</b>\n\n1️⃣ Qarz beruvchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Qarz.beruvchi_fio)
async def q1(m: Message, state: FSMContext):
    await state.update_data(beruvchi_fio=m.text); await state.set_state(Qarz.beruvchi_pasport)
    await m.answer("2️⃣ Beruvchining pasporti:")

@dp.message(Qarz.beruvchi_pasport)
async def q2(m: Message, state: FSMContext):
    await state.update_data(beruvchi_pasport=m.text); await state.set_state(Qarz.oluvchi_fio)
    await m.answer("3️⃣ Qarz oluvchining to'liq ismi:")

@dp.message(Qarz.oluvchi_fio)
async def q3(m: Message, state: FSMContext):
    await state.update_data(oluvchi_fio=m.text); await state.set_state(Qarz.oluvchi_pasport)
    await m.answer("4️⃣ Oluvchining pasporti:")

@dp.message(Qarz.oluvchi_pasport)
async def q4(m: Message, state: FSMContext):
    await state.update_data(oluvchi_pasport=m.text); await state.set_state(Qarz.miqdor)
    await m.answer("5️⃣ Qarz miqdori (so'mda):")

@dp.message(Qarz.miqdor)
async def q5(m: Message, state: FSMContext):
    await state.update_data(miqdor=m.text); await state.set_state(Qarz.muddat)
    await m.answer("6️⃣ Qaytarish muddati:")

@dp.message(Qarz.muddat)
async def q6(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text); await state.set_state(Qarz.foiz)
    await m.answer("7️⃣ Foiz (0 bo'lsa 0 yozing):")

@dp.message(Qarz.foiz)
async def q7(m: Message, state: FSMContext):
    await state.update_data(foiz=m.text); await state.set_state(Qarz.shahar)
    await m.answer("8️⃣ Shahar:")

@dp.message(Qarz.shahar)
async def q8(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "qarz", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Beruvchi: {data['beruvchi_fio']}\n• Oluvchi: {data['oluvchi_fio']}\n• Miqdor: {data['miqdor']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_mehnat")
async def mehnat_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Mehnat.ish_beruvchi)
    await cb.message.edit_text("📝 <b>Mehnat Shartnomasi</b>\n\n1️⃣ Ish beruvchi tashkilot/shaxs:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Mehnat.ish_beruvchi)
async def meh1(m: Message, state: FSMContext):
    await state.update_data(ish_beruvchi=m.text); await state.set_state(Mehnat.xodim_fio)
    await m.answer("2️⃣ Xodimning to'liq ismi:")

@dp.message(Mehnat.xodim_fio)
async def meh2(m: Message, state: FSMContext):
    await state.update_data(xodim_fio=m.text); await state.set_state(Mehnat.xodim_pasport)
    await m.answer("3️⃣ Xodimning pasporti:")

@dp.message(Mehnat.xodim_pasport)
async def meh3(m: Message, state: FSMContext):
    await state.update_data(xodim_pasport=m.text); await state.set_state(Mehnat.lavozim)
    await m.answer("4️⃣ Lavozimi:")

@dp.message(Mehnat.lavozim)
async def meh4(m: Message, state: FSMContext):
    await state.update_data(lavozim=m.text); await state.set_state(Mehnat.oylik)
    await m.answer("5️⃣ Oylik maosh (so'mda):")

@dp.message(Mehnat.oylik)
async def meh5(m: Message, state: FSMContext):
    await state.update_data(oylik=m.text); await state.set_state(Mehnat.muddat)
    await m.answer("6️⃣ Shartnoma muddati:")

@dp.message(Mehnat.muddat)
async def meh6(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text); await state.set_state(Mehnat.shahar)
    await m.answer("7️⃣ Shahar:")

@dp.message(Mehnat.shahar)
async def meh7(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "mehnat", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Xodim: {data['xodim_fio']}\n• Lavozim: {data['lavozim']}\n• Oylik: {data['oylik']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_pudrat")
async def pudrat_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Pudrat.buyurtmachi)
    await cb.message.edit_text("🏗️ <b>Pudrat Shartnomasi</b>\n\n1️⃣ Buyurtmachi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Pudrat.buyurtmachi)
async def pud1(m: Message, state: FSMContext):
    await state.update_data(buyurtmachi=m.text); await state.set_state(Pudrat.pudratchi)
    await m.answer("2️⃣ Pudratchi:")

@dp.message(Pudrat.pudratchi)
async def pud2(m: Message, state: FSMContext):
    await state.update_data(pudratchi=m.text); await state.set_state(Pudrat.pudratchi_pasport)
    await m.answer("3️⃣ Pudratchi pasporti/STIR:")

@dp.message(Pudrat.pudratchi_pasport)
async def pud3(m: Message, state: FSMContext):
    await state.update_data(pudratchi_pasport=m.text); await state.set_state(Pudrat.ish_tavsifi)
    await m.answer("4️⃣ Ish tavsifi:")

@dp.message(Pudrat.ish_tavsifi)
async def pud4(m: Message, state: FSMContext):
    await state.update_data(ish_tavsifi=m.text); await state.set_state(Pudrat.narx)
    await m.answer("5️⃣ Ish narxi (so'mda):")

@dp.message(Pudrat.narx)
async def pud5(m: Message, state: FSMContext):
    await state.update_data(narx=m.text); await state.set_state(Pudrat.muddat)
    await m.answer("6️⃣ Bajarish muddati:")

@dp.message(Pudrat.muddat)
async def pud6(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text); await state.set_state(Pudrat.shahar)
    await m.answer("7️⃣ Shahar:")

@dp.message(Pudrat.shahar)
async def pud7(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "pudrat", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• Buyurtmachi: {data['buyurtmachi']}\n• Pudratchi: {data['pudratchi']}\n• Narx: {data['narx']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_aliment")
async def aliment_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Aliment.tolovchi)
    await cb.message.edit_text("👨‍👩‍👧 <b>Aliment Shartnomasi</b>\n\n1️⃣ To'lovchining to'liq ismi:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Aliment.tolovchi)
async def al1(m: Message, state: FSMContext):
    await state.update_data(tolovchi=m.text); await state.set_state(Aliment.tolovchi_pasport)
    await m.answer("2️⃣ To'lovchining pasporti:")

@dp.message(Aliment.tolovchi_pasport)
async def al2(m: Message, state: FSMContext):
    await state.update_data(tolovchi_pasport=m.text); await state.set_state(Aliment.oluvchi)
    await m.answer("3️⃣ Oluvchining to'liq ismi:")

@dp.message(Aliment.oluvchi)
async def al3(m: Message, state: FSMContext):
    await state.update_data(oluvchi=m.text); await state.set_state(Aliment.bola_ismi)
    await m.answer("4️⃣ Farzand(lar) ismi va tug'ilgan yili:")

@dp.message(Aliment.bola_ismi)
async def al4(m: Message, state: FSMContext):
    await state.update_data(bola_ismi=m.text); await state.set_state(Aliment.miqdor)
    await m.answer("5️⃣ Oylik miqdori (so'mda):")

@dp.message(Aliment.miqdor)
async def al5(m: Message, state: FSMContext):
    await state.update_data(miqdor=m.text); await state.set_state(Aliment.muddat)
    await m.answer("6️⃣ Muddat:")

@dp.message(Aliment.muddat)
async def al6(m: Message, state: FSMContext):
    await state.update_data(muddat=m.text); await state.set_state(Aliment.shahar)
    await m.answer("7️⃣ Shahar:")

@dp.message(Aliment.shahar)
async def al7(m: Message, state: FSMContext):
    await state.update_data(shahar=m.text)
    data = await state.get_data(); await state.clear()
    lang = await get_lang(m.from_user.id)
    order_id = await create_order(m.from_user.id, "aliment", PRICE, data)
    await m.answer(f"✅ <b>Ma'lumotlar qabul qilindi!</b>\n\n• To'lovchi: {data['tolovchi']}\n• Farzand: {data['bola_ismi']}\n• Miqdor: {data['miqdor']} so'm\n\n💳 To'lov qiling:", parse_mode="HTML", reply_markup=pay_menu(order_id, lang))

@dp.callback_query(F.data == "doc_hamkorlik")
async def hamkorlik_start(cb: CallbackQuery, state: FSMContext):
    lang = await get_lang(cb.from_user.id)
    await state.set_state(Hamkorlik.tomon1)
    await cb.message.edit_text("🏢 <b>Hamkorlik Shartnomasi</b>\n\n1️⃣ 1-tomon:", parse_mode="HTML", reply_markup=cancel_menu(lang))

@dp.message(Hamkorlik.tomon1)
async def ham1(m: Message, state: FSMContext):
    await state.update_data(tomon1=m.text); await state.set_state(Hamkorlik.tomon1_pasport)
    await m.answer("2️⃣ 1-tomon pasporti/STIR:")

@dp.message(Hamkorlik.tomon1_pasport)
async def ham2(m: Message, state: FSMContext):
    await state.update_data(tomon1_pasport=m.text); await state.set_state(Hamkorlik.tomon2)
    await m.answer("3️⃣ 2-tomon:")

@dp.message(Hamkorlik.tomon2)
asy
