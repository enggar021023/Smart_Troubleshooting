"""
main.py — Bot Telegram Smart Troubleshooting Ammonia Plant
Jalankan: python main.py   (atau via Docker / systemd untuk 24 jam)

Struktur file proyek:
    main.py          ← file ini (bot + logika)
    setup_db.py      ← jalankan SEKALI untuk buat database
    data/
        ammonia.db   ← database SQLite (edit via DB Browser for SQLite)
    .env             ← TELEGRAM_TOKEN dan GROQ_API_KEY
"""

import os
import re
import sqlite3
import logging
from dotenv import load_dotenv
from telegram import Update
from telegram.ext import (
    Application, CommandHandler, MessageHandler, filters, ContextTypes
)
from groq import Groq

# ─────────────────────────────────────────────
# KONFIGURASI
# ─────────────────────────────────────────────
load_dotenv()

TELEGRAM_TOKEN = os.getenv("TELEGRAM_TOKEN")
GROQ_API_KEY   = os.getenv("GROQ_API_KEY")
DB_PATH        = "data/ammonia.db"

client = Groq(api_key=GROQ_API_KEY)

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
)
logger = logging.getLogger(__name__)

# ─────────────────────────────────────────────
# DATABASE HELPER
# (tidak perlu file db_helper.py terpisah)
# ─────────────────────────────────────────────
def _get_conn() -> sqlite3.Connection:
    """Buka koneksi SQLite. Setiap request buka-tutup sendiri (thread-safe)."""
    return sqlite3.connect(DB_PATH, check_same_thread=False)


def find_equipment(query: str) -> dict | None:
    if not os.path.exists(DB_PATH):
        logger.error("Database tidak ditemukan.")
        return None

    conn   = _get_conn()
    cursor = conn.cursor()

    try:
        # Normalisasi query: hilangkan spasi & strip
        query_clean = query.upper().strip()
        query_nospace = query_clean.replace(" ", "").replace("-", "")

        # ── Tahap 1: coba cocokkan seluruh query (tanpa spasi) ke tag_number ──
        cursor.execute('''
            SELECT tag_number, pid_number, service_description, pid_link, interlock_link
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(tag_number), ' ', ''), '-', '') = ?
            LIMIT 1
        ''', (query_nospace,))
        row = cursor.fetchone()
        if row:
            return _row_to_dict(row)

        # ── Tahap 2: coba LIKE (query_nospace ada di dalam tag) ──
        cursor.execute('''
            SELECT tag_number, pid_number, service_description, pid_link, interlock_link
            FROM equipment
            WHERE REPLACE(REPLACE(UPPER(tag_number), ' ', ''), '-', '') LIKE ?
            LIMIT 1
        ''', (f'%{query_nospace}%',))
        row = cursor.fetchone()
        if row:
            return _row_to_dict(row)

        # ── Tahap 3: cari token yang mengandung huruf+angka ──
        # Gabungkan dulu semua token yang berdekatan
        tokens = re.findall(r'[A-Z0-9]+', query_clean)
        for i in range(len(tokens)):
            # coba kombinasi 1, 2, 3 token berurutan
            for j in range(i+1, min(i+4, len(tokens)+1)):
                combo = "".join(tokens[i:j])
                if re.search(r'[A-Z]', combo) and re.search(r'\d', combo):
                    cursor.execute('''
                        SELECT tag_number, pid_number, service_description, pid_link, interlock_link
                        FROM equipment
                        WHERE REPLACE(REPLACE(UPPER(tag_number), ' ', ''), '-', '') = ?
                        LIMIT 1
                    ''', (combo,))
                    row = cursor.fetchone()
                    if row:
                        return _row_to_dict(row)

        # ── Tahap 4: fallback service_description ──
        keywords = [w for w in query_clean.split() if len(w) > 3]
        for kw in keywords:
            cursor.execute('''
                SELECT tag_number, pid_number, service_description, pid_link, interlock_link
                FROM equipment
                WHERE UPPER(service_description) LIKE ?
                LIMIT 1
            ''', (f'%{kw}%',))
            row = cursor.fetchone()
            if row:
                return _row_to_dict(row)

    except sqlite3.Error as e:
        logger.error(f"Database error: {e}")
    finally:
        conn.close()

    return None


def _row_to_dict(row: tuple) -> dict:
    return {
        'tag_number':          row[0],
        'pid_number':          row[1],
        'service_description': row[2],
        'pid_link':            row[3],
        'interlock_link':      row[4],
    }


def get_all_tags() -> list[str]:
    """Ambil semua tag number (untuk debug atau /list command)."""
    if not os.path.exists(DB_PATH):
        return []
    conn = _get_conn()
    cursor = conn.cursor()
    cursor.execute("SELECT tag_number FROM equipment ORDER BY tag_number")
    tags = [r[0] for r in cursor.fetchall()]
    conn.close()
    return tags

# ─────────────────────────────────────────────
# INTENT DETECTION
# ─────────────────────────────────────────────
def detect_intent(query: str) -> str:
    q = query.lower()
    if any(k in q for k in ["setpoint", "parameter", "nilai"]):
        return "parameter"
    elif any(k in q for k in ["cara kerja", "logic", "prinsip"]):
        return "logic"
    elif any(k in q for k in ["interlock", "safety", "trip", "shutdown"]):
        return "interlock"
    elif any(k in q for k in ["lokasi", "posisi", "dimana", "letak", "pid"]):
        return "lokasi"
    else:
        return "general"

# ─────────────────────────────────────────────
# FORMAT RESPONSE
# ─────────────────────────────────────────────
def format_output(eq: dict, intent: str) -> str:
    tag     = eq.get('tag_number', '-')
    pid     = eq.get('pid_number', '-')
    service = eq.get('service_description', '-')
    link    = eq.get('pid_link', '-')
    ilock   = eq.get('interlock_link', '-')

    link_text  = f"[Buka P&ID]({link})"        if link  and link  != '-' else "Belum tersedia"
    ilock_text = f"[Buka Interlock]({ilock})"  if ilock and ilock != '-' else "Belum tersedia"

    if intent == "lokasi":
        return (
            f"📍 *LOKASI P\\&ID*\n"
            f"━━━━━━━━━━━━━━━━━━━━━\n"
            f"Tag Number  : `{tag}`\n"
            f"P\\&ID Number : {pid}\n"
            f"Service     : {service}\n\n"
            f"🔗 Link P\\&ID : {link_text}"
        )

    elif intent == "interlock":
        return (
            f"⚡ *SAFETY / INTERLOCK*\n"
            f"━━━━━━━━━━━━━━━━━━━━━\n"
            f"Tag Number  : `{tag}`\n"
            f"P\\&ID Number : {pid}\n"
            f"Service     : {service}\n\n"
            f"🔗 Link Interlock : {ilock_text}"
        )

    else:  # general / parameter / logic
        return (
            f"📋 *INFORMASI EQUIPMENT*\n"
            f"━━━━━━━━━━━━━━━━━━━━━\n"
            f"Tag Number  : `{tag}`\n"
            f"P\\&ID Number : {pid}\n"
            f"Service     : {service}\n\n"
            f"🔗 Link P\\&ID     : {link_text}\n"
            f"🔗 Link Interlock : {ilock_text}"
        )

# ─────────────────────────────────────────────
# AI FALLBACK (Groq)
# ─────────────────────────────────────────────
async def get_groq_response(user_query: str, eq_context: dict | None = None) -> str:
    try:
        context_text = ""
        if eq_context:
            context_text = (
                f"TAG: {eq_context.get('tag_number', '-')}\n"
                f"P&ID: {eq_context.get('pid_number', '-')}\n"
                f"SERVICE: {eq_context.get('service_description', '-')}\n"
            )

        prompt = (
            "Anda adalah engineer ammonia plant.\n"
            "ATURAN KETAT:\n"
            "- Gunakan HANYA data yang diberikan di bawah\n"
            "- DILARANG mengarang atau menambah asumsi\n"
            "- Jika data tidak ada → jawab: 'Informasi tidak tersedia di database'\n\n"
            f"DATA TERSEDIA:\n{context_text if context_text else 'Tidak ada data.'}\n\n"
            f"PERTANYAAN: {user_query}\n\n"
            "Jawab singkat, teknis, dan jelas dalam Bahasa Indonesia."
        )

        response = client.chat.completions.create(
            messages=[{"role": "user", "content": prompt}],
            model="llama-3.1-8b-instant",
            max_tokens=400,
        )
        return response.choices[0].message.content

    except Exception as e:
        logger.error(f"Groq AI Error: {e}")
        return "⚠️ Sistem AI sedang tidak tersedia."

# ─────────────────────────────────────────────
# TELEGRAM HANDLERS
# ─────────────────────────────────────────────
async def cmd_start(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "👋 *Selamat datang di Smart Troubleshooting Bot\\!*\n\n"
        "Saya membantu mencari informasi P\\&ID dan interlock equipment ammonia plant\\.\n\n"
        "🔍 *Cara penggunaan:*\n"
        "Ketik tag number equipment, contoh:\n"
        "  → `AE 1001`\n"
        "  → `lokasi PIC102`\n"
        "  → `interlock PSV\\-101`\n\n"
        "📌 *Command tersedia:*\n"
        "  /start — Pesan ini\n"
        "  /list  — Tampilkan 20 tag pertama di database\n"
        "  /help  — Panduan lengkap",
        parse_mode="MarkdownV2",
    )


async def cmd_list(update: Update, context: ContextTypes.DEFAULT_TYPE):
    tags = get_all_tags()
    if not tags:
        await update.message.reply_text("⚠️ Database kosong atau belum dibuat\\. Jalankan `setup\\_db\\.py`\\.", parse_mode="MarkdownV2")
        return
    sample = tags[:20]
    tag_list = "\n".join(f"  • `{t}`" for t in sample)
    more = f"\n_\\.\\.\\. dan {len(tags) - 20} tag lainnya_" if len(tags) > 20 else ""
    await update.message.reply_text(
        f"📋 *Daftar Tag Number* \\({len(tags)} total\\):\n{tag_list}{more}",
        parse_mode="MarkdownV2",
    )


async def cmd_help(update: Update, context: ContextTypes.DEFAULT_TYPE):
    await update.message.reply_text(
        "📖 *PANDUAN BOT*\n\n"
        "*Format query yang didukung:*\n"
        "  `AE 1001` — info umum\n"
        "  `lokasi AE 1001` — info P\\&ID\n"
        "  `interlock PSV\\-101` — info safety/interlock\n\n"
        "*Tips:*\n"
        "  \\- Tag bisa ditulis dengan atau tanpa spasi/strip\n"
        "  \\- Database dikelola via DB Browser for SQLite\n"
        "  \\- Setelah edit database, restart bot",
        parse_mode="MarkdownV2",
    )


async def handle_message(update: Update, context: ContextTypes.DEFAULT_TYPE):
    user_msg = update.message.text.strip()
    logger.info(f"Query asli    : {user_msg}")

    await update.message.chat.send_action("typing")

    intent = detect_intent(user_msg)
    eq     = find_equipment(user_msg)

    logger.info(f"Intent        : {intent}")
    logger.info(f"Equipment hit : {eq}")

    if eq:
        reply = format_output(eq, intent)
        await update.message.reply_text(reply, parse_mode="MarkdownV2")
        logger.info(f"  ✓ Ditemukan: {eq['tag_number']}")
    else:
        ai_reply = await get_groq_response(user_msg)
        await update.message.reply_text(
            f"⚠️ *Tag tidak ditemukan di database*\n\n"
            f"💬 Jawaban AI:\n{ai_reply}",
            parse_mode="Markdown",
        )
        logger.info("  ✗ Tag tidak ditemukan, fallback ke AI.")

# ─────────────────────────────────────────────
# MAIN
# ─────────────────────────────────────────────
def main():
    if not TELEGRAM_TOKEN:
        logger.critical("TELEGRAM_TOKEN tidak ditemukan di .env")
        return
    if not os.path.exists(DB_PATH):
        logger.critical(f"Database tidak ada di '{DB_PATH}'. Jalankan: python setup_db.py")
        return

    app = Application.builder().token(TELEGRAM_TOKEN).build()

    app.add_handler(CommandHandler("start", cmd_start))
    app.add_handler(CommandHandler("list",  cmd_list))
    app.add_handler(CommandHandler("help",  cmd_help))
    app.add_handler(MessageHandler(filters.TEXT & ~filters.COMMAND, handle_message))

    logger.info("🤖 Bot berjalan... (Ctrl+C untuk stop)")
    app.run_polling(drop_pending_updates=True)


if __name__ == "__main__":
    main()