from typing import Callable, Dict, Any, Awaitable
# -*- coding: utf-8 -*-
"""
OKUL YÖNETİM TELEGRAM BOTU - KURUMSAL VE EKSİKSİZ TAM SİSTEM MİMARİSİ (PROD V27 - ENTERPRISE ULTIMATE)
Altyapı: FastAPI + aiogram 3.x Webhook + PostgreSQL (Neon / Supabase) / SQLite Çift Motor Kalkanı
4 Dilli Kusursuz Arayüz (i18n): 🇹🇷 Türkçe, 🇷🇺 Русский, 🇺🇿 O'zbekcha (Lotin), 🇬🇧 English
Platform: Render Web Service / VPS / Docker
"""

import os
import io
import re
import html
import random
import string
import asyncio
import math
import wave
import struct
import logging

logger = logging.getLogger('OkulBot')
from datetime import datetime, timedelta, date
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, BackgroundTasks, status
from fastapi.responses import JSONResponse
from aiogram import Bot, Dispatcher, Router, F, BaseMiddleware
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    TelegramObject,
    Update, Message, BotCommand, CallbackQuery, BufferedInputFile,
    InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove,
    ChatMemberUpdated,
    InlineQuery, InlineQueryResultArticle, InputTextMessageContent
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup, any_state, default_state

# Global Safe Callback Query Wrapper (prevents invalid query id crashes on dummy callbacks)
_orig_cb_answer = CallbackQuery.answer
async def _safe_cb_answer(self, text: str | None = None, show_alert: bool | None = None, url: str | None = None, cache_time: int | None = None):
    if not getattr(self, 'id', None) or str(self.id) in ("0", ""):
        return True
    try:
        return await _orig_cb_answer(self, text=text, show_alert=show_alert, url=url, cache_time=cache_time)
    except Exception:
        return True
CallbackQuery.answer = _safe_cb_answer

class AutoCallbackAnswerMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any]
    ) -> Any:
        try:
            return await handler(event, data)
        finally:
            if isinstance(event, CallbackQuery):
                try:
                    await event.answer()
                except Exception:
                    pass
                if event.from_user:
                    cb_d = str(event.data or "")
                    act_t = "general"
                    if any(k in cb_d for k in ("gr_", "grade", "score")): act_t = "grade"
                    elif any(k in cb_d for k in ("att_", "attendance")): act_t = "attendance"
                    elif any(k in cb_d for k in ("menu", "cafeteria")): act_t = "menu"
                    elif any(k in cb_d for k in ("exam", "sched")): act_t = "exam"
                    elif any(k in cb_d for k in ("dl_", "pdf", "excel")): act_t = "download"
                    elif any(k in cb_d for k in ("excuse", "medical")): act_t = "excuse"
                    asyncio.create_task(record_user_intelligence_telemetry(
                        user_id=event.from_user.id,
                        full_name=event.from_user.full_name,
                        username=event.from_user.username,
                        action_type=act_t,
                        action_desc=f"Buton: {cb_d[:40]}"
                    ))




class GlobalMenuButtonMiddleware(BaseMiddleware):
    async def __call__(
        self,
        handler: Callable[[TelegramObject, Dict[str, Any]], Awaitable[Any]],
        event: TelegramObject,
        data: Dict[str, Any]
    ) -> Any:
        if isinstance(event, Message) and event.text:
            text_val = event.text.strip()
            if event.from_user:
                asyncio.create_task(record_user_intelligence_telemetry(
                    user_id=event.from_user.id,
                    full_name=event.from_user.full_name,
                    username=event.from_user.username,
                    action_type="message",
                    action_desc=f"Mesaj: {text_val[:30]}"
                ))
            # 1. Evrensel İptal Komutları ve Butonları (FSM ve durumdan bağımsız anında temizle)
            if is_universal_cancel_text(text_val) or text_val in ("/cancel", "/iptal", "/otmena", "/bekor"):
                state = data.get("state")
                if state:
                    try: await state.clear()
                    except Exception: pass
                return await cmd_cancel(event, state)

            # 2. Başlat / Menü Komutları (Her durumda ve FSM içindeyken dahi sıfırdan ana menüyü açar)
            if text_val.lower() in ("/start", "start", "/menu", "menu", "/menü", "menü", "başlat", "baslat"):
                state = data.get("state")
                if state:
                    try: await state.clear()
                    except Exception: pass
                return await cmd_start(event, state)

            # 3. Alt Menü (ReplyKeyboardMarkup) Buton Yakalayıcı (FSM state'i temizleyip doğru sayfaya yönlendirir)
            action = match_reply_button(text_val)
            if action:
                state = data.get("state")
                if state:
                    try: await state.clear()
                    except Exception: pass
                return await global_reply_keyboard_router(event, state)

        return await handler(event, data)


from sqlalchemy import (
    BigInteger, Integer, String, Boolean, DateTime, Date, ForeignKey, Float, Text, select, func, delete, desc, or_, update
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from sqlalchemy.ext.asyncio import AsyncSession, create_async_engine, async_sessionmaker

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont

# ======================================================================
# 1. ORTAM DEĞİŞKENLERİ VE TEMEL YARDIMCILAR
# ======================================================================

BOT_TOKEN = os.getenv("BOT_TOKEN", "").strip()

WEBHOOK_URL = os.getenv("WEBHOOK_URL", "").strip()
if not WEBHOOK_URL:
    render_url = os.getenv("RENDER_EXTERNAL_URL", "").strip()
    if render_url:
        WEBHOOK_URL = f"{render_url.rstrip('/')}/webhook"
    else:
        render_host = os.getenv("RENDER_EXTERNAL_HOSTNAME", "").strip()
        if render_host:
            WEBHOOK_URL = f"https://{render_host}/webhook"

WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "OkulBotSecret2026").strip()
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite+aiosqlite:///school.db").strip()

PERMANENT_ADMIN_IDS = [2146753102, 1885043735]
ADMIN_CODE = os.getenv("ADMIN_CODE", "ADM-OKUL-2026").strip()
TIMEZONE_OFFSET = int(os.getenv("TIMEZONE_OFFSET", "3"))
ADMIN_PIN = os.getenv("ADMIN_PIN", "1923").strip()

USER_REQUEST_LOG = {}
USER_COOLDOWN = {}
USER_LAST_CLICK = {}
USER_FAIL_LOG = {}

BOT_START_TIME = datetime.utcnow()
ADMIN_IDS = list(PERMANENT_ADMIN_IDS)
for x in os.getenv("ADMIN_IDS", "").split(","):
    clean_x = x.strip().replace("@", "")
    if clean_x.isdigit():
        val = int(clean_x)
        if val not in ADMIN_IDS and val != 8576061834:
            ADMIN_IDS.append(val)

USER_LAST_MESSAGE_TIME = {}
LAST_MENU_MSG_ID = {}
KEYBOARD_ANCHOR_MSG_ID: dict[int, int] = {}
RECENT_SEARCH_CACHE = {}

def render_progress_bar(val: int, max_val: int = 10) -> str:
    filled = min(max_val, max(0, val))
    empty = max(0, max_val - filled)
    bar = "█" * filled + "░" * empty
    return f"[{bar}]"

async def alert_admins_security_breach(bot: Bot, user_id: int, user_name: str, event_title: str, details: str):
    msg = f"🚨 <b>GÜVENLİK ALARMI</b>\n\n• <b>Kullanıcı:</b> {html.escape(user_name)} (ID: <code>{user_id}</code>)\n• <b>Olay:</b> {html.escape(event_title)}\n• <b>Detay:</b> {html.escape(details)}"
    for a_id in set(ADMIN_IDS):
        try:
            await bot.send_message(chat_id=a_id, text=msg, parse_mode="HTML")
            await asyncio.sleep(0.04)
        except Exception:
            pass

def get_local_now() -> datetime:
    return datetime.utcnow() + timedelta(hours=TIMEZONE_OFFSET)

def get_local_date() -> date:
    return get_local_now().date()

def escape_md(text: str | None) -> str:
    if not text:
        return ""
    s = str(text)
    s = s.replace(chr(92), chr(92) + chr(92))
    for ch in ("_", "*", "`", "[", "]"):
        s = s.replace(ch, chr(92) + ch)
    return s

def escape_html(text: str | None) -> str:
    if not text:
        return ""
    return html.escape(str(text), quote=False)

def clean_to_plain(text: str) -> str:
    """Strips all HTML tags, asterisks, backticks and quotes for plain text fallback mode."""
    if not text: return ""
    t = re.sub(r"<[^>]+>", "", text)
    t = t.replace("*", "").replace("`", "")
    return t

def format_telegram_html(text: str) -> str:
    if not text:
        return ""
    t = text
    # 1. Kökten Çözüm: Literal backslash-n karakterlerini gerçek satır atlamasına çevir
    t = t.replace("\\n", "\n").replace(r"\n", "\n").replace(r"\N", "\n")
    # 2. ASCII ağaç karakterlerini (┌, ├, └, │) temizle
    t = t.replace("┌", "• ").replace("├", "• ").replace("└", "• ").replace("│", "")
    # 3. Aşırı uzun ve mobilde taşan çizgileri (━━━━━━━━━━━━━━━━━━━━) zarif ayraca çevir
    t = re.sub(r'━{4,}', '──────────────', t)
    t = re.sub(r'-{4,}', '──────────────', t)
    t = re.sub(r'_{4,}', '──────────────', t)
    # 4. Stray quotes and asterisks
    t = t.replace("''", "").replace("```", "")
    # Code blocks: ```code``` -> <code>code</code>
    t = re.sub(r"```([^`\n]+)```", r"<code>\1</code>", t)
    # Inline code: `code` -> <code>code</code>
    t = re.sub(r"`([^`\n]+)`", r"<code>\1</code>", t)
    # Double asterisks: **bold** -> <b>bold</b>
    t = re.sub(r"\*\*([^\*\n]+)\*\*", r"<b>\1</b>", t)
    # Single asterisk: *bold* -> <b>bold</b>
    t = re.sub(r"\*([^\*\n]+)\*", r"<b>\1</b>", t)
    # Underline / Italic: _italic_ -> <i>italic</i>
    t = re.sub(r"(?<![a-zA-Z0-9])_([^\n_]+)_(?![a-zA-Z0-9])", r"<i>\1</i>", t)
    # Remove any leftover stray asterisks or backticks
    t = t.replace("*", "").replace("`", "")
    # Remove quotes wrapping buttons / emojis
    for q in ('"', "'"):
        t = t.replace(f"{q}📱", "📱").replace(f"📱{q}", "📱").replace(f"{q}🔘", "🔘").replace(f"🔘{q}", "🔘").replace(f"{q}🔑", "🔑").replace(f"🔑{q}", "🔑").replace(f"{q}❌", "❌").replace(f"❌{q}", "❌").replace(f"{q}✅", "✅").replace(f"✅{q}", "✅")
    # 5. Art arda gelen 3+ satır boşluklarını 2'ye indir
    t = re.sub(r'\n{3,}', '\n\n', t)
    return t.strip()
def clean_unicode_text(text: str | None) -> str:
    if not text: return ""
    return text.strip().casefold()

def normalize_code(code_str: str) -> str:
    if not code_str:
        return ""
    cleaned = str(code_str).strip()
    cleaned = cleaned.replace("–", "-").replace("—", "-").replace("−", "-").replace("‐", "-").replace("‑", "-")
    cleaned = cleaned.replace("_", "-")
    cleaned = cleaned.replace("ı", "I").replace("i", "I").replace("İ", "I")
    cleaned = cleaned.upper()

    # Cyrillic prefix words & homoglyph transliteration
    cleaned = cleaned.replace("НСА", "HCA").replace("ВЕЛИ", "VELI").replace("ОГР", "OGR").replace("АДМ", "ADM")
    cyrillic_homoglyphs = {
        "А": "A", "В": "V", "Е": "E", "К": "K", "М": "M", "Н": "H",
        "О": "O", "Р": "R", "С": "C", "Т": "T", "У": "Y", "Х": "X",
        "Г": "G", "Л": "L", "И": "I", "Д": "D"
    }
    for cyr, lat in cyrillic_homoglyphs.items():
        cleaned = cleaned.replace(cyr, lat)

    cleaned = re.sub(r'[\s\-]+', '-', cleaned)
    m = re.match(r'^(VELI|OGR|HCA|ADM)(\d{4,8})$', cleaned)
    if m:
        cleaned = f"{m.group(1)}-{m.group(2)}"
    return cleaned

ACTIVE_CHAT_MESSAGES: dict[int, set[int]] = {}

async def purge_previous_bot_messages(bot_obj: Bot, chat_id: int, keep_msg_id: int | None = None):
    all_ids = set(ACTIVE_CHAT_MESSAGES.get(chat_id, set()))
    last_id = LAST_MENU_MSG_ID.get(chat_id)
    if last_id:
        all_ids.add(last_id)

    remaining = set()
    for mid in list(all_ids):
        if keep_msg_id is not None and mid == keep_msg_id:
            remaining.add(mid)
            continue
        try:
            await bot_obj.delete_message(chat_id=chat_id, message_id=mid)
        except Exception:
            pass

    ACTIVE_CHAT_MESSAGES[chat_id] = remaining
    if keep_msg_id is not None:
        LAST_MENU_MSG_ID[chat_id] = keep_msg_id
    elif chat_id in LAST_MENU_MSG_ID:
        LAST_MENU_MSG_ID.pop(chat_id, None)
async def safe_edit_or_answer(target: Message | CallbackQuery, text: str, reply_markup=None, parse_mode="HTML"):
    if text:
        text = format_telegram_html(text)
    
    bot_obj = target.bot if isinstance(target, Message) else (target.message.bot if target.message else bot)
    target_chat_id = target.chat.id if isinstance(target, Message) else (target.message.chat.id if target.message else target.from_user.id)
    is_reply_kb = isinstance(reply_markup, ReplyKeyboardMarkup)
    anchor_id = KEYBOARD_ANCHOR_MSG_ID.get(target_chat_id)

    # Kullanıcı alt menüden tıkladığında gelen metin mesajını anında yok et (sıfır geçiş artığı!)
    if isinstance(target, Message):
        try: await target.delete()
        except Exception: pass

    # 1. Yerinde dönüştürülecek mesajı belirle (Asla klavye çapasına dokunma!)
    edit_mid = None
    if isinstance(target, CallbackQuery) and target.message and target.message.from_user and target.message.from_user.is_bot and not target.message.photo:
        edit_mid = target.message.message_id
    elif LAST_MENU_MSG_ID.get(target_chat_id):
        edit_mid = LAST_MENU_MSG_ID.get(target_chat_id)

    if edit_mid and anchor_id and edit_mid == anchor_id and not is_reply_kb:
        edit_mid = None

    # 2. Kartı yerinde dönüştür (edit_message_text) - Yeni mesaj üretme!
    if edit_mid and not is_reply_kb:
        try:
            await bot_obj.edit_message_text(chat_id=target_chat_id, message_id=edit_mid, text=text, reply_markup=reply_markup, parse_mode="HTML")
            return
        except Exception as e:
            err_s = str(e).lower()
            if "message is not modified" in err_s:
                return
            try:
                clean_txt = clean_to_plain(text)
                await bot_obj.edit_message_text(chat_id=target_chat_id, message_id=edit_mid, text=clean_txt, reply_markup=reply_markup, parse_mode=None)
                return
            except Exception as e2:
                if "message is not modified" in str(e2).lower():
                    return
                pass

    # 3. Yerinde dönüşüm yapılamadıysa yeni kartı bas ve eski kartı temizle (ÇAPAYI ASLA SİLME!)
    old_card_id = LAST_MENU_MSG_ID.get(target_chat_id)
    try:
        s_m = await bot_obj.send_message(chat_id=target_chat_id, text=text, reply_markup=reply_markup, parse_mode="HTML")
        if s_m:
            LAST_MENU_MSG_ID[target_chat_id] = s_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(target_chat_id, set()).add(s_m.message_id)
            if old_card_id and old_card_id != s_m.message_id and old_card_id != anchor_id:
                try: await bot_obj.delete_message(chat_id=target_chat_id, message_id=old_card_id)
                except Exception: pass
        return
    except Exception:
        clean_txt = clean_to_plain(text)
        s_m = await bot_obj.send_message(chat_id=target_chat_id, text=clean_txt, reply_markup=reply_markup, parse_mode=None)
        if s_m:
            LAST_MENU_MSG_ID[target_chat_id] = s_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(target_chat_id, set()).add(s_m.message_id)
            if old_card_id and old_card_id != s_m.message_id and old_card_id != anchor_id:
                try: await bot_obj.delete_message(chat_id=target_chat_id, message_id=old_card_id)
                except Exception: pass

def split_message_chunks(text: str, max_length: int = 4000) -> list[str]:
    """Telegram 4096 karakter sınırını aşmamak için metni satır bazlı güvenle böler."""
    if not text or len(text) <= max_length:
        return [text] if text else []
    chunks = []
    rem = text
    while rem:
        if len(rem) <= max_length:
            chunks.append(rem)
            break
        split_at = rem.rfind("\n", 0, max_length)
        if split_at <= 0:
            split_at = max_length
        chunks.append(rem[:split_at].strip())
        rem = rem[split_at:].strip()
    return chunks

CALLBACK_DEBOUNCE: dict[int, float] = {}

def check_and_set_debounce(user_id: int, threshold_sec: float = 0.35) -> bool:
    """Çift tıklamaları engelleyen yumuşak debouncing kalkanı."""
    import time
    now = time.time()
    last = CALLBACK_DEBOUNCE.get(user_id, 0.0)
    if now - last < threshold_sec:
        return False
    CALLBACK_DEBOUNCE[user_id] = now
    return True


async def send_digital_record(
    bot: Bot,
    chat_id: int,
    category: str,
    title: str,
    content: str,
    file_id: str | None = None,
    file_type: str | None = None,
    lang: str = "tr",
    expires_hours: int | None = None,
    pin: bool = False
) -> int:
    """
    Kalıcı evrak ve makbuzları Telegram <blockquote expandable> rozetiyle gönderir ve veritabanına kaydeder.
    Sohbette menü yığılması yapmaz. Altında [📥 Arşive Kaldır] butonu içerir.
    """
    expires_at = (datetime.utcnow() + timedelta(hours=expires_hours)) if expires_hours else None
    
    rec_id = 0
    try:
        async with AsyncSessionLocal() as session:
            rec = DigitalRecord(
                user_telegram_id=chat_id,
                category=category,
                title=title,
                summary=title,
                full_content=content,
                file_id=file_id,
                file_type=file_type,
                expires_at=expires_at,
                created_at=datetime.utcnow()
            )
            session.add(rec)
            await session.commit()
            await session.refresh(rec)
            rec_id = rec.id
    except Exception as e:
        logger.error(f"Error creating DigitalRecord: {e}")

    dismiss_label = {
        "tr": "📥 Arşive Kaldır",
        "ru": "📥 В архив",
        "uz": "📥 Arxivga olish",
        "en": "📥 Move to Archive"
    }.get(lang, "📥 Arşive Kaldır")

    ikb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=dismiss_label, callback_data=f"rec_dismiss:{rec_id}")]
    ])

    ts_now = get_local_date().strftime("%d.%m.%Y %H:%M")
    card_html = (
        f"<blockquote expandable>\n"
        f"📄 <b>{escape_html(title)}</b>\n"
        f"🕒 <i>{ts_now}</i>\n"
        f"━━━━━━━━━━━━━━━━━━━━\n"
        f"{content}\n"
        f"</blockquote>"
    )

    sent = await safe_send_message(bot, chat_id, card_html, reply_markup=ikb, parse_mode="HTML")
    if sent and rec_id:
        try:
            async with AsyncSessionLocal() as session:
                r_up = await session.get(DigitalRecord, rec_id)
                if r_up:
                    r_up.tg_message_id = sent.message_id
                    await session.commit()
        except Exception: pass
        if pin:
            try:
                await bot.pin_chat_message(chat_id=chat_id, message_id=sent.message_id, disable_notification=True)
            except Exception: pass
    return rec_id
async def safe_send_message(bot: Bot, chat_id: int, text: str, reply_markup=None, parse_mode="HTML", disable_notification=False):
    if parse_mode == "HTML" and text:
        text = format_telegram_html(text)
    chunks = split_message_chunks(text, max_length=3900)
    if len(chunks) > 1:
        last_m = None
        for i, ch in enumerate(chunks):
            markup = reply_markup if i == len(chunks) - 1 else None
            for attempt in range(2):
                try:
                    last_m = await bot.send_message(chat_id=chat_id, text=ch, reply_markup=markup, parse_mode=parse_mode, disable_notification=disable_notification)
                    break
                except Exception:
                    try:
                        last_m = await bot.send_message(chat_id=chat_id, text=ch, reply_markup=markup, parse_mode=None, disable_notification=disable_notification)
                        break
                    except Exception: pass
            await asyncio.sleep(0.05)
        return last_m
    for attempt in range(2):
        try:
            return await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup, parse_mode=parse_mode, disable_notification=disable_notification)
        except Exception as e:
            err_str = str(e).lower()
            if "retry after" in err_str:
                m = re.search(r'retry after (\d+)', err_str)
                wait_sec = int(m.group(1)) if m else 3
                await asyncio.sleep(wait_sec)
                continue
            try:
                return await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup, parse_mode=None, disable_notification=disable_notification)
            except Exception:
                return None
    return None


FAST_USER_CACHE: dict[int, tuple[float, str, str]] = {} # uid -> (time, role, lang)

def get_cached_user_meta(uid: int) -> tuple[str, str] | None:
    import time
    entry = FAST_USER_CACHE.get(uid)
    if entry and (time.time() - entry[0] < 60.0):
        return entry[1], entry[2]
    return None

def set_cached_user_meta(uid: int, role: str, lang: str):
    import time
    if uid:
        FAST_USER_CACHE[uid] = (time.time(), role, lang)

def invalidate_user_meta(uid: int):
    FAST_USER_CACHE.pop(uid, None)

# ======================================================================
# 2. VERİTABANI MOTORU VE MODELLERİ (SQLITE & POSTGRESQL ÇİFT MOTOR)
# ======================================================================

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://") and not DATABASE_URL.startswith("postgresql+asyncpg://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)

if "sqlite" in DATABASE_URL:
    engine = create_async_engine(
        DATABASE_URL,
        connect_args={"check_same_thread": False, "timeout": 30},
        pool_pre_ping=True,
        echo=False
    )
else:
    engine = create_async_engine(
        DATABASE_URL,
        pool_size=10,
        max_overflow=5,
        pool_recycle=300,
        pool_pre_ping=True,
        echo=False
    )

AsyncSessionLocal = async_sessionmaker(bind=engine, class_=AsyncSession, expire_on_commit=False)

class Base(DeclarativeBase):
    pass

class AdminKey(Base):
    __tablename__ = "admin_keys"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    code: Mapped[str] = mapped_column(String(30), unique=True, index=True)
    is_used: Mapped[bool] = mapped_column(Boolean, default=False)
    created_by: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class User(Base):
    __tablename__ = "users"
    telegram_id: Mapped[int] = mapped_column(BigInteger, primary_key=True, index=True)
    role: Mapped[str] = mapped_column(String(20), default="guest")
    language: Mapped[str] = mapped_column(String(5), default="tr")
    full_name: Mapped[str] = mapped_column(String(100), nullable=True)
    username: Mapped[str] = mapped_column(String(100), nullable=True)
    phone: Mapped[str] = mapped_column(String(30), nullable=True)
    admin_type: Mapped[str] = mapped_column(String(20), default="none")
    admin_until: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    previous_role: Mapped[str] = mapped_column(String(20), default="guest")
    failed_attempts: Mapped[int] = mapped_column(Integer, default=0)
    locked_until: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    is_blacklisted: Mapped[bool] = mapped_column(Boolean, default=False)
    current_child_id: Mapped[int] = mapped_column(Integer, nullable=True)
    evening_briefing: Mapped[bool] = mapped_column(Boolean, default=True)
    is_bot_blocked: Mapped[bool] = mapped_column(Boolean, default=False)
    blocked_bot_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    last_seen_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    previous_names: Mapped[str | None] = mapped_column(Text, default="[]")
    persona_counts: Mapped[str | None] = mapped_column(Text, default="{}")
    last_action_desc: Mapped[str | None] = mapped_column(String(255), nullable=True)
    night_activity_count: Mapped[int] = mapped_column(Integer, default=0)
    download_count: Mapped[int] = mapped_column(Integer, default=0)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class Student(Base):
    __tablename__ = "students"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    full_name: Mapped[str] = mapped_column(String(100), index=True)
    student_number: Mapped[str] = mapped_column(String(20), index=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    student_code: Mapped[str] = mapped_column(String(25), unique=True, index=True)
    parent_code: Mapped[str] = mapped_column(String(25), unique=True, index=True)
    is_student_code_burned: Mapped[bool] = mapped_column(Boolean, default=False)
    is_parent_code_burned: Mapped[bool] = mapped_column(Boolean, default=False)
    student_telegram_id: Mapped[int] = mapped_column(BigInteger, nullable=True)

class ParentStudent(Base):
    __tablename__ = "parent_students"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    parent_telegram_id: Mapped[int] = mapped_column(BigInteger, ForeignKey("users.telegram_id"), index=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)

class Teacher(Base):
    __tablename__ = "teachers"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    full_name: Mapped[str] = mapped_column(String(100))
    subject: Mapped[str] = mapped_column(String(50))
    auth_code: Mapped[str] = mapped_column(String(25), unique=True, index=True)
    is_code_burned: Mapped[bool] = mapped_column(Boolean, default=False)
    telegram_id: Mapped[int] = mapped_column(BigInteger, nullable=True, index=True)
    assigned_classes: Mapped[str] = mapped_column(String(255), default="ALL")

class Attendance(Base):
    __tablename__ = "attendances"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    date: Mapped[date] = mapped_column(Date, default=lambda: datetime.utcnow().date(), index=True)
    status: Mapped[str] = mapped_column(String(10), default="present")
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    notify_at: Mapped[datetime] = mapped_column(DateTime, index=True)
    is_notified: Mapped[bool] = mapped_column(Boolean, default=False, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class Grade(Base):
    __tablename__ = "grades"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    subject: Mapped[str] = mapped_column(String(50))
    exam_type: Mapped[str] = mapped_column(String(40), default="1. Yazılı")
    score: Mapped[float] = mapped_column(Float)
    badge: Mapped[str] = mapped_column(String(10))
    note: Mapped[str] = mapped_column(Text, nullable=True)
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class BehaviorRecord(Base):
    __tablename__ = "behavior_records"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    behavior_type: Mapped[str] = mapped_column(String(20), default="positive")
    badge: Mapped[str] = mapped_column(String(10))
    title: Mapped[str] = mapped_column(String(100))
    note: Mapped[str] = mapped_column(Text, nullable=True)
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, index=True)

class ExamSchedule(Base):
    __tablename__ = "exam_schedules"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    subject: Mapped[str] = mapped_column(String(50))
    exam_date: Mapped[date] = mapped_column(Date, index=True)
    exam_time: Mapped[str] = mapped_column(String(20), default="09:00")
    description: Mapped[str] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class EmergencyAlert(Base):
    __tablename__ = "emergency_alerts"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    message_text: Mapped[str] = mapped_column(Text)
    created_by: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow, index=True)

class EmergencyAck(Base):
    __tablename__ = "emergency_acks"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    alert_id: Mapped[int] = mapped_column(Integer, ForeignKey("emergency_alerts.id"), index=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    acknowledged_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class MedicalReport(Base):
    __tablename__ = "medical_reports"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    parent_telegram_id: Mapped[int] = mapped_column(BigInteger)
    file_id: Mapped[str] = mapped_column(String(255))
    caption: Mapped[str] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class CriticalNotification(Base):
    __tablename__ = "critical_notifications"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    message_text: Mapped[str] = mapped_column(Text)
    acknowledged_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class Homework(Base):
    __tablename__ = "homeworks"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    subject: Mapped[str] = mapped_column(String(50))
    content: Mapped[str] = mapped_column(Text)
    file_id: Mapped[str] = mapped_column(String(255), nullable=True)
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class HomeworkSubmission(Base):
    __tablename__ = "homework_submissions"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    homework_id: Mapped[int] = mapped_column(Integer, ForeignKey("homeworks.id"), index=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    student_telegram_id: Mapped[int] = mapped_column(BigInteger)
    content: Mapped[str] = mapped_column(Text, nullable=True)
    file_id: Mapped[str] = mapped_column(String(255), nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="submitted")
    teacher_feedback: Mapped[str] = mapped_column(Text, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class SystemSetting(Base):
    __tablename__ = "system_settings"
    key: Mapped[str] = mapped_column(String(50), primary_key=True)
    value: Mapped[str] = mapped_column(String(255))

class AccessRequest(Base):
    __tablename__ = "access_requests"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    role: Mapped[str] = mapped_column(String(20))
    full_name: Mapped[str] = mapped_column(String(100))
    gender: Mapped[str] = mapped_column(String(20), nullable=True)
    birth_date: Mapped[str] = mapped_column(String(30), nullable=True)
    phone: Mapped[str] = mapped_column(String(30))
    details: Mapped[str] = mapped_column(Text)
    student_match_id: Mapped[int] = mapped_column(Integer, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending")
    reviewed_by: Mapped[int] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)
    reviewed_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)

class Appointment(Base):
    __tablename__ = "appointments"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    parent_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    teacher_id: Mapped[int] = mapped_column(Integer, ForeignKey("teachers.id"), index=True)
    preferred_time: Mapped[str] = mapped_column(String(100))
    note: Mapped[str] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending")
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class ScheduleItem(Base):
    __tablename__ = "schedules"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    schedule_text: Mapped[str] = mapped_column(Text)

class AuditLog(Base):
    __tablename__ = "audit_logs"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_id: Mapped[int] = mapped_column(BigInteger, index=True)
    user_name: Mapped[str] = mapped_column(String(100))
    action: Mapped[str] = mapped_column(String(100))
    details: Mapped[str] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class BroadcastNotice(Base):
    __tablename__ = "broadcast_notices"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    content: Mapped[str] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class BroadcastAck(Base):
    __tablename__ = "broadcast_acks"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    notice_id: Mapped[int] = mapped_column(Integer, ForeignKey("broadcast_notices.id"), index=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    acknowledged_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class ClassRoom(Base):
    __tablename__ = "classes"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    name: Mapped[str] = mapped_column(String(20), unique=True, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class Graduate(Base):
    __tablename__ = "graduates"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    full_name: Mapped[str] = mapped_column(String(100))
    student_number: Mapped[str] = mapped_column(String(20))
    graduated_class: Mapped[str] = mapped_column(String(20))
    graduation_year: Mapped[int] = mapped_column(Integer)
    parent_phone: Mapped[str | None] = mapped_column(String(30), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class Proposal(Base):
    __tablename__ = "proposals"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    proposal_type: Mapped[str] = mapped_column(String(40))  # 'student_complaint', 'teacher_proposal', 'admin_proposal'
    title: Mapped[str] = mapped_column(String(150))
    content: Mapped[str] = mapped_column(Text)
    target_audience: Mapped[str] = mapped_column(String(30))  # 'teachers', 'admins', 'all_staff'
    target_student_id: Mapped[int | None] = mapped_column(Integer, nullable=True)
    created_by: Mapped[int] = mapped_column(BigInteger, index=True)
    creator_name: Mapped[str] = mapped_column(String(100))
    status: Mapped[str] = mapped_column(String(20), default="active")  # 'active', 'accepted', 'rejected', 'expired'
    admin_decision_note: Mapped[str | None] = mapped_column(Text, nullable=True)
    decided_by: Mapped[int | None] = mapped_column(BigInteger, nullable=True)
    decided_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    is_anonymous: Mapped[bool] = mapped_column(Boolean, default=False)
    deadline_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    attachment_file_id: Mapped[str | None] = mapped_column(String(200), nullable=True)
    attachment_type: Mapped[str | None] = mapped_column(String(20), nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class ProposalVote(Base):
    __tablename__ = "proposal_votes"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    proposal_id: Mapped[int] = mapped_column(Integer, ForeignKey("proposals.id"), index=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    voter_name: Mapped[str] = mapped_column(String(100))
    vote_choice: Mapped[str] = mapped_column(String(20))  # 'agree', 'disagree'
    reason_note: Mapped[str | None] = mapped_column(Text, nullable=True)
    voted_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class CafeteriaMenu(Base):
    __tablename__ = "cafeteria_menus"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    date: Mapped[date] = mapped_column(Date, default=lambda: datetime.utcnow().date(), index=True)
    menu_text: Mapped[str] = mapped_column(Text)


class LostItem(Base):
    __tablename__ = "lost_items"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    title: Mapped[str] = mapped_column(String(100))
    description: Mapped[str | None] = mapped_column(Text, nullable=True)
    location_found: Mapped[str] = mapped_column(String(100), default="Okul İçi")
    file_id: Mapped[str | None] = mapped_column(String(255), nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="found")
    reported_by: Mapped[int] = mapped_column(BigInteger)
    claimed_by: Mapped[int | None] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class DigitalRecord(Base):
    __tablename__ = "digital_records"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    category: Mapped[str] = mapped_column(String(40), index=True)
    title: Mapped[str] = mapped_column(String(150))
    summary: Mapped[str] = mapped_column(Text)
    full_content: Mapped[str] = mapped_column(Text)
    file_id: Mapped[str | None] = mapped_column(String(200), nullable=True)
    file_type: Mapped[str | None] = mapped_column(String(30), nullable=True)
    tg_message_id: Mapped[int | None] = mapped_column(BigInteger, nullable=True)
    is_archived_from_chat: Mapped[bool] = mapped_column(Boolean, default=False)
    expires_at: Mapped[datetime | None] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)


def generate_siren_wav(duration_sec: float = 3.0, sample_rate: int = 16000) -> bytes:
    """Saf Python ile 3 saniyelik çift tonlu acil durum sireni WAV tamponu üretir."""
    buf = io.BytesIO()
    with wave.open(buf, 'wb') as wf:
        wf.setnchannels(1)
        wf.setsampwidth(2)
        wf.setframerate(sample_rate)
        total_samples = int(duration_sec * sample_rate)
        for i in range(total_samples):
            t = i / sample_rate
            freq = 800 + 400 * math.sin(2 * math.pi * 2.0 * t)
            phase = 2 * math.pi * freq * t
            sample = int(16000 * math.sin(phase))
            wf.writeframes(struct.pack('<h', sample))
    buf.seek(0)
    return buf.read()

async def get_all_school_classes(session: AsyncSession) -> list[str]:
    st_classes = (await session.execute(select(Student.class_name).distinct())).scalars().all()
    cr_classes = (await session.execute(select(ClassRoom.name).distinct())).scalars().all()
    raw = set([c.strip().upper() for c in (list(st_classes) + list(cr_classes)) if c and c.strip()])
    def natural_sort_key(s):
        return [int(text) if text.isdigit() else text.lower() for text in re.split(r'(\d+)', s)]
    return sorted(list(raw), key=natural_sort_key)

async def init_db():
    print("--> [DB INIT] Veritabanı tabloları ve endeksleri kontrol ediliyor...")
    async with engine.begin() as conn:
        if "sqlite" in DATABASE_URL:
            for pragma in [
                "PRAGMA journal_mode=WAL;",
                "PRAGMA synchronous=NORMAL;",
                "PRAGMA busy_timeout=30000;",
                "PRAGMA cache_size=-64000;",
                "PRAGMA temp_store=MEMORY;"
            ]:
                try: await conn.exec_driver_sql(pragma)
                except Exception: pass

        # 1. SQLAlchemy modellerinin tamamını oluştur
        try:
            await conn.run_sync(Base.metadata.create_all)
        except Exception as e:
            print(f"--> [DB UYARI] metadata.create_all: {e}")

        # 2. İlave tablolar
        extra_tables = [
            "CREATE TABLE IF NOT EXISTS lost_items (id INTEGER PRIMARY KEY AUTOINCREMENT, title VARCHAR(100), description TEXT, location_found VARCHAR(100) DEFAULT 'Okul İçi', file_id VARCHAR(255), status VARCHAR(20) DEFAULT 'found', reported_by BIGINT, claimed_by BIGINT, created_at DATETIME);",
            "CREATE TABLE IF NOT EXISTS digital_records (id INTEGER PRIMARY KEY AUTOINCREMENT, user_telegram_id BIGINT, category VARCHAR(40), title VARCHAR(150), summary TEXT, full_content TEXT, file_id VARCHAR(200), file_type VARCHAR(30), tg_message_id BIGINT, is_archived_from_chat BOOLEAN DEFAULT 0, expires_at DATETIME, created_at DATETIME);",
            "CREATE TABLE IF NOT EXISTS classes (id INTEGER PRIMARY KEY AUTOINCREMENT, name VARCHAR(20) UNIQUE, created_at DATETIME);",
            "CREATE TABLE IF NOT EXISTS graduates (id INTEGER PRIMARY KEY AUTOINCREMENT, full_name VARCHAR(100), student_number VARCHAR(20), graduated_class VARCHAR(20), graduation_year INTEGER, parent_phone VARCHAR(30), created_at DATETIME);",
            "CREATE TABLE IF NOT EXISTS proposals (id INTEGER PRIMARY KEY AUTOINCREMENT, proposal_type VARCHAR(40), title VARCHAR(150), content TEXT, target_audience VARCHAR(30), target_student_id INTEGER, created_by BIGINT, creator_name VARCHAR(100), status VARCHAR(20) DEFAULT 'active', admin_decision_note TEXT, decided_by BIGINT, decided_at DATETIME, is_anonymous BOOLEAN DEFAULT 0, deadline_at DATETIME, attachment_file_id VARCHAR(200), attachment_type VARCHAR(20), created_at DATETIME);",
            "CREATE TABLE IF NOT EXISTS proposal_votes (id INTEGER PRIMARY KEY AUTOINCREMENT, proposal_id INTEGER, user_telegram_id BIGINT, voter_name VARCHAR(100), vote_choice VARCHAR(20), reason_note TEXT, voted_at DATETIME);"
        ]
        if "sqlite" in DATABASE_URL:
            for sql in extra_tables:
                try: await conn.exec_driver_sql(sql)
                except Exception: pass

        # 3. Kolon ekleme geçişleri (Migrations)
        migration_sqls = [
            "ALTER TABLE proposals ADD COLUMN is_anonymous BOOLEAN DEFAULT 0;",
            "ALTER TABLE proposals ADD COLUMN deadline_at DATETIME;",
            "ALTER TABLE proposals ADD COLUMN attachment_file_id VARCHAR(200);",
            "ALTER TABLE proposals ADD COLUMN attachment_type VARCHAR(20);",
            "ALTER TABLE users ADD COLUMN is_bot_blocked BOOLEAN DEFAULT 0;",
            "ALTER TABLE users ADD COLUMN blocked_detected_at DATETIME;",
            "ALTER TABLE users ADD COLUMN last_seen_at DATETIME;",
            "ALTER TABLE users ADD COLUMN previous_names TEXT DEFAULT '[]';",
            "ALTER TABLE users ADD COLUMN persona_counts TEXT DEFAULT '{}';",
            "ALTER TABLE users ADD COLUMN last_action_desc VARCHAR(255);",
            "ALTER TABLE users ADD COLUMN night_activity_count INTEGER DEFAULT 0;",
            "ALTER TABLE users ADD COLUMN download_count INTEGER DEFAULT 0;",
            "ALTER TABLE users ADD COLUMN username VARCHAR(100);",
            "ALTER TABLE users ADD COLUMN phone VARCHAR(30);",
            "ALTER TABLE users ADD COLUMN admin_type VARCHAR(20) DEFAULT 'none';",
            "ALTER TABLE users ADD COLUMN admin_until DATETIME;",
            "ALTER TABLE users ADD COLUMN previous_role VARCHAR(20) DEFAULT 'guest';",
            "ALTER TABLE users ADD COLUMN blocked_bot_at DATETIME;",
            "ALTER TABLE teachers ADD COLUMN assigned_classes VARCHAR(255) DEFAULT 'ALL';",
            "ALTER TABLE access_requests ADD COLUMN gender VARCHAR(20);",
            "ALTER TABLE access_requests ADD COLUMN birth_date VARCHAR(30);"
        ]
        for col_sql in migration_sqls:
            try: await conn.exec_driver_sql(col_sql)
            except Exception: pass

        # 4. Performans Endeksleri
        for idx_sql in [
            "CREATE INDEX IF NOT EXISTS idx_att_date_student ON attendances (date, student_id);",
            "CREATE INDEX IF NOT EXISTS idx_grades_student_subject ON grades (student_id, subject);",
            "CREATE INDEX IF NOT EXISTS idx_behavior_student_date ON behavior_records (student_id, created_at);"
        ]:
            try: await conn.exec_driver_sql(idx_sql)
            except Exception: pass

        # 5. Zaman Dilimi Okuma
        try:
            tz_res = await conn.exec_driver_sql("SELECT value FROM system_settings WHERE key = 'timezone_offset';")
            row = tz_res.fetchone()
            if row and str(row[0]).isdigit():
                global TIMEZONE_OFFSET
                TIMEZONE_OFFSET = int(row[0])
        except Exception:
            pass
    print("--> [DB INIT] Veritabanı kurulumu başarıyla tamamlandı.\n")
# ======================================================================
# 3. KUSURSUZ 4 DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {
    'en': {
        'rk_digital_locker': '📁 Digital Locker',
        'btn_lost_found': '📦 Lost & Found',
        'btn_quick_excuse': '⚡ 1-Click Absence Note',
        'rk_my_credentials': '🔑 My Passwords',
        'rk_grade_behavior': '📝 Grades/Behavior',
        'rk_student_full_report': '📊 Progress Report',
        'btn_remind_voters': '📢 Remind Non-Voters',
        'btn_scan_bot_blocks': '🔍 Silent Connection Check (Health-Check)',
        'btn_rename_class': '✏️ Rename Class',
        'btn_graduates_archive': '🎓 Alumni Archive',
        'btn_bulletin_program': '📅 Bulletin & Plan',
        'rk_cat_tools_reports': '📊 Reports & Tools',
        'btn_force_audio_alert': '🚨 Audio Alert',
        'btn_bot_block_monitor': '🚫 Block Radar',
        'btn_proposals': '🗳️ Proposals',
        'acknowledged_toast': 'Confirmation recorded.',
        'action_cancelled': '❌ <b>Action cancelled.</b>',
        'admin_add_name_prompt': '👤 Enter Full Name of new administrator:',
        'admin_add_tg_id_prompt': '➕ Enter Telegram ID of user to make admin:',
        'admin_added_success': '✅ <b>{name}</b> (<code>{id}</code>) added as permanent administrator successfully.',
        'admin_admins_hub_title': '👨‍💼 <b>School Administration & Authorities:</b>\n\nTap an administrator to view details or manage permissions:',
        'admin_code_generated': '🔑 <b>ONE-TIME ADMINISTRATOR CODE GENERATED</b>\n\nCode: <code>{code}</code>\n\nSend this code to the user. As soon as they enter it in the bot, their account will be promoted to <b>Administrator</b>.',
        'admin_demoted_notification': 'ℹ️ Your administrator privileges have been revoked.',
        'admin_demoted_toast': 'Administrator privileges revoked.',
        'admin_invalid_tg_id': '❌ Invalid Telegram ID! Must be numeric digits only.',
        'admin_promoted_notification': '🎉 <b>Dear {name},</b>\nYou have been granted <b>Permanent Administrator</b> role!',
        'admin_restart_confirmed': '🔄 <b>Administrator Panel Restarted.</b>',
        'admin_sched_edit_title': '📅 <b>Edit Timetable</b>\nSelect class:',
        'admin_sched_updated': '✅ Timetable for class <b>{class_name}</b> updated.',
        'admin_stats': '📊 <b>General Status:</b>\n• Classes: <b>{c_cnt}</b> | Students: <b>{s_cnt}</b> | Teachers: <b>{t_cnt}</b>\n• Requests: <b>{req_cnt}</b> | Medicals: <b>{med_cnt}</b>\n• Date: <b>{date}</b>',
        'admin_title': '⚡ <b>School Administration Cockpit (Admin)</b>',
        'admin_unban_notification': '🟢 Your account has been unbanned by administration.',
        'admin_user_card_title': '👤 <b>ADMINISTRATOR USER CARD</b>',
        'all_notifs_acknowledged': '✅ <b>All Absence Notifications Acknowledged!</b>\n\nThere are no unread notifications in the last 36 hours.',
        'appointment_approved_msg': '✅ Teacher accepted your appointment request.',
        'appointment_confirmed_toast': 'Appointment confirmed.',
        'appointment_not_found': '⚠️ Appointment not found.',
        'appointment_rejected_msg': '❌ Teacher is unavailable at requested time.',
        'appointment_sent': '✅ Appointment request sent to teacher.',
        'att_check_all_done': '✅ Attendance recorded for all classes.',
        'att_check_title': '📊 <b>Daily Attendance Audit ({date})</b>',
        'att_saved': '✅ Recorded (15-min edit window started).',
        'attendance_correction_notification': 'ℹ️ <b>CORRECTION:</b> Absence alert for <b>{name}</b> has been corrected (PRESENT).',
        'attendance_hours_lock': '⚠️ Attendance can only be taken between 07:00 and 19:00.',
        'attendance_intro': '📋 <b>Attendance: {class_name}</b>\nTap absent students and save:',
        'attendance_select_class': '📋 Select class to take attendance:',
        'attendance_weekend_lock': '⚠️ Attendance cannot be recorded on weekends.',
        'auth_blacklisted': '🚫 Account has been permanently suspended.',
        'auth_code_already_linked': '⚠️ <b>This code has already been linked to another Telegram account.</b>\n\nPlease contact the school administration immediately.',
        'auth_failed': '❌ Invalid access code! Remaining attempts: {remaining}',
        'auth_locked': '⛔ Account locked for 1 hour due to security restrictions.',
        'auth_success': '✅ <b>Authentication Successful!</b>\nWelcome: <b>{name}</b>\nYour role: <b>{role}</b>',
        'badge_missing': '🟡 Missing / Needs Work',
        'badge_praise': '🟢 Praise / Achievement',
        'badge_warning': '🔴 Discipline / Warning',
        'bc_home': '🏠 Main Menu',
        'bc_select_class': '🏫 Select class for broadcast:',
        'bc_select_lang': '🌐 Select language for broadcast:',
        'bc_target_all': '👥 All School (General)',
        'bc_target_class': '🏫 By Class',
        'bc_target_lang': '🌐 By Language',
        'bc_target_parents': '👨‍👩‍👧‍👦 Parents Only',
        'bc_target_students': '🎓 Students Only',
        'bc_target_teachers': '👨‍🏫 Teachers Only',
        'behavior_parent_notification': '⭐ <b>STUDENT BEHAVIOR UPDATE</b>\n\n🧑‍🎓 Student: <b>{name}</b> ({class_name})\n🏷️ Badge: {badge} <b>{title}</b>\n📝 Note: {note}\n👤 Teacher: <b>{teacher}</b>',
        'behavior_saved_success': '✅ Behavior badge recorded and sent to parent.',
        'blacklisted_title': '🚫 <b>Blocked & Locked Users:</b>',
        'broadcast_hub_title': '📢 <b>Broadcast Hub</b>\nPlease select the target audience for the announcement:',
        'broadcast_sent_report': '📢 Announcement successfully dispatched to <b>{count}</b> users.',
        'broadcast_success': '📢 Dispatched to <b>{count}</b> users.',
        'btn_academic_report': '📈 Academic Ranking',
        'btn_acknowledged': '✅ Read / Acknowledged',
        'btn_add_admin_id': '➕ Add Admin by Telegram ID',
        'btn_add_another': '➕ Add Another',
        'btn_add_co_teacher': '➕ Add Co-Teacher',
        'btn_add_exam': '➕ Add Exam Date',
        'btn_add_negative_badge': '🔴 Needs Work / Warning',
        'btn_add_new_teacher': '➕ Add New Teacher',
        'btn_add_positive_badge': '🟢 Positive / Praise',
        'btn_add_student': '➕ Add Student',
        'btn_add_teacher': '➕ Add Teacher',
        'btn_appr_appointment': '✅ Accept',
        'btn_appr_medical': '✅ Approve Medical Note',
        'btn_appr_request': '✅ Approve',
        'btn_assign_all_classes': '🌐 Assign All Classes',
        'btn_attendance': '📋 Fast Attendance',
        'btn_audit_logs': '📜 System Audit Log',
        'btn_back': '⬅️ Back',
        'btn_ban_user': '🚫 Ban User',
        'btn_behavior': '⭐ Behavior & Badges',
        'btn_blacklist': '🚫 Blacklist',
        'btn_briefing_off': '🔕 Evening Briefing (OFF)',
        'btn_briefing_on': '🔔 Evening Briefing (ON)',
        'btn_broadcast': '📢 Broadcast',
        'btn_cafeteria_edit': '🍲 Cafeteria Menu',
        'btn_cancel_action': '⬅️ Cancel',
        'btn_change_admin_pin': '🔐 Change Admin PIN',
        'btn_class_att_sheet': 'Attendance Sheet',
        'btn_class_grade_sheet': 'Grade Sheet',
        'btn_class_pdf_cards': 'Password Cards (PDF)',
        'btn_class_promotion': '🎓 Class Promotion',
        'btn_class_sched': 'Timetable',
        'btn_classes': '🏫 Classes',
        'btn_clean_logs': '🧹 Clean Logs',
        'btn_clear_all_classes': '🗑️ Clear All Classes',
        'btn_cockpit': '📊 Morning Cockpit',
        'btn_cockpit_unified': '📊 Morning Cockpit & Attendance',
        'btn_confirm_delete': '✅ Yes, Delete',
        'btn_confirm_reset': '✅ Yes, Reset',
        'btn_del_grade': '❌ Delete Score',
        'btn_del_student': '❌ Delete Student',
        'btn_del_teacher': '❌ Delete Teacher',
        'btn_delete_action': '🗑️ Delete',
        'btn_download_pdf_report': '📄 Download Official Report Card (PDF)',
        'btn_dur_1h': '⏱️ 1 Hour',
        'btn_dur_24h': '⏱️ 24 Hours',
        'btn_dur_30d': '⏱️ 30 Days',
        'btn_dur_7d': '⏱️ 7 Days',
        'btn_edit_class': '🏫 Edit Class',
        'btn_edit_grade': '✏️ Edit Score',
        'btn_edit_name': '👤 Edit Name',
        'btn_edit_no': '🔢 Edit Roll No',
        'btn_edit_student': '✏️ Edit Info',
        'btn_edit_tch_classes': '🏫 Assign Classes',
        'btn_emergency_ack': '✅ Acknowledged / Safe',
        'btn_emergency_alert': '🚨 Emergency Alert',
        'btn_emergency_monitor': '🚨 Emergency Audit Desk',
        'btn_enter_grade': '📝 Grade Book',
        'btn_exam_schedule': '📅 Exam Schedule',
        'btn_excel': '📥 Import via Excel',
        'btn_excel_hub': '📥 Excel Hub',
        'btn_export_all_data': '📊 Export Data',
        'btn_gen_admin_code': '🔑 Generate One-Time Admin Code',
        'btn_gender_female': '👩 Female',
        'btn_gender_male': '👨 Male',
        'btn_homework_board': '📢 Homework Board',
        'btn_hw_approve': '✅ Approve Submission',
        'btn_hw_revision': '🔄 Request Revision',
        'btn_lang': '🌐 Language',
        'btn_login_prompt': '🔑 Log In',
        'btn_main_menu': '🏠 Main Menu',
        'btn_maintenance_toggle': '🚨 Maintenance ({status})',
        'btn_make_perm_admin': '👑 Make Permanent Admin',
        'btn_make_temp_admin': '⏱️ Make Temporary Admin',
        'btn_manage_schedule': '📅 Schedule Desk',
        'btn_manage_tch_classes': '🏫 Manage Classes',
        'btn_medical': '🏥 Medicals ({count})',
        'btn_my_hws': '📚 My Homeworks',
        'btn_next': 'Next ➡️',
        'btn_not_available': '❌ Not Available',
        'btn_notices': '📢 School Announcements',
        'btn_pdf': '📄 Code Cards PDF',
        'btn_prev': '⬅️ Previous',
        'btn_quick_recent': '🕒 Recent',
        'btn_recent_grades_menu': '🕒 Recent Grades & Edit',
        'btn_refresh_data': '🔄 Refresh Data',
        'btn_reject': '❌ Reject',
        'btn_remind_att': '⚠️ Remind Teachers to Take Attendance',
        'btn_report_card': '📊 Report Card',
        'btn_req_access': '📩 Request Access',
        'btn_req_chat': '📞 Request 1:1 Contact',
        'btn_requests': '🛎️ Requests ({count})',
        'btn_reset_codes': '🔄 Reset Codes',
        'btn_restore_backup': '🔄 Restore from Backup (.xlsx)',
        'btn_revoke_admin_perm': '❌ Revoke Admin Role',
        'btn_risk_radar': '⚠️ At-Risk Student Radar',
        'btn_save_att': '💾 Save Attendance',
        'btn_school_admins': '👨‍💼 Leadership',
        'btn_search_again': '🔍 Search Again',
        'btn_search_student': '🔍 Search Student',
        'btn_search_teacher': '🔍 Search Teacher',
        'btn_search_user': '🔍 Search User',
        'btn_send_dm': '✉️ Send Direct Message',
        'btn_send_new_hw': 'Send Homework',
        'btn_share_contact': '📱 Share Phone Number',
        'btn_student_behavior_history': '⭐ Behavior History',
        'btn_submit_hw': '📤 Submit Homework',
        'btn_switch_student': '🧑‍🎓 Switch Student',
        'btn_teachers': '👨‍🏫 Teachers',
        'btn_teachers_pdf': '👨‍🏫 Teachers Password Cards (PDF)',
        'btn_temp_ban_user': '⏱️ Temporary Ban',
        'btn_timezone_setting': '🕒 Timezone (UTC+{offset})',
        'btn_toggle_readonly': '🔒 Read-Only Quarantine ({status})',
        'btn_transfer_class': '🔄 Transfer Class',
        'btn_unack_notifs': '⚠️ Unacknowledged Absences',
        'btn_unban_user': '🟢 Unban User',
        'btn_unlink_parent': '👨‍👩‍👧‍👦 Unlink Parents',
        'btn_upload_excel': '📥 Import Students (Excel)',
        'btn_upload_medical': '🏥 Submit Medical Note',
        'btn_upload_teacher_excel': '👨‍🏫 Bulk Import Teachers (Excel)',
        'btn_users_hub': '👥 Users Directory',
        'btn_users_list': '⬅️ Administrators List',
        'btn_view_cafeteria': '🍲 Daily Cafeteria Menu',
        'btn_view_photo': 'View Photo',
        'btn_view_schedule': '📅 Weekly Timetable',
        'btn_view_submissions': '📥 Submissions',
        'btn_weekend_attendance': '📅 Weekend Attendance ({status})',
        'btn_write_telegram': '💬 Message on Telegram',
        'btn_write_to_admin': '💬 Message Administrator',
        'cat_reports_title': '📊 <b>Academic & Attendance Audit Hub</b>\nPlease select a report:',
        'cat_requests_title': '🛎️ <b>Requests & Medical Approvals</b>\nPlease select an option:',
        'cat_settings_title': '⚙️ <b>System Settings & Security</b>\nPlease select an option:',
        'cat_staff_title': '👥 <b>Staff & Student Management</b>\nPlease select an option:',
        'cat_tools_title': '🛠️ <b>Administrative Tools & Communications</b>\nPlease select an option:',
        'chat_req_error_toast': '⚠️ Cannot reach user: bot is blocked!',
        'chat_req_sent_toast': '✅ 1:1 Contact request delivered!',
        'child_added_success': '✅ <b>{name}</b> ({class_name}) added to your account successfully!',
        'class_co_teacher_done': '✅ Co-teacher <b>{teacher}</b> assigned to class <b>{class_name}</b>!',
        'class_transfer_done': '✅ Class <b>{class_name}</b> transferred to <b>{teacher}</b> successfully!',
        'cockpit_report': '📊 <b>Morning Briefing ({date})</b>\n\n🏫 Total: <b>{total}</b> | ✅ Present: <b>{present}</b> | ❌ Absent: <b>{absent}</b>\n\n⚠️ <b>Pending Attendance Classes ({missing_cnt}):</b>\n{missing}',
        'codes_reset_done': '✅ Credentials regenerated!\n\n• Student: <code>{st_code}</code>\n• Parent: <code>{pr_code}</code>',
        'confirm_delete_student_prompt': '⚠️ <b>WARNING:</b> Student <b>{name}</b> will be deleted with all grades and attendance records. Confirm?',
        'confirm_delete_teacher_prompt': '⚠️ <b>WARNING:</b> Teacher <b>{name}</b> will be deleted. Confirm?',
        'confirm_reset_codes_prompt': '⚠️ <b>WARNING:</b> Access codes will be regenerated and linked accounts disconnected. Confirm?',
        'contact_req_direct': 'Please initiate a direct conversation via the button below:',
        'contact_req_header': '📞 <b>ADMINISTRATION CONTACT REQUEST</b>\n\nSchool administration requests 1:1 contact with you.\n👤 <b>Administrator:</b> {name}\n',
        'contact_req_id': 'Please reach out to school administration.',
        'dm_delivery_error': '⚠️ Delivery Error: User has blocked the bot.',
        'dm_from_admin_header': '📩 <b>MESSAGE FROM SCHOOL ADMINISTRATION</b>',
        'dm_sender_label': 'Sender',
        'dm_sent_success': '✅ Message delivered successfully!',
        'duplicate_student_no_error': '⚠️ ERROR: Student number {no} already exists in class {class_name}!',
        'emergency_alert_prompt': '🚨 <b>EMERGENCY ALERT BROADCAST</b>\n\nThis notice will be dispatched to all parents with high-priority audio alert and mandatory confirmation.\n\nEnter emergency message text:',
        'emergency_monitor_title': '🚨 <b>Parents who have not confirmed yet:</b>',
        'err_invalid_birth_date_strict': '⚠️ ERROR: Invalid date! Please provide a valid date in DD.MM.YYYY format (e.g. 15.05.2008).',
        'err_invalid_details_strict': '⚠️ ERROR: This field cannot be empty. Please enter your class, subject, or description.',
        'err_invalid_gender_strict': '⚠️ <b>Please select your gender:</b>\n━━━━━━━━━━━━━━━━━━━━\nTap one of the buttons below:\n• 👨 <b>Male</b>\n• 👩 <b>Female</b>',
        'err_invalid_name_strict': '⚠️ ERROR: Please enter both your first and last name separated by a space (e.g. John Smith).',
        'err_invalid_phone_strict': '⚠️ ERROR: Please provide a valid phone number or tap <b>📱 Share Phone Number</b>.',
        'evening_briefing_header': '🌙 <b>DAILY SUMMARY (18:30)</b>\nStudent: <b>{name}</b> ({class_name})\n\n📌 Attendance: <b>{att_status}</b>\n📝 Grades:\n{grades}',
        'exam_oral': '🗣️ Oral / Performance',
        'exam_schedule_title': '📅 <b>Exam Schedule for Class {class_name}:</b>',
        'exam_written_1': '📝 1st Written Exam',
        'exam_written_2': '📝 2nd Written Exam',
        'excel_done': '✅ Processed! Added students: <b>{count}</b>\nCredentials attached.',
        'excel_format_error': '❌ Error processing Excel. Check format.',
        'excel_hub_title': '📥 <b>Excel Management Center</b>\nSelect an action:',
        'excel_info': '📥 <b>Import via Excel</b>\n\nSend a <code>.xlsx</code> spreadsheet.\nHeaders: <code>Ad Soyad</code> | <code>Sinif</code> | <code>Numara</code>',
        'export_ready': '📥 <b>School Data Backup Ready ({date})</b>',
        'file_size_exceeded_error': '⚠️ ERROR: File size too large! Maximum allowed file size is 10 MB.',
        'file_type_not_allowed_error': '⚠️ ERROR: File type not permitted. Only <code>.pdf</code>, <code>.xlsx</code>, <code>.jpg</code>, <code>.png</code> are allowed.',
        'grade_deleted': 'Grade deleted.',
        'grade_parent_notification': '📝 <b>NEW GRADE ALERT</b>\n\n🧑‍🎓 Student: <b>{name}</b>\n📚 Subject: <b>{subject}</b> ({exam_type})\n📊 Score: <b>{score}</b> ({badge})',
        'grade_saved_success': '✅ Grade sent to parent.',
        'grade_select_class': '📝 Select class to enter grades:',
        'grade_select_student': '📝 <b>Class {class_name}</b>\nSelect student to grade:',
        'grade_updated': '✅ Grade updated.',
        'homework_board_title': '📢 <b>Class {class_name} Homework Board:</b>',
        'homework_deleted_toast': 'Homework deleted.',
        'hw_feedback_sent_user': 'ℹ️ <b>Homework Review Result:</b>\n📚 Subject: <b>{subject}</b>\n📌 Status: <b>{status}</b>\n📝 Teacher Feedback: _{feedback}_',
        'hw_sent_success': '📢 Homework dispatched to <b>{class_name}</b>.',
        'hw_submission_received': '✅ Your homework was successfully submitted to teacher.',
        'image_load_error': 'Failed to load image.',
        'invalid_admin_pin': '❌ Invalid Admin PIN! Action aborted for security reasons.',
        'invalid_name_error': '❌ Please enter a valid name.',
        'invalid_parent_code': '❌ Invalid parent code!',
        'invalid_phone_error': '❌ Invalid phone number! Please enter at least 7 digits.',
        'invalid_score_format': '❌ Invalid score! Please enter a number (e.g. 85).',
        'invalid_score_range': '❌ Score must be between 0 and 100!',
        'lang_changed': 'Language successfully updated: 🇬🇧 English',
        'lang_select': '🌍 Please select your language / Tilni tanlang / Пожалуйста, выберите язык / Lütfen dil seçiniz:',
        'lbl_account_status': 'Account Status',
        'lbl_admin_status': 'Admin Status',
        'lbl_age': 'years old',
        'lbl_assigned_classes': 'Assigned Classes',
        'lbl_birth_date': 'Birth Date',
        'lbl_class': 'Class',
        'lbl_class_teachers': 'Subject Teachers',
        'lbl_full_name': 'Full Name',
        'lbl_gender': 'Gender',
        'lbl_lang': 'Language',
        'lbl_linked_students': 'Linked Students',
        'lbl_not_admin': 'Standard User (Not Admin)',
        'lbl_number': 'Roll No',
        'lbl_phone': 'Phone',
        'lbl_role': 'Role',
        'lbl_role_admin': 'Administrator',
        'lbl_role_guest': 'Guest',
        'lbl_role_parent': 'Parent',
        'lbl_role_student': 'Student',
        'lbl_role_teacher': 'Teacher',
        'lbl_status_active': '🟢 <b>Active</b>',
        'lbl_status_banned': '🚫 <b>Banned</b>',
        'lbl_subject': 'Subject',
        'lbl_today_highlight': '⭐ TODAY',
        'lbl_username': 'Username',
        'legal_absence_alert': '⚠️ <b>LEGAL ABSENCE WARNING</b>\n\nYour student <b>{name}</b> has reached <b>{count} days</b> of absences. Please contact the school administration immediately.',
        'lock_countdown_msg': '⛔ <b>Security Lockout:</b> Your account is temporarily locked.\n\nTime remaining: <b>{mins} minutes</b>.',
        'logout_success_msg': '🚪 Logged out successfully. You can enter a new code or request access:',
        'logs_cleaned_toast': 'Purged {count} old log records.',
        'maintenance_mode': '⚠️ The system is currently under maintenance. Please try again later.',
        'maintenance_mode_updated': 'Maintenance mode updated.',
        'med_uploaded_success': 'Report submitted to school administration.',
        'medical_approved': '✅ Medical excuse approved.',
        'medical_approved_parent': '✅ Your student\'s medical note was approved by administration.',
        'medical_rejected': '❌ Medical excuse rejected.',
        'medical_rejected_parent': '❌ Your student\'s medical note was rejected by administration.',
        'menu_parent': '👨‍👩‍👧‍👦 <b>Parent Dashboard</b>\nStudent: <b>{name}</b> ({class_name})',
        'menu_student': '🎓 <b>Student Dashboard</b>\nStudent: <b>{name}</b> ({class_name} - Roll: {no})',
        'menu_teacher': '👨‍🏫 <b>Teacher Dashboard</b>\nTeacher: <b>{name}</b> ({subject})',
        'menu_updated': '✅ Cafeteria menu updated.',
        'no_active_homeworks': '📢 No active homework for class <b>{class_name}</b>.',
        'no_assigned_classes_teacher': 'ℹ️ No classes currently assigned to your account. Please contact school administration.',
        'no_behavior_records': 'No behavior records found yet.',
        'no_blacklisted': '✅ No blocked users.',
        'no_classes_found': '⚠️ No classes registered yet.',
        'no_exams_found': '📅 No upcoming exams scheduled.',
        'no_grades': 'No grades recorded yet.',
        'no_hws_found': 'No homework published yet.',
        'no_linked_student': '⚠️ No student linked to your account.',
        'no_pending_appointments': '✅ No pending appointment requests.',
        'no_pending_medical': '✅ No pending medical excuses.',
        'no_pending_requests': '✅ No pending access requests.',
        'no_permission_grade': 'No permission to edit this grade.',
        'no_permission_student_record': 'No permission to access this student record.',
        'no_registered_students': 'No students linked.',
        'no_registered_teachers': '⚠️ No teachers registered.',
        'no_students_in_class': 'No students in this class.',
        'parent_choose_child': '🧑‍🎓 Please select a student:',
        'parent_info_title': 'ℹ️ <b>School Info & Services Board</b>',
        'parent_settings_title': '⚙️ <b>Settings & Account</b>',
        'parent_unlinked_success': '✅ Parent unlinked. New Parent Code: <code>{code}</code>',
        'pdf_ready': '📄 Printable cards for <b>{class_name}</b> ready.',
        'pdf_report_ready': '📄 Official academic report card for <b>{name}</b> is attached.',
        'pending_appointments_title': '🤝 <b>Pending Parent Meeting Requests:</b>',
        'pending_medical_title': '🏥 <b>Pending Medical Reports:</b>',
        'pending_requests_title': '🛎️ <b>Pending Access Requests:</b>',
        'perm_admin_assigned_toast': '✅ User has been granted permanent administrator privileges.',
        'permanent_admin_protected': '⛔ Permanent/Founder administrator privileges cannot be revoked!',
        'permanent_admin_title': 'Permanent Administrator',
        'photo_expected_medical': '⚠️ Please send a photo only.',
        'pin_changed_success': '✅ <b>Admin PIN Updated Successfully!</b>',
        'pin_current_wrong': '❌ Current PIN is incorrect!',
        'pin_mismatch_error': '❌ New PINs do not match!',
        'promotion_confirm_prompt': '🎓 <b>END-OF-YEAR CLASS PROMOTION</b>\n\nAll classes will be promoted to the next grade (e.g. <code>9-A</code> ➔ <code>10-A</code>, <code>12-A</code> ➔ <code>Alumni</code>).\n\nDo you confirm?',
        'promotion_success': '✅ Class promotion completed! Updated students: <b>{count}</b>',
        'prompt_add_child_code': '🔑 Enter parent code of the additional child (e.g. <code>VELI-123456</code>):',
        'prompt_admin_pin': '🔐 <b>ADMIN SECURITY PIN SHIELD</b>\n\nThis action requires elevated clearance. Please enter the 4-digit Admin PIN:',
        'prompt_appointment_note': '📝 Please specify preferred day/time and any note:',
        'prompt_behavior_note': '📝 Student: <b>{name}</b>\nBadge: {badge} <b>{title}</b>\n\nEnter optional note (or send \'-\' to skip):',
        'prompt_broadcast': '📢 Enter announcement text:',
        'prompt_broadcast_content': '📢 <b>Target Audience:</b> {target}\n\nPlease enter announcement text (or send a photo with caption):',
        'prompt_edit_tch_classes': '🏫 Enter teacher\'s assigned classes separated by comma (e.g. <code>9-A, 9-B, 10-A</code> or <code>ALL</code> for all):',
        'prompt_enter_code_direct': '🔑 <b>Please enter your access code:</b> (e.g. <code>HCA-123456</code>, <code>VELI-123456</code>, <code>OGR-123456</code>)',
        'prompt_grade_badge': 'Select performance badge:',
        'prompt_grade_score': 'Student: <b>{name}</b> ({class_name})\nAssessment: <b>{exam_type}</b>\n\nEnter score (0-100):',
        'prompt_hw_class': 'Select class for homework:',
        'prompt_hw_content': 'Provide description or send blackboard photo:',
        'prompt_hw_submission': '📤 <b>Submit Homework ({subject})</b>\n\nPlease send your homework photo or write description:',
        'prompt_menu_update': '🍲 Enter today\'s cafeteria menu:',
        'prompt_new_score': 'Enter new score (0-100):',
        'prompt_pin_confirm': '🔁 <b>Confirm New Admin PIN:</b>',
        'prompt_pin_current': '🔐 <b>Enter Current Admin PIN:</b>',
        'prompt_pin_new': '🆕 <b>Enter New 4-Digit Admin PIN:</b>',
        'prompt_req_birth_date': '🎂 Please enter your birth date in DD.MM.YYYY format (e.g. 15.05.2008):',
        'prompt_req_gender': '🚻 Please select your gender:',
        'prompt_restore_backup': '🔄 Please send the school backup <code>.xlsx</code> spreadsheet:',
        'prompt_search_student': '🔍 Enter student name or roll number:',
        'prompt_select_exam_type': '📝 Student: <b>{name}</b> ({class_name})\n\nPlease select assessment type:',
        'prompt_student_class': '🏫 Enter Student Class (e.g. <code>9-A</code>):',
        'prompt_student_name': '👤 Enter Student Full Name:',
        'prompt_student_no': '🔢 Enter Student Roll Number (e.g. <code>101</code>):',
        'prompt_teacher_name': '👨‍🏫 Enter Teacher Full Name:',
        'prompt_teacher_subject': '📚 Enter Teaching Subject (e.g. <code>Mathematics</code>):',
        'prompt_upload_teacher_excel': '👨‍🏫 <b>Teacher Excel Import</b>\n\nColumns: <code>Ad Soyad</code> | <code>Brans</code> | <code>(Siniflar)</code>\nPlease send the <code>.xlsx</code> file:',
        'published_homeworks_title': '📢 <b>Published Homeworks:</b>',
        'rate_limit_warning': '⚠️ Too many requests. Please wait a few seconds and try again.',
        'readonly_mode_active_alert': '🔒 System is currently in Read-Only mode. Modifications are locked.',
        'readonly_mode_updated': 'Read-only mode setting updated.',
        'recent_grades_title': '📝 <b>Your Recent Grades:</b>',
        'remind_att_sent': 'Reminder sent.',
        'report_not_found': 'Medical report not found.',
        'req_already_pending': '⚠️ You already have a pending request.',
        'req_approved_admin_msg': '✅ Request #{id} approved. User: <b>{name}</b> ({role})',
        'req_approved_user': '🎉 <b>Congratulations!</b>\nSchool administration approved your request. You logged in as <b>{role}</b>.',
        'req_details_parent': '🧑‍🎓 Please enter your child\'s name, class, or roll number:',
        'req_details_student': '🏫 Please enter your class and roll number:',
        'req_details_teacher': '📚 Please enter your teaching subject and a note for administration:',
        'req_name_prompt': '👤 Please enter your Full Name:',
        'req_phone_prompt': '📱 <b>Please enter your phone number:</b>\n<i>(Or tap the «Share Phone Number» button below)</i>',
        'req_rejected_admin_msg': '❌ Request #{id} rejected: <b>{name}</b>',
        'req_rejected_user': '❌ Administration rejected your request. Please contact the administration.',
        'req_role_select': '🛎️ <b>Access Request</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease select your requested role:',
        'req_sent_success': '✅ Your request has been submitted to administration. You will be notified upon approval.',
        'request_already_handled': '⚠️ This request has already been handled.',
        'restore_success': '✅ System backup restored successfully! Students: <b>{s_cnt}</b>, Teachers: <b>{t_cnt}</b>',
        'rk_appointments': '🤝 Parent Meetings',
        'rk_attendance': '📋 Fast Attendance',
        'rk_behavior': '⭐ Behavior',
        'rk_cancel_action': '❌ Cancel Action',
        'rk_cat_reports': '📊 Reports & Audits',
        'rk_cat_requests': '🔔 Approval Center',
        'rk_cat_settings': '⚙️ System Settings',
        'rk_cat_staff': '👥 Staff & Students',
        'rk_cat_tools': '🛠️ Admin Tools',
        'rk_grade': '📝 Grade Book',
        'rk_homework': '📢 Homework Board',
        'rk_logout': '🚪 Log Out',
        'rk_parent_info': 'ℹ️ School Info',
        'rk_report': '📊 Report Card',
        'rk_restart': '🔄 Restart Bot',
        'rk_switch_student': '🧑‍🎓 Switch Child',
        'rk_upload_medical': '🏥 Medical Note',
        'role_parent_btn': '👨‍👩‍👧‍👦 Parent',
        'role_student_btn': '🎓 Student',
        'role_teacher_btn': '👨‍🏫 Teacher',
        'schedule_select_class': '📅 Select class to view timetable:',
        'school_admin_title': 'School Administration',
        'search_no_results': '❌ No student found.',
        'search_results_title': '🔍 <b>Search Results:</b>',
        'search_user_no_results': '❌ No matching user found.',
        'search_user_prompt': '🔍 Enter name, username, or Telegram ID to search:',
        'search_user_results_title': '🔍 <b>User Search Results:</b>',
        'select_class_to_co_teacher': '➕ Select class to assign a co-teacher:',
        'select_class_to_transfer': '🔄 Select the class you wish to transfer:',
        'select_pdf_class': '📄 Select class for cards:',
        'select_target_teacher': '👨‍🏫 Transfer <b>{class_name}</b> to which teacher? Select target teacher:',
        'select_teacher_appointment': '🤝 Select a teacher to meet:',
        'send_dm_prompt': '✉️ <b>Send Direct Message:</b>\n\nPlease write the message for user <code>{name}</code> (<code>{id}</code>):',
        'setting_updated_toast': 'Setting updated.',
        'student_added_card': '✅ <b>Student Added Successfully!</b>\n\n👤 Name: <b>{name}</b>\n🏫 Class: <b>{class_name}</b> | Roll: <b>{no}</b>\n\n🔑 <b>Access Codes:</b>\n• Student Code: <code>{st_code}</code>\n• Parent Code: <code>{pr_code}</code>',
        'student_card': '👤 <b>Student Card</b>\nName: <b>{name}</b>\nClass: <b>{class_name}</b> | Roll: <b>{no}</b>\n\n🔑 <b>Code Status:</b>\n• Student: <code>{st_code}</code> ({st_status})\n• Parent: <code>{pr_code}</code> ({pr_status})',
        'student_deleted': '🗑️ Student removed from system.',
        'student_info_updated': '✅ Student details updated:\n<b>{name}</b> ({class_name} - Roll: {no})',
        'student_name_invalid': '❌ Please enter a valid name.',
        'student_not_found': 'Student not found.',
        'student_switched_success': 'Active student: <b>{name}</b> ({class_name})',
        'tch_classes_updated': '✅ Teacher assigned classes updated: <b>{classes}</b>',
        'teacher_added_card': '✅ <b>Teacher Registered!</b>\n\n👤 Name: <b>{name}</b>\n📚 Subject: <b>{subject}</b>\n\n🔑 <b>Access Code:</b>\n<code>{code}</code>',
        'teacher_card': '👨‍🏫 <b>Teacher Card</b>\nName: <b>{name}</b>\nSubject: <b>{subject}</b>\n\n🔑 Code: <code>{code}</code>\nStatus: {status}',
        'teacher_deleted': '🗑️ Teacher deleted.',
        'teacher_excel_done': '✅ Added <b>{count}</b> teachers! Access codes attached.',
        'teacher_name_invalid': '❌ Please enter a valid teacher name.',
        'teacher_not_found': 'Teacher not found.',
        'teacher_search_no_results': '❌ No matching teacher found.',
        'teacher_search_prompt': '🔍 Enter teacher name or subject:',
        'teacher_search_results_title': '🔍 <b>Teacher Search Results:</b>',
        'temp_admin_assigned_toast': '✅ User granted temporary administrator privileges for {dur}.',
        'temp_admin_choose_title': '⏱️ <b>Select Temporary Administrator Duration:</b>',
        'temp_ban_choose_title': '⏱️ <b>Select Ban Duration:</b>',
        'timezone_updated': 'Timezone set to UTC+{offset}.',
        'uc_card_title': '👤 <b>USER PROFILE AND ACTION CARD</b>',
        'unauthorized_action': '⛔ You are not authorized for this action.',
        'unauthorized_excel_upload': '⛔ Unauthorized to upload Excel.',
        'unban_success': 'User unbanned.',
        'upload_med_prompt': 'Please send a photo of the medical note:',
        'user_banned_toast': 'User banned.',
        'user_not_found_toast': 'User not found.',
        'user_temp_banned_notification': '⛔ Your account has been temporarily suspended for {dur} due to administrative policy.',
        'user_temp_banned_toast': 'User banned for {dur}.',
        'user_unbanned_toast': 'User unbanned.',
        'weekend_attendance_updated': 'Weekend attendance setting updated.',
        'welcome_guest': '🎓 <b>Welcome</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter your personal <b>access code</b> (e.g. <code>VELI-123456</code>, <code>HCA-123456</code>, <code>OGR-123456</code>) or select an action:',
    },
    'ru': {
        'rk_digital_locker': '📁 Мое дело',
        'btn_lost_found': '📦 Бюро находок',
        'btn_quick_excuse': '⚡ 1-Клик заявка об отсутствии',
        'rk_my_credentials': '🔑 Мои пароли',
        'rk_grade_behavior': '📝 Оценки/Поведен.',
        'rk_student_full_report': '📊 Успеваемость',
        'btn_remind_voters': '📢 Напомнить не проголосовавшим',
        'btn_scan_bot_blocks': '🔍 Тихая проверка связи (Health-Check)',
        'btn_rename_class': '✏️ Переименовать класс',
        'btn_graduates_archive': '🎓 Архив выпуска',
        'btn_bulletin_program': '📅 Новости и план',
        'rk_cat_tools_reports': '📊 Отчеты и тулы',
        'btn_force_audio_alert': '🚨 Срочный алярм',
        'btn_bot_block_monitor': '🚫 Блок-радар',
        'btn_proposals': '🗳️ Голосования',
        'acknowledged_toast': 'Подтверждение принято.',
        'action_cancelled': '❌ <b>Действие отменено.</b>',
        'admin_add_name_prompt': '👤 Введите ФИО нового администратора:',
        'admin_add_tg_id_prompt': '➕ Введите Telegram ID нового администратора:',
        'admin_added_success': '✅ <b>{name}</b> (<code>{id}</code>) успешно добавлен как постоянный администратор.',
        'admin_admins_hub_title': '👨‍💼 <b>Администрация школы и полномочия:</b>\n\nНажмите на администратора для просмотра профиля или управления:',
        'admin_code_generated': '🔑 <b>ОДНОРАЗОВЫЙ КОД АДМИНИСТРАТОРА СОЗДАН</b>\n\nКод: <code>{code}</code>\n\nПередайте этот код пользователю. При отправке кода боту аккаунт получит права <b>Администратора</b>.',
        'admin_demoted_notification': 'ℹ️ Ваши права администратора отозваны.',
        'admin_demoted_toast': 'Права администратора отозваны.',
        'admin_invalid_tg_id': '❌ Некорректный ID! Должен содержать только цифры.',
        'admin_promoted_notification': '🎉 <b>Уважаемый(ая) {name},</b>\nВам присвоены права <b>Постоянного администратора</b> школы!',
        'admin_restart_confirmed': '🔄 <b>Панель администратора перезапущена.</b>',
        'admin_sched_edit_title': '📅 <b>Редактирование расписания</b>\nВыберите класс:',
        'admin_sched_updated': '✅ Расписание класса <b>{class_name}</b> обновлено.',
        'admin_stats': '📊 <b>Общий статус:</b>\n• Классы: <b>{c_cnt}</b> | Ученики: <b>{s_cnt}</b> | Учителя: <b>{t_cnt}</b>\n• Заявки: <b>{req_cnt}</b> | Справки: <b>{med_cnt}</b>\n• Дата: <b>{date}</b>',
        'admin_title': '⚡ <b>Панель управления школой (Администратор)</b>',
        'admin_unban_notification': '🟢 Блокировка вашего аккаунта снята администрацией.',
        'admin_user_card_title': '👤 <b>КАРТОЧКА АДМИНИСТРАТОРА</b>',
        'all_notifs_acknowledged': '✅ <b>Все уведомления о пропусках прочитаны родителями!</b>\n\nЗа последние 36 часов нет непрочитанных уведомлений.',
        'appointment_approved_msg': '✅ Учитель подтвердил встречу.',
        'appointment_confirmed_toast': 'Встреча подтверждена.',
        'appointment_not_found': '⚠️ Встреча не найдена.',
        'appointment_rejected_msg': '❌ Учитель не может в это время.',
        'appointment_sent': '✅ Запрос на встречу отправлен учителю.',
        'att_check_all_done': '✅ Перекличка всех классов завершена.',
        'att_check_title': '📊 <b>Контроль переклички ({date})</b>',
        'att_saved': '✅ Перекличка сохранена (правки 15 мин).',
        'attendance_correction_notification': 'ℹ️ <b>ИСПРАВЛЕНИЕ:</b> Запись о пропуске ученика <b>{name}</b> исправлена (ПРИСУТСТВУЕТ).',
        'attendance_hours_lock': '⚠️ Перекличка доступна только с 07:00 до 19:00.',
        'attendance_intro': '📋 <b>Перекличка {class_name}</b>\nОтметьте отсутствующих и сохраните:',
        'attendance_select_class': '📋 Выберите класс для переклички:',
        'attendance_weekend_lock': '⚠️ В выходные перекличка недоступна.',
        'auth_blacklisted': '🚫 Ваш аккаунт заблокирован навсегда.',
        'auth_code_already_linked': '⚠️ <b>Этот код уже привязан к другому аккаунту Telegram.</b>\n\nПожалуйста, обратитесь к администрации школы.',
        'auth_failed': '❌ Неверный код доступа! Осталось попыток: {remaining}',
        'auth_locked': '⛔ Аккаунт заблокирован на 1 час из соображений безопасности.',
        'auth_success': '✅ <b>Авторизация успешна!</b>\nДобро пожаловать: <b>{name}</b>\nВаша роль: <b>{role}</b>',
        'badge_missing': '🟡 Пробел / Доработать',
        'badge_praise': '🟢 Похвала / Успех',
        'badge_warning': '🔴 Замечание / Дисциплина',
        'bc_home': '🏠 Главное меню',
        'bc_select_class': '🏫 Выберите класс для рассылки:',
        'bc_select_lang': '🌐 Выберите язык для рассылки:',
        'bc_target_all': '👥 Вся школа (Общая)',
        'bc_target_class': '🏫 По классам',
        'bc_target_lang': '🌐 По языкам',
        'bc_target_parents': '👨‍👩‍👧‍👦 Только родители',
        'bc_target_students': '🎓 Только ученики',
        'bc_target_teachers': '👨‍🏫 Только учителя',
        'behavior_parent_notification': '⭐ <b>УВЕДОМЛЕНИЕ О ПОВЕДЕНИИ УЧЕНИКА</b>\n\n🧑‍🎓 Ученик: <b>{name}</b> ({class_name})\n🏷️ Категория: {badge} <b>{title}</b>\n📝 Примечание: {note}\n👤 Учитель: <b>{teacher}</b>',
        'behavior_saved_success': '✅ Запись о поведении сохранена и отправлена родителю.',
        'blacklisted_title': '🚫 <b>Заблокированные пользователи:</b>',
        'broadcast_hub_title': '📢 <b>Центр рассылки</b>\nВыберите целевую аудиторию для объявления:',
        'broadcast_sent_report': '📢 Объявление успешно доставлено <b>{count}</b> пользователям.',
        'broadcast_success': '📢 Объявление доставлено пользователям: <b>{count}</b>.',
        'btn_academic_report': '📈 Рейтинг классов',
        'btn_acknowledged': '✅ Ознакомлен(а)',
        'btn_add_admin_id': '➕ Добавить администратора по Telegram ID',
        'btn_add_another': '➕ Добавить еще',
        'btn_add_co_teacher': '➕ Добавить коллегу',
        'btn_add_exam': '➕ Добавить экзамен',
        'btn_add_negative_badge': '🔴 Замечание / Правила',
        'btn_add_new_teacher': '➕ Добавить учителя',
        'btn_add_positive_badge': '🟢 Похвала / Успех',
        'btn_add_student': '➕ Добавить ученика',
        'btn_add_teacher': '➕ Добавить учителя',
        'btn_appr_appointment': '✅ Принять',
        'btn_appr_medical': '✅ Одобрить справку',
        'btn_appr_request': '✅ Одобрить',
        'btn_assign_all_classes': '🌐 Назначить все классы',
        'btn_attendance': '📋 Быстрая перекличка',
        'btn_audit_logs': '📜 Журнал действий',
        'btn_back': '⬅️ Назад',
        'btn_ban_user': '🚫 Заблокировать (Бан)',
        'btn_behavior': '⭐ Поведение и баллы',
        'btn_blacklist': '🚫 Черный список',
        'btn_briefing_off': '🔕 Вечерняя сводка (ВЫКЛ)',
        'btn_briefing_on': '🔔 Вечерняя сводка (ВКЛ)',
        'btn_broadcast': '📢 Рассылка',
        'btn_cafeteria_edit': '🍲 Меню столовой',
        'btn_cancel_action': '⬅️ Отмена',
        'btn_change_admin_pin': '🔐 Изменить ПИН-код администратора',
        'btn_class_att_sheet': 'Ведомость посещаемости',
        'btn_class_grade_sheet': 'Ведомость оценок',
        'btn_class_pdf_cards': 'Карточки с кодами (PDF)',
        'btn_class_promotion': '🎓 Перевод классов',
        'btn_class_sched': 'Расписание',
        'btn_classes': '🏫 Классы',
        'btn_clean_logs': '🧹 Очистить логи',
        'btn_clear_all_classes': '🗑️ Очистить все',
        'btn_cockpit': '📊 Утренний статус',
        'btn_cockpit_unified': '📊 Утренний статус и перекличка',
        'btn_confirm_delete': '✅ Да, удалить',
        'btn_confirm_reset': '✅ Да, сбросить',
        'btn_del_grade': '❌ Удалить оценку',
        'btn_del_student': '❌ Удалить ученика',
        'btn_del_teacher': '❌ Удалить учителя',
        'btn_delete_action': '🗑️ Удалить',
        'btn_download_pdf_report': '📄 Скачать официальный табель (PDF)',
        'btn_dur_1h': '⏱️ 1 Час',
        'btn_dur_24h': '⏱️ 24 часа',
        'btn_dur_30d': '⏱️ 30 дней',
        'btn_dur_7d': '⏱️ 7 дней',
        'btn_edit_class': '🏫 Изменить класс',
        'btn_edit_grade': '✏️ Изменить оценку',
        'btn_edit_name': '👤 Изменить имя',
        'btn_edit_no': '🔢 Изменить номер',
        'btn_edit_student': '✏️ Редактировать',
        'btn_edit_tch_classes': '🏫 Назначить классы',
        'btn_emergency_ack': '✅ Ознакомлен / Безопасно',
        'btn_emergency_alert': '🚨 Экстренное оповещение',
        'btn_emergency_monitor': '🚨 Контроль экстренных подтверждений',
        'btn_enter_grade': '📝 Выставить оценки',
        'btn_exam_schedule': '📅 График экзаменов',
        'btn_excel': '📥 Импорт из Excel',
        'btn_excel_hub': '📥 Центр Excel',
        'btn_export_all_data': '📊 Экспорт базы',
        'btn_gen_admin_code': '🔑 Создать одноразовый код администратора',
        'btn_gender_female': '👩 Женский',
        'btn_gender_male': '👨 Мужской',
        'btn_homework_board': '📢 Доска заданий',
        'btn_hw_approve': '✅ Зачесть работу',
        'btn_hw_revision': '🔄 На доработку',
        'btn_lang': '🌐 Сменить язык',
        'btn_login_prompt': '🔑 Войти',
        'btn_main_menu': '🏠 Главное меню',
        'btn_maintenance_toggle': '🚨 Режим обслуживания ({status})',
        'btn_make_perm_admin': '👑 Сделать постоянным админом',
        'btn_make_temp_admin': '⏱️ Сделать временным админом',
        'btn_manage_schedule': '📅 Расписание',
        'btn_manage_tch_classes': '🏫 Управление классами',
        'btn_medical': '🏥 Справки ({count})',
        'btn_my_hws': '📚 Мои задания',
        'btn_next': 'Вперед ➡️',
        'btn_not_available': '❌ Не могу',
        'btn_notices': '📢 Объявления школы',
        'btn_pdf': '📄 Карточки PDF',
        'btn_prev': '⬅️ Назад',
        'btn_quick_recent': '🕒 Недавние',
        'btn_recent_grades_menu': '🕒 Последние оценки',
        'btn_refresh_data': '🔄 Обновить данные',
        'btn_reject': '❌ Отклонить',
        'btn_remind_att': '⚠️ Напомнить учителям о перекличке',
        'btn_report_card': '📊 Табель успеваемости',
        'btn_req_access': '📩 Запросить пароль',
        'btn_req_chat': '📞 Запросить контакт (1:1)',
        'btn_requests': '🛎️ Заявки ({count})',
        'btn_reset_codes': '🔄 Сбросить коды',
        'btn_restore_backup': '🔄 Восстановить из архива (Restore)',
        'btn_revoke_admin_perm': '❌ Отозвать права администратора',
        'btn_risk_radar': '⚠️ Радар успеваемости и рисков',
        'btn_save_att': '💾 Сохранить',
        'btn_school_admins': '👨‍💼 Администрация',
        'btn_search_again': '🔍 Искать снова',
        'btn_search_student': '🔍 Поиск ученика',
        'btn_search_teacher': '🔍 Поиск учителей',
        'btn_search_user': '🔍 Поиск пользователя',
        'btn_send_dm': '✉️ Отправить личное сообщение',
        'btn_send_new_hw': 'Отправить задание',
        'btn_share_contact': '📱 Поделиться номером телефона',
        'btn_student_behavior_history': '⭐ История баллов',
        'btn_submit_hw': '📤 Сдать задание',
        'btn_switch_student': '🧑‍🎓 Сменить ученика',
        'btn_teachers': '👨‍🏫 Учителя',
        'btn_teachers_pdf': '👨‍🏫 Карточки учителей (PDF)',
        'btn_temp_ban_user': '⏱️ Временный бан',
        'btn_timezone_setting': '🕒 Часовой пояс (UTC+{offset})',
        'btn_toggle_readonly': '🔒 Режим только чтение ({status})',
        'btn_transfer_class': '🔄 Передать класс',
        'btn_unack_notifs': '⚠️ Непрочитанные пропуски',
        'btn_unban_user': '🟢 Разблокировать (Разбан)',
        'btn_unlink_parent': '👨‍👩‍👧‍👦 Отвязать родителей',
        'btn_upload_excel': '📥 Загрузить учеников (Excel)',
        'btn_upload_medical': '🏥 Отправить справку',
        'btn_upload_teacher_excel': '👨‍🏫 Загрузить учителей (Excel)',
        'btn_users_hub': '👥 Пользователи',
        'btn_users_list': '⬅️ Список администраторов',
        'btn_view_cafeteria': '🍲 Меню столовой',
        'btn_view_photo': 'Открыть фото',
        'btn_view_schedule': '📅 Расписание уроков',
        'btn_view_submissions': '📥 Сданные работы',
        'btn_weekend_attendance': '📅 Перекличка в выходные ({status})',
        'btn_write_telegram': '💬 Написать в Telegram',
        'btn_write_to_admin': '💬 Написать администратору',
        'cat_reports_title': '📊 <b>Академический контроль и отчеты</b>\nВыберите раздел:',
        'cat_requests_title': '🛎️ <b>Заявки и медицинские справки</b>\nВыберите действие:',
        'cat_settings_title': '⚙️ <b>Системные настройки и безопасность</b>\nВыберите действие:',
        'cat_staff_title': '👥 <b>Управление учениками и учителями</b>\nВыберите действие:',
        'cat_tools_title': '🛠️ <b>Инструменты управления и рассылки</b>\nВыберите раздел:',
        'chat_req_error_toast': '⚠️ Не удалось связаться: пользователь заблокировал бота!',
        'chat_req_sent_toast': '✅ Запрос на контакт отправлен!',
        'child_added_success': '✅ <b>{name}</b> ({class_name}) успешно добавлен к вашему аккаунту!',
        'class_co_teacher_done': '✅ К классу <b>{class_name}</b> также прикреплен учитель <b>{teacher}</b>!',
        'class_transfer_done': '✅ Класс <b>{class_name}</b> успешно передан учителю <b>{teacher}</b>!',
        'cockpit_report': '📊 <b>Утренняя сводка ({date})</b>\n\n🏫 Всего: <b>{total}</b> | ✅ Есть: <b>{present}</b> | ❌ Нет: <b>{absent}</b>\n\n⚠️ <b>Классы без переклички ({missing_cnt}):</b>\n{missing}',
        'codes_reset_done': '✅ Коды обновлены!\n\n• Новый код ученика: <code>{st_code}</code>\n• Новый код родителя: <code>{pr_code}</code>',
        'confirm_delete_student_prompt': '⚠️ <b>ВНИМАНИЕ:</b> Ученик <b>{name}</b> будет удален из системы со всеми оценками и пропусками. Вы уверены?',
        'confirm_delete_teacher_prompt': '⚠️ <b>ВНИМАНИЕ:</b> Учитель <b>{name}</b> будет удален из системы. Вы уверены?',
        'confirm_reset_codes_prompt': '⚠️ <b>ВНИМАНИЕ:</b> Коды доступа будут сброшены, привязанные аккаунты потеряют доступ. Вы уверены?',
        'contact_req_direct': 'Пожалуйста, начните диалог по кнопке ниже:',
        'contact_req_header': '📞 <b>ВЫЗОВ НА СВЯЗЬ ОТ АДМИНИСТРАЦИИ</b>\n\nАдминистрация школы приглашает вас к личному общению (1:1).\n👤 <b>Администратор:</b> {name}\n',
        'contact_req_id': 'Пожалуйста, напишите администрации школы.',
        'dm_delivery_error': '⚠️ Ошибка доставки: Пользователь заблокировал бота.',
        'dm_from_admin_header': '📩 <b>СООБЩЕНИЕ ОТ АДМИНИСТРАЦИИ</b>',
        'dm_sender_label': 'Отправитель',
        'dm_sent_success': '✅ Сообщение успешно доставлено!',
        'duplicate_student_no_error': '⚠️ ОШИБКА: В классе {class_name} уже зарегистрирован ученик с номером {no}!',
        'emergency_alert_prompt': '🚨 <b>ЭКСТРЕННОЕ ОПОВЕЩЕНИЕ</b>\n\nЭто сообщение будет отправлено всем родителям со звуковым уведомлением и кнопкой подтверждения.\n\nВведите текст экстренного сообщения:',
        'emergency_monitor_title': '🚨 <b>Родители, не подтвердившие экстренное сообщение:</b>',
        'err_invalid_birth_date_strict': '⚠️ ОШИБКА: Неверная дата! Пожалуйста, укажите реальную дату в формате ДД.ММ.ГГГГ (напр: 15.05.2008).',
        'err_invalid_details_strict': '⚠️ ОШИБКА: Это поле обязательно. Укажите ваш класс, предмет или пояснение.',
        'err_invalid_gender_strict': '⚠️ <b>Пожалуйста, выберите пол:</b>\n━━━━━━━━━━━━━━━━━━━━\nНажмите одну из кнопок ниже:\n• 👨 <b>Мужской</b>\n• 👩 <b>Женский</b>',
        'err_invalid_name_strict': '⚠️ ОШИБКА: Пожалуйста, введите имя и фамилию через пробел (напр: Иван Иванов).',
        'err_invalid_phone_strict': '⚠️ ОШИБКА: Введите корректный номер телефона или нажмите <b>📱 Поделиться номером телефона</b>.',
        'evening_briefing_header': '🌙 <b>ИТОГИ ДНЯ (18:30)</b>\nУченик: <b>{name}</b> ({class_name})\n\n📌 Посещаемость: <b>{att_status}</b>\n📝 Оценки:\n{grades}',
        'exam_oral': '🗣️ Устный опрос / Активность',
        'exam_schedule_title': '📅 <b>Расписание экзаменов класса {class_name}:</b>',
        'exam_written_1': '📝 1-я Контрольная',
        'exam_written_2': '📝 2-я Контрольная',
        'excel_done': '✅ Обработано! Добавлено учеников: <b>{count}</b>\nКоды прикреплены в файле.',
        'excel_format_error': '❌ Ошибка при обработке Excel. Проверьте формат.',
        'excel_hub_title': '📥 <b>Центр Excel</b>\nВыберите действие:',
        'excel_info': '📥 <b>Импорт через Excel</b>\n\nОтправьте файл <code>.xlsx</code>.\nЗаголовки: <code>Ad Soyad</code> | <code>Sinif</code> | <code>Numara</code>',
        'export_ready': '📥 <b>Архив данных школы готов ({date})</b>',
        'file_size_exceeded_error': '⚠️ ОШИБКА: Файл слишком большой! Максимальный размер файла — 10 МБ.',
        'file_type_not_allowed_error': '⚠️ ОШИБКА: Этот тип файла запрещен. Разрешены только <code>.pdf</code>, <code>.xlsx</code>, <code>.jpg</code>, <code>.png</code>.',
        'grade_deleted': 'Оценка удалена.',
        'grade_parent_notification': '📝 <b>НОВАЯ ОЦЕНКА</b>\n\n🧑‍🎓 Ученик: <b>{name}</b>\n📚 Предмет: <b>{subject}</b> ({exam_type})\n📊 Оценка: <b>{score}</b> ({badge})',
        'grade_saved_success': '✅ Оценка отправлена родителю.',
        'grade_select_class': '📝 Выберите класс для выставления оценок:',
        'grade_select_student': '📝 <b>Класс {class_name}</b>\nВыберите ученика для оценки:',
        'grade_updated': '✅ Оценка обновлена.',
        'homework_board_title': '📢 <b>Доска заданий класса {class_name}:</b>',
        'homework_deleted_toast': 'Задание удалено.',
        'hw_feedback_sent_user': 'ℹ️ <b>Результат проверки задания:</b>\n📚 Предмет: <b>{subject}</b>\n📌 Статус: <b>{status}</b>\n📝 Комментарий учителя: _{feedback}_',
        'hw_sent_success': '📢 Задание отправлено классу <b>{class_name}</b>.',
        'hw_submission_received': '✅ Ваше домашнее задание отправлено учителю.',
        'image_load_error': 'Не удалось загрузить фото.',
        'invalid_admin_pin': '❌ Неверный ПИН-код! Действие отменено в целях безопасности.',
        'invalid_name_error': '❌ Пожалуйста, введите корректное имя.',
        'invalid_parent_code': '❌ Неверный код родителя!',
        'invalid_phone_error': '❌ Некорректный номер! Введите не менее 7 цифр.',
        'invalid_score_format': '❌ Некорректный балл! Введите число (напр: 85).',
        'invalid_score_range': '❌ Оценка должна быть от 0 до 100!',
        'lang_changed': 'Язык успешно изменен: 🇷🇺 Русский',
        'lang_select': '🌍 Пожалуйста, выберите язык:',
        'lbl_account_status': 'Состояние аккаунта',
        'lbl_admin_status': 'Статус администратора',
        'lbl_age': 'лет',
        'lbl_assigned_classes': 'Назначенные классы',
        'lbl_birth_date': 'Дата рождения',
        'lbl_class': 'Класс',
        'lbl_class_teachers': 'Учителя предмета',
        'lbl_full_name': 'ФИО',
        'lbl_gender': 'Пол',
        'lbl_lang': 'Язык',
        'lbl_linked_students': 'Привязанные ученики',
        'lbl_not_admin': 'Обычный пользователь (Не админ)',
        'lbl_number': 'Номер',
        'lbl_phone': 'Телефон',
        'lbl_role': 'Роль',
        'lbl_role_admin': 'Администратор',
        'lbl_role_guest': 'Гость',
        'lbl_role_parent': 'Родитель',
        'lbl_role_student': 'Ученик',
        'lbl_role_teacher': 'Учитель',
        'lbl_status_active': '🟢 <b>Активен</b>',
        'lbl_status_banned': '🚫 <b>Заблокирован (Бан)</b>',
        'lbl_subject': 'Предмет',
        'lbl_today_highlight': '⭐ СЕГОДНЯ',
        'lbl_username': 'Имя пользователя',
        'legal_absence_alert': '⚠️ <b>ПРЕДУПРЕЖДЕНИЕ О ПРОПУСКАХ</b>\n\nВаш ребенок <b>{name}</b> достиг(ла) отметки <b>{count} дней</b> пропусков. Пожалуйста, обратитесь к администрации школы.',
        'lock_countdown_msg': '⛔ <b>Блокировка безопасности:</b> Аккаунт временно заблокирован.\n\nОсталось: <b>{mins} мин.</b>',
        'logout_success_msg': '🚪 Вы успешно вышли из системы. Вы можете войти по новому коду или запросить пароль:',
        'logs_cleaned_toast': 'Удалено {count} старых записей журнала.',
        'maintenance_mode': '⚠️ В системе ведутся технические работы. Пожалуйста, попробуйте позже.',
        'maintenance_mode_updated': 'Режим обслуживания обновлен.',
        'med_uploaded_success': 'Справка отправлена администрации.',
        'medical_approved': '✅ Справка одобрена.',
        'medical_approved_parent': '✅ Справка вашего ребенка одобрена администрацией.',
        'medical_rejected': '❌ Справка отклонена.',
        'medical_rejected_parent': '❌ Справка вашего ребенка отклонена администрацией.',
        'menu_parent': '👨‍👩‍👧‍👦 <b>Панель родителя</b>\nУченик: <b>{name}</b> ({class_name})',
        'menu_student': '🎓 <b>Панель ученика</b>\nУченик: <b>{name}</b> ({class_name} - №: {no})',
        'menu_teacher': '👨‍🏫 <b>Панель учителя</b>\nУчитель: <b>{name}</b> ({subject})',
        'menu_updated': '✅ Меню столовой обновлено.',
        'no_active_homeworks': '📢 Для класса <b>{class_name}</b> нет активных заданий.',
        'no_assigned_classes_teacher': 'ℹ️ За вами пока не закреплены классы. Пожалуйста, обратитесь к администрации.',
        'no_behavior_records': 'Записей о поведении пока нет.',
        'no_blacklisted': '✅ Нет заблокированных пользователей.',
        'no_classes_found': '⚠️ Пока не зарегистрировано ни одного класса.',
        'no_exams_found': '📅 Запланированных экзаменов пока нет.',
        'no_grades': 'Оценки пока не выставлены.',
        'no_hws_found': 'Вы еще не отправляли заданий.',
        'no_linked_student': '⚠️ Ученик не найден в системе.',
        'no_pending_appointments': '✅ Нет запросов на встречу.',
        'no_pending_medical': '✅ Нет справок, ожидающих проверки.',
        'no_pending_requests': '✅ Нет ожидающих заявок на доступ.',
        'no_permission_grade': 'Нет прав для редактирования этой оценки.',
        'no_permission_student_record': 'Нет доступа к записи этого ученика.',
        'no_registered_students': 'Нет привязанных учеников.',
        'no_registered_teachers': '⚠️ Учителя не найдены.',
        'no_students_in_class': 'В этом классе нет учеников.',
        'parent_choose_child': '🧑‍🎓 Выберите ученика:',
        'parent_info_title': 'ℹ️ <b>Информационная панель школы</b>',
        'parent_settings_title': '⚙️ <b>Настройки и аккаунт</b>',
        'parent_unlinked_success': '✅ Родители отвязаны. Новый код родителя: <code>{code}</code>',
        'pdf_ready': '📄 Карточки для <b>{class_name}</b> готовы.',
        'pdf_report_ready': '📄 Официальный табель успеваемости ученика <b>{name}</b> прикреплен.',
        'pending_appointments_title': '🤝 <b>Запросы родителей на встречу:</b>',
        'pending_medical_title': '🏥 <b>Справки на рассмотрении:</b>',
        'pending_requests_title': '🛎️ <b>Заявки на рассмотрении:</b>',
        'perm_admin_assigned_toast': '✅ Пользователю присвоены права постоянного администратора.',
        'permanent_admin_protected': '⛔ Права главного/постоянного администратора не могут быть отозваны!',
        'permanent_admin_title': 'Постоянный администратор',
        'photo_expected_medical': '⚠️ Пожалуйста, отправьте только фотографию.',
        'pin_changed_success': '✅ <b>ПИН-код администратора успешно обновлен!</b>',
        'pin_current_wrong': '❌ Неверный текущий ПИН-код!',
        'pin_mismatch_error': '❌ Введенные ПИН-коды не совпадают!',
        'promotion_confirm_prompt': '🎓 <b>ПЕРЕВОД В СЛЕДУЮЩИЙ КЛАСС</b>\n\nВсе классы будут переведены на ступень выше (напр: <code>9-A</code> ➔ <code>10-A</code>, <code>12-A</code> ➔ <code>Выпускники</code>).\n\nВы подтверждаете действие?',
        'promotion_success': '✅ Перевод завершен! Обновлено учеников: <b>{count}</b>',
        'prompt_add_child_code': '🔑 Введите код родителя второго ребенка (Например: <code>VELI-123456</code>):',
        'prompt_admin_pin': '🔐 <b>ПИН-КОД БЕЗОПАСНОСТИ АДМИНИСТРАТОРА</b>\n\nЭто действие требует подтверждения. Введите 4-значный ПИН-код администратора:',
        'prompt_appointment_note': '📝 Укажите удобное время встречи и примечание:',
        'prompt_behavior_note': '📝 Ученик: <b>{name}</b>\nКатегория: {badge} <b>{title}</b>\n\nВведите примечание (или \'-\' чтобы пропустить):',
        'prompt_broadcast': '📢 Введите текст объявления:',
        'prompt_broadcast_content': '📢 <b>Целевая аудитория:</b> {target}\n\nВведите текст объявления (или отправьте фото с описанием):',
        'prompt_edit_tch_classes': '🏫 Укажите классы учителя через запятую (напр: <code>9-A, 9-B, 10-A</code> или <code>ВСЕ</code>):',
        'prompt_enter_code_direct': '🔑 <b>Введите ваш код доступа:</b> (Например: <code>HCA-123456</code>, <code>VELI-123456</code>, <code>OGR-123456</code>)',
        'prompt_grade_badge': 'Выберите категорию оценки:',
        'prompt_grade_score': 'Ученик: <b>{name}</b> ({class_name})\nВид оценки: <b>{exam_type}</b>\n\nВведите оценку (0-100):',
        'prompt_hw_class': 'Выберите класс для задания:',
        'prompt_hw_content': 'Отправьте текст задания или фото доски:',
        'prompt_hw_submission': '📤 <b>Сдача домашнего задания ({subject})</b>\n\nОтправьте фото выполненного задания или текст:',
        'prompt_menu_update': '🍲 Введите сегодняшнее меню столовой:',
        'prompt_new_score': 'Введите новую оценку (0-100):',
        'prompt_pin_confirm': '🔁 <b>Повторите новый ПИН-код (Подтверждение):</b>',
        'prompt_pin_current': '🔐 <b>Введите текущий ПИН-код:</b>',
        'prompt_pin_new': '🆕 <b>Введите новый 4-значный ПИН-код:</b>',
        'prompt_req_birth_date': '🎂 Введите дату рождения в формате ДД.ММ.ГГГГ (напр: 15.05.2008):',
        'prompt_req_gender': '🚻 Пожалуйста, выберите ваш пол:',
        'prompt_restore_backup': '🔄 Отправьте файл резервной копии школы <code>.xlsx</code>:',
        'prompt_search_student': '🔍 Введите фамилию или номер ученика:',
        'prompt_select_exam_type': '📝 Ученик: <b>{name}</b> ({class_name})\n\nВыберите вид оценивания:',
        'prompt_student_class': '🏫 Введите класс ученика (Например: <code>9-A</code>):',
        'prompt_student_name': '👤 Введите Фамилию и Имя ученика:',
        'prompt_student_no': '🔢 Введите номер ученика (Например: <code>101</code>):',
        'prompt_teacher_name': '👨‍🏫 Введите ФИО учителя:',
        'prompt_teacher_subject': '📚 Введите предмет (Например: <code>Математика</code>):',
        'prompt_upload_teacher_excel': '👨‍🏫 <b>Импорт учителей через Excel</b>\n\nКолонки: <code>Ad Soyad</code> | <code>Brans</code> | <code>(Siniflar)</code>\nОтправьте файл <code>.xlsx</code>:',
        'published_homeworks_title': '📢 <b>Опубликованные задания:</b>',
        'rate_limit_warning': '⚠️ Слишком много запросов. Пожалуйста, подождите несколько секунд.',
        'readonly_mode_active_alert': '🔒 Система в режиме только чтение. Изменения запрещены.',
        'readonly_mode_updated': 'Режим только для чтения обновлен.',
        'recent_grades_title': '📝 <b>Ваши последние оценки:</b>',
        'remind_att_sent': 'Напоминание отправлено.',
        'report_not_found': 'Справка не найдена.',
        'req_already_pending': '⚠️ У вас уже есть заявка на рассмотрении.',
        'req_approved_admin_msg': '✅ Заявка #{id} одобрена. Пользователь: <b>{name}</b> ({role})',
        'req_approved_user': '🎉 <b>Поздравляем!</b>\nАдминистрация одобрила вашу заявку. Вы вошли как <b>{role}</b>.',
        'req_details_parent': '🧑‍🎓 Укажите ФИО, класс или номер вашего ребенка:',
        'req_details_student': '🏫 Укажите ваш класс и номер в школе:',
        'req_details_teacher': '📚 Укажите ваш предмет и примечание для администрации:',
        'req_name_prompt': '👤 Пожалуйста, введите ваши Фамилию и Имя:',
        'req_phone_prompt': '📱 <b>Введите ваш номер телефона:</b>\n<i>(Или нажмите кнопку «Поделиться контактом» ниже)</i>',
        'req_rejected_admin_msg': '❌ Заявка #{id} отклонена: <b>{name}</b>',
        'req_rejected_user': '❌ Администрация отклонила вашу заявку. Обратитесь к администрации.',
        'req_role_select': '🛎️ <b>Запрос Доступа</b>\n━━━━━━━━━━━━━━━━━━━━\nВыберите роль для запроса:',
        'req_sent_success': '✅ Ваша заявка отправлена администрации школы. Вы получите уведомление после проверки.',
        'request_already_handled': '⚠️ Эта заявка уже обработана.',
        'restore_success': '✅ Данные успешно восстановлены! Учеников: <b>{s_cnt}</b>, Учителей: <b>{t_cnt}</b>',
        'rk_appointments': '🤝 Записи родителей',
        'rk_attendance': '📋 Перекличка',
        'rk_behavior': '⭐ Поведение',
        'rk_cancel_action': '❌ Отменить действие',
        'rk_cat_reports': '📊 Отчеты и контроль',
        'rk_cat_requests': '🔔 Центр одобрений',
        'rk_cat_settings': '⚙️ Настройки',
        'rk_cat_staff': '👥 Персонал и дети',
        'rk_cat_tools': '🛠️ Инструменты',
        'rk_grade': '📝 Выставить оценки',
        'rk_homework': '📢 Доска заданий',
        'rk_logout': '🚪 Выйти',
        'rk_parent_info': 'ℹ️ Инфо о школе',
        'rk_report': '📊 Табель оценок',
        'rk_restart': '🔄 Перезапуск',
        'rk_switch_student': '🧑‍🎓 Смена ученика',
        'rk_upload_medical': '🏥 Медсправка',
        'role_parent_btn': '👨‍👩‍👧‍👦 Родитель',
        'role_student_btn': '🎓 Ученик',
        'role_teacher_btn': '👨‍🏫 Учитель',
        'schedule_select_class': '📅 Выберите класс для расписания:',
        'school_admin_title': 'Администрация школы',
        'search_no_results': '❌ Ученик не найден.',
        'search_results_title': '🔍 <b>Результаты поиска:</b>',
        'search_user_no_results': '❌ Пользователь не найден.',
        'search_user_prompt': '🔍 Введите имя, юзернейм или Telegram ID пользователя:',
        'search_user_results_title': '🔍 <b>Результаты поиска пользователей:</b>',
        'select_class_to_co_teacher': '➕ Выберите класс для добавления коллеги:',
        'select_class_to_transfer': '🔄 Выберите класс, который хотите передать:',
        'select_pdf_class': '📄 Выберите класс для карточек:',
        'select_target_teacher': '👨‍🏫 Кому передать класс <b>{class_name}</b>? Выберите учителя:',
        'select_teacher_appointment': '🤝 Выберите учителя для встречи:',
        'send_dm_prompt': '✉️ <b>Отправить личное сообщение:</b>\n\nВведите сообщение для пользователя <code>{name}</code> (<code>{id}</code>):',
        'setting_updated_toast': 'Настройка обновлена.',
        'student_added_card': '✅ <b>Ученик успешно добавлен!</b>\n\n👤 ФИО: <b>{name}</b>\n🏫 Класс: <b>{class_name}</b> | №: <b>{no}</b>\n\n🔑 <b>Коды доступа:</b>\n• Код ученика: <code>{st_code}</code>\n• Код родителя: <code>{pr_code}</code>',
        'student_card': '👤 <b>Карточка ученика</b>\nФИО: <b>{name}</b>\nКласс: <b>{class_name}</b> | №: <b>{no}</b>\n\n🔑 <b>Статус кодов:</b>\n• Ученик: <code>{st_code}</code> ({st_status})\n• Родитель: <code>{pr_code}</code> ({pr_status})',
        'student_deleted': '🗑️ Ученик удален из системы.',
        'student_info_updated': '✅ Данные ученика обновлены:\n<b>{name}</b> ({class_name} - №: {no})',
        'student_name_invalid': '❌ Введите корректное имя.',
        'student_not_found': 'Ученик не найден.',
        'student_switched_success': 'Выбран ученик: <b>{name}</b> ({class_name})',
        'tch_classes_updated': '✅ Назначенные классы учителя обновлены: <b>{classes}</b>',
        'teacher_added_card': '✅ <b>Учитель зарегистрирован!</b>\n\n👤 ФИО: <b>{name}</b>\n📚 Предмет: <b>{subject}</b>\n\n🔑 <b>Код доступа:</b>\n<code>{code}</code>',
        'teacher_card': '👨‍🏫 <b>Карточка учителя</b>\nФИО: <b>{name}</b>\nПредмет: <b>{subject}</b>\n\n🔑 Код: <code>{code}</code>\nСтатус: {status}',
        'teacher_deleted': '🗑️ Учитель удален.',
        'teacher_excel_done': '✅ Добавлено <b>{count}</b> учителей! Коды прикреплены в файле.',
        'teacher_name_invalid': '❌ Введите корректное имя учителя.',
        'teacher_not_found': 'Учитель не найден.',
        'teacher_search_no_results': '❌ Учитель не найден.',
        'teacher_search_prompt': '🔍 Введите ФИО учителя или предмет:',
        'teacher_search_results_title': '🔍 <b>Результаты поиска учителей:</b>',
        'temp_admin_assigned_toast': '✅ Пользователю присвоены права временного администратора на {dur}.',
        'temp_admin_choose_title': '⏱️ <b>Выберите срок временных прав администратора:</b>',
        'temp_ban_choose_title': '⏱️ <b>Выберите срок блокировки:</b>',
        'timezone_updated': 'Часовой пояс установлен на UTC+{offset}.',
        'uc_card_title': '👤 <b>КАРТОЧКА ПОЛЬЗОВАТЕЛЯ И ДЕЙСТВИЯ</b>',
        'unauthorized_action': '⛔ У вас нет прав для этого действия.',
        'unauthorized_excel_upload': '⛔ У вас нет прав загрузки Excel.',
        'unban_success': 'Пользователь разблокирован.',
        'upload_med_prompt': 'Пожалуйста, отправьте фото справки:',
        'user_banned_toast': 'Пользователь заблокирован.',
        'user_not_found_toast': 'Пользователь не найден.',
        'user_temp_banned_notification': '⛔ Ваш аккаунт временно заблокирован на {dur} по соображениям безопасности.',
        'user_temp_banned_toast': 'Пользователь заблокирован на {dur}.',
        'user_unbanned_toast': 'Пользователь разблокирован.',
        'weekend_attendance_updated': 'Настройка переклички в выходные обновлена.',
        'welcome_guest': '🎓 <b>Добро пожаловать</b>\n━━━━━━━━━━━━━━━━━━━━\nВведите ваш <b>код доступа</b> (Напр: <code>VELI-123456</code>, <code>HCA-123456</code>, <code>OGR-123456</code>) или выберите действие:',
    },
    'tr': {
        'rk_digital_locker': '📁 Dijital Dosyam',
        'btn_lost_found': '📦 Kayıp Eşya Masası',
        'btn_quick_excuse': '⚡ 1-Tıkla İzin Bildir',
        'rk_my_credentials': '🔑 Şifrelerim',
        'rk_grade_behavior': '📝 Not & Davranış',
        'rk_student_full_report': '📊 Öğrenci Karnesi',
        'btn_remind_voters': '📢 Oy Kullanmayanlara Hatırlat',
        'btn_scan_bot_blocks': '🔍 Sessiz Sağlık Taraması (Health-Check)',
        'btn_rename_class': '✏️ Sınıf Adını Değiştir',
        'btn_graduates_archive': '🎓 Mezunlar Arşivi',
        'btn_bulletin_program': '📅 Bülten & Program',
        'rk_cat_tools_reports': '📊 Rapor & Araçlar',
        'btn_force_audio_alert': '🚨 Acil Alarm',
        'btn_bot_block_monitor': '🚫 Blok Radarı',
        'btn_proposals': '🗳️ Oylamalar',
        'acknowledged_toast': 'Bildirim onaylandı.',
        'action_cancelled': '❌ <b>İşlem iptal edildi.</b>',
        'admin_add_name_prompt': '👤 Yeni yöneticinin Adını ve Soyadını yazınız:',
        'admin_add_tg_id_prompt': '➕ Yönetici yapılacak kullanıcının Telegram ID numarasını yazınız:',
        'admin_added_success': '✅ <b>{name}</b> (<code>{id}</code>) kalıcı yönetici olarak başarıyla yetkilendirildi.',
        'admin_admins_hub_title': '👨‍💼 <b>Okul Yönetim Kadrosu ve Yetkililer:</b>\n\nProfilini görüntülemek veya yetkisini yönetmek istediğiniz yöneticiye tıklayınız:',
        'admin_code_generated': '🔑 <b>TEK SEFERLİK YÖNETİCİ KODU ÜRETİLDİ</b>\n\nKod: <code>{code}</code>\n\nBu kodu ilgili kişiye iletiniz. Kişi bota bu kodu yazdığı anda hesabı <b>Yönetici (Admin)</b> olarak yetkilendirilecektir.',
        'admin_demoted_notification': 'ℹ️ Yönetici (Admin) yetkileriniz geri alınmıştır.',
        'admin_demoted_toast': 'Yönetici yetkileri geri alındı.',
        'admin_invalid_tg_id': '❌ Geçersiz Telegram ID! Sadece rakamlardan oluşmalıdır.',
        'admin_promoted_notification': '🎉 <b>Sayın {name},</b>\nOkul yönetim sistemi tarafından size <b>Kalıcı Yönetici (Admin)</b> yetkisi tanımlandı!',
        'admin_restart_confirmed': '🔄 <b>Yönetici Paneli Yeniden Başlatıldı.</b>',
        'admin_sched_edit_title': '📅 <b>Ders Programı Düzenle</b>\nSınıf seçiniz:',
        'admin_sched_updated': '✅ <b>{class_name}</b> sınıfının ders programı güncellendi.',
        'admin_stats': '📊 <b>Genel Durum:</b>\n• Sınıf: <b>{c_cnt}</b> | Öğrenci: <b>{s_cnt}</b> | Öğretmen: <b>{t_cnt}</b>\n• Bekleyen Başvuru: <b>{req_cnt}</b> | Rapor: <b>{med_cnt}</b>\n• Tarih: <b>{date}</b>',
        'admin_title': '⚡ <b>Okul Yönetim Masası (Admin)</b>',
        'admin_unban_notification': '🟢 Hesabınızın sistem engeli yönetim tarafından kaldırılmıştır.',
        'admin_user_card_title': '👤 <b>YÖNETİCİ KULLANICI KARTI</b>',
        'all_notifs_acknowledged': '✅ <b>Tüm Devamsızlık Bildirimleri Veliler Tarafından Okundu!</b>\n\nSon 36 saatte velisi tarafından okunmamış hiçbir devamsızlık bildirimi bulunmuyor.',
        'appointment_approved_msg': '✅ Öğretmen randevu talebinizi onayladı.',
        'appointment_confirmed_toast': 'Randevu onaylandı.',
        'appointment_not_found': '⚠️ Randevu bulunamadı.',
        'appointment_rejected_msg': '❌ Öğretmen belirtilen saatte müsait değil.',
        'appointment_sent': '✅ Randevu talebiniz öğretmene iletildi.',
        'att_check_all_done': '✅ Tüm sınıfların yoklaması tamamlandı.',
        'att_check_title': '📊 <b>Günün Yoklama Denetimi ({date})</b>',
        'att_saved': '✅ Yoklama kaydedildi (15 dk düzenlenebilir).',
        'attendance_correction_notification': 'ℹ️ <b>DÜZELTME:</b> <b>{name}</b> isimli öğrencinin devamsızlık kaydı düzeltilmiştir (VAR).',
        'attendance_hours_lock': '⚠️ Yoklama sadece 07:00 - 19:00 saatleri arasında alınabilir.',
        'attendance_intro': '📋 <b>{class_name} Yoklaması</b>\nOlmayan öğrencilere tıklayıp kaydedin:',
        'attendance_select_class': '📋 Yoklama almak istediğiniz sınıfı seçin:',
        'attendance_weekend_lock': '⚠️ Hafta sonu yoklama alınamaz.',
        'auth_blacklisted': '🚫 Hesabınız kalıcı olarak engellenmiştir.',
        'auth_code_already_linked': '⚠️ <b>Bu kod zaten başka bir Telegram hesabına bağlanmış.</b>\n\nLütfen okul idaresi ile iletişime geçiniz.',
        'auth_failed': '❌ Geçersiz kod! Kalan deneme hakkınız: {remaining}',
        'auth_locked': '⛔ Güvenlik kısıtlaması: Hesabınız 1 saat süreyle kilitlendi.',
        'auth_success': '✅ <b>Giriş Başarılı!</b>\nHoş geldiniz: <b>{name}</b>\nRolünüz: <b>{role}</b>',
        'badge_missing': '🟡 Eksik / Geliştirilmeli',
        'badge_praise': '🟢 Tebrik / Başarı',
        'badge_warning': '🔴 Uyarı / Disiplin',
        'bc_home': '🏠 Ana Menü',
        'bc_select_class': '🏫 Duyuru yapılacak sınıfı seçiniz:',
        'bc_select_lang': '🌐 Duyuru yapılacak dili seçiniz:',
        'bc_target_all': '👥 Tüm Okul (Genel)',
        'bc_target_class': '🏫 Sınıfa Göre',
        'bc_target_lang': '🌐 Dile Göre',
        'bc_target_parents': '👨‍👩‍👧‍👦 Sadece Veliler',
        'bc_target_students': '🎓 Sadece Öğrenciler',
        'bc_target_teachers': '👨‍🏫 Sadece Öğretmenler',
        'behavior_parent_notification': '⭐ <b>ÖĞRENCİ DAVRANIŞ BİLGİLENDİRMESİ</b>\n\n🧑‍🎓 Öğrenci: <b>{name}</b> ({class_name})\n🏷️ Değerlendirme: {badge} <b>{title}</b>\n📝 Açıklama: {note}\n👤 Öğretmen: <b>{teacher}</b>',
        'behavior_saved_success': '✅ Davranış notu kaydedildi ve veliye bildirildi.',
        'blacklisted_title': '🚫 <b>Sistemde Engellenen ve Kilitlenen Kullanıcılar:</b>',
        'broadcast_hub_title': '📢 <b>Hedefli Duyuru Masası</b>\nLütfen duyurunun ulaştırılacağı hedef kitleyi seçiniz:',
        'broadcast_sent_report': '📢 Duyuru başarıyla <b>{count}</b> kullanıcıya ulaştırıldı.',
        'broadcast_success': '📢 Duyuru <b>{count}</b> kişiye ulaştırıldı.',
        'btn_academic_report': '📈 Akademik Sıralama',
        'btn_acknowledged': '✅ Okudum / Onaylıyorum',
        'btn_add_admin_id': '➕ Telegram ID ile Yönetici Ekle',
        'btn_add_another': '➕ Yenisini Ekle',
        'btn_add_co_teacher': '➕ Ek Öğretmen Ekle',
        'btn_add_exam': '➕ Sınav Tarihi Ekle',
        'btn_add_negative_badge': '🔴 Uyarı / Kural İhlali',
        'btn_add_new_teacher': '➕ Yeni Öğretmen Ekle',
        'btn_add_positive_badge': '🟢 Tebrik / Başarı',
        'btn_add_student': '➕ Öğrenci Ekle',
        'btn_add_teacher': '➕ Öğretmen Ekle',
        'btn_appr_appointment': '✅ Onayla',
        'btn_appr_medical': '✅ Raporu Onayla',
        'btn_appr_request': '✅ Onayla',
        'btn_assign_all_classes': '🌐 Tüm Sınıfları Ata',
        'btn_attendance': '📋 Hızlı Yoklama',
        'btn_audit_logs': '📜 Güvenlik & İşlem Günlüğü',
        'btn_back': '⬅️ Geri',
        'btn_ban_user': '🚫 Kullanıcıyı Engelle (Ban)',
        'btn_behavior': '⭐ Davranış & Rozet',
        'btn_blacklist': '🚫 Kara Liste',
        'btn_briefing_off': '🔕 Akşam Özeti (KAPALI)',
        'btn_briefing_on': '🔔 Akşam Özeti (AÇIK)',
        'btn_broadcast': '📢 Hedefli Duyuru',
        'btn_cafeteria_edit': '🍲 Yemek Menüsü',
        'btn_cancel_action': '⬅️ İptal Et',
        'btn_change_admin_pin': '🔐 İdari PIN Değiştir',
        'btn_class_att_sheet': 'Yoklama Çizelgesi',
        'btn_class_grade_sheet': 'Not Çizelgesi',
        'btn_class_pdf_cards': 'Şifre Kartları (PDF)',
        'btn_class_promotion': '🎓 Sınıf Terfi',
        'btn_class_sched': 'Ders Programı',
        'btn_classes': '🏫 Sınıflar',
        'btn_clean_logs': '🧹 Log Temizliği',
        'btn_clear_all_classes': '🗑️ Tümünü Kaldır',
        'btn_cockpit': '📊 Sabah Durumu',
        'btn_cockpit_unified': '📊 Sabah Kokpiti & Yoklama',
        'btn_confirm_delete': '✅ Evet, Sil',
        'btn_confirm_reset': '✅ Evet, Sıfırla',
        'btn_del_grade': '❌ Notu Sil',
        'btn_del_student': '❌ Öğrenciyi Sil',
        'btn_del_teacher': '❌ Öğretmeni Sil',
        'btn_delete_action': '🗑️ Sil',
        'btn_download_pdf_report': '📄 Resmi Karne İndir (PDF)',
        'btn_dur_1h': '⏱️ 1 Saat',
        'btn_dur_24h': '⏱️ 24 Saat',
        'btn_dur_30d': '⏱️ 30 Gün',
        'btn_dur_7d': '⏱️ 7 Gün',
        'btn_edit_class': '🏫 Sınıfı Değiştir',
        'btn_edit_grade': '✏️ Notu Düzenle',
        'btn_edit_name': '👤 İsmi Değiştir',
        'btn_edit_no': '🔢 Numarayı Değiştir',
        'btn_edit_student': '✏️ Düzenle',
        'btn_edit_tch_classes': '🏫 Sınıf Yetkileri',
        'btn_emergency_ack': '✅ Durumu Onaylıyorum / Güvendeyiz',
        'btn_emergency_alert': '🚨 Acil Durum / Kırmızı Alarm',
        'btn_emergency_monitor': '🚨 Acil Durum Denetim Masası',
        'btn_enter_grade': '📝 Not Girişi',
        'btn_exam_schedule': '📅 Sınav Takvimi',
        'btn_excel': '📥 Excel ile İçe Aktar',
        'btn_excel_hub': '📥 Excel Masası',
        'btn_export_all_data': '📊 Dışa Aktar',
        'btn_gen_admin_code': '🔑 Tek Seferlik Yönetici Kodu Üret',
        'btn_gender_female': '👩 Kız',
        'btn_gender_male': '👨 Erkek',
        'btn_homework_board': '📢 Ödev Panosu',
        'btn_hw_approve': '✅ Ödevi Onayla',
        'btn_hw_revision': '🔄 Düzeltme İste',
        'btn_lang': '🌐 Dil Seçimi',
        'btn_login_prompt': '🔑 Giriş Yap',
        'btn_main_menu': '🏠 Ana Menü',
        'btn_maintenance_toggle': '🚨 Bakım Modu ({status})',
        'btn_make_perm_admin': '👑 Kalıcı Yönetici Yap',
        'btn_make_temp_admin': '⏱️ Geçici Yönetici Yap',
        'btn_manage_schedule': '📅 Ders Programı',
        'btn_manage_tch_classes': '🏫 Sınıf Yönetimi',
        'btn_medical': '🏥 Raporlar ({count})',
        'btn_my_hws': '📚 Ödevlerim',
        'btn_next': 'İleri ➡️',
        'btn_not_available': '❌ Müsait Değilim',
        'btn_notices': '📢 Okul Duyuruları',
        'btn_pdf': '📄 Şifre Kartları',
        'btn_prev': '⬅️ Geri',
        'btn_quick_recent': '🕒 Son İşlemler',
        'btn_recent_grades_menu': '🕒 Son Notlar & Düzeltme',
        'btn_refresh_data': '🔄 Verileri Yenile',
        'btn_reject': '❌ Reddet',
        'btn_remind_att': '⚠️ Yoklama Almayanlara Bildir',
        'btn_report_card': '📊 Not Karnesi',
        'btn_req_access': '📩 Şifre Talep Et',
        'btn_req_chat': '📞 Özel İletişim (1:1) Talep Et',
        'btn_requests': '🛎️ Başvurular ({count})',
        'btn_reset_codes': '🔄 Kodları Sıfırla',
        'btn_restore_backup': '🔄 Yedekten Geri Yükle (Restore)',
        'btn_revoke_admin_perm': '❌ Yönetici Yetkisini Al',
        'btn_risk_radar': '⚠️ Devamsızlık & Risk Radarı',
        'btn_save_att': '💾 Yoklamayı Kaydet',
        'btn_school_admins': '👨‍💼 İdareciler',
        'btn_search_again': '🔍 Yeniden Ara',
        'btn_search_student': '🔍 Öğrenci Ara',
        'btn_search_teacher': '🔍 Öğretmen Ara',
        'btn_search_user': '🔍 Kullanıcı Ara',
        'btn_send_dm': '✉️ Özel Mesaj Gönder',
        'btn_send_new_hw': 'Ödev Gönder',
        'btn_share_contact': '📱 Telefon Numaramı Paylaş',
        'btn_student_behavior_history': '⭐ Davranış Geçmişi',
        'btn_submit_hw': '📤 Ödev Teslim Et',
        'btn_switch_student': '🧑‍🎓 Öğrenci Değiştir',
        'btn_teachers': '👨‍🏫 Öğretmenler',
        'btn_teachers_pdf': '👨‍🏫 Öğretmen Şifre Kartları (PDF)',
        'btn_temp_ban_user': '⏱️ Süreli Engelle (Mute/Ban)',
        'btn_timezone_setting': '🕒 Saat Dilimi (UTC+{offset})',
        'btn_toggle_readonly': '🔒 Sadece Okunabilir Karantina ({status})',
        'btn_transfer_class': '🔄 Sınıfı Başka Öğretmene Devret',
        'btn_unack_notifs': '⚠️ Okunmamış Bildirimler',
        'btn_unban_user': '🟢 Engeli Kaldır (Unban)',
        'btn_unlink_parent': '👨‍👩‍👧‍👦 Velileri Kopar',
        'btn_upload_excel': '📥 Öğrenci Yükle (Excel)',
        'btn_upload_medical': '🏥 Rapor Gönder',
        'btn_upload_teacher_excel': '👨‍🏫 Toplu Öğretmen Yükle (Excel)',
        'btn_users_hub': '👥 Kullanıcılar',
        'btn_users_list': '⬅️ Yönetici Listesi',
        'btn_view_cafeteria': '🍲 Yemekhane Menüsü',
        'btn_view_photo': 'Fotoğrafı Gör',
        'btn_view_schedule': '📅 Haftalık Ders Programı',
        'btn_view_submissions': '📥 Teslimleri İncele',
        'btn_weekend_attendance': '📅 Hafta Sonu Yoklama ({status})',
        'btn_write_telegram': '💬 Telegram Üzerinden Yaz',
        'btn_write_to_admin': '💬 İdareye Mesaj Gönder',
        'cat_reports_title': '📊 <b>Akademik Denetim ve Rapor Masası</b>\nLütfen incelemek istediğiniz raporu seçiniz:',
        'cat_requests_title': '🛎️ <b>Başvurular ve Rapor Onay Masası</b>\nLütfen işlem yapmak istediğiniz alanı seçiniz:',
        'cat_settings_title': '⚙️ <b>Sistem & Güvenlik Ayarları</b>\nLütfen yapılandırmak istediğiniz ayarı seçiniz:',
        'cat_staff_title': '👥 <b>Kadro & Öğrenci Yönetim Masası</b>\nLütfen işlem yapmak istediğiniz alanı seçiniz:',
        'cat_tools_title': '🛠️ <b>Yönetim Araçları ve Duyuru Masası</b>\nLütfen kullanmak istediğiniz aracı seçiniz:',
        'chat_req_error_toast': '⚠️ İletişim kurulamadı: Kullanıcı botu engellemiş!',
        'chat_req_sent_toast': '✅ 1:1 İletişim talebi iletildi!',
        'child_added_success': '✅ <b>{name}</b> ({class_name}) başarıyla hesabınıza eklendi!',
        'class_co_teacher_done': '✅ <b>{class_name}</b> sınıfına <b>{teacher}</b> da ortak öğretmen olarak atandı!',
        'class_transfer_done': '✅ <b>{class_name}</b> sınıfı başarıyla <b>{teacher}</b> öğretmenine devredildi!',
        'cockpit_report': '📊 <b>Sabah Özeti ({date})</b>\n\n🏫 Toplam: <b>{total}</b> | ✅ Var: <b>{present}</b> | ❌ Yok: <b>{absent}</b>\n\n⚠️ <b>Yoklama Alınmayan Sınıflar ({missing_cnt}):</b>\n{missing}',
        'codes_reset_done': '✅ Kodlar sıfırlandı!\n\n• Yeni Öğrenci Kodu: <code>{st_code}</code>\n• Yeni Veli Kodu: <code>{pr_code}</code>',
        'confirm_delete_student_prompt': '⚠️ <b>DİKKAT:</b> <b>{name}</b> isimli öğrenci tüm not ve devamsızlık kayıtlarıyla birlikte silinecektir. Onaylıyor musunuz?',
        'confirm_delete_teacher_prompt': '⚠️ <b>DİKKAT:</b> <b>{name}</b> isimli öğretmen silinecektir. Onaylıyor musunuz?',
        'confirm_reset_codes_prompt': '⚠️ <b>DİKKAT:</b> Erişim kodları sıfırlanacak ve bağlı hesapların erişimi kesilecektir. Onaylıyor musunuz?',
        'contact_req_direct': 'Lütfen aşağıdaki butonu kullanarak doğrudan görüşme başlatınız:',
        'contact_req_header': '📞 <b>OKUL İDARESİ İLETİŞİM TALEBİ</b>\n\nOkul idaresi sizinle 1:1 özel görüşme talep etmektedir.\n👤 <b>Görüşme Talep Eden İdareci:</b> {name}\n',
        'contact_req_id': 'Lütfen okul idaresine yazınız.',
        'dm_delivery_error': '⚠️ İletim Hatası: Kullanıcı botu engellemiş.',
        'dm_from_admin_header': '📩 <b>OKUL İDARESİNDEN ÖZEL BİLDİRİM</b>',
        'dm_sender_label': 'Gönderen Yetkili',
        'dm_sent_success': '✅ Özel mesaj kullanıcıya başarıyla iletildi!',
        'duplicate_student_no_error': '⚠️ HATA: {class_name} sınıfında {no} numaralı öğrenci zaten kayıtlı!',
        'emergency_alert_prompt': '🚨 <b>ACİL DURUM / KIRMIZI ALARM DUYURUSU</b>\n\nBu bildirim tüm velilere sesli yüksek öncelikle iletilecek ve ekranlarında onay butonu çıkacaktır.\n\nLütfen acil durum mesajını yazınız:',
        'emergency_monitor_title': '🚨 <b>Acil Durumu Henüz Onaylamayan Veliler:</b>',
        'err_invalid_birth_date_strict': '⚠️ HATA: Geçersiz tarih! Lütfen GG.AA.YYYY formatında gerçek bir tarih giriniz (Örn: 15.05.2008).',
        'err_invalid_details_strict': '⚠️ HATA: Bu alan boş bırakılamaz. Lütfen sınıfınızı, branşınızı veya açıklamanızı yazınız.',
        'err_invalid_gender_strict': '⚠️ <b>Lütfen cinsiyetinizi seçiniz:</b>\n━━━━━━━━━━━━━━━━━━━━\nAşağıdaki butonlardan birine dokununuz:\n• 👨 <b>Erkek</b>\n• 👩 <b>Kız</b>',
        'err_invalid_name_strict': '⚠️ HATA: Lütfen aralarında boşluk olacak şekilde Adınızı ve Soyadınızı tam giriniz (Örn: Ahmet Yılmaz).',
        'err_invalid_phone_strict': '⚠️ HATA: Lütfen geçerli bir telefon numarası giriniz veya <b>📱 Telefon Numaramı Paylaş</b> butonuna basınız.',
        'evening_briefing_header': '🌙 <b>GÜN SONU ÖZETİ (18:30)</b>\nÖğrenci: <b>{name}</b> ({class_name})\n\n📌 Devamsızlık: <b>{att_status}</b>\n📝 Notlar:\n{grades}',
        'exam_oral': '🗣️ Sözlü / Performans',
        'exam_schedule_title': '📅 <b>{class_name} Sınıfı Sınav Takvimi:</b>',
        'exam_written_1': '📝 1. Yazılı Sınav',
        'exam_written_2': '📝 2. Yazılı Sınav',
        'excel_done': '✅ İşlendi! Eklenen öğrenci: <b>{count}</b>\nŞifreler ektedir.',
        'excel_format_error': '❌ Excel işlenirken hata oluştu. Formatı kontrol edin.',
        'excel_hub_title': '📥 <b>Excel Yönetim Masası</b>\nYapmak istediğiniz işlemi seçiniz:',
        'excel_info': '📥 <b>Excel ile Yükleme</b>\n\n<code>.xlsx</code> dosyası gönderin.\nBaşlıklar: <code>Ad Soyad</code> | <code>Sinif</code> | <code>Numara</code>',
        'export_ready': '📥 <b>Okul Veri Yedeği Hazır ({date})</b>',
        'file_size_exceeded_error': '⚠️ HATA: Gönderilen dosya boyutu çok yüksek! Maksimum dosya boyutu 10 MB olabilir.',
        'file_type_not_allowed_error': '⚠️ HATA: Bu dosya türü kabul edilmemektedir. Sadece <code>.pdf</code>, <code>.xlsx</code>, <code>.jpg</code>, <code>.png</code> gönderebilirsiniz.',
        'grade_deleted': 'Not silindi.',
        'grade_parent_notification': '📝 <b>YENİ NOT GİRİŞİ</b>\n\n🧑‍🎓 Öğrenci: <b>{name}</b>\n📚 Ders: <b>{subject}</b> ({exam_type})\n📊 Not: <b>{score}</b> ({badge})',
        'grade_saved_success': '✅ Not veliye iletildi.',
        'grade_select_class': '📝 Not girmek istediğiniz sınıfı seçin:',
        'grade_select_student': '📝 <b>{class_name} Sınıfı</b>\nNot verilecek öğrenciyi seçin:',
        'grade_updated': '✅ Not güncellendi.',
        'homework_board_title': '📢 <b>{class_name} Ödev Panosu:</b>',
        'homework_deleted_toast': 'Ödev silindi.',
        'hw_feedback_sent_user': 'ℹ️ <b>Ödev Değerlendirme Sonucu:</b>\n📚 Ders: <b>{subject}</b>\n📌 Durum: <b>{status}</b>\n📝 Öğretmen Notu: _{feedback}_',
        'hw_sent_success': '📢 Ödev <b>{class_name}</b> sınıfına iletildi.',
        'hw_submission_received': '✅ Ödeviniz başarıyla öğretmene iletildi.',
        'image_load_error': 'Görsel yüklenemedi.',
        'invalid_admin_pin': '❌ Geçersiz Yönetici PIN Kodu! Güvenlik nedeniyle işlem iptal edildi.',
        'invalid_name_error': '❌ Lütfen geçerli bir isim giriniz.',
        'invalid_parent_code': '❌ Geçersiz veli kodu!',
        'invalid_phone_error': '❌ Geçersiz telefon numarası! En az 7 haneli rakam giriniz.',
        'invalid_score_format': '❌ Geçersiz not! Lütfen sayısal bir değer giriniz (Örn: 85).',
        'invalid_score_range': '❌ Not 0 ile 100 arasında olmalıdır!',
        'lang_changed': 'Dil başarıyla güncellendi: 🇹🇷 Türkçe',
        'lang_select': '🌍 Lütfen bir dil seçiniz:',
        'lbl_account_status': 'Hesap Durumu',
        'lbl_admin_status': 'Yönetici Durumu',
        'lbl_age': 'yaşında',
        'lbl_assigned_classes': 'Atanmış Sınıflar',
        'lbl_birth_date': 'Doğum Tarihi',
        'lbl_class': 'Sınıf',
        'lbl_class_teachers': 'Ders Öğretmenleri',
        'lbl_full_name': 'Adı Soyadı',
        'lbl_gender': 'Cinsiyet',
        'lbl_lang': 'Sistem Dili',
        'lbl_linked_students': 'Bağlı Öğrenciler',
        'lbl_not_admin': 'Standart Kullanıcı (Admin Değil)',
        'lbl_number': 'Numara',
        'lbl_phone': 'Telefon',
        'lbl_role': 'Kullanıcı Rolü',
        'lbl_role_admin': 'Yönetici',
        'lbl_role_guest': 'Misafir',
        'lbl_role_parent': 'Veli',
        'lbl_role_student': 'Öğrenci',
        'lbl_role_teacher': 'Öğretmen',
        'lbl_status_active': '🟢 <b>Aktif</b>',
        'lbl_status_banned': '🚫 <b>Engellendi (Ban)</b>',
        'lbl_subject': 'Branş / Ders',
        'lbl_today_highlight': '⭐ BUGÜN',
        'lbl_username': 'Kullanıcı Adı',
        'legal_absence_alert': '⚠️ <b>YASAL DEVAMSIZLIK UYARISI</b>\n\nÖğrenciniz <b>{name}</b> toplam <b>{count} gün</b> devamsızlığa ulaşmıştır. Lütfen okul idaresi ile iletişime geçiniz.',
        'lock_countdown_msg': '⛔ <b>Güvenlik Karantinası:</b> Hesabınız geçici olarak kilitlenmiştir.\n\nKalan süre: <b>{mins} dakika</b>.',
        'logout_success_msg': '🚪 Başarıyla çıkış yapıldı. Yeni bir kod girebilir veya başvuru yapabilirsiniz:',
        'logs_cleaned_toast': '{count} eski log kaydı temizlendi.',
        'maintenance_mode': '⚠️ Sistem şu anda bakım modundadır. Lütfen daha sonra tekrar deneyiniz.',
        'maintenance_mode_updated': 'Bakım modu ayarı güncellendi.',
        'med_uploaded_success': 'Rapor okul idaresine iletildi.',
        'medical_approved': '✅ Sağlık raporu onaylandı.',
        'medical_approved_parent': '✅ Öğrencinizin sağlık raporu okul idaresi tarafından onaylandı.',
        'medical_rejected': '❌ Sağlık raporu reddedildi.',
        'medical_rejected_parent': '❌ Öğrencinizin sağlık raporu okul idaresi tarafından reddedildi.',
        'menu_parent': '👨‍👩‍👧‍👦 <b>Veli Masası</b>\nÖğrenci: <b>{name}</b> ({class_name})',
        'menu_student': '🎓 <b>Öğrenci Masası</b>\nÖğrenci: <b>{name}</b> ({class_name} - No: {no})',
        'menu_teacher': '👨‍🏫 <b>Öğretmen Masası</b>\nÖğretmen: <b>{name}</b> ({subject})',
        'menu_updated': '✅ Yemekhane menüsü güncellendi.',
        'no_active_homeworks': '📢 <b>{class_name}</b> sınıfı için aktif ödev bulunmuyor.',
        'no_assigned_classes_teacher': 'ℹ️ Hesabınıza tanımlanmış aktif bir sınıf bulunmuyor. Lütfen okul idaresi ile iletişime geçiniz.',
        'no_behavior_records': 'Henüz davranış kaydı bulunmuyor.',
        'no_blacklisted': '✅ Engellenen kullanıcı bulunmuyor.',
        'no_classes_found': '⚠️ Henüz kayıtlı sınıf bulunmuyor.',
        'no_exams_found': '📅 Yaklaşan sınav bulunmuyor.',
        'no_grades': 'Henüz not girilmemiş.',
        'no_hws_found': 'Henüz yayınlanan ödev bulunmuyor.',
        'no_linked_student': '⚠️ Hesabınıza bağlı öğrenci bulunamadı.',
        'no_pending_appointments': '✅ Bekleyen randevu talebi yok.',
        'no_pending_medical': '✅ Bekleyen sağlık raporu yok.',
        'no_pending_requests': '✅ Bekleyen başvuru yok.',
        'no_permission_grade': 'Bu notu düzenleme yetkiniz yok.',
        'no_permission_student_record': 'Bu öğrencinin kaydına erişim yetkiniz yok.',
        'no_registered_students': 'Bağlı öğrenci yok.',
        'no_registered_teachers': '⚠️ Kayıtlı öğretmen bulunamadı.',
        'no_students_in_class': 'Bu sınıfta kayıtlı öğrenci yok.',
        'parent_choose_child': '🧑‍🎓 Lütfen işlem yapmak istediğiniz öğrenciyi seçiniz:',
        'parent_info_title': 'ℹ️ <b>Okul Bilgi & İletişim Panosu</b>',
        'parent_settings_title': '⚙️ <b>Ayarlar & Hesap Masası</b>',
        'parent_unlinked_success': '✅ Veliler koparıldı. Yeni Veli Kodu: <code>{code}</code>',
        'pdf_ready': '📄 <b>{class_name}</b> şifre kartları hazır.',
        'pdf_report_ready': '📄 <b>{name}</b> isimli öğrencinin resmi karne belgesi ektedir.',
        'pending_appointments_title': '🤝 <b>Bekleyen Veli Görüşme Talepleri:</b>',
        'pending_medical_title': '🏥 <b>Onay Bekleyen Sağlık Raporları:</b>',
        'pending_requests_title': '🛎️ <b>Onay Bekleyen Yetki Başvuruları:</b>',
        'perm_admin_assigned_toast': '✅ Kullanıcı kalıcı yönetici olarak yetkilendirildi.',
        'permanent_admin_protected': '⛔ Kurucu/Kalıcı yöneticilerin yetkisi alınamaz!',
        'permanent_admin_title': 'Kalıcı Yönetici',
        'photo_expected_medical': '⚠️ Lütfen sadece fotoğraf gönderiniz.',
        'pin_changed_success': '✅ <b>İdari PIN Kodu Başarıyla Değiştirildi!</b>',
        'pin_current_wrong': '❌ Mevcut PIN kodu hatalı!',
        'pin_mismatch_error': '❌ Girilen yeni PIN kodları eşleşmiyor!',
        'promotion_confirm_prompt': '🎓 <b>YIL SONU SINIF TERFİ İŞLEMİ</b>\n\nTüm sınıflar bir üst kademeye aktarılacaktır (Örn: <code>9-A</code> ➔ <code>10-A</code>, <code>12-A</code> ➔ <code>Mezun</code>).\n\nİşlemi onaylıyor musunuz?',
        'promotion_success': '✅ Sınıf terfi işlemi başarıyla tamamlandı! Güncellenen: <b>{count}</b>',
        'prompt_add_child_code': '🔑 Lütfen diğer çocuğunuzun veli kodunu giriniz (Örn: <code>VELI-123456</code>):',
        'prompt_admin_pin': '🔐 <b>İDARİ GÜVENLİK PİN KALKANI</b>\n\nBu kritik işlem üst düzey güvenlik onayı gerektirir. Lütfen 4 haneli İdari PIN kodunu tuşlayınız:',
        'prompt_appointment_note': '📝 Randevu için uygun olduğunuz zamanı ve notunuzu yazınız:',
        'prompt_behavior_note': '📝 Öğrenci: <b>{name}</b>\nRozet: {badge} <b>{title}</b>\n\nVarsa açıklama notu yazınız (veya \'-\' yazıp geçiniz):',
        'prompt_broadcast': '📢 Duyuru metnini yazın:',
        'prompt_broadcast_content': '📢 <b>Hedef Kitle:</b> {target}\n\nLütfen duyuru metnini yazınız (veya fotoğrafla birlikte açıklama gönderiniz):',
        'prompt_edit_tch_classes': '🏫 Öğretmenin sınıflarını virgülle yazınız (Örn: <code>9-A, 9-B, 10-A</code> veya tümü için <code>TUMU</code>):',
        'prompt_enter_code_direct': '🔑 <b>Lütfen erişim kodunuzu giriniz:</b> (Örn: <code>HCA-123456</code>, <code>VELI-123456</code>, <code>OGR-123456</code>)',
        'prompt_grade_badge': 'Performans değerlendirmesini seçiniz:',
        'prompt_grade_score': 'Öğrenci: <b>{name}</b> ({class_name})\nDeğerlendirme: <b>{exam_type}</b>\n\nNotu giriniz (0-100):',
        'prompt_hw_class': 'Ödev gönderilecek sınıfı seçin:',
        'prompt_hw_content': 'Ödev açıklamasını yazın veya tahta fotoğrafı gönderin:',
        'prompt_hw_submission': '📤 <b>Ödev Teslimi ({subject})</b>\n\nLütfen ödevinizin fotoğrafını gönderiniz veya açıklama yazınız:',
        'prompt_menu_update': '🍲 Bugünün yemekhane menüsünü yazınız:',
        'prompt_new_score': 'Yeni notu giriniz (0-100):',
        'prompt_pin_confirm': '🔁 <b>Yeni PIN Kodunu Tekrar Tuşlayınız (Onay):</b>',
        'prompt_pin_current': '🔐 <b>Mevcut İdari PIN Kodunu Tuşlayınız:</b>',
        'prompt_pin_new': '🆕 <b>Yeni 4 Haneli PIN Kodunu Tuşlayınız:</b>',
        'prompt_req_birth_date': '🎂 Doğum tarihinizi GG.AA.YYYY formatında yazınız (Örn: 15.05.2008):',
        'prompt_req_gender': '🚻 Lütfen cinsiyetinizi seçiniz:',
        'prompt_restore_backup': '🔄 Lütfen okul yedekleme <code>.xlsx</code> dosyasını gönderiniz:',
        'prompt_search_student': '🔍 Öğrenci adı veya numarası yazınız:',
        'prompt_select_exam_type': '📝 Öğrenci: <b>{name}</b> ({class_name})\n\nLütfen not türünü seçiniz:',
        'prompt_student_class': '🏫 Öğrencinin Sınıfını giriniz (Örn: <code>9-A</code>):',
        'prompt_student_name': '👤 Öğrencinin Adını ve Soyadını giriniz:',
        'prompt_student_no': '🔢 Öğrencinin Okul Numarasını giriniz (Örn: <code>101</code>):',
        'prompt_teacher_name': '👨‍🏫 Öğretmenin Adını ve Soyadını giriniz:',
        'prompt_teacher_subject': '📚 Öğretmenin Branşını giriniz (Örn: <code>Matematik</code>):',
        'prompt_upload_teacher_excel': '👨‍🏫 <b>Öğretmen Excel Yükleme</b>\n\nSütunlar: <code>Ad Soyad</code> | <code>Brans</code> | <code>(Siniflar)</code>\nDosyayı <code>.xlsx</code> formatında gönderiniz:',
        'published_homeworks_title': '📢 <b>Yayınlanan Ödevler:</b>',
        'rate_limit_warning': '⚠️ Çok fazla istek gönderdiniz. Lütfen birkaç saniye bekleyin.',
        'readonly_mode_active_alert': '🔒 Sistem şu anda Sadece Okunabilir (Read-Only) modundadır. Veri değişikliği yapılamaz.',
        'readonly_mode_updated': 'Sadece okunabilir mod ayarı güncellendi.',
        'recent_grades_title': '📝 <b>Son Girdiğiniz Notlar:</b>',
        'remind_att_sent': 'Hatırlatma iletildi.',
        'report_not_found': 'Rapor bulunamadı.',
        'req_already_pending': '⚠️ Zaten onay bekleyen bir başvurunuz bulunmaktadır.',
        'req_approved_admin_msg': '✅ Başvuru #{id} onaylandı. Kullanıcı: <b>{name}</b> ({role})',
        'req_approved_user': '🎉 <b>Tebrikler!</b>\nOkul idaresi başvurunuzu onayladı. <b>{role}</b> olarak giriş yaptınız.',
        'req_details_parent': '🧑‍🎓 Çocuğunuzun Adı, Sınıfı veya Numarasını yazınız:',
        'req_details_student': '🏫 Okul Numaranızı ve Sınıfınızı yazınız:',
        'req_details_teacher': '📚 Branşınızı ve idareye iletmek istediğiniz notu yazınız:',
        'req_name_prompt': '👤 Lütfen Adınızı ve Soyadınızı tam olarak yazınız:',
        'req_phone_prompt': '📱 <b>Lütfen telefon numaranızı giriniz:</b>\n<i>(Veya aşağıdaki «Telefon Numaramı Paylaş» butonuna basınız)</i>',
        'req_rejected_admin_msg': '❌ Başvuru #{id} reddedildi: <b>{name}</b>',
        'req_rejected_user': '❌ Başvurunuz idare tarafından onaylanmadı. Okul idaresi ile iletişime geçiniz.',
        'req_role_select': '🛎️ <b>Erişim ve Şifre Başvurusu</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen başvuru rolünüzü seçiniz:',
        'req_sent_success': '✅ Başvurunuz okul idaresine iletildi. İncelendikten sonra bilgilendirileceksiniz.',
        'request_already_handled': '⚠️ Bu başvuru daha önce işlem görmüş.',
        'restore_success': '✅ Sistem yedeği başarıyla yüklendi! Öğrenci: <b>{s_cnt}</b>, Öğretmen: <b>{t_cnt}</b>',
        'rk_appointments': '🤝 Veli Görüşmeleri',
        'rk_attendance': '📋 Hızlı Yoklama',
        'rk_behavior': '⭐ Davranış & Puan',
        'rk_cancel_action': '❌ İşlemi İptal Et',
        'rk_cat_reports': '📊 Raporlar & Denetim',
        'rk_cat_requests': '🔔 Onay Masası',
        'rk_cat_settings': '⚙️ Sistem Ayarları',
        'rk_cat_staff': '👥 Kadro & Öğrenci',
        'rk_cat_tools': '🛠️ Yönetim Araçları',
        'rk_grade': '📝 Not Girişi',
        'rk_homework': '📢 Ödev Panosu',
        'rk_logout': '🚪 Çıkış Yap',
        'rk_parent_info': 'ℹ️ Okul Bilgisi',
        'rk_report': '📊 Not Karnesi',
        'rk_restart': '🔄 Yeniden Başlat',
        'rk_switch_student': '🧑‍🎓 Öğrenci Seç',
        'rk_upload_medical': '🏥 Sağlık Raporu',
        'role_parent_btn': '👨‍👩‍👧‍👦 Veli',
        'role_student_btn': '🎓 Öğrenci',
        'role_teacher_btn': '👨‍🏫 Öğretmen',
        'schedule_select_class': '📅 Programını görmek istediğiniz sınıfı seçin:',
        'school_admin_title': 'Okul İdaresi',
        'search_no_results': '❌ Öğrenci bulunamadı.',
        'search_results_title': '🔍 <b>Arama Sonuçları:</b>',
        'search_user_no_results': '❌ Eşleşen kullanıcı bulunamadı.',
        'search_user_prompt': '🔍 Kullanıcı adı, ismi veya Telegram ID yazınız:',
        'search_user_results_title': '🔍 <b>Kullanıcı Arama Sonuçları:</b>',
        'select_class_to_co_teacher': '➕ Ortak öğretmen eklemek istediğiniz sınıfı seçiniz:',
        'select_class_to_transfer': '🔄 Devretmek istediğiniz sınıfı seçiniz:',
        'select_pdf_class': '📄 Kartlarını indirmek istediğiniz sınıfı seçin:',
        'select_target_teacher': '👨‍🏫 <b>{class_name}</b> sınıfı kime devredilecek? Hedef öğretmeni seçin:',
        'select_teacher_appointment': '🤝 Randevu almak istediğiniz öğretmeni seçin:',
        'send_dm_prompt': '✉️ <b>Kullanıcıya Özel Mesaj Gönder:</b>\n\n<code>{name}</code> (<code>{id}</code>) kullanıcısına iletilecek mesajı yazınız:',
        'setting_updated_toast': 'Ayar güncellendi.',
        'student_added_card': '✅ <b>Öğrenci Başarıyla Eklendi!</b>\n\n👤 Adı Soyadı: <b>{name}</b>\n🏫 Sınıf: <b>{class_name}</b> | No: <b>{no}</b>\n\n🔑 <b>Giriş Kodları:</b>\n• Öğrenci Kodu: <code>{st_code}</code>\n• Veli Kodu: <code>{pr_code}</code>',
        'student_card': '👤 <b>Öğrenci Kartı</b>\nAdı Soyadı: <b>{name}</b>\nSınıf: <b>{class_name}</b> | No: <b>{no}</b>\n\n🔑 <b>Kod Durumu:</b>\n• Öğrenci: <code>{st_code}</code> ({st_status})\n• Veli: <code>{pr_code}</code> ({pr_status})',
        'student_deleted': '🗑️ Öğrenci sistemden silindi.',
        'student_info_updated': '✅ Öğrenci bilgileri güncellendi:\n<b>{name}</b> ({class_name} - No: {no})',
        'student_name_invalid': '❌ Lütfen geçerli bir isim giriniz.',
        'student_not_found': 'Öğrenci bulunamadı.',
        'student_switched_success': 'Aktif öğrenci: <b>{name}</b> ({class_name})',
        'tch_classes_updated': '✅ Öğretmenin sınıfları güncellendi: <b>{classes}</b>',
        'teacher_added_card': '✅ <b>Öğretmen Kaydedildi!</b>\n\n👤 Adı Soyadı: <b>{name}</b>\n📚 Branşı: <b>{subject}</b>\n\n🔑 <b>Giriş Kodu:</b>\n<code>{code}</code>',
        'teacher_card': '👨‍🏫 <b>Öğretmen Kartı</b>\nAdı Soyadı: <b>{name}</b>\nBranş: <b>{subject}</b>\n\n🔑 Kod: <code>{code}</code>\nDurum: {status}',
        'teacher_deleted': '🗑️ Öğretmen silindi.',
        'teacher_excel_done': '✅ <b>{count}</b> öğretmen eklendi! Şifreler ektedir.',
        'teacher_name_invalid': '❌ Lütfen geçerli bir öğretmen adı giriniz.',
        'teacher_not_found': 'Öğretmen bulunamadı.',
        'teacher_search_no_results': '❌ Öğretmen bulunamadı.',
        'teacher_search_prompt': '🔍 Öğretmen adı veya branş yazınız:',
        'teacher_search_results_title': '🔍 <b>Öğretmen Arama Sonuçları:</b>',
        'temp_admin_assigned_toast': '✅ Kullanıcıya {dur} süreyle geçici yönetici yetkisi verildi.',
        'temp_admin_choose_title': '⏱️ <b>Geçici Yönetici Süresini Seçiniz:</b>',
        'temp_ban_choose_title': '⏱️ <b>Engelleme (Ban) Süresini Seçiniz:</b>',
        'timezone_updated': 'Saat dilimi UTC+{offset} olarak ayarlandı.',
        'uc_card_title': '👤 <b>KULLANICI BİLGİ VE YETKİ KARTI</b>',
        'unauthorized_action': '⛔ Bu işlem için yetkiniz bulunmuyor.',
        'unauthorized_excel_upload': '⛔ Excel yükleme yetkiniz yok.',
        'unban_success': 'Kullanıcının engeli kaldırıldı.',
        'upload_med_prompt': 'Lütfen sağlık raporunun fotoğrafını gönderin:',
        'user_banned_toast': 'Kullanıcı engellendi.',
        'user_not_found_toast': 'Kullanıcı bulunamadı.',
        'user_temp_banned_notification': '⛔ Hesabınız idari güvenlik gerekçesiyle {dur} süreyle geçici olarak engellenmiştir.',
        'user_temp_banned_toast': 'Kullanıcı {dur} süreyle engellendi.',
        'user_unbanned_toast': 'Kullanıcının engeli kaldırıldı.',
        'weekend_attendance_updated': 'Hafta sonu yoklama ayarı güncellendi.',
        'welcome_guest': '🎓 <b>Hoş Geldiniz</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen size verilen <b>erişim kodunu</b> (Örn: <code>VELI-123456</code>, <code>HCA-123456</code>, <code>OGR-123456</code>) yazınız veya işlem seçiniz:',
    },
    'uz': {
        'rk_digital_locker': '📁 Mening ishim',
        'btn_lost_found': '📦 Topilmalar byurosi',
        'btn_quick_excuse': '⚡ 1-Bosishda ruxsatnoma',
        'rk_my_credentials': '🔑 Parollarim',
        'rk_grade_behavior': '📝 Baho va xulq',
        'rk_student_full_report': '📊 O\'quvchi tabeli',
        'btn_remind_voters': '📢 Ovoz bermaganlarga eslatish',
        'btn_scan_bot_blocks': '🔍 Tinch aloqa tekshiruvi (Health-Check)',
        'btn_rename_class': "✏️ Sinf nomini o'zgartirish",
        'btn_graduates_archive': '🎓 Bitiruvchilar',
        'btn_bulletin_program': '📅 Xabar & Jadval',
        'rk_cat_tools_reports': '📊 Hisobot & Vosita',
        'btn_force_audio_alert': '🚨 Tezkor signal',
        'btn_bot_block_monitor': '🚫 Blok nazorati',
        'btn_proposals': '🗳️ Ovoz berish',
        'acknowledged_toast': 'Tasdiqlandi.',
        'action_cancelled': '❌ <b>Amal bekor qilindi.</b>',
        'admin_add_name_prompt': '👤 Yangi ma\'murning Ism va Familiyasini kiriting:',
        'admin_add_tg_id_prompt': '➕ Ma\'mur qilinadigan foydalanuvchining Telegram ID raqamini kiriting:',
        'admin_added_success': '✅ <b>{name}</b> (<code>{id}</code>) doimiy ma\'mur sifatida muvaffaqiyatli tayinlandi.',
        'admin_admins_hub_title': '👨‍💼 <b>Maktab rahbariyati va vakolatli xodimlar:</b>\n\nProfilini ko\'rish yoki vakolatini boshqarish uchun ma\'murni tanlang:',
        'admin_code_generated': '🔑 <b>BIR MARTALIK MA\'MUR KODI YARATILDI</b>\n\nKod: <code>{code}</code>\n\nUshbu kodni tegishli shaxsga yuboring. U botga ushbu kodni kiritishi bilan hisobi <b>Ma\'mur (Admin)</b> sifatida faollashadi.',
        'admin_demoted_notification': 'ℹ️ Ma\'mur (Admin) vakolatlaringiz bekor qilindi.',
        'admin_demoted_toast': 'Ma\'mur vakolatlari bekor qilindi.',
        'admin_invalid_tg_id': '❌ Noto\'g\'ri Telegram ID! Faqat raqamlardan iborat bo\'lishi kerak.',
        'admin_promoted_notification': '🎉 <b>Hurmatli {name},</b>\nMaktab boshqaruv tizimi tomonidan sizga <b>Doimiy ma\'mur (Admin)</b> vakolati berildi!',
        'admin_restart_confirmed': '🔄 <b>Boshqaruv paneli qayta ishga tushirildi.</b>',
        'admin_sched_edit_title': '📅 <b>Dars jadvalini tahrirlash</b>\nSinfni tanlang:',
        'admin_sched_updated': '✅ <b>{class_name}</b> sinfining dars jadvali yangilandi.',
        'admin_stats': '📊 <b>Umumiy holat:</b>\n• Sinflar: <b>{c_cnt}</b> | O\'quvchilar: <b>{s_cnt}</b> | O\'qituvchilar: <b>{t_cnt}</b>\n• Arizalar: <b>{req_cnt}</b> | Ma\'lumotnomalar: <b>{med_cnt}</b>\n• Sana: <b>{date}</b>',
        'admin_title': '⚡ <b>Maktab boshqaruv markazi (Admin)</b>',
        'admin_unban_notification': '🟢 Hisobingizning tizim blokirovkasi ma\'muriyat tomonidan olib tashlandi.',
        'admin_user_card_title': '👤 <b>MA\'MUR FOYDALANUVCHI KARTASI</b>',
        'all_notifs_acknowledged': '✅ <b>Barcha dars qoldirish xabarnomalari ota-onalar tomonidan o\'qildi!</b>\n\nOxirgi 36 soatda o\'qilmagan xabarnomalar mavjud emas.',
        'appointment_approved_msg': '✅ O\'qituvchi uchrashuv so\'rovingizni tasdiqladi.',
        'appointment_confirmed_toast': 'Uchrashuv tasdiqlandi.',
        'appointment_not_found': '⚠️ Uchrashuv topilmadi.',
        'appointment_rejected_msg': '❌ O\'qituvchi ushbu vaqtda band.',
        'appointment_sent': '✅ Uchrashuv so\'rovi o\'qituvchiga yuborildi.',
        'att_check_all_done': '✅ Barcha sinflarning davomati olindi.',
        'att_check_title': '📊 <b>Kunlik davomat nazorati ({date})</b>',
        'att_saved': '✅ Davomat saqlandi (15 daqiqa tuzatish mumkin).',
        'attendance_correction_notification': 'ℹ️ <b>TUZATISH:</b> <b>{name}</b> ismli o\'quvchining dars qoldirish qaydi to\'g\'rilandi (BOR).',
        'attendance_hours_lock': '⚠️ Davomat faqat 07:00 dan 19:00 gacha olinishi mumkin.',
        'attendance_intro': '📋 <b>{class_name} davomati</b>\nDarsda yo\'qlarni belgilang va saqlang:',
        'attendance_select_class': '📋 Davomat olmoqchi bo\'lgan sinfni tanlang:',
        'attendance_weekend_lock': '⚠️ Dam olish kunlari davomat olinmaydi.',
        'auth_blacklisted': '🚫 Hisobingiz butunlay bloklangan.',
        'auth_code_already_linked': '⚠️ <b>Ushbu kod allaqachon boshqa Telegram hisobiga ulangan.</b>\n\nIltimos, maktab ma\'muriyatiga murojaat qiling.',
        'auth_failed': '❌ Noto\'g\'ri kod! Qolgan urinishlar: {remaining}',
        'auth_locked': '⛔ Xavfsizlik cheklovi: Hisobingiz 1 soatga bloklandi.',
        'auth_success': '✅ <b>Kirish muvaffaqiyatli!</b>\nXush kelibsiz: <b>{name}</b>\nSizning rolingiz: <b>{role}</b>',
        'badge_missing': '🟡 Kamchilik / Ishlash kerak',
        'badge_praise': '🟢 Maqtov / Muvaffaqiyat',
        'badge_warning': '🔴 Ogohlantirish / Intizom',
        'bc_home': '🏠 Asosiy Menyu',
        'bc_select_class': '🏫 E\'lon yuboriladigan sinfni tanlang:',
        'bc_select_lang': '🌐 E\'lon yuboriladigan tilni tanlang:',
        'bc_target_all': '👥 Butun maktab (Umumiy)',
        'bc_target_class': '🏫 Sinflar bo\'yicha',
        'bc_target_lang': '🌐 Tillar bo\'yicha',
        'bc_target_parents': '👨‍👩‍👧‍👦 Faqat ota-onalar',
        'bc_target_students': '🎓 Faqat o\'quvchilar',
        'bc_target_teachers': '👨‍🏫 Faqat o\'qituvchilar',
        'behavior_parent_notification': '⭐ <b>O\'QUVCHI XULQ-ATVORI HAQIDA MA\'LUMOT</b>\n\n🧑‍🎓 O\'quvchi: <b>{name}</b> ({class_name})\n🏷️ Baholash: {badge} <b>{title}</b>\n📝 Izoh: {note}\n👤 O\'qituvchi: <b>{teacher}</b>',
        'behavior_saved_success': '✅ Xulq-atvor qaydi saqlandi va ota-onaga yuborildi.',
        'blacklisted_title': '🚫 <b>Bloklangan va cheklangan foydalanuvchilar:</b>',
        'broadcast_hub_title': '📢 <b>Maqsadli e\'lonlar markazi</b>\nIltimos, e\'lon yuboriladigan auditoriyani tanlang:',
        'broadcast_sent_report': '📢 E\'lon muvaffaqiyatli ravishda <b>{count}</b> ta foydalanuvchiga yetkazildi.',
        'broadcast_success': '📢 E\'lon <b>{count}</b> kishiga yetkazildi.',
        'btn_academic_report': '📈 Akademik reyting',
        'btn_acknowledged': "✅ O\'qidim / Tasdiqlayman",
        'btn_add_admin_id': "➕ Telegram ID orqali ma\'mur qo\'shish",
        'btn_add_another': "➕ Yana qo\'shish",
        'btn_add_co_teacher': "➕ Qo\'shimcha o\'qituvchi biriktirish",
        'btn_add_exam': "➕ Imtihon sanasi qo\'shish",
        'btn_add_negative_badge': '🔴 Ogohlantirish / Qoida buzilishi',
        'btn_add_new_teacher': "➕ Yangi o\'qituvchi qo\'shish",
        'btn_add_positive_badge': '🟢 Maqtov / Muvaffaqiyat',
        'btn_add_student': "➕ O\'quvchi qo\'shish",
        'btn_add_teacher': "➕ O\'qituvchi qo\'shish",
        'btn_appr_appointment': '✅ Tasdiqlash',
        'btn_appr_medical': "✅ Ma\'lumotnomani tasdiqlash",
        'btn_appr_request': '✅ Tasdiqlash',
        'btn_assign_all_classes': '🌐 Barcha sinflarni biriktirish',
        'btn_attendance': '📋 Tezkor davomat',
        'btn_audit_logs': '📜 Xavfsizlik jurnali',
        'btn_back': '⬅️ Orqaga',
        'btn_ban_user': '🚫 Bloklash (Ban)',
        'btn_behavior': '⭐ Xulq-atvor va nishonlar',
        'btn_blacklist': "🚫 Qora ro'yxat",
        'btn_briefing_off': "🔕 Kechki hisobot (O\'CHIQ)",
        'btn_briefing_on': '🔔 Kechki hisobot (YONIQ)',
        'btn_broadcast': "📢 E'lon",
        'btn_cafeteria_edit': '🍲 Oshxona menyusi',
        'btn_cancel_action': '⬅️ Bekor qilish',
        'btn_change_admin_pin': "🔐 Ma\'muriy PIN kodni o\'zgartirish",
        'btn_class_att_sheet': 'Davomat varaqasi',
        'btn_class_grade_sheet': 'Baholar qaydnomasi',
        'btn_class_pdf_cards': 'Parol kartalari (PDF)',
        'btn_class_promotion': '🎓 Sinf koʻchirish',
        'btn_class_sched': 'Dars jadvali',
        'btn_classes': '🏫 Sinflar',
        'btn_clean_logs': '🧹 Log tozalash',
        'btn_clear_all_classes': '🗑️ Barchasini tozalash',
        'btn_cockpit': '📊 Ertalabki holat',
        'btn_cockpit_unified': '📊 Ertalabki holat va davomat',
        'btn_confirm_delete': "✅ Ha, o\'chirilsin",
        'btn_confirm_reset': "✅ Ha, qayta o\'rnatilsin",
        'btn_del_grade': "❌ Bahoni o\'chirish",
        'btn_del_student': "❌ O\'quvchini o\'chirish",
        'btn_del_teacher': "❌ O\'qituvchini o\'chirish",
        'btn_delete_action': "🗑️ O\'chirish",
        'btn_download_pdf_report': '📄 Rasmiy baholar tabelini yuklab olish (PDF)',
        'btn_dur_1h': '⏱️ 1 Soat',
        'btn_dur_24h': '⏱️ 24 Soat',
        'btn_dur_30d': '⏱️ 30 Kun',
        'btn_dur_7d': '⏱️ 7 Kun',
        'btn_edit_class': "🏫 Sinfni o\'zgartirish",
        'btn_edit_grade': '✏️ Bahoni tahrirlash',
        'btn_edit_name': "👤 Ismni o\'zgartirish",
        'btn_edit_no': "🔢 Raqamni o\'zgartirish",
        'btn_edit_student': '✏️ Tahrirlash',
        'btn_edit_tch_classes': '🏫 Sinf vakolatlari',
        'btn_emergency_ack': '✅ Holatni tasdiqlayman / Xavfsizdamiz',
        'btn_emergency_alert': '🚨 Favqulodda holat / Qizil signal',
        'btn_emergency_monitor': '🚨 Favqulodda nazorat markazi',
        'btn_enter_grade': "📝 Baho qo\'yish",
        'btn_exam_schedule': '📅 Imtihon jadvali',
        'btn_excel': '📥 Excel orqali yuklash',
        'btn_excel_hub': '📥 Excel markazi',
        'btn_export_all_data': "📊 Eksport (.xlsx)",
        'btn_gen_admin_code': "🔑 Bir martalik ma\'mur kodi yaratish",
        'btn_gender_female': '👩 Qiz bola',
        'btn_gender_male': "👨 O\'g\'il bola",
        'btn_homework_board': '📢 Vazifalar paneli',
        'btn_hw_approve': '✅ Vazifani qabul qilish',
        'btn_hw_revision': '🔄 Qayta ishlashga qaytarish',
        'btn_lang': '🌐 Tilni tanlash',
        'btn_login_prompt': '🔑 Kirish',
        'btn_main_menu': '🏠 Asosiy Menyu',
        'btn_maintenance_toggle': '🚨 Texnik xizmat rejimi ({status})',
        'btn_make_perm_admin': "👑 Doimiy ma\'mur qilish",
        'btn_make_temp_admin': "⏱️ Vaqtinchalik ma\'mur qilish",
        'btn_manage_schedule': '📅 Dars jadvali',
        'btn_manage_tch_classes': '🏫 Sinflarni boshqarish',
        'btn_medical': "🏥 Ma\'lumotnomalar ({count})",
        'btn_my_hws': '📚 Mening vazifalarim',
        'btn_next': 'Oldinga ➡️',
        'btn_not_available': "❌ Vaqtim yo\'q",
        'btn_notices': "📢 Maktab e\'lonlari",
        'btn_pdf': '📄 Parol kartalari',
        'btn_prev': '⬅️ Orqaga',
        'btn_quick_recent': "🕒 So\'nggi amallar",
        'btn_recent_grades_menu': "🕒 So\'nggi baholar",
        'btn_refresh_data': "🔄 Ma\'lumotlarni yangilash",
        'btn_reject': '❌ Rad etish',
        'btn_remind_att': '⚠️ Davomat olmaganlarga eslatish',
        'btn_report_card': '📊 Baholar tabeli',
        'btn_req_access': "📩 Parol so\'rash",
        'btn_req_chat': "📞 Bog\'lanish (1:1) so\'rash",
        'btn_requests': '🛎️ Arizalar ({count})',
        'btn_reset_codes': '🔄 Kodlarni yangilash',
        'btn_restore_backup': '🔄 Arxivdan tiklash (Restore)',
        'btn_revoke_admin_perm': "❌ Ma\'mur vakolatini olish",
        'btn_risk_radar': '⚠️ Davomat va xavf radari',
        'btn_save_att': '💾 Davomatni saqlash',
        'btn_school_admins': '👨‍💼 Rahbariyat',
        'btn_search_again': '🔍 Qayta qidirish',
        'btn_search_student': '🔍 Oʻquvchi izlash',
        'btn_search_teacher': "🔍 O'qituvchi qidirish",
        'btn_search_user': '🔍 Foydalanuvchi qidirish',
        'btn_send_dm': '✉️ Shaxsiy xabar yuborish',
        'btn_send_new_hw': 'Vazifa yuborish',
        'btn_share_contact': '📱 Telefon raqamimni ulashish',
        'btn_student_behavior_history': '⭐ Ballar tarixi',
        'btn_submit_hw': '📤 Vazifani topshirish',
        'btn_switch_student': "🧑‍🎓 O\'quvchini almashtirish",
        'btn_teachers': '👨‍🏫 Ustozlar',
        'btn_teachers_pdf': "👨‍🏫 O\'qituvchi kartalari (PDF)",
        'btn_temp_ban_user': '⏱️ Muddatli bloklash',
        'btn_timezone_setting': '🕒 Vaqt mintaqasi (UTC+{offset})',
        'btn_toggle_readonly': "🔒 Faqat o\'qish rejimi ({status})",
        'btn_transfer_class': "🔄 Sinfni boshqa o\'qituvchiga o\'tkazish",
        'btn_unack_notifs': "⚠️ O\'qilmagan xabarnomalar",
        'btn_unban_user': '🟢 Blokdan chiqarish (Unban)',
        'btn_unlink_parent': '👨‍👩‍👧‍👦 Ota-onalarni uzish',
        'btn_upload_excel': "📥 O\'quvchilarni yuklash (Excel)",
        'btn_upload_medical': "🏥 Ma\'lumotnoma yuborish",
        'btn_upload_teacher_excel': "👨‍🏫 O\'qituvchilarni yuklash (Excel)",
        'btn_users_hub': "👥 Foydalanuvchi'yxati",
        'btn_users_list': "⬅️ Ma\'murlar ro\'yxati",
        'btn_view_cafeteria': '🍲 Oshxona menyusi',
        'btn_view_photo': "Rasmni ko\'rish",
        'btn_view_schedule': '📅 Haftalik dars jadvali',
        'btn_view_submissions': "📥 Topshiriqlarni ko\'rish",
        'btn_weekend_attendance': '📅 Dam olish kunlari davomat ({status})',
        'btn_write_telegram': '💬 Telegram orqali yozish',
        'btn_write_to_admin': '💬 Rahbariyatga xabar yuborish',
        'cat_reports_title': '📊 <b>Akademik nazorat va hisobotlar markazi</b>\nIltimos, kerakli bo\'limni tanlang:',
        'cat_requests_title': '🛎️ <b>Arizalar va ma\'lumotnomalar markazi</b>\nIltimos, kerakli amalni tanlang:',
        'cat_settings_title': '⚙️ <b>Tizim va xavfsizlik sozlamalari</b>\nIltimos, sozlamani tanlang:',
        'cat_staff_title': '👥 <b>Xodimlar va o\'quvchilarni boshqarish</b>\nIltimos, kerakli bo\'limni tanlang:',
        'cat_tools_title': '🛠️ <b>Boshqaruv vositalari va e\'lonlar</b>\nIltimos, kerakli vositani tanlang:',
        'chat_req_error_toast': '⚠️ Bog\'lanib bo\'lmadi: Foydalanuvchi botni bloklagan!',
        'chat_req_sent_toast': '✅ Bog\'lanish so\'rovi yuborildi!',
        'child_added_success': '✅ <b>{name}</b> ({class_name}) hisobingizga muvaffaqiyatli qo\'shildi!',
        'class_co_teacher_done': '✅ <b>{class_name}</b> sinfiga <b>{teacher}</b> ham o\'qituvchi sifatida biriktirildi!',
        'class_transfer_done': '✅ <b>{class_name}</b> sinfi muvaffaqiyatli ravishda <b>{teacher}</b> ga o\'tkazildi!',
        'cockpit_report': '📊 <b>Ertalabki hisobot ({date})</b>\n\n🏫 Jami: <b>{total}</b> | ✅ Bor: <b>{present}</b> | ❌ Yo\'q: <b>{absent}</b>\n\n⚠️ <b>Davomat olinmagan sinflar ({missing_cnt}):</b>\n{missing}',
        'codes_reset_done': '✅ Kodlar yangilandi!\n\n• Yangi O\'quvchi kodi: <code>{st_code}</code>\n• Yangi Ota-ona kodi: <code>{pr_code}</code>',
        'confirm_delete_student_prompt': '⚠️ <b>DIQQAT:</b> <b>{name}</b> ismli o\'quvchi barcha baholari va davomat qaydlari bilan birga o\'chiriladi. Tasdiqlaysizmi?',
        'confirm_delete_teacher_prompt': '⚠️ <b>DIQQAT:</b> <b>{name}</b> ismli o\'qituvchi o\'chiriladi. Tasdiqlaysizmi?',
        'confirm_reset_codes_prompt': '⚠️ <b>DIQQAT:</b> Kirish kodlari yangilanadi va ulangan hisoblar uziladi. Tasdiqlaysizmi?',
        'contact_req_direct': 'Iltimos, quyidagi tugma orqali to\'g\'ridan-to\'g\'ri suhbatni boshlang:',
        'contact_req_header': '📞 <b>MAKTAB RAHBARIYATIDAN BOG\'LANISH SO\'ROVI</b>\n\nMaktab ma\'muriyati siz bilan 1:1 shaxsiy suhbat so\'ramoqda.\n👤 <b>So\'rov yuborgan ma\'mur:</b> {name}\n',
        'contact_req_id': 'Iltimos, maktab rahbariyatiga yozing.',
        'dm_delivery_error': '⚠️ Yetkazish xatosi: Foydalanuvchi botni bloklagan.',
        'dm_from_admin_header': '📩 <b>MAKTAB RAHBARIYATIDAN XABAR</b>',
        'dm_sender_label': 'Yuboruvchi',
        'dm_sent_success': '✅ Xabar foydalanuvchiga muvaffaqiyatli yetkazildi!',
        'duplicate_student_no_error': '⚠️ XATOLIK: {class_name} sinfida {no} raqamli o\'quvchi allaqachon mavjud!',
        'emergency_alert_prompt': '🚨 <b>FAVQULODDA HOLAT / QIZIL SIGNAL</b>\n\nUshbu xabar barcha ota-onalarga ovozli signal bilan yetkaziladi va tasdiqlash tugmasi chiqadi.\n\nFavqulodda xabar matnini kiriting:',
        'emergency_monitor_title': '🚨 <b>Favqulodda xabarni hali tasdiqlamagan ota-onalar:</b>',
        'err_invalid_birth_date_strict': '⚠️ XATOLIK: Noto\'g\'ri sana! Iltimos, KK.OO.YYYY formatida haqiqiy sanani kiriting (Masalan: 15.05.2008).',
        'err_invalid_details_strict': '⚠️ XATOLIK: Bu maydon bo\'sh bo\'lishi mumkin emas. Sinfingiz, faningiz yoki izohingizni yozing.',
        'err_invalid_gender_strict': '⚠️ <b>Iltimos, jinsingizni tanlang:</b>\n━━━━━━━━━━━━━━━━━━━━\nQuyidagi tugmalardan birini bosing:\n• 👦 <b>O\'g\'il bola</b>\n• 👧 <b>Qiz bola</b>',
        'err_invalid_name_strict': '⚠️ XATOLIK: Iltimos, Ism va Familiyangizni oralarida bo\'sh joy qoldirib to\'liq kiriting (Masalan: Ali Valiyev).',
        'err_invalid_phone_strict': '⚠️ XATOLIK: Iltimos, to\'g\'ri telefon raqamini kiriting yoki <b>📱 Telefon raqamimni ulashish</b> tugmasini bosing.',
        'evening_briefing_header': '🌙 <b>KUN YAKUNI HISOBOTI (18:30)</b>\nO\'quvchi: <b>{name}</b> ({class_name})\n\n📌 Davomat: <b>{att_status}</b>\n📝 Baholar:\n{grades}',
        'exam_oral': '🗣️ Og\'zaki / Faollik',
        'exam_schedule_title': '📅 <b>{class_name} sinfi imtihonlar jadvali:</b>',
        'exam_written_1': '📝 1-Yozma imtihon',
        'exam_written_2': '📝 2-Yozma imtihon',
        'excel_done': '✅ Bajarildi! Qo\'shilgan o\'quvchilar: <b>{count}</b>\nParollar biriktirilgan.',
        'excel_format_error': '❌ Excel faylni o\'qishda xatolik yuz berdi. Formatni tekshiring.',
        'excel_hub_title': '📥 <b>Excel boshqaruv markazi</b>\nKerakli amalni tanlang:',
        'excel_info': '📥 <b>Excel orqali yuklash</b>\n\n<code>.xlsx</code> fayl yuboring.\nUstunlar: <code>Ad Soyad</code> | <code>Sinif</code> | <code>Numara</code>',
        'export_ready': '📥 <b>Maktab ma\'lumotlar arxivi tayyor ({date})</b>',
        'file_size_exceeded_error': '⚠️ XATOLIK: Fayl hajmi juda katta! Maksimal fayl hajmi 10 MB bo\'lishi mumkin.',
        'file_type_not_allowed_error': '⚠️ XATOLIK: Ushbu fayl turiga ruxsat berilmagan. Faqat <code>.pdf</code>, <code>.xlsx</code>, <code>.jpg</code>, <code>.png</code> yuborishingiz mumkin.',
        'grade_deleted': 'Baho o\'chirildi.',
        'grade_parent_notification': '📝 <b>YANGI BAHO QO\'YILDI</b>\n\n🧑‍🎓 O\'quvchi: <b>{name}</b>\n📚 Fan: <b>{subject}</b> ({exam_type})\n📊 Baho: <b>{score}</b> ({badge})',
        'grade_saved_success': '✅ Baho ota-onaga yuborildi.',
        'grade_select_class': '📝 Baho qo\'ymoqchi bo\'lgan sinfni tanlang:',
        'grade_select_student': '📝 <b>{class_name} sinfi</b>\nBaholanadigan o\'quvchini tanlang:',
        'grade_updated': '✅ Baho yangilandi.',
        'homework_board_title': '📢 <b>{class_name} vazifalar paneli:</b>',
        'homework_deleted_toast': 'Vazifa o\'chirildi.',
        'hw_feedback_sent_user': 'ℹ️ <b>Vazifani tekshirish natijasi:</b>\n📚 Fan: <b>{subject}</b>\n📌 Holat: <b>{status}</b>\n📝 O\'qituvchi izohi: _{feedback}_',
        'hw_sent_success': '📢 Vazifa <b>{class_name}</b> sinfiga yuborildi.',
        'hw_submission_received': '✅ Vazifangiz o\'qituvchiga muvaffaqiyatli yuborildi.',
        'image_load_error': 'Rasmni yuklab bo\'lmadi.',
        'invalid_admin_pin': '❌ Noto\'g\'ri Ma\'muriy PIN kod! Xavfsizlik sababli amal bekor qilindi.',
        'invalid_name_error': '❌ Iltimos, to\'g\'ri ism kiriting.',
        'invalid_parent_code': '❌ Ota-ona kodi noto\'g\'ri!',
        'invalid_phone_error': '❌ Noto\'g\'ri telefon raqami! Kamida 7 xonali raqam kiriting.',
        'invalid_score_format': '❌ Noto\'g\'ri baho! Raqam kiriting (Masalan: 85).',
        'invalid_score_range': '❌ Baho 0 dan 100 gacha bo\'lishi kerak!',
        'lang_changed': 'Til muvaffaqiyatli yangilandi: 🇺🇿 O\'zbekcha',
        'lang_select': '🌍 Iltimos, tilni tanlang:',
        'lbl_account_status': 'Hisob holati',
        'lbl_admin_status': 'Ma\'mur holati',
        'lbl_age': 'yoshda',
        'lbl_assigned_classes': 'Biriktirilgan sinflar',
        'lbl_birth_date': 'Tug\'ilgan sana',
        'lbl_class': 'Sinf',
        'lbl_class_teachers': 'Fan o\'qituvchilari',
        'lbl_full_name': 'F.I.O',
        'lbl_gender': 'Jinsi',
        'lbl_lang': 'Tizim tili',
        'lbl_linked_students': 'Biriktirilgan o\'quvchilar',
        'lbl_not_admin': 'Oddiy foydalanuvchi (Admin emas)',
        'lbl_number': 'Raqam',
        'lbl_phone': 'Telefon',
        'lbl_role': 'Rol',
        'lbl_role_admin': 'Ma\'mur',
        'lbl_role_guest': 'Mehmon',
        'lbl_role_parent': 'Ota-ona',
        'lbl_role_student': 'O\'quvchi',
        'lbl_role_teacher': 'O\'qituvchi',
        'lbl_status_active': '🟢 <b>Faol</b>',
        'lbl_status_banned': '🚫 <b>Bloklangan (Ban)</b>',
        'lbl_subject': 'Mutaxassislik / Fan',
        'lbl_today_highlight': '⭐ BUGUN',
        'lbl_username': 'Foydalanuvchi nomi',
        'legal_absence_alert': '⚠️ <b>QONUNIY DAVOMAT OGOHLANTIRISHI</b>\n\nO\'quvchingiz <b>{name}</b> jami <b>{count} kun</b> dars qoldirdi. Iltimos, maktab ma\'muriyati bilan bog\'laning.',
        'lock_countdown_msg': '⛔ <b>Xavfsizlik karantini:</b> Hisobingiz vaqtincha bloklangan.\n\nQolgan vaqt: <b>{mins} daqiqa</b>.',
        'logout_success_msg': '🚪 Tizimdan chiqildi. Yangi kod kiritishingiz yoki parol so\'rashingiz mumkin:',
        'logs_cleaned_toast': '{count} ta eski log tozalandi.',
        'maintenance_mode': '⚠️ Tizimda ta\'mirlash ishlari olib borilmoqda. Iltimos, keyinroq urinib ko\'ring.',
        'maintenance_mode_updated': 'Texnik rejim sozlamasi yangilandi.',
        'med_uploaded_success': 'Ma\'lumotnoma ma\'muriyatga yuborildi.',
        'medical_approved': '✅ Ma\'lumotnoma tasdiqlandi.',
        'medical_approved_parent': '✅ O\'quvchingizning ma\'lumotnomasi ma\'muriyat tomonidan tasdiqlandi.',
        'medical_rejected': '❌ Ma\'lumotnoma rad etildi.',
        'medical_rejected_parent': '❌ O\'quvchingizning ma\'lumotnomasi ma\'muriyat tomonidan rad etildi.',
        'menu_parent': '👨‍👩‍👧‍👦 <b>Ota-ona paneli</b>\nO\'quvchi: <b>{name}</b> ({class_name})',
        'menu_student': '🎓 <b>O\'quvchi paneli</b>\nO\'quvchi: <b>{name}</b> ({class_name} - №: {no})',
        'menu_teacher': '👨‍🏫 <b>O\'qituvchi paneli</b>\nO\'qituvchi: <b>{name}</b> ({subject})',
        'menu_updated': '✅ Oshxona menyusi yangilandi.',
        'no_active_homeworks': '📢 <b>{class_name}</b> sinfi uchun faol vazifalar yo\'q.',
        'no_assigned_classes_teacher': 'ℹ️ Hisobingizga biriktirilgan faol sinflar topilmadi. Iltimos, ma\'muriyatga murojaat qiling.',
        'no_behavior_records': 'Hozircha xulq-atvor qaydlari yo\'q.',
        'no_blacklisted': '✅ Bloklangan foydalanuvchilar yo\'q.',
        'no_classes_found': '⚠️ Hozircha sinflar ro\'yxatga olinmagan.',
        'no_exams_found': '📅 Rejalashtirilgan imtihonlar yo\'q.',
        'no_grades': 'Hali baholar qo\'yilmagan.',
        'no_hws_found': 'Hali e\'lon qilingan vazifalar yo\'q.',
        'no_linked_student': '⚠️ Hisobingizga bog\'langan o\'quvchi topilmadi.',
        'no_pending_appointments': '✅ Kutilayotgan uchrashuv so\'rovlari yo\'q.',
        'no_pending_medical': '✅ Kutilayotgan ma\'lumotnomalar yo\'q.',
        'no_pending_requests': '✅ Kutilayotgan arizalar yo\'q.',
        'no_permission_grade': 'Bu bahoni tahrirlashga huquqingiz yo\'q.',
        'no_permission_student_record': 'Ushbu o\'quvchi ma\'lumotlarini ko\'rishga huquqingiz yo\'q.',
        'no_registered_students': 'Bog\'langan o\'quvchilar yo\'q.',
        'no_registered_teachers': '⚠️ Ro\'yxatga olingan o\'qituvchilar topilmadi.',
        'no_students_in_class': 'Bu sinfda o\'quvchilar yo\'q.',
        'parent_choose_child': '🧑‍🎓 Iltimos, o\'quvchini tanlang:',
        'parent_info_title': 'ℹ️ <b>Maktab ma\'lumot paneli</b>',
        'parent_settings_title': '⚙️ <b>Sozlamalar va hisob</b>',
        'parent_unlinked_success': '✅ Ota-onalar uzildi. Yangi ota-ona kodi: <code>{code}</code>',
        'pdf_ready': '📄 <b>{class_name}</b> sinfi parol kartalari tayyor.',
        'pdf_report_ready': '📄 <b>{name}</b> ismli o\'quvchining rasmiy baholar tabeli biriktirildi.',
        'pending_appointments_title': '🤝 <b>Ota-onalar bilan uchrashuv so\'rovlari:</b>',
        'pending_medical_title': '🏥 <b>Ko\'rib chiqilayotgan ma\'lumotnomalar:</b>',
        'pending_requests_title': '🛎️ <b>Ko\'rib chiqilayotgan arizalar:</b>',
        'perm_admin_assigned_toast': '✅ Foydalanuvchiga doimiy ma\'mur vakolati berildi.',
        'permanent_admin_protected': '⛔ Asosiy ma\'murning vakolatini bekor qilib bo\'lmaydi!',
        'permanent_admin_title': 'Doimiy ma\'mur',
        'photo_expected_medical': '⚠️ Iltimos, faqat rasm yuboring.',
        'pin_changed_success': '✅ <b>Ma\'muriy PIN kod muvaffaqiyatli o\'zgartirildi!</b>',
        'pin_current_wrong': '❌ Joriy PIN kod noto\'g\'ri!',
        'pin_mismatch_error': '❌ Kiritilgan yangi PIN kodlar bir-biriga mos kelmadi!',
        'promotion_confirm_prompt': '🎓 <b>YIL YAKUNI SINF KO\'CHIRISH AMALI</b>\n\nBarcha sinflar bir pog\'ona yuqoriga ko\'chiriladi (Masalan: <code>9-A</code> ➔ <code>10-A</code>, <code>12-A</code> ➔ <code>Bitiruvchi</code>).\n\nAmalni tasdiqlaysizmi?',
        'promotion_success': '✅ Sinflar muvaffaqiyatli ko\'chirildi! Yangilanganlar: <b>{count}</b>',
        'prompt_add_child_code': '🔑 Boshqa farzandingizning ota-ona kodini kiriting (Masalan: <code>VELI-123456</code>):',
        'prompt_admin_pin': '🔐 <b>MA\'MURIY XAVFSIZLIK PIN QALQONI</b>\n\nUshbu muhim amal yuqori darajadagi ruxsatni talab qiladi. Iltimos, 4 xonali Ma\'muriy PIN kodni tering:',
        'prompt_appointment_note': '📝 Uchrashuv uchun qulay vaqtingizni va izohingizni yozing:',
        'prompt_behavior_note': '📝 O\'quvchi: <b>{name}</b>\nNishon: {badge} <b>{title}</b>\n\nIzoh yozing (yoki \'-\' deb yuboring):',
        'prompt_broadcast': '📢 E\'lon matnini kiriting:',
        'prompt_broadcast_content': '📢 <b>Maqsadli auditoriya:</b> {target}\n\nE\'lon matnini kiriting (yoki rasm bilan birga izoh yuboring):',
        'prompt_edit_tch_classes': '🏫 O\'qituvchi sinflarini vergul bilan yozing (Masalan: <code>9-A, 9-B, 10-A</code> yoki barchasi uchun <code>BARCHASI</code>):',
        'prompt_enter_code_direct': '🔑 <b>Iltimos, kirish kodini kiriting:</b> (Masalan: <code>HCA-123456</code>, <code>VELI-123456</code>, <code>OGR-123456</code>)',
        'prompt_grade_badge': 'Baho toifasini tanlang:',
        'prompt_grade_score': 'O\'quvchi: <b>{name}</b> ({class_name})\nBaholash turi: <b>{exam_type}</b>\n\nBahoni kiriting (0-100):',
        'prompt_hw_class': 'Vazifa yuboriladigan sinfni tanlang:',
        'prompt_hw_content': 'Vazifa matnini yozing yoki doska rasmini yuboring:',
        'prompt_hw_submission': '📤 <b>Vazifa topshirish ({subject})</b>\n\nVazifa rasmini yuboring yoki izoh yozing:',
        'prompt_menu_update': '🍲 Bugungi oshxona menyusini kiriting:',
        'prompt_new_score': 'Yangi bahoni kiriting (0-100):',
        'prompt_pin_confirm': '🔁 <b>Yangi PIN kodni qayta tering (Tasdiqlash):</b>',
        'prompt_pin_current': '🔐 <b>Joriy Ma\'muriy PIN kodni tering:</b>',
        'prompt_pin_new': '🆕 <b>Yangi 4 xonali PIN kodni tering:</b>',
        'prompt_req_birth_date': '🎂 Tug\'ilgan sanangizni KK.OO.YYYY formatida kiriting (Masalan: 15.05.2008):',
        'prompt_req_gender': '🚻 Iltimos, jinsingizni tanlang:',
        'prompt_restore_backup': '🔄 Iltimos, maktab zaxira <code>.xlsx</code> faylini yuboring:',
        'prompt_search_student': '🔍 O\'quvchi ismi yoki raqamini kiriting:',
        'prompt_select_exam_type': '📝 O\'quvchi: <b>{name}</b> ({class_name})\n\nBaholash turini tanlang:',
        'prompt_student_class': '🏫 O\'quvchi sinfini kiriting (Masalan: <code>9-A</code>):',
        'prompt_student_name': '👤 O\'quvchining Ism va Familiyasini kiriting:',
        'prompt_student_no': '🔢 O\'quvchining maktab raqamini kiriting (Masalan: <code>101</code>):',
        'prompt_teacher_name': '👨‍🏫 O\'qituvchining Ism va Familiyasini kiriting:',
        'prompt_teacher_subject': '📚 O\'qituvchining fanini kiriting (Masalan: <code>Matematika</code>):',
        'prompt_upload_teacher_excel': '👨‍🏫 <b>O\'qituvchilarni Excel orqali yuklash</b>\n\nUstunlar: <code>Ad Soyad</code> | <code>Brans</code> | <code>(Siniflar)</code>\nFaylni <code>.xlsx</code> formatida yuboring:',
        'published_homeworks_title': '📢 <b>E\'lon qilingan vazifalar:</b>',
        'rate_limit_warning': '⚠️ Juda ko\'p so\'rov yubordingiz. Iltimos, bir necha soniya kuting.',
        'readonly_mode_active_alert': '🔒 Tizim hozirda faqat o\'qish (Read-Only) rejimida. O\'zgartirishlar kiritish taqiqlangan.',
        'readonly_mode_updated': 'Faqat o\'qish rejimi sozlamasi yangilandi.',
        'recent_grades_title': '📝 <b>Oxirgi qo\'yilgan baholar:</b>',
        'remind_att_sent': 'Eslatma yuborildi.',
        'report_not_found': 'Ma\'lumotnoma topilmadi.',
        'req_already_pending': '⚠️ Sizda allaqachon ko\'rib chiqilayotgan ariza mavjud.',
        'req_approved_admin_msg': '✅ Ariza #{id} tasdiqlandi. Foydalanuvchi: <b>{name}</b> ({role})',
        'req_approved_user': '🎉 <b>Tabriklaymiz!</b>\nMaktab ma\'muriyati arizangizni tasdiqladi. Siz <b>{role}</b> sifatida kirdingiz.',
        'req_details_parent': '🧑‍🎓 Farzandingizning Ismi, Sinfiga yoki Raqamini yozing:',
        'req_details_student': '🏫 Maktab raqamingiz va sinfingizni yozing:',
        'req_details_teacher': '📚 Faningiz va ma\'muriyatga eslatmangizni yozing:',
        'req_name_prompt': '👤 Iltimos, Ism va Familiyangizni to\'liq kiriting:',
        'req_phone_prompt': '📱 <b>Telefon raqamingizni kiriting:</b>\n<i>(Yoki pastdagi «Telefon raqamni ulashish» tugmasini bosing)</i>',
        'req_rejected_admin_msg': '❌ Ariza #{id} rad etildi: <b>{name}</b>',
        'req_rejected_user': '❌ Arizangiz ma\'muriyat tomonidan tasdiqlanmadi. Maktab ma\'muriyati bilan bog\'laning.',
        'req_role_select': '🛎️ <b>Kirish va Parol So\'rovi</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, rolingizni tanlang:',
        'req_sent_success': '✅ Arizangiz maktab ma\'muriyatiga yuborildi. Ko\'rib chiqilgach sizga xabar beriladi.',
        'request_already_handled': '⚠️ Ushbu ariza allaqachon ko\'rib chiqilgan.',
        'restore_success': '✅ Tizim arxivi muvaffaqiyatli yuklandi! O\'quvchilar: <b>{s_cnt}</b>, O\'qituvchilar: <b>{t_cnt}</b>',
        'rk_appointments': '🤝 Uchrashuvlar',
        'rk_attendance': '📋 Tezkor davomat',
        'rk_behavior': '⭐ Xulq-atvor',
        'rk_cancel_action': '❌ Amalni bekor qilish',
        'rk_cat_reports': '📊 Hisobotlar va nazorat',
        'rk_cat_requests': '🔔 Tasdiqlash',
        'rk_cat_settings': '⚙️ Sozlamalar',
        'rk_cat_staff': '👥 Xodim & Oʻquvchi',
        'rk_cat_tools': '🛠️ Boshqaruv vositalari',
        'rk_grade': "📝 Baho qo\'yish",
        'rk_homework': '📢 Vazifalar paneli',
        'rk_logout': '🚪 Chiqish',
        'rk_parent_info': "ℹ️ Maktab haqida",
        'rk_report': '📊 Baholar tabeli',
        'rk_restart': '🔄 Qayta ishga tushirish',
        'rk_switch_student': "🧑‍🎓 Bolani tanlash",
        'rk_upload_medical': "🏥 Ma\'lumotnoma",
        'role_parent_btn': '👨‍👩‍👧‍👦 Ota-ona',
        'role_student_btn': '🎓 O\'quvchi',
        'role_teacher_btn': '👨‍🏫 O\'qituvchi',
        'schedule_select_class': '📅 Jadvalini ko\'rmoqchi bo\'lgan sinfni tanlang:',
        'school_admin_title': 'Maktab ma\'muriyati',
        'search_no_results': '❌ O\'quvchi topilmadi.',
        'search_results_title': '🔍 <b>Qidiruv natijalari:</b>',
        'search_user_no_results': '❌ Mos keluvchi foydalanuvchi topilmadi.',
        'search_user_prompt': '🔍 Foydalanuvchi nomi, ismi yoki Telegram ID raqamini kiriting:',
        'search_user_results_title': '🔍 <b>Foydalanuvchi qidiruv natijalari:</b>',
        'select_class_to_co_teacher': '➕ Qo\'shimcha o\'qituvchi biriktirmoqchi bo\'lgan sinfni tanlang:',
        'select_class_to_transfer': '🔄 Boshqa o\'qituvchiga o\'tkazmoqchi bo\'lgan sinfni tanlang:',
        'select_pdf_class': '📄 Kartalarini yuklab olmoqchi bo\'lgan sinfni tanlang:',
        'select_target_teacher': '👨‍🏫 <b>{class_name}</b> sinfi kimga o\'tkazilsin? O\'qituvchini tanlang:',
        'select_teacher_appointment': '🤝 Uchrashuv belgilamoqchi bo\'lgan o\'qituvchini tanlang:',
        'send_dm_prompt': '✉️ <b>Foydalanuvchiga shaxsiy xabar yuborish:</b>\n\n<code>{name}</code> (<code>{id}</code>) uchun xabar matnini kiriting:',
        'setting_updated_toast': 'Sozlama yangilandi.',
        'student_added_card': '✅ <b>O\'quvchi muvaffaqiyatli qo<b>shildi!</b>\n\n👤 F.I.O: <b>{name}</b>\n🏫 Sinf: <b>{class_name}</b> | №: <b>{no}</b>\n\n🔑 <b>Kirish kodlari:</b>\n• O</b>quvchi kodi: <code>{st_code}</code>\n• Ota-ona kodi: <code>{pr_code}</code>',
        'student_card': '👤 <b>O<b>quvchi kartasi</b>\nF.I.O: <b>{name}</b>\nSinf: <b>{class_name}</b> | №: <b>{no}</b>\n\n🔑 <b>Kod holati:</b>\n• O</b>quvchi: <code>{st_code}</code> ({st_status})\n• Ota-ona: <code>{pr_code}</code> ({pr_status})',
        'student_deleted': '🗑️ O\'quvchi tizimdan o\'chirildi.',
        'student_info_updated': '✅ O\'quvchi ma\'lumotlari yangilandi:\n<b>{name}</b> ({class_name} - №: {no})',
        'student_name_invalid': '❌ Iltimos, to\'g\'ri ism kiriting.',
        'student_not_found': 'O\'quvchi topilmadi.',
        'student_switched_success': 'Tanlangan o\'quvchi: <b>{name}</b> ({class_name})',
        'tch_classes_updated': '✅ O\'qituvchining sinflari yangilandi: <b>{classes}</b>',
        'teacher_added_card': '✅ <b>O\'qituvchi ro\'yxatga olindi!</b>\n\n👤 F.I.O: <b>{name}</b>\n📚 Fani: <b>{subject}</b>\n\n🔑 <b>Kirish kodi:</b>\n<code>{code}</code>',
        'teacher_card': '👨‍🏫 <b>O\'qituvchi kartasi</b>\nF.I.O: <b>{name}</b>\nFan: <b>{subject}</b>\n\n🔑 Kod: <code>{code}</code>\nHolat: {status}',
        'teacher_deleted': '🗑️ O\'qituvchi o\'chirildi.',
        'teacher_excel_done': '✅ <b>{count}</b> ta o\'qituvchi qo\'shildi! Parollar biriktirilgan.',
        'teacher_name_invalid': '❌ Iltimos, to\'g\'ri o\'qituvchi ismini kiriting.',
        'teacher_not_found': 'O\'qituvchi topilmadi.',
        'teacher_search_no_results': '❌ O\'qituvchi topilmadi.',
        'teacher_search_prompt': '🔍 O\'qituvchi ismi yoki fanini kiriting:',
        'teacher_search_results_title': '🔍 <b>O\'qituvchi qidiruv natijalari:</b>',
        'temp_admin_assigned_toast': '✅ Foydalanuvchiga {dur} muddatga vaqtinchalik ma\'mur vakolati berildi.',
        'temp_admin_choose_title': '⏱️ <b>Vaqtinchalik ma\'mur muddatini tanlang:</b>',
        'temp_ban_choose_title': '⏱️ <b>Bloklash muddatini tanlang:</b>',
        'timezone_updated': 'Vaqt mintaqasi UTC+{offset} ga sozlandi.',
        'uc_card_title': '👤 <b>FOYDALANUVCHI MA\'LUMOTI VA AMALLAR KARTASI</b>',
        'unauthorized_action': '⛔ Bu amal uchun sizda vakolat yo\'q.',
        'unauthorized_excel_upload': '⛔ Excel yuklashga ruxsatingiz yo\'q.',
        'unban_success': 'Foydalanuvchi blokdan chiqarildi.',
        'upload_med_prompt': 'Iltimos, ma\'lumotnomaning rasmini yuboring:',
        'user_banned_toast': 'Foydalanuvchi bloklandi.',
        'user_not_found_toast': 'Foydalanuvchi topilmadi.',
        'user_temp_banned_notification': '⛔ Hisobingiz xavfsizlik sababli {dur} muddatga vaqtincha bloklandi.',
        'user_temp_banned_toast': 'Foydalanuvchi {dur} muddatga bloklandi.',
        'user_unbanned_toast': 'Foydalanuvchi blokdan chiqarildi.',
        'weekend_attendance_updated': 'Dam olish kunlari davomat sozlamasi yangilandi.',
        'welcome_guest': '🎓 <b>Xush kelibsiz</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, sizga berilgan <b>kirish kodini</b> (Masalan: <code>VELI-123456</code>, <code>HCA-123456</code>, <code>OGR-123456</code>) kiriting yoki kerakli amalni tanlang:',
    },
}
def get_text(key: str, lang: str = "tr", **kwargs) -> str:
    selected_lang = lang if lang in LOCALES and LOCALES[lang] else "tr"
    template = LOCALES.get(selected_lang, {}).get(key) or LOCALES.get("tr", {}).get(key) or f"[{key}]"
    if kwargs:
        try:
            return template.format(**kwargs)
        except Exception:
            return template
    return template

# ======================================================================
# 4. ARAYÜZ VE ETKİLEŞİM KLAVYELERİ (SADE & DERLİ TOPLU DÜZEN)
# ======================================================================

def get_role_label(role: str | None, lang: str = "tr") -> str:
    role_map = {
        "admin": {"tr": "Yönetici", "ru": "Администратор", "uz": "Ma'mur", "en": "Admin"},
        "teacher": {"tr": "Öğretmen", "ru": "Учитель", "uz": "O'qituvchi", "en": "Teacher"},
        "student": {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"},
        "parent": {"tr": "Veli", "ru": "Родитель", "uz": "Ota-ona", "en": "Parent"},
        "guest": {"tr": "Misafir", "ru": "Гость", "uz": "Mehmon", "en": "Guest"},
    }
    r = str(role or "guest").strip().lower()
    return role_map.get(r, {}).get(lang, str(role or "User"))

def get_role_reply_kb(role: str, lang: str = "tr", linked_count: int = 1) -> ReplyKeyboardMarkup:
    keyboard = []
    if role == "admin":
        # Ultra-sade 3 satırlı yönetici kokpit klavyesi (dosyam ve şifre kartı kaldırıldı)
        keyboard = [
            [KeyboardButton(text=get_text("rk_cat_staff", lang)), KeyboardButton(text=get_text("rk_cat_tools_reports", lang))],
            [KeyboardButton(text=get_text("rk_cat_requests", lang)), KeyboardButton(text=get_text("rk_cat_settings", lang))],
            [KeyboardButton(text=get_text("btn_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "teacher":
        # Sadeleştirilmiş 4 satırlı öğretmen klavyesi (Not ve davranış birleştirildi)
        keyboard = [
            [KeyboardButton(text=get_text("rk_attendance", lang)), KeyboardButton(text=get_text("rk_grade_behavior", lang))],
            [KeyboardButton(text=get_text("rk_homework", lang)), KeyboardButton(text=get_text("rk_appointments", lang))],
            [KeyboardButton(text=get_text("btn_proposals", lang)), KeyboardButton(text=get_text("rk_my_credentials", lang))],
            [KeyboardButton(text=get_text("btn_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "parent":
        # Birleşik akıllı veli klavyesi (tek çocuğu olanda Öğrenci Değiştir gizlenir)
        p_row3 = [KeyboardButton(text=get_text("rk_digital_locker", lang))]
        if linked_count > 1:
            p_row3.append(KeyboardButton(text=get_text("rk_switch_student", lang)))
        else:
            p_row3.append(KeyboardButton(text=get_text("rk_my_credentials", lang)))

        keyboard = [
            [KeyboardButton(text=get_text("rk_student_full_report", lang)), KeyboardButton(text=get_text("rk_upload_medical", lang))],
            [KeyboardButton(text=get_text("rk_appointments", lang)), KeyboardButton(text=get_text("btn_bulletin_program", lang))],
            p_row3,
            [KeyboardButton(text=get_text("btn_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "student":
        # Sadeleştirilmiş 3 satırlı öğrenci klavyesi (ölü veli bilgisi kaldırıldı)
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_homework", lang))],
            [KeyboardButton(text=get_text("btn_exam_schedule", lang)), KeyboardButton(text=get_text("rk_digital_locker", lang))],
            [KeyboardButton(text=get_text("btn_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    else:
        keyboard = [
            [KeyboardButton(text=get_text("btn_login_prompt", lang)), KeyboardButton(text=get_text("btn_req_access", lang))],
            [KeyboardButton(text=get_text("btn_lang", lang))]
        ]
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True, one_time_keyboard=True, is_persistent=False)
def get_cancel_reply_kb(lang: str = "tr") -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=get_text("rk_cancel_action", lang))]],
        resize_keyboard=True,
        is_persistent=False
    )


def get_inline_cancel_kb(lang: str = "tr", back_callback: str = "cancel_action") -> InlineKeyboardMarkup:
    cb_data = f"cancel_action:{back_callback}" if back_callback != "cancel_action" and not back_callback.startswith("cancel_action:") else back_callback
    return InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=cb_data)
    ]])


def get_language_inline_kb() -> InlineKeyboardMarkup:
    buttons = [
        [
            InlineKeyboardButton(text="🇹🇷 Türkçe", callback_data="set_lang:tr"),
            InlineKeyboardButton(text="🇷🇺 Русский", callback_data="set_lang:ru")
        ],
        [
            InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="set_lang:uz"),
            InlineKeyboardButton(text="🇬🇧 English", callback_data="set_lang:en")
        ]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def get_nav_buttons(lang: str = "tr", back_callback: str = "adm:dashboard") -> list:
    if not back_callback:
        return []
    return [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=back_callback)]
def get_attendance_grid_kb(students: list, attendance_state: dict, class_name: str, lang: str = "tr") -> InlineKeyboardMarkup:
    inline_keyboard = []
    
    btn_all_txt = {
        "tr": "🟢 Herkes Burada (Hızlı Kaydet)",
        "ru": "🟢 Все присутствуют (Сохранить в 1 клик)",
        "uz": "🟢 Barcha darsda (1 bosishda saqlash)",
        "en": "🟢 All Present (1-Click Save)"
    }.get(lang, "🟢 All Present (1-Click Save)")
    inline_keyboard.append([InlineKeyboardButton(text=btn_all_txt, callback_data=f"att_all_pres:{class_name}")])

    row = []
    for s in students:
        is_absent = attendance_state.get(s.id, False)
        status_icon = "❌" if is_absent else "✅"
        btn_text = f"{status_icon} {s.full_name}"
        row.append(InlineKeyboardButton(text=btn_text, callback_data=f"att_toggle:{s.id}"))
        if len(row) == 2:
            inline_keyboard.append(row)
            row = []
    if row:
        inline_keyboard.append(row)

    inline_keyboard.append([
        InlineKeyboardButton(text=get_text("btn_save_att", lang), callback_data=f"att_save:{class_name}")
    ])
    inline_keyboard.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
    return InlineKeyboardMarkup(inline_keyboard=inline_keyboard)

def get_ack_notification_kb(notification_id: int, lang: str = "tr") -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [
            InlineKeyboardButton(text=get_text("btn_acknowledged", lang), callback_data=f"ack_notif:{notification_id}"),
            InlineKeyboardButton(text=get_text("btn_upload_medical", lang), callback_data="upload_medical_init")
        ]
    ])

def get_exam_type_kb(class_name: str, student_id: int, lang: str = "tr") -> InlineKeyboardMarkup:
    buttons = [
        [
            InlineKeyboardButton(text=get_text("exam_written_1", lang), callback_data="gr_type:1_yazili"),
            InlineKeyboardButton(text=get_text("exam_written_2", lang), callback_data="gr_type:2_yazili")
        ],
        [
            InlineKeyboardButton(text=get_text("exam_oral", lang), callback_data="gr_type:sozlu")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"gr_cls:{class_name}")
        ]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

# --- iPHONE TARZI SESSİZ INLINE & ALT MENÜ TUŞ TAKIMI ---
def get_pin_inline_kb(lang: str = "tr", callback_prefix: str = "pinkey", is_perm_admin: bool = False, step: str = "") -> InlineKeyboardMarkup:
    del_txt = "⌫ Sil" if lang == "tr" else ("⌫ Стереть" if lang == "ru" else ("⌫ O'chirish" if lang == "uz" else "⌫ Del"))
    cancel_txt = "❌ Vazgeç" if lang == "tr" else ("❌ Отмена" if lang == "ru" else ("❌ Bekor" if lang == "uz" else "❌ Cancel"))
    rows = [
        [
            InlineKeyboardButton(text="1", callback_data=f"{callback_prefix}:1"),
            InlineKeyboardButton(text="2", callback_data=f"{callback_prefix}:2"),
            InlineKeyboardButton(text="3", callback_data=f"{callback_prefix}:3")
        ],
        [
            InlineKeyboardButton(text="4", callback_data=f"{callback_prefix}:4"),
            InlineKeyboardButton(text="5", callback_data=f"{callback_prefix}:5"),
            InlineKeyboardButton(text="6", callback_data=f"{callback_prefix}:6")
        ],
        [
            InlineKeyboardButton(text="7", callback_data=f"{callback_prefix}:7"),
            InlineKeyboardButton(text="8", callback_data=f"{callback_prefix}:8"),
            InlineKeyboardButton(text="9", callback_data=f"{callback_prefix}:9")
        ],
        [
            InlineKeyboardButton(text=cancel_txt, callback_data=f"{callback_prefix}:cancel"),
            InlineKeyboardButton(text="0", callback_data=f"{callback_prefix}:0"),
            InlineKeyboardButton(text=del_txt, callback_data=f"{callback_prefix}:del")
        ]
    ]
    if callback_prefix == "chgpin" and is_perm_admin and step == "verify_current":
        rst_txt = "🔑 PIN Sıfırla (Doğrudan Yeni PIN)" if lang == "tr" else ("🔑 Сбросить ПИН / Задать новый" if lang == "ru" else ("🔑 PINni tiklash / Yangi kod" if lang == "uz" else "🔑 Reset PIN / Set New"))
        rows.append([InlineKeyboardButton(text=rst_txt, callback_data="chgpin:perm_reset")])
    return InlineKeyboardMarkup(inline_keyboard=rows)

def get_pin_reply_kb(lang: str = "tr", is_perm_admin: bool = False, step: str = "") -> ReplyKeyboardMarkup:
    del_txt = "⌫ Sil" if lang == "tr" else ("⌫ Стереть" if lang == "ru" else ("⌫ O'chirish" if lang == "uz" else "⌫ Del"))
    cancel_txt = "❌ Vazgeç" if lang == "tr" else ("❌ Отмена" if lang == "ru" else ("❌ Bekor" if lang == "uz" else "❌ Cancel"))
    keyboard = [
        [KeyboardButton(text="1"), KeyboardButton(text="2"), KeyboardButton(text="3")],
        [KeyboardButton(text="4"), KeyboardButton(text="5"), KeyboardButton(text="6")],
        [KeyboardButton(text="7"), KeyboardButton(text="8"), KeyboardButton(text="9")],
        [KeyboardButton(text=cancel_txt), KeyboardButton(text="0"), KeyboardButton(text=del_txt)]
    ]
    if is_perm_admin and step == "verify_current":
        rst_txt = "🔑 PIN Sıfırla" if lang == "tr" else ("🔑 Сбросить ПИН" if lang == "ru" else ("🔑 PINni tiklash" if lang == "uz" else "🔑 Reset PIN"))
        keyboard.append([KeyboardButton(text=rst_txt)])
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True, is_persistent=True)

# ======================================================================
# 5. YARDIMCI SERVİSLER, RESMİ PDF KARNE, EXCEL VE ŞİFRE KARTLARI
# ======================================================================

def setup_pdf_fonts() -> tuple[str, str]:
    font_regular = "Helvetica"
    font_bold = "Helvetica-Bold"
    candidates = [
        ("/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf", "/usr/share/fonts/truetype/droid/DroidSansFallbackFull.ttf"),
        ("/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans-Bold.ttf"),
        ("/usr/share/fonts/dejavu/DejaVuSans.ttf", "/usr/share/fonts/dejavu/DejaVuSans-Bold.ttf"),
        ("C:\\Windows\\Fonts\\arial.ttf", "C:\\Windows\\Fonts\\arialbd.ttf")
    ]
    for f_path, b_path in candidates:
        if os.path.exists(f_path) and os.path.exists(b_path):
            try:
                pdfmetrics.registerFont(TTFont("AppUnicodeFont", f_path))
                pdfmetrics.registerFont(TTFont("AppUnicodeFont-Bold", b_path))
                return "AppUnicodeFont", "AppUnicodeFont-Bold"
            except Exception:
                pass
    return font_regular, font_bold

async def log_audit(session: AsyncSession, user_id: int, user_name: str, action: str, details: str):
    try:
        entry = AuditLog(
            user_id=user_id,
            user_name=user_name,
            action=action,
            details=details
        )
        session.add(entry)
        if "GÜVENLİK ALARMI" in action:
            asyncio.create_task(alert_admins_security_breach(bot, user_id, user_name, action, details))
    except Exception:
        pass

SETTINGS_CACHE = {}
SETTINGS_CACHE_TTL = {}

async def get_cached_setting(key: str, default_val: str = "") -> str:
    import time
    now_t = time.time()
    if key in SETTINGS_CACHE and now_t < SETTINGS_CACHE_TTL.get(key, 0):
        return SETTINGS_CACHE[key]
    try:
        async with AsyncSessionLocal() as session:
            s = await session.get(SystemSetting, key)
            val = s.value if s else default_val
            SETTINGS_CACHE[key] = val
            SETTINGS_CACHE_TTL[key] = now_t + 30.0
            return val
    except Exception:
        return default_val

async def is_readonly_mode_active() -> bool:
    return (await get_cached_setting("readonly_mode", "false")) == "true"

async def get_current_admin_pin() -> str:
    return await get_cached_setting("admin_pin", ADMIN_PIN)


def ensure_role_authorized(user: User | None, allowed_roles: list[str]) -> bool:
    """Kullanıcının belirtilen rollere sahip olup olmadığını doğrulayan merkezi güvenlik filtresi."""
    if not user:
        return False
    if user.is_blacklisted:
        return False
    if user.role in allowed_roles:
        return True
    if "admin" in allowed_roles and (user.admin_type in ["permanent", "temporary"] or user.telegram_id in ADMIN_IDS):
        return True
    return False

def is_admin_user(user: User | None, user_id: int) -> bool:
    if user_id in ADMIN_IDS:
        return True
    if not user or user.role != "admin":
        return False
    if user.admin_type == "temporary" and user.admin_until:
        if user.admin_until <= datetime.utcnow():
            return False
    return True

def generate_secure_code(prefix: str) -> str:
    num_part = str(random.randint(100000, 999999))
    return f"{prefix}-{num_part}"

async def export_all_school_data_excel() -> io.BytesIO:
    async with AsyncSessionLocal() as session:
        students = (await session.execute(select(Student).order_by(Student.class_name, Student.student_number))).scalars().all()
        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        attendances = (await session.execute(select(Attendance).order_by(desc(Attendance.date)))).scalars().all()
        grades = (await session.execute(select(Grade).order_by(desc(Grade.created_at)))).scalars().all()
        behaviors = (await session.execute(select(BehaviorRecord).order_by(desc(BehaviorRecord.created_at)))).scalars().all()

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Ogrenciler"
    ws1.append(["ID", "Ad Soyad", "Sinif", "Okul Numarasi", "Ogrenci Kodu", "Veli Kodu", "Ogrenci Telegram ID"])
    for s in students:
        ws1.append([s.id, s.full_name, s.class_name, s.student_number, s.student_code, s.parent_code, s.student_telegram_id or "-"])

    ws2 = wb.create_sheet(title="Ogretmenler")
    ws2.append(["ID", "Ad Soyad", "Brans", "Giris Kodu", "Telegram ID", "Sorumlu Siniflar"])
    for t in teachers:
        ws2.append([t.id, t.full_name, t.subject, t.auth_code, t.telegram_id or "-", t.assigned_classes or "ALL"])

    ws3 = wb.create_sheet(title="Yoklamalar")
    ws3.append(["Tarih", "Sinif", "Ogrenci ID", "Durum"])
    for a in attendances:
        ws3.append([str(a.date), a.class_name, a.student_id, a.status])

    ws4 = wb.create_sheet(title="Ders Notlari")
    ws4.append(["Tarih", "Ogrenci ID", "Ders", "Sinav Turu", "Not", "Rozet", "Aciklama"])
    for g in grades:
        ws4.append([g.created_at.strftime('%d.%m.%Y'), g.student_id, g.subject, g.exam_type or "1. Yazılı", g.score, g.badge, g.note or ""])

    ws5 = wb.create_sheet(title="Davranis_Rozetleri")
    ws5.append(["Tarih", "Ogrenci ID", "Tur", "Rozet", "Baslik", "Not"])
    for b in behaviors:
        ws5.append([b.created_at.strftime('%d.%m.%Y'), b.student_id, b.behavior_type, b.badge, b.title, b.note or ""])

    out_buf = io.BytesIO()
    await asyncio.to_thread(wb.save, out_buf)
    out_buf.seek(0)
    wb.close()
    return out_buf

async def process_student_excel(file_bytes: bytes) -> tuple[int, io.BytesIO]:
    in_buffer = io.BytesIO(file_bytes)
    wb_in = await asyncio.to_thread(openpyxl.load_workbook, in_buffer, read_only=True, data_only=True)
    sheet = wb_in.active

    created_students = []
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, None)

    async with AsyncSessionLocal() as session:
        existing_codes = set(
            (await session.execute(select(Student.student_code))).scalars().all() +
            (await session.execute(select(Student.parent_code))).scalars().all()
        )

        def get_unique_code(prefix: str) -> str:
            while True:
                c = generate_secure_code(prefix)
                if c not in existing_codes:
                    existing_codes.add(c)
                    return c

        for row in rows:
            if not row or not row[0]:
                continue
            full_name = str(row[0]).strip()
            class_name = str(row[1]).strip().upper() if len(row) > 1 and row[1] else "GENEL"
            raw_no = row[2] if len(row) > 2 else "0"
            if isinstance(raw_no, float) and raw_no.is_integer():
                student_no = str(int(raw_no))
            else:
                student_no = str(raw_no).strip()

            st_code = get_unique_code("OGR")
            pr_code = get_unique_code("VELI")

            student = Student(
                full_name=full_name,
                class_name=class_name,
                student_number=student_no,
                student_code=st_code,
                parent_code=pr_code
            )
            session.add(student)
            created_students.append((full_name, class_name, student_no, st_code, pr_code))

        await session.commit()
    wb_in.close()
    in_buffer.close()

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Giris_Kodlari"
    ws_out.append(["Ad Soyad", "Sinif", "Numara", "Ogrenci_Kodu", "Veli_Kodu"])
    for item in created_students:
        ws_out.append(list(item))

    out_buffer = io.BytesIO()
    wb_out.save(out_buffer)
    out_buffer.seek(0)
    wb_out.close()
    return len(created_students), out_buffer

async def process_teacher_excel(file_bytes: bytes) -> tuple[int, io.BytesIO]:
    in_buffer = io.BytesIO(file_bytes)
    wb_in = await asyncio.to_thread(openpyxl.load_workbook, in_buffer, read_only=True, data_only=True)
    sheet = wb_in.active

    created_teachers = []
    rows = sheet.iter_rows(values_only=True)
    header = next(rows, None)

    async with AsyncSessionLocal() as session:
        existing_codes = set((await session.execute(select(Teacher.auth_code))).scalars().all())

        def get_unique_teacher_code() -> str:
            while True:
                c = generate_secure_code("HCA")
                if c not in existing_codes:
                    existing_codes.add(c)
                    return c

        for row in rows:
            if not row or not row[0]:
                continue
            full_name = str(row[0]).strip()
            subject = str(row[1]).strip() if len(row) > 1 and row[1] else "Genel"
            assigned_cls = str(row[2]).strip() if len(row) > 2 and row[2] else "ALL"

            existing_tch = (await session.execute(select(Teacher).where(Teacher.full_name == full_name))).scalar_one_or_none()
            if existing_tch:
                existing_tch.subject = subject
                existing_tch.assigned_classes = assigned_cls
                auth_code = existing_tch.auth_code
            else:
                auth_code = get_unique_teacher_code()
                teacher = Teacher(
                    full_name=full_name,
                    subject=subject,
                    auth_code=auth_code,
                    assigned_classes=assigned_cls
                )
                session.add(teacher)
            created_teachers.append((full_name, subject, assigned_cls, auth_code))

        await session.commit()
    wb_in.close()
    in_buffer.close()

    wb_out = openpyxl.Workbook()
    ws_out = wb_out.active
    ws_out.title = "Ogretmen_Kodlari"
    ws_out.append(["Ad Soyad", "Brans", "Sorumlu_Siniflar", "Giris_Kodu"])
    for item in created_teachers:
        ws_out.append(list(item))

    out_buffer = io.BytesIO()
    wb_out.save(out_buffer)
    out_buffer.seek(0)
    wb_out.close()
    return len(created_teachers), out_buffer

async def restore_all_school_data_excel(file_bytes: bytes) -> tuple[int, int]:
    in_buffer = io.BytesIO(file_bytes)
    wb = openpyxl.load_workbook(in_buffer, data_only=True)

    st_cnt, tch_cnt = 0, 0
    async with AsyncSessionLocal() as session:
        if "Ogrenciler" in wb.sheetnames:
            ws = wb["Ogrenciler"]
            rows = ws.iter_rows(values_only=True)
            next(rows, None)
            for r in rows:
                if not r or not r[1]: continue
                fn = str(r[1]).strip()
                cls_n = str(r[2]).strip().upper() if r[2] else "GENEL"
                no_v = str(r[3]).strip() if r[3] else "0"
                st_c = str(r[4]).strip() if r[4] else generate_secure_code("OGR")
                pr_c = str(r[5]).strip() if r[5] else generate_secure_code("VELI")

                existing = (await session.execute(select(Student).where(Student.student_code == st_c))).scalar_one_or_none()
                if not existing:
                    session.add(Student(full_name=fn, class_name=cls_n, student_number=no_v, student_code=st_c, parent_code=pr_c))
                    st_cnt += 1

        if "Ogretmenler" in wb.sheetnames:
            ws = wb["Ogretmenler"]
            rows = ws.iter_rows(values_only=True)
            next(rows, None)
            for r in rows:
                if not r or not r[1]: continue
                fn = str(r[1]).strip()
                sb = str(r[2]).strip() if r[2] else "Ders"
                ac = str(r[3]).strip() if r[3] else generate_secure_code("HCA")
                as_cls = str(r[5]).strip() if len(r) > 5 and r[5] else "ALL"

                existing = (await session.execute(select(Teacher).where(Teacher.auth_code == ac))).scalar_one_or_none()
                if not existing:
                    session.add(Teacher(full_name=fn, subject=sb, auth_code=ac, assigned_classes=as_cls))
                    tch_cnt += 1

        await session.commit()
    wb.close()
    in_buffer.close()
    return st_cnt, tch_cnt

async def generate_teachers_pdf_cards(lang: str = "tr") -> io.BytesIO:
    async with AsyncSessionLocal() as session:
        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()

    fn_reg, fn_bold = setup_pdf_fonts()
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name="TCardTitle", parent=styles["Heading2"], fontName=fn_bold, fontSize=13, leading=16, textColor=colors.HexColor("#1A365D"), alignment=1)
    body_style = ParagraphStyle(name="TCardBody", parent=styles["Normal"], fontName=fn_reg, fontSize=8.5, leading=12, textColor=colors.HexColor("#2D3748"))

    header_text = {
        "tr": "<b>ÖĞRETMEN GİRİŞ KARTLARI</b>",
        "ru": "<b>КАРТОЧКИ УЧИТЕЛЕЙ</b>",
        "uz": "<b>MAKTAB BOSHQARUV TIZIMI - O'QITUVCHILAR PAROL KARTALARI</b>",
        "en": "<b>SCHOOL MANAGEMENT SYSTEM - TEACHER ACCESS CARDS</b>"
    }.get(lang, "<b>TEACHER ACCESS CARDS</b>")

    story = [
        Paragraph(header_text, title_style),
        Spacer(1, 12)
    ]
    card_cells = []
    row = []
    lbl_t = {"tr": "Öğretmen", "ru": "Учитель", "uz": "O'qituvchi", "en": "Teacher"}.get(lang, "Teacher")
    lbl_s = {"tr": "Branş", "ru": "Предмет", "uz": "Fan", "en": "Subject"}.get(lang, "Subject")
    lbl_c = {"tr": "Giriş Kodu", "ru": "Код доступа", "uz": "Kirish kodi", "en": "Access Code"}.get(lang, "Code")
    lbl_st = {"tr": "Durum", "ru": "Статус", "uz": "Holat", "en": "Status"}.get(lang, "Status")

    for t in teachers:
        status_txt = {
            "tr": "Aktif Bağlı" if t.telegram_id else "Henüz Giriş Yapmadı",
            "ru": "Активен" if t.telegram_id else "Не вошел",
            "uz": "Faol ulangan" if t.telegram_id else "Hali kirmagan",
            "en": "Active Linked" if t.telegram_id else "Not logged in"
        }.get(lang, "Active" if t.telegram_id else "Pending")

        text = (
            f"<b>{lbl_t}:</b> {escape_md(t.full_name)}<br/>"
            f"<b>{lbl_s}:</b> {escape_md(t.subject)}<br/>"
            f"<b>{lbl_c}:</b> <code>{t.auth_code}</code><br/>"
            f"<b>{lbl_st}:</b> {status_txt}"
        )
        row.append(Paragraph(text, body_style))
        if len(row) == 2:
            card_cells.append(row)
            row = []
    if row:
        row.append(Paragraph("", body_style))
        card_cells.append(row)

    if card_cells:
        table = Table(card_cells, colWidths=[265, 265])
        table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(table)

    await asyncio.to_thread(doc.build, story)
    pdf_buffer.seek(0)
    return pdf_buffer

async def generate_classroom_pdf_cards(class_name: str, lang: str = "tr") -> io.BytesIO:
    async with AsyncSessionLocal() as session:
        stmt = select(Student).where(Student.class_name == class_name).order_by(Student.student_number)
        result = await session.execute(stmt)
        students = result.scalars().all()

    fn_reg, fn_bold = setup_pdf_fonts()
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=25, leftMargin=25, topMargin=25, bottomMargin=25)
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name="CardTitle", parent=styles["Heading2"], fontName=fn_bold, fontSize=13, leading=16, textColor=colors.HexColor("#1A365D"), alignment=1)
    body_style = ParagraphStyle(name="CardBody", parent=styles["Normal"], fontName=fn_reg, fontSize=8.5, leading=12, textColor=colors.HexColor("#2D3748"))

    t_header = {
        "tr": f"<b>{class_name} SINIFI - TELEGRAM GİRİŞ ŞİFRE KARTLARI</b>",
        "ru": f"<b>КЛАСС {class_name} - КАРТОЧКИ ВХОДА TELEGRAM</b>",
        "uz": f"<b>{class_name} SINFI - TELEGRAM KIRISH KARTALARI</b>",
        "en": f"<b>CLASS {class_name} - TELEGRAM ACCESS CARDS</b>"
    }.get(lang, f"<b>{class_name} - ACCESS CARDS</b>")

    lbl_st = {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"}.get(lang, "Student")
    lbl_st_code = {"tr": "Öğrenci Kodu", "ru": "Код ученика", "uz": "O'quvchi kodi", "en": "Student Code"}.get(lang, "Student Code")
    lbl_pr_code = {"tr": "Veli Kodu", "ru": "Код родителя", "uz": "Ota-ona kodi", "en": "Parent Code"}.get(lang, "Parent Code")

    story = [
        Paragraph(t_header, title_style),
        Spacer(1, 12)
    ]

    card_cells = []
    row = []
    for s in students:
        text = (
            f"<b>{lbl_st}:</b> {escape_md(s.full_name)} (№ {s.student_number})<br/>"
            f"<b>{lbl_st_code}:</b> <code>{s.student_code}</code><br/>"
            f"<b>{lbl_pr_code}:</b> <code>{s.parent_code}</code>"
        )
        row.append(Paragraph(text, body_style))
        if len(row) == 2:
            card_cells.append(row)
            row = []
    if row:
        row.append(Paragraph("", body_style))
        card_cells.append(row)

    if card_cells:
        table = Table(card_cells, colWidths=[265, 265])
        table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("TOPPADDING", (0, 0), (-1, -1), 7),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ("LEFTPADDING", (0, 0), (-1, -1), 10),
            ("RIGHTPADDING", (0, 0), (-1, -1), 10),
            ("VALIGN", (0, 0), (-1, -1), "TOP"),
        ]))
        story.append(table)

    await asyncio.to_thread(doc.build, story)
    pdf_buffer.seek(0)
    return pdf_buffer

async def generate_student_report_card_pdf(student_id: int, lang: str = "tr") -> io.BytesIO:
    async with AsyncSessionLocal() as session:
        st = await session.get(Student, student_id)
        if not st: raise ValueError("Öğrenci bulunamadı")

        att_records = (await session.execute(select(Attendance).where(Attendance.student_id == st.id))).scalars().all()
        absent_cnt = sum(1 for a in att_records if a.status == "absent")
        excused_cnt = sum(1 for a in att_records if a.status == "excused")
        grades = (await session.execute(select(Grade).where(Grade.student_id == st.id).order_by(Grade.created_at.desc()))).scalars().all()
        behaviors = (await session.execute(select(BehaviorRecord).where(BehaviorRecord.student_id == st.id).order_by(BehaviorRecord.created_at.desc()))).scalars().all()

    fn_reg, fn_bold = setup_pdf_fonts()
    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(pdf_buffer, pagesize=A4, rightMargin=24, leftMargin=24, topMargin=24, bottomMargin=24)
    styles = getSampleStyleSheet()

    t_style = ParagraphStyle('RTitle', parent=styles['Heading1'], fontName=fn_bold, fontSize=13, leading=16, textColor=colors.HexColor('#1A365D'), alignment=1)
    c_bold = ParagraphStyle('RCB', parent=styles['Normal'], fontName=fn_bold, fontSize=8, leading=10, textColor=colors.HexColor('#1A202C'))
    c_norm = ParagraphStyle('RCN', parent=styles['Normal'], fontName=fn_reg, fontSize=7.5, leading=9.5, textColor=colors.HexColor('#2D3748'))
    c_center_b = ParagraphStyle('RCCB', parent=styles['Normal'], fontName=fn_bold, fontSize=8, leading=10, textColor=colors.HexColor('#1A202C'), alignment=1)
    c_center = ParagraphStyle('RCC', parent=styles['Normal'], fontName=fn_reg, fontSize=7.5, leading=9.5, textColor=colors.HexColor('#2D3748'), alignment=1)

    t_main = {
        "tr": "<b>RESMİ ÖĞRENCİ GELİŞİM VE NOT KARNESİ</b>",
        "ru": "<b>ОФИЦИАЛЬНЫЙ ТАБЕЛЬ УСПЕВАЕМОСТИ</b>",
        "uz": "<b>MAKTAB BOSHQARUV TIZIMI - RASMIY O'QUVCHI BAHOLAR VA RIVOJLANISH KUNDALIGI</b>",
        "en": "<b>SCHOOL MANAGEMENT SYSTEM - OFFICIAL STUDENT REPORT CARD</b>"
    }.get(lang, "<b>OFFICIAL STUDENT REPORT CARD</b>")

    story = [Paragraph(t_main, t_style), Spacer(1, 10)]
    date_str = get_local_date().strftime('%d.%m.%Y')

    subject_map = {}
    for g in grades:
        subject_map.setdefault(g.subject, []).append(g.score)
    
    subject_averages = {s: sum(scores)/len(scores) for s, scores in subject_map.items()}
    avg_score = (sum(subject_averages.values()) / len(subject_averages)) if subject_averages else 100.0

    lbl_st = {"tr": "Öğrenci:", "ru": "Ученик:", "uz": "O'quvchi:", "en": "Student:"}.get(lang, "Student:")
    lbl_dt = {"tr": "Tarih:", "ru": "Дата:", "uz": "Sana:", "en": "Date:"}.get(lang, "Date:")
    lbl_cl = {"tr": "Sınıf:", "ru": "Класс:", "uz": "Sinf:", "en": "Class:"}.get(lang, "Class:")
    lbl_no = {"tr": "Numara:", "ru": "Номер:", "uz": "Raqam:", "en": "Roll No:"}.get(lang, "Roll No:")
    lbl_att = {"tr": "Devamsızlık:", "ru": "Пропуски:", "uz": "Davomat:", "en": "Absences:"}.get(lang, "Absences:")
    lbl_avg = {"tr": "Ağırlıklı Ortalama:", "ru": "Средний балл:", "uz": "O'rtacha ball:", "en": "Weighted GPA:"}.get(lang, "GPA:")

    att_text = f"{absent_cnt} (İzinli: {excused_cnt})" if lang == "tr" else (f"{absent_cnt} (Уваж: {excused_cnt})" if lang == "ru" else (f"{absent_cnt} (Ruxsatli: {excused_cnt})" if lang == "uz" else f"{absent_cnt} (Excused: {excused_cnt})"))

    meta = [
        [Paragraph(f"<b>{lbl_st}</b>", c_bold), Paragraph(st.full_name, c_norm), Paragraph(f"<b>{lbl_dt}</b>", c_bold), Paragraph(date_str, c_norm)],
        [Paragraph(f"<b>{lbl_cl}</b>", c_bold), Paragraph(st.class_name, c_norm), Paragraph(f"<b>{lbl_no}</b>", c_bold), Paragraph(st.student_number, c_norm)],
        [Paragraph(f"<b>{lbl_att}</b>", c_bold), Paragraph(att_text, c_norm), Paragraph(f"<b>{lbl_avg}</b>", c_bold), Paragraph(f"<b>{avg_score:.1f} / 100</b>", c_bold)],
    ]
    t_meta = Table(meta, colWidths=[130, 140, 130, 147])
    t_meta.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor('#F7FAFC')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 6),
        ('RIGHTPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_meta)
    story.append(Spacer(1, 10))

    col_sub = {"tr": "Ders Adı", "ru": "Предмет", "uz": "Fan", "en": "Subject"}.get(lang, "Subject")
    col_type = {"tr": "Sınav Türü", "ru": "Вид оценки", "uz": "Baho turi", "en": "Type"}.get(lang, "Type")
    col_score = {"tr": "Puan", "ru": "Балл", "uz": "Baho", "en": "Score"}.get(lang, "Score")
    col_badge = {"tr": "Durum", "ru": "Статус", "uz": "Holat", "en": "Status"}.get(lang, "Status")
    col_date = {"tr": "Tarih", "ru": "Дата", "uz": "Sana", "en": "Date"}.get(lang, "Date")

    table_data = [[Paragraph(f"<b>{col_sub}</b>", c_center_b), Paragraph(f"<b>{col_type}</b>", c_center_b), Paragraph(f"<b>{col_score}</b>", c_center_b), Paragraph(f"<b>{col_badge}</b>", c_center_b), Paragraph(f"<b>{col_date}</b>", c_center_b)]]
    if not grades:
        table_data.append([Paragraph("-", c_norm), Paragraph("-", c_center), Paragraph("-", c_center), Paragraph("-", c_center), Paragraph("-", c_center)])
    else:
        for g in grades:
            table_data.append([Paragraph(g.subject, c_norm), Paragraph(g.exam_type or "1. Yazılı", c_norm), Paragraph(str(g.score), c_center_b), Paragraph(g.badge, c_center), Paragraph(g.created_at.strftime('%d.%m.%Y'), c_center)])

    t_grades = Table(table_data, colWidths=[140, 140, 80, 90, 97])
    t_grades.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#EDF2F7')),
        ('BOX', (0,0), (-1,-1), 1, colors.HexColor('#CBD5E0')),
        ('INNERGRID', (0,0), (-1,-1), 0.5, colors.HexColor('#E2E8F0')),
        ('TOPPADDING', (0,0), (-1,-1), 4),
        ('BOTTOMPADDING', (0,0), (-1,-1), 4),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
    ]))
    story.append(t_grades)
    await asyncio.to_thread(doc.build, story)
    pdf_buffer.seek(0)
    return pdf_buffer

# ======================================================================
# 6. ARKA PLAN İŞÇİLERİ
# ======================================================================

async def run_attendance_delay_worker(bot: Bot):
    async with AsyncSessionLocal() as session:
        now = datetime.utcnow()
        stmt = select(Attendance, Student).join(Student, Attendance.student_id == Student.id).where(Attendance.notify_at <= now, Attendance.is_notified == False)
        records = (await session.execute(stmt)).all()

        for att, st in records:
            if att.status == "absent":
                parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st.id))).scalars().all()
                st_name_esc = escape_md(st.full_name)
                for p_id in parents:
                    user = await session.get(User, p_id)
                    p_lang = user.language if user else "tr"
                    msg_text = {
                        "tr": f"🚨 *DEVAMSIZLIK BİLDİRİMİ*\n\nÖğrenciniz *{st_name_esc}*, bugün ({att.date}) okul yoklamasında *GELMEDİ (YOK)* olarak kaydedilmiştir.",
                        "ru": f"🚨 *УВЕДОМЛЕНИЕ О ПРОПУСКЕ*\n\nВаш ребенок *{st_name_esc}* сегодня ({att.date}) отмечен(а) как *ОТСУТСТВУЕТ*.",
                        "uz": f"🚨 *DAVOMAT OGOHLANTIRISHI*\n\nFarzandingiz *{st_name_esc}* bugun ({att.date}) maktab davomatida *KELMADI* deb qayd etildi.",
                        "en": f"🚨 *ABSENCE ALERT*\n\nYour student *{st_name_esc}* has been marked *ABSENT* today ({att.date})."
                    }.get(p_lang, f"🚨 ABSENCE: {st_name_esc}")

                    notif = CriticalNotification(user_telegram_id=p_id, message_text=msg_text)
                    session.add(notif)
                    await session.flush()
                    await safe_send_message(bot=bot, chat_id=p_id, text=msg_text, reply_markup=get_ack_notification_kb(notif.id, lang=p_lang), parse_mode="HTML")

                total_absent = (await session.execute(select(func.count(Attendance.id)).where(Attendance.student_id == st.id, Attendance.status == "absent"))).scalar() or 0
                if total_absent in (5, 10, 15):
                    for p_id in parents:
                        p_user = await session.get(User, p_id)
                        p_l = p_user.language if p_user else "tr"
                        warn_txt = get_text("legal_absence_alert", p_l, name=st_name_esc, count=total_absent)
                        await safe_send_message(bot=bot, chat_id=p_id, text=warn_txt, parse_mode="HTML")

            att.is_notified = True
        await session.commit()

async def run_evening_briefing_worker(bot: Bot):
    async with AsyncSessionLocal() as session:
        today = datetime.utcnow().date()
        parents = (await session.execute(select(User).where(User.role == "parent", User.evening_briefing == True))).scalars().all()
        for p in parents:
            if not p.current_child_id: continue
            st = await session.get(Student, p.current_child_id)
            if not st: continue

            att = (await session.execute(select(Attendance).where(Attendance.student_id == st.id, Attendance.date == today))).scalar_one_or_none()
            att_status = "✅ Geldi" if not att or att.status == "present" else ("❌ Gelmedi" if att.status == "absent" else "🏥 İzinli")
            grades = (await session.execute(select(Grade).where(Grade.student_id == st.id, func.date(Grade.created_at) == today))).scalars().all()
            grades_str = "\n".join([f"• {g.subject}: *{g.score}* ({g.badge})" for g in grades]) if grades else get_text("no_grades", p.language)

            text = get_text("evening_briefing_header", p.language, name=escape_md(st.full_name), class_name=escape_md(st.class_name), att_status=att_status, grades=grades_str)
            await safe_send_message(bot, p.telegram_id, text, parse_mode="HTML", disable_notification=True)
            await asyncio.sleep(0.05)

async def run_morning_briefing_worker(bot: Bot):
    async with AsyncSessionLocal() as session:
        today = get_local_date()
        total_students = (await session.execute(select(func.count(Student.id)))).scalar() or 0
        present = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "present"))).scalar() or 0
        absent = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "absent"))).scalar() or 0
        all_classes = (await session.execute(select(Student.class_name).distinct())).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()
        missing = [c for c in all_classes if c not in taken_classes]
        missing_str = ", ".join(missing) if missing else "Tümü Alındı"

        admin_users = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
        admin_lang_map = {u.telegram_id: u.language for u in admin_users}
        for a_id in set(ADMIN_IDS + list(admin_lang_map.keys())):
            a_lang = admin_lang_map.get(a_id, "tr")
            msg = {
                "tr": f"☀️ *GÜNLÜK İDARİ BRİFİNG (09:30)*\n\n🏫 Toplam: {total_students} | ✅ Gelen: {present} | ❌ Gelmeyen: {absent}\n⚠️ Yoklama Girmeyen: {missing_str}",
                "ru": f"☀️ *ЕЖЕДНЕВНЫЙ ОТЧЕТ ДИРЕКТОРУ (09:30)*\n\n🏫 Всего: {total_students} | ✅ Присутствуют: {present} | ❌ Отсутствуют: {absent}\n⚠️ Без переклички: {missing_str}",
                "uz": f"☀️ *KUNLIK TONGGI HISOBOT (09:30)*\n\n🏫 Jami: {total_students} | ✅ Kelgan: {present} | ❌ Kelmagan: {absent}\n⚠️ Davomat olinmagan: {missing_str}",
                "en": f"☀️ *DAILY EXECUTIVE BRIEFING (09:30)*\n\n🏫 Total: {total_students} | ✅ Present: {present} | ❌ Absent: {absent}\n⚠️ Pending Attendance: {missing_str}"
            }.get(a_lang, f"☀️ *DAILY BRIEFING (09:30)*")
            await safe_send_message(bot, a_id, msg, parse_mode="HTML")

async def run_friday_backup_worker(bot: Bot):
    try:
        buf = await export_all_school_data_excel()
        today_str = get_local_date().strftime('%d_%m_%Y')
        data_bytes = buf.read()
        buf.close()
        backup_captions = {
            "tr": f"📦 *HAFTALIK OTOMATİK OKUL VERİ YEDEĞİ (CUMA 18:00)*\n📅 Tarih: {today_str}\nOkulunuzun tüm güncel verileri ektedir.",
            "ru": f"📦 *ЕЖЕНЕДЕЛЬНЫЙ АВТОМАТИЧЕСКИЙ АРХИВ (ПЯТНИЦА 18:00)*\n📅 Дата: {today_str}\nВсе актуальные данные школы прикреплены.",
            "uz": f"📦 *HAFTALIK AVTOMATIK MAKTAB ZAXIRASI (JUMA 18:00)*\n📅 Sana: {today_str}\nMaktabingizning barcha joriy ma'lumotlari biriktirildi.",
            "en": f"📦 *WEEKLY AUTOMATIC SCHOOL DATA BACKUP (FRIDAY 18:00)*\n📅 Date: {today_str}\nComplete school data backup is attached."
        }
        async with AsyncSessionLocal() as session:
            admin_users = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
            admin_lang_map = {u.telegram_id: u.language for u in admin_users}
        for a_id in set(ADMIN_IDS + list(admin_lang_map.keys())):
            a_lang = admin_lang_map.get(a_id, "tr")
            caption_txt = backup_captions.get(a_lang, backup_captions["en"])
            doc_file = BufferedInputFile(data_bytes, filename=f"Okul_Yedek_{today_str}.xlsx")
            await bot.send_document(chat_id=a_id, document=doc_file, caption=caption_txt, parse_mode="HTML")
            await asyncio.sleep(0.05)
    except Exception as e:
        pass

# ======================================================================
# 7. FSM DURUMLARI VE ÖNBELLEK
# ======================================================================

class Form(StatesGroup):
    waiting_school_name = State()
    waiting_auth_code = State()
    waiting_admin_tg_id = State()
    waiting_admin_name = State()
    waiting_search_teacher_query = State()
    sched_update_text = State()
    waiting_search_user_query = State()
    waiting_admin_dm_text = State()
    req_role = State()
    req_name = State()
    req_gender = State()
    req_birth_date = State()
    req_phone = State()
    req_details = State()
    rename_class_name = State()
    waiting_prop_deadline = State()
    waiting_prop_anon = State()
    waiting_prop_attachment = State()
    add_class_name = State()
    waiting_prop_student_class = State()
    waiting_prop_student_id = State()
    waiting_prop_content = State()
    waiting_prop_reason = State()
    waiting_admin_prop_title = State()
    waiting_admin_prop_content = State()
    waiting_admin_decision_note = State()
    waiting_force_alert_text = State()
    add_student_name = State()
    add_student_class = State()
    add_student_no = State()
    add_teacher_name = State()
    add_teacher_subject = State()
    waiting_search_query = State()
    waiting_broadcast_text = State()
    waiting_medical_photo = State()
    grade_score = State()
    grade_badge = State()
    app_note = State()
    menu_update_text = State()
    edit_student_val = State()
    parent_add_child_code = State()
    edit_grade_val = State()
    waiting_behavior_note = State()
    waiting_hw_submission = State()
    waiting_hw_feedback = State()
    waiting_exam_subject = State()
    waiting_exam_date = State()
    waiting_emergency_text = State()
    hw_content = State()
    waiting_teacher_excel = State()
    waiting_restore_excel = State()
    appr_st_class = State()
    appr_st_no = State()

router = Router()
ATTENDANCE_CACHE = {}
GRADE_CACHE = {}
HW_CACHE = {}
REQ_CACHE = {}
APP_CACHE = {}
BEHAVIOR_CACHE = {}
EXAM_CACHE = {}
SUBMISSION_CACHE = {}
BC_CACHE = {}
PROP_CACHE = {}
FORCE_ALERT_CACHE = {}
VOTE_REASON_CACHE = {}
PIN_PENDING_ACTIONS = {}
PIN_CHANGE_SESSION = {}
ADMIN_PIN_INPUT = {}
ADMIN_PIN_FAILURES = {}
PIN_MSG_ID = {}
PIN_CHAT_ID = {}
PIN_SUB_MSG_ID = {}
ADMIN_DISPATCHED_NOTIFS = {}


async def record_user_intelligence_telemetry(user_id: int, full_name: str | None = None, username: str | None = None, action_type: str = "general", action_desc: str = ""):
    if not user_id:
        return
    now = datetime.utcnow()
    try:
        async with AsyncSessionLocal() as session:
            u = await session.get(User, user_id)
            if not u:
                return
            u.last_seen_at = now
            if action_desc:
                u.last_action_desc = action_desc[:250]

            local_hour = (now.hour + TIMEZONE_OFFSET) % 24
            if 1 <= local_hour < 6:
                u.night_activity_count = (u.night_activity_count or 0) + 1

            # Check identity morphing
            name_changes = []
            try:
                name_changes = json.loads(u.previous_names) if u.previous_names else []
            except Exception:
                name_changes = []

            changed = False
            if username and u.username and u.username != username:
                name_changes.append({"type": "username", "old": f"@{u.username}", "new": f"@{username}", "date": now.strftime("%d.%m.%Y %H:%M")})
                u.username = username
                changed = True
            elif username and not u.username:
                u.username = username
                changed = True

            if full_name and u.full_name and u.full_name.strip() != full_name.strip():
                name_changes.append({"type": "full_name", "old": u.full_name, "new": full_name, "date": now.strftime("%d.%m.%Y %H:%M")})
                u.full_name = full_name
                changed = True

            if changed:
                u.previous_names = json.dumps(name_changes[-10:], ensure_ascii=False)

            if action_type:
                p_counts = {}
                try:
                    p_counts = json.loads(u.persona_counts) if u.persona_counts else {}
                except Exception:
                    p_counts = {}
                p_counts[action_type] = p_counts.get(action_type, 0) + 1
                u.persona_counts = json.dumps(p_counts, ensure_ascii=False)

            if action_type == "download":
                u.download_count = (u.download_count or 0) + 1

            await session.commit()
    except Exception:
        pass


async def sync_admin_notif_resolution(bot_obj: Bot, key: str, actor_id: int, actor_name: str, status: str, summary: str):
    dispatched = ADMIN_DISPATCHED_NOTIFS.pop(key, [])

    async with AsyncSessionLocal() as session:
        admin_users = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
        admin_lang_map = {u.telegram_id: (u.language or "tr") for u in admin_users}
        for a_id in ADMIN_IDS:
            if a_id not in admin_lang_map:
                admin_lang_map[a_id] = "tr"

    now_str = datetime.now().strftime("%d.%m.%Y %H:%M")

    for a_id, msg_id in dispatched:
        # 1. Silme işlemi: Tüm adminlerin sohbetindeki beklemede olan bildirim kartını yok et!
        try:
            await bot_obj.delete_message(chat_id=a_id, message_id=msg_id)
        except Exception:
            try:
                await bot_obj.edit_message_reply_markup(chat_id=a_id, message_id=msg_id, reply_markup=None)
            except Exception:
                pass

        # 2. Bilgilendirme: İşlemi yapan admin HARİÇ diğer tüm adminlere anlık haber ver!
        if a_id != actor_id:
            a_lang = admin_lang_map.get(a_id, "tr")
            status_lbl = {
                "approved": {"tr": "✅ ONAYLANDI", "ru": "✅ ОДОБРЕНО", "uz": "✅ TASDIQLANDI", "en": "✅ APPROVED"},
                "rejected": {"tr": "❌ REDDEDİLDİ", "ru": "❌ ОТКЛОНЕНО", "uz": "❌ RAD ETILDI", "en": "❌ REJECTED"}
            }.get(status, {}).get(a_lang, status.upper())

            info_card = {
                "tr": (
                    "📢 <b>YÖNETİCİ BİLGİLENDİRMESİ</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Durum:</b> {status_lbl}\n"
                    f"👤 <b>İşlemi Yapan:</b> <b>{escape_html(actor_name)}</b>\n"
                    f"📝 <b>Açıklama:</b> {escape_html(summary)}\n"
                    f"📅 <b>Tarih:</b> {now_str}\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "ru": (
                    "📢 <b>УВЕДОМЛЕНИЕ ДЛЯ АДМИНИСТРАЦИИ</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Статус:</b> {status_lbl}\n"
                    f"👤 <b>Администратор:</b> <b>{escape_html(actor_name)}</b>\n"
                    f"📝 <b>Детали:</b> {escape_html(summary)}\n"
                    f"📅 <b>Дата:</b> {now_str}\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "uz": (
                    "📢 <b>MA'MURIYAT UCHUN XABARNOMA</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Holat:</b> {status_lbl}\n"
                    f"👤 <b>Bajaruvchi ma'mur:</b> <b>{escape_html(actor_name)}</b>\n"
                    f"📝 <b>Tafsilot:</b> {escape_html(summary)}\n"
                    f"📅 <b>Sana:</b> {now_str}\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "en": (
                    "📢 <b>ADMINISTRATIVE UPDATE</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Action:</b> {status_lbl}\n"
                    f"👤 <b>Handled by:</b> <b>{escape_html(actor_name)}</b>\n"
                    f"📝 <b>Details:</b> {escape_html(summary)}\n"
                    f"📅 <b>Date:</b> {now_str}\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                )
            }.get(a_lang, f"Request {status} by {actor_name}")

            try:
                await bot_obj.send_message(chat_id=a_id, text=info_card, parse_mode="HTML")
            except Exception:
                pass


# ======================================================================
# 8. ROUTER: BAŞLANGIÇ, PROFİL, KİMLİK DOĞRULAMA VE OTURUM KAPATMA
# ======================================================================

@router.message(any_state, Command("profil", "profile"))
async def cmd_profile(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not user: return
        lang = user.language
        role_label = {
            "admin": get_text("lbl_role_admin", lang),
            "teacher": get_text("lbl_role_teacher", lang),
            "parent": get_text("lbl_role_parent", lang),
            "student": get_text("lbl_role_student", lang)
        }.get(user.role, user.role)

        lbl_s = get_text("lbl_subject", lang)
        lbl_cl = get_text("lbl_class", lang)
        lbl_no = get_text("lbl_number", lang)
        lbl_st = get_text("lbl_role_student", lang)

        extra = ""
        if user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: extra = f"• {lbl_st}: {escape_md(st.full_name)} ({escape_md(st.class_name)})\n"
        elif user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            if st: extra = f"• {lbl_cl}: {escape_md(st.class_name)} | {lbl_no}: {escape_md(st.student_number)}\n"
        elif user.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
            if tch: extra = f"• {lbl_s}: {escape_md(tch.subject)}\n"

        p_title = get_text("uc_card_title", lang)
        lbl_name = get_text("lbl_full_name", lang)
        lbl_r = get_text("lbl_role", lang)
        lbl_user = get_text("lbl_role_guest", lang)

        text = (
            f"{p_title}\n\n"
            f"• {lbl_name}: {escape_md(user.full_name or lbl_user)}\n"
            f"• {lbl_r}: {role_label}\n"
            f"• Telegram ID: {user.telegram_id}\n"
            f"{extra}"
        )
        await safe_edit_or_answer(message, text, parse_mode=None)

@router.message(any_state, Command("help", "yardim"))
async def cmd_help(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        role_label = {
            "admin": get_text("lbl_role_admin", lang),
            "teacher": get_text("lbl_role_teacher", lang),
            "parent": get_text("lbl_role_parent", lang),
            "student": get_text("lbl_role_student", lang),
            "guest": get_text("lbl_role_guest", lang)
        }.get(user.role if user else "guest", "User")

        help_texts = {
            "tr": f"📖 Okul Yönetim Sistemi Kılavuzu\n\nRolünüz: {role_label}\n• Alt menüyü kullanarak işlemlerinizi yapabilirsiniz.\n• İptal için /cancel yazabilirsiniz.\n• Destek için okul idaresine başvurabilirsiniz.",
            "ru": f"📖 Справка по системе управления школой\n\nВаша роль: {role_label}\n• Используйте нижнее меню для навигации.\n• Для отмены напишите /cancel.\n• По вопросам обращайтесь к администрации.",
            "uz": f"📖 Maktab boshqaruv tizimi bo'yicha qo'llanma\n\nSizning rolingiz: {role_label}\n• Pastdagi menyu orqali amallarni bajarishingiz mumkin.\n• Bekor qilish uchun /cancel yozishingiz mumkin.\n• Savollar bo'yicha maktab ma'muriyatiga murojaat qiling.",
            "en": f"📖 School Management System Guide\n\nYour Role: {role_label}\n• Use the bottom menu to navigate.\n• Type /cancel to abort any action.\n• Contact school administration for support."
        }
        await safe_edit_or_answer(message, help_texts.get(lang, help_texts["en"]), parse_mode=None)

def is_universal_cancel_text(text: str | None) -> bool:
    if not text: return False
    clean = text.strip().lower()
    clean_alnum = re.sub(r'[^\w\s]', '', clean).strip()
    keywords = ["iptal", "bekor", "отмен", "cancel", "vazgeç", "vazgec", "amalni bekor", "отменить", "islemi iptal"]
    return any(kw in clean_alnum for kw in keywords)

@router.message(any_state, Command("cancel", "iptal", "bekor", "otmena"))
@router.message(any_state, F.text.func(is_universal_cancel_text))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    for cache_dict in [ATTENDANCE_CACHE, GRADE_CACHE, HW_CACHE, REQ_CACHE, APP_CACHE, BEHAVIOR_CACHE, SUBMISSION_CACHE, EXAM_CACHE, BC_CACHE, PIN_PENDING_ACTIONS, PIN_CHANGE_SESSION, ADMIN_PIN_INPUT, ADMIN_PIN_FAILURES]:
        cache_dict.pop(user_id, None)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    # Kullanıcının gönderdiği iptal mesajını sil
    try: await message.delete()
    except Exception: pass

    # Açık olan PIN mesajını veya bekleyen kartları temizle
    p_mid = PIN_MSG_ID.pop(user_id, None)
    PIN_CHAT_ID.pop(user_id, None)
    if p_mid:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=p_mid)
        except Exception: pass

    last_mid = LAST_MENU_MSG_ID.pop(message.chat.id, None)
    if last_mid:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=last_mid)
        except Exception: pass

    await purge_previous_bot_messages(message.bot, message.chat.id)

    # İptal sonrası menünün kaybolmaması için doğrudan kalıcı alt menülü dashboard gönder
    if user and user.role != "guest":
        role_kb = get_role_reply_kb(user.role, lang)
        await render_clean_dashboard(message, user, reply_kb=role_kb)
    else:
        guest_kb = get_role_reply_kb("guest", lang)
        s_m = await message.bot.send_message(chat_id=message.chat.id, text=get_text("welcome_guest", lang), reply_markup=guest_kb)
        if s_m:
            LAST_MENU_MSG_ID[message.chat.id] = s_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(s_m.message_id)


@router.callback_query(F.data.func(lambda d: d == "cancel_action" or (isinstance(d, str) and d.startswith("cancel_action:"))))
async def cb_cancel_action(query: CallbackQuery, state: FSMContext):
    await state.clear()
    user_id = query.from_user.id
    for cache_dict in [ATTENDANCE_CACHE, GRADE_CACHE, HW_CACHE, REQ_CACHE, APP_CACHE, BEHAVIOR_CACHE, SUBMISSION_CACHE, EXAM_CACHE, BC_CACHE, PIN_PENDING_ACTIONS, PIN_CHANGE_SESSION, ADMIN_PIN_INPUT, ADMIN_PIN_FAILURES, PROP_CACHE]:
        cache_dict.pop(user_id, None)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    p_mid = PIN_MSG_ID.pop(user_id, None)
    PIN_CHAT_ID.pop(user_id, None)
    if p_mid:
        try: await query.bot.delete_message(chat_id=query.message.chat.id, message_id=p_mid)
        except Exception: pass

    target = query.data.split(":", 1)[1] if ":" in query.data else ""
    if target:
        query.data = target
        if target.startswith("prop:results:"):
            query.data = target
            await cb_proposal_view_results(query)
            return
        elif target.startswith("prop:open:"):
            query.data = target
            await cb_open_proposal_for_vote(query)
            return
        elif target == "adm:classes":
            await cb_classes_list(query, state)
            return
        elif target == "adm:cat_staff":
            await cb_cat_staff(query, state)
            return
        elif target == "adm:cat_settings":
            await cb_cat_settings(query, state)
            return
        elif target == "adm:cat_tools_reports":
            await cb_cat_tools_reports(query, state)
            return
        elif target == "adm:excel_hub":
            await cb_admin_excel_hub(query)
            return
        elif target == "adm:proposals_hub":
            await cb_admin_proposals_hub(query)
            return
        elif target == "adm:teachers":
            await cb_admin_teachers_list(query)
            return
        elif target.startswith("adm:show_class:"):
            await cb_show_class_students(query)
            return

    # Default fallback: render clean dashboard in-place
    if user and user.role != "guest":
        role_kb = get_role_reply_kb(user.role, lang)
        await render_clean_dashboard(query, user, reply_kb=role_kb)
    else:
        guest_kb = get_role_reply_kb("guest", lang)
        await safe_edit_or_answer(query, get_text("unreg_welcome", lang), reply_markup=guest_kb)


@router.message(any_state, Command("logout", "cikis"))
@router.message(any_state, F.text.in_(["🚪 Çıkış Yap", "🚪 Выйти", "🚪 Chiqish", "🚪 Log Out", "/logout", "/cikis"]))
async def handle_bot_logout_cmd(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if user:
            user.role = "guest"
            user.current_child_id = None
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            lang = user.language
        else:
            lang = "tr"

    try: await message.delete()
    except Exception: pass
    KEYBOARD_ANCHOR_MSG_ID.pop(message.chat.id, None)
    if user:
        await render_clean_dashboard(message, user)
    else:
        guest_kb = get_role_reply_kb("guest", lang)
        combined_lgt = f"{get_text('logout_success_msg', lang)}\n──────────────\n{get_text('welcome_guest', lang)}"
        await safe_edit_or_answer(message, combined_lgt, reply_markup=guest_kb, parse_mode="HTML")

@router.my_chat_member()
async def on_my_chat_member_updated(event: ChatMemberUpdated, bot: Bot | None = None):
    user_id = event.from_user.id
    old_state = event.old_chat_member.status
    new_state = event.new_chat_member.status
    now = datetime.utcnow()
    bot_obj = bot or event.bot

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        u_name = user.full_name if user and user.full_name else (event.from_user.full_name or "Kullanıcı")
        u_username = f"@{event.from_user.username}" if event.from_user.username else "-"
        u_role = user.role if user else "guest"

        if new_state in ("kicked", "left"):
            if user:
                user.is_bot_blocked = True
                user.blocked_bot_at = now
                await session.commit()
            
            await log_audit(session, user_id, u_name, "GÜVENLİK ALARMI: BOT ENGELLEME", f"Kullanıcı botu engelledi/sildi (Rol: {u_role}).")
            await session.commit()

            admins = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
            admin_tg_ids = set(ADMIN_IDS + [a.telegram_id for a in admins])
            for a_id in admin_tg_ids:
                adm_user = await session.get(User, a_id)
                a_lang = adm_user.language if adm_user else "tr"
                alert_txt = {
                    "tr": f"🚨 <b>KULLANICI BOTU ENGELLEDİ / SİLDİ</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Kullanıcı:</b> {escape_html(u_name)}\n🆔 <b>Telegram ID:</b> <code>{user_id}</code>\n🔗 <b>Kullanıcı Adı:</b> {u_username}\n🎭 <b>Sistem Rolü:</b> {u_role}\n🕒 <b>Zaman:</b> {get_local_now().strftime('%d.%m.%Y %H:%M')}\n━━━━━━━━━━━━━━━━━━━━",
                    "ru": f"🚨 <b>ПОЛЬЗОВАТЕЛЬ ЗАБЛОКИРОВАЛ БОТА / ВЫШЕЛ</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Пользователь:</b> {escape_html(u_name)}\n🆔 <b>Telegram ID:</b> <code>{user_id}</code>\n🔗 <b>Юзернейм:</b> {u_username}\n🎭 <b>Роль:</b> {u_role}\n🕒 <b>Время:</b> {get_local_now().strftime('%d.%m.%Y %H:%M')}\n━━━━━━━━━━━━━━━━━━━━",
                    "uz": f"🚨 <b>FOYDALANUVCHI BOTNI BLOKLADI / CHIQIB KETDI</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Foydalanuvchi:</b> {escape_html(u_name)}\n🆔 <b>Telegram ID:</b> <code>{user_id}</code>\n🔗 <b>Foydalanuvchi nomi:</b> {u_username}\n🎭 <b>Roli:</b> {u_role}\n🕒 <b>Vaqt:</b> {get_local_now().strftime('%d.%m.%Y %H:%M')}\n━━━━━━━━━━━━━━━━━━━━",
                    "en": f"🚨 <b>USER BLOCKED BOT / LEFT</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>User:</b> {escape_html(u_name)}\n🆔 <b>Telegram ID:</b> <code>{user_id}</code>\n🔗 <b>Username:</b> {u_username}\n🎭 <b>Role:</b> {u_role}\n🕒 <b>Time:</b> {get_local_now().strftime('%d.%m.%Y %H:%M')}\n━━━━━━━━━━━━━━━━━━━━"
                }.get(a_lang, "🚨 <b>USER BLOCKED BOT</b>")
                await safe_send_message(bot_obj, a_id, alert_txt, parse_mode="HTML")

        elif new_state in ("member",) and old_state in ("kicked", "left"):
            if user:
                user.is_bot_blocked = False
                user.blocked_bot_at = None
                await session.commit()
            await log_audit(session, user_id, u_name, "BOT ENGELİ KALKTI", f"Kullanıcı bot engelini kaldırdı (Rol: {u_role}).")
            await session.commit()


@router.message(Command("temizle", "clear", "sil"))
async def cmd_clear_chat(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    chat_id = message.chat.id

    # 1. Purge all bot messages
    await purge_previous_bot_messages(message.bot, chat_id)
    KEYBOARD_ANCHOR_MSG_ID.pop(chat_id, None)
    LAST_MENU_MSG_ID.pop(chat_id, None)

    # 2. Delete user's command message
    try: await message.delete()
    except Exception: pass

    # 3. Get user language
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    reset_txt = {
        "tr": "🧹 <b>Sohbet Temizlendi</b>\n━━━━━━━━━━━━━━━━━━━━\nSistemi ve menüyü yeniden başlatmak için aşağıdaki butona tıklayabilir veya <code>/start</code> yazabilirsiniz:",
        "ru": "🧹 <b>Чат очищен</b>\n━━━━━━━━━━━━━━━━━━━━\nНажмите кнопку ниже или введите <code>/start</code> для запуска меню:",
        "uz": "🧹 <b>Chat tozalandi</b>\n━━━━━━━━━━━━━━━━━━━━\nMenyuni ishga tushirish uchun quyidagi tugmani bosing yoki <code>/start</code> yozing:",
        "en": "🧹 <b>Chat Cleared</b>\n━━━━━━━━━━━━━━━━━━━━\nClick the button below or type <code>/start</code> to launch the menu:"
    }.get(lang, "🧹 <b>Sohbet Temizlendi</b>")

    start_btn_txt = {
        "tr": "🚀 Botu Başlat (START)",
        "ru": "🚀 Запустить бота (СТАРТ)",
        "uz": "🚀 Botni ishga tushirish (BOSHLASH)",
        "en": "🚀 Launch Bot (START)"
    }.get(lang, "🚀 Botu Başlat")

    ikb = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=start_btn_txt, callback_data="act_relaunch_start")]
    ])

    sent = await message.answer(reset_txt, reply_markup=ikb, parse_mode="HTML")
    if sent:
        LAST_MENU_MSG_ID[chat_id] = sent.message_id
        ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(sent.message_id)


@router.callback_query(F.data == "act_relaunch_start")
async def cb_relaunch_start(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user:
            await render_clean_dashboard(query, user)
        else:
            await cmd_start(query.message, state)
    await query.answer()

@router.message(any_state, CommandStart())
@router.message(any_state, Command("start", "menu", "baslat", "başlat"))
@router.message(any_state, F.text.lower().in_(["/start", "start", "/menu", "menu", "menü", "/menü", "başlat", "baslat"]))
async def cmd_start(message: Message, state: FSMContext | None = None):
    print(f"--> [CMD_START TETİKLENDİ] ID: {message.from_user.id} | İsim: {message.from_user.full_name} | Metin: '{message.text}'")
    if state is not None:
        try: await state.clear()
        except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        is_admin_id = message.from_user.id in ADMIN_IDS
        tg_user = message.from_user

        if not user:
            user = User(
                telegram_id=tg_user.id,
                role="admin" if is_admin_id else "guest",
                language="tr",
                full_name=tg_user.full_name,
                username=tg_user.username,
                admin_type="permanent" if is_admin_id else "none"
            )
            session.add(user)
            await session.commit()
            await safe_edit_or_answer(
                message,
                "🌍 Iltimos, tilni tanlang / Пожалуйста, выберите язык / Lütfen dil seçiniz / Please select language:",
                reply_markup=get_language_inline_kb()
            )
            return

        if is_admin_id:
            user.role = "admin"
            user.admin_type = "permanent"
        if tg_user.username and user.username != tg_user.username:
            user.username = tg_user.username
        if tg_user.full_name:
            user.full_name = tg_user.full_name

        now = datetime.utcnow()
        if user.admin_type == "temporary" and user.admin_until and user.admin_until <= now:
            user.role = user.previous_role or "guest"
            user.admin_type = "none"
            user.admin_until = None
        if user.locked_until and user.locked_until <= now:
            user.locked_until = None
            user.failed_attempts = 0

        await session.commit()

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language))
            return

        if user.locked_until and user.locked_until > now:
            rem_mins = max(1, int((user.locked_until - now).total_seconds() / 60))
            await message.answer(get_text("lock_countdown_msg", user.language, mins=rem_mins), parse_mode="HTML")
            return

        if user.role != "admin":
            maint = await session.get(SystemSetting, "maintenance_mode")
            if maint and maint.value == "true":
                await message.answer(get_text("maintenance_mode", user.language))
                return

        if user.role == "guest":
            await prompt_guest_screen(message, user, state)
            return

        # Render clean dashboard with permanent keyboard anchor and single active card
        await render_clean_dashboard(message, user)

async def prompt_guest_screen(target: Message | CallbackQuery, user: User, state: FSMContext):
    lang = user.language
    text = get_text("welcome_guest", lang)
    reply_kb = get_role_reply_kb("guest", lang)
    await safe_edit_or_answer(target, text, reply_markup=reply_kb, parse_mode="HTML")

@router.callback_query(F.data.startswith("set_lang:"))
async def cb_set_lang(query: CallbackQuery, state: FSMContext):
    await state.clear()
    lang_code = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if not user:
            is_admin_id = query.from_user.id in ADMIN_IDS
            user = User(telegram_id=query.from_user.id, language=lang_code, role="admin" if is_admin_id else "guest")
            session.add(user)
        else:
            user.language = lang_code
        await session.commit()

    toast_msg = get_text("lang_changed", user.language)
    await query.answer(toast_msg, show_alert=False)

    bot_obj = query.bot if getattr(query, 'bot', None) else (query.message.bot if query.message else bot)
    target_chat_id = query.message.chat.id if query.message else query.from_user.id

    reply_kb = get_role_reply_kb(user.role, user.language)
    if user.role == "guest":
        text = get_text("welcome_guest", user.language)
        sent_m = await bot_obj.send_message(chat_id=target_chat_id, text=text, reply_markup=reply_kb, parse_mode="HTML")
        if sent_m:
            LAST_MENU_MSG_ID[target_chat_id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(target_chat_id, set()).add(sent_m.message_id)
            await purge_previous_bot_messages(bot_obj, target_chat_id, keep_msg_id=sent_m.message_id)
    else:
        dash_text = await get_dashboard_card_text(user)
        sent_dash = await bot_obj.send_message(chat_id=target_chat_id, text=dash_text, reply_markup=reply_kb, parse_mode="HTML")
        if sent_dash:
            LAST_MENU_MSG_ID[target_chat_id] = sent_dash.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(target_chat_id, set()).add(sent_dash.message_id)
            await purge_previous_bot_messages(bot_obj, target_chat_id, keep_msg_id=sent_dash.message_id)

    if query.message:
        try: await query.message.delete()
        except Exception: pass

async def get_dashboard_card_text(user: User) -> str:
    lang = user.language
    async with AsyncSessionLocal() as session:
        sn_obj = await session.get(SystemSetting, "school_name")
        school_name = (sn_obj.value or "").strip() if sn_obj else ""
        s_header = f"🏛️ <b>{escape_html(school_name.upper())}</b>\n" if school_name else ""
        c_cnt, s_cnt, t_cnt, med_cnt, req_cnt = 0, 0, 0, 0, 0
        name, cls_name, num_val, subj = "-", "-", "-", "Ders"

        if user.role == "admin":
            c_cnt = (await session.execute(select(func.count(Student.class_name.distinct())))).scalar() or 0
            s_cnt = (await session.execute(select(func.count(Student.id)))).scalar() or 0
            t_cnt = (await session.execute(select(func.count(Teacher.id)))).scalar() or 0
            med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0
            req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
        elif user.role == "parent":
            st = await session.get(Student, user.current_child_id) if user.current_child_id else None
            name = escape_md(st.full_name) if st else "-"
            cls_name = escape_md(st.class_name) if st else "-"
        elif user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            name = escape_md(st.full_name) if st else (escape_md(user.full_name or "Öğrenci"))
            cls_name = escape_md(st.class_name) if st else "-"
            num_val = escape_md(st.student_number) if st else "-"
        elif user.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
            raw_s = tch.subject if (tch and tch.subject and tch.subject.strip()) else ""
            if not raw_s or raw_s.lower() in ("genel", "nothing", "none", "-"):
                subj = {"tr": "Genel", "ru": "Общий", "uz": "Umumiy", "en": "General"}.get(lang, "General")
            else:
                subj = raw_s
            name = escape_md(tch.full_name) if (tch and tch.full_name) else (escape_md(user.full_name or get_role_label("teacher", lang)))

    date_str = get_local_date().strftime('%d.%m.%Y')
    w_greet = f"👋 <i>{escape_html(user.full_name or 'Kullanıcı')}</i>\n" if user.full_name else ""

    if user.role == "admin":
        t_adm_title = s_header + {"tr": "⚡ <b>OKUL YÖNETİM KOKPİTİ (ADMİN)</b>", "ru": "⚡ <b>ПАНЕЛЬ УПРАВЛЕНИЯ ШКОЛОЙ (АДМИН)</b>", "uz": "⚡ <b>MAKTAB BOSHQARUV MARKAZI (ADMIN)</b>", "en": "⚡ <b>SCHOOL ADMINISTRATION COCKPIT (ADMIN)</b>"}.get(lang, "⚡ <b>SCHOOL ADMINISTRATION COCKPIT (ADMIN)</b>")
        t_adm_sec1 = {"tr": "📊 <b>GENEL OKUL DURUMU</b>", "ru": "📊 <b>ОБЩИЙ СТАТУС ШКОЛЫ</b>", "uz": "📊 <b>UMUMIY MAKTAB HOLATI</b>", "en": "📊 <b>GENERAL SCHOOL STATUS</b>"}.get(lang, "📊 <b>GENERAL SCHOOL STATUS</b>")
        t_adm_sec2 = {"tr": "🛎️ <b>BEKLEYEN İŞLEMLER</b>", "ru": "🛎️ <b>ОЖИДАЮЩИЕ ЗАЯВКИ</b>", "uz": "🛎️ <b>KUTILAYOTGAN AMALLAR</b>", "en": "🛎️ <b>PENDING ACTIONS</b>"}.get(lang, "🛎️ <b>PENDING ACTIONS</b>")
        lbl_c = {"tr": "Sınıflar", "ru": "Классы", "uz": "Sinflar", "en": "Classes"}.get(lang, "Classes")
        lbl_s = {"tr": "Öğrenciler", "ru": "Ученики", "uz": "O'quvchilar", "en": "Students"}.get(lang, "Students")
        lbl_t = {"tr": "Öğretmenler", "ru": "Учителя", "uz": "O'qituvchilar", "en": "Teachers"}.get(lang, "Teachers")
        lbl_a = {"tr": "Yöneticiler", "ru": "Админы", "uz": "Ma'murlar", "en": "Admins"}.get(lang, "Admins")
        lbl_req = {"tr": "Erişim Başvuruları", "ru": "Заявки на доступ", "uz": "Kirish arizalari", "en": "Access Requests"}.get(lang, "Access Requests")
        lbl_med = {"tr": "Mazeret Raporları", "ru": "Медицинские справки", "uz": "Tibbiy ma'lumotnomalar", "en": "Medical Notes"}.get(lang, "Medical Notes")
        lbl_date = {"tr": "Tarih", "ru": "Дата", "uz": "Sana", "en": "Date"}.get(lang, "Date")

        return (
            f"{t_adm_title}\n"
            f"{w_greet}"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📅 <b>{lbl_date}:</b> {date_str}\n\n"
            f"{t_adm_sec1}\n"
            f"┌ 🏫 <b>{lbl_c}:</b> {c_cnt}      👥 <b>{lbl_s}:</b> {s_cnt}\n"
            f"└ 👨‍🏫 <b>{lbl_t}:</b> {t_cnt}  👑 <b>{lbl_a}:</b> {len(ADMIN_IDS)}\n\n"
            f"{t_adm_sec2}\n"
            f"┌ 📩 <b>{lbl_req}:</b> {req_cnt}\n"
            f"└ 🏥 <b>{lbl_med}:</b> {med_cnt}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )
    elif user.role == "teacher":
        t_title = {"tr": "👨‍🏫 <b>ÖĞRETMEN YÖNETİM MASASI</b>", "ru": "👨‍🏫 <b>ПАНЕЛЬ УЧИТЕЛЯ</b>", "uz": "👨‍🏫 <b>O'QITUVCHI BOSHQARUV PANELI</b>", "en": "👨‍🏫 <b>TEACHER DASHBOARD</b>"}.get(lang, "👨‍🏫 <b>TEACHER DASHBOARD</b>")
        lbl_t = {"tr": "Öğretmen", "ru": "Учитель", "uz": "O'qituvchi", "en": "Teacher"}.get(lang, "Teacher")
        lbl_d = {"tr": "Tarih", "ru": "Дата", "uz": "Sana", "en": "Date"}.get(lang, "Date")
        return (
            f"{t_title}\n"
            f"{w_greet}"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 <b>{lbl_t}:</b> {escape_html(name)} ({escape_html(subj)})\n"
            f"📅 <b>{lbl_d}:</b> {date_str}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )
    elif user.role == "parent":
        p_title = {"tr": "👨‍👩‍👧‍👦 <b>VELİ BİLGİLENDİRME MASASI</b>", "ru": "👨‍👩‍👧‍👦 <b>ПАНЕЛЬ РОДИТЕЛЯ</b>", "uz": "👨‍👩‍👧‍👦 <b>OTA-ONA PANELI</b>", "en": "👨‍👩‍👧‍👦 <b>PARENT DASHBOARD</b>"}.get(lang, "👨‍👩‍👧‍👦 <b>PARENT DASHBOARD</b>")
        lbl_s = {"tr": "Aktif Öğrenci", "ru": "Текущий ученик", "uz": "Faol o'quvchi", "en": "Active Student"}.get(lang, "Active Student")
        lbl_d = {"tr": "Tarih", "ru": "Дата", "uz": "Sana", "en": "Date"}.get(lang, "Date")

        other_children_str = ""
        kids = (await session.execute(select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == user.telegram_id))).scalars().all()
        if len(kids) > 1:
            lbl_other = {"tr": "Diğer Çocuklar", "ru": "Другие дети", "uz": "Boshqa farzandlar", "en": "Other Children"}.get(lang, "Other Children")
            others = [f"• {k.full_name} ({k.class_name})" for k in kids if k.id != user.current_child_id]
            if others:
                other_children_str = f"\n\n👥 <b>{lbl_other}:</b>\n" + "\n".join(others)

        return (
            f"{p_title}\n"
            f"{w_greet}"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🧑‍🎓 <b>{lbl_s}:</b> {escape_html(name)} ({escape_html(cls_name)})\n"
            f"📅 <b>{lbl_d}:</b> {date_str}"
            f"{other_children_str}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )
    elif user.role == "student":
        s_title = {"tr": "🎓 <b>ÖĞRENCİ PORTALI</b>", "ru": "🎓 <b>ПАНЕЛЬ УЧЕНИКА</b>", "uz": "🎓 <b>O'QUVCHI PORTALI</b>", "en": "🎓 <b>STUDENT PORTAL</b>"}.get(lang, "🎓 <b>STUDENT PORTAL</b>")
        lbl_s = {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"}.get(lang, "Student")
        lbl_d = {"tr": "Tarih", "ru": "Дата", "uz": "Sana", "en": "Date"}.get(lang, "Date")
        lbl_no = {"tr": "No", "ru": "№", "uz": "№", "en": "Roll"}.get(lang, "No")
        return (
            f"{s_title}\n"
            f"{w_greet}"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🧑‍🎓 <b>{lbl_s}:</b> {escape_html(name)} ({escape_html(cls_name)} - {lbl_no}: {escape_html(num_val)})\n"
            f"📅 <b>{lbl_d}:</b> {date_str}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )
    else:
        raw = get_text("welcome_guest", lang)
        return re.sub(r'\*([^\*]+)\*', r'<b>\1</b>', raw.replace("`", "<code>").replace("`", "</code>"))

async def render_clean_dashboard(target: Message | CallbackQuery | Bot, user: User, chat_id: int | None = None, reply_kb=None):
    lang = user.language
    text = await get_dashboard_card_text(user)
    target_reply_kb = reply_kb or get_role_reply_kb(user.role, lang)

    bot_obj = target if isinstance(target, Bot) else (target.bot if isinstance(target, Message) else (target.message.bot if target.message else bot))
    target_chat_id = chat_id or (target.chat.id if isinstance(target, Message) else (target.message.chat.id if isinstance(target, CallbackQuery) and target.message else user.telegram_id))

    if isinstance(target, Message):
        try: await target.delete()
        except Exception: pass

    # Deliver single clean operational card equipped with role reply keyboard
    # Zero anchor messages, zero redundant bubbles!
    sent_m = await bot_obj.send_message(
        chat_id=target_chat_id,
        text=text,
        reply_markup=target_reply_kb,
        parse_mode="HTML"
    )
    if sent_m:
        new_mid = sent_m.message_id
        LAST_MENU_MSG_ID[target_chat_id] = new_mid
        chat_msgs = ACTIVE_CHAT_MESSAGES.setdefault(target_chat_id, set())
        chat_msgs.add(new_mid)
        if len(chat_msgs) > 100:
            ACTIVE_CHAT_MESSAGES[target_chat_id] = set(sorted(list(chat_msgs))[-30:])
        await purge_previous_bot_messages(bot_obj, target_chat_id, keep_msg_id=new_mid)

    if isinstance(target, CallbackQuery) and target.message and target.message.message_id != LAST_MENU_MSG_ID.get(target_chat_id):
        try: await target.message.delete()
        except Exception: pass
async def process_auth_code_string(code: str, user_id: int, message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass

    # If this is a menu button or cancel, never treat it as an auth code attempt
    action = match_reply_button(code)
    if action:
        await state.clear()
        await global_reply_keyboard_router(message, state)
        return
    if is_universal_cancel_text(code) or str(code).startswith("/"):
        await state.clear()
        if str(code).startswith("/start"):
            await cmd_start(message, state)
        else:
            await cmd_cancel(message, state)
        return

    clean_code = normalize_code(code)
    auth_result = None

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            user = User(telegram_id=user_id, language="tr", role="guest")
            session.add(user)
            await session.commit()
        lang = user.language

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", lang))
            return

        now = datetime.utcnow()
        if user.locked_until and user.locked_until > now:
            rem_mins = max(1, int((user.locked_until - now).total_seconds() / 60))
            await message.answer(get_text("lock_countdown_msg", lang, mins=rem_mins), parse_mode="HTML")
            return
        elif user.locked_until and user.locked_until <= now:
            user.locked_until = None
            user.failed_attempts = 0

        # 1. ADMIN CODE
        if clean_code == ADMIN_CODE or user_id in ADMIN_IDS:
            user.role = "admin"
            user.admin_type = "permanent"
            user.failed_attempts = 0
            user.locked_until = None
            if message.from_user.full_name:
                user.full_name = message.from_user.full_name
            await session.commit()
            await state.clear()
            auth_result = ("admin", user.full_name or get_text("permanent_admin_title", lang), get_text("lbl_role_admin", lang), user)
        elif clean_code.startswith("ADM-"):
            adm_key = (await session.execute(select(AdminKey).where(func.upper(func.trim(AdminKey.code)) == clean_code, AdminKey.is_used == False))).scalar_one_or_none()
            if adm_key:
                adm_key.is_used = True
                user.role = "admin"
                user.failed_attempts = 0
                user.locked_until = None
                if message.from_user.full_name:
                    user.full_name = message.from_user.full_name
                await session.commit()
                await state.clear()
                auth_result = ("admin", user.full_name or get_text("permanent_admin_title", lang), get_text("lbl_role_admin", lang), user)

        # 2. TEACHER CODE
        if not auth_result:
            teacher = (await session.execute(select(Teacher).where(func.upper(func.trim(Teacher.auth_code)) == clean_code))).scalar_one_or_none()
            if teacher:
                if teacher.telegram_id is not None and teacher.telegram_id != user_id:
                    await message.answer(get_text("auth_code_already_linked", lang), parse_mode="HTML")
                    await state.clear()
                    return
                teacher.telegram_id = user_id
                teacher.is_code_burned = True
                user.role = "teacher"
                user.full_name = teacher.full_name
                user.failed_attempts = 0
                user.locked_until = None
                await session.commit()
                await state.clear()
                auth_result = ("teacher", teacher.full_name, f"{get_text('lbl_role_teacher', lang)} ({teacher.subject})", user)

        # 3. STUDENT CODE
        if not auth_result:
            student = (await session.execute(select(Student).where(func.upper(func.trim(Student.student_code)) == clean_code))).scalar_one_or_none()
            if student:
                if student.student_telegram_id is not None and student.student_telegram_id != user_id:
                    await message.answer(get_text("auth_code_already_linked", lang), parse_mode="HTML")
                    await state.clear()
                    return
                student.student_telegram_id = user_id
                student.is_student_code_burned = True
                user.role = "student"
                user.full_name = student.full_name
                user.failed_attempts = 0
                user.locked_until = None
                await session.commit()
                await state.clear()
                auth_result = ("student", student.full_name, f"{get_text('lbl_role_student', lang)} ({student.class_name})", user)

        # 4. PARENT CODE
        if not auth_result:
            p_student = (await session.execute(select(Student).where(func.upper(func.trim(Student.parent_code)) == clean_code))).scalar_one_or_none()
            if p_student:
                p_student.is_parent_code_burned = True
                user.role = "parent"
                user.failed_attempts = 0
                user.locked_until = None
                user.current_child_id = p_student.id
                rel = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == user_id, ParentStudent.student_id == p_student.id))).scalar_one_or_none()
                if not rel:
                    session.add(ParentStudent(parent_telegram_id=user_id, student_id=p_student.id))
                await session.commit()
                await state.clear()
                auth_result = ("parent", f"{p_student.full_name} ({get_text('lbl_role_parent', lang)})", f"{get_text('lbl_role_parent', lang)} ({p_student.class_name})", user)

        if not auth_result:
            user.failed_attempts += 1
            if user.failed_attempts >= 5:
                user.is_blacklisted = True
                await session.commit()
                await message.answer(get_text("auth_blacklisted", lang))
                await state.clear()
                return
            elif user.failed_attempts >= 3:
                user.locked_until = datetime.utcnow() + timedelta(hours=1)
                await session.commit()
                await message.answer(get_text("auth_locked", lang))
                await state.clear()
                return

            rem = 3 - user.failed_attempts
            await session.commit()
            await message.answer(get_text("auth_failed", lang, remaining=rem))
            return

    if auth_result:
        r_type, n_val, r_lbl, u_obj = auth_result
        reply_kb = get_role_reply_kb(r_type, u_obj.language)
        dash_text = await get_dashboard_card_text(u_obj)
        welcome_banner = get_text("auth_success", u_obj.language, name=escape_md(n_val), role=escape_md(r_lbl))
        full_text = f"{welcome_banner}\n──────────────\n{dash_text}"
        sent_m = await message.bot.send_message(
            chat_id=message.chat.id,
            text=full_text,
            reply_markup=reply_kb,
            parse_mode="HTML"
        )
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)
            await purge_previous_bot_messages(message.bot, message.chat.id, keep_msg_id=sent_m.message_id)

@router.message(Form.waiting_auth_code)
async def handle_auth_code_fsm(message: Message, state: FSMContext):
    await process_auth_code_string(message.text, message.from_user.id, message, state)

@router.callback_query(F.data == "act_enter_code")
async def cb_act_enter_code(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:dashboard")]])
    await safe_edit_or_answer(query, get_text("prompt_enter_code_direct", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_auth_code)
    await query.answer()

# ======================================================================
# 9. KATI VE DOĞRULAMALI 6 ADIMLI BAŞVURU SİHİRBAZI
# ======================================================================

@router.callback_query(F.data == "act_req_access")
async def cb_act_req_access(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        pending = (await session.execute(
            select(AccessRequest).where(AccessRequest.telegram_id == query.from_user.id, AccessRequest.status == "pending")
        )).scalar_one_or_none()
        if pending:
            await query.answer(get_text("req_already_pending", lang), show_alert=True)
            return

    role_kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=get_text("role_teacher_btn", lang)), KeyboardButton(text=get_text("role_parent_btn", lang))],
            [KeyboardButton(text=get_text("role_student_btn", lang)), KeyboardButton(text=get_text("rk_cancel_action", lang))]
        ],
        resize_keyboard=True,
        is_persistent=True
    )
    await query.message.answer(get_text("req_role_select", lang), reply_markup=role_kb, parse_mode="HTML")
    await state.set_state(Form.req_role)
    await query.answer()

@router.message(Form.req_role)
async def process_req_role(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    txt = message.text.strip().lower()
    role_code = None
    if any(k in txt for k in ["öğretmen", "учитель", "o'qituvchi", "teacher"]):
        role_code = "teacher"
    elif any(k in txt for k in ["veli", "родитель", "ota-ona", "parent"]):
        role_code = "parent"
    elif any(k in txt for k in ["öğrenci", "ученик", "o'quvchi", "student"]):
        role_code = "student"

    if not role_code:
        p_role_sel = {
            "tr": "⚠️ Lütfen aşağıdaki butonlardan bir rol seçiniz:",
            "ru": "⚠️ Пожалуйста, выберите роль с помощью кнопок ниже:",
            "uz": "⚠️ Iltimos, quyidagi tugmalar orqali rolni tanlang:",
            "en": "⚠️ Please select a role using the buttons below:"
        }.get(lang, "⚠️ Please select a role:")
        await message.answer(p_role_sel)
        return

    await state.update_data(role=role_code)
    cancel_kb = get_cancel_reply_kb(lang)
    await message.answer(get_text("req_name_prompt", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.req_name)

@router.message(Form.req_name)
async def process_req_name(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    name_val = message.text.strip()
    parts = name_val.split()
    if len(parts) < 2 or len(name_val) < 4 or any(c.isdigit() for c in name_val) or "..." in name_val:
        await message.answer(get_text("err_invalid_name_strict", lang))
        return

    await state.update_data(full_name=name_val)

    gender_kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=get_text("btn_gender_male", lang)), KeyboardButton(text=get_text("btn_gender_female", lang))],
            [KeyboardButton(text=get_text("rk_cancel_action", lang))]
        ],
        resize_keyboard=True,
        is_persistent=True
    )
    await message.answer(get_text("prompt_req_gender", lang), reply_markup=gender_kb, parse_mode="HTML")
    await state.set_state(Form.req_gender)

@router.message(Form.req_gender)
async def process_req_gender(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    raw_txt = (message.text or "").strip()
    clean_t = raw_txt.lower()
    gender = None

    # 1. Exact match against defined buttons across all languages
    for l in ["tr", "ru", "uz", "en"]:
        male_btn = get_text("btn_gender_male", l).strip()
        female_btn = get_text("btn_gender_female", l).strip()
        if raw_txt == male_btn:
            gender = "Erkek"
            break
        elif raw_txt == female_btn:
            gender = "Kadın"
            break

    # 2. Broad keyword matching (Uzbek, Turkish, Russian, English and emojis)
    if not gender:
        male_keys = ["erkek", "erkak", "мужской", "male", "o'g'il", "o‘g‘il", "o`g`il", "ogil", "boy", "👨", "👦"]
        female_keys = ["kadın", "kadin", "kız", "kiz", "женский", "ayol", "female", "qiz", "girl", "👩", "👧"]
        if any(k in clean_t for k in male_keys):
            gender = "Erkek"
        elif any(k in clean_t for k in female_keys):
            gender = "Kadın"

    if not gender:
        err_card = get_text("err_invalid_gender_strict", lang)
        gender_kb = ReplyKeyboardMarkup(
            keyboard=[
                [KeyboardButton(text=get_text("btn_gender_male", lang)), KeyboardButton(text=get_text("btn_gender_female", lang))],
                [KeyboardButton(text=get_text("rk_cancel_action", lang))]
            ],
            resize_keyboard=True,
            is_persistent=True
        )
        await safe_send_message(message.bot, message.chat.id, err_card, reply_markup=gender_kb, parse_mode="HTML")
        return

    await state.update_data(gender=gender)
    cancel_kb = get_cancel_reply_kb(lang)
    prompt_txt = get_text("prompt_req_birth_date", lang)
    await safe_send_message(message.bot, message.chat.id, prompt_txt, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.req_birth_date)

@router.message(Form.req_birth_date)
async def process_req_birth_date(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    date_str = message.text.strip()
    m = re.match(r"^(\d{1,2})[\.\/\-](\d{1,2})[\.\/\-](\d{4})$", date_str)
    if not m:
        await message.answer(get_text("err_invalid_birth_date_strict", lang))
        return

    day, month, year = int(m.group(1)), int(m.group(2)), int(m.group(3))
    try:
        b_date = datetime(year, month, day)
        if not (1940 <= year <= (datetime.utcnow().year - 3)):
            await message.answer(get_text("err_invalid_birth_date_strict", lang))
            return
    except Exception:
        await message.answer(get_text("err_invalid_birth_date_strict", lang))
        return

    clean_birth = f"{day:02d}.{month:02d}.{year}"
    await state.update_data(birth_date=clean_birth)

    phone_kb = ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text=get_text("btn_share_contact", lang), request_contact=True)],
            [KeyboardButton(text=get_text("rk_cancel_action", lang))]
        ],
        resize_keyboard=True,
        is_persistent=True
    )
    await message.answer(get_text("req_phone_prompt", lang), reply_markup=phone_kb, parse_mode="HTML")
    await state.set_state(Form.req_phone)

@router.message(Form.req_phone)
async def process_req_phone(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    if message.contact and message.contact.phone_number:
        phone_val = str(message.contact.phone_number)
    else:
        phone_val = message.text.strip()
        digits_only = re.sub(r"\D", "", phone_val)
        if len(digits_only) < 10:
            await message.answer(get_text("err_invalid_phone_strict", lang))
            return

    await state.update_data(phone=phone_val)
    data = await state.get_data()
    role = data.get("role", "student")

    prompt_key = "req_details_teacher" if role == "teacher" else ("req_details_parent" if role == "parent" else "req_details_student")
    await message.answer(get_text(prompt_key, lang), reply_markup=get_cancel_reply_kb(lang), parse_mode=None)
    await state.set_state(Form.req_details)

@router.message(Form.req_details)
async def process_req_details(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    details_text = message.text.strip()
    if len(details_text) < 2:
        await message.answer(get_text("err_invalid_details_strict", "tr"))
        return

    data = await state.get_data()
    await state.clear()

    role = data.get("role", "student")
    full_name = data.get("full_name", "")
    gender = data.get("gender", "Erkek")
    birth_date = data.get("birth_date", "")
    phone_val = data.get("phone", "")

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        student_match = None
        clean_det = details_text.lower()
        if role in ["parent", "student"]:
            all_st = (await session.execute(select(Student))).scalars().all()
            for s in all_st:
                if s.student_number in clean_det or s.full_name.lower() in clean_det:
                    student_match = s
                    break

        req = AccessRequest(
            telegram_id=message.from_user.id,
            role=role,
            full_name=full_name,
            gender=gender,
            birth_date=birth_date,
            phone=phone_val,
            details=details_text,
            student_match_id=student_match.id if student_match else None,
            status="pending"
        )
        session.add(req)
        await session.commit()

        st_sent = get_text("req_sent_success", lang)
        w_guest = get_text("welcome_guest", lang)
        combined_card = f"{st_sent}\n──────────────\n{w_guest}"
        guest_kb = get_role_reply_kb("guest", lang)
        await safe_edit_or_answer(message, combined_card, reply_markup=guest_kb, parse_mode="HTML")

        age_str = ""
        try:
            parts = birth_date.split(".")
            if len(parts) == 3:
                age_val = datetime.utcnow().year - int(parts[2])
                age_str = f" ({age_val} {get_text('lbl_age', 'tr')})"
        except Exception: pass

        admin_users = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
        admin_lang_map = {u.telegram_id: (u.language or "tr") for u in admin_users}
        for a_id in ADMIN_IDS:
            if a_id not in admin_lang_map:
                admin_lang_map[a_id] = "tr"

        for a_id, a_lang in admin_lang_map.items():
            r_lbl = {
                "student": {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"},
                "teacher": {"tr": "Öğretmen", "ru": "Учитель", "uz": "O'qituvchi", "en": "Teacher"},
                "parent": {"tr": "Veli", "ru": "Родитель", "uz": "Ota-ona", "en": "Parent"}
            }.get(role, {}).get(a_lang, role)

            auto_badge = {
                "tr": f"🟢 Sistem Eşleşmesi: {escape_html(student_match.full_name)} ({escape_html(student_match.class_name)} - No: {escape_html(student_match.student_number)})" if student_match else "🔴 Otomatik Eşleşme Yok",
                "ru": f"🟢 Найдено в базе: {escape_html(student_match.full_name)} ({escape_html(student_match.class_name)} - №: {escape_html(student_match.student_number)})" if student_match else "🔴 Автосовпадений не найдено",
                "uz": f"🟢 Tizimdan topildi: {escape_html(student_match.full_name)} ({escape_html(student_match.class_name)} - №: {escape_html(student_match.student_number)})" if student_match else "🔴 Avtomatik moslik topilmadi",
                "en": f"🟢 Matched in system: {escape_html(student_match.full_name)} ({escape_html(student_match.class_name)} - No: {escape_html(student_match.student_number)})" if student_match else "🔴 No Automatic Match Found"
            }.get(a_lang, "No match")

            adm_card = {
                "tr": (
                    f"🛎️ <b>YENİ ERİŞİM VE ŞİFRE BAŞVURUSU (#{req.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Ad Soyad:</b> {escape_html(full_name)}\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(phone_val)}</code>\n"
                    f"📋 <b>Talep Edilen Rol:</b> <b>{r_lbl}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"📝 <b>Açıklama:</b> {escape_html(details_text or '-')}\n\n"
                    f"🔍 <i>{auto_badge}</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "ru": (
                    f"🛎️ <b>НОВАЯ ЗАЯВКА НА ДОСТУП И ПАРОЛЬ (#{req.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>ФИО:</b> {escape_html(full_name)}\n"
                    f"📱 <b>Телефон:</b> <code>{escape_html(phone_val)}</code>\n"
                    f"📋 <b>Роль:</b> <b>{r_lbl}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"📝 <b>Примечание:</b> {escape_html(details_text or '-')}\n\n"
                    f"🔍 <i>{auto_badge}</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "uz": (
                    f"🛎️ <b>YANGI KIRISH VA PAROL ARIZASI (#{req.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>F.I.O:</b> {escape_html(full_name)}\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(phone_val)}</code>\n"
                    f"📋 <b>So'ralgan lavozim:</b> <b>{r_lbl}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"📝 <b>Izoh:</b> {escape_html(details_text or '-')}\n\n"
                    f"🔍 <i>{auto_badge}</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "en": (
                    f"🛎️ <b>NEW ACCESS & PASSWORD REQUEST (#{req.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Full Name:</b> {escape_html(full_name)}\n"
                    f"📱 <b>Phone:</b> <code>{escape_html(phone_val)}</code>\n"
                    f"📋 <b>Requested Role:</b> <b>{r_lbl}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"📝 <b>Note:</b> {escape_html(details_text or '-')}\n\n"
                    f"🔍 <i>{auto_badge}</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                )
            }.get(a_lang, "New Request")

            adm_kb = InlineKeyboardMarkup(inline_keyboard=[
                [
                    InlineKeyboardButton(text=get_text("btn_appr_request", a_lang), callback_data=f"adm:appr_req:{req.id}"),
                    InlineKeyboardButton(text=get_text("btn_reject", a_lang), callback_data=f"adm:rej_req:{req.id}")
                ]
            ])
            s_m = await safe_send_message(message.bot, a_id, adm_card, reply_markup=adm_kb, parse_mode="HTML")
            if s_m:
                ADMIN_DISPATCHED_NOTIFS.setdefault(f"req:{req.id}", []).append((a_id, s_m.message_id))

@router.callback_query(F.data == "adm:requests_list")
async def cb_admin_requests_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        requests = (await session.execute(select(AccessRequest).where(AccessRequest.status == "pending").order_by(desc(AccessRequest.created_at)))).scalars().all()
        if not requests:
            buttons = [get_nav_buttons(lang, back_callback="adm:cat_requests")]
            await safe_edit_or_answer(query, get_text("no_pending_requests", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode=None)
            await query.answer()
            return

        buttons = []
        for r in requests:
            icon = "👨‍🏫" if r.role == "teacher" else ("👨‍👩‍👧‍👦" if r.role == "parent" else "🎓")
            buttons.append([InlineKeyboardButton(text=f"{icon} {r.full_name} ({r.created_at.strftime('%H:%M')})", callback_data=f"adm:view_req:{r.id}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_requests"))

        await safe_edit_or_answer(query, get_text("pending_requests_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode=None)
    await query.answer()

@router.callback_query(F.data.startswith("adm:view_req:"))
async def cb_admin_view_request(query: CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer(get_text("request_already_handled", lang), show_alert=True)
            if query.message:
                try: await query.message.delete()
                except Exception:
                    try: await query.message.edit_reply_markup(reply_markup=None)
                    except Exception: pass
            await cb_admin_requests_list(query)
            return

        role_label = {
            "teacher": get_text("lbl_role_teacher", lang),
            "parent": get_text("lbl_role_parent", lang),
            "student": get_text("lbl_role_student", lang)
        }.get(req.role, req.role)

        st_match_info = ""
        if req.student_match_id:
            st = await session.get(Student, req.student_match_id)
            if st:
                match_lbl = {"tr": "Eşleşen Öğrenci", "ru": "Найден ученик", "uz": "Mos o'quvchi", "en": "Matched Student"}.get(lang, "Matched Student")
                st_match_info = f"\n🔍 {match_lbl}: {escape_html(st.full_name)} ({escape_html(st.class_name)} - No: {escape_html(st.student_number)})"

        age_str = ""
        if req.birth_date:
            try:
                parts = req.birth_date.split(".")
                if len(parts) == 3:
                    age_val = datetime.utcnow().year - int(parts[2])
                    age_str = f" ({age_val} {get_text('lbl_age', lang)})"
            except Exception: pass

        t_header = {
            "tr": f"🛎️ <b>Erişim & Şifre Başvurusu (#{req.id})</b>\n━━━━━━━━━━━━━━━━━━━━",
            "ru": f"🛎️ <b>Заявка на доступ (#{req.id})</b>\n━━━━━━━━━━━━━━━━━━━━",
            "uz": f"🛎️ <b>Kirish va parol arizasi (#{req.id})</b>\n━━━━━━━━━━━━━━━━━━━━",
            "en": f"🛎️ <b>Access & Password Request (#{req.id})</b>\n━━━━━━━━━━━━━━━━━━━━"
        }.get(lang, "🛎️ <b>Access Request</b>\n━━━━━━━━━━━━━━━━━━━━")

        prompt_choose = {
            "tr": "Lütfen yapılacak işlemi seçiniz:",
            "ru": "Выберите действие:",
            "uz": "Iltimos, amalni tanlang:",
            "en": "Please select action:"
        }.get(lang, "Select action:")

        lbl_details = {"tr": "Açıklama / Detay", "ru": "Описание", "uz": "Izoh", "en": "Details"}.get(lang, "Details")

        text = (
            f"{t_header}\n"
            f"• <b>{get_text('lbl_full_name', lang)}:</b> {escape_html(req.full_name)}\n"
            f"• <b>{get_text('lbl_gender', lang)}:</b> {req.gender or '-'}\n"
            f"• <b>{get_text('lbl_birth_date', lang)}:</b> {req.birth_date or '-'}{age_str}\n"
            f"• <b>{get_text('lbl_phone', lang)}:</b> {req.phone}\n"
            f"• <b>{get_text('lbl_role', lang)}:</b> {role_label}\n"
            f"• <b>{lbl_details}:</b> {escape_html(req.details)}{st_match_info}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{prompt_choose}"
        )
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_appr_request", lang), callback_data=f"adm:appr_req:{req.id}"), InlineKeyboardButton(text=get_text("btn_reject", lang), callback_data=f"adm:rej_req:{req.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:requests_list")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:appr_req:"))
async def cb_admin_approve_request(query: CallbackQuery, state: FSMContext):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        lang = admin_user.language if admin_user else "tr"
        if not is_admin_user(admin_user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer(get_text("request_already_handled", lang), show_alert=True)
            if query.message:
                try: await query.message.delete()
                except Exception:
                    try: await query.message.edit_reply_markup(reply_markup=None)
                    except Exception: pass
            return

        if req.role == "student":
            st = await session.get(Student, req.student_match_id) if req.student_match_id else None
            if not st:
                st = (await session.execute(select(Student).where(func.lower(func.trim(Student.full_name)) == req.full_name.strip().lower()))).scalar_one_or_none()

            if not st:
                await state.update_data(appr_req_id=req.id, appr_st_name=req.full_name, appr_st_tg_id=req.telegram_id, appr_st_phone=req.phone)
                await state.set_state(Form.appr_st_class)

                existing_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
                cls_buttons = []
                row = []
                for c in existing_classes:
                    row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:appr_cls:{c}"))
                    if len(row) == 3:
                        cls_buttons.append(row)
                        row = []
                if row: cls_buttons.append(row)
                cls_buttons.append([InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:view_req:{req.id}")])

                prompt_c = {
                    "tr": f"🎓 <b>Öğrenci Kayıt Onayı (#{req.id})</b>\n──────────────\n👤 <b>Öğrenci:</b> {escape_html(req.full_name)}\n📱 <b>Telefon:</b> {escape_html(req.phone or '-')}\n\nLütfen öğrencinin atanacağı <b>Sınıfı</b> seçiniz veya yazınız (Örn: <code>9-A</code>):",
                    "ru": f"🎓 <b>Одобрение регистрации ученика (#{req.id})</b>\n──────────────\n👤 <b>Ученик:</b> {escape_html(req.full_name)}\n📱 <b>Телефон:</b> {escape_html(req.phone or '-')}\n\nВыберите <b>Класс</b> или напишите сообщением (Напр: <code>9-A</code>):",
                    "uz": f"🎓 <b>O'quvchini ro'yxatga olishni tasdiqlash (#{req.id})</b>\n──────────────\n👤 <b>O'quvchi:</b> {escape_html(req.full_name)}\n📱 <b>Telefon:</b> {escape_html(req.phone or '-')}\n\nIltimos, o'quvchi biriktiriladigan <b>Sinfni</b> tanlang yoki yozing (Masalan: <code>9-A</code>):",
                    "en": f"🎓 <b>Student Registration Approval (#{req.id})</b>\n──────────────\n👤 <b>Student:</b> {escape_html(req.full_name)}\n📱 <b>Phone:</b> {escape_html(req.phone or '-')}\n\nPlease select the student's <b>Class</b> or type it (e.g. <code>9-A</code>):"
                }.get(lang, "Select Class:")

                await safe_edit_or_answer(query, prompt_c, reply_markup=InlineKeyboardMarkup(inline_keyboard=cls_buttons), parse_mode="HTML")
                await query.answer()
                return
            else:
                st.student_telegram_id = req.telegram_id
                st.is_student_code_burned = True
                req.student_match_id = st.id

        req.status = "approved"
        req.reviewed_by = query.from_user.id
        req.reviewed_at = datetime.utcnow()

        target_user = await session.get(User, req.telegram_id)
        if not target_user:
            target_user = User(telegram_id=req.telegram_id, language="tr")
            session.add(target_user)

        target_user.role = req.role
        target_user.full_name = req.full_name
        target_user.phone = req.phone
        target_user.failed_attempts = 0
        target_user.locked_until = None

        adm_card = ""
        extra_btn = None

        if req.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.full_name == req.full_name))).scalar_one_or_none()
            if not tch:
                tch = Teacher(full_name=req.full_name, subject=(req.details or "Genel"), auth_code=generate_secure_code("HCA"), telegram_id=req.telegram_id, is_code_burned=True, assigned_classes="ALL")
                session.add(tch)
            else:
                tch.telegram_id = req.telegram_id
                tch.is_code_burned = True
                if req.details: tch.subject = req.details
            await session.flush()

            adm_card = {
                "tr": (
                    f"✅ <b>ÖĞRETMEN BAŞVURUSU ONAYLANDI (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Yetki:</b> 👨‍🏫 Öğretmen\n"
                    f"👤 <b>Adı Soyadı:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"📚 <b>Branş / Ders:</b> <code>{escape_html(tch.subject or 'Genel')}</code>\n"
                    f"🏫 <b>Atanan Sınıflar:</b> <code>{escape_html(tch.assigned_classes or 'Tüm Sınıflar')}</code>\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Öğretmen Giriş Kodu:</b> <code>{tch.auth_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Kullanıcı 'Öğretmen' olarak onaylandı. Aşağıdaki butona tıklayarak giriş şifresini ve kodunu doğrudan kullanıcıya gönderebilirsiniz.</i>"
                ),
                "ru": (
                    f"✅ <b>ЗАЯВКА УЧИТЕЛЯ ОДОБРЕНА (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Роль:</b> 👨‍🏫 Учитель\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"📚 <b>Предмет:</b> <code>{escape_html(tch.subject or 'Общий')}</code>\n"
                    f"🏫 <b>Классы:</b> <code>{escape_html(tch.assigned_classes or 'Все классы')}</code>\n"
                    f"📱 <b>Телефон:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Код доступа:</b> <code>{tch.auth_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Пользователь одобрен как Учитель. Нажмите кнопку ниже для отправки пароля пользователю.</i>"
                ),
                "uz": (
                    f"✅ <b>O'QITUVCHI ARIZASI TASDIQLANDI (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Lavozim:</b> 👨‍🏫 O'qituvchi\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"📚 <b>Fani:</b> <code>{escape_html(tch.subject or 'Umumiy')}</code>\n"
                    f"🏫 <b>Biriktirilgan sinflar:</b> <code>{escape_html(tch.assigned_classes or 'Barcha sinflar')}</code>\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Kirish kodi:</b> <code>{tch.auth_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Foydalanuvchi O'qituvchi sifatida tasdiqlandi. Quyidagi tugma orqali parolni foydalanuvchiga yuborishingiz mumkin.</i>"
                ),
                "en": (
                    f"✅ <b>TEACHER REQUEST APPROVED (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Role:</b> 👨‍🏫 Teacher\n"
                    f"👤 <b>Name:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"📚 <b>Subject:</b> <code>{escape_html(tch.subject or 'General')}</code>\n"
                    f"🏫 <b>Assigned Classes:</b> <code>{escape_html(tch.assigned_classes or 'All Classes')}</code>\n"
                    f"📱 <b>Phone:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Access Code:</b> <code>{tch.auth_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>User approved as Teacher. Click the button below to send credentials to the user.</i>"
                )
            }.get(lang, "Teacher request approved.")

        elif req.role == "student":
            st = await session.get(Student, req.student_match_id) if req.student_match_id else None
            st_cls = st.class_name if st else "-"
            st_no = st.student_number if st else "-"
            st_code = st.student_code if st else "-"
            pr_code = st.parent_code if st else "-"

            adm_card = {
                "tr": (
                    f"✅ <b>ÖĞRENCİ BAŞVURUSU ONAYLANDI (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Yetki:</b> 🎓 Öğrenci\n"
                    f"👤 <b>Adı Soyadı:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🏫 <b>Sınıfı:</b> <code>{escape_html(st_cls)}</code>\n"
                    f"🔢 <b>Okul Numarası:</b> <code>{escape_html(st_no)}</code>\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Öğrenci Kodu:</b> <code>{st_code}</code>\n"
                    f"👨‍👩‍👧 <b>Veli Kodu:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Öğrenci kaydı eşleştirildi ve onaylandı.</i>"
                ),
                "ru": (
                    f"✅ <b>ЗАЯВКА УЧЕНИКА ОДОБРЕНА (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Роль:</b> 🎓 Ученик\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🏫 <b>Класс:</b> <code>{escape_html(st_cls)}</code>\n"
                    f"🔢 <b>Номер:</b> <code>{escape_html(st_no)}</code>\n"
                    f"📱 <b>Телефон:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Код ученика:</b> <code>{st_code}</code>\n"
                    f"🔑 <b>Код родителя:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Запись ученика привязана и подтверждена.</i>"
                ),
                "uz": (
                    f"✅ <b>O'QUVCHI ARIZASI TASDIQLANDI (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Lavozim:</b> 🎓 O'quvchi\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🏫 <b>Sinfi:</b> <code>{escape_html(st_cls)}</code>\n"
                    f"🔢 <b>Maktab raqami:</b> <code>{escape_html(st_no)}</code>\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>O'quvchi kodi:</b> <code>{st_code}</code>\n"
                    f"🔑 <b>Ota-ona kodi:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>O'quvchi qaydi biriktirildi va tasdiqlandi.</i>"
                ),
                "en": (
                    f"✅ <b>STUDENT REQUEST APPROVED (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Role:</b> 🎓 Student\n"
                    f"👤 <b>Name:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🏫 <b>Class:</b> <code>{escape_html(st_cls)}</code>\n"
                    f"🔢 <b>Roll:</b> <code>{escape_html(st_no)}</code>\n"
                    f"📱 <b>Phone:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Student Code:</b> <code>{st_code}</code>\n"
                    f"🔑 <b>Parent Code:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Student record matched and approved.</i>"
                )
            }.get(lang, "Student request approved.")

            if st and st.class_name:
                btn_view_cls = {"tr": f"🏫 {st.class_name} Sınıfını Gör", "ru": f"🏫 Класс {st.class_name}", "uz": f"🏫 {st.class_name} sinfi", "en": f"🏫 Class {st.class_name}"}.get(lang, f"🏫 {st.class_name}")
                extra_btn = InlineKeyboardButton(text=btn_view_cls, callback_data=f"adm:show_class:{st.class_name}")

        elif req.role == "parent":
            st_id = req.student_match_id
            if not st_id:
                all_st = (await session.execute(select(Student))).scalars().all()
                for s in all_st:
                    if s.full_name.lower() in req.details.lower():
                        st_id = s.id
                        break
            st = await session.get(Student, st_id) if st_id else None
            st_name = st.full_name if st else "-"
            st_cls = st.class_name if st else "-"
            st_no = st.student_number if st else "-"
            pr_code = st.parent_code if st else "-"

            if st_id:
                target_user.current_child_id = st_id
                rel = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == req.telegram_id, ParentStudent.student_id == st_id))).scalar_one_or_none()
                if not rel:
                    session.add(ParentStudent(parent_telegram_id=req.telegram_id, student_id=st_id))
            await session.flush()

            adm_card = {
                "tr": (
                    f"✅ <b>VELİ BAŞVURUSU ONAYLANDI (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Yetki:</b> 👨‍👩‍👧 Veli\n"
                    f"👤 <b>Adı Soyadı:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🧑‍🎓 <b>Bağlanan Öğrenci:</b> <b>{escape_html(st_name)}</b> (<code>{escape_html(st_cls)}</code> - No: <code>{escape_html(st_no)}</code>)\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Veli Giriş Kodu:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Veli hesabı öğrenciyle eşleştirildi. Şifreyi kullanıcıya iletmek için aşağıdaki butonu kullanabilirsiniz.</i>"
                ),
                "ru": (
                    f"✅ <b>ЗАЯВКА РОДИТЕЛЯ ОДОБРЕНА (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Роль:</b> 👨‍👩‍👧 Родитель\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🧑‍🎓 <b>Ученик:</b> <b>{escape_html(st_name)}</b> (<code>{escape_html(st_cls)}</code> - № <code>{escape_html(st_no)}</code>)\n"
                    f"📱 <b>Телефон:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Код родителя:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Родитель успешно привязан к ученику. Нажмите кнопку ниже для отправки доступа.</i>"
                ),
                "uz": (
                    f"✅ <b>OTA-ONA ARIZASI TASDIQLANDI (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Lavozim:</b> 👨‍👩‍👧 Ota-ona\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🧑‍🎓 <b>Biriktirilgan o'quvchi:</b> <b>{escape_html(st_name)}</b> (<code>{escape_html(st_cls)}</code> - № <code>{escape_html(st_no)}</code>)\n"
                    f"📱 <b>Telefon:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Ota-ona kodi:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Ota-ona profili o'quvchiga muvaffaqiyatli biriktirildi.</i>"
                ),
                "en": (
                    f"✅ <b>PARENT REQUEST APPROVED (#{req.id})</b>\n"
                    "──────────────\n"
                    f"📋 <b>Role:</b> 👨‍👩‍👧 Parent\n"
                    f"👤 <b>Name:</b> <b>{escape_html(req.full_name)}</b>\n"
                    f"🧑‍🎓 <b>Linked Student:</b> <b>{escape_html(st_name)}</b> (<code>{escape_html(st_cls)}</code> - No: <code>{escape_html(st_no)}</code>)\n"
                    f"📱 <b>Phone:</b> <code>{escape_html(req.phone or '-')}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{req.telegram_id}</code>\n"
                    f"🔑 <b>Parent Code:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Parent profile successfully linked to student.</i>"
                )
            }.get(lang, "Parent request approved.")

        await session.commit()

        await sync_admin_notif_resolution(
            query.message.bot,
            f"req:{req.id}",
            query.from_user.id,
            (admin_user.full_name if admin_user else "Yönetici"),
            "approved",
            f"Başvuru #{req.id}: {req.full_name} ({req.role})"
        )

        u_lang = target_user.language or "tr"
        dash_t = await get_dashboard_card_text(target_user)
        appr_b = get_text("req_approved_user", u_lang, role=req.role)
        full_u_msg = f"{appr_b}\n──────────────\n{dash_t}"
        u_kb = get_role_reply_kb(req.role, u_lang)
        old_u_mid = LAST_MENU_MSG_ID.get(req.telegram_id)
        s_u = await safe_send_message(query.message.bot, req.telegram_id, full_u_msg, reply_markup=u_kb, parse_mode="HTML")
        if s_u:
            LAST_MENU_MSG_ID[req.telegram_id] = s_u.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(req.telegram_id, set()).add(s_u.message_id)
            if old_u_mid and old_u_mid != s_u.message_id:
                try: await query.message.bot.delete_message(chat_id=req.telegram_id, message_id=old_u_mid)
                except Exception: pass

        btn_send = {"tr": "📲 Kullanıcıya Şifreyi Gönder", "ru": "📲 Отправить пароль пользователю", "uz": "📲 Foydalanuvchiga parolni yuborish", "en": "📲 Send Credentials to User"}.get(lang, "Send Credentials")
        buttons = [[InlineKeyboardButton(text=btn_send, callback_data=f"adm:send_creds:{req.id}")]]
        if extra_btn:
            buttons.append([extra_btn])
        buttons.append([
            InlineKeyboardButton(text=get_text("btn_requests", lang).split("(")[0].strip(), callback_data="adm:requests_list")
        ])

        await safe_edit_or_answer(query, adm_card, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer(get_text("acknowledged_toast", lang))


@router.callback_query(F.data.startswith("adm:send_creds:"))
async def cb_admin_send_credentials(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        lang = admin_user.language if admin_user else "tr"
        if not is_admin_user(admin_user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        req = await session.get(AccessRequest, req_id)
        if not req:
            await query.answer("⚠️ Başvuru bulunamadı." if lang == "tr" else "Not found", show_alert=True)
            return

        target_user = await session.get(User, req.telegram_id)
        u_lang = target_user.language if target_user else "tr"

        if req.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == req.telegram_id))).scalar_one_or_none()
            if not tch:
                tch = (await session.execute(select(Teacher).where(Teacher.full_name == req.full_name))).scalar_one_or_none()
            code = tch.auth_code if tch else "-"
            subj = tch.subject if tch else (req.details or "Genel")
            
            cred_msg = {
                "tr": (
                    f"🎉 <b>Sayın {escape_html(req.full_name)},</b>\n\n"
                    "Okul idaresi başvurunuzu onayladı ve sisteme erişim yetkiniz tanımlandı.\n\n"
                    "📋 <b>Giriş Bilgileriniz:</b>\n"
                    f"• <b>Rolünüz:</b> 👨‍🏫 Öğretmen\n"
                    f"• <b>Branşınız:</b> {escape_html(subj)}\n"
                    f"• <b>Özel Giriş Şifreniz / Kodunuz:</b> <code>{code}</code>\n\n"
                    "<i>Aşağıdaki menüden dilediğiniz işlemi başlatabilirsiniz. İyi dersler dileriz!</i>"
                ),
                "ru": (
                    f"🎉 <b>Уважаемый(ая) {escape_html(req.full_name)},</b>\n\n"
                    "Администрация школы одобрила вашу заявку и предоставила доступ к системе.\n\n"
                    "📋 <b>Данные для входа:</b>\n"
                    f"• <b>Должность:</b> 👨‍🏫 Учитель\n"
                    f"• <b>Предмет:</b> {escape_html(subj)}\n"
                    f"• <b>Ваш код доступа (пароль):</b> <code>{code}</code>\n\n"
                    "<i>Используйте меню ниже для работы с ботом. Желаем успешных занятий!</i>"
                ),
                "uz": (
                    f"🎉 <b>Hurmatli {escape_html(req.full_name)},</b>\n\n"
                    "Maktab ma'muriyati arizangizni tasdiqladi va tizimga kirish ruxsatini berdi.\n\n"
                    "📋 <b>Kirish ma'lumotlaringiz:</b>\n"
                    f"• <b>Lavozimingiz:</b> 👨‍🏫 O'qituvchi\n"
                    f"• <b>Faningiz:</b> {escape_html(subj)}\n"
                    f"• <b>Maxsus kirish kodingiz (parol):</b> <code>{code}</code>\n\n"
                    "<i>Quyidagi menyu orqali faoliyatni boshlashingiz mumkin. Omadli darslar!</i>"
                ),
                "en": (
                    f"🎉 <b>Dear {escape_html(req.full_name)},</b>\n\n"
                    "The school administration has approved your request and granted system access.\n\n"
                    "📋 <b>Your Credentials:</b>\n"
                    f"• <b>Role:</b> 👨‍🏫 Teacher\n"
                    f"• <b>Subject:</b> {escape_html(subj)}\n"
                    f"• <b>Access Code / Password:</b> <code>{code}</code>\n\n"
                    "<i>You can use the menu below to navigate the bot. Have a great teaching session!</i>"
                )
            }.get(u_lang, "Access approved.")

        elif req.role == "student":
            st = await session.get(Student, req.student_match_id) if req.student_match_id else None
            if not st:
                st = (await session.execute(select(Student).where(Student.student_telegram_id == req.telegram_id))).scalar_one_or_none()
            st_code = st.student_code if st else "-"
            pr_code = st.parent_code if st else "-"
            cls_name = st.class_name if st else "-"
            st_no = st.student_number if st else "-"

            cred_msg = {
                "tr": (
                    f"🎉 <b>Sayın {escape_html(req.full_name)},</b>\n\n"
                    "Okul idaresi öğrenci kaydınızı ve başvurunuzu onayladı!\n\n"
                    "📋 <b>Öğrenci Bilgileriniz:</b>\n"
                    f"• <b>Sınıfınız:</b> <code>{cls_name}</code>\n"
                    f"• <b>Okul Numaranız:</b> <code>{st_no}</code>\n"
                    f"• <b>Öğrenci Giriş Kodunuz:</b> <code>{st_code}</code>\n"
                    f"• <b>Veli Bağlantı Kodunuz:</b> <code>{pr_code}</code>\n\n"
                    "<i>Aşağıdaki menüyü kullanarak notlarınızı, ödevlerinizi ve ders durumunuzu takip edebilirsiniz. Başarılar dileriz!</i>"
                ),
                "ru": (
                    f"🎉 <b>Уважаемый(ая) {escape_html(req.full_name)},</b>\n\n"
                    "Администрация школы одобрила вашу регистрацию в качестве ученика!\n\n"
                    "📋 <b>Данные ученика:</b>\n"
                    f"• <b>Класс:</b> <code>{cls_name}</code>\n"
                    f"• <b>Номер в школе:</b> <code>{st_no}</code>\n"
                    f"• <b>Код ученика:</b> <code>{st_code}</code>\n"
                    f"• <b>Код для родителя:</b> <code>{pr_code}</code>\n\n"
                    "<i>Используйте меню ниже для просмотра оценок и заданий. Желаем успехов!</i>"
                ),
                "uz": (
                    f"🎉 <b>Hurmatli {escape_html(req.full_name)},</b>\n\n"
                    "Maktab ma'muriyati o'quvchi arizangizni tasdiqladi!\n\n"
                    "📋 <b>O'quvchi ma'lumotlari:</b>\n"
                    f"• <b>Sinfingiz:</b> <code>{cls_name}</code>\n"
                    f"• <b>Maktab raqamingiz:</b> <code>{st_no}</code>\n"
                    f"• <b>O'quvchi kodingiz:</b> <code>{st_code}</code>\n"
                    f"• <b>Ota-ona ulanish kodi:</b> <code>{pr_code}</code>\n\n"
                    "<i>Baholar va vazifalarni kuzatish uchun quyidagi menyudan foydalaning. O'qishlaringizda zafarlar!</i>"
                ),
                "en": (
                    f"🎉 <b>Dear {escape_html(req.full_name)},</b>\n\n"
                    "The school administration has approved your student registration!\n\n"
                    "📋 <b>Student Information:</b>\n"
                    f"• <b>Class:</b> <code>{cls_name}</code>\n"
                    f"• <b>Roll Number:</b> <code>{st_no}</code>\n"
                    f"• <b>Student Code:</b> <code>{st_code}</code>\n"
                    f"• <b>Parent Code:</b> <code>{pr_code}</code>\n\n"
                    "<i>Use the menu below to view your grades and homework. Best of luck!</i>"
                )
            }.get(u_lang, "Student credentials.")

        else: # parent
            st = None
            if req.student_match_id:
                st = await session.get(Student, req.student_match_id)
            if not st:
                ps = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == req.telegram_id))).scalar_one_or_none()
                if ps:
                    st = await session.get(Student, ps.student_id)
            
            st_name = st.full_name if st else "Öğrenciniz"
            pr_code = st.parent_code if st else "-"

            cred_msg = {
                "tr": (
                    f"🎉 <b>Sayın {escape_html(req.full_name)},</b>\n\n"
                    "Okul idaresi veli başvurunuzu onayladı ve öğrencinizle bağlantınız kuruldu.\n\n"
                    "📋 <b>Veli Erişim Bilgileriniz:</b>\n"
                    f"• <b>Rolünüz:</b> 👨‍👩‍👧 Veli\n"
                    f"• <b>Bağlanan Öğrenci:</b> <b>{escape_html(st_name)}</b>\n"
                    f"• <b>Veli Giriş Kodunuz:</b> <code>{pr_code}</code>\n\n"
                    "<i>Aşağıdaki menüyü kullanarak öğrencimizin notlarını, devamsızlığını ve sağlık raporlarını takip edebilirsiniz.</i>"
                ),
                "ru": (
                    f"🎉 <b>Уважаемый(ая) {escape_html(req.full_name)},</b>\n\n"
                    "Администрация школы одобрила вашу заявку родителя.\n\n"
                    "📋 <b>Данные доступа:</b>\n"
                    f"• <b>Роль:</b> 👨‍👩‍👧 Родитель\n"
                    f"• <b>Ученик:</b> <b>{escape_html(st_name)}</b>\n"
                    f"• <b>Код доступа:</b> <code>{pr_code}</code>\n\n"
                    "<i>Используйте меню ниже для просмотра успеваемости и посещаемости ребенка.</i>"
                ),
                "uz": (
                    f"🎉 <b>Hurmatli {escape_html(req.full_name)},</b>\n\n"
                    "Maktab ma'muriyati ota-ona arizangizni tasdiqladi.\n\n"
                    "📋 <b>Kirish ma'lumotlari:</b>\n"
                    f"• <b>Rol:</b> 👨‍👩‍👧 Ota-ona\n"
                    f"• <b>Bog'langan o'quvchi:</b> <b>{escape_html(st_name)}</b>\n"
                    f"• <b>Ota-ona kodingiz:</b> <code>{pr_code}</code>\n\n"
                    "<i>Farzandingizning baholari va davomatini quyidagi menyu orqali kuzatib boring.</i>"
                ),
                "en": (
                    f"🎉 <b>Dear {escape_html(req.full_name)},</b>\n\n"
                    "The school administration has approved your parent request.\n\n"
                    "📋 <b>Parent Access Info:</b>\n"
                    f"• <b>Role:</b> 👨‍👩‍👧 Parent\n"
                    f"• <b>Linked Student:</b> <b>{escape_html(st_name)}</b>\n"
                    f"• <b>Parent Code:</b> <code>{pr_code}</code>\n\n"
                    "<i>Use the menu below to monitor student grades, attendance, and medical reports.</i>"
                )
            }.get(u_lang, "Parent credentials.")

        u_role = req.role or "guest"
        role_kb = get_role_reply_kb(u_role, u_lang)
        await safe_send_message(query.message.bot, req.telegram_id, cred_msg, reply_markup=role_kb, parse_mode="HTML")

        if query.message and query.message.reply_markup:
            new_rows = []
            for row in query.message.reply_markup.inline_keyboard:
                new_row = []
                for b in row:
                    if b.callback_data == query.data:
                        sent_lbl = {"tr": "✅ Şifre Gönderildi", "ru": "✅ Пароль отправлен", "uz": "✅ Parol yuborildi", "en": "✅ Password Sent"}.get(lang, "✅ Sent")
                        new_row.append(InlineKeyboardButton(text=sent_lbl, callback_data="noop"))
                    else:
                        new_row.append(b)
                new_rows.append(new_row)
            try:
                await query.message.edit_reply_markup(reply_markup=InlineKeyboardMarkup(inline_keyboard=new_rows))
            except Exception:
                pass

        alert_txt = {
            "tr": "✅ Şifre ve giriş bilgileri kullanıcıya başarıyla gönderildi!",
            "ru": "✅ Пароль и данные для входа успешно отправлены пользователю!",
            "uz": "✅ Parol va kirish ma'lumotlari foydalanuvchiga muvaffaqiyatli yuborildi!",
            "en": "✅ Credentials and password successfully sent to user!"
        }.get(lang, "Credentials sent successfully!")
        await query.answer(alert_txt, show_alert=True)

        await log_audit(session, query.from_user.id, (admin_user.full_name if admin_user else "Yönetici"), "ŞİFRE İLETİLDİ", f"Başvuru #{req.id} ({req.full_name}) kullanıcısına şifre/kodlar iletildi.")
        await session.commit()

@router.callback_query(any_state, F.data.startswith("adm:appr_cls:"))
async def cb_appr_select_class(query: CallbackQuery, state: FSMContext):
    class_name = query.data.split(":")[2].strip().upper()
    await state.update_data(appr_class_name=class_name)
    await state.set_state(Form.appr_st_no)

    user_id = query.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        data = await state.get_data()
        st_name = data.get("appr_st_name", "Öğrenci")
        req_id = data.get("appr_req_id")

        existing_st = (await session.execute(select(Student).where(Student.class_name == class_name))).scalars().all()
        rec_num = 101
        num_list = []
        for s in existing_st:
            try: num_list.append(int(s.student_number))
            except Exception: pass
        if num_list: rec_num = max(num_list) + 1

        prompt_no = {
            "tr": f"🎓 <b>Öğrenci Kayıt Onayı (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Öğrenci:</b> {escape_html(st_name)}\n🏫 <b>Sınıf:</b> <code>{class_name}</code>\n\nLütfen öğrencinin <b>Okul Numarasını</b> yazınız (Öneri: <code>{rec_num}</code>):",
            "ru": f"🎓 <b>Одобрение регистрации ученика (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Ученик:</b> {escape_html(st_name)}\n🏫 <b>Класс:</b> <code>{class_name}</code>\n\nВведите <b>Номер ученика</b> в школе (Рекомендация: <code>{rec_num}</code>):",
            "uz": f"🎓 <b>O'quvchini ro'yxatga olishni tasdiqlash (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>O'quvchi:</b> {escape_html(st_name)}\n🏫 <b>Sinf:</b> <code>{class_name}</code>\n\nIltimos, o'quvchining <b>Maktab raqamini</b> kiriting (Tavsiya: <code>{rec_num}</code>):",
            "en": f"🎓 <b>Student Registration Approval (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Student:</b> {escape_html(st_name)}\n🏫 <b>Class:</b> <code>{class_name}</code>\n\nPlease enter the <b>Student Number</b> (Recommended: <code>{rec_num}</code>):"
        }.get(lang, "Enter Student Number:")

        cancel_btn = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:view_req:{req_id}")]
        ])
        await safe_edit_or_answer(query, prompt_no, reply_markup=cancel_btn, parse_mode="HTML")
        await query.answer()

@router.message(Form.appr_st_class)
async def process_appr_type_class(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    cls_name = message.text.strip().upper()
    await state.update_data(appr_class_name=cls_name)
    await state.set_state(Form.appr_st_no)

    user_id = message.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        data = await state.get_data()
        st_name = data.get("appr_st_name", "Öğrenci")
        req_id = data.get("appr_req_id")

        prompt_no = {
            "tr": f"🎓 <b>Öğrenci Kayıt Onayı (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Öğrenci:</b> {escape_html(st_name)}\n🏫 <b>Sınıf:</b> <code>{cls_name}</code>\n\nLütfen öğrencinin <b>Okul Numarasını</b> yazınız (Örn: <code>101</code>):",
            "ru": f"🎓 <b>Одобрение регистрации ученика (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Ученик:</b> {escape_html(st_name)}\n🏫 <b>Класс:</b> <code>{cls_name}</code>\n\nВведите <b>Номер ученика</b> в школе (Напр: <code>101</code>):",
            "uz": f"🎓 <b>O'quvchini ro'yxatga olishni tasdiqlash (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>O'quvchi:</b> {escape_html(st_name)}\n🏫 <b>Sinf:</b> <code>{cls_name}</code>\n\nIltimos, o'quvchining <b>Maktab raqamini</b> kiriting (Masalan: <code>101</code>):",
            "en": f"🎓 <b>Student Registration Approval (#{req_id})</b>\n━━━━━━━━━━━━━━━━━━━━\n👤 <b>Student:</b> {escape_html(st_name)}\n🏫 <b>Class:</b> <code>{cls_name}</code>\n\nPlease enter the <b>Student Number</b> (e.g. <code>101</code>):"
        }.get(lang, "Enter Student Number:")

        cancel_btn = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:view_req:{req_id}")]
        ])
        last_mid = LAST_MENU_MSG_ID.get(message.chat.id)
        if last_mid:
            try:
                await message.bot.edit_message_text(chat_id=message.chat.id, message_id=last_mid, text=prompt_no, reply_markup=cancel_btn, parse_mode="HTML")
                return
            except Exception: pass
        s_m = await message.answer(prompt_no, reply_markup=cancel_btn, parse_mode="HTML")
        if s_m: LAST_MENU_MSG_ID[message.chat.id] = s_m.message_id

@router.message(Form.appr_st_no)
async def process_appr_st_no(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    st_no = message.text.strip()
    data = await state.get_data()
    req_id = data.get("appr_req_id")
    cls_name = data.get("appr_class_name", "9-A")
    st_name = data.get("appr_st_name", "Öğrenci")
    st_tg_id = data.get("appr_st_tg_id")
    st_phone = data.get("appr_st_phone")

    user_id = message.from_user.id
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, user_id)
        lang = admin_user.language if admin_user else "tr"

        dup = (await session.execute(select(Student).where(Student.class_name == cls_name, Student.student_number == st_no))).scalar_one_or_none()
        if dup:
            dup_warn = {
                "tr": f"⚠️ <b>Hata:</b> <code>{cls_name}</code> sınıfında <code>{st_no}</code> numaralı öğrenci ({escape_html(dup.full_name)}) zaten kayıtlı! Lütfen farklı bir numara yazınız:",
                "ru": f"⚠️ <b>Ошибка:</b> В классе <code>{cls_name}</code> ученик с номером <code>{st_no}</code> ({escape_html(dup.full_name)}) уже зарегистрирован! Введите другой номер:",
                "uz": f"⚠️ <b>Xatolik:</b> <code>{cls_name}</code> sinfida <code>{st_no}</code> raqamli o'quvchi ({escape_html(dup.full_name)}) allaqachon mavjud! Boshqa raqam kiriting:",
                "en": f"⚠️ <b>Error:</b> Roll number <code>{st_no}</code> in class <code>{cls_name}</code> is already taken by {escape_html(dup.full_name)}! Enter another number:"
            }.get(lang, "Duplicate number, enter another:")
            last_mid = LAST_MENU_MSG_ID.get(message.chat.id)
            if last_mid:
                try:
                    await message.bot.edit_message_text(chat_id=message.chat.id, message_id=last_mid, text=dup_warn, parse_mode="HTML")
                    return
                except Exception: pass
            s_m = await message.answer(dup_warn, parse_mode="HTML")
            if s_m: LAST_MENU_MSG_ID[message.chat.id] = s_m.message_id
            return

        await state.update_data(appr_st_no=st_no)

        confirm_card = {
            "tr": (
                f"🎓 <b>ÖĞRENCİ KAYIT VE ONAY KARTI</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"👤 <b>Adı Soyadı:</b> {escape_html(st_name)}\n"
                f"🏫 <b>Sınıf:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Okul Numarası:</b> <code>{st_no}</code>\n"
                f"📱 <b>Telefon:</b> {escape_html(st_phone or '-')}\n"
                f"🆔 <b>Telegram ID:</b> <code>{st_tg_id}</code>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                "Bu öğrenciyi onaylayıp sisteme kaydetmek istiyor musunuz?"
            ),
            "ru": (
                f"🎓 <b>КАРТОЧКА ПОДТВЕРЖДЕНИЯ УЧЕНИКА</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"👤 <b>ФИО:</b> {escape_html(st_name)}\n"
                f"🏫 <b>Класс:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Номер в школе:</b> <code>{st_no}</code>\n"
                f"📱 <b>Телефон:</b> {escape_html(st_phone or '-')}\n"
                f"🆔 <b>Telegram ID:</b> <code>{st_tg_id}</code>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                "Подтвердить и добавить ученика в систему?"
            ),
            "uz": (
                f"🎓 <b>O'QUVCHINI TASDIQLASH VA SAQLASH KARTASI</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"👤 <b>F.I.O:</b> {escape_html(st_name)}\n"
                f"🏫 <b>Sinf:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Maktab raqami:</b> <code>{st_no}</code>\n"
                f"📱 <b>Telefon:</b> {escape_html(st_phone or '-')}\n"
                f"🆔 <b>Telegram ID:</b> <code>{st_tg_id}</code>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                "Ushbu o'quvchini tasdiqlab tizimga qo'shasizmi?"
            ),
            "en": (
                f"🎓 <b>STUDENT APPROVAL CARD</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"👤 <b>Full Name:</b> {escape_html(st_name)}\n"
                f"🏫 <b>Class:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Roll Number:</b> <code>{st_no}</code>\n"
                f"📱 <b>Phone:</b> {escape_html(st_phone or '-')}\n"
                f"🆔 <b>Telegram ID:</b> <code>{st_tg_id}</code>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                "Do you confirm and add this student to the system?"
            )
        }.get(lang, "Confirm and add student?")

        btn_appr_label = get_text("btn_appr_request", lang)
        buttons = [
            [InlineKeyboardButton(text=btn_appr_label, callback_data=f"adm:appr_st_fin:{req_id}")],
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:view_req:{req_id}")]
        ]
        last_mid = LAST_MENU_MSG_ID.get(message.chat.id)
        if last_mid:
            try:
                await message.bot.edit_message_text(chat_id=message.chat.id, message_id=last_mid, text=confirm_card, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
                return
            except Exception: pass
        s_m = await message.answer(confirm_card, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        if s_m: LAST_MENU_MSG_ID[message.chat.id] = s_m.message_id

@router.callback_query(F.data.startswith("adm:appr_st_fin:"))
async def cb_appr_st_finalize(query: CallbackQuery, state: FSMContext):
    req_id = int(query.data.split(":")[2])
    data = await state.get_data()
    cls_name = data.get("appr_class_name", "9-A")
    st_no = data.get("appr_st_no", "101")
    await state.clear()

    user_id = query.from_user.id
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, user_id)
        lang = admin_user.language if admin_user else "tr"

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer(get_text("request_already_handled", lang), show_alert=True)
            await cb_admin_requests_list(query)
            return

        st_code = generate_secure_code("OGR")
        pr_code = generate_secure_code("VELI")

        student = Student(
            full_name=req.full_name,
            class_name=cls_name,
            student_number=st_no,
            student_code=st_code,
            parent_code=pr_code,
            is_student_code_burned=True,
            is_parent_code_burned=False,
            student_telegram_id=req.telegram_id
        )
        session.add(student)
        await session.flush()

        target_user = await session.get(User, req.telegram_id)
        if not target_user:
            target_user = User(telegram_id=req.telegram_id, language="tr")
            session.add(target_user)

        target_user.role = "student"
        target_user.full_name = req.full_name
        target_user.phone = req.phone
        target_user.failed_attempts = 0
        target_user.locked_until = None

        req.status = "approved"
        req.student_match_id = student.id
        req.reviewed_by = user_id
        req.reviewed_at = datetime.utcnow()

        await session.commit()

        await sync_admin_notif_resolution(
            query.message.bot,
            f"req:{req.id}",
            user_id,
            (admin_user.full_name if admin_user else "Yönetici"),
            "approved",
            f"Öğrenci Başvurusu: {req.full_name} ({cls_name} - No: {st_no})"
        )

        st_lang = target_user.language or "tr"
        st_welcome = {
            "tr": f"🎉 <b>Tebrikler! Başvurunuz Onaylandı.</b>\n\n👤 <b>Öğrenci:</b> {escape_html(req.full_name)}\n🏫 <b>Sınıfınız:</b> <code>{cls_name}</code>\n🔢 <b>Okul Numaranız:</b> <code>{st_no}</code>\n\nSisteme <b>Öğrenci</b> olarak başarıyla giriş yaptınız. Aşağıdaki menüyü kullanabilirsiniz:",
            "ru": f"🎉 <b>Поздравляем! Ваша заявка одобрена.</b>\n\n👤 <b>Ученик:</b> {escape_html(req.full_name)}\n🏫 <b>Класс:</b> <code>{cls_name}</code>\n🔢 <b>Номер в школе:</b> <code>{st_no}</code>\n\nВы успешно вошли как <b>Ученик</b>. Используйте меню ниже:",
            "uz": f"🎉 <b>Tabriklaymiz! Arizangiz tasdiqlandi.</b>\n\n👤 <b>O'quvchi:</b> {escape_html(req.full_name)}\n🏫 <b>Sinfingiz:</b> <code>{cls_name}</code>\n🔢 <b>Maktab raqamingiz:</b> <code>{st_no}</code>\n\nTizimga <b>O'quvchi</b> sifatida kirdingiz. Quyidagi menyudan foydalanishingiz mumkin:",
            "en": f"🎉 <b>Congratulations! Your Request Has Been Approved.</b>\n\n👤 <b>Student:</b> {escape_html(req.full_name)}\n🏫 <b>Class:</b> <code>{cls_name}</code>\n🔢 <b>Roll Number:</b> <code>{st_no}</code>\n\nYou are now logged in as <b>Student</b>. Use the menu below:"
        }.get(st_lang, "Approved.")
        old_u_mid = LAST_MENU_MSG_ID.get(req.telegram_id)
        s_u = await safe_send_message(query.message.bot, req.telegram_id, st_welcome, reply_markup=get_role_reply_kb("student", st_lang), parse_mode="HTML")
        if s_u:
            LAST_MENU_MSG_ID[req.telegram_id] = s_u.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(req.telegram_id, set()).add(s_u.message_id)
            if old_u_mid and old_u_mid != s_u.message_id:
                try: await query.message.bot.delete_message(chat_id=req.telegram_id, message_id=old_u_mid)
                except Exception: pass

        adm_success = {
            "tr": (
                f"✅ <b>Öğrenci Başarıyla Sisteme Kaydedildi ve Onaylandı!</b>\n\n"
                f"👤 <b>Öğrenci:</b> {escape_html(req.full_name)}\n"
                f"🏫 <b>Sınıf:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Okul No:</b> <code>{st_no}</code>\n"
                f"🔑 <b>Öğrenci Kodu:</b> <code>{st_code}</code>\n"
                f"🔑 <b>Veli Kodu:</b> <code>{pr_code}</code>"
            ),
            "ru": (
                f"✅ <b>Ученик успешно зарегистрирован и одобрен!</b>\n\n"
                f"👤 <b>ФИО:</b> {escape_html(req.full_name)}\n"
                f"🏫 <b>Класс:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Номер:</b> <code>{st_no}</code>\n"
                f"🔑 <b>Код ученика:</b> <code>{st_code}</code>\n"
                f"🔑 <b>Код родителя:</b> <code>{pr_code}</code>"
            ),
            "uz": (
                f"✅ <b>O'quvchi muvaffaqiyatli ro'yxatga olindi va tasdiqlandi!</b>\n\n"
                f"👤 <b>F.I.O:</b> {escape_html(req.full_name)}\n"
                f"🏫 <b>Sinf:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Raqami:</b> <code>{st_no}</code>\n"
                f"🔑 <b>O'quvchi kodi:</b> <code>{st_code}</code>\n"
                f"🔑 <b>Ota-ona kodi:</b> <code>{pr_code}</code>"
            ),
            "en": (
                f"✅ <b>Student Successfully Registered and Approved!</b>\n\n"
                f"👤 <b>Student:</b> {escape_html(req.full_name)}\n"
                f"🏫 <b>Class:</b> <code>{cls_name}</code>\n"
                f"🔢 <b>Roll:</b> <code>{st_no}</code>\n"
                f"🔑 <b>Student Code:</b> <code>{st_code}</code>\n"
                f"🔑 <b>Parent Code:</b> <code>{pr_code}</code>"
            )
        }.get(lang, "Student Registered.")

        btn_view_cls = {
            "tr": f"🏫 {cls_name} Sınıfını Gör",
            "ru": f"🏫 Посмотреть класс {cls_name}",
            "uz": f"🏫 {cls_name} sinfini ko'rish",
            "en": f"🏫 View Class {cls_name}"
        }.get(lang, f"🏫 View Class {cls_name}")

        btn_send_cred = {
            "tr": "📲 Kullanıcıya Şifreyi Gönder",
            "ru": "📲 Отправить пароль пользователю",
            "uz": "📲 Foydalanuvchiga parolni yuborish",
            "en": "📲 Send Credentials to User"
        }.get(lang, "📲 Send Credentials")

        buttons = [
            [InlineKeyboardButton(text=btn_send_cred, callback_data=f"adm:send_creds:{req_id}")],
            [InlineKeyboardButton(text=btn_view_cls, callback_data=f"adm:show_class:{cls_name}")],
            [InlineKeyboardButton(text=get_text("btn_requests", lang).split("(")[0].strip(), callback_data="adm:requests_list")]
        ]
        await safe_edit_or_answer(query, adm_success, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer("✅ " + ("Öğrenci eklendi!" if lang == "tr" else "O'quvchi qo'shildi!"))

@router.callback_query(F.data == "adm:approve_all_requests")
async def cb_admin_approve_all_requests(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        lang = admin_user.language if admin_user else "tr"
        if not is_admin_user(admin_user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        pending_requests = (await session.execute(select(AccessRequest).where(AccessRequest.status == "pending"))).scalars().all()
        if not pending_requests:
            await query.answer({"tr": "Bekleyen başvuru bulunamadı.", "ru": "Нет ожидающих заявок.", "uz": "Kutilayotgan arizalar topilmadi.", "en": "No pending applications found."}.get(lang, "No pending applications found."), show_alert=True)
            return

        approved_count = 0
        for req in pending_requests:
            req.status = "approved"
            req.reviewed_by = query.from_user.id
            req.reviewed_at = datetime.utcnow()
            target_user = await session.get(User, req.telegram_id)
            if target_user:
                target_user.role = req.role
                target_user.full_name = req.full_name
                target_user.phone = req.phone
                if req.role == "teacher":
                    new_code = generate_secure_code("HCA")
                    tch = Teacher(full_name=req.full_name, subject="Genel", auth_code=new_code, is_code_burned=True, telegram_id=req.telegram_id, assigned_classes="ALL")
                    session.add(tch)
                elif req.role in ["student", "parent"] and req.student_match_id:
                    st = await session.get(Student, req.student_match_id)
                    if st:
                        if req.role == "student":
                            st.is_student_code_burned = True
                            st.student_telegram_id = req.telegram_id
                        elif req.role == "parent":
                            st.is_parent_code_burned = True
                            target_user.current_child_id = st.id
                            session.add(ParentStudent(parent_telegram_id=req.telegram_id, student_id=st.id))
            approved_count += 1

        await log_audit(session, query.from_user.id, (admin_user.full_name if admin_user else "Yönetici"), "TOPLU BAŞVURU ONAYI", f"{approved_count} adet başvuru tek tıkla onaylandı.")
        await session.commit()

        toast = {"tr": f"✅ {approved_count} adet başvuru başarıyla onaylandı!", "ru": f"✅ {approved_count} заявок успешно одобрено!", "uz": f"✅ {approved_count} ta ariza muvaffaqiyatli tasdiqlandi!", "en": f"✅ {approved_count} requests approved successfully!"}.get(lang, "Approved.")
        await query.answer(toast, show_alert=True)
        await cb_cat_requests(query, None)

@router.callback_query(F.data.startswith('adm:rej_req:'))
async def cb_admin_reject_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        lang = admin_user.language if admin_user else "tr"
        if not is_admin_user(admin_user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer(get_text("request_already_handled", lang), show_alert=True)
            if query.message:
                try: await query.message.delete()
                except Exception:
                    try: await query.message.edit_reply_markup(reply_markup=None)
                    except Exception: pass
            return

        req.status = "rejected"
        req.reviewed_by = query.from_user.id
        req.reviewed_at = datetime.utcnow()
        target_user = await session.get(User, req.telegram_id)
        await session.commit()

        await sync_admin_notif_resolution(
            query.message.bot,
            f"req:{req.id}",
            query.from_user.id,
            (admin_user.full_name if admin_user else "Yönetici"),
            "rejected",
            f"Başvuru #{req.id}: {req.full_name} ({req.role})"
        )

        await safe_send_message(query.message.bot, req.telegram_id, get_text("req_rejected_user", target_user.language if target_user else "tr"), parse_mode="HTML")
        await safe_edit_or_answer(query, get_text("req_rejected_admin_msg", lang, id=req.id, name=escape_html(req.full_name)), parse_mode="HTML")
    lbl_rej = {"tr": "Talep reddedildi.", "ru": "Заявка отклонена.", "uz": "Ariza rad etildi.", "en": "Request rejected."}.get(lang, "Rejected.")
    await query.answer(lbl_rej)

# ======================================================================
# 10. RANDEVU, DERS PROGRAMI, YEMEKHANE VE KARA LİSTE
# ======================================================================

@router.callback_query(F.data == "act_book_app")
async def cb_parent_book_app_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        if not teachers:
            await query.answer(get_text("no_registered_teachers", lang), show_alert=True)
            return

        buttons = []
        for t in teachers:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {t.full_name} ({t.subject})", callback_data=f"book_tch:{t.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")])

        now_h = get_local_now().hour
        quiet_shield = ""
        if now_h >= 18 or now_h < 8:
            quiet_shield = "🌙 <i>" + {"tr": "Öğretmenlerimiz mesai saatleri (08:30 - 17:30) dışındadır. Talebinizi oluşturabilirsiniz, öğretmenimiz mesai başladığında inceleyecektir.", "ru": "Вне рабочих часов (08:30 - 17:30). Учитель рассмотрит заявку в рабочее время.", "uz": "Ish vaqtidan tashqari (08:30 - 17:30). O'qituvchi ish vaqtida ko'rib chiqadi.", "en": "Outside working hours (08:30 - 17:30). Teacher will review during work hours."}.get(lang, "Mesai dışı") + "</i>\n──────────────\n\n"

        app_text = quiet_shield + get_text("select_teacher_appointment", lang)
        await safe_edit_or_answer(query, app_text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("book_tch:"))
async def cb_parent_select_teacher(query: CallbackQuery, state: FSMContext):
    tch_id = int(query.data.split(":")[1])
    APP_CACHE[query.from_user.id] = {"teacher_id": tch_id}

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    await safe_edit_or_answer(query, get_text("prompt_appointment_note", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
    await state.set_state(Form.app_note)
    await query.answer()

@router.message(Form.app_note)
async def process_appointment_note(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    note_text = message.text.strip()
    data = APP_CACHE.pop(message.from_user.id, {})
    tch_id = data.get("teacher_id")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id) if tch_id else None
        st = await session.get(Student, user.current_child_id) if user and user.current_child_id else None

        app = Appointment(
            parent_telegram_id=message.from_user.id,
            teacher_id=tch_id,
            preferred_time=note_text,
            note=note_text,
            status="pending"
        )
        session.add(app)
        await session.commit()

        await message.answer(get_text("appointment_sent", lang))
        await render_clean_dashboard(message, user)

        if tch and tch.telegram_id:
            tch_user = await session.get(User, tch.telegram_id)
            tch_lang = tch_user.language if tch_user else "tr"
            st_name = st.full_name if st else get_role_label("student", tch_lang)
            cls_name = st.class_name if st else ""
            t_text = {
                "tr": (
                    f"🤝 <b>YENİ VELİ GÖRÜŞME TALEBİ</b>\n\n"
                    f"👤 <b>Veli:</b> {escape_html(user.full_name or 'Veli')}\n"
                    f"🧑‍🎓 <b>Öğrenci:</b> {escape_html(st_name)} ({escape_html(cls_name)})\n"
                    f"📝 <b>Talep / Zaman:</b> {escape_html(note_text)}"
                ),
                "ru": (
                    f"🤝 <b>НОВЫЙ ЗАПРОС НА ВСТРЕЧУ С РОДИТЕЛЕМ</b>\n\n"
                    f"👤 <b>Родитель:</b> {escape_html(user.full_name or 'Родитель')}\n"
                    f"🧑‍🎓 <b>Ученик:</b> {escape_html(st_name)} ({escape_html(cls_name)})\n"
                    f"📝 <b>Запрос / Время:</b> {escape_html(note_text)}"
                ),
                "uz": (
                    f"🤝 <b>OTA-ONA BILAN UCHRASHUV SO'ROVI</b>\n\n"
                    f"👤 <b>Ota-ona:</b> {escape_html(user.full_name or 'Ota-ona')}\n"
                    f"🧑‍🎓 <b>O'quvchi:</b> {escape_html(st_name)} ({escape_html(cls_name)})\n"
                    f"📝 <b>So'rov / Vaqt:</b> {escape_html(note_text)}"
                ),
                "en": (
                    f"🤝 <b>NEW PARENT-TEACHER MEETING REQUEST</b>\n\n"
                    f"👤 <b>Parent:</b> {escape_html(user.full_name or 'Parent')}\n"
                    f"🧑‍🎓 <b>Student:</b> {escape_html(st_name)} ({escape_html(cls_name)})\n"
                    f"📝 <b>Request / Time:</b> {escape_html(note_text)}"
                )
            }.get(tch_lang, "New Meeting Request")
            t_kb = InlineKeyboardMarkup(inline_keyboard=[
                [
                    InlineKeyboardButton(text=get_text("btn_appr_appointment", tch_lang), callback_data=f"tch:appr_app:{app.id}"),
                    InlineKeyboardButton(text=get_text("btn_not_available", tch_lang), callback_data=f"tch:rej_app:{app.id}")
                ]
            ])
            await safe_send_message(message.bot, tch.telegram_id, t_text, reply_markup=t_kb, parse_mode="HTML")

@router.callback_query(F.data == "tch:appointments")
async def cb_teacher_appointments_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        tch_id = tch.id if tch else 0

        apps = (await session.execute(
            select(Appointment).where(Appointment.teacher_id == tch_id, Appointment.status == "pending")
        )).scalars().all()

        if not apps:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_pending_appointments", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        for a in apps:
            btn_txt = f"🤝 Randevu #{a.id} ({a.preferred_time[:20]})"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"tch:view_app:{a.id}")])
        buttons.append(get_nav_buttons(lang))

        await safe_edit_or_answer(query, get_text("pending_appointments_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:view_app:"))
async def cb_teacher_view_appointment(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        app = await session.get(Appointment, app_id)
        if not app:
            await query.answer(get_text("appointment_not_found", lang), show_alert=True)
            return

        parent_user = await session.get(User, app.parent_telegram_id)
        p_name = escape_md(parent_user.full_name) if parent_user else "Veli"

        text = (
            f"🤝 *Veli Görüşme Talebi* (#{app.id})\n\n"
            f"👤 *Veli:* {p_name}\n"
            f"🕒 *Talep Edilen Zaman:* {escape_md(app.preferred_time)}\n"
            f"📝 *Not / Konu:* {escape_md(app.note or '-')}\n"
            f"📌 *Durum:* {app.status}\n\n"
            f"Lütfen yapılacak işlemi seçiniz:"
        )
        buttons = [
            [
                InlineKeyboardButton(text=get_text("btn_appr_appointment", lang), callback_data=f"tch:appr_app:{app.id}"),
                InlineKeyboardButton(text=get_text("btn_not_available", lang), callback_data=f"tch:rej_app:{app.id}")
            ],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:appointments")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:appr_app:"))
async def cb_teacher_approve_app(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        app = await session.get(Appointment, app_id)
        if not tch or not app or app.teacher_id != tch.id:
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return
        if app and app.status == "pending":
            app.status = "approved"
            await session.commit()
            p_u = await session.get(User, app.parent_telegram_id)
            p_lang = p_u.language if p_u else "tr"
            app_receipt_txt = (
                f"🗓️ <b>Veli - Öğretmen Randevu Onay Fişi</b>\n"
                f"👤 <b>Öğretmen:</b> {escape_html(tch.full_name)}\n"
                f"🕒 <b>Randevu Zamanı:</b> <code>{escape_html(app.preferred_time)}</code>\n"
                f"🆔 <b>Fiş No:</b> <code>#RND-{app.id:04d}</code>\n"
                f"✅ <i>Randevunuz öğretmen tarafından onaylanmıştır. Lütfen belirtilen saatte hazır bulununuz.</i>"
            )
            await send_digital_record(
                bot=query.message.bot,
                chat_id=app.parent_telegram_id,
                category="appointment",
                title=f"Randevu Onay Fişi: {tch.full_name}",
                content=app_receipt_txt,
                lang=p_lang,
                expires_hours=168
            )
            await safe_send_message(query.message.bot, app.parent_telegram_id, get_text("appointment_approved_msg", p_lang), parse_mode="HTML")
            await safe_edit_or_answer(query, get_text("appointment_confirmed_toast", lang))
    await query.answer()

@router.callback_query(F.data.startswith("tch:rej_app:"))
async def cb_teacher_reject_app(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        app = await session.get(Appointment, app_id)
        if app and app.status == "pending":
            app.status = "rejected"
            await session.commit()
            p_u = await session.get(User, app.parent_telegram_id)
            p_lang = p_u.language if p_u else "tr"
            await safe_send_message(query.message.bot, app.parent_telegram_id, get_text("appointment_rejected_msg", p_lang), parse_mode="HTML")
            await safe_edit_or_answer(query, get_text("appointment_rejected_msg", lang))
    await query.answer()

@router.callback_query(F.data == "parent:toggle_briefing")
async def cb_parent_toggle_briefing(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if user:
            user.evening_briefing = not user.evening_briefing
            await session.commit()
            await render_clean_dashboard(query, user)
    await query.answer(get_text("setting_updated_toast", lang))

@router.callback_query(F.data == "act_view_notices")
async def cb_view_notices(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        notices = (await session.execute(select(BroadcastNotice).order_by(desc(BroadcastNotice.created_at)).limit(5))).scalars().all()
        if not notices:
            text = {"tr": "📢 Henüz yayınlanmış bir duyuru bulunmamaktadır.", "ru": "📢 Объявлений пока нет.", "uz": "📢 Hozircha e'lonlar mavjud emas.", "en": "📢 No announcements yet."}.get(lang, "📢 No notices.")
        else:
            header_n = {"tr": "📢 *Okul Duyuruları (Son Duyurular):*\n\n", "ru": "📢 *Объявления школы:*\n\n", "uz": "📢 *Maktab e'lonlari:*\n\n", "en": "📢 *School Announcements:*\n\n"}.get(lang, "📢 *Announcements:*\n\n")
            text = header_n
            for n in notices:
                text += f"📌 *{n.created_at.strftime('%d.%m.%Y %H:%M')}*\n{escape_md(n.content)}\n\n"

        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "act_view_sched")
async def cb_view_schedule(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        if user and user.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
            all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
            if tch and tch.assigned_classes and tch.assigned_classes != "ALL":
                allowed = [c.strip() for c in tch.assigned_classes.split(",") if c.strip()]
                classes = [c for c in all_classes if c in allowed] or all_classes
            else:
                classes = all_classes

            buttons = []
            row = []
            for c in classes:
                row.append(InlineKeyboardButton(text=f"📅 {c}", callback_data=f"sched_cls:{c}"))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row: buttons.append(row)
            buttons.append(get_nav_buttons(lang))
            await safe_edit_or_answer(query, get_text("schedule_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        cls_name = None
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: cls_name = st.class_name
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == query.from_user.id))).scalar_one_or_none()
            if st: cls_name = st.class_name

        if not cls_name:
            buttons = [get_nav_buttons(lang)]
            no_info_txt = get_text("no_linked_student", lang) if user and user.role in ("parent", "student") else get_text("no_classes_found", lang)
            await safe_edit_or_answer(query, no_info_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == cls_name))).scalar_one_or_none()
        content = sched.schedule_text if sched else "• 1. Ders: 09:00 - Matematik\n• 2. Ders: 09:50 - Fizik\n• 3. Ders: 10:40 - Türkçe\n• 4. Ders: 11:30 - Tarih\n• 5. Ders: 13:00 - Biyoloji"

        now_dt = get_local_now()
        cur_hm = now_dt.strftime("%H:%M")
        is_weekday = now_dt.weekday() < 5

        # Canlı Ders Hesaplama (Örn: 09:00-09:40, 09:50-10:30, 10:40-11:20, 11:30-12:10, 13:00-13:40)
        periods = [
            ("09:00", "09:40", "1. Ders"),
            ("09:50", "10:30", "2. Ders"),
            ("10:40", "11:20", "3. Ders"),
            ("11:30", "12:10", "4. Ders"),
            ("13:00", "13:40", "5. Ders"),
            ("13:50", "14:30", "6. Ders")
        ]
        live_status = ""
        if is_weekday:
            for p_start, p_end, p_label in periods:
                if p_start <= cur_hm <= p_end:
                    live_status = f"🟢 <b>Şu An Derste:</b> {p_label} (Bitiş: {p_end})\n"
                    break
                elif cur_hm < p_start:
                    live_status = f"⏳ <b>Sıradaki Ders:</b> {p_label} ({p_start})\n"
                    break
            if not live_status:
                if cur_hm > "14:30":
                    live_status = "🌙 <b>Bugünkü Dersler Tamamlandı</b>\n"
                else:
                    live_status = "🔔 <b>Teneffüs / Ara</b>\n"
        else:
            live_status = "🌴 <b>Hafta Sonu Tatili</b> (Pazartesi 09:00'da başlar)\n"

        header_sc = {"tr": f"📅 <b>{escape_html(cls_name)} Sınıfı Ders Programı:</b>", "ru": f"📅 <b>Расписание уроков {escape_html(cls_name)} класса:</b>", "uz": f"📅 <b>{escape_html(cls_name)} sinfi dars jadvali:</b>", "en": f"📅 <b>Timetable for Class {escape_html(cls_name)}:</b>"}.get(lang, f"📅 <b>{escape_html(cls_name)} Timetable:</b>")
        text = f"{header_sc}\n──────────────\n{live_status}\n{format_telegram_html(content)}"
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("sched_cls:"))
async def cb_view_schedule_class(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == class_name))).scalar_one_or_none()
        content = sched.schedule_text if sched else "• 1. Ders: 09:00 - Matematik\n• 2. Ders: 09:50 - Fizik\n• 3. Ders: 10:40 - Türkçe\n• 4. Ders: 11:30 - Tarih\n• 5. Ders: 13:00 - Biyoloji"

        header_sc = {"tr": f"📅 *{escape_md(class_name)} Sınıfı Ders Programı:*", "ru": f"📅 *Расписание уроков {escape_md(class_name)} класса:*", "uz": f"📅 *{escape_md(class_name)} sinfi dars jadvali:*", "en": f"📅 *Timetable for Class {escape_md(class_name)}:*"}.get(lang, f"📅 *{escape_md(class_name)} Timetable:*")
        text = f"{header_sc}\n\n{escape_md(content)}"
        buttons = [get_nav_buttons(lang, back_callback="act_view_sched")]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "act_view_cafe")
async def cb_view_cafeteria(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = get_local_date()

        menu = (await session.execute(select(CafeteriaMenu).where(CafeteriaMenu.date == today))).scalar_one_or_none()
        content = menu.menu_text if menu else "🍲 Mercimek Çorbası\n🍗 Fırında Tavuk & Pilav\n🥗 Mevsim Salatası\n🍎 Meyve / Ayran"

        header_m = {"tr": f"🍲 *Bugünün Yemek Menüsü ({today.strftime('%d.%m.%Y')}):*", "ru": f"🍲 *Меню столовой на сегодня ({today.strftime('%d.%m.%Y')}):*", "uz": f"🍲 *Bugungi oshxona menyusi ({today.strftime('%d.%m.%Y')}):*", "en": f"🍲 *Today's Cafeteria Menu ({today.strftime('%d.%m.%Y')}):*"}.get(lang, f"🍲 *Menu ({today.strftime('%d.%m.%Y')}):*")
        text = f"{header_m}\n\n{escape_md(content)}"
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:blacklist")
async def cb_admin_blacklist(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        now = datetime.utcnow()
        blocked_users = (await session.execute(select(User).where((User.is_blacklisted == True) | (User.locked_until > now)))).scalars().all()

        if not blocked_users:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_blacklisted", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        lbl_bl = {"tr": "🚫 Kara Liste", "ru": "🚫 Черный список", "uz": "🚫 Qora ro'yxat", "en": "🚫 Blacklist"}.get(lang, "🚫 Blacklist")
        lbl_lock = {"tr": "⏱️ Kilitli", "ru": "⏱️ Заблокирован", "uz": "⏱️ Bloklangan", "en": "⏱️ Locked"}.get(lang, "⏱️ Locked")
        lbl_unban = {"tr": "🟢 Engeli Kaldır", "ru": "🟢 Разблокировать", "uz": "🟢 Blokdan chiqarish", "en": "🟢 Unban"}.get(lang, "🟢 Unban")
        for u in blocked_users:
            b_type = lbl_bl if u.is_blacklisted else lbl_lock
            name_str = u.full_name or f"ID: {u.telegram_id}"
            buttons.append([InlineKeyboardButton(text=f"{lbl_unban}: {b_type} - {name_str}", callback_data=f"adm:unban:{u.telegram_id}")])

        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("blacklisted_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:unban:"))
async def cb_admin_unban(query: CallbackQuery):
    t_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        target_u = await session.get(User, t_id)
        if target_u:
            target_u.is_blacklisted = False
            target_u.failed_attempts = 0
            target_u.locked_until = None
            await session.commit()
            await safe_send_message(query.message.bot, t_id, get_text("admin_unban_notification", target_u.language if target_u else "tr"), parse_mode="HTML")
            await query.answer(get_text("unban_success", lang), show_alert=True)
            await cb_admin_blacklist(query)
            return
    await query.answer()

@router.callback_query(F.data == "adm:sched_edit_menu")
async def cb_admin_sched_edit_menu(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = await get_all_school_classes(session)
        if not classes:
            buttons = [
                [InlineKeyboardButton(text="➕ " + ("Sınıf Ekle" if lang=="tr" else ("Добавить класс" if lang=="ru" else ("Sinf qo'shish" if lang=="uz" else "Add Class"))), callback_data="adm:add_class")],
                get_nav_buttons(lang, back_callback="adm:cat_tools_reports")
            ]
            await safe_edit_or_answer(query, get_text("no_classes_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"✏️ {c} {get_text('btn_class_sched', lang)}", callback_data=f"adm:sched_edit_cls:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append(get_nav_buttons(lang))

        await safe_edit_or_answer(query, get_text("admin_sched_edit_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:sched_edit_cls:"))
async def cb_admin_sched_edit_class(query: CallbackQuery, state: FSMContext):
    class_name = query.data.split(":")[2]
    await state.update_data(edit_sched_class=class_name)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == class_name))).scalar_one_or_none()
        current_txt = sched.schedule_text if sched else "Henüz özel program girilmemiş."

    prompt_t = {
        "tr": f"📅 *{escape_md(class_name)} Sınıfı Ders Programı*\n\n📌 *Mevcut Program:*\n{escape_md(current_txt)}\n\n✍️ Lütfen yeni haftalık programı yazıp gönderiniz:",
        "ru": f"📅 *Расписание уроков класса {escape_md(class_name)}*\n\n📌 *Текущее расписание:*\n{escape_md(current_txt)}\n\n✍️ Введите новое расписание уроков:",
        "uz": f"📅 *{escape_md(class_name)} sinfi dars jadvali*\n\n📌 *Amaldagi jadval:*\n{escape_md(current_txt)}\n\n✍️ Iltimos, yangi haftalik dars jadvalini yuboring:",
        "en": f"📅 *Timetable for Class {escape_md(class_name)}*\n\n📌 *Current Timetable:*\n{escape_md(current_txt)}\n\n✍️ Please send the updated weekly timetable:"
    }.get(lang, f"📅 *{escape_md(class_name)} Timetable*\n\n✍️ Enter new timetable:")

    buttons = [get_nav_buttons(lang, back_callback="adm:sched_edit_menu")]
    await safe_edit_or_answer(query, prompt_t, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.sched_update_text)
    await query.answer()

@router.message(Form.sched_update_text)
async def process_sched_update_text(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    new_text = message.text.strip()
    data = await state.get_data()
    class_name = data.get("edit_sched_class", "9-A")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == class_name))).scalar_one_or_none()
        if not sched:
            sched = ScheduleItem(class_name=class_name, schedule_text=new_text)
            session.add(sched)
        else:
            sched.schedule_text = new_text
        await session.commit()

        await message.answer(get_text("admin_sched_updated", lang, class_name=escape_md(class_name)), reply_markup=get_role_reply_kb("admin", lang), parse_mode="HTML")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data == "adm:menu_edit")
async def cb_admin_menu_edit(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("prompt_menu_update", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.menu_update_text)
    await query.answer()

@router.message(Form.menu_update_text)
async def process_menu_update_text(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    m_text = message.text.strip()
    await state.clear()
    today = get_local_date()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        menu = (await session.execute(select(CafeteriaMenu).where(CafeteriaMenu.date == today))).scalar_one_or_none()
        if not menu:
            menu = CafeteriaMenu(date=today, menu_text=m_text)
            session.add(menu)
        else:
            menu.menu_text = m_text
        await session.commit()

        await message.answer(get_text("menu_updated", lang))
        await render_clean_dashboard(message, user)

# ======================================================================
# 11. YÖNETİCİ İŞLEMLERİ (ÖĞRENCİ, ÖĞRETMEN, İNTERAKTİF SINIF MASASI & DEVRETME)
# ======================================================================

@router.callback_query(F.data == "adm:add_class")
async def cb_admin_add_class_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:classes")
    prompt_add_c = {
        "tr": "🏫 <b>Yeni Sınıf Ekleme</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen eklemek istediğiniz yeni sınıfın adını yazınız (Örn: <code>10-A</code>, <code>11-B</code>, <code>5-C</code>):",
        "ru": "🏫 <b>Добавление нового класса</b>\n━━━━━━━━━━━━━━━━━━━━\nПожалуйста, введите название нового класса (Напр: <code>10-A</code>, <code>11-B</code>, <code>5-C</code>):",
        "uz": "🏫 <b>Yangi sinf qo'shish</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, yangi sinf nomini kiriting (Masalan: <code>10-A</code>, <code>11-B</code>, <code>5-C</code>):",
        "en": "🏫 <b>Add New Class</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter the new class name (e.g. <code>10-A</code>, <code>11-B</code>, <code>5-C</code>):"
    }.get(lang, "Enter new class name:")

    await safe_edit_or_answer(query, prompt_add_c, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_class_name)
    await query.answer()

@router.message(Form.add_class_name)
async def process_admin_add_class_name(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    raw_name = message.text.strip().upper()
    raw_name = raw_name.replace("İ", "I").replace("ı", "I")
    raw_name = re.sub(r'\s+', '', raw_name)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        if not raw_name or len(raw_name) < 1 or len(raw_name) > 15 or not re.match(r'^[A-Z0-9\-_]+$', raw_name):
            err_txt = {
                "tr": "⚠️ Geçersiz sınıf adı! Lütfen harf ve rakamlardan oluşan geçerli bir sınıf adı giriniz (Örn: <code>10-A</code>):",
                "ru": "⚠️ Некорректное название класса! Введите правильное название (Напр: <code>10-A</code>):",
                "uz": "⚠️ Noto'g'ri sinf nomi! Iltimos, to'g'ri nom kiriting (Masalan: <code>10-A</code>):",
                "en": "⚠️ Invalid class name! Please enter a valid name (e.g. <code>10-A</code>):"
            }.get(lang, "Invalid class name:")
            await message.answer(err_txt, parse_mode="HTML")
            return

        existing = (await session.execute(select(ClassRoom).where(ClassRoom.name == raw_name))).scalar_one_or_none()
        st_existing = (await session.execute(select(Student).where(Student.class_name == raw_name))).scalar_one_or_none()
        if existing or st_existing:
            dup_txt = {
                "tr": f"⚠️ <b>{raw_name}</b> sınıfı sistemde zaten kayıtlı! Lütfen farklı bir sınıf adı giriniz:",
                "ru": f"⚠️ Класс <b>{raw_name}</b> уже существует! Введите другое название:",
                "uz": f"⚠️ <b>{raw_name}</b> sinfi tizimda allaqachon mavjud! Boshqa nom kiriting:",
                "en": f"⚠️ Class <b>{raw_name}</b> already exists! Please enter another name:"
            }.get(lang, f"Class {raw_name} already exists!")
            await message.answer(dup_txt, parse_mode="HTML")
            return

        new_cr = ClassRoom(name=raw_name)
        session.add(new_cr)
        await log_audit(session, message.from_user.id, (user.full_name or "Yönetici"), "SINIF EKLENDİ", f"Yeni sınıf oluşturuldu: {raw_name}")
        await session.commit()
        await state.clear()

        succ_txt = {
            "tr": f"✅ <b>{raw_name} Sınıfı Başarıyla Eklendi!</b>\n━━━━━━━━━━━━━━━━━━━━\nSınıf sisteme kaydedildi. Şimdi bu sınıfa öğrenci veya ders programı tanımlayabilirsiniz.",
            "ru": f"✅ <b>Класс {raw_name} успешно добавлен!</b>\n━━━━━━━━━━━━━━━━━━━━\nКласс зарегистрирован. Теперь вы можете добавить в него учеников или расписание.",
            "uz": f"✅ <b>{raw_name} sinfi muvaffaqiyatli qo'shildi!</b>\n━━━━━━━━━━━━━━━━━━━━\nSinf tizimga saqlandi. Endi ushbu sinfga o'quvchilar qo'shishingiz mumkin.",
            "en": f"✅ <b>Class {raw_name} Added Successfully!</b>\n━━━━━━━━━━━━━━━━━━━━\nClass registered. You can now add students or timetables to this class."
        }.get(lang, f"Class {raw_name} added successfully!")

        buttons = [
            [InlineKeyboardButton(text=f"🏫 {raw_name} Sınıfını Gör", callback_data=f"adm:show_class:{raw_name}")],
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes")]
        ]
        role_kb = get_role_reply_kb("admin", lang)
        sent_m = await safe_send_message(
            message.bot,
            chat_id=message.chat.id,
            text=succ_txt,
            reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
            parse_mode="HTML"
        )
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data.startswith("adm:del_class:"))
async def cb_admin_del_class(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        st_cnt = (await session.execute(select(func.count(Student.id)).where(Student.class_name == class_name))).scalar() or 0
        if st_cnt > 0:
            warn = {
                "tr": f"⚠️ {class_name} sınıfında {st_cnt} öğrenci kayıtlı olduğu için sınıf doğrudan silinemez! Önce öğrencileri siliniz veya başka sınıfa aktarınız.",
                "ru": f"⚠️ В классе {class_name} есть {st_cnt} учеников. Сначала удалите или переведите учеников.",
                "uz": f"⚠️ {class_name} sinfida {st_cnt} ta o'quvchi borligi sababli sinfni o'chirib bo'lmaydi!",
                "en": f"⚠️ Cannot delete class {class_name} because {st_cnt} students are enrolled."
            }.get(lang, f"Cannot delete class with {st_cnt} students.")
            await query.answer(warn, show_alert=True)
            return

        await session.execute(delete(ClassRoom).where(ClassRoom.name == class_name))
        await session.execute(delete(ScheduleItem).where(ScheduleItem.class_name == class_name))
        await session.execute(delete(ExamSchedule).where(ExamSchedule.class_name == class_name))
        await session.execute(delete(Homework).where(Homework.class_name == class_name))
        await log_audit(session, query.from_user.id, (user.full_name or "Yönetici"), "SINIF SİLİNDİ", f"Sınıf silindi: {class_name}")
        await session.commit()
        tst = {"tr": f"🗑️ {class_name} sınıfı silindi.", "ru": f"🗑️ Класс {class_name} удален.", "uz": f"🗑️ {class_name} sinfi o'chirildi.", "en": f"🗑️ Class {class_name} deleted."}.get(lang, f"Class {class_name} deleted.")
        await query.answer(tst, show_alert=True)
        await cb_classes_list(query, None)

@router.callback_query(F.data.startswith("adm:add_st_to_cls:"))
async def cb_add_student_to_specific_class(query: CallbackQuery, state: FSMContext):
    await state.clear()
    class_name = query.data.split(":")[2]
    await state.update_data(class_name=class_name)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback=f"adm:show_class:{class_name}")
    prompt_msg = {
        "tr": f"👤 <b>{class_name} Sınıfına Öğrenci Ekleme</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen öğrencinin Adını ve Soyadını yazınız:",
        "ru": f"👤 <b>Добавление ученика в класс {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\nВведите ФИО ученика:",
        "uz": f"👤 <b>{class_name} sinfiga o'quvchi qo'shish</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, o'quvchining Ism va Familiyasini kiriting:",
        "en": f"👤 <b>Add Student to {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter the student's Full Name:"
    }.get(lang, "Enter Full Name:")

    await safe_edit_or_answer(query, prompt_msg, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_student_name)
    await query.answer()

@router.callback_query(F.data == "adm:add_student")
async def cb_start_add_student(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:classes")
    await safe_edit_or_answer(query, get_text("prompt_student_name", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_student_name)
    await query.answer()

@router.message(Form.add_student_name)
async def process_student_name(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    name_val = message.text.strip()
    if not name_val or len(name_val) < 1:
        await message.answer(get_text("student_name_invalid", lang))
        return
    await state.update_data(name=name_val)

    data = await state.get_data()
    cls_target = data.get("class_name")
    back_cb = f"adm:show_class:{cls_target}" if cls_target else "adm:classes"
    cancel_kb = get_inline_cancel_kb(lang, back_callback=back_cb)
    if cls_target:
        await safe_edit_or_answer(message, get_text("prompt_student_no", lang), reply_markup=cancel_kb, parse_mode="HTML")
        await state.set_state(Form.add_student_no)
        return

    await safe_edit_or_answer(message, get_text("prompt_student_class", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_student_class)

@router.message(Form.add_student_class)
async def process_student_class(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    cls_name = message.text.strip().upper()
    await state.update_data(class_name=cls_name)
    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:classes")
    await safe_edit_or_answer(message, get_text("prompt_student_no", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_student_no)

@router.message(Form.add_student_no)
async def process_student_no(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        data = await state.get_data()
        await state.clear()
        full_name = data.get("name")
        class_name = data.get("class_name")
        student_no = message.text.strip()
        existing_no = (await session.execute(select(Student).where(Student.class_name == class_name, Student.student_number == student_no))).scalar_one_or_none()
        if existing_no:
            await message.answer(get_text("duplicate_student_no_error", lang, class_name=class_name, no=student_no))
            return

        st_code = generate_secure_code("OGR")
        pr_code = generate_secure_code("VELI")

        student = Student(
            full_name=full_name,
            class_name=class_name,
            student_number=student_no,
            student_code=st_code,
            parent_code=pr_code
        )
        session.add(student)
        await session.commit()

        text = get_text("student_added_card", lang, name=escape_md(full_name), class_name=escape_md(class_name), no=escape_md(student_no), st_code=st_code, pr_code=pr_code)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_add_another", lang), callback_data="adm:add_student")]
        ]
        await safe_edit_or_answer(message, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@router.callback_query(F.data == "adm:add_teacher")
async def cb_start_add_teacher(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:teachers")
    await safe_edit_or_answer(query, get_text("prompt_teacher_name", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_teacher_name)
    await query.answer()

@router.message(Form.add_teacher_name)
async def process_teacher_name(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    name_val = message.text.strip()
    if not name_val or len(name_val) < 1:
        await message.answer(get_text("teacher_name_invalid", lang))
        return
    await state.update_data(name=name_val)
    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:teachers")
    await safe_edit_or_answer(message, get_text("prompt_teacher_subject", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.add_teacher_subject)

@router.message(Form.add_teacher_subject)
async def process_teacher_subject(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        data = await state.get_data()
        await state.clear()
        full_name = data.get("name")
        subject = message.text.strip()
        auth_code = generate_secure_code("HCA")

        teacher = Teacher(full_name=full_name, subject=subject, auth_code=auth_code, assigned_classes="ALL")
        session.add(teacher)
        await session.commit()

        text = get_text("teacher_added_card", lang, name=escape_md(full_name), subject=escape_md(subject), code=auth_code)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_add_another", lang), callback_data="adm:add_teacher")]
        ]
        await safe_edit_or_answer(message, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@router.callback_query(F.data == "adm:search_student")
async def cb_search_student_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:cat_staff")]])
    prompt_card = f"🔍 <b>{get_text('prompt_search_student', lang)}</b>\n━━━━━━━━━━━━━━━━━━━━\n<i>Lütfen aramak istediğiniz öğrencinin adını veya okul numarasını yazınız:</i>"
    await safe_edit_or_answer(query, prompt_card, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_search_query)
    await query.answer()

@router.message(Form.waiting_search_query)
async def process_search_query(message: Message, state: FSMContext):
    raw_q = message.text.strip()
    clean_q = clean_unicode_text(raw_q)
    await state.clear()
    try: await message.delete()
    except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        conds = [
            func.lower(Student.full_name).contains(clean_q),
            Student.student_number == raw_q
        ]
        results = (await session.execute(select(Student).where(or_(*conds)).order_by(Student.class_name, Student.student_number).limit(15))).scalars().all()

        if not results:
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_search_again", lang), callback_data="adm:search_student")]
            ]
            await message.answer(get_text("search_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        if len(results) == 1:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data=f"adm:st_card:{results[0].id}")
            await cb_student_card(dummy_q)
            return

        buttons = []
        for s in results:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.class_name} - №{s.student_number})", callback_data=f"adm:st_card:{s.id}")])

        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(message, get_text("search_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@router.callback_query(F.data == "adm:search_teacher")
async def cb_search_teacher_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:teachers")]])
    prompt_card = f"🔍 <b>{get_text('teacher_search_prompt', lang)}</b>\n━━━━━━━━━━━━━━━━━━━━\n<i>Lütfen aramak istediğiniz öğretmenin adını veya branşını yazınız:</i>"
    await safe_edit_or_answer(query, prompt_card, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_search_teacher_query)
    await query.answer()

@router.message(Form.waiting_search_teacher_query)
async def process_search_teacher_query(message: Message, state: FSMContext):
    clean_q = clean_unicode_text(message.text)
    await state.clear()
    try: await message.delete()
    except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        conds = [
            func.lower(Teacher.full_name).contains(clean_q),
            func.lower(Teacher.subject).contains(clean_q)
        ]
        results = (await session.execute(select(Teacher).where(or_(*conds)).order_by(Teacher.full_name).limit(15))).scalars().all()

        if not results:
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_search_again", lang), callback_data="adm:search_teacher")]
            ]
            await message.answer(get_text("teacher_search_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        if len(results) == 1:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data=f"adm:tch_card:{results[0].id}")
            await cb_admin_teacher_card(dummy_q)
            return

        buttons = []
        for t in results:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {t.full_name} ({t.subject})", callback_data=f"adm:tch_card:{t.id}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(message, get_text("teacher_search_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

@router.callback_query(F.data.startswith("adm:teachers"))
async def cb_admin_teachers_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        parts = query.data.split(":")
        page = int(parts[2]) if len(parts) > 2 and parts[2].isdigit() else 0
        PAGE_SIZE = 10

        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        total_teachers = len(teachers)
        total_pages = max(1, (total_teachers + PAGE_SIZE - 1) // PAGE_SIZE)
        page = max(0, min(page, total_pages - 1))
        paged_teachers = teachers[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]

        buttons = []
        for t in paged_teachers:
            status_dot = "🟢" if t.telegram_id else "⚪"
            buttons.append([InlineKeyboardButton(text=f"{status_dot} {t.full_name} ({t.subject})", callback_data=f"adm:tch_card:{t.id}")])

        if total_pages > 1:
            nav_row = []
            if page > 0:
                nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:teachers:{page-1}"))
            nav_row.append(InlineKeyboardButton(text=f"📄 {page+1}/{total_pages}", callback_data="noop"))
            if page < total_pages - 1:
                nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:teachers:{page+1}"))
            buttons.append(nav_row)

        buttons.append([
            InlineKeyboardButton(text=get_text("btn_search_teacher", lang), callback_data="adm:search_teacher"),
            InlineKeyboardButton(text=get_text("btn_add_new_teacher", lang), callback_data="adm:add_teacher")
        ])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))

        bc_t = {"tr": "🏠 Ana Menü ➔ 👥 Kadro ➔ 👨‍🏫 Öğretmenler", "ru": "🏠 Главное меню ➔ 👥 Ученики и учителя ➔ 👨‍🏫 Учителя", "uz": "🏠 Asosiy menyu ➔ 👥 Kadro ➔ 👨‍🏫 O'qituvchilar", "en": "🏠 Main Menu ➔ 👥 Staff ➔ 👨‍🏫 Teachers"}.get(lang, "👨‍🏫 Teachers")
        tch_roster_hdr = {
            "tr": f"👨‍🏫 <b>Öğretmenler Listesi</b> (Toplam {len(teachers)} Öğretmen):",
            "ru": f"👨‍🏫 <b>Список учителей</b> (Всего: {len(teachers)} учителей):",
            "uz": f"👨‍🏫 <b>O'qituvchilar ro'yxati</b> (Jami: {len(teachers)} ta o'qituvchi):",
            "en": f"👨‍🏫 <b>Teachers List</b> (Total: {len(teachers)} Teachers):"
        }.get(lang, "👨‍🏫 <b>Teachers List:</b>")
        tch_tap_hint = {
            "tr": "Detay veya yetki işlemleri için tıklayınız:",
            "ru": "Нажмите для просмотра профиля или прав:",
            "uz": "Tafsilotlar yoki vakolatlar uchun tanlang:",
            "en": "Tap to manage profile or permissions:"
        }.get(lang, "Tap to manage:")
        teachers_list_title = (
            f"<b>{bc_t}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{tch_roster_hdr}\n"
            f"{tch_tap_hint}"
        )
        await safe_edit_or_answer(query, teachers_list_title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

# --- ÖĞRETMEN KARTI (KUTU TASARIMI) ---
@router.callback_query(F.data.startswith("adm:tch_card:"))
async def cb_admin_teacher_card(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch:
            await query.answer(get_text("teacher_not_found", lang), show_alert=True)
            return

        u_linked = await session.get(User, tch.telegram_id) if tch.telegram_id else None
        lbl_act = {"tr": f"🟢 Aktif Bağlı (<code>{tch.telegram_id}</code>)", "ru": f"🟢 Активен / Привязан (<code>{tch.telegram_id}</code>)", "uz": f"🟢 Faol ulangan (<code>{tch.telegram_id}</code>)", "en": f"🟢 Active Connected (<code>{tch.telegram_id}</code>)"}.get(lang, f"🟢 Active (<code>{tch.telegram_id}</code>)")
        lbl_inact = {"tr": "⚪ Henüz Giriş Yapmadı", "ru": "⚪ Еще не вошел", "uz": "⚪ Hali kirmagan", "en": "⚪ Not Logged In Yet"}.get(lang, "⚪ Not Logged In Yet")
        status_str = lbl_act if tch.telegram_id else lbl_inact

        assigned_display = tch.assigned_classes or ""
        lbl_all_sch = {"tr": "🌐 TÜM OKUL (Genel)", "ru": "🌐 ВСЯ ШКОЛА (Общий)", "uz": "🌐 BUTUN MAKTAB (Umumiy)", "en": "🌐 ALL SCHOOL (General)"}.get(lang, "🌐 ALL SCHOOL")
        lbl_none_as = {"tr": "Henüz Atanmadı", "ru": "Еще не назначен", "uz": "Hali biriktirilmagan", "en": "Not Assigned Yet"}.get(lang, "Not Assigned Yet")
        if assigned_display == "ALL":
            assigned_display = lbl_all_sch
        elif not assigned_display:
            assigned_display = lbl_none_as

        bc_tc = {"tr": f"🏠 Ana Menü ➔ 👥 Kadro ➔ 👨‍🏫 {tch.full_name}", "ru": f"🏠 Главное меню ➔ 👥 Учителя ➔ 👨‍🏫 {tch.full_name}", "uz": f"🏠 Asosiy menyu ➔ 👥 Kadro ➔ 👨‍🏫 {tch.full_name}", "en": f"🏠 Main Menu ➔ 👥 Staff ➔ 👨‍🏫 {tch.full_name}"}.get(lang, f"👨‍🏫 {tch.full_name}")

        tc_title = {"tr": "👨‍🏫 <b>ÖĞRETMEN BİLGİ KARTI</b>", "ru": "👨‍🏫 <b>КАРТОЧКА УЧИТЕЛЯ</b>", "uz": "👨‍🏫 <b>O'QITUVCHI KARTASI</b>", "en": "👨‍🏫 <b>TEACHER CARD</b>"}.get(lang, "👨‍🏫 <b>TEACHER CARD</b>")
        tc_sec1 = {"tr": "📌 <b>BİLGİLER VE KİMLİK</b>", "ru": "📌 <b>ДАННЫЕ И ПРОФИЛЬ</b>", "uz": "📌 <b>MA'LUMOTLAR VA PROFIL</b>", "en": "📌 <b>INFO & IDENTITY</b>"}.get(lang, "📌 <b>INFO</b>")
        tc_sec2 = {"tr": "🏫 <b>SORUMLU SINIFLAR</b>", "ru": "🏫 <b>ЗАКРЕПЛЕННЫЕ КЛАССЫ</b>", "uz": "🏫 <b>BIRIKTIRILGAN SINFLAR</b>", "en": "🏫 <b>ASSIGNED CLASSES</b>"}.get(lang, "🏫 <b>CLASSES</b>")

        lbl_tc_name = {"tr": "İsim", "ru": "ФИО", "uz": "F.I.O", "en": "Full Name"}.get(lang, "Name")
        lbl_tc_subj = {"tr": "Branş", "ru": "Предмет", "uz": "Fani", "en": "Subject"}.get(lang, "Subject")
        lbl_tc_code = {"tr": "Giriş Kodu", "ru": "Код доступа", "uz": "Kirish kodi", "en": "Access Code"}.get(lang, "Code")
        lbl_tc_stat = {"tr": "Durum", "ru": "Статус", "uz": "Holat", "en": "Status"}.get(lang, "Status")
        lbl_tc_cls = {"tr": "Sınıflar", "ru": "Классы", "uz": "Sinflar", "en": "Classes"}.get(lang, "Classes")

        text = (
            f"<b>{bc_tc}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{tc_title}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{tc_sec1}\n"
            f"• <b>{lbl_tc_name}:</b> {escape_html(tch.full_name)}\n"
            f"• <b>{lbl_tc_subj}:</b> {escape_html(tch.subject)}\n"
            f"• <b>{lbl_tc_code}:</b> <code>{tch.auth_code}</code>\n"
            f"• <b>{lbl_tc_stat}:</b> {status_str}\n\n"
            f"{tc_sec2}\n"
            f"• <b>{lbl_tc_cls}:</b> {escape_html(assigned_display)}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = []
        if u_linked and u_linked.username:
            buttons.append([InlineKeyboardButton(text=get_text("btn_write_telegram", lang), url=f"https://t.me/{u_linked.username}")])

        buttons.append([
            InlineKeyboardButton(text=get_text("btn_manage_tch_classes", lang), callback_data=f"adm:manage_tch_cls:{tch.id}"),
            InlineKeyboardButton(text=get_text("btn_transfer_class", lang), callback_data=f"adm:xfer_cls_init:{tch.id}")
        ])
        buttons.append([
            InlineKeyboardButton(text=get_text("btn_add_co_teacher", lang), callback_data=f"adm:co_tch_init:{tch.id}"),
            InlineKeyboardButton(text=get_text("btn_reset_codes", lang), callback_data=f"adm:prompt_reset_tch:{tch.id}")
        ])
        buttons.append([InlineKeyboardButton(text=get_text("btn_del_teacher", lang), callback_data=f"adm:del_tch:{tch.id}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:teachers"))

        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:manage_tch_cls:"))
async def cb_admin_manage_teacher_classes(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch: return

        all_classes = await get_all_school_classes(session)
        if not all_classes:
            buttons = [
                [InlineKeyboardButton(text="➕ " + ("Sınıf Ekle" if lang=="tr" else ("Добавить класс" if lang=="ru" else ("Sinf qo'shish" if lang=="uz" else "Add Class"))), callback_data="adm:add_class")],
                [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"adm:tch_card:{tch.id}")]
            ]
            await safe_edit_or_answer(query, get_text("no_classes_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        assigned_raw = tch.assigned_classes or ""
        is_all = (assigned_raw == "ALL")
        assigned_set = set(all_classes) if is_all else set([c.strip() for c in assigned_raw.split(",") if c.strip()])

        bc_mc = {"tr": f"🏠 Ana Menü ➔ 👨‍🏫 {tch.full_name} ➔ 🏫 Sınıflar", "ru": f"🏠 Главное ➔ 👨‍🏫 {tch.full_name} ➔ 🏫 Классы", "uz": f"🏠 Asosiy ➔ 👨‍🏫 {tch.full_name} ➔ 🏫 Sinflar", "en": f"🏠 Main ➔ 👨‍🏫 {tch.full_name} ➔ 🏫 Classes"}.get(lang, "🏫 Classes")
        prompt_txt = (
            f"<b>{bc_mc}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"👨‍🏫 <b>{escape_md(tch.full_name)}</b> ({escape_md(tch.subject)})\n"
            f"Lütfen öğretmenin sorumlu olacağı sınıfları seçiniz (Dokunarak açıp kapatabilirsiniz):"
        )

        buttons = []
        row = []
        for c in all_classes:
            is_checked = c in assigned_set
            icon = "✅" if is_checked else "⬜"
            row.append(InlineKeyboardButton(text=f"{icon} {c}", callback_data=f"adm:toggle_tch_cls:{tch.id}:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)

        buttons.append([
            InlineKeyboardButton(text=get_text("btn_assign_all_classes", lang), callback_data=f"adm:tch_cls_all:{tch.id}"),
            InlineKeyboardButton(text=get_text("btn_clear_all_classes", lang), callback_data=f"adm:tch_cls_clear:{tch.id}")
        ])
        buttons.append([InlineKeyboardButton(text=f"💾 {get_text('btn_save_att', lang)} / Geri", callback_data=f"adm:tch_card:{tch.id}")])

        await safe_edit_or_answer(query, prompt_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:toggle_tch_cls:"))
async def cb_admin_toggle_teacher_class(query: CallbackQuery):
    parts = query.data.split(":")
    tch_id = int(parts[2])
    cls_name = parts[3]

    async with AsyncSessionLocal() as session:
        tch = await session.get(Teacher, tch_id)
        if tch:
            all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
            if tch.assigned_classes == "ALL":
                assigned_set = set(all_classes)
            else:
                assigned_set = set([c.strip() for c in (tch.assigned_classes or "").split(",") if c.strip()])

            if cls_name in assigned_set:
                assigned_set.remove(cls_name)
            else:
                assigned_set.add(cls_name)

            if len(assigned_set) == len(all_classes) and len(all_classes) > 0:
                tch.assigned_classes = "ALL"
            else:
                tch.assigned_classes = ",".join(sorted(list(assigned_set)))
            await session.commit()

    await cb_admin_manage_teacher_classes(query)

@router.callback_query(F.data.startswith("adm:tch_cls_all:"))
async def cb_admin_set_all_teacher_classes(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        tch = await session.get(Teacher, tch_id)
        if tch:
            tch.assigned_classes = "ALL"
            await session.commit()
    await cb_admin_manage_teacher_classes(query)

@router.callback_query(F.data.startswith("adm:tch_cls_clear:"))
async def cb_admin_clear_all_teacher_classes(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        tch = await session.get(Teacher, tch_id)
        if tch:
            tch.assigned_classes = ""
            await session.commit()
    await cb_admin_manage_teacher_classes(query)

@router.callback_query(F.data.startswith("adm:xfer_cls_init:"))
async def cb_admin_transfer_class_init(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch: return

        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        assigned_set = set(all_classes) if tch.assigned_classes == "ALL" else set([c.strip() for c in (tch.assigned_classes or "").split(",") if c.strip()])

        if not assigned_set:
            no_xfer_tst = {
                "tr": "Bu öğretmene ait devredilebilecek sorumlu sınıf bulunmuyor.",
                "ru": "У этого учителя нет закрепленных классов для передачи.",
                "uz": "Bu o'qituvchiga biriktirilgan sinflar mavjud emas.",
                "en": "This teacher has no assigned classes to transfer."
            }.get(lang, "No classes to transfer.")
            await query.answer(no_xfer_tst, show_alert=True)
            return

        buttons = []
        row = []
        for c in sorted(list(assigned_set)):
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:xfer_pick:{tch.id}:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"adm:tch_card:{tch.id}")])

        text = f"🔄 *{escape_md(tch.full_name)}*\n\n" + get_text("select_class_to_transfer", lang)
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:xfer_pick:"))
async def cb_admin_xfer_pick_target(query: CallbackQuery):
    parts = query.data.split(":")
    tch_id = int(parts[2])
    class_name = parts[3]

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        other_teachers = (await session.execute(select(Teacher).where(Teacher.id != tch_id).order_by(Teacher.full_name))).scalars().all()

        if not other_teachers:
            no_oth_tst = {
                "tr": "Okulda devredilebilecek başka öğretmen bulunamadı.",
                "ru": "Других учителей в школе не найдено.",
                "uz": "Maktabda boshqa o'qituvchilar topilmadi.",
                "en": "No other teachers found to transfer to."
            }.get(lang, "No other teachers found.")
            await query.answer(no_oth_tst, show_alert=True)
            return

        buttons = []
        for ot in other_teachers:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {ot.full_name} ({ot.subject})", callback_data=f"adm:xfer_target:{tch_id}:{class_name}:{ot.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"adm:xfer_cls_init:{tch_id}")])

        text = get_text("select_target_teacher", lang, class_name=escape_md(class_name))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:xfer_target:"))
async def cb_admin_xfer_execute(query: CallbackQuery):
    parts = query.data.split(":")
    from_tch_id = int(parts[2])
    class_name = parts[3]
    to_tch_id = int(parts[4])

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        from_tch = await session.get(Teacher, from_tch_id)
        to_tch = await session.get(Teacher, to_tch_id)

        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()

        from_set = set(all_classes) if from_tch.assigned_classes == "ALL" else set([c.strip() for c in (from_tch.assigned_classes or "").split(",") if c.strip()])
        from_set.discard(class_name)
        from_tch.assigned_classes = ",".join(sorted(list(from_set)))

        if to_tch.assigned_classes != "ALL":
            to_set = set([c.strip() for c in (to_tch.assigned_classes or "").split(",") if c.strip()])
            to_set.add(class_name)
            to_tch.assigned_classes = ",".join(sorted(list(to_set)))

        await session.commit()
        await query.answer(get_text("class_transfer_done", lang, class_name=class_name, teacher=to_tch.full_name), show_alert=True)
        query.data = f"adm:tch_card:{from_tch_id}"
        await cb_admin_teacher_card(query)

@router.callback_query(F.data.startswith("adm:co_tch_init:"))
async def cb_admin_co_teacher_init(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch: return

        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        assigned_set = set(all_classes) if tch.assigned_classes == "ALL" else set([c.strip() for c in (tch.assigned_classes or "").split(",") if c.strip()])

        if not assigned_set:
            no_co_tst = {
                "tr": "Bu öğretmene ait ortak yapılabilecek sınıf bulunmuyor.",
                "ru": "У этого учителя нет классов для назначения второго преподавателя.",
                "uz": "Bu o'qituvchiga tegishli sinflar mavjud emas.",
                "en": "This teacher has no classes for co-teaching."
            }.get(lang, "No classes for co-teaching.")
            await query.answer(no_co_tst, show_alert=True)
            return

        buttons = []
        row = []
        for c in sorted(list(assigned_set)):
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:co_pick:{tch.id}:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"adm:tch_card:{tch.id}")])

        text = f"➕ *{escape_md(tch.full_name)}*\n\n" + get_text("select_class_to_co_teacher", lang)
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:co_pick:"))
async def cb_admin_co_pick_target(query: CallbackQuery):
    parts = query.data.split(":")
    tch_id = int(parts[2])
    class_name = parts[3]

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        other_teachers = (await session.execute(select(Teacher).where(Teacher.id != tch_id).order_by(Teacher.full_name))).scalars().all()

        if not other_teachers:
            no_oth_co = {
                "tr": "Okulda atanabilecek başka öğretmen bulunamadı.",
                "ru": "Других учителей для назначения не найдено.",
                "uz": "Biriktirish uchun boshqa o'qituvchilar topilmadi.",
                "en": "No other teachers found to assign."
            }.get(lang, "No other teachers found.")
            await query.answer(no_oth_co, show_alert=True)
            return

        buttons = []
        for ot in other_teachers:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {ot.full_name} ({ot.subject})", callback_data=f"adm:co_target:{tch_id}:{class_name}:{ot.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"adm:co_tch_init:{tch_id}")])

        prompt_co = {"tr": f"👨‍🏫 *{escape_md(class_name)}* sınıfına hangi hoca ortak olarak atansın?", "ru": f"👨‍🏫 Кого прикрепить к классу *{escape_md(class_name)}*?", "uz": f"👨‍🏫 *{escape_md(class_name)}* sinfiga qaysi ustoz hamkor qilinsin?", "en": f"👨‍🏫 Add co-teacher for *{escape_md(class_name)}*:"}.get(lang, "Select co-teacher:")
        await safe_edit_or_answer(query, prompt_co, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:co_target:"))
async def cb_admin_co_execute(query: CallbackQuery):
    parts = query.data.split(":")
    from_tch_id = int(parts[2])
    class_name = parts[3]
    to_tch_id = int(parts[4])

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        to_tch = await session.get(Teacher, to_tch_id)

        if to_tch.assigned_classes != "ALL":
            to_set = set([c.strip() for c in (to_tch.assigned_classes or "").split(",") if c.strip()])
            to_set.add(class_name)
            to_tch.assigned_classes = ",".join(sorted(list(to_set)))

        await session.commit()
        await query.answer(get_text("class_co_teacher_done", lang, class_name=class_name, teacher=to_tch.full_name), show_alert=True)
        query.data = f"adm:tch_card:{from_tch_id}"
        await cb_admin_teacher_card(query)

@router.callback_query(F.data.startswith("adm:prompt_reset_tch:"))
async def cb_admin_prompt_reset_teacher_codes(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch: return

        text = f"👨‍🏫 *{escape_md(tch.full_name)}*\n\n" + get_text("confirm_reset_codes_prompt", lang)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_confirm_reset", lang), callback_data=f"adm:confirm_reset_tch:{tch.id}")],
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:tch_card:{tch.id}")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_reset_tch:"))
async def cb_admin_reset_tch_code(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if tch:
            if tch.telegram_id:
                old_u = await session.get(User, tch.telegram_id)
                if old_u: old_u.role = "guest"
            tch.auth_code = generate_secure_code("HCA")
            tch.is_code_burned = False
            tch.telegram_id = None
            await session.commit()
            lbl_ac_rst = {"tr": "Giriş Kodu", "ru": "Код доступа", "uz": "Kirish kodi", "en": "Access Code"}.get(lang, "Access Code")
            text = f"✅ {lbl_ac_rst}: `{tch.auth_code}`"
            buttons = [get_nav_buttons(lang, back_callback=f"adm:tch_card:{tch.id}")]
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:del_tch:"))
async def cb_admin_delete_teacher_confirm_prompt(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch:
            await query.answer(get_text("teacher_not_found", lang), show_alert=True)
            return

        text = get_text("confirm_delete_teacher_prompt", lang, name=escape_md(tch.full_name))
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_confirm_delete", lang), callback_data=f"adm:confirm_del_tch:{tch.id}")],
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:tch_card:{tch.id}")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_del_tch:"))
async def cb_admin_delete_teacher_confirmed(query: CallbackQuery):
    if await is_readonly_mode_active():
        await query.answer(get_text("readonly_mode_active_alert", "tr"), show_alert=True)
        return
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if tch:
            if tch.telegram_id:
                old_u = await session.get(User, tch.telegram_id)
                if old_u:
                    old_u.role = "guest"
                    old_u.full_name = None
            await session.execute(delete(Appointment).where(Appointment.teacher_id == tch.id))
            await session.delete(tch)
            await session.commit()
            await query.answer(get_text("teacher_deleted", lang), show_alert=True)
            await cb_admin_teachers_list(query)
            return
    await query.answer()

@router.callback_query(F.data == "adm:classes")
async def cb_classes_list(query: CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        classes = await get_all_school_classes(session)
        btn_add_cls = {"tr": "➕ Sınıf Ekle", "ru": "➕ Добавить класс", "uz": "➕ Sinf qo'shish", "en": "➕ Add Class"}.get(lang, "➕ Add Class")
        btn_add_st = get_text("btn_add_student", lang)
        btn_excel = get_text("btn_excel_hub", lang)

        if not classes:
            text = get_text("no_classes_found", lang) + "\n\n" + {
                "tr": "Yeni bir sınıf oluşturmak veya öğrenci eklemek için aşağıdaki butonları kullanabilirsiniz:",
                "ru": "Используйте кнопки ниже для создания нового класса или добавления учеников:",
                "uz": "Yangi sinf yaratish yoki o'quvchi qo'shish uchun quyidagi tugmalardan foydalaning:",
                "en": "Use the buttons below to create a new class or add students:"
            }.get(lang, "Use the buttons below to add a class or student:")
            buttons = [
                [InlineKeyboardButton(text=btn_add_cls, callback_data="adm:add_class"), InlineKeyboardButton(text=btn_add_st, callback_data="adm:add_student")],
                [InlineKeyboardButton(text=btn_excel, callback_data="adm:excel_hub")],
                get_nav_buttons(lang, back_callback="adm:cat_staff")
            ]
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        row = []
        for c in classes:
            st_cnt = (await session.execute(select(func.count(Student.id)).where(Student.class_name == c))).scalar() or 0
            row.append(InlineKeyboardButton(text=f"🏫 {c} ({st_cnt})", callback_data=f"adm:show_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append([
            InlineKeyboardButton(text=btn_add_cls, callback_data="adm:add_class"),
            InlineKeyboardButton(text=btn_add_st, callback_data="adm:add_student")
        ])
        buttons.append([InlineKeyboardButton(text=btn_excel, callback_data="adm:excel_hub")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))

        bc_c = {"tr": "🏠 Ana Menü ➔ 👥 Kadro ➔ 🏫 Sınıflar", "ru": "🏠 Главное меню ➔ 👥 Ученики и учителя ➔ 🏫 Классы", "uz": "🏠 Asosiy menyu ➔ 👥 Kadro ➔ 🏫 Sinflar", "en": "🏠 Main Menu ➔ 👥 Staff ➔ 🏫 Classes"}.get(lang, "🏫 Classes")
        sub_c = {"tr": "İncelemek istediğiniz sınıfı seçiniz veya yeni bir sınıf ekleyiniz:", "ru": "Выберите класс для просмотра или добавьте новый:", "uz": "Ko'rmoqchi bo'lgan sinfni tanlang yoki yangi sinf qo'shing:", "en": "Select a class to view or add a new class:"}.get(lang, "Select class:")
        prompt_c = (
            f"<b>{bc_c}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🏫 <b>{get_text('btn_classes', lang)}</b> (Toplam {len(classes)} Sınıf)\n\n"
            f"{sub_c}"
        )
        await safe_edit_or_answer(query, prompt_c, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "noop")
async def cb_noop(query: CallbackQuery):
    await query.answer()

@router.callback_query(F.data.startswith("adm:show_class:"))
async def cb_show_class_students(query: CallbackQuery):
    parts = query.data.split(":")
    class_name = parts[2]
    page = int(parts[3]) if len(parts) > 3 and parts[3].isdigit() else 0
    PAGE_SIZE = 10

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        all_students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()
        total_students = len(all_students)
        total_pages = max(1, (total_students + PAGE_SIZE - 1) // PAGE_SIZE)
        page = max(0, min(page, total_pages - 1))

        all_teachers = (await session.execute(select(Teacher))).scalars().all()
        class_teachers = [t for t in all_teachers if t.assigned_classes == "ALL" or class_name in [x.strip() for x in (t.assigned_classes or "").split(",")]]
        tch_names = [f"{t.full_name} ({t.subject})" for t in class_teachers]
        tch_summary = ", ".join(tch_names) if tch_names else "Henüz Atanmadı"

        bc_sc = {"tr": f"🏠 Ana Menü ➔ 👥 Kadro ➔ 🏫 {class_name}", "ru": f"🏠 Главное меню ➔ 👥 Ученики ➔ 🏫 {class_name}", "uz": f"🏠 Asosiy menyu ➔ 👥 Kadro ➔ 🏫 {class_name}", "en": f"🏠 Main Menu ➔ 👥 Staff ➔ 🏫 {class_name}"}.get(lang, f"🏫 {class_name}")
        header_roster = {
            "tr": f"🏫 <b>{escape_html(class_name)} Sınıfı Listesi</b> (Toplam {total_students} Öğrenci):",
            "ru": f"🏫 <b>Список класса {escape_html(class_name)}</b> (Всего: {total_students} уч.):",
            "uz": f"🏫 <b>{escape_html(class_name)} sinfi ro'yxati</b> (Jami: {total_students} o'quvchi):",
            "en": f"🏫 <b>Class {escape_html(class_name)} Roster</b> (Total: {total_students} Students):"
        }.get(lang, f"🏫 <b>Class {escape_html(class_name)}</b>:")

        btn_add_to_cls = {"tr": "➕ Bu Sınıfa Öğrenci Ekle", "ru": "➕ Добавить ученика в этот класс", "uz": "➕ Bu sinfga o'quvchi qo'shish", "en": "➕ Add Student to Class"}.get(lang, "➕ Add Student")
        btn_del_cls = {"tr": "🗑️ Sınıfı Sil", "ru": "🗑️ Удалить класс", "uz": "🗑️ Sinfni o'chirish", "en": "🗑️ Delete Class"}.get(lang, "🗑️ Delete Class")

        buttons = []

        if total_students == 0:
            empty_hint = {
                "tr": "ℹ️ <i>Bu sınıfta henüz kayıtlı öğrenci bulunmamaktadır.</i>",
                "ru": "ℹ️ <i>В этом классе пока нет зарегистрированных учеников.</i>",
                "uz": "ℹ️ <i>Ushbu sinfda hali o'quvchilar ro'yxatga olinmagan.</i>",
                "en": "ℹ️ <i>No students are currently registered in this class.</i>"
            }.get(lang, "No students registered.")
            buttons = [
                [InlineKeyboardButton(text=btn_add_to_cls, callback_data=f"adm:add_st_to_cls:{class_name}")],
                [InlineKeyboardButton(text=btn_del_cls, callback_data=f"adm:del_class:{class_name}")],
                get_nav_buttons(lang, back_callback="adm:classes")
            ]
            class_list_title = (
                f"<b>{bc_sc}</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"{header_roster}\n"
                f"👨‍🏫 <b>{get_text('lbl_class_teachers', lang)}:</b> <i>{escape_html(tch_summary)}</i>\n\n"
                f"{empty_hint}"
            )
        else:
            start_idx = page * PAGE_SIZE
            paged_students = all_students[start_idx : start_idx + PAGE_SIZE]

            for s in paged_students:
                buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"adm:st_card:{s.id}:{page}")])

            if total_pages > 1:
                nav_row = []
                if page > 0:
                    nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:show_class:{class_name}:{page - 1}"))
                nav_row.append(InlineKeyboardButton(text=f"📄 {page + 1}/{total_pages}", callback_data="noop"))
                if page < total_pages - 1:
                    nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:show_class:{class_name}:{page + 1}"))
                buttons.append(nav_row)

            buttons.append([
                InlineKeyboardButton(text=btn_add_to_cls, callback_data=f"adm:add_st_to_cls:{class_name}"),
                InlineKeyboardButton(text="✏️ " + {"tr": "Sınıf Adını Değiştir", "ru": "Переименовать", "uz": "Nomini o'zgartirish", "en": "Rename"}.get(lang, "Rename"), callback_data=f"adm:rename_cls:{class_name}")
            ])
            buttons.append([InlineKeyboardButton(text=f"📋 {class_name} {get_text('btn_class_att_sheet', lang)}", callback_data=f"adm:class_att_sheet:{class_name}")])
            buttons.append([InlineKeyboardButton(text=f"📄 {class_name} {get_text('btn_class_pdf_cards', lang)}", callback_data=f"adm:gen_pdf:{class_name}")])
            buttons.append(get_nav_buttons(lang, back_callback="adm:classes"))

            tap_hint = {
                "tr": "Detay veya şifre işlemleri için öğrenciye tıklayınız:",
                "ru": "Нажмите на ученика для просмотра данных или кодов:",
                "uz": "Ma'lumotlar yoki kodlar uchun o'quvchini tanlang:",
                "en": "Tap a student to manage details or credentials:"
            }.get(lang, "Tap a student:")
            class_list_title = (
                f"<b>{bc_sc}</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"{header_roster}\n"
                f"👨‍🏫 <b>{get_text('lbl_class_teachers', lang)}:</b> <i>{escape_html(tch_summary)}</i>\n\n"
                f"{tap_hint}"
            )

        await safe_edit_or_answer(query, class_list_title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:edit_st:"))
async def cb_admin_edit_student_menu(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st: return

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_edit_name", lang), callback_data=f"adm:edf:{st.id}:name")],
            [InlineKeyboardButton(text=get_text("btn_edit_class", lang), callback_data=f"adm:edf:{st.id}:class")],
            [InlineKeyboardButton(text=get_text("btn_edit_no", lang), callback_data=f"adm:edf:{st.id}:no")],
            get_nav_buttons(lang, back_callback=f"adm:st_card:{st.id}")
        ]
        prompt_es = {
            "tr": f"✏️ <b>{escape_md(st.full_name)}</b> ({escape_md(st.class_name)} - No: {escape_md(st.student_number)})\nDüzenlemek istediğiniz bilgiyi seçiniz:",
            "ru": f"✏️ <b>{escape_md(st.full_name)}</b> ({escape_md(st.class_name)} - №: {escape_md(st.student_number)})\nВыберите параметр для изменения:",
            "uz": f"✏️ <b>{escape_md(st.full_name)}</b> ({escape_md(st.class_name)} - №: {escape_md(st.student_number)})\nTahrirlamoqchi bo'lgan ma'lumotni tanlang:",
            "en": f"✏️ <b>{escape_md(st.full_name)}</b> ({escape_md(st.class_name)} - Roll: {escape_md(st.student_number)})\nSelect information to edit:"
        }.get(lang, f"✏️ <b>{escape_md(st.full_name)}</b>\n")
        await safe_edit_or_answer(query, prompt_es, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:edf:"))
async def cb_admin_edit_field_init(query: CallbackQuery, state: FSMContext):
    parts = query.data.split(":")
    st_id = int(parts[2])
    field = parts[3]
    await state.update_data(edit_st_id=st_id, edit_field=field)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    prompts = {
        "name": get_text("prompt_student_name", lang),
        "class": get_text("prompt_student_class", lang),
        "no": get_text("prompt_student_no", lang)
    }
    await safe_edit_or_answer(query, prompts.get(field, "Lütfen yeni değeri yazınız:"), reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:st_card:{st_id}")]]))
    await state.set_state(Form.edit_student_val)
    await query.answer()

@router.message(Form.edit_student_val)
async def process_student_edit_val(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    data = await state.get_data()
    st_id = data.get("edit_st_id")
    field = data.get("edit_field")
    new_val = message.text.strip()
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if st:
            if field == "name": st.full_name = new_val
            elif field == "class": st.class_name = new_val.upper()
            elif field == "no": st.student_number = new_val
            await session.commit()
            await message.answer(get_text("student_info_updated", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name), no=escape_md(st.student_number)), parse_mode="HTML")

# --- ÖĞRENCİ KARTI (KUTU TASARIMI) ---
@router.callback_query(F.data.startswith("adm:st_card:"))
async def cb_student_card(query: CallbackQuery):
    parts = query.data.split(":")
    st_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 and parts[3].isdigit() else 0
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st:
            await query.answer(get_text("student_not_found", lang), show_alert=True)
            return

        lbl_used = {"tr": "Kullanıldı", "ru": "Использован", "uz": "Ishlatildi", "en": "Used"}.get(lang, "Used")
        lbl_free = {"tr": "Aktif / Boşta", "ru": "Активен / Свободен", "uz": "Faol / Bo'sh", "en": "Active / Unused"}.get(lang, "Active")
        st_status = lbl_used if st.is_student_code_burned else lbl_free
        pr_status = lbl_used if st.is_parent_code_burned else lbl_free

        parents = (await session.execute(select(User).join(ParentStudent, ParentStudent.parent_telegram_id == User.telegram_id).where(ParentStudent.student_id == st.id))).scalars().all()
        parent_names = [f"{p.full_name or 'Veli'} (<code>{p.telegram_id}</code>)" for p in parents]
        p_info = ", ".join(parent_names) if parent_names else "Henüz bağlı veli yok"

        bc_st = {"tr": f"🏠 Ana Menü ➔ 👥 Kadro ➔ 🏫 {st.class_name} ➔ 👤 {st.full_name}", "ru": f"🏠 Главное меню ➔ 👥 Ученики ➔ 🏫 {st.class_name} ➔ 👤 {st.full_name}", "uz": f"🏠 Asosiy menyu ➔ 👥 Kadro ➔ 🏫 {st.class_name} ➔ 👤 {st.full_name}", "en": f"🏠 Main Menu ➔ 👥 Staff ➔ 🏫 {st.class_name} ➔ 👤 {st.full_name}"}.get(lang, f"👤 {st.full_name}")
        
        t_card_title = {"tr": "🧑‍🎓 <b>ÖĞRENCİ BİLGİ KARTI</b>", "ru": "🧑‍🎓 <b>КАРТОЧКА УЧЕНИКА</b>", "uz": "🧑‍🎓 <b>O'QUVCHI KARTASI</b>", "en": "🧑‍🎓 <b>STUDENT CARD</b>"}.get(lang, "🧑‍🎓 <b>STUDENT CARD</b>")
        t_id_sec = {"tr": "📌 <b>ÖĞRENCİ KİMLİĞİ</b>", "ru": "📌 <b>ЛИЧНЫЕ ДАННЫЕ УЧЕНИКА</b>", "uz": "📌 <b>O'QUVCHI SHAXSI</b>", "en": "📌 <b>STUDENT IDENTITY</b>"}.get(lang, "📌 <b>IDENTITY</b>")
        t_cd_sec = {"tr": "🔑 <b>GİRİŞ KODLARI VE DURUM</b>", "ru": "🔑 <b>КОДЫ ДОСТУПА И СTATUS</b>", "uz": "🔑 <b>KIRISH KODLARI VA HOLAT</b>", "en": "🔑 <b>ACCESS CODES & STATUS</b>"}.get(lang, "🔑 <b>CODES</b>")
        t_pr_sec = {"tr": "👨‍👩‍👧‍👦 <b>BAĞLI VELİLER</b>", "ru": "👨‍👩‍👧‍👦 <b>ПРИВЯЗАННЫЕ РОДИТЕЛИ</b>", "uz": "👨‍👩‍👧‍👦 <b>ULANGAN OTA-ONALAR</b>", "en": "👨‍👩‍👧‍👦 <b>LINKED PARENTS</b>"}.get(lang, "👨‍👩‍👧‍👦 <b>PARENTS</b>")

        lbl_name_w = {"tr": "Ad Soyad", "ru": "ФИО", "uz": "F.I.O", "en": "Full Name"}.get(lang, "Name")
        lbl_cls_w = {"tr": "Sınıf", "ru": "Класс", "uz": "Sinf", "en": "Class"}.get(lang, "Class")
        lbl_no_w = {"tr": "Okul No", "ru": "Номер в школе", "uz": "Maktab raqami", "en": "Roll No"}.get(lang, "No")
        lbl_st_code = {"tr": "Öğrenci Kodu", "ru": "Код ученика", "uz": "O'quvchi kodi", "en": "Student Code"}.get(lang, "Student Code")
        lbl_pr_code = {"tr": "Veli Kodu", "ru": "Код родителя", "uz": "Ota-ona kodi", "en": "Parent Code"}.get(lang, "Parent Code")

        text = (
            f"<b>{bc_st}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{t_card_title}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{t_id_sec}\n"
            f"• <b>{lbl_name_w}:</b> {escape_html(st.full_name)}\n"
            f"• <b>{lbl_cls_w}:</b> {escape_html(st.class_name)}\n"
            f"• <b>{lbl_no_w}:</b> <code>{escape_html(st.student_number)}</code>\n\n"
            f"{t_cd_sec}\n"
            f"• <b>{lbl_st_code}:</b> <code>{st.student_code}</code> ({st_status})\n"
            f"• <b>{lbl_pr_code}:</b> <code>{st.parent_code}</code> ({pr_status})\n\n"
            f"{t_pr_sec}\n"
            f"• {p_info}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = [
            [
                InlineKeyboardButton(text=get_text("btn_edit_student", lang), callback_data=f"adm:edit_st:{st.id}"),
                InlineKeyboardButton(text=get_text("btn_behavior", lang), callback_data=f"bh_st:{st.id}")
            ],
            [
                InlineKeyboardButton(text=get_text("btn_reset_codes", lang), callback_data=f"adm:prompt_reset_st:{st.id}"),
                InlineKeyboardButton(text=get_text("btn_unlink_parent", lang), callback_data=f"adm:unlink_pr:{st.id}")
            ],
            [InlineKeyboardButton(text=get_text("btn_del_student", lang), callback_data=f"adm:del_student:{st.id}")],
            get_nav_buttons(lang, back_callback=f"adm:show_class:{st.class_name}:{page}")
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:prompt_reset_st:"))
async def cb_prompt_reset_student_codes(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st: return

        text = f"🧑‍🎓 *{escape_md(st.full_name)}*\n\n" + get_text("confirm_reset_codes_prompt", lang)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_confirm_reset", lang), callback_data=f"adm:confirm_reset_st:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:st_card:{st.id}")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_reset_st:"))
async def cb_reset_student_codes(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if st:
            if st.student_telegram_id:
                st_u = await session.get(User, st.student_telegram_id)
                if st_u: st_u.role = "guest"
            st.student_code = generate_secure_code("OGR")
            st.parent_code = generate_secure_code("VELI")
            st.is_student_code_burned = False
            st.is_parent_code_burned = False
            st.student_telegram_id = None
            await session.commit()
            text = get_text("codes_reset_done", lang, st_code=st.student_code, pr_code=st.parent_code)
            buttons = [get_nav_buttons(lang, back_callback=f"adm:st_card:{st.id}")]
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:unlink_pr:"))
async def cb_admin_unlink_parent(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if st:
            parent_ids = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st.id))).scalars().all()
            for p_id in parent_ids:
                p_user = await session.get(User, p_id)
                if p_user:
                    other_rel = (await session.execute(select(ParentStudent.student_id).where(ParentStudent.parent_telegram_id == p_id, ParentStudent.student_id != st.id))).scalars().first()
                    p_user.current_child_id = other_rel if other_rel else None
                    if not other_rel: p_user.role = "guest"
            await session.execute(delete(ParentStudent).where(ParentStudent.student_id == st.id))
            st.parent_code = generate_secure_code("VELI")
            st.is_parent_code_burned = False
            await session.commit()
            await query.answer(get_text("parent_unlinked_success", lang, code=st.parent_code), show_alert=True)
            await cb_student_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:del_student:"))
async def cb_delete_student_confirm_prompt(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st:
            await query.answer(get_text("student_not_found", lang), show_alert=True)
            return

        text = get_text("confirm_delete_student_prompt", lang, name=escape_md(st.full_name))
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_confirm_delete", lang), callback_data=f"adm:confirm_del_st:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:st_card:{st.id}")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_del_st:"))
async def cb_delete_student_confirmed(query: CallbackQuery):
    if await is_readonly_mode_active():
        await query.answer(get_text("readonly_mode_active_alert", "tr"), show_alert=True)
        return
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if st:
            cls_name = st.class_name
            if st.student_telegram_id:
                st_u = await session.get(User, st.student_telegram_id)
                if st_u: st_u.role = "guest"

            parents_linked = (await session.execute(select(User).where(User.current_child_id == st.id))).scalars().all()
            for p_u in parents_linked:
                other_kid = (await session.execute(select(ParentStudent.student_id).where(ParentStudent.parent_telegram_id == p_u.telegram_id, ParentStudent.student_id != st.id))).scalars().first()
                p_u.current_child_id = other_kid if other_kid else None
                if not other_kid: p_u.role = "guest"

            await session.execute(delete(ParentStudent).where(ParentStudent.student_id == st.id))
            await session.execute(delete(Grade).where(Grade.student_id == st.id))
            await session.execute(delete(BehaviorRecord).where(BehaviorRecord.student_id == st.id))
            await session.execute(delete(Attendance).where(Attendance.student_id == st.id))
            await session.execute(delete(MedicalReport).where(MedicalReport.student_id == st.id))
            await session.execute(delete(HomeworkSubmission).where(HomeworkSubmission.student_id == st.id))
            await session.execute(delete(AccessRequest).where(AccessRequest.student_match_id == st.id))
            await session.delete(st)
            await session.commit()
            await query.answer(get_text("student_deleted", lang), show_alert=True)
            query.data = f"adm:show_class:{cls_name}"
            await cb_show_class_students(query)
            return
    await query.answer()

@router.callback_query(F.data.in_(["adm:dashboard", "tch:menu", "pr:menu", "st:menu", "nav:main_menu"]))
async def cb_nav_to_main_menu(query: CallbackQuery, state: FSMContext | None = None):
    if state:
        try: await state.clear()
        except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if not user:
            role = "admin" if query.from_user.id in ADMIN_IDS else "guest"
            user = User(telegram_id=query.from_user.id, language="tr", role=role, full_name=query.from_user.full_name)
            session.add(user)
            await session.commit()
        await render_clean_dashboard(query, user)
    try:
        await query.answer()
    except Exception:
        pass

@router.callback_query(F.data == "act_change_lang")
async def cb_act_change_lang(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("lang_select", lang), reply_markup=get_language_inline_kb())
    await query.answer()

@router.callback_query(F.data.startswith("ack_notif:"))
async def cb_ack_notif(query: CallbackQuery):
    notif_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        notif = await session.get(CriticalNotification, notif_id)
        if notif:
            notif.acknowledged_at = datetime.utcnow()
            await session.commit()
    await query.answer(get_text("acknowledged_toast", lang))
    if query.message:
        try:
            await query.message.edit_reply_markup(reply_markup=None)
        except Exception:
            pass

async def cb_admin_dashboard(query: CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = query.from_user.id
    chat_id = query.message.chat.id if (query and query.message) else user_id
    bot_obj = query.bot if getattr(query, 'bot', None) else (query.message.bot if query.message else bot)

    for cache_dict in [PIN_PENDING_ACTIONS, PIN_CHANGE_SESSION, ADMIN_PIN_INPUT, ADMIN_PIN_FAILURES]:
        cache_dict.pop(user_id, None)

    sub_id = PIN_SUB_MSG_ID.pop(user_id, None)
    if sub_id:
        try: await bot_obj.delete_message(chat_id=chat_id, message_id=sub_id)
        except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if user:
            # Geri tıklandığında alt menüyü (ReplyKeyboardMarkup) Telegram'a kesin olarak geri yükle!
            await render_clean_dashboard(query, user)
    await query.answer()

@router.callback_query(F.data == "adm:unack_notifs")
async def cb_admin_unack_notifs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        since = datetime.utcnow() - timedelta(hours=36)
        unacks = (await session.execute(select(CriticalNotification).where(CriticalNotification.created_at >= since, CriticalNotification.acknowledged_at == None).order_by(desc(CriticalNotification.created_at)).limit(20))).scalars().all()

        if not unacks:
            buttons = [get_nav_buttons(lang, back_callback="adm:cat_tools_reports")]
            await safe_edit_or_answer(query, get_text("all_notifs_acknowledged", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        header_u = {
            "tr": f"⚠️ *Okunmamış / Onaylanmamış Devamsızlıklar ({len(unacks)}):*\n",
            "ru": f"⚠️ *Непрочитанные уведомления о пропусках ({len(unacks)}):*\n",
            "uz": f"⚠️ *Tasdiqlanmagan davomat xabarlari ({len(unacks)}):*\n",
            "en": f"⚠️ *Unacknowledged Absence Alerts ({len(unacks)}):*\n"
        }.get(lang, f"⚠️ *Unacknowledged Alerts ({len(unacks)}):*\n")

        lines = [header_u]
        lbl_p = get_text("lbl_role_parent", lang)
        for un in unacks:
            p_user = await session.get(User, un.user_telegram_id)
            p_name = p_user.full_name if (p_user and p_user.full_name) else lbl_p
            dt_str = un.created_at.strftime("%d.%m %H:%M")
            lines.append(f"• 👤 *{escape_md(p_name)}* (`{un.user_telegram_id}`) - _{dt_str}_")

        buttons = [get_nav_buttons(lang, back_callback="adm:cat_tools_reports")]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:cockpit")
async def cb_cockpit(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = get_local_date()
        total_students = (await session.execute(select(func.count(Student.id)))).scalar() or 0
        present_count = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "present"))).scalar() or 0
        absent_count = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "absent"))).scalar() or 0

        all_classes = (await session.execute(select(Student.class_name).distinct())).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()
        missing = [c for c in all_classes if c not in taken_classes]
        all_done_txt = {"tr": "✅ Tümü Alındı", "ru": "✅ Все приняты", "uz": "✅ Hammasi olingan", "en": "✅ All Recorded"}.get(lang, "✅ All Done")
        missing_str = "\n".join([f"• ❌ <b>{escape_md(c)}</b>" for c in missing]) if missing else all_done_txt

        pct = round((present_count / total_students * 100), 1) if total_students > 0 else 100.0
        prog_bar = render_progress_bar(int(pct // 10), 10)

        bc_cp = {"tr": "🏠 Ana Menü ➔ 📊 Raporlar ➔ 🎛️ Kokpit", "ru": "🏠 Главное меню ➔ 📊 Отчеты ➔ 🎛️ Панель", "uz": "🏠 Asosiy menyu ➔ 📊 Hisobotlar ➔ 🎛️ Kokpit", "en": "🏠 Main Menu ➔ 📊 Reports ➔ 🎛️ Cockpit"}.get(lang, "🎛️ Cockpit")
        cp_title = {"tr": "📊 <b>GÜNLÜK SABAH KOKPİTİ</b>", "ru": "📊 <b>УТРЕННЯЯ СВОДКА И КОКПИТ</b>", "uz": "📊 <b>KUNLIK ERTALABKI KOKPIT</b>", "en": "📊 <b>DAILY MORNING COCKPIT</b>"}.get(lang, "📊 <b>DAILY MORNING COCKPIT</b>")
        lbl_cp_dt = {"tr": "Tarih", "ru": "Дата", "uz": "Sana", "en": "Date"}.get(lang, "Date")
        lbl_cp_rate = {"tr": "📈 <b>KATILIM ORANI</b>", "ru": "📈 <b>ПРОЦЕНТ ПОСЕЩАЕМОСТИ</b>", "uz": "📈 <b>QATNASHISH DARAJASI</b>", "en": "📈 <b>ATTENDANCE RATE</b>"}.get(lang, "📈 <b>ATTENDANCE RATE</b>")
        lbl_cp_stat = {"tr": "👥 <b>ÖĞRENCİ DURUMU</b>", "ru": "👥 <b>СТАТИСТИКА УЧЕНИКОВ</b>", "uz": "👥 <b>O'QUVCHILAR HOLATI</b>", "en": "👥 <b>STUDENT BREAKDOWN</b>"}.get(lang, "👥 <b>STUDENT BREAKDOWN</b>")
        lbl_cp_tot = {"tr": "Toplam Öğrenci", "ru": "Всего учеников", "uz": "Jami o'quvchilar", "en": "Total Students"}.get(lang, "Total Students")
        lbl_cp_prs = {"tr": "Gelen (Mevcut)", "ru": "Присутствуют", "uz": "Kelgan (Bor)", "en": "Present"}.get(lang, "Present")
        lbl_cp_abs = {"tr": "Gelmeyen (Yok)", "ru": "Отсутствуют", "uz": "Kelmagan (Yo'q)", "en": "Absent"}.get(lang, "Absent")
        lbl_cp_mis = {"tr": "⚠️ <b>YOKLAMA GİRMEYEN SINIFLAR</b>", "ru": "⚠️ <b>КЛАССЫ БЕЗ ПЕРЕКЛИЧКИ</b>", "uz": "⚠️ <b>DAVOMAT KIRITILMAGAN SINFLAR</b>", "en": "⚠️ <b>CLASSES WITHOUT ATTENDANCE</b>"}.get(lang, "⚠️ <b>CLASSES WITHOUT ATTENDANCE</b>")

        text = (
            f"<b>{bc_cp}</b>\n\n"
            f"{cp_title}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📅 <b>{lbl_cp_dt}:</b> {today.strftime('%d.%m.%Y')}\n\n"
            f"{lbl_cp_rate}\n"
            f"<code>{prog_bar}</code> <b>%{pct}</b>\n\n"
            f"{lbl_cp_stat}\n"
            f"┌ 🏫 <b>{lbl_cp_tot}:</b> {total_students}\n"
            f"├ 🟢 <b>{lbl_cp_prs}:</b> {present_count}\n"
            f"└ 🔴 <b>{lbl_cp_abs}:</b> {absent_count}\n\n"
            f"{lbl_cp_mis}\n"
            f"{missing_str}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )
        buttons = []
        if missing:
            buttons.append([InlineKeyboardButton(text=get_text("btn_remind_att", lang), callback_data="adm:remind_all_att")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_tools_reports"))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:risk_radar")
async def cb_admin_risk_radar(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        # 1. Kritik Öğrenciler (Devamsızlık >= 7 veya Not < 50)
        students = (await session.execute(select(Student).order_by(Student.class_name, Student.student_number))).scalars().all()
        risky_students = []
        for s in students:
            abs_cnt = (await session.execute(select(func.count(Attendance.id)).where(Attendance.student_id == s.id, Attendance.status == "absent"))).scalar() or 0
            grades = (await session.execute(select(Grade).where(Grade.student_id == s.id))).scalars().all()
            avg_score = (sum(g.score for g in grades) / len(grades)) if grades else 100.0
            reasons = []
            if abs_cnt >= 7: reasons.append(f"{abs_cnt} Gün Devamsızlık")
            if avg_score < 50.0: reasons.append(f"Not Ort: {round(avg_score, 1)}")
            if reasons:
                risky_students.append(f"• <b>{escape_html(s.full_name)}</b> ({escape_html(s.class_name)}): <i>{', '.join(reasons)}</i>")

        # 2. İletişim Kopuk Veliler (14+ gündür bota girmeyenler)
        now = datetime.utcnow()
        fourteen_days_ago = now - timedelta(days=14)
        inactive_parents = (await session.execute(
            select(User).where(User.role == "parent", (User.last_seen_at == None) | (User.last_seen_at < fourteen_days_ago)).limit(5)
        )).scalars().all()
        parent_risks = []
        for p in inactive_parents:
            p_name = p.full_name or f"ID: {p.telegram_id}"
            parent_risks.append(f"• <b>{escape_html(p_name)}</b>: <i>14+ gündür bota girmedi</i>")

        # 3. Geciken İşlemler (Bekleyen randevular)
        pending_apps = (await session.execute(
            select(Appointment).where(Appointment.status == "pending").limit(5)
        )).scalars().all()
        pending_risks = []
        for a in pending_apps:
            pending_risks.append(f"• Randevu #{a.id}: <i>Öğretmen onayı bekleniyor ({a.preferred_time})</i>")

        # Rapor Metni
        lines = [
            "🚨 <b>MÜDÜR DENETİM & RİSK RADARI</b>",
            "━━━━━━━━━━━━━━━━━━━━"
        ]

        if risky_students:
            lines.append(f"🔴 <b>Kritik Öğrenciler ({len(risky_students)}):</b>")
            lines.extend(risky_students[:5])
        else:
            lines.append("🔴 <b>Kritik Öğrenciler:</b> ✅ Yok")

        lines.append("")
        if parent_risks:
            lines.append(f"🟡 <b>İletişim Kopuk Veliler ({len(parent_risks)}):</b>")
            lines.extend(parent_risks[:3])
        else:
            lines.append("🟡 <b>İletişim Kopuk Veliler:</b> ✅ Tüm veliler aktif")

        lines.append("")
        if pending_risks:
            lines.append(f"🔵 <b>Geciken / Bekleyen İşlemler ({len(pending_risks)}):</b>")
            lines.extend(pending_risks[:3])
        else:
            lines.append("🔵 <b>Geciken İşlemler:</b> ✅ Bekleyen işlem yok")

        lines.append("━━━━━━━━━━━━━━━━━━━━")

        lbl_ref = "🔄 " + get_text("btn_refresh_data", lang)
        lbl_excel = "📥 " + {"tr": "Olay Günlüğü (Excel)", "ru": "Журнал аудита (Excel)", "uz": "Audit jurnali (Excel)", "en": "Audit Log (Excel)"}.get(lang, "Audit Log (Excel)")

        buttons = [
            [InlineKeyboardButton(text=lbl_ref, callback_data="adm:risk_radar"), InlineKeyboardButton(text=lbl_excel, callback_data="adm:export_audit_excel")],
            get_nav_buttons(lang, back_callback="adm:cat_tools_reports")
        ]

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:export_audit_excel")
async def cb_admin_export_audit_excel(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        logs = (await session.execute(select(AuditLog).order_by(desc(AuditLog.created_at)).limit(500))).scalars().all()

    import io
    import openpyxl
    from openpyxl.styles import Font, PatternFill, Alignment

    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Adli Olay Günlüğü"

    headers = ["ID", "Tarih / Saat", "Kullanıcı ID", "Kullanıcı Adı", "İşlem Türü", "Ayrıntılar / Detay"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")

    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for l in logs:
        dt_str = (l.created_at + timedelta(hours=TIMEZONE_OFFSET)).strftime("%d.%m.%Y %H:%M:%S")
        ws.append([l.id, dt_str, str(l.user_id), l.user_name, l.action, l.details])

    for col in ws.columns:
        max_len = max(len(str(cell.value or '')) for cell in col)
        col_letter = openpyxl.utils.get_column_letter(col[0].column)
        ws.column_dimensions[col_letter].width = min(50, max(12, max_len + 3))

    bio = io.BytesIO()
    wb.save(bio)
    bio.seek(0)

    doc_file = BufferedInputFile(bio.read(), filename=f"Adli_Olay_Gunlugu_{datetime.utcnow().strftime('%Y%m%d_%H%M')}.xlsx")

    caption_txt = {
        "tr": f"📑 <b>Okul Adli Olay Günlüğü (Excel)</b>\nSon <b>{(len(logs) if isinstance(logs, (list, tuple)) else 0)}</b> kritik işlem ve güvenlik kaydı aktarıldı.",
        "ru": f"📑 <b>Журнал аудита безопасности (Excel)</b>\nВыгружено <b>{(len(logs) if isinstance(logs, (list, tuple)) else 0)}</b> записей действий.",
        "uz": f"📑 <b>Xavfsizlik va harakatlar jurnali (Excel)</b>\nSo'nggi <b>{(len(logs) if isinstance(logs, (list, tuple)) else 0)}</b> ta yozuv yuklandi.",
        "en": f"📑 <b>Audit & Security Log (Excel)</b>\nExported <b>{(len(logs) if isinstance(logs, (list, tuple)) else 0)}</b> recent actions."
    }.get(lang, "Audit Log.")

    await query.bot.send_document(chat_id=query.message.chat.id, document=doc_file, caption=caption_txt, parse_mode="HTML")
    await query.answer("✅ Excel aktarıldı!", show_alert=False)


@router.callback_query(F.data == "adm:audit_logs")
async def cb_admin_audit_logs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        logs = (await session.execute(select(AuditLog).order_by(desc(AuditLog.created_at)).limit(12))).scalars().all()
        no_logs_txt = {"tr": "✅ Henüz kayıtlı işlem geçmişi bulunmamaktadır.", "ru": "✅ Журнал действий пуст.", "uz": "✅ Hozircha tizim jurnali bo'sh.", "en": "✅ No audit logs recorded yet."}.get(lang, "✅ No logs.")
        bc_al = {"tr": "🏠 Ana Menü ➔ 🛎️ Onay Masası ➔ 📜 İşlem Geçmişi", "ru": "🏠 Главное меню ➔ 🛎️ Центр ➔ 📜 Журнал", "uz": "🏠 Asosiy menyu ➔ 🛎️ Tasdiqlash ➔ 📜 Tizim jurnali", "en": "🏠 Main Menu ➔ 🛎️ Approvals ➔ 📜 Audit Logs"}.get(lang, "📜 Audit Logs")
        header_al = f"<b>{bc_al}</b>\n\n📜 <b>Son Sistem İşlem Geçmişi:</b>\n━━━━━━━━━━━━━━━━━━━━\n"

        if not logs:
            content = no_logs_txt
        else:
            lines = [f"📌 <b>{l.created_at.strftime('%d.%m %H:%M')}</b> - <b>{escape_md(l.action)}</b>\n↳ {escape_md(l.user_name)}: <i>{escape_md(l.details)}</i>" for l in logs]
            content = "\n\n".join(lines)

        text = f"{header_al}{content}"
        buttons = [get_nav_buttons(lang, back_callback="adm:cat_requests")]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:academic_report")
async def cb_admin_academic_report(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        rankings = []
        for c in classes:
            st_ids = (await session.execute(select(Student.id).where(Student.class_name == c))).scalars().all()
            grades = (await session.execute(select(Grade.score).where(Grade.student_id.in_(st_ids)))).scalars().all()
            if grades:
                rankings.append((c, round(sum(grades)/len(grades), 1), len(grades)))

        rankings.sort(key=lambda x: x[1], reverse=True)
        lbl_gr_cnt = {"tr": "not", "ru": "оценок", "uz": "baho", "en": "grades"}.get(lang, "grades")
        lbl_avg_word = {"tr": "Ortalama", "ru": "Средний", "uz": "O'rtacha", "en": "Average"}.get(lang, "Avg")
        lbl_cls_word = {"tr": "Sınıfı", "ru": "Класс", "uz": "sinfi", "en": "Class"}.get(lang, "Class")

        if not rankings:
            content = {"tr": "Henüz girilmiş ders notu bulunmamaktadır.", "ru": "Оценки пока не выставлены.", "uz": "Hozircha baholar kiritilmagan.", "en": "No grades recorded yet."}.get(lang, "No grades.")
        else:
            lines = [f"🥇 <b>{escape_md(c)} {lbl_cls_word}:</b> {lbl_avg_word} <b>{avg}</b> ({cnt} {lbl_gr_cnt})" if idx==1 else (f"🥈 <b>{escape_md(c)} {lbl_cls_word}:</b> {lbl_avg_word} <b>{avg}</b> ({cnt} {lbl_gr_cnt})" if idx==2 else f"{idx}. <b>{escape_md(c)} {lbl_cls_word}:</b> {lbl_avg_word} <b>{avg}</b> ({cnt} {lbl_gr_cnt})") for idx, (c, avg, cnt) in enumerate(rankings, 1)]
            content = "\n".join(lines)

        bc_ar = {"tr": "🏠 Ana Menü ➔ 📊 Raporlar ➔ 📈 Sıralama", "ru": "🏠 Главное меню ➔ 📊 Отчеты ➔ 📈 Рейтинг", "uz": "🏠 Asosiy menyu ➔ 📊 Hisobotlar ➔ 📈 Reyting", "en": "🏠 Main Menu ➔ 📊 Reports ➔ 📈 Ranking"}.get(lang, "📈 Ranking")
        header_ar = f"<b>{bc_ar}</b>\n\n📈 <b>Okul Akademik Başarı Sıralaması (Sınıflar):</b>\n━━━━━━━━━━━━━━━━━━━━\n"
        text = f"{header_ar}{content}"
        buttons = [get_nav_buttons(lang, back_callback="adm:cat_tools_reports")]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:teacher_attendance_check")
async def cb_admin_att_check(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        today = get_local_date()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()

        missing = [c for c in all_classes if c not in taken_classes]
        buttons = []
        if not missing:
            text = f"{get_text('att_check_title', lang, date=today.strftime('%d.%m.%Y'))}\n\n{get_text('att_check_all_done', lang)}"
        else:
            missing_lines = [f"• ❌ <b>{escape_md(c)}</b>" for c in missing]
            missing_title = {"tr": "⚠️ <b>Yoklama Almayan Sınıflar:</b>", "ru": "⚠️ <b>Классы без переклички:</b>", "uz": "⚠️ <b>Davomat olinmagan sinflar:</b>", "en": "⚠️ <b>Classes without attendance:</b>"}.get(lang, "⚠️ <b>Pending:</b>")
            text = (
                f"{get_text('att_check_title', lang, date=today.strftime('%d.%m.%Y'))}\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"{missing_title} ({len(missing)}):\n"
                + "\n".join(missing_lines)
            )
            buttons.append([InlineKeyboardButton(text=get_text("btn_remind_att", lang), callback_data="adm:remind_all_att")])

        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_tools_reports"))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:remind_all_att")
async def cb_admin_remind_all_att(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = get_local_date()
        all_classes = (await session.execute(select(Student.class_name).distinct())).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()
        missing = [c for c in all_classes if c not in taken_classes]

        teachers = (await session.execute(select(Teacher).where(Teacher.telegram_id.isnot(None)))).scalars().all()
        for t in teachers:
            tch_u = await session.get(User, t.telegram_id)
            tch_l = tch_u.language if tch_u else "tr"
            remind_m = {
                "tr": f"⚠️ *YOKLAMA HATIRLATMASI*\n\nSayın Hocam, *{', '.join(missing)}* sınıflarının sabah yoklaması henüz sisteme girilmemiştir. Lütfen ders yoklamasını alınız.",
                "ru": f"⚠️ *НАПОМИНАНИЕ О ПЕРЕКЛИЧКЕ*\n\nУважаемый(ая) учитель, перекличка для классов *{', '.join(missing)}* еще не внесена в систему.",
                "uz": f"⚠️ *DAVOMAT ESLATMASI*\n\nHurmatli ustoz, *{', '.join(missing)}* sinflarining tonggi davomati hali tizimga kiritilmagan. Iltimos, davomatni oling.",
                "en": f"⚠️ *ATTENDANCE REMINDER*\n\nDear Teacher, morning attendance for classes *{', '.join(missing)}* is still pending. Please take attendance."
            }.get(tch_l, f"⚠️ Attendance Reminder for: {', '.join(missing)}")
            await safe_send_message(query.message.bot, t.telegram_id, remind_m, parse_mode="HTML")
            await asyncio.sleep(0.05)

        await query.answer(get_text("remind_att_sent", lang), show_alert=True)
    await query.answer()

@router.callback_query(F.data == "adm:pdf_menu")
async def cb_pdf_menu(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            buttons = [get_nav_buttons(lang, back_callback="adm:cat_tools_reports")]
            await safe_edit_or_answer(query, get_text("no_classes_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = [[InlineKeyboardButton(text=get_text("btn_teachers_pdf", lang), callback_data="adm:gen_pdf_teachers")]]
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"📄 {c}", callback_data=f"adm:gen_pdf:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_tools_reports"))
        await safe_edit_or_answer(query, get_text("select_pdf_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:gen_pdf_teachers")
async def cb_generate_pdf_teachers(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    pdf_buffer = await generate_teachers_pdf_cards(lang=lang)
    file = BufferedInputFile(pdf_buffer.read(), filename="Ogretmen_Sifre_Kartlari.pdf")
    pdf_tch_caption = {
        "tr": "👨‍🏫 *Tüm Öğretmenlerin Giriş Şifre Kartları Ektedir.*",
        "ru": "👨‍🏫 *Карточки с кодами доступа всех учителей прикреплены к сообщению.*",
        "uz": "👨‍🏫 *Barcha o'qituvchilarning kirish parol kartalari biriktirildi.*",
        "en": "👨‍🏫 *Teacher Access Cards Attached.*"
    }.get(lang, "👨‍🏫 *Teacher Access Cards Attached.*")
    await query.message.answer_document(file, caption=pdf_tch_caption, parse_mode="HTML")
    pdf_buffer.close()
    await query.answer()

@router.callback_query(F.data.startswith("adm:gen_pdf:"))
async def cb_generate_pdf(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        pdf_buffer = await generate_classroom_pdf_cards(class_name, lang=lang)
        file = BufferedInputFile(pdf_buffer.read(), filename=f"{class_name}_Sifre_Kartlari.pdf")
        caption = get_text("pdf_ready", lang, class_name=escape_md(class_name))
        await query.message.answer_document(file, caption=caption, parse_mode="HTML")
        pdf_buffer.close()
    await query.answer()

# ======================================================================
# 12. KATEGORİ HUB'LARI VE RAPORLAR MASASI (ÇİFT YÖNLENDİRİCİ & IN-PLACE DÖNÜŞÜM)
# ======================================================================

@router.callback_query(F.data == "adm:cat_staff")
@router.message(any_state, F.text.in_(["👥 Kadro & Öğrenci", "👥 Ученики и учителя", "👥 Kadro va o'quvchilar", "👥 Staff & Students"]))
async def cb_cat_staff(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id):
            await log_audit(session, user_id, (user.full_name if user else "Bilinmeyen"), "GÜVENLİK ALARMI", f"Yetkisiz İdari Callback Erişimi Engellendi: adm:cat_staff")
            await session.commit()
            if isinstance(event, CallbackQuery): await event.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"

        bc = {"tr": "🏠 Ana Menü ➔ 👥 Kadro & Öğrenci", "ru": "🏠 Главное меню ➔ 👥 Ученики и учителя", "uz": "🏠 Asosiy menyu ➔ 👥 Kadro va o'quvchilar", "en": "🏠 Main Menu ➔ 👥 Staff & Students"}.get(lang, "👥 Staff & Students")
        cat_hdr = {"tr": "👥 <b>KADRO VE ÖĞRENCİ YÖNETİMİ</b>", "ru": "👥 <b>УПРАВЛЕНИЕ КАДРАМИ И УЧЕНИКАМИ</b>", "uz": "👥 <b>XODIMLAR VA O'QUVCHILAR BOSHQARUVI</b>", "en": "👥 <b>STAFF & STUDENT MANAGEMENT</b>"}.get(lang, "👥 <b>STAFF & STUDENTS</b>")
        cat_desc = {"tr": "📌 Öğrenci ve öğretmen kayıtlarını yönetmek, sınıf atamaları yapmak veya arama gerçekleştirmek için bir işlem seçiniz:", "ru": "📌 Выберите раздел для управления учениками, учителями или классами:", "uz": "📌 O'quvchilar, o'qituvchilar yoki sinflarni boshqarish uchun bo'limni tanlang:", "en": "📌 Select an option to manage students, teachers, classes, or promotions:"}.get(lang, "Select an option:")
        title = f"<b>{bc}</b>\n━━━━━━━━━━━━━━━━━━━━\n{cat_hdr}\n━━━━━━━━━━━━━━━━━━━━\n{cat_desc}\n━━━━━━━━━━━━━━━━━━━━"

        lbl_smart_search = {"tr": "🔍 Akıllı Kullanıcı Arama", "ru": "🔍 Умный поиск людей", "uz": "🔍 Aqlli qidiruv", "en": "🔍 Smart User Search"}.get(lang, "🔍 Smart Search")
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"), InlineKeyboardButton(text=get_text("btn_teachers", lang), callback_data="adm:teachers")],
            [InlineKeyboardButton(text=lbl_smart_search, callback_data="adm:search_user_init"), InlineKeyboardButton(text=get_text("btn_users_hub", lang), callback_data="adm:users_hub:0")],
            [InlineKeyboardButton(text=get_text("btn_school_admins", lang), callback_data="adm:admins_list"), InlineKeyboardButton(text=get_text("btn_class_promotion", lang), callback_data="adm:class_promotion_init")],
            [InlineKeyboardButton(text="🎓 " + {"tr": "Mezunlar Arşivi", "ru": "Архив выпускников", "uz": "Bitiruvchilar arxivi", "en": "Alumni Archive"}.get(lang, "Alumni"), callback_data="adm:graduates:0")]
        ]
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "adm:cat_reports")
@router.message(any_state, F.text.in_(["📊 Raporlar & Denetim", "📊 Отчеты и контроль", "📊 Hisobotlar va nazorat", "📊 Reports & Audits"]))
async def cb_cat_reports(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id):
            await log_audit(session, user_id, (user.full_name if user else "Bilinmeyen"), "GÜVENLİK ALARMI", f"Yetkisiz İdari Callback Erişimi Engellendi: adm:cat_reports")
            await session.commit()
            if isinstance(event, CallbackQuery): await event.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"

        bc = {"tr": "🏠 Ana Menü ➔ 📊 Raporlar & Denetim", "ru": "🏠 Главное меню ➔ 📊 Отчеты и контроль", "uz": "🏠 Asosiy menyu ➔ 📊 Hisobotlar va nazorat", "en": "🏠 Main Menu ➔ 📊 Reports & Audits"}.get(lang, "📊 Reports & Audits")
        cat_hdr = {"tr": "📊 <b>RAPORLAR VE AKADEMİK DENETİM</b>", "ru": "📊 <b>ОТЧЕТЫ И АКАДЕМИЧЕСКИЙ КОНТРОЛЬ</b>", "uz": "📊 <b>HISOBOTLAR VA AKADEMIK NAZORAT</b>", "en": "📊 <b>REPORTS & ACADEMIC AUDIT</b>"}.get(lang, "📊 <b>REPORTS & AUDIT</b>")
        cat_desc = {"tr": "📌 Okulun katılım kokpitini, devamsızlık risk radarını veya akademik sıralamasını incelemek için bir rapor seçiniz:", "ru": "📌 Выберите отчет для анализа посещаемости, академического рейтинга или экстренных подтверждений:", "uz": "📌 Davomat monitoringi, akademik reyting yoki hisobotlarni ko'rish uchun tanlang:", "en": "📌 Select a report to review attendance cockpit, academic rankings, or audits:"}.get(lang, "Select a report:")
        title = f"<b>{bc}</b>\n━━━━━━━━━━━━━━━━━━━━\n{cat_hdr}\n━━━━━━━━━━━━━━━━━━━━\n{cat_desc}\n━━━━━━━━━━━━━━━━━━━━"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_cockpit_unified", lang), callback_data="adm:cockpit")],
            [InlineKeyboardButton(text=get_text("btn_risk_radar", lang), callback_data="adm:risk_radar"), InlineKeyboardButton(text=get_text("btn_academic_report", lang), callback_data="adm:academic_report")],
            [InlineKeyboardButton(text=get_text("btn_unack_notifs", lang), callback_data="adm:unack_notifs"), InlineKeyboardButton(text=get_text("btn_emergency_monitor", lang), callback_data="adm:emergency_monitor")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "adm:cat_requests")
@router.message(any_state, F.text.in_(["🛎️ Onay Masası", "🛎️ Центр одобрений", "🛎️ Tasdiqlash markazi", "🛎️ Approval Center"]))
async def cb_cat_requests(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id):
            await log_audit(session, user_id, (user.full_name if user else "Bilinmeyen"), "GÜVENLİK ALARMI", f"Yetkisiz İdari Callback Erişimi Engellendi: adm:cat_requests")
            await session.commit()
            if isinstance(event, CallbackQuery): await event.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"

        req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
        med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0

        bc = {"tr": "🏠 Ana Menü ➔ 🛎️ Onay Masası", "ru": "🏠 Главное меню ➔ 🛎️ Центр одобрений", "uz": "🏠 Asosiy menyu ➔ 🛎️ Tasdiqlash markazi", "en": "🏠 Main Menu ➔ 🛎️ Approval Center"}.get(lang, "🛎️ Approval Center")
        cat_hdr = {"tr": "🛎️ <b>BAŞVURU VE RAPOR ONAY MASASI</b>", "ru": "🛎️ <b>ЦЕНТР ОДОБРЕНИЯ ЗАЯВОК И СПРАВОК</b>", "uz": "🛎️ <b>ARIZALAR VA MA'LUMOTNOMALARNI TASDIQLASH</b>", "en": "🛎️ <b>APPROVAL CENTER & MEDICALS</b>"}.get(lang, "🛎️ <b>APPROVAL CENTER</b>")
        cat_desc = {"tr": f"📌 Okula yapılan giriş şifresi başvurularını ve veli sağlık raporlarını inceleyip onaylayabilirsiniz.\n\n• 📩 <b>Bekleyen Şifre Başvurusu:</b> {req_cnt} Adet\n• 🏥 <b>Bekleyen Sağlık Raporu:</b> {med_cnt} Adet", "ru": f"📌 Рассмотрите заявки на получение кодов и медицинские справки учеников:\n\n• 📩 <b>Ожидающие заявки:</b> {req_cnt} шт.\n• 🏥 <b>Ожидающие справки:</b> {med_cnt} шт.", "uz": f"📌 Kirish kodi arizalari va tibbiy ma'lumotnomalarni ko'rib chiqing:\n\n• 📩 <b>Kutilayotgan arizalar:</b> {req_cnt} ta\n• 🏥 <b>Kutilayotgan ma'lumotnomalar:</b> {med_cnt} ta", "en": f"📌 Review and approve pending access requests and medical notes:\n\n• 📩 <b>Pending Requests:</b> {req_cnt}\n• 🏥 <b>Pending Medicals:</b> {med_cnt}"}.get(lang, "Pending items:")
        title = f"<b>{bc}</b>\n━━━━━━━━━━━━━━━━━━━━\n{cat_hdr}\n━━━━━━━━━━━━━━━━━━━━\n{cat_desc}\n━━━━━━━━━━━━━━━━━━━━"

        buttons = [
            [
                InlineKeyboardButton(text=get_text("btn_requests", lang, count=req_cnt), callback_data="adm:requests_list"),
                InlineKeyboardButton(text=get_text("btn_medical", lang, count=med_cnt), callback_data="adm:medical_list")
            ]
        ]
        if req_cnt > 1:
            btn_app_all_txt = {"tr": f"🟢 Tüm Bekleyenleri Onayla ({req_cnt})", "ru": f"🟢 Одобрить все ({req_cnt})", "uz": f"🟢 Barchasini tasdiqlash ({req_cnt})", "en": f"🟢 Approve All ({req_cnt})"}.get(lang, "🟢 Approve All")
            buttons.append([InlineKeyboardButton(text=btn_app_all_txt, callback_data="adm:approve_all_requests")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_audit_logs", lang), callback_data="adm:audit_logs")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "adm:cat_tools")
@router.message(any_state, F.text.in_(["🛠️ İdari Araçlar", "🛠️ Инструменты", "🛠️ Boshqaruv vositalari", "🛠️ Admin Tools"]))
async def cb_cat_tools(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id):
            await log_audit(session, user_id, (user.full_name if user else "Bilinmeyen"), "GÜVENLİK ALARMI", f"Yetkisiz İdari Callback Erişimi Engellendi: adm:cat_tools")
            await session.commit()
            if isinstance(event, CallbackQuery): await event.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"

        bc = {"tr": "🏠 Ana Menü ➔ 🛠️ İdari Araçlar", "ru": "🏠 Главное меню ➔ 🛠️ Инструменты", "uz": "🏠 Asosiy menyu ➔ 🛠️ Boshqaruv vositalari", "en": "🏠 Main Menu ➔ 🛠️ Admin Tools"}.get(lang, "🛠️ Admin Tools")
        cat_hdr = {"tr": "🛠️ <b>İDARİ YÖNETİM ARAÇLARI</b>", "ru": "🛠️ <b>ИНСТРУМЕНТЫ УПРАВЛЕНИЯ</b>", "uz": "🛠️ <b>BOSHQARUV VOSITALARI</b>", "en": "🛠️ <b>ADMINISTRATIVE TOOLS</b>"}.get(lang, "🛠️ <b>ADMIN TOOLS</b>")
        cat_desc = {"tr": "📌 Toplu duyuru gönderme, acil alarm tetikleme, ders programı ve yemekhane menüsü düzenleme araçlarını aşağıdan seçiniz:", "ru": "📌 Инструменты рассылки, экстренного оповещения, расписания уроков и меню столовой:", "uz": "📌 Ommaviy e'lonlar, favqulodda xabar, dars jadvali va oshxona menyusini boshqarish:", "en": "📌 Broadcasts, emergency alert, schedules, and cafeteria menu tools:"}.get(lang, "Select tool:")
        title = f"<b>{bc}</b>\n━━━━━━━━━━━━━━━━━━━━\n{cat_hdr}\n━━━━━━━━━━━━━━━━━━━━\n{cat_desc}\n━━━━━━━━━━━━━━━━━━━━"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_hub"), InlineKeyboardButton(text=get_text("btn_emergency_alert", lang), callback_data="adm:emergency_init")],
            [InlineKeyboardButton(text=get_text("btn_excel_hub", lang), callback_data="adm:excel_hub"), InlineKeyboardButton(text=get_text("btn_pdf", lang), callback_data="adm:pdf_menu")],
            [InlineKeyboardButton(text=get_text("btn_manage_schedule", lang), callback_data="adm:sched_edit_menu"), InlineKeyboardButton(text=get_text("btn_cafeteria_edit", lang), callback_data="adm:menu_edit")],
            [InlineKeyboardButton(text="🗳️ " + {"tr": "Oylama & İstişare Masası", "ru": "Голосования и предложения", "uz": "Ovoz berish va takliflar", "en": "Voting & Proposals"}.get(lang, "Voting & Proposals"), callback_data="adm:proposals_hub")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        await event.answer()


# --- DİNAMİK OKUL ADI YÖNETİMİ & VERİTABANI YEDEKLEME HANDLERLARI ---
@router.callback_query(F.data == "adm:edit_school_name")
async def cb_admin_edit_school_name(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if not is_admin_user(user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"
        sn_obj = await session.get(SystemSetting, "school_name")
        cur_name = sn_obj.value if (sn_obj and sn_obj.value) else "-"

    prompt_txt = {
        "tr": (
            "🏛️ <b>OKUL / KURUM ADI DÜZENLEME</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>Mevcut Tanım:</b> <code>{escape_html(cur_name)}</code>\n\n"
            "Lütfen sistem panolarında ve raporlarda görüntülenecek yeni okul adını yazınız:\n"
            "<i>(Varsayılana döndürmek için <b>-</b> veya <b>varsayılan</b> yazabilirsiniz)</i>"
        ),
        "ru": (
            "🏛️ <b>ИЗМЕНЕНИЕ НАЗВАНИЯ ШКОЛЫ</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>Текущее:</b> <code>{escape_html(cur_name)}</code>\n\n"
            "Введите новое название школы для отображения в отчетах и панелях:\n"
            "<i>(Для сброса введите <b>-</b>)</i>"
        ),
        "uz": (
            "🏛️ <b>MAKTAB NOMINI O'ZGARTIRISH</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>Joriy nom:</b> <code>{escape_html(cur_name)}</code>\n\n"
            "Hisobotlar va panellarda ko'rinadigan yangi maktab nomini kiriting:\n"
            "<i>(Birlamchi holatga qaytarish uchun <b>-</b> deb yozing)</i>"
        ),
        "en": (
            "🏛️ <b>EDIT SCHOOL NAME</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>Current Name:</b> <code>{escape_html(cur_name)}</code>\n\n"
            "Please enter the new school/institution name:\n"
            "<i>(Type <b>-</b> to reset to default)</i>"
        )
    }.get(lang, "Please enter new school name:")

    cancel_kb = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:cat_settings")
    ]])
    await safe_edit_or_answer(query, prompt_txt, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_school_name)
    await query.answer()

@router.message(Form.waiting_school_name)
async def process_waiting_school_name(message: Message, state: FSMContext):
    new_val = (message.text or "").strip()
    await state.clear()
    try: await message.delete()
    except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id):
            await message.answer(get_text("unauthorized_action", lang))
            return

        sn_obj = await session.get(SystemSetting, "school_name")
        if not sn_obj:
            sn_obj = SystemSetting(key="school_name", value="")
            session.add(sn_obj)

        if new_val in ["-", "varsayılan", "default", "sbros"]:
            sn_obj.value = ""
            display_saved = {"tr": "Varsayılan", "ru": "По умолчанию", "uz": "Standart", "en": "Default"}.get(lang, "Default")
        else:
            sn_obj.value = new_val[:64]
            display_saved = sn_obj.value

        await session.commit()
        await log_audit(session, message.from_user.id, (user.full_name if user else "Admin"), "AYAR GÜNCELLEME", f"Okul Adı Güncellendi: {display_saved}")
        await session.commit()

    done_msg = {
        "tr": f"✅ Okul adı başarıyla güncellendi: <b>{escape_html(display_saved)}</b>",
        "ru": f"✅ Название школы успешно обновлено: <b>{escape_html(display_saved)}</b>",
        "uz": f"✅ Maktab nomi muvaffaqiyatli saqlandi: <b>{escape_html(display_saved)}</b>",
        "en": f"✅ School name successfully updated: <b>{escape_html(display_saved)}</b>"
    }.get(lang, "School name updated successfully.")

    # 1. Update anchor header message in-place
    anchor_id = KEYBOARD_ANCHOR_MSG_ID.get(message.chat.id)
    if anchor_id:
        try:
            await message.bot.edit_message_text(
                chat_id=message.chat.id,
                message_id=anchor_id,
                text=f"🏛️ <b>{escape_html(display_saved)}</b>",
                parse_mode="HTML"
            )
        except Exception:
            pass

    # 2. Re-render settings desk in-place with zero stray messages
    await cb_cat_settings(message)

@router.callback_query(F.data == "adm:backup_db_now")
async def cb_admin_backup_db_now(query: CallbackQuery):
    user_id = query.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id):
            await query.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"

    db_path = "/working_dir/c_aa37f5c5f27a7602/school.db"
    import os
    from aiogram.types import FSInputFile
    if not os.path.exists(db_path):
        await query.answer({"tr": "❌ school.db bulunamadı!", "ru": "❌ school.db не найден!", "uz": "❌ school.db topilmadi!", "en": "❌ school.db not found!"}.get(lang, "❌ school.db not found!"), show_alert=True)
        return

    now_str = datetime.now().strftime("%Y%m%d_%H%M%S")
    doc = FSInputFile(db_path, filename=f"school_backup_{now_str}.db")
    cap = {
        "tr": f"💾 <b>GÜNCEL VERİTABANI YEDEĞİ</b>\n📅 Tarih: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n📂 Dosya: <code>school.db</code>",
        "ru": f"💾 <b>АКТУАЛЬНЫЙ БЭКАП БАЗЫ ДАННЫХ</b>\n📅 Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n📂 Файл: <code>school.db</code>",
        "uz": f"💾 <b>MA'LUMOTLAR BAZASI ZAXIRASI</b>\n📅 Sana: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n📂 Fayl: <code>school.db</code>",
        "en": f"💾 <b>DATABASE BACKUP FILE</b>\n📅 Date: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n📂 File: <code>school.db</code>"
    }.get(lang, "Database Backup File")

    try:
        await query.message.bot.send_document(chat_id=query.message.chat.id, document=doc, caption=cap, parse_mode="HTML")
        await query.answer("✅ Veritabanı yedeği sohbete gönderildi.", show_alert=False)
    except Exception as e:
        await query.answer(f"Hata: {e}", show_alert=True)

@router.callback_query(F.data == "adm:cat_settings")
@router.message(any_state, F.text.in_(["⚙️ Sistem & Ayarlar", "⚙️ Настройки системы", "⚙️ Tizim va sozlamalar", "⚙️ System & Settings"]))
async def cb_cat_settings(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id):
            await log_audit(session, user_id, (user.full_name if user else "Bilinmeyen"), "GÜVENLİK ALARMI", f"Yetkisiz İdari Callback Erişimi Engellendi: adm:cat_settings")
            await session.commit()
            if isinstance(event, CallbackQuery): await event.answer(get_text("unauthorized_action", "tr"), show_alert=True)
            return
        lang = user.language if user else "tr"

        sn_obj = await session.get(SystemSetting, "school_name")
        school_name = (sn_obj.value or "").strip() if sn_obj else ""
        display_sn = school_name if school_name else {"tr": "Belirlenmedi", "ru": "Не задано", "uz": "Kiritilmagan", "en": "Not Set"}.get(lang, "Not Set")
        maint = await session.get(SystemSetting, "maintenance_mode")
        is_maint = maint.value == "true" if maint else False

        status_lbl = {
            "tr": "AÇIK" if is_maint else "KAPALI",
            "ru": "ВКЛ" if is_maint else "ВЫКЛ",
            "uz": "TO'XTATILGAN" if is_maint else "ISHLAMOQDA",
            "en": "ON" if is_maint else "OFF"
        }.get(lang, "ON" if is_maint else "OFF")
        maint_txt = get_text("btn_maintenance_toggle", lang, status=status_lbl)

        wk_setting = await session.get(SystemSetting, "weekend_attendance_allowed")
        is_wk = wk_setting.value == "true" if wk_setting else False
        wk_lbl = "AÇIK" if is_wk else "KAPALI"
        wk_btn_txt = get_text("btn_weekend_attendance", lang, status=wk_lbl)

        bc = {"tr": "🏠 Ana Menü ➔ ⚙️ Sistem & Ayarlar", "ru": "🏠 Главное меню ➔ ⚙️ Настройки", "uz": "🏠 Asosiy menyu ➔ ⚙️ Tizim va sozlamalar", "en": "🏠 Main Menu ➔ ⚙️ System & Settings"}.get(lang, "⚙️ Settings")
        cat_hdr = {"tr": "⚙️ <b>SİSTEM VE GÜVENLİK AYARLARI</b>", "ru": "⚙️ <b>СИСТЕМНЫЕ НАСТРОЙКИ И БЕЗОПАСНОСТЬ</b>", "uz": "⚙️ <b>TIZIM VA XAVFSIZLIK SOZLAMALARI</b>", "en": "⚙️ <b>SYSTEM SETTINGS & SECURITY</b>"}.get(lang, "⚙️ <b>SETTINGS</b>")
        cat_desc = {"tr": "📌 Okul sistemi, bakım modu, saat dilimi ve sistem araçlarını bu merkezden yönetebilirsiniz:", "ru": "📌 Настройки ПИН-кода, режима обслуживания, часового пояса и безопасности:", "uz": "📌 PIN kodni o'zgartirish, texnik rejim, vaqt mintaqasi va xavfsizlik sozlamalari:", "en": "📌 Admin PIN, maintenance mode, timezone, and security controls:"}.get(lang, "Select setting:")
        title = f"<b>{bc}</b>\n━━━━━━━━━━━━━━━━━━━━\n{cat_hdr}\n━━━━━━━━━━━━━━━━━━━━\n{cat_desc}\n━━━━━━━━━━━━━━━━━━━━"

        ro_setting = await session.get(SystemSetting, "readonly_mode")
        is_ro = ro_setting.value == "true" if ro_setting else False
        ro_lbl = "AÇIK" if is_ro else "KAPALI"
        ro_btn_txt = get_text("btn_toggle_readonly", lang, status=ro_lbl)

        buttons = [
            
            [InlineKeyboardButton(text=maint_txt, callback_data="adm:toggle_maint"), InlineKeyboardButton(text=ro_btn_txt, callback_data="adm:toggle_readonly")],
            [InlineKeyboardButton(text=wk_btn_txt, callback_data="adm:toggle_weekend_att"), InlineKeyboardButton(text=get_text("btn_blacklist", lang), callback_data="adm:blacklist")],
            [InlineKeyboardButton(text=get_text("btn_clean_logs", lang), callback_data="adm:clean_old_logs"), InlineKeyboardButton(text=get_text("btn_export_all_data", lang), callback_data="adm:export_all_excel")],
            [InlineKeyboardButton(text="🏛️ " + {"tr": f"Okul Adı: {display_sn[:12]}", "ru": f"Школа: {display_sn[:12]}", "uz": f"Maktab: {display_sn[:12]}", "en": f"School: {display_sn[:12]}"}.get(lang, f"School: {display_sn[:12]}"), callback_data="adm:edit_school_name"), InlineKeyboardButton(text="💾 " + {"tr": "Veritabanı Yedeği", "ru": "Бэкап базы", "uz": "Baza nusxasi", "en": "DB Backup"}.get(lang, "DB Backup"), callback_data="adm:backup_db_now")],
            [InlineKeyboardButton(text=get_text("btn_timezone_setting", lang, offset=TIMEZONE_OFFSET), callback_data="adm:timezone_menu"), InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")],
        ]
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "adm:toggle_readonly")
async def cb_admin_toggle_readonly(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        setting = await session.get(SystemSetting, "readonly_mode")
        if not setting:
            setting = SystemSetting(key="readonly_mode", value="true")
            session.add(setting)
        else:
            setting.value = "false" if setting.value == "true" else "true"
        await session.commit()
        SETTINGS_CACHE["readonly_mode"] = setting.value
        await query.answer(get_text("readonly_mode_updated", lang), show_alert=True)
        await cb_cat_settings(query, None)

@router.callback_query(F.data == "adm:toggle_weekend_att")
async def cb_admin_toggle_weekend_att(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        wk_setting = await session.get(SystemSetting, "weekend_attendance_allowed")
        if not wk_setting:
            wk_setting = SystemSetting(key="weekend_attendance_allowed", value="true")
            session.add(wk_setting)
        else:
            wk_setting.value = "false" if wk_setting.value == "true" else "true"
        await session.commit()
        await query.answer(get_text("weekend_attendance_updated", lang), show_alert=True)
        await cb_cat_settings(query, None)

# ======================================================================
# 13. KULLANICI REHBERİ, YÖNETİCİ KADROSU VE HEDEFLİ DUYURU MASASI
# ======================================================================

@router.callback_query(F.data.startswith("adm:users_hub:"))
async def cb_admin_users_hub(query: CallbackQuery):
    page = int(query.data.split(":")[2])
    per_page = 8
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):
            await query.answer()
            return

        total_users = (await session.execute(select(func.count(User.telegram_id)))).scalar() or 0
        total_pages = max(1, (total_users + per_page - 1) // per_page)
        users_list = (await session.execute(select(User).order_by(User.created_at.desc()).offset(page * per_page).limit(per_page))).scalars().all()

        bc_uh = {"tr": "🏠 Ana Menü ➔ 👥 Kadro ➔ 👥 Kullanıcı Rehberi", "ru": "🏠 Главное меню ➔ 👥 Ученики ➔ 👥 Справочник", "uz": "🏠 Asosiy menyu ➔ 👥 Kadro ➔ 👥 Foydalanuvchilar", "en": "🏠 Main Menu ➔ 👥 Staff ➔ 👥 User Directory"}.get(lang, "👥 User Directory")
        hub_title = (
            f"<b>{bc_uh}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"👥 <b>Kullanıcı Rehberi</b> (Sayfa {page + 1}/{total_pages})\n"
            f"Toplam Kayıtlı: <b>{total_users}</b>\n\n"
            "Detay veya işlem için dokunun:\n"
        )

        lines = [hub_title]
        buttons = []
        role_icons = {"admin": "👑", "teacher": "👨‍🏫", "parent": "👨‍👩‍👧‍👦", "student": "🎓", "guest": "👤"}

        for u in users_list:
            icon = role_icons.get(u.role, "👤")
            if u.is_blacklisted: icon = "🚫"
            elif u.locked_until and u.locked_until > datetime.utcnow(): icon = "⏱️"
            elif u.admin_type == "temporary": icon = "⏱️"
            
            raw_n = u.full_name or ""
            clean_n = re.sub(r'[\.\s\-_]+', '', str(raw_n or ''))
            if not clean_n:
                u_name = get_text("permanent_admin_title", lang) if (u.telegram_id in PERMANENT_ADMIN_IDS or u.admin_type == "permanent") else get_text("lbl_role_guest", lang)
            else:
                u_name = raw_n

            u_tag = f" (@{u.username})" if u.username else ""
            btn_txt = f"{icon} {u_name}{u_tag} [{get_role_label(u.role, lang)}]"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:user_card:{u.telegram_id}:{page}")])

        nav_row = []
        if page > 0: nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:users_hub:{page - 1}"))
        if (page + 1) * per_page < total_users: nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:users_hub:{page + 1}"))
        if nav_row: buttons.append(nav_row)

        blocked_btn_txt = {"tr": "🚫 Botu Engelleyenler / Çıkanlar", "ru": "🚫 Заблокировавшие бота", "uz": "🚫 Botni bloklaganlar", "en": "🚫 Blocked Bot Users"}.get(lang, "🚫 Blocked Users")
        buttons.append([
            InlineKeyboardButton(text=get_text("btn_search_user", lang), callback_data="adm:search_user_init"),
            InlineKeyboardButton(text=blocked_btn_txt, callback_data="adm:blocked_bot_users:0")
        ])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:blocked_bot_users:"))
async def cb_admin_blocked_bot_users(query: CallbackQuery):
    page = int(query.data.split(":")[2])
    per_page = 8
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):
            await query.answer()
            return

        total_cnt = (await session.execute(select(func.count(User.telegram_id)).where(User.is_bot_blocked == True))).scalar() or 0
        total_pages = max(1, (total_cnt + per_page - 1) // per_page)
        blocked_list = (await session.execute(select(User).where(User.is_bot_blocked == True).order_by(desc(User.blocked_bot_at)).offset(page * per_page).limit(per_page))).scalars().all()

        b_title = {
            "tr": f"🚫 <b>Botu Engelleyen veya Çıkan Kullanıcılar</b> ({total_cnt})\n━━━━━━━━━━━━━━━━━━━━\n",
            "ru": f"🚫 <b>Пользователи, заблокировавшие бота</b> ({total_cnt})\n━━━━━━━━━━━━━━━━━━━━\n",
            "uz": f"🚫 <b>Botni bloklagan yoki chiqib ketgan foydalanuvchilar</b> ({total_cnt})\n━━━━━━━━━━━━━━━━━━━━\n",
            "en": f"🚫 <b>Users Who Blocked Bot or Left</b> ({total_cnt})\n━━━━━━━━━━━━━━━━━━━━\n"
        }.get(lang, "🚫 <b>Blocked Users</b>\n")

        if not blocked_list:
            no_blocked = {
                "tr": "✅ <i>Botu engelleyen veya çıkan hiçbir kullanıcı bulunmuyor.</i>",
                "ru": "✅ <i>Нет пользователей, заблокировавших бота.</i>",
                "uz": "✅ <i>Botni bloklagan hech qanday foydalanuvchi yo'q.</i>",
                "en": "✅ <i>No users have blocked the bot.</i>"
            }.get(lang, "✅ No blocked users.")
            buttons = [[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:users_hub:0")]]
            await safe_edit_or_answer(query, b_title + no_blocked, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        lines = [b_title]
        buttons = []
        for u in blocked_list:
            u_name = u.full_name or "Kullanıcı"
            u_tag = f" (@{u.username})" if u.username else ""
            t_str = u.blocked_bot_at.strftime('%d.%m.%Y %H:%M') if u.blocked_bot_at else "-"
            lines.append(f"• 👤 <b>{escape_html(u_name)}</b>{u_tag} [{get_role_label(u.role, lang)}]\n  🆔 <code>{u.telegram_id}</code> | 🕒 {t_str}\n")
            buttons.append([InlineKeyboardButton(text=f"👤 {u_name} ({get_role_label(u.role, lang)})", callback_data=f"adm:user_card:{u.telegram_id}:{page}")])

        nav_row = []
        if page > 0: nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:blocked_bot_users:{page - 1}"))
        if (page + 1) * per_page < total_cnt: nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:blocked_bot_users:{page + 1}"))
        if nav_row: buttons.append(nav_row)
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:users_hub:0")])

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:search_user_init")
async def cb_admin_search_user_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    buttons = [[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:users_hub:0")]]
    await safe_edit_or_answer(query, get_text("search_user_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.waiting_search_user_query)
    await query.answer()

@router.message(Form.waiting_search_user_query)
async def process_search_user_query(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    q_txt = message.text.strip().replace("@", "")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id):
            await message.answer(get_text("unauthorized_action", lang))
            return

        conds = [
            func.lower(User.full_name).contains(q_txt.lower()),
            func.lower(User.username).contains(q_txt.lower())
        ]
        if q_txt.isdigit(): conds.append(User.telegram_id == int(q_txt))

        res = (await session.execute(select(User).where(or_(*conds)).limit(10))).scalars().all()
        if not res:
            page = 0
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_search_again", lang), callback_data="adm:search_user_init")],
                [InlineKeyboardButton(text=get_text("btn_users_list", lang), callback_data=f"adm:users_hub:{page}")]
            ]
            await message.answer(get_text("search_user_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        if len(res) == 1:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data=f"adm:user_card:{res[0].telegram_id}")
            await cb_admin_user_card(dummy_q)
            return

        buttons = []
        for u in res:
            tag = f" (@{u.username})" if u.username else ""
            raw_n = u.full_name or ""
            clean_n = re.sub(r'[\.\s\-_]+', '', str(raw_n or ''))
            u_name = raw_n if clean_n else get_text("lbl_role_guest", lang)
            buttons.append([InlineKeyboardButton(text=f"👤 {u_name}{tag} [{get_role_label(u.role, lang)}]", callback_data=f"adm:user_card:{u.telegram_id}:{page}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_users_list", lang), callback_data=f"adm:users_hub:{page}")])
        await safe_edit_or_answer(message, get_text("search_user_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")

# --- KULLANICI PROFİL KARTI (KUTU TASARIMI) ---
@router.callback_query(F.data.startswith("adm:user_card:"))
async def cb_admin_user_card(query: CallbackQuery):
    parts = query.data.split(":")
    tg_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 and parts[3].isdigit() else 0
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id):
            await query.answer()
            return

        target_u = await session.get(User, tg_id)
        if not target_u:
            await query.answer(get_text("user_not_found_toast", lang), show_alert=True)
            return

        role_label = {
            "admin": get_text("lbl_role_admin", lang),
            "teacher": get_text("lbl_role_teacher", lang),
            "parent": get_text("lbl_role_parent", lang),
            "student": get_text("lbl_role_student", lang),
            "guest": get_text("lbl_role_guest", lang)
        }.get(target_u.role, target_u.role)

        extra_info = []
        lbl_s = get_text("lbl_subject", lang)
        lbl_cl = get_text("lbl_class", lang)
        lbl_no = get_text("lbl_number", lang)
        lbl_lk = get_text("lbl_linked_students", lang)

        if target_u.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == tg_id))).scalar_one_or_none()
            if tch: extra_info.append(f"• <b>{lbl_s}:</b> {escape_md(tch.subject)} (Kod: <code>{tch.auth_code}</code>)")
        elif target_u.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == tg_id))).scalar_one_or_none()
            if st: extra_info.append(f"• <b>{lbl_cl}:</b> {escape_md(st.class_name)} | <b>{lbl_no}:</b> {escape_md(st.student_number)}")
        elif target_u.role == "parent":
            kids = (await session.execute(select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == tg_id))).scalars().all()
            if kids:
                k_names = ", ".join([f"{k.full_name} ({k.class_name})" for k in kids])
                extra_info.append(f"• <b>{lbl_lk}:</b> {escape_md(k_names)}")

        lbl_none = {"uz": "Yo'q", "ru": "Нет", "en": "None", "tr": "Tanımlanmamış"}.get(lang, "None")
        lbl_nophone = {"uz": "Kiritilmagan", "ru": "Не указан", "en": "Not provided", "tr": "Kayıtlı Değil"}.get(lang, "Not provided")
        phone_str = target_u.phone if target_u.phone else lbl_nophone
        username_display = f"@{target_u.username}" if target_u.username else lbl_none

        if tg_id in PERMANENT_ADMIN_IDS or target_u.admin_type == "permanent":
            adm_status = "👑 <b>Kalıcı / Kurucu Yönetici</b>"
        elif target_u.admin_type == "temporary" and target_u.admin_until:
            exp_str = (target_u.admin_until + timedelta(hours=TIMEZONE_OFFSET)).strftime('%d.%m.%Y %H:%M')
            adm_status = f"⏱️ <b>Geçici Yönetici</b> (Bitiş: <code>{exp_str}</code>)"
        else:
            adm_status = get_text("lbl_not_admin", lang)

        now = datetime.utcnow()
        if target_u.is_blacklisted:
            acc_status = get_text("lbl_status_banned", lang)
        elif target_u.locked_until and target_u.locked_until > now:
            rem_m = max(1, int((target_u.locked_until - now).total_seconds() / 60))
            acc_status = f"⏱️ <b>Kilitli ({rem_m} dk)</b>"
        else:
            acc_status = "🟢 <b>Aktif</b>"

        raw_name = target_u.full_name or ""
        clean_name = re.sub(r'[\.\s\-_]+', '', str(raw_name or ''))
        display_name = raw_name if clean_name else (get_text("permanent_admin_title", lang) if (tg_id in PERMANENT_ADMIN_IDS or target_u.admin_type == "permanent") else lbl_none)

        bc_uc = {"tr": f"🏠 Ana Menü ➔ 👥 Kadro ➔ 👤 {display_name}", "ru": f"🏠 Главное меню ➔ 👥 Ученики ➔ 👤 {display_name}", "uz": f"🏠 Asosiy menyu ➔ 👥 Kadro ➔ 👤 {display_name}", "en": f"🏠 Main Menu ➔ 👥 Staff ➔ 👤 {display_name}"}.get(lang, f"👤 {display_name}")

        uc_hdr = {"tr": "👤 <b>KULLANICI PROFİL KARTI</b>", "ru": "👤 <b>ПРОФИЛЬ ПОЛЬЗОВАТЕЛЯ</b>", "uz": "👤 <b>FOYDALANUVCHI PROFILI</b>", "en": "👤 <b>USER PROFILE CARD</b>"}.get(lang, "👤 <b>USER PROFILE CARD</b>")
        lbl_u_name = {"tr": "Ad Soyad", "ru": "ФИО", "uz": "F.I.O", "en": "Full Name"}.get(lang, "Name")
        lbl_u_usr = {"tr": "Kullanıcı Adı", "ru": "Имя пользователя", "uz": "Foydalanuvchi nomi", "en": "Username"}.get(lang, "Username")
        lbl_u_phn = {"tr": "Telefon", "ru": "Телефон", "uz": "Telefon", "en": "Phone"}.get(lang, "Phone")
        lbl_u_rol = {"tr": "Sistem Rolü", "ru": "Роль в системе", "uz": "Tizimdagi roli", "en": "System Role"}.get(lang, "Role")
        lbl_u_adm = {"tr": "İdari Statü", "ru": "Статус админа", "uz": "Ma'muriy maqomi", "en": "Admin Status"}.get(lang, "Admin Status")
        lbl_u_acc = {"tr": "Hesap Durumu", "ru": "Состояние аккаунта", "uz": "Hisob holati", "en": "Account Status"}.get(lang, "Account Status")
        lbl_u_lng = {"tr": "Dil Tercihi", "ru": "Язык интерфейса", "uz": "Til tanlovi", "en": "Language Preference"}.get(lang, "Language")

        text = (
            f"<b>{bc_uc}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{uc_hdr}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"• <b>Telegram ID:</b> <code>{target_u.telegram_id}</code>\n"
            f"• <b>{lbl_u_name}:</b> {escape_html(display_name)}\n"
            f"• <b>{lbl_u_usr}:</b> {escape_html(username_display)}\n"
            f"• <b>{lbl_u_phn}:</b> <code>{escape_html(phone_str)}</code>\n"
            f"• <b>{lbl_u_rol}:</b> <b>{role_label}</b>\n"
            f"• <b>{lbl_u_adm}:</b> {adm_status}\n"
            f"• <b>{lbl_u_acc}:</b> {acc_status}\n"
            f"• <b>{lbl_u_lng}:</b> {target_u.language.upper()}\n"
        )
        if extra_info:
            text += "\n" + "\n".join(extra_info) + "\n"
        text += "━━━━━━━━━━━━━━━━━━━━"

        buttons = []
        lbl_rntgen = {"tr": "🔍 360° İstihbarat & Röntgen", "ru": "🔍 360° Досье разведки", "uz": "🔍 360° Razvedka dosyesi", "en": "🔍 360° Intelligence Dossier"}.get(lang, "🔍 360° Intelligence")
        buttons.append([InlineKeyboardButton(text=lbl_rntgen, callback_data=f"adm:user_360:{tg_id}:{page}")])

        if target_u.role != "admin" or target_u.admin_type == "none":
            buttons.append([
                InlineKeyboardButton(text=get_text("btn_make_perm_admin", lang), callback_data=f"adm:make_perm:{tg_id}"),
                InlineKeyboardButton(text=get_text("btn_make_temp_admin", lang), callback_data=f"adm:make_temp:{tg_id}")
            ])
        else:
            if tg_id not in PERMANENT_ADMIN_IDS:
                buttons.append([InlineKeyboardButton(text=get_text("btn_revoke_admin_perm", lang), callback_data=f"adm:revoke_adm:{tg_id}")])

        buttons.append([
            InlineKeyboardButton(text=get_text("btn_send_dm", lang), callback_data=f"adm:send_dm:{tg_id}"),
            InlineKeyboardButton(text=get_text("btn_req_chat", lang), callback_data=f"adm:req_chat:{tg_id}")
        ])

        if tg_id not in PERMANENT_ADMIN_IDS:
            if target_u.is_blacklisted or (target_u.locked_until and target_u.locked_until > now):
                buttons.append([InlineKeyboardButton(text=get_text("btn_unban_user", lang), callback_data=f"adm:unban_u:{tg_id}")])
            else:
                buttons.append([
                    InlineKeyboardButton(text=get_text("btn_ban_user", lang), callback_data=f"adm:ban_u:{tg_id}"),
                    InlineKeyboardButton(text=get_text("btn_temp_ban_user", lang), callback_data=f"adm:temp_ban_init:{tg_id}")
                ])

        buttons.append([
            InlineKeyboardButton(text=get_text("btn_refresh_data", lang), callback_data=f"adm:user_card:{tg_id}"),
            InlineKeyboardButton(text=get_text("btn_users_list", lang), callback_data=f"adm:users_hub:{page}")
        ])

        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()


@router.callback_query(F.data.startswith("adm:user_360:"))
async def cb_admin_user_360(query: CallbackQuery):
    parts = query.data.split(":")
    tg_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 and parts[3].isdigit() else 0

    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        target_u = await session.get(User, tg_id)
        if not target_u:
            await query.answer(get_text("user_not_found_toast", lang), show_alert=True)
            return

        raw_name = target_u.full_name or ""
        clean_name = re.sub(r'[\.\s\-_]+', '', str(raw_name or ''))
        display_name = raw_name if clean_name else (get_text("permanent_admin_title", lang) if (tg_id in PERMANENT_ADMIN_IDS or target_u.admin_type == "permanent") else "Bilinmeyen")
        usr_display = f"@{target_u.username}" if target_u.username else "Yok"
        role_label = {
            "admin": get_text("lbl_role_admin", lang),
            "teacher": get_text("lbl_role_teacher", lang),
            "parent": get_text("lbl_role_parent", lang),
            "student": get_text("lbl_role_student", lang),
            "guest": get_text("lbl_role_guest", lang)
        }.get(target_u.role, target_u.role)

        now = datetime.utcnow()
        live_status = "⚪ Henüz etkileşim yok"
        if target_u.is_bot_blocked:
            live_status = "🔴 Botu Engelledi / Çıktı"
        elif target_u.last_seen_at:
            diff_secs = (now - target_u.last_seen_at).total_seconds()
            if diff_secs < 900:
                live_status = f"🟢 Çevrimiçi ({int(diff_secs // 60)} dk önce)"
            elif diff_secs < 86400:
                live_status = f"🟡 Bugün Aktif ({int(diff_secs // 3600)} sa önce)"
            else:
                live_status = f"🔴 Pasif ({int(diff_secs // 86400)} gündür girmedi)"

        acks = (await session.execute(
            select(BroadcastAck, BroadcastNotice)
            .join(BroadcastNotice, BroadcastAck.notice_id == BroadcastNotice.id)
            .where(BroadcastAck.user_telegram_id == tg_id)
            .order_by(desc(BroadcastAck.acknowledged_at))
            .limit(5)
        )).tuples().all()

        lats = []
        if acks:
            for row in acks:
                try:
                    ack, notice = row[0], row[1]
                    if getattr(ack, 'acknowledged_at', None) and getattr(notice, 'created_at', None):
                        lats.append((ack.acknowledged_at - notice.created_at).total_seconds() / 60)
                except Exception:
                    pass

        if lats:
            avg_m = round(sum(lats) / len(lats), 1)
            if avg_m <= 5: itaat_txt = f"~{avg_m} dk (A Sınıfı - Anında Onay)"
            elif avg_m <= 30: itaat_txt = f"~{avg_m} dk (B Sınıfı - Standart)"
            else: itaat_txt = f"~{avg_m} dk (C Sınıfı - Gecikmeli / İhmal)"
        else:
            itaat_txt = "Henüz onay kaydı yok"

        score = 100
        if target_u.is_blacklisted:
            score = 0
        else:
            score -= (target_u.failed_attempts or 0) * 15
            score -= (target_u.night_activity_count or 0) * 10
            if target_u.is_bot_blocked: score -= 25
            score = max(0, min(100, score))

        if score >= 85:
            score_txt = f"%{score} (A Sınıfı - Çok Güvenilir)"
            score_badge = "🛡️"
        elif score >= 60:
            score_txt = f"%{score} (B Sınıfı - Standart)"
            score_badge = "🟡"
        elif score >= 40:
            score_txt = f"%{score} (C Sınıfı - Takipte)"
            score_badge = "🟠"
        else:
            score_txt = f"%{score} (D Sınıfı - Yüksek Risk)"
            score_badge = "🔴"

        night_cnt = target_u.night_activity_count or 0
        if night_cnt > 0:
            rutin_txt = f"⚠️ Gece {night_cnt} şüpheli işlem tespit edildi! (01:00-06:00)"
        else:
            rutin_txt = "Gündüz Rutini (Anomali Yok)"

        names_log = []
        try: names_log = json.loads(target_u.previous_names) if target_u.previous_names else []
        except Exception: pass
        if names_log:
            last_n = names_log[-1]
            alias_txt = f"{len(names_log)} Değişiklik ({last_n.get('old')} ➔ {last_n.get('new')})"
        else:
            alias_txt = "Sabit (Hiç isim/kullanıcı adı değiştirmedi)"

        p_map = {}
        try: p_map = json.loads(target_u.persona_counts) if target_u.persona_counts else {}
        except Exception: pass
        top_act = max(p_map, key=p_map.get) if p_map else "standart"
        if top_act in ("grade", "exam"): persona_txt = "🎯 Başarı & Not Odaklı"
        elif top_act in ("attendance", "excuse"): persona_txt = "⚠️ Devamsızlık & İzin Odaklı"
        elif top_act == "menu": persona_txt = "🍲 Yemek & Genel Bilgi Odaklı"
        elif top_act == "download": persona_txt = "📑 Evrak & Dosya Odaklı"
        else: persona_txt = "💤 Standart / Asgari Temas"

        dl_cnt = target_u.download_count or 0
        if dl_cnt >= 10: dl_txt = f"🚨 Yüksek Hacim ({dl_cnt} dosya indirildi!)"
        else: dl_txt = f"Normal ({dl_cnt} indirme)"

        last_act = target_u.last_action_desc or "Kayıtlı son işlem yok"

        titles = {
            "tr": ("🔍 <b>360° KULLANICI İSTİHBARAT RÖNTGENİ</b>", "Güven Skoru", "Canlı Durum", "İtaat Hızı", "Biyolojik Saat", "Kimlik Geçmişi", "Tipoloji", "Veri Güvenliği", "Son Hareket"),
            "ru": ("🔍 <b>360° ДОСЬЕ РАЗВЕДКИ И БЕЗОПАСНОСТИ</b>", "Индекс доверия", "Онлайн-статус", "Скорость отклика", "Биоритм/Часы", "История смены имени", "Типология", "Безопасность данных", "Посл. действие"),
            "uz": ("🔍 <b>360° RAZVEDKA VA XAVFSIZLIK DOSYESI</b>", "Ishonch indeksi", "Jonli holat", "Itoat tezligi", "Bioritm/Soat", "Ism o'zgarishi", "Tipologiya", "Ma'lumot xavfsizligi", "Oxirgi amal"),
            "en": ("🔍 <b>360° USER INTELLIGENCE DOSSIER</b>", "Trust Index", "Live Status", "Response Latency", "Circadian Rhythm", "Identity History", "Typology", "Data Safety", "Last Action")
        }.get(lang, ("🔍 <b>360° USER INTELLIGENCE DOSSIER</b>", "Trust Index", "Live Status", "Response Latency", "Circadian Rhythm", "Identity History", "Typology", "Data Safety", "Last Action"))

        t_hdr, t_scr, t_liv, t_lat, t_bio, t_ali, t_tip, t_dat, t_act = titles

        card_text = (
            f"<b>{t_hdr}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"👤 <b>{escape_html(display_name)}</b> ({usr_display}) | ID: <code>{tg_id}</code>\n"
            f"Rol: <b>{role_label}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{score_badge} <b>{t_scr}:</b> <b>{score_txt}</b>\n"
            f"• <b>{t_liv}:</b> {live_status}\n"
            f"• <b>{t_lat}:</b> {itaat_txt}\n"
            f"• <b>{t_bio}:</b> {rutin_txt}\n"
            f"• <b>{t_ali}:</b> {escape_html(alias_txt)}\n"
            f"• <b>{t_tip}:</b> {persona_txt}\n"
            f"• <b>{t_dat}:</b> {dl_txt}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"👣 <b>{t_act}:</b> <i>{escape_html(last_act)}</i>\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        btn_kill = {"tr": "⛔ Karantina / Kodu Yak", "ru": "⛔ Изоляция / Сжечь код", "uz": "⛔ Karantin / Kodni bekor qilish", "en": "⛔ Quarantine / Burn Code"}.get(lang, "⛔ Quarantine")
        btn_dm = get_text("btn_send_dm", lang)
        btn_prof = {"tr": "👤 Standart Profil", "ru": "👤 Обычный профиль", "uz": "👤 Oddiy profil", "en": "👤 Standard Profile"}.get(lang, "👤 Standard Profile")
        btn_back = get_text("btn_users_list", lang)

        buttons = [
            [InlineKeyboardButton(text=btn_kill, callback_data=f"adm:kill_switch:{tg_id}:{page}"), InlineKeyboardButton(text=btn_dm, callback_data=f"adm:send_dm:{tg_id}")],
            [InlineKeyboardButton(text=btn_prof, callback_data=f"adm:user_card:{tg_id}:{page}"), InlineKeyboardButton(text=btn_back, callback_data=f"adm:users_hub:{page}")]
        ]

        await safe_edit_or_answer(query, card_text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:kill_switch:"))
async def cb_admin_kill_switch(query: CallbackQuery):
    parts = query.data.split(":")
    tg_id = int(parts[2])
    page = int(parts[3]) if len(parts) > 3 and parts[3].isdigit() else 0

    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        target_u = await session.get(User, tg_id)
        if not target_u:
            await query.answer(get_text("user_not_found_toast", lang), show_alert=True)
            return

        target_u.is_blacklisted = True
        target_u.locked_until = datetime.utcnow() + timedelta(days=3650)
        target_u.role = "guest"
        target_u.previous_role = "banned"

        if target_u.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == tg_id))).scalar_one_or_none()
            if tch:
                tch.is_code_burned = True
                tch.telegram_id = None
        elif target_u.role == "parent":
            ps_list = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == tg_id))).scalars().all()
            for ps in ps_list:
                await session.delete(ps)
        elif target_u.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == tg_id))).scalar_one_or_none()
            if st:
                st.student_telegram_id = None
                st.is_student_code_burned = True

        session.add(AuditLog(
            user_id=query.from_user.id,
            user_name=admin_u.full_name or "Admin",
            action="KILL_SWITCH",
            details=f"Kullanıcı ID {tg_id} ({target_u.full_name}) tek tıkla karantinaya alındı, kodları yakıldı ve engellendi."
        ))

        await session.commit()
        FAST_USER_CACHE.pop(tg_id, None)

    toast_txt = {
        "tr": "⛔ Kullanıcı anında karantinaya alındı, kodları yakıldı ve erişimi kesildi!",
        "ru": "⛔ Пользователь изолирован, коды сожжены, доступ заблокирован!",
        "uz": "⛔ Foydalanuvchi karantinga olindi, kodlari bekor qilindi!",
        "en": "⛔ User quarantined, codes burned, and access terminated!"
    }.get(lang, "User quarantined.")
    await query.answer(toast_txt, show_alert=True)
    await cb_admin_user_360(query)


@router.callback_query(F.data.startswith("adm:temp_ban_init:"))
async def cb_admin_temp_ban_init(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
    buttons = [
        [InlineKeyboardButton(text=get_text("btn_dur_1h", lang), callback_data=f"adm:set_tban:{tg_id}:1h"), InlineKeyboardButton(text=get_text("btn_dur_24h", lang), callback_data=f"adm:set_tban:{tg_id}:24h")],
        [InlineKeyboardButton(text=get_text("btn_dur_7d", lang), callback_data=f"adm:set_tban:{tg_id}:7d"), InlineKeyboardButton(text=get_text("btn_dur_30d", lang), callback_data=f"adm:set_tban:{tg_id}:30d")],
        [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:user_card:{tg_id}")]
    ]
    await safe_edit_or_answer(query, get_text("temp_ban_choose_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:set_tban:"))
async def cb_admin_set_temp_ban(query: CallbackQuery):
    parts = query.data.split(":")
    tg_id = int(parts[2])
    dur = parts[3]
    hours_map = {"1h": 1, "24h": 24, "7d": 24 * 7, "30d": 24 * 30}
    until = datetime.utcnow() + timedelta(hours=hours_map.get(dur, 24))

    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if tg_id in PERMANENT_ADMIN_IDS:
            await query.answer(get_text("permanent_admin_protected", lang), show_alert=True)
            return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.locked_until = until
            await session.commit()
            t_lang = target_u.language
            await safe_send_message(query.message.bot, tg_id, get_text("user_temp_banned_notification", t_lang, dur=dur), parse_mode="HTML")
            await query.answer(get_text("user_temp_banned_toast", lang, dur=dur), show_alert=True)
            await cb_admin_user_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:make_perm:"))
async def cb_admin_make_perm_admin(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.previous_role = target_u.role if target_u.role != "admin" else "guest"
            target_u.role = "admin"
            target_u.admin_type = "permanent"
            target_u.admin_until = None
            await session.commit()

            t_lang = target_u.language
            notify_msg = get_text("admin_promoted_notification", t_lang, name=escape_md(target_u.full_name or get_text("lbl_role_guest", t_lang)))
            await safe_send_message(query.message.bot, tg_id, notify_msg, reply_markup=get_role_reply_kb("admin", t_lang), parse_mode="HTML")
            await query.answer(get_text("perm_admin_assigned_toast", lang), show_alert=True)
            await cb_admin_user_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:make_temp:"))
async def cb_admin_choose_temp_admin(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
    buttons = [
        [InlineKeyboardButton(text=get_text("btn_dur_24h", lang), callback_data=f"adm:set_temp:{tg_id}:24h"), InlineKeyboardButton(text=get_text("btn_dur_7d", lang), callback_data=f"adm:set_temp:{tg_id}:7d")],
        [InlineKeyboardButton(text=get_text("btn_dur_30d", lang), callback_data=f"adm:set_temp:{tg_id}:30d"), InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:user_card:{tg_id}")]
    ]
    await safe_edit_or_answer(query, get_text("temp_admin_choose_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:set_temp:"))
async def cb_admin_set_temp_admin(query: CallbackQuery):
    parts = query.data.split(":")
    tg_id = int(parts[2])
    dur = parts[3]
    hours_map = {"24h": 24, "7d": 24 * 7, "30d": 24 * 30}
    until = datetime.utcnow() + timedelta(hours=hours_map.get(dur, 24))

    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.previous_role = target_u.role if target_u.role != "admin" else "guest"
            target_u.role = "admin"
            target_u.admin_type = "temporary"
            target_u.admin_until = until
            await session.commit()

            t_lang = target_u.language
            exp_str = (until + timedelta(hours=TIMEZONE_OFFSET)).strftime('%d.%m.%Y %H:%M')
            notify_msg = f"⏱️ *{escape_md(target_u.full_name or '')}*\n\n" + get_text("temp_admin_assigned_toast", t_lang, dur=dur) + f" (Bitiş: `{exp_str}`)"
            await safe_send_message(query.message.bot, tg_id, notify_msg, reply_markup=get_role_reply_kb("admin", t_lang), parse_mode="HTML")
            await query.answer(get_text("temp_admin_assigned_toast", lang, dur=dur), show_alert=True)
            await cb_admin_user_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:revoke_adm:"))
async def cb_admin_revoke_admin(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if tg_id in PERMANENT_ADMIN_IDS:
            await query.answer(get_text("permanent_admin_protected", lang), show_alert=True)
            return
        if not is_admin_user(admin_u, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.role = target_u.previous_role or "guest"
            target_u.admin_type = "none"
            target_u.admin_until = None
            await session.commit()
            t_lang = target_u.language
            await safe_send_message(query.message.bot, tg_id, get_text("admin_demoted_notification", t_lang), reply_markup=get_role_reply_kb(target_u.role, t_lang), parse_mode="HTML")
            await query.answer(get_text("admin_demoted_toast", lang), show_alert=True)
            await cb_admin_user_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:ban_u:"))
async def cb_admin_ban_user(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if tg_id in PERMANENT_ADMIN_IDS:
            await query.answer(get_text("permanent_admin_protected", lang), show_alert=True)
            return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.is_blacklisted = True
            await session.commit()
            await safe_send_message(query.message.bot, tg_id, get_text("auth_blacklisted", target_u.language if target_u else "tr"), parse_mode="HTML")

        await query.answer(get_text("user_banned_toast", lang), show_alert=True)
        await cb_admin_user_card(query)

@router.callback_query(F.data.startswith("adm:unban_u:"))
async def cb_admin_unban_user(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.is_blacklisted = False
            target_u.failed_attempts = 0
            target_u.locked_until = None
            await session.commit()
            await safe_send_message(query.message.bot, tg_id, get_text("admin_unban_notification", target_u.language if target_u else "tr"), parse_mode="HTML")

        await query.answer(get_text("user_unbanned_toast", lang), show_alert=True)
        await cb_admin_user_card(query)

@router.callback_query(F.data.startswith("adm:send_dm:"))
async def cb_admin_send_dm_init(query: CallbackQuery, state: FSMContext):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        target_u = await session.get(User, tg_id)
        target_name = target_u.full_name if (target_u and target_u.full_name) else f"ID: {tg_id}"
    await state.update_data(target_dm_id=tg_id)
    buttons = [[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"adm:user_card:{tg_id}")]]
    await safe_edit_or_answer(query, get_text("send_dm_prompt", lang, id=tg_id, name=escape_md(target_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.waiting_admin_dm_text)
    await query.answer()

@router.message(Form.waiting_admin_dm_text)
async def process_admin_dm_text(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    dm_text = message.text.strip()
    data = await state.get_data()
    target_id = data.get("target_dm_id")
    await state.clear()

    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, message.from_user.id)
        adm_lang = admin_u.language if admin_u else "tr"
        admin_name = admin_u.full_name if (admin_u and admin_u.full_name) else get_text("school_admin_title", adm_lang)

        target_u = await session.get(User, target_id) if target_id else None
        target_lang = target_u.language if target_u else "tr"

        header_txt = get_text("dm_from_admin_header", target_lang)
        sender_lbl = get_text("dm_sender_label", target_lang)
        formatted_msg = f"{header_txt}\n\n{escape_md(dm_text)}\n\n👤 _{sender_lbl}: {escape_md(admin_name)}_"
        res = await safe_send_message(message.bot, target_id, formatted_msg, parse_mode="HTML")
        if res:
            await message.answer(get_text("dm_sent_success", adm_lang, id=target_id), reply_markup=get_role_reply_kb("admin", adm_lang), parse_mode="HTML")
        else:
            await message.answer(get_text("dm_delivery_error", adm_lang), reply_markup=get_role_reply_kb("admin", adm_lang), parse_mode="HTML")

        await render_clean_dashboard(message, admin_u)

@router.callback_query(F.data.startswith("adm:req_chat:"))
async def cb_admin_req_contact(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        adm_lang = admin_u.language if admin_u else "tr"
        admin_name = admin_u.full_name if (admin_u and admin_u.full_name) else get_text("school_admin_title", adm_lang)

        target_u = await session.get(User, tg_id)
        target_lang = target_u.language if target_u else "tr"

        chat_msg = get_text("contact_req_header", target_lang, name=escape_md(admin_name))
        buttons = []
        if admin_u and admin_u.username:
            chat_msg += get_text("contact_req_direct", target_lang)
            buttons.append([InlineKeyboardButton(text=get_text("btn_write_to_admin", target_lang), url=f"https://t.me/{admin_u.username}")])
        else:
            chat_msg += f"ID: `{query.from_user.id}`\n" + get_text("contact_req_id", target_lang)

        kb = InlineKeyboardMarkup(inline_keyboard=buttons) if buttons else None
        res = await safe_send_message(query.message.bot, tg_id, chat_msg, reply_markup=kb, parse_mode="HTML")
        if res:
            await query.answer(get_text("chat_req_sent_toast", adm_lang), show_alert=True)
        else:
            await query.answer(get_text("chat_req_error_toast", adm_lang), show_alert=True)

# --- YÖNETİCİ KADROSU MASASI (ADM:ADMINS_LIST) ---
@router.callback_query(F.data == "adm:admins_list")
async def cb_admin_admins_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        db_admins = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
        admin_map = {u.telegram_id: u for u in db_admins}

        bc_al = {"tr": "🏠 Ana Menü ➔ 👥 Kadro ➔ 👑 Yöneticiler", "ru": "🏠 Главное меню ➔ 👥 Ученики ➔ 👑 Администраторы", "uz": "🏠 Asosiy menyu ➔ 👥 Kadro ➔ 👑 Ma'murlar", "en": "🏠 Main Menu ➔ 👥 Staff ➔ 👑 Administrators"}.get(lang, "👑 Administrators")
        hub_title = {"tr": "👨‍💼 <b>OKUL YÖNETİM KADROSU</b>", "ru": "👨‍💼 <b>РУКОВОДСТВО ШКОЛЫ</b>", "uz": "👨‍💼 <b>MAKTAB RAHBARIYATI</b>", "en": "👨‍💼 <b>SCHOOL LEADERSHIP & ADMINS</b>"}.get(lang, "👨‍💼 <b>SCHOOL LEADERSHIP</b>")

        founder_count = len(PERMANENT_ADMIN_IDS)
        other_admins = [u for u in db_admins if u.telegram_id not in PERMANENT_ADMIN_IDS]
        other_count = len(other_admins)

        desc_text = {
            "tr": f"📌 Sistemde yetkilendirilmiş okul idarecileri aşağıda listelenmiştir. Yetkilerini düzenlemek, süre tanımlamak veya profillerini görüntülemek için ilgili yöneticiye dokununuz.\n\n👑 <b>Kalıcı Kurucu Yöneticiler:</b> {founder_count} Kişi\n👨‍💼 <b>Ek Görevli Yöneticiler:</b> {other_count} Kişi",
            "ru": f"📌 Ниже представлены администраторы школы. Нажмите на нужного администратора для просмотра профиля или изменения полномочий:\n\n👑 <b>Постоянные основатели:</b> {founder_count} чел.\n👨‍💼 <b>Назначенные администраторы:</b> {other_count} чел.",
            "uz": f"📌 Maktab ma'murlari quyida keltirilgan. Vakolatlarni tahrirlash yoki profilni ko'rish uchun tegishli ma'murni tanlang:\n\n👑 <b>Doimiy asosiy ma'murlar:</b> {founder_count} kishi\n👨‍💼 <b>Qo'shimcha ma'murlar:</b> {other_count} kishi",
            "en": f"📌 School administrators are listed below. Tap an administrator to manage permissions, assign temporary duration, or view details:\n\n👑 <b>Permanent Founders:</b> {founder_count}\n👨‍💼 <b>Staff Admins:</b> {other_count}"
        }.get(lang, "School administrators are listed below:")

        card_text = (
            f"<b>{bc_al}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{hub_title}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{desc_text}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = []

        # 1. Kalıcı Kurucular
        for p_id in PERMANENT_ADMIN_IDS:
            u_obj = admin_map.get(p_id)
            if not u_obj:
                u_obj = await session.get(User, p_id)
            if not u_obj or not u_obj.full_name or u_obj.full_name in ("Kalıcı İdareci", "Kurucu İdareci", "Asosiy ma'mur", "Главный админ", "Yönetici", "İsimsiz"):
                try:
                    chat_info = await query.bot.get_chat(p_id)
                    tg_name = chat_info.full_name or chat_info.title or chat_info.first_name
                    if tg_name:
                        if not u_obj:
                            u_obj = User(telegram_id=p_id, role="admin", full_name=tg_name, admin_type="permanent", language=lang)
                            session.add(u_obj)
                            admin_map[p_id] = u_obj
                        else:
                            u_obj.full_name = tg_name
                        await session.commit()
                except Exception: pass

            name_str = u_obj.full_name if (u_obj and u_obj.full_name) else get_text("permanent_admin_title", lang)
            btn_txt = f"👑 {name_str} ({p_id})"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:user_card:{p_id}")])

        # 2. Diğer Yöneticiler
        all_other_admin_ids = set(ADMIN_IDS) - set(PERMANENT_ADMIN_IDS)
        for extra_id in all_other_admin_ids:
            if extra_id not in admin_map:
                u_extra = await session.get(User, extra_id)
                if u_extra: admin_map[extra_id] = u_extra

        for tg_id, u_obj in admin_map.items():
            if tg_id not in PERMANENT_ADMIN_IDS:
                name_str = u_obj.full_name if (u_obj and u_obj.full_name) else f"ID: {tg_id}"
                badge = "⏱️" if (u_obj.admin_type == "temporary" and u_obj.admin_until) else "👨‍💼"
                btn_txt = f"{badge} {name_str} ({tg_id})"
                buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:user_card:{tg_id}")])

        buttons.append([InlineKeyboardButton(text=get_text("btn_add_admin_id", lang), callback_data="adm:add_admin_id")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_gen_admin_code", lang), callback_data="adm:gen_admin_code")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:cat_staff")])

        await safe_edit_or_answer(query, card_text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == 'adm:gen_admin_code')
async def cb_admin_gen_code(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        code_val = generate_secure_code("ADM")
        adm_key = AdminKey(code=code_val, created_by=query.from_user.id)
        session.add(adm_key)
        await session.commit()

        text = get_text("admin_code_generated", lang, code=code_val)
        buttons = [get_nav_buttons(lang, back_callback="adm:admins_list")]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:add_admin_id")
async def cb_admin_add_id_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:cat_settings")
    await safe_edit_or_answer(query, get_text("admin_add_tg_id_prompt", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_admin_tg_id)
    await query.answer()

@router.message(Form.waiting_admin_tg_id)
async def process_admin_tg_id(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    val = message.text.strip().replace("@", "")
    if not val.isdigit():
        await message.answer(get_text("admin_invalid_tg_id", lang))
        return
    await state.update_data(new_admin_id=int(val))
    await message.answer(get_text("admin_add_name_prompt", lang), parse_mode="HTML")
    await state.set_state(Form.waiting_admin_name)

@router.message(Form.waiting_admin_name)
async def process_admin_name(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    name_val = message.text.strip()
    data = await state.get_data()
    new_admin_id = data.get("new_admin_id")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        target_u = await session.get(User, new_admin_id)
        if not target_u:
            target_u = User(telegram_id=new_admin_id, language="tr", role="admin", full_name=name_val, admin_type="permanent")
            session.add(target_u)
        else:
            target_u.role = "admin"
            target_u.full_name = name_val
            target_u.admin_type = "permanent"
            target_u.failed_attempts = 0
            target_u.locked_until = None
        await session.commit()

        t_lang = target_u.language if target_u else "tr"
        await safe_send_message(message.bot, new_admin_id, get_text("admin_promoted_notification", t_lang, name=escape_md(name_val)), reply_markup=get_role_reply_kb("admin", t_lang), parse_mode="HTML")
        await message.answer(get_text("admin_added_success", lang, name=escape_md(name_val), id=new_admin_id), reply_markup=get_role_reply_kb("admin", lang), parse_mode="HTML")
        await render_clean_dashboard(message, await session.get(User, message.from_user.id))

@router.callback_query(F.data.startswith("adm:del_admin:"))
async def cb_admin_delete_admin(query: CallbackQuery):
    target_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        if target_id in PERMANENT_ADMIN_IDS:
            await query.answer(get_text("permanent_admin_protected", lang), show_alert=True)
            return

        target_u = await session.get(User, target_id)
        if target_u:
            target_u.role = "guest"
            await session.commit()
            await safe_send_message(query.message.bot, target_id, get_text("admin_demoted_notification", target_u.language if target_u else "tr"), reply_markup=get_role_reply_kb("guest", target_u.language), parse_mode="HTML")

        await query.answer(get_text("admin_demoted_toast", lang), show_alert=True)
        await cb_admin_admins_list(query)

# --- GELİŞMİŞ VE HEDEFLİ TOPLU DUYURU MASASI ---
@router.callback_query(F.data == "adm:broadcast_hub")
async def cb_admin_broadcast_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        bc_bh = {"tr": "🏠 Ana Menü ➔ 🛠️ İdari Araçlar ➔ 📢 Toplu Duyuru", "ru": "🏠 Главное меню ➔ 🛠️ Инструменты ➔ 📢 Центр рассылки", "uz": "🏠 Asosiy menyu ➔ 🛠️ Boshqaruv ➔ 📢 E'lonlar", "en": "🏠 Main Menu ➔ 🛠️ Admin Tools ➔ 📢 Broadcast Hub"}.get(lang, "📢 Broadcast Hub")
        text = f"<b>{bc_bh}</b>\n━━━━━━━━━━━━━━━━━━━━\n" + get_text("broadcast_hub_title", lang)
        buttons = [
            [InlineKeyboardButton(text=get_text("bc_target_all", lang), callback_data="adm:bc_target:all")],
            [
                InlineKeyboardButton(text=get_text("bc_target_teachers", lang), callback_data="adm:bc_target:teachers"),
                InlineKeyboardButton(text=get_text("bc_target_parents", lang), callback_data="adm:bc_target:parents")
            ],
            [
                InlineKeyboardButton(text=get_text("bc_target_students", lang), callback_data="adm:bc_target:students"),
                InlineKeyboardButton(text=get_text("bc_target_class", lang), callback_data="adm:bc_target:class")
            ],
            [InlineKeyboardButton(text=get_text("bc_target_lang", lang), callback_data="adm:bc_target:lang")]
        ]
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_tools_reports"))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:bc_target:"))
async def cb_admin_bc_target_selected(query: CallbackQuery, state: FSMContext):
    target_type = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        if target_type == "class":
            classes = await get_all_school_classes(session)
            if not classes:
                buttons = [[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:broadcast_hub")]]
                await safe_edit_or_answer(query, get_text("no_classes_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
                await query.answer()
                return
            buttons = []
            row = []
            for c in classes:
                row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:bc_class:{c}"))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row: buttons.append(row)
            buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:broadcast_hub")])
            await safe_edit_or_answer(query, get_text("bc_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        elif target_type == "lang":
            counts = {}
            for l_code in ["tr", "uz", "ru", "en"]:
                cnt = (await session.execute(select(func.count(User.telegram_id)).where(User.language == l_code))).scalar() or 0
                counts[l_code] = cnt
            buttons = [
                [
                    InlineKeyboardButton(text=f"🇹🇷 Türkçe ({counts['tr']})", callback_data="adm:bc_lang:tr"),
                    InlineKeyboardButton(text=f"🇺🇿 O'zbekcha ({counts['uz']})", callback_data="adm:bc_lang:uz")
                ],
                [
                    InlineKeyboardButton(text=f"🇷🇺 Русский ({counts['ru']})", callback_data="adm:bc_lang:ru"),
                    InlineKeyboardButton(text=f"🇬🇧 English ({counts['en']})", callback_data="adm:bc_lang:en")
                ],
                [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:broadcast_hub")]
            ]
            await safe_edit_or_answer(query, get_text("bc_select_lang", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        target_names = {
            "all": get_text("bc_target_all", lang),
            "teachers": get_text("bc_target_teachers", lang),
            "parents": get_text("bc_target_parents", lang),
            "students": get_text("bc_target_students", lang)
        }
        target_display = target_names.get(target_type, target_type)
        BC_CACHE[query.from_user.id] = {"type": target_type, "filter": None, "display": target_display}

        buttons = [[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:broadcast_hub")]]
        await safe_edit_or_answer(query, get_text("prompt_broadcast_content", lang, target=target_display), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        await state.set_state(Form.waiting_broadcast_text)
    await query.answer()

@router.callback_query(F.data.startswith("adm:bc_class:"))
async def cb_admin_bc_class_selected(query: CallbackQuery, state: FSMContext):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    target_display = f"{class_name} ({get_text('bc_target_class', lang)})"
    BC_CACHE[query.from_user.id] = {"type": "class", "filter": class_name, "display": target_display}

    buttons = [[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:broadcast_hub")]]
    await safe_edit_or_answer(query, get_text("prompt_broadcast_content", lang, target=target_display), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.waiting_broadcast_text)
    await query.answer()

@router.callback_query(F.data.startswith("adm:bc_lang:"))
async def cb_admin_bc_lang_selected(query: CallbackQuery, state: FSMContext):
    chosen_lang = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    lang_labels = {"tr": "🇹🇷 Türkçe", "ru": "🇷🇺 Русский", "uz": "🇺🇿 O'zbekcha", "en": "🇬🇧 English"}
    target_display = lang_labels.get(chosen_lang, chosen_lang)
    BC_CACHE[query.from_user.id] = {"type": "lang", "filter": chosen_lang, "display": target_display}

    buttons = [[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:broadcast_hub")]]
    await safe_edit_or_answer(query, get_text("prompt_broadcast_content", lang, target=target_display), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.waiting_broadcast_text)
    await query.answer()

@router.message(Form.waiting_broadcast_text)
async def process_broadcast_text(message: Message, state: FSMContext):
    b_text = message.caption or message.text or ""
    photo_id = message.photo[-1].file_id if message.photo else None
    bc_data = BC_CACHE.pop(message.from_user.id, {"type": "all", "filter": None, "display": "Genel"})
    await state.clear()

    if not b_text and not photo_id:
        return

    safe_text = html.escape(b_text)
    sent_cnt = 0

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        target_type = bc_data.get("type", "all")
        target_filter = bc_data.get("filter")

        target_ids = []
        if target_type == "all":
            target_ids = (await session.execute(select(User.telegram_id))).scalars().all()
        elif target_type == "teachers":
            target_ids = (await session.execute(select(User.telegram_id).where(User.role == "teacher"))).scalars().all()
        elif target_type == "parents":
            target_ids = (await session.execute(select(User.telegram_id).where(User.role == "parent"))).scalars().all()
        elif target_type == "students":
            target_ids = (await session.execute(select(User.telegram_id).where(User.role == "student"))).scalars().all()
        elif target_type == "class":
            st_ids = (await session.execute(select(Student.id).where(Student.class_name == target_filter))).scalars().all()
            p_ids = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id.in_(st_ids)))).scalars().all()
            s_ids = (await session.execute(select(Student.student_telegram_id).where(Student.id.in_(st_ids), Student.student_telegram_id.isnot(None)))).scalars().all()
            target_ids = list(set(p_ids + s_ids))
        elif target_type == "lang":
            target_ids = (await session.execute(select(User.telegram_id).where(User.language == target_filter))).scalars().all()

        notice = BroadcastNotice(content=b_text)
        session.add(notice)
        await session.commit()

        for u_id in set(target_ids):
            try:
                b_read = {"tr": "👀 Okudum / Bilgim Var", "ru": "👀 Прочитано", "uz": "👀 O'qidim / Xabardorman", "en": "👀 Read / Acknowledged"}.get(lang, "👀 Read / Acknowledged")
                read_btn = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=b_read, callback_data=f"ack_bc:{notice.id}")]])
                if photo_id:
                    await message.bot.send_photo(chat_id=u_id, photo=photo_id, caption=f"📢 <b>OKUL DUYURUSU</b>\n\n{safe_text}", reply_markup=read_btn, parse_mode="HTML")
                else:
                    await message.bot.send_message(chat_id=u_id, text=f"📢 <b>OKUL DUYURUSU</b>\n\n{safe_text}", reply_markup=read_btn, parse_mode="HTML")
                sent_cnt += 1
                await asyncio.sleep(0.04)
            except Exception:
                pass

                blocked_cnt = max(0, total_recipients - sent_cnt)
        report_card = {
            "tr": (
                "📢 <b>DUYURU İLETİM RAPORU</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ <b>Başarıyla İletilen:</b> {sent_cnt} kullanıcı\n"
                f"🚫 <b>Ulaşılamayan / Engelleyen:</b> {blocked_cnt} kullanıcı\n"
                f"👥 <b>Toplam Hedef Kitle:</b> {total_recipients} kullanıcı\n"
                f"📅 <b>Tarih:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                "━━━━━━━━━━━━━━━━━━━━"
            ),
            "ru": (
                "📢 <b>ОТЧЕТ О ДОСТАВКЕ ОБЪЯВЛЕНИЯ</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ <b>Успешно доставлено:</b> {sent_cnt}\n"
                f"🚫 <b>Не доставлено (блок):</b> {blocked_cnt}\n"
                f"👥 <b>Всего получателей:</b> {total_recipients}\n"
                f"📅 <b>Дата:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                "━━━━━━━━━━━━━━━━━━━━"
            ),
            "uz": (
                "📢 <b>E'LON YUBORISH HISOBOTI</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ <b>Yetkazildi:</b> {sent_cnt}\n"
                f"🚫 <b>Yetkazilmadi (bloklangan):</b> {blocked_cnt}\n"
                f"👥 <b>Jami qabul qiluvchilar:</b> {total_recipients}\n"
                f"📅 <b>Sana:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                "━━━━━━━━━━━━━━━━━━━━"
            ),
            "en": (
                "📢 <b>BROADCAST DELIVERY REPORT</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"✅ <b>Successfully Delivered:</b> {sent_cnt} users\n"
                f"🚫 <b>Unreachable / Blocked:</b> {blocked_cnt} users\n"
                f"👥 <b>Total Target Audience:</b> {total_recipients} users\n"
                f"📅 <b>Date:</b> {datetime.now().strftime('%d.%m.%Y %H:%M')}\n"
                "━━━━━━━━━━━━━━━━━━━━"
            )
        }.get(lang, f"Dispatched: {sent_cnt}/{total_recipients}")
        await message.answer(report_card, parse_mode="HTML")
        await render_clean_dashboard(message, user)

# --- SINIF ATLATMA (DÖNEM SONU TERFİ) - PIN KORUMALI & OTOMATİK EXCEL YEDEKLEMELİ ---
@router.callback_query(F.data == "adm:class_promotion_init")
async def cb_admin_class_promotion_init(query: CallbackQuery, state: FSMContext | None = None):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        buttons = [
            [InlineKeyboardButton(text={"tr": "✅ Evet, Sınıfları Yükselt (PIN Gerekir)", "ru": "✅ Да, перевести классы (Требуется ПИН)", "uz": "✅ Ha, sinflarni ko'chirish (PIN kerak)", "en": "✅ Yes, Promote Classes (PIN Required)"}.get(lang, "✅ Promote Classes"), callback_data="adm:class_promotion_pin_prompt")]
        ]
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(query, get_text("promotion_confirm_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:class_promotion_pin_prompt")
async def cb_admin_class_promotion_pin_prompt(query: CallbackQuery, state: FSMContext | None = None):
    await cb_admin_class_promotion_execute(query)

@router.callback_query(F.data == 'adm:class_promotion_confirm')
async def cb_admin_class_promotion_confirm(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        target_chat_id = query.message.chat.id if query.message else query.from_user.id
        bot_obj = query.bot if getattr(query, 'bot', None) else (query.message.bot if query.message else bot)
        try:
            buf = await export_all_school_data_excel()
            today_str = datetime.utcnow().strftime("%d_%m_%Y")
            file = BufferedInputFile(buf.read(), filename=f"Terfi_Oncesi_Yedek_{today_str}.xlsx")
            caption = {"tr": "📦 <b>Sınıf Atlatma Öncesi Otomatik Tam Sistem Yedeği</b>", "ru": "📦 <b>Автоматический архив системы перед переводом классов</b>", "uz": "📦 <b>Sinf ko'chirishdan oldingi avtomatik to'liq tizim zaxirasi</b>", "en": "📦 <b>Pre-Promotion Automatic Full System Backup</b>"}.get(lang, "📦 <b>Automatic Pre-Promotion System Backup</b>")
            await bot_obj.send_document(chat_id=target_chat_id, document=file, caption=caption, parse_mode="HTML")
            buf.close()
        except Exception:
            pass

        students = (await session.execute(select(Student))).scalars().all()
        promoted_cnt = 0
        for s in students:
            m = re.match(r"^(\d+)(.*)$", s.class_name.strip())
            if m:
                grade_num = int(m.group(1))
                suffix = m.group(2)
                if grade_num >= 12:
                    s.class_name = f"MEZUN{suffix}"
                else:
                    s.class_name = f"{grade_num + 1}{suffix}"
                promoted_cnt += 1

        await session.commit()
        await safe_edit_or_answer(query, get_text("promotion_success", lang, count=promoted_cnt), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang, back_callback="adm:cat_staff")]), parse_mode="HTML")
    await query.answer()

# --- SİSTEM LOG TEMİZLİĞİ & ZAMAN DİLİMİ ---
@router.callback_query(F.data == "adm:clean_old_logs")
async def cb_admin_clean_old_logs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        cutoff = datetime.utcnow() - timedelta(days=90)
        res = await session.execute(delete(AuditLog).where(AuditLog.created_at < cutoff))
        cleaned_cnt = res.rowcount or 0
        await session.commit()

        await query.answer(get_text("logs_cleaned_toast", lang, count=cleaned_cnt), show_alert=True)
        await cb_cat_settings(query, None)

@router.callback_query(F.data == "adm:timezone_menu")
async def cb_admin_timezone_menu(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        bc_tz = {"tr": "🏠 Ana Menü ➔ ⚙️ Ayarlar ➔ 🕒 Zaman Dilimi", "ru": "🏠 Главное меню ➔ ⚙️ Настройки ➔ 🕒 Часовой пояс", "uz": "🏠 Asosiy menyu ➔ ⚙️ Sozlamalar ➔ 🕒 Vaqt mintaqasi", "en": "🏠 Main Menu ➔ ⚙️ Settings ➔ 🕒 Timezone"}.get(lang, "🕒 Timezone")
        buttons = [
            [
                InlineKeyboardButton(text="UTC+3 (İstanbul)", callback_data="adm:set_tz:3"),
                InlineKeyboardButton(text="UTC+4 (Bakü/Tiflis)", callback_data="adm:set_tz:4")
            ],
            [
                InlineKeyboardButton(text="UTC+5 (Taşkent)", callback_data="adm:set_tz:5"),
                InlineKeyboardButton(text="UTC+6 (Almatı)", callback_data="adm:set_tz:6")
            ]
        ]
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_settings"))
        prompt_tz = f"<b>{bc_tz}</b>\n━━━━━━━━━━━━━━━━━━━━\n" + {"tr": "🕒 <b>Zaman Dilimi Seçiniz:</b>", "ru": "🕒 <b>Выберите часовой пояс:</b>", "uz": "🕒 <b>Vaqt mintaqasini tanlang:</b>", "en": "🕒 <b>Select Timezone:</b>"}.get(lang, "🕒 <b>Timezone:</b>")
        await safe_edit_or_answer(query, prompt_tz, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:set_tz:"))
async def cb_admin_set_tz(query: CallbackQuery):
    offset = int(query.data.split(":")[2])
    global TIMEZONE_OFFSET
    TIMEZONE_OFFSET = offset
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tz_setting = await session.get(SystemSetting, "timezone_offset")
        if not tz_setting:
            session.add(SystemSetting(key="timezone_offset", value=str(offset)))
        else:
            tz_setting.value = str(offset)
        await session.commit()

        await query.answer(get_text("timezone_updated", lang, offset=offset), show_alert=True)
        await cb_cat_settings(query, None)

@router.callback_query(F.data.startswith("ack_bc:"))
async def cb_ack_broadcast(query: CallbackQuery):
    notice_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        existing = (await session.execute(select(BroadcastAck).where(BroadcastAck.notice_id == notice_id, BroadcastAck.user_telegram_id == query.from_user.id))).scalar_one_or_none()
        if not existing:
            session.add(BroadcastAck(notice_id=notice_id, user_telegram_id=query.from_user.id))
            await session.commit()
    not_ack_tst = {"tr": "✅ Duyuruyu okuduğunuz kaydedildi.", "ru": "✅ Вы подтвердили прочтение объявления.", "uz": "✅ E'lonni o'qiganingiz qayd etildi.", "en": "✅ Announcement marked as read."}.get(lang, "Marked as read.")
    await query.answer(not_ack_tst, show_alert=False)
    await query.message.edit_reply_markup(reply_markup=None)

# ======================================================================
# 14. ÖĞRETMEN, VELİ VE ÖĞRENCİ İŞLEMLERİ (NOT, DAVRANIŞ, ÖDEV, SINAV, ACİL DURUM)
# ======================================================================

@router.callback_query(F.data == "tch:classes")
async def cb_teacher_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()

        if tch and tch.assigned_classes and tch.assigned_classes != "ALL":
            allowed = [c.strip() for c in tch.assigned_classes.split(",") if c.strip()]
            classes = [c for c in all_classes if c in allowed]
            if not classes:
                buttons = [get_nav_buttons(lang, back_callback="adm:dashboard")]
                await safe_edit_or_answer(query, get_text("no_assigned_classes_teacher", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
                await query.answer()
                return
        else:
            classes = all_classes

        if len(classes) == 1:
            query.data = f"att_class:{classes[0]}"
            await cb_attendance_class(query)
            return

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"att_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        await safe_edit_or_answer(query, get_text("attendance_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("att_class:"))
async def cb_attendance_class(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.full_name))).scalars().all()

        today = get_local_date()
        today_att = (await session.execute(select(Attendance).where(Attendance.class_name == class_name, Attendance.date == today))).scalars().all()
        today_map = {a.student_id: (a.status == "absent") for a in today_att}

        ATTENDANCE_CACHE[query.from_user.id] = {s.id: today_map.get(s.id, False) for s in students}
        kb = get_attendance_grid_kb(students, ATTENDANCE_CACHE[query.from_user.id], class_name, lang=lang)
        text = get_text("attendance_intro", lang, class_name=escape_md(class_name))
        await safe_edit_or_answer(query, text, reply_markup=kb, parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("att_toggle:"))
async def cb_attendance_toggle(query: CallbackQuery):
    student_id = int(query.data.split(":")[1])
    cache = ATTENDANCE_CACHE.get(query.from_user.id, {})
    cache[student_id] = not cache.get(student_id, False)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, student_id)
        students = (await session.execute(select(Student).where(Student.class_name == st.class_name).order_by(Student.full_name))).scalars().all()

        kb = get_attendance_grid_kb(students, cache, st.class_name, lang=lang)
        await query.message.edit_reply_markup(reply_markup=kb)
    await query.answer()

@router.callback_query(F.data.startswith("att_save:"))
async def cb_attendance_save(query: CallbackQuery):
    if await is_readonly_mode_active():
        await query.answer(get_text("readonly_mode_active_alert", "tr"), show_alert=True)
        return
    class_name = query.data.split(":")[1]
    now_local = get_local_now()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        wk_setting = await session.get(SystemSetting, "weekend_attendance_allowed")
        is_weekend_allowed = (wk_setting and wk_setting.value == "true")

        if now_local.weekday() in (5, 6) and not is_weekend_allowed:
            await query.answer(get_text("attendance_weekend_lock", lang), show_alert=True)
            return

        if not (7 <= now_local.hour < 21):
            await query.answer(get_text("attendance_hours_lock", lang), show_alert=True)
            return

        cache = ATTENDANCE_CACHE.pop(query.from_user.id, {})
        now = datetime.utcnow()
        att_date = get_local_date()
        notify_time = now + timedelta(minutes=15)

        for st_id, is_absent in cache.items():
            new_status = "absent" if is_absent else "present"
            existing = (await session.execute(select(Attendance).where(Attendance.student_id == st_id, Attendance.date == att_date))).scalar_one_or_none()
            if existing:
                was_notified_absent = (existing.status == "absent" and existing.is_notified)
                if existing.status != "excused":
                    existing.status = new_status
                    existing.notify_at = notify_time
                    existing.is_notified = False

                if was_notified_absent and new_status in ("present", "excused"):
                    st_obj = await session.get(Student, st_id)
                    st_n = escape_md(st_obj.full_name) if st_obj else "Öğrenci"
                    parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
                    for p_id in parents:
                        p_u = await session.get(User, p_id)
                        p_lang = p_u.language if p_u else "tr"
                        corr_text = get_text("attendance_correction_notification", p_lang, name=st_n)
                        await safe_send_message(query.message.bot, p_id, corr_text, parse_mode="HTML")
            else:
                att = Attendance(student_id=st_id, class_name=class_name, date=att_date, status=new_status, teacher_id=query.from_user.id, notify_at=notify_time, is_notified=False)
                session.add(att)

        await session.commit()
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, get_text("att_saved", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("att_all_pres:"))
async def cb_attendance_all_present(query: CallbackQuery):
    if await is_readonly_mode_active():
        await query.answer(get_text("readonly_mode_active_alert", "tr"), show_alert=True)
        return
    class_name = query.data.split(":")[1]
    now_local = get_local_now()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        wk_setting = await session.get(SystemSetting, "weekend_attendance_allowed")
        is_weekend_allowed = (wk_setting and wk_setting.value == "true")

        if now_local.weekday() in (5, 6) and not is_weekend_allowed:
            await query.answer(get_text("attendance_weekend_lock", lang), show_alert=True)
            return

        if not (7 <= now_local.hour < 21):
            await query.answer(get_text("attendance_hours_lock", lang), show_alert=True)
            return

        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.full_name))).scalars().all()

        now = datetime.utcnow()
        att_date = get_local_date()
        notify_time = now + timedelta(minutes=15)

        for s in students:
            existing = (await session.execute(select(Attendance).where(Attendance.student_id == s.id, Attendance.date == att_date))).scalar_one_or_none()
            if existing:
                if existing.status != "excused":
                    existing.status = "present"
                    existing.notify_at = notify_time
                    existing.is_notified = False
            else:
                att = Attendance(
                    student_id=s.id,
                    class_name=class_name,
                    date=att_date,
                    status="present",
                    teacher_id=query.from_user.id,
                    notify_at=notify_time,
                    is_notified=False
                )
                session.add(att)

        await session.commit()
        ATTENDANCE_CACHE.pop(query.from_user.id, None)

        buttons = [get_nav_buttons(lang)]
        success_all_txt = {
            "tr": f"✅ <b>{escape_html(class_name)} Sınıfı Yoklaması Tamamlandı!</b>\n━━━━━━━━━━━━━━━━━━━━\nTüm öğrenciler (<b>{len(students)} kişi</b>) eksiksiz olarak <b>VAR</b> kaydedildi.",
            "ru": f"✅ <b>Перекличка класса {escape_html(class_name)} завершена!</b>\n━━━━━━━━━━━━━━━━━━━━\nВсе ученики (<b>{len(students)} чел.</b>) отмечены как <b>ПРИСУТСТВУЮЩИЕ</b>.",
            "uz": f"✅ <b>{escape_html(class_name)} sinfi davomati yakunlandi!</b>\n━━━━━━━━━━━━━━━━━━━━\nBarcha o'quvchilar (<b>{len(students)} nafar</b>) to'liq <b>BOR</b> deb belgilandi.",
            "en": f"✅ <b>Class {escape_html(class_name)} Attendance Saved!</b>\n━━━━━━━━━━━━━━━━━━━━\nAll students (<b>{len(students)} total</b>) marked as <b>PRESENT</b>."
        }.get(lang, f"✅ <b>Attendance Saved!</b>\nAll students marked as <b>PRESENT</b>.")
        await safe_edit_or_answer(query, success_all_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        await query.answer("✅", show_alert=False)

@router.callback_query(F.data == "tch:grade_classes")
async def cb_grade_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if tch and tch.assigned_classes and tch.assigned_classes != "ALL":
            allowed = [c.strip() for c in tch.assigned_classes.split(",") if c.strip()]
            classes = [c for c in all_classes if c in allowed] or all_classes
        else:
            classes = all_classes

        buttons = [[InlineKeyboardButton(text=get_text("btn_recent_grades_menu", lang), callback_data="tch:recent_grades")]]
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"gr_cls:{c}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(query, get_text("grade_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:grade_sheet:"))
async def cb_teacher_class_grade_sheet(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        subject_name = tch.subject if tch else "Ders"

        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()
        if not students:
            await query.answer(get_text("no_students_in_class", lang), show_alert=True)
            return

        gr_sheet_title = {
            "tr": f"📊 <b>{escape_md(class_name)} Sınıfı {escape_md(subject_name)} Not Çizelgesi:</b>\n━━━━━━━━━━━━━━━━━━━━\n",
            "ru": f"📊 <b>Ведомость оценок класса {escape_md(class_name)} ({escape_md(subject_name)}):</b>\n━━━━━━━━━━━━━━━━━━━━\n",
            "uz": f"📊 <b>{escape_md(class_name)} sinfining {escape_md(subject_name)} fanidan baholar qaydnomasi:</b>\n━━━━━━━━━━━━━━━━━━━━\n",
            "en": f"📊 <b>Grade Sheet for Class {escape_md(class_name)} ({escape_md(subject_name)}):</b>\n━━━━━━━━━━━━━━━━━━━━\n"
        }.get(lang, f"📊 <b>Grade Sheet: {escape_md(class_name)}</b>\n")

        no_gr_str = {"tr": "<i>Not girilmedi</i>", "ru": "<i>Оценок нет</i>", "uz": "<i>Baho qo'yilmagan</i>", "en": "<i>No grades recorded</i>"}.get(lang, "<i>No grades</i>")
        lbl_avg_cls = {"tr": "Sınıf Dersi Ortalaması", "ru": "Средний балл по предмету", "uz": "Sinf fan o'rtachasi", "en": "Class Subject Average"}.get(lang, "Average")
        lbl_cnt = {"tr": "not", "ru": "оценок", "uz": "baho", "en": "grades"}.get(lang, "grades")

        lines = [gr_sheet_title]
        scores_all = []
        for s in students:
            grades = (await session.execute(select(Grade).where(Grade.student_id == s.id, Grade.subject == subject_name).order_by(Grade.created_at))).scalars().all()
            if grades:
                g_str = ", ".join([f"{g.exam_type or '1. Yazılı'}: <b>{g.score}</b>" for g in grades])
                for g in grades: scores_all.append(g.score)
            else:
                g_str = no_gr_str
            lines.append(f"• <b>{escape_md(s.full_name)}</b> (№{s.student_number}): {g_str}")

        avg_class = round(sum(scores_all)/len(scores_all), 1) if scores_all else 0
        lines.append(f"\n📈 <b>{lbl_avg_cls}:</b> <b>{avg_class}</b> ({len(scores_all)} {lbl_cnt})")

        buttons = [get_nav_buttons(lang, back_callback=f"gr_cls:{class_name}")]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("gr_cls:"))
async def cb_grade_students(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.full_name))).scalars().all()

        buttons = []
        for s in students:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"gr_st:{s.id}")])
        buttons.append([InlineKeyboardButton(text=f"📊 {class_name} {get_text('btn_class_grade_sheet', lang)}", callback_data=f"tch:grade_sheet:{class_name}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:grade_classes")])
        await safe_edit_or_answer(query, get_text("grade_select_student", lang, class_name=escape_md(class_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("gr_st:"))
async def cb_grade_select_exam_type(query: CallbackQuery, state: FSMContext):
    st_id = int(query.data.split(":")[1])
    GRADE_CACHE[query.from_user.id] = {"student_id": st_id}

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)

    prompt = get_text("prompt_select_exam_type", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name))
    kb = get_exam_type_kb(st.class_name, st.id, lang=lang)
    await safe_edit_or_answer(query, prompt, reply_markup=kb, parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("gr_type:"))
async def cb_grade_enter_score(query: CallbackQuery, state: FSMContext):
    type_code = query.data.split(":")[1]
    type_labels = {"1_yazili": "1. Yazılı", "2_yazili": "2. Yazılı", "sozlu": "Sözlü / Performans"}
    exam_type = type_labels.get(type_code, "1. Yazılı")
    GRADE_CACHE[query.from_user.id]["exam_type"] = exam_type

    st_id = GRADE_CACHE[query.from_user.id].get("student_id")
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)

    text = get_text("prompt_grade_score", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name), exam_type=escape_md(exam_type))
    await safe_edit_or_answer(query, text, parse_mode="HTML")
    await state.set_state(Form.grade_score)
    await query.answer()

@router.message(Form.grade_score)
async def process_grade_score(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    score_str = message.text.strip().replace(",", ".")
    try:
        score_val = float(score_str)
    except ValueError:
        await message.answer(get_text("invalid_score_format", lang))
        return

    if not (0.0 <= score_val <= 100.0):
        await message.answer(get_text("invalid_score_range", lang))
        return

    GRADE_CACHE[message.from_user.id]["score"] = score_val
    buttons = [
        [InlineKeyboardButton(text=get_text("badge_praise", lang), callback_data="gr_bdg:🟢")],
        [InlineKeyboardButton(text=get_text("badge_missing", lang), callback_data="gr_bdg:🟡")],
        [InlineKeyboardButton(text=get_text("badge_warning", lang), callback_data="gr_bdg:🔴")]
    ]
    await message.answer(get_text("prompt_grade_badge", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await state.set_state(Form.grade_badge)

@router.message(Form.grade_badge)
async def fallback_grade_badge_text(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
    buttons = [
        [InlineKeyboardButton(text=get_text("badge_praise", lang), callback_data="gr_bdg:🟢")],
        [InlineKeyboardButton(text=get_text("badge_missing", lang), callback_data="gr_bdg:🟡")],
        [InlineKeyboardButton(text=get_text("badge_warning", lang), callback_data="gr_bdg:🔴")]
    ]
    await message.answer(get_text("prompt_grade_badge", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))

@router.callback_query(F.data.startswith("gr_bdg:"))
async def cb_grade_save_final(query: CallbackQuery, state: FSMContext):
    if await is_readonly_mode_active():
        await query.answer(get_text("readonly_mode_active_alert", "tr"), show_alert=True)
        return
    badge_val = query.data.split(":")[1]
    data = GRADE_CACHE.pop(query.from_user.id, {})
    st_id = data.get("student_id")
    score_val = data.get("score", 100.0)
    exam_type_val = data.get("exam_type", "1. Yazılı")
    await state.clear()

    saved_grade_id = None
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        subject_name = tch.subject if tch else "Ders"

        st = await session.get(Student, st_id)
        existing_grade = (await session.execute(select(Grade).where(Grade.student_id == st_id, Grade.subject == subject_name, Grade.exam_type == exam_type_val))).scalar_one_or_none()

        if existing_grade:
            existing_grade.score = score_val
            existing_grade.badge = badge_val
            existing_grade.created_at = datetime.utcnow()
            saved_grade_id = existing_grade.id
        else:
            grade = Grade(student_id=st_id, subject=subject_name, exam_type=exam_type_val, score=score_val, badge=badge_val, note="Öğretmen Değerlendirmesi", teacher_id=query.from_user.id)
            session.add(grade)
            await session.flush()
            saved_grade_id = grade.id
        await session.commit()

        parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
        for p_id in parents:
            p_u = await session.get(User, p_id)
            p_lang = p_u.language if p_u else "tr"
            alert_text = get_text("grade_parent_notification", p_lang, name=escape_md(st.full_name), subject=escape_md(subject_name), exam_type=escape_md(exam_type_val), score=score_val, badge=badge_val)
            await safe_send_message(query.message.bot, p_id, alert_text, parse_mode="HTML")

        buttons = []
        if saved_grade_id:
            undo_txt = "⌫ Bu Notu Geri Al / Sil" if lang == "tr" else ("⌫ Отозвать / Удалить оценку" if lang == "ru" else ("⌫ Ushbu bahoni bekor qilish" if lang == "uz" else "⌫ Undo / Delete Score"))
            buttons.append([InlineKeyboardButton(text=undo_txt, callback_data=f"tch:del_gr:{saved_grade_id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("grade_saved_success", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data == "tch:recent_grades")
async def cb_teacher_recent_grades(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        grades = (await session.execute(select(Grade, Student).join(Student, Grade.student_id == Student.id).where(Grade.teacher_id == query.from_user.id).order_by(desc(Grade.created_at)).limit(15))).all()
        if not grades:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_grades", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        buttons = []
        for g, s in grades:
            btn_txt = f"{g.badge} {s.full_name} ({s.class_name}): {g.score} - {g.subject}"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"tch:view_gr:{g.id}")])
        await safe_edit_or_answer(query, get_text("recent_grades_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:view_gr:"))
async def cb_teacher_view_grade(query: CallbackQuery):
    gr_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        grade = await session.get(Grade, gr_id)
        if not grade or grade.teacher_id != query.from_user.id:
            await query.answer(get_text("no_permission_grade", lang), show_alert=True)
            return

        st = await session.get(Student, grade.student_id)
        st_name = escape_md(st.full_name) if st else "Öğrenci"
        st_cls = escape_md(st.class_name) if st else ""
        date_str = grade.created_at.strftime("%d.%m.%Y %H:%M")

        lbl_st = get_text("lbl_role_student", lang)
        lbl_s = get_text("lbl_subject", lang)
        lbl_dt = {"tr": "Tarih:", "ru": "Дата:", "uz": "Sana:", "en": "Date:"}.get(lang, "Date:")
        text = f"📝 <b>{get_text('btn_enter_grade', lang)}</b>\n━━━━━━━━━━━━━━━━━━━━\n• <b>{lbl_st}:</b> {st_name} ({st_cls})\n• <b>{lbl_s}:</b> {escape_md(grade.subject)} ({escape_md(grade.exam_type or '1. Yazılı')})\n• <b>Not:</b> <b>{grade.score}</b> ({grade.badge})\n• <b>{lbl_dt}</b> {date_str}"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_edit_grade", lang), callback_data=f"tch:edit_gr:{grade.id}")],
            [InlineKeyboardButton(text=get_text("btn_del_grade", lang), callback_data=f"tch:del_gr:{grade.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:recent_grades")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:del_gr:"))
async def cb_teacher_delete_grade(query: CallbackQuery):
    gr_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        grade = await session.get(Grade, gr_id)
        if grade and grade.teacher_id == query.from_user.id:
            await session.delete(grade)
            await session.commit()
            await query.answer(get_text("grade_deleted", lang), show_alert=True)
            await cb_teacher_recent_grades(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("tch:edit_gr:"))
async def cb_teacher_edit_grade_init(query: CallbackQuery, state: FSMContext):
    gr_id = int(query.data.split(":")[2])
    await state.update_data(edit_grade_id=gr_id)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="cancel_action")
    await safe_edit_or_answer(query, get_text("prompt_new_score", lang), reply_markup=cancel_kb)
    await state.set_state(Form.edit_grade_val)
    await query.answer()

@router.message(Form.edit_grade_val)
async def process_grade_edit_val(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    score_str = message.text.strip().replace(",", ".")
    try: score_val = float(score_str)
    except ValueError:
        await message.answer(get_text("invalid_score_format", lang))
        return

    if not (0.0 <= score_val <= 100.0):
        await message.answer(get_text("invalid_score_range", lang))
        return

    data = await state.get_data()
    gr_id = data.get("edit_grade_id")
    await state.clear()

    async with AsyncSessionLocal() as session:
        grade = await session.get(Grade, gr_id)
        if grade and grade.teacher_id == message.from_user.id:
            grade.score = score_val
            await session.commit()
            await message.answer(get_text("grade_updated", lang), reply_markup=get_role_reply_kb("teacher", lang))
            await render_clean_dashboard(message, user)

# --- ÖĞRENCİ DAVRANIŞ & ROZET MODÜLÜ ---
@router.callback_query(F.data == "tch:behavior_classes")
async def cb_teacher_behavior_classes(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if tch and tch.assigned_classes and tch.assigned_classes != "ALL":
            allowed = [c.strip() for c in tch.assigned_classes.split(",") if c.strip()]
            classes = [c for c in all_classes if c in allowed] or all_classes
        else:
            classes = all_classes

        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"bh_cls:{c}")])
        prompt_b = {"tr": "⭐ Davranış değerlendirmesi yapmak istediğiniz sınıfı seçiniz:", "ru": "⭐ Выберите класс:", "uz": "⭐ Xulq-atvor baholash uchun sinfni tanlang:", "en": "⭐ Select class for behavior evaluation:"}.get(lang, "⭐ Select class:")
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(query, prompt_b, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("bh_cls:"))
async def cb_behavior_select_class(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.full_name))).scalars().all()

        buttons = []
        for s in students:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"bh_st:{s.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:behavior_classes")])
        prompt_st = {"tr": f"⭐ <b>{class_name} Sınıfı</b>\nÖğrenci seçiniz:", "ru": f"⭐ <b>Класс {class_name}</b>\nВыберите ученика:", "uz": f"⭐ <b>{class_name} sinfi</b>\nO'quvchini tanlang:", "en": f"⭐ <b>Class {class_name}</b>\nSelect student:"}.get(lang, "Select student:")
        await safe_edit_or_answer(query, prompt_st, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("bh_st:"))
async def cb_behavior_select_type(query: CallbackQuery):
    st_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)

    BEHAVIOR_CACHE[query.from_user.id] = {"student_id": st_id}
    buttons = [
        [InlineKeyboardButton(text=get_text("btn_add_positive_badge", lang), callback_data=f"bh_type:pos:{st_id}")],
        [InlineKeyboardButton(text=get_text("btn_add_negative_badge", lang), callback_data=f"bh_type:neg:{st_id}")],
        [InlineKeyboardButton(text=get_text("btn_student_behavior_history", lang), callback_data=f"bh_hist:{st_id}")],
        [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"bh_cls:{st.class_name}")]
    ]
    prompt_t = {"tr": f"⭐ <b>{escape_md(st.full_name)}</b> ({escape_md(st.class_name)})\nLütfen davranış türünü seçiniz:", "ru": f"⭐ <b>{escape_md(st.full_name)}</b>\nВыберите категорию:", "uz": f"⭐ <b>{escape_md(st.full_name)}</b>\nBaholash turini tanlang:", "en": f"⭐ <b>{escape_md(st.full_name)}</b>\nSelect evaluation type:"}.get(lang, "Select evaluation type:")
    await safe_edit_or_answer(query, prompt_t, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("bh_type:"))
async def cb_behavior_type_chosen(query: CallbackQuery, state: FSMContext):
    parts = query.data.split(":")
    b_type = parts[1]
    st_id = int(parts[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    BEHAVIOR_CACHE[query.from_user.id]["type"] = b_type
    if b_type == "pos":
        badges = [
            ("⭐ Derse Aktif Katılım", "⭐"),
            ("🏆 Üstün Başarı", "🏆"),
            ("👏 Nezaket & Yardımlaşma", "👏"),
            ("📖 Kitap Okuma & Gayret", "📖")
        ]
    else:
        badges = [
            ("⚠️ Derse Geç Kalma", "⚠️"),
            ("🔴 Kural İhlali", "🔴"),
            ("📵 Telefon Kuralı", "📵"),
            ("🗣️ Dersi Kaynatma", "🗣️")
        ]

    buttons = []
    for title, icon in badges:
        buttons.append([InlineKeyboardButton(text=f"{icon} {title}", callback_data=f"bh_bdg:{icon}:{title}")])
    buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"bh_st:{st_id}")])

    prompt_bg = {"tr": "Lütfen rozeti seçiniz:", "ru": "Выберите категорию:", "uz": "Nishonni tanlang:", "en": "Select badge:"}.get(lang, "Select badge:")
    await safe_edit_or_answer(query, prompt_bg, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("bh_bdg:"))
async def cb_behavior_badge_chosen(query: CallbackQuery, state: FSMContext):
    parts = query.data.split(":")
    icon = parts[1]
    title = parts[2]
    BEHAVIOR_CACHE[query.from_user.id]["badge"] = icon
    BEHAVIOR_CACHE[query.from_user.id]["title"] = title

    st_id = BEHAVIOR_CACHE[query.from_user.id]["student_id"]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)

    prompt = get_text("prompt_behavior_note", lang, name=escape_md(st.full_name), badge=icon, title=escape_md(title))
    cancel_kb = get_inline_cancel_kb(lang, back_callback="cancel_action")
    await safe_edit_or_answer(query, prompt, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_behavior_note)
    await query.answer()

@router.message(Form.waiting_behavior_note)
async def process_behavior_note(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    note_val = message.text.strip()
    if note_val in ["-", "İptal", "Bekor qilish", "Отмена"]:
        note_val = None

    data = BEHAVIOR_CACHE.pop(message.from_user.id, {})
    st_id = data.get("student_id")
    b_type = data.get("type", "pos")
    icon = data.get("badge", "⭐")
    title = data.get("title", "Davranış Değerlendirmesi")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        teacher_name = tch.full_name if tch else (user.full_name or "Öğretmen")

        b_rec = BehaviorRecord(
            student_id=st_id,
            behavior_type="positive" if b_type == "pos" else "negative",
            badge=icon,
            title=title,
            note=note_val,
            teacher_id=message.from_user.id
        )
        session.add(b_rec)
        await session.commit()

        parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
        for p_id in parents:
            p_u = await session.get(User, p_id)
            p_l = p_u.language if p_u else "tr"
            alert_msg = get_text("behavior_parent_notification", p_l, name=escape_md(st.full_name), class_name=escape_md(st.class_name), badge=icon, title=escape_md(title), note=escape_md(note_val or "-"), teacher=escape_md(teacher_name))
            await safe_send_message(message.bot, p_id, alert_msg, parse_mode="HTML")

        await message.answer(get_text("behavior_saved_success", lang), reply_markup=get_role_reply_kb(user.role, lang), parse_mode="HTML")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data.startswith("bh_hist:"))
async def cb_student_behavior_history(query: CallbackQuery):
    st_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        records = (await session.execute(select(BehaviorRecord).where(BehaviorRecord.student_id == st_id).order_by(desc(BehaviorRecord.created_at)).limit(10))).scalars().all()

        if not records:
            buttons = [get_nav_buttons(lang, back_callback=f"bh_st:{st_id}")]
            await safe_edit_or_answer(query, get_text("no_behavior_records", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        header = {
            "tr": f"⭐ <b>{escape_md(st.full_name)}</b> ({escape_md(st.class_name)}) Davranış Geçmişi:\n━━━━━━━━━━━━━━━━━━━━\n",
            "ru": f"⭐ История поведения <b>{escape_md(st.full_name)}</b>:\n━━━━━━━━━━━━━━━━━━━━\n",
            "uz": f"⭐ <b>{escape_md(st.full_name)}</b> xulq-atvor tarixi:\n━━━━━━━━━━━━━━━━━━━━\n",
            "en": f"⭐ Behavior History for <b>{escape_md(st.full_name)}</b>:\n━━━━━━━━━━━━━━━━━━━━\n"
        }.get(lang, "⭐ Behavior History:\n")

        lines = [header]
        for r in records:
            note_str = f" - <i>{escape_md(r.note)}</i>" if r.note else ""
            lines.append(f"• {r.badge} <b>{escape_md(r.title)}</b> ({r.created_at.strftime('%d.%m')}){note_str}")

        buttons = [get_nav_buttons(lang, back_callback=f"bh_st:{st_id}")]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

# --- ÖDEV TESLİM & KONTROL DÖNGÜSÜ ---
@router.callback_query(F.data == "tch:hw_classes")
async def cb_hw_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if tch and tch.assigned_classes and tch.assigned_classes != "ALL":
            allowed = [c.strip() for c in tch.assigned_classes.split(",") if c.strip()]
            classes = [c for c in all_classes if c in allowed] or all_classes
        else:
            classes = all_classes

        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"📢 {c} - {get_text('btn_send_new_hw', lang)}", callback_data=f"hw_cls:{c}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_my_hws", lang), callback_data="tch:view_my_hws")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(query, get_text("prompt_hw_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data == "tch:view_my_hws")
async def cb_teacher_view_my_hws(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        hws = (await session.execute(select(Homework).where(Homework.teacher_id == query.from_user.id).order_by(desc(Homework.created_at)).limit(10))).scalars().all()
        if not hws:
            buttons = [get_nav_buttons(lang, back_callback="tch:hw_classes")]
            await safe_edit_or_answer(query, get_text("no_hws_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        buttons = []
        for h in hws:
            sub_count = (await session.execute(select(func.count(HomeworkSubmission.id)).where(HomeworkSubmission.homework_id == h.id))).scalar() or 0
            btn_txt = f"📚 {h.class_name} ({h.subject}) [{sub_count} Teslim]"
            buttons.append([
                InlineKeyboardButton(text=btn_txt, callback_data=f"tch:hw_subs:{h.id}"),
                InlineKeyboardButton(text=get_text("btn_delete_action", lang), callback_data=f"tch:del_hw:{h.id}")
            ])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:hw_classes")])
        await safe_edit_or_answer(query, get_text("published_homeworks_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("tch:del_hw:"))
async def cb_teacher_del_hw(query: CallbackQuery):
    hw_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        hw = await session.get(Homework, hw_id)
        if hw and hw.teacher_id == query.from_user.id:
            await session.execute(delete(HomeworkSubmission).where(HomeworkSubmission.homework_id == hw.id))
            await session.delete(hw)
            await session.commit()
            await query.answer(get_text("homework_deleted_toast", lang), show_alert=True)
            await cb_teacher_view_my_hws(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("hw_cls:"))
async def cb_hw_enter_content(query: CallbackQuery, state: FSMContext):
    class_name = query.data.split(":")[1]
    HW_CACHE[query.from_user.id] = class_name
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("prompt_hw_content", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
    await state.set_state(Form.hw_content)
    await query.answer()

@router.message(Form.hw_content)
async def process_hw_content(message: Message, state: FSMContext):
    class_name = HW_CACHE.pop(message.from_user.id, "9-A")
    await state.clear()
    photo_id = message.photo[-1].file_id if message.photo else None
    voice_id = message.voice.file_id if message.voice else (message.audio.file_id if message.audio else None)
    file_id = voice_id or photo_id

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        def_hw_txt = {"tr": "Ödev Panosu", "ru": "Доска заданий", "uz": "Vazifalar paneli", "en": "Homework Board"}.get(lang, "Homework Board")
        hw_text = message.caption or message.text or def_hw_txt
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        subj = tch.subject if tch else "Ders"

        hw = Homework(class_name=class_name, subject=subj, content=hw_text, file_id=file_id, teacher_id=message.from_user.id)
        session.add(hw)
        await session.commit()

        st_ids = (await session.execute(select(Student.id).where(Student.class_name == class_name))).scalars().all()
        parent_tg_ids = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id.in_(st_ids)))).scalars().all()
        student_tg_ids = (await session.execute(select(Student.student_telegram_id).where(Student.id.in_(st_ids), Student.student_telegram_id.isnot(None)))).scalars().all()

        target_ids = set(parent_tg_ids + student_tg_ids)
        for t_id in target_ids:
            u_rec = await session.get(User, t_id)
            u_l = u_rec.language if u_rec else "tr"
            dispatch_texts = {
                "tr": f"📢 *{escape_md(class_name)} ÖDEV PANOSU ({escape_md(subj)})*\n\n{escape_md(hw_text)}",
                "ru": f"📢 *ДОСКА ЗАДАНИЙ КЛАССА {escape_md(class_name)} ({escape_md(subj)})*\n\n{escape_md(hw_text)}",
                "uz": f"📢 *{escape_md(class_name)} SINFI VAZIFALAR PANELI ({escape_md(subj)})*\n\n{escape_md(hw_text)}",
                "en": f"📢 *CLASS {escape_md(class_name)} HOMEWORK BOARD ({escape_md(subj)})*\n\n{escape_md(hw_text)}"
            }
            caption = dispatch_texts.get(u_l, dispatch_texts["en"])
            if voice_id:
                try: await message.bot.send_voice(chat_id=t_id, voice=voice_id, caption=caption, parse_mode="HTML")
                except Exception: pass
            elif photo_id:
                try: await message.bot.send_photo(chat_id=t_id, photo=photo_id, caption=caption, parse_mode="HTML")
                except Exception: pass
            else:
                await safe_send_message(message.bot, t_id, caption, parse_mode="HTML")

        await message.answer(get_text("hw_sent_success", lang, class_name=escape_md(class_name)), parse_mode="HTML")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data == "act_view_hw")
async def cb_view_homeworks_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        cls_name = None
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: cls_name = st.class_name
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == query.from_user.id))).scalar_one_or_none()
            if st: cls_name = st.class_name

        if not cls_name:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_linked_student", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        hws = (await session.execute(select(Homework).where(Homework.class_name == cls_name).order_by(desc(Homework.created_at)).limit(6))).scalars().all()
        if not hws:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_active_homeworks", lang, class_name=escape_md(cls_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        lines = [f"{get_text('homework_board_title', lang, class_name=escape_md(cls_name))}\n"]
        buttons = []
        for h in hws:
            lines.append(f"📌 *{escape_md(h.subject)}* ({h.created_at.strftime('%d.%m.%Y')}):\n{escape_md(h.content)}")
            row = []
            if h.file_id:
                row.append(InlineKeyboardButton(text=f"📷 {get_text('btn_view_photo', lang)}: {h.subject}", callback_data=f"hw_photo:{h.id}"))
            if user and user.role == "student":
                row.append(InlineKeyboardButton(text=get_text("btn_submit_hw", lang), callback_data=f"hw_submit:{h.id}"))
            if row: buttons.append(row)
            lines.append("")

        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("hw_photo:"))
async def cb_view_hw_photo(query: CallbackQuery):
    hw_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        hw = await session.get(Homework, hw_id)
        if hw and hw.file_id:
            hw_caption_texts = {
                "tr": f"📚 *{escape_md(hw.class_name)} Ödev Görseli ({escape_md(hw.subject)})*\n\n{escape_md(hw.content)}",
                "ru": f"📚 *Класс {escape_md(hw.class_name)} Фото задания ({escape_md(hw.subject)})*\n\n{escape_md(hw.content)}",
                "uz": f"📚 *{escape_md(hw.class_name)} sinfi vazifa rasmi ({escape_md(hw.subject)})*\n\n{escape_md(hw.content)}",
                "en": f"📚 *Class {escape_md(hw.class_name)} Homework Image ({escape_md(hw.subject)})*\n\n{escape_md(hw.content)}"
            }
            caption = hw_caption_texts.get(lang, hw_caption_texts["en"])
            try:
                await query.message.bot.send_photo(chat_id=query.from_user.id, photo=hw.file_id, caption=caption, parse_mode="HTML")
                await query.answer()
                return
            except Exception: pass
    await query.answer(get_text("image_load_error", lang), show_alert=True)

@router.callback_query(F.data.startswith("hw_submit:"))
async def cb_submit_homework_init(query: CallbackQuery, state: FSMContext):
    hw_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        hw = await session.get(Homework, hw_id)
        if not hw: return

    SUBMISSION_CACHE[query.from_user.id] = {"homework_id": hw_id}
    cancel_kb = get_inline_cancel_kb(lang, back_callback="cancel_action")
    await safe_edit_or_answer(query, get_text("prompt_hw_submission", lang, subject=escape_md(hw.subject)), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_hw_submission)
    await query.answer()

@router.message(Form.waiting_hw_submission)
async def process_homework_submission(message: Message, state: FSMContext):
    data = SUBMISSION_CACHE.pop(message.from_user.id, {})
    hw_id = data.get("homework_id")
    await state.clear()

    photo_id = message.photo[-1].file_id if message.photo else None

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        def_sub_txt = {"tr": "Ödev Teslimi", "ru": "Сдача задания", "uz": "Vazifa topshirish", "en": "Homework Submission"}.get(lang, "Homework Submission")
        content_text = message.caption or message.text or def_sub_txt
        st = (await session.execute(select(Student).where(Student.student_telegram_id == message.from_user.id))).scalar_one_or_none()
        hw = await session.get(Homework, hw_id) if hw_id else None

        if not st or not hw:
            await message.answer(get_text("no_linked_student", lang), reply_markup=get_role_reply_kb("student", lang))
            return

        sub = HomeworkSubmission(
            homework_id=hw.id,
            student_id=st.id,
            student_telegram_id=message.from_user.id,
            content=content_text,
            file_id=photo_id,
            status="submitted"
        )
        session.add(sub)
        await session.commit()

        hw_receipt_txt = (
            f"📘 <b>Ödev:</b> {escape_html(hw.title)}\n"
            f"👤 <b>Öğrenci:</b> {escape_html(st.full_name)} (No: {st.student_number})\n"
            f"📅 <b>Teslim Zamanı:</b> <code>{get_local_date().strftime('%d.%m.%Y %H:%M')}</code>\n"
            f"🆔 <b>Kayıt No:</b> <code>#ODV-{sub.id:04d}</code>\n"
            f"✅ <i>Ödeviniz başarıyla öğretmen incelemesine gönderildi.</i>"
        )
        await send_digital_record(
            bot=message.bot,
            chat_id=message.from_user.id,
            category="homework",
            title=f"Ödev Teslim Makbuzu: {hw.title}",
            content=hw_receipt_txt,
            file_id=photo_id,
            file_type="photo" if photo_id else None,
            lang=lang
        )
        await message.answer(get_text("hw_submission_received", lang), reply_markup=get_role_reply_kb("student", lang), parse_mode="HTML")
        await render_clean_dashboard(message, user)

        if hw.teacher_id:
            tch_u = await session.get(User, hw.teacher_id)
            tch_l = tch_u.language if tch_u else "tr"
            t_msg = {
                "tr": f"📥 <b>YENİ ÖDEV TESLİMİ</b>\n\n🧑‍🎓 <b>Öğrenci:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n📚 <b>Ders:</b> {escape_html(hw.subject)}\n📝 <b>Not:</b> {escape_html(content_text or '-')}",
                "ru": f"📥 <b>НОВАЯ СДАЧА ДОМАШНЕГО ЗАДАНИЯ</b>\n\n🧑‍🎓 <b>Ученик:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n📚 <b>Предмет:</b> {escape_html(hw.subject)}\n📝 <b>Комментарий:</b> {escape_html(content_text or '-')}",
                "uz": f"📥 <b>YANGI VAZIFA TOPSHIRILDI</b>\n\n🧑‍🎓 <b>O'quvchi:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n📚 <b>Fan:</b> {escape_html(hw.subject)}\n📝 <b>Izoh:</b> {escape_html(content_text or '-')}",
                "en": f"📥 <b>NEW HOMEWORK SUBMISSION</b>\n\n🧑‍🎓 <b>Student:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n📚 <b>Subject:</b> {escape_html(hw.subject)}\n📝 <b>Note:</b> {escape_html(content_text or '-')}"
            }.get(tch_l, "New Homework Submission")
            btn_view = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_view_submissions", tch_l), callback_data=f"tch:view_sub:{sub.id}")]
            ])
            if photo_id:
                try: await message.bot.send_photo(chat_id=hw.teacher_id, photo=photo_id, caption=t_msg, reply_markup=btn_view, parse_mode="HTML")
                except Exception: pass
            else:
                await safe_send_message(message.bot, hw.teacher_id, t_msg, reply_markup=btn_view, parse_mode="HTML")

@router.callback_query(F.data.startswith("tch:hw_subs:"))
async def cb_teacher_hw_submissions_list(query: CallbackQuery):
    hw_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        hw = await session.get(Homework, hw_id)
        subs = (await session.execute(select(HomeworkSubmission, Student).join(Student, HomeworkSubmission.student_id == Student.id).where(HomeworkSubmission.homework_id == hw_id))).all()

        no_subs_txt = {
            "tr": "Henüz teslim edilen ödev yok.",
            "ru": "Сданных заданий пока нет.",
            "uz": "Hali topshirilgan vazifalar yo'q.",
            "en": "No submitted assignments yet."
        }.get(lang, "No submissions yet.")

        if not subs:
            buttons = [get_nav_buttons(lang, back_callback="tch:view_my_hws")]
            await safe_edit_or_answer(query, no_subs_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        buttons = []
        for sub, st in subs:
            status_icon = "✅" if sub.status == "approved" else ("🔄" if sub.status == "revision" else "⚪")
            buttons.append([InlineKeyboardButton(text=f"{status_icon} {st.full_name} ({st.class_name})", callback_data=f"tch:view_sub:{sub.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:view_my_hws")])

        hdr_subs = {
            "tr": f"📥 *{escape_md(hw.subject)} Ödev Teslimleri ({len(subs)}):*",
            "ru": f"📥 *Сданные задания по предмету {escape_md(hw.subject)} ({len(subs)}):*",
            "uz": f"📥 *{escape_md(hw.subject)} fanidan topshirilgan vazifalar ({len(subs)}):*",
            "en": f"📥 *{escape_md(hw.subject)} Homework Submissions ({len(subs)}):*"
        }.get(lang, f"📥 *{escape_md(hw.subject)} Submissions ({len(subs)}):*")

        await safe_edit_or_answer(query, hdr_subs, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:view_sub:"))
async def cb_teacher_view_submission(query: CallbackQuery):
    sub_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        sub = await session.get(HomeworkSubmission, sub_id)
        if not sub: return
        st = await session.get(Student, sub.student_id)
        hw = await session.get(Homework, sub.homework_id)

        t_sub_hdr = {"tr": "📥 *Öğrenci Ödev Teslimi*", "ru": "📥 *Сдача домашнего задания*", "uz": "📥 *O'quvchi vazifa topshirishi*", "en": "📥 *Student Homework Submission*"}.get(lang, "📥 *Homework Submission*")
        lbl_sub_st = {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"}.get(lang, "Student")
        lbl_sub_sb = {"tr": "Ders", "ru": "Предмет", "uz": "Fan", "en": "Subject"}.get(lang, "Subject")
        lbl_sub_stt = {"tr": "Durum", "ru": "Статус", "uz": "Holat", "en": "Status"}.get(lang, "Status")
        lbl_sub_nt = {"tr": "Açıklama", "ru": "Описание", "uz": "Izoh", "en": "Description"}.get(lang, "Description")

        text = f"{t_sub_hdr}\n\n🧑🎓 {lbl_sub_st}: *{escape_md(st.full_name)}* ({escape_md(st.class_name)})\n📚 {lbl_sub_sb}: *{escape_md(hw.subject)}*\n📌 {lbl_sub_stt}: *{sub.status.upper()}*\n📝 {lbl_sub_nt}: {escape_md(sub.content or '-')}"
        buttons = [
            [
                InlineKeyboardButton(text=get_text("btn_hw_approve", lang), callback_data=f"tch:sub_appr:{sub.id}"),
                InlineKeyboardButton(text=get_text("btn_hw_revision", lang), callback_data=f"tch:sub_rev:{sub.id}")
            ],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"tch:hw_subs:{hw.id}")]
        ]

        if sub.file_id:
            try: await query.message.delete()
            except Exception: pass
            await query.message.bot.send_photo(chat_id=query.from_user.id, photo=sub.file_id, caption=text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        else:
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:sub_appr:"))
async def cb_teacher_approve_sub(query: CallbackQuery):
    sub_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        sub = await session.get(HomeworkSubmission, sub_id)
        if sub:
            sub.status = "approved"
            sub.teacher_feedback = "Ödev başarıyla kabul edildi."
            await session.commit()

            st = await session.get(Student, sub.student_id)
            hw = await session.get(Homework, sub.homework_id)
            s_u = await session.get(User, sub.student_telegram_id)
            s_l = s_u.language if s_u else "tr"
            await safe_send_message(query.message.bot, sub.student_telegram_id, get_text("hw_feedback_sent_user", s_l, subject=hw.subject, status={"tr": "✅ Kabul Edildi / Onaylandı", "ru": "✅ Принято / Одобрено", "uz": "✅ Qabul qilindi / Tasdiqlandi", "en": "✅ Accepted / Approved"}.get(s_l, "✅ Accepted"), feedback=sub.teacher_feedback), parse_mode="HTML")

            appr_msg_tst = {"tr": "Ödev onaylandı.", "ru": "Работа принята.", "uz": "Vazifa qabul qilindi.", "en": "Homework approved."}.get(lang, "Approved.")
            await query.answer(appr_msg_tst, show_alert=True)
            query.data = f"tch:hw_subs:{hw.id}"
            await cb_teacher_hw_submissions_list(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("tch:sub_rev:"))
async def cb_teacher_sub_rev_init(query: CallbackQuery, state: FSMContext):
    sub_id = int(query.data.split(":")[2])
    SUBMISSION_CACHE[query.from_user.id] = {"sub_id": sub_id}
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="cancel_action")
    p_rev = {
        "tr": "📝 Lütfen öğrenciye iletilecek düzeltme notunu yazınız:",
        "ru": "📝 Введите комментарий для доработки задания:",
        "uz": "📝 O'quvchiga yuboriladigan tuzatish izohini yozing:",
        "en": "📝 Please enter revision notes for the student:"
    }.get(lang, "Enter revision notes:")
    await safe_edit_or_answer(query, p_rev, reply_markup=cancel_kb)
    await state.set_state(Form.waiting_hw_feedback)
    await query.answer()

@router.message(Form.waiting_hw_feedback)
async def process_hw_feedback(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    fb_text = message.text.strip()
    data = SUBMISSION_CACHE.pop(message.from_user.id, {})
    sub_id = data.get("sub_id")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        sub = await session.get(HomeworkSubmission, sub_id)
        if sub:
            sub.status = "revision"
            sub.teacher_feedback = fb_text
            await session.commit()

            hw = await session.get(Homework, sub.homework_id)
            s_u = await session.get(User, sub.student_telegram_id)
            s_l = s_u.language if s_u else "tr"
            await safe_send_message(message.bot, sub.student_telegram_id, get_text("hw_feedback_sent_user", s_l, subject=hw.subject, status={"tr": "⚠️ Düzeltme İsteniyor", "ru": "⚠️ Требуется доработка", "uz": "⚠️ Qayta ishlash so'ralmoqda", "en": "⚠️ Revision Requested"}.get(s_l, "⚠️ Revision Requested"), feedback=fb_text), parse_mode="HTML")

        done_rev = {
            "tr": "✅ Düzeltme talebi öğrenciye iletildi.",
            "ru": "✅ Запрос на доработку отправлен ученику.",
            "uz": "✅ Qayta ishlash so'rovi o'quvchiga yuborildi.",
            "en": "✅ Revision request sent to student."
        }.get(lang, "Revision request sent.")
        await message.answer(done_rev, reply_markup=get_role_reply_kb("teacher", lang))
        await render_clean_dashboard(message, user)

# --- SINAV TAKVİMİ & AKADEMİK TAKVİM ---
@router.callback_query(F.data == "act_view_exams")
async def cb_view_exams_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        cls_name = None
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: cls_name = st.class_name
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == query.from_user.id))).scalar_one_or_none()
            if st: cls_name = st.class_name

        if not cls_name:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_linked_student", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        exams = (await session.execute(select(ExamSchedule).where(ExamSchedule.class_name == cls_name).order_by(ExamSchedule.exam_date))).scalars().all()
        if not exams:
            buttons = []
            if user and user.role in ("admin", "teacher"):
                buttons.append([InlineKeyboardButton(text=get_text("btn_add_exam", lang), callback_data=f"adm:add_exam:{cls_name}")])
            buttons.append(get_nav_buttons(lang))
            await safe_edit_or_answer(query, get_text("no_exams_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        lines = [get_text("exam_schedule_title", lang, class_name=escape_md(cls_name)) + "\n"]
        today_cur = get_local_date()
        nearest_exam = None
        min_days = 999
        for e in exams:
            days_left = (e.exam_date - today_cur).days
            if 0 <= days_left < min_days:
                min_days = days_left
                nearest_exam = (e, days_left)

            if days_left == 0:
                badge = {"tr": "🚨 BUGÜN!", "ru": "🚨 СЕГОДНЯ!", "uz": "🚨 BUGUN!", "en": "🚨 TODAY!"}.get(lang, "🚨 BUGÜN!")
            elif days_left == 1:
                badge = {"tr": "⚡ YARIN!", "ru": "⚡ ЗАВТРА!", "uz": "⚡ ERTAGA!", "en": "⚡ TOMORROW!"}.get(lang, "⚡ YARIN!")
            elif days_left > 1:
                badge = {"tr": f"⏳ {days_left} gün kaldı", "ru": f"⏳ осталось {days_left} дн.", "uz": f"⏳ {days_left} kun qoldi", "en": f"⏳ {days_left} days left"}.get(lang, f"⏳ {days_left} gün")
            else:
                badge = {"tr": "✅ Tamamlandı", "ru": "✅ Завершен", "uz": "✅ Yakunlandi", "en": "✅ Completed"}.get(lang, "✅ Tamamlandı")

            desc_str = f" - <i>{escape_html(e.description)}</i>" if e.description else ""
            lines.append(f"• 📅 <b>{e.exam_date.strftime('%d.%m.%Y')}</b> ({e.exam_time}) — <b>{escape_html(e.subject)}</b> [{badge}]{desc_str}")

        if nearest_exam:
            ne_obj, ne_days = nearest_exam
            top_banner = "📌 <b>" + {"tr": f"EN YAKIN SINAV: {ne_obj.subject} (Kalan: {ne_days} Gün)", "ru": f"БЛИЖАЙШИЙ ЭКЗАМЕН: {ne_obj.subject} ({ne_days} дн.)", "uz": f"ENG YAQIN IMTIHON: {ne_obj.subject} ({ne_days} kun)", "en": f"NEXT EXAM: {ne_obj.subject} ({ne_days} days left)"}.get(lang, f"NEXT: {ne_obj.subject}") + "</b>\n──────────────\n"
            lines.insert(1, top_banner)

        buttons = []
        if user and user.role in ("admin", "teacher"):
            buttons.append([InlineKeyboardButton(text=get_text("btn_add_exam", lang), callback_data=f"adm:add_exam:{cls_name}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:add_exam:"))
async def cb_admin_add_exam_init(query: CallbackQuery, state: FSMContext):
    class_name = query.data.split(":")[2]
    EXAM_CACHE[query.from_user.id] = {"class_name": class_name}
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback=f"adm:show_class:{class_name}")
    p_ex_sb = {
        "tr": f"📅 *{class_name} Sınav Girişi*\n\nLütfen Ders Adını yazınız (Örn: `Matematik`):",
        "ru": f"📅 *Ввод экзамена для класса {class_name}*\n\nВведите название предмета (Напр: `Математика`):",
        "uz": f"📅 *{class_name} sinfi imtihon kiritish*\n\nFan nomini kiriting (Masalan: `Matematika`):",
        "en": f"📅 *Add Exam for Class {class_name}*\n\nEnter Subject Name (e.g. `Mathematics`):"
    }.get(lang, f"Add Exam for {class_name}:")
    await safe_edit_or_answer(query, p_ex_sb, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_exam_subject)
    await query.answer()

@router.message(Form.waiting_exam_subject)
async def process_exam_subject(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    subj = message.text.strip()
    EXAM_CACHE[message.from_user.id]["subject"] = subj
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
    p_ex_dt = {
        "tr": "📅 Lütfen sınav tarihini ve saatini yazınız (Örn: `2026-06-15 09:30`):",
        "ru": "📅 Введите дату и время экзамена (Напр: `2026-06-15 09:30`):",
        "uz": "📅 Imtihon sanasi va vaqtini kiriting (Masalan: `2026-06-15 09:30`):",
        "en": "📅 Please enter exam date and time (e.g. `2026-06-15 09:30`):"
    }.get(lang, "Enter date and time:")
    await message.answer(p_ex_dt, parse_mode="HTML")
    await state.set_state(Form.waiting_exam_date)

@router.message(Form.waiting_exam_date)
async def process_exam_date(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    dt_str = message.text.strip()
    data = EXAM_CACHE.pop(message.from_user.id, {})
    class_name = data.get("class_name", "9-A")
    subj = data.get("subject", "Ders")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        exam_d = get_local_date()
        exam_t = "09:00"

        parts = dt_str.split(" ")
        if parts:
            try: exam_d = datetime.strptime(parts[0], "%Y-%m-%d").date()
            except Exception: pass
            if len(parts) > 1: exam_t = parts[1]

        ex = ExamSchedule(class_name=class_name, subject=subj, exam_date=exam_d, exam_time=exam_t, description="Dönem Sınavı")
        session.add(ex)
        await session.commit()

        done_ex = {
            "tr": f"✅ *{class_name}* sınıfı için *{subj}* sınav tarihi kaydedildi: `{exam_d} {exam_t}`",
            "ru": f"✅ Для класса *{class_name}* сохранен экзамен по предмету *{subj}*: `{exam_d} {exam_t}`",
            "uz": f"✅ *{class_name}* sinfi uchun *{subj}* fanidan imtihon sanasi saqlandi: `{exam_d} {exam_t}`",
            "en": f"✅ Exam for class *{class_name}* ({subj}) saved: `{exam_d} {exam_t}`"
        }.get(lang, "Exam saved.")
        await message.answer(done_ex, reply_markup=get_role_reply_kb(user.role, lang), parse_mode="HTML")
        await render_clean_dashboard(message, user)

# --- 🚨 ACİL DURUM / KIRMIZI ALARM ---
@router.callback_query(F.data == "adm:emergency_init")
async def cb_admin_emergency_pin_guard(query: CallbackQuery, state: FSMContext):
    await cb_admin_emergency_init(query, state)

async def cb_admin_emergency_init(query: CallbackQuery, state: FSMContext | None = None):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:cat_tools_reports")
    await safe_edit_or_answer(query, get_text("emergency_alert_prompt", lang), reply_markup=cancel_kb, parse_mode="HTML")
    if state: await state.set_state(Form.waiting_emergency_text)

@router.message(Form.waiting_emergency_text)
async def process_emergency_text(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    em_text = message.text.strip()
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id):
            await message.answer(get_text("unauthorized_action", lang))
            return

        alert = EmergencyAlert(message_text=em_text, created_by=message.from_user.id)
        session.add(alert)
        await session.commit()

        parents = (await session.execute(select(User).where(User.role == "parent"))).scalars().all()
        sent_cnt = 0
        for p in parents:
            p_l = p.language
            em_msg = f"🚨 *ACİL DURUM / KIRMIZI ALARM* 🚨\n\n{escape_md(em_text)}\n\n⚠️ Lütfen aşağıdaki butona basarak durumu onaylayınız:"
            ack_btn = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_emergency_ack", p_l), callback_data=f"ack_em:{alert.id}")]
            ])
            res = await safe_send_message(message.bot, p.telegram_id, em_msg, reply_markup=ack_btn, parse_mode="HTML")
            if res: sent_cnt += 1
            await asyncio.sleep(0.04)

        em_rep = {
            "tr": f"🚨 Acil durum alarmı *{sent_cnt}* veliye sesli bildirimle iletildi!",
            "ru": f"🚨 Экстренное оповещение отправлено *{sent_cnt}* родителям со звуковым сигналом!",
            "uz": f"🚨 Favqulodda xabar *{sent_cnt}* ta ota-onaga ovozli signal bilan yetkazildi!",
            "en": f"🚨 Emergency alert dispatched to *{sent_cnt}* parents with high-priority audio!"
        }.get(lang, "Alert sent.")
        await message.answer(em_rep, reply_markup=get_role_reply_kb("admin", lang), parse_mode="HTML")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data.startswith("ack_em:"))
async def cb_ack_emergency(query: CallbackQuery):
    alert_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        existing = (await session.execute(select(EmergencyAck).where(EmergencyAck.alert_id == alert_id, EmergencyAck.user_telegram_id == query.from_user.id))).scalar_one_or_none()
        if not existing:
            session.add(EmergencyAck(alert_id=alert_id, user_telegram_id=query.from_user.id))
            await session.commit()

    ack_tst = {
        "tr": "✅ Onayınız idareye iletildi.",
        "ru": "✅ Ваше подтверждение отправлено администрации.",
        "uz": "✅ Tasdig'ingiz ma'muriyatga yetkazildi.",
        "en": "✅ Your confirmation has been sent to administration."
    }.get(lang, "Confirmed.")
    await query.answer(ack_tst, show_alert=True)
    await query.message.edit_reply_markup(reply_markup=None)

@router.callback_query(F.data == "adm:emergency_monitor")
async def cb_admin_emergency_monitor(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        latest_alert = (await session.execute(select(EmergencyAlert).order_by(desc(EmergencyAlert.created_at)).limit(1))).scalar_one_or_none()
        no_em_txt = {
            "tr": "Kayıtlı acil durum bulunamadı.",
            "ru": "Активных экстренных ситуаций не найдено.",
            "uz": "Faol favqulodda holat topilmadi.",
            "en": "No active emergency alerts found."
        }.get(lang, "No emergency alerts.")

        if not latest_alert:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, no_em_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        acked_ids = (await session.execute(select(EmergencyAck.user_telegram_id).where(EmergencyAck.alert_id == latest_alert.id))).scalars().all()
        all_parents = (await session.execute(select(User).where(User.role == "parent"))).scalars().all()
        unacked = [p for p in all_parents if p.telegram_id not in acked_ids]

        all_acked_txt = {
            "tr": "✅ Tüm veliler acil durumu onayladı.",
            "ru": "✅ Все родители подтвердили экстренное сообщение.",
            "uz": "✅ Barcha ota-onalar favqulodda xabarni tasdiqlashdi.",
            "en": "✅ All parents have confirmed the emergency alert."
        }.get(lang, "All parents confirmed.")

        if not unacked:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, all_acked_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        lines = [get_text("emergency_monitor_title", lang) + f" ({len(unacked)}):\n"]
        for p in unacked[:25]:
            p_phone = f" (`{p.phone}`)" if p.phone else ""
            lines.append(f"• 👤 *{escape_md(p.full_name or 'Veli')}*{p_phone} - ID: `{p.telegram_id}`")

        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

# --- VELİ VE ÖĞRENCİ DİĞER AKSİYONLARI ---
@router.callback_query(F.data == "parent:switch_student")
async def cb_parent_switch_student(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        stmt = select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == user.telegram_id)
        children = (await session.execute(stmt)).scalars().all()

        if not children:
            await query.answer(get_text("no_registered_students", lang), show_alert=True)
            return

        buttons = []
        for c in children:
            is_active = "⭐ " if c.id == user.current_child_id else ""
            buttons.append([InlineKeyboardButton(text=f"{is_active}🧑🎓 {c.full_name} ({c.class_name})", callback_data=f"set_child:{c.id}")])
        await safe_edit_or_answer(query, get_text("parent_choose_child", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data == "parent:add_child_code")
async def cb_parent_add_child_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="cancel_action")
    await safe_edit_or_answer(query, get_text("prompt_add_child_code", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.parent_add_child_code)
    await query.answer()

@router.message(Form.parent_add_child_code)
async def process_parent_add_child_code(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    code_raw = message.text.strip()
    clean_code = normalize_code(code_raw)
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        p_stmt = select(Student).where(func.upper(func.trim(Student.parent_code)) == clean_code)
        p_student = (await session.execute(p_stmt)).scalar_one_or_none()

        reply_kb = get_role_reply_kb("parent", lang)
        if not p_student:
            await message.answer(get_text("invalid_parent_code", lang), reply_markup=reply_kb, parse_mode="HTML")
            await render_clean_dashboard(message, user)
            return

        rel = (await session.execute(select(ParentStudent).where(
            ParentStudent.parent_telegram_id == message.from_user.id,
            ParentStudent.student_id == p_student.id
        ))).scalar_one_or_none()

        if not rel:
            session.add(ParentStudent(parent_telegram_id=message.from_user.id, student_id=p_student.id))
        user.current_child_id = p_student.id
        p_student.is_parent_code_burned = True
        await session.commit()

        success_text = get_text("child_added_success", lang, name=escape_md(p_student.full_name), class_name=escape_md(p_student.class_name))
        await message.answer(success_text, reply_markup=reply_kb, parse_mode="HTML")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data.startswith("set_child:"))
async def cb_set_child(query: CallbackQuery):
    child_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        rel = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == query.from_user.id, ParentStudent.student_id == child_id))).scalar_one_or_none()
        if not rel:
            await query.answer(get_text("no_permission_student_record", lang), show_alert=True)
            return

        user.current_child_id = child_id
        await session.commit()
        st = await session.get(Student, child_id)
        await query.answer(get_text("student_switched_success", lang, name=st.full_name, class_name=st.class_name), show_alert=True)
    await render_clean_dashboard(query, user)

# --- E-OKUL FORMATINDA KARNE & GÖRSEL DEVAMSIZLIK ÇUBUĞU ---
@router.callback_query(F.data == "act_view_report")
async def cb_view_report_card(query: CallbackQuery, state: FSMContext):
    await state.clear()
    user_id = query.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        st = None
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user_id))).scalar_one_or_none()

        if not st:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_linked_student", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        attendances = (await session.execute(select(Attendance).where(Attendance.student_id == st.id))).scalars().all()
        absent_count = sum(1 for a in attendances if a.status == "absent")
        excused_count = sum(1 for a in attendances if a.status == "excused")

        legal_limit = 10
        rem_att = max(0, legal_limit - absent_count)
        att_bar = render_progress_bar(absent_count, legal_limit)

        grades = (await session.execute(select(Grade).where(Grade.student_id == st.id).order_by(Grade.created_at.desc()))).scalars().all()

        subj_map = {}
        for g in grades:
            subj_map.setdefault(g.subject, []).append(g.score)

        subj_lines = []
        all_avgs = []
        for s_name, scores in subj_map.items():
            s_avg = round(sum(scores) / len(scores), 1)
            all_avgs.append(s_avg)
            scores_str = " | ".join([str(int(sc) if sc.is_integer() else sc) for sc in scores])
            badge = "🟢" if s_avg >= 70 else ("🟡" if s_avg >= 50 else "🔴")
            subj_lines.append(f"┌ 📚 <b>{escape_md(s_name)}:</b> [ {scores_str} ]\n└ ↳ <b>Ort:</b> <code>{s_avg}</code> {badge}")

        general_gpa = round(sum(all_avgs) / len(all_avgs), 1) if all_avgs else 100.0
        gpa_bar = render_progress_bar(int(round(general_gpa / 10)), 10)
        honor_badge = "🏆 " + {"tr": "Üstün Başarı", "ru": "Отличник", "uz": "A'lochi", "en": "Honor Roll"}.get(lang, "Honor") if general_gpa >= 85 else ("🌟 " + {"tr": "Teşekkür", "ru": "Хорошист", "uz": "Yaxshi", "en": "Good"}.get(lang, "Good") if general_gpa >= 70 else ("📈 " + {"tr": "Gelişmekte", "ru": "В процессе", "uz": "O'rtacha", "en": "Progress"}.get(lang, "Progress") if general_gpa >= 50 else "⚠️ " + {"tr": "Destek Gerekli", "ru": "Требуется помощь", "uz": "Yordam kerak", "en": "Needs Support"}.get(lang, "Support")))

        behaviors = (await session.execute(select(BehaviorRecord).where(BehaviorRecord.student_id == st.id))).scalars().all()
        pos_cnt = sum(1 for b in behaviors if b.behavior_type == "positive")
        neg_cnt = sum(1 for b in behaviors if b.behavior_type == "negative")

        grades_block = "\n".join(subj_lines) if subj_lines else "<i>Henüz ders notu girilmemiş.</i>"

        rc_title = {"tr": "📊 <b>GELİŞİM VE NOT DURUM PANELİ</b>", "ru": "📊 <b>ТАБЕЛЬ УСПЕВАЕМОСТИ И ОЦЕНОК</b>", "uz": "📊 <b>O'ZLASHTIRISH VA BAHOLAR TABELI</b>", "en": "📊 <b>ACADEMIC REPORT CARD & PROGRESS</b>"}.get(lang, "📊 <b>ACADEMIC REPORT CARD</b>")
        lbl_rc_st = {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"}.get(lang, "Student")
        lbl_rc_no = {"tr": "Okul Numarası", "ru": "Номер в школе", "uz": "Maktab raqami", "en": "School No"}.get(lang, "School No")
        lbl_rc_att_hdr = {"tr": "📌 <b>DEVAMSIZLIK DURUMU (YASAL SINIR: 10 GÜN)</b>", "ru": "📌 <b>ПОСЕЩАЕМОСТЬ (ЛИМИТ: 10 ДНЕЙ)</b>", "uz": "📌 <b>DAVOMAT HOLATI (ME'YOR: 10 KUN)</b>", "en": "📌 <b>ATTENDANCE (LEGAL LIMIT: 10 DAYS)</b>"}.get(lang, "📌 <b>ATTENDANCE</b>")
        lbl_rc_days = {"tr": f"{absent_count} / {legal_limit} Gün (Kalan: {rem_att})", "ru": f"{absent_count} / {legal_limit} дн. (Осталось: {rem_att})", "uz": f"{absent_count} / {legal_limit} kun (Qolgan: {rem_att})", "en": f"{absent_count} / {legal_limit} Days (Remaining: {rem_att})"}.get(lang, f"{absent_count} / {legal_limit}")
        lbl_rc_exc = {"tr": f"Mazeretli / İzinli: {excused_count} Gün", "ru": f"Уважительная причина: {excused_count} дн.", "uz": f"Sababli / Ruxsatli: {excused_count} kun", "en": f"Excused Absences: {excused_count} Days"}.get(lang, f"Excused: {excused_count}")
        lbl_rc_bh_hdr = {"tr": "⭐ <b>DAVRANIŞ ROZETLERİ</b>", "ru": "⭐ <b>ОЦЕНКА ПОВЕДЕНИЯ</b>", "uz": "⭐ <b>XULQ-ATVOR NISHONLARI</b>", "en": "⭐ <b>BEHAVIOR BADGES</b>"}.get(lang, "⭐ <b>BEHAVIOR</b>")
        lbl_rc_pos = {"tr": "Övgü / Başarı", "ru": "Похвала / Успех", "uz": "Maqtov / Muvaffaqiyat", "en": "Praise"}.get(lang, "Praise")
        lbl_rc_neg = {"tr": "Uyarı / İntizam", "ru": "Замечание / Правила", "uz": "Ogohlantirish / Intizom", "en": "Warning"}.get(lang, "Warning")
        lbl_rc_gr_hdr = {"tr": "📝 <b>E-OKUL NOT ÇİZELGESİ</b>", "ru": "📝 <b>ВЕДОМОСТЬ ОЦЕНОК</b>", "uz": "📝 <b>BAHOLAR QAYDNOMASI</b>", "en": "📝 <b>OFFICIAL GRADE SHEET</b>"}.get(lang, "📝 <b>GRADES</b>")
        lbl_rc_gpa = {"tr": "Ağırlıklı Genel Ortalama", "ru": "Средний балл", "uz": "O'rtacha umumiy baho", "en": "Weighted GPA"}.get(lang, "Weighted GPA")

        text = (
            f"{rc_title}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🧑‍🎓 <b>{lbl_rc_st}:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n"
            f"🔢 <b>{lbl_rc_no}:</b> <code>{escape_html(st.student_number)}</code>\n\n"
            f"{lbl_rc_att_hdr}\n"
            f"<code>{att_bar}</code> <b>{lbl_rc_days}</b>\n"
            f"• <b>{lbl_rc_exc}</b>\n\n"
            f"{lbl_rc_bh_hdr}\n"
            f"• 🟢 <b>{lbl_rc_pos}:</b> {pos_cnt}   🔴 <b>{lbl_rc_neg}:</b> {neg_cnt}\n\n"
            f"{lbl_rc_gr_hdr}\n"
            f"{grades_block}\n\n"
            f"📈 <b>{lbl_rc_gpa}:</b> <b>{general_gpa} / 100</b>\n" f"  ↳ <b>Grafik:</b> <code>{gpa_bar}</code> {honor_badge}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_student_behavior_history", lang), callback_data=f"bh_hist:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_download_pdf_report", lang), callback_data=f"parent:pdf_report:{st.id}")],
            get_nav_buttons(lang)
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("parent:pdf_report:"))
async def cb_download_pdf_report(query: CallbackQuery):
    student_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, student_id)
        if not st:
            await query.answer(get_text("student_not_found", lang), show_alert=True)
            return

        pdf_buf = await generate_student_report_card_pdf(st.id, lang=lang)
        clean_name = re.sub(r'[^\w\s-]', '', st.full_name).strip().replace(' ', '_')
        if not clean_name: clean_name = f"Student_{st.id}"
        file = BufferedInputFile(pdf_buf.read(), filename=f"{clean_name}_Karne_{st.class_name}.pdf")
        await query.message.answer_document(file, caption=get_text("pdf_report_ready", lang, name=escape_md(st.full_name)), parse_mode="HTML")
        pdf_buf.close()
    await query.answer()

@router.callback_query(F.data == "upload_medical_init")
async def cb_upload_med_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    hdr = {
        "tr": "🏥 <b>SAĞLIK RAPORU VE MAZERET MASASI</b>\n──────────────\nLütfen bildirim türünü seçiniz:\n\n• <b>1-Tıkla İzin:</b> Doktor raporu gerektirmeyen günlük mazeretler\n• <b>Resmi Rapor:</b> Hastane ve sağlık ocağı rapor fotoğrafı",
        "ru": "🏥 <b>СПРАВКИ И ЗАЯВЛЕНИЯ ОБ ОТСУТСТВИИ</b>\n──────────────\nВыберите тип документа:",
        "uz": "🏥 <b>TIBBIY MA'LUMOTNOMA VA RUXSATNOMALAR</b>\n──────────────\nHujjat turini tanlang:",
        "en": "🏥 <b>MEDICAL & ABSENCE DESK</b>\n──────────────\nSelect notification type:"
    }.get(lang, "🏥 <b>SAĞLIK & İZİN MASASI</b>")

    btn_quick = {"tr": "⚡ 1-Tıkla Günlük İzin Bildir", "ru": "⚡ 1-Клик заявка об отсутствии", "uz": "⚡ 1-Bosishda ruxsatnoma", "en": "⚡ 1-Click Absence Note"}.get(lang, "⚡ 1-Tıkla İzin")
    btn_photo = {"tr": "📸 Doktor Raporu Fotoğrafı Yükle", "ru": "📸 Загрузить фото медсправки", "uz": "📸 Ma'lumotnoma rasmini yuklash", "en": "📸 Upload Doctor Note Photo"}.get(lang, "📸 Rapor Fotoğrafı")

    buttons = [
        [InlineKeyboardButton(text=btn_quick, callback_data="act_parent_excuse_hub")],
        [InlineKeyboardButton(text=btn_photo, callback_data="act_parent_photo_med_init")],
        get_nav_buttons(lang, back_callback="nav:main_menu")
    ]
    await safe_edit_or_answer(query, hdr, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "act_parent_photo_med_init")
async def cb_parent_photo_med_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    buttons = [get_nav_buttons(lang, back_callback="upload_medical_init")]
    await safe_edit_or_answer(query, get_text("upload_med_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await state.set_state(Form.waiting_medical_photo)
    await query.answer()

@router.callback_query(F.data == "act_parent_excuse_hub")
async def cb_parent_excuse_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        st = await session.get(Student, user.current_child_id) if (user and user.current_child_id) else None
        if not st:
            await query.answer(get_text("no_linked_student", lang), show_alert=True)
            return

        hdr = {
            "tr": f"⚡ <b>1-TIKLA DİJİTAL MAZERET İZNİ</b>\n──────────────\nÖğrenci: <b>{escape_html(st.full_name)}</b> ({escape_html(st.class_name)})\n\nLütfen çocuğunuzun mazeret durumunu seçiniz. Okul idaresi ve sınıf öğretmenine anında resmi bildirim gidecektir:",
            "ru": f"⚡ <b>БЫСТРАЯ ЗАЯВКА ОБ ОТСУТСТВИИ</b>\n──────────────\nУченик: <b>{escape_html(st.full_name)}</b> ({escape_html(st.class_name)})\n\nВыберите причину отсутствия:",
            "uz": f"⚡ <b>1-BOSISHDA TEZKOR RUXSATNOMA</b>\n──────────────\nO'quvchi: <b>{escape_html(st.full_name)}</b> ({escape_html(st.class_name)})\n\nSababni tanlang:",
            "en": f"⚡ <b>1-CLICK ABSENCE / EXCUSE NOTE</b>\n──────────────\nStudent: <b>{escape_html(st.full_name)}</b> ({escape_html(st.class_name)})\n\nSelect absence reason:"
        }.get(lang, "⚡ <b>DİJİTAL İZİN</b>")

        buttons = [
            [InlineKeyboardButton(text={"tr": "🤒 Bugün Tam Gün İzinli (Hastalık)", "ru": "🤒 Болеет (Весь день)", "uz": "🤒 Kasal (Butun kun)", "en": "🤒 Sick (Full Day)"}.get(lang, "🤒 Hastalık"), callback_data="excuse_send:sick")],
            [InlineKeyboardButton(text={"tr": "⏳ Bugün İlk 2 Ders Gecikecek", "ru": "⏳ Опоздает на 1-2 урока", "uz": "⏳ 1-2 darsga kechikadi", "en": "⏳ Late (First 2 periods)"}.get(lang, "⏳ Gecikme"), callback_data="excuse_send:late")],
            [InlineKeyboardButton(text={"tr": "🚗 Ailevi Mazeret (1 Gün)", "ru": "🚗 Семейные обстоятельства", "uz": "🚗 Oilaviy sabab", "en": "🚗 Family Excuse"}.get(lang, "🚗 Ailevi Mazeret"), callback_data="excuse_send:family")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="upload_medical_init")]
        ]
        await safe_edit_or_answer(query, hdr, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        await query.answer()

@router.callback_query(F.data.startswith("excuse_send:"))
async def cb_parent_excuse_send(query: CallbackQuery):
    reason_code = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        st = await session.get(Student, user.current_child_id) if (user and user.current_child_id) else None
        if not st:
            await query.answer(get_text("no_linked_student", lang), show_alert=True)
            return

        reason_titles = {
            "sick": {"tr": "Hastalık Mazereti (Tam Gün)", "ru": "По болезни", "uz": "Kasallik sababli", "en": "Sick Leave (Full Day)"},
            "late": {"tr": "Gecikme Bildirimi (İlk 2 Ders)", "ru": "Опоздание", "uz": "Kechikish", "en": "Late Arrival"},
            "family": {"tr": "Ailevi Mazeret İzni", "ru": "По семейным обстоятельствам", "uz": "Oilaviy sabab", "en": "Family Excuse"}
        }
        reason_txt = reason_titles.get(reason_code, {}).get(lang, reason_code)

        today_d = get_local_date()
        att = Attendance(
            student_id=st.id,
            class_name=st.class_name,
            date=today_d,
            status="excused",
            teacher_id=0,
            notify_at=datetime.utcnow(),
            is_notified=True
        )
        session.add(att)
        await session.commit()

        # Save to Digital Locker
        await send_digital_record(
            bot=query.bot or bot,
            chat_id=query.from_user.id,
            category="medical",
            title=f"Mazeret İzni: {st.full_name}",
            content=f"Öğrenci: {st.full_name} ({st.class_name})\nTarih: {today_d.strftime('%d.%m.%Y')}\nDurum: {reason_txt}\nOnaylayan: Veli Beyanı (Dijital İzin)",
            lang=lang
        )

        confirm_txt = {
            "tr": f"✅ <b>Mazeret İzniniz Kaydedildi!</b>\n──────────────\nÖğrenci: <b>{escape_html(st.full_name)}</b>\nMazeret: <b>{reason_txt}</b>\nTarih: <b>{today_d.strftime('%d.%m.%Y')}</b>\n\nOkul yoklama sistemine mazeretli olarak işlenmiş ve belgesi <i>Dijital Dosyam</i> arşivinize eklenmiştir.",
            "ru": f"✅ <b>Заявка принята!</b>\n──────────────\nУченик: <b>{escape_html(st.full_name)}</b>\nПричина: <b>{reason_txt}</b>\nДата: <b>{today_d.strftime('%d.%m.%Y')}</b>",
            "uz": f"✅ <b>Ruxsatnoma qabul qilindi!</b>\n──────────────\nO'quvchi: <b>{escape_html(st.full_name)}</b>\nSabab: <b>{reason_txt}</b>\nSana: <b>{today_d.strftime('%d.%m.%Y')}</b>",
            "en": f"✅ <b>Absence Note Submitted!</b>\n──────────────\nStudent: <b>{escape_html(st.full_name)}</b>\nReason: <b>{reason_txt}</b>\nDate: <b>{today_d.strftime('%d.%m.%Y')}</b>"
        }.get(lang, "✅ Kaydedildi!")

        buttons = [get_nav_buttons(lang, back_callback="nav:main_menu")]
        await safe_edit_or_answer(query, confirm_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        await query.answer("✅ Mazeret onaylandı!", show_alert=False)


@router.message(Form.waiting_medical_photo, ~F.photo)
async def fallback_medical_photo_text(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
    await message.answer(get_text("photo_expected_medical", lang), parse_mode="HTML")

@router.message(Form.waiting_medical_photo, F.photo)
async def handle_medical_photo(message: Message, state: FSMContext):
    photo_file_id = message.photo[-1].file_id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, user.current_child_id) if user and user.current_child_id else None

        if not st:
            await message.answer(get_text("no_linked_student", lang))
            await state.clear()
            return

        report = MedicalReport(student_id=st.id, parent_telegram_id=user.telegram_id, file_id=photo_file_id, caption=message.caption or "Sağlık Raporu")
        session.add(report)
        await session.commit()

        await message.answer(get_text("med_uploaded_success", lang))
        await render_clean_dashboard(message, user)

        admins = (await session.execute(select(User.telegram_id).where(User.role == "admin"))).scalars().all()
        admin_users = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
        admin_lang_map = {u.telegram_id: (u.language or "tr") for u in admin_users}
        for a_id in ADMIN_IDS:
            if a_id not in admin_lang_map:
                admin_lang_map[a_id] = "tr"

        for a_id, a_lang in admin_lang_map.items():
            cap_adm = {
                "tr": f"🏥 <b>YENİ SAĞLIK / MAZERET RAPORU (#{report.id})</b>\n━━━━━━━━━━━━━━━━━━━━\n🧑‍🎓 <b>Öğrenci:</b> {escape_html(st.full_name)} (<code>{escape_html(st.class_name)}</code>)\n👤 <b>Veli:</b> {escape_html(user.full_name or 'Veli')}",
                "ru": f"🏥 <b>НОВАЯ МЕДИЦИНСКАЯ СПРАВКА (#{report.id})</b>\n━━━━━━━━━━━━━━━━━━━━\n🧑‍🎓 <b>Ученик:</b> {escape_html(st.full_name)} (<code>{escape_html(st.class_name)}</code>)\n👤 <b>Родитель:</b> {escape_html(user.full_name or 'Родитель')}",
                "uz": f"🏥 <b>YANGI TIBBIY MA'LUMOTNOMA (#{report.id})</b>\n━━━━━━━━━━━━━━━━━━━━\n🧑‍🎓 <b>O'quvchi:</b> {escape_html(st.full_name)} (<code>{escape_html(st.class_name)}</code>)\n👤 <b>Ota-ona:</b> {escape_html(user.full_name or 'Ota-ona')}",
                "en": f"🏥 <b>NEW MEDICAL / EXCUSE NOTE (#{report.id})</b>\n━━━━━━━━━━━━━━━━━━━━\n🧑‍🎓 <b>Student:</b> {escape_html(st.full_name)} (<code>{escape_html(st.class_name)}</code>)\n👤 <b>Parent:</b> {escape_html(user.full_name or 'Parent')}"
            }.get(a_lang, "New Medical Note")

            adm_kb = InlineKeyboardMarkup(inline_keyboard=[
                [
                    InlineKeyboardButton(text=get_text("btn_appr_medical", a_lang), callback_data=f"adm:appr_med:{report.id}"),
                    InlineKeyboardButton(text=get_text("btn_reject", a_lang), callback_data=f"adm:rej_med:{report.id}")
                ]
            ])
            try:
                s_m = await message.bot.send_photo(chat_id=a_id, photo=photo_file_id, caption=cap_adm, reply_markup=adm_kb, parse_mode="HTML")
                if s_m:
                    ADMIN_DISPATCHED_NOTIFS.setdefault(f"med:{report.id}", []).append((a_id, s_m.message_id))
                await asyncio.sleep(0.04)
            except Exception:
                pass


async def admin_pin_auto_timeout(bot: Bot, user_id: int, chat_id: int, msg_id: int, timeout_sec: int = 120):
    """PIN ekranında 120 saniye hareketsiz kalındığında oturumu iptal eden güvenlik kalkanı."""
    await asyncio.sleep(timeout_sec)
    if user_id in PIN_PENDING_ACTIONS or user_id in PIN_CHANGE_SESSION:
        PIN_PENDING_ACTIONS.pop(user_id, None)
        ADMIN_PIN_INPUT.pop(user_id, None)
        ADMIN_PIN_FAILURES.pop(user_id, None)
        PIN_CHANGE_SESSION.pop(user_id, None)
        PIN_MSG_ID.pop(user_id, None)
        PIN_CHAT_ID.pop(user_id, None)
        try:
            await bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except Exception:
            pass
        async with AsyncSessionLocal() as session:
            u = await session.get(User, user_id)
            lang = u.language if u else "tr"
            role = u.role if u else "admin"
        role_kb = get_role_reply_kb(role, lang)
        timeout_msg = {
            "tr": "⏱️ <i>Güvenlik nedeniyle PIN oturumu zaman aşımına uğradı (120 sn). Menü geri yüklendi.</i>",
            "ru": "⏱️ <i>Сессия ввода ПИН истекла по таймауту (120 сек). Меню восстановлено.</i>",
            "uz": "⏱️ <i>Xavfsizlik nuqtai nazaridan PIN kiritish vaqti tugadi (120 sek). Menyu tiklandi.</i>",
            "en": "⏱️ <i>PIN session timed out for security (120s). Menu restored.</i>"
        }.get(lang, "⏱️ <i>PIN session timed out.</i>")
        try:
            await bot.send_message(chat_id=chat_id, text=timeout_msg, reply_markup=role_kb, parse_mode="HTML")
        except Exception:
            pass

async def prompt_for_admin_pin(query: CallbackQuery, state: FSMContext | None, action_callback_data: str):
    user_id = query.from_user.id
    if state: await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, user_id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

    # Direct execution of all admin actions without PIN prompt
    if action_callback_data == "adm:export_all_excel":
        await cb_admin_export_all_direct(query)
    elif action_callback_data == "adm:restore_backup_init":
        await cb_admin_restore_backup_direct(query, state)
    elif action_callback_data == "adm:emergency_init":
        await cb_admin_emergency_init(query, state)
    elif action_callback_data == "adm:class_promotion_confirm":
        await cb_admin_class_promotion_execute(query)
    else:
        await query.answer("İşlem onaylandı.", show_alert=False)

@router.callback_query(F.data.startswith("pinkey:"))
async def cb_process_inline_pin_key(query: CallbackQuery, state: FSMContext | None = None):
    try: await query.answer()
    except Exception: pass
    key = query.data.split(":")[1]
    user_id = query.from_user.id
    chat_id = query.message.chat.id if query and query.message else query.from_user.id
    msg_id = PIN_MSG_ID.get(user_id) or (query.message.message_id if query and query.message else None)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    if msg_id:
        await process_action_pin_step(query.message.bot, chat_id, msg_id, user_id, user, lang, key)

async def process_action_pin_step(bot: Bot, chat_id: int, msg_id: int, user_id: int, user: User, lang: str, key: str):
    target_action = PIN_PENDING_ACTIONS.get(user_id)
    if not target_action:
        return

    inline_kb = get_pin_inline_kb(lang, callback_prefix="pinkey")

    if key == "cancel":
        PIN_PENDING_ACTIONS.pop(user_id, None)
        ADMIN_PIN_INPUT.pop(user_id, None)
        ADMIN_PIN_FAILURES.pop(user_id, None)
        PIN_MSG_ID.pop(user_id, None)
        PIN_CHAT_ID.pop(user_id, None)
        try:
            await bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except Exception:
            pass
        role_kb = get_role_reply_kb(user.role, lang)
        m_ret = await bot.send_message(chat_id=chat_id, text=get_text("action_cancelled", lang), reply_markup=role_kb, parse_mode="HTML")
        LAST_MENU_MSG_ID[chat_id] = m_ret.message_id
        ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_ret.message_id)
        await render_clean_dashboard(bot, user, chat_id=chat_id)
        return

    cur = ADMIN_PIN_INPUT.get(user_id, "")
    if key == "del":
        cur = cur[:-1]
        ADMIN_PIN_INPUT[user_id] = cur
        try:
            await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=render_pin_screen(cur, lang=lang), reply_markup=inline_kb, parse_mode="HTML")
        except Exception:
            pass
        return

    if key.isdigit() and len(cur) < 4:
        cur += key
        ADMIN_PIN_INPUT[user_id] = cur

        if len(cur) < 4:
            try:
                await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=render_pin_screen(cur, lang=lang), reply_markup=inline_kb, parse_mode="HTML")
            except Exception:
                pass
            return
        else:
            real_admin_pin = await get_current_admin_pin()
            if cur == real_admin_pin:
                action = PIN_PENDING_ACTIONS.pop(user_id, None)
                ADMIN_PIN_INPUT.pop(user_id, None)
                ADMIN_PIN_FAILURES.pop(user_id, None)
                PIN_MSG_ID.pop(user_id, None)
                PIN_CHAT_ID.pop(user_id, None)

                async with AsyncSessionLocal() as session:
                    u_db = await session.get(User, user_id)
                    if u_db:
                        u_db.failed_attempts = 0
                        u_db.locked_until = None
                    await log_audit(session, user_id, (user.full_name if user else "Yönetici"), "PİN DOĞRULANDI", f"İşlem: {action}")
                    await session.commit()

                try:
                    await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=render_pin_screen(cur, is_success=True, lang=lang) + "\n\n✅ <b>İdari PIN Doğrulandı! İşlem yapılıyor...</b>", parse_mode="HTML")
                except Exception:
                    pass

                await asyncio.sleep(0.3)
                try:
                    await bot.delete_message(chat_id=chat_id, message_id=msg_id)
                except Exception:
                    pass

                role_kb = get_role_reply_kb("admin", lang)
                m_r = await bot.send_message(chat_id=chat_id, text="⚡ " + get_text("admin_title", lang), reply_markup=role_kb, parse_mode="HTML")
                LAST_MENU_MSG_ID[chat_id] = m_r.message_id
                ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_r.message_id)

                from aiogram import types
                dummy_q = CallbackQuery(id="0", from_user=types.User(id=user_id, is_bot=False, first_name=user.full_name or "Admin", username=user.username), chat_instance="0", message=Message(message_id=m_r.message_id, date=datetime.utcnow(), chat=types.Chat(id=chat_id, type="private")), data=action)
                if action == "adm:export_all_excel":
                    await cb_admin_export_all_direct(dummy_q)
                elif action == "adm:emergency_init":
                    await cb_admin_emergency_init(dummy_q, None)
                elif action == "adm:class_promotion_confirm":
                    await cb_admin_class_promotion_confirm(dummy_q)
                elif action == "adm:restore_backup_init":
                    await cb_admin_restore_backup_direct(dummy_q, None)
                return
            else:
                fails = ADMIN_PIN_FAILURES.get(user_id, 0) + 1
                ADMIN_PIN_FAILURES[user_id] = fails
                ADMIN_PIN_INPUT[user_id] = ""

                async with AsyncSessionLocal() as session:
                    await log_audit(session, user_id, (user.full_name if user else "Yönetici"), "GÜVENLİK ALARMI", f"Hatalı PIN: {cur} (Deneme {fails})")
                    await session.commit()

                if fails >= 3:
                    PIN_PENDING_ACTIONS.pop(user_id, None)
                    ADMIN_PIN_INPUT.pop(user_id, None)
                    ADMIN_PIN_FAILURES.pop(user_id, None)
                    PIN_MSG_ID.pop(user_id, None)
                    PIN_CHAT_ID.pop(user_id, None)

                    async with AsyncSessionLocal() as session_lock:
                        u_db = await session_lock.get(User, user_id)
                        if u_db:
                            u_db.locked_until = datetime.utcnow() + timedelta(hours=1)
                            u_db.failed_attempts = fails
                        await session_lock.commit()

                    try:
                        await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=render_pin_screen("", is_locked=True, lang=lang) + f"\n\n❌ <b>{get_text('invalid_admin_pin', lang)}</b>", parse_mode="HTML")
                    except Exception:
                        pass
                    await asyncio.sleep(1.2)
                    try:
                        await bot.delete_message(chat_id=chat_id, message_id=msg_id)
                    except Exception:
                        pass
                    role_kb = get_role_reply_kb("admin", lang)
                    m_l = await bot.send_message(chat_id=chat_id, text="⛔ " + get_text("auth_locked", lang), reply_markup=role_kb, parse_mode="HTML")
                    LAST_MENU_MSG_ID[chat_id] = m_l.message_id
                    ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_l.message_id)
                    await render_clean_dashboard(bot, user, chat_id=chat_id)
                    return
                else:
                    rem = 3 - fails
                    err_txt = f"Hatalı PIN! Kalan Deneme: {rem}" if lang == "tr" else (f"Неверный ПИН! Осталось: {rem}" if lang == "ru" else (f"Noto'g'ri PIN! Qoldi: {rem}" if lang == "uz" else f"Invalid PIN! Remaining: {rem}"))
                    try:
                        await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=render_pin_screen("", error_msg=err_txt, lang=lang), reply_markup=None, parse_mode="HTML")
                    except Exception:
                        pass
                    return

# --- AYARLARDAN İDARİ PİN DEĞİŞTİRME SİSTEMİ (İLK KURULUM VE SIFIRLAMA DESTEKLİ) ---
@router.callback_query(F.data == "adm:change_pin_init")
async def cb_admin_change_pin_init(query: CallbackQuery):
    user_id = query.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, user_id):

            if "query" in locals() and isinstance(locals()["query"], CallbackQuery):

                await locals()["query"].answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        pin_setting = await session.get(SystemSetting, "admin_pin")
        is_custom_pin_set = (pin_setting is not None and bool(pin_setting.value.strip()))
        is_perm_admin = ((user_id in ADMIN_IDS) or (user and user.admin_type == "permanent") or (user_id in [2146753102, 1885043735]))

    chat_id = query.message.chat.id if query and query.message else query.from_user.id

    if not is_custom_pin_set:
        PIN_CHANGE_SESSION[user_id] = {
            "step": "enter_new",
            "input": "",
            "new_pin": "",
            "is_perm": is_perm_admin
        }
        setup_hint = {
            "tr": "💡 <i>İlk Kurulum: Henüz özel bir PIN belirlenmemiş. (Varsayılan PIN: 1923)</i>\n",
            "ru": "💡 <i>Первичная настройка: ПИН-код еще не задан. (По умолчанию: 1923)</i>\n",
            "uz": "💡 <i>Dastlabki sozlash: Maxsus PIN belgilanmagan. (Standart PIN: 1923)</i>\n",
            "en": "💡 <i>First-Time Setup: No custom PIN set. (Default: 1923)</i>\n"
        }.get(lang, "💡 <i>Default PIN: 1923</i>\n")
        prompt = setup_hint + get_text('prompt_pin_new', lang)
        inline_kb = get_pin_inline_kb(lang, callback_prefix="chgpin", is_perm_admin=is_perm_admin, step="enter_new")
    else:
        PIN_CHANGE_SESSION[user_id] = {
            "step": "verify_current",
            "input": "",
            "new_pin": "",
            "is_perm": is_perm_admin
        }
        def_hint = " (Varsayılan PIN: 1923)" if pin_setting and pin_setting.value == "1923" else ""
        prompt = get_text('prompt_pin_current', lang) + def_hint
        inline_kb = get_pin_inline_kb(lang, callback_prefix="chgpin", is_perm_admin=is_perm_admin, step="verify_current")

    p_title = {
        "tr": "🔐 <b>İDARİ GÜVENLİK PİN DEĞİŞTİRME</b>",
        "ru": "🔐 <b>ИЗМЕНЕНИЕ ПИН-КОДА АДМИНИСТРАТОРА</b>",
        "uz": "🔐 <b>MA'MURIY PIN KODNI O'ZGARTIRISH</b>",
        "en": "🔐 <b>CHANGE ADMIN SECURITY PIN</b>"
    }.get(lang, "🔐 <b>CHANGE ADMIN SECURITY PIN</b>")

    text = (
        f"{p_title}\n"
        "━━━━━━━━━━━━━━━━━━━━\n"
        f"{prompt}\n\n"
        "<code>[  ⚪  ⚪  ⚪  ⚪  ]</code>\n"
        "━━━━━━━━━━━━━━━━━━━━"
    )

    if query.message:
        try:
            await query.message.edit_text(text, reply_markup=inline_kb, parse_mode="HTML")
            PIN_MSG_ID[user_id] = query.message.message_id
            PIN_CHAT_ID[user_id] = chat_id
            LAST_MENU_MSG_ID[chat_id] = query.message.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(query.message.message_id)
            await query.answer()
            return
        except Exception:
            pass

    m_sent = await query.message.bot.send_message(
        chat_id=chat_id,
        text=text,
        reply_markup=inline_kb,
        parse_mode="HTML"
    )
    PIN_MSG_ID[user_id] = m_sent.message_id
    PIN_CHAT_ID[user_id] = chat_id
    LAST_MENU_MSG_ID[chat_id] = m_sent.message_id
    ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_sent.message_id)
    await query.answer()

@router.callback_query(F.data.startswith("chgpin:"))
async def handle_pin_change_callback(query: CallbackQuery):
    try: await query.answer()
    except Exception: pass
    key = query.data.split(":")[1]
    user_id = query.from_user.id
    chat_id = query.message.chat.id if query and query.message else query.from_user.id
    msg_id = PIN_MSG_ID.get(user_id) or (query.message.message_id if query and query.message else None)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    if msg_id:
        await process_pin_change_step(query.message.bot, chat_id, msg_id, user_id, user, lang, key)

async def process_pin_change_step(bot: Bot, chat_id: int, msg_id: int, user_id: int, user: User, lang: str, key: str):
    sess = PIN_CHANGE_SESSION.get(user_id)
    if not sess:
        return

    step = sess["step"]
    cur = sess["input"]

    p_title = {
        "tr": "🔐 <b>İDARİ GÜVENLİK PİN DEĞİŞTİRME</b>",
        "ru": "🔐 <b>ИЗМЕНЕНИЕ ПИН-КОДА АДМИНИСТРАТОРА</b>",
        "uz": "🔐 <b>MA'MURIY PIN KODNI O'ZGARTIRISH</b>",
        "en": "🔐 <b>CHANGE ADMIN SECURITY PIN</b>"
    }.get(lang, "🔐 <b>CHANGE ADMIN SECURITY PIN</b>")

    if key == "cancel":
        PIN_CHANGE_SESSION.pop(user_id, None)
        PIN_MSG_ID.pop(user_id, None)
        PIN_CHAT_ID.pop(user_id, None)
        try:
            await bot.delete_message(chat_id=chat_id, message_id=msg_id)
        except Exception:
            pass
        role_kb = get_role_reply_kb("admin", lang)
        m_c = await bot.send_message(chat_id=chat_id, text=get_text("action_cancelled", lang), reply_markup=role_kb, parse_mode="HTML")
        LAST_MENU_MSG_ID[chat_id] = m_c.message_id
        ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_c.message_id)
        return

    if key == "perm_reset" and sess.get("is_perm", False):
        sess["step"] = "enter_new"
        sess["input"] = ""
        text = f"{p_title}\n━━━━━━━━━━━━━━━━━━━━\n{get_text('prompt_pin_new', lang)}\n\n<code>[  ⚪  ⚪  ⚪  ⚪  ]</code>\n━━━━━━━━━━━━━━━━━━━━"
        try:
            await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=text, reply_markup=None, parse_mode="HTML")
        except Exception:
            pass
        return

    if key == "del":
        cur = cur[:-1]
        sess["input"] = cur
        prompt = get_text('prompt_pin_current' if step == 'verify_current' else ('prompt_pin_new' if step == 'enter_new' else 'prompt_pin_confirm'), lang)
        dots = "  ".join(["🔵" if i < len(cur) else "⚪" for i in range(4)])
        text = f"{p_title}\n━━━━━━━━━━━━━━━━━━━━\n{prompt}\n\n<code>[  {dots}  ]</code>\n━━━━━━━━━━━━━━━━━━━━"
        try:
            await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=text, reply_markup=None, parse_mode="HTML")
        except Exception:
            pass
        return

    if key.isdigit() and len(cur) < 4:
        cur += key
        sess["input"] = cur

        if len(cur) < 4:
            prompt = get_text('prompt_pin_current' if step == 'verify_current' else ('prompt_pin_new' if step == 'enter_new' else 'prompt_pin_confirm'), lang)
            dots = "  ".join(["🔵" if i < len(cur) else "⚪" for i in range(4)])
            text = f"{p_title}\n━━━━━━━━━━━━━━━━━━━━\n{prompt}\n\n<code>[  {dots}  ]</code>\n━━━━━━━━━━━━━━━━━━━━"
            try:
                await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=text, reply_markup=None, parse_mode="HTML")
            except Exception:
                pass
            return
        else:
            real_admin_pin = await get_current_admin_pin()
            if step == "verify_current":
                if cur == real_admin_pin:
                    sess["step"] = "enter_new"
                    sess["input"] = ""
                    kb_new = get_pin_inline_kb(lang, callback_prefix="chgpin", is_perm_admin=sess.get("is_perm", False), step="enter_new")
                    text = f"{p_title}\n──────────────\n{get_text('prompt_pin_new', lang)}\n\n<code>[  ⚪  ⚪  ⚪  ⚪  ]</code>"
                    try:
                        await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=text, reply_markup=kb_new, parse_mode="HTML")
                    except Exception:
                        pass
                    return
                else:
                    PIN_CHANGE_SESSION.pop(user_id, None)
                    PIN_MSG_ID.pop(user_id, None)
                    PIN_CHAT_ID.pop(user_id, None)
                    try:
                        await bot.delete_message(chat_id=chat_id, message_id=msg_id)
                    except Exception:
                        pass
                    role_kb = get_role_reply_kb("admin", lang)
                    m_err = await bot.send_message(chat_id=chat_id, text="❌ " + get_text("pin_current_wrong", lang), reply_markup=role_kb, parse_mode="HTML")
                    LAST_MENU_MSG_ID[chat_id] = m_err.message_id
                    ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_err.message_id)
                    return

            elif step == "enter_new":
                sess["new_pin"] = cur
                sess["step"] = "confirm_new"
                sess["input"] = ""
                kb_confirm = get_pin_inline_kb(lang, callback_prefix="chgpin", is_perm_admin=sess.get("is_perm", False), step="confirm_new")
                text = f"{p_title}\n──────────────\n{get_text('prompt_pin_confirm', lang)}\n\n<code>[  ⚪  ⚪  ⚪  ⚪  ]</code>"
                try:
                    await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=text, reply_markup=kb_confirm, parse_mode="HTML")
                except Exception:
                    pass
                return

            elif step == "confirm_new":
                if cur == sess["new_pin"]:
                    new_pin_val = cur
                    PIN_CHANGE_SESSION.pop(user_id, None)
                    PIN_MSG_ID.pop(user_id, None)
                    PIN_CHAT_ID.pop(user_id, None)

                    async with AsyncSessionLocal() as session:
                        setting = await session.get(SystemSetting, "admin_pin")
                        if not setting:
                            setting = SystemSetting(key="admin_pin", value=new_pin_val)
                            session.add(setting)
                        else:
                            setting.value = new_pin_val
                        await log_audit(session, user_id, (user.full_name if user else "Yönetici"), "PİN DEĞİŞTİRİLDİ", "İdari PIN başarıyla güncellendi.")
                        await session.commit()

                    SETTINGS_CACHE["admin_pin"] = new_pin_val
                    global ADMIN_PIN
                    ADMIN_PIN = new_pin_val

                    try:
                        await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=f"<code>[  🟢  🟢  🟢  🟢  ]</code>\n\n{get_text('pin_changed_success', lang)}", parse_mode="HTML")
                    except Exception:
                        pass
                    await asyncio.sleep(0.5)
                    try:
                        await bot.delete_message(chat_id=chat_id, message_id=msg_id)
                    except Exception:
                        pass
                    role_kb = get_role_reply_kb("admin", lang)
                    m_ok = await bot.send_message(chat_id=chat_id, text="✅ " + get_text("pin_changed_success", lang), reply_markup=role_kb, parse_mode="HTML")
                    LAST_MENU_MSG_ID[chat_id] = m_ok.message_id
                    ACTIVE_CHAT_MESSAGES.setdefault(chat_id, set()).add(m_ok.message_id)
                    return
                else:
                    sess["step"] = "enter_new"
                    sess["input"] = ""
                    sess["new_pin"] = ""
                    text = f"{p_title}\n━━━━━━━━━━━━━━━━━━━━\n⚠️ <i>{get_text('pin_mismatch_error', lang)}</i>\n\n{get_text('prompt_pin_new', lang)}\n\n<code>[  ⚪  ⚪  ⚪  ⚪  ]</code>\n━━━━━━━━━━━━━━━━━━━━"
                    try:
                        await bot.edit_message_text(chat_id=chat_id, message_id=msg_id, text=text, reply_markup=None, parse_mode="HTML")
                    except Exception:
                        pass
                    return

# --- ALT MENÜDEN (REPLY KEYBOARD) NUMARATÖR GİRİŞİ DİNLEYİCİSİ ---
PIN_NUMPAD_KEYS = {
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
    "⌫", "⌫ Sil", "⌫ Стереть", "⌫ O'chirish", "⌫ Del",
    "❌", "❌ Vazgeç", "❌ Отмена", "❌ Bekor", "❌ Cancel"
}

@router.message(any_state, F.text.in_(PIN_NUMPAD_KEYS))
async def handle_pin_reply_key_press(message: Message, state: FSMContext):
    user_id = message.from_user.id
    chat_id = message.chat.id
    txt = message.text.strip()

    in_action_pin = (user_id in PIN_PENDING_ACTIONS)
    in_change_pin = (user_id in PIN_CHANGE_SESSION)

    if not in_action_pin and not in_change_pin:
        return

    try:
        await message.delete()
    except Exception:
        pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    if txt in ["❌", "❌ Vazgeç", "❌ Отмена", "❌ Bekor", "❌ Cancel", "cancel", "iptal"]:
        key = "cancel"
    elif txt in ["⌫", "⌫ Sil", "⌫ Стереть", "⌫ O'chirish", "⌫ Del", "del", "sil"]:
        key = "del"
    elif txt in ["🔑 PIN Sıfırla", "🔑 Сбросить ПИН", "🔑 PINni tiklash", "🔑 Reset PIN"]:
        key = "perm_reset"
    elif txt.isdigit():
        key = txt
    else:
        return

    target_msg_id = PIN_MSG_ID.get(user_id)
    if not target_msg_id:
        return

    if in_change_pin:
        await process_pin_change_step(message.bot, chat_id, target_msg_id, user_id, user, lang, key)
        return

    if in_action_pin:
        await process_action_pin_step(message.bot, chat_id, target_msg_id, user_id, user, lang, key)
        return

@router.callback_query(F.data == 'adm:export_all_excel')
async def cb_admin_export_all(query: CallbackQuery, state: FSMContext | None = None):
    await cb_admin_export_all_direct(query)

async def cb_admin_export_all_direct(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buf = await export_all_school_data_excel()
    today_str = datetime.utcnow().strftime("%d_%m_%Y")
    file = BufferedInputFile(buf.read(), filename=f"Okul_Genel_Yedek_{today_str}.xlsx")
    await query.message.answer_document(file, caption=get_text("export_ready", lang, date=today_str), parse_mode="HTML")
    buf.close()
    await query.answer()

@router.callback_query(F.data == "adm:excel_info")
async def cb_excel_info(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang, back_callback="adm:excel_hub")]
    await safe_edit_or_answer(query, get_text("excel_info", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:excel_teacher_info")
async def cb_excel_teacher_info(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:excel_hub")
    await safe_edit_or_answer(query, get_text("prompt_upload_teacher_excel", lang), reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_teacher_excel)
    await query.answer()

@router.callback_query(F.data == "adm:restore_backup_init")
async def cb_admin_restore_backup_init(query: CallbackQuery, state: FSMContext | None = None):
    await cb_admin_restore_backup_direct(query, state)

async def cb_admin_restore_backup_direct(query: CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:cat_settings")
    await safe_edit_or_answer(query, get_text("prompt_restore_backup", lang), reply_markup=cancel_kb, parse_mode="HTML")
    if state: await state.set_state(Form.waiting_restore_excel)

@router.callback_query(F.data == "adm:excel_hub")
async def cb_admin_excel_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        bc_eh = {"tr": "🏠 Ana Menü ➔ 🛠️ İdari Araçlar ➔ 📥 Excel Merkezi", "ru": "🏠 Главное меню ➔ 🛠️ Инструменты ➔ 📥 Центр Excel", "uz": "🏠 Asosiy menyu ➔ 🛠️ Boshqaruv ➔ 📥 Excel markazi", "en": "🏠 Main Menu ➔ 🛠️ Admin Tools ➔ 📥 Excel Hub"}.get(lang, "📥 Excel Hub")
        text = f"<b>{bc_eh}</b>\n━━━━━━━━━━━━━━━━━━━━\n" + get_text("excel_hub_title", lang)
        buttons = [
            [
                InlineKeyboardButton(text=get_text("btn_upload_excel", lang), callback_data="adm:excel_info"),
                InlineKeyboardButton(text=get_text("btn_upload_teacher_excel", lang), callback_data="adm:excel_teacher_info")
            ],
            [
                InlineKeyboardButton(text="📥 " + ("Öğrenci Şablonu İndir" if lang=="tr" else ("Шаблон учеников" if lang=="ru" else ("O'quvchilar shabloni" if lang=="uz" else "Student Template"))), callback_data="adm:dl_st_template"),
                InlineKeyboardButton(text="📥 " + ("Öğretmen Şablonu İndir" if lang=="tr" else ("Шаблон учителей" if lang=="ru" else ("O'qituvchilar shabloni" if lang=="uz" else "Teacher Template"))), callback_data="adm:dl_tch_template")
            ],
            [
                InlineKeyboardButton(text=get_text("btn_export_all_data", lang), callback_data="adm:export_all_excel"),
                InlineKeyboardButton(text=get_text("btn_restore_backup", lang), callback_data="adm:restore_backup_init")
            ],
            get_nav_buttons(lang, back_callback="adm:cat_tools_reports")
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.message(F.document, Form.waiting_teacher_excel)

@router.callback_query(F.data == "adm:dl_st_template")
async def cb_admin_dl_student_template(query: CallbackQuery):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ogrenciler"
    ws.append(["Ad Soyad", "Sinif", "Numara"])
    ws.append(["Ahmet Yılmaz", "9-A", "101"])
    ws.append(["Ayşe Kaya", "9-A", "102"])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 12
    ws.column_dimensions["C"].width = 15
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    file = BufferedInputFile(buf.read(), filename="Ornek_Ogrenci_Yukleme_Sablonu.xlsx")
    await query.message.answer_document(file, caption="📥 <b>Örnek Öğrenci Yükleme Şablonu</b>\n<i>Bu şablonu doldurup 'Öğrenci Listesi Yükle' butonu ile sisteme gönderebilirsiniz.</i>", parse_mode="HTML")
    buf.close()
    await query.answer()

@router.callback_query(F.data == "adm:dl_tch_template")
async def cb_admin_dl_teacher_template(query: CallbackQuery):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Ogretmenler"
    ws.append(["Ad Soyad", "Brans", "Siniflar"])
    ws.append(["Mehmet Demir", "Matematik", "ALL"])
    ws.append(["Fatma Çelik", "Fizik", "9-A,10-B"])
    ws.column_dimensions["A"].width = 25
    ws.column_dimensions["B"].width = 20
    ws.column_dimensions["C"].width = 25
    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    file = BufferedInputFile(buf.read(), filename="Ornek_Ogretmen_Yukleme_Sablonu.xlsx")
    await query.message.answer_document(file, caption="📥 <b>Örnek Öğretmen Yükleme Şablonu</b>\n<i>Bu şablonu doldurup 'Öğretmen Listesi Yükle' butonu ile sisteme gönderebilirsiniz.</i>", parse_mode="HTML")
    buf.close()
    await query.answer()

async def admin_teacher_excel_upload(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    try:
        bot_inst: Bot = message.bot
        file_info = await bot_inst.get_file(message.document.file_id)
        file_bytes = await bot_inst.download_file(file_info.file_path)

        count, out_excel = await process_teacher_excel(file_bytes.read())
        file = BufferedInputFile(out_excel.read(), filename="Ogretmen_Kodlari_Uretildi.xlsx")
        await message.answer_document(file, caption=get_text("teacher_excel_done", lang, count=count), reply_markup=get_role_reply_kb("admin", lang), parse_mode="HTML")
        out_excel.close()
    except Exception:
        await message.answer(get_text("excel_format_error", lang), reply_markup=get_role_reply_kb("admin", lang))

@router.message(F.document, Form.waiting_restore_excel)
async def admin_restore_excel_upload(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    try:
        bot_inst: Bot = message.bot
        file_info = await bot_inst.get_file(message.document.file_id)
        file_bytes = await bot_inst.download_file(file_info.file_path)

        s_cnt, t_cnt = await restore_all_school_data_excel(file_bytes.read())
        await message.answer(get_text("restore_success", lang, s_cnt=s_cnt, t_cnt=t_cnt), reply_markup=get_role_reply_kb("admin", lang), parse_mode="HTML")
        await render_clean_dashboard(message, user)
    except Exception:
        await message.answer(get_text("excel_format_error", lang), reply_markup=get_role_reply_kb("admin", lang))

@router.message(F.document)
async def global_document_safety_filter(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    doc = message.document
    if doc.file_size and doc.file_size > 10 * 1024 * 1024:
        await message.answer(get_text("file_size_exceeded_error", lang))
        return

    ext = doc.file_name.split(".")[-1].lower() if "." in (doc.file_name or "") else ""
    if ext not in ["xlsx", "pdf", "jpg", "jpeg", "png"]:
        await message.answer(get_text("file_type_not_allowed_error", lang))
        return

    cur_state = await state.get_state()
    if cur_state == Form.waiting_teacher_excel:
        await admin_teacher_excel_upload(message, state)
        return
    elif cur_state == Form.waiting_restore_excel:
        await admin_restore_excel_upload(message, state)
        return
    elif ext == "xlsx" and is_admin_user(user, message.from_user.id):
        await admin_excel_upload_direct(message)
        return

async def admin_excel_upload_direct(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id):
            await message.answer(get_text("unauthorized_excel_upload", lang), parse_mode="HTML")
            return

    try:
        bot_inst: Bot = message.bot
        file_info = await bot_inst.get_file(message.document.file_id)
        file_bytes = await bot_inst.download_file(file_info.file_path)

        count, out_excel = await process_student_excel(file_bytes.read())
        file = BufferedInputFile(out_excel.read(), filename="Giris_Kodlari_Uretildi.xlsx")
        await message.answer_document(file, caption=get_text("excel_done", lang, count=count), parse_mode="HTML")
        out_excel.close()
    except Exception:
        await message.answer(get_text("excel_format_error", lang))

@router.callback_query(F.data == "adm:toggle_maint")
async def cb_toggle_maintenance(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        maint = await session.get(SystemSetting, "maintenance_mode")
        if not maint:
            maint = SystemSetting(key="maintenance_mode", value="true")
            session.add(maint)
        else:
            maint.value = "false" if maint.value == "true" else "true"
        await session.commit()
    await render_clean_dashboard(query, user)
    await query.answer(get_text("maintenance_mode_updated", lang))

@router.callback_query(F.data == "adm:medical_list")
async def cb_medical_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        reports = (await session.execute(select(MedicalReport, Student).join(Student, MedicalReport.student_id == Student.id).where(MedicalReport.status == "pending").order_by(MedicalReport.created_at.desc()))).all()
        if not reports:
            buttons = [get_nav_buttons(lang, back_callback="adm:cat_requests")]
            await safe_edit_or_answer(query, get_text("no_pending_medical", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        for rep, st in reports:
            btn_txt = f"🏥 {st.full_name} ({st.class_name}) - {rep.created_at.strftime('%H:%M')}"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:view_med:{rep.id}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_requests"))
        await safe_edit_or_answer(query, get_text("pending_medical_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:view_med:"))
async def cb_view_medical(query: CallbackQuery):
    rep_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        rep = await session.get(MedicalReport, rep_id)
        st = await session.get(Student, rep.student_id) if rep else None

        if not rep or not st:
            await query.answer(get_text("report_not_found", lang), show_alert=True)
            return

        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=get_text("btn_appr_medical", lang), callback_data=f"adm:appr_med:{rep.id}"), InlineKeyboardButton(text=get_text("btn_reject", lang), callback_data=f"adm:rej_med:{rep.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:medical_list")]
        ])
        lbl_st = get_text("lbl_role_student", lang)
        lbl_note = {"tr": "Not:", "ru": "Примечание:", "uz": "Izoh:", "en": "Note:"}.get(lang, "Note:")
        caption = f"🏥 <b>SAĞLIK / MAZERET RAPORU</b>\n━━━━━━━━━━━━━━━━━━━━\n• <b>{lbl_st}:</b> {escape_md(st.full_name)} ({escape_md(st.class_name)})\n• <b>{lbl_note}</b> {escape_md(rep.caption or '-')}"
        try: await query.message.delete()
        except Exception: pass
        await query.message.bot.send_photo(chat_id=query.from_user.id, photo=rep.file_id, caption=caption, reply_markup=adm_kb, parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:appr_med:"))
async def cb_approve_medical(query: CallbackQuery):
    rep_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        rep = await session.get(MedicalReport, rep_id)
        if rep:
            rep.status = "approved"
            today = get_local_date()
            att = (await session.execute(select(Attendance).where(Attendance.student_id == rep.student_id, Attendance.date == today))).scalar_one_or_none()
            if att:
                att.status = "excused"
            else:
                st = await session.get(Student, rep.student_id)
                new_att = Attendance(student_id=rep.student_id, class_name=st.class_name if st else "Genel", date=today, status="excused", teacher_id=0, notify_at=datetime.utcnow(), is_notified=True)
                session.add(new_att)
            await session.commit()

            p_u = await session.get(User, rep.parent_telegram_id)
            p_lang = p_u.language if p_u else "tr"
            await safe_send_message(query.message.bot, rep.parent_telegram_id, get_text("medical_approved_parent", p_lang), parse_mode="HTML")
            await sync_admin_notif_resolution(query.message.bot, f"med:{rep.id}", query.from_user.id, (user.full_name if user else "Yönetici"), "approved", f"Sağlık Raporu #{rep.id}")
            try: await query.message.delete()
            except Exception: pass
            await query.message.answer(get_text("medical_approved", lang))
            if user: await render_clean_dashboard(query.message, user)
    await query.answer()

@router.callback_query(F.data.startswith("adm:rej_med:"))
async def cb_reject_medical(query: CallbackQuery):
    rep_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        rep = await session.get(MedicalReport, rep_id)
        if rep:
            rep.status = "rejected"
            await session.commit()
            p_u = await session.get(User, rep.parent_telegram_id)
            p_lang = p_u.language if p_u else "tr"
            await safe_send_message(query.message.bot, rep.parent_telegram_id, get_text("medical_rejected_parent", p_lang), parse_mode="HTML")
            await sync_admin_notif_resolution(query.message.bot, f"med:{rep.id}", query.from_user.id, (user.full_name if user else "Yönetici"), "rejected", f"Sağlık Raporu #{rep.id}")
            try: await query.message.delete()
            except Exception: pass
            await query.message.answer(get_text("medical_rejected", lang))
            if user: await render_clean_dashboard(query.message, user)
    await query.answer()

@router.callback_query(F.data.startswith("adm:class_att_sheet:"))
async def cb_class_attendance_sheet(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()

        if not students:
            await query.answer(get_text("no_students_in_class", lang), show_alert=True)
            return

        header_att = {
            "tr": f"📋 <b>{escape_md(class_name)} Sınıfı Toplu Devamsızlık Çizelgesi:</b>\n━━━━━━━━━━━━━━━━━━━━\n",
            "ru": f"📋 <b>Ведомость посещаемости класса {escape_md(class_name)}:</b>\n━━━━━━━━━━━━━━━━━━━━\n",
            "uz": f"📋 <b>{escape_md(class_name)} sinfining umumiy davomat qaydnomasi:</b>\n━━━━━━━━━━━━━━━━━━━━\n",
            "en": f"📋 <b>Class Attendance Sheet for {escape_md(class_name)}:</b>\n━━━━━━━━━━━━━━━━━━━━\n"
        }.get(lang, f"📋 <b>Attendance: {escape_md(class_name)}</b>\n")

        lines = [header_att]
        total_abs_class = 0
        lbl_day = {"tr": "gün", "ru": "дн.", "uz": "kun", "en": "days"}.get(lang, "days")
        lbl_crit = {"tr": " (Kritik)", "ru": " (Критично)", "uz": " (Xavfli)", "en": " (Critical)"}.get(lang, " (Critical)")
        lbl_exc = {"tr": "İzinli", "ru": "Уваж.", "uz": "Ruxsatli", "en": "Excused"}.get(lang, "Excused")

        for s in students:
            abs_cnt = (await session.execute(select(func.count(Attendance.id)).where(Attendance.student_id == s.id, Attendance.status == "absent"))).scalar() or 0
            exc_cnt = (await session.execute(select(func.count(Attendance.id)).where(Attendance.student_id == s.id, Attendance.status == "excused"))).scalar() or 0
            total_abs_class += abs_cnt
            warn_badge = f" ⚠️{lbl_crit}" if abs_cnt >= 7 else ""
            lines.append(f"• <b>{escape_md(s.full_name)}</b> (№{s.student_number}): <b>{abs_cnt} {lbl_day}</b>{warn_badge} <i>({lbl_exc}: {exc_cnt})</i>")

        avg_abs = round(total_abs_class / len(students), 1) if students else 0
        lbl_total_st = {"tr": "Sınıf Mevcudu", "ru": "Всего учеников", "uz": "Sinf mevcudi", "en": "Class Total"}.get(lang, "Total")
        lbl_avg_abs = {"tr": "Ortalama Devamsızlık", "ru": "Средний пропуск", "uz": "O'rtacha davomat", "en": "Average Absence"}.get(lang, "Avg")
        lines.append(f"\n📊 <b>{lbl_total_st}:</b> {len(students)} | <b>{lbl_avg_abs}:</b> {avg_abs} {lbl_day}")

        buttons = [get_nav_buttons(lang, back_callback=f"adm:show_class:{class_name}")]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

# ======================================================================
# 15. SABİT ALT MENÜ (REPLY KEYBOARD) VE DOĞRUDAN KOD YAZMA MOTORU
# ======================================================================

REPLY_BUTTON_ACTIONS = {
    "rk_digital_locker": "act_digital_locker",
    "rk_my_credentials": "act_my_credentials",
    "rk_logout": "act_logout",
    "rk_restart": "act_restart",
    "rk_main_menu": "act_main_menu",
    "rk_admin_dash": "act_main_menu",
    "btn_main_menu": "act_main_menu",
    "btn_proposals": "act_proposals",
    "btn_bot_block_monitor": "act_bot_block_monitor",
    "bc_home": "act_main_menu",
    "rk_lang": "act_lang",
    "btn_lang": "act_lang",
    "btn_login_prompt": "act_enter_code",
    "btn_req_access": "act_req_access",
    "rk_enter_code": "act_enter_code",
    "rk_req_access": "act_req_access",
    "rk_cat_staff": "act_cat_staff",
    "rk_cat_reports": "act_cat_reports",
    "rk_cat_requests": "act_cat_requests",
    "rk_cat_tools": "act_cat_tools",
    "rk_cat_tools_reports": "act_cat_tools_reports",
    "btn_bulletin_program": "act_parent_bulletin",
    "rk_cat_settings": "act_cat_settings",
    "rk_cockpit": "act_cockpit",
    "rk_classes": "act_classes",
    "rk_teachers": "act_teachers",
    "rk_requests": "act_requests",
    "rk_attendance": "act_attendance",
    "rk_grade": "act_grade",
    "rk_homework": "act_homework",
    "rk_behavior": "act_behavior",
    "btn_student_behavior_history": "act_behavior_hist",
    "btn_exam_schedule": "act_view_exams",
    "rk_parent_info": "act_parent_info",
    "rk_parent_settings": "act_parent_settings",
    "rk_report": "act_report",
    "rk_appointments": "act_appointments",
    "rk_schedule": "act_schedule",
    "rk_notices": "act_notices",
    "rk_cafeteria": "act_cafeteria",
    "rk_switch_student": "act_switch_student",
    "rk_upload_medical": "act_upload_medical"
}

def match_reply_button(text: str) -> str | None:
    if not text: return None
    clean_text = text.strip()
    for key, action in REPLY_BUTTON_ACTIONS.items():
        for lang_code in LOCALES:
            btn_txt = LOCALES[lang_code].get(key, "")
            if btn_txt and btn_txt == clean_text:
                return action

    clean_norm = clean_text.lower().replace("🏠", "").replace("⚡", "").replace("•", "").replace("📊", "").replace("🛠️", "").replace("👥", "").replace("🔔", "").replace("⚙️", "").replace("📁", "").replace("🔑", "").replace("🌐", "").replace("🚪", "").replace("🗳️", "").replace("🗳", "").replace("📅", "").replace("🧑‍🎓", "").replace("ℹ️", "").replace("⭐", "").replace("📋", "").replace("📝", "").replace("📢", "").replace("🤝", "").replace("🏥", "").strip()
    
    # Fuzzy & Normalized matches for all languages & historical button labels
    if any(m in clean_norm for m in ["kadro", "ogrenci", "öğrenci", "ученики", "учителя", "staff", "students", "xodim", "персонал"]):
        return "act_cat_staff"
    if any(m in clean_norm for m in ["rapor", "arac", "araç", "отчеты", "инструменты", "тулы", "reports", "tools", "hisobot"]):
        return "act_cat_tools_reports"
    if any(m in clean_norm for m in ["onay", "одобрений", "approval", "tasdiqlash"]):
        return "act_cat_requests"
    if any(m in clean_norm for m in ["ayar", "настройки", "settings", "sozlamalar"]):
        return "act_cat_settings"
    if any(m in clean_norm for m in ["dijital", "dosyam", "дело", "locker", "ishim"]):
        return "act_digital_locker"
    if any(m in clean_norm for m in ["sifre", "şifre", "пароли", "credentials", "passwords", "parol"]):
        return "act_my_credentials"
    if any(m in clean_norm for m in ["dil", "язык", "language", "til"]):
        return "act_lang"
    if any(m in clean_norm for m in ["cikis", "çıkış", "выйти", "logout", "chiqish"]):
        return "act_logout"
    if any(m in clean_norm for m in ["teklif", "oylama", "голосован", "предложен", "taklif", "ovoz", "proposal", "voting", "ballot", "oylar"]):
        return "act_proposals"
    if any(m in clean_norm for m in ["yoklama", "перекличка", "davomat", "attendance"]):
        return "act_attendance"
    if any(m in clean_norm for m in ["not gir", "оценки", "baho", "grade"]):
        return "act_grade"
    if any(m in clean_norm for m in ["davranış", "davranis", "поведение", "xulq", "behavior"]):
        return "act_behavior"
    if any(m in clean_norm for m in ["ödev", "odev", "задани", "vazifa", "homework"]):
        return "act_homework"
    if any(m in clean_norm for m in ["görüşme", "gorusme", "записи", "uchrashuv", "meeting"]):
        return "act_appointments"
    if any(m in clean_norm for m in ["karne", "табель", "tabel", "report"]):
        return "act_report"
    if any(m in clean_norm for m in ["tarix", "история", "geçmiş", "gecmis", "history"]):
        return "act_behavior_hist"
    if any(m in clean_norm for m in ["rapor", "справка", "медсправка", "ma'lumotnoma", "medical"]):
        return "act_upload_medical"
    if any(m in clean_norm for m in ["bülten", "bulten", "программ", "вестник", "график", "jadval", "bulletin", "plan"]):
        return "act_parent_bulletin"
    if any(m in clean_norm for m in ["öğrenci seç", "ogrenci sec", "смена", "almashtir", "tanlash", "switch"]):
        return "act_switch_student"
    if any(m in clean_norm for m in ["sınav", "sinav", "экзамен", "imtihon", "exam"]):
        return "act_view_exams"
    if any(m in clean_norm for m in ["bilgi", "инфо", "axborot", "info"]):
        return "act_parent_info"
    if any(m in clean_norm for m in ["giriş yap", "giris yap", "войти", "kirish", "log in"]):
        return "act_enter_code"
    if any(m in clean_norm for m in ["başvuru", "basvuru", "запрос", "talep", "ariza", "parol so", "request"]):
        return "act_req_access"
    if clean_norm in [
        "ana menu", "ana menü", "baş menü", "bas menu", "главное меню", "главное",
        "glavnoe menu", "glavnoye menyu", "asosiy menyu", "main menu", "menu", "меню"
    ]:
        return "act_main_menu"
    return None
@router.callback_query(F.data.startswith("rec_dismiss:"))
async def cb_rec_dismiss(query: CallbackQuery):
    try:
        rec_id = int(query.data.split(":")[1])
        async with AsyncSessionLocal() as session:
            rec = await session.get(DigitalRecord, rec_id)
            if rec:
                rec.is_archived_from_chat = True
                await session.commit()
    except Exception:
        pass

    user_lang = "tr"
    try:
        async with AsyncSessionLocal() as session:
            u = await session.get(User, query.from_user.id)
            if u and u.language:
                user_lang = u.language
    except Exception:
        pass

    toast_msg = {
        "tr": "📁 Evrak 'Dijital Dosyam' bölümüne kaldırıldı.",
        "ru": "📁 Документ перемещен в раздел 'Мое цифровое дело'.",
        "uz": "📁 Hujjat 'Raqamli hujjatlarim' bo'limiga olindi.",
        "en": "📁 Document moved to 'Digital Locker'."
    }.get(user_lang, "📁 Belge arşivlendi.")

    try:
        await query.answer(toast_msg, show_alert=False)
    except Exception:
        pass
    try:
        await query.message.delete()
    except Exception:
        pass


@router.callback_query(F.data == "cb_digital_locker")
async def cb_digital_locker(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        categories = ["receipt", "exam", "homework", "attendance", "appointment", "medical", "notice"]
        counts = {}
        for cat in categories:
            cnt = (await session.execute(
                select(func.count(DigitalRecord.id)).where(
                    DigitalRecord.user_telegram_id == user_id,
                    DigitalRecord.category == cat
                )
            )).scalar() or 0
            counts[cat] = cnt

        total_records = sum(counts.values())

        title = {
            "tr": "📁 <b>DİJİTAL DOSYAM & EVRAK ARŞİVİ</b>",
            "ru": "📁 <b>МОЕ ЦИФРОВОЕ ДЕЛО И АРХИВ</b>",
            "uz": "📁 <b>RAQAMLI HUJJATLARIM VA ARXIV</b>",
            "en": "📁 <b>DIGITAL LOCKER & ARCHIVE</b>"
        }.get(lang, "📁 <b>DİJİTAL DOSYAM</b>")

        desc = {
            "tr": (
                f"{title}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"Adınıza kayıtlı toplam <b>{total_records}</b> adet resmi evrak ve makbuz bulunmaktadır.\n\n"
                f"İncelemek istediğiniz evrak kategorisini seçiniz:"
            ),
            "ru": (
                f"{title}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"В вашем деле <b>{total_records}</b> официальных документов и квитанций.\n\n"
                f"Выберите категорию документов для просмотра:"
            ),
            "uz": (
                f"{title}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"Sizning nomingizda <b>{total_records}</b> ta rasmiy hujjat va kvitansiya mavjud.\n\n"
                f"Ko'rmoqchi bo'lgan hujjat toifasini tanlang:"
            ),
            "en": (
                f"{title}\n"
                f"━━━━━━━━━━━━━━━━━━━━\n"
                f"You have <b>{total_records}</b> official documents and receipts in your file.\n\n"
                f"Select a category to view:"
            )
        }.get(lang, f"{title}\nSeçim yapınız:")

        cat_labels = {
            "exam": {"tr": f"📑 Sınav Takvimleri ({counts['exam']})", "ru": f"📑 Расписание экзаменов ({counts['exam']})", "uz": f"📑 Imtihon jadvallari ({counts['exam']})", "en": f"📑 Exam Schedules ({counts['exam']})"},
            "homework": {"tr": f"📝 Ödev Makbuzları ({counts['homework']})", "ru": f"📝 Квитанции ДЗ ({counts['homework']})", "uz": f"📝 Vazifa kvitansiyalari ({counts['homework']})", "en": f"📝 Homework Receipts ({counts['homework']})"},
            "attendance": {"tr": f"⚠️ Devamsızlık Bildirimleri ({counts['attendance']})", "ru": f"⚠️ Пропуски ({counts['attendance']})", "uz": f"⚠️ Davomat xabarlari ({counts['attendance']})", "en": f"⚠️ Attendance Notices ({counts['attendance']})"},
            "appointment": {"tr": f"🗓️ Randevu Fişleri ({counts['appointment']})", "ru": f"🗓️ Талоны встреч ({counts['appointment']})", "uz": f"🗓️ Uchrashuv talonlari ({counts['appointment']})", "en": f"🗓️ Appointments ({counts['appointment']})"},
            "medical": {"tr": f"🏥 Sağlık & İzin Belgeleri ({counts['medical']})", "ru": f"🏥 Мед. справки ({counts['medical']})", "uz": f"🏥 Tibbiy ma'lumotnomalar ({counts['medical']})", "en": f"🏥 Medical Reports ({counts['medical']})"},
            "notice": {"tr": f"📢 Okul Duyuruları ({counts['notice']})", "ru": f"📢 Объявления школы ({counts['notice']})", "uz": f"📢 Maktab e'lonlari ({counts['notice']})", "en": f"📢 School Notices ({counts['notice']})"},
            "receipt": {"tr": f"📜 Kimlik & Giriş Makbuzları ({counts['receipt']})", "ru": f"📜 Учетные данные ({counts['receipt']})", "uz": f"📜 Kirish ma'lumotlari ({counts['receipt']})", "en": f"📜 Credentials & Receipts ({counts['receipt']})"}
        }

        buttons = [
            [InlineKeyboardButton(text=cat_labels["exam"][lang], callback_data="dl_cat:exam:1"), InlineKeyboardButton(text=cat_labels["homework"][lang], callback_data="dl_cat:homework:1")],
            [InlineKeyboardButton(text=cat_labels["attendance"][lang], callback_data="dl_cat:attendance:1"), InlineKeyboardButton(text=cat_labels["appointment"][lang], callback_data="dl_cat:appointment:1")],
            [InlineKeyboardButton(text=cat_labels["medical"][lang], callback_data="dl_cat:medical:1"), InlineKeyboardButton(text=cat_labels["notice"][lang], callback_data="dl_cat:notice:1")],
            [InlineKeyboardButton(text=cat_labels["receipt"][lang], callback_data="dl_cat:receipt:1")],
            get_nav_buttons(lang)
        ]

        await safe_edit_or_answer(event, desc, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")


@router.callback_query(F.data.startswith("dl_cat:"))
async def cb_dl_cat(query: CallbackQuery):
    parts = query.data.split(":")
    cat = parts[1]
    page = int(parts[2]) if len(parts) > 2 else 1

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        per_page = 5
        offset = (page - 1) * per_page

        total_cnt = (await session.execute(
            select(func.count(DigitalRecord.id)).where(
                DigitalRecord.user_telegram_id == query.from_user.id,
                DigitalRecord.category == cat
            )
        )).scalar() or 0

        total_pages = max(1, (total_cnt + per_page - 1) // per_page)

        records = (await session.execute(
            select(DigitalRecord).where(
                DigitalRecord.user_telegram_id == query.from_user.id,
                DigitalRecord.category == cat
            ).order_by(desc(DigitalRecord.created_at)).limit(per_page).offset(offset)
        )).scalars().all()

        cat_names = {
            "exam": {"tr": "Sınav Takvimleri", "ru": "Расписание экзаменов", "uz": "Imtihon jadvallari", "en": "Exam Schedules"},
            "homework": {"tr": "Ödev Makbuzları", "ru": "Квитанции ДЗ", "uz": "Vazifa kvitansiyalari", "en": "Homework Receipts"},
            "attendance": {"tr": "Devamsızlık Bildirimleri", "ru": "Пропуски", "uz": "Davomat xabarlari", "en": "Attendance Notices"},
            "appointment": {"tr": "Randevu Fişleri", "ru": "Талоны встреч", "uz": "Uchrashuv talonlari", "en": "Appointments"},
            "medical": {"tr": "Sağlık Belgeleri", "ru": "Мед. справки", "uz": "Tibbiy ma'lumotnomalar", "en": "Medical Reports"},
            "notice": {"tr": "Duyurular", "ru": "Объявления", "uz": "E'lonlar", "en": "Notices"},
            "receipt": {"tr": "Giriş & Makbuzlar", "ru": "Квитанции", "uz": "Kvitansiyalar", "en": "Receipts"}
        }
        c_title = cat_names.get(cat, {}).get(lang, cat)

        if not records:
            empty_txt = {
                "tr": f"📁 <b>{c_title}</b>\n━━━━━━━━━━━━━━━━━━━━\nBu kategoride henüz kayıtlı bir evrak bulunmamaktadır.",
                "ru": f"📁 <b>{c_title}</b>\n━━━━━━━━━━━━━━━━━━━━\nВ этой категории пока нет документов.",
                "uz": f"📁 <b>{c_title}</b>\n━━━━━━━━━━━━━━━━━━━━\nUshbu toifada hozircha hujjat mavjud emas.",
                "en": f"📁 <b>{c_title}</b>\n━━━━━━━━━━━━━━━━━━━━\nNo records found in this category yet."
            }.get(lang, "Kayıt yok.")
            buttons = [[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="cb_digital_locker")]]
            await safe_edit_or_answer(query, empty_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            return

        txt = (
            f"📁 <b>{c_title}</b> (Sayfa {page}/{total_pages})\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"Detayını incelemek istediğiniz evraka tıklayınız:\n"
        )

        buttons = []
        for r in records:
            d_str = r.created_at.strftime("%d.%m.%Y %H:%M")
            btn_txt = f"📄 {r.title[:24]} ({d_str[:10]})"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"dl_view:{r.id}:{cat}:{page}")])

        nav_row = []
        if page > 1:
            nav_row.append(InlineKeyboardButton(text="⬅️", callback_data=f"dl_cat:{cat}:{page-1}"))
        nav_row.append(InlineKeyboardButton(text=f"{page}/{total_pages}", callback_data="noop"))
        if page < total_pages:
            nav_row.append(InlineKeyboardButton(text="➡️", callback_data=f"dl_cat:{cat}:{page+1}"))
        
        if nav_row:
            buttons.append(nav_row)

        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="cb_digital_locker")])

        await safe_edit_or_answer(query, txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")


@router.callback_query(F.data.startswith("dl_view:"))
async def cb_dl_view(query: CallbackQuery):
    parts = query.data.split(":")
    rec_id = int(parts[1])
    cat = parts[2]
    page = int(parts[3])

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        rec = await session.get(DigitalRecord, rec_id)
        if not rec:
            await query.answer({"tr": "⚠️ Kayıt bulunamadı.", "ru": "⚠️ Запись не найдена.", "uz": "⚠️ Yozuv topilmadi.", "en": "⚠️ Record not found."}.get(lang, "⚠️ Record not found."), show_alert=True)
            return

        d_str = rec.created_at.strftime("%d.%m.%Y %H:%M")
        detail_html = (
            f"📁 <b>DİJİTAL EVRAK DETAYI</b>\n"
            f"━━━━━━━━━━━━━━━━━━━━\n"
            f"📄 <b>Evrak:</b> {escape_html(rec.title)}\n"
            f"📅 <b>Kayıt Tarihi:</b> <code>{d_str}</code>\n"
            f"🆔 <b>Kayıt No:</b> <code>#EVR-{rec.id:04d}</code>\n"
            f"━━━━━━━━━━━━━━━━━━━━\n\n"
            f"{rec.full_content}"
        )

        buttons = []
        if rec.file_id:
            dl_file_label = {
                "tr": "📥 Dosyayı Tekrar İndir",
                "ru": "📥 Скачать файл снова",
                "uz": "📥 Faylni qayta yuklash",
                "en": "📥 Download File Again"
            }.get(lang, "📥 Dosyayı İndir")
            buttons.append([InlineKeyboardButton(text=dl_file_label, callback_data=f"dl_file:{rec.id}")])

        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"dl_cat:{cat}:{page}")])

        await safe_edit_or_answer(query, detail_html, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")


@router.callback_query(F.data.startswith("dl_file:"))
async def cb_dl_file(query: CallbackQuery):
    rec_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        rec = await session.get(DigitalRecord, rec_id)
        if not rec or not rec.file_id:
            await query.answer({"tr": "⚠️ Dosya bulunamadı.", "ru": "⚠️ Файл не найден.", "uz": "⚠️ Fayl topilmadi.", "en": "⚠️ File not found."}.get(lang, "⚠️ File not found."), show_alert=True)
            return
        
        try:
            if rec.file_type == "photo":
                await query.message.bot.send_photo(chat_id=query.from_user.id, photo=rec.file_id, caption=f"📄 {rec.title}")
            else:
                await query.message.bot.send_document(chat_id=query.from_user.id, document=rec.file_id, caption=f"📄 {rec.title}")
            await query.answer("Dosya iletildi.")
        except Exception as e:
            await query.answer("Dosya gönderilemedi.", show_alert=True)

async def cb_show_my_credentials(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    chat_id = event.chat.id if isinstance(event, Message) else (event.message.chat.id if event.message else user_id)
    
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            role = "admin" if user_id in ADMIN_IDS else "guest"
            user = User(telegram_id=user_id, language="tr", role=role, full_name=event.from_user.full_name)
            session.add(user)
            await session.commit()
        lang = user.language or "tr"

        if user.role == "guest":
            guest_txt = {
                "tr": "⚠️ <b>Henüz Giriş Yapmadınız</b>\n──────────────\nŞifrelerinizi ve yetkilerinizi görebilmek için lütfen okul idaresinden aldığınız giriş kodunu giriniz veya <b>📩 Başvuru Yap</b> butonuna basarak kayıt talebinde bulununuz.",
                "ru": "⚠️ <b>Вы еще не вошли в систему</b>\n──────────────\nЧтобы увидеть ваши данные доступа, пожалуйста, введите код доступа или нажмите кнопку <b>📩 Запросить пароль</b> для подачи заявки.",
                "uz": "⚠️ <b>Siz hali tizimga kirmagansiz</b>\n──────────────\nKirish ma'lumotlaringizni ko'rish uchun maktab ma'muriyatidan berilgan kodni kiriting yoki <b>📩 Ariza berish</b> tugmasini bosing.",
                "en": "⚠️ <b>You are not logged in</b>\n──────────────\nPlease enter your access code or tap <b>📩 Request Access</b> to submit a registration request."
            }.get(lang, "Please log in.")
            await safe_edit_or_answer(event, guest_txt, parse_mode="HTML")
            if isinstance(event, CallbackQuery):
                try: await event.answer()
                except Exception: pass
            return

        title = {
            "tr": "🔑 <b>GİRİŞ VE KOD BİLGİLERİM</b>",
            "ru": "🔑 <b>МОИ ДАННЫЕ ДЛЯ ВХОДА И ПАРОЛИ</b>",
            "uz": "🔑 <b>KIRISH VA PAROL MA'LUMOTLARIM</b>",
            "en": "🔑 <b>MY ACCESS CREDENTIALS & CODES</b>"
        }.get(lang, "🔑 <b>MY ACCESS CREDENTIALS</b>")

        details_txt = ""
        if user.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user_id))).scalar_one_or_none()
            if not tch and user.full_name:
                tch = (await session.execute(select(Teacher).where(Teacher.full_name == user.full_name))).scalar_one_or_none()
            code = tch.auth_code if tch else "-"
            
            raw_subj = tch.subject if (tch and tch.subject and tch.subject.strip()) else ""
            if not raw_subj or raw_subj.lower() in ("genel", "nothing", "none", "-"):
                subj = {"tr": "Genel", "ru": "Общий", "uz": "Umumiy", "en": "General"}.get(lang, "General")
            else:
                subj = raw_subj

            raw_classes = tch.assigned_classes if (tch and tch.assigned_classes and tch.assigned_classes.strip()) else "ALL"
            if raw_classes.upper() == "ALL":
                cls_assigned = {"tr": "Tüm Sınıflar", "ru": "Все классы", "uz": "Barcha sinflar", "en": "All Classes"}.get(lang, "All Classes")
            else:
                cls_assigned = raw_classes
            
            details_txt = {
                "tr": (
                    f"📋 <b>Rolünüz:</b> 👨‍🏫 Öğretmen\n"
                    f"👤 <b>Adınız:</b> <b>{escape_html(user.full_name or 'Öğretmen')}</b>\n"
                    f"📚 <b>Branşınız:</b> <code>{escape_html(subj)}</code>\n"
                    f"🏫 <b>Atanan Sınıflar:</b> <code>{escape_html(cls_assigned)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Özel Öğretmen Şifreniz / Kodunuz:</b> <code>{code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Bu kod sisteme kayıtlı size özel erişim kodunuzdur.</i>"
                ),
                "ru": (
                    f"📋 <b>Роль:</b> 👨‍🏫 Учитель\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(user.full_name or 'Учитель')}</b>\n"
                    f"📚 <b>Предмет:</b> <code>{escape_html(subj)}</code>\n"
                    f"🏫 <b>Классы:</b> <code>{escape_html(cls_assigned)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Ваш код доступа (пароль):</b> <code>{code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Этот код привязан к вашей учетной записи учителя.</i>"
                ),
                "uz": (
                    f"📋 <b>Lavozimingiz:</b> 👨‍🏫 O'qituvchi\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(user.full_name or 'O`qituvchi')}</b>\n"
                    f"📚 <b>Faningiz:</b> <code>{escape_html(subj)}</code>\n"
                    f"🏫 <b>Sinflar:</b> <code>{escape_html(cls_assigned)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Maxsus o'qituvchi kodingiz (parol):</b> <code>{code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Ushbu kod sizning shaxsiy kirish kodingizdir.</i>"
                ),
                "en": (
                    f"📋 <b>Role:</b> 👨‍🏫 Teacher\n"
                    f"👤 <b>Name:</b> <b>{escape_html(user.full_name or 'Teacher')}</b>\n"
                    f"📚 <b>Subject:</b> <code>{escape_html(subj)}</code>\n"
                    f"🏫 <b>Assigned Classes:</b> <code>{escape_html(cls_assigned)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Teacher Access Code:</b> <code>{code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>This access code is linked to your teacher profile.</i>"
                )
            }.get(lang, "")

        elif user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user_id))).scalar_one_or_none()
            if not st and user.full_name:
                st = (await session.execute(select(Student).where(Student.full_name == user.full_name))).scalar_one_or_none()
            st_name = st.full_name if st else (user.full_name or "Öğrenci")
            cls_name = st.class_name if st else "-"
            st_no = st.student_number if st else "-"
            st_code = st.student_code if st else "-"
            pr_code = st.parent_code if st else "-"

            details_txt = {
                "tr": (
                    f"📋 <b>Rolünüz:</b> 🎓 Öğrenci\n"
                    f"👤 <b>Adınız:</b> <b>{escape_html(st_name)}</b>\n"
                    f"🏫 <b>Sınıfınız:</b> <code>{escape_html(cls_name)}</code>\n"
                    f"🔢 <b>Okul Numaranız:</b> <code>{escape_html(st_no)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Öğrenci Giriş Kodunuz:</b> <code>{st_code}</code>\n"
                    f"👨‍👩‍👧 <b>Veli Bağlantı Kodunuz:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Veliniz bu Veli Kodunu kullanarak bota kaydolabilir ve durumunuzu takip edebilir.</i>"
                ),
                "ru": (
                    f"📋 <b>Роль:</b> 🎓 Ученик\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(st_name)}</b>\n"
                    f"🏫 <b>Класс:</b> <code>{escape_html(cls_name)}</code>\n"
                    f"🔢 <b>Номер в школе:</b> <code>{escape_html(st_no)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Код ученика:</b> <code>{st_code}</code>\n"
                    f"👨‍👩‍👧 <b>Код для родителя:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Родитель может использовать данный код для подключения к боту.</i>"
                ),
                "uz": (
                    f"📋 <b>Lavozim:</b> 🎓 O'quvchi\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(st_name)}</b>\n"
                    f"🏫 <b>Sinfingiz:</b> <code>{escape_html(cls_name)}</code>\n"
                    f"🔢 <b>Maktab raqamingiz:</b> <code>{escape_html(st_no)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>O'quvchi kodingiz:</b> <code>{st_code}</code>\n"
                    f"👨‍👩‍👧 <b>Ota-ona ulanish kodi:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Ota-onangiz ushbu kod orqali botga ulanishi mumkin.</i>"
                ),
                "en": (
                    f"📋 <b>Role:</b> 🎓 Student\n"
                    f"👤 <b>Name:</b> <b>{escape_html(st_name)}</b>\n"
                    f"🏫 <b>Class:</b> <code>{escape_html(cls_name)}</code>\n"
                    f"🔢 <b>Roll Number:</b> <code>{escape_html(st_no)}</code>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🔑 <b>Student Code:</b> <code>{st_code}</code>\n"
                    f"👨‍👩‍👧 <b>Parent Code:</b> <code>{pr_code}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Your parent can use this Parent Code to link their account.</i>"
                )
            }.get(lang, "")

        elif user.role == "parent":
            st_list = []
            rels = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == user_id))).scalars().all()
            for r in rels:
                s = await session.get(Student, r.student_id)
                if s: st_list.append(s)
            
            lbl_pr_code = {"tr": "Veli Kodu", "ru": "Код родителя", "uz": "Ota-ona kodi", "en": "Parent Code"}.get(lang, "Parent Code")
            children_str = ""
            if st_list:
                for s in st_list:
                    children_str += f"\n• <b>{escape_html(s.full_name)}</b> (<code>{escape_html(s.class_name)}</code> - No: <code>{escape_html(s.student_number)}</code>)\n  └ 🔑 {lbl_pr_code}: <code>{s.parent_code}</code>"
            else:
                children_str = {
                    "tr": "\n• <i>Henüz bağlı öğrenci bulunmuyor.</i>",
                    "ru": "\n• <i>Нет привязанных учеников.</i>",
                    "uz": "\n• <i>Hozircha biriktirilgan o'quvchi yo'q.</i>",
                    "en": "\n• <i>No linked students yet.</i>"
                }.get(lang, "\n• <i>No linked students.</i>")

            details_txt = {
                "tr": (
                    f"📋 <b>Rolünüz:</b> 👨‍👩‍👧 Veli\n"
                    f"👤 <b>Adınız:</b> <b>{escape_html(user.full_name or 'Veli')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🧑‍🎓 <b>Bağlı Öğrenciler:</b>{children_str}\n"
                    "──────────────\n"
                    "ℹ️ <i>Yeni bir öğrenci bağlamak için okul idaresinden aldığınız Veli Kodunu giriniz.</i>"
                ),
                "ru": (
                    f"📋 <b>Роль:</b> 👨‍👩‍👧 Родитель\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(user.full_name or 'Родитель')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🧑‍🎓 <b>Привязанные ученики:</b>{children_str}\n"
                    "──────────────\n"
                    "ℹ️ <i>Используйте код родителя для привязки новых учеников.</i>"
                ),
                "uz": (
                    f"📋 <b>Lavozim:</b> 👨‍👩‍👧 Ota-ona\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(user.full_name or 'Ota-ona')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🧑‍🎓 <b>Biriktirilgan o'quvchilar:</b>{children_str}\n"
                    "──────────────\n"
                    "ℹ️ <i>Yangi o'quvchi qo'shish uchun maktab ma'muriyatiga murojaat qiling.</i>"
                ),
                "en": (
                    f"📋 <b>Role:</b> 👨‍👩‍👧 Parent\n"
                    f"👤 <b>Name:</b> <b>{escape_html(user.full_name or 'Parent')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    f"🧑‍🎓 <b>Linked Students:</b>{children_str}\n"
                    "──────────────\n"
                    "ℹ️ <i>Use the parent codes to link additional children.</i>"
                )
            }.get(lang, "")

        else: # admin
            admin_title = get_text("permanent_admin_title", lang) if (user_id in ADMIN_IDS or user.admin_type == "permanent") else get_text("temporary_admin_title", lang)
            details_txt = {
                "tr": (
                    f"📋 <b>Rolünüz:</b> ⚡ {admin_title}\n"
                    f"👤 <b>Adınız:</b> <b>{escape_html(user.full_name or 'Yönetici')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Tüm okul yönetim paneline ve yetkilerine doğrudan erişiminiz bulunmaktadır.</i>"
                ),
                "ru": (
                    f"📋 <b>Роль:</b> ⚡ {admin_title}\n"
                    f"👤 <b>ФИО:</b> <b>{escape_html(user.full_name or 'Администратор')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>У вас есть полный доступ ко всем функциям управления школой.</i>"
                ),
                "uz": (
                    f"📋 <b>Lavozim:</b> ⚡ {admin_title}\n"
                    f"👤 <b>F.I.O:</b> <b>{escape_html(user.full_name or 'Ma`mur')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>Barcha maktab boshqaruv vositalariga to'liq ruxsatingiz mavjud.</i>"
                ),
                "en": (
                    f"📋 <b>Role:</b> ⚡ {admin_title}\n"
                    f"👤 <b>Name:</b> <b>{escape_html(user.full_name or 'Admin')}</b>\n"
                    f"🆔 <b>Telegram ID:</b> <code>{user.telegram_id}</code>\n"
                    "──────────────\n"
                    "ℹ️ <i>You have full administrative access to all school management tools.</i>"
                )
            }.get(lang, "Admin access active.")

        back_cb = "adm:dashboard" if user.role == "admin" else ("tch:menu" if user.role == "teacher" else ("pr:menu" if user.role == "parent" else "st:menu"))
        creds_kb = InlineKeyboardMarkup(inline_keyboard=[
            get_nav_buttons(lang, back_callback=back_cb)
        ])
        full_card = f"{title}\n──────────────\n{details_txt}"
        await safe_edit_or_answer(event, full_card, reply_markup=creds_kb, parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        try: await event.answer()
        except Exception: pass

@router.message(any_state, F.text.func(lambda text: match_reply_button(text) is not None))
@router.callback_query(F.data == "act_lost_found_hub")
async def cb_lost_found_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        items = (await session.execute(
            select(LostItem).where(LostItem.status == "found").order_by(desc(LostItem.created_at)).limit(6)
        )).scalars().all()

        hdr = {
            "tr": "📦 <b>OKUL KAYIP EŞYA MASASI</b>\n──────────────\nOkul içinde bulunup idareye/öğretmenlere teslim edilen eşyalar aşağıda listelenmiştir. Size ait olan eşyayı seçerek talep edebilirsiniz:",
            "ru": "📦 <b>ШКОЛЬНОЕ БЮРО НАХОДОК</b>\n──────────────\nНайденные вещи в школе:",
            "uz": "📦 <b>MAKTAB TOPILMALAR BYUROSI</b>\n──────────────\nMaktabda topilgan buyumlar:",
            "en": "📦 <b>SCHOOL LOST & FOUND DESK</b>\n──────────────\nItems found on school grounds:"
        }.get(lang, "📦 <b>KAYIP EŞYA MASASI</b>")

        buttons = []
        if not items:
            hdr += "\n\n<i>" + {"tr": "Şu anda kayıtlı kayıp eşya bulunmamaktadır.", "ru": "В настоящее время найденных вещей нет.", "uz": "Hozirda topilgan buyumlar yo'q.", "en": "No lost items currently recorded."}.get(lang, "Kayıt yok.") + "</i>"
        else:
            for it in items:
                d_str = it.created_at.strftime("%d.%m")
                b_txt = f"🔍 {it.title[:20]} ({it.location_found}, {d_str})"
                buttons.append([InlineKeyboardButton(text=b_txt, callback_data=f"lost_view:{it.id}")])

        buttons.append(get_nav_buttons(lang, back_callback="nav:main_menu"))
        await safe_edit_or_answer(query, hdr, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        await query.answer()

@router.callback_query(F.data.startswith("lost_view:"))
async def cb_lost_item_view(query: CallbackQuery):
    it_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        item = await session.get(LostItem, it_id)
        if not item:
            await query.answer(get_text("record_not_found", lang), show_alert=True)
            return

        claim_btn_txt = {"tr": "🙋‍♂️ Bu Eşya Benim / Sahiplen", "ru": "🙋‍♂️ Это моя вещь", "uz": "🙋‍♂️ Bu mening buyumim", "en": "🙋‍♂️ This is Mine / Claim"}.get(lang, "🙋‍♂️ Sahiplen")
        d_str = item.created_at.strftime("%d.%m.%Y %H:%M")
        card_text = (
            f"📦 <b>KAYIP EŞYA DETAYI</b>\n"
            f"──────────────\n"
            f"• <b>Eşya:</b> {escape_html(item.title)}\n"
            f"• <b>Bulunduğu Yer:</b> {escape_html(item.location_found)}\n"
            f"• <b>Tarih:</b> <i>{d_str}</i>\n"
            f"• <b>Açıklama:</b> {escape_html(item.description or 'Belirtilmedi')}\n"
            f"• <b>Durum:</b> 🟢 TESLİM BEKLİYOR\n\n"
            f"<i>Bu eşya size aitse aşağıdaki butona tıklayarak idareye bildirim gönderebilirsiniz.</i>"
        )
        buttons = [
            [InlineKeyboardButton(text=claim_btn_txt, callback_data=f"lost_claim:{item.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="act_lost_found_hub")]
        ]
        await safe_edit_or_answer(query, card_text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        await query.answer()

@router.callback_query(F.data.startswith("lost_claim:"))
async def cb_lost_item_claim(query: CallbackQuery):
    it_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        item = await session.get(LostItem, it_id)
        if item:
            item.status = "claimed"
            item.claimed_by = query.from_user.id
            await session.commit()

        toast = {"tr": "✅ Talebiniz kaydedildi! Eşyayı teslim almak için okul idaresine başvurunuz.", "ru": "✅ Заявка принята! Обратитесь к администрации.", "uz": "✅ Qabul qilindi! Buyumni olish uchun ma'muriyatga murojaat qiling.", "en": "✅ Request recorded! Visit school office to collect."}.get(lang, "✅ Kaydedildi!")
        await query.answer(toast, show_alert=True)
        await cb_lost_found_hub(query)

async def global_reply_keyboard_router(message: Message, state: FSMContext):
    action = match_reply_button(message.text)
    user_id = message.from_user.id
    chat_id = message.chat.id

    try:
        await message.delete()
    except Exception:
        pass


    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            role = "admin" if user_id in ADMIN_IDS else "guest"
            user = User(telegram_id=user_id, language="tr", role=role, full_name=message.from_user.full_name)
            session.add(user)
            await session.commit()
        elif user_id in ADMIN_IDS and user.role != "admin":
            user.role = "admin"
            await session.commit()
        if message.from_user.full_name and user.full_name != message.from_user.full_name:
            user.full_name = message.from_user.full_name
            await session.commit()
        lang = user.language

    try:
        await message.delete()
    except Exception:
        pass

    if action:
        await state.clear()
        if action == "act_digital_locker":
            await cb_digital_locker(message, state)
            return
        elif action == "act_my_credentials":
            await cb_show_my_credentials(message, state)
            return
        elif action == "act_logout":
            await handle_bot_logout_cmd(message, state)
            return
        elif action == "act_restart":
            await cmd_start(message, state)
            return
        elif action == "act_cat_staff" and user.role == "admin":
            await cb_cat_staff(message, state)
            return
        elif action == "act_cat_reports" and user.role == "admin":
            await cb_cat_reports(message, state)
            return
        elif action == "act_cat_requests" and user.role == "admin":
            await cb_cat_requests(message, state)
            return
        elif action == "act_cat_tools" and user.role == "admin":
            await cb_cat_tools(message, state)
            return
        elif action == "act_cat_settings" and user.role == "admin":
            await cb_cat_settings(message, state)
            return
        elif action == "act_main_menu":
            await render_clean_dashboard(message, user)
            return
        elif action == "act_lang":
            await safe_edit_or_answer(message, get_text("lang_select", lang), reply_markup=get_language_inline_kb())
            return
        elif action == "act_enter_code":
            await message.answer(get_text("prompt_enter_code_direct", lang), parse_mode="HTML")
            await state.set_state(Form.waiting_auth_code)
            return
        elif action == "act_req_access":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_req_access")
            await cb_act_req_access(dummy_q, state)
            return
        elif action == "act_cockpit" and user.role == "admin":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="adm:cockpit")
            await cb_cockpit(dummy_q)
            return
        elif action == "act_classes" and user.role == "admin":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="adm:classes")
            await cb_classes_list(dummy_q, state)
            return
        elif action == "act_teachers" and user.role == "admin":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="adm:teachers")
            await cb_admin_teachers_list(dummy_q)
            return
        elif action == "act_requests" and user.role == "admin":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="adm:requests_list")
            await cb_admin_requests_list(dummy_q)
            return
        elif action == "act_attendance" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:classes")
            await cb_teacher_classes(dummy_q, state)
            return
        elif action == "act_grade" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:grade_classes")
            await cb_grade_classes(dummy_q, state)
            return
        elif action == "act_grade_behavior" and user.role == "teacher":
            lbl_gr = {"tr": "📝 Ders Notu Ver", "ru": "📝 Выставить оценку", "uz": "📝 Dars bahosi berish", "en": "📝 Enter Grade"}.get(user.language, "Grade")
            lbl_bh = {"tr": "⭐ Davranış Rozeti Ver", "ru": "⭐ Оценка поведения", "uz": "⭐ Xulq-atvor bahosi", "en": "Behavior"}.get(user.language, "Behavior")
            lbl_t = {"tr": "📝 <b>Öğrenci Not & Davranış Değerlendirme Masası</b>\nLütfen işlem yapmak istediğiniz alanı seçiniz:", "ru": "📝 <b>Оценки и поведение учеников</b>\nВыберите необходимое действие:", "uz": "📝 <b>O'quvchi bahosi va xulq-atvori</b>\nKerakli amalni tanlang:", "en": "📝 <b>Student Assessment & Behavior</b>\nSelect operation:"}.get(user.language, "Select operation:")
            btns = [
                [InlineKeyboardButton(text=lbl_gr, callback_data="tch:grade_classes"), InlineKeyboardButton(text=lbl_bh, callback_data="tch:behavior_classes")],
                get_nav_buttons(user.language)
            ]
            await message.answer(lbl_t, reply_markup=InlineKeyboardMarkup(inline_keyboard=btns), parse_mode="HTML")
            return
        elif action == "act_student_full_report":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="parent:report")
            await cb_parent_report(dummy_q)
            return
        elif action == "act_behavior" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:behavior_classes")
            await cb_teacher_behavior_classes(dummy_q)
            return
        elif action == "act_behavior_hist" and user.role in ("parent", "student"):
            st_id = user.current_child_id
            if user.role == "student":
                async with AsyncSessionLocal() as session:
                    st = (await session.execute(select(Student).where(Student.student_telegram_id == user_id))).scalar_one_or_none()
                    st_id = st.id if st else None
            if st_id:
                dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data=f"bh_hist:{st_id}")
                await cb_student_behavior_history(dummy_q)
            return
        elif action == "act_view_exams":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_exams")
            await cb_view_exams_list(dummy_q)
            return
        elif action == "act_homework":
            if user.role == "teacher":
                dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:hw_classes")
                await cb_hw_classes(dummy_q, state)
                return
            elif user.role in ("student", "parent"):
                dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_hw")
                await cb_view_homeworks_list(dummy_q)
                return
        elif action == "act_appointments" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:appointments")
            await cb_teacher_appointments_list(dummy_q)
            return
        elif action == "act_report" and user.role in ["parent", "student"]:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_report")
            await cb_view_report_card(dummy_q, state)
            return
        elif action == "act_appointments" and user.role == "parent":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_book_app")
            await cb_parent_book_app_init(dummy_q, state)
            return
        elif action == "act_schedule" and user.role in ["teacher", "parent", "student"]:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_sched")
            await cb_view_schedule(dummy_q)
            return
        elif action == "act_notices" and user.role in ["parent", "student"]:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_notices")
            await cb_view_notices(dummy_q)
            return
        elif action == "act_cafeteria" and user.role in ["parent", "student"]:
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_cafe")
            await cb_view_cafeteria(dummy_q)
            return
        elif action == "act_switch_student" and user.role == "parent":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="parent:switch_student")
            await cb_parent_switch_student(dummy_q, state)
            return
        elif action == "act_upload_medical" and user.role == "parent":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="upload_medical_init")
            await cb_upload_med_init(dummy_q, state)
            return
        
        elif action == "act_parent_info":
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_notices", lang), callback_data="act_view_notices"), InlineKeyboardButton(text=get_text("btn_view_schedule", lang), callback_data="act_view_sched")],
                [InlineKeyboardButton(text=get_text("btn_view_cafeteria", lang), callback_data="act_view_cafe"), InlineKeyboardButton(text=get_text("btn_exam_schedule", lang), callback_data="act_view_exams")],
                [InlineKeyboardButton(text=get_text("btn_lost_found", lang), callback_data="act_lost_found_hub")]
            ]
            if user.role == "parent":
                buttons.append([InlineKeyboardButton(text=get_text("rk_appointments", lang), callback_data="act_book_app"), InlineKeyboardButton(text=get_text("rk_upload_medical", lang), callback_data="upload_medical_init")])

            title = get_text("parent_info_title", lang)
            info_desc = {
                "tr": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nOkul duyurularını, haftalık ders programını, yemekhane menüsünü veya sınav takvimini görüntülemek için bir işlem seçiniz:",
                "ru": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nВыберите раздел для просмотра объявлений школы, расписания уроков, меню или экзаменов:",
                "uz": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nMaktab e'lonlari, dars jadvali, oshxona menyusi yoki imtihonlar jadvalini ko'rish uchun tanlang:",
                "en": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nSelect an option to view school announcements, timetable, cafeteria menu, or exams:"
            }.get(lang, f"<b>{title}</b>\nSelect an option:")

            sent_m = await message.bot.send_message(chat_id=message.chat.id, text=info_desc, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            if sent_m:
                LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
                ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)
            return
        elif action == "act_parent_settings":
            btn_brief = get_text("btn_briefing_on", lang) if user.evening_briefing else get_text("btn_briefing_off", lang)
            buttons = [
                [InlineKeyboardButton(text=btn_brief, callback_data="parent:toggle_briefing")],
                [InlineKeyboardButton(text="➕ " + {"tr": "Başka Çocuk Ekle", "ru": "Привязать ребенка", "uz": "Boshqa farzandni ulash", "en": "Link Child"}.get(lang, "Link Child"), callback_data="parent:add_child_code")],
                [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")]
            ]
            title = get_text("parent_settings_title", lang)
            p_desc = f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\n" + {"tr": "Ayarlarınızı yapılandırabilirsiniz:", "ru": "Настройки аккаунта:", "uz": "Sozlamalarni boshqarish:", "en": "Configure your settings:"}.get(lang, "Settings:")
            sent_m = await message.bot.send_message(chat_id=message.chat.id, text=p_desc, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            if sent_m:
                LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
                ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)
            return
        elif action == "act_proposals":
            if user.role == "teacher":
                dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:prop_hub")
                await cb_teacher_proposal_hub(dummy_q)
                return
            elif user.role == "admin":
                dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="adm:proposals_hub")
                await cb_admin_proposals_hub(dummy_q)
                return
        elif action == "act_bot_block_monitor" and user.role == "admin":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="adm:bot_block_monitor:0")
            await cb_admin_bot_block_monitor(dummy_q)
            return
        elif action == "act_cat_tools_reports" and user.role == "admin":
            await cb_cat_tools_reports(message, state)
            return
        elif action == "act_parent_bulletin" and user.role in ("parent", "student"):
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="parent:tab_bulletin:notices")
            await cb_parent_tab_bulletin(dummy_q)
            return

@router.message(any_state, F.text.func(lambda text: normalize_code(text).startswith(("VELI-", "OGR-", "HCA-", "ADM-")) or normalize_code(text) == ADMIN_CODE))
@router.message(default_state, F.text)
async def smart_text_auth_router(message: Message, state: FSMContext):
    user_id = message.from_user.id
    clean_code = normalize_code(message.text)
    is_code = clean_code.startswith(("VELI-", "OGR-", "HCA-", "ADM-")) or clean_code == ADMIN_CODE

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            role = "admin" if user_id in ADMIN_IDS else "guest"
            user = User(telegram_id=user_id, language="tr", role=role, full_name=message.from_user.full_name)
            session.add(user)
            await session.commit()
        if message.from_user.full_name and user.full_name != message.from_user.full_name:
            user.full_name = message.from_user.full_name
            await session.commit()
        if is_code or user.role == "guest":
            await process_auth_code_string(message.text, user_id, message, state)
        else:
            try: await message.delete()
            except Exception: pass

# ======================================================================
# 17. BOT BLOK TAKİP MASASI VE ZORLA SESLİ BİLDİRİM (FORCE AUDIO ALERT)
# ======================================================================

@router.callback_query(F.data.startswith("adm:bot_block_monitor:"))
async def cb_admin_bot_block_monitor(query: CallbackQuery):
    page = int(query.data.split(":")[2])
    per_page = 8
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):
            await query.answer()
            return

        total_users = (await session.execute(select(func.count(User.telegram_id)))).scalar() or 0
        blocked_users = (await session.execute(select(User).where((User.is_bot_blocked == True) | (User.is_blacklisted == True)).order_by(desc(User.blocked_bot_at)).offset(page * per_page).limit(per_page))).scalars().all()
        total_blocked = (await session.execute(select(func.count(User.telegram_id)).where((User.is_bot_blocked == True) | (User.is_blacklisted == True)))).scalar() or 0
        active_reach = max(0, total_users - total_blocked)
        total_pages = max(1, (total_blocked + per_page - 1) // per_page)

        bc_bb = {"tr": "🏠 Ana Menü ➔ 👥 Kadro ➔ 🚫 Bot İletişim Takip", "ru": "🏠 Главное меню ➔ 👥 Ученики ➔ 🚫 Контроль блокировок", "uz": "🏠 Asosiy menyu ➔ 👥 Kadro ➔ 🚫 Bot bloklarini nazorat", "en": "🏠 Main Menu ➔ 👥 Staff ➔ 🚫 Bot Block Tracking"}.get(lang, "🚫 Bot Block Tracking")
        title_bb = {"tr": "🚫 <b>BOT İLETİŞİM & ENGEL TAKİP MASASI</b>", "ru": "🚫 <b>КОНТРОЛЬ БЛОКИРОВОК И ДОСТАВКИ БОТА</b>", "uz": "🚫 <b>BOT BLOKLARINI NAZORAT QILISH MARKAZI</b>", "en": "🚫 <b>BOT DELIVERY & BLOCK TRACKING DESK</b>"}.get(lang, "🚫 <b>BOT TRACKING DESK</b>")

        stat_box = (
            f"<b>{bc_bb}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{title_bb}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📊 <b>İletişim Sağlığı ve İstatistikler:</b>\n"
            f"┌ 👥 <b>Toplam Kayıtlı:</b> {total_users} Kullanıcı\n"
            f"├ 🟢 <b>Aktif Ulaşılabilir:</b> {active_reach} Kullanıcı\n"
            f"└ 🚫 <b>Botu Engelleyen / Sessize Alan:</b> {total_blocked} Kullanıcı\n\n"
            f"⚠️ <i>Botu engelleyen veya sessize alan kullanıcılara acil durumlarda sesli zorla bildirim gönderebilir veya profillerini inceleyebilirsiniz.</i>\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        lines = [stat_box]
        buttons = []

        if not blocked_users:
            no_bl = {"tr": "✅ <i>Harika! Şu anda botu engelleyen hiçbir kullanıcı bulunmuyor.</i>", "ru": "✅ <i>Отлично! Нет пользователей, заблокировавших бота.</i>", "uz": "✅ <i>Ajoyib! Hozirda botni bloklagan foydalanuvchilar yo'q.</i>", "en": "✅ <i>Great! No users have blocked the bot currently.</i>"}.get(lang, "✅ No blocked users.")
            lines.append(f"\n{no_bl}")
        else:
            list_hdr = {"tr": f"📋 <b>Engelli Kullanıcılar Listesi (Sayfa {page + 1}/{total_pages}):</b>", "ru": f"📋 <b>Список заблокировавших (Стр. {page + 1}/{total_pages}):</b>", "uz": f"📋 <b>Bloklaganlar ro'yxati ({page + 1}/{total_pages}):</b>", "en": f"📋 <b>Blocked Users List (Page {page + 1}/{total_pages}):</b>"}.get(lang, "📋 Blocked Users:")
            lines.append(f"\n{list_hdr}\n")
            for u in blocked_users:
                u_name = u.full_name or "Kullanıcı"
                t_str = u.blocked_bot_at.strftime('%d.%m %H:%M') if u.blocked_bot_at else "Bilinmiyor"
                lines.append(f"• 🚫 <b>{escape_html(u_name)}</b> [{get_role_label(u.role, lang)}] - <code>{u.telegram_id}</code> (🕒 {t_str})")
                buttons.append([InlineKeyboardButton(text=f"👤 {u_name} ({get_role_label(u.role, lang)})", callback_data=f"adm:user_card:{u.telegram_id}:{page}")])

        nav_row = []
        if page > 0: nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:bot_block_monitor:{page - 1}"))
        if (page + 1) * per_page < total_blocked: nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:bot_block_monitor:{page + 1}"))
        if nav_row: buttons.append(nav_row)

        btn_force = {"tr": "🚨 Zorla Sesli Acil Bildirim Gönder", "ru": "🚨 Отправить звуковое оповещение", "uz": "🚨 Ovozli shoshilinch ogohlantirish", "en": "🚨 Send Force Audio Alert"}.get(lang, "🚨 Force Audio Alert")
        buttons.append([InlineKeyboardButton(text=btn_force, callback_data="adm:force_audio_alert_init")])
        buttons.append([
            InlineKeyboardButton(text="🔍 " + {"tr": "Sessiz Sağlık Taraması (Health-Check)", "ru": "Тихая проверка связи", "uz": "Tinch tekshiruv", "en": "Silent Health-Check"}.get(lang, "Health Check"), callback_data="adm:scan_bot_blocks")
        ])
        buttons.append([
            InlineKeyboardButton(text=get_text("btn_refresh_data", lang), callback_data=f"adm:bot_block_monitor:{page}"),
            InlineKeyboardButton(text=get_text("btn_users_hub", lang), callback_data="adm:users_hub:0")
        ])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:force_audio_alert_init")
async def cb_force_audio_alert_init(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):
            await query.answer(get_text("unauthorized_action", lang), show_alert=True)
            return

        hdr = {
            "tr": "🚨 <b>ACİL DURUM BİLDİRİM MASASI (YÜKSEK ÖNCELİK)</b>\n━━━━━━━━━━━━━━━━━━━━\nBu işlem botu sessize alan veya bildirimleri kapatan kullanıcılara Telegram'ın <code>disable_notification=False</code> ve mesaj sabitleme (pin) parametreleri ile <b>yüksek öncelikli sesli ve görsel alarm</b> gönderir.\n\nLütfen hedef kitleyi seçiniz:",
            "ru": "🚨 <b>ЭКСТРЕННОЕ ОПОВЕЩЕНИЕ (ВЫСОКИЙ ПРИОРИТЕТ)</b>\n━━━━━━━━━━━━━━━━━━━━\nОтправляет экстренное оповещение с принудительным звуковым сигналом и закреплением сообщения поверх всех чатов.\n\nВыберите получателей:",
            "uz": "🚨 <b>SHOSHILINCH XABAR (YUQORI DARAJALI)</b>\n━━━━━━━━━━━━━━━━━━━━\nFoydalanuvchilarga ovozli bildirishnoma va qadab qo'yish (pin) orqali yuqori darajadagi shoshilinch xabar yuboradi.\n\nAuditoriyani tanlang:",
            "en": "🚨 <b>EMERGENCY ALERT (HIGH PRIORITY)</b>\n━━━━━━━━━━━━━━━━━━━━\nDispatches high-priority alert with audio notification and chat pin signal.\n\nSelect target audience:"
        }.get(lang, "Select target audience:")

        buttons = [
            [InlineKeyboardButton(text="👨‍👩‍👧‍👦 " + {"tr": "Tüm Velilere Acil Bildirim", "ru": "Всем родителям (Экстренно)", "uz": "Barcha ota-onalarga (Shoshilinch)", "en": "All Parents (Urgent)"}.get(lang, "All Parents"), callback_data="adm:force_aud:parents")],
            [InlineKeyboardButton(text="👨‍🏫 " + {"tr": "Tüm Öğretmenlere Acil Bildirim", "ru": "Всем учителям (Экстренно)", "uz": "Barcha o'qituvchilarga (Shoshilinch)", "en": "All Teachers (Urgent)"}.get(lang, "All Teachers"), callback_data="adm:force_aud:teachers")],
            [InlineKeyboardButton(text="👥 " + {"tr": "Tüm Okula Acil Bildirim (Genel)", "ru": "Всей школе (Экстренно)", "uz": "Butun maktabga (Shoshilinch)", "en": "All School (Urgent)"}.get(lang, "All School"), callback_data="adm:force_aud:all")],
            [InlineKeyboardButton(text="❌ " + get_text("btn_cancel_action", lang), callback_data="adm:cat_tools_reports")]
        ]
        await safe_edit_or_answer(query, hdr, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:force_audio_alert_cancel")
async def cb_force_audio_alert_cancel(query: CallbackQuery, state: FSMContext):
    await state.clear()
    FORCE_ALERT_CACHE.pop(query.from_user.id, None)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    cancel_txt = {
        "tr": "❌ Acil bildirim gönderimi iptal edildi.",
        "ru": "❌ Отправка экстренного оповещения отменена.",
        "uz": "❌ Shoshilinch xabar yuborish bekor qilindi.",
        "en": "❌ Emergency alert cancelled."
    }.get(lang, "Cancelled.")
    back_btn = InlineKeyboardMarkup(inline_keyboard=[[
        InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:cat_tools_reports")
    ]])
    await safe_edit_or_answer(query, cancel_txt, reply_markup=back_btn, parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:force_aud:"))
async def cb_send_force_alert_pick(query: CallbackQuery, state: FSMContext):
    target_aud = query.data.split(":")[2]
    FORCE_ALERT_CACHE[query.from_user.id] = {"target": target_aud}
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    prompt_t = {
        "tr": f"🚨 <b>Acil Durum Bildirim Metni ({target_aud.upper()})</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen kullanıcılara yüksek öncelikle ve sesli iletilecek acil durum metnini yazınız (veya fotoğrafla açıklama gönderiniz):",
        "ru": f"🚨 <b>Текст экстренного оповещения ({target_aud.upper()})</b>\n━━━━━━━━━━━━━━━━━━━━\nВведите текст экстренного оповещения (или отправьте фото с текстом):",
        "uz": f"🚨 <b>Shoshilinch xabar matni ({target_aud.upper()})</b>\n━━━━━━━━━━━━━━━━━━━━\nFoydalanuvchilarga ovozli va yuqori darajada yuboriladigan xabarni kiriting (yoki rasmli izoh yuboring):",
        "en": f"🚨 <b>Emergency Alert Text ({target_aud.upper()})</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter the urgent alert text (or send a photo with caption):"
    }.get(lang, "Enter alert text:")

    cancel_btn = InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="❌ " + get_text("btn_cancel_action", lang), callback_data="adm:force_audio_alert_cancel")]
    ])
    await safe_edit_or_answer(query, prompt_t, reply_markup=cancel_btn, parse_mode="HTML")
    await state.set_state(Form.waiting_force_alert_text)
    await query.answer()

@router.message(Form.waiting_force_alert_text)
async def process_force_audio_alert_text(message: Message, state: FSMContext):
    raw_text = message.caption or message.text or ""
    photo_id = message.photo[-1].file_id if message.photo else None
    user_id = message.from_user.id

    # 1. İptal Kontrolü: Kullanıcı metinle veya klavyeden iptal ederse
    if is_universal_cancel_text(raw_text) or (raw_text and match_reply_button(raw_text)):
        FORCE_ALERT_CACHE.pop(user_id, None)
        await state.clear()
        async with AsyncSessionLocal() as session:
            user = await session.get(User, user_id)
            lang = user.language if user else "tr"
        c_msg = {
            "tr": "❌ Acil durum bildirimi iptal edildi.",
            "ru": "❌ Экстренное оповещение отменено.",
            "uz": "❌ Shoshilinch xabar bekor qilindi.",
            "en": "❌ Emergency alert cancelled."
        }.get(lang, "Cancelled.")
        role_kb = get_role_reply_kb("admin", lang)
        await message.bot.send_message(chat_id=message.chat.id, text=c_msg, reply_markup=role_kb, parse_mode="HTML")
        return

    cache = FORCE_ALERT_CACHE.pop(user_id, {"target": "all"})
    await state.clear()

    if not raw_text and not photo_id:
        return

    target_aud = cache.get("target", "all")

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        target_ids = []
        if target_aud == "parents":
            u_parents = (await session.execute(select(User.telegram_id).where(User.role.in_(["parent", "Parent"])))).scalars().all()
            ps_parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.parent_telegram_id.isnot(None)))).scalars().all()
            target_ids = list(set(list(u_parents) + list(ps_parents)))
        elif target_aud == "teachers":
            u_teachers = (await session.execute(select(User.telegram_id).where(User.role.in_(["teacher", "Teacher"])))).scalars().all()
            t_teachers = (await session.execute(select(Teacher.telegram_id).where(Teacher.telegram_id.isnot(None)))).scalars().all()
            target_ids = list(set(list(u_teachers) + list(t_teachers)))
        else: # "all"
            all_u = (await session.execute(select(User.telegram_id))).scalars().all()
            ps_parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.parent_telegram_id.isnot(None)))).scalars().all()
            t_teachers = (await session.execute(select(Teacher.telegram_id).where(Teacher.telegram_id.isnot(None)))).scalars().all()
            st_users = (await session.execute(select(Student.student_telegram_id).where(Student.student_telegram_id.isnot(None)))).scalars().all()
            target_ids = list(set(list(all_u) + list(ps_parents) + list(t_teachers) + list(st_users)))

        # Gönderen yöneticiyi kesinlikle hariç tut (kendisine 'Okudum/Onayladım' çıkmaz!)
        target_ids = [tid for tid in set(target_ids) if tid and tid != user_id]

        if not target_ids:
            empty_txt = {
                "tr": f"⚠️ Seçilen hedef kitlede (<b>{target_aud.upper()}</b>) sizden başka kayıtlı aktif kullanıcı bulunamadı.\nLütfen veli veya öğretmenlerin sisteme giriş yaptığından emin olunuz.",
                "ru": f"⚠️ В выбранной аудитории (<b>{target_aud.upper()}</b>) кроме вас нет зарегистрированных пользователей.\nУбедитесь, что родители или учителя вошли в систему.",
                "uz": f"⚠️ Tanlangan auditoriyada (<b>{target_aud.upper()}</b>) sizdan boshqa foydalanuvchi topilmadi.\nIltimos, ota-onalar yoki o'qituvchilar tizimga kirganligini tekshiring.",
                "en": f"⚠️ In the selected target audience (<b>{target_aud.upper()}</b>), no registered users were found other than you.\nPlease ensure parents or teachers are registered."
            }.get(lang, "No recipients found.")
            role_kb = get_role_reply_kb("admin", lang)
            sent_m = await message.bot.send_message(chat_id=message.chat.id, text=empty_txt, reply_markup=role_kb, parse_mode="HTML")
            if sent_m:
                LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
                ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)
            return

        sent_count = 0
        for tid in target_ids:
            u_db = await session.get(User, tid)
            u_l = u_db.language if u_db else "tr"
            ack_btn_txt = {"tr": "✅ Okudum / Onayladım", "ru": "✅ Ознакомлен / Принято", "uz": "✅ O'qidim / Tasdiqlayman", "en": "✅ Read & Acknowledged"}.get(u_l, "Acknowledge")
            ack_kb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=ack_btn_txt, callback_data=f"ack_fa:{user_id}")]])
            
            hdr_fa = {
                "tr": "🚨🔊 <b>OKUL YÖNETİMİ ACİL BİLDİRİMİ</b>",
                "ru": "🚨🔊 <b>ЭКСТРЕННОЕ ОПОВЕЩЕНИЕ ШКОЛЫ</b>",
                "uz": "🚨🔊 <b>MAKTAB MA'MURIYATI SHOSHILINCH XABARI</b>",
                "en": "🚨🔊 <b>SCHOOL ADMINISTRATION URGENT ALERT</b>"
            }.get(u_l, "🚨🔊 <b>URGENT ALERT</b>")

            ftr_fa = {
                "tr": "⚠️ <i>Bu bildirim okul idaresi tarafından yüksek öncelikle ve sesli olarak iletilmiştir.</i>",
                "ru": "⚠️ <i>Это экстренное сообщение высокой важности со звуковым оповещением.</i>",
                "uz": "⚠️ <i>Ushbu xabar maktab ma'muriyati tomonidan yuqori darajadagi ovozli bildirishnoma bilan yuborildi.</i>",
                "en": "⚠️ <i>This urgent notification was dispatched with high-priority audio alert.</i>"
            }.get(u_l, "⚠️ <i>Urgent alert.</i>")

            card_msg = (
                f"{hdr_fa}\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"{escape_html(raw_text)}\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"{ftr_fa}"
            )
            try:
                sent_msg = None
                if photo_id:
                    sent_msg = await message.bot.send_photo(chat_id=tid, photo=photo_id, caption=card_msg, reply_markup=ack_kb, parse_mode="HTML", disable_notification=False)
                else:
                    sent_msg = await message.bot.send_message(chat_id=tid, text=card_msg, reply_markup=ack_kb, parse_mode="HTML", disable_notification=False)

                # Ekranın En Üstüne Sabitle (Pin Sinyali ile ek sesli uyarı)
                if sent_msg:
                    try:
                        await message.bot.pin_chat_message(chat_id=tid, message_id=sent_msg.message_id, disable_notification=False)
                    except Exception:
                        pass

                sent_count += 1
                await asyncio.sleep(0.04)
            except Exception:
                pass

        done_txt = {
            "tr": f"🚨 Acil durum bildirimi başarıyla <b>{sent_count}</b> kullanıcıya ulaştırıldı!",
            "ru": f"🚨 Экстренное оповещение успешно доставлено <b>{sent_count}</b> пользователям!",
            "uz": f"🚨 Shoshilinch xabar muvaffaqiyatli <b>{sent_count}</b> ta foydalanuvchiga yetkazildi!",
            "en": f"🚨 Emergency alert successfully delivered to <b>{sent_count}</b> users!"
        }.get(lang, "Delivered.")
        role_kb = get_role_reply_kb("admin", lang)
        sent_m = await message.bot.send_message(chat_id=message.chat.id, text=done_txt, reply_markup=role_kb, parse_mode="HTML")
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data.startswith("ack_fa:"))
async def cb_ack_force_alert(query: CallbackQuery):
    tst = {"tr": "✅ Acil durum onayınız kaydedildi.", "ru": "✅ Ваше подтверждение принято.", "uz": "✅ Tasdig'ingiz qabul qilindi.", "en": "✅ Confirmed."}.get("tr", "Confirmed.")
    await query.answer(tst, show_alert=True)
    if query.message:
        try:
            await query.bot.unpin_chat_message(chat_id=query.message.chat.id, message_id=query.message.message_id)
        except Exception:
            pass
        try:
            await query.message.edit_reply_markup(reply_markup=None)
        except Exception:
            pass

# ======================================================================
# 18. ÖĞRETMEN & YÖNETİCİ OYLAMA, ŞİKAYET VE İSTİŞARE SİSTEMİ
# ======================================================================

@router.callback_query(F.data.startswith("prop:view_att:"))
async def cb_view_proposal_attachment(query: CallbackQuery):
    prop_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        prop = await session.get(Proposal, prop_id)
        if not prop or not prop.attachment_file_id:
            await query.answer({"tr": "⚠️ Bu teklife ait ekli belge bulunamadı.", "ru": "⚠️ К этому предложению нет прикрепленного документа.", "uz": "⚠️ Ushbu taklifga biriktirilgan hujjat topilmadi.", "en": "⚠️ No attached document found for this proposal."}.get(lang, "⚠️ No attached document found for this proposal."), show_alert=True)
            return

        caption = f"📎 <b>{escape_html(prop.title)}</b> — Ekli Belge / Tutanak"
        try:
            if prop.attachment_type == "photo":
                await query.bot.send_photo(chat_id=query.message.chat.id, photo=prop.attachment_file_id, caption=caption, parse_mode="HTML")
            else:
                await query.bot.send_document(chat_id=query.message.chat.id, document=prop.attachment_file_id, caption=caption, parse_mode="HTML")
            await query.answer()
        except Exception as e:
            await query.answer(f"⚠️ Belge gönderilemedi: {str(e)[:50]}", show_alert=True)

@router.callback_query(F.data.startswith("prop:open:"))
async def cb_open_proposal_for_vote(query: CallbackQuery):
    prop_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        prop = await session.get(Proposal, prop_id)
        if not prop or prop.status != "active":
            await query.answer("⚠️ Bu teklif artık aktif değil.", show_alert=True)
            return

        agree_cnt = (await session.execute(select(func.count(ProposalVote.id)).where(ProposalVote.proposal_id == prop.id, ProposalVote.vote_choice == "agree"))).scalar() or 0
        disagree_cnt = (await session.execute(select(func.count(ProposalVote.id)).where(ProposalVote.proposal_id == prop.id, ProposalVote.vote_choice == "disagree"))).scalar() or 0

        deadline_line = f"⏳ <b>Son Oy:</b> <code>{prop.deadline_at.strftime('%d.%m.%Y %H:%M')}</code>\n" if prop.deadline_at else ""
        anon_line = "🕵️ <b>Gizli Oylama (Anonim)</b>\n" if prop.is_anonymous else "👁️ <b>Açık Oylama</b>\n"

        p_card = (
            f"🗳️ <b>{escape_html(prop.title)}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{anon_line}"
            f"{deadline_line}"
            f"📝 <b>Açıklama:</b>\n{escape_html(prop.content)}\n\n"
            f"📊 <b>Mevcut Durum:</b> ✅ {agree_cnt} | ❌ {disagree_cnt}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "<i>Lütfen aşağıdaki butonlarla oyunuzu kullanınız:</i>"
        )
        b_agr = {"tr": "✅ Katılıyorum", "ru": "✅ Согласен", "uz": "✅ Qo'shilaman", "en": "✅ Agree"}.get(lang, "✅ Agree")
        b_dis = {"tr": "❌ Katılmıyorum", "ru": "❌ Не согласен", "uz": "❌ Qo'shilmayman", "en": "❌ Disagree"}.get(lang, "❌ Disagree")
        b_rsn = {"tr": "📝 Gerekçe / Fikir Ekle", "ru": "📝 Добавить комментарий", "uz": "📝 Izoh / Fikr qo'shish", "en": "📝 Add Note"}.get(lang, "📝 Add Note")
        b_res = {"tr": "📊 Oylama Sonuçları", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Results"}.get(lang, "📊 Results")
        back_cb = "adm:proposals_hub" if is_admin_user(user, query.from_user.id) else "tch:active_props:0"
        vote_buttons = [
            [
                InlineKeyboardButton(text=b_agr, callback_data=f"prop:vote:{prop.id}:agree"),
                InlineKeyboardButton(text=b_dis, callback_data=f"prop:vote:{prop.id}:disagree")
            ],
            [InlineKeyboardButton(text=b_rsn, callback_data=f"prop:reason:{prop.id}")],
            [InlineKeyboardButton(text=b_res, callback_data=f"prop:results:{prop.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=back_cb)]
        ]
        if prop.attachment_file_id:
            b_att = {"tr": "📎 Ekli Belgeyi Gör", "ru": "📎 Прикрепленный файл", "uz": "📎 Biriktirilgan fayl", "en": "📎 View Attachment"}.get(lang, "📎 View Attachment")
            vote_buttons.append([InlineKeyboardButton(text=b_att, callback_data=f"prop:view_att:{prop.id}")])
        await safe_edit_or_answer(query, p_card, reply_markup=InlineKeyboardMarkup(inline_keyboard=vote_buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "tch:prop_hub")
async def cb_teacher_proposal_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        bc_tp = {"tr": "🏠 Ana Menü ➔ 🗳️ Teklif & Şikayet Masası", "ru": "🏠 Главное меню ➔ 🗳️ Голосования и предложения", "uz": "🏠 Asosiy menyu ➔ 🗳️ Taklif va shikoyatlar", "en": "🏠 Main Menu ➔ 🗳️ Voting & Proposals"}.get(lang, "🗳️ Voting & Proposals")
        title_tp = {"tr": "🗳️ <b>ÖĞRETMEN TEKLİF, ŞİKAYET VE OYLAMA MASASI</b>", "ru": "🗳️ <b>ПРЕДЛОЖЕНИЯ, ЖАЛОБЫ И ГОЛОСОВАНИЯ</b>", "uz": "🗳️ <b>O'QITUVCHILAR TAKLIF VA OVOZ BERISH MARKAZI</b>", "en": "🗳️ <b>TEACHER PROPOSALS & VOTING DESK</b>"}.get(lang, "🗳️ <b>PROPOSALS & VOTING</b>")
        desc_tp = {"tr": "📌 Bir öğrenci hakkında idareye şikayet/disiplin oylaması başlatabilir, kurumsal bir eğitim teklifi sunabilir veya aktif oylamalarda oy kullanabilirsiniz:", "ru": "📌 Создайте обращение по поведению ученика, внесите предложение или проголосуйте в текущих обсуждениях:", "uz": "📌 O'quvchi xulqi bo'yicha shikoyat yoki taklif kiritishingiz va faol ovoz berishlarda qatnashishingiz mumkin:", "en": "📌 Submit a student disciplinary complaint, create a school proposal, or cast your vote on active ballots:"}.get(lang, "Select option:")

        active_cnt = (await session.execute(select(func.count(Proposal.id)).where(Proposal.status == 'active'))).scalar() or 0

        active_cnt_str = {
            "tr": f"🗳️ <b>Aktif Oylama Sayısı:</b> {active_cnt} Adet",
            "ru": f"🗳️ <b>Активных голосований:</b> {active_cnt}",
            "uz": f"🗳️ <b>Faol ovoz berishlar soni:</b> {active_cnt} ta",
            "en": f"🗳️ <b>Active Ballots:</b> {active_cnt}"
        }.get(lang, f"🗳️ <b>Active Ballots:</b> {active_cnt}")

        text_content = (
            f"<b>{bc_tp}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{title_tp}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{desc_tp}\n\n"
            f"{active_cnt_str}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = [
            [InlineKeyboardButton(text="⚠️ " + {"tr": "Öğrenci Disiplin Şikayeti", "ru": "Дисциплинарная жалоба", "uz": "O'quvchi intizom shikoyati", "en": "Disciplinary Complaint"}.get(lang, "Disciplinary Complaint"), callback_data="tch:prop_new:student_complaint")],
            [InlineKeyboardButton(text="💡 " + {"tr": "Kurumsal Teklif & Öneri", "ru": "Школьное предложение", "uz": "Maktab bo'yicha taklif", "en": "School Proposal"}.get(lang, "School Proposal"), callback_data="tch:prop_new:teacher_proposal")],
            [InlineKeyboardButton(text="🗳️ " + {"tr": f"Aktif Oylamalar ({active_cnt})", "ru": f"Голосования ({active_cnt})", "uz": f"Ovoz berishlar ({active_cnt})", "en": f"Active Ballots ({active_cnt})"}.get(lang, f"Active Ballots ({active_cnt})"), callback_data="tch:active_props:0")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:menu")]
        ]
        await safe_edit_or_answer(query, text_content, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:prop_new:"))
async def cb_teacher_new_proposal_init(query: CallbackQuery, state: FSMContext):
    prop_type = query.data.split(":")[2]
    await state.clear()
    PROP_CACHE[query.from_user.id] = {"type": prop_type}

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        if prop_type == "student_complaint":
            classes = await get_all_school_classes(session)
            if not classes:
                no_cls_txt = {
                    "tr": "⚠️ <b>Öğrenci Şikayeti / Disiplin</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ Okulda kayıtlı sınıf bulunamadı. Şikayet oluşturabilmek için önce sisteme sınıflar ve öğrenciler eklenmelidir.",
                    "ru": "⚠️ <b>Жалоба на ученика / Дисциплина</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ В системе пока нет зарегистрированных классов. Сначала необходимо добавить классы и учеников.",
                    "uz": "⚠️ <b>O'quvchi shikoyati / Intizom</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ Maktabda ro'yxatga olingan sinflar topilmadi. Shikoyat kiritish uchun avval sinflar va o'quvchilar qo'shilishi kerak.",
                    "en": "⚠️ <b>Student Disciplinary Complaint</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ No classes registered yet in the system. Classes and students must be added first."
                }.get(lang, "No classes registered yet.")
                no_cls_kb = InlineKeyboardMarkup(inline_keyboard=[
                    [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_hub")]
                ])
                await safe_edit_or_answer(query, no_cls_txt, reply_markup=no_cls_kb, parse_mode="HTML")
                await query.answer()
                return

            buttons = []
            row = []
            for c in classes:
                row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"tch:prop_cls:{c}"))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row: buttons.append(row)
            buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_hub")])
            prompt_cls = {
                "tr": "⚠️ <b>Öğrenci Şikayeti / Disiplin</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen şikayet konusu olan öğrencinin sınıfını seçiniz:",
                "ru": "⚠️ <b>Жалоба на ученика / Дисциплина</b>\n━━━━━━━━━━━━━━━━━━━━\nВыберите класс ученика:",
                "uz": "⚠️ <b>O'quvchi shikoyati / Intizom</b>\n━━━━━━━━━━━━━━━━━━━━\nShikoyat qilinadigan o'quvchi sinfini tanlang:",
                "en": "⚠️ <b>Student Disciplinary Complaint</b>\n━━━━━━━━━━━━━━━━━━━━\nSelect the student's class:"
            }.get(lang, "Select class:")
            await safe_edit_or_answer(query, prompt_cls, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return
        else:
            prompt_prop = {
                "tr": "💡 <b>Kurumsal Teklif / Öneri</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen diğer öğretmenlerin ve okul yönetiminin oylamasına sunmak istediğiniz teklifinizi mesaj olarak yazıp gönderiniz:\n\n<i>Bu teklif tüm kadro için ortak oylamaya sunulacaktır.</i>",
                "ru": "💡 <b>Школьное предложение / Идея</b>\n━━━━━━━━━━━━━━━━━━━━\nВведите тему и подробное описание вашего предложения для голосования учителей и администрации:\n\n<i>Предложение будет вынесено на общее голосование.</i>",
                "uz": "💡 <b>Maktab taklifi / Tashabbus</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, barcha o'qituvchilar va ma'muriyat ovoz berishiga qo'ymoqchi bo'lgan taklifingiz mazmuni va tavsifini yozib yuboring:\n\n<i>Ushbu taklif barcha xodimlar uchun umumiy ovoz berishga qo'yiladi.</i>",
                "en": "💡 <b>School Proposal / Initiative</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease type and send the title and detailed description of your proposal for staff voting:\n\n<i>This proposal will be submitted for ballot review.</i>"
            }.get(lang, "Enter proposal:")
            cancel_inline_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="tch:prop_hub")]
            ])
            await safe_edit_or_answer(query, prompt_prop, reply_markup=cancel_inline_kb, parse_mode="HTML")
            await state.set_state(Form.waiting_prop_content)
            await query.answer()

@router.callback_query(F.data.startswith("tch:prop_cls:"))
async def cb_teacher_prop_select_class(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.full_name))).scalars().all()

        if not students:
            no_st_txt = {
                "tr": f"⚠️ <b>{class_name} Sınıfı</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ Bu sınıfta henüz kayıtlı öğrenci bulunmuyor.",
                "ru": f"⚠️ <b>Класс {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ В этом классе пока нет зарегистрированных учеников.",
                "uz": f"⚠️ <b>{class_name} sinfi</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ Ushbu sinfda hozircha ro'yxatdan o'tgan o'quvchilar mavjud emas.",
                "en": f"⚠️ <b>Class {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\n❌ No registered students found in this class."
            }.get(lang, "No students in class.")
            no_st_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_new:student_complaint")]
            ])
            await safe_edit_or_answer(query, no_st_txt, reply_markup=no_st_kb, parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        for s in students:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"tch:prop_st:{s.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_new:student_complaint")])

        prompt_st = {
            "tr": f"⚠️ <b>{class_name} Sınıfı</b>\n━━━━━━━━━━━━━━━━━━━━\nŞikayet konusu olan öğrenciyi seçiniz:",
            "ru": f"⚠️ <b>Класс {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\nВыберите ученика:",
            "uz": f"⚠️ <b>{class_name} sinfi</b>\n━━━━━━━━━━━━━━━━━━━━\nShikoyat qilinadigan o'quvchini tanlang:",
            "en": f"⚠️ <b>Class {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\nSelect student:"
        }.get(lang, "Select student:")
        await safe_edit_or_answer(query, prompt_st, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:prop_st:"))
async def cb_teacher_prop_select_student(query: CallbackQuery, state: FSMContext):
    st_id = int(query.data.split(":")[2])
    PROP_CACHE.setdefault(query.from_user.id, {})["student_id"] = st_id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        st_name = st.full_name if st else "Öğrenci"

        prompt_c = {
            "tr": f"⚠️ <b>Öğrenci Şikayeti: {escape_html(st_name)}</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen öğrencinin sergilediği disiplinsiz davranışı, olayın detaylarını ve gerekçenizi mesaj olarak yazıp gönderiniz:\n\n<i>Bu şikayet tüm öğretmenler ve okul yönetimi kurulunun ortak değerlendirmesine ve oylamasına sunulacaktır.</i>",
            "ru": f"⚠️ <b>Жалоба на ученика: {escape_html(st_name)}</b>\n━━━━━━━━━━━━━━━━━━━━\nОпишите суть нарушения и подробности поведения ученика сообщением:\n\n<i>Жалоба будет направлена на голосование и рассмотрение педсовета.</i>",
            "uz": f"⚠️ <b>O'quvchi shikoyati: {escape_html(st_name)}</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, o'quvchining intizomsiz xulq-atvori va voqea tafsilotlarini xabar tarzida yozib yuboring:\n\n<i>Ushbu shikoyat barcha o'qituvchilar va ma'muriyat muhokamasiga qo'yiladi.</i>",
            "en": f"⚠️ <b>Student Complaint: {escape_html(st_name)}</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease type and send the details of the disciplinary issue:\n\n<i>This complaint will be submitted for staff review and voting.</i>"
        }.get(lang, "Enter complaint details:")

        cancel_inline_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"tch:prop_cls:{st.class_name if st else ''}")]
        ])
        await safe_edit_or_answer(query, prompt_c, reply_markup=cancel_inline_kb, parse_mode="HTML")
        await state.set_state(Form.waiting_prop_content)
        await query.answer()
@router.message(Form.waiting_prop_content)
async def process_teacher_proposal_content(message: Message, state: FSMContext):
    user_id = message.from_user.id
    text_content = (message.text or "").strip()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        # 1. Validation check (minimum 3 characters)
        if len(text_content) < 3:
            warn_msg = {
                "tr": "⚠️ <b>Geçersiz Teklif Metni</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen teklifiniz veya şikayetiniz için en az 3 karakterden oluşan açıklayıcı bir metin yazınız:",
                "ru": "⚠️ <b>Слишком короткое описание</b>\n━━━━━━━━━━━━━━━━━━━━\nПожалуйста, введите содержательное описание предложения (не менее 3 символов):",
                "uz": "⚠️ <b>Matn juda qisqa</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, taklif yoki shikoyatingiz bo'yicha kamida 3 ta belgidan iborat aniq ma'lumot kiriting:",
                "en": "⚠️ <b>Text Too Short</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter a meaningful proposal description (at least 3 characters):"
            }.get(lang, "Please enter at least 3 characters.")
            cancel_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="tch:prop_hub")]
            ])
            await message.answer(warn_msg, reply_markup=cancel_kb, parse_mode="HTML")
            return

        # 2. Extract cache & state
        cache = PROP_CACHE.pop(user_id, {})
        prop_type = cache.get("type", "teacher_proposal")
        st_id = cache.get("student_id")
        await state.clear()

        # Delete prompt card from chat to keep screen clean
        old_prompt_id = LAST_MENU_MSG_ID.get(message.chat.id)
        if old_prompt_id:
            try: await message.bot.delete_message(chat_id=message.chat.id, message_id=old_prompt_id)
            except Exception: pass

        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user_id))).scalar_one_or_none()
        creator_name = tch.full_name if tch else (user.full_name or get_role_label("teacher", lang))
        st = await session.get(Student, st_id) if st_id else None

        if prop_type == "student_complaint" and st:
            title = f"Öğrenci Şikayeti: {st.full_name} ({st.class_name})"
        else:
            title = text_content[:50] + ("..." if len(text_content) > 50 else "")

        proposal = Proposal(
            proposal_type=prop_type,
            title=title,
            content=text_content,
            target_audience="all_staff",
            target_student_id=st_id,
            created_by=user_id,
            creator_name=creator_name,
            status="active"
        )
        session.add(proposal)
        await session.commit()

        # 3. Dispatch localized ballot card to other staff members
        target_staff_ids = set(ADMIN_IDS)
        db_staff_users = (await session.execute(
            select(User).where(
                or_(
                    User.role.in_(["teacher", "admin"]),
                    User.telegram_id.in_(ADMIN_IDS)
                ),
                User.is_bot_blocked == False,
                User.is_blacklisted == False
            )
        )).scalars().all()
        user_lang_map = {u.telegram_id: (u.language or "tr") for u in db_staff_users if u.telegram_id}
        for u in db_staff_users:
            if u.telegram_id:
                target_staff_ids.add(u.telegram_id)

        tch_records = (await session.execute(select(Teacher.telegram_id).where(Teacher.telegram_id.isnot(None)))).scalars().all()
        for tid in tch_records:
            if tid:
                target_staff_ids.add(tid)

        dispatched_count = 0
        for target_tid in target_staff_ids:
            if target_tid == user_id:
                continue  # Creator receives confirmation card instead of ballot
            su_lang = user_lang_map.get(target_tid, "tr")
            type_lbl_su = {
                "student_complaint": {
                    "tr": "⚠️ Öğrenci Şikayeti & Disiplin",
                    "ru": "⚠️ Жалоба на ученика и дисциплина",
                    "uz": "⚠️ O'quvchi shikoyati va intizom",
                    "en": "⚠️ Student Disciplinary Complaint"
                }.get(su_lang, "Student Complaint"),
                "teacher_proposal": {
                    "tr": "💡 Kurumsal Eğitim Teklifi",
                    "ru": "💡 Школьное предложение",
                    "uz": "💡 Maktab ta'lim taklifi",
                    "en": "💡 School Initiative Proposal"
                }.get(su_lang, "School Proposal"),
                "admin_proposal": {
                    "tr": "👑 Yönetici İdari Teklifi",
                    "ru": "👑 Административное предложение",
                    "uz": "👑 Ma'muriy tashabbus",
                    "en": "👑 Administrative Proposal"
                }.get(su_lang, "Admin Proposal")
            }.get(prop_type, "Teklif")

            st_row_su = {
                "tr": f"🧑‍🎓 <b>İlgili Öğrenci:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n",
                "ru": f"🧑‍🎓 <b>Ученик:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n",
                "uz": f"🧑‍🎓 <b>Tegishli o'quvchi:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n",
                "en": f"🧑‍🎓 <b>Student:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)})\n"
            }.get(su_lang, "") if st else ""

            vote_card_su = {
                "tr": (
                    f"🗳️ <b>YENİ OYLAMA VE İSTİŞARE TALEBİ (#{proposal.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Teklif Eden:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Kategori:</b> {type_lbl_su}\n"
                    f"{st_row_su}"
                    f"📝 <b>Açıklama / Gerekçe:</b>\n{escape_html(text_content)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Lütfen görüşünüzü belirtiniz (Gerekçenizi de yazabilirsiniz):"
                ),
                "ru": (
                    f"🗳️ <b>НОВОЕ ГОЛОСОВАНИЕ И ОБСУЖДЕНИЕ (#{proposal.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Автор:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Категория:</b> {type_lbl_su}\n"
                    f"{st_row_su}"
                    f"📝 <b>Описание / Обоснование:</b>\n{escape_html(text_content)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Пожалуйста, выразите ваше мнение (вы также можете оставить комментарий):"
                ),
                "uz": (
                    f"🗳️ <b>YANGI OVOZ BERISH VA MUHOKAMA (#{proposal.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Taklif egasi:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Toifa:</b> {type_lbl_su}\n"
                    f"{st_row_su}"
                    f"📝 <b>Mazmuni / Asos:</b>\n{escape_html(text_content)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Iltimos, o'z fikringizni bildiring (izoh ham qoldirishingiz mumkin):"
                ),
                "en": (
                    f"🗳️ <b>NEW BALLOT & DISCUSSION REQUEST (#{proposal.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Submitted by:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Category:</b> {type_lbl_su}\n"
                    f"{st_row_su}"
                    f"📝 <b>Details / Rationale:</b>\n{escape_html(text_content)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Please cast your vote (you may also attach your reasoning):"
                )
            }.get(su_lang, "New Ballot")

            vote_buttons_su = [
                [
                    InlineKeyboardButton(text={"tr": "✅ Katılıyorum", "ru": "✅ Согласен", "uz": "✅ Qo'shilaman", "en": "✅ Agree"}.get(su_lang, "Agree"), callback_data=f"prop:vote:{proposal.id}:agree"),
                    InlineKeyboardButton(text={"tr": "❌ Katılmıyorum", "ru": "❌ Не согласен", "uz": "❌ Qo'shilmayman", "en": "❌ Disagree"}.get(su_lang, "Disagree"), callback_data=f"prop:vote:{proposal.id}:disagree")
                ],
                [InlineKeyboardButton(text={"tr": "📝 Gerekçe / Fikir Ekle", "ru": "📝 Добавить комментарий", "uz": "📝 Izoh / Fikr qo'shish", "en": "📝 Add Note"}.get(su_lang, "Add Note"), callback_data=f"prop:reason:{proposal.id}")],
                [InlineKeyboardButton(text={"tr": "📊 Oylama Sonuçları", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Results"}.get(su_lang, "Results"), callback_data=f"prop:results:{proposal.id}")]
            ]

            try:
                sent_msg = await safe_send_message(
                    message.bot,
                    chat_id=target_tid,
                    text=vote_card_su,
                    reply_markup=InlineKeyboardMarkup(inline_keyboard=vote_buttons_su),
                    parse_mode="HTML"
                )
                if sent_msg:
                    dispatched_count += 1
                await asyncio.sleep(0.04)
            except Exception:
                pass

        # 4. Immediate confirmation card to the creator
        succ_txt = {
            "tr": (
                f"✅ <b>Teklifiniz Başarıyla Yayınlandı (#{proposal.id})!</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"📌 <b>Başlık:</b> {escape_html(title)}\n\n"
                "Teklifiniz tüm öğretmenlerin ve okul yönetiminin oylamasına sunuldu. Oylama sonuçlarını canlı olarak takip edebilirsiniz."
            ),
            "ru": (
                f"✅ <b>Ваше предложение успешно опубликовано (#{proposal.id})!</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"📌 <b>Тема:</b> {escape_html(title)}\n\n"
                "Предложение вынесено на голосование учителей и администрации. Вы можете отслеживать ход голосования."
            ),
            "uz": (
                f"✅ <b>Taklifingiz muvaffaqiyatli e'lon qilindi (#{proposal.id})!</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"📌 <b>Mavzu:</b> {escape_html(title)}\n\n"
                "Taklifingiz barcha o'qituvchilar va ma'muriyat ovoz berishiga taqdim etildi. Ovoz berish natijalarini kuzatib borishingiz mumkin."
            ),
            "en": (
                f"✅ <b>Your Proposal Was Successfully Published (#{proposal.id})!</b>\n"
                "━━━━━━━━━━━━━━━━━━━━\n"
                f"📌 <b>Title:</b> {escape_html(title)}\n\n"
                "Submitted for faculty and administration review and voting."
            )
        }.get(lang, "Proposal published.")

        buttons_ret = [
            [InlineKeyboardButton(text={"tr": "📊 Oylama Sonuçları", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Results"}.get(lang, "Results"), callback_data=f"prop:results:{proposal.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_hub")]
        ]
        role_kb = get_role_reply_kb(user.role, lang)
        sent_m = await message.answer(succ_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons_ret), parse_mode="HTML")
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data.startswith("prop:vote:"))
async def cb_proposal_cast_vote(query: CallbackQuery):
    parts = query.data.split(":")
    prop_id = int(parts[2])
    choice = parts[3]  # 'agree' or 'disagree'
    user_id = query.from_user.id

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        prop = await session.get(Proposal, prop_id)
        if not prop or prop.status != "active":
            closed_tst = {"tr": "⚠️ Bu oylama tamamlanmış veya kapatılmıştır.", "ru": "⚠️ Это голосование уже завершено.", "uz": "⚠️ Ushbu ovoz berish yakunlangan.", "en": "⚠️ This ballot is closed."}.get(lang, "Closed.")
            await query.answer(closed_tst, show_alert=True)
            return

        existing_vote = (await session.execute(select(ProposalVote).where(ProposalVote.proposal_id == prop_id, ProposalVote.user_telegram_id == user_id))).scalar_one_or_none()
        voter_name = user.full_name if (user and user.full_name) else "Öğretmen"

        if existing_vote:
            existing_vote.vote_choice = choice
            existing_vote.voted_at = datetime.utcnow()
        else:
            new_vote = ProposalVote(
                proposal_id=prop_id,
                user_telegram_id=user_id,
                voter_name=voter_name,
                vote_choice=choice
            )
            session.add(new_vote)

        await session.commit()

        choice_lbl = {
            "agree": {"tr": "✅ Katılıyorum", "ru": "✅ Согласен", "uz": "✅ Qo'shilaman", "en": "✅ Agree"}.get(lang, "Agree"),
            "disagree": {"tr": "❌ Katılmıyorum", "ru": "❌ Не согласен", "uz": "❌ Qo'shilmayman", "en": "❌ Disagree"}.get(lang, "Disagree")
        }.get(choice, choice)
        tst = {"tr": f"✅ Oyunuz kaydedildi: {choice_lbl}", "ru": f"✅ Ваш голос принят: {choice_lbl}", "uz": f"✅ Ovozingiz qabul qilindi: {choice_lbl}", "en": f"✅ Vote recorded: {choice_lbl}"}.get(lang, "Vote recorded.")
        await query.answer(tst, show_alert=True)

@router.callback_query(F.data.startswith("prop:reason:"))
async def cb_proposal_reason_init(query: CallbackQuery, state: FSMContext):
    prop_id = int(query.data.split(":")[2])
    VOTE_REASON_CACHE[query.from_user.id] = prop_id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    p_prompt = {
        "tr": "📝 <b>Oylama Gerekçesi / Fikriniz</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen teklife katılım veya ret gerekçenizi yazınız (Yönetim ve kurul kararında incelenecektir):",
        "ru": "📝 <b>Причина / Ваше мнение</b>\n━━━━━━━━━━━━━━━━━━━━\nВведите обоснование вашего голоса:",
        "uz": "📝 <b>Izoh / Sabab</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, fikringiz sababini yozing:",
        "en": "📝 <b>Vote Justification / Note</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter your reason or note:"
    }.get(lang, "Enter justification:")

    cancel_btn = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data=f"prop:results:{prop_id}")]])
    await safe_edit_or_answer(query, p_prompt, reply_markup=cancel_btn, parse_mode="HTML")
    await state.set_state(Form.waiting_prop_reason)
    await query.answer()

@router.message(Form.waiting_prop_reason)
async def process_proposal_reason_text(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    reason_txt = message.text.strip()
    prop_id = VOTE_REASON_CACHE.pop(message.from_user.id, None)
    await state.clear()

    user_id = message.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        if not prop_id: return

        voter_name = user.full_name if (user and user.full_name) else "Öğretmen"
        vote = (await session.execute(select(ProposalVote).where(ProposalVote.proposal_id == prop_id, ProposalVote.user_telegram_id == user_id))).scalar_one_or_none()
        if not vote:
            vote = ProposalVote(
                proposal_id=prop_id,
                user_telegram_id=user_id,
                voter_name=voter_name,
                vote_choice="agree",
                reason_note=reason_txt
            )
            session.add(vote)
        else:
            vote.reason_note = reason_txt

        await session.commit()

        # Notify the creator of the proposal about the new feedback / opinion
        prop = await session.get(Proposal, prop_id)
        if prop and prop.created_by and prop.created_by != user_id:
            cr_user = await session.get(User, prop.created_by)
            cr_lang = cr_user.language if (cr_user and cr_user.language) else "tr"
            voter_role_desc = {
                "tr": "Yönetici" if is_admin_user(user, user_id) else "Öğretmen",
                "ru": "Руководство" if is_admin_user(user, user_id) else "Учитель",
                "uz": "Ma'muriyat" if is_admin_user(user, user_id) else "O'qituvchi",
                "en": "Administration" if is_admin_user(user, user_id) else "Teacher"
            }.get(cr_lang, "Staff")

            vote_icon = "✅" if (vote and vote.vote_choice == "agree") else "❌"
            vote_status_desc = {
                "tr": "Destekliyor (Kabul)" if (vote and vote.vote_choice == "agree") else "Karşı Çıkıyor (Ret)",
                "ru": "Поддерживает (За)" if (vote and vote.vote_choice == "agree") else "Против (Не согласен)",
                "uz": "Qo'llab-quvvatlaydi (Rozilik)" if (vote and vote.vote_choice == "agree") else "Qarshi",
                "en": "In Favor (Agrees)" if (vote and vote.vote_choice == "agree") else "Opposed (Disagrees)"
            }.get(cr_lang, vote_icon)

            cr_notif = {
                "tr": (
                    f"💬 <b>TEKLİFİNİZE YENİ GÖRÜŞ / GEREKÇE BİLDİRİLDİ (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Teklifiniz:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>Görüş Bildiren:</b> {escape_html(voter_name)} ({voter_role_desc})\n"
                    f"🏁 <b>Oy Tercihi:</b> {vote_icon} <b>{vote_status_desc}</b>\n\n"
                    f"📝 <b>İletilen Görüş / Gerekçe:</b>\n<i>\"{escape_html(reason_txt)}\"</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "ru": (
                    f"💬 <b>НОВЫЙ КОММЕНТАРИЙ К ВАШЕМУ ПРЕДЛОЖЕНИЮ (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Ваше предложение:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>От кого:</b> {escape_html(voter_name)} ({voter_role_desc})\n"
                    f"🏁 <b>Позиция:</b> {vote_icon} <b>{vote_status_desc}</b>\n\n"
                    f"📝 <b>Текст комментария / Обоснование:</b>\n<i>\"{escape_html(reason_txt)}\"</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "uz": (
                    f"💬 <b>TAKLIFINGIZGA YANGI FIKR / IZOH KELDI (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Taklifingiz:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>Fikr bildiruvchi:</b> {escape_html(voter_name)} ({voter_role_desc})\n"
                    f"🏁 <b>Ovozi:</b> {vote_icon} <b>{vote_status_desc}</b>\n\n"
                    f"📝 <b>Bildirilgan fikr / Izoh:</b>\n<i>\"{escape_html(reason_txt)}\"</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                ),
                "en": (
                    f"💬 <b>NEW FEEDBACK ON YOUR PROPOSAL (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"📌 <b>Your Proposal:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>From:</b> {escape_html(voter_name)} ({voter_role_desc})\n"
                    f"🏁 <b>Vote Position:</b> {vote_icon} <b>{vote_status_desc}</b>\n\n"
                    f"📝 <b>Comment / Feedback:</b>\n<i>\"{escape_html(reason_txt)}\"</i>\n"
                    "━━━━━━━━━━━━━━━━━━━━"
                )
            }.get(cr_lang, f"New note on proposal #{prop.id}")

            btn_v_res = {"tr": "📊 Oylama Sonuçları", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Results"}.get(cr_lang, "Results")
            cr_ikb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=btn_v_res, callback_data=f"prop:results:{prop.id}")]])
            try:
                await safe_send_message(message.bot, chat_id=prop.created_by, text=cr_notif, reply_markup=cr_ikb, parse_mode="HTML")
            except Exception:
                pass

        done_r = {"tr": "✅ Gerekçeniz başarıyla oylama kaydınıza eklendi!", "ru": "✅ Ваш комментарий успешно добавлен!", "uz": "✅ Izohingiz muvaffaqiyatli saqlandi!", "en": "✅ Reason successfully attached to your vote!"}.get(lang, "Reason saved.")
        b_res = {"tr": "📊 Sonuçları Gör", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Results"}.get(lang, "📊 Results")
        buttons = [
            [InlineKeyboardButton(text=b_res, callback_data=f"prop:results:{prop_id}")]
        ]
        role_kb = get_role_reply_kb(user.role, lang)
        sent_m = await safe_send_message(message.bot, chat_id=message.chat.id, text=done_r, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data.startswith("prop:results:"))
async def cb_proposal_view_results(query: CallbackQuery):
    prop_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        prop = await session.get(Proposal, prop_id)
        if not prop:
            not_found_tst = {"tr": "Teklif bulunamadı.", "ru": "Предложение не найдено.", "uz": "Taklif topilmadi.", "en": "Proposal not found."}.get(lang, "Not found.")
            await query.answer(not_found_tst, show_alert=True)
            return

        votes = (await session.execute(select(ProposalVote).where(ProposalVote.proposal_id == prop_id))).scalars().all()
        agrees = [v for v in votes if v.vote_choice == "agree"]
        disagrees = [v for v in votes if v.vote_choice == "disagree"]
        total_votes = len(votes)

        pct_agree = round((len(agrees) / total_votes * 100), 1) if total_votes > 0 else 0.0
        prog_bar = render_progress_bar(int(pct_agree // 10), 10)

        # Reasons list
        reasons_list = []
        for v in votes:
            if v.reason_note:
                icon = "✅" if v.vote_choice == "agree" else "❌"
                reasons_list.append(f"• {icon} <b>{escape_html(v.voter_name)}:</b> <i>\"{escape_html(v.reason_note)}\"</i>")

        no_reasons_txt = {
            "tr": "<i>Henüz gerekçe eklenmemiş.</i>",
            "ru": "<i>Комментариев пока нет.</i>",
            "uz": "<i>Hozircha izohlar qo'shilmagan.</i>",
            "en": "<i>No reasons added yet.</i>"
        }.get(lang, "No reasons yet.")
        reasons_block = "\n".join(reasons_list[:8]) if reasons_list else no_reasons_txt

        status_badge = {
            "active": {
                "tr": "🟢 <b>Oylama Devam Ediyor</b>",
                "ru": "🟢 <b>Голосование продолжается</b>",
                "uz": "🟢 <b>Ovoz berish davom etmoqda</b>",
                "en": "🟢 <b>Ballot in Progress</b>"
            }.get(lang, "In Progress"),
            "accepted": {
                "tr": "🏆 <b>Yönetim Tarafından Kabul Edildi</b>",
                "ru": "🏆 <b>Одобрено администрацией</b>",
                "uz": "🏆 <b>Ma'muriyat tomonidan qabul qilindi</b>",
                "en": "🏆 <b>Approved by Administration</b>"
            }.get(lang, "Approved"),
            "rejected": {
                "tr": "❌ <b>Yönetim Tarafından Reddedildi</b>",
                "ru": "❌ <b>Отклонено администрацией</b>",
                "uz": "❌ <b>Ma'muriyat tomonidan rad etildi</b>",
                "en": "❌ <b>Rejected by Administration</b>"
            }.get(lang, "Rejected")
        }.get(prop.status, prop.status)

        dec_note_block = ""
        if prop.admin_decision_note:
            dec_lbl = {
                "tr": "📌 <b>Yönetim Karar Gerekçesi:</b>",
                "ru": "📌 <b>Обоснование руководства:</b>",
                "uz": "📌 <b>Ma'muriyat qaror asosi:</b>",
                "en": "📌 <b>Administrative Rationale:</b>"
            }.get(lang, "Administrative Decision:")
            dec_note_block = f"\n{dec_lbl}\n<i>{escape_html(prop.admin_decision_note)}</i>\n"

        card_title = {
            "tr": f"📊 <b>OYLAMA VE İSTİŞARE SONUÇLARI (#{prop.id})</b>",
            "ru": f"📊 <b>РЕЗУЛЬТАТЫ ГОЛОСОВАНИЯ (#{prop.id})</b>",
            "uz": f"📊 <b>OVOZ BERISH VA MUHOKAMA NATIJALARI (#{prop.id})</b>",
            "en": f"📊 <b>BALLOT & DISCUSSION RESULTS (#{prop.id})</b>"
        }.get(lang, "Ballot Results")

        lbl_title = {"tr": "Başlık", "ru": "Тема", "uz": "Mavzu", "en": "Title"}.get(lang, "Title")
        lbl_author = {"tr": "Teklif Eden", "ru": "Автор", "uz": "Taklif egasi", "en": "Submitted by"}.get(lang, "Submitted by")
        lbl_status = {"tr": "Durum", "ru": "Статус", "uz": "Holat", "en": "Status"}.get(lang, "Status")
        lbl_distrib = {"tr": "Oy Dağılımı", "ru": "Распределение голосов", "uz": "Ovozlar taqsimoti", "en": "Vote Distribution"}.get(lang, "Distribution")
        lbl_support = {"tr": "Destek", "ru": "Поддержка", "uz": "Qo'llab-quvvatlash", "en": "Support"}.get(lang, "Support")
        lbl_agrees = {"tr": "Katılanlar (Destekleyenler)", "ru": "Согласны (Поддерживают)", "uz": "Qo'shilganlar (Qo'llaganlar)", "en": "In Favor (Agreed)"}.get(lang, "In Favor")
        lbl_disagrees = {"tr": "Karşı Olanlar", "ru": "Против", "uz": "Qarshilar", "en": "Opposed"}.get(lang, "Opposed")
        lbl_total = {"tr": "Toplam Kullanılan Oy", "ru": "Всего проголосовало", "uz": "Jami ovozlar", "en": "Total Votes Cast"}.get(lang, "Total Votes")
        lbl_reasons = {"tr": "Gerekçeler ve Öğretmen Görüşleri", "ru": "Комментарии и мнения учителей", "uz": "Izohlar va o'qituvchilar fikri", "en": "Staff Notes & Feedback"}.get(lang, "Feedback")
        lbl_unit = {"tr": "Kişi", "ru": "чел.", "uz": "ta", "en": "votes"}.get(lang, "")

        card_text = (
            f"{card_title}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📌 <b>{lbl_title}:</b> {escape_html(prop.title)}\n"
            f"👤 <b>{lbl_author}:</b> {escape_html(prop.creator_name)}\n"
            f"🏁 <b>{lbl_status}:</b> {status_badge}\n"
            f"{dec_note_block}"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"📈 <b>{lbl_distrib}:</b>\n"
            f"<code>{prog_bar}</code> <b>%{pct_agree} {lbl_support}</b>\n"
            f"• ✅ <b>{lbl_agrees}:</b> {len(agrees)} {lbl_unit}\n"
            f"• ❌ <b>{lbl_disagrees}:</b> {len(disagrees)} {lbl_unit}\n"
            f"• 👥 <b>{lbl_total}:</b> {total_votes}\n\n"
            f"💬 <b>{lbl_reasons}:</b>\n"
            f"{reasons_block}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = []
        if prop.status == "active":
            b_vote_btn = {"tr": "🗳️ Oy Ver / Değiştir", "ru": "🗳️ Голосовать / Изменить", "uz": "🗳️ Ovoz berish", "en": "🗳️ Cast / Change Vote"}.get(lang, "🗳️ Vote")
            buttons.append([InlineKeyboardButton(text=b_vote_btn, callback_data=f"prop:open:{prop.id}")])

        if is_admin_user(user, query.from_user.id) and prop.status == "active":
            btn_acc = {"tr": "🏆 Kabul Et & Onayla", "ru": "🏆 Принять решение", "uz": "🏆 Qabul qilish", "en": "🏆 Accept & Enact"}.get(lang, "Accept")
            btn_rej = {"tr": "❌ Reddet", "ru": "❌ Отклонить", "uz": "❌ Rad etish", "en": "❌ Reject"}.get(lang, "Reject")
            buttons.append([
                InlineKeyboardButton(text=btn_acc, callback_data=f"adm:prop_decide:{prop.id}:accept"),
                InlineKeyboardButton(text=btn_rej, callback_data=f"adm:prop_decide:{prop.id}:reject")
            ])

        back_cb = "adm:proposals_hub" if is_admin_user(user, query.from_user.id) else "tch:active_props:0"
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=back_cb)])

        await safe_edit_or_answer(query, card_text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:proposals_hub")
async def cb_admin_proposals_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        props = (await session.execute(select(Proposal).order_by(desc(Proposal.created_at)).limit(10))).scalars().all()

        bc_ap = {"tr": "🏠 Ana Menü ➔ 🛠️ Araçlar ➔ 🗳️ Oylamalar", "ru": "🏠 Главное меню ➔ 🛠️ Инструменты ➔ 🗳️ Голосования", "uz": "🏠 Asosiy menyu ➔ 🛠️ Boshqaruv ➔ 🗳️ Ovoz berish", "en": "🏠 Main Menu ➔ 🛠️ Admin Tools ➔ 🗳️ Proposals"}.get(lang, "🗳️ Voting Desk")
        title_ap = {
            "tr": "🗳️ <b>YÖNETİM OYLAMA, ŞİKAYET VE KARAR MASASI</b>",
            "ru": "🗳️ <b>ЦЕНТР ГОЛОСОВАНИЙ И АДМИНИСТРАТИВНЫХ РЕШЕНИЙ</b>",
            "uz": "🗳️ <b>MA'MURIYAT OVOZ BERISH VA QAROR MARKAZI</b>",
            "en": "🗳️ <b>ADMINISTRATIVE BALLOTS & DECISIONS DESK</b>"
        }.get(lang, "🗳️ <b>PROPOSALS & DECISIONS</b>")

        desc_ap = {
            "tr": "Öğretmenlerin sunduğu öğrenci disiplin şikayetlerini ve teklifleri inceleyebilir, oylama sonuçlarına göre idari kararı açıklayabilirsiniz:",
            "ru": "Изучайте предложения и дисциплинарные жалобы учителей, принимайте решения по итогам голосования:",
            "uz": "O'qituvchilar taklif va intizom shikoyatlarini ko'rib chiqishingiz va ovoz berish natijalariga ko'ra qaror qabul qilishingiz mumkin:",
            "en": "Review faculty proposals and student disciplinary complaints, and issue official administrative decisions:"
        }.get(lang, "Manage proposals:")

        lbl_recent = {
            "tr": f"📋 <b>Son Teklif ve Şikayetler ({len(props)}):</b>",
            "ru": f"📋 <b>Недавние предложения и жалобы ({len(props)}):</b>",
            "uz": f"📋 <b>So'nggi taklif va shikoyatlar ({len(props)}):</b>",
            "en": f"📋 <b>Recent Proposals ({len(props)}):</b>"
        }.get(lang, f"Proposals ({len(props)}):")

        text_content = (
            f"<b>{bc_ap}</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{title_ap}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{desc_ap}\n\n"
            f"{lbl_recent}\n"
            "━━━━━━━━━━━━━━━━━━━━"
        )

        buttons = []
        for p in props:
            st_icon = "🟢" if p.status == "active" else ("🏆" if p.status == "accepted" else "❌")
            buttons.append([InlineKeyboardButton(text=f"{st_icon} #{p.id} {p.title[:25]}", callback_data=f"prop:results:{p.id}")])

        buttons.append([InlineKeyboardButton(text={"tr": "➕ Yeni Oylama / Teklif", "ru": "➕ Создать голосование", "uz": "➕ Yangi ovoz berish", "en": "➕ Start New Proposal"}.get(lang, "➕ Start New Proposal"), callback_data="adm:new_admin_prop")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_tools_reports"))

        await safe_edit_or_answer(query, text_content, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:prop_decide:"))
async def cb_admin_proposal_decide_init(query: CallbackQuery, state: FSMContext):
    parts = query.data.split(":")
    prop_id = int(parts[2])
    decision = parts[3]  # 'accept' or 'reject'
    PROP_CACHE[query.from_user.id] = {"prop_id": prop_id, "decision": decision}

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    dec_labels = {
        "accept": {
            "tr": "KABUL VE ONAY",
            "ru": "ОДОБРЕНИЕ И ПРИНЯТИЕ",
            "uz": "QABUL VA TASDIQ",
            "en": "APPROVAL & ADOPTION"
        },
        "reject": {
            "tr": "RET / ARŞİV",
            "ru": "ОТКЛОНЕНИЕ / В АРХИВ",
            "uz": "RAD ETISH / ARXIV",
            "en": "REJECTION / ARCHIVE"
        }
    }
    cur_dec_lbl = dec_labels.get(decision, {}).get(lang, decision.upper())

    prompt_d = {
        "tr": (
            f"⚖️ <b>İdari Karar Belirleme: #{prop_id} ({cur_dec_lbl})</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "Lütfen personele ve teklif sahibine iletilecek resmi idari karar gerekçesini yazınız:"
        ),
        "ru": (
            f"⚖️ <b>Административное решение: #{prop_id} ({cur_dec_lbl})</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "Пожалуйста, введите официальное обоснование решения для автора и коллектива:"
        ),
        "uz": (
            f"⚖️ <b>Ma'muriy qaror qabul qilish: #{prop_id} ({cur_dec_lbl})</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "Iltimos, xodimlar va taklif egasiga yuboriladigan rasmiy qaror asosini kiriting:"
        ),
        "en": (
            f"⚖️ <b>Administrative Decision: #{prop_id} ({cur_dec_lbl})</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "Please enter the official administrative rationale for the author and staff:"
        )
    }.get(lang, f"Decision for #{prop_id}:")

    cancel_kb = get_inline_cancel_kb(lang, back_callback=f"prop:results:{prop_id}")
    await safe_edit_or_answer(query, prompt_d, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_admin_decision_note)
    await query.answer()

@router.message(Form.waiting_admin_decision_note)
async def process_admin_decision_note(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    note_txt = message.text.strip()
    cache = PROP_CACHE.pop(message.from_user.id, {})
    prop_id = cache.get("prop_id")
    decision = cache.get("decision", "accept")
    await state.clear()

    # Delete prompt card from chat to keep screen clean
    old_prompt_id = LAST_MENU_MSG_ID.get(message.chat.id)
    if old_prompt_id:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=old_prompt_id)
        except Exception: pass

    user_id = message.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        prop = await session.get(Proposal, prop_id)
        if not prop: return

        prop.status = "accepted" if decision == "accept" else "rejected"
        prop.admin_decision_note = note_txt
        prop.decided_by = user_id
        prop.decided_at = datetime.utcnow()

        await log_audit(session, user_id, (user.full_name or "Yönetici"), f"OYLAMA KARARI: {prop.status.upper()}", f"Teklif #{prop.id} karara bağlandı: {note_txt}")
        await session.commit()

        # Gather all target staff IDs (admins + staff + creator)
        target_ids = set(ADMIN_IDS)
        if prop.created_by:
            target_ids.add(prop.created_by)

        db_staff_users = (await session.execute(
            select(User).where(
                or_(
                    User.role.in_(["teacher", "admin"]),
                    User.telegram_id.in_(ADMIN_IDS)
                ),
                User.is_bot_blocked == False,
                User.is_blacklisted == False
            )
        )).scalars().all()
        user_lang_map = {u.telegram_id: (u.language or "tr") for u in db_staff_users if u.telegram_id}
        for u in db_staff_users:
            if u.telegram_id:
                target_ids.add(u.telegram_id)

        tch_records = (await session.execute(select(Teacher.telegram_id).where(Teacher.telegram_id.isnot(None)))).scalars().all()
        for tid in tch_records:
            if tid:
                target_ids.add(tid)

        for target_tid in target_ids:
            if target_tid == user_id:
                continue  # The deciding admin receives the dedicated confirmation card below!
            t_lang = user_lang_map.get(target_tid, "tr")
            badge_t = {
                "accepted": {
                    "tr": "🏆 KABUL EDİLDİ",
                    "ru": "🏆 ОДОБРЕНО",
                    "uz": "🏆 QABUL QILINDI",
                    "en": "🏆 APPROVED"
                },
                "rejected": {
                    "tr": "❌ REDDEDİLDİ",
                    "ru": "❌ ОТКЛОНЕНО",
                    "uz": "❌ RAD ETILDI",
                    "en": "❌ REJECTED"
                }
            }.get(prop.status, {}).get(t_lang, prop.status.upper())

            notif_t = {
                "tr": (
                    f"⚖️ <b>İdari Karar Açıklandı (#{prop.id})</b>\n"
                    "──────────────\n"
                    f"📌 <b>Teklif:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>Sahibi:</b> {escape_html(prop.creator_name)}\n"
                    f"🏁 <b>Karar:</b> <b>{badge_t}</b>\n\n"
                    f"📝 <b>Gerekçe:</b>\n<i>{escape_html(note_txt)}</i>\n"
                    "──────────────"
                ),
                "ru": (
                    f"⚖️ <b>Решение руководства (#{prop.id})</b>\n"
                    "──────────────\n"
                    f"📌 <b>Тема:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>Автор:</b> {escape_html(prop.creator_name)}\n"
                    f"🏁 <b>Решение:</b> <b>{badge_t}</b>\n\n"
                    f"📝 <b>Обоснование:</b>\n<i>{escape_html(note_txt)}</i>\n"
                    "──────────────"
                ),
                "uz": (
                    f"⚖️ <b>Ma'muriy qaror e'lon qilindi (#{prop.id})</b>\n"
                    "──────────────\n"
                    f"📌 <b>Taklif:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>Muallif:</b> {escape_html(prop.creator_name)}\n"
                    f"🏁 <b>Qaror:</b> <b>{badge_t}</b>\n\n"
                    f"📝 <b>Asos:</b>\n<i>{escape_html(note_txt)}</i>\n"
                    "──────────────"
                ),
                "en": (
                    f"⚖️ <b>Administrative Decision (#{prop.id})</b>\n"
                    "──────────────\n"
                    f"📌 <b>Proposal:</b> {escape_html(prop.title)}\n"
                    f"👤 <b>Author:</b> {escape_html(prop.creator_name)}\n"
                    f"🏁 <b>Decision:</b> <b>{badge_t}</b>\n\n"
                    f"📝 <b>Rationale:</b>\n<i>{escape_html(note_txt)}</i>\n"
                    "──────────────"
                )
            }.get(t_lang, f"Decision on #{prop.id}")

            b_lbl = {"tr": "📊 Sonuçları İncele", "ru": "📊 Изучить результаты", "uz": "📊 Natijalarni ko'rish", "en": "📊 View Results"}.get(t_lang, "View Results")
            ikb = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=b_lbl, callback_data=f"prop:results:{prop.id}")]])
            try:
                await safe_send_message(message.bot, chat_id=target_tid, text=notif_t, reply_markup=ikb, parse_mode="HTML")
                await asyncio.sleep(0.04)
            except Exception:
                pass

        succ_txt = {
            "tr": f"✅ <b>Teklif #{prop.id} karara bağlandı!</b>\nKarar tüm personele ve teklif sahibine duyuruldu.",
            "ru": f"✅ <b>Решение по предложению #{prop.id} вынесено!</b>\nУведомление отправлено коллективу и автору.",
            "uz": f"✅ <b>#{prop.id}-sonli taklif bo'yicha qaror qabul qilindi!</b>\nXabarnoma jamoa va muallifga yuborildi.",
            "en": f"✅ <b>Decision enacted for Proposal #{prop.id}!</b>\nStaff and author have been notified."
        }.get(lang, f"Decision recorded for #{prop.id}!")
        buttons = [
            [InlineKeyboardButton(text={"tr": "📊 Oylama Masası", "ru": "📊 Центр голосований", "uz": "📊 Ovoz berish markazi", "en": "📊 Proposals Hub"}.get(lang, "📊 Proposals Hub"), callback_data="adm:proposals_hub")]
        ]
        role_kb = get_role_reply_kb("admin", lang)
        sent_m = await safe_send_message(
            message.bot,
            chat_id=message.chat.id,
            text=succ_txt,
            reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons),
            parse_mode="HTML"
        )
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data == "adm:new_admin_prop")
async def cb_admin_new_proposal_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    prompt_t = {
        "tr": "👑 <b>Yönetici Teklifi / İdari Oylama Başlatma</b>\n━━━━━━━━━━━━━━━━━━━━\nLütfen oylamaya sunmak istediğiniz teklifin <b>Başlığını</b> yazınız (Örn: Hafta Sonu Kurs Programı Düzenlemesi):",
        "ru": "👑 <b>Административное голосование / Инициатива</b>\n━━━━━━━━━━━━━━━━━━━━\nВведите <b>название</b> темы голосования (Напр: Организация дополнительных занятий):",
        "uz": "👑 <b>Ma'muriy tashabbus / Ovoz berish</b>\n━━━━━━━━━━━━━━━━━━━━\nIltimos, ovoz berishga qo'yiladigan mavzuning <b>nomini</b> kiriting (Masalan: Dam olish kunlari dars jadvali):",
        "en": "👑 <b>Administrative Ballot / Proposal</b>\n━━━━━━━━━━━━━━━━━━━━\nPlease enter the <b>title</b> of the ballot topic (e.g. Weekend Class Schedule Arrangement):"
    }.get(lang, "Enter proposal title:")

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:proposals_hub")
    await safe_edit_or_answer(query, prompt_t, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_admin_prop_title)
    await query.answer()

@router.message(Form.waiting_admin_prop_title)
async def process_admin_prop_title(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    title_text = message.text.strip()
    await state.update_data(prop_title=title_text)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_inline_cancel_kb(lang, back_callback="adm:proposals_hub")
    prompt_c = {
        "tr": f"👑 <b>Başlık:</b> {escape_html(title_text)}\n━━━━━━━━━━━━━━━━━━━━\nLütfen personele sunulacak teklifin <b>Detaylı Açıklamasını</b> yazınız:",
        "ru": f"👑 <b>Тема:</b> {escape_html(title_text)}\n━━━━━━━━━━━━━━━━━━━━\nВведите <b>подробное описание</b> предложения для коллектива:",
        "uz": f"👑 <b>Mavzu:</b> {escape_html(title_text)}\n━━━━━━━━━━━━━━━━━━━━\nIltimos, jamoaga taqdim etiladigan taklifning <b>batafsil tavsifini</b> kiriting:",
        "en": f"👑 <b>Title:</b> {escape_html(title_text)}\n━━━━━━━━━━━━━━━━━━━━\nPlease enter the <b>detailed description</b> of the proposal for the staff:"
    }.get(lang, "Enter description:")

    await safe_edit_or_answer(message, prompt_c, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.waiting_admin_prop_content)

@router.message(Form.waiting_admin_prop_content)
async def process_admin_prop_content(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    content_text = message.text.strip()
    data = await state.get_data()
    title_text = data.get("prop_title", "İdari Teklif")
    await state.clear()

    # Delete prompt card from chat
    old_prompt_id = LAST_MENU_MSG_ID.get(message.chat.id)
    if old_prompt_id:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=old_prompt_id)
        except Exception: pass

    user_id = message.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        creator_name = user.full_name or "Okul Yönetimi"

        prop = Proposal(
            proposal_type="admin_proposal",
            title=title_text,
            content=content_text,
            target_audience="all_staff",
            created_by=user_id,
            creator_name=creator_name,
            status="active"
        )
        session.add(prop)
        await session.commit()

        # Oylama kartını personele dağıt

        b_acc = {"tr": "✅ Kabul / Katılıyorum", "ru": "✅ Согласен", "uz": "✅ Qo'shilaman", "en": "✅ Agree"}.get(lang, "✅ Agree")
        b_rej = {"tr": "❌ Red / Karşıyım", "ru": "❌ Не согласен", "uz": "❌ Qo'shilmayman", "en": "❌ Disagree"}.get(lang, "❌ Disagree")
        b_rsn_a = {"tr": "📝 Gerekçe / Fikir Ekle", "ru": "📝 Добавить комментарий", "uz": "📝 Izoh / Fikr qo'shish", "en": "📝 Add Note"}.get(lang, "📝 Add Note")
        b_live = {"tr": "📊 Canlı Sonuçlar", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Live Results"}.get(lang, "📊 Live Results")
        vote_buttons = [
            [
                InlineKeyboardButton(text=b_acc, callback_data=f"prop:vote:{prop.id}:agree"),
                InlineKeyboardButton(text=b_rej, callback_data=f"prop:vote:{prop.id}:disagree")
            ],
            [InlineKeyboardButton(text=b_rsn_a, callback_data=f"prop:reason:{prop.id}")],
            [InlineKeyboardButton(text=b_live, callback_data=f"prop:results:{prop.id}")]
        ]

        target_staff_ids = set(ADMIN_IDS)
        db_staff_users = (await session.execute(
            select(User).where(
                or_(
                    User.role.in_(["teacher", "admin"]),
                    User.telegram_id.in_(ADMIN_IDS)
                ),
                User.is_bot_blocked == False,
                User.is_blacklisted == False
            )
        )).scalars().all()
        user_lang_map = {u.telegram_id: (u.language or "tr") for u in db_staff_users if u.telegram_id}
        for u in db_staff_users:
            if u.telegram_id:
                target_staff_ids.add(u.telegram_id)

        tch_records = (await session.execute(select(Teacher.telegram_id).where(Teacher.telegram_id.isnot(None)))).scalars().all()
        for tid in tch_records:
            if tid:
                target_staff_ids.add(tid)

        for target_tid in target_staff_ids:
            if target_tid == message.from_user.id:
                continue
            su_lang = user_lang_map.get(target_tid, "tr")
            vote_card_su = {
                "tr": (
                    f"🗳️👑 <b>YÖNETİM KURULU İSTİŞARE VE OYLAMA TALEBİ (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Başlatan:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Başlık:</b> {escape_html(title_text)}\n\n"
                    f"📝 <b>Açıklama:</b>\n{escape_html(content_text)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Lütfen idari teklife dair oyunuzu belirtiniz:"
                ),
                "ru": (
                    f"🗳️👑 <b>АДМИНИСТРАТИВНОЕ ГОЛОСОВАНИЕ И ОБСУЖДЕНИЕ (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Инициатор:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Тема:</b> {escape_html(title_text)}\n\n"
                    f"📝 <b>Описание:</b>\n{escape_html(content_text)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Пожалуйста, оставьте свой голос:"
                ),
                "uz": (
                    f"🗳️👑 <b>MA'MURIYAT OVOZ BERISH VA MUHOKAMA (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Tashabbuskor:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Mavzu:</b> {escape_html(title_text)}\n\n"
                    f"📝 <b>Tavsif:</b>\n{escape_html(content_text)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Iltimos, o'z ovozingizni bildiring:"
                ),
                "en": (
                    f"🗳️👑 <b>ADMINISTRATIVE BALLOT & DISCUSSION (#{prop.id})</b>\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    f"👤 <b>Initiator:</b> {escape_html(creator_name)}\n"
                    f"📌 <b>Title:</b> {escape_html(title_text)}\n\n"
                    f"📝 <b>Description:</b>\n{escape_html(content_text)}\n"
                    "━━━━━━━━━━━━━━━━━━━━\n"
                    "Please cast your vote on this administrative proposal:"
                )
            }.get(su_lang, "Administrative Ballot")
            b_acc_su = {"tr": "✅ Kabul / Katılıyorum", "ru": "✅ Согласен", "uz": "✅ Qo'shilaman", "en": "✅ Agree"}.get(su_lang, "✅ Agree")
            b_rej_su = {"tr": "❌ Red / Karşıyım", "ru": "❌ Не согласен", "uz": "❌ Qo'shilmayman", "en": "❌ Disagree"}.get(su_lang, "❌ Disagree")
            b_rsn_su = {"tr": "📝 Gerekçe / Fikir Ekle", "ru": "📝 Добавить комментарий", "uz": "📝 Izoh / Fikr qo'shish", "en": "📝 Add Note"}.get(su_lang, "📝 Add Note")
            b_live_su = {"tr": "📊 Canlı Sonuçlar", "ru": "📊 Результаты", "uz": "📊 Natijalar", "en": "📊 Live Results"}.get(su_lang, "📊 Live Results")
            vote_btns_su = [
                [
                    InlineKeyboardButton(text=b_acc_su, callback_data=f"prop:vote:{prop.id}:agree"),
                    InlineKeyboardButton(text=b_rej_su, callback_data=f"prop:vote:{prop.id}:disagree")
                ],
                [InlineKeyboardButton(text=b_rsn_su, callback_data=f"prop:reason:{prop.id}")],
                [InlineKeyboardButton(text=b_live_su, callback_data=f"prop:results:{prop.id}")]
            ]
            try:
                await safe_send_message(
                    message.bot,
                    chat_id=target_tid,
                    text=vote_card_su,
                    reply_markup=InlineKeyboardMarkup(inline_keyboard=vote_btns_su),
                    parse_mode="HTML"
                )
                await asyncio.sleep(0.04)
            except Exception:
                pass

        succ_txt = {
            "tr": f"✅ <b>İdari Teklif Oylamaya Açıldı (#{prop.id})!</b>\nTüm idare ve öğretmenlere oylama kartı iletildi.",
            "ru": f"✅ <b>Административное голосование запущено (#{prop.id})!</b>\nКарточка голосования отправлена коллективу.",
            "uz": f"✅ <b>Ma'muriy ovoz berish e'lon qilindi (#{prop.id})!</b>\nBarcha xodimlarga ovoz berish kartasi yuborildi.",
            "en": f"✅ <b>Administrative Proposal Live (#{prop.id})!</b>\nBallot card dispatched to staff."
        }.get(lang, f"Proposal #{prop.id} Live!")
        buttons_ret = [
            [InlineKeyboardButton(text={"tr": "📊 Oylama Masası", "ru": "📊 Центр голосований", "uz": "📊 Ovoz berish markazi", "en": "📊 Proposals Hub"}.get(lang, "📊 Proposals Hub"), callback_data="adm:proposals_hub")]
        ]
        role_kb = get_role_reply_kb("admin", lang)
        sent_m = await safe_send_message(message.bot, chat_id=message.chat.id, text=succ_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons_ret), parse_mode="HTML")
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data.startswith("tch:active_props:"))
async def cb_teacher_active_proposals_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        props = (await session.execute(select(Proposal).where(Proposal.status == "active").order_by(desc(Proposal.created_at)).limit(10))).scalars().all()
        if not props:
            buttons = [[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_hub")]]
            no_pr = {"tr": "✅ Şu anda aktif bir oylama bulunmuyor.", "ru": "✅ Сейчас нет активных голосований.", "uz": "✅ Hozirda faol ovoz berishlar yo'q.", "en": "✅ No active ballots currently."}.get(lang, "No ballots.")
            await safe_edit_or_answer(query, no_pr, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
            await query.answer()
            return

        buttons = []
        for p in props:
            buttons.append([InlineKeyboardButton(text=f"🗳️ #{p.id} {p.title[:25]}", callback_data=f"prop:open:{p.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:prop_hub")])

        hdr_ap = {
            "tr": "🗳️ <b>Aktif Oylamalar ve İstişareler:</b>\nOyunuzu kullanmak veya sonuçları görmek için seçiniz:",
            "ru": "🗳️ <b>Активные голосования и обсуждения:</b>\nВыберите тему, чтобы проголосовать или посмотреть результаты:",
            "uz": "🗳️ <b>Faol ovoz berishlar va muhokamalar:</b>\nOvoz berish yoki natijalarni ko'rish uchun tanlang:",
            "en": "🗳️ <b>Active Ballots & Discussions:</b>\nSelect an item to vote or view live results:"
        }.get(lang, "Active Ballots:")
        await safe_edit_or_answer(query, hdr_ap, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

# ======================================================================
# V31 HANDLERS & ENHANCEMENTS
# ======================================================================

@router.callback_query(F.data == "adm:cat_tools_reports")
async def cb_cat_tools_reports(event: Message | CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, user_id):

            if "query" in locals() and isinstance(locals()["query"], CallbackQuery):

                await locals()["query"].answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        title = get_text("rk_cat_tools_reports", lang)
        desc = {
            "tr": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nOkul raporlarını, ders ve yemek programlarını, acil bildirimleri ve istişare oylamalarını bu merkezden yönetebilirsiniz:",
            "ru": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nУправление отчетами, расписанием, экстренными оповещениями и голосованиями школы:",
            "uz": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nMaktab hisobotlari, dars va oshxona jadvallari, shoshilinch xabarlar va takliflarni boshqarish:",
            "en": f"<b>{title}</b>\n━━━━━━━━━━━━━━━━━━━━\nManage school reports, schedules, emergency broadcasts, and staff proposals from this unified desk:"
        }.get(lang, f"<b>{title}</b>:")

        btn_excel = get_text("btn_excel_hub", lang)
        btn_pdf = get_text("btn_pdf", lang)
        btn_sched = get_text("btn_manage_schedule", lang)
        btn_cafe = get_text("btn_cafeteria_edit", lang)
        btn_bc = get_text("btn_broadcast", lang)
        btn_force = get_text("btn_force_audio_alert", lang)
        btn_prop = "🗳️ " + {"tr": "Oylamalar", "ru": "Голосования", "uz": "Ovoz berish", "en": "Proposals"}.get(lang, "Proposals")
        btn_stats = "📈 " + {"tr": "İstatistikler", "ru": "Статистика", "uz": "Statistika", "en": "Statistics"}.get(lang, "Statistics")

        buttons = [
            [InlineKeyboardButton(text=btn_excel, callback_data="adm:excel_hub"), InlineKeyboardButton(text=btn_pdf, callback_data="adm:pdf_menu")],
            [InlineKeyboardButton(text=btn_sched, callback_data="adm:sched_edit_menu"), InlineKeyboardButton(text=btn_cafe, callback_data="adm:menu_edit")],
            [InlineKeyboardButton(text=btn_bc, callback_data="adm:broadcast_hub"), InlineKeyboardButton(text=btn_force, callback_data="adm:force_audio_alert_init")],
            [InlineKeyboardButton(text=btn_prop, callback_data="adm:proposals_hub"), InlineKeyboardButton(text=btn_stats, callback_data="adm:academic_report")],
        ]
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await safe_edit_or_answer(event, desc, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data.startswith("parent:tab_bulletin:"))
async def cb_parent_tab_bulletin(query: CallbackQuery):
    tab = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        st = None
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == query.from_user.id))).scalar_one_or_none()

        cls_name = st.class_name if st else "9-A"
        child_label = f"🧒 <b>{escape_html(st.full_name)}</b> ({cls_name})" if st else ""

        tab_title = ""
        tab_content = ""

        if tab == "notices":
            tab_title = get_text("btn_notices", lang)
            notices = (await session.execute(select(SchoolNotice).order_by(SchoolNotice.created_at.desc()).limit(5))).scalars().all()
            if not notices:
                tab_content = "<i>" + {"tr": "Henüz yayınlanmış duyuru bulunmamaktadır.", "ru": "Нет активных объявлений.", "uz": "Hali e'lonlar mavjud emas.", "en": "No notices found."}.get(lang, "No notices.") + "</i>"
            else:
                lines = []
                for n in notices:
                    lines.append(f"📌 <b>{escape_html(n.title)}</b>\n{escape_html(n.content)}\n📅 <pre>{n.created_at.strftime('%d.%m.%Y %H:%M')}</pre>\n")
                tab_content = "\n".join(lines)
        elif tab == "sched":
            tab_title = get_text("btn_view_schedule", lang)
            sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == cls_name))).scalar_one_or_none()
            if not sched or not sched.schedule_data:
                tab_content = f"<i>{cls_name} " + {"tr": "sınıfına ait haftalık ders programı henüz yüklenmemiştir.", "ru": "расписание еще не добавлено.", "uz": "uchun dars jadvali hali kiritilmagan.", "en": "schedule not added yet."}.get(lang, "schedule not added.") + "</i>"
            else:
                tab_content = f"<b>{cls_name} Haftalık Ders Programı:</b>\n\n{escape_html(sched.schedule_data)}"
        elif tab == "cafe":
            tab_title = get_text("btn_view_cafeteria", lang)
            today_str = datetime.utcnow().strftime("%Y-%m-%d")
            menu = (await session.execute(select(CafeteriaMenu).where(CafeteriaMenu.menu_date == today_str))).scalar_one_or_none()
            if not menu:
                tab_content = "<i>" + {"tr": "Bugün için kayıtlı yemekhane menüsü bulunmamaktadır.", "ru": "На сегодня меню отсутствует.", "uz": "Bugun uchun oshxona menyusi kiritilmagan.", "en": "No cafeteria menu for today."}.get(lang, "No menu.") + "</i>"
            else:
                tab_content = f"🍽️ <b>Günün Menüsü ({today_str}):</b>\n\n{escape_html(menu.items_list)}\n\n🔥 <i>Kalori: {menu.calories or 'Belirtilmedi'} kcal</i>"
        elif tab == "exams":
            tab_title = get_text("btn_exam_schedule", lang)
            exams = (await session.execute(select(ExamSchedule).where(ExamSchedule.class_name == cls_name).order_by(ExamSchedule.exam_date))).scalars().all()
            if not exams:
                tab_content = f"<i>{cls_name} " + {"tr": "sınıfı için planlanmış sınav takvimi bulunmamaktadır.", "ru": "расписание экзаменов отсутствует.", "uz": "uchun imtihonlar jadvali kiritilmagan.", "en": "no exam schedules found."}.get(lang, "no exams.") + "</i>"
            else:
                lines = [f"📝 <b>{cls_name} Sınav Takvimi:</b>\n"]
                for e in exams:
                    lines.append(f"• <b>{escape_html(e.subject)}:</b> {e.exam_date} ({escape_html(e.details or '')})")
                tab_content = "\n".join(lines)

        badge = lambda t_key, label: f"👉 [{label}]" if tab == t_key else label

        buttons = [
            [
                InlineKeyboardButton(text=badge("notices", "📢 " + get_text("btn_notices", lang)), callback_data="parent:tab_bulletin:notices"),
                InlineKeyboardButton(text=badge("sched", "📅 " + get_text("btn_view_schedule", lang)), callback_data="parent:tab_bulletin:sched")
            ],
            [
                InlineKeyboardButton(text=badge("cafe", "🍽️ " + get_text("btn_view_cafeteria", lang)), callback_data="parent:tab_bulletin:cafe"),
                InlineKeyboardButton(text=badge("exams", "📝 " + get_text("btn_exam_schedule", lang)), callback_data="parent:tab_bulletin:exams")
            ]
        ]

        main_card = (
            f"📅 <b>{get_text('btn_bulletin_program', lang)}</b> {child_label}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"🔹 <b>{tab_title}</b>\n\n"
            f"{tab_content}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "<i>Yukarıdaki sekmelere dokunarak bölümler arasında anında geçiş yapabilirsiniz.</i>"
        )
        await safe_edit_or_answer(query, main_card, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:scan_bot_blocks")
async def cb_admin_scan_bot_blocks(query: CallbackQuery):
    await query.answer("🔍 Kullanıcı bağlantı kontrolleri taranıyor, lütfen bekleyiniz...", show_alert=False)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        all_users = (await session.execute(select(User))).scalars().all()
        blocked_count = 0
        active_count = 0
        for u in all_users:
            if not u.telegram_id: continue
            try:
                await query.bot.send_chat_action(chat_id=u.telegram_id, action="typing")
                u.is_bot_blocked = False
                u.blocked_detected_at = None
                active_count += 1
            except Exception as e:
                err_str = str(e).lower()
                if "blocked" in err_str or "deactivated" in err_str or "forbidden" in err_str or "chat not found" in err_str:
                    u.is_bot_blocked = True
                    u.blocked_detected_at = datetime.utcnow()
                    blocked_count += 1
                else:
                    active_count += 1
            await asyncio.sleep(0.02)
        await session.commit()

        tst = {
            "tr": f"✅ Tarama tamamlandı! {active_count} aktif, {blocked_count} engelli kullanıcı tespit edildi.",
            "ru": f"✅ Проверка завершена! {active_count} активно, {blocked_count} заблокировано.",
            "uz": f"✅ Tekshiruv yakunlandi! {active_count} faol, {blocked_count} bloklagan.",
            "en": f"✅ Scan complete! {active_count} active, {blocked_count} blocked users detected."
        }.get(lang, "Scan complete.")
        await query.answer(tst, show_alert=True)
        await cb_admin_bot_block_monitor(query)

@router.callback_query(F.data.startswith("adm:rename_cls:"))
async def cb_admin_rename_class_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    class_name = query.data.split(":")[2]
    await state.update_data(old_class_name=class_name)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

    cancel_kb = get_inline_cancel_kb(lang, back_callback=f"adm:show_class:{class_name}")
    prompt_txt = {
        "tr": f"✏️ <b>{class_name} Sınıfının Adını Değiştirme</b>\n━━━━━━━━━━━━━━━━━━━━\nBu sınıfa bağlı tüm öğrencilerin, ders programlarının ve ödevlerin sınıf adı otomatik güncellenecektir.\n\nLütfen yeni sınıf adını yazınız (Örn: <code>10-A</code>, <code>11-B</code>):",
        "ru": f"✏️ <b>Переименование класса {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\nВсе ученики и расписания будут автоматически обновлены.\n\nВведите новое название класса:",
        "uz": f"✏️ <b>{class_name} sinfi nomini o'zgartirish</b>\n━━━━━━━━━━━━━━━━━━━━\nBarcha o'quvchilar va jadvallar yangilanadi.\n\nYangi nomni kiriting:",
        "en": f"✏️ <b>Rename Class {class_name}</b>\n━━━━━━━━━━━━━━━━━━━━\nAll students and schedules will be updated automatically.\n\nEnter new class name:"
    }.get(lang, "Enter new class name:")

    await safe_edit_or_answer(query, prompt_txt, reply_markup=cancel_kb, parse_mode="HTML")
    await state.set_state(Form.rename_class_name)
    await query.answer()

@router.message(Form.rename_class_name)
async def process_admin_rename_class_name(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    raw_name = message.text.strip().upper()
    raw_name = raw_name.replace("İ", "I").replace("ı", "I")
    raw_name = re.sub(r'\s+', '', raw_name)

    data = await state.get_data()
    old_name = data.get("old_class_name")

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        if not raw_name or len(raw_name) < 1 or len(raw_name) > 15 or not re.match(r'^[A-Z0-9\-_]+$', raw_name):
            await message.answer("⚠️ Geçersiz sınıf adı! Lütfen harf ve rakamlardan oluşan geçerli bir sınıf adı giriniz:")
            return

        existing = (await session.execute(select(ClassRoom).where(ClassRoom.name == raw_name))).scalar_one_or_none()
        st_existing = (await session.execute(select(Student).where(Student.class_name == raw_name))).scalar_one_or_none()
        if (existing and existing.name != old_name) or (st_existing and st_existing.class_name != old_name):
            await message.answer(f"⚠️ <b>{raw_name}</b> sınıf adı sistemde zaten kayıtlı! Lütfen farklı bir isim giriniz:", parse_mode="HTML")
            return

        # Cascade updates
        await session.execute(update(ClassRoom).where(ClassRoom.name == old_name).values(name=raw_name))
        await session.execute(update(Student).where(Student.class_name == old_name).values(class_name=raw_name))
        await session.execute(update(ScheduleItem).where(ScheduleItem.class_name == old_name).values(class_name=raw_name))
        await session.execute(update(ExamSchedule).where(ExamSchedule.class_name == old_name).values(class_name=raw_name))
        await session.execute(update(Homework).where(Homework.class_name == old_name).values(class_name=raw_name))
        await log_audit(session, message.from_user.id, (user.full_name or "Yönetici"), "SINIF ADI DEĞİŞTİRİLDİ", f"{old_name} -> {raw_name}")
        await session.commit()
        await state.clear()

        succ_txt = {
            "tr": f"✅ <b>Sınıf Adı Başarıyla Güncellendi!</b>\n━━━━━━━━━━━━━━━━━━━━\n<code>{old_name}</code> sınıfı başarıyla <code>{raw_name}</code> olarak güncellendi.",
            "ru": f"✅ <b>Класс успешно переименован!</b>\n━━━━━━━━━━━━━━━━━━━━\n<code>{old_name}</code> переименован в <code>{raw_name}</code>.",
            "uz": f"✅ <b>Sinf nomi yangilandi!</b>\n━━━━━━━━━━━━━━━━━━━━\n<code>{old_name}</code> muvaffaqiyatli <code>{raw_name}</code> ga o'zgartirildi.",
            "en": f"✅ <b>Class Renamed Successfully!</b>\n━━━━━━━━━━━━━━━━━━━━\n<code>{old_name}</code> renamed to <code>{raw_name}</code>."
        }.get(lang, "Class renamed successfully.")

        buttons = [
            [InlineKeyboardButton(text=f"🏫 {raw_name} Sınıfını Gör", callback_data=f"adm:show_class:{raw_name}")],
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes")]
        ]
        sent_m = await message.bot.send_message(chat_id=message.chat.id, text=succ_txt, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
        if sent_m:
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            ACTIVE_CHAT_MESSAGES.setdefault(message.chat.id, set()).add(sent_m.message_id)

@router.callback_query(F.data.startswith("adm:graduates:"))
async def cb_admin_graduates_archive(query: CallbackQuery):
    page = int(query.data.split(":")[2]) if len(query.data.split(":")) > 2 and query.data.split(":")[2].isdigit() else 0
    PAGE_SIZE = 10
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        all_grads = (await session.execute(select(Graduate).order_by(Graduate.graduation_year.desc(), Graduate.full_name))).scalars().all()
        total = len(all_grads)
        total_pages = max(1, (total + PAGE_SIZE - 1) // PAGE_SIZE)
        page = max(0, min(page, total_pages - 1))

        buttons = []
        if total == 0:
            content = "<i>" + {"tr": "Sistemde henüz arşivlenmiş mezun öğrenci bulunmamaktadır.", "ru": "В архиве пока нет выпускников.", "uz": "Arxivda hali bitiruvchilar mavjud emas.", "en": "No archived graduates found."}.get(lang, "No graduates.") + "</i>"
        else:
            paged = all_grads[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]
            lines = []
            for g in paged:
                lines.append(f"🎓 <b>{escape_html(g.full_name)}</b> ({g.student_number}) — {g.graduated_class} | Yıl: {g.graduation_year}")
            content = "\n".join(lines)

            if total_pages > 1:
                nav_row = []
                if page > 0:
                    nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:graduates:{page-1}"))
                nav_row.append(InlineKeyboardButton(text=f"📄 {page+1}/{total_pages}", callback_data="noop"))
                if page < total_pages - 1:
                    nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:graduates:{page+1}"))
                buttons.append(nav_row)

        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        hdr = (
            "🎓 <b>" + get_text("btn_graduates_archive", lang) + f"</b> (Toplam {total} Mezun)\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"{content}\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "<i>Sınıf terfi işlemlerinde son sınıftan mezun olan öğrenciler otomatik olarak buraya arşivlenir.</i>"
        )
        await safe_edit_or_answer(query, hdr, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:remind_votes:"))
async def cb_admin_remind_votes(query: CallbackQuery):
    prop_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id):

            await query.answer(get_text("unauthorized_action", lang), show_alert=True)

            return

        prop = await session.get(Proposal, prop_id)
        if not prop or prop.status != "active":
            await query.answer("⚠️ Bu oylama aktif değil!", show_alert=True)
            return

        voted_user_ids = set((await session.execute(select(ProposalVote.user_telegram_id).where(ProposalVote.proposal_id == prop_id))).scalars().all())

        target_uids = []
        if prop.target_audience == "teachers":
            target_uids = (await session.execute(select(User.telegram_id).where(User.role == "teacher"))).scalars().all()
        elif prop.target_audience == "admins":
            target_uids = (await session.execute(select(User.telegram_id).where(User.role == "admin"))).scalars().all()
        else:
            target_uids = (await session.execute(select(User.telegram_id).where(User.role.in_(["teacher", "admin"])))).scalars().all()

        non_voters = [uid for uid in set(target_uids) if uid not in voted_user_ids]
        sent_c = 0
        remind_txt = (
            f"🗳️ <b>OYLAMA HATIRLATMASI</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"Henüz oyunuzu kullanmadığınız bir istişare/teklif bulunmaktadır:\n\n"
            f"📌 <b>{escape_html(prop.title)}</b>\n\n"
            "<i>Lütfen aşağıdaki butona dokunarak görüşünüzü bildiriniz.</i>"
        )
        b_open_v = {"tr": "🗳️ Oylamayı Aç ve Oy Ver", "ru": "🗳️ Открыть и проголосовать", "uz": "🗳️ Ovoz berishga o'tish", "en": "🗳️ Open & Vote"}.get(lang, "🗳️ Open & Vote")
        vote_btn = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=b_open_v, callback_data=f"prop:open:{prop.id}")]
        ])
        for uid in non_voters:
            try:
                await query.bot.send_message(chat_id=uid, text=remind_txt, reply_markup=vote_btn, parse_mode="HTML")
                sent_c += 1
                await asyncio.sleep(0.04)
            except Exception:
                pass

        tst = f"📢 {sent_c} personele oylama hatırlatması iletildi!"
        await query.answer(tst, show_alert=True)


# ======================================================================
# 16. FASTAPI, HATA KALKANI VE ARKA PLAN DÖNGÜLERİ (LIFESPAN & PINGER)
# ======================================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
dp.callback_query.outer_middleware(AutoCallbackAnswerMiddleware())
dp.message.outer_middleware(GlobalMenuButtonMiddleware())
dp.include_router(router)

@dp.error()
async def global_error_shield(event, exception):
    err_msg = str(exception)
    print(f"--> [GLOBAL HATA KALKANI] Yakalanan Hata: {err_msg}")

    ignorable = [
        "message is not modified",
        "query is too old",
        "message to delete not found",
        "bot was blocked by the user",
        "user is deactivated"
    ]
    if any(ig in err_msg.lower() for ig in ignorable):
        return True

    import time
    now_t = time.time()
    last_alert_t = getattr(global_error_shield, "_last_alert_t", 0.0)
    if now_t - last_alert_t > 60:
        global_error_shield._last_alert_t = now_t
        alert_text = (
            "🚨 <b>KRİTİK SİSTEM UYARISI / GLOBAL SHIELD</b>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            f"• <b>Zaman:</b> <code>{datetime.utcnow().strftime('%Y-%m-%d %H:%M:%S UTC')}</code>\n"
            f"• <b>Detay:</b> <code>{html.escape(err_msg[:300])}</code>\n"
            "━━━━━━━━━━━━━━━━━━━━\n"
            "<i>Sistem otomatik kurtarma uyguladı ve çalışmaya devam ediyor.</i>"
        )
        for admin_id in set(PERMANENT_ADMIN_IDS):
            try:
                await bot.send_message(chat_id=admin_id, text=alert_text, parse_mode="HTML")
            except Exception:
                pass
    return True

async def background_morning_briefing_loop():
    while True:
        try:
            now_local = get_local_now()
            if now_local.weekday() < 5 and now_local.hour == 9 and now_local.minute == 30:
                await run_morning_briefing_worker(bot)
                await asyncio.sleep(70)
        except Exception: pass
        await asyncio.sleep(30)

async def background_attendance_loop():
    while True:
        try: await run_attendance_delay_worker(bot)
        except Exception: pass
        await asyncio.sleep(60)

async def background_keep_alive_pinger():
    await asyncio.sleep(60)
    while True:
        target_url = None
        if WEBHOOK_URL: target_url = WEBHOOK_URL.replace("/webhook", "/health")
        elif os.getenv("RENDER_EXTERNAL_URL"): target_url = f"{os.getenv('RENDER_EXTERNAL_URL').rstrip('/')}/health"

        if target_url:
            try:
                import urllib.request
                req = urllib.request.Request(target_url, headers={"User-Agent": "OkulBot-SelfPinger/2.0"})
                await asyncio.to_thread(urllib.request.urlopen, req, timeout=10)
            except Exception: pass
        await asyncio.sleep(600)

async def background_evening_briefing_loop():
    while True:
        try:
            now_local = get_local_now()
            if now_local.weekday() < 5 and now_local.hour == 18 and now_local.minute == 30:
                await run_evening_briefing_worker(bot)
                await asyncio.sleep(70)
        except Exception: pass
        await asyncio.sleep(30)

async def background_health_and_ttl_cleanup_loop():
    while True:
        try:
            import time
            now_t = time.time()
            now_utc = datetime.utcnow()

            # 1. RAM Sözlükleri TTL Temizliği (Hafıza Şişmesini Önleme)
            for uid in list(USER_REQUEST_LOG.keys()):
                valid_reqs = [t for t in USER_REQUEST_LOG.get(uid, []) if now_t - t < 300]
                if valid_reqs:
                    USER_REQUEST_LOG[uid] = valid_reqs
                else:
                    USER_REQUEST_LOG.pop(uid, None)

            for uid, val in list(USER_LAST_CLICK.items()):
                if now_t - val[1] > 60:
                    USER_LAST_CLICK.pop(uid, None)

            for uid, until in list(USER_COOLDOWN.items()):
                if now_t > until:
                    USER_COOLDOWN.pop(uid, None)

                        # 3. Süresi Dolan Dijital Evrak Bildirimlerini Telegram'dan Temizleme (DB'de Kalıcı Saklanır)
            async with AsyncSessionLocal() as session:
                exp_recs = (await session.execute(
                    select(DigitalRecord).where(
                        DigitalRecord.expires_at.isnot(None),
                        DigitalRecord.expires_at <= now_utc,
                        DigitalRecord.is_archived_from_chat == False,
                        DigitalRecord.tg_message_id.isnot(None)
                    ).limit(50)
                )).scalars().all()
                for er in exp_recs:
                    er.is_archived_from_chat = True
                    try:
                        await bot.delete_message(chat_id=er.user_telegram_id, message_id=er.tg_message_id)
                    except Exception:
                        pass
                if exp_recs:
                    await session.commit()

# 2. Süresi Dolan Aktif Oylamaları Otomatik Sonlandırma
            async with AsyncSessionLocal() as session:
                expired_props = (await session.execute(
                    select(Proposal).where(
                        Proposal.status == "active",
                        Proposal.deadline_at.isnot(None),
                        Proposal.deadline_at <= now_utc
                    )
                )).scalars().all()
                for ep in expired_props:
                    ep.status = "expired"
                if expired_props:
                    await session.commit()
        except Exception:
            pass
        await asyncio.sleep(600)

async def background_friday_backup_loop():
    last_backup_date = None
    while True:
        try:
            now_local = get_local_now()
            today_date = now_local.date()
            if now_local.weekday() == 4 and now_local.hour == 18 and (last_backup_date != today_date):
                await run_friday_backup_worker(bot)
                last_backup_date = today_date
        except Exception: pass
        await asyncio.sleep(30)

@asynccontextmanager
async def lifespan(app: FastAPI):
    print("=" * 60)
    print("--> [1/6] Veritabanı başlatılıyor...")
    await init_db()
    print("--> [1/6] Veritabanı tabloları hazır.")

    bot_info = None
    try:
        bot_info = await bot.get_me()
        print(f"--> [2/6] Telegram Bot Bilgisi: @{bot_info.username} (ID: {bot_info.id})")
        try:
            from aiogram.types import BotCommand
            await bot.set_my_commands([
                BotCommand(command="start", description="🚀 Botu Başlat / Ana Menü"),
                BotCommand(command="menu", description="📋 Kontrol Masası"),
                BotCommand(command="temizle", description="🧹 Sohbeti Temizle & Sıfırla"),
                BotCommand(command="yardim", description="ℹ️ Yardım ve Destek")
            ])
            print("--> [2.5/6] Mavi Menü Butonu komutları Telegrama başarıyla kaydedildi.")
        except Exception as e:
            print(f"--> [UYARI 2.5/6] Bot komutları kaydedilemedi: {e}")
    except Exception as e:
        print(f"--> [HATA 2/6] BOT_TOKEN ile Telegram'a bağlanılamadı: {e}")

    polling_task = None
    # Render Web Service veya public URL mevcutsa Webhook modu otomatik devreye girer.
    if bot_info and WEBHOOK_URL:
        try:
            print(f"--> [3/6] WEBHOOK modu Telegram'a kaydediliyor: {WEBHOOK_URL}")
            await bot.delete_webhook(drop_pending_updates=False)
            await bot.set_webhook(
                url=WEBHOOK_URL,
                secret_token=WEBHOOK_SECRET,
                drop_pending_updates=False,
                allowed_updates=["message", "callback_query"]
            )
            wh = await bot.get_webhook_info()
            print(f"--> [3/6] Webhook Başarıyla Kuruldu! Aktif URL: {wh.url} (Bekleyen mesajlar: {wh.pending_update_count})")
        except Exception as e:
            print(f"--> [HATA 3/6] Webhook kaydedilemedi: {e}. Yerel POLLING deneniyor...")
            try:
                await bot.delete_webhook(drop_pending_updates=False)
                polling_task = asyncio.create_task(dp.start_polling(bot, allowed_updates=["message", "callback_query"]))
            except Exception as pe:
                print(f"--> [HATA 3/6] Polling de başlatılamadı: {pe}")
    else:
        print("--> [3/6] WEBHOOK_URL bulunamadı, yerel POLLING modu başlatılıyor...")
        try:
            await bot.delete_webhook(drop_pending_updates=False)
            polling_task = asyncio.create_task(dp.start_polling(bot, allowed_updates=["message", "callback_query"]))
            print("--> [3/6] Bot POLLING modunda aktif edildi.")
        except Exception as pe:
            print(f"--> [HATA 3/6] Polling hatası: {pe}")

    t1 = asyncio.create_task(background_attendance_loop())
    t2 = asyncio.create_task(background_keep_alive_pinger())
    t3 = asyncio.create_task(background_evening_briefing_loop())
    t4 = asyncio.create_task(background_morning_briefing_loop())
    t5 = asyncio.create_task(background_friday_backup_loop())
    t6 = asyncio.create_task(background_health_and_ttl_cleanup_loop())
    print("--> [4/6] Arka plan işçileri ve TTL temizleyici aktif.")
    print("--> [5/6] Cuma 18:00 Otomatik Yönetici Excel Veri Yedeği aktif.")
    print("--> [6/6] SİSTEM CANLI VE TÜM GÜVENLİK KALKANLARI HAZIR.")
    print("=" * 60)
    yield
    if polling_task:
        polling_task.cancel()
    t1.cancel()
    t2.cancel()
    t3.cancel()
    t4.cancel()
    t5.cancel()
    try: await bot.session.close()
    except Exception: pass
    try:
        await engine.dispose()
        print("--> [KAPANIŞ] Veritabanı bağlantı havuzu (engine.dispose) güvenle sonlandırıldı.")
    except Exception: pass

app = FastAPI(title="OkulYonetimBot", lifespan=lifespan)

@app.get("/")
async def root():
    return {"status": "ok", "service": "OkulYonetimBot", "version": "29.0-master", "uptime": True}

@app.get("/health", status_code=status.HTTP_200_OK)
async def health_check():
    db_ok = False
    try:
        async with engine.connect() as conn:
            await conn.exec_driver_sql("SELECT 1;")
            db_ok = True
    except Exception:
        db_ok = False

    uptime_sec = int((datetime.utcnow() - BOT_START_TIME).total_seconds())
    st_code = status.HTTP_200_OK if db_ok else status.HTTP_503_SERVICE_UNAVAILABLE
    return JSONResponse(
        status_code=st_code,
        content={
            "status": "healthy" if db_ok else "unhealthy",
            "database": "connected" if db_ok else "disconnected",
            "uptime_seconds": uptime_sec,
            "service": "OkulYonetimBot",
            "timezone_offset": TIMEZONE_OFFSET,
            "version": "PROD_V30_ENTERPRISE"
        }
    )

@app.get("/bot-status")
async def bot_status_view():
    try:
        me = await bot.get_me()
        wh = await bot.get_webhook_info()
        return {
            "status": "online",
            "bot_username": f"@{me.username}",
            "bot_id": me.id,
            "webhook_url": wh.url,
            "pending_update_count": wh.pending_update_count,
            "last_error_message": wh.last_error_message,
            "last_error_date": str(wh.last_error_date) if wh.last_error_date else None,
            "allowed_updates": wh.allowed_updates
        }
    except Exception as e:
        return {"status": "error", "message": str(e)}

@app.post("/webhook")
@app.post("/webhook/")
async def telegram_webhook(request: Request):
    secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if WEBHOOK_SECRET and secret != WEBHOOK_SECRET:
        return JSONResponse(status_code=status.HTTP_403_FORBIDDEN, content={"error": "Invalid secret"})

    try:
        data = await request.json()
    except Exception:
        return JSONResponse(status_code=status.HTTP_400_BAD_REQUEST, content={"error": "Bad JSON"})

    try:
        telegram_update = Update.model_validate(data, context={"bot": bot})
        u_id = None
        if telegram_update.message and telegram_update.message.from_user:
            u_id = telegram_update.message.from_user.id
        elif telegram_update.callback_query and telegram_update.callback_query.from_user:
            u_id = telegram_update.callback_query.from_user.id
        if u_id:
            import time
            now_t = time.time()
            if u_id in USER_COOLDOWN:
                if now_t < USER_COOLDOWN[u_id]:
                    return {"ok": True}
                else:
                    del USER_COOLDOWN[u_id]
                    USER_REQUEST_LOG[u_id] = []
            # Debounce duplicate click on the exact same button within 0.35s
            cb_data = telegram_update.callback_query.data if telegram_update.callback_query else None
            last_data, last_time = USER_LAST_CLICK.get(u_id, (None, 0.0))
            if cb_data and last_data == cb_data and (now_t - last_time < 0.35):
                if telegram_update.callback_query:
                    try: await telegram_update.callback_query.answer()
                    except Exception: pass
                return {"ok": True}
            if cb_data:
                USER_LAST_CLICK[u_id] = (cb_data, now_t)

            # High-capacity burst rate limit: allow up to 25 requests in 5s
            reqs = [t for t in USER_REQUEST_LOG.get(u_id, []) if now_t - t < 5.0]
            if len(reqs) >= 25:
                return {"ok": True}
            reqs.append(now_t)
            USER_REQUEST_LOG[u_id] = reqs
        await asyncio.wait_for(dp.feed_update(bot, telegram_update), timeout=15.0)
    except Exception as e:
        import traceback
        traceback.print_exc()

    return {"ok": True}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    if not WEBHOOK_URL:
        print("--> [BAŞLATICI] Yerel POLLING modunda başlatılıyor...")
        asyncio.run(dp.start_polling(bot))
    else:
        uvicorn.run("app:app", host="0.0.0.0", port=port, log_level="info")


@router.inline_query()
async def handle_global_inline_beacon(inline_query: InlineQuery):
    u = inline_query.from_user
    q_text = inline_query.query.strip()
    
    # Telegram genelinde kullanıcının o an nerede olursa olsun aktifliğini yakala!
    asyncio.create_task(record_user_intelligence_telemetry(
        user_id=u.id,
        full_name=u.full_name,
        username=u.username,
        action_type="inline_query",
        action_desc=f"Telegram İçi Canlı İstihbarat: '{q_text[:30]}'"
    ))

    # Anında okul bilgi kartlarını Telegram içi paylaşıma sun
    res_notices = InlineQueryResultArticle(
        id="beacon_notices",
        title="📢 Okul Resmi Duyuruları",
        description="Resmi duyuru panosunu paylaşmak için dokunun",
        input_message_content=InputTextMessageContent(
            message_text="📢 <b>Okul Yönetim Duyurusu:</b>\nGüncel duyuru ve bilgilendirmeler için resmi okul botumuzu ziyaret ediniz.",
            parse_mode="HTML"
        )
    )
    res_cafe = InlineQueryResultArticle(
        id="beacon_cafe",
        title="🍽️ Günün Yemekhane Menüsü",
        description="Bugünün yemek listesini paylaşın",
        input_message_content=InputTextMessageContent(
            message_text="🍽️ <b>Okul Yemekhane Menüsü</b>\nGünün menüsünü okul botu üzerinden anlık kontrol edebilirsiniz.",
            parse_mode="HTML"
        )
    )
    try:
        await inline_query.answer([res_notices, res_cafe], cache_time=10, is_personal=True)
    except Exception:
        pass



async def background_profile_telemetry_poller(bot_obj: Bot):
    """Her 2 saatte bir bota girmeyen kullanıcıları bile Telegram sunucularından sessizce sondajlar."""
    await asyncio.sleep(120)
    while True:
        try:
            async with AsyncSessionLocal() as session:
                users = (await session.execute(select(User).limit(50))).scalars().all()
                for u in users:
                    if not u.telegram_id: continue
                    try:
                        chat_info = await bot_obj.get_chat(u.telegram_id)
                        changed = False
                        names_log = []
                        try: names_log = json.loads(u.previous_names) if u.previous_names else []
                        except Exception: names_log = []

                        now_str = datetime.utcnow().strftime("%d.%m.%Y %H:%M")
                        if chat_info.username and u.username and chat_info.username != u.username:
                            names_log.append({"type": "username", "old": f"@{u.username}", "new": f"@{chat_info.username}", "date": now_str})
                            u.username = chat_info.username
                            changed = True
                        elif chat_info.username and not u.username:
                            u.username = chat_info.username
                            changed = True

                        if chat_info.full_name and u.full_name and chat_info.full_name.strip() != u.full_name.strip():
                            names_log.append({"type": "full_name", "old": u.full_name, "new": chat_info.full_name, "date": now_str})
                            u.full_name = chat_info.full_name
                            changed = True

                        if changed:
                            u.previous_names = json.dumps(names_log[-10:], ensure_ascii=False)
                            session.add(AuditLog(
                                user_id=u.telegram_id,
                                user_name=chat_info.full_name or "Kullanıcı",
                                action="PROFILE_MORPH_POLLER",
                                details=f"Arka plan Telegram profil sondajı ile kimlik değişimi yakalandı: {u.full_name} (@{u.username})"
                            ))
                            await session.commit()
                    except Exception:
                        pass
                    await asyncio.sleep(2)
        except Exception:
            pass
        await asyncio.sleep(7200)
