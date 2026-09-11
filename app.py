# -*- coding: utf-8 -*-
"""
OKUL YÖNETİM TELEGRAM BOTU - KURUMSAL VE EKSİKSİZ TAM SİSTEM MİMARİSİ (PROD V30 - COMPACT BULLETPROOF)
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
from datetime import datetime, timedelta, date
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, BackgroundTasks, status
from fastapi.responses import JSONResponse
from aiogram import Bot, Dispatcher, Router, F
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    Update, Message, CallbackQuery, BufferedInputFile,
    InlineKeyboardMarkup, InlineKeyboardButton, ReplyKeyboardMarkup, KeyboardButton, ReplyKeyboardRemove
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup, any_state

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

PERMANENT_ADMIN_IDS = [8576061834, 2146753102, 1885043735]
ADMIN_CODE = os.getenv("ADMIN_CODE", "ADM-OKUL-2026").strip()
TIMEZONE_OFFSET = int(os.getenv("TIMEZONE_OFFSET", "3"))

ADMIN_IDS = list(PERMANENT_ADMIN_IDS)
for x in os.getenv("ADMIN_IDS", "").split(","):
    clean_x = x.strip().replace("@", "")
    if clean_x.isdigit():
        val = int(clean_x)
        if val not in ADMIN_IDS:
            ADMIN_IDS.append(val)

def get_local_now() -> datetime:
    return datetime.utcnow() + timedelta(hours=TIMEZONE_OFFSET)

def get_local_date() -> date:
    return get_local_now().date()

def escape_md(text: str | None) -> str:
    if not text: return ""
    s = str(text)
    s = s.replace(chr(92), chr(92) + chr(92))
    for ch in ("_", "*", "`", "[", "]"):
        s = s.replace(ch, chr(92) + ch)
    return s

def escape_html(text: str | None) -> str:
    if not text: return ""
    return html.escape(str(text))

def clean_unicode_text(text: str | None) -> str:
    if not text: return ""
    return str(text).strip()

def normalize_code(code_str: str) -> str:
    if not code_str: return ""
    cleaned = str(code_str).strip()
    cleaned = cleaned.replace("–", "-").replace("—", "-").replace("−", "-").replace("‐", "-").replace("‑", "-")
    cleaned = cleaned.replace("_", "-")
    cleaned = cleaned.replace("ı", "I").replace("i", "I").replace("İ", "I")
    cleaned = cleaned.upper()
    cleaned = re.sub(r'[\s\-]+', '-', cleaned)
    m = re.match(r'^(VELI|OGR|HCA|ADM)(\d{4,8})$', cleaned)
    if m:
        cleaned = f"{m.group(1)}-{m.group(2)}"
    return cleaned

async def safe_edit_or_answer(target: Message | CallbackQuery, text: str, reply_markup=None, parse_mode="Markdown"):
    if isinstance(target, CallbackQuery):
        msg = target.message
        if msg and msg.from_user and msg.from_user.is_bot and not msg.photo:
            try:
                await msg.edit_text(text, reply_markup=reply_markup, parse_mode=parse_mode)
                return
            except Exception:
                try:
                    await msg.edit_text(text, reply_markup=reply_markup, parse_mode=None)
                    return
                except Exception:
                    pass
        try:
            s_m = await msg.answer(text, reply_markup=reply_markup, parse_mode=parse_mode)
            return s_m
        except Exception:
            return await msg.answer(text, reply_markup=reply_markup, parse_mode=None)
    elif isinstance(target, Message):
        try:
            return await target.answer(text, reply_markup=reply_markup, parse_mode=parse_mode)
        except Exception:
            return await target.answer(text, reply_markup=reply_markup, parse_mode=None)

async def safe_send_message(bot: Bot, chat_id: int, text: str, reply_markup=None, parse_mode="Markdown"):
    try:
        return await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup, parse_mode=parse_mode)
    except Exception:
        try:
            return await bot.send_message(chat_id=chat_id, text=text, reply_markup=reply_markup, parse_mode=None)
        except Exception:
            return None

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
    is_bot_blocked: Mapped[bool] = mapped_column(Boolean, default=False)
    blocked_bot_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    current_child_id: Mapped[int] = mapped_column(Integer, nullable=True)
    evening_briefing: Mapped[bool] = mapped_column(Boolean, default=True)
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
    behavior_type: Mapped[str] = mapped_column(String(20))
    badge: Mapped[str] = mapped_column(String(10))
    title: Mapped[str] = mapped_column(String(100))
    note: Mapped[str] = mapped_column(Text, nullable=True)
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

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
    gender: Mapped[str] = mapped_column(String(10), nullable=True)
    birth_date: Mapped[str] = mapped_column(String(20), nullable=True)
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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class CafeteriaMenu(Base):
    __tablename__ = "cafeteria_menus"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    date: Mapped[date] = mapped_column(Date, default=lambda: datetime.utcnow().date(), index=True)
    menu_text: Mapped[str] = mapped_column(Text)

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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class EmergencyAck(Base):
    __tablename__ = "emergency_acks"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    alert_id: Mapped[int] = mapped_column(Integer, ForeignKey("emergency_alerts.id"), index=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

async def init_db():
    async with engine.begin() as conn:
        if "sqlite" in DATABASE_URL:
            try:
                await conn.exec_driver_sql("PRAGMA journal_mode=WAL;")
                await conn.exec_driver_sql("PRAGMA synchronous=NORMAL;")
                await conn.exec_driver_sql("PRAGMA busy_timeout=30000;")
                await conn.exec_driver_sql("PRAGMA cache_size=-64000;")
            except Exception: pass
        await conn.run_sync(Base.metadata.create_all)
        try:
            if "sqlite" in DATABASE_URL:
                res = await conn.exec_driver_sql("PRAGMA table_info(users);")
                cols = [r[1] for r in res.fetchall()]
                for c_name, c_type in [("username", "VARCHAR(100)"), ("phone", "VARCHAR(30)"), ("admin_type", "VARCHAR(20) DEFAULT 'none'"), ("admin_until", "DATETIME"), ("previous_role", "VARCHAR(20) DEFAULT 'guest'"), ("is_bot_blocked", "BOOLEAN DEFAULT 0"), ("blocked_bot_at", "DATETIME")]:
                    if c_name not in cols: await conn.exec_driver_sql(f"ALTER TABLE users ADD COLUMN {c_name} {c_type};")
            else:
                for c_name, c_type in [("username", "VARCHAR(100)"), ("phone", "VARCHAR(30)"), ("admin_type", "VARCHAR(20) DEFAULT 'none'"), ("admin_until", "TIMESTAMP"), ("previous_role", "VARCHAR(20) DEFAULT 'guest'"), ("is_bot_blocked", "BOOLEAN DEFAULT FALSE"), ("blocked_bot_at", "TIMESTAMP")]:
                    await conn.exec_driver_sql(f"ALTER TABLE users ADD COLUMN IF NOT EXISTS {c_name} {c_type};")
        except Exception: pass

# ======================================================================
# 3. KUSURSUZ 4 DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {
    "tr": {
        "lang_select": "🌍 Lütfen bir dil seçiniz / Tilni tanlang / Пожалуйста, выберите язык / Please select language:",
        "lang_changed": "Dil başarıyla güncellendi: 🇹🇷 Türkçe",
        "welcome_guest": "🎓 *Okul Yönetim Sistemine Hoş Geldiniz.*\n\nLütfen sisteme giriş yapmak için size verilen **erişim kodunu** (Örn: `VELI-123456`, `HCA-123456`, `OGR-123456`) yazınız veya işlem seçiniz:",
        "btn_login_prompt": "🔑 Sisteme Giriş Yap (Kod Gir)",
        "prompt_enter_code_direct": "🔑 *Lütfen size verilen giriş kodunu yazınız:* (Örn: `HCA-123456`, `VELI-123456`, `OGR-123456`)",
        "auth_success": "✅ *Giriş Başarılı!*\nHoş geldiniz: *{name}*\nRolünüz: *{role}*",
        "auth_failed": "❌ Geçersiz giriş kodu! Kalan deneme hakkınız: {remaining}",
        "auth_locked": "⛔ Güvenlik nedeniyle hesabınız 1 saat süreyle kilitlendi.",
        "auth_blacklisted": "🚫 Hesabınız güvenlik nedeniyle kalıcı olarak askıya alındı.",
        "auth_code_already_linked": "⚠️ *Bu kod zaten başka bir Telegram hesabına bağlanmıştır.*\n\nKodunuzun çalındığını düşünüyorsanız lütfen derhal okul idaresine başvurunuz.",
        "maintenance_mode": "⚠️ Sistem şu anda planlı bakım modundadır. Lütfen daha sonra tekrar deneyiniz.",
        "lock_countdown_msg": "⛔ *Güvenlik Kilidi:* Hesabınız geçici olarak kilitlidir.\n\nKalan süre: *{mins} dakika*.",
        "admin_title": "⚡ *Okul Yönetim Kokpiti (Admin)*",
        "admin_stats": "📊 *Genel Durum:*\n• Sınıflar: *{c_cnt}* | Öğrenciler: *{s_cnt}* | Öğretmenler: *{t_cnt}*\n• Bekleyen Başvurular: *{req_cnt}* | Mazeretler: *{med_cnt}*\n• Tarih: *{date}*",
        "rk_cat_staff": "👥 Kadro & Öğrenci",
        "rk_cat_reports": "📊 Raporlar & Denetim",
        "rk_cat_requests": "🛎️ Onay Masası",
        "rk_cat_tools": "🛠️ İdari Araçlar",
        "rk_cat_settings": "⚙️ Sistem & Ayarlar",
        "rk_restart": "🔄 Yeniden Başlat",
        "rk_main_menu": "🏠 Ana Menü",
        "rk_admin_dash": "🏠 Ana Menü",
        "rk_lang": "🌐 Dil Değiştir",
        "rk_enter_code": "🔑 Kod Gir",
        "rk_req_access": "📝 Şifre Talep Et",
        "btn_classes": "🏫 Sınıflar & Öğrenciler",
        "btn_teachers": "👨‍🏫 Öğretmenler",
        "btn_search_student": "🔍 Öğrenci Ara",
        "btn_search_teacher": "🔍 Öğretmen Ara",
        "btn_school_admins": "👑 Okul Yöneticileri",
        "btn_users_hub": "👥 Kullanıcı Rehberi",
        "btn_cockpit_unified": "📊 Sabah Kokpiti & Yoklama",
        "btn_risk_radar": "⚠️ Riskli Öğrenci Radarı",
        "btn_academic_report": "📈 Başarı Analizi",
        "btn_unack_notifs": "⚠️ Okunmamış Devamsızlıklar",
        "btn_requests": "🛎️ Başvurular ({count})",
        "btn_medical": "🏥 Mazeretler ({count})",
        "btn_audit_logs": "📜 İşlem Geçmişi (Audit Log)",
        "btn_broadcast": "📢 Toplu Duyuru",
        "btn_excel_hub": "📥 Excel Merkezi",
        "btn_manage_schedule": "📅 Ders Programı Yönetimi",
        "btn_cafeteria_edit": "🍲 Günün Menüsünü Güncelle",
        "btn_pdf": "📄 Şifre Kartları (PDF)",
        "btn_maintenance_toggle": "🚨 Bakım Modu ({status})",
        "btn_blacklist": "🚫 Engelli Kullanıcılar",
        "btn_lang": "🌐 Dil Değiştir",
        "btn_add_student": "➕ Öğrenci Ekle",
        "btn_add_teacher": "➕ Öğretmen Ekle",
        "btn_excel": "📥 Excel ile Yükle",
        "btn_cockpit": "📊 Sabah Kokpiti",
        "btn_back": "⬅️ Geri",
        "btn_main_menu": "🏠 Ana Menü",
        "btn_prev": "⬅️ Önceki",
        "btn_next": "Sonraki ➡️",
        "btn_acknowledged": "✅ Okudum / Bilgilendirildim",
        "acknowledged_toast": "Onayınız kaydedildi.",
        "no_classes_found": "⚠️ Henüz kayıtlı bir sınıf bulunmamaktadır.",
        "prompt_student_name": "👤 Öğrencinin Adını ve Soyadını yazınız:",
        "prompt_student_class": "🏫 Öğrencinin Sınıfını yazınız (Örn: `9-A`):",
        "prompt_student_no": "🔢 Öğrencinin Okul Numarasını yazınız (Örn: `101`):",
        "student_added_card": "✅ *Öğrenci Başarıyla Kaydedildi!*\n\n👤 İsim: *{name}*\n🏫 Sınıf: *{class_name}* | No: *{no}*\n\n🔑 *Giriş Şifreleri:*\n• Öğrenci Kodu: `{st_code}`\n• Veli Kodu: `{pr_code}`",
        "prompt_teacher_name": "👨‍🏫 Öğretmenin Adını ve Soyadını yazınız:",
        "prompt_teacher_subject": "📚 Öğretmenin Branşını yazınız (Örn: `Matematik`):",
        "teacher_added_card": "✅ *Öğretmen Başarıyla Kaydedildi!*\n\n👤 İsim: *{name}*\n📚 Branş: *{subject}*\n\n🔑 *Öğretmen Giriş Kodu:*\n`{code}`",
        "btn_edit_student": "✏️ Bilgileri Düzenle",
        "btn_reset_codes": "🔄 Kodları Sıfırla",
        "btn_unlink_parent": "👨‍👩‍👧‍👦 Veli Bağlantısını Kopar",
        "btn_del_student": "❌ Öğrenciyi Sil",
        "student_deleted": "🗑️ Öğrenci sistemden silindi.",
        "codes_reset_done": "✅ Giriş kodları yenilendi!\n\n• Yeni Öğrenci: `{st_code}`\n• Yeni Veli: `{pr_code}`",
        "parent_unlinked_success": "✅ Veli bağlantıları koparıldı. Yeni Veli Kodu: `{code}`",
        "excel_info": "📥 *Excel ile Toplu Öğrenci Yükleme*\n\nLütfen `.xlsx` dosyasını gönderiniz.\nBaşlıklar: `Ad Soyad` | `Sinif` | `Numara`",
        "excel_done": "✅ Excel işlendi! Eklenen öğrenci: *{count}*\nŞifreler ekteki dosyada üretildi.",
        "select_pdf_class": "📄 Şifre kartlarını indirmek istediğiniz sınıfı seçiniz:",
        "pdf_ready": "📄 *{class_name}* şifre kartları ektedir.",
        "pending_medical_title": "🏥 *Bekleyen Mazeret Raporları:*",
        "no_pending_medical": "✅ Bekleyen mazeret raporu bulunmamaktadır.",
        "medical_approved": "✅ Rapor onaylandı. Öğrenci izinli sayıldı.",
        "medical_rejected": "❌ Rapor reddedildi.",
        "prompt_broadcast": "📢 Göndermek istediğiniz duyuru metnini yazınız:",
        "broadcast_success": "📢 Duyuru *{count}* kullanıcıya iletildi.",
        "prompt_search_student": "🔍 Öğrenci adı veya numarası yazınız:",
        "search_no_results": "❌ Eşleşen öğrenci bulunamadı.",
        "search_results_title": "🔍 *Arama Sonuçları:*",
        "btn_add_admin_id": "➕ Telegram ID ile Yönetici Ekle",
        "btn_gen_admin_code": "🔑 Tek Kullanımlık Yönetici Kodu Üret",
        "btn_make_perm_admin": "👑 Kalıcı Yönetici Yap",
        "btn_make_temp_admin": "⏱️ Geçici Yönetici Yap",
        "btn_revoke_admin_perm": "❌ Yönetici Yetkisini Al",
        "btn_send_dm": "✉️ Özel Mesaj Gönder",
        "btn_req_chat": "📞 1:1 İletişim İsteği Gönder",
        "btn_ban_user": "🚫 Engelle (Ban)",
        "btn_unban_user": "🟢 Engeli Kaldır (Unban)",
        "btn_refresh_data": "🔄 Verileri Yenile",
        "btn_users_list": "⬅️ Kullanıcılar Listesi",
        "perm_admin_assigned_toast": "✅ Kullanıcı kalıcı yönetici yapıldı.",
        "temp_admin_choose_title": "⏱️ *Geçici Yöneticilik Süresi Seçiniz:*",
        "btn_dur_1h": "⏱️ 1 Saat",
        "btn_dur_24h": "⏱️ 24 Saat",
        "btn_dur_7d": "⏱️ 7 Gün",
        "btn_dur_30d": "⏱️ 30 Gün",
        "temp_admin_assigned_toast": "✅ Kullanıcı {dur} süreyle geçici yönetici yapıldı.",
        "admin_demoted_toast": "Yönetici yetkisi kaldırıldı.",
        "permanent_admin_protected": "⛔ Kalıcı / Kurucu yöneticinin yetkisi kaldırılamaz!",
        "user_banned_toast": "Kullanıcı engellendi.",
        "user_unbanned_toast": "Kullanıcının engeli kaldırıldı.",
        "send_dm_prompt": "✉️ *Özel Mesaj Gönder:*\n\nLütfen `{id}` ID'li kullanıcıya iletmek istediğiniz mesajı yazınız:",
        "dm_sent_success": "✅ Mesaj kullanıcıya başarıyla iletildi!",
        "dm_delivery_error": "⚠️ İletim Hatası: Kullanıcı botu engellemiş.",
        "chat_req_sent_toast": "✅ 1:1 İletişim isteği kullanıcıya iletildi!",
        "chat_req_error_toast": "⚠️ Kullanıcı botu engellediği için iletilemedi!",
        "all_notifs_acknowledged": "✅ *Tüm Devamsızlık Bildirimleri Velilerce Onaylandı!*\n\nSon 36 saatte velisi tarafından okunmamış hiçbir bildirim bulunmamaktadır.",
        "menu_teacher": "👨‍🏫 *Öğretmen Masası*\nÖğretmen: *{name}* ({subject})",
        "rk_attendance": "📋 Hızlı Yoklama",
        "rk_grade": "📝 Not Girişi",
        "rk_homework": "📢 Ödev Panosu",
        "rk_appointments": "🤝 Veli Randevuları",
        "rk_behavior": "⭐ Davranış & Rozet",
        "btn_attendance": "📋 Hızlı Yoklama",
        "btn_enter_grade": "📝 Not Girişi",
        "btn_homework_board": "📢 Ödev Panosu",
        "attendance_intro": "📋 *{class_name} Yoklaması*\nGelmeyenlerin üzerine tıklayıp kaydediniz:",
        "btn_save_att": "💾 Yoklamayı Kaydet",
        "att_saved": "✅ Yoklama kaydedildi. 15 dakikalık düzeltme süresi başladı.",
        "prompt_select_exam_type": "📝 Öğrenci: *{name}* ({class_name})\n\nLütfen sınav / değerlendirme türünü seçiniz:",
        "exam_written_1": "📝 1. Yazılı",
        "exam_written_2": "📝 2. Yazılı",
        "exam_oral": "🗣️ Sözlü / Performans",
        "prompt_grade_score": "Öğrenci: *{name}* ({class_name})\nSınav Türü: *{exam_type}*\n\nNotu giriniz (0-100):",
        "prompt_grade_badge": "Değerlendirme rozeti seçiniz:",
        "badge_praise": "🟢 Başarılı / Övgü",
        "badge_missing": "🟡 Eksik / Tekrar",
        "badge_warning": "🔴 Uyarı / Dikkat",
        "grade_saved_success": "✅ Not kaydedildi ve veliye bildirildi.",
        "prompt_hw_class": "Ödevin sınıfını seçiniz:",
        "prompt_hw_content": "Ödev açıklamasını yazınız veya fotoğraf gönderiniz:",
        "hw_sent_success": "📢 Ödev *{class_name}* sınıfına iletildi.",
        "menu_parent": "👨‍👩‍👧‍👦 *Veli Masası*\nÖğrenci: *{name}* ({class_name})",
        "rk_report": "📊 Durum Paneli (Karne)",
        "rk_upload_medical": "🏥 Mazeret / Rapor",
        "rk_parent_info": "ℹ️ Okul Bilgi Panosu",
        "rk_switch_student": "🧑‍🎓 Öğrenci Değiştir",
        "rk_schedule": "📅 Ders Programı",
        "rk_notices": "📢 Duyurular",
        "rk_cafeteria": "🍲 Yemekhane Menüsü",
        "btn_report_card": "📊 Durum Paneli (Karne)",
        "btn_switch_student": "🧑‍🎓 Öğrenci Değiştir",
        "btn_upload_medical": "🏥 Mazeret / Rapor Yükle",
        "btn_download_pdf_report": "📄 Resmi PDF Karne İndir",
        "pdf_report_ready": "📄 *{name}* öğrencimizin resmi dönem not ve gelişim karnesi ektedir.",
        "upload_med_prompt": "Lütfen mazeret veya rapor belgesinin fotoğrafını gönderiniz:",
        "med_uploaded_success": "Rapor okul idaresine iletildi.",
        "student_switched_success": "Aktif öğrenci seçildi: *{name}* ({class_name})",
        "menu_student": "🎓 *Öğrenci Masası*\nÖğrenci: *{name}* ({class_name} - No: {no})",
        "no_grades": "Henüz girilmiş bir ders notu bulunmamaktadır.",
        "no_linked_student": "⚠️ Hesabınıza tanımlı bir öğrenci bulunamadı.",
        "evening_briefing_header": "🌙 *GÜN SONU BÜLTENİ (18:30)*\nÖğrenci: *{name}* ({class_name})\n\n📌 Devamsızlık: *{att_status}*\n📝 Notlar:\n{grades}",
        "btn_req_access": "📝 Şifre / Erişim Talep Et",
        "btn_cancel_action": "⬅️ İptal",
        "rk_cancel_action": "❌ İşlemi İptal Et",
        "action_cancelled": "❌ *İşlem iptal edildi.*",
        "admin_restart_confirmed": "🔄 *Yönetici Paneli Yeniden Başlatıldı.*",
        "btn_appr_appointment": "✅ Kabul Et",
        "btn_view_photo": "Görseli Aç",
        "btn_class_att_sheet": "Devamsızlık Çizelgesi",
        "btn_class_pdf_cards": "Şifre Kartları (PDF)",
        "btn_class_grade_sheet": "Not Çizelgesi",
        "btn_send_new_hw": "Yeni Ödev Gönder",
        "btn_class_sched": "Programı",
        "btn_change_admin_pin": "🔐 İdari PIN Kodunu Değiştir",
        "btn_toggle_readonly": "🔒 Salt-Okunur Modu ({status})",
        "btn_weekend_attendance": "📅 Hafta Sonu Yoklaması ({status})",
        "btn_clean_logs": "🧹 90 Günlük Logları Temizle",
        "btn_export_all_data": "📦 Tam Sistem Veri Yedeği (.xlsx)",
        "btn_restore_backup": "♻️ Excel'den Veri Geri Yükle",
        "btn_timezone_setting": "🕒 Zaman Dilimi (UTC+{offset})",
        "btn_class_promotion": "🎓 Dönem Sonu Sınıf Atlatma",
        "prompt_pin_current": "Lütfen MEVCUT 4 haneli PIN kodunuzu tuşlayınız:",
        "prompt_pin_new": "Lütfen YENİ 4 haneli PIN kodunuzu tuşlayınız:",
        "prompt_pin_confirm": "Lütfen yeni PIN kodunuzu ONAYLAMAK için tekrar tuşlayınız:",
        "pin_changed_success": "✅ <b>İdari PIN Kodu Başarıyla Değiştirildi!</b>",
        "pin_current_wrong": "Mevcut PIN kodu hatalı!",
        "pin_mismatch_error": "Girdiğiniz PIN kodları birbiriyle uyuşmuyor!",
        "invalid_admin_pin": "Hatalı PIN! Güvenlik nedeniyle işlem iptal edildi.",
        "permanent_admin_title": "Kalıcı İdareci",
        "school_admin_title": "Okul Yönetimi",
        "lbl_role_admin": "Yönetici",
        "lbl_role_teacher": "Öğretmen",
        "lbl_role_parent": "Veli",
        "lbl_role_student": "Öğrenci",
        "lbl_role_guest": "Misafir",
        "lbl_not_admin": "Yönetici Değil",
        "lbl_status_banned": "🚫 Engellendi",
        "lbl_subject": "Branş",
        "lbl_class": "Sınıf",
        "lbl_number": "Numara",
        "lbl_linked_students": "Bağlı Öğrenciler",
        "req_already_pending": "⚠️ Zaten bekleyen bir başvurunuz bulunmaktadır.",
        "req_role_select": "Lütfen başvurmak istediğiniz rolü seçiniz:",
        "role_teacher_btn": "👨‍🏫 Öğretmen Başvurusu",
        "role_parent_btn": "👨‍👩‍👧‍👦 Veli Başvurusu",
        "role_student_btn": "🎓 Öğrenci Başvurusu",
        "req_name_prompt": "Lütfen Adınızı ve Soyadınızı yazınız (Örn: `Ahmet Yılmaz`):",
        "err_invalid_name_strict": "⚠️ Lütfen en az iki kelimeden oluşan gerçek bir Ad ve Soyad yazınız:",
        "prompt_req_gender": "Lütfen Cinsiyetinizi seçiniz:",
        "btn_gender_male": "Erkek 👨",
        "btn_gender_female": "Kadın 👩",
        "err_invalid_gender_strict": "⚠️ Lütfen geçerli bir cinsiyet seçiniz:",
        "prompt_req_birth_date": "Lütfen Doğum Tarihinizi Gün.Ay.Yıl olarak yazınız (Örn: `15.04.1988`):",
        "err_invalid_birth_date_strict": "⚠️ Geçersiz doğum tarihi! Lütfen `GG.AA.YYYY` formatında yazınız:",
        "req_phone_prompt": "Lütfen İletişim Telefon Numaranızı yazınız veya paylaşınız:",
        "btn_share_contact": "📱 Telefon Numaramı Paylaş",
        "err_invalid_phone_strict": "⚠️ Lütfen geçerli bir telefon numarası giriniz:",
        "req_details_teacher": "Lütfen Branşınızı ve Görev Bilginizi yazınız:",
        "req_details_parent": "Lütfen velisi olduğunuz Öğrencinin Adını, Sınıfını veya Okul Numarasını yazınız:",
        "req_details_student": "Lütfen Sınıfınızı ve Okul Numaranızı yazınız:",
        "err_invalid_details_strict": "⚠️ Lütfen açıklayıcı bilgi giriniz:",
        "req_sent_success": "✅ Başvurunuz okul idaresine iletildi. Onaylandığında bilgilendirileceksiniz.",
        "no_pending_requests": "✅ Bekleyen başvuru bulunmamaktadır.",
        "pending_requests_title": "🛎️ *Bekleyen Erişim ve Şifre Başvuruları:*",
        "request_already_handled": "Bu başvuru zaten sonuçlandırılmıştır.",
        "unauthorized_action": "Bu işlem için yetkiniz bulunmamaktadır.",
        "req_approved_user": "🎉 Tebrikler! Başvurunuz onaylandı. Sisteme **{role}** olarak erişebilirsiniz.",
        "req_approved_admin_msg": "✅ Başvuru onaylandı: #{id} - {name} ({role})",
        "req_rejected_user": "❌ Başvurunuz okul idaresi tarafından uygun görülmedi.",
        "req_rejected_admin_msg": "❌ Başvuru reddedildi: #{id} - {name}",
        "btn_appr_request": "✅ Onayla",
        "btn_reject": "❌ Reddet",
        "lbl_full_name": "Ad Soyad",
        "lbl_gender": "Cinsiyet",
        "lbl_birth_date": "Doğum Tarihi",
        "lbl_phone": "Telefon",
        "lbl_role": "Rol",
        "lbl_age": "yaş",
        "btn_add_another": "➕ Yeni Ekle",
        "student_name_invalid": "⚠️ Lütfen geçerli bir isim giriniz.",
        "teacher_name_invalid": "⚠️ Lütfen geçerli bir isim giriniz.",
        "duplicate_student_no_error": "⚠️ {class_name} sınıfında {no} numaralı öğrenci zaten kayıtlı!",
        "btn_search_again": "🔍 Yeniden Ara",
        "teacher_search_prompt": "Öğretmen adı veya branşı yazınız:",
        "teacher_search_no_results": "❌ Eşleşen öğretmen bulunamadı.",
        "teacher_search_results_title": "🔍 *Öğretmen Arama Sonuçları:*",
        "btn_add_new_teacher": "➕ Yeni Öğretmen Ekle",
        "teacher_not_found": "Öğretmen bulunamadı.",
        "btn_write_telegram": "💬 Telegram'dan Yaz",
        "btn_manage_tch_classes": "🏫 Sorumlu Sınıflar",
        "btn_transfer_class": "🔄 Sınıfı Devret",
        "btn_add_co_teacher": "➕ Ortak Hoca Ekle",
        "btn_del_teacher": "❌ Öğretmeni Sil",
        "btn_confirm_reset": "✅ Evet, Sıfırla",
        "confirm_reset_codes_prompt": "Giriş kodlarını sıfırlamak istediğinize emin misiniz?",
        "confirm_delete_teacher_prompt": "*{name}* öğretmenini sistemden silmek istediğinize emin misiniz?",
        "btn_confirm_delete": "🗑️ Evet, Sil",
        "teacher_deleted": "Öğretmen sistemden silindi.",
        "lbl_class_teachers": "Sorumlu Öğretmenler",
        "btn_edit_name": "✏️ İsim Düzenle",
        "btn_edit_class": "🏫 Sınıf Düzenle",
        "btn_edit_no": "🔢 Okul No Düzenle",
        "student_info_updated": "✅ Bilgiler güncellendi: *{name}* ({class_name} - No: {no})",
        "student_not_found": "Öğrenci bulunamadı.",
        "confirm_delete_student_prompt": "*{name}* öğrencisini silmek istediğinize emin misiniz?",
        "user_not_found_toast": "Kullanıcı bulunamadı.",
        "admin_promoted_notification": "👑 *Tebrikler!*\n*{name}*, okul yöneticisi olarak yetkilendirildiniz.",
        "admin_demoted_notification": "ℹ️ Okul yöneticisi yetkiniz kaldırılmıştır.",
        "admin_add_tg_id_prompt": "Lütfen eklenecek yöneticinin Telegram ID'sini yazınız:",
        "admin_invalid_tg_id": "⚠️ Geçersiz Telegram ID!",
        "admin_add_name_prompt": "Yöneticinin Adını ve Soyadını yazınız:",
        "admin_added_success": "✅ Yönetici eklendi: *{name}* (`{id}`)",
        "admin_code_generated": "🔑 *Yönetici Erişim Kodu Üretildi:*\n\n`{code}`\n\nBu kod tek kullanımlıktır.",
        "broadcast_hub_title": "📢 *Toplu Duyuru Merkezi*\nLütfen hedef kitleyi seçiniz:",
        "bc_target_all": "🌐 Tüm Okul (Herkes)",
        "bc_target_teachers": "👨‍🏫 Sadece Öğretmenler",
        "bc_target_parents": "👨‍👩‍👧‍👦 Sadece Veliler",
        "bc_target_students": "🎓 Sadece Öğrenciler",
        "bc_target_class": "🏫 Belirli Bir Sınıf",
        "bc_target_lang": "🌍 Dile Göre Hedefle",
        "bc_select_class": "Duyuru yapılacak sınıfı seçiniz:",
        "bc_select_lang": "Duyurunun iletileceği dili seçiniz:",
        "prompt_broadcast_content": "📢 *Hedef:* {target}\nLütfen duyuru metnini yazınız veya görsel gönderiniz:",
        "broadcast_sent_report": "📢 Duyuru *{count}* kullanıcıya başarıyla ulaştırıldı.",
        "promotion_confirm_prompt": "⚠️ *Dönem Sonu Sınıf Atlatma*\n\nTüm sınıflar 1 seviye yükseltilecek, 12. sınıflar MEZUN durumuna getirilecektir.\n\nİşlem öncesinde tüm veritabanı otomatik olarak Excel'e yedeklenecektir.\nDevam etmek istiyor musunuz?",
        "promotion_success": "🎉 Sınıf atlatma tamamlandı! Toplam *{count}* öğrenci terfi ettirildi.",
        "logs_cleaned_toast": "🧹 {count} adet eski sistem kaydı temizlendi.",
        "timezone_updated": "🕒 Zaman dilimi UTC+{offset} olarak güncellendi.",
        "select_teacher_appointment": "Randevu almak istediğiniz öğretmeni seçiniz:",
        "no_registered_teachers": "Kayıtlı öğretmen bulunamadı.",
        "prompt_appointment_note": "Görüşmek istediğiniz konuyu veya uygun olduğunuz zamanı yazınız:",
        "appointment_sent": "✅ Randevu talebiniz öğretmene iletildi.",
        "no_pending_appointments": "Bekleyen randevu talebi bulunmamaktadır.",
        "pending_appointments_title": "🤝 *Bekleyen Veli Randevu Talepleri:*",
        "appointment_not_found": "Randevu bulunamadı.",
        "btn_not_available": "❌ Uygun Değilim",
        "appointment_approved_msg": "✅ Randevu talebiniz öğretmen tarafından kabul edildi.",
        "appointment_rejected_msg": "❌ Randevu talebiniz uygun görülmedi.",
        "appointment_confirmed_toast": "Randevu kabul edildi.",
        "setting_updated_toast": "Ayar güncellendi.",
        "schedule_select_class": "Programını görmek istediğiniz sınıfı seçiniz:",
        "no_blacklisted": "Engellenmiş kullanıcı bulunmamaktadır.",
        "blacklisted_title": "🚫 *Engelli ve Kilitli Kullanıcılar:*",
        "unban_success": "Kullanıcının engeli kaldırıldı.",
        "admin_unban_notification": "🟢 Hesabınızın engeli kaldırılmıştır.",
        "admin_sched_edit_title": "Düzenlemek istediğiniz sınıfın ders programını seçiniz:",
        "admin_sched_updated": "✅ *{class_name}* sınıfı ders programı güncellendi.",
        "prompt_menu_update": "Lütfen günün yemek menüsünü yazınız:",
        "menu_updated": "✅ Yemekhane menüsü güncellendi.",
        "search_user_prompt": "Aramak istediğiniz kullanıcının adını, kullanıcı adını veya Telegram ID'sini yazınız:",
        "search_user_no_results": "Kullanıcı bulunamadı.",
        "search_user_results_title": "🔍 *Kullanıcı Arama Sonuçları:*",
        "user_temp_banned_notification": "⏱️ Hesabınız {dur} süreyle askıya alınmıştır.",
        "user_temp_banned_toast": "Kullanıcı {dur} süreyle askıya alındı.",
        "contact_req_header": "📞 *Okul Yönetimi Sizinle İletişime Geçmek İstiyor!*\nİdareci: *{name}*",
        "contact_req_direct": "\n\nAşağıdaki butona basarak doğrudan yazabilirsiniz:",
        "btn_write_to_admin": "💬 İdareciye Mesaj Yaz",
        "contact_req_id": "Lütfen Telegram arama kısmından iletişime geçiniz.",
        "prompt_restore_backup": "♻️ *Excel Veri Geri Yükleme*\n\nLütfen daha önce sistemden indirdiğiniz `.xlsx` yedek dosyasını gönderiniz:",
        "restore_success": "✅ Veri geri yükleme tamamlandı! ({s_cnt} öğrenci, {t_cnt} öğretmen güncellendi).",
        "prompt_upload_teacher_excel": "👨‍🏫 Lütfen öğretmen listesini içeren `.xlsx` dosyasını gönderiniz:\nBaşlıklar: `Ad Soyad` | `Brans`",
        "teacher_excel_done": "✅ {count} adet öğretmen sisteme eklendi ve kodları üretildi.",
        "file_size_exceeded_error": "⚠️ Dosya boyutu çok büyük! (Maksimum: 10 MB)",
        "file_type_not_allowed_error": "⚠️ Geçersiz dosya formatı! Sadece `.xlsx`, `.pdf` veya görsel kabul edilir.",
        "excel_format_error": "⚠️ Excel dosyası okunamadı veya sütun başlıkları hatalı!",
        "unauthorized_excel_upload": "Bu işlem için yönetici yetkisi gereklidir.",
        "maintenance_mode_updated": "Bakım modu güncellendi.",
        "readonly_mode_updated": "Salt-okunur modu güncellendi.",
        "readonly_mode_active_alert": "🔒 Sistem şu anda salt-okunur modundadır. Veri değişikliği yapılamaz!",
        "weekend_attendance_updated": "Hafta sonu yoklama ayarı güncellendi.",
        "no_students_in_class": "Bu sınıfta öğrenci bulunmuyor.",
        "no_assigned_classes_teacher": "Size atanmış sorumlu sınıf bulunmuyor.",
        "attendance_select_class": "Yoklama almak istediğiniz sınıfı seçiniz:",
        "attendance_weekend_lock": "⚠️ Hafta sonu yoklama girişi kilitlidir.",
        "attendance_hours_lock": "⚠️ Yoklama girişi sadece 07:00 - 21:00 saatleri arasında yapılabilir.",
        "attendance_correction_notification": "ℹ️ *Yoklama Düzeltmesi:* *{name}* adlı öğrencinin yoklama durumu güncellenmiştir.",
        "grade_select_class": "Not girişi yapmak istediğiniz sınıfı seçiniz:",
        "btn_recent_grades_menu": "🕒 Son Girilen Notlar",
        "btn_class_grade_sheet": "Sınıf Not Çizelgesi",
        "grade_select_student": "*{class_name}* sınıfı için öğrenci seçiniz:",
        "invalid_score_format": "⚠️ Lütfen geçerli bir sayı giriniz (Örn: 85 veya 92.5).",
        "invalid_score_range": "⚠️ Not 0 ile 100 arasında olmalıdır!",
        "grade_parent_notification": "📝 *YENİ NOT GİRİŞİ*\n\n🧑‍🎓 Öğrenci: *{name}*\n📚 Ders: *{subject}* ({exam_type})\n📊 Not: *{score}* ({badge})",
        "recent_grades_title": "🕒 *Son Girilen Notlar (Öğretmen Masası):*",
        "no_permission_grade": "Bu notu düzenleme yetkiniz yok.",
        "btn_edit_grade": "✏️ Notu Düzenle",
        "btn_del_grade": "🗑️ Notu Sil",
        "prompt_new_score": "Lütfen yeni notu giriniz (0-100):",
        "grade_updated": "✅ Not güncellendi.",
        "grade_deleted": "Not sistemden silindi.",
        "btn_add_positive_badge": "🟢 Olumlu Davranış / Övgü",
        "btn_add_negative_badge": "🔴 Geliştirilmeli / Uyarı",
        "btn_student_behavior_history": "📜 Davranış Geçmişi",
        "prompt_behavior_note": "Öğrenci: *{name}*\nRozet: {badge} *{title}*\n\nVarsa açıklama notu yazınız (veya `-` gönderiniz):",
        "behavior_parent_notification": "⭐ *DAVRANIŞ DEĞERLENDİRMESİ*\n\n🧑‍🎓 Öğrenci: *{name}* ({class_name})\n🏷️ Durum: {badge} *{title}*\n📝 Açıklama: _{note}_\n👨‍🏫 Öğretmen: {teacher}",
        "behavior_saved_success": "✅ Davranış değerlendirmesi kaydedildi ve veliye bildirildi.",
        "no_behavior_records": "Kayıtlı davranış değerlendirmesi bulunmuyor.",
        "btn_send_new_hw": "📢 Yeni Ödev Gönder",
        "btn_my_hws": "📚 Yayınladığım Ödevler",
        "no_hws_found": "Yayınlanmış ödeviniz bulunmuyor.",
        "homework_deleted_toast": "Ödev silindi.",
        "no_active_homeworks": "*{class_name}* sınıfı için aktif ödev bulunmuyor.",
        "homework_board_title": "📚 *{class_name} Sınıfı Ödev Panosu:*",
        "btn_submit_hw": "📤 Ödev Teslim Et",
        "image_load_error": "Görsel yüklenemedi.",
        "prompt_hw_submission": "📚 *{subject} Ödevi Teslimi*\n\nLütfen ödevinizi anlatan bir metin yazınız veya ödevinizin fotoğrafını gönderiniz:",
        "hw_submission_received": "✅ Ödeviniz öğretmene iletildi.",
        "btn_view_submissions": "📥 Teslimleri Gör",
        "btn_hw_approve": "✅ Kabul Et",
        "btn_hw_revision": "🔄 Düzeltme İste",
        "hw_feedback_sent_user": "📝 *Ödev Durumu Güncellemesi*\n\nDers: *{subject}*\nDurum: *{status}*\nÖğretmen Notu: _{feedback}_",
        "no_exams_found": "Kayıtlı sınav takvimi bulunmuyor.",
        "btn_add_exam": "➕ Sınav Ekle",
        "exam_schedule_title": "📅 *{class_name} Sınav Takvimi:*",
        "emergency_alert_prompt": "🚨 *KIRMIZI ALARM & ACİL DURUM BİLDİRİMİ*\n\nLütfen tüm velilere yüksek sesli acil bildirim olarak gönderilecek mesajı yazınız:",
        "btn_emergency_ack": "🚨 Okudum, Bilgim Var",
        "emergency_monitor_title": "🚨 *Acil Durumu Henüz Onaylamayan Veliler*",
        "parent_choose_child": "Lütfen işlem yapmak istediğiniz öğrenciyi seçiniz:",
        "no_registered_students": "Kayıtlı öğrenci bulunamadı.",
        "prompt_add_child_code": "Lütfen eklemek istediğiniz diğer öğrencinin Veli Kodunu (`VELI-XXXX`) yazınız:",
        "invalid_parent_code": "⚠️ Geçersiz veli kodu!",
        "child_added_success": "✅ *{name}* ({class_name}) hesabınıza eklendi.",
        "no_permission_student_record": "Bu öğrenci için yetkiniz yok.",
        "export_ready": "📦 Okul genel veri yedeği ({date}) ektedir.",
        "excel_hub_title": "📥 *Kurumsal Excel ve Veri Merkezi*",
        "btn_upload_excel": "📥 Toplu Öğrenci Yükle",
        "btn_upload_teacher_excel": "👨‍🏫 Toplu Öğretmen Yükle",
        "select_class_to_transfer": "Devretmek istediğiniz sınıfı seçiniz:",
        "select_target_teacher": "*{class_name}* sınıfı hangi öğretmene devredilsin?",
        "class_transfer_done": "✅ {class_name} sınıfı {teacher} hocaya devredildi.",
        "select_class_to_co_teacher": "Ortak hoca atamak istediğiniz sınıfı seçiniz:",
        "class_co_teacher_done": "✅ {teacher} hoca {class_name} sınıfına ortak atandı.",
        "btn_assign_all_classes": "🌐 Tüm Sınıfları Ata",
        "btn_clear_all_classes": "🗑️ Tümünü Temizle",
        "remind_att_sent": "Öğretmenlere yoklama hatırlatması iletildi.",
        "att_check_title": "📋 *Yoklama Denetim Masası ({date}):*",
        "att_check_all_done": "✅ Tüm sınıfların sabah yoklaması eksiksiz sisteme girilmiştir.",
        "btn_remind_att": "⚠️ Yoklamayı Hatırlat",
        "btn_teachers_pdf": "👨‍🏫 Tüm Öğretmenler (PDF)",
        "photo_expected_medical": "⚠️ Lütfen mazeret belgesinin fotoğrafını gönderiniz.",
        "btn_appr_medical": "✅ Kabul Et",
        "btn_cancel_pin": "❌ Vazgeç"
    }
}

# Diğer dilleri Türkçeden türetip özel metinleri zenginleştiriyoruz (Ultra-kompakt & eksiksiz)
LOCALES["ru"] = dict(LOCALES["tr"])
LOCALES["uz"] = dict(LOCALES["tr"])
LOCALES["en"] = dict(LOCALES["tr"])

# Rusça özel anahtarlar
LOCALES["ru"].update({
    "lang_changed": "Язык интерфейса: 🇷🇺 Русский",
    "welcome_guest": "🎓 *Добро пожаловать в школьную систему управления.*\n\nВведите ваш код доступа (Напр: `VELI-123456`, `HCA-123456`, `OGR-123456`):",
    "btn_login_prompt": "🔑 Войти по коду",
    "auth_success": "✅ *Успешный вход!*\nДобро пожаловать: *{name}*\nРоль: *{role}*",
    "auth_failed": "❌ Неверный код! Осталось попыток: {remaining}",
    "auth_locked": "⛔ Аккаунт заблокирован на 1 час из соображений безопасности.",
    "rk_cat_staff": "👥 Ученики и учителя",
    "rk_cat_reports": "📊 Отчеты и контроль",
    "rk_cat_requests": "🛎️ Центр одобрений",
    "rk_cat_tools": "🛠️ Инструменты",
    "rk_cat_settings": "⚙️ Настройки системы",
    "rk_restart": "🔄 Перезапуск",
    "rk_main_menu": "🏠 Главное меню",
    "rk_cancel_action": "❌ Отмена",
    "action_cancelled": "❌ *Действие отменено.*",
    "btn_classes": "🏫 Классы и ученики",
    "btn_teachers": "👨‍🏫 Учителя",
    "btn_back": "⬅️ Назад",
    "btn_main_menu": "🏠 Главное меню",
    "btn_cancel_action": "⬅️ Отмена",
    "btn_cancel_pin": "❌ Отмена",
    "pin_changed_success": "✅ <b>ПИН-код администратора успешно обновлен!</b>",
    "prompt_pin_current": "Пожалуйста, введите ТЕКУЩИЙ 4-значный ПИН-код:",
    "prompt_pin_new": "Пожалуйста, введите НОВЫЙ 4-значный ПИН-код:",
    "prompt_pin_confirm": "Пожалуйста, подтвердите новый ПИН-код:",
    "invalid_admin_pin": "Неверный ПИН! В целях безопасности операция отменена.",
    "lbl_role_admin": "Администратор",
    "lbl_role_teacher": "Учитель",
    "lbl_role_parent": "Родитель",
    "lbl_role_student": "Ученик",
    "lbl_role_guest": "Гость"
})

# Özbekçe özel anahtarlar
LOCALES["uz"].update({
    "lang_changed": "Til muvaffaqiyatli yangilandi: 🇺🇿 O'zbekcha",
    "welcome_guest": "🎓 *Maktab boshqaruv tizimiga xush kelibsiz.*\n\nTizimga kirish uchun maxsus kirish kodini kiriting (Masalan: `VELI-123456`, `HCA-123456`, `OGR-123456`):",
    "btn_login_prompt": "🔑 Kod orqali kirish",
    "auth_success": "✅ *Muvaffaqiyatli kirildi!*\nXush kelibsiz: *{name}*\nRolingiz: *{role}*",
    "auth_failed": "❌ Noto'g'ri kod! Qolgan urinishlar: {remaining}",
    "auth_locked": "⛔ Xavfsizlik sababli hisobingiz 1 soatga bloklandi.",
    "rk_cat_staff": "👥 Kadro va o'quvchilar",
    "rk_cat_reports": "📊 Hisobotlar va nazorat",
    "rk_cat_requests": "🛎️ Tasdiqlash markazi",
    "rk_cat_tools": "🛠️ Boshqaruv vositalari",
    "rk_cat_settings": "⚙️ Tizim va sozlamalar",
    "rk_restart": "🔄 Qayta ishga tushirish",
    "rk_main_menu": "🏠 Asosiy menyu",
    "rk_cancel_action": "❌ Bekor qilish",
    "action_cancelled": "❌ *Amal bekor qilindi.*",
    "btn_classes": "🏫 Sinflar va o'quvchilar",
    "btn_teachers": "👨‍🏫 O'qituvchilar",
    "btn_back": "⬅️ Orqaga",
    "btn_main_menu": "🏠 Asosiy menyu",
    "btn_cancel_action": "⬅️ Bekor qilish",
    "btn_cancel_pin": "❌ Bekor",
    "pin_changed_success": "✅ <b>Ma'muriy PIN kod muvaffaqiyatli o'zgartirildi!</b>",
    "prompt_pin_current": "Iltimos, AMALDAGI 4 xonali PIN kodni tering:",
    "prompt_pin_new": "Iltimos, YANGI 4 xonali PIN kodni tering:",
    "prompt_pin_confirm": "Iltimos, yangi PIN kodni TASDIQLASH uchun qayta tering:",
    "invalid_admin_pin": "Noto'g'ri PIN! Xavfsizlik sababli amal to'xtatildi.",
    "lbl_role_admin": "Ma'mur",
    "lbl_role_teacher": "O'qituvchi",
    "lbl_role_parent": "Ota-ona",
    "lbl_role_student": "O'quvchi",
    "lbl_role_guest": "Mehmon"
})

# İngilizce özel anahtarlar
LOCALES["en"].update({
    "lang_changed": "Language updated: 🇬🇧 English",
    "welcome_guest": "🎓 *Welcome to School Management System.*\n\nPlease enter your access code (e.g. `VELI-123456`, `HCA-123456`, `OGR-123456`):",
    "btn_login_prompt": "🔑 Log In with Code",
    "auth_success": "✅ *Login Successful!*\nWelcome: *{name}*\nRole: *{role}*",
    "auth_failed": "❌ Invalid code! Remaining attempts: {remaining}",
    "auth_locked": "⛔ Account locked for 1 hour for security.",
    "rk_cat_staff": "👥 Staff & Students",
    "rk_cat_reports": "📊 Reports & Audits",
    "rk_cat_requests": "🛎️ Approval Center",
    "rk_cat_tools": "🛠️ Admin Tools",
    "rk_cat_settings": "⚙️ System & Settings",
    "rk_restart": "🔄 Restart Bot",
    "rk_main_menu": "🏠 Main Menu",
    "rk_cancel_action": "❌ Cancel Action",
    "action_cancelled": "❌ *Action cancelled.*",
    "btn_classes": "🏫 Classes & Students",
    "btn_teachers": "👨‍🏫 Teachers",
    "btn_back": "⬅️ Back",
    "btn_main_menu": "🏠 Main Menu",
    "btn_cancel_action": "⬅️ Cancel",
    "btn_cancel_pin": "❌ Cancel",
    "pin_changed_success": "✅ <b>Admin PIN Updated Successfully!</b>",
    "prompt_pin_current": "Please enter your CURRENT 4-digit PIN:",
    "prompt_pin_new": "Please enter your NEW 4-digit PIN:",
    "prompt_pin_confirm": "Please RE-ENTER new PIN to confirm:",
    "invalid_admin_pin": "Invalid PIN! Security triggered, action cancelled.",
    "lbl_role_admin": "Admin",
    "lbl_role_teacher": "Teacher",
    "lbl_role_parent": "Parent",
    "lbl_role_student": "Student",
    "lbl_role_guest": "Guest"
})

def get_text(key: str, lang: str = "tr", **kwargs) -> str:
    lang_dict = LOCALES.get(lang, LOCALES["tr"])
    text = lang_dict.get(key, LOCALES["tr"].get(key, f"[{key}]"))
    if kwargs:
        try: text = text.format(**kwargs)
        except Exception: pass
    return text

# ======================================================================
# 4. ARAYÜZ VE ROLE ÖZEL KALICI ALT MENÜ MOTORU (REPLY KEYBOARD)
# ======================================================================

def get_role_reply_kb(role: str, lang: str = "tr") -> ReplyKeyboardMarkup:
    kb_list = []
    if role == "admin":
        kb_list = [
            [KeyboardButton(text=get_text("rk_cat_staff", lang)), KeyboardButton(text=get_text("rk_cat_reports", lang))],
            [KeyboardButton(text=get_text("rk_cat_requests", lang)), KeyboardButton(text=get_text("rk_cat_tools", lang))],
            [KeyboardButton(text=get_text("rk_cat_settings", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    elif role == "teacher":
        kb_list = [
            [KeyboardButton(text=get_text("rk_attendance", lang)), KeyboardButton(text=get_text("rk_grade", lang))],
            [KeyboardButton(text=get_text("rk_homework", lang)), KeyboardButton(text=get_text("rk_behavior", lang))],
            [KeyboardButton(text=get_text("rk_appointments", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    elif role == "parent":
        kb_list = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_upload_medical", lang))],
            [KeyboardButton(text=get_text("rk_switch_student", lang)), KeyboardButton(text=get_text("rk_homework", lang))],
            [KeyboardButton(text=get_text("rk_appointments", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    elif role == "student":
        kb_list = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_homework", lang))],
            [KeyboardButton(text=get_text("rk_schedule", lang)), KeyboardButton(text=get_text("rk_notices", lang))],
            [KeyboardButton(text=get_text("rk_cafeteria", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    else:
        kb_list = [
            [KeyboardButton(text=get_text("btn_login_prompt", lang)), KeyboardButton(text=get_text("btn_req_access", lang))],
            [KeyboardButton(text=get_text("btn_lang", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    return ReplyKeyboardMarkup(keyboard=kb_list, resize_keyboard=True, is_persistent=True)

def get_cancel_reply_kb(lang: str = "tr") -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=get_text("rk_cancel_action", lang))]],
        resize_keyboard=True,
        is_persistent=True
    )

def get_pin_reply_kb(lang: str = "tr") -> ReplyKeyboardMarkup:
    """Ekranda sıfır inline buton olmasını sağlayan, sadece alttan çalışan numaratör klavyesi"""
    return ReplyKeyboardMarkup(
        keyboard=[
            [KeyboardButton(text="1"), KeyboardButton(text="2"), KeyboardButton(text="3")],
            [KeyboardButton(text="4"), KeyboardButton(text="5"), KeyboardButton(text="6")],
            [KeyboardButton(text="7"), KeyboardButton(text="8"), KeyboardButton(text="9")],
            [KeyboardButton(text=get_text("rk_cancel_action", lang)), KeyboardButton(text="0"), KeyboardButton(text="⌫")]
        ],
        resize_keyboard=True,
        is_persistent=True
    )

def get_language_inline_kb() -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text="🇹🇷 Türkçe", callback_data="set_lang:tr"), InlineKeyboardButton(text="🇷🇺 Русский", callback_data="set_lang:ru")],
        [InlineKeyboardButton(text="🇺🇿 O'zbekcha", callback_data="set_lang:uz"), InlineKeyboardButton(text="🇬🇧 English", callback_data="set_lang:en")]
    ])

def get_nav_buttons(lang: str = "tr", back_callback: str = "adm:dashboard") -> list[InlineKeyboardButton]:
    return [
        InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=back_callback),
        InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")
    ]

def get_attendance_grid_kb(students: list, attendance_map: dict, class_name: str, lang: str = "tr") -> InlineKeyboardMarkup:
    buttons = []
    row = []
    for s in students:
        is_absent = attendance_map.get(s.id, False)
        icon = "🔴" if is_absent else "🟢"
        row.append(InlineKeyboardButton(text=f"{icon} {s.full_name}", callback_data=f"att_toggle:{s.id}"))
        if len(row) == 2:
            buttons.append(row)
            row = []
    if row: buttons.append(row)
    buttons.append([InlineKeyboardButton(text=f"🟢 {get_text('btn_save_att', lang)} (Tümü Mevcut)", callback_data=f"att_all_pres:{class_name}")])
    buttons.append([InlineKeyboardButton(text=get_text("btn_save_att", lang), callback_data=f"att_save:{class_name}")])
    buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:classes")])
    return InlineKeyboardMarkup(inline_keyboard=buttons)

def get_exam_type_kb(class_name: str, student_id: int, lang: str = "tr") -> InlineKeyboardMarkup:
    buttons = [
        [InlineKeyboardButton(text=get_text("exam_written_1", lang), callback_data="gr_type:1_yazili")],
        [InlineKeyboardButton(text=get_text("exam_written_2", lang), callback_data="gr_type:2_yazili")],
        [InlineKeyboardButton(text=get_text("exam_oral", lang), callback_data="gr_type:sozlu")],
        [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=f"gr_cls:{class_name}")]
    ]
    return InlineKeyboardMarkup(inline_keyboard=buttons)

# ======================================================================
# 5. YARDIMCI SERVİSLER, PDF, EXCEL VE SİSTEM AYARLARI
# ======================================================================

SETTINGS_CACHE = {}

async def get_cached_setting(key: str, default_val: str = "") -> str:
    if key in SETTINGS_CACHE: return SETTINGS_CACHE[key]
    async with AsyncSessionLocal() as session:
        st = await session.get(SystemSetting, key)
        val = st.value if st else default_val
        SETTINGS_CACHE[key] = val
        return val

async def is_readonly_mode_active() -> bool:
    val = await get_cached_setting("readonly_mode", "false")
    return val.strip().lower() == "true"

async def get_current_admin_pin() -> str:
    val = await get_cached_setting("admin_pin", "")
    if val and val.strip(): return val.strip()
    return "1923"

def is_admin_user(user: User | None, telegram_id: int) -> bool:
    if telegram_id in ADMIN_IDS: return True
    if not user: return False
    if user.admin_type == "permanent": return True
    if user.role == "admin":
        if user.admin_type == "temporary" and user.admin_until:
            if datetime.utcnow() > user.admin_until: return False
        return True
    return False

def generate_secure_code(prefix: str) -> str:
    rand = "".join(random.choices(string.digits, k=6))
    return f"{prefix}-{rand}"

async def log_audit(session: AsyncSession, user_id: int, user_name: str, action: str, details: str):
    log = AuditLog(user_id=user_id, user_name=user_name, action=action, details=details)
    session.add(log)

def render_progress_bar(val: int, total: int = 10) -> str:
    filled = max(0, min(val, total))
    return "█" * filled + "░" * (total - filled)

def setup_pdf_fonts():
    try:
        font_paths = ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/usr/share/fonts/TTF/DejaVuSans.ttf"]
        for p in font_paths:
            if os.path.exists(p):
                pdfmetrics.registerFont(TTFont("DejaVu", p))
                return "DejaVu"
    except Exception: pass
    return "Helvetica"

async def export_all_school_data_excel() -> io.BytesIO:
    wb = openpyxl.Workbook()
    ws_students = wb.active
    ws_students.title = "Öğrenciler"
    ws_students.append(["ID", "Ad Soyad", "Sınıf", "Okul No", "Öğrenci Kodu", "Veli Kodu", "Kullanıldı"])

    ws_teachers = wb.create_sheet(title="Öğretmenler")
    ws_teachers.append(["ID", "Ad Soyad", "Branş", "Giriş Kodu", "Sorumlu Sınıflar"])

    ws_grades = wb.create_sheet(title="Notlar")
    ws_grades.append(["Öğrenci ID", "Ders", "Sınav Türü", "Not", "Rozet", "Tarih"])

    ws_att = wb.create_sheet(title="Devamsızlık")
    ws_att.append(["Öğrenci ID", "Sınıf", "Tarih", "Durum"])

    async with AsyncSessionLocal() as session:
        st_list = (await session.execute(select(Student).order_by(Student.class_name, Student.student_number))).scalars().all()
        for s in st_list:
            ws_students.append([s.id, s.full_name, s.class_name, s.student_number, s.student_code, s.parent_code, "Evet" if s.is_student_code_burned else "Hayır"])

        tc_list = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        for t in tc_list:
            ws_teachers.append([t.id, t.full_name, t.subject, t.auth_code, t.assigned_classes or "ALL"])

        gr_list = (await session.execute(select(Grade).order_by(desc(Grade.created_at)).limit(500))).scalars().all()
        for g in gr_list:
            ws_grades.append([g.student_id, g.subject, g.exam_type, g.score, g.badge, g.created_at.strftime("%d.%m.%Y")])

        at_list = (await session.execute(select(Attendance).order_by(desc(Attendance.date)).limit(500))).scalars().all()
        for a in at_list:
            ws_att.append([a.student_id, a.class_name, a.date.strftime("%d.%m.%Y"), a.status])

    buf = io.BytesIO()
    wb.save(buf)
    buf.seek(0)
    return buf

async def process_student_excel(file_content: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_content))
    ws = wb.active
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Üretilen Kodlar"
    out_ws.append(["Ad Soyad", "Sınıf", "Okul No", "Öğrenci Kodu", "Veli Kodu"])

    count = 0
    async with AsyncSessionLocal() as session:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            name = str(row[0]).strip()
            class_name = str(row[1]).strip().upper() if len(row) > 1 and row[1] else "9-A"
            student_no = str(row[2]).strip() if len(row) > 2 and row[2] else str(100 + count)

            st_code = generate_secure_code("OGR")
            pr_code = generate_secure_code("VELI")

            st = Student(full_name=name, class_name=class_name, student_number=student_no, student_code=st_code, parent_code=pr_code)
            session.add(st)
            out_ws.append([name, class_name, student_no, st_code, pr_code])
            count += 1
        await session.commit()

    out_buf = io.BytesIO()
    out_wb.save(out_buf)
    out_buf.seek(0)
    return count, out_buf

async def process_teacher_excel(file_content: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_content))
    ws = wb.active
    out_wb = openpyxl.Workbook()
    out_ws = out_wb.active
    out_ws.title = "Öğretmen Kodları"
    out_ws.append(["Ad Soyad", "Branş", "Giriş Kodu"])

    count = 0
    async with AsyncSessionLocal() as session:
        for row in ws.iter_rows(min_row=2, values_only=True):
            if not row or not row[0]: continue
            name = str(row[0]).strip()
            subject = str(row[1]).strip() if len(row) > 1 and row[1] else "Genel"
            code = generate_secure_code("HCA")
            tch = Teacher(full_name=name, subject=subject, auth_code=code, assigned_classes="ALL")
            session.add(tch)
            out_ws.append([name, subject, code])
            count += 1
        await session.commit()

    out_buf = io.BytesIO()
    out_wb.save(out_buf)
    out_buf.seek(0)
    return count, out_buf

async def restore_all_school_data_excel(file_content: bytes):
    wb = openpyxl.load_workbook(io.BytesIO(file_content))
    s_cnt = 0
    t_cnt = 0
    async with AsyncSessionLocal() as session:
        if "Öğrenciler" in wb.sheetnames:
            ws_st = wb["Öğrenciler"]
            for row in ws_st.iter_rows(min_row=2, values_only=True):
                if not row or not row[1]: continue
                name, cls_n, no, st_c, pr_c = str(row[1]).strip(), str(row[2]).strip(), str(row[3]).strip(), str(row[4]).strip(), str(row[5]).strip()
                existing = (await session.execute(select(Student).where(Student.student_number == no, Student.class_name == cls_n))).scalar_one_or_none()
                if not existing:
                    st = Student(full_name=name, class_name=cls_n, student_number=no, student_code=st_c, parent_code=pr_c)
                    session.add(st)
                    s_cnt += 1
        if "Öğretmenler" in wb.sheetnames:
            ws_tc = wb["Öğretmenler"]
            for row in ws_tc.iter_rows(min_row=2, values_only=True):
                if not row or not row[1]: continue
                name, subj, code = str(row[1]).strip(), str(row[2]).strip(), str(row[3]).strip()
                existing = (await session.execute(select(Teacher).where(Teacher.full_name == name))).scalar_one_or_none()
                if not existing:
                    t = Teacher(full_name=name, subject=subj, auth_code=code, assigned_classes="ALL")
                    session.add(t)
                    t_cnt += 1
        await session.commit()
    return s_cnt, t_cnt

async def generate_classroom_pdf_cards(class_name: str, lang: str = "tr") -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    font_name = setup_pdf_fonts()
    styles = getSampleStyleSheet()

    title_style = ParagraphStyle(name="TStyle", fontName=font_name, fontSize=16, leading=20, alignment=1, textColor=colors.HexColor("#1A365D"))
    norm_style = ParagraphStyle(name="NStyle", fontName=font_name, fontSize=9, leading=12)

    story = [
        Paragraph(f"<b>OKUL YÖNETİM SİSTEMİ - GİRİŞ ŞİFRE KARTLARI</b>", title_style),
        Spacer(1, 5),
        Paragraph(f"<b>Sınıf: {escape_html(class_name)}</b> | Tarih: {datetime.utcnow().strftime('%d.%m.%Y')}", norm_style),
        Spacer(1, 10)
    ]

    async with AsyncSessionLocal() as session:
        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()

    table_data = [["No", "Öğrenci Adı Soyadı", "Öğrenci Kodu", "Veli Kodu"]]
    for s in students:
        table_data.append([str(s.student_number), s.full_name, s.student_code, s.parent_code])

    t = Table(table_data, colWidths=[40, 220, 140, 140])
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#EDF2F7")]),
    ]))
    story.append(t)
    doc.build(story)
    buf.seek(0)
    return buf

async def generate_teachers_pdf_cards(lang: str = "tr") -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    font_name = setup_pdf_fonts()
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(name="TStyle2", fontName=font_name, fontSize=16, leading=20, alignment=1, textColor=colors.HexColor("#1A365D"))
    story = [
        Paragraph("<b>ÖĞRETMENLER GİRİŞ ŞİFRE KARTLARI</b>", title_style),
        Spacer(1, 15)
    ]
    async with AsyncSessionLocal() as session:
        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()

    data = [["Adı Soyadı", "Branşı", "Giriş Kodu"]]
    for tch in teachers:
        data.append([tch.full_name, tch.subject, tch.auth_code])

    t = Table(data, colWidths=[240, 150, 150])
    t.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2C5282")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#EBF8FF")]),
    ]))
    story.append(t)
    doc.build(story)
    buf.seek(0)
    return buf

async def generate_student_report_card_pdf(student_id: int, lang: str = "tr") -> io.BytesIO:
    buf = io.BytesIO()
    doc = SimpleDocTemplate(buf, pagesize=A4, rightMargin=20, leftMargin=20, topMargin=20, bottomMargin=20)
    font_name = setup_pdf_fonts()
    styles = getSampleStyleSheet()

    t_title = ParagraphStyle(name="RP_T", fontName=font_name, fontSize=18, leading=22, alignment=1, textColor=colors.HexColor("#1A365D"))
    t_sub = ParagraphStyle(name="RP_S", fontName=font_name, fontSize=11, leading=14, alignment=1)
    norm = ParagraphStyle(name="RP_N", fontName=font_name, fontSize=10, leading=13)

    story = [
        Paragraph("<b>T.C. MİLLÎ EĞİTİM BAKANLIĞI</b>", t_title),
        Paragraph("<b>RESMİ ÖĞRENCİ GELİŞİM VE NOT DÖKÜM BELGESİ</b>", t_sub),
        Spacer(1, 15)
    ]

    async with AsyncSessionLocal() as session:
        st = await session.get(Student, student_id)
        grades = (await session.execute(select(Grade).where(Grade.student_id == student_id).order_by(Grade.created_at))).scalars().all()
        atts = (await session.execute(select(Attendance).where(Attendance.student_id == student_id))).scalars().all()

    abs_cnt = sum(1 for a in atts if a.status == "absent")
    exc_cnt = sum(1 for a in atts if a.status == "excused")

    info_data = [
        [f"Öğrenci: {st.full_name}", f"Sınıf: {st.class_name}", f"Okul No: {st.student_number}"],
        [f"Özürsüz Devamsızlık: {abs_cnt} Gün", f"Mazeretli: {exc_cnt} Gün", f"Tarih: {datetime.utcnow().strftime('%d.%m.%Y')}"]
    ]
    t_info = Table(info_data, colWidths=[200, 170, 170])
    t_info.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('GRID', (0,0), (-1,-1), 0.5, colors.HexColor("#CBD5E0")),
        ('BACKGROUND', (0,0), (-1,-1), colors.HexColor("#F7FAFC")),
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
    ]))
    story.append(t_info)
    story.append(Spacer(1, 15))

    gr_data = [["Ders Adı", "Sınav Türü", "Not", "Sonuç"]]
    scores = []
    for g in grades:
        scores.append(g.score)
        res_str = "BAŞARILI" if g.score >= 50 else "GELİŞTİRİLMELİ"
        gr_data.append([g.subject, g.exam_type or "1. Yazılı", str(g.score), res_str])

    gpa = round(sum(scores)/len(scores), 1) if scores else 100.0

    t_gr = Table(gr_data, colWidths=[200, 160, 90, 90])
    t_gr.setStyle(TableStyle([
        ('FONTNAME', (0,0), (-1,-1), font_name),
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#2B6CB0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.white),
        ('ALIGN', (0,0), (-1,-1), 'CENTER'),
        ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
        ('ROWBACKGROUNDS', (0,1), (-1,-1), [colors.white, colors.HexColor("#EDF2F7")]),
    ]))
    story.append(t_gr)
    story.append(Spacer(1, 15))
    story.append(Paragraph(f"<b>Genel Ağırlıklı Başarı Ortalaması: {gpa} / 100</b>", norm))

    doc.build(story)
    buf.seek(0)
    return buf

# ======================================================================
# 6. ARKA PLAN İŞÇİLERİ VE OTOMASYON
# ======================================================================

async def run_attendance_delay_worker(bot: Bot):
    now = datetime.utcnow()
    async with AsyncSessionLocal() as session:
        pending_atts = (await session.execute(
            select(Attendance).where(Attendance.status == "absent", Attendance.is_notified == False, Attendance.notify_at <= now)
        )).scalars().all()

        for att in pending_atts:
            st = await session.get(Student, att.student_id)
            if not st:
                att.is_notified = True
                continue

            parents = (await session.execute(
                select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st.id)
            )).scalars().all()

            msg_text = (
                f"🚨 <b>DEVAMSIZLIK BİLDİRİMİ</b>\n\n"
                f"Sayın Velimiz, <b>{escape_html(st.full_name)}</b> ({escape_html(st.class_name)}) "
                f"adlı öğrencimiz bugün (<b>{att.date.strftime('%d.%m.%Y')}</b>) sabah yoklamasında okulda bulunmamaktadır.\n\n"
                f"ℹ️ Bir mazereti varsa aşağıdaki butondan sağlık raporu veya mazeret belgesi yükleyebilirsiniz:"
            )
            report_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="🏥 Mazeret / Rapor Yükle", callback_data="upload_medical_init")]
            ])

            for p_id in parents:
                await safe_send_message(bot, p_id, msg_text, reply_markup=report_kb, parse_mode="HTML")
                crit = CriticalNotification(user_telegram_id=p_id, message_text=msg_text)
                session.add(crit)

            att.is_notified = True

        await session.commit()

async def run_evening_briefing_worker(bot: Bot):
    today = get_local_date()
    async with AsyncSessionLocal() as session:
        parents = (await session.execute(select(User).where(User.role == "parent", User.evening_briefing == True))).scalars().all()
        for p in parents:
            kids = (await session.execute(select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == p.telegram_id))).scalars().all()
            for k in kids:
                att = (await session.execute(select(Attendance).where(Attendance.student_id == k.id, Attendance.date == today))).scalar_one_or_none()
                att_status = "✅ Okulda Mevcut" if not att or att.status == "present" else ("🔴 Gelmedi (Devamsız)" if att.status == "absent" else "🟡 İzinli / Raporlu")

                since = datetime.utcnow() - timedelta(hours=14)
                today_grades = (await session.execute(select(Grade).where(Grade.student_id == k.id, Grade.created_at >= since))).scalars().all()
                gr_txt = "\n".join([f"• {g.subject}: {g.score} ({g.badge})" for g in today_grades]) if today_grades else "Bugün yeni not girilmedi."

                msg = get_text("evening_briefing_header", p.language, name=escape_md(k.full_name), class_name=escape_md(k.class_name), att_status=att_status, grades=escape_md(gr_txt))
                await safe_send_message(bot, p.telegram_id, msg, parse_mode="Markdown")
                await asyncio.sleep(0.05)

async def run_morning_briefing_worker(bot: Bot):
    today = get_local_date()
    async with AsyncSessionLocal() as session:
        teachers = (await session.execute(select(Teacher).where(Teacher.telegram_id.isnot(None)))).scalars().all()
        for tch in teachers:
            tch_u = await session.get(User, tch.telegram_id)
            lang = tch_u.language if tch_u else "tr"
            m = f"☀️ *GÜNAYDIN SAYIN HOCAM!*\n\nBugün *{today.strftime('%d.%m.%Y')}*. Lütfen ilk dersinizde yoklama almayı unutmayınız."
            await safe_send_message(bot, tch.telegram_id, m, parse_mode="Markdown")

async def run_friday_backup_worker(bot: Bot):
    try:
        buf = await export_all_school_data_excel()
        today_str = datetime.utcnow().strftime("%d_%m_%Y")
        file = BufferedInputFile(buf.read(), filename=f"Haftalik_Okul_Yedegi_{today_str}.xlsx")
        for a_id in ADMIN_IDS:
            try:
                await bot.send_document(chat_id=a_id, document=file, caption=f"📦 *Haftalık Otomatik Sistem Veri Yedeği* ({today_str})", parse_mode="Markdown")
                await asyncio.sleep(0.1)
            except Exception: pass
        buf.close()
    except Exception: pass

# ======================================================================
# 7. FSM DURUMLARI VE BELLEK ÖNBELLEKLERİ
# ======================================================================

class Form(StatesGroup):
    waiting_auth_code = State()
    req_role = State()
    req_name = State()
    req_gender = State()
    req_birth_date = State()
    req_phone = State()
    req_details = State()
    appr_st_class = State()
    appr_st_no = State()
    app_note = State()
    sched_update_text = State()
    menu_update_text = State()
    add_student_name = State()
    add_student_class = State()
    add_student_no = State()
    add_teacher_name = State()
    add_teacher_subject = State()
    waiting_search_query = State()
    waiting_search_teacher_query = State()
    waiting_search_user_query = State()
    edit_student_val = State()
    waiting_admin_dm_text = State()
    waiting_admin_tg_id = State()
    waiting_admin_name = State()
    waiting_broadcast_text = State()
    waiting_teacher_excel = State()
    waiting_restore_excel = State()
    grade_score = State()
    grade_badge = State()
    edit_grade_val = State()
    waiting_behavior_note = State()
    hw_content = State()
    waiting_hw_submission = State()
    waiting_hw_feedback = State()
    waiting_exam_subject = State()
    waiting_exam_date = State()
    waiting_emergency_text = State()
    parent_add_child_code = State()
    waiting_medical_photo = State()

ATTENDANCE_CACHE = {}
GRADE_CACHE = {}
BEHAVIOR_CACHE = {}
HW_CACHE = {}
SUBMISSION_CACHE = {}
EXAM_CACHE = {}
BC_CACHE = {}
APP_CACHE = {}
LAST_MENU_MSG_ID = {}
ACTIVE_CHAT_MESSAGES = {}
ADMIN_DISPATCHED_NOTIFS = {}
USER_COOLDOWN = {}
USER_REQUEST_LOG = {}

# ======================================================================
# 8. ROUTER: BAŞLANGIÇ, PROFİL VE KİMLİK DOĞRULAMA MOTORU
# ======================================================================

router = Router()

async def render_clean_dashboard(target: Message | CallbackQuery, user: User, chat_id: int | None = None):
    lang = user.language if user else "tr"
    role = user.role if user else "guest"
    target_chat_id = chat_id or (target.chat.id if isinstance(target, Message) else target.message.chat.id)

    last_mid = LAST_MENU_MSG_ID.pop(target_chat_id, None)
    if last_mid:
        try:
            bot_obj = target.bot if isinstance(target, Message) else target.message.bot
            await bot_obj.delete_message(chat_id=target_chat_id, message_id=last_mid)
        except Exception: pass

    role_kb = get_role_reply_kb(role, lang)

    if role == "admin":
        async with AsyncSessionLocal() as session:
            c_cnt = (await session.execute(select(func.count(Student.class_name.distinct())))).scalar() or 0
            s_cnt = (await session.execute(select(func.count(Student.id)))).scalar() or 0
            t_cnt = (await session.execute(select(func.count(Teacher.id)))).scalar() or 0
            req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
            med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0
        text = get_text("admin_title", lang) + "\\n\\n" + get_text("admin_stats", lang, c_cnt=c_cnt, s_cnt=s_cnt, t_cnt=t_cnt, req_cnt=req_cnt, med_cnt=med_cnt, date=get_local_date().strftime("%d.%m.%Y"))
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"), InlineKeyboardButton(text=get_text("btn_teachers", lang), callback_data="adm:teachers")],
            [InlineKeyboardButton(text=get_text("btn_cockpit_unified", lang), callback_data="adm:cockpit"), InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_hub")],
            [InlineKeyboardButton(text=get_text("btn_requests", lang, count=req_cnt), callback_data="adm:requests_list"), InlineKeyboardButton(text=get_text("btn_medical", lang, count=med_cnt), callback_data="adm:medical_list")],
            [InlineKeyboardButton(text=get_text("btn_excel_hub", lang), callback_data="adm:excel_hub"), InlineKeyboardButton(text=get_text("btn_users_hub", lang), callback_data="adm:users_hub:0")]
        ]
        sent_m = await safe_edit_or_answer(target, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
        if sent_m: LAST_MENU_MSG_ID[target_chat_id] = sent_m.message_id
    elif role == "teacher":
        async with AsyncSessionLocal() as session:
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
            t_name = tch.full_name if tch else (user.full_name or "Öğretmen")
            t_subj = tch.subject if tch else "Branş"
        text = get_text("menu_teacher", lang, name=escape_md(t_name), subject=escape_md(t_subj))
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_attendance", lang), callback_data="tch:classes"), InlineKeyboardButton(text=get_text("btn_enter_grade", lang), callback_data="tch:grade_classes")],
            [InlineKeyboardButton(text=get_text("btn_homework_board", lang), callback_data="tch:hw_classes"), InlineKeyboardButton(text=get_text("rk_behavior", lang), callback_data="tch:behavior_classes")],
            [InlineKeyboardButton(text=get_text("rk_appointments", lang), callback_data="tch:appointments")]
        ]
        sent_m = await safe_edit_or_answer(target, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
        if sent_m: LAST_MENU_MSG_ID[target_chat_id] = sent_m.message_id
    elif role == "parent":
        async with AsyncSessionLocal() as session:
            st = await session.get(Student, user.current_child_id) if user.current_child_id else None
            s_name = st.full_name if st else "Öğrenci"
            c_name = st.class_name if st else ""
        text = get_text("menu_parent", lang, name=escape_md(s_name), class_name=escape_md(c_name))
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_report_card", lang), callback_data="act_view_report"), InlineKeyboardButton(text=get_text("btn_upload_medical", lang), callback_data="upload_medical_init")],
            [InlineKeyboardButton(text=get_text("btn_switch_student", lang), callback_data="parent:switch_student"), InlineKeyboardButton(text=get_text("rk_appointments", lang), callback_data="act_book_app")]
        ]
        sent_m = await safe_edit_or_answer(target, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
        if sent_m: LAST_MENU_MSG_ID[target_chat_id] = sent_m.message_id
    elif role == "student":
        async with AsyncSessionLocal() as session:
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            s_name = st.full_name if st else (user.full_name or "Öğrenci")
            c_name = st.class_name if st else ""
            s_no = st.student_number if st else ""
        text = get_text("menu_student", lang, name=escape_md(s_name), class_name=escape_md(c_name), no=escape_md(s_no))
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_report_card", lang), callback_data="act_view_report"), InlineKeyboardButton(text=get_text("rk_homework", lang), callback_data="act_view_hw")],
            [InlineKeyboardButton(text=get_text("rk_schedule", lang), callback_data="act_view_sched"), InlineKeyboardButton(text=get_text("rk_notices", lang), callback_data="act_view_notices")]
        ]
        sent_m = await safe_edit_or_answer(target, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
        if sent_m: LAST_MENU_MSG_ID[target_chat_id] = sent_m.message_id
    else:
        await prompt_guest_screen(target, user)

async def prompt_guest_screen(target: Message | CallbackQuery, user: User):
    lang = user.language if user else "tr"
    text = get_text("welcome_guest", lang)
    buttons = [
        [InlineKeyboardButton(text=get_text("btn_login_prompt", lang), callback_data="act_enter_code")],
        [InlineKeyboardButton(text=get_text("btn_req_access", lang), callback_data="act_req_access")],
        [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")]
    ]
    chat_id = target.chat.id if isinstance(target, Message) else target.message.chat.id
    sent_m = await safe_edit_or_answer(target, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    if sent_m: LAST_MENU_MSG_ID[chat_id] = sent_m.message_id

@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    try: await message.delete()
    except Exception: pass

    last_mid = LAST_MENU_MSG_ID.pop(message.chat.id, None)
    if last_mid:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=last_mid)
        except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            role = "admin" if user_id in ADMIN_IDS else "guest"
            user = User(telegram_id=user_id, language="tr", role=role, full_name=message.from_user.full_name, username=message.from_user.username)
            if role == "admin": user.admin_type = "permanent"
            session.add(user)
            await session.commit()
            sent_m = await message.answer(get_text("lang_select", "tr"), reply_markup=get_language_inline_kb())
            LAST_MENU_MSG_ID[message.chat.id] = sent_m.message_id
            return

        if user_id in ADMIN_IDS and user.role != "admin":
            user.role = "admin"
            user.admin_type = "permanent"
            await session.commit()

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language), parse_mode="Markdown")
            return

        now = datetime.utcnow()
        if user.locked_until and user.locked_until > now:
            mins_left = max(1, int((user.locked_until - now).total_seconds() / 60))
            await message.answer(get_text("lock_countdown_msg", user.language, mins=mins_left), parse_mode="Markdown")
            return

        if user.role != "guest":
            await render_clean_dashboard(message, user)
        else:
            await prompt_guest_screen(message, user)

@router.callback_query(F.data.startswith("set_lang:"))
async def cb_set_lang(query: CallbackQuery):
    lang_code = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if not user:
            role = "admin" if query.from_user.id in ADMIN_IDS else "guest"
            user = User(telegram_id=query.from_user.id, language=lang_code, role=role, full_name=query.from_user.full_name)
            session.add(user)
        else:
            user.language = lang_code
        await session.commit()

    # Eski dil kartını sohbette kirlilik kalmaması için anında siliyoruz
    try: await query.message.delete()
    except Exception: pass

    role_kb = get_role_reply_kb(user.role, user.language)
    await query.message.answer(get_text("lang_changed", user.language), reply_markup=role_kb, parse_mode="Markdown")
    await render_clean_dashboard(query.message, user)
    await query.answer()

@router.callback_query(F.data == "act_change_lang")
async def cb_change_lang_screen(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    try: await query.message.delete()
    except Exception: pass
    sent_m = await query.message.answer(get_text("lang_select", lang), reply_markup=get_language_inline_kb())
    LAST_MENU_MSG_ID[query.message.chat.id] = sent_m.message_id
    await query.answer()

def is_universal_cancel_text(text: str | None) -> bool:
    if not text: return False
    cleaned = text.strip().replace("İ", "i").replace("I", "ı").lower()
    cancel_keywords = ["iptal", "bekor", "отмена", "cancel", "vazgeç", "çıkış", "geri", "amalni bekor qilish", "işlem", "❌"]
    return any(k in cleaned for k in cancel_keywords)

@router.message(any_state, Command("cancel", "iptal", "bekor"))
@router.message(any_state, F.text.func(is_universal_cancel_text))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    try: await message.delete()
    except Exception: pass

    last_mid = LAST_MENU_MSG_ID.pop(message.chat.id, None)
    if last_mid:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=last_mid)
        except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        role = user.role if user else "guest"

    role_kb = get_role_reply_kb(role, lang)
    m = await message.answer(get_text("action_cancelled", lang), reply_markup=role_kb, parse_mode="Markdown")
    LAST_MENU_MSG_ID[message.chat.id] = m.message_id
    if user:
        await render_clean_dashboard(message, user)

@router.message(any_state, Command("restart", "reset", "cikis", "logout"))
@router.message(any_state, F.text.in_(["🔄 Yeniden Başlat", "🔄 Перезапуск", "🔄 Qayta ishga tushirish", "🔄 Restart Bot", "🚪 Çıkış Yap", "🚪 Выйти", "🚪 Chiqish", "🚪 Log Out"]))
async def handle_bot_restart_cmd(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id
    try: await message.delete()
    except Exception: pass

    last_mid = LAST_MENU_MSG_ID.pop(message.chat.id, None)
    if last_mid:
        try: await message.bot.delete_message(chat_id=message.chat.id, message_id=last_mid)
        except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if user and is_admin_user(user, user_id):
            user.role = "admin"
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            await message.answer(get_text("admin_restart_confirmed", user.language), reply_markup=get_role_reply_kb("admin", user.language), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        if user:
            user.role = "guest"
            user.current_child_id = None
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
        else:
            user = User(telegram_id=user_id, role="guest", language="tr", full_name=message.from_user.full_name)
            session.add(user)
            await session.commit()

    await message.answer(get_text("lang_select", "tr"), reply_markup=get_language_inline_kb())

@router.callback_query(F.data == "act_enter_code")
async def cb_act_enter_code(query: CallbackQuery, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    cancel_kb = get_cancel_reply_kb(lang)
    await safe_edit_or_answer(query, get_text("prompt_enter_code_direct", lang), reply_markup=cancel_kb, parse_mode="Markdown")
    await state.set_state(Form.waiting_auth_code)
    await query.answer()

@router.message(Form.waiting_auth_code)
async def handle_auth_code_fsm(message: Message, state: FSMContext):
    code_text = message.text.strip()
    await state.clear()
    await process_auth_code_string(code_text, message.from_user.id, message, state)

async def process_auth_code_string(code_text: str, user_id: int, message: Message, state: FSMContext):
    clean_code = normalize_code(code_text)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            user = User(telegram_id=user_id, language="tr", role="guest", full_name=message.from_user.full_name)
            session.add(user)
            await session.commit()
        lang = user.language

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", lang), parse_mode="Markdown")
            return

        now = datetime.utcnow()
        if user.locked_until and user.locked_until > now:
            mins_left = max(1, int((user.locked_until - now).total_seconds() / 60))
            await message.answer(get_text("lock_countdown_msg", lang, mins=mins_left), parse_mode="Markdown")
            return

        # 1. Admin Kodu
        if clean_code == ADMIN_CODE:
            user.role = "admin"
            user.admin_type = "permanent"
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            await message.answer(get_text("auth_success", lang, name=escape_md(user.full_name or "Yönetici"), role="Yönetici"), reply_markup=get_role_reply_kb("admin", lang), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        adm_key = (await session.execute(select(AdminKey).where(AdminKey.code == clean_code, AdminKey.is_used == False))).scalar_one_or_none()
        if adm_key:
            adm_key.is_used = True
            user.role = "admin"
            user.admin_type = "permanent"
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            await message.answer(get_text("auth_success", lang, name=escape_md(user.full_name or "Yönetici"), role="Yönetici"), reply_markup=get_role_reply_kb("admin", lang), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # 2. Öğretmen Kodu (HCA)
        tch = (await session.execute(select(Teacher).where(func.upper(func.trim(Teacher.auth_code)) == clean_code))).scalar_one_or_none()
        if tch:
            if tch.is_code_burned and tch.telegram_id and tch.telegram_id != user_id:
                await message.answer(get_text("auth_code_already_linked", lang), parse_mode="Markdown")
                return
            tch.telegram_id = user_id
            tch.is_code_burned = True
            user.role = "teacher"
            user.full_name = tch.full_name
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            await message.answer(get_text("auth_success", lang, name=escape_md(tch.full_name), role="Öğretmen"), reply_markup=get_role_reply_kb("teacher", lang), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # 3. Öğrenci Kodu (OGR)
        st_student = (await session.execute(select(Student).where(func.upper(func.trim(Student.student_code)) == clean_code))).scalar_one_or_none()
        if st_student:
            if st_student.is_student_code_burned and st_student.student_telegram_id and st_student.student_telegram_id != user_id:
                await message.answer(get_text("auth_code_already_linked", lang), parse_mode="Markdown")
                return
            st_student.student_telegram_id = user_id
            st_student.is_student_code_burned = True
            user.role = "student"
            user.full_name = st_student.full_name
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            await message.answer(get_text("auth_success", lang, name=escape_md(st_student.full_name), role="Öğrenci"), reply_markup=get_role_reply_kb("student", lang), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # 4. Veli Kodu (VELI)
        p_student = (await session.execute(select(Student).where(func.upper(func.trim(Student.parent_code)) == clean_code))).scalar_one_or_none()
        if p_student:
            rel = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == user_id, ParentStudent.student_id == p_student.id))).scalar_one_or_none()
            if not rel:
                session.add(ParentStudent(parent_telegram_id=user_id, student_id=p_student.id))
            user.role = "parent"
            user.current_child_id = p_student.id
            p_student.is_parent_code_burned = True
            user.failed_attempts = 0
            user.locked_until = None
            await session.commit()
            await message.answer(get_text("auth_success", lang, name=escape_md(user.full_name or "Veli"), role="Veli"), reply_markup=get_role_reply_kb("parent", lang), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # Başarısız Kod
        user.failed_attempts += 1
        rem = max(0, 5 - user.failed_attempts)
        if user.failed_attempts >= 5:
            user.locked_until = datetime.utcnow() + timedelta(hours=1)
            await session.commit()
            await message.answer(get_text("auth_locked", lang), parse_mode="Markdown")
        else:
            await session.commit()
            await message.answer(get_text("auth_failed", lang, remaining=rem), parse_mode="Markdown")

# ======================================================================
# 9. ERİŞİM TALEBİ VE ŞİFRE BAŞVURU MOTORU
# ======================================================================

@router.callback_query(F.data == "act_req_access")
async def cb_act_req_access(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        pending = (await session.execute(select(AccessRequest).where(AccessRequest.telegram_id == query.from_user.id, AccessRequest.status == "pending"))).scalar_one_or_none()
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
    await query.message.answer(get_text("req_role_select", lang), reply_markup=role_kb, parse_mode=None)
    await state.set_state(Form.req_role)
    await query.answer()

@router.message(Form.req_role)
async def process_req_role(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    txt = message.text.strip().lower()
    role_code = None
    if any(k in txt for k in ["öğretmen", "учитель", "o'qituvchi", "teacher"]): role_code = "teacher"
    elif any(k in txt for k in ["veli", "родитель", "ota-ona", "parent"]): role_code = "parent"
    elif any(k in txt for k in ["öğrenci", "ученик", "o'quvchi", "student"]): role_code = "student"

    if not role_code:
        await message.answer("⚠️ " + get_text("req_role_select", lang))
        return

    await state.update_data(role=role_code)
    await message.answer(get_text("req_name_prompt", lang), reply_markup=get_cancel_reply_kb(lang), parse_mode=None)
    await state.set_state(Form.req_name)

@router.message(Form.req_name)
async def process_req_name(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    name_val = message.text.strip()
    if len(name_val.split()) < 2 or len(name_val) < 4:
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
    await message.answer(get_text("prompt_req_gender", lang), reply_markup=gender_kb, parse_mode=None)
    await state.set_state(Form.req_gender)

@router.message(Form.req_gender)
async def process_req_gender(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    txt = message.text.strip().lower()
    gender = "Erkek" if any(k in txt for k in ["erkek", "мужской", "erkak", "male"]) else ("Kadın" if any(k in txt for k in ["kadın", "женский", "ayol", "female"]) else None)
    if not gender:
        await message.answer(get_text("err_invalid_gender_strict", lang))
        return

    await state.update_data(gender=gender)
    await message.answer(get_text("prompt_req_birth_date", lang), reply_markup=get_cancel_reply_kb(lang), parse_mode=None)
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
    await message.answer(get_text("req_phone_prompt", lang), reply_markup=phone_kb, parse_mode=None)
    await state.set_state(Form.req_phone)

@router.message(Form.req_phone)
async def process_req_phone(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    phone_val = str(message.contact.phone_number) if (message.contact and message.contact.phone_number) else message.text.strip()
    await state.update_data(phone=phone_val)
    data = await state.get_data()
    role = data.get("role", "student")

    prompt_key = "req_details_teacher" if role == "teacher" else ("req_details_parent" if role == "parent" else "req_details_student")
    await message.answer(get_text(prompt_key, lang), reply_markup=get_cancel_reply_kb(lang), parse_mode=None)
    await state.set_state(Form.req_details)

@router.message(Form.req_details)
async def process_req_details(message: Message, state: FSMContext):
    details_text = message.text.strip()
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

        req = AccessRequest(telegram_id=message.from_user.id, role=role, full_name=full_name, gender=gender, birth_date=birth_date, phone=phone_val, details=details_text, status="pending")
        session.add(req)
        await session.commit()

        await message.answer(get_text("req_sent_success", lang), reply_markup=get_role_reply_kb("guest", lang), parse_mode=None)
        await prompt_guest_screen(message, user)

        adm_msg = (
            f"🛎️ <b>YENİ ERİŞİM / ŞİFRE BAŞVURUSU (#{req.id})</b>\n\n"
            f"• <b>Ad Soyad:</b> {escape_html(full_name)}\n"
            f"• <b>Cinsiyet:</b> {gender} | <b>D. Tarihi:</b> {birth_date}\n"
            f"• <b>Telefon:</b> {phone_val}\n"
            f"• <b>Rol:</b> {role}\n"
            f"• <b>Detay:</b> {escape_html(details_text)}"
        )
        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Onayla", callback_data=f"adm:appr_req:{req.id}"), InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_req:{req.id}")]
        ])
        for a_id in ADMIN_IDS:
            await safe_send_message(message.bot, a_id, adm_msg, reply_markup=adm_kb, parse_mode="HTML")

@router.callback_query(F.data == "adm:requests_list")
async def cb_admin_requests_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        requests = (await session.execute(select(AccessRequest).where(AccessRequest.status == "pending").order_by(desc(AccessRequest.created_at)))).scalars().all()
        if not requests:
            await safe_edit_or_answer(query, get_text("no_pending_requests", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return
        buttons = []
        for r in requests:
            buttons.append([InlineKeyboardButton(text=f"🛎️ {r.full_name} ({r.role})", callback_data=f"adm:view_req:{r.id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("pending_requests_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:view_req:"))
async def cb_admin_view_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer(get_text("request_already_handled", lang), show_alert=True)
            return

        text = f"🛎️ <b>Başvuru (#{req.id})</b>\n\n• Ad Soyad: {escape_html(req.full_name)}\n• Telefon: {req.phone}\n• Rol: {req.role}\n• Detay: {escape_html(req.details)}"
        buttons = [
            [InlineKeyboardButton(text="✅ Onayla", callback_data=f"adm:appr_req:{req.id}"), InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_req:{req.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:requests_list")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:appr_req:"))
async def cb_admin_approve_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        lang = admin_user.language if admin_user else "tr"
        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending": return

        req.status = "approved"
        req.reviewed_by = query.from_user.id
        req.reviewed_at = datetime.utcnow()

        target_u = await session.get(User, req.telegram_id)
        if not target_u:
            target_u = User(telegram_id=req.telegram_id, language="tr")
            session.add(target_u)
        target_u.role = req.role
        target_u.full_name = req.full_name

        if req.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.full_name == req.full_name))).scalar_one_or_none()
            if not tch:
                session.add(Teacher(full_name=req.full_name, subject=req.details, auth_code=generate_secure_code("HCA"), telegram_id=req.telegram_id, is_code_burned=True, assigned_classes="ALL"))
            else:
                tch.telegram_id = req.telegram_id
                tch.is_code_burned = True
        elif req.role == "student":
            st = (await session.execute(select(Student).where(Student.full_name == req.full_name))).scalar_one_or_none()
            if not st:
                st = Student(full_name=req.full_name, class_name="9-A", student_number="101", student_code=generate_secure_code("OGR"), parent_code=generate_secure_code("VELI"), is_student_code_burned=True, student_telegram_id=req.telegram_id)
                session.add(st)
            else:
                st.student_telegram_id = req.telegram_id
                st.is_student_code_burned = True

        await session.commit()
        await safe_send_message(query.message.bot, req.telegram_id, get_text("req_approved_user", target_u.language, role=req.role), reply_markup=get_role_reply_kb(req.role, target_u.language), parse_mode="Markdown")
        await safe_edit_or_answer(query, get_text("req_approved_admin_msg", lang, id=req.id, name=escape_html(req.full_name), role=req.role))
    await query.answer()

@router.callback_query(F.data.startswith("adm:rej_req:"))
async def cb_admin_reject_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        req = await session.get(AccessRequest, req_id)
        if req and req.status == "pending":
            req.status = "rejected"
            await session.commit()
            t_u = await session.get(User, req.telegram_id)
            await safe_send_message(query.message.bot, req.telegram_id, get_text("req_rejected_user", t_u.language if t_u else "tr"), parse_mode="Markdown")
            await safe_edit_or_answer(query, get_text("req_rejected_admin_msg", lang, id=req.id, name=escape_html(req.full_name)))
    await query.answer()

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
        await safe_edit_or_answer(query, get_text("select_teacher_appointment", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
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
        app = Appointment(parent_telegram_id=message.from_user.id, teacher_id=tch_id, preferred_time=note_text, note=note_text, status="pending")
        session.add(app)
        await session.commit()
        await message.answer(get_text("appointment_sent", lang))
        await render_clean_dashboard(message, user)

@router.callback_query(F.data == "tch:appointments")
async def cb_teacher_appointments_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        tch_id = tch.id if tch else 0
        apps = (await session.execute(select(Appointment).where(Appointment.teacher_id == tch_id, Appointment.status == "pending"))).scalars().all()
        if not apps:
            await safe_edit_or_answer(query, get_text("no_pending_appointments", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return
        buttons = []
        for a in apps:
            buttons.append([InlineKeyboardButton(text=f"🤝 #{a.id} ({a.preferred_time[:20]})", callback_data=f"tch:view_app:{a.id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("pending_appointments_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("tch:view_app:"))
async def cb_teacher_view_appointment(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        app = await session.get(Appointment, app_id)
        if not app: return
        p_u = await session.get(User, app.parent_telegram_id)
        p_name = p_u.full_name if p_u else "Veli"
        text = f"🤝 <b>Veli Randevu Talebi (#{app.id})</b>\n\n• Veli: {p_name}\n• Zaman / Not: {escape_html(app.preferred_time)}"
        buttons = [
            [InlineKeyboardButton(text="✅ Kabul Et", callback_data=f"tch:appr_app:{app.id}"), InlineKeyboardButton(text="❌ Reddet", callback_data=f"tch:rej_app:{app.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:appointments")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("tch:appr_app:"))
async def cb_teacher_approve_app(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        app = await session.get(Appointment, app_id)
        if app:
            app.status = "approved"
            await session.commit()
            p_u = await session.get(User, app.parent_telegram_id)
            await safe_send_message(query.message.bot, app.parent_telegram_id, get_text("appointment_approved_msg", p_u.language if p_u else "tr"))
            await query.answer(get_text("appointment_confirmed_toast", "tr"), show_alert=True)
            await cb_teacher_appointments_list(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("tch:rej_app:"))
async def cb_teacher_reject_app(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        app = await session.get(Appointment, app_id)
        if app:
            app.status = "rejected"
            await session.commit()
            p_u = await session.get(User, app.parent_telegram_id)
            await safe_send_message(query.message.bot, app.parent_telegram_id, get_text("appointment_rejected_msg", p_u.language if p_u else "tr"))
            await cb_teacher_appointments_list(query)
            return
    await query.answer()

@router.callback_query(F.data == "act_view_sched")
async def cb_view_schedule(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        cls_name = "9-A"
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: cls_name = st.class_name
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == query.from_user.id))).scalar_one_or_none()
            if st: cls_name = st.class_name

        sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == cls_name))).scalar_one_or_none()
        content = sched.schedule_text if sched else "• 1. Ders: 09:00 - Matematik\n• 2. Ders: 09:50 - Fizik\n• 3. Ders: 10:40 - Türkçe"
        text = f"📅 <b>{escape_html(cls_name)} Ders Programı:</b>\n\n{escape_html(content)}"
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "act_view_cafe")
async def cb_view_cafeteria(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = get_local_date()
        menu = (await session.execute(select(CafeteriaMenu).where(CafeteriaMenu.date == today))).scalar_one_or_none()
        content = menu.menu_text if menu else "🍲 Mercimek Çorbası\n🍗 Tavuk & Pilav\n🥗 Mevsim Salatası"
        text = f"🍲 <b>Günün Yemek Menüsü ({today.strftime('%d.%m.%Y')}):</b>\n\n{escape_html(content)}"
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "act_view_notices")
async def cb_view_notices(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        notices = (await session.execute(select(BroadcastNotice).order_by(desc(BroadcastNotice.created_at)).limit(5))).scalars().all()
        if not notices:
            text = "📢 Henüz yayınlanmış bir duyuru bulunmamaktadır."
        else:
            text = "📢 <b>Okul Duyuruları:</b>\n\n" + "\n\n".join([f"📌 <i>{n.created_at.strftime('%d.%m %H:%M')}</i>\n{escape_html(n.content)}" for n in notices])
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:blacklist")
async def cb_admin_blacklist(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        now = datetime.utcnow()
        blocked_users = (await session.execute(select(User).where((User.is_blacklisted == True) | (User.locked_until > now)))).scalars().all()
        if not blocked_users:
            await safe_edit_or_answer(query, get_text("no_blacklisted", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return
        buttons = []
        for u in blocked_users:
            buttons.append([InlineKeyboardButton(text=f"🟢 Engeli Kaldır: {u.full_name or u.telegram_id}", callback_data=f"adm:unban:{u.telegram_id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("blacklisted_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("adm:unban:"))
async def cb_admin_unban(query: CallbackQuery):
    t_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        target_u = await session.get(User, t_id)
        if target_u:
            target_u.is_blacklisted = False
            target_u.failed_attempts = 0
            target_u.locked_until = None
            await session.commit()
            await query.answer(get_text("unban_success", user.language if user else "tr"), show_alert=True)
            await cb_admin_blacklist(query)
            return
    await query.answer()

# ======================================================================
# 11. YÖNETİCİ İŞLEMLERİ (ÖĞRENCİ, ÖĞRETMEN, SINIFLAR & KARTLAR)
# ======================================================================

@router.callback_query(F.data == "adm:classes")
async def cb_classes_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            await safe_edit_or_answer(query, get_text("no_classes_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return
        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:show_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(query, "🏫 <b>Sınıflar Listesi:</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
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
        all_st = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()
        total_pages = max(1, (len(all_st) + PAGE_SIZE - 1) // PAGE_SIZE)
        paged_st = all_st[page * PAGE_SIZE : (page + 1) * PAGE_SIZE]

        buttons = []
        for s in paged_st:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"adm:st_card:{s.id}")])

        if total_pages > 1:
            nav_row = []
            if page > 0: nav_row.append(InlineKeyboardButton(text="⬅️", callback_data=f"adm:show_class:{class_name}:{page - 1}"))
            nav_row.append(InlineKeyboardButton(text=f"{page + 1}/{total_pages}", callback_data="noop"))
            if page < total_pages - 1: nav_row.append(InlineKeyboardButton(text="➡️", callback_data=f"adm:show_class:{class_name}:{page + 1}"))
            buttons.append(nav_row)

        buttons.append([InlineKeyboardButton(text=f"📄 {class_name} Şifre Kartları (PDF)", callback_data=f"adm:gen_pdf:{class_name}")])
        buttons.append([InlineKeyboardButton(text="⬅️ Geri", callback_data="adm:classes")])
        text = f"🏫 <b>{escape_html(class_name)} Sınıfı Listesi ({len(all_st)} Öğrenci):</b>"
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:st_card:"))
async def cb_student_card(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st: return

        text = (
            f"🧑‍🎓 <b>ÖĞRENCİ BİLGİ KARTI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"• <b>Ad Soyad:</b> {escape_html(st.full_name)}\n"
            f"• <b>Sınıf:</b> {escape_html(st.class_name)} | <b>No:</b> {escape_html(st.student_number)}\n"
            f"• <b>Öğrenci Kodu:</b> <code>{st.student_code}</code>\n"
            f"• <b>Veli Kodu:</b> <code>{st.parent_code}</code>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )
        buttons = [
            [InlineKeyboardButton(text="🔄 Kodları Sıfırla", callback_data=f"adm:confirm_reset_st:{st.id}")],
            [InlineKeyboardButton(text="👨‍👩‍👧‍👦 Veli Bağlantısını Kopar", callback_data=f"adm:unlink_pr:{st.id}")],
            [InlineKeyboardButton(text="🗑️ Öğrenciyi Sil", callback_data=f"adm:confirm_del_st:{st.id}")],
            [InlineKeyboardButton(text="⬅️ Geri", callback_data=f"adm:show_class:{st.class_name}")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_reset_st:"))
async def cb_reset_student_codes(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        st = await session.get(Student, st_id)
        if st:
            st.student_code = generate_secure_code("OGR")
            st.parent_code = generate_secure_code("VELI")
            st.is_student_code_burned = False
            st.is_parent_code_burned = False
            st.student_telegram_id = None
            await session.commit()
            await query.answer("✅ Kodlar yenilendi!", show_alert=True)
            await cb_student_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:unlink_pr:"))
async def cb_admin_unlink_parent(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        st = await session.get(Student, st_id)
        if st:
            await session.execute(delete(ParentStudent).where(ParentStudent.student_id == st.id))
            st.parent_code = generate_secure_code("VELI")
            st.is_parent_code_burned = False
            await session.commit()
            await query.answer("✅ Veli bağlantısı koparıldı!", show_alert=True)
            await cb_student_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_del_st:"))
async def cb_delete_student_confirmed(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        st = await session.get(Student, st_id)
        if st:
            cls = st.class_name
            await session.execute(delete(ParentStudent).where(ParentStudent.student_id == st.id))
            await session.execute(delete(Attendance).where(Attendance.student_id == st.id))
            await session.execute(delete(Grade).where(Grade.student_id == st.id))
            await session.delete(st)
            await session.commit()
            await query.answer("Öğrenci silindi.", show_alert=True)
            query.data = f"adm:show_class:{cls}"
            await cb_show_class_students(query)
            return
    await query.answer()

@router.callback_query(F.data == "adm:teachers")
async def cb_admin_teachers_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        buttons = []
        for t in teachers:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {t.full_name} ({t.subject})", callback_data=f"adm:tch_card:{t.id}")])
        buttons.append([InlineKeyboardButton(text="➕ Yeni Öğretmen Ekle", callback_data="adm:add_teacher")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(query, "👨‍🏫 <b>Öğretmenler Listesi:</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:tch_card:"))
async def cb_admin_teacher_card(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        tch = await session.get(Teacher, tch_id)
        if not tch: return
        status_str = f"🟢 Bağlı (<code>{tch.telegram_id}</code>)" if tch.telegram_id else "⚪ Giriş Yapmadı"
        text = (
            f"👨‍🏫 <b>ÖĞRETMEN BİLGİ KARTI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"• <b>İsim:</b> {escape_html(tch.full_name)}\n"
            f"• <b>Branş:</b> {escape_html(tch.subject)}\n"
            f"• <b>Giriş Kodu:</b> <code>{tch.auth_code}</code>\n"
            f"• <b>Durum:</b> {status_str}\n"
            f"• <b>Sınıflar:</b> {escape_html(tch.assigned_classes or 'ALL')}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )
        buttons = [
            [InlineKeyboardButton(text="🔄 Kodu Sıfırla", callback_data=f"adm:confirm_reset_tch:{tch.id}")],
            [InlineKeyboardButton(text="🗑️ Öğretmeni Sil", callback_data=f"adm:confirm_del_tch:{tch.id}")],
            [InlineKeyboardButton(text="⬅️ Geri", callback_data="adm:teachers")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_reset_tch:"))
async def cb_admin_reset_tch_code(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        tch = await session.get(Teacher, tch_id)
        if tch:
            tch.auth_code = generate_secure_code("HCA")
            tch.is_code_burned = False
            tch.telegram_id = None
            await session.commit()
            await query.answer("✅ Öğretmen kodu sıfırlandı!", show_alert=True)
            await cb_admin_teacher_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:confirm_del_tch:"))
async def cb_admin_delete_teacher_confirmed(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        tch = await session.get(Teacher, tch_id)
        if tch:
            await session.delete(tch)
            await session.commit()
            await query.answer("Öğretmen silindi.", show_alert=True)
            await cb_admin_teachers_list(query)
            return
    await query.answer()

@router.callback_query(F.data == "noop")
async def cb_noop(query: CallbackQuery):
    await query.answer()

# ======================================================================
# 12. KATEGORİ HUB'LARI VE RAPORLAR MASASI
# ======================================================================

@router.callback_query(F.data == "adm:cat_staff")
@router.message(any_state, F.text.in_(["👥 Kadro & Öğrenci", "👥 Ученики и учителя", "👥 Kadro va o'quvchilar", "👥 Staff & Students"]))
async def cb_cat_staff(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id): return
        lang = user.language if user else "tr"
        title = "👥 <b>KADRO VE ÖĞRENCİ YÖNETİMİ</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nLütfen işlem seçiniz:"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"), InlineKeyboardButton(text=get_text("btn_teachers", lang), callback_data="adm:teachers")],
            [InlineKeyboardButton(text=get_text("btn_school_admins", lang), callback_data="adm:admins_list"), InlineKeyboardButton(text=get_text("btn_users_hub", lang), callback_data="adm:users_hub:0")],
            [InlineKeyboardButton(text=get_text("btn_class_promotion", lang), callback_data="adm:class_promotion_init")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery): await event.answer()

@router.callback_query(F.data == "adm:cat_reports")
@router.message(any_state, F.text.in_(["📊 Raporlar & Denetim", "📊 Отчеты и контроль", "📊 Hisobotlar va nazorat", "📊 Reports & Audits"]))
async def cb_cat_reports(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id): return
        lang = user.language if user else "tr"
        title = "📊 <b>RAPORLAR VE AKADEMİK DENETİM</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nLütfen incelemek istediğiniz raporu seçiniz:"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_cockpit_unified", lang), callback_data="adm:cockpit")],
            [InlineKeyboardButton(text=get_text("btn_unack_notifs", lang), callback_data="adm:unack_notifs"), InlineKeyboardButton(text=get_text("btn_emergency_monitor", lang), callback_data="adm:emergency_monitor")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery): await event.answer()

@router.callback_query(F.data == "adm:cat_requests")
@router.message(any_state, F.text.in_(["🛎️ Onay Masası", "🛎️ Центр одобрений", "🛎️ Tasdiqlash markazi", "🛎️ Approval Center"]))
async def cb_cat_requests(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id): return
        lang = user.language if user else "tr"
        req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
        med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0
        title = f"🛎️ <b>BAŞVURU VE ONAY MASASI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n• Bekleyen Şifre Başvurusu: {req_cnt}\n• Bekleyen Sağlık Raporu: {med_cnt}"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_requests", lang, count=req_cnt), callback_data="adm:requests_list")],
            [InlineKeyboardButton(text=get_text("btn_medical", lang, count=med_cnt), callback_data="adm:medical_list")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery): await event.answer()

@router.callback_query(F.data == "adm:cat_tools")
@router.message(any_state, F.text.in_(["🛠️ İdari Araçlar", "🛠️ Инструменты", "🛠️ Boshqaruv vositalari", "🛠️ Admin Tools"]))
async def cb_cat_tools(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id): return
        lang = user.language if user else "tr"
        title = "🛠️ <b>İDARİ YÖNETİM ARAÇLARI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nLütfen işlem seçiniz:"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_hub"), InlineKeyboardButton(text=get_text("btn_emergency_alert", lang), callback_data="adm:emergency_init")],
            [InlineKeyboardButton(text=get_text("btn_excel_hub", lang), callback_data="adm:excel_hub"), InlineKeyboardButton(text=get_text("btn_pdf", lang), callback_data="adm:pdf_menu")],
            [InlineKeyboardButton(text=get_text("btn_cafeteria_edit", lang), callback_data="adm:menu_edit"), InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery): await event.answer()

@router.callback_query(F.data == "adm:cat_settings")
@router.message(any_state, F.text.in_(["⚙️ Sistem & Ayarlar", "⚙️ Настройки системы", "⚙️ Tizim va sozlamalar", "⚙️ System & Settings"]))
async def cb_cat_settings(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not is_admin_user(user, user_id): return
        lang = user.language if user else "tr"
        maint = await session.get(SystemSetting, "maintenance_mode")
        is_maint = maint.value == "true" if maint else False
        m_status = "AÇIK" if is_maint else "KAPALI"

        ro_setting = await session.get(SystemSetting, "readonly_mode")
        is_ro = ro_setting.value == "true" if ro_setting else False
        ro_status = "AÇIK" if is_ro else "KAPALI"

        title = "⚙️ <b>SİSTEM VE GÜVENLİK AYARLARI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\nLütfen işlem seçiniz:"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_change_admin_pin", lang), callback_data="adm:change_pin_init")],
            [InlineKeyboardButton(text=get_text("btn_maintenance_toggle", lang, status=m_status), callback_data="adm:toggle_maint"), InlineKeyboardButton(text=get_text("btn_toggle_readonly", lang, status=ro_status), callback_data="adm:toggle_readonly")],
            [InlineKeyboardButton(text=get_text("btn_blacklist", lang), callback_data="adm:blacklist"), InlineKeyboardButton(text=get_text("btn_export_all_data", lang), callback_data="adm:export_all_excel")],
            [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang"), InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(event, title, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    if isinstance(event, CallbackQuery): await event.answer()

@router.callback_query(F.data == "adm:toggle_readonly")
async def cb_admin_toggle_readonly(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        setting = await session.get(SystemSetting, "readonly_mode")
        if not setting:
            setting = SystemSetting(key="readonly_mode", value="true")
            session.add(setting)
        else:
            setting.value = "false" if setting.value == "true" else "true"
        await session.commit()
        SETTINGS_CACHE["readonly_mode"] = setting.value
        await query.answer(get_text("readonly_mode_updated", lang), show_alert=True)
        await cb_cat_settings(query)

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
        await query.answer(get_text("maintenance_mode_updated", lang), show_alert=True)
        await cb_cat_settings(query)

# ======================================================================
# 13. KULLANICI REHBERİ, YÖNETİCİ KADROSU VE DUYURU MASASI
# ======================================================================

@router.callback_query(F.data.startswith("adm:users_hub:"))
async def cb_admin_users_hub(query: CallbackQuery):
    page = int(query.data.split(":")[2])
    per_page = 8
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        total = (await session.execute(select(func.count(User.telegram_id)))).scalar() or 0
        total_pages = max(1, (total + per_page - 1) // per_page)
        users = (await session.execute(select(User).order_by(desc(User.created_at)).offset(page * per_page).limit(per_page))).scalars().all()

        buttons = []
        for u in users:
            name = u.full_name or f"ID: {u.telegram_id}"
            buttons.append([InlineKeyboardButton(text=f"👤 {name} [{u.role.upper()}]", callback_data=f"adm:user_card:{u.telegram_id}")])

        nav = []
        if page > 0: nav.append(InlineKeyboardButton(text="⬅️", callback_data=f"adm:users_hub:{page - 1}"))
        if (page + 1) * per_page < total: nav.append(InlineKeyboardButton(text="➡️", callback_data=f"adm:users_hub:{page + 1}"))
        if nav: buttons.append(nav)
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(query, f"👥 <b>Kullanıcı Rehberi ({total}):</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:user_card:"))
async def cb_admin_user_card(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        target_u = await session.get(User, tg_id)
        if not target_u: return

        text = (
            f"👤 <b>KULLANICI PROFİL KARTI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"• <b>Telegram ID:</b> <code>{target_u.telegram_id}</code>\n"
            f"• <b>İsim:</b> {escape_html(target_u.full_name or 'Bilinmiyor')}\n"
            f"• <b>Rol:</b> <b>{target_u.role.upper()}</b>\n"
            f"• <b>Durum:</b> {'🚫 Engelli' if target_u.is_blacklisted else '🟢 Aktif'}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )
        buttons = []
        if tg_id not in PERMANENT_ADMIN_IDS:
            if target_u.is_blacklisted:
                buttons.append([InlineKeyboardButton(text="🟢 Engeli Kaldır", callback_data=f"adm:unban_u:{tg_id}")])
            else:
                buttons.append([InlineKeyboardButton(text="🚫 Engelle (Ban)", callback_data=f"adm:ban_u:{tg_id}")])
        buttons.append([InlineKeyboardButton(text="⬅️ Kullanıcılar Listesi", callback_data="adm:users_hub:0")])
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:ban_u:"))
async def cb_admin_ban_user(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        target_u = await session.get(User, tg_id)
        if target_u and tg_id not in PERMANENT_ADMIN_IDS:
            target_u.is_blacklisted = True
            await session.commit()
            await query.answer("Kullanıcı engellendi.", show_alert=True)
            await cb_admin_user_card(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:unban_u:"))
async def cb_admin_unban_user(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.is_blacklisted = False
            await session.commit()
            await query.answer("Engeli kaldırıldı.", show_alert=True)
            await cb_admin_user_card(query)
            return
    await query.answer()

@router.callback_query(F.data == "adm:admins_list")
async def cb_admin_admins_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        buttons = []
        for a_id in PERMANENT_ADMIN_IDS:
            u_obj = await session.get(User, a_id)
            name = u_obj.full_name if u_obj and u_obj.full_name else f"Kurucu ID: {a_id}"
            buttons.append([InlineKeyboardButton(text=f"👑 {name}", callback_data=f"adm:user_card:{a_id}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))
        await safe_edit_or_answer(query, "👑 <b>Okul Yönetim Kadrosu:</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:broadcast_hub")
async def cb_admin_broadcast_hub(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    buttons = [
        [InlineKeyboardButton(text="🌐 Tüm Okula Duyur", callback_data="adm:bc_target:all")],
        [InlineKeyboardButton(text="👨‍🏫 Sadece Öğretmenler", callback_data="adm:bc_target:teachers"), InlineKeyboardButton(text="👨‍👩‍👧‍👦 Sadece Veliler", callback_data="adm:bc_target:parents")],
        get_nav_buttons(lang, back_callback="adm:cat_tools")
    ]
    await safe_edit_or_answer(query, get_text("broadcast_hub_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:bc_target:"))
async def cb_admin_bc_target_selected(query: CallbackQuery, state: FSMContext):
    target = query.data.split(":")[2]
    BC_CACHE[query.from_user.id] = target
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("prompt_broadcast_content", lang, target=target.upper()), reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ İptal", callback_data="adm:broadcast_hub")]]), parse_mode="Markdown")
    await state.set_state(Form.waiting_broadcast_text)
    await query.answer()

@router.message(Form.waiting_broadcast_text)
async def process_broadcast_text(message: Message, state: FSMContext):
    target = BC_CACHE.pop(message.from_user.id, "all")
    b_text = message.caption or message.text or ""
    await state.clear()
    if not b_text: return

    sent_cnt = 0
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if target == "all":
            u_ids = (await session.execute(select(User.telegram_id))).scalars().all()
        elif target == "teachers":
            u_ids = (await session.execute(select(User.telegram_id).where(User.role == "teacher"))).scalars().all()
        elif target == "parents":
            u_ids = (await session.execute(select(User.telegram_id).where(User.role == "parent"))).scalars().all()
        else:
            u_ids = (await session.execute(select(User.telegram_id))).scalars().all()

        notice = BroadcastNotice(content=b_text)
        session.add(notice)
        await session.commit()

        for uid in set(u_ids):
            res = await safe_send_message(message.bot, uid, f"📢 <b>OKUL DUYURUSU</b>\n\n{escape_html(b_text)}", parse_mode="HTML")
            if res: sent_cnt += 1
            await asyncio.sleep(0.04)

        await message.answer(get_text("broadcast_sent_report", lang, count=sent_cnt), reply_markup=get_role_reply_kb("admin", lang), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

# ======================================================================
# 14. ÖĞRETMEN, VELİ VE ÖĞRENCİ MODÜLLERİ (YOKLAMA, NOT, ÖDEV, KARNE)
# ======================================================================

@router.callback_query(F.data == "tch:classes")
async def cb_teacher_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all() or ["9-A"]
        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"att_class:{c}")])
        buttons.append([InlineKeyboardButton(text="🏠 Ana Menü", callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("attendance_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, get_text("attendance_intro", lang, class_name=escape_md(class_name)), reply_markup=kb, parse_mode="Markdown")
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

@router.callback_query(F.data.startswith("att_all_pres:"))
async def cb_attendance_all_present(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student).where(Student.class_name == class_name))).scalars().all()
        now = datetime.utcnow()
        att_date = get_local_date()
        for s in students:
            existing = (await session.execute(select(Attendance).where(Attendance.student_id == s.id, Attendance.date == att_date))).scalar_one_or_none()
            if existing:
                existing.status = "present"
            else:
                session.add(Attendance(student_id=s.id, class_name=class_name, date=att_date, status="present", teacher_id=query.from_user.id, notify_at=now + timedelta(minutes=15)))
        await session.commit()
        ATTENDANCE_CACHE.pop(query.from_user.id, None)
        await safe_edit_or_answer(query, f"✅ <b>{escape_html(class_name)}</b> sınıfı tümü mevcut kaydedildi!", reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("att_save:"))
async def cb_attendance_save(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        cache = ATTENDANCE_CACHE.pop(query.from_user.id, {})
        now = datetime.utcnow()
        att_date = get_local_date()
        for st_id, is_absent in cache.items():
            st_stat = "absent" if is_absent else "present"
            existing = (await session.execute(select(Attendance).where(Attendance.student_id == st_id, Attendance.date == att_date))).scalar_one_or_none()
            if existing:
                if existing.status != "excused":
                    existing.status = st_stat
                    existing.notify_at = now + timedelta(minutes=15)
                    existing.is_notified = False
            else:
                session.add(Attendance(student_id=st_id, class_name=class_name, date=att_date, status=st_stat, teacher_id=query.from_user.id, notify_at=now + timedelta(minutes=15)))
        await session.commit()
        await safe_edit_or_answer(query, get_text("att_saved", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
    await query.answer()

@router.callback_query(F.data == "tch:grade_classes")
async def cb_grade_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all() or ["9-A"]
        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"gr_cls:{c}")])
        buttons.append([InlineKeyboardButton(text="🏠 Ana Menü", callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("grade_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
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
        buttons.append([InlineKeyboardButton(text="⬅️ Geri", callback_data="tch:grade_classes")])
        await safe_edit_or_answer(query, get_text("grade_select_student", lang, class_name=escape_md(class_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("gr_st:"))
async def cb_grade_select_exam_type(query: CallbackQuery):
    st_id = int(query.data.split(":")[1])
    GRADE_CACHE[query.from_user.id] = {"student_id": st_id}
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        prompt = get_text("prompt_select_exam_type", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name))
        kb = get_exam_type_kb(st.class_name, st.id, lang=lang)
        await safe_edit_or_answer(query, prompt, reply_markup=kb, parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("gr_type:"))
async def cb_grade_enter_score(query: CallbackQuery, state: FSMContext):
    type_code = query.data.split(":")[1]
    exam_type = "1. Yazılı" if type_code == "1_yazili" else ("2. Yazılı" if type_code == "2_yazili" else "Sözlü / Performans")
    GRADE_CACHE[query.from_user.id]["exam_type"] = exam_type
    st_id = GRADE_CACHE[query.from_user.id].get("student_id")
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        text = get_text("prompt_grade_score", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name), exam_type=escape_md(exam_type))
        await safe_edit_or_answer(query, text, parse_mode="Markdown")
    await state.set_state(Form.grade_score)
    await query.answer()

@router.message(Form.grade_score)
async def process_grade_score(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
    try:
        score_val = float(message.text.strip().replace(",", "."))
        if not (0.0 <= score_val <= 100.0): raise ValueError
    except ValueError:
        await message.answer(get_text("invalid_score_range", lang))
        return

    GRADE_CACHE[message.from_user.id]["score"] = score_val
    buttons = [
        [InlineKeyboardButton(text="🟢 Başarılı / Övgü", callback_data="gr_bdg:🟢")],
        [InlineKeyboardButton(text="🟡 Eksik / Tekrar", callback_data="gr_bdg:🟡")],
        [InlineKeyboardButton(text="🔴 Uyarı / Dikkat", callback_data="gr_bdg:🔴")]
    ]
    await message.answer(get_text("prompt_grade_badge", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await state.set_state(Form.grade_badge)

@router.callback_query(F.data.startswith("gr_bdg:"))
async def cb_grade_save_final(query: CallbackQuery, state: FSMContext):
    badge_val = query.data.split(":")[1]
    data = GRADE_CACHE.pop(query.from_user.id, {})
    st_id = data.get("student_id")
    score_val = data.get("score", 100.0)
    exam_type_val = data.get("exam_type", "1. Yazılı")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        subject_name = tch.subject if tch else "Ders"
        st = await session.get(Student, st_id)

        grade = Grade(student_id=st_id, subject=subject_name, exam_type=exam_type_val, score=score_val, badge=badge_val, teacher_id=query.from_user.id)
        session.add(grade)
        await session.commit()

        parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
        for p_id in parents:
            p_u = await session.get(User, p_id)
            alert = get_text("grade_parent_notification", p_u.language if p_u else "tr", name=escape_md(st.full_name), subject=escape_md(subject_name), exam_type=escape_md(exam_type_val), score=score_val, badge=badge_val)
            await safe_send_message(query.message.bot, p_id, alert, parse_mode="Markdown")

        await safe_edit_or_answer(query, get_text("grade_saved_success", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
    await query.answer()

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
            await safe_edit_or_answer(query, get_text("no_linked_student", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return

        atts = (await session.execute(select(Attendance).where(Attendance.student_id == st.id))).scalars().all()
        abs_cnt = sum(1 for a in atts if a.status == "absent")
        exc_cnt = sum(1 for a in atts if a.status == "excused")
        grades = (await session.execute(select(Grade).where(Grade.student_id == st.id).order_by(desc(Grade.created_at)))).scalars().all()

        subj_map = {}
        for g in grades: subj_map.setdefault(g.subject, []).append(g.score)
        gr_lines = [f"• 📚 <b>{escape_md(s)}:</b> {round(sum(sc)/len(sc), 1)} / 100" for s, sc in subj_map.items()]
        gr_txt = "\n".join(gr_lines) if gr_lines else "<i>Henüz ders notu girilmemiş.</i>"

        text = (
            f"📊 <b>GELİŞİM VE NOT DURUM PANELİ</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"🧑‍🎓 <b>Öğrenci:</b> {escape_html(st.full_name)} ({escape_html(st.class_name)} - No: {st.student_number})\n\n"
            f"📌 <b>Devamsızlık Durumu:</b>\n"
            f"• Özürsüz: <b>{abs_cnt} Gün</b> (Yasal Sınır: 10)\n• Mazeretli / İzinli: <b>{exc_cnt} Gün</b>\n\n"
            f"📝 <b>Ders Notları:</b>\n{gr_txt}\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )
        buttons = [
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
        st = await session.get(Student, student_id)
        if not st: return
        pdf_buf = await generate_student_report_card_pdf(st.id, lang=user.language if user else "tr")
        file = BufferedInputFile(pdf_buf.read(), filename=f"Karne_{st.student_number}_{st.class_name}.pdf")
        await query.message.answer_document(file, caption=get_text("pdf_report_ready", user.language if user else "tr", name=escape_md(st.full_name)), parse_mode="Markdown")
        pdf_buf.close()
    await query.answer()

@router.callback_query(F.data == "upload_medical_init")
async def cb_upload_med_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("upload_med_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="Markdown")
    await state.set_state(Form.waiting_medical_photo)
    await query.answer()

@router.message(Form.waiting_medical_photo, F.photo)
async def handle_medical_photo(message: Message, state: FSMContext):
    photo_file_id = message.photo[-1].file_id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        st = await session.get(Student, user.current_child_id) if user and user.current_child_id else None
        if not st:
            await message.answer("Öğrenci bulunamadı.")
            return

        report = MedicalReport(student_id=st.id, parent_telegram_id=user.telegram_id, file_id=photo_file_id, caption=message.caption or "Sağlık Raporu")
        session.add(report)
        await session.commit()
        await message.answer(get_text("med_uploaded_success", user.language if user else "tr"))
        await render_clean_dashboard(message, user)

        for a_id in ADMIN_IDS:
            caption_adm = f"🏥 <b>YENİ SAĞLIK RAPORU</b> (#{report.id})\n• Öğrenci: {escape_html(st.full_name)} ({escape_html(st.class_name)})"
            adm_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text="✅ Onayla", callback_data=f"adm:appr_med:{report.id}"), InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_med:{report.id}")]
            ])
            await message.bot.send_photo(chat_id=a_id, photo=photo_file_id, caption=caption_adm, reply_markup=adm_kb, parse_mode="HTML")
    await state.clear()

@router.callback_query(F.data == "adm:medical_list")
async def cb_medical_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        reports = (await session.execute(select(MedicalReport, Student).join(Student, MedicalReport.student_id == Student.id).where(MedicalReport.status == "pending"))).all()
        if not reports:
            await safe_edit_or_answer(query, get_text("no_pending_medical", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return
        buttons = []
        for rep, st in reports:
            buttons.append([InlineKeyboardButton(text=f"🏥 {st.full_name} ({st.class_name})", callback_data=f"adm:view_med:{rep.id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("pending_medical_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("adm:view_med:"))
async def cb_view_medical(query: CallbackQuery):
    rep_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        rep = await session.get(MedicalReport, rep_id)
        st = await session.get(Student, rep.student_id) if rep else None
        if not rep or not st: return
        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="✅ Onayla", callback_data=f"adm:appr_med:{rep.id}"), InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_med:{rep.id}")],
            [InlineKeyboardButton(text="⬅️ Geri", callback_data="adm:medical_list")]
        ])
        await query.message.bot.send_photo(chat_id=query.from_user.id, photo=rep.file_id, caption=f"🏥 <b>Sağlık Raporu:</b> {escape_html(st.full_name)}", reply_markup=adm_kb, parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data.startswith("adm:appr_med:"))
async def cb_approve_medical(query: CallbackQuery):
    rep_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        rep = await session.get(MedicalReport, rep_id)
        if rep:
            rep.status = "approved"
            att = (await session.execute(select(Attendance).where(Attendance.student_id == rep.student_id, Attendance.date == get_local_date()))).scalar_one_or_none()
            if att: att.status = "excused"
            await session.commit()
            p_u = await session.get(User, rep.parent_telegram_id)
            await safe_send_message(query.message.bot, rep.parent_telegram_id, get_text("medical_approved", p_u.language if p_u else "tr"))
            await query.answer("Rapor onaylandı.", show_alert=True)
            await cb_medical_list(query)
            return
    await query.answer()

@router.callback_query(F.data.startswith("adm:rej_med:"))
async def cb_reject_medical(query: CallbackQuery):
    rep_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        rep = await session.get(MedicalReport, rep_id)
        if rep:
            rep.status = "rejected"
            await session.commit()
            await query.answer("Rapor reddedildi.", show_alert=True)
            await cb_medical_list(query)
            return
    await query.answer()

@router.callback_query(F.data == "parent:switch_student")
async def cb_parent_switch_student(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        kids = (await session.execute(select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == user.telegram_id))).scalars().all()
        if not kids:
            await query.answer("Kayıtlı öğrenci yok.", show_alert=True)
            return
        buttons = []
        for k in kids:
            act = "⭐ " if k.id == user.current_child_id else ""
            buttons.append([InlineKeyboardButton(text=f"{act}{k.full_name} ({k.class_name})", callback_data=f"set_child:{k.id}")])
        buttons.append([InlineKeyboardButton(text="🏠 Ana Menü", callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("parent_choose_child", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("set_child:"))
async def cb_set_child(query: CallbackQuery):
    child_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        user.current_child_id = child_id
        await session.commit()
    await render_clean_dashboard(query, user)
    await query.answer("Öğrenci seçildi.")

# ======================================================================
# 15. GİZLİ NUMARATÖR (REPLY KEYBOARD) & SESSİZ ANİMASYONLU PIN MOTORU
# ======================================================================

PIN_PENDING_ACTIONS = {}
ADMIN_PIN_INPUT = {}
ADMIN_PIN_FAILURES = {}
PIN_CHANGE_SESSION = {}
PIN_MSG_ID = {}
PIN_CHAT_ID = {}

def render_pin_screen(cur_pin: str, error_msg: str = "", is_success: bool = False, is_locked: bool = False, lang: str = "tr", custom_prompt: str = "") -> str:
    if is_success: dots = "🟢  🟢  🟢  🟢"
    elif is_locked: dots = "🔴  🔴  🔴  🔴"
    else: dots = "  ".join(["🔵" if i < len(cur_pin) else "⚪" for i in range(4)])

    prompt = custom_prompt or get_text("prompt_pin_current", lang)
    box = f"<code>[  {dots}  ]</code>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
    err_part = f"\n\n⚠️ <i>{error_msg}</i>" if error_msg else ""
    return f"🔐 <b>İDARİ GÜVENLİK PİN KALKANI</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n{prompt}\n\n{box}{err_part}"

async def prompt_for_admin_pin(query: CallbackQuery, state: FSMContext | None, action_callback_data: str):
    user_id = query.from_user.id
    if state: await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    PIN_PENDING_ACTIONS[user_id] = action_callback_data
    ADMIN_PIN_INPUT[user_id] = ""
    ADMIN_PIN_FAILURES[user_id] = 0

    chat_id = query.message.chat.id if (query and query.message) else user_id
    try: await query.message.delete()
    except Exception: pass

    # Ekranda HİÇBİR inline buton yok! Sadece alt ReplyKeyboardMarkup numaratörü var.
    pin_numpad_kb = get_pin_reply_kb(lang)
    m_sent = await query.message.bot.send_message(
        chat_id=chat_id,
        text=render_pin_screen("", lang=lang),
        reply_markup=pin_numpad_kb,
        parse_mode="HTML"
    )
    PIN_MSG_ID[user_id] = m_sent.message_id
    PIN_CHAT_ID[user_id] = chat_id
    await query.answer()

@router.callback_query(F.data == "adm:change_pin_init")
async def cb_admin_change_pin_init(query: CallbackQuery, state: FSMContext | None = None):
    user_id = query.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, user_id): return
        pin_setting = await session.get(SystemSetting, "admin_pin")
        is_custom_set = (pin_setting is not None and bool(pin_setting.value.strip()))

    chat_id = query.message.chat.id if query and query.message else user_id
    try: await query.message.delete()
    except Exception: pass

    if not is_custom_set:
        PIN_CHANGE_SESSION[user_id] = {"step": "enter_new", "input": "", "new_pin": ""}
        prompt_txt = get_text("prompt_pin_new", lang)
    else:
        PIN_CHANGE_SESSION[user_id] = {"step": "verify_current", "input": "", "new_pin": ""}
        prompt_txt = get_text("prompt_pin_current", lang)

    pin_numpad_kb = get_pin_reply_kb(lang)
    m_sent = await query.message.bot.send_message(
        chat_id=chat_id,
        text=render_pin_screen("", lang=lang, custom_prompt=prompt_txt),
        reply_markup=pin_numpad_kb,
        parse_mode="HTML"
    )
    PIN_MSG_ID[user_id] = m_sent.message_id
    PIN_CHAT_ID[user_id] = chat_id
    await query.answer()

PIN_NUMPAD_KEYS = {
    "0", "1", "2", "3", "4", "5", "6", "7", "8", "9",
    "⌫", "⌫ Sil", "⌫ Стереть", "⌫ O'chirish", "⌫ Del",
    "❌", "❌ Vazgeç", "❌ Отмена", "❌ Bekor", "❌ Cancel", "❌ İşlemi İptal Et"
}

@router.message(any_state, F.text.in_(PIN_NUMPAD_KEYS))
async def handle_pin_reply_key_press(message: Message, state: FSMContext):
    user_id = message.from_user.id
    chat_id = message.chat.id
    txt = message.text.strip()

    in_action_pin = (user_id in PIN_PENDING_ACTIONS)
    in_change_pin = (user_id in PIN_CHANGE_SESSION)
    if not in_action_pin and not in_change_pin: return

    # SOHBETE MESAJ DÜŞMEMESİ İÇİN GÖNDERİLEN NUMARAYI ANINDA SİLİYORUZ (SESSİZ ARKA PLAN)
    try: await message.delete()
    except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    key = "cancel" if ("❌" in txt or "İptal" in txt or "Bekor" in txt) else ("del" if "⌫" in txt else txt)
    target_msg_id = PIN_MSG_ID.get(user_id)
    if not target_msg_id: return

    # 1. PIN Değiştirme Akışı
    if in_change_pin:
        sess = PIN_CHANGE_SESSION[user_id]
        step = sess["step"]
        cur = sess["input"]

        if key == "cancel":
            PIN_CHANGE_SESSION.pop(user_id, None)
            PIN_MSG_ID.pop(user_id, None)
            try: await message.bot.delete_message(chat_id=chat_id, message_id=target_msg_id)
            except Exception: pass
            role_kb = get_role_reply_kb("admin", lang)
            await message.bot.send_message(chat_id=chat_id, text=get_text("action_cancelled", lang), reply_markup=role_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        if key == "del": cur = cur[:-1]
        elif key.isdigit() and len(cur) < 4: cur += key
        sess["input"] = cur

        if len(cur) < 4:
            p_prompt = get_text("prompt_pin_current" if step == "verify_current" else ("prompt_pin_new" if step == "enter_new" else "prompt_pin_confirm"), lang)
            try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen(cur, lang=lang, custom_prompt=p_prompt), parse_mode="HTML")
            except Exception: pass
            return
        else:
            real_admin_pin = await get_current_admin_pin()
            if step == "verify_current":
                if cur == real_admin_pin:
                    sess["step"] = "enter_new"
                    sess["input"] = ""
                    try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen("", lang=lang, custom_prompt=get_text("prompt_pin_new", lang)), parse_mode="HTML")
                    except Exception: pass
                    return
                else:
                    PIN_CHANGE_SESSION.pop(user_id, None)
                    PIN_MSG_ID.pop(user_id, None)
                    try: await message.bot.delete_message(chat_id=chat_id, message_id=target_msg_id)
                    except Exception: pass
                    role_kb = get_role_reply_kb("admin", lang)
                    await message.bot.send_message(chat_id=chat_id, text="❌ " + get_text("pin_current_wrong", lang), reply_markup=role_kb, parse_mode="Markdown")
                    await render_clean_dashboard(message, user)
                    return
            elif step == "enter_new":
                sess["new_pin"] = cur
                sess["step"] = "confirm_new"
                sess["input"] = ""
                try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen("", lang=lang, custom_prompt=get_text("prompt_pin_confirm", lang)), parse_mode="HTML")
                except Exception: pass
                return
            elif step == "confirm_new":
                if cur == sess["new_pin"]:
                    new_val = cur
                    PIN_CHANGE_SESSION.pop(user_id, None)
                    PIN_MSG_ID.pop(user_id, None)
                    async with AsyncSessionLocal() as session_w:
                        st = await session_w.get(SystemSetting, "admin_pin")
                        if not st: session_w.add(SystemSetting(key="admin_pin", value=new_val))
                        else: st.value = new_val
                        await session_w.commit()
                    SETTINGS_CACHE["admin_pin"] = new_val
                    try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen(cur, is_success=True, lang=lang) + "\n\n" + get_text("pin_changed_success", lang), parse_mode="HTML")
                    except Exception: pass
                    await asyncio.sleep(0.5)
                    try: await message.bot.delete_message(chat_id=chat_id, message_id=target_msg_id)
                    except Exception: pass
                    role_kb = get_role_reply_kb("admin", lang)
                    await message.bot.send_message(chat_id=chat_id, text="✅ " + get_text("pin_changed_success", lang), reply_markup=role_kb, parse_mode="HTML")
                    await render_clean_dashboard(message, user)
                    return
                else:
                    sess["step"] = "enter_new"
                    sess["input"] = ""
                    try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen("", error_msg=get_text("pin_mismatch_error", lang), lang=lang, custom_prompt=get_text("prompt_pin_new", lang)), parse_mode="HTML")
                    except Exception: pass
                    return

    # 2. İdari Aksiyon PIN Akışı
    if in_action_pin:
        action = PIN_PENDING_ACTIONS[user_id]
        cur = ADMIN_PIN_INPUT.get(user_id, "")

        if key == "cancel":
            PIN_PENDING_ACTIONS.pop(user_id, None)
            ADMIN_PIN_INPUT.pop(user_id, None)
            PIN_MSG_ID.pop(user_id, None)
            try: await message.bot.delete_message(chat_id=chat_id, message_id=target_msg_id)
            except Exception: pass
            role_kb = get_role_reply_kb(user.role, lang)
            await message.bot.send_message(chat_id=chat_id, text=get_text("action_cancelled", lang), reply_markup=role_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        if key == "del": cur = cur[:-1]
        elif key.isdigit() and len(cur) < 4: cur += key
        ADMIN_PIN_INPUT[user_id] = cur

        if len(cur) < 4:
            try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen(cur, lang=lang), parse_mode="HTML")
            except Exception: pass
            return
        else:
            real_admin_pin = await get_current_admin_pin()
            if cur == real_admin_pin:
                PIN_PENDING_ACTIONS.pop(user_id, None)
                ADMIN_PIN_INPUT.pop(user_id, None)
                PIN_MSG_ID.pop(user_id, None)
                try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen(cur, is_success=True, lang=lang) + "\n\n✅ <b>İdari PIN Doğrulandı!</b>", parse_mode="HTML")
                except Exception: pass
                await asyncio.sleep(0.3)
                try: await message.bot.delete_message(chat_id=chat_id, message_id=target_msg_id)
                except Exception: pass

                role_kb = get_role_reply_kb("admin", lang)
                m_r = await message.bot.send_message(chat_id=chat_id, text="⚡ " + get_text("admin_title", lang), reply_markup=role_kb, parse_mode="Markdown")

                from aiogram import types
                dummy_q = CallbackQuery(id="0", from_user=types.User(id=user_id, is_bot=False, first_name=user.full_name or "Admin"), chat_instance="0", message=Message(message_id=m_r.message_id, date=datetime.utcnow(), chat=types.Chat(id=chat_id, type="private")), data=action)
                if action == "adm:export_all_excel": await cb_admin_export_all_direct(dummy_q)
                elif action == "adm:class_promotion_confirm": await cb_admin_class_promotion_confirm(dummy_q)
                elif action == "adm:emergency_init": await cb_admin_emergency_init(dummy_q, None)
                return
            else:
                fails = ADMIN_PIN_FAILURES.get(user_id, 0) + 1
                ADMIN_PIN_FAILURES[user_id] = fails
                ADMIN_PIN_INPUT[user_id] = ""
                if fails >= 3:
                    PIN_PENDING_ACTIONS.pop(user_id, None)
                    ADMIN_PIN_INPUT.pop(user_id, None)
                    PIN_MSG_ID.pop(user_id, None)
                    try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen("", is_locked=True, lang=lang) + f"\n\n❌ <b>{get_text('invalid_admin_pin', lang)}</b>", parse_mode="HTML")
                    except Exception: pass
                    await asyncio.sleep(1.0)
                    try: await message.bot.delete_message(chat_id=chat_id, message_id=target_msg_id)
                    except Exception: pass
                    role_kb = get_role_reply_kb("admin", lang)
                    await message.bot.send_message(chat_id=chat_id, text="⛔ " + get_text("auth_locked", lang), reply_markup=role_kb, parse_mode="Markdown")
                    await render_clean_dashboard(message, user)
                    return
                else:
                    rem = 3 - fails
                    err_txt = f"Hatalı PIN! Kalan: {rem}" if lang == "tr" else f"Invalid PIN! Remaining: {rem}"
                    try: await message.bot.edit_message_text(chat_id=chat_id, message_id=target_msg_id, text=render_pin_screen("", error_msg=err_txt, lang=lang), parse_mode="HTML")
                    except Exception: pass
                    return

@router.callback_query(F.data == "adm:export_all_excel")
async def cb_admin_export_all(query: CallbackQuery, state: FSMContext | None = None):
    await prompt_for_admin_pin(query, state, "adm:export_all_excel")

async def cb_admin_export_all_direct(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    buf = await export_all_school_data_excel()
    today_str = datetime.utcnow().strftime("%d_%m_%Y")
    file = BufferedInputFile(buf.read(), filename=f"Okul_Yedegi_{today_str}.xlsx")
    await query.message.answer_document(file, caption=get_text("export_ready", lang, date=today_str), parse_mode="Markdown")
    buf.close()
    await query.answer()

@router.callback_query(F.data == "adm:class_promotion_init")
async def cb_admin_class_promotion_init(query: CallbackQuery, state: FSMContext | None = None):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        buttons = [
            [InlineKeyboardButton(text="✅ Sınıfları Yükselt (PIN Gerekir)", callback_data="adm:class_promotion_pin_prompt")],
            [get_nav_buttons(lang, back_callback="adm:cat_staff")]
        ]
        await safe_edit_or_answer(query, get_text("promotion_confirm_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:class_promotion_pin_prompt")
async def cb_admin_class_promotion_pin_prompt(query: CallbackQuery, state: FSMContext | None = None):
    await prompt_for_admin_pin(query, state, "adm:class_promotion_confirm")

async def cb_admin_class_promotion_confirm(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        students = (await session.execute(select(Student))).scalars().all()
        cnt = 0
        for s in students:
            m = re.match(r"^(\d+)(.*)$", s.class_name.strip())
            if m:
                gr = int(m.group(1))
                sfx = m.group(2)
                s.class_name = f"MEZUN{sfx}" if gr >= 12 else f"{gr+1}{sfx}"
                cnt += 1
        await session.commit()
        await safe_edit_or_answer(query, get_text("promotion_success", lang, count=cnt), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang, back_callback="adm:cat_staff")]), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:emergency_init")
async def cb_admin_emergency_pin_guard(query: CallbackQuery, state: FSMContext):
    await prompt_for_admin_pin(query, state, "adm:emergency_init")

async def cb_admin_emergency_init(query: CallbackQuery, state: FSMContext | None = None):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    cancel_kb = get_cancel_reply_kb(lang)
    await safe_edit_or_answer(query, get_text("emergency_alert_prompt", lang), reply_markup=cancel_kb, parse_mode="Markdown")
    if state: await state.set_state(Form.waiting_emergency_text)

@router.message(Form.waiting_emergency_text)
async def process_emergency_text(message: Message, state: FSMContext):
    try: await message.delete()
    except Exception: pass
    em_text = message.text.strip()
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        alert = EmergencyAlert(message_text=em_text, created_by=message.from_user.id)
        session.add(alert)
        await session.commit()

        parents = (await session.execute(select(User).where(User.role == "parent"))).scalars().all()
        for p in parents:
            em_msg = f"🚨 *ACİL DURUM / KIRMIZI ALARM* 🚨\n\n{escape_md(em_text)}"
            ack_btn = InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="🚨 Okudum, Bilgim Var", callback_data=f"ack_em:{alert.id}")]])
            await safe_send_message(message.bot, p.telegram_id, em_msg, reply_markup=ack_btn, parse_mode="Markdown")
            await asyncio.sleep(0.04)

        await message.answer(f"🚨 Acil durum alarmı {len(parents)} veliye iletildi!", reply_markup=get_role_reply_kb("admin", user.language if user else "tr"), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data.startswith("ack_em:"))
async def cb_ack_emergency(query: CallbackQuery):
    alert_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        existing = (await session.execute(select(EmergencyAck).where(EmergencyAck.alert_id == alert_id, EmergencyAck.user_telegram_id == query.from_user.id))).scalar_one_or_none()
        if not existing:
            session.add(EmergencyAck(alert_id=alert_id, user_telegram_id=query.from_user.id))
            await session.commit()
    await query.answer("✅ Onayınız idareye iletildi.", show_alert=True)
    await query.message.edit_reply_markup(reply_markup=None)

@router.callback_query(F.data == "adm:emergency_monitor")
async def cb_admin_emergency_monitor(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        latest = (await session.execute(select(EmergencyAlert).order_by(desc(EmergencyAlert.created_at)).limit(1))).scalar_one_or_none()
        if not latest:
            await safe_edit_or_answer(query, "Kayıtlı acil durum bulunamadı.", reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
            await query.answer()
            return
        acked = (await session.execute(select(EmergencyAck.user_telegram_id).where(EmergencyAck.alert_id == latest.id))).scalars().all()
        parents = (await session.execute(select(User).where(User.role == "parent"))).scalars().all()
        unacked = [p for p in parents if p.telegram_id not in acked]
        lines = [f"🚨 <b>Onaylamayan Veliler ({len(unacked)}):</b>\n"]
        for p in unacked[:20]: lines.append(f"• 👤 {p.full_name or 'Veli'} (ID: <code>{p.telegram_id}</code>)")
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:cockpit")
async def cb_cockpit(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = get_local_date()
        total_st = (await session.execute(select(func.count(Student.id)))).scalar() or 0
        pres_cnt = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "present"))).scalar() or 0
        abs_cnt = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "absent"))).scalar() or 0

        pct = round((pres_cnt / total_st * 100), 1) if total_st > 0 else 100.0
        bar = render_progress_bar(int(pct // 10), 10)
        text = (
            f"📊 <b>GÜNLÜK SABAH KOKPİTİ ({today.strftime('%d.%m.%Y')})</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━\n"
            f"📈 <b>Katılım:</b> <code>{bar}</code> <b>%{pct}</b>\n\n"
            f"• Toplam: <b>{total_st}</b> | Gelen: <b>{pres_cnt}</b> | Gelmeyen: <b>{abs_cnt}</b>\n━━━━━━━━━━━━━━━━━━━━━━━━━━━━"
        )
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang, back_callback="adm:cat_reports")]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:unack_notifs")
async def cb_admin_unack_notifs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        since = datetime.utcnow() - timedelta(hours=36)
        unacks = (await session.execute(select(CriticalNotification).where(CriticalNotification.created_at >= since, CriticalNotification.acknowledged_at == None).limit(15))).scalars().all()
        if not unacks:
            await safe_edit_or_answer(query, get_text("all_notifs_acknowledged", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang, back_callback="adm:cat_reports")]))
            await query.answer()
            return
        lines = [f"⚠️ <b>Okunmamış Devamsızlıklar ({len(unacks)}):</b>\n"]
        for un in unacks: lines.append(f"• ID: <code>{un.user_telegram_id}</code> - {un.created_at.strftime('%d.%m %H:%M')}")
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang, back_callback="adm:cat_reports")]), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:pdf_menu")
async def cb_pdf_menu(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all() or ["9-A"]
        buttons = [[InlineKeyboardButton(text="👨‍🏫 Tüm Öğretmenler (PDF)", callback_data="adm:gen_pdf_teachers")]]
        for c in classes: buttons.append([InlineKeyboardButton(text=f"📄 {c}", callback_data=f"adm:gen_pdf:{c}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_tools"))
        await safe_edit_or_answer(query, get_text("select_pdf_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data == "adm:gen_pdf_teachers")
async def cb_generate_pdf_teachers(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
    pdf_buf = await generate_teachers_pdf_cards(lang=user.language if user else "tr")
    file = BufferedInputFile(pdf_buf.read(), filename="Ogretmenler_Sifreleri.pdf")
    await query.message.answer_document(file, caption="👨‍🏫 Öğretmen Şifre Kartları Ektedir.")
    pdf_buf.close()
    await query.answer()

@router.callback_query(F.data.startswith("adm:gen_pdf:"))
async def cb_generate_pdf(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
    pdf_buf = await generate_classroom_pdf_cards(class_name, lang=user.language if user else "tr")
    file = BufferedInputFile(pdf_buf.read(), filename=f"{class_name}_Sifre_Kartlari.pdf")
    await query.message.answer_document(file, caption=f"📄 {class_name} Şifre Kartları Ektedir.")
    pdf_buf.close()
    await query.answer()

@router.callback_query(F.data == "adm:excel_hub")
async def cb_admin_excel_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        buttons = [
            [InlineKeyboardButton(text="📥 Toplu Öğrenci Yükle", callback_data="adm:excel_info")],
            [InlineKeyboardButton(text="📦 Tam Sistem Veri Yedeği (.xlsx)", callback_data="adm:export_all_excel")],
            get_nav_buttons(lang, back_callback="adm:cat_tools")
        ]
        await safe_edit_or_answer(query, "📥 <b>Kurumsal Excel Merkezi:</b>", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="HTML")
    await query.answer()

@router.callback_query(F.data == "adm:excel_info")
async def cb_excel_info(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("excel_info", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang, back_callback="adm:excel_hub")]), parse_mode="Markdown")
    await query.answer()

@router.message(F.document)
async def admin_excel_upload_direct(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id): return

    doc = message.document
    if not (doc.file_name or "").endswith(".xlsx"):
        await message.answer(get_text("file_type_not_allowed_error", lang))
        return

    try:
        file_info = await message.bot.get_file(doc.file_id)
        file_bytes = await message.bot.download_file(file_info.file_path)
        count, out_excel = await process_student_excel(file_bytes.read())
        file = BufferedInputFile(out_excel.read(), filename="Giris_Kodlari.xlsx")
        await message.answer_document(file, caption=get_text("excel_done", lang, count=count), parse_mode="Markdown")
        out_excel.close()
    except Exception:
        await message.answer(get_text("excel_format_error", lang))

@router.callback_query(F.data == "adm:menu_edit")
async def cb_admin_menu_edit(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("prompt_menu_update", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
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
        menu = (await session.execute(select(CafeteriaMenu).where(CafeteriaMenu.date == today))).scalar_one_or_none()
        if not menu: session.add(CafeteriaMenu(date=today, menu_text=m_text))
        else: menu.menu_text = m_text
        await session.commit()
        await message.answer(get_text("menu_updated", user.language if user else "tr"))
        await render_clean_dashboard(message, user)

@router.callback_query(F.data == "adm:dashboard")
async def cb_admin_dashboard(query: CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user: await render_clean_dashboard(query, user)
    await query.answer()

REPLY_BUTTON_ACTIONS = {
    "rk_logout": "act_restart",
    "rk_restart": "act_restart",
    "rk_main_menu": "act_main_menu",
    "rk_admin_dash": "act_main_menu",
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
    "rk_cat_settings": "act_cat_settings",
    "rk_attendance": "act_attendance",
    "rk_grade": "act_grade",
    "rk_homework": "act_homework",
    "rk_behavior": "act_behavior",
    "rk_appointments": "act_appointments",
    "rk_report": "act_report",
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
    return None

@router.message(any_state, F.text.func(lambda text: match_reply_button(text) is not None))
async def global_reply_keyboard_router(message: Message, state: FSMContext):
    action = match_reply_button(message.text)
    user_id = message.from_user.id
    try: await message.delete()
    except Exception: pass

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            role = "admin" if user_id in ADMIN_IDS else "guest"
            user = User(telegram_id=user_id, language="tr", role=role, full_name=message.from_user.full_name)
            session.add(user)
            await session.commit()

    if action in ("act_logout", "act_restart"):
        await handle_bot_restart_cmd(message, state)
        return
    elif action == "act_cat_staff": await cb_cat_staff(message)
    elif action == "act_cat_reports": await cb_cat_reports(message)
    elif action == "act_cat_requests": await cb_cat_requests(message)
    elif action == "act_cat_tools": await cb_cat_tools(message)
    elif action == "act_cat_settings": await cb_cat_settings(message)
    elif action == "act_main_menu": await render_clean_dashboard(message, user)
    elif action == "act_lang": await safe_edit_or_answer(message, get_text("lang_select", user.language), reply_markup=get_language_inline_kb())
    elif action == "act_enter_code":
        cancel_kb = get_cancel_reply_kb(user.language)
        await message.answer(get_text("prompt_enter_code_direct", user.language), reply_markup=cancel_kb, parse_mode="Markdown")
        await state.set_state(Form.waiting_auth_code)
    elif action == "act_req_access":
        dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_req_access")
        await cb_act_req_access(dummy_q, state)
    elif action == "act_attendance":
        dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:classes")
        await cb_teacher_classes(dummy_q, state)
    elif action == "act_grade":
        dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:grade_classes")
        await cb_grade_classes(dummy_q, state)
    elif action == "act_report":
        dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="act_view_report")
        await cb_view_report_card(dummy_q, state)
    elif action == "act_upload_medical":
        dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="upload_medical_init")
        await cb_upload_med_init(dummy_q, state)
    elif action == "act_switch_student":
        dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="parent:switch_student")
        await cb_parent_switch_student(dummy_q, state)

@router.message(any_state, F.text.func(lambda text: normalize_code(text).startswith(("VELI-", "OGR-", "HCA-", "ADM-")) or normalize_code(text) == ADMIN_CODE))
@router.message(F.text)
async def smart_text_auth_router(message: Message, state: FSMContext):
    clean_code = normalize_code(message.text)
    is_code = clean_code.startswith(("VELI-", "OGR-", "HCA-", "ADM-")) or clean_code == ADMIN_CODE
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if is_code or (user and user.role == "guest"):
            await process_auth_code_string(message.text, message.from_user.id, message, state)

# ======================================================================
# 16. FASTAPI VE ASENKRON DÖNGÜLER (LIFESPAN & WEBHOOK)
# ======================================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
dp.include_router(router)

@dp.error()
async def global_error_shield(event, exception):
    print(f"--> [HATA KALKANI] {exception}")
    return True

async def background_attendance_loop():
    while True:
        try: await run_attendance_delay_worker(bot)
        except Exception: pass
        await asyncio.sleep(60)

async def background_evening_briefing_loop():
    while True:
        try:
            now_l = get_local_now()
            if now_l.weekday() < 5 and now_l.hour == 18 and now_l.minute == 30:
                await run_evening_briefing_worker(bot)
                await asyncio.sleep(70)
        except Exception: pass
        await asyncio.sleep(30)

async def background_keep_alive_pinger():
    await asyncio.sleep(60)
    while True:
        url = WEBHOOK_URL.replace("/webhook", "/health") if WEBHOOK_URL else None
        if url:
            try:
                import urllib.request
                req = urllib.request.Request(url, headers={"User-Agent": "OkulBot/2.0"})
                await asyncio.to_thread(urllib.request.urlopen, req, timeout=10)
            except Exception: pass
        await asyncio.sleep(600)

@asynccontextmanager
async def lifespan(app: FastAPI):
    print("--> [1/4] Veritabanı başlatılıyor...")
    await init_db()
    print("--> [2/4] Veritabanı hazır.")
    if WEBHOOK_URL:
        try:
            await bot.delete_webhook(drop_pending_updates=False)
            await bot.set_webhook(url=WEBHOOK_URL, secret_token=WEBHOOK_SECRET, drop_pending_updates=False)
            print(f"--> [3/4] Webhook kuruldu: {WEBHOOK_URL}")
        except Exception as e: print(f"Webhook error: {e}")
    t1 = asyncio.create_task(background_attendance_loop())
    t2 = asyncio.create_task(background_evening_briefing_loop())
    t3 = asyncio.create_task(background_keep_alive_pinger())
    print("--> [4/4] Sistem hazır.")
    yield
    t1.cancel(); t2.cancel(); t3.cancel()
    try: await bot.session.close()
    except Exception: pass

app = FastAPI(title="OkulYonetimBot", lifespan=lifespan)

@app.get("/")
@app.get("/health")
async def health_check():
    return {"status": "ok", "service": "OkulYonetimBot", "uptime": True}

@app.post("/webhook")
async def telegram_webhook(request: Request):
    secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if WEBHOOK_SECRET and secret != WEBHOOK_SECRET:
        return JSONResponse(status_code=status.HTTP_403_FORBIDDEN, content={"error": "Invalid secret"})
    try: data = await request.json()
    except Exception: return JSONResponse(status_code=status.HTTP_400_BAD_REQUEST, content={"error": "Bad JSON"})

    try:
        telegram_update = Update.model_validate(data, context={"bot": bot})
        u_id = telegram_update.message.from_user.id if (telegram_update.message and telegram_update.message.from_user) else (telegram_update.callback_query.from_user.id if (telegram_update.callback_query and telegram_update.callback_query.from_user) else None)
        if u_id:
            import time
            now_t = time.time()
            if u_id in USER_COOLDOWN:
                if now_t < USER_COOLDOWN[u_id]: return {"ok": True}
                else: del USER_COOLDOWN[u_id]; USER_REQUEST_LOG[u_id] = []
            reqs = [t for t in USER_REQUEST_LOG.get(u_id, []) if now_t - t < 60.0]
            sec_reqs = [t for t in reqs if now_t - t < 1.0]
            # PIN tuşlamasında ardışık tıklamaların düşmemesi için 10 req/s
            if len(sec_reqs) >= 10 or len(reqs) >= 60:
                USER_COOLDOWN[u_id] = now_t + 15.0
                USER_REQUEST_LOG[u_id] = reqs
                return {"ok": True}
            reqs.append(now_t)
            USER_REQUEST_LOG[u_id] = reqs
        await asyncio.wait_for(dp.feed_update(bot, telegram_update), timeout=15.0)
    except Exception: pass
    return {"ok": True}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    if not WEBHOOK_URL: asyncio.run(dp.start_polling(bot))
    else: uvicorn.run("app:app", host="0.0.0.0", port=port)
