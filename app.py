# -*- coding: utf-8 -*-
"""
OKUL YÖNETİM TELEGRAM BOTU - KURUMSAL VE EKSİKSİZ TAM SİSTEM MİMARİSİ (PROD V22 - FINAL BULLETPROOF)
Altyapı: FastAPI + aiogram 3.x Webhook + PostgreSQL (Neon / Supabase) / SQLite Çift Motor Kalkanı
4 Dilli Kusursuz Arayüz (i18n): 🇹🇷 Türkçe, 🇷🇺 Русский, 🇺🇿 O'zbekcha (Lotin), 🇬🇧 English
Platform: Render Web Service / VPS / Docker

Öne Çıkan Mimari Çözümler:
- 🛡️ Çift Cihaz ve Hesap Çalma Koruması (HCA ve OGR kodları ilk Telegram hesabına kilitlenir, VELI kodları çoklu ebeveyn kullanımına açıktır).
- 🔑 Akıllı Doğrudan Kod Tanıma (Butona basılmasa dahi HCA, OGR, VELI, ADM kodları anında tanınır).
- 🔤 Özel Karakter Koruması (escape_md ile Markdown çökmesi sıfırlanmıştır).
- 📄 Veli ve Öğrenci İçin Resmi PDF Karne / Not Dökümü (Dersler, sınav türleri, rozetler ve genel ortalama).
- 📝 Sınav Türü Seçimli Not Girişi (1. Yazılı, 2. Yazılı, Sözlü / Performans) ve İsteğe Bağlı Öğretmen Görüşü.
- 👨‍💼 Yönetici Kadrosu & Kullanıcı Rehberi Masası (Profil kartları, kalıcı/geçici yönetici atama, ban/unban, 1:1 mesaj ve arama).
- 🚨 Devamsızlık Bildiriminde Tek Tıkla Mazeret Bildirme ve 15 Dk Düzeltme Kalkanı.
- 📦 Cuma 18:00 Otomatik Çok Sayfalı Okul Veri Yedeği (.xlsx) ve Render 7/24 Uyku Önleyici (Self-Pinger).
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
    if not text:
        return ""
    s = str(text)
    s = s.replace(chr(92), chr(92) + chr(92))
    for ch in ("_", "*", "`", "[", "]"):
        s = s.replace(ch, chr(92) + ch)
    return s

def normalize_code(code_str: str) -> str:
    if not code_str:
        return ""
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
            await msg.answer(text, reply_markup=reply_markup, parse_mode=parse_mode)
        except Exception:
            await msg.answer(text, reply_markup=reply_markup, parse_mode=None)
    elif isinstance(target, Message):
        try:
            await target.answer(text, reply_markup=reply_markup, parse_mode=parse_mode)
        except Exception:
            await target.answer(text, reply_markup=reply_markup, parse_mode=None)

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
    role: Mapped[str] = mapped_column(String(20), default="guest") # admin, teacher, parent, student, guest
    language: Mapped[str] = mapped_column(String(5), default="tr")
    full_name: Mapped[str] = mapped_column(String(100), nullable=True)
    username: Mapped[str] = mapped_column(String(100), nullable=True)
    phone: Mapped[str] = mapped_column(String(30), nullable=True)
    admin_type: Mapped[str] = mapped_column(String(20), default="none") # permanent, temporary, none
    admin_until: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    previous_role: Mapped[str] = mapped_column(String(20), default="guest")
    failed_attempts: Mapped[int] = mapped_column(Integer, default=0)
    locked_until: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    is_blacklisted: Mapped[bool] = mapped_column(Boolean, default=False)
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

class Attendance(Base):
    __tablename__ = "attendances"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    date: Mapped[date] = mapped_column(Date, default=lambda: datetime.utcnow().date(), index=True)
    status: Mapped[str] = mapped_column(String(10), default="present") # present, absent, excused
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
    badge: Mapped[str] = mapped_column(String(10)) # 🟢, 🟡, 🔴
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
    status: Mapped[str] = mapped_column(String(20), default="pending") # pending, approved, rejected
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

class CafeteriaMenu(Base):
    __tablename__ = "cafeteria_menus"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    date: Mapped[date] = mapped_column(Date, default=lambda: datetime.utcnow().date(), index=True)
    menu_text: Mapped[str] = mapped_column(Text)

async def init_db():
    async with engine.begin() as conn:
        if "sqlite" in DATABASE_URL:
            try:
                await conn.exec_driver_sql("PRAGMA journal_mode=WAL;")
                await conn.exec_driver_sql("PRAGMA synchronous=NORMAL;")
                await conn.exec_driver_sql("PRAGMA busy_timeout=30000;")
                await conn.exec_driver_sql("PRAGMA cache_size=-64000;")
            except Exception:
                pass
        await conn.run_sync(Base.metadata.create_all)

        # Otomatik Sütun Göçü (conn oturumu içinde güvenle çalışır)
        try:
            if "sqlite" in DATABASE_URL:
                res = await conn.exec_driver_sql("PRAGMA table_info(users);")
                cols = [r[1] for r in res.fetchall()]
                for c_name, c_type in [
                    ("username", "VARCHAR(100)"),
                    ("phone", "VARCHAR(30)"),
                    ("admin_type", "VARCHAR(20) DEFAULT 'none'"),
                    ("admin_until", "DATETIME"),
                    ("previous_role", "VARCHAR(20) DEFAULT 'guest'")
                ]:
                    if c_name not in cols:
                        await conn.exec_driver_sql(f"ALTER TABLE users ADD COLUMN {c_name} {c_type};")
            else:
                for c_name, c_type in [
                    ("username", "VARCHAR(100)"),
                    ("phone", "VARCHAR(30)"),
                    ("admin_type", "VARCHAR(20) DEFAULT 'none'"),
                    ("admin_until", "TIMESTAMP"),
                    ("previous_role", "VARCHAR(20) DEFAULT 'guest'")
                ]:
                    await conn.exec_driver_sql(f"ALTER TABLE users ADD COLUMN IF NOT EXISTS {c_name} {c_type};")
        except Exception:
            pass

        try:
            if "sqlite" in DATABASE_URL:
                res = await conn.exec_driver_sql("PRAGMA table_info(grades);")
                cols = [r[1] for r in res.fetchall()]
                if "exam_type" not in cols:
                    await conn.exec_driver_sql("ALTER TABLE grades ADD COLUMN exam_type VARCHAR(40) DEFAULT '1. Yazılı';")
            else:
                await conn.exec_driver_sql("ALTER TABLE grades ADD COLUMN IF NOT EXISTS exam_type VARCHAR(40) DEFAULT '1. Yazılı';")
        except Exception:
            pass


# ======================================================================
# 3. KUSURSUZ 4 DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {   'en': {   'acknowledged_toast': 'Confirmation recorded.',
              'action_cancelled': '❌ *Action cancelled.*',
              'admin_add_name_prompt': '👤 Enter Full Name of new administrator:',
              'admin_add_tg_id_prompt': '➕ Enter Telegram ID of user to make admin:',
              'admin_added_success': '✅ *{name}* (`{id}`) added as permanent administrator successfully.',
              'admin_admins_hub_title': '👨\u200d💼 *School Administration & Authorities:*\n'
                                        '\n'
                                        'Tap an administrator to view details or manage permissions:',
              'admin_code_generated': '🔑 *ONE-TIME ADMINISTRATOR CODE GENERATED*\n'
                                      '\n'
                                      'Code: `{code}`\n'
                                      '\n'
                                      'Send this code to the user. As soon as they enter it in the bot, their account '
                                      'will be promoted to *Administrator*.',
              'admin_demoted_notification': 'ℹ️ Your administrator privileges have been revoked.',
              'admin_demoted_toast': 'Administrator privileges revoked.',
              'admin_invalid_tg_id': '❌ Invalid Telegram ID! Must be numeric digits only.',
              'admin_promoted_notification': '🎉 *Dear {name},*\nYou have been granted *Permanent Administrator* role!',
              'admin_restart_confirmed': '🔄 *Administrator Panel Restarted.*',
              'admin_sched_edit_title': '📅 *Edit Timetable*\nSelect class:',
              'admin_sched_updated': '✅ Timetable for class *{class_name}* updated.',
              'admin_stats': '📊 *General Status:*\n'
                             '• Classes: *{c_cnt}* | Students: *{s_cnt}* | Teachers: *{t_cnt}*\n'
                             '• Requests: *{req_cnt}* | Medicals: *{med_cnt}*\n'
                             '• Date: *{date}*',
              'admin_title': '⚡ *School Administration Cockpit (Admin)*',
              'admin_unban_notification': '🟢 Your account has been unbanned by administration.',
              'admin_user_card_title': '👤 *ADMINISTRATOR USER CARD*',
              'all_notifs_acknowledged': '✅ *All Absence Notifications Acknowledged!*\n'
                                         '\n'
                                         'There are no unread notifications in the last 36 hours.',
              'appointment_approved_msg': '✅ Teacher accepted your appointment request.',
              'appointment_confirmed_toast': 'Appointment confirmed.',
              'appointment_not_found': '⚠️ Appointment not found.',
              'appointment_rejected_msg': '❌ Teacher is unavailable at requested time.',
              'appointment_sent': '✅ Appointment request sent to teacher.',
              'att_check_all_done': '✅ Attendance recorded for all classes.',
              'att_check_title': '📊 *Daily Attendance Audit ({date})*',
              'att_saved': '✅ Recorded (15-min edit window started).',
              'attendance_correction_notification': 'ℹ️ *CORRECTION:* Absence alert for *{name}* has been corrected '
                                                    '(PRESENT).',
              'attendance_hours_lock': '⚠️ Attendance can only be taken between 07:00 and 19:00.',
              'attendance_intro': '📋 *Attendance: {class_name}*\nTap absent students and save:',
              'attendance_select_class': '📋 Select class to take attendance:',
              'attendance_weekend_lock': '⚠️ Attendance cannot be recorded on weekends.',
              'auth_blacklisted': '🚫 Account has been permanently suspended.',
              'auth_code_already_linked': '⚠️ *This code has already been linked to another Telegram account.*\n'
                                          '\n'
                                          'Please contact the school administration immediately.',
              'auth_failed': '❌ Invalid access code! Remaining attempts: {remaining}',
              'auth_locked': '⛔ Account locked for 1 hour due to security restrictions.',
              'auth_success': '✅ *Authentication Successful!*\nWelcome: *{name}*\nYour role: *{role}*',
              'badge_missing': '🟡 Missing / Needs Work',
              'badge_praise': '🟢 Praise / Achievement',
              'badge_warning': '🔴 Discipline / Warning',
              'blacklisted_title': '🚫 *Blocked & Locked Users:*',
              'broadcast_success': '📢 Dispatched to *{count}* users.',
              'btn_academic_report': '📈 Academic Ranking',
              'btn_acknowledged': '✅ Read / Acknowledged',
              'btn_add_admin_id': '➕ Add Admin by Telegram ID',
              'btn_add_another': '➕ Add Another',
              'btn_add_new_teacher': '➕ Add New Teacher',
              'btn_add_student': '➕ Add Student',
              'btn_add_teacher': '➕ Add Teacher',
              'btn_appr_appointment': '✅ Accept',
              'btn_appr_medical': '✅ Approve Medical Note',
              'btn_appr_request': '✅ Approve',
              'btn_attendance': '📋 Fast Attendance',
              'btn_audit_logs': '📜 System Audit Log',
              'btn_back': '⬅️ Back',
              'btn_ban_user': '🚫 Ban User',
              'btn_blacklist': '🚫 Blocked Users',
              'btn_briefing_off': '🔕 Evening Briefing (OFF)',
              'btn_briefing_on': '🔔 Evening Briefing (ON)',
              'btn_broadcast': '📢 Broadcast Notice',
              'btn_cafeteria_edit': '🍲 Update Cafeteria Menu',
              'btn_cancel_action': '⬅️ Cancel',
              'btn_class_att_sheet': 'Attendance Sheet',
              'btn_class_grade_sheet': 'Grade Sheet',
              'btn_class_pdf_cards': 'Password Cards (PDF)',
              'btn_class_sched': 'Timetable',
              'btn_classes': '🏫 Classes & Students',
              'btn_cockpit': '📊 Morning Cockpit',
              'btn_cockpit_unified': '📊 Morning Cockpit & Attendance',
              'btn_del_grade': '❌ Delete Score',
              'btn_del_student': '❌ Delete Student',
              'btn_del_teacher': '❌ Delete Teacher',
              'btn_delete_action': '🗑️ Delete',
              'btn_download_pdf_report': '📄 Download Official Report Card (PDF)',
              'btn_dur_24h': '⏱️ 24 Hours',
              'btn_dur_30d': '⏱️ 30 Days',
              'btn_dur_7d': '⏱️ 7 Days',
              'btn_edit_class': '🏫 Edit Class',
              'btn_edit_grade': '✏️ Edit Score',
              'btn_edit_name': '👤 Edit Name',
              'btn_edit_no': '🔢 Edit Roll No',
              'btn_edit_student': '✏️ Edit Info',
              'btn_enter_grade': '📝 Grade Book',
              'btn_excel': '📥 Import via Excel',
              'btn_excel_hub': '📥 Excel Center',
              'btn_export_all_data': '📊 Export Full School Data (.xlsx)',
              'btn_gen_admin_code': '🔑 Generate One-Time Admin Code',
              'btn_homework_board': '📢 Homework Board',
              'btn_lang': '🌐 Change Language',
              'btn_login_prompt': '🔑 Log In (Access Code)',
              'btn_main_menu': '🏠 Main Menu',
              'btn_maintenance_toggle': '🚨 Maintenance ({status})',
              'btn_make_perm_admin': '👑 Make Permanent Admin',
              'btn_make_temp_admin': '⏱️ Make Temporary Admin',
              'btn_manage_schedule': '📅 Timetable Management',
              'btn_medical': '🏥 Medicals ({count})',
              'btn_my_hws': '📚 My Homeworks',
              'btn_next': 'Next ➡️',
              'btn_not_available': '❌ Not Available',
              'btn_notices': '📢 School Announcements',
              'btn_pdf': '📄 Password Cards (PDF)',
              'btn_prev': '⬅️ Previous',
              'btn_recent_grades_menu': '🕒 Recent Grades & Edit',
              'btn_refresh_data': '🔄 Refresh Data',
              'btn_reject': '❌ Reject',
              'btn_remind_att': '⚠️ Remind Teachers to Take Attendance',
              'btn_report_card': '📊 Report Card',
              'btn_req_access': '📩 Request Access',
              'btn_req_chat': '📞 Request 1:1 Contact',
              'btn_requests': '🛎️ Requests ({count})',
              'btn_reset_codes': '🔄 Reset Codes',
              'btn_revoke_admin_perm': '❌ Revoke Admin Role',
              'btn_risk_radar': '⚠️ At-Risk Student Radar',
              'btn_save_att': '💾 Save Attendance',
              'btn_school_admins': '👨\u200d💼 School Administrators',
              'btn_search_again': '🔍 Search Again',
              'btn_search_student': '🔍 Search Student',
              'btn_search_teacher': '🔍 Search Teachers',
              'btn_search_user': '🔍 Search User',
              'btn_send_dm': '✉️ Send Direct Message',
              'btn_send_new_hw': 'Send Homework',
              'btn_switch_student': '🧑\u200d🎓 Switch Student',
              'btn_teachers': '👨\u200d🏫 Teachers',
              'btn_teachers_pdf': '👨\u200d🏫 Teachers Password Cards (PDF)',
              'btn_unack_notifs': '⚠️ Unacknowledged Absences',
              'btn_unban_user': '🟢 Unban User',
              'btn_unlink_parent': '👨\u200d👩\u200d👧\u200d👦 Unlink Parents',
              'btn_upload_excel': '📥 Import Students (Excel)',
              'btn_upload_medical': '🏥 Submit Medical Note',
              'btn_users_hub': '👥 User Directory',
              'btn_users_list': '⬅️ Administrators List',
              'btn_view_cafeteria': '🍲 Daily Cafeteria Menu',
              'btn_view_photo': 'View Photo',
              'btn_view_schedule': '📅 Weekly Timetable',
              'btn_write_to_admin': '💬 Message Administrator',
              'cat_reports_title': '📊 *Academic & Attendance Audit Hub*\nPlease select a report:',
              'cat_requests_title': '🛎️ *Requests & Medical Approvals*\nPlease select an option:',
              'cat_settings_title': '⚙️ *System Settings & Security*\nPlease select an option:',
              'cat_staff_title': '👥 *Staff & Student Management*\nPlease select an option:',
              'cat_tools_title': '🛠️ *Administrative Tools & Communications*\nPlease select an option:',
              'chat_req_error_toast': '⚠️ Cannot reach user: bot is blocked!',
              'chat_req_sent_toast': '✅ 1:1 Contact request delivered!',
              'child_added_success': '✅ *{name}* ({class_name}) added to your account successfully!',
              'cockpit_report': '📊 *Morning Briefing ({date})*\n'
                                '\n'
                                '🏫 Total: *{total}* | ✅ Present: *{present}* | ❌ Absent: *{absent}*\n'
                                '\n'
                                '⚠️ *Pending Attendance Classes ({missing_cnt}):*\n'
                                '{missing}',
              'codes_reset_done': '✅ Credentials regenerated!\n\n• Student: `{st_code}`\n• Parent: `{pr_code}`',
              'contact_req_direct': 'Please initiate a direct conversation via the button below:',
              'contact_req_header': '📞 *ADMINISTRATION CONTACT REQUEST*\n'
                                    '\n'
                                    'School administration requests 1:1 contact with you.\n'
                                    '👤 *Administrator:* {name}\n',
              'contact_req_id': 'Please reach out to school administration.',
              'dm_delivery_error': '⚠️ Delivery Error: User has blocked the bot.',
              'dm_from_admin_header': '📩 *MESSAGE FROM SCHOOL ADMINISTRATION*',
              'dm_sender_label': 'Sender',
              'dm_sent_success': '✅ Message delivered successfully!',
              'evening_briefing_header': '🌙 *DAILY SUMMARY (18:30)*\n'
                                         'Student: *{name}* ({class_name})\n'
                                         '\n'
                                         '📌 Attendance: *{att_status}*\n'
                                         '📝 Grades:\n'
                                         '{grades}',
              'exam_oral': '🗣️ Oral / Performance',
              'exam_written_1': '📝 1st Written Exam',
              'exam_written_2': '📝 2nd Written Exam',
              'excel_done': '✅ Processed! Added students: *{count}*\nCredentials attached.',
              'excel_format_error': '❌ Error processing Excel. Check format.',
              'excel_hub_title': '📥 *Excel Management Center*\nSelect an action:',
              'excel_info': '📥 *Import via Excel*\n'
                            '\n'
                            'Send a `.xlsx` spreadsheet.\n'
                            'Headers: `Ad Soyad` | `Sinif` | `Numara`',
              'export_ready': '📥 *School Data Backup Ready ({date})*',
              'grade_deleted': 'Grade deleted.',
              'grade_parent_notification': '📝 *NEW GRADE ALERT*\n'
                                           '\n'
                                           '🧑\u200d🎓 Student: *{name}*\n'
                                           '📚 Subject: *{subject}* ({exam_type})\n'
                                           '📊 Score: *{score}* ({badge})',
              'grade_saved_success': '✅ Grade sent to parent.',
              'grade_select_class': '📝 Select class to enter grades:',
              'grade_select_student': '📝 *Class {class_name}*\nSelect student to grade:',
              'grade_updated': '✅ Grade updated.',
              'homework_board_title': '📢 *Class {class_name} Homework Board:*',
              'homework_deleted_toast': 'Homework deleted.',
              'hw_sent_success': '📢 Homework dispatched to *{class_name}*.',
              'image_load_error': 'Failed to load image.',
              'invalid_name_error': '❌ Please enter a valid name.',
              'invalid_parent_code': '❌ Invalid parent code!',
              'invalid_phone_error': '❌ Invalid phone number! Please enter at least 7 digits.',
              'invalid_score_format': '❌ Invalid score! Please enter a number (e.g. 85).',
              'invalid_score_range': '❌ Score must be between 0 and 100!',
              'lang_changed': 'Language successfully updated: 🇬🇧 English',
              'lang_select': '🌍 Please select your language / Tilni tanlang / Пожалуйста, выберите язык / Lütfen dil '
                             'seçiniz:',
              'lbl_account_status': 'Account Status',
              'lbl_admin_status': 'Admin Status',
              'lbl_class': 'Class',
              'lbl_full_name': 'Full Name',
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
              'lbl_status_active': '🟢 *Active*',
              'lbl_status_banned': '🚫 *Banned*',
              'lbl_subject': 'Subject',
              'lbl_username': 'Username',
              'lock_countdown_msg': '⛔ *Security Lockout:* Your account is temporarily locked.\n'
                                    '\n'
                                    'Time remaining: *{mins} minutes*.',
              'maintenance_mode': '⚠️ The system is currently under maintenance. Please try again later.',
              'maintenance_mode_updated': 'Maintenance mode updated.',
              'med_uploaded_success': 'Report submitted to school administration.',
              'medical_approved': '✅ Medical excuse approved.',
              'medical_approved_parent': "✅ Your student's medical note was approved by administration.",
              'medical_rejected': '❌ Medical excuse rejected.',
              'medical_rejected_parent': "❌ Your student's medical note was rejected by administration.",
              'menu_parent': '👨\u200d👩\u200d👧\u200d👦 *Parent Dashboard*\nStudent: *{name}* ({class_name})',
              'menu_student': '🎓 *Student Dashboard*\nStudent: *{name}* ({class_name} - Roll: {no})',
              'menu_teacher': '👨\u200d🏫 *Teacher Dashboard*\nTeacher: *{name}* ({subject})',
              'menu_updated': '✅ Cafeteria menu updated.',
              'no_active_homeworks': '📢 No active homework for class *{class_name}*.',
              'no_blacklisted': '✅ No blocked users.',
              'no_classes_found': '⚠️ No classes registered yet.',
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
              'parent_choose_child': '🧑\u200d🎓 Please select a student:',
              'parent_info_title': 'ℹ️ *School Info & Services Board*',
              'parent_settings_title': '⚙️ *Settings & Account*',
              'parent_unlinked_success': '✅ Parent unlinked. New Parent Code: `{code}`',
              'pdf_ready': '📄 Printable cards for *{class_name}* ready.',
              'pdf_report_ready': '📄 Official academic report card for *{name}* is attached.',
              'pending_appointments_title': '🤝 *Pending Parent Meeting Requests:*',
              'pending_medical_title': '🏥 *Pending Medical Reports:*',
              'pending_requests_title': '🛎️ *Pending Access Requests:*',
              'perm_admin_assigned_toast': '✅ User has been granted permanent administrator privileges.',
              'permanent_admin_protected': '⛔ Permanent/Founder administrator privileges cannot be revoked!',
              'permanent_admin_title': 'Permanent Administrator',
              'photo_expected_medical': '⚠️ Please send a photo only.',
              'prompt_add_child_code': '🔑 Enter parent code of the additional child (e.g. `VELI-123456`):',
              'prompt_appointment_note': '📝 Please specify preferred day/time and any note:',
              'prompt_broadcast': '📢 Enter announcement text:',
              'prompt_enter_code_direct': '🔑 *Please enter your access code:* (e.g. `HCA-123456`, `VELI-123456`, '
                                          '`OGR-123456`)',
              'prompt_grade_badge': 'Select performance badge:',
              'prompt_grade_score': 'Student: *{name}* ({class_name})\n'
                                    'Assessment: *{exam_type}*\n'
                                    '\n'
                                    'Enter score (0-100):',
              'prompt_hw_class': 'Select class for homework:',
              'prompt_hw_content': 'Provide description or send blackboard photo:',
              'prompt_menu_update': "🍲 Enter today's cafeteria menu:",
              'prompt_new_score': 'Enter new score (0-100):',
              'prompt_search_student': '🔍 Enter student name or roll number:',
              'prompt_select_exam_type': '📝 Student: *{name}* ({class_name})\n\nPlease select assessment type:',
              'prompt_student_class': '🏫 Enter Student Class (e.g. `9-A`):',
              'prompt_student_name': '👤 Enter Student Full Name:',
              'prompt_student_no': '🔢 Enter Student Roll Number (e.g. `101`):',
              'prompt_teacher_name': '👨\u200d🏫 Enter Teacher Full Name:',
              'prompt_teacher_subject': '📚 Enter Teaching Subject (e.g. `Mathematics`):',
              'published_homeworks_title': '📢 *Published Homeworks:*',
              'recent_grades_title': '📝 *Your Recent Grades:*',
              'remind_att_sent': 'Reminder sent.',
              'report_not_found': 'Medical report not found.',
              'req_already_pending': '⚠️ You already have a pending request.',
              'req_approved_admin_msg': '✅ Request #{id} approved. User: *{name}* ({role})',
              'req_approved_user': '🎉 *Congratulations!*\n'
                                   'School administration approved your request. You logged in as *{role}*.',
              'req_details_parent': "🧑\u200d🎓 Please enter your child's name, class, or roll number:",
              'req_details_student': '🏫 Please enter your class and roll number:',
              'req_details_teacher': '📚 Please enter your teaching subject and a note for administration:',
              'req_name_prompt': '👤 Please enter your Full Name:',
              'req_phone_prompt': '📱 Please enter your phone number (e.g. `+1...` or `+998...`):',
              'req_rejected_admin_msg': '❌ Request #{id} rejected: *{name}*',
              'req_rejected_user': '❌ Administration rejected your request. Please contact the administration.',
              'req_role_select': '🛎️ *Access Request*\n\nPlease select your requested role:',
              'req_sent_success': '✅ Your request has been submitted to administration. You will be notified upon '
                                  'approval.',
              'request_already_handled': '⚠️ This request has already been handled.',
              'rk_appointments': '🤝 Parent Meetings',
              'rk_attendance': '📋 Fast Attendance',
              'rk_cancel_action': '❌ Cancel Action',
              'rk_cat_reports': '📊 Reports & Audits',
              'rk_cat_requests': '🛎️ Approval Center',
              'rk_cat_settings': '⚙️ System & Settings',
              'rk_cat_staff': '👥 Staff & Students',
              'rk_cat_tools': '🛠️ Admin Tools',
              'rk_grade': '📝 Grade Book',
              'rk_homework': '📢 Homework Board',
              'rk_logout': '🚪 Log Out',
              'rk_parent_info': 'ℹ️ School Info Board',
              'rk_report': '📊 Report Card',
              'rk_restart': '🔄 Restart Bot',
              'rk_switch_student': '🧑\u200d🎓 Switch Student',
              'rk_upload_medical': '🏥 Medical Note',
              'role_parent_btn': '👨\u200d👩\u200d👧\u200d👦 Parent',
              'role_student_btn': '🎓 Student',
              'role_teacher_btn': '👨\u200d🏫 Teacher',
              'schedule_select_class': '📅 Select class to view timetable:',
              'school_admin_title': 'School Administration',
              'search_no_results': '❌ No student found.',
              'search_results_title': '🔍 *Search Results:*',
              'search_user_no_results': '❌ No matching user found.',
              'search_user_prompt': '🔍 Enter name, username, or Telegram ID to search:',
              'search_user_results_title': '🔍 *User Search Results:*',
              'select_pdf_class': '📄 Select class for cards:',
              'select_teacher_appointment': '🤝 Select a teacher to meet:',
              'send_dm_prompt': '✉️ *Send Direct Message:*\n\nPlease write the message for user `{name}` (`{id}`):',
              'setting_updated_toast': 'Setting updated.',
              'student_added_card': '✅ *Student Added Successfully!*\n'
                                    '\n'
                                    '👤 Name: *{name}*\n'
                                    '🏫 Class: *{class_name}* | Roll: *{no}*\n'
                                    '\n'
                                    '🔑 *Access Codes:*\n'
                                    '• Student Code: `{st_code}`\n'
                                    '• Parent Code: `{pr_code}`',
              'student_card': '👤 *Student Card*\n'
                              'Name: *{name}*\n'
                              'Class: *{class_name}* | Roll: *{no}*\n'
                              '\n'
                              '🔑 *Code Status:*\n'
                              '• Student: `{st_code}` ({st_status})\n'
                              '• Parent: `{pr_code}` ({pr_status})',
              'student_deleted': '🗑️ Student removed from system.',
              'student_info_updated': '✅ Student details updated:\n*{name}* ({class_name} - Roll: {no})',
              'student_name_invalid': '❌ Please enter a valid name.',
              'student_not_found': 'Student not found.',
              'student_switched_success': 'Active student: *{name}* ({class_name})',
              'teacher_added_card': '✅ *Teacher Registered!*\n'
                                    '\n'
                                    '👤 Name: *{name}*\n'
                                    '📚 Subject: *{subject}*\n'
                                    '\n'
                                    '🔑 *Access Code:*\n'
                                    '`{code}`',
              'teacher_card': '👨\u200d🏫 *Teacher Card*\n'
                              'Name: *{name}*\n'
                              'Subject: *{subject}*\n'
                              '\n'
                              '🔑 Code: `{code}`\n'
                              'Status: {status}',
              'teacher_deleted': '🗑️ Teacher deleted.',
              'teacher_name_invalid': '❌ Please enter a valid teacher name.',
              'teacher_not_found': 'Teacher not found.',
              'teacher_search_no_results': '❌ No matching teacher found.',
              'teacher_search_prompt': '🔍 Enter teacher name or subject:',
              'teacher_search_results_title': '🔍 *Teacher Search Results:*',
              'temp_admin_assigned_toast': '✅ User granted temporary administrator privileges for {dur}.',
              'temp_admin_choose_title': '⏱️ *Select Temporary Administrator Duration:*',
              'uc_card_title': '👤 *USER PROFILE AND ACTION CARD*',
              'unauthorized_action': '⛔ You are not authorized for this action.',
              'unauthorized_excel_upload': '⛔ Unauthorized to upload Excel.',
              'unban_success': 'User unbanned.',
              'upload_med_prompt': 'Please send a photo of the medical note:',
              'user_banned_toast': 'User banned.',
              'user_not_found_toast': 'User not found.',
              'user_unbanned_toast': 'User unbanned.',
              'welcome_guest': '🎓 *Welcome to School Management Ecosystem.*\n'
                               '\n'
                               'Please enter your personal **access code** (e.g. `VELI-123456`, `HCA-123456`, '
                               '`OGR-123456`) or select an action:'},
    'ru': {   'acknowledged_toast': 'Подтверждение принято.',
              'action_cancelled': '❌ *Действие отменено.*',
              'admin_add_name_prompt': '👤 Введите ФИО нового администратора:',
              'admin_add_tg_id_prompt': '➕ Введите Telegram ID нового администратора:',
              'admin_added_success': '✅ *{name}* (`{id}`) успешно добавлен как постоянный администратор.',
              'admin_admins_hub_title': '👨\u200d💼 *Администрация школы и полномочия:*\n'
                                        '\n'
                                        'Нажмите на администратора для просмотра профиля или управления:',
              'admin_code_generated': '🔑 *ОДНОРАЗОВЫЙ КОД АДМИНИСТРАТОРА СОЗДАН*\n'
                                      '\n'
                                      'Код: `{code}`\n'
                                      '\n'
                                      'Передайте этот код пользователю. При отправке кода боту аккаунт получит права '
                                      '*Администратора*.',
              'admin_demoted_notification': 'ℹ️ Ваши права администратора отозваны.',
              'admin_demoted_toast': 'Права администратора отозваны.',
              'admin_invalid_tg_id': '❌ Некорректный ID! Должен содержать только цифры.',
              'admin_promoted_notification': '🎉 *Уважаемый(ая) {name},*\n'
                                             'Вам присвоены права *Постоянного администратора* школы!',
              'admin_restart_confirmed': '🔄 *Панель администратора перезапущена.*',
              'admin_sched_edit_title': '📅 *Редактирование расписания*\nВыберите класс:',
              'admin_sched_updated': '✅ Расписание класса *{class_name}* обновлено.',
              'admin_stats': '📊 *Общий статус:*\n'
                             '• Классы: *{c_cnt}* | Ученики: *{s_cnt}* | Учителя: *{t_cnt}*\n'
                             '• Заявки: *{req_cnt}* | Справки: *{med_cnt}*\n'
                             '• Дата: *{date}*',
              'admin_title': '⚡ *Панель управления школой (Администратор)*',
              'admin_unban_notification': '🟢 Блокировка вашего аккаунта снята администрацией.',
              'admin_user_card_title': '👤 *КАРТОЧКА АДМИНИСТРАТОРА*',
              'all_notifs_acknowledged': '✅ *Все уведомления о пропусках прочитаны родителями!*\n'
                                         '\n'
                                         'За последние 36 часов нет непрочитанных уведомлений.',
              'appointment_approved_msg': '✅ Учитель подтвердил встречу.',
              'appointment_confirmed_toast': 'Встреча подтверждена.',
              'appointment_not_found': '⚠️ Встреча не найдена.',
              'appointment_rejected_msg': '❌ Учитель не может в это время.',
              'appointment_sent': '✅ Запрос на встречу отправлен учителю.',
              'att_check_all_done': '✅ Перекличка всех классов завершена.',
              'att_check_title': '📊 *Контроль переклички ({date})*',
              'att_saved': '✅ Перекличка сохранена (правки 15 мин).',
              'attendance_correction_notification': 'ℹ️ *ИСПРАВЛЕНИЕ:* Запись о пропуске ученика *{name}* исправлена '
                                                    '(ПРИСУТСТВУЕТ).',
              'attendance_hours_lock': '⚠️ Перекличка доступна только с 07:00 до 19:00.',
              'attendance_intro': '📋 *Перекличка {class_name}*\nОтметьте отсутствующих и сохраните:',
              'attendance_select_class': '📋 Выберите класс для переклички:',
              'attendance_weekend_lock': '⚠️ В выходные перекличка недоступна.',
              'auth_blacklisted': '🚫 Ваш аккаунт заблокирован навсегда.',
              'auth_code_already_linked': '⚠️ *Этот код уже привязан к другому аккаунту Telegram.*\n'
                                          '\n'
                                          'Пожалуйста, обратитесь к администрации школы.',
              'auth_failed': '❌ Неверный код доступа! Осталось попыток: {remaining}',
              'auth_locked': '⛔ Аккаунт заблокирован на 1 час из соображений безопасности.',
              'auth_success': '✅ *Авторизация успешна!*\nДобро пожаловать: *{name}*\nВаша роль: *{role}*',
              'badge_missing': '🟡 Пробел / Доработать',
              'badge_praise': '🟢 Похвала / Успех',
              'badge_warning': '🔴 Замечание / Дисциплина',
              'blacklisted_title': '🚫 *Заблокированные пользователи:*',
              'broadcast_success': '📢 Объявление доставлено пользователям: *{count}*.',
              'btn_academic_report': '📈 Рейтинг классов',
              'btn_acknowledged': '✅ Ознакомлен(а)',
              'btn_add_admin_id': '➕ Добавить администратора по Telegram ID',
              'btn_add_another': '➕ Добавить еще',
              'btn_add_new_teacher': '➕ Добавить учителя',
              'btn_add_student': '➕ Добавить ученика',
              'btn_add_teacher': '➕ Добавить учителя',
              'btn_appr_appointment': '✅ Принять',
              'btn_appr_medical': '✅ Одобрить справку',
              'btn_appr_request': '✅ Одобрить',
              'btn_attendance': '📋 Быстрая перекличка',
              'btn_audit_logs': '📜 Журнал действий',
              'btn_back': '⬅️ Назад',
              'btn_ban_user': '🚫 Заблокировать (Бан)',
              'btn_blacklist': '🚫 Заблокированные',
              'btn_briefing_off': '🔕 Вечерняя сводка (ВЫКЛ)',
              'btn_briefing_on': '🔔 Вечерняя сводка (ВКЛ)',
              'btn_broadcast': '📢 Рассылка объявления',
              'btn_cafeteria_edit': '🍲 Обновить меню столовой',
              'btn_cancel_action': '⬅️ Отмена',
              'btn_class_att_sheet': 'Ведомость посещаемости',
              'btn_class_grade_sheet': 'Ведомость оценок',
              'btn_class_pdf_cards': 'Карточки с кодами (PDF)',
              'btn_class_sched': 'Расписание',
              'btn_classes': '🏫 Классы и ученики',
              'btn_cockpit': '📊 Утренний статус',
              'btn_cockpit_unified': '📊 Утренний статус и перекличка',
              'btn_del_grade': '❌ Удалить оценку',
              'btn_del_student': '❌ Удалить ученика',
              'btn_del_teacher': '❌ Удалить учителя',
              'btn_delete_action': '🗑️ Удалить',
              'btn_download_pdf_report': '📄 Скачать официальный табель (PDF)',
              'btn_dur_24h': '⏱️ 24 часа',
              'btn_dur_30d': '⏱️ 30 дней',
              'btn_dur_7d': '⏱️ 7 дней',
              'btn_edit_class': '🏫 Изменить класс',
              'btn_edit_grade': '✏️ Изменить оценку',
              'btn_edit_name': '👤 Изменить имя',
              'btn_edit_no': '🔢 Изменить номер',
              'btn_edit_student': '✏️ Редактировать',
              'btn_enter_grade': '📝 Выставить оценки',
              'btn_excel': '📥 Импорт из Excel',
              'btn_excel_hub': '📥 Центр Excel',
              'btn_export_all_data': '📊 Выгрузить все данные (.xlsx)',
              'btn_gen_admin_code': '🔑 Создать одноразовый код администратора',
              'btn_homework_board': '📢 Доска заданий',
              'btn_lang': '🌐 Сменить язык',
              'btn_login_prompt': '🔑 Войти в систему (Ввести код)',
              'btn_main_menu': '🏠 Главное меню',
              'btn_maintenance_toggle': '🚨 Режим обслуживания ({status})',
              'btn_make_perm_admin': '👑 Сделать постоянным админом',
              'btn_make_temp_admin': '⏱️ Сделать временным админом',
              'btn_manage_schedule': '📅 Расписание уроков',
              'btn_medical': '🏥 Справки ({count})',
              'btn_my_hws': '📚 Мои задания',
              'btn_next': 'Вперед ➡️',
              'btn_not_available': '❌ Не могу',
              'btn_notices': '📢 Объявления школы',
              'btn_pdf': '📄 Карточки с кодами (PDF)',
              'btn_prev': '⬅️ Назад',
              'btn_recent_grades_menu': '🕒 Последние оценки',
              'btn_refresh_data': '🔄 Обновить данные',
              'btn_reject': '❌ Отклонить',
              'btn_remind_att': '⚠️ Напомнить учителям о перекличке',
              'btn_report_card': '📊 Табель успеваемости',
              'btn_req_access': '📩 Запросить доступ',
              'btn_req_chat': '📞 Запросить контакт (1:1)',
              'btn_requests': '🛎️ Заявки ({count})',
              'btn_reset_codes': '🔄 Сбросить коды',
              'btn_revoke_admin_perm': '❌ Отозвать права администратора',
              'btn_risk_radar': '⚠️ Радар успеваемости и рисков',
              'btn_save_att': '💾 Сохранить',
              'btn_school_admins': '👨\u200d💼 Администрация школы',
              'btn_search_again': '🔍 Искать снова',
              'btn_search_student': '🔍 Поиск ученика',
              'btn_search_teacher': '🔍 Поиск учителей',
              'btn_search_user': '🔍 Поиск пользователя',
              'btn_send_dm': '✉️ Отправить личное сообщение',
              'btn_send_new_hw': 'Отправить задание',
              'btn_switch_student': '🧑\u200d🎓 Сменить ученика',
              'btn_teachers': '👨\u200d🏫 Учителя',
              'btn_teachers_pdf': '👨\u200d🏫 Карточки учителей (PDF)',
              'btn_unack_notifs': '⚠️ Непрочитанные пропуски',
              'btn_unban_user': '🟢 Разблокировать (Разбан)',
              'btn_unlink_parent': '👨\u200d👩\u200d👧\u200d👦 Отвязать родителей',
              'btn_upload_excel': '📥 Загрузить учеников (Excel)',
              'btn_upload_medical': '🏥 Отправить справку',
              'btn_users_hub': '👥 Пользователи',
              'btn_users_list': '⬅️ Список администраторов',
              'btn_view_cafeteria': '🍲 Меню столовой',
              'btn_view_photo': 'Открыть фото',
              'btn_view_schedule': '📅 Расписание уроков',
              'btn_write_to_admin': '💬 Написать администратору',
              'cat_reports_title': '📊 *Академический контроль и отчеты*\nВыберите раздел:',
              'cat_requests_title': '🛎️ *Заявки и медицинские справки*\nВыберите действие:',
              'cat_settings_title': '⚙️ *Системные настройки и безопасность*\nВыберите действие:',
              'cat_staff_title': '👥 *Управление учениками и учителями*\nВыберите действие:',
              'cat_tools_title': '🛠️ *Инструменты управления и рассылки*\nВыберите раздел:',
              'chat_req_error_toast': '⚠️ Не удалось связаться: пользователь заблокировал бота!',
              'chat_req_sent_toast': '✅ Запрос на контакт отправлен!',
              'child_added_success': '✅ *{name}* ({class_name}) успешно добавлен к вашему аккаунту!',
              'cockpit_report': '📊 *Утренняя сводка ({date})*\n'
                                '\n'
                                '🏫 Всего: *{total}* | ✅ Есть: *{present}* | ❌ Нет: *{absent}*\n'
                                '\n'
                                '⚠️ *Классы без переклички ({missing_cnt}):*\n'
                                '{missing}',
              'codes_reset_done': '✅ Коды обновлены!\n'
                                  '\n'
                                  '• Новый код ученика: `{st_code}`\n'
                                  '• Новый код родителя: `{pr_code}`',
              'contact_req_direct': 'Пожалуйста, начните диалог по кнопке ниже:',
              'contact_req_header': '📞 *ВЫЗОВ НА СВЯЗЬ ОТ АДМИНИСТРАЦИИ*\n'
                                    '\n'
                                    'Администрация школы приглашает вас к личному общению (1:1).\n'
                                    '👤 *Администратор:* {name}\n',
              'contact_req_id': 'Пожалуйста, напишите администрации школы.',
              'dm_delivery_error': '⚠️ Ошибка доставки: Пользователь заблокировал бота.',
              'dm_from_admin_header': '📩 *СООБЩЕНИЕ ОТ АДМИНИСТРАЦИИ*',
              'dm_sender_label': 'Отправитель',
              'dm_sent_success': '✅ Сообщение успешно доставлено!',
              'evening_briefing_header': '🌙 *ИТОГИ ДНЯ (18:30)*\n'
                                         'Ученик: *{name}* ({class_name})\n'
                                         '\n'
                                         '📌 Посещаемость: *{att_status}*\n'
                                         '📝 Оценки:\n'
                                         '{grades}',
              'exam_oral': '🗣️ Устный опрос / Активность',
              'exam_written_1': '📝 1-я Контрольная',
              'exam_written_2': '📝 2-я Контрольная',
              'excel_done': '✅ Обработано! Добавлено учеников: *{count}*\nКоды прикреплены в файле.',
              'excel_format_error': '❌ Ошибка при обработке Excel. Проверьте формат.',
              'excel_hub_title': '📥 *Центр Excel*\nВыберите действие:',
              'excel_info': '📥 *Импорт через Excel*\n'
                            '\n'
                            'Отправьте файл `.xlsx`.\n'
                            'Заголовки: `Ad Soyad` | `Sinif` | `Numara`',
              'export_ready': '📥 *Архив данных школы готов ({date})*',
              'grade_deleted': 'Оценка удалена.',
              'grade_parent_notification': '📝 *НОВАЯ ОЦЕНКА*\n'
                                           '\n'
                                           '🧑\u200d🎓 Ученик: *{name}*\n'
                                           '📚 Предмет: *{subject}* ({exam_type})\n'
                                           '📊 Оценка: *{score}* ({badge})',
              'grade_saved_success': '✅ Оценка отправлена родителю.',
              'grade_select_class': '📝 Выберите класс для выставления оценок:',
              'grade_select_student': '📝 *Класс {class_name}*\nВыберите ученика для оценки:',
              'grade_updated': '✅ Оценка обновлена.',
              'homework_board_title': '📢 *Доска заданий класса {class_name}:*',
              'homework_deleted_toast': 'Задание удалено.',
              'hw_sent_success': '📢 Задание отправлено классу *{class_name}*.',
              'image_load_error': 'Не удалось загрузить фото.',
              'invalid_name_error': '❌ Пожалуйста, введите корректное имя.',
              'invalid_parent_code': '❌ Неверный код родителя!',
              'invalid_phone_error': '❌ Некорректный номер! Введите не менее 7 цифр.',
              'invalid_score_format': '❌ Некорректный балл! Введите число (напр: 85).',
              'invalid_score_range': '❌ Оценка должна быть от 0 до 100!',
              'lang_changed': 'Язык успешно изменен: 🇷🇺 Русский',
              'lang_select': '🌍 Пожалуйста, выберите язык:',
              'lbl_account_status': 'Состояние аккаунта',
              'lbl_admin_status': 'Статус администратора',
              'lbl_class': 'Класс',
              'lbl_full_name': 'ФИО',
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
              'lbl_status_active': '🟢 *Активен*',
              'lbl_status_banned': '🚫 *Заблокирован (Бан)*',
              'lbl_subject': 'Предмет',
              'lbl_username': 'Имя пользователя',
              'lock_countdown_msg': '⛔ *Блокировка безопасности:* Аккаунт временно заблокирован.\n'
                                    '\n'
                                    'Осталось: *{mins} мин.*',
              'maintenance_mode': '⚠️ В системе ведутся технические работы. Пожалуйста, попробуйте позже.',
              'maintenance_mode_updated': 'Режим обслуживания обновлен.',
              'med_uploaded_success': 'Справка отправлена администрации.',
              'medical_approved': '✅ Справка одобрена.',
              'medical_approved_parent': '✅ Справка вашего ребенка одобрена администрацией.',
              'medical_rejected': '❌ Справка отклонена.',
              'medical_rejected_parent': '❌ Справка вашего ребенка отклонена администрацией.',
              'menu_parent': '👨\u200d👩\u200d👧\u200d👦 *Панель родителя*\nУченик: *{name}* ({class_name})',
              'menu_student': '🎓 *Панель ученика*\nУченик: *{name}* ({class_name} - №: {no})',
              'menu_teacher': '👨\u200d🏫 *Панель учителя*\nУчитель: *{name}* ({subject})',
              'menu_updated': '✅ Меню столовой обновлено.',
              'no_active_homeworks': '📢 Для класса *{class_name}* нет активных заданий.',
              'no_blacklisted': '✅ Нет заблокированных пользователей.',
              'no_classes_found': '⚠️ Пока не зарегистрировано ни одного класса.',
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
              'parent_choose_child': '🧑\u200d🎓 Выберите ученика:',
              'parent_info_title': 'ℹ️ *Информационная панель школы*',
              'parent_settings_title': '⚙️ *Настройки и аккаунт*',
              'parent_unlinked_success': '✅ Родители отвязаны. Новый код родителя: `{code}`',
              'pdf_ready': '📄 Карточки для *{class_name}* готовы.',
              'pdf_report_ready': '📄 Официальный табель успеваемости ученика *{name}* прикреплен.',
              'pending_appointments_title': '🤝 *Запросы родителей на встречу:*',
              'pending_medical_title': '🏥 *Справки на рассмотрении:*',
              'pending_requests_title': '🛎️ *Заявки на рассмотрении:*',
              'perm_admin_assigned_toast': '✅ Пользователю присвоены права постоянного администратора.',
              'permanent_admin_protected': '⛔ Права главного/постоянного администратора не могут быть отозваны!',
              'permanent_admin_title': 'Постоянный администратор',
              'photo_expected_medical': '⚠️ Пожалуйста, отправьте только фотографию.',
              'prompt_add_child_code': '🔑 Введите код родителя второго ребенка (Например: `VELI-123456`):',
              'prompt_appointment_note': '📝 Укажите удобное время встречи и примечание:',
              'prompt_broadcast': '📢 Введите текст объявления:',
              'prompt_enter_code_direct': '🔑 *Введите ваш код доступа:* (Например: `HCA-123456`, `VELI-123456`, '
                                          '`OGR-123456`)',
              'prompt_grade_badge': 'Выберите категорию оценки:',
              'prompt_grade_score': 'Ученик: *{name}* ({class_name})\n'
                                    'Вид оценки: *{exam_type}*\n'
                                    '\n'
                                    'Введите оценку (0-100):',
              'prompt_hw_class': 'Выберите класс для задания:',
              'prompt_hw_content': 'Отправьте текст задания или фото доски:',
              'prompt_menu_update': '🍲 Введите сегодняшнее меню столовой:',
              'prompt_new_score': 'Введите новую оценку (0-100):',
              'prompt_search_student': '🔍 Введите фамилию или номер ученика:',
              'prompt_select_exam_type': '📝 Ученик: *{name}* ({class_name})\n\nВыберите вид оценивания:',
              'prompt_student_class': '🏫 Введите класс ученика (Например: `9-A`):',
              'prompt_student_name': '👤 Введите Фамилию и Имя ученика:',
              'prompt_student_no': '🔢 Введите номер ученика (Например: `101`):',
              'prompt_teacher_name': '👨\u200d🏫 Введите ФИО учителя:',
              'prompt_teacher_subject': '📚 Введите предмет (Например: `Математика`):',
              'published_homeworks_title': '📢 *Опубликованные задания:*',
              'recent_grades_title': '📝 *Ваши последние оценки:*',
              'remind_att_sent': 'Напоминание отправлено.',
              'report_not_found': 'Справка не найдена.',
              'req_already_pending': '⚠️ У вас уже есть заявка на рассмотрении.',
              'req_approved_admin_msg': '✅ Заявка #{id} одобрена. Пользователь: *{name}* ({role})',
              'req_approved_user': '🎉 *Поздравляем!*\nАдминистрация одобрила вашу заявку. Вы вошли как *{role}*.',
              'req_details_parent': '🧑\u200d🎓 Укажите ФИО, класс или номер вашего ребенка:',
              'req_details_student': '🏫 Укажите ваш класс и номер в школе:',
              'req_details_teacher': '📚 Укажите ваш предмет и примечание для администрации:',
              'req_name_prompt': '👤 Пожалуйста, введите ваши Фамилию и Имя:',
              'req_phone_prompt': '📱 Введите номер телефона (Например: `+7...` или `+998...`):',
              'req_rejected_admin_msg': '❌ Заявка #{id} отклонена: *{name}*',
              'req_rejected_user': '❌ Администрация отклонила вашу заявку. Обратитесь к администрации.',
              'req_role_select': '🛎️ *Запрос доступа*\n\nВыберите роль для запроса:',
              'req_sent_success': '✅ Ваша заявка отправлена администрации школы. Вы получите уведомление после '
                                  'проверки.',
              'request_already_handled': '⚠️ Эта заявка уже обработана.',
              'rk_appointments': '🤝 Записи родителей',
              'rk_attendance': '📋 Быстрая перекличка',
              'rk_cancel_action': '❌ Отменить действие',
              'rk_cat_reports': '📊 Отчеты и контроль',
              'rk_cat_requests': '🛎️ Центр одобрений',
              'rk_cat_settings': '⚙️ Настройки системы',
              'rk_cat_staff': '👥 Ученики и учителя',
              'rk_cat_tools': '🛠️ Инструменты',
              'rk_grade': '📝 Выставить оценки',
              'rk_homework': '📢 Доска заданий',
              'rk_logout': '🚪 Выйти',
              'rk_parent_info': 'ℹ️ Инфопанель школы',
              'rk_report': '📊 Табель успеваемости',
              'rk_restart': '🔄 Перезапуск',
              'rk_switch_student': '🧑\u200d🎓 Сменить ученика',
              'rk_upload_medical': '🏥 Отправить справку',
              'role_parent_btn': '👨\u200d👩\u200d👧\u200d👦 Родитель',
              'role_student_btn': '🎓 Ученик',
              'role_teacher_btn': '👨\u200d🏫 Учитель',
              'schedule_select_class': '📅 Выберите класс для расписания:',
              'school_admin_title': 'Администрация школы',
              'search_no_results': '❌ Ученик не найден.',
              'search_results_title': '🔍 *Результаты поиска:*',
              'search_user_no_results': '❌ Пользователь не найден.',
              'search_user_prompt': '🔍 Введите имя, юзернейм или Telegram ID пользователя:',
              'search_user_results_title': '🔍 *Результаты поиска пользователей:*',
              'select_pdf_class': '📄 Выберите класс для карточек:',
              'select_teacher_appointment': '🤝 Выберите учителя для встречи:',
              'send_dm_prompt': '✉️ *Отправить личное сообщение:*\n'
                                '\n'
                                'Введите сообщение для пользователя `{name}` (`{id}`):',
              'setting_updated_toast': 'Настройка обновлена.',
              'student_added_card': '✅ *Ученик успешно добавлен!*\n'
                                    '\n'
                                    '👤 ФИО: *{name}*\n'
                                    '🏫 Класс: *{class_name}* | №: *{no}*\n'
                                    '\n'
                                    '🔑 *Коды доступа:*\n'
                                    '• Код ученика: `{st_code}`\n'
                                    '• Код родителя: `{pr_code}`',
              'student_card': '👤 *Карточка ученика*\n'
                              'ФИО: *{name}*\n'
                              'Класс: *{class_name}* | №: *{no}*\n'
                              '\n'
                              '🔑 *Статус кодов:*\n'
                              '• Ученик: `{st_code}` ({st_status})\n'
                              '• Родитель: `{pr_code}` ({pr_status})',
              'student_deleted': '🗑️ Ученик удален из системы.',
              'student_info_updated': '✅ Данные ученика обновлены:\n*{name}* ({class_name} - №: {no})',
              'student_name_invalid': '❌ Введите корректное имя.',
              'student_not_found': 'Ученик не найден.',
              'student_switched_success': 'Выбран ученик: *{name}* ({class_name})',
              'teacher_added_card': '✅ *Учитель зарегистрирован!*\n'
                                    '\n'
                                    '👤 ФИО: *{name}*\n'
                                    '📚 Предмет: *{subject}*\n'
                                    '\n'
                                    '🔑 *Код доступа:*\n'
                                    '`{code}`',
              'teacher_card': '👨\u200d🏫 *Карточка учителя*\n'
                              'ФИО: *{name}*\n'
                              'Предмет: *{subject}*\n'
                              '\n'
                              '🔑 Код: `{code}`\n'
                              'Статус: {status}',
              'teacher_deleted': '🗑️ Учитель удален.',
              'teacher_name_invalid': '❌ Введите корректное имя учителя.',
              'teacher_not_found': 'Учитель не найден.',
              'teacher_search_no_results': '❌ Учитель не найден.',
              'teacher_search_prompt': '🔍 Введите ФИО учителя или предмет:',
              'teacher_search_results_title': '🔍 *Результаты поиска учителей:*',
              'temp_admin_assigned_toast': '✅ Пользователь назначен временным администратором на {dur}.',
              'temp_admin_choose_title': '⏱️ *Выберите срок временных прав администратора:*',
              'uc_card_title': '👤 *КАРТОЧКА ПОЛЬЗОВАТЕЛЯ И ДЕЙСТВИЯ*',
              'unauthorized_action': '⛔ У вас нет прав для этого действия.',
              'unauthorized_excel_upload': '⛔ У вас нет прав загрузки Excel.',
              'unban_success': 'Пользователь разблокирован.',
              'upload_med_prompt': 'Пожалуйста, отправьте фото справки:',
              'user_banned_toast': 'Пользователь заблокирован.',
              'user_not_found_toast': 'Пользователь не найден.',
              'user_unbanned_toast': 'Пользователь разблокирован.',
              'welcome_guest': '🎓 *Добро пожаловать в систему управления школой.*\n'
                               '\n'
                               'Введите код доступа (Напр: `VELI-123456`, `HCA-123456`, `OGR-123456`) или выберите '
                               'действие:'},
    'tr': {   'acknowledged_toast': 'Onayınız kaydedildi.',
              'action_cancelled': '❌ *İşlem iptal edildi.*',
              'admin_add_name_prompt': '👤 Yeni yöneticinin Adını ve Soyadını yazınız:',
              'admin_add_tg_id_prompt': '➕ Yönetici yapmak istediğiniz kişinin Telegram ID numarasını yazınız:',
              'admin_added_success': '✅ *{name}* (`{id}`) başarıyla kalıcı yönetici olarak eklendi.',
              'admin_admins_hub_title': '👨\u200d💼 *Okul Yönetici Kadrosu & İdari Yetkiler:*\n'
                                        '\n'
                                        'Detaylarını görmek veya yetkilerini yönetmek için bir yöneticiye dokunun:',
              'admin_code_generated': '🔑 *TEK KULLANIMLIK YÖNETİCİ KODU ÜRETİLDİ*\n'
                                      '\n'
                                      'Kod: `{code}`\n'
                                      '\n'
                                      'Bu kodu yetkilendirmek istediğiniz kişiye iletiniz. Kişi bota bu kodu yazdığı '
                                      'anda hesabı *Okul İdaresi (Admin)* rolüne yükselecektir.',
              'admin_demoted_notification': 'ℹ️ Yönetici yetkileriniz okul idaresi tarafından sonlandırıldı.',
              'admin_demoted_toast': 'Yönetici yetkisi kaldırıldı.',
              'admin_invalid_tg_id': '❌ Geçersiz Telegram ID! Sadece rakamlardan oluşmalıdır.',
              'admin_promoted_notification': '🎉 *Sayın {name},*\n'
                                             'Okul yönetim sisteminde *Kalıcı Yönetici (Admin)* olarak '
                                             'yetkilendirildiniz!',
              'admin_restart_confirmed': '🔄 *Yönetici Paneli Yeniden Başlatıldı.*',
              'admin_sched_edit_title': '📅 *Ders Programı Düzenleme*\nLütfen sınıf seçiniz:',
              'admin_sched_updated': '✅ *{class_name}* sınıfı ders programı güncellendi.',
              'admin_stats': '📊 *Genel Durum:*\n'
                             '• Sınıflar: *{c_cnt}* | Öğrenciler: *{s_cnt}* | Öğretmenler: *{t_cnt}*\n'
                             '• Bekleyen Başvurular: *{req_cnt}* | Mazeretler: *{med_cnt}*\n'
                             '• Tarih: *{date}*',
              'admin_title': '⚡ *Okul Yönetim Kokpiti (Admin)*',
              'admin_unban_notification': '🟢 Hesabınızın engeli okul idaresi tarafından kaldırıldı.',
              'admin_user_card_title': '👤 *YÖNETİCİ KULLANICI KARTI*',
              'all_notifs_acknowledged': '✅ *Tüm Devamsızlık Bildirimleri Velilerce Onaylandı!*\n'
                                         '\n'
                                         'Son 36 saatte velisi tarafından okunmamış hiçbir bildirim bulunmamaktadır.',
              'appointment_approved_msg': '✅ Öğretmen randevu talebinizi kabul etti.',
              'appointment_confirmed_toast': 'Randevu onaylandı.',
              'appointment_not_found': '⚠️ Randevu bulunamadı.',
              'appointment_rejected_msg': '❌ Öğretmen belirtilen saatte müsait olmadığını bildirdi.',
              'appointment_sent': '✅ Randevu talebiniz öğretmene iletildi.',
              'att_check_all_done': '✅ Tüm sınıfların yoklaması eksiksiz alınmıştır.',
              'att_check_title': '📊 *Günlük Yoklama Denetimi ({date})*',
              'att_saved': '✅ Yoklama kaydedildi. 15 dakikalık düzeltme süresi başladı.',
              'attendance_correction_notification': 'ℹ️ *DÜZELTME:* Öğrenciniz *{name}* hakkındaki devamsızlık kaydı '
                                                    'düzeltilmiştir (GELDI).',
              'attendance_hours_lock': '⚠️ Yoklama sadece 07:00 - 19:00 saatleri arasında alınabilir.',
              'attendance_intro': '📋 *{class_name} Yoklaması*\nGelmeyenlerin üzerine tıklayıp kaydediniz:',
              'attendance_select_class': '📋 Yoklama almak istediğiniz sınıfı seçiniz:',
              'attendance_weekend_lock': '⚠️ Hafta sonu yoklama girişi yapılamaz.',
              'auth_blacklisted': '🚫 Hesabınız güvenlik nedeniyle kalıcı olarak askıya alındı.',
              'auth_code_already_linked': '⚠️ *Bu kod zaten başka bir Telegram hesabına bağlanmıştır.*\n'
                                          '\n'
                                          'Kodunuzun çalındığını düşünüyorsanız lütfen derhal okul idaresine '
                                          'başvurunuz.',
              'auth_failed': '❌ Geçersiz giriş kodu! Kalan deneme hakkınız: {remaining}',
              'auth_locked': '⛔ Güvenlik nedeniyle hesabınız 1 saat süreyle kilitlendi.',
              'auth_success': '✅ *Giriş Başarılı!*\nHoş geldiniz: *{name}*\nRolünüz: *{role}*',
              'badge_missing': '🟡 Eksik / Tekrar',
              'badge_praise': '🟢 Başarılı / Övgü',
              'badge_warning': '🔴 Uyarı / Dikkat',
              'blacklisted_title': '🚫 *Engelli & Kilitli Kullanıcılar:*',
              'broadcast_success': '📢 Duyuru *{count}* kullanıcıya iletildi.',
              'btn_academic_report': '📈 Başarı Analizi',
              'btn_acknowledged': '✅ Okudum / Bilgilendirildim',
              'btn_add_admin_id': '➕ Telegram ID ile Yönetici Ekle',
              'btn_add_another': '➕ Başka Ekle',
              'btn_add_new_teacher': '➕ Yeni Öğretmen Ekle',
              'btn_add_student': '➕ Öğrenci Ekle',
              'btn_add_teacher': '➕ Öğretmen Ekle',
              'btn_appr_appointment': '✅ Kabul Et',
              'btn_appr_medical': '✅ Onayla (İzinli Say)',
              'btn_appr_request': '✅ Onayla',
              'btn_attendance': '📋 Hızlı Yoklama',
              'btn_audit_logs': '📜 İşlem Geçmişi (Audit Log)',
              'btn_back': '⬅️ Geri',
              'btn_ban_user': '🚫 Engelle (Ban)',
              'btn_blacklist': '🚫 Engelli Kullanıcılar',
              'btn_briefing_off': '🔕 Gün Sonu Bülteni (KAPALI)',
              'btn_briefing_on': '🔔 Gün Sonu Bülteni (AÇIK)',
              'btn_broadcast': '📢 Toplu Duyuru',
              'btn_cafeteria_edit': '🍲 Günün Menüsünü Güncelle',
              'btn_cancel_action': '⬅️ İptal',
              'btn_class_att_sheet': 'Devamsızlık Çizelgesi',
              'btn_class_grade_sheet': 'Not Çizelgesi',
              'btn_class_pdf_cards': 'Şifre Kartları (PDF)',
              'btn_class_sched': 'Programı',
              'btn_classes': '🏫 Sınıflar & Öğrenciler',
              'btn_cockpit': '📊 Sabah Kokpiti',
              'btn_cockpit_unified': '📊 Sabah Kokpiti & Yoklama',
              'btn_del_grade': '❌ Notu Sil',
              'btn_del_student': '❌ Öğrenciyi Sil',
              'btn_del_teacher': '❌ Öğretmeni Sil',
              'btn_delete_action': '🗑️ Sil',
              'btn_download_pdf_report': '📄 Resmi PDF Karne İndir',
              'btn_dur_24h': '⏱️ 24 Saat',
              'btn_dur_30d': '⏱️ 30 Gün',
              'btn_dur_7d': '⏱️ 7 Gün',
              'btn_edit_class': '🏫 Sınıf Düzenle',
              'btn_edit_grade': '✏️ Notu Düzenle',
              'btn_edit_name': '👤 İsim Düzenle',
              'btn_edit_no': '🔢 Numara Düzenle',
              'btn_edit_student': '✏️ Bilgileri Düzenle',
              'btn_enter_grade': '📝 Not Girişi',
              'btn_excel': '📥 Excel ile Yükle',
              'btn_excel_hub': '📥 Excel Merkezi',
              'btn_export_all_data': '📊 Tüm Okul Verilerini İndir (.xlsx)',
              'btn_gen_admin_code': '🔑 Tek Kullanımlık Yönetici Kodu Üret',
              'btn_homework_board': '📢 Ödev Panosu',
              'btn_lang': '🌐 Dil Değiştir',
              'btn_login_prompt': '🔑 Sisteme Giriş Yap (Kod Gir)',
              'btn_main_menu': '🏠 Ana Menü',
              'btn_maintenance_toggle': '🚨 Bakım Modu ({status})',
              'btn_make_perm_admin': '👑 Kalıcı Yönetici Yap',
              'btn_make_temp_admin': '⏱️ Geçici Yönetici Yap',
              'btn_manage_schedule': '📅 Ders Programı Yönetimi',
              'btn_medical': '🏥 Mazeretler ({count})',
              'btn_my_hws': '📚 Gönderdiğim Ödevler',
              'btn_next': 'Sonraki ➡️',
              'btn_not_available': '❌ Uygun Değil',
              'btn_notices': '📢 Okul Duyuruları',
              'btn_pdf': '📄 Şifre Kartları (PDF)',
              'btn_prev': '⬅️ Önceki',
              'btn_recent_grades_menu': '🕒 Son Notlar & Düzenle',
              'btn_refresh_data': '🔄 Verileri Yenile',
              'btn_reject': '❌ Reddet',
              'btn_remind_att': '⚠️ Öğretmenlere Yoklama Hatırlat',
              'btn_report_card': '📊 Durum Paneli (Karne)',
              'btn_req_access': '📩 Şifre / Erişim Talep Et',
              'btn_req_chat': '📞 1:1 İletişim İsteği Gönder',
              'btn_requests': '🛎️ Başvurular ({count})',
              'btn_reset_codes': '🔄 Kodları Sıfırla',
              'btn_revoke_admin_perm': '❌ Yönetici Yetkisini Al',
              'btn_risk_radar': '⚠️ Riskli Öğrenci Radarı',
              'btn_save_att': '💾 Yoklamayı Kaydet',
              'btn_school_admins': '👨\u200d💼 Okul Yöneticileri',
              'btn_search_again': '🔍 Tekrar Ara',
              'btn_search_student': '🔍 Öğrenci Ara',
              'btn_search_teacher': '🔍 Öğretmen Ara',
              'btn_search_user': '🔍 Kullanıcı Ara',
              'btn_send_dm': '✉️ Özel Mesaj Gönder',
              'btn_send_new_hw': 'Yeni Ödev Gönder',
              'btn_switch_student': '🧑\u200d🎓 Öğrenci Değiştir',
              'btn_teachers': '👨\u200d🏫 Öğretmenler',
              'btn_teachers_pdf': '👨\u200d🏫 Öğretmen Şifre Kartları (PDF)',
              'btn_unack_notifs': '⚠️ Okunmamış Devamsızlıklar',
              'btn_unban_user': '🟢 Engeli Kaldır (Unban)',
              'btn_unlink_parent': '👨\u200d👩\u200d👧\u200d👦 Veli Bağlantısını Kopar',
              'btn_upload_excel': '📥 Excel ile Öğrenci Yükle',
              'btn_upload_medical': '🏥 Mazeret / Rapor Yükle',
              'btn_users_hub': '👥 Kullanıcı Rehberi',
              'btn_users_list': '⬅️ Yöneticiler Listesi',
              'btn_view_cafeteria': '🍲 Günün Yemek Menüsü',
              'btn_view_photo': 'Görseli Aç',
              'btn_view_schedule': '📅 Haftalık Ders Programı',
              'btn_write_to_admin': '💬 İdareciye Mesaj Yaz',
              'cat_reports_title': '📊 *Akademik & Yoklama Denetim Masası*\n'
                                   'Lütfen incelemek istediğiniz raporu seçiniz:',
              'cat_requests_title': '🛎️ *Başvuru & Mazeret Onay Masası*\nLütfen işlem seçiniz:',
              'cat_settings_title': '⚙️ *Sistem & Güvenlik Ayarları*\nLütfen işlem seçiniz:',
              'cat_staff_title': '👥 *Okul Kadrosu & Öğrenci Yönetimi*\nLütfen işlem seçiniz:',
              'cat_tools_title': '🛠️ *İdari Araçlar & İletişim Masası*\nLütfen işlem seçiniz:',
              'chat_req_error_toast': '⚠️ Kullanıcı botu engellediği için iletilemedi!',
              'chat_req_sent_toast': '✅ 1:1 İletişim isteği kullanıcıya iletildi!',
              'child_added_success': '✅ *{name}* ({class_name}) başarıyla hesabınıza eklendi!',
              'cockpit_report': '📊 *Sabah Kokpiti ({date})*\n'
                                '\n'
                                '🏫 Toplam: *{total}* | ✅ Gelen: *{present}* | ❌ Gelmeyen: *{absent}*\n'
                                '\n'
                                '⚠️ *Yoklama Almayan Sınıflar ({missing_cnt}):*\n'
                                '{missing}',
              'codes_reset_done': '✅ Giriş kodları yenilendi!\n\n• Yeni Öğrenci: `{st_code}`\n• Yeni Veli: `{pr_code}`',
              'contact_req_direct': 'Lütfen aşağıdaki butondan doğrudan mesajlaşmayı başlatınız:',
              'contact_req_header': '📞 *OKUL İDARESİ İLETİŞİM ÇAĞRISI*\n'
                                    '\n'
                                    'Okul yönetimi sizinle 1:1 iletişime geçmek istemektedir.\n'
                                    '👤 *İdareci:* {name}\n',
              'contact_req_id': 'Lütfen okul idaresine yazınız.',
              'dm_delivery_error': '⚠️ İletim Hatası: Kullanıcı botu engellemiş.',
              'dm_from_admin_header': '📩 *OKUL İDARESİNDEN BİLDİRİM*',
              'dm_sender_label': 'Gönderen',
              'dm_sent_success': '✅ Mesaj kullanıcıya başarıyla iletildi!',
              'evening_briefing_header': '🌙 *GÜN SONU BÜLTENİ (18:30)*\n'
                                         'Öğrenci: *{name}* ({class_name})\n'
                                         '\n'
                                         '📌 Devamsızlık: *{att_status}*\n'
                                         '📝 Notlar:\n'
                                         '{grades}',
              'exam_oral': '🗣️ Sözlü / Performans',
              'exam_written_1': '📝 1. Yazılı',
              'exam_written_2': '📝 2. Yazılı',
              'excel_done': '✅ Excel işlendi! Eklenen öğrenci: *{count}*\nŞifreler ekteki dosyada üretildi.',
              'excel_format_error': '❌ Excel dosyası işlenirken hata oluştu. Lütfen formatı kontrol ediniz.',
              'excel_hub_title': '📥 *Excel Yönetim Merkezi*\nLütfen işlem seçiniz:',
              'excel_info': '📥 *Excel ile Toplu Öğrenci Yükleme*\n'
                            '\n'
                            'Lütfen `.xlsx` dosyasını gönderiniz.\n'
                            'Başlıklar: `Ad Soyad` | `Sinif` | `Numara`',
              'export_ready': '📥 *Okul Veri Yedeği Hazır ({date})*',
              'grade_deleted': 'Not silindi.',
              'grade_parent_notification': '📝 *YENİ DERS NOTU BİLDİRİMİ*\n'
                                           '\n'
                                           '🧑\u200d🎓 Öğrenci: *{name}*\n'
                                           '📚 Ders: *{subject}* ({exam_type})\n'
                                           '📊 Not: *{score}* ({badge})',
              'grade_saved_success': '✅ Not kaydedildi ve veliye bildirildi.',
              'grade_select_class': '📝 Not girişi yapmak istediğiniz sınıfı seçiniz:',
              'grade_select_student': '📝 *{class_name} Sınıfı*\nNot girmek istediğiniz öğrenciyi seçiniz:',
              'grade_updated': '✅ Not güncellendi.',
              'homework_board_title': '📢 *{class_name} Sınıfı Ödev Panosu:*',
              'homework_deleted_toast': 'Ödev silindi.',
              'hw_sent_success': '📢 Ödev *{class_name}* sınıfına iletildi.',
              'image_load_error': 'Görsel yüklenemedi.',
              'invalid_name_error': '❌ Lütfen geçerli bir ad soyad giriniz.',
              'invalid_parent_code': '❌ Geçersiz veli kodu!',
              'invalid_phone_error': '❌ Geçersiz telefon numarası! Lütfen en az 7 haneli geçerli bir numara giriniz.',
              'invalid_score_format': '❌ Geçersiz not! Lütfen sayı giriniz (Örn: 85).',
              'invalid_score_range': '❌ Not 0 ile 100 arasında olmalıdır!',
              'lang_changed': 'Dil başarıyla güncellendi: 🇹🇷 Türkçe',
              'lang_select': '🌍 Lütfen bir dil seçiniz / Tilni tanlang / Пожалуйста, выберите язык / Please select '
                             'language:',
              'lbl_account_status': 'Hesap Durumu',
              'lbl_admin_status': 'Yönetici Statüsü',
              'lbl_class': 'Sınıf',
              'lbl_full_name': 'Ad Soyad',
              'lbl_lang': 'Dil',
              'lbl_linked_students': 'Bağlı Öğrenciler',
              'lbl_not_admin': 'Standart Kullanıcı (Yönetici Değil)',
              'lbl_number': 'Numara',
              'lbl_phone': 'Telefon',
              'lbl_role': 'Rol',
              'lbl_role_admin': 'Okul İdaresi (Admin)',
              'lbl_role_guest': 'Misafir',
              'lbl_role_parent': 'Veli',
              'lbl_role_student': 'Öğrenci',
              'lbl_role_teacher': 'Öğretmen',
              'lbl_status_active': '🟢 *Aktif*',
              'lbl_status_banned': '🚫 *Engelli (Banlı)*',
              'lbl_subject': 'Branş',
              'lbl_username': 'Kullanıcı Adı',
              'lock_countdown_msg': '⛔ *Güvenlik Kilidi:* Hesabınız geçici olarak kilitlidir.\n'
                                    '\n'
                                    'Kalan süre: *{mins} dakika*.',
              'maintenance_mode': '⚠️ Sistem şu anda planlı bakım modundadır. Lütfen daha sonra tekrar deneyiniz.',
              'maintenance_mode_updated': 'Bakım modu güncellendi.',
              'med_uploaded_success': 'Rapor okul idaresine iletildi.',
              'medical_approved': '✅ Rapor onaylandı. Öğrenci izinli sayıldı.',
              'medical_approved_parent': '✅ Öğrencinizin sağlık raporu okul idaresince onaylanmış ve izinli '
                                         'kaydedilmiştir.',
              'medical_rejected': '❌ Rapor reddedildi.',
              'medical_rejected_parent': '❌ Öğrencinizin sağlık raporu okul idaresince reddedilmiştir.',
              'menu_parent': '👨\u200d👩\u200d👧\u200d👦 *Veli Masası*\nÖğrenci: *{name}* ({class_name})',
              'menu_student': '🎓 *Öğrenci Masası*\nÖğrenci: *{name}* ({class_name} - No: {no})',
              'menu_teacher': '👨\u200d🏫 *Öğretmen Masası*\nÖğretmen: *{name}* ({subject})',
              'menu_updated': '✅ Yemek menüsü güncellendi.',
              'no_active_homeworks': '📢 *{class_name}* sınıfı için aktif ödev bulunmamaktadır.',
              'no_blacklisted': '✅ Engellenmiş kullanıcı bulunmamaktadır.',
              'no_classes_found': '⚠️ Henüz kayıtlı bir sınıf bulunmamaktadır.',
              'no_grades': 'Henüz girilmiş bir ders notu bulunmamaktadır.',
              'no_hws_found': 'Henüz gönderilmiş ödeviniz bulunmamaktadır.',
              'no_linked_student': '⚠️ Hesabınıza tanımlı bir öğrenci bulunamadı.',
              'no_pending_appointments': '✅ Bekleyen randevu talebi bulunmamaktadır.',
              'no_pending_medical': '✅ Bekleyen mazeret raporu bulunmamaktadır.',
              'no_pending_requests': '✅ Bekleyen erişim başvurusu bulunmamaktadır.',
              'no_permission_grade': 'Bu notu düzenleme yetkiniz yok.',
              'no_permission_student_record': 'Bu öğrenci kaydına erişim yetkiniz yok.',
              'no_registered_students': 'Hesabınıza bağlı öğrenci bulunamadı.',
              'no_registered_teachers': '⚠️ Kayıtlı öğretmen bulunamadı.',
              'no_students_in_class': 'Bu sınıfta kayıtlı öğrenci yok.',
              'parent_choose_child': '🧑\u200d🎓 Lütfen aktif işlem yapmak istediğiniz öğrenciyi seçiniz:',
              'parent_info_title': 'ℹ️ *Okul Bilgi ve Hizmet Panosu*',
              'parent_settings_title': '⚙️ *Ayarlar & Hesap*',
              'parent_unlinked_success': '✅ Veli bağlantıları koparıldı. Yeni Veli Kodu: `{code}`',
              'pdf_ready': '📄 *{class_name}* şifre kartları ektedir.',
              'pdf_report_ready': '📄 *{name}* öğrencimizin resmi dönem not ve gelişim karnesi ektedir.',
              'pending_appointments_title': '🤝 *Bekleyen Veli Randevu Talepleri:*',
              'pending_medical_title': '🏥 *Bekleyen Mazeret Raporları:*',
              'pending_requests_title': '🛎️ *Bekleyen Erişim Başvuruları:*',
              'perm_admin_assigned_toast': '✅ Kullanıcı kalıcı yönetici yapıldı.',
              'permanent_admin_protected': '⛔ Kalıcı / Kurucu yöneticinin yetkisi kaldırılamaz!',
              'permanent_admin_title': 'Kalıcı / Kurucu Yönetici',
              'photo_expected_medical': '⚠️ Lütfen sadece fotoğraf gönderiniz.',
              'prompt_add_child_code': '🔑 Eklemek istediğiniz diğer öğrencinin Veli Kodunu yazınız (Örn: '
                                       '`VELI-123456`):',
              'prompt_appointment_note': '📝 Randevu için uygun olduğunuz gün/saati ve varsa notunuzu yazınız:',
              'prompt_broadcast': '📢 Göndermek istediğiniz duyuru metnini yazınız:',
              'prompt_enter_code_direct': '🔑 *Lütfen size verilen giriş kodunu yazınız:* (Örn: `HCA-123456`, '
                                          '`VELI-123456`, `OGR-123456`)',
              'prompt_grade_badge': 'Değerlendirme rozeti seçiniz:',
              'prompt_grade_score': 'Öğrenci: *{name}* ({class_name})\n'
                                    'Sınav Türü: *{exam_type}*\n'
                                    '\n'
                                    'Notu giriniz (0-100):',
              'prompt_hw_class': 'Ödevin sınıfını seçiniz:',
              'prompt_hw_content': 'Ödev açıklamasını yazınız veya fotoğraf gönderiniz:',
              'prompt_menu_update': '🍲 Bugünün yemek menüsünü yazıp gönderiniz:',
              'prompt_new_score': 'Yeni notu giriniz (0-100):',
              'prompt_search_student': '🔍 Öğrenci adı veya numarası yazınız:',
              'prompt_select_exam_type': '📝 Öğrenci: *{name}* ({class_name})\n'
                                         '\n'
                                         'Lütfen sınav / değerlendirme türünü seçiniz:',
              'prompt_student_class': '🏫 Öğrencinin Sınıfını yazınız (Örn: `9-A`):',
              'prompt_student_name': '👤 Öğrencinin Adını ve Soyadını yazınız:',
              'prompt_student_no': '🔢 Öğrencinin Okul Numarasını yazınız (Örn: `101`):',
              'prompt_teacher_name': '👨\u200d🏫 Öğretmenin Adını ve Soyadını yazınız:',
              'prompt_teacher_subject': '📚 Öğretmenin Branşını yazınız (Örn: `Matematik`):',
              'published_homeworks_title': '📢 *Yayınladığınız Ödevler:*',
              'recent_grades_title': '📝 *Girdiğiniz Son Notlar:*',
              'remind_att_sent': 'Hatırlatma gönderildi.',
              'report_not_found': 'Mazeret raporu bulunamadı.',
              'req_already_pending': '⚠️ Zaten onay bekleyen bir başvurunuz bulunmaktadır.',
              'req_approved_admin_msg': '✅ #{id} numaralı başvuru onaylandı. Kullanıcı: *{name}* ({role})',
              'req_approved_user': '🎉 *Tebrikler!*\n'
                                   'Okul idaresi başvurunuzu onayladı. Sisteme *{role}* olarak giriş yaptınız.',
              'req_details_parent': '🧑\u200d🎓 Lütfen öğrencinizin adını, sınıfını veya okul numarasını yazınız:',
              'req_details_student': '🏫 Lütfen sınıfınızı ve okul numaranızı yazınız:',
              'req_details_teacher': '📚 Lütfen branşınızı ve okul idaresine iletmek istediğiniz notu yazınız:',
              'req_name_prompt': '👤 Lütfen Adınızı ve Soyadınızı yazınız:',
              'req_phone_prompt': '📱 Lütfen telefon numaranızı yazınız (Örn: `+905...` veya `+998...`):',
              'req_rejected_admin_msg': '❌ #{id} numaralı başvuru reddedildi: *{name}*',
              'req_rejected_user': '❌ Okul idaresi erişim başvurunuzu reddetti. Lütfen idare ile iletişime geçiniz.',
              'req_role_select': '🛎️ *Erişim Talebi Başvurusu*\n\nLütfen talep ettiğiniz rolü seçiniz:',
              'req_sent_success': '✅ Başvurunuz okul idaresine iletildi. Onaylandığında bilgilendirileceksiniz.',
              'request_already_handled': '⚠️ Bu başvuru zaten işleme alınmış.',
              'rk_appointments': '🤝 Veli Randevuları',
              'rk_attendance': '📋 Hızlı Yoklama',
              'rk_cancel_action': '❌ İşlemi İptal Et',
              'rk_cat_reports': '📊 Raporlar & Denetim',
              'rk_cat_requests': '🛎️ Onay Masası',
              'rk_cat_settings': '⚙️ Sistem & Ayarlar',
              'rk_cat_staff': '👥 Kadro & Öğrenci',
              'rk_cat_tools': '🛠️ İdari Araçlar',
              'rk_grade': '📝 Not Girişi',
              'rk_homework': '📢 Ödev Panosu',
              'rk_logout': '🚪 Çıkış Yap',
              'rk_parent_info': 'ℹ️ Okul Bilgi Panosu',
              'rk_report': '📊 Durum Paneli (Karne)',
              'rk_restart': '🔄 Yeniden Başlat',
              'rk_switch_student': '🧑\u200d🎓 Öğrenci Değiştir',
              'rk_upload_medical': '🏥 Mazeret / Rapor',
              'role_parent_btn': '👨\u200d👩\u200d👧\u200d👦 Veli',
              'role_student_btn': '🎓 Öğrenci',
              'role_teacher_btn': '👨\u200d🏫 Öğretmen',
              'schedule_select_class': '📅 Programını görmek istediğiniz sınıfı seçiniz:',
              'school_admin_title': 'Okul İdaresi',
              'search_no_results': '❌ Eşleşen öğrenci bulunamadı.',
              'search_results_title': '🔍 *Arama Sonuçları:*',
              'search_user_no_results': '❌ Eşleşen kullanıcı bulunamadı.',
              'search_user_prompt': "🔍 Aramak istediğiniz kullanıcının adını, kullanıcı adını veya Telegram ID'sini "
                                    'yazınız:',
              'search_user_results_title': '🔍 *Kullanıcı Arama Sonuçları:*',
              'select_pdf_class': '📄 Şifre kartlarını indirmek istediğiniz sınıfı seçiniz:',
              'select_teacher_appointment': '🤝 Görüşmek istediğiniz öğretmeni seçiniz:',
              'send_dm_prompt': '✉️ *Özel Mesaj Gönder:*\n'
                                '\n'
                                'Lütfen `{name}` (`{id}`) kullanıcısına iletmek istediğiniz mesajı yazınız:',
              'setting_updated_toast': 'Ayar güncellendi.',
              'student_added_card': '✅ *Öğrenci Başarıyla Kaydedildi!*\n'
                                    '\n'
                                    '👤 İsim: *{name}*\n'
                                    '🏫 Sınıf: *{class_name}* | No: *{no}*\n'
                                    '\n'
                                    '🔑 *Giriş Şifreleri:*\n'
                                    '• Öğrenci Kodu: `{st_code}`\n'
                                    '• Veli Kodu: `{pr_code}`',
              'student_card': '👤 *Öğrenci Kartı*\n'
                              'İsim: *{name}*\n'
                              'Sınıf: *{class_name}* | No: *{no}*\n'
                              '\n'
                              '🔑 *Kod Durumu:*\n'
                              '• Öğrenci: `{st_code}` ({st_status})\n'
                              '• Veli: `{pr_code}` ({pr_status})',
              'student_deleted': '🗑️ Öğrenci sistemden silindi.',
              'student_info_updated': '✅ Öğrenci bilgileri güncellendi:\n*{name}* ({class_name} - No: {no})',
              'student_name_invalid': '❌ Lütfen geçerli bir isim yazınız.',
              'student_not_found': 'Öğrenci bulunamadı.',
              'student_switched_success': 'Aktif öğrenci seçildi: *{name}* ({class_name})',
              'teacher_added_card': '✅ *Öğretmen Başarıyla Kaydedildi!*\n'
                                    '\n'
                                    '👤 İsim: *{name}*\n'
                                    '📚 Branş: *{subject}*\n'
                                    '\n'
                                    '🔑 *Öğretmen Giriş Kodu:*\n'
                                    '`{code}`',
              'teacher_card': '👨\u200d🏫 *Öğretmen Kartı*\n'
                              'İsim: *{name}*\n'
                              'Branş: *{subject}*\n'
                              '\n'
                              '🔑 Kod: `{code}`\n'
                              'Durum: {status}',
              'teacher_deleted': '🗑️ Öğretmen sistemden silindi.',
              'teacher_name_invalid': '❌ Lütfen geçerli bir öğretmen ismi yazınız.',
              'teacher_not_found': 'Öğretmen bulunamadı.',
              'teacher_search_no_results': '❌ Eşleşen öğretmen bulunamadı.',
              'teacher_search_prompt': '🔍 Öğretmen adı veya branş yazınız:',
              'teacher_search_results_title': '🔍 *Öğretmen Arama Sonuçları:*',
              'temp_admin_assigned_toast': '✅ Kullanıcı {dur} süreyle geçici yönetici yapıldı.',
              'temp_admin_choose_title': '⏱️ *Geçici Yöneticilik Süresi Seçiniz:*',
              'uc_card_title': '👤 *KULLANICI PROFİL VE İŞLEM KARTI*',
              'unauthorized_action': '⛔ Bu işlem için yetkiniz bulunmamaktadır.',
              'unauthorized_excel_upload': '⛔ Excel yükleme yetkiniz yok.',
              'unban_success': 'Kullanıcının engeli kaldırıldı.',
              'upload_med_prompt': 'Lütfen mazeret veya rapor belgesinin fotoğrafını gönderiniz:',
              'user_banned_toast': 'Kullanıcı engellendi.',
              'user_not_found_toast': 'Kullanıcı bulunamadı.',
              'user_unbanned_toast': 'Kullanıcının engeli kaldırıldı.',
              'welcome_guest': '🎓 *Okul Yönetim Sistemine Hoş Geldiniz.*\n'
                               '\n'
                               'Lütfen sisteme giriş yapmak için size verilen **erişim kodunu** (Örn: `VELI-123456`, '
                               '`HCA-123456`, `OGR-123456`) yazınız veya işlem seçiniz:'},
    'uz': {   'acknowledged_toast': 'Tasdiqlaganingiz saqlandi.',
              'action_cancelled': '❌ *Amal bekor qilindi.*',
              'admin_add_name_prompt': "👤 Yangi ma'murning Familiyasi va Ismini kiriting:",
              'admin_add_tg_id_prompt': "➕ Ma'mur qilmoqchi bo'lgan shaxsning Telegram ID raqamini kiriting:",
              'admin_added_success': "✅ *{name}* (`{id}`) muvaffaqiyatli doimiy ma'mur sifatida qo'shildi.",
              'admin_admins_hub_title': "👨\u200d💼 *Maktab ma'muriyati va idoraviy vakolatlar:*\n"
                                        '\n'
                                        "Batafsil ma'lumot ko'rish yoki vakolatlarni boshqarish uchun ma'murni "
                                        'tanlang:',
              'admin_code_generated': "🔑 *BIR MARTALIK MA'MUR KODI YARATILDI*\n"
                                      '\n'
                                      'Kod: `{code}`\n'
                                      '\n'
                                      "Ushbu kodni ma'mur qilmoqchi bo'lgan shaxsga yuboring. Foydalanuvchi botga bu "
                                      "kodni yozishi bilan hisobi *Maktab ma'muriyati (Admin)* darajasiga ko'tariladi.",
              'admin_demoted_notification': "ℹ️ Ma'murlik vakolatlaringiz bekor qilindi.",
              'admin_demoted_toast': "Ma'mur vakolatlari bekor qilindi.",
              'admin_invalid_tg_id': "❌ Noto'g'ri Telegram ID! Faqat raqamlardan iborat bo'lishi kerak.",
              'admin_promoted_notification': '🎉 *Hurmatli {name},*\n'
                                             "Maktab boshqaruv tizimida *Doimiy ma'mur* etib tayinlandingiz!",
              'admin_restart_confirmed': '🔄 *Boshqaruv paneli qayta ishga tushirildi.*',
              'admin_sched_edit_title': '📅 *Dars jadvalini tahrirlash*\nSinfni tanlang:',
              'admin_sched_updated': '✅ *{class_name}* sinfi dars jadvali yangilandi.',
              'admin_stats': '📊 *Umumiy holat:*\n'
                             "• Sinflar: *{c_cnt}* | O'quvchilar: *{s_cnt}* | O'qituvchilar: *{t_cnt}*\n"
                             "• Arizalar: *{req_cnt}* | Ma'lumotnomalar: *{med_cnt}*\n"
                             '• Sana: *{date}*',
              'admin_title': '⚡ *Maktab boshqaruv markazi (Admin)*',
              'admin_unban_notification': "🟢 Hisobingizdagi blok maktab ma'muriyati tomonidan olib tashlandi.",
              'admin_user_card_title': "👤 *MA'MUR FOYDALANUVCHI KARTASI*",
              'all_notifs_acknowledged': '✅ *Barcha davomat xabarlari ota-onalar tomonidan tasdiqlandi!*\n'
                                         '\n'
                                         "So'nggi 36 soatda o'qilmagan xabarlar mavjud emas.",
              'appointment_approved_msg': "✅ O'qituvchi uchrashuv so'rovingizni qabul qildi.",
              'appointment_confirmed_toast': 'Uchrashuv tasdiqlandi.',
              'appointment_not_found': '⚠️ Uchrashuv topilmadi.',
              'appointment_rejected_msg': "❌ O'qituvchi bu vaqtda bo'sh emasligini bildirdi.",
              'appointment_sent': "✅ Uchrashuv so'rovingiz o'qituvchiga yuborildi.",
              'att_check_all_done': "✅ Barcha sinflar davomati to'liq olingan.",
              'att_check_title': '📊 *Kunlik davomat nazorati ({date})*',
              'att_saved': '✅ Davomat saqlandi (15 daqiqa tuzatish vaqti).',
              'attendance_correction_notification': 'ℹ️ *TUZATISH:* Farzandingiz *{name}* haqidagi davomat yozuvi '
                                                    "to'g'rilandi (KELDI).",
              'attendance_hours_lock': '⚠️ Davomat faqat 07:00 dan 19:00 gacha olinadi.',
              'attendance_intro': '📋 *{class_name} davomati*\nKelmaganlarni belgilab saqlang:',
              'attendance_select_class': "📋 Davomat olmoqchi bo'lgan sinfni tanlang:",
              'attendance_weekend_lock': '⚠️ Dam olish kunlari davomat olinmaydi.',
              'auth_blacklisted': "🚫 Hisobingiz butunlay to'xtatildi.",
              'auth_code_already_linked': '⚠️ *Bu kod allaqachon boshqa Telegram hisobiga ulangan.*\n'
                                          '\n'
                                          "Kodingiz o'g'irlangan deb hisoblasangiz, darhol maktab ma'muriyatiga "
                                          'murojaat qiling.',
              'auth_failed': "❌ Noto'g'ri kod! Qolgan urinishlar soni: {remaining}",
              'auth_locked': '⛔ Xavfsizlik sababli hisobingiz 1 soatga bloklandi.',
              'auth_success': '✅ *Kirish muvaffaqiyatli!*\nXush kelibsiz: *{name}*\nSizning rolingiz: *{role}*',
              'badge_missing': "🟡 Kamchilik / O'rganish",
              'badge_praise': '🟢 Maqtov / Muvaffaqiyat',
              'badge_warning': '🔴 Intizom / Ogohlantirish',
              'blacklisted_title': '🚫 *Bloklangan va cheklangan foydalanuvchilar:*',
              'broadcast_success': "📢 E'lon *{count}* ta foydalanuvchiga yetkazildi.",
              'btn_academic_report': '📈 Sinf reytingi',
              'btn_acknowledged': "✅ O'qidim / Xabardorman",
              'btn_add_admin_id': "➕ Telegram ID orqali ma'mur qo'shish",
              'btn_add_another': "➕ Yana qo'shish",
              'btn_add_new_teacher': "➕ Yangi o'qituvchi qo'shish",
              'btn_add_student': "➕ O'quvchi qo'shish",
              'btn_add_teacher': "➕ O'qituvchi qo'shish",
              'btn_appr_appointment': '✅ Qabul qilish',
              'btn_appr_medical': '✅ Tasdiqlash (Ruxsatli)',
              'btn_appr_request': '✅ Tasdiqlash',
              'btn_attendance': '📋 Tezkor davomat',
              'btn_audit_logs': '📜 Tizim jurnali',
              'btn_back': '⬅️ Orqaga',
              'btn_ban_user': '🚫 Bloklash (Ban)',
              'btn_blacklist': '🚫 Bloklanganlar',
              'btn_briefing_off': "🔕 Kunlik hisobot (O'CHIRILGAN)",
              'btn_briefing_on': '🔔 Kunlik hisobot (YOQILGAN)',
              'btn_broadcast': "📢 Ommaviy e'lon",
              'btn_cafeteria_edit': '🍲 Oshxona menyusini yangilash',
              'btn_cancel_action': '⬅️ Bekor qilish',
              'btn_class_att_sheet': 'Davomat qaydnomasi',
              'btn_class_grade_sheet': 'Baholar qaydnomasi',
              'btn_class_pdf_cards': 'Parol kartalari (PDF)',
              'btn_class_sched': 'Jadvali',
              'btn_classes': "🏫 Sinflar va o'quvchilar",
              'btn_cockpit': '📊 Tonggi hisobot',
              'btn_cockpit_unified': '📊 Tonggi hisobot va davomat',
              'btn_del_grade': "❌ Bahoni o'chirish",
              'btn_del_student': "❌ O'quvchini o'chirish",
              'btn_del_teacher': "❌ O'qituvchini o'chirish",
              'btn_delete_action': "🗑️ O'chirish",
              'btn_download_pdf_report': '📄 Rasmiy PDF kundalikni yuklab olish',
              'btn_dur_24h': '⏱️ 24 soat',
              'btn_dur_30d': '⏱️ 30 kun',
              'btn_dur_7d': '⏱️ 7 kun',
              'btn_edit_class': '🏫 Sinfni tahrirlash',
              'btn_edit_grade': '✏️ Bahoni tahrirlash',
              'btn_edit_name': '👤 Ismni tahrirlash',
              'btn_edit_no': '🔢 Raqamni tahrirlash',
              'btn_edit_student': '✏️ Tahrirlash',
              'btn_enter_grade': "📝 Baho qo'yish",
              'btn_excel': '📥 Excel orqali yuklash',
              'btn_excel_hub': '📥 Excel markazi',
              'btn_export_all_data': "📊 Barcha maktab ma'lumotlarini yuklash (.xlsx)",
              'btn_gen_admin_code': "🔑 Bir martalik ma'mur kodini yaratish",
              'btn_homework_board': '📢 Vazifalar paneli',
              'btn_lang': "🌐 Tilni o'zgartirish",
              'btn_login_prompt': '🔑 Tizimga kirish (Kodni kiritish)',
              'btn_main_menu': '🏠 Asosiy menyu',
              'btn_maintenance_toggle': "🚨 Ta'mir rejimi ({status})",
              'btn_make_perm_admin': "👑 Doimiy ma'mur qilish",
              'btn_make_temp_admin': "⏱️ Vaqtinchalik ma'mur qilish",
              'btn_manage_schedule': '📅 Dars jadvali boshqaruvi',
              'btn_medical': "🏥 Ma'lumotnomalar ({count})",
              'btn_my_hws': '📚 Yuborgan vazifalarim',
              'btn_next': 'Keyingi ➡️',
              'btn_not_available': '❌ Bandman',
              'btn_notices': "📢 Maktab e'lonlari",
              'btn_pdf': '📄 Parol kartalari (PDF)',
              'btn_prev': '⬅️ Oldingi',
              'btn_recent_grades_menu': "🕒 So'nggi baholar va tahrir",
              'btn_refresh_data': "🔄 Ma'lumotlarni yangilash",
              'btn_reject': '❌ Rad etish',
              'btn_remind_att': "⚠️ O'qituvchilarga davomatni eslatish",
              'btn_report_card': '📊 Holat paneli (Kundalik)',
              'btn_req_access': "📩 Ruxsat so'rash",
              'btn_req_chat': "📞 1:1 Aloqa so'rovi yuborish",
              'btn_requests': '🛎️ Arizalar ({count})',
              'btn_reset_codes': '🔄 Kodlarni yangilash',
              'btn_revoke_admin_perm': "❌ Ma'mur vakolatini olish",
              'btn_risk_radar': "⚠️ Xavf ostidagi o'quvchilar",
              'btn_save_att': '💾 Saqlash',
              'btn_school_admins': "👨\u200d💼 Maktab ma'muriyati",
              'btn_search_again': '🔍 Qayta qidirish',
              'btn_search_student': "🔍 O'quvchini qidirish",
              'btn_search_teacher': "🔍 O'qituvchini qidirish",
              'btn_search_user': '🔍 Foydalanuvchi qidirish',
              'btn_send_dm': '✉️ Shaxsiy xabar yuborish',
              'btn_send_new_hw': 'Yangi vazifa yuborish',
              'btn_switch_student': "🧑\u200d🎓 O'quvchini tanlash",
              'btn_teachers': "👨\u200d🏫 O'qituvchilar",
              'btn_teachers_pdf': "👨\u200d🏫 O'qituvchilar parol kartalari (PDF)",
              'btn_unack_notifs': '⚠️ Tasdiqlanmagan davomatlar',
              'btn_unban_user': '🟢 Blokdan chiqarish (Unban)',
              'btn_unlink_parent': '👨\u200d👩\u200d👧\u200d👦 Ota-onani ajratish',
              'btn_upload_excel': "📥 Excel orqali o'quvchi yuklash",
              'btn_upload_medical': "🏥 Ma'lumotnoma yuborish",
              'btn_users_hub': '👥 Foydalanuvchilar',
              'btn_users_list': "⬅️ Ma'murlar ro'yxati",
              'btn_view_cafeteria': '🍲 Kunlik oshxona menyusi',
              'btn_view_photo': 'Rasmni ochish',
              'btn_view_schedule': '📅 Haftalik dars jadvali',
              'btn_write_to_admin': "💬 Ma'murga xabar yozish",
              'cat_reports_title': "📊 *Akademik hisobotlar va davomat nazorati*\nBo'limni tanlang:",
              'cat_requests_title': "🛎️ *Arizalar va ma'lumotnomalarni tasdiqlash*\nAmalni tanlang:",
              'cat_settings_title': '⚙️ *Tizim va xavfsizlik sozlamalari*\nAmalni tanlang:',
              'cat_staff_title': "👥 *O'quvchilar va o'qituvchilar boshqaruvi*\nAmalni tanlang:",
              'cat_tools_title': "🛠️ *Boshqaruv vositalari va e'lonlar*\nBo'limni tanlang:",
              'chat_req_error_toast': '⚠️ Foydalanuvchi botni bloklagani sababli yetkazilmadi!',
              'chat_req_sent_toast': "✅ 1:1 Aloqa so'rovi yuborildi!",
              'child_added_success': "✅ *{name}* ({class_name}) hisobingizga muvaffaqiyatli qo'shildi!",
              'cockpit_report': '📊 *Tonggi hisobot ({date})*\n'
                                '\n'
                                '🏫 Jami: *{total}* | ✅ Kelgan: *{present}* | ❌ Kelmagan: *{absent}*\n'
                                '\n'
                                '⚠️ *Davomat olinmagan ({missing_cnt}):*\n'
                                '{missing}',
              'codes_reset_done': "✅ Kodlar yangilandi!\n\n• Yangi O'quvchi: `{st_code}`\n• Yangi Ota-ona: `{pr_code}`",
              'contact_req_direct': 'Iltimos, quyidagi tugma orqali bevosita suhbatni boshlang:',
              'contact_req_header': "📞 *MAKTAB MA'MURIYATI ALOQA SO'ROVI*\n"
                                    '\n'
                                    "Maktab ma'muriyati siz bilan 1:1 bog'lanishni xohlamoqda.\n"
                                    "👤 *Ma'mur:* {name}\n",
              'contact_req_id': "Iltimos, maktab ma'muriyatiga yozing.",
              'dm_delivery_error': '⚠️ Yetkazish xatosi: Foydalanuvchi botni bloklagan.',
              'dm_from_admin_header': "📩 *MAKTAB MA'MURIYATIDAN XABAR*",
              'dm_sender_label': 'Yuboruvchi',
              'dm_sent_success': '✅ Xabar muvaffaqiyatli yetkazildi!',
              'evening_briefing_header': '🌙 *KUNLIK YAKUNIY HISOBOT (18:30)*\n'
                                         "O'quvchi: *{name}* ({class_name})\n"
                                         '\n'
                                         '📌 Davomat: *{att_status}*\n'
                                         '📝 Baholar:\n'
                                         '{grades}',
              'exam_oral': "🗣️ Og'zaki / Faollik",
              'exam_written_1': '📝 1-Yozma ish',
              'exam_written_2': '📝 2-Yozma ish',
              'excel_done': "✅ Qabul qilindi! O'quvchilar: *{count}*\nParollar biriktirilgan faylda berildi.",
              'excel_format_error': "❌ Excel faylini o'qishda xatolik yuz berdi.",
              'excel_hub_title': '📥 *Excel boshqaruv markazi*\nAmalni tanlang:',
              'excel_info': '📥 *Excel orqali yuklash*\n'
                            '\n'
                            '`.xlsx` faylini yuboring.\n'
                            'Ustunlar: `Ad Soyad` | `Sinif` | `Numara`',
              'export_ready': "📥 *Maktab ma'lumotlar zaxirasi tayyor ({date})*",
              'grade_deleted': "Baho o'chirildi.",
              'grade_parent_notification': '📝 *YANGI BAHO BILDIRISHNOMASI*\n'
                                           '\n'
                                           "🧑\u200d🎓 O'quvchi: *{name}*\n"
                                           '📚 Fan: *{subject}* ({exam_type})\n'
                                           '📊 Baho: *{score}* ({badge})',
              'grade_saved_success': '✅ Baho saqlandi va ota-onaga yuborildi.',
              'grade_select_class': "📝 Baho qo'ymoqchi bo'lgan sinfni tanlang:",
              'grade_select_student': "📝 *{class_name} sinfi*\nBaho qo'ymoqchi bo'lgan o'quvchini tanlang:",
              'grade_updated': '✅ Baho yangilandi.',
              'homework_board_title': '📢 *{class_name} sinfi vazifalar paneli:*',
              'homework_deleted_toast': "Vazifa o'chirildi.",
              'hw_sent_success': '📢 Vazifa *{class_name}* sinfiga yuborildi.',
              'image_load_error': "Rasmni ochib bo'lmadi.",
              'invalid_name_error': "❌ Iltimos, to'g'ri ism-familiya kiriting.",
              'invalid_parent_code': "❌ Noto'g'ri ota-ona kodi!",
              'invalid_phone_error': "❌ Noto'g'ri telefon raqami! Kamida 7 ta raqam kiriting.",
              'invalid_score_format': "❌ Noto'g'ri baho! Son kiriting (masalan: 85).",
              'invalid_score_range': "❌ Baho 0 dan 100 gacha bo'lishi kerak!",
              'lang_changed': "Til muvaffaqiyatli yangilandi: 🇺🇿 O'zbekcha",
              'lang_select': '🌍 Iltimos, tilni tanlang / Пожалуйста, выберите язык / Lütfen dil seçiniz / Please '
                             'select language:',
              'lbl_account_status': 'Hisob holati',
              'lbl_admin_status': "Ma'murlik holati",
              'lbl_class': 'Sinf',
              'lbl_full_name': 'F.I.SH',
              'lbl_lang': 'Til',
              'lbl_linked_students': "Biriktirilgan o'quvchilar",
              'lbl_not_admin': "Oddiy foydalanuvchi (Ma'mur emas)",
              'lbl_number': 'Raqam',
              'lbl_phone': 'Telefon',
              'lbl_role': 'Rol',
              'lbl_role_admin': "Maktab ma'muri (Admin)",
              'lbl_role_guest': 'Mehmon',
              'lbl_role_parent': 'Ota-ona',
              'lbl_role_student': "O'quvchi",
              'lbl_role_teacher': "O'qituvchi",
              'lbl_status_active': '🟢 *Faol*',
              'lbl_status_banned': '🚫 *Bloklangan (Ban)*',
              'lbl_subject': 'Fan',
              'lbl_username': 'Foydalanuvchi nomi',
              'lock_countdown_msg': '⛔ *Xavfsizlik bloki:* Hisobingiz vaqtincha bloklangan.\n'
                                    '\n'
                                    'Qolgan vaqt: *{mins} daqiqa*.',
              'maintenance_mode': "⚠️ Tizimda rejaviy texnik ishlar olib borilmoqda. Iltimos, keyinroq urinib ko'ring.",
              'maintenance_mode_updated': "Ta'mir rejimi yangilandi.",
              'med_uploaded_success': "Ma'lumotnoma ma'muriyatga yuborildi.",
              'medical_approved': "✅ Ma'lumotnoma tasdiqlandi.",
              'medical_approved_parent': "✅ Farzandingizning ma'lumotnomasi ma'muriyat tomonidan tasdiqlandi.",
              'medical_rejected': "❌ Ma'lumotnoma rad etildi.",
              'medical_rejected_parent': "❌ Farzandingizning ma'lumotnomasi ma'muriyat tomonidan rad etildi.",
              'menu_parent': "👨\u200d👩\u200d👧\u200d👦 *Ota-ona paneli*\nO'quvchi: *{name}* ({class_name})",
              'menu_student': "🎓 *O'quvchi paneli*\nO'quvchi: *{name}* ({class_name} - №: {no})",
              'menu_teacher': "👨\u200d🏫 *O'qituvchi paneli*\nO'qituvchi: *{name}* ({subject})",
              'menu_updated': '✅ Taomlar menyusi yangilandi.',
              'no_active_homeworks': "📢 *{class_name}* sinfi uchun faol vazifalar yo'q.",
              'no_blacklisted': "✅ Bloklangan foydalanuvchilar yo'q.",
              'no_classes_found': '⚠️ Hozircha birorta sinf mavjud emas.',
              'no_grades': 'Hozircha baholar mavjud emas.',
              'no_hws_found': "Hozircha yuborilgan vazifalar yo'q.",
              'no_linked_student': "⚠️ Sizga biriktirilgan o'quvchi topilmadi.",
              'no_pending_appointments': "✅ Kutilayotgan uchrashuv so'rovlari yo'q.",
              'no_pending_medical': "✅ Kutilayotgan ma'lumotnomalar mavjud emas.",
              'no_pending_requests': '✅ Kutilayotgan arizalar mavjud emas.',
              'no_permission_grade': "Bu bahoni tahrirlash vakolatingiz yo'q.",
              'no_permission_student_record': "Ushbu o'quvchi ma'lumotiga kirish huquqingiz yo'q.",
              'no_registered_students': "Sizga biriktirilgan o'quvchi topilmadi.",
              'no_registered_teachers': "⚠️ O'qituvchilar topilmadi.",
              'no_students_in_class': "Bu sinfda o'quvchilar yo'q.",
              'parent_choose_child': "🧑\u200d🎓 Iltimos, o'quvchini tanlang:",
              'parent_info_title': "ℹ️ *Maktab ma'lumotlari va xizmatlar paneli*",
              'parent_settings_title': '⚙️ *Sozlamalar va hisob*',
              'parent_unlinked_success': '✅ Ota-ona ajratildi. Yangi ota-ona kodi: `{code}`',
              'pdf_ready': '📄 *{class_name}* kartalari tayyorlandi.',
              'pdf_report_ready': "📄 *{name}* o'quvchimizning rasmiy baholar hisoboti biriktirildi.",
              'pending_appointments_title': "🤝 *Ota-onalarning uchrashuv so'rovlari:*",
              'pending_medical_title': "🏥 *Ko'rib chiqilishi kerak bo'lgan ma'lumotnomalar:*",
              'pending_requests_title': "🛎️ *Ko'rib chiqilishi kerak bo'lgan arizalar:*",
              'perm_admin_assigned_toast': "✅ Foydalanuvchiga doimiy ma'mur vakolati berildi.",
              'permanent_admin_protected': "⛔ Asosiy/doimiy ma'mur vakolatlarini bekor qilib bo'lmaydi!",
              'permanent_admin_title': "Doimiy / Asosiy ma'mur",
              'photo_expected_medical': '⚠️ Iltimos, faqat rasm yuboring.',
              'prompt_add_child_code': "🔑 Qo'shmoqchi bo'lgan boshqa o'quvchining Ota-ona kodini kiriting (Masalan: "
                                       '`VELI-123456`):',
              'prompt_appointment_note': '📝 Uchrashuv uchun qulay vaqt va izohingizni yozing:',
              'prompt_broadcast': "📢 E'lon matnini yozing:",
              'prompt_enter_code_direct': '🔑 *Iltimos, sizga berilgan kirish kodini yozing:* (Masalan: `HCA-123456`, '
                                          '`VELI-123456`, `OGR-123456`)',
              'prompt_grade_badge': 'Toifani tanlang:',
              'prompt_grade_score': "O'quvchi: *{name}* ({class_name})\n"
                                    'Baholash turi: *{exam_type}*\n'
                                    '\n'
                                    'Bahoni kiriting (0-100):',
              'prompt_hw_class': 'Vazifa sinfini tanlang:',
              'prompt_hw_content': 'Vazifa matnini yozing yoki doska rasmini yuboring:',
              'prompt_menu_update': '🍲 Bugungi oshxona menyusini yozib yuboring:',
              'prompt_new_score': 'Yangi bahoni kiriting (0-100):',
              'prompt_search_student': '🔍 Ism yoki raqam kiriting:',
              'prompt_select_exam_type': "📝 O'quvchi: *{name}* ({class_name})\n\nIltimos, baholash turini tanlang:",
              'prompt_student_class': "🏫 O'quvchining sinfini kiriting (Masalan: `9-A`):",
              'prompt_student_name': "👤 O'quvchining Familiyasi va Ismini kiriting:",
              'prompt_student_no': "🔢 O'quvchining raqamini kiriting (Masalan: `101`):",
              'prompt_teacher_name': "👨\u200d🏫 O'qituvchining Familiyasi va Ismini kiriting:",
              'prompt_teacher_subject': '📚 Fanni kiriting (Masalan: `Matematika`):',
              'published_homeworks_title': "📢 *Siz e'lon qilgan vazifalar:*",
              'recent_grades_title': "📝 *Kiritilgan so'nggi baholar:*",
              'remind_att_sent': 'Eslatma yuborildi.',
              'report_not_found': "Ma'lumotnoma topilmadi.",
              'req_already_pending': "⚠️ Sizda allaqachon ko'rib chiqilayotgan ariza mavjud.",
              'req_approved_admin_msg': '✅ #{id}-sonli ariza tasdiqlandi. Foydalanuvchi: *{name}* ({role})',
              'req_approved_user': '🎉 *Tabriklaymiz!*\n'
                                   "Maktab ma'muriyati arizangizni tasdiqladi. Tizimga *{role}* sifatida kirdingiz.",
              'req_details_parent': '🧑\u200d🎓 Iltimos, farzandingizning ismi, sinfi yoki raqamini kiriting:',
              'req_details_student': '🏫 Iltimos, sinfingiz va maktab raqamingizni kiriting:',
              'req_details_teacher': "📚 Iltimos, faningizni va ma'muriyatga xabaringizni kiriting:",
              'req_name_prompt': '👤 Iltimos, Familiyangiz va Ismingizni kiriting:',
              'req_phone_prompt': '📱 Iltimos, telefon raqamingizni kiriting (Masalan: `+998...`):',
              'req_rejected_admin_msg': '❌ #{id}-sonli ariza rad etildi: *{name}*',
              'req_rejected_user': "❌ Maktab ma'muriyati arizangizni rad etdi. Iltimos, ma'muriyatga murojaat qiling.",
              'req_role_select': "🛎️ *Kirish so'rovi arizasi*\n\nIltimos, so'ralayotgan rolni tanlang:",
              'req_sent_success': "✅ Arizangiz maktab ma'muriyatiga yuborildi. Tasdiqlangach xabar beriladi.",
              'request_already_handled': "⚠️ Bu ariza allaqachon ko'rib chiqilgan.",
              'rk_appointments': '🤝 Ota-onalar uchrashuvi',
              'rk_attendance': '📋 Tezkor davomat',
              'rk_cancel_action': '❌ Bekor qilish',
              'rk_cat_reports': '📊 Hisobotlar va nazorat',
              'rk_cat_requests': '🛎️ Tasdiqlash markazi',
              'rk_cat_settings': '⚙️ Tizim va sozlamalar',
              'rk_cat_staff': "👥 Kadro va o'quvchilar",
              'rk_cat_tools': '🛠️ Boshqaruv vositalari',
              'rk_grade': "📝 Baho qo'yish",
              'rk_homework': '📢 Vazifalar paneli',
              'rk_logout': '🚪 Chiqish',
              'rk_parent_info': "ℹ️ Maktab ma'lumotlari",
              'rk_report': '📊 Holat paneli (Kundalik)',
              'rk_restart': '🔄 Qayta ishga tushirish',
              'rk_switch_student': "🧑\u200d🎓 O'quvchini almashtirish",
              'rk_upload_medical': "🏥 Ma'lumotnoma",
              'role_parent_btn': '👨\u200d👩\u200d👧\u200d👦 Ota-ona',
              'role_student_btn': "🎓 O'quvchi",
              'role_teacher_btn': "👨\u200d🏫 O'qituvchi",
              'schedule_select_class': "📅 Jadvalini ko'rmoqchi bo'lgan sinfni tanlang:",
              'school_admin_title': "Maktab ma'muriyati",
              'search_no_results': "❌ O'quvchi topilmadi.",
              'search_results_title': '🔍 *Qidiruv natijalari:*',
              'search_user_no_results': '❌ Foydalanuvchi topilmadi.',
              'search_user_prompt': '🔍 Foydalanuvchining ismi, username yoki Telegram ID raqamini kiriting:',
              'search_user_results_title': '🔍 *Foydalanuvchilar qidiruv natijalari:*',
              'select_pdf_class': '📄 Parol kartalari uchun sinfni tanlang:',
              'select_teacher_appointment': "🤝 Uchrashmoqchi bo'lgan o'qituvchini tanlang:",
              'send_dm_prompt': '✉️ *Shaxsiy xabar yuborish:*\n'
                                '\n'
                                "Iltimos, `{name}` (`{id}`) foydalanuvchisiga yubormoqchi bo'lgan xabaringizni yozing:",
              'setting_updated_toast': 'Sozlama yangilandi.',
              'student_added_card': "✅ *O'quvchi saqlandi!*\n"
                                    '\n'
                                    '👤 Ism: *{name}*\n'
                                    '🏫 Sinf: *{class_name}* | №: *{no}*\n'
                                    '\n'
                                    '🔑 *Kodlar:*\n'
                                    "• O'quvchi: `{st_code}`\n"
                                    '• Ota-ona: `{pr_code}`',
              'student_card': "👤 *O'quvchi kartasi*\n"
                              'F.I.SH: *{name}*\n'
                              'Sinf: *{class_name}* | №: *{no}*\n'
                              '\n'
                              '🔑 *Kodlar:*\n'
                              "• O'quvchi: `{st_code}` ({st_status})\n"
                              '• Ota-ona: `{pr_code}` ({pr_status})',
              'student_deleted': "🗑️ O'quvchi tizimdan o'chirildi.",
              'student_info_updated': "✅ O'quvchi ma'lumotlari yangilandi:\n*{name}* ({class_name} - №: {no})",
              'student_name_invalid': "❌ Iltimos, to'g'ri ism kiriting.",
              'student_not_found': "O'quvchi topilmadi.",
              'student_switched_success': "Tanlangan o'quvchi: *{name}* ({class_name})",
              'teacher_added_card': "✅ *O'qituvchi saqlandi!*\n"
                                    '\n'
                                    '👤 Ism: *{name}*\n'
                                    '📚 Fan: *{subject}*\n'
                                    '\n'
                                    "🔑 *O'qituvchi kodi:*\n"
                                    '`{code}`',
              'teacher_card': "👨\u200d🏫 *O'qituvchi kartasi*\n"
                              'Ism: *{name}*\n'
                              'Fan: *{subject}*\n'
                              '\n'
                              '🔑 Kod: `{code}`\n'
                              'Holat: {status}',
              'teacher_deleted': "🗑️ O'qituvchi o'chirildi.",
              'teacher_name_invalid': "❌ Iltimos, to'g'ri o'qituvchi ismini kiriting.",
              'teacher_not_found': "O'qituvchi topilmadi.",
              'teacher_search_no_results': "❌ O'qituvchi topilmadi.",
              'teacher_search_prompt': "🔍 O'qituvchi ismi yoki fanini kiriting:",
              'teacher_search_results_title': "🔍 *O'qituvchilar qidiruv natijalari:*",
              'temp_admin_assigned_toast': "✅ Foydalanuvchi {dur} muddatga vaqtinchalik ma'mur qilindi.",
              'temp_admin_choose_title': "⏱️ *Vaqtinchalik ma'murlik muddatini tanlang:*",
              'uc_card_title': '👤 *FOYDALANUVCHI PROFILI VA AMALLAR KARTASI*',
              'unauthorized_action': "⛔ Bu amal uchun vakolatingiz yo'q.",
              'unauthorized_excel_upload': "⛔ Excel yuklash vakolatingiz yo'q.",
              'unban_success': 'Foydalanuvchi blokdan chiqarildi.',
              'upload_med_prompt': "Iltimos, ma'lumotnoma rasmini yuboring:",
              'user_banned_toast': 'Foydalanuvchi bloklandi.',
              'user_not_found_toast': 'Foydalanuvchi topilmadi.',
              'user_unbanned_toast': 'Foydalanuvchi blokdan chiqarildi.',
              'welcome_guest': '🎓 *Maktab boshqaruv tizimiga xush kelibsiz.*\n'
                               '\n'
                               'Iltimos, tizimga kirish uchun berilgan **kirish kodini** (Masalan: `VELI-123456`, '
                               '`HCA-123456`, `OGR-123456`) yozing yoki amalni tanlang:'}}

def get_text(key: str, lang: str = "tr", **kwargs) -> str:
    selected_lang = lang if lang in LOCALES else "tr"
    template = LOCALES.get(selected_lang, {}).get(key) or LOCALES["tr"].get(key) or f"[{key}]"
    if kwargs:
        try:
            return template.format(**kwargs)
        except Exception:
            return template
    return template

# ======================================================================
# 4. ARAYÜZ VE ETKİLEŞİM KLAVYELERİ (SADE & DAĞINIK OLMAYAN HUB DÜZENİ)
# ======================================================================

def get_role_reply_kb(role: str, lang: str = "tr") -> ReplyKeyboardMarkup:
    keyboard = []
    if role == "admin":
        keyboard = [
            [KeyboardButton(text=get_text("rk_cat_staff", lang)), KeyboardButton(text=get_text("rk_cat_reports", lang))],
            [KeyboardButton(text=get_text("rk_cat_requests", lang)), KeyboardButton(text=get_text("rk_cat_tools", lang))],
            [KeyboardButton(text=get_text("rk_cat_settings", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    elif role == "teacher":
        keyboard = [
            [KeyboardButton(text=get_text("rk_attendance", lang)), KeyboardButton(text=get_text("rk_grade", lang))],
            [KeyboardButton(text=get_text("rk_homework", lang)), KeyboardButton(text=get_text("rk_appointments", lang))],
            [KeyboardButton(text=get_text("btn_manage_schedule", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    elif role == "parent":
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_upload_medical", lang))],
            [KeyboardButton(text=get_text("rk_appointments", lang)), KeyboardButton(text=get_text("rk_parent_info", lang))],
            [KeyboardButton(text=get_text("rk_switch_student", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    elif role == "student":
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_homework", lang))],
            [KeyboardButton(text=get_text("rk_parent_info", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    else:
        keyboard = [
            [KeyboardButton(text=get_text("btn_login_prompt", lang)), KeyboardButton(text=get_text("btn_req_access", lang))],
            [KeyboardButton(text=get_text("btn_lang", lang)), KeyboardButton(text=get_text("rk_restart", lang))]
        ]
    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True, is_persistent=True)

def get_cancel_reply_kb(lang: str = "tr") -> ReplyKeyboardMarkup:
    return ReplyKeyboardMarkup(
        keyboard=[[KeyboardButton(text=get_text("rk_cancel_action", lang))]],
        resize_keyboard=True,
        is_persistent=True
    )

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
    return [
        InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=back_callback),
        InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")
    ]

def get_attendance_grid_kb(students: list, attendance_state: dict, class_name: str, lang: str = "tr") -> InlineKeyboardMarkup:
    inline_keyboard = []
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
        ("C:\Windows\Fonts\arial.ttf", "C:\Windows\Fonts\arialbd.ttf")
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
    except Exception:
        pass

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

    wb = openpyxl.Workbook()
    ws1 = wb.active
    ws1.title = "Ogrenciler"
    ws1.append(["ID", "Ad Soyad", "Sinif", "Okul Numarasi", "Ogrenci Kodu", "Veli Kodu", "Ogrenci Telegram ID"])
    for s in students:
        ws1.append([s.id, s.full_name, s.class_name, s.student_number, s.student_code, s.parent_code, s.student_telegram_id or "-"])

    ws2 = wb.create_sheet(title="Ogretmenler")
    ws2.append(["ID", "Ad Soyad", "Brans", "Giris Kodu", "Telegram ID"])
    for t in teachers:
        ws2.append([t.id, t.full_name, t.subject, t.auth_code, t.telegram_id or "-"])

    ws3 = wb.create_sheet(title="Yoklamalar")
    ws3.append(["Tarih", "Sinif", "Ogrenci ID", "Durum"])
    for a in attendances:
        ws3.append([str(a.date), a.class_name, a.student_id, a.status])

    ws4 = wb.create_sheet(title="Ders Notlari")
    ws4.append(["Tarih", "Ogrenci ID", "Ders", "Sinav Turu", "Not", "Rozet", "Aciklama"])
    for g in grades:
        ws4.append([g.created_at.strftime('%d.%m.%Y'), g.student_id, g.subject, g.exam_type or "1. Yazılı", g.score, g.badge, g.note or ""])

    out_buf = io.BytesIO()
    wb.save(out_buf)
    out_buf.seek(0)
    wb.close()
    return out_buf

async def process_student_excel(file_bytes: bytes) -> tuple[int, io.BytesIO]:
    in_buffer = io.BytesIO(file_bytes)
    wb_in = openpyxl.load_workbook(in_buffer, read_only=True, data_only=True)
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
        "tr": "<b>OKUL YÖNETİM SİSTEMİ - ÖĞRETMEN GİRİŞ KARTLARI</b>",
        "ru": "<b>СИСТЕМА УПРАВЛЕНИЯ ШКОЛОЙ - КАРТОЧКИ УЧИТЕЛЕЙ</b>",
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

    doc.build(story)
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

    doc.build(story)
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
        "tr": "<b>OKUL YÖNETİM SİSTEMİ - RESMİ ÖĞRENCİ GELİŞİM VE NOT KARNESİ</b>",
        "ru": "<b>СИСТЕМА УПРАВЛЕНИЯ ШКОЛОЙ - ОФИЦИАЛЬНЫЙ ТАБЕЛЬ УСПЕВАЕМОСТИ</b>",
        "uz": "<b>MAKTAB BOSHQARUV TIZIMI - RASMIY O'QUVCHI BAHOLAR VA RIVOJLANISH KUNDALIGI</b>",
        "en": "<b>SCHOOL MANAGEMENT SYSTEM - OFFICIAL STUDENT REPORT CARD</b>"
    }.get(lang, "<b>OFFICIAL STUDENT REPORT CARD</b>")

    story = [Paragraph(t_main, t_style), Spacer(1, 10)]
    date_str = get_local_date().strftime('%d.%m.%Y')
    avg_score = (sum(g.score for g in grades) / len(grades)) if grades else 100.0

    lbl_st = {"tr": "Öğrenci:", "ru": "Ученик:", "uz": "O'quvchi:", "en": "Student:"}.get(lang, "Student:")
    lbl_dt = {"tr": "Tarih:", "ru": "Дата:", "uz": "Sana:", "en": "Date:"}.get(lang, "Date:")
    lbl_cl = {"tr": "Sınıf:", "ru": "Класс:", "uz": "Sinf:", "en": "Class:"}.get(lang, "Class:")
    lbl_no = {"tr": "Numara:", "ru": "Номер:", "uz": "Raqam:", "en": "Roll No:"}.get(lang, "Roll No:")
    lbl_att = {"tr": "Devamsızlık:", "ru": "Пропуски:", "uz": "Davomat:", "en": "Absences:"}.get(lang, "Absences:")
    lbl_avg = {"tr": "Ortalama:", "ru": "Средний балл:", "uz": "O'rtacha:", "en": "Average:"}.get(lang, "Average:")

    att_text = f"{absent_cnt} (İzinli: {excused_cnt})" if lang == "tr" else (f"{absent_cnt} (Уваж: {excused_cnt})" if lang == "ru" else f"{absent_cnt} (Ruxsatli: {excused_cnt})")

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
    doc.build(story)
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
                    await safe_send_message(bot=bot, chat_id=p_id, text=msg_text, reply_markup=get_ack_notification_kb(notif.id, lang=p_lang), parse_mode="Markdown")

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
            await safe_send_message(bot, p.telegram_id, text, parse_mode="Markdown")
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
            await safe_send_message(bot, a_id, msg, parse_mode="Markdown")

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
            await bot.send_document(chat_id=a_id, document=doc_file, caption=caption_txt, parse_mode="Markdown")
            await asyncio.sleep(0.05)
    except Exception as e:
        pass

# ======================================================================
# 7. FSM DURUMLARI
# ======================================================================

class Form(StatesGroup):
    waiting_auth_code = State()
    waiting_admin_tg_id = State()
    waiting_admin_name = State()
    waiting_search_teacher_query = State()
    sched_update_text = State()
    waiting_search_user_query = State()
    waiting_admin_dm_text = State()
    req_role = State()
    req_name = State()
    req_phone = State()
    req_details = State()
    add_student_name = State()
    add_student_class = State()
    add_student_no = State()
    add_teacher_name = State()
    add_teacher_subject = State()
    waiting_search_query = State()
    waiting_broadcast_text = State()
    waiting_medical_photo = State()
    grade_exam_type = State()
    grade_score = State()
    grade_badge = State()
    hw_content = State()
    app_teacher = State()
    app_note = State()
    menu_update_text = State()
    edit_student_val = State()
    parent_add_child_code = State()
    edit_grade_val = State()

router = Router()
ATTENDANCE_CACHE = {}
GRADE_CACHE = {}
HW_CACHE = {}
REQ_CACHE = {}
APP_CACHE = {}

# ======================================================================
# 8. ROUTER: BAŞLANGIÇ, PROFİL VE KİMLİK DOĞRULAMA MOTORU
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
            if st: extra = f"{lbl_st}: *{escape_md(st.full_name)}* ({escape_md(st.class_name)})\n"
        elif user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            if st: extra = f"{lbl_cl}: *{escape_md(st.class_name)}* | {lbl_no}: *{escape_md(st.student_number)}*\n"
        elif user.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
            if tch: extra = f"{lbl_s}: *{escape_md(tch.subject)}*\n"

        p_title = get_text("uc_card_title", lang)
        lbl_name = get_text("lbl_full_name", lang)
        lbl_r = get_text("lbl_role", lang)
        lbl_user = get_text("lbl_role_guest", lang)

        text = f"{p_title}\n\n• {lbl_name}: *{escape_md(user.full_name or lbl_user)}*\n• {lbl_r}: *{role_label}*\n• Telegram ID: `{user.telegram_id}`\n{extra}"
        await safe_edit_or_answer(message, text, parse_mode="Markdown")

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
            "tr": f"📖 *Okul Yönetim Sistemi Kılavuzu*\n\nRolünüz: *{role_label}*\n• Alt menüyü kullanarak işlemlerinizi yapabilirsiniz.\n• İptal için /cancel yazabilirsiniz.\n• Destek için okul idaresine başvurabilirsiniz.",
            "ru": f"📖 *Справка по системе управления школой*\n\nВаша роль: *{role_label}*\n• Используйте нижнее меню для навигации.\n• Для отмены напишите /cancel.\n• По вопросам обращайтесь к администрации.",
            "uz": f"📖 *Maktab boshqaruv tizimi bo'yicha qo'llanma*\n\nSizning rolingiz: *{role_label}*\n• Pastdagi menyu orqali amallarni bajarishingiz mumkin.\n• Bekor qilish uchun /cancel yozishingiz mumkin.\n• Savollar bo'yicha maktab ma'muriyatiga murojaat qiling.",
            "en": f"📖 *School Management System Guide*\n\nYour Role: *{role_label}*\n• Use the bottom menu to navigate.\n• Type /cancel to abort any action.\n• Contact school administration for support."
        }
        await safe_edit_or_answer(message, help_texts.get(lang, help_texts["en"]), parse_mode="Markdown")

@router.message(any_state, Command("cancel", "iptal"))
@router.message(any_state, F.text.in_(["❌ İşlemi İptal Et", "❌ Отменить действие", "❌ Bekor qilish", "❌ Cancel Action", "iptal", "İptal", "cancel"]))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if user:
            reply_kb = get_role_reply_kb(user.role, user.language)
            await message.answer(get_text("action_cancelled", user.language), reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)

@router.message(CommandStart())
@router.message(Command("menu"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
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
            await message.answer(
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
        await session.commit()

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language))
            return

        if user.locked_until and user.locked_until > datetime.utcnow():
            rem_mins = max(1, int((user.locked_until - datetime.utcnow()).total_seconds() / 60))
            await message.answer(get_text("lock_countdown_msg", user.language, mins=rem_mins), parse_mode="Markdown")
            return

        if user.role != "admin":
            maint = await session.get(SystemSetting, "maintenance_mode")
            if maint and maint.value == "true":
                await message.answer(get_text("maintenance_mode", user.language))
                return

        if user.role == "guest":
            await prompt_guest_screen(message, user, state)
            return

        reply_kb = get_role_reply_kb(user.role, user.language)
        gen_welcome = {
            "tr": f"👋 *Hoş Geldiniz, {escape_md(user.full_name or '')}*",
            "ru": f"👋 *Добро пожаловать, {escape_md(user.full_name or '')}*",
            "uz": f"👋 *Xush kelibsiz, {escape_md(user.full_name or '')}*",
            "en": f"👋 *Welcome, {escape_md(user.full_name or '')}*"
        }.get(user.language, f"👋 *Welcome, {escape_md(user.full_name or '')}*")
        await message.answer(
            gen_welcome,
            reply_markup=reply_kb,
            parse_mode="Markdown"
        )
        await render_clean_dashboard(message, user)

async def prompt_guest_screen(target: Message | CallbackQuery, user: User, state: FSMContext):
    lang = user.language
    text = get_text("welcome_guest", lang)
    reply_kb = get_role_reply_kb("guest", lang)
    if isinstance(target, CallbackQuery):
        try: await target.message.delete()
        except Exception: pass
        await target.message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
    else:
        await target.answer(text, reply_markup=reply_kb, parse_mode="Markdown")

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

    if user.role == "guest":
        await prompt_guest_screen(query, user, state)
    else:
        reply_kb = get_role_reply_kb(user.role, user.language)
        await query.message.answer(get_text("lang_changed", user.language), reply_markup=reply_kb)
        await render_clean_dashboard(query, user)
    await query.answer()

async def render_clean_dashboard(target: Message | CallbackQuery | Bot, user: User, chat_id: int | None = None):
    lang = user.language
    async with AsyncSessionLocal() as session:
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
            subj = escape_md(tch.subject) if tch else "Ders"
            name = escape_md(tch.full_name) if tch else (escape_md(user.full_name or "Öğretmen"))

    date_str = get_local_date().strftime('%d.%m.%Y')
    if user.role == "admin":
        text = f"{get_text('admin_title', lang)}\n\n{get_text('admin_stats', lang, c_cnt=c_cnt, s_cnt=s_cnt, t_cnt=t_cnt, req_cnt=req_cnt, med_cnt=med_cnt, date=date_str)}"
    elif user.role == "teacher":
        text = get_text("menu_teacher", lang, name=name, subject=subj)
    elif user.role == "parent":
        text = get_text("menu_parent", lang, name=name, class_name=cls_name)
    elif user.role == "student":
        text = get_text("menu_student", lang, name=name, class_name=cls_name, no=num_val)
    else:
        text = get_text("welcome_guest", lang)

    if isinstance(target, Bot):
        target_id = chat_id or user.telegram_id
        await safe_send_message(target, target_id, text, parse_mode="Markdown")
        return

    await safe_edit_or_answer(target, text, parse_mode="Markdown")

async def process_auth_code_string(code: str, user_id: int, message: Message, state: FSMContext):
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

        if user.locked_until and user.locked_until > datetime.utcnow():
            rem_mins = max(1, int((user.locked_until - datetime.utcnow()).total_seconds() / 60))
            await message.answer(get_text("lock_countdown_msg", lang, mins=rem_mins), parse_mode="Markdown")
            return

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
                    await message.answer(get_text("auth_code_already_linked", lang), parse_mode="Markdown")
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
                    await message.answer(get_text("auth_code_already_linked", lang), parse_mode="Markdown")
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
        txt = get_text("auth_success", u_obj.language, name=escape_md(n_val), role=escape_md(r_lbl))
        await message.answer(txt, reply_markup=reply_kb, parse_mode="Markdown")
        await render_clean_dashboard(message, u_obj)

@router.message(Form.waiting_auth_code)
async def handle_auth_code_fsm(message: Message, state: FSMContext):
    await process_auth_code_string(message.text, message.from_user.id, message, state)

@router.callback_query(F.data == "act_enter_code")
async def cb_act_enter_code(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await query.message.answer(get_text("prompt_enter_code_direct", lang), parse_mode="Markdown")
    await state.set_state(Form.waiting_auth_code)
    await query.answer()

# ======================================================================
# 9. ERİŞİM TALEBİ / ŞİFRE BAŞVURU MOTORU
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

        buttons = [
            [InlineKeyboardButton(text=get_text("role_teacher_btn", lang), callback_data="req_role:teacher"), InlineKeyboardButton(text=get_text("role_parent_btn", lang), callback_data="req_role:parent")],
            [InlineKeyboardButton(text=get_text("role_student_btn", lang), callback_data="req_role:student")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(query, get_text("req_role_select", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("req_role:"))
async def cb_req_choose_role(query: CallbackQuery, state: FSMContext):
    chosen_role = query.data.split(":")[1]
    await state.update_data(role=chosen_role)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_cancel_reply_kb(lang)
    await query.message.answer(get_text("req_name_prompt", lang), reply_markup=cancel_kb, parse_mode="Markdown")
    await state.set_state(Form.req_name)
    await query.answer()

@router.message(Form.req_name)
async def process_req_name(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    full_name = message.text.strip()
    if not full_name or len(full_name) < 1:
        await message.answer(get_text("invalid_name_error", lang), parse_mode="Markdown")
        return
    await state.update_data(full_name=full_name)

    await message.answer(get_text("req_phone_prompt", lang), parse_mode="Markdown")
    await state.set_state(Form.req_phone)

@router.message(Form.req_phone)
async def process_req_phone(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    phone_val = message.text.strip()
    digits_only = re.sub(r"\D", "", phone_val)
    if len(digits_only) < 7:
        await message.answer(get_text("invalid_phone_error", lang), parse_mode="Markdown")
        return
    await state.update_data(phone=phone_val)

    data = await state.get_data()
    role = data.get("role", "student")

    prompt_key = "req_details_teacher" if role == "teacher" else ("req_details_parent" if role == "parent" else "req_details_student")
    await message.answer(get_text(prompt_key, lang), parse_mode="Markdown")
    await state.set_state(Form.req_details)

@router.message(Form.req_details)
async def process_req_details(message: Message, state: FSMContext):
    details_text = message.text.strip()
    data = await state.get_data()
    await state.clear()

    role = data.get("role", "student")
    full_name = data.get("full_name", "")
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
            phone=phone_val,
            details=details_text,
            student_match_id=student_match.id if student_match else None,
            status="pending"
        )
        session.add(req)
        await session.commit()

        await message.answer(get_text("req_sent_success", lang), parse_mode="Markdown")
        await prompt_guest_screen(message, user, state)

        role_label = {
            "teacher": get_text("lbl_role_teacher", lang),
            "parent": get_text("lbl_role_parent", lang),
            "student": get_text("lbl_role_student", lang)
        }.get(role, role)

        auto_check_badge = f"🟢 *Sistem Eşleşmesi:* {escape_md(student_match.full_name)} ({escape_md(student_match.class_name)} - No: {escape_md(student_match.student_number)})" if student_match else "🔴 *Otomatik Eşleşme Yok*"

        adm_msg = (
            f"🛎️ *YENİ ERİŞİM / ŞİFRE TALEBİ* (#{req.id})\n\n"
            f"👤 *Başvuran:* {escape_md(full_name)}\n"
            f"🆔 *Telegram:* @{escape_md(message.from_user.username or 'gizli')} (`{message.from_user.id}`)\n"
            f"📱 *Telefon:* `{phone_val}`\n"
            f"🎯 *Talep Rolü:* *{role_label}*\n"
            f"📝 *Açıklama:* {escape_md(details_text)}\n\n"
            f"🔍 *Durum:* {auto_check_badge}"
        )
        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text=get_text("btn_appr_request", lang), callback_data=f"adm:appr_req:{req.id}"),
                InlineKeyboardButton(text=get_text("btn_reject", lang), callback_data=f"adm:rej_req:{req.id}")
            ]
        ])

        admins = (await session.execute(select(User.telegram_id).where(User.role == "admin"))).scalars().all()
        for a_id in set(ADMIN_IDS + list(admins)):
            await safe_send_message(message.bot, a_id, adm_msg, reply_markup=adm_kb, parse_mode="Markdown")

@router.callback_query(F.data == "adm:requests_list")
async def cb_admin_requests_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        requests = (await session.execute(select(AccessRequest).where(AccessRequest.status == "pending").order_by(desc(AccessRequest.created_at)))).scalars().all()
        if not requests:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_pending_requests", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        for r in requests:
            icon = "👨‍🏫" if r.role == "teacher" else ("👨‍👩‍👧‍👦" if r.role == "parent" else "🎓")
            buttons.append([InlineKeyboardButton(text=f"{icon} {r.full_name} ({r.created_at.strftime('%H:%M')})", callback_data=f"adm:view_req:{r.id}")])
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
            if st: st_match_info = f"\n🔍 *Eşleşen Öğrenci:* {escape_md(st.full_name)} ({escape_md(st.class_name)} - No: {escape_md(st.student_number)})"

        text = (
            f"🛎️ *Erişim Başvurusu Detayı* (#{req.id})\n\n"
            f"👤 *İsim:* {escape_md(req.full_name)}\n"
            f"📱 *Telefon:* `{req.phone}`\n"
            f"🎯 *Rol:* *{role_label}*\n"
            f"📝 *Detay:* {escape_md(req.details)}{st_match_info}\n\n"
            f"Lütfen yapılacak işlemi seçiniz:"
        )
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_appr_request", lang), callback_data=f"adm:appr_req:{req.id}"), InlineKeyboardButton(text=get_text("btn_reject", lang), callback_data=f"adm:rej_req:{req.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:requests_list")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:appr_req:"))
async def cb_admin_approve_request(query: CallbackQuery):
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
            return

        req.status = "approved"
        req.reviewed_by = query.from_user.id
        req.reviewed_at = datetime.utcnow()

        target_user = await session.get(User, req.telegram_id)
        if not target_user:
            target_user = User(telegram_id=req.telegram_id, language="tr")
            session.add(target_user)

        target_user.role = req.role
        target_user.full_name = req.full_name
        target_user.failed_attempts = 0
        target_user.locked_until = None

        if req.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.full_name == req.full_name))).scalar_one_or_none()
            if not tch:
                tch = Teacher(full_name=req.full_name, subject=req.details, auth_code=generate_secure_code("HCA"), telegram_id=req.telegram_id, is_code_burned=True)
                session.add(tch)
            else:
                tch.telegram_id = req.telegram_id
                tch.is_code_burned = True
        elif req.role == "parent":
            st_id = req.student_match_id
            if not st_id:
                all_st = (await session.execute(select(Student))).scalars().all()
                for s in all_st:
                    if s.full_name.lower() in req.details.lower():
                        st_id = s.id
                        break
            if st_id:
                target_user.current_child_id = st_id
                rel = (await session.execute(select(ParentStudent).where(ParentStudent.parent_telegram_id == req.telegram_id, ParentStudent.student_id == st_id))).scalar_one_or_none()
                if not rel:
                    session.add(ParentStudent(parent_telegram_id=req.telegram_id, student_id=st_id))
        elif req.role == "student":
            if req.student_match_id:
                st = await session.get(Student, req.student_match_id)
                if st:
                    st.student_telegram_id = req.telegram_id
                    st.is_student_code_burned = True

        await session.commit()
        await safe_send_message(query.message.bot, req.telegram_id, get_text("req_approved_user", target_user.language, role=req.role), reply_markup=get_role_reply_kb(req.role, target_user.language), parse_mode="Markdown")
        await safe_edit_or_answer(query, get_text("req_approved_admin_msg", lang, id=req.id, name=escape_md(req.full_name), role=req.role), parse_mode="Markdown")
    await query.answer(get_text("acknowledged_toast", lang))

@router.callback_query(F.data.startswith("adm:rej_req:"))
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
            return

        req.status = "rejected"
        req.reviewed_by = query.from_user.id
        req.reviewed_at = datetime.utcnow()
        target_user = await session.get(User, req.telegram_id)
        await session.commit()

        await safe_send_message(query.message.bot, req.telegram_id, get_text("req_rejected_user", target_user.language if target_user else "tr"), parse_mode="Markdown")
        await safe_edit_or_answer(query, get_text("req_rejected_admin_msg", lang, id=req.id, name=escape_md(req.full_name)), parse_mode="Markdown")
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
            st_name = escape_md(st.full_name) if st else "Öğrenci"
            cls_name = escape_md(st.class_name) if st else ""
            t_text = (
                f"🤝 *YENİ VELİ GÖRÜŞME TALEBİ*\n\n"
                f"👤 *Veli:* {escape_md(user.full_name or 'Veli')}\n"
                f"🧑‍🎓 *Öğrenci:* {st_name} ({cls_name})\n"
                f"📝 *Talep / Zaman:* {escape_md(note_text)}"
            )
            t_kb = InlineKeyboardMarkup(inline_keyboard=[
                [
                    InlineKeyboardButton(text=get_text("btn_appr_appointment", tch_lang), callback_data=f"tch:appr_app:{app.id}"),
                    InlineKeyboardButton(text=get_text("btn_not_available", tch_lang), callback_data=f"tch:rej_app:{app.id}")
                ]
            ])
            await safe_send_message(message.bot, tch.telegram_id, t_text, reply_markup=t_kb, parse_mode="Markdown")

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
            await safe_edit_or_answer(query, get_text("no_pending_appointments", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        for a in apps:
            btn_txt = f"🤝 Randevu #{a.id} ({a.preferred_time[:20]})"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"tch:view_app:{a.id}")])
        buttons.append(get_nav_buttons(lang))

        await safe_edit_or_answer(query, get_text("pending_appointments_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, app.parent_telegram_id, get_text("appointment_approved_msg", p_lang), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, app.parent_telegram_id, get_text("appointment_rejected_msg", p_lang), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "act_view_sched")
async def cb_view_schedule(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        if user and user.role == "teacher":
            classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
            if not classes: classes = ["9-A"]
            buttons = []
            row = []
            for c in classes:
                row.append(InlineKeyboardButton(text=f"📅 {c}", callback_data=f"sched_cls:{c}"))
                if len(row) == 2:
                    buttons.append(row)
                    row = []
            if row: buttons.append(row)
            buttons.append(get_nav_buttons(lang))
            await safe_edit_or_answer(query, get_text("schedule_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        cls_name = "9-A"
        if user and user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: cls_name = st.class_name
        elif user and user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == query.from_user.id))).scalar_one_or_none()
            if st: cls_name = st.class_name

        sched = (await session.execute(select(ScheduleItem).where(ScheduleItem.class_name == cls_name))).scalar_one_or_none()
        content = sched.schedule_text if sched else "• 1. Ders: 09:00 - Matematik\n• 2. Ders: 09:50 - Fizik\n• 3. Ders: 10:40 - Türkçe\n• 4. Ders: 11:30 - Tarih\n• 5. Ders: 13:00 - Biyoloji"

        header_sc = {"tr": f"📅 *{escape_md(cls_name)} Sınıfı Ders Programı:*", "ru": f"📅 *Расписание уроков {escape_md(cls_name)} класса:*", "uz": f"📅 *{escape_md(cls_name)} sinfi dars jadvali:*", "en": f"📅 *Timetable for Class {escape_md(cls_name)}:*"}.get(lang, f"📅 *{escape_md(cls_name)} Timetable:*")
        text = f"{header_sc}\n\n{escape_md(content)}"
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await safe_edit_or_answer(query, get_text("no_blacklisted", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, get_text("blacklisted_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, t_id, get_text("admin_unban_notification", target_u.language if target_u else "tr"), parse_mode="Markdown")
            await query.answer(get_text("unban_success", lang), show_alert=True)
            await cb_admin_blacklist(query)
            return
    await query.answer()

@router.callback_query(F.data == "adm:sched_edit_menu")
async def cb_admin_sched_edit_menu(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes: classes = ["9-A"]

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"✏️ {c} {get_text('btn_class_sched', lang)}", callback_data=f"adm:sched_edit_cls:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append(get_nav_buttons(lang))

        await safe_edit_or_answer(query, get_text("admin_sched_edit_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
    await safe_edit_or_answer(query, prompt_t, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.sched_update_text)
    await query.answer()

@router.message(Form.sched_update_text)
async def process_sched_update_text(message: Message, state: FSMContext):
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

        await message.answer(get_text("admin_sched_updated", lang, class_name=escape_md(class_name)), reply_markup=get_role_reply_kb("admin", lang), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data == "adm:menu_edit")
async def cb_admin_menu_edit(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("prompt_menu_update", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.menu_update_text)
    await query.answer()

@router.message(Form.menu_update_text)
async def process_menu_update_text(message: Message, state: FSMContext):
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
# 11. YÖNETİCİ İŞLEMLERİ (ÖĞRENCİ, ÖĞRETMEN, SINIFLAR, KOKPİT, EXCEL, RAPORLAR)
# ======================================================================

@router.callback_query(F.data == "adm:add_student")
async def cb_start_add_student(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_cancel_reply_kb(lang)
    await query.message.answer(get_text("prompt_student_name", lang), reply_markup=cancel_kb, parse_mode="Markdown")
    await state.set_state(Form.add_student_name)
    await query.answer()

@router.message(Form.add_student_name)
async def process_student_name(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    name_val = message.text.strip()
    if not name_val or len(name_val) < 1:
        await message.answer(get_text("student_name_invalid", lang))
        return
    await state.update_data(name=name_val)

    await message.answer(get_text("prompt_student_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]), parse_mode="Markdown")
    await state.set_state(Form.add_student_class)

@router.message(Form.add_student_class)
async def process_student_class(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    await state.update_data(class_name=message.text.strip().upper())
    await message.answer(get_text("prompt_student_no", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]), parse_mode="Markdown")
    await state.set_state(Form.add_student_no)

@router.message(Form.add_student_no)
async def process_student_no(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        data = await state.get_data()
        await state.clear()
        full_name = data.get("name")
        class_name = data.get("class_name")
        student_no = message.text.strip()

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
            [InlineKeyboardButton(text=get_text("btn_add_another", lang), callback_data="adm:add_student")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(message, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "adm:add_teacher")
async def cb_start_add_teacher(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_cancel_reply_kb(lang)
    await query.message.answer(get_text("prompt_teacher_name", lang), reply_markup=cancel_kb, parse_mode="Markdown")
    await state.set_state(Form.add_teacher_name)
    await query.answer()

@router.message(Form.add_teacher_name)
async def process_teacher_name(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    name_val = message.text.strip()
    if not name_val or len(name_val) < 1:
        await message.answer(get_text("teacher_name_invalid", lang))
        return
    await state.update_data(name=name_val)

    await message.answer(get_text("prompt_teacher_subject", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]), parse_mode="Markdown")
    await state.set_state(Form.add_teacher_subject)

@router.message(Form.add_teacher_subject)
async def process_teacher_subject(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        data = await state.get_data()
        await state.clear()
        full_name = data.get("name")
        subject = message.text.strip()
        auth_code = generate_secure_code("HCA")

        teacher = Teacher(full_name=full_name, subject=subject, auth_code=auth_code)
        session.add(teacher)
        await session.commit()

        text = get_text("teacher_added_card", lang, name=escape_md(full_name), subject=escape_md(subject), code=auth_code)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_add_another", lang), callback_data="adm:add_teacher")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await safe_edit_or_answer(message, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "adm:search_student")
async def cb_search_student_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("prompt_search_student", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_search_query)
    await query.answer()

@router.message(Form.waiting_search_query)
async def process_search_query(message: Message, state: FSMContext):
    query_text = message.text.strip()
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        stmt = select(Student).where(
            (func.lower(Student.full_name).contains(query_text.lower())) |
            (Student.student_number == query_text)
        ).limit(10)
        results = (await session.execute(stmt)).scalars().all()

        if not results:
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_search_again", lang), callback_data="adm:search_student")],
                [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
            ]
            await message.answer(get_text("search_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        buttons = []
        for s in results:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.class_name} - №{s.student_number})", callback_data=f"adm:st_card:{s.id}")])

        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(message, get_text("search_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "adm:search_teacher")
async def cb_search_teacher_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("teacher_search_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_search_teacher_query)
    await query.answer()

@router.message(Form.waiting_search_teacher_query)
async def process_search_teacher_query(message: Message, state: FSMContext):
    query_text = message.text.strip()
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        stmt = select(Teacher).where((func.lower(Teacher.full_name).contains(query_text.lower())) | (func.lower(Teacher.subject).contains(query_text.lower()))).limit(10)
        results = (await session.execute(stmt)).scalars().all()

        if not results:
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_search_again", lang), callback_data="adm:search_teacher")],
                [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
            ]
            await message.answer(get_text("teacher_search_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        buttons = []
        for t in results:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {t.full_name} ({t.subject})", callback_data=f"adm:tch_card:{t.id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(message, get_text("teacher_search_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "adm:teachers")
async def cb_admin_teachers_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        buttons = []
        for t in teachers:
            status_dot = "🟢" if t.telegram_id else "⚪"
            buttons.append([InlineKeyboardButton(text=f"{status_dot} {t.full_name} ({t.subject})", callback_data=f"adm:tch_card:{t.id}")])

        buttons.append([InlineKeyboardButton(text=get_text("btn_add_new_teacher", lang), callback_data="adm:add_teacher")])
        buttons.append(get_nav_buttons(lang))

        teachers_list_title = {
            "tr": f"👨‍🏫 *Öğretmenler Listesi* (Toplam {len(teachers)} Öğretmen):\nDetay veya şifre işlemleri için tıklayınız:",
            "ru": f"👨‍🏫 *Список учителей* (Всего {len(teachers)}):\nНажмите для просмотра данных или пароля:",
            "uz": f"👨‍🏫 *O'qituvchilar ro'yxati* (Jami {len(teachers)} ta o'qituvchi):\nBatafsil ma'lumot yoki parollar uchun tanlang:",
            "en": f"👨‍🏫 *Teachers List* (Total {len(teachers)} Teachers):\nTap to view details or credentials:"
        }.get(lang, f"👨‍🏫 *Teachers List*\n")
        text = teachers_list_title
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

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

        status_str = f"Bağlı (`{tch.telegram_id}`)" if tch.telegram_id else "Henüz Giriş Yapmadı"
        text = get_text("teacher_card", lang, name=escape_md(tch.full_name), subject=escape_md(tch.subject), code=tch.auth_code, status=status_str)

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_reset_codes", lang), callback_data=f"adm:reset_tch:{tch.id}")],
            [InlineKeyboardButton(text=get_text("btn_del_teacher", lang), callback_data=f"adm:del_tch:{tch.id}")],
            get_nav_buttons(lang, back_callback="adm:teachers")
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:reset_tch:"))
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
            text = f"✅ Giriş Kodu: `{tch.auth_code}`"
            buttons = [get_nav_buttons(lang, back_callback=f"adm:tch_card:{tch.id}")]
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:del_tch:"))
async def cb_admin_delete_teacher(query: CallbackQuery):
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
async def cb_classes_list(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            text = get_text("no_classes_found", lang)
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_add_student", lang), callback_data="adm:add_student")],
                [InlineKeyboardButton(text=get_text("btn_excel", lang), callback_data="adm:excel_info")],
                get_nav_buttons(lang)
            ]
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        buttons.append(get_nav_buttons(lang))
        prompt_c = {
            "tr": f"🏫 *{get_text('btn_classes', lang)}*\n\nİncelemek istediğiniz sınıfı seçiniz:",
            "ru": f"🏫 *{get_text('btn_classes', lang)}*\n\nВыберите класс для просмотра:",
            "uz": f"🏫 *{get_text('btn_classes', lang)}*\n\nKo'rmoqchi bo'lgan sinfni tanlang:",
            "en": f"🏫 *{get_text('btn_classes', lang)}*\n\nSelect a class to view:"
        }.get(lang, f"🏫 *{get_text('btn_classes', lang)}*\n")
        await safe_edit_or_answer(query, prompt_c, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "noop")
async def cb_noop(query: CallbackQuery):
    await query.answer()

# ======================================================================
# 12. KATEGORİ HUB'LARI VE SINIF/ÖĞRENCİ İŞLEMLERİ
# ======================================================================

@router.message(any_state, F.text.in_(["👥 Kadro & Öğrenci", "👥 Ученики и учителя", "👥 Kadro va o'quvchilar", "👥 Staff & Students"]))
async def cb_cat_staff(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"), InlineKeyboardButton(text=get_text("btn_teachers", lang), callback_data="adm:teachers")],
            [InlineKeyboardButton(text=get_text("btn_search_student", lang), callback_data="adm:search_student"), InlineKeyboardButton(text=get_text("btn_search_teacher", lang), callback_data="adm:search_teacher")],
            [InlineKeyboardButton(text=get_text("btn_school_admins", lang), callback_data="adm:admins_list"), InlineKeyboardButton(text=get_text("btn_users_hub", lang), callback_data="adm:users_hub:0")]
        ]
        await safe_edit_or_answer(message, get_text("cat_staff_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_(["📊 Raporlar & Denetim", "📊 Отчеты и контроль", "📊 Hisobotlar va nazorat", "📊 Reports & Audits"]))
async def cb_cat_reports(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_cockpit_unified", lang), callback_data="adm:cockpit")],
            [InlineKeyboardButton(text=get_text("btn_risk_radar", lang), callback_data="adm:risk_radar"), InlineKeyboardButton(text=get_text("btn_academic_report", lang), callback_data="adm:academic_report")],
            [InlineKeyboardButton(text=get_text("btn_unack_notifs", lang), callback_data="adm:unack_notifs")]
        ]
        await safe_edit_or_answer(message, get_text("cat_reports_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_(["🛎️ Onay Masası", "🛎️ Центр одобрений", "🛎️ Tasdiqlash markazi", "🛎️ Approval Center"]))
async def cb_cat_requests(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
        med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_requests", lang, count=req_cnt), callback_data="adm:requests_list")],
            [InlineKeyboardButton(text=get_text("btn_medical", lang, count=med_cnt), callback_data="adm:medical_list")],
            [InlineKeyboardButton(text=get_text("btn_audit_logs", lang), callback_data="adm:audit_logs")]
        ]
        await safe_edit_or_answer(message, get_text("cat_requests_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_(["🛠️ İdari Araçlar", "🛠️ Инструменты", "🛠️ Boshqaruv vositalari", "🛠️ Admin Tools"]))
async def cb_cat_tools(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_init"), InlineKeyboardButton(text=get_text("btn_excel_hub", lang), callback_data="adm:excel_hub")],
            [InlineKeyboardButton(text=get_text("btn_manage_schedule", lang), callback_data="adm:sched_edit_menu"), InlineKeyboardButton(text=get_text("btn_cafeteria_edit", lang), callback_data="adm:menu_edit")],
            [InlineKeyboardButton(text=get_text("btn_pdf", lang), callback_data="adm:pdf_menu")]
        ]
        await safe_edit_or_answer(message, get_text("cat_tools_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_(["⚙️ Sistem & Ayarlar", "⚙️ Настройки системы", "⚙️ Tizim va sozlamalar", "⚙️ System & Settings"]))
async def cb_cat_settings(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        maint = await session.get(SystemSetting, "maintenance_mode")
        is_maint = maint.value == "true" if maint else False
        maint_txt = get_text("btn_maintenance_toggle", lang, status="AÇIK" if is_maint else "KAPALI")

        buttons = [
            [InlineKeyboardButton(text=maint_txt, callback_data="adm:toggle_maint"), InlineKeyboardButton(text=get_text("btn_blacklist", lang), callback_data="adm:blacklist")],
            [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")]
        ]
        await safe_edit_or_answer(message, get_text("cat_settings_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

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
            "tr": f"📋 *{escape_md(class_name)} Sınıfı Toplu Devamsızlık Çizelgesi:*\n",
            "ru": f"📋 *Ведомость посещаемости класса {escape_md(class_name)}:*\n",
            "uz": f"📋 *{escape_md(class_name)} sinfining umumiy davomat qaydnomasi:*\n",
            "en": f"📋 *Class Attendance Sheet for {escape_md(class_name)}:*\n"
        }.get(lang, f"📋 *Attendance: {escape_md(class_name)}*\n")

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
            lines.append(f"• *{escape_md(s.full_name)}* (№{s.student_number}): *{abs_cnt} {lbl_day}*{warn_badge} _({lbl_exc}: {exc_cnt})_")

        avg_abs = round(total_abs_class / len(students), 1) if students else 0
        lbl_total_st = {"tr": "Sınıf Mevcudu", "ru": "Всего учеников", "uz": "Sinf mevcudi", "en": "Class Total"}.get(lang, "Total")
        lbl_avg_abs = {"tr": "Ortalama Devamsızlık", "ru": "Средний пропуск", "uz": "O'rtacha davomat", "en": "Average Absence"}.get(lang, "Avg")
        lines.append(f"\n📊 *{lbl_total_st}:* {len(students)} | *{lbl_avg_abs}:* {avg_abs} {lbl_day}")

        buttons = [get_nav_buttons(lang, back_callback=f"adm:show_class:{class_name}")]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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

        start_idx = page * PAGE_SIZE
        paged_students = all_students[start_idx : start_idx + PAGE_SIZE]

        buttons = []
        for s in paged_students:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"adm:st_card:{s.id}")])

        if total_pages > 1:
            nav_row = []
            if page > 0:
                nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:show_class:{class_name}:{page - 1}"))
            nav_row.append(InlineKeyboardButton(text=f"📄 {page + 1}/{total_pages}", callback_data="noop"))
            if page < total_pages - 1:
                nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:show_class:{class_name}:{page + 1}"))
            buttons.append(nav_row)

        buttons.append([InlineKeyboardButton(text=f"📋 {class_name} {get_text('btn_class_att_sheet', lang)}", callback_data=f"adm:class_att_sheet:{class_name}")])
        buttons.append([InlineKeyboardButton(text=f"📄 {class_name} {get_text('btn_class_pdf_cards', lang)}", callback_data=f"adm:gen_pdf:{class_name}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:classes"))

        class_list_title = {
            "tr": f"🏫 *{escape_md(class_name)} Sınıfı Listesi* (Toplam {total_students} Öğrenci):\nDetay veya şifre işlemleri için öğrenciye tıklayınız:",
            "ru": f"🏫 *Список класса {escape_md(class_name)}* (Всего {total_students} учеников):\nНажмите на ученика для просмотра данных или кодов:",
            "uz": f"🏫 *{escape_md(class_name)} sinfi ro'yxati* (Jami {total_students} o'quvchi):\nMa'lumot yoki parollar uchun o'quvchini tanlang:",
            "en": f"🏫 *Class {escape_md(class_name)} List* (Total {total_students} Students):\nTap a student to view details or access credentials:"
        }.get(lang, f"🏫 *Class {escape_md(class_name)}*\n")
        text = class_list_title
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            "tr": f"✏️ *{escape_md(st.full_name)}* ({escape_md(st.class_name)} - No: {escape_md(st.student_number)})\nDüzenlemek istediğiniz bilgiyi seçiniz:",
            "ru": f"✏️ *{escape_md(st.full_name)}* ({escape_md(st.class_name)} - №: {escape_md(st.student_number)})\nВыберите параметр для изменения:",
            "uz": f"✏️ *{escape_md(st.full_name)}* ({escape_md(st.class_name)} - №: {escape_md(st.student_number)})\nTahrirlamoqchi bo'lgan ma'lumotni tanlang:",
            "en": f"✏️ *{escape_md(st.full_name)}* ({escape_md(st.class_name)} - Roll: {escape_md(st.student_number)})\nSelect information to edit:"
        }.get(lang, f"✏️ *{escape_md(st.full_name)}*\n")
        await safe_edit_or_answer(query, prompt_es, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await message.answer(get_text("student_info_updated", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name), no=escape_md(st.student_number)), parse_mode="Markdown")

@router.callback_query(F.data.startswith("adm:st_card:"))
async def cb_student_card(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
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

        text = get_text("student_card", lang, name=escape_md(st.full_name), class_name=escape_md(st.class_name), no=escape_md(st.student_number), st_code=st.student_code, st_status=st_status, pr_code=st.parent_code, pr_status=pr_status)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_edit_student", lang), callback_data=f"adm:edit_st:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_reset_codes", lang), callback_data=f"adm:reset_codes:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_unlink_parent", lang), callback_data=f"adm:unlink_pr:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_del_student", lang), callback_data=f"adm:del_student:{st.id}")],
            get_nav_buttons(lang, back_callback=f"adm:show_class:{st.class_name}")
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:reset_codes:"))
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
            await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
async def cb_delete_student(query: CallbackQuery):
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

@router.callback_query(F.data == "adm:dashboard")
async def cb_admin_dashboard(query: CallbackQuery, state: FSMContext | None = None):
    if state: await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user: await render_clean_dashboard(query, user)
    await query.answer()

@router.callback_query(F.data == "adm:unack_notifs")
async def cb_admin_unack_notifs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        since = datetime.utcnow() - timedelta(hours=36)
        unacks = (await session.execute(select(CriticalNotification).where(CriticalNotification.created_at >= since, CriticalNotification.acknowledged_at == None).order_by(desc(CriticalNotification.created_at)).limit(20))).scalars().all()

        if not unacks:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("all_notifs_acknowledged", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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

        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        missing_str = "\n".join([f"• ❌ *{escape_md(c)}*" for c in missing]) if missing else all_done_txt

        text = get_text("cockpit_report", lang, date=today.strftime("%d.%m.%Y"), total=total_students, present=present_count, absent=absent_count, missing_cnt=len(missing), missing=missing_str)
        buttons = []
        if missing:
            buttons.append([InlineKeyboardButton(text=get_text("btn_remind_att", lang), callback_data="adm:remind_all_att")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:risk_radar")
async def cb_admin_risk_radar(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        students = (await session.execute(select(Student).order_by(Student.class_name, Student.student_number))).scalars().all()
        risky_items = []
        lbl_day = {"tr": "gün", "ru": "дн.", "uz": "kun", "en": "days"}.get(lang, "days")
        lbl_att_r = {"tr": "Devamsızlık:", "ru": "Пропуски:", "uz": "Davomat:", "en": "Absences:"}.get(lang, "Absences:")
        lbl_low_gr = {"tr": "Düşük Not:", "ru": "Низкий балл:", "uz": "Past baho:", "en": "Low Grade:"}.get(lang, "Low Grade:")
        lbl_warn_bdg = {"tr": "🔴 Uyarı Rozeti", "ru": "🔴 Замечание", "uz": "🔴 Ogohlantirish", "en": "🔴 Warning Badge"}.get(lang, "🔴 Warning")

        for s in students:
            abs_cnt = (await session.execute(select(func.count(Attendance.id)).where(Attendance.student_id == s.id, Attendance.status == "absent"))).scalar() or 0
            grades = (await session.execute(select(Grade).where(Grade.student_id == s.id))).scalars().all()
            avg_score = (sum(g.score for g in grades) / len(grades)) if grades else 100.0
            has_warning = any(g.badge == "🔴" for g in grades)

            reasons = []
            if abs_cnt >= 7: reasons.append(f"{lbl_att_r} {abs_cnt} {lbl_day}")
            if avg_score < 50.0: reasons.append(f"{lbl_low_gr} {round(avg_score, 1)}")
            if has_warning: reasons.append(lbl_warn_bdg)

            if reasons:
                risky_items.append(f"• *{escape_md(s.full_name)}* ({escape_md(s.class_name)} - No: {escape_md(s.student_number)}):\n  ↳ _{escape_md(', '.join(reasons))}_")

        no_risk_txt = {"tr": "✅ Şu anda risk grubunda bulunan öğrenci bulunmamaktadır.", "ru": "✅ Нет учеников в группе риска.", "uz": "✅ Hozirda xavf guruhidagi o'quvchilar yo'q.", "en": "✅ No at-risk students currently."}.get(lang, "✅ No at-risk students.")
        header_rr = {"tr": "⚠️ *Riskli Öğrenci Radarı (Devamsızlık & Düşük Not):*\n\n", "ru": "⚠️ *Радар рисков (Пропуски и низкие оценки):*\n\n", "uz": "⚠️ *Xavf ostidagi o'quvchilar radari (Davomat va past baholar):*\n\n", "en": "⚠️ *At-Risk Student Radar (Absences & Low Grades):*\n\n"}.get(lang, "⚠️ *At-Risk Student Radar:*\n\n")
        content = "\n\n".join(risky_items[:20]) if risky_items else no_risk_txt
        text = f"{header_rr}{content}"
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:audit_logs")
async def cb_admin_audit_logs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        logs = (await session.execute(select(AuditLog).order_by(desc(AuditLog.created_at)).limit(12))).scalars().all()
        no_logs_txt = {"tr": "✅ Henüz kayıtlı işlem geçmişi bulunmamaktadır.", "ru": "✅ Журнал действий пуст.", "uz": "✅ Hozircha tizim jurnali bo'sh.", "en": "✅ No audit logs recorded yet."}.get(lang, "✅ No logs.")
        header_al = {"tr": "📜 *Son Sistem İşlem Geçmişi:*\n\n", "ru": "📜 *Журнал действий системы:*\n\n", "uz": "📜 *So'nggi tizim amallari jurnali:*\n\n", "en": "📜 *Recent System Audit Logs:*\n\n"}.get(lang, "📜 *System Audit Logs:*\n\n")

        if not logs:
            content = no_logs_txt
        else:
            lines = [f"📌 *{l.created_at.strftime('%d.%m %H:%M')}* - *{escape_md(l.action)}*\n↳ {escape_md(l.user_name)}: _{escape_md(l.details)}_" for l in logs]
            content = "\n\n".join(lines)

        text = f"{header_al}{content}"
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:academic_report")
async def cb_admin_academic_report(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

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
            lines = [f"🥇 *{escape_md(c)} {lbl_cls_word}:* {lbl_avg_word} *{avg}* ({cnt} {lbl_gr_cnt})" if idx==1 else (f"🥈 *{escape_md(c)} {lbl_cls_word}:* {lbl_avg_word} *{avg}* ({cnt} {lbl_gr_cnt})" if idx==2 else f"{idx}. *{escape_md(c)} {lbl_cls_word}:* {lbl_avg_word} *{avg}* ({cnt} {lbl_gr_cnt})") for idx, (c, avg, cnt) in enumerate(rankings, 1)]
            content = "\n".join(lines)

        header_ar = {"tr": "📈 *Okul Akademik Başarı Sıralaması (Sınıflar):*\n\n", "ru": "📈 *Рейтинг классов по успеваемости:*\n\n", "uz": "📈 *Sinflarning akademik reytingi:*\n\n", "en": "📈 *Academic Ranking of Classes:*\n\n"}.get(lang, "📈 *Academic Ranking:*\n\n")
        text = f"{header_ar}{content}"
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:teacher_attendance_check")
async def cb_admin_att_check(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        today = get_local_date()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()

        missing = [c for c in all_classes if c not in taken_classes]
        buttons = []
        if not missing:
            text = f"{get_text('att_check_title', lang, date=today.strftime('%d.%m.%Y'))}\n\n{get_text('att_check_all_done', lang)}"
        else:
            missing_lines = [f"• ❌ *{escape_md(c)}*" for c in missing]
            missing_title = {"tr": "⚠️ *Yoklama Almayan Sınıflar:*", "ru": "⚠️ *Классы без переклички:*", "uz": "⚠️ *Davomat olinmagan sinflar:*", "en": "⚠️ *Classes without attendance:*"}.get(lang, "⚠️ *Pending:*")
            text = (
                f"{get_text('att_check_title', lang, date=today.strftime('%d.%m.%Y'))}\n\n"
                f"{missing_title} ({len(missing)}):\n"
                + "\n".join(missing_lines)
            )
            buttons.append([InlineKeyboardButton(text=get_text("btn_remind_att", lang), callback_data="adm:remind_all_att")])

        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, t.telegram_id, remind_m, parse_mode="Markdown")
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
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_classes_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("select_pdf_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        "en": "👨‍🏫 *Access credential cards for all teachers are attached.*"
    }.get(lang, "👨‍🏫 *Teacher Access Cards Attached.*")
    await query.message.answer_document(file, caption=pdf_tch_caption, parse_mode="Markdown")
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
        await query.message.answer_document(file, caption=caption, parse_mode="Markdown")
        pdf_buffer.close()
    await query.answer()

@router.callback_query(F.data == "adm:export_all_excel")
async def cb_admin_export_all(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buf = await export_all_school_data_excel()
    today_str = datetime.utcnow().strftime("%d_%m_%Y")
    file = BufferedInputFile(buf.read(), filename=f"Okul_Genel_Yedek_{today_str}.xlsx")
    await query.message.answer_document(file, caption=get_text("export_ready", lang, date=today_str), parse_mode="Markdown")
    buf.close()
    await query.answer()

@router.callback_query(F.data == "adm:excel_info")
async def cb_excel_info(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("excel_info", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.message(F.document, F.document.file_name.endswith(".xlsx"))
async def admin_excel_upload(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id):
            await message.answer(get_text("unauthorized_excel_upload", lang), parse_mode="Markdown")
            return

    try:
        bot: Bot = message.bot
        file_info = await bot.get_file(message.document.file_id)
        file_bytes = await bot.download_file(file_info.file_path)

        count, out_excel = await process_student_excel(file_bytes.read())
        file = BufferedInputFile(out_excel.read(), filename="Giris_Kodlari_Uretildi.xlsx")
        await message.answer_document(file, caption=get_text("excel_done", lang, count=count), parse_mode="Markdown")
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
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_pending_medical", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        for rep, st in reports:
            btn_txt = f"🏥 {st.full_name} ({st.class_name}) - {rep.created_at.strftime('%H:%M')}"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:view_med:{rep.id}")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, get_text("pending_medical_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        caption = f"🏥 *{get_text('btn_medical', lang, count=1)}*\n{lbl_st}: *{escape_md(st.full_name)}* ({escape_md(st.class_name)})\n{lbl_note} {escape_md(rep.caption or '-')}"
        try: await query.message.delete()
        except Exception: pass
        await query.message.bot.send_photo(chat_id=query.from_user.id, photo=rep.file_id, caption=caption, reply_markup=adm_kb, parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, rep.parent_telegram_id, get_text("medical_approved_parent", p_lang), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, rep.parent_telegram_id, get_text("medical_rejected_parent", p_lang), parse_mode="Markdown")
            try: await query.message.delete()
            except Exception: pass
            await query.message.answer(get_text("medical_rejected", lang))
            if user: await render_clean_dashboard(query.message, user)
    await query.answer()

@router.callback_query(F.data == "adm:broadcast_init")
async def cb_broadcast_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("prompt_broadcast", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_broadcast_text)
    await query.answer()

@router.message(Form.waiting_broadcast_text)
async def process_broadcast_text(message: Message, state: FSMContext):
    b_text = message.text.strip()
    if len(b_text) > 4000: b_text = b_text[:4000] + "\n...(Metin kısaltıldı)"
    await state.clear()
    sent_cnt = 0
    safe_text = html.escape(b_text)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        notice = BroadcastNotice(content=b_text)
        session.add(notice)
        await session.commit()

        users = (await session.execute(select(User.telegram_id))).scalars().all()
        for u_id in users:
            try:
                await message.bot.send_message(chat_id=u_id, text=f"📢 <b>OKUL DUYURUSU</b>\n\n{safe_text}", parse_mode="HTML")
                sent_cnt += 1
                await asyncio.sleep(0.05)
            except Exception: pass

        await message.answer(get_text("broadcast_success", lang, count=sent_cnt), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

# ======================================================================
# 13. KULLANICI REHBERİ VE YÖNETİCİ KADROSU MASASI (ADMIN MANAGEMENT)
# ======================================================================

@router.callback_query(F.data.startswith("adm:users_hub:"))
async def cb_admin_users_hub(query: CallbackQuery):
    page = int(query.data.split(":")[2])
    per_page = 8
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        total_users = (await session.execute(select(func.count(User.telegram_id)))).scalar() or 0
        total_pages = max(1, (total_users + per_page - 1) // per_page)
        users_list = (await session.execute(select(User).order_by(User.created_at.desc()).offset(page * per_page).limit(per_page))).scalars().all()

        hub_title = {
            "tr": f"👥 *Okul Botu Kullanıcı Rehberi* (Sayfa {page + 1}/{total_pages}):\nToplam Kayıtlı: *{total_users}*\nDetay veya işlem için kullanıcıya dokunun:\n",
            "ru": f"👥 *Справочник пользователей* (Стр {page + 1}/{total_pages}):\nВсего: *{total_users}*\nНажмите на пользователя для действий:\n",
            "uz": f"👥 *Foydalanuvchilar ma'lumotnomasi* (Sahifa {page + 1}/{total_pages}):\nJami: *{total_users}*\nAmal bajarish uchun tanlang:\n",
            "en": f"👥 *User Directory* (Page {page + 1}/{total_pages}):\nTotal: *{total_users}*\nTap a user for actions:\n"
        }.get(lang, f"👥 *User Directory*\n")

        lines = [hub_title]
        buttons = []
        role_icons = {"admin": "👑", "teacher": "👨‍🏫", "parent": "👨‍👩‍👧‍👦", "student": "🎓", "guest": "👤"}

        for u in users_list:
            icon = role_icons.get(u.role, "👤")
            if u.is_blacklisted: icon = "🚫"
            elif u.admin_type == "temporary": icon = "⏱️"
            u_name = u.full_name or get_text("lbl_role_guest", lang)
            u_tag = f" (@{u.username})" if u.username else ""
            btn_txt = f"{icon} {u_name}{u_tag} [{u.role.upper()}]"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:user_card:{u.telegram_id}")])

        nav_row = []
        if page > 0: nav_row.append(InlineKeyboardButton(text=get_text("btn_prev", lang), callback_data=f"adm:users_hub:{page - 1}"))
        if (page + 1) * per_page < total_users: nav_row.append(InlineKeyboardButton(text=get_text("btn_next", lang), callback_data=f"adm:users_hub:{page + 1}"))
        if nav_row: buttons.append(nav_row)

        buttons.append([InlineKeyboardButton(text=get_text("btn_search_user", lang), callback_data="adm:search_user_init")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:search_user_init")
async def cb_admin_search_user_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    buttons = [[InlineKeyboardButton(text=get_text("btn_cancel_action", lang), callback_data="adm:users_hub:0")]]
    await safe_edit_or_answer(query, get_text("search_user_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_search_user_query)
    await query.answer()

@router.message(Form.waiting_search_user_query)
async def process_search_user_query(message: Message, state: FSMContext):
    q_txt = message.text.strip().replace("@", "")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, message.from_user.id): return

        conds = [
            func.lower(User.full_name).contains(q_txt.lower()),
            func.lower(User.username).contains(q_txt.lower())
        ]
        if q_txt.isdigit(): conds.append(User.telegram_id == int(q_txt))

        res = (await session.execute(select(User).where(or_(*conds)).limit(10))).scalars().all()
        if not res:
            buttons = [
                [InlineKeyboardButton(text=get_text("btn_search_again", lang), callback_data="adm:search_user_init")],
                [InlineKeyboardButton(text=get_text("btn_users_list", lang), callback_data="adm:users_hub:0")]
            ]
            await message.answer(get_text("search_user_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        buttons = []
        for u in res:
            tag = f" (@{u.username})" if u.username else ""
            buttons.append([InlineKeyboardButton(text=f"👤 {u.full_name or get_text('lbl_role_guest', lang)}{tag} [{u.role.upper()}]", callback_data=f"adm:user_card:{u.telegram_id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_users_list", lang), callback_data="adm:users_hub:0")])
        await safe_edit_or_answer(message, get_text("search_user_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data.startswith("adm:user_card:"))
async def cb_admin_user_card(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id): return

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
            if tch: extra_info.append(f"• {lbl_s}: *{escape_md(tch.subject)}* (Kod: `{tch.auth_code}`)")
        elif target_u.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == tg_id))).scalar_one_or_none()
            if st: extra_info.append(f"• {lbl_cl}: *{escape_md(st.class_name)}* | {lbl_no}: *{escape_md(st.student_number)}*")
        elif target_u.role == "parent":
            kids = (await session.execute(select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == tg_id))).scalars().all()
            if kids:
                k_names = ", ".join([f"{k.full_name} ({k.class_name})" for k in kids])
                extra_info.append(f"• {lbl_lk}: *{escape_md(k_names)}*")

        lbl_none = {"uz": "Yo'q", "ru": "Нет", "en": "None", "tr": "Tanımlanmamış"}.get(lang, "None")
        lbl_nophone = {"uz": "Kiritilmagan", "ru": "Не указан", "en": "Not provided", "tr": "Kayıtlı Değil"}.get(lang, "Not provided")
        phone_str = target_u.phone if target_u.phone else lbl_nophone
        username_display = f"@{target_u.username}" if target_u.username else lbl_none

        if tg_id in PERMANENT_ADMIN_IDS or target_u.admin_type == "permanent":
            adm_status = {
                "uz": "👑 *Doimiy / Asosiy ma'mur*",
                "ru": "👑 *Постоянный / Главный администратор*",
                "en": "👑 *Permanent / Founder Administrator*",
                "tr": "👑 *Kalıcı / Kurucu Yönetici*"
            }.get(lang, "👑 *Permanent Administrator*")
        elif target_u.admin_type == "temporary" and target_u.admin_until:
            exp_str = (target_u.admin_until + timedelta(hours=TIMEZONE_OFFSET)).strftime('%d.%m.%Y %H:%M')
            adm_status = {
                "uz": f"⏱️ *Vaqtinchalik ma'mur* (Muddati: `{exp_str}`)",
                "ru": f"⏱️ *Временный администратор* (До: `{exp_str}`)",
                "en": f"⏱️ *Temporary Administrator* (Expires: `{exp_str}`)",
                "tr": f"⏱️ *Geçici Yönetici* (Bitiş: `{exp_str}`)"
            }.get(lang, f"⏱️ *Temporary Administrator* (Expires: `{exp_str}`)")
        else:
            adm_status = get_text("lbl_not_admin", lang)

        acc_status = get_text("lbl_status_banned", lang) if target_u.is_blacklisted else get_text("lbl_status_active", lang)

        card_title = get_text("uc_card_title", lang)
        lbl_fn = get_text("lbl_full_name", lang)
        lbl_un = get_text("lbl_username", lang)
        lbl_ph = get_text("lbl_phone", lang)
        lbl_r = get_text("lbl_role", lang)
        lbl_as = get_text("lbl_admin_status", lang)
        lbl_acc = get_text("lbl_account_status", lang)
        lbl_l = get_text("lbl_lang", lang)

        lines = [
            f"{card_title}\n",
            f"• Telegram ID: `{target_u.telegram_id}`",
            f"• {lbl_fn}: *{escape_md(target_u.full_name or lbl_none)}*",
            f"• {lbl_un}: {username_display}",
            f"• {lbl_ph}: `{phone_str}`",
            f"• {lbl_r}: *{role_label}*",
            f"• {lbl_as}: {adm_status}",
            f"• {lbl_acc}: {acc_status}",
            f"• {lbl_l}: {target_u.language.upper()}"
        ]
        if extra_info: lines.extend(extra_info)

        buttons = []
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
            if target_u.is_blacklisted:
                buttons.append([InlineKeyboardButton(text=get_text("btn_unban_user", lang), callback_data=f"adm:unban_u:{tg_id}")])
            else:
                buttons.append([InlineKeyboardButton(text=get_text("btn_ban_user", lang), callback_data=f"adm:ban_u:{tg_id}")])

        buttons.append([
            InlineKeyboardButton(text=get_text("btn_refresh_data", lang), callback_data=f"adm:user_card:{tg_id}"),
            InlineKeyboardButton(text=get_text("btn_users_list", lang), callback_data="adm:users_hub:0")
        ])

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:make_perm:"))
async def cb_admin_make_perm_admin(query: CallbackQuery):
    tg_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_u = await session.get(User, query.from_user.id)
        lang = admin_u.language if admin_u else "tr"
        if not is_admin_user(admin_u, query.from_user.id): return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.previous_role = target_u.role if target_u.role != "admin" else "guest"
            target_u.role = "admin"
            target_u.admin_type = "permanent"
            target_u.admin_until = None
            await session.commit()

            t_lang = target_u.language
            notify_msg = get_text("admin_promoted_notification", t_lang, name=escape_md(target_u.full_name or get_text("lbl_role_guest", t_lang)))
            await safe_send_message(query.message.bot, tg_id, notify_msg, reply_markup=get_role_reply_kb("admin", t_lang), parse_mode="Markdown")
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
    await safe_edit_or_answer(query, get_text("temp_admin_choose_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        if not is_admin_user(admin_u, query.from_user.id): return

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
            await safe_send_message(query.message.bot, tg_id, notify_msg, reply_markup=get_role_reply_kb("admin", t_lang), parse_mode="Markdown")
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
        if not is_admin_user(admin_u, query.from_user.id): return

        target_u = await session.get(User, tg_id)
        if target_u:
            target_u.role = target_u.previous_role or "guest"
            target_u.admin_type = "none"
            target_u.admin_until = None
            await session.commit()
            t_lang = target_u.language
            await safe_send_message(query.message.bot, tg_id, get_text("admin_demoted_notification", t_lang), reply_markup=get_role_reply_kb(target_u.role, t_lang), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, tg_id, get_text("auth_blacklisted", target_u.language if target_u else "tr"), parse_mode="Markdown")

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
            await safe_send_message(query.message.bot, tg_id, get_text("admin_unban_notification", target_u.language if target_u else "tr"), parse_mode="Markdown")

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
    await safe_edit_or_answer(query, get_text("send_dm_prompt", lang, id=tg_id, name=escape_md(target_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_admin_dm_text)
    await query.answer()

@router.message(Form.waiting_admin_dm_text)
async def process_admin_dm_text(message: Message, state: FSMContext):
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
        res = await safe_send_message(message.bot, target_id, formatted_msg, parse_mode="Markdown")
        if res:
            await message.answer(get_text("dm_sent_success", adm_lang, id=target_id), reply_markup=get_role_reply_kb("admin", adm_lang), parse_mode="Markdown")
        else:
            await message.answer(get_text("dm_delivery_error", adm_lang), reply_markup=get_role_reply_kb("admin", adm_lang), parse_mode="Markdown")

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
        res = await safe_send_message(query.message.bot, tg_id, chat_msg, reply_markup=kb, parse_mode="Markdown")
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
        if not is_admin_user(user, query.from_user.id): return

        db_admins = (await session.execute(select(User).where(User.role == "admin"))).scalars().all()
        admin_map = {u.telegram_id: u for u in db_admins}

        founder_label = {
            "tr": "Kalıcı Kurucu", "ru": "Постоянный админ", "uz": "Doimiy / Asosiy", "en": "Permanent Founder"
        }.get(lang, "Permanent Founder")

        lines = [get_text("admin_admins_hub_title", lang) + "\n"]
        buttons = []

        # 1. Kalıcı Kurucular
        for p_id in PERMANENT_ADMIN_IDS:
            u_obj = admin_map.get(p_id)
            if not u_obj:
                u_obj = await session.get(User, p_id)
            
            # Real Telegram name lookup if not in DB or is generic placeholder
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
                except Exception:
                    pass

            name_str = u_obj.full_name if (u_obj and u_obj.full_name) else get_text("permanent_admin_title", lang)
            lines.append(f"• 👑 *{escape_md(name_str)}* (`{p_id}`) - _{founder_label}_")
            buttons.append([InlineKeyboardButton(text=f"👑 {name_str} ({p_id})", callback_data=f"adm:user_card:{p_id}")])

        # 2. Eklenen Yöneticiler
        all_other_admin_ids = set(ADMIN_IDS) - set(PERMANENT_ADMIN_IDS)
        for extra_id in all_other_admin_ids:
            if extra_id not in admin_map:
                u_extra = await session.get(User, extra_id)
                if not u_extra or not u_extra.full_name or u_extra.full_name in ("Kalıcı İdareci", "Kurucu İdareci", "Asosiy ma'mur", "Главный админ", "Yönetici", "İsimsiz"):
                    try:
                        chat_info = await query.bot.get_chat(extra_id)
                        tg_name = chat_info.full_name or chat_info.title or chat_info.first_name
                        if tg_name:
                            if not u_extra:
                                u_extra = User(telegram_id=extra_id, role="admin", full_name=tg_name, admin_type="permanent", language=lang)
                                session.add(u_extra)
                            else:
                                u_extra.full_name = tg_name
                            await session.commit()
                    except Exception:
                        pass
                if u_extra:
                    admin_map[extra_id] = u_extra

        for tg_id, u_obj in admin_map.items():
            if tg_id not in PERMANENT_ADMIN_IDS:
                name_str = u_obj.full_name if (u_obj and u_obj.full_name) else f"ID: {tg_id}"
                if u_obj.admin_type == "temporary" and u_obj.admin_until:
                    exp_txt = (u_obj.admin_until + timedelta(hours=TIMEZONE_OFFSET)).strftime('%d.%m %H:%M')
                    temp_label = {
                        "tr": f"⏱️ Geçici ({exp_txt})", "ru": f"⏱️ Временный ({exp_txt})", "uz": f"⏱️ Vaqtinchalik ({exp_txt})", "en": f"⏱️ Temp ({exp_txt})"
                    }.get(lang, f"⏱️ Temp ({exp_txt})")
                    lines.append(f"• 👨‍💼 *{escape_md(name_str)}* (`{tg_id}`) - _{temp_label}_")
                else:
                    perm_label = {
                        "tr": "👑 Kalıcı", "ru": "👑 Постоянный", "uz": "👑 Doimiy", "en": "👑 Permanent"
                    }.get(lang, "👑 Permanent")
                    lines.append(f"• 👨‍💼 *{escape_md(name_str)}* (`{tg_id}`) - _{perm_label}_")
                buttons.append([InlineKeyboardButton(text=f"👨‍💼 {name_str} ({tg_id})", callback_data=f"adm:user_card:{tg_id}")])

        buttons.append([InlineKeyboardButton(text=get_text("btn_add_admin_id", lang), callback_data="adm:add_admin_id")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_gen_admin_code", lang), callback_data="adm:gen_admin_code")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:cat_staff"))

        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:gen_admin_code")
async def cb_admin_gen_code(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        code_val = generate_secure_code("ADM")
        adm_key = AdminKey(code=code_val, created_by=query.from_user.id)
        session.add(adm_key)
        await session.commit()

        text = get_text("admin_code_generated", lang, code=code_val)
        buttons = [get_nav_buttons(lang, back_callback="adm:admins_list")]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:add_admin_id")
async def cb_admin_add_id_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await query.message.answer(get_text("admin_add_tg_id_prompt", lang), reply_markup=get_cancel_reply_kb(lang), parse_mode="Markdown")
    await state.set_state(Form.waiting_admin_tg_id)
    await query.answer()

@router.message(Form.waiting_admin_tg_id)
async def process_admin_tg_id(message: Message, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    val = message.text.strip().replace("@", "")
    if not val.isdigit():
        await message.answer(get_text("admin_invalid_tg_id", lang))
        return
    await state.update_data(new_admin_id=int(val))
    await message.answer(get_text("admin_add_name_prompt", lang), parse_mode="Markdown")
    await state.set_state(Form.waiting_admin_name)

@router.message(Form.waiting_admin_name)
async def process_admin_name(message: Message, state: FSMContext):
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
        await safe_send_message(message.bot, new_admin_id, get_text("admin_promoted_notification", t_lang, name=escape_md(name_val)), reply_markup=get_role_reply_kb("admin", t_lang), parse_mode="Markdown")
        await message.answer(get_text("admin_added_success", lang, name=escape_md(name_val), id=new_admin_id), reply_markup=get_role_reply_kb("admin", lang), parse_mode="Markdown")
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
            await safe_send_message(query.message.bot, target_id, get_text("admin_demoted_notification", target_u.language if target_u else "tr"), reply_markup=get_role_reply_kb("guest", target_u.language), parse_mode="Markdown")

        await query.answer(get_text("admin_demoted_toast", lang), show_alert=True)
        await cb_admin_admins_list(query)

# ======================================================================
# 14. ÖĞRETMEN, VELİ VE ÖĞRENCİ İŞLEMLERİ (NOT, ÖDEV, KARNE, DEVAMSIZLIK)
# ======================================================================

@router.callback_query(F.data == "tch:classes")
async def cb_teacher_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes: classes = ["9-A"]

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"att_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row: buttons.append(row)
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
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
        text = get_text("attendance_intro", lang, class_name=escape_md(class_name))
        await safe_edit_or_answer(query, text, reply_markup=kb, parse_mode="Markdown")
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
    class_name = query.data.split(":")[1]
    now_local = get_local_now()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        if now_local.weekday() in (5, 6):
            await query.answer(get_text("attendance_weekend_lock", lang), show_alert=True)
            return

        if not (7 <= now_local.hour < 19):
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
                        await safe_send_message(query.message.bot, p_id, corr_text, parse_mode="Markdown")
            else:
                att = Attendance(student_id=st_id, class_name=class_name, date=att_date, status=new_status, teacher_id=query.from_user.id, notify_at=notify_time, is_notified=False)
                session.add(att)

        await session.commit()
        buttons = [get_nav_buttons(lang)]
        await safe_edit_or_answer(query, get_text("att_saved", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "tch:grade_classes")
async def cb_grade_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes: classes = ["9-A"]

        buttons = [[InlineKeyboardButton(text=get_text("btn_recent_grades_menu", lang), callback_data="tch:recent_grades")]]
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"gr_cls:{c}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("grade_select_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            "tr": f"📊 *{escape_md(class_name)} Sınıfı {escape_md(subject_name)} Not Çizelgesi:*\n",
            "ru": f"📊 *Ведомость оценок класса {escape_md(class_name)} ({escape_md(subject_name)}):*\n",
            "uz": f"📊 *{escape_md(class_name)} sinfining {escape_md(subject_name)} fanidan baholar qaydnomasi:*\n",
            "en": f"📊 *Grade Sheet for Class {escape_md(class_name)} ({escape_md(subject_name)}):*\n"
        }.get(lang, f"📊 *Grade Sheet: {escape_md(class_name)}*\n")

        no_gr_str = {"tr": "_Not girilmedi_", "ru": "_Оценок нет_", "uz": "_Baho qo'yilmagan_", "en": "_No grades recorded_"}.get(lang, "_No grades_")
        lbl_avg_cls = {"tr": "Sınıf Dersi Ortalaması", "ru": "Средний балл по предмету", "uz": "Sinf fan o'rtachasi", "en": "Class Subject Average"}.get(lang, "Average")
        lbl_cnt = {"tr": "not", "ru": "оценок", "uz": "baho", "en": "grades"}.get(lang, "grades")

        lines = [gr_sheet_title]
        scores_all = []
        for s in students:
            grades = (await session.execute(select(Grade).where(Grade.student_id == s.id, Grade.subject == subject_name).order_by(Grade.created_at))).scalars().all()
            if grades:
                g_str = ", ".join([f"{g.exam_type or '1. Yazılı'}: *{g.score}*" for g in grades])
                for g in grades: scores_all.append(g.score)
            else:
                g_str = no_gr_str
            lines.append(f"• *{escape_md(s.full_name)}* (№{s.student_number}): {g_str}")

        avg_class = round(sum(scores_all)/len(scores_all), 1) if scores_all else 0
        lines.append(f"\n📈 *{lbl_avg_cls}:* {avg_class} ({len(scores_all)} {lbl_cnt})")

        buttons = [get_nav_buttons(lang, back_callback=f"gr_cls:{class_name}")]
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await safe_edit_or_answer(query, get_text("grade_select_student", lang, class_name=escape_md(class_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
    await safe_edit_or_answer(query, prompt, reply_markup=kb, parse_mode="Markdown")
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
    await safe_edit_or_answer(query, text, parse_mode="Markdown")
    await state.set_state(Form.grade_score)
    await query.answer()

@router.message(Form.grade_score)
async def process_grade_score(message: Message, state: FSMContext):
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
        existing_grade = (await session.execute(select(Grade).where(Grade.student_id == st_id, Grade.subject == subject_name, Grade.exam_type == exam_type_val))).scalar_one_or_none()

        if existing_grade:
            existing_grade.score = score_val
            existing_grade.badge = badge_val
            existing_grade.created_at = datetime.utcnow()
        else:
            grade = Grade(student_id=st_id, subject=subject_name, exam_type=exam_type_val, score=score_val, badge=badge_val, note="Öğretmen Değerlendirmesi", teacher_id=query.from_user.id)
            session.add(grade)
        await session.commit()

        parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
        for p_id in parents:
            p_u = await session.get(User, p_id)
            p_lang = p_u.language if p_u else "tr"
            alert_text = get_text("grade_parent_notification", p_lang, name=escape_md(st.full_name), subject=escape_md(subject_name), exam_type=escape_md(exam_type_val), score=score_val, badge=badge_val)
            await safe_send_message(query.message.bot, p_id, alert_text, parse_mode="Markdown")

        buttons = [get_nav_buttons(lang)]
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
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("recent_grades_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        text = f"📝 *{get_text('btn_enter_grade', lang)}*\n{lbl_st}: *{st_name}* ({st_cls})\n{lbl_s}: *{escape_md(grade.subject)}* ({escape_md(grade.exam_type or '1. Yazılı')})\nNot: *{grade.score}* ({grade.badge})\n{lbl_dt} {date_str}"
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_edit_grade", lang), callback_data=f"tch:edit_gr:{grade.id}")],
            [InlineKeyboardButton(text=get_text("btn_del_grade", lang), callback_data=f"tch:del_gr:{grade.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:recent_grades")]
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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

    cancel_kb = get_cancel_reply_kb(lang)
    await query.message.answer(get_text("prompt_new_score", lang), reply_markup=cancel_kb)
    await state.set_state(Form.edit_grade_val)
    await query.answer()

@router.message(Form.edit_grade_val)
async def process_grade_edit_val(message: Message, state: FSMContext):
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

@router.callback_query(F.data == "adm:excel_hub")
async def cb_admin_excel_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_upload_excel", lang), callback_data="adm:excel_info")],
            [InlineKeyboardButton(text=get_text("btn_export_all_data", lang), callback_data="adm:export_all_excel")],
            get_nav_buttons(lang)
        ]
        await safe_edit_or_answer(query, get_text("excel_hub_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.message(any_state, F.text.in_(["ℹ️ Okul Bilgi Panosu", "ℹ️ Инфопанель школы", "ℹ️ Maktab ma'lumotlari", "ℹ️ School Info Board"]))
async def cb_parent_info_board(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        briefing_txt = get_text("btn_briefing_on", lang) if (user and user.evening_briefing) else get_text("btn_briefing_off", lang)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_view_schedule", lang), callback_data="act_view_sched")],
            [InlineKeyboardButton(text=get_text("btn_notices", lang), callback_data="act_view_notices")],
            [InlineKeyboardButton(text=get_text("btn_view_cafeteria", lang), callback_data="act_view_cafe")],
            [InlineKeyboardButton(text=briefing_txt, callback_data="parent:toggle_briefing")]
        ]
        await safe_edit_or_answer(message, get_text("parent_info_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_(["⚙️ Ayarlar & Çıkış", "⚙️ Настройки и выход", "⚙️ Sozlamalar va chiqish", "⚙️ Settings & Exit"]))
async def cb_parent_settings_hub(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        briefing_txt = get_text("btn_briefing_on", lang) if (user and user.evening_briefing) else get_text("btn_briefing_off", lang)
        buttons = [
            [InlineKeyboardButton(text=briefing_txt, callback_data="parent:toggle_briefing")],
            [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")],
            [InlineKeyboardButton(text=get_text("rk_logout", lang), callback_data="act_logout")]
        ]
        await safe_edit_or_answer(message, get_text("parent_settings_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "act_view_hw")
async def cb_view_homeworks_list(query: CallbackQuery):
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

        hws = (await session.execute(select(Homework).where(Homework.class_name == cls_name).order_by(desc(Homework.created_at)).limit(6))).scalars().all()
        if not hws:
            buttons = [get_nav_buttons(lang)]
            await safe_edit_or_answer(query, get_text("no_active_homeworks", lang, class_name=escape_md(cls_name)), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        lines = [f"{get_text('homework_board_title', lang, class_name=escape_md(cls_name))}\n"]
        buttons = []
        for h in hws:
            lines.append(f"📌 *{escape_md(h.subject)}* ({h.created_at.strftime('%d.%m.%Y')}):\n{escape_md(h.content)}")
            if h.file_id:
                buttons.append([InlineKeyboardButton(text=f"📷 {get_text('btn_view_photo', lang)}: {h.subject}", callback_data=f"hw_photo:{h.id}")])
            lines.append("")

        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, "\n".join(lines), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
                await query.message.bot.send_photo(chat_id=query.from_user.id, photo=hw.file_id, caption=caption, parse_mode="Markdown")
                await query.answer()
                return
            except Exception: pass
    await query.answer(get_text("image_load_error", lang), show_alert=True)

@router.callback_query(F.data == "tch:hw_classes")
async def cb_hw_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes: classes = ["9-A"]

        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"📢 {c} - {get_text('btn_send_new_hw', lang)}", callback_data=f"hw_cls:{c}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_my_hws", lang), callback_data="tch:view_my_hws")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
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
            btn_txt = f"📚 {h.class_name} ({h.subject}): {h.content[:15]}..."
            buttons.append([
                InlineKeyboardButton(text=btn_txt, callback_data="noop"),
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
    hw_text = message.caption or message.text or "Ödev Panosu"

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        subj = tch.subject if tch else "Ders"

        hw = Homework(class_name=class_name, subject=subj, content=hw_text, file_id=photo_id, teacher_id=message.from_user.id)
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
            if photo_id:
                try: await message.bot.send_photo(chat_id=t_id, photo=photo_id, caption=caption, parse_mode="Markdown")
                except Exception: pass
            else:
                await safe_send_message(message.bot, t_id, caption, parse_mode="Markdown")

        await message.answer(get_text("hw_sent_success", lang, class_name=escape_md(class_name)), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

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
            buttons.append([InlineKeyboardButton(text=f"{is_active}🧑‍🎓 {c.full_name} ({c.class_name})", callback_data=f"set_child:{c.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("parent_choose_child", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data == "parent:add_child_code")
async def cb_parent_add_child_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    cancel_kb = get_cancel_reply_kb(lang)
    await query.message.answer(get_text("prompt_add_child_code", lang), reply_markup=cancel_kb, parse_mode="Markdown")
    await state.set_state(Form.parent_add_child_code)
    await query.answer()

@router.message(Form.parent_add_child_code)
async def process_parent_add_child_code(message: Message, state: FSMContext):
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
            await message.answer(get_text("invalid_parent_code", lang), reply_markup=reply_kb, parse_mode="Markdown")
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
        await message.answer(success_text, reply_markup=reply_kb, parse_mode="Markdown")
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

        grades = (await session.execute(select(Grade).where(Grade.student_id == st.id).order_by(Grade.created_at.desc()))).scalars().all()
        grades_text = "\n".join([f"• {escape_md(g.subject)} ({escape_md(g.exam_type or '1. Yazılı')}): *{g.score}* ({g.badge})" for g in grades]) if grades else get_text("no_grades", lang)

        lbl_tot_att = {"tr": "Toplam Devamsızlık:", "ru": "Всего пропусков:", "uz": "Umumiy davomat:", "en": "Total Absences:"}.get(lang, "Absences:")
        lbl_day = {"tr": "gün", "ru": "дн.", "uz": "kun", "en": "days"}.get(lang, "days")
        lbl_exc = {"tr": "Özürlü:", "ru": "Уважительно:", "uz": "Ruxsatli:", "en": "Excused:"}.get(lang, "Excused:")
        lbl_gr_sec = {"tr": "Ders Notları:", "ru": "Оценки по предметам:", "uz": "Fan baholari:", "en": "Grades:"}.get(lang, "Grades:")
        lbl_status_rep = {"tr": "Durum Raporu", "ru": "Отчет об успеваемости", "uz": "O'zlashtirish hisoboti", "en": "Progress Report"}.get(lang, "Report")

        text = (
            f"📊 *{escape_md(st.full_name)} ({escape_md(st.class_name)}) - {lbl_status_rep}*\n\n"
            f"📌 {lbl_tot_att} *{absent_count} {lbl_day}* ({lbl_exc} {excused_count} {lbl_day})\n\n"
            f"📝 *{lbl_gr_sec}*\n{grades_text}"
        )
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_download_pdf_report", lang), callback_data=f"parent:pdf_report:{st.id}")],
            get_nav_buttons(lang)
        ]
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
        await query.message.answer_document(file, caption=get_text("pdf_report_ready", lang, name=escape_md(st.full_name)), parse_mode="Markdown")
        pdf_buf.close()
    await query.answer()

@router.callback_query(F.data == "upload_medical_init")
async def cb_upload_med_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await safe_edit_or_answer(query, get_text("upload_med_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_medical_photo)
    await query.answer()

@router.message(Form.waiting_medical_photo, ~F.photo)
async def fallback_medical_photo_text(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
    await message.answer(get_text("photo_expected_medical", lang), parse_mode="Markdown")

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
        caption_adm = f"🏥 *YENİ SAĞLIK / MAZERET RAPORU* (#{report.id})\n\n🧑‍🎓 *Öğrenci:* {escape_md(st.full_name)} ({escape_md(st.class_name)})\n👤 *Veli:* {escape_md(user.full_name or 'Veli')}"
        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=get_text("btn_appr_medical", "tr"), callback_data=f"adm:appr_med:{report.id}"), InlineKeyboardButton(text=get_text("btn_reject", "tr"), callback_data=f"adm:rej_med:{report.id}")]
        ])
        for a_id in set(ADMIN_IDS + list(admins)):
            try:
                await message.bot.send_photo(chat_id=a_id, photo=photo_file_id, caption=caption_adm, reply_markup=adm_kb, parse_mode="Markdown")
                await asyncio.sleep(0.05)
            except Exception: pass
    await state.clear()

@router.callback_query(F.data.startswith("ack_notif:"))
async def cb_ack_notification(query: CallbackQuery):
    notif_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        notif = await session.get(CriticalNotification, notif_id)
        if notif:
            notif.acknowledged_at = datetime.utcnow()
            await session.commit()
    await query.answer(get_text("acknowledged_toast", lang), show_alert=True)
    await query.message.edit_reply_markup(reply_markup=None)

@router.callback_query(F.data == "act_logout")
async def cb_act_logout(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        u = await session.get(User, query.from_user.id)
        if u:
            u.role = "guest"
            u.current_child_id = None
            u.failed_attempts = 0
            u.locked_until = None
            await session.commit()
    try: await query.message.delete()
    except Exception: pass
    await query.message.answer("🌍 Iltimos, tilni tanlang / Пожалуйста, выберите язык / Lütfen bir dil seçiniz / Select language:", reply_markup=get_language_inline_kb())

@router.message(any_state, Command("restart", "reset", "cikis", "logout"))
@router.message(any_state, F.text.in_(["🔄 Yeniden Başlat", "🔄 Перезапуск", "🔄 Qayta ishga tushirish", "🔄 Restart Bot", "🚪 Çıkış Yap", "🚪 Выйти", "🚪 Chiqish", "🚪 Log Out", "/restart", "/reset", "/cikis", "/logout"]))
async def handle_bot_restart_cmd(message: Message, state: FSMContext):
    await state.clear()
    user_id = message.from_user.id

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        is_rec_admin = (user_id in ADMIN_IDS) or (user and user.admin_type == "permanent") or (user and user.role == "admin" and (user.admin_until is None or user.admin_until > datetime.utcnow()))

        if is_rec_admin:
            if user:
                user.role = "admin"
                user.admin_type = "permanent" if (user_id in ADMIN_IDS or user.admin_type == "permanent") else "temporary"
                user.failed_attempts = 0
                user.locked_until = None
                if message.from_user.full_name:
                    user.full_name = message.from_user.full_name
                await session.commit()
                lang = user.language
            else:
                user = User(telegram_id=user_id, role="admin", language="tr", admin_type="permanent", full_name=message.from_user.full_name or get_text("permanent_admin_title", "tr"))
                session.add(user)
                await session.commit()
                lang = "tr"

            try: await message.delete()
            except Exception: pass
            await message.answer(get_text("admin_restart_confirmed", lang), reply_markup=get_role_reply_kb("admin", lang), parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        if user:
            user.role = "guest"
            user.current_child_id = None
            user.failed_attempts = 0
            user.locked_until = None
            if message.from_user.full_name:
                user.full_name = message.from_user.full_name
            await session.commit()
        else:
            user = User(telegram_id=user_id, role="guest", language="tr", full_name=message.from_user.full_name)
            session.add(user)
            await session.commit()

    try: await message.delete()
    except Exception: pass
    await message.answer("🌍 Iltimos, tilni tanlang / Пожалуйста, выберите язык / Lütfen bir dil seçiniz / Select language:", reply_markup=get_language_inline_kb())

@router.callback_query(F.data == "act_change_lang")
async def cb_change_lang_screen(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
    await safe_edit_or_answer(query, get_text("lang_select", lang), reply_markup=get_language_inline_kb())
    await query.answer()

# ======================================================================
# 15. SABİT ALT MENÜ (REPLY KEYBOARD) VE DOĞRUDAN KOD YAZMA MOTORU
# ======================================================================

REPLY_BUTTON_ACTIONS = {
    "rk_logout": "act_restart",
    "rk_restart": "act_restart",
    "rk_main_menu": "act_main_menu",
    "rk_admin_dash": "act_main_menu",
    "rk_lang": "act_lang",
    "rk_enter_code": "act_enter_code",
    "rk_req_access": "act_req_access",
    "rk_cat_staff": "act_cat_staff",
    "rk_cat_reports": "act_cat_reports",
    "rk_cat_requests": "act_cat_requests",
    "rk_cat_tools": "act_cat_tools",
    "rk_cat_settings": "act_cat_settings",
    "rk_cockpit": "act_cockpit",
    "rk_classes": "act_classes",
    "rk_teachers": "act_teachers",
    "rk_requests": "act_requests",
    "rk_attendance": "act_attendance",
    "rk_grade": "act_grade",
    "rk_homework": "act_homework",
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
    return None

@router.message(any_state, F.text.func(lambda text: match_reply_button(text) is not None))
async def global_reply_keyboard_router(message: Message, state: FSMContext):
    action = match_reply_button(message.text)
    user_id = message.from_user.id

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
        if message.from_user.full_name:
            user.full_name = message.from_user.full_name
            await session.commit()
        lang = user.language

    if action:
        await state.clear()
        if action in ("act_logout", "act_restart"):
            await handle_bot_restart_cmd(message, state)
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
            await message.answer(get_text("lang_select", lang), reply_markup=get_language_inline_kb())
            return
        elif action == "act_enter_code":
            await message.answer(get_text("prompt_enter_code_direct", lang), parse_mode="Markdown")
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

@router.message(any_state, F.text.func(lambda text: normalize_code(text).startswith(("VELI-", "OGR-", "HCA-", "ADM-")) or normalize_code(text) == ADMIN_CODE))
@router.message(F.text)
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

# ======================================================================
# 16. FASTAPI VE ARKA PLAN DÖNGÜLERİ (LIFESPAN & PINGER)
# ======================================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
dp.include_router(router)

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
    except Exception as e:
        print(f"--> [HATA 2/6] BOT_TOKEN ile Telegram'a bağlanılamadı: {e}")

    if bot_info and WEBHOOK_URL:
        try:
            print(f"--> [3/6] Webhook Telegram'a kaydediliyor: {WEBHOOK_URL}")
            await bot.delete_webhook(drop_pending_updates=False)
            await bot.set_webhook(
                url=WEBHOOK_URL,
                secret_token=WEBHOOK_SECRET,
                drop_pending_updates=False,
                allowed_updates=["message", "callback_query"]
            )
            wh = await bot.get_webhook_info()
            print(f"--> [3/6] Webhook Başarıyla Kuruldu! Aktif URL: {wh.url}")
        except Exception as e:
            print(f"--> [HATA 3/6] Webhook kurulum hatası: {e}")

    t1 = asyncio.create_task(background_attendance_loop())
    t2 = asyncio.create_task(background_keep_alive_pinger())
    t3 = asyncio.create_task(background_evening_briefing_loop())
    t4 = asyncio.create_task(background_morning_briefing_loop())
    t5 = asyncio.create_task(background_friday_backup_loop())
    print("--> [4/6] Arka plan işçileri aktif.")
    print("--> [5/6] Cuma 18:00 Otomatik Yönetici Excel Veri Yedeği aktif.")
    print("--> [6/6] SİSTEM CANLI VE TÜM GÜVENLİK KALKANLARI HAZIR.")
    print("=" * 60)
    yield
    t1.cancel()
    t2.cancel()
    t3.cancel()
    t4.cancel()
    t5.cancel()
    try: await bot.session.close()
    except Exception: pass

app = FastAPI(title="OkulYonetimBot", lifespan=lifespan)

@app.get("/")
async def root():
    return {"status": "ok", "service": "OkulYonetimBot", "version": "22.0-enterprise", "uptime": True}

@app.get("/health", status_code=status.HTTP_200_OK)
async def health_check():
    return {"status": "ok", "service": "OkulYonetimBot", "uptime": True}

@app.post("/webhook")
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
