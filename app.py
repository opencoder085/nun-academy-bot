# -*- coding: utf-8 -*-
"""
OKUL YÖNETİM TELEGRAM BOTU - KURUMSAL VE EKSİKSİZ TAM SİSTEM MİMARİSİ (PROD V9 - ENTERPRISE)
Altyapı: FastAPI + aiogram 3.x Webhook + PostgreSQL (Neon / Supabase) / SQLite Çift Motor Kalkanı
4 Dilli Kusursuz Arayüz (i18n): 🇹🇷 Türkçe, 🇷🇺 Русский, 🇺🇿 O'zbekcha, 🇬🇧 English
Platform: Render Web Service / VPS / Docker

Yeni Mimari Özellikler:
- 📩 Yöneticiden Şifre / Erişim İsteme Modülü (Kimlik seçimi, akıllı veri eşleşme teyidi, spam & cooldown kalkanı).
- 🛎️ Yönetici Tek Tuşla Onay / Ret Masası (Otomatik rol aktivasyonu ve Telegram üzerinden anlık bildirim).
- 🤝 Veli - Öğretmen Randevu Sistemi (Görüşme talepleri ve öğretmen onay paneli).
- 📅 Sınıf Ders Programı & 🍲 Günlük Yemekhane Menüsü Entegrasyonu.
- 📱 Sade, Dağınık Olmayan, Maksimum Kolay Tuş Tasarımı (Rol bazlı odaklı Reply & Inline Hub'lar).
- 🔒 Güçlendirilmiş Güvenlik ve İzolasyon: Kaba kuvvet kalkanı, 6 haneli benzersiz şifre üretimi.
"""

import os
import io
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
    BigInteger, Integer, String, Boolean, DateTime, Date, ForeignKey, Float, Text, select, func, delete, desc
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from sqlalchemy.ext.asyncio import AsyncSession, create_async_engine, async_sessionmaker

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

def get_local_now() -> datetime:
    """Okulun yerel saatini döner (Varsayılan UTC+3)"""
    return datetime.utcnow() + timedelta(hours=TIMEZONE_OFFSET)

def get_local_date() -> date:
    """Okulun yerel takvim tarihini döner"""
    return get_local_now().date()

async def safe_edit_or_answer(target: Message | CallbackQuery, text: str, reply_markup=None, parse_mode="Markdown"):
    """
    Kullanıcının gönderdiği mesajlarda 'message can't be edited' hatasını önleyen,
    inline butonlarda yerinde güncelleme, metinlerde ise temiz yeni mesaj gönderen akıllı fonksiyon.
    """
    if isinstance(target, CallbackQuery):
        msg = target.message
        if msg and msg.from_user and msg.from_user.is_bot and not msg.photo:
            try:
                await msg.edit_text(text, reply_markup=reply_markup, parse_mode=parse_mode)
                return
            except Exception:
                pass
        await msg.answer(text, reply_markup=reply_markup, parse_mode=parse_mode)
    elif isinstance(target, Message):
        await target.answer(text, reply_markup=reply_markup, parse_mode=parse_mode)

# ======================================================================
# 1. ORTAM DEĞİŞKENLERİ VE AKILLI WEBHOOK TESPİTİ
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

# Doğrudan Tanımlı Kalıcı Yöneticiler (Kullanıcı tarafından belirlenen ID'ler)
PERMANENT_ADMIN_IDS = [8576061834, 2146753102, 1885043735]

ADMIN_CODE = os.getenv("ADMIN_CODE", "ADM-OKUL-2026").strip()
TIMEZONE_OFFSET = int(os.getenv("TIMEZONE_OFFSET", "3")) # Varsayılan UTC+3 (Türkiye)

ADMIN_IDS = list(PERMANENT_ADMIN_IDS)
for x in os.getenv("ADMIN_IDS", "").split(","):
    clean_x = x.strip().replace("@", "")
    if clean_x.isdigit():
        val = int(clean_x)
        if val not in ADMIN_IDS:
            ADMIN_IDS.append(val)

# ======================================================================
# 2. VERİTABANI MOTORU VE MODELLERİ (SQLITE & POSTGRESQL ÇİFT MOTOR)
# ======================================================================

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://") and not DATABASE_URL.startswith("postgresql+asyncpg://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)

from sqlalchemy import event

if "sqlite" in DATABASE_URL:
    engine = create_async_engine(
        DATABASE_URL,
        connect_args={"check_same_thread": False, "timeout": 30},
        pool_pre_ping=True,
        echo=False
    )
    @event.listens_for(engine.sync_engine, "connect")
    def set_sqlite_pragma(dbapi_connection, connection_record):
        cursor = dbapi_connection.cursor()
        cursor.execute("PRAGMA journal_mode=WAL")
        cursor.execute("PRAGMA synchronous=NORMAL")
        cursor.execute("PRAGMA busy_timeout=30000")
        cursor.execute("PRAGMA cache_size=-64000")
        cursor.close()
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

class User(Base):
    __tablename__ = "users"
    telegram_id: Mapped[int] = mapped_column(BigInteger, primary_key=True, index=True)
    role: Mapped[str] = mapped_column(String(20), default="guest") # admin, teacher, parent, student, guest
    language: Mapped[str] = mapped_column(String(5), default="tr")
    full_name: Mapped[str] = mapped_column(String(100), nullable=True)
    failed_attempts: Mapped[int] = mapped_column(Integer, default=0)
    locked_until: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    is_blacklisted: Mapped[bool] = mapped_column(Boolean, default=False)
    current_child_id: Mapped[int] = mapped_column(Integer, nullable=True)
    evening_briefing: Mapped[bool] = mapped_column(Boolean, default=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

class Grade(Base):
    __tablename__ = "grades"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    subject: Mapped[str] = mapped_column(String(50))
    score: Mapped[float] = mapped_column(Float)
    badge: Mapped[str] = mapped_column(String(10)) # 🟢, 🟡, 🔴
    note: Mapped[str] = mapped_column(Text, nullable=True)
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

class MedicalReport(Base):
    __tablename__ = "medical_reports"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    student_id: Mapped[int] = mapped_column(Integer, ForeignKey("students.id"), index=True)
    parent_telegram_id: Mapped[int] = mapped_column(BigInteger)
    file_id: Mapped[str] = mapped_column(String(255))
    caption: Mapped[str] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending") # pending, approved, rejected
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

class CriticalNotification(Base):
    __tablename__ = "critical_notifications"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    user_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    message_text: Mapped[str] = mapped_column(Text)
    acknowledged_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

class Homework(Base):
    __tablename__ = "homeworks"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    class_name: Mapped[str] = mapped_column(String(20), index=True)
    subject: Mapped[str] = mapped_column(String(50))
    content: Mapped[str] = mapped_column(Text)
    file_id: Mapped[str] = mapped_column(String(255), nullable=True)
    teacher_id: Mapped[int] = mapped_column(BigInteger)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

class SystemSetting(Base):
    __tablename__ = "system_settings"
    key: Mapped[str] = mapped_column(String(50), primary_key=True)
    value: Mapped[str] = mapped_column(String(255))

class AccessRequest(Base):
    __tablename__ = "access_requests"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    role: Mapped[str] = mapped_column(String(20)) # teacher, parent, student
    full_name: Mapped[str] = mapped_column(String(100))
    phone: Mapped[str] = mapped_column(String(30))
    details: Mapped[str] = mapped_column(Text)
    student_match_id: Mapped[int] = mapped_column(Integer, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending") # pending, approved, rejected
    reviewed_by: Mapped[int] = mapped_column(BigInteger, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())
    reviewed_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)

class Appointment(Base):
    __tablename__ = "appointments"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    parent_telegram_id: Mapped[int] = mapped_column(BigInteger, index=True)
    teacher_id: Mapped[int] = mapped_column(Integer, ForeignKey("teachers.id"), index=True)
    preferred_time: Mapped[str] = mapped_column(String(100))
    note: Mapped[str] = mapped_column(Text, nullable=True)
    status: Mapped[str] = mapped_column(String(20), default="pending") # pending, approved, rejected
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

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
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

class BroadcastNotice(Base):
    __tablename__ = "broadcast_notices"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    content: Mapped[str] = mapped_column(Text)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=lambda: datetime.utcnow())

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
            except Exception:
                pass
        await conn.run_sync(Base.metadata.create_all)

# ======================================================================
# 3. KUSURSUZ 4 DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {
    "tr": {
        "lang_select": "🌍 Lütfen bir dil seçiniz / Tilni tanlang / Select language:",
        "lang_changed": "Dil başarıyla güncellendi: 🇹🇷 Türkçe",
        "prompt_enter_code_direct": "🔑 *Lütfen size verilen giriş kodunu yazınız:* (Örn: `HCA-123456`, `VELI-123456`, `OGR-123456`)",
        "welcome_guest": "🎓 *Okul Yönetim Sistemine Hoş Geldiniz.*\n\nLütfen giriş yapmak için bir seçenek belirleyiniz veya doğrudan size verilen **erişim kodunu** yazınız:",
        "auth_success": "✅ Giriş başarılı!\nHoş geldiniz: *{name}*\nRolünüz: *{role}*",
        "auth_failed": "❌ Geçersiz giriş kodu! Kalan deneme hakkınız: {remaining}",
        "auth_locked": "⛔ Güvenlik nedeniyle hesabınız 1 saat süreyle kilitlendi.",
        "auth_blacklisted": "🚫 Hesabınız güvenlik nedeniyle kalıcı olarak askıya alındı.",
        "maintenance_mode": "⚠️ Sistem şu anda planlı bakım modundadır. Lütfen daha sonra tekrar deneyiniz.",
        "admin_title": "⚡ *Okul Yönetim Kokpiti (Admin)*",
        "admin_stats": "📊 *Genel Durum:*\n• Sınıflar: *{c_cnt}* | Öğrenciler: *{s_cnt}* | Öğretmenler: *{t_cnt}*\n• Bekleyen Talepler: *{req_cnt}* | Mazeretler: *{med_cnt}*\n• Tarih: *{date}*",
        "btn_add_student": "➕ Öğrenci Ekle",
        "btn_add_teacher": "➕ Öğretmen Ekle",
        "btn_classes": "🏫 Sınıflar & Öğrenciler",
        "btn_search_student": "🔍 Öğrenci Ara",
        "btn_excel": "📥 Excel ile Yükle",
        "btn_cockpit": "📊 Sabah Kokpiti",
        "btn_pdf": "📄 Şifre Kartları (PDF)",
        "btn_medical": "🏥 Mazeretler ({count})",
        "btn_requests": "🛎️ Erişim Talepleri ({count})",
        "btn_broadcast": "📢 Toplu Duyuru",
        "btn_maintenance_toggle": "🚨 Bakım Modu ({status})",
        "btn_cafeteria_edit": "🍲 Günün Menüsünü Güncelle",
        "btn_lang": "🌐 Dil Değiştir",
        "btn_back": "⬅️ Geri",
        "btn_main_menu": "🏠 Ana Menü",
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
        "student_card": "👤 *Öğrenci Kartı*\nİsim: *{name}*\nSınıf: *{class_name}* | No: *{no}*\n\n🔑 *Kod Durumu:*\n• Öğrenci: `{st_code}` ({st_status})\n• Veli: `{pr_code}` ({pr_status})",
        "btn_reset_codes": "🔄 Kodları Sıfırla",
        "btn_del_student": "❌ Öğrenciyi Sil",
        "btn_student_grades": "📊 Notları & Devamsızlığı",
        "student_deleted": "🗑️ Öğrenci sistemden silindi.",
        "codes_reset_done": "✅ Giriş kodları yenilendi!\n\n• Yeni Öğrenci: `{st_code}`\n• Yeni Veli: `{pr_code}`",
        "excel_info": "📥 *Excel ile Toplu Öğrenci Yükleme*\n\nLütfen `.xlsx` dosyasını gönderiniz.\nBaşlıklar: `Ad Soyad` | `Sinif` | `Numara`",
        "excel_done": "✅ Excel işlendi! Eklenen öğrenci: *{count}*\nŞifreler ekteki dosyada üretildi.",
        "cockpit_report": "📊 *Sabah Kokpiti ({date})*\n\n🏫 Toplam: *{total}* | ✅ Gelen: *{present}* | ❌ Gelmeyen: *{absent}*\n\n⚠️ *Yoklama Almayan Sınıflar ({missing_cnt}):*\n{missing}",
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
        "menu_teacher": "👨‍🏫 *Öğretmen Masası*\nÖğretmen: *{name}* ({subject})",
        "btn_attendance": "📋 Hızlı Yoklama",
        "btn_enter_grade": "📝 Not Girişi",
        "btn_homework_board": "📢 Ödev Panosu",
        "btn_teacher_appointments": "🤝 Veli Randevuları ({count})",
        "attendance_intro": "📋 *{class_name} Yoklaması*\nGelmeyenlerin üzerine tıklayıp kaydediniz:",
        "btn_save_att": "💾 Yoklamayı Kaydet",
        "att_saved": "✅ Yoklama kaydedildi. 15 dakikalık düzeltme süresi başladı.",
        "prompt_grade_score": "Öğrenci: *{name}* ({class_name})\nNotu giriniz (0-100):",
        "prompt_grade_badge": "Değerlendirme rozeti seçiniz:",
        "badge_praise": "🟢 Başarılı / Övgü",
        "badge_missing": "🟡 Eksik / Tekrar",
        "badge_warning": "🔴 Uyarı / Dikkat",
        "grade_saved_success": "✅ Not kaydedildi ve veliye bildirildi.",
        "prompt_hw_class": "Ödevin sınıfını seçiniz:",
        "prompt_hw_content": "Ödev açıklamasını yazınız veya fotoğraf gönderiniz:",
        "hw_sent_success": "📢 Ödev *{class_name}* sınıfına iletildi.",
        "menu_parent": "👨‍👩‍👧‍👦 *Veli Masası*\nÖğrenci: *{name}* ({class_name})",
        "btn_report_card": "📊 Durum Paneli (Karne)",
        "btn_switch_student": "🧑‍🎓 Öğrenci Değiştir",
        "btn_upload_medical": "🏥 Mazeret / Rapor Yükle",
        "btn_book_appointment": "🤝 Öğretmen Randevusu",
        "btn_view_schedule": "📅 Ders Programı",
        "btn_view_cafeteria": "🍲 Günün Menüsü",
        "upload_med_prompt": "Lütfen mazeret veya rapor belgesinin fotoğrafını gönderiniz:",
        "med_uploaded_success": "Rapor okul idaresine iletildi.",
        "student_switched_success": "Aktif öğrenci seçildi: *{name}* ({class_name})",
        "menu_student": "🎓 *Öğrenci Masası*\nÖğrenci: *{name}* ({class_name} - No: {no})",
        "no_grades": "Henüz girilmiş bir ders notu bulunmamaktadır.",
        "no_linked_student": "⚠️ Hesabınıza tanımlı bir öğrenci bulunamadı.",
        "evening_briefing_header": "🌙 *GÜN SONU BÜLTENİ (18:30)*\nÖğrenci: *{name}* ({class_name})\n\n📌 Devamsızlık: *{att_status}*\n📝 Notlar:\n{grades}",
        "logged_out": "🚪 Güvenli çıkış yapıldı. Tekrar giriş yapabilirsiniz.",
        # Talep Sistemi (Request Access)
        "btn_req_access": "📩 Şifre / Erişim Talep Et",
        "req_role_select": "📩 *Erişim Talebi Başvurusu*\n\nLütfen kurumdaki rolünüzü seçiniz:",
        "req_name_prompt": "👤 Lütfen Adınızı ve Soyadınızı yazınız:",
        "req_phone_prompt": "📱 Lütfen iletişim telefon numaranızı yazınız (Örn: `05551234567`):",
        "req_details_teacher": "📚 Öğretmenlik branşınızı yazınız (Örn: `Matematik`):",
        "req_details_student": "🏫 Sınıfınızı ve Okul Numaranızı yazınız (Örn: `9-A / 101`):",
        "req_details_parent": "🧑‍🎓 Öğrencinizin Adı Soyadı, Sınıfı ve Okul Numarasını yazınız (Örn: `Ali Yılmaz, 9-A, 101`):",
        "req_sent_success": "✅ *Erişim Talebiniz İdareye İletildi!*\n\nBaşvurunuz incelendikten sonra onay ve erişim bildiriminiz Telegram üzerinden otomatik iletilecektir.",
        "req_already_pending": "⚠️ Zaten incelenmekte olan bir başvurunuz bulunmaktadır. Lütfen yönetici onayını bekleyiniz.",
        "req_cooldown": "⚠️ Önceki başvurunuz reddedilmiştir. Güvenlik gereği 24 saat sonra tekrar deneyebilirsiniz.",
        "req_approved_user": "🎉 *Tebrikler! Okul Yönetimi Erişim Talebinizi Onayladı.*\n\nRolünüz: *{role}*\nArtık tüm menüleri kullanabilirsiniz.",
        "req_rejected_user": "❌ *Erişim Talebiniz Okul İdaresi Tarafından Reddedildi.*\n\nLütfen okul yönetimi ile doğrudan iletişime geçiniz.",
        "no_pending_requests": "✅ Bekleyen erişim talebi bulunmamaktadır.",
        "pending_requests_title": "🛎️ *İncelenmeyi Bekleyen Erişim Talepleri:*",
        # Randevu ve Menü
        "select_teacher_appointment": "🤝 Görüşmek istediğiniz öğretmeni seçiniz:",
        "prompt_appointment_note": "Lütfen görüşmek istediğiniz konu ve uygun olduğunuz zaman dilimini yazınız:",
        "appointment_sent": "✅ Randevu talebiniz öğretmene iletildi.",
        "appointment_approved_msg": "✅ Randevu talebiniz öğretmen tarafından kabul edildi!",
        "appointment_rejected_msg": "❌ Öğretmen bu saatte müsait olmadığını bildirdi.",
        "no_pending_appointments": "✅ Bekleyen randevu talebiniz bulunmamaktadır.",
        "cafeteria_today_title": "🍲 *Bugünün Yemek Menüsü ({date}):*\n\n{content}",
        "schedule_title": "📅 *{class_name} Sınıfı Ders Programı:*\n\n{content}",
        "prompt_menu_update": "🍲 Lütfen bugünün yemek menüsünü yazınız:",
        "menu_updated": "✅ Günün yemek menüsü güncellendi.",
                "btn_add_child": "➕ Başka Çocuk Ekle (Veli Kodu)",
        "prompt_add_child_code": "Lütfen diğer öğrencinizin Veli Erişim Kodunu yazınız (Örn: `VELI-123456`):",
        "child_added_success": "✅ Yeni öğrenci hesabınıza başarıyla bağlandı:\n👤 *{name}* ({class_name})",
        "btn_recent_grades": "📝 Son Girilen Notlar",
        "recent_grades_title": "📝 *Son Girilen Ders Notları:*\nDüzenlemek veya silmek için nota tıklayınız:",
        "no_recent_grades": "✅ Henüz girdiğiniz bir ders notu bulunmamaktadır.",
        "grade_detail_card": "📝 *Not Detayı*\nÖğrenci: *{name}* ({class_name})\nDers: *{subject}*\nNot: *{score}* ({badge})\nTarih: {date}",
        "btn_edit_grade": "✏️ Notu Değiştir",
        "btn_del_grade": "❌ Notu Sil",
        "grade_deleted": "🗑️ Ders notu silindi.",
        "prompt_new_score": "Lütfen yeni notu giriniz (0-100):",
        "grade_updated": "✅ Not başarıyla güncellendi.",
        "absence_critical_warning": "🚨 *KRİTİK DEVAMSIZLIK UYARISI*\n\nÖğrenciniz *{name}* için toplam devamsızlık *{count} güne* ulaştı!\n(Yasal devamsızlık sınırı 10 gündür. Lütfen okul idaresi ile görüşünüz.)",
        "btn_profile": "👤 Profilim / Bilgilerim",
        "profile_title": "👤 *Kullanıcı Profil Kartı*\n\nAd Soyad: *{name}*\nRol: *{role}*\nTelegram ID: `{id}`\n{details}",
        "help_text": "📖 *Okul Yönetim Sistemi Kılavuzu*\n\nRolünüz: *{role}*\n\n• Alt menüyü kullanarak işlemlerinizi gerçekleştirebilirsiniz.\n• Herhangi bir formdayken iptal etmek için '❌ İşlemi İptal Et' veya /cancel yazabilirsiniz.\n• Sorularınız için okul idaresine başvurabilirsiniz.",
        "btn_risk_radar": "⚠️ Riskli Öğrenci Radarı",
        "risk_radar_title": "⚠️ *Riskli Öğrenci Radarı (Devamsızlık & Düşük Not):*",
        "no_risky_students": "✅ Şu anda risk grubunda bulunan öğrenci bulunmamaktadır.",
        "btn_audit_logs": "📜 İşlem Geçmişi (Audit Log)",
        "audit_logs_title": "📜 *Son Sistem İşlem Geçmişi:*",
        "no_audit_logs": "✅ Henüz kayıtlı işlem geçmişi bulunmamaktadır.",
        "btn_academic_report": "📈 Başarı Analizi",
        "academic_report_title": "📈 *Okul Akademik Başarı Sıralaması (Sınıflar):*",
        "btn_att_check": "📋 Yoklama Denetimi",
        "att_check_title": "📋 *Günün Yoklama Denetimi ({date}):*",
        "att_check_all_done": "✅ Tebrikler! Tüm sınıfların yoklaması eksiksiz alınmıştır.",
        "btn_remind_att": "🔔 Yoklama Hatırlat",
        "remind_att_sent": "📢 Yoklama almayan sınıfların öğretmenlerine anlık hatırlatma gönderildi.",
        "teacher_att_reminder_msg": "⚠️ *YOKLAMA HATIRLATMASI*\n\nSayın Hocam, *{class_name}* sınıfının sabah yoklaması henüz sisteme girilmemiştir. Lütfen ders yoklamasını alınız.",
        "morning_briefing_header": "☀️ *GÜNAYDIN SAYIN MÜDÜRÜM (İDARİ BRİFİNG - 09:30)*\n\n🏫 Toplam: *{total}* | ✅ Okulda: *{present}* (%{pct})\n❌ Gelmeyen: *{absent}* | 🏥 Mazeretli: *{excused}*\n\n⚠️ *Yoklama Almayan Sınıflar ({missing_cnt}):*\n{missing}\n\n🛎️ Bekleyen İşler: *{req_cnt}* Başvuru, *{med_cnt}* Rapor",
                "rk_cat_staff": "👥 Kadro & Öğrenci",
        "rk_cat_reports": "📊 Raporlar & Denetim",
        "rk_cat_requests": "🛎️ Onay Masası",
        "rk_cat_tools": "🛠️ İdari Araçlar",
        "rk_cat_settings": "⚙️ Sistem & Ayarlar",
        "cat_staff_title": "👥 *Okul Kadrosu & Öğrenci Yönetimi*\nLütfen işlem seçiniz:",
        "cat_reports_title": "📊 *Akademik & Yoklama Denetim Masası*\nLütfen incelemek istediğiniz raporu seçiniz:",
        "cat_requests_title": "🛎️ *Başvuru & Mazeret Onay Masası*\nLütfen işlem seçiniz:",
        "cat_tools_title": "🛠️ *İdari Araçlar & İletişim Masası*\nLütfen işlem seçiniz:",
        "cat_settings_title": "⚙️ *Sistem & Güvenlik Ayarları*\nLütfen işlem seçiniz:",
                "rk_parent_info": "ℹ️ Okul Bilgi Panosu",
        "rk_parent_settings": "⚙️ Ayarlar & Çıkış",
        "btn_excel_hub": "📥 Excel Merkezi (Yükle / İndir)",
        "excel_hub_title": "📥 *Excel İşlemleri Merkezi*\n\nLütfen yapmak istediğiniz işlemi seçiniz:",
        "parent_info_title": "ℹ️ *Okul Bilgi Panosu*\n\nLütfen incelemek istediğiniz bölümü seçiniz:",
        "parent_settings_title": "⚙️ *Veli Hesap & Bildirim Ayarları*\n\nLütfen ayar seçiniz:",
        "btn_my_hws": "📚 Yayınlanan Ödevler",
        "no_hws_found": "✅ Henüz yayınlanmış bir ödeviniz bulunmamaktadır.",
        "btn_cockpit_unified": "📊 Sabah Kokpiti & Yoklama",
        "btn_teachers": "👨‍🏫 Öğretmenler",
        "btn_export_all": "📥 Genel Veri Yedeği (Excel)",
        "btn_notices": "📢 Son Duyurular",
        "btn_blacklist": "🚫 Engelli Kullanıcılar",
        "btn_edit_student": "✏️ Bilgileri Düzenle",
        "btn_briefing_on": "🔔 18:30 Bülteni: Açık",
        "btn_briefing_off": "🔕 18:30 Bülteni: Kapalı",
        "teacher_card": "👨‍🏫 *Öğretmen Kartı*\nİsim: *{name}*\nBranş: *{subject}*\n🔑 Giriş Kodu: `{code}`\n📱 Telegram: {status}",
        "teacher_deleted": "🗑️ Öğretmen sistemden silindi.",
        "tch_code_reset": "✅ Öğretmen kodu sıfırlandı:\n`{code}`",
        "student_updated": "✅ Öğrenci bilgisi güncellendi.",
        "unban_success": "✅ Kullanıcının engeli kaldırıldı.",
        "no_notices": "📢 Henüz yayınlanmış bir duyuru bulunmamaktadır.",
        "notices_title": "📢 *Okul Duyuruları (Son Duyurular):*",
        "no_blacklisted": "✅ Sistemde engelli kullanıcı bulunmamaktadır.",
        "blacklisted_title": "🚫 *Engelli / Kilitli Kullanıcılar:*",
        "export_ready": "📥 *Tüm Okul Veri Yedeği Ektedir.* ({date})\n(Öğrenciler, Öğretmenler, Yoklamalar, Notlar)",
        "prompt_edit_value": "Lütfen yeni değeri yazınız:",
        "rk_notices": "📢 Son Duyurular",
        "rk_teachers": "👨‍🏫 Öğretmenler",
        # Reply Keyboard Alt Menü Butonları (4 Dilli Eşleştirme İçin)
        "rk_cancel_action": "❌ İşlemi İptal Et",
        "rk_admin_dash": "⚡ Yönetici Paneli",
        "rk_cockpit": "📊 Sabah Kokpiti",
        "rk_classes": "🏫 Sınıflar & Öğrenciler",
        "rk_requests": "🛎️ Talepler & Onaylar",
        "rk_attendance": "📋 Hızlı Yoklama",
        "rk_grade": "📝 Not Girişi",
        "rk_homework": "📢 Ödev Panosu",
        "rk_appointments": "🤝 Veli Randevuları",
        "rk_report": "📊 Durum Paneli (Karne)",
        "rk_schedule": "📅 Ders Programı",
        "rk_cafeteria": "🍲 Günün Menüsü",
        "rk_switch_student": "🧑‍🎓 Öğrenci Değiştir",
        "rk_upload_medical": "🏥 Mazeret / Rapor",
        "rk_main_menu": "🏠 Ana Menü",
        "rk_lang": "🌐 Dil Değiştir",
        "rk_logout": "🚪 Çıkış Yap",
        "rk_enter_code": "🔑 Şifre / Kod Gir",
        "rk_req_access": "📩 Şifre / Erişim İste"
    },
    "ru": {
        "lang_select": "🌍 Пожалуйста, выберите язык:",
        "lang_changed": "Язык успешно изменен: 🇷🇺 Русский",
        "prompt_enter_code_direct": "🔑 *Введите ваш код доступа:* (Например: `HCA-123456`, `VELI-123456`, `OGR-123456`)",
        "welcome_guest": "🎓 *Добро пожаловать в систему управления школой.*\n\nВыберите действие или введите выданный **код доступа**:",
        "auth_success": "✅ Авторизация успешна!\nДобро пожаловать: *{name}*\nВаша роль: *{role}*",
        "auth_failed": "❌ Неверный код доступа! Осталось попыток: {remaining}",
        "auth_locked": "⛔ Аккаунт заблокирован на 1 час из соображений безопасности.",
        "auth_blacklisted": "🚫 Ваш аккаунт заблокирован навсегда.",
        "maintenance_mode": "⚠️ В системе ведутся технические работы. Пожалуйста, попробуйте позже.",
        "admin_title": "⚡ *Панель управления школой (Администратор)*",
        "admin_stats": "📊 *Общий статус:*\n• Классы: *{c_cnt}* | Ученики: *{s_cnt}* | Учителя: *{t_cnt}*\n• Заявки: *{req_cnt}* | Справки: *{med_cnt}*\n• Дата: *{date}*",
        "btn_add_student": "➕ Добавить ученика",
        "btn_add_teacher": "➕ Добавить учителя",
        "btn_classes": "🏫 Классы и ученики",
        "btn_search_student": "🔍 Поиск ученика",
        "btn_excel": "📥 Импорт из Excel",
        "btn_cockpit": "📊 Утренний статус",
        "btn_pdf": "📄 Карточки с кодами (PDF)",
        "btn_medical": "🏥 Справки ({count})",
        "btn_requests": "🛎️ Заявки на доступ ({count})",
        "btn_broadcast": "📢 Рассылка объявления",
        "btn_maintenance_toggle": "🚨 Режим обслуживания ({status})",
        "btn_cafeteria_edit": "🍲 Обновить меню столовой",
        "btn_lang": "🌐 Сменить язык",
        "btn_back": "⬅️ Назад",
        "btn_main_menu": "🏠 Главное меню",
        "btn_acknowledged": "✅ Ознакомлен(а)",
        "acknowledged_toast": "Подтверждение принято.",
        "no_classes_found": "⚠️ Пока не зарегистрировано ни одного класса.",
        "prompt_student_name": "👤 Введите Фамилию и Имя ученика:",
        "prompt_student_class": "🏫 Введите класс ученика (Например: `9-A`):",
        "prompt_student_no": "🔢 Введите номер ученика (Например: `101`):",
        "student_added_card": "✅ *Ученик успешно добавлен!*\n\n👤 ФИО: *{name}*\n🏫 Класс: *{class_name}* | №: *{no}*\n\n🔑 *Коды доступа:*\n• Код ученика: `{st_code}`\n• Код родителя: `{pr_code}`",
        "prompt_teacher_name": "👨‍🏫 Введите ФИО учителя:",
        "prompt_teacher_subject": "📚 Введите предмет (Например: `Математика`):",
        "teacher_added_card": "✅ *Учитель зарегистрирован!*\n\n👤 ФИО: *{name}*\n📚 Предмет: *{subject}*\n\n🔑 *Код доступа:*\n`{code}`",
        "student_card": "👤 *Карточка ученика*\nФИО: *{name}*\nКласс: *{class_name}* | №: *{no}*\n\n🔑 *Статус кодов:*\n• Ученик: `{st_code}` ({st_status})\n• Родитель: `{pr_code}` ({pr_status})",
        "btn_reset_codes": "🔄 Сбросить коды",
        "btn_del_student": "❌ Удалить ученика",
        "btn_student_grades": "📊 Оценки и пропуски",
        "student_deleted": "🗑️ Ученик удален из системы.",
        "codes_reset_done": "✅ Коды обновлены!\n\n• Новый код ученика: `{st_code}`\n• Новый код родителя: `{pr_code}`",
        "excel_info": "📥 *Импорт через Excel*\n\nОтправьте файл `.xlsx`.\nЗаголовки: `Ad Soyad` | `Sinif` | `Numara`",
        "excel_done": "✅ Обработано! Добавлено учеников: *{count}*\nКоды прикреплены в файле.",
        "cockpit_report": "📊 *Утренняя сводка ({date})*\n\n🏫 Всего: *{total}* | ✅ Есть: *{present}* | ❌ Нет: *{absent}*\n\n⚠️ *Классы без переклички ({missing_cnt}):*\n{missing}",
        "select_pdf_class": "📄 Выберите класс для карточек:",
        "pdf_ready": "📄 Карточки для *{class_name}* готовы.",
        "pending_medical_title": "🏥 *Справки на рассмотрении:*",
        "no_pending_medical": "✅ Нет справок, ожидающих проверки.",
        "medical_approved": "✅ Справка одобрена.",
        "medical_rejected": "❌ Справка отклонена.",
        "prompt_broadcast": "📢 Введите текст объявления:",
        "broadcast_success": "📢 Объявление доставлено пользователям: *{count}*.",
        "prompt_search_student": "🔍 Введите фамилию или номер ученика:",
        "search_no_results": "❌ Ученик не найден.",
        "search_results_title": "🔍 *Результаты поиска:*",
        "menu_teacher": "👨‍🏫 *Панель учителя*\nУчитель: *{name}* ({subject})",
        "btn_attendance": "📋 Быстрая перекличка",
        "btn_enter_grade": "📝 Выставить оценки",
        "btn_homework_board": "📢 Доска заданий",
        "btn_teacher_appointments": "🤝 Записи родителей ({count})",
        "attendance_intro": "📋 *Перекличка {class_name}*\nОтметьте отсутствующих и сохраните:",
        "btn_save_att": "💾 Сохранить",
        "att_saved": "✅ Перекличка сохранена (правки 15 мин).",
        "prompt_grade_score": "Ученик: *{name}* ({class_name})\nВведите оценку (0-100):",
        "prompt_grade_badge": "Выберите категорию оценки:",
        "badge_praise": "🟢 Похвала / Успех",
        "badge_missing": "🟡 Пробел / Доработать",
        "badge_warning": "🔴 Замечание / Дисциплина",
        "grade_saved_success": "✅ Оценка отправлена родителю.",
        "prompt_hw_class": "Выберите класс для задания:",
        "prompt_hw_content": "Отправьте текст задания или фото доски:",
        "hw_sent_success": "📢 Задание отправлено классу *{class_name}*.",
        "menu_parent": "👨‍👩‍👧‍👦 *Панель родителя*\nУченик: *{name}* ({class_name})",
        "btn_report_card": "📊 Табель успеваемости",
        "btn_switch_student": "🧑‍🎓 Сменить ученика",
        "btn_upload_medical": "🏥 Отправить справку",
        "btn_book_appointment": "🤝 Запись к учителю",
        "btn_view_schedule": "📅 Расписание уроков",
        "btn_view_cafeteria": "🍲 Меню столовой",
        "upload_med_prompt": "Пожалуйста, отправьте фото справки:",
        "med_uploaded_success": "Справка отправлена администрации.",
        "student_switched_success": "Выбран ученик: *{name}* ({class_name})",
        "menu_student": "🎓 *Панель ученика*\nУченик: *{name}* ({class_name} - №: {no})",
        "no_grades": "Оценки пока не выставлены.",
        "no_linked_student": "⚠️ Ученик не найден в системе.",
        "evening_briefing_header": "🌙 *ИТОГИ ДНЯ (18:30)*\nУченик: *{name}* ({class_name})\n\n📌 Посещаемость: *{att_status}*\n📝 Оценки:\n{grades}",
        "logged_out": "🚪 Вы успешно вышли из системы.",
        # Talep Sistemi
        "btn_req_access": "📩 Запросить доступ",
        "req_role_select": "📩 *Заявка на доступ к системе*\n\nВыберите вашу роль в школе:",
        "req_name_prompt": "👤 Введите Фамилию и Имя:",
        "req_phone_prompt": "📱 Введите номер телефона (Например: `+79991234567`):",
        "req_details_teacher": "📚 Введите предмет преподавания (Например: `Математика`):",
        "req_details_student": "🏫 Введите класс и номер (Например: `9-A / 101`):",
        "req_details_parent": "🧑‍🎓 Введите ФИО ребенка, класс и номер (Например: `Иван Иванов, 9-A, 101`):",
        "req_sent_success": "✅ *Заявка отправлена администрации!*\n\nПосле проверки вам придет уведомление с доступом в этот чат.",
        "req_already_pending": "⚠️ У вас уже есть активная заявка на рассмотрении. Ожидайте подтверждения.",
        "req_cooldown": "⚠️ Предыдущая заявка была отклонена. Повторная отправка доступна через 24 часа.",
        "req_approved_user": "🎉 *Поздравляем! Администрация одобрила доступ.*\n\nВаша роль: *{role}*\nВсе функции открыты.",
        "req_rejected_user": "❌ *Администрация отклонила заявку на доступ.*\n\nОбратитесь к руководству школы напрямую.",
        "no_pending_requests": "✅ Нет заявок, ожидающих рассмотрения.",
        "pending_requests_title": "🛎️ *Заявки на регистрацию в системе:*",
        # Randevu ve Menü
        "select_teacher_appointment": "🤝 Выберите учителя для записи:",
        "prompt_appointment_note": "Укажите тему встречи и удобное для вас время:",
        "appointment_sent": "✅ Заявка на встречу отправлена учителю.",
        "appointment_approved_msg": "✅ Учитель подтвердил встречу!",
        "appointment_rejected_msg": "❌ Учитель сообщил, что не сможет в это время.",
        "no_pending_appointments": "✅ Нет записей, ожидающих ответа.",
        "cafeteria_today_title": "🍲 *Меню школьной столовой ({date}):*\n\n{content}",
        "schedule_title": "📅 *Расписание класса {class_name}:*\n\n{content}",
        "prompt_menu_update": "🍲 Введите меню школьной столовой на сегодня:",
        "menu_updated": "✅ Меню столовой успешно обновлено.",
                "btn_add_child": "➕ Привязать еще ученика",
        "prompt_add_child_code": "Введите код родителя для второго ребенка (Например: `VELI-123456`):",
        "child_added_success": "✅ Новый ученик успешно привязан:\n👤 *{name}* ({class_name})",
        "btn_recent_grades": "📝 Выставленные оценки",
        "recent_grades_title": "📝 *Недавние оценки:*\nНажмите для редактирования или удаления:",
        "no_recent_grades": "✅ Вы еще не выставляли оценок.",
        "grade_detail_card": "📝 *Детали оценки*\nУченик: *{name}* ({class_name})\nПредмет: *{subject}*\nОценка: *{score}* ({badge})\nДата: {date}",
        "btn_edit_grade": "✏️ Изменить оценку",
        "btn_del_grade": "❌ Удалить оценку",
        "grade_deleted": "🗑️ Оценка удалена.",
        "prompt_new_score": "Введите новую оценку (0-100):",
        "grade_updated": "✅ Оценка обновлена.",
        "absence_critical_warning": "🚨 *ВНИМАНИЕ: КРИТИЧЕСКИЙ ПРОПУСК*\n\nОбщее число пропусков ученика *{name}* достигло *{count} дней*!\n(Лимит 10 дней. Пожалуйста, обратитесь к администрации.)",
        "btn_profile": "👤 Мой профиль",
        "profile_title": "👤 *Профиль пользователя*\n\nФИО: *{name}*\nРоль: *{role}*\nTelegram ID: `{id}`\n{details}",
        "help_text": "📖 *Справка по системе*\n\nВаша роль: *{role}*\n\n• Используйте меню снизу для навигации.\n• Для отмены действия напишите /cancel или нажмите '❌ Отменить действие'.",
        "btn_risk_radar": "⚠️ Радар успеваемости и рисков",
        "risk_radar_title": "⚠️ *Радар рисков (Пропуски и низкие баллы):*",
        "no_risky_students": "✅ Учеников в группе риска нет.",
        "btn_audit_logs": "📜 Журнал действий (Audit Log)",
        "audit_logs_title": "📜 *История операций в системе:*",
        "no_audit_logs": "✅ Журнал действий пуст.",
        "btn_academic_report": "📈 Рейтинг классов",
        "academic_report_title": "📈 *Академический рейтинг классов:*",
        "btn_att_check": "📋 Контроль переклички",
        "att_check_title": "📋 *Контроль переклички на сегодня ({date}):*",
        "att_check_all_done": "✅ Все классы успешно прошли перекличку.",
        "btn_remind_att": "🔔 Напомнить учителям",
        "remind_att_sent": "📢 Уведомление учителям отправлено.",
        "teacher_att_reminder_msg": "⚠️ *НАПОМИНАНИЕ О ПЕРЕКЛИЧКЕ*\n\nУважаемый учитель, перекличка класса *{class_name}* еще не внесена. Пожалуйста, заполните журнал.",
        "morning_briefing_header": "☀️ *УТРЕННИЙ ОТЧЕТ ДИРЕКТОРУ (09:30)*\n\n🏫 Всего: *{total}* | ✅ В школе: *{present}* (%{pct})\n❌ Отсутствуют: *{absent}* | 🏥 Справки: *{excused}*\n\n⚠️ *Классы без переклички ({missing_cnt}):*\n{missing}\n\n🛎️ Заявки: *{req_cnt}*, Справки: *{med_cnt}*",
                "rk_cat_staff": "👥 Ученики и учителя",
        "rk_cat_reports": "📊 Отчеты и контроль",
        "rk_cat_requests": "🛎️ Центр одобрений",
        "rk_cat_tools": "🛠️ Инструменты",
        "rk_cat_settings": "⚙️ Настройки системы",
        "cat_staff_title": "👥 *Управление учениками и учителями*\nВыберите действие:",
        "cat_reports_title": "📊 *Академический контроль и отчеты*\nВыберите раздел:",
        "cat_requests_title": "🛎️ *Заявки и медицинские справки*\nВыберите действие:",
        "cat_tools_title": "🛠️ *Инструменты управления и рассылки*\nВыберите раздел:",
        "cat_settings_title": "⚙️ *Системные настройки и безопасность*\nВыберите действие:",
                "rk_parent_info": "ℹ️ Инфопанель школы",
        "rk_parent_settings": "⚙️ Настройки и выход",
        "btn_excel_hub": "📥 Центр Excel (Импорт / Экспорт)",
        "excel_hub_title": "📥 *Центр операций Excel*\n\nВыберите действие:",
        "parent_info_title": "ℹ️ *Информационная панель школы*\n\nВыберите раздел:",
        "parent_settings_title": "⚙️ *Настройки уведомлений и аккаунта*\n\nВыберите действие:",
        "btn_my_hws": "📚 Мои задания",
        "no_hws_found": "✅ Опубликованных заданий пока нет.",
        "btn_cockpit_unified": "📊 Утренний статус и перекличка",
        "btn_teachers": "👨‍🏫 Учителя",
        "btn_export_all": "📥 Экспорт базы (Excel)",
        "btn_notices": "📢 Объявления школы",
        "btn_blacklist": "🚫 Заблокированные",
        "btn_edit_student": "✏️ Редактировать",
        "btn_briefing_on": "🔔 Итоги дня: Вкл",
        "btn_briefing_off": "🔕 Итоги дня: Выкл",
        "teacher_card": "👨‍🏫 *Карточка учителя*\nФИО: *{name}*\nПредмет: *{subject}*\n🔑 Код доступа: `{code}`\n📱 Telegram: {status}",
        "teacher_deleted": "🗑️ Учитель удален из системы.",
        "tch_code_reset": "✅ Код доступа учителя обновлен:\n`{code}`",
        "student_updated": "✅ Данные ученика обновлены.",
        "unban_success": "✅ Блокировка успешно снята.",
        "no_notices": "📢 Объявлений пока нет.",
        "notices_title": "📢 *Объявления школы:*",
        "no_blacklisted": "✅ Заблокированных пользователей нет.",
        "blacklisted_title": "🚫 *Заблокированные аккаунты:*",
        "export_ready": "📥 *Полный архив базы данных школы* ({date})\n(Ученики, Учителя, Перекличка, Оценки)",
        "prompt_edit_value": "Введите новое значение:",
        "rk_notices": "📢 Объявления",
        "rk_teachers": "👨‍🏫 Учителя",
        # Reply Keyboard
        "rk_cancel_action": "❌ Отменить действие",
        "rk_admin_dash": "⚡ Панель управления",
        "rk_cockpit": "📊 Утренний статус",
        "rk_classes": "🏫 Классы и ученики",
        "rk_requests": "🛎️ Заявки и справки",
        "rk_attendance": "📋 Быстрая перекличка",
        "rk_grade": "📝 Выставить оценки",
        "rk_homework": "📢 Доска заданий",
        "rk_appointments": "🤝 Записи родителей",
        "rk_report": "📊 Табель успеваемости",
        "rk_schedule": "📅 Расписание уроков",
        "rk_cafeteria": "🍲 Меню столовой",
        "rk_switch_student": "🧑‍🎓 Сменить ученика",
        "rk_upload_medical": "🏥 Отправить справку",
        "rk_main_menu": "🏠 Главное меню",
        "rk_lang": "🌐 Сменить язык",
        "rk_logout": "🚪 Выйти",
        "rk_enter_code": "🔑 Ввести код доступа",
        "rk_req_access": "📩 Запросить доступ"
    },
    "uz": {
        "lang_select": "🌍 Iltimos, tilni tanlang:",
        "lang_changed": "Til muvaffaqiyatli yangilandi: 🇺🇿 O'zbekcha",
        "prompt_enter_code_direct": "🔑 *Iltimos, sizga berilgan kirish kodini yozing:* (Masalan: `HCA-123456`, `VELI-123456`, `OGR-123456`)",
        "welcome_guest": "🎓 *Maktab boshqaruv tizimiga xush kelibsiz.*\n\nIltimos, amalni tanlang yoki berilgan **kirish kodini** yozing:",
        "auth_success": "✅ Kirish muvaffaqiyatli!\nXush kelibsiz: *{name}*\nSizning rolingiz: *{role}*",
        "auth_failed": "❌ Noto'g'ri kod! Qolgan urinishlar soni: {remaining}",
        "auth_locked": "⛔ Xavfsizlik sababli hisobingiz 1 soatga bloklandi.",
        "auth_blacklisted": "🚫 Hisobingiz butunlay to'xtatildi.",
        "maintenance_mode": "⚠️ Tizimda rejaviy texnik ishlar olib borilmoqda. Iltimos, keyinroq urinib ko'ring.",
        "admin_title": "⚡ *Maktab boshqaruv markazi (Admin)*",
        "admin_stats": "📊 *Umumiy holat:*\n• Sinflar: *{c_cnt}* | O'quvchilar: *{s_cnt}* | O'qituvchilar: *{t_cnt}*\n• Arizalar: *{req_cnt}* | Ma'lumotnomalar: *{med_cnt}*\n• Sana: *{date}*",
        "btn_add_student": "➕ O'quvchi qo'shish",
        "btn_add_teacher": "➕ O'qituvchi qo'shish",
        "btn_classes": "🏫 Sinflar va o'quvchilar",
        "btn_search_student": "🔍 O'quvchini qidirish",
        "btn_excel": "📥 Excel orqali yuklash",
        "btn_cockpit": "📊 Tonggi hisobot",
        "btn_pdf": "📄 Parol kartalari (PDF)",
        "btn_medical": "🏥 Ma'lumotnomalar ({count})",
        "btn_requests": "🛎️ Kirish arizalari ({count})",
        "btn_broadcast": "📢 Ommaviy e'lon",
        "btn_maintenance_toggle": "🚨 Ta'mir rejimi ({status})",
        "btn_cafeteria_edit": "🍲 Oshxona menyusini yangilash",
        "btn_lang": "🌐 Tilni o'zgartirish",
        "btn_back": "⬅️ Orqaga",
        "btn_main_menu": "🏠 Asosiy menyu",
        "btn_acknowledged": "✅ O'qidim / Xabardorman",
        "acknowledged_toast": "Tasdiqlaganingiz saqlandi.",
        "no_classes_found": "⚠️ Hozircha birorta sinf mavjud emas.",
        "prompt_student_name": "👤 O'quvchining Familiyasi va Ismini kiriting:",
        "prompt_student_class": "🏫 O'quvchining sinfini kiriting (Masalan: `9-A`):",
        "prompt_student_no": "🔢 O'quvchining raqamini kiriting (Masalan: `101`):",
        "student_added_card": "✅ *O'quvchi saqlandi!*\n\n👤 Ism: *{name}*\n🏫 Sinf: *{class_name}* | №: *{no}*\n\n🔑 *Kodlar:*\n• O'quvchi: `{st_code}`\n• Ota-ona: `{pr_code}`",
        "prompt_teacher_name": "👨‍🏫 O'qituvchining Familiyasi va Ismini kiriting:",
        "prompt_teacher_subject": "📚 Fanni kiriting (Masalan: `Matematika`):",
        "teacher_added_card": "✅ *O'qituvchi saqlandi!*\n\n👤 Ism: *{name}*\n📚 Fan: *{subject}*\n\n🔑 *O'qituvchi kodi:*\n`{code}`",
        "student_card": "👤 *O'quvchi kartasi*\nF.I.SH: *{name}*\nSinf: *{class_name}* | №: *{no}*\n\n🔑 *Kodlar:*\n• O'quvchi: `{st_code}` ({st_status})\n• Ota-ona: `{pr_code}` ({pr_status})",
        "btn_reset_codes": "🔄 Kodlarni yangilash",
        "btn_del_student": "❌ O'quvchini o'chirish",
        "btn_student_grades": "📊 Baholari va davomati",
        "student_deleted": "🗑️ O'quvchi tizimdan o'chirildi.",
        "codes_reset_done": "✅ Kodlar yangilandi!\n\n• Yangi O'quvchi: `{st_code}`\n• Yangi Ota-ona: `{pr_code}`",
        "excel_info": "📥 *Excel orqali yuklash*\n\n`.xlsx` faylini yuboring.\nUstunlar: `Ad Soyad` | `Sinif` | `Numara`",
        "excel_done": "✅ Qabul qilindi! O'quvchilar: *{count}*\nParollar biriktirilgan faylda berildi.",
        "cockpit_report": "📊 *Tonggi hisobot ({date})*\n\n🏫 Jami: *{total}* | ✅ Kelgan: *{present}* | ❌ Kelmagan: *{absent}*\n\n⚠️ *Davomat olinmagan ({missing_cnt}):*\n{missing}",
        "select_pdf_class": "📄 Parol kartalari uchun sinfni tanlang:",
        "pdf_ready": "📄 *{class_name}* kartalari tayyorlandi.",
        "pending_medical_title": "🏥 *Ko'rib chiqilishi kerak bo'lgan ma'lumotnomalar:*",
        "no_pending_medical": "✅ Kutilayotgan ma'lumotnomalar mavjud emas.",
        "medical_approved": "✅ Ma'lumotnoma tasdiqlandi.",
        "medical_rejected": "❌ Ma'lumotnoma rad etildi.",
        "prompt_broadcast": "📢 E'lon matnini yozing:",
        "broadcast_success": "📢 E'lon *{count}* ta foydalanuvchiga yetkazildi.",
        "prompt_search_student": "🔍 Ism yoki raqam kiriting:",
        "search_no_results": "❌ O'quvchi topilmadi.",
        "search_results_title": "🔍 *Qidiruv natijalari:*",
        "menu_teacher": "👨‍🏫 *O'qituvchi paneli*\nO'qituvchi: *{name}* ({subject})",
        "btn_attendance": "📋 Tezkor davomat",
        "btn_enter_grade": "📝 Baho qo'yish",
        "btn_homework_board": "📢 Vazifalar paneli",
        "btn_teacher_appointments": "🤝 Ota-onalar uchrashuvi ({count})",
        "attendance_intro": "📋 *{class_name} davomati*\nKelmaganlarni belgilab saqlang:",
        "btn_save_att": "💾 Saqlash",
        "att_saved": "✅ Davomat saqlandi (15 daqiqa tuzatish vaqti).",
        "prompt_grade_score": "O'quvchi: *{name}* ({class_name})\nBahoni kiriting (0-100):",
        "prompt_grade_badge": "Toifani tanlang:",
        "badge_praise": "🟢 Maqtov / Muvaffaqiyat",
        "badge_missing": "🟡 Kamchilik / O'rganish",
        "badge_warning": "🔴 Intizom / Ogohlantirish",
        "grade_saved_success": "✅ Baho saqlandi va ota-onaga yuborildi.",
        "prompt_hw_class": "Vazifa sinfini tanlang:",
        "prompt_hw_content": "Vazifa matnini yozing yoki doska rasmini yuboring:",
        "hw_sent_success": "📢 Vazifa *{class_name}* sinfiga yuborildi.",
        "menu_parent": "👨‍👩‍👧‍👦 *Ota-ona paneli*\nO'quvchi: *{name}* ({class_name})",
        "btn_report_card": "📊 Holat paneli (Kundalik)",
        "btn_switch_student": "🧑‍🎓 O'quvchini tanlash",
        "btn_upload_medical": "🏥 Ma'lumotnoma yuborish",
        "btn_book_appointment": "🤝 O'qituvchi bilan uchrashuv",
        "btn_view_schedule": "📅 Dars jadvali",
        "btn_view_cafeteria": "🍲 Bugungi oshxona menyusi",
        "upload_med_prompt": "Iltimos, ma'lumotnoma rasmini yuboring:",
        "med_uploaded_success": "Ma'lumotnoma ma'muriyatga yuborildi.",
        "student_switched_success": "Tanlangan o'quvchi: *{name}* ({class_name})",
        "menu_student": "🎓 *O'quvchi paneli*\nO'quvchi: *{name}* ({class_name} - №: {no})",
        "no_grades": "Hozircha baholar mavjud emas.",
        "no_linked_student": "⚠️ Sizga biriktirilgan o'quvchi topilmadi.",
        "evening_briefing_header": "🌙 *KUNLIK YAKUNIY HISOBOT (18:30)*\nO'quvchi: *{name}* ({class_name})\n\n📌 Davomat: *{att_status}*\n📝 Baholar:\n{grades}",
        "logged_out": "🚪 Tizimdan chiqildi.",
        # Talep Sistemi
        "btn_req_access": "📩 Ruxsat / Parol so'rash",
        "req_role_select": "📩 *Tizimga kirish uchun ariza*\n\nIltimos, maktabdagi rolingizni tanlang:",
        "req_name_prompt": "👤 Familiyangiz va Ismingizni kiriting:",
        "req_phone_prompt": "📱 Telefon raqamingizni kiriting (Masalan: `+998901234567`):",
        "req_details_teacher": "📚 O'qitadigan faningizni yozing (Masalan: `Matematika`):",
        "req_details_student": "🏫 Sinfingiz va raqamingizni yozing (Masalan: `9-A / 101`):",
        "req_details_parent": "🧑‍🎓 Farzandingizning F.I.SH, sinfi va raqamini yozing (Masalan: `Ali Valiyev, 9-A, 101`):",
        "req_sent_success": "✅ *Arizangiz maktab ma'muriyatiga yuborildi!*\n\nTasdiqlangandan so'ng Telegram orqali xabarnoma olasiz.",
        "req_already_pending": "⚠️ Sizda allaqachon ko'rib chiqilayotgan ariza mavjud. Iltimos, kuting.",
        "req_cooldown": "⚠️ Oldingi arizangiz rad etilgan. 24 soatdan so'ng qayta ariza topshirishingiz mumkin.",
        "req_approved_user": "🎉 *Tabriklaymiz! Maktab ma'muriyati arizangizni tasdiqladi.*\n\nRolingiz: *{role}*\nBarcha xizmatlar ochildi.",
        "req_rejected_user": "❌ *Arizangiz maktab ma'muriyati tomonidan rad etildi.*\n\nIltimos, maktab bilan to'g'ridan-to'g'ri bog'laning.",
        "no_pending_requests": "✅ Kutilayotgan arizalar mavjud emas.",
        "pending_requests_title": "🛎️ *Ko'rib chiqilishi kerak bo'lgan arizalar:*",
        # Randevu ve Menü
        "select_teacher_appointment": "🤝 Uchrashmoqchi bo'lgan o'qituvchini tanlang:",
        "prompt_appointment_note": "Uchrashuv mavzusi va o'zingizga qulay vaqtni yozing:",
        "appointment_sent": "✅ Uchrashuv taklifi o'qituvchiga yuborildi.",
        "appointment_approved_msg": "✅ O'qituvchi uchrashuvni qabul qildi!",
        "appointment_rejected_msg": "❌ O'qituvchi bu vaqtda bo'sh emasligini bildirdi.",
        "no_pending_appointments": "✅ Kutilayotgan uchrashuvlar mavjud emas.",
        "cafeteria_today_title": "🍲 *Bugungi oshxona menyusi ({date}):*\n\n{content}",
        "schedule_title": "📅 *{class_name} sinfi dars jadvali:*\n\n{content}",
        "prompt_menu_update": "🍲 Bugungi oshxona menyusini kiriting:",
        "menu_updated": "✅ Oshxona menyusi muvaffaqiyatli yangilandi.",
                "btn_add_child": "➕ Boshqa farzandni ulash",
        "prompt_add_child_code": "Iltimos, boshqa farzandingizning Ota-ona kodini kiriting (Masalan: `VELI-123456`):",
        "child_added_success": "✅ Yangi o'quvchi hisobingizga muvaffaqiyatli ulandi:\n👤 *{name}* ({class_name})",
        "btn_recent_grades": "📝 Qo'yilgan baholar",
        "recent_grades_title": "📝 *So'nggi qo'yilgan baholar:*\nTahrirlash yoki o'chirish uchun tanlang:",
        "no_recent_grades": "✅ Hozircha baholar qo'yilmagan.",
        "grade_detail_card": "📝 *Baho tafsiloti*\nO'quvchi: *{name}* ({class_name})\nFan: *{subject}*\nBaho: *{score}* ({badge})\nSana: {date}",
        "btn_edit_grade": "✏️ Bahoni o'zgartirish",
        "btn_del_grade": "❌ Bahoni o'chirish",
        "grade_deleted": "🗑️ Baho o'chirildi.",
        "prompt_new_score": "Yangi bahoni kiriting (0-100):",
        "grade_updated": "✅ Baho muvaffaqiyatli yangilandi.",
        "absence_critical_warning": "🚨 *DIQQAT: KRITIK DAVOMAT OGOHLANTIRISHI*\n\nFarzandingiz *{name}* jami davomati *{count} kunga* yetdi!\n(Qonuniy chegara 10 kun. Iltimos, ma'muriyat bilan bog'laning.)",
        "btn_profile": "👤 Mening profilim",
        "profile_title": "👤 *Foydalanuvchi profili*\n\nF.I.SH: *{name}*\nRol: *{role}*\nTelegram ID: `{id}`\n{details}",
        "help_text": "📖 *Tizimdan foydalanish bo'yicha qo'llanma*\n\nRolingiz: *{role}*\n\n• Pastdagi menyudan foydalaning.\n• Bekor qilish uchun /cancel yozing yoki '❌ Bekor qilish' tugmasini bosing.",
        "btn_risk_radar": "⚠️ Xavf ostidagi o'quvchilar",
        "risk_radar_title": "⚠️ *Xavf radari (Ko'p qoldirganlar va past baholar):*",
        "no_risky_students": "✅ Xavf guruhidagi o'quvchilar mavjud emas.",
        "btn_audit_logs": "📜 Tizim harakatlari (Audit Log)",
        "audit_logs_title": "📜 *So'nggi tizim amallari jurnali:*",
        "no_audit_logs": "✅ Harakatlar jurnali bo'sh.",
        "btn_academic_report": "📈 Sinf reytingi",
        "academic_report_title": "📈 *Maktab sinflarining o'zlashtirish reytingi:*",
        "btn_att_check": "📋 Davomat nazorati",
        "att_check_title": "📋 *Bugungi davomat nazorati ({date}):*",
        "att_check_all_done": "✅ Barcha sinflar davomati to'liq olingan.",
        "btn_remind_att": "🔔 Eslatma yuborish",
        "remind_att_sent": "📢 Davomat olmagan o'qituvchilarga eslatma yuborildi.",
        "teacher_att_reminder_msg": "⚠️ *DAVOMAT ESLATMASI*\n\nHurmatli ustoz, *{class_name}* sinfining tonggi davomati hali kiritilmagan. Iltimos, davomatni oling.",
        "morning_briefing_header": "☀️ *HURMATLI DIREKTOR, TONGGI HISOBOT (09:30)*\n\n🏫 Jami: *{total}* | ✅ Maktabda: *{present}* (%{pct})\n❌ Kelmagan: *{absent}* | 🏥 Sababli: *{excused}*\n\n⚠️ *Davomat olinmagan ({missing_cnt}):*\n{missing}\n\n🛎️ Arizalar: *{req_cnt}*, Ma'lumotnomalar: *{med_cnt}*",
                "rk_cat_staff": "👥 Kadro va o'quvchilar",
        "rk_cat_reports": "📊 Hisobotlar va nazorat",
        "rk_cat_requests": "🛎️ Tasdiqlash markazi",
        "rk_cat_tools": "🛠️ Boshqaruv vositalari",
        "rk_cat_settings": "⚙️ Tizim va sozlamalar",
        "cat_staff_title": "👥 *O'quvchilar va o'qituvchilar boshqaruvi*\nAmalni tanlang:",
        "cat_reports_title": "📊 *Akademik hisobotlar va davomat nazorati*\nBo'limni tanlang:",
        "cat_requests_title": "🛎️ *Arizalar va ma'lumotnomalarni tasdiqlash*\nAmalni tanlang:",
        "cat_tools_title": "🛠️ *Boshqaruv vositalari va e'lonlar*\nBo'limni tanlang:",
        "cat_settings_title": "⚙️ *Tizim va xavfsizlik sozlamalari*\nAmalni tanlang:",
                "rk_parent_info": "ℹ️ Maktab ma'lumotlari",
        "rk_parent_settings": "⚙️ Sozlamalar va chiqish",
        "btn_excel_hub": "📥 Excel markazi (Yuklash / Ko'chirish)",
        "excel_hub_title": "📥 *Excel amallari markazi*\n\nAmalni tanlang:",
        "parent_info_title": "ℹ️ *Maktab axborot paneli*\n\nBo'limni tanlang:",
        "parent_settings_title": "⚙️ *Xabarnoma va hisob sozlamalari*\n\nSozlamani tanlang:",
        "btn_my_hws": "📚 Berilgan vazifalar",
        "no_hws_found": "✅ Hozircha berilgan vazifalar yo'q.",
        "btn_cockpit_unified": "📊 Tonggi hisobot va davomat",
        "btn_teachers": "👨‍🏫 O'qituvchilar",
        "btn_export_all": "📥 Umumiy arxiv (Excel)",
        "btn_notices": "📢 Maktab e'lonlari",
        "btn_blacklist": "🚫 Bloklanganlar",
        "btn_edit_student": "✏️ Tahrirlash",
        "btn_briefing_on": "🔔 Kunlik hisobot: Yoqiq",
        "btn_briefing_off": "🔕 Kunlik hisobot: O'chiq",
        "teacher_card": "👨‍🏫 *O'qituvchi kartasi*\nF.I.SH: *{name}*\nFan: *{subject}*\n🔑 Kirish kodi: `{code}`\n📱 Telegram: {status}",
        "teacher_deleted": "🗑️ O'qituvchi tizimdan o'chirildi.",
        "tch_code_reset": "✅ O'qituvchi kodi yangilandi:\n`{code}`",
        "student_updated": "✅ O'quvchi ma'lumoti yangilandi.",
        "unban_success": "✅ Blokdan muvaffaqiyatli chiqarildi.",
        "no_notices": "📢 Hozircha e'lonlar mavjud emas.",
        "notices_title": "📢 *Maktab e'lonlari:*",
        "no_blacklisted": "✅ Bloklangan foydalanuvchilar yo'q.",
        "blacklisted_title": "🚫 *Bloklangan hisoblar:*",
        "export_ready": "📥 *Maktabning to'liq ma'lumotlar arxivi* ({date})\n(O'quvchilar, O'qituvchilar, Davomat, Baholar)",
        "prompt_edit_value": "Yangi qiymatni kiriting:",
        "rk_notices": "📢 So'nggi e'lonlar",
        "rk_teachers": "👨‍🏫 O'qituvchilar",
        # Reply Keyboard
        "rk_cancel_action": "❌ Bekor qilish",
        "rk_admin_dash": "⚡ Boshqaruv paneli",
        "rk_cockpit": "📊 Tonggi hisobot",
        "rk_classes": "🏫 Sinflar va o'quvchilar",
        "rk_requests": "🛎️ Arizalar va arizachilar",
        "rk_attendance": "📋 Tezkor davomat",
        "rk_grade": "📝 Baho qo'yish",
        "rk_homework": "📢 Vazifalar paneli",
        "rk_appointments": "🤝 Ota-onalar uchrashuvi",
        "rk_report": "📊 Holat paneli",
        "rk_schedule": "📅 Dars jadvali",
        "rk_cafeteria": "🍲 Oshxona menyusi",
        "rk_switch_student": "🧑‍🎓 O'quvchini almashtirish",
        "rk_upload_medical": "🏥 Ma'lumotnoma",
        "rk_main_menu": "🏠 Asosiy menyu",
        "rk_lang": "🌐 Tilni o'zgartirish",
        "rk_logout": "🚪 Chiqish",
        "rk_enter_code": "🔑 Kod kiritish",
        "rk_req_access": "📩 Ruxsat so'rash"
    },
    "en": {
        "lang_select": "🌍 Please select your language:",
        "lang_changed": "Language successfully updated: 🇬🇧 English",
        "prompt_enter_code_direct": "🔑 *Please enter your access code:* (e.g. `HCA-123456`, `VELI-123456`, `OGR-123456`)",
        "welcome_guest": "🎓 *Welcome to School Management Ecosystem.*\n\nPlease choose an action or enter your provided **access code**:",
        "auth_success": "✅ Authentication successful!\nWelcome: *{name}*\nYour role: *{role}*",
        "auth_failed": "❌ Invalid access code! Remaining attempts: {remaining}",
        "auth_locked": "⛔ Account locked for 1 hour due to security restrictions.",
        "auth_blacklisted": "🚫 Account has been permanently suspended.",
        "maintenance_mode": "⚠️ The system is currently under maintenance. Please try again later.",
        "admin_title": "⚡ *School Administration Cockpit (Admin)*",
        "admin_stats": "📊 *General Status:*\n• Classes: *{c_cnt}* | Students: *{s_cnt}* | Teachers: *{t_cnt}*\n• Requests: *{req_cnt}* | Medicals: *{med_cnt}*\n• Date: *{date}*",
        "btn_add_student": "➕ Add Student",
        "btn_add_teacher": "➕ Add Teacher",
        "btn_classes": "🏫 Classes & Students",
        "btn_search_student": "🔍 Search Student",
        "btn_excel": "📥 Import via Excel",
        "btn_cockpit": "📊 Morning Cockpit",
        "btn_pdf": "📄 Password Cards (PDF)",
        "btn_medical": "🏥 Medicals ({count})",
        "btn_requests": "🛎️ Access Requests ({count})",
        "btn_broadcast": "📢 Broadcast Notice",
        "btn_maintenance_toggle": "🚨 Maintenance ({status})",
        "btn_cafeteria_edit": "🍲 Update Cafeteria Menu",
        "btn_lang": "🌐 Change Language",
        "btn_back": "⬅️ Back",
        "btn_main_menu": "🏠 Main Menu",
        "btn_acknowledged": "✅ Read / Acknowledged",
        "acknowledged_toast": "Confirmation recorded.",
        "no_classes_found": "⚠️ No classes registered yet.",
        "prompt_student_name": "👤 Enter Student Full Name:",
        "prompt_student_class": "🏫 Enter Student Class (e.g. `9-A`):",
        "prompt_student_no": "🔢 Enter Student Roll Number (e.g. `101`):",
        "student_added_card": "✅ *Student Added Successfully!*\n\n👤 Name: *{name}*\n🏫 Class: *{class_name}* | Roll: *{no}*\n\n🔑 *Access Codes:*\n• Student Code: `{st_code}`\n• Parent Code: `{pr_code}`",
        "prompt_teacher_name": "👨‍🏫 Enter Teacher Full Name:",
        "prompt_teacher_subject": "📚 Enter Teaching Subject (e.g. `Mathematics`):",
        "teacher_added_card": "✅ *Teacher Registered!*\n\n👤 Name: *{name}*\n📚 Subject: *{subject}*\n\n🔑 *Access Code:*\n`{code}`",
        "student_card": "👤 *Student Card*\nName: *{name}*\nClass: *{class_name}* | Roll: *{no}*\n\n🔑 *Code Status:*\n• Student: `{st_code}` ({st_status})\n• Parent: `{pr_code}` ({pr_status})",
        "btn_reset_codes": "🔄 Reset Codes",
        "btn_del_student": "❌ Delete Student",
        "btn_student_grades": "📊 Grades & Absences",
        "student_deleted": "🗑️ Student removed from system.",
        "codes_reset_done": "✅ Credentials regenerated!\n\n• Student: `{st_code}`\n• Parent: `{pr_code}`",
        "excel_info": "📥 *Import via Excel*\n\nSend a `.xlsx` spreadsheet.\nHeaders: `Ad Soyad` | `Sinif` | `Numara`",
        "excel_done": "✅ Processed! Added students: *{count}*\nCredentials attached.",
        "cockpit_report": "📊 *Morning Briefing ({date})*\n\n🏫 Total: *{total}* | ✅ Present: *{present}* | ❌ Absent: *{absent}*\n\n⚠️ *Pending Attendance Classes ({missing_cnt}):*\n{missing}",
        "select_pdf_class": "📄 Select class for cards:",
        "pdf_ready": "📄 Printable cards for *{class_name}* ready.",
        "pending_medical_title": "🏥 *Pending Medical Reports:*",
        "no_pending_medical": "✅ No pending medical excuses.",
        "medical_approved": "✅ Medical excuse approved.",
        "medical_rejected": "❌ Medical excuse rejected.",
        "prompt_broadcast": "📢 Enter announcement text:",
        "broadcast_success": "📢 Dispatched to *{count}* users.",
        "prompt_search_student": "🔍 Enter student name or roll number:",
        "search_no_results": "❌ No student found.",
        "search_results_title": "🔍 *Search Results:*",
        "menu_teacher": "👨‍🏫 *Teacher Dashboard*\nTeacher: *{name}* ({subject})",
        "btn_attendance": "📋 Fast Attendance",
        "btn_enter_grade": "📝 Grade Book",
        "btn_homework_board": "📢 Homework Board",
        "btn_teacher_appointments": "🤝 Parent Meetings ({count})",
        "attendance_intro": "📋 *Attendance: {class_name}*\nTap absent students and save:",
        "btn_save_att": "💾 Save Attendance",
        "att_saved": "✅ Recorded (15-min edit window started).",
        "prompt_grade_score": "Student: *{name}* ({class_name})\nEnter score (0-100):",
        "prompt_grade_badge": "Select performance badge:",
        "badge_praise": "🟢 Praise / Achievement",
        "badge_missing": "🟡 Missing / Needs Work",
        "badge_warning": "🔴 Discipline / Warning",
        "grade_saved_success": "✅ Grade sent to parent.",
        "prompt_hw_class": "Select class for homework:",
        "prompt_hw_content": "Provide description or send blackboard photo:",
        "hw_sent_success": "📢 Homework dispatched to *{class_name}*.",
        "menu_parent": "👨‍👩‍👧‍👦 *Parent Dashboard*\nStudent: *{name}* ({class_name})",
        "btn_report_card": "📊 Report Card",
        "btn_switch_student": "🧑‍🎓 Switch Student",
        "btn_upload_medical": "🏥 Submit Medical Note",
        "btn_book_appointment": "🤝 Teacher Meeting",
        "btn_view_schedule": "📅 Timetable",
        "btn_view_cafeteria": "🍲 Cafeteria Menu",
        "upload_med_prompt": "Please send a photo of the medical note:",
        "med_uploaded_success": "Report submitted to school administration.",
        "student_switched_success": "Active student: *{name}* ({class_name})",
        "menu_student": "🎓 *Student Dashboard*\nStudent: *{name}* ({class_name} - Roll: {no})",
        "no_grades": "No grades recorded yet.",
        "no_linked_student": "⚠️ No student linked to your account.",
        "evening_briefing_header": "🌙 *DAILY SUMMARY (18:30)*\nStudent: *{name}* ({class_name})\n\n📌 Attendance: *{att_status}*\n📝 Grades:\n{grades}",
        "logged_out": "🚪 Logged out successfully.",
        # Talep Sistemi
        "btn_req_access": "📩 Request Access",
        "req_role_select": "📩 *Request System Access*\n\nPlease select your role at school:",
        "req_name_prompt": "👤 Enter your Full Name:",
        "req_phone_prompt": "📱 Enter your contact phone number (e.g. `+1234567890`):",
        "req_details_teacher": "📚 Enter your teaching subject (e.g. `Mathematics`):",
        "req_details_student": "🏫 Enter your Class & Roll Number (e.g. `9-A / 101`):",
        "req_details_parent": "🧑‍🎓 Enter your child's Name, Class & Roll Number (e.g. `John Doe, 9-A, 101`):",
        "req_sent_success": "✅ *Access Request Submitted!*\n\nOnce reviewed by the school administration, you will receive confirmation and credentials right here.",
        "req_already_pending": "⚠️ You already have an active request pending review. Please wait.",
        "req_cooldown": "⚠️ Your previous request was rejected. You can re-apply after 24 hours.",
        "req_approved_user": "🎉 *Congratulations! School Administration approved your request.*\n\nYour role: *{role}*\nAll features are now unlocked.",
        "req_rejected_user": "❌ *Your access request was rejected by administration.*\n\nPlease contact the school directly.",
        "no_pending_requests": "✅ No pending registration requests.",
        "pending_requests_title": "🛎️ *Pending Access Requests:*",
        # Randevu ve Menü
        "select_teacher_appointment": "🤝 Select a teacher to book a meeting:",
        "prompt_appointment_note": "Please write the topic and your preferred time slot:",
        "appointment_sent": "✅ Meeting request sent to the teacher.",
        "appointment_approved_msg": "✅ The teacher approved your meeting request!",
        "appointment_rejected_msg": "❌ The teacher is unavailable at the requested time.",
        "no_pending_appointments": "✅ No pending appointments.",
        "cafeteria_today_title": "🍲 *Cafeteria Menu for Today ({date}):*\n\n{content}",
        "schedule_title": "📅 *Timetable for Class {class_name}:*\n\n{content}",
        "prompt_menu_update": "🍲 Enter cafeteria menu for today:",
        "menu_updated": "✅ Cafeteria menu updated successfully.",
                "btn_add_child": "➕ Link Another Student",
        "prompt_add_child_code": "Enter the Parent Code for your other student (e.g. `VELI-123456`):",
        "child_added_success": "✅ New student linked to your account:\n👤 *{name}* ({class_name})",
        "btn_recent_grades": "📝 Recent Grades",
        "recent_grades_title": "📝 *Recently Recorded Grades:*\nTap to edit or delete:",
        "no_recent_grades": "✅ No grades entered yet.",
        "grade_detail_card": "📝 *Grade Detail*\nStudent: *{name}* ({class_name})\nSubject: *{subject}*\nScore: *{score}* ({badge})\nDate: {date}",
        "btn_edit_grade": "✏️ Edit Score",
        "btn_del_grade": "❌ Delete Grade",
        "grade_deleted": "🗑️ Grade deleted.",
        "prompt_new_score": "Enter new score (0-100):",
        "grade_updated": "✅ Grade updated successfully.",
        "absence_critical_warning": "🚨 *CRITICAL ABSENCE ALERT*\n\nTotal absences for student *{name}* reached *{count} days*!\n(Legal limit is 10 days. Please contact administration.)",
        "btn_profile": "👤 My Profile",
        "profile_title": "👤 *User Profile*\n\nName: *{name}*\nRole: *{role}*\nTelegram ID: `{id}`\n{details}",
        "help_text": "📖 *School System User Guide*\n\nYour role: *{role}*\n\n• Use the bottom keyboard menu to navigate.\n• To cancel any active form, tap '❌ Cancel Action' or write /cancel.",
        "btn_risk_radar": "⚠️ At-Risk Student Radar",
        "risk_radar_title": "⚠️ *At-Risk Student Radar (Absences & Low Scores):*",
        "no_risky_students": "✅ No at-risk students currently identified.",
        "btn_audit_logs": "📜 Audit & Activity Log",
        "audit_logs_title": "📜 *System Audit Log:*",
        "no_audit_logs": "✅ Audit log is empty.",
        "btn_academic_report": "📈 Academic Ranking",
        "academic_report_title": "📈 *Academic Ranking Across Classes:*",
        "btn_att_check": "📋 Attendance Compliance",
        "att_check_title": "📋 *Today's Attendance Compliance ({date}):*",
        "att_check_all_done": "✅ All classes have submitted attendance today.",
        "btn_remind_att": "🔔 Remind Teachers",
        "remind_att_sent": "📢 Attendance reminders sent to teachers.",
        "teacher_att_reminder_msg": "⚠️ *ATTENDANCE REMINDER*\n\nDear Teacher, attendance for class *{class_name}* has not been recorded yet. Please submit attendance.",
        "morning_briefing_header": "☀️ *EXECUTIVE BRIEFING FOR PRINCIPAL (09:30)*\n\n🏫 Total: *{total}* | ✅ Present: *{present}* (%{pct})\n❌ Absent: *{absent}* | 🏥 Excused: *{excused}*\n\n⚠️ *Missing Attendance ({missing_cnt}):*\n{missing}\n\n🛎️ Requests: *{req_cnt}*, Medicals: *{med_cnt}*",
                "rk_cat_staff": "👥 Staff & Students",
        "rk_cat_reports": "📊 Reports & Audits",
        "rk_cat_requests": "🛎️ Approval Center",
        "rk_cat_tools": "🛠️ Admin Tools",
        "rk_cat_settings": "⚙️ System & Settings",
        "cat_staff_title": "👥 *Staff & Student Management*\nPlease select an option:",
        "cat_reports_title": "📊 *Academic & Attendance Audit Hub*\nPlease select a report:",
        "cat_requests_title": "🛎️ *Requests & Medical Approvals*\nPlease select an option:",
        "cat_tools_title": "🛠️ *Administrative Tools & Communications*\nPlease select an option:",
        "cat_settings_title": "⚙️ *System Settings & Security*\nPlease select an option:",
                "rk_parent_info": "ℹ️ School Info Board",
        "rk_parent_settings": "⚙️ Settings & Exit",
        "btn_excel_hub": "📥 Excel Center (Import / Export)",
        "excel_hub_title": "📥 *Excel Operations Hub*\n\nPlease select an action:",
        "parent_info_title": "ℹ️ *School Information Board*\n\nPlease select a section:",
        "parent_settings_title": "⚙️ *Parent Account & Notification Settings*\n\nPlease select an option:",
        "btn_my_hws": "📚 Published Homeworks",
        "no_hws_found": "✅ No published homeworks found.",
        "btn_cockpit_unified": "📊 Morning Cockpit & Attendance",
        "btn_teachers": "👨‍🏫 Teachers",
        "btn_export_all": "📥 Full Data Backup (Excel)",
        "btn_notices": "📢 School Notices",
        "btn_blacklist": "🚫 Blocked Users",
        "btn_edit_student": "✏️ Edit Info",
        "btn_briefing_on": "🔔 Daily Summary: On",
        "btn_briefing_off": "🔕 Daily Summary: Off",
        "teacher_card": "👨‍🏫 *Teacher Card*\nName: *{name}*\nSubject: *{subject}*\n🔑 Access Code: `{code}`\n📱 Telegram: {status}",
        "teacher_deleted": "🗑️ Teacher removed from system.",
        "tch_code_reset": "✅ Teacher credentials regenerated:\n`{code}`",
        "student_updated": "✅ Student record updated.",
        "unban_success": "✅ User unblocked successfully.",
        "no_notices": "📢 No notices published yet.",
        "notices_title": "📢 *School Notices Board:*",
        "no_blacklisted": "✅ No blocked users.",
        "blacklisted_title": "🚫 *Blocked Accounts:*",
        "export_ready": "📥 *Complete School Data Backup* ({date})\n(Students, Teachers, Attendance, Grades)",
        "prompt_edit_value": "Enter new value:",
        "rk_notices": "📢 School Notices",
        "rk_teachers": "👨‍🏫 Teachers",
        # Reply Keyboard
        "rk_cancel_action": "❌ Cancel Action",
        "rk_admin_dash": "⚡ Admin Dashboard",
        "rk_cockpit": "📊 Morning Cockpit",
        "rk_classes": "🏫 Classes & Students",
        "rk_requests": "🛎️ Requests & Reports",
        "rk_attendance": "📋 Fast Attendance",
        "rk_grade": "📝 Grade Book",
        "rk_homework": "📢 Homework Board",
        "rk_appointments": "🤝 Parent Meetings",
        "rk_report": "📊 Report Card",
        "rk_schedule": "📅 Timetable",
        "rk_cafeteria": "🍲 Cafeteria Menu",
        "rk_switch_student": "🧑‍🎓 Switch Student",
        "rk_upload_medical": "🏥 Medical Note",
        "rk_main_menu": "🏠 Main Menu",
        "rk_lang": "🌐 Change Language",
        "rk_logout": "🚪 Log Out",
        "rk_enter_code": "🔑 Enter Access Code",
        "rk_req_access": "📩 Request Access"
    }
}

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
    """
    Sadeleştirilmiş, gözü yormayan kalıcı alt menü (Reply Keyboard).
    Maksimum 3 satır, derli toplu ve en kritik işlemlere anında erişim.
    """
    keyboard = []
    if role == "admin":
        keyboard = [
            [KeyboardButton(text=get_text("rk_cat_staff", lang)), KeyboardButton(text=get_text("rk_cat_reports", lang))],
            [KeyboardButton(text=get_text("rk_cat_requests", lang)), KeyboardButton(text=get_text("rk_cat_tools", lang))],
            [KeyboardButton(text=get_text("rk_cat_settings", lang)), KeyboardButton(text=get_text("rk_main_menu", lang))]
        ]
    elif role == "teacher":
        keyboard = [
            [KeyboardButton(text=get_text("rk_attendance", lang)), KeyboardButton(text=get_text("rk_grade", lang))],
            [KeyboardButton(text=get_text("rk_homework", lang)), KeyboardButton(text=get_text("rk_appointments", lang))],
            [KeyboardButton(text=get_text("rk_schedule", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "parent":
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_upload_medical", lang))],
            [KeyboardButton(text=get_text("rk_appointments", lang)), KeyboardButton(text=get_text("rk_parent_info", lang))],
            [KeyboardButton(text=get_text("rk_switch_student", lang)), KeyboardButton(text=get_text("rk_parent_settings", lang))]
        ]
    elif role == "student":
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_homework", lang))],
            [KeyboardButton(text=get_text("rk_schedule", lang)), KeyboardButton(text=get_text("rk_parent_info", lang))],
            [KeyboardButton(text=get_text("rk_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    else: # guest / giriş yapılmamış
        keyboard = [
            [KeyboardButton(text=get_text("rk_enter_code", lang)), KeyboardButton(text=get_text("rk_req_access", lang))],
            [KeyboardButton(text=get_text("rk_lang", lang))]
        ]

    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True, is_persistent=True)

def get_cancel_reply_kb(lang: str = "tr") -> ReplyKeyboardMarkup:
    """Aktif bir form veya işlem sırasında sadece İptal seçeneği sunarak kullanıcının kilitli kalmasını sağlar"""
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

def get_admin_main_inline_kb(lang: str = "tr", med_count: int = 0, req_count: int = 0, is_maintenance: bool = False) -> InlineKeyboardMarkup:
    """Temiz, düzenli ve 2'li bloklar halinde Yönetici Masası"""
    med_text = get_text("btn_medical", lang, count=med_count)
    req_text = get_text("btn_requests", lang, count=req_count)
    maint_status = "AÇIK" if is_maintenance else "KAPALI"
    if lang == "ru":
        maint_status = "ВКЛ" if is_maintenance else "ВЫКЛ"
    elif lang == "uz":
        maint_status = "YOQIK" if is_maintenance else "OCHIQ"
    elif lang == "en":
        maint_status = "ON" if is_maintenance else "OFF"

    maint_text = get_text("btn_maintenance_toggle", lang, status=maint_status)

    buttons = [
        [
            InlineKeyboardButton(text=get_text("btn_cockpit", lang), callback_data="adm:cockpit"),
            InlineKeyboardButton(text=get_text("btn_att_check", lang), callback_data="adm:teacher_attendance_check")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_risk_radar", lang), callback_data="adm:risk_radar"),
            InlineKeyboardButton(text=get_text("btn_academic_report", lang), callback_data="adm:academic_report")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"),
            InlineKeyboardButton(text=get_text("btn_teachers", lang), callback_data="adm:teachers")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_search_student", lang), callback_data="adm:search_student"),
            InlineKeyboardButton(text=req_text, callback_data="adm:requests_list")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_add_student", lang), callback_data="adm:add_student"),
            InlineKeyboardButton(text=get_text("btn_add_teacher", lang), callback_data="adm:add_teacher")
        ],
        [
            InlineKeyboardButton(text=med_text, callback_data="adm:medical_list"),
            InlineKeyboardButton(text=get_text("btn_audit_logs", lang), callback_data="adm:audit_logs")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_export_all", lang), callback_data="adm:export_all_excel"),
            InlineKeyboardButton(text=get_text("btn_pdf", lang), callback_data="adm:pdf_menu")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_excel", lang), callback_data="adm:excel_info"),
            InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_init")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_blacklist", lang), callback_data="adm:blacklist"),
            InlineKeyboardButton(text=get_text("btn_cafeteria_edit", lang), callback_data="adm:menu_edit")
        ],
        [
            InlineKeyboardButton(text=maint_text, callback_data="adm:toggle_maint"),
            InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")
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
        [InlineKeyboardButton(text=get_text("btn_acknowledged", lang), callback_data=f"ack_notif:{notification_id}")]
    ])

# ======================================================================
# 5. YARDIMCI SERVİSLER (KOD ÜRETİMİ, EXCEL, PDF VE ARKA PLAN İŞÇİLERİ)
# ======================================================================

def normalize_code(code_str: str) -> str:
    """Türkçe karakterleri ve boşlukları düzelterek kodu standartlaştırır"""
    if not code_str:
        return ""
    cleaned = code_str.strip().replace("ı", "I").replace("i", "I").replace("İ", "I").upper()
    cleaned = re.sub(r'[\s\-]+', '-', cleaned)
    return cleaned

async def log_audit(session: AsyncSession, user_id: int, user_name: str, action: str, details: str):
    """Sistemdeki tüm kritik idari ve akademik işlemleri şeffaf şekilde günlüğe kaydeder"""
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
    """Kullanıcının yönetici yetkisini tam ve kesin olarak doğrular"""
    if user_id in ADMIN_IDS:
        return True
    return bool(user and user.role == "admin")

def generate_secure_code(prefix: str) -> str:
    """6 haneli, yüksek entropili güvenli kod üretir"""
    num_part = str(random.randint(100000, 999999))
    return f"{prefix}-{num_part}"

async def export_all_school_data_excel() -> io.BytesIO:
    """Tüm okul verilerini (öğrenciler, öğretmenler, yoklamalar, notlar) çok sayfalı Excel olarak üretir"""
    async with AsyncSessionLocal() as session:
        students = (await session.execute(select(Student).order_by(Student.class_name, Student.student_number))).scalars().all()
        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        attendances = (await session.execute(select(Attendance).order_by(desc(Attendance.date)))).scalars().all()
        grades = (await session.execute(select(Grade).order_by(desc(Grade.created_at)))).scalars().all()

    wb = openpyxl.Workbook()

    # Sayfa 1: Öğrenciler
    ws1 = wb.active
    ws1.title = "Ogrenciler"
    ws1.append(["ID", "Ad Soyad", "Sinif", "Okul Numarasi", "Ogrenci Kodu", "Veli Kodu", "Ogrenci Telegram ID"])
    for s in students:
        ws1.append([s.id, s.full_name, s.class_name, s.student_number, s.student_code, s.parent_code, s.student_telegram_id or "-"])

    # Sayfa 2: Öğretmenler
    ws2 = wb.create_sheet(title="Ogretmenler")
    ws2.append(["ID", "Ad Soyad", "Brans", "Giris Kodu", "Telegram ID"])
    for t in teachers:
        ws2.append([t.id, t.full_name, t.subject, t.auth_code, t.telegram_id or "-"])

    # Sayfa 3: Yoklamalar
    ws3 = wb.create_sheet(title="Yoklamalar")
    ws3.append(["Tarih", "Sinif", "Ogrenci ID", "Durum"])
    for a in attendances:
        ws3.append([str(a.date), a.class_name, a.student_id, a.status])

    # Sayfa 4: Notlar
    ws4 = wb.create_sheet(title="Ders Notlari")
    ws4.append(["Tarih", "Ogrenci ID", "Ders", "Not", "Rozet", "Not"])
    for g in grades:
        ws4.append([g.created_at.strftime("%d.%m.%Y"), g.student_id, g.subject, g.score, g.badge, g.note or ""])

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
        for row in rows:
            if not row or not row[0]:
                continue
            full_name = str(row[0]).strip()
            class_name = str(row[1]).strip() if len(row) > 1 and row[1] else "Genel"
            student_no = str(row[2]).strip() if len(row) > 2 and row[2] else "0"

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

async def generate_classroom_pdf_cards(class_name: str, lang: str = "tr") -> io.BytesIO:
    async with AsyncSessionLocal() as session:
        stmt = select(Student).where(Student.class_name == class_name).order_by(Student.student_number)
        result = await session.execute(stmt)
        students = result.scalars().all()

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=25,
        leftMargin=25,
        topMargin=25,
        bottomMargin=25
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="CardTitle", parent=styles["Heading2"], fontSize=13, leading=16, textColor=colors.HexColor("#1A365D"), alignment=1
    )
    body_style = ParagraphStyle(
        name="CardBody", parent=styles["Normal"], fontSize=8.5, leading=12, textColor=colors.HexColor("#2D3748")
    )

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
            f"<b>{lbl_st}:</b> {s.full_name} (№ {s.student_number})<br/>"
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

async def run_attendance_delay_worker(bot: Bot):
    async with AsyncSessionLocal() as session:
        now = datetime.utcnow()
        stmt = (
            select(Attendance, Student)
            .join(Student, Attendance.student_id == Student.id)
            .where(Attendance.notify_at <= now, Attendance.is_notified == False)
        )
        result = await session.execute(stmt)
        records = result.all()

        for att, st in records:
            if att.status == "absent":
                p_stmt = select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st.id)
                parents = (await session.execute(p_stmt)).scalars().all()

                for p_id in parents:
                    u_stmt = select(User).where(User.telegram_id == p_id)
                    user = (await session.execute(u_stmt)).scalar_one_or_none()
                    lang = user.language if user else "tr"

                    msg_text = {
                        "tr": f"🚨 *DEVAMSIZLIK BİLDİRİMİ*\n\nÖğrenciniz *{st.full_name}*, bugün ({att.date}) okul yoklamasında *GELMEDİ (YOK)* olarak kaydedilmiştir.",
                        "ru": f"🚨 *УВЕДОМЛЕНИЕ О ПРОПУСКЕ*\n\nВаш ребенок *{st.full_name}* сегодня ({att.date}) отмечен(а) как *ОТСУТСТВУЕТ*.",
                        "uz": f"🚨 *DAVOMAT OGOHLANTIRISHI*\n\nFarzandingiz *{st.full_name}* bugun ({att.date}) maktab davomatida *KELMADI* deb qayd etildi.",
                        "en": f"🚨 *ABSENCE ALERT*\n\nYour student *{st.full_name}* has been marked *ABSENT* today ({att.date})."
                    }.get(lang, f"🚨 *ABSENCE ALERT*\n\n{st.full_name} is marked ABSENT.")

                    notif = CriticalNotification(user_telegram_id=p_id, message_text=msg_text)
                    session.add(notif)
                    await session.flush()

                    try:
                        await bot.send_message(
                            chat_id=p_id,
                            text=msg_text,
                            reply_markup=get_ack_notification_kb(notif.id, lang=lang),
                            parse_mode="Markdown"
                        )
                    except Exception:
                        pass

            att.is_notified = True

        await session.commit()

async def run_evening_briefing_worker(bot: Bot):
    """18:30 Gün Sonu Özeti: Devamsızlık ve notları veliye bülten olarak sunar"""
    async with AsyncSessionLocal() as session:
        now = datetime.utcnow()
        today = now.date()

        parents = (await session.execute(select(User).where(User.role == "parent", User.evening_briefing == True))).scalars().all()
        for p in parents:
            st_id = p.current_child_id
            if not st_id:
                continue
            st = await session.get(Student, st_id)
            if not st:
                continue

            lang = p.language
            att = (await session.execute(select(Attendance).where(Attendance.student_id == st.id, Attendance.date == today))).scalar_one_or_none()
            att_status = "✅ Geldi" if not att or att.status == "present" else ("❌ Gelmedi (Yok)" if att.status == "absent" else "🏥 İzinli")

            grades = (await session.execute(select(Grade).where(Grade.student_id == st.id, func.date(Grade.created_at) == today))).scalars().all()
            grades_str = "\n".join([f"• {g.subject}: *{g.score}* ({g.badge})" for g in grades]) if grades else get_text("no_grades", lang)

            text = get_text("evening_briefing_header", lang, name=st.full_name, class_name=st.class_name, att_status=att_status, grades=grades_str)
            try:
                await bot.send_message(chat_id=p.telegram_id, text=text, parse_mode="Markdown")
            except Exception:
                pass

# ======================================================================
# 6. FSM DURUMLARI (STATE MACHINE)
# ======================================================================

class Form(StatesGroup):
    waiting_auth_code = State()
    # Erişim Talebi Başvurusu Adımları
    req_role = State()
    req_name = State()
    req_phone = State()
    req_details = State()
    # Öğrenci & Öğretmen Ekleme
    add_student_name = State()
    add_student_class = State()
    add_student_no = State()
    add_teacher_name = State()
    add_teacher_subject = State()
    # Arama, Duyuru ve Mazeret
    waiting_search_query = State()
    waiting_broadcast_text = State()
    waiting_medical_photo = State()
    # Not & Ödev
    grade_score = State()
    grade_badge = State()
    hw_content = State()
    # Randevu ve Yemekhane
    app_teacher = State()
    app_note = State()
    menu_update_text = State()
    # Öğrenci Bilgi Düzenleme
    edit_student_val = State()
    # Veli Çoklu Çocuk Ekleme & Öğretmen Not Düzenleme
    parent_add_child_code = State()
    edit_grade_val = State()

router = Router()
ATTENDANCE_CACHE = {}
GRADE_CACHE = {}
HW_CACHE = {}
REQ_CACHE = {}
APP_CACHE = {}

# ======================================================================
# 7. ROUTER: BAŞLANGIÇ, DİL VE ROLE ÖZEL KONTROL MASALARI
# ======================================================================

@router.message(any_state, Command("profil", "profile"))
async def cmd_profile(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not user: return
        lang = user.language

        role_label = {"admin": "Okul İdaresi (Admin)", "teacher": "Öğretmen", "parent": "Veli", "student": "Öğrenci", "guest": "Giriş Yapılmamış"}.get(user.role, user.role)
        extra_details = ""
        if user.role == "parent" and user.current_child_id:
            st = await session.get(Student, user.current_child_id)
            if st: extra_details = f"Öğrenci: *{st.full_name}* ({st.class_name})\n"
        elif user.role == "student":
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            if st: extra_details = f"Sınıf: *{st.class_name}* | No: *{st.student_number}*\n"
        elif user.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
            if tch: extra_details = f"Branş: *{tch.subject}*\n"

        text = get_text("profile_title", lang, name=user.full_name or "Kullanıcı", role=role_label, id=user.telegram_id, details=extra_details)
        await message.answer(text, parse_mode="Markdown")

@router.message(any_state, Command("help", "yardim"))
async def cmd_help(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        role_label = {"admin": "Yönetici", "teacher": "Öğretmen", "parent": "Veli", "student": "Öğrenci"}.get(user.role if user else "guest", "Kullanıcı")
        await message.answer(get_text("help_text", lang, role=role_label), parse_mode="Markdown")

@router.message(any_state, Command("cancel", "iptal"))
@router.message(any_state, F.text.in_(["❌ İşlemi İptal Et", "❌ Отменить действие", "❌ Bekor qilish", "❌ Cancel Action", "iptal", "İptal", "cancel"]))
async def cmd_cancel(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if user:
            reply_kb = get_role_reply_kb(user.role, user.language)
            await message.answer("❌ *İşlem iptal edildi.*", reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)

@router.message(CommandStart())
@router.message(Command("menu"))
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        is_admin_id = message.from_user.id in ADMIN_IDS

        if not user:
            user = User(
                telegram_id=message.from_user.id,
                role="admin" if is_admin_id else "guest",
                language="tr"
            )
            session.add(user)
            await session.commit()

            await message.answer(
                "🌍 Iltimos, tilni tanlang / Пожалуйста, выберите язык / Lütfen dil seçiniz / Please select language:",
                reply_markup=get_language_inline_kb()
            )
            return

        if is_admin_id:
            if user.role != "admin":
                user.role = "admin"
                await session.commit()
            reply_kb = get_role_reply_kb("admin", user.language)
            await message.answer(
                f"👑 *Yönetici Girişi Yapıldı* (ID: `{message.from_user.id}`)",
                reply_markup=reply_kb,
                parse_mode="Markdown"
            )
            await render_clean_dashboard(message, user)
            return

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language))
            return

        if user.locked_until and user.locked_until > datetime.utcnow():
            await message.answer(get_text("auth_locked", user.language))
            return

        # Bakım modu kontrolü (Admin hariç)
        if user.role != "admin":
            maint = await session.get(SystemSetting, "maintenance_mode")
            if maint and maint.value == "true":
                await message.answer(get_text("maintenance_mode", user.language))
                return

        # Eğer misafir ise sade karşılama ve alt menüyü göster
        if user.role == "guest":
            await prompt_guest_screen(message, user, state)
            return

        reply_kb = get_role_reply_kb(user.role, user.language)
        await message.answer(
            f"👋 *Hoş Geldiniz, {user.full_name or ''}*",
            reply_markup=reply_kb,
            parse_mode="Markdown"
        )
        await render_clean_dashboard(message, user)

async def prompt_guest_screen(target: Message | CallbackQuery, user: User, state: FSMContext):
    """
    Misafire TEK ve net bir mesaj gönderir.
    Ekranı kirleten mükerrer inline butonlar ve çift mesaj balonu tamamen kaldırılmıştır.
    İşlemler doğrudan alt menü (Reply Keyboard) üzerinden seçilir veya şifre doğrudan yazılır.
    """
    lang = user.language
    text = get_text("welcome_guest", lang)
    reply_kb = get_role_reply_kb("guest", lang)

    if isinstance(target, CallbackQuery):
        try:
            await target.message.delete()
        except Exception:
            pass
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
    """Sohbeti kirletmeden TEK BİR MESAJDA mükemmel in-place kontrol paneli çizer"""
    lang = user.language

    if user.role == "admin":
        async with AsyncSessionLocal() as session:
            c_cnt = (await session.execute(select(func.count(Student.class_name.distinct())))).scalar() or 0
            s_cnt = (await session.execute(select(func.count(Student.id)))).scalar() or 0
            t_cnt = (await session.execute(select(func.count(Teacher.id)))).scalar() or 0
            med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0
            req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
            maint_setting = await session.get(SystemSetting, "maintenance_mode")
            is_maint = maint_setting.value == "true" if maint_setting else False
            date_str = datetime.utcnow().strftime("%d.%m.%Y")

            text = (
                f"{get_text('admin_title', lang)}\n\n"
                f"{get_text('admin_stats', lang, c_cnt=c_cnt, s_cnt=s_cnt, t_cnt=t_cnt, req_cnt=req_cnt, med_cnt=med_cnt, date=date_str)}"
            )
            kb = None

    elif user.role == "teacher":
        async with AsyncSessionLocal() as session:
            tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
            subj = tch.subject if tch else "Ders"
            tch_name = tch.full_name if tch else (user.full_name or "Öğretmen")
            tch_id = tch.id if tch else 0
            app_cnt = (await session.execute(select(func.count(Appointment.id)).where(Appointment.teacher_id == tch_id, Appointment.status == "pending"))).scalar() or 0

            text = get_text("menu_teacher", lang, name=tch_name, subject=subj)
            buttons = [
                [
                    InlineKeyboardButton(text=get_text("btn_attendance", lang), callback_data="tch:classes"),
                    InlineKeyboardButton(text=get_text("btn_enter_grade", lang), callback_data="tch:grade_classes")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_recent_grades", lang), callback_data="tch:recent_grades"),
                    InlineKeyboardButton(text=get_text("btn_homework_board", lang), callback_data="tch:hw_classes")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_teacher_appointments", lang, count=app_cnt), callback_data="tch:appointments")
                ]
            ]
            kb = None

    elif user.role == "parent":
        async with AsyncSessionLocal() as session:
            st = await session.get(Student, user.current_child_id) if user.current_child_id else None
            name = st.full_name if st else "-"
            cls_name = st.class_name if st else "-"
            briefing_txt = get_text("btn_briefing_on", lang) if user.evening_briefing else get_text("btn_briefing_off", lang)
            text = get_text("menu_parent", lang, name=name, class_name=cls_name)
            buttons = [
                [
                    InlineKeyboardButton(text=get_text("btn_report_card", lang), callback_data="act_view_report"),
                    InlineKeyboardButton(text=get_text("btn_book_appointment", lang), callback_data="act_book_app")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_view_schedule", lang), callback_data="act_view_sched"),
                    InlineKeyboardButton(text=get_text("btn_notices", lang), callback_data="act_view_notices")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_view_cafeteria", lang), callback_data="act_view_cafe"),
                    InlineKeyboardButton(text=get_text("btn_upload_medical", lang), callback_data="upload_medical_init")
                ],
                [
                    InlineKeyboardButton(text=briefing_txt, callback_data="parent:toggle_briefing"),
                    InlineKeyboardButton(text=get_text("btn_switch_student", lang), callback_data="parent:switch_student")
                ]
            ]
            kb = None

    elif user.role == "student":
        async with AsyncSessionLocal() as session:
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            st_name = st.full_name if st else (user.full_name or "Öğrenci")
            cls_name = st.class_name if st else "-"
            num_val = st.student_number if st else "-"
            text = get_text("menu_student", lang, name=st_name, class_name=cls_name, no=num_val)
            buttons = [
                [
                    InlineKeyboardButton(text=get_text("btn_student_grades", lang), callback_data="act_view_report"),
                    InlineKeyboardButton(text=get_text("btn_view_schedule", lang), callback_data="act_view_sched")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_view_cafeteria", lang), callback_data="act_view_cafe")
                ]
            ]
            kb = None

    else: # guest
        text = get_text("welcome_guest", lang)
        kb = None

    if isinstance(target, Bot):
        target_id = chat_id or user.telegram_id
        await target.send_message(chat_id=target_id, text=text, reply_markup=kb, parse_mode="Markdown")
        return

    if isinstance(target, CallbackQuery):
        try:
            if target.message and target.message.photo:
                await target.message.delete()
                await target.message.answer(text, reply_markup=kb, parse_mode="Markdown")
            else:
                await target.message.edit_text(text, reply_markup=kb, parse_mode="Markdown")
        except Exception:
            await target.message.answer(text, reply_markup=kb, parse_mode="Markdown")
    else:
        await target.answer(text, reply_markup=kb, parse_mode="Markdown")

# ======================================================================
# 8. KİMLİK DOĞRULAMA (AUTHENTICATION ENGINE - TEK KAPIDAN TÜM ROLLER)
# ======================================================================

async def process_auth_code_string(code: str, user_id: int, message: Message, state: FSMContext):
    clean_code = normalize_code(code)

    try:
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
                await message.answer(get_text("auth_locked", lang))
                return

            # 1. YÖNETİCİ ŞİFRESİ KONTROLÜ
            if clean_code == ADMIN_CODE or user_id in ADMIN_IDS:
                user.role = "admin"
                user.failed_attempts = 0
                user.locked_until = None
                await session.commit()
                await state.clear()

                reply_kb = get_role_reply_kb("admin", lang)
                success_txt = get_text("auth_success", lang, name="Yönetici", role="Okul İdaresi (Admin)")
                await message.answer(success_txt, reply_markup=reply_kb)
                await render_clean_dashboard(message, user)
                return

            # 2. ÖĞRETMEN GİRİŞ KODU KONTROLÜ (HCA-XXXXXX)
            t_stmt = select(Teacher).where(func.upper(func.trim(Teacher.auth_code)) == clean_code)
            teacher = (await session.execute(t_stmt)).scalar_one_or_none()
            if teacher:
                teacher.telegram_id = user_id
                teacher.is_code_burned = True
                user.role = "teacher"
                user.full_name = teacher.full_name
                user.failed_attempts = 0
                user.locked_until = None
                await session.commit()
                await state.clear()

                reply_kb = get_role_reply_kb("teacher", lang)
                success_txt = get_text("auth_success", lang, name=teacher.full_name, role=f"Öğretmen ({teacher.subject})")
                await message.answer(success_txt, reply_markup=reply_kb)
                await render_clean_dashboard(message, user)
                return

            # 3. ÖĞRENCİ GİRİŞ KODU KONTROLÜ (OGR-XXXXXX)
            s_stmt = select(Student).where(func.upper(func.trim(Student.student_code)) == clean_code)
            student = (await session.execute(s_stmt)).scalar_one_or_none()
            if student:
                student.student_telegram_id = user_id
                student.is_student_code_burned = True
                user.role = "student"
                user.full_name = student.full_name
                user.failed_attempts = 0
                user.locked_until = None
                await session.commit()
                await state.clear()

                reply_kb = get_role_reply_kb("student", lang)
                success_txt = get_text("auth_success", lang, name=student.full_name, role=f"Öğrenci ({student.class_name})")
                await message.answer(success_txt, reply_markup=reply_kb)
                await render_clean_dashboard(message, user)
                return

            # 4. VELİ GİRİŞ KODU KONTROLÜ (VELI-XXXXXX)
            p_stmt = select(Student).where(func.upper(func.trim(Student.parent_code)) == clean_code)
            p_student = (await session.execute(p_stmt)).scalar_one_or_none()
            if p_student:
                p_student.is_parent_code_burned = True
                user.role = "parent"
                user.failed_attempts = 0
                user.locked_until = None
                user.current_child_id = p_student.id

                rel_check = select(ParentStudent).where(
                    ParentStudent.parent_telegram_id == user_id,
                    ParentStudent.student_id == p_student.id
                )
                exists = (await session.execute(rel_check)).scalar_one_or_none()
                if not exists:
                    session.add(ParentStudent(parent_telegram_id=user_id, student_id=p_student.id))

                await session.commit()
                await state.clear()

                reply_kb = get_role_reply_kb("parent", lang)
                success_txt = get_text("auth_success", lang, name=f"{p_student.full_name} Velisi", role=f"Veli ({p_student.class_name})")
                await message.answer(success_txt, reply_markup=reply_kb)
                await render_clean_dashboard(message, user)
                return

            # BAŞARISIZ GİRİŞ
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

            remaining = 3 - user.failed_attempts
            await session.commit()
            await message.answer(get_text("auth_failed", lang, remaining=remaining))
    except Exception as e:
        import traceback
        traceback.print_exc()
        await message.answer(f"⚠️ Kod işlenirken beklenmeyen bir hata oluştu: {type(e).__name__}. Lütfen tekrar deneyiniz.")

@router.message(Form.waiting_auth_code)
async def handle_auth_code_fsm(message: Message, state: FSMContext):
    await process_auth_code_string(message.text, message.from_user.id, message, state)

@router.callback_query(F.data == "act_enter_code")
async def cb_act_enter_code(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    await query.message.answer(
        "🔑 *Lütfen size verilen giriş kodunu yazınız:* (Örn: `HCA-123456`, `VELI-123456`, `OGR-123456`)",
        parse_mode="Markdown"
    )
    await state.set_state(Form.waiting_auth_code)
    await query.answer()

# ======================================================================
# 9. YÖNETİCİDEN ERİŞİM / ŞİFRE TALEP ETME SİSTEMİ (SECURITY SHIELD)
# ======================================================================

@router.callback_query(F.data == "act_req_access")
async def cb_act_req_access(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        # 1. Bekleyen aktif talep var mı? (Spam Kalkanı)
        pending = (await session.execute(
            select(AccessRequest).where(AccessRequest.telegram_id == query.from_user.id, AccessRequest.status == "pending")
        )).scalar_one_or_none()
        if pending:
            await query.answer(get_text("req_already_pending", lang), show_alert=True)
            return

        # 2. Son 24 saatte reddedilen talep var mı? (Cooldown Kalkanı)
        yesterday = datetime.utcnow() - timedelta(hours=24)
        rejected = (await session.execute(
            select(AccessRequest).where(
                AccessRequest.telegram_id == query.from_user.id,
                AccessRequest.status == "rejected",
                AccessRequest.reviewed_at >= yesterday
            )
        )).scalar_one_or_none()
        if rejected:
            await query.answer(get_text("req_cooldown", lang), show_alert=True)
            return

        buttons = [
            [
                InlineKeyboardButton(text="👨‍🏫 Öğretmen", callback_data="req_role:teacher"),
                InlineKeyboardButton(text="👨‍👩‍👧‍👦 Veli", callback_data="req_role:parent")
            ],
            [
                InlineKeyboardButton(text="🎓 Öğrenci", callback_data="req_role:student")
            ],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await query.message.edit_text(get_text("req_role_select", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
    full_name = message.text.strip()
    if len(full_name) < 3:
        await message.answer("⚠️ *Hatalı İsim:* Lütfen adınızı ve soyadınızı eksiksiz yazınız (en az 3 karakter):", parse_mode="Markdown")
        return
    await state.update_data(full_name=full_name)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    await message.answer(
        get_text("req_phone_prompt", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="act_req_access")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.req_phone)

@router.message(Form.req_phone)
async def process_req_phone(message: Message, state: FSMContext):
    phone_val = message.text.strip()
    digits_only = re.sub(r'\D', '', phone_val)
    if len(digits_only) < 7:
        await message.answer("⚠️ *Hatalı Telefon:* Lütfen geçerli bir telefon numarası giriniz (en az 7 rakam):", parse_mode="Markdown")
        return
    await state.update_data(phone=phone_val)

    data = await state.get_data()
    role = data.get("role", "student")

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        prompt_key = "req_details_teacher" if role == "teacher" else ("req_details_parent" if role == "parent" else "req_details_student")
        await message.answer(
            get_text(prompt_key, lang),
            reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="act_req_access")]]),
            parse_mode="Markdown"
        )
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

        # Veritabanında otomatik teyit araması
        student_match = None
        clean_det = details_text.lower()
        if role in ["parent", "student"]:
            all_st = (await session.execute(select(Student))).scalars().all()
            for s in all_st:
                if s.student_number in clean_det or s.full_name.lower() in clean_det:
                    student_match = s
                    break

        # Talebi veritabanına kaydet
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

        # Kullanıcıya başarı mesajı
        await message.answer(get_text("req_sent_success", lang), parse_mode="Markdown")
        await prompt_guest_screen(message, user, state)

        # Tüm yöneticilere Telegram üzerinden anlık onay kartı gönder
        role_label = {"teacher": "Öğretmen", "parent": "Veli", "student": "Öğrenci"}.get(role, role)
        auto_check_badge = f"🟢 *Sistemde Doğrulandı:* {student_match.full_name} ({student_match.class_name} - No: {student_match.student_number})" if student_match else "🔴 *Otomatik Eşleşme Bulunamadı (Manuel İnceleyiniz)*"

        adm_msg = (
            f"🛎️ *YENİ ERİŞİM & ŞİFRE TALEBİ* (#{req.id})\n\n"
            f"👤 *Başvuran:* {full_name}\n"
            f"🆔 *Telegram:* @{message.from_user.username or 'gizli'} (`{message.from_user.id}`)\n"
            f"📱 *Telefon:* `{phone_val}`\n"
            f"🎯 *Talep Rolü:* *{role_label}*\n"
            f"📝 *Açıklama / Bilgi:* {details_text}\n\n"
            f"🔍 *Sistem Teyidi:*\n{auto_check_badge}"
        )

        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Onayla ve Yetkilendir", callback_data=f"adm:appr_req:{req.id}"),
                InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_req:{req.id}")
            ]
        ])

        admins = (await session.execute(select(User.telegram_id).where(User.role == "admin"))).scalars().all()
        target_admin_ids = set(ADMIN_IDS + list(admins))
        for a_id in target_admin_ids:
            try:
                await message.bot.send_message(chat_id=a_id, text=adm_msg, reply_markup=adm_kb, parse_mode="Markdown")
            except Exception:
                pass

# --- YÖNETİCİ: ERİŞİM TALEPLERİ LİSTESİ VE ONAY İŞLEMLERİ ---

@router.callback_query(F.data == "adm:requests_list")
async def cb_admin_requests_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        requests = (await session.execute(
            select(AccessRequest).where(AccessRequest.status == "pending").order_by(desc(AccessRequest.created_at))
        )).scalars().all()

        if not requests:
            buttons = [get_nav_buttons(lang)]
            await query.message.edit_text(get_text("no_pending_requests", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        for r in requests:
            role_icon = "👨‍🏫" if r.role == "teacher" else ("👨‍👩‍👧‍👦" if r.role == "parent" else "🎓")
            btn_text = f"{role_icon} {r.full_name} ({r.created_at.strftime('%H:%M')})"
            buttons.append([InlineKeyboardButton(text=btn_text, callback_data=f"adm:view_req:{r.id}")])

        buttons.append(get_nav_buttons(lang))
        await query.message.edit_text(get_text("pending_requests_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:view_req:"))
async def cb_admin_view_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer("Bu talep zaten sonuçlandırılmış!", show_alert=True)
            await cb_admin_requests_list(query)
            return

        role_label = {"teacher": "Öğretmen", "parent": "Veli", "student": "Öğrenci"}.get(req.role, req.role)
        st_match_info = ""
        if req.student_match_id:
            st = await session.get(Student, req.student_match_id)
            if st:
                st_match_info = f"\n🔍 *Eşleşen Öğrenci:* {st.full_name} ({st.class_name} - No: {st.student_number})"

        text = (
            f"🛎️ *Erişim Talebi Detayı* (#{req.id})\n\n"
            f"👤 *İsim:* {req.full_name}\n"
            f"📱 *Telefon:* `{req.phone}`\n"
            f"🎯 *Rol:* *{role_label}*\n"
            f"📝 *Detay:* {req.details}"
            f"{st_match_info}\n\n"
            f"Lütfen yapılacak işlemi seçiniz:"
        )

        buttons = [
            [
                InlineKeyboardButton(text="✅ Onayla ve Yetkilendir", callback_data=f"adm:appr_req:{req.id}"),
                InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_req:{req.id}")
            ],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:requests_list")]
        ]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:appr_req:"))
async def cb_admin_approve_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        is_admin = (admin_user and admin_user.role == "admin") or (query.from_user.id in ADMIN_IDS)
        if not is_admin:
            try:
                await query.answer("Yetkisiz işlem!", show_alert=True)
            except Exception:
                pass
            return

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer("Bu talep daha önce işleme alınmış!", show_alert=True)
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

        # Rol bazlı atama
        if req.role == "teacher":
            tch = (await session.execute(select(Teacher).where(Teacher.full_name == req.full_name))).scalar_one_or_none()
            if not tch:
                auth_c = generate_secure_code("HCA")
                tch = Teacher(full_name=req.full_name, subject=req.details, auth_code=auth_c, telegram_id=req.telegram_id, is_code_burned=True)
                session.add(tch)
            else:
                tch.telegram_id = req.telegram_id
                tch.is_code_burned = True

        elif req.role == "parent":
            st_id = req.student_match_id
            if not st_id:
                # İsimden tekrar ara
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

        # Kullanıcıya Telegram üzerinden anında bildirim ve alt menü ilet
        try:
            role_tr = {"teacher": "Öğretmen", "parent": "Veli", "student": "Öğrenci"}.get(req.role, req.role)
            user_msg = get_text("req_approved_user", target_user.language, role=role_tr)
            reply_kb = get_role_reply_kb(req.role, target_user.language)
            await query.message.bot.send_message(chat_id=req.telegram_id, text=user_msg, reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(query.message.bot, target_user, chat_id=req.telegram_id)
        except Exception:
            pass

        await query.message.edit_text(f"✅ *Talep Onaylandı (# {req.id})*\n{req.full_name} ({req.role}) başarıyla sisteme kaydedildi.", parse_mode="Markdown")
    await query.answer("Onaylandı!")

@router.callback_query(F.data.startswith("adm:rej_req:"))
async def cb_admin_reject_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        is_admin = (admin_user and admin_user.role == "admin") or (query.from_user.id in ADMIN_IDS)
        if not is_admin:
            try:
                await query.answer("Yetkisiz işlem!", show_alert=True)
            except Exception:
                pass
            return

        req = await session.get(AccessRequest, req_id)
        if not req or req.status != "pending":
            await query.answer("Bu talep daha önce işleme alınmış!", show_alert=True)
            return

        req.status = "rejected"
        req.reviewed_by = query.from_user.id
        req.reviewed_at = datetime.utcnow()
        await session.commit()

        # Kullanıcıya bilgilendirme gönder
        try:
            target_user = await session.get(User, req.telegram_id)
            lang = target_user.language if target_user else "tr"
            await query.message.bot.send_message(chat_id=req.telegram_id, text=get_text("req_rejected_user", lang), parse_mode="Markdown")
        except Exception:
            pass

        await query.message.edit_text(f"❌ *Talep Reddedildi (# {req.id})*\n{req.full_name} adlı kullanıcının başvurusu reddedildi.", parse_mode="Markdown")
    await query.answer("Reddedildi.")

# ======================================================================
# 10. VELİ - ÖĞRETMEN RANDEVU SİSTEMİ & DERS PROGRAMI / MENÜ
# ======================================================================

@router.callback_query(F.data == "act_book_app")
async def cb_parent_book_app_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        teachers = (await session.execute(select(Teacher).order_by(Teacher.full_name))).scalars().all()
        if not teachers:
            await query.answer("Kayıtlı öğretmen bulunamadı!", show_alert=True)
            return

        buttons = []
        for t in teachers:
            buttons.append([InlineKeyboardButton(text=f"👨‍🏫 {t.full_name} ({t.subject})", callback_data=f"book_tch:{t.id}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_add_child", lang), callback_data="parent:add_child_code")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])

        await query.message.edit_text(get_text("select_teacher_appointment", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data.startswith("book_tch:"))
async def cb_parent_select_teacher(query: CallbackQuery, state: FSMContext):
    tch_id = int(query.data.split(":")[1])
    APP_CACHE[query.from_user.id] = {"teacher_id": tch_id}

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    await query.message.edit_text(get_text("prompt_appointment_note", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
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

        # Öğretmene anlık kart gönder
        if tch and tch.telegram_id:
            try:
                st_name = st.full_name if st else "Öğrenci"
                cls_name = st.class_name if st else ""
                t_text = (
                    f"🤝 *YENİ VELİ GÖRÜŞME TALEBİ*\n\n"
                    f"👤 *Veli:* {user.full_name or 'Veli'}\n"
                    f"🧑‍🎓 *Öğrenci:* {st_name} ({cls_name})\n"
                    f"📝 *Talep / Zaman:* {note_text}"
                )
                t_kb = InlineKeyboardMarkup(inline_keyboard=[
                    [
                        InlineKeyboardButton(text="✅ Kabul Et", callback_data=f"tch:appr_app:{app.id}"),
                        InlineKeyboardButton(text="❌ Müsait Değilim", callback_data=f"tch:rej_app:{app.id}")
                    ]
                ])
                await message.bot.send_message(chat_id=tch.telegram_id, text=t_text, reply_markup=t_kb, parse_mode="Markdown")
            except Exception:
                pass

@router.callback_query(F.data.startswith("tch:view_app:"))
async def cb_teacher_view_appointment(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        app = await session.get(Appointment, app_id)
        if not app:
            try:
                await query.answer("Randevu bulunamadı!", show_alert=True)
            except Exception:
                pass
            return

        parent_user = await session.get(User, app.parent_telegram_id)
        p_name = parent_user.full_name if parent_user else "Veli"

        text = (
            f"🤝 *Veli Görüşme Talebi* (#{app.id})\n\n"
            f"👤 *Veli:* {p_name}\n"
            f"🕒 *Talep Edilen Zaman:* {app.preferred_time}\n"
            f"📝 *Not / Konu:* {app.note or '-'}\n"
            f"📌 *Durum:* {app.status}\n\n"
            f"Lütfen yapılacak işlemi seçiniz:"
        )
        buttons = [
            [
                InlineKeyboardButton(text="✅ Kabul Et", callback_data=f"tch:appr_app:{app.id}"),
                InlineKeyboardButton(text="❌ Müsait Değilim", callback_data=f"tch:rej_app:{app.id}")
            ],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:appointments")]
        ]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try:
        await query.answer()
    except Exception:
        pass

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
            await query.message.edit_text(get_text("no_pending_appointments", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        for a in apps:
            btn_txt = f"🤝 Randevu #{a.id} ({a.preferred_time[:20]})"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"tch:view_app:{a.id}")])
        buttons.append(get_nav_buttons(lang))

        await query.message.edit_text("🤝 *Bekleyen Veli Randevuları:*", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("tch:appr_app:"))
async def cb_teacher_approve_app(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == query.from_user.id))).scalar_one_or_none()
        app = await session.get(Appointment, app_id)
        if not tch or not app or app.teacher_id != tch.id:
            try: await query.answer("⛔ Bu randevuyu yanıtlama yetkiniz bulunmuyor!", show_alert=True)
            except Exception: pass
            return
        if app and app.status == "pending":
            app.status = "approved"
            await session.commit()
            try:
                await query.message.bot.send_message(chat_id=app.parent_telegram_id, text=get_text("appointment_approved_msg", "tr"), parse_mode="Markdown")
            except Exception:
                pass
            await query.message.edit_text("✅ Randevu onaylandı ve veliye bildirildi.")
    await query.answer()

@router.callback_query(F.data.startswith("tch:rej_app:"))
async def cb_teacher_reject_app(query: CallbackQuery):
    app_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        app = await session.get(Appointment, app_id)
        if app and app.status == "pending":
            app.status = "rejected"
            await session.commit()
            try:
                await query.message.bot.send_message(chat_id=app.parent_telegram_id, text=get_text("appointment_rejected_msg", "tr"), parse_mode="Markdown")
            except Exception:
                pass
            await query.message.edit_text("❌ Randevu talebi reddedildi.")
    await query.answer()

# --- DERS PROGRAMI VE GÜNÜN MENÜSÜ ---

@router.callback_query(F.data == "parent:toggle_briefing")
async def cb_parent_toggle_briefing(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user:
            user.evening_briefing = not user.evening_briefing
            await session.commit()
            await render_clean_dashboard(query, user)
    try: await query.answer("Ayar güncellendi.")
    except Exception: pass

@router.callback_query(F.data == "act_view_notices")
async def cb_view_notices(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        notices = (await session.execute(
            select(BroadcastNotice).order_by(desc(BroadcastNotice.created_at)).limit(5)
        )).scalars().all()

        if not notices:
            text = get_text("no_notices", lang)
        else:
            text = f"{get_text('notices_title', lang)}\n\n"
            for n in notices:
                text += f"📌 *{n.created_at.strftime('%d.%m.%Y %H:%M')}*\n{n.content}\n\n"

        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

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
        content = sched.schedule_text if sched else "• 1. Ders: 09:00 - Matematik\n• 2. Ders: 09:50 - Fizik\n• 3. Ders: 10:40 - Türkçe\n• 4. Ders: 11:30 - Tarih\n• 5. Ders: 13:00 - Biyoloji"

        text = get_text("schedule_title", lang, class_name=cls_name, content=content)
        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "act_view_cafe")
async def cb_view_cafeteria(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = datetime.utcnow().date()

        menu = (await session.execute(select(CafeteriaMenu).where(CafeteriaMenu.date == today))).scalar_one_or_none()
        content = menu.menu_text if menu else "🍲 Mercimek Çorbası\n🍗 Fırında Tavuk & Pilav\n🥗 Mevsim Salatası\n🍎 Meyve / Ayran"

        text = get_text("cafeteria_today_title", lang, date=today.strftime("%d.%m.%Y"), content=content)
        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:blacklist")
async def cb_admin_blacklist(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        now = datetime.utcnow()
        blocked_users = (await session.execute(
            select(User).where((User.is_blacklisted == True) | (User.locked_until > now))
        )).scalars().all()

        if not blocked_users:
            buttons = [get_nav_buttons(lang)]
            await query.message.edit_text(get_text("no_blacklisted", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            try: await query.answer()
            except Exception: pass
            return

        buttons = []
        for u in blocked_users:
            b_type = "🚫 Kara Liste" if u.is_blacklisted else "⏱️ Kilitli"
            name_str = u.full_name or f"ID: {u.telegram_id}"
            btn_txt = f"{b_type}: {name_str}"
            buttons.append([InlineKeyboardButton(text=f"🟢 Engeli Kaldır: {btn_txt}", callback_data=f"adm:unban:{u.telegram_id}")])

        buttons.append(get_nav_buttons(lang))
        await query.message.edit_text(get_text("blacklisted_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

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
            try:
                await query.message.bot.send_message(chat_id=t_id, text="🟢 *Hesabınızın engeli okul yönetimi tarafından kaldırılmıştır. Tekrar giriş yapabilirsiniz.*", parse_mode="Markdown")
            except Exception:
                pass
            try: await query.answer(get_text("unban_success", lang), show_alert=True)
            except Exception: pass
            await cb_admin_blacklist(query)
            return
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data == "adm:menu_edit")
async def cb_admin_menu_edit(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await query.message.edit_text(get_text("prompt_menu_update", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.menu_update_text)
    await query.answer()

@router.message(Form.menu_update_text)
async def process_menu_update_text(message: Message, state: FSMContext):
    m_text = message.text.strip()
    await state.clear()
    today = datetime.utcnow().date()

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
# 11. YÖNETİCİ: ÖĞRENCİ, ÖĞRETMEN, SINIFLAR, KOKPİT VE PDF İŞLEMLERİ
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
    name_val = message.text.strip()
    if len(name_val) < 3:
        await message.answer("⚠️ *Hatalı Giriş:* Lütfen öğrencinin adını ve soyadını eksiksiz yazınız (en az 3 karakter):", parse_mode="Markdown")
        return
    await state.update_data(name=name_val)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    await message.answer(
        get_text("prompt_student_class", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.add_student_class)

@router.message(Form.add_student_class)
async def process_student_class(message: Message, state: FSMContext):
    await state.update_data(class_name=message.text.strip().upper())
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    await message.answer(
        get_text("prompt_student_no", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.add_student_no)

@router.message(Form.add_student_no)
async def process_student_no(message: Message, state: FSMContext):
    data = await state.get_data()
    await state.clear()
    full_name = data.get("name")
    class_name = data.get("class_name")
    student_no = message.text.strip()

    st_code = generate_secure_code("OGR")
    pr_code = generate_secure_code("VELI")

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        student = Student(
            full_name=full_name,
            class_name=class_name,
            student_number=student_no,
            student_code=st_code,
            parent_code=pr_code
        )
        session.add(student)
        await session.commit()

        text = get_text("student_added_card", lang, name=full_name, class_name=class_name, no=student_no, st_code=st_code, pr_code=pr_code)
        buttons = [
            [InlineKeyboardButton(text="➕ Başka Ekle", callback_data="adm:add_student")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

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
    name_val = message.text.strip()
    if len(name_val) < 3:
        await message.answer("⚠️ *Hatalı Giriş:* Lütfen öğretmenin adını ve soyadını eksiksiz yazınız (en az 3 karakter):", parse_mode="Markdown")
        return
    await state.update_data(name=name_val)
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    await message.answer(
        get_text("prompt_teacher_subject", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.add_teacher_subject)

@router.message(Form.add_teacher_subject)
async def process_teacher_subject(message: Message, state: FSMContext):
    data = await state.get_data()
    await state.clear()
    full_name = data.get("name")
    subject = message.text.strip()
    auth_code = generate_secure_code("HCA")

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        teacher = Teacher(full_name=full_name, subject=subject, auth_code=auth_code)
        session.add(teacher)
        await session.commit()

        text = get_text("teacher_added_card", lang, name=full_name, subject=subject, code=auth_code)
        buttons = [
            [InlineKeyboardButton(text="➕ Başka Ekle", callback_data="adm:add_teacher")],
            [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
        ]
        await message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "adm:search_student")
async def cb_search_student_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await query.message.edit_text(get_text("prompt_search_student", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            (Student.full_name.ilike(f"%{query_text}%")) |
            (Student.student_number == query_text)
        ).limit(10)
        results = (await session.execute(stmt)).scalars().all()

        if not results:
            buttons = [
                [InlineKeyboardButton(text="🔍 Tekrar Ara", callback_data="adm:search_student")],
                [InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")]
            ]
            await message.answer(get_text("search_no_results", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            return

        buttons = []
        for s in results:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.class_name} - №{s.student_number})", callback_data=f"adm:st_card:{s.id}")])

        buttons.append(get_nav_buttons(lang))
        await message.answer(get_text("search_results_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

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

        buttons.append([InlineKeyboardButton(text="➕ Yeni Öğretmen Ekle", callback_data="adm:add_teacher")])
        buttons.append(get_nav_buttons(lang))

        text = f"👨‍🏫 *Öğretmenler Listesi* (Toplam {len(teachers)} Öğretmen):\nDetay veya şifre işlemleri için tıklayınız:"
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try:
        await query.answer()
    except Exception:
        pass

@router.callback_query(F.data.startswith("adm:tch_card:"))
async def cb_admin_teacher_card(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if not tch:
            try: await query.answer("Öğretmen bulunamadı!", show_alert=True)
            except Exception: pass
            return

        status_str = f"Bağlı (`{tch.telegram_id}`)" if tch.telegram_id else "Henüz Giriş Yapmadı"
        text = get_text("teacher_card", lang, name=tch.full_name, subject=tch.subject, code=tch.auth_code, status=status_str)

        buttons = [
            [InlineKeyboardButton(text="🔄 Kodu Sıfırla", callback_data=f"adm:reset_tch:{tch.id}")],
            [InlineKeyboardButton(text="❌ Öğretmeni Sil", callback_data=f"adm:del_tch:{tch.id}")],
            get_nav_buttons(lang, back_callback="adm:teachers")
        ]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try:
        await query.answer()
    except Exception:
        pass

@router.callback_query(F.data.startswith("adm:reset_tch:"))
async def cb_admin_reset_tch_code(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if tch:
            tch.auth_code = generate_secure_code("HCA")
            tch.is_code_burned = False
            tch.telegram_id = None
            await session.commit()
            text = get_text("tch_code_reset", lang, code=tch.auth_code)
            buttons = [get_nav_buttons(lang, back_callback=f"adm:tch_card:{tch.id}")]
            await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try:
        await query.answer()
    except Exception:
        pass

@router.callback_query(F.data.startswith("adm:del_tch:"))
async def cb_admin_delete_teacher(query: CallbackQuery):
    tch_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = await session.get(Teacher, tch_id)
        if tch:
            await session.execute(delete(Appointment).where(Appointment.teacher_id == tch.id))
            await session.delete(tch)
            await session.commit()
            try: await query.answer(get_text("teacher_deleted", lang), show_alert=True)
            except Exception: pass
            await cb_admin_teachers_list(query)
            return
    try:
        await query.answer()
    except Exception:
        pass

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
            await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:show_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row:
            buttons.append(row)

        buttons.append(get_nav_buttons(lang))
        text = f"🏫 *{get_text('btn_classes', lang)}*\n\nİncelemek istediğiniz sınıfı seçiniz:"
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "noop")
async def cb_noop(query: CallbackQuery):
    try:
        await query.answer()
    except Exception:
        pass

@router.message(any_state, F.text.in_([
    "👥 Kadro & Öğrenci", "👥 Ученики и учителя", "👥 Kadro va o'quvchilar", "👥 Staff & Students"
]))
async def cb_cat_staff(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"), InlineKeyboardButton(text=get_text("btn_teachers", lang), callback_data="adm:teachers")],
            [InlineKeyboardButton(text=get_text("btn_search_student", lang), callback_data="adm:search_student")]
        ]
        await message.answer(get_text("cat_staff_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_([
    "📊 Raporlar & Denetim", "📊 Отчеты и контроль", "📊 Hisobotlar va nazorat", "📊 Reports & Audits"
]))
async def cb_cat_reports(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_cockpit_unified", lang), callback_data="adm:cockpit")],
            [InlineKeyboardButton(text=get_text("btn_risk_radar", lang), callback_data="adm:risk_radar"), InlineKeyboardButton(text=get_text("btn_academic_report", lang), callback_data="adm:academic_report")]
        ]
        await message.answer(get_text("cat_reports_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_([
    "🛎️ Onay Masası", "🛎️ Центр одобрений", "🛎️ Tasdiqlash markazi", "🛎️ Approval Center"
]))
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
        await message.answer(get_text("cat_requests_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_([
    "🛠️ İdari Araçlar", "🛠️ Инструменты", "🛠️ Boshqaruv vositalari", "🛠️ Admin Tools"
]))
async def cb_cat_tools(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not is_admin_user(user, message.from_user.id): return
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_init"), InlineKeyboardButton(text=get_text("btn_excel_hub", lang), callback_data="adm:excel_hub")],
            [InlineKeyboardButton(text=get_text("btn_pdf", lang), callback_data="adm:pdf_menu"), InlineKeyboardButton(text=get_text("btn_cafeteria_edit", lang), callback_data="adm:menu_edit")]
        ]
        await message.answer(get_text("cat_tools_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_([
    "⚙️ Sistem & Ayarlar", "⚙️ Настройки системы", "⚙️ Tizim va sozlamalar", "⚙️ System & Settings"
]))
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
            [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang"), InlineKeyboardButton(text=get_text("rk_logout", lang), callback_data="act_logout")]
        ]
        await message.answer(get_text("cat_settings_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data.startswith("adm:show_class:"))
async def cb_show_class_students(query: CallbackQuery):
    parts = query.data.split(":")
    class_name = parts[2]
    page = int(parts[3]) if len(parts) > 3 and parts[3].isdigit() else 0
    PAGE_SIZE = 10

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        all_students = (await session.execute(
            select(Student).where(Student.class_name == class_name).order_by(Student.student_number)
        )).scalars().all()

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
                nav_row.append(InlineKeyboardButton(text="⬅️ Önceki", callback_data=f"adm:show_class:{class_name}:{page - 1}"))
            nav_row.append(InlineKeyboardButton(text=f"📄 {page + 1}/{total_pages}", callback_data="noop"))
            if page < total_pages - 1:
                nav_row.append(InlineKeyboardButton(text="Sonraki ➡️", callback_data=f"adm:show_class:{class_name}:{page + 1}"))
            buttons.append(nav_row)

        buttons.append([InlineKeyboardButton(text=f"📄 {class_name} Şifre Kartları (PDF)", callback_data=f"adm:gen_pdf:{class_name}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:classes"))

        text = f"🏫 *{class_name} Sınıfı Listesi* (Toplam {total_students} Öğrenci):\nDetay veya şifre işlemleri için öğrenciye tıklayınız:"
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try:
        await query.answer()
    except Exception:
        pass

@router.callback_query(F.data.startswith("adm:edit_st:"))
async def cb_admin_edit_student_menu(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st: return

        buttons = [
            [InlineKeyboardButton(text="👤 İsim Değiştir", callback_data=f"adm:edf:{st.id}:name")],
            [InlineKeyboardButton(text="🏫 Sınıf / Nakil", callback_data=f"adm:edf:{st.id}:class")],
            [InlineKeyboardButton(text="🔢 Numara Değiştir", callback_data=f"adm:edf:{st.id}:no")],
            get_nav_buttons(lang, back_callback=f"adm:st_card:{st.id}")
        ]
        text = f"✏️ *{st.full_name}* ({st.class_name} - No: {st.student_number})\nDüzenlemek istediğiniz bilgiyi seçiniz:"
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data.startswith("adm:edf:"))
async def cb_admin_edit_field_init(query: CallbackQuery, state: FSMContext):
    parts = query.data.split(":")
    st_id = int(parts[2])
    field = parts[3]
    await state.update_data(edit_st_id=st_id, edit_field=field)

    prompts = {
        "name": "👤 Öğrencinin yeni Adını ve Soyadını yazınız:",
        "class": "🏫 Öğrencinin yeni Sınıfını yazınız (Örn: `10-B`):",
        "no": "🔢 Öğrencinin yeni Okul Numarasını yazınız (Örn: `205`):"
    }
    await query.message.edit_text(prompts.get(field, "Lütfen yeni değeri yazınız:"), reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text="⬅️ İptal", callback_data=f"adm:st_card:{st_id}")]]))
    await state.set_state(Form.edit_student_val)
    try: await query.answer()
    except Exception: pass

@router.message(Form.edit_student_val)
async def process_student_edit_val(message: Message, state: FSMContext):
    data = await state.get_data()
    st_id = data.get("edit_st_id")
    field = data.get("edit_field")
    new_val = message.text.strip()
    await state.clear()

    async with AsyncSessionLocal() as session:
        st = await session.get(Student, st_id)
        if st:
            if field == "name":
                st.full_name = new_val
            elif field == "class":
                st.class_name = new_val.upper()
            elif field == "no":
                st.student_number = new_val
            await session.commit()
            await message.answer(f"✅ *Öğrenci Bilgisi Güncellendi:* {st.full_name} ({st.class_name} - No: {st.student_number})", parse_mode="Markdown")
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data=f"adm:st_card:{st.id}")
            await cb_student_card(dummy_q)

@router.callback_query(F.data.startswith("adm:st_card:"))
async def cb_student_card(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if not st:
            await query.answer("Öğrenci bulunamadı!", show_alert=True)
            return

        st_status = "Kullanıldı" if st.is_student_code_burned else "Aktif / Boşta"
        pr_status = "Kullanıldı" if st.is_parent_code_burned else "Aktif / Boşta"

        text = get_text("student_card", lang, name=st.full_name, class_name=st.class_name, no=st.student_number, st_code=st.student_code, st_status=st_status, pr_code=st.parent_code, pr_status=pr_status)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_edit_student", lang), callback_data=f"adm:edit_st:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_reset_codes", lang), callback_data=f"adm:reset_codes:{st.id}")],
            [InlineKeyboardButton(text=get_text("btn_del_student", lang), callback_data=f"adm:del_student:{st.id}")],
            get_nav_buttons(lang, back_callback=f"adm:show_class:{st.class_name}")
        ]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:reset_codes:"))
async def cb_reset_student_codes(query: CallbackQuery):
    st_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)
        if st:
            st.student_code = generate_secure_code("OGR")
            st.parent_code = generate_secure_code("VELI")
            st.is_student_code_burned = False
            st.is_parent_code_burned = False
            await session.commit()
            text = get_text("codes_reset_done", lang, st_code=st.student_code, pr_code=st.parent_code)
            buttons = [get_nav_buttons(lang, back_callback=f"adm:st_card:{st.id}")]
            await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await session.execute(delete(ParentStudent).where(ParentStudent.student_id == st.id))
            await session.execute(delete(Grade).where(Grade.student_id == st.id))
            await session.execute(delete(Attendance).where(Attendance.student_id == st.id))
            await session.execute(delete(MedicalReport).where(MedicalReport.student_id == st.id))
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
    if state:
        await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user:
            await render_clean_dashboard(query, user)
    await query.answer()

@router.callback_query(F.data == "adm:cockpit")
async def cb_cockpit(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        today = datetime.utcnow().date()
        total_students = (await session.execute(select(func.count(Student.id)))).scalar() or 0
        present_count = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "present"))).scalar() or 0
        absent_count = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "absent"))).scalar() or 0

        all_classes_stmt = select(Student.class_name).distinct()
        all_classes = (await session.execute(all_classes_stmt)).scalars().all()
        taken_classes_stmt = select(Attendance.class_name).where(Attendance.date == today).distinct()
        taken_classes = (await session.execute(taken_classes_stmt)).scalars().all()

        missing = [c for c in all_classes if c not in taken_classes]
        missing_str = "\n".join([f"• ❌ *{c}*" for c in missing]) if missing else "✅ Tümü Alındı"

        text = get_text("cockpit_report", lang, date=today.strftime("%d.%m.%Y"), total=total_students, present=present_count, absent=absent_count, missing_cnt=len(missing), missing=missing_str)
        buttons = []
        if missing:
            buttons.append([InlineKeyboardButton(text=get_text("btn_remind_att", lang), callback_data="adm:remind_all_att")])
        buttons.append(get_nav_buttons(lang))
        await safe_edit_or_answer(query, text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

@router.callback_query(F.data == "adm:risk_radar")
async def cb_admin_risk_radar(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        students = (await session.execute(select(Student).order_by(Student.class_name, Student.student_number))).scalars().all()

        risky_items = []
        for s in students:
            # Devamsızlık hesapla
            abs_cnt = (await session.execute(
                select(func.count(Attendance.id)).where(Attendance.student_id == s.id, Attendance.status == "absent")
            )).scalar() or 0

            # Ortalama ve uyarı rozeti hesapla
            grades = (await session.execute(select(Grade).where(Grade.student_id == s.id))).scalars().all()
            avg_score = (sum(g.score for g in grades) / len(grades)) if grades else 100.0
            has_warning = any(g.badge == "🔴" for g in grades)

            reasons = []
            if abs_cnt >= 7:
                reasons.append(f"Devamsızlık: {abs_cnt} gün")
            if avg_score < 50.0:
                reasons.append(f"Düşük Not: {round(avg_score, 1)}")
            if has_warning:
                reasons.append("🔴 Uyarı Rozeti")

            if reasons:
                risky_items.append(f"• *{s.full_name}* ({s.class_name} - No: {s.student_number}):\n  ↳ _{', '.join(reasons)}_")

        if not risky_items:
            content = get_text("no_risky_students", lang)
        else:
            content = "\n\n".join(risky_items[:20])

        text = f"{get_text('risk_radar_title', lang)}\n\n{content}"
        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data == "adm:audit_logs")
async def cb_admin_audit_logs(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        logs = (await session.execute(select(AuditLog).order_by(desc(AuditLog.created_at)).limit(12))).scalars().all()
        if not logs:
            content = get_text("no_audit_logs", lang)
        else:
            lines = []
            for l in logs:
                lines.append(f"📌 *{l.created_at.strftime('%d.%m %H:%M')}* - *{l.action}*\n↳ {l.user_name}: _{l.details}_")
            content = "\n\n".join(lines)

        text = f"{get_text('audit_logs_title', lang)}\n\n{content}"
        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data == "adm:academic_report")
async def cb_admin_academic_report(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        # Sınıf bazlı başarı ortalaması
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        rankings = []
        for c in classes:
            st_ids = (await session.execute(select(Student.id).where(Student.class_name == c))).scalars().all()
            grades = (await session.execute(select(Grade.score).where(Grade.student_id.in_(st_ids)))).scalars().all()
            if grades:
                avg = sum(grades) / len(grades)
                rankings.append((c, round(avg, 1), len(grades)))

        rankings.sort(key=lambda x: x[1], reverse=True)
        if not rankings:
            content = "Henüz girilmiş ders notu bulunmamaktadır."
        else:
            lines = []
            for idx, (c, avg, cnt) in enumerate(rankings, 1):
                medal = "🥇" if idx == 1 else ("🥈" if idx == 2 else ("🥉" if idx == 3 else f"{idx}."))
                lines.append(f"{medal} *{c} Sınıfı:* Ortalama *{avg}* ({cnt} not girişi)")
            content = "\n".join(lines)

        text = f"{get_text('academic_report_title', lang)}\n\n{content}"
        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data == "adm:teacher_attendance_check")
async def cb_admin_att_check(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        if not is_admin_user(user, query.from_user.id): return

        today = datetime.utcnow().date()
        all_classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()

        missing = [c for c in all_classes if c not in taken_classes]
        buttons = []
        if not missing:
            text = f"{get_text('att_check_title', lang, date=today.strftime('%d.%m.%Y'))}\n\n{get_text('att_check_all_done', lang)}"
        else:
            missing_lines = [f"• ❌ *{c} Sınıfı* (Yoklama girilmedi)" for c in missing]
            text = (
                f"{get_text('att_check_title', lang, date=today.strftime('%d.%m.%Y'))}\n\n"
                f"⚠️ *Yoklama Almayan Sınıflar ({len(missing)}):*\n"
                + "\n".join(missing_lines)
            )
            buttons.append([InlineKeyboardButton(text=get_text("btn_remind_att", lang), callback_data="adm:remind_all_att")])

        buttons.append(get_nav_buttons(lang))
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data == "adm:remind_all_att")
async def cb_admin_remind_all_att(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        today = datetime.utcnow().date()
        all_classes = (await session.execute(select(Student.class_name).distinct())).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()
        missing = [c for c in all_classes if c not in taken_classes]

        # Öğretmenlere hatırlatma mesajı at
        teachers = (await session.execute(select(Teacher).where(Teacher.telegram_id.isnot(None)))).scalars().all()
        for t in teachers:
            try:
                msg = get_text("teacher_att_reminder_msg", "tr", class_name=", ".join(missing))
                await query.message.bot.send_message(chat_id=t.telegram_id, text=msg, parse_mode="Markdown")
                await asyncio.sleep(0.05)
            except Exception:
                pass

        try: await query.answer(get_text("remind_att_sent", "tr"), show_alert=True)
        except Exception: pass
        await cb_admin_att_check(query)

@router.callback_query(F.data == "adm:pdf_menu")
async def cb_pdf_menu(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            text = get_text("no_classes_found", lang)
            buttons = [get_nav_buttons(lang)]
            await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"📄 {c}", callback_data=f"adm:gen_pdf:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row:
            buttons.append(row)

        buttons.append(get_nav_buttons(lang))
        await query.message.edit_text(get_text("select_pdf_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("adm:gen_pdf:"))
async def cb_generate_pdf(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        pdf_buffer = await generate_classroom_pdf_cards(class_name, lang=lang)
        file = BufferedInputFile(pdf_buffer.read(), filename=f"{class_name}_Sifre_Kartlari.pdf")
        caption = get_text("pdf_ready", lang, class_name=class_name)
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
    caption = get_text("export_ready", lang, date=today_str)
    await query.message.answer_document(file, caption=caption, parse_mode="Markdown")
    buf.close()
    try:
        await query.answer()
    except Exception:
        pass

@router.callback_query(F.data == "adm:excel_info")
async def cb_excel_info(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await query.message.edit_text(get_text("excel_info", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.message(F.document, F.document.file_name.endswith(".xlsx"))
async def admin_excel_upload(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    try:
        bot: Bot = message.bot
        file_info = await bot.get_file(message.document.file_id)
        file_bytes = await bot.download_file(file_info.file_path)

        count, out_excel = await process_student_excel(file_bytes.read())
        file = BufferedInputFile(out_excel.read(), filename="Giris_Kodlari_Uretildi.xlsx")
        await message.answer_document(file, caption=get_text("excel_done", lang, count=count), parse_mode="Markdown")
        out_excel.close()
    except Exception:
        await message.answer("⚠️ *Dosya İşleme Hatası:* Excel dosyası okunamadı veya formatı geçersiz.\n\nLütfen ilk satır başlıklarının `Ad Soyad | Sinif | Numara` olduğundan ve dosyanın bozuk olmadığından emin olunuz.", parse_mode="Markdown")

@router.callback_query(F.data == "adm:toggle_maint")
async def cb_toggle_maintenance(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        maint = await session.get(SystemSetting, "maintenance_mode")
        if not maint:
            maint = SystemSetting(key="maintenance_mode", value="true")
            session.add(maint)
        else:
            maint.value = "false" if maint.value == "true" else "true"
        await session.commit()
    await render_clean_dashboard(query, user)
    await query.answer("Bakım modu güncellendi.")

# --- MAZERET RAPORLARI ---

@router.callback_query(F.data == "adm:medical_list")
async def cb_medical_list(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        reports = (await session.execute(
            select(MedicalReport, Student)
            .join(Student, MedicalReport.student_id == Student.id)
            .where(MedicalReport.status == "pending")
            .order_by(MedicalReport.created_at.desc())
        )).all()

        if not reports:
            buttons = [get_nav_buttons(lang)]
            await query.message.edit_text(get_text("no_pending_medical", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
            await query.answer()
            return

        buttons = []
        for rep, st in reports:
            btn_txt = f"🏥 {st.full_name} ({st.class_name}) - {rep.created_at.strftime('%H:%M')}"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"adm:view_med:{rep.id}")])

        buttons.append(get_nav_buttons(lang))
        await query.message.edit_text(get_text("pending_medical_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
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
            await query.answer("Rapor bulunamadı!", show_alert=True)
            return

        adm_kb = InlineKeyboardMarkup(inline_keyboard=[
            [
                InlineKeyboardButton(text="✅ Onayla (İzinli Say)", callback_data=f"adm:appr_med:{rep.id}"),
                InlineKeyboardButton(text="❌ Reddet", callback_data=f"adm:rej_med:{rep.id}")
            ],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:medical_list")]
        ])

        caption = f"🏥 *Mazeret Raporu*\nÖğrenci: *{st.full_name}* ({st.class_name})\nNot: {rep.caption}"
        try:
            await query.message.delete()
        except Exception:
            pass
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
            today = datetime.utcnow().date()
            att_stmt = select(Attendance).where(Attendance.student_id == rep.student_id, Attendance.date == today)
            att = (await session.execute(att_stmt)).scalar_one_or_none()
            if att:
                att.status = "excused"
            else:
                st = await session.get(Student, rep.student_id)
                new_att = Attendance(
                    student_id=rep.student_id,
                    class_name=st.class_name if st else "Genel",
                    date=today,
                    status="excused",
                    teacher_id=0,
                    notify_at=datetime.utcnow(),
                    is_notified=True
                )
                session.add(new_att)
            await session.commit()

            try:
                await query.message.bot.send_message(chat_id=rep.parent_telegram_id, text="✅ *Sağlık / Mazeret raporunuz onaylandı. Öğrenci izinli sayıldı.*", parse_mode="Markdown")
            except Exception:
                pass

            try:
                await query.message.delete()
            except Exception:
                pass
            await query.message.answer(get_text("medical_approved", lang))
            if user:
                await render_clean_dashboard(query.message, user)
    try:
        await query.answer()
    except Exception:
        pass

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
            try:
                await query.message.bot.send_message(chat_id=rep.parent_telegram_id, text="❌ *Sağlık / Mazeret raporunuz reddedildi.*", parse_mode="Markdown")
            except Exception:
                pass
            try:
                await query.message.delete()
            except Exception:
                pass
            await query.message.answer(get_text("medical_rejected", lang))
            if user:
                await render_clean_dashboard(query.message, user)
    try:
        await query.answer()
    except Exception:
        pass

# --- TOPLU DUYURU ---

@router.callback_query(F.data == "adm:broadcast_init")
async def cb_broadcast_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await query.message.edit_text(get_text("prompt_broadcast", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_broadcast_text)
    await query.answer()

@router.message(Form.waiting_broadcast_text)
async def process_broadcast_text(message: Message, state: FSMContext):
    b_text = message.text.strip()
    if len(b_text) > 4000:
        b_text = b_text[:4000] + "\n...(Metin kısaltıldı)"
    await state.clear()
    sent_cnt = 0

    import html
    safe_text = html.escape(b_text)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        # Veritabanına kalıcı duyuru kaydı ekle
        notice = BroadcastNotice(content=b_text)
        session.add(notice)
        await session.commit()

        users = (await session.execute(select(User.telegram_id))).scalars().all()
        for u_id in users:
            try:
                await message.bot.send_message(chat_id=u_id, text=f"📢 <b>OKUL DUYURUSU</b>\n\n{safe_text}", parse_mode="HTML")
                sent_cnt += 1
                await asyncio.sleep(0.05)
            except Exception:
                pass

        await message.answer(get_text("broadcast_success", lang, count=sent_cnt), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

# ======================================================================
# 12. ÖĞRETMEN MODÜLÜ: HIZLI YOKLAMA, NOT VE ÖDEV PANOSU
# ======================================================================

@router.callback_query(F.data == "tch:classes")
async def cb_teacher_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            classes = ["9-A"]

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"att_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row:
            buttons.append(row)

        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await query.message.edit_text("📋 Yoklama alacağınız sınıfı seçiniz:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("att_class:"))
async def cb_attendance_class(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        stmt = select(Student).where(Student.class_name == class_name).order_by(Student.full_name)
        students = (await session.execute(stmt)).scalars().all()

        # Veritabanında bugüne ait önceden girilmiş yoklama durumunu çekerek hatırla
        today = datetime.utcnow().date()
        today_att = (await session.execute(
            select(Attendance).where(Attendance.class_name == class_name, Attendance.date == today)
        )).scalars().all()
        today_map = {a.student_id: (a.status == "absent") for a in today_att}

        ATTENDANCE_CACHE[query.from_user.id] = {s.id: today_map.get(s.id, False) for s in students}
        kb = get_attendance_grid_kb(students, ATTENDANCE_CACHE[query.from_user.id], class_name, lang=lang)
        text = get_text("attendance_intro", lang, class_name=class_name)
        await query.message.edit_text(text, reply_markup=kb, parse_mode="Markdown")
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
        stmt = select(Student).where(Student.class_name == st.class_name).order_by(Student.full_name)
        students = (await session.execute(stmt)).scalars().all()

        kb = get_attendance_grid_kb(students, cache, st.class_name, lang=lang)
        await query.message.edit_reply_markup(reply_markup=kb)
    await query.answer()

@router.callback_query(F.data.startswith("att_save:"))
async def cb_attendance_save(query: CallbackQuery):
    class_name = query.data.split(":")[1]
    cache = ATTENDANCE_CACHE.pop(query.from_user.id, {})
    now = datetime.utcnow()
    notify_time = now + timedelta(minutes=15)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        for st_id, is_absent in cache.items():
            new_status = "absent" if is_absent else "present"
            att_stmt = select(Attendance).where(Attendance.student_id == st_id, Attendance.date == now.date())
            existing = (await session.execute(att_stmt)).scalar_one_or_none()
            if existing:
                if existing.status != "excused":
                    existing.status = new_status
                    existing.notify_at = notify_time
                    existing.is_notified = False
            else:
                att = Attendance(
                    student_id=st_id,
                    class_name=class_name,
                    date=now.date(),
                    status=new_status,
                    teacher_id=query.from_user.id,
                    notify_at=notify_time,
                    is_notified=False
                )
                session.add(att)

            # Kritik Devamsızlık Eşiği Uyarısı (5, 8 veya 10 gün)
            if new_status == "absent":
                tot_abs = (await session.execute(
                    select(func.count(Attendance.id)).where(Attendance.student_id == st_id, Attendance.status == "absent")
                )).scalar() or 0
                if tot_abs in [5, 8, 10]:
                    parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
                    st_obj = await session.get(Student, st_id)
                    st_n = st_obj.full_name if st_obj else "Öğrenci"
                    for p_id in parents:
                        try:
                            warn_msg = get_text("absence_critical_warning", "tr", name=st_n, count=tot_abs)
                            await query.message.bot.send_message(chat_id=p_id, text=warn_msg, parse_mode="Markdown")
                        except Exception:
                            pass

        await session.commit()
        await log_audit(session, query.from_user.id, user.full_name or "Öğretmen", "Yoklama Kaydı", f"{class_name} sınıfı yoklaması kaydedildi.")
        await session.commit()

        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(get_text("att_saved", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

# --- NOT GİRİŞİ ---

@router.callback_query(F.data == "tch:grade_classes")
async def cb_grade_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            classes = ["9-A"]

        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"gr_cls:{c}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await query.message.edit_text("📝 Not gireceğiniz sınıfı seçiniz:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
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
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:grade_classes")])
        await query.message.edit_text(f"📝 *{class_name} Sınıfı* - Öğrenciyi seçiniz:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data.startswith("gr_st:"))
async def cb_grade_enter_score(query: CallbackQuery, state: FSMContext):
    st_id = int(query.data.split(":")[1])
    GRADE_CACHE[query.from_user.id] = {"student_id": st_id}

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, st_id)

    text = get_text("prompt_grade_score", lang, name=st.full_name, class_name=st.class_name)
    await query.message.edit_text(text, parse_mode="Markdown")
    await state.set_state(Form.grade_score)
    await query.answer()

@router.message(Form.grade_score)
async def process_grade_score(message: Message, state: FSMContext):
    score_str = message.text.strip().replace(",", ".")
    try:
        score_val = float(score_str)
    except ValueError:
        await message.answer("Lütfen geçerli bir sayı giriniz (0-100):")
        return

    if not (0.0 <= score_val <= 100.0):
        await message.answer("⚠️ Not 0 ile 100 arasında olmalıdır. Lütfen tekrar giriniz (0-100):")
        return

    GRADE_CACHE[message.from_user.id]["score"] = score_val

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

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
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        tch = (await session.execute(select(Teacher).where(Teacher.telegram_id == user.telegram_id))).scalar_one_or_none()
        subject_name = tch.subject if tch else "Ders"

        st = await session.get(Student, st_id)
        grade = Grade(
            student_id=st_id,
            subject=subject_name,
            score=score_val,
            badge=badge_val,
            note="Öğretmen Değerlendirmesi",
            teacher_id=query.from_user.id
        )
        session.add(grade)
        await session.commit()

        parents = (await session.execute(select(ParentStudent.parent_telegram_id).where(ParentStudent.student_id == st_id))).scalars().all()
        for p_id in parents:
            try:
                alert_text = f"📝 *YENİ DERS NOTU*\n\nÖğrenciniz *{st.full_name}*, {subject_name} dersinden *{score_val}* aldı. ({badge_val})"
                await query.message.bot.send_message(chat_id=p_id, text=alert_text, parse_mode="Markdown")
            except Exception:
                pass

        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(get_text("grade_saved_success", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

# --- ÖDEV PANOSU ---

@router.callback_query(F.data == "tch:recent_grades")
async def cb_teacher_recent_grades(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        grades = (await session.execute(
            select(Grade, Student)
            .join(Student, Grade.student_id == Student.id)
            .where(Grade.teacher_id == query.from_user.id)
            .order_by(desc(Grade.created_at))
            .limit(15)
        )).all()

        if not grades:
            buttons = [get_nav_buttons(lang)]
            await query.message.edit_text(get_text("no_recent_grades", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            try: await query.answer()
            except Exception: pass
            return

        buttons = []
        for g, s in grades:
            btn_txt = f"{g.badge} {s.full_name} ({s.class_name}): {g.score} - {g.subject}"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"tch:view_gr:{g.id}")])

        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await query.message.edit_text(get_text("recent_grades_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data.startswith("tch:view_gr:"))
async def cb_teacher_view_grade(query: CallbackQuery):
    gr_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        grade = await session.get(Grade, gr_id)
        if not grade or grade.teacher_id != query.from_user.id:
            try: await query.answer("Bu nota erişim yetkiniz yok!", show_alert=True)
            except Exception: pass
            return

        st = await session.get(Student, grade.student_id)
        st_name = st.full_name if st else "Öğrenci"
        st_cls = st.class_name if st else ""
        date_str = grade.created_at.strftime("%d.%m.%Y %H:%M")

        text = get_text("grade_detail_card", lang, name=st_name, class_name=st_cls, subject=grade.subject, score=grade.score, badge=grade.badge, date=date_str)
        buttons = [
            [InlineKeyboardButton(text=get_text("btn_edit_grade", lang), callback_data=f"tch:edit_gr:{grade.id}")],
            [InlineKeyboardButton(text=get_text("btn_del_grade", lang), callback_data=f"tch:del_gr:{grade.id}")],
            [InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:recent_grades")]
        ]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    try: await query.answer()
    except Exception: pass

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
            try: await query.answer(get_text("grade_deleted", lang), show_alert=True)
            except Exception: pass
            await cb_teacher_recent_grades(query)
            return
    try: await query.answer()
    except Exception: pass

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
    try: await query.answer()
    except Exception: pass

@router.message(Form.edit_grade_val)
async def process_grade_edit_val(message: Message, state: FSMContext):
    score_str = message.text.strip().replace(",", ".")
    try:
        score_val = float(score_str)
    except ValueError:
        await message.answer("⚠️ Lütfen geçerli bir sayı giriniz (0-100):")
        return

    if not (0.0 <= score_val <= 100.0):
        await message.answer("⚠️ Not 0 ile 100 arasında olmalıdır. Tekrar giriniz:")
        return

    data = await state.get_data()
    gr_id = data.get("edit_grade_id")
    await state.clear()

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        grade = await session.get(Grade, gr_id)
        if grade and grade.teacher_id == message.from_user.id:
            grade.score = score_val
            await session.commit()
            reply_kb = get_role_reply_kb("teacher", lang)
            await message.answer(get_text("grade_updated", lang), reply_markup=reply_kb)
            await render_clean_dashboard(message, user)

@router.callback_query(F.data == "adm:excel_hub")
async def cb_admin_excel_hub(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text="📥 Toplu Öğrenci Yükle (Excel)", callback_data="adm:excel_info")],
            [InlineKeyboardButton(text="📤 Tüm Okul Verisini İndir (Yedek)", callback_data="adm:export_all_excel")],
            get_nav_buttons(lang)
        ]
        await safe_edit_or_answer(query, get_text("excel_hub_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    try: await query.answer()
    except Exception: pass

@router.message(any_state, F.text.in_([
    "ℹ️ Okul Bilgi Panosu", "ℹ️ Инфопанель школы", "ℹ️ Maktab ma'lumotlari", "ℹ️ School Info Board"
]))
async def cb_parent_info_board(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_view_schedule", lang), callback_data="act_view_sched")],
            [InlineKeyboardButton(text=get_text("btn_notices", lang), callback_data="act_view_notices")],
            [InlineKeyboardButton(text=get_text("btn_view_cafeteria", lang), callback_data="act_view_cafe")]
        ]
        await message.answer(get_text("parent_info_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.message(any_state, F.text.in_([
    "⚙️ Ayarlar & Çıkış", "⚙️ Настройки и выход", "⚙️ Sozlamalar va chiqish", "⚙️ Settings & Exit"
]))
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
        await message.answer(get_text("parent_settings_title", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

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
            buttons.append([InlineKeyboardButton(text=f"📢 {c} İçin Yeni Ödev Gönder", callback_data=f"hw_cls:{c}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_my_hws", lang), callback_data="tch:view_my_hws")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await safe_edit_or_answer(query, get_text("prompt_hw_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data == "tch:view_my_hws")
async def cb_teacher_view_my_hws(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        hws = (await session.execute(
            select(Homework).where(Homework.teacher_id == query.from_user.id).order_by(desc(Homework.created_at)).limit(10)
        )).scalars().all()

        if not hws:
            buttons = [get_nav_buttons(lang, back_callback="tch:hw_classes")]
            await safe_edit_or_answer(query, get_text("no_hws_found", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            try: await query.answer()
            except Exception: pass
            return

        buttons = []
        for h in hws:
            btn_txt = f"📚 {h.class_name} ({h.subject}): {h.content[:15]}..."
            buttons.append([
                InlineKeyboardButton(text=btn_txt, callback_data="noop"),
                InlineKeyboardButton(text="❌ Sil", callback_data=f"tch:del_hw:{h.id}")
            ])
        buttons.append([InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="tch:hw_classes")])

        await safe_edit_or_answer(query, "📚 *Yayınladığınız Son Ödevler:*", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    try: await query.answer()
    except Exception: pass

@router.callback_query(F.data.startswith("tch:del_hw:"))
async def cb_teacher_del_hw(query: CallbackQuery):
    hw_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        hw = await session.get(Homework, hw_id)
        if hw and hw.teacher_id == query.from_user.id:
            await session.delete(hw)
            await session.commit()
            try: await query.answer("Ödev silindi.", show_alert=True)
            except Exception: pass
            await cb_teacher_view_my_hws(query)
            return
    try: await query.answer()
    except Exception: pass


# --- HW CLASSES HANDLER MIGRATED ---

@router.callback_query(F.data.startswith("hw_cls:"))
async def cb_hw_enter_content(query: CallbackQuery, state: FSMContext):
    class_name = query.data.split(":")[1]
    HW_CACHE[query.from_user.id] = class_name

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    await query.message.edit_text(get_text("prompt_hw_content", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]))
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
            try:
                caption = f"📢 *{class_name} ÖDEV PANOSU ({subj})*\n\n{hw_text}"
                if photo_id:
                    await message.bot.send_photo(chat_id=t_id, photo=photo_id, caption=caption, parse_mode="Markdown")
                else:
                    await message.bot.send_message(chat_id=t_id, text=caption, parse_mode="Markdown")
            except Exception:
                pass

        await message.answer(get_text("hw_sent_success", lang, class_name=class_name), parse_mode="Markdown")
        await render_clean_dashboard(message, user)

# ======================================================================
# 13. VELİ & ÖĞRENCİ MODÜLÜ: KARNE, ÖĞRENCİ SEÇİCİ VE MAZERET BELGESİ
# ======================================================================

@router.callback_query(F.data == "parent:switch_student")
async def cb_parent_switch_student(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        stmt = select(Student).join(ParentStudent, ParentStudent.student_id == Student.id).where(ParentStudent.parent_telegram_id == user.telegram_id)
        children = (await session.execute(stmt)).scalars().all()

        if not children:
            await query.answer("Kayıtlı öğrenci bulunamadı!", show_alert=True)
            return

        buttons = []
        for c in children:
            is_active = "⭐ " if c.id == user.current_child_id else ""
            btn_txt = f"{is_active}🧑‍🎓 {c.full_name} ({c.class_name})"
            buttons.append([InlineKeyboardButton(text=btn_txt, callback_data=f"set_child:{c.id}")])

        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await query.message.edit_text("🧑‍🎓 Lütfen durumunu incelemek istediğiniz öğrenciyi seçiniz:", reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
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
    try: await query.answer()
    except Exception: pass

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
            await message.answer("❌ *Geçersiz Veli Kodu!* Lütfen kodu kontrol ediniz.", reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # Zaten bağlı mı?
        rel = (await session.execute(select(ParentStudent).where(
            ParentStudent.parent_telegram_id == message.from_user.id,
            ParentStudent.student_id == p_student.id
        ))).scalar_one_or_none()

        if not rel:
            session.add(ParentStudent(parent_telegram_id=message.from_user.id, student_id=p_student.id))
        user.current_child_id = p_student.id
        p_student.is_parent_code_burned = True
        await session.commit()

        success_text = get_text("child_added_success", lang, name=p_student.full_name, class_name=p_student.class_name)
        await message.answer(success_text, reply_markup=reply_kb, parse_mode="Markdown")
        await render_clean_dashboard(message, user)

@router.callback_query(F.data.startswith("set_child:"))
async def cb_set_child(query: CallbackQuery):
    child_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        # GÜVENLİK KONTROLÜ: Çocuğun bu veliye ait olduğu doğrulanır
        rel = (await session.execute(
            select(ParentStudent).where(
                ParentStudent.parent_telegram_id == query.from_user.id,
                ParentStudent.student_id == child_id
            )
        )).scalar_one_or_none()
        if not rel:
            try: await query.answer("⛔ Yetkisiz İşlem: Bu öğrenci kaydına erişim izniniz yok!", show_alert=True)
            except Exception: pass
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
            text = get_text("no_linked_student", lang)
            buttons = [get_nav_buttons(lang)]
            await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
            await query.answer()
            return

        att_stmt = select(Attendance).where(Attendance.student_id == st.id)
        attendances = (await session.execute(att_stmt)).scalars().all()
        absent_count = sum(1 for a in attendances if a.status == "absent")
        excused_count = sum(1 for a in attendances if a.status == "excused")

        gr_stmt = select(Grade).where(Grade.student_id == st.id).order_by(Grade.created_at.desc())
        grades = (await session.execute(gr_stmt)).scalars().all()

        grades_text = "\n".join([f"• {g.subject}: *{g.score}* ({g.badge}) - _{g.note}_" for g in grades]) if grades else get_text("no_grades", lang)

        text = (
            f"📊 *{st.full_name} ({st.class_name}) - Durum Raporu*\n\n"
            f"📌 Toplam Devamsızlık: *{absent_count} gün* (Özürlü/İzinli: {excused_count} gün)\n\n"
            f"📝 *Ders Notları:*\n{grades_text}"
        )

        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "upload_medical_init")
async def cb_upload_med_init(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

    buttons = [get_nav_buttons(lang)]
    await query.message.edit_text(get_text("upload_med_prompt", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await state.set_state(Form.waiting_medical_photo)
    await query.answer()

@router.message(Form.waiting_medical_photo, ~F.photo)
async def fallback_medical_photo_text(message: Message):
    await message.answer("⚠️ *Fotoğraf Bekleniyor:* Lütfen sağlık raporu veya mazeret belgesinin net bir fotoğrafını gönderiniz.\n\n*(İptal etmek için aşağıdaki '❌ İşlemi İptal Et' butonuna basabilirsiniz.)*", parse_mode="Markdown")

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

        report = MedicalReport(
            student_id=st.id,
            parent_telegram_id=user.telegram_id,
            file_id=photo_file_id,
            caption=message.caption or "Sağlık Raporu"
        )
        session.add(report)
        await session.commit()

        await message.answer(get_text("med_uploaded_success", lang))
        await render_clean_dashboard(message, user)
    await state.clear()

@router.callback_query(F.data.startswith("ack_notif:"))
async def cb_ack_notification(query: CallbackQuery):
    notif_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        notif = await session.get(CriticalNotification, notif_id)
        if notif:
            notif.acknowledged_at = datetime.utcnow()
            await session.commit()
    await query.answer(get_text("acknowledged_toast", "tr"), show_alert=True)
    await query.message.edit_reply_markup(reply_markup=None)

@router.callback_query(F.data == "act_logout")
async def cb_act_logout(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        u = await session.get(User, query.from_user.id)
        if u:
            u.role = "guest"
            await session.commit()
        lang = u.language if u else "tr"
    reply_kb = get_role_reply_kb("guest", lang)
    try: await query.message.delete()
    except Exception: pass
    await query.message.answer(get_text("logged_out", lang), reply_markup=reply_kb)
    await prompt_guest_screen(query, u, state)

@router.callback_query(F.data == "act_change_lang")
async def cb_change_lang_screen(query: CallbackQuery, state: FSMContext):
    await state.clear()
    await query.message.edit_text(get_text("lang_select", "tr"), reply_markup=get_language_inline_kb())
    await query.answer()

# ======================================================================
# 14. SABİT ALT MENÜ (REPLY KEYBOARD) VE GLOBAL METİN YÖNLENDİRİCİSİ
# ======================================================================

REPLY_BUTTON_ACTIONS = {
    # Ortak
    "rk_logout": "act_logout",
    "rk_main_menu": "act_main_menu",
    "rk_admin_dash": "act_main_menu",
    "rk_lang": "act_lang",
    "rk_enter_code": "act_enter_code",
    "rk_req_access": "act_req_access",
    # Yönetici Kategori Hub'ları
    "rk_cat_staff": "act_cat_staff",
    "rk_cat_reports": "act_cat_reports",
    "rk_cat_requests": "act_cat_requests",
    "rk_cat_tools": "act_cat_tools",
    "rk_cat_settings": "act_cat_settings",
    # Yönetici
    "rk_cockpit": "act_cockpit",
    "rk_classes": "act_classes",
    "rk_teachers": "act_teachers",
    "rk_requests": "act_requests",
    # Öğretmen
    "rk_attendance": "act_attendance",
    "rk_grade": "act_grade",
    "rk_homework": "act_homework",
    # Veli & Öğrenci
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
    if not text:
        return None
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
            user = User(telegram_id=user_id, language="tr", role=role)
            session.add(user)
            await session.commit()
        elif user_id in ADMIN_IDS and user.role != "admin":
            user.role = "admin"
            await session.commit()
        lang = user.language

    if action:
        await state.clear()
        if action == "act_logout":
            async with AsyncSessionLocal() as session:
                u = await session.get(User, user_id)
                if u:
                    u.role = "guest"
                    await session.commit()
            reply_kb = get_role_reply_kb("guest", lang)
            await message.answer(get_text("logged_out", lang), reply_markup=reply_kb)
            await prompt_guest_screen(message, user, state)
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

        # Yönetici Eylemleri
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

        # Öğretmen Eylemleri
        elif action == "act_attendance" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:classes")
            await cb_teacher_classes(dummy_q, state)
            return

        elif action == "act_grade" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:grade_classes")
            await cb_grade_classes(dummy_q, state)
            return

        elif action == "act_homework" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:hw_classes")
            await cb_hw_classes(dummy_q, state)
            return

        elif action == "act_appointments" and user.role == "teacher":
            dummy_q = CallbackQuery(id="0", from_user=message.from_user, chat_instance="0", message=message, data="tch:appointments")
            await cb_teacher_appointments_list(dummy_q)
            return

        # Veli & Öğrenci Eylemleri
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

@router.message(any_state, F.text.func(lambda text: normalize_code(text).startswith(("VELI-", "OGR-", "HCA-", "ADM-"))))
@router.message(F.text)
async def smart_text_auth_router(message: Message, state: FSMContext):
    """Herhangi bir aşamada kod yazıldığında onu akıllıca yakalayıp doğrudan doğrular"""
    user_id = message.from_user.id
    clean_code = normalize_code(message.text)
    is_code = clean_code.startswith(("VELI-", "OGR-", "HCA-", "ADM-")) or clean_code == ADMIN_CODE

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            role = "admin" if user_id in ADMIN_IDS else "guest"
            user = User(telegram_id=user_id, language="tr", role=role)
            session.add(user)
            await session.commit()
        if is_code or user.role == "guest":
            await process_auth_code_string(message.text, user_id, message, state)

# ======================================================================
# 15. FASTAPI VE ARKA PLAN DÖNGÜLERİ (LIFESPAN & PINGER)
# ======================================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
dp.include_router(router)

async def run_morning_briefing_worker(bot: Bot):
    """Her sabah 09:30'da okul idarecilerine anlık durum raporu iletir"""
    async with AsyncSessionLocal() as session:
        today = datetime.utcnow().date()
        total_students = (await session.execute(select(func.count(Student.id)))).scalar() or 0
        present = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "present"))).scalar() or 0
        absent = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "absent"))).scalar() or 0
        excused = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "excused"))).scalar() or 0

        all_classes = (await session.execute(select(Student.class_name).distinct())).scalars().all()
        taken_classes = (await session.execute(select(Attendance.class_name).where(Attendance.date == today).distinct())).scalars().all()
        missing = [c for c in all_classes if c not in taken_classes]
        missing_str = ", ".join(missing) if missing else "Yok (Tümü Alındı)"

        req_cnt = (await session.execute(select(func.count(AccessRequest.id)).where(AccessRequest.status == "pending"))).scalar() or 0
        med_cnt = (await session.execute(select(func.count(MedicalReport.id)).where(MedicalReport.status == "pending"))).scalar() or 0

        pct = round((present / total_students * 100), 1) if total_students > 0 else 100.0

        for adm_id in ADMIN_IDS:
            try:
                msg = get_text("morning_briefing_header", "tr", total=total_students, present=present, pct=pct, absent=absent, excused=excused, missing_cnt=len(missing), missing=missing_str, req_cnt=req_cnt, med_cnt=med_cnt)
                await bot.send_message(chat_id=adm_id, text=msg, parse_mode="Markdown")
                await asyncio.sleep(0.05)
            except Exception:
                pass

async def background_morning_briefing_loop():
    """Hafta içi her gün sabah 09:30'da idareci brifingi döngüsü"""
    while True:
        try:
            now_local = datetime.utcnow() + timedelta(hours=TIMEZONE_OFFSET)
            if now_local.weekday() < 5 and now_local.hour == 9 and now_local.minute == 30:
                await run_morning_briefing_worker(bot)
                await asyncio.sleep(70)
        except Exception:
            pass
        await asyncio.sleep(30)

async def background_attendance_loop():
    while True:
        try:
            await run_attendance_delay_worker(bot)
        except Exception:
            pass
        await asyncio.sleep(60)

async def background_keep_alive_pinger():
    """Render Free Tier uyku modunu önleyen hafif dahili pinger (10 dakikada bir)"""
    await asyncio.sleep(60)
    while True:
        target_url = None
        if WEBHOOK_URL:
            target_url = WEBHOOK_URL.replace("/webhook", "/health")
        elif os.getenv("RENDER_EXTERNAL_URL"):
            target_url = f"{os.getenv('RENDER_EXTERNAL_URL').rstrip('/')}/health"

        if target_url:
            try:
                import urllib.request
                req = urllib.request.Request(target_url, headers={"User-Agent": "OkulBot-SelfPinger/2.0"})
                with urllib.request.urlopen(req, timeout=10) as resp:
                    pass
            except Exception:
                pass
        await asyncio.sleep(600)

async def background_evening_briefing_loop():
    """Her gün 18:30'da velilere gün sonu özeti gönderen zamanlayıcı"""
    while True:
        try:
            now_local = datetime.utcnow() + timedelta(hours=TIMEZONE_OFFSET)
            # Sadece hafta içi günlerde (Pazartesi - Cuma) gönderim yap
            if now_local.weekday() < 5 and now_local.hour == 18 and now_local.minute == 30:
                await run_evening_briefing_worker(bot)
                await asyncio.sleep(70)
        except Exception:
            pass
        await asyncio.sleep(30)

@asynccontextmanager
async def lifespan(app: FastAPI):
    print("=" * 60)
    print("--> [1/5] Veritabanı başlatılıyor...")
    await init_db()
    print("--> [1/5] Veritabanı tabloları hazır.")

    bot_info = None
    try:
        bot_info = await bot.get_me()
        print(f"--> [2/5] Telegram Bot Bilgisi: @{bot_info.username} (ID: {bot_info.id})")
    except Exception as e:
        print(f"--> [HATA 2/5] BOT_TOKEN ile Telegram'a bağlanılamadı: {e}")

    if bot_info and WEBHOOK_URL:
        try:
            print(f"--> [3/5] Webhook Telegram'a kaydediliyor: {WEBHOOK_URL}")
            await bot.delete_webhook(drop_pending_updates=False)
            await bot.set_webhook(
                url=WEBHOOK_URL,
                secret_token=WEBHOOK_SECRET,
                drop_pending_updates=False,
                allowed_updates=["message", "callback_query"]
            )
            wh = await bot.get_webhook_info()
            print(f"--> [3/5] Webhook Başarıyla Kuruldu! Aktif URL: {wh.url}")
            if wh.last_error_message:
                print(f"--> [UYARI] Telegram Son Hata: {wh.last_error_message}")
        except Exception as e:
            print(f"--> [HATA 3/5] Webhook kurulum hatası: {e}")
    else:
        print(f"--> [UYARI 3/5] Webhook kurulamadı! WEBHOOK_URL='{WEBHOOK_URL}'")

    t1 = asyncio.create_task(background_attendance_loop())
    t2 = asyncio.create_task(background_keep_alive_pinger())
    t3 = asyncio.create_task(background_evening_briefing_loop())
    t4 = asyncio.create_task(background_morning_briefing_loop())
    print("--> [4/5] Gecikmeli yoklama, Self-Pinger ve 18:30 Akşam Brifingi aktif.")
    print("--> [5/5] SİSTEM CANLI VE HAZIR.")
    print("=" * 60)

    yield

    t1.cancel()
    t2.cancel()
    t3.cancel()
    t4.cancel()
    try:
        await bot.session.close()
    except Exception:
        pass

app = FastAPI(title="OkulYonetimBot", lifespan=lifespan)

@app.get("/")
async def root():
    return {
        "status": "ok",
        "service": "OkulYonetimBot",
        "webhook_url": WEBHOOK_URL,
        "uptime": True
    }

@app.get("/health", status_code=status.HTTP_200_OK)
async def health_check():
    return {"status": "ok", "service": "OkulYonetimBot", "uptime": True}

@app.post("/webhook")
async def telegram_webhook(request: Request, background_tasks: BackgroundTasks):
    secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if WEBHOOK_SECRET and secret and secret != WEBHOOK_SECRET:
        print(f"--> [GÜVENLİK ENGELİ] Geçersiz Webhook Secret: {secret}")
        return JSONResponse(status_code=status.HTTP_403_FORBIDDEN, content={"error": "Invalid secret"})

    try:
        data = await request.json()
    except Exception as e:
        print(f"--> [HATA] JSON okunamadı: {e}")
        return JSONResponse(status_code=status.HTTP_400_BAD_REQUEST, content={"error": "Bad JSON"})

    update_id = data.get("update_id", 0)
    print(f"--> [MESAJ GELDİ] Update ID: {update_id}")

    try:
        telegram_update = Update.model_validate(data, context={"bot": bot})
        await dp.feed_update(bot, telegram_update)
    except Exception as e:
        import traceback
        traceback.print_exc()
        print(f"--> [HATA] Webhook işleme hatası: {e}")

    return {"ok": True}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    if not WEBHOOK_URL:
        print("--> [BAŞLATICI] WEBHOOK_URL tanımlı değil, yerel POLLING modunda başlatılıyor...")
        asyncio.run(dp.start_polling(bot))
    else:
        uvicorn.run("app:app", host="0.0.0.0", port=port, log_level="info")

