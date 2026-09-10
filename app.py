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
from aiogram.fsm.state import State, StatesGroup

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

# Yönetici Tanımlamaları ve Güvenlik
ADMIN_CODE = os.getenv("ADMIN_CODE", "ADM-OKUL-2026").strip()
TIMEZONE_OFFSET = int(os.getenv("TIMEZONE_OFFSET", "3")) # Varsayılan UTC+3 (Türkiye)

ADMIN_IDS = []
for x in os.getenv("ADMIN_IDS", "").split(","):
    clean_x = x.strip().replace("@", "")
    if clean_x.isdigit():
        ADMIN_IDS.append(int(clean_x))

# ======================================================================
# 2. VERİTABANI MOTORU VE MODELLERİ (SQLITE & POSTGRESQL ÇİFT MOTOR)
# ======================================================================

if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://") and not DATABASE_URL.startswith("postgresql+asyncpg://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)

if "sqlite" in DATABASE_URL:
    engine = create_async_engine(DATABASE_URL, echo=False)
else:
    engine = create_async_engine(
        DATABASE_URL,
        pool_size=5,
        max_overflow=2,
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

class CafeteriaMenu(Base):
    __tablename__ = "cafeteria_menus"
    id: Mapped[int] = mapped_column(Integer, primary_key=True, autoincrement=True)
    date: Mapped[date] = mapped_column(Date, default=lambda: datetime.utcnow().date(), index=True)
    menu_text: Mapped[str] = mapped_column(Text)

async def init_db():
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)

# ======================================================================
# 3. KUSURSUZ 4 DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {
    "tr": {
        "lang_select": "🌍 Lütfen bir dil seçiniz / Tilni tanlang / Select language:",
        "lang_changed": "Dil başarıyla güncellendi: 🇹🇷 Türkçe",
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
        # Reply Keyboard Alt Menü Butonları (4 Dilli Eşleştirme İçin)
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
        # Reply Keyboard
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
        # Reply Keyboard
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
        # Reply Keyboard
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
            [KeyboardButton(text=get_text("rk_admin_dash", lang)), KeyboardButton(text=get_text("rk_cockpit", lang))],
            [KeyboardButton(text=get_text("rk_classes", lang)), KeyboardButton(text=get_text("rk_requests", lang))],
            [KeyboardButton(text=get_text("rk_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "teacher":
        keyboard = [
            [KeyboardButton(text=get_text("rk_attendance", lang)), KeyboardButton(text=get_text("rk_grade", lang))],
            [KeyboardButton(text=get_text("rk_homework", lang)), KeyboardButton(text=get_text("rk_appointments", lang))],
            [KeyboardButton(text=get_text("rk_schedule", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "parent":
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_appointments", lang))],
            [KeyboardButton(text=get_text("rk_upload_medical", lang)), KeyboardButton(text=get_text("rk_schedule", lang))],
            [KeyboardButton(text=get_text("rk_cafeteria", lang)), KeyboardButton(text=get_text("rk_switch_student", lang))],
            [KeyboardButton(text=get_text("rk_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    elif role == "student":
        keyboard = [
            [KeyboardButton(text=get_text("rk_report", lang)), KeyboardButton(text=get_text("rk_homework", lang))],
            [KeyboardButton(text=get_text("rk_schedule", lang)), KeyboardButton(text=get_text("rk_cafeteria", lang))],
            [KeyboardButton(text=get_text("rk_lang", lang)), KeyboardButton(text=get_text("rk_logout", lang))]
        ]
    else: # guest / giriş yapılmamış
        keyboard = [
            [KeyboardButton(text=get_text("rk_enter_code", lang)), KeyboardButton(text=get_text("rk_req_access", lang))],
            [KeyboardButton(text=get_text("rk_lang", lang))]
        ]

    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True, is_persistent=True)

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
            InlineKeyboardButton(text=get_text("btn_classes", lang), callback_data="adm:classes"),
            InlineKeyboardButton(text=get_text("btn_search_student", lang), callback_data="adm:search_student")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_add_student", lang), callback_data="adm:add_student"),
            InlineKeyboardButton(text=get_text("btn_add_teacher", lang), callback_data="adm:add_teacher")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_cockpit", lang), callback_data="adm:cockpit"),
            InlineKeyboardButton(text=req_text, callback_data="adm:requests_list")
        ],
        [
            InlineKeyboardButton(text=med_text, callback_data="adm:medical_list"),
            InlineKeyboardButton(text=get_text("btn_pdf", lang), callback_data="adm:pdf_menu")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_excel", lang), callback_data="adm:excel_info"),
            InlineKeyboardButton(text=get_text("btn_broadcast", lang), callback_data="adm:broadcast_init")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_cafeteria_edit", lang), callback_data="adm:menu_edit"),
            InlineKeyboardButton(text=maint_text, callback_data="adm:toggle_maint")
        ],
        [
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

def generate_secure_code(prefix: str) -> str:
    """6 haneli, yüksek entropili güvenli kod üretir"""
    num_part = str(random.randint(100000, 999999))
    return f"{prefix}-{num_part}"

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

router = Router()
ATTENDANCE_CACHE = {}
GRADE_CACHE = {}
HW_CACHE = {}
REQ_CACHE = {}
APP_CACHE = {}

# ======================================================================
# 7. ROUTER: BAŞLANGIÇ, DİL VE ROLE ÖZEL KONTROL MASALARI
# ======================================================================

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

        if is_admin_id and user.role != "admin":
            user.role = "admin"
            await session.commit()

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
    """Misafire en sade, iki temel seçenekli giriş kartı sunar"""
    lang = user.language
    text = get_text("welcome_guest", lang)
    reply_kb = get_role_reply_kb("guest", lang)

    buttons = [
        [InlineKeyboardButton(text="🔑 Şifre / Kod Gir", callback_data="act_enter_code")],
        [InlineKeyboardButton(text="📩 Şifrem Yok (Erişim İste)", callback_data="act_req_access")],
        [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")]
    ]
    inline_kb = InlineKeyboardMarkup(inline_keyboard=buttons)

    if isinstance(target, CallbackQuery):
        await target.message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
        await target.message.answer("Lütfen seçiminizi yapınız:", reply_markup=inline_kb)
    else:
        await target.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
        await target.answer("Lütfen seçiminizi yapınız:", reply_markup=inline_kb)

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

async def render_clean_dashboard(target: Message | CallbackQuery, user: User):
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
            kb = get_admin_main_inline_kb(lang, med_count=med_cnt, req_count=req_cnt, is_maintenance=is_maint)

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
                    InlineKeyboardButton(text=get_text("btn_homework_board", lang), callback_data="tch:hw_classes"),
                    InlineKeyboardButton(text=get_text("btn_teacher_appointments", lang, count=app_cnt), callback_data="tch:appointments")
                ]
            ]
            kb = InlineKeyboardMarkup(inline_keyboard=buttons)

    elif user.role == "parent":
        async with AsyncSessionLocal() as session:
            st = await session.get(Student, user.current_child_id) if user.current_child_id else None
            name = st.full_name if st else "-"
            cls_name = st.class_name if st else "-"
            text = get_text("menu_parent", lang, name=name, class_name=cls_name)
            buttons = [
                [
                    InlineKeyboardButton(text=get_text("btn_report_card", lang), callback_data="act_view_report"),
                    InlineKeyboardButton(text=get_text("btn_book_appointment", lang), callback_data="act_book_app")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_view_schedule", lang), callback_data="act_view_sched"),
                    InlineKeyboardButton(text=get_text("btn_view_cafeteria", lang), callback_data="act_view_cafe")
                ],
                [
                    InlineKeyboardButton(text=get_text("btn_upload_medical", lang), callback_data="upload_medical_init"),
                    InlineKeyboardButton(text=get_text("btn_switch_student", lang), callback_data="parent:switch_student")
                ]
            ]
            kb = InlineKeyboardMarkup(inline_keyboard=buttons)

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
            kb = InlineKeyboardMarkup(inline_keyboard=buttons)

    else: # guest
        text = get_text("welcome_guest", lang)
        buttons = [
            [InlineKeyboardButton(text=get_text("rk_enter_code", lang), callback_data="act_enter_code")],
            [InlineKeyboardButton(text=get_text("btn_req_access", lang), callback_data="act_req_access")],
            [InlineKeyboardButton(text=get_text("btn_lang", lang), callback_data="act_change_lang")]
        ]
        kb = InlineKeyboardMarkup(inline_keyboard=buttons)

    if isinstance(target, CallbackQuery):
        try:
            await target.message.edit_text(text, reply_markup=kb, parse_mode="Markdown")
        except Exception:
            await target.message.answer(text, reply_markup=kb, parse_mode="Markdown")
    else:
        await target.answer(text, reply_markup=kb, parse_mode="Markdown")

# ======================================================================
# 8. KİMLİK DOĞRULAMA (AUTHENTICATION ENGINE - TEK KAPIDAN TÜM ROLLER)
# ======================================================================

async def process_auth_code_string(code: str, user_id: int, message: Message, state: FSMContext):
    clean_code = code.strip().upper()

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
            await message.answer(success_txt, reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # 2. ÖĞRETMEN GİRİŞ KODU KONTROLÜ (HCA-XXXXXX)
        t_stmt = select(Teacher).where(Teacher.auth_code == clean_code)
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
            await message.answer(success_txt, reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # 3. ÖĞRENCİ GİRİŞ KODU KONTROLÜ (OGR-XXXXXX)
        s_stmt = select(Student).where(Student.student_code == clean_code)
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
            await message.answer(success_txt, reply_markup=reply_kb, parse_mode="Markdown")
            await render_clean_dashboard(message, user)
            return

        # 4. VELİ GİRİŞ KODU KONTROLÜ (VELI-XXXXXX)
        p_stmt = select(Student).where(Student.parent_code == clean_code)
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
            await message.answer(success_txt, reply_markup=reply_kb, parse_mode="Markdown")
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

    await query.message.edit_text(
        get_text("req_name_prompt", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="act_req_access")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.req_name)
    await query.answer()

@router.message(Form.req_name)
async def process_req_name(message: Message, state: FSMContext):
    full_name = message.text.strip()
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
        if not admin_user or admin_user.role != "admin":
            await query.answer("Yetkisiz işlem!", show_alert=True)
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
            await render_clean_dashboard(query.message.bot, target_user)
        except Exception:
            pass

        await query.message.edit_text(f"✅ *Talep Onaylandı (# {req.id})*\n{req.full_name} ({req.role}) başarıyla sisteme kaydedildi.", parse_mode="Markdown")
    await query.answer("Onaylandı!")

@router.callback_query(F.data.startswith("adm:rej_req:"))
async def cb_admin_reject_request(query: CallbackQuery):
    req_id = int(query.data.split(":")[2])
    async with AsyncSessionLocal() as session:
        admin_user = await session.get(User, query.from_user.id)
        if not admin_user or admin_user.role != "admin":
            await query.answer("Yetkisiz işlem!", show_alert=True)
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
        app = await session.get(Appointment, app_id)
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

    await query.message.edit_text(
        get_text("prompt_student_name", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.add_student_name)
    await query.answer()

@router.message(Form.add_student_name)
async def process_student_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text.strip())
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

    await query.message.edit_text(
        get_text("prompt_teacher_name", lang),
        reply_markup=InlineKeyboardMarkup(inline_keyboard=[[InlineKeyboardButton(text=get_text("btn_back", lang), callback_data="adm:dashboard")]]),
        parse_mode="Markdown"
    )
    await state.set_state(Form.add_teacher_name)
    await query.answer()

@router.message(Form.add_teacher_name)
async def process_teacher_name(message: Message, state: FSMContext):
    await state.update_data(name=message.text.strip())
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

@router.callback_query(F.data.startswith("adm:show_class:"))
async def cb_show_class_students(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()
        buttons = []
        for s in students:
            buttons.append([InlineKeyboardButton(text=f"👤 {s.full_name} ({s.student_number})", callback_data=f"adm:st_card:{s.id}")])

        buttons.append([InlineKeyboardButton(text=f"📄 {class_name} Şifre Kartları (PDF)", callback_data=f"adm:gen_pdf:{class_name}")])
        buttons.append(get_nav_buttons(lang, back_callback="adm:classes"))

        text = f"🏫 *{class_name} Sınıfı Listesi* ({len(students)} Öğrenci):\nDetay veya şifre işlemleri için öğrenciye tıklayınız:"
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

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
            await session.delete(st)
            await session.commit()
            await query.answer(get_text("student_deleted", lang), show_alert=True)
            query.data = f"adm:show_class:{cls_name}"
            await cb_show_class_students(query)
            return
    await query.answer()

@router.callback_query(F.data == "adm:dashboard")
async def cb_admin_dashboard(query: CallbackQuery, state: FSMContext):
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
        missing_str = "\n".join([f"• {c}" for c in missing]) if missing else "✅"

        text = get_text("cockpit_report", lang, date=today.strftime("%d.%m.%Y"), total=total_students, present=present_count, absent=absent_count, missing_cnt=len(missing), missing=missing_str)
        buttons = [get_nav_buttons(lang)]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

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

    bot: Bot = message.bot
    file_info = await bot.get_file(message.document.file_id)
    file_bytes = await bot.download_file(file_info.file_path)

    count, out_excel = await process_student_excel(file_bytes.read())
    file = BufferedInputFile(out_excel.read(), filename="Giris_Kodlari_Uretildi.xlsx")
    await message.answer_document(file, caption=get_text("excel_done", lang, count=count), parse_mode="Markdown")
    out_excel.close()

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

            await query.message.answer(get_text("medical_approved", lang))
    await query.answer()
    await cb_admin_dashboard(query, None)

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
            await query.message.answer(get_text("medical_rejected", lang))
    await query.answer()
    await cb_admin_dashboard(query, None)

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
    await state.clear()
    sent_cnt = 0

    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

        users = (await session.execute(select(User.telegram_id))).scalars().all()
        for u_id in users:
            try:
                await message.bot.send_message(chat_id=u_id, text=f"📢 *OKUL DUYURUSU*\n\n{b_text}", parse_mode="Markdown")
                sent_cnt += 1
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

        ATTENDANCE_CACHE[query.from_user.id] = {s.id: False for s in students}
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

@router.callback_query(F.data == "tch:hw_classes")
async def cb_hw_classes(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        if not classes:
            classes = ["9-A"]

        buttons = []
        for c in classes:
            buttons.append([InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"hw_cls:{c}")])
        buttons.append([InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="adm:dashboard")])
        await query.message.edit_text(get_text("prompt_hw_class", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons))
    await query.answer()

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

@router.callback_query(F.data.startswith("set_child:"))
async def cb_set_child(query: CallbackQuery):
    child_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
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
    # Yönetici
    "rk_cockpit": "act_cockpit",
    "rk_classes": "act_classes",
    "rk_requests": "act_requests",
    # Öğretmen
    "rk_attendance": "act_attendance",
    "rk_grade": "act_grade",
    "rk_homework": "act_homework",
    # Veli & Öğrenci
    "rk_report": "act_report",
    "rk_appointments": "act_appointments",
    "rk_schedule": "act_schedule",
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

@router.message(F.text)
async def global_text_router(message: Message, state: FSMContext):
    action = match_reply_button(message.text)
    user_id = message.from_user.id

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        if not user:
            user = User(telegram_id=user_id, language="tr", role="guest")
            session.add(user)
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
            await message.answer("🔑 *Lütfen giriş kodunu yazınız:* (Örn: `HCA-123456`, `VELI-123456`, `OGR-123456`)", parse_mode="Markdown")
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

    # Eğer misafir kullanıcı doğrudan şifresini yazdıysa
    if user.role == "guest":
        await process_auth_code_string(message.text, user_id, message, state)
        return

# ======================================================================
# 15. FASTAPI VE ARKA PLAN DÖNGÜLERİ (LIFESPAN & PINGER)
# ======================================================================

bot = Bot(token=BOT_TOKEN)
dp = Dispatcher()
dp.include_router(router)

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
            if now_local.hour == 18 and now_local.minute == 30:
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
    print("--> [4/5] Gecikmeli yoklama, Self-Pinger ve 18:30 Akşam Brifingi aktif.")
    print("--> [5/5] SİSTEM CANLI VE HAZIR.")
    print("=" * 60)

    yield

    t1.cancel()
    t2.cancel()
    t3.cancel()
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
        background_tasks.add_task(dp.feed_update, bot, telegram_update)
    except Exception as e:
        print(f"--> [HATA] aiogram Update çeviri hatası: {e}")

    return {"ok": True}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, log_level="info")

