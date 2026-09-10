# -*- coding: utf-8 -*-
"""
OKUL YÖNETİM TELEGRAM BOTU - TEK DOSYA TAM SİSTEM MİMARİSİ
Altyapı: FastAPI + aiogram 3.x Webhook + PostgreSQL (Neon / Supabase) / SQLite Çift Motor Kalkanı
Dil Desteği: Türkçe (TR), Русский (RU), O'zbekcha (UZ), English (EN)
Render Free Tier Uyumlu (512 MB RAM, 0.1 vCPU, Ephemeral Disk)
"""

import os
import io
import random
import string
import asyncio
from datetime import datetime, timedelta
from contextlib import asynccontextmanager

from fastapi import FastAPI, Request, BackgroundTasks, status
from fastapi.responses import JSONResponse
from aiogram import Bot, Dispatcher, Router, F
from aiogram.filters import CommandStart, Command
from aiogram.types import (
    Update, Message, CallbackQuery, BufferedInputFile,
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup

from sqlalchemy import (
    BigInteger, Integer, String, Boolean, DateTime, Date, ForeignKey, Float, Text, select, func
)
from sqlalchemy.orm import DeclarativeBase, Mapped, mapped_column
from sqlalchemy.ext.asyncio import AsyncSession, create_async_engine, async_sessionmaker

import openpyxl
from reportlab.lib.pagesizes import A4
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib import colors

# ======================================================================
# 1. ORTAM DEĞİŞKENLERİ VE YAPILANDIRMA
# ======================================================================

BOT_TOKEN = os.getenv("BOT_TOKEN", "")
WEBHOOK_URL = os.getenv("WEBHOOK_URL", "")
WEBHOOK_SECRET = os.getenv("WEBHOOK_SECRET", "okul_bot_secret_token_2026")
DATABASE_URL = os.getenv("DATABASE_URL", "sqlite+aiosqlite:///school.db")
ADMIN_IDS = [int(x.strip()) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()]

# ======================================================================
# 2. VERİTABANI MOTORU (ÇİFT MOTOR KALKANI - SQLITE & POSTGRESQL)
# ======================================================================

# Neon veya Supabase linkleri 'postgres://' formatında gelirse 'postgresql+asyncpg://' yapar
if DATABASE_URL.startswith("postgres://"):
    DATABASE_URL = DATABASE_URL.replace("postgres://", "postgresql+asyncpg://", 1)
elif DATABASE_URL.startswith("postgresql://") and not DATABASE_URL.startswith("postgresql+asyncpg://"):
    DATABASE_URL = DATABASE_URL.replace("postgresql://", "postgresql+asyncpg://", 1)

# SQLite için havuz parametreleri kapatılır, PostgreSQL için havuz kalkanı devreye girer
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
    role: Mapped[str] = mapped_column(String(20), default="guest")
    language: Mapped[str] = mapped_column(String(5), default="tr")
    full_name: Mapped[str] = mapped_column(String(100), nullable=True)
    failed_attempts: Mapped[int] = mapped_column(Integer, default=0)
    locked_until: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    is_blacklisted: Mapped[bool] = mapped_column(Boolean, default=False)
    current_child_id: Mapped[int] = mapped_column(Integer, nullable=True)
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
    date: Mapped[datetime.date] = mapped_column(Date, default=datetime.utcnow().date, index=True)
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
    score: Mapped[float] = mapped_column(Float)
    badge: Mapped[str] = mapped_column(String(10))
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
    message_text: Mapped[Text] = mapped_column(Text)
    acknowledged_at: Mapped[datetime] = mapped_column(DateTime, nullable=True)
    created_at: Mapped[datetime] = mapped_column(DateTime, default=datetime.utcnow)

class SystemSetting(Base):
    __tablename__ = "system_settings"
    key: Mapped[str] = mapped_column(String(50), primary_key=True)
    value: Mapped[str] = mapped_column(String(255))

async def init_db():
    async with engine.begin() as conn:
        await conn.run_sync(Base.metadata.create_all)

# ======================================================================
# 3. ÇOK DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {
    "tr": {
        "lang_select": "Lütfen bir dil seçiniz / Пожалуйста, выберите язык / Tilni tanlang / Please select a language:",
        "lang_changed": "Dil başarıyla güncellendi: 🇹🇷 Türkçe",
        "welcome_guest": "🎓 *Okul Yönetim Sistemine Hoş Geldiniz.*\n\nLütfen size okul idaresi tarafından verilen giriş kodunu (Örn: `VELI-1234`, `OGR-5678`, `HCA-9012`) yazınız:",
        "auth_success": "✅ Giriş başarılı. Rolünüz: *{role}*",
        "auth_failed": "❌ Geçersiz kod! Kalan deneme hakkınız: {remaining}",
        "auth_locked": "⛔ Çok fazla hatalı deneme yaptınız. Hesabınız 1 saat süreyle kilitlendi.",
        "auth_blacklisted": "🚫 Hesabınız güvenlik nedeniyle kalıcı olarak askıya alınmıştır.",
        "menu_parent": "👨‍👩‍👧‍👦 *Veli Paneli*\nSeçili Öğrenci: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *Öğretmen Paneli*\nBranş: *{subject}*",
        "menu_student": "🎓 *Öğrenci Paneli*\nSınıf: *{class_name}* | No: *{student_number}*",
        "menu_admin": "⚡ *Yönetim Kokpiti (Admin)*\nHoş geldiniz.",
        "btn_parent_dashboard": "📊 Durum Paneli",
        "btn_parent_switch_child": "👶 Çocuk Değiştir",
        "btn_parent_settings": "⚙️ Ayarlar / Dil",
        "btn_teacher_attendance": "📋 Hızlı Yoklama",
        "btn_teacher_grades": "📝 Not Girişi",
        "btn_teacher_homework": "📢 Ödev / Duyuru",
        "btn_student_grades": "📊 Karnem / Notlar",
        "btn_student_exams": "📝 Aktif Sınavlar",
        "btn_student_upload_hw": "📤 Ödev Yükle",
        "btn_admin_cockpit": "⚡ Sabah Kokpiti",
        "btn_admin_students": "👥 Öğrenci Yönetimi",
        "btn_admin_announcement": "📢 Toplu Duyuru",
        "btn_admin_maintenance": "⚙️ Sistem & Bakım",
        "btn_back": "⬅️ Geri",
        "btn_main_menu": "🏠 Ana Menü",
        "btn_acknowledged": "✅ Okudum / Bilgilendirildim",
        "acknowledged_toast": "Bildirim onayınız kaydedildi.",
        "report_card_title": "📊 *Akademik Durum ve Devamsızlık*",
        "attendance_summary": "📌 Toplam Devamsızlık: *{absent} gün* (Özürlü: {excused} gün)",
        "grades_list": "📝 *Ders Notları:*",
        "no_grades": "Henüz girilmiş bir ders notu bulunmamaktadır.",
        "upload_medical_prompt": "Lütfen sağlık raporu veya mazeret belgesinin fotoğrafını gönderiniz:",
        "medical_uploaded": "Sağlık raporu okul idaresine iletildi. İncelendiğinde bildirim alacaksınız.",
        "select_class_attendance": "Lütfen yoklama alacağınız sınıfı seçiniz:",
        "attendance_started": "📋 *{class_name} Yoklaması*\nHerkes varsayılan olarak 'Geldi' kabul edilir. Gelmeyen öğrencinin üzerine dokunarak 'Gelmedi' yapınız, ardından kaydediniz.",
        "btn_save_attendance": "💾 Yoklamayı Onayla ve Kaydet",
        "attendance_saved": "✅ Yoklama kaydedildi. 15 dakikalık düzeltme penceresi başlatıldı. Süre bitiminde velilere otomatik bildirim gidecektir.",
        "cockpit_report": "⚡ *Sabah Kokpiti Özeti ({date})*\n\n🏫 Toplam Öğrenci: {total_students}\n✅ Gelen: {present_count}\n❌ Gelmeyen: {absent_count}\n\n⚠️ *Yoklama Girmeyen Sınıflar ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Excel işlendi. {created_count} öğrenci sisteme eklendi. Üretilen erişim şifreleri ekteki dosyada yer almaktadır.",
        "medical_request_admin": "🏥 *Yeni Mazeret Raporu*\nÖğrenci: *{student_name}* ({class_name})\nVeli ID: `{parent_id}`",
        "panic_mode_status": "🚨 Panik Modu (Sistem Bakımı): *{status}*",
        "btn_toggle_panic": "🚨 Bakım Modunu Aç/Kapat"
    },
    "ru": {
        "lang_select": "Пожалуйста, выберите язык:",
        "lang_changed": "Язык успешно изменен: 🇷🇺 Русский",
        "welcome_guest": "🎓 *Добро пожаловать в систему управления школой.*\n\nПожалуйста, введите код доступа, выданный администрацией (Напр: `VELI-1234`, `OGR-5678`, `HCA-9012`):",
        "auth_success": "✅ Авторизация успешна. Ваша роль: *{role}*",
        "auth_failed": "❌ Неверный код! Осталось попыток: {remaining}",
        "auth_locked": "⛔ Слишком много неверных попыток. Аккаунт заблокирован на 1 час.",
        "auth_blacklisted": "🚫 Ваш аккаунт навсегда заблокирован из соображений безопасности.",
        "menu_parent": "👨‍👩‍👧‍👦 *Панель родителя*\nВыбран ученик: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *Панель учителя*\nПредмет: *{subject}*",
        "menu_student": "🎓 *Панель ученика*\nКласс: *{class_name}* | №: *{student_number}*",
        "menu_admin": "⚡ *Панель администратора*\nДобро пожаловать.",
        "btn_parent_dashboard": "📊 Панель успеваемости",
        "btn_parent_switch_child": "👶 Сменить ребенка",
        "btn_parent_settings": "⚙️ Настройки / Язык",
        "btn_teacher_attendance": "📋 Быстрая перекличка",
        "btn_teacher_grades": "📝 Выставление оценок",
        "btn_teacher_homework": "📢 Д/З и объявления",
        "btn_student_grades": "📊 Мой табель",
        "btn_student_exams": "📝 Активные тесты",
        "btn_student_upload_hw": "📤 Сдать Д/З",
        "btn_admin_cockpit": "⚡ Утренний статус",
        "btn_admin_students": "👥 Управление учениками",
        "btn_admin_announcement": "📢 Рассылка",
        "btn_admin_maintenance": "⚙️ Система и сервис",
        "btn_back": "⬅️ Назад",
        "btn_main_menu": "🏠 Главное меню",
        "btn_acknowledged": "✅ Ознакомлен(а)",
        "acknowledged_toast": "Ваше подтверждение зафиксировано.",
        "report_card_title": "📊 *Успеваемость и посещаемость*",
        "attendance_summary": "📌 Пропущено занятий: *{absent} дн.* (По уважительной: {excused} дн.)",
        "grades_list": "📝 *Оценки по предметам:*",
        "no_grades": "Оценки пока не выставлены.",
        "upload_medical_prompt": "Пожалуйста, отправьте фото медицинской справки или заявления:",
        "medical_uploaded": "Справка отправлена администрации школы.",
        "select_class_attendance": "Выберите класс для переклички:",
        "attendance_started": "📋 *Перекличка класса {class_name}*\nВсе отмечены как 'Присутствует'. Нажмите на отсутствующих, затем сохраните.",
        "btn_save_attendance": "💾 Сохранить перекличку",
        "attendance_saved": "✅ Перекличка сохранена. Окно правок (15 минут) запущено. Затем родители получат уведомления.",
        "cockpit_report": "⚡ *Утренняя сводка ({date})*\n\n🏫 Всего учеников: {total_students}\n✅ Присутствуют: {present_count}\n❌ Отсутствуют: {absent_count}\n\n⚠️ *Классы без переклички ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Файл обработан. Добавлено учеников: {created_count}. Коды доступа в прикрепленном файле.",
        "medical_request_admin": "🏥 *Новая медицинская справка*\nУченик: *{student_name}* ({class_name})\nID родителя: `{parent_id}`",
        "panic_mode_status": "🚨 Режим обслуживания: *{status}*",
        "btn_toggle_panic": "🚨 Переключить режим"
    },
    "uz": {
        "lang_select": "Iltimos, tilni tanlang:",
        "lang_changed": "Til muvaffaqiyatli yangilandi: 🇺🇿 O'zbekcha",
        "welcome_guest": "🎓 *Maktab boshqaruv tizimiga xush kelibsiz.*\n\nIltimos, maktab ma'muriyati bergan maxsus kodni kiriting (Masalan: `VELI-1234`, `OGR-5678`, `HCA-9012`):",
        "auth_success": "✅ Kirish muvaffaqiyatli. Sizning rolingiz: *{role}*",
        "auth_failed": "❌ Noto'g'ri kod! Qolgan urinishlar soni: {remaining}",
        "auth_locked": "⛔ Xato urinishlar ko'payib ketdi. Hisobingiz 1 soatga bloklandi.",
        "auth_blacklisted": "🚫 Hisobingiz xavfsizlik sababli doimiy ravishda to'xtatildi.",
        "menu_parent": "👨‍👩‍👧‍👦 *Ota-ona paneli*\nTanlangan o'quvchi: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *O'qituvchi paneli*\nFan: *{subject}*",
        "menu_student": "🎓 *O'quvchi paneli*\nSinf: *{class_name}* | №: *{student_number}*",
        "menu_admin": "⚡ *Boshqaruv markazi (Admin)*\nXush kelibsiz.",
        "btn_parent_dashboard": "📊 Holat paneli",
        "btn_parent_switch_child": "👶 Farzandni almashtirish",
        "btn_parent_settings": "⚙️ Sozlamalar / Til",
        "btn_teacher_attendance": "📋 Tezkor davomat",
        "btn_teacher_grades": "📝 Baho qo'yish",
        "btn_teacher_homework": "📢 Vazifa / E'lon",
        "btn_student_grades": "📊 Baholarim / Tabel",
        "btn_student_exams": "📝 Faol testlar",
        "btn_student_upload_hw": "📤 Vazifa yuklash",
        "btn_admin_cockpit": "⚡ Tonggi hisobot",
        "btn_admin_students": "👥 O'quvchilarni boshqarish",
        "btn_admin_announcement": "📢 Ommaviy xabar",
        "btn_admin_maintenance": "⚙️ Tizim va ta'mirlash",
        "btn_back": "⬅️ Orqaga",
        "btn_main_menu": "🏠 Asosiy menyu",
        "btn_acknowledged": "✅ O'qidim / Xabardorman",
        "acknowledged_toast": "Tasdig'ingiz qabul qilindi.",
        "report_card_title": "📊 *O'quv ko'rsatkichlari va davomat*",
        "attendance_summary": "📌 Qoldirilgan darslar: *{absent} kun* (Sababli: {excused} kun)",
        "grades_list": "📝 *Fan baholari:*",
        "no_grades": "Hozircha baholar mavjud emas.",
        "upload_medical_prompt": "Iltimos, tibbiy ma'lumotnoma rasmini yuboring:",
        "medical_uploaded": "Ma'lumotnoma ma'muriyatga yuborildi.",
        "select_class_attendance": "Davomat olinadigan sinfni tanlang:",
        "attendance_started": "📋 *{class_name} davomati*\nHamma 'Keldi' deb belgilangan. Kelmagan o'quvchi ustiga bosib 'Kelmadi' qiling va saqlang.",
        "btn_save_attendance": "💾 Davomatni saqlash",
        "attendance_saved": "✅ Davomat saqlandi. 15 daqiqalik tuzatish oynasi boshlandi.",
        "cockpit_report": "⚡ *Tonggi umumiy hisobot ({date})*\n\n🏫 Jami o'quvchilar: {total_students}\n✅ Kelganlar: {present_count}\n❌ Kelmaganlar: {absent_count}\n\n⚠️ *Davomat olinmagan sinflar ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Fayl yuklandi. Qo'shildi: {created_count} ta o'quvchi. Kodlar biriktirilgan faylda.",
        "medical_request_admin": "🏥 *Yangi tibbiy ma'lumotnoma*\nO'quvchi: *{student_name}* ({class_name})\nOta-ona ID: `{parent_id}`",
        "panic_mode_status": "🚨 Texnik xizmat rejimi: *{status}*",
        "btn_toggle_panic": "🚨 Rejimni o'zgartirish"
    },
    "en": {
        "lang_select": "Please select your language:",
        "lang_changed": "Language successfully updated: 🇬🇧 English",
        "welcome_guest": "🎓 *Welcome to School Management Ecosystem.*\n\nPlease enter the access code provided by administration (e.g., `VELI-1234`, `OGR-5678`, `HCA-9012`):",
        "auth_success": "✅ Authentication successful. Your role: *{role}*",
        "auth_failed": "❌ Invalid code! Remaining attempts: {remaining}",
        "auth_locked": "⛔ Too many failed attempts. Your account is locked for 1 hour.",
        "auth_blacklisted": "🚫 Your account has been permanently blacklisted for security violations.",
        "menu_parent": "👨‍👩‍👧‍👦 *Parent Dashboard*\nSelected Student: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *Teacher Dashboard*\nSubject: *{subject}*",
        "menu_student": "🎓 *Student Dashboard*\nClass: *{class_name}* | Roll: *{student_number}*",
        "menu_admin": "⚡ *Administration Cockpit*\nWelcome back.",
        "btn_parent_dashboard": "📊 Overview",
        "btn_parent_switch_child": "👶 Switch Child",
        "btn_parent_settings": "⚙️ Settings / Language",
        "btn_teacher_attendance": "📋 Fast Attendance",
        "btn_teacher_grades": "📝 Enter Grades",
        "btn_teacher_homework": "📢 Homework / Notice",
        "btn_student_grades": "📊 My Grades",
        "btn_student_exams": "📝 Active Quizzes",
        "btn_student_upload_hw": "📤 Submit Homework",
        "btn_admin_cockpit": "⚡ Morning Cockpit",
        "btn_admin_students": "👥 Student Directory",
        "btn_admin_announcement": "📢 Broadcast",
        "btn_admin_maintenance": "⚙️ System & Health",
        "btn_back": "⬅️ Back",
        "btn_main_menu": "🏠 Main Menu",
        "btn_acknowledged": "✅ Read / Acknowledged",
        "acknowledged_toast": "Acknowledgment recorded.",
        "report_card_title": "📊 *Academic Record & Attendance*",
        "attendance_summary": "📌 Total Absences: *{absent} days* (Excused: {excused} days)",
        "grades_list": "📝 *Grade Book:*",
        "no_grades": "No grades recorded yet.",
        "upload_medical_prompt": "Please send a photo of the medical report or excuse letter:",
        "medical_uploaded": "Report submitted to school administration.",
        "select_class_attendance": "Select class for attendance:",
        "attendance_started": "📋 *Attendance: {class_name}*\nAll students are marked 'Present' by default. Tap absent students to toggle, then click save.",
        "btn_save_attendance": "💾 Confirm & Save",
        "attendance_saved": "✅ Attendance recorded. 15-minute edit window started. Parents will be notified automatically thereafter.",
        "cockpit_report": "⚡ *Morning Status Brief ({date})*\n\n🏫 Total Students: {total_students}\n✅ Present: {present_count}\n❌ Absent: {absent_count}\n\n⚠️ *Pending Attendance Classes ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Processed. Added {created_count} students. Generated codes are in the attached file.",
        "medical_request_admin": "🏥 *New Medical Excuse Submission*\nStudent: *{student_name}* ({class_name})\nParent ID: `{parent_id}`",
        "panic_mode_status": "🚨 Maintenance Mode: *{status}*",
        "btn_toggle_panic": "🚨 Toggle Maintenance"
    }
}

def get_text(key: str, lang: str = "tr", **kwargs) -> str:
    selected_lang = lang if lang in LOCALES else "tr"
    template = LOCALES[selected_lang].get(key) or LOCALES["tr"].get(key) or f"[{key}]"
    if kwargs:
        return template.format(**kwargs)
    return template

# ======================================================================
# 4. KLAVYELER VE ARAYÜZ (3 TIK KURALI)
# ======================================================================

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

def get_role_reply_kb(role: str, lang: str = "tr") -> ReplyKeyboardMarkup:
    keyboard = []
    if role == "parent":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_parent_dashboard", lang)),
                KeyboardButton(text=get_text("btn_parent_switch_child", lang))
            ],
            [KeyboardButton(text=get_text("btn_parent_settings", lang))]
        ]
    elif role == "teacher":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_teacher_attendance", lang)),
                KeyboardButton(text=get_text("btn_teacher_grades", lang))
            ],
            [KeyboardButton(text=get_text("btn_teacher_homework", lang))]
        ]
    elif role == "student":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_student_grades", lang)),
                KeyboardButton(text=get_text("btn_student_exams", lang))
            ],
            [KeyboardButton(text=get_text("btn_student_upload_hw", lang))]
        ]
    elif role == "admin":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_admin_cockpit", lang)),
                KeyboardButton(text=get_text("btn_admin_students", lang))
            ],
            [
                KeyboardButton(text=get_text("btn_admin_announcement", lang)),
                KeyboardButton(text=get_text("btn_admin_maintenance", lang))
            ]
        ]
    else:
        keyboard = [[KeyboardButton(text="/start")]]

    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def get_nav_buttons(lang: str = "tr", back_callback: str = "nav:main") -> list:
    return [
        InlineKeyboardButton(text=get_text("btn_back", lang), callback_data=back_callback),
        InlineKeyboardButton(text=get_text("btn_main_menu", lang), callback_data="nav:main")
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
        InlineKeyboardButton(text=get_text("btn_save_attendance", lang), callback_data=f"att_save:{class_name}")
    ])
    inline_keyboard.append(get_nav_buttons(lang, back_callback="nav:main"))
    return InlineKeyboardMarkup(inline_keyboard=inline_keyboard)

def get_ack_notification_kb(notification_id: int, lang: str = "tr") -> InlineKeyboardMarkup:
    return InlineKeyboardMarkup(inline_keyboard=[
        [InlineKeyboardButton(text=get_text("btn_acknowledged", lang), callback_data=f"ack_notif:{notification_id}")]
    ])

# ======================================================================
# 5. YARDIMCI SERVİSLER (EXCEL, PDF VE BİLDİRİM İŞÇİSİ)
# ======================================================================

def generate_secure_code(prefix: str) -> str:
    nums = "".join(random.choices(string.digits, k=4))
    return f"{prefix}-{nums}"

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
            class_name = str(row).strip() if len(row) > 1 and row else "Genel"
            student_no = str(row).strip() if len(row) > 2 and row else "0"

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

async def generate_classroom_pdf_cards(class_name: str) -> io.BytesIO:
    async with AsyncSessionLocal() as session:
        stmt = select(Student).where(Student.class_name == class_name).order_by(Student.full_name)
        result = await session.execute(stmt)
        students = result.scalars().all()

    pdf_buffer = io.BytesIO()
    doc = SimpleDocTemplate(
        pdf_buffer,
        pagesize=A4,
        rightMargin=20,
        leftMargin=20,
        topMargin=20,
        bottomMargin=20
    )
    styles = getSampleStyleSheet()
    title_style = ParagraphStyle(
        name="CardTitle", parent=styles["Heading2"], fontSize=12, leading=14, textColor=colors.HexColor("#1A365D")
    )
    body_style = ParagraphStyle(
        name="CardBody", parent=styles["Normal"], fontSize=9, leading=11, textColor=colors.HexColor("#2D3748")
    )

    story = [
        Paragraph(f"<b>{class_name} SINIFI - TELEGRAM GİRİŞ KARTLARI</b>", title_style),
        Spacer(1, 15)
    ]

    card_data = []
    for s in students:
        text = (
            f"<b>Öğrenci:</b> {s.full_name} (No: {s.student_number})<br/>"
            f"<b>Öğrenci Kodu:</b> <code>{s.student_code}</code><br/>"
            f"<b>Veli Kodu:</b> <code>{s.parent_code}</code><br/>"
            f"<i>Telegram Botu üzerinden giriş yapınız.</i>"
        )
        card_data.append([Paragraph(text, body_style)])

    if card_data:
        table = Table(card_data, colWidths=[540])
        table.setStyle(TableStyle([
            ("BOX", (0, 0), (-1, -1), 1, colors.HexColor("#CBD5E0")),
            ("INNERGRID", (0, 0), (-1, -1), 0.5, colors.HexColor("#E2E8F0")),
            ("TOPPADDING", (0, 0), (-1, -1), 8),
            ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
            ("LEFTPADDING", (0, 0), (-1, -1), 12),
            ("RIGHTPADDING", (0, 0), (-1, -1), 12),
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

                    msg_text = (
                        f"🚨 *DEVAMSIZLIK BİLDİRİMİ*\n\n"
                        f"Öğrenciniz *{st.full_name}*, bugün ({att.date}) okul yoklamasında *GELMEDİ (YOK)* olarak işaretlenmiştir."
                    )
                    
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

# ======================================================================
# 6. TELEGRAM ROUTER VE OLAY YÖNETİCİLERİ
# ======================================================================

router = Router()
ATTENDANCE_CACHE = {}

class Form(StatesGroup):
    waiting_auth_code = State()
    waiting_medical_photo = State()

@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not user:
            await message.answer(get_text("lang_select", "tr"), reply_markup=get_language_inline_kb())
            return

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language))
            return

        if user.locked_until and user.locked_until > datetime.utcnow():
            await message.answer(get_text("auth_locked", user.language))
            return

        if user.role == "guest":
            await message.answer(get_text("welcome_guest", user.language), parse_mode="Markdown")
            await state.set_state(Form.waiting_auth_code)
            return

        await send_role_home(message, user)

@router.callback_query(F.data.startswith("set_lang:"))
async def cb_set_lang(query: CallbackQuery, state: FSMContext):
    lang_code = query.data.split(":")
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if not user:
            role = "admin" if query.from_user.id in ADMIN_IDS else "guest"
            user = User(telegram_id=query.from_user.id, language=lang_code, role=role)
            session.add(user)
        else:
            user.language = lang_code
        await session.commit()

    await query.message.edit_text(get_text("lang_changed", lang_code))
    if user.role == "guest":
        await query.message.answer(get_text("welcome_guest", lang_code), parse_mode="Markdown")
        await state.set_state(Form.waiting_auth_code)
    else:
        await send_role_home(query.message, user)
    await query.answer()

async def send_role_home(message: Message, user: User):
    reply_kb = get_role_reply_kb(user.role, user.language)
    lang = user.language

    async with AsyncSessionLocal() as session:
        if user.role == "parent":
            st_stmt = (
                select(Student)
                .join(ParentStudent, ParentStudent.student_id == Student.id)
                .where(ParentStudent.parent_telegram_id == user.telegram_id)
            )
            children = (await session.execute(st_stmt)).scalars().all()
            if children:
                current_child = next((c for c in children if c.id == user.current_child_id), children[0])
                if user.current_child_id != current_child.id:
                    user.current_child_id = current_child.id
                    await session.commit()
                text = get_text("menu_parent", lang, student_name=current_child.full_name, class_name=current_child.class_name)
            else:
                text = get_text("menu_parent", lang, student_name="Kayıtlı Çocuk Yok", class_name="-")
        elif user.role == "teacher":
            t_stmt = select(Teacher).where(Teacher.telegram_id == user.telegram_id)
            teacher = (await session.execute(t_stmt)).scalar_one_or_none()
            subj = teacher.subject if teacher else "Öğretmen"
            text = get_text("menu_teacher", lang, subject=subj)
        elif user.role == "student":
            st_stmt = select(Student).where(Student.student_telegram_id == user.telegram_id)
            st = (await session.execute(st_stmt)).scalar_one_or_none()
            cls = st.class_name if st else "-"
            num = st.student_number if st else "-"
            text = get_text("menu_student", lang, class_name=cls, student_number=num)
        elif user.role == "admin":
            text = get_text("menu_admin", lang)
        else:
            text = get_text("welcome_guest", lang)

    await message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")

@router.message(Form.waiting_auth_code)
async def handle_auth_code(message: Message, state: FSMContext):
    code = message.text.strip().upper()
    user_id = message.from_user.id

    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        t_stmt = select(Teacher).where(Teacher.auth_code == code, Teacher.is_code_burned == False)
        teacher = (await session.execute(t_stmt)).scalar_one_or_none()
        if teacher:
            teacher.is_code_burned = True
            teacher.telegram_id = user_id
            user.role = "teacher"
            user.full_name = teacher.full_name
            user.failed_attempts = 0
            await session.commit()
            await state.clear()
            await message.answer(get_text("auth_success", lang, role="Öğretmen"))
            await send_role_home(message, user)
            return

        s_stmt = select(Student).where(Student.student_code == code, Student.is_student_code_burned == False)
        student = (await session.execute(s_stmt)).scalar_one_or_none()
        if student:
            student.is_student_code_burned = True
            student.student_telegram_id = user_id
            user.role = "student"
            user.full_name = student.full_name
            user.failed_attempts = 0
            await session.commit()
            await state.clear()
            await message.answer(get_text("auth_success", lang, role="Öğrenci"))
            await send_role_home(message, user)
            return

        p_stmt = select(Student).where(Student.parent_code == code)
        p_student = (await session.execute(p_stmt)).scalar_one_or_none()
        if p_student:
            user.role = "parent"
            user.failed_attempts = 0
            if not user.current_child_id:
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
            await message.answer(get_text("auth_success", lang, role="Veli"))
            await send_role_home(message, user)
            return

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

@router.message(F.text.in_(["⚡ Sabah Kokpiti", "⚡ Утренний статус", "⚡ Tonggi hisobot", "⚡ Morning Cockpit"]))
async def admin_morning_cockpit(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    async with AsyncSessionLocal() as session:
        today = datetime.utcnow().date()
        total_students = (await session.execute(select(func.count(Student.id)))).scalar() or 0
        present_count = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "present"))).scalar() or 0
        absent_count = (await session.execute(select(func.count(Attendance.id)).where(Attendance.date == today, Attendance.status == "absent"))).scalar() or 0

        all_classes_stmt = select(Student.class_name).distinct()
        all_classes = (await session.execute(all_classes_stmt)).scalars().all()
        taken_classes_stmt = select(Attendance.class_name).where(Attendance.date == today).distinct()
        taken_classes = (await session.execute(taken_classes_stmt)).scalars().all()

        missing = [c for c in all_classes if c not in taken_classes]
        missing_str = "\n".join([f"• {c}" for c in missing]) if missing else "✅ Tüm sınıfların yoklaması tamamlandı."

        text = get_text(
            "cockpit_report", "tr",
            date=today.strftime("%d.%m.%Y"),
            total_students=total_students,
            present_count=present_count,
            absent_count=absent_count,
            missing_classes_count=len(missing),
            missing_classes=missing_str
        )
        await message.answer(text, parse_mode="Markdown")

@router.message(Command("kodlar"))
async def admin_pdf_cards_command(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    args = message.text.split(maxsplit=1)
    if len(args) < 2:
        await message.answer("Lütfen sınıf belirtiniz: `/kodlar 9-A`", parse_mode="Markdown")
        return
    class_name = args.strip()
    pdf_buffer = await generate_classroom_pdf_cards(class_name)
    file = BufferedInputFile(pdf_buffer.read(), filename=f"{class_name}_Sifre_Kartlari.pdf")
    await message.answer_document(file, caption=f"📄 *{class_name}* sınıfı şifre kartları hazır.")
    pdf_buffer.close()

@router.message(F.document, F.document.file_name.endswith(".xlsx"))
async def admin_excel_upload(message: Message):
    if message.from_user.id not in ADMIN_IDS:
        return
    bot: Bot = message.bot
    file_info = await bot.get_file(message.document.file_id)
    file_bytes = await bot.download_file(file_info.file_path)

    count, out_excel = await process_student_excel(file_bytes.read())
    file = BufferedInputFile(out_excel.read(), filename="Giris_Kodlari_Uretildi.xlsx")
    await message.answer_document(file, caption=get_text("excel_processed", "tr", created_count=count))
    out_excel.close()

@router.message(F.text.in_(["📋 Hızlı Yoklama", "📋 Быстрая перекличка", "📋 Tezkor davomat", "📋 Fast Attendance"]))
async def teacher_start_attendance(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not user or user.role != "teacher":
            return
        stmt = select(Student.class_name).distinct()
        classes = (await session.execute(stmt)).scalars().all()
        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"att_class:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row:
            buttons.append(row)
        buttons.append(get_nav_buttons(user.language, back_callback="nav:main"))
        await message.answer(
            get_text("select_class_attendance", user.language),
            reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
        )

@router.callback_query(F.data.startswith("att_class:"))
async def cb_attendance_class(query: CallbackQuery):
    class_name = query.data.split(":")
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        stmt = select(Student).where(Student.class_name == class_name).order_by(Student.full_name)
        students = (await session.execute(stmt)).scalars().all()

        ATTENDANCE_CACHE[query.from_user.id] = {s.id: False for s in students}
        kb = get_attendance_grid_kb(students, ATTENDANCE_CACHE[query.from_user.id], class_name, lang=lang)
        await query.message.edit_text(
            get_text("attendance_started", lang, class_name=class_name),
            reply_markup=kb,
            parse_mode="Markdown"
        )
    await query.answer()

@router.callback_query(F.data.startswith("att_toggle:"))
async def cb_attendance_toggle(query: CallbackQuery):
    student_id = int(query.data.split(":"))
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
    class_name = query.data.split(":")
    cache = ATTENDANCE_CACHE.pop(query.from_user.id, {})
    now = datetime.utcnow()
    notify_time = now + timedelta(minutes=15)

    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        for st_id, is_absent in cache.items():
            status_val = "absent" if is_absent else "present"
            att = Attendance(
                student_id=st_id,
                class_name=class_name,
                date=now.date(),
                status=status_val,
                teacher_id=query.from_user.id,
                notify_at=notify_time,
                is_notified=False
            )
            session.add(att)
        await session.commit()

        await query.message.edit_text(get_text("attendance_saved", lang))
    await query.answer()

@router.message(F.text.in_(["📊 Durum Paneli", "📊 Панель успеваемости", "📊 Holat paneli", "📊 Overview"]))
async def parent_dashboard(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not user or user.role != "parent" or not user.current_child_id:
            return
        lang = user.language
        st = await session.get(Student, user.current_child_id)

        att_stmt = select(Attendance).where(Attendance.student_id == st.id)
        attendances = (await session.execute(att_stmt)).scalars().all()
        absent_count = sum(1 for a in attendances if a.status == "absent")
        excused_count = sum(1 for a in attendances if a.status == "excused")

        gr_stmt = select(Grade).where(Grade.student_id == st.id).order_by(Grade.created_at.desc())
        grades = (await session.execute(gr_stmt)).scalars().all()

        grades_text = "\n".join([f"• {g.subject}: *{g.score}* ({g.badge})" for g in grades]) if grades else get_text("no_grades", lang)

        text = (
            f"{get_text('report_card_title', lang)}\n"
            f"👤 Öğrenci: *{st.full_name}* ({st.class_name})\n\n"
            f"{get_text('attendance_summary', lang, absent=absent_count, excused=excused_count)}\n\n"
            f"{get_text('grades_list', lang)}\n{grades_text}"
        )
        buttons = [
            [InlineKeyboardButton(text="🏥 Mazeret / Rapor Yükle", callback_data="upload_medical_init")],
            get_nav_buttons(lang, back_callback="nav:main")
        ]
        await message.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")

@router.callback_query(F.data == "upload_medical_init")
async def cb_upload_medical_init(query: CallbackQuery, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"
        await query.message.answer(get_text("upload_medical_prompt", lang))
        await state.set_state(Form.waiting_medical_photo)
    await query.answer()

@router.message(Form.waiting_medical_photo, F.photo)
async def handle_medical_photo(message: Message, state: FSMContext):
    photo_file_id = message.photo[-1].file_id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, user.current_child_id)

        report = MedicalReport(
            student_id=st.id,
            parent_telegram_id=user.telegram_id,
            file_id=photo_file_id,
            caption=message.caption or "Rapor görseli"
        )
        session.add(report)
        await session.commit()

        for adm_id in ADMIN_IDS:
            try:
                adm_kb = InlineKeyboardMarkup(inline_keyboard=[
                    [
                        InlineKeyboardButton(text="✅ Onayla", callback_data=f"med_appr:{report.id}"),
                        InlineKeyboardButton(text="❌ Reddet", callback_data=f"med_rej:{report.id}")
                    ]
                ])
                await message.bot.send_photo(
                    chat_id=adm_id,
                    photo=photo_file_id,
                    caption=get_text("medical_request_admin", "tr", student_name=st.full_name, class_name=st.class_name, parent_id=user.telegram_id),
                    reply_markup=adm_kb,
                    parse_mode="Markdown"
                )
            except Exception:
                pass

        await message.answer(get_text("medical_uploaded", lang))
    await state.clear()

@router.callback_query(F.data.startswith("ack_notif:"))
async def cb_acknowledge_notification(query: CallbackQuery):
    notif_id = int(query.data.split(":"))
    async with AsyncSessionLocal() as session:
        notif = await session.get(CriticalNotification, notif_id)
        if notif:
            notif.acknowledged_at = datetime.utcnow()
            await session.commit()
    await query.answer(get_text("acknowledged_toast", "tr"), show_alert=True)
    await query.message.edit_reply_markup(reply_markup=None)

@router.callback_query(F.data == "nav:main")
async def cb_nav_main(query: CallbackQuery, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user:
            await query.message.delete()
            await send_role_home(query.message, user)
    await query.answer()

# ======================================================================
# 7. FASTAPI WEBHOOK SUNUCUSU VE LIFESPAN
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

@asynccontextmanager
async def lifespan(app: FastAPI):
    await init_db()
    if WEBHOOK_URL:
        try:
            await bot.set_webhook(
                url=WEBHOOK_URL,
                secret_token=WEBHOOK_SECRET,
                drop_pending_updates=True
            )
        except Exception:
            pass
    worker_task = asyncio.create_task(background_attendance_loop())
    yield
    worker_task.cancel()
    try:
        await bot.delete_webhook()
    except Exception:
        pass
    await bot.session.close()

app = FastAPI(title="OkulYonetimBot", lifespan=lifespan)

@app.get("/health", status_code=status.HTTP_200_OK)
async def health_check():
    return {"status": "ok", "service": "OkulYonetimBot", "uptime": True}

@app.post("/webhook")
async def telegram_webhook(request: Request, background_tasks: BackgroundTasks):
    secret = request.headers.get("X-Telegram-Bot-Api-Secret-Token")
    if secret != WEBHOOK_SECRET:
        return JSONResponse(status_code=status.HTTP_403_FORBIDDEN, content={"error": "Invalid secret"})

    data = await request.json()
    telegram_update = Update.model_validate(data, context={"bot": bot})
    background_tasks.add_task(dp.feed_update, bot, telegram_update)
    return {"ok": True}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, log_level="info")
