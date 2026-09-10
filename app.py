# -*- coding: utf-8 -*-
"""
OKUL YÖNETİM TELEGRAM BOTU - KURUMSAL VE EKSİKSİZ TAM SİSTEM MİMARİSİ (PROD V5)
Altyapı: FastAPI + aiogram 3.x Webhook + PostgreSQL (Neon / Supabase) / SQLite Çift Motor Kalkanı
4 Dilli Kusursuz Arayüz (i18n): 🇹🇷 Türkçe, 🇷🇺 Русский, 🇺🇿 O'zbekcha, 🇬🇧 English
Platform: Render Free Web Service (512 MB RAM, 0.1 vCPU, Ephemeral Disk)
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

ADMIN_IDS = []
for x in os.getenv("ADMIN_IDS", "").split(","):
    clean_x = x.strip().replace("@", "")
    if clean_x.isdigit():
        ADMIN_IDS.append(int(clean_x))

# ======================================================================
# 2. VERİTABANI MOTORU (SQLITE & POSTGRESQL ÇİFT MOTOR KALKANI)
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
# 3. KUSURSUZ 4 DİLLİ METİN SÖZLÜĞÜ (TR, RU, UZ, EN)
# ======================================================================

LOCALES = {
    "tr": {
        "lang_select": "Lütfen tercih ettiğiniz dili seçiniz:",
        "lang_changed": "Dil başarıyla güncellendi: 🇹🇷 Türkçe",
        "welcome_guest": "🎓 *Okul Yönetim Sistemine Hoş Geldiniz.*\n\nLütfen size okul idaresi tarafından verilen giriş kodunu yazınız:",
        "auth_success": "✅ Giriş başarılı. Rolünüz: *{role}*",
        "auth_failed": "❌ Geçersiz giriş kodu! Kalan deneme hakkınız: {remaining}",
        "auth_locked": "⛔ Güvenlik nedeniyle hesabınız 1 saat süreyle kilitlendi.",
        "auth_blacklisted": "🚫 Hesabınız kalıcı olarak askıya alınmıştır.",
        "menu_parent": "👨‍👩‍👧‍👦 *Veli Paneli*\nÖğrenci: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *Öğretmen Paneli*\nBranş: *{subject}*",
        "menu_student": "🎓 *Öğrenci Paneli*\nSınıf: *{class_name}* | No: *{student_number}*",
        "menu_admin": "⚡ *Okul Yönetim Kokpiti (Admin)*\nHoş geldiniz.",
        "admin_stats_text": "📊 *Genel Okul İstatistikleri:*\n• Toplam Sınıf: *{classes_count}*\n• Toplam Öğrenci: *{students_count}*\n• Tarih: *{date}*",
        "btn_admin_cockpit": "📊 Sabah Kokpiti",
        "btn_admin_classes": "🏫 Sınıflar & Öğrenciler",
        "btn_admin_excel": "📥 Excel ile Öğrenci Yükle",
        "btn_admin_pdf": "📄 Şifre Kartları (PDF)",
        "btn_admin_announcement": "📢 Toplu Duyuru",
        "btn_admin_settings": "⚙️ Ayarlar & Dil",
        "btn_parent_dashboard": "📊 Durum Paneli",
        "btn_parent_switch_child": "👶 Çocuk Değiştir",
        "btn_parent_medical": "🏥 Mazeret / Rapor Yükle",
        "btn_teacher_attendance": "📋 Hızlı Yoklama",
        "btn_teacher_grades": "📝 Not Girişi",
        "btn_teacher_homework": "📢 Ödev Panosu",
        "btn_student_grades": "📊 Karnem / Notlar",
        "btn_student_upload_hw": "📤 Ödev Teslim Et",
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
        "excel_instructions": "📥 *Excel ile Öğrenci Yükleme Rehberi*\n\nLütfen bir `.xlsx` dosyasını doğrudan bu sohbete gönderiniz.\nExcel dosyanızın ilk satırı şu sütun başlıklarını içermelidir:\n\n`Ad Soyad` | `Sinif` | `Numara`\n\nÖrnek:\n`Ali Yılmaz` | `9-A` | `101`\n`Ayşe Kaya` | `9-A` | `102`\n\nDosyayı gönderdiğinizde bot tüm şifreleri otomatik üretecek ve size şifreli Excel dosyasını geri teslim edecektir.",
        "no_classes_found": "⚠️ *Henüz kayıtlı bir sınıf bulunmamaktadır.*\n\nÖğrencileri sisteme eklemek için Excel dosyası yükleyebilir veya test için örnek bir sınıf oluşturabilirsiniz.",
        "btn_create_sample_class": "➕ Örnek Sınıf Oluştur (9-A)",
        "sample_class_created": "✅ Örnek 9-A sınıfı (3 öğrenci ve 1 öğretmen) başarıyla oluşturuldu.",
        "select_class_for_pdf": "📄 Şifre kartlarını PDF olarak indirmek istediğiniz sınıfı seçiniz:",
        "pdf_cards_ready": "📄 *{class_name}* sınıfı için kesilip dağıtılmaya hazır şifre kartları ekte sunulmuştur.",
        "medical_request_admin": "🏥 *Yeni Mazeret Raporu*\nÖğrenci: *{student_name}* ({class_name})\nVeli ID: `{parent_id}`",
        "btn_approve_medical": "✅ Onayla (İzinli Say)",
        "btn_reject_medical": "❌ Reddet",
        "class_detail_title": "🏫 *{class_name} Sınıfı Detayı*\nToplam Öğrenci: *{count}*",
        "btn_download_class_pdf": "📄 Sınıfın Şifre Kartlarını İndir (PDF)"
    },
    "ru": {
        "lang_select": "Пожалуйста, выберите предпочитаемый язык:",
        "lang_changed": "Язык успешно изменен: 🇷🇺 Русский",
        "welcome_guest": "🎓 *Добро пожаловать в систему управления школой.*\n\nВведите код доступа, выданный администрацией школы:",
        "auth_success": "✅ Авторизация успешна. Ваша роль: *{role}*",
        "auth_failed": "❌ Неверный код! Осталось попыток: {remaining}",
        "auth_locked": "⛔ Аккаунт заблокирован на 1 час из соображений безопасности.",
        "auth_blacklisted": "🚫 Ваш аккаунт навсегда заблокирован.",
        "menu_parent": "👨‍👩‍👧‍👦 *Панель родителя*\nУченик: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *Панель учителя*\nПредмет: *{subject}*",
        "menu_student": "🎓 *Панель ученика*\nКласс: *{class_name}* | №: *{student_number}*",
        "menu_admin": "⚡ *Панель управления школой (Администратор)*\nДобро пожаловать.",
        "admin_stats_text": "📊 *Общая школьная статистика:*\n• Всего классов: *{classes_count}*\n• Всего учеников: *{students_count}*\n• Дата: *{date}*",
        "btn_admin_cockpit": "📊 Утренний статус",
        "btn_admin_classes": "🏫 Классы и ученики",
        "btn_admin_excel": "📥 Импорт через Excel",
        "btn_admin_pdf": "📄 Карточки с кодами (PDF)",
        "btn_admin_announcement": "📢 Общее объявление",
        "btn_admin_settings": "⚙️ Настройки и язык",
        "btn_parent_dashboard": "📊 Табель успеваемости",
        "btn_parent_switch_child": "👶 Сменить ребенка",
        "btn_parent_medical": "🏥 Отправить справку",
        "btn_teacher_attendance": "📋 Быстрая перекличка",
        "btn_teacher_grades": "📝 Выставление оценок",
        "btn_teacher_homework": "📢 Домашнее задание",
        "btn_student_grades": "📊 Мои оценки",
        "btn_student_upload_hw": "📤 Сдать задание",
        "btn_back": "⬅️ Назад",
        "btn_main_menu": "🏠 Главное меню",
        "btn_acknowledged": "✅ Ознакомлен(а)",
        "acknowledged_toast": "Подтверждение записано.",
        "report_card_title": "📊 *Успеваемость и посещаемость*",
        "attendance_summary": "📌 Пропущено занятий: *{absent} дн.* (По уважительной: {excused} дн.)",
        "grades_list": "📝 *Оценки по предметам:*",
        "no_grades": "Оценки пока не выставлены.",
        "upload_medical_prompt": "Пожалуйста, отправьте фото медицинской справки или заявления:",
        "medical_uploaded": "Справка отправлена администрации школы.",
        "select_class_attendance": "Выберите класс для переклички:",
        "attendance_started": "📋 *Перекличка: класс {class_name}*\nВсе отмечены как 'Присутствует'. Нажмите на отсутствующих, затем сохраните.",
        "btn_save_attendance": "💾 Сохранить перекличку",
        "attendance_saved": "✅ Перекличка сохранена. Окно правок (15 мин) активно. Затем родителям придет уведомление.",
        "cockpit_report": "⚡ *Утренняя сводка ({date})*\n\n🏫 Всего учеников: {total_students}\n✅ Присутствуют: {present_count}\n❌ Отсутствуют: {absent_count}\n\n⚠️ *Классы без переклички ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Excel обработан. Добавлено учеников: {created_count}. Коды доступа в прикрепленном файле.",
        "excel_instructions": "📥 *Инструкция по загрузке через Excel*\n\nОтправьте файл `.xlsx` прямо в этот чат.\nПервая строка таблицы должна содержать следующие заголовки:\n\n`Ad Soyad` | `Sinif` | `Numara`\n\nПример:\n`Иван Иванов` | `9-A` | `101`\n`Анна Смирнова` | `9-A` | `102`\n\nПосле отправки бот сгенерирует коды доступа и вернет готовый файл.",
        "no_classes_found": "⚠️ *Пока не зарегистрировано ни одного класса.*\n\nВы можете загрузить Excel-файл с учениками или создать тестовый класс.",
        "btn_create_sample_class": "➕ Создать пример класса (9-A)",
        "sample_class_created": "✅ Тестовый класс 9-A (3 ученика и 1 учитель) успешно создан.",
        "select_class_for_pdf": "📄 Выберите класс для скачивания карточек с кодами доступа:",
        "pdf_cards_ready": "📄 Карточки с кодами доступа для класса *{class_name}* готовы к печати.",
        "medical_request_admin": "🏥 *Новая медицинская справка*\nУченик: *{student_name}* ({class_name})\nID родителя: `{parent_id}`",
        "btn_approve_medical": "✅ Одобрить",
        "btn_reject_medical": "❌ Отклонить",
        "class_detail_title": "🏫 *Класс {class_name}*\nВсего учеников: *{count}*",
        "btn_download_class_pdf": "📄 Скачать карточки класса (PDF)"
    },
    "uz": {
        "lang_select": "Iltimos, o'zingizga qulay tilni tanlang:",
        "lang_changed": "Til muvaffaqiyatli yangilandi: 🇺🇿 O'zbekcha",
        "welcome_guest": "🎓 *Maktab boshqaruv tizimiga xush kelibsiz.*\n\nIltimos, maktab ma'muriyati bergan maxsus kodni kiriting:",
        "auth_success": "✅ Kirish muvaffaqiyatli. Sizning rolingiz: *{role}*",
        "auth_failed": "❌ Noto'g'ri kod! Qolgan urinishlar: {remaining}",
        "auth_locked": "⛔ Xavfsizlik sababli hisobingiz 1 soatga bloklandi.",
        "auth_blacklisted": "🚫 Hisobingiz butunlay to'xtatildi.",
        "menu_parent": "👨‍👩‍👧‍👦 *Ota-ona paneli*\nO'quvchi: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *O'qituvchi paneli*\nFan: *{subject}*",
        "menu_student": "🎓 *O'quvchi paneli*\nSinf: *{class_name}* | №: *{student_number}*",
        "menu_admin": "⚡ *Maktab boshqaruv markazi (Admin)*\nXush kelibsiz.",
        "admin_stats_text": "📊 *Umumiy maktab statistikasi:*\n• Jami sinflar: *{classes_count}*\n• Jami o'quvchilar: *{students_count}*\n• Sana: *{date}*",
        "btn_admin_cockpit": "📊 Tonggi hisobot",
        "btn_admin_classes": "🏫 Sinflar va o'quvchilar",
        "btn_admin_excel": "📥 Excel orqali yuklash",
        "btn_admin_pdf": "📄 Parol kartalari (PDF)",
        "btn_admin_announcement": "📢 Ommaviy e'lon",
        "btn_admin_settings": "⚙️ Sozlamalar va til",
        "btn_parent_dashboard": "📊 Holat paneli",
        "btn_parent_switch_child": "👶 Farzandni tanlash",
        "btn_parent_medical": "🏥 Ma'lumotnoma yuborish",
        "btn_teacher_attendance": "📋 Tezkor davomat",
        "btn_teacher_grades": "📝 Baho qo'yish",
        "btn_teacher_homework": "📢 Vazifalar paneli",
        "btn_student_grades": "📊 Baholarim / Tabel",
        "btn_student_upload_hw": "📤 Vazifa topshirish",
        "btn_back": "⬅️ Orqaga",
        "btn_main_menu": "🏠 Asosiy menyu",
        "btn_acknowledged": "✅ O'qidim / Xabardorman",
        "acknowledged_toast": "Tasdiqlaganingiz qayd etildi.",
        "report_card_title": "📊 *O'quv ko'rsatkichlari va davomat*",
        "attendance_summary": "📌 Qoldirilgan darslar: *{absent} kun* (Sababli: {excused} kun)",
        "grades_list": "📝 *Fan baholari:*",
        "no_grades": "Hozircha baholar mavjud emas.",
        "upload_medical_prompt": "Iltimos, tibbiy ma'lumotnoma rasmini yuboring:",
        "medical_uploaded": "Ma'lumotnoma maktab ma'muriyatiga yuborildi.",
        "select_class_attendance": "Davomat olinadigan sinfni tanlang:",
        "attendance_started": "📋 *{class_name} sinfi davomati*\nHamma 'Keldi' deb belgilangan. Kelmagan o'quvchi ustiga bosib 'Kelmadi' qiling va saqlang.",
        "btn_save_attendance": "💾 Davomatni tasdiqlash",
        "attendance_saved": "✅ Davomat saqlandi. 15 daqiqalik tuzatish vaqti boshlandi. Vaqt tugagach ota-onalarga xabar yuboriladi.",
        "cockpit_report": "⚡ *Tonggi umumiy hisobot ({date})*\n\n🏫 Jami o'quvchilar: {total_students}\n✅ Kelganlar: {present_count}\n❌ Kelmaganlar: {absent_count}\n\n⚠️ *Davomat olinmagan sinflar ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Excel qabul qilindi. {created_count} ta o'quvchi qo'shildi. Parollar biriktirilgan faylda.",
        "excel_instructions": "📥 *Excel orqali o'quvchilarni yuklash bo'yicha qo'llanma*\n\nIltimos, `.xlsx` faylini to'g'ridan-to'g'ri ushbu chatga yuboring.\nFaylning birinchi qatorida quyidagi ustun nomlari bo'lishi shart:\n\n`Ad Soyad` | `Sinif` | `Numara`\n\nMisol:\n`Ali Valiyev` | `9-A` | `101`\n`Zuhra Karimova` | `9-A` | `102`\n\nFaylni yuborganingizdan so'ng bot barcha o'quvchi va ota-onalar uchun maxsus parollarni yaratib, tayyor faylni qaytaradi.",
        "no_classes_found": "⚠️ *Hozircha birorta sinf ro'yxatdan o'tmagan.*\n\nO'quvchilarni kiritish uchun Excel faylini yuborishingiz yoki sinov uchun namunaviy sinf yaratishingiz mumkin.",
        "btn_create_sample_class": "➕ Namunaviy sinf yaratish (9-A)",
        "sample_class_created": "✅ Sinov uchun 9-A sinfi (3 ta o'quvchi va 1 ta o'qituvchi) muvaffaqiyatli yaratildi.",
        "select_class_for_pdf": "📄 Parol kartalarini yuklab olish uchun sinfni tanlang:",
        "pdf_cards_ready": "📄 *{class_name}* sinfi uchun tarqatishga tayyor parol kartalari (PDF) biriktirildi.",
        "medical_request_admin": "🏥 *Yangi tibbiy ma'lumotnoma*\nO'quvchi: *{student_name}* ({class_name})\nOta-ona ID: `{parent_id}`",
        "btn_approve_medical": "✅ Tasdiqlash",
        "btn_reject_medical": "❌ Rad etish",
        "class_detail_title": "🏫 *{class_name} sinfi*\nJami o'quvchilar soni: *{count}*",
        "btn_download_class_pdf": "📄 Ushbu sinf parol kartalarini yuklash (PDF)"
    },
    "en": {
        "lang_select": "Please select your preferred language:",
        "lang_changed": "Language successfully updated: 🇬🇧 English",
        "welcome_guest": "🎓 *Welcome to School Management Ecosystem.*\n\nPlease enter the access code provided by administration:",
        "auth_success": "✅ Authentication successful. Your role: *{role}*",
        "auth_failed": "❌ Invalid code! Remaining attempts: {remaining}",
        "auth_locked": "⛔ Account locked for 1 hour due to multiple failed attempts.",
        "auth_blacklisted": "🚫 Your account has been permanently suspended.",
        "menu_parent": "👨‍👩‍👧‍👦 *Parent Dashboard*\nStudent: *{student_name}* ({class_name})",
        "menu_teacher": "👨‍🏫 *Teacher Dashboard*\nSubject: *{subject}*",
        "menu_student": "🎓 *Student Dashboard*\nClass: *{class_name}* | Roll: *{student_number}*",
        "menu_admin": "⚡ *School Administration Cockpit (Admin)*\nWelcome back.",
        "admin_stats_text": "📊 *School Overview:*\n• Total Classes: *{classes_count}*\n• Total Students: *{students_count}*\n• Date: *{date}*",
        "btn_admin_cockpit": "📊 Morning Cockpit",
        "btn_admin_classes": "🏫 Classes & Students",
        "btn_admin_excel": "📥 Import via Excel",
        "btn_admin_pdf": "📄 Password Cards (PDF)",
        "btn_admin_announcement": "📢 Broadcast",
        "btn_admin_settings": "⚙️ Settings & Language",
        "btn_parent_dashboard": "📊 Overview",
        "btn_parent_switch_child": "👶 Switch Child",
        "btn_parent_medical": "🏥 Submit Medical Note",
        "btn_teacher_attendance": "📋 Fast Attendance",
        "btn_teacher_grades": "📝 Grade Book",
        "btn_teacher_homework": "📢 Homework Board",
        "btn_student_grades": "📊 My Report Card",
        "btn_student_upload_hw": "📤 Submit Homework",
        "btn_back": "⬅️ Back",
        "btn_main_menu": "🏠 Main Menu",
        "btn_acknowledged": "✅ Read / Acknowledged",
        "acknowledged_toast": "Acknowledgment recorded.",
        "report_card_title": "📊 *Academic Record & Attendance*",
        "attendance_summary": "📌 Total Absences: *{absent} days* (Excused: {excused} days)",
        "grades_list": "📝 *Grade Book:*",
        "no_grades": "No grades recorded yet.",
        "upload_medical_prompt": "Please send a photo of the medical report or excuse note:",
        "medical_uploaded": "Report submitted to school administration.",
        "select_class_attendance": "Select class for attendance:",
        "attendance_started": "📋 *Attendance: Class {class_name}*\nAll students marked Present by default. Tap to toggle absent, then save.",
        "btn_save_attendance": "💾 Confirm & Save",
        "attendance_saved": "✅ Attendance recorded. 15-minute edit window started. Parents will be notified automatically thereafter.",
        "cockpit_report": "⚡ *Morning Status Brief ({date})*\n\n🏫 Total Students: {total_students}\n✅ Present: {present_count}\n❌ Absent: {absent_count}\n\n⚠️ *Pending Attendance Classes ({missing_classes_count}):*\n{missing_classes}",
        "excel_processed": "✅ Processed. Added {created_count} students. Generated codes are in the attached file.",
        "excel_instructions": "📥 *Excel Import Guide*\n\nPlease send a `.xlsx` spreadsheet directly into this chat.\nThe first row must include these column headers:\n\n`Ad Soyad` | `Sinif` | `Numara`\n\nExample:\n`John Doe` | `9-A` | `101`\n`Jane Smith` | `9-A` | `102`\n\nOnce sent, the bot will generate unique access codes and return the completed file.",
        "no_classes_found": "⚠️ *No classes registered yet.*\n\nYou can upload an Excel file with students or generate a sample class for testing.",
        "btn_create_sample_class": "➕ Create Sample Class (9-A)",
        "sample_class_created": "✅ Sample class 9-A (3 students and 1 teacher) created successfully.",
        "select_class_for_pdf": "📄 Select a class to download printable password cards:",
        "pdf_cards_ready": "📄 Printable password cards for class *{class_name}* are attached.",
        "medical_request_admin": "🏥 *New Medical Excuse Submission*\nStudent: *{student_name}* ({class_name})\\nParent ID: `{parent_id}`",
        "btn_approve_medical": "✅ Approve (Excused)",
        "btn_reject_medical": "❌ Reject",
        "class_detail_title": "🏫 *Class {class_name}*\nTotal Students: *{count}*",
        "btn_download_class_pdf": "📄 Download Class Password Cards (PDF)"
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
# 4. ARAYÜZ VE ERGONOMİK MENÜ MOTORU (SEÇİLEN DİLDE %100 UYUMLU)
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
    if role == "admin":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_admin_cockpit", lang)),
                KeyboardButton(text=get_text("btn_admin_classes", lang))
            ],
            [
                KeyboardButton(text=get_text("btn_admin_excel", lang)),
                KeyboardButton(text=get_text("btn_admin_settings", lang))
            ]
        ]
    elif role == "teacher":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_teacher_attendance", lang)),
                KeyboardButton(text=get_text("btn_teacher_grades", lang))
            ],
            [
                KeyboardButton(text=get_text("btn_teacher_homework", lang)),
                KeyboardButton(text=get_text("btn_main_menu", lang))
            ]
        ]
    elif role == "parent":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_parent_dashboard", lang)),
                KeyboardButton(text=get_text("btn_parent_switch_child", lang))
            ],
            [
                KeyboardButton(text=get_text("btn_parent_medical", lang)),
                KeyboardButton(text=get_text("btn_main_menu", lang))
            ]
        ]
    elif role == "student":
        keyboard = [
            [
                KeyboardButton(text=get_text("btn_student_grades", lang)),
                KeyboardButton(text=get_text("btn_student_upload_hw", lang))
            ],
            [KeyboardButton(text=get_text("btn_main_menu", lang))]
        ]
    else:
        keyboard = [
            [KeyboardButton(text="/start")]
        ]

    return ReplyKeyboardMarkup(keyboard=keyboard, resize_keyboard=True)

def get_admin_dashboard_inline_kb(lang: str = "tr") -> InlineKeyboardMarkup:
    buttons = [
        [
            InlineKeyboardButton(text=get_text("btn_admin_cockpit", lang), callback_data="adm:cockpit"),
            InlineKeyboardButton(text=get_text("btn_admin_classes", lang), callback_data="adm:classes")
        ],
        [
            InlineKeyboardButton(text=get_text("btn_admin_excel", lang), callback_data="adm:excel_info"),
            InlineKeyboardButton(text=get_text("btn_admin_pdf", lang), callback_data="adm:pdf_menu")
        ],
        [
            InlineKeyboardButton(text="🌐 " + get_text("btn_admin_settings", lang), callback_data="act_change_lang")
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
        InlineKeyboardButton(text=get_text("btn_save_attendance", lang), callback_data=f"att_save:{class_name}")
    ])
    inline_keyboard.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
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

    t_header = {
        "tr": f"<b>{class_name} SINIFI - TELEGRAM GİRİŞ KARTLARI</b>",
        "ru": f"<b>КЛАСС {class_name} - КАРТОЧКИ ВХОДА TELEGRAM</b>",
        "uz": f"<b>{class_name} SINFI - TELEGRAM KIRISH KARTALARI</b>",
        "en": f"<b>CLASS {class_name} - TELEGRAM ACCESS CARDS</b>"
    }.get(lang, f"<b>{class_name} - ACCESS CARDS</b>")

    lbl_st = {"tr": "Öğrenci", "ru": "Ученик", "uz": "O'quvchi", "en": "Student"}.get(lang, "Student")
    lbl_st_code = {"tr": "Öğrenci Kodu", "ru": "Код ученика", "uz": "O'quvchi kodi", "en": "Student Code"}.get(lang, "Student Code")
    lbl_pr_code = {"tr": "Veli Kodu", "ru": "Код родителя", "uz": "Ota-ona kodi", "en": "Parent Code"}.get(lang, "Parent Code")

    story = [
        Paragraph(t_header, title_style),
        Spacer(1, 15)
    ]

    card_data = []
    for s in students:
        text = (
            f"<b>{lbl_st}:</b> {s.full_name} (№: {s.student_number})<br/>"
            f"<b>{lbl_st_code}:</b> <code>{s.student_code}</code><br/>"
            f"<b>{lbl_pr_code}:</b> <code>{s.parent_code}</code>"
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

# ======================================================================
# 6. TELEGRAM ROUTER VE EKSİKSİZ ETKİLEŞİM YÖNETİCİLERİ
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
        
        total_admins = (await session.execute(select(func.count(User.telegram_id)).where(User.role == "admin"))).scalar() or 0
        auto_admin = (len(ADMIN_IDS) == 0 and total_admins == 0) or (message.from_user.id in ADMIN_IDS)

        if not user:
            await message.answer(
                "🌍 Iltimos, tilni tanlang / Пожалуйста, выберите язык / Lütfen dil seçiniz / Please select language:",
                reply_markup=get_language_inline_kb()
            )
            return

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language))
            return

        if user.locked_until and user.locked_until > datetime.utcnow():
            await message.answer(get_text("auth_locked", user.language))
            return

        if auto_admin and user.role != "admin":
            user.role = "admin"
            await session.commit()

        await render_admin_or_role_dashboard(message, user)

@router.callback_query(F.data.startswith("set_lang:"))
async def cb_set_lang(query: CallbackQuery, state: FSMContext):
    lang_code = query.data.split(":")[1]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        
        total_admins = (await session.execute(select(func.count(User.telegram_id)).where(User.role == "admin"))).scalar() or 0
        auto_admin = (len(ADMIN_IDS) == 0 and total_admins == 0) or (query.from_user.id in ADMIN_IDS)
        role = "admin" if auto_admin else "guest"

        if not user:
            user = User(telegram_id=query.from_user.id, language=lang_code, role=role)
            session.add(user)
        else:
            user.language = lang_code
            if auto_admin:
                user.role = "admin"
        await session.commit()

    try:
        await query.message.delete()
    except Exception:
        pass

    await render_admin_or_role_dashboard(query.message, user)
    await query.answer()

async def render_admin_or_role_dashboard(message: Message, user: User):
    lang = user.language
    reply_kb = get_role_reply_kb(user.role, lang)

    if user.role == "admin":
        async with AsyncSessionLocal() as session:
            classes_count = (await session.execute(select(func.count(Student.class_name.distinct())))).scalar() or 0
            students_count = (await session.execute(select(func.count(Student.id)))).scalar() or 0
            today_str = datetime.utcnow().strftime("%d.%m.%Y")

            dashboard_text = (
                f"{get_text('menu_admin', lang)}\n\n"
                f"{get_text('admin_stats_text', lang, classes_count=classes_count, students_count=students_count, date=today_str)}"
            )

            inline_kb = get_admin_dashboard_inline_kb(lang)
            await message.answer(dashboard_text, reply_markup=reply_kb, parse_mode="Markdown")
            await message.answer("👇", reply_markup=inline_kb, parse_mode="Markdown")

    elif user.role == "teacher":
        text = get_text("menu_teacher", lang, subject="Öğretmen")
        inline_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text=get_text("btn_teacher_attendance", lang), callback_data="tch:select_class")],
            [InlineKeyboardButton(text=get_text("btn_admin_settings", lang), callback_data="act_change_lang")]
        ])
        await message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
        await message.answer("👇", reply_markup=inline_kb, parse_mode="Markdown")

    elif user.role == "parent":
        async with AsyncSessionLocal() as session:
            st = await session.get(Student, user.current_child_id if user.current_child_id else 1)
            st_name = st.full_name if st else "-"
            cls_name = st.class_name if st else "-"
            text = get_text("menu_parent", lang, student_name=st_name, class_name=cls_name)
            inline_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_parent_dashboard", lang), callback_data="act_view_report")],
                [InlineKeyboardButton(text=get_text("btn_parent_medical", lang), callback_data="upload_medical_init")],
                [InlineKeyboardButton(text=get_text("btn_admin_settings", lang), callback_data="act_change_lang")]
            ])
            await message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
            await message.answer("👇", reply_markup=inline_kb, parse_mode="Markdown")

    elif user.role == "student":
        async with AsyncSessionLocal() as session:
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user.telegram_id))).scalar_one_or_none()
            cls_name = st.class_name if st else "-"
            num_val = st.student_number if st else "-"
            text = get_text("menu_student", lang, class_name=cls_name, student_number=num_val)
            inline_kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_student_grades", lang), callback_data="act_view_report")],
                [InlineKeyboardButton(text=get_text("btn_admin_settings", lang), callback_data="act_change_lang")]
            ])
            await message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
            await message.answer("👇", reply_markup=inline_kb, parse_mode="Markdown")

    else:
        text = get_text("welcome_guest", lang)
        inline_kb = InlineKeyboardMarkup(inline_keyboard=[
            [InlineKeyboardButton(text="🔑 Giriş Kodu Yaz", callback_data="act_enter_code")],
            [InlineKeyboardButton(text="🌐 Dili Değiştir", callback_data="act_change_lang")]
        ])
        await message.answer(text, reply_markup=reply_kb, parse_mode="Markdown")
        await message.answer("👇", reply_markup=inline_kb, parse_mode="Markdown")

@router.callback_query(F.data == "adm:dashboard")
async def cb_admin_dashboard(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        if user:
            try:
                await query.message.delete()
            except Exception:
                pass
            await render_admin_or_role_dashboard(query.message, user)
    await query.answer()

@router.callback_query(F.data == "adm:cockpit")
@router.message(F.text.in_([LOCALES[l]["btn_admin_cockpit"] for l in LOCALES]))
async def admin_morning_cockpit(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
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

        text = get_text(
            "cockpit_report", lang,
            date=today.strftime("%d.%m.%Y"),
            total_students=total_students,
            present_count=present_count,
            absent_count=absent_count,
            missing_classes_count=len(missing),
            missing_classes=missing_str
        )
    msg = event.message if isinstance(event, CallbackQuery) else event
    kb = InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)])
    await msg.answer(text, reply_markup=kb, parse_mode="Markdown")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "adm:classes")
@router.message(F.text.in_([LOCALES[l]["btn_admin_classes"] for l in LOCALES]))
async def admin_classes_list(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        msg = event.message if isinstance(event, CallbackQuery) else event

        if not classes:
            text = get_text("no_classes_found", lang)
            kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_create_sample_class", lang), callback_data="adm:create_sample")],
                [InlineKeyboardButton(text=get_text("btn_admin_excel", lang), callback_data="adm:excel_info")],
                get_nav_buttons(lang)
            ])
            await msg.answer(text, reply_markup=kb, parse_mode="Markdown")
            if isinstance(event, CallbackQuery):
                await event.answer()
            return

        buttons = []
        row = []
        for c in classes:
            row.append(InlineKeyboardButton(text=f"🏫 {c}", callback_data=f"adm:class_detail:{c}"))
            if len(row) == 2:
                buttons.append(row)
                row = []
        if row:
            buttons.append(row)

        buttons.append(get_nav_buttons(lang))
        text = f"🏫 *{get_text('btn_admin_classes', lang)}*\n{get_text('select_class_attendance', lang)}"
        await msg.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
        if isinstance(event, CallbackQuery):
            await event.answer()

@router.callback_query(F.data == "adm:create_sample")
async def cb_create_sample_class(query: CallbackQuery):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        existing = (await session.execute(select(Student).where(Student.class_name == "9-A"))).scalars().all()
        if not existing:
            s1 = Student(full_name="Ali Yılmaz", student_number="101", class_name="9-A", student_code="OGR-1001", parent_code="VELI-1001")
            s2 = Student(full_name="Ayşe Kaya", student_number="102", class_name="9-A", student_code="OGR-1002", parent_code="VELI-1002")
            s3 = Student(full_name="Mehmet Demir", student_number="103", class_name="9-A", student_code="OGR-1003", parent_code="VELI-1003")
            t1 = Teacher(full_name="Ahmet Hoca", subject="Matematik", auth_code="HCA-1001")
            session.add_all([s1, s2, s3, t1])
            await session.commit()

        await query.message.answer(get_text("sample_class_created", lang))
        await admin_classes_list(query)
    await query.answer()

@router.callback_query(F.data.startswith("adm:class_detail:"))
async def cb_class_detail(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        students = (await session.execute(select(Student).where(Student.class_name == class_name).order_by(Student.student_number))).scalars().all()
        text = f"{get_text('class_detail_title', lang, class_name=class_name, count=len(students))}\n\n"
        for s in students:
            text += f"• *{s.full_name}* (№: `{s.student_number}`)\n   Code: `{s.student_code}` | Parent: `{s.parent_code}`\n"

        buttons = [
            [InlineKeyboardButton(text=get_text("btn_download_class_pdf", lang), callback_data=f"adm:gen_pdf:{class_name}")],
            [InlineKeyboardButton(text=get_text("btn_teacher_attendance", lang), callback_data=f"att_class:{class_name}")],
            get_nav_buttons(lang, back_callback="adm:classes")
        ]
        await query.message.edit_text(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
    await query.answer()

@router.callback_query(F.data == "adm:excel_info")
@router.message(F.text.in_([LOCALES[l]["btn_admin_excel"] for l in LOCALES]))
async def cb_excel_info(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

    msg = event.message if isinstance(event, CallbackQuery) else event
    kb = InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)])
    await msg.answer(get_text("excel_instructions", lang), reply_markup=kb, parse_mode="Markdown")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "adm:pdf_menu")
@router.message(F.text.in_([LOCALES[l]["btn_admin_pdf"] for l in LOCALES]))
async def cb_pdf_menu(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        classes = (await session.execute(select(Student.class_name).distinct().order_by(Student.class_name))).scalars().all()
        msg = event.message if isinstance(event, CallbackQuery) else event

        if not classes:
            text = get_text("no_classes_found", lang)
            kb = InlineKeyboardMarkup(inline_keyboard=[
                [InlineKeyboardButton(text=get_text("btn_create_sample_class", lang), callback_data="adm:create_sample")],
                get_nav_buttons(lang)
            ])
            await msg.answer(text, reply_markup=kb, parse_mode="Markdown")
            if isinstance(event, CallbackQuery):
                await event.answer()
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
        await msg.answer(get_text("select_class_for_pdf", lang), reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons), parse_mode="Markdown")
        if isinstance(event, CallbackQuery):
            await event.answer()

@router.callback_query(F.data.startswith("adm:gen_pdf:"))
async def cb_generate_pdf(query: CallbackQuery):
    class_name = query.data.split(":")[2]
    async with AsyncSessionLocal() as session:
        user = await session.get(User, query.from_user.id)
        lang = user.language if user else "tr"

        pdf_buffer = await generate_classroom_pdf_cards(class_name, lang=lang)
        file = BufferedInputFile(pdf_buffer.read(), filename=f"{class_name}_Sifre_Kartlari.pdf")
        caption = get_text("pdf_cards_ready", lang, class_name=class_name)
        await query.message.answer_document(file, caption=caption, parse_mode="Markdown")
        pdf_buffer.close()
    await query.answer()

@router.message(Command("kodlar"))
async def admin_pdf_cards_command(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"

    args = message.text.split(maxsplit=1)
    class_name = args[1].strip() if len(args) > 1 else "9-A"
    pdf_buffer = await generate_classroom_pdf_cards(class_name, lang=lang)
    file = BufferedInputFile(pdf_buffer.read(), filename=f"{class_name}_Sifre_Kartlari.pdf")
    await message.answer_document(file, caption=get_text("pdf_cards_ready", lang, class_name=class_name), parse_mode="Markdown")
    pdf_buffer.close()

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
    await message.answer_document(file, caption=get_text("excel_processed", lang, created_count=count))
    out_excel.close()

@router.callback_query(F.data == "act_change_lang")
@router.message(F.text.in_([LOCALES[l]["btn_admin_settings"] for l in LOCALES]))
async def cb_change_lang_screen(event: Message | CallbackQuery):
    msg = event.message if isinstance(event, CallbackQuery) else event
    await msg.answer(
        "🌍 Lütfen bir dil seçiniz / Пожалуйста, выберите язык / Tilni tanlang / Select a language:",
        reply_markup=get_language_inline_kb()
    )
    if isinstance(event, CallbackQuery):
        await event.answer()

# -------------------------------------------------------------
# ÖĞRETMEN MODÜLÜ (YOKLAMA GİRİŞİ)
# -------------------------------------------------------------

@router.message(F.text.in_([LOCALES[l]["btn_teacher_attendance"] for l in LOCALES]))
async def teacher_start_attendance(message: Message):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
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
        buttons.append(get_nav_buttons(lang, back_callback="adm:dashboard"))
        await message.answer(
            get_text("select_class_attendance", lang),
            reply_markup=InlineKeyboardMarkup(inline_keyboard=buttons)
        )

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
        await query.message.edit_text(
            get_text("attendance_started", lang, class_name=class_name),
            reply_markup=kb,
            parse_mode="Markdown"
        )
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

# -------------------------------------------------------------
# VELİ VE ÖĞRENCİ MODÜLÜ
# -------------------------------------------------------------

@router.callback_query(F.data == "act_view_report")
@router.message(F.text.in_([LOCALES[l]["btn_parent_dashboard"] for l in LOCALES] + [LOCALES[l]["btn_student_grades"] for l in LOCALES]))
async def view_report_card(event: Message | CallbackQuery):
    user_id = event.from_user.id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, user_id)
        lang = user.language if user else "tr"

        st_id = user.current_child_id if user and user.role == "parent" else None
        if not st_id:
            st = (await session.execute(select(Student).where(Student.student_telegram_id == user_id))).scalar_one_or_none()
            if not st:
                st = (await session.execute(select(Student).limit(1))).scalar_one_or_none()
            st_id = st.id if st else None

        msg = event.message if isinstance(event, CallbackQuery) else event
        if not st_id:
            await msg.answer(get_text("no_grades", lang), parse_mode="Markdown")
            if isinstance(event, CallbackQuery):
                await event.answer()
            return

        st = await session.get(Student, st_id)
        att_stmt = select(Attendance).where(Attendance.student_id == st.id)
        attendances = (await session.execute(att_stmt)).scalars().all()
        absent_count = sum(1 for a in attendances if a.status == "absent")
        excused_count = sum(1 for a in attendances if a.status == "excused")

        gr_stmt = select(Grade).where(Grade.student_id == st.id).order_by(Grade.created_at.desc())
        grades = (await session.execute(gr_stmt)).scalars().all()

        grades_text = "\n".join([f"• {g.subject}: *{g.score}* ({g.badge}) - _{g.note}_" for g in grades]) if grades else get_text("no_grades", lang)

        text = (
            f"{get_text('report_card_title', lang)}\n"
            f"👤 Öğrenci: *{st.full_name}* ({st.class_name})\n\n"
            f"{get_text('attendance_summary', lang, absent=absent_count, excused=excused_count)}\n\n"
            f"{get_text('grades_list', lang)}\n{grades_text}"
        )

    await msg.answer(text, reply_markup=InlineKeyboardMarkup(inline_keyboard=[get_nav_buttons(lang)]), parse_mode="Markdown")
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.callback_query(F.data == "upload_medical_init")
@router.message(F.text.in_([LOCALES[l]["btn_parent_medical"] for l in LOCALES]))
async def cb_upload_medical_init(event: Message | CallbackQuery, state: FSMContext):
    async with AsyncSessionLocal() as session:
        user = await session.get(User, event.from_user.id)
        lang = user.language if user else "tr"
    msg = event.message if isinstance(event, CallbackQuery) else event
    await msg.answer(get_text("upload_medical_prompt", lang))
    await state.set_state(Form.waiting_medical_photo)
    if isinstance(event, CallbackQuery):
        await event.answer()

@router.message(Form.waiting_medical_photo, F.photo)
async def handle_medical_photo(message: Message, state: FSMContext):
    photo_file_id = message.photo[-1].file_id
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        lang = user.language if user else "tr"
        st = await session.get(Student, user.current_child_id if user and user.current_child_id else 1)
        st_name = st.full_name if st else "Öğrenci"
        st_class = st.class_name if st else "Genel"

        report = MedicalReport(
            student_id=st.id if st else 1,
            parent_telegram_id=user.telegram_id,
            file_id=photo_file_id,
            caption=message.caption or "Rapor görseli"
        )
        session.add(report)
        await session.commit()

        for adm_id in ADMIN_IDS or [message.from_user.id]:
            try:
                adm_kb = InlineKeyboardMarkup(inline_keyboard=[
                    [
                        InlineKeyboardButton(text=get_text("btn_approve_medical", "tr"), callback_data=f"med_appr:{report.id}"),
                        InlineKeyboardButton(text=get_text("btn_reject_medical", "tr"), callback_data=f"med_rej:{report.id}")
                    ]
                ])
                await message.bot.send_photo(
                    chat_id=adm_id,
                    photo=photo_file_id,
                    caption=get_text("medical_request_admin", "tr", student_name=st_name, class_name=st_class, parent_id=user.telegram_id),
                    reply_markup=adm_kb,
                    parse_mode="Markdown"
                )
            except Exception:
                pass

        await message.answer(get_text("medical_uploaded", lang))
    await state.clear()

@router.callback_query(F.data.startswith("ack_notif:"))
async def cb_acknowledge_notification(query: CallbackQuery):
    notif_id = int(query.data.split(":")[1])
    async with AsyncSessionLocal() as session:
        notif = await session.get(CriticalNotification, notif_id)
        if notif:
            notif.acknowledged_at = datetime.utcnow()
            await session.commit()
    await query.answer(get_text("acknowledged_toast", "tr"), show_alert=True)
    await query.message.edit_reply_markup(reply_markup=None)

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
    print("=" * 60)
    print("--> [1/4] Veritabanı başlatılıyor...")
    await init_db()
    print("--> [1/4] Veritabanı tabloları hazır.")

    bot_info = None
    try:
        bot_info = await bot.get_me()
        print(f"--> [2/4] Telegram Bot Bilgisi: @{bot_info.username} (ID: {bot_info.id})")
    except Exception as e:
        print(f"--> [HATA 2/4] BOT_TOKEN ile Telegram'a bağlanılamadı: {e}")

    if bot_info and WEBHOOK_URL:
        try:
            print(f"--> [3/4] Webhook Telegram'a kaydediliyor: {WEBHOOK_URL}")
            await bot.delete_webhook(drop_pending_updates=False)
            await bot.set_webhook(
                url=WEBHOOK_URL,
                secret_token=WEBHOOK_SECRET,
                drop_pending_updates=False,
                allowed_updates=["message", "callback_query"]
            )
            wh = await bot.get_webhook_info()
            print(f"--> [3/4] Webhook Başarıyla Kuruldu! Aktif URL: {wh.url}")
            if wh.last_error_message:
                print(f"--> [UYARI] Telegram Son Hata: {wh.last_error_message}")
        except Exception as e:
            print(f"--> [HATA 3/4] Webhook kurulum hatası: {e}")
    else:
        print(f"--> [UYARI 3/4] Webhook kurulamadı! WEBHOOK_URL='{WEBHOOK_URL}'")

    worker_task = asyncio.create_task(background_attendance_loop())
    print("--> [4/4] Gecikmeli bildirim kuyruk işçisi aktif.")
    print("--> SISTEM CANLI VE TELEGRAM MESAJLARINI BEKLIYOR.")
    print("=" * 60)
    
    yield
    
    worker_task.cancel()
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
        print(f"--> [GÜVENLİK ENGELİ] Geçersiz Webhook Secret. Gelen: {secret}")
        return JSONResponse(status_code=status.HTTP_403_FORBIDDEN, content={"error": "Invalid secret"})

    try:
        data = await request.json()
    except Exception as e:
        print(f"--> [HATA] JSON okunamadı: {e}")
        return JSONResponse(status_code=status.HTTP_400_BAD_REQUEST, content={"error": "Bad JSON"})

    update_id = data.get("update_id", 0)
    print(f"--> [MESAJ GELDİ] Telegram Update ID: {update_id}")

    try:
        telegram_update = Update.model_validate(data, context={"bot": bot})
        background_tasks.add_task(dp.feed_update, bot, telegram_update)
    except Exception as e:
        print(f"--> [HATA] aiogram Update nesnesine çevrilemedi: {e}")

    return {"ok": True}

if __name__ == "__main__":
    import uvicorn
    port = int(os.getenv("PORT", 8000))
    uvicorn.run("app:app", host="0.0.0.0", port=port, log_level="info")
