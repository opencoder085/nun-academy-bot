# -*- coding: utf-8 -*-
"""
aiogram 3.x Telegram Router ve Olay Yöneticileri
"""

import os
from datetime import datetime, timedelta
from aiogram import Router, F, Bot
from aiogram.filters import CommandStart, Command
from aiogram.types import Message, CallbackQuery, BufferedInputFile
from aiogram.fsm.context import FSMContext
from aiogram.fsm.state import State, StatesGroup
from sqlalchemy import select, func

from database import (
    AsyncSessionLocal, User, Student, Teacher, ParentStudent,
    Attendance, Grade, MedicalReport, CriticalNotification, SystemSetting
)
from locales import get_text
from keyboards import (
    get_language_inline_kb, get_role_reply_kb, get_nav_buttons,
    get_attendance_grid_kb
)
from services import process_student_excel, generate_classroom_pdf_cards

router = Router()
ADMIN_IDS = [int(x.strip()) for x in os.getenv("ADMIN_IDS", "").split(",") if x.strip()]

class Form(StatesGroup):
    waiting_auth_code = State()
    waiting_medical_photo = State()

@router.message(CommandStart())
async def cmd_start(message: Message, state: FSMContext):
    await state.clear()
    async with AsyncSessionLocal() as session:
        user = await session.get(User, message.from_user.id)
        if not user:
            await message.answer(
                get_text("lang_select", "tr"),
                reply_markup=get_language_inline_kb()
            )
            return

        if user.is_blacklisted:
            await message.answer(get_text("auth_blacklisted", user.language))
            return

        if user.locked_until and user.locked_until > datetime.utcnow():
            await message.answer(get_text("auth_locked", user.language))
            return

        if user.role == "guest":
            await message.answer(
                get_text("welcome_guest", user.language),
                parse_mode="Markdown"
            )
            await state.set_state(Form.waiting_auth_code)
            return

        await send_role_home(message, user)

@router.callback_query(F.data.startswith("set_lang:"))
async def cb_set_lang(query: CallbackQuery, state: FSMContext):
    lang_code = query.data.split(":")[1]
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
        await query.message.answer(
            get_text("welcome_guest", lang_code),
            parse_mode="Markdown"
        )
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
    class_name = args[1].strip()
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
    await message.answer_document(
        file,
        caption=get_text("excel_processed", "tr", created_count=count)
    )
    out_excel.close()

ATTENDANCE_CACHE = {}

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
            status = "absent" if is_absent else "present"
            att = Attendance(
                student_id=st_id,
                class_name=class_name,
                date=now.date(),
                status=status,
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
    notif_id = int(query.data.split(":")[1])
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
