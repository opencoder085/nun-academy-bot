# -*- coding: utf-8 -*-
"""
Minimalist ve Ergonomik Arayüz Motoru (3 Tık Kuralı, Izgara Düzeni)
"""

from aiogram.types import (
    ReplyKeyboardMarkup, KeyboardButton,
    InlineKeyboardMarkup, InlineKeyboardButton
)
from locales import get_text

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
