import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox, scrolledtext
from datetime import datetime
import threading
import webbrowser
import socket
import io
from pathlib import Path

from config import config, _check_required
from logger import setup_logging, get_logger
from backup import make_backup, list_backups, start_backup_scheduler
from database import db, PARA_TIMES, DAYS_OF_WEEK, ATTENDANCE_STATUSES, GRADE_TYPES, ROLES, SEMESTERS

_check_required()

setup_logging()
logger = get_logger(__name__)

try:
    from flask import Flask, render_template, request, jsonify, send_file, redirect
    FLASK_AVAILABLE = True
except ImportError:
    FLASK_AVAILABLE = False

from openpyxl.styles import Font

try:
    from waitress import serve
    WAITRESS_AVAILABLE = True
except ImportError:
    WAITRESS_AVAILABLE = False

try:
    from flask_limiter import Limiter
    from flask_limiter.util import get_remote_address
    LIMITER_AVAILABLE = True
except ImportError:
    LIMITER_AVAILABLE = False
    logger.warning("flask-limiter не установлен")

os.makedirs("templates", exist_ok=True)
os.makedirs("static", exist_ok=True)

limiter = None
if LIMITER_AVAILABLE:
    limiter = Limiter(
        key_func=get_remote_address,
        storage_uri="memory://",
        default_limits=[],
    )


class ScheduleApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Управление расписанием")
        self.root.geometry("900x700")

        self.server_thread = None
        self.web_server_running = False
        self.flask_app = None

        self.create_flask_app()
        self.create_widgets()

    def create_flask_app(self):
        if not FLASK_AVAILABLE:
            logger.error("Flask не установлен")
            return

        app = Flask(__name__)
        app.secret_key = config.SECRET_KEY

        if LIMITER_AVAILABLE and limiter:
            limiter.init_app(app)
        else:
            logger.warning("Rate limiting отключён")

        def get_current_user():
            token = request.cookies.get('token')
            if token:
                return db.get_user_by_token(token)
            return None

        def _excel_header_styles():
            from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
            return {
                'header_fill': PatternFill(start_color='1A3A5C', end_color='1A3A5C', fill_type='solid'),
                'header_font': Font(color='FFFFFF', bold=True, size=11),
                'thin_border': Border(
                    left=Side(style='thin', color='CCCCCC'),
                    right=Side(style='thin', color='CCCCCC'),
                    top=Side(style='thin', color='CCCCCC'),
                    bottom=Side(style='thin', color='CCCCCC')
                ),
                'center': Alignment(horizontal='center', vertical='center'),
                'left': Alignment(horizontal='left', vertical='center', wrap_text=True),
            }

        def _period_str(date_from, date_to):
            if date_from and date_to:
                return f'с {date_from} по {date_to}'
            if date_from:
                return f'с {date_from}'
            if date_to:
                return f'по {date_to}'
            return 'за всё время'

        def _build_student_diary_excel(student, group_name, records, date_from, date_to):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame().to_excel(writer, index=False, sheet_name='Дневник')
                ws = writer.sheets['Дневник']
                styles = _excel_header_styles()

                ws['A1'] = 'Дневник ученика'
                ws['A1'].font = styles['header_font']
                ws['A1'].fill = styles['header_fill']
                ws.merge_cells('A1:J1')

                ws['A2'] = f'{student["full_name"]} — группа {group_name}'
                ws.merge_cells('A2:J2')

                ws['A3'] = f'Период: {_period_str(date_from, date_to)}'
                ws.merge_cells('A3:J3')

                headers = ['Дата', 'День', 'Пара', 'Предмет', 'Преподаватель',
                           'Кабинет', 'Оценка', 'Тип', 'Посещение', 'Комментарий']
                for col, h in enumerate(headers, 1):
                    cell = ws.cell(row=5, column=col, value=h)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                att_labels = {'present': 'Был', 'late': 'Опоздал', 'absent': 'Н'}
                grade_types = {
                    'regular': 'Текущая', 'control': 'Контрольная',
                    'test': 'Самостоятельная', 'oral': 'Устный ответ',
                    'homework': 'Домашняя работа',
                }

                for i, r in enumerate(records, start=6):
                    comment_parts = []
                    if r.get('grade_comment'):
                        comment_parts.append(r['grade_comment'])
                    if r.get('attendance_comment'):
                        comment_parts.append(r['attendance_comment'])
                    row_vals = [
                        r.get('date') or '',
                        r.get('day_of_week') or '',
                        r.get('para_number') or '',
                        r.get('subject_name') or '',
                        r.get('teacher_name') or '',
                        r.get('room_number') or '',
                        r.get('grade') if r.get('grade') is not None else '',
                        grade_types.get(r.get('grade_type'), r.get('grade_type') or ''),
                        att_labels.get(r.get('attendance_status'), ''),
                        ' / '.join(comment_parts),
                    ]
                    for col, v in enumerate(row_vals, 1):
                        cell = ws.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['center'] if col in (1, 2, 3, 7, 9) else styles['left']

                widths = {'A': 12, 'B': 12, 'C': 6, 'D': 32, 'E': 28, 'F': 10, 'G': 8, 'H': 16, 'I': 12, 'J': 30}
                for col, w in widths.items():
                    ws.column_dimensions[col].width = w

                ws.freeze_panes = 'A6'

            return output

        def _build_group_diary_excel(group_name, data_group, date_from, date_to):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame().to_excel(writer, index=False, sheet_name='Сводка')
                summary = writer.sheets['Сводка']
                styles = _excel_header_styles()

                summary['A1'] = f'Дневник группы {group_name}'
                summary['A1'].font = styles['header_font']
                summary['A1'].fill = styles['header_fill']
                summary.merge_cells('A1:F1')

                summary['A2'] = f'Период: {_period_str(date_from, date_to)}'
                summary.merge_cells('A2:F2')

                headers = ['№', 'ФИО', 'Оценок', 'Средний балл', 'Пропусков', 'Посещаемость, %']
                for col, h in enumerate(headers, 1):
                    cell = summary.cell(row=4, column=col, value=h)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                students = data_group.get('students', [])
                records_by_student = data_group.get('records', {})

                for i, st in enumerate(students, start=5):
                    recs = records_by_student.get(st['id'], [])
                    grades = [r['grade'] for r in recs if r.get('grade') is not None]
                    avg = round(sum(grades) / len(grades), 2) if grades else 0
                    absents = sum(1 for r in recs if r.get('attendance_status') == 'absent')
                    total_att = sum(1 for r in recs if r.get('attendance_status') in ('present', 'late', 'absent'))
                    presents = sum(1 for r in recs if r.get('attendance_status') in ('present', 'late'))
                    pct = round(presents / total_att * 100, 1) if total_att else 0

                    row_vals = [i - 4, st['full_name'], len(grades), avg, absents, pct]
                    for col, v in enumerate(row_vals, 1):
                        cell = summary.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['center'] if col != 2 else styles['left']

                widths = {'A': 6, 'B': 40, 'C': 12, 'D': 14, 'E': 12, 'F': 16}
                for col, w in widths.items():
                    summary.column_dimensions[col].width = w
                summary.freeze_panes = 'A5'

                att_labels = {'present': 'Был', 'late': 'Опоздал', 'absent': 'Н'}
                grade_types = {
                    'regular': 'Текущая', 'control': 'Контрольная',
                    'test': 'Самостоятельная', 'oral': 'Устный ответ',
                    'homework': 'Домашняя работа',
                }

                for st in students:
                    recs = records_by_student.get(st['id'], [])
                    safe_name = (st['full_name'] or f'id{st["id"]}')[:31]
                    for ch in [':', '\\', '/', '?', '*', '[', ']']:
                        safe_name = safe_name.replace(ch, '_')
                    if safe_name in writer.sheets:
                        safe_name = f"{safe_name[:28]}_{st['id']}"

                    pd.DataFrame().to_excel(writer, index=False, sheet_name=safe_name)
                    ws = writer.sheets[safe_name]

                    ws['A1'] = st['full_name']
                    ws['A1'].font = styles['header_font']
                    ws['A1'].fill = styles['header_fill']
                    ws.merge_cells('A1:J1')

                    ws['A2'] = f'Период: {_period_str(date_from, date_to)}'
                    ws.merge_cells('A2:J2')

                    headers2 = ['Дата', 'День', 'Пара', 'Предмет', 'Преподаватель',
                                'Кабинет', 'Оценка', 'Тип', 'Посещение', 'Комментарий']
                    for col, h in enumerate(headers2, 1):
                        cell = ws.cell(row=4, column=col, value=h)
                        cell.fill = styles['header_fill']
                        cell.font = styles['header_font']
                        cell.alignment = styles['center']
                        cell.border = styles['thin_border']

                    for j, r in enumerate(recs, start=5):
                        comment_parts = []
                        if r.get('grade_comment'):
                            comment_parts.append(r['grade_comment'])
                        if r.get('attendance_comment'):
                            comment_parts.append(r['attendance_comment'])
                        row_vals = [
                            r.get('date') or '',
                            r.get('day_of_week') or '',
                            r.get('para_number') or '',
                            r.get('subject_name') or '',
                            r.get('teacher_name') or '',
                            r.get('room_number') or '',
                            r.get('grade') if r.get('grade') is not None else '',
                            grade_types.get(r.get('grade_type'), r.get('grade_type') or ''),
                            att_labels.get(r.get('attendance_status'), ''),
                            ' / '.join(comment_parts),
                        ]
                        for col, v in enumerate(row_vals, 1):
                            cell = ws.cell(row=j, column=col, value=v)
                            cell.border = styles['thin_border']
                            cell.alignment = styles['center'] if col in (1, 2, 3, 7, 9) else styles['left']

                    widths2 = {'A': 12, 'B': 12, 'C': 6, 'D': 32, 'E': 28, 'F': 10, 'G': 8, 'H': 16, 'I': 12, 'J': 30}
                    for col, w in widths2.items():
                        ws.column_dimensions[col].width = w
                    ws.freeze_panes = 'A5'

            return output

        def _build_curriculum_excel(rows, group_name=None, semester=None, academic_year=None):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                pd.DataFrame().to_excel(writer, index=False, sheet_name='Учебный план')
                ws = writer.sheets['Учебный план']
                styles = _excel_header_styles()

                ws['A1'] = 'Учебный план'
                ws['A1'].font = styles['header_font']
                ws['A1'].fill = styles['header_fill']
                ws.merge_cells('A1:L1')

                sub_parts = []
                if group_name:
                    sub_parts.append(f'Группа: {group_name}')
                if semester:
                    sub_parts.append(f'Семестр: {semester}')
                if academic_year:
                    sub_parts.append(f'Учебный год: {academic_year}')
                ws['A2'] = ' • '.join(sub_parts) if sub_parts else 'Все группы'
                ws.merge_cells('A2:L2')

                headers = ['Группа', 'Предмет', 'Преподаватель',
                           'План, ч', 'Пар в неделю', 'Факт, ч', 'Отменено, ч',
                           'Остаток, ч', '%', 'Семестр', 'Учебный год', 'Приоритет']
                for col, h in enumerate(headers, 1):
                    cell = ws.cell(row=4, column=col, value=h)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                priority_labels = {1: 'Высокий', 2: 'Средний', 3: 'Низкий'}
                for i, r in enumerate(rows, start=5):
                    row_vals = [
                        r.get('group_name') or '',
                        r.get('subject_name') or '',
                        r.get('teacher_short') or r.get('teacher_name') or '',
                        r.get('planned_hours') or 0,
                        r.get('pairs_per_week') or 0,
                        r.get('fact_hours') or 0,
                        r.get('cancelled_hours') or 0,
                        r.get('remaining_hours') or 0,
                        r.get('percent') or 0,
                        r.get('semester') or '',
                        r.get('academic_year') or '',
                        priority_labels.get(r.get('priority'), r.get('priority') or ''),
                    ]
                    for col, v in enumerate(row_vals, 1):
                        cell = ws.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['center'] if col not in (2, 3) else styles['left']

                widths = {'A': 12, 'B': 32, 'C': 24, 'D': 10, 'E': 14, 'F': 10,
                          'G': 12, 'H': 12, 'I': 8, 'J': 12, 'K': 14, 'L': 12}
                for col, w in widths.items():
                    ws.column_dimensions[col].width = w
                ws.freeze_panes = 'A5'

            return output

        def _build_archive_excel(snap, user):
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                styles = _excel_header_styles()

                grades = snap.get('grades') or []
                attendance = snap.get('attendance') or []
                schedule_data = snap.get('schedule') or {}
                schedule_base = schedule_data.get('base') or []
                schedule_overrides = schedule_data.get('overrides') or []
                curriculum = snap.get('curriculum') or []

                if user['role'] == 'student':
                    grades = [g for g in grades if g.get('student_name') == user.get('full_name')]
                    attendance = [a for a in attendance if a.get('student_name') == user.get('full_name')]
                elif user['role'] == 'parent':
                    children = db.get_parent_children_detailed(user.get('parent_of', ''))
                    names = {c['full_name'] for c in children}
                    if names:
                        grades = [g for g in grades if g.get('student_name') in names]
                        attendance = [a for a in attendance if a.get('student_name') in names]
                elif user['role'] == 'teacher':
                    teacher = db.get_teacher_by_user(user['id'])
                    if teacher:
                        tname = teacher['full_name']
                        grades = [g for g in grades if g.get('teacher_name') == tname]
                        attendance = [a for a in attendance if a.get('teacher_name') == tname]
                        schedule_base = [s for s in schedule_base if s.get('teacher_name') == tname]
                        schedule_overrides = [o for o in schedule_overrides if o.get('new_teacher_name') == tname]

                pd.DataFrame().to_excel(writer, index=False, sheet_name='Сводка')
                ws = writer.sheets['Сводка']
                ws['A1'] = snap.get('name') or 'Архив'
                ws['A1'].font = styles['header_font']
                ws['A1'].fill = styles['header_fill']
                ws.merge_cells('A1:B1')

                ws['A2'] = f"Семестр: {snap.get('semester')} {snap.get('academic_year')}"
                ws.merge_cells('A2:B2')

                ws['A3'] = f"Период: с {snap.get('date_from')} по {snap.get('date_to')}"
                ws.merge_cells('A3:B3')

                ws['A4'] = f"Создан: {snap.get('created_at')} ({snap.get('created_by_name') or '—'})"
                ws.merge_cells('A4:B4')

                ws['A6'] = 'Показатель'
                ws['B6'] = 'Значение'
                for c in ('A6', 'B6'):
                    ws[c].fill = styles['header_fill']
                    ws[c].font = styles['header_font']
                    ws[c].border = styles['thin_border']

                stats = snap.get('stats') or {}
                stats_rows = [
                    ('Учеников', stats.get('total_students', 0)),
                    ('Преподавателей', stats.get('total_teachers', 0)),
                    ('Оценок', len(grades)),
                    ('Отметок посещения', len(attendance)),
                    ('Пар в расписании', len(schedule_base)),
                    ('Замен/отмен', len(schedule_overrides)),
                    ('Средний балл', stats.get('avg_grade', 0)),
                ]
                for i, (label, val) in enumerate(stats_rows, start=7):
                    ws.cell(row=i, column=1, value=label).border = styles['thin_border']
                    ws.cell(row=i, column=2, value=val).border = styles['thin_border']

                ws.column_dimensions['A'].width = 30
                ws.column_dimensions['B'].width = 25

                pd.DataFrame().to_excel(writer, index=False, sheet_name='Оценки')
                ws2 = writer.sheets['Оценки']
                headers = ['Дата', 'Ученик', 'Группа', 'Предмет', 'Преподаватель', 'Оценка', 'Тип', 'Комментарий']
                for col, h in enumerate(headers, 1):
                    cell = ws2.cell(row=1, column=col, value=h)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                for i, g in enumerate(grades, start=2):
                    row_vals = [
                        g.get('date') or '',
                        g.get('student_name') or '',
                        g.get('student_group') or '',
                        g.get('subject_name') or '',
                        g.get('teacher_name') or '',
                        g.get('grade') if g.get('grade') is not None else '',
                        g.get('grade_type') or '',
                        g.get('comment') or '',
                    ]
                    for col, v in enumerate(row_vals, 1):
                        cell = ws2.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['center'] if col in (1, 6) else styles['left']

                for col, w in {'A': 12, 'B': 28, 'C': 10, 'D': 28, 'E': 24, 'F': 8, 'G': 16, 'H': 30}.items():
                    ws2.column_dimensions[col].width = w
                ws2.freeze_panes = 'A2'

                pd.DataFrame().to_excel(writer, index=False, sheet_name='Посещение')
                ws3 = writer.sheets['Посещение']
                headers3 = ['Дата', 'Ученик', 'Группа', 'Предмет', 'Преподаватель', 'Статус', 'Комментарий']
                for col, h in enumerate(headers3, 1):
                    cell = ws3.cell(row=1, column=col, value=h)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                att_labels = {'present': 'Был', 'late': 'Опоздал', 'absent': 'Н'}
                for i, a in enumerate(attendance, start=2):
                    row_vals = [
                        a.get('date') or '',
                        a.get('student_name') or '',
                        a.get('student_group') or '',
                        a.get('subject_name') or '',
                        a.get('teacher_name') or '',
                        att_labels.get(a.get('status'), a.get('status') or ''),
                        a.get('comment') or '',
                    ]
                    for col, v in enumerate(row_vals, 1):
                        cell = ws3.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['center'] if col in (1, 6) else styles['left']

                for col, w in {'A': 12, 'B': 28, 'C': 10, 'D': 28, 'E': 24, 'F': 12, 'G': 30}.items():
                    ws3.column_dimensions[col].width = w
                ws3.freeze_panes = 'A2'

                pd.DataFrame().to_excel(writer, index=False, sheet_name='Расписание')
                ws4 = writer.sheets['Расписание']
                headers4 = ['Группа', 'День', 'Пара', 'Предмет', 'Преподаватель', 'Кабинет', 'Часы']
                for col, h in enumerate(headers4, 1):
                    cell = ws4.cell(row=1, column=col, value=h)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                for i, s in enumerate(schedule_base, start=2):
                    row_vals = [
                        s.get('group_name') or '',
                        s.get('day_of_week') or '',
                        s.get('para_number') or '',
                        s.get('subject_name') or '',
                        s.get('teacher_name') or '',
                        s.get('room_number') or '',
                        s.get('hours') or 2,
                    ]
                    for col, v in enumerate(row_vals, 1):
                        cell = ws4.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['center'] if col in (2, 3, 7) else styles['left']

                for col, w in {'A': 12, 'B': 14, 'C': 6, 'D': 32, 'E': 28, 'F': 10, 'G': 8}.items():
                    ws4.column_dimensions[col].width = w
                ws4.freeze_panes = 'A2'

                if curriculum:
                    pd.DataFrame().to_excel(writer, index=False, sheet_name='Учебный план')
                    ws5 = writer.sheets['Учебный план']
                    headers5 = ['Группа', 'Предмет', 'Преподаватель', 'План, ч', 'Пар/нед',
                                'Факт, ч', 'Остаток, ч', '%']
                    for col, h in enumerate(headers5, 1):
                        cell = ws5.cell(row=1, column=col, value=h)
                        cell.fill = styles['header_fill']
                        cell.font = styles['header_font']
                        cell.alignment = styles['center']
                        cell.border = styles['thin_border']

                    for i, c in enumerate(curriculum, start=2):
                        row_vals = [
                            c.get('group_name') or '',
                            c.get('subject_name') or '',
                            c.get('teacher_short') or c.get('teacher_name') or '',
                            c.get('planned_hours') or 0,
                            c.get('pairs_per_week') or 0,
                            c.get('fact_hours') or 0,
                            c.get('remaining_hours') or 0,
                            c.get('percent') or 0,
                        ]
                        for col, v in enumerate(row_vals, 1):
                            cell = ws5.cell(row=i, column=col, value=v)
                            cell.border = styles['thin_border']
                            cell.alignment = styles['center'] if col not in (2, 3) else styles['left']

                    for col, w in {'A': 12, 'B': 32, 'C': 24, 'D': 10, 'E': 10, 'F': 10, 'G': 12, 'H': 8}.items():
                        ws5.column_dimensions[col].width = w
                    ws5.freeze_panes = 'A2'

            return output

        @app.route('/login', methods=['GET', 'POST'])
        def login():
            if request.method == 'POST':
                username = request.form.get('username')
                password = request.form.get('password')
                user = db.authenticate(username, password)
                if user:
                    logger.info(f"Успешный вход: {username}")
                    next_page = request.args.get('next', '/')
                    response = redirect(next_page)
                    response.set_cookie('token', user['token'], max_age=7 * 24 * 3600,
                                        httponly=True, samesite='Lax')
                    return response
                logger.warning(f"Неудачная попытка входа: {username}")
                return render_template('login.html', error='Неверный логин или пароль')
            return render_template('login.html')

        if LIMITER_AVAILABLE and limiter:
            login = limiter.limit(config.LOGIN_RATE_LIMIT, methods=["POST"])(login)

        @app.route('/register')
        def register():
            return redirect('/login')

        @app.route('/logout')
        def logout():
            user = get_current_user()
            if user:
                logger.info(f"Выход: {user['username']}")
            response = redirect('/')
            response.delete_cookie('token')
            return response

        @app.route('/')
        def index():
            user = get_current_user()
            view_mode = request.args.get('view', 'day')
            current_day = request.args.get('day', 'Понедельник')
            selected_group = request.args.get('group', '')

            if user and user['role'] == 'student' and user.get('group_name') and not selected_group:
                selected_group = user['group_name']

            schedule = db.get_schedule_by_day(current_day) if view_mode != 'week' else {}

            return render_template('schedule.html',
                                   user=user, active_page='schedule', schedule=schedule,
                                   current_day=current_day, days=DAYS_OF_WEEK,
                                   para_times=PARA_TIMES,
                                   groups=db.get_all_groups(),
                                   teachers=db.get_all_teachers(),
                                   subjects=db.get_all_subjects(),
                                   rooms=db.get_all_rooms(),
                                   selected_group=selected_group,
                                   view_mode=view_mode)

        @app.route('/calendar')
        def calendar_page():
            user = get_current_user()
            year = int(request.args.get('year', datetime.now().year))
            month = int(request.args.get('month', datetime.now().month))
            calendar_data = db.get_calendar_month(year, month)
            month_names = ['', 'Январь', 'Февраль', 'Март', 'Апрель', 'Май', 'Июнь',
                           'Июль', 'Август', 'Сентябрь', 'Октябрь', 'Ноябрь', 'Декабрь']
            return render_template('calendar.html',
                                   user=user, active_page='calendar',
                                   year=year, month=month, month_name=month_names[month],
                                   calendar_data={k: dict(v) for k, v in calendar_data.items()},
                                   viewer_group=user.get('group_name') if user and user['role'] == 'student' else None)

        @app.route('/editor')
        def editor_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/editor')
            current_day = request.args.get('day', 'Понедельник')
            today = datetime.now().strftime('%Y-%m-%d')
            current_date = request.args.get('date', today)
            mode = request.args.get('mode', 'weekday')
            return render_template('editor.html',
                                   user=user, active_page='editor',
                                   current_day=current_day, days=DAYS_OF_WEEK,
                                   today=today,
                                   current_date=current_date,
                                   mode=mode,
                                   groups=db.get_all_groups(),
                                   teachers=db.get_all_teachers(),
                                   subjects=db.get_all_subjects(),
                                   rooms=db.get_all_rooms())

        def _render_diary(user, student, is_parent=False, children=None):
            if not student:
                return redirect('/diary')
            date_from = request.args.get('from', '')
            date_to = request.args.get('to', '')

            diary_records = db.get_full_diary(student['id'], date_from or None, date_to or None)
            absent_records = db.get_attendance_only_diary(student['id'], date_from or None, date_to or None)
            subject_summary = db.get_grades_by_subject_period(student['id'], date_from or None, date_to or None)
            stats_grades = db.get_grades_stats(student['id'])
            stats_att = db.get_attendance_stats_period(student['id'], date_from or None, date_to or None)

            return render_template('diary.html',
                                   user=user, active_page='diary',
                                   student=student, children=children or [],
                                   is_parent=is_parent,
                                   diary_records=diary_records,
                                   absent_records=absent_records,
                                   subject_summary=subject_summary,
                                   stats_grades=stats_grades,
                                   stats_att=stats_att,
                                   date_from=date_from, date_to=date_to)

        @app.route('/diary')
        def diary_page():
            user = get_current_user()
            if not user:
                return redirect('/login?next=/diary')

            if user['role'] == 'student':
                return _render_diary(user, user)
            elif user['role'] == 'parent':
                children = db.get_parent_children_detailed(user.get('parent_of', ''))
                if not children:
                    return render_template('diary.html',
                                           user=user, active_page='diary',
                                           student=None, children=[], is_parent=True,
                                           diary_records=[], absent_records=[], subject_summary=[],
                                           stats_grades={'total': 0, 'avg': 0},
                                           stats_att={'total': 0, 'present': 0, 'late': 0, 'absent': 0, 'percent': 0},
                                           date_from=None, date_to=None)
                return redirect(f'/diary/{children[0]["id"]}')
            else:
                return redirect('/')

        @app.route('/diary/<int:student_id>')
        def diary_student(student_id):
            user = get_current_user()
            if not user:
                return redirect('/login?next=/diary')

            if user['role'] == 'student':
                if user['id'] != student_id:
                    return redirect('/diary')
                return _render_diary(user, user)
            elif user['role'] == 'parent':
                children = db.get_parent_children_detailed(user.get('parent_of', ''))
                child_ids = [c['id'] for c in children]
                if student_id not in child_ids:
                    return redirect('/diary')
                student = db.get_user(student_id)
                return _render_diary(user, student, is_parent=True, children=children)
            else:
                return redirect('/')

        @app.route('/gradebook')
        def gradebook_page():
            user = get_current_user()
            if not user or user['role'] not in ['teacher', 'admin']:
                return redirect('/login?next=/gradebook')

            if user['role'] == 'teacher':
                teacher = db.get_teacher_by_user(user['id'])
                if not teacher:
                    return "Ваш аккаунт не привязан к преподавателю.", 403

            current_day = request.args.get('day', 'Понедельник')
            today = datetime.now().strftime('%Y-%m-%d')

            return render_template('gradebook.html',
                                   user=user, active_page='gradebook',
                                   current_day=current_day, days=DAYS_OF_WEEK,
                                   today=today)

        @app.route('/api/teacher/lessons')
        def api_teacher_lessons():
            user = get_current_user()
            if not user or user['role'] not in ['teacher', 'admin']:
                return jsonify([])

            day = request.args.get('day', 'Понедельник')
            date_param = request.args.get('date', '').strip()

            if user['role'] == 'teacher':
                teacher = db.get_teacher_by_user(user['id'])
                if not teacher:
                    return jsonify([])

                lessons = db.get_teacher_schedule(teacher['id'], day)

                if date_param:
                    for l in lessons:
                        ov = db.get_override_for_lesson_date(l['lesson_id'], date_param)
                        if ov:
                            l['override'] = {
                                'id': ov['id'],
                                'type': ov['override_type'],
                                'comment': ov['comment'] or '',
                            }
                            if ov['override_type'] == 'cancelled':
                                l['cancelled'] = True
                                l['subject_name'] = 'ОТМЕНЕНО: ' + (l.get('subject_name') or '')
                            elif ov['override_type'] == 'replaced':
                                if ov.get('new_subject_name'):
                                    l['subject_name'] = ov['new_subject_name']
                                    l['replaced'] = True
                                if ov.get('new_room_number'):
                                    l['room_number'] = ov['new_room_number']
                return jsonify(lessons)

            all_lessons = []
            schedule = db.get_schedule_by_day(day)
            for group_name, paras in schedule.items():
                for para_num, lesson in paras.items():
                    item = {
                        'lesson_id': lesson['id'],
                        'group_name': group_name,
                        'para_number': int(para_num),
                        'subject_name': lesson['предмет'],
                        'room_number': lesson['кабинет']
                    }
                    if date_param:
                        ov = db.get_override_for_lesson_date(lesson['id'], date_param)
                        if ov:
                            item['override'] = {
                                'id': ov['id'],
                                'type': ov['override_type'],
                                'comment': ov['comment'] or '',
                            }
                            if ov['override_type'] == 'cancelled':
                                item['cancelled'] = True
                                item['subject_name'] = 'ОТМЕНЕНО: ' + (item['subject_name'] or '')
                            elif ov['override_type'] == 'replaced':
                                if ov.get('new_subject_name'):
                                    item['subject_name'] = ov['new_subject_name']
                                    item['replaced'] = True
                                if ov.get('new_room_number'):
                                    item['room_number'] = ov['new_room_number']
                    all_lessons.append(item)
            return jsonify(all_lessons)

        @app.route('/api/gradebook/lesson/<int:lesson_id>')
        def api_gradebook_lesson(lesson_id):
            user = get_current_user()
            if not user or user['role'] not in ['teacher', 'admin']:
                return jsonify([])

            if user['role'] == 'teacher':
                teacher = db.get_teacher_by_user(user['id'])
                if not teacher:
                    return jsonify([])
                lesson = db.get_lesson(lesson_id)
                if not lesson or lesson.get('teacher_id') != teacher['id']:
                    return jsonify({'error': 'Нет прав на эту пару'}), 403

            date = request.args.get('date', datetime.now().strftime('%Y-%m-%d'))
            return jsonify(db.get_lesson_students_with_marks(lesson_id, date, auto_mark_present=False))

        @app.route('/api/gradebook/save', methods=['POST'])
        def api_gradebook_save():
            user = get_current_user()
            if not user or user['role'] not in ['teacher', 'admin']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json
            records = data.get('records', [])
            saved_grades = 0
            saved_att = 0

            if user['role'] == 'teacher':
                teacher = db.get_teacher_by_user(user['id'])
                if not teacher:
                    return jsonify({'success': False, 'error': 'Не привязан к преподавателю'})
                lesson_ids = set(r.get('lesson_id') for r in records if r.get('lesson_id'))
                for lid in lesson_ids:
                    lesson = db.get_lesson(lid)
                    if not lesson or lesson.get('teacher_id') != teacher['id']:
                        return jsonify({'success': False, 'error': 'Нет прав на пару'}), 403

            _lesson_ids = set(r.get('lesson_id') for r in records if r.get('lesson_id'))
            _dates = set(r.get('date') for r in records if r.get('date'))
            for _lid in _lesson_ids:
                for _d in _dates:
                    try:
                        db.auto_mark_present_for_lesson(_lid, _d)
                    except Exception as e:
                        logger.warning(f"auto_mark_present failed: {_lid}, {_d}: {e}")

            for rec in records:
                student_id = rec.get('student_id')
                lesson_id = rec.get('lesson_id')
                date = rec.get('date')
                grade = rec.get('grade')
                attendance = rec.get('attendance')
                comment = rec.get('comment', '')

                if not (student_id and lesson_id and date):
                    continue

                existing_grades = db.get_grades_for_lesson_date(lesson_id, date)
                existing_grade = next((g for g in existing_grades if g['student_id'] == student_id), None)

                if grade is not None and grade != '':
                    try:
                        grade_val = int(grade)
                        if grade_val in [2, 3, 4, 5]:
                            if existing_grade:
                                db.delete_grade(existing_grade['id'])
                            db.add_grade(student_id, lesson_id, date, grade_val,
                                         grade_type='regular',
                                         comment=comment, marked_by=user['id'])
                            saved_grades += 1
                    except (ValueError, TypeError):
                        pass
                else:
                    if existing_grade:
                        db.delete_grade(existing_grade['id'])

                if attendance in ('present', 'late', 'absent'):
                    db.mark_attendance(student_id, lesson_id, date,
                                       attendance, comment=comment, marked_by=user['id'])
                    saved_att += 1

            logger.info(f"Журнал сохранён: {saved_grades} оценок, {saved_att} отметок")
            return jsonify({'success': True, 'grades': saved_grades, 'attendance': saved_att})        @app.route('/profile')
        def profile_page():
            user = get_current_user()
            if not user or user['role'] != 'teacher':
                return redirect('/login?next=/profile')

            teacher = db.get_teacher_by_user(user['id'])
            if not teacher:
                return "Ваш аккаунт не привязан к преподавателю.", 403

            date_from = request.args.get('from', '')
            date_to = request.args.get('to', '')

            stats = db.get_teacher_full_stats(teacher['id'])
            groups = db.get_teacher_groups_list(teacher['id'])
            subjects = db.get_teacher_subjects_list(teacher['id'])
            today_lessons = db.get_teacher_lessons_today(teacher['id'])

            avg_grades = db.get_teacher_avg_grade(teacher['id'], date_from or None, date_to or None)
            attendance = db.get_teacher_attendance_by_group(teacher['id'], date_from or None, date_to or None)

            return render_template('profile.html',
                                   user=user, active_page='profile',
                                   teacher=teacher, stats=stats,
                                   groups=groups, subjects=subjects,
                                   today_lessons=today_lessons,
                                   avg_grades=avg_grades,
                                   attendance=attendance,
                                   date_from=date_from, date_to=date_to)

        @app.route('/api/teacher/stats')
        def api_teacher_stats():
            user = get_current_user()
            if not user or user['role'] != 'teacher':
                return jsonify({'success': False, 'error': 'Нет прав'})
            teacher = db.get_teacher_by_user(user['id'])
            if not teacher:
                return jsonify({'success': False, 'error': 'Не привязан'})

            date_from = request.args.get('from', '')
            date_to = request.args.get('to', '')

            return jsonify({
                'success': True,
                'stats': db.get_teacher_full_stats(teacher['id']),
                'avg_grades': db.get_teacher_avg_grade(teacher['id'], date_from or None, date_to or None),
                'attendance': db.get_teacher_attendance_by_group(teacher['id'], date_from or None, date_to or None)
            })

        @app.route('/my-schedule')
        def my_schedule_page():
            user = get_current_user()
            if not user or user['role'] != 'teacher':
                return redirect('/login?next=/my-schedule')

            teacher = db.get_teacher_by_user(user['id'])
            if not teacher:
                return "Ваш аккаунт не привязан к преподавателю.", 403

            schedule = db.get_teacher_schedule_full(teacher['id'])
            has_lessons = any(len(day_data) > 0 for day_data in schedule.values())

            return render_template('my-schedule.html',
                                   user=user, active_page='my-schedule',
                                   teacher=teacher, schedule=schedule,
                                   days=DAYS_OF_WEEK,
                                   para_times=PARA_TIMES,
                                   has_lessons=has_lessons)

        @app.route('/change-password', methods=['GET', 'POST'])
        def change_password_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/change-password')

            error = None
            success = None

            if request.method == 'POST':
                old_password = request.form.get('old_password', '')
                new_password = request.form.get('new_password', '')
                new_password2 = request.form.get('new_password2', '')

                if not db.verify_password(old_password, user['password_hash']):
                    error = 'Неверный текущий пароль'
                elif len(new_password) < 6:
                    error = 'Новый пароль минимум 6 символов'
                elif new_password != new_password2:
                    error = 'Пароли не совпадают'
                elif new_password == old_password:
                    error = 'Новый пароль должен отличаться от старого'
                else:
                    db.change_password(user['id'], new_password)
                    success = 'Пароль успешно изменён'
                    logger.info(f"Пароль изменён: {user['username']}")

            return render_template('change-password.html',
                                   user=user, active_page='change-password',
                                   error=error, success=success)

        @app.route('/performance')
        def performance_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/performance')

            selected_group = request.args.get('group', '')
            date_from = request.args.get('from', '')
            date_to = request.args.get('to', '')

            performance = []
            if selected_group:
                performance = db.get_group_performance(selected_group,
                                                       date_from or None,
                                                       date_to or None)

            return render_template('performance.html',
                                   user=user, active_page='performance',
                                   groups=db.get_all_groups(),
                                   selected_group=selected_group,
                                   date_from=date_from, date_to=date_to,
                                   performance=performance)
        @app.route('/curriculum-v2')
        def curriculum_v2_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/curriculum-v2')

            selected_group_id = request.args.get('group_id', type=int)

            groups = db.get_active_groups_for_schedule()

            if not selected_group_id and groups:
                selected_group_id = groups[0]['id']

            plan = []
            summary = {'by_semester': [], 'total_disciplines': 0,
                       'total_records': 0, 'total_hours': 0}

            if selected_group_id:
                plan = db.get_curriculum_plan_v2(selected_group_id)
                summary = db.get_curriculum_summary(selected_group_id)

            selected_group = None
            if selected_group_id:
                selected_group = db.get_group_by_id(selected_group_id)

            return render_template('curriculum_v2.html',
                                   user=user, active_page='curriculum-v2',
                                   groups=groups,
                                   selected_group_id=selected_group_id,
                                   selected_group=selected_group,
                                   plan=plan,
                                   summary=summary)

        @app.route('/api/curriculum-v2/update-pairs', methods=['POST'])
        def api_curriculum_v2_update_pairs():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            row_id = data.get('id')
            pairs = data.get('pairs_per_week')

            if not row_id:
                return jsonify({'success': False, 'error': 'Не указан ID'})

            try:
                pairs = int(pairs)
                if pairs < 0 or pairs > 20:
                    return jsonify({'success': False, 'error': 'Пары должны быть от 0 до 20'})
            except (ValueError, TypeError):
                return jsonify({'success': False, 'error': 'Неверное значение'})

            conn = db.get_connection()
            cursor = conn.cursor()
            try:
                cursor.execute('''
                    UPDATE curriculum_plan_v2
                    SET pairs_per_week = ?
                    WHERE id = ?
                ''', (pairs, row_id))
                conn.commit()
                return jsonify({'success': True})
            except Exception as e:
                logger.error(f"Ошибка обновления пар: {e}")
                return jsonify({'success': False, 'error': str(e)})
            finally:
                conn.close()

        @app.route('/api/curriculum-v2/clear/<int:group_id>', methods=['POST'])
        def api_curriculum_v2_clear(group_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            deleted = db.delete_curriculum_plan_v2(group_id)
            logger.info(f"Удалён план группы ID={group_id}: {deleted} записей")
            return jsonify({'success': True, 'deleted': deleted})
        @app.route('/api/curriculum-v2/upload/preview', methods=['POST'])
        def api_curriculum_v2_upload_preview():
            """Принимает файл .xls/.xlsx, парсит, возвращает превью."""
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            if 'file' not in request.files:
                return jsonify({'success': False, 'error': 'Файл не выбран'}), 400

            file = request.files['file']
            if not file.filename:
                return jsonify({'success': False, 'error': 'Пустой файл'}), 400

            ext = os.path.splitext(file.filename)[1].lower()
            if ext not in ('.xls', '.xlsx'):
                return jsonify({'success': False, 'error': f'Неверный формат: {ext}. Нужен .xls или .xlsx'}), 400

            import tempfile
            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name

            try:
                from parse_curriculum_v2 import parse_curriculum
                parsed = parse_curriculum(Path(tmp_path))

                if not parsed:
                    return jsonify({'success': False, 'error': 'Не удалось разобрать файл'}), 400

                # Сводка
                total_semesters = {}
                total_disciplines = len(parsed['disciplines'])
                for d in parsed['disciplines']:
                    for sem_num in d['semesters'].keys():
                        total_semesters[sem_num] = total_semesters.get(sem_num, 0) + 1

                return jsonify({
                    'success': True,
                    'direction': parsed.get('direction', ''),
                    'education_form': parsed.get('education_form', ''),
                    'total_disciplines': total_disciplines,
                    'by_semester': total_semesters,
                    'disciplines': [
                        {
                            'index': d['index'],
                            'name': d['name'],
                            'type': d['type'],
                            'registry': d.get('registry', ''),
                            'total_hours': d.get('total_hours', 0),
                            'semesters': list(d['semesters'].keys()),
                        }
                        for d in parsed['disciplines']
                    ],
                    # Для сохранения целиком в БД
                    'raw': parsed,
                })
            except Exception as e:
                logger.exception("Ошибка парсинга УП")
                return jsonify({'success': False, 'error': f'{type(e).__name__}: {e}'}), 500
            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

        @app.route('/api/curriculum-v2/upload/apply', methods=['POST'])
        def api_curriculum_v2_upload_apply():
            """Применяет распарсенный УП к группе (пересохраняет)."""
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            data = request.json or {}
            group_id = data.get('group_id')
            parsed = data.get('parsed')

            if not group_id:
                return jsonify({'success': False, 'error': 'Не указана группа'}), 400
            if not parsed or not parsed.get('disciplines'):
                return jsonify({'success': False, 'error': 'Нет данных для импорта'}), 400

            # Удаляем старый план
            db.delete_curriculum_plan_v2(group_id)

            # Weeks по семестрам (общие для всех УП этого колледжа)
            WEEKS = {1: 17, 2: 23, 3: 17, 4: 19, 5: 14, 6: 11, 7: 14, 8: 11}

            created = 0
            errors = []

            for d in parsed['disciplines']:
                # Находим/создаём предмет
                subject = db.get_subject_by_name(d['name'])
                if subject:
                    subject_id = subject['id']
                else:
                    subject_id = db.add_subject(d['name'])
                    if not subject_id:
                        errors.append(f"Не удалось создать предмет: {d['name']}")
                        continue

                for sem_num_raw, sem in d['semesters'].items():
                    sem_num = int(sem_num_raw)
                    course = (sem_num + 1) // 2
                    semester_in_course = 2 if sem_num % 2 == 0 else 1
                    total_hours = sem.get('total_hours', 0)
                    weeks = WEEKS.get(sem_num, 16)
                    # Формула: Труд-сть / недели / 2
                    pairs = round(total_hours / weeks / 2) if weeks and total_hours else 0
                    if pairs < 1 and total_hours > 0:
                        pairs = 1

                    ok = db.add_curriculum_plan_v2(
                        group_id=group_id,
                        subject_id=subject_id,
                        course=course,
                        semester_in_course=semester_in_course,
                        global_semester=sem_num,
                        total_hours=sem.get('total_hours', 0),
                        lectures=sem.get('lectures', 0),
                        lab=sem.get('lab', 0),
                        practical=sem.get('practical', 0),
                        seminars=sem.get('seminars', 0),
                        individual=sem.get('individual', 0),
                        group_lessons=sem.get('group_lessons', 0),
                        subgroup_lessons=sem.get('subgroup_lessons', 0),
                        self_work=sem.get('self_work', 0),
                        control=sem.get('control', 0),
                        independent=sem.get('independent', 0),
                        contact=sem.get('contact', 0),
                        pairs_per_week=pairs,
                        weeks=weeks,
                        index_code=d['index'],
                        registry_number=d.get('registry', ''),
                        priority=2,
                        created_by=user['id'],
                    )
                    if ok:
                        created += 1

            logger.info(f"Импорт УП: группа ID={group_id}, создано {created} записей")
            return jsonify({
                'success': True,
                'created': created,
                'errors': errors,
            })

        @app.route('/generator-v2')
        def generator_v2_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/generator-v2')
            return render_template('generator_v2.html',
                                   user=user, active_page='generator-v2')

        @app.route('/api/generator-v2/preview', methods=['POST'])
        def api_generator_v2_preview():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            try:
                from auto_generator_v2 import generate_for_all_groups
                result = generate_for_all_groups()

                if result.get('error'):
                    return jsonify({'success': False, 'error': result['error']})

                return jsonify({'success': True, **result})
            except Exception as e:
                logger.exception("Ошибка генерации расписания v2")
                return jsonify({'success': False, 'error': f'{type(e).__name__}: {e}'}), 500

        @app.route('/api/generator-v2/apply', methods=['POST'])
        def api_generator_v2_apply():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            data = request.json or {}
            placed = data.get('placed', [])

            if not placed:
                return jsonify({'success': False, 'error': 'Нет данных'}), 400

            try:
                from auto_generator_v2 import apply_generation
                result = apply_generation(placed, created_by=user['id'])
                logger.info(f"Применено авторасписание: {result['inserted']} вставлено, {result['skipped']} пропущено")
                return jsonify({'success': True, **result})
            except Exception as e:
                logger.exception("Ошибка применения генерации")
                return jsonify({'success': False, 'error': f'{type(e).__name__}: {e}'}), 500
        @app.route('/personal-schedule')
        def personal_schedule_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/personal-schedule')

            teachers = db.get_all_teachers()
            groups = db.get_active_groups_for_schedule()
            subjects = db.get_all_subjects()
            rooms = db.get_all_rooms()

            # Текущий семестр
            now = datetime.now()
            academic_year = f"{now.year}/{now.year + 1}" if now.month >= 9 else f"{now.year - 1}/{now.year}"
            semester_num = 1 if now.month in [9, 10, 11, 12, 1] else 2

            return render_template('personal_schedule.html',
                                   user=user, active_page='personal-schedule',
                                   teachers=teachers,
                                   groups=groups,
                                   subjects=subjects,
                                   rooms=rooms,
                                   academic_year=academic_year,
                                   semester_num=semester_num)

        @app.route('/api/personal-schedule/teacher/<int:teacher_id>')
        def api_personal_schedule_get(teacher_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            entries = db.get_teacher_personal_schedule(teacher_id)
            return jsonify({'success': True, 'entries': entries})

        @app.route('/api/personal-schedule/add', methods=['POST'])
        def api_personal_schedule_add():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            try:
                teacher_id = int(data['teacher_id'])
                day = data['day_of_week']
                para = int(data['para_number'])

                student1_id = int(data['student1_id']) if data.get('student1_id') else None
                student2_id = int(data['student2_id']) if data.get('student2_id') else None
                group_id = int(data['group_id']) if data.get('group_id') else None
                subject_id = int(data['subject_id']) if data.get('subject_id') else None
                room_id = int(data['room_id']) if data.get('room_id') else None

                # Проверка: у преподавателя не занят ли слот
                existing = db.get_teacher_personal_schedule(teacher_id)
                for e in existing:
                    if e['day_of_week'] == day and e['para_number'] == para:
                        return jsonify({'success': False, 'error': f'У преподавателя уже есть занятие на {day}, пара {para}'})

                entry_id = db.add_personal_schedule(
                    teacher_id=teacher_id,
                    day_of_week=day,
                    para_number=para,
                    student1_id=student1_id,
                    student2_id=student2_id,
                    students_raw=data.get('students_raw', ''),
                    group_id=group_id,
                    subject_id=subject_id,
                    room_id=room_id,
                    lesson_type=data.get('lesson_type', 'individual'),
                    semester_num=int(data.get('semester_num') or 1),
                    academic_year=data.get('academic_year', ''),
                    note=data.get('note', ''),
                    created_by=user['id'],
                )
                if entry_id:
                    return jsonify({'success': True, 'id': entry_id})
                return jsonify({'success': False, 'error': 'Не удалось создать'})
            except Exception as e:
                logger.exception("Ошибка api_personal_schedule_add")
                return jsonify({'success': False, 'error': f'{type(e).__name__}: {e}'}), 500

        @app.route('/api/personal-schedule/update/<int:entry_id>', methods=['POST'])
        def api_personal_schedule_update(entry_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            fields = {}
            for key in ('student1_id', 'student2_id', 'group_id', 'subject_id',
                        'room_id', 'lesson_type', 'note'):
                if key in data:
                    v = data[key]
                    if key in ('student1_id', 'student2_id', 'group_id', 'subject_id', 'room_id'):
                        fields[key] = int(v) if v else None
                    else:
                        fields[key] = v

            ok = db.update_personal_schedule(entry_id, **fields)
            return jsonify({'success': ok})

        @app.route('/api/personal-schedule/delete/<int:entry_id>', methods=['POST'])
        def api_personal_schedule_delete(entry_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            ok = db.delete_personal_schedule(entry_id)
            return jsonify({'success': ok})

        @app.route('/api/personal-schedule/students/<int:group_id>')
        def api_personal_schedule_students(group_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            students = db.get_students_in_group(group_id)
            return jsonify({'success': True, 'students': students})
        @app.route('/personal-schedule/all')
        def personal_schedule_all_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/personal-schedule/all')

            # Текущий семестр
            now = datetime.now()
            academic_year = f"{now.year}/{now.year + 1}" if now.month >= 9 else f"{now.year - 1}/{now.year}"
            semester_num = 1 if now.month in [9, 10, 11, 12, 1] else 2

            # Фильтры
            teacher_id = request.args.get('teacher_id', type=int)
            day_of_week = request.args.get('day', '').strip()
            group_id = request.args.get('group_id', type=int)

            entries = db.get_all_personal_schedules_with_teacher(
                semester_num=semester_num,
                academic_year=academic_year,
                teacher_id=teacher_id,
                day_of_week=day_of_week or None,
                group_id=group_id,
            )

            teachers = db.get_all_teachers()
            groups = db.get_active_groups_for_schedule()

            return render_template('personal_schedule_all.html',
                                   user=user, active_page='personal-schedule',
                                   entries=entries,
                                   teachers=teachers,
                                   groups=groups,
                                   teacher_id=teacher_id,
                                   day_of_week=day_of_week,
                                   group_id=group_id,
                                   academic_year=academic_year,
                                   semester_num=semester_num)

        @app.route('/personal-schedule/export')
        def personal_schedule_export():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login')

            now = datetime.now()
            academic_year = f"{now.year}/{now.year + 1}" if now.month >= 9 else f"{now.year - 1}/{now.year}"
            semester_num = 1 if now.month in [9, 10, 11, 12, 1] else 2

            teacher_id = request.args.get('teacher_id', type=int)
            day_of_week = request.args.get('day', '').strip()
            group_id = request.args.get('group_id', type=int)

            entries = db.get_all_personal_schedules_with_teacher(
                semester_num=semester_num,
                academic_year=academic_year,
                teacher_id=teacher_id,
                day_of_week=day_of_week or None,
                group_id=group_id,
            )

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                rows = []
                for e in entries:
                    s1 = e.get('student1_name') or ''
                    s2 = e.get('student2_name') or ''
                    students = s1 + (f', {s2}' if s2 else '')
                    rows.append({
                        'Преподаватель': e.get('teacher_name') or '',
                        'День': e.get('day_of_week') or '',
                        'Пара': e.get('para_number') or '',
                        'Группа': e.get('group_name') or '',
                        'Студенты': students,
                        'Предмет': e.get('subject_name') or '',
                        'Кабинет': e.get('room_number') or '',
                        'Тип': 'Индивидуальное' if e.get('lesson_type') == 'individual' else ('Консультация' if e.get('lesson_type') == 'consultation' else (e.get('lesson_type') or '')),
                    })

                df = pd.DataFrame(rows)
                df.to_excel(writer, index=False, sheet_name='Личные расписания')

                ws = writer.sheets['Личные расписания']
                styles = _excel_header_styles()

                for col in range(1, 9):
                    cell = ws.cell(row=1, column=col)
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                widths = {'A': 30, 'B': 14, 'C': 6, 'D': 10, 'E': 45, 'F': 30, 'G': 10, 'H': 16}
                for col, w in widths.items():
                    ws.column_dimensions[col].width = w

                ws.row_dimensions[1].height = 26
                ws.freeze_panes = 'A2'

            output.seek(0)
            filename = f'Личные_расписания_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name=filename)

        @app.route('/curriculum')
        def curriculum_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/curriculum')

            group_filter = request.args.get('group', '').strip()
            semester = request.args.get('semester', '').strip()
            academic_year = request.args.get('year', '').strip()

            group_id = None
            if group_filter:
                for g in db.get_all_groups():
                    if g['name'] == group_filter:
                        group_id = g['id']
                        break

            rows = db.get_curriculum_with_fact(
                group_id=group_id,
                semester=semester or None,
                academic_year=academic_year or None,
            )

            total_planned = sum(r.get('planned_hours') or 0 for r in rows)
            total_fact = sum(r.get('fact_hours') or 0 for r in rows)
            total_remaining = total_planned - total_fact

            all_years = set()
            for r in db.get_curriculum():
                y = (r.get('academic_year') or '').strip()
                if y:
                    all_years.add(y)
            all_years = sorted(all_years, reverse=True)
            if not all_years:
                all_years = ['2025/2026']

            return render_template('curriculum.html',
                                   user=user, active_page='curriculum',
                                   rows=rows,
                                   groups=db.get_all_groups(),
                                   subjects=db.get_all_subjects(),
                                   teachers=db.get_all_teachers(),
                                   semesters=SEMESTERS,
                                   all_years=all_years,
                                   selected_group=group_filter,
                                   selected_semester=semester,
                                   selected_year=academic_year,
                                   total_planned=total_planned,
                                   total_fact=total_fact,
                                   total_remaining=total_remaining)

        @app.route('/api/curriculum/add', methods=['POST'])
        def api_curriculum_add():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            try:
                group_id = int(data.get('group_id'))
                subject_id = int(data.get('subject_id'))
                teacher_id = int(data['teacher_id']) if data.get('teacher_id') else None
                planned_hours = int(data.get('planned_hours') or 0)
                pairs_per_week = int(data.get('pairs_per_week') or 2)
                semester = str(data.get('semester') or '1').strip()
                academic_year = (data.get('academic_year') or '2025/2026').strip()
                priority = int(data.get('priority') or 2)
                comment = (data.get('comment') or '').strip()

                if planned_hours <= 0:
                    return jsonify({'success': False, 'error': 'План часов должен быть > 0'})
                if pairs_per_week <= 0:
                    return jsonify({'success': False, 'error': 'Пар в неделю должно быть > 0'})

                row_id = db.add_curriculum_row(
                    group_id=group_id,
                    subject_id=subject_id,
                    teacher_id=teacher_id,
                    planned_hours=planned_hours,
                    pairs_per_week=pairs_per_week,
                    semester=semester,
                    academic_year=academic_year,
                    priority=priority,
                    comment=comment,
                    created_by=user['id'],
                )
                if row_id:
                    return jsonify({'success': True, 'id': row_id})
                return jsonify({'success': False, 'error': 'Не удалось добавить'})
            except Exception as e:
                logger.error(f"Ошибка api_curriculum_add: {e}")
                return jsonify({'success': False, 'error': str(e)})

        @app.route('/api/curriculum/update', methods=['POST'])
        def api_curriculum_update():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            row_id = data.get('id')
            if not row_id:
                return jsonify({'success': False, 'error': 'Не указан id'})

            fields = {}
            for key in ('group_id', 'subject_id', 'teacher_id', 'planned_hours',
                        'pairs_per_week', 'semester', 'academic_year', 'priority', 'comment'):
                if key in data:
                    fields[key] = data[key]

            for k in ('group_id', 'subject_id', 'planned_hours', 'pairs_per_week', 'priority'):
                if k in fields and fields[k] is not None:
                    try:
                        fields[k] = int(fields[k])
                    except (ValueError, TypeError):
                        return jsonify({'success': False, 'error': f'Неверное поле: {k}'})

            if 'teacher_id' in fields:
                t = fields['teacher_id']
                fields['teacher_id'] = int(t) if t else None

            ok = db.update_curriculum_row(int(row_id), fields)
            if ok:
                return jsonify({'success': True})
            return jsonify({'success': False, 'error': 'Не удалось обновить'})

        @app.route('/api/curriculum/delete/<int:row_id>', methods=['POST'])
        def api_curriculum_delete(row_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            ok = db.delete_curriculum_row(row_id)
            if ok:
                logger.info(f"Удалена строка плана: {row_id}")
                return jsonify({'success': True})
            return jsonify({'success': False, 'error': 'Не удалось удалить'})
        @app.route('/api/curriculum/count-filtered', methods=['POST'])
        def api_curriculum_count_filtered():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            group_name = (data.get('group') or '').strip()
            semester = (data.get('semester') or '').strip()
            academic_year = (data.get('year') or '').strip()

            group_id = None
            if group_name:
                for g in db.get_all_groups():
                    if g['name'] == group_name:
                        group_id = g['id']
                        break
                if group_id is None:
                    return jsonify({'success': False, 'error': 'Группа не найдена'})

            rows = db.get_curriculum(
                group_id=group_id,
                semester=semester or None,
                academic_year=academic_year or None,
            )

            return jsonify({
                'success': True,
                'count': len(rows),
                'group_name': group_name,
                'semester': semester,
                'academic_year': academic_year,
            })

        @app.route('/api/curriculum/delete-filtered', methods=['POST'])
        def api_curriculum_delete_filtered():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            group_name = (data.get('group') or '').strip()
            semester = (data.get('semester') or '').strip()
            academic_year = (data.get('year') or '').strip()
            confirm_count = data.get('confirm_count')

            if not group_name:
                return jsonify({'success': False, 'error': 'Выберите группу — иначе удалятся строки всех групп'}), 400

            group_id = None
            for g in db.get_all_groups():
                if g['name'] == group_name:
                    group_id = g['id']
                    break
            if group_id is None:
                return jsonify({'success': False, 'error': 'Группа не найдена'}), 404

            rows = db.get_curriculum(
                group_id=group_id,
                semester=semester or None,
                academic_year=academic_year or None,
            )

            if confirm_count is not None and confirm_count != len(rows):
                return jsonify({
                    'success': False,
                    'error': f'Количество изменилось: было {confirm_count}, стало {len(rows)}. Обновите страницу и повторите.',
                }), 409

            deleted = 0
            errors = []
            for r in rows:
                try:
                    if db.delete_curriculum_row(r['id']):
                        deleted += 1
                except Exception as e:
                    errors.append(f"ID {r['id']}: {e}")

            logger.info(f"Массовое удаление УП: {deleted} строк (группа={group_name}, семестр={semester or 'все'}, год={academic_year or 'все'})")

            return jsonify({
                'success': True,
                'deleted': deleted,
                'errors': errors,
            })

        @app.route('/api/curriculum/<int:row_id>')
        def api_curriculum_get(row_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            row = db.get_curriculum_row(row_id)
            if row:
                return jsonify({'success': True, 'row': row})
            return jsonify({'success': False, 'error': 'Строка не найдена'})

        @app.route('/api/curriculum/export')
        def api_curriculum_export():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login')

            group_filter = request.args.get('group', '').strip()
            semester = request.args.get('semester', '').strip()
            academic_year = request.args.get('year', '').strip()

            group_id = None
            if group_filter:
                for g in db.get_all_groups():
                    if g['name'] == group_filter:
                        group_id = g['id']
                        break

            rows = db.get_curriculum_with_fact(
                group_id=group_id,
                semester=semester or None,
                academic_year=academic_year or None,
            )

            output = _build_curriculum_excel(rows, group_filter or None,
                                              semester or None, academic_year or None)
            output.seek(0)
            filename = f'Учебный_план_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name=filename)

        @app.route('/api/curriculum/import/preview', methods=['POST'])
        def api_curriculum_import_preview():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            if 'file' not in request.files:
                return jsonify({'success': False, 'error': 'Файл не выбран'}), 400

            file = request.files['file']
            if not file.filename:
                return jsonify({'success': False, 'error': 'Пустое имя файла'}), 400

            group_name = (request.form.get('group_name') or '').strip()
            if not group_name:
                return jsonify({'success': False, 'error': 'Выберите группу'}), 400

            import tempfile
            ext = os.path.splitext(file.filename)[1].lower()
            if ext not in ('.xls', '.xlsx'):
                return jsonify({'success': False, 'error': f'Неверное расширение: {ext}. Нужен .xls или .xlsx'}), 400

            with tempfile.NamedTemporaryFile(suffix=ext, delete=False) as tmp:
                file.save(tmp.name)
                tmp_path = tmp.name

            try:
                with open(tmp_path, 'rb') as f:
                    head = f.read(8)

                real_ext = ext
                if head.startswith(b'PK\x03\x04'):
                    real_ext = '.xlsx'
                elif head.startswith(b'\xd0\xcf\x11\xe0'):
                    real_ext = '.xls'
                else:
                    try:
                        with open(tmp_path, 'rb') as f:
                            text_head = f.read(512).decode('utf-8', errors='ignore').lower()
                        if '<html' in text_head or '<?xml' in text_head or '<table' in text_head:
                            real_ext = '.html'
                    except Exception:
                        pass

                if real_ext != ext:
                    new_path = tmp_path + real_ext
                    os.rename(tmp_path, new_path)
                    tmp_path = new_path

                if real_ext == '.html':
                    return jsonify({
                        'success': False,
                        'error': 'Файл имеет расширение .xls, но внутри — HTML-таблица. Откройте его в Excel и сохраните как «Книга Excel 97-2003 (*.xls)» или «Книга Excel (*.xlsx)».'
                    }), 400

                from parse_curriculum_excel import parse_curriculum
                parsed = parse_curriculum(tmp_path)

            except ImportError as e:
                return jsonify({'success': False, 'error': f'Не установлена библиотека: {e}. Выполните: pip install xlrd openpyxl'}), 500
            except ValueError as e:
                return jsonify({'success': False, 'error': str(e)}), 400
            except Exception as e:
                logger.exception("Ошибка парсинга УП")
                return jsonify({'success': False, 'error': f'{type(e).__name__}: {e}'}), 500
            finally:
                try:
                    os.unlink(tmp_path)
                except Exception:
                    pass

            known = {s['name'].strip().lower(): s['id'] for s in db.get_all_subjects()}
            for d in parsed['disciplines']:
                key = d['name'].strip().lower()
                d['subject_id'] = known.get(key)
                d['matched'] = d['subject_id'] is not None

            return jsonify({
                'success': True,
                'meta': parsed['meta'],
                'disciplines': parsed['disciplines'],
                'total': len(parsed['disciplines']),
            })

        @app.route('/api/curriculum/import/apply', methods=['POST'])
        def api_curriculum_import_apply():
            try:
                user = get_current_user()
                if not user or user['role'] not in ['admin', 'editor']:
                    return jsonify({'success': False, 'error': 'Нет прав'}), 403

                data = request.json or {}
                group_name = (data.get('group_name') or '').strip()
                academic_year = (data.get('academic_year') or '2025/2026').strip()
                filename = (data.get('filename') or '').strip()
                meta = data.get('meta') or {}
                items = data.get('items') or []

                if not group_name:
                    return jsonify({'success': False, 'error': 'Группа не указана'}), 400
                if not items:
                    return jsonify({'success': False, 'error': 'Нет строк для импорта'}), 400

                group = next((g for g in db.get_all_groups() if g['name'] == group_name), None)
                if not group:
                    return jsonify({'success': False, 'error': 'Группа не найдена'}), 404

                created = 0
                errors = []

                db.save_curriculum_import(
                    group_id=group['id'],
                    filename=filename,
                    meta=meta,
                    disciplines=items,
                    imported_by=user['id'],
                )

                for item in items:
                    subject_id = item.get('subject_id')
                    if not subject_id:
                        name = (item.get('name') or '').strip()
                        if not name:
                            errors.append(f"Пустое название для {item.get('index', '')}")
                            continue
                        existing = next(
                            (s for s in db.get_all_subjects() if s['name'].strip().lower() == name.lower()),
                            None
                        )
                        if existing:
                            subject_id = existing['id']
                        else:
                            subject_id = db.add_subject(name)
                            if not subject_id:
                                errors.append(f"Не удалось создать предмет: {name}")
                                continue

                    auditory_hours = int(item.get('auditory_hours') or 0)
                    if auditory_hours <= 0:
                        errors.append(f"{item.get('index', '')}: часы = 0, пропуск")
                        continue

                    semesters = item.get('semesters') or []
                    if not semesters:
                        semesters = [1]

                    per_sem = max(2, auditory_hours // len(semesters))
                    per_sem = (per_sem // 2) * 2
                    if per_sem < 2:
                        per_sem = 2

                    for sem in semesters:
                        db.add_curriculum_row(
                            group_id=group['id'],
                            subject_id=subject_id,
                            teacher_id=None,
                            planned_hours=per_sem,
                            pairs_per_week=2,
                            semester=str(sem),
                            academic_year=academic_year,
                            priority=2,
                            comment=f"Импорт: {item.get('index', '')} ({item.get('registry', '')})",
                            created_by=user['id'],
                        )
                        created += 1

                return jsonify({
                    'success': True,
                    'created': created,
                    'errors': errors,
                })
            except Exception as e:
                logger.exception("Ошибка apply УП")
                return jsonify({'success': False, 'error': f'{type(e).__name__}: {e}'}), 500

        @app.route('/reports')
        def reports_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/reports')

            mode = request.args.get('mode', 'student')
            group_filter = request.args.get('group', '').strip()
            student_id = request.args.get('student_id', '').strip()
            date_from = request.args.get('from', '')
            date_to = request.args.get('to', '')

            students_in_group = []
            if group_filter:
                students_in_group = db.get_students_by_group(group_filter)

            return render_template('reports.html',
                                   user=user, active_page='reports',
                                   mode=mode,
                                   groups=db.get_all_groups(),
                                   selected_group=group_filter,
                                   students_in_group=students_in_group,
                                   selected_student_id=student_id,
                                   date_from=date_from, date_to=date_to)

        @app.route('/reports/export', methods=['POST'])
        def reports_export():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            data = request.json or {}
            mode = data.get('mode', 'student')
            group_name = (data.get('group') or '').strip()
            student_id = data.get('student_id')
            date_from = (data.get('from') or '').strip()
            date_to = (data.get('to') or '').strip()

            if not group_name:
                return jsonify({'success': False, 'error': 'Выберите группу'}), 400

            try:
                if mode == 'student':
                    if not student_id:
                        return jsonify({'success': False, 'error': 'Выберите ученика'}), 400
                    student = db.get_user(int(student_id))
                    if not student or student['role'] not in ('student', 'viewer'):
                        return jsonify({'success': False, 'error': 'Ученик не найден'}), 400

                    records = db.get_student_diary_for_export(
                        student['id'],
                        date_from or None, date_to or None
                    )
                    output = _build_student_diary_excel(student, group_name, records,
                                                        date_from, date_to)
                    safe_name = (student['full_name'] or 'student').replace(' ', '_')
                    filename = f'Дневник_{safe_name}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
                else:
                    data_group = db.get_group_diary_for_export(
                        group_name, date_from or None, date_to or None
                    )
                    output = _build_group_diary_excel(group_name, data_group,
                                                      date_from, date_to)
                    filename = f'Дневник_группы_{group_name}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
            except Exception as e:
                logger.error(f"Ошибка формирования отчёта: {e}")
                return jsonify({'success': False, 'error': str(e)}), 500

            output.seek(0)
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name=filename)

        @app.route('/api/group/<path:group_name>/students')
        def api_group_students(group_name):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            students = db.get_students_by_group(group_name)
            return jsonify({
                'success': True,
                'students': [{'id': s['id'], 'full_name': s['full_name']} for s in students]
            })

        @app.route('/archive')
        def archive_page():
            user = get_current_user()
            if not user:
                return redirect('/login?next=/archive')
            snapshots = db.list_archive_snapshots()
            return render_template('archive.html',
                                   user=user, active_page='archive',
                                   snapshots=snapshots,
                                   semesters=SEMESTERS)

        @app.route('/api/archive/list')
        def api_archive_list():
            user = get_current_user()
            if not user:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403
            return jsonify({'success': True, 'snapshots': db.list_archive_snapshots()})

        @app.route('/api/archive/create', methods=['POST'])
        def api_archive_create():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            data = request.json or {}
            name = (data.get('name') or '').strip()
            semester = str(data.get('semester') or '').strip()
            academic_year = (data.get('academic_year') or '').strip()
            date_from = (data.get('date_from') or '').strip()
            date_to = (data.get('date_to') or '').strip()

            if not name:
                return jsonify({'success': False, 'error': 'Укажите название снимка'}), 400
            if semester not in SEMESTERS:
                return jsonify({'success': False, 'error': 'Неверный семестр'}), 400
            if not academic_year:
                return jsonify({'success': False, 'error': 'Укажите учебный год'}), 400
            if not date_from or not date_to:
                return jsonify({'success': False, 'error': 'Укажите обе даты'}), 400
            if date_from > date_to:
                return jsonify({'success': False, 'error': 'Дата начала позже даты конца'}), 400

            try:
                snapshot_id = db.create_archive_snapshot(
                    name=name,
                    semester=semester,
                    academic_year=academic_year,
                    date_from=date_from,
                    date_to=date_to,
                    created_by=user['id'],
                )
                if snapshot_id:
                    return jsonify({'success': True, 'id': snapshot_id})
                return jsonify({'success': False, 'error': 'Не удалось создать снимок'}), 500
            except Exception as e:
                logger.error(f"Ошибка создания архива: {e}")
                return jsonify({'success': False, 'error': str(e)}), 500

        @app.route('/api/archive/<int:snapshot_id>')
        def api_archive_get(snapshot_id):
            user = get_current_user()
            if not user:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            snap = db.get_archive_snapshot(snapshot_id)
            if not snap:
                return jsonify({'success': False, 'error': 'Снимок не найден'}), 404

            grades = snap.get('grades') or []
            attendance = snap.get('attendance') or []
            schedule = snap.get('schedule') or {}

            if user['role'] == 'student':
                grades = [g for g in grades if g.get('student_name') == user.get('full_name')]
                attendance = [a for a in attendance if a.get('student_name') == user.get('full_name')]
            elif user['role'] == 'parent':
                children = db.get_parent_children_detailed(user.get('parent_of', ''))
                child_names = {c['full_name'] for c in children}
                if child_names:
                    grades = [g for g in grades if g.get('student_name') in child_names]
                    attendance = [a for a in attendance if a.get('student_name') in child_names]
            elif user['role'] == 'teacher':
                teacher = db.get_teacher_by_user(user['id'])
                if teacher:
                    teacher_name = teacher['full_name']
                    grades = [g for g in grades if g.get('teacher_name') == teacher_name]
                    attendance = [a for a in attendance if a.get('teacher_name') == teacher_name]
                    base = schedule.get('base') or []
                    overrides = schedule.get('overrides') or []
                    schedule = {
                        'base': [s for s in base if s.get('teacher_name') == teacher_name],
                        'overrides': [o for o in overrides if o.get('new_teacher_name') == teacher_name or o.get('teacher_name') == teacher_name],
                    }

            return jsonify({
                'success': True,
                'snapshot': {
                    'id': snap['id'],
                    'name': snap['name'],
                    'semester': snap['semester'],
                    'academic_year': snap['academic_year'],
                    'date_from': snap['date_from'],
                    'date_to': snap['date_to'],
                    'created_at': snap['created_at'],
                    'created_by_name': snap.get('created_by_name'),
                    'stats': snap.get('stats') or {},
                    'schedule': schedule,
                    'grades': grades,
                    'attendance': attendance,
                    'curriculum': snap.get('curriculum') or [],
                }
            })

        @app.route('/api/archive/delete/<int:snapshot_id>', methods=['POST'])
        def api_archive_delete(snapshot_id):
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'}), 403
            ok = db.delete_archive_snapshot(snapshot_id)
            if ok:
                return jsonify({'success': True})
            return jsonify({'success': False, 'error': 'Не удалось удалить'}), 500

        @app.route('/api/archive/export/<int:snapshot_id>')
        def api_archive_export(snapshot_id):
            user = get_current_user()
            if not user:
                return redirect('/login')

            snap = db.get_archive_snapshot(snapshot_id)
            if not snap:
                return "Снимок не найден", 404

            output = _build_archive_excel(snap, user)
            output.seek(0)
            safe_name = (snap['name'] or 'archive').replace(' ', '_')
            filename = f'Архив_{safe_name}_{datetime.now().strftime("%Y%m%d")}.xlsx'
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name=filename)

        @app.route('/rating')
        def rating_page():
            user = get_current_user()
            if not user:
                return redirect('/login?next=/rating')

            if user['role'] in ['student', 'parent']:
                return redirect('/diary')

            can_view_teachers = user['role'] in ['admin', 'editor']

            tab = request.args.get('tab', 'students')
            if tab not in ('students', 'teachers'):
                tab = 'students'
            if tab == 'teachers' and not can_view_teachers:
                tab = 'students'

            date_from = request.args.get('from', '')
            date_to = request.args.get('to', '')

            try:
                min_grades = int(request.args.get('min_grades', 1))
                if min_grades < 0:
                    min_grades = 0
            except (ValueError, TypeError):
                min_grades = 1

            group_filter = request.args.get('group', '').strip()

            students_rating = db.get_students_rating(
                date_from=date_from or None,
                date_to=date_to or None,
                group_name=group_filter or None,
                min_grades=min_grades,
            )

            teachers_rating = []
            if can_view_teachers:
                teachers_rating = db.get_teacher_rating(
                    date_from=date_from or None,
                    date_to=date_to or None,
                    min_grades=min_grades,
                )

            return render_template('rating.html',
                                   user=user, active_page='rating',
                                   tab=tab,
                                   can_view_teachers=can_view_teachers,
                                   students_rating=students_rating,
                                   teachers_rating=teachers_rating,
                                   groups=db.get_all_groups(),
                                   selected_group=group_filter,
                                   date_from=date_from,
                                   date_to=date_to,
                                   min_grades=min_grades)

        @app.route('/generator')
        def generator_page():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return redirect('/login?next=/generator')
            return render_template('generator.html',
                                   user=user, active_page='generator')

        @app.route('/api/generator/parse-excel', methods=['POST'])
        def api_generator_parse_excel():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})

            if 'file' not in request.files:
                return jsonify({'success': False, 'error': 'Файл не выбран'})
            file = request.files['file']
            if file.filename == '':
                return jsonify({'success': False, 'error': 'Файл не выбран'})

            try:
                df = pd.read_excel(file)
                data_list = df.to_dict('records')
                students, errors = db.parse_students_excel(data_list)

                preview = []
                for st in students:
                    prefix = db.make_group_prefix(st['group_name'])
                    preview.append({
                        'full_name': st['full_name'],
                        'group_name': st['group_name'],
                        'prefix': prefix,
                    })

                return jsonify({
                    'success': True,
                    'students': preview,
                    'errors': errors,
                    'total': len(preview),
                })
            except Exception as e:
                return jsonify({'success': False, 'error': str(e)})

        @app.route('/api/generator/create', methods=['POST'])
        def api_generator_create():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})

            data = request.json
            students = data.get('students', [])
            if not students:
                return jsonify({'success': False, 'error': 'Нет студентов для создания'})

            result = db.generate_students_and_parents(students, created_by=user['id'])

            return jsonify({
                'success': True,
                'created': result['created'],
                'errors': result['errors'],
                'total_created': len(result['created']),
            })

        @app.route('/api/generator/credentials')
        def api_generator_credentials():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})
            role_filter = request.args.get('role', '')
            records = db.get_all_credentials(role_filter or None)
            return jsonify({'success': True, 'records': records})

        @app.route('/api/generator/search')
        def api_generator_search():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})
            query = request.args.get('q', '').strip()
            records = db.search_credentials(query) if query else []
            return jsonify({'success': True, 'records': records})

        @app.route('/api/generator/export-excel', methods=['POST'])
        def api_generator_export_excel():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})

            data = request.json
            created = data.get('created', [])
            if not created:
                return jsonify({'success': False, 'error': 'Нет данных для экспорта'})

            rows = []
            for c in created:
                rows.append({
                    'ФИО': c.get('full_name', ''),
                    'Группа': c.get('group_name', ''),
                    'Логин студента': c.get('student_login', ''),
                    'Пароль студента': c.get('student_password', ''),
                    'Логин родителя': c.get('parent_login') or '',
                    'Пароль родителя': c.get('parent_password') or '',
                })

            df = pd.DataFrame(rows)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Логины и пароли')
                worksheet = writer.sheets['Логины и пароли']
                from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
                header_fill = PatternFill(start_color='1A3A5C', end_color='1A3A5C', fill_type='solid')
                header_font = Font(color='FFFFFF', bold=True, size=12)
                thin_border = Border(
                    left=Side(style='thin', color='CCCCCC'),
                    right=Side(style='thin', color='CCCCCC'),
                    top=Side(style='thin', color='CCCCCC'),
                    bottom=Side(style='thin', color='CCCCCC')
                )
                column_widths = {'A': 35, 'B': 12, 'C': 22, 'D': 22, 'E': 25, 'F': 25}
                for col, width in column_widths.items():
                    worksheet.column_dimensions[col].width = width
                for cell in worksheet[1]:
                    cell.fill = header_fill
                    cell.font = header_font
                    cell.alignment = Alignment(horizontal='center', vertical='center')
                    cell.border = thin_border
                for row in worksheet.iter_rows(min_row=2, max_row=worksheet.max_row, max_col=6):
                    for cell in row:
                        cell.border = thin_border
                        cell.alignment = Alignment(vertical='center')
                        if cell.column in [3, 4, 5, 6]:
                            cell.font = Font(name='Consolas', size=11)
                worksheet.row_dimensions[1].height = 25
                for i in range(2, worksheet.max_row + 1):
                    worksheet.row_dimensions[i].height = 22
                worksheet.auto_filter.ref = worksheet.dimensions

            output.seek(0)
            filename = f'Логины_пароли_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name=filename)

        @app.route('/api/generator/reset-password', methods=['POST'])
        def api_generator_reset_password():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})

            data = request.json
            user_id = data.get('user_id')
            role = data.get('role')

            if not user_id or not role:
                return jsonify({'success': False, 'error': 'Не указан пользователь'})

            if role == 'student':
                new_password = db.generate_student_password()
            elif role == 'parent':
                new_password = db.generate_parent_password()
            else:
                return jsonify({'success': False, 'error': 'Только student/parent'})

            db.reset_user_password(user_id, new_password)
            db.update_credentials_password(user_id, new_password)
            return jsonify({'success': True, 'new_password': new_password})

        @app.route('/generator-schedule')
        def generator_schedule_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/generator-schedule')
            return render_template('generator_schedule.html',
                                   user=user, active_page='generator-schedule',
                                   groups=db.get_all_groups(),
                                   semesters=SEMESTERS)

        @app.route('/api/generator/auto-fill', methods=['POST'])
        def api_generator_auto_fill():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            data = request.json or {}
            group_filter = (data.get('group') or '').strip()
            semester = (data.get('semester') or '').strip()
            academic_year = (data.get('year') or '').strip()

            group_id = None
            if group_filter:
                for g in db.get_all_groups():
                    if g['name'] == group_filter:
                        group_id = g['id']
                        break
                if group_id is None:
                    return jsonify({'success': False, 'error': 'Группа не найдена'}), 400

            try:
                result = db.auto_fill_schedule(
                    group_id=group_id,
                    semester=semester or None,
                    academic_year=academic_year or None,
                )
                return jsonify({'success': True, **result})
            except Exception as e:
                logger.error(f"Ошибка auto-fill: {e}")
                return jsonify({'success': False, 'error': str(e)}), 500

        @app.route('/api/generator/apply', methods=['POST'])
        def api_generator_apply():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            data = request.json or {}
            placed = data.get('placed', [])
            if not placed:
                return jsonify({'success': False, 'error': 'Нет данных для применения'}), 400

            try:
                result = db.apply_auto_fill(placed, created_by=user['id'])
                return jsonify({'success': True, **result})
            except Exception as e:
                logger.error(f"Ошибка apply_auto_fill: {e}")
                return jsonify({'success': False, 'error': str(e)}), 500

        @app.route('/api/editor/hints')
        def api_editor_hints():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            group_id = request.args.get('group_id', type=int)
            subject_id = request.args.get('subject_id', type=int)
            teacher_id = request.args.get('teacher_id', type=int)
            day_of_week = (request.args.get('day') or '').strip()
            date_str = (request.args.get('date') or '').strip()
            lesson_id = request.args.get('lesson_id', type=int)

            result = {
                'success': True,
                'curriculum': None,
                'group_day_count': 0,
                'teacher_day_count': 0,
                'warnings': [],
            }

            conn = db.get_connection()
            cursor = conn.cursor()

            try:
                if group_id and subject_id:
                    cursor.execute('''
                        SELECT c.planned_hours, c.pairs_per_week,
                               c.semester, c.academic_year
                        FROM curriculum c
                        WHERE c.group_id = ? AND c.subject_id = ?
                        ORDER BY c.id DESC
                        LIMIT 1
                    ''', (group_id, subject_id))
                    row = cursor.fetchone()

                    if row:
                        cursor.execute('''
                            SELECT COALESCE(SUM(hours), 0) AS fact_hours
                            FROM schedule
                            WHERE group_id = ? AND subject_id = ? AND is_active = 1
                        ''', (group_id, subject_id))
                        fact = cursor.fetchone()['fact_hours'] or 0

                        cursor.execute('''
                            SELECT COALESCE(SUM(s.hours), 0) AS cancelled
                            FROM schedule_overrides o
                            JOIN schedule s ON s.id = o.lesson_id
                            WHERE o.override_type = 'cancelled'
                              AND o.is_active = 1
                              AND s.group_id = ? AND s.subject_id = ? AND s.is_active = 1
                        ''', (group_id, subject_id))
                        cancelled = cursor.fetchone()['cancelled'] or 0

                        fact = max(0, fact - cancelled)
                        planned = row['planned_hours'] or 0
                        remaining = planned - fact

                        result['curriculum'] = {
                            'planned_hours': planned,
                            'fact_hours': fact,
                            'remaining_hours': remaining,
                            'pairs_per_week': row['pairs_per_week'] or 0,
                            'semester': row['semester'],
                            'academic_year': row['academic_year'],
                        }

                        if remaining <= 0:
                            result['warnings'].append({
                                'type': 'plan_exceeded',
                                'text': f'План по предмету исчерпан: {fact} из {planned} ч.',
                            })
                        elif remaining <= 4:
                            result['warnings'].append({
                                'type': 'plan_low',
                                'text': f'Осталось мало часов: {remaining} из {planned} ч.',
                            })

                if group_id and day_of_week:
                    query = '''
                        SELECT COUNT(*) AS cnt FROM schedule
                        WHERE group_id = ? AND day_of_week = ? AND is_active = 1
                    '''
                    params = [group_id, day_of_week]
                    if lesson_id:
                        query += " AND id != ?"
                        params.append(lesson_id)
                    cursor.execute(query, params)
                    result['group_day_count'] = cursor.fetchone()['cnt'] or 0

                if teacher_id and day_of_week:
                    query = '''
                        SELECT COUNT(*) AS cnt FROM schedule
                        WHERE teacher_id = ? AND day_of_week = ? AND is_active = 1
                    '''
                    params = [teacher_id, day_of_week]
                    if lesson_id:
                        query += " AND id != ?"
                        params.append(lesson_id)
                    cursor.execute(query, params)
                    result['teacher_day_count'] = cursor.fetchone()['cnt'] or 0

                if result['group_day_count'] >= 6:
                    result['warnings'].append({
                        'type': 'group_day_full',
                        'text': f"У группы уже {result['group_day_count']} пар в этот день (максимум — 6).",
                    })
                if result['teacher_day_count'] >= 6:
                    result['warnings'].append({
                        'type': 'teacher_day_full',
                        'text': f"У преподавателя уже {result['teacher_day_count']} пар в этот день (максимум — 6).",
                    })

            except Exception as e:
                logger.error(f"Ошибка api_editor_hints: {e}")
                return jsonify({'success': False, 'error': str(e)}), 500
            finally:
                conn.close()

            return jsonify(result)

        @app.route('/import')
        def import_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/import')
            return render_template('import.html',
                                   user=user, active_page='import',
                                   import_log=db.get_import_log() if hasattr(db, 'get_import_log') else [])

        @app.route('/import/template')
        def download_template():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login')

            template_data = [
                {'Группа': '10', 'День': 'Понедельник', 'Пара': 1, 'Предмет': 'Математика',
                 'Преподаватель': 'Новикова Наталья Васильевна', 'Кабинет': '121', 'Часы': 2},
                {'Группа': '10', 'День': 'Понедельник', 'Пара': 2, 'Предмет': 'Русский язык',
                 'Преподаватель': 'Попова Нина Михайловна', 'Кабинет': '213', 'Часы': 2},
            ]
            df = pd.DataFrame(template_data)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Расписание')
            output.seek(0)
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name='Шаблон_расписания.xlsx')
        @app.route('/users/export')
        def users_export():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login')

            role_filter = request.args.get('role', '').strip()
            if role_filter not in ('', 'student', 'parent', 'teacher', 'editor', 'admin'):
                role_filter = ''

            rows = db.get_users_for_export(role_filter or None)

            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df = pd.DataFrame(rows)
                if df.empty:
                    df = pd.DataFrame(columns=[
                        'student_login', 'student_password', 'student_name', 'group_name',
                        'parent_login', 'parent_password', 'parent_name', 'role'
                    ])
                df.to_excel(writer, index=False, sheet_name='Пользователи')

                ws = writer.sheets['Пользователи']
                styles = _excel_header_styles()

                headers_ru = {
                    'student_login': 'Логин студента',
                    'student_password': 'Пароль студента',
                    'student_name': 'ФИО студента',
                    'group_name': 'Группа',
                    'parent_login': 'Логин родителя',
                    'parent_password': 'Пароль родителя',
                    'parent_name': 'ФИО родителя',
                    'role': 'Роль',
                }
                column_order = [
                    'student_login', 'student_password', 'student_name', 'group_name',
                    'parent_login', 'parent_password', 'parent_name', 'role',
                ]

                for col, key in enumerate(column_order, 1):
                    cell = ws.cell(row=1, column=col, value=headers_ru[key])
                    cell.fill = styles['header_fill']
                    cell.font = styles['header_font']
                    cell.alignment = styles['center']
                    cell.border = styles['thin_border']

                for i, r in enumerate(rows, start=2):
                    for col, key in enumerate(column_order, 1):
                        v = r.get(key) or ''
                        if key == 'role':
                            v = {
                                'student': 'Студент',
                                'viewer': 'Студент',
                                'parent': 'Родитель',
                                'teacher': 'Преподаватель',
                                'editor': 'Редактор',
                                'admin': 'Администратор',
                            }.get(v, v)
                        cell = ws.cell(row=i, column=col, value=v)
                        cell.border = styles['thin_border']
                        cell.alignment = styles['left'] if key in (
                            'student_name', 'parent_name', 'group_name'
                        ) else styles['center']
                        if key in ('student_password', 'parent_password', 'student_login', 'parent_login'):
                            cell.font = Font(name='Consolas', size=11)

                widths = {
                    'A': 18, 'B': 22, 'C': 32, 'D': 10,
                    'E': 20, 'F': 22, 'G': 32, 'H': 16,
                }
                for col, w in widths.items():
                    ws.column_dimensions[col].width = w

                ws.row_dimensions[1].height = 26
                for i in range(2, ws.max_row + 1):
                    ws.row_dimensions[i].height = 20

                ws.freeze_panes = 'A2'
                ws.auto_filter.ref = ws.dimensions

            output.seek(0)
            role_label = role_filter or 'all'
            filename = f'Пользователи_{role_label}_{datetime.now().strftime("%Y%m%d_%H%M")}.xlsx'
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True, download_name=filename)

        @app.route('/import/export')
        def export_schedule():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login')

            schedule = db.get_schedule()
            rows = []
            for item in schedule:
                rows.append({
                    'Группа': item['group_name'], 'День': item['day_of_week'],
                    'Пара': item['para_number'], 'Предмет': item['subject_name'],
                    'Преподаватель': item['teacher_name'], 'Кабинет': item['room_number'],
                    'Часы': item['hours']
                })
            if not rows:
                return redirect('/import')
            df = pd.DataFrame(rows)
            output = io.BytesIO()
            with pd.ExcelWriter(output, engine='openpyxl') as writer:
                df.to_excel(writer, index=False, sheet_name='Расписание')
            output.seek(0)
            return send_file(output,
                             mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
                             as_attachment=True,
                             download_name=f'Расписание_{datetime.now().strftime("%Y%m%d")}.xlsx')

        @app.route('/admin')
        def admin_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/admin')

            teachers = db.get_all_teachers()
            duplicates = db.find_duplicate_teachers()
            duplicates_count = len(duplicates)

            return render_template('admin.html',
                                   user=user, active_page='admin',
                                   groups=db.get_all_groups(),
                                   teachers=teachers,
                                   subjects=db.get_all_subjects(),
                                   rooms=db.get_all_rooms(),
                                   duplicates_count=duplicates_count)

        @app.route('/users')
        def users_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/users')
            return render_template('users.html',
                                   user=user, active_page='users',
                                   users=db.get_all_users(),
                                   all_groups=db.get_all_groups(),
                                   all_teachers=db.get_all_teachers())

        @app.route('/stats')
        def stats_page():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return redirect('/login?next=/stats')
            return render_template('stats.html',
                                   user=user, active_page='stats',
                                   overall=db.get_overall_stats(),
                                   teacher_stats=db.get_teacher_stats(),
                                   group_stats=db.get_group_stats())

        @app.route('/admin/backups')
        def admin_backups():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return redirect('/login?next=/admin/backups')
            return render_template('backups.html',
                                   user=user, active_page='backups',
                                   backups=list_backups())

        @app.route('/api/admin/backups/create', methods=['POST'])
        def api_create_backup():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Нет прав'})
            path = make_backup(reason='manual')
            if path:
                return jsonify({'success': True, 'name': path.name})
            return jsonify({'success': False, 'error': 'Не удалось создать'})

        @app.route('/api/admin/backups/download/<name>')
        def api_download_backup(name):
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return redirect('/login')
            safe_name = Path(name).name
            path = config.BACKUPS_PATH / safe_name
            if not path.exists() or not path.is_file():
                return "Not found", 404
            return send_file(path, as_attachment=True, download_name=safe_name)

        @app.route('/api/schedule/<day>')
        def api_get_schedule(day):
            return jsonify(db.get_schedule_by_day(day))

        @app.route('/api/schedule/week/<path:group_name>')
        def api_get_schedule_week(group_name):
            result = {}
            for day in DAYS_OF_WEEK:
                day_data = db.get_schedule_by_day(day, group_name)
                result[day] = day_data.get(group_name, {})
            return jsonify(result)

        @app.route('/api/schedule/week')
        def api_get_schedule_week_all():
            result = {}
            for day in DAYS_OF_WEEK:
                result[day] = db.get_schedule_by_day(day)
            return jsonify(result)

        @app.route('/api/schedule/date/<date_str>')
        def api_get_schedule_date(date_str):
            return jsonify(db.get_schedule_for_date(date_str))

        @app.route('/api/schedule/date-with-overrides/<date_str>')
        def api_schedule_date_with_overrides(date_str):
            try:
                data = db.get_schedule_for_date_with_overrides(date_str)
                return jsonify(data)
            except Exception as e:
                logger.error(f"Ошибка api_schedule_date_with_overrides: {e}")
                return jsonify({'error': str(e)}), 500

        @app.route('/api/groups')
        def api_get_groups():
            return jsonify(db.get_all_groups())

        @app.route('/api/subjects')
        def api_get_subjects():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify([])
            return jsonify(db.get_all_subjects())

        @app.route('/api/rooms/status')
        def api_rooms_status():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'}), 403

            day = (request.args.get('day') or '').strip()
            para = request.args.get('para', type=int)
            lesson_id = request.args.get('lesson_id', type=int)

            if not day or not para:
                return jsonify({'success': False, 'error': 'Укажите day и para'}), 400

            try:
                busy = db.get_busy_rooms_for_slot(day, para, lesson_id)
                return jsonify({'success': True, 'busy_rooms': busy})
            except Exception as e:
                logger.error(f"Ошибка api_rooms_status: {e}")
                return jsonify({'success': False, 'error': str(e)}), 500

        @app.route('/api/lesson/<int:lesson_id>')
        def api_get_lesson(lesson_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            lesson = db.get_lesson(lesson_id)
            if lesson:
                return jsonify({'success': True, 'lesson': lesson})
            return jsonify({'success': False, 'error': 'Пара не найдена'})

        @app.route('/api/check-conflicts', methods=['POST'])
        def api_check_conflicts():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'conflicts': [], 'warnings': []})
            data = request.json
            group_id = int(data.get('group_id'))
            day_of_week = data.get('day_of_week')
            para_number = int(data.get('para_number'))
            teacher_id = int(data.get('teacher_id')) if data.get('teacher_id') else None
            room_id = int(data.get('room_id')) if data.get('room_id') else None
            lesson_id = int(data.get('lesson_id')) if data.get('lesson_id') else None

            conflicts = db.check_conflicts(
                group_id=group_id,
                day_of_week=day_of_week,
                para_number=para_number,
                teacher_id=teacher_id,
                room_id=room_id,
                lesson_id=lesson_id
            )

            warnings = []
            try:
                conn = db.get_connection()
                cursor = conn.cursor()

                cursor.execute('''
                    SELECT COUNT(*) AS cnt FROM schedule
                    WHERE group_id = ? AND day_of_week = ? AND is_active = 1
                      AND (? IS NULL OR id != ?)
                ''', (group_id, day_of_week, lesson_id, lesson_id))
                g_cnt = cursor.fetchone()['cnt'] or 0
                if g_cnt >= 6:
                    warnings.append({
                        'type': 'group_day_full',
                        'description': f"У группы уже <strong>{g_cnt}</strong> пар в этот день (максимум — 6).",
                    })

                if teacher_id:
                    cursor.execute('''
                        SELECT COUNT(*) AS cnt FROM schedule
                        WHERE teacher_id = ? AND day_of_week = ? AND is_active = 1
                          AND (? IS NULL OR id != ?)
                    ''', (teacher_id, day_of_week, lesson_id, lesson_id))
                    t_cnt = cursor.fetchone()['cnt'] or 0
                    if t_cnt >= 6:
                        warnings.append({
                            'type': 'teacher_day_full',
                            'description': f"У преподавателя уже <strong>{t_cnt}</strong> пар в этот день (максимум — 6).",
                        })

                conn.close()
            except Exception as e:
                logger.warning(f"Ошибка warnings в check_conflicts: {e}")

            return jsonify({'success': True, 'conflicts': conflicts, 'warnings': warnings})

        @app.route('/api/schedule/add', methods=['POST'])
        def api_add_lesson():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json
            ignore_conflicts = data.get('ignore_conflicts', False)

            try:
                group_id = int(data['group_id'])
                day_of_week = data['day_of_week']
                para_number = int(data['para_number'])
                subject_id = int(data['subject_id'])
                teacher_id = int(data['teacher_id']) if data.get('teacher_id') else None
                room_id = int(data['room_id']) if data.get('room_id') else None
                hours = int(data.get('hours', 2))

                if not ignore_conflicts:
                    conflicts = db.check_conflicts(
                        group_id=group_id, day_of_week=day_of_week,
                        para_number=para_number, teacher_id=teacher_id, room_id=room_id)
                    if conflicts:
                        return jsonify({'success': False, 'error': 'Обнаружены конфликты', 'conflicts': conflicts})

                conn = db.get_connection()
                cursor = conn.cursor()
                cursor.execute('''
                    UPDATE schedule SET is_active = 0 
                    WHERE group_id = ? AND day_of_week = ? AND para_number = ? AND is_active = 1
                ''', (group_id, day_of_week, para_number))
                conn.commit()
                conn.close()

                lesson_id = db.add_lesson(
                    group_id=group_id, day_of_week=day_of_week, para_number=para_number,
                    subject_id=subject_id, teacher_id=teacher_id, room_id=room_id,
                    hours=hours, created_by=user['id'])
                return jsonify({'success': True, 'id': lesson_id})
            except Exception as e:
                logger.error(f"Ошибка добавления пары: {e}")
                return jsonify({'success': False, 'error': str(e)})

        @app.route('/api/schedule/delete/<int:lesson_id>', methods=['POST'])
        def api_delete_lesson(lesson_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            db.delete_lesson(lesson_id, deleted_by=user['id'])
            return jsonify({'success': True})

        @app.route('/api/schedule/editor/<date_str>')
        def api_schedule_editor(date_str):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'error': 'Нет прав'}), 403
            try:
                data = db.get_schedule_for_date_editor(date_str)
                return jsonify(data)
            except Exception as e:
                logger.error(f"Ошибка api_schedule_editor: {e}")
                return jsonify({'error': str(e)}), 500

        @app.route('/api/schedule/editor/add', methods=['POST'])
        def api_schedule_editor_add():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            try:
                group_id = int(data['group_id'])
                para_number = int(data['para_number'])
                date_str = data['date']
                subject_id = int(data['subject_id'])
                teacher_id = int(data['teacher_id']) if data.get('teacher_id') else None
                room_id = int(data['room_id']) if data.get('room_id') else None
                hours = int(data.get('hours', 2))
                comment = data.get('comment', '').strip()

                override_id = db.add_override(
                    lesson_id=None,
                    date=date_str,
                    override_type='added',
                    new_teacher_id=teacher_id,
                    new_room_id=room_id,
                    new_subject_id=subject_id,
                    new_hours=hours,
                    comment=comment,
                    created_by=user['id'],
                    group_id=group_id,
                    para_number=para_number,
                )
                if override_id:
                    return jsonify({'success': True, 'id': override_id})
                return jsonify({'success': False, 'error': 'Не удалось создать'})
            except Exception as e:
                logger.error(f"Ошибка api_schedule_editor_add: {e}")
                return jsonify({'success': False, 'error': str(e)})

        @app.route('/api/schedule/editor/copy-day', methods=['POST'])
        def api_schedule_editor_copy_day():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            source_day = data.get('source_day')
            target_day = data.get('target_day')
            overwrite = bool(data.get('overwrite', False))

            if not source_day or not target_day:
                return jsonify({'success': False, 'error': 'Не указаны день-источник или день-цель'})

            result = db.copy_schedule_day(source_day, target_day,
                                           created_by=user['id'],
                                           overwrite=overwrite)
            if result.get('error'):
                return jsonify({'success': False, 'error': result['error']})

            return jsonify({
                'success': True,
                'copied': result['copied'],
                'deleted': result['deleted'],
            })

        @app.route('/api/override/add', methods=['POST'])
        def api_override_add():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            try:
                date_str = data.get('date', '').strip()
                override_type = data.get('override_type', '').strip()
                lesson_id = data.get('lesson_id')

                if not date_str or override_type not in ('cancelled', 'replaced', 'added'):
                    return jsonify({'success': False, 'error': 'Неверные параметры'})

                new_teacher_id = data.get('new_teacher_id')
                new_room_id = data.get('new_room_id')
                new_subject_id = data.get('new_subject_id')
                new_hours = data.get('new_hours')
                comment = data.get('comment', '').strip()
                group_id = data.get('group_id')
                para_number = data.get('para_number')

                new_teacher_id = int(new_teacher_id) if new_teacher_id else None
                new_room_id = int(new_room_id) if new_room_id else None
                new_subject_id = int(new_subject_id) if new_subject_id else None
                new_hours = int(new_hours) if new_hours else None
                lesson_id = int(lesson_id) if lesson_id else None
                group_id = int(group_id) if group_id else None
                para_number = int(para_number) if para_number else None

                if override_type == 'added':
                    if not group_id or not para_number or not new_subject_id:
                        return jsonify({'success': False, 'error': 'Укажите группу, номер пары и предмет'})
                else:
                    if not lesson_id:
                        return jsonify({'success': False, 'error': 'Не указана пара'})
                    lesson = db.get_lesson(lesson_id)
                    if not lesson:
                        return jsonify({'success': False, 'error': 'Пара не найдена'})

                if override_type == 'replaced' and not any([new_teacher_id, new_room_id, new_subject_id, new_hours]):
                    return jsonify({'success': False, 'error': 'Укажите хотя бы одно изменение'})

                override_id = db.add_override(
                    lesson_id=lesson_id,
                    date=date_str,
                    override_type=override_type,
                    new_teacher_id=new_teacher_id,
                    new_room_id=new_room_id,
                    new_subject_id=new_subject_id,
                    new_hours=new_hours,
                    comment=comment,
                    created_by=user['id'],
                    group_id=group_id,
                    para_number=para_number,
                )

                if override_id:
                    return jsonify({'success': True, 'id': override_id})
                return jsonify({'success': False, 'error': 'Не удалось создать'})
            except Exception as e:
                logger.error(f"Ошибка api_override_add: {e}")
                return jsonify({'success': False, 'error': str(e)})

        @app.route('/api/override/delete/<int:override_id>', methods=['POST'])
        def api_override_delete(override_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            if db.delete_override(override_id, deleted_by=user['id']):
                return jsonify({'success': True})
            return jsonify({'success': False, 'error': 'Не удалось удалить'})

        @app.route('/api/override/lesson/<int:lesson_id>')
        def api_override_get_for_lesson(lesson_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor', 'teacher']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            date = request.args.get('date', '').strip()
            if not date:
                return jsonify({'success': False, 'error': 'Не указана дата'})
            override = db.get_override_for_lesson_date(lesson_id, date)
            return jsonify({'success': True, 'override': override})

        @app.route('/api/override/date/<date>')
        def api_override_get_for_date(date):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            overrides = db.get_overrides_for_date(date)
            return jsonify({'success': True, 'overrides': overrides})

        @app.route('/api/users/list')
        def api_users_list():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            users = db.get_all_users()
            if user['role'] == 'editor':
                users = [u for u in users if u['role'] != 'admin']
            return jsonify({'success': True, 'users': users})

        @app.route('/api/users/<int:user_id>')
        def api_user_detail(user_id):
            current = get_current_user()
            if not current or current['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            target = db.get_user(user_id)
            if not target:
                return jsonify({'success': False, 'error': 'Пользователь не найден'})
            if current['role'] == 'editor' and target['role'] == 'admin':
                return jsonify({'success': False, 'error': 'Редактор не может просматривать админа'})
            return jsonify({'success': True, 'user': target})

        @app.route('/api/users/update', methods=['POST'])
        def api_update_user():
            current = get_current_user()
            if not current or current['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json
            user_id = data.get('user_id')
            if not user_id:
                return jsonify({'success': False, 'error': 'Не указан пользователь'})

            target = db.get_user(user_id)
            if not target:
                return jsonify({'success': False, 'error': 'Пользователь не найден'})

            if current['role'] == 'editor' and target['role'] == 'admin':
                return jsonify({'success': False, 'error': 'Редактор не может редактировать админа'})

            fields = {}
            if 'full_name' in data:
                fn = data['full_name'].strip()
                if not fn:
                    return jsonify({'success': False, 'error': 'ФИО не может быть пустым'})
                fields['full_name'] = fn
            if 'email' in data:
                fields['email'] = data['email'].strip()
            if 'phone' in data:
                fields['phone'] = data['phone'].strip()

            if target['role'] == 'student' and 'group_name' in data:
                gn = data['group_name'].strip()
                if not gn:
                    return jsonify({'success': False, 'error': 'Группа обязательна'})
                fields['group_name'] = gn

            if target['role'] == 'parent' and 'parent_of' in data:
                po = data['parent_of'].strip()
                if not po:
                    return jsonify({'success': False, 'error': 'Укажите логин ребёнка'})
                fields['parent_of'] = po

            if target['role'] == 'teacher' and 'teacher_id' in data:
                tid = data['teacher_id']
                fields['teacher_id'] = int(tid) if tid else None

            ok = db.update_user(user_id, fields)
            if not ok:
                return jsonify({'success': False, 'error': 'Не удалось обновить'})

            return jsonify({'success': True, 'user': db.get_user(user_id)})

        @app.route('/api/users/reset-password', methods=['POST'])
        def api_users_reset_password():
            current = get_current_user()
            if not current or current['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json
            user_id = data.get('user_id')
            new_password = data.get('new_password', '').strip()
            auto_generate = data.get('auto_generate', False)

            if not user_id:
                return jsonify({'success': False, 'error': 'Не указан пользователь'})

            target = db.get_user(user_id)
            if not target:
                return jsonify({'success': False, 'error': 'Пользователь не найден'})

            if current['role'] == 'editor' and target['role'] == 'admin':
                return jsonify({'success': False, 'error': 'Нельзя менять пароль админа'})

            if current['id'] == user_id:
                return jsonify({'success': False, 'error': 'Для смены своего пароля используйте /change-password'})

            if auto_generate:
                role = target['role']
                if role == 'student':
                    new_password = db.generate_student_password()
                elif role == 'parent':
                    new_password = db.generate_parent_password()
                else:
                    new_password = db.generate_random_password(10)
            else:
                if len(new_password) < 6:
                    return jsonify({'success': False, 'error': 'Пароль минимум 6 символов'})

            db.reset_user_password(user_id, new_password)
            db.update_credentials_password(user_id, new_password)
            return jsonify({'success': True, 'new_password': new_password})

        @app.route('/api/users/add', methods=['POST'])
        def api_add_user():
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})

            data = request.json
            username = data.get('username', '').strip()
            password = data.get('password', '')
            full_name = data.get('full_name', '').strip()
            role = data.get('role', 'student')
            email = data.get('email', '').strip()
            group_name = data.get('group_name', '').strip()
            parent_of = data.get('parent_of', '').strip()
            teacher_id = data.get('teacher_id')

            if not username or len(username) < 3:
                return jsonify({'success': False, 'error': 'Логин минимум 3 символа'})
            if not password or len(password) < 6:
                return jsonify({'success': False, 'error': 'Пароль минимум 6 символов'})
            if not full_name:
                return jsonify({'success': False, 'error': 'Введите ФИО'})
            if role not in ['student', 'editor', 'admin', 'teacher', 'parent']:
                return jsonify({'success': False, 'error': 'Неверная роль'})
            if db.user_exists(username):
                return jsonify({'success': False, 'error': f'Логин "{username}" занят'})
            if role == 'student' and not group_name:
                return jsonify({'success': False, 'error': 'Для студента нужна группа'})
            if role == 'parent' and not parent_of:
                return jsonify({'success': False, 'error': 'Укажите логин ребёнка'})

            created_new_teacher = False
            if role == 'teacher':
                if not teacher_id:
                    tid, created_new_teacher, _ = db.find_or_create_teacher(full_name, email)
                    teacher_id = tid

            new_id = db.create_user(
                username=username, password=password, full_name=full_name,
                role=role, email=email,
                group_name=group_name if role == 'student' else '',
                parent_of=parent_of if role == 'parent' else '',
                teacher_id=teacher_id if role == 'teacher' else None
            )

            if new_id:
                if role == 'teacher' and teacher_id:
                    db.link_teacher_user(teacher_id, new_id)
                db.save_credentials(
                    user_id=new_id, username=username, password_plain=password,
                    role=role, full_name=full_name,
                    group_name=group_name if role == 'student' else '',
                    created_by=user['id']
                )
                return jsonify({
                    'success': True,
                    'id': new_id,
                    'teacher_linked': (role == 'teacher'),
                    'teacher_created': created_new_teacher
                })
            return jsonify({'success': False, 'error': 'Не удалось создать'})

        @app.route('/api/users/delete/<int:user_id>', methods=['POST'])
        def api_delete_user(user_id):
            current = get_current_user()
            if not current or current['role'] != 'admin':
                return jsonify({'success': False, 'error': 'Только администратор'})
            if current['id'] == user_id:
                return jsonify({'success': False, 'error': 'Нельзя удалить себя'})
            target = db.get_user(user_id)
            if not target:
                return jsonify({'success': False, 'error': 'Пользователь не найден'})
            db.delete_user(user_id)
            return jsonify({'success': True, 'deleted': target['username']})

        @app.route('/api/admin/group', methods=['POST'])
        def api_add_group():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False})
            data = request.json
            group_id = db.add_group(data['name'], data.get('specialty'))
            if group_id:
                return jsonify({'success': True, 'id': group_id})
            return jsonify({'success': False, 'error': 'Группа существует'})

        @app.route('/api/admin/group/<int:group_id>', methods=['DELETE'])
        def api_delete_group(group_id):
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False})
            db.delete_group(group_id)
            return jsonify({'success': True})

        @app.route('/api/admin/teacher', methods=['POST'])
        def api_add_teacher():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False})
            data = request.json
            teacher_id = db.add_teacher(data['full_name'])
            return jsonify({'success': True, 'id': teacher_id})

        @app.route('/api/admin/teacher/<int:teacher_id>', methods=['DELETE'])
        def api_delete_teacher(teacher_id):
            user = get_current_user()
            if not user or user['role'] != 'admin':
                return jsonify({'success': False})
            db.delete_teacher(teacher_id)
            return jsonify({'success': True})

        @app.route('/api/admin/subject', methods=['POST'])
        def api_add_subject():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False})
            data = request.json
            subject_id = db.add_subject(data['name'])
            return jsonify({'success': True, 'id': subject_id})

        @app.route('/api/admin/subject/<int:subject_id>', methods=['DELETE'])
        def api_delete_subject(subject_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            result = db.delete_subject(subject_id)
            return jsonify(result)

        @app.route('/api/admin/room', methods=['POST'])
        def api_add_room():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False})
            data = request.json
            room_id = db.add_room(data['number'])
            return jsonify({'success': True, 'id': room_id})

        @app.route('/api/admin/room/<int:room_id>', methods=['DELETE'])
        def api_delete_room(room_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            result = db.delete_room(room_id)
            return jsonify(result)

        @app.route('/api/admin/room/<int:room_id>')
        def api_get_room(room_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            room = db.get_room(room_id)
            if not room:
                return jsonify({'success': False, 'error': 'Кабинет не найден'})

            return jsonify({'success': True, 'room': room})
        @app.route('/api/admin/room/<int:room_id>/update', methods=['POST'])
        def api_update_room(room_id):
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            data = request.json or {}
            fields = {}

            if 'number' in data:
                number = (data['number'] or '').strip()
                if not number:
                    return jsonify({'success': False, 'error': 'Номер не может быть пустым'})
                fields['number'] = number

            if 'building' in data:
                fields['building'] = (data['building'] or '').strip()

            if 'room_type' in data:
                fields['room_type'] = (data['room_type'] or '').strip()

            if 'capacity' in data:
                try:
                    cap = int(data['capacity']) if data['capacity'] else None
                    fields['capacity'] = cap
                except (ValueError, TypeError):
                    return jsonify({'success': False, 'error': 'Вместимость — число'})

            if 'subjects_allowed' in data:
                # Список ID предметов через запятую
                raw = data['subjects_allowed'] or []
                if isinstance(raw, list):
                    fields['subjects_allowed'] = ','.join(str(int(x)) for x in raw if x)
                else:
                    fields['subjects_allowed'] = str(raw).strip()

            if not fields:
                return jsonify({'success': False, 'error': 'Нет полей для обновления'})

            ok = db.update_room(room_id, fields)
            if ok:
                logger.info(f"Кабинет обновлён: id={room_id}, поля={list(fields.keys())}")
                return jsonify({'success': True})
            return jsonify({'success': False, 'error': 'Не удалось обновить'})

        @app.route('/api/admin/teachers/duplicates')
        def api_teachers_duplicates():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            duplicates = db.find_duplicate_teachers()
            total_dup = sum(len(g['duplicates']) for g in duplicates)
            return jsonify({
                'success': True,
                'groups': duplicates,
                'total_groups': len(duplicates),
                'total_duplicates': total_dup,
            })

        @app.route('/api/admin/teachers/cleanup', methods=['POST'])
        def api_teachers_cleanup():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})
            result = db.delete_teacher_duplicates()
            return jsonify({
                'success': True,
                'deleted': result['deleted'],
                'skipped': result['skipped'],
            })

        @app.route('/api/import', methods=['POST'])
        def api_import():
            user = get_current_user()
            if not user or user['role'] not in ['admin', 'editor']:
                return jsonify({'success': False, 'error': 'Нет прав'})

            if 'file' not in request.files:
                return jsonify({'success': False, 'error': 'Файл не выбран'})
            file = request.files['file']
            if file.filename == '':
                return jsonify({'success': False, 'error': 'Файл не выбран'})

            try:
                df = pd.read_excel(file)
                required = ['Группа', 'День', 'Пара', 'Предмет']
                missing = [col for col in required if col not in df.columns]
                if missing:
                    return jsonify({'success': False, 'error': f'Отсутствуют: {", ".join(missing)}'})
                data_list = df.to_dict('records')
                result = db.import_schedule_from_list(data_list, user_id=user['id'], filename=file.filename)
                return jsonify({
                    'success': True,
                    'imported': result['imported'],
                    'errors': result['errors'],
                    'error_messages': result['error_messages'][:50]
                })
            except Exception as e:
                logger.error(f"Ошибка импорта: {e}")
                return jsonify({'success': False, 'error': str(e)})
        self.flask_app = app
        logger.info("Flask-приложение создано")

    def create_widgets(self):
        notebook = ttk.Notebook(self.root)
        notebook.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        web_frame = ttk.Frame(notebook)
        notebook.add(web_frame, text="Веб-сервер")
        self.create_web_tab(web_frame)
        info_frame = ttk.Frame(notebook)
        notebook.add(info_frame, text="Информация")
        self.create_info_tab(info_frame)

    def create_web_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        ttk.Label(main_frame, text="Управление веб-сервером",
                  font=('Arial', 16, 'bold')).pack(pady=10)
        info_frame = ttk.LabelFrame(main_frame, text="Информация о сервере")
        info_frame.pack(fill=tk.X, pady=10)
        status_frame = ttk.Frame(info_frame)
        status_frame.pack(fill=tk.X, pady=5, padx=10)
        ttk.Label(status_frame, text="Статус:", font=('Arial', 11, 'bold')).pack(side=tk.LEFT)
        self.web_status = tk.StringVar(value="Остановлен")
        ttk.Label(status_frame, textvariable=self.web_status, font=('Arial', 11)).pack(side=tk.LEFT, padx=10)
        url_frame = ttk.Frame(info_frame)
        url_frame.pack(fill=tk.X, pady=5, padx=10)
        ttk.Label(url_frame, text="Локальный:", font=('Arial', 11, 'bold')).pack(side=tk.LEFT)
        ttk.Label(url_frame, text=f"http://localhost:{config.PORT}", font=('Arial', 11),
                  foreground='blue').pack(side=tk.LEFT, padx=10)
        network_frame = ttk.Frame(info_frame)
        network_frame.pack(fill=tk.X, pady=5, padx=10)
        ttk.Label(network_frame, text="В сети:", font=('Arial', 11, 'bold')).pack(side=tk.LEFT)
        self.network_url = tk.StringVar(value="Определяется...")
        ttk.Label(network_frame, textvariable=self.network_url, font=('Arial', 11),
                  foreground='blue').pack(side=tk.LEFT, padx=10)
        cred_frame = ttk.LabelFrame(main_frame, text="Учётные данные")
        cred_frame.pack(fill=tk.X, pady=10)
        cred_text = tk.Text(cred_frame, height=15, width=60, font=('Consolas', 10))
        cred_text.pack(fill=tk.X, padx=10, pady=10)
        cred_text.insert(1.0,
                         "Администратор: admin / admin123\n\n"
                         "РОЛИ:\n"
                         "admin   - полный доступ\n"
                         "editor  - расписание + пользователи\n"
                         "teacher - журнал + профиль\n"
                         "student - расписание + дневник\n"
                         "parent  - дневник ребёнка\n\n"
                         "Смена пароля: /change-password\n"
                         "Очистка дублей: /admin\n"
                         "Бэкапы БД: /admin/backups\n"
                         "Рейтинг: /rating\n"
                         "Отчёты (Excel): /reports\n"
                         "Учебный план: /curriculum\n"
                         "Генератор расписания: /generator-schedule\n"
                         "Архив семестров: /archive\n"
                         "Замены пар: /editor")
        cred_text.config(state=tk.DISABLED)
        btn_frame = ttk.Frame(main_frame)
        btn_frame.pack(pady=15)
        ttk.Button(btn_frame, text="Запустить сервер",
                   command=self.start_web_server, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Остановить",
                   command=self.stop_web_server, width=20).pack(side=tk.LEFT, padx=5)
        ttk.Button(btn_frame, text="Открыть браузер",
                   command=self.open_browser, width=20).pack(side=tk.LEFT, padx=5)
        log_frame = ttk.LabelFrame(main_frame, text="Лог")
        log_frame.pack(fill=tk.BOTH, expand=True, pady=10)
        self.web_log = scrolledtext.ScrolledText(log_frame, height=10, font=('Consolas', 9))
        self.web_log.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)

    def create_info_tab(self, parent):
        main_frame = ttk.Frame(parent)
        main_frame.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        ttk.Label(main_frame, text="О системе", font=('Arial', 16, 'bold')).pack(pady=10)
        info_text = tk.Text(main_frame, height=25, width=80, font=('Arial', 10))
        info_text.pack(fill=tk.BOTH, expand=True, pady=10)
        info_text.insert(1.0, """УПРАВЛЕНИЕ РАСПИСАНИЕМ - ВЕРСИЯ 16.6

БЕЗОПАСНОСТЬ:
- bcrypt для паролей
- Rate limiting на /login
- SECRET_KEY обязателен (.env)

БЭКАПЫ:
- Автобэкап при запуске
- Ежедневный в 03:00
- /admin/backups

ПРОИЗВОДИТЕЛЬНОСТЬ:
- WAL-режим SQLite
- Индексы

ЛОГИРОВАНИЕ:
- logs/app.log

АВТО-ПОСЕЩАЕМОСТЬ:
- При сохранении журнала

РЕЙТИНГ:
- /rating - вкладки "Студенты" и "Преподаватели"

ОТЧЁТЫ:
- /reports - экспорт дневника ученика или группы в Excel

УЧЕБНЫЙ ПЛАН:
- /curriculum - план часов, пар в неделю
- Факт из расписания минус отмены
- Экспорт в Excel
- Импорт из Excel (.xls/.xlsx)

ГЕНЕРАТОР РАСПИСАНИЯ:
- /generator-schedule - автогенерация по плану
- Приоритеты, конфликты, превью
- Применение одним нажатием

АРХИВ СЕМЕСТРОВ:
- /archive - снимки за период
- Ролевой доступ (ученик/родитель/учитель)
- Экспорт в Excel

ЗАМЕНЫ И ОТМЕНЫ:
- Override на конкретную дату

МАССОВОЕ КОПИРОВАНИЕ:
- Копирование дня недели

УДАЛЕНИЕ СПРАВОЧНИКОВ:
- Предметы и кабинеты - с проверкой использования

ПОДСКАЗКИ В РЕДАКТОРЕ:
- План/факт/остаток по предмету
- Пары в день у группы и преподавателя
- Свободные кабинеты на слот
""")
        info_text.config(state=tk.DISABLED)

    def start_web_server(self):
        if not FLASK_AVAILABLE:
            messagebox.showerror("Ошибка", "Flask не установлен")
            return
        if self.web_server_running:
            messagebox.showinfo("Информация", "Сервер уже запущен")
            return

        def run_server():
            try:
                hostname = socket.gethostname()
                local_ip = socket.gethostbyname(hostname)
                server_url = f"http://{local_ip}:{config.PORT}"
                self.network_url.set(server_url)
                self.log(f"Запуск сервера на {local_ip}:{config.PORT}")
                self.log(f"Доступ: {server_url}")
                self.log(f"Локально: http://localhost:{config.PORT}")
                self.web_status.set("Запущен")
                self.web_server_running = True
                if WAITRESS_AVAILABLE:
                    self.log("Используется Waitress")
                    serve(self.flask_app, host=config.HOST, port=config.PORT, threads=4)
                else:
                    self.log("Используется встроенный сервер Flask")
                    import logging
                    logging.getLogger('werkzeug').setLevel(logging.ERROR)
                    self.flask_app.run(host=config.HOST, port=config.PORT, debug=False,
                                       use_reloader=False, threaded=True)
            except Exception as e:
                self.log(f"Ошибка: {e}")
                logger.error(f"Ошибка веб-сервера: {e}")
                self.web_status.set("Ошибка")
                self.web_server_running = False

        self.web_server_running = True
        self.web_status.set("Запуск...")
        self.log("Запуск веб-сервера...")
        self.server_thread = threading.Thread(target=run_server, daemon=True)
        self.server_thread.start()

    def stop_web_server(self):
        self.web_server_running = False
        self.web_status.set("Остановлен")
        self.log("Сервер остановлен")

    def open_browser(self):
        try:
            webbrowser.open(f"http://localhost:{config.PORT}")
            self.log("Открыт браузер")
        except Exception as e:
            messagebox.showerror("Ошибка", str(e))

    def log(self, message):
        timestamp = datetime.now().strftime("%H:%M:%S")
        self.web_log.insert(tk.END, f"[{timestamp}] {message}\n")
        self.web_log.see(tk.END)
        self.root.update()

def main():
    logger.info("=" * 60)
    logger.info("ЗАПУСК СИСТЕМЫ РАСПИСАНИЯ")
    logger.info("=" * 60)
    logger.info("Создание стартового бэкапа...")
    make_backup(reason='startup')
    current_year = datetime.now().year
    db.generate_calendar(current_year)
    db.generate_calendar(current_year + 1)
    start_backup_scheduler()
    root = tk.Tk()
    app = ScheduleApp(root)
    root.mainloop()


if __name__ == '__main__':
    main()