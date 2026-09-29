"""
Генератор расписания v2 — на семестр для всех активных групп.
"""
import sys
from pathlib import Path
from typing import List, Dict, Any, Optional
from datetime import datetime

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from database import db, DAYS_OF_WEEK
from config import config
from logger import get_logger

logger = get_logger(__name__)

PARAS = [1, 2, 3, 4, 5, 6]
MAX_PAIRS_PER_DAY = 6


def get_current_semester_info():
    """
    Определяем текущий семестр по дате:
    - Сентябрь-Декабрь + Январь → осенний (semester_in_course=1)
    - Февраль-Июнь → весенний (semester_in_course=2)
    Возвращает (academic_year_start, semester_in_course).
    """
    now = datetime.now()
    if now.month >= 9:
        # Сентябрь-Декабрь → 2025/2026, осень
        return now.year, 1
    elif now.month == 1:
        # Январь → тот же учебный год, осень
        return now.year - 1, 1
    else:
        # Февраль-Август → весна второго полугодия
        return now.year - 1, 2


def get_busy_slots():
    """Возвращает занятые слоты (день, пара) для групп, преподавателей и кабинетов."""
    conn = db.get_connection()
    cursor = conn.cursor()
    cursor.execute('''
        SELECT group_id, teacher_id, room_id, day_of_week, para_number
        FROM schedule
        WHERE is_active = 1
    ''')
    result = {
        'group': {},
        'teacher': {},
        'room': {},
    }
    for row in cursor.fetchall():
        key = (row['day_of_week'], row['para_number'])
        if row['group_id']:
            result['group'].setdefault(row['group_id'], set()).add(key)
        if row['teacher_id']:
            result['teacher'].setdefault(row['teacher_id'], set()).add(key)
        if row['room_id']:
            result['room'].setdefault(row['room_id'], set()).add(key)
    conn.close()
    return result


def generate_for_all_groups():
    """
    Генерирует расписание для всех активных групп.
    Возвращает превью: список пар + статистику.
    """
    academic_year, semester_in_course = get_current_semester_info()
    semester_label = 'осень' if semester_in_course == 1 else 'весна'

    logger.info(f"🎯 Генерация расписания: {academic_year}/{academic_year+1}, {semester_label}")

    groups = db.get_active_groups_for_schedule()

    if not groups:
        return {
            'error': 'Нет активных групп',
            'placed': [],
            'conflicts': [],
            'stats': {'total_groups': 0, 'total_placed': 0, 'total_conflicts': 0},
        }

    # Занятые слоты (существующие)
    busy = get_busy_slots()

    placed = []
    conflicts = []
    group_stats = []

    for group in groups:
        group_id = group['id']
        course = group['current_course']

        # Берём план группы на её текущий курс и текущий семестр
        plan_rows = db.get_curriculum_plan_v2(
            group_id=group_id,
            course=course,
            semester_in_course=semester_in_course,
        )

        if not plan_rows:
            group_stats.append({
                'group_id': group_id,
                'group_name': group['name'],
                'course': course,
                'placed': 0,
                'wanted': 0,
                'reason': 'Нет плана',
            })
            continue

        wanted_total = sum(r['pairs_per_week'] or 0 for r in plan_rows)
        placed_count = 0

        # Сортируем: сначала с большими часами, потом по приоритету
        plan_rows.sort(key=lambda r: (-(r['pairs_per_week'] or 0), r.get('priority', 2)))

        for plan_row in plan_rows:
            subject_id = plan_row['subject_id']
            teacher_id = plan_row.get('teacher_id')
            pairs_to_place = plan_row['pairs_per_week'] or 0

            if pairs_to_place <= 0:
                continue

            placed_for_subject = 0

            for day in DAYS_OF_WEEK:
                if placed_for_subject >= pairs_to_place:
                    break

                # Считаем, сколько пар уже стоит у группы в этот день
                group_day_pairs = sum(
                    1 for p in placed
                    if p['group_id'] == group_id and p['day_of_week'] == day
                )
                if group_day_pairs >= MAX_PAIRS_PER_DAY:
                    continue

                for para in PARAS:
                    if placed_for_subject >= pairs_to_place:
                        break

                    key = (day, para)

                    # Проверяем занятость группы
                    if key in busy['group'].get(group_id, set()):
                        continue
                    if any(
                        p['group_id'] == group_id and p['day_of_week'] == day
                        and p['para_number'] == para
                        for p in placed
                    ):
                        continue

                    # Проверяем занятость преподавателя
                    if teacher_id:
                        if key in busy['teacher'].get(teacher_id, set()):
                            continue
                        if any(
                            p.get('teacher_id') == teacher_id and p['day_of_week'] == day
                            and p['para_number'] == para
                            for p in placed
                        ):
                            continue

                        # Проверяем, что у преподавателя не больше 6 пар в этот день
                        teacher_day_pairs = sum(
                            1 for p in placed
                            if p.get('teacher_id') == teacher_id and p['day_of_week'] == day
                        )
                        if teacher_day_pairs >= MAX_PAIRS_PER_DAY:
                            continue

                    # Слот свободен — ставим
                    placed.append({
                        'group_id': group_id,
                        'group_name': group['name'],
                        'course': course,
                        'subject_id': subject_id,
                        'subject_name': plan_row.get('subject_name', ''),
                        'teacher_id': teacher_id,
                        'teacher_name': plan_row.get('teacher_name', ''),
                        'teacher_short': plan_row.get('teacher_short', ''),
                        'day_of_week': day,
                        'para_number': para,
                        'hours': 2,
                        'source': 'auto',
                    })

                    # Помечаем слот как занятый
                    busy['group'].setdefault(group_id, set()).add(key)
                    if teacher_id:
                        busy['teacher'].setdefault(teacher_id, set()).add(key)

                    placed_for_subject += 1
                    placed_count += 1

            if placed_for_subject < pairs_to_place:
                conflicts.append({
                    'group_name': group['name'],
                    'subject_name': plan_row.get('subject_name', ''),
                    'teacher_name': plan_row.get('teacher_name', ''),
                    'placed': placed_for_subject,
                    'wanted': pairs_to_place,
                })

        group_stats.append({
            'group_id': group_id,
            'group_name': group['name'],
            'course': course,
            'placed': placed_count,
            'wanted': wanted_total,
        })

    return {
        'academic_year': f"{academic_year}/{academic_year + 1}",
        'semester_label': semester_label,
        'semester_in_course': semester_in_course,
        'placed': placed,
        'conflicts': conflicts,
        'group_stats': group_stats,
        'stats': {
            'total_groups': len(groups),
            'total_placed': len(placed),
            'total_conflicts': len(conflicts),
        },
    }


def apply_generation(placed: List[Dict[str, Any]],
                     created_by: Optional[int] = None) -> Dict[str, int]:
    """Применяет сгенерированные пары: пишет в schedule с source='auto'."""
    if not placed:
        return {'inserted': 0, 'skipped': 0}

    academic_year, semester_in_course = get_current_semester_info()

    inserted = 0
    skipped = 0

    with db.connect() as conn:
        cursor = conn.cursor()
        for item in placed:
            gid = item.get('group_id')
            sid = item.get('subject_id')
            tid = item.get('teacher_id')
            day = item.get('day_of_week')
            para = item.get('para_number')
            hours = item.get('hours', 2)

            if not (gid and sid and day and para):
                skipped += 1
                continue

            # Проверяем, не занят ли слот
            cursor.execute('''
                SELECT id FROM schedule
                WHERE group_id = ? AND day_of_week = ? AND para_number = ? AND is_active = 1
            ''', (gid, day, para))
            if cursor.fetchone():
                skipped += 1
                continue

            if tid:
                cursor.execute('''
                    SELECT id FROM schedule
                    WHERE teacher_id = ? AND day_of_week = ? AND para_number = ? AND is_active = 1
                ''', (tid, day, para))
                if cursor.fetchone():
                    skipped += 1
                    continue

            cursor.execute('''
                INSERT INTO schedule
                (group_id, day_of_week, para_number, subject_id, teacher_id, room_id,
                 hours, created_by, source, semester_num, academic_year, is_active)
                VALUES (?, ?, ?, ?, ?, NULL, ?, ?, 'auto', ?, ?, 1)
            ''', (gid, day, para, sid, tid, hours, created_by,
                  semester_in_course, f"{academic_year}/{academic_year + 1}"))
            inserted += 1

    return {'inserted': inserted, 'skipped': skipped}