"""
Парсер учебного плана УП XX.xls (v2) — с динамическим определением колонок.
"""
import sys
import argparse
from pathlib import Path

BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from database import db
from config import config
from logger import get_logger

logger = get_logger(__name__)


# Маппинг русских названий метрик → английские ключи
METRIC_MAP = {
    'Труд-сть': 'total_hours',
    'Лекции': 'lectures',
    'Лаб.зан-я': 'lab',
    'Практ.зан-я': 'practical',
    'Семинары': 'seminars',
    'Инд. занят.': 'individual',
    'Занятия в группах': 'group_lessons',
    'Занятия в подгруппах': 'subgroup_lessons',
    'Инд. проект': 'individual_proj',
    'Сам. раб.': 'self_work',
    'Консультации': 'consultations',
    'Контроль': 'control',
    'Самост.': 'independent',
    'Контактн.': 'contact',
}


def col_letter(idx):
    s = ""
    idx += 1
    while idx:
        idx, rem = divmod(idx - 1, 26)
        s = chr(65 + rem) + s
    return s


def read_excel(path):
    ext = path.suffix.lower()
    if ext == '.xls':
        import xlrd
        wb = xlrd.open_workbook(str(path))
        ws = wb.sheet_by_index(0)
        rows = [[ws.cell_value(r, c) for c in range(ws.ncols)] for r in range(ws.nrows)]
    elif ext == '.xlsx':
        import openpyxl
        wb = openpyxl.load_workbook(path, data_only=True)
        ws = wb.active
        rows = [list(r) for r in ws.iter_rows(values_only=True)]
    else:
        raise ValueError(f"Неподдерживаемый формат: {ext}")
    return rows


def to_num(v):
    if v is None:
        return 0
    if isinstance(v, (int, float)):
        return int(v)
    s = str(v).strip()
    if not s:
        return 0
    if '(' in s:
        s = s.split('(')[0].strip()
    try:
        return int(float(s))
    except (ValueError, TypeError):
        return 0


def find_header_row(rows):
    for i, row in enumerate(rows):
        if len(row) > 4 and row[4] and 'Название' in str(row[4]):
            return i
    return None


def find_semester_start(rows):
    for r_idx, row in enumerate(rows):
        for c_idx, cell in enumerate(row):
            if cell and '1 курс' in str(cell):
                if r_idx + 1 < len(rows):
                    row2 = rows[r_idx + 1]
                    for c2 in range(c_idx, len(row2)):
                        v = row2[c2]
                        if v and '1 (' in str(v):
                            return r_idx, c2
    return None, None


def detect_semester_columns(rows, hdr_r, sem_start_col):
    """
    Определяет, сколько колонок в одном семестре и как называются метрики.
    """
    keys_row = None
    for r_idx in range(hdr_r, min(hdr_r + 8, len(rows))):
        row = rows[r_idx]
        trud_count = sum(1 for v in row if v and str(v).strip() == 'Труд-сть')
        if trud_count >= 2:
            keys_row = row
            break

    if keys_row is None:
        # Fallback
        return 14, ['total_hours', 'lectures', 'lab', 'practical', 'seminars',
                    'individual', 'group_lessons', 'subgroup_lessons', 'individual_proj',
                    'self_work', 'consultations', 'control', 'independent', 'contact']

    first = None
    second = None
    for c in range(sem_start_col, len(keys_row)):
        v = keys_row[c]
        if v and str(v).strip() == 'Труд-сть':
            if first is None:
                first = c
            elif second is None:
                second = c
                break

    if first is None or second is None:
        return 14, ['total_hours', 'lectures', 'lab', 'practical', 'seminars',
                    'individual', 'group_lessons', 'subgroup_lessons', 'individual_proj',
                    'self_work', 'consultations', 'control', 'independent', 'contact']

    cols_per_semester = second - first

    keys = []
    for i in range(cols_per_semester):
        v = keys_row[first + i] if first + i < len(keys_row) else None
        name = str(v).strip() if v else f"col_{i}"
        keys.append(METRIC_MAP.get(name, name))

    return cols_per_semester, keys


def parse_curriculum(path, verbose=True):
    rows = read_excel(path)
    if verbose:
        print(f"📄 Файл: {path.name}")
        print(f"📏 Строк: {len(rows)}, Колонок: {len(rows[0]) if rows else 0}")

    direction = ""
    education_form = ""
    for r in rows[:10]:
        for c, v in enumerate(r):
            if v and 'Направление подготовки' in str(v):
                if c + 1 < len(r) and r[c+1]:
                    direction = str(r[c+1]).strip()
            if v and 'Форма обучения' in str(v):
                if c + 1 < len(r) and r[c+1]:
                    education_form = str(r[c+1]).strip()

    if verbose:
        print(f"🎓 Направление: {direction[:80]}")

    sem_hdr_r, sem_start_col = find_semester_start(rows)
    if sem_start_col is None:
        if verbose: print("❌ Не найден блок «Нагрузка по семестрам»")
        return None

    hdr_r = find_header_row(rows)
    if hdr_r is None:
        if verbose: print("❌ Не найдена шапка")
        return None

    # Динамическое определение колонок семестров
    cols_per_semester, semester_keys = detect_semester_columns(rows, hdr_r, sem_start_col)

    if verbose:
        print(f"🎯 Блок семестров: строка {sem_hdr_r + 1}, колонка {col_letter(sem_start_col)}")
        print(f"📊 Колонок в семестре: {cols_per_semester}")
        print(f"📊 Метрики: {semester_keys}")

    # Ключевые колонки шапки
    col_index_code = 1
    col_name = 4
    col_type = 20
    col_registry = 21
    col_total = 34

    hdr = rows[hdr_r]
    for c, v in enumerate(hdr):
        if v:
            s = str(v).strip()
            if s == 'Индекс': col_index_code = c
            if s == 'Название': col_name = c
            if s == 'Тип': col_type = c
            if s == '№ (реестр)': col_registry = c
            if s == 'Труд-сть (всего)': col_total = c

    disciplines = []
    for r_idx in range(hdr_r + 1, len(rows)):
        row = rows[r_idx]
        if not row or len(row) < 5:
            continue

        code = str(row[col_index_code]).strip() if col_index_code < len(row) and row[col_index_code] else ""
        name = str(row[col_name]).strip() if col_name < len(row) and row[col_name] else ""
        typ = str(row[col_type]).strip() if col_type < len(row) and row[col_type] else ""

        if not code or not name:
            continue
        if typ in ('Цикл', 'Часть'):
            continue
        if code in ('ПА', 'ГИА', 'ДП', 'ВКР'):
            continue

        semesters = {}
        for sem_num in range(1, 9):
            offset = (sem_num - 1) * cols_per_semester
            sem_col_start = sem_start_col + offset
            if sem_col_start + cols_per_semester - 1 >= len(row):
                break

            total = to_num(row[sem_col_start])
            if total == 0:
                continue

            sem_data = {}
            for i, key in enumerate(semester_keys):
                v = to_num(row[sem_col_start + i]) if sem_col_start + i < len(row) else 0
                sem_data[key] = v
            semesters[sem_num] = sem_data

        if not semesters:
            continue

        total_hours = to_num(row[col_total]) if col_total < len(row) else 0
        registry = str(row[col_registry]).strip() if col_registry < len(row) and row[col_registry] else ""

        disciplines.append({
            'index': code,
            'name': name,
            'type': typ,
            'registry': registry,
            'total_hours': total_hours,
            'semesters': semesters,
        })

    if verbose:
        print(f"\n📚 Найдено дисциплин: {len(disciplines)}")

    return {
        'direction': direction,
        'education_form': education_form,
        'disciplines': disciplines,
    }


def save_to_db(parsed, group_name):
    group = db.get_group_by_name(group_name)
    if not group:
        print(f"❌ Группа '{group_name}' не найдена")
        return 0

    group_id = group['id']
    deleted = db.delete_curriculum_plan_v2(group_id)
    if deleted:
        print(f"🗑 Удалён старый план: {deleted} строк")

    WEEKS = {1: 17, 2: 23, 3: 17, 4: 19, 5: 14, 6: 11, 7: 14, 8: 11}

    created = 0
    for d in parsed['disciplines']:
        subject = db.get_subject_by_name(d['name'])
        if subject:
            subject_id = subject['id']
        else:
            subject_id = db.add_subject(d['name'])
            if not subject_id:
                continue

        for sem_num, sem in d['semesters'].items():
            course = (sem_num + 1) // 2
            semester_in_course = 2 if sem_num % 2 == 0 else 1
            total_hours = sem.get('total_hours', 0)
            weeks = WEEKS.get(sem_num, 16)
            # Формула: Труд-сть / недели / 2 часа на пару
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
                contact=contact,
                pairs_per_week=pairs,
                weeks=weeks,
                index_code=d['index'],
                registry_number=d.get('registry', ''),
                priority=2,
            )
            if ok:
                created += 1

    print(f"\n✅ Сохранено: {created} записей")
    return created


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("file", help="Путь к файлу УП")
    parser.add_argument("--group", required=True, help="Имя группы")
    args = parser.parse_args()

    path = Path(args.file)
    if not path.exists():
        print(f"❌ Файл не найден: {path}")
        return

    print("=" * 70)
    print("📥 ИМПОРТ УЧЕБНОГО ПЛАНА v2")
    print("=" * 70)

    parsed = parse_curriculum(path)
    if not parsed:
        return

    print("\n" + "=" * 70)
    print("💾 СОХРАНЕНИЕ В БД")
    print("=" * 70)
    save_to_db(parsed, args.group)


if __name__ == "__main__":
    main()