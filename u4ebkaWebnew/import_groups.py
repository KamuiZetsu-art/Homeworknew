"""
Импорт групп из Excel-файла колледжа.
Файл: 1790577215771.xlsx (лист "Группы")
"""
import sys
from pathlib import Path

# Путь к корню проекта
BASE_DIR = Path(__file__).resolve().parent
sys.path.insert(0, str(BASE_DIR))

from database import db
from config import config
from logger import get_logger

logger = get_logger(__name__)

GROUPS_FILE = BASE_DIR / "1790577215771.xlsx"
CURRENT_ACADEMIC_YEAR = config.CURRENT_ACADEMIC_YEAR  # 2025


def import_groups():
    if not GROUPS_FILE.exists():
        print(f"❌ Файл не найден: {GROUPS_FILE}")
        return

    try:
        import openpyxl
    except ImportError:
        print("❌ Установи openpyxl: pip install openpyxl")
        return

    print("=" * 70)
    print("📥 ИМПОРТ ГРУПП ИЗ EXCEL")
    print("=" * 70)
    print(f"📄 Файл: {GROUPS_FILE.name}")
    print(f"📅 Учебный год: {CURRENT_ACADEMIC_YEAR}/{CURRENT_ACADEMIC_YEAR + 1}")
    print()

    wb = openpyxl.load_workbook(GROUPS_FILE, data_only=True)
    ws = wb.active

    # Читаем все строки
    rows = list(ws.iter_rows(values_only=True))

    # Ищем строку-заголовок
    header_row_idx = None
    for i, row in enumerate(rows):
        if row and "Название" in [str(c).strip() if c else "" for c in row]:
            header_row_idx = i
            break

    if header_row_idx is None:
        print("❌ Не найдена строка с заголовками (ищу 'Название')")
        return

    header = [str(c).strip() if c else "" for c in rows[header_row_idx]]
    col = {name: idx for idx, name in enumerate(header) if name}

    print(f"📋 Найдено колонок: {len(col)}")
    print(f"   Ключевые: {', '.join(list(col.keys())[:5])}...")
    print()

    def get(row, name):
        idx = col.get(name)
        if idx is None or idx >= len(row):
            return None
        return row[idx]

    # Счётчики
    created = 0
    updated = 0
    skipped = 0
    full_time = 0      # очных
    part_time = 0      # заочных

    for row_num, row in enumerate(rows[header_row_idx + 1:], start=header_row_idx + 2):
        if not row or not row[0]:
            continue

        name = str(get(row, "Название") or "").strip()
        if not name or name in ("None", ""):
            continue

        # Пропускаем строки-заголовки повторные
        if name == "Название":
            continue

        try:
            course = int(float(get(row, "Курс") or 0))
        except (ValueError, TypeError):
            skipped += 1
            continue

        if course < 1 or course > 4:
            skipped += 1
            continue

        # Специальность: "53.02.05 Сольное и хоровое народное пение (специальность СПО, СПО 2022)"
        specialty_full = str(get(row, "Направление подготовки (специальность)") or "").strip()
        specialty_code = ""
        specialty_name = ""
        if specialty_full:
            parts = specialty_full.split(" ", 1)
            if len(parts) >= 2 and "." in parts[0]:
                specialty_code = parts[0]  # "53.02.05"
                rest = parts[1]
                # Убираем хвост "(специальность СПО, СПО 2022)"
                if "(" in rest:
                    rest = rest.split("(")[0]
                specialty_name = rest.strip()

        education_form = str(get(row, "Форма освоения") or "Очная").strip()
        department = str(get(row, "Формирующее подр.") or "").strip()
        curator = str(get(row, "Куратор") or "").strip()

        try:
            students_count = int(float(get(row, "Активных обучающихся") or 0))
        except (ValueError, TypeError):
            students_count = 0

        # Год поступления: вычисляем
        # Учебный год 2025 = "2025/2026", 1 курс поступил в 2025
        # Если группа на курсе N в 2025/2026 → admission_year = 2025 - (N - 1)
        admission_year = CURRENT_ACADEMIC_YEAR - (course - 1)

        # Очная = активная для расписания, заочная = неактивная
        is_active = 1 if education_form == "Очная" else 0

        if is_active:
            full_time += 1
        else:
            part_time += 1

        # Проверяем, есть ли уже такая группа
        existing = db.get_group_by_name(name)

        # Upsert
        group_id = db.upsert_group(
            name=name,
            specialty_code=specialty_code,
            specialty_name=specialty_name,
            current_course=course,
            admission_year=admission_year,
            education_form=education_form,
            department=department,
            curator=curator,
            students_count=students_count,
            is_active=is_active,
        )

        if group_id:
            if existing:
                updated += 1
                mark = "🔄"
            else:
                created += 1
                mark = "✅"
            print(f"   {mark} {name:35s} | курс {course} | {education_form:8s} | {specialty_code}")
        else:
            skipped += 1
            print(f"   ❌ {name} — ошибка")

    print()
    print("=" * 70)
    print("📊 ИТОГО")
    print("=" * 70)
    print(f"   ✅ Создано:     {created}")
    print(f"   🔄 Обновлено:   {updated}")
    print(f"   ⏭  Пропущено:   {skipped}")
    print()
    print(f"   📚 Всего очных:     {full_time}")
    print(f"   📚 Всего заочных:   {part_time}")
    print("=" * 70)


def verify():
    """Проверка после импорта."""
    print("\n🔍 ПРОВЕРКА БД:")
    groups = db.get_active_groups_for_schedule()
    print(f"   Активных очных групп: {len(groups)}")
    print()
    print("   Курс | Групп")
    print("   -----|------")
    for c in range(1, 5):
        cnt = sum(1 for g in groups if g['current_course'] == c)
        print(f"   {c}    | {cnt}")


if __name__ == "__main__":
    import_groups()
    verify()