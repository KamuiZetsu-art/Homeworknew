"""
Миграция БД v2: курсы, семестры, учебный план по семестрам.
Запускается ОДИН РАЗ.
"""
import sqlite3
import sys
from pathlib import Path

sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

DB_FILE = Path(__file__).resolve().parent.parent / "schedule.db"


def column_exists(cursor, table, column):
    cursor.execute(f"PRAGMA table_info({table})")
    return column in [r[1] for r in cursor.fetchall()]


def table_exists(cursor, table):
    cursor.execute(
        "SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=?",
        (table,)
    )
    return cursor.fetchone()[0] > 0


def migrate():
    if not DB_FILE.exists():
        print(f"❌ Файл {DB_FILE} не найден!")
        return False

    print("=" * 60)
    print("🔧 МИГРАЦИЯ БД v2")
    print("=" * 60)

    conn = sqlite3.connect(str(DB_FILE))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    try:
        # ========== 1. groups ==========
        print("\n📋 Таблица groups:")
        groups_new_cols = [
            ("specialty_code", "TEXT"),
            ("specialty_name", "TEXT"),
            ("current_course", "INTEGER DEFAULT 1"),
            ("admission_year", "INTEGER"),
            ("education_form", "TEXT DEFAULT 'Очная'"),
            ("department", "TEXT"),
            ("curator", "TEXT"),
            ("students_count", "INTEGER DEFAULT 0"),
            ("is_active", "INTEGER DEFAULT 1"),
        ]
        for col, typ in groups_new_cols:
            if not column_exists(cursor, "groups", col):
                cursor.execute(f"ALTER TABLE groups ADD COLUMN {col} {typ}")
                print(f"   ✅ + {col} {typ}")
            else:
                print(f"   ⏭  {col} уже есть")

        # ========== 2. schedule ==========
        print("\n📋 Таблица schedule:")
        schedule_new_cols = [
            ("source", "TEXT DEFAULT 'manual'"),
            ("semester_num", "INTEGER"),
            ("academic_year", "TEXT"),
        ]
        for col, typ in schedule_new_cols:
            if not column_exists(cursor, "schedule", col):
                cursor.execute(f"ALTER TABLE schedule ADD COLUMN {col} {typ}")
                print(f"   ✅ + {col} {typ}")
            else:
                print(f"   ⏭  {col} уже есть")

        # ========== 3. curriculum_plan_v2 ==========
        print("\n📋 Таблица curriculum_plan_v2:")
        if not table_exists(cursor, "curriculum_plan_v2"):
            cursor.execute('''
                CREATE TABLE curriculum_plan_v2 (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    group_id INTEGER NOT NULL,
                    subject_id INTEGER NOT NULL,
                    teacher_id INTEGER,
                    course INTEGER NOT NULL,
                    semester_in_course INTEGER NOT NULL,
                    global_semester INTEGER NOT NULL,
                    total_hours INTEGER DEFAULT 0,
                    lectures INTEGER DEFAULT 0,
                    lab INTEGER DEFAULT 0,
                    practical INTEGER DEFAULT 0,
                    seminars INTEGER DEFAULT 0,
                    individual INTEGER DEFAULT 0,
                    group_lessons INTEGER DEFAULT 0,
                    subgroup_lessons INTEGER DEFAULT 0,
                    self_work INTEGER DEFAULT 0,
                    control INTEGER DEFAULT 0,
                    independent INTEGER DEFAULT 0,
                    contact INTEGER DEFAULT 0,
                    pairs_per_week INTEGER DEFAULT 0,
                    weeks INTEGER DEFAULT 16,
                    index_code TEXT,
                    registry_number TEXT,
                    control_forms TEXT,
                    block TEXT,
                    priority INTEGER DEFAULT 2,
                    comment TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    created_by INTEGER,
                    FOREIGN KEY (group_id) REFERENCES groups(id),
                    FOREIGN KEY (subject_id) REFERENCES subjects(id),
                    FOREIGN KEY (teacher_id) REFERENCES teachers(id),
                    UNIQUE(group_id, subject_id, course, semester_in_course)
                )
            ''')
            print("   ✅ Таблица создана")
        else:
            print("   ⏭  Уже существует")

        # ========== 4. Индексы ==========
        print("\n📋 Индексы:")
        indexes = [
            ("idx_cpv2_group", "curriculum_plan_v2(group_id, course, semester_in_course)"),
            ("idx_cpv2_subject", "curriculum_plan_v2(subject_id)"),
            ("idx_cpv2_teacher", "curriculum_plan_v2(teacher_id)"),
            ("idx_sched_source", "schedule(source, is_active)"),
            ("idx_groups_active_course", "groups(is_active, current_course)"),
        ]
        for idx_name, idx_def in indexes:
            cursor.execute(f"CREATE INDEX IF NOT EXISTS {idx_name} ON {idx_def}")
            print(f"   ✅ {idx_name}")

        conn.commit()
        print("\n" + "=" * 60)
        print("✅ МИГРАЦИЯ v2 ЗАВЕРШЕНА")
        print("=" * 60)
        return True

    except Exception as e:
        print(f"\n❌ ОШИБКА: {e}")
        conn.rollback()
        return False
    finally:
        conn.close()


def verify():
    print("\n🔍 ПРОВЕРКА:")
    conn = sqlite3.connect(str(DB_FILE))
    conn.row_factory = sqlite3.Row
    cursor = conn.cursor()

    cursor.execute("PRAGMA table_info(groups)")
    cols = [r[1] for r in cursor.fetchall()]
    for c in ["specialty_code", "current_course", "admission_year",
              "education_form", "is_active"]:
        status = "✅" if c in cols else "❌"
        print(f"   {status} groups.{c}")

    cursor.execute("PRAGMA table_info(schedule)")
    cols = [r[1] for r in cursor.fetchall()]
    for c in ["source", "semester_num", "academic_year"]:
        status = "✅" if c in cols else "❌"
        print(f"   {status} schedule.{c}")

    cursor.execute(
        "SELECT name FROM sqlite_master WHERE type='table' AND name='curriculum_plan_v2'"
    )
    status = "✅" if cursor.fetchone() else "❌"
    print(f"   {status} table curriculum_plan_v2")

    conn.close()


if __name__ == "__main__":
    if migrate():
        verify()