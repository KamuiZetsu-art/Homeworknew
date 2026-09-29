"""
Миграция v3: личные расписания (индивидуальные занятия).
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
    print("🔧 МИГРАЦИЯ БД v3 — Личные расписания")
    print("=" * 60)

    conn = sqlite3.connect(str(DB_FILE))
    cursor = conn.cursor()

    try:
        # ========== 1. teacher_personal_schedule ==========
        print("\n📋 Таблица teacher_personal_schedule:")
        if not table_exists(cursor, "teacher_personal_schedule"):
            cursor.execute('''
                CREATE TABLE teacher_personal_schedule (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    teacher_id INTEGER NOT NULL,
                    day_of_week TEXT NOT NULL,
                    para_number INTEGER NOT NULL,

                    student1_id INTEGER,
                    student2_id INTEGER,
                    students_raw TEXT,

                    group_id INTEGER,
                    subject_id INTEGER,
                    room_id INTEGER,

                    lesson_type TEXT DEFAULT 'individual',
                    semester_num INTEGER,
                    academic_year TEXT,
                    note TEXT,
                    is_active INTEGER DEFAULT 1,
                    created_by INTEGER,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,

                    FOREIGN KEY (teacher_id) REFERENCES teachers(id),
                    FOREIGN KEY (student1_id) REFERENCES users(id),
                    FOREIGN KEY (student2_id) REFERENCES users(id),
                    FOREIGN KEY (group_id) REFERENCES groups(id),
                    FOREIGN KEY (subject_id) REFERENCES subjects(id),
                    FOREIGN KEY (room_id) REFERENCES rooms(id),
                    FOREIGN KEY (created_by) REFERENCES users(id)
                )
            ''')
            print("   ✅ Таблица создана")
        else:
            print("   ⏭  Уже существует")

        # Индексы
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_tps_teacher ON teacher_personal_schedule(teacher_id, day_of_week, para_number, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_tps_group ON teacher_personal_schedule(group_id, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_tps_student1 ON teacher_personal_schedule(student1_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_tps_student2 ON teacher_personal_schedule(student2_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_tps_semester ON teacher_personal_schedule(semester_num, academic_year)")
        print("   ✅ Индексы добавлены")

        # ========== 2. schedule — lesson_type ==========
        print("\n📋 Таблица schedule:")
        if not column_exists(cursor, "schedule", "lesson_type"):
            cursor.execute("ALTER TABLE schedule ADD COLUMN lesson_type TEXT DEFAULT 'group'")
            print("   ✅ + lesson_type TEXT DEFAULT 'group'")
        else:
            print("   ⏭  lesson_type уже есть")

        # ========== 3. rooms — расширение ==========
        print("\n📋 Таблица rooms:")
        rooms_new_cols = [
            ("room_type", "TEXT"),
            ("subjects_allowed", "TEXT"),
            ("capacity", "INTEGER"),
        ]
        for col, typ in rooms_new_cols:
            if not column_exists(cursor, "rooms", col):
                cursor.execute(f"ALTER TABLE rooms ADD COLUMN {col} {typ}")
                print(f"   ✅ + {col} {typ}")
            else:
                print(f"   ⏭  {col} уже есть")

        conn.commit()
        print("\n" + "=" * 60)
        print("✅ МИГРАЦИЯ v3 ЗАВЕРШЕНА")
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

    cursor.execute("SELECT name FROM sqlite_master WHERE type='table' AND name='teacher_personal_schedule'")
    status = "✅" if cursor.fetchone() else "❌"
    print(f"   {status} table teacher_personal_schedule")

    cursor.execute("PRAGMA table_info(schedule)")
    cols = [r[1] for r in cursor.fetchall()]
    status = "✅" if "lesson_type" in cols else "❌"
    print(f"   {status} schedule.lesson_type")

    cursor.execute("PRAGMA table_info(rooms)")
    cols = [r[1] for r in cursor.fetchall()]
    for c in ["room_type", "subjects_allowed", "capacity"]:
        status = "✅" if c in cols else "❌"
        print(f"   {status} rooms.{c}")

    conn.close()


if __name__ == "__main__":
    if migrate():
        verify()