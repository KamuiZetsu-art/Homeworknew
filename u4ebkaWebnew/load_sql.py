import sqlite3
import os

SQL_FILE = "nokki_data.sql"
DB_FILE = "schedule.db"


def load_sql():
    if not os.path.exists(SQL_FILE):
        print(f"❌ Файл {SQL_FILE} не найден")
        return

    if not os.path.exists(DB_FILE):
        print(f"⚠️ Файл {DB_FILE} не найден. Запустите сначала app.py, чтобы создать базу.")
        return

    print(f"📂 Загрузка данных из {SQL_FILE}...")

    conn = sqlite3.connect(DB_FILE)

    try:
        with open(SQL_FILE, 'r', encoding='utf-8') as f:
            sql = f.read()

        conn.executescript(sql)
        conn.commit()

        # Проверяем результат
        cursor = conn.cursor()

        cursor.execute("SELECT COUNT(*) FROM groups")
        groups = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM teachers")
        teachers = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM subjects")
        subjects = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM rooms")
        rooms = cursor.fetchone()[0]

        cursor.execute("SELECT COUNT(*) FROM schedule WHERE is_active=1")
        lessons = cursor.fetchone()[0]

        print()
        print("=" * 50)
        print("✅ ДАННЫЕ УСПЕШНО ЗАГРУЖЕНЫ!")
        print("=" * 50)
        print(f"📚 Групп: {groups}")
        print(f"👨‍🏫 Преподавателей: {teachers}")
        print(f"📖 Предметов: {subjects}")
        print(f"🚪 Кабинетов: {rooms}")
        print(f"📅 Занятий: {lessons}")
        print("=" * 50)

    except Exception as e:
        print(f"❌ Ошибка: {e}")
        conn.rollback()
    finally:
        conn.close()


if __name__ == '__main__':
    load_sql()
    input("\nНажмите Enter для выхода...")