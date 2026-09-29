import sqlite3
import os
import glob

OLD_PROJECT = r"Z:\Техотдел\!учебчасть\u4ebkaWeb старый"


def check(path):
    print(f"\n{'=' * 70}")
    print(f"📁 {path}")
    print(f"   Размер: {os.path.getsize(path) / 1024:.1f} KB")
    print(f"{'=' * 70}")

    try:
        conn = sqlite3.connect(path)
        cur = conn.cursor()

        tables = [r[0] for r in cur.execute(
            "SELECT name FROM sqlite_master WHERE type='table'").fetchall()]

        if 'rooms' not in tables:
            print("❌ Нет таблицы rooms")
            conn.close()
            return

        for label, q in [
            ("🏫 Кабинетов", "SELECT COUNT(*) FROM rooms"),
            ("📚 Предметов", "SELECT COUNT(*) FROM subjects"),
            ("📅 Всего пар", "SELECT COUNT(*) FROM schedule"),
            ("📅 Активных", "SELECT COUNT(*) FROM schedule WHERE is_active=1"),
            ("🔗 С room_id", "SELECT COUNT(*) FROM schedule WHERE room_id IS NOT NULL"),
            ("👨‍🏫 Учителей", "SELECT COUNT(*) FROM teachers"),
            ("👥 Групп", "SELECT COUNT(*) FROM groups"),
        ]:
            try:
                cur.execute(q)
                print(f"{label}: {cur.fetchone()[0]}")
            except Exception as e:
                print(f"❌ {label}: {e}")

        # Битые ссылки
        try:
            cur.execute("""
                SELECT COUNT(*) FROM schedule s
                LEFT JOIN rooms r ON s.room_id = r.id
                WHERE s.room_id IS NOT NULL AND r.id IS NULL
            """)
            print(f"💔 Битых ссылок: {cur.fetchone()[0]}")
        except Exception as e:
            print(f"❌ Битые ссылки: {e}")

        # Примеры пар с кабинетами
        try:
            cur.execute("""
                SELECT s.id, s.room_id, r.number, r.building
                FROM schedule s
                JOIN rooms r ON s.room_id = r.id
                WHERE s.is_active=1
                LIMIT 5
            """)
            rows = cur.fetchall()
            if rows:
                print("\n✅ Примеры привязок:")
                for row in rows:
                    print(f"   пара #{row[0]}: room_id={row[1]} → каб. №{row[2]} (корп. {row[3]})")
        except Exception as e:
            print(f"❌ Примеры: {e}")

        conn.close()
    except Exception as e:
        print(f"❌ Не удалось открыть: {e}")


if __name__ == "__main__":
    # Все .db в корне старого проекта
    root_dbs = glob.glob(os.path.join(OLD_PROJECT, "*.db"))
    # Все .db в backups/
    backup_dbs = glob.glob(os.path.join(OLD_PROJECT, "backups", "*.db"))

    all_dbs = root_dbs + backup_dbs
    print(f"Найдено .db файлов: {len(all_dbs)}\n")

    for db in all_dbs:
        check(db)