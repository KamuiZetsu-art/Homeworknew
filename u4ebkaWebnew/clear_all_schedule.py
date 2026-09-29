import sqlite3
from pathlib import Path

DB = Path('schedule.db')
BACKUP = Path('backups') / 'schedule_before_clear.db'

# Сначала бэкап
import shutil
BACKUP.parent.mkdir(exist_ok=True)
shutil.copy(DB, BACKUP)
print(f"💾 Бэкап: {BACKUP}")

conn = sqlite3.connect(str(DB))
cur = conn.cursor()

cur.execute("SELECT COUNT(*) FROM schedule WHERE is_active=1")
total = cur.fetchone()[0]

cur.execute("UPDATE schedule SET is_active=0 WHERE is_active=1")
affected = cur.rowcount
conn.commit()

print(f"✅ Было активных: {total}")
print(f"✅ Деактивировано: {affected}")

cur.execute("SELECT COUNT(*) FROM schedule WHERE is_active=1")
print(f"✅ Осталось активных: {cur.fetchone()[0]}")

conn.close()