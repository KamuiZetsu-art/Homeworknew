import sqlite3

conn = sqlite3.connect('schedule.db')
cur = conn.cursor()
cur.execute("UPDATE schedule SET is_active=0 WHERE source='auto'")
conn.commit()
print(f"✅ Деактивировано пар: {cur.rowcount}")
conn.close()