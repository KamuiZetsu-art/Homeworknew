import sqlite3

conn = sqlite3.connect('schedule.db')
conn.row_factory = sqlite3.Row
cur = conn.cursor()

# 1. Сколько пар в schedule для группы 28
cur.execute("SELECT id FROM groups WHERE name='28'")
row = cur.fetchone()
if not row:
    print("❌ Группа 28 не найдена")
    conn.close()
    exit()

gid = row['id']
print(f"Группа 28: group_id = {gid}")

cur.execute("SELECT COUNT(*) FROM schedule WHERE group_id=? AND is_active=1", (gid,))
print(f"📅 Пар в schedule для 28: {cur.fetchone()[0]}")

# 2. Что в curriculum_plan_v2
cur.execute('''
    SELECT cp.subject_id, cp.teacher_id, cp.pairs_per_week,
           s.name AS subject_name, t.short_name AS teacher_short
    FROM curriculum_plan_v2 cp
    LEFT JOIN subjects s ON cp.subject_id = s.id
    LEFT JOIN teachers t ON cp.teacher_id = t.id
    WHERE cp.group_id=? AND cp.pairs_per_week > 0
    ORDER BY s.name
''', (gid,))

rows = cur.fetchall()
print(f"\n📋 Дисциплин в УП с pairs_per_week > 0: {len(rows)}")
print(f"{'Предмет':<40} {'Преподаватель':<25} {'Пар/нед':>7}")
print("-" * 75)
for r in rows:
    teacher = r['teacher_short'] or '❌ НЕ НАЗНАЧЕН'
    print(f"{r['subject_name'][:40]:<40} {teacher[:25]:<25} {r['pairs_per_week']:>7}")

# 3. Занятость по дням для группы 28
print("\n📅 Занятость группы 28 по дням:")
cur.execute('''
    SELECT day_of_week, COUNT(*) AS cnt
    FROM schedule
    WHERE group_id=? AND is_active=1
    GROUP BY day_of_week
    ORDER BY day_of_week
''', (gid,))
for r in cur.fetchall():
    print(f"  {r['day_of_week']}: {r['cnt']} пар")

# 4. Проверка: у каких дисциплин нет преподавателя
print("\n⚠️  Дисциплины без преподавателя:")
cur.execute('''
    SELECT s.name AS subject_name, cp.pairs_per_week
    FROM curriculum_plan_v2 cp
    LEFT JOIN subjects s ON cp.subject_id = s.id
    WHERE cp.group_id=? AND cp.pairs_per_week > 0 AND cp.teacher_id IS NULL
''', (gid,))
empty = cur.fetchall()
if not empty:
    print("  ✅ Все дисциплины имеют преподавателя")
else:
    for r in empty:
        print(f"  ❌ {r['subject_name']} ({r['pairs_per_week']} пар/нед)")

conn.close()