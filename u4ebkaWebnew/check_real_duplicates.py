from database import db

print("=" * 70)
print("🔍 НАСТОЯЩИЕ ДУБЛИКАТЫ (одинаковый предмет в одном семестре)")
print("=" * 70)

group = db.get_group_by_name("28")
rows = db.get_curriculum_plan_v2(group['id'])

seen = {}
for r in rows:
    key = (r['subject_id'], r['course'], r['semester_in_course'])
    if key not in seen:
        seen[key] = []
    seen[key].append(r)

dupes = {k: v for k, v in seen.items() if len(v) > 1}

print(f"Всего записей: {len(rows)}")
print(f"Уникальных (subject, course, semester): {len(seen)}")
print(f"Реальных дублей: {len(dupes)}")
print()

if dupes:
    print("Дубли:")
    for k, items in list(dupes.items())[:20]:
        name = items[0].get('subject_name', '?')
        print(f"  subject_id={k[0]}, course={k[1]}, sem={k[2]}: {name[:50]} — {len(items)} раз")

# Сводка по семестрам
print()
print("=" * 70)
print("📊 СВОДКА ПО СЕМЕСТРАМ")
print("=" * 70)
print(f"{'Курс':>4} {'Сем':>4} {'Дисциплин':>10} {'Часов':>8} {'Пар/нед':>8}")
print("-" * 45)

from collections import defaultdict
by_sem = defaultdict(lambda: {'count': 0, 'hours': 0, 'pairs': 0})
for r in rows:
    key = (r['course'], r['semester_in_course'])
    by_sem[key]['count'] += 1
    by_sem[key]['hours'] += r['total_hours']
    by_sem[key]['pairs'] += r['pairs_per_week']

for key in sorted(by_sem.keys()):
    c, s = key
    d = by_sem[key]
    print(f"{c:>4} {s:>4} {d['count']:>10} {d['hours']:>8} {d['pairs']:>8}")