from database import db

print("=" * 70)
print("🔍 ПРОВЕРКА ДУБЛИКАТОВ (группа 28)")
print("=" * 70)

group = db.get_group_by_name("28")
rows = db.get_curriculum_plan_v2(group['id'])

# Группируем по subject_id
by_subject = {}
for r in rows:
    sid = r['subject_id']
    if sid not in by_subject:
        by_subject[sid] = []
    by_subject[sid].append(r)

# Ищем дубли
dupes = {k: v for k, v in by_subject.items() if len(v) > 1}

print(f"Всего записей: {len(rows)}")
print(f"Уникальных предметов: {len(by_subject)}")
print(f"С дублями (>1): {len(dupes)}")
print()

if dupes:
    print("Примеры дублей:")
    for sid, items in list(dupes.items())[:10]:
        name = items[0].get('subject_name', '?')
        sems = [(i['course'], i['semester_in_course']) for i in items]
        print(f"  ID {sid}: {name[:50]:50s} | семестры: {sems}")
