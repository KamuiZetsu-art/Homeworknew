from database import db

g = db.get_group_by_name('43н')
if not g:
    print("❌ Группа 43н не найдена")
    exit(1)

rows = db.get_curriculum_plan_v2(g['id'])

print(f"Всего записей: {len(rows)}")
print()
print(f"{'Сем':>4} {'Дисциплина':<45} {'contact':>8} {'pairs':>6} {'hours':>6}")
print("-" * 75)

for r in rows[:20]:
    name = (r.get('subject_name') or '')[:43]
    print(f"{r['global_semester']:>4} {name:<45} "
          f"{r['contact']:>8} {r['pairs_per_week']:>6} {r['total_hours']:>6}")