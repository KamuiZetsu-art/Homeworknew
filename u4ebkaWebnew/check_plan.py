from database import db

print("=" * 70)
print("📊 УЧЕБНЫЕ ПЛАНЫ ПО ВСЕМ ГРУППАМ")
print("=" * 70)
print(f"{'Группа':<10} {'Курс':>5} {'Дисциплин':>10} {'Часов':>8} {'Пар/нед':>8}")
print("-" * 70)

groups = db.get_all_groups()
for g in groups:
    summary = db.get_curriculum_summary(g['id'])
    if summary['total_disciplines'] > 0:
        print(f"{g['name']:<10} {g.get('current_course', '?'):>5} "
              f"{summary['total_disciplines']:>10} "
              f"{summary['total_hours']:>8} "
              f"{sum(r['total_pairs_per_week'] or 0 for r in summary['by_semester']):>8}")