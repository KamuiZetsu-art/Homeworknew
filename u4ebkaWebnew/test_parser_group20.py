from parse_curriculum_v2 import parse_curriculum
from pathlib import Path
import sys

if len(sys.argv) < 2:
    print("Использование: python test_parser_group20.py 'УП 20 гр.xls'")
    sys.exit(1)

p = parse_curriculum(Path(sys.argv[1]))
if not p:
    print("❌ Не удалось распарсить")
    sys.exit(1)

print(f"\n📚 Всего дисциплин: {len(p['disciplines'])}")

# Ищем Математику
for d in p['disciplines']:
    if 'Математика' in d['name']:
        print(f"\n📋 {d['name']}")
        print(f"   total_hours = {d['total_hours']}")
        for sem_num, sem in d['semesters'].items():
            print(f"   Семестр {sem_num}:")
            for k, v in sem.items():
                if v != 0:
                    print(f"      {k} = {v}")
        break

# Ищем Музыкальную литературу
for d in p['disciplines']:
    if 'Музыкальная литература' in d['name']:
        print(f"\n📋 {d['name']}")
        print(f"   total_hours = {d['total_hours']}")
        for sem_num, sem in d['semesters'].items():
            print(f"   Семестр {sem_num}:")
            for k, v in sem.items():
                if v != 0:
                    print(f"      {k} = {v}")
        break