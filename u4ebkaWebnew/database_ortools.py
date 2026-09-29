import sqlite3
from datetime import datetime
from typing import Optional, List, Dict, Any

from logger import get_logger

logger = get_logger(__name__)

try:
    from ortools.sat.python import cp_model
    ORTOOLS_AVAILABLE = True
except ImportError:
    ORTOOLS_AVAILABLE = False
    logger.warning("OR-Tools не установлен — используется жадный алгоритм")


DAYS_OF_WEEK = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']
PARAS = [1, 2, 3, 4, 5, 6]


def is_available() -> bool:
    return ORTOOLS_AVAILABLE


def solve_schedule(db_file: str,
                    curriculum_rows: List[Dict[str, Any]],
                    existing_busy: Dict[str, Any],
                    max_pairs_per_day: int = 6,
                    time_limit_sec: int = 30) -> Dict[str, Any]:
    if not ORTOOLS_AVAILABLE:
        return {'error': 'OR-Tools недоступен'}

    if not curriculum_rows:
        return {
            'placed': [],
            'conflicts': [],
            'stats': {'total_requests': 0, 'total_placed': 0, 'total_conflicts': 0},
        }

    model = cp_model.CpModel()

    slots = [(day_idx, para) for day_idx in range(len(DAYS_OF_WEEK)) for para in PARAS]
    n_slots = len(slots)

    requests = []
    for r in curriculum_rows:
        ppr = r.get('pairs_per_week') or 0
        if ppr <= 0:
            continue
        requests.append({
            'row': r,
            'group_id': r['group_id'],
            'subject_id': r['subject_id'],
            'teacher_id': r.get('teacher_id'),
            'pairs_per_week': ppr,
            'priority': r.get('priority', 2),
            'group_name': r.get('group_name') or '',
            'subject_name': r.get('subject_name') or '',
            'teacher_name': r.get('teacher_name') or '',
            'teacher_short': r.get('teacher_short') or '',
        })

    if not requests:
        return {
            'placed': [],
            'conflicts': [],
            'stats': {'total_requests': 0, 'total_placed': 0, 'total_conflicts': 0},
        }

    x = {}
    for i, req in enumerate(requests):
        for j, _ in enumerate(slots):
            x[(i, j)] = model.NewBoolVar(f'x_{i}_{j}')

    for i, req in enumerate(requests):
        model.Add(sum(x[(i, j)] for j in range(n_slots)) <= req['pairs_per_week'])

    busy_group = existing_busy.get('group', {})
    busy_teacher = existing_busy.get('teacher', {})

    for j, (day_idx, para) in enumerate(slots):
        day_name = DAYS_OF_WEEK[day_idx]
        key = (day_name, para)

        for i, req in enumerate(requests):
            gid = req['group_id']
            if busy_group.get(gid, {}).get(key):
                model.Add(x[(i, j)] == 0)

            tid = req['teacher_id']
            if tid and busy_teacher.get(tid, {}).get(key):
                model.Add(x[(i, j)] == 0)

    for j, (day_idx, para) in enumerate(slots):
        day_name = DAYS_OF_WEEK[day_idx]

        by_group = {}
        by_teacher = {}

        for i, req in enumerate(requests):
            gid = req['group_id']
            by_group.setdefault(gid, []).append(i)

            tid = req['teacher_id']
            if tid:
                by_teacher.setdefault(tid, []).append(i)

        for gid, idxs in by_group.items():
            model.Add(sum(x[(i, j)] for i in idxs) <= 1)

        for tid, idxs in by_teacher.items():
            model.Add(sum(x[(i, j)] for i in idxs) <= 1)

    for gid in set(r['group_id'] for r in requests):
        for day_idx in range(len(DAYS_OF_WEEK)):
            day_name = DAYS_OF_WEEK[day_idx]
            idxs = [i for i, r in enumerate(requests) if r['group_id'] == gid]
            day_slots_j = [j for j, (d, p) in enumerate(slots) if d == day_idx]
            model.Add(
                sum(x[(i, j)] for i in idxs for j in day_slots_j) <= max_pairs_per_day
            )

    total_possible = sum(r['pairs_per_week'] for r in requests)
    model.Maximize(sum(x[(i, j)] for i in range(len(requests)) for j in range(n_slots)))

    solver = cp_model.CpSolver()
    solver.parameters.max_time_in_seconds = time_limit_sec
    solver.parameters.num_search_workers = 4

    status = solver.Solve(model)

    placed = []
    placed_count_by_req = {i: 0 for i in range(len(requests))}

    if status in (cp_model.OPTIMAL, cp_model.FEASIBLE):
        for i, req in enumerate(requests):
            for j, (day_idx, para) in enumerate(slots):
                if solver.Value(x[(i, j)]) == 1:
                    day_name = DAYS_OF_WEEK[day_idx]
                    placed.append({
                        'curriculum_id': req['row']['id'],
                        'group_id': req['group_id'],
                        'group_name': req['group_name'],
                        'subject_id': req['subject_id'],
                        'subject_name': req['subject_name'],
                        'teacher_id': req['teacher_id'],
                        'teacher_name': req['teacher_name'],
                        'teacher_short': req['teacher_short'],
                        'day_of_week': day_name,
                        'para_number': para,
                        'hours': 2,
                    })
                    placed_count_by_req[i] += 1

    conflicts = []
    for i, req in enumerate(requests):
        wanted = req['pairs_per_week']
        placed_count = placed_count_by_req[i]
        if placed_count < wanted:
            conflicts.append({
                'curriculum_id': req['row']['id'],
                'group_id': req['group_id'],
                'group_name': req['group_name'],
                'subject_name': req['subject_name'],
                'teacher_name': req['teacher_name'],
                'placed_count': placed_count,
                'wanted': wanted,
                'reason': f"OR-Tools: размещено {placed_count} из {wanted} пар",
            })

    return {
        'placed': placed,
        'conflicts': conflicts,
        'stats': {
            'total_requests': len(requests),
            'total_placed': len(placed),
            'total_conflicts': len(conflicts),
            'solver_status': solver.StatusName(status),
            'solve_time': solver.WallTime(),
        },
    }