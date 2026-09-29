import re
import os
from typing import List, Dict, Any, Optional


INDEX_PATTERN = re.compile(r'^[А-Я]{2,4}\.\d{2}(\.\d{2})?$')
EXCLUDED_INDEXES = {'ПМ', 'ПА', 'ПП', 'ГИА', 'ДР', 'ДЭ', 'УП'}
SEMESTER_HEADER_PATTERN = re.compile(r'^(\d)\s*\(\s*\d+\s*нед\.?\s*\)')


def parse_curriculum(file_path: str) -> Dict[str, Any]:
    ext = os.path.splitext(file_path)[1].lower()

    if ext == '.xls':
        rows = _read_xls(file_path)
    elif ext == '.xlsx':
        rows = _read_xlsx(file_path)
    else:
        raise ValueError(f"Неподдерживаемый формат: {ext}. Только .xls или .xlsx")

    if not rows:
        return {'meta': {'filename': os.path.basename(file_path)}, 'disciplines': []}

    meta = _parse_meta(rows)
    meta['filename'] = os.path.basename(file_path)
    disciplines = _parse_disciplines(rows)

    return {'meta': meta, 'disciplines': disciplines}


def _read_xls(path: str) -> List[List[Any]]:
    import xlrd
    wb = xlrd.open_workbook(path)
    ws = wb.sheet_by_index(0)
    rows = []
    for r in range(ws.nrows):
        row = []
        for c in range(ws.ncols):
            cell = ws.cell(r, c)
            val = cell.value
            if isinstance(val, float) and val.is_integer():
                val = int(val)
            row.append(val)
        rows.append(row)
    return rows


def _read_xlsx(path: str) -> List[List[Any]]:
    import openpyxl
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb.active
    rows = []
    for row in ws.iter_rows(values_only=True):
        rows.append(list(row))
    return rows


def _parse_meta(rows: List[List[Any]]) -> Dict[str, str]:
    meta = {'direction': '', 'form': ''}
    for row in rows[:20]:
        for i, cell in enumerate(row):
            if not cell:
                continue
            s = str(cell).strip()
            if 'Направление подготовки' in s and i + 1 < len(row):
                for j in range(i + 1, min(i + 25, len(row))):
                    if row[j] and str(row[j]).strip():
                        meta['direction'] = str(row[j]).strip()
                        break
            if 'Форма обучения' in s and i + 1 < len(row):
                for j in range(i + 1, min(i + 25, len(row))):
                    if row[j] and str(row[j]).strip():
                        meta['form'] = str(row[j]).strip()
                        break
    return meta


def _parse_disciplines(rows: List[List[Any]]) -> List[Dict[str, Any]]:
    disciplines = []

    header_row = None
    for i, row in enumerate(rows):
        joined = ' '.join(str(c) for c in row if c)
        if 'Индекс' in joined and 'Название' in joined:
            header_row = i
            break

    if header_row is None:
        return []

    header_rows = []
    for j in range(max(0, header_row - 3), min(len(rows), header_row + 5)):
        header_rows.append(rows[j])

    for i in range(header_row + 1, len(rows)):
        row = rows[i]
        if not row or len(row) < 6:
            continue

        num = row[0] if len(row) > 0 else None
        if not isinstance(num, int):
            continue

        index = str(row[1]).strip() if len(row) > 1 and row[1] else ''
        if not index or not INDEX_PATTERN.match(index):
            continue

        prefix = index.split('.')[0]
        if prefix in EXCLUDED_INDEXES:
            continue

        name = ''
        for c in range(2, min(8, len(row))):
            if row[c] and isinstance(row[c], str) and len(str(row[c]).strip()) > 2:
                name = str(row[c]).strip()
                break

        if not name:
            continue

        auditory_hours = _find_hours(row)

        if auditory_hours <= 0:
            continue

        semesters = _find_semesters(row, header_rows)
        controls = _find_controls(row)

        registry = ''
        for c in range(len(row) - 1, 20, -1):
            v = row[c]
            if v and isinstance(v, str) and '№' in v:
                registry = v.strip()
                break

        disciplines.append({
            'row_num': num,
            'index': index,
            'name': name,
            'auditory_hours': auditory_hours,
            'semesters': semesters,
            'controls': controls,
            'registry': registry,
        })

    return disciplines


def _find_hours(row: List[Any]) -> int:
    candidates = []
    for c in range(30, min(50, len(row))):
        v = row[c]
        if isinstance(v, (int, float)) and v > 0:
            candidates.append((c, int(v)))

    if not candidates:
        return 0

    for c, v in candidates:
        if 35 <= c <= 40:
            return v

    return candidates[0][1]


def _find_semester_header_cols(header_rows: List[List[Any]]) -> Dict[int, int]:
    semester_cols = {}
    for hrow in header_rows:
        found = 0
        local_cols = {}
        for c, cell in enumerate(hrow):
            if not cell:
                continue
            s = str(cell).strip()
            m = SEMESTER_HEADER_PATTERN.match(s)
            if m:
                local_cols[int(m.group(1))] = c
                found += 1
        if found >= 4:
            semester_cols = local_cols
            break
    return semester_cols


def _find_semesters(row: List[Any], header_rows: List[List[Any]]) -> List[int]:
    semester_cols = _find_semester_header_cols(header_rows)

    if not semester_cols:
        return _find_semesters_fallback(row)

    sorted_sems = sorted(semester_cols.items())
    if len(sorted_sems) >= 2:
        stride = sorted_sems[1][1] - sorted_sems[0][1]
    else:
        stride = 14

    semesters = []
    for sem_num in range(1, 9):
        if sem_num not in semester_cols:
            continue
        col = semester_cols[sem_num]
        if col >= len(row):
            continue

        v = row[col]
        if isinstance(v, (int, float)) and v > 0:
            semesters.append(sem_num)
            continue

        for offset in range(1, min(4, stride)):
            if col + offset >= len(row):
                break
            v2 = row[col + offset]
            if isinstance(v2, (int, float)) and v2 > 0:
                semesters.append(sem_num)
                break

    return semesters


def _find_semesters_fallback(row: List[Any]) -> List[int]:
    semesters = []
    start = 68
    stride = 14
    for i in range(8):
        col = start + i * stride
        if col >= len(row):
            break
        v = row[col]
        if isinstance(v, (int, float)) and v > 0:
            semesters.append(i + 1)
    return semesters


def _find_controls(row: List[Any]) -> List[str]:
    valid = {'Э', 'З', 'ЗД', 'КР', 'КП', 'КтР', 'ПЗП', 'Эссе'}
    found = []
    for c in range(18, min(30, len(row))):
        v = row[c]
        if v and isinstance(v, str):
            s = v.strip()
            if s in valid and s not in found:
                found.append(s)
    return found