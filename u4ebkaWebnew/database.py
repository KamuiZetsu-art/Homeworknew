import sqlite3
import os
from datetime import datetime, timedelta
import hashlib
import secrets
import json
from typing import Optional, List, Dict, Any, Tuple
from contextlib import contextmanager

try:
    import bcrypt
    BCRYPT_AVAILABLE = True
except ImportError:
    BCRYPT_AVAILABLE = False

from config import config
from logger import get_logger

logger = get_logger(__name__)

DB_FILE = config.DB_FILE

PARA_TIMES = {
    1: "08:30-10:10", 2: "10:20-12:00", 3: "12:35-14:15",
    4: "14:25-16:05", 5: "16:15-17:55", 6: "18:05-19:45"
}

DAYS_OF_WEEK = ['Понедельник', 'Вторник', 'Среда', 'Четверг', 'Пятница', 'Суббота']

ATTENDANCE_STATUSES = {
    'present': {'label': 'Был',     'short': '',  'color': '#2e7d32'},
    'late':    {'label': 'Опоздал', 'short': 'о', 'color': '#f57c00'},
    'absent':  {'label': 'Н',       'short': 'н', 'color': '#c62828'},
}

GRADE_TYPES = {
    'regular':  'Текущая',
    'control':  'Контрольная',
    'test':     'Самостоятельная',
    'oral':     'Устный ответ',
    'homework': 'Домашняя работа',
}

ROLES = {
    'admin':   {'label': 'Администратор', 'color': '#c62828'},
    'editor':  {'label': 'Редактор',      'color': '#f57c00'},
    'teacher': {'label': 'Преподаватель', 'color': '#1565c0'},
    'student': {'label': 'Студент',       'color': '#2e7d32'},
    'parent':  {'label': 'Родитель',      'color': '#6a1b9a'},
}

SEMESTERS = ['1', '2', '3', '4', '5', '6', '7', '8']

SEMESTER_LABELS = {
    '1': '1 семестр',
    '2': '2 семестр',
    '3': '3 семестр',
    '4': '4 семестр',
    '5': '5 семестр',
    '6': '6 семестр',
    '7': '7 семестр',
    '8': '8 семестр',
}


class Database:
    def __init__(self, db_file: str = DB_FILE):
        self.db_file = db_file
        self.init_database()
        self.migrate_database()
        if BCRYPT_AVAILABLE:
            logger.info("bcrypt доступен")
        else:
            logger.warning("bcrypt НЕ установлен")

    def get_connection(self) -> sqlite3.Connection:
        conn = sqlite3.connect(self.db_file, timeout=10.0)
        conn.row_factory = sqlite3.Row
        conn.execute("PRAGMA journal_mode=WAL")
        conn.execute("PRAGMA synchronous=NORMAL")
        conn.execute("PRAGMA foreign_keys=ON")
        conn.execute("PRAGMA cache_size=-64000")
        conn.execute("PRAGMA temp_store=MEMORY")
        return conn

    @contextmanager
    def connect(self):
        conn = self.get_connection()
        try:
            yield conn
            conn.commit()
        except Exception:
            conn.rollback()
            raise
        finally:
            conn.close()

    def table_exists(self, cursor, table_name: str) -> bool:
        cursor.execute("SELECT COUNT(*) FROM sqlite_master WHERE type='table' AND name=?", (table_name,))
        return cursor.fetchone()[0] > 0

    def column_exists(self, cursor, table_name: str, column_name: str) -> bool:
        cursor.execute(f"PRAGMA table_info({table_name})")
        columns = [row[1] for row in cursor.fetchall()]
        return column_name in columns

    def init_database(self):
        conn = self.get_connection()
        cursor = conn.cursor()

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS users (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                username TEXT UNIQUE NOT NULL,
                password_hash TEXT NOT NULL,
                full_name TEXT,
                role TEXT NOT NULL DEFAULT 'student',
                email TEXT DEFAULT '',
                phone TEXT DEFAULT '',
                group_name TEXT DEFAULT '',
                parent_of TEXT DEFAULT '',
                teacher_id INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                last_login TIMESTAMP,
                is_active INTEGER DEFAULT 1
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS groups (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                specialty TEXT,
                course INTEGER,
                students_count INTEGER DEFAULT 0,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS teachers (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                full_name TEXT NOT NULL,
                short_name TEXT,
                department TEXT,
                email TEXT,
                phone TEXT,
                FOREIGN KEY (user_id) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS subjects (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT UNIQUE NOT NULL,
                short_name TEXT,
                department TEXT
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS rooms (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                number TEXT UNIQUE NOT NULL,
                building TEXT,
                capacity INTEGER,
                room_type TEXT
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schedule (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                teacher_id INTEGER,
                subject_id INTEGER,
                room_id INTEGER,
                day_of_week TEXT NOT NULL,
                para_number INTEGER NOT NULL,
                hours INTEGER DEFAULT 2,
                comment TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_by INTEGER,
                is_active INTEGER DEFAULT 1,
                FOREIGN KEY (group_id) REFERENCES groups(id),
                FOREIGN KEY (teacher_id) REFERENCES teachers(id),
                FOREIGN KEY (subject_id) REFERENCES subjects(id),
                FOREIGN KEY (room_id) REFERENCES rooms(id),
                FOREIGN KEY (created_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schedule_history (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                schedule_id INTEGER,
                action TEXT NOT NULL,
                old_data TEXT,
                new_data TEXT,
                changed_by INTEGER,
                changed_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (changed_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS schedule_overrides (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lesson_id INTEGER,
                date DATE NOT NULL,
                group_id INTEGER,
                para_number INTEGER,
                override_type TEXT NOT NULL,
                new_teacher_id INTEGER,
                new_room_id INTEGER,
                new_subject_id INTEGER,
                new_hours INTEGER,
                comment TEXT,
                created_by INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                is_active INTEGER DEFAULT 1,
                FOREIGN KEY (lesson_id) REFERENCES schedule(id),
                FOREIGN KEY (group_id) REFERENCES groups(id),
                FOREIGN KEY (new_teacher_id) REFERENCES teachers(id),
                FOREIGN KEY (new_room_id) REFERENCES rooms(id),
                FOREIGN KEY (new_subject_id) REFERENCES subjects(id),
                FOREIGN KEY (created_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS calendar (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date DATE NOT NULL,
                day_of_week TEXT NOT NULL,
                is_holiday INTEGER DEFAULT 0,
                holiday_name TEXT,
                is_working_day INTEGER DEFAULT 1,
                comment TEXT,
                UNIQUE(date)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS sessions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                token TEXT UNIQUE,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                expires_at TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS import_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                filename TEXT,
                imported_count INTEGER,
                errors_count INTEGER,
                imported_by INTEGER,
                imported_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (imported_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS attendance (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                lesson_id INTEGER NOT NULL,
                date DATE NOT NULL,
                status TEXT DEFAULT 'present',
                comment TEXT,
                marked_by INTEGER,
                marked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(student_id, lesson_id, date),
                FOREIGN KEY (student_id) REFERENCES users(id),
                FOREIGN KEY (lesson_id) REFERENCES schedule(id),
                FOREIGN KEY (marked_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS grades (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                student_id INTEGER NOT NULL,
                lesson_id INTEGER NOT NULL,
                date DATE NOT NULL,
                grade INTEGER,
                grade_type TEXT DEFAULT 'regular',
                comment TEXT,
                marked_by INTEGER,
                marked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (student_id) REFERENCES users(id),
                FOREIGN KEY (lesson_id) REFERENCES schedule(id),
                FOREIGN KEY (marked_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS lessons_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                lesson_id INTEGER NOT NULL,
                date DATE NOT NULL,
                topic TEXT,
                homework TEXT,
                comment TEXT,
                is_conducted INTEGER DEFAULT 1,
                marked_by INTEGER,
                marked_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                UNIQUE(lesson_id, date),
                FOREIGN KEY (lesson_id) REFERENCES schedule(id),
                FOREIGN KEY (marked_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS credentials_log (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                user_id INTEGER,
                username TEXT NOT NULL,
                password_plain TEXT NOT NULL,
                role TEXT NOT NULL,
                full_name TEXT,
                group_name TEXT,
                linked_student_id INTEGER,
                linked_student_login TEXT,
                created_by INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (user_id) REFERENCES users(id),
                FOREIGN KEY (created_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS group_prefixes (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_name TEXT UNIQUE NOT NULL,
                prefix TEXT NOT NULL,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS curriculum (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                subject_id INTEGER NOT NULL,
                teacher_id INTEGER,
                planned_hours INTEGER NOT NULL DEFAULT 0,
                pairs_per_week INTEGER NOT NULL DEFAULT 2,
                semester TEXT NOT NULL DEFAULT '1',
                academic_year TEXT NOT NULL DEFAULT '2025/2026',
                priority INTEGER NOT NULL DEFAULT 2,
                comment TEXT,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                updated_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                created_by INTEGER,
                FOREIGN KEY (group_id) REFERENCES groups(id),
                FOREIGN KEY (subject_id) REFERENCES subjects(id),
                FOREIGN KEY (teacher_id) REFERENCES teachers(id),
                FOREIGN KEY (created_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS curriculum_imports (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                group_id INTEGER NOT NULL,
                source_filename TEXT,
                meta_json TEXT,
                disciplines_json TEXT,
                imported_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                imported_by INTEGER,
                FOREIGN KEY (group_id) REFERENCES groups(id),
                FOREIGN KEY (imported_by) REFERENCES users(id)
            )
        ''')

        cursor.execute('''
            CREATE TABLE IF NOT EXISTS archive_snapshots (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                name TEXT NOT NULL,
                semester TEXT NOT NULL,
                academic_year TEXT NOT NULL,
                date_from DATE NOT NULL,
                date_to DATE NOT NULL,
                schedule_json TEXT,
                grades_json TEXT,
                attendance_json TEXT,
                curriculum_json TEXT,
                stats_json TEXT,
                created_by INTEGER,
                created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                FOREIGN KEY (created_by) REFERENCES users(id)
            )
        ''')

        cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedule_day_para ON schedule(day_of_week, para_number, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedule_group ON schedule(group_id, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_schedule_teacher ON schedule(teacher_id, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_grades_lesson_date ON grades(lesson_id, date)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_grades_student ON grades(student_id, date)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_attendance_lesson_date ON attendance(lesson_id, date)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_attendance_student ON attendance(student_id, date)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_group_role ON users(group_name, role, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_users_username ON users(username)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_sessions_token ON sessions(token)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_overrides_date ON schedule_overrides(date, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_overrides_lesson ON schedule_overrides(lesson_id, date, is_active)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_curriculum_group ON curriculum(group_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_curriculum_subject ON curriculum(subject_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_curriculum_semester ON curriculum(semester, academic_year)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_curr_imports_group ON curriculum_imports(group_id)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_curr_imports_date ON curriculum_imports(imported_at)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_archive_semester ON archive_snapshots(semester, academic_year)")
        cursor.execute("CREATE INDEX IF NOT EXISTS idx_archive_dates ON archive_snapshots(date_from, date_to)")

        conn.commit()
        conn.close()

    def migrate_database(self):
        conn = self.get_connection()
        cursor = conn.cursor()

        try:
            if not self.column_exists(cursor, 'users', 'parent_of'):
                cursor.execute("ALTER TABLE users ADD COLUMN parent_of TEXT DEFAULT ''")
                conn.commit()

            if not self.column_exists(cursor, 'users', 'teacher_id'):
                cursor.execute("ALTER TABLE users ADD COLUMN teacher_id INTEGER")
                conn.commit()

            if self.table_exists(cursor, 'schedule_overrides'):
                if not self.column_exists(cursor, 'schedule_overrides', 'group_id'):
                    cursor.execute("ALTER TABLE schedule_overrides ADD COLUMN group_id INTEGER")
                    conn.commit()
                if not self.column_exists(cursor, 'schedule_overrides', 'para_number'):
                    cursor.execute("ALTER TABLE schedule_overrides ADD COLUMN para_number INTEGER")
                    conn.commit()

                cursor.execute("CREATE INDEX IF NOT EXISTS idx_overrides_group_date ON schedule_overrides(group_id, date, is_active)")
                conn.commit()

            cursor.execute("SELECT COUNT(*) FROM users WHERE role='viewer'")
            viewer_count = cursor.fetchone()[0]
            if viewer_count > 0:
                cursor.execute("UPDATE users SET role='student' WHERE role='viewer'")
                conn.commit()

            cursor.execute("SELECT COUNT(*) FROM users WHERE role='admin'")
            if cursor.fetchone()[0] == 0:
                password_hash = self.hash_password("admin123")
                cursor.execute('''
                    INSERT INTO users (username, password_hash, full_name, role)
                    VALUES (?, ?, ?, ?)
                ''', ("admin", password_hash, "Администратор", "admin"))
                conn.commit()

        except Exception as e:
            logger.error(f"Ошибка миграции: {e}")
            conn.rollback()
        finally:
            conn.close()

    def hash_password(self, password: str) -> str:
        if BCRYPT_AVAILABLE:
            salt = bcrypt.gensalt(rounds=12)
            return bcrypt.hashpw(password.encode('utf-8'), salt).decode('utf-8')
        return 'sha256$' + hashlib.sha256(password.encode()).hexdigest()

    def verify_password(self, password: str, password_hash: str) -> bool:
        if not password_hash:
            return False
        if password_hash.startswith(('$2a$', '$2b$', '$2y$')):
            if not BCRYPT_AVAILABLE:
                return False
            try:
                return bcrypt.checkpw(password.encode('utf-8'), password_hash.encode('utf-8'))
            except Exception:
                return False
        if password_hash.startswith('sha256$'):
            old_hash = password_hash[7:]
        else:
            old_hash = password_hash
        return hashlib.sha256(password.encode()).hexdigest() == old_hash

    def needs_rehash(self, password_hash: str) -> bool:
        if not password_hash or not BCRYPT_AVAILABLE:
            return False
        if not password_hash.startswith(('$2a$', '$2b$', '$2y$')):
            return True
        try:
            if password_hash.startswith('$2b$'):
                rounds = int(password_hash.split('$')[2])
                return rounds < 12
        except (IndexError, ValueError):
            pass
        return False

    def _normalize_name(self, full_name: str) -> str:
        if not full_name:
            return ''
        s = str(full_name).strip()
        s = ' '.join(s.split())
        s = s.replace('ё', 'е').replace('Ё', 'Е')
        return s.lower()

    def create_user(self, username: str, password: str, full_name: str, role: str = 'student',
                    email: str = '', phone: str = '', group_name: str = '',
                    parent_of: str = '', teacher_id: Optional[int] = None) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            password_hash = self.hash_password(password)
            cursor.execute('''
                INSERT INTO users (username, password_hash, full_name, role, email, phone,
                                   group_name, parent_of, teacher_id)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (username, password_hash, full_name, role, email or '',
                  phone or '', group_name or '', parent_of or '', teacher_id))
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()

    def user_exists(self, username: str) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM users WHERE username=?", (username,))
        count = cursor.fetchone()[0]
        conn.close()
        return count > 0

    def authenticate(self, username: str, password: str) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE username=? AND is_active=1", (username,))
        user = cursor.fetchone()
        if not user:
            conn.close()
            return None
        if not self.verify_password(password, user['password_hash']):
            conn.close()
            return None
        if self.needs_rehash(user['password_hash']):
            try:
                new_hash = self.hash_password(password)
                cursor.execute("UPDATE users SET password_hash=? WHERE id=?", (new_hash, user['id']))
            except Exception as e:
                logger.error(f"Ошибка перехеширования: {e}")
        cursor.execute("UPDATE users SET last_login=CURRENT_TIMESTAMP WHERE id=?", (user['id'],))
        conn.commit()
        token = secrets.token_urlsafe(32)
        expires = datetime.now() + timedelta(days=7)
        cursor.execute('''
            INSERT INTO sessions (user_id, token, expires_at)
            VALUES (?, ?, ?)
        ''', (user['id'], token, expires))
        conn.commit()
        result = dict(user)
        result['token'] = token
        conn.close()
        return result

    def get_user_by_token(self, token: str) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT u.* FROM users u
            JOIN sessions s ON u.id = s.user_id
            WHERE s.token=? AND s.expires_at > CURRENT_TIMESTAMP AND u.is_active=1
        ''', (token,))
        user = cursor.fetchone()
        conn.close()
        return dict(user) if user else None

    def get_all_users(self, only_active: bool = True) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        where = "WHERE is_active=1" if only_active else ""
        cursor.execute(f"""
            SELECT id, username, full_name, role, email, phone, group_name, parent_of,
                   teacher_id, created_at, last_login, is_active
            FROM users
            {where}
            ORDER BY
                CASE role
                    WHEN 'admin' THEN 1
                    WHEN 'editor' THEN 2
                    WHEN 'teacher' THEN 3
                    WHEN 'student' THEN 4
                    WHEN 'parent' THEN 5
                END,
                full_name
        """)
        users = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return users

    def get_students_by_group(self, group_name: str) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT * FROM users
            WHERE role IN ('student', 'viewer') AND group_name=? AND is_active=1
            ORDER BY full_name
        ''', (group_name,))
        students = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return students

    def get_user(self, user_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM users WHERE id=?", (user_id,))
        user = cursor.fetchone()
        conn.close()
        return dict(user) if user else None

    def delete_user(self, user_id: int):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET is_active=0 WHERE id=?", (user_id,))
        conn.commit()
        conn.close()

    def change_password(self, user_id: int, new_password: str):
        conn = self.get_connection()
        cursor = conn.cursor()
        new_hash = self.hash_password(new_password)
        cursor.execute("UPDATE users SET password_hash=? WHERE id=?", (new_hash, user_id))
        cursor.execute("UPDATE credentials_log SET password_plain=? WHERE user_id=?", (new_password, user_id))
        conn.commit()
        conn.close()

    def update_user(self, user_id: int, fields: Dict[str, Any]) -> bool:
        allowed = ['full_name', 'email', 'phone', 'group_name', 'parent_of', 'teacher_id']
        updates = []
        params = []
        for key, value in fields.items():
            if key in allowed:
                updates.append(f"{key}=?")
                params.append(value)
        if not updates:
            return False
        params.append(user_id)
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(f"UPDATE users SET {', '.join(updates)} WHERE id=?", params)
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Ошибка обновления пользователя: {e}")
            return False
        finally:
            conn.close()

    def get_parent_children_detailed(self, parent_of: str) -> List[Dict[str, Any]]:
        if not parent_of:
            return []
        logins = [l.strip() for l in parent_of.split(',') if l.strip()]
        if not logins:
            return []
        conn = self.get_connection()
        cursor = conn.cursor()
        placeholders = ','.join('?' * len(logins))
        cursor.execute(f'''
            SELECT * FROM users
            WHERE username IN ({placeholders})
              AND role IN ('student', 'viewer')
              AND is_active = 1
            ORDER BY full_name
        ''', logins)
        children = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return children

    def get_all_groups(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM groups ORDER BY name")
        groups = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return groups

    def add_group(self, name: str, specialty: Optional[str] = None) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('INSERT INTO groups (name, specialty) VALUES (?, ?)', (name, specialty))
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()

    def delete_group(self, group_id: int):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM groups WHERE id=?", (group_id,))
        conn.commit()
        conn.close()

    def get_all_teachers(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM teachers ORDER BY full_name")
        teachers = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return teachers

    def add_teacher(self, full_name: str, short_name: Optional[str] = None,
                    department: Optional[str] = None, user_id: Optional[int] = None) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        if not short_name:
            parts = full_name.split()
            if len(parts) >= 3:
                short_name = f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
            elif len(parts) >= 2:
                short_name = f"{parts[0]} {parts[1][0]}."
            else:
                short_name = full_name
        cursor.execute('''
            INSERT INTO teachers (user_id, full_name, short_name, department)
            VALUES (?, ?, ?, ?)
        ''', (user_id, full_name, short_name, department))
        conn.commit()
        teacher_id = cursor.lastrowid
        conn.close()
        return teacher_id

    def delete_teacher(self, teacher_id: int):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM teachers WHERE id=?", (teacher_id,))
        conn.commit()
        conn.close()

    def get_teacher_by_user(self, user_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM teachers WHERE user_id=?", (user_id,))
        teacher = cursor.fetchone()
        conn.close()
        return dict(teacher) if teacher else None

    def get_teacher_by_id(self, teacher_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM teachers WHERE id=?", (teacher_id,))
        teacher = cursor.fetchone()
        conn.close()
        return dict(teacher) if teacher else None

    def find_or_create_teacher(self, full_name: str, email: str = '') -> Tuple[Optional[int], bool, Optional[int]]:
        norm = self._normalize_name(full_name)
        if not norm:
            return None, False, None
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT id, user_id, full_name FROM teachers ORDER BY id")
        all_teachers = cursor.fetchall()
        matches = [dict(t) for t in all_teachers if self._normalize_name(t['full_name']) == norm]
        if matches:
            with_user = [m for m in matches if m['user_id']]
            if with_user:
                conn.close()
                return with_user[0]['id'], False, None
            conn.close()
            return matches[0]['id'], False, None
        parts = full_name.split()
        if len(parts) >= 3:
            short = f"{parts[0]} {parts[1][0]}.{parts[2][0]}."
        elif len(parts) >= 2:
            short = f"{parts[0]} {parts[1][0]}."
        else:
            short = full_name
        cursor.execute('''
            INSERT INTO teachers (full_name, short_name, email)
            VALUES (?, ?, ?)
        ''', (full_name.strip(), short, email))
        conn.commit()
        new_id = cursor.lastrowid
        conn.close()
        return new_id, True, None

    def link_teacher_user(self, teacher_id: int, user_id: int):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE teachers SET user_id=? WHERE id=?", (user_id, teacher_id))
        conn.commit()
        conn.close()

    def find_duplicate_teachers(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT t.id, t.full_name, t.user_id, t.short_name, u.username
            FROM teachers t
            LEFT JOIN users u ON t.user_id = u.id
            ORDER BY t.id
        ''')
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        groups = {}
        for r in rows:
            norm = self._normalize_name(r['full_name'])
            if not norm:
                continue
            if norm not in groups:
                groups[norm] = []
            groups[norm].append(r)
        result = []
        for norm, items in groups.items():
            if len(items) <= 1:
                continue
            with_user = [x for x in items if x['user_id']]
            without_user = [x for x in items if not x['user_id']]
            if with_user:
                main = with_user[0]
                duplicates = with_user[1:] + without_user
            else:
                main = items[0]
                duplicates = items[1:]
            if not duplicates:
                continue
            result.append({
                'norm_name': norm,
                'main': main,
                'duplicates': duplicates,
                'total': len(items),
            })
        return result

    def delete_teacher_duplicates(self) -> Dict[str, int]:
        groups = self.find_duplicate_teachers()
        deleted = 0
        skipped = 0
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            for g in groups:
                main_id = g['main']['id']
                for dup in g['duplicates']:
                    if dup['user_id']:
                        skipped += 1
                        continue
                    cursor.execute("UPDATE schedule SET teacher_id=? WHERE teacher_id=?",
                                   (main_id, dup['id']))
                    cursor.execute("DELETE FROM teachers WHERE id=?", (dup['id'],))
                    deleted += 1
            conn.commit()
        except Exception as e:
            logger.error(f"Ошибка очистки дублей: {e}")
            conn.rollback()
        finally:
            conn.close()
        return {'deleted': deleted, 'skipped': skipped}

    def get_all_subjects(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM subjects ORDER BY name")
        subjects = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return subjects

    def add_subject(self, name: str) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('INSERT INTO subjects (name) VALUES (?)', (name,))
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()
    def get_subject_by_name(self, name: str) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM subjects WHERE name=?", (name,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def delete_subject(self, subject_id: int) -> Dict[str, Any]:
        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT COUNT(*) FROM schedule
                WHERE subject_id = ? AND is_active = 1
            ''', (subject_id,))
            usage = cursor.fetchone()[0]
            if usage > 0:
                return {
                    'success': False,
                    'error': f'Предмет используется в {usage} парах.',
                    'usage': usage,
                }
            cursor.execute("DELETE FROM subjects WHERE id = ?", (subject_id,))
            return {'success': True}

    def get_all_rooms(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM rooms ORDER BY number")
        rooms = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return rooms

    def get_busy_rooms_for_slot(self, day_of_week: str, para_number: int,
                                  exclude_lesson_id: Optional[int] = None) -> Dict[int, Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT s.room_id,
                   g.name AS group_name,
                   sub.name AS subject_name,
                   t.short_name AS teacher_short,
                   t.full_name AS teacher_name
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            WHERE s.day_of_week = ?
              AND s.para_number = ?
              AND s.is_active = 1
              AND s.room_id IS NOT NULL
        '''
        params = [day_of_week, int(para_number)]
        if exclude_lesson_id:
            query += " AND s.id != ?"
            params.append(int(exclude_lesson_id))

        cursor.execute(query, params)
        result = {}
        for row in cursor.fetchall():
            result[row['room_id']] = {
                'group_name': row['group_name'] or '',
                'subject_name': row['subject_name'] or '',
                'teacher_short': row['teacher_short'] or row['teacher_name'] or '',
            }
        conn.close()
        return result

    def add_room(self, number: str) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('INSERT INTO rooms (number) VALUES (?)', (number,))
            conn.commit()
            return cursor.lastrowid
        except sqlite3.IntegrityError:
            return None
        finally:
            conn.close()

    def get_room(self, room_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM rooms WHERE id=?", (room_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def update_room(self, room_id: int, fields: Dict[str, Any]) -> bool:
        """Обновляет поля кабинета. Разрешены: number, building, capacity, room_type, subjects_allowed."""
        allowed = ['number', 'building', 'capacity', 'room_type', 'subjects_allowed']
        updates = []
        params = []
        for key, value in fields.items():
            if key in allowed:
                updates.append(f"{key} = ?")
                params.append(value)
        if not updates:
            return False
        params.append(room_id)

        with self.connect() as conn:
            cursor = conn.cursor()
            try:
                cursor.execute(
                    f"UPDATE rooms SET {', '.join(updates)} WHERE id = ?",
                    params
                )
                return cursor.rowcount > 0
            except Exception as e:
                logger.error(f"update_room error: {e}")
                raise

    def delete_room(self, room_id: int) -> Dict[str, Any]:
        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT COUNT(*) FROM schedule
                WHERE room_id = ? AND is_active = 1
            ''', (room_id,))
            usage = cursor.fetchone()[0]
            if usage > 0:
                return {
                    'success': False,
                    'error': f'Кабинет используется в {usage} парах.',
                    'usage': usage,
                }
            cursor.execute("DELETE FROM rooms WHERE id = ?", (room_id,))
            return {'success': True}

    def add_curriculum_row(self, group_id: int, subject_id: int,
                           teacher_id: Optional[int],
                           planned_hours: int,
                           pairs_per_week: int = 2,
                           semester: str = '1',
                           academic_year: str = '2025/2026',
                           priority: int = 2,
                           comment: str = '',
                           created_by: Optional[int] = None) -> Optional[int]:
        if priority not in (1, 2, 3):
            priority = 2
        if str(semester) not in SEMESTERS:
            semester = '1'

        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                INSERT INTO curriculum
                (group_id, subject_id, teacher_id, planned_hours, pairs_per_week,
                 semester, academic_year, priority, comment, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (group_id, subject_id, teacher_id, planned_hours, pairs_per_week,
                  str(semester), academic_year, priority, comment, created_by))
            return cursor.lastrowid

    def get_curriculum(self, group_id: Optional[int] = None,
                       semester: Optional[str] = None,
                       academic_year: Optional[str] = None) -> List[Dict[str, Any]]:
        query = '''
            SELECT c.*,
                   g.name AS group_name,
                   sub.name AS subject_name,
                   t.full_name AS teacher_name,
                   t.short_name AS teacher_short
            FROM curriculum c
            LEFT JOIN groups g ON c.group_id = g.id
            LEFT JOIN subjects sub ON c.subject_id = sub.id
            LEFT JOIN teachers t ON c.teacher_id = t.id
            WHERE 1=1
        '''
        params: List[Any] = []
        if group_id:
            query += " AND c.group_id = ?"
            params.append(group_id)
        if semester:
            query += " AND c.semester = ?"
            params.append(str(semester))
        if academic_year:
            query += " AND c.academic_year = ?"
            params.append(academic_year)
        query += " ORDER BY g.name, c.semester, sub.name, c.priority"

        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def get_curriculum_row(self, row_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT c.*,
                   g.name AS group_name,
                   sub.name AS subject_name,
                   t.full_name AS teacher_name,
                   t.short_name AS teacher_short
            FROM curriculum c
            LEFT JOIN groups g ON c.group_id = g.id
            LEFT JOIN subjects sub ON c.subject_id = sub.id
            LEFT JOIN teachers t ON c.teacher_id = t.id
            WHERE c.id = ?
        ''', (row_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def update_curriculum_row(self, row_id: int, fields: Dict[str, Any]) -> bool:
        allowed = ['group_id', 'subject_id', 'teacher_id', 'planned_hours',
                   'pairs_per_week', 'semester', 'academic_year', 'priority', 'comment']
        updates = []
        params = []
        for key, value in fields.items():
            if key in allowed:
                updates.append(f"{key} = ?")
                params.append(value)
        if not updates:
            return False
        updates.append("updated_at = CURRENT_TIMESTAMP")
        params.append(row_id)

        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute(f"UPDATE curriculum SET {', '.join(updates)} WHERE id = ?", params)
            return cursor.rowcount > 0

    def delete_curriculum_row(self, row_id: int) -> bool:
        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM curriculum WHERE id = ?", (row_id,))
            return cursor.rowcount > 0

    def get_curriculum_with_fact(self, group_id: Optional[int] = None,
                                  semester: Optional[str] = None,
                                  academic_year: Optional[str] = None) -> List[Dict[str, Any]]:
        rows = self.get_curriculum(group_id, semester, academic_year)
        if not rows:
            return rows

        conn = self.get_connection()
        cursor = conn.cursor()

        for row in rows:
            gid = row['group_id']
            sid = row['subject_id']

            cursor.execute('''
                SELECT COALESCE(SUM(hours), 0) AS total_hours,
                       COUNT(*) AS lessons_count
                FROM schedule
                WHERE group_id = ? AND subject_id = ? AND is_active = 1
            ''', (gid, sid))
            sch = cursor.fetchone()
            plan_from_schedule = sch['total_hours'] or 0
            lessons_count = sch['lessons_count'] or 0

            cursor.execute('''
                SELECT COALESCE(SUM(s.hours), 0) AS cancelled_hours
                FROM schedule_overrides o
                JOIN schedule s ON s.id = o.lesson_id
                WHERE o.override_type = 'cancelled'
                  AND o.is_active = 1
                  AND s.group_id = ? AND s.subject_id = ? AND s.is_active = 1
            ''', (gid, sid))
            cancelled = cursor.fetchone()['cancelled_hours'] or 0

            fact_hours = max(0, plan_from_schedule - cancelled)
            planned = row['planned_hours'] or 0
            remaining = planned - fact_hours
            percent = round(fact_hours / planned * 100, 1) if planned else 0

            row['fact_hours'] = fact_hours
            row['cancelled_hours'] = cancelled
            row['remaining_hours'] = remaining
            row['percent'] = percent
            row['lessons_count'] = lessons_count

        conn.close()
        return rows

    def save_curriculum_import(self, group_id: int, filename: str,
                                 meta: dict, disciplines: list,
                                 imported_by: int) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO curriculum_imports
                (group_id, source_filename, meta_json, disciplines_json, imported_by)
                VALUES (?, ?, ?, ?, ?)
            ''', (group_id, filename,
                  json.dumps(meta, ensure_ascii=False),
                  json.dumps(disciplines, ensure_ascii=False),
                  imported_by))
            conn.commit()
            return cursor.lastrowid
        except Exception as e:
            logger.error(f"Ошибка save_curriculum_import: {e}")
            return None
        finally:
            conn.close()

    def get_curriculum_imports(self, group_id: Optional[int] = None,
                                 limit: int = 20) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT ci.id, ci.group_id, ci.source_filename, ci.imported_at,
                   ci.disciplines_json,
                   g.name AS group_name,
                   u.full_name AS imported_by_name
            FROM curriculum_imports ci
            LEFT JOIN groups g ON ci.group_id = g.id
            LEFT JOIN users u ON ci.imported_by = u.id
        '''
        params: List[Any] = []
        if group_id:
            query += " WHERE ci.group_id = ?"
            params.append(group_id)
        query += " ORDER BY ci.imported_at DESC LIMIT ?"
        params.append(limit)
        cursor.execute(query, params)
        rows = []
        for r in cursor.fetchall():
            d = dict(r)
            try:
                d['disciplines'] = json.loads(d.pop('disciplines_json') or '[]')
            except Exception:
                d['disciplines'] = []
            rows.append(d)
        conn.close()
        return rows

    def create_archive_snapshot(self, name: str, semester: str,
                                 academic_year: str,
                                 date_from: str, date_to: str,
                                 created_by: Optional[int] = None) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()

        try:
            cursor.execute('''
                SELECT s.id, s.day_of_week, s.para_number, s.hours,
                       g.name AS group_name,
                       t.full_name AS teacher_name,
                       t.short_name AS teacher_short,
                       sub.name AS subject_name,
                       r.number AS room_number
                FROM schedule s
                LEFT JOIN groups g ON s.group_id = g.id
                LEFT JOIN teachers t ON s.teacher_id = t.id
                LEFT JOIN subjects sub ON s.subject_id = sub.id
                LEFT JOIN rooms r ON s.room_id = r.id
                WHERE s.is_active = 1
                ORDER BY g.name, s.day_of_week, s.para_number
            ''')
            schedule_base = [dict(r) for r in cursor.fetchall()]

            cursor.execute('''
                SELECT o.id, o.date, o.override_type, o.comment,
                       o.lesson_id, o.group_id, o.para_number,
                       o.new_teacher_id, o.new_room_id, o.new_subject_id, o.new_hours,
                       g.name AS group_name,
                       nt.full_name AS new_teacher_name, nt.short_name AS new_teacher_short,
                       nr.number AS new_room_number,
                       ns.name AS new_subject_name
                FROM schedule_overrides o
                LEFT JOIN groups g ON o.group_id = g.id
                LEFT JOIN teachers nt ON o.new_teacher_id = nt.id
                LEFT JOIN rooms nr ON o.new_room_id = nr.id
                LEFT JOIN subjects ns ON o.new_subject_id = ns.id
                WHERE o.is_active = 1
                  AND o.date >= ? AND o.date <= ?
                ORDER BY o.date, o.id
            ''', (date_from, date_to))
            schedule_overrides = [dict(r) for r in cursor.fetchall()]

            schedule_data = {
                'base': schedule_base,
                'overrides': schedule_overrides,
            }

            cursor.execute('''
                SELECT g.id, g.date, g.grade, g.grade_type, g.comment,
                       u.full_name AS student_name, u.group_name AS student_group,
                       sub.name AS subject_name,
                       t.full_name AS teacher_name,
                       s.day_of_week, s.para_number
                FROM grades g
                LEFT JOIN users u ON g.student_id = u.id
                LEFT JOIN schedule s ON g.lesson_id = s.id
                LEFT JOIN subjects sub ON s.subject_id = sub.id
                LEFT JOIN teachers t ON s.teacher_id = t.id
                WHERE g.date >= ? AND g.date <= ?
                ORDER BY g.date, u.full_name
            ''', (date_from, date_to))
            grades_data = [dict(r) for r in cursor.fetchall()]

            cursor.execute('''
                SELECT a.id, a.date, a.status, a.comment,
                       u.full_name AS student_name, u.group_name AS student_group,
                       sub.name AS subject_name,
                       t.full_name AS teacher_name,
                       s.day_of_week, s.para_number
                FROM attendance a
                LEFT JOIN users u ON a.student_id = u.id
                LEFT JOIN schedule s ON a.lesson_id = s.id
                LEFT JOIN subjects sub ON s.subject_id = sub.id
                LEFT JOIN teachers t ON s.teacher_id = t.id
                WHERE a.date >= ? AND a.date <= ?
                ORDER BY a.date, u.full_name
            ''', (date_from, date_to))
            attendance_data = [dict(r) for r in cursor.fetchall()]

            curriculum_data = self.get_curriculum_with_fact(
                semester=semester,
                academic_year=academic_year,
            )

            total_students = len(set(g['student_name'] for g in grades_data if g.get('student_name')))
            total_teachers = len(set(g['teacher_name'] for g in grades_data if g.get('teacher_name')))
            grades_count = len(grades_data)
            attendance_count = len(attendance_data)
            schedule_count = len(schedule_base)
            overrides_count = len(schedule_overrides)

            avg_grade = 0
            grade_values = [g['grade'] for g in grades_data if g.get('grade') is not None]
            if grade_values:
                avg_grade = round(sum(grade_values) / len(grade_values), 2)

            stats = {
                'total_students': total_students,
                'total_teachers': total_teachers,
                'grades_count': grades_count,
                'attendance_count': attendance_count,
                'schedule_count': schedule_count,
                'overrides_count': overrides_count,
                'avg_grade': avg_grade,
                'date_from': date_from,
                'date_to': date_to,
            }

            cursor.execute('''
                INSERT INTO archive_snapshots
                (name, semester, academic_year, date_from, date_to,
                 schedule_json, grades_json, attendance_json,
                 curriculum_json, stats_json, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
            ''', (
                name, semester, academic_year, date_from, date_to,
                json.dumps(schedule_data, ensure_ascii=False),
                json.dumps(grades_data, ensure_ascii=False),
                json.dumps(attendance_data, ensure_ascii=False),
                json.dumps(curriculum_data, ensure_ascii=False, default=str),
                json.dumps(stats, ensure_ascii=False),
                created_by,
            ))
            conn.commit()
            snapshot_id = cursor.lastrowid
            return snapshot_id

        except Exception as e:
            logger.error(f"Ошибка создания снимка: {e}")
            conn.rollback()
            return None
        finally:
            conn.close()

    def list_archive_snapshots(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT a.id, a.name, a.semester, a.academic_year,
                   a.date_from, a.date_to, a.stats_json,
                   a.created_at,
                   u.full_name AS created_by_name
            FROM archive_snapshots a
            LEFT JOIN users u ON a.created_by = u.id
            ORDER BY a.created_at DESC
        ''')
        rows = []
        for r in cursor.fetchall():
            row = dict(r)
            try:
                row['stats'] = json.loads(row.pop('stats_json') or '{}')
            except Exception:
                row['stats'] = {}
            rows.append(row)
        conn.close()
        return rows

    def get_archive_snapshot(self, snapshot_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT a.*, u.full_name AS created_by_name
            FROM archive_snapshots a
            LEFT JOIN users u ON a.created_by = u.id
            WHERE a.id = ?
        ''', (snapshot_id,))
        row = cursor.fetchone()
        conn.close()
        if not row:
            return None
        result = dict(row)
        for field in ('schedule_json', 'grades_json', 'attendance_json',
                      'curriculum_json', 'stats_json'):
            try:
                result[field.replace('_json', '')] = json.loads(result.pop(field) or 'null')
            except Exception:
                result[field.replace('_json', '')] = None
        return result

    def delete_archive_snapshot(self, snapshot_id: int) -> bool:
        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute("DELETE FROM archive_snapshots WHERE id = ?", (snapshot_id,))
            return cursor.rowcount > 0

    def get_schedule_by_day(self, day_of_week: str, group_name: Optional[str] = None) -> Dict[str, Dict[int, Dict[str, Any]]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                s.id, s.para_number, s.hours,
                g.name as group_name, g.id as group_id,
                t.full_name as teacher_name, t.short_name as teacher_short, t.id as teacher_id,
                sub.name as subject_name, sub.id as subject_id,
                r.number as room_number, r.id as room_id
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.day_of_week = ? AND s.is_active = 1
        '''
        params = [day_of_week]
        if group_name:
            query += " AND g.name = ?"
            params.append(group_name)
        query += " ORDER BY g.name, s.para_number"
        cursor.execute(query, params)
        result = {}
        for row in cursor.fetchall():
            row_dict = dict(row)
            group = row_dict['group_name']
            para = row_dict['para_number']
            if group not in result:
                result[group] = {}
            result[group][para] = {
                'id': row_dict['id'],
                'предмет': row_dict['subject_name'] or '',
                'преподаватель': row_dict['teacher_name'] or '',
                'преподаватель_коротко': row_dict['teacher_short'] or '',
                'кабинет': row_dict['room_number'] or '',
                'часы': row_dict['hours'],
                'время': PARA_TIMES.get(para, '')
            }
        conn.close()
        return result

    def get_schedule(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.id, s.day_of_week, s.para_number, s.hours,
                   g.name as group_name,
                   t.full_name as teacher_name,
                   sub.name as subject_name,
                   r.number as room_number
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.is_active = 1
            ORDER BY g.name, s.day_of_week, s.para_number
        ''')
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return rows

    def get_lesson(self, lesson_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.*, g.name as group_name, t.full_name as teacher_name,
                   sub.name as subject_name, r.number as room_number
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.id = ? AND s.is_active = 1
        ''', (lesson_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def add_lesson(self, group_id: int, day_of_week: str, para_number: int,
                   subject_id: int, teacher_id: Optional[int], room_id: Optional[int],
                   hours: int = 2, created_by: Optional[int] = None) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO schedule
            (group_id, day_of_week, para_number, subject_id, teacher_id, room_id, hours, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?)
        ''', (group_id, day_of_week, para_number, subject_id, teacher_id, room_id, hours, created_by))
        schedule_id = cursor.lastrowid
        conn.commit()
        cursor.execute('''
            INSERT INTO schedule_history (schedule_id, action, new_data, changed_by)
            VALUES (?, 'create', ?, ?)
        ''', (schedule_id, f"Группа ID:{group_id}, {day_of_week}, пара {para_number}", created_by))
        conn.commit()
        conn.close()
        return schedule_id

    def delete_lesson(self, schedule_id: int, deleted_by: Optional[int] = None):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE schedule SET is_active=0 WHERE id=?", (schedule_id,))
        cursor.execute('''
            INSERT INTO schedule_history (schedule_id, action, changed_by)
            VALUES (?, 'delete', ?)
        ''', (schedule_id, deleted_by))
        conn.commit()
        conn.close()

    def check_conflicts(self, group_id: int, day_of_week: str, para_number: int,
                        teacher_id: Optional[int], room_id: Optional[int],
                        lesson_id: Optional[int] = None) -> List[Dict[str, str]]:
        conflicts = []
        conn = self.get_connection()
        cursor = conn.cursor()
        if room_id:
            query = '''
                SELECT s.id, s.para_number, g.name as group_name,
                       sub.name as subject_name, r.number as room_number
                FROM schedule s
                LEFT JOIN groups g ON s.group_id = g.id
                LEFT JOIN subjects sub ON s.subject_id = sub.id
                LEFT JOIN rooms r ON s.room_id = r.id
                WHERE s.day_of_week = ? AND s.para_number = ? AND s.room_id = ?
                AND s.is_active = 1 AND s.group_id != ?
            '''
            params = [day_of_week, para_number, room_id, group_id]
            if lesson_id:
                query += " AND s.id != ?"
                params.append(lesson_id)
            cursor.execute(query, params)
            for row in cursor.fetchall():
                conflicts.append({
                    'type': 'Конфликт кабинета',
                    'description': f"Кабинет <strong>{row['room_number']}</strong> занят группой <strong>{row['group_name']}</strong>"
                })
        if teacher_id:
            query = '''
                SELECT s.id, s.para_number, g.name as group_name,
                       sub.name as subject_name, t.full_name as teacher_name
                FROM schedule s
                LEFT JOIN groups g ON s.group_id = g.id
                LEFT JOIN subjects sub ON s.subject_id = sub.id
                LEFT JOIN teachers t ON s.teacher_id = t.id
                WHERE s.day_of_week = ? AND s.para_number = ? AND s.teacher_id = ?
                AND s.is_active = 1 AND s.group_id != ?
            '''
            params = [day_of_week, para_number, teacher_id, group_id]
            if lesson_id:
                query += " AND s.id != ?"
                params.append(lesson_id)
            cursor.execute(query, params)
            for row in cursor.fetchall():
                conflicts.append({
                    'type': 'Конфликт преподавателя',
                    'description': f"<strong>{row['teacher_name']}</strong> ведёт пару у группы <strong>{row['group_name']}</strong>"
                })
        conn.close()
        return conflicts

    def mark_attendance(self, student_id: int, lesson_id: int, date: str,
                        status: str, comment: str = '', marked_by: Optional[int] = None) -> bool:
        if status not in ATTENDANCE_STATUSES:
            status = 'present'
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO attendance (student_id, lesson_id, date, status, comment, marked_by)
                VALUES (?, ?, ?, ?, ?, ?)
                ON CONFLICT(student_id, lesson_id, date)
                DO UPDATE SET status=excluded.status, comment=excluded.comment,
                              marked_by=excluded.marked_by, marked_at=CURRENT_TIMESTAMP
            ''', (student_id, lesson_id, date, status, comment, marked_by))
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Ошибка отметки: {e}")
            return False
        finally:
            conn.close()

    def auto_mark_present_for_lesson(self, lesson_id: int, date: str):
        with self.connect() as conn:
            cursor = conn.cursor()
            cursor.execute('''
                SELECT s.group_id, g.name as group_name
                FROM schedule s
                LEFT JOIN groups g ON s.group_id = g.id
                WHERE s.id = ?
            ''', (lesson_id,))
            lesson_info = cursor.fetchone()
            if not lesson_info:
                return
            group_name = lesson_info['group_name']

            cursor.execute('''
                INSERT INTO attendance (student_id, lesson_id, date, status, comment)
                SELECT u.id, ?, ?, 'present', ''
                FROM users u
                WHERE u.role IN ('student', 'viewer')
                  AND u.group_name = ?
                  AND u.is_active = 1
                  AND NOT EXISTS (
                      SELECT 1 FROM attendance a
                      WHERE a.student_id = u.id
                        AND a.lesson_id = ?
                        AND a.date = ?
                  )
            ''', (lesson_id, date, group_name, lesson_id, date))

    def get_attendance_stats_period(self, student_id: int, date_from: Optional[str] = None,
                                     date_to: Optional[str] = None) -> Dict[str, Any]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                COUNT(*) as total,
                SUM(CASE WHEN status='present' THEN 1 ELSE 0 END) as present,
                SUM(CASE WHEN status='late'    THEN 1 ELSE 0 END) as late,
                SUM(CASE WHEN status='absent'  THEN 1 ELSE 0 END) as absent
            FROM attendance WHERE student_id=?
        '''
        params = [student_id]
        if date_from:
            query += " AND date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND date <= ?"
            params.append(date_to)
        cursor.execute(query, params)
        row = cursor.fetchone()
        conn.close()
        if not row or not row['total']:
            return {'total': 0, 'present': 0, 'late': 0, 'absent': 0, 'percent': 0}
        total = row['total'] or 0
        present = row['present'] or 0
        late = row['late'] or 0
        absent = row['absent'] or 0
        percent = round((present + late) / total * 100, 1) if total else 0
        return {'total': total, 'present': present, 'late': late, 'absent': absent, 'percent': percent}

    def add_grade(self, student_id: int, lesson_id: int, date: str, grade: int,
                  grade_type: str = 'regular', comment: str = '',
                  marked_by: Optional[int] = None) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO grades (student_id, lesson_id, date, grade, grade_type, comment, marked_by)
                VALUES (?, ?, ?, ?, ?, ?, ?)
            ''', (student_id, lesson_id, date, grade, grade_type, comment, marked_by))
            conn.commit()
            return cursor.lastrowid
        except Exception as e:
            logger.error(f"Ошибка добавления оценки: {e}")
            return None
        finally:
            conn.close()

    def delete_grade(self, grade_id: int):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("DELETE FROM grades WHERE id=?", (grade_id,))
        conn.commit()
        conn.close()

    def get_grades_stats(self, student_id: int) -> Dict[str, Any]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                COUNT(*) as total,
                AVG(CAST(grade AS REAL)) as avg_grade,
                SUM(CASE WHEN grade=5 THEN 1 ELSE 0 END) as g5,
                SUM(CASE WHEN grade=4 THEN 1 ELSE 0 END) as g4,
                SUM(CASE WHEN grade=3 THEN 1 ELSE 0 END) as g3,
                SUM(CASE WHEN grade=2 THEN 1 ELSE 0 END) as g2
            FROM grades WHERE student_id=?
        ''', (student_id,))
        row = cursor.fetchone()
        conn.close()
        if not row or row['total'] == 0:
            return {'total': 0, 'avg': 0, 'g5': 0, 'g4': 0, 'g3': 0, 'g2': 0}
        return {
            'total': row['total'],
            'avg': round(row['avg_grade'] or 0, 2),
            'g5': row['g5'] or 0, 'g4': row['g4'] or 0,
            'g3': row['g3'] or 0, 'g2': row['g2'] or 0
        }

    def get_grades_for_lesson_date(self, lesson_id: int, date: str) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT g.*, u.full_name as student_name
            FROM grades g
            LEFT JOIN users u ON g.student_id = u.id
            WHERE g.lesson_id = ? AND g.date = ?
        ''', (lesson_id, date))
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def get_full_diary(self, student_id: int, date_from: Optional[str] = None,
                       date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                g.date,
                sub.name as subject_name,
                t.full_name as teacher_name,
                t.short_name as teacher_short,
                s.para_number,
                g.grade,
                g.grade_type,
                g.comment as grade_comment,
                a.status as attendance_status,
                a.comment as attendance_comment
            FROM grades g
            LEFT JOIN schedule s ON g.lesson_id = s.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            LEFT JOIN attendance a ON a.student_id = g.student_id
                                   AND a.lesson_id = g.lesson_id
                                   AND a.date = g.date
            WHERE g.student_id = ?
        '''
        params = [student_id]
        if date_from:
            query += " AND g.date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND g.date <= ?"
            params.append(date_to)
        query += " ORDER BY g.date DESC, s.para_number"
        cursor.execute(query, params)
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def get_attendance_only_diary(self, student_id: int, date_from: Optional[str] = None,
                                   date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                a.date, a.status, a.comment,
                sub.name as subject_name,
                t.full_name as teacher_name, t.short_name as teacher_short,
                s.para_number
            FROM attendance a
            LEFT JOIN schedule s ON a.lesson_id = s.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            WHERE a.student_id = ? AND a.status = 'absent'
        '''
        params = [student_id]
        if date_from:
            query += " AND a.date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND a.date <= ?"
            params.append(date_to)
        query += " ORDER BY a.date DESC"
        cursor.execute(query, params)
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def get_grades_by_subject_period(self, student_id: int, date_from: Optional[str] = None,
                                      date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                sub.name as subject_name,
                COUNT(g.id) as total,
                AVG(CAST(g.grade AS REAL)) as avg_grade,
                GROUP_CONCAT(g.grade, ',') as grades_list
            FROM grades g
            LEFT JOIN schedule s ON g.lesson_id = s.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            WHERE g.student_id = ?
        '''
        params = [student_id]
        if date_from:
            query += " AND g.date >= ?"
            params.append(date_from)
        if date_to:
            query += " AND g.date <= ?"
            params.append(date_to)
        query += " GROUP BY sub.name ORDER BY sub.name"
        cursor.execute(query, params)
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def get_student_diary_for_export(self, student_id: int,
                                     date_from: Optional[str] = None,
                                     date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                COALESCE(g.date, a.date)                AS date,
                s.day_of_week                            AS day_of_week,
                s.para_number                            AS para_number,
                sub.name                                 AS subject_name,
                t.full_name                              AS teacher_name,
                r.number                                 AS room_number,
                g.grade                                  AS grade,
                g.grade_type                             AS grade_type,
                g.comment                                AS grade_comment,
                a.status                                 AS attendance_status,
                a.comment                                AS attendance_comment
            FROM schedule s
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN teachers t   ON s.teacher_id = t.id
            LEFT JOIN rooms r      ON s.room_id = r.id
            LEFT JOIN grades g
                ON g.lesson_id = s.id AND g.student_id = ?
            LEFT JOIN attendance a
                ON a.lesson_id = s.id AND a.student_id = ?
                   AND (g.date IS NULL OR a.date = g.date)
            WHERE (g.id IS NOT NULL OR a.id IS NOT NULL)
        '''
        params: List[Any] = [student_id, student_id]
        if date_from:
            query += " AND COALESCE(g.date, a.date) >= ?"
            params.append(date_from)
        if date_to:
            query += " AND COALESCE(g.date, a.date) <= ?"
            params.append(date_to)
        query += " ORDER BY COALESCE(g.date, a.date) DESC, s.para_number"

        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def get_group_diary_for_export(self, group_name: str,
                                    date_from: Optional[str] = None,
                                    date_to: Optional[str] = None) -> Dict[str, Any]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, full_name, username
            FROM users
            WHERE role IN ('student', 'viewer')
              AND group_name = ?
              AND is_active = 1
            ORDER BY full_name
        ''', (group_name,))
        students = [dict(r) for r in cursor.fetchall()]
        conn.close()

        records_by_student: Dict[int, List[Dict[str, Any]]] = {}
        for st in students:
            records_by_student[st['id']] = self.get_student_diary_for_export(
                st['id'], date_from, date_to
            )

        return {
            'students': students,
            'records': records_by_student,
        }

    def get_calendar_month(self, year: int, month: int) -> Dict[str, Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        start_date = f"{year}-{month:02d}-01"
        if month == 12:
            end_date = f"{year + 1}-01-01"
        else:
            end_date = f"{year}-{month + 1:02d}-01"
        cursor.execute('''
            SELECT * FROM calendar
            WHERE date >= ? AND date < ?
            ORDER BY date
        ''', (start_date, end_date))
        calendar_data = {row['date']: dict(row) for row in cursor.fetchall()}
        conn.close()
        return calendar_data

    def generate_calendar(self, year: int):
        conn = self.get_connection()
        cursor = conn.cursor()
        from datetime import date, timedelta
        start = date(year, 1, 1)
        end = date(year, 12, 31)
        current = start
        while current <= end:
            day_of_week = DAYS_OF_WEEK[current.weekday()] if current.weekday() < 6 else 'Воскресенье'
            cursor.execute('''
                INSERT OR IGNORE INTO calendar (date, day_of_week)
                VALUES (?, ?)
            ''', (current.isoformat(), day_of_week))
            current += timedelta(days=1)
        conn.commit()
        conn.close()

    def get_schedule_for_date_with_overrides(self, date_str: str) -> Dict[str, Any]:
        from datetime import datetime
        date_obj = datetime.strptime(date_str, "%Y-%m-%d")
        day_of_week = DAYS_OF_WEEK[date_obj.weekday()] if date_obj.weekday() < 6 else 'Воскресенье'

        conn = self.get_connection()
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM calendar WHERE date=?", (date_str,))
        calendar_day = cursor.fetchone()
        if calendar_day and calendar_day['is_holiday']:
            conn.close()
            return {'holiday': True, 'name': calendar_day['holiday_name']}

        cursor.execute('''
            SELECT o.*,
                   nt.full_name as new_teacher_name, nt.short_name as new_teacher_short,
                   nr.number as new_room_number,
                   ns.name as new_subject_name,
                   g.name as group_name
            FROM schedule_overrides o
            LEFT JOIN teachers nt ON o.new_teacher_id = nt.id
            LEFT JOIN rooms nr ON o.new_room_id = nr.id
            LEFT JOIN subjects ns ON o.new_subject_id = ns.id
            LEFT JOIN groups g ON o.group_id = g.id
            WHERE o.date = ? AND o.is_active = 1
        ''', (date_str,))
        all_overrides = [dict(r) for r in cursor.fetchall()]

        base_overrides_map = {}
        added_list = []
        for ov in all_overrides:
            if ov['override_type'] == 'added':
                added_list.append(ov)
            elif ov['lesson_id']:
                base_overrides_map[ov['lesson_id']] = ov

        cursor.execute('''
            SELECT
                s.id, s.para_number, s.hours,
                g.name as group_name,
                t.full_name as teacher_name, t.short_name as teacher_short,
                sub.name as subject_name,
                r.number as room_number
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN teachers t ON s.teacher_id = t.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.day_of_week = ? AND s.is_active = 1
            ORDER BY g.name, s.para_number
        ''', (day_of_week,))

        result = {}
        for row in cursor.fetchall():
            row_dict = dict(row)
            group = row_dict['group_name']
            para = row_dict['para_number']
            lesson_id = row_dict['id']

            if group not in result:
                result[group] = {}

            lesson = {
                'id': lesson_id,
                'предмет': row_dict['subject_name'] or '',
                'преподаватель': row_dict['teacher_name'] or '',
                'преподаватель_коротко': row_dict['teacher_short'] or '',
                'кабинет': row_dict['room_number'] or '',
                'часы': row_dict['hours'],
                'время': PARA_TIMES.get(para, ''),
                'override': None,
            }

            ov = base_overrides_map.get(lesson_id)
            if ov:
                lesson['override'] = {
                    'id': ov['id'],
                    'type': ov['override_type'],
                    'comment': ov['comment'] or '',
                }
                if ov['override_type'] == 'cancelled':
                    lesson['отменено'] = True
                elif ov['override_type'] == 'replaced':
                    if ov['new_subject_name']:
                        lesson['предмет'] = ov['new_subject_name']
                        lesson['заменён_предмет'] = True
                    if ov['new_teacher_name']:
                        lesson['преподаватель'] = ov['new_teacher_name']
                        lesson['преподаватель_коротко'] = ov['new_teacher_short'] or ov['new_teacher_name']
                        lesson['заменён_преподаватель'] = True
                    if ov['new_room_number']:
                        lesson['кабинет'] = ov['new_room_number']
                        lesson['заменён_кабинет'] = True
                    if ov['new_hours']:
                        lesson['часы'] = ov['new_hours']

            result[group][para] = lesson

        for ov in added_list:
            group = ov.get('group_name')
            para = ov.get('para_number')
            if not group or not para:
                continue
            if group not in result:
                result[group] = {}
            result[group][para] = {
                'id': None,
                'override_id': ov['id'],
                'предмет': ov.get('new_subject_name') or '',
                'преподаватель': ov.get('new_teacher_name') or '',
                'преподаватель_коротко': ov.get('new_teacher_short') or '',
                'кабинет': ov.get('new_room_number') or '',
                'часы': ov.get('new_hours') or 2,
                'время': PARA_TIMES.get(para, ''),
                'override': {
                    'id': ov['id'],
                    'type': 'added',
                    'comment': ov['comment'] or '',
                },
                'добавлено': True,
            }

        conn.close()
        return result

    def get_schedule_for_date(self, date_str: str) -> Dict[str, Any]:
        return self.get_schedule_for_date_with_overrides(date_str)

    def get_schedule_for_date_editor(self, date_str: str) -> Dict[str, Any]:
        return self.get_schedule_for_date_with_overrides(date_str)

    def copy_schedule_day(self, source_day: str, target_day: str,
                          created_by: Optional[int] = None,
                          overwrite: bool = False) -> Dict[str, Any]:
        if source_day not in DAYS_OF_WEEK or target_day not in DAYS_OF_WEEK:
            return {'copied': 0, 'deleted': 0, 'error': 'Неверный день недели'}
        if source_day == target_day:
            return {'copied': 0, 'deleted': 0, 'error': 'Источник и цель совпадают'}

        conn = self.get_connection()
        cursor = conn.cursor()
        copied = 0
        deleted = 0

        try:
            cursor.execute('BEGIN IMMEDIATE')

            cursor.execute('''
                SELECT id, group_id, para_number, subject_id, teacher_id, room_id, hours
                FROM schedule
                WHERE day_of_week = ? AND is_active = 1
            ''', (source_day,))
            source_lessons = [dict(r) for r in cursor.fetchall()]

            if overwrite:
                cursor.execute('''
                    UPDATE schedule SET is_active = 0
                    WHERE day_of_week = ? AND is_active = 1
                ''', (target_day,))
                deleted = cursor.rowcount

            cursor.execute('''
                SELECT group_id, para_number, id FROM schedule
                WHERE day_of_week = ? AND is_active = 1
            ''', (target_day,))
            existing_map = {(r['group_id'], r['para_number']): r['id'] for r in cursor.fetchall()}

            to_insert = []
            to_update = []

            for lesson in source_lessons:
                key = (lesson['group_id'], lesson['para_number'])
                existing_id = existing_map.get(key)

                if existing_id:
                    if overwrite:
                        to_update.append((
                            lesson['subject_id'], lesson['teacher_id'],
                            lesson['room_id'], lesson['hours'], existing_id
                        ))
                        copied += 1
                    continue

                to_insert.append((
                    lesson['group_id'], target_day, lesson['para_number'],
                    lesson['subject_id'], lesson['teacher_id'],
                    lesson['room_id'], lesson['hours'], created_by
                ))
                copied += 1

            if to_update:
                cursor.executemany('''
                    UPDATE schedule SET subject_id=?, teacher_id=?, room_id=?, hours=?
                    WHERE id=?
                ''', to_update)

            if to_insert:
                cursor.executemany('''
                    INSERT INTO schedule
                    (group_id, day_of_week, para_number, subject_id, teacher_id, room_id, hours, created_by)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                ''', to_insert)

            cursor.execute('COMMIT')
        except Exception as e:
            try:
                cursor.execute('ROLLBACK')
            except Exception:
                pass
            logger.error(f"Ошибка копирования дня: {e}")
            return {'copied': 0, 'deleted': 0, 'error': str(e)}
        finally:
            conn.close()

        return {'copied': copied, 'deleted': deleted, 'error': None}

    def get_override_for_lesson_date(self, lesson_id: int, date: str) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT o.*,
                   nt.full_name as new_teacher_name,
                   nr.number as new_room_number,
                   ns.name as new_subject_name,
                   u.full_name as created_by_name
            FROM schedule_overrides o
            LEFT JOIN teachers nt ON o.new_teacher_id = nt.id
            LEFT JOIN rooms nr ON o.new_room_id = nr.id
            LEFT JOIN subjects ns ON o.new_subject_id = ns.id
            LEFT JOIN users u ON o.created_by = u.id
            WHERE o.lesson_id = ? AND o.date = ? AND o.is_active = 1
        ''', (lesson_id, date))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_overrides_for_date(self, date: str) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT o.*,
                   nt.full_name as new_teacher_name,
                   nr.number as new_room_number,
                   ns.name as new_subject_name,
                   u.full_name as created_by_name,
                   g.name as group_name
            FROM schedule_overrides o
            LEFT JOIN teachers nt ON o.new_teacher_id = nt.id
            LEFT JOIN rooms nr ON o.new_room_id = nr.id
            LEFT JOIN subjects ns ON o.new_subject_id = ns.id
            LEFT JOIN users u ON o.created_by = u.id
            LEFT JOIN groups g ON o.group_id = g.id
            WHERE o.date = ? AND o.is_active = 1
            ORDER BY o.id
        ''', (date,))
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def add_override(self, lesson_id: Optional[int], date: str, override_type: str,
                     new_teacher_id: Optional[int] = None,
                     new_room_id: Optional[int] = None,
                     new_subject_id: Optional[int] = None,
                     new_hours: Optional[int] = None,
                     comment: str = '',
                     created_by: Optional[int] = None,
                     group_id: Optional[int] = None,
                     para_number: Optional[int] = None) -> Optional[int]:
        if override_type not in ('cancelled', 'replaced', 'added'):
            return None
        if override_type in ('cancelled', 'replaced') and not lesson_id:
            return None
        if override_type == 'added' and (not group_id or not para_number):
            return None

        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            if override_type in ('cancelled', 'replaced'):
                cursor.execute('''
                    UPDATE schedule_overrides SET is_active = 0
                    WHERE lesson_id = ? AND date = ? AND is_active = 1
                ''', (lesson_id, date))
            elif override_type == 'added':
                cursor.execute('''
                    UPDATE schedule_overrides SET is_active = 0
                    WHERE group_id = ? AND date = ? AND para_number = ?
                      AND override_type = 'added' AND is_active = 1
                ''', (group_id, date, para_number))

            cursor.execute('''
                INSERT INTO schedule_overrides
                (lesson_id, date, group_id, para_number, override_type,
                 new_teacher_id, new_room_id, new_subject_id, new_hours,
                 comment, created_by, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
            ''', (lesson_id, date, group_id, para_number, override_type,
                  new_teacher_id, new_room_id, new_subject_id, new_hours,
                  comment, created_by))

            conn.commit()
            override_id = cursor.lastrowid
            return override_id
        except Exception as e:
            logger.error(f"Ошибка создания override: {e}")
            conn.rollback()
            return None
        finally:
            conn.close()

    def delete_override(self, override_id: int, deleted_by: Optional[int] = None) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('UPDATE schedule_overrides SET is_active = 0 WHERE id = ?', (override_id,))
            conn.commit()
            return True
        except Exception as e:
            logger.error(f"Ошибка удаления override: {e}")
            return False
        finally:
            conn.close()

    def get_teacher_stats(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT t.id, t.full_name, t.short_name,
                COUNT(s.id) as lessons_count,
                COALESCE(SUM(s.hours), 0) as total_hours
            FROM teachers t
            LEFT JOIN schedule s ON t.id = s.teacher_id AND s.is_active = 1
            GROUP BY t.id
            ORDER BY total_hours DESC
        ''')
        stats = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return stats

    def get_group_stats(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT g.id, g.name, g.specialty,
                COUNT(s.id) as lessons_count,
                COALESCE(SUM(s.hours), 0) as total_hours
            FROM groups g
            LEFT JOIN schedule s ON g.id = s.group_id AND s.is_active = 1
            GROUP BY g.id
            ORDER BY g.name
        ''')
        stats = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return stats

    def get_overall_stats(self) -> Dict[str, int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT COUNT(*) FROM groups")
        groups_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM teachers")
        teachers_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM subjects")
        subjects_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*) FROM rooms")
        rooms_count = cursor.fetchone()[0]
        cursor.execute("SELECT COUNT(*), COALESCE(SUM(hours), 0) FROM schedule WHERE is_active=1")
        row = cursor.fetchone()
        lessons_count = row[0]
        total_hours = row[1]
        cursor.execute("SELECT COUNT(*) FROM users WHERE role IN ('student', 'viewer') AND is_active=1")
        students_count = cursor.fetchone()[0]
        conn.close()
        return {
            'groups': groups_count,
            'teachers': teachers_count,
            'subjects': subjects_count,
            'rooms': rooms_count,
            'lessons': lessons_count,
            'hours': total_hours,
            'viewers': students_count,
            'students': students_count
        }

    def get_lesson_students_with_marks(self, lesson_id: int, date: str,
                                        auto_mark_present: bool = True) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.group_id, g.name as group_name
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            WHERE s.id = ?
        ''', (lesson_id,))
        lesson_info = cursor.fetchone()
        if not lesson_info:
            conn.close()
            return []
        group_name = lesson_info['group_name']

        cursor.execute('''
            SELECT id, full_name, username FROM users
            WHERE role IN ('student', 'viewer') AND group_name=? AND is_active=1
            ORDER BY full_name
        ''', (group_name,))
        students = [dict(row) for row in cursor.fetchall()]

        cursor.execute('SELECT * FROM grades WHERE lesson_id=? AND date=?', (lesson_id, date))
        grades_map = {row['student_id']: dict(row) for row in cursor.fetchall()}

        cursor.execute('SELECT * FROM attendance WHERE lesson_id=? AND date=?', (lesson_id, date))
        att_map = {row['student_id']: dict(row) for row in cursor.fetchall()}

        if auto_mark_present:
            for st in students:
                if st['id'] not in att_map:
                    cursor.execute('''
                        INSERT INTO attendance (student_id, lesson_id, date, status, comment)
                        VALUES (?, ?, ?, 'present', '')
                    ''', (st['id'], lesson_id, date))
                    att_map[st['id']] = {
                        'student_id': st['id'], 'lesson_id': lesson_id,
                        'date': date, 'status': 'present', 'comment': ''
                    }
            conn.commit()

        for st in students:
            st['grade'] = grades_map.get(st['id'], {}).get('grade')
            st['grade_id'] = grades_map.get(st['id'], {}).get('id')
            st['grade_comment'] = grades_map.get(st['id'], {}).get('comment', '')
            st['attendance'] = att_map.get(st['id'], {}).get('status', 'present')
            st['attendance_comment'] = att_map.get(st['id'], {}).get('comment', '')

        conn.close()
        return students

    def get_all_busy_slots(self) -> Dict[str, Any]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT group_id, teacher_id, day_of_week, para_number
            FROM schedule
            WHERE is_active = 1
        ''')
        result = {'group': {}, 'teacher': {}}
        for row in cursor.fetchall():
            gid = row['group_id']
            tid = row['teacher_id']
            key = (row['day_of_week'], int(row['para_number']))

            if gid:
                result['group'].setdefault(gid, {})[key] = True
            if tid:
                result['teacher'].setdefault(tid, {})[key] = True
        conn.close()
        return result

    def auto_fill_schedule(self, group_id: Optional[int] = None,
                            semester: Optional[str] = None,
                            academic_year: Optional[str] = None) -> Dict[str, Any]:
        rows = self.get_curriculum_with_fact(
            group_id=group_id,
            semester=semester,
            academic_year=academic_year,
        )
        if not rows:
            return {
                'placed': [],
                'conflicts': [],
                'stats': {'total_requests': 0, 'total_placed': 0, 'total_conflicts': 0}
            }

        busy = self.get_all_busy_slots()

        try:
            import database_ortools
            if database_ortools.is_available():
                logger.info("Используется OR-Tools (CP-SAT)")
                result = database_ortools.solve_schedule(
                    db_file=self.db_file,
                    curriculum_rows=rows,
                    existing_busy=busy,
                    max_pairs_per_day=6,
                    time_limit_sec=30,
                )
                if result and not result.get('error'):
                    return result
                logger.warning(f"OR-Tools вернул ошибку: {result.get('error')}, откат на жадный")
        except ImportError:
            logger.debug("database_ortools не найден — используется жадный алгоритм")
        except Exception as e:
            logger.warning(f"Ошибка OR-Tools: {e}, откат на жадный")

        logger.info("Используется жадный алгоритм")
        busy_group = busy['group']
        busy_teacher = busy['teacher']

        days = list(DAYS_OF_WEEK)
        paras = [1, 2, 3, 4, 5, 6]

        rows_sorted = sorted(rows, key=lambda r: (r.get('priority', 2), -(r.get('pairs_per_week') or 0)))

        placed: List[Dict[str, Any]] = []
        conflicts: List[Dict[str, Any]] = []

        for r in rows_sorted:
            gid = r['group_id']
            sid = r['subject_id']
            tid = r.get('teacher_id')
            ppr = r.get('pairs_per_week') or 0
            if ppr <= 0:
                continue

            placed_count = 0
            for day in days:
                if placed_count >= ppr:
                    break
                for para in paras:
                    if placed_count >= ppr:
                        break
                    key = (day, para)

                    if busy_group.get(gid, {}).get(key):
                        continue
                    if tid and busy_teacher.get(tid, {}).get(key):
                        continue

                    placed.append({
                        'curriculum_id': r['id'],
                        'group_id': gid,
                        'group_name': r.get('group_name') or '',
                        'subject_id': sid,
                        'subject_name': r.get('subject_name') or '',
                        'teacher_id': tid,
                        'teacher_name': r.get('teacher_name') or '',
                        'teacher_short': r.get('teacher_short') or '',
                        'day_of_week': day,
                        'para_number': para,
                        'hours': 2,
                    })
                    busy_group.setdefault(gid, {})[key] = True
                    if tid:
                        busy_teacher.setdefault(tid, {})[key] = True
                    placed_count += 1

            if placed_count < ppr:
                reason = f"Удалось разместить {placed_count} из {ppr} пар"
                conflicts.append({
                    'curriculum_id': r['id'],
                    'group_id': gid,
                    'group_name': r.get('group_name') or '',
                    'subject_name': r.get('subject_name') or '',
                    'teacher_name': r.get('teacher_name') or '',
                    'placed_count': placed_count,
                    'wanted': ppr,
                    'reason': reason,
                })

        return {
            'placed': placed,
            'conflicts': conflicts,
            'stats': {
                'total_requests': len(rows),
                'total_placed': len(placed),
                'total_conflicts': len(conflicts),
                'solver_status': 'greedy',
            },
        }

    def apply_auto_fill(self, placed: List[Dict[str, Any]],
                         created_by: Optional[int] = None) -> Dict[str, int]:
        if not placed:
            return {'inserted': 0, 'skipped': 0}

        inserted = 0
        skipped = 0

        with self.connect() as conn:
            cursor = conn.cursor()
            for item in placed:
                gid = item.get('group_id')
                sid = item.get('subject_id')
                tid = item.get('teacher_id')
                day = item.get('day_of_week')
                para = item.get('para_number')
                hours = item.get('hours', 2)

                if not (gid and sid and day and para):
                    skipped += 1
                    continue

                cursor.execute('''
                    SELECT id FROM schedule
                    WHERE group_id = ? AND day_of_week = ? AND para_number = ? AND is_active = 1
                ''', (gid, day, para))
                if cursor.fetchone():
                    skipped += 1
                    continue

                if tid:
                    cursor.execute('''
                        SELECT id FROM schedule
                        WHERE teacher_id = ? AND day_of_week = ? AND para_number = ? AND is_active = 1
                    ''', (tid, day, para))
                    if cursor.fetchone():
                        skipped += 1
                        continue

                cursor.execute('''
                    INSERT INTO schedule
                    (group_id, day_of_week, para_number, subject_id, teacher_id, room_id, hours, created_by)
                    VALUES (?, ?, ?, ?, ?, NULL, ?, ?)
                ''', (gid, day, para, sid, tid, hours, created_by))
                inserted += 1

        return {'inserted': inserted, 'skipped': skipped}

    def get_teacher_full_stats(self, teacher_id: int, date_from=None, date_to=None) -> Dict[str, int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT
                COUNT(*) as total_lessons,
                COALESCE(SUM(s.hours), 0) as total_hours,
                COUNT(DISTINCT s.group_id) as total_groups,
                COUNT(DISTINCT s.subject_id) as total_subjects
            FROM schedule s
            WHERE s.teacher_id = ? AND s.is_active = 1
        ''', (teacher_id,))
        row = cursor.fetchone()
        conn.close()
        return {
            'total_lessons': row['total_lessons'] or 0,
            'total_hours': row['total_hours'] or 0,
            'total_groups': row['total_groups'] or 0,
            'total_subjects': row['total_subjects'] or 0
        }

    def get_teacher_groups_list(self, teacher_id: int) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT DISTINCT g.id, g.name, g.specialty
            FROM schedule s
            JOIN groups g ON s.group_id = g.id
            WHERE s.teacher_id = ? AND s.is_active = 1
            ORDER BY g.name
        ''', (teacher_id,))
        groups = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return groups

    def get_teacher_subjects_list(self, teacher_id: int) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT DISTINCT sub.id, sub.name
            FROM schedule s
            JOIN subjects sub ON s.subject_id = sub.id
            WHERE s.teacher_id = ? AND s.is_active = 1
            ORDER BY sub.name
        ''', (teacher_id,))
        subjects = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return subjects

    def get_teacher_schedule(self, teacher_id: int, day_of_week: str) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.id as lesson_id, s.para_number, s.hours,
                   g.id as group_id, g.name as group_name,
                   sub.id as subject_id, sub.name as subject_name,
                   r.number as room_number
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.teacher_id = ? AND s.day_of_week = ? AND s.is_active = 1
            ORDER BY s.para_number
        ''', (teacher_id, day_of_week))
        lessons = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return lessons

    def get_teacher_schedule_full(self, teacher_id: int) -> Dict[str, Dict[int, Dict[str, Any]]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.id, s.day_of_week, s.para_number, s.hours,
                   g.name as group_name,
                   sub.name as subject_name,
                   r.number as room_number
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.teacher_id = ? AND s.is_active = 1
            ORDER BY s.para_number
        ''', (teacher_id,))
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()
        result = {}
        for day in DAYS_OF_WEEK:
            result[day] = {}
        for r in rows:
            day = r['day_of_week']
            para = r['para_number']
            if day not in result:
                result[day] = {}
            result[day][para] = {
                'id': r['id'],
                'предмет': r['subject_name'] or '',
                'группа': r['group_name'] or '',
                'кабинет': r['room_number'] or '',
                'часы': r['hours'],
                'время': PARA_TIMES.get(para, '')
            }
        return result

    def get_teacher_lessons_today(self, teacher_id: int) -> List[Dict[str, Any]]:
        today = datetime.now()
        weekday = today.weekday()
        if weekday >= 6:
            return []
        day_name = DAYS_OF_WEEK[weekday]
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT s.id, s.para_number, s.hours,
                   g.name as group_name,
                   sub.name as subject_name,
                   r.number as room_number
            FROM schedule s
            LEFT JOIN groups g ON s.group_id = g.id
            LEFT JOIN subjects sub ON s.subject_id = sub.id
            LEFT JOIN rooms r ON s.room_id = r.id
            WHERE s.teacher_id = ? AND s.day_of_week = ? AND s.is_active = 1
            ORDER BY s.para_number
        ''', (teacher_id, day_name))
        lessons = [dict(row) for row in cursor.fetchall()]
        conn.close()
        for l in lessons:
            l['время'] = PARA_TIMES.get(l['para_number'], '')
        return lessons

    def get_teacher_avg_grade(self, teacher_id: int, date_from=None, date_to=None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                sub.name as subject_name,
                g.name as group_name,
                COUNT(gr.id) as grades_count,
                AVG(CAST(gr.grade AS REAL)) as avg_grade
            FROM schedule s
            JOIN subjects sub ON s.subject_id = sub.id
            JOIN groups g ON s.group_id = g.id
            LEFT JOIN grades gr ON gr.lesson_id = s.id
            WHERE s.teacher_id = ? AND s.is_active = 1
        '''
        params = [teacher_id]
        if date_from:
            query += " AND (gr.date >= ? OR gr.date IS NULL)"
            params.append(date_from)
        if date_to:
            query += " AND (gr.date <= ? OR gr.date IS NULL)"
            params.append(date_to)
        query += '''
            GROUP BY sub.name, g.name
            HAVING COUNT(gr.id) > 0
            ORDER BY sub.name, g.name
        '''
        cursor.execute(query, params)
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()
        for r in rows:
            r['avg_grade'] = round(r['avg_grade'] or 0, 2)
        return rows

    def get_teacher_attendance_by_group(self, teacher_id: int, date_from=None, date_to=None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        query = '''
            SELECT
                g.name as group_name,
                COUNT(a.id) as total_marks,
                SUM(CASE WHEN a.status='present' THEN 1 ELSE 0 END) as present,
                SUM(CASE WHEN a.status='late'    THEN 1 ELSE 0 END) as late,
                SUM(CASE WHEN a.status='absent'  THEN 1 ELSE 0 END) as absent
            FROM schedule s
            JOIN groups g ON s.group_id = g.id
            LEFT JOIN attendance a ON a.lesson_id = s.id
            WHERE s.teacher_id = ? AND s.is_active = 1
        '''
        params = [teacher_id]
        if date_from:
            query += " AND (a.date >= ? OR a.date IS NULL)"
            params.append(date_from)
        if date_to:
            query += " AND (a.date <= ? OR a.date IS NULL)"
            params.append(date_to)
        query += '''
            GROUP BY g.name
            HAVING COUNT(a.id) > 0
            ORDER BY g.name
        '''
        cursor.execute(query, params)
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()
        for r in rows:
            total = r['total_marks'] or 0
            present = r['present'] or 0
            late = r['late'] or 0
            r['present'] = present
            r['late'] = late
            r['absent'] = r['absent'] or 0
            r['percent'] = round((present + late) / total * 100, 1) if total else 0
        return rows

    def get_teacher_rating(self, date_from=None, date_to=None, min_grades: int = 1) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()

        where_rating = ""
        params_rating: List[Any] = []
        if date_from or date_to:
            where_rating = " WHERE 1=1"
            if date_from:
                where_rating += " AND (g.date >= ? OR g.date IS NULL)"
                params_rating.append(date_from)
            if date_to:
                where_rating += " AND (g.date <= ? OR g.date IS NULL)"
                params_rating.append(date_to)

        query = f'''
            SELECT
                t.id as teacher_id,
                t.full_name,
                t.short_name,
                COUNT(g.id) as grades_count,
                AVG(CAST(g.grade AS REAL)) as avg_grade,
                COUNT(DISTINCT g.student_id) as students_count,
                SUM(CASE WHEN g.grade=5 THEN 1 ELSE 0 END) as g5,
                SUM(CASE WHEN g.grade=4 THEN 1 ELSE 0 END) as g4,
                SUM(CASE WHEN g.grade=3 THEN 1 ELSE 0 END) as g3,
                SUM(CASE WHEN g.grade=2 THEN 1 ELSE 0 END) as g2,
                COALESCE((
                    SELECT COUNT(*) FROM schedule s
                    WHERE s.teacher_id = t.id AND s.is_active = 1
                ), 0) as lessons_count,
                COALESCE((
                    SELECT SUM(s.hours) FROM schedule s
                    WHERE s.teacher_id = t.id AND s.is_active = 1
                ), 0) as total_hours,
                COALESCE((
                    SELECT COUNT(DISTINCT s.group_id) FROM schedule s
                    WHERE s.teacher_id = t.id AND s.is_active = 1
                ), 0) as groups_count
            FROM teachers t
            JOIN schedule s ON s.teacher_id = t.id AND s.is_active = 1
            LEFT JOIN grades g ON g.lesson_id = s.id
            {where_rating}
            GROUP BY t.id
            HAVING COUNT(g.id) >= ?
        '''
        params_rating.append(min_grades)

        cursor.execute(query, params_rating)
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()

        for r in rows:
            r['avg_grade'] = round(r['avg_grade'] or 0, 2)
            r['grades_count'] = r['grades_count'] or 0
            r['students_count'] = r['students_count'] or 0
            r['g5'] = r['g5'] or 0
            r['g4'] = r['g4'] or 0
            r['g3'] = r['g3'] or 0
            r['g2'] = r['g2'] or 0
            r['lessons_count'] = r['lessons_count'] or 0
            r['total_hours'] = r['total_hours'] or 0
            r['groups_count'] = r['groups_count'] or 0
            total = r['grades_count']
            quality = r['g5'] + r['g4']
            r['quality_percent'] = round(quality / total * 100, 1) if total else 0

        rows.sort(key=lambda x: (-x['avg_grade'], -x['grades_count']))
        return rows

    def get_students_rating(self, date_from=None, date_to=None,
                            group_name: Optional[str] = None,
                            min_grades: int = 1) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()

        query = '''
            SELECT
                u.id            AS student_id,
                u.full_name,
                u.username,
                u.group_name,
                COUNT(g.id)     AS grades_count,
                AVG(CAST(g.grade AS REAL)) AS avg_grade,
                SUM(CASE WHEN g.grade = 5 THEN 1 ELSE 0 END) AS g5,
                SUM(CASE WHEN g.grade = 4 THEN 1 ELSE 0 END) AS g4,
                SUM(CASE WHEN g.grade = 3 THEN 1 ELSE 0 END) AS g3,
                SUM(CASE WHEN g.grade = 2 THEN 1 ELSE 0 END) AS g2
            FROM users u
            LEFT JOIN grades g ON g.student_id = u.id
        '''
        params: List[Any] = []
        where = ["u.role IN ('student', 'viewer')", "u.is_active = 1"]

        if group_name:
            where.append("u.group_name = ?")
            params.append(group_name)
        if date_from:
            where.append("(g.date >= ? OR g.date IS NULL)")
            params.append(date_from)
        if date_to:
            where.append("(g.date <= ? OR g.date IS NULL)")
            params.append(date_to)

        query += " WHERE " + " AND ".join(where)
        query += " GROUP BY u.id HAVING COUNT(g.id) >= ?"
        params.append(min_grades)
        query += " ORDER BY avg_grade DESC, grades_count DESC"

        cursor.execute(query, params)
        rows = [dict(row) for row in cursor.fetchall()]
        conn.close()

        for r in rows:
            r['avg_grade']    = round(r['avg_grade'] or 0, 2)
            r['grades_count'] = r['grades_count'] or 0
            r['g5']           = r['g5'] or 0
            r['g4']           = r['g4'] or 0
            r['g3']           = r['g3'] or 0
            r['g2']           = r['g2'] or 0
            total   = r['grades_count']
            quality = r['g5'] + r['g4']
            r['quality_percent'] = round(quality / total * 100, 1) if total else 0

        return rows

    def get_group_performance(self, group_name: str, date_from: Optional[str] = None,
                              date_to: Optional[str] = None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()

        cursor.execute('''
            SELECT id, full_name, username FROM users
            WHERE role IN ('student', 'viewer') AND group_name=? AND is_active=1
            ORDER BY full_name
        ''', (group_name,))
        students = [dict(row) for row in cursor.fetchall()]
        if not students:
            conn.close()
            return []

        student_ids = [s['id'] for s in students]
        placeholders = ','.join('?' * len(student_ids))

        grade_query = f'''
            SELECT
                student_id,
                COUNT(*) as total,
                AVG(CAST(grade AS REAL)) as avg_grade,
                SUM(CASE WHEN grade=5 THEN 1 ELSE 0 END) as g5,
                SUM(CASE WHEN grade=4 THEN 1 ELSE 0 END) as g4,
                SUM(CASE WHEN grade=3 THEN 1 ELSE 0 END) as g3,
                SUM(CASE WHEN grade=2 THEN 1 ELSE 0 END) as g2
            FROM grades
            WHERE student_id IN ({placeholders})
        '''
        grade_params: List[Any] = list(student_ids)
        if date_from:
            grade_query += " AND date >= ?"
            grade_params.append(date_from)
        if date_to:
            grade_query += " AND date <= ?"
            grade_params.append(date_to)
        grade_query += " GROUP BY student_id"

        cursor.execute(grade_query, grade_params)
        grades_map = {r['student_id']: dict(r) for r in cursor.fetchall()}

        att_query = f'''
            SELECT
                student_id,
                COUNT(*) as total,
                SUM(CASE WHEN status='present' THEN 1 ELSE 0 END) as present,
                SUM(CASE WHEN status='late'    THEN 1 ELSE 0 END) as late,
                SUM(CASE WHEN status='absent'  THEN 1 ELSE 0 END) as absent
            FROM attendance
            WHERE student_id IN ({placeholders})
        '''
        att_params: List[Any] = list(student_ids)
        if date_from:
            att_query += " AND date >= ?"
            att_params.append(date_from)
        if date_to:
            att_query += " AND date <= ?"
            att_params.append(date_to)
        att_query += " GROUP BY student_id"

        cursor.execute(att_query, att_params)
        att_map = {r['student_id']: dict(r) for r in cursor.fetchall()}

        conn.close()

        for st in students:
            g_row = grades_map.get(st['id'], {})
            a_row = att_map.get(st['id'], {})

            st['grades_total'] = g_row.get('total') or 0
            st['avg_grade'] = round(g_row.get('avg_grade') or 0, 2)
            st['g5'] = g_row.get('g5') or 0
            st['g4'] = g_row.get('g4') or 0
            st['g3'] = g_row.get('g3') or 0
            st['g2'] = g_row.get('g2') or 0
            st['att_total'] = a_row.get('total') or 0
            st['att_present'] = a_row.get('present') or 0
            st['att_late'] = a_row.get('late') or 0
            st['att_absent'] = a_row.get('absent') or 0
            st['att_percent'] = round((st['att_present'] + st['att_late']) / st['att_total'] * 100, 1) if st['att_total'] else 0

        return students

    TRANSLIT = {
        'А': 'a', 'Б': 'b', 'В': 'v', 'Г': 'g', 'Д': 'd', 'Е': 'e', 'Ё': 'e',
        'Ж': 'zh', 'З': 'z', 'И': 'i', 'Й': 'y', 'К': 'k', 'Л': 'l', 'М': 'm',
        'Н': 'n', 'О': 'o', 'П': 'p', 'Р': 'r', 'С': 's', 'Т': 't', 'У': 'u',
        'Ф': 'f', 'Х': 'h', 'Ц': 'c', 'Ч': 'ch', 'Ш': 'sh', 'Щ': 'sch',
        'Ъ': '', 'Ы': 'y', 'Ь': '', 'Э': 'e', 'Ю': 'yu', 'Я': 'ya',
        'A': 'a', 'B': 'b', 'C': 'c', 'D': 'd', 'E': 'e', 'F': 'f', 'G': 'g',
        'H': 'h', 'I': 'i', 'J': 'j', 'K': 'k', 'L': 'l', 'M': 'm', 'N': 'n',
        'O': 'o', 'P': 'p', 'Q': 'q', 'R': 'r', 'S': 's', 'T': 't', 'U': 'u',
        'V': 'v', 'W': 'w', 'X': 'x', 'Y': 'y', 'Z': 'z',
    }

    def make_group_prefix(self, group_name: str) -> str:
        if not group_name:
            return 'grp'
        result = []
        for ch in group_name:
            if ch.isdigit():
                result.append(ch)
            elif ch in self.TRANSLIT:
                result.append(self.TRANSLIT[ch])
            elif ch.isalpha():
                result.append(ch.lower())
        prefix = ''.join(result).lower()
        return prefix if prefix else 'grp'

    def get_or_create_group_prefix(self, group_name: str) -> str:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT prefix FROM group_prefixes WHERE group_name=?", (group_name,))
        row = cursor.fetchone()
        if row:
            conn.close()
            return row['prefix']
        prefix = self.make_group_prefix(group_name)
        base_prefix = prefix
        counter = 1
        while True:
            cursor.execute("SELECT COUNT(*) FROM group_prefixes WHERE prefix=?", (prefix,))
            if cursor.fetchone()[0] == 0:
                break
            counter += 1
            prefix = f"{base_prefix}{counter}"
        cursor.execute("INSERT INTO group_prefixes (group_name, prefix) VALUES (?, ?)",
                       (group_name, prefix))
        conn.commit()
        conn.close()
        return prefix

    def generate_unique_username(self, base: str, max_attempts: int = 100) -> Optional[str]:
        for _ in range(max_attempts):
            suffix = ''.join(str(secrets.randbelow(10)) for _ in range(3))
            username = f"{base}{suffix}"
            if not self.user_exists(username):
                return username
        for _ in range(max_attempts):
            suffix = ''.join(str(secrets.randbelow(10)) for _ in range(6))
            username = f"{base}{suffix}"
            if not self.user_exists(username):
                return username
        return None

    def generate_unique_parent_username(self, student_login: str) -> Optional[str]:
        base = f"parent_{student_login}"
        if len(base) > 40:
            base = base[:40]
        if not self.user_exists(base):
            return base
        for _ in range(100):
            suffix = ''.join(str(secrets.randbelow(10)) for _ in range(2))
            username = f"{base}{suffix}"
            if not self.user_exists(username):
                return username
        return None

    def generate_student_password(self) -> str:
        digits = ''.join(str(secrets.randbelow(10)) for _ in range(8))
        return f"Student{digits}"

    def generate_parent_password(self) -> str:
        digits = ''.join(str(secrets.randbelow(10)) for _ in range(7))
        return f"parent{digits}"

    def generate_random_password(self, length: int = 10) -> str:
        chars = 'abcdefghijkmnpqrstuvwxyzABCDEFGHJKLMNPQRSTUVWXYZ23456789'
        return ''.join(chars[secrets.randbelow(len(chars))] for _ in range(length))

    def save_credentials(self, user_id, username, password_plain, role,
                         full_name='', group_name='', linked_student_id=None,
                         linked_student_login='', created_by=None):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            INSERT INTO credentials_log
            (user_id, username, password_plain, role, full_name, group_name,
             linked_student_id, linked_student_login, created_by)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
        ''', (user_id, username, password_plain, role, full_name, group_name,
              linked_student_id, linked_student_login, created_by))
        conn.commit()
        conn.close()

    def get_all_credentials(self, role_filter=None) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        if role_filter:
            cursor.execute('''
                SELECT cl.*, u.is_active as user_active
                FROM credentials_log cl
                LEFT JOIN users u ON cl.user_id = u.id
                WHERE cl.role = ?
                ORDER BY cl.created_at DESC
            ''', (role_filter,))
        else:
            cursor.execute('''
                SELECT cl.*, u.is_active as user_active
                FROM credentials_log cl
                LEFT JOIN users u ON cl.user_id = u.id
                ORDER BY cl.created_at DESC
            ''')
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records
    def get_users_for_export(self, role_filter: Optional[str] = None) -> List[Dict[str, Any]]:
        """
        Возвращает список пользователей для экспорта:
        студенты с привязанными родителями, отдельно преподаватели/редакторы/админы.
        Пароли берутся из credentials_log.
        """
        conn = self.get_connection()
        cursor = conn.cursor()

        # Все credentials для быстрого доступа
        cursor.execute('''
            SELECT user_id, username, password_plain, role,
                   linked_student_id, linked_student_login
            FROM credentials_log
            ORDER BY id DESC
        ''')
        creds_by_user = {}
        creds_by_linked = {}
        for row in cursor.fetchall():
            r = dict(row)
            uid = r['user_id']
            if uid and uid not in creds_by_user:
                creds_by_user[uid] = r
            linked = r.get('linked_student_id')
            if linked:
                creds_by_linked.setdefault(linked, []).append(r)

        # Все активные пользователи
        query = '''
            SELECT id, username, full_name, role, group_name, parent_of,
                   teacher_id, is_active, last_login
            FROM users
            WHERE is_active = 1
        '''
        params: List[Any] = []
        if role_filter:
            query += " AND role = ?"
            params.append(role_filter)
        query += " ORDER BY role, group_name, full_name"

        cursor.execute(query, params)
        users = [dict(r) for r in cursor.fetchall()]
        conn.close()

        # Индексируем студентов по username для привязки родителей
        students_by_login = {u['username']: u for u in users if u['role'] in ('student', 'viewer')}

        # Собираем строки
        result = []
        used_parent_ids = set()

        # 1) Студенты + привязанные родители
        for u in users:
            if u['role'] not in ('student', 'viewer'):
                continue

            student_cred = creds_by_user.get(u['id'], {})
            student_login = student_cred.get('username') or u['username']
            student_password = student_cred.get('password_plain') or ''

            # Находим родителя
            parent_login = ''
            parent_password = ''
            parent_full_name = ''
            parent_found = None
            for p in users:
                if p['role'] != 'parent':
                    continue
                parent_of_list = [x.strip() for x in (p.get('parent_of') or '').split(',') if x.strip()]
                if u['username'] in parent_of_list:
                    parent_found = p
                    break

            if parent_found:
                used_parent_ids.add(parent_found['id'])
                pcred = creds_by_user.get(parent_found['id'], {})
                parent_login = pcred.get('username') or parent_found['username']
                parent_password = pcred.get('password_plain') or ''
                parent_full_name = parent_found.get('full_name') or ''

            result.append({
                'student_login': student_login,
                'student_password': student_password,
                'student_name': u.get('full_name') or '',
                'group_name': u.get('group_name') or '',
                'parent_login': parent_login,
                'parent_password': parent_password,
                'parent_name': parent_full_name,
                'role': u['role'],
                'user_id': u['id'],
            })

        # 2) Остальные (преподаватели, редакторы, админы, "бесхозные" родители)
        for u in users:
            if u['role'] in ('student', 'viewer'):
                continue
            if u['id'] in used_parent_ids:
                continue

            cred = creds_by_user.get(u['id'], {})
            login = cred.get('username') or u['username']
            password = cred.get('password_plain') or ''

            result.append({
                'student_login': '',
                'student_password': '',
                'student_name': '',
                'group_name': '',
                'parent_login': login,
                'parent_password': password,
                'parent_name': u.get('full_name') or '',
                'role': u['role'],
                'user_id': u['id'],
            })

        # Сортируем: сначала студенты по группе, потом остальные
        def sort_key(r):
            return (0 if r['role'] in ('student', 'viewer') else 1,
                    r.get('group_name') or 'zzz',
                    r.get('student_name') or r.get('parent_name') or '')

        result.sort(key=sort_key)
        return result

    def search_credentials(self, query: str) -> List[Dict[str, Any]]:
        if not query:
            return []
        q = f"%{query.lower()}%"
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT cl.*, u.is_active as user_active
            FROM credentials_log cl
            LEFT JOIN users u ON cl.user_id = u.id
            WHERE LOWER(cl.username) LIKE ?
               OR LOWER(cl.full_name) LIKE ?
               OR LOWER(cl.group_name) LIKE ?
               OR LOWER(cl.linked_student_login) LIKE ?
            ORDER BY cl.role, cl.full_name
            LIMIT 200
        ''', (q, q, q, q))
        records = [dict(row) for row in cursor.fetchall()]
        conn.close()
        return records

    def parse_students_excel(self, data_list):
        students = []
        errors = []
        seen = set()
        for i, row in enumerate(data_list, 1):
            full_name = str(row.get('ФИО', '') or row.get('ФИО студента', '') or '').strip()
            group_name = str(row.get('Группа', '') or row.get('Группа (код)', '') or '').strip()
            if not full_name:
                errors.append(f"Строка {i}: пустое ФИО")
                continue
            if not group_name:
                errors.append(f"Строка {i}: пустая группа ({full_name})")
                continue
            key = (full_name.lower(), group_name.lower())
            if key in seen:
                errors.append(f"Строка {i}: дубликат ({full_name}, {group_name})")
                continue
            seen.add(key)
            students.append({'full_name': full_name, 'group_name': group_name})
        return students, errors

    def generate_students_and_parents(self, students, created_by=None):
        created = []
        errors = []
        for st in students:
            full_name = st['full_name']
            group_name = st['group_name']
            try:
                conn = self.get_connection()
                cursor = conn.cursor()
                cursor.execute("SELECT id FROM groups WHERE name=?", (group_name,))
                if not cursor.fetchone():
                    cursor.execute("INSERT INTO groups (name) VALUES (?)", (group_name,))
                    conn.commit()
                conn.close()
                prefix = self.get_or_create_group_prefix(group_name)
                student_login = self.generate_unique_username(prefix)
                if not student_login:
                    errors.append(f"{full_name}: не удалось создать логин")
                    continue
                student_password = self.generate_student_password()
                student_id = self.create_user(
                    username=student_login, password=student_password,
                    full_name=full_name, role='student', group_name=group_name
                )
                if not student_id:
                    errors.append(f"{full_name}: ошибка создания студента")
                    continue
                self.save_credentials(student_id, student_login, student_password,
                                      'student', full_name, group_name, created_by=created_by)
                parent_login = self.generate_unique_parent_username(student_login)
                parent_password = None
                if parent_login:
                    parent_password = self.generate_parent_password()
                    parent_id = self.create_user(
                        username=parent_login, password=parent_password,
                        full_name=f"Родитель {full_name}", role='parent',
                        parent_of=student_login
                    )
                    if parent_id:
                        self.save_credentials(
                            parent_id, parent_login, parent_password, 'parent',
                            f"Родитель {full_name}", group_name,
                            linked_student_id=student_id,
                            linked_student_login=student_login,
                            created_by=created_by
                        )
                created.append({
                    'full_name': full_name, 'group_name': group_name,
                    'student_login': student_login, 'student_password': student_password,
                    'parent_login': parent_login, 'parent_password': parent_password,
                })
            except Exception as e:
                errors.append(f"{full_name}: {e}")
        return {'created': created, 'errors': errors}

    def reset_user_password(self, user_id: int, new_password: str):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("UPDATE users SET password_hash=? WHERE id=?",
                       (self.hash_password(new_password), user_id))
        conn.commit()
        conn.close()

    def update_credentials_password(self, user_id: int, new_password_plain: str):
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('UPDATE credentials_log SET password_plain=? WHERE user_id=?',
                       (new_password_plain, user_id))
        conn.commit()
        conn.close()
    def get_group_by_id(self, group_id: int) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM groups WHERE id=?", (group_id,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def get_group_by_name(self, name: str) -> Optional[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute("SELECT * FROM groups WHERE name=?", (name,))
        row = cursor.fetchone()
        conn.close()
        return dict(row) if row else None

    def upsert_group(self, name: str, **fields) -> Optional[int]:
        allowed = {
            'specialty_code', 'specialty_name', 'current_course',
            'admission_year', 'education_form', 'department',
            'curator', 'students_count', 'is_active'
        }
        data = {k: v for k, v in fields.items() if k in allowed}

        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("SELECT id FROM groups WHERE name=?", (name,))
            existing = cursor.fetchone()
            if existing:
                if data:
                    updates = ", ".join(f"{k}=?" for k in data.keys())
                    params = list(data.values()) + [name]
                    cursor.execute(f"UPDATE groups SET {updates} WHERE name=?", params)
                conn.commit()
                return existing['id']
            else:
                cols = ['name'] + list(data.keys())
                vals = [name] + list(data.values())
                placeholders = ','.join('?' * len(cols))
                cursor.execute(
                    f"INSERT INTO groups ({','.join(cols)}) VALUES ({placeholders})",
                    vals
                )
                conn.commit()
                return cursor.lastrowid
        except Exception as e:
            logger.error(f"upsert_group error: {e}")
            conn.rollback()
            return None
        finally:
            conn.close()

    def promote_groups_if_needed(self, new_academic_year: int) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                SELECT id, name, current_course, admission_year
                FROM groups WHERE is_active = 1
            ''')
            groups = [dict(r) for r in cursor.fetchall()]
            promoted = 0
            for g in groups:
                new_course = g['current_course'] + 1
                if new_course > 4:
                    cursor.execute("UPDATE groups SET is_active=0 WHERE id=?", (g['id'],))
                    logger.info(f"🎓 Группа {g['name']} выпустилась")
                else:
                    cursor.execute("UPDATE groups SET current_course=? WHERE id=?", (new_course, g['id']))
                    logger.info(f"⬆️ Группа {g['name']}: {g['current_course']} → {new_course} курс")
                    promoted += 1
            conn.commit()
            return promoted
        except Exception as e:
            logger.error(f"promote_groups error: {e}")
            conn.rollback()
            return 0
        finally:
            conn.close()

    def get_active_groups_for_schedule(self) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT id, name, current_course, admission_year,
                   specialty_code, specialty_name, students_count
            FROM groups
            WHERE is_active = 1
              AND current_course BETWEEN 1 AND 4
              AND (education_form IS NULL OR education_form = 'Очная')
            ORDER BY current_course, name
        ''')
        groups = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return groups

    def get_curriculum_plan_v2(self, group_id: int,
                                course: Optional[int] = None,
                                semester_in_course: Optional[int] = None
                                ) -> List[Dict[str, Any]]:
        query = '''
            SELECT cp.*,
                   s.name AS subject_name,
                   t.full_name AS teacher_name,
                   t.short_name AS teacher_short
            FROM curriculum_plan_v2 cp
            LEFT JOIN subjects s ON cp.subject_id = s.id
            LEFT JOIN teachers t ON cp.teacher_id = t.id
            WHERE cp.group_id = ?
        '''
        params = [group_id]
        if course:
            query += " AND cp.course = ?"
            params.append(course)
        if semester_in_course:
            query += " AND cp.semester_in_course = ?"
            params.append(semester_in_course)
        query += " ORDER BY cp.course, cp.semester_in_course, s.name"

        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def add_curriculum_plan_v2(self, group_id: int, subject_id: int,
                                course: int, semester_in_course: int,
                                global_semester: int,
                                total_hours: int = 0,
                                lectures: int = 0, lab: int = 0,
                                practical: int = 0, seminars: int = 0,
                                individual: int = 0, group_lessons: int = 0,
                                subgroup_lessons: int = 0, self_work: int = 0,
                                control: int = 0, independent: int = 0,
                                contact: int = 0,
                                pairs_per_week: int = 0, weeks: int = 16,
                                index_code: str = '', registry_number: str = '',
                                control_forms: str = '', block: str = '',
                                teacher_id: Optional[int] = None,
                                priority: int = 2, comment: str = '',
                                created_by: Optional[int] = None
                                ) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO curriculum_plan_v2
                (group_id, subject_id, teacher_id,
                 course, semester_in_course, global_semester,
                 total_hours, lectures, lab, practical, seminars,
                 individual, group_lessons, subgroup_lessons,
                 self_work, control, independent, contact,
                 pairs_per_week, weeks,
                 index_code, registry_number, control_forms, block,
                 priority, comment, created_by)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?,
                        ?, ?, ?, ?, ?, ?, ?)
                ON CONFLICT(group_id, subject_id, course, semester_in_course)
                DO UPDATE SET
                    total_hours=excluded.total_hours,
                    lectures=excluded.lectures,
                    lab=excluded.lab,
                    practical=excluded.practical,
                    seminars=excluded.seminars,
                    individual=excluded.individual,
                    group_lessons=excluded.group_lessons,
                    subgroup_lessons=excluded.subgroup_lessons,
                    self_work=excluded.self_work,
                    control=excluded.control,
                    independent=excluded.independent,
                    contact=excluded.contact,
                    pairs_per_week=excluded.pairs_per_week,
                    weeks=excluded.weeks,
                    index_code=excluded.index_code,
                    registry_number=excluded.registry_number,
                    control_forms=excluded.control_forms,
                    block=excluded.block,
                    priority=excluded.priority,
                    comment=excluded.comment
            ''', (
                group_id, subject_id, teacher_id,
                course, semester_in_course, global_semester,
                total_hours, lectures, lab, practical, seminars,
                individual, group_lessons, subgroup_lessons,
                self_work, control, independent, contact,
                pairs_per_week, weeks,
                index_code, registry_number, control_forms, block,
                priority, comment, created_by
            ))
            conn.commit()
            return cursor.lastrowid
        except Exception as e:
            logger.error(f"add_curriculum_plan_v2 error: {e}")
            conn.rollback()
            return None
        finally:
            conn.close()

    def delete_curriculum_plan_v2(self, group_id: int) -> int:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute("DELETE FROM curriculum_plan_v2 WHERE group_id=?", (group_id,))
            conn.commit()
            return cursor.rowcount
        except Exception as e:
            logger.error(f"delete_curriculum_plan_v2 error: {e}")
            conn.rollback()
            return 0
        finally:
            conn.close()

    def get_curriculum_summary(self, group_id: int) -> Dict[str, Any]:
        """Сводка по группе: плановые часы по курсам и семестрам."""
        conn = self.get_connection()
        cursor = conn.cursor()

        cursor.execute('''
            SELECT course, semester_in_course,
                   COUNT(DISTINCT subject_id) as disciplines,
                   SUM(total_hours) as total_hours,
                   SUM(pairs_per_week) as total_pairs_per_week
            FROM curriculum_plan_v2
            WHERE group_id = ?
            GROUP BY course, semester_in_course
            ORDER BY course, semester_in_course
        ''', (group_id,))
        rows = [dict(r) for r in cursor.fetchall()]

        cursor.execute('''
            SELECT COUNT(DISTINCT subject_id) as uniq_disciplines,
                   COUNT(*) as total_records,
                   COALESCE(SUM(total_hours), 0) as total_hours
            FROM curriculum_plan_v2
            WHERE group_id = ?
        ''', (group_id,))
        overall = dict(cursor.fetchone())

        conn.close()
        return {
            'by_semester': rows,
            'total_disciplines': overall['uniq_disciplines'] or 0,
            'total_records': overall['total_records'] or 0,
            'total_hours': overall['total_hours'] or 0,
        }
    # ==================== ЛИЧНЫЕ РАСПИСАНИЯ (v3) ====================

    def get_teacher_personal_schedule(self, teacher_id: int,
                                       semester_num: Optional[int] = None,
                                       academic_year: Optional[str] = None
                                       ) -> List[Dict[str, Any]]:
        query = '''
            SELECT tps.*,
                   u1.full_name AS student1_name,
                   u2.full_name AS student2_name,
                   g.name AS group_name,
                   s.name AS subject_name,
                   r.number AS room_number
            FROM teacher_personal_schedule tps
            LEFT JOIN users u1 ON tps.student1_id = u1.id
            LEFT JOIN users u2 ON tps.student2_id = u2.id
            LEFT JOIN groups g ON tps.group_id = g.id
            LEFT JOIN subjects s ON tps.subject_id = s.id
            LEFT JOIN rooms r ON tps.room_id = r.id
            WHERE tps.teacher_id = ? AND tps.is_active = 1
        '''
        params = [teacher_id]
        if semester_num:
            query += " AND tps.semester_num = ?"
            params.append(semester_num)
        if academic_year:
            query += " AND tps.academic_year = ?"
            params.append(academic_year)
        query += " ORDER BY tps.day_of_week, tps.para_number"

        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    def get_all_personal_schedules(self, semester_num: Optional[int] = None,
                                    academic_year: Optional[str] = None
                                    ) -> List[Dict[str, Any]]:
        query = '''
            SELECT tps.*,
                   t.full_name AS teacher_name,
                   t.short_name AS teacher_short,
                   u1.full_name AS student1_name,
                   u2.full_name AS student2_name,
                   g.name AS group_name,
                   s.name AS subject_name,
                   r.number AS room_number
            FROM teacher_personal_schedule tps
            LEFT JOIN teachers t ON tps.teacher_id = t.id
            LEFT JOIN users u1 ON tps.student1_id = u1.id
            LEFT JOIN users u2 ON tps.student2_id = u2.id
            LEFT JOIN groups g ON tps.group_id = g.id
            LEFT JOIN subjects s ON tps.subject_id = s.id
            LEFT JOIN rooms r ON tps.room_id = r.id
            WHERE tps.is_active = 1
        '''
        params: List[Any] = []
        if semester_num:
            query += " AND tps.semester_num = ?"
            params.append(semester_num)
        if academic_year:
            query += " AND tps.academic_year = ?"
            params.append(academic_year)
        query += " ORDER BY t.full_name, tps.day_of_week, tps.para_number"

        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows
    def get_all_personal_schedules_with_teacher(self,
                                                  semester_num: Optional[int] = None,
                                                  academic_year: Optional[str] = None,
                                                  teacher_id: Optional[int] = None,
                                                  day_of_week: Optional[str] = None,
                                                  group_id: Optional[int] = None
                                                  ) -> List[Dict[str, Any]]:
        """Полный список личных расписаний с фильтрами."""
        query = '''
            SELECT tps.*,
                   t.full_name AS teacher_name,
                   t.short_name AS teacher_short,
                   u1.full_name AS student1_name,
                   u2.full_name AS student2_name,
                   g.name AS group_name,
                   s.name AS subject_name,
                   r.number AS room_number
            FROM teacher_personal_schedule tps
            LEFT JOIN teachers t ON tps.teacher_id = t.id
            LEFT JOIN users u1 ON tps.student1_id = u1.id
            LEFT JOIN users u2 ON tps.student2_id = u2.id
            LEFT JOIN groups g ON tps.group_id = g.id
            LEFT JOIN subjects s ON tps.subject_id = s.id
            LEFT JOIN rooms r ON tps.room_id = r.id
            WHERE tps.is_active = 1
        '''
        params: List[Any] = []
        if semester_num:
            query += " AND tps.semester_num = ?"
            params.append(semester_num)
        if academic_year:
            query += " AND tps.academic_year = ?"
            params.append(academic_year)
        if teacher_id:
            query += " AND tps.teacher_id = ?"
            params.append(teacher_id)
        if day_of_week:
            query += " AND tps.day_of_week = ?"
            params.append(day_of_week)
        if group_id:
            query += " AND tps.group_id = ?"
            params.append(group_id)
        query += " ORDER BY t.short_name, tps.day_of_week, tps.para_number"

        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute(query, params)
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows
    def add_personal_schedule(self, teacher_id: int, day_of_week: str,
                               para_number: int,
                               student1_id: Optional[int] = None,
                               student2_id: Optional[int] = None,
                               students_raw: str = '',
                               group_id: Optional[int] = None,
                               subject_id: Optional[int] = None,
                               room_id: Optional[int] = None,
                               lesson_type: str = 'individual',
                               semester_num: int = 1,
                               academic_year: str = '2025/2026',
                               note: str = '',
                               created_by: Optional[int] = None
                               ) -> Optional[int]:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute('''
                INSERT INTO teacher_personal_schedule
                (teacher_id, day_of_week, para_number,
                 student1_id, student2_id, students_raw,
                 group_id, subject_id, room_id,
                 lesson_type, semester_num, academic_year,
                 note, created_by, is_active)
                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?, 1)
            ''', (teacher_id, day_of_week, para_number,
                  student1_id, student2_id, students_raw,
                  group_id, subject_id, room_id,
                  lesson_type, semester_num, academic_year,
                  note, created_by))
            conn.commit()
            return cursor.lastrowid
        except Exception as e:
            logger.error(f"add_personal_schedule error: {e}")
            conn.rollback()
            return None
        finally:
            conn.close()

    def update_personal_schedule(self, entry_id: int, **fields) -> bool:
        allowed = {
            'teacher_id', 'day_of_week', 'para_number',
            'student1_id', 'student2_id', 'students_raw',
            'group_id', 'subject_id', 'room_id',
            'lesson_type', 'semester_num', 'academic_year', 'note'
        }
        data = {k: v for k, v in fields.items() if k in allowed}
        if not data:
            return False

        updates = ", ".join(f"{k}=?" for k in data.keys())
        params = list(data.values()) + [entry_id]

        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                f"UPDATE teacher_personal_schedule SET {updates}, updated_at=CURRENT_TIMESTAMP WHERE id=?",
                params
            )
            conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            logger.error(f"update_personal_schedule error: {e}")
            conn.rollback()
            return False
        finally:
            conn.close()

    def delete_personal_schedule(self, entry_id: int) -> bool:
        conn = self.get_connection()
        cursor = conn.cursor()
        try:
            cursor.execute(
                "UPDATE teacher_personal_schedule SET is_active=0 WHERE id=?",
                (entry_id,)
            )
            conn.commit()
            return cursor.rowcount > 0
        except Exception as e:
            logger.error(f"delete_personal_schedule error: {e}")
            return False
        finally:
            conn.close()

    def get_students_in_group(self, group_id: int) -> List[Dict[str, Any]]:
        conn = self.get_connection()
        cursor = conn.cursor()
        cursor.execute('''
            SELECT u.id, u.full_name, u.username
            FROM users u
            JOIN groups g ON u.group_name = g.name
            WHERE g.id = ? AND u.role IN ('student', 'viewer') AND u.is_active = 1
            ORDER BY u.full_name
        ''', (group_id,))
        rows = [dict(r) for r in cursor.fetchall()]
        conn.close()
        return rows

    # ==================== КОНЕЦ ЛИЧНЫХ РАСПИСАНИЙ ====================
db = Database()

if __name__ == '__main__':
    logger.info("База данных инициализирована")