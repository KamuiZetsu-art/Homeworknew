"""
Конфигурация приложения. Читает переменные из .env
"""
import os
import sys
from pathlib import Path
from dotenv import load_dotenv

# Загружаем .env из корня проекта
BASE_DIR = Path(__file__).resolve().parent
load_dotenv(BASE_DIR / '.env')


class Config:
    # Безопасность
    SECRET_KEY = os.environ.get('SECRET_KEY')

    # База данных
    DB_FILE = os.environ.get('DB_FILE', 'schedule.db')

    # Бэкапы
    BACKUP_DIR = os.environ.get('BACKUP_DIR', 'backups')
    BACKUP_KEEP = int(os.environ.get('BACKUP_KEEP', '10'))

    # Логирование
    LOG_LEVEL = os.environ.get('LOG_LEVEL', 'INFO').upper()
    LOG_FILE = os.environ.get('LOG_FILE', 'logs/app.log')

    # Rate limiting
    LOGIN_RATE_LIMIT = os.environ.get('LOGIN_RATE_LIMIT', '5 per minute')

    # Flask
    HOST = os.environ.get('HOST', '0.0.0.0')
    PORT = int(os.environ.get('PORT', '5000'))

    # Папки
    BASE_DIR = BASE_DIR
    BACKUPS_PATH = BASE_DIR / BACKUP_DIR
    LOGS_PATH = BASE_DIR / 'logs'

    # ===== НОВЫЕ ПАРАМЕТРЫ =====
    CURRENT_ACADEMIC_YEAR = int(os.environ.get('CURRENT_ACADEMIC_YEAR', '2025'))
    DEFAULT_WEEKS_PER_SEMESTER = int(os.environ.get('DEFAULT_WEEKS_PER_SEMESTER', '16'))
    MAX_PAIRS_PER_DAY = int(os.environ.get('MAX_PAIRS_PER_DAY', '6'))

    @classmethod
    def ensure_dirs(cls):
        """Создаёт нужные папки, если их нет."""
        cls.BACKUPS_PATH.mkdir(exist_ok=True)
        cls.LOGS_PATH.mkdir(exist_ok=True)


config = Config()
config.ensure_dirs()


def _check_required():
    """
    Проверяет, что критичные переменные заданы.
    Вызывается на старте приложения (в app.py) — не при импорте конфига,
    чтобы тесты/скрипты могли импортировать config без .env.
    """
    if not config.SECRET_KEY:
        print()
        print("=" * 60)
        print("❌ ОШИБКА КОНФИГУРАЦИИ: не задан SECRET_KEY")
        print("=" * 60)
        print()
        print("Что сделать:")
        print(f"  1. Создайте файл: {config.BASE_DIR / '.env'}")
        print("  2. Добавьте в него строку:")
        print("       SECRET_KEY=ваш-секретный-ключ")
        print()
        print("  Где взять ключ? Сгенерируйте так (Python):")
        print("       python -c \"import secrets; print(secrets.token_urlsafe(48))\"")
        print()
        print("=" * 60)
        sys.exit(1)


if __name__ == '__main__':
    _check_required()
    print("✅ Конфигурация корректна")