"""
Резервное копирование базы данных.
- При запуске
- Раз в сутки (через фоновый поток)
"""
import shutil
import sqlite3
import threading
import time
from datetime import datetime, timedelta
from pathlib import Path

from config import config
from logger import get_logger

logger = get_logger(__name__)


def make_backup(reason='manual'):
    """
    Создаёт резервную копию БД через sqlite3.backup().
    Это безопаснее, чем shutil.copy, если БД используется.
    """
    db_path = Path(config.DB_FILE)
    if not db_path.exists():
        logger.warning(f"⚠️ БД {db_path} не найдена, бэкап пропущен")
        return None

    timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
    backup_name = f"schedule_{reason}_{timestamp}.db"
    backup_path = config.BACKUPS_PATH / backup_name

    try:
        src = sqlite3.connect(str(db_path))
        dst = sqlite3.connect(str(backup_path))
        with dst:
            src.backup(dst)
        dst.close()
        src.close()

        logger.info(f"💾 Бэкап создан: {backup_name}")
        _rotate_backups()
        return backup_path
    except Exception as e:
        logger.error(f"❌ Ошибка бэкапа: {e}")
        return None


def _rotate_backups():
    """Удаляет старые бэкапы, оставляя последние BACKUP_KEEP штук."""
    try:
        backups = sorted(
            config.BACKUPS_PATH.glob('schedule_*.db'),
            key=lambda p: p.stat().st_mtime,
            reverse=True
        )
        for old in backups[config.BACKUP_KEEP:]:
            old.unlink()
            logger.info(f"🗑 Удалён старый бэкап: {old.name}")
    except Exception as e:
        logger.error(f"❌ Ошибка ротации бэкапов: {e}")


def list_backups():
    """Список существующих бэкапов (для UI)."""
    if not config.BACKUPS_PATH.exists():
        return []
    backups = sorted(
        config.BACKUPS_PATH.glob('schedule_*.db'),
        key=lambda p: p.stat().st_mtime,
        reverse=True
    )
    return [
        {
            'name': p.name,
            'size': p.stat().st_size,
            'mtime': datetime.fromtimestamp(p.stat().st_mtime).strftime('%Y-%m-%d %H:%M:%S'),
        }
        for p in backups
    ]


# ==================== ФОНОВЫЙ ПЛАНИРОВЩИК ====================

_scheduler_thread = None
_stop_event = threading.Event()


def _scheduler_loop():
    """Каждый час проверяет: если сегодня ещё не было бэкапа — делает его."""
    logger.info("⏰ Планировщик бэкапов запущен")
    last_backup_date = None

    while not _stop_event.is_set():
        now = datetime.now()
        today = now.date()

        # Делаем бэкап раз в сутки в 03:00
        if last_backup_date != today and now.hour >= 3:
            make_backup(reason='daily')
            last_backup_date = today

        # Спим час, но проверяем stop_event каждые 60 сек
        for _ in range(60):
            if _stop_event.is_set():
                return
            time.sleep(60)

    logger.info("⏰ Планировщик бэкапов остановлен")


def start_backup_scheduler():
    """Запускает фоновый планировщик."""
    global _scheduler_thread
    if _scheduler_thread and _scheduler_thread.is_alive():
        return
    _stop_event.clear()
    _scheduler_thread = threading.Thread(target=_scheduler_loop, daemon=True)
    _scheduler_thread.start()


def stop_backup_scheduler():
    """Останавливает планировщик."""
    _stop_event.set()