"""
Настройка логирования. Логи пишутся в файл + в консоль.
"""
import logging
import sys
from logging.handlers import RotatingFileHandler
from config import config


def setup_logging():
    """Инициализирует корневой логгер."""
    level = getattr(logging, config.LOG_LEVEL, logging.INFO)

    # Формат
    formatter = logging.Formatter(
        '[%(asctime)s] [%(levelname)s] [%(name)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )

    # Корневой логгер
    root = logging.getLogger()
    root.setLevel(level)

    # Убираем старые хендлеры (на случай повторного вызова)
    for h in list(root.handlers):
        root.removeHandler(h)

    # Файл с ротацией (5 МБ × 5 файлов)
    file_handler = RotatingFileHandler(
        config.LOG_FILE,
        maxBytes=5 * 1024 * 1024,
        backupCount=5,
        encoding='utf-8'
    )
    file_handler.setFormatter(formatter)
    root.addHandler(file_handler)

    # Консоль
    console_handler = logging.StreamHandler(sys.stdout)
    console_handler.setFormatter(formatter)
    root.addHandler(console_handler)

    # Приглушаем werkzeug (слишком болтливый)
    logging.getLogger('werkzeug').setLevel(logging.WARNING)

    logging.info(f"📝 Логирование инициализировано (level={config.LOG_LEVEL})")


def get_logger(name):
    """Получить логгер для модуля."""
    return logging.getLogger(name)