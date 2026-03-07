# Core модуль - базовые компоненты и утилиты
from .logger_config import setup_logger
from .constants import MONTH_NAMES, DEFAULT_EMAIL_BODY

__all__ = ['setup_logger', 'MONTH_NAMES', 'DEFAULT_EMAIL_BODY']
