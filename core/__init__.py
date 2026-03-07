# Core модуль - базовые компоненты и утилиты
from .logger_config import setup_logger
from .constants import (
    MONTH_NAMES, 
    DEFAULT_EMAIL_BODY, 
    get_previous_month_subject,
    REQUIRED_COLUMNS,
    DEFAULT_THREAD_COUNT,
    DEFAULT_WINDOW_WIDTH,
    DEFAULT_WINDOW_HEIGHT
)

__all__ = [
    'setup_logger', 
    'MONTH_NAMES', 
    'DEFAULT_EMAIL_BODY',
    'get_previous_month_subject',
    'REQUIRED_COLUMNS',
    'DEFAULT_THREAD_COUNT',
    'DEFAULT_WINDOW_WIDTH',
    'DEFAULT_WINDOW_HEIGHT'
]
