"""Константы приложения"""

from datetime import datetime

MONTH_NAMES = {
    1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
    5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
    9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
}

DEFAULT_EMAIL_BODY = (
    "Добрый день!\n\n"
    "Сообщение сформировано автоматически, отвечать на него не нужно.\n\n"
    "Пароль от файлов в скором времени будет направлен вашему Куратору"
)


def get_previous_month_subject() -> str:
    """
    Генерирует тему письма с прошлым месяцем.
    
    Returns:
        Строка темы в формате "Расчетные листы <Месяц> <Год>"
        
    Пример:
        Если сейчас Март 2026 -> "Расчетные листы Февраль 2026"
        Если сейчас Январь 2026 -> "Расчетные листы Декабрь 2025"
    """
    today = datetime.now()
    # Вычисляем предыдущий месяц
    if today.month == 1:
        # Если январь, то предыдущий месяц - декабрь прошлого года
        previous_month = 12
        previous_year = today.year - 1
    else:
        # В остальных случаях просто уменьшаем месяц на 1
        previous_month = today.month - 1
        previous_year = today.year
    
    month_name = MONTH_NAMES[previous_month]
    return f"Расчетные листы {month_name} {previous_year}"

# Колонки Excel
REQUIRED_COLUMNS = ['email', 'файл_01', 'файл_02', 'файл_03']

# Настройки по умолчанию
DEFAULT_THREAD_COUNT = 3
MIN_THREAD_COUNT = 1
MAX_THREAD_COUNT = 10
DEFAULT_WINDOW_WIDTH = 720
DEFAULT_WINDOW_HEIGHT = 800
