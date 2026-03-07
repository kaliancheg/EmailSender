"""Настройка логирования"""

import logging
from pathlib import Path


def setup_logger(log_file: str = "email_sender.log", level: int = logging.INFO) -> logging.Logger:
    """
    Настраивает и возвращает логгер.
    
    Args:
        log_file: Имя файла для логирования
        level: Уровень логирования
        
    Returns:
        Настроенный логгер
    """
    logger = logging.getLogger("email_sender")
    logger.setLevel(level)
    
    # Очищаем существующие обработчики
    logger.handlers.clear()
    
    # Форматтер
    formatter = logging.Formatter(
        '%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        datefmt='%H:%M:%S'
    )
    
    # Файловый обработчик
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(level)
    file_handler.setFormatter(formatter)
    logger.addHandler(file_handler)
    
    # Консольный обработчик
    console_handler = logging.StreamHandler()
    console_handler.setLevel(level)
    console_handler.setFormatter(formatter)
    logger.addHandler(console_handler)
    
    return logger
