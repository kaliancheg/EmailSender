"""Сервис для работы с файлами"""

import os
from pathlib import Path
from typing import Optional
import logging

logger = logging.getLogger("email_sender")


class FileService:
    """Сервис для поиска и управления файлами"""
    
    @staticmethod
    def find_file_in_folder(folder_path: str, filename: str) -> Optional[str]:
        """
        Ищет файл в указанной папке.
        
        Args:
            folder_path: Путь к папке
            filename: Имя файла для поиска
            
        Returns:
            Полный путь к файлу или None, если не найден
        """
        if not folder_path or not filename:
            return None
        
        folder = Path(folder_path)
        if not folder.exists():
            logger.warning(f"Папка не существует: {folder_path}")
            return None
        
        for file in folder.iterdir():
            if file.is_file() and file.name.lower() == str(filename).lower():
                return str(file)
        
        return None
    
    @staticmethod
    def validate_folder(folder_path: str) -> tuple[bool, str]:
        """
        Проверяет существование папки.
        
        Args:
            folder_path: Путь к папке
            
        Returns:
            Кортеж (успех, сообщение)
        """
        if not folder_path:
            return False, "Путь к папке не указан"
        
        folder = Path(folder_path)
        if not folder.exists():
            return False, f"Папка не существует: {folder_path}"
        
        if not folder.is_dir():
            return False, f"Указанный путь не является папкой: {folder_path}"
        
        return True, "Папка существует"
    
    @staticmethod
    def get_files_count(folder_path: str) -> int:
        """
        Подсчитывает количество файлов в папке.
        
        Args:
            folder_path: Путь к папке
            
        Returns:
            Количество файлов
        """
        if not folder_path:
            return 0
        
        folder = Path(folder_path)
        if not folder.exists():
            return 0
        
        return sum(1 for f in folder.iterdir() if f.is_file())
