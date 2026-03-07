"""Менеджер настроек приложения"""

import json
import os
from typing import Dict, Any
from pathlib import Path
import logging

logger = logging.getLogger("email_sender")


class SettingsManager:
    """Менеджер для сохранения и загрузки настроек"""
    
    def __init__(self, settings_file: str = "settings.json"):
        """
        Инициализация менеджера настроек.
        
        Args:
            settings_file: Путь к файлу настроек
        """
        self.settings_file = settings_file
    
    def save(self, settings: Dict[str, Any]) -> bool:
        """
        Сохраняет настройки в файл.
        
        Args:
            settings: Словарь с настройками
            
        Returns:
            True если успешно
        """
        try:
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
            
            logger.info("Настройки сохранены")
            return True
            
        except Exception as e:
            logger.error(f"Ошибка сохранения настроек: {str(e)}")
            return False
    
    def load(self) -> Dict[str, Any]:
        """
        Загружает настройки из файла.
        
        Returns:
            Словарь с настройками
        """
        if not os.path.exists(self.settings_file):
            return {}
        
        try:
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            logger.info("Настройки загружены")
            return settings
            
        except Exception as e:
            logger.error(f"Ошибка загрузки настроек: {str(e)}")
            return {}
    
    def exists(self) -> bool:
        """
        Проверяет существование файла настроек.
        
        Returns:
            True если файл существует
        """
        return os.path.exists(self.settings_file)
