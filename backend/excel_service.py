"""Сервис для работы с Excel файлами"""

import pandas as pd
from typing import List, Dict
import logging

from models.email_data import EmailRecipient
from core.constants import REQUIRED_COLUMNS

logger = logging.getLogger("email_sender")


class ExcelService:
    """Сервис для чтения и обработки Excel файлов"""
    
    @staticmethod
    def read_recipients(file_path: str) -> List[EmailRecipient]:
        """
        Читает данные из Excel файла и возвращает список получателей.
        
        Args:
            file_path: Путь к Excel файлу
            
        Returns:
            Список получателей
            
        Raises:
            ValueError: Если отсутствуют обязательные колонки
            FileNotFoundError: Если файл не найден
        """
        try:
            df = pd.read_excel(file_path)
            
            # Проверяем наличие обязательных колонок
            missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
            
            if missing_columns:
                raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing_columns)}")
            
            # Преобразуем в список объектов
            recipients = []
            for _, row in df.iterrows():
                recipient = EmailRecipient(
                    email=str(row.get('email', '')).strip(),
                    file_01=ExcelService._get_safe_value(row.get('файл_01')),
                    file_02=ExcelService._get_safe_value(row.get('файл_02')),
                    file_03=ExcelService._get_safe_value(row.get('файл_03'))
                )
                recipients.append(recipient)
            
            logger.info(f"Прочитано {len(recipients)} записей из Excel")
            return recipients
            
        except FileNotFoundError:
            logger.error(f"Файл не найден: {file_path}")
            raise
        except Exception as e:
            logger.error(f"Ошибка чтения Excel: {str(e)}")
            raise
    
    @staticmethod
    def _get_safe_value(value) -> str:
        """
        Безопасно получает значение, обрабатывая NaN.
        
        Args:
            value: Значение
            
        Returns:
            Строка или пустая строка
        """
        if pd.isna(value):
            return ""
        return str(value).strip()
    
    @staticmethod
    def validate_file(file_path: str) -> tuple[bool, str]:
        """
        Проверяет корректность Excel файла.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            Кортеж (успех, сообщение)
        """
        if not file_path:
            return False, "Путь к файлу не указан"
        
        try:
            df = pd.read_excel(file_path)
            
            missing_columns = [col for col in REQUIRED_COLUMNS if col not in df.columns]
            
            if missing_columns:
                return False, f"Отсутствуют обязательные колонки: {', '.join(missing_columns)}"
            
            return True, "Файл корректен"
            
        except FileNotFoundError:
            return False, f"Файл не найден: {file_path}"
        except Exception as e:
            return False, f"Ошибка чтения файла: {str(e)}"
