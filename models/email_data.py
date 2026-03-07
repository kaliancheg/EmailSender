"""Модели данных для email рассылки"""

from dataclasses import dataclass, field
from typing import List, Optional
from datetime import datetime


@dataclass
class EmailRecipient:
    """Модель получателя письма"""
    email: str
    file_01: Optional[str] = None
    file_02: Optional[str] = None
    file_03: Optional[str] = None
    
    @property
    def has_files(self) -> bool:
        """Есть ли файлы для прикрепления"""
        return any([self.file_01, self.file_02, self.file_03])


@dataclass
class EmailConfig:
    """Конфигурация email рассылки"""
    account: str
    subject: str
    body: str
    folder_paths: List[str]
    thread_count: int = 3
    
    def get_folder_path(self, index: int) -> Optional[str]:
        """Получить путь к папке по индексу (0-2)"""
        if 0 <= index < len(self.folder_paths):
            return self.folder_paths[index]
        return None


@dataclass
class SendResult:
    """Результат отправки письма"""
    success: bool
    email: str
    error: Optional[str] = None
    attached_files: List[str] = field(default_factory=list)
    timestamp: datetime = field(default_factory=datetime.now)
    
    @property
    def status_text(self) -> str:
        """Текстовое представление статуса"""
        if self.success:
            return "Отправлено"
        return f"Ошибка: {self.error}"
