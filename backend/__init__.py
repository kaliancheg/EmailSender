# Backend модуль - бизнес-логика приложения
from .email_service import EmailService
from .file_service import FileService
from .settings_manager import SettingsManager

__all__ = ['EmailService', 'FileService', 'SettingsManager']
