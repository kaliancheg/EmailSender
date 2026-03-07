"""Модели данных для SMTP отправки"""

from dataclasses import dataclass, field
from typing import Optional
from datetime import datetime
from enum import Enum


class EmailStatus(Enum):
    """Статус отправки письма"""
    PENDING = "pending"      # В очереди
    SENDING = "sending"      # Отправляется
    SENT = "sent"           # Отправлено
    FAILED = "failed"       # Ошибка
    RETRY = "retry"         # Повторная попытка


@dataclass
class SMTPConfig:
    """Конфигурация SMTP сервера"""
    smtp_server: str
    smtp_port: int
    email_login: str
    email_password: str
    use_ssl: bool = True
    use_tls: bool = True  # Для порта 587
    sender_name: str = ""  # Имя отправителя
    
    @property
    def display_name(self) -> str:
        """Имя для отображения"""
        if self.sender_name:
            return f"{self.sender_name} <{self.email_login}>"
        return self.email_login


@dataclass
class QueuedEmail:
    """Письмо в очереди на отправку"""
    recipient_email: str
    subject: str
    body: str
    attachments: list[str] = field(default_factory=list)
    
    # Статус
    status: EmailStatus = EmailStatus.PENDING
    retry_count: int = 0
    max_retries: int = 3
    error_message: Optional[str] = None
    
    # Временные метки
    created_at: datetime = field(default_factory=datetime.now)
    sent_at: Optional[datetime] = None
    last_attempt: Optional[datetime] = None
    
    @property
    def can_retry(self) -> bool:
        """Можно ли повторить попытку"""
        return self.retry_count < self.max_retries
    
    @property
    def status_display(self) -> str:
        """Человекочитаемый статус"""
        status_map = {
            EmailStatus.PENDING: "В очереди",
            EmailStatus.SENDING: "Отправляется",
            EmailStatus.SENT: "Отправлено",
            EmailStatus.FAILED: "Ошибка",
            EmailStatus.RETRY: "Повтор"
        }
        return status_map.get(self.status, "Неизвестно")


@dataclass
class SendStatistics:
    """Статистика рассылки"""
    total: int = 0
    sent: int = 0
    failed: int = 0
    pending: int = 0
    retry: int = 0
    
    @property
    def progress_percent(self) -> float:
        """Процент выполнения"""
        if self.total == 0:
            return 0.0
        return ((self.sent + self.failed) / self.total) * 100
    
    @property
    def success_rate(self) -> float:
        """Процент успешных"""
        completed = self.sent + self.failed
        if completed == 0:
            return 0.0
        return (self.sent / completed) * 100
