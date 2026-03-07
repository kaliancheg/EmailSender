# Models модуль - модели данных
from .email_data import EmailRecipient, EmailConfig, SendResult
from .smtp_models import SMTPConfig, QueuedEmail, EmailStatus, SendStatistics

__all__ = [
    'EmailRecipient', 'EmailConfig', 'SendResult',
    'SMTPConfig', 'QueuedEmail', 'EmailStatus', 'SendStatistics'
]
