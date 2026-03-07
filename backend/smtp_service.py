"""SMTP сервис для отправки email"""

import smtplib
import ssl
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
from email.mime.base import MIMEBase
from email import encoders
from pathlib import Path
from typing import List, Optional, Callable
from datetime import datetime
import logging

from models.smtp_models import SMTPConfig, QueuedEmail, EmailStatus, SendStatistics

logger = logging.getLogger("email_sender")


class SMTPService:
    """Сервис для отправки email через SMTP"""
    
    def __init__(self, config: SMTPConfig):
        """
        Инициализация SMTP сервиса.
        
        Args:
            config: Конфигурация SMTP сервера
        """
        self.config = config
        self.is_cancelled = False
        self.is_paused = False
    
    def send_email(self, queued_email: QueuedEmail) -> bool:
        """
        Отправляет одно письмо.

        Args:
            queued_email: Письмо для отправки

        Returns:
            True если успешно
        """
        try:
            queued_email.status = EmailStatus.SENDING
            queued_email.last_attempt = datetime.now()

            # Создаём сообщение
            msg = MIMEMultipart()
            
            # Важно: From должен совпадать с email_login для авторизации
            # Для совместимости с Yandex, Gmail и другими
            msg['From'] = self.config.email_login
            if self.config.sender_name:
                # Если есть имя, добавляем его в формате: "Имя <email>"
                msg['From'] = f"{self.config.sender_name} <{self.config.email_login}>"
            
            msg['To'] = queued_email.recipient_email
            msg['Subject'] = queued_email.subject

            # Добавляем тело письма
            msg.attach(MIMEText(queued_email.body, 'plain', 'utf-8'))

            # Добавляем вложения
            for file_path in queued_email.attachments:
                if Path(file_path).exists():
                    attachment = self._create_attachment(file_path)
                    if attachment:
                        msg.attach(attachment)

            # Подключение к серверу
            if self.config.use_ssl:
                # SSL подключение (порт 465)
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(self.config.smtp_server, self.config.smtp_port, context=context)
            else:
                # Обычное подключение с TLS (порт 587)
                server = smtplib.SMTP(self.config.smtp_server, self.config.smtp_port)
                if self.config.use_tls:
                    server.starttls(context=ssl.create_default_context())

            try:
                # Авторизация
                server.login(self.config.email_login, self.config.email_password)

                # Отправка
                server.send_message(msg)

                queued_email.status = EmailStatus.SENT
                queued_email.sent_at = datetime.now()

                logger.info(f"Письмо отправлено: {queued_email.recipient_email}")
                return True

            finally:
                server.quit()

        except smtplib.SMTPAuthenticationError as e:
            error_msg = f"Ошибка авторизации SMTP: {str(e)}"
            logger.error(error_msg)
            queued_email.status = EmailStatus.FAILED
            queued_email.error_message = error_msg
            return False

        except smtplib.SMTPConnectError as e:
            error_msg = f"Ошибка подключения к SMTP серверу: {str(e)}"
            logger.error(error_msg)
            queued_email.status = EmailStatus.FAILED
            queued_email.error_message = error_msg
            return False

        except smtplib.SMTPSenderRefused as e:
            # Ошибка "Sender address rejected" - отклонён адрес отправителя
            error_msg = f"Адрес отправителя отклонён: {str(e)}. Убедитесь, что From совпадает с логином авторизации."
            logger.error(error_msg)
            queued_email.status = EmailStatus.FAILED
            queued_email.error_message = error_msg
            return False

        except Exception as e:
            error_msg = f"Ошибка отправки письма: {str(e)}"
            logger.error(error_msg)
            queued_email.status = EmailStatus.FAILED
            queued_email.error_message = error_msg
            return False
    
    def _create_attachment(self, file_path: str) -> Optional[MIMEBase]:
        """
        Создаёт вложение для письма.
        
        Args:
            file_path: Путь к файлу
            
        Returns:
            MIMEBase объект или None
        """
        try:
            path = Path(file_path)
            if not path.exists():
                logger.warning(f"Файл не найден: {file_path}")
                return None
            
            with open(path, "rb") as attachment:
                part = MIMEBase('application', 'octet-stream')
                part.set_payload(attachment.read())
            
            encoders.encode_base64(part)
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{path.name}"'
            )
            return part
            
        except Exception as e:
            logger.error(f"Ошибка добавления вложения {file_path}: {str(e)}")
            return None
    
    def send_bulk(
        self,
        emails: List[QueuedEmail],
        progress_callback: Optional[Callable[[int, int, QueuedEmail], None]] = None
    ) -> SendStatistics:
        """
        Массовая отправка писем.
        
        Args:
            emails: Список писем для отправки
            progress_callback: Callback (current, total, email)
            
        Returns:
            Статистика отправки
        """
        self.is_cancelled = False
        self.is_paused = False
        
        stats = SendStatistics(total=len(emails))
        
        logger.info(f"Начало SMTP рассылки: {len(emails)} писем")
        
        for i, email in enumerate(emails):
            if self.is_cancelled:
                logger.warning("Рассылка отменена пользователем")
                break
            
            # Пауза
            while self.is_paused and not self.is_cancelled:
                import time
                time.sleep(1)
            
            # Отправка
            success = self.send_email(email)
            
            if success:
                stats.sent += 1
            else:
                stats.failed += 1
                
                # Попытка повтора
                if email.can_retry:
                    email.retry_count += 1
                    email.status = EmailStatus.RETRY
                    stats.retry += 1
                    logger.warning(f"Попытка повтора {email.retry_count}/{email.max_retries} для {email.recipient_email}")
            
            # Обновление статистики
            stats.pending = len(emails) - stats.sent - stats.failed
            
            if progress_callback:
                progress_callback(i + 1, len(emails), email)
        
        return stats
    
    def cancel(self):
        """Отменяет рассылку"""
        self.is_cancelled = True
        logger.warning("SMTP рассылка отменена")
    
    def toggle_pause(self) -> bool:
        """Переключает паузу"""
        self.is_paused = not self.is_paused
        status = "приостановлена" if self.is_paused else "продолжена"
        logger.info(f"SMTP рассылка {status}")
        return self.is_paused
    
    def test_connection(self) -> tuple[bool, str]:
        """
        Проверяет подключение к SMTP серверу.
        
        Returns:
            (успех, сообщение)
        """
        try:
            if self.config.use_ssl:
                context = ssl.create_default_context()
                server = smtplib.SMTP_SSL(self.config.smtp_server, self.config.smtp_port, context=context)
            else:
                server = smtplib.SMTP(self.config.smtp_server, self.config.smtp_port)
                if self.config.use_tls:
                    server.starttls(context=ssl.create_default_context())
            
            try:
                server.login(self.config.email_login, self.config.email_password)
                server.quit()
                return True, "Подключение успешно"
            finally:
                try:
                    server.quit()
                except:
                    pass
                    
        except smtplib.SMTPAuthenticationError:
            return False, "Ошибка авторизации (неверный логин/пароль)"
        except smtplib.SMTPConnectError:
            return False, "Не удалось подключиться к серверу"
        except Exception as e:
            return False, f"Ошибка: {str(e)}"
