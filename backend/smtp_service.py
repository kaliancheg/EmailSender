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
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging
import threading

from models.smtp_models import SMTPConfig, QueuedEmail, EmailStatus, SendStatistics

logger = logging.getLogger("email_sender")


class SMTPService:
    """Сервис для отправки email через SMTP"""

    def __init__(self, config: SMTPConfig, thread_count: int = 3, delay_between_emails: float = 0.0):
        """
        Инициализация SMTP сервиса.

        Args:
            config: Конфигурация SMTP сервера
            thread_count: Количество потоков для отправки (по умолчанию 3)
            delay_between_emails: Задержка между письмами в секундах (по умолчанию 0)
        """
        self.config = config
        self.thread_count = min(max(thread_count, 1), 10)  # Ограничение 1-10
        self.delay_between_emails = delay_between_emails  # Задержка в секундах
        self.is_cancelled = False
        self.is_paused = False
        self._pause_lock = threading.Lock()
        self._status_lock = threading.Lock()  # Блокировка для защиты статусов писем
    
    def send_email(self, queued_email: QueuedEmail) -> bool:
        """
        Отправляет одно письмо (потокобезопасная версия).

        Args:
            queued_email: Письмо для отправки

        Returns:
            True если успешно
        """
        return self._send_single_email(queued_email)
    
    def _send_single_email(self, queued_email: QueuedEmail) -> bool:
        """
        Отправляет одно письмо (внутренний метод для потоков).

        Args:
            queued_email: Письмо для отправки

        Returns:
            True если успешно
        """
        try:
            # Устанавливаем статус SENDING с блокировкой
            with self._status_lock:
                queued_email.status = EmailStatus.SENDING
                queued_email.last_attempt = datetime.now()

            # Задержка перед отправкой (для соблюдения лимитов SMTP)
            if self.delay_between_emails > 0:
                import time
                time.sleep(self.delay_between_emails)

            # Создаём сообщение
            msg = MIMEMultipart()

            # Важно: From должен совпадать с email_login для авторизации
            # Для совместимости с Yandex, Gmail и другими
            if self.config.sender_name:
                # Формат: "Имя <email>" - правильно для всех SMTP серверов
                msg['From'] = f"{self.config.sender_name} <{self.config.email_login}>"
            else:
                # Просто email
                msg['From'] = self.config.email_login

            msg['To'] = queued_email.recipient_email
            msg['Subject'] = queued_email.subject

            # Добавляем тело письма
            msg.attach(MIMEText(queued_email.body, 'plain', 'utf-8'))

            # Добавляем вложения (создаём копию списка для безопасности потока)
            attachments_copy = list(queued_email.attachments)
            attached_count = 0
            for file_path in attachments_copy:
                if Path(file_path).exists():
                    attachment = self._create_attachment(file_path)
                    if attachment:
                        msg.attach(attachment)
                        attached_count += 1
                        logger.debug(f"Вложение добавлено: {file_path} ({attached_count} из {len(attachments_copy)})")
                else:
                    logger.warning(f"Файл не найден при отправке: {file_path}")

            logger.info(f"Подготовлено письмо для {queued_email.recipient_email}: {attached_count} вложений")

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

                # Успешная отправка - устанавливаем статус с блокировкой
                with self._status_lock:
                    queued_email.status = EmailStatus.SENT
                    queued_email.sent_at = datetime.now()

                logger.info(f"Письмо отправлено: {queued_email.recipient_email} ({attached_count} вложений)")
                return True

            finally:
                server.quit()

        except smtplib.SMTPAuthenticationError as e:
            error_msg = f"Ошибка авторизации SMTP: {str(e)}"
            logger.error(error_msg)
            with self._status_lock:
                queued_email.status = EmailStatus.FAILED
                queued_email.error_message = error_msg
            return False

        except smtplib.SMTPConnectError as e:
            error_msg = f"Ошибка подключения к SMTP серверу: {str(e)}"
            logger.error(error_msg)
            with self._status_lock:
                queued_email.status = EmailStatus.FAILED
                queued_email.error_message = error_msg
            return False

        except smtplib.SMTPSenderRefused as e:
            # Ошибка "Sender address rejected" - отклонён адрес отправителя
            error_msg = f"Адрес отправителя отклонён: {str(e)}. Убедитесь, что From совпадает с логином авторизации."
            logger.error(error_msg)
            with self._status_lock:
                queued_email.status = EmailStatus.FAILED
                queued_email.error_message = error_msg
            return False

        except Exception as e:
            error_msg = f"Ошибка отправки письма: {str(e)}"
            logger.error(error_msg)
            with self._status_lock:
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

            # Получаем размер файла для логирования
            file_size = path.stat().st_size
            logger.debug(f"Чтение файла: {file_path} (размер: {file_size} байт)")

            # Читаем файл полностью перед кодированием
            with open(path, "rb") as attachment_file:
                payload = attachment_file.read()

            # Проверяем, что файл прочитан полностью
            if len(payload) != file_size:
                logger.error(f"Файл прочитан не полностью: {file_path} (ожидалось {file_size}, прочитано {len(payload)})")
                return None

            part = MIMEBase('application', 'octet-stream')
            part.set_payload(payload)

            encoders.encode_base64(part)
            
            # Используем ASCII-safe имя файла
            filename = path.name
            part.add_header(
                'Content-Disposition',
                f'attachment; filename="{filename}"'
            )
            
            logger.debug(f"Вложение создано: {filename} ({file_size} байт)")
            return part

        except Exception as e:
            logger.error(f"Ошибка добавления вложения {file_path}: {str(e)}")
            return None
    
    def send_bulk(
        self,
        emails: List[QueuedEmail],
        progress_callback: Optional[Callable[[int, int, QueuedEmail, SendStatistics], None]] = None
    ) -> SendStatistics:
        """
        Массовая многопоточная отправка писем.

        Args:
            emails: Список писем для отправки
            progress_callback: Callback (current, total, email, stats)

        Returns:
            Статистика отправки
        """
        self.is_cancelled = False
        self.is_paused = False

        stats = SendStatistics(total=len(emails))

        logger.info(f"Начало SMTP рассылки: {len(emails)} писем, потоков: {self.thread_count}")

        # Ограничиваем количество потоков количеством писем
        max_workers = min(self.thread_count, len(emails))

        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Создаём задачи для каждого письма
            future_to_email = {
                executor.submit(self._send_single_email, email): email
                for email in emails
            }

            # Обрабатываем результаты по мере завершения
            for i, future in enumerate(as_completed(future_to_email)):
                if self.is_cancelled:
                    logger.warning("Рассылка отменена пользователем")
                    # Отменяем оставшиеся задачи
                    for f in future_to_email:
                        f.cancel()
                    break

                # Пауза
                while self.is_paused and not self.is_cancelled:
                    import time
                    time.sleep(0.5)

                email = future_to_email[future]

                try:
                    success = future.result()

                    # Статус уже установлен в _send_single_email, просто считаем
                    with self._status_lock:
                        if success:
                            stats.sent += 1
                        else:
                            stats.failed += 1
                            logger.warning(f"Ошибка отправки для {email.recipient_email}: {email.error_message}")

                        # Подсчёт статистики с блокировкой
                        stats.pending = sum(1 for e in emails if e.status == EmailStatus.PENDING)
                        stats.sending = len(emails) - stats.sent - stats.failed - stats.pending

                    if progress_callback:
                        progress_callback(i + 1, len(emails), email, stats)

                except Exception as e:
                    logger.error(f"Ошибка в потоке для {email.recipient_email}: {str(e)}")
                    with self._status_lock:
                        stats.failed += 1
                        # Подсчёт статистики
                        stats.pending = sum(1 for e in emails if e.status == EmailStatus.PENDING)
                        stats.sending = len(emails) - stats.sent - stats.failed - stats.pending
                    if progress_callback:
                        progress_callback(i + 1, len(emails), email, stats)

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
