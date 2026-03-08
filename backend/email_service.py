"""Сервис для отправки email через Outlook"""

import os
import pythoncom
import win32com.client as win32
from typing import List, Optional, Callable
from concurrent.futures import ThreadPoolExecutor, as_completed
import logging

from models.email_data import EmailRecipient, EmailConfig, SendResult
from backend.file_service import FileService

logger = logging.getLogger("email_sender")


class EmailService:
    """Сервис для отправки email через Outlook"""

    def __init__(self, config: EmailConfig):
        """
        Инициализация сервиса.

        Args:
            config: Конфигурация email рассылки
        """
        self.config = config
        self.file_service = FileService()
        self.is_cancelled = False
        self.is_paused = False
        # Outlook создаётся в каждом потоке отдельно
    
    def send_email(self, recipient: EmailRecipient) -> SendResult:
        """
        Отправляет письмо одному получателю.

        Args:
            recipient: Получатель письма

        Returns:
            Результат отправки
        """
        try:
            # Инициализируем COM для потока
            pythoncom.CoInitialize()

            try:
                # Создаём Outlook в этом потоке
                outlook = win32.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To = recipient.email
                mail.Subject = self.config.subject
                mail.Body = self.config.body

                # Прикрепляем файлы
                attached_files = self._attach_files(mail, recipient)

                if not attached_files:
                    return SendResult(
                        success=False,
                        email=recipient.email,
                        error="Нет файлов для прикрепления"
                    )

                logger.info(f"Прикреплены файлы для {recipient.email}: {', '.join(attached_files)}")

                # Отправляем письмо
                mail.Send()

                logger.info(f"Письмо отправлено: {recipient.email}")

                return SendResult(
                    success=True,
                    email=recipient.email,
                    attached_files=attached_files
                )

            except Exception as e:
                error_msg = f"Ошибка отправки письма для {recipient.email}: {str(e)}"
                logger.error(error_msg)
                return SendResult(
                    success=False,
                    email=recipient.email,
                    error=str(e)
                )

            finally:
                # Освобождаем COM
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            error_msg = f"Критическая ошибка в потоке для {recipient.email}: {str(e)}"
            logger.error(error_msg)
            return SendResult(
                success=False,
                email=recipient.email,
                error=str(e)
            )
    
    def _attach_files(self, mail, recipient: EmailRecipient) -> List[str]:
        """
        Прикрепляет файлы к письму.
        
        Args:
            mail: Outlook Mail объект
            recipient: Получатель с файлами
            
        Returns:
            Список прикреплённых файлов
        """
        attached_files = []
        
        files = [
            (recipient.file_01, 0),
            (recipient.file_02, 1),
            (recipient.file_03, 2)
        ]
        
        for filename, folder_index in files:
            if filename:
                folder_path = self.config.get_folder_path(folder_index)
                if folder_path:
                    file_path = self.file_service.find_file_in_folder(folder_path, filename)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))
        
        return attached_files
    
    def send_bulk(self, recipients: List[EmailRecipient], 
                  progress_callback: Optional[Callable[[int, int, SendResult], None]] = None) -> tuple[int, int]:
        """
        Массовая отправка писем.
        
        Args:
            recipients: Список получателей
            progress_callback: Callback для обновления прогресса (current, total, result)
            
        Returns:
            Кортеж (успешно, ошибок)
        """
        self.is_cancelled = False
        self.is_paused = False
        
        max_workers = min(self.config.thread_count, len(recipients))
        success_count = 0
        failed_count = 0
        
        logger.info(f"Запуск {max_workers} потоков для отправки {len(recipients)} писем")
        
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            futures = {
                executor.submit(self.send_email, recipient): i 
                for i, recipient in enumerate(recipients)
            }
            
            for i, future in enumerate(as_completed(futures)):
                if self.is_cancelled:
                    for f in futures:
                        f.cancel()
                    break
                
                result = future.result()
                
                if result.success:
                    success_count += 1
                else:
                    failed_count += 1
                
                if progress_callback:
                    progress_callback(i + 1, len(recipients), result)
        
        return success_count, failed_count
    
    def cancel(self):
        """Отменяет рассылку"""
        self.is_cancelled = True
        logger.warning("Рассылка отменена пользователем")
    
    def toggle_pause(self) -> bool:
        """
        Переключает состояние паузы.
        
        Returns:
            Текущее состояние паузы
        """
        self.is_paused = not self.is_paused
        status = "приостановлена" if self.is_paused else "продолжена"
        logger.info(f"Рассылка {status}")
        return self.is_paused
    
    def preview_email(self, recipient: EmailRecipient) -> bool:
        """
        Показывает предварительный просмотр письма.

        Args:
            recipient: Получатель для превью

        Returns:
            True если успешно
        """
        try:
            # Инициализируем COM
            pythoncom.CoInitialize()

            try:
                outlook = win32.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To = recipient.email
                mail.Subject = self.config.subject
                mail.Body = self.config.body

                # Прикрепляем файлы
                attached_files = self._attach_files(mail, recipient)

                # Показываем письмо
                mail.Display()

                logger.info(f"Предварительный просмотр письма для: {recipient.email}")
                if attached_files:
                    logger.info(f"Прикрепленные файлы: {', '.join(attached_files)}")

                return True

            except Exception as e:
                logger.error(f"Ошибка предварительного просмотра: {str(e)}")
                return False

            finally:
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            logger.error(f"Критическая ошибка в предварительном просмотре: {str(e)}")
            return False
