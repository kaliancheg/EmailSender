"""Главный класс приложения с поддержкой Outlook и SMTP"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import queue
import threading
import random
import os
import subprocess
import time
from datetime import datetime
from typing import Dict, Any, List

from core.constants import (
    MONTH_NAMES, DEFAULT_EMAIL_BODY,
    DEFAULT_THREAD_COUNT, DEFAULT_WINDOW_WIDTH, DEFAULT_WINDOW_HEIGHT,
    get_previous_month_subject
)
from core.logger_config import setup_logger
from models.email_data import EmailRecipient, EmailConfig
from models.smtp_models import SMTPConfig, QueuedEmail, EmailStatus, SendStatistics
from backend.email_service import EmailService
from backend.smtp_service import SMTPService
from backend.excel_service import ExcelService
from backend.settings_manager import SettingsManager
from backend.file_service import FileService
from frontend.ui_components import SettingsFrame
from frontend.smtp_settings import SMTPSettingsDialog


logger = setup_logger()


class EmailSenderApp:
    """Главный класс приложения с поддержкой Outlook и SMTP"""

    def __init__(self, root: tk.Tk):
        self.root = root
        self.root.title("Почтовая рассылка расчетных листов ФитоФарм")
        self.root.geometry(f"{DEFAULT_WINDOW_WIDTH}x{DEFAULT_WINDOW_HEIGHT}")
        self.root.resizable(False, False)

        # Сервисы
        self.settings_manager = SettingsManager()
        self.file_service = FileService()
        self.email_service: EmailService = None
        self.smtp_service: SMTPService = None

        # Режим отправки
        self.send_mode = tk.StringVar(value="smtp")  # "outlook" или "smtp" (SMTP по умолчанию)

        # Состояние
        self.is_paused = False
        self.is_cancelled = False
        self.total_emails = 0
        self.sent_count = 0
        self.failed_count = 0

        # Таймер рассылки
        self.send_start_time: Optional[datetime] = None
        self.send_elapsed_time: float = 0.0
        self.send_timer_job: Optional[str] = None
        self.last_stats: Optional[SendStatistics] = None  # Последняя статистика для обновления таймера

        # Список email-адресов с ошибками отправки
        self.failed_emails: List[str] = []
        self.failed_emails_lock = threading.Lock()
        self.failed_emails_file_path: str = None  # Путь к файлу с ошибками (создаётся один раз за сессию)

        # Очередь писем (для SMTP)
        self.email_queue: List[QueuedEmail] = []

        # Очередь UI
        self.ui_queue = queue.Queue()

        # Кэш аккаунтов Outlook (чтобы не запрашивать каждый раз)
        self.outlook_accounts_cached: List[str] = []
        self.outlook_accounts_loaded = False

        # Настройки по умолчанию
        self._init_default_settings()

        # Callbacks
        self._setup_callbacks()

        # UI
        self._setup_ui()
        self._load_settings()  # Загружаем настройки ДО инициализации UI
        self._update_mode_ui()  # Обновляем состояние кнопки SMTP
        self._init_stats_display()  # Инициализируем статистику текстовыми метками

        # Outlook НЕ загружаем при инициализации - только при переключении на Outlook режим

        # Обработка очереди UI
        self.root.after(100, self._process_ui_queue)

    def _init_default_settings(self):
        """Инициализирует настройки по умолчанию"""
        self.default_subject = get_previous_month_subject()
        self.default_body = DEFAULT_EMAIL_BODY
        self.folder_paths: Dict[int, str] = {}
        self.smtp_settings: Dict[str, Any] = {}

    def _setup_callbacks(self):
        """Настраивает callback функции"""
        self.callbacks = {
            'browse_excel': self._browse_excel,
            'browse_folder': self._browse_folder,
            'smtp_settings': self._open_smtp_settings
        }

    def _setup_ui(self):
        """Создаёт пользовательский интерфейс"""
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=tk.BOTH, expand=True)

        # Фрейм настроек
        self.settings_frame = SettingsFrame(main_frame, self.callbacks)
        self.settings_frame.pack(fill=tk.X, pady=(0, 10))

        # Выбор режима отправки
        self._setup_mode_frame(main_frame)

        # Фрейм управления
        self._setup_control_frame(main_frame)

        # Фрейм лога
        self._setup_log_frame(main_frame)

        # Центрирование
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - DEFAULT_WINDOW_WIDTH) // 2
        y = (self.root.winfo_screenheight() - DEFAULT_WINDOW_HEIGHT) // 2
        self.root.geometry(f"+{x}+{y}")

    def _setup_mode_frame(self, parent):
        """Создаёт фрейм выбора режима отправки"""
        mode_frame = ttk.LabelFrame(parent, text="Режим отправки", padding="10")
        mode_frame.pack(fill=tk.X, pady=(0, 10))

        # Переключатель
        ttk.Radiobutton(
            mode_frame, text="Outlook (COM)",
            variable=self.send_mode, value="outlook",
            command=self._on_mode_changed
        ).pack(side=tk.LEFT, padx=10)

        ttk.Radiobutton(
            mode_frame, text="SMTP (прямая отправка)",
            variable=self.send_mode, value="smtp",
            command=self._on_mode_changed
        ).pack(side=tk.LEFT, padx=10)

        # Кнопка настроек SMTP
        self.smtp_config_button = ttk.Button(
            mode_frame, text="⚙️ Настройки SMTP",
            command=self._open_smtp_settings
        )
        self.smtp_config_button.pack(side=tk.RIGHT, padx=10)
        
        # Статус SMTP
        self.smtp_status_label = ttk.Label(
            mode_frame, text="",
            foreground="gray"
        )
        self.smtp_status_label.pack(side=tk.RIGHT, padx=5)
        
        self._update_mode_ui()

    def _setup_control_frame(self, parent):
        """Создаёт фрейм управления"""
        control_frame = ttk.LabelFrame(parent, text="Управление рассылкой", padding="10")
        control_frame.pack(fill=tk.X, pady=(0, 10))

        # Статус
        status_frame = ttk.Frame(control_frame)
        status_frame.pack(fill=tk.X, pady=(0, 10))

        self.status_label = ttk.Label(
            status_frame, text="Готов к отправке",
            font=("TkDefaultFont", 10, "bold")
        )
        self.status_label.pack(side=tk.LEFT)

        self.counter_label = ttk.Label(
            status_frame, text="0/0",
            font=("TkDefaultFont", 10)
        )
        self.counter_label.pack(side=tk.RIGHT)

        # Прогресс-бар
        self.progress_var = tk.DoubleVar()
        self.progress_bar = ttk.Progressbar(
            control_frame, variable=self.progress_var,
            maximum=100, mode='determinate'
        )
        self.progress_bar.pack(fill=tk.X, pady=(0, 10))

        # Кнопки
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=tk.X)

        self.send_button = ttk.Button(
            button_frame, text="Начать рассылку",
            command=self._start_send
        )
        self.send_button.pack(side=tk.LEFT, padx=(0, 5))

        self.pause_button = ttk.Button(
            button_frame, text="Пауза",
            command=self._toggle_pause, state=tk.DISABLED
        )
        self.pause_button.pack(side=tk.LEFT, padx=(0, 5))

        self.preview_button = ttk.Button(
            button_frame, text="Предварительный просмотр",
            command=self._preview_email
        )
        self.preview_button.pack(side=tk.LEFT, padx=(0, 5))

        self.cancel_button = ttk.Button(
            button_frame, text="Отмена",
            command=self._cancel_send, state=tk.DISABLED
        )
        self.cancel_button.pack(side=tk.LEFT)

        # Кнопка просмотра ошибок
        self.errors_button = ttk.Button(
            button_frame, text="📋 Ошибки",
            command=self._open_failed_emails_file, state=tk.DISABLED
        )
        self.errors_button.pack(side=tk.LEFT, padx=(5, 0))

        # Статистика (для SMTP)
        self.stats_frame = ttk.Frame(control_frame)
        self.stats_frame.pack(fill=tk.X, pady=(5, 0))

        self.stats_label = ttk.Label(
            self.stats_frame,
            text="✅ 0 | ❌ 0 | 🕐 0 | 🔄 0",
            font=("TkDefaultFont", 9)
        )
        self.stats_label.pack(side=tk.LEFT)

        # Счётчик писем в файле с кнопкой обновления
        count_frame = ttk.Frame(self.stats_frame)
        count_frame.pack(side=tk.RIGHT)
        
        self.recipients_count_label = ttk.Label(
            count_frame,
            text="📋 Писем в файле: 0",
            font=("TkDefaultFont", 9),
            foreground="blue"
        )
        self.recipients_count_label.pack(side=tk.LEFT, padx=(0, 5))
        
        self.refresh_count_button = ttk.Button(
            count_frame,
            text="Ref.",
            command=self._refresh_recipients_count,
            width=5
        )
        self.refresh_count_button.pack(side=tk.LEFT)

    def _setup_log_frame(self, parent):
        """Создаёт фрейм лога"""
        log_frame = ttk.LabelFrame(parent, text="Лог выполнения", padding="10")
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = tk.Text(log_frame, height=12, width=80, wrap=tk.WORD)
        self.log_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, pady=(0, 5))

        # Контекстное меню
        if hasattr(self.settings_frame, 'add_context_menu'):
            self.settings_frame.add_context_menu(self.log_text)

        # Scrollbar
        scrollbar = ttk.Scrollbar(log_frame, orient=tk.VERTICAL, command=self.log_text.yview)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        self.log_text.configure(yscrollcommand=scrollbar.set)

    def _on_mode_changed(self):
        """Обработчик изменения режима отправки"""
        self._update_mode_ui()
        self._log_message(f"Режим отправки: {self.send_mode.get()}")
        
        # Загружаем аккаунты Outlook только при переключении на Outlook
        if self.send_mode.get() == "outlook":
            self._load_outlook_accounts()

    def _update_mode_ui(self):
        """Обновляет UI в зависимости от режима"""
        is_smtp = self.send_mode.get() == "smtp"
        
        # Активация кнопки настроек SMTP
        self.smtp_config_button.configure(state=tk.NORMAL if is_smtp else tk.DISABLED)
        
        # Статус SMTP
        if is_smtp and self.smtp_settings:
            self.smtp_status_label.config(
                text="✓ SMTP настроен",
                foreground="green"
            )
        elif is_smtp:
            self.smtp_status_label.config(
                text="⚠️ SMTP не настроен",
                foreground="orange"
            )
        else:
            self.smtp_status_label.config(text="")

    def _open_smtp_settings(self):
        """Открывает диалог настройки SMTP"""
        dialog = SMTPSettingsDialog(self.root, self.smtp_settings)
        result = dialog.show()
        
        if result:
            self.smtp_settings = {
                'smtp_server': result.smtp_server,
                'smtp_port': result.smtp_port,
                'email_login': result.email_login,
                'email_password': result.email_password,
                'use_ssl': result.use_ssl,
                'use_tls': result.use_tls,
                'sender_name': result.sender_name
            }
            self._update_mode_ui()
            self._log_message("Настройки SMTP сохранены")

    def _load_outlook_accounts(self):
        """Загружает аккаунты Outlook (с кэшированием)"""
        # Если уже загружали, используем кэш
        if self.outlook_accounts_loaded and self.outlook_accounts_cached:
            self.settings_frame.set_account_values(self.outlook_accounts_cached)
            return
        
        try:
            import win32com.client as win32
            outlook = win32.Dispatch('Outlook.Application')
            namespace = outlook.GetNamespace("MAPI")
            accounts = [account.SmtpAddress for account in namespace.Accounts]

            # Кэшируем аккаунты
            self.outlook_accounts_cached = accounts
            self.outlook_accounts_loaded = True
            
            self.settings_frame.set_account_values(accounts)
            self._log_message(f"Найдено {len(accounts)} аккаунтов Outlook")

        except Exception as e:
            self._log_message(f"Ошибка загрузки аккаунтов Outlook: {str(e)}", "ERROR")

    def _browse_excel(self):
        """Выбор Excel файла"""
        file = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.settings_frame.excel_path.set(file)
            # Обновляем счётчик писем
            self._update_recipients_count(file)

    def _update_recipients_count(self, file_path: str):
        """Обновляет счётчик получателей из Excel файла"""
        try:
            recipients = ExcelService.read_recipients(file_path)
            count = len(recipients)
            self.recipients_count_label.config(
                text=f"📋 Писем в файле: {count}",
                foreground="green" if count > 0 else "red"
            )
            self._log_message(f"Загружено {count} получателей из Excel")
            
            # Обновляем статистику: все письма в очереди
            self.ui_queue.put({
                'type': 'stats',
                'stats': SendStatistics(
                    total=count,
                    sent=0,
                    failed=0,
                    pending=count,
                    sending=0
                )
            })
        except Exception as e:
            self.recipients_count_label.config(
                text="📋 Писем в файле: ошибка",
                foreground="red"
            )
            self._log_message(f"Ошибка подсчёта получателей: {str(e)}", "ERROR")

    def _refresh_recipients_count(self):
        """Повторный подсчёт получателей из текущего файла"""
        file_path = self.settings_frame.excel_path.get()
        if not file_path:
            messagebox.showwarning("Предупреждение", "Сначала выберите Excel файл")
            return

        self._log_message("Обновление количества получателей...")
        # Обновляем только счётчик в файле, НЕ сбрасывая статистику рассылки
        try:
            recipients = ExcelService.read_recipients(file_path)
            count = len(recipients)
            self.recipients_count_label.config(
                text=f"📋 Писем в файле: {count}",
                foreground="green" if count > 0 else "red"
            )
            self._log_message(f"Загружено {count} получателей из Excel")
        except Exception as e:
            self.recipients_count_label.config(
                text="📋 Писем в файле: ошибка",
                foreground="red"
            )
            self._log_message(f"Ошибка подсчёта получателей: {str(e)}", "ERROR")

    def _browse_folder(self, folder_number: int):
        """Выбор папки"""
        folder = filedialog.askdirectory(title=f"Выберите папку {folder_number} с файлами")
        if folder:
            self.folder_paths[folder_number] = folder
            if folder_number == 1:
                self.settings_frame.folder_path_1.set(folder)
            elif folder_number == 2:
                self.settings_frame.folder_path_2.set(folder)
            elif folder_number == 3:
                self.settings_frame.folder_path_3.set(folder)

    def _log_message(self, message: str, level: str = "INFO"):
        """Добавляет сообщение в лог"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"
        self.log_text.insert(tk.END, formatted_message)
        self.log_text.see(tk.END)
        self.root.update_idletasks()

    def _process_ui_queue(self):
        """Обрабатывает очередь UI"""
        try:
            while True:
                item = self.ui_queue.get_nowait()
                item_type = item.get('type')

                if item_type == 'log':
                    self._log_message(item['message'], item['level'])
                elif item_type == 'status':
                    self.status_label.config(text=item['message'])
                elif item_type == 'progress':
                    self._update_progress(item['current'], item['total'])
                elif item_type == 'stats':
                    self._update_stats(item['stats'])
                elif item_type == 'complete':
                    self._on_send_complete(item)

                self.ui_queue.task_done()

        except queue.Empty:
            pass

        self.root.after(100, self._process_ui_queue)

    def _update_progress(self, current: int, total: int):
        """Обновляет прогресс"""
        self.counter_label.config(text=f"{current}/{total}")
        if total > 0:
            progress = (current / total) * 100
            self.progress_var.set(progress)

    def _update_stats(self, stats: SendStatistics):
        """Обновляет статистику SMTP"""
        # Сохраняем последнюю статистику для обновления таймера
        self.last_stats = stats
        
        # Форматируем elapsed time
        elapsed_str = self._format_elapsed_time(self.send_elapsed_time)

        # Используем фиксированную ширину полей для стабильного отображения
        # sending: 2 знака, sent: 3 знака, failed: 3 знака, время: 8 знаков (ЧЧ:ММ:СС)
        text = (
            f"📤 В процессе: {stats.sending:2d} | "
            f"✅ Отправлено: {stats.sent:3d} | "
            f"❌ Ошибка: {stats.failed:3d} | "
            f"⏱️ Время: {elapsed_str}"
        )
        self.stats_label.config(text=text)

    def _format_elapsed_time(self, seconds: float) -> str:
        """Форматирует прошедшее время в ЧЧ:ММ:СС"""
        hours = int(seconds // 3600)
        minutes = int((seconds % 3600) // 60)
        secs = int(seconds % 60)
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"

    def _start_send_timer(self):
        """Запускает таймер рассылки"""
        self.send_start_time = datetime.now()
        self.send_elapsed_time = 0.0
        self._update_send_timer()

    def _update_send_timer(self):
        """Обновляет таймер рассылки"""
        if self.send_start_time and not self.is_cancelled:
            self.send_elapsed_time = (datetime.now() - self.send_start_time).total_seconds()
            
            # Обновляем статистику с последними данными
            if self.last_stats:
                elapsed_str = self._format_elapsed_time(self.send_elapsed_time)
                stats = self.last_stats
                text = (
                    f"📤 В процессе: {stats.sending:2d} | "
                    f"✅ Отправлено: {stats.sent:3d} | "
                    f"❌ Ошибка: {stats.failed:3d} | "
                    f"⏱️ Время: {elapsed_str}"
                )
                self.stats_label.config(text=text)
            
            # Планируем следующее обновление через 1 секунду
            self.send_timer_job = self.root.after(1000, self._update_send_timer)

    def _stop_send_timer(self):
        """Останавливает таймер рассылки"""
        if self.send_timer_job:
            self.root.after_cancel(self.send_timer_job)
            self.send_timer_job = None
        if self.send_start_time:
            self.send_elapsed_time = (datetime.now() - self.send_start_time).total_seconds()

    def _init_stats_display(self):
        """Инициализирует отображение статистики текстовыми метками"""
        text = (
            f"📤 В процессе: 0 | "
            f"✅ Отправлено: 0 | "
            f"❌ Ошибка: 0 | "
            f"⏱️ Время: 00:00:00"
        )
        self.stats_label.config(text=text)

    def _update_stats_display(self):
        """Обновляет отображение статистики в начальное состояние"""
        self._init_stats_display()
        # Обновляем счётчик получателей без изменения статистики
        file_path = self.settings_frame.excel_path.get()
        if file_path:
            try:
                recipients = ExcelService.read_recipients(file_path)
                count = len(recipients)
                self.recipients_count_label.config(
                    text=f"📋 Писем в файле: {count}",
                    foreground="green" if count > 0 else "red"
                )
            except Exception:
                self.recipients_count_label.config(
                    text="📋 Писем в файле: 0",
                    foreground="red"
                )

    def _on_send_complete(self, item: dict):
        """Обработка завершения рассылки"""
        # Останавливаем таймер
        self._stop_send_timer()
        
        self.send_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.DISABLED)
        self.preview_button.config(state=tk.NORMAL)
        self.refresh_count_button.config(state=tk.NORMAL)

        # Активируем кнопку ошибок если есть неудачи
        self.errors_button.config(state=tk.NORMAL if self.failed_emails else tk.DISABLED)

        success = item.get('success_count', 0)
        failed = item.get('failed_count', 0)
        total = item.get('total', 0)
        cancelled = item.get('cancelled', False)

        if cancelled:
            status = f"Рассылка отменена. Отправлено: {success}/{total}"
            self._log_message(status, "WARNING")
        else:
            status = f"Рассылка завершена: {success} отправлено, {failed} ошибок"
            self._log_message(status, "INFO")
            messagebox.showinfo(
                "Завершено",
                f"Рассылка завершена!\n\n"
                f"✅ Отправлено: {success}\n"
                f"❌ Ошибок: {failed}\n"
                f"📊 Успех: {(success/total*100) if total > 0 else 0:.1f}%"
            )

    def _validate_settings(self) -> bool:
        """Проверяет настройки"""
        if self.send_mode.get() == "outlook":
            if not self.settings_frame.email_account.get():
                messagebox.showerror("Ошибка", "Выберите аккаунт Outlook")
                return False
        else:
            if not self.smtp_settings:
                messagebox.showerror("Ошибка", "Настройте SMTP сервер")
                return False

        if not self.settings_frame.excel_path.get():
            messagebox.showerror("Ошибка", "Укажите Excel файл")
            return False

        folders = [
            self.settings_frame.folder_path_1.get(),
            self.settings_frame.folder_path_2.get(),
            self.settings_frame.folder_path_3.get()
        ]
        if not any(folders):
            messagebox.showerror("Ошибка", "Укажите хотя бы одну папку с файлами")
            return False

        return True

    def _start_send(self):
        """Запуск рассылки"""
        try:
            # Валидация
            if not self._validate_settings():
                return

            # Чтение данных
            recipients = ExcelService.read_recipients(self.settings_frame.excel_path.get())

            if not recipients:
                messagebox.showerror("Ошибка", "Нет данных для рассылки")
                return

            # Обновляем счётчик перед отправкой
            self._log_message(f"Подготовлено к отправке: {len(recipients)} писем")

            # Сброс состояния
            self.is_cancelled = False
            self.is_paused = False
            # Сброс списка ошибок и пути к файлу
            with self.failed_emails_lock:
                self.failed_emails.clear()
            self.failed_emails_file_path = None

            # Сброс статистики в начальное состояние
            self.last_stats = None
            initial_stats = SendStatistics(
                total=0,
                sent=0,
                failed=0,
                pending=0,
                sending=0
            )
            self.ui_queue.put({
                'type': 'stats',
                'stats': initial_stats
            })
            self.send_elapsed_time = 0.0
            self._update_stats_display()

            # Блокировка кнопок
            self.send_button.config(state=tk.DISABLED)
            self.pause_button.config(state=tk.NORMAL)
            self.cancel_button.config(state=tk.NORMAL)
            self.preview_button.config(state=tk.DISABLED)
            self.refresh_count_button.config(state=tk.DISABLED)

            # Запуск в потоке
            self.total_emails = len(recipients)
            self._update_progress(0, self.total_emails)

            if self.send_mode.get() == "outlook":
                self._start_outlook_send(recipients)
            else:
                self._start_smtp_send(recipients)

        except Exception as e:
            self._restore_buttons()
            self._log_message(f"Критическая ошибка: {str(e)}", "ERROR")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

    def _start_outlook_send(self, recipients: List[EmailRecipient]):
        """Запускает отправку через Outlook"""
        config = EmailConfig(
            account=self.settings_frame.email_account.get(),
            subject=self.settings_frame.email_subject.get(),
            body=self.settings_frame.get_email_body(),
            folder_paths=[
                self.settings_frame.folder_path_1.get(),
                self.settings_frame.folder_path_2.get(),
                self.settings_frame.folder_path_3.get()
            ],
            thread_count=self.settings_frame.thread_count.get()
        )

        self.email_service = EmailService(config)

        self._log_message(f"Подготовка к отправке через Outlook: 0/{self.total_emails}")

        # Запускаем таймер
        self._start_send_timer()

        # Обновляем статистику с правильным total
        self.ui_queue.put({
            'type': 'stats',
            'stats': SendStatistics(
                total=len(recipients),
                sent=0,
                failed=0,
                pending=len(recipients),
                sending=0
            )
        })

        threading.Thread(
            target=self._send_outlook_thread,
            args=(recipients,),
            daemon=True
        ).start()

    def _start_smtp_send(self, recipients: List[EmailRecipient]):
        """Запускает отправку через SMTP"""
        config = SMTPConfig(
            smtp_server=self.smtp_settings['smtp_server'],
            smtp_port=self.smtp_settings['smtp_port'],
            email_login=self.smtp_settings['email_login'],
            email_password=self.smtp_settings['email_password'],
            use_ssl=self.smtp_settings['use_ssl'],
            use_tls=self.smtp_settings['use_tls'],
            sender_name=self.smtp_settings['sender_name']
        )

        # Оптимальные настройки для mail.ru (350 писем, 1 раз в месяц):
        # - 3 потока: баланс между скоростью и безопасностью
        # - 1.5 сек задержка: соблюдение лимитов mail.ru (~50-100 писем/час)
        # - 20 писем warm-up с задержкой 2.5 сек: "разогрев" аккаунта
        # - 50 писем в пакете + 10 сек пауза: предотвращение блокировок
        # - jitter ±0.3 сек: случайность для избежания паттернов бота
        self.smtp_service = SMTPService(
            config,
            thread_count=self.settings_frame.thread_count.get(),
            delay_between_emails=1.5,      # Задержка между письмами (сек)
            batch_size=50,                 # Писем в пакете перед паузой
            batch_delay=10.0,              # Пауза между пакетами (сек)
            jitter=0.3,                    # Разброс задержки ±0.3 сек
            warmup_count=20,               # Первые 20 писем с увеличенной задержкой
            warmup_delay=2.5               # Задержка для warm-up (сек)
        )

        # Подготовка очереди писем
        self.email_queue = []
        for recipient in recipients:
            # Поиск файлов
            attachments = []
            folders = [
                self.settings_frame.folder_path_1.get(),
                self.settings_frame.folder_path_2.get(),
                self.settings_frame.folder_path_3.get()
            ]
            files = [recipient.file_01, recipient.file_02, recipient.file_03]

            for folder, filename in zip(folders, files):
                if folder and filename:
                    file_path = self.file_service.find_file_in_folder(folder, filename)
                    if file_path:
                        attachments.append(file_path)

            queued_email = QueuedEmail(
                recipient_email=recipient.email,
                subject=self.settings_frame.email_subject.get(),
                body=self.settings_frame.get_email_body(),
                attachments=attachments
            )
            self.email_queue.append(queued_email)

        self._log_message(f"Подготовка к отправке через SMTP: 0/{self.total_emails}")

        # Запускаем таймер
        self._start_send_timer()

        # Обновляем статистику с правильным total
        self.ui_queue.put({
            'type': 'stats',
            'stats': SendStatistics(
                total=len(self.email_queue),
                sent=0,
                failed=0,
                pending=len(self.email_queue),
                sending=0
            )
        })

        threading.Thread(
            target=self._send_smtp_thread,
            args=(recipients,),
            daemon=True
        ).start()

    def _send_outlook_thread(self, recipients: list):
        """Поток отправки через Outlook"""
        try:
            failed_list = []

            def progress_callback(current: int, total: int, result):
                self.ui_queue.put({
                    'type': 'progress',
                    'current': current,
                    'total': total
                })

                # Собираем email с ошибками в локальный список
                if not result.success:
                    failed_list.append(result.email)

            success, failed = self.email_service.send_bulk(
                recipients,
                progress_callback=progress_callback
            )

            # После завершения send_bulk копируем все ошибки в общий список
            with self.failed_emails_lock:
                self.failed_emails.extend(failed_list)

            self.ui_queue.put({
                'type': 'complete',
                'success_count': success,
                'failed_count': failed,
                'total': len(recipients),
                'cancelled': self.email_service.is_cancelled if self.email_service else False
            })

        except Exception as e:
            self.ui_queue.put({
                'type': 'status',
                'message': f"Критическая ошибка в потоке: {str(e)}",
                'level': 'ERROR'
            })

    def _send_smtp_thread(self, recipients: list):
        """Поток отправки через SMTP"""
        try:
            failed_list = []

            def progress_callback(current: int, total: int, email: QueuedEmail, stats: SendStatistics):
                self.ui_queue.put({
                    'type': 'progress',
                    'current': current,
                    'total': total
                })
                # Обновление статистики в реальном времени
                self.ui_queue.put({
                    'type': 'stats',
                    'stats': stats
                })

            # Запускаем отправку
            stats = self.smtp_service.send_bulk(
                self.email_queue,
                progress_callback=progress_callback
            )

            # После завершения собираем все письма с ошибками
            for email in self.email_queue:
                if email.status == EmailStatus.FAILED:
                    failed_list.append(email.recipient_email)
                    self.ui_queue.put({
                        'type': 'log',
                        'message': f"❌ Ошибка: {email.recipient_email} - {email.error_message}",
                        'level': 'ERROR'
                    })

            # Копируем все ошибки в общий список
            with self.failed_emails_lock:
                self.failed_emails.extend(failed_list)

            # Финальное обновление статистики
            final_stats = SendStatistics(
                total=len(self.email_queue),
                sent=stats.sent,
                failed=stats.failed,
                pending=0,
                sending=0
            )
            self.ui_queue.put({
                'type': 'stats',
                'stats': final_stats
            })

            self.ui_queue.put({
                'type': 'complete',
                'success_count': stats.sent,
                'failed_count': stats.failed,
                'total': len(self.email_queue),
                'cancelled': self.smtp_service.is_cancelled if self.smtp_service else False
            })

        except Exception as e:
            self.ui_queue.put({
                'type': 'status',
                'message': f"Критическая ошибка в SMTP потоке: {str(e)}",
                'level': 'ERROR'
            })

    def _toggle_pause(self):
        """Пауза/продолжение"""
        if self.send_mode.get() == "outlook":
            if self.email_service:
                self.is_paused = self.email_service.toggle_pause()
                if self.is_paused:
                    self.pause_button.config(text="Продолжить")
                    self._log_message("Рассылка приостановлена", "WARNING")
                else:
                    self.pause_button.config(text="Пауза")
                    self._log_message("Рассылка продолжена", "INFO")
        else:
            if self.smtp_service:
                self.is_paused = self.smtp_service.toggle_pause()
                if self.is_paused:
                    self.pause_button.config(text="Продолжить")
                    self._log_message("SMTP рассылка приостановлена", "WARNING")
                else:
                    self.pause_button.config(text="Пауза")
                    self._log_message("SMTP рассылка продолжена", "INFO")

    def _cancel_send(self):
        """Отмена рассылки"""
        self.is_cancelled = True
        self.is_paused = False
        if self.email_service:
            self.email_service.cancel()
        if self.smtp_service:
            self.smtp_service.cancel()
        self._log_message("Рассылка отменена пользователем", "WARNING")
        
        # Сброс кнопки паузы в исходное состояние
        self.pause_button.config(text="Пауза")

    def _preview_email(self):
        """Предварительный просмотр случайного получателя"""
        try:
            if not self._validate_settings():
                return

            recipients = ExcelService.read_recipients(self.settings_frame.excel_path.get())

            if not recipients:
                messagebox.showerror("Ошибка", "Нет данных для просмотра")
                return

            # Выбираем случайного получателя
            recipient = random.choice(recipients)
            recipient_index = recipients.index(recipient) + 1

            if not recipient.email:
                messagebox.showerror("Ошибка", "Нет email для просмотра")
                return

            if self.send_mode.get() == "outlook":
                config = EmailConfig(
                    account=self.settings_frame.email_account.get(),
                    subject=self.settings_frame.email_subject.get(),
                    body=self.settings_frame.get_email_body(),
                    folder_paths=[
                        self.settings_frame.folder_path_1.get(),
                        self.settings_frame.folder_path_2.get(),
                        self.settings_frame.folder_path_3.get()
                    ]
                )
                email_service = EmailService(config)
                email_service.preview_email(recipient)
            else:
                # Для SMTP просто показываем информацию
                attachments = []
                folders = [
                    self.settings_frame.folder_path_1.get(),
                    self.settings_frame.folder_path_2.get(),
                    self.settings_frame.folder_path_3.get()
                ]
                files = [recipient.file_01, recipient.file_02, recipient.file_03]

                for folder, filename in zip(folders, files):
                    if folder and filename:
                        file_path = self.file_service.find_file_in_folder(folder, filename)
                        if file_path:
                            attachments.append(file_path)

                info = (
                    f"Получатель: {recipient.email}\n\n"
                    f"Тема: {self.settings_frame.email_subject.get()}\n\n"
                    f"Текст:\n{self.settings_frame.get_email_body()}\n\n"
                    f"Вложения ({len(attachments)}):\n" +
                    "\n".join(f"  • {a}" for a in attachments)
                )
                messagebox.showinfo("Предварительный просмотр (SMTP)", info)

            self._log_message(f"Предварительный просмотр для: {recipient.email} (строка {recipient_index})")

        except Exception as e:
            self._log_message(f"Ошибка预览: {str(e)}", "ERROR")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

    def _restore_buttons(self):
        """Восстанавливает кнопки"""
        self.send_button.config(state=tk.NORMAL)
        self.pause_button.config(state=tk.DISABLED)
        self.cancel_button.config(state=tk.DISABLED)
        self.preview_button.config(state=tk.NORMAL)
        self.refresh_count_button.config(state=tk.NORMAL)
        # Кнопка ошибок активна только если есть ошибки
        self.errors_button.config(state=tk.NORMAL if self.failed_emails else tk.DISABLED)
        
        # Сброс кнопки паузы в исходное состояние
        self.pause_button.config(text="Пауза")
        self.is_paused = False

    def _save_failed_emails_to_file(self):
        """Сохраняет список email-адресов с ошибками в текстовый файл (один раз за сессию)"""
        if not self.failed_emails:
            return None
        
        # Если файл уже создан, возвращаем его путь
        if self.failed_emails_file_path and os.path.exists(self.failed_emails_file_path):
            return self.failed_emails_file_path

        # Формируем имя файла с датой и временем
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        filename = f"failed_emails_{timestamp}.txt"
        filepath = os.path.join(os.getcwd(), filename)

        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                f.write("Список e-mail адресов с ошибками отправки\n")
                f.write(f"Дата: {datetime.now().strftime('%d.%m.%Y %H:%M:%S')}\n")
                f.write(f"Всего ошибок: {len(self.failed_emails)}\n")
                f.write("=" * 50 + "\n\n")
                for email in self.failed_emails:
                    f.write(f"{email}\n")
            
            # Сохраняем путь к файлу
            self.failed_emails_file_path = filepath
            
            self._log_message(f"Список ошибок сохранён: {filename}")
            return filepath
        
        except Exception as e:
            self._log_message(f"Ошибка сохранения файла ошибок: {str(e)}", "ERROR")
            return None

    def _open_failed_emails_file(self):
        """Открывает файл со списком email-адресов с ошибками"""
        if not self.failed_emails:
            messagebox.showinfo("Информация", "Нет email-адресов с ошибками отправки")
            return
        
        # Сохраняем файл (или получаем путь к существующему)
        filepath = self._save_failed_emails_to_file()
        
        if filepath and os.path.exists(filepath):
            # Открываем файл в текстовом редакторе по умолчанию
            try:
                filename = os.path.basename(filepath)
                os.startfile(filepath)
                self._log_message(f"Открыт файл с ошибками: {filename}")
            except Exception as e:
                self._log_message(f"Ошибка открытия файла: {str(e)}", "ERROR")
                messagebox.showerror("Ошибка", f"Не удалось открыть файл:\n{filepath}")

    def _load_settings(self):
        """Загружает сохранённые настройки"""
        settings = self.settings_manager.load()

        # Всегда устанавливаем актуальную тему с прошлым месяцем
        self.settings_frame.email_subject.set(self.default_subject)

        if not settings:
            # Устанавливаем текст письма по умолчанию
            self.settings_frame.set_email_body(self.default_body)
            return

        if 'excel_path' in settings:
            self.settings_frame.excel_path.set(settings['excel_path'])
        if 'email_account' in settings:
            self.settings_frame.email_account.set(settings['email_account'])
        if 'thread_count' in settings:
            self.settings_frame.thread_count.set(settings['thread_count'])
        # email_subject НЕ загружаем - используем актуальный
        if 'email_body' in settings:
            self.settings_frame.set_email_body(settings['email_body'])
        if 'folder_path_1' in settings:
            self.settings_frame.folder_path_1.set(settings['folder_path_1'])
            self.folder_paths[1] = settings['folder_path_1']
        if 'folder_path_2' in settings:
            self.settings_frame.folder_path_2.set(settings['folder_path_2'])
            self.folder_paths[2] = settings['folder_path_2']
        if 'folder_path_3' in settings:
            self.settings_frame.folder_path_3.set(settings['folder_path_3'])
            self.folder_paths[3] = settings['folder_path_3']
        
        # Загрузка SMTP настроек
        if 'smtp_settings' in settings:
            self.smtp_settings = settings['smtp_settings']
        
        # Загрузка режима отправки
        if 'send_mode' in settings:
            self.send_mode.set(settings['send_mode'])

        self._log_message("Настройки загружены")
        
        # Обновляем счётчик получателей если файл указан
        if 'excel_path' in settings and settings['excel_path']:
            self._update_recipients_count(settings['excel_path'])

    def _save_settings(self):
        """Сохраняет настройки"""
        settings = {
            'excel_path': self.settings_frame.excel_path.get(),
            'email_account': self.settings_frame.email_account.get(),
            'thread_count': self.settings_frame.thread_count.get(),
            'email_subject': self.settings_frame.email_subject.get(),
            'email_body': self.settings_frame.get_email_body(),
            'folder_path_1': self.settings_frame.folder_path_1.get(),
            'folder_path_2': self.settings_frame.folder_path_2.get(),
            'folder_path_3': self.settings_frame.folder_path_3.get(),
            'smtp_settings': self.smtp_settings,
            'send_mode': self.send_mode.get()
        }

        self.settings_manager.save(settings)

    def on_closing(self):
        """Обработчик закрытия окна"""
        self._save_settings()

        # Принудительная отмена рассылки если она активна
        if self.email_service:
            self.email_service.cancel()
            self.email_service.is_paused = False  # Сброс паузы для выхода из цикла
        if self.smtp_service:
            self.smtp_service.cancel()
            self.smtp_service.is_paused = False  # Сброс паузы для выхода из цикла

        # Небольшая задержка для завершения потоков
        time.sleep(0.3)

        self.root.destroy()
