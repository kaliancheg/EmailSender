"""Диалог настройки SMTP"""

import tkinter as tk
from tkinter import ttk, messagebox
from typing import Optional, Dict, Any

from models.smtp_models import SMTPConfig


class SMTPSettingsDialog:
    """Диалоговое окно настройки SMTP"""
    
    def __init__(self, parent: tk.Tk, settings: Optional[Dict[str, Any]] = None):
        """
        Инициализация диалога.
        
        Args:
            parent: Родительское окно
            settings: Существующие настройки (опционально)
        """
        self.parent = parent
        self.result: Optional[SMTPConfig] = None
        self.settings = settings or {}
        
        self._create_dialog()
    
    def _create_dialog(self):
        """Создаёт диалоговое окно"""
        self.dialog = tk.Toplevel(self.parent)
        self.dialog.title("Настройка SMTP сервера")
        self.dialog.geometry("500x400")
        self.dialog.resizable(False, False)
        self.dialog.transient(self.parent)
        self.dialog.grab_set()
        
        # Фрейм настроек
        settings_frame = ttk.Frame(self.dialog, padding="20")
        settings_frame.pack(fill=tk.BOTH, expand=True)
        
        # SMTP сервер
        ttk.Label(settings_frame, text="SMTP сервер:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.server_entry = ttk.Entry(settings_frame, width=40)
        self.server_entry.grid(row=0, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.server_entry.insert(0, self.settings.get('smtp_server', 'smtp.gmail.com'))
        
        # Порт
        ttk.Label(settings_frame, text="Порт:").grid(row=1, column=0, sticky=tk.W, pady=5)
        self.port_entry = ttk.Entry(settings_frame, width=10)
        self.port_entry.grid(row=1, column=1, padx=5, pady=5, sticky=tk.W)
        self.port_entry.insert(0, str(self.settings.get('smtp_port', '587')))
        
        # SSL
        self.use_ssl_var = tk.BooleanVar(value=self.settings.get('use_ssl', True))
        ttk.Checkbutton(
            settings_frame, text="Использовать SSL (порт 465)",
            variable=self.use_ssl_var
        ).grid(row=2, column=1, padx=5, pady=5, sticky=tk.W)
        
        # TLS
        self.use_tls_var = tk.BooleanVar(value=self.settings.get('use_tls', True))
        ttk.Checkbutton(
            settings_frame, text="Использовать TLS (порт 587)",
            variable=self.use_tls_var
        ).grid(row=3, column=1, padx=5, pady=5, sticky=tk.W)
        
        # Логин
        ttk.Label(settings_frame, text="Email (логин):").grid(row=4, column=0, sticky=tk.W, pady=5)
        self.login_entry = ttk.Entry(settings_frame, width=40)
        self.login_entry.grid(row=4, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.login_entry.insert(0, self.settings.get('email_login', ''))
        
        # Пароль
        ttk.Label(settings_frame, text="Пароль приложения:").grid(row=5, column=0, sticky=tk.W, pady=5)
        self.password_entry = ttk.Entry(settings_frame, width=40, show="*")
        self.password_entry.grid(row=5, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.password_entry.insert(0, self.settings.get('email_password', ''))
        
        # Имя отправителя
        ttk.Label(settings_frame, text="Имя отправителя:").grid(row=6, column=0, sticky=tk.W, pady=5)
        self.name_entry = ttk.Entry(settings_frame, width=40)
        self.name_entry.grid(row=6, column=1, padx=5, pady=5, sticky=(tk.W, tk.E))
        self.name_entry.insert(0, self.settings.get('sender_name', ''))
        
        # Пресеты
        ttk.Label(settings_frame, text="Быстрые настройки:").grid(row=7, column=0, sticky=tk.W, pady=10)
        preset_frame = ttk.Frame(settings_frame)
        preset_frame.grid(row=7, column=1, padx=5, pady=5, sticky=tk.W)
        
        ttk.Button(preset_frame, text="Gmail", command=self._set_gmail).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="Yandex", command=self._set_yandex).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="Mail.ru", command=self._set_mailru).pack(side=tk.LEFT, padx=2)
        ttk.Button(preset_frame, text="Outlook", command=self._set_outlook).pack(side=tk.LEFT, padx=2)
        
        # Кнопки
        button_frame = ttk.Frame(settings_frame)
        button_frame.grid(row=8, column=0, columnspan=2, pady=20)
        
        ttk.Button(button_frame, text="Тест подключения", command=self._test_connection).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Сохранить", command=self._save).pack(side=tk.LEFT, padx=5)
        ttk.Button(button_frame, text="Отмена", command=self.dialog.destroy).pack(side=tk.LEFT, padx=5)
        
        # Настройка веса колонок
        settings_frame.grid_columnconfigure(1, weight=1)
        
        # Центрирование
        self.dialog.update_idletasks()
        x = (self.parent.winfo_screenwidth() - self.dialog.winfo_width()) // 2
        y = (self.parent.winfo_screenheight() - self.dialog.winfo_height()) // 2
        self.dialog.geometry(f"+{x}+{y}")
    
    def _set_gmail(self):
        """Настройки Gmail"""
        self.server_entry.delete(0, tk.END)
        self.server_entry.insert(0, 'smtp.gmail.com')
        self.port_entry.delete(0, tk.END)
        self.port_entry.insert(0, '587')
        self.use_ssl_var.set(False)
        self.use_tls_var.set(True)
    
    def _set_yandex(self):
        """Настройки Yandex"""
        self.server_entry.delete(0, tk.END)
        self.server_entry.insert(0, 'smtp.yandex.ru')
        self.port_entry.delete(0, tk.END)
        self.port_entry.insert(0, '465')
        self.use_ssl_var.set(True)
        self.use_tls_var.set(False)
    
    def _set_mailru(self):
        """Настройки Mail.ru"""
        self.server_entry.delete(0, tk.END)
        self.server_entry.insert(0, 'smtp.mail.ru')
        self.port_entry.delete(0, tk.END)
        self.port_entry.insert(0, '465')
        self.use_ssl_var.set(True)
        self.use_tls_var.set(False)
    
    def _set_outlook(self):
        """Настройки Outlook/Hotmail"""
        self.server_entry.delete(0, tk.END)
        self.server_entry.insert(0, 'smtp.office365.com')
        self.port_entry.delete(0, tk.END)
        self.port_entry.insert(0, '587')
        self.use_ssl_var.set(False)
        self.use_tls_var.set(True)
    
    def _test_connection(self):
        """Тест подключения"""
        config = self._get_config()
        if not config:
            messagebox.showerror("Ошибка", "Заполните все поля")
            return
        
        from backend.smtp_service import SMTPService
        service = SMTPService(config)
        success, message = service.test_connection()
        
        if success:
            messagebox.showinfo("Успех", f"✅ {message}")
        else:
            messagebox.showerror("Ошибка", f"❌ {message}")
    
    def _get_config(self) -> Optional[SMTPConfig]:
        """Получает конфигурацию из полей"""
        try:
            return SMTPConfig(
                smtp_server=self.server_entry.get().strip(),
                smtp_port=int(self.port_entry.get()),
                email_login=self.login_entry.get().strip(),
                email_password=self.password_entry.get(),
                use_ssl=self.use_ssl_var.get(),
                use_tls=self.use_tls_var.get(),
                sender_name=self.name_entry.get().strip()
            )
        except ValueError:
            return None
    
    def _save(self):
        """Сохраняет настройки"""
        config = self._get_config()
        if config:
            self.result = config
            self.dialog.destroy()
        else:
            messagebox.showerror("Ошибка", "Проверьте правильность заполнения полей")
    
    def show(self) -> Optional[SMTPConfig]:
        """
        Показывает диалог и возвращает результат.
        
        Returns:
            SMTPConfig или None
        """
        self.parent.wait_window(self.dialog)
        return self.result
