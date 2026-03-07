"""UI компоненты приложения"""

import tkinter as tk
from tkinter import ttk


class ContextMenuMixin:
    """Миксин для добавления контекстного меню к текстовым виджетам"""
    
    def add_context_menu(self, text_widget: tk.Text):
        """
        Добавляет контекстное меню с функцией копирования.
        
        Args:
            text_widget: Текстовый виджет
        """
        context_menu = tk.Menu(text_widget, tearoff=0)
        context_menu.add_command(
            label="Копировать выделенное", 
            command=lambda: self._copy_text(text_widget)
        )
        context_menu.add_command(
            label="Копировать все", 
            command=lambda: self._copy_all(text_widget)
        )
        
        # Привязка к правой кнопке мыши
        text_widget.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))
        
        # Горячие клавиши
        text_widget.bind("<Control-c>", lambda e: self._copy_text(text_widget))
        text_widget.bind("<Control-C>", lambda e: self._copy_text(text_widget))
        text_widget.bind("<Control-Insert>", lambda e: self._copy_text(text_widget))
        text_widget.bind("<Control-a>", lambda e: text_widget.tag_add("sel", "1.0", "end"))
        text_widget.bind("<Control-A>", lambda e: text_widget.tag_add("sel", "1.0", "end"))
        text_widget.bind("<Control-Shift-c>", lambda e: self._copy_all(text_widget))
        text_widget.bind("<Control-Shift-C>", lambda e: self._copy_all(text_widget))
    
    def _copy_text(self, text_widget: tk.Text):
        """Копирует выделенный текст"""
        try:
            selected_text = text_widget.selection_get()
            if selected_text:
                text_widget.master.clipboard_clear()
                text_widget.master.clipboard_append(selected_text)
                text_widget.master.update()
        except tk.TclError:
            pass
    
    def _copy_all(self, text_widget: tk.Text):
        """Копирует весь текст"""
        try:
            content = text_widget.get("1.0", tk.END).strip()
            if content:
                text_widget.master.clipboard_clear()
                text_widget.master.clipboard_append(content)
                text_widget.master.update()
        except Exception:
            pass


class SettingsFrame(ttk.LabelFrame, ContextMenuMixin):
    """Фрейм настроек рассылки"""
    
    def __init__(self, parent, callbacks: dict):
        super().__init__(parent, text="Настройки рассылки", padding="10")
        self.callbacks = callbacks
        
        # Переменные
        self.email_account = tk.StringVar()
        self.excel_path = tk.StringVar()
        self.folder_path_1 = tk.StringVar()
        self.folder_path_2 = tk.StringVar()
        self.folder_path_3 = tk.StringVar()
        self.thread_count = tk.IntVar(value=3)
        self.email_subject = tk.StringVar()
        
        self._setup_ui()
    
    def _setup_ui(self):
        """Создаёт элементы UI"""
        # Аккаунт Outlook
        ttk.Label(self, text="Аккаунт Outlook:").grid(row=0, column=0, sticky=tk.W, pady=5)
        self.account_combo = ttk.Combobox(
            self, textvariable=self.email_account, 
            width=60, state="readonly"
        )
        self.account_combo.grid(row=0, column=1, padx=5, pady=5, columnspan=2, sticky=(tk.W, tk.E))
        
        # Excel файл
        ttk.Label(self, text="Путь к Excel файлу:").grid(row=1, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self, textvariable=self.excel_path, width=65).grid(
            row=1, column=1, padx=5, pady=5, sticky=(tk.W, tk.E)
        )
        ttk.Button(self, text="Обзор", command=self.callbacks.get('browse_excel'), width=10).grid(
            row=1, column=2, padx=5, pady=5
        )
        
        # Папка 1
        ttk.Label(self, text="Папка с расчетными листами:").grid(row=2, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self, textvariable=self.folder_path_1, width=65).grid(
            row=2, column=1, padx=5, pady=5, sticky=(tk.W, tk.E)
        )
        ttk.Button(self, text="Обзор", command=lambda: self.callbacks.get('browse_folder')(1), width=10).grid(
            row=2, column=2, padx=5, pady=5
        )
        
        # Папка 2
        ttk.Label(self, text="Папка с реестрами выдачи:").grid(row=3, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self, textvariable=self.folder_path_2, width=65).grid(
            row=3, column=1, padx=5, pady=5, sticky=(tk.W, tk.E)
        )
        ttk.Button(self, text="Обзор", command=lambda: self.callbacks.get('browse_folder')(2), width=10).grid(
            row=3, column=2, padx=5, pady=5
        )
        
        # Папка 3
        ttk.Label(self, text="Папка с доп. файлами:").grid(row=4, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self, textvariable=self.folder_path_3, width=50).grid(
            row=4, column=1, padx=5, pady=5, sticky=(tk.W, tk.E)
        )
        ttk.Button(self, text="Обзор", command=lambda: self.callbacks.get('browse_folder')(3), width=10).grid(
            row=4, column=2, padx=5, pady=5
        )
        
        # Потоки
        ttk.Label(self, text="Количество потоков:").grid(row=5, column=0, sticky=tk.W, pady=5)
        ttk.Spinbox(self, from_=1, to=10, textvariable=self.thread_count, width=5).grid(
            row=5, column=1, padx=5, pady=5, sticky=tk.W
        )
        ttk.Label(self, text="(Рекомендуется 3-5 потоков)").grid(
            row=5, column=1, padx=80, pady=5, sticky=tk.W
        )
        
        # Тема
        ttk.Label(self, text="Тема письма:").grid(row=6, column=0, sticky=tk.W, pady=5)
        ttk.Entry(self, textvariable=self.email_subject, width=50).grid(
            row=6, column=1, padx=5, pady=5, columnspan=2, sticky=(tk.W, tk.E)
        )
        
        # Текст письма
        ttk.Label(self, text="Текст письма:").grid(row=7, column=0, sticky=tk.NW, pady=5)
        self.body_text = tk.Text(self, height=8, width=50, wrap=tk.WORD)
        self.body_text.grid(row=7, column=1, padx=5, pady=5, columnspan=2, sticky=(tk.W, tk.E))
        
        # Контекстное меню
        self.add_context_menu(self.body_text)
        
        # Настройка веса колонок
        self.grid_columnconfigure(1, weight=1)
    
    def set_email_body(self, text: str):
        """Устанавливает текст письма"""
        self.body_text.delete("1.0", tk.END)
        self.body_text.insert("1.0", text)
    
    def get_email_body(self) -> str:
        """Получает текст письма"""
        return self.body_text.get("1.0", tk.END).strip()
    
    def set_account_values(self, values: list):
        """Устанавливает значения аккаунтов"""
        self.account_combo['values'] = values
        if values:
            self.email_account.set(values[0])
