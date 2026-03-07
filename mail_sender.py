import os
import sys
import logging
import json
import random
from pathlib import Path
import pandas as pd
import win32com.client as win32
import pythoncom  # Добавляем для инициализации COM в потоках
from tkinter import *
from tkinter import ttk, filedialog, messagebox
import tkinter as tk
from datetime import datetime, timedelta
from typing import List, Dict
import threading
import queue
import time
from concurrent.futures import ThreadPoolExecutor, as_completed

# Настройка логирования
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler("email_sender.log"),
        logging.StreamHandler()
    ]
)


class EmailSenderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Почтовая рассылка расчетных листов ФитоФарм")
        self.root.geometry("720x800")  # Увеличенный размер для новых элементов
        self.root.resizable(False, False)  # Запрет на изменение размера окна

        # Файл для сохранения настроек
        self.settings_file = "settings.json"
        
        # Переменные для хранения путей и данных
        self.excel_path = StringVar()
        self.email_account = StringVar()
        self.thread_count = IntVar(value=3)  # Количество потоков по умолчанию

        # Автоматическая тема письма
        previous_month = (datetime.now() - timedelta(days=30)).month
        month_names = {
            1: "Январь", 2: "Февраль", 3: "Март", 4: "Апрель",
            5: "Май", 6: "Июнь", 7: "Июль", 8: "Август",
            9: "Сентябрь", 10: "Октябрь", 11: "Ноябрь", 12: "Декабрь"
        }
        current_year = datetime.now().strftime("%Y")
        self.email_subject = StringVar(value=f"Расчетные листы {month_names[previous_month]} {current_year}")

        # Стандартный текст письма
        self.email_body = ("Добрый день!\n\n"
                           "Сообщение сформировано автоматически, отвечать на него не нужно.\n\n"
                           "Пароль от файлов в скором времени будет направлен вашему Куратору")

        # Пути к трем папкам с файлами
        self.folder_path_1 = StringVar()
        self.folder_path_2 = StringVar()
        self.folder_path_3 = StringVar()

        # Переменные для управления рассылкой
        self.is_paused = False
        self.is_cancelled = False
        self.total_emails = 0
        self.sent_count = 0
        self.failed_count = 0
        self.executor = None
        self.futures = []
        
        # Очередь для обновления UI из потоков
        self.ui_queue = queue.Queue()

        self.setup_ui()
        self.load_outlook_accounts()
        self.load_settings()  # Загружаем сохраненные настройки
        
        # Запускаем обработку очереди UI
        self.root.after(100, self.process_ui_queue)

    def setup_ui(self):
        # Основной фрейм
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.pack(fill=BOTH, expand=True)

        # Секция настроек (группировка)
        settings_frame = ttk.LabelFrame(main_frame, text="Настройки рассылки", padding="10")
        settings_frame.pack(fill=X, pady=(0, 10))

        # Выбор аккаунта Outlook
        ttk.Label(settings_frame, text="Аккаунт Outlook:").grid(row=0, column=0, sticky=W, pady=5)
        self.account_combo = ttk.Combobox(settings_frame, textvariable=self.email_account, width=60, state="readonly")
        self.account_combo.grid(row=0, column=1, padx=5, pady=5, columnspan=2, sticky=(W, E))

        # Путь к Excel файлу
        ttk.Label(settings_frame, text="Путь к Excel файлу:").grid(row=1, column=0, sticky=W, pady=5)
        ttk.Entry(settings_frame, textvariable=self.excel_path, width=65).grid(row=1, column=1, padx=5, pady=5,
                                                                              sticky=(W, E))
        ttk.Button(settings_frame, text="Обзор", command=self.browse_excel, width=10).grid(row=1, column=2, padx=5, pady=5)

        # Путь к папке 1
        ttk.Label(settings_frame, text="Папка с расчетными листами:").grid(row=2, column=0, sticky=W, pady=5)
        ttk.Entry(settings_frame, textvariable=self.folder_path_1, width=65).grid(row=2, column=1, padx=5, pady=5,
                                                                                 sticky=(W, E))
        ttk.Button(settings_frame, text="Обзор", command=lambda: self.browse_folder(1), width=10).grid(row=2, column=2,
                                                                                                      padx=5, pady=5)

        # Путь к папке 2
        ttk.Label(settings_frame, text="Папка с реестрами выдачи:").grid(row=3, column=0, sticky=W, pady=5)
        ttk.Entry(settings_frame, textvariable=self.folder_path_2, width=65).grid(row=3, column=1, padx=5, pady=5,
                                                                                 sticky=(W, E))
        ttk.Button(settings_frame, text="Обзор", command=lambda: self.browse_folder(2), width=10).grid(row=3, column=2,
                                                                                                      padx=5, pady=5)

        # Путь к папке 3
        ttk.Label(settings_frame, text="Папка с доп. файлами:").grid(row=4, column=0, sticky=W, pady=5)
        ttk.Entry(settings_frame, textvariable=self.folder_path_3, width=50).grid(row=4, column=1, padx=5, pady=5,
                                                                                 sticky=(W, E))
        ttk.Button(settings_frame, text="Обзор", command=lambda: self.browse_folder(3), width=10).grid(row=4, column=2,
                                                                                                      padx=5, pady=5)

        # Настройки многопоточности
        ttk.Label(settings_frame, text="Количество потоков:").grid(row=5, column=0, sticky=W, pady=5)
        self.thread_spinbox = ttk.Spinbox(settings_frame, from_=1, to=10, textvariable=self.thread_count, width=5)
        self.thread_spinbox.grid(row=5, column=1, padx=5, pady=5, sticky=W)
        ttk.Label(settings_frame, text="(Рекомендуется 3-5 потоков)").grid(row=5, column=1, padx=80, pady=5, sticky=W)

        # Тема письма
        ttk.Label(settings_frame, text="Тема письма:").grid(row=6, column=0, sticky=W, pady=5)
        ttk.Entry(settings_frame, textvariable=self.email_subject, width=50).grid(row=6, column=1, padx=5, pady=5,
                                                                                 columnspan=2, sticky=(W, E))

        # Текст письма
        ttk.Label(settings_frame, text="Текст письма:").grid(row=7, column=0, sticky=NW, pady=5)
        self.body_text = Text(settings_frame, height=8, width=50)
        self.body_text.grid(row=7, column=1, padx=5, pady=5, columnspan=2, sticky=(W, E))
        self.body_text.insert("1.0", self.email_body)
        
        
        # Добавляем контекстное меню для текста письма
        self.add_context_menu(self.body_text)

        # Секция управления рассылкой
        control_frame = ttk.LabelFrame(main_frame, text="Управление рассылкой", padding="10")
        control_frame.pack(fill=X, pady=(0, 10))

        # Статус и счетчики
        status_frame = ttk.Frame(control_frame)
        status_frame.pack(fill=X, pady=(0, 10))

        self.status_label = ttk.Label(status_frame, text="Готов к отправке", font=("TkDefaultFont", 10, "bold"))
        self.status_label.pack(side=LEFT)

        self.counter_label = ttk.Label(status_frame, text="0/0", font=("TkDefaultFont", 10))
        self.counter_label.pack(side=RIGHT)

        # Прогресс-бар
        self.progress_var = DoubleVar()
        self.progress_bar = ttk.Progressbar(control_frame, variable=self.progress_var, maximum=100, mode='determinate')
        self.progress_bar.pack(fill=X, pady=(0, 10))

        # Кнопки управления
        button_frame = ttk.Frame(control_frame)
        button_frame.pack(fill=X)

        self.send_button = ttk.Button(button_frame, text="Начать рассылку", command=self.send_emails)
        self.send_button.pack(side=LEFT, padx=(0, 5))

        self.pause_button = ttk.Button(button_frame, text="Пауза", command=self.toggle_pause, state=DISABLED)
        self.pause_button.pack(side=LEFT, padx=(0, 5))

        self.preview_button = ttk.Button(button_frame, text="Предварительный просмотр", command=self.preview_email)
        self.preview_button.pack(side=LEFT, padx=(0, 5))
        
        
        self.cancel_button = ttk.Button(button_frame, text="Отмена", command=self.cancel_send, state=DISABLED)
        self.cancel_button.pack(side=LEFT)

        # Секция лога
        log_frame = ttk.LabelFrame(main_frame, text="Лог выполнения", padding="10")
        log_frame.pack(fill=BOTH, expand=True)

        # Текстовое поле для лога
        self.log_text = Text(log_frame, height=12, width=80, wrap=WORD)
        self.log_text.pack(side=LEFT, fill=BOTH, expand=True, pady=(0, 5))
        
        # Добавляем контекстное меню для лога
        self.add_context_menu(self.log_text)

        # Scrollbar для лога
        log_scrollbar = ttk.Scrollbar(log_frame, orient=VERTICAL, command=self.log_text.yview)
        log_scrollbar.pack(side=RIGHT, fill=Y)
        self.log_text.configure(yscrollcommand=log_scrollbar.set)

        # Центрирование окна
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - self.root.winfo_width()) // 2
        y = (self.root.winfo_screenheight() - self.root.winfo_height()) // 2
        self.root.geometry(f"+{x}+{y}")
    
    def add_context_menu(self, text_widget):
        """Добавляет контекстное меню с функцией копирования для текстового виджета"""
        # Создаем контекстное меню
        context_menu = Menu(text_widget, tearoff=0)
        context_menu.add_command(label="Копировать выделенное", command=lambda: self.copy_text(text_widget))
        context_menu.add_command(label="Копировать все", command=lambda: self.copy_all_from_widget(text_widget))
        
        # Привязываем контекстное меню к правой кнопке мыши
        text_widget.bind("<Button-3>", lambda e: context_menu.tk_popup(e.x_root, e.y_root))
        
        # Привязываем горячую клавишу Ctrl+C для копирования выделенного текста
        text_widget.bind("<Control-c>", lambda e: self.copy_text(text_widget))
        text_widget.bind("<Control-C>", lambda e: self.copy_text(text_widget))
        
        # Привязываем горячую клавишу Ctrl+Insert для копирования
        text_widget.bind("<Control-Insert>", lambda e: self.copy_text(text_widget))
        
        # Привязываем горячую клавишу Ctrl+A для выделения всего текста
        text_widget.bind("<Control-a>", lambda e: text_widget.tag_add("sel", "1.0", "end"))
        text_widget.bind("<Control-A>", lambda e: text_widget.tag_add("sel", "1.0", "end"))
        
        # Привязываем горячую клавишу Ctrl+Shift+C для копирования всего текста
        text_widget.bind("<Control-Shift-c>", lambda e: self.copy_all_from_widget(text_widget))
        text_widget.bind("<Control-Shift-C>", lambda e: self.copy_all_from_widget(text_widget))
    
    def copy_text(self, text_widget):
        """Копирует выделенный текст в буфер обмена"""
        try:
            # Получаем выделенный текст
            selected_text = text_widget.selection_get()
            if selected_text:
                # Копируем в буфер обмена
                self.root.clipboard_clear()
                self.root.clipboard_append(selected_text)
                self.root.update()  # Обновляем буфер обмена
                self.log_message("Текст скопирован в буфер обмена", "INFO")
        except tk.TclError:
            # Нет выделенного текста
            self.log_message("Нет выделенного текста для копирования", "WARNING")
    
    def copy_all_from_widget(self, text_widget):
        """Копирует весь текст из виджета в буфер обмена"""
        try:
            content = text_widget.get("1.0", END).strip()
            if content:
                self.root.clipboard_clear()
                self.root.clipboard_append(content)
                self.root.update()
                self.log_message("Весь текст скопирован в буфер обмена", "INFO")
            else:
                self.log_message("Текст пуст, нечего копировать", "WARNING")
        except Exception as e:
            self.log_message(f"Ошибка копирования текста: {str(e)}", "ERROR")

    def load_outlook_accounts(self):
        """Загружает список email аккаунтов из Outlook"""
        try:
            outlook = win32.Dispatch('Outlook.Application')
            namespace = outlook.GetNamespace("MAPI")
            accounts = []

            for account in namespace.Accounts:
                accounts.append(account.SmtpAddress)

            self.account_combo['values'] = accounts
            if accounts:
                self.email_account.set(accounts[0])

            self.log_message(f"Найдено {len(accounts)} аккаунтов Outlook")

        except Exception as e:
            self.log_message(f"Ошибка загрузки аккаунтов Outlook: {str(e)}")

    def browse_folder(self, folder_number: int):
        folder = filedialog.askdirectory(title=f"Выберите папку {folder_number} с файлами")
        if folder:
            if folder_number == 1:
                self.folder_path_1.set(folder)
            elif folder_number == 2:
                self.folder_path_2.set(folder)
            elif folder_number == 3:
                self.folder_path_3.set(folder)

    def browse_excel(self):
        file = filedialog.askopenfilename(
            title="Выберите Excel файл",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if file:
            self.excel_path.set(file)
    

    def log_message(self, message: str, level: str = "INFO"):
        """Добавляет сообщение в лог с цветовой индикацией"""
        timestamp = datetime.now().strftime("%H:%M:%S")
        formatted_message = f"[{timestamp}] {level}: {message}\n"
        
        # Добавляем сообщение
        self.log_text.insert(END, formatted_message)
        
        # Прокручиваем вниз
        self.log_text.see(END)
        self.root.update_idletasks()

    def process_ui_queue(self):
        """Обрабатывает очередь обновлений UI из потоков"""
        try:
            while True:
                item = self.ui_queue.get_nowait()
                if item['type'] == 'log':
                    self.log_message(item['message'], item['level'])
                elif item['type'] == 'status':
                    self.update_status(item['message'], item['level'])
                elif item['type'] == 'progress':
                    self.update_progress(item['current'], item['total'])
                elif item['type'] == 'complete':
                    self.send_button.config(state='normal')
                    self.pause_button.config(state='disabled')
                    self.cancel_button.config(state='disabled')
                    self.preview_button.config(state='normal')
                    if item['cancelled']:
                        self.update_status(f"Рассылка отменена. Отправлено: {item['success_count']}/{item['total']}", "WARNING")
                    else:
                        self.update_status(f"Рассылка завершена: {item['success_count']} отправлено, {item['failed_count']} ошибок", "INFO")
                        messagebox.showinfo("Завершено",
                                          f"Рассылка завершена!\nОтправлено: {item['success_count']}/{item['total']}\nОшибок: {item['failed_count']}")
                self.ui_queue.task_done()
        except queue.Empty:
            pass
        self.root.after(100, self.process_ui_queue)

    def send_email_worker(self, row, index, total):
        """Рабочая функция для отправки одного письма"""
        try:
            email = row.get('email', '')
            file_01 = row.get('файл_01', '')
            file_02 = row.get('файл_02', '')
            file_03 = row.get('файл_03', '')

            if not email:
                self.ui_queue.put({
                    'type': 'log',
                    'message': f"Пропуск строки {index}: нет email адреса",
                    'level': 'WARNING'
                })
                return {'success': False, 'email': email, 'error': 'Нет email адреса'}

            # Получаем текст письма
            email_body = self.body_text.get("1.0", END).strip()
            email_subject = self.email_subject.get()

            # Инициализируем COM для каждого потока
            pythoncom.CoInitialize()
            
            try:
                # Создаем новый экземпляр Outlook для каждого потока
                outlook = win32.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = email_subject
                mail.Body = email_body

                # Ищем и прикрепляем файлы из всех трех папок
                attached_files = []

                # Файл из папки 1
                if file_01 and self.folder_path_1.get():
                    file_path = self.find_file_in_folder(self.folder_path_1.get(), file_01)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))

                # Файл из папки 2
                if file_02 and self.folder_path_2.get():
                    file_path = self.find_file_in_folder(self.folder_path_2.get(), file_02)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))

                # Файл из папки 3
                if file_03 and self.folder_path_3.get():
                    file_path = self.find_file_in_folder(self.folder_path_3.get(), file_03)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))

                if attached_files:
                    self.ui_queue.put({
                        'type': 'log',
                        'message': f"Прикреплены файлы для {email}: {', '.join(attached_files)}",
                        'level': 'INFO'
                    })
                else:
                    self.ui_queue.put({
                        'type': 'log',
                        'message': f"Для {email} не найдены файлы для прикрепления",
                        'level': 'WARNING'
                    })
                    return {'success': False, 'email': email, 'error': 'Нет файлов для прикрепления'}

                # Отправляем письмо
                mail.Send()
                
                self.ui_queue.put({
                    'type': 'log',
                    'message': f"Письмо {index}/{total} отправлено: {email}",
                    'level': 'INFO'
                })
                
                return {'success': True, 'email': email, 'error': None}

            except Exception as e:
                error_msg = f"Ошибка отправки письма для {email}: {str(e)}"
                self.ui_queue.put({
                    'type': 'log',
                    'message': error_msg,
                    'level': 'ERROR'
                })
                return {'success': False, 'email': email, 'error': str(e)}

            finally:
                # Освобождаем COM в любом случае
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            error_msg = f"Критическая ошибка в потоке для {email}: {str(e)}"
            self.ui_queue.put({
                'type': 'log',
                'message': error_msg,
                'level': 'ERROR'
            })
            return {'success': False, 'email': email, 'error': str(e)}

    def update_status(self, message: str, level: str = "INFO"):
        """Обновляет статус и лог"""
        self.status_label.config(text=message)
        self.log_message(message, level)

    def update_progress(self, current: int, total: int):
        """Обновляет прогресс-бар и счетчик"""
        self.total_emails = total
        self.sent_count = current
        self.counter_label.config(text=f"{current}/{total}")
        
        if total > 0:
            progress = (current / total) * 100
            self.progress_var.set(progress)

    def read_excel_data(self) -> List[Dict]:
        """Читает данные из Excel файла"""
        try:
            df = pd.read_excel(self.excel_path.get())

            # Проверяем наличие обязательных колонок
            required_columns = ['email', 'файл_01', 'файл_02', 'файл_03']
            missing_columns = [col for col in required_columns if col not in df.columns]

            if missing_columns:
                raise ValueError(f"Отсутствуют обязательные колонки: {', '.join(missing_columns)}")

            # Преобразуем в список словарей
            data = df.to_dict('records')
            self.log_message(f"Прочитано {len(data)} записей из Excel")
            return data

        except Exception as e:
            self.log_message(f"Ошибка чтения Excel: {str(e)}")
            raise

    def find_file_in_folder(self, folder_path: str, filename: str) -> str:
        """Ищет файл в указанной папке"""
        if not folder_path or not filename or pd.isna(filename):
            return None

        folder = Path(folder_path)
        if not folder.exists():
            self.log_message(f"Папка не существует: {folder_path}")
            return None

        for file in folder.iterdir():
            if file.is_file() and file.name.lower() == str(filename).lower():
                return str(file)

        return None

    def toggle_pause(self):
        """Переключение между паузой и продолжением рассылки"""
        self.is_paused = not self.is_paused
        if self.is_paused:
            self.pause_button.config(text="Продолжить")
            self.update_status("Рассылка приостановлена", "WARNING")
            self.log_message("Рассылка поставлена на паузу", "WARNING")
        else:
            self.pause_button.config(text="Пауза")
            self.update_status("Рассылка продолжается", "INFO")
            self.log_message("Рассылка продолжается", "INFO")

    def cancel_send(self):
        """Отмена рассылки"""
        self.is_cancelled = True
        self.update_status("Рассылка отменена", "WARNING")
        self.log_message("Рассылка отменена пользователем", "WARNING")

    def preview_email(self):
        """Предварительный просмотр первого письма"""
        try:
            # Проверка обязательных полей
            if not self.email_account.get():
                messagebox.showerror("Ошибка", "Выберите аккаунт Outlook для отправки")
                return

            if not self.excel_path.get():
                messagebox.showerror("Ошибка", "Укажите путь к Excel файлу с адресами")
                return

            # Проверка хотя бы одной папки
            folders = [self.folder_path_1.get(), self.folder_path_2.get(), self.folder_path_3.get()]
            if not any(folders):
                messagebox.showerror("Ошибка", "Укажите хотя бы одну папку с файлами")
                return

            # Читаем данные
            email_data = self.read_excel_data()

            if not email_data:
                messagebox.showerror("Ошибка", "Нет данных для предварительного просмотра")
                return

            # Берем случайное письмо для просмотра
            row = random.choice(email_data)
            email = row.get('email', '')
            file_01 = row.get('файл_01', '')
            file_02 = row.get('файл_02', '')
            file_03 = row.get('файл_03', '')

            if not email:
                messagebox.showerror("Ошибка", "Нет email адреса для предварительного просмотра")
                return

            # Получаем текст письма
            email_body = self.body_text.get("1.0", END).strip()
            email_subject = self.email_subject.get()

            # Инициализируем COM для предварительного просмотра
            pythoncom.CoInitialize()
            
            try:
                # Подключаемся к Outlook
                outlook = win32.Dispatch('Outlook.Application')
                mail = outlook.CreateItem(0)
                mail.To = email
                mail.Subject = email_subject
                mail.Body = email_body

                # Ищем и прикрепляем файлы из всех трех папок
                attached_files = []

                # Файл из папки 1
                if file_01 and self.folder_path_1.get():
                    file_path = self.find_file_in_folder(self.folder_path_1.get(), file_01)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))

                # Файл из папки 2
                if file_02 and self.folder_path_2.get():
                    file_path = self.find_file_in_folder(self.folder_path_2.get(), file_02)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))

                # Файл из папки 3
                if file_03 and self.folder_path_3.get():
                    file_path = self.find_file_in_folder(self.folder_path_3.get(), file_03)
                    if file_path:
                        mail.Attachments.Add(file_path)
                        attached_files.append(os.path.basename(file_path))

                # Показываем письмо
                mail.Display()
                
                self.log_message(f"Предварительный просмотр письма для: {email}")
                if attached_files:
                    self.log_message(f"Прикрепленные файлы: {', '.join(attached_files)}")
                
            except Exception as e:
                self.log_message(f"Ошибка предварительного просмотра: {str(e)}", "ERROR")
                messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")
            
            finally:
                # Освобождаем COM в любом случае
                try:
                    pythoncom.CoUninitialize()
                except:
                    pass

        except Exception as e:
            self.log_message(f"Критическая ошибка в предварительном просмотре: {str(e)}", "ERROR")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

    def send_emails(self):
        """Основная функция отправки писем (многопоточная)"""
        try:
            # Проверка обязательных полей
            if not self.email_account.get():
                messagebox.showerror("Ошибка", "Выберите аккаунт Outlook для отправки")
                return

            if not self.excel_path.get():
                messagebox.showerror("Ошибка", "Укажите путь к Excel файлу с адресами")
                return

            # Проверка хотя бы одной папки
            folders = [self.folder_path_1.get(), self.folder_path_2.get(), self.folder_path_3.get()]
            if not any(folders):
                messagebox.showerror("Ошибка", "Укажите хотя бы одну папку с файлами")
                return

            # Сбрасываем флаги
            self.is_paused = False
            self.is_cancelled = False

            # Блокируем кнопки
            self.send_button.config(state='disabled')
            self.pause_button.config(state='normal')
            self.cancel_button.config(state='normal')
            self.preview_button.config(state='disabled')

            # Читаем данные
            email_data = self.read_excel_data()

            if not email_data:
                self.send_button.config(state='normal')
                self.pause_button.config(state='disabled')
                self.cancel_button.config(state='disabled')
                self.preview_button.config(state='normal')
                return

            total_emails = len(email_data)
            self.total_emails = total_emails

            # Обновляем прогресс
            self.update_progress(0, total_emails)
            self.update_status(f"Подготовка к отправке: 0/{total_emails}")

            # Запускаем многопоточную отправку в отдельном потоке
            threading.Thread(target=self.send_emails_multithreaded, args=(email_data, total_emails), daemon=True).start()

        except Exception as e:
            self.send_button.config(state='normal')
            self.pause_button.config(state='disabled')
            self.cancel_button.config(state='disabled')
            self.preview_button.config(state='normal')
            self.update_status(f"Критическая ошибка: {str(e)}", "ERROR")
            self.log_message(f"Критическая ошибка: {str(e)}", "ERROR")
            messagebox.showerror("Ошибка", f"Произошла ошибка: {str(e)}")

    def send_emails_multithreaded(self, email_data, total_emails):
        """Многопоточная отправка писем"""
        try:
            # Создаем пул потоков
            max_workers = min(self.thread_count.get(), total_emails)
            success_count = 0
            failed_count = 0

            self.ui_queue.put({
                'type': 'status',
                'message': f"Запуск {max_workers} потоков для отправки {total_emails} писем",
                'level': 'INFO'
            })

            # Отправляем письма параллельно
            with ThreadPoolExecutor(max_workers=max_workers) as executor:
                # Создаем задачи для каждого письма
                futures = []
                for i, row in enumerate(email_data, 1):
                    future = executor.submit(self.send_email_worker, row, i, total_emails)
                    futures.append(future)

                # Обрабатываем результаты по мере завершения
                for i, future in enumerate(as_completed(futures)):
                    if self.is_cancelled:
                        # Отменяем оставшиеся задачи
                        for f in futures:
                            f.cancel()
                        break
                    
                    # Пауза не нужна здесь, так как потоки работают независимо
                    result = future.result()
                    
                    if result['success']:
                        success_count += 1
                    else:
                        failed_count += 1
                    
                    # Обновляем прогресс
                    self.ui_queue.put({
                        'type': 'progress',
                        'current': i + 1,
                        'total': total_emails
                    })

            # Завершение рассылки
            self.ui_queue.put({
                'type': 'complete',
                'success_count': success_count,
                'failed_count': failed_count,
                'total': total_emails,
                'cancelled': self.is_cancelled
            })

        except Exception as e:
            self.ui_queue.put({
                'type': 'status',
                'message': f"Критическая ошибка в потоке: {str(e)}",
                'level': 'ERROR'
            })
            self.ui_queue.put({
                'type': 'complete',
                'success_count': success_count,
                'failed_count': failed_count,
                'total': total_emails,
                'cancelled': True
            })

    def save_settings(self):
        """Сохраняет настройки в файл"""
        try:
            settings = {
                'excel_path': self.excel_path.get(),
                'email_account': self.email_account.get(),
                'thread_count': self.thread_count.get(),
                'email_subject': self.email_subject.get(),
                'email_body': self.body_text.get("1.0", END).strip(),
                'folder_path_1': self.folder_path_1.get(),
                'folder_path_2': self.folder_path_2.get(),
                'folder_path_3': self.folder_path_3.get()
            }
            
            with open(self.settings_file, 'w', encoding='utf-8') as f:
                json.dump(settings, f, ensure_ascii=False, indent=2)
                
            self.log_message("Настройки сохранены")
            
        except Exception as e:
            self.log_message(f"Ошибка сохранения настроек: {str(e)}", "ERROR")

    def load_settings(self):
        """Загружает настройки из файла"""
        try:
            if not os.path.exists(self.settings_file):
                return
                
            with open(self.settings_file, 'r', encoding='utf-8') as f:
                settings = json.load(f)
            
            # Восстанавливаем пути к файлам и папкам
            if 'excel_path' in settings:
                self.excel_path.set(settings['excel_path'])
            
            if 'email_account' in settings:
                self.email_account.set(settings['email_account'])
                
            if 'thread_count' in settings:
                self.thread_count.set(settings['thread_count'])
                
            if 'email_subject' in settings:
                self.email_subject.set(settings['email_subject'])
                
            if 'email_body' in settings:
                self.body_text.delete("1.0", END)
                self.body_text.insert("1.0", settings['email_body'])
                
            if 'folder_path_1' in settings:
                self.folder_path_1.set(settings['folder_path_1'])
                
            if 'folder_path_2' in settings:
                self.folder_path_2.set(settings['folder_path_2'])
                
            if 'folder_path_3' in settings:
                self.folder_path_3.set(settings['folder_path_3'])
                
            self.log_message("Настройки загружены")
            
        except Exception as e:
            self.log_message(f"Ошибка загрузки настроек: {str(e)}", "ERROR")


def main():
    root = Tk()
    app = EmailSenderApp(root)
    
    # Обработчик закрытия окна
    def on_closing():
        app.save_settings()  # Сохраняем настройки перед выходом
        if app.executor and app.executor._threads:
            app.is_cancelled = True
            app.executor.shutdown(wait=False)
        root.destroy()
    
    root.protocol("WM_DELETE_WINDOW", on_closing)
    root.mainloop()




if __name__ == "__main__":
    main()