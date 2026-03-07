#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Почтовая рассылка расчетных листов ФитоФарм
Главный файл запуска приложения
"""

import tkinter as tk
import sys
from pathlib import Path

# Добавляем корень проекта в path
project_root = Path(__file__).parent
if str(project_root) not in sys.path:
    sys.path.insert(0, str(project_root))

from frontend.app import EmailSenderApp


def main():
    """Точка входа в приложение"""
    root = tk.Tk()
    app = EmailSenderApp(root)
    
    # Обработчик закрытия
    root.protocol("WM_DELETE_WINDOW", app.on_closing)
    
    root.mainloop()


if __name__ == "__main__":
    main()
