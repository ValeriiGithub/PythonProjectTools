"""
программа на python, которая из таблицы excel создает таблицу в markdown формате
v.0.2.0
- вызов диалогового окна для выбора файла и места сохранения файла с помощью библиотеки tkinter.
Пожалуйста, убедитесь, что у вас установлены необходимые библиотеки (pandas, tabulate и tkinter)
перед запуском этого кода.
"""

import pandas as pd
import tkinter as tk
from tkinter import filedialog


def excel_to_markdown():
    # Создаем окно для выбора файла
    root = tk.Tk()
    root.withdraw()

    # Запрашиваем у пользователя файл Excel
    file_path = filedialog.askopenfilename(title="Выберите файл Excel",
                                           filetypes=(("Excel Files", "*.xlsx"), ("All Files", "*.*")))

    # Читаем файл Excel
    df = pd.read_excel(file_path)

    # Конвертируем DataFrame в Markdown
    markdown = df.to_markdown()

    # Запрашиваем у пользователя место для сохранения файла
    save_path = filedialog.asksaveasfilename(defaultextension=".md",
                                             filetypes=(("Markdown Files", "*.md"), ("All Files", "*.*")))

    # Сохраняем Markdown в файл
    with open(save_path, 'w') as f:
        f.write(markdown)


# Пример использования:
excel_to_markdown()
