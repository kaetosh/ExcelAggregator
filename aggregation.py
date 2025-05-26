# -*- coding: utf-8 -*-
"""
Created on Mon May 26 10:02:36 2025

@author: a.karabedyan
"""

import os
import tkinter as tk
from tkinter import filedialog
from pathlib import Path
import pandas as pd
from typing import List
from textual.widgets import ProgressBar

from data_text import NAME_OUTPUT_FILE

def aggregating_data_from_excel_files(pr_bar: ProgressBar, excel_files: List[Path], sheet_name) -> List[str]:
    dict_df = {}
    missing_files = []
    for index, file_excel in enumerate(excel_files):
        try:
            # Читаем файл без заголовков
            df = pd.read_excel(file_excel, sheet_name, header=None)

            # Добавляем столбец с именем файла
            df['Исх.файл'] = file_excel.name

            # Перемещаем столбец 'Исх.файл' в начало
            cols = df.columns.to_list()
            cols.insert(0, cols.pop(cols.index('Исх.файл')))
            df = df.loc[:, cols]

            # Сохраняем DataFrame в словаре
            dict_df[file_excel] = df
        except ValueError:
            missing_files.append(file_excel.name)
        percentage = ((index+1) / len(excel_files)) * 100
        pr_bar.update(progress=percentage)
    pr_bar.update(progress=100)
    try:
        if dict_df:
            result = pd.concat(list(dict_df.values()), ignore_index=True)
            result.to_excel(NAME_OUTPUT_FILE, index = False)
            os.startfile(os.path.abspath(NAME_OUTPUT_FILE))
    except PermissionError:
        raise PermissionError(f"Ошибка доступа к {NAME_OUTPUT_FILE}")
    except FileNotFoundError:
        raise FileNotFoundError(f"Ошибка, файл {NAME_OUTPUT_FILE} не найден.")
    except OSError:
        raise OSError("Ошибка, не найдено приложение Excel.")
    except Exception:
        raise Exception("Неизвестная ошибка!")
    return missing_files

class NoExcelFilesError(Exception):
    """Custom exception for no Excel files found."""
    pass

def get_excel_files(folder_path: Path) -> List[Path]:
    # Определяем расширения файлов Excel
    excel_extensions = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.odf')

    # Получаем список файлов в указанной папке
    files = [f for f in folder_path.iterdir() if f.suffix in excel_extensions and not f.name.startswith('~') and f.name != 'consolidated.xlsx']

    # Проверяем, есть ли файлы Excel
    if not files:
        raise NoExcelFilesError("В указанной папке нет файлов Excel.")
    return files

def select_folder(current_path) -> Path:
    # Создаем скрытое окно
    root = tk.Tk()
    root.withdraw()  # Скрываем главное окно

    # Открываем диалог выбора папки
    folder_path = filedialog.askdirectory(title="Выберите папку")

    return Path(folder_path) if folder_path else current_path
