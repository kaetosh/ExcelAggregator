# -*- coding: utf-8 -*-
"""
Created on Mon May 26 10:02:36 2025

@author: a.karabedyan
"""

import os
import numpy as np
import pandas as pd
import win32com.client
import pythoncom
from pathlib import Path
from typing import List, Dict, Set, Callable
from pyexcelerate import Workbook

from utils import read_config, fix_excel_filename
from data_text import (NAME_OUTPUT_FILE,
                       TEXT_CONCAT_PROCESS,
                       TEXT_LOAD_FILE_XLS,
                       TEXT_OPEN_FILE_XLS,
                       TEXT_GENERATING_LIST_SHEETS,
                       TEXT_GENERATING_CONSOLIDATED_FILE)

class NoExcelFilesError(Exception):
    """Custom exception for no Excel files found."""
    pass

class NoSelectSheetsError(Exception):
    """Custom exception for no select sheets."""
    pass

class LargeDataError(Exception):
    """Custom exception for if len(df) > 1_000_000"""

def is_excel_file_open(filepath: str) -> bool:
    """Проверяем, существует ли файл"""
    if not os.path.exists(filepath):
        return False

    # Инициализируем COM в текущем потоке
    pythoncom.CoInitialize()

    try:
        # Получаем объект Excel
        excel = win32com.client.Dispatch("Excel.Application")

        # Проходим по всем открытым книгам
        for workbook in excel.Workbooks:
            if workbook.FullName.lower() == os.path.abspath(filepath).lower():
                raise PermissionError

        return False
    finally:
        # Освобождаем COM
        pythoncom.CoUninitialize()


def get_excel_files(folder_path: Path) -> List[Path]:
    """
    Получить список Excel файлов в указанной папке.
    Исключает временные файлы и файл с именем 'consolidated.xlsx'.
    """
    excel_extensions = ('.xls', '.xlsx', '.xlsm', '.xlsb', '.odf')
    files = [
        f for f in folder_path.iterdir()
        if f.is_file()
        and f.suffix.lower() in excel_extensions
        and not f.name.startswith('~')
        and f.name.lower() != 'consolidated.xlsx'
    ]
    if not files:
        raise NoExcelFilesError("В указанной папке нет файлов Excel.")
    return files


def get_sheet_names(file_path: Path) -> List[str]:
    """
    Получить список имен листов в Excel файле.
    """
    try:
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names
    except KeyError:
        fix_excel_filename(file_path)
        xls = pd.ExcelFile(file_path)
        return xls.sheet_names

def get_unique_sheet_names(file_paths: List[Path],
                           on_status: Callable[[str], None]) -> List[str]:
    """
    Получить отсортированный список уникальных листов из всех файлов.
    Обновляет прогресс бар во время обработки.
    """
    on_status(TEXT_GENERATING_LIST_SHEETS)
    unique_sheets: Set[str] = set()
    for index, file_path in enumerate(file_paths):
        if not file_path.exists() or file_path.suffix.lower() not in ['.xlsx', '.xls', '.xlsm', '.xlsb', '.odf']:
            continue  # Пропускаем несуществующие или неподдерживаемые файлы
        try:
            workbook_sheetnames = get_sheet_names(file_path)
            unique_sheets.update(workbook_sheetnames)
        except Exception:
            # Логирование или обработка ошибок чтения листов можно добавить здесь
            pass
    list_unique_sheets = sorted(unique_sheets, key=str.casefold)
    return list_unique_sheets


def aggregating_data_from_excel_files(excel_files: List[Path],
                                      sheet_name_list: List[str],
                                      on_status: Callable[[str], None]
                                      ) -> Dict[str, List[str]]:
    """
    Агрегирует данные из указанных листов Excel-файлов в один файл.
    Возвращает словарь файлов, которые не удалось обработать, где ключ - имя файла, а значения - список отсутствующих листов этого файла.
    Обновляет виджеты Textual (прогресс бар и notify) во время обработки.
    
    Добавлена поддержка общей шапки на основе config.json:
    - general_header: 0 — без шапки (данные с номерами колонок).
    - general_header: 1 — с шапкой (первая строка файла — заголовки).
    """
    if not excel_files:
        return {}  # Нет файлов для обработки
    on_status(TEXT_GENERATING_CONSOLIDATED_FILE)
    
    # Читаем конфиг для настройки шапки
    config = read_config()
    general_header = config.get("general_settings", {}).get("general_header", 0)
    header_param = 0 if general_header == 1 else None  # 0 для шапки, None для номеров
    
    dict_df: Dict[Path, pd.DataFrame] = {}
    missing_files: Dict[str, List[str]] = {}
    
    number_of_files = len(excel_files)
    checkpoints = [int(number_of_files * i / 10) for i in range(1, 11)]
    for index, file_excel in enumerate(excel_files):
        try:
            # Получаем доступные листы
            lists_current_file = get_sheet_names(file_excel)
            available_sheets = set(lists_current_file)
            # Фильтруем и сортируем листы (игнорируем регистр)
            sheets_to_read = sorted([sheet for sheet in sheet_name_list if sheet in available_sheets], key=str.lower)
            missing_sheets = [sheet for sheet in sheet_name_list if sheet not in available_sheets]
            
            if not sheets_to_read:
                missing_files[file_excel.name] = missing_sheets
                continue
            
            # Читаем листы с учётом настройки шапки
            df_dict = pd.read_excel(file_excel, sheet_name=sheets_to_read, header=header_param)
            
            # Добавляем колонку с именем листа в каждый DataFrame
            for key, df in df_dict.items():
                df.insert(0, 'Имя листа', key)
            
            # Объединяем все листы текущего файла
            df = pd.concat(df_dict.values(), ignore_index=True)
            
            # Добавляем колонку с именем файла
            df.insert(0, 'Имя файла', file_excel.name)
            
            dict_df[file_excel] = df
            
            if (index + 1) in checkpoints:
                percent_complete = ((index + 1) * 100) // number_of_files
                on_status(f"Обработано {percent_complete}% файлов ({index + 1} из {number_of_files})")
            
        except (FileNotFoundError, PermissionError, ValueError):
            # Файл не найден, нет доступа или ошибка чтения Excel (например, повреждённый файл или неверный лист)
            # Предполагаем, что все листы отсутствуют
            missing_files[file_excel.name] = sheet_name_list.copy()
        except Exception:
            # Другие неожиданные ошибки
            # Предполагаем, что все листы отсутствуют
            missing_files[file_excel.name] = sheet_name_list.copy()
  
    # Если есть данные, сохраняем и открываем
    if dict_df:
        try:
            on_status(TEXT_CONCAT_PROCESS)
            result = pd.concat(dict_df.values(), ignore_index=True)
            
            if len(result) > 1_000_000:
                raise LargeDataError("В сводном файле будет более млн. строк., что превышает лимит листа Excel.")
            
            on_status(TEXT_LOAD_FILE_XLS)
            
            # Конвертируем DataFrame в список списков
            result = result.replace(np.nan, None)
            data = [result.columns.tolist()] + result.values.tolist()
            wb = Workbook()
            wb.new_sheet('sheet1', data=data)
            wb.save(NAME_OUTPUT_FILE)
            
            # result.to_excel(NAME_OUTPUT_FILE, index=False)
            
            on_status(TEXT_OPEN_FILE_XLS)
            if os.name == 'nt':
                os.startfile(os.path.abspath(NAME_OUTPUT_FILE))
        except LargeDataError:
            raise LargeDataError("В сводном файле будет более млн. строк., что превышает лимит листа Excel.")
        except PermissionError:
            raise PermissionError(f"Ошибка доступа к файлу {NAME_OUTPUT_FILE}")
        except FileNotFoundError:
            raise FileNotFoundError(f"Файл {NAME_OUTPUT_FILE} не найден.")
        except OSError:
            raise OSError("Ошибка: не найдено приложение для открытия файла.")
        except Exception as e:
            raise Exception(f"Неизвестная ошибка при сохранении: {e}")
    
    return missing_files
