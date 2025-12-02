# -*- coding: utf-8 -*-
"""
Created on Fri Nov 28 15:12:12 2025

@author: karab
"""

import json, os, shutil, tempfile
import tkinter as tk
from tkinter import filedialog
from typing import Dict, Any
from pathlib import Path
from zipfile import ZipFile


# Глобальная переменная для кэширования конфигурации
_config_cache: Dict[str, Any] = {}
CONFIG_FILE_PATH = "config.json"


# Значения по умолчанию для config.json
# Используем значения из config.json, который вы предоставили в первом сообщении
DEFAULT_CONFIG = {
    "general_settings": {"general_header": 0}}

def write_default_config(config_path: str = None):
    """Создает файл config.json со значениями по умолчанию."""
    if config_path is None:
        config_path = CONFIG_FILE_PATH
    
    try:
        with open(config_path, 'w', encoding='utf-8') as file:
            json.dump(DEFAULT_CONFIG, file, ensure_ascii=False, indent=4)
        # print(f"Создан файл конфигурации по умолчанию: {config_path}")
    except Exception as e:
        print(f"Ошибка при создании файла конфигурации по умолчанию: {e}")

def read_config(config_path: str = None) -> dict:
    """
    Читает и возвращает текущую конфигурацию из JSON файла.
    Если файл не найден или некорректен, создает его со значениями по умолчанию.
    """
    if config_path is None:
        config_path = CONFIG_FILE_PATH
    
    # 1. Попытка прочитать файл
    try:
        with open(config_path, 'r', encoding='utf-8') as file:
            config = json.load(file)
            return config
    
    # 2. Обработка FileNotFoundError: файл не найден
    except FileNotFoundError:
        # print(f"Файл конфигурации {config_path} не найден. Создание файла по умолчанию.")
        write_default_config(config_path)
        return DEFAULT_CONFIG
        
    # 3. Обработка json.JSONDecodeError: некорректный формат
    except json.JSONDecodeError:
        # print(f"Ошибка: Неверный формат JSON в файле {config_path}. Файл будет перезаписан значениями по умолчанию.")
        write_default_config(config_path)
        return DEFAULT_CONFIG
        
    # 4. Обработка других ошибок
    except Exception as e:
        print(f"Непредвиденная ошибка при чтении конфигурации: {e}. Возврат значений по умолчанию.")
        return DEFAULT_CONFIG 

def update_config(updates, config_path: str = None):
    
    """
    Вносит изменения в файл конфигурации JSON
    
    Args:
        config_path (str): Путь к файлу config.json
        updates (dict): Словарь с обновлениями для конфигурации
    """
    
    if config_path is None:
        config_path = CONFIG_FILE_PATH
    
    # Сначала читаем конфигурацию, которая теперь гарантированно вернет либо существующую, либо дефолтную
    config = read_config(config_path)
    
    try:
        # Рекурсивное обновление конфигурации
        def deep_update(current_dict, update_dict):
            for key, value in update_dict.items():
                if (key in current_dict and 
                    isinstance(current_dict[key], dict) and 
                    isinstance(value, dict)):
                    deep_update(current_dict[key], value)
                else:
                    current_dict[key] = value
        
        # Применение обновлений
        deep_update(config, updates)
        
        # Запись обновленной конфигурации
        with open(config_path, 'w', encoding='utf-8') as file:
            json.dump(config, file, ensure_ascii=False, indent=4)
        
        # print("Конфигурация успешно обновлена")
        clear_config_cache()
        return True
        
    except Exception as e:
        print(f"Ошибка при обновлении конфигурации: {e}")
        return False

def load_config(file_path: str = CONFIG_FILE_PATH) -> Dict[str, Any]:
    global _config_cache
    if not _config_cache:
        _config_cache = read_config(file_path)
    return _config_cache


# Добавляем функцию для очистки кэша, чтобы можно было перечитать конфиг
def clear_config_cache():
    global _config_cache
    _config_cache = {}

def select_folder(current_path: Path) -> Path:
    """
    Открыть диалог выбора папки и вернуть выбранный путь.
    Если пользователь отменил выбор, вернуть current_path.
    Окно диалога будет на переднем плане.
    """
    root = tk.Tk()
    root.withdraw()
    
    # Сделать окно на переднем плане перед открытием диалога
    root.attributes('-topmost', True)
    root.focus_force()
    
    folder_path = filedialog.askdirectory(title="Выберите папку")
    
    # Снять флаг topmost после выбора (опционально, если хотите вернуть нормальное поведение)
    root.attributes('-topmost', False)
    
    root.destroy()
    return Path(folder_path) if folder_path else current_path

def generate_compact_report(problem_files: dict) -> str:
    """Компактный отчет в виде таблицы"""
    
    report = """
# 🚫 Пропущенные файлы

Файлы не были обработаны из-за отсутствия листов:

| № | Файл | Отсутствующие листы |
|---|---|---|
"""
    
    for idx, (file_name, sheets) in enumerate(problem_files.items(), 1):
        # Обрабатываем слишком длинные имена файлов
        short_name = file_name if len(file_name) < 40 else file_name[:37] + "..."
        
        # Объединяем листы через запятую
        sheets_list = ", ".join(f"`{sheet}`" for sheet in sheets)
        
        report += f"| {idx} | `{short_name}` | {sheets_list} |\n"
    
    report += f"""
---
**Затронуто файлов:** {len(problem_files)}
"""
    
    return report

def fix_excel_filename(excel_file_path: Path) -> None:
    """
    Исправляет проблему с регистром в именах файлов внутри Excel-файлов (.xlsx), которые на самом деле являются ZIP-архивами.
    Иногда 1С некорректно генерируют Excel-файлы, используя SharedStrings.xml вместо правильного sharedStrings.xml (с маленькой буквы "s").
    Это вызывает ошибки при открытии файла в pandas.
    

    Parameters
    ----------
    excel_file_path : Path
        Путь к aqke Excel.

    Returns
    -------
    None
        изменяет файл на месте (in-place).

    """
    with tempfile.TemporaryDirectory() as tmp_dir:
        tmp_folder = Path(tmp_dir)

        with ZipFile(excel_file_path) as excel_container:
            excel_container.extractall(tmp_folder)

        wrong_file_path = tmp_folder / 'xl' / 'SharedStrings.xml'
        correct_file_path = tmp_folder / 'xl' / 'sharedStrings.xml'

        if wrong_file_path.exists():
            os.rename(wrong_file_path, correct_file_path)

        # Создаем архив с новым именем
        tmp_zip_path = excel_file_path.with_suffix('.zip')
        shutil.make_archive(str(excel_file_path.with_suffix('')), 'zip', tmp_folder)

        # Удаляем исходный файл и переименовываем новый архив
        if excel_file_path.exists():
            os.remove(excel_file_path)
        os.rename(tmp_zip_path, excel_file_path)