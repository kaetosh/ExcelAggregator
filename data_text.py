# -*- coding: utf-8 -*-
"""
Created on Mon May 26 09:17:28 2025

@author: a.karabedyan
"""

NAME_APP = 'ExcelAggregator'
SUB_TITLE_APP = 'простой аналог PowerQuery из Excel'
NAME_DATA_FILE = 'data_comparison.xlsx'
NAME_OUTPUT_FILE = 'consolidated.xlsx'
correct_columns = ['data1', 'data2']

TEXT_INTRODUCTION = f"""{NAME_APP} позволяет быстро собрать данные из множества файлов Excel в одном, указав только папку с исходыми файлами и лист, с которго нужно собрать эти данные.
Данные будут расположены друг под другом 'как есть', без дополнительных обработки.
Это удобно в случае, когда данные из разных файлов имеют одинаковую структуру и/или одинаковую шапку в таблицах.
Следуйте приведенной инструкции!
"""

TEXT_GENERAL = """Чтобы указать имя листа, введите его наименование в поле справа и нажмите Enter.
Чтобы указать папку с исходными файлами Excel, нажмите Ctrl+o или соответствующую кнопку внизу.
Чтобы сформировать сводный файл, нажмите Ctrl+r или соответствующую кнопку внизу.

После обработки сводный файл {NAME_OUTPUT_FILE} будет доступен в папке, из которой запущен {NAME_APP}.

Текущая папка: {file_path}
Текущее имя листа: {sheet_name}
"""

TEXT_ALL_PROCESSED_FILES = """{NAME_APP} обработал все файлы Excel в указанной папке.

"""+TEXT_GENERAL


TEXT_ERR_FILES_EXCEL = """Извините, но {NAME_APP} не нашел в указанной папке файлы Excel.

"""+TEXT_GENERAL


TEXT_APP_EXCEL_NOT_FIND = """Ошибка при загрузке {NAME_OUTPUT_FILE}: Не найдено приложение для работы с Excel-файлами.
Описание ошибки: {error_app_xls}.
Убедитесь, что Excel установлен и нажмите Ctrl+r, чтобы сформировать и открыть для редактирования {NAME_OUTPUT_FILE}

"""+TEXT_GENERAL


TEXT_ERR_NO_PROCESSED_FILES = """Извините, но {NAME_APP} не смог обработать файлы. Убедитесь, что в обрабатываемых файлах существует лист с именем {sheet_name}.

"""+TEXT_GENERAL


TEXT_ERR_NOT_ALL_PROCESSED_FILES = """Обработка завершена частично. {NAME_APP} не смог обработать все файлы. Убедитесь, что в обрабатываемых файлах:
{missing_files}
существует лист с именем {sheet_name}.

"""+TEXT_GENERAL


TEXT_ERR_PERMISSION = """Нет доступа к {NAME_OUTPUT_FILE}. Пожалуйста, закройте данный файл.

"""+TEXT_GENERAL


TEXT_ERR_FILE_NOT_FOUND = """Не найден {NAME_OUTPUT_FILE}. Повторите операцию, нажав Ctrl+r или соответствующую кнопку внизу.

"""+TEXT_GENERAL


TEXT_UNKNOW_ERR = """Извините, произошла неучтенная ошибка: {text_err}.
Перезапустите приложение и попробуйте снова. Если ошибка повторяется, свяжитесь в телеграм @kaetosh"""
