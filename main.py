# -*- coding: utf-8 -*-
"""
Created on Mon May 26 09:16:02 2025

@author: a.karabedyan
"""

from typing import List, Literal
from pathlib import Path

from textual import work
from textual.app import App, ComposeResult
from textual.reactive import reactive
from textual.containers import Horizontal
from textual.widgets import Header, Footer, LoadingIndicator, Markdown, Button
from textual.binding import Binding

from modal_screen import SheetsScreen, SettingsScreen, ReportScreen

from utils import select_folder

from aggregation import (get_excel_files,
                         aggregating_data_from_excel_files,
                         NoExcelFilesError,
                         NoSelectSheetsError,
                         LargeDataError,
                         get_unique_sheet_names,
                         is_excel_file_open)

from data_text import (NAME_APP,
                       SUB_TITLE_APP,
                       TEXT_INTRODUCTION,
                       NAME_OUTPUT_FILE,
                       TEXT_ERR_NO_SELECT_SHEETS,
                       TEXT_ERR_FILES_EXCEL,
                       TEXT_ERR_PERMISSION,
                       TEXT_ERR_FILE_NOT_FOUND,
                       TEXT_APP_EXCEL_NOT_FIND,
                       TEXT_UNKNOW_ERR,
                       TEXT_ERR_FOLDER_NOT_SELECTED,
                       TEXT_ERR_NO_PROCESSED_FILES,
                       TEXT_ALL_PROCESSED_FILES,
                       TEXT_SHEETS_READY,
                       TEXT_ERR_LARGE_DATA
                       )


OperationStatus = Literal["before", "after"]



class ExcelAggregatorApp(App):
    # CSS_PATH = "stile.tcss"
    
    CSS = """
    ToastRack {
            position: relative;
            offset: 0 -6;
        }

    .introduction {
        height: auto;
        border: solid #0087d7;
    }

    #buttons {
        align: center middle;
        layout: grid;
        grid-size: 3 1;
        height: 5;
    }

    Button {
        border: tall $background;
        width: 100%;
    }

    LoadingIndicator {
        height: auto;
    }


    ReportScreen {
        align: center middle;
        border: solid $accent;
        background: $surface;
        }

    SheetsScreen {
        align: center middle;
        }
        #container-sheetsscreen-modal{
            width: 45; 
            height: 20;
            border: solid $accent;
            background: $surface;
            }
        SelectionList{
            height: 5fr;
            }
        #horizontals-button-settings-modal{
            height: 1fr;

        }
    SettingsScreen {
            align: center middle;
        }
            #container-settings-modal {
               width: 45; 
               height: 15;
               border: solid $accent;
               background: $surface;
               padding: 1;
            }
           .statics-settings-modal {
              content-align: left middle;
              height: 3;
              width: 80%; 
           }
           .switchs-settings-modal {
               content-align: right middle;
               width: 20%;
           }
           #horizontals-button-settings-modal{
               align: center bottom;
               }
           #button-settings-modal {
                width: 45%;
                        }

    """

    BINDINGS = [Binding(key="f3",
                        action="push_screen('settings')",
                        description="Настройки",
                        key_display="F3")]

    file_path: Path = reactive(Path.cwd()) # путь к выбранной папке с файлами для обработки
    sheet_names: List[str] = reactive(['НЕ ВЫБРАНЫ']) # список всех листов фалов из выбранной папки
    sheet_selected_names: List[str] = reactive(['НЕ ВЫБРАНЫ']) # список выбранных листов для обработки
    names_files_excel: List[Path] = reactive(None) # список путей к файлам выбранной папки
    missing_files: dict[str, list[str]] = reactive({}) # словарь пропущенных из-за несуществующих листов файлов (ключ - файл, значение - список листов)

    def compose(self) -> ComposeResult:
        markdown = Markdown(TEXT_INTRODUCTION, classes='introduction')
        markdown.code_indent_guides = False
        yield Header(show_clock=True)
        yield markdown
        yield Horizontal(
            Button("📁 Выбрать папку", id="button_select_folder", variant="primary"),
            Button("📑 Выбрать листы", id="button_select_sheets", variant="primary"),
            Button("📥 Агрегировать", id="button_aggregate", variant="primary"),
            id="buttons")
        yield LoadingIndicator()
        yield Footer(show_command_palette = False)

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(LoadingIndicator).visible = False
        self.install_screen(SettingsScreen(), name="settings")
        
    def action_push_screen(self, screen_name: str) -> None:
        """Действие для открытия экрана по имени."""
        self.push_screen(screen_name)
    
    def updating_interface_status(self, status: OperationStatus) -> None:
        """
    Обновляет визуальное состояние интерфейса в зависимости от этапа операции.
    
    Args:
        status: Статус выполнения операции:
                - 'before': отображает состояние 'в процессе' перед операцией
                - 'after': обновляет интерфейс после завершения операции
    """
        is_loading = status == "before"
        self.query_one(LoadingIndicator).visible = is_loading
        self.query_one(Footer).display = not is_loading
        for button in self.query("Button"):
            button.disabled = is_loading
    
    def update_sheet_names(self, sheet_names):
        self.sheet_names = sheet_names
    
    def on_button_pressed(self, event: Button.Pressed) -> None:
        
        if event.button.id == "button_select_folder":
            self.updating_interface_status('before')
            self.file_path = select_folder(self.file_path) or Path.cwd()
            self.load_files_thread()
            
        elif event.button.id == "button_select_sheets":
            if self.sheet_names != ['НЕ ВЫБРАНЫ']:
                self.push_screen(SheetsScreen())
            else:
                self.notify(TEXT_ERR_FOLDER_NOT_SELECTED,
                            title="Ошибка",
                            severity='error',
                            timeout=5)
            
        elif event.button.id == "button_aggregate":
            if not self.names_files_excel:
                self.notify(TEXT_ERR_FOLDER_NOT_SELECTED,
                            title="Ошибка",
                            severity='error',
                            timeout=5)
            elif self.sheet_selected_names == ['НЕ ВЫБРАНЫ']:
                self.notify(TEXT_ERR_NO_SELECT_SHEETS,
                            title="Ошибка",
                            severity='error',
                            timeout=5)
            else:
                self.updating_interface_status('before')
                self.action_open_consolidate()
    
    def get_error_message(self, error):
        error_messages = {
            NoSelectSheetsError: TEXT_ERR_NO_SELECT_SHEETS,
            NoExcelFilesError: TEXT_ERR_FILES_EXCEL,
            PermissionError: TEXT_ERR_PERMISSION,
            FileNotFoundError: TEXT_ERR_FILE_NOT_FOUND,
            OSError: TEXT_APP_EXCEL_NOT_FIND,
            TypeError: TEXT_ERR_FOLDER_NOT_SELECTED,
            LargeDataError: TEXT_ERR_LARGE_DATA
        }
        return error_messages.get(type(error), TEXT_UNKNOW_ERR.format(text_err=error))
    
    def handle_aggregation_results(self, missing_files):
        if len(missing_files) == len(self.names_files_excel):
            self.notify(TEXT_ERR_NO_PROCESSED_FILES,
                        title="Ошибка",
                        severity='error',
                        timeout=5)
        elif missing_files:
            self.missing_files = missing_files
            self.push_screen(ReportScreen())
        else:
            self.notify(TEXT_ALL_PROCESSED_FILES,
                        title="Статус",
                        severity='info',
                        timeout=5)

    
    @work(thread=True)
    def load_files_thread(self) -> None:
        def status_callback(status: str):
            self.call_from_thread(self.notify,
                                  status,
                                  title="Статус",
                                  severity="info",
                                  timeout=5)
        try:
            self.names_files_excel = get_excel_files(self.file_path)
            sheet_names = get_unique_sheet_names(self.names_files_excel,
                                                 on_status=status_callback)
            self.call_from_thread(self.update_sheet_names, sheet_names)
            self.call_from_thread(self.notify,
                                  TEXT_SHEETS_READY,
                                  title="Статус",
                                  severity="info",
                                  timeout=5)
        except NoExcelFilesError as e:
            message_error = self.get_error_message(e)
            self.call_from_thread(self.notify,
                                  message_error,
                                  title="Ошибка",
                                  severity="error",
                                  timeout=5)
        finally:
            self.call_from_thread(self.updating_interface_status, 'after')
        
    @work(thread=True)
    def action_open_consolidate(self) -> None:
        def status_callback(status: str):
            self.call_from_thread(self.notify,
                                  status,
                                  title="Статус",
                                  severity="info",
                                  timeout=5)
        try:
            is_excel_file_open(NAME_OUTPUT_FILE)
            missing_files = aggregating_data_from_excel_files(self.names_files_excel,
                                                              self.sheet_selected_names,
                                                              on_status=status_callback)
            self.call_from_thread(self.handle_aggregation_results, missing_files)
        except (NoSelectSheetsError, NoExcelFilesError, PermissionError, FileNotFoundError, OSError, TypeError, LargeDataError) as e:
            message_error = self.get_error_message(e)
            self.call_from_thread(self.notify,
                                  message_error,
                                  title="Ошибка",
                                  severity="error",
                                  timeout=5)
        except Exception as e:
            self.call_from_thread(self.notify,
                                  TEXT_UNKNOW_ERR.format(text_err=e),
                                  title="Ошибка",
                                  severity="error",
                                  timeout=5)
        finally:
            self.call_from_thread(self.updating_interface_status, 'after')

if __name__ == "__main__":
    app = ExcelAggregatorApp()
    app.run()
