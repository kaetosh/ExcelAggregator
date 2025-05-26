# -*- coding: utf-8 -*-
"""
Created on Mon May 26 09:16:02 2025

@author: a.karabedyan
"""

from typing import List
from pathlib import Path

from textual import work, on
from textual.app import App, ComposeResult
from textual.reactive import reactive
from textual.containers import Horizontal
from textual.widgets import Header, Footer, Static, LoadingIndicator, ProgressBar, Input
from textual.containers import Middle, Center

from aggregation import (select_folder,
                         get_excel_files,
                         aggregating_data_from_excel_files,
                         NoExcelFilesError)

from data_text import (NAME_APP,
                       NAME_OUTPUT_FILE,
                       SUB_TITLE_APP,
                       TEXT_INTRODUCTION,
                       TEXT_GENERAL,
                       TEXT_ERR_FILES_EXCEL,
                       TEXT_ERR_NO_PROCESSED_FILES,
                       TEXT_ERR_NOT_ALL_PROCESSED_FILES,
                       TEXT_ERR_PERMISSION,
                       TEXT_ERR_FILE_NOT_FOUND,
                       TEXT_APP_EXCEL_NOT_FIND,
                       TEXT_UNKNOW_ERR,
                       TEXT_ALL_PROCESSED_FILES
                       )

class ExcelAggregatorApp(App):
    CSS = """
    .introduction {
        height: auto;
        border: solid #0087d7;
    }
    .steps_l {
        height: auto;
        width: 9fr;
    }
    .steps_r {
        height: 100%;
        border: solid #0087d7;
        width: 3fr;
        margin: 0 0 0 1
    }
    Horizontal {
        height: auto;
        border: solid #0087d7;
    }
    LoadingIndicator {
        dock: bottom;
        height: 10%;
    }
    ProgressBar {
        padding-left: 3;
    }
    """

    BINDINGS = [
        ("ctrl+o", "open_dir", "Выбрать папку с исходными файлами"),
        ("ctrl+r", "open_consolidate", "Сформировать и открыть сводный файл"),
    ]
    file_path = reactive(Path.cwd())
    sheet_name = reactive('Лист1')


    def compose(self) -> ComposeResult:
        yield Header(show_clock=True, icon='<>')
        yield Static(TEXT_INTRODUCTION, classes='introduction')
        yield Horizontal(
                Static(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                           NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                           file_path=self.file_path,
                                           sheet_name=self.sheet_name), classes='steps_l'),
                Input(placeholder="Введи и нажми Enter",
                      type="text",
                      max_length = 31,
                      classes='steps_r'), id='example')
        yield Footer()
        yield LoadingIndicator()
        with Middle():
            with Center():
                yield ProgressBar(show_eta=False)

    def on_mount(self) -> None:
        self.title = NAME_APP
        self.sub_title = SUB_TITLE_APP
        self.query_one(LoadingIndicator).visible = False
        self.query_one(ProgressBar).visible = False


    @on(Input.Submitted)
    def save_name_sheet(self, event: Input.Submitted) -> None:
        self.sheet_name = self.query_one(Input).value
        self.query_one('.steps_l').update(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                                              NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                              file_path=self.file_path,
                                                              sheet_name=self.sheet_name))

    def action_open_dir(self) -> None:
        self.file_path: Path = select_folder(self.file_path)
        self.query_one('.steps_l').update(TEXT_GENERAL.format(NAME_APP=NAME_APP,
                                                              NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                              file_path=self.file_path,
                                                              sheet_name=self.sheet_name))

    @work(thread=True)
    def action_open_consolidate(self) -> None:
        self.query_one(LoadingIndicator).visible = True
        self.query_one('.steps_l').update('Идет обработка данных...')
        self.query_one('.steps_r').visible=False
        try:
            self.names_files_excel: List[Path] =  get_excel_files(self.file_path)
            self.sheet_name = self.query_one(Input).value
            self.query_one(ProgressBar).update(total=100)
            self.query_one(ProgressBar).visible = True
            missing_files= aggregating_data_from_excel_files(self.query_one(ProgressBar), self.names_files_excel, self.sheet_name)
            if len(missing_files) == len(self.names_files_excel):
                self.query_one('.steps_l').update(TEXT_ERR_NO_PROCESSED_FILES.format(NAME_APP=NAME_APP,
                                                                                     NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                                     file_path=self.file_path,
                                                                                     sheet_name=self.sheet_name))
            elif missing_files:
                self.query_one('.steps_l').update(TEXT_ERR_NOT_ALL_PROCESSED_FILES.format(NAME_APP=NAME_APP,
                                                                                          NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                                          missing_files='\n'.join(missing_files),
                                                                                          file_path=self.file_path,
                                                                                          sheet_name=self.sheet_name))
            else:
                self.query_one('.steps_l').update(TEXT_ALL_PROCESSED_FILES.format(NAME_APP=NAME_APP,
                                                                                  NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                                  file_path=self.file_path,
                                                                                  sheet_name=self.sheet_name))
        except NoExcelFilesError:
            self.query_one('.steps_l').update(TEXT_ERR_FILES_EXCEL.format(NAME_APP=NAME_APP,
                                                                          NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                          file_path=self.file_path,
                                                                          sheet_name=self.sheet_name))
        except PermissionError:
            self.query_one('.steps_l').update(TEXT_ERR_PERMISSION.format(NAME_APP=NAME_APP,
                                                                         NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                         file_path=self.file_path,
                                                                         sheet_name=self.sheet_name))
        except FileNotFoundError:
            self.query_one('.steps_l').update(TEXT_ERR_FILE_NOT_FOUND.format(NAME_APP=NAME_APP,
                                                                             NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                             file_path=self.file_path,
                                                                             sheet_name=self.sheet_name))
        except OSError as e:
            self.query_one('.steps_l').update(TEXT_APP_EXCEL_NOT_FIND.format(NAME_APP=NAME_APP,
                                                                             NAME_OUTPUT_FILE=NAME_OUTPUT_FILE,
                                                                             error_app_xls = e,
                                                                             file_path=self.file_path,
                                                                             sheet_name=self.sheet_name))
        except Exception as e:
            self.query_one('.steps_l').update(TEXT_UNKNOW_ERR.format(text_err=e))

        finally:
            self.query_one(LoadingIndicator).visible = False
            self.query_one('.steps_r').visible=True
            self.query_one(ProgressBar).visible = False
            self.query_one(ProgressBar).update(progress=0)


if __name__ == "__main__":
    app = ExcelAggregatorApp()
    app.run()
