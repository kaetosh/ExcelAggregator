# -*- coding: utf-8 -*-
"""
Created on Sun Nov 30 16:15:26 2025

@author: karab
"""

from textual import on
from textual.app import ComposeResult
from textual.containers import Horizontal
from textual.widgets import Footer, Static, Button, Switch, SelectionList, Markdown
from textual.containers import Container
from textual.screen import ModalScreen
from textual.binding import Binding

from utils import update_config, read_config, DEFAULT_CONFIG, generate_compact_report

class ReportScreen(ModalScreen):
    """
    Окно со списком необработанных файлов
    """
    BINDINGS = [
        Binding(key="escape", action="exit_windows", description="Закрыть", key_display="escape"),
    ]

    def compose(self) -> ComposeResult:
        # Создаем пустой Markdown виджет
        self.markdown = Markdown("")
        self.markdown.code_indent_guides = False
        yield self.markdown
        yield Footer()
    
    def on_show(self) -> None:
        report = generate_compact_report(self.app.missing_files)
        self.markdown.update(report)
        
    def action_exit_windows(self) -> None:
        self.app.pop_screen()
            

class SheetsScreen(ModalScreen):
    """
    Окно с выбором листов из книг Excel для агрегирования
    """
    BINDINGS = [
        Binding(key="backspace", action="deselect", description="Очистить выбор", key_display="backspace"),
        Binding(key="escape", action="exit_windows", description="Закрыть", key_display="escape"),
    ]
    def compose(self) -> ComposeResult:
        yield Container(
        SelectionList(),
        Horizontal(
            Button("Сохранить", variant="success", id="button-sheetsscreen-modal"),
            id="horizontals-button-sheetsscreen-modal"),
        id="container-sheetsscreen-modal"
        )
        yield Footer()
        
    def on_mount(self) -> None:
        self.query_one(Button).disabled = True
        self.query_one(SelectionList).border_title = "Выберите листы:"
        self.query_one(SelectionList).clear_options()
        self.query_one(SelectionList).add_options([(name, name) for name in self.app.sheet_names])
        
        # Восстановление выбора из предыдущей сессии
        if hasattr(self.app, 'sheet_selected_names'):
            # Упрощенная проверка
            if (self.app.sheet_selected_names and 
                self.app.sheet_selected_names != ['НЕ ВЫБРАНЫ']):
                
                # Восстанавливаем выбранные элементы
                for name in self.app.sheet_selected_names:
                    if name in self.app.sheet_names and name != 'НЕ ВЫБРАНЫ':
                        self.query_one(SelectionList).select(name)
    
    def on_button_pressed(self, event: Button.Pressed):
        """Обрабатывает нажатие кнопки "Сохранить"."""
        if event.button.id == "button-sheetsscreen-modal":
            self.app.sheet_selected_names = self.query_one(SelectionList).selected
            self.dismiss()
    
    @on(SelectionList.SelectedChanged)
    def handle_select_sheet(self):
        selected = self.query_one(SelectionList).selected
        self.query_one(Button).disabled = False if selected else True
    
    def action_deselect(self) -> None:
        self.query_one(SelectionList).deselect_all()
    
    def action_exit_windows(self) -> None:
        self.query_one(SelectionList).deselect_all()
        self.app.pop_screen()
        
        

class SettingsScreen(ModalScreen):
    """
    Окно с настройками.
    """
    
    def compose(self) -> ComposeResult:
        config = read_config()
        general_options = config.get("general_settings", DEFAULT_CONFIG.get("general_settings", {}))
        general_header_value = bool(general_options.get("general_header", 0))
        
        yield Container(
            Horizontal(
                Static("Общая шапка:", classes="statics-settings-modal"),
                Switch(value=general_header_value, id='switch-general-header', classes="switchs-settings-modal"),
                id='horizontal-general-header-settings-modal'
                ),
            Horizontal(
                Button("Сохранить", variant="success", id="button-settings-modal"),
                id="horizontals-button-settings-modal"),
            id="container-settings-modal"
        )
    
    def on_mount(self) -> None:
        self.query_one('#horizontal-general-header-settings-modal').tooltip = 'Автоматически объединить данные под общими названиями столбцов'

    
    def on_button_pressed(self, event: Button.Pressed):
        """Обрабатывает нажатие кнопки "Сохранить"."""
        if event.button.id == "button-settings-modal":
            general_header_val = int(self.query_one('#switch-general-header', Switch).value)
            updates = {
                        "general_settings": {
                            "general_header": general_header_val,
                                            }
                      }

            update_config(updates=updates)
            self.app.pop_screen()