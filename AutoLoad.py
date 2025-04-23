import sys
import os
import subprocess
import winreg
import uuid
import json
import logging
from logging.handlers import RotatingFileHandler
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QLineEdit, QPushButton, QCheckBox, QFileDialog, QMessageBox,
    QTabWidget, QTableWidget, QTableWidgetItem, QHeaderView, QTextEdit, QDialogButtonBox
)
from PySide6.QtCore import Qt
from PySide6.QtGui import QDragEnterEvent, QDragMoveEvent, QDropEvent
import re

logger = logging.getLogger('AutostartManager')
logger.setLevel(logging.INFO)
handler = RotatingFileHandler('autostart_manager.log', maxBytes=1_000_000, backupCount=5)
formatter = logging.Formatter('%(asctime)s - %(levelname)s - %(message)s')
handler.setFormatter(formatter)
logger.addHandler(handler)

class DragDropLineEdit(QLineEdit):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setToolTip("Drag and drop an executable file (.exe, .bat, .cmd) here")

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dragMoveEvent(self, event: QDragMoveEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            file_path = event.mimeData().urls()[0].toLocalFile()
            if self.is_valid_file(file_path):
                self.setText(file_path)
                parent = self.parent()
                while parent and not hasattr(parent, 'app_name_input'):
                    parent = parent.parent()
                if parent:
                    app_name = os.path.splitext(os.path.basename(file_path))[0]
                    parent.app_name_input.setText(app_name)
                    logger.info(f"Dropped file for autostart: {file_path}")
            else:
                logger.warning(f"Invalid file dropped for autostart: {file_path}")
                QMessageBox.warning(self, "Warning", f"Invalid file: {file_path}\nOnly .exe, .bat, .cmd files are allowed.")

    def is_valid_file(self, file_path):
        return (os.path.exists(file_path) and 
                file_path.lower().endswith(('.exe', '.bat', '.cmd')) and
                not any(c in file_path for c in '<>|&'))

class DragDropTableWidget(QTableWidget):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setAcceptDrops(True)
        self.setSelectionMode(QTableWidget.SingleSelection)
        self.setEditTriggers(QTableWidget.NoEditTriggers)
        self.setToolTip("Drag and drop executable files here to add to batch execution or reorder by dragging rows")
        self.setSelectionBehavior(QTableWidget.SelectRows)
        self.setDragEnabled(True)
        self.setDropIndicatorShown(True)

    def dragEnterEvent(self, event: QDragEnterEvent):
        if event.mimeData().hasUrls() or event.source() == self:
            event.acceptProposedAction()

    def dragMoveEvent(self, event: QDragMoveEvent):
        if event.mimeData().hasUrls() or event.source() == self:
            event.acceptProposedAction()

    def dropEvent(self, event: QDropEvent):
        if event.source() == self:
            selected_rows = self.selectedIndexes()
            if selected_rows:
                from_row = selected_rows[0].row()
                to_row = self.rowAt(event.position().y())
                if to_row == -1:
                    to_row = self.rowCount()
                if from_row == to_row or from_row + 1 == to_row:
                    return
                items = [self.takeItem(from_row, col) for col in range(self.columnCount())]
                self.removeRow(from_row)
                if from_row < to_row:
                    to_row -= 1
                self.insertRow(to_row)
                for col, item in enumerate(items):
                    self.setItem(to_row, col, item)
                self.selectRow(to_row)
                parent = self.parent()
                while parent and not hasattr(parent, 'save_batch_list'):
                    parent = parent.parent()
                if parent:
                    parent.save_batch_list()
                event.accept()
        else:
            if event.mimeData().hasUrls():
                event.acceptProposedAction()
                for url in event.mimeData().urls():
                    file_path = url.toLocalFile()
                    if DragDropLineEdit().is_valid_file(file_path):
                        for row in range(self.rowCount()):
                            if self.item(row, 0).text() == file_path:
                                logger.warning(f"Duplicate file dropped for batch: {file_path}")
                                return
                        row = self.rowCount()
                        self.insertRow(row)
                        self.setItem(row, 0, QTableWidgetItem(file_path))
                        self.setItem(row, 1, QTableWidgetItem("Stopped"))
                        logger.info(f"Dropped file for batch: {file_path}")
                        parent = self.parent()
                        while parent and not hasattr(parent, 'save_batch_list'):
                            parent = parent.parent()
                        if parent:
                            parent.save_batch_list()
                    else:
                        logger.warning(f"Invalid file dropped for batch: {file_path}")
                        QMessageBox.warning(self, "Warning", f"Invalid file: {file_path}\nOnly .exe, .bat, .cmd files are allowed.")

class AutostartManager(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Autostart Manager")
        self.setFixedSize(700, 550)
        self.processes = {}
        self.batch_list_file = "batch_list.json"
        self.config_file = "config.json"
        self.current_theme = self.load_theme()
        logger.info("Application started")
        
        central_widget = QWidget()
        self.setCentralWidget(central_widget)
        main_layout = QVBoxLayout(central_widget)
        main_layout.setContentsMargins(20, 20, 20, 20)
        main_layout.setSpacing(10)
        
        theme_layout = QHBoxLayout()
        theme_button = QPushButton("Toggle Theme")
        theme_button.setToolTip("Switch between light and dark themes")
        theme_button.clicked.connect(self.toggle_theme)
        theme_layout.addStretch()
        theme_layout.addWidget(theme_button)
        main_layout.addLayout(theme_layout)
        
        tabs = QTabWidget()
        main_layout.addWidget(tabs)
        
        add_tab = QWidget()
        list_tab = QWidget()
        batch_tab = QWidget()
        log_tab = QWidget()
        tabs.addTab(add_tab, "Add to Autostart")
        tabs.addTab(list_tab, "Current Autostart")
        tabs.addTab(batch_tab, "Batch Execution")
        tabs.addTab(log_tab, "Logs")
        
        add_layout = QVBoxLayout(add_tab)
        add_layout.setSpacing(10)
        
        add_layout.addWidget(QLabel("Select file for autostart (or drag and drop):"))
        file_layout = QHBoxLayout()
        self.file_path_input = DragDropLineEdit()
        self.file_path_input.setPlaceholderText("File path...")
        file_layout.addWidget(self.file_path_input)
        browse_button = QPushButton("Browse...")
        browse_button.setToolTip("Browse for an executable file")
        browse_button.clicked.connect(self.browse_file)
        file_layout.addWidget(browse_button)
        add_layout.addLayout(file_layout)
        
        self.admin_checkbox = QCheckBox("Run as administrator")
        self.admin_checkbox.setToolTip("Run the application with elevated privileges")
        add_layout.addWidget(self.admin_checkbox)
        
        add_layout.addWidget(QLabel("Name in autostart:"))
        self.app_name_input = QLineEdit()
        self.app_name_input.setPlaceholderText("Enter application name...")
        self.app_name_input.setToolTip("Enter a unique name for the autostart entry")
        add_layout.addWidget(self.app_name_input)
        
        add_button = QPushButton("Add to autostart")
        add_button.setToolTip("Add the selected application to autostart")
        add_button.clicked.connect(self.add_to_autostart)
        add_layout.addWidget(add_button)
        
        config_layout = QHBoxLayout()
        export_button = QPushButton("Export Settings")
        export_button.setToolTip("Export autostart and batch settings to a file")
        export_button.clicked.connect(self.export_settings)
        config_layout.addWidget(export_button)
        import_button = QPushButton("Import Settings")
        import_button.setToolTip("Import settings from a file")
        import_button.clicked.connect(self.import_settings)
        config_layout.addWidget(import_button)
        add_layout.addLayout(config_layout)
        
        add_layout.addStretch()
        
        self_program_layout = QHBoxLayout()
        self_program_layout.addStretch()
        self_program_button = QPushButton("This Manager")
        self_program_button.setObjectName("selfProgramButton")
        self_program_button.setToolTip("Add or remove this application from autostart")
        self_program_button.clicked.connect(self.toggle_self_autostart)
        self_program_layout.addWidget(self_program_button)
        self_program_layout.addStretch()
        add_layout.addLayout(self_program_layout)
        
        list_layout = QVBoxLayout(list_tab)
        list_layout.setSpacing(10)
        
        self.process_table = QTableWidget()
        self.process_table.setColumnCount(3)
        self.process_table.setHorizontalHeaderLabels(["Name", "Path", "Action"])
        self.process_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.process_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Stretch)
        self.process_table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Fixed)
        self.process_table.setColumnWidth(2, 100)
        self.process_table.setSelectionMode(QTableWidget.NoSelection)
        self.process_table.setEditTriggers(QTableWidget.NoEditTriggers)
        self.process_table.setSortingEnabled(True)
        list_layout.addWidget(self.process_table)
        
        refresh_button = QPushButton("Refresh List")
        refresh_button.setToolTip("Reload the list of autostart applications")
        refresh_button.clicked.connect(self.load_autostart_processes)
        list_layout.addWidget(refresh_button)
        
        batch_layout = QVBoxLayout(batch_tab)
        batch_layout.setSpacing(10)
        
        batch_layout.addWidget(QLabel("Batch Execution Files (Drag and Drop files here):"))
        self.batch_table = DragDropTableWidget()
        self.batch_table.setColumnCount(2)
        self.batch_table.setHorizontalHeaderLabels(["File Path", "Status"])
        self.batch_table.horizontalHeader().setSectionResizeMode(0, QHeaderView.Stretch)
        self.batch_table.horizontalHeader().setSectionResizeMode(1, QHeaderView.Fixed)
        self.batch_table.setColumnWidth(1, 100)
        self.batch_table.setSelectionMode(QTableWidget.SingleSelection)
        self.batch_table.setEditTriggers(QTableWidget.NoEditTriggers)
        batch_layout.addWidget(self.batch_table)
        
        batch_buttons_layout = QHBoxLayout()
        add_batch_button = QPushButton("Add File")
        add_batch_button.setToolTip("Browse for a file to add to batch execution")
        add_batch_button.clicked.connect(self.add_batch_file)
        batch_buttons_layout.addWidget(add_batch_button)
        
        remove_batch_button = QPushButton("Remove Selected")
        remove_batch_button.setToolTip("Remove the selected file from batch execution")
        remove_batch_button.clicked.connect(self.remove_batch_file)
        batch_buttons_layout.addWidget(remove_batch_button)
        
        remove_all_button = QPushButton("Remove All")
        remove_all_button.setToolTip("Remove all files from batch execution")
        remove_all_button.clicked.connect(self.remove_all_batch_files)
        batch_buttons_layout.addWidget(remove_all_button)
        
        batch_layout.addLayout(batch_buttons_layout)
        
        control_buttons_layout = QHBoxLayout()
        start_all_button = QPushButton("Start All")
        start_all_button.setToolTip("Start all batch execution files")
        start_all_button.clicked.connect(self.start_all_batch)
        control_buttons_layout.addWidget(start_all_button)
        
        stop_all_button = QPushButton("Stop All")
        stop_all_button.setToolTip("Stop all running batch execution files")
        stop_all_button.clicked.connect(self.stop_all_batch)
        control_buttons_layout.addWidget(stop_all_button)
        batch_layout.addLayout(control_buttons_layout)
        
        batch_layout.addStretch()
        
        log_layout = QVBoxLayout(log_tab)
        log_layout.setSpacing(10)
        
        self.log_text = QTextEdit()
        self.log_text.setReadOnly(True)
        log_layout.addWidget(self.log_text)
        
        clear_log_button = QPushButton("Clear Logs")
        clear_log_button.setToolTip("Clear the displayed log entries")
        clear_log_button.clicked.connect(self.clear_logs)
        log_layout.addWidget(clear_log_button)
        
        self.load_autostart_processes()
        self.load_batch_list()
        self.load_logs()
        self.apply_theme()
    
    def get_light_theme(self):
        return """
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 14px;
                background-color: #f5f7fa;
            }
            QTabWidget::pane {
                border: 1px solid #dcdcdc;
                border-radius: 8px;
                background-color: #ffffff;
            }
            QTabBar::tab {
                background-color: #e0e6ed;
                color: #333333;
                padding: 10px 20px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #ffffff;
                color: #007BFF;
                border-bottom: 2px solid #007BFF;
            }
            QLineEdit, QTableWidget, QTextEdit {
                padding: 8px;
                border: 1px solid #dcdcdc;
                border-radius: 5px;
                background-color: #ffffff;
            }
            QLineEdit:focus, QTableWidget:focus, QTextEdit:focus {
                border: 2px solid #007BFF;
                background-color: #f0f8ff;
            }
            QTableWidget::item:selected {
                background-color: #007BFF;
                color: white;
            }
            QPushButton {
                padding: 10px;
                border-radius: 5px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #007BFF, stop:1 #0056b3);
                color: white;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #0056b3, stop:1 #003d80);
            }
            QPushButton:pressed {
                background: #003d80;
            }
            QPushButton#selfProgramButton {
                padding: 6px 12px;
                border-radius: 5px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #dc3545, stop:1 #b02a37);
                color: white;
                border: none;
                font-size: 12px;
                font-weight: bold;
                min-width: 100px;
                min-height: 24px;
            }
            QPushButton#selfProgramButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #b02a37, stop:1 #8b1e29);
            }
            QPushButton#selfProgramButton:pressed {
                background: #8b1e29;
            }
            QPushButton#removeButton {
                padding: 6px 16px;
                min-width: 50px;
                min-height: 10px;
                border-radius: 3px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #dc3545, stop:1 #b02a37);
                color: white;
                border: none;
                font-size: 12px;
            }
            QPushButton#removeButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #b02a37, stop:1 #8b1e29);
            }
            QCheckBox {
                padding: 5px;
                color: #333333;
            }
            QTableWidget {
                border: 1px solid #dcdcdc;
                border-radius: 5px;
                background-color: #ffffff;
            }
            QTableWidget::item {
                padding: 5px;
                color: #333333;
            }
            QHeaderView::section {
                background-color: #e0e6ed;
                padding: 5px;
                border: none;
                color: #333333;
                font-weight: bold;
            }
        """
    
    def get_dark_theme(self):
        return """
            QWidget {
                font-family: 'Segoe UI', Arial, sans-serif;
                font-size: 14px;
                background-color: #2c2f33;
                color: #ffffff;
            }
            QTabWidget::pane {
                border: 1px solid #555555;
                border-radius: 8px;
                background-color: #23272a;
            }
            QTabBar::tab {
                background-color: #3a3f44;
                color: #ffffff;
                padding: 10px 20px;
                border-top-left-radius: 8px;
                border-top-right-radius: 8px;
                margin-right: 2px;
            }
            QTabBar::tab:selected {
                background-color: #23272a;
                color: #1e90ff;
                border-bottom: 2px solid #1e90ff;
            }
            QLineEdit, QTableWidget, QTextEdit {
                padding: 8px;
                border: 1px solid #555555;
                border-radius: 5px;
                background-color: #3a3f44;
                color: #ffffff;
            }
            QLineEdit:focus, QTableWidget:focus, QTextEdit:focus {
                border: 2px solid #1e90ff;
                background-color: #2c2f33;
            }
            QTableWidget::item:selected {
                background-color: #1e90ff;
                color: white;
            }
            QPushButton {
                padding: 10px;
                border-radius: 5px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #1e90ff, stop:1 #1565c0);
                color: white;
                border: none;
                font-weight: bold;
            }
            QPushButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #1565c0, stop:1 #0d47a1);
            }
            QPushButton:pressed {
                background: #0d47a1;
            }
            QPushButton#selfProgramButton {
                padding: 6px 12px;
                border-radius: 5px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #dc3545, stop:1 #b02a37);
                color: white;
                border: none;
                font-size: 12px;
                font-weight: bold;
                min-width: 100px;
                min-height: 24px;
            }
            QPushButton#selfProgramButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #b02a37, stop:1 #8b1e29);
            }
            QPushButton#selfProgramButton:pressed {
                background: #8b1e29;
            }
            QPushButton#removeButton {
                padding: 6px 16px;
                min-width: 50px;
                min-height: 10px;
                border-radius: 3px;
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #dc3545, stop:1 #b02a37);
                color: white;
                border: none;
                font-size: 12px;
            }
            QPushButton#removeButton:hover {
                background: qlineargradient(x1:0, y1:0, x2:0, y2:1,
                                            stop:0 #b02a37, stop:1 #8b1e29);
            }
            QCheckBox {
                padding: 5px;
                color: #ffffff;
            }
            QTableWidget {
                border: 1px solid #555555;
                border-radius: 5px;
                background-color: #3a3f44;
                color: #ffffff;
            }
            QTableWidget::item {
                padding: 5px;
                color: #ffffff;
            }
            QHeaderView::section {
                background-color: #3a3f44;
                padding: 5px;
                border: none;
                color: #ffffff;
                font-weight: bold;
            }
        """
    
    def apply_theme(self):
        if self.current_theme == "dark":
            self.setStyleSheet(self.get_dark_theme())
        else:
            self.setStyleSheet(self.get_light_theme())
        logger.info(f"Applied {self.current_theme} theme")
    
    def toggle_theme(self):
        old_theme = self.current_theme
        self.current_theme = "dark" if self.current_theme == "light" else "light"
        self.apply_theme()
        self.save_theme()
        logger.info(f"Toggled theme from {old_theme} to {self.current_theme}")
    
    def save_theme(self):
        try:
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump({"theme": self.current_theme}, f)
            logger.info(f"Saved theme preference: {self.current_theme}")
        except Exception as e:
            logger.error(f"Failed to save theme preference: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to save theme preference: {str(e)}")
    
    def load_theme(self):
        try:
            if os.path.exists(self.config_file):
                with open(self.config_file, 'r', encoding='utf-8') as f:
                    config = json.load(f)
                    theme = config.get("theme", "light")
                    logger.info(f"Loaded theme preference: {theme}")
                    return theme
        except Exception as e:
            logger.error(f"Failed to load theme preference: {str(e)}")
        logger.info("Defaulting to light theme")
        return "light"
    
    def browse_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "",
            "Executable files (*.exe *.bat *.cmd);;All files (*.*)"
        )
        if file_path:
            self.file_path_input.setText(file_path)
            app_name = os.path.splitext(os.path.basename(file_path))[0]
            self.app_name_input.setText(app_name)
            logger.info(f"Selected file for autostart via browse: {file_path}")
    
    def validate_app_name(self, app_name):
        if not app_name:
            return False, "Application name cannot be empty"
        if not re.match(r'^[a-zA-Z0-9_-]+$', app_name):
            return False, "Application name can only contain letters, numbers, underscores, or hyphens"
        return True, ""
    
    def add_to_autostart(self):
        file_path = self.file_path_input.text()
        app_name = self.app_name_input.text()
        run_as_admin = self.admin_checkbox.isChecked()
        
        if not file_path:
            logger.error("Attempted to add to autostart with missing file path")
            QMessageBox.critical(self, "Error", "Specify file path!")
            return
        
        is_valid, error_msg = self.validate_app_name(app_name)
        if not is_valid:
            logger.error(f"Invalid application name: {app_name} - {error_msg}")
            QMessageBox.critical(self, "Error", error_msg)
            return
        
        if not os.path.exists(file_path):
            logger.error(f"Attempted to add non-existent file to autostart: {file_path}")
            QMessageBox.critical(self, "Error", "The specified file does not exist!")
            return
        
        try:
            if run_as_admin:
                self.add_to_task_scheduler(file_path, app_name)
                logger.info(f"Added to autostart (Task Scheduler): {app_name} - {file_path}")
            else:
                self.add_to_registry(file_path, app_name)
                logger.info(f"Added to autostart (Registry): {app_name} - {file_path}")
            QMessageBox.information(self, "Success", f"Application '{app_name}' added to autostart!")
            self.load_autostart_processes()
        except Exception as e:
            logger.error(f"Failed to add to autostart: {app_name} - {file_path} - {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to add to autostart: {str(e)}")
    
    def toggle_self_autostart(self):
        app_name = "AutostartManager"
        current_exe = sys.executable if getattr(sys, 'frozen', False) else __file__
        
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_READ
            )
            try:
                winreg.QueryValueEx(key, app_name)
                winreg.CloseKey(key)
                reply = QMessageBox.question(
                    self, "Confirm", "Remove Autostart Manager from autostart?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    key = winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER,
                        r"Software\Microsoft\Windows\CurrentVersion\Run",
                        0,
                        winreg.KEY_SET_VALUE
                    )
                    winreg.DeleteValue(key, app_name)
                    winreg.CloseKey(key)
                    logger.info("Removed Autostart Manager from autostart")
                    QMessageBox.information(self, "Success", "Autostart Manager removed from autostart!")
            except FileNotFoundError:
                winreg.CloseKey(key)
                reply = QMessageBox.question(
                    self, "Confirm", "Add Autostart Manager to autostart?",
                    QMessageBox.Yes | QMessageBox.No, QMessageBox.No
                )
                if reply == QMessageBox.Yes:
                    key = winreg.OpenKey(
                        winreg.HKEY_CURRENT_USER,
                        r"Software\Microsoft\Windows\CurrentVersion\Run",
                        0,
                        winreg.KEY_SET_VALUE
                    )
                    winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, f'"{current_exe}"')
                    winreg.CloseKey(key)
                    logger.info("Added Autostart Manager to autostart")
                    QMessageBox.information(self, "Success", "Autostart Manager added to autostart!")
            self.load_autostart_processes()
        except Exception as e:
            logger.error(f"Failed to modify Autostart Manager autostart: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to modify autostart: {str(e)}")
    
    def add_to_registry(self, file_path, app_name):
        key = winreg.OpenKey(
            winreg.HKEY_CURRENT_USER,
            r"Software\Microsoft\Windows\CurrentVersion\Run",
            0,
            winreg.KEY_SET_VALUE
        )
        command = f'"{file_path}"'
        winreg.SetValueEx(key, app_name, 0, winreg.REG_SZ, command)
        winreg.CloseKey(key)
    
    def add_to_task_scheduler(self, file_path, app_name):
        task_name = f"Autostart_{app_name}_{uuid.uuid4().hex[:8]}"
        task_xml = f"""<?xml version="1.0" encoding="UTF-16"?>
<Task version="1.2" xmlns="http://schemas.microsoft.com/windows/2004/02/mit/task">
  <RegistrationInfo>
    <Description>Autostart task for {app_name}</Description>
  </RegistrationInfo>
  <Triggers>
    <LogonTrigger>
      <Enabled>true</Enabled>
    </LogonTrigger>
  </Triggers>
  <Principals>
    <Principal id="Author">
      <RunLevel>HighestAvailable</RunLevel>
      <UserId>{os.getlogin()}</UserId>
    </Principal>
  </Principals>
  <Settings>
    <MultipleInstancesPolicy>IgnoreNew</MultipleInstancesPolicy>
    <DisallowStartIfOnBatteries>false</DisallowStartIfOnBatteries>
    <StopIfGoingOnBatteries>false</StopIfGoingOnBatteries>
    <AllowHardTerminate>true</AllowHardTerminate>
    <StartWhenAvailable>false</StartWhenAvailable>
    <RunOnlyIfNetworkAvailable>false</RunOnlyIfNetworkAvailable>
    <IdleSettings>
      <StopOnIdleEnd>true</StopOnIdleEnd>
      <RestartOnIdle>false</RestartOnIdle>
    </IdleSettings>
    <AllowStartOnDemand>true</AllowStartOnDemand>
    <Enabled>true</Enabled>
    <Hidden>false</Hidden>
    <RunOnlyIfIdle>false</RunOnlyIfIdle>
    <WakeToRun>false</WakeToRun>
    <ExecutionTimeLimit>PT72H</ExecutionTimeLimit>
    <Priority>7</Priority>
  </Settings>
  <Actions Context="Author">
    <Exec>
      <Command>"{file_path}"</Command>
    </Exec>
  </Actions>
</Task>"""
        
        temp_xml = f"{task_name}.xml"
        try:
            with open(temp_xml, "w", encoding="utf-16") as f:
                f.write(task_xml)
            
            cmd = f'schtasks /Create /XML "{temp_xml}" /TN "{task_name}"'
            result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
            
            if result.returncode != 0:
                raise Exception(f"Task creation error: {result.stderr}")
        finally:
            if os.path.exists(temp_xml):
                os.remove(temp_xml)
    
    def load_autostart_processes(self):
        self.process_table.setRowCount(0)
        processes = []
        
        try:
            key = winreg.OpenKey(
                winreg.HKEY_CURRENT_USER,
                r"Software\Microsoft\Windows\CurrentVersion\Run",
                0,
                winreg.KEY_READ
            )
            i = 0
            while True:
                try:
                    name, path, _ = winreg.EnumValue(key, i)
                    processes.append(("Registry", name, path))
                    i += 1
                except OSError:
                    break
            winreg.CloseKey(key)
        except FileNotFoundError:
            pass
        
        cmd = 'schtasks /Query /FO CSV /V'
        result = subprocess.run(cmd, shell=True, capture_output=True, text=True)
        if result.returncode == 0:
            lines = result.stdout.splitlines()
            for line in lines[1:]:
                if line and "Autostart_" in line:
                    parts = line.split('","')
                    if len(parts) > 8:
                        task_name = parts[0].strip('"')
                        if task_name.startswith("Autostart_"):
                            cmd_line = parts[8].strip('"')
                            processes.append(("Scheduler", task_name, cmd_line))
        
        for source, name, path in processes:
            row = self.process_table.rowCount()
            self.process_table.insertRow(row)
            
            self.process_table.setItem(row, 0, QTableWidgetItem(name))
            self.process_table.setItem(row, 1, QTableWidgetItem(path))
            
            remove_button = QPushButton("Remove")
            remove_button.setObjectName("removeButton")
            remove_button.setToolTip("Remove this application from autostart")
            remove_button.clicked.connect(lambda _, s=source, n=name: self.remove_process(s, n))
            self.process_table.setCellWidget(row, 2, remove_button)
        logger.info("Loaded autostart processes")
    
    def remove_process(self, source, name):
        reply = QMessageBox.question(
            self, "Confirm", f"Remove '{name}' from autostart?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply != QMessageBox.Yes:
            return
        
        try:
            if source == "Registry":
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Run",
                    0,
                    winreg.KEY_SET_VALUE
                )
                winreg.DeleteValue(key, name)
                winreg.CloseKey(key)
            elif source == "Scheduler":
                subprocess.run(f'schtasks /Delete /TN "{name}" /F', shell=True)
            logger.info(f"Removed from autostart: {name} ({source})")
            QMessageBox.information(self, "Success", f"Application '{name}' removed from autostart!")
            self.load_autostart_processes()
        except Exception as e:
            logger.error(f"Failed to remove from autostart: {name} - {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to remove from autostart: {str(e)}")
    
    def add_batch_file(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Select File", "",
            "Executable files (*.exe *.bat *.cmd);;All files (*.*)"
        )
        if file_path and DragDropLineEdit().is_valid_file(file_path):
            for row in range(self.batch_table.rowCount()):
                if self.batch_table.item(row, 0).text() == file_path:
                    logger.warning(f"Attempted to add duplicate batch file: {file_path}")
                    return
            row = self.batch_table.rowCount()
            self.batch_table.insertRow(row)
            self.batch_table.setItem(row, 0, QTableWidgetItem(file_path))
            self.batch_table.setItem(row, 1, QTableWidgetItem("Stopped"))
            self.save_batch_list()
            logger.info(f"Added batch file: {file_path}")
    
    def remove_batch_file(self):
        selected_rows = sorted(set(index.row() for index in self.batch_table.selectedIndexes()), reverse=True)
        if not selected_rows:
            logger.warning("No batch files selected for removal")
            QMessageBox.warning(self, "Warning", "No items selected to remove!")
            return
        for row in selected_rows:
            file_path = self.batch_table.item(row, 0).text()
            self.batch_table.removeRow(row)
            logger.info(f"Removed batch file: {file_path}")
        self.save_batch_list()
    
    def remove_all_batch_files(self):
        reply = QMessageBox.question(
            self, "Confirm", "Remove all batch files?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            self.batch_table.setRowCount(0)
            self.save_batch_list()
            logger.info("Removed all batch files")
    
    def save_batch_list(self):
        batch_items = [self.batch_table.item(row, 0).text() for row in range(self.batch_table.rowCount())]
        try:
            with open(self.batch_list_file, 'w', encoding='utf-8') as f:
                json.dump(batch_items, f)
            logger.info("Saved batch list")
        except Exception as e:
            logger.error(f"Failed to save batch list: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to save batch list: {str(e)}")
    
    def load_batch_list(self):
        try:
            if os.path.exists(self.batch_list_file):
                with open(self.batch_list_file, 'r', encoding='utf-8') as f:
                    batch_items = json.load(f)
                self.batch_table.setRowCount(0)
                for item in batch_items:
                    if DragDropLineEdit().is_valid_file(item):
                        row = self.batch_table.rowCount()
                        self.batch_table.insertRow(row)
                        self.batch_table.setItem(row, 0, QTableWidgetItem(item))
                        status = "Running" if item in self.processes else "Stopped"
                        self.batch_table.setItem(row, 1, QTableWidgetItem(status))
                logger.info("Loaded batch list")
        except Exception as e:
            logger.error(f"Failed to load batch list: {str(e)}")
            QMessageBox.critical(self, "Error", f"Failed to load batch list: {str(e)}")
    
    def start_all_batch(self):
        self.processes.clear()
        logger.info("Starting all batch processes")
        for row in range(self.batch_table.rowCount()):
            file_path = self.batch_table.item(row, 0).text()
            process_name = os.path.basename(file_path)
            try:
                subprocess.Popen(file_path, shell=True)
                self.processes[file_path] = process_name
                self.batch_table.setItem(row, 1, QTableWidgetItem("Running"))
                logger.info(f"Started batch process: {process_name}")
            except Exception as e:
                logger.error(f"Failed to start batch process: {process_name} - {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to start {process_name}: {str(e)}")
                self.batch_table.setItem(row, 1, QTableWidgetItem("Stopped"))
        self.load_logs()
    
    def stop_all_batch(self):
        logger.info("Stopping all batch processes")
        for file_path, process_name in list(self.processes.items()):
            try:
                result = subprocess.run(
                    f'taskkill /IM "{process_name}" /F',
                    shell=True,
                    capture_output=True,
                    text=True
                )
                if result.returncode != 0:
                    logger.warning(f"Could not stop batch process: {process_name} - {result.stderr}")
                    QMessageBox.warning(self, "Warning", f"Could not stop {process_name}: {result.stderr}")
                else:
                    logger.info(f"Stopped batch process: {process_name}")
                del self.processes[file_path]
            except Exception as e:
                logger.error(f"Failed to stop batch process: {process_name} - {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to stop {process_name}: {str(e)}")
                del self.processes[file_path]
        for row in range(self.batch_table.rowCount()):
            self.batch_table.setItem(row, 1, QTableWidgetItem("Stopped"))
        self.load_logs()
    
    def export_settings(self):
        file_path, _ = QFileDialog.getSaveFileName(
            self, "Export Settings", "", "JSON files (*.json);;All files (*.*)"
        )
        if file_path:
            settings = {
                "batch_list": [self.batch_table.item(row, 0).text() for row in range(self.batch_table.rowCount())],
                "theme": self.current_theme
            }
            try:
                with open(file_path, 'w', encoding='utf-8') as f:
                    json.dump(settings, f, indent=2)
                logger.info(f"Exported settings to {file_path}")
                QMessageBox.information(self, "Success", "Settings exported successfully!")
            except Exception as e:
                logger.error(f"Failed to export settings: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to export settings: {str(e)}")
    
    def import_settings(self):
        file_path, _ = QFileDialog.getOpenFileName(
            self, "Import Settings", "", "JSON files (*.json);;All files (*.*)"
        )
        if file_path:
            try:
                with open(file_path, 'r', encoding='utf-8') as f:
                    settings = json.load(f)
                
                self.batch_table.setRowCount(0)
                for item in settings.get("batch_list", []):
                    if DragDropLineEdit().is_valid_file(item):
                        row = self.batch_table.rowCount()
                        self.batch_table.insertRow(row)
                        self.batch_table.setItem(row, 0, QTableWidgetItem(item))
                        self.batch_table.setItem(row, 1, QTableWidgetItem("Stopped"))
                self.save_batch_list()
                
                theme = settings.get("theme", "light")
                if theme in ["light", "dark"]:
                    self.current_theme = theme
                    self.apply_theme()
                    self.save_theme()
                
                logger.info(f"Imported settings from {file_path}")
                QMessageBox.information(self, "Success", "Settings imported successfully!")
            except Exception as e:
                logger.error(f"Failed to import settings: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to import settings: {str(e)}")
    
    def load_logs(self):
        try:
            with open('autostart_manager.log', 'r', encoding='utf-8') as f:
                lines = f.readlines()[-100:]
                self.log_text.setText(''.join(lines))
            logger.info("Loaded logs")
        except Exception as e:
            logger.error(f"Failed to load logs: {str(e)}")
            self.log_text.setText(f"Failed to load logs: {str(e)}")
    
    def clear_logs(self):
        reply = QMessageBox.question(
            self, "Confirm", "Clear all logs?",
            QMessageBox.Yes | QMessageBox.No, QMessageBox.No
        )
        if reply == QMessageBox.Yes:
            try:
                open('autostart_manager.log', 'w').close()
                self.log_text.clear()
                logger.info("Cleared logs")
            except Exception as e:
                logger.error(f"Failed to clear logs: {str(e)}")
                QMessageBox.critical(self, "Error", f"Failed to clear logs: {str(e)}")

if __name__ == "__main__":
    app = QApplication(sys.argv)
    window = AutostartManager()
    window.show()
    sys.exit(app.exec())