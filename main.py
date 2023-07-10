import json
import random
import sys

import pandas as pd
from PyQt5.QtCore import Qt
from PyQt5.QtGui import QImage, QPixmap, QColor
from PyQt5.QtWidgets import (QApplication, QMainWindow, QWidget, QLabel, QTextEdit, QVBoxLayout, QHBoxLayout,
                             QPushButton, QComboBox, QFileDialog, QMessageBox)

# Constantes para os textos
APP_WINDOW_NAME = "Aplicativo de Sorteio"
IMPORT_BUTTON_TEXT = "Importar"
CLEAR_BUTTON_TEXT = "Limpar"
EXPORT_FORMAT_SELECT = "-- SELECIONE --"
EXPORT_FORMAT_JSON = "JSON"
EXPORT_FORMAT_EXCEL = "EXCEL"
EXPORT_FORMAT_OPTIONS = [EXPORT_FORMAT_SELECT, EXPORT_FORMAT_JSON, EXPORT_FORMAT_EXCEL]
JSON_FILE_EXTENSION = "json"
EXCEL_FILE_EXTENSION = "xlsx"
EXPORT_BUTTON_TEXT = "Exportar"
EXIT_BUTTON_TEXT = "Sair"
RESUME_LABEL_TEXT = "OS IMPERIALISTAS SÃO TIGRES DE PAPEL"
QUANTITY_LABEL_TEXT = "Quantidade:"
PICK_BUTTON_TEXT = "Sortear"
ERROR_TITLE = "Erro"
ERROR_INVALID_QUANTITY = "A quantidade inserida é inválida."
ERROR_INVALID_EXPORT_FORMAT = "Selecione um formato de exportação válido."
JSON_FILE_TYPE = f"Arquivo JSON (*.{JSON_FILE_EXTENSION})"
EXCEL_FILE_TYPE = f"Planilha Excel (*.{EXCEL_FILE_EXTENSION})"
IMPORT_SHORTCUT = " (Ctrl+V)"
EXPORT_SHORTCUT = " (Ctrl+S)"
CLEAR_SHORTCUT = " (Ctrl+L)"
EXIT_SHORTCUT = " (Ctrl+Q)"
PICK_SHORTCUT = " (Ctrl+R)"
BACKGROUND_IMG_PATH = "assets/wallpaper.jpg"
COLOR_WHITE = "white"
NEW_ENTRY_TEXT = "Novas entradas:"
REPEATED_ENTRIES_TEXT = "Entradas repetidas:"
TOTAL_ENTRIES_TEXT = "Total de entradas:"
EK_TAB = "\t"
EK_LINE_BREAK = "\n"
SAVE_AS_TEXT = "Salvar como"
COLUMN_ID = "ID"
COLUMN_NAME = "Nome"


class App(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle(APP_WINDOW_NAME)
        self.setMinimumSize(800, 450)
        self.setMaximumSize(1600, 900)

        self.bg_image = QImage(BACKGROUND_IMG_PATH)
        self.bg_label = QLabel(self)
        self.bg_label.setPixmap(QPixmap.fromImage(self.bg_image))
        self.setCentralWidget(self.bg_label)

        self.main_layout = QVBoxLayout(self.bg_label)

        self.create_top_bar()
        self.create_center()
        self.create_bottom_bar()

        self.ctrl_pressed = False
        self.update_button_shortcuts()

    def create_top_bar(self):
        self.top_bar_widget = QWidget(self)
        self.top_bar_layout = QHBoxLayout(self.top_bar_widget)

        self.import_button = QPushButton(IMPORT_BUTTON_TEXT, self)
        self.top_bar_layout.addWidget(self.import_button)
        self.import_button.clicked.connect(self.import_names)
        button_palette = self.import_button.palette()
        button_palette.setColor(self.import_button.foregroundRole(), QColor("black"))
        self.import_button.setPalette(button_palette)
        self.import_button.setStyleSheet("background-color: lightgray;")

        self.clear_button = QPushButton(CLEAR_BUTTON_TEXT, self)
        self.top_bar_layout.addWidget(self.clear_button)
        self.clear_button.clicked.connect(self.clear_sheet)
        button_palette = self.clear_button.palette()
        button_palette.setColor(self.clear_button.foregroundRole(), QColor("black"))
        self.clear_button.setPalette(button_palette)
        self.clear_button.setStyleSheet("background-color: lightgray;")

        self.export_dropdown = QComboBox(self)
        self.export_dropdown.addItems(EXPORT_FORMAT_OPTIONS)
        self.top_bar_layout.addWidget(self.export_dropdown)
        button_palette = self.export_dropdown.palette()
        button_palette.setColor(self.export_dropdown.foregroundRole(), QColor("black"))
        self.export_dropdown.setPalette(button_palette)
        self.export_dropdown.setStyleSheet("background-color: lightgray;")

        self.export_button = QPushButton(EXPORT_BUTTON_TEXT, self)
        self.top_bar_layout.addWidget(self.export_button)
        self.export_button.clicked.connect(self.export_results)
        button_palette = self.export_button.palette()
        button_palette.setColor(self.export_button.foregroundRole(), QColor("black"))
        self.export_button.setPalette(button_palette)
        self.export_button.setStyleSheet("background-color: lightgray;")

        self.exit_button = QPushButton(EXIT_BUTTON_TEXT, self)
        self.top_bar_layout.addWidget(self.exit_button)
        self.exit_button.clicked.connect(self.close)
        button_palette = self.exit_button.palette()
        button_palette.setColor(self.exit_button.foregroundRole(), QColor("black"))
        self.exit_button.setPalette(button_palette)
        self.exit_button.setStyleSheet("background-color: lightgray;")

        self.main_layout.addWidget(self.top_bar_widget)

    def create_center(self):
        self.center_widget = QWidget(self)
        self.center_layout = QHBoxLayout(self.center_widget)

        self.sheet_textbox = QTextEdit(self)
        self.sheet_textbox.setStyleSheet("background-color: white; color: black;")
        self.center_layout.addWidget(self.sheet_textbox)

        self.results_textbox = QTextEdit(self)
        self.results_textbox.setStyleSheet("background-color: white; color: black;")
        self.center_layout.addWidget(self.results_textbox)

        self.main_layout.addWidget(self.center_widget, 1)

    def create_bottom_bar(self):
        self.bottom_bar_widget = QWidget(self)
        self.bottom_bar_layout = QVBoxLayout(self.bottom_bar_widget)

        self.bottom_bar_layout.setContentsMargins(20, 10, 20, 10)
        self.bottom_bar_layout.setSpacing(10)

        # Top Row
        top_row_widget = QWidget(self)
        top_row_layout = QHBoxLayout(top_row_widget)

        self.summary_label = QLabel(RESUME_LABEL_TEXT, self)
        summary_label_palette = self.summary_label.palette()
        summary_label_palette.setColor(self.summary_label.foregroundRole(), QColor(COLOR_WHITE))
        self.summary_label.setPalette(summary_label_palette)
        top_row_layout.addWidget(self.summary_label)

        self.bottom_bar_layout.addWidget(top_row_widget)

        # Bottom Row
        bottom_row_widget = QWidget(self)
        bottom_row_layout = QHBoxLayout(bottom_row_widget)

        self.quantity_label = QLabel(QUANTITY_LABEL_TEXT, self)
        bottom_row_layout.addWidget(self.quantity_label)

        self.quantity_entry = QTextEdit(self)
        self.quantity_entry.setMaximumHeight(30)
        bottom_row_layout.addWidget(self.quantity_entry)

        self.pick_button = QPushButton(PICK_BUTTON_TEXT, self)
        bottom_row_layout.addWidget(self.pick_button)
        self.pick_button.clicked.connect(self.pick_names)
        button_palette = self.pick_button.palette()
        button_palette.setColor(self.pick_button.foregroundRole(), QColor("black"))
        self.pick_button.setPalette(button_palette)
        self.pick_button.setStyleSheet("background-color: lightgray;")

        self.bottom_bar_layout.addWidget(bottom_row_widget)

        self.main_layout.addWidget(self.bottom_bar_widget, alignment=Qt.AlignBottom)

    def keyPressEvent(self, event):
        if event.key() == Qt.Key_Control:
            self.ctrl_pressed = True
            self.update_button_shortcuts()

    def keyReleaseEvent(self, event):
        if event.key() == Qt.Key_Control:
            self.ctrl_pressed = False
            self.update_button_shortcuts()

    def update_button_shortcuts(self):
        import_button_text = f"{IMPORT_BUTTON_TEXT} {IMPORT_SHORTCUT if self.ctrl_pressed else ''}"
        export_button_text = f"{EXPORT_BUTTON_TEXT} {EXPORT_SHORTCUT if self.ctrl_pressed else ''}"
        clear_button_text = f"{CLEAR_BUTTON_TEXT} {CLEAR_SHORTCUT if self.ctrl_pressed else ''}"
        exit_button_text = f"{EXIT_BUTTON_TEXT} {EXIT_SHORTCUT if self.ctrl_pressed else ''}"
        pick_button_text = f"{PICK_BUTTON_TEXT} {PICK_SHORTCUT if self.ctrl_pressed else ''}"

        self.import_button.setText(import_button_text)
        self.export_button.setText(export_button_text)
        self.clear_button.setText(clear_button_text)
        self.exit_button.setText(exit_button_text)
        self.pick_button.setText(pick_button_text)

    def import_names(self):
        clipboard_text = QApplication.clipboard().text()
        names = clipboard_text.split(EK_LINE_BREAK)
        new_entries = 0
        repeated_entries = 0

        for name in names:
            name = name.strip()
            if name and not any(name in entry for entry in self.sheet_textbox.toPlainText().split(EK_LINE_BREAK)):
                new_entries += 1
                self.sheet_textbox.append(f"{COLUMN_ID}: {new_entries}\t{COLUMN_NAME}: {name}")
            elif name:
                repeated_entries += 1

        self.summary_label.setText(
            f"{TOTAL_ENTRIES_TEXT} {new_entries} | "
            f"{REPEATED_ENTRIES_TEXT} {repeated_entries} | "
            f"{TOTAL_ENTRIES_TEXT} {new_entries + repeated_entries}"
        )

    def clear_sheet(self):
        self.sheet_textbox.clear()
        self.results_textbox.clear()
        self.summary_label.setText(RESUME_LABEL_TEXT)

    def pick_names(self):
        quantity = self.quantity_entry.toPlainText()
        names_list = self.sheet_textbox.toPlainText().split(EK_LINE_BREAK)
        names_list = [name.split(EK_TAB)[1] for name in names_list if name]

        if quantity.isdigit() and int(quantity) <= len(names_list):
            results = random.sample(names_list, int(quantity))
            self.results_textbox.clear()
            self.results_textbox.append(EK_LINE_BREAK.join(results))
        else:
            QMessageBox.critical(self, ERROR_TITLE, ERROR_INVALID_QUANTITY)

    def export_results(self):
        export_format = self.export_dropdown.currentText()

        if export_format == EXPORT_FORMAT_JSON:
            self.export_to_json()
        elif export_format == EXPORT_FORMAT_EXCEL:
            self.export_to_excel()
        else:
            QMessageBox.critical(self, ERROR_TITLE, ERROR_INVALID_EXPORT_FORMAT)

    def export_to_json(self):
        file_dialog = QFileDialog(self)
        file_dialog.setDefaultSuffix(JSON_FILE_EXTENSION)
        file_path, _ = file_dialog.getSaveFileName(
            self, SAVE_AS_TEXT, "", JSON_FILE_TYPE
        )

        if file_path:
            results = self.results_textbox.toPlainText().split(EK_LINE_BREAK)
            results = [result.strip() for result in results if result]
            data = {"results": results}

            with open(file_path, "w") as file:
                json.dump(data, file)

    def export_to_excel(self):
        file_dialog = QFileDialog(self)
        file_dialog.setDefaultSuffix(EXCEL_FILE_EXTENSION)
        file_path, _ = file_dialog.getSaveFileName(
            self, SAVE_AS_TEXT, "", EXCEL_FILE_TYPE
        )

        if file_path:
            results = self.results_textbox.toPlainText().split(EK_LINE_BREAK)
            results = [result.strip() for result in results if result]
            data = {"results": results}
            df = pd.DataFrame(data)

            with pd.ExcelWriter(file_path) as writer:
                df.to_excel(writer, index=False)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setStyle("Fusion")
    window = App()
    window.show()
    sys.exit(app.exec_())
