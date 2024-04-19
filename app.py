import sys

import pandas as pd
from PyQt6.QtWidgets import QApplication, QMainWindow, QFileDialog, QMessageBox, QTextEdit, QPushButton, QInputDialog


class SpreadsheetApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spreadsheet Application")
        self.setGeometry(100, 100, 800, 600)

        self.file_path = ""
        self.data = None

        self.load_button = QPushButton("Load Spreadsheet", self)
        self.load_button.move(20, 20)
        self.load_button.clicked.connect(self.load_file)

        self.append_button = QPushButton("Append", self)
        self.append_button.move(20, 60)
        self.append_button.clicked.connect(self.append_data)

        self.merge_button = QPushButton("Merge", self)
        self.merge_button.move(20, 100)
        self.merge_button.clicked.connect(self.merge_data)

        self.pivot_button = QPushButton("Pivot", self)
        self.pivot_button.move(20, 140)
        self.pivot_button.clicked.connect(self.pivot_data)

        self.unpivot_button = QPushButton("Unpivot", self)
        self.unpivot_button.move(20, 180)
        self.unpivot_button.clicked.connect(self.unpivot_data)

        self.save_button = QPushButton("Save Spreadsheet", self)
        self.save_button.move(20, 220)
        self.save_button.clicked.connect(self.save_file)

        self.preview_text = QTextEdit(self)
        self.preview_text.setGeometry(200, 20, 580, 560)

    def load_file(self):
        options = QFileDialog.Option.ReadOnly
        self.file_path, _ = QFileDialog.getOpenFileName(self, "Load Spreadsheet", "",
                                                        "Excel files (*.xlsx);;CSV files (*.csv)", options=options)
        if self.file_path:
            if self.file_path.endswith(".xlsx"):
                self.data = pd.read_excel(self.file_path)
            elif self.file_path.endswith(".csv"):
                self.data = pd.read_csv(self.file_path)
            self.preview_text.clear()
            self.preview_text.insertPlainText(self.data.to_string(index=False))

    def append_data(self):
        if self.data is not None:
            options = QFileDialog.Option.ReadOnly
            file_path, _ = QFileDialog.getOpenFileName(self, "Append Data", "",
                                                       "Excel files (*.xlsx);;CSV files (*.csv)", options=options)
            if file_path:
                if file_path.endswith(".xlsx"):
                    new_data = pd.read_excel(file_path)
                elif file_path.endswith(".csv"):
                    new_data = pd.read_csv(file_path)
                self.data = self.data.append(new_data, ignore_index=True)
                self.preview_text.clear()
                self.preview_text.insertPlainText(self.data.to_string(index=False))
        else:
            QMessageBox.information(self, "Error", "No data loaded. Please load a spreadsheet first.")

    def merge_data(self):
        if self.data is not None:
            options = QFileDialog.Option.ReadOnly
            file_path, _ = QFileDialog.getOpenFileName(self, "Merge Data", "",
                                                       "Excel files (*.xlsx);;CSV files (*.csv)", options=options)
            if file_path:
                if file_path.endswith(".xlsx"):
                    new_data = pd.read_excel(file_path)
                elif file_path.endswith(".csv"):
                    new_data = pd.read_csv(file_path)
                merge_column, ok = QInputDialog.getText(self, "Merge Column", "Enter the column name to merge on:")
                if ok and merge_column:
                    self.data = pd.merge(self.data, new_data, on=merge_column)
                    self.preview_text.clear()
                    self.preview_text.insertPlainText(self.data.to_string(index=False))
        else:
            QMessageBox.information(self, "Error", "No data loaded. Please load a spreadsheet first.")

    def pivot_data(self):
        if self.data is not None:
            pivot_index, ok = QInputDialog.getText(self, "Pivot Index", "Enter the column name for the pivot index:")
            if ok and pivot_index:
                pivot_columns, ok = QInputDialog.getText(self, "Pivot Columns",
                                                         "Enter the column name for the pivot columns:")
                if ok and pivot_columns:
                    pivot_values, ok = QInputDialog.getText(self, "Pivot Values",
                                                            "Enter the column name for the pivot values:")
                    if ok and pivot_values:
                        self.data = self.data.pivot_table(index=pivot_index, columns=pivot_columns, values=pivot_values)
                        self.preview_text.clear()
                        self.preview_text.insertPlainText(self.data.to_string())
        else:
            QMessageBox.information(self, "Error", "No data loaded. Please load a spreadsheet first.")

    def unpivot_data(self):
        if self.data is not None:
            self.data = self.data.reset_index()
            self.data = pd.melt(self.data, id_vars=self.data.columns[0], value_vars=self.data.columns[1:])
            self.preview_text.clear()
            self.preview_text.insertPlainText(self.data.to_string(index=False))
        else:
            QMessageBox.information(self, "Error", "No data loaded. Please load a spreadsheet first.")

    def save_file(self):
        if self.data is not None:
            options = QFileDialog.Option.WriteOnly
            file_path, _ = QFileDialog.getSaveFileName(self, "Save Spreadsheet", "",
                                                       "Excel files (*.xlsx);;CSV files (*.csv)", options=options)
            if file_path:
                if file_path.endswith(".xlsx"):
                    self.data.to_excel(file_path, index=False)
                elif file_path.endswith(".csv"):
                    self.data.to_csv(file_path, index=False)
                QMessageBox.information(self, "Save", "Data saved successfully.")
        else:
            QMessageBox.information(self, "Error", "No data loaded. Please load a spreadsheet first.")


if __name__ == "__main__":
    app = QApplication(sys.argv)
    spreadsheet_app = SpreadsheetApp()
    spreadsheet_app.show()
    sys.exit(app.exec())
