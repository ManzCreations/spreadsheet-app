import sys

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction
from PyQt6.QtWidgets import QApplication, QMainWindow, QTabWidget, QWidget, QVBoxLayout, QHBoxLayout, QListWidget, \
    QTableWidget, QTableWidgetItem, QHeaderView, QMenu, QFileDialog, QMessageBox


class SpreadsheetApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spreadsheet Application")
        self.setGeometry(100, 100, 800, 600)

        self.tables = {}

        self.init_ui()

    def init_ui(self):
        main_layout = QHBoxLayout()

        # File Organizer
        self.file_list = QListWidget()
        self.file_list.itemClicked.connect(self.show_table)
        main_layout.addWidget(self.file_list)

        right_layout = QVBoxLayout()

        # Tab Widget
        self.tab_widget = QTabWidget()
        self.tab_widget.setTabShape(QTabWidget.TabShape.Triangular)
        right_layout.addWidget(self.tab_widget)

        # Table View
        self.table_view = QTableWidget()
        self.table_view.setColumnCount(0)
        self.table_view.setRowCount(0)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.Stretch)
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self.show_context_menu)
        right_layout.addWidget(self.table_view)

        main_layout.addLayout(right_layout)

        central_widget = QWidget()
        central_widget.setLayout(main_layout)
        self.setCentralWidget(central_widget)

        self.init_menu()

    def init_menu(self):
        menubar = self.menuBar()

        file_menu = menubar.addMenu("File")

        add_table_action = QAction("Add Table", self)
        add_table_action.triggered.connect(self.add_table)
        file_menu.addAction(add_table_action)

        operations_menu = menubar.addMenu("Operations")

        merge_action = QAction("Merge", self)
        merge_action.triggered.connect(self.merge_tables)
        operations_menu.addAction(merge_action)

        append_action = QAction("Append", self)
        append_action.triggered.connect(self.append_tables)
        operations_menu.addAction(append_action)

        pivot_action = QAction("Pivot", self)
        pivot_action.triggered.connect(self.pivot_table)
        operations_menu.addAction(pivot_action)

    def add_table(self):
        options = QFileDialog.Option.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, "Add Table", "", "Excel files (*.xlsx);;CSV files (*.csv)",
                                                   options=options)
        if file_path:
            if file_path.endswith(".xlsx"):
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                for sheet_name in sheet_names:
                    data = excel_file.parse(sheet_name)
                    table_name = f"{file_path.split('/')[-1]} - {sheet_name}"
                    self.tables[table_name] = data
                    self.file_list.addItem(table_name)
            elif file_path.endswith(".csv"):
                data = pd.read_csv(file_path)
                table_name = file_path.split("/")[-1]
                self.tables[table_name] = data
                self.file_list.addItem(table_name)

    def show_table(self, item):
        table_name = item.text()
        data = self.tables[table_name]
        self.populate_table(data)

    def populate_table(self, data):
        self.table_view.setColumnCount(len(data.columns))
        self.table_view.setRowCount(len(data))
        self.table_view.setHorizontalHeaderLabels([f"{col} ({dtype})" for col, dtype in zip(data.columns, data.dtypes)])

        for i in range(len(data)):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                self.table_view.setItem(i, j, item)

    def show_context_menu(self, pos):
        menu = QMenu(self)

        delete_action = QAction("Delete", self)
        delete_action.triggered.connect(self.delete_item)
        menu.addAction(delete_action)

        add_column_action = QAction("Add Column", self)
        add_column_action.triggered.connect(self.add_column)
        menu.addAction(add_column_action)

        add_row_action = QAction("Add Row", self)
        add_row_action.triggered.connect(self.add_row)
        menu.addAction(add_row_action)

        menu.exec(self.table_view.mapToGlobal(pos))

    def delete_item(self):
        current_row = self.table_view.currentRow()
        current_column = self.table_view.currentColumn()

        if current_row != -1:
            self.table_view.removeRow(current_row)
        elif current_column != -1:
            self.table_view.removeColumn(current_column)

    def add_column(self):
        self.table_view.insertColumn(self.table_view.columnCount())

    def add_row(self):
        self.table_view.insertRow(self.table_view.rowCount())

    def merge_tables(self):
        if len(self.tables) < 2:
            QMessageBox.warning(self, "Error", "At least two tables are required for merging.")
            return

        merged_data = None
        for data in self.tables.values():
            if merged_data is None:
                merged_data = data
            else:
                merged_data = pd.merge(merged_data, data, how="outer")

        self.populate_table(merged_data)

    def append_tables(self):
        if len(self.tables) < 2:
            QMessageBox.warning(self, "Error", "At least two tables are required for appending.")
            return

        appended_data = pd.concat(list(self.tables.values()), ignore_index=True)
        self.populate_table(appended_data)

    def pivot_table(self):
        if len(self.tables) == 0:
            QMessageBox.warning(self, "Error", "No tables available for pivoting.")
            return

        data = list(self.tables.values())[0]
        pivot_data = data.pivot_table(index=data.columns[0], values=data.columns[1], aggfunc="sum")
        self.populate_table(pivot_data)


if __name__ == "__main__":
    app = QApplication(sys.argv)
    spreadsheet_app = SpreadsheetApp()
    spreadsheet_app.show()
    sys.exit(app.exec())
