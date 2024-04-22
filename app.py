import sys

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QFont, QIcon
from PyQt6.QtWidgets import *


class MergeDialog(QDialog):
    def __init__(self, tables, selected_table, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Merge Tables")
        self.setWindowIcon(QIcon("images/crm-icon-high-seas.png"))
        self.setGeometry(100, 100, 800, 500)

        self.tables = tables  # Store the tables dictionary as an instance variable
        self.selected_column1 = None
        self.selected_column2 = None

        layout = QVBoxLayout()

        # Dropdowns
        dropdown_layout = QHBoxLayout()
        self.table1_dropdown = QComboBox()
        self.table1_dropdown.addItems(list(tables.keys()))
        self.table1_dropdown.setCurrentText(selected_table)
        dropdown_layout.addWidget(QLabel("Table 1:"))
        dropdown_layout.addWidget(self.table1_dropdown)
        dropdown_layout.addStretch()
        layout.addLayout(dropdown_layout)

        # Table Views
        table_layout = QVBoxLayout()
        self.table1_view = self.create_table_view(tables[selected_table])
        table_layout.addWidget(self.table1_view)
        layout.addLayout(table_layout)

        # Dropdowns
        dropdown_layout = QHBoxLayout()
        self.table2_dropdown = QComboBox()
        self.table2_dropdown.addItems([""] + list(tables.keys()))
        dropdown_layout.addWidget(QLabel("Table 2:"))
        dropdown_layout.addWidget(self.table2_dropdown)
        dropdown_layout.addStretch()
        layout.addLayout(dropdown_layout)

        # Table Views
        table_layout = QVBoxLayout()
        self.table2_view = QTableWidget()
        table_layout.addWidget(self.table2_view)
        layout.addLayout(table_layout)

        # Join Type Dropdown
        join_layout = QHBoxLayout()
        self.join_dropdown = QComboBox()
        self.join_dropdown.addItems(["Inner Join", "Left Join", "Right Join", "Outer Join"])
        join_layout.addWidget(QLabel("Join Type:"))
        join_layout.addWidget(self.join_dropdown)
        layout.addLayout(join_layout)

        # Button
        self.merge_button = QPushButton("Merge")
        self.merge_button.clicked.connect(self.merge_tables)
        layout.addWidget(self.merge_button)

        self.setLayout(layout)

        self.table1_dropdown.currentTextChanged.connect(self.update_table1_view)
        self.table2_dropdown.currentTextChanged.connect(self.update_table2_view)

    def create_table_view(self, table_revision):
        data = table_revision.revisions[table_revision.current_revision]
        table_view = QTableWidget()
        table_view.setColumnCount(len(data.columns))
        table_view.setRowCount(3)
        table_view.setHorizontalHeaderLabels(data.columns)
        table_view.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        table_view.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectColumns)

        for i in range(3):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                table_view.setItem(i, j, item)

        table_view.resizeColumnsToContents()
        table_view.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        table_view.setVerticalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(table_view)

        table_view.itemSelectionChanged.connect(self.update_selected_column)

        return scroll_area

    def update_table1_view(self, table_name):
        table_revision = self.tables[table_name]
        self.update_table_view(self.table1_view, table_revision)

    def update_table2_view(self, table_name):
        if table_name:
            table_revision = self.tables[table_name]
            self.update_table_view(self.table2_view, table_revision)
        else:
            self.table2_view.clear()
            self.table2_view.setColumnCount(0)
            self.table2_view.setRowCount(0)

    def update_table_view(self, table_view, table_revision):
        data = table_revision.revisions[table_revision.current_revision]

        if isinstance(table_view, QScrollArea):
            table_widget = table_view.widget()
        else:
            table_widget = table_view

        table_widget.setColumnCount(len(data.columns))
        table_widget.setRowCount(3)
        table_widget.setHorizontalHeaderLabels(data.columns)

        for i in range(3):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                table_widget.setItem(i, j, item)

        table_widget.resizeColumnsToContents()
        table_widget.itemSelectionChanged.connect(self.update_selected_column)

    def update_selected_column(self):
        sender = self.sender()
        if sender == self.table1_view.widget():
            selected_indexes = sender.selectedIndexes()
            if selected_indexes:
                self.selected_column1 = selected_indexes[0].column()
        elif sender == self.table2_view:
            selected_indexes = sender.selectedIndexes()
            if selected_indexes:
                self.selected_column2 = selected_indexes[0].column()

    def merge_tables(self):
        table1_name = self.table1_dropdown.currentText()
        table2_name = self.table2_dropdown.currentText()
        table1_revision = self.tables[table1_name]
        table2_revision = self.tables[table2_name]

        table1_data = table1_revision.revisions[table1_revision.current_revision]
        table2_data = table2_revision.revisions[table2_revision.current_revision]

        if self.selected_column1 is None or self.selected_column2 is None:
            QMessageBox.warning(self, "Error", "Please select a column from each table.")
            return

        merge_column1 = table1_data.columns[self.selected_column1]
        merge_column2 = table2_data.columns[self.selected_column2]

        join_type = self.join_dropdown.currentText().lower().split(" ")[0]

        merged_data = pd.merge(table1_data, table2_data, left_on=merge_column1, right_on=merge_column2, how=join_type)
        self.parent().populate_table(merged_data)
        self.close()


class AppendDialog(QDialog):
    def __init__(self, tables, selected_table, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Append Tables")
        self.setWindowIcon(QIcon("images/crm-icon-high-seas.png"))
        self.setGeometry(100, 100, 800, 500)

        self.tables = tables  # Store the tables dictionary as an instance variable
        self.selected_column1 = None
        self.selected_column2 = None

        layout = QVBoxLayout()

        # Dropdowns
        dropdown_layout = QHBoxLayout()
        self.table1_dropdown = QComboBox()
        self.table1_dropdown.addItems(list(tables.keys()))
        self.table1_dropdown.setCurrentText(selected_table)
        dropdown_layout.addWidget(QLabel("Table 1:"))
        dropdown_layout.addWidget(self.table1_dropdown)
        dropdown_layout.addStretch()
        layout.addLayout(dropdown_layout)

        # Table Views
        table_layout = QVBoxLayout()
        self.table1_view = self.create_table_view(tables[selected_table])
        table_layout.addWidget(self.table1_view)
        layout.addLayout(table_layout)

        # Dropdowns
        dropdown_layout = QHBoxLayout()
        self.table2_dropdown = QComboBox()
        self.table2_dropdown.addItems([""] + list(tables.keys()))
        dropdown_layout.addWidget(QLabel("Table 2:"))
        dropdown_layout.addWidget(self.table2_dropdown)
        dropdown_layout.addStretch()
        layout.addLayout(dropdown_layout)

        # Table Views
        table_layout = QVBoxLayout()
        self.table2_view = QTableWidget()
        table_layout.addWidget(self.table2_view)
        layout.addLayout(table_layout)

        # Append Direction Dropdown
        direction_layout = QHBoxLayout()
        self.direction_dropdown = QComboBox()
        self.direction_dropdown.addItems(["Vertically", "Horizontally"])
        direction_layout.addWidget(QLabel("Append Direction:"))
        direction_layout.addWidget(self.direction_dropdown)
        layout.addLayout(direction_layout)

        # Button
        self.append_button = QPushButton("Append")
        self.append_button.clicked.connect(self.append_tables)
        layout.addWidget(self.append_button)

        self.setLayout(layout)

        self.table1_dropdown.currentTextChanged.connect(self.update_table1_view)
        self.table2_dropdown.currentTextChanged.connect(self.update_table2_view)

    def create_table_view(self, table_revision):
        data = table_revision.revisions[table_revision.current_revision]
        table_view = QTableWidget()
        table_view.setColumnCount(len(data.columns))
        table_view.setRowCount(3)
        table_view.setHorizontalHeaderLabels(data.columns)
        table_view.setSelectionMode(QTableWidget.SelectionMode.SingleSelection)
        table_view.setSelectionBehavior(QTableWidget.SelectionBehavior.SelectColumns)

        for i in range(3):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                table_view.setItem(i, j, item)

        table_view.resizeColumnsToContents()
        table_view.setHorizontalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)
        table_view.setVerticalScrollMode(QTableWidget.ScrollMode.ScrollPerPixel)

        scroll_area = QScrollArea()
        scroll_area.setWidgetResizable(True)
        scroll_area.setWidget(table_view)

        table_view.itemSelectionChanged.connect(self.update_selected_column)

        return scroll_area

    def update_table1_view(self, table_name):
        table_revision = self.tables[table_name]
        self.update_table_view(self.table1_view, table_revision)

    def update_table2_view(self, table_name):
        if table_name:
            table_revision = self.tables[table_name]
            self.update_table_view(self.table2_view, table_revision)
        else:
            self.table2_view.clear()
            self.table2_view.setColumnCount(0)
            self.table2_view.setRowCount(0)

    def update_table_view(self, table_view, table_revision):
        data = table_revision.revisions[table_revision.current_revision]

        if isinstance(table_view, QScrollArea):
            table_widget = table_view.widget()
        else:
            table_widget = table_view

        table_widget.setColumnCount(len(data.columns))
        table_widget.setRowCount(3)
        table_widget.setHorizontalHeaderLabels(data.columns)

        for i in range(3):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                table_widget.setItem(i, j, item)

        table_widget.resizeColumnsToContents()
        table_widget.itemSelectionChanged.connect(self.update_selected_column)

    def update_selected_column(self):
        sender = self.sender()
        if sender == self.table1_view.widget():
            selected_indexes = sender.selectedIndexes()
            if selected_indexes:
                self.selected_column1 = selected_indexes[0].column()
        elif sender == self.table2_view:
            selected_indexes = sender.selectedIndexes()
            if selected_indexes:
                self.selected_column2 = selected_indexes[0].column()

    def append_tables(self):
        table1_name = self.table1_dropdown.currentText()
        table2_name = self.table2_dropdown.currentText()
        table1_revision = self.tables[table1_name]
        table2_revision = self.tables[table2_name]

        table1_data = table1_revision.revisions[table1_revision.current_revision]
        table2_data = table2_revision.revisions[table2_revision.current_revision]

        append_direction = self.direction_dropdown.currentText().lower()

        if append_direction == "vertically":
            appended_data = pd.concat([table1_data, table2_data], ignore_index=True)
        else:
            appended_data = pd.concat([table1_data, table2_data], axis=1)

        self.parent().populate_table(appended_data)
        self.close()


class TableRevision:
    def __init__(self, data):
        self.data = data
        self.revisions = [data]
        self.current_revision = 0

    def add_revision(self, data):
        if len(self.revisions) >= 10:
            self.revisions.pop(0)
        self.revisions.append(data)
        self.current_revision = len(self.revisions) - 1

    def undo(self):
        if self.current_revision > 0:
            self.current_revision -= 1
            self.data = self.revisions[self.current_revision]
            return 0
        else:
            return -1

    def redo(self):
        if self.current_revision < len(self.revisions) - 1:
            self.current_revision += 1
            self.data = self.revisions[self.current_revision]
            return 0
        else:
            return -1


class SpreadsheetApp(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spreadsheet Application")
        self.setWindowIcon(QIcon("images/crm-icon-high-seas.png"))
        self.setGeometry(100, 100, 800, 600)

        self.tables = {}
        self.pressed_keys = set()

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # Revision buttons
        revision_button_layout = QHBoxLayout()
        undo_button = QPushButton("Undo")
        undo_button.clicked.connect(self.undo_revision)
        revision_button_layout.addWidget(undo_button)

        redo_button = QPushButton("Redo")
        redo_button.clicked.connect(self.redo_revision)
        revision_button_layout.addWidget(redo_button)

        revision_button_layout.addStretch()
        main_layout.addLayout(revision_button_layout)

        file_table_layout = QHBoxLayout()

        # File Organizer
        self.file_list = QListWidget()
        self.file_list.setMaximumWidth(200)  # Set a maximum width for the file list
        self.file_list.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.file_list.itemDoubleClicked.connect(self.show_table)
        self.file_list.itemChanged.connect(self.rename_table)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_file_context_menu)
        file_table_layout.addWidget(self.file_list)

        # Table View
        self.table_view = QTableWidget()
        self.table_view.setColumnCount(0)
        self.table_view.setRowCount(0)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table_view.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table_view.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.customContextMenuRequested.connect(self.show_context_menu)
        self.table_view.horizontalHeader().setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.table_view.horizontalHeader().sectionDoubleClicked.connect(self.rename_column)
        file_table_layout.addWidget(self.table_view)

        main_layout.addLayout(file_table_layout)

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
                    self.tables[table_name] = TableRevision(data)
                    item = QListWidgetItem(table_name)
                    self.file_list.addItem(item)
                    self.file_list.setCurrentItem(item)
            elif file_path.endswith(".csv"):
                data = pd.read_csv(file_path)
                table_name = file_path.split("/")[-1]
                self.tables[table_name] = TableRevision(data)
                self.file_list.addItem(table_name)
            self.show_table(self.file_list.currentItem())

    def populate_table(self, data):
        self.table_view.setColumnCount(len(data.columns))
        self.table_view.setRowCount(len(data))
        self.table_view.setHorizontalHeaderLabels([f"{col} ({dtype})" for col, dtype in zip(data.columns, data.dtypes)])

        for i in range(len(data)):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                self.table_view.setItem(i, j, item)

    def show_file_context_menu(self, pos):
        item = self.file_list.itemAt(pos)
        if item:
            menu = QMenu(self)
            rename_action = menu.addAction("Rename Table")
            delete_action = menu.addAction("Delete Table")
            rollback_action = menu.addAction("Rollback to Original")
            move_up_action = menu.addAction("Move Table Up")
            move_down_action = menu.addAction("Move Table Down")

            action = menu.exec(self.file_list.mapToGlobal(pos))

            if action == rename_action:
                self.rename_table(item)
            elif action == delete_action:
                self.delete_table(item)
            elif action == rollback_action:
                self.rollback_table(item)
            elif action == move_up_action:
                self.move_table_up(item)
            elif action == move_down_action:
                self.move_table_down(item)

    def rename_table(self, item):
        old_name = item.text()
        new_name, ok = QInputDialog.getText(self, "Rename Table", "Enter new table name:", QLineEdit.EchoMode.Normal,
                                            old_name)
        if ok and new_name != old_name:
            self.tables[new_name] = self.tables.pop(old_name)
            item.setText(new_name)

    def delete_table(self, item):
        table_name = item.text()
        del self.tables[table_name]

        current_row = self.file_list.row(item)
        self.file_list.takeItem(current_row)

        if self.file_list.count() > 0:
            if current_row > 0:
                new_current_item = self.file_list.item(current_row - 1)
            else:
                new_current_item = self.file_list.item(0)

            self.file_list.setCurrentItem(new_current_item)
            self.show_table(new_current_item)
        else:
            self.table_view.setColumnCount(0)
            self.table_view.setRowCount(0)

    def rollback_table(self, item):
        table_name = item.text()
        table_revision = self.tables[table_name]
        table_revision.current_revision = 0
        table_revision.data = table_revision.revisions[0]
        self.populate_table(table_revision.data)

    def move_table_up(self, item):
        current_row = self.file_list.row(item)
        if current_row > 0:
            self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row - 1, item)
            self.file_list.setCurrentItem(item)
        else:
            QMessageBox.information(self, "Info", "The table is already at the top.")

    def move_table_down(self, item):
        current_row = self.file_list.row(item)
        if current_row < self.file_list.count() - 1:
            self.file_list.takeItem(current_row)
            self.file_list.insertItem(current_row + 1, item)
            self.file_list.setCurrentItem(item)
        else:
            QMessageBox.information(self, "Info", "The table is already at the bottom.")

    def rename_column(self, column_index):
        table_name = self.file_list.currentItem().text()
        table_revision = self.tables[table_name]
        data = table_revision.revisions[table_revision.current_revision].copy()
        old_name = data.columns[column_index]
        old_dtype = str(data.dtypes[column_index])
        new_name, ok = QInputDialog.getText(self, "Rename Column", "Enter new column name:", QLineEdit.EchoMode.Normal,
                                            old_name)
        if ok and new_name != old_name:
            data.rename(columns={old_name: new_name}, inplace=True)
            table_revision.add_revision(data)
            self.populate_table(data)
            self.table_view.horizontalHeaderItem(column_index).setText(f"{new_name} ({old_dtype})")

    def show_context_menu(self, pos):
        menu = QMenu(self)

        # Get the selected columns and rows
        selected_columns = self.table_view.selectionModel().selectedColumns()
        selected_rows = self.table_view.selectionModel().selectedRows()

        if selected_columns:
            # One or more columns are selected
            menu.addAction("Delete Selected Columns", self.delete_selected_columns)
            menu.addAction("Insert Column Left", self.insert_column_left)
            menu.addAction("Insert Column Right", self.insert_column_right)
        elif selected_rows:
            # One or more rows are selected
            menu.addAction("Delete Selected Rows", self.delete_selected_rows)
            menu.addAction("Insert Row Above", self.insert_row_above)
            menu.addAction("Insert Row Below", self.insert_row_below)
        else:
            # No selection or a single cell is selected
            menu.addAction("Delete Selected Rows", self.delete_selected_rows)
            menu.addAction("Delete Selected Columns", self.delete_selected_columns)
            menu.addSeparator()
            menu.addAction("Insert Row Above", self.insert_row_above)
            menu.addAction("Insert Row Below", self.insert_row_below)
            menu.addAction("Insert Column Left", self.insert_column_left)
            menu.addAction("Insert Column Right", self.insert_column_right)

        menu.exec(self.table_view.mapToGlobal(pos))

    def delete_selected_rows(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            rows_to_delete = set(index.row() for index in selected_indexes)
            data.drop(data.index[list(rows_to_delete)], inplace=True)
            table_revision.add_revision(data)
            self.populate_table(data)

    def delete_selected_columns(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            columns_to_delete = set(data.columns[index.column()] for index in selected_indexes)
            data.drop(columns=list(columns_to_delete), inplace=True)
            table_revision.add_revision(data)
            self.populate_table(data)

    def insert_row_above(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            current_row = min(index.row() for index in selected_indexes)
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            new_row = pd.DataFrame({column: "" for column in data.columns}, index=[current_row])
            data = pd.concat([data.iloc[:current_row], new_row, data.iloc[current_row:]], ignore_index=True)
            table_revision.add_revision(data)
            self.populate_table(data)

    def insert_row_below(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            current_row = max(index.row() for index in selected_indexes)
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            new_row = pd.DataFrame({column: "" for column in data.columns}, index=[current_row + 1])
            data = pd.concat([data.iloc[:current_row + 1], new_row, data.iloc[current_row + 1:]], ignore_index=True)
            table_revision.add_revision(data)
            self.populate_table(data)

    def insert_column_left(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            current_column = min(index.column() for index in selected_indexes)
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            new_column_name = f"New Column {current_column}"
            data.insert(current_column, new_column_name, "")
            table_revision.add_revision(data)
            self.populate_table(data)

    def insert_column_right(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            current_column = max(index.column() for index in selected_indexes)
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            new_column_name = f"New Column {current_column + 1}"
            data.insert(current_column + 1, new_column_name, "")
            table_revision.add_revision(data)
            self.populate_table(data)

    def show_table(self, item):
        table_name = item.text()
        table_revision = self.tables[table_name]
        self.populate_table(table_revision.data)
        self.file_list.setCurrentItem(item)

    def merge_tables(self):
        if len(self.tables) < 2:
            QMessageBox.warning(self, "Error", "At least two tables are required for merging.")
            return

        selected_table = self.file_list.currentItem().text()
        dialog = MergeDialog(self.tables, selected_table, parent=self)  # Pass self as the parent
        dialog.exec()

    def append_tables(self):
        if len(self.tables) < 2:
            QMessageBox.warning(self, "Error", "At least two tables are required for appending.")
            return

        selected_table = self.file_list.currentItem().text()
        dialog = AppendDialog(self.tables, selected_table, parent=self)  # Pass self as the parent
        dialog.exec()

    def pivot_table(self):
        if len(self.tables) == 0:
            QMessageBox.warning(self, "Error", "No tables available for pivoting.")
            return

        data = list(self.tables.values())[0]
        pivot_data = data.pivot_table(index=data.columns[0], values=data.columns[1], aggfunc="sum")
        self.populate_table(pivot_data)

    def undo_revision(self):
        if len(self.tables) < 1:
            QMessageBox.warning(self, "Error", "Nothing to undo.")
            return

        table_name = self.file_list.currentItem().text()
        table_revision = self.tables[table_name]
        status = table_revision.undo()
        if status != -1:
            self.populate_table(table_revision.data)
        else:
            QMessageBox.warning(self, "Error", "Nothing to undo.")

    def redo_revision(self):
        if len(self.tables) < 1:
            QMessageBox.warning(self, "Error", "Nothing to redo.")
            return

        table_name = self.file_list.currentItem().text()
        table_revision = self.tables[table_name]
        status = table_revision.redo()
        if status != -1:
            self.populate_table(table_revision.data)
        else:
            QMessageBox.warning(self, "Error", "Nothing to redo.")


def load_stylesheet() -> str:
    """
    Load the stylesheet from a given file path.

    :return: The stylesheet content as a string.
    """
    # Determine the full path to the icons folder
    icons_path = "images"
    try:
        with open("stylesheet.qss", "r") as file:
            return file.read().replace('{{ICON_PATH}}', str(icons_path).replace("\\", "/"))
    except IOError:
        print(
            f"Error opening stylesheet file: "
            f"stylesheet.qss")
        return ""


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setFont(QFont("Arial", 10))
    stylesheet = load_stylesheet()
    app.setStyleSheet(stylesheet)
    spreadsheet_app = SpreadsheetApp()
    spreadsheet_app.show()
    sys.exit(app.exec())
