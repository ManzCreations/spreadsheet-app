import os
import sys
import re

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QFont, QIcon, QColor
from PyQt6.QtWidgets import *
from openpyxl import load_workbook
from openpyxl.utils.dataframe import dataframe_to_rows


# TODO: What the hell am I doing with the loading window? Create a gif maybe?
# TODO: Update pivoting and un-pivoting
# TODO: Allow option to merge as new.
# TODO: Add a button for filtering that allows for whole word searching


class ExportDialog(QDialog):
    def __init__(self, tables, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Export Tables")
        self.setWindowIcon(QIcon("images/crm-icon-high-seas.png"))
        self.setGeometry(100, 100, 500, 400)

        self.tables = tables
        self.output_location = ""

        layout = QVBoxLayout()

        # Header
        header_label = QLabel("Export Tables")
        header_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        layout.addWidget(header_label)

        # Explanation
        explanation_label = QLabel(
            "Select the tables to export, choose the output location, and specify the file names and extensions.")
        explanation_label.setStyleSheet("font-size: 10pt; color: #888;")
        layout.addWidget(explanation_label)

        self.update_name_label = QLabel("Tables with red text were found in your output directory. "
                                        "Consider changing the name before exporting.")
        self.update_name_label.setStyleSheet("font-size: 10pt; color: #ff0000;")
        self.update_name_label.setVisible(False)
        layout.addWidget(self.update_name_label)

        if not tables:
            no_tables_label = QLabel("No tables available to export. Please add tables to the app first.")
            no_tables_label.setStyleSheet("font-size: 12pt; color: #888;")
            layout.addWidget(no_tables_label)
            self.export_button = QPushButton("Close")
            self.export_button.clicked.connect(self.close)
            layout.addWidget(self.export_button)
        else:
            # Table List
            self.table_list = QTableWidget()
            self.table_list.setColumnCount(4)
            self.table_list.setHorizontalHeaderLabels(["Table Name", "Spreadsheet Name", "Sheet Name", "Extension"])
            self.table_list.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
            self.table_list.setEditTriggers(QAbstractItemView.EditTrigger.AllEditTriggers)
            self.table_list.setToolTip("Select the tables to export.")
            for row, (table_name, table_revision) in enumerate(tables.items()):
                self.table_list.insertRow(row)
                self.table_list.setItem(row, 0, QTableWidgetItem(table_name))
                self.table_list.setItem(row, 1, QTableWidgetItem(table_revision.spreadsheet_name))
                self.table_list.setItem(row, 2, QTableWidgetItem(table_revision.sheet_name))
                self.table_list.setItem(row, 3, QTableWidgetItem(table_revision.extension))
            self.table_list.resizeColumnsToContents()
            self.table_list.horizontalHeader().setStretchLastSection(True)
            self.table_list.itemChanged.connect(self.update_table_data)
            layout.addWidget(self.table_list)

            # Output Location
            output_layout = QHBoxLayout()
            self.output_line_edit = QLineEdit()
            self.output_line_edit.setPlaceholderText("Output Location")
            self.output_line_edit.setToolTip("Choose the output location for the exported tables.")
            output_layout.addWidget(self.output_line_edit)
            browse_button = QPushButton("Browse")
            browse_button.clicked.connect(self.browse_output_location)
            output_layout.addWidget(browse_button)
            layout.addLayout(output_layout)

            # Export Button
            self.export_button = QPushButton("Export")
            self.export_button.setToolTip("Export the selected tables to the specified output location.")
            self.export_button.clicked.connect(self.export_selected_tables)
            layout.addWidget(self.export_button)

        self.setLayout(layout)

    def browse_output_location(self):
        options = QFileDialog.Option.DontResolveSymlinks | QFileDialog.Option.ShowDirsOnly
        directory = QFileDialog.getExistingDirectory(self, "Select Output Location", self.output_location,
                                                     options=options)
        if directory:
            self.output_location = directory
            self.output_line_edit.setText(directory)
            self.check_existing_files()

    def check_existing_files(self):
        if self.output_location:
            existing_files = os.listdir(self.output_location)
            show_update_label = False
            for row in range(self.table_list.rowCount()):
                spreadsheet_name = self.table_list.item(row, 1).text()
                extension = self.table_list.item(row, 3).text()
                file_name = f"{spreadsheet_name}{extension}"
                if file_name in existing_files:
                    for col in range(4):
                        self.table_list.item(row, col).setForeground(QColor(255, 0, 0))
                    self.table_list.item(row, 0).setToolTip(
                        "A file with the same name already exists in the output location.")
                    show_update_label = True
                else:
                    for col in range(4):
                        self.table_list.item(row, col).setForeground(QColor(255, 255, 255))
                    self.table_list.item(row, 0).setToolTip("")
            self.update_name_label.setVisible(show_update_label)

    def update_table_data(self, item):
        row = item.row()
        col = item.column()
        old_table_name = self.table_list.item(row, 0).text()
        table_revision = self.tables[old_table_name]

        if col == 1:  # Spreadsheet Name
            table_revision.spreadsheet_name = item.text()
        elif col == 2:  # Sheet Name
            table_revision.sheet_name = item.text()
        elif col == 3:  # Extension
            table_revision.extension = item.text()

        # Update Table Name
        new_table_name = f"{table_revision.spreadsheet_name} - {table_revision.sheet_name}"
        table_revision.table_name = new_table_name

        # Update the dictionary key
        self.tables[new_table_name] = self.tables.pop(old_table_name)

        # Update the table list item
        self.table_list.item(row, 0).setText(new_table_name)

        # Check if the table exists in the output folder
        if self.output_location:
            file_name = f"{table_revision.spreadsheet_name}{table_revision.extension}"
            file_path = os.path.join(self.output_location, file_name)
            if os.path.exists(file_path):
                for col in range(4):
                    self.table_list.item(row, col).setForeground(QColor(255, 0, 0))
                self.table_list.item(row, 0).setToolTip(
                    "A file with the same name already exists in the output location.")
                self.update_name_label.setVisible(True)
            else:
                for col in range(4):
                    self.table_list.item(row, col).setForeground(QColor(255, 255, 255))
                self.table_list.item(row, 0).setToolTip("")

    def export_selected_tables(self):
        selected_rows = self.table_list.selectionModel().selectedRows()
        if not selected_rows:
            QMessageBox.warning(self, "No Tables Selected", "Please select at least one table to export.")
            return

        if not self.output_location:
            QMessageBox.warning(self, "No Output Location", "Please specify an output location.")
            return

        open_files = []
        file_data = {}
        for row in [index.row() for index in selected_rows]:
            table_name = self.table_list.item(row, 0).text()
            spreadsheet_name = self.table_list.item(row, 1).text()
            sheet_name = self.table_list.item(row, 2).text()
            extension = self.table_list.item(row, 3).text()
            file_name = f"{spreadsheet_name}{extension}"
            file_path = os.path.join(self.output_location, file_name)
            if os.path.exists(file_path):
                if os.path.isfile(file_path) and sys.platform == 'win32':
                    try:
                        os.rename(file_path, file_path)
                    except OSError as e:
                        if e.errno == 13:
                            open_files.append(file_name)

            if file_name not in file_data:
                file_data[file_name] = {}
            file_data[file_name][sheet_name] = self.tables[table_name]

        if open_files:
            reply = QMessageBox.question(self, "Files Open",
                                         "The following files are currently open:\n\n"
                                         + "\n".join(open_files) +
                                         "\n\nPlease save any unsaved changes in the open files. "
                                         "This program will close the files and overwrite them. "
                                         "Do you want to proceed?",
                                         QMessageBox.StandardButton.Cancel | QMessageBox.StandardButton.Ok,
                                         QMessageBox.StandardButton.Cancel)
            if reply == QMessageBox.StandardButton.Cancel:
                return
            else:
                for file_name in open_files:
                    file_path = os.path.join(self.output_location, file_name)
                    os.system(f'taskkill /F /IM "{os.path.basename(file_path)}"')

        for file_name, sheet_data in file_data.items():
            extension = os.path.splitext(file_name)[1]
            file_path = os.path.join(self.output_location, file_name)

            if extension == ".csv":
                if len(sheet_data) > 1:
                    QMessageBox.warning(self, "Multiple Sheets",
                                        f"The file '{file_name}' contains multiple sheets. "
                                        "Only the first sheet will be exported as CSV.")
                sheet_name = next(iter(sheet_data))
                sheet_data[sheet_name].revisions[sheet_data[sheet_name].current_revision].to_csv(file_path, index=False)
            elif extension in [".xlsx", ".xls", ".xlsm"]:
                if os.path.exists(file_path):
                    # Create new file with _transformed appended to the name
                    file_name, extension = os.path.splitext(file_path)
                    with pd.ExcelWriter(f"{file_name}_transformed{extension}", engine='openpyxl') as writer:
                        for sheet_name, table_revision in sheet_data.items():
                            table_revision.revisions[table_revision.current_revision].to_excel(writer,
                                                                                               sheet_name=sheet_name,
                                                                                               index=False)
                else:
                    # Create new file
                    with pd.ExcelWriter(file_path, engine='openpyxl') as writer:
                        for sheet_name, table_revision in sheet_data.items():
                            table_revision.revisions[table_revision.current_revision].to_excel(writer,
                                                                                               sheet_name=sheet_name,
                                                                                               index=False)
            elif extension == ".txt":
                if len(sheet_data) > 1:
                    QMessageBox.warning(self, "Multiple Sheets",
                                        f"The file '{file_name}' contains multiple sheets. "
                                        "Only the first sheet will be exported as TXT.")
                sheet_name = next(iter(sheet_data))
                sheet_data[sheet_name].revisions[sheet_data[sheet_name].current_revision].to_csv(file_path, sep="\t",
                                                                                                 index=False)
            else:
                QMessageBox.warning(self, "Unsupported Extension",
                                    f"The extension '{extension}' is not supported for export.")
                continue

        QMessageBox.information(self, "Export Completed", "The selected tables have been exported successfully.")
        self.close()


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

        # Header
        header_label = QLabel("Merge Tables")
        header_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        layout.addWidget(header_label)

        # Explanation
        explanation_label = QLabel("Select two tables to merge and specify the join type.")
        explanation_label.setStyleSheet("font-size: 10pt; color: #888;")
        layout.addWidget(explanation_label)

        # Dropdowns
        dropdown_layout = QHBoxLayout()
        self.table1_dropdown = QComboBox()
        self.table1_dropdown.addItems(list(tables.keys()))
        self.table1_dropdown.setCurrentText(selected_table)
        self.table1_dropdown.setToolTip("Select the first table to merge.")
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
        self.table2_dropdown.setToolTip("Select the second table to merge.")
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
        self.join_dropdown.setToolTip("Select the type of join to perform.")
        join_layout.addWidget(QLabel("Join Type:"))
        join_layout.addWidget(self.join_dropdown)
        layout.addLayout(join_layout)

        # Button
        self.merge_button = QPushButton("Merge")
        self.merge_button.setToolTip("Perform the merge operation.")
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
        table_view.setToolTip("Select a column to merge on.")

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

        # Header
        header_label = QLabel("Append Tables")
        header_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        layout.addWidget(header_label)

        # Explanation
        explanation_label = QLabel("Select two tables to append and specify the append direction.")
        explanation_label.setStyleSheet("font-size: 10pt; color: #888;")
        layout.addWidget(explanation_label)

        # Dropdowns
        dropdown_layout = QHBoxLayout()
        self.table1_dropdown = QComboBox()
        self.table1_dropdown.addItems(list(tables.keys()))
        self.table1_dropdown.setCurrentText(selected_table)
        self.table1_dropdown.setToolTip("Select the first table to append.")
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
        self.table2_dropdown.setToolTip("Select the second table to append.")
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
        self.direction_dropdown.setToolTip("Select the direction to append the tables.")
        direction_layout.addWidget(QLabel("Append Direction:"))
        direction_layout.addWidget(self.direction_dropdown)
        layout.addLayout(direction_layout)

        # Button
        self.append_button = QPushButton("Append")
        self.append_button.setToolTip("Perform the append operation.")
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
        self.spreadsheet_name = ""
        self.sheet_name = ""
        self.extension = ""

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
    """
    Represents the main window of the Spreadsheet Application. It includes a file list
    for managing tables, a table view for displaying data, and various tools for data
    manipulation and visualization.

    Functions:
    - __init__: Initializes the main window with a layout, headers, file list, table view, and tools.
    - init_ui: Sets up the user interface components and layouts.
    - init_menu: Creates the menu bar with file and operations menus.
    - add_table: Adds a new table to the application from an Excel or CSV file.
    - populate_table: Populates the table view with data from the selected table.
    - show_file_context_menu: Displays a context menu for file operations.
    - rename_table: Renames the selected table.
    - delete_table: Deletes the selected table.
    - rollback_table: Rolls back the selected table to its original state.
    - move_table_up: Moves the selected table up in the file list.
    - move_table_down: Moves the selected table down in the file list.
    - rename_column: Renames the selected column in the table view.
    - show_context_menu: Displays a context menu for table operations.
    - delete_selected_rows: Deletes the selected rows from the table.
    - delete_selected_columns: Deletes the selected columns from the table.
    - insert_row_above: Inserts a new row above the selected row.
    - insert_row_below: Inserts a new row below the selected row.
    - insert_column_left: Inserts a new column to the left of the selected column.
    - insert_column_right: Inserts a new column to the right of the selected column.
    - show_table: Displays the selected table in the table view.
    - merge_tables: Opens a dialog to merge two tables.
    - append_tables: Opens a dialog to append tables.
    - pivot_table: Performs a pivot operation on the selected table.
    - undo_revision: Undoes the last revision made to the selected table.
    - redo_revision: Redoes the last undone revision made to the selected table.
    - updateButtonStyle: Updates the style of the filter buttons based on their state.
    - filterTable: Filters the table based on the entered text and selected column.
    - sort_column: Sorts the selected column in the table view.
    """

    def __init__(self):
        super().__init__()
        self.setWindowTitle("Spreadsheet Application")
        self.setWindowIcon(QIcon("images/crm-icon-high-seas.png"))
        self.setGeometry(100, 100, 800, 600)

        self.tables = {}
        self.pressed_keys = set()
        self.current_showing_table = None

        self.init_ui()

    def init_ui(self):
        main_layout = QVBoxLayout()

        # Header
        header_layout = QHBoxLayout()
        title_label = QLabel("Spreadsheet Application")
        title_label.setStyleSheet("font-size: 18pt; font-weight: bold;")
        header_layout.addWidget(title_label)
        header_layout.addStretch()
        main_layout.addLayout(header_layout)

        # Explanation
        explanation_label = QLabel("Use the File menu to load/export tables into the application. "
                                   "Use the Operations menu to do actions like merge, append, pivot, or un-pivot.")
        explanation_label.setStyleSheet("font-size: 10pt; color: #888;")
        main_layout.addWidget(explanation_label)

        # Revision buttons
        revision_button_layout = QHBoxLayout()
        undo_button = QPushButton("Undo")
        undo_button.setToolTip("Undo the last revision made to the selected table.")
        undo_button.clicked.connect(self.undo_revision)
        revision_button_layout.addWidget(undo_button)

        redo_button = QPushButton("Redo")
        redo_button.setToolTip("Redo the last undone revision made to the selected table.")
        redo_button.clicked.connect(self.redo_revision)
        revision_button_layout.addWidget(redo_button)

        revision_button_layout.addStretch()
        main_layout.addLayout(revision_button_layout)

        file_table_layout = QHBoxLayout()

        # File List
        file_list_layout = QVBoxLayout()
        file_list_label = QLabel("Tables")
        file_list_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        file_list_layout.addWidget(file_list_label)

        self.file_list = QListWidget()
        self.file_list.setMaximumWidth(200)  # Set a maximum width for the file list
        self.file_list.setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.file_list.setToolTip("Double-click a table to show its data in the table view. "
                                  "Right-click for more options.")
        self.file_list.itemDoubleClicked.connect(self.show_table)
        self.file_list.itemChanged.connect(self.rename_table)
        self.file_list.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.file_list.customContextMenuRequested.connect(self.show_file_context_menu)
        file_list_layout.addWidget(self.file_list)
        file_table_layout.addLayout(file_list_layout)

        # Table View
        table_view_layout = QVBoxLayout()
        table_view_label = QLabel("Data")
        table_view_label.setStyleSheet("font-size: 12pt; font-weight: bold;")
        table_view_layout.addWidget(table_view_label)

        self.table_view = QTableWidget()
        self.table_view.setColumnCount(0)
        self.table_view.setRowCount(0)
        self.table_view.horizontalHeader().setSectionResizeMode(QHeaderView.ResizeMode.ResizeToContents)
        self.table_view.setHorizontalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table_view.setVerticalScrollMode(QAbstractItemView.ScrollMode.ScrollPerPixel)
        self.table_view.setContextMenuPolicy(Qt.ContextMenuPolicy.CustomContextMenu)
        self.table_view.setToolTip("Double-click a column header to rename it. Right-click for more options.")
        self.table_view.customContextMenuRequested.connect(self.show_context_menu)
        self.table_view.horizontalHeader().setEditTriggers(QAbstractItemView.EditTrigger.DoubleClicked)
        self.table_view.horizontalHeader().sectionDoubleClicked.connect(self.rename_column)
        table_view_layout.addWidget(self.table_view)
        file_table_layout.addLayout(table_view_layout)

        main_layout.addLayout(file_table_layout)

        # Filter Row
        self.filterRowLayout = QHBoxLayout()
        self.filterRowLayout.setSpacing(10)

        # Filter and Sorting Section Header
        self.filterSectionHeader = QLabel(
            "Filter Table View (Select a column, then type to filter. Double-click a column header to sort.)")
        self.filterSectionHeader.setStyleSheet("font-size: 10pt; margin-top: 10px;")
        main_layout.addWidget(self.filterSectionHeader)

        # Text Editor for Filtering
        self.filterTextEditor = QLineEdit()
        self.filterTextEditor.setPlaceholderText("Filter...")
        self.filterTextEditor.textChanged.connect(self.filterTable)
        self.filterTextEditor.setToolTip("Enter text to filter the data in the selected column.")
        self.filterRowLayout.addWidget(self.filterTextEditor)

        # Starts With Button
        self.swButton = QPushButton("Sw")
        self.swButton.setCheckable(True)
        self.swButton.clicked.connect(self.updateButtonStyle)
        self.swButton.setToolTip(
            "Starts With: Select this button if you want to filter only to words that start with the letters/numbers in your filter criteria.")
        self.filterRowLayout.addWidget(self.swButton)

        # Match Case Button
        self.ccButton = QPushButton("Cc")
        self.ccButton.setCheckable(True)
        self.ccButton.clicked.connect(self.updateButtonStyle)
        self.ccButton.setToolTip(
            "Match Case: Select this button if you want to match the case that you have in the filter criteria.")
        self.filterRowLayout.addWidget(self.ccButton)

        main_layout.addLayout(self.filterRowLayout)

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

        export_action = QAction("Export Table(s)", self)
        export_action.triggered.connect(self.export_tables)
        file_menu.addAction(export_action)

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

    def updateButtonStyle(self):
        """
        Updates the style of the Sw and Cc buttons based on their checked state.
        """
        sender = self.sender()
        if sender.isChecked():
            sender.setStyleSheet("background-color: #4d8d9c;")  # Darker color when checked
        else:
            sender.setStyleSheet("")  # Revert to default stylesheet

    def filterTable(self):
        """
        Filters the table rows based on the text entered in the filterTextEditor and the selected column.

        The method performs a case-insensitive comparison of the filter text with the data in the selected column.
        It hides rows that do not contain the filter text in the selected column. The comparison takes into
        account the data type of the column (numeric, date, or string) for appropriate formatting and comparison.
        """
        filter_text = self.filterTextEditor.text()
        column_index = self.table_view.currentColumn()

        if column_index == -1:
            return  # Exit if no column is selected

        for row in range(self.table_view.rowCount()):
            item = self.table_view.item(row, column_index)
            if item is None:
                self.table_view.setRowHidden(row, True)
                continue

            cell_value = item.text()

            # Adjust for case sensitivity based on Cc button
            if not self.ccButton.isChecked():
                filter_text = filter_text.lower()
                cell_value = cell_value.lower()

            # Adjust for "Starts With" functionality based on Sw button
            if self.swButton.isChecked():
                self.table_view.setRowHidden(row, not cell_value.startswith(filter_text))
            else:
                self.table_view.setRowHidden(row, filter_text not in cell_value)

    def add_table(self):
        options = QFileDialog.Option.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, "Add Table", "",
                                                   "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv);;Text files (*.txt)",
                                                   options=options)
        if file_path:
            file_name = os.path.basename(file_path)
            file_name_without_ext, extension = os.path.splitext(file_name)
            if extension == ".txt":
                data = pd.read_csv(file_path, sep="\t")
                sheet_name = "Sheet1"
            elif extension == ".csv":
                data = pd.read_csv(file_path)
                sheet_name = "Sheet1"
            else:
                excel_file = pd.ExcelFile(file_path)
                sheet_names = excel_file.sheet_names
                for sheet_name in sheet_names:
                    data = excel_file.parse(sheet_name)
                    table_name = f"{file_name_without_ext} - {sheet_name}"

                    if table_name in self.tables:
                        pattern = rf"{table_name}\s*\((\d+)\)"
                        max_number = 0
                        for existing_table_name in self.tables:
                            match = re.match(pattern, existing_table_name)
                            if match:
                                number = int(match.group(1))
                                max_number = max(max_number, number)
                        # Increment the number for the new table name
                        if max_number == 0:
                            table_name = f"{table_name} (1)"
                        else:
                            table_name = f"{table_name} ({max_number + 1})"

                    self.tables[table_name] = TableRevision(data)
                    self.tables[table_name].spreadsheet_name = file_name_without_ext
                    self.tables[table_name].sheet_name = sheet_name
                    self.tables[table_name].extension = extension if extension else ".xlsx"
                    item = QListWidgetItem(table_name)
                    self.file_list.addItem(item)
                    self.file_list.setCurrentItem(item)
                self.show_table(self.file_list.currentItem())
                return
            table_name = file_name_without_ext
            self.tables[table_name] = TableRevision(data)
            self.tables[table_name].spreadsheet_name = file_name_without_ext
            self.tables[table_name].sheet_name = sheet_name
            self.tables[table_name].extension = extension if extension else ".xlsx"
            self.file_list.addItem(table_name)
        self.show_table(self.file_list.currentItem())

    def export_tables(self):
        if not self.tables:
            QMessageBox.warning(self, "No Tables", "No tables available to export. Please add tables to the app first.")
            return

        dialog = ExportDialog(self.tables, parent=self)
        dialog.exec()

    def populate_table(self, data):
        self.table_view.clear()
        self.table_view.setColumnCount(len(data.columns))
        self.table_view.setRowCount(len(data))
        self.table_view.setHorizontalHeaderLabels([f"{col} ({dtype})" for col, dtype in zip(data.columns, data.dtypes)])

        for i in range(len(data)):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                self.table_view.setItem(i, j, item)

    def show_table(self, item):
        table_name = item.text()
        table_revision = self.tables[table_name]
        if self.table_view.horizontalHeader().count() > 0 and table_name == self.current_showing_table:
            return  # Table is already displayed, no need to repopulate
        self.populate_table(table_revision.data)
        self.file_list.setCurrentItem(item)
        self.current_showing_table = table_name

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
            menu.addAction("Sort Column (ascending)", self.sort_column_ascending)
            menu.addAction("Sort Column (descending)", self.sort_column_descending)
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
            menu.addSeparator()
            menu.addAction("Sort Column (ascending)", self.sort_column_ascending)
            menu.addAction("Sort Column (descending)", self.sort_column_descending)

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

    def sort_column_ascending(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            column_index = selected_indexes[0].column()
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            column_name = data.columns[column_index]

            data.sort_values(by=column_name, ascending=True, inplace=True)

            table_revision.add_revision(data)
            self.populate_table(data)

    def sort_column_descending(self):
        selected_indexes = self.table_view.selectedIndexes()
        if selected_indexes:
            column_index = selected_indexes[0].column()
            table_name = self.file_list.currentItem().text()
            table_revision = self.tables[table_name]
            data = table_revision.revisions[table_revision.current_revision].copy()
            column_name = data.columns[column_index]

            data.sort_values(by=column_name, ascending=False, inplace=True)

            table_revision.add_revision(data)
            self.populate_table(data)


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
