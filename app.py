import os
import re
import sys
from typing import Dict, Optional

import pandas as pd
from PyQt6.QtCore import Qt
from PyQt6.QtGui import QAction, QFont, QIcon, QColor, QPixmap
from PyQt6.QtWidgets import *


# TODO: Make code device and OS agnostic


def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and for PyInstaller """
    try:
        # PyInstaller creates a temp folder and stores path in _MEIPASS
        base_path = sys._MEIPASS2
    except Exception:
        base_path = os.path.abspath(".")

    return os.path.join(base_path, relative_path)


class LoadingDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Loading...")
        self.setFixedSize(150, 150)
        self.setWindowFlag(Qt.WindowType.FramelessWindowHint)
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        # Set the background color of the dialog to transparent
        self.setStyleSheet("background-color: rgba(0, 0, 0, 0);")

        layout = QVBoxLayout(self)
        layout.setContentsMargins(0, 0, 0, 0)

        # Create a QLabel to hold the loading image
        self.label = QLabel(self)
        self.label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        layout.addWidget(self.label)

        # Load the image using QPixmap and resize it
        pixmap = QPixmap(resource_path(os.path.join("assets", "images", "loading_with_text.png")))
        pixmap = pixmap.scaled(128, 128, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation)
        self.label.setPixmap(pixmap)

    def keyPressEvent(self, event):
        # Prevent the dialog from being closed by pressing Esc key
        if event.key() != Qt.Key.Key_Escape:
            super().keyPressEvent(event)


class ExportDialog(QDialog):
    """
    A dialog for exporting tables from the Spreadsheet Application.

    The ExportDialog allows the user to select tables to export, choose the output location,
    and specify file names and extensions. It provides options to update table names and
    handles the export process for different file formats.

    Functions:
    - __init__: Initializes the ExportDialog with the necessary components and layout.
    - browse_output_location: Opens a file dialog for the user to select the output location.
    - check_existing_files: Checks if the selected tables already exist in the output location.
    - update_table_data: Updates the table data when the user modifies the table list.
    - export_selected_tables: Exports the selected tables to the specified output location.
    """

    def __init__(self, tables: Dict[str, 'TableRevision'], parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Export Tables")
        self.setWindowIcon(QIcon(resource_path(os.path.join("assets", "images", "crm-icon-high-seas.png"))))
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

        self.parent().loading_dialog.show()

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
                self.parent().loading_dialog.hide()
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
        self.parent().loading_dialog.hide()
        self.close()


class PivotDialog(QDialog):
    """
    A dialog for selecting pivot options in the Spreadsheet Application.

    The PivotDialog allows the user to choose the values column for pivoting a selected column.
    It displays the selected column name and provides a dropdown menu to select the values column.

    Functions:
    - __init__: Initializes the PivotDialog with the necessary components and layout.
    - get_values_column: Retrieves the selected values column from the dropdown menu.
    """

    def __init__(self, data: pd.DataFrame, selected_column: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.setWindowTitle("Pivot Table")
        self.setGeometry(100, 100, 400, 200)

        layout = QVBoxLayout()

        # Header
        header_label = QLabel("Pivot Table")
        header_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        layout.addWidget(header_label)

        # Explanation
        explanation_label = QLabel(f"Pivoting the selected column: {selected_column}")
        explanation_label.setWordWrap(True)
        explanation_label.setToolTip("The selected column will be used as the new column headers.")
        layout.addWidget(explanation_label)

        # Values Column Dropdown
        values_label = QLabel("Select the values column:")
        values_label.setToolTip("Choose the column that will provide the values for the pivoted cells.")
        layout.addWidget(values_label)

        self.values_dropdown = QComboBox()
        self.values_dropdown.addItems(data.columns)
        layout.addWidget(self.values_dropdown)

        # Accept Button
        self.accept_button = QPushButton("Accept")
        self.accept_button.clicked.connect(self.accept)
        layout.addWidget(self.accept_button)

        self.setLayout(layout)

    def get_values_column(self):
        return self.values_dropdown.currentText()


class MergeDialog(QDialog):
    """
    A dialog for merging tables in the Spreadsheet Application.

    The MergeDialog allows the user to select two tables to merge, specify the join type,
    and choose the columns to merge on. It provides a preview of the selected tables and
    displays information about different join types.

    Functions:
    - __init__: Initializes the MergeDialog with the necessary components and layout.
    - create_table_view: Creates a table view widget for displaying a preview of the selected table.
    - update_table1_view: Updates the table view for the first selected table.
    - update_table2_view: Updates the table view for the second selected table.
    - update_table_view: Updates the table view with the given table revision data.
    - update_selected_column: Updates the selected column when the user selects a column in the table views.
    - show_join_info: Displays information about different join types in a scrollable dialog.
    - accept: Performs the merge operation when the user accepts the dialog.
    """

    def __init__(self, tables: Dict[str, 'TableRevision'], selected_table: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.merged_data = None
        self.setWindowTitle("Merge Tables")
        self.setWindowIcon(QIcon(resource_path(os.path.join("assets", "images", "crm-icon-high-seas.png"))))
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

        # Info Button
        info_button = QPushButton()
        info_button.setIcon(QIcon(resource_path(os.path.join("assets", "images", "iconmonstr-info-9-240.png"))))
        info_button.clicked.connect(self.show_join_info)
        join_layout.addWidget(info_button)

        layout.addLayout(join_layout)

        # Button
        self.merge_button = QPushButton("Merge")
        self.merge_button.setToolTip("Perform the merge operation.")
        self.merge_button.clicked.connect(self.accept)
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

    def show_join_info(self):
        info_dialog = QDialog(self)
        info_dialog.setWindowTitle("Join Type Information")
        info_dialog.setGeometry(100, 100, 600, 400)

        scroll_area = QScrollArea(info_dialog)
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget(scroll_area)
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_area.setWidget(scroll_content)

        main_layout = QVBoxLayout(info_dialog)
        main_layout.addWidget(scroll_area)

        # Introduction
        intro_label = QLabel("What Are Joins?")
        intro_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        scroll_layout.addWidget(intro_label)

        intro_text = QLabel(
            "Joins allow you to fetch data that is scattered across multiple tables in a database. "
            "They enable you to combine rows from different tables based on a related column between them.")
        intro_text.setWordWrap(True)
        scroll_layout.addWidget(intro_text)

        # Inner Join
        inner_join_label = QLabel("Inner Join")
        inner_join_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        scroll_layout.addWidget(inner_join_label)

        inner_join_text = QLabel("An inner join returns only the rows that have matching values in both tables.")
        inner_join_text.setWordWrap(True)
        scroll_layout.addWidget(inner_join_text)

        inner_join_example = QLabel("Example: Table1 INNER JOIN Table2 ON Table1.ID = Table2.ID")
        inner_join_example.setWordWrap(True)
        scroll_layout.addWidget(inner_join_example)

        inner_join_image = QLabel()
        inner_join_image.setPixmap(QPixmap(resource_path(os.path.join("assets", "images", "inner_join.png"))))
        scroll_layout.addWidget(inner_join_image)

        # Left Join
        left_join_label = QLabel("Left Join")
        left_join_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        scroll_layout.addWidget(left_join_label)

        left_join_text = QLabel(
            "A left join returns all the rows from the left table and the matching rows from the right table. "
            "If there is no match, NULL values are returned for the right table.")
        left_join_text.setWordWrap(True)
        scroll_layout.addWidget(left_join_text)

        left_join_example = QLabel("Example: Table1 LEFT JOIN Table2 ON Table1.ID = Table2.ID")
        left_join_example.setWordWrap(True)
        scroll_layout.addWidget(left_join_example)

        left_join_image = QLabel()
        left_join_image.setPixmap(QPixmap(resource_path(os.path.join("assets", "images", "left_join.png"))))
        scroll_layout.addWidget(left_join_image)

        # Right Join
        right_join_label = QLabel("Right Join")
        right_join_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        scroll_layout.addWidget(right_join_label)

        right_join_text = QLabel(
            "A right join returns all the rows from the right table and the matching rows from the left table. "
            "If there is no match, NULL values are returned for the left table.")
        right_join_text.setWordWrap(True)
        scroll_layout.addWidget(right_join_text)

        right_join_example = QLabel("Example: Table1 RIGHT JOIN Table2 ON Table1.ID = Table2.ID")
        right_join_example.setWordWrap(True)
        scroll_layout.addWidget(right_join_example)

        right_join_image = QLabel()
        right_join_image.setPixmap(QPixmap(resource_path(os.path.join("assets", "images", "right_join.jpg"))))
        scroll_layout.addWidget(right_join_image)

        # Outer Join
        outer_join_label = QLabel("Outer Join")
        outer_join_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        scroll_layout.addWidget(outer_join_label)

        outer_join_text = QLabel(
            "An outer join returns all the rows from both tables, including the non-matching rows. "
            "NULL values are used for the missing values.")
        outer_join_text.setWordWrap(True)
        scroll_layout.addWidget(outer_join_text)

        outer_join_example = QLabel("Example: Table1 FULL OUTER JOIN Table2 ON Table1.ID = Table2.ID")
        outer_join_example.setWordWrap(True)
        scroll_layout.addWidget(outer_join_example)

        outer_join_image = QLabel()
        outer_join_image.setPixmap(QPixmap(resource_path(os.path.join("assets", "images", "outer_join.png"))))
        scroll_layout.addWidget(outer_join_image)

        info_dialog.exec()

    def accept(self):
        table1_name = self.table1_dropdown.currentText()
        table2_name = self.table2_dropdown.currentText()
        table1_revision = self.tables[table1_name]
        table2_revision = self.tables[table2_name]

        table1_data = table1_revision.revisions[table1_revision.current_revision]
        table2_data = table2_revision.revisions[table2_revision.current_revision]

        if self.selected_column1 is None or self.selected_column2 is None:
            QMessageBox.warning(self, "Error", "Please select a column from each table.")
            return

        self.parent().loading_dialog.show()

        merge_column1 = table1_data.columns[self.selected_column1]
        merge_column2 = table2_data.columns[self.selected_column2]

        join_type = self.join_dropdown.currentText().lower().split(" ")[0]

        self.merged_data = pd.merge(table1_data, table2_data, left_on=merge_column1, right_on=merge_column2,
                                    how=join_type)
        self.parent().loading_dialog.hide()
        super().accept()


class AppendDialog(QDialog):
    """
    A dialog for appending tables in the Spreadsheet Application.

    The AppendDialog allows the user to select two tables to append, specify the append direction,
    and choose the columns to append. It provides a preview of the selected tables and
    displays information about different append directions.

    Functions:
    - __init__: Initializes the AppendDialog with the necessary components and layout.
    - create_table_view: Creates a table view widget for displaying a preview of the selected table.
    - update_table1_view: Updates the table view for the first selected table.
    - update_table2_view: Updates the table view for the second selected table.
    - update_table_view: Updates the table view with the given table revision data.
    - update_selected_column: Updates the selected column when the user selects a column in the table views.
    - show_append_info: Displays information about different append directions in a scrollable dialog.
    - accept: Performs the append operation when the user accepts the dialog.
    """

    def __init__(self, tables: Dict[str, 'TableRevision'], selected_table: str, parent: Optional[QWidget] = None):
        super().__init__(parent)
        self.appended_data = None
        self.setWindowTitle("Append Tables")
        self.setWindowIcon(QIcon(resource_path(os.path.join("assets", "images", "crm-icon-high-seas.png"))))
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

        # Info Button
        info_button = QPushButton()
        info_button.setIcon(QIcon(resource_path(os.path.join("assets", "images", "iconmonstr-info-9-240.png"))))
        info_button.clicked.connect(self.show_append_info)
        direction_layout.addWidget(info_button)

        layout.addLayout(direction_layout)

        # Button
        self.append_button = QPushButton("Append")
        self.append_button.setToolTip("Perform the append operation.")
        self.append_button.clicked.connect(self.accept)
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

    def show_append_info(self):
        info_dialog = QDialog(self)
        info_dialog.setWindowTitle("Append Direction Information")
        info_dialog.setGeometry(100, 100, 600, 400)

        scroll_area = QScrollArea(info_dialog)
        scroll_area.setWidgetResizable(True)
        scroll_content = QWidget(scroll_area)
        scroll_layout = QVBoxLayout(scroll_content)
        scroll_area.setWidget(scroll_content)

        main_layout = QVBoxLayout(info_dialog)
        main_layout.addWidget(scroll_area)

        # Introduction
        intro_label = QLabel("Append Directions")
        intro_label.setStyleSheet("font-size: 18px; font-weight: bold;")
        scroll_layout.addWidget(intro_label)

        intro_text = QLabel(
            "Append allows you to combine data from multiple tables by stacking them either "
            "vertically or horizontally.")
        intro_text.setWordWrap(True)
        scroll_layout.addWidget(intro_text)

        # Vertical Append
        vertical_append_label = QLabel("Vertical Append")
        vertical_append_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        scroll_layout.addWidget(vertical_append_label)

        vertical_append_text = QLabel(
            "A vertical append stacks the rows of the tables on top of each other. "
            "The resulting table will have the same columns as the input tables.")
        vertical_append_text.setWordWrap(True)
        scroll_layout.addWidget(vertical_append_text)

        vertical_append_example = QLabel("Example: Table1 UNION Table2")
        vertical_append_example.setWordWrap(True)
        scroll_layout.addWidget(vertical_append_example)

        # Horizontal Append
        horizontal_append_label = QLabel("Horizontal Append")
        horizontal_append_label.setStyleSheet("font-size: 16px; font-weight: bold;")
        scroll_layout.addWidget(horizontal_append_label)

        horizontal_append_text = QLabel(
            "A horizontal append combines the columns of the tables side by side. "
            "The resulting table will have the same number of rows as the input tables.")
        horizontal_append_text.setWordWrap(True)
        scroll_layout.addWidget(horizontal_append_text)

        horizontal_append_example = QLabel("Example: Table1 JOIN Table2")
        horizontal_append_example.setWordWrap(True)
        scroll_layout.addWidget(horizontal_append_example)

        append_image = QLabel()
        append_image.setPixmap(QPixmap(resource_path(os.path.join("assets", "images", "appends.png"))))
        scroll_layout.addWidget(append_image)

        info_dialog.exec()

    def accept(self):
        table1_name = self.table1_dropdown.currentText()
        table2_name = self.table2_dropdown.currentText()
        table1_revision = self.tables[table1_name]
        table2_revision = self.tables[table2_name]

        table1_data = table1_revision.revisions[table1_revision.current_revision]
        table2_data = table2_revision.revisions[table2_revision.current_revision]

        append_direction = self.direction_dropdown.currentText().lower()

        self.parent().loading_dialog.show()
        if append_direction == "vertically":
            self.appended_data = pd.concat([table1_data, table2_data], ignore_index=True)
        else:
            self.appended_data = pd.concat([table1_data, table2_data], axis=1)
        self.parent().loading_dialog.hide()

        super().accept()


class TableRevision:
    """
    Represents a revision of a table in the Spreadsheet Application.

    The TableRevision class stores the data, revisions, and current revision of a table.
    It provides methods to add a new revision, undo changes, and redo changes.

    Functions:
    - __init__: Initializes the TableRevision with the given data.
    - add_revision: Adds a new revision to the table.
    - undo: Undoes the last revision made to the table.
    - redo: Redoes the last undone revision made to the table.
    """

    def __init__(self, data: pd.DataFrame):
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
        self.wButton = None
        self.ccButton = None
        self.swButton = None
        self.filterTextEditor = None
        self.filterSectionHeader = None
        self.filterRowLayout = None
        self.table_view = None
        self.file_list = None
        self.setWindowTitle("Spreadsheet Application")
        self.setWindowIcon(QIcon(resource_path(os.path.join("assets", "images", "crm-icon-high-seas.png"))))
        self.setGeometry(100, 100, 800, 600)

        self.tables = {}
        self.pressed_keys = set()
        self.current_showing_table = None

        self.loading_dialog = LoadingDialog(self)

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
            "Starts With: Select this button if you want to filter only to words that start with "
            "the letters/numbers in your filter criteria.")
        self.filterRowLayout.addWidget(self.swButton)

        # Match Case Button
        self.ccButton = QPushButton("Cc")
        self.ccButton.setCheckable(True)
        self.ccButton.clicked.connect(self.updateButtonStyle)
        self.ccButton.setToolTip(
            "Match Case: Select this button if you want to match the case that you have in the filter criteria.")
        self.filterRowLayout.addWidget(self.ccButton)

        # Entire Word Button
        self.wButton = QPushButton("W")
        self.wButton.setCheckable(True)
        self.wButton.clicked.connect(self.updateButtonStyle)
        self.wButton.setToolTip("Entire Word: Select this button to filter only exact word matches.")
        self.filterRowLayout.addWidget(self.wButton)

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

        merge_menu = QMenu("Merge", self)
        merge_as_same_action = QAction("Merge as Same", self)
        merge_as_same_action.triggered.connect(lambda: self.merge_tables(as_same=True))
        merge_menu.addAction(merge_as_same_action)
        merge_as_new_action = QAction("Merge as New", self)
        merge_as_new_action.triggered.connect(lambda: self.merge_tables(as_same=False))
        merge_menu.addAction(merge_as_new_action)
        operations_menu.addMenu(merge_menu)

        append_menu = QMenu("Append", self)
        append_as_same_action = QAction("Append as Same", self)
        append_as_same_action.triggered.connect(lambda: self.append_tables(as_same=True))
        append_menu.addAction(append_as_same_action)
        append_as_new_action = QAction("Append as New", self)
        append_as_new_action.triggered.connect(lambda: self.append_tables(as_same=False))
        append_menu.addAction(append_as_new_action)
        operations_menu.addMenu(append_menu)

        pivot_menu = QMenu("Pivot", self)
        pivot_as_same_action = QAction("Pivot as Same", self)
        pivot_as_same_action.triggered.connect(lambda: self.pivot_table(as_same=True))
        pivot_menu.addAction(pivot_as_same_action)
        pivot_as_new_action = QAction("Pivot as New", self)
        pivot_as_new_action.triggered.connect(lambda: self.pivot_table(as_same=False))
        pivot_menu.addAction(pivot_as_new_action)
        operations_menu.addMenu(pivot_menu)

        unpivot_menu = QMenu("Unpivot", self)
        unpivot_as_same_action = QAction("Unpivot as Same", self)
        unpivot_as_same_action.triggered.connect(lambda: self.unpivot_table(as_same=True))
        unpivot_menu.addAction(unpivot_as_same_action)
        unpivot_as_new_action = QAction("Unpivot as New", self)
        unpivot_as_new_action.triggered.connect(lambda: self.unpivot_table(as_same=False))
        unpivot_menu.addAction(unpivot_as_new_action)
        operations_menu.addMenu(unpivot_menu)

    def updateButtonStyle(self):
        """
        Updates the style of the Sw and Cc buttons based on their checked state.
        """
        sender = self.sender()
        if sender.isChecked():
            sender.setStyleSheet("background-color: #4d8d9c;")  # Darker color when checked
        else:
            sender.setStyleSheet("")  # Revert to default stylesheet
        self.filterTable()

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

            # Adjust for "Entire Word" functionality based on W button
            if self.wButton.isChecked():
                match = cell_value == filter_text
            # Adjust for "Starts With" functionality based on Sw button
            elif self.swButton.isChecked():
                match = cell_value.startswith(filter_text)
            else:
                match = filter_text in cell_value

            self.table_view.setRowHidden(row, not match)

    def add_table(self):
        self.loading_dialog.show()
        options = QFileDialog.Option.ReadOnly
        file_path, _ = QFileDialog.getOpenFileName(self, "Add Table", "",
                                                   "Excel files (*.xlsx *.xls *.xlsm);;CSV files (*.csv);;"
                                                   "Text files (*.txt)",
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
                self.loading_dialog.hide()
                return
            table_name = file_name_without_ext
            self.tables[table_name] = TableRevision(data)
            self.tables[table_name].spreadsheet_name = file_name_without_ext
            self.tables[table_name].sheet_name = sheet_name
            self.tables[table_name].extension = extension if extension else ".xlsx"
            self.file_list.addItem(table_name)
        self.show_table(self.file_list.currentItem())
        self.loading_dialog.hide()

    def export_tables(self):
        if not self.tables:
            QMessageBox.warning(self, "No Tables", "No tables available to export. Please add tables to the app first.")
            return

        dialog = ExportDialog(self.tables, parent=self)
        dialog.exec()

    def populate_table(self, data):
        self.loading_dialog.show()
        self.table_view.clear()
        self.table_view.setColumnCount(len(data.columns))
        self.table_view.setRowCount(len(data))
        self.table_view.setHorizontalHeaderLabels([f"{col} ({dtype})" for col, dtype in zip(data.columns, data.dtypes)])

        for i in range(len(data)):
            for j in range(len(data.columns)):
                item = QTableWidgetItem(str(data.iloc[i, j]))
                self.table_view.setItem(i, j, item)

        self.loading_dialog.hide()

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

    def merge_tables(self, as_same=True):
        if len(self.tables) < 2:
            QMessageBox.warning(self, "Error", "At least two tables are required for merging.")
            return

        selected_table = self.file_list.currentItem().text()
        dialog = MergeDialog(self.tables, selected_table, parent=self)  # Pass self as the parent
        result = dialog.exec()

        if result == QDialog.DialogCode.Accepted:
            merged_data = dialog.merged_data
            if as_same:
                table_name = self.file_list.currentItem().text()
                table_revision = self.tables[table_name]
                table_revision.add_revision(merged_data)  # Add merged data as a new revision
                self.populate_table(merged_data)
                current_item = self.file_list.currentItem()
                self.file_list.setCurrentItem(current_item)
                self.show_table(current_item)
            else:
                new_table_name = self.generate_new_table_name("Query")
                self.tables[new_table_name] = TableRevision(merged_data)
                new_item = QListWidgetItem(new_table_name)
                self.file_list.addItem(new_item)
                # Select the new table in the file list and show it
                self.file_list.setCurrentItem(new_item)
                self.show_table(new_item)

    def append_tables(self, as_same=True):
        if len(self.tables) < 2:
            QMessageBox.warning(self, "Error", "At least two tables are required for appending.")
            return

        selected_table = self.file_list.currentItem().text()
        dialog = AppendDialog(self.tables, selected_table, parent=self)  # Pass self as the parent
        result = dialog.exec()

        if result == QDialog.DialogCode.Accepted:
            appended_data = dialog.appended_data
            if as_same:
                table_name = self.file_list.currentItem().text()
                table_revision = self.tables[table_name]
                table_revision.add_revision(appended_data)  # Add appended data as a new revision
                self.populate_table(appended_data)
                current_item = self.file_list.currentItem()
                self.file_list.setCurrentItem(current_item)
                self.show_table(current_item)
            else:
                new_table_name = self.generate_new_table_name("Query")
                self.tables[new_table_name] = TableRevision(appended_data)
                new_item = QListWidgetItem(new_table_name)
                self.file_list.addItem(new_item)
                # Select the new table in the file list and show it
                self.file_list.setCurrentItem(new_item)
                self.show_table(new_item)

    def pivot_table(self, as_same=True):
        if len(self.tables) == 0:
            QMessageBox.warning(self, "Error", "No tables available for pivoting.")
            return

        selected_indexes = self.table_view.selectedIndexes()
        if len(selected_indexes) == 0:
            QMessageBox.warning(self, "Error", "Please select a single column to pivot.")
            return

        selected_column_with_dtype = self.table_view.horizontalHeaderItem(selected_indexes[0].column()).text()
        selected_column = selected_column_with_dtype.split(" (")[0]  # Extract column name without data type

        selected_table = self.file_list.currentItem().text()
        table_revision = self.tables[selected_table]
        data = table_revision.revisions[table_revision.current_revision]

        dialog = PivotDialog(data, selected_column, parent=self)
        if dialog.exec() == QDialog.DialogCode.Accepted:
            values_column_with_dtype = dialog.get_values_column()
            values_column = values_column_with_dtype.split(" (")[0]  # Extract column name without data type

            if values_column == selected_column:
                QMessageBox.warning(self, "Error", "The values column cannot be the same as the selected column.")
                return

            self.loading_dialog.show()
            # Perform the pivot operation
            pivot_data = data.pivot_table(index=data.columns[0], columns=selected_column, values=values_column,
                                          aggfunc="sum")
            self.loading_dialog.hide()

            if as_same:
                table_revision.add_revision(pivot_data)  # Add pivot data as a new revision
                self.populate_table(pivot_data)
                current_item = self.file_list.currentItem()
                self.file_list.setCurrentItem(current_item)
                self.show_table(current_item)
            else:
                new_table_name = self.generate_new_table_name("Query")
                self.tables[new_table_name] = TableRevision(pivot_data)
                new_item = QListWidgetItem(new_table_name)
                self.file_list.addItem(new_item)
                self.file_list.setCurrentItem(new_item)
                self.show_table(new_item)

    def generate_new_table_name(self, prefix):
        i = 1
        while True:
            new_table_name = f"{prefix} {i}"
            if new_table_name not in self.tables:
                return new_table_name
            i += 1

    def unpivot_table(self, as_same=True):
        if len(self.tables) == 0:
            QMessageBox.warning(self, "Error", "No tables available for unpivoting.")
            return

        selected_indexes = self.table_view.selectedIndexes()
        if len(selected_indexes) < 2:
            QMessageBox.warning(self, "Error", "Please select multiple columns to unpivot.")
            return

        selected_table = self.file_list.currentItem().text()
        table_revision = self.tables[selected_table]
        data = table_revision.revisions[table_revision.current_revision]

        selected_columns = [index.column() for index in selected_indexes]
        unique_columns = list(set(selected_columns))

        if len(unique_columns) < 2:
            QMessageBox.warning(self, "Error", "Please select at least two unique columns to unpivot.")
            return

        self.loading_dialog.show()

        column_names = [data.columns[col].split(" (")[0] for col in
                        unique_columns]  # Extract column names without data types

        # Perform the unpivot operation
        id_vars = [col for col in data.columns if col not in column_names]
        unpivoted_data = data.melt(id_vars=id_vars, value_vars=column_names, var_name='Variable', value_name='Value')

        self.loading_dialog.hide()

        if as_same:
            table_revision.add_revision(unpivoted_data)  # Add unpivoted data as a new revision
            self.populate_table(unpivoted_data)
            current_item = self.file_list.currentItem()
            self.file_list.setCurrentItem(current_item)
            self.show_table(current_item)
        else:
            new_table_name = self.generate_new_table_name("Query")
            self.tables[new_table_name] = TableRevision(unpivoted_data)
            new_item = QListWidgetItem(new_table_name)
            self.file_list.addItem(new_item)
            self.file_list.setCurrentItem(new_item)
            self.show_table(new_item)

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
    assets_path = "assets"
    icons_path = os.path.join(assets_path, "images")
    style_path = os.path.join(assets_path, "style")
    try:
        stylesheet = os.path.join(resource_path(style_path), "stylesheet.qss")
        with open(stylesheet, "r") as file:
            return file.read().replace('{{ICON_PATH}}', str(resource_path(icons_path)).replace("\\", "/"))
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
