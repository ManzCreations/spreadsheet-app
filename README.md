# Spreadsheet Application

The Spreadsheet Application is a Python-based desktop application that allows users to load, append, merge, pivot, and unpivot data from Excel or CSV files. It provides a user-friendly interface for manipulating spreadsheet data and saving the results to new files.

## Features

- Load Excel or CSV files into the application
- Append data from another file to the currently loaded data
- Merge data from another file with the currently loaded data based on a specified column
- Create a pivot table from the currently loaded data by specifying the index, columns, and values
- Unpivot the currently loaded data, converting columns to rows
- Save the modified data to a new Excel or CSV file
- Preview the loaded and modified data in the application

## Requirements

- Python 3.x
- PyQt5
- PySide2
- Pandas
- openpyxl (for Excel file support)

## Installation

1. Clone the repository:
git clone https://github.com/yourusername/spreadsheet-app.git
2. Install the required dependencies:
pip install PyQt5 PySide2 pandas openpyxl
3. Run the application:
python spreadsheet_app.py
## Usage

1. Launch the Spreadsheet Application.
2. Click the "Load Spreadsheet" button to select an Excel or CSV file to load into the application.
3. Use the provided buttons to perform various operations on the loaded data:
- "Append" button: Append data from another file to the currently loaded data.
- "Merge" button: Merge data from another file with the currently loaded data based on a specified column.
- "Pivot" button: Create a pivot table from the currently loaded data by specifying the index, columns, and values.
- "Unpivot" button: Unpivot the currently loaded data, converting columns to rows.
4. Preview the loaded and modified data in the application's text area.
5. Click the "Save Spreadsheet" button to save the modified data to a new Excel or CSV file.

## Contributing

Contributions are welcome! If you find any issues or have suggestions for improvements, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE).
