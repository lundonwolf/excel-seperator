# Excel Sheet Separator

This application helps you separate an Excel spreadsheet into multiple sheets based on the content in the first column.

## Features

- Simple GUI interface for selecting Excel files
- Automatically detects table boundaries based on first column content
- Creates a new Excel file with separated tables in different sheets
- Preserves data formatting and structure
- Handles both .xlsx and .xls files

## Requirements

- Python 3.x
- pandas
- openpyxl

## Installation

1. Install the required packages:
```bash
pip install -r requirements.txt
```

## Usage

1. Run the application:
```bash
python excel_separator.py
```

2. Click "Select Excel File" to choose your Excel file
3. The application will automatically process the file and create a new Excel workbook with separated sheets
4. The output file will be saved in the same directory as the input file with "_separated" added to the filename

## How it works

The application separates tables based on these rules:
- Each non-empty cell in the first column marks the beginning of a new table
- Empty cells in the first column indicate that the row belongs to the previous table
- Each table is placed in a separate sheet, named after its first column value
