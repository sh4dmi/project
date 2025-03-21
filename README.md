# Excel Operations Library

A comprehensive Python library for Excel operations with JSON input support and reward feedback.

## Overview

This library provides a high-level interface for Excel operations, including reading, writing, and clearing data. It's designed to be easy to use, with robust error handling and detailed logging. The library supports both direct API calls and JSON-based operations with reward feedback.

## Features

- **Comprehensive Excel Operations**: Read, write, and clear data in Excel files
- **Robust Error Handling**: Detailed error messages and logging
- **JSON Input Support**: Process Excel operations from JSON input
- **Reward Feedback**: Get reward and feedback for operation success/failure
- **Input Validation**: Thorough validation of all inputs
- **Flexible Column Indexing**: Support for both numerical indices and column letters (A, B, C, etc.)

## Files

- `excel_functions.py`: Core Excel operations and JSON processing functionality
- `test.py`: Comprehensive test suite for all operations
- `main.py`: Interactive playground for testing operations
- `README.md`: Documentation (this file)

## Basic Usage

```python
from excel_functions import ExcelHandler

# Create an Excel handler
excel = ExcelHandler("example.xlsx")

# Write data
excel.write_row(1, ["ID", "Name", "Age", "Department"])  # Write header
excel.write_row(2, [1, "John Smith", 35, "Engineering"])  # Write data row
excel.write_cell(3, "B", "Mary Johnson")  # Write to cell B3

# Read data
header_row, _ = excel.read_header_row()  # Read header row
cell_value, _ = excel.read_cell(2, "B")  # Read cell B2
col_index, _ = excel.get_column_index_by_header("Name")  # Find column index by header

# Clear data
excel.clear_cell(2, 3)  # Clear cell at row 2, column 3
excel.clear_row(3)  # Delete row 3
excel.clear_column("B")  # Delete column B
excel.clear_sheet()  # Clear the entire sheet
```

## JSON Input Usage

The library also supports JSON-based operations with reward feedback:

```python
import json
from excel_functions import ExcelHandler

excel = ExcelHandler("example.xlsx")

# Example: Write a header row
json_input = json.dumps({
    "function_name": "excel_write_row",
    "parameters": {
        "row_index": 1,
        "row_data": ["ID", "Name", "Age", "Department"]
    }
})
reward, feedback = excel.process_json_operation(json_input)
print(f"Reward: {reward}, Feedback: {feedback}")

# Example: Read a cell
json_input = json.dumps({
    "function_name": "excel_read_cell",
    "parameters": {
        "row_index": 2,
        "col_index": "B"
    }
})
reward, feedback = excel.process_json_operation(json_input)
print(f"Reward: {reward}, Feedback: {feedback}")
```

## JSON Function Call Format

All Excel operations can be performed using JSON input in the following format:

```json
{
    "function_name": "excel_function_name",
    "parameters": {
        "param1": value1,
        "param2": value2,
        ...
    }
}
```

For example:

```json
{
    "function_name": "excel_write_cell",
    "parameters": {
        "row_index": 2,
        "col_index": "B",
        "text": "John Smith"
    }
}
```

## Reward Function

The reward function evaluates the success of an operation and provides feedback:

- Reward `1`: Operation was successful
- Reward `-1`: Operation failed (with detailed error message)

The feedback string provides detailed information about the operation result, including any error messages or return values.

## Available Methods

### File Operations

- `ExcelHandler(filename)`: Initialize with an Excel file
- `clear_sheet()`: Clear all data from the active sheet

### Writing Data

- `add_row(row_index, text)`: Add a new row with text in the first cell
- `write_cell(row_index, col_index, text)`: Write text to a specific cell
- `write_row(row_index, row_data)`: Write data to an entire row

### Reading Data

- `read_header_row()`: Read the header row (first row)
- `read_column(col_index)`: Read an entire column
- `read_cell(row_index, col_index)`: Read a specific cell
- `read_row(row_index)`: Read an entire row
- `get_column_index_by_header(header_name)`: Find column index by header name

### Clearing Data

- `clear_cell(row_index, col_index)`: Clear a specific cell
- `clear_row(row_index)`: Delete an entire row
- `clear_column(col_index)`: Delete an entire column

### JSON Processing

- `process_json_operation(json_input)`: Process a JSON-formatted Excel operation

## Input Handling

- Row indices are 1-based (as in Excel)
- Column indices can be numerical (1, 2, 3) or letters (A, B, C)
- The special value "next_available" can be used for row_index in add_row()

## Running Tests

Run the comprehensive test suite:

```bash
python test.py
```

## Interactive Playground

Use the interactive playground to test operations:

```bash
python main.py
```

In the playground, you can:
- Enter JSON commands to perform operations
- Get immediate feedback and reward
- Set up demo data with sample employees
- View help information about available operations

## Requirements

- Python 3.6+
- openpyxl

## Error Handling

The library provides detailed error messages for all operations. Each function returns a tuple:

- For operations: `(success_flag, message)`
- For JSON processing: `(reward, feedback)`

## Examples

### Adding and Writing Data

```python
excel = ExcelHandler("employees.xlsx")

# Add empty rows
excel.add_row(1, "Header Row")
excel.add_row("next_available", "First Employee")

# Write header row
excel.write_row(1, ["ID", "Name", "Age", "Department", "Salary"])

# Write employee data
excel.write_row(2, [1, "John Smith", 35, "Engineering", 75000])
excel.write_row(3, [2, "Mary Johnson", 42, "Finance", 82000])
```

### Reading and Finding Data

```python
# Read the header row
header_row, _ = excel.read_header_row()
print("Headers:", header_row)

# Find column index by header
dept_col_index, _ = excel.get_column_index_by_header("Department")
print(f"Department column index: {dept_col_index}")

# Read a specific cell using the found column index
dept_name, _ = excel.read_cell(2, dept_col_index)
print(f"John's department: {dept_name}")

# Read an entire row
employee2_data, _ = excel.read_row(2)
print("Employee 1 data:", employee2_data)

# Read an entire column
names_column, _ = excel.read_column("B")
print("Names column:", names_column)
```

### Clearing and Modifying Data

```python
# Clear a specific cell
excel.clear_cell(3, "E")  # Clear Mary's salary

# Delete a row
excel.clear_row(2)  # Remove John's row

# Delete a column
excel.clear_column("C")  # Remove Age column

# Clear the entire sheet
excel.clear_sheet()
``` 