# Synthetic Excel Data Generation

This project generates synthetic data for training language models to interact with Excel files. It creates a dataset of user instructions and corresponding JSON commands to update cells in Excel based on lookup.

## Overview

The main script `generate_json.py` contains three key components:

1. **ExcelHandler Class**: A simulation of Excel file operations for development purposes.
2. **generate_project_excel_file Function**: Creates synthetic Excel files with project data.
3. **generate_instruction_json_from_excel Function**: Generates natural language instructions and JSON objects for updating cells.

## How It Works

### 1. Generating Excel Files

The script first creates Excel files with synthetic project data:

- Each file contains a "Projects" table with headers like "Project Name", "Description", "Status", etc.
- Data types are specified for certain columns (numbers, dates, selection lists)
- Random data is generated for each cell based on the column type
- For this simulation, Excel files are represented as JSON files with the same base name

### 2. Generating Instructions and Function Calls

For each Excel file:

- A random cell is selected
- A new value is generated based on the column's data type
- A natural language instruction is created using templates (e.g., "Change the Status of Project X to In Progress")
- A corresponding JSON object is generated that represents the function call to update the cell

### 3. Creating the Dataset

The script compiles these examples into a dataset saved as `synthetic_excel_data_excel_files_v2.json`:

- Each example includes:
  - `instruction`: Natural language instruction for updating a cell
  - `context`: JSON representation of the Excel data (headers and rows)
  - `function_call`: JSON object representing the update function call
  - `excel_file_path`: Path to the Excel file

## Output Format

The output dataset has the following structure:

```json
[
  {
    "instruction": "Change the Status of Project X to In Progress",
    "context": {
      "PROJECTS": {
        "headers": ["Project Name", "Description", "Status", "..."],
        "rows": [
          ["Project-123: ...", "Description text", "Planning", "..."],
          ...
        ]
      }
    },
    "function_call": {
      "function_name": "excel_update_cell_by_lookup",
      "parameters": {
        "row_header": "Project Name",
        "row_value": "Project-123: ...",
        "col_header": "Status",
        "new_value": "In Progress"
      },
      "target_cell": "C2",
      "expected_cell_value": "In Progress"
    },
    "excel_file_path": "path/to/excel/file.xlsx"
  },
  ...
]
```

## Usage

To generate synthetic data:

1. Install the required libraries:
   ```
   pip install faker
   ```

2. Run the script:
   ```
   python generate_json.py
   ```

3. The script will generate:
   - Multiple Excel files (simulated as JSON files in this implementation)
   - A dataset file (`synthetic_excel_data_excel_files_v2.json`)

## Customization

You can customize the generation process by modifying:

- `num_examples_to_generate`: Number of examples to create
- `project_schema`: Headers and data types for the project data
- Instruction templates in `instruction_templates`
- Data generation logic for different column types

## Notes for Production Use

In a production environment:

1. Replace `ExcelHandler` with actual Excel file operations using libraries like `openpyxl` or `pandas`
2. Implement proper error handling for file operations
3. Add more diverse data generation logic for realistic data
4. Consider adding more complex instruction templates and scenarios 