import json

# Default test scenarios focused on write_cell
default_scenarios = [
    {
        "prompt": "Write 'Hello World' to cell B2",
        "expected_params": {
            "row_index": 2,
            "col_index": "B", 
            "text": "Hello World"
        }
    },
    {
        "prompt": "Please add the text 'Employee Name' to cell A1",
        "expected_params": {
            "row_index": 1,
            "col_index": "A", 
            "text": "Employee Name"
        }
    },
    {
        "prompt": "Put the value 42 in cell C3",
        "expected_params": {
            "row_index": 3,
            "col_index": "C", 
            "text": "42"
        }
    },
    {
        "prompt": "Write 'Department' to the cell at column D, row 1",
        "expected_params": {
            "row_index": 1,
            "col_index": "D", 
            "text": "Department"
        }
    },
    {
        "prompt": "Add the text 'Total: $1000' to cell E5",
        "expected_params": {
            "row_index": 5,
            "col_index": "E", 
            "text": "Total: $1000"
        }
    }
]

# Write scenarios to file
with open("write_cell_scenarios.json", "w") as f:
    json.dump(default_scenarios, f, indent=4)

print("Created write_cell_scenarios.json with default test scenarios.") 