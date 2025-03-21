#!/usr/bin/env python3
"""
Excel Operations Interactive Playground
======================================

This script provides an interactive playground for testing Excel operations.
Users can test various Excel operations using JSON input format and see the
results and reward feedback in real-time.
"""

import os
import json
from excel_functions import ExcelHandler

def print_help():
    """Print help information about available JSON operations."""
    print("\nAvailable Excel Operations:")
    print("---------------------------")
    print("1. excel_clear_sheet - Clear the entire sheet")
    print("   Example: {\"function_name\": \"excel_clear_sheet\", \"parameters\": {}}")
    print()
    print("2. excel_add_row - Add a row at specific index or next available")
    print("   Example: {\"function_name\": \"excel_add_row\", \"parameters\": {\"row_index\": 5, \"text\": \"New row\"}}")
    print()
    print("3. excel_write_cell - Write text to a specific cell")
    print("   Example: {\"function_name\": \"excel_write_cell\", \"parameters\": {\"row_index\": 2, \"col_index\": 3, \"text\": \"Cell content\"}}")
    print()
    print("4. excel_write_row - Write data to an entire row")
    print("   Example: {\"function_name\": \"excel_write_row\", \"parameters\": {\"row_index\": 2, \"row_data\": [1, \"John\", 30, \"IT\"]}}")
    print()
    print("5. excel_clear_cell - Clear the content of a specific cell")
    print("   Example: {\"function_name\": \"excel_clear_cell\", \"parameters\": {\"row_index\": 2, \"col_index\": 3}}")
    print()
    print("6. excel_clear_row - Clear/delete an entire row")
    print("   Example: {\"function_name\": \"excel_clear_row\", \"parameters\": {\"row_index\": 3}}")
    print()
    print("7. excel_clear_column - Clear/delete an entire column")
    print("   Example: {\"function_name\": \"excel_clear_column\", \"parameters\": {\"col_index\": 2}}")
    print()
    print("8. excel_read_header_row - Read the header row")
    print("   Example: {\"function_name\": \"excel_read_header_row\"}")
    print()
    print("9. excel_read_column - Read an entire column")
    print("   Example: {\"function_name\": \"excel_read_column\", \"parameters\": {\"col_index\": 2}}")
    print()
    print("10. excel_read_cell - Read the content of a specific cell")
    print("   Example: {\"function_name\": \"excel_read_cell\", \"parameters\": {\"row_index\": 2, \"col_index\": 3}}")
    print()
    print("11. excel_read_row - Read an entire row")
    print("   Example: {\"function_name\": \"excel_read_row\", \"parameters\": {\"row_index\": 2}}")
    print()
    print("12. excel_get_column_index_by_header - Find column index by header name")
    print("   Example: {\"function_name\": \"excel_get_column_index_by_header\", \"parameters\": {\"header_name\": \"Name\"}}")
    print()
    print("Commands:")
    print("  help       - Display this help information")
    print("  exit       - Exit the program")
    print("  save       - Save the current Excel file")
    print("  clear      - Clear the console screen")
    print("  setup_demo - Set up a demo with sample data")
    print("  inspect    - Inspect the first few cells of the sheet (for debugging)")

def inspect_sheet(excel_handler):
    """Display the first few cells of the sheet for debugging purposes."""
    print("\n=== Sheet Inspection ===")
    print("First 6 cells of first 6 rows:")
    
    # Get max rows/cols to check (up to 6 of each)
    max_row = min(6, excel_handler.sheet.max_row)
    
    # Display header
    print("\n    | A      | B      | C      | D      | E      | F      |")
    print("----|--------|--------|--------|--------|--------|--------|")
    
    # Display cells
    for row in range(1, max_row + 1):
        row_str = f"{row:3} |"
        for col in range(1, 7):  # Columns A through F
            try:
                value = excel_handler.sheet.cell(row=row, column=col).value
                if value is None:
                    cell_str = "(empty)"
                else:
                    # Truncate long values
                    cell_str = str(value)[:6] + "..." if len(str(value)) > 6 else str(value)
                row_str += f" {cell_str:6} |"
            except Exception:
                row_str += " ERROR  |"
        print(row_str)
    
    # Special check for cell A1 (debugging the A1 issue)
    print("\nCell A1 specific check:")
    try:
        a1_value = excel_handler.sheet.cell(row=1, column=1).value
        print(f"  A1 value = {a1_value}")
        print(f"  A1 type = {type(a1_value).__name__}")
    except Exception as e:
        print(f"  Error reading A1: {str(e)}")
    
    print("\nUse read commands for more detailed inspection.")

def setup_demo_data(excel_handler):
    """Set up sample data for demonstration purposes."""
    print("\nSetting up demo data...")
    
    # Clear any existing data
    excel_handler.clear_sheet()
    
    # Create header row
    headers = ["ID", "Name", "Age", "Department", "Salary"]
    excel_handler.write_row(1, headers)
    
    # Add sample employee data
    employees = [
        [1, "John Smith", 35, "Engineering", 75000],
        [2, "Mary Johnson", 42, "Finance", 82000],
        [3, "Robert Brown", 28, "Marketing", 65000],
        [4, "Michael Davis", 33, "HR", 68000],
        [5, "Jennifer Wilson", 38, "Operations", 72000]
    ]
    
    for i, employee in enumerate(employees):
        excel_handler.write_row(i + 2, employee)
    
    # Save the workbook
    excel_handler.workbook.save(excel_handler.filename)
    
    print("Demo data setup complete. Added header row and 5 employees.")
    print("Try reading the data with: {\"function_name\": \"excel_read_header_row\"}")

def clear_screen():
    """Clear the console screen."""
    os.system('cls' if os.name == 'nt' else 'clear')

def main():
    """Main function for the interactive Excel operations playground."""
    # Create an Excel file for testing
    excel_file = "playground.xlsx"
    
    # Check if the file exists, and remove it if it does
    if os.path.exists(excel_file):
        try:
            os.remove(excel_file)
            print(f"Removed existing file: {excel_file}")
        except PermissionError:
            print(f"Warning: Could not remove existing file {excel_file}. It may be open in another program.")
    
    # Create an instance of ExcelHandler
    excel = ExcelHandler(excel_file)
    
    print("=" * 80)
    print("Excel Operations Interactive Playground".center(80))
    print("=" * 80)
    print(f"Working with Excel file: {excel_file}")
    print("Enter JSON commands to perform Excel operations.")
    print("Type 'help' for available operations, 'exit' to quit, or 'setup_demo' for sample data.")
    
    while True:
        print("\n" + "-" * 80)
        user_input = input("Enter JSON command (or help/exit/save/clear/setup_demo/inspect): ").strip()
        
        if user_input.lower() == 'exit':
            print("Exiting Excel Operations Playground...")
            break
        
        elif user_input.lower() == 'help':
            print_help()
            
        elif user_input.lower() == 'save':
            try:
                excel.workbook.save(excel_file)
                print(f"Excel file saved successfully: {excel_file}")
            except Exception as e:
                print(f"Error saving Excel file: {str(e)}")
                
        elif user_input.lower() == 'clear':
            clear_screen()
            
        elif user_input.lower() == 'setup_demo':
            setup_demo_data(excel)
            
        elif user_input.lower() == 'inspect':
            inspect_sheet(excel)
            
        else:
            try:
                # Before processing JSON, make sure we're not altering anything unintentionally
                print(f"Processing JSON command...")
                
                # Check cell A1 before processing (for debugging the A1 issue)
                try:
                    a1_before = excel.sheet.cell(row=1, column=1).value
                    print(f"DEBUG: A1 value BEFORE command = {a1_before}")
                except Exception as e:
                    print(f"DEBUG: Error reading A1 before command: {str(e)}")
                
                # Process the JSON command
                reward, feedback = excel.process_json_operation(user_input)
                
                # Check cell A1 after processing (for debugging the A1 issue)
                try:
                    a1_after = excel.sheet.cell(row=1, column=1).value
                    print(f"DEBUG: A1 value AFTER command = {a1_after}")
                    if a1_before != a1_after:
                        print(f"WARNING: A1 changed from '{a1_before}' to '{a1_after}'")
                except Exception as e:
                    print(f"DEBUG: Error reading A1 after command: {str(e)}")
                
                # Print the result with color coding
                if reward == 1:  # Success
                    print("\n✅ SUCCESS | Reward: 1")
                    print(f"Feedback: {feedback}")
                    
                    # Offer to display the current state of the data for verification
                    print("\nTip: To verify the results, you can read data with commands like:")
                    print('{"function_name": "excel_read_header_row"}')
                    print('{"function_name": "excel_read_cell", "parameters": {"row_index": 2, "col_index": "B"}}')
                else:  # Error
                    print("\n❌ ERROR | Reward: -1")
                    print(f"Feedback: {feedback}")
                    
            except Exception as e:
                print(f"\n❌ ERROR: {str(e)}")
    
    # Save and close the workbook before exiting
    try:
        excel.workbook.close()
        print(f"Excel file closed: {excel_file}")
    except Exception as e:
        print(f"Error closing Excel file: {str(e)}")

if __name__ == "__main__":
    main()