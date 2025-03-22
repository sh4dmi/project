#!/usr/bin/env python3
"""
Comprehensive Excel Functions Test Suite
=======================================

This script provides a comprehensive test suite for all Excel operations,
testing both direct method calls and JSON-based operations.
It tests success cases, failure cases, and edge cases.
"""

import unittest
import json
import os
from excel_functions import ExcelHandler

class TestExcelFunctions(unittest.TestCase):
    """
    Comprehensive test suite for the ExcelHandler class.
    Tests all operations via direct API calls and JSON input.
    """
    
    def setUp(self):
        """Prepare test environment before each test case."""
        self.test_file = "test_excel.xlsx"
        # Remove test file if it exists from previous test runs
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        self.excel = ExcelHandler(self.test_file)
        
        # Set up some initial data for tests that need existing data
        self.setup_initial_data()
    
    def tearDown(self):
        """Clean up after each test case."""
        # Close workbook to release file handle
        self.excel.workbook.close()
        # Remove test file after tests
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
    
    def setup_initial_data(self):
        """Set up initial data for tests that need existing data."""
        # Create a header row
        self.excel.write_row(1, ["ID", "Name", "Age", "Department", "Salary"])
        
        # Add some sample data
        employees = [
            [1, "John Smith", 35, "Engineering", 75000],
            [2, "Mary Johnson", 42, "Finance", 82000],
            [3, "Robert Brown", 28, "Marketing", 65000]
        ]
        
        for i, employee in enumerate(employees):
            self.excel.write_row(i + 2, employee)
        
        # Save the workbook
        self.excel.workbook.save(self.test_file)
    
    #
    # DIRECT API TESTS
    #
    
    def test_direct_clear_sheet(self):
        """Test clearing a sheet directly."""
        # Clear the sheet
        success, message = self.excel.clear_sheet()
        
        # Verify
        self.assertTrue(success)
        self.assertIn("Sheet", message)
        self.assertIn("clear", message)
        self.assertEqual(self.excel.sheet.max_row, 1)  # Empty sheet has 1 row
        
        # Verify cell is empty
        cell_value, _ = self.excel.read_cell(1, 1)
        self.assertIsNone(cell_value)
    
    def test_direct_add_row(self):
        """Test adding a row directly."""
        # Add a row at a specific index
        success1, message1 = self.excel.add_row(5, "New employee row")
        
        # Add a row at next available index
        success2, message2 = self.excel.add_row("next_available", "Next available row")
        
        # Verify
        self.assertTrue(success1)
        self.assertTrue(success2)
        self.assertIn("row", message1.lower())
        self.assertIn("5", message1)
        
        # Verify content
        cell_value1, _ = self.excel.read_cell(5, 1)
        self.assertEqual(cell_value1, "New employee row")
        
        # Test invalid input
        success_invalid, message_invalid = self.excel.add_row(-1, "Invalid row")
        self.assertFalse(success_invalid)
        self.assertIn("must be positive", message_invalid)
    
    def test_direct_write_cell(self):
        """Test writing to a cell directly."""
        # Write to cell
        success, message = self.excel.write_cell(2, 2, "Updated Name")
        
        # Verify
        self.assertTrue(success)
        self.assertIn("Value", message)
        self.assertIn("written", message)
        
        # Verify content
        cell_value, _ = self.excel.read_cell(2, 2)
        self.assertEqual(cell_value, "Updated Name")
        
        # Test with column letter
        success_letter, message_letter = self.excel.write_cell(3, "B", "Column Letter Test")
        self.assertTrue(success_letter)
        
        # Verify content
        cell_value_letter, _ = self.excel.read_cell(3, "B")
        self.assertEqual(cell_value_letter, "Column Letter Test")
        
        # Test invalid input
        success_invalid, message_invalid = self.excel.write_cell("invalid", 1, "Test")
        self.assertFalse(success_invalid)
    
    def test_direct_write_row(self):
        """Test writing an entire row directly."""
        # Write a row
        row_data = [1, "Updated Employee", 30, "IT", 70000]
        success, message = self.excel.write_row(2, row_data)
        
        # Verify
        self.assertTrue(success)
        self.assertIn("Data written to row", message)
        
        # Verify content
        row_value, _ = self.excel.read_row(2)
        self.assertEqual(row_value, row_data)
        
        # Test with mixed data types
        mixed_row = ["String", 42, 3.14, True, None]
        success_mixed, _ = self.excel.write_row(3, mixed_row)
        self.assertTrue(success_mixed)
        
        # Verify mixed content
        mixed_result, _ = self.excel.read_row(3)
        self.assertEqual(mixed_result, mixed_row)
        
        # Test invalid input
        success_invalid, message_invalid = self.excel.write_row("invalid", row_data)
        self.assertFalse(success_invalid)
        
        # Test non-iterable input
        success_non_iterable, message_non_iterable = self.excel.write_row(2, "not iterable")
        self.assertFalse(success_non_iterable)
        self.assertIn("iterable collection, not a string", message_non_iterable)
    
    def test_direct_clear_cell(self):
        """Test clearing a cell directly."""
        # Set up a cell with data
        self.excel.write_cell(2, 3, "Test Cell")
        
        # Verify initial content
        cell_value_before, _ = self.excel.read_cell(2, 3)
        self.assertEqual(cell_value_before, "Test Cell")
        
        # Clear the cell
        success, message = self.excel.clear_cell(2, 3)
        
        # Verify
        self.assertTrue(success)
        self.assertIn("cleared", message.lower())
        
        # Verify cell is cleared
        cell_value_after, _ = self.excel.read_cell(2, 3)
        self.assertIsNone(cell_value_after)
        
        # Test with column letter
        self.excel.write_cell(3, "C", "Test Cell Letter")
        success_letter, _ = self.excel.clear_cell(3, "C")
        self.assertTrue(success_letter)
        
        # Test invalid input
        success_invalid, _ = self.excel.clear_cell("invalid", 3)
        self.assertFalse(success_invalid)
    
    def test_direct_clear_row(self):
        """Test clearing a row directly."""
        # Get data before clearing
        row3_before, _ = self.excel.read_row(3)
        
        # Clear row 2
        success, message = self.excel.clear_row(2)
        
        # Verify
        self.assertTrue(success)
        self.assertIn("Row", message)
        self.assertIn("2", message)
        self.assertIn("deleted", message)
        
        # Verify row 3 is now at position 2
        row2_after, _ = self.excel.read_row(2)
        self.assertEqual(row2_after, row3_before)
        
        # Test invalid input
        success_invalid, _ = self.excel.clear_row("invalid")
        self.assertFalse(success_invalid)
    
    def test_direct_clear_column(self):
        """Test clearing a column directly."""
        # Get data before clearing
        header_before, _ = self.excel.read_header_row()
        
        # Clear column 2 (Name)
        success, message = self.excel.clear_column(2)
        
        # Verify
        self.assertTrue(success)
        self.assertIn("Column", message)
        self.assertIn("deleted", message)
        
        # Verify column is removed and Age moved to position 2
        header_after, _ = self.excel.read_header_row()
        self.assertEqual(header_after[1], header_before[2])  # Age was at 3, now at 2
        
        # Test with column letter
        success_letter, _ = self.excel.clear_column("B")  # Now Age becomes column B
        self.assertTrue(success_letter)
        
        # Test invalid input
        success_invalid, _ = self.excel.clear_column("invalid")
        self.assertFalse(success_invalid)
    
    def test_direct_read_header_row(self):
        """Test reading the header row directly."""
        # Read header row
        header, message = self.excel.read_header_row()
        
        # Verify
        self.assertIsNotNone(header)
        self.assertIn("Header row read", message)
        self.assertEqual(header[0], "ID")
        self.assertEqual(header[1], "Name")
    
    def test_direct_read_column(self):
        """Test reading a column directly."""
        # Read column by index
        column, message = self.excel.read_column(2)
        
        # Verify
        self.assertIsNotNone(column)
        self.assertIn("Column", message)
        self.assertIn("read", message)
        self.assertEqual(column[0], "Name")
        self.assertEqual(column[1], "John Smith")
        
        # Read column by letter
        column_letter, message_letter = self.excel.read_column("B")
        
        # Verify
        self.assertIsNotNone(column_letter)
        self.assertEqual(column_letter, column)
        
        # Test invalid input
        column_invalid, _ = self.excel.read_column("invalid")
        self.assertIsNone(column_invalid)
    
    def test_direct_read_cell(self):
        """Test reading a cell directly."""
        # Read cell
        cell_value, message = self.excel.read_cell(2, 2)
        
        # Verify
        self.assertEqual(cell_value, "John Smith")
        self.assertIn("Value", message)
        self.assertIn("read", message)
        
        # Read with column letter
        cell_letter, _ = self.excel.read_cell(2, "B")
        self.assertEqual(cell_letter, "John Smith")
        
        # Test invalid input
        cell_invalid, _ = self.excel.read_cell(-1, 1)
        self.assertIsNone(cell_invalid)
    
    def test_direct_read_row(self):
        """Test reading a row directly."""
        # Read row
        row, message = self.excel.read_row(2)
        
        # Verify
        self.assertIsNotNone(row)
        self.assertIn("Row", message)
        self.assertIn("read", message)
        self.assertEqual(row[0], 1)  # ID
        self.assertEqual(row[1], "John Smith")  # Name
        
        # Test invalid input
        row_invalid, _ = self.excel.read_row("invalid")
        self.assertIsNone(row_invalid)
    
    def test_direct_get_column_index_by_header(self):
        """Test finding column index by header directly."""
        # Find column index
        col_index, message = self.excel.get_column_index_by_header("Department")
        
        # Verify
        self.assertEqual(col_index, 4)
        self.assertIn("found by header", message)
        
        # Test non-existent header
        col_invalid, message_invalid = self.excel.get_column_index_by_header("NonExistent")
        self.assertIsNone(col_invalid)
        self.assertIn("not found", message_invalid)
    
    #
    # JSON API TESTS
    #
    
    def test_json_clear_sheet(self):
        """Test clearing a sheet with JSON."""
        # Clear the sheet
        json_input = json.dumps({
            "function_name": "excel_clear_sheet",
            "parameters": {}
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify sheet is empty
        self.assertEqual(self.excel.sheet.max_row, 1)
    
    def test_json_add_row(self):
        """Test adding a row with JSON."""
        # Add a row
        json_input = json.dumps({
            "function_name": "excel_add_row",
            "parameters": {
                "row_index": 5,
                "text": "New employee row"
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify row was added
        cell_value, _ = self.excel.read_cell(5, 1)
        self.assertEqual(cell_value, "New employee row")
        
        # Test with next_available
        json_input = json.dumps({
            "function_name": "excel_add_row",
            "parameters": {
                "row_index": "next_available",
                "text": "Next available row"
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Test invalid input
        json_invalid = json.dumps({
            "function_name": "excel_add_row",
            "parameters": {
                "row_index": -1,
                "text": "Invalid row"
            }
        })
        
        reward_invalid, feedback_invalid = self.excel.process_json_operation(json_invalid)
        
        # Verify
        self.assertEqual(reward_invalid, -1)
        self.assertIn("Error", feedback_invalid)
    
    def test_json_write_cell(self):
        """Test writing to a cell with JSON."""
        # Write to cell
        json_input = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": 2,
                "col_index": 2,
                "text": "Updated JSON Name"
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify cell was updated
        cell_value, _ = self.excel.read_cell(2, 2)
        self.assertEqual(cell_value, "Updated JSON Name")
        
        # Test with column letter
        json_letter = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": 3,
                "col_index": "B",
                "text": "Column Letter JSON"
            }
        })
        
        reward_letter, _ = self.excel.process_json_operation(json_letter)
        self.assertEqual(reward_letter, 1)
        
        # Test invalid input
        json_invalid = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": "invalid",
                "col_index": 2,
                "text": "Invalid"
            }
        })
        
        reward_invalid, _ = self.excel.process_json_operation(json_invalid)
        self.assertEqual(reward_invalid, -1)
    
    def test_json_write_row(self):
        """Test writing a row with JSON."""
        # Write row
        json_input = json.dumps({
            "function_name": "excel_write_row",
            "parameters": {
                "row_index": 2,
                "row_data": [1, "JSON Updated", 30, "IT", 70000]
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify row was updated
        row_data, _ = self.excel.read_row(2)
        self.assertEqual(row_data[1], "JSON Updated")
        
        # Test invalid input
        json_invalid = json.dumps({
            "function_name": "excel_write_row",
            "parameters": {
                "row_index": "invalid",
                "row_data": [1, 2, 3]
            }
        })
        
        reward_invalid, _ = self.excel.process_json_operation(json_invalid)
        self.assertEqual(reward_invalid, -1)
    
    def test_json_clear_cell(self):
        """Test clearing a cell with JSON."""
        # Clear cell
        json_input = json.dumps({
            "function_name": "excel_clear_cell",
            "parameters": {
                "row_index": 2,
                "col_index": 2
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify cell was cleared
        cell_value, _ = self.excel.read_cell(2, 2)
        self.assertIsNone(cell_value)
    
    def test_json_clear_row(self):
        """Test clearing a row with JSON."""
        # Get data from row 3 before clearing
        row3_before, _ = self.excel.read_row(3)
        
        # Clear row
        json_input = json.dumps({
            "function_name": "excel_clear_row",
            "parameters": {
                "row_index": 2
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify row 3 is now at position 2
        row2_after, _ = self.excel.read_row(2)
        self.assertEqual(row2_after, row3_before)
    
    def test_json_clear_column(self):
        """Test clearing a column with JSON."""
        # Get header before clearing
        header_before, _ = self.excel.read_header_row()
        
        # Clear column
        json_input = json.dumps({
            "function_name": "excel_clear_column",
            "parameters": {
                "col_index": 2
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify Age is now at position 2
        header_after, _ = self.excel.read_header_row()
        self.assertEqual(header_after[1], header_before[2])
    
    def test_json_read_header_row(self):
        """Test reading the header row with JSON."""
        # Read header row
        json_input = json.dumps({
            "function_name": "excel_read_header_row"
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        self.assertIn("ID", feedback)
        self.assertIn("Name", feedback)
    
    def test_json_read_column(self):
        """Test reading a column with JSON."""
        # Read column
        json_input = json.dumps({
            "function_name": "excel_read_column",
            "parameters": {
                "col_index": 2
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        self.assertIn("Name", feedback)
        self.assertIn("John Smith", feedback)
        
        # Test with column letter
        json_letter = json.dumps({
            "function_name": "excel_read_column",
            "parameters": {
                "col_index": "B"
            }
        })
        
        reward_letter, feedback_letter = self.excel.process_json_operation(json_letter)
        
        # Verify
        self.assertEqual(reward_letter, 1)
        self.assertIn("John Smith", feedback_letter)
    
    def test_json_read_cell(self):
        """Test reading a cell with JSON."""
        # Read cell
        json_input = json.dumps({
            "function_name": "excel_read_cell",
            "parameters": {
                "row_index": 2,
                "col_index": 2
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        self.assertIn("John Smith", feedback)
    
    def test_json_read_row(self):
        """Test reading a row with JSON."""
        # Read row
        json_input = json.dumps({
            "function_name": "excel_read_row",
            "parameters": {
                "row_index": 2
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        self.assertIn("John Smith", feedback)
    
    def test_json_get_column_index_by_header(self):
        """Test finding column index by header with JSON."""
        # Get column index
        json_input = json.dumps({
            "function_name": "excel_get_column_index_by_header",
            "parameters": {
                "header_name": "Department"
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        self.assertIn("Column index found by header", feedback)
        self.assertIn("Result: 4", feedback)
        
        # Test non-existent header
        json_invalid = json.dumps({
            "function_name": "excel_get_column_index_by_header",
            "parameters": {
                "header_name": "NonExistent"
            }
        })
        
        reward_invalid, feedback_invalid = self.excel.process_json_operation(json_invalid)
        
        # Verify
        self.assertEqual(reward_invalid, -1)
        self.assertIn("Error", feedback_invalid)
        self.assertIn("not found", feedback_invalid)
    
    def test_json_invalid_function(self):
        """Test invalid function name with JSON."""
        # Invalid function
        json_input = json.dumps({
            "function_name": "excel_invalid_function",
            "parameters": {}
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, -1)
        self.assertIn("Error", feedback)
        self.assertIn("Unknown function", feedback)
    
    def test_json_invalid_format(self):
        """Test invalid JSON format."""
        # Invalid JSON
        json_input = "This is not valid JSON"
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, -1)
        self.assertIn("Error", feedback)
        self.assertIn("Invalid JSON", feedback)
    
    def test_json_missing_function_name(self):
        """Test missing function name in JSON."""
        # Missing function name
        json_input = json.dumps({
            "parameters": {
                "row_index": 1,
                "col_index": 1,
                "text": "Test"
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify
        self.assertEqual(reward, -1)
        self.assertIn("Error", feedback)
        self.assertIn("missing", feedback)

    def test_write_cell_does_not_affect_a1(self):
        """Test that writing to a cell does not affect cell A1."""
        # Set up initial state - write something to A1
        initial_a1_value = "Initial A1 Value"
        self.excel.write_cell(1, 1, initial_a1_value)
        
        # Verify A1 has the value
        a1_value_before, _ = self.excel.read_cell(1, 1)
        self.assertEqual(a1_value_before, initial_a1_value)
        
        # Write to another cell (B2)
        b2_value = "Value in B2"
        self.excel.write_cell(2, 2, b2_value)
        
        # Verify B2 has the right value
        b2_value_after, _ = self.excel.read_cell(2, 2)
        self.assertEqual(b2_value_after, b2_value)
        
        # Verify A1 still has its original value (not modified)
        a1_value_after, _ = self.excel.read_cell(1, 1)
        self.assertEqual(a1_value_after, initial_a1_value, 
                         "Cell A1 was modified when writing to B2")
    
    def test_json_write_cell_does_not_affect_a1(self):
        """Test that writing to a cell via JSON does not affect cell A1."""
        # Set up initial state - write something to A1
        initial_a1_value = "Initial A1 JSON"
        self.excel.write_cell(1, 1, initial_a1_value)
        
        # Verify A1 has the value
        a1_value_before, _ = self.excel.read_cell(1, 1)
        self.assertEqual(a1_value_before, initial_a1_value)
        
        # Write to another cell (B2) using JSON
        json_input = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": 2,
                "col_index": "B",
                "text": "JSON Value in B2"
            }
        })
        
        reward, feedback = self.excel.process_json_operation(json_input)
        
        # Verify operation succeeded
        self.assertEqual(reward, 1)
        self.assertIn("Success", feedback)
        
        # Verify B2 has the right value
        b2_value_after, _ = self.excel.read_cell(2, 2)
        self.assertEqual(b2_value_after, "JSON Value in B2")
        
        # Verify A1 still has its original value (not modified)
        a1_value_after, _ = self.excel.read_cell(1, 1)
        self.assertEqual(a1_value_after, initial_a1_value, 
                         "Cell A1 was modified when writing to B2 via JSON")

    def test_comprehensive_excel_operations(self):
        """A comprehensive test of all Excel operations to ensure they work as expected."""
        # 1. Clear the sheet and verify
        self.excel.clear_sheet()
        header_row, _ = self.excel.read_header_row()
        self.assertEqual(len(header_row), 1, "Sheet should have one empty header row after clearing")
        self.assertEqual(header_row[0], None, "Header row should be empty (None) after clearing")
        
        # 2. Write header row
        self.excel.write_row(1, ["ID", "Name", "Age", "Department"])
        header_row, _ = self.excel.read_header_row()
        self.assertEqual(header_row, ["ID", "Name", "Age", "Department"], "Header row should match what was written")
        
        # 3. Write to specific cells - IMPORTANT TEST for the write issue
        # Write to A1 (which should already have "ID")
        self.excel.write_cell(1, "A", "Employee ID")
        # Write to B2
        self.excel.write_cell(2, "B", "Jane Doe")
        # Write to C3
        self.excel.write_cell(3, "C", 28)
        # Write to D4
        self.excel.write_cell(4, "D", "Finance")
        
        # 4. Read and verify each cell
        a1_value, _ = self.excel.read_cell(1, "A")
        b2_value, _ = self.excel.read_cell(2, "B")
        c3_value, _ = self.excel.read_cell(3, "C")
        d4_value, _ = self.excel.read_cell(4, "D")
        
        self.assertEqual(a1_value, "Employee ID", "A1 should contain 'Employee ID'")
        self.assertEqual(b2_value, "Jane Doe", "B2 should contain 'Jane Doe'")
        self.assertEqual(c3_value, 28, "C3 should contain 28")
        self.assertEqual(d4_value, "Finance", "D4 should contain 'Finance'")
        
        # 5. Test write_row again, ensuring it doesn't affect other cells
        self.excel.write_row(2, [101, "John Smith", 35, "Engineering"])
        
        # Verify row 2 was updated
        row2_data, _ = self.excel.read_row(2)
        self.assertEqual(row2_data, [101, "John Smith", 35, "Engineering"], "Row 2 should match what was written")
        
        # Verify other cells remain unchanged
        a1_check, _ = self.excel.read_cell(1, "A")
        c3_check, _ = self.excel.read_cell(3, "C")
        d4_check, _ = self.excel.read_cell(4, "D")
        
        self.assertEqual(a1_check, "Employee ID", "A1 should still contain 'Employee ID'")
        self.assertEqual(c3_check, 28, "C3 should still contain 28")
        self.assertEqual(d4_check, "Finance", "D4 should still contain 'Finance'")
        
        # 6. Test JSON operations for writing to cells
        json_write_b3 = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": 3,
                "col_index": "B",
                "text": "Sarah Johnson"
            }
        })
        
        reward, _ = self.excel.process_json_operation(json_write_b3)
        self.assertEqual(reward, 1, "JSON write to B3 should succeed")
        
        # Verify B3 was updated
        b3_value, _ = self.excel.read_cell(3, "B")
        self.assertEqual(b3_value, "Sarah Johnson", "B3 should contain 'Sarah Johnson'")
        
        # Verify A1 was not affected by JSON write to B3
        a1_final, _ = self.excel.read_cell(1, "A")
        self.assertEqual(a1_final, "Employee ID", "A1 should still contain 'Employee ID' after JSON write to B3")
        
        # 7. Test column operations
        col_b, _ = self.excel.read_column("B")
        self.assertEqual(col_b[2], "Sarah Johnson", "Column B, row 3 should contain 'Sarah Johnson'")
        
        # 8. Test get_column_index_by_header
        name_col_index, _ = self.excel.get_column_index_by_header("Name")
        self.assertEqual(name_col_index, 2, "Name column should be at index 2 (column B)")
        
        # 9. Test clearing specific cells
        self.excel.clear_cell(3, "B")
        b3_after_clear, _ = self.excel.read_cell(3, "B")
        self.assertIsNone(b3_after_clear, "B3 should be None after clearing")
        
        # Final verification that everything works as expected
        header_final, _ = self.excel.read_header_row()
        self.assertEqual(header_final, ["Employee ID", "Name", "Age", "Department"], 
                         "Header row should still be intact after all operations")

    def test_invalid_inputs_handling(self):
        """Test how the Excel functions handle invalid inputs."""
        # Test invalid row index
        success, _ = self.excel.write_cell("invalid", 1, "Test")
        self.assertFalse(success, "Should fail with invalid row index")
        
        # Test invalid column index
        success, _ = self.excel.write_cell(1, "invalid$column", "Test")
        self.assertFalse(success, "Should fail with invalid column index")
        
        # Test negative indices
        success, _ = self.excel.write_cell(-1, 1, "Test") 
        self.assertFalse(success, "Should fail with negative row index")
        
        success, _ = self.excel.write_cell(1, -1, "Test")
        self.assertFalse(success, "Should fail with negative column index")
        
        # Test JSON invalid inputs
        json_invalid_row = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": "invalid",
                "col_index": "B",
                "text": "Test"
            }
        })
        
        reward, _ = self.excel.process_json_operation(json_invalid_row)
        self.assertEqual(reward, -1, "Should fail with invalid row index in JSON")
        
        json_invalid_col = json.dumps({
            "function_name": "excel_write_cell",
            "parameters": {
                "row_index": 1,
                "col_index": "invalid$$",
                "text": "Test"
            }
        })
        
        reward, _ = self.excel.process_json_operation(json_invalid_col)
        self.assertEqual(reward, -1, "Should fail with invalid column index in JSON")


if __name__ == '__main__':
    print("=== Comprehensive Excel Functions Test Suite ===\n")
    unittest.main(verbosity=2) 