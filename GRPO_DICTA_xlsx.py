import torch
import json
import os
import logging
from transformers import AutoTokenizer, AutoModelForCausalLM, BitsAndBytesConfig
from peft import PeftModel
from excel_functions import ExcelHandler

# Configure logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='llm_excel_test.log'
)
logger = logging.getLogger('excel_llm_test')

# Model configuration
model_id = "dicta-il/dictalm2.0-instruct"
device = "cuda" if torch.cuda.is_available() else "cpu"

# Configure quantization
bnb_config = BitsAndBytesConfig(
    load_in_4bit=True,
    bnb_4bit_use_double_quant=True,
    bnb_4bit_quant_type="nf4",
    bnb_4bit_compute_dtype=torch.bfloat16
)

print(f"Loading model on device: {device}")
model = AutoModelForCausalLM.from_pretrained(
    model_id,
    quantization_config=bnb_config,
    device_map="auto",
    trust_remote_code=True,
)

tokenizer = AutoTokenizer.from_pretrained(model_id, add_bos_token=True, trust_remote_code=True)

# Excel-focused system prompt - focused on write_cell operation
system_prompt = """
You are an Excel operations assistant that helps users work with Excel files.
You can perform operations on Excel files using JSON commands.

IMPORTANT: Your response MUST include a valid JSON command that updates a cell in Excel by looking up the appropriate row and column.

The update_cell_by_lookup operation allows you to find and modify cells by their data context rather than direct cell references.

Respond with a valid JSON object that follows this exact format:
{
    "function_name": "excel_update_cell_by_lookup",
    "parameters": {
        "row_header": "ROW_IDENTIFIER_COLUMN_HEADER",
        "row_value": "VALUE_TO_FIND_IN_ROW_IDENTIFIER_COLUMN",
        "col_header": "TARGET_COLUMN_HEADER",
        "new_value": "NEW_TEXT_TO_WRITE"
    }
}

PARAMETER RULES:
- row_header: Name of a column header that uniquely identifies the row (e.g., "Project ID", "Name", etc.)
- row_value: The value to search for in the row_header column (e.g., "Project Alpha", "John Smith")
- col_header: Name of the column header where the update should happen (e.g., "Description", "Status")
- new_value: String value to write to the cell
- language: The language of the headers and value is in the most time in Hebrew, keep your mind on that.

EXAMPLES:

1. Update the description of Project Alpha:
{
    "function_name": "excel_update_cell_by_lookup",
    "parameters": {
        "row_header": "Project Name",
        "row_value": "Project Alpha",
        "col_header": "Description",
        "new_value": "An educational initiative to improve literacy."
    }
}

2. Change the status of Task ID 1001:
{
    "function_name": "excel_update_cell_by_lookup",
    "parameters": {
        "row_header": "Task ID",
        "row_value": "1001",
        "col_header": "Status",
        "new_value": "Completed"
    }
}

First carefully analyze the user's request to understand:
1. Which column contains the row identifier (row_header)
2. What value to look for in that column (row_value)
3. Which column needs to be updated (col_header)
4. What new value to write (new_value)

Then generate the appropriate JSON command to update the cell based on these context clues rather than direct cell references.
"""

model.resize_token_embeddings(len(tokenizer))
model.config.pad_token_id = tokenizer.pad_token_id  # Set pad token ID


def generate_response(user_input, chat_history=None, excel_handler=None):
    """Generate a response from the model"""
    if chat_history is None:
        chat_history = []

    # Add data context to the prompt if Excel handler is provided
    context_prompt = ""
    if excel_handler:
        context_prompt = "Here is the current Excel structure:\n\n"

        # Function to format rows with headers
        def format_rows_with_headers(headers, start_row, end_row, section_name):
            nonlocal context_prompt  # Allow modification of outer scope variable
            if headers and any(headers):
                context_prompt += f"{section_name.upper()}:\n"
                context_prompt += f"Headers (Row {start_row}): {', '.join([str(h) for h in headers if h])}\n"
                for row_idx in range(start_row + 1, end_row + 1):
                    try:
                        row_data, _ = excel_handler.read_row(row_idx)
                        if row_data and any(row_data):
                            formatted_row = []
                            for i, cell in enumerate(row_data):
                                if cell and headers and i < len(headers) and headers[i]:
                                    formatted_row.append(f"{headers[i]}: {cell}")
                                elif cell:
                                    formatted_row.append(str(cell))
                            context_prompt += f"Row {row_idx}: {', '.join(formatted_row)}\n"
                    except Exception as e:  # Catching specific exceptions is better, but for now keep it broad and log
                        logger.error(f"Error reading row {row_idx} for context: {e}")
                context_prompt += "\n"

        # Add Projects section
        project_headers, _ = excel_handler.read_row(1)
        format_rows_with_headers(project_headers, 1, 4, "Projects")

        # Add Tasks section if it exists
        try:
            task_headers, _ = excel_handler.read_row(6)
            format_rows_with_headers(task_headers, 6, 9, "Tasks")
        except Exception as e:
            logger.error(f"Error reading Task headers for context: {e}")

        # Add Employees section if it exists
        try:
            emp_headers, _ = excel_handler.read_row(11)
            format_rows_with_headers(emp_headers, 11, 14, "Employees")
        except Exception as e:
            logger.error(f"Error reading Employee headers for context: {e}")


    messages = chat_history + [{"role": "user", "content": f"{system_prompt}\n\n{context_prompt}\n\nUser instruction: {user_input}"}]
    encoded = tokenizer.apply_chat_template(messages, return_tensors="pt").to(device)

    model.eval()
    with torch.no_grad():
        outputs = model.generate(
            input_ids=encoded,
            max_new_tokens=500,
            do_sample=True,
            pad_token_id=tokenizer.eos_token_id
        )

        decoded_output = tokenizer.decode(outputs[0], skip_special_tokens=True)

    # Extract the assistant's response
    assistant_response = decoded_output.split("[/INST]")[-1].strip()

    # For Excel command testing, we want to keep the entire response including JSON
    return assistant_response

def extract_json_from_response(response):
    """Try to extract a JSON object from the LLM response"""
    try:
        # Look for JSON pattern
        start_idx = response.find('{')
        end_idx = response.rfind('}')

        if start_idx != -1 and end_idx != -1:
            json_str = response[start_idx:end_idx+1]
            return json.loads(json_str)
        return None
    except json.JSONDecodeError:
        logger.error(f"Failed to parse JSON from response: {response}")
        return None

def create_default_scenarios_file(filename="write_cell_scenarios.json"):
    """Create a default scenarios file if it doesn't exist"""
    if os.path.exists(filename):
        return

    # Default test scenarios focused on cell lookup and update
    default_scenarios = [
        {
            "prompt": "Change the description of Project Alpha to 'Educational initiative for elementary schools'",
            "expected_params": {
                "row_header": "Project Name",
                "row_value": "Project Alpha",
                "col_header": "Description",
                "new_value": "Educational initiative for elementary schools"
            }
        },
        {
            "prompt": "Update the status of Task 1001 to 'Completed'",
            "expected_params": {
                "row_header": "Task ID",
                "row_value": "1001",
                "col_header": "Status",
                "new_value": "Completed"
            }
        },
        {
            "prompt": "Set John Smith's department to 'Marketing'",
            "expected_params": {
                "row_header": "Name",
                "row_value": "John Smith",
                "col_header": "Department",
                "new_value": "Marketing"
            }
        },
        {
            "prompt": "For Project Beta, change the budget to 50000",
            "expected_params": {
                "row_header": "Project Name",
                "row_value": "Project Beta",
                "col_header": "Budget",
                "new_value": "50000"
            }
        },
        {
            "prompt": "Update the deadline for Project Gamma to 'December 31, 2023'",
            "expected_params": {
                "row_header": "Project Name",
                "row_value": "Project Gamma",
                "col_header": "Deadline",
                "new_value": "December 31, 2023"
            }
        }
    ]

    with open(filename, 'w') as f:
        json.dump(default_scenarios, f, indent=4)

    print(f"Created default scenarios file: {filename}")

def load_scenarios(filename="write_cell_scenarios.json"):
    """Load test scenarios from a JSON file"""
    # Create default file if it doesn't exist
    create_default_scenarios_file(filename)

    try:
        with open(filename, 'r') as f:
            scenarios = json.load(f)
        return scenarios
    except Exception as e:
        logger.error(f"Error loading scenarios from {filename}: {str(e)}")
        return []

class WriteExcelTest:
    """Test the LLM's ability to generate update_cell_by_lookup commands"""

    def __init__(self, test_file="write_cell_test.xlsx"):
        """Initialize the test environment"""
        self.test_file = test_file
        if os.path.exists(test_file):
            os.remove(test_file)
        self.excel = ExcelHandler(test_file)
        self.results = []

        # Set up test data
        self._create_test_data()

    def _create_test_data(self):
        """Create sample data for testing - HEBREW VERSION"""
        # Headers for projects - HEBREW
        headers = ["שם פרויקט", "תיאור", "סטטוס", "תקציב", "תאריך יעד", "מנהל"]
        self.excel.write_row(1, headers)

        # Data rows for projects - HEBREW
        project_rows = [
            ["פרויקט אלפא", "יוזמת חינוך בסיסי", "בביצוע", "25000", "15 באוקטובר, 2023", "ג'ון דו"],
            ["פרויקט בטא", "שיפור שירותי בריאות", "בתכנון", "35000", "30 בנובמבר, 2023", "ג'יין סמית'"],
            ["פרויקט גמא", "שימור הסביבה", "בהמתנה", "15000", "15 בינואר, 2024", "בוב ג'ונסון"]
        ]

        for i, row in enumerate(project_rows, 2):
            self.excel.write_row(i, row)

        # Tasks section - HEBREW
        self.excel.write_row(6, ["מזהה משימה", "שם משימה", "סטטוס", "משויך ל", "עדיפות"])

        task_rows = [
            ["1001", "יצירת תוכנית פרויקט", "בביצוע", "אליס בראון", "גבוהה"],
            ["1002", "פגישת בעלי עניין", "לא התחיל", "ג'ון דו", "בינונית"],
            ["1003", "אישור תקציב", "הושלם", "ג'יין סמית'", "גבוהה"]
        ]

        for i, row in enumerate(task_rows, 7):
            self.excel.write_row(i, row)

        # Employees section - HEBREW
        self.excel.write_row(11, ["מזהה", "שם", "מחלקה", "תפקיד", "תאריך גיוס"])

        employee_rows = [
            ["E001", "ג'ון סמית'", "מחשוב", "מפתח תוכנה", "2020-01-15"],
            ["E002", "מרי ג'ונסון", "משאבי אנוש", "מנהל", "2018-05-20"],
            ["E003", "רוברט דייויס", "כספים", "אנליסט", "2021-03-10"]
        ]

        for i, row in enumerate(employee_rows, 12):
            self.excel.write_row(i, row)

    def run_test_case(self, test_case, expected_params):
        """Run a single test case and evaluate the result"""
        logger.info(f"Running test case: {test_case}")

        # Generate response from the model
        response = generate_response(test_case, excel_handler=self.excel)
        logger.info(f"Model response: {response[:100]}..." if len(response) > 100 else response)

        # Extract JSON from response
        json_data = extract_json_from_response(response)

        result = {
            "test_case": test_case,
            "response": response,
            "expected_params": expected_params,
            "json_extracted": json_data is not None,
            "correct_function": False,
            "correct_params": False,
            "excel_success": False,
            "excel_feedback": ""
        }

        if not json_data:
            logger.warning("No valid JSON found in response")
            self.results.append(result)
            return result

        # Check if the function name is correct
        if "function_name" in json_data and json_data["function_name"] == "excel_update_cell_by_lookup":
            result["correct_function"] = True

        # Check if the parameters are correct
        if "parameters" in json_data:
            params_match = True
            for key, value in expected_params.items():
                if key not in json_data["parameters"] or str(json_data["parameters"][key]) != str(value):
                    params_match = False
                    logger.info(f"Parameter mismatch: expected {key}={value}, got {json_data['parameters'].get(key, 'missing')}")
                    break
            result["correct_params"] = params_match

        # Execute the Excel operation
        try:
            json_str = json.dumps(json_data)
            reward, feedback = self.excel.process_json_operation(json_str)
            result["excel_success"] = (reward == 1)
            result["excel_feedback"] = feedback
            logger.info(f"Excel operation result: reward={reward}, feedback={feedback}")
        except Exception as e:
            logger.error(f"Error executing Excel operation: {str(e)}")
            result["excel_feedback"] = f"Error: {str(e)}"

        self.results.append(result)
        return result

    def run_all_tests(self, scenarios):
        """Run all test cases from the scenarios list"""
        print("\n--- Automated Test Results ---") # Added section title for clarity
        for i, scenario in enumerate(scenarios):
            prompt = scenario["prompt"]
            expected_params = scenario["expected_params"]

            print(f"\nTest {i+1}/{len(scenarios)}: {prompt}") # Keep test case prompt visible
            result = self.run_test_case(prompt, expected_params)

            # Print result
            json_status = "✅" if result["json_extracted"] else "❌"
            function_status = "✅" if result["correct_function"] else "❌"
            params_status = "✅" if result["correct_params"] else "❌"
            excel_status = "✅" if result["excel_success"] else "❌"

            print(f"  JSON Extracted: {json_status}")
            print(f"  Correct Function: {function_status}")
            print(f"  Correct Parameters: {params_status}")
            print(f"  Excel Success: {excel_status}")
            print(f"  Feedback: {result['excel_feedback']}")

        print("\n--- End of Automated Test Results ---") # End section title for clarity


    def calculate_metrics(self):
        """Calculate performance metrics"""
        total_tests = len(self.results)
        if total_tests == 0:
            return {"total_tests": 0}

        json_extraction_success = sum(1 for r in self.results if r["json_extracted"])
        correct_function = sum(1 for r in self.results if r["correct_function"])
        correct_params = sum(1 for r in self.results if r["correct_params"])
        excel_success = sum(1 for r in self.results if r["excel_success"])

        metrics = {
            "total_tests": total_tests,
            "json_extraction_rate": json_extraction_success / total_tests,
            "function_accuracy": correct_function / total_tests,
            "parameter_accuracy": correct_params / total_tests,
            "excel_success_rate": excel_success / total_tests
        }

        return metrics

    def cleanup(self):
        """Clean up test resources"""
        self.excel.workbook.close()
        if os.path.exists(self.test_file):
            os.remove(self.test_file)
        logger.info("Test resources cleaned up")

def run_automated_tests():
    """Run automated tests using scenarios from JSON file"""
    scenarios = load_scenarios()
    if not scenarios:
        print("No test scenarios found. Please check the scenarios file.")
        return

    tester = WriteExcelTest()

    # Display a preview of the test data - HEBREW HEADERS in description
    print("\nTest data structure:")
    print("1. Projects (rows 1-4):")
    print("   Headers: שם פרויקט, תיאור, סטטוס, תקציב, תאריך יעד, מנהל")
    print("   Row 2: פרויקט אלפא - יוזמת חינוך בסיסי (סטטוס: בביצוע)")
    print("   Row 3: פרויקט בטא - שיפור שירותי בריאות (סטטוס: בתכנון)")
    print("   Row 4: פרויקט גמא - שימור הסביבה (סטטוס: בהמתנה)")

    print("\n2. Tasks (rows 6-9):")
    print("   Headers: מזהה משימה, שם משימה, סטטוס, משויך ל, עדיפות")
    print("   Row 7: 1001 - יצירת תוכנית פרויקט (סטטוס: בביצוע)")
    print("   Row 8: 1002 - פגישת בעלי עניין (סטטוס: לא התחיל)")
    print("   Row 9: 1003 - אישור תקציב (סטטוס: הושלם)")

    print("\n3. Employees (rows 11-14):")
    print("   Headers: מזהה, שם, מחלקה, תפקיד, תאריך גיוס")
    print("   Row 12: E001 - ג'ון סמית' - מחשוב (תפקיד: מפתח תוכנה)")
    print("   Row 13: E002 - מרי ג'ונסון - משאבי אנוש (תפקיד: מנהל)")
    print("   Row 14: E003 - רוברט דייויס - כספים (תפקיד: אנליסט)")

    try:
        print(f"\nRunning {len(scenarios)} cell lookup and update test scenarios...")
        tester.run_all_tests(scenarios)

        # Print summary metrics
        metrics = tester.calculate_metrics()
        print("\n--- Overall Test Summary ---") # Added section title for clarity
        print(f"Total Tests: {metrics['total_tests']}")
        print(f"JSON Extraction Rate: {metrics['json_extraction_rate']:.2%}")
        print(f"Function Accuracy: {metrics['function_accuracy']:.2%}")
        print(f"Parameter Accuracy: {metrics['parameter_accuracy']:.2%}")
        print(f"Excel Success Rate: {metrics['excel_success_rate']:.2%}")
        print("\n--- End of Test Summary ---") # End section title for clarity
    finally:
        tester.cleanup()

def run_interactive_test():
    """Run interactive testing with the model - HEBREW DATA"""
    test_file = "interactive_test.xlsx"
    if os.path.exists(test_file):
        os.remove(test_file)

    # Create a new test Excel file with sample data
    test_excel = ExcelHandler(test_file)

    # Create sample data for projects - HEBREW
    print("\nCreating sample data for testing...")

    # Headers - HEBREW
    headers = ["שם פרויקט", "תיאור", "סטטוס", "תקציב", "תאריך יעד", "מנהל"]
    test_excel.write_row(1, headers)

    # Data rows - HEBREW
    data_rows = [
        ["פרויקט אלפא", "יוזמת חינוך בסיסי", "בביצוע", "25000", "15 באוקטובר, 2023", "ג'ון דו"],
        ["פרויקט בטא", "שיפור שירותי בריאות", "בתכנון", "35000", "30 בנובמבר, 2023", "ג'יין סמית'"],
        ["פרויקט גמא", "שימור הסביבה", "בהמתנה", "15000", "15 בינואר, 2024", "בוב ג'ונסון"]
    ]

    for i, row in enumerate(data_rows, 2):
        test_excel.write_row(i, row)

    # Create another section for tasks - HEBREW
    test_excel.write_row(6, ["מזהה משימה", "שם משימה", "סטטוס", "משויך ל", "עדיפות"])

    task_rows = [
        ["1001", "יצירת תוכנית פרויקט", "בביצוע", "אליס בראון", "גבוהה"],
        ["1002", "פגישת בעלי עניין", "לא התחיל", "ג'ון דו", "בינונית"],
        ["1003", "אישור תקציב", "הושלם", "ג'יין סמית'", "גבוהה"]
    ]

    for i, row in enumerate(task_rows, 7):
        test_excel.write_row(i, row)

    # Create another section for employees - HEBREW
    test_excel.write_row(11, ["מזהה", "שם", "מחלקה", "תפקיד", "תאריך גיוס"])

    employee_rows = [
        ["E001", "ג'ון סמית'", "מחשוב", "מפתח תוכנה", "2020-01-15"],
        ["E002", "מרי ג'ונסון", "משאבי אנוש", "מנהל", "2018-05-20"],
        ["E003", "רוברט דייויס", "כספים", "אנליסט", "2021-03-10"]
    ]

    for i, row in enumerate(employee_rows, 12):
        test_excel.write_row(i, row)

    print("Sample data created successfully! (Hebrew)")

    # Display the structure of the data - HEBREW HEADERS in description
    print("\nData structure:")
    print("1. Projects (rows 1-4):")
    headers_str = ", ".join(headers)
    print(f"   Headers: {headers_str}")
    for i, row in enumerate(data_rows, 2):
        print(f"   Row {i}: {row[0]} - {row[1]} (סטטוס: {row[2]})") # סטטוס in Hebrew

    print("\n2. Tasks (rows 6-9):")
    task_headers = ["מזהה משימה", "שם משימה", "סטטוס", "משויך ל", "עדיפות"]
    print(f"   Headers: {', '.join(task_headers)}")
    for i, row in enumerate(task_rows, 7):
        print(f"   Row {i}: {row[0]} - {row[1]} (סטטוס: {row[2]})") # סטטוס in Hebrew

    print("\n3. Employees (rows 11-14):")
    employee_headers = ["מזהה", "שם", "מחלקה", "תפקיד", "תאריך גיוס"]
    print(f"   Headers: {', '.join(employee_headers)}")
    for i, row in enumerate(employee_rows, 12):
        print(f"   Row {i}: {row[1]} - {row[2]} (תפקיד: {row[3]})") # תפקיד in Hebrew

    print("\nExcel Cell Update Interactive Test (Hebrew Data)")
    print("Type 'exit' to quit")
    print("Type 'show' to display current data")
    print("Type 'debug' to see what's being sent to the LLM")

    chat_history = []
    debug_mode = False

    while True:
        user_input = input("\nEnter your instruction (e.g., 'Change Project Alpha's status to Completed'): ") # Keep instructions in English for now

        if user_input.lower() == 'exit':
            break

        if user_input.lower() == 'debug':
            debug_mode = not debug_mode
            print(f"Debug mode: {'ON' if debug_mode else 'OFF'}")
            continue

        if user_input.lower() == 'show':
            # Display current data
            print("\nCurrent Data:")

            # Show Projects
            print("\nProjects:")
            project_headers, _ = test_excel.read_row(1)
            print(f"   Headers: {', '.join([str(h) for h in project_headers if h])}")
            for row_idx in range(2, 5):
                row_data, _ = test_excel.read_row(row_idx)
                # Show all columns for each row clearly
                formatted_row = []
                for i, cell in enumerate(row_data):
                    if cell and project_headers and i < len(project_headers) and project_headers[i]:
                        formatted_row.append(f"{project_headers[i]}: {cell}")
                    elif cell:
                        formatted_row.append(str(cell))
                print(f"   Row {row_idx}: {', '.join(formatted_row)}")

            # Show Tasks
            print("\nTasks:")
            task_headers, _ = test_excel.read_row(6)
            print(f"   Headers: {', '.join([str(h) for h in task_headers if h])}")
            for row_idx in range(7, 10):
                row_data, _ = test_excel.read_row(row_idx)
                # Show all columns for each row clearly
                formatted_row = []
                for i, cell in enumerate(row_data):
                    if cell and task_headers and i < len(task_headers) and task_headers[i]:
                        formatted_row.append(f"{task_headers[i]}: {cell}")
                    elif cell:
                        formatted_row.append(str(cell))
                print(f"   Row {row_idx}: {', '.join(formatted_row)}")

            # Show Employees
            print("\nEmployees:")
            emp_headers, _ = test_excel.read_row(11)
            print(f"   Headers: {', '.join([str(h) for h in emp_headers if h])}")
            for row_idx in range(12, 15):
                row_data, _ = test_excel.read_row(row_idx)
                # Show all columns for each row clearly
                formatted_row = []
                for i, cell in enumerate(row_data):
                    if cell and emp_headers and i < len(emp_headers) and emp_headers[i]:
                        formatted_row.append(f"{emp_headers[i]}: {cell}")
                    elif cell:
                        formatted_row.append(str(cell))
                print(f"   Row {row_idx}: {', '.join(formatted_row)}")
            continue

        print(f"\nProcessing: '{user_input}'")

        # Create context for the LLM
        context_prompt = "Here is the current Excel structure:\n\n"

        # Function to format rows with headers for interactive mode (duplicated from generate_response, consider refactoring if used more widely)
        def format_rows_with_headers_interactive(headers, start_row, end_row, section_name):
            nonlocal context_prompt
            if headers and any(headers):
                context_prompt += f"{section_name.upper()}:\n"
                context_prompt += f"Headers (Row {start_row}): {', '.join([str(h) for h in headers if h])}\n"
                for row_idx in range(start_row + 1, end_row + 1):
                    try:
                        row_data, _ = test_excel.read_row(row_idx)
                        if row_data and any(row_data):
                            formatted_row = []
                            for i, cell in enumerate(row_data):
                                if cell and headers and i < len(headers) and headers[i]:
                                    formatted_row.append(f"{headers[i]}: {cell}")
                                elif cell:
                                    formatted_row.append(str(cell))
                            context_prompt += f"Row {row_idx}: {', '.join(formatted_row)}\n"
                    except Exception as e:
                        logger.error(f"Error reading row {row_idx} for interactive context: {e}")
                context_prompt += "\n"

        # Add Projects section
        project_headers, _ = test_excel.read_row(1)
        format_rows_with_headers_interactive(project_headers, 1, 4, "Projects")

        # Add Tasks section
        try:
            task_headers, _ = test_excel.read_row(6)
            format_rows_with_headers_interactive(task_headers, 6, 9, "Tasks")
        except Exception as e:
            logger.error(f"Error reading Task headers for interactive context: {e}")


        # Add Employees section
        try:
            emp_headers, _ = test_excel.read_row(11)
            format_rows_with_headers_interactive(emp_headers, 11, 14, "Employees")
        except Exception as e:
            logger.error(f"Error reading Employee headers for interactive context: {e}")


        if debug_mode:
            print("\n=== FULL LLM PROMPT ===")
            print("--- SYSTEM PROMPT ---")
            print(system_prompt)
            print("\n--- DATA CONTEXT ---")
            print(context_prompt)
            print("\n--- USER INSTRUCTION ---")
            print(user_input)
            print("=== END PROMPT ===\n")

        # Generate response from the model
        response = generate_response(user_input, chat_history, test_excel)
        print("\nLLM Response:", response)

        # Try to execute the command if it contains JSON
        json_data = extract_json_from_response(response)
        if json_data:
            print("\nDetected JSON command:")
            print(json.dumps(json_data, indent=2))

            print("\nExecuting command...")
            try:
                json_str = json.dumps(json_data)
                reward, feedback = test_excel.process_json_operation(json_str)
                print(f"Result: {'✅ Success' if reward == 1 else '❌ Failed'}")
                print(f"Feedback: {feedback}")
            except Exception as e:
                print(f"Error executing command: {str(e)}")
        else:
            print("\nNo valid JSON command found in the response.")

        # Update chat history
        chat_history.append({"role": "user", "content": user_input})
        chat_history.append({"role": "assistant", "content": response})

    # Clean up
    test_excel.workbook.close()
    if os.path.exists(test_file):
        os.remove(test_file)
    print("\nTest complete. Excel file cleaned up.")

def add_new_test_scenario():
    """Add a new test scenario to the JSON file"""
    scenarios = load_scenarios()

    print("\nAdd New Test Scenario\n")
    prompt = input("Enter the prompt (e.g., 'Change Project Alpha's status to Completed'): ")
    row_header = input("Row identifier column header (e.g., 'Project Name'): ")
    row_value = input("Value to find in row identifier column (e.g., 'Project Alpha'): ")
    col_header = input("Column header to update (e.g., 'Status'): ")
    new_value = input("New value to write (e.g., 'Completed'): ")

    new_scenario = {
        "prompt": prompt,
        "expected_params": {
            "row_header": row_header,
            "row_value": row_value,
            "col_header": col_header,
            "new_value": new_value
        }
    }

    scenarios.append(new_scenario)

    # Save updated scenarios
    with open("write_cell_scenarios.json", 'w') as f:
        json.dump(scenarios, f, indent=4)

    print(f"Added new scenario. Total scenarios: {len(scenarios)}")

if __name__ == "__main__":
    print("Excel Cell Lookup & Update Test Suite")
    print("1. Run automated tests (from scenarios.json)")
    print("2. Run interactive test (Hebrew Data)") # Indicate Hebrew data in menu
    print("3. Add new test scenario")

    choice = input("Enter your choice (1-3): ")

    if choice == "1":
        run_automated_tests()
    elif choice == "2":
        run_interactive_test() # Run the modified interactive test
    elif choice == "3":
        add_new_test_scenario()
    else:
        print("Invalid choice. Exiting.")