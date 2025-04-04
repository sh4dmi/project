import torch
import json
import os
import logging
# LLM imports are kept for other options, but not used in modified Option 1
from transformers import AutoTokenizer, AutoModelForCausalLM, BitsAndBytesConfig
# from peft import PeftModel # Not used in original loading, commented out
from excel_functions import ExcelHandler
import shutil

# --- Logging Configuration (Unchanged) ---
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    filename='llm_excel_test.log'
)
logger = logging.getLogger('excel_llm_test')

# --- LLM Configuration & Loading (Kept for options 2/3, but skipped if only running option 1) ---
# These are loaded lazily now only if needed for interactive mode
model = None
tokenizer = None
model_loaded = False

def load_llm_resources():
    global model, tokenizer, model_loaded
    if model_loaded:
        return

    logger.info("LLM Resources not loaded yet. Loading...")
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

    print(f"Loading model '{model_id}' on device: {device}...")
    try:
        model = AutoModelForCausalLM.from_pretrained(
            model_id,
            quantization_config=bnb_config,
            device_map="auto",
            trust_remote_code=True,
        )
        tokenizer = AutoTokenizer.from_pretrained(model_id, add_bos_token=True, trust_remote_code=True)

        model.resize_token_embeddings(len(tokenizer))
        model.config.pad_token_id = tokenizer.pad_token_id  # Set pad token ID
        model_loaded = True
        print("Model and Tokenizer loaded successfully.")
        logger.info("LLM Resources loaded.")
    except Exception as e:
        print(f"Error loading LLM resources: {e}")
        logger.error(f"Failed to load LLM model/tokenizer: {e}", exc_info=True)
        # Exit or handle appropriately if LLM is critical for the chosen option
        # For now, we allow proceeding as Option 1 doesn't need it.


# --- System Prompt (Kept for options 2/3) ---
system_prompt = """
You are an Excel operations assistant... [REST OF YOUR PROMPT]...
"""

# --- generate_response Function (Kept for options 2/3) ---
def generate_response(user_input, chat_history=None, excel_handler=None):
    """Generate a response from the model (Used for Interactive Mode)"""
    if not model_loaded:
        load_llm_resources() # Load if not already loaded
        if not model_loaded: # Check again if loading failed
             return "Error: LLM Model could not be loaded. Cannot generate response."

    # [Your existing generate_response logic remains here]
    # ... (Ensure it uses the global 'model' and 'tokenizer')
    # ... (Make sure context prompt generation is still correct)

    if chat_history is None:
        chat_history = []

    context_prompt = ""
    if excel_handler:
        # ... (your context prompt generation logic) ...
        context_prompt = "Here is the current Excel structure:\n\n"

        # Function to format rows with headers
        def format_rows_with_headers(headers, start_row, end_row, section_name):
            nonlocal context_prompt
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
                    except Exception as e:
                        logger.error(f"Error reading row {row_idx} for context: {e}")
                context_prompt += "\n"

        # Add sections (adjust row numbers if your test data is different)
        try:
            project_headers, _ = excel_handler.read_row(1)
            format_rows_with_headers(project_headers, 1, 4, "Projects") # Example range
        except Exception as e: logger.error(f"Ctx Prj Err: {e}")

        try:
            task_headers, _ = excel_handler.read_row(6)
            format_rows_with_headers(task_headers, 6, 9, "Tasks") # Example range
        except Exception as e: logger.error(f"Ctx Tsk Err: {e}")

        try:
            emp_headers, _ = excel_handler.read_row(11)
            format_rows_with_headers(emp_headers, 11, 14, "Employees") # Example range
        except Exception as e: logger.error(f"Ctx Emp Err: {e}")

    messages = chat_history + [{"role": "user", "content": f"{system_prompt}\n\n{context_prompt}\n\nUser instruction: {user_input}"}]
    encoded = tokenizer.apply_chat_template(messages, return_tensors="pt").to(model.device) # Use model.device

    model.eval()
    with torch.no_grad():
        outputs = model.generate(
            input_ids=encoded,
            max_new_tokens=500,
            do_sample=True,
            pad_token_id=tokenizer.eos_token_id
        )
        decoded_output = tokenizer.decode(outputs[0], skip_special_tokens=True)

    assistant_response = decoded_output.split("[/INST]")[-1].strip()
    return assistant_response


# --- extract_json_from_response Function (Kept for options 2/3) ---
def extract_json_from_response(response):
    # [Your existing extract_json_from_response logic remains here]
    try:
        start_idx = response.find('{')
        end_idx = response.rfind('}')
        if start_idx != -1 and end_idx != -1:
            json_str = response[start_idx:end_idx+1]
            return json.loads(json_str)
        return None
    except json.JSONDecodeError:
        logger.error(f"Failed to parse JSON from response: {response}")
        return None


# --- Scenario Management Functions (MODIFIED) ---
# Removed create_default_scenarios_file as it created the old format
# Now loads the format created by convert_scenarios.py

def load_scenarios(filename="converted_scenarios.json"): # Default to the converted file
    """Load test scenarios from the JSON file generated by convert_scenarios.py"""
    if not os.path.exists(filename):
        logger.error(f"Scenario file '{filename}' not found.")
        print(f"Error: Scenario file '{filename}' not found.")
        print(f"Please create it first, for example by running 'python convert_scenarios.py input_data.json {filename}'")
        return []

    try:
        with open(filename, 'r', encoding='utf-8') as f: # Ensure utf-8 reading
            scenarios = json.load(f)
        # Basic validation of the loaded structure
        if isinstance(scenarios, list) and all(isinstance(s, dict) and 'instruction' in s and 'command' in s for s in scenarios):
             logger.info(f"Loaded {len(scenarios)} scenarios from {filename}")
             return scenarios
        else:
            logger.error(f"Invalid format in scenario file {filename}. Expected list of {{'instruction': '...', 'command': {{...}}}}")
            print(f"Error: Invalid format in scenario file {filename}.")
            return []
    except json.JSONDecodeError as e:
        logger.error(f"Error decoding JSON from {filename}: {str(e)}")
        print(f"Error: Could not parse JSON in {filename}. Please check the file format.")
        return []
    except Exception as e:
        logger.error(f"Error loading scenarios from {filename}: {str(e)}")
        print(f"Error: Failed to load scenarios from {filename}.")
        return []

# --- WriteExcelTest Class (MODIFIED for Option 1) ---
class WriteExcelTest:
    """Test the ExcelHandler's ability to execute pre-defined JSON commands"""

    def __init__(self, test_file="write_cell_test.xlsx"):
        self.test_file = test_file
        # Create a backup if the file exists, otherwise create from scratch
        self.backup_file = None
        if os.path.exists(self.test_file):
             self.backup_file = self.test_file + ".bak"
             shutil.copyfile(self.test_file, self.backup_file)
             logger.info(f"Backed up existing test file to {self.backup_file}")
             # Use the existing file
             self.excel = ExcelHandler(self.test_file)
        else:
             logger.info(f"Test file {self.test_file} not found. Creating new one with sample data.")
             self.excel = ExcelHandler(self.test_file)
             # Set up test data only if file is new
             self._create_test_data() # Using HEBREW version

        self.results = []


    def _create_test_data(self):
        """Create sample data for testing - HEBREW VERSION"""
        # Headers for projects - HEBREW
        headers = ["שם פרויקט", "תיאור", "סטטוס", "תקציב", "תאריך יעד", "מנהל", "תוספת אומדן שיקלי", "אומדן ציוד תקשוב שיקלי", "מס' פרויקט מולטימדיה", "תאריך פתיחת פרויקט"] # Added columns from example
        self.excel.write_row(1, headers)

        # Data rows for projects - HEBREW - Adjusted to include example project names/values
        project_rows = [
            ["פרויקט אלפא", "יוזמת חינוך בסיסי", "בביצוע", "25000", "15 באוקטובר, 2023", "ג'ון דו", "1000", "50000", "1111", "01/01/2023"],
            ["פרויקט בטא", "שיפור שירותי בריאות", "בתכנון", "35000", "30 בנובמבר, 2023", "ג'יין סמית'", "0", "60000", "2222", "02/01/2023"],
            ["פרויקט גמא", "שימור הסביבה", "בהמתנה", "15000", "15 בינואר, 2024", "בוב ג'ונסון", "500", "45000", "3333", "03/01/2023"],
            ["התקנת מערכות אזעקה מדויק", "תיאור כלשהו", "בתכנון", "40000", "01/12/2024", "אנה לוי", "15000", "70000", "4444", "04/01/2023"],
            ["פיתוח מערכת שליטה ובקרה ניסיוני יישובי", "פרויקט גדול", "בביצוע", "500000", "01/06/2025", "משה כהן", "5000", "600000", "5555", "05/01/2023"],
            ["הקמת חדר בקרה מבצעי מרכזי יישובי", "פרויקט אסטרטגי", "לא התחיל", "300000", "01/09/2024", "דוד יצחק", "2000", "550000", "6666", "06/01/2023"],
            ["שדרוג מערכות מיזוג אוויר רב-שנתי - בסיס 34", "תחזוקה", "בהמתנה", "120000", "31/12/2024", "שרה גל", "0", "0", "5900", "07/01/2023"], # Initial MM project number
            ["החלפת תשתיות חשמל מהיר", "תשתיות חיוניות", "בתכנון", "90000", "01/03/2024", "יוסי פרץ", "1000", "10000", "7777", "08/01/2023"],
        ]

        for i, row in enumerate(project_rows, 2):
            self.excel.write_row(i, row)

        # NOTE: Removed Task/Employee sections for simplicity, as the example JSON
        # only targeted the project data structure implied by its headers/values.
        # Add them back if your converted_scenarios.json targets them.
        logger.info("Created test data in Excel file.")


    def run_test_case(self, instruction, json_command):
        """Run a single test case using a pre-defined JSON command."""
        logger.info(f"Running test for instruction: {instruction}")
        logger.info(f"Using JSON command: {json.dumps(json_command, ensure_ascii=False)}")

        # Simulate result structure, but bypass LLM generation/extraction checks
        result = {
            "instruction": instruction,
            "command_provided": json_command,
            "valid_command_structure": False, # Assume false initially
            "excel_success": False,
            "excel_feedback": "Command not executed yet."
        }

        # Basic check of the command structure
        if isinstance(json_command, dict) and \
           "function_name" in json_command and \
           json_command["function_name"] == "excel_update_cell_by_lookup" and \
           "parameters" in json_command and \
           isinstance(json_command["parameters"], dict) and \
           all(key in json_command["parameters"] for key in ["row_header", "row_value", "col_header", "new_value"]):
            result["valid_command_structure"] = True
        else:
            logger.warning(f"Invalid command structure provided: {json_command}")
            result["excel_feedback"] = "Invalid command structure provided."
            self.results.append(result)
            return result # Don't attempt to execute invalid structure

        # Execute the Excel operation directly using the provided command
        try:
            json_str = json.dumps(json_command, ensure_ascii=False) # Ensure Hebrew chars are okay
            reward, feedback = self.excel.process_json_operation(json_str)
            result["excel_success"] = (reward == 1)
            result["excel_feedback"] = feedback
            logger.info(f"Excel operation result: reward={reward}, feedback={feedback}")
        except Exception as e:
            logger.error(f"Error executing Excel operation: {str(e)}", exc_info=True)
            result["excel_feedback"] = f"Runtime Error during Excel execution: {str(e)}"
            result["excel_success"] = False # Ensure failure on exception

        self.results.append(result)
        return result

    def run_all_tests(self, scenarios):
        """Run all test cases from the scenarios list (using pre-defined commands)."""
        print("\n--- Automated Excel Command Execution Test Results ---")
        if not scenarios:
            print("No scenarios loaded to test.")
            return

        for i, scenario in enumerate(scenarios):
            instruction = scenario.get("instruction", "No instruction provided")
            command = scenario.get("command", None)

            print(f"\nTest {i+1}/{len(scenarios)}: {instruction}")
            if command is None:
                print("  ❌ Error: Scenario missing 'command' object.")
                # Log this as a failure in results?
                self.results.append({
                    "instruction": instruction, "command_provided": None,
                    "valid_command_structure": False, "excel_success": False,
                    "excel_feedback": "Scenario missing 'command' object."
                 })
                continue

            result = self.run_test_case(instruction, command)

            # Print result
            structure_status = "✅" if result["valid_command_structure"] else "❌"
            excel_status = "✅" if result["excel_success"] else "❌"

            print(f"  Command Structure Valid: {structure_status}")
            print(f"  Excel Execution Success: {excel_status}")
            # Provide more detail on failure
            if not result["excel_success"]:
                print(f"  Feedback: {result['excel_feedback']}")
            elif result.get("excel_feedback"): # Show feedback even on success if provided
                 print(f"  Feedback: {result['excel_feedback']}")


        print("\n--- End of Automated Test Results ---")


    def calculate_metrics(self):
        """Calculate performance metrics based on command execution."""
        total_tests = len(self.results)
        if total_tests == 0:
            return {"total_tests": 0, "valid_command_rate": 0, "excel_success_rate": 0}

        valid_commands = sum(1 for r in self.results if r["valid_command_structure"])
        excel_success = sum(1 for r in self.results if r["excel_success"]) # Success implies valid structure was attempted

        metrics = {
            "total_tests": total_tests,
            "valid_command_rate": valid_commands / total_tests if total_tests > 0 else 0,
            "excel_success_rate": excel_success / total_tests if total_tests > 0 else 0,
            # Could add: success rate among valid commands
            "excel_success_rate_given_valid": excel_success / valid_commands if valid_commands > 0 else 0
        }
        return metrics

    def cleanup(self):
        """Clean up test resources, restore backup if necessary."""
        if self.excel and self.excel.workbook:
            self.excel.workbook.close()
            logger.info("Closed Excel workbook.")

        # Restore from backup if one was made
        if self.backup_file and os.path.exists(self.backup_file):
            try:
                os.remove(self.test_file) # Remove the modified file
                os.rename(self.backup_file, self.test_file) # Restore backup
                logger.info(f"Restored original test file from {self.backup_file}")
            except Exception as e:
                logger.error(f"Failed to restore backup: {e}")
                print(f"Warning: Failed to restore {self.test_file} from backup {self.backup_file}")
        elif os.path.exists(self.test_file) and self.backup_file is None:
            # If no backup was made (meaning file was created fresh), remove it
            try:
                os.remove(self.test_file)
                logger.info(f"Removed newly created test file: {self.test_file}")
            except Exception as e:
                 logger.error(f"Failed to remove test file {self.test_file}: {e}")


# --- run_automated_tests Function (MODIFIED) ---
def run_automated_tests(scenario_file="converted_scenarios.json", excel_file="write_cell_test.xlsx"):
    """Run automated tests using pre-defined JSON commands from file"""
    scenarios = load_scenarios(scenario_file)
    if not scenarios:
        # Error message already printed in load_scenarios
        return

    # Create the tester - it handles backup/creation of the excel_file
    tester = WriteExcelTest(excel_file)

    # Display a preview of the test data structure used (optional but helpful)
    print("\nTest data structure in use (example, actual rows may vary):")
    # You might want to read the actual headers/first few rows from tester.excel here
    # For now, using the hardcoded Hebrew headers from _create_test_data
    print("   Headers: שם פרויקט, תיאור, סטטוס, תקציב, תאריך יעד, מנהל, תוספת אומדן שיקלי, ...")
    print("   Sample Row: פרויקט אלפא, יוזמת חינוך בסיסי, בביצוע, ...")
    print("   Sample Row: התקנת מערכות אזעקה מדויק, תיאור כלשהו, בתכנון, ...")


    try:
        print(f"\nRunning {len(scenarios)} Excel command execution tests from '{scenario_file}' on '{excel_file}'...")
        tester.run_all_tests(scenarios)

        # Print summary metrics
        metrics = tester.calculate_metrics()
        print("\n--- Overall Test Summary ---")
        print(f"Total Tests Attempted: {metrics['total_tests']}")
        print(f"Valid Command Structure Rate: {metrics['valid_command_rate']:.2%}")
        print(f"Overall Excel Execution Success Rate: {metrics['excel_success_rate']:.2%}")
        print(f"Success Rate (given valid command): {metrics['excel_success_rate_given_valid']:.2%}")
        print("\n--- End of Test Summary ---")

    except Exception as e:
        logger.error(f"An error occurred during automated test execution: {e}", exc_info=True)
        print(f"\nAn unexpected error occurred: {e}")
    finally:
        # Cleanup (close file, restore backup)
        tester.cleanup()


# --- run_interactive_test Function (Requires LLM - Unchanged functionally, but uses lazy loading) ---
def run_interactive_test():
    """Run interactive testing with the model - HEBREW DATA"""
    # Ensure LLM is loaded for this mode
    load_llm_resources()
    if not model_loaded:
        print("Cannot start interactive test without loaded LLM resources.")
        return

    # [Your existing run_interactive_test logic remains here]
    # ... (Make sure it calls generate_response correctly)
    # ... (Make sure it creates and uses its own ExcelHandler instance)
    # ... (Make sure it calls cleanup on its specific test file)
    test_file = "interactive_test.xlsx"
    if os.path.exists(test_file):
        os.remove(test_file)

    test_excel = ExcelHandler(test_file)
    print("\nCreating sample data for interactive testing...")
    # Use the same _create_test_data structure or a simplified one
    # For consistency, let's reuse the structure from WriteExcelTest
    temp_tester = WriteExcelTest(test_file) # Use it to create data
    temp_tester.excel.workbook.close() # Close the file handle used by temp_tester
    del temp_tester # Don't need it anymore

    # Re-open with the main handler for interactive mode
    test_excel = ExcelHandler(test_file) # Re-open needed after WriteExcelTest closes it

    # Display the structure (adapt based on _create_test_data)
    print("\nData structure:")
    # ... (print structure similar to how it was before) ...
    print("1. Projects (rows 1-approx 9):")
    headers, _ = test_excel.read_row(1)
    if headers: print(f"   Headers: {', '.join(headers)}")
    # Print a couple of sample rows
    for r_idx in range(2, 5):
        try:
            r_data, _ = test_excel.read_row(r_idx)
            if r_data and r_data[0]: print(f"   Row {r_idx}: {r_data[0]} - {r_data[1]} (סטטוס: {r_data[2]})")
        except: pass # Ignore errors reading sample rows


    print("\nExcel Cell Update Interactive Test (Hebrew Data)")
    print("Type 'exit' to quit")
    print("Type 'show' to display current data")
    print("Type 'debug' to see what's being sent to the LLM")

    chat_history = []
    debug_mode = False

    while True:
        user_input = input("\nEnter your instruction (e.g., 'שנה סטטוס של פרויקט אלפא להושלם'): ") # Hebrew example

        if user_input.lower() == 'exit':
            break
        # ... (rest of your interactive loop logic: show, debug, generate, extract, execute) ...
        if user_input.lower() == 'debug':
            debug_mode = not debug_mode
            print(f"Debug mode: {'ON' if debug_mode else 'OFF'}")
            continue

        if user_input.lower() == 'show':
            print("\nCurrent Data:")
            # ... (your detailed 'show' logic to read and print rows/headers) ...
            print("\nProjects:")
            project_headers, _ = test_excel.read_row(1)
            print(f"   Headers: {', '.join([str(h) for h in project_headers if h])}")
            # Determine approx last row dynamically or use a reasonable upper limit
            last_proj_row = 10 # Adjust as needed
            for row_idx in range(2, last_proj_row):
                row_data, _ = test_excel.read_row(row_idx)
                if not row_data or not any(row_data): break # Stop if empty row
                formatted_row = []
                for i, cell in enumerate(row_data):
                    if cell and project_headers and i < len(project_headers) and project_headers[i]:
                        formatted_row.append(f"{project_headers[i]}: {cell}")
                    elif cell:
                        formatted_row.append(str(cell))
                print(f"   Row {row_idx}: {', '.join(formatted_row)}")
            # Add logic for Tasks/Employees if they exist in _create_test_data
            continue

        print(f"\nProcessing: '{user_input}'")

        # Create context prompt (logic can be reused or adapted)
        context_prompt = "..." # Generate context using test_excel

        if debug_mode:
             print("\n=== FULL LLM PROMPT ===")
             # ... (print system, context, user instruction) ...
             print("=== END PROMPT ===\n")


        response = generate_response(user_input, chat_history, test_excel)
        print("\nLLM Response:", response)

        json_data = extract_json_from_response(response)
        if json_data:
            print("\nDetected JSON command:")
            print(json.dumps(json_data, indent=2, ensure_ascii=False))
            print("\nExecuting command...")
            try:
                json_str = json.dumps(json_data, ensure_ascii=False)
                reward, feedback = test_excel.process_json_operation(json_str)
                print(f"Result: {'✅ Success' if reward == 1 else '❌ Failed'}")
                print(f"Feedback: {feedback}")
            except Exception as e:
                print(f"Error executing command: {str(e)}")
        else:
            print("\nNo valid JSON command found in the response.")

        chat_history.append({"role": "user", "content": user_input})
        chat_history.append({"role": "assistant", "content": response})


    # Clean up interactive test file
    test_excel.workbook.close()
    if os.path.exists(test_file):
        os.remove(test_file)
    print("\nInteractive test complete. Excel file cleaned up.")


# --- add_new_test_scenario Function (MODIFIED - Adds to CONVERTED format) ---
def add_new_test_scenario(filename="converted_scenarios.json"):
    """Add a new test scenario (instruction + command) to the JSON file."""
    scenarios = load_scenarios(filename) # Load existing converted scenarios

    print("\nAdd New Test Scenario (Command Execution Test)\n")
    instruction = input("Enter the original instruction/prompt: ")
    row_header = input("Row identifier column header (e.g., 'שם הפרויקט'): ")
    row_value = input("Value to find in row identifier column (e.g., 'פרויקט אלפא'): ")
    col_header = input("Column header to update (e.g., 'סטטוס'): ")
    new_value = input("New value to write (e.g., 'הושלם'): ")

    new_scenario = {
        "instruction": instruction,
        "command": {
             "function_name": "excel_update_cell_by_lookup", # Hardcoded for this test type
             "parameters": {
                 "row_header": row_header,
                 "row_value": row_value,
                 "col_header": col_header,
                 "new_value": new_value
             }
        }
    }

    scenarios.append(new_scenario)

    # Save updated scenarios
    try:
        with open(filename, 'w', encoding='utf-8') as f:
            json.dump(scenarios, f, indent=4, ensure_ascii=False)
        print(f"Added new scenario to '{filename}'. Total scenarios: {len(scenarios)}")
    except Exception as e:
        print(f"Error saving updated scenarios to '{filename}': {e}")


# --- create_default_test_files Function (MODIFIED) ---
def create_default_test_files():
    """Creates default converted_scenarios.json and playground.xlsx files."""
    scenario_file = "converted_scenarios.json"
    input_file_for_conversion = "input_data.json" # The file convert_scenarios expects
    playground_file = "playground.xlsx"
    test_file_template = "write_cell_test.xlsx" # Template for Excel structure

    # 1. Create default input_data.json if needed, then run conversion
    if not os.path.exists(input_file_for_conversion):
         print(f"'{input_file_for_conversion}' not found.")
         # Create a dummy one so conversion script doesn't fail immediately
         dummy_data = [
            {
                "parsed_function_call_json": {
                    "instruction": "עדכן סטטוס של פרויקט אלפא להושלם",
                    "function_name": "excel_update_cell_by_lookup",
                    "parameters": { "row_header": "שם פרויקט", "row_value": "פרויקט אלפא", "col_header": "סטטוס", "new_value": "הושלם" }
                }, "excel_headers": ["שם פרויקט", "סטטוס"]
            }
         ]
         try:
            with open(input_file_for_conversion, 'w', encoding='utf-8') as f:
                json.dump(dummy_data, f, ensure_ascii=False, indent=4)
            print(f"Created dummy '{input_file_for_conversion}'. Please edit it with real data.")
         except Exception as e: print(f"Error creating dummy input file: {e}")

    # Run conversion (it handles its own input file check now)
    print("\nRunning conversion script...")
    convert_json_format(input_file_for_conversion, scenario_file)
    print("-" * 20)


    # 2. Create playground.xlsx
    # Ensure the template exists or create it
    if not os.path.exists(test_file_template):
        print(f"Template '{test_file_template}' not found, creating a default one...")
        try:
            # Use WriteExcelTest just to create the file with data
            temp_tester = WriteExcelTest(test_file_template)
            temp_tester.cleanup() # This creates the file, then closes & removes it if no backup existed. We need it to stay.
            # Re-create it cleanly
            temp_tester = WriteExcelTest(test_file_template)
            temp_tester.excel.workbook.close() # Just close, don't cleanup/remove
            print(f"Created template file '{test_file_template}'.")
        except Exception as e:
            print(f"Error creating template file '{test_file_template}': {e}")
            # If template creation fails, cannot create playground
            return

    # Copy template to playground
    if os.path.exists(test_file_template):
        try:
            shutil.copyfile(test_file_template, playground_file)
            print(f"Created '{playground_file}' from template '{test_file_template}'")
        except Exception as e:
            print(f"Error copying template to '{playground_file}': {e}")
    else:
         print(f"Error: Could not create '{playground_file}' because template '{test_file_template}' could not be created/found.")


# --- Main Execution Block (MODIFIED) ---
if __name__ == "__main__":
    print("Excel Command Execution Test Suite")
    print("1. Run automated tests (Executes commands from converted_scenarios.json on write_cell_test.xlsx - NO LLM)") # Updated description
    print("2. Run interactive test (Uses LLM - Hebrew Data)")
    print("3. Add new test scenario (to converted_scenarios.json)") # Updated description
    print("4. Create/Update default test files (input_data.json -> converted_scenarios.json, playground.xlsx)") # Updated description

    choice = input("Enter your choice (1-4): ")

    if choice == "1":
        # Run automated tests using the converted scenarios and specific Excel file
        run_automated_tests(scenario_file="converted_scenarios.json", excel_file="write_cell_test.xlsx")
    elif choice == "2":
        run_interactive_test()
    elif choice == "3":
        add_new_test_scenario(filename="converted_scenarios.json") # Add to the converted file
    elif choice == "4":
        create_default_test_files()
    else:
        print("Invalid choice. Exiting.")