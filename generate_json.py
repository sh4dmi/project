"""
Generate JSON dataset from Excel using Gemini API (Hebrew instructions and JSON) with robust JSON extraction,
rate limiting, loading feedback, dotenv, enhanced prompt with row data and headers, error handling, and saving
parsed JSON responses with headers (removed raw text).

This script enhances JSON extraction from Gemini API responses to be more robust,
handling potential surrounding text and variations in output format. It includes
rate limiting, loading feedback, dotenv for API key, an enhanced prompt that includes
the entire Excel row data and headers for context, comprehensive error handling,
and saves the PARSED Gemini API JSON responses along with context (headers, row number)
to a separate file for debugging and analysis. The raw text response is no longer saved
in this file to keep it focused.

The prompt encourages Gemini to generate more varied and human-like Hebrew instructions.

Usage:
    python generate_json_from_excel_gemini_humanlike_prompt.py <excel_file_path> [--start_row <row_number>]

Options:
    --start_row <row_number>  Row number to start processing from (default is 1, the second row in Excel after headers).

Requirements:
    - Python 3.6+
    - pandas
    - faker
    - openpyxl (for .xlsx) or xlrd (for .xls)
    - google-generativeai
    - python-dotenv (pip install python-dotenv)
"""

import json
import random
import os
import sys
import time  # Import time for rate limiting
from datetime import datetime, timedelta
from faker import Faker # Still import Faker, might be used for fallback or other purposes later
import pandas as pd
import google.generativeai as genai
from google.generativeai import types # Although types might not be directly used here, keep for potential future use or if Client expects it indirectly
from dotenv import load_dotenv  # Import load_dotenv
import re # Import regular expression library

# Load environment variables from .env file
load_dotenv()

fake = Faker() # Initialize Faker, potentially for future use

# This function seems unused in the current flow, but kept for potential future utility
def get_excel_cell_address_from_pandas(row_index, col_index):
    """Convert pandas 0-based indices to Excel cell address (A1, B2, etc.)"""
    col_letters = ""
    col_idx_excel = col_index + 1
    while col_idx_excel > 0:
        col_idx_excel, remainder = divmod(col_idx_excel - 1, 26)
        col_letters = chr(65 + remainder) + col_letters
    return f"{col_letters}{row_index + 2}"


# No need to re-import genai here if already imported above
# from google import genai

# ... (Keep all imports and other functions as they are) ...

def generate_instruction_and_json_with_gemini(project_name, special_column_header, current_value, excel_row_number, raw_response_file, row_data, headers_list):
    """
    Generates a Hebrew instruction and JSON function call using Gemini API (compatible method)
    with robust JSON extraction, saves parsed response details, and includes row data/headers in prompt.
    The prompt encourages more human-like and varied Hebrew instructions, avoiding repetitive quoting.
    """
    gemini_api_key = os.environ.get("GEMINI_API_KEY")
    if not gemini_api_key:
        raise EnvironmentError("GEMINI_API_KEY environment variable not set. Ensure GEMINI_API_KEY is in your .env file or environment variables.")

    # **API Initialization (compatible method)**
    try:
        genai.configure(api_key=gemini_api_key)
        # Consider trying different models if 'flash' is too basic or repetitive
        # model_name = 'gemini-1.5-pro-latest' # Potentially better at complex instructions
        model_name = 'gemini-1.5-flash' # Stick with flash for now if preferred
        model = genai.GenerativeModel(model_name)
    except Exception as e:
        print(f"\n❌ Failed to configure Gemini or find model '{model_name}': {e}")
        raise EnvironmentError(f"Could not configure Gemini. Check API key and model name. Error: {e}")

    # Enhanced prompt context presentation
    row_data_text = "\n".join([f"- {headers_list[i]}: {row_data.get(header, '')}" for i, header in enumerate(headers_list)])
    headers_text = ", ".join(headers_list)

    # ----- ***** REFINED PROMPT TEXT ***** -----
    prompt_text = f"""
    You are an expert in generating realistic, human-like instructions for updating Excel sheets in Hebrew.
    Your goal is to create varied and natural-sounding requests.

    Here is the context from an Excel row:
    **Available Excel Column Headers:** [{headers_text}]
    **Project Name:** {project_name}
    **Column to Update:** {special_column_header}
    **Current Value in '{special_column_header}' column:** {current_value}
    **Full Row Data:**
    {row_data_text}

    Please perform the following tasks:
    1. Decide on a relevant *new* value for the '{special_column_header}' column related to the project "{project_name}", considering all the row data.
    2. Generate a natural language instruction *in Hebrew* telling someone to make this update.
    3. Generate a JSON object representing a function call to 'excel_update_cell_by_lookup' to execute the update.

    **Instructions for the Hebrew Instruction (Task 2):**
    - **Variety is key!** Phrase the instruction as a real person might ask a colleague. Use different sentence structures: direct commands, questions, polite requests, statements about what needs doing.
    - **Refer to the project, column, and values naturally.** Don't just repeat the names rigidly. Examples: "לגבי {project_name}, צריך לשנות את ה{special_column_header} ל...", "אפשר לעדכן בבקשה את הסטטוס של הפרויקט הזה?", "הערך הנוכחי ({current_value}) בעמודה {special_column_header} לא נכון, שנה ל...", "מה הסטטוס החדש של {project_name}?".
    - **CRITICAL: Avoid consistently putting project names, column names, or values inside single quotes (' ') in the instruction text.** Use quotes only if they are truly natural in Hebrew for emphasis in that specific context, which should be rare. Refer to things directly by name or description.
    - Minor, realistic-sounding grammatical quirks or informalities are okay if they sound natural, but the instruction must be clear.

    **Instructions for the JSON Object (Task 3):**
    - Generate a JSON object in *exactly* this format:
    ```json
    {{
      "function_name": "excel_update_cell_by_lookup",
      "parameters": {{
        "row_header": "שם הפרויקט",
        "row_value": "<project_name_in_hebrew>",
        "col_header": "<special_column_header_in_hebrew>",
        "new_value": "<new_value_in_hebrew>"
      }}
    }}
    ```
    - Ensure:
        a) The `instruction` key holds the Hebrew instruction generated in Task 2.
        b) All JSON string values (`row_value`, `col_header`, `new_value`) are in Hebrew. Note: These JSON *values* must be strings, even if the instruction doesn't use quotes.
        c) Output *only* the complete JSON object, enclosed in ```json ... ``` markers. Do not include any other text before or after the JSON block.

    Example of desired JSON output structure (the instruction text itself should vary greatly based on the rules above):
    ```json
    {{
      "instruction": "בפרויקט {project_name} צריך לעדכן את {special_column_header} שיהיה [ערך חדש]", // Natural phrasing, replace placeholders
      "function_name": "excel_update_cell_by_lookup",
      "parameters": {{
        "row_header": "שם הפרויקט",
        "row_value": "שם פרויקט לדוגמה", // Value is a string
        "col_header": "כותרת עמודה לדוגמה", // Value is a string
        "new_value": "ערך חדש לדוגמה" // Value is a string
      }}
    }}
    ```
    """
    # ----- ***** END OF REFINED PROMPT TEXT ***** -----


    print(f"  ⏳ Querying Gemini API for Hebrew instruction and JSON for '{project_name}' (Row {excel_row_number})...", end=" ", flush=True)

    try:
        # API Call (compatible method)
        response = model.generate_content(
            contents=[prompt_text],
            # safety_settings={'HARASSMENT': 'BLOCK_NONE', 'HATE_SPEECH': 'BLOCK_NONE', 'SEXUAL': 'BLOCK_NONE', 'DANGEROUS': 'BLOCK_NONE'} # Uncomment if needed
             generation_config=genai.types.GenerationConfig(
                 # temperature=0.9 # Increase temperature for more creativity/variation if needed (e.g., 0.7 to 1.0)
             )
        )

        # --- Response processing and JSON extraction (Keep this logic the same) ---
        gemini_output = ""
        if not response.candidates:
             print(f"❌ No candidates returned. Possible safety block or other issue.")
             # Add improved feedback check if available
             try:
                 feedback = response.prompt_feedback
                 print(f"   Response prompt feedback: {feedback}")
                 reason = f"{feedback}"
             except Exception:
                  print("   Could not retrieve detailed prompt feedback.")
                  reason = "Unknown reason (no candidates)"

             error_raw_info = {
                 "excel_row_number": excel_row_number,
                 "parsed_function_call_json": {"error": "No Content Generated", "reason": reason},
                 "excel_headers": headers_list
             }
             raw_response_file.write(json.dumps(error_raw_info, ensure_ascii=False, indent=2) + ",\n")
             raw_response_file.flush()
             return None, None

        try:
            # Prioritize getting text from parts if available
            if response.candidates[0].content and response.candidates[0].content.parts:
                 gemini_output = "".join(part.text for part in response.candidates[0].content.parts).strip()
            # Fallback to response.text if parts aren't structured as expected but text attribute exists
            elif hasattr(response, 'text'):
                 gemini_output = response.text.strip()
            else:
                 gemini_output = str(response.candidates[0].content)
                 print("⚠️ Gemini response content structure unexpected, using basic string conversion.")
        except (IndexError, AttributeError, Exception) as e:
            print(f"❌ Error accessing response content: {e}")
            gemini_output = "" # Ensure it's empty if access fails


        if not isinstance(gemini_output, str) or not gemini_output:
             print("⚠️ Gemini response format unexpected or empty.")
             # Log empty/failed response detail
             error_raw_info = {
                 "excel_row_number": excel_row_number,
                 "parsed_function_call_json": {"error": "Empty or Invalid Response Format"},
                 "excel_headers": headers_list
             }
             raw_response_file.write(json.dumps(error_raw_info, ensure_ascii=False, indent=2) + ",\n")
             raw_response_file.flush()
             return None, None # Treat as failure

        print("✅ Done.")

        # --- JSON Extraction Logic ---
        instruction_hebrew = None
        function_call_json = None
        extracted_json_string = None

        # Regex to find JSON block ```json ... ``` or just { ... }
        json_match = re.search(r'```json\s*(\{[\s\S]*?\})\s*```|\s*(\{[\s\S]*?\})\s*', gemini_output, re.IGNORECASE | re.MULTILINE)

        if json_match:
            extracted_json_string = json_match.group(1) or json_match.group(2)
            try:
                function_call_json = json.loads(extracted_json_string)
                # Validate expected keys exist
                if not isinstance(function_call_json, dict) or \
                   "instruction" not in function_call_json or \
                   "function_name" not in function_call_json or \
                   "parameters" not in function_call_json or \
                   not isinstance(function_call_json.get("parameters"), dict):
                    print(f"  ⚠️ Parsed JSON is missing required keys (instruction, function_name, parameters dict).")
                    instruction_hebrew = function_call_json.get("instruction", "Instruction Missing, JSON Structure Invalid") # Try to get instruction anyway
                    function_call_json = None # Mark JSON as invalid
                else:
                    instruction_hebrew = function_call_json.get("instruction")
                    if not instruction_hebrew: # Handle empty instruction string
                         print("  ⚠️ 'instruction' key exists in JSON but the value is empty.")
                         instruction_hebrew = "Instruction Empty in JSON"
                         function_call_json = None # Treat as invalid if instruction is mandatory and empty


            except json.JSONDecodeError as e:
                print(f"  ⚠️ Gemini generated invalid JSON string: {e}")
                instruction_hebrew = gemini_output.strip() if gemini_output.strip() else f"JSON Parsing Error: {e}"
                function_call_json = None

        else:
            print(f"  ⚠️ No JSON block found in Gemini output.")
            instruction_hebrew = gemini_output.strip() if gemini_output.strip() else "No JSON and No Text Output"
            function_call_json = None

        # --- Save Parsed Details ---
        raw_response_details_json = {
            "excel_row_number": excel_row_number,
            "parsed_function_call_json": function_call_json, # Parsed JSON or None
            "excel_headers": headers_list # Context
        }
        # Use try-except for file writing for robustness
        try:
            raw_response_file.write(json.dumps(raw_response_details_json, ensure_ascii=False, indent=2) + ",\n")
            raw_response_file.flush()
        except Exception as file_err:
             print(f"  ⚠️ Failed to write details to raw log file: {file_err}")


        # --- Return Value Logic ---
        # Ensure consistency: if function_call_json is None (invalid/missing), instruction should also be None for the main dataset processing logic
        if not function_call_json:
             instruction_hebrew = None # Signal failure clearly

        return instruction_hebrew, function_call_json

    # --- Exception Handling ---
    except Exception as e:
        print(f"❌ Unhandled Error during Gemini API call or processing: {e}")
        import traceback
        print(traceback.format_exc())
        print(f"  ⚠️ Gemini API call failed for '{project_name}' (Row {excel_row_number}). Skipping.")
        # Log minimal info to raw file
        error_raw_info = {
             "excel_row_number": excel_row_number,
             "parsed_function_call_json": {"error": "Unhandled Exception", "message": str(e)},
             "excel_headers": headers_list
         }
        try:
            raw_response_file.write(json.dumps(error_raw_info, ensure_ascii=False, indent=2) + ",\n")
            raw_response_file.flush()
        except Exception as file_err:
             print(f"  ⚠️ Also failed to write error details to raw log file: {file_err}")
        return None, None


# --- IMPORTANT ---
# Ensure the rest of the script (generate_data_point_from_excel_row,
# the main __name__ == "__main__" block) remains unchanged from the previous version.
# You only need to replace the generate_instruction_and_json_with_gemini function.


def generate_data_point_from_excel_row(excel_file_path, row_data, row_index, headers, raw_response_file, headers_list):
    """
    Generates a data point (instruction, function call, context) from an Excel row using Gemini and saves parsed response details.
    Handles JSON validation and error cases.
    """
    project_name_header = "שם הפרויקט" # Assuming this is the key column in Hebrew

    if project_name_header not in headers:
        print(f"Error: Column '{project_name_header}' not found in Excel headers.")
        return None, "MissingProjectNameColumn"

    # Handle potential NaN or None values in project name gracefully
    project_name_val = row_data.get(project_name_header)
    project_name = str(project_name_val) if pd.notna(project_name_val) else "Unknown Project"


    available_columns = [header for header in headers if header != project_name_header]
    if not available_columns:
        print(f"Warning: No columns available to update other than '{project_name_header}'. Skipping row {row_index + 2}.")
        return None, "NoColumnsToUpdate"

    special_column_header = random.choice(available_columns)
    # special_col_index = headers.get_loc(special_column_header) # Index not directly needed here
    current_value_val = row_data.get(special_column_header)
    current_value = str(current_value_val) if pd.notna(current_value_val) else "" # Handle NaN/None current values

    excel_row_number = row_index + 2 # Excel row number for context

    instruction, function_call_json = generate_instruction_and_json_with_gemini(
        project_name, special_column_header, current_value, excel_row_number, raw_response_file, row_data.to_dict(), headers_list # Pass row_data as dict
    )

    if instruction and function_call_json:
        # Basic JSON validation
        if not isinstance(function_call_json, dict) or \
           "function_name" not in function_call_json or \
           "parameters" not in function_call_json or \
           not isinstance(function_call_json.get("parameters"), dict) or \
           "instruction" not in function_call_json: # Check for instruction key as well
            print(f"  ⚠️ Gemini JSON structure is invalid or missing 'instruction' key for row {excel_row_number}.")
            # The error details were already saved to the raw file inside the Gemini function
            return None, "InvalidJSONStructure"

        data_point = {
            "instruction": instruction, # Hebrew instruction (extracted, possibly varied)
            "context": {
                "PROJECTS": { # Assuming this structure is desired for the final dataset
                    "headers": headers_list, # Use the consistent list
                    "rows": [row_data.to_dict()] # Current row data as context
                }
            },
            "function_call": function_call_json, # The extracted and validated function call JSON
            "excel_file_path": excel_file_path,
            "processing_status": "success" # Add status to data point
        }
        return data_point, "success"
    else:
        # Error details (like API failure, JSON parse error) were already saved to raw file
        # Determine more specific status if possible
        status = "GeminiError"
        if instruction and not function_call_json: # e.g., JSON parse error
            status = "InvalidJSONOutput"
        elif not instruction and not function_call_json: # e.g., API error, safety stop
             status = "GeminiAPIFailureOrSafety"

        print(f"  ⚠️ Failed to generate valid data point for row {excel_row_number}. Status: {status}")
        return None, status


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print(f"Usage: python {os.path.basename(__file__)} <excel_file_path> [--start_row <row_number>]")
        sys.exit(1)

    excel_file_path = sys.argv[1]
    start_row = 2 # Default start row is the second row in Excel (index 1, after headers)

    # Parse --start_row argument
    try:
        if '--start_row' in sys.argv:
            start_row_index_arg = sys.argv.index('--start_row') + 1
            if start_row_index_arg < len(sys.argv):
                start_row = int(sys.argv[start_row_index_arg])
                if start_row < 2:
                    print("Warning: --start_row cannot be less than 2 (first data row). Setting to 2.")
                    start_row = 2
            else:
                 raise ValueError("--start_row option requires a number.")
    except ValueError as e:
        print(f"Error: Invalid --start_row argument. {e}")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred while parsing command line arguments: {e}")
        sys.exit(1)


    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file not found at path: {excel_file_path}")
        sys.exit(1)

    num_examples_to_generate = 700 # Adjust as needed
    # Define output file names relative to the script location
    script_dir = os.path.dirname(__file__) or '.' # Handle running from current dir
    output_file = os.path.join(script_dir, "synthetic_excel_data_gemini_hebrew_humanlike.json")
    raw_responses_details_file_path = os.path.join(script_dir, "gemini_parsed_responses_details.json") # Renamed to reflect content

    generated_examples = [] # Store successful examples
    error_examples = [] # Store examples with errors

    try:
        print(f"Reading Excel file: {excel_file_path}...")
        # Specify engine based on file extension if needed, especially for older .xls
        engine = 'openpyxl' if excel_file_path.endswith('.xlsx') else None
        df = pd.read_excel(excel_file_path, engine=engine)
        headers_list = list(df.columns) # Get headers list here

        project_name_header = "שם הפרויקט" # Make sure this matches your Excel
        if project_name_header not in headers_list:
            print(f"Error: Required column '{project_name_header}' not found in the Excel file headers: {headers_list}")
            sys.exit(1)

        # Convert Excel row number (1-based, header is row 1) to pandas 0-based index
        # start_row = 2 means first data row, which is pandas index 0
        start_pandas_row_index = start_row - 2
        if start_pandas_row_index < 0:
            start_pandas_row_index = 0 # Should not happen if start_row >= 2

        if start_pandas_row_index >= len(df):
            print(f"Warning: start_row ({start_row}) is beyond the last data row ({len(df)+1}) of the Excel file. No rows to process.")
            sys.exit(0)

        # Slice the DataFrame to start from the correct row
        df_processed = df.iloc[start_pandas_row_index:].copy() # Use .copy() to avoid SettingWithCopyWarning if modifying later
        print(f"Processing {len(df_processed)} rows starting from Excel row {start_row} (pandas index {start_pandas_row_index}).")


        example_count = 0
        total_rows_to_process = min(num_examples_to_generate, len(df_processed))
        print(f"Attempting to generate up to {num_examples_to_generate} examples...")

        # Prepare raw responses JSON file - start with JSON array opening
        # Use 'w' to overwrite or start fresh each run
        with open(raw_responses_details_file_path, 'w', encoding='utf-8') as raw_response_file:
            raw_response_file.write("[\n") # Start JSON array

            # Iterate using df_processed.iterrows() which gives (index, Series) pairs
            # The index here will be the original DataFrame index
            for original_index, row in df_processed.iterrows():
                if example_count >= num_examples_to_generate:
                    print(f"Reached target of {num_examples_to_generate} examples. Stopping.")
                    break

                excel_row_number = original_index + 2 # Calculate Excel row number (original index + 2)
                project_name_display = row.get(project_name_header, 'N/A')
                print(f"\nProcessing Excel Row {excel_row_number} (Project: '{project_name_display}')...") # Row processing feedback

                # Pass the original index for accurate row tracking if needed elsewhere,
                # and the row Series, headers list, and file object.
                data_point, status = generate_data_point_from_excel_row(
                    excel_file_path, row, original_index, df.columns, raw_response_file, headers_list # Pass df.columns (Index object) and headers_list (list)
                )

                if status == "success" and data_point:
                    generated_examples.append(data_point)
                    print(f"  ✅ Example {example_count + 1}/{num_examples_to_generate} generated successfully for row {excel_row_number}.")
                    # print(f"     Instruction: {data_point.get('instruction', 'N/A')[:80]}...") # Print snippet
                    example_count += 1
                else:
                    # Error details already printed and logged to raw file within the functions
                    error_data_point = {
                        "excel_file_path": excel_file_path,
                        "excel_row_number": excel_row_number,
                        "project_name": str(row.get(project_name_header, 'N/A')),
                        "error_type": status, # Store the error status string
                        "processing_status": "error" # Mark as error in general output
                    }
                    error_examples.append(error_data_point)
                    # print(f"  ⚠️ Warning handled for row {excel_row_number}. Status: {status}") # Already printed inside functions


                # Rate limiting AFTER processing each row
                print(f"  ⏱️ Waiting 5 seconds before next row...")
                time.sleep(5)


            # Clean up trailing comma and close raw responses JSON array
            # Go back 2 characters (comma and newline) and write the closing bracket
            raw_response_file.seek(raw_response_file.tell() - 2, os.SEEK_SET)
            raw_response_file.truncate()
            raw_response_file.write("\n]") # Close JSON array properly

    except FileNotFoundError:
        print(f"Error: Excel file not found at: {excel_file_path}")
        sys.exit(1)
    except pd.errors.EmptyDataError:
         print(f"Error: Excel file is empty: {excel_file_path}")
         sys.exit(1)
    except ImportError as e:
         print(f"Error: Missing library required for reading Excel file. Maybe 'pip install openpyxl'? Error: {e}")
         sys.exit(1)
    except Exception as e:
        print(f"\nAn unexpected error occurred during processing: {e}")
        import traceback
        traceback.print_exc() # Print detailed traceback for debugging
         # Attempt to close the raw JSON file gracefully even on error
        try:
             with open(raw_responses_details_file_path, 'a', encoding='utf-8') as raw_file_close:
                # Check if file is non-empty before trying to fix array structure
                if raw_file_close.tell() > 5: # Arbitrary small number > size of '['\n'
                     raw_file_close.seek(raw_file_close.tell() - 2, os.SEEK_SET)
                     raw_file_close.truncate()
                     raw_file_close.write("\n]")
                else: # File likely only has '[' or is empty, just close array
                     raw_file_close.write("\n]")

        except Exception as file_e:
             print(f"Additionally, could not properly close the raw JSON file: {file_e}")
        sys.exit(1)


    # Combine successful and error examples into final JSON
    final_dataset = {
        "generation_summary": {
            "timestamp": datetime.now().isoformat(),
            "excel_file": os.path.basename(excel_file_path),
            "target_examples": num_examples_to_generate,
            "rows_processed_from": start_row,
            "successful_examples": len(generated_examples),
            "error_examples": len(error_examples)
        },
        "generated_examples": generated_examples,
        "error_examples": error_examples
    }


    try:
        with open(output_file, 'w', encoding='utf-8') as f_out:
            json.dump(final_dataset, f_out, ensure_ascii=False, indent=2)
        print(f"\n--- Generation Complete ---") # End process feedback
        print(f"Successfully generated examples: {len(generated_examples)}")
        print(f"Examples with errors: {len(error_examples)}")
        print(f"Main dataset saved to: {output_file}")
        print(f"Parsed response details saved to: {raw_responses_details_file_path}")
    except Exception as e:
         print(f"\nError saving final JSON output file: {e}")
         sys.exit(1)