"""
Generate JSON dataset from Excel using Gemini API (Hebrew instructions and JSON) with robust JSON extraction,
rate limiting, loading feedback, dotenv, enhanced prompt with row data and headers, error handling, and saving raw JSON responses with headers.

This script enhances JSON extraction from Gemini API responses to be more robust,
handling potential surrounding text and variations in output format. It includes
rate limiting, loading feedback, dotenv for API key, an enhanced prompt that includes
the entire Excel row data and headers for context, comprehensive error handling,
and saves the raw Gemini API responses as JSON to a separate file, now including
headers for better debugging and analysis of column choices.

Usage:
    python generate_json_from_excel_gemini_headers_in_raw_json.py <excel_file_path> [--start_row <row_number>]

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
from google.generativeai import types
from dotenv import load_dotenv  # Import load_dotenv
import re # Import regular expression library

# Load environment variables from .env file
load_dotenv()

fake = Faker() # Initialize Faker, potentially for future use

def get_excel_cell_address_from_pandas(row_index, col_index):
    """Convert pandas 0-based indices to Excel cell address (A1, B2, etc.)"""
    col_letters = ""
    col_idx_excel = col_index + 1
    while col_idx_excel > 0:
        col_idx_excel, remainder = divmod(col_idx_excel - 1, 26)
        col_letters = chr(65 + remainder) + col_letters
    return f"{col_letters}{row_index + 2}"


from google import genai # Updated import


def generate_instruction_and_json_with_gemini(project_name, special_column_header, current_value, excel_row_number, raw_response_file, row_data, headers_list):
    """
    Generates a Hebrew instruction and JSON function call using Gemini API with robust JSON extraction, saves raw response as JSON, and includes row data and headers in prompt.
    Gemini selects the new value and provides JSON in Hebrew.
    """
    gemini_api_key = os.environ.get("GEMINI_API_KEY")
    if not gemini_api_key:
        raise EnvironmentError("GEMINI_API_KEY environment variable not set. Ensure GEMINI_API_KEY is in your .env file or environment variables.")

    # Initialize Gemini Client with API key directly in constructor (newest way)
    client = genai.Client(api_key=gemini_api_key)  # API Key passed here

    model_name = 'gemini-2.0-flash' # Or 'gemini-pro' if you prefer, adjust as needed

    # Enhanced prompt to include all row data and headers
    row_data_text = "\n".join([f"- {headers_list[i]}: {row_data[header]}" for i, header in enumerate(headers_list)])
    headers_text = ", ".join(headers_list) # Comma-separated string of headers for prompt

    prompt_text = f"""
    You are an expert in generating instructions for updating Excel sheets in Hebrew.
    For the following project and data from an Excel sheet, please generate:
    1. A concise natural language instruction in Hebrew to update the column '{special_column_header}'.
    2. A JSON object representing a function call to 'excel_update_cell_by_lookup' to perform this update.

    **Available Excel Column Headers:** [{headers_text}]
    **Project:** '{project_name}'
    **Column to update:** '{special_column_header}'
    **Current value of '{special_column_header}':** '{current_value}'
    **All data from the Excel row:**
    {row_data_text}

    **Instructions:**
    - Decide on a relevant new value for '{special_column_header}' for project '{project_name}', considering the context of the entire row and available column headers.
    - Generate a JSON object in the following format:
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
        a) The instruction is a short, natural-sounding request in Hebrew.
        b) All JSON values ('row_value', 'col_header', 'new_value') are in Hebrew.
        c) Include a key called "instruction" in the JSON, with the Hebrew instruction as its value.
        d) Output the JSON object. It's okay if there's descriptive text before or after the JSON, I will extract it.

    Example Output:
    הנה ה JSON:
    ```json
    {{
      "instruction": "עדכן סטטוס של פרויקט פרויקט אלף ל'בביצוע'",
      "function_name": "excel_update_cell_by_lookup",
      "parameters": {{
        "row_header": "שם הפרויקט",
        "row_value": "פרויקט אלף",
        "col_header": "סטטוס",
        "new_value": "בביצוע"
      }}
    }}
    ```
    """

    print(f"  ⏳ Querying Gemini API for Hebrew instruction and JSON for '{project_name}' (Row {excel_row_number})...", end=" ", flush=True) # Loading feedback

    try:
        response = client.models.generate_content(
            model=model_name,  # Specify the model name
            contents=prompt_text  # Pass the prompt directly as 'contents'
        )
        gemini_output = response.text.strip()
        print("✅ Done.") # End loading feedback - success

        instruction_hebrew = None
        function_call_json = None

        # Robust JSON extraction using regex (same as before)
        json_match = re.search(r'```json\s*(\{[\s\S]*?\})\s*```', gemini_output, re.IGNORECASE) # Look for JSON within ```json ... ``` blocks
        extracted_json_string = None # Variable to hold extracted JSON string

        if json_match:
            extracted_json_string = json_match.group(1) # Extracted JSON string from Gemini's response
            try:
                function_call_json = json.loads(extracted_json_string)

                # Attempt to extract instruction - assuming it's before the JSON block.
                instruction_candidate = gemini_output[:json_match.start()].strip()
                if instruction_candidate:
                    instruction_hebrew = instruction_candidate
                else: # If no clear instruction before JSON, use the whole text *without* the json part as instruction (less ideal, but handles cases)
                    instruction_hebrew = gemini_output[:json_match.start()].strip()


            except json.JSONDecodeError as e:
                print(f"  ⚠️ Gemini generated invalid JSON: {e}")
                instruction_hebrew = gemini_output # Still return the whole Gemini output as instruction for context in error case
                function_call_json = None # Ensure function_call_json is None in case of JSON decode error

        else:
            print(f"  ⚠️ No JSON block found in Gemini output.")
            instruction_hebrew = gemini_output # Return the whole output as instruction
            function_call_json = None # No JSON found


        if instruction_hebrew is None: # Fallback if instruction extraction failed, use the whole output.
            instruction_hebrew = gemini_output


        # Prepare raw response JSON object - simplified and with row data
        raw_response_json = {
            "response": gemini_output,
            "excel_row_number": excel_row_number,
            "parsed_function_call_json": function_call_json,
            "excel_headers": headers_list # Include headers in raw JSON
        }


        # Save raw response as JSON to file
        raw_response_file.write(json.dumps(raw_response_json, ensure_ascii=False, indent=2) + ",\n") # Write JSON object and a comma, for array format
        raw_response_file.flush() # Ensure it's written immediately


        return instruction_hebrew, function_call_json

    except Exception as e:
        print(f"❌ Error: {e}") # End loading feedback - error
        print(f"  ⚠️ Gemini API call failed for '{project_name}' (Row {excel_row_number}). Skipping instruction and JSON generation.")
        return None, None



def generate_data_point_from_excel_row(excel_file_path, row_data, row_index, headers, raw_response_file, headers_list):
    """
    Generates a data point (instruction, function call, context) from an Excel row using Gemini and saves raw responses as JSON with row data.
    Handles JSON validation and error cases.
    """
    project_name_header = "שם הפרויקט"

    if project_name_header not in headers:
        print(f"Error: Column '{project_name_header}' not found in Excel headers.")
        return None, "MissingProjectNameColumn"

    project_name = row_data[project_name_header]

    available_columns = [header for header in headers if header != project_name_header]
    if not available_columns:
        print(f"Warning: No columns available to update other than '{project_name_header}'. Skipping row.")
        return None, "NoColumnsToUpdate"

    special_column_header = random.choice(available_columns)
    special_col_index = headers.get_loc(special_column_header)
    current_value = row_data[special_column_header]
    excel_row_number = row_index + 2 # Excel row number for context

    instruction, function_call_json = generate_instruction_and_json_with_gemini(
        project_name, special_column_header, current_value, excel_row_number, raw_response_file, row_data, headers_list
    )

    if instruction and function_call_json:
        # Basic JSON validation (you can add more checks if needed)
        if not isinstance(function_call_json, dict) or "function_name" not in function_call_json or "parameters" not in function_call_json:
            print(f"  ⚠️ Gemini JSON structure is invalid.")
            return None, "InvalidJSONStructure"

        data_point = {
            "instruction": instruction, # Hebrew instruction
            "context": {
                "PROJECTS": {
                    "headers": list(headers),
                    "rows": [row_data.to_dict()]
                }
            },
            "function_call": function_call_json, # The extracted function call JSON
            "excel_file_path": excel_file_path,
            "processing_status": "success" # Add status to data point
        }
        return data_point, "success"
    else:
        return None, "GeminiError" # Or "InvalidJSON" if generate_instruction_and_json_with_gemini returns (instruction, None)


if __name__ == "__main__":
    if len(sys.argv) < 2:
        print("Usage: python generate_json_from_excel_headers_in_raw_json.py <excel_file_path> [--start_row <row_number>]")
        sys.exit(1)

    excel_file_path = sys.argv[1]
    start_row = 1 # Default start row is the second row in Excel (after headers)

    try:
        start_row_index_arg = sys.argv.index('--start_row') + 1
        start_row = int(sys.argv[start_row_index_arg])
        if start_row < 1:
            start_row = 1 # Ensure start_row is at least 1
    except ValueError:
        print("Error: --start_row must be followed by a valid row number.")
        sys.exit(1)
    except IndexError:
        pass # --start_row argument not provided, use default start_row
    except Exception as e:
        print(f"An error occurred while parsing command line arguments: {e}")
        sys.exit(1)


    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file not found at path: {excel_file_path}")
        sys.exit(1)

    num_examples_to_generate = 700
    output_file = os.path.join(os.path.dirname(__file__), "synthetic_excel_data_gemini_hebrew_headers_raw_json_responses.json")
    raw_responses_file_path = os.path.join(os.path.dirname(__file__), "gemini_raw_responses.json") # File to save raw Gemini outputs as JSON

    generated_examples = [] # Store successful examples
    error_examples = [] # Store examples with errors

    try:
        df = pd.read_excel(excel_file_path)
        headers_pd_index = df.columns
        headers = pd.Index(headers_pd_index)
        headers_list = list(headers) # Get headers list here to pass to functions


        if "שם הפרויקט" not in headers:
            print(f"Error: Required column 'שם הפרויקט' not found in the Excel file.")
            sys.exit(1)

        start_pandas_row_index = start_row - 2 # Convert Excel row number to pandas 0-based index
        if start_pandas_row_index < 0:
            start_pandas_row_index = 0

        if start_pandas_row_index >= len(df):
            print(f"Warning: start_row is beyond the last row of the Excel file. No rows to process.")
            sys.exit(0)

        df_processed = df.iloc[start_pandas_row_index:] # Process from start_row onwards

        example_count = 0
        print(f"Starting synthetic data generation from Excel row {start_row}...") # Start process feedback

        # Prepare raw responses JSON file - start with JSON array opening
        with open(raw_responses_file_path, 'w', encoding='utf-8') as raw_response_file: # Open raw responses file in write mode to start fresh as JSON array
            raw_response_file.write("[\n") # Start JSON array

            is_first_raw_response = True # Flag - not used anymore, but kept for potential future use

            for index, row in df_processed.iterrows():
                if example_count >= num_examples_to_generate:
                    break

                excel_row_number = index + 2 # Excel row number is index + 2
                print(f"Processing row {excel_row_number} (Project: '{row['שם הפרויקט'] if 'שם הפרויקט' in row else 'N/A'}')...") # Row processing feedback

                data_point, status = generate_data_point_from_excel_row(
                    excel_file_path, row, index, headers, raw_response_file, headers_list # Pass headers_list
                )

                if status == "success" and data_point:
                    generated_examples.append(data_point)
                    print(f"  ✅ Example {example_count + 1}/{num_examples_to_generate} generated: {data_point['instruction']}")
                    example_count += 1
                else:
                    error_data_point = {
                        "excel_file_path": excel_file_path,
                        "excel_row_number": excel_row_number,
                        "project_name": row['שם הפרויקט'] if 'שם הפרויקט' in row else 'N/A',
                        "error_type": status, # Store the error status
                        "processing_status": "error" # Mark as error in general output
                    }
                    error_examples.append(error_data_point)
                    print(f"  ⚠️ Warning: Failed to generate example for row {excel_row_number} due to: {status}")


                time.sleep(5) # Rate limiting: wait for 5 seconds before next iteration (and Gemini API call indirectly)


            # Close raw responses JSON array
            with open(raw_responses_file_path, 'a', encoding='utf-8') as raw_response_file_close: # Reopen in append mode to close array
                raw_response_file_close.write("\n]") # Close JSON array


    except FileNotFoundError:
        print(f"Error: Excel file not found at: {excel_file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred during Excel processing: {e}")
        sys.exit(1)


    # Combine successful and error examples into final JSON
    final_dataset = {
        "generated_examples": generated_examples,
        "error_examples": error_examples
    }


    with open(output_file, 'w', encoding='utf-8') as f_out:
        json.dump(final_dataset, f_out, ensure_ascii=False, indent=2)


    print(f"\nGeneration complete: {example_count} successful examples generated, {len(error_examples)} errors.") # End process feedback
    print(f"Dataset saved to {output_file}")
    print(f"Raw Gemini responses saved to: {raw_responses_file_path} (as JSON array)") # Inform user about raw responses file