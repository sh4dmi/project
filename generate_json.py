"""
Generate JSON dataset from Excel using Gemini API (Hebrew instructions) with rate limiting, loading feedback, and dotenv for API key.

This script reads an Excel file, processes rows, and uses Gemini API to generate
Hebrew instructions for updating cells. It incorporates rate limiting to respect
API usage, provides loading feedback, and uses dotenv to load the Gemini API key
from a .env file.

Usage:
    python generate_json_from_excel_gemini_rate_limit_dotenv.py <excel_file_path>

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
from faker import Faker
import pandas as pd
import google.generativeai as genai
from google.generativeai import types
from dotenv import load_dotenv  # Import load_dotenv

# Load environment variables from .env file
load_dotenv()

fake = Faker()

def get_excel_cell_address_from_pandas(row_index, col_index):
    """Convert pandas 0-based indices to Excel cell address (A1, B2, etc.)"""
    col_letters = ""
    col_idx_excel = col_index + 1
    while col_idx_excel > 0:
        col_idx_excel, remainder = divmod(col_idx_excel - 1, 26)
        col_letters = chr(65 + remainder) + col_letters
    return f"{col_letters}{row_index + 2}"

from google import genai # Updated import

from google import genai # Updated import

def generate_instruction_with_gemini(project_name, special_column_header, new_value):
    """Generates a Hebrew instruction using Gemini API with loading feedback."""
    gemini_api_key = os.environ.get("GEMINI_API_KEY")
    if not gemini_api_key:
        raise EnvironmentError("GEMINI_API_KEY environment variable not set. Ensure GEMINI_API_KEY is in your .env file or environment variables.")

    # Initialize Gemini Client with API key directly in constructor (newest way)
    client = genai.Client(api_key=gemini_api_key)  # API Key passed here

    model_name = 'gemini-2.0-flash' # Or 'gemini-pro' if you prefer, adjust as needed

    prompt_text = f"""
    Generate a concise natural language instruction in Hebrew to update the '{special_column_header}' of project '{project_name}' to '{new_value}' in an Excel sheet.
    The instruction should be a simple request, as if asking someone to make the change.
    Keep it short and natural-sounding in Hebrew.
    """

    print(f"  ⏳ Generating Hebrew instruction for '{project_name}' with Gemini API...", end=" ", flush=True) # Loading feedback

    try:
        # Call generate_content using the new client and simplified arguments
        response = client.models.generate_content(
            model=model_name,  # Specify the model name
            contents=prompt_text  # Pass the prompt directly as 'contents'
        )
        instruction_hebrew = response.text.strip()
        print("✅ Done.") # End loading feedback - success
        return instruction_hebrew
    except Exception as e:
        print(f"❌ Error: {e}") # End loading feedback - error
        print(f"  ⚠️ Gemini API call failed for '{project_name}'. Skipping instruction generation.")
        return None
    
    
def generate_instruction_json_from_excel_row(excel_file_path, row_data, row_index, headers):
    """Generate instruction and JSON for updating a cell, with Gemini and random column choice."""
    project_name_header = "שם הפרויקט"

    if project_name_header not in headers:
        print(f"Error: Column '{project_name_header}' not found in Excel headers.")
        return None, None, excel_file_path

    project_name = row_data[project_name_header]

    available_columns = [header for header in headers if header != project_name_header]
    if not available_columns:
        print(f"Warning: No columns available to update other than '{project_name_header}'. Skipping row.")
        return None, None, excel_file_path

    special_column_header = random.choice(available_columns)
    special_col_index = headers.get_loc(special_column_header)
    current_value = row_data[special_column_header]
    new_value = None

    project_schema = {
        "data_types": {
            # Add data types for columns if specific handling is needed
        }
    }

    if special_column_header in project_schema["data_types"]:
        data_type = project_schema["data_types"][special_column_header]
        # ... (Specific data type handling if needed) ...
        pass
    else:
        if isinstance(current_value, (int, float)) or (isinstance(current_value, str) and current_value.isdigit()):
            new_value = random.randint(100, 1000)
        else:
            new_value = fake.word()

    if pd.isna(new_value) or new_value is None:
        print(f"Warning: Could not generate valid new value for {special_column_header} in '{project_name}'. Skipping row.")
        return None, None, excel_file_path

    instruction = generate_instruction_with_gemini(project_name, special_column_header, new_value)
    if not instruction:
        return None, None, excel_file_path

    target_cell_address = get_excel_cell_address_from_pandas(row_index, special_col_index)

    function_call = {
        "function_name": "excel_update_cell_by_lookup",
        "parameters": {
            "row_header": project_name_header,
            "row_value": project_name,
            "col_header": special_column_header,
            "new_value": new_value
        },
        "target_cell": target_cell_address,
        "expected_cell_value": new_value
    }

    return instruction, function_call, excel_file_path


if __name__ == "__main__":
    if len(sys.argv) != 2:
        print("Usage: python generate_json_from_excel_gemini_rate_limit_dotenv.py <excel_file_path>")
        sys.exit(1)

    excel_file_path = sys.argv[1]

    if not os.path.exists(excel_file_path):
        print(f"Error: Excel file not found at path: {excel_file_path}")
        sys.exit(1)

    num_examples_to_generate = 50
    synthetic_dataset = []

    try:
        df = pd.read_excel(excel_file_path)
        headers_pd_index = df.columns
        headers = pd.Index(headers_pd_index)
        headers_list = list(headers)

        if "שם הפרויקט" not in headers:
            print(f"Error: Required column 'שם הפרויקט' not found in the Excel file.")
            sys.exit(1)

        example_count = 0
        print("Starting synthetic data generation...") # Start process feedback
        for index, row in df.iterrows():
            if example_count >= num_examples_to_generate:
                break

            print(f"Processing row {index + 2} (Project: '{row['שם הפרויקט'] if 'שם הפרויקט' in row else 'N/A'}')...") # Row processing feedback

            instruction, function_call, filepath = generate_instruction_json_from_excel_row(
                excel_file_path, row, index, headers
            )

            if instruction and function_call:
                data_point = {
                    "instruction": instruction,
                    "context": {
                        "PROJECTS": {
                            "headers": headers_list,
                            "rows": [row.to_dict()]
                        }
                    },
                    "function_call": function_call,
                    "excel_file_path": filepath
                }
                synthetic_dataset.append(data_point)
                print(f"  ✅ Example {len(synthetic_dataset)}/{num_examples_to_generate} generated: {instruction}")
                example_count += 1
            else:
                print(f"  ⚠️ Warning: Failed to generate instruction and JSON for row {index + 2} (Excel row number).")

            time.sleep(5) # Rate limiting: wait for 5 seconds before next iteration (and Gemini API call indirectly)


    except FileNotFoundError:
        print(f"Error: Excel file not found at: {excel_file_path}")
        sys.exit(1)
    except Exception as e:
        print(f"An error occurred during Excel processing: {e}")
        sys.exit(1)

    output_file = os.path.join(os.path.dirname(__file__), "synthetic_excel_data_gemini_hebrew_rate_limit_dotenv.json")
    with open(output_file, 'w', encoding='utf-8') as f:
        json.dump(synthetic_dataset, f, ensure_ascii=False, indent=2)

    print(f"\nGeneration complete: {len(synthetic_dataset)} synthetic examples generated.") # End process feedback
    print(f"Dataset saved to {output_file}")