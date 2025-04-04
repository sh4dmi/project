import json

def convert_json_structure(input_filepath, output_filepath):
    """
    Converts a JSON file from the given input structure to the desired output structure
    for write_cell_scenarios.json.

    Args:
        input_filepath (str): Path to the input JSON file.
        output_filepath (str): Path to save the converted JSON file.
    """
    try:
        with open(input_filepath, 'r', encoding='utf-8') as f:
            data = json.load(f)
    except FileNotFoundError:
        print(f"Error: Input file not found at {input_filepath}")
        return
    except json.JSONDecodeError:
        print(f"Error: Invalid JSON format in {input_filepath}")
        return

    converted_scenarios = []
    for item in data:
        parsed_json = item.get("parsed_function_call_json")
        if parsed_json:
            instruction = parsed_json.get("instruction")
            parameters = parsed_json.get("parameters")

            if instruction and parameters:
                expected_params = {
                    "row_header": parameters.get("row_header"),
                    "row_value": parameters.get("row_value"),
                    "col_header": parameters.get("col_header"),
                    "new_value": parameters.get("new_value")
                }

                converted_scenario = {
                    "prompt": instruction,
                    "expected_params": expected_params
                }
                converted_scenarios.append(converted_scenario)
            else:
                print(f"Warning: Missing 'instruction' or 'parameters' in parsed_function_call_json for an item.")
        else:
            print(f"Warning: 'parsed_function_call_json' key not found in an item.")

    try:
        with open(output_filepath, 'w', encoding='utf-8') as outfile:
            json.dump(converted_scenarios, outfile, indent=4, ensure_ascii=False) # ensure_ascii=False to handle Hebrew characters
        print(f"Successfully converted JSON structure and saved to {output_filepath}")
    except Exception as e:
        print(f"Error: Failed to save converted JSON to {output_filepath}: {e}")

if __name__ == "__main__":
    input_json_file = "write_cell_scenarios.json"  # Replace with your input file name
    output_json_file = "write_cell_scenarios2.json" # Replace with your desired output file name

    convert_json_structure(input_json_file, output_json_file)
    print(f"\nConversion process finished. Check '{output_json_file}' for the result.")