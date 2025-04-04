import json
import sys  # To get command-line arguments

def remove_fields_from_json_file(json_filepath, output_filepath):
    """
    Loads JSON data from a file, removes 'response' and 'excel_row_number' fields
    from each object in the array, and saves the modified JSON to a new file.

    Args:
        json_filepath (str): The path to the input JSON file.
        output_filepath (str): The path to the output JSON file where modified data will be saved.
    """
    try:
        with open(json_filepath, 'r', encoding='utf-8') as f: # Open input file in read mode, specify encoding for Hebrew
            data = json.load(f) # Load JSON data from the file
    except FileNotFoundError:
        print(f"Error: Input file not found at path: {json_filepath}")
        return
    except json.JSONDecodeError as e:
        print(f"Error: Invalid JSON format in file: {json_filepath}")
        print(f"JSONDecodeError details: {e}") # Print detailed error message
        return

    modified_data = []
    for item in data:
        modified_item = {}
        for key, value in item.items():
            if key not in ["response", "excel_row_number"]:
                modified_item[key] = value
        modified_data.append(modified_item)

    modified_json_string = json.dumps(modified_data, indent=2, ensure_ascii=False)

    try:
        with open(output_filepath, 'w', encoding='utf-8') as outfile: # Open output file in write mode, specify encoding
            outfile.write(modified_json_string) # Write the modified JSON string to the output file
        print(f"Modified JSON data saved to: {output_filepath}")
    except Exception as e:
        print(f"Error: Failed to write to output file: {output_filepath}")
        print(f"Error details: {e}")


if __name__ == "__main__":
    input_json_file = "gemini_raw_responses.json" # Default input filename
    output_json_file = "modified_gemini_responses.json" # Default output filename

    if len(sys.argv) > 1:
        input_json_file = sys.argv[1] # Get input filename from command line argument
    if len(sys.argv) > 2:
        output_json_file = sys.argv[2] # Get output filename from command line argument

    print(f"Processing input file: {input_json_file}")
    print(f"Saving output to file: {output_json_file}")

    remove_fields_from_json_file(input_json_file, output_json_file)