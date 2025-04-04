import json
from datasets import Dataset, Features, Value, Sequence
from huggingface_hub import HfApi, HfFolder

# --- Configuration ---
input_json_path = 'gemini_parsed_responses_details.json'  # <--- CHANGE THIS to the actual path of your JSON file
hf_repo_id = "SH4DMI/XLSX1"  # <--- CHANGE THIS to your desired Hugging Face repo ID (e.g., "jsmith/excel-instructions-he")
# --- End Configuration ---

def transform_data(input_path):
    """Loads the original JSON, transforms it, and returns a list of dictionaries."""
    transformed_entries = []
    try:
        with open(input_path, 'r', encoding='utf-8') as f:
            original_data = json.load(f)
    except FileNotFoundError:
        print(f"Error: Input file not found at '{input_path}'")
        return None
    except json.JSONDecodeError:
        print(f"Error: Could not decode JSON from '{input_path}'. Check file format.")
        return None

    print(f"Loaded {len(original_data)} entries from {input_path}")

    for entry in original_data:
        try:
            # Extract original instruction
            instruction = entry['parsed_function_call_json']['instruction']

            # Create the ground truth function JSON object
            ground_truth_obj = {
                "function_name": entry['parsed_function_call_json']['function_name'],
                "parameters": entry['parsed_function_call_json']['parameters']
            }
            # Convert the ground truth object to a JSON string
            ground_truth_string = json.dumps(ground_truth_obj, ensure_ascii=False) # ensure_ascii=False preserves Hebrew characters

            # Extract headers
            excel_headers = entry['excel_headers']

            # Create the new entry for the dataset
            transformed_entries.append({
                "instruction": instruction,
                "ground_truth_function": ground_truth_string,
                "excel_headers": excel_headers
            })
        except KeyError as e:
            print(f"Warning: Skipping entry due to missing key: {e}. Entry data: {entry}")
        except Exception as e:
            print(f"Warning: Skipping entry due to unexpected error: {e}. Entry data: {entry}")


    print(f"Successfully transformed {len(transformed_entries)} entries.")
    return transformed_entries

def push_to_huggingface(data_list, repo_id):
    """Creates a Hugging Face Dataset and pushes it to the Hub."""
    if not data_list:
        print("No data to push.")
        return

    print(f"\nPreparing to push data to Hugging Face repository: {repo_id}")

    # Define the features (schema) of the dataset
    features = Features({
        'instruction': Value('string'),
        'ground_truth_function': Value('string'),
        'excel_headers': Sequence(Value('string'))
    })

    # Create Hugging Face Dataset object from the list of dictionaries
    # Using from_list is generally efficient for moderate sized lists
    try:
        hf_dataset = Dataset.from_list(data_list, features=features)
        print("Hugging Face Dataset object created.")

        # Push the dataset to the Hub
        print(f"Pushing dataset to '{repo_id}'...")
        hf_dataset.push_to_hub(repo_id)
        print("\nDataset successfully pushed to Hugging Face Hub!")
        print(f"You can view it at: https://huggingface.co/datasets/{repo_id}")

    except Exception as e:
        print(f"\nError during Hugging Face dataset creation or push: {e}")
        print("Please check:")
        print("1. You are logged in (`huggingface-cli login`).")
        print("2. The repository ID is correct and you have write access.")
        print("3. Your internet connection.")

# --- Main Execution ---
if __name__ == "__main__":
    # 1. Transform the data
    transformed_data = transform_data(input_json_path)

    # 2. Push to Hugging Face (only if transformation was successful)
    if transformed_data:
        # Make sure you are logged in before running this
        push_to_huggingface(transformed_data, hf_repo_id)
    else:
        print("Data transformation failed. Nothing pushed to Hugging Face.")