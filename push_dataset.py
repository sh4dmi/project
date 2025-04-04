from huggingface_hub import HfApi

api = HfApi()
api.upload_file(
    path_or_fileobj="synthetic_excel_data_gemini_hebrew_headers_raw_json_responses.json",
    path_in_repo="data.json",
    repo_id="SH4DMI/XLSX_JSON",
    repo_type="dataset",
)
