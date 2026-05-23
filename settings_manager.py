# File: settings_manager.py
import os
import json

SETTINGS_FILE = os.path.join(os.path.dirname(__file__), "temps", "app_settings.json")

DEFAULT_PROMPT = """please extract header and item infromation from this file and put in json format. return only json without any additional text
this json template is your reference only. json should content actual extracted data

{
  "header": {
    "header_infor1": "content of header infor 1",
    "header_infor2": "content of header infor 2"
  },
  "items": [
    {
      "item1_info1": "content of item1 infor 1",
      "item1_info2": "content of item1 infor 2"
    },
    {
      "item2_info1": "content of item2 infor 1",
      "item2_info2": "content of item2 infor 2"
    }
  ],
  "other infor": {
    "Quantity": "12.282",
    "FOB Amount": "$106.162.950"
  }
}"""

DEFAULT_SETTINGS = {
    "input_folder": "",
    "output_folder": "",
    "ai_engine": "Local AI",
    "gemini_api_key": "",
    "gpt_api_key": "",
    "local_api_url": "http://api-localai.germantest.net",
    "gemini_model": "gemini-2.5-flash",
    "gpt_model": "gpt-4o-mini",
    "local_model": "qwen2.5vl:7b",
    "batch_limit": "All Files",
    "inter_file_delay": 2,
    "model_fail_delay": 5,
    "model_fail_attempts": 3,
    "custom_prompt": DEFAULT_PROMPT,
    "selected_columns": [], # Holds last selected schema columns
    "selected_schema": None
}

def load_settings() -> dict:
    """Loads settings from JSON file or returns defaults."""
    os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
    if os.path.exists(SETTINGS_FILE):
        try:
            with open(SETTINGS_FILE, "r", encoding="utf-8") as f:
                data = json.load(f)
                # Merge loaded data with defaults to ensure all keys exist
                settings = DEFAULT_SETTINGS.copy()
                settings.update(data)
                return settings
        except Exception as e:
            print(f"Error loading settings: {e}")
    return DEFAULT_SETTINGS.copy()

def save_settings(settings: dict):
    """Saves settings to JSON file."""
    os.makedirs(os.path.dirname(SETTINGS_FILE), exist_ok=True)
    try:
        with open(SETTINGS_FILE, "w", encoding="utf-8") as f:
            json.dump(settings, f, indent=4, ensure_ascii=False)
    except Exception as e:
        print(f"Error saving settings: {e}")
