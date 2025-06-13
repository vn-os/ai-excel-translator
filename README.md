# EXCEL FILE TRANSLATION TOOL GUIDE

## Introduction

This is an automatic tool for translating Excel file content between Japanese, Vietnamese, and English, using an LLM API.
This tool can translate text in cells and shapes within Excel files.

## Requirements

1. Python must be installed (Python 3.7 or higher is recommended).
2. Microsoft Excel must be installed (tool uses xlwings library which requires Excel).
3. This tool only works on Windows or macOS (xlwings requires Excel on these platforms).
4. Install the required libraries with the following command:

   ```
   pip install -r requirements.txt
   ```

   (This file will be created automatically when you run the program for the first time)

## API Setup

1. Configure your LLM API settings in a .env file with the following format:
   ```
   LLM_API_URL=your_api_url_here
   LLM_API_KEY=your_api_key_here
   LLM_MODEL_NAME=your_model_name_here
   ```
2. Place the .env file in the same directory as the main.py file

## How to Use

1. Run the main.py file for the first time to create the directory structure:

   ```
   python main.py
   ```
2. Place the Excel files to be translated in the "input" folder
3. Run the program with optional parameters to specify the source and target languages:

   - Translate from Japanese to English (default):
     ```
     python main.py
     ```
   - Translate from Japanese to Vietnamese:
     ```
     python main.py --to vi
     ```
   - Translate from Vietnamese to Japanese:
     ```
     python main.py --from vi --to ja
     ```
   - Translate from Vietnamese to English:
     ```
     python main.py --from vi
     ```
4. Translation results will be saved in the "output" folder

## Language Support

The tool currently supports the following languages:
- Japanese (ja)
- Vietnamese (vi)
- English (en)

Default settings:
- Source language: Japanese (ja)
- Target language: English (en)

## Supported Languages

The script supports translation between the following languages:

| Code | Language    | Code | Language     |
|------|-------------|------|--------------|
| en   | English     | ko   | Korean       |
| zh   | Chinese     | id   | Indonesian   |
| hi   | Hindi       | fr   | French       |
| es   | Spanish     | de   | German       |
| ar   | Arabic      | tr   | Turkish      |
| bn   | Bengali     | it   | Italian      |
| pt   | Portuguese  | th   | Thai         |
| ru   | Russian     | pl   | Polish       |
| ja   | Japanese    | uk   | Ukrainian    |
| vi   | Vietnamese  | nl   | Dutch        |

## API Performance Customization

The default configuration is set to work with most LLM APIs. You can customize these settings based on your API provider and model:

1. **API Delay Adjustment** - The current delay between API calls is set to 2 seconds to avoid hitting rate limits:

   ```python
   # Find this line near the beginning of the script
   API_DELAY = 2  # Delay 2 seconds between API calls
   ```

   - For models with higher rate limits, you can reduce this value
   - For free tier APIs with stricter limits, you might need to increase this value

2. **Batch Size Adjustment** - The default batch size is 100 cells/shapes per API call:

   ```python
   # Find this line near the beginning of the script
   BATCH_SIZE = 100  # Maximum number of cells in a batch
   ```

   - Increase this value for more powerful models that can handle larger contexts
   - Decrease this value if you're getting context length errors or incomplete translations

3. **Using Different API Providers** - Configure your API settings in the .env file:

   ```
   LLM_API_URL=your_api_url_here
   LLM_API_KEY=your_api_key_here
   LLM_MODEL_NAME=your_model_name_here
   ```

   Examples for different providers:
   - For OpenAI: `https://api.openai.com/v1/`
   - For Azure OpenAI: `https://your-resource.openai.azure.com/openai/deployments/your-deployment/`
   - For LM Studio: `http://localhost:1234/v1`

## Customizing System Prompt

The default system prompt is optimized for general translations. You can customize it by:

1. Locate the `system-prompt.txt` file in your project directory (it's created automatically after first run)
2. Open it with a text editor
3. Modify the content to match your specific requirements

## Directory Structure

- /main.py: Main program file
- /input/: Directory containing Excel files to be translated
- /output/: Directory containing translated Excel files
- /system-prompt.txt: File containing system prompt (will be created automatically)
- /requirements.txt: File containing library requirements (will be created automatically)
- /.env: File containing API settings (needs to be created manually)

## Features

- Translates text content in Excel cells
- Translates text content in shapes such as TextBox, WordArt, etc.
- Preserves the original format of the Excel file
- Skips cells that only contain numbers, formulas, or very short content
- Processes multiple Excel files in a directory
- Supports multiple language pairs (Japanese, Vietnamese, English)

## Notes

- The translation process may take time depending on the amount of content to be translated
- Translation quality depends on the LLM API being used
- Do not edit Excel files while the program is running

## Troubleshooting

If you encounter errors:

1. Check if the API key is correctly set up in the .env file
2. Ensure all dependent libraries are installed
3. Check file and directory access permissions

## Usage

1. Place your Excel files in the `input` directory
2. Run the script with source and target language codes:

```bash
python main.py --from <source_lang> --to <target_lang>
```

For example:
```bash
# Translate from Japanese to English
python main.py --from ja --to en

# Translate from English to Vietnamese
python main.py --from en --to vi

# Translate from Chinese to Spanish
python main.py --from zh --to es
```

The translated files will be saved in the `output` directory with the target language code appended to the filename.
