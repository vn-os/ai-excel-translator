# AI Excel Translator

A Python tool that uses AI to translate Excel files while preserving formatting. The tool processes both cell contents and shapes (text boxes, comments, etc.) in Excel files.

## Features

- Translates Excel files while preserving all formatting
- Supports translation of text in cells and shapes
- Processes multiple Excel files in batch
- Preserves Excel formulas and special formatting
- Handles multiple sheets within workbooks
- Supports 20 major languages

## Requirements

1. Python 3.7 or higher
2. Microsoft Excel must be installed (tool uses xlwings library which requires Excel).
3. This tool only works on Windows or macOS (xlwings requires Excel on these platforms).

## Setup

1. Clone this repository
2. Install required packages:
   ```bash
   pip install -r requirements.txt
   ```
3. Create a `.env` file in the project root with your LLM API settings:
   ```
   LLM_API_URL=your_api_url
   LLM_API_KEY=your_api_key
   LLM_MODEL_NAME=your_model_name
   ```

   Examples:
   ```
   *Gemini*
   LLM_API_URL=https://generativelanguage.googleapis.com/v1beta/openai/
   LLM_API_KEY=
   LLM_MODEL_NAME=gemini-2.0-flash-lite
   LLM_MODEL_SUFFIX=

   *Locally with LM Studio*
   LLM_API_URL=http://localhost:1234/v1/
   LLM_API_KEY=lm-studio
   LLM_MODEL_NAME=qwen3-1.7b
   LLM_MODEL_NO_THINK=1
   LLM_MODEL_SUFFIX=/no_think

   ...
   ```

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

You can customize the API performance by setting these optional environment variables in your `.env` file:

- `LLM_MODEL_SUFFIX`: Additional instructions for the LLM model
- `LLM_MODEL_NO_THINK`: Set to "1" or "true" to remove thinking steps from the model's output

## Troubleshooting

If you encounter issues:

1. Ensure all dependent libraries are installed
2. Check file and directory access permissions
3. Verify your API credentials in the `.env` file
4. Make sure Excel is properly installed and accessible
5. Check that input files are valid Excel files (.xlsx or .xls)

## License

This project is licensed under the MIT License - see the LICENSE file for details.
