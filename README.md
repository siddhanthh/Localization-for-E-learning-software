# DOCX Document Translator

An enterprise-grade, high-performance Python utility designed to translate specific columns in DOCX document tables while maintaining the exact original layout and styles. This tool is optimized for localizing e-learning modules exported from **Articulate Storyline** or **Articulate Rise 360**.

### Articulate E-Learning Localization Workflow
1. **Export**: Export your translation document from Articulate Storyline or Rise 360 as a `.docx` file.
2. **Translate**: Run this utility pointing to the exported `.docx` file and specifying the desired target language. The tool parses the tables, translates only the target column (default header `"Translation"`), and maintains all formatting and styling configurations.
3. **Import**: Import the resulting translated `.docx` file back into Articulate Storyline or Rise 360 to instantly update your course slides and contents.

## Features

- **Format Preservation**: Automatically copies style structures and dimensions from the source document to ensure visual parity.
- **Multiple Translation Engines**: Supports fallbacks to `deep-translator` (Google Translate) or advanced contextual translation utilizing Google's Gemini API (with support for system instructions).
- **Concurrency**: High-throughput translation utilizing a thread pool executor to fetch translations concurrently before applying them.
- **Translation Memory & Glossary**: Localized exact, normalized, and case-insensitive caching of translation keys to minimize API requests and ensure terminological consistency.
- **Robust Tag-Based Translation**: Translates paragraphs containing multiple inline format runs using inline XML tagging, validating tags in response, and falling back gracefully to run-by-run mode if necessary.
- **Dynamic Script Detection**: Automatically skips translation for strings containing native characters matching the target language script or matching protected syntax patterns (URLs, variables, placeholders, product codes).

---

## Installation

### Prerequisites

- Python 3.8 or higher

### 1. Setup Virtual Environment

Clone or copy the project to your computer, navigate into the directory, and initialize the Python virtual environment:

```bash
# Create the virtual environment
python -m venv venv

# Activate the virtual environment
# On Windows (Command Prompt)
venv\Scripts\activate.bat
# On Windows (PowerShell)
venv\Scripts\Activate.ps1
# On macOS / Linux
source venv/bin/activate
```

### 2. Install Dependencies

Install the required packages:

```bash
pip install -r requirements.txt
```

*(Optional)* Install this utility as a CLI tool in editable mode:

```bash
pip install -e .
```

---

## Configuration

You can customize execution parameters using a JSON configuration file. By default, the application falls back to built-in settings.

Create a file named `config.json` in your workspace:

```json
{
  "translation_backend": "google_translator",
  "gemini": {
    "api_key": "",
    "model": "gemini-1.5-flash",
    "endpoint": "https://generativelanguage.googleapis.com/v1beta/models/{model}:generateContent"
  },
  "rate_limiting": {
    "max_requests_per_minute": 300,
    "max_tokens_per_minute": 40000,
    "max_batch_size": 20,
    "retry_backoff_factor": 2.0,
    "max_retries": 5
  },
  "concurrency": {
    "num_workers": 8
  }
}
```

### Environment Variables
For security, keep API keys out of your code files. Set the following environment variables to override default values:
- `GEMINI_API_KEY`: Your Gemini API access key.
- `TRANSLATION_BACKEND`: Override engine selection (e.g. `gemini` or `google_translator`).

---

## Usage

Once installed, use the `docx-translate` command directly, or run the package entry module via Python.

### Standard Translate (Single Step)
Inputs the `.docx` document and translates the `Translation` column to the target language:

```bash
# Using CLI entry point
docx-translate input.docx th

# Or running through python directly
python -m docx_translator.cli input.docx th
```
*Output will be saved as `input_th.docx`.*

### Command-line Reference

```
usage: docx-translate [-h] [--output OUTPUT] [--source SOURCE] [--config CONFIG]
                      [--backend {google_translator,gemini}] [--tm TM]
                      [--glossary GLOSSARY] [--column-header COLUMN_HEADER]
                      [--analyze-only] [--test-mode] [--review-report REVIEW_REPORT]
                      [--verbose]
                      input target

AI-Powered DOCX Translation Tool

positional arguments:
  input                 Path to the input Word Document (.docx)
  target                Target language code (e.g., 'th', 'es', 'fr', 'ru', 'zh', 'ja')

options:
  -h, --help            show this help message and exit
  --output OUTPUT, -o OUTPUT
                        Path to save the translated document. Default: <input_name>_<target>.docx
  --source SOURCE, -s SOURCE
                        Source language of the document (default: 'en')
  --config CONFIG, -c CONFIG
                        Path to custom JSON configuration file
  --backend {google_translator,gemini}, -b {google_translator,gemini}
                        Override the translation backend (google_translator or gemini)
  --tm TM               Path to Translation Memory JSON file (default: translation_memory.json)
  --glossary GLOSSARY   Path to Glossary JSON file (default: glossary.json)
  --column-header COLUMN_HEADER
                        The header name of the table column containing text to translate (default: 'Translation')
  --analyze-only        Analyze the document structures, print translation statistics, and exit
  --test-mode           Run translation in test mode (limits execution to 5 tables / 50 cells)
  --review-report REVIEW_REPORT
                        Path to save the post-translation CSV review report
  -v, --verbose         Enable detailed debug logs
```

### Examples

#### Spanish Translation with Gemini Engine
```bash
# Set key
set GEMINI_API_KEY=AIzaSy...

docx-translate course_export.docx es --backend gemini --review-report review_es.csv
```

#### Analyze a document without translating
```bash
docx-translate course_export.docx th --analyze-only
```

---

## Project Structure

```
docx-translator/
├── .gitignore              # Files to ignore in Git repository
├── README.md               # User manual and configuration details
├── requirements.txt        # Python dependency list
├── setup.py                # Setup script for setuptools packaging
├── docx_translator/        # Core package folder
│   ├── __init__.py         # Package level initialization
│   ├── cli.py              # Command-Line Parser and runner
│   ├── config.py           # Configuration parser and settings
│   ├── layout.py           # Formatting logic
│   ├── pipeline.py         # Table parsing and pipeline runner
│   └── translator.py       # Rate-limited translation client classes
```
