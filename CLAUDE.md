# CLAUDE.md

This file provides guidance to Claude Code (claude.ai/code) when working with code in this repository.

## Project Overview

This is a Python-based Excel translation tool that translates Chinese text to English while preserving all formatting, formulas, and cell positioning. The tool uses OpenAI's GPT-4o model with streaming API for efficient translation and includes caching to avoid retranslating the same text.

## Environment Setup

The project requires:
```bash
pip install -r requirements.txt
```

Key dependencies include OpenAI API client, openpyxl for Excel manipulation, pandas, and asyncio for concurrent processing.

## Configuration

The tool requires environment variables in `.env`:
- `OPENAI_API_KEY`: Required for translation API access
- Other optional API keys for extended functionality

**Important**: The `.env` file contains sensitive API keys and should never be committed to version control.

## Running the Tool

Basic usage:
```bash
python ExcelTranslate.py --input input_file.xlsx --output output_file.xlsx
```

Common options:
- `--context "questionnaire"`: Provide translation context
- `--batch-size 10`: Number of cells to translate per batch (default: 5)
- `--max-retries 3`: Retry attempts for failed translations (default: 5)
- `--save-interval 50`: Save progress every N cells (default: 20)
- `--no-backup`: Skip creating backup file
- `--clear-cache`: Clear translation cache before starting

## Testing

Run tests using:
```bash
pytest
```

## Architecture

The codebase consists of a single main file `ExcelTranslate.py` with these key components:

### Core Translation Functions
- `stream_translation()`: Async streaming translation using OpenAI API
- `batch_translate_texts_async()`: Processes multiple texts concurrently with retry logic
- `is_chinese()`: Detects Chinese characters using Unicode ranges

### Excel Processing
- `translate_excel_file()`: Main function that preserves exact cell positioning and formatting
- Handles both regular cells and formula cells separately
- Creates backups and saves progress incrementally

### Caching System
- JSON-based translation cache in `translation_cache/` directory
- Cache keys generated using MD5 hashes of original text
- Persistent across runs to avoid retranslating the same content

### Error Handling & Resilience
- Exponential backoff retry logic for API failures
- Signal handling for graceful interruption (Ctrl+C)
- Automatic progress saving at configurable intervals
- Comprehensive logging to timestamped log files

## File Structure

- `ExcelTranslate.py`: Main application file
- `requirements.txt`: Python dependencies
- `.env`: Environment variables (not in repo)
- `translation_cache/`: Directory for cached translations
- `*.log`: Generated log files with timestamps