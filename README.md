# Excel Translator 📊

[![Python 3.10+](https://img.shields.io/badge/python-3.10+-blue.svg)](https://www.python.org/downloads/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](https://opensource.org/licenses/MIT)
[![Code style: black](https://img.shields.io/badge/code%20style-black-000000.svg)](https://github.com/psf/black)

A powerful Excel translation tool that preserves all formatting while translating content using various Large Language Model (LLM) providers. Translate Excel spreadsheets from Chinese to English (or other language pairs) while maintaining exact cell positioning, formulas, styling, and formatting.

## ✨ Features

- **🎯 Formatting Preservation**: Maintains all Excel formatting, fonts, colors, borders, and cell styles
- **⚡ Fast Translation**: Sub-2 second translation speeds with intelligent caching
- **🔄 Batch Processing**: Efficient batch translation with configurable batch sizes
- **🧠 Smart Caching**: Avoids retranslating identical content across sessions
- **📊 Formula Support**: Translates text within Excel formulas while preserving formula logic
- **🔧 Robust Error Handling**: Exponential backoff retry logic for reliability
- **💾 Progress Saving**: Automatic incremental saves to prevent data loss
- **🛡️ Backup Creation**: Automatic backup creation before translation
- **📈 Progress Tracking**: Real-time progress bars and detailed logging
- **🔀 Multiple Providers**: Support for OpenAI GPT models (GPT-4o, GPT-5)
- **🌐 Language Flexibility**: Configurable source and target languages
- **⚙️ Highly Configurable**: Extensive CLI options and environment variable support

## 🚀 Quick Start

### Installation

#### Using UV (Recommended)
```bash
# Install UV if you haven't already
curl -LsSf https://astral.sh/uv/install.sh | sh

# Clone the repository
git clone https://github.com/tristan-mcinnis/Excel-Translator-Formatting-Intact-with-LLMs.git
cd Excel-Translator-Formatting-Intact-with-LLMs

# Install dependencies
uv sync
```

#### Using pip
```bash
# Clone the repository
git clone https://github.com/tristan-mcinnis/Excel-Translator-Formatting-Intact-with-LLMs.git
cd Excel-Translator-Formatting-Intact-with-LLMs

# Create virtual environment
python -m venv venv
source venv/bin/activate  # On Windows: venv\Scripts\activate

# Install dependencies
pip install -e .
```

### Configuration

1. **Copy environment template:**
   ```bash
   cp example.env .env
   ```

2. **Edit `.env` file with your API key:**
   ```bash
   OPENAI_API_KEY=your_openai_api_key_here
   ```

### Basic Usage

```bash
# Simple translation
python main.py input.xlsx --output output.xlsx

# With specific model and languages
python main.py input.xlsx --output output.xlsx \\
  --provider openai --model gpt-5 \\
  --source-lang zh --target-lang en

# Batch processing with custom settings
python main.py input.xlsx --output output.xlsx \\
  --batch-size 10 --context "financial report" \\
  --max-retries 3 --save-interval 50
```

## 📋 Requirements

- **Python**: 3.10 or higher
- **Operating System**: macOS, Linux, or Windows
- **API Keys**: OpenAI API key for translation services
- **Memory**: Minimum 4GB RAM (8GB+ recommended for large files)

## 🔧 Configuration Options

### Command Line Arguments

| Argument | Description | Default |
|----------|-------------|---------|
| `input_file` | Input Excel file path | Required |
| `--output, -o` | Output Excel file path | Required |
| `--provider` | Translation provider (`openai`) | `openai` |
| `--model` | Model name (`gpt-4o`, `gpt-5`) | `gpt-4o` |
| `--source-lang` | Source language code | `zh` |
| `--target-lang` | Target language code | `en` |
| `--context` | Translation context | `spreadsheet` |
| `--batch-size` | Cells per batch | `5` |
| `--max-retries` | Maximum retry attempts | `5` |
| `--save-interval` | Save every N cells | `20` |
| `--no-backup` | Skip backup creation | `False` |
| `--clear-cache` | Clear cache before start | `False` |
| `--cache-dir` | Cache directory path | `translation_cache` |
| `--log-level` | Logging level | `INFO` |
| `--log-file` | Log file path | Auto-generated |

### Environment Variables

| Variable | Description | Default |
|----------|-------------|---------|
| `OPENAI_API_KEY` | OpenAI API key | Required |
| `DEFAULT_SOURCE_LANG` | Default source language | `zh` |
| `DEFAULT_TARGET_LANG` | Default target language | `en` |
| `DEFAULT_CONTEXT` | Default translation context | `spreadsheet` |
| `DEFAULT_BATCH_SIZE` | Default batch size | `5` |
| `LOG_LEVEL` | Logging level | `INFO` |
| `CACHE_DIR` | Cache directory | `translation_cache` |

## 📖 Usage Examples

### Basic Translation
```bash
# Translate Chinese Excel file to English
python main.py chinese_report.xlsx --output english_report.xlsx
```

### Advanced Usage
```bash
# Financial report with custom context
python main.py financial_data.xlsx --output translated_financial.xlsx \\
  --context "financial report with accounting terms" \\
  --batch-size 8 --save-interval 30

# Large file with performance optimization
python main.py large_dataset.xlsx --output translated_dataset.xlsx \\
  --batch-size 15 --max-retries 3 \\
  --log-level DEBUG
```

### Using Different Models
```bash
# Use GPT-5 for higher quality translation
python main.py input.xlsx --output output.xlsx \\
  --model gpt-5 --context "technical documentation"

# Use GPT-4o for faster translation
python main.py input.xlsx --output output.xlsx \\
  --model gpt-4o --batch-size 10
```

## 🏗️ Project Structure

```
excel_translator/
├── excel_translator/           # Main package
│   ├── __init__.py            # Package initialization
│   ├── cli.py                 # Command-line interface
│   ├── translation.py         # Core translation logic
│   ├── utils.py               # Utility functions
│   └── providers/             # Translation providers
│       ├── __init__.py
│       ├── base_provider.py   # Abstract base provider
│       └── openai_provider.py # OpenAI implementation
├── tests/                     # Test suite
├── main.py                    # Entry point
├── pyproject.toml            # Project configuration
├── example.env               # Environment template
├── README.md                 # This file
└── CLAUDE.md                 # Development guidelines
```

## 🧪 Testing

```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=excel_translator --cov-report=html

# Run specific test file
pytest tests/test_translation.py

# Run with verbose output
pytest -v
```

## 🔍 Supported File Types

- **Excel Files**: `.xlsx`, `.xlsm`
- **Language Detection**: Automatic Chinese character detection
- **Formula Support**: Excel formulas with embedded text strings
- **Formatting**: All Excel formatting elements (fonts, colors, borders, etc.)

## ⚡ Performance Tips

1. **Optimize Batch Size**: Start with default (5), increase for faster processing
2. **Use Caching**: Keep cache enabled to avoid retranslating identical content
3. **Save Intervals**: Adjust based on file size (smaller intervals for large files)
4. **Model Selection**: Use GPT-4o for speed, GPT-5 for quality
5. **Context Specification**: Provide specific context for better translations

## 🛠️ Development

### Development Installation
```bash
# Install with development dependencies
uv sync --extra dev

# Or with pip
pip install -e ".[dev]"
```

### Code Quality Tools
```bash
# Format code
black excel_translator/

# Sort imports
isort excel_translator/

# Lint code
flake8 excel_translator/

# Type checking
mypy excel_translator/
```

### Running Tests
```bash
# Run all tests
pytest

# Run with coverage
pytest --cov=excel_translator

# Run specific tests
pytest tests/test_providers.py -v
```

## 🤝 Contributing

1. **Fork** the repository
2. **Create** a feature branch (`git checkout -b feature/amazing-feature`)
3. **Commit** your changes (`git commit -m 'Add amazing feature'`)
4. **Push** to the branch (`git push origin feature/amazing-feature`)
5. **Open** a Pull Request

### Development Guidelines

- Follow PEP 8 style guidelines
- Add type hints to all functions
- Write tests for new features
- Update documentation for API changes
- Use meaningful commit messages

## 📝 License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.

## 🙏 Acknowledgments

- **OpenAI** for providing powerful language models
- **openpyxl** library for Excel file manipulation
- **tqdm** for progress tracking
- The open-source community for various tools and libraries

## 📞 Support

- **Issues**: [GitHub Issues](https://github.com/tristan-mcinnis/Excel-Translator-Formatting-Intact-with-LLMs/issues)
- **Documentation**: This README and inline code documentation
- **Examples**: See the `examples/` directory for sample usage

## 🚧 Roadmap

- [ ] Support for additional LLM providers (Anthropic, DeepSeek, Grok)
- [ ] Web interface for non-technical users
- [ ] Support for additional file formats (.xls, .csv)
- [ ] Real-time collaborative translation
- [ ] Translation quality metrics and validation
- [ ] Custom translation models fine-tuning

---

**Made with ❤️ by [Tristan McInnis](https://github.com/tristan-mcinnis)**