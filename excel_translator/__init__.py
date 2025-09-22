"""Excel Translator - Preserving formatting while translating Excel files with LLMs."""

__version__ = "1.0.0"
__author__ = "Tristan McInnis"
__email__ = "your.email@example.com"
__description__ = "A powerful Excel translation tool that preserves all formatting while translating content using various LLM providers."

from .translation import ExcelTranslator
from .cli import main

__all__ = ["ExcelTranslator", "main"]