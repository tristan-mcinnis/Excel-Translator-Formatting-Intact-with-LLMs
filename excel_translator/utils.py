"""Utility functions for Excel translation."""

import os
import re
import json
import hashlib
from typing import Any
import logging

logger = logging.getLogger(__name__)


def is_chinese(text: str) -> bool:
    """Check if text contains Chinese characters.

    Args:
        text: Text to check

    Returns:
        True if text contains Chinese characters
    """
    if not isinstance(text, str):
        return False

    # Chinese Unicode ranges (expanded):
    # Basic Chinese: 4E00-9FFF
    # Extended Chinese A: 3400-4DBF
    # Extended Chinese B: 20000-2A6DF
    # Extended Chinese C: 2A700-2B73F
    # Extended Chinese D: 2B740-2B81F
    # Extended Chinese E: 2B820-2CEAF
    # CJK Radicals: 2F00-2FDF
    # CJK Symbols and Punctuation: 3000-303F
    chinese_pattern = re.compile(r'[\u4E00-\u9FFF\u3400-\u4DBF\u20000-\u2A6DF\u2A700-\u2B73F\u2B740-\u2B81F\u2B820-\u2CEAF\u2F00-\u2FDF\u3000-\u303F]')
    return bool(chinese_pattern.search(text))


def get_cache_key(text: str) -> str:
    """Generate a unique cache key for the text.

    Args:
        text: Text to generate key for

    Returns:
        MD5 hash of the text
    """
    return hashlib.md5(text.encode('utf-8')).hexdigest()


def load_cache(cache_file: str) -> dict:
    """Load cache from disk.

    Args:
        cache_file: Path to cache file

    Returns:
        Cache dictionary
    """
    cache = {}
    try:
        if os.path.exists(cache_file):
            with open(cache_file, 'r', encoding='utf-8') as f:
                cache = json.load(f)
            logger.info(f"Loaded {len(cache)} translations from cache")
    except Exception as e:
        logger.warning(f"Failed to load cache: {e}")
    return cache


def save_cache(cache: dict, cache_file: str) -> None:
    """Save cache to disk.

    Args:
        cache: Cache dictionary to save
        cache_file: Path to cache file
    """
    try:
        os.makedirs(os.path.dirname(cache_file), exist_ok=True)
        with open(cache_file, 'w', encoding='utf-8') as f:
            json.dump(cache, f, ensure_ascii=False, indent=2)
        logger.debug(f"Saved {len(cache)} translations to cache")
    except Exception as e:
        logger.warning(f"Failed to save cache: {e}")


def setup_logging(log_level: str = "INFO", log_file: str = None) -> None:
    """Set up logging configuration.

    Args:
        log_level: Logging level (DEBUG, INFO, WARNING, ERROR)
        log_file: Optional log file path
    """
    handlers = [logging.StreamHandler()]

    if log_file:
        handlers.append(logging.FileHandler(log_file))

    logging.basicConfig(
        level=getattr(logging, log_level.upper()),
        format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
        handlers=handlers
    )