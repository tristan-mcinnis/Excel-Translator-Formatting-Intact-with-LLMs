"""Tests for utility functions."""

import pytest
import tempfile
import os
import json
from excel_translator.utils import is_chinese, get_cache_key, load_cache, save_cache


class TestUtilityFunctions:
    """Test utility functions."""

    def test_is_chinese_with_chinese_text(self):
        """Test Chinese text detection with Chinese characters."""
        assert is_chinese("你好") is True
        assert is_chinese("中文测试") is True
        assert is_chinese("Hello 你好") is True
        assert is_chinese("测试123") is True

    def test_is_chinese_with_non_chinese_text(self):
        """Test Chinese text detection with non-Chinese text."""
        assert is_chinese("Hello") is False
        assert is_chinese("123") is False
        assert is_chinese("Test") is False
        assert is_chinese("") is False

    def test_is_chinese_with_invalid_input(self):
        """Test Chinese text detection with invalid input."""
        assert is_chinese(None) is False
        assert is_chinese(123) is False
        assert is_chinese([]) is False

    def test_get_cache_key(self):
        """Test cache key generation."""
        key1 = get_cache_key("test")
        key2 = get_cache_key("test")
        key3 = get_cache_key("different")

        assert key1 == key2  # Same input should give same key
        assert key1 != key3  # Different input should give different key
        assert len(key1) == 32  # MD5 hash length

    def test_cache_operations(self):
        """Test cache save and load operations."""
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_file = os.path.join(temp_dir, "test_cache.json")

            # Test saving cache
            test_cache = {"key1": "value1", "key2": "value2"}
            save_cache(test_cache, cache_file)

            # Test loading cache
            loaded_cache = load_cache(cache_file)
            assert loaded_cache == test_cache

    def test_load_nonexistent_cache(self):
        """Test loading cache from non-existent file."""
        cache = load_cache("nonexistent_file.json")
        assert cache == {}

    def test_save_cache_creates_directory(self):
        """Test that save_cache creates directory if it doesn't exist."""
        with tempfile.TemporaryDirectory() as temp_dir:
            cache_file = os.path.join(temp_dir, "subdir", "cache.json")
            test_cache = {"test": "data"}

            save_cache(test_cache, cache_file)

            assert os.path.exists(cache_file)
            with open(cache_file, 'r') as f:
                loaded_data = json.load(f)
            assert loaded_data == test_cache