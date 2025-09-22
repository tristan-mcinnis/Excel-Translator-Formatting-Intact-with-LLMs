"""Base provider class for translation services."""

from abc import ABC, abstractmethod
from typing import List, Dict, Optional
import logging

logger = logging.getLogger(__name__)


class BaseProvider(ABC):
    """Abstract base class for translation providers."""

    def __init__(self, api_key: str, model: str, timeout: float = 90.0):
        """Initialize the provider.

        Args:
            api_key: API key for the service
            model: Model name to use
            timeout: Request timeout in seconds
        """
        self.api_key = api_key
        self.model = model
        self.timeout = timeout

    @abstractmethod
    async def translate_batch(
        self,
        texts: List[str],
        source_lang: str = "zh",
        target_lang: str = "en",
        context: str = "spreadsheet"
    ) -> Dict[str, str]:
        """Translate a batch of texts.

        Args:
            texts: List of texts to translate
            source_lang: Source language code
            target_lang: Target language code
            context: Translation context

        Returns:
            Dict mapping original texts to translations
        """
        pass

    @abstractmethod
    async def translate_single(
        self,
        text: str,
        source_lang: str = "zh",
        target_lang: str = "en",
        context: str = "spreadsheet"
    ) -> str:
        """Translate a single text.

        Args:
            text: Text to translate
            source_lang: Source language code
            target_lang: Target language code
            context: Translation context

        Returns:
            Translated text
        """
        pass