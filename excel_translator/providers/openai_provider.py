"""OpenAI provider for translation services."""

import asyncio
from typing import List, Dict
import logging
from openai import OpenAI

from .base_provider import BaseProvider
from ..utils import is_chinese, get_cache_key

logger = logging.getLogger(__name__)


class OpenAIProvider(BaseProvider):
    """OpenAI translation provider."""

    def __init__(self, api_key: str, model: str = "gpt-4o", timeout: float = 90.0):
        """Initialize OpenAI provider.

        Args:
            api_key: OpenAI API key
            model: OpenAI model to use (default: gpt-4o, can use gpt-5 when available)
            timeout: Request timeout in seconds
        """
        super().__init__(api_key, model, timeout)
        self.client = OpenAI(api_key=api_key, timeout=timeout)

    async def translate_single(
        self,
        text: str,
        source_lang: str = "zh",
        target_lang: str = "en",
        context: str = "spreadsheet"
    ) -> str:
        """Translate a single text using OpenAI API.

        Args:
            text: Text to translate
            source_lang: Source language code
            target_lang: Target language code
            context: Translation context

        Returns:
            Translated text
        """
        if not isinstance(text, str) or not text.strip():
            return text

        if not is_chinese(text):
            return text

        messages = [
            {
                "role": "system",
                "content": "You are a direct Chinese-to-English translator. ONLY translate the exact Chinese text provided. Do not add ANY formatting, explanations, or extra words. Keep punctuation, numbers, and special characters exactly as they appear in the original. Your job is ONLY literal translation."
            },
            {
                "role": "user",
                "content": f"Translate this text to English (provide ONLY the direct translation without ANY additional text): {text}"
            }
        ]

        try:
            # Use streaming for better performance
            stream = await asyncio.to_thread(
                lambda: self.client.chat.completions.create(
                    model=self.model,
                    messages=messages,
                    temperature=0.1,
                    stream=True
                )
            )

            collected_chunks = []
            for chunk in stream:
                collected_chunks.append(chunk)

            content_list = []
            for chunk in collected_chunks:
                if hasattr(chunk.choices[0].delta, 'content') and chunk.choices[0].delta.content is not None:
                    content_list.append(chunk.choices[0].delta.content)

            full_content = ''.join(content_list)
            return full_content.strip()

        except Exception as e:
            logger.error(f"Translation error for text '{text[:50]}...': {e}")
            return text

    async def translate_batch(
        self,
        texts: List[str],
        source_lang: str = "zh",
        target_lang: str = "en",
        context: str = "spreadsheet"
    ) -> Dict[str, str]:
        """Translate a batch of texts using OpenAI API.

        Args:
            texts: List of texts to translate
            source_lang: Source language code
            target_lang: Target language code
            context: Translation context

        Returns:
            Dict mapping original texts to translations
        """
        results = {}

        # Filter texts that need translation
        texts_to_translate = []
        for text in texts:
            if not isinstance(text, str) or not text.strip():
                results[text] = text
                continue

            if not is_chinese(text):
                results[text] = text
                continue

            texts_to_translate.append(text)

        if not texts_to_translate:
            return results

        # Process each text individually for better error handling
        for text in texts_to_translate:
            translation = await self.translate_single(text, source_lang, target_lang, context)
            results[text] = translation

        return results