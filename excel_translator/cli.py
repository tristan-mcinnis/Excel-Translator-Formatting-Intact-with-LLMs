"""Command-line interface for Excel translator."""

import argparse
import os
import sys
from datetime import datetime
from typing import Optional

from .translation import ExcelTranslator
from .providers import OpenAIProvider
from .utils import setup_logging


def get_provider(provider_name: str, model: str) -> Optional[OpenAIProvider]:
    """Get translation provider instance.

    Args:
        provider_name: Name of the provider (currently only 'openai')
        model: Model name to use

    Returns:
        Provider instance or None if unsupported
    """
    if provider_name.lower() == "openai":
        api_key = os.getenv("OPENAI_API_KEY")
        if not api_key:
            print("Error: OPENAI_API_KEY environment variable not set")
            sys.exit(1)
        return OpenAIProvider(api_key=api_key, model=model)
    else:
        print(f"Error: Unsupported provider '{provider_name}'")
        print("Supported providers: openai")
        sys.exit(1)


def main():
    """Main CLI entry point."""
    parser = argparse.ArgumentParser(
        description="Excel Translator - Translate Excel files while preserving formatting",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
Examples:
  # Basic translation
  python -m excel_translator input.xlsx --output output.xlsx

  # With specific provider and model
  python -m excel_translator input.xlsx --output output.xlsx \\
    --provider openai --model gpt-5 --source-lang zh --target-lang en

  # Batch processing with custom settings
  python -m excel_translator input.xlsx --output output.xlsx \\
    --batch-size 10 --max-workers 4 --context "financial report"

Supported providers:
  - openai (models: gpt-4o, gpt-5)

Environment variables:
  OPENAI_API_KEY    Required for OpenAI provider
        """
    )

    # Required arguments
    parser.add_argument(
        "input_file",
        help="Input Excel file path"
    )

    # Optional arguments
    parser.add_argument(
        "--output", "-o",
        required=True,
        help="Output Excel file path"
    )

    parser.add_argument(
        "--provider",
        default="openai",
        choices=["openai"],
        help="Translation provider to use (default: openai)"
    )

    parser.add_argument(
        "--model",
        default="gpt-4o",
        help="Model to use for translation (default: gpt-4o, can use gpt-5 when available)"
    )

    parser.add_argument(
        "--source-lang",
        default="zh",
        help="Source language code (default: zh)"
    )

    parser.add_argument(
        "--target-lang",
        default="en",
        help="Target language code (default: en)"
    )

    parser.add_argument(
        "--context",
        default="spreadsheet",
        help="Translation context to help with accuracy (default: spreadsheet)"
    )

    parser.add_argument(
        "--batch-size",
        type=int,
        default=5,
        help="Number of cells to translate per batch (default: 5)"
    )

    parser.add_argument(
        "--max-retries",
        type=int,
        default=5,
        help="Maximum number of retries for failed translations (default: 5)"
    )

    parser.add_argument(
        "--save-interval",
        type=int,
        default=20,
        help="Save progress every N cells (default: 20)"
    )

    parser.add_argument(
        "--no-backup",
        action="store_true",
        help="Skip creating a backup file"
    )

    parser.add_argument(
        "--clear-cache",
        action="store_true",
        help="Clear the translation cache before starting"
    )

    parser.add_argument(
        "--cache-dir",
        default="translation_cache",
        help="Directory for translation cache (default: translation_cache)"
    )

    parser.add_argument(
        "--log-level",
        default="INFO",
        choices=["DEBUG", "INFO", "WARNING", "ERROR"],
        help="Logging level (default: INFO)"
    )

    parser.add_argument(
        "--log-file",
        help="Optional log file path (default: None - console only)"
    )

    args = parser.parse_args()

    # Set up logging
    if not args.log_file:
        log_filename = f"excel_translation_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"
        args.log_file = log_filename

    setup_logging(args.log_level, args.log_file)

    # Clean up file paths
    input_file = args.input_file.strip("\"'")
    output_file = args.output.strip("\"'")

    # Validate input file
    if not os.path.exists(input_file):
        print(f"Error: Input file '{input_file}' does not exist")
        sys.exit(1)

    # Create output directory if needed
    output_dir = os.path.dirname(output_file)
    if output_dir and not os.path.exists(output_dir):
        os.makedirs(output_dir)
        print(f"Created output directory: {output_dir}")

    # Get provider
    provider = get_provider(args.provider, args.model)

    # Initialize translator
    translator = ExcelTranslator(
        provider=provider,
        cache_dir=args.cache_dir,
        batch_size=args.batch_size,
        max_retries=args.max_retries,
        save_interval=args.save_interval
    )

    # Clear cache if requested
    if args.clear_cache:
        print("Clearing translation cache as requested")
        translator.cache = {}
        from .utils import save_cache
        save_cache({}, translator.cache_file)

    # Perform translation
    print(f"Translating {input_file} -> {output_file}")
    print(f"Provider: {args.provider}, Model: {args.model}")
    print(f"Languages: {args.source_lang} -> {args.target_lang}")

    try:
        translator.translate_file_sync(
            input_file=input_file,
            output_file=output_file,
            source_lang=args.source_lang,
            target_lang=args.target_lang,
            context=args.context,
            create_backup_file=not args.no_backup
        )
        print(f"Translation completed successfully!")
        print(f"Output saved to: {output_file}")
    except KeyboardInterrupt:
        print("\nTranslation interrupted by user")
        sys.exit(1)
    except Exception as e:
        print(f"Translation failed: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()