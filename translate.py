import argparse
import os
import sys
from translator import ExcelTranslator

def main():
    """
    Command-line interface for the Excel Translator.
    """
    parser = argparse.ArgumentParser(
        description="Translate specific columns in an Excel/CSV file using the Gemini API.",
        formatter_class=argparse.RawTextHelpFormatter
    )

    parser.add_argument(
        "input_file",
        help="Path to the input Excel or CSV file."
    )
    parser.add_argument(
        "--api-key",
        dest="api_key",
        help="Your Google Gemini API key. Recommended to use GEMINI_API_KEY environment variable instead.",
        default=os.environ.get("GEMINI_API_KEY")
    )
    parser.add_argument(
        "--prompt-file",
        dest="prompt_file",
        help="Path to a custom prompt text file. (Optional)"
    )
    parser.add_argument(
        "--delay",
        type=float,
        default=1.0,
        help="Delay in seconds between API calls. (Default: 1.0)"
    )
    parser.add_argument(
        "-o", "--output-file",
        dest="output_file",
        help="Path to save the translated file. (Optional)\nIf not provided, it will be saved as 'input_filename_translated.ext'."
    )

    args = parser.parse_args()

    if not args.api_key:
        print("Error: Gemini API key not found.")
        print("Please provide it using the --api-key argument or by setting the GEMINI_API_KEY environment variable.")
        sys.exit(1)

    if not os.path.exists(args.input_file):
        print(f"Error: Input file not found at '{args.input_file}'")
        sys.exit(1)

    print("--- Starting Translation ---")

    # Instantiate the translator
    try:
        translator = ExcelTranslator(
            api_key=args.api_key,
            prompt_file=args.prompt_file,
            log_callback=print  # Log messages directly to the console
        )
    except Exception as e:
        print(f"Error initializing translator: {e}")
        sys.exit(1)

    # Process the file
    result = translator.process_single_file(
        input_file_path=args.input_file,
        output_file_path=args.output_file,
        delay=args.delay
    )

    print("\n--- Translation Summary ---")
    if result['success']:
        print(f"✅ Success!")
        print(f"   - Translations made: {result['translations_made']}")
        print(f"   - Output file saved to: {result['output_file']}")
    else:
        print(f"❌ Failure!")
        print(f"   - Error: {result['error']}")
    print("-------------------------")


if __name__ == "__main__":
    main()
