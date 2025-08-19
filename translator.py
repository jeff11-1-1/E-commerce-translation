import pandas as pd
import google.generativeai as genai
import time
import os
import glob
from typing import Optional, List
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading


class ExcelTranslator:
    def __init__(self, api_key: str, prompt_file: str = None, log_callback=None, stop_flag_callback=None):
        """
        Initialize the translator with Gemini API key and optional custom prompt

        Args:
            api_key (str): Your Google Gemini API key
            prompt_file (str): Path to text file containing custom translation prompt
            log_callback: Function to call for logging
            stop_flag_callback: Function to check if translation should stop
        """
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-pro')
        self.custom_prompt = self._load_custom_prompt(prompt_file)
        self.lock = threading.Lock()
        self.log_callback = log_callback if log_callback else print
        self.should_stop = stop_flag_callback if stop_flag_callback else lambda: False

    def log(self, message):
        """Log a message using the provided callback"""
        self.log_callback(message)

    def _load_custom_prompt(self, prompt_file: str) -> Optional[str]:
        """Load custom prompt from text file"""
        if not prompt_file or not os.path.exists(prompt_file):
            return None

        try:
            with open(prompt_file, 'r', encoding='utf-8') as f:
                prompt = f.read().strip()
                if prompt:
                    self.log(f"📄 Loaded custom prompt from: {os.path.basename(prompt_file)}")
                    return prompt
        except Exception as e:
            self.log(f"❌ Error loading prompt file: {e}")

        return None

    def translate_text(self, text: str, delay: float = 1.0) -> Optional[str]:
        """Translate text from English to Arabic using Gemini API"""
        if self.should_stop():
            return None

        try:
            # Use custom prompt if available, otherwise use default
            if self.custom_prompt:
                prompt = self.custom_prompt.replace("{text}", text)
            else:
                prompt = f"Translate the following text from English to Arabic. Only provide the translation, no additional text:\n\n{text}"

            response = self.model.generate_content(prompt)

            # Add delay to respect rate limits
            time.sleep(delay)

            return response.text.strip()

        except Exception as e:
            self.log(f"❌ Error translating '{text[:30]}...': {str(e)}")
            return None

    def process_single_file(self, input_file_path: str, output_file_path: str = None, delay: float = 1.0) -> dict:
        """Process a single Excel file and translate specified cells"""
        result = {
            'file': input_file_path,
            'success': False,
            'translations_made': 0,
            'total_rows': 0,
            'error': None,
            'output_file': None
        }

        try:
            # Read the file
            if input_file_path.endswith('.csv'):
                df = pd.read_csv(input_file_path)
            else:
                df = pd.read_excel(input_file_path)

            result['total_rows'] = len(df)
            self.log(f"📂 Processing: {os.path.basename(input_file_path)} ({len(df)} rows)")

            # Column indices
            english_col_idx = 2  # Third column
            arabic_col_idx = 3  # Fourth column
            check_col_idx = 4  # Fifth column

            if len(df.columns) < 5:
                result['error'] = "File must have at least 5 columns"
                return result

            # Find rows to translate
            rows_to_translate = df[df.iloc[:, check_col_idx] == 1]
            self.log(f"🔍 Found {len(rows_to_translate)} rows marked for translation")

            # Translate each qualifying row
            translations_made = 0
            for idx, row in rows_to_translate.iterrows():
                if self.should_stop():
                    self.log("⏹️ Translation stopped by user")
                    break

                english_text = str(row.iloc[english_col_idx])

                # Skip empty values
                if pd.isna(english_text) or english_text.strip() == '' or english_text.lower() == 'nan':
                    continue

                self.log(f"🔄 Translating row {idx}: '{english_text[:50]}...'")

                arabic_translation = self.translate_text(english_text, delay)

                if arabic_translation:
                    df.at[idx, df.columns[arabic_col_idx]] = arabic_translation
                    translations_made += 1
                    self.log(f"✅ Row {idx} translated successfully")
                else:
                    self.log(f"❌ Failed to translate row {idx}")

            # Save the updated file
            if output_file_path is None:
                name, ext = os.path.splitext(input_file_path)
                output_file_path = f"{name}_translated{ext}"

            if output_file_path.endswith('.csv'):
                df.to_csv(output_file_path, index=False)
            else:
                df.to_excel(output_file_path, index=False)

            result['success'] = True
            result['translations_made'] = translations_made
            result['output_file'] = output_file_path

            self.log(f"💾 Saved: {os.path.basename(output_file_path)} ({translations_made} translations)")

        except Exception as e:
            result['error'] = str(e)
            self.log(f"❌ Error processing {os.path.basename(input_file_path)}: {str(e)}")

        return result

    def batch_process_folder(self, folder_path: str, output_folder: str = None,
                             max_workers: int = 3, delay: float = 1.0,
                             file_extensions: List[str] = None) -> List[dict]:
        """Process all Excel/CSV files in a folder with parallel processing"""
        if file_extensions is None:
            file_extensions = ['*.xlsx', '*.xls', '*.csv']

        # Find all matching files
        all_files = []
        for ext in file_extensions:
            pattern = os.path.join(folder_path, ext)
            all_files.extend(glob.glob(pattern))

        if not all_files:
            self.log(f"❌ No files found in {folder_path} with extensions {file_extensions}")
            return []

        self.log(f"📁 Found {len(all_files)} files to process")
        self.log(f"⚡ Using {max_workers} parallel workers")

        # Set up output folder
        if output_folder is None:
            output_folder = folder_path
        os.makedirs(output_folder, exist_ok=True)

        # Process files in parallel
        results = []
        with ThreadPoolExecutor(max_workers=max_workers) as executor:
            # Submit all tasks
            future_to_file = {}
            for file_path in all_files:
                if self.should_stop():
                    break

                # Create output path
                filename = os.path.basename(file_path)
                name, ext = os.path.splitext(filename)
                output_path = os.path.join(output_folder, f"{name}_translated{ext}")

                # Submit task
                future = executor.submit(
                    self.process_single_file,
                    file_path,
                    output_path,
                    delay
                )
                future_to_file[future] = file_path

            # Collect results as they complete
            for future in as_completed(future_to_file):
                if self.should_stop():
                    break

                file_path = future_to_file[future]
                try:
                    result = future.result()
                    results.append(result)
                except Exception as e:
                    results.append({
                        'file': file_path,
                        'success': False,
                        'error': f"Processing failed: {str(e)}",
                        'translations_made': 0,
                        'total_rows': 0,
                        'output_file': None
                    })

        # Print summary
        self._print_batch_summary(results)
        return results

    def _print_batch_summary(self, results: List[dict]):
        """Print summary of batch processing results"""
        total_files = len(results)
        successful_files = sum(1 for r in results if r['success'])
        total_translations = sum(r['translations_made'] for r in results)

        self.log("\n" + "=" * 60)
        self.log("📊 BATCH PROCESSING SUMMARY")
        self.log("=" * 60)
        self.log(f"Total files processed: {total_files}")
        self.log(f"Successful files: {successful_files}")
        self.log(f"Failed files: {total_files - successful_files}")
        self.log(f"Total translations made: {total_translations}")

        if successful_files < total_files:
            self.log("\n❌ Failed files:")
            for result in results:
                if not result['success']:
                    self.log(f"  - {os.path.basename(result['file'])}: {result['error']}")

        if successful_files > 0:
            self.log("\n✅ Successful files:")
            for result in results:
                if result['success']:
                    self.log(f"  - {os.path.basename(result['file'])}: {result['translations_made']} translations")
