import pandas as pd
import google.generativeai as genai
import time
import os
import glob
from typing import Optional, List
from concurrent.futures import ThreadPoolExecutor, as_completed
import threading
from pathlib import Path
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import queue
import json


class TranslationGUI:
    def __init__(self, root):
        self.root = root
        self.root.title("Excel Translator - Gemini API")
        self.root.geometry("800x700")
        self.root.resizable(True, True)

        # Variables
        self.api_key_var = tk.StringVar()
        self.prompt_file_var = tk.StringVar()
        self.input_file_var = tk.StringVar()
        self.input_folder_var = tk.StringVar()
        self.output_folder_var = tk.StringVar()
        self.delay_var = tk.DoubleVar(value=1.0)
        self.workers_var = tk.IntVar(value=3)
        self.mode_var = tk.StringVar(value="single")

        # Queue for thread communication
        self.log_queue = queue.Queue()

        # Load saved settings
        self.load_settings()

        self.create_widgets()
        self.update_log_display()

    def create_widgets(self):
        # Main container with scrolling
        main_frame = ttk.Frame(self.root, padding="10")
        main_frame.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Configure grid weights
        self.root.columnconfigure(0, weight=1)
        self.root.rowconfigure(0, weight=1)
        main_frame.columnconfigure(1, weight=1)

        current_row = 0

        # Title
        title_label = ttk.Label(main_frame, text="Excel Translator with Gemini API",
                                font=('Arial', 16, 'bold'))
        title_label.grid(row=current_row, column=0, columnspan=3, pady=(0, 20))
        current_row += 1

        # API Key Section
        api_frame = ttk.LabelFrame(main_frame, text="API Configuration", padding="10")
        api_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        api_frame.columnconfigure(1, weight=1)
        current_row += 1

        ttk.Label(api_frame, text="Gemini API Key:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        api_entry = ttk.Entry(api_frame, textvariable=self.api_key_var, show="*", width=50)
        api_entry.grid(row=0, column=1, sticky=(tk.W, tk.E), padx=(0, 10))

        ttk.Button(api_frame, text="Show/Hide",
                   command=self.toggle_api_visibility).grid(row=0, column=2)

        # Custom Prompt Section
        prompt_frame = ttk.LabelFrame(main_frame, text="Custom Prompt (Optional)", padding="10")
        prompt_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        prompt_frame.columnconfigure(1, weight=1)
        current_row += 1

        ttk.Label(prompt_frame, text="Prompt File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(prompt_frame, textvariable=self.prompt_file_var, width=50).grid(row=0, column=1, sticky=(tk.W, tk.E),
                                                                                  padx=(0, 10))
        ttk.Button(prompt_frame, text="Browse",
                   command=self.browse_prompt_file).grid(row=0, column=2, padx=(10, 0))

        ttk.Button(prompt_frame, text="Create Sample",
                   command=self.create_sample_prompt).grid(row=1, column=0, pady=(10, 0))
        ttk.Button(prompt_frame, text="Edit Prompt",
                   command=self.edit_prompt).grid(row=1, column=1, pady=(10, 0), sticky=tk.W, padx=(10, 0))

        # Processing Mode Section
        mode_frame = ttk.LabelFrame(main_frame, text="Processing Mode", padding="10")
        mode_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        current_row += 1

        ttk.Radiobutton(mode_frame, text="Single File", variable=self.mode_var,
                        value="single", command=self.on_mode_change).grid(row=0, column=0, sticky=tk.W)
        ttk.Radiobutton(mode_frame, text="Batch Process Folder", variable=self.mode_var,
                        value="batch", command=self.on_mode_change).grid(row=0, column=1, sticky=tk.W, padx=(20, 0))

        # Single File Section
        self.single_frame = ttk.LabelFrame(main_frame, text="Single File Translation", padding="10")
        self.single_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.single_frame.columnconfigure(1, weight=1)
        current_row += 1

        ttk.Label(self.single_frame, text="Input File:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(self.single_frame, textvariable=self.input_file_var, width=50).grid(row=0, column=1,
                                                                                      sticky=(tk.W, tk.E), padx=(0, 10))
        ttk.Button(self.single_frame, text="Browse",
                   command=self.browse_input_file).grid(row=0, column=2)

        # Batch Processing Section
        self.batch_frame = ttk.LabelFrame(main_frame, text="Batch Processing", padding="10")
        self.batch_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        self.batch_frame.columnconfigure(1, weight=1)
        current_row += 1

        ttk.Label(self.batch_frame, text="Input Folder:").grid(row=0, column=0, sticky=tk.W, padx=(0, 10))
        ttk.Entry(self.batch_frame, textvariable=self.input_folder_var, width=50).grid(row=0, column=1,
                                                                                       sticky=(tk.W, tk.E),
                                                                                       padx=(0, 10))
        ttk.Button(self.batch_frame, text="Browse",
                   command=self.browse_input_folder).grid(row=0, column=2)

        ttk.Label(self.batch_frame, text="Output Folder:").grid(row=1, column=0, sticky=tk.W, padx=(0, 10),
                                                                pady=(10, 0))
        ttk.Entry(self.batch_frame, textvariable=self.output_folder_var, width=50).grid(row=1, column=1,
                                                                                        sticky=(tk.W, tk.E),
                                                                                        padx=(0, 10), pady=(10, 0))
        ttk.Button(self.batch_frame, text="Browse",
                   command=self.browse_output_folder).grid(row=1, column=2, pady=(10, 0))

        # Settings Section
        settings_frame = ttk.LabelFrame(main_frame, text="Translation Settings", padding="10")
        settings_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 10))
        current_row += 1

        ttk.Label(settings_frame, text="Delay between API calls (seconds):").grid(row=0, column=0, sticky=tk.W,
                                                                                  padx=(0, 10))
        delay_spin = ttk.Spinbox(settings_frame, from_=0.5, to=5.0, increment=0.5,
                                 textvariable=self.delay_var, width=10)
        delay_spin.grid(row=0, column=1, sticky=tk.W)

        ttk.Label(settings_frame, text="Parallel Workers (batch):").grid(row=0, column=2, sticky=tk.W, padx=(20, 10))
        workers_spin = ttk.Spinbox(settings_frame, from_=1, to=10, increment=1,
                                   textvariable=self.workers_var, width=10)
        workers_spin.grid(row=0, column=3, sticky=tk.W)

        # Control Buttons
        button_frame = ttk.Frame(main_frame)
        button_frame.grid(row=current_row, column=0, columnspan=3, pady=20)
        current_row += 1

        self.translate_btn = ttk.Button(button_frame, text="Start Translation",
                                        command=self.start_translation, style="Accent.TButton")
        self.translate_btn.pack(side=tk.LEFT, padx=(0, 10))

        self.stop_btn = ttk.Button(button_frame, text="Stop",
                                   command=self.stop_translation, state=tk.DISABLED)
        self.stop_btn.pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(button_frame, text="Clear Log",
                   command=self.clear_log).pack(side=tk.LEFT, padx=(0, 10))

        ttk.Button(button_frame, text="Save Settings",
                   command=self.save_settings).pack(side=tk.LEFT)

        # Progress Bar
        self.progress_var = tk.StringVar(value="Ready")
        self.progress = ttk.Progressbar(main_frame, mode='indeterminate')
        self.progress.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E), pady=(0, 5))
        current_row += 1

        progress_label = ttk.Label(main_frame, textvariable=self.progress_var)
        progress_label.grid(row=current_row, column=0, columnspan=3)
        current_row += 1

        # Log Display
        log_frame = ttk.LabelFrame(main_frame, text="Translation Log", padding="10")
        log_frame.grid(row=current_row, column=0, columnspan=3, sticky=(tk.W, tk.E, tk.N, tk.S), pady=(10, 0))
        log_frame.columnconfigure(0, weight=1)
        log_frame.rowconfigure(0, weight=1)
        main_frame.rowconfigure(current_row, weight=1)

        self.log_text = scrolledtext.ScrolledText(log_frame, height=15, width=80, wrap=tk.WORD)
        self.log_text.grid(row=0, column=0, sticky=(tk.W, tk.E, tk.N, tk.S))

        # Initialize mode
        self.on_mode_change()

        # Translation control
        self.stop_translation_flag = False

    def log(self, message):
        """Add message to log queue for thread-safe logging"""
        self.log_queue.put(message)

    def update_log_display(self):
        """Update log display from queue (called periodically)"""
        try:
            while True:
                message = self.log_queue.get_nowait()
                self.log_text.insert(tk.END, f"{message}\n")
                self.log_text.see(tk.END)
                self.log_text.update()
        except queue.Empty:
            pass

        # Schedule next update
        self.root.after(100, self.update_log_display)

    def toggle_api_visibility(self):
        """Toggle API key visibility"""
        current_widget = None
        for widget in self.root.winfo_children():
            if isinstance(widget, ttk.Frame):
                for child in widget.winfo_children():
                    if isinstance(child, ttk.LabelFrame) and "API Configuration" in str(child.cget("text")):
                        for entry in child.winfo_children():
                            if isinstance(entry, ttk.Entry) and entry.cget("show") == "*":
                                entry.config(show="")
                                return
                            elif isinstance(entry, ttk.Entry) and entry.cget("show") == "":
                                entry.config(show="*")
                                return

    def browse_prompt_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Prompt File",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")]
        )
        if file_path:
            self.prompt_file_var.set(file_path)

    def browse_input_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel/CSV File",
            filetypes=[("Excel files", "*.xlsx"), ("Excel files", "*.xls"),
                       ("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if file_path:
            self.input_file_var.set(file_path)

    def browse_input_folder(self):
        folder_path = filedialog.askdirectory(title="Select Input Folder")
        if folder_path:
            self.input_folder_var.set(folder_path)

    def browse_output_folder(self):
        folder_path = filedialog.askdirectory(title="Select Output Folder")
        if folder_path:
            self.output_folder_var.set(folder_path)

    def on_mode_change(self):
        """Handle mode change between single file and batch processing"""
        if self.mode_var.get() == "single":
            self.single_frame.grid()
            self.batch_frame.grid_remove()
        else:
            self.single_frame.grid_remove()
            self.batch_frame.grid()

    def create_sample_prompt(self):
        """Create a sample prompt file"""
        sample_prompt = """Translate the following text from English to Arabic. 
Make sure the translation is:
- Accurate and contextually appropriate
- Suitable for e-commerce product descriptions
- Professional and clear
- Maintains technical terms appropriately

Text to translate: {text}

Provide only the Arabic translation, no additional text or explanations."""

        file_path = filedialog.asksaveasfilename(
            title="Save Sample Prompt File",
            defaultextension=".txt",
            filetypes=[("Text files", "*.txt"), ("All files", "*.*")],
            initialvalue="sample_prompt.txt"
        )

        if file_path:
            try:
                with open(file_path, "w", encoding="utf-8") as f:
                    f.write(sample_prompt)
                self.prompt_file_var.set(file_path)
                self.log(f"Sample prompt file created: {file_path}")
                messagebox.showinfo("Success", "Sample prompt file created successfully!")
            except Exception as e:
                messagebox.showerror("Error", f"Failed to create prompt file: {str(e)}")

    def edit_prompt(self):
        """Open prompt file for editing"""
        if not self.prompt_file_var.get():
            messagebox.showwarning("No File", "Please select a prompt file first")
            return

        try:
            # Open prompt editor window
            self.open_prompt_editor()
        except Exception as e:
            messagebox.showerror("Error", f"Failed to open prompt editor: {str(e)}")

    def open_prompt_editor(self):
        """Open a simple text editor for the prompt file"""
        editor_window = tk.Toplevel(self.root)
        editor_window.title("Prompt Editor")
        editor_window.geometry("600x400")

        # Load current prompt
        current_prompt = ""
        if os.path.exists(self.prompt_file_var.get()):
            try:
                with open(self.prompt_file_var.get(), 'r', encoding='utf-8') as f:
                    current_prompt = f.read()
            except Exception as e:
                current_prompt = f"Error loading file: {str(e)}"

        # Text editor
        text_editor = scrolledtext.ScrolledText(editor_window, wrap=tk.WORD)
        text_editor.pack(fill=tk.BOTH, expand=True, padx=10, pady=10)
        text_editor.insert('1.0', current_prompt)

        # Save button
        def save_prompt():
            try:
                content = text_editor.get('1.0', tk.END).strip()
                with open(self.prompt_file_var.get(), 'w', encoding='utf-8') as f:
                    f.write(content)
                messagebox.showinfo("Saved", "Prompt file saved successfully!")
                editor_window.destroy()
            except Exception as e:
                messagebox.showerror("Error", f"Failed to save: {str(e)}")

        button_frame = ttk.Frame(editor_window)
        button_frame.pack(fill=tk.X, padx=10, pady=(0, 10))

        ttk.Button(button_frame, text="Save", command=save_prompt).pack(side=tk.RIGHT, padx=(10, 0))
        ttk.Button(button_frame, text="Cancel", command=editor_window.destroy).pack(side=tk.RIGHT)

    def validate_inputs(self):
        """Validate user inputs before starting translation"""
        if not self.api_key_var.get().strip():
            messagebox.showerror("Error", "Please enter your Gemini API key")
            return False

        if self.mode_var.get() == "single":
            if not self.input_file_var.get() or not os.path.exists(self.input_file_var.get()):
                messagebox.showerror("Error", "Please select a valid input file")
                return False
        else:
            if not self.input_folder_var.get() or not os.path.exists(self.input_folder_var.get()):
                messagebox.showerror("Error", "Please select a valid input folder")
                return False

        return True

    def start_translation(self):
        """Start the translation process"""
        if not self.validate_inputs():
            return

        # Disable start button, enable stop button
        self.translate_btn.config(state=tk.DISABLED)
        self.stop_btn.config(state=tk.NORMAL)
        self.stop_translation_flag = False

        # Start progress bar
        self.progress.start()
        self.progress_var.set("Initializing translation...")

        # Save current settings
        self.save_settings()

        # Start translation in separate thread
        translation_thread = threading.Thread(target=self.run_translation, daemon=True)
        translation_thread.start()

    def stop_translation(self):
        """Stop the translation process"""
        self.stop_translation_flag = True
        self.log("Translation stop requested...")

    def run_translation(self):
        """Run translation in background thread"""
        try:
            # Create translator
            translator = ExcelTranslator(
                api_key=self.api_key_var.get(),
                prompt_file=self.prompt_file_var.get() if self.prompt_file_var.get() else None,
                log_callback=self.log,
                stop_flag_callback=lambda: self.stop_translation_flag
            )

            if self.mode_var.get() == "single":
                # Single file processing
                self.progress_var.set("Translating single file...")
                result = translator.process_single_file(
                    input_file_path=self.input_file_var.get(),
                    delay=self.delay_var.get()
                )

                if result['success']:
                    self.log(f"✅ Translation completed! {result['translations_made']} cells translated.")
                    self.log(f"Output saved to: {result['output_file']}")
                    messagebox.showinfo("Success",
                                        f"Translation completed!\n{result['translations_made']} cells translated.")
                else:
                    self.log(f"❌ Translation failed: {result['error']}")
                    messagebox.showerror("Error", f"Translation failed: {result['error']}")

            else:
                # Batch processing
                self.progress_var.set("Batch processing files...")
                output_folder = self.output_folder_var.get() if self.output_folder_var.get() else None

                results = translator.batch_process_folder(
                    folder_path=self.input_folder_var.get(),
                    output_folder=output_folder,
                    max_workers=self.workers_var.get(),
                    delay=self.delay_var.get()
                )

                # Show summary
                successful = sum(1 for r in results if r['success'])
                total_translations = sum(r['translations_made'] for r in results)

                summary = f"Batch processing completed!\n"
                summary += f"Files processed: {len(results)}\n"
                summary += f"Successful: {successful}\n"
                summary += f"Total translations: {total_translations}"

                if successful == len(results):
                    messagebox.showinfo("Success", summary)
                else:
                    messagebox.showwarning("Partial Success", summary)

        except Exception as e:
            self.log(f"❌ Fatal error: {str(e)}")
            messagebox.showerror("Error", f"Translation failed: {str(e)}")
        finally:
            # Re-enable controls
            self.root.after(0, self.translation_finished)

    def translation_finished(self):
        """Called when translation is finished"""
        self.translate_btn.config(state=tk.NORMAL)
        self.stop_btn.config(state=tk.DISABLED)
        self.progress.stop()
        self.progress_var.set("Ready")

    def clear_log(self):
        """Clear the log display"""
        self.log_text.delete('1.0', tk.END)

    def save_settings(self):
        """Save current settings to file"""
        settings = {
            'api_key': self.api_key_var.get(),
            'prompt_file': self.prompt_file_var.get(),
            'input_file': self.input_file_var.get(),
            'input_folder': self.input_folder_var.get(),
            'output_folder': self.output_folder_var.get(),
            'delay': self.delay_var.get(),
            'workers': self.workers_var.get(),
            'mode': self.mode_var.get()
        }

        try:
            with open('translator_settings.json', 'w') as f:
                json.dump(settings, f, indent=2)
            self.log("Settings saved")
        except Exception as e:
            self.log(f"Failed to save settings: {str(e)}")

    def load_settings(self):
        """Load settings from file"""
        try:
            if os.path.exists('translator_settings.json'):
                with open('translator_settings.json', 'r') as f:
                    settings = json.load(f)

                self.api_key_var.set(settings.get('api_key', ''))
                self.prompt_file_var.set(settings.get('prompt_file', ''))
                self.input_file_var.set(settings.get('input_file', ''))
                self.input_folder_var.set(settings.get('input_folder', ''))
                self.output_folder_var.set(settings.get('output_folder', ''))
                self.delay_var.set(settings.get('delay', 1.0))
                self.workers_var.set(settings.get('workers', 3))
                self.mode_var.set(settings.get('mode', 'single'))
        except Exception as e:
            pass  # Ignore errors loading settings


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
        # Set up logging and stop callback first
        self.log_callback = log_callback if log_callback else print
        self.should_stop = stop_flag_callback if stop_flag_callback else lambda: False

        # Initialize API and model
        genai.configure(api_key=api_key)
        self.model = genai.GenerativeModel('gemini-pro')
        self.lock = threading.Lock()

        # Load custom prompt after logging is set up
        self.custom_prompt = self._load_custom_prompt(prompt_file)

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


def main():
    """Launch the GUI application"""
    root = tk.Tk()

    # Set up the style
    style = ttk.Style()

    # Try to use a modern theme
    try:
        style.theme_use('clam')
    except:
        pass

    # Configure colors
    style.configure('Accent.TButton', foreground='white', background='#0078d4')

    app = TranslationGUI(root)

    # Handle window closing
    def on_closing():
        app.save_settings()
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_closing)

    # Center window on screen
    root.eval('tk::PlaceWindow . center')

    # Add menu bar
    menubar = tk.Menu(root)
    root.config(menu=menubar)

    # File menu
    file_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="File", menu=file_menu)
    file_menu.add_command(label="Create Sample Prompt", command=app.create_sample_prompt)
    file_menu.add_separator()
    file_menu.add_command(label="Save Settings", command=app.save_settings)
    file_menu.add_separator()
    file_menu.add_command(label="Exit", command=on_closing)

    # Help menu
    help_menu = tk.Menu(menubar, tearoff=0)
    menubar.add_cascade(label="Help", menu=help_menu)
    help_menu.add_command(label="About", command=lambda: messagebox.showinfo(
        "About",
        "Excel Translator with Gemini API\n\n"
        "This application translates specific cells in Excel/CSV files\n"
        "from English to Arabic using Google's Gemini API.\n\n"
        "Features:\n"
        "• Single file or batch processing\n"
        "• Custom translation prompts\n"
        "• Parallel processing for speed\n"
        "• Automatic settings saving"
    ))

    root.mainloop()


if __name__ == "__main__":
    main()