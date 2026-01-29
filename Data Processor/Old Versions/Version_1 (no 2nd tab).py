"""
Data Processor - GUI Version
============================================
A user-friendly graphical interface for processing scholarship data.

INSTALLATION:
1. Install Python 3.7 or higher
2. Install required packages:
   pip install pandas openpyxl

USAGE:
Simply run this file:
   python .py file

No coding required - just click buttons!
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re
from pathlib import Path
import threading
from datetime import datetime


class ORGProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Processor App")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Variables
        self.mts_file = tk.StringVar()
        self.ccis_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.processing = False
        
        # Output file names
        self.output_names = {
            'PSLE': 'Organisation_MTSCTP PSLE',
            'NA': 'Organisation_MTSCTP SEC 4 NA',
            'NT': 'Organisation_MTSCTP SEC 4 NT',
            'EX': 'Organisation_MTSCTP SEC 4 EX'
        }
        
        self.setup_ui()
    
    def setup_ui(self):
        """Create the user interface."""
        # Header
        header = tk.Frame(self.root, bg="#2c3e50", height=80)
        header.pack(fill=tk.X)
        
        title = tk.Label(
            header,
            text="Organisation Data Processor",
            font=("Arial", 18, "bold"),
            bg="#2c3e50",
            fg="white"
        )
        title.pack(pady=20)
        
        # Main content
        main_frame = tk.Frame(self.root, padx=20, pady=20)
        main_frame.pack(fill=tk.BOTH, expand=True)
        
        # File Selection Section
        self.create_file_selection(main_frame)
        
        # Options Section
        self.create_options_section(main_frame)
        
        # Action Buttons
        self.create_action_buttons(main_frame)
        
        # Progress Section
        self.create_progress_section(main_frame)
        
        # Status bar
        self.status_bar = tk.Label(
            self.root,
            text="Ready to process files",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg="#ecf0f1"
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)
    
    def create_file_selection(self, parent):
        """Create file selection section."""
        file_frame = tk.LabelFrame(
            parent,
            text="1. Select Input Files",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=15
        )
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        # MTS File
        tk.Label(
            file_frame,
            text="MTS Students Attendance File (.csv):",
            font=("Arial", 10)
        ).grid(row=0, column=0, sticky=tk.W, pady=5)
        
        mts_entry = tk.Entry(file_frame, textvariable=self.mts_file, width=60)
        mts_entry.grid(row=0, column=1, padx=10, pady=5)
        
        tk.Button(
            file_frame,
            text="Browse...",
            command=self.browse_mts_file,
            width=12
        ).grid(row=0, column=2, pady=5)
        
        # CCIS File
        tk.Label(
            file_frame,
            text="CCIS/Dynamics 365 File (.xlsx):",
            font=("Arial", 10)
        ).grid(row=1, column=0, sticky=tk.W, pady=5)
        
        ccis_entry = tk.Entry(file_frame, textvariable=self.ccis_file, width=60)
        ccis_entry.grid(row=1, column=1, padx=10, pady=5)
        
        tk.Button(
            file_frame,
            text="Browse...",
            command=self.browse_ccis_file,
            width=12
        ).grid(row=1, column=2, pady=5)
        
        # Output Folder
        tk.Label(
            file_frame,
            text="Output Folder:",
            font=("Arial", 10)
        ).grid(row=2, column=0, sticky=tk.W, pady=5)
        
        output_entry = tk.Entry(file_frame, textvariable=self.output_folder, width=60)
        output_entry.grid(row=2, column=1, padx=10, pady=5)
        
        tk.Button(
            file_frame,
            text="Browse...",
            command=self.browse_output_folder,
            width=12
        ).grid(row=2, column=2, pady=5)
    
    def create_options_section(self, parent):
        """Create processing options section."""
        options_frame = tk.LabelFrame(
            parent,
            text="2. Processing Options",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=15
        )
        options_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.remove_invalid_nric = tk.BooleanVar(value=True)
        self.remove_duplicates = tk.BooleanVar(value=True)
        self.clean_names = tk.BooleanVar(value=True)
        self.generate_txt = tk.BooleanVar(value=True)
        
        tk.Checkbutton(
            options_frame,
            text="Remove invalid NRICs (not matching S1234567A format)",
            variable=self.remove_invalid_nric,
            font=("Arial", 10)
        ).grid(row=0, column=0, sticky=tk.W, pady=3)
        
        tk.Checkbutton(
            options_frame,
            text="Remove duplicate entries",
            variable=self.remove_duplicates,
            font=("Arial", 10)
        ).grid(row=1, column=0, sticky=tk.W, pady=3)
        
        tk.Checkbutton(
            options_frame,
            text="Clean names (remove numbers and special characters)",
            variable=self.clean_names,
            font=("Arial", 10)
        ).grid(row=2, column=0, sticky=tk.W, pady=3)
        
        tk.Checkbutton(
            options_frame,
            text="Generate .txt files (fixed-width format)",
            variable=self.generate_txt,
            font=("Arial", 10)
        ).grid(row=3, column=0, sticky=tk.W, pady=3)
    
    def create_action_buttons(self, parent):
        """Create action buttons."""
        button_frame = tk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.process_btn = tk.Button(
            button_frame,
            text="▶ Process Files",
            command=self.start_processing,
            font=("Arial", 12, "bold"),
            bg="#27ae60",
            fg="white",
            height=2,
            width=20
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="📁 Open Output Folder",
            command=self.open_output_folder,
            font=("Arial", 11),
            height=2,
            width=20
        ).pack(side=tk.LEFT, padx=5)
        
        tk.Button(
            button_frame,
            text="Clear Log",
            command=self.clear_log,
            font=("Arial", 11),
            height=2,
            width=15
        ).pack(side=tk.LEFT, padx=5)
    
    def create_progress_section(self, parent):
        """Create progress and log section."""
        progress_frame = tk.LabelFrame(
            parent,
            text="3. Processing Log",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=15
        )
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        # Log text area
        self.log_text = scrolledtext.ScrolledText(
            progress_frame,
            height=15,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#ecf0f1",
            insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    def browse_mts_file(self):
        """Browse for MTS file."""
        filename = filedialog.askopenfilename(
            title="Select MTS Students Attendance File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.mts_file.set(filename)
            self.log(f"Selected MTS file: {Path(filename).name}")
    
    def browse_ccis_file(self):
        """Browse for CCIS file."""
        filename = filedialog.askopenfilename(
            title="Select CCIS/Dynamics 365 File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.ccis_file.set(filename)
            self.log(f"Selected CCIS file: {Path(filename).name}")
    
    def browse_output_folder(self):
        """Browse for output folder."""
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
            self.log(f"Output folder: {folder}")
    
    def open_output_folder(self):
        """Open output folder in file explorer."""
        folder = self.output_folder.get()
        if folder and Path(folder).exists():
            os.startfile(folder) if os.name == 'nt' else os.system(f'open "{folder}"')
        else:
            messagebox.showwarning("Warning", "Output folder not set or doesn't exist")
    
    def log(self, message, level="INFO"):
        """Add message to log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {
            "INFO": "ℹ️",
            "SUCCESS": "✅",
            "WARNING": "⚠️",
            "ERROR": "❌"
        }.get(level, "•")
        
        self.log_text.insert(tk.END, f"[{timestamp}] {prefix} {message}\n")
        self.log_text.see(tk.END)
        self.root.update()
    
    def clear_log(self):
        """Clear the log."""
        self.log_text.delete(1.0, tk.END)
    
    def validate_inputs(self):
        """Validate all inputs before processing."""
        if not self.mts_file.get():
            messagebox.showerror("Error", "Please select the MTS file")
            return False
        
        if not self.ccis_file.get():
            messagebox.showerror("Error", "Please select the CCIS file")
            return False
        
        if not self.output_folder.get():
            messagebox.showerror("Error", "Please select an output folder")
            return False
        
        if not Path(self.mts_file.get()).exists():
            messagebox.showerror("Error", f"MTS file not found:\n{self.mts_file.get()}")
            return False
        
        if not Path(self.ccis_file.get()).exists():
            messagebox.showerror("Error", f"CCIS file not found:\n{self.ccis_file.get()}")
            return False
        
        return True
    
    def start_processing(self):
        """Start processing in a separate thread."""
        if not self.validate_inputs():
            return
        
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        
        # Run in separate thread to keep GUI responsive
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Main processing logic."""
        self.processing = True
        self.process_btn.config(state=tk.DISABLED, bg="#95a5a6")
        self.status_bar.config(text="Processing...", bg="#f39c12")
        
        try:
            self.log("="*70)
            self.log("Starting Data Processing", "INFO")
            self.log("="*70)
            
            # Create output folder
            output_path = Path(self.output_folder.get())
            output_path.mkdir(parents=True, exist_ok=True)
            
            # Load MTS data
            self.log("Loading MTS attendance file...")
            df_mts = pd.read_csv(self.mts_file.get())
            self.log(f"Loaded {len(df_mts)} records from MTS file", "SUCCESS")
            
            # Filter datasets
            datasets = {
                'PSLE': df_mts.query("Level == 'P6'"),
                'NA': df_mts.query("Level == 'S4' and Stream in ['G2', 'Normal Academic']"),
                'NT': df_mts.query("Level == 'S4' and Stream in ['G1', 'Normal Technical']"),
                'EX': df_mts.query("Level == 'S4' and Stream in ['G3', 'Express']")
            }
            
            # Load CCIS data
            self.log("Loading CCIS data...")
            df_ccis = pd.read_excel(self.ccis_file.get())
            self.log(f"Loaded {len(df_ccis)} records from CCIS file", "SUCCESS")
            
            # Process each category
            all_schools = set()
            total_processed = 0
            
            for category, df in datasets.items():
                self.log(f"\n{'='*70}")
                self.log(f"Processing: {self.output_names[category]}", "INFO")
                self.log(f"{'='*70}")
                
                result = self.process_category(
                    category, df, df_ccis, output_path, all_schools
                )
                total_processed += result['final_count']
            
            # Summary
            self.log(f"\n{'='*70}")
            self.log(f"ALL UNIQUE SCHOOLS ({len(all_schools)}):", "INFO")
            self.log(f"{'='*70}")
            for school in sorted(all_schools):
                self.log(f"  • {school}")
            
            self.log(f"\n{'='*70}")
            self.log(f"PROCESSING COMPLETE!", "SUCCESS")
            self.log(f"Total records processed: {total_processed}")
            self.log(f"Output folder: {output_path}")
            self.log(f"{'='*70}")
            
            self.status_bar.config(text="Processing completed successfully!", bg="#27ae60")
            messagebox.showinfo(
                "Success",
                f"Processing completed!\n\n"
                f"Total records: {total_processed}\n"
                f"Schools found: {len(all_schools)}\n\n"
                f"Files saved to:\n{output_path}"
            )
            
        except Exception as e:
            self.log(f"ERROR: {str(e)}", "ERROR")
            self.status_bar.config(text="Error occurred during processing", bg="#e74c3c")
            messagebox.showerror("Error", f"Processing failed:\n\n{str(e)}")
        
        finally:
            self.processing = False
            self.process_btn.config(state=tk.NORMAL, bg="#27ae60")
    
    def process_category(self, category, df, df_ccis, output_path, all_schools):
        """Process a single category."""
        # Merge datasets
        df_merged = pd.merge(
            df, df_ccis,
            left_on='Student Ref No.',
            right_on='Registration ID',
            how='inner'
        )
        
        # Select columns
        df_export = df_merged[[
            'NRIC (Main Applicant) (Contact)',
            'Current School/Institution (Main Applicant) (Contact)',
            'Main Applicant'
        ]].rename(columns={
            'NRIC (Main Applicant) (Contact)': 'NRIC',
            'Current School/Institution (Main Applicant) (Contact)': 'SCHOOL NAME',
            'Main Applicant': 'STATUTORY NAME'
        })
        
        initial_count = len(df_export)
        self.log(f"Initial records: {initial_count}")
        
        # Remove invalid NRICs
        invalid_count = 0
        if self.remove_invalid_nric.get():
            valid_nric = df_export['NRIC'].apply(self.validate_nric)
            invalid_count = (~valid_nric).sum()
            if invalid_count > 0:
                self.log(f"Removing {invalid_count} invalid NRICs", "WARNING")
            df_export = df_export[valid_nric]
        
        # Clean names
        if self.clean_names.get():
            df_export['STATUTORY NAME'] = df_export['STATUTORY NAME'].apply(
                self.clean_text
            ).str.upper()
            df_export['SCHOOL NAME'] = df_export['SCHOOL NAME'].str.upper()
            self.log("Names cleaned and standardized")
        
        # Remove duplicates
        duplicates = 0
        if self.remove_duplicates.get():
            before = len(df_export)
            df_export = df_export.drop_duplicates()
            duplicates = before - len(df_export)
            if duplicates > 0:
                self.log(f"Removed {duplicates} duplicates", "WARNING")
        
        # Collect schools
        schools = set(df_export['SCHOOL NAME'].unique())
        all_schools.update(schools)
        
        # Save Excel
        excel_path = output_path / f"{self.output_names[category]}.xlsx"
        df_export.to_excel(excel_path, index=False)
        self.log(f"Saved Excel: {excel_path.name}", "SUCCESS")
        
        # Generate text file
        if self.generate_txt.get():
            text_content, warnings = self.fixed_width_format(df_export)
            txt_path = output_path / f"{self.output_names[category]}.txt"
            txt_path.write_text(text_content, encoding='utf-8')
            self.log(f"Saved text file: {txt_path.name}", "SUCCESS")
            
            if warnings:
                self.log(f"Text format warnings: {len(warnings)}", "WARNING")
                for warning in warnings[:3]:
                    self.log(f"  {warning}", "WARNING")
                if len(warnings) > 3:
                    self.log(f"  ... and {len(warnings) - 3} more", "WARNING")
        
        # Summary
        final_count = len(df_export)
        self.log(f"Summary: {initial_count} → {final_count} records")
        self.log(f"  Invalid NRICs: {invalid_count}")
        self.log(f"  Duplicates: {duplicates}")
        self.log(f"  Schools: {len(schools)}")
        
        return {
            'final_count': final_count,
            'invalid_count': invalid_count,
            'duplicates': duplicates,
            'schools': len(schools)
        }
    
    @staticmethod
    def clean_text(text):
        """Remove numbers and special characters."""
        text = re.sub(r'\d+', '', str(text))
        text = re.sub(r"[^a-zA-Z ']", '', text)
        return text.strip()
    
    @staticmethod
    def validate_nric(nric):
        """Validate NRIC format."""
        pattern = r'^[STFGM]\d{7}[A-Z]$'
        return bool(re.match(pattern, str(nric).upper()))
    
    @staticmethod
    def fixed_width_format(df):
        """Convert to fixed-width format."""
        lines = []
        warnings = []
        
        for index, row in df.iterrows():
            nric = str(row['NRIC']).strip()
            school = str(row['SCHOOL NAME']).strip()
            name = str(row['STATUTORY NAME']).strip()
            
            if len(nric) > 9:
                warnings.append(f"Row {index + 1}: NRIC too long ({len(nric)} chars)")
            if len(school) > 66:
                warnings.append(f"Row {index + 1}: School name too long ({len(school)} chars)")
            if len(name) > 66:
                warnings.append(f"Row {index + 1}: Name too long ({len(name)} chars)")
            
            line = nric[:9].ljust(9) + school[:66].ljust(66) + name[:66].ljust(66)
            lines.append(line)
        
        return '\n'.join(lines), warnings


def main():
    """Main entry point."""
    root = tk.Tk()
    app = ORGProcessor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
