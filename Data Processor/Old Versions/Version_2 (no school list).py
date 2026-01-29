# Add this section after the imports at the top
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re
from pathlib import Path
import threading
from datetime import datetime

# ============= NEW: ADD LOGGING MODULE =============
class ProcessingLogger:
    """Handles both GUI and file logging."""
    
    def __init__(self, output_folder):
        self.output_folder = Path(output_folder)
        self.log_file = self.output_folder / f"Processing_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        self.removed_file = self.output_folder / f"Removed_Records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        self.messages = []
        self.removed_records = []
        
    def log(self, message, level="INFO"):
        """Store message for both GUI and file."""
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        formatted = f"[{timestamp}] [{level}] {message}"
        self.messages.append(formatted)
        return formatted
    
    def add_removed_record(self, category, reason, nric, school, name):
        """Track removed records."""
        record = {
            'Category': category,
            'Reason': reason,
            'NRIC': nric,
            'School': school,
            'Name': name
        }
        self.removed_records.append(record)
    
    def save_logs(self):
        """Save all logs and removed records to files."""
        # Save processing log
        with open(self.log_file, 'w', encoding='utf-8') as f:
            f.write("ORG DATA PROCESSOR - PROCESSING LOG\n")
            f.write("=" * 80 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 80 + "\n\n")
            f.write('\n'.join(self.messages))
        
        print(f"✅ Log saved: {self.log_file}")
        
        # Save removed records
        if self.removed_records:
            df_removed = pd.DataFrame(self.removed_records)
            excel_removed = self.output_folder / f"Removed_Records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            df_removed.to_excel(excel_removed, index=False)
            
            # Also save as TXT
            with open(self.removed_file, 'w', encoding='utf-8') as f:
                f.write("REMOVED RECORDS - REFERENCE DOCUMENT\n")
                f.write("=" * 100 + "\n")
                f.write(f"Total removed: {len(self.removed_records)}\n")
                f.write("=" * 100 + "\n\n")
                
                # Group by reason
                for reason in set(r['Reason'] for r in self.removed_records):
                    records_of_reason = [r for r in self.removed_records if r['Reason'] == reason]
                    f.write(f"\n{reason} ({len(records_of_reason)} records):\n")
                    f.write("-" * 100 + "\n")
                    f.write(f"{'Category':<10} {'NRIC':<12} {'School':<40} {'Name':<35}\n")
                    f.write("-" * 100 + "\n")
                    
                    for rec in records_of_reason:
                        f.write(f"{rec['Category']:<10} {str(rec['NRIC'])[:11]:<12} {str(rec['School'])[:39]:<40} {str(rec['Name'])[:34]:<35}\n")
            
            print(f"✅ Removed records saved: {self.removed_file}")
            print(f"✅ Removed records Excel: {excel_removed}")


class ORGProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("ORG Data Processor")
        self.root.geometry("900x700")
        self.root.resizable(True, True)
        
        # Variables
        self.mts_file = tk.StringVar()
        self.ccis_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.processing = False
        self.logger = None  # Will be initialized during processing
        
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
        
        self.create_file_selection(main_frame)
        self.create_options_section(main_frame)
        self.create_action_buttons(main_frame)
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
        
        tk.Label(file_frame, text="MTS Students Attendance File (.csv):", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        mts_entry = tk.Entry(file_frame, textvariable=self.mts_file, width=60)
        mts_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse...", command=self.browse_mts_file, width=12).grid(row=0, column=2, pady=5)
        
        tk.Label(file_frame, text="CCIS/Dynamics 365 File (.xlsx):", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=5)
        ccis_entry = tk.Entry(file_frame, textvariable=self.ccis_file, width=60)
        ccis_entry.grid(row=1, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse...", command=self.browse_ccis_file, width=12).grid(row=1, column=2, pady=5)
        
        tk.Label(file_frame, text="Output Folder:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=5)
        output_entry = tk.Entry(file_frame, textvariable=self.output_folder, width=60)
        output_entry.grid(row=2, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse...", command=self.browse_output_folder, width=12).grid(row=2, column=2, pady=5)
    
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
        
        tk.Checkbutton(options_frame, text="Remove invalid NRICs (not matching S1234567A format)", variable=self.remove_invalid_nric, font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=3)
        tk.Checkbutton(options_frame, text="Remove duplicate entries", variable=self.remove_duplicates, font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=3)
        tk.Checkbutton(options_frame, text="Clean names (remove numbers and special characters)", variable=self.clean_names, font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=3)
        tk.Checkbutton(options_frame, text="Generate .txt files (fixed-width format)", variable=self.generate_txt, font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=3)
    
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
        
        tk.Button(button_frame, text="📁 Open Output Folder", command=self.open_output_folder, font=("Arial", 11), height=2, width=20).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="📋 View Removed Records", command=self.view_removed_records, font=("Arial", 11), height=2, width=20).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Clear Log", command=self.clear_log, font=("Arial", 11), height=2, width=15).pack(side=tk.LEFT, padx=5)
    
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
        filename = filedialog.askopenfilename(
            title="Select MTS Students Attendance File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.mts_file.set(filename)
            self.log_gui(f"Selected MTS file: {Path(filename).name}")
    
    def browse_ccis_file(self):
        filename = filedialog.askopenfilename(
            title="Select CCIS/Dynamics 365 File",
            filetypes=[("Excel files", "*.xlsx *.xls"), ("All files", "*.*")]
        )
        if filename:
            self.ccis_file.set(filename)
            self.log_gui(f"Selected CCIS file: {Path(filename).name}")
    
    def browse_output_folder(self):
        folder = filedialog.askdirectory(title="Select Output Folder")
        if folder:
            self.output_folder.set(folder)
            self.log_gui(f"Output folder: {folder}")
    
    def open_output_folder(self):
        folder = self.output_folder.get()
        if folder and Path(folder).exists():
            if os.name == 'nt':
                os.startfile(folder)
            else:
                os.system(f'open "{folder}"' if os.uname().sysname == 'Darwin' else f'xdg-open "{folder}"')
        else:
            messagebox.showwarning("Warning", "Output folder not set or doesn't exist")
    
    def view_removed_records(self):
        """Open removed records file if it exists."""
        if not self.logger:
            messagebox.showinfo("Info", "No processing has been done yet.")
            return
        
        removed_file = self.logger.removed_file
        if removed_file.exists():
            if os.name == 'nt':
                os.startfile(removed_file)
            else:
                os.system(f'open "{removed_file}"' if os.uname().sysname == 'Darwin' else f'xdg-open "{removed_file}"')
            messagebox.showinfo("Success", f"Removed records file opened:\n{removed_file}")
        else:
            messagebox.showinfo("Info", "No records were removed in the last processing.")
    
    def log_gui(self, message, level="INFO"):
        """Display message in GUI log."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {
            "INFO": "ℹ️",
            "SUCCESS": "✅",
            "WARNING": "⚠️",
            "ERROR": "❌"
        }.get(level, "•")
        
        gui_message = f"[{timestamp}] {prefix} {message}"
        self.log_text.insert(tk.END, gui_message + "\n")
        self.log_text.see(tk.END)
        self.root.update()
        
        # Also log to file if logger exists
        if self.logger:
            self.logger.log(message, level)
    
    def clear_log(self):
        self.log_text.delete(1.0, tk.END)
    
    def validate_inputs(self):
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
        if not self.validate_inputs():
            return
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        
        thread = threading.Thread(target=self.process_files)
        thread.daemon = True
        thread.start()
    
    def process_files(self):
        """Main processing logic with logging."""
        self.processing = True
        self.process_btn.config(state=tk.DISABLED, bg="#95a5a6")
        self.status_bar.config(text="Processing...", bg="#f39c12")
        
        # Initialize logger
        output_path = Path(self.output_folder.get())
        output_path.mkdir(parents=True, exist_ok=True)
        self.logger = ProcessingLogger(output_path)
        
        try:
            self.log_gui("="*70)
            self.log_gui("Starting Data Processing", "INFO")
            self.log_gui("="*70)
            
            self.log_gui("Loading MTS attendance file...")
            df_mts = pd.read_csv(self.mts_file.get())
            self.log_gui(f"Loaded {len(df_mts)} records from MTS file", "SUCCESS")
            
            datasets = {
                'PSLE': df_mts.query("Level == 'P6'"),
                'NA': df_mts.query("Level == 'S4' and Stream in ['G2', 'Normal Academic']"),
                'NT': df_mts.query("Level == 'S4' and Stream in ['G1', 'Normal Technical']"),
                'EX': df_mts.query("Level == 'S4' and Stream in ['G3', 'Express']")
            }
            
            self.log_gui("Loading CCIS data...")
            df_ccis = pd.read_excel(self.ccis_file.get())
            self.log_gui(f"Loaded {len(df_ccis)} records from CCIS file", "SUCCESS")
            
            all_schools = set()
            total_processed = 0
            
            for category, df in datasets.items():
                self.log_gui(f"\n{'='*70}")
                self.log_gui(f"Processing: {self.output_names[category]}", "INFO")
                self.log_gui(f"{'='*70}")
                
                result = self.process_category(category, df, df_ccis, output_path, all_schools)
                total_processed += result['final_count']
            
            self.log_gui(f"\n{'='*70}")
            self.log_gui(f"ALL UNIQUE SCHOOLS ({len(all_schools)}):", "INFO")
            self.log_gui(f"{'='*70}")
            for school in sorted(all_schools):
                self.log_gui(f"  • {school}")
            
            self.log_gui(f"\n{'='*70}")
            self.log_gui(f"PROCESSING COMPLETE!", "SUCCESS")
            self.log_gui(f"Total records processed: {total_processed}")
            self.log_gui(f"Total removed records: {len(self.logger.removed_records)}")
            self.log_gui(f"Output folder: {output_path}")
            self.log_gui(f"{'='*70}")
            
            # Save logs to file
            self.logger.save_logs()
            
            self.status_bar.config(text="Processing completed successfully!", bg="#27ae60")
            messagebox.showinfo(
                "Success",
                f"Processing completed!\n\n"
                f"Total records: {total_processed}\n"
                f"Removed records: {len(self.logger.removed_records)}\n"
                f"Schools found: {len(all_schools)}\n\n"
                f"Log files saved to:\n{output_path}"
            )
            
        except Exception as e:
            self.log_gui(f"ERROR: {str(e)}", "ERROR")
            self.status_bar.config(text="Error occurred during processing", bg="#e74c3c")
            messagebox.showerror("Error", f"Processing failed:\n\n{str(e)}")
        
        finally:
            self.processing = False
            self.process_btn.config(state=tk.NORMAL, bg="#27ae60")
    
    def process_category(self, category, df, df_ccis, output_path, all_schools):
        """Process a single category with removed record tracking."""
        df_merged = pd.merge(
            df, df_ccis,
            left_on='Student Ref No.',
            right_on='Registration ID',
            how='inner'
        )
        
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
        self.log_gui(f"Initial records: {initial_count}")
        
        # Remove invalid NRICs
        invalid_count = 0
        if self.remove_invalid_nric.get():
            valid_nric = df_export['NRIC'].apply(self.validate_nric)
            invalid_rows = df_export[~valid_nric]
            invalid_count = (~valid_nric).sum()
            
            # Track removed records
            for idx, row in invalid_rows.iterrows():
                self.logger.add_removed_record(
                    category=category,
                    reason="Invalid NRIC Format",
                    nric=str(row['NRIC']),
                    school=str(row['SCHOOL NAME']),
                    name=str(row['STATUTORY NAME'])
                )
            
            if invalid_count > 0:
                self.log_gui(f"Removing {invalid_count} invalid NRICs", "WARNING")
            df_export = df_export[valid_nric]
        
        # Clean names
        if self.clean_names.get():
            df_export['STATUTORY NAME'] = df_export['STATUTORY NAME'].apply(self.clean_text).str.upper()
            df_export['SCHOOL NAME'] = df_export['SCHOOL NAME'].str.upper()
            self.log_gui("Names cleaned and standardized")
        
        # Remove duplicates
        duplicates = 0
        if self.remove_duplicates.get():
            before = len(df_export)
            duplicate_rows = df_export[df_export.duplicated(keep=False)]
            
            # Track duplicate records (keep first, mark rest as removed)
            df_export_dedup = df_export.drop_duplicates()
            removed_dupes = df_export[df_export.duplicated(keep='first')]
            
            for idx, row in removed_dupes.iterrows():
                self.logger.add_removed_record(
                    category=category,
                    reason="Duplicate Entry",
                    nric=str(row['NRIC']),
                    school=str(row['SCHOOL NAME']),
                    name=str(row['STATUTORY NAME'])
                )
            
            duplicates = before - len(df_export_dedup)
            if duplicates > 0:
                self.log_gui(f"Removed {duplicates} duplicates", "WARNING")
            df_export = df_export_dedup
        
        # Collect schools
        schools = set(df_export['SCHOOL NAME'].unique())
        all_schools.update(schools)
        
        # Save Excel
        excel_path = output_path / f"{self.output_names[category]}.xlsx"
        df_export.to_excel(excel_path, index=False)
        self.log_gui(f"Saved Excel: {excel_path.name}", "SUCCESS")
        
        # Generate text file
        if self.generate_txt.get():
            text_content, warnings = self.fixed_width_format(df_export)
            txt_path = output_path / f"{self.output_names[category]}.txt"
            txt_path.write_text(text_content, encoding='utf-8')
            self.log_gui(f"Saved text file: {txt_path.name}", "SUCCESS")
        
        final_count = len(df_export)
        self.log_gui(f"Summary: {initial_count} → {final_count} records")
        self.log_gui(f"  Invalid NRICs: {invalid_count}")
        self.log_gui(f"  Duplicates: {duplicates}")
        self.log_gui(f"  Schools: {len(schools)}")
        
        return {
            'final_count': final_count,
            'invalid_count': invalid_count,
            'duplicates': duplicates,
            'schools': len(schools)
        }
    
    @staticmethod
    def clean_text(text):
        text = re.sub(r'\d+', '', str(text))
        text = re.sub(r"[^a-zA-Z ']", '', text)
        return text.strip()
    
    @staticmethod
    def validate_nric(nric):
        pattern = r'^[STFGM]\d{7}[A-Z]$'
        return bool(re.match(pattern, str(nric).upper()))
    
    @staticmethod
    def fixed_width_format(df):
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
    root = tk.Tk()
    app = ORGProcessor(root)
    root.mainloop()


if __name__ == "__main__":
    main()
