# Organisation Data Processor App working GUI - Version 3
import tkinter as tk
from tkinter import ttk, filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re
from pathlib import Path
import threading
from datetime import datetime


# ============= PROCESSING LOGGER CLASS =============
class ProcessingLogger:
    """Handles both GUI and file logging."""
    
    def __init__(self, output_folder):
        self.output_folder = Path(output_folder)
        self.log_file = self.output_folder / f"Processing_Log_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        self.removed_file = self.output_folder / f"Removed_Records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        self.removed_excel = self.output_folder / f"Removed_Records_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
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
            f.write("DATA PROCESSOR - PROCESSING LOG\n")
            f.write("=" * 100 + "\n")
            f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
            f.write("=" * 100 + "\n\n")
            f.write('\n'.join(self.messages))
        
        # Save removed records as Excel
        if self.removed_records:
            df_removed = pd.DataFrame(self.removed_records)
            df_removed.to_excel(self.removed_excel, index=False)
            
            # Also save as TXT
            with open(self.removed_file, 'w', encoding='utf-8') as f:
                f.write("REMOVED RECORDS - REFERENCE DOCUMENT\n")
                f.write("=" * 120 + "\n")
                f.write(f"Total removed: {len(self.removed_records)}\n")
                f.write("=" * 120 + "\n\n")
                
                # Group by reason
                for reason in set(r['Reason'] for r in self.removed_records):
                    records_of_reason = [r for r in self.removed_records if r['Reason'] == reason]
                    f.write(f"\n{reason} ({len(records_of_reason)} records):\n")
                    f.write("-" * 120 + "\n")
                    f.write(f"{'Category':<10} {'NRIC':<12} {'School':<45} {'Name':<40}\n")
                    f.write("-" * 120 + "\n")
                    
                    for rec in records_of_reason:
                        f.write(f"{str(rec['Category']):<10} {str(rec['NRIC'])[:11]:<12} {str(rec['School'])[:44]:<45} {str(rec['Name'])[:39]:<40}\n")


class ORGProcessor:
    def __init__(self, root):
        self.root = root
        self.root.title("Data Processor App")
        self.root.geometry("1000x800")
        self.root.resizable(True, True)
        
        # Variables
        self.mts_file = tk.StringVar()
        self.ccis_file = tk.StringVar()
        self.output_folder = tk.StringVar()
        self.processing = False
        self.logger = None
        self.mode = "process"
        
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
        
        # Mode selection tabs
        self.create_mode_tabs(main_frame)
        
        # Content frames
        self.process_frame = tk.Frame(main_frame)
        self.process_frame.pack(fill=tk.BOTH, expand=True)
        
        self.regenerate_frame = tk.Frame(main_frame)
        
        # Setup process mode content
        self.create_file_selection(self.process_frame)
        self.create_options_section(self.process_frame)
        self.create_action_buttons_process(self.process_frame)
        self.create_progress_section(self.process_frame)
        
        # Setup regenerate mode content
        self.create_regenerate_section(self.regenerate_frame)
        
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
    
    def create_mode_tabs(self, parent):
        """Create mode selection tabs."""
        tab_frame = tk.Frame(parent, bg="#ecf0f1")
        tab_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.process_tab_btn = tk.Button(
            tab_frame,
            text="1️⃣ Process New Data",
            command=lambda: self.switch_mode("process"),
            font=("Arial", 11, "bold"),
            bg="#27ae60",
            fg="white",
            height=2,
            width=25,
            relief=tk.RAISED,
            bd=3
        )
        self.process_tab_btn.pack(side=tk.LEFT, padx=10, pady=10)
        
        self.regenerate_tab_btn = tk.Button(
            tab_frame,
            text="2️⃣ Regenerate TXT Files",
            command=lambda: self.switch_mode("regenerate"),
            font=("Arial", 11, "bold"),
            bg="#95a5a6",
            fg="white",
            height=2,
            width=25,
            relief=tk.SUNKEN,
            bd=3
        )
        self.regenerate_tab_btn.pack(side=tk.LEFT, padx=10, pady=10)
        
        info_label = tk.Label(
            tab_frame,
            text="Tip: Use 'Process New Data' first, then 'Regenerate TXT Files' after manual edits",
            font=("Arial", 9, "italic"),
            bg="#ecf0f1",
            fg="#34495e"
        )
        info_label.pack(side=tk.LEFT, padx=10)
    
    def switch_mode(self, new_mode):
        """Switch between process and regenerate modes."""
        self.mode = new_mode
        
        if new_mode == "process":
            self.regenerate_frame.pack_forget()
            self.process_frame.pack(fill=tk.BOTH, expand=True)
            
            self.process_tab_btn.config(bg="#27ae60", relief=tk.RAISED, bd=3)
            self.regenerate_tab_btn.config(bg="#95a5a6", relief=tk.SUNKEN, bd=3)
            
            self.log_gui("Switched to: Process New Data mode", "INFO")
        else:
            self.process_frame.pack_forget()
            self.regenerate_frame.pack(fill=tk.BOTH, expand=True)
            
            self.process_tab_btn.config(bg="#95a5a6", relief=tk.SUNKEN, bd=3)
            self.regenerate_tab_btn.config(bg="#27ae60", relief=tk.RAISED, bd=3)
            
            self.regen_log("Switched to: Regenerate TXT Files mode", "INFO")
    
    # ===== PROCESS MODE SECTIONS =====
    
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
    
    def create_action_buttons_process(self, parent):
        """Create action buttons for process mode."""
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
            width=18
        )
        self.process_btn.pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="📁 Open Output Folder", command=self.open_output_folder, font=("Arial", 11), height=2, width=18).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="📋 Removed Records", command=self.view_removed_records, font=("Arial", 11), height=2, width=18).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="🏫 School List", command=self.view_school_list, font=("Arial", 11), height=2, width=18).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Clear Log", command=self.clear_log, font=("Arial", 11), height=2, width=12).pack(side=tk.LEFT, padx=5)
    
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
            height=12,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#ecf0f1",
            insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)
    
    # ===== REGENERATE MODE SECTIONS =====
    
    def create_regenerate_section(self, parent):
        """Create regenerate TXT files section."""
        # Instructions
        instr_frame = tk.LabelFrame(
            parent,
            text="ℹ️ How to Use This Mode",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=15,
            bg="#e8f4f8"
        )
        instr_frame.pack(fill=tk.X, pady=(0, 15))
        
        instructions = """
1. After initial processing, you'll have 4 Excel files (PSLE, NA, NT, EX)
2. Open and edit these files to fix school names or student names
3. Save the edited Excel files
4. Come back here and select the folder containing your edited files
5. Click "Regenerate TXT Files" to create new TXT files from your edits
6. Your cleaned TXT files will be ready for submission!
        """
        
        instr_label = tk.Label(instr_frame, text=instructions, font=("Arial", 10), justify=tk.LEFT, bg="#e8f4f8")
        instr_label.pack(anchor=tk.W)
        
        # File selection
        file_frame = tk.LabelFrame(
            parent,
            text="1. Select Folder with Edited Excel Files",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=15
        )
        file_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.regen_folder = tk.StringVar()
        
        tk.Label(file_frame, text="Folder containing edited Excel files:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        folder_entry = tk.Entry(file_frame, textvariable=self.regen_folder, width=70)
        folder_entry.grid(row=0, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse...", command=self.browse_regen_folder, width=12).grid(row=0, column=2, pady=5)
        
        # Action buttons
        button_frame = tk.Frame(parent)
        button_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.regen_btn = tk.Button(
            button_frame,
            text="▶ Regenerate TXT Files",
            command=self.start_regeneration,
            font=("Arial", 12, "bold"),
            bg="#2980b9",
            fg="white",
            height=2,
            width=25
        )
        self.regen_btn.pack(side=tk.LEFT, padx=5)
        
        tk.Button(button_frame, text="📁 Open Folder", command=self.open_regen_folder, font=("Arial", 11), height=2, width=25).pack(side=tk.LEFT, padx=5)
        tk.Button(button_frame, text="Clear Log", command=self.clear_log, font=("Arial", 11), height=2, width=15).pack(side=tk.LEFT, padx=5)
        
        # Log area
        progress_frame = tk.LabelFrame(
            parent,
            text="2. Processing Log",
            font=("Arial", 11, "bold"),
            padx=15,
            pady=15
        )
        progress_frame.pack(fill=tk.BOTH, expand=True)
        
        self.regen_log_text = scrolledtext.ScrolledText(
            progress_frame,
            height=12,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#ecf0f1",
            insertbackground="white"
        )
        self.regen_log_text.pack(fill=tk.BOTH, expand=True)
    
    def browse_regen_folder(self):
        """Browse for regenerate folder."""
        folder = filedialog.askdirectory(title="Select Folder with Edited Excel Files")
        if folder:
            self.regen_folder.set(folder)
            self.regen_log("Selected folder: " + folder)
    
    def open_regen_folder(self):
        """Open regenerate folder."""
        folder = self.regen_folder.get()
        if folder and Path(folder).exists():
            if os.name == 'nt':
                os.startfile(folder)
            else:
                os.system(f'open "{folder}"' if os.uname().sysname == 'Darwin' else f'xdg-open "{folder}"')
        else:
            messagebox.showwarning("Warning", "Folder not set or doesn't exist")
    
    def regen_log(self, message, level="INFO"):
        """Log to regenerate section."""
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {
            "INFO": "ℹ️",
            "SUCCESS": "✅",
            "WARNING": "⚠️",
            "ERROR": "❌"
        }.get(level, "•")
        
        gui_message = f"[{timestamp}] {prefix} {message}"
        self.regen_log_text.insert(tk.END, gui_message + "\n")
        self.regen_log_text.see(tk.END)
        self.root.update()
    
    def view_removed_records(self):
        """Open removed records file if it exists."""
        if not self.logger:
            messagebox.showinfo("Info", "No processing has been done yet.")
            return
        
        removed_file = self.logger.removed_excel
        if removed_file.exists():
            if os.name == 'nt':
                os.startfile(removed_file)
            else:
                os.system(f'open "{removed_file}"' if os.uname().sysname == 'Darwin' else f'xdg-open "{removed_file}"')
        else:
            messagebox.showinfo("Info", "No records were removed in the last processing.")
    
    def view_school_list(self):
        """Create and display school list."""
        if not self.logger:
            messagebox.showinfo("Info", "No processing has been done yet.")
            return
        
        # Create school list file
        school_list_file = self.logger.output_folder / f"School_Names_List_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        
        # Extract all schools from removed and processed records
        all_schools_set = set()
        
        # Try to read all Excel files in output folder
        for category, file_name in self.output_names.items():
            excel_path = self.logger.output_folder / f"{file_name}.xlsx"
            try:
                if excel_path.exists():
                    df = pd.read_excel(excel_path)
                    if 'SCHOOL NAME' in df.columns:
                        schools = set(df['SCHOOL NAME'].unique())
                        all_schools_set.update(schools)
            except:
                pass
        
        if all_schools_set:
            with open(school_list_file, 'w', encoding='utf-8') as f:
                f.write("ALL UNIQUE SCHOOL NAMES - REFERENCE DOCUMENT\n")
                f.write("=" * 80 + "\n")
                f.write(f"Generated: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n")
                f.write(f"Total unique schools: {len(all_schools_set)}\n")
                f.write("=" * 80 + "\n\n")
                
                for idx, school in enumerate(sorted(all_schools_set), 1):
                    f.write(f"{idx:3d}. {school}\n")
            
            if os.name == 'nt':
                os.startfile(school_list_file)
            else:
                os.system(f'open "{school_list_file}"' if os.uname().sysname == 'Darwin' else f'xdg-open "{school_list_file}"')
        else:
            messagebox.showinfo("Info", "No schools found in the processed data.")
    
    def start_regeneration(self):
        """Start TXT file regeneration."""
        if not self.regen_folder.get():
            messagebox.showerror("Error", "Please select the folder with edited Excel files")
            return
        
        if not Path(self.regen_folder.get()).exists():
            messagebox.showerror("Error", f"Folder not found:\n{self.regen_folder.get()}")
            return
        
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        
        thread = threading.Thread(target=self.regenerate_txt_files)
        thread.daemon = True
        thread.start()
    
    def regenerate_txt_files(self):
        """Regenerate TXT files from cleaned Excel files."""
        self.processing = True
        self.regen_btn.config(state=tk.DISABLED, bg="#95a5a6")
        self.status_bar.config(text="Regenerating TXT files...", bg="#f39c12")
        
        try:
            folder = Path(self.regen_folder.get())
            
            self.regen_log("="*70)
            self.regen_log("Starting TXT File Regeneration", "INFO")
            self.regen_log("="*70)
            self.regen_log(f"Reading from: {folder}")
            
            files_found = 0
            files_processed = 0
            warnings_total = 0
            
            for category, file_name in self.output_names.items():
                excel_path = folder / f"{file_name}.xlsx"
                txt_path = folder / f"{file_name}.txt"
                
                if not excel_path.exists():
                    self.regen_log(f"⏭️  Skipping: {excel_path.name} (not found)", "WARNING")
                    continue
                
                files_found += 1
                self.regen_log(f"\nProcessing: {file_name}")
                self.regen_log(f"  Reading: {excel_path.name}")
                
                try:
                    df = pd.read_excel(excel_path)
                    
                    required_cols = ['NRIC', 'SCHOOL NAME', 'STATUTORY NAME']
                    missing_cols = [col for col in required_cols if col not in df.columns]
                    
                    if missing_cols:
                        self.regen_log(f"  ❌ ERROR: Missing columns: {missing_cols}", "ERROR")
                        self.regen_log(f"     Found columns: {list(df.columns)}", "ERROR")
                        continue
                    
                    text_content, warnings = self.fixed_width_format(df)
                    txt_path.write_text(text_content, encoding='utf-8')
                    
                    self.regen_log(f"  ✅ Saved: {txt_path.name}")
                    self.regen_log(f"     Records: {len(df)}")
                    
                    if warnings:
                        self.regen_log(f"     ⚠️  Warnings: {len(warnings)}", "WARNING")
                        warnings_total += len(warnings)
                        for warning in warnings[:3]:
                            self.regen_log(f"        {warning}", "WARNING")
                        if len(warnings) > 3:
                            self.regen_log(f"        ... and {len(warnings) - 3} more", "WARNING")
                    
                    files_processed += 1
                
                except Exception as e:
                    self.regen_log(f"  ❌ Error processing {file_name}: {str(e)}", "ERROR")
            
            self.regen_log("\n" + "="*70)
            self.regen_log(f"REGENERATION COMPLETE!", "SUCCESS")
            self.regen_log(f"Files found: {files_found}")
            self.regen_log(f"Files processed: {files_processed}")
            self.regen_log(f"Total warnings: {warnings_total}")
            self.regen_log(f"Output folder: {folder}")
            self.regen_log("="*70)
            
            if files_processed > 0:
                self.status_bar.config(text=f"Successfully regenerated {files_processed} TXT files!", bg="#27ae60")
                messagebox.showinfo(
                    "Success",
                    f"TXT file regeneration complete!\n\n"
                    f"Files processed: {files_processed}\n"
                    f"Warnings: {warnings_total}\n\n"
                    f"All TXT files are ready in:\n{folder}"
                )
            else:
                self.status_bar.config(text="No Excel files found to process", bg="#e74c3c")
                messagebox.showwarning(
                    "Warning",
                    f"No Excel files found in the selected folder.\n\n"
                    f"Make sure the folder contains files named:\n"
                    f"  - Organisation_MTSCTP PSLE.xlsx\n"
                    f"  - Organisation_MTSCTP SEC 4 NA.xlsx\n"
                    f"  - Organisation_MTSCTP SEC 4 NT.xlsx\n"
                    f"  - Organisation_MTSCTP SEC 4 EX.xlsx"
                )
        
        except Exception as e:
            self.regen_log(f"ERROR: {str(e)}", "ERROR")
            self.status_bar.config(text="Error during regeneration", bg="#e74c3c")
            messagebox.showerror("Error", f"Regeneration failed:\n\n{str(e)}")
        
        finally:
            self.processing = False
            self.regen_btn.config(state=tk.NORMAL, bg="#2980b9")
    
    # ===== UTILITY METHODS =====
    
    def browse_mts_file(self):
        filename = filedialog.askopenfilename(
            title="Select MAP Students Attendance File",
            filetypes=[("CSV files", "*.csv"), ("All files", "*.*")]
        )
        if filename:
            self.mts_file.set(filename)
            self.log_gui(f"Selected MTS file: {Path(filename).name}")
    
    def browse_ccis_file(self):
        filename = filedialog.askopenfilename(
            title="Select Salesforce/YM Hub File",
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
        
        if self.logger:
            self.logger.log(message, level)
    
    def clear_log(self):
        if self.mode == "process":
            self.log_text.delete(1.0, tk.END)
        else:
            self.regen_log_text.delete(1.0, tk.END)
    
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
            removed_dupes = df_export[df_export.duplicated(keep='first')]
            
            for idx, row in removed_dupes.iterrows():
                self.logger.add_removed_record(
                    category=category,
                    reason="Duplicate Entry",
                    nric=str(row['NRIC']),
                    school=str(row['SCHOOL NAME']),
                    name=str(row['STATUTORY NAME'])
                )
            
            df_export = df_export.drop_duplicates()
            duplicates = before - len(df_export)
            if duplicates > 0:
                self.log_gui(f"Removed {duplicates} duplicates", "WARNING")
        
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
