import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re
from pathlib import Path
import threading
from datetime import datetime
import sys


class ExcelToTXTCheckerApp:
    """
    Updated Flow:

    1) Select Excel file
       - reads headers immediately
       - enables column mapping dropdowns

    2) Column Mapping:
       - user maps: NRIC / SCHOOL / NAME / LEVEL / STREAM

    3) Validate Excel:
       - required columns = mapped columns
       - NRIC: exactly 9 chars AND regex ^[STFGM]\\d{7}[A-Z]$
       - SCHOOL max 66
       - NAME max 66
       - LEVEL/STREAM must map into one of: PSLE / NA / NT / EX (else row is problematic)
       - logs problematic rows (shows preview)

    4) If problems:
       A) Remove problematic rows (drops in-memory)
          - ALSO exports an Excel file showing removed rows + reasons
          - ALSO splits cleaned data into separate Excel files by group (PSLE, NA, NT, EX)
       B) Or user edits Excel + re-validates
       (TXT generation is blocked until clean)

    5) Generate TXT:
       - reads from the split Excel files
       - writes TXT_OUTPUT_YYYYMMDD_HHMMSS inside the Excel file's folder
       - creates up to 4 .txt files based on LEVEL+STREAM grouping
    """

    def __init__(self, root):
        self.root = root
        self.root.title("MOE-JTA Excel Checker → TXT Generator")
        self.root.geometry("1050x780")
        self.root.resizable(True, True)

        # UI / State
        self.processing = False
        self.excel_path_var = tk.StringVar()

        # Output file names (base)
        self.output_names = {
            "PSLE": "MENDAKI_MTSCTP PSLE",
            "NA": "MENDAKI_MTSCTP SEC 4 NA",
            "NT": "MENDAKI_MTSCTP SEC 4 NT",
            "EX": "MENDAKI_MTSCTP SEC 4 EX",
        }

        # Column mapping
        self.col_nric = tk.StringVar(value="")
        self.col_school = tk.StringVar(value="")
        self.col_name = tk.StringVar(value="")
        self.col_level = tk.StringVar(value="")
        self.col_stream = tk.StringVar(value="")
        self.available_headers = []

        # Data storage
        self.original_df = None              # the loaded Excel data
        self.cleaned_df = None               # in-memory cleaned data after removal
        self.bad_row_mask = None             # boolean series True=bad
        self.bad_row_reasons = None          # list of reasons per row (same length as df)
        self.total_bad_rows = 0

        # Removed rows audit storage
        self.removed_rows_audit_df = None    # DataFrame of removed rows + reasons + group

        # Split Excel files storage
        self.excel_output_folder = None      # Path to folder containing split Excel files

        # Control flags
        self.file_loaded = False
        self.validation_passed = False
        self.block_txt_generation = True

        self.setup_ui()

    def open_moe_school_list(self, event=None):
        """
        Opens the bundled MOE school list file that sits next to the app.
        Works in dev (.py) and packaged (.exe).
        """
        if getattr(sys, 'frozen', False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        file_path = os.path.join(base_dir, "MOE School List.xlsx")

        if not os.path.exists(file_path):
            messagebox.showerror(
                "File not found",
                "MOE School List.xlsx was not found.\n\n"
                "Please ensure the app was installed with the school list file."
            )
            return

        try:
            if os.name == "nt":
                os.startfile(file_path)
            elif sys.platform == "darwin":
                os.system(f'open "{file_path}"')
            else:
                os.system(f'xdg-open "{file_path}"')
        except Exception as e:
            messagebox.showerror("Error opening file", f"Could not open MOE school list:\n\n{e}")

    # ---------------- UI ----------------
    def setup_ui(self):
        header = tk.Frame(self.root, bg="#2c3e50", height=70)
        header.pack(fill=tk.X)

        title = tk.Label(
            header,
            text="MOE-JTA Excel Checker → TXT Generator",
            font=("Arial", 16, "bold"),
            bg="#2c3e50",
            fg="white",
        )
        title.pack(pady=18)

        main = tk.Frame(self.root, padx=18, pady=18)
        main.pack(fill=tk.BOTH, expand=True)

        # Excel selection (replaces folder+scan)
        file_frame = tk.LabelFrame(main, text="1) Select Excel File", font=("Arial", 11, "bold"), padx=12, pady=12)
        file_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(file_frame, text="Excel file:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        tk.Entry(file_frame, textvariable=self.excel_path_var, width=70).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse...", width=12, command=self.browse_excel).grid(row=0, column=2, pady=5)

        hint = tk.Label(
            file_frame,
            text="Selecting an Excel file will load headers and enable column mapping.",
            font=("Arial", 9, "italic")
        )
        hint.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=6)

        # Column mapping (disabled until file load)
        mapping_frame = tk.LabelFrame(main, text="2) Column Mapping (enabled after file selection)", font=("Arial", 11, "bold"), padx=12, pady=12)
        mapping_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(mapping_frame, text="NRIC column:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=4)
        self.nric_menu = tk.OptionMenu(mapping_frame, self.col_nric, "")
        self.nric_menu.config(width=45, state=tk.DISABLED)
        self.nric_menu.grid(row=0, column=1, sticky=tk.W, padx=10)

        tk.Label(mapping_frame, text="SCHOOL column:", font=("Arial", 10)).grid(row=1, column=0, sticky=tk.W, pady=4)
        self.school_menu = tk.OptionMenu(mapping_frame, self.col_school, "")
        self.school_menu.config(width=45, state=tk.DISABLED)
        self.school_menu.grid(row=1, column=1, sticky=tk.W, padx=10)

        tk.Label(mapping_frame, text="NAME column:", font=("Arial", 10)).grid(row=2, column=0, sticky=tk.W, pady=4)
        self.name_menu = tk.OptionMenu(mapping_frame, self.col_name, "")
        self.name_menu.config(width=45, state=tk.DISABLED)
        self.name_menu.grid(row=2, column=1, sticky=tk.W, padx=10)

        tk.Label(mapping_frame, text="LEVEL column:", font=("Arial", 10)).grid(row=3, column=0, sticky=tk.W, pady=4)
        self.level_menu = tk.OptionMenu(mapping_frame, self.col_level, "")
        self.level_menu.config(width=45, state=tk.DISABLED)
        self.level_menu.grid(row=3, column=1, sticky=tk.W, padx=10)

        tk.Label(mapping_frame, text="STREAM column:", font=("Arial", 10)).grid(row=4, column=0, sticky=tk.W, pady=4)
        self.stream_menu = tk.OptionMenu(mapping_frame, self.col_stream, "")
        self.stream_menu.config(width=45, state=tk.DISABLED)
        self.stream_menu.grid(row=4, column=1, sticky=tk.W, padx=10)

        instruction_text = (
            "Before proceeding:\n"
            "1) Verify SCHOOL spelling against the official 'MOE School List.xlsx' file\n"
            "2) Ensure NRIC follows the standard format (e.g. S1234567A)\n"
            "3) Ensure SCHOOL and STATUTORY NAMES do not exceed 66 characters\n"
            "4) Ensure LEVEL+STREAM values can be mapped into PSLE / NA / NT / EXPRESS\n\n"
            "Note: The MOE school list is installed together with this application."
        )

        instruction_label = tk.Label(
            mapping_frame,
            text=instruction_text,
            font=("Arial", 9),
            justify=tk.LEFT,
            fg="#2c3e50",
            wraplength=800
        )
        instruction_label.grid(row=0, column=2, rowspan=5, sticky=tk.NW, padx=(20, 0), pady=2)

        moe_file_link = tk.Label(
            mapping_frame,
            text="📄 Open official MOE school list (reference)",
            font=("Arial", 9, "underline"),
            fg="#1a73e8",
            cursor="hand2"
        )
        moe_file_link.grid(row=6, column=2, sticky=tk.W, padx=(20, 0), pady=(6, 0))
        moe_file_link.bind("<Button-1>", self.open_moe_school_list)

        # Actions
        btn_frame = tk.LabelFrame(main, text="3) Actions (Validate -> Remove Rows -> Generate TXT files)", font=("Arial", 11, "bold"), padx=12, pady=12)
        btn_frame.pack(fill=tk.X, pady=(0, 12))

        self.validate_btn = tk.Button(
            btn_frame,
            text="✅ Validate Excel File",
            command=self.start_validate,
            font=("Arial", 12, "bold"),
            height=2,
            width=24,
            state=tk.DISABLED
        )
        self.validate_btn.pack(side=tk.LEFT, padx=6)

        self.remove_btn = tk.Button(
            btn_frame,
            text="🧹 Remove Problem Rows",
            command=self.start_remove_rows,
            font=("Arial", 12, "bold"),
            height=2,
            width=24,
            state=tk.DISABLED
        )
        self.remove_btn.pack(side=tk.LEFT, padx=6)

        self.generate_btn = tk.Button(
            btn_frame,
            text="▶ Generate TXT Files",
            command=self.start_generate_txt,
            font=("Arial", 12, "bold"),
            height=2,
            width=22,
            state=tk.DISABLED
        )
        self.generate_btn.pack(side=tk.LEFT, padx=6)

        tk.Button(
            btn_frame,
            text="📁 Open Folder",
            command=self.open_folder,
            font=("Arial", 11),
            height=2,
            width=16
        ).pack(side=tk.LEFT, padx=6)

        tk.Button(
            btn_frame,
            text="Clear Log",
            command=self.clear_log,
            font=("Arial", 11),
            height=2,
            width=12
        ).pack(side=tk.LEFT, padx=6)

        # Log
        log_frame = tk.LabelFrame(main, text="Log", font=("Arial", 11, "bold"), padx=12, pady=12)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=18,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#ecf0f1",
            insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

        # Status bar
        self.status_bar = tk.Label(
            self.root,
            text="Ready",
            bd=1,
            relief=tk.SUNKEN,
            anchor=tk.W,
            bg="#ecf0f1"
        )
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X)

    # ---------------- UI helpers ----------------
    def browse_excel(self):
        file_path = filedialog.askopenfilename(
            title="Select Excel File",
            filetypes=[("Excel Files", "*.xlsx *.xls")]
        )
        if not file_path:
            return

        self.excel_path_var.set(file_path)
        self.log(f"Selected Excel: {file_path}", "INFO")

        # Reset state when file changes
        self.reset_state_after_file_change()

        # Load headers immediately (in a thread so UI doesn't freeze)
        t = threading.Thread(target=self.load_excel_headers)
        t.daemon = True
        t.start()

    def reset_state_after_file_change(self):
        self.file_loaded = False
        self.validation_passed = False
        self.block_txt_generation = True

        self.available_headers = []
        self.col_nric.set("")
        self.col_school.set("")
        self.col_name.set("")
        self.col_level.set("")
        self.col_stream.set("")

        self.original_df = None
        self.cleaned_df = None
        self.bad_row_mask = None
        self.bad_row_reasons = None
        self.total_bad_rows = 0
        self.removed_rows_audit_df = None
        self.excel_output_folder = None

        self.set_mapping_enabled(False)
        self.validate_btn.config(state=tk.DISABLED)
        self.remove_btn.config(state=tk.DISABLED)
        self.generate_btn.config(state=tk.DISABLED)

    def set_mapping_enabled(self, enabled: bool):
        state = tk.NORMAL if enabled else tk.DISABLED
        self.nric_menu.config(state=state)
        self.school_menu.config(state=state)
        self.name_menu.config(state=state)
        self.level_menu.config(state=state)
        self.stream_menu.config(state=state)

    def populate_column_menus(self, headers):
        def reset_menu(menu_widget, var):
            menu_widget["menu"].delete(0, "end")
            for h in headers:
                menu_widget["menu"].add_command(label=h, command=tk._setit(var, h))

        reset_menu(self.nric_menu, self.col_nric)
        reset_menu(self.school_menu, self.col_school)
        reset_menu(self.name_menu, self.col_name)
        reset_menu(self.level_menu, self.col_level)
        reset_menu(self.stream_menu, self.col_stream)

    def open_folder(self):
        """
        Opens the folder containing the selected Excel file.
        """
        excel_path = self.excel_path_var.get().strip()
        if not excel_path or not Path(excel_path).exists():
            messagebox.showwarning("Warning", "Excel file not set or doesn't exist")
            return

        folder = str(Path(excel_path).parent)
        try:
            if os.name == "nt":
                os.startfile(folder)
            else:
                if hasattr(os, "uname") and os.uname().sysname == "Darwin":
                    os.system(f'open "{folder}"')
                else:
                    os.system(f'xdg-open "{folder}"')
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"INFO": "ℹ️", "SUCCESS": "✅", "WARNING": "⚠️", "ERROR": "❌"}.get(level, "•")
        line = f"[{timestamp}] {prefix} {message}"
        self.log_text.insert(tk.END, line + "\n")
        self.log_text.see(tk.END)
        self.root.update()

    def set_busy(self, busy: bool, status_text: str, status_color: str):
        self.processing = busy
        self.status_bar.config(text=status_text, bg=status_color)

        if busy:
            self.validate_btn.config(state=tk.DISABLED)
            self.remove_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)

    # ---------------- Load Excel headers ----------------
    def load_excel_headers(self):
        self.set_busy(True, "Loading Excel headers...", "#f39c12")
        try:
            excel_path = self.excel_path_var.get().strip()
            if not excel_path or not Path(excel_path).exists():
                raise FileNotFoundError("Excel file not found.")

            df = pd.read_excel(excel_path)
            self.original_df = df
            self.cleaned_df = df.copy()

            self.available_headers = list(df.columns)
            self.populate_column_menus(self.available_headers)
            self.set_mapping_enabled(True)

            # Auto-select only if exact names exist
            self.col_nric.set("NRIC" if "NRIC" in self.available_headers else "")
            self.col_school.set("SCHOOL NAME" if "SCHOOL NAME" in self.available_headers else "")
            self.col_name.set("STATUTORY NAME" if "STATUTORY NAME" in self.available_headers else "")
            self.col_level.set("LEVEL" if "LEVEL" in self.available_headers else "")
            self.col_stream.set("STREAM" if "STREAM" in self.available_headers else "")

            self.file_loaded = True
            self.validate_btn.config(state=tk.NORMAL)
            self.remove_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)

            self.status_bar.config(text="File loaded. Select columns, then validate.", bg="#27ae60")

            if not (self.col_nric.get() and self.col_school.get() and self.col_name.get() and self.col_level.get() and self.col_stream.get()):
                messagebox.showinfo(
                    "Select Columns",
                    "File loaded.\n\nPlease choose NRIC / SCHOOL / NAME / LEVEL / STREAM columns, then click 'Validate Excel File'."
                )
            else:
                messagebox.showinfo(
                    "Ready to Validate",
                    "File loaded and some standard headers were auto-detected.\n\nClick 'Validate Excel File' to proceed."
                )

        except Exception as e:
            self.log(f"ERROR loading Excel: {e}", "ERROR")
            self.status_bar.config(text="Error loading Excel.", bg="#e74c3c")
            messagebox.showerror("Error", f"Failed to load Excel:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            if self.file_loaded:
                self.status_bar.config(text="File loaded. Select columns, then validate.", bg="#27ae60")

    # ---------------- Validation rules ----------------
    @staticmethod
    def safe_str(x) -> str:
        if pd.isna(x):
            return ""
        return str(x)

    @staticmethod
    def nric_is_valid(nric_value: str) -> bool:
        s = str(nric_value).strip().upper()
        if len(s) != 9:
            return False
        pattern = r"^[STFGM]\d{7}[A-Z]$"
        return bool(re.match(pattern, s))

    @staticmethod
    def normalize_text(x: str) -> str:
        """
        Used only for LEVEL/STREAM grouping logic:
        - uppercases
        - collapses multiple spaces
        """
        s = str(x).strip().upper()
        s = re.sub(r"\s+", " ", s)
        return s

    def infer_group(self, level_value: str, stream_value: str):
        """
        Uses the EXACT same logic as the old data processor app.
        
        Grouping rules:
        - PSLE: Level == 'P6'
        - NA: Level == 'S4' AND Stream in ['G2', 'Normal Academic']
        - NT: Level == 'S4' AND Stream in ['G1', 'Normal Technical']
        - EX: Level == 'S4' AND Stream in ['G3', 'Express']
        """
        lvl = str(level_value).strip()
        stm = str(stream_value).strip()

        # PSLE: Level == 'P6'
        if lvl == 'P6':
            return "PSLE"

        # Secondary 4 streams (must be S4 AND match stream)
        if lvl == 'S4':
            if stm in ['G2', 'Normal Academic']:
                return "NA"
            if stm in ['G1', 'Normal Technical']:
                return "NT"
            if stm in ['G3', 'Express']:
                return "EX"

        return None

    def validate_dataframe(self, df: pd.DataFrame, nric_col: str, school_col: str, name_col: str, level_col: str, stream_col: str):
        required_cols = [nric_col, school_col, name_col, level_col, stream_col]
        missing_cols = [c for c in required_cols if c not in df.columns]

        if missing_cols:
            bad_mask = pd.Series([False] * len(df), index=df.index)
            reasons = [""] * len(df)
            return missing_cols, bad_mask, reasons, []

        nric_series = df[nric_col].apply(self.safe_str)
        school_series = df[school_col].apply(self.safe_str)
        name_series = df[name_col].apply(self.safe_str)
        level_series = df[level_col].apply(self.safe_str)
        stream_series = df[stream_col].apply(self.safe_str)

        bad_nric = ~nric_series.apply(self.nric_is_valid)
        bad_school_len = school_series.str.strip().str.len() > 66
        bad_name_len = name_series.str.strip().str.len() > 66

        # grouping check
        groups = []
        bad_group = []
        for lvl, stm in zip(level_series.tolist(), stream_series.tolist()):
            g = self.infer_group(lvl, stm)
            groups.append(g)
            bad_group.append(g is None)
        bad_group = pd.Series(bad_group, index=df.index)

        bad_mask = bad_nric | bad_school_len | bad_name_len | bad_group

        # build reasons list (same length as df)
        reasons = []
        for i, idx in enumerate(df.index):
            r = []
            nric = self.safe_str(df.loc[idx, nric_col]).strip()
            school = self.safe_str(df.loc[idx, school_col]).strip()
            name = self.safe_str(df.loc[idx, name_col]).strip()
            lvl = self.safe_str(df.loc[idx, level_col]).strip()
            stm = self.safe_str(df.loc[idx, stream_col]).strip()

            if not self.nric_is_valid(nric):
                r.append(f"NRIC invalid (value='{nric}')")
            if len(school) > 66:
                r.append(f"SCHOOL too long ({len(school)})")
            if len(name) > 66:
                r.append(f"NAME too long ({len(name)})")
            if self.infer_group(lvl, stm) is None:
                r.append(f"LEVEL/STREAM cannot map (LEVEL='{lvl}', STREAM='{stm}')")

            reasons.append("; ".join(r))

        messages = []
        bad_rows = df[bad_mask]

        for idx, row in bad_rows.iterrows():
            excel_row_num = df.index.get_loc(idx) + 2

            nric = self.safe_str(row[nric_col]).strip()
            school = self.safe_str(row[school_col]).strip()
            name = self.safe_str(row[name_col]).strip()
            lvl = self.safe_str(row[level_col]).strip()
            stm = self.safe_str(row[stream_col]).strip()

            school_preview = school[:60] + ("..." if len(school) > 60 else "")
            name_preview = name[:60] + ("..." if len(name) > 60 else "")

            messages.append(
                f"Row {excel_row_num}: {reasons[df.index.get_loc(idx)]} | "
                f"NRIC='{nric}' | SCHOOL='{school_preview}' | NAME='{name_preview}' | LEVEL='{lvl}' | STREAM='{stm}'"
            )

        return [], bad_mask, reasons, messages

    # ---------------- Stage 1: Validate ----------------
    def start_validate(self):
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        if not self.file_loaded or self.cleaned_df is None:
            messagebox.showwarning("Select File First", "Please select an Excel file first.")
            return

        nric_col = self.col_nric.get().strip()
        school_col = self.col_school.get().strip()
        name_col = self.col_name.get().strip()
        level_col = self.col_level.get().strip()
        stream_col = self.col_stream.get().strip()

        if not (nric_col and school_col and name_col and level_col and stream_col):
            messagebox.showerror("Missing Mapping", "Please select NRIC / SCHOOL / NAME / LEVEL / STREAM columns before validating.")
            return

        t = threading.Thread(target=self.validate_excel)
        t.daemon = True
        t.start()

    def validate_excel(self):
        self.set_busy(True, "Validating Excel file...", "#f39c12")

        # Reset validation-related state
        self.bad_row_mask = None
        self.bad_row_reasons = None
        self.total_bad_rows = 0
        self.validation_passed = False
        self.block_txt_generation = True
        self.remove_btn.config(state=tk.DISABLED)
        self.generate_btn.config(state=tk.DISABLED)
        self.removed_rows_audit_df = None
        self.excel_output_folder = None

        try:
            df = self.cleaned_df.copy()
            nric_col = self.col_nric.get().strip()
            school_col = self.col_school.get().strip()
            name_col = self.col_name.get().strip()
            level_col = self.col_level.get().strip()
            stream_col = self.col_stream.get().strip()

            self.log("=" * 90, "INFO")
            self.log("VALIDATION START", "INFO")
            self.log(f"Mapping: NRIC='{nric_col}', SCHOOL='{school_col}', NAME='{name_col}', LEVEL='{level_col}', STREAM='{stream_col}'", "INFO")
            self.log("=" * 90, "INFO")

            missing_cols, bad_mask, reasons, row_messages = self.validate_dataframe(
                df, nric_col, school_col, name_col, level_col, stream_col
            )

            if missing_cols:
                self.log(f"❌ FAIL: Missing mapped columns: {missing_cols}", "ERROR")
                self.log(f"Found columns: {list(df.columns)}", "ERROR")
                self.status_bar.config(text="Validation failed: missing mapped columns.", bg="#e74c3c")
                messagebox.showerror(
                    "Validation Failed",
                    "The Excel file is missing one or more columns you mapped.\n\n"
                    "Fix the Excel headers or change the mapping, then validate again."
                )
                return

            self.bad_row_mask = bad_mask
            self.bad_row_reasons = reasons
            self.total_bad_rows = int(bad_mask.sum())

            self.log("VALIDATION COMPLETE", "INFO")
            self.log(f"Rows checked: {len(df)}", "INFO")
            self.log(f"Problematic rows: {self.total_bad_rows}", "INFO")
            self.log("=" * 90, "INFO")

            if self.total_bad_rows > 0:
                self.log("⚠️ Issues found. Showing up to 10 examples:", "WARNING")
                shown = 0
                for msg in row_messages:
                    self.log(f"  {msg}", "WARNING")
                    shown += 1
                    if shown >= 10:
                        remaining = self.total_bad_rows - shown
                        if remaining > 0:
                            self.log(f"  ... and {remaining} more problematic rows", "WARNING")
                        break

                self.status_bar.config(text="Issues found. Remove/edit rows before TXT generation.", bg="#f39c12")
                self.validation_passed = False
                self.block_txt_generation = True
                self.remove_btn.config(state=tk.NORMAL)
                self.generate_btn.config(state=tk.DISABLED)

                messagebox.showwarning(
                    "Issues Found",
                    f"Found {self.total_bad_rows} problematic rows.\n\n"
                    "Option A: Remove problem rows (in the app)\n"
                    "Option B: Edit Excel manually and validate again\n\n"
                    "TXT generation is blocked until clean."
                )
                return

            # Clean data - if no problems, go straight to splitting Excel files
            self.log("Data is clean. Proceeding to split into Excel files...", "SUCCESS")
            self.split_excel_files()

        except Exception as e:
            self.log(f"ERROR during validation: {e}", "ERROR")
            self.status_bar.config(text="Error during validation.", bg="#e74c3c")
            messagebox.showerror("Error", f"Validation failed:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            if self.validation_passed:
                self.status_bar.config(text="Excel files created. Ready to generate TXT.", bg="#27ae60")

    # ---------------- Stage 2: Remove rows (+ export audit Excel + split Excel files) ----------------
    def start_remove_rows(self):
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        if self.cleaned_df is None or self.bad_row_mask is None:
            messagebox.showerror("Error", "Validate first.")
            return
        if self.total_bad_rows == 0:
            messagebox.showinfo("Info", "No problematic rows to remove.")
            return

        confirm = messagebox.askyesno(
            "Remove Problem Rows",
            "This will DROP all problematic rows in-memory.\n"
            "It will ALSO export an Excel file showing removed rows + reasons.\n"
            "Then it will split the cleaned data into separate Excel files by group.\n\n"
            "Proceed?"
        )
        if not confirm:
            return

        t = threading.Thread(target=self.remove_problem_rows)
        t.daemon = True
        t.start()

    def remove_problem_rows(self):
        self.set_busy(True, "Removing problematic rows...", "#f39c12")

        try:
            df = self.cleaned_df.copy()
            mask = self.bad_row_mask.copy()
            reasons = list(self.bad_row_reasons)

            # Build removed rows audit table
            removed_df = df[mask].copy()
            if not removed_df.empty:
                removed_df.insert(0, "REMOVAL_REASON", [reasons[df.index.get_loc(i)] for i in removed_df.index])

                # add inferred group for removed rows (useful for debugging)
                level_col = self.col_level.get().strip()
                stream_col = self.col_stream.get().strip()

                groups = []
                for _, r in removed_df.iterrows():
                    g = self.infer_group(r.get(level_col, ""), r.get(stream_col, ""))
                    groups.append(g if g is not None else "UNKNOWN")
                removed_df.insert(1, "OUTPUT_GROUP", groups)

            # Remove in-memory
            self.cleaned_df = df[~mask].copy()

            removed_total = int(mask.sum())
            self.total_bad_rows = 0
            self.validation_passed = True
            self.block_txt_generation = False
            self.remove_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.NORMAL)

            # Save audit excel next to the input excel
            excel_path = Path(self.excel_path_var.get().strip())
            base_folder = excel_path.parent
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            audit_path = base_folder / f"REMOVED_ROWS_AUDIT_{ts}.xlsx"

            if removed_total > 0 and not removed_df.empty:
                with pd.ExcelWriter(audit_path, engine="openpyxl") as writer:
                    removed_df.to_excel(writer, sheet_name="Removed Rows", index=False)
                self.removed_rows_audit_df = removed_df
                self.log(f"Removed {removed_total} problematic rows.", "SUCCESS")
                self.log(f"Saved removed-rows audit: {audit_path.name}", "SUCCESS")
            else:
                self.log("No rows removed (unexpected state).", "WARNING")

            # Now split the cleaned data into Excel files
            self.split_excel_files()

            self.status_bar.config(text="Rows removed. Excel files created. Ready to generate TXT.", bg="#27ae60")
            messagebox.showinfo(
                "Removed & Split",
                f"Removed {removed_total} problematic rows.\n\n"
                f"Split cleaned data into Excel files in:\n{self.excel_output_folder}\n\n"
                f"Audit Excel saved:\n{audit_path}"
            )

        except Exception as e:
            self.log(f"ERROR during removal: {e}", "ERROR")
            self.status_bar.config(text="Error during removal.", bg="#e74c3c")
            messagebox.showerror("Error", f"Removal failed:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            if self.validation_passed:
                self.status_bar.config(text="Rows removed. Excel files created. Ready to generate TXT.", bg="#27ae60")

    # ---------------- Split Excel files helper ----------------
    def split_excel_files(self):
        """
        Splits the cleaned dataframe into separate Excel files by group (PSLE, NA, NT, EX).
        This is called either after validation (if data is clean) or after removal.
        """
        try:
            excel_path = Path(self.excel_path_var.get().strip())
            base_folder = excel_path.parent

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            excel_out_folder = base_folder / f"EXCEL_OUTPUT_{ts}"
            excel_out_folder.mkdir(parents=True, exist_ok=True)
            self.excel_output_folder = excel_out_folder

            self.log("=" * 90, "INFO")
            self.log("SPLITTING CLEANED DATA INTO EXCEL FILES BY GROUP", "INFO")
            self.log(f"Excel output folder: {excel_out_folder}", "INFO")

            nric_col = self.col_nric.get().strip()
            school_col = self.col_school.get().strip()
            name_col = self.col_name.get().strip()
            level_col = self.col_level.get().strip()
            stream_col = self.col_stream.get().strip()

            # Group cleaned data
            grouped = {"PSLE": [], "NA": [], "NT": [], "EX": []}
            for _, row in self.cleaned_df.iterrows():
                g = self.infer_group(row.get(level_col, ""), row.get(stream_col, ""))
                if g in grouped:
                    grouped[g].append(row)

            excel_files_created = 0
            for group_key, rows in grouped.items():
                if not rows:
                    self.log(f"Skipping {group_key}: no rows in this group", "WARNING")
                    continue

                group_df = pd.DataFrame(rows)
                excel_name = self.output_names[group_key]
                excel_file_path = excel_out_folder / f"{excel_name}.xlsx"

                with pd.ExcelWriter(excel_file_path, engine="openpyxl") as writer:
                    group_df.to_excel(writer, sheet_name="Data", index=False)

                self.log(f"Created Excel: {excel_name}.xlsx (records={len(group_df)})", "SUCCESS")
                excel_files_created += 1

            self.log(f"Excel files created: {excel_files_created}", "SUCCESS")
            self.log("=" * 90, "INFO")

            # Enable TXT generation
            self.validation_passed = True
            self.block_txt_generation = False
            self.generate_btn.config(state=tk.NORMAL)

            if self.total_bad_rows == 0:
                # This means we came here directly from validation (no removal needed)
                self.status_bar.config(text="Excel files created. Ready to generate TXT.", bg="#27ae60")
                messagebox.showinfo(
                    "Excel Files Created",
                    f"Data is clean.\n\n"
                    f"Created {excel_files_created} Excel files by group in:\n{excel_out_folder}\n\n"
                    f"Click 'Generate TXT Files' to create the final .txt outputs."
                )

        except Exception as e:
            self.log(f"ERROR during Excel splitting: {e}", "ERROR")
            self.status_bar.config(text="Error during Excel splitting.", bg="#e74c3c")
            messagebox.showerror("Error", f"Excel splitting failed:\n\n{e}")

    # ---------------- Stage 3: Generate TXT ----------------
    def start_generate_txt(self):
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        if not self.validation_passed or self.block_txt_generation:
            messagebox.showwarning(
                "TXT Generation Blocked",
                "Data is not clean yet.\n\n"
                "Please validate and fix/remove problematic rows first."
            )
            return
        if self.excel_output_folder is None or not self.excel_output_folder.exists():
            messagebox.showwarning("No Excel Files", "Excel output folder not found. Please validate/remove rows first.")
            return

        t = threading.Thread(target=self.generate_txt_files)
        t.daemon = True
        t.start()

    def generate_txt_files(self):
        self.set_busy(True, "Generating TXT files...", "#f39c12")

        try:
            excel_path = Path(self.excel_path_var.get().strip())
            base_folder = excel_path.parent

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            txt_out_folder = base_folder / f"TXT_OUTPUT_{ts}"
            txt_out_folder.mkdir(parents=True, exist_ok=True)

            nric_col = self.col_nric.get().strip()
            school_col = self.col_school.get().strip()
            name_col = self.col_name.get().strip()

            self.log("=" * 90, "INFO")
            self.log("TXT GENERATION START", "INFO")
            self.log(f"Reading from: {self.excel_output_folder}", "INFO")
            self.log(f"TXT output folder: {txt_out_folder}", "INFO")
            self.log("=" * 90, "INFO")

            written = 0
            for group_key, base_name in self.output_names.items():
                excel_file = self.excel_output_folder / f"{base_name}.xlsx"

                if not excel_file.exists():
                    self.log(f"Skipping {group_key}: Excel file not found ({excel_file.name})", "WARNING")
                    continue

                # Read the Excel file
                group_df = pd.read_excel(excel_file)

                if group_df.empty:
                    self.log(f"Skipping {group_key}: no rows in Excel file", "WARNING")
                    continue

                # Generate TXT content
                txt_content = self.fixed_width_format(group_df, nric_col, school_col, name_col)
                txt_name = base_name
                txt_path = txt_out_folder / f"{txt_name}.txt"
                txt_path.write_text(txt_content, encoding="utf-8")

                self.log(f"Saved: {txt_path.name} (records={len(group_df)})", "SUCCESS")
                written += 1

            self.log("\n" + "=" * 90, "INFO")
            self.log("TXT GENERATION COMPLETE", "SUCCESS")
            self.log(f"TXT files written: {written}", "INFO")
            self.log("=" * 90, "INFO")

            self.status_bar.config(text=f"Done. Wrote {written} TXT files.", bg="#27ae60")
            messagebox.showinfo(
                "Done",
                f"TXT generation complete.\n\n"
                f"TXT files written: {written}\n\n"
                f"Output folder:\n{txt_out_folder}"
            )

        except Exception as e:
            self.log(f"ERROR during TXT generation: {e}", "ERROR")
            self.status_bar.config(text="Error during TXT generation.", bg="#e74c3c")
            messagebox.showerror("Error", f"TXT generation failed:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            self.status_bar.config(text="Ready", bg="#ecf0f1")

    # ---------------- Fixed-width formatting ----------------
    @staticmethod
    def fixed_width_format(df: pd.DataFrame, nric_col: str, school_col: str, name_col: str) -> str:
        lines = []
        for _, row in df.iterrows():
            nric = str(row[nric_col]).strip()[:9].ljust(9)
            school = str(row[school_col]).strip()[:66].ljust(66)
            name = str(row[name_col]).strip()[:66].ljust(66)
            lines.append(nric + school + name)
        return "\n".join(lines)


def main():
    root = tk.Tk()
    app = ExcelToTXTCheckerApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()