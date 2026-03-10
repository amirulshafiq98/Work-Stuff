import tkinter as tk
from tkinter import filedialog, messagebox, scrolledtext
import pandas as pd
import os
import re
from pathlib import Path
import threading
from datetime import datetime
import sys


class ExcelToMOEOutputsApp:
    """
    UI flow from app_ver2.py + strict export logic based on MOE-JTA Sample.ipynb.

    Expected setup:
    - User selects ONE master Excel/CSV file
    - File is already standardised before running
    - Exact standard values are expected after trimming spaces and uppercasing

    Required logic:
    - Keep only rows where SCHOOL CHECK == TRUE
    - Keep only rows where RACE == MALAY
    - Output groups:
        P6 + PROGRAM == MHC   -> ORG_MHC
        P6 + PROGRAM == SIPMS -> ORG_SIPMS
        P6                    -> ORG_MTSCTP PSLE
        S4 + G2               -> ORG_MTSCTP SEC 4 NA
        S4 + G1               -> ORG_MTSCTP SEC 4 NT
        S4 + G3               -> ORG_MTSCTP SEC 4 EX

    For each non-empty group the app exports:
    - Excel file
    - TXT file

    Also exports:
    - ALL_SCHOOLS.xlsx
    - REMOVED_ROWS_AUDIT_<timestamp>.xlsx when bad rows are removed in-app
    """

    def __init__(self, root):
        self.root = root
        self.root.title("MOE-JTA Excel Checker -> Excel + TXT Generator")
        self.root.geometry("1180x860")
        self.root.resizable(True, True)

        self.processing = False
        self.file_path_var = tk.StringVar()

        self.output_names = {
            "PSLE": "ORG_MTSCTP PSLE",
            "NA": "ORG_MTSCTP SEC 4 NA",
            "NT": "ORG_MTSCTP SEC 4 NT",
            "EX": "ORG_MTSCTP SEC 4 EX",
            "MHC": "ORG_MHC",
            "SIPMS": "ORG_SIPMS",
        }

        # Required column mappings
        self.col_nric = tk.StringVar(value="")
        self.col_school = tk.StringVar(value="")
        self.col_name = tk.StringVar(value="")
        self.col_level = tk.StringVar(value="")
        self.col_stream = tk.StringVar(value="")
        self.col_race = tk.StringVar(value="")
        self.col_school_check = tk.StringVar(value="")
        self.col_program = tk.StringVar(value="")

        self.available_headers = []

        self.original_df = None
        self.cleaned_df = None
        self.bad_row_mask = None
        self.bad_row_reasons = None
        self.total_bad_rows = 0
        self.removed_rows_audit_df = None

        self.file_loaded = False
        self.validation_passed = False
        self.block_generation = True

        self.setup_ui()

    # ---------------- UI ----------------
    def setup_ui(self):
        header = tk.Frame(self.root, bg="#2c3e50", height=70)
        header.pack(fill=tk.X)

        title = tk.Label(
            header,
            text="MOE-JTA Excel Checker -> Excel + TXT Generator",
            font=("Arial", 16, "bold"),
            bg="#2c3e50",
            fg="white",
        )
        title.pack(pady=18)

        main = tk.Frame(self.root, padx=18, pady=18)
        main.pack(fill=tk.BOTH, expand=True)

        file_frame = tk.LabelFrame(main, text="1) Select Master File", font=("Arial", 11, "bold"), padx=12, pady=12)
        file_frame.pack(fill=tk.X, pady=(0, 12))

        tk.Label(file_frame, text="Excel / CSV file:", font=("Arial", 10)).grid(row=0, column=0, sticky=tk.W, pady=5)
        tk.Entry(file_frame, textvariable=self.file_path_var, width=78).grid(row=0, column=1, padx=10, pady=5)
        tk.Button(file_frame, text="Browse...", width=12, command=self.browse_file).grid(row=0, column=2, pady=5)

        hint = tk.Label(
            file_frame,
            text="Selecting a file will load headers and enable column mapping.",
            font=("Arial", 9, "italic")
        )
        hint.grid(row=1, column=0, columnspan=3, sticky=tk.W, pady=6)

        mapping_frame = tk.LabelFrame(main, text="2) Column Mapping", font=("Arial", 11, "bold"), padx=12, pady=12)
        mapping_frame.pack(fill=tk.X, pady=(0, 12))

        labels_and_vars = [
            ("NRIC column:", "nric_menu", self.col_nric),
            ("SCHOOL column:", "school_menu", self.col_school),
            ("NAME column:", "name_menu", self.col_name),
            ("LEVEL column:", "level_menu", self.col_level),
            ("STREAM column:", "stream_menu", self.col_stream),
            ("RACE column:", "race_menu", self.col_race),
            ("SCHOOL CHECK column:", "school_check_menu", self.col_school_check),
            ("PROGRAM column:", "program_menu", self.col_program),
        ]

        for row_idx, (label_text, attr_name, var) in enumerate(labels_and_vars):
            tk.Label(mapping_frame, text=label_text, font=("Arial", 10)).grid(row=row_idx, column=0, sticky=tk.W, pady=4)
            menu = tk.OptionMenu(mapping_frame, var, "")
            menu.config(width=45, state=tk.DISABLED)
            menu.grid(row=row_idx, column=1, sticky=tk.W, padx=10)
            setattr(self, attr_name, menu)

        instruction_text = (
            "Strict standardised input expected:\n"
            "1) SCHOOL CHECK must be TRUE\n"
            "2) RACE must be MALAY\n"
            "3) LEVEL must be P6 or S4\n"
            "4) STREAM for S4 must be G1/G2/G3\n"
            "5) PROGRAM can be blank / MHC / SIPMS\n"
            "6) Output is Excel + TXT for each non-empty group\n"
        )

        instruction_label = tk.Label(
            mapping_frame,
            text=instruction_text,
            font=("Arial", 9),
            justify=tk.LEFT,
            fg="#2c3e50",
            wraplength=420
        )
        instruction_label.grid(row=0, column=2, rowspan=8, sticky=tk.NW, padx=(20, 0), pady=2)

        moe_file_link = tk.Label(
            mapping_frame,
            text="Open official MOE school list (reference)",
            font=("Arial", 9, "underline"),
            fg="#1a73e8",
            cursor="hand2"
        )
        moe_file_link.grid(row=8, column=2, sticky=tk.W, padx=(20, 0), pady=(6, 0))
        moe_file_link.bind("<Button-1>", self.open_moe_school_list)

        btn_frame = tk.LabelFrame(main, text="3) Actions", font=("Arial", 11, "bold"), padx=12, pady=12)
        btn_frame.pack(fill=tk.X, pady=(0, 12))

        self.validate_btn = tk.Button(
            btn_frame,
            text="Validate File",
            command=self.start_validate,
            font=("Arial", 12, "bold"),
            height=2,
            width=20,
            state=tk.DISABLED
        )
        self.validate_btn.pack(side=tk.LEFT, padx=6)

        self.remove_btn = tk.Button(
            btn_frame,
            text="Remove Problem Rows",
            command=self.start_remove_rows,
            font=("Arial", 12, "bold"),
            height=2,
            width=20,
            state=tk.DISABLED
        )
        self.remove_btn.pack(side=tk.LEFT, padx=6)

        self.generate_btn = tk.Button(
            btn_frame,
            text="Generate Excel + TXT",
            command=self.start_generate_outputs,
            font=("Arial", 12, "bold"),
            height=2,
            width=22,
            state=tk.DISABLED
        )
        self.generate_btn.pack(side=tk.LEFT, padx=6)

        tk.Button(
            btn_frame,
            text="Open Folder",
            command=self.open_folder,
            font=("Arial", 11),
            height=2,
            width=14
        ).pack(side=tk.LEFT, padx=6)

        tk.Button(
            btn_frame,
            text="Clear Log",
            command=self.clear_log,
            font=("Arial", 11),
            height=2,
            width=12
        ).pack(side=tk.LEFT, padx=6)

        log_frame = tk.LabelFrame(main, text="Log", font=("Arial", 11, "bold"), padx=12, pady=12)
        log_frame.pack(fill=tk.BOTH, expand=True)

        self.log_text = scrolledtext.ScrolledText(
            log_frame,
            height=20,
            font=("Consolas", 9),
            bg="#2c3e50",
            fg="#ecf0f1",
            insertbackground="white"
        )
        self.log_text.pack(fill=tk.BOTH, expand=True)

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
    def browse_file(self):
        file_path = filedialog.askopenfilename(
            title="Select Master Excel/CSV File",
            filetypes=[("Excel or CSV Files", "*.xlsx *.xls *.csv")]
        )
        if not file_path:
            return

        self.file_path_var.set(file_path)
        self.log(f"Selected file: {file_path}", "INFO")
        self.reset_state_after_file_change()

        t = threading.Thread(target=self.load_file_headers)
        t.daemon = True
        t.start()

    def reset_state_after_file_change(self):
        self.file_loaded = False
        self.validation_passed = False
        self.block_generation = True

        self.available_headers = []
        for var in [
            self.col_nric, self.col_school, self.col_name, self.col_level,
            self.col_stream, self.col_race, self.col_school_check, self.col_program,
        ]:
            var.set("")

        self.original_df = None
        self.cleaned_df = None
        self.bad_row_mask = None
        self.bad_row_reasons = None
        self.total_bad_rows = 0
        self.removed_rows_audit_df = None

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
        self.race_menu.config(state=state)
        self.school_check_menu.config(state=state)
        self.program_menu.config(state=state)

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
        reset_menu(self.race_menu, self.col_race)
        reset_menu(self.school_check_menu, self.col_school_check)
        reset_menu(self.program_menu, self.col_program)

    def open_folder(self):
        file_path = self.file_path_var.get().strip()
        if not file_path or not Path(file_path).exists():
            messagebox.showwarning("Warning", "Input file not set or doesn't exist")
            return

        folder = str(Path(file_path).parent)
        try:
            if os.name == "nt":
                os.startfile(folder)
            elif sys.platform == "darwin":
                os.system(f'open "{folder}"')
            else:
                os.system(f'xdg-open "{folder}"')
        except Exception as e:
            messagebox.showerror("Error", f"Could not open folder:\n{e}")

    def clear_log(self):
        self.log_text.delete(1.0, tk.END)

    def log(self, message, level="INFO"):
        timestamp = datetime.now().strftime("%H:%M:%S")
        prefix = {"INFO": "[INFO]", "SUCCESS": "[OK]", "WARNING": "[WARN]", "ERROR": "[ERR]"}.get(level, "[.]")
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

    def open_moe_school_list(self, event=None):
        if getattr(sys, "frozen", False):
            base_dir = os.path.dirname(sys.executable)
        else:
            base_dir = os.path.dirname(os.path.abspath(__file__))

        file_path = os.path.join(base_dir, "MOE School List.xlsx")

        if not os.path.exists(file_path):
            messagebox.showerror(
                "File not found",
                "MOE School List.xlsx was not found.\n\nPlease ensure the app was installed with the school list file."
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

    # ---------------- Core helpers ----------------
    def read_input_file(self, path: str) -> pd.DataFrame:
        p = Path(path)
        ext = p.suffix.lower()

        if ext == ".csv":
            return pd.read_csv(path)
        if ext in [".xlsx", ".xls"]:
            return pd.read_excel(path)
        raise ValueError("Unsupported file type. Please use .xlsx, .xls, or .csv")

    @staticmethod
    def safe_str(x) -> str:
        if pd.isna(x):
            return ""
        return str(x)

    @staticmethod
    def normalize_text(x: str) -> str:
        s = str(x).strip().upper()
        s = re.sub(r"\s+", " ", s)
        return s

    @staticmethod
    def nric_is_valid(nric_value: str) -> bool:
        s = str(nric_value).strip().upper()
        if len(s) != 9:
            return False
        return bool(re.match(r"^[STFGM]\d{7}[A-Z]$", s))

    def clean_export_text(self, series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.replace("\u00A0", " ", regex=False)
            .str.replace(r"\s+", " ", regex=True)
            .str.strip()
            .str.upper()
        )

    def clean_nric_series(self, series: pd.Series) -> pd.Series:
        return (
            series.astype(str)
            .str.upper()
            .str.strip()
            .str.replace(r"\s+", "", regex=True)
        )

    def school_check_is_true(self, value: str) -> bool:
        return self.normalize_text(value) == "TRUE"

    def race_is_malay(self, value: str) -> bool:
        return self.normalize_text(value) == "MALAY"

    def infer_group(self, level_value: str, stream_value: str, program_value: str):
        """
        Strict standardised grouping only.
        No fallback spelling support beyond trim + uppercase.
        """
        lvl = self.normalize_text(level_value)
        stm = self.normalize_text(stream_value)
        prog = self.normalize_text(program_value)

        if lvl == "P6" and prog == "MHC":
            return "MHC"
        if lvl == "P6" and prog == "SIPMS":
            return "SIPMS"
        if lvl == "P6":
            return "PSLE"
        if lvl == "S4" and stm == "G2":
            return "NA"
        if lvl == "S4" and stm == "G1":
            return "NT"
        if lvl == "S4" and stm == "G3":
            return "EX"
        return None

    # ---------------- Load file ----------------
    def load_file_headers(self):
        self.set_busy(True, "Loading file headers...", "#f39c12")
        try:
            file_path = self.file_path_var.get().strip()
            if not file_path or not Path(file_path).exists():
                raise FileNotFoundError("Input file not found.")

            df = self.read_input_file(file_path)
            self.original_df = df
            self.cleaned_df = df.copy()

            self.available_headers = list(df.columns)
            self.populate_column_menus(self.available_headers)
            self.set_mapping_enabled(True)

            auto_map_candidates = [
                (self.col_nric, ["NRIC"]),
                (self.col_school, ["School Name", "SCHOOL NAME"]),
                (self.col_name, ["Name of Student", "STATUTORY NAME"]),
                (self.col_level, ["Level", "LEVEL"]),
                (self.col_stream, ["Stream", "STREAM"]),
                (self.col_race, ["Race", "RACE"]),
                (self.col_school_check, ["School Check", "SCHOOL CHECK"]),
                (self.col_program, ["Program", "PROGRAM"]),
            ]

            for var, choices in auto_map_candidates:
                matched = next((c for c in choices if c in self.available_headers), "")
                var.set(matched)

            self.file_loaded = True
            self.validate_btn.config(state=tk.NORMAL)
            self.remove_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.DISABLED)
            self.status_bar.config(text="File loaded. Map columns, then validate.", bg="#27ae60")

            if not all([
                self.col_nric.get(), self.col_school.get(), self.col_name.get(),
                self.col_level.get(), self.col_stream.get(), self.col_race.get(),
                self.col_school_check.get(), self.col_program.get(),
            ]):
                messagebox.showinfo(
                    "Check Mapping",
                    "File loaded.\n\nPlease confirm all 8 mapped columns before validating."
                )

        except Exception as e:
            self.log(f"ERROR loading file: {e}", "ERROR")
            self.status_bar.config(text="Error loading file.", bg="#e74c3c")
            messagebox.showerror("Error", f"Failed to load file:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            if self.file_loaded:
                self.status_bar.config(text="File loaded. Map columns, then validate.", bg="#27ae60")

    # ---------------- Validation ----------------
    def validate_dataframe(
        self,
        df: pd.DataFrame,
        nric_col: str,
        school_col: str,
        name_col: str,
        level_col: str,
        stream_col: str,
        race_col: str,
        school_check_col: str,
        program_col: str,
    ):
        required_cols = [
            nric_col, school_col, name_col, level_col,
            stream_col, race_col, school_check_col, program_col,
        ]
        missing_cols = [c for c in required_cols if c not in df.columns]

        if missing_cols:
            bad_mask = pd.Series([False] * len(df), index=df.index)
            reasons = [""] * len(df)
            return missing_cols, bad_mask, reasons, []

        bad_mask_list = []
        reasons = []
        row_messages = []

        for idx, row in df.iterrows():
            r = []

            nric = self.safe_str(row[nric_col]).strip()
            school = self.safe_str(row[school_col]).strip()
            name = self.safe_str(row[name_col]).strip()
            level = self.safe_str(row[level_col]).strip()
            stream = self.safe_str(row[stream_col]).strip()
            race = self.safe_str(row[race_col]).strip()
            school_check = self.safe_str(row[school_check_col]).strip()
            program = self.safe_str(row[program_col]).strip()

            if not self.nric_is_valid(nric):
                r.append(f"NRIC invalid (value='{nric}')")
            if len(school) > 66:
                r.append(f"SCHOOL too long ({len(school)})")
            if len(name) > 66:
                r.append(f"NAME too long ({len(name)})")

            # Strict checks only for rows that are meant to be eligible after pre-filter.
            relevant_row = self.school_check_is_true(school_check) and self.race_is_malay(race)
            if relevant_row:
                inferred = self.infer_group(level, stream, program)
                if inferred is None:
                    r.append(
                        f"LEVEL/STREAM/PROGRAM cannot map "
                        f"(LEVEL='{level}', STREAM='{stream}', PROGRAM='{program}')"
                    )

            is_bad = len(r) > 0
            bad_mask_list.append(is_bad)
            reasons.append("; ".join(r))

            if is_bad:
                excel_row_num = df.index.get_loc(idx) + 2
                school_preview = school[:60] + ("..." if len(school) > 60 else "")
                name_preview = name[:60] + ("..." if len(name) > 60 else "")
                row_messages.append(
                    f"Row {excel_row_num}: {'; '.join(r)} | "
                    f"NRIC='{nric}' | SCHOOL='{school_preview}' | NAME='{name_preview}' | "
                    f"LEVEL='{level}' | STREAM='{stream}' | RACE='{race}' | "
                    f"SCHOOL CHECK='{school_check}' | PROGRAM='{program}'"
                )

        bad_mask = pd.Series(bad_mask_list, index=df.index)
        return [], bad_mask, reasons, row_messages

    def start_validate(self):
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        if not self.file_loaded or self.cleaned_df is None:
            messagebox.showwarning("Select File First", "Please select an input file first.")
            return

        mapped = [
            self.col_nric.get().strip(),
            self.col_school.get().strip(),
            self.col_name.get().strip(),
            self.col_level.get().strip(),
            self.col_stream.get().strip(),
            self.col_race.get().strip(),
            self.col_school_check.get().strip(),
            self.col_program.get().strip(),
        ]
        if not all(mapped):
            messagebox.showerror("Missing Mapping", "Please select all 8 required columns before validating.")
            return

        t = threading.Thread(target=self.validate_file)
        t.daemon = True
        t.start()

    def validate_file(self):
        self.set_busy(True, "Validating file...", "#f39c12")

        self.bad_row_mask = None
        self.bad_row_reasons = None
        self.total_bad_rows = 0
        self.validation_passed = False
        self.block_generation = True
        self.remove_btn.config(state=tk.DISABLED)
        self.generate_btn.config(state=tk.DISABLED)
        self.removed_rows_audit_df = None

        try:
            df = self.cleaned_df.copy()

            nric_col = self.col_nric.get().strip()
            school_col = self.col_school.get().strip()
            name_col = self.col_name.get().strip()
            level_col = self.col_level.get().strip()
            stream_col = self.col_stream.get().strip()
            race_col = self.col_race.get().strip()
            school_check_col = self.col_school_check.get().strip()
            program_col = self.col_program.get().strip()

            self.log("=" * 90, "INFO")
            self.log("VALIDATION START", "INFO")
            self.log(
                f"Mapping: NRIC='{nric_col}', SCHOOL='{school_col}', NAME='{name_col}', "
                f"LEVEL='{level_col}', STREAM='{stream_col}', RACE='{race_col}', "
                f"SCHOOL CHECK='{school_check_col}', PROGRAM='{program_col}'",
                "INFO"
            )
            self.log("=" * 90, "INFO")

            missing_cols, bad_mask, reasons, row_messages = self.validate_dataframe(
                df, nric_col, school_col, name_col, level_col, stream_col,
                race_col, school_check_col, program_col
            )

            if missing_cols:
                self.log(f"Missing mapped columns: {missing_cols}", "ERROR")
                self.status_bar.config(text="Validation failed: missing mapped columns.", bg="#e74c3c")
                messagebox.showerror("Validation Failed", f"Missing columns:\n{missing_cols}")
                return

            self.bad_row_mask = bad_mask
            self.bad_row_reasons = reasons
            self.total_bad_rows = int(bad_mask.sum())

            self.log("VALIDATION COMPLETE", "INFO")
            self.log(f"Rows checked: {len(df)}", "INFO")
            self.log(f"Problematic rows: {self.total_bad_rows}", "INFO")
            self.log("=" * 90, "INFO")

            if self.total_bad_rows > 0:
                self.log("Issues found. Showing up to 10 examples:", "WARNING")
                shown = 0
                for msg in row_messages:
                    self.log(msg, "WARNING")
                    shown += 1
                    if shown >= 10:
                        remaining = self.total_bad_rows - shown
                        if remaining > 0:
                            self.log(f"... and {remaining} more problematic rows", "WARNING")
                        break

                self.validation_passed = False
                self.block_generation = True
                self.remove_btn.config(state=tk.NORMAL)
                self.generate_btn.config(state=tk.DISABLED)
                self.status_bar.config(text="Issues found. Remove or fix rows first.", bg="#f39c12")
                messagebox.showwarning(
                    "Issues Found",
                    f"Found {self.total_bad_rows} problematic rows.\n\n"
                    "You can remove them in-app or fix the file manually."
                )
                return

            self.validation_passed = True
            self.block_generation = False
            self.remove_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.NORMAL)
            self.status_bar.config(text="File clean. Ready to generate outputs.", bg="#27ae60")
            messagebox.showinfo("All Good", "The file passed validation.\nYou can generate outputs now.")

        except Exception as e:
            self.log(f"ERROR during validation: {e}", "ERROR")
            self.status_bar.config(text="Error during validation.", bg="#e74c3c")
            messagebox.showerror("Error", f"Validation failed:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            if self.validation_passed:
                self.status_bar.config(text="File clean. Ready to generate outputs.", bg="#27ae60")

    # ---------------- Remove bad rows ----------------
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
            "This will drop all problematic rows in memory and save an audit Excel file.\n\nProceed?"
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

            removed_df = df[mask].copy()
            if not removed_df.empty:
                removed_df.insert(0, "REMOVAL_REASON", [reasons[df.index.get_loc(i)] for i in removed_df.index])

            self.cleaned_df = df[~mask].copy()

            removed_total = int(mask.sum())
            self.total_bad_rows = 0
            self.validation_passed = True
            self.block_generation = False
            self.remove_btn.config(state=tk.DISABLED)
            self.generate_btn.config(state=tk.NORMAL)

            file_path = Path(self.file_path_var.get().strip())
            base_folder = file_path.parent
            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            audit_path = base_folder / f"REMOVED_ROWS_AUDIT_{ts}.xlsx"

            if removed_total > 0 and not removed_df.empty:
                with pd.ExcelWriter(audit_path, engine="openpyxl") as writer:
                    removed_df.to_excel(writer, sheet_name="Removed Rows", index=False)
                self.removed_rows_audit_df = removed_df
                self.log(f"Removed {removed_total} problematic rows.", "SUCCESS")
                self.log(f"Saved audit file: {audit_path.name}", "SUCCESS")

            self.status_bar.config(text="Rows removed. Ready to generate outputs.", bg="#27ae60")
            messagebox.showinfo("Removed", f"Removed {removed_total} problematic rows.\n\nAudit saved:\n{audit_path}")

        except Exception as e:
            self.log(f"ERROR during removal: {e}", "ERROR")
            self.status_bar.config(text="Error during removal.", bg="#e74c3c")
            messagebox.showerror("Error", f"Removal failed:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            if self.validation_passed:
                self.status_bar.config(text="Rows removed. Ready to generate outputs.", bg="#27ae60")

    # ---------------- Export helpers ----------------
    def build_export_df(self, df_subset: pd.DataFrame, nric_col: str, school_col: str, name_col: str):
        df_export = df_subset[[nric_col, school_col, name_col]].copy()
        df_export.columns = ["NRIC", "SCHOOL NAME", "STATUTORY NAME"]
        df_export = df_export.dropna().copy()

        for col in df_export.columns:
            df_export[col] = self.clean_export_text(df_export[col])

        df_export["NRIC"] = self.clean_nric_series(df_export["NRIC"])

        before = len(df_export)
        df_export = df_export.drop_duplicates(subset=["NRIC"], keep="first").copy()
        duplicates_removed = before - len(df_export)
        return df_export, duplicates_removed

    @staticmethod
    def fixed_width_format(df_export: pd.DataFrame):
        lines = []
        warnings = []

        for index, row in df_export.iterrows():
            nric = str(row.get("NRIC", "")).strip().upper()
            school = str(row.get("SCHOOL NAME", "")).strip().upper()
            statutory = str(row.get("STATUTORY NAME", "")).strip().upper()

            nric_bad = False
            if len(nric) != 9:
                warnings.append(f"Row {index + 1}: NRIC must be exactly 9 characters - '{nric}'")
                nric_bad = True
            elif nric[0] not in {"S", "T", "F", "G", "M"}:
                warnings.append(f"Row {index + 1}: NRIC must start with S/T/F/G/M - '{nric}'")
                nric_bad = True
            elif not nric[1:8].isdigit():
                warnings.append(f"Row {index + 1}: NRIC middle 7 characters must be digits - '{nric}'")
                nric_bad = True
            elif not nric[8].isalpha():
                warnings.append(f"Row {index + 1}: NRIC must end with a letter - '{nric}'")
                nric_bad = True

            if nric_bad:
                continue

            if len(school) > 66:
                warnings.append(f"Row {index + 1}: SCHOOL NAME exceeds 66 characters - '{school}' ({len(school)} chars)")
                school = school[:66]
            if len(statutory) > 66:
                warnings.append(f"Row {index + 1}: STATUTORY NAME exceeds 66 characters - '{statutory}' ({len(statutory)} chars)")
                statutory = statutory[:66]

            line = nric.ljust(9) + school.ljust(66) + statutory.ljust(66)
            lines.append(line)

        return "\n".join(lines), warnings

    # ---------------- Generate outputs ----------------
    def start_generate_outputs(self):
        if self.processing:
            messagebox.showinfo("Info", "Processing already in progress")
            return
        if not self.validation_passed or self.block_generation:
            messagebox.showwarning(
                "Generation Blocked",
                "Data is not clean yet.\n\nPlease validate and fix/remove problematic rows first."
            )
            return
        if self.cleaned_df is None or self.cleaned_df.empty:
            messagebox.showwarning("No Data", "No rows available to export.")
            return

        t = threading.Thread(target=self.generate_outputs)
        t.daemon = True
        t.start()

    def generate_outputs(self):
        self.set_busy(True, "Generating Excel + TXT outputs...", "#f39c12")
        try:
            file_path = Path(self.file_path_var.get().strip())
            base_folder = file_path.parent

            ts = datetime.now().strftime("%Y%m%d_%H%M%S")
            out_folder = base_folder / f"MOE_OUTPUT_{ts}"
            out_folder.mkdir(parents=True, exist_ok=True)

            nric_col = self.col_nric.get().strip()
            school_col = self.col_school.get().strip()
            name_col = self.col_name.get().strip()
            level_col = self.col_level.get().strip()
            stream_col = self.col_stream.get().strip()
            race_col = self.col_race.get().strip()
            school_check_col = self.col_school_check.get().strip()
            program_col = self.col_program.get().strip()

            self.log("=" * 90, "INFO")
            self.log("OUTPUT GENERATION START", "INFO")
            self.log(f"Output folder: {out_folder}", "INFO")
            self.log("=" * 90, "INFO")

            df = self.cleaned_df.copy()

            # Notebook pre-filter logic
            df = df[
                df[school_check_col].apply(self.school_check_is_true) &
                df[race_col].apply(self.race_is_malay)
            ].copy()

            self.log(f"Rows after SCHOOL CHECK + RACE filter: {len(df)}", "INFO")

            grouped_raw = {
                "PSLE": df[df.apply(lambda r: self.infer_group(r[level_col], r[stream_col], r[program_col]) == "PSLE", axis=1)].copy(),
                "NA": df[df.apply(lambda r: self.infer_group(r[level_col], r[stream_col], r[program_col]) == "NA", axis=1)].copy(),
                "NT": df[df.apply(lambda r: self.infer_group(r[level_col], r[stream_col], r[program_col]) == "NT", axis=1)].copy(),
                "EX": df[df.apply(lambda r: self.infer_group(r[level_col], r[stream_col], r[program_col]) == "EX", axis=1)].copy(),
                "MHC": df[df.apply(lambda r: self.infer_group(r[level_col], r[stream_col], r[program_col]) == "MHC", axis=1)].copy(),
                "SIPMS": df[df.apply(lambda r: self.infer_group(r[level_col], r[stream_col], r[program_col]) == "SIPMS", axis=1)].copy(),
            }

            all_school_names = []
            files_written = 0

            for key, raw_df in grouped_raw.items():
                if raw_df.empty:
                    self.log(f"Skipping {key}: no rows", "WARNING")
                    continue

                df_export, duplicates_removed = self.build_export_df(raw_df, nric_col, school_col, name_col)
                output_base_name = self.output_names[key]

                excel_out = out_folder / f"{output_base_name}.xlsx"
                txt_out = out_folder / f"{output_base_name}.txt"

                df_export.to_excel(excel_out, index=False)

                formatted_text, warnings = self.fixed_width_format(df_export)
                txt_out.write_text(formatted_text, encoding="utf-8")

                all_school_names.extend(df_export["SCHOOL NAME"].tolist())

                self.log(
                    f"Saved {output_base_name}: rows={len(df_export)}, duplicate NRIC removed={duplicates_removed}",
                    "SUCCESS"
                )
                if warnings:
                    self.log(f"{output_base_name}: {len(warnings)} formatting warning(s)", "WARNING")

                files_written += 2

            if all_school_names:
                unique_schools = sorted(set(all_school_names))
                df_schools = pd.DataFrame(unique_schools, columns=["SCHOOL NAME"])
                schools_path = out_folder / "ALL_SCHOOLS.xlsx"
                df_schools.to_excel(schools_path, index=False)
                self.log(f"Saved ALL_SCHOOLS.xlsx ({len(df_schools)} schools)", "SUCCESS")
                files_written += 1

            self.log("=" * 90, "INFO")
            self.log("OUTPUT GENERATION COMPLETE", "SUCCESS")
            self.log(f"Files written: {files_written}", "INFO")
            self.log("=" * 90, "INFO")

            self.status_bar.config(text=f"Done. Wrote {files_written} files.", bg="#27ae60")
            messagebox.showinfo(
                "Done",
                f"Generation complete.\n\nFiles written: {files_written}\n\nOutput folder:\n{out_folder}"
            )

        except Exception as e:
            self.log(f"ERROR during output generation: {e}", "ERROR")
            self.status_bar.config(text="Error during output generation.", bg="#e74c3c")
            messagebox.showerror("Error", f"Generation failed:\n\n{e}")
        finally:
            self.processing = False
            self.set_busy(False, "Ready", "#ecf0f1")
            self.status_bar.config(text="Ready", bg="#ecf0f1")


def main():
    root = tk.Tk()
    app = ExcelToMOEOutputsApp(root)
    root.mainloop()


if __name__ == "__main__":
    main()
