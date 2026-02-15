![JTA](https://dam.mediacorp.sg/image/upload/s--e7I8o5qw--/c_fill,g_auto,h_676,w_1200/f_auto,q_auto/v1/tdy-migration/file6rv1xtrdjpy1347083k9.jpg?itok=DWUIJt8g)

# Organisation Data Checker & TXT Generator (v2)

A desktop tool that validates Excel student records, splits them by level/stream, and generates submission-ready fixed-width TXT files.

**Download:**  
➡️ **Download the latest ZIP from the Releases section of this repository**  
(Extract the ZIP and run the application inside the folder)

**Or you can click [Download](https://github.com/amirulshafiq98/Work-Stuff/releases/download/v2/Data.Processor.App.zip)**

---

## What This Tool Does

<img width="954" height="496" alt="updated1" src="https://github.com/user-attachments/assets/2f8d0294-b619-4b38-8471-38c56d27137b" />

This tool takes a **single prepared Excel file**, validates the data, splits it by student level/stream into separate Excel files, and converts them into **submission-ready TXT files**.

It is designed for situations where:
- school names have already been manually verified
- data has already been cleaned upstream
- strict formatting rules must be enforced consistently
- records need to be separated by educational level/stream

The tool helps by:

- ✅ Loading a single Excel file with all student records  
- ✅ Validating NRIC format and field length rules  
- ✅ Checking that LEVEL+STREAM can be mapped to output groups (PSLE / NA / NT / EX)  
- ✅ Flagging problematic rows clearly (with row numbers)  
- ✅ Letting users decide whether to fix or remove invalid rows  
- ✅ Splitting cleaned data into separate Excel files by group  
- ✅ Generating fixed-width TXT files only when data is clean  
- ✅ Keeping the process simple and repeatable  

**Why this matters:**  
Manual checking and splitting of hundreds of records is error-prone. This tool ensures the final submission files are properly grouped and always meet the required format before submission.

---

## Who Should Use This

- Staff preparing student data for official submissions  
- Users who already have consolidated Excel outputs and need validation + splitting  
- Non-technical users (simple point-and-click interface)  

No programming knowledge is required.

---

## How to Use It

### Step 1: Download and Extract
[Download](https://github.com/amirulshafiq98/Work-Stuff/releases/download/v2/Data.Processor.App.zip) the ZIP file and extract it to a folder on your computer.

> **Important:**  
> Do not move the `.exe` out of the extracted folder. All reference files must stay together. That includes the school list file.

---

### Step 2: Prepare Your Excel File

Your Excel file should contain **all student records** with these columns:

- NRIC  
- School Name  
- Student Name  
- Level  
- Stream  

(Exact column names can be selected inside the app.)

**Example file structure:**
- `Student_Records_2026.xlsx`
- Only **<ins>1 sheet</ins>** in the Excel file

The tool will automatically split this into separate files based on Level + Stream values.

---

### Step 3: Select Excel File

1. Open the application  
2. Click **Browse...** and select your Excel file  
3. The tool will load column headers automatically  

---

### Step 4: Map Columns

Choose which columns represent:
- NRIC  
- School Name  
- Student Name  
- Level  
- Stream  

> **Note:** If your column names match exactly (e.g., "NRIC", "SCHOOL NAME", "STATUTORY NAME", "LEVEL", "STREAM"), they will be auto-selected.

---

### Step 5: Validate Excel File

Click **Validate Excel File**.

The tool checks:
- NRIC format (exactly 9 characters, valid pattern like S1234567A)  
- School name length (max 66 characters)  
- Student name length (max 66 characters)  
- Level + Stream mapping (must match PSLE / NA / NT / EXPRESS)  

If issues are found:
- they are shown clearly in the log  
- row numbers and value previews are displayed  
- splitting and TXT generation are blocked until resolved  

---

### Step 6: Fix or Remove Problem Rows

If validation fails, you have two options:

**Option A: Fix in Excel**
- Open your Excel file  
- Correct the values  
- Save  
- Re-run validation  

**Option B: Remove Problem Rows**
- Click **Remove Problem Rows**  
- The tool drops invalid rows in memory  
- An audit Excel file is saved showing what was removed and why  
- Cleaned data is automatically split into Excel files by group  
- Validation passes once all issues are removed  

---

### Step 7: Review Split Excel Files

After validation passes (or after removing rows), the tool creates:

**Folder:** `EXCEL_OUTPUT_YYYYMMDD_HHMMSS`

**Files created:**
- `ORG_MTSCTP PSLE.xlsx`
- `ORG_MTSCTP SEC 4 NA.xlsx`
- `ORG_MTSCTP SEC 4 NT.xlsx`
- `ORG_MTSCTP SEC 4 EX.xlsx`

> **Tip:** You can open these Excel files to review how records were grouped before generating final TXT files.

---

### Step 8: Generate TXT Files

Once Excel files are created:
1. Click **Generate TXT Files**  
2. The tool creates a new folder:

`TXT_OUTPUT_YYYYMMDD_HHMMSS`

3. TXT files are written inside this folder with matching names

---

## Output Files Explained

<img width="262" height="213" alt="Untitled" src="https://github.com/user-attachments/assets/7ed23ab7-d72f-494c-88dc-63ae16419cf9" />

### Excel Files (Intermediate Output)

Created in `EXCEL_OUTPUT_YYYYMMDD_HHMMSS/` folder.

One Excel file per group:
- PSLE students  
- Secondary 4 NA students  
- Secondary 4 NT students  
- Secondary 4 EX students  

These files contain the cleaned, validated data ready for review before TXT generation.

---

### TXT Files (Fixed-Width Format)

Created in `TXT_OUTPUT_YYYYMMDD_HHMMSS/` folder.

One TXT file is generated per Excel group file.

Each line is formatted as:
- **NRIC:** 9 characters  
- **School Name:** 66 characters  
- **Student Name:** 66 characters  

If a value is too long, it is **trimmed automatically**.

The output is ready for submission without further formatting.

---

### Audit File (When Rows Are Removed)

<img width="373" height="162" alt="updated2" src="https://github.com/user-attachments/assets/2e05c2fc-1361-43c2-9644-44983124385f" />

If you use the "Remove Problem Rows" option, an audit file is created:

**File:** `REMOVED_ROWS_AUDIT_YYYYMMDD_HHMMSS.xlsx`

This file shows:
- which rows were removed  
- why they were removed  
- what group they would have belonged to  

This provides a clear record of what was excluded from the final submission.

---

### Log Output

<img width="1013" height="218" alt="Screenshot 2026-02-15 115458" src="https://github.com/user-attachments/assets/0435b1f4-ef1f-46b4-9785-107474de319a" />

The log window shows:
- which file was loaded  
- which rows failed validation  
- why validation failed  
- how many Excel files were created  
- how many TXT files were generated  

This makes it easy to verify decisions before submission.

---

## Reference Files

This application includes a **local School List** file for reference.

- It can be opened directly from the app via the clickable link  
- It is provided for **manual verification only**  
- School name spelling is **not auto-corrected**  

Users are responsible for confirming school names before submission.

---

## Common Workflow

### Standard Submission Flow
1. Prepare single Excel file with all students  
2. Verify school names manually using reference list  
3. Select Excel file in the app  
4. Map columns (NRIC, School, Name, Level, Stream)  
5. Validate file  
6. Fix or remove invalid rows  
7. Review split Excel files (grouped by level/stream)  
8. Generate TXT files  
9. Submit TXT files  

---

## Level/Stream Grouping Rules

The tool automatically groups records based on Level and Stream values:

**PSLE Group:**
- Level must be exactly: **"P6"**

**Secondary Groups (Level must be exactly "S4"):**
- **NA:** Stream must be **"G2"** or **"Normal Academic"**  
- **NT:** Stream must be **"G1"** or **"Normal Technical"**  
- **EX:** Stream must be **"G3"** or **"Express"**  

> **Important:** If a row's Level + Stream cannot be mapped to any group, it will be flagged as problematic during validation.

---

## Technical Details (For Reference)

- Built with Python  
- Desktop GUI using Tkinter  
- Excel handling via Pandas / OpenPyXL  
- No internet connection required  
- No data is sent anywhere  

---

## Privacy & Security Note

This repository does **not** contain:
- real student data  
- sample Excel inputs  
- example submission files  

The tool is designed to operate on sensitive data locally without uploading or sharing any information.

---

## Previous Versions

- **App v1** – Full data processing pipeline for CCIS-based files (merge, clean, regenerate)  
  See `Old Versions/App v1/README.md`
