![JTA](https://dam.mediacorp.sg/image/upload/s--e7I8o5qw--/c_fill,g_auto,h_676,w_1200/f_auto,q_auto/v1/tdy-migration/file6rv1xtrdjpy1347083k9.jpg?itok=DWUIJt8g)

# Organisation Data Processor App

A desktop tool that automatically cleans, validates, and formats student records for reporting submissions.

**Download:** [Latest Release](https://github.com/amirulshafiq98/Work-Stuff/releases/download/Data_Cleaning/Data.Processor.App.zip)

---

## What This Tool Does

This tool takes messy student data from multiple sources and turns it into clean, submission-ready files. It saves hours of manual work by:

- ✅ Automatically merging data from different systems
- ✅ Checking and removing invalid records
- ✅ Sorting students into the correct reporting categories
- ✅ Creating both Excel files (for review) and fixed-width text files
- ✅ Keeping a record of what was changed or removed

**Why this matters:** Manual data cleaning is slow, error-prone, and hard to verify. This tool makes the process repeatable, fast, and creates an audit trail.

---

## Who Should Use This

- Staff who prepare student data for submissions
- Anyone who needs to process Organisation student records
- Non-technical users (the tool has a simple point-and-click interface)

---

## How to Use It

![Picture1](https://github.com/user-attachments/assets/7aac9e6a-e90c-498a-8da8-d04d82641c56)

### Step 1: Download and Extract
Download the ZIP file from the link above and extract it to a folder on your computer.

### Step 2: Run the Program
Double-click the program file to open it. You'll see a window with two options:

**Option 1: Process New Data**
- Use this when you have fresh raw data files
- Click "Select Files" and choose your input files (CSV and Excel)
- Choose which cleaning rules to apply (checkboxes)
- Click "Process" and wait for it to finish
- Your clean files will appear in the output folder

**Option 2: Regenerate TXT Files**
- Use this when you've already processed data but made manual edits to the Excel files
- Select the Excel files you edited
- Click "Generate TXT" to create new submission files
- This saves you from re-processing everything from scratch

### Step 3: Review the Outputs
The tool creates several files:
- **Excel files** (one for each student category) - for human review
- **TXT files** (matching each Excel file) - for ORG_2 submission
- **Log file** - shows what the tool did
- **Removed records file** - lists any students that were excluded and why

---

## What Data This Tool Processes

### Input Files Required
You need two files:
1. **Internal attendance file** (CSV format) - from your internal system
2. **Salesforce export** (Excel format) - from your CRM

### Student Categories
The tool automatically sorts students into four groups based on their level and stream:
- PSLE (Primary 6)
- Secondary 4 Normal Academic (NA)
- Secondary 4 Normal Technical (NT)  
- Secondary 4 Express (EX)

### Data Cleaning Rules
You can turn these on or off:
- **Remove invalid NRICs** - Removes students with incorrectly formatted ID numbers
- **Remove duplicates** - Keeps only one record per student
- **Clean names** - Removes weird characters from student names and makes them uppercase

---

## Output Files Explained

### Excel Files
One Excel file per student category, named like:
- `ORG_MTSCTP PSLE.xlsx`
- `ORG_MTSCTP SEC 4 NA.xlsx`
- `ORG_MTSCTP SEC 4 NT.xlsx`
- `ORG_MTSCTP SEC 4 EX.xlsx`

These files are for human review. You can open them, check the data, and make manual corrections if needed.

### TXT Files (Fixed-Width Format)

<img width="390" height="148" alt="Picture3" src="https://github.com/user-attachments/assets/588f9c59-37e4-4146-a34c-3b50012980ab" />

One TXT file per student category (matching the Excel files). These are formatted exactly how ORG_2 wants them:
- NRIC: 9 characters wide
- School Name: 66 characters wide  
- Student Name: 66 characters wide

**Important:** If a school name or student name is too long, the tool will automatically cut it to fit and write a warning in the log file.

### Log Files
- `Processing_Log_<timestamp>.txt` - Shows everything the tool did
- `Removed_Records_<timestamp>.xlsx` - Lists students who were excluded
- `Removed_Records_<timestamp>.txt` - Same as above, in text format

These files help you verify what changed and why.

---

## Common Workflows

### First Time Processing
1. Get your two input files (attendance CSV + CRM Excel)
2. Open the tool and choose "Process New Data"
3. Select both input files
4. Turn on the cleaning rules you want
5. Click Process
6. Review the Excel outputs
7. Submit the TXT files to ORG_2

### After Making Manual Edits

![Picture2](https://github.com/user-attachments/assets/dcbc134b-0879-43e6-af14-e5d637b0d67a)

Sometimes you need to fix specific records by hand (like correcting a school name spelling). Here's how:
1. Open the Excel file from the previous processing
2. Make your corrections directly in Excel
3. Save the Excel file
4. Open the tool and choose "Regenerate TXT Files"
5. Select your edited Excel files
6. Click Generate TXT
7. The tool creates fresh TXT files without re-running all the validation

This saves time and preserves your manual fixes.

---

## Technical Details

### Built With
- Python (programming language)
- Tkinter (creates the window interface)
- Pandas (handles spreadsheet data)

### File Merging Logic
The tool matches records using:
- **Internal system:** Uses the `Student Ref No.` field
- **CRM system:** Uses the `Registration ID` field

Both must exist for a record to be included.

### Known Limitations
- Input files must have the expected column names (tool will show an error if they're missing)
- Long school/student names get automatically shortened to fit the fixed-width format (check the log file for warnings)
- School name spelling is NOT checked automatically - you must verify this manually before submission

---

## Future Improvements

Things that could make this tool even better:
- Add a school name checker (auto-correct common spelling variations)
- Show a preview of the data before saving files
- Add a summary dashboard showing how many records were processed/removed
- Package as a standalone .exe file that doesn't need Python installed

---

## Privacy & Security Note

This repository does not include:
- Real student data
- Sample input files  
- Example outputs

The underlying data represents sensitive student information and is not publicly shared.

---

## Questions or Issues?

If the tool isn't working as expected:
1. Check that your input files have the correct column names
2. Review the log file for error messages
3. Check the removed records file to see if your data was excluded by a validation rule

---

## License

_[Add your license here if applicable]_
