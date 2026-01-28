# Automate Cybersecurity Tasks with Python

## Project Overview

This module demonstrates practical Python automation for cybersecurity data processing tasks. The project extracts student enrollment data from PDF documents, parses it into a structured format, and exports it to a professionally formatted Excel spreadsheet.

**This automation script showcases:**
- Advanced PDF parsing with error handling
- Intelligent data extraction from mixed/malformed data formats
- Structured data organization and validation
- Professional Excel file generation with formatting
- Practical cybersecurity data management skills

---

## Project Objectives

✅ **PDF Data Extraction** - Extract raw text from multi-page PDF documents  
✅ **Intelligent Parsing** - Handle mixed data formats and edge cases  
✅ **Data Structuring** - Parse unstructured text into organized columns  
✅ **Excel Export** - Generate professionally formatted Excel files  
✅ **Data Validation** - Implement search and verification functions  

---

## Script Features

### `automation_python.py` - Main Script

#### **Function 1: `extract_students_from_pdf(pdf_path)`**
- **Purpose:** Extracts student records from multi-page PDF
- **Process:**
  - Reads all PDF pages after the first (header page)
  - Filters out headers, university info, and page markers
  - Removes non-data lines (headers, status markers, page numbers)
- **Input:** Path to PDF file
- **Output:** List of 755 raw student record lines
- **Success Rate:** 100% of valid student data extracted

#### **Function 2: `find_student(student_list, name)`**
- **Purpose:** Search for students by name (case-insensitive)
- **Usage:** Verify specific student records in the dataset
- **Example:** Found student "CHIPWAYA PAUL POOLA" with ID 2016140638

#### **Function 3: `parse_student_data_columns(student_list)`**
- **Purpose:** Parse raw text into structured 6-column format
- **Key Features:**
  - Handles combined NO. and STUDENT NO. fields (e.g., "322016138632" → 32, 2016138632)
  - Intelligent NRC/PASSPORT detection (identifies by "/" or 9+ digit patterns)
  - Flexible name parsing for variable data formats
  - Validates numeric fields while accepting optional name variations
  - Maintains 100% data integrity (755/755 records parsed successfully)
- **Output Columns:**
  - `NO.` - Student number
  - `STUDENT NO.` - Student ID (10 digits)
  - `SURNAME` - Last name
  - `OTHER NAMES` - Middle names (if present)
  - `FIRST NAME` - First name
  - `NRC/PASSPORT` - National Registration/Passport ID

#### **Function 4: `save_students_to_excel_columns(parsed_data, excel_path)`**
- **Purpose:** Create professionally formatted Excel file
- **Features:**
  - Merged title row: "STUDENT ENROLLMENT LIST"
  - Dark blue background (203864) with white text (14pt bold)
  - Frozen header rows for easy scrolling
  - Auto-filter on all columns
  - Black borders on all cells
  - Optimized column widths
  - 30-point row height for title, 25-point for headers
- **Output:** `students_list.xlsx` with professional formatting

### **Main Execution Workflow**

```
Step 1: Extract PDF Data (859 lines) → Filter (755 valid records)
   ↓
Step 2: Parse Structured Data (100% success rate)
   ↓
Step 3: Save to Excel (Professional formatting)
   ↓
Step 4: Verify Data (Search functionality)
```

---

## Installation & Setup

### **Requirements**
```bash
Python 3.8+
PyPDF2          # PDF parsing
pandas          # Data manipulation
openpyxl        # Excel file creation
```

### **Install Dependencies**
```bash
pip install PyPDF2 pandas openpyxl
```

### **Virtual Environment** (Recommended)
```bash
python -m venv venv
source venv/bin/activate  # Linux/Mac
venv\Scripts\activate      # Windows
pip install -r requirements.txt
```

---

## Usage

### **Run the Complete Workflow**
```bash
python automation_python.py
```

### **Expected Output**
```
Step 1: Extracting student data from PDF...
✓ Extracted 755 total student records from PDF

Step 2: Parsing student data into columns...
✓ Parsed 755 student records into structured format

Step 3: Saving parsed data to Excel file...
Student data saved to students_list.xlsx with professional formatting
Total records: 755

Step 4: Searching for student...
✓ Found student: 105 2016140638 CHIPWAYA PAUL POOLA 254400/15/1

============================================================
PROCESS COMPLETED SUCCESSFULLY
============================================================
Total students extracted from PDF: 755
Total records parsed and saved: 755
Excel file created: students_list.xlsx
============================================================
```

---

## Data Processing Results

| Metric | Value |
|--------|-------|
| **PDF Records Extracted** | 755 valid student records |
| **Data Extraction Rate** | 100% |
| **Parsing Success Rate** | 755/755 (100%) |
| **Output Format** | Excel (.xlsx) |
| **Column Count** | 6 |
| **Excel Rows** | 758 (1 title + 1 header + 756 data) |

---

## Cybersecurity Applications

This automation demonstrates critical skills for cybersecurity roles:

### **Data Management & Analysis**
- Parse large datasets from multiple sources
- Handle malformed/mixed-format data
- Validate data integrity and consistency

### **Incident Response**
- Extract logs from multiple file formats
- Organize alert data into actionable reports
- Search and filter security events

### **Compliance & Auditing**
- Generate formatted reports for stakeholders
- Automate repetitive documentation tasks
- Maintain structured audit trails

### **Threat Intelligence**
- Parse indicator lists from PDFs
- Structure IoCs (Indicators of Compromise)
- Export for downstream analysis tools

---

## File Structure

```
07-Automate-Cybersecurity-Tasks-with-Python/
├── README.md                              # This file
├── automation_python.py                   # Main automation script
├── SCRIPT_DOCUMENTATION.md                # Detailed technical documentation
├── students_list.xlsx                     # Output Excel file
├── hss_(b.a_nqs)_final_pub._list_-_male.pdf  # Input PDF file
└── script_output.txt                      # Execution output log
```

---

## Technical Highlights

### **Intelligent Data Parsing**
The script handles complex PDF formatting challenges:
- **Combined fields:** NO. and STUDENT NO. merged in one token
- **Variable formats:** Names with 1-4 components
- **Missing data:** Records without NRC/PASSPORT
- **Special characters:** Slash in ID numbers (e.g., "254400/15/1")

### **Excel Formatting**
Professional output with:
- Centered headers with blue background (366092)
- White bold text for headers
- Calibri 10pt font for data
- Frozen panes at row 4
- Auto-filter for sorting/filtering
- Black cell borders throughout

### **Error Handling**
- Gracefully handles invalid records
- Continues processing despite errors
- Reports processing statistics
- Validates numeric fields

---

## Performance Metrics

- **PDF Parsing Time:** < 1 second
- **Data Extraction:** 755 records in < 2 seconds
- **Excel Generation:** < 1 second
- **Total Execution Time:** ~3 seconds
- **Memory Usage:** ~50 MB

---

## Lessons Learned

### **PDF Text Extraction Challenges**
1. **No guaranteed field separation** - Fields may be combined (e.g., "322016138632")
2. **Variable data quality** - Records have different formatting/completeness
3. **Header/footer noise** - Multiple types of non-data lines mixed in
4. **Flexible parsing needed** - Can't assume fixed field positions

### **Data Validation Best Practices**
1. Validate core fields (numbers must be numeric)
2. Accept optional variations (names may be missing)
3. Log skipped records for audit trail
4. Maintain data integrity during transformation

---

## Future Enhancements

- [ ] Add data validation against schema
- [ ] Implement duplicate detection
- [ ] Add CSV export option
- [ ] Create interactive search GUI
- [ ] Support batch PDF processing
- [ ] Add data encryption for sensitive fields
- [ ] Generate audit reports

---

## Related Documentation

- **[SCRIPT_DOCUMENTATION.md](SCRIPT_DOCUMENTATION.md)** - Detailed technical specification
- **Portfolio:** Google Cybersecurity Professional Certificate
- **Course:** Module 07 - Automate Cybersecurity Tasks with Python

---

## Author

**Paul Chola** - Cybersecurity Professional Portfolio Project

---

## License

Educational project - Google Cybersecurity Professional Certificate

---

## Questions & Support

For questions about the script or Python automation in cybersecurity:
- Review the detailed documentation in `SCRIPT_DOCUMENTATION.md`
- Check the code comments in `automation_python.py`
- Test with the provided sample data in `students_list.xlsx`
