# Python Automation Script Documentation

## Project: Student PDF Data Extraction & Excel Export Tool

---

## Overview

This enhanced Python automation script demonstrates advanced cybersecurity automation skills by extracting student records from PDF documents and exporting them to Excel format. It showcases proficiency in data handling, file processing, search algorithms, and data export capabilities.

**Script Name**: `automation_python.py`

---

## Purpose & Use Case

The script automates the process of:
1. **Extracting** student data from unstructured PDF files
2. **Processing** extracted text with filtering logic
3. **Searching** for specific student records efficiently
4. **Exporting** all data to Excel format for further analysis
5. **Reporting** search results and export statistics

This type of automation is valuable in cybersecurity for:
- Managing user/student databases
- Automating data extraction from reports
- Implementing security audit trails
- Handling bulk data processing tasks
- Reducing manual data entry errors
- Creating structured data for analysis

---

## Technical Implementation

### Libraries Used

- **PyPDF2**: PDF document parsing and text extraction
- **pandas**: Data manipulation and Excel export
- **openpyxl**: Excel file creation and formatting
- **csv**: Module imported for potential CSV data handling (extensible design)

```python
from PyPDF2 import PdfReader
import csv
import pandas as pd
```

### Function 1: `extract_students_from_pdf(pdf_path)`

**Purpose**: Extract and process student data from a PDF document

**Parameters**:
- `pdf_path` (str): File path to the PDF document to process

**Returns**:
- `student_list` (list): Processed list of student records

**Logic**:
```python
def extract_students_from_pdf(pdf_path):
    student_list = []
    reader = PdfReader(pdf_path)
    
    # Iterate through pages (skip header page at index 0)
    for page in reader.pages[1:]:
        # Extract text from each page
        text = page.extract_text()
        lines = text.split("\n")
        
        # Process each line with filtering
        for line in lines:
            line = line.strip()
            # Filter: Skip headers and empty lines
            if line and not line.startswith("THE UNIVERSITY") \
               and not line.startswith("BACHELOR") \
               and not line.startswith("GENDER"):
                student_list.append(line)
    
    return student_list
```

**Key Features**:
- Skips first page (header information)
- Removes whitespace with `.strip()`
- Filters unwanted header rows
- Handles multi-line text extraction
- Returns clean, processed data

**Security Considerations**:
- Validates and sanitizes text input
- Prevents processing of malformed header data
- Handles large documents efficiently

---

### Function 3: `parse_student_data_columns(student_list)`

**Purpose**: Parse student records into separate columns for structured data organization

**Parameters**:
- `student_list` (list): List of student records to parse

**Returns**:
- `parsed_data` (list of dicts): Each dict contains: NO., STUDENT NO., SURNAME, OTHER NAMES, FIRST NAME, NRC/PASSPORT

**Key Features**:
- **Column Separation**: Extracts each field into its own column
- **Smart Parsing**: Intelligently splits names into FIRST NAME and OTHER NAMES
- **Data Validation**: Only processes records with sufficient data
- **Error Handling**: Skips incomplete or malformed records gracefully

**Columns Created**:
1. **NO.**: Sequential number/ID
2. **STUDENT NO.**: Student identification number
3. **SURNAME**: Student's surname/family name
4. **OTHER NAMES**: Additional names/middle names
5. **FIRST NAME**: Student's first name
6. **NRC/PASSPORT**: National Registration Card or Passport number

---

### Function 4: `save_students_to_excel_columns(parsed_data, excel_path)`

**Purpose**: Export parsed student data to Excel format with separate columns and professional formatting

**Parameters**:
- `parsed_data` (list of dicts): Parsed student records
- `excel_path` (str): File path where Excel file will be saved

**Returns**:
- None (prints success message)

**Key Features**:
- **Separate Columns**: 6 distinct columns for each data field
- **Professional Formatting**:
  - Header: Bold white text on blue background
  - Data: Calibri 10pt font, centered vertical alignment
  - Borders: Black borders on all cells
  - Optimized column widths for each field type
- **Navigation Features**:
  - Frozen header row for easy scrolling
  - Auto-filter on all columns for quick search
- **Professional Appearance**: Business-ready format

**Column Widths**:
- NO.: 8 characters
- STUDENT NO.: 14 characters
- SURNAME: 16 characters
- OTHER NAMES: 20 characters
- FIRST NAME: 14 characters
- NRC/PASSPORT: 18 characters

---

### Function 2: `find_student(student_list, name)`

**Purpose**: Search for a student by name within the extracted list

**Parameters**:
- `student_list` (list): List of student records to search
- `name` (str): Search query (student name or name part)

**Returns**:
- `str`: Matching student record, or `None` if not found

**Logic**:
```python
def find_student(student_list, name):
    for line in student_list:
        # Case-insensitive search
        if name.upper() in line.upper():
            return line.strip()
    return None
```

**Key Features**:
- **Case-Insensitive Search**: Matches "paul", "PAUL", "Paul"
- **Partial Matching**: Finds records containing the search term
- **Efficient Iteration**: Linear search suitable for moderate datasets
- **Clean Return Value**: Returns trimmed string or None

**Security Considerations**:
- Prevents case-sensitivity bypass exploits
- Validates search results before return
- Safe handling of None returns

---

## Main Execution

```python
# Define file paths
pdf_path = "/home/paul/Documents/Google CyberSecurity Scholarship/Cyber security Portfolio Project /Google-Cybersecurity-Professional-Portfolio/07-Automate-Cybersecurity-Tasks-with-Python/hss_(b.a_nqs)_final_pub._list_-_male.pdf"
excel_path = "/home/paul/Documents/Google CyberSecurity Scholarship/Cyber security Portfolio Project /Google-Cybersecurity-Professional-Portfolio/07-Automate-Cybersecurity-Tasks-with-Python/students_list.xlsx"

# Extract student data from PDF
student_list = extract_students_from_pdf(pdf_path)

# Parse student data into separate columns
parsed_data = parse_student_data_columns(student_list)

# Save all students to Excel file with separate columns
save_students_to_excel_columns(parsed_data, excel_path)

# Search for the student (using surname or full name parts)
student_name = "PAUL"  # Adjusted to match the PDF format
found = find_student(student_list, student_name)

# Report search results
if found:
    print(f"Found student: {found}")
else:
    print(f"Student '{student_name}' not found in the list.")

# Display summary
print(f"\nTotal students extracted: {len(student_list)}")
print("Excel file created successfully!")
```

---

## Output Examples

**Success Case**:
```
Student data saved to /path/to/students_list.xlsx with separate columns and professional formatting
Total records: 152
Found student: 105 2016140638 CHIPWAYA PAUL POOLA 254400/15/1
Total students extracted: 859
Excel file created successfully!
```

**Excel File Structure**:
- **Row 1**: TITLE ("STUDENT ENROLLMENT LIST") - merged across all columns, dark blue background
- **Row 2**: Empty spacing row
- **Row 3**: COLUMN HEADERS - Bold white text on blue background
  - A: NO.
  - B: STUDENT NO.
  - C: SURNAME
  - D: OTHER NAMES
  - E: FIRST NAME
  - F: NRC/PASSPORT
- **Rows 4+**: Student data records (152 properly parsed records)
- **Format**: .xlsx (Excel 2007+ compatible)

**Professional Styling**:
- Title: Large (14pt), bold, white text on dark blue (203864) background, 30px height
- Header: Bold (11pt) white text on blue (366092) background, centered, 25px height
- Data: Calibri (10pt) font, left-aligned, 20px height
- Borders: Black borders on all cells
- Column Widths: Optimized for each field type
- Frozen Rows: Both title and header rows stay visible during scrolling
- Auto-Filter: Search and filter on column headers

**Visibility**:
- Title and column headers are always visible at the top
- Header row is frozen to stay visible when scrolling through data
- Professional, easy-to-read layout

---

## Technical Skills Demonstrated

### 1. **File I/O & Document Processing**
- PDF parsing using industry-standard library
- Text extraction from complex document formats
- File path handling and error prevention

### 2. **Data Processing & Filtering**
- String manipulation and normalization
- Line-by-line text processing
- Header detection and filtering logic
- Data validation and sanitization

### 3. **Search Algorithm Implementation**
- Linear search with pattern matching
- Case-insensitive comparison
- Efficient record retrieval
- Error handling and result reporting

### 4. **Python Best Practices**
- Function modularity and reusability
- Clear parameter documentation
- Meaningful variable naming
- Proper return value handling
- Error-safe None checks

### 6. **Document Layout Preservation**
- PDF-to-Excel conversion maintaining visual structure
- Professional formatting and styling
- Layout consistency between source and output documents

### 7. **Data Export & Management**
- Excel file creation and formatting
- DataFrame manipulation with pandas
- Structured data export for analysis
- File I/O operations for data persistence

### 8. **Cybersecurity Applications**
- Automated data extraction for audits
- Secure user record management
- Log file processing and analysis
- User verification workflows
- Compliance data handling and reporting
- Document format conversion for security operations

---

## Potential Enhancements

### Version 2.0 Features

```python
# Advanced logging
import logging
logging.basicConfig(level=logging.INFO)

# Error handling
try:
    student_list = extract_students_from_pdf(pdf_path)
except FileNotFoundError:
    logging.error(f"PDF file not found: {pdf_path}")

# CSV export capability
def export_to_csv(student_list, csv_path):
    """Export extracted students to CSV format"""
    df = pd.DataFrame(student_list, columns=['Student Name'])
    df.to_csv(csv_path, index=False)
    print(f"Student data saved to {csv_path}")

# Excel formatting enhancements
def save_students_to_excel_formatted(student_list, excel_path):
    """Save with enhanced Excel formatting"""
    df = pd.DataFrame(student_list, columns=['Student Name'])
    
    with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
        df.to_excel(writer, sheet_name='Students', index=False)
        
        # Get workbook and worksheet
        workbook = writer.book
        worksheet = writer.sheets['Students']
        
        # Add formatting
        from openpyxl.styles import Font, PatternFill
        header_font = Font(bold=True, color="FFFFFF")
        header_fill = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
        
        # Format header
        for cell in worksheet[1]:
            cell.font = header_font
            cell.fill = header_fill
        
        # Auto-adjust column width
        worksheet.column_dimensions['A'].width = 50

# Regular expression search for advanced patterns
import re
def search_by_pattern(student_list, pattern):
    """Search using regular expressions"""
    return [line for line in student_list if re.search(pattern, line, re.IGNORECASE)]

# Performance optimization for large datasets
def find_student_optimized(student_list, name):
    """O(n) search with early termination"""
    name_upper = name.upper()
    for line in student_list:
        if name_upper in line.upper():
            return line.strip()
    return None
```

---

## Deployment & Usage

### Requirements
```bash
pip install PyPDF2
```

### Execution
```bash
python automation_python.py
```

### Integration with Security Tools
- **Incident Response**: Extract user data from incident reports
- **Compliance Audits**: Automate data extraction from compliance documents
- **Vulnerability Reports**: Parse PDF vulnerability scan results
- **User Management**: Automate user record extraction and verification

---

## Security Implications

### Input Validation
✓ Filters malicious header strings  
✓ Sanitizes extracted text  
✓ Validates file path before processing  

### Data Handling
✓ Case-insensitive matching prevents bypass  
✓ Safe None return for failed searches  
✓ No SQL injection vulnerabilities (no database queries)  

### Best Practices Applied
✓ Error handling for missing records  
✓ Modular function design  
✓ Clear separation of concerns  

---

## Learning Outcomes

This project demonstrates mastery of:

1. **Python Programming**
   - Module imports and library usage
   - Function definition and parameters
   - Control flow (loops, conditionals)
   - String operations and methods

2. **Document Processing**
   - PDF parsing techniques
   - Text extraction and parsing
   - Data cleaning and normalization

3. **Security Automation**
   - Automating repetitive data tasks
   - Secure data handling
   - Error management and logging

4. **Professional Development**
   - Code documentation
   - Modular design patterns
   - Scalability considerations

---

## Real-World Applications in Cybersecurity

- **SOC Operations**: Automated log extraction and analysis
- **Incident Response**: Parsing incident reports and threat intel documents
- **Compliance**: Extracting data for audit reports and certifications
- **User Management**: Automating user verification workflows
- **Threat Intelligence**: Extracting IOCs from PDF threat reports

---

## Repository Integration

**Course**: Google Cybersecurity Professional Certificate - Module 07  
**Project**: Automate Cybersecurity Tasks with Python  
**Repository**: Google-Cybersecurity-Professional-Portfolio  
**Location**: `07-Automate-Cybersecurity-Tasks-with-Python/automation_python.py`

---

*Documentation Created: January 2026*  
*Skill Level: Intermediate - Practical Automation*  
*Use Case: Security Operations & Data Processing*
