# KID Document Generator v1.2

## **Overview**
KID Document Generator is a Python application that generates Key Information Documents (KIDs) for contractors. It uses contractor data from a `.csv` file and a customizable KID template to populate and produce PDF documents for each contractor. The program provides a clean graphical user interface (GUI) to guide users through selecting files, monitoring progress, and viewing the status of document generation.

---

## **Features**
- GUI-based interface for ease of use.
- Supports `.csv` data files for bulk document generation.
- Customizable KID templates.
- Automatic calculation of basic deductions (tax, National Insurance Contributions, umbrella fees).
- Real-time progress bar to display generation progress.
- Summary dialog box showing the number of successful document generations.

---

## **Development Information**
**Version:** 1.1
**Date of Development:** 12 February 2025
**Developer:** Peter Grint

---

## **Installation and Requirements**

### **Prerequisites:**
- Python 3.7 or higher
- Required Python packages:
  - `tkinter`
  - `pandas`
  - `fpdf`

### **Installing Dependencies:**
To install the necessary dependencies, run:
```bash
pip install pandas fpdf
```

---

## **How to Use**

### **1. Run the Program:**
Execute the script `kid_generator_gui.py` using Python:
python kid_generator_gui.py


### **2. Steps to Generate KID Documents:**
1. **Select Data Input File (.csv):** Choose a CSV file containing contractor data.
2. **Select KID Template:** Select a text file containing the KID template with placeholders.
3. **Set Save Destination:** Choose a directory to save the generated KID documents.
4. **Click 'Generate KID Documents':** The program will create a KID document for each row in the CSV file and save it as a PDF.

### **3. CSV Format:**
Ensure the CSV file has the following headers:
- `CandidateFirstName`
- `CandidateLastName`
- `ContractStartDate`
- `ContractEndDate`
- `DailyRate`
- `UmbrellaFee`

**Sample CSV:**
```csv
CandidateFirstName,CandidateLastName,ContractStartDate,ContractEndDate,DailyRate,UmbrellaFee
Joe,Bloggs,01/01/2025,31/03/2025,500.0,25.0
Jane,Doe,01/02/2025,30/06/2025,450.0,20.0
```

### **4. KID Template Format:**
The KID template should be a plain text file with placeholders. Example placeholders:
- `{CandidateFirstName}`
- `{CandidateLastName}`
- `{ContractStartDate}`
- `{ContractEndDate}`
- `{DailyRate}`
- `{TaxDeduction}`
- `{NICDeduction}`
- `{UmbrellaFee}`
- `{NetPay}`

Example template content:
```
Dear {CandidateFirstName} {CandidateLastName},

Your contract details are as follows:
Start Date: {ContractStartDate}
End Date: {ContractEndDate}
Daily Rate: {DailyRate}
Tax Deduction: {TaxDeduction}
NIC Deduction: {NICDeduction}
Umbrella Fee: {UmbrellaFee}
Net Pay: {NetPay}
```

---

## **Output File Naming Convention**
The generated KID files will be saved with the following format:
```
CandidateFirstName_CandidateLastName_-_ContractStartDate_-_ContractEndDate_KID.pdf
```
For example:
```
Joe_Bloggs_-_01012025_-_31032025_KID.pdf
```

---

## **Known Limitations**
- Only static deductions are calculated (20% tax, 12% NIC). Custom deductions must be added manually.
- Assumes the CSV file contains all required headers.

---

## **Updates v1.1**

- Added "Skip Row" functionality for default CSV export format from CRM
- Accounted for Candidate name being one value, not 2 values (First, Last)
- Collect currency, populate template with palceholder to allow for GBP or EUR
- Collect Job Title, added placeholder to template
- Collect Placement ID and use in file name of saved docs
- Accounted for pension contribution at 3%
- Fixed bug where program expected .txt template file, resolved to .docx
- BugFix: resolved disconnect between "input_file_path" and "Data_file_path" - resolved to latter
- Added dynamic deduction rates
- Changed Placeholder names to match default output from CRM

## **Updates v1.2**

- Removed "Skip Row" function when reading CSV as get.row was grabbing column headers not data
- Point above requires user to delete ROW 2 from the CSV before it is imported to this program



---

## **Future Enhancements**
- Add more robust error handling for missing or invalid data.
- Support additional file formats for data input and templates.
- Provide example .csv file for users
- Provide example Template KID FORM for users
- Bug where get.row pulls column header and not cell value, fix in v1.2

---

## **License**
Copyright (c) [2025] [RainyGrinch]

Permission is hereby granted, free of charge, to any person obtaining a copy of this software
and associated documentation files (the "Software"), to use, copy, modify, and
distribute the Software for personal, educational, or non-profit purposes,
subject to the following conditions:

Non-Commercial Clause:
The Software may not be used for any commercial purpose, including but not limited to
selling, offering as a service, or any use that generates direct or indirect profit.

Retention of Copyright Notice:
All copies or substantial portions of the Software must include this copyright notice
and this license.

Disclaimer of Warranty:
The Software is provided "as is," without warranty of any kind, express or implied.
In no event shall the author be liable for any damages arising from the
use of this Software.

---

## **Contact**
For questions or support, contact Peter Grint.

