KID Document Generator (Version 1.3)

Overview
---------
The KID Document Generator (Version 1.3) is a tool designed to automate the generation of KID (Key Information Document) Word files for contractors. The program allows you to select a data input file (CSV format) containing contractor information, a Word template with placeholders, and a save destination. It then processes the data and fills in the placeholders in the template, generating individual KID documents for each contractor.

Features
--------
- Select CSV Input File: Choose a CSV file containing contractor data (e.g., daily rate, assignment dates, etc.).
- Select Template File: Choose a Word document template (.docx) containing placeholders for contractor information.
- Set Save Destination: Specify the directory where the generated KID documents will be saved.
- Progress Tracking: A progress bar to indicate the document generation process.
- Financial Calculations: The program performs financial calculations, including National Insurance (NI), tax, and pension deductions, based on contractor data.
- Employer Contributions: Includes calculations for the employer's NI and pension contributions.
- File Naming: The generated documents are named using the contractor's name, assignment ID, and contract start/end dates for easy identification.
- Support for Multiple Contractors: Process multiple rows of contractor data from the input CSV file in one go.

Requirements
------------
- Python 3.x: The program is written in Python and requires Python 3.x to run.
- Required Libraries: You need to install the following libraries:
    - pandas (for reading and processing CSV data)
    - tkinter (for the graphical user interface)
    - docx (for manipulating Word documents)
    - os (for file handling)

To install the required libraries, use the following command:
pip install pandas python-docx

How to Use
-----------
1. Launch the Program: Run the KIDDocumentGenerator.py script to open the application.
2. Select the Data Input File: Click the "Browse" button under "1. Select Data Input File" to choose the CSV file containing contractor data.
3. Select the KID Template: Click the "Browse" button under "2. Select KID Template" to choose the Word document template with placeholders.
4. Set the Save Destination: Click the "Browse" button under "3. Set Save Destination" to select the folder where the generated documents will be saved.
5. Generate KID Documents: Once all the fields are filled in, click the "Generate KID Documents" button. The program will process the contractor data, calculate deductions (NI, tax, pension), replace the placeholders in the template, and save each generated document in the specified folder.

Financial Calculations
----------------------
- National Insurance (NI): The program calculates both employee and employer's NI contributions based on the monthly rate and UK thresholds.
- Tax: The program calculates the tax deduction based on UK income tax bands and personal allowances.
- Pension Contribution: The program calculates pension contributions based on the daily rate and an average pension percentage.

Employer Contributions:
-----------------------
- Employer's NI: The program calculates the employer's NI based on the monthly earnings above the threshold.
- Employer's Pension: The program calculates the employer's pension contribution based on qualifying earnings above a set threshold.

File Naming
-----------
The generated documents are saved using the following naming format:
CandidateName_AssignmentID_ContractStartDate_ContractEndDate_KIDFORM.docx

For example:
JohnDoe_12345_2025-02-01_2025-02-28_KIDFORM.docx

Notes
-----
- Ensure that the input CSV file has the following columns: Candidate, ID, Start Date, End Date, Pay Rate, Umbrella Company, Vacancy, Pay Unit.
- The program assumes a default working month of 20 working days unless specified otherwise.
- If a document fails to generate for a particular row due to missing or incorrect data, it will be skipped and an error message will be displayed in the console.

Version History
---------------
Version 1.3
- Initial release of the KID Document Generator with all primary functionality for generating KID documents from contractor data.


### Functionality Updates:

1. **Enhanced Document Population:**
   - Additional fields like Net Pay, Pension Contribution, Apprenticeship Levy, Total Deductions, and Monthly Rate are populated in the template.
   - Currency symbols are added to monetary fields.

2. **Calculation of Deductions:**
   - Tax deduction (20%) and NIC deduction (12%) are calculated based on the daily rate.
   - Pension contribution (5% of monthly rate) is calculated.
   - Umbrella fee (£20 per week) is included in total deductions.
   - Apprenticeship levy is calculated as 0.5% of the annual salary.

3. **Monthly Rate Calculation:**
   - Monthly rate is calculated by multiplying the daily rate by 22 working days.

4. **PDF Generation with Dynamic Data:**
   - PDF is generated using the populated template, which includes formatted financial data like deductions, net pay, and other figures.

5. **Dynamic File Naming:**
   - File names for PDFs include Candidate Name, Assignment ID, Contract Start Date, and Contract End Date.

6. **Umbrella Fee Fix:**
   - The umbrella fee (£20 per week) is now correctly included in the deductions section of the template.

7. **Support for Currency:**
   - Currency symbol is dynamically inserted for financial fields like daily rate, deductions, and net pay.

### Bug Fixes and Enhancements:

1. **CSV Column Mapping:**
   - The `get()` method is used to fetch CSV columns, preventing errors if a column is missing.

2. **File Naming Issue:**
   - File names now avoid invalid characters by replacing `/` with `-`, improving file management.

3. **Progress Bar Enhancement:**
   - The progress bar updates more accurately as documents are generated.

4. **Error Handling for Template and File Saving:**
   - Error handling for template reading and PDF generation is added, with error messages shown if an issue occurs.

5. **File Path Checks:**
   - Checks are in place to ensure files and directories are selected before document generation proceeds.

These updates improve calculations, data handling, template population, user experience, and error handling.


Contact
-------
For any issues or suggestions, feel free to reach out to the development team.
