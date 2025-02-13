# KID Document Generator v1.5

## Overview

The **KID Document Generator** is a simple desktop application designed to automate the creation of KID (Key Information Document) files. By using your data from a CSV file and a Word template, this tool helps you generate personalized KID documents for each employee.

It automatically calculates taxes, National Insurance, pensions, and other deductions, and then populates a pre-defined Word template with the necessary information.

---

## Features

- **CSV File Input**: Upload a CSV file with employee data (e.g., name, pay rate, contract start/end dates).
- **Word Template**: Select a Word document template where placeholders will be replaced with data from the CSV.
- **Destination Folder**: Choose where to save the generated KID documents.
- **Progress Tracking**: View the progress of document generation in real-time.
- **Automatic Calculations**: The tool calculates taxes, NI deductions, pensions, and more.
- **Final Document Creation**: The KID documents are generated and saved in the selected destination folder.

---

## How to Use

### Step 1: Open the Application
Once you've downloaded the application, simply double-click the `.exe` file to open it.

### Step 2: Select Your Files
1. **Select Data Input File (.csv)**:
   - Click the “Browse” button under the **1. Select Data Input File** label.
   - Choose the CSV file that contains your employee data (ensure the file is formatted correctly with the necessary fields like name, pay rate, contract dates, etc.).

2. **Select KID Template**:
   - Click the “Browse” button under **2. Select KID Template**.
   - Choose the Word document template (ensure it has placeholders like `{CandidateName}`, `{NetPay}`, etc.).

3. **Set Save Destination**:
   - Click the “Browse” button under **3. Set Save Destination**.
   - Select the folder where you want to save the generated KID documents.

### Step 3: Generate KID Documents
- Once all the files are selected, the **Generate KID Documents** button will become active.
- Click the button to begin the document generation process.
- The progress bar will update as each document is generated.

### Step 4: Check Your Results
- After the process is complete, a message will appear showing how many documents were successfully generated. You can now open the folder you selected to find your completed KID documents.

---

## Requirements

- **Windows** operating system (This version is packaged as an `.exe` file).
- **Microsoft Word** (to open the generated KID documents).
- **Python** (Not required to be installed by the user as the app is packaged as an executable).

---

## Troubleshooting

- **"No files selected" error**: Ensure you have selected all required files (CSV and Word template) and the destination folder.
- **Document saving issues**: Check if you have write permissions for the folder where you're saving the documents.
- **Missing Data in CSV**: If some employees have missing information (e.g., name, pay rate), the tool will skip them and report it in the log.

---

## License & Disclaimer

This application is provided as-is. The developers are not responsible for any incorrect calculations or data loss. Please review the generated documents before sharing or storing them.

---

This program was created by Peter Grint (RainyGrinch Studios) and is for private use only - not for external distribution

