# 1 - Locate Data Input File (.csv)
# 2 - Locate KID FORM Template
# 3 - Set SAVE FILE Destination
# 4 - Assign values from Columns
# 5 - Calculate deductions
# 6 - Populate new figures into template
# 7 - Produce PDF

import tkinter as tk
from tkinter import filedialog, messagebox
import pandas as pd
from fpdf import FPDF
import os
import time
from tkinter.ttk import Progressbar

class KIDGeneratorApp:
    def __init__(self, root):
        self.root = root
        self.root.title("KID Document Generator")
        self.create_widgets()

    def create_widgets(self):
        # Labels and buttons for file selection
        tk.Label(self.root, text="1. Select Data Input File (.csv):").pack(anchor='w', padx=10)
        self.data_file_button = tk.Button(self.root, text="Browse", command=self.select_data_file)
        self.data_file_button.pack(padx=10, pady=5, anchor='w')

        tk.Label(self.root, text="2. Select KID Template:").pack(anchor='w', padx=10)
        self.template_button = tk.Button(self.root, text="Browse", command=self.select_template_file)
        self.template_button.pack(padx=10, pady=5, anchor='w')

        tk.Label(self.root, text="3. Set Save Destination:").pack(anchor='w', padx=10)
        self.save_button = tk.Button(self.root, text="Browse", command=self.set_save_destination)
        self.save_button.pack(padx=10, pady=5, anchor='w')

        # Progress Bar
        self.progress = Progressbar(self.root, orient="horizontal", length=300, mode="determinate")
        self.progress.pack(pady=10)

        # Generate Button
        self.generate_button = tk.Button(self.root, text="Generate KID Documents", command=self.generate_documents, state=tk.DISABLED)
        self.generate_button.pack(pady=10)

        # File paths
        self.data_file_path = ""
        self.template_file_path = ""
        self.save_directory = ""

    def select_data_file(self):
        self.data_file_path = filedialog.askopenfilename(filetypes=[("CSV Files", "*.csv")])
        if self.data_file_path:
            self.check_ready_to_generate()

    def select_template_file(self):
        self.template_file_path = filedialog.askopenfilename(filetypes=[("Text Files", "*.txt")])
        if self.template_file_path:
            self.check_ready_to_generate()

    def set_save_destination(self):
        self.save_directory = filedialog.askdirectory()
        if self.save_directory:
            self.check_ready_to_generate()

    def check_ready_to_generate(self):
        if self.data_file_path and self.template_file_path and self.save_directory:
            self.generate_button.config(state=tk.NORMAL)

    def generate_documents(self):
        # Load data from CSV
        data = pd.read_csv(self.data_file_path)

        successful_generations = 0
        total_rows = len(data)

        for index, row in data.iterrows():
            candidate_first_name = row.get("CandidateFirstName", "")
            candidate_last_name = row.get("CandidateLastName", "")
            contract_start_date = row.get("ContractStartDate", "")
            contract_end_date = row.get("ContractEndDate", "")
            daily_rate = row.get("DailyRate", 0.0)
            umbrella_fee = row.get("UmbrellaFee", 0.0)

            # Calculate deductions (example: simple static tax and fee calculation)
            tax_deduction = daily_rate * 0.2  # 20% tax
            nic_deduction = daily_rate * 0.12  # 12% NIC
            net_pay = daily_rate - tax_deduction - nic_deduction - umbrella_fee

            # Read template content
            with open(self.template_file_path, 'r') as template_file:
                template_content = template_file.read()

            # Populate template with dynamic data
            filled_template = template_content.format(
                CandidateFirstName=candidate_first_name,
                CandidateLastName=candidate_last_name,
                ContractStartDate=contract_start_date,
                ContractEndDate=contract_end_date,
                DailyRate=f"£{daily_rate:.2f}",
                TaxDeduction=f"£{tax_deduction:.2f}",
                NICDeduction=f"£{nic_deduction:.2f}",
                UmbrellaFee=f"£{umbrella_fee:.2f}",
                NetPay=f"£{net_pay:.2f}"
            )

            # Generate PDF
            pdf = FPDF()
            pdf.add_page()
            pdf.set_auto_page_break(auto=True, margin=15)
            pdf.set_font("Arial", size=12)
            for line in filled_template.split('\n'):
                pdf.multi_cell(0, 10, line)

            # File name formatting
            file_name = f"{candidate_first_name}_{candidate_last_name}_-_{contract_start_date}_-_{contract_end_date}_KID.pdf"
            save_path = os.path.join(self.save_directory, file_name)

            try:
                pdf.output(save_path)
                successful_generations += 1
            except Exception as e:
                print(f"Error generating {file_name}: {e}")

            # Update progress bar
            self.progress['value'] = ((index + 1) / total_rows) * 100
            self.root.update_idletasks()

        # Show completion dialog
        messagebox.showinfo("Generation Complete", f"Successfully generated {successful_generations} of {total_rows} documents.")

if __name__ == "__main__":
    root = tk.Tk()
    app = KIDGeneratorApp(root)
    root.mainloop()
