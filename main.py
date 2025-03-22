import tabula
import pdfplumber
import pandas as pd
import re
import tkinter as tk
from tkinter import filedialog


def upload_pdf():
    root = tk.Tk()
    root.withdraw() 
    file_path = filedialog.askopenfilename(
        title="Select a PDF file",
        filetypes=[("PDF files", "*.pdf")]
    )
    return file_path

try:
   
    file_path = upload_pdf()
    if not file_path:
        print("No file selected.")
    else:
        to_excel_path = "output.xlsx"

        account_number = None
        customer_id = None

        with pdfplumber.open(file_path) as pdf:
            for page in pdf.pages:
                text = page.extract_text()
                if text:
                    
                    acc_match = re.search(r'(?:Account Number|Account No|A/C No|A/c No|Account #)[:\s-]+(\d+)', text, re.IGNORECASE)
                    if acc_match:
                        account_number = acc_match.group(1)

                
                    cust_match = re.search(r'(?:Customer ID|Customer No|Cust ID)[:\s-]+(\d+)', text, re.IGNORECASE)
                    if cust_match:
                        customer_id = cust_match.group(1)

                    if account_number and customer_id:
                        break

  
        df_list = tabula.read_pdf(file_path, pages='all', multiple_tables=True, lattice=True, guess=True)
        if not df_list or len(df_list) == 0:
            df_list = tabula.read_pdf(file_path, pages='all', multiple_tables=True, stream=True, guess=True)

        if df_list:
            main_df = df_list[0]

            cleaned_dfs = []
            for df in df_list:
                df.columns = main_df.columns
                df = df[df[df.columns[0]] != df.columns[0]]
                cleaned_dfs.append(df)

            final_df = pd.concat(cleaned_dfs, ignore_index=True)
            final_df = final_df.dropna(how='all')

            final_df["Account Number"] = account_number if account_number else "Not Found"
            final_df["Customer ID"] = customer_id if customer_id else "Not Found"

          
            final_df.to_excel(to_excel_path, index=False)
            print(f"Data successfully exported to {to_excel_path}")
        else:
            print("No tables found in the PDF. Try using OCR for scanned PDFs.")

except Exception as e:
    print(f"Error extracting table: {e}")
