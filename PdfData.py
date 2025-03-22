import os
import re
import pdfplumber
import tabula
import pandas as pd
import pyodbc
from flask import Flask, request, render_template, redirect, url_for
from werkzeug.utils import secure_filename
from datetime import datetime

app = Flask(__name__)
app.config["IMAGE_UPLOADS"] = "C:\\phpproject\\newtest\\abcd"
ALLOWED_EXTENSIONS = {'pdf'}

server = '192.168.1.77,2131'  
database = 'LQR_Dummy' 
username = 'dev'  
password = 'dev@abgrain'  

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_db_connection():
    conn = pyodbc.connect(f'DRIVER={{ODBC Driver 17 for SQL Server}};'
                          f'SERVER={server};'
                          f'DATABASE={database};'
                          f'UID={username};'
                          f'PWD={password}')
    return conn

def safe_float(value, precision=2):
    try:
        if isinstance(value, (int, float)):
            return round(value, precision)
        
        if isinstance(value, str):
            value = re.sub(r'[^\d.-]', '', value) 
            if value:
                return round(float(value), precision)
        return 0.0
    except (ValueError, TypeError) as e:
        print(f"Error converting value '{value}' to float: {e}")
        return 0.0

def format_date(value):
    try:
        if isinstance(value, str):
            value = value.strip()
            
           
            try:
                date_obj = datetime.strptime(value, "%m-%d-%Y %H:%M:%S")
                return date_obj.strftime("%Y-%m-%d %H:%M:%S")  
            except ValueError:
               
                date_obj = datetime.strptime(value, "%m-%d-%Y")
                return date_obj.strftime("%Y-%m-%d")  

        elif isinstance(value, datetime):
            
            return value.strftime("%Y-%m-%d %H:%M:%S") if value.hour or value.minute or value.second else value.strftime("%Y-%m-%d")

        else:
            return None  

    except Exception as e:
        print(f"Error formatting date: {e}")
        return None


@app.route('/home', methods=["GET", "POST"])
def upload_image():
    if request.method == "POST":

        if 'file' not in request.files:
            print("No file part")
            return redirect(request.url)

        image = request.files['file']
        if image.filename == '':
            print("No selected file")
            return redirect(request.url)

        if not allowed_file(image.filename):
            print("File type not allowed")
            return redirect(request.url)

        filename = secure_filename(image.filename)
        basedir = os.path.abspath(os.path.dirname(__file__))
        image.save(os.path.join(basedir, app.config["IMAGE_UPLOADS"], filename))
        file_path = os.path.join(basedir, app.config["IMAGE_UPLOADS"], filename)
        print(f"File saved at: {file_path}")
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
        if account_number and customer_id:
          
            try:
                df_list = tabula.read_pdf(file_path, pages='all', multiple_tables=True, lattice=True, guess=True)
                if not df_list or len(df_list) == 0:
                    df_list = tabula.read_pdf(file_path, pages='all', multiple_tables=True, stream=True, guess=True)

                if df_list:
                    main_df = df_list[
                        0]

                    cleaned_dfs = []
                    for df in df_list:
                        df.columns = main_df.columns
                        df = df[df[df.columns[0]] != df.columns[0]]  
                        cleaned_dfs.append(df)
                    final_df = pd.concat(cleaned_dfs, ignore_index=True)
                    final_df = final_df.dropna(how='all')  
                    final_df['Withdrawals'] = final_df['Withdrawals'].fillna(0).apply(safe_float)
                    final_df['Deposits'] = final_df['Deposits'].fillna(0).apply(safe_float)
                    final_df['Balance'] = final_df['Balance'].fillna(0).apply(safe_float)
                    final_df["Instr. ID"] = final_df["Instr. ID"].fillna(0)
                    final_df["Account Number"] = account_number if account_number else "Not Found"
                    final_df["Customer ID"] = customer_id if customer_id else "Not Found"
                    print("Extracted Data:\n", final_df)
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    for index, row in final_df.iterrows():
                        try:
                            formatted_date = format_date(row['Date'])
                            cursor.execute(""" 
                                INSERT INTO GetPdfData (Date, Remarks, [Tran Id-1], [UTR_Number], [Instr_ID], Withdrawals, Deposits, Balance, AccountNo, CustomerID,Uploadby)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?,1)
                            """, formatted_date, row['Remarks'], row['Tran Id-1'], row['UTR Number'], row['Instr. ID'],
                                safe_float(row['Withdrawals']), safe_float(row['Deposits']), safe_float(row['Balance']),
                                row['Account Number'], row['Customer ID'])
                            conn.commit()
                        except Exception as e:
                            print(f"Error inserting row {index}: {e}")
                            conn.rollback()
                    print("Data successfully inserted into GetPdfData table.")
            except Exception as e:
                print(f"Error processing PDF for Form 1: {e}")

        else:
            try:
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
                    final_df['Dr Amount'] = final_df['Dr Amount'].fillna(0).apply(safe_float)
                    final_df['Balance'] = final_df['Balance'].fillna(0).apply(lambda x: str(x).replace('Cr.', '').replace(',', '').strip())
                    final_df['Balance'] = final_df['Balance'].apply(safe_float)
                    final_df['Cr Amount'] = final_df['Cr Amount'].fillna(0).apply(lambda x: str(x).replace(',', '').replace('\r', '').strip())
                    final_df['Cr Amount'] = final_df['Cr Amount'].apply(safe_float)
                    final_df["Txn No."] = final_df["Txn No."].fillna("Not Found")
                    final_df["Description"] = final_df["Description"].fillna("Not Found")
                    final_df["Branch Name"] = final_df["Branch Name"].fillna("Not Found")
                    final_df["Cheque No."] = final_df["Cheque No."].fillna("Not Found")
                    final_df["KIMS\rRemarks"] = final_df["KIMS\rRemarks"].fillna("Not Found")
                    final_df["Account Number"] = account_number if account_number else "Not Found"
                    final_df["Customer ID"] = customer_id if customer_id else "Not Found"
                    print("Extracted Data:\n", final_df)
                    conn = get_db_connection()
                    cursor = conn.cursor()
                    for index, row in final_df.iterrows():
                        try:
                            formatted_date = format_date(row['Txn Date'])
                            cursor.execute(""" 
                                INSERT INTO GetPdfData (TxnNo, TxnDate, Description, BranchName, ChequeNo, DrAmount, CrAmount, Balance, AccountNo, KIMSRemarks,Uploadby)
                                VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?,?,1)
                            """, row['Txn No.'], formatted_date, row['Description'], row['Branch Name'],
                                row['Cheque No.'], row['Dr Amount'], row['Cr Amount'], row['Balance'], row['Account Number'], row['KIMS\rRemarks'])

                            conn.commit()

                        except Exception as e:
                            print(f"Error inserting row {index}: {e}")
                            conn.rollback()

                    print("Data successfully inserted into GetPdfData table.")
            except Exception as e:
                print(f"Error processing PDF for Form 2: {e}")

        return render_template("main.html", filename=filename)

    return render_template('main.html')

@app.route('/display/<filename>')
def display_image(filename):
    return redirect(url_for('static', filename="/Images" + filename), code=301)

if __name__ == '__main__':
    app.run(debug=True, port=2000)

