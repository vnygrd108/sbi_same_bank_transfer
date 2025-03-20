from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

@app.route('/', methods=['GET', 'POST'])
def upload_form():
    if request.method == 'POST':
        return process_file()
    return render_template('sbt.html')

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file uploaded"
    
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    
    file_path = os.path.join(UPLOAD_FOLDER, file.filename)
    file.save(file_path)
    
    # Processing the Excel file
    same_bank_transfer = pd.read_excel(file_path, engine="xlrd", dtype={'Bank-A/C': str})
    same_bank_transfer = same_bank_transfer[same_bank_transfer['IFSC'].str.contains('SBIN', na=False)]
    same_bank_transfer = same_bank_transfer[['Bank-A/C', 'IFSC', 'Amount', 'Vendor']]
    same_bank_transfer = same_bank_transfer.rename(columns={
        'Bank-A/C': 'Account Number',
        'IFSC': 'IFSC Code',
        'Vendor': 'Name'
    })
    
    today_date = datetime.today().strftime('%d/%m/%y')
    same_bank_transfer.insert(2, 'Date', today_date)
    
    start_serial = int(request.form.get("start_serial", 1))
    total_amount = same_bank_transfer["Amount"].sum()
    
    default_row = pd.DataFrame([{
        "Account Number": "41256726637",
        "IFSC Code": "07773",
        "Date": today_date,
        "Amount": total_amount,
        "Name": "SCST GORHE",
        "Serial": start_serial,
        "Mode": "",
        "Formula": ""
    }])
    
    same_bank_transfer = pd.concat([default_row, same_bank_transfer], ignore_index=True)
    same_bank_transfer["Serial"] = range(start_serial, start_serial + len(same_bank_transfer))
    same_bank_transfer["IFSC Code"] = same_bank_transfer["IFSC Code"].astype(str).str[-5:]
    
    same_bank_transfer.loc[0, "Formula"] = same_bank_transfer.loc[0, "Account Number"] + "#" + \
                                            same_bank_transfer.loc[0, "IFSC Code"] + "#" + \
                                            same_bank_transfer.loc[0, "Date"] + "#" + \
                                            str(same_bank_transfer.loc[0, "Amount"]) + "##" + \
                                            str(same_bank_transfer.loc[0, "Serial"]) + "#" + \
                                            same_bank_transfer.loc[0, "Name"] + "#"
    
    same_bank_transfer.loc[1:, "Formula"] = same_bank_transfer.loc[1:, "Account Number"] + "#" + \
                                            same_bank_transfer.loc[1:, "IFSC Code"] + "#" + \
                                            same_bank_transfer.loc[1:, "Date"] + "##" + \
                                            same_bank_transfer.loc[1:, "Amount"].astype(str) + "#" + \
                                            same_bank_transfer.loc[1:, "Serial"].astype(str) + "#" + \
                                            same_bank_transfer.loc[1:, "Name"] + "#"
    
    date_str = datetime.today().strftime('%d.%m.%Y')
    output_filename = f"SCHCT_{date_str}_Same_Bank_Transfer.xlsx"
    output_path = os.path.join(PROCESSED_FOLDER, output_filename)
    same_bank_transfer.to_excel(output_path, index=False)
    
    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
