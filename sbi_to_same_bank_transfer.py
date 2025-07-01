from flask import Flask, render_template, request, send_file
import pandas as pd
from datetime import datetime
import os
from werkzeug.utils import secure_filename

app = Flask(__name__)
UPLOAD_FOLDER = "uploads"
PROCESSED_FOLDER = "processed"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PROCESSED_FOLDER, exist_ok=True)

# Allowed account numbers and their details
ACCOUNT_MAP = {
    "43557115725": {"name": "SBI SCHCT", "ifsc": "07773"},
    "41256726637": {"name": "SBI SCHCT", "ifsc": "07773"},
    "34889306900": {"name": "SCHCT Gorhe", "ifsc": "07773"},
    "40237416058": {"name": "SBI SCHCT", "ifsc": "07773"}
}

@app.route('/', methods=['GET'])
def upload_form():
    return render_template('sbt.html', account_map=ACCOUNT_MAP)

@app.route('/process', methods=['POST'])
def process_file():
    if 'file' not in request.files:
        return "No file uploaded"

    file = request.files['file']
    if file.filename == '':
        return "No selected file"

    file_path = os.path.join(UPLOAD_FOLDER, secure_filename(file.filename))
    file.save(file_path)

    account_input = request.form.get("account_input", "").strip()
    start_serial = int(request.form.get("start_serial", 1))

    if account_input not in ACCOUNT_MAP:
        return render_template("sbt.html", error="Invalid account number", account_map=ACCOUNT_MAP)

    details = ACCOUNT_MAP[account_input]
    account_name = details["name"]
    ifsc_code = details["ifsc"][-5:]

    # Read Excel
    ext = os.path.splitext(file_path)[1]
    if ext == ".xlsx":
        df = pd.read_excel(file_path, engine="openpyxl", dtype={'Bank-A/C': str})
    elif ext == ".xls":
        df = pd.read_excel(file_path, engine="xlrd", dtype={'Bank-A/C': str})
    else:
        return "Unsupported file format"

    same_bank_transfer = df[df['IFSC'].str.contains('SBIN', na=False)]
    same_bank_transfer = same_bank_transfer[['Bank-A/C', 'IFSC', 'Amount', 'Vendor']].copy()
    same_bank_transfer.columns = ['Account Number', 'IFSC Code', 'Amount', 'Name']
    today_date = datetime.today().strftime('%d/%m/%y')
    same_bank_transfer.insert(2, 'Date', today_date)

    total_amount = same_bank_transfer["Amount"].sum()

    default_row = pd.DataFrame([{
        "Account Number": account_input,
        "IFSC Code": ifsc_code,
        "Date": today_date,
        "Amount": total_amount,
        "Name": account_name,
        "Serial": start_serial,
        "Mode": "",
        "Formula": ""
    }])

    same_bank_transfer = pd.concat([default_row, same_bank_transfer], ignore_index=True)
    same_bank_transfer["Serial"] = range(start_serial, start_serial + len(same_bank_transfer))
    same_bank_transfer["IFSC Code"] = same_bank_transfer["IFSC Code"].astype(str).str[-5:]

    same_bank_transfer.loc[0, "Formula"] = f"{same_bank_transfer.loc[0, 'Account Number']}#" \
                                           f"{same_bank_transfer.loc[0, 'IFSC Code']}#" \
                                           f"{same_bank_transfer.loc[0, 'Date']}#" \
                                           f"{same_bank_transfer.loc[0, 'Amount']}##" \
                                           f"{same_bank_transfer.loc[0, 'Serial']}#" \
                                           f"{same_bank_transfer.loc[0, 'Name']}#"

    same_bank_transfer.loc[1:, "Formula"] = same_bank_transfer.loc[1:, "Account Number"] + "#" + \
                                            same_bank_transfer.loc[1:, "IFSC Code"] + "#" + \
                                            same_bank_transfer.loc[1:, "Date"] + "##" + \
                                            same_bank_transfer.loc[1:, "Amount"].astype(str) + "#" + \
                                            same_bank_transfer.loc[1:, "Serial"].astype(str) + "#" + \
                                            same_bank_transfer.loc[1:, "Name"] + "#"

    output_filename = f"SCHCT_{datetime.today().strftime('%d.%m.%Y')}_Same_Bank_Transfer.xlsx"
    output_path = os.path.join(PROCESSED_FOLDER, output_filename)
    same_bank_transfer.to_excel(output_path, index=False)

    return send_file(output_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
