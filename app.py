from flask import Flask, request, send_file
import io
import os
import pandas as pd
from PyPDF2 import PdfReader, PdfWriter
import xlsxwriter
import tempfile

app = Flask(__name__)

def encrypt_pdf(file_contents, password):
    input_pdf = io.BytesIO(file_contents)
    reader = PdfReader(input_pdf)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(password)
    output_pdf = io.BytesIO()
    writer.write(output_pdf)
    return output_pdf.getvalue()

def create_protected_excel(dataframe, password):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        dataframe.to_excel(writer, index=False, sheet_name='Sheet1')

        workbook = writer.book
        worksheet = writer.sheets['Sheet1']

        worksheet.protect(password)
        workbook.set_properties({'title': 'Encrypted Output'})
    return output.getvalue()

def encrypt_xlsx(file_contents, password):
    # Load the Excel file into pandas and re-write it with password protection
    with tempfile.NamedTemporaryFile(suffix=".xlsx") as tmp:
        tmp.write(file_contents)
        tmp.flush()
        df = pd.read_excel(tmp.name)
        return create_protected_excel(df, password)

def convert_and_encrypt(file_contents, filename, password):
    ext = os.path.splitext(filename)[1].lower()

    if ext == '.csv':
        df = pd.read_csv(io.BytesIO(file_contents))
    elif ext == '.json':
        df = pd.read_json(io.BytesIO(file_contents))
    elif ext == '.xml':
        df = pd.read_xml(io.BytesIO(file_contents))
    else:
        return None  # Unsupported format

    return create_protected_excel(df, password)

@app.route('/encrypt', methods=['POST'])
def encrypt():
    if 'file' not in request.files:
        return "No file part", 400

    file = request.files['file']
    password = request.form.get('password', 'rajesh')
    file_contents = file.read()
    filename = file.filename.lower()

    if filename.endswith('.pdf'):
        encrypted_file = encrypt_pdf(file_contents, password)
        return send_file(
            io.BytesIO(encrypted_file),
            mimetype='application/pdf',
            as_attachment=True,
            download_name='encrypted_output.pdf'
        )
    elif filename.endswith('.xlsx'):
        encrypted_excel = encrypt_xlsx(file_contents, password)
        return send_file(
            io.BytesIO(encrypted_excel),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='protected_output.xlsx'
        )
    elif filename.endswith(('.csv', '.json', '.xml')):
        encrypted_excel = convert_and_encrypt(file_contents, filename, password)
        if not encrypted_excel:
            return "Unsupported file format", 400

        return send_file(
            io.BytesIO(encrypted_excel),
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name='protected_output.xlsx'
        )
    else:
        return "Unsupported file format", 400

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
