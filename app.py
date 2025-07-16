from flask import Flask, request, send_file
import os
import io
import base64
import tempfile
import time
from PyPDF2 import PdfReader, PdfWriter
from cryptography.fernet import Fernet
import win32com.client as win32
import pythoncom

app = Flask(__name__)

def generate_cipher(password):
    key = base64.urlsafe_b64encode(password.encode().ljust(32, b'0'))
    return Fernet(key)

# -------------------- ENCRYPTION --------------------
def encrypt_pdf(data, password):
    input_pdf = io.BytesIO(data)
    reader = PdfReader(input_pdf)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    writer.encrypt(password)
    output_pdf = io.BytesIO()
    writer.write(output_pdf)
    return output_pdf.getvalue(), 'application/pdf', 'encrypted.pdf'

def encrypt_text(data, password, ext):
    cipher = generate_cipher(password)
    encrypted_data = cipher.encrypt(data)
    mime_map = {
        '.csv': 'text/csv',
        '.json': 'application/json',
        '.xml': 'application/xml'
    }
    return encrypted_data, mime_map[ext], f'encrypted{ext}'

def encrypt_excel(data, password):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_input:
        temp_input.write(data)
        input_path = temp_input.name

    output_path = input_path.replace('.xlsx', '_encrypted.xlsx')
    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    wb = excel.Workbooks.Open(input_path)
    wb.SaveAs(output_path, FileFormat=51, Password=password)
    wb.Close(False)
    excel.Quit()

    with open(output_path, 'rb') as f:
        encrypted_data = f.read()

    os.remove(input_path)
    os.remove(output_path)

    return encrypted_data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'encrypted.xlsx'

# -------------------- DECRYPTION --------------------
def decrypt_pdf(data, password):
    input_pdf = io.BytesIO(data)
    reader = PdfReader(input_pdf)
    if reader.is_encrypted:
        reader.decrypt(password)
    writer = PdfWriter()
    for page in reader.pages:
        writer.add_page(page)
    output_pdf = io.BytesIO()
    writer.write(output_pdf)
    return output_pdf.getvalue(), 'application/pdf', 'decrypted.pdf'

def decrypt_text(data, password, ext):
    cipher = generate_cipher(password)
    decrypted_data = cipher.decrypt(data)
    mime_map = {
        '.csv': 'text/csv',
        '.json': 'application/json',
        '.xml': 'application/xml'
    }
    return decrypted_data, mime_map[ext], f'decrypted{ext}'

def decrypt_excel(data, password):
    pythoncom.CoInitialize()

    with tempfile.NamedTemporaryFile(delete=False, suffix='.xlsx') as temp_input:
        temp_input.write(data)
        input_path = temp_input.name

    output_path = input_path.replace('.xlsx', '_decrypted.xlsx')

    excel = win32.gencache.EnsureDispatch('Excel.Application')
    excel.DisplayAlerts = False
    excel.Visible = False

    wb = excel.Workbooks.Open(input_path, Password=password)
    new_wb = excel.Workbooks.Add()
    ws_src = wb.Worksheets(1)
    ws_dst = new_wb.Worksheets(1)
    ws_src.UsedRange.Copy(Destination=ws_dst.Range("A1"))

    wb.Close(SaveChanges=False)
    new_wb.SaveAs(output_path, FileFormat=51)
    new_wb.Close(False)
    excel.Quit()

    with open(output_path, 'rb') as f:
        decrypted_data = f.read()

    os.remove(input_path)
    os.remove(output_path)

    return decrypted_data, 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'decrypted.xlsx'

# -------------------- ROUTES --------------------

@app.route('/encrypt', methods=['POST'])
def encrypt_route():
    if 'file' not in request.files:
        return "❌ No file uploaded", 400
    file = request.files['file']
    password = request.form.get('password', 'rajesh')
    file_ext = os.path.splitext(file.filename)[1].lower()
    data = file.read()

    try:
        if file_ext == '.pdf':
            encrypted_data, mimetype, name = encrypt_pdf(data, password)
        elif file_ext in ['.csv', '.json', '.xml']:
            encrypted_data, mimetype, name = encrypt_text(data, password, file_ext)
        elif file_ext in ['.xls', '.xlsx']:
            encrypted_data, mimetype, name = encrypt_excel(data, password)
        else:
            return "❌ Unsupported file type", 400

        return send_file(io.BytesIO(encrypted_data), mimetype=mimetype, as_attachment=True, download_name=name)
    except Exception as e:
        return f"❌ Encryption failed: {e}", 500

@app.route('/decrypt', methods=['POST'])
def decrypt_route():
    if 'file' not in request.files:
        return "❌ No file uploaded", 400
    file = request.files['file']
    password = request.form.get('password', 'rajesh')
    file_ext = os.path.splitext(file.filename)[1].lower()
    data = file.read()

    try:
        if file_ext == '.pdf':
            decrypted_data, mimetype, name = decrypt_pdf(data, password)
        elif file_ext in ['.csv', '.json', '.xml']:
            decrypted_data, mimetype, name = decrypt_text(data, password, file_ext)
        elif file_ext in ['.xls', '.xlsx']:
            decrypted_data, mimetype, name = decrypt_excel(data, password)
        else:
            return "❌ Unsupported file type", 400

        return send_file(io.BytesIO(decrypted_data), mimetype=mimetype, as_attachment=True, download_name=name)
    except Exception as e:
        return f"❌ Decryption failed: {e}", 500

if __name__ == '__main__':
    app.run(debug=True, port=5000)
