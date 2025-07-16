from flask import Flask, request, send_file
import os
import io
import base64
import tempfile
from PyPDF2 import PdfReader, PdfWriter
from cryptography.fernet import Fernet
import msoffcrypto

app = Flask(__name__)

# ----------------- Utilities -----------------

def generate_cipher(password):
    key = base64.urlsafe_b64encode(password.encode().ljust(32, b'0'))
    return Fernet(key)

# ----------------- PDF -----------------

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

# ----------------- CSV / JSON / XML -----------------

def encrypt_text(data, password, ext):
    cipher = generate_cipher(password)
    encrypted_data = cipher.encrypt(data)
    mime_map = {
        '.csv': 'text/csv',
        '.json': 'application/json',
        '.xml': 'application/xml'
    }
    return encrypted_data, mime_map[ext], f'encrypted{ext}'

def decrypt_text(data, password, ext):
    cipher = generate_cipher(password)
    decrypted_data = cipher.decrypt(data)
    mime_map = {
        '.csv': 'text/csv',
        '.json': 'application/json',
        '.xml': 'application/xml'
    }
    return decrypted_data, mime_map[ext], f'decrypted{ext}'

# ----------------- Excel -----------------

def decrypt_excel(data, password):
    try:
        decrypted = io.BytesIO()
        office_file = msoffcrypto.OfficeFile(io.BytesIO(data))
        office_file.load_key(password=password)
        office_file.decrypt(decrypted)
        decrypted.seek(0)
        return decrypted.read(), 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', 'decrypted.xlsx'
    except Exception as e:
        raise Exception(f"Excel decryption failed: {e}")

def encrypt_excel_generic(data, password):
    try:
        cipher = generate_cipher(password)
        encrypted_data = cipher.encrypt(data)
        return encrypted_data, 'application/octet-stream', 'encrypted_excel.bin'
    except Exception as e:
        raise Exception(f"Excel encryption failed: {e}")

# ----------------- Routes -----------------

@app.route('/encrypt', methods=['POST'])
def encrypt_route():
    if 'file' not in request.files:
        return "❌ No file uploaded", 400

    file = request.files['file']
    password = request.form.get('password', 'rajesh')

    # Fallback: Use a default extension if filename is missing
    filename = file.filename or 'file.pdf'  # default to PDF
    file_ext = os.path.splitext(filename)[1].lower()

    if not file_ext:
        return "❌ Unable to determine file extension", 400

    data = file.read()

    try:
        if file_ext == '.pdf':
            output_data, mimetype, name = encrypt_pdf(data, password)
        elif file_ext in ['.csv', '.json', '.xml']:
            output_data, mimetype, name = encrypt_text(data, password, file_ext)
        elif file_ext in ['.xls', '.xlsx']:
            output_data, mimetype, name = encrypt_excel_generic(data, password)
        else:
            return f"❌ Unsupported file type: {file_ext}", 400

        return send_file(
            io.BytesIO(output_data),
            mimetype=mimetype,
            as_attachment=True,
            download_name=name
        )
    except Exception as e:
        return f"❌ Encryption failed: {str(e)}", 500


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
            output_data, mimetype, name = decrypt_pdf(data, password)
        elif file_ext in ['.csv', '.json', '.xml']:
            output_data, mimetype, name = decrypt_text(data, password, file_ext)
        elif file_ext in ['.xls', '.xlsx']:
            output_data, mimetype, name = decrypt_excel(data, password)
        elif file_ext == '.bin':
            # Decrypt previously encrypted Excel binary
            output_data, mimetype, name = decrypt_text(data, password, '.xlsx')
        else:
            return "❌ Unsupported file type", 400

        return send_file(io.BytesIO(output_data), mimetype=mimetype, as_attachment=True, download_name=name)
    except Exception as e:
        return f"❌ Decryption failed: {e}", 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000)
