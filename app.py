from flask import Flask, request, render_template, send_file, after_this_request
import os
import pdfplumber
from openpyxl import Workbook
from werkzeug.utils import secure_filename
from openpyxl.styles import Font

app = Flask(__name__)

# Configure upload folder and allowed extensions
UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf'}
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Check allowed file types
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# Extract table data from text-based PDFs using pdfplumber
def extract_table_pdf(file_path):
    tables = []
    bank_details = []
    with pdfplumber.open(file_path) as pdf:
        for i, page in enumerate(pdf.pages):
            if i == 0:  # Assume the first page contains bank details and transaction data
                text = page.extract_text()
                if text:
                    for line in text.splitlines():
                        if ":" in line:  # Identify lines with key-value pairs for bank details
                            bank_details.append(line.split(":", 1))
                page_tables = page.extract_tables()
                if page_tables:
                    for table in page_tables:
                        tables.append(table)  # Add all tables to be processed for transactions
            else:
                page_tables = page.extract_tables()
                print(f"Page {i + 1} Tables: {page_tables}")  # Debugging
                tables.extend(page_tables)
    return bank_details, tables

# Fuzzy match columns to expected headings
def fuzzy_match_columns(row, expected_columns):
    matched_row = []
    for col in row:
        matched = None
        for expected in expected_columns:
            if expected.lower() in (col or '').lower():
                matched = expected
                break
        matched_row.append(matched if matched else col)
    return matched_row

# Process PDF and save all tables into a single sheet in Excel
# First sheet contains bank details, second sheet contains combined tables
def process_pdf(file_path):
    bank_details, tables = extract_table_pdf(file_path)

    if not tables or len(tables) == 0:
        return None, "No tables found in the PDF."

    workbook = Workbook()

    # Add bank details to the first sheet
    sheet1 = workbook.active
    sheet1.title = "Bank Details"
    if bank_details:
        for detail in bank_details:
            sheet1.append(detail)

    # Add combined tables to the second sheet
    sheet2 = workbook.create_sheet(title="Bank Transactions")
    header_written = False  # Track if the header has been written

    expected_columns = [
        "Date",
        "Particulars / Description",
        "Cheque / Ref",
        "Debit",
        "Credit",
        "Balance / Closing balance"
    ]
    for table in tables:
     for i, row in enumerate(table):
            # Perform fuzzy matching for the first row of each table
            if i == 0:
                 # Detect the first row of the table as a potential header
                if not header_written:
                    sheet2.append(row)  # Write the first detected header
                    header_written = True
                # Make the header row bold
                    for cell in sheet2[1]:  # First row in the sheet
                        cell.font = Font(bold=True)
                else:
                    # Skip subsequent headers (i.e., the first row of subsequent tables)
                    continue
            else:
                # Append non-header rows (data rows)
                sheet2.append(row)

    return workbook, None

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/bank_statement')
def bank_statement():
    return render_template('bank_statement.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return "No file part"
    file = request.files['file']
    if file.filename == '':
        return "No selected file"
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)

        # Process the file
        workbook, error = process_pdf(file_path)

        if error:
            return error

        # Save workbook to a temporary file
        output_excel = os.path.join(app.config['UPLOAD_FOLDER'], 'Bank Transaction.xlsx')
        workbook.save(output_excel)
         # Schedule cleanup after the response
        @after_this_request
        def cleanup_files(response):
            try:
                # Delete the uploaded PDF file
                if os.path.exists(file_path):
                    os.remove(file_path)
                    print(f"Deleted uploaded file: {file_path}")

                # Delete the generated Excel file
                if os.path.exists(output_excel):
                    os.remove(output_excel)
                    print(f"Deleted generated Excel file: {output_excel}")

                # Remove any additional files in the uploads folder
                for f in os.listdir(app.config['UPLOAD_FOLDER']):
                    file_to_delete = os.path.join(app.config['UPLOAD_FOLDER'], f)
                    if os.path.exists(file_to_delete):
                        os.remove(file_to_delete)
                        print(f"Deleted additional file: {file_to_delete}")
            except Exception as e:
                print(f"Error during cleanup: {e}")
            return response

        # Send the Excel file as a response
        return send_file(output_excel, as_attachment=True, download_name='Bank Transaction.xlsx')

    else:
        return "Invalid file format. Please upload a PDF."

if __name__ == '__main__':
    app.run(debug=True)
