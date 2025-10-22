from flask import Flask, render_template, request, flash, send_file, redirect, url_for
import os
import pandas as pd
import openpyxl
from docxtpl import DocxTemplate
from docx2pdf import convert
from PIL import Image, ImageDraw, ImageFont
from pptx import Presentation
from pptx.util import Inches
import zipfile
from datetime import datetime
import tempfile
import shutil

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Change to a secure key

# Persistent work folder for uploads and reading
WORK_FOLDER = 'temp/uploads'
GENERATE_FOLDER = 'temp/generated'
os.makedirs(WORK_FOLDER, exist_ok=True)
os.makedirs(GENERATE_FOLDER, exist_ok=True)

# Function for generating certificates (PNG or PPTX with placeholders)
def generate_certificates(output_folder, font_path="arialbd.ttf", font_size=100):
    excel_file = os.path.join(WORK_FOLDER, 'data.xlsx')
    template_file = os.path.join(WORK_FOLDER, 'template.png')  # Or .pptx
    if not os.path.exists(excel_file) or not os.path.exists(template_file):
        raise FileNotFoundError("Required files (data.xlsx or template) not found in work_folder.")
    
    data = pd.read_excel(excel_file)
    os.makedirs(output_folder, exist_ok=True)
    if template_file.lower().endswith('.png') or template_file.lower().endswith('.jpg'):
        # Image: Draw name (no placeholders, overlay text)
        font_name = ImageFont.truetype(font_path, font_size)
        for index, row in data.iterrows():
            name = row["Name"]  # Full name from Excel
            certificate = Image.open(template_file)
            draw = ImageDraw.Draw(certificate)
            width, height = certificate.size()
            draw.text((width//2,700), name, fill="navy", font=font_name)
            output_path = os.path.join(output_folder, f"certificate_{name}.png")
            certificate.save(output_path)
    elif template_file.lower().endswith('.pptx'):
        # PPTX: Replace placeholders in text boxes
        for index, row in data.iterrows():
            name = row["Name"]
            prs = Presentation(template_file)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text") and shape.text:
                        # Replace placeholders (e.g., "Your Name" -> actual name)
                        shape.text = shape.text.replace("{{your_name}}", name).replace("Your Name", name)
            output_path = os.path.join(output_folder, f"certificate_{name}.pptx")
            prs.save(output_path)

# Functions for generating Transcripts (DOCX with placeholders)
def TranscriptExcel_data():
    filename = os.path.join(WORK_FOLDER, 'data.xlsx')
    if not os.path.exists(filename):
        raise FileNotFoundError("data.xlsx not found in work_folder.")
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)

def TranscriptDocument(output_directory, row_data):
    template = os.path.join(WORK_FOLDER, 'template.docx')
    if not os.path.exists(template):
        raise FileNotFoundError("template.docx not found in work_folder.")
    doc = DocxTemplate(template)
    current_date = datetime.now().strftime("%B %d, %Y")
    # Map placeholders to Excel data (adjust columns as needed)
    doc.render({
        'your_name': row_data[1],  # First name
        'your_surname': row_data[2],  # Last name
        'student_id': row_data[0],
        # Add more mappings if needed, e.g., 'your_grade': row_data[3]
        'cur_date': current_date
    })
    doc_name = os.path.join(output_directory, f"{row_data[1]}.docx")
    doc.save(doc_name)
    return doc_name

def TranscriptPdf(doc_path, pdf_directory):
    pdf_path = os.path.join(pdf_directory, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")
    convert(doc_path, pdf_path)
    return pdf_path

def generate_transcripts(output_folder, option):
    data_rows = TranscriptExcel_data()
    docx_dir = os.path.join(output_folder, 'docx')
    pdf_dir = os.path.join(output_folder, 'pdf')
    os.makedirs(docx_dir, exist_ok=True)
    os.makedirs(pdf_dir, exist_ok=True)
    for row in data_rows[1:]:
        if option in ["doc", "both"]:
            doc_path = TranscriptDocument(docx_dir, row)
        if option in ["pdf", "both"]:
            if option == "pdf":
                doc_path = TranscriptDocument(pdf_dir, row)
            TranscriptPdf(doc_path, pdf_dir)
            if option == "pdf":
                os.remove(doc_path)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/generate', methods=['POST'])
def generate():
    doc_type = request.form.get('doc_type')
    template = request.files.get('template')
    excel = request.files.get('excel')
    option = request.form.get('option', 'both')

    if not template or not excel:
        flash('Please upload both template and Excel files.')
        return redirect(url_for('index'))

    # Save uploaded files to work_folder with fixed names
    template_ext = os.path.splitext(template.filename)[1]
    excel_ext = os.path.splitext(excel.filename)[1]
    template_path = os.path.join(WORK_FOLDER, f'template{template_ext}')
    excel_path = os.path.join(WORK_FOLDER, f'data{excel_ext}')
    template.save(template_path)
    excel.save(excel_path)

    # Generate based on type (reads from work_folder)
    output_folder = os.path.join(GENERATE_FOLDER, doc_type)
    try:
        if doc_type == 'certificate':
            generate_certificates(output_folder)
        elif doc_type == 'transcript':
            generate_transcripts(output_folder, option)
        flash('Documents generated successfully! Placeholders replaced with names from Excel.')
    except Exception as e:
        flash(f'Error generating documents: {str(e)}')

    return redirect(url_for('preview', doc_type=doc_type))

@app.route('/preview/<doc_type>')
def preview(doc_type):
    output_folder = os.path.join(GENERATE_FOLDER, doc_type)
    files = []
    if os.path.exists(output_folder):
        for root, dirs, filenames in os.walk(output_folder):
            for filename in filenames:
                files.append(os.path.join(root, filename))
    return render_template('preview.html', doc_type=doc_type, files=files[:5])

@app.route('/download/<doc_type>')
def download(doc_type):
    output_folder = os.path.join(GENERATE_FOLDER, doc_type)
    zip_path = os.path.join(GENERATE_FOLDER, f'{doc_type}_documents.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, dirs, files in os.walk(output_folder):
            for file in files:
                zipf.write(os.path.join(root, file), os.path.relpath(os.path.join(root, file), output_folder))
    shutil.rmtree(output_folder)
    return send_file(zip_path, as_attachment=True, download_name=f'{doc_type}_documents.zip')

@app.route('/uploaded_files')
def uploaded_files():
    files = []
    if os.path.exists(WORK_FOLDER):
        files = [f for f in os.listdir(WORK_FOLDER) if os.path.isfile(os.path.join(WORK_FOLDER, f))]
    return render_template('uploaded_files.html', files=files)

if __name__ == '__main__':
    app.run(debug=True)