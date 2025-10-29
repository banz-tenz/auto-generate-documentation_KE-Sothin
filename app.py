import os
from flask import Flask, render_template, request, flash, redirect, url_for, send_file
from werkzeug.utils import secure_filename
from docxtpl import DocxTemplate
from docx2pdf import convert
from PIL import Image, ImageDraw, ImageFont
import openpyxl
import pandas as pd  # Removed pandas dependency for certificates to use openpyxl consistently
from datetime import datetime
from pptx import Presentation
from pptx.enum.text import PP_ALIGN  # For text alignment
from pptx.util import Pt  # For font size in points
from pptx.dml.color import RGBColor  # For font color
import zipfile
import io
# from flask_mysqldb import MySql
import mysql.connector

db_config = {
    'host': 'localhost',
    'user': 'root',
    'password': '',
    'database': 'user_result'
}

# Function to insert multiple certificates into MySQL (assuming for bulk generation)
def insert_multiple_certificate(nb, first_name, last_name, certificate_filename, certificate_file):
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        query = "INSERT INTO certificates (nb, first_name, last_name, certificate_filename, certificate_file) VALUES (%s, %s, %s, %s, %s)"
        cursor.execute(query, (nb, first_name, last_name, certificate_filename, certificate_file))
        conn.commit()
    except mysql.connector.Error as err:
        print(f"Error inserting into DB: {err}")
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()

# Function to insert individual certificate into MySQL
def insert_individual_certificate(nb, first_name, last_name, certificate_filename, certificate_file):
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        query = "INSERT INTO individual_certificate (nb, first_name, last_name, certificate_filename, certificate_file) VALUES (%s, %s, %s, %s, %s)"
        cursor.execute(query, (nb, first_name, last_name, certificate_filename, certificate_file))
        conn.commit()
    except mysql.connector.Error as err:
        print(f"Error inserting into DB: {err}")
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()

# Function to get the next nb (assuming nb is auto-increment or we need to calculate it)
def get_next_nb():
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor()
        cursor.execute("SELECT MAX(nb) FROM individual_certificate")  # Assuming nb is shared or per table; adjust if needed
        result = cursor.fetchone()
        if result[0] is None:
            return 1
        else:
            return result[0] + 1
    except mysql.connector.Error as err:
        print(f"Error getting next nb: {err}")
        return 1  # Default to 1 if error
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()

# Function to get all individual certificates from MySQL
def get_individual_certificates_from_db():
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM individual_certificate")
        results = cursor.fetchall()
        return results
    except mysql.connector.Error as err:
        print(f"Error fetching individual certificates from DB: {err}")
        return []
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()

# Function to get all certificates from MySQL
def get_certificates_from_db():
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM certificates")
        results = cursor.fetchall()
        return results
    except mysql.connector.Error as err:
        print(f"Error fetching certificates from DB: {err}")
        return []
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()

# Function to get all transcripts from MySQL
def get_transcripts_from_db():
    conn = None
    cursor = None
    try:
        conn = mysql.connector.connect(**db_config)
        cursor = conn.cursor(dictionary=True)
        cursor.execute("SELECT * FROM Transcript")
        results = cursor.fetchall()
        return results
    except mysql.connector.Error as err:
        print(f"Error fetching transcripts from DB: {err}")
        return []
    finally:
        if cursor:
            cursor.close()
        if conn and conn.is_connected():
            conn.close()

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Change this to a secure key
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function for generating certificates (modified to use openpyxl consistently, assuming name in first column after header)
def generate_certificates(excel_file, template_file, output_folder, template_filename, option, font_path="arialbd.ttf", font_size=100):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values)[1:]  # Skip header row
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    font_name = ImageFont.truetype(font_path, font_size)
    generated_files = []
    for row in data:
        name = row[0]  # Assuming name is in the first column
        if template_filename.lower().endswith('.png') or template_filename.lower().endswith('.jpg'):
            certificate = Image.open(template_file)
            width, height = certificate.size
            draw = ImageDraw.Draw(certificate)
            
            name_position = (width//2, (height//2)+40)
            draw.text(name_position, str(name), fill="navy", font=font_name, anchor='mm')
            output_path = os.path.join(output_folder, "certificate_" + str(name) + ".png")
            certificate.save(output_path)
            generated_files.append(output_path)
            # Optionally insert into DB for multiple certificates
            # Assuming name is full name, split if needed; for now, skip or adjust
            print("Certificate generated for {} and saved to {}".format(name, output_path))
        elif template_filename.lower().endswith(".pptx"):
            # Use python-pptx for PPTX templates with placeholders
            prs = Presentation(template_file)
            for slide in prs.slides:
                for shape in slide.shapes:
                    if hasattr(shape, "text_frame"):
                        text_frame = shape.text_frame
                        for paragraph in text_frame.paragraphs:
                            full_text = paragraph.text
                            if "{{your_name}}" in full_text:
                                paragraph.text = full_text.replace("{{your_name}}", str(name))
                                # Apply formatting to the runs containing the name
                                for run in paragraph.runs:
                                    if str(name) in run.text:  # Ensure the run contains the replaced name
                                        run.font.name = 'Arial'
                                        run.font.size = Pt(54)  # Use Pt() for points
                                        run.font.bold = True
                                        run.font.color.rgb = RGBColor(0, 0, 128)
                                # Set paragraph alignment to justify
                                paragraph.alignment = PP_ALIGN.CENTER
                                print(f"Replaced {{your_name}} with {name} in paragraph.")  # Debug print
            # Save as PPTX
            output_path_pptx = os.path.join(output_folder, "certificate_" + str(name) + ".pptx")
            prs.save(output_path_pptx)
            output_path = output_path_pptx
            generated_files.append(output_path)
            # If PDF is needed, convert
            if option in ['pdf', 'both']:
                pdf_path = os.path.join(output_folder, "certificate_" + str(name) + ".pdf")
                convert(output_path_pptx, pdf_path)
                if option == 'pdf':
                    os.remove(output_path_pptx)
                    output_path = pdf_path
                generated_files.append(pdf_path)
    print("All certificates have been generated!")
    return generated_files

# Function for generating individual certificate
def generate_individual_certificate(name, template_file, output_folder, template_filename, font_path="arialbd.ttf", font_size=100):
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    font_name = ImageFont.truetype(font_path, font_size)
    if template_filename.lower().endswith('.png') or template_filename.lower().endswith('.jpg'):
        certificate = Image.open(template_file)
        width, height = certificate.size
        draw = ImageDraw.Draw(certificate)
        
        name_position = (width//2, (height//2)+40)
        draw.text(name_position, str(name), fill="navy", font=font_name, anchor='mm')
        output_path = os.path.join(output_folder, "certificate_" + str(name).replace(" ", "_") + ".png")
        certificate.save(output_path)
        return output_path
    # Assuming only PNG/JPG for individual, as PPTX might be complex for individual download

# Functions for generating Transcripts (unchanged)
def TranscriptExcel_data(filename):
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)

def TranscriptDocument(template, output_directory, row_data):
    doc = DocxTemplate(template)
    current_date = datetime.now().strftime("%B %d, %Y")
    doc.render({
        "student_id": row_data[0],
        "first_name": row_data[1],
        "last_name": row_data[2],
        "logic": row_data[3],
        "l_g": row_data[4],
        "bcum": row_data[5],
        "bc_g": row_data[6],
        "design": row_data[7],
        "d_g": row_data[8],
        "p1": row_data[9],
        "p1_g": row_data[10],
        "e1": row_data[11],
        "e1_g": row_data[12],
        "wd": row_data[13],
        "wd_g": row_data[14],
        "algo": row_data[15],
        "al_g": row_data[16],
        "p2": row_data[17],
        "p2_g": row_data[18],
        "e2": row_data[19],
        "e2_g": row_data[20],
        "sd": row_data[21],
        "sd_g": row_data[22],
        "js": row_data[23],
        "js_g": row_data[24],
        "php": row_data[25],
        "ph_g": row_data[26],
        "db": row_data[27],
        "db_g": row_data[28],
        "vc1": row_data[29],
        "v1_g": row_data[30],
        "node": row_data[31],
        "no_g": row_data[32],
        "e3": row_data[33],
        "e3_g": row_data[34],
        "p3": row_data[35],
        "p3_g": row_data[36],
        "oop": row_data[37],
        "op_g": row_data[38],
        "lar": row_data[39],
        "lar_g": row_data[40],
        "vue": row_data[41],
        "vu_g": row_data[42],
        "vc2": row_data[43],
        "v2_g": row_data[44],
        "e4": row_data[45],
        "e4_g": row_data[46],
        "p4": row_data[47],
        "p4_g": row_data[48],
        "int": row_data[49],
        "in_g": row_data[50],
        'cur_date': current_date
    })
    doc_name = os.path.join(output_directory, "{}.docx".format(row_data[1]))
    doc.save(doc_name)
    return doc_name

def TranscriptPdf(doc_path, pdf_directory):
    pdf_path = os.path.join(pdf_directory, os.path.splitext(os.path.basename(doc_path))[0] + ".pdf")
    convert(doc_path, pdf_path)
    return pdf_path

def generate_transcripts(excel_file, template_file, docx_directory, pdf_directory, option):
    os.makedirs(docx_directory, exist_ok=True)
    os.makedirs(pdf_directory, exist_ok=True)
    data_rows = TranscriptExcel_data(excel_file)
    generated_files = []

    for row in data_rows[1:]:
        if option in ["doc", "both"]:
            doc_path = TranscriptDocument(template_file, docx_directory, row)
            generated_files.append(doc_path)
        if option in ["pdf", "both"]:
            if option == "pdf":
                doc_path = TranscriptDocument(template_file, pdf_directory, row)
            pdf_path = TranscriptPdf(doc_path, pdf_directory)
            generated_files.append(pdf_path)
            if option == "pdf":
                os.remove(doc_path)
    print("All files for option '{}' have been generated!".format(option))
    return generated_files

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/complete-info-multiple')
def complete_info_multiple():
    return render_template('complete_certificate_info_multiple.html')

@app.route('/generate-multipule', methods=['POST'])
def generate_multiple():
    if request.method == 'POST':
        doc_type = request.form.get('doc_type')
        option = request.form.get('option')
        template = request.files.get('template')
        excel = request.files.get('excel')
        

        if not template or not excel:
            flash('Please upload both template and Excel files.')
            return redirect(url_for('complete_info_multiple'))
        
        template_filename = secure_filename(template.filename)
        excel_filename = secure_filename(excel.filename)
        template_path = os.path.join(app.config['UPLOAD_FOLDER'], template_filename)
        excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
        template.save(template_path)
        excel.save(excel_path)
        
        try:
            if doc_type == 'transcript':
                docx_directory = 'Transcript_Doc'
                pdf_directory = 'Transcript_PDF'
                generated_files = generate_transcripts(excel_path, template_path, docx_directory, pdf_directory, option)
            elif doc_type == 'certificate':
                output_folder = 'Certificates'
                generated_files = generate_certificates(excel_path, template_path, output_folder, template_filename, option)
            else:
                flash('Invalid document type.')
                return redirect(url_for('complete_info_multiple'))
            
            flash('Documents generated successfully!')
            return redirect(url_for('view'))
        except Exception as e:
            flash(f'An error occurred: {str(e)}')
            return redirect(url_for('complete_info_multiple'))

@app.route('/complete-info-certificate')
def complete_info_certificate():
    return render_template('complete-certificate.html')

@app.route('/complete-info-certificate-individual')
def complete_info_certificate_individual():
    return render_template('complete-certificate-individual.html')

@app.route('/generate_certificate_individual', methods=['POST'])
def generate_certificate_individual():
    first_name = request.form.get('first_name')
    last_name = request.form.get('last_name')
    # Use default template
    template_path = os.path.join(app.root_path, 'template', 'certificate_template.png')  # Updated to use 'template/certificate_template.png'
    template_filename = os.path.basename(template_path)
    
    if not first_name or not last_name:
        flash('Please provide both first and last name.')
        return redirect(url_for('complete_info_certificate_individual'))
    
    if not os.path.exists(template_path):
        flash('Default certificate template not found. Please contact admin.')
        return redirect(url_for('complete_info_certificate_individual'))
    
    name = first_name.upper() + " " + last_name.upper()
    output_folder = 'Certificates_Individual'
    output_path = generate_individual_certificate(name, template_path, output_folder, template_filename)
    
    # Get next nb
    nb = get_next_nb()
    certificate_filename = os.path.basename(output_path)
    certificate_file = name  # Assuming certificate_file is the full name
    
    # Insert into DB
    insert_individual_certificate(nb, first_name.upper(), last_name.upper(), certificate_filename, certificate_file)
    
    flash('Certificate generated successfully!')
    return redirect(url_for('view'))

@app.route('/complet-info-transcript')
def complete_info_transcript():
    return render_template('complete_transcript.html')

@app.route('/download-zip/<doc_type>')
def download_zip(doc_type):
    if doc_type == 'certificate':
        folder = 'Certificates'
    elif doc_type == 'transcript':
        folder = 'Transcript_PDF'  # Assuming PDF for zip
    else:
        flash('Invalid type')
        return redirect(url_for('view'))
    
    zip_buffer = io.BytesIO()
    with zipfile.ZipFile(zip_buffer, 'w', zipfile.ZIP_DEFLATED) as zip_file:
        for root, dirs, files in os.walk(folder):
            for file in files:
                zip_file.write(os.path.join(root, file), file)
    zip_buffer.seek(0)
    return send_file(zip_buffer, as_attachment=True, download_name=f'{doc_type}_files.zip', mimetype='application/zip')

@app.route('/download_file/<dir_name>/<filename>')
def download_file(dir_name, filename):
    file_path = os.path.join(dir_name, filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    else:
        flash('File not found.')
        return redirect(url_for('view'))

@app.route('/view')
def view():
    # Get data from all tables
    individual_certificates = get_individual_certificates_from_db()
    certificates = get_certificates_from_db()
    transcripts = get_transcripts_from_db()
    # For other files, still list from directories if needed
    generated_files = {}
    directories = ['Transcript_Doc', 'Transcript_PDF', 'Certificates', 'Associate_Documents', 'Associate_PDF']
    for dir_name in directories:
        if os.path.exists(dir_name):
            generated_files[dir_name] = os.listdir(dir_name)
    return render_template('view.html', generated_files=generated_files, individual_certificates=individual_certificates, certificates=certificates, transcripts=transcripts)

if __name__ == '__main__':
    app.run(debug=True)
