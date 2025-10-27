import os
from flask import Flask, render_template, request, flash, redirect, url_for
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

app = Flask(__name__)
app.secret_key = 'your_secret_key_here'  # Change this to a secure key
UPLOAD_FOLDER = 'uploads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
os.makedirs(UPLOAD_FOLDER, exist_ok=True)

# Function for generating certificates (modified to use openpyxl consistently, assuming name in first column after header)
def generate_certificates(excel_file, template_file, output_folder, template_filename,option,font_path="arialbd.ttf", font_size=100):
    workbook = openpyxl.load_workbook(excel_file)
    sheet = workbook.active
    data = list(sheet.values)[1:]  # Skip header row
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    font_name = ImageFont.truetype(font_path, font_size)
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
            # If PDF is needed, convert
            if option in ['pdf', 'both']:
                pdf_path = os.path.join(output_folder, "certificate_" + str(name) + ".pdf")
                convert(output_path_pptx, pdf_path)
                if option == 'pdf':
                    os.remove(output_path_pptx)
                    output_path = pdf_path
    print("All certificates have been generated!")


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

    for row in data_rows[1:]:
        if option in ["doc", "both"]:
            doc_path = TranscriptDocument(template_file, docx_directory, row)
        if option in ["pdf", "both"]:
            if option == "pdf":
                doc_path = TranscriptDocument(template_file, pdf_directory, row)
            TranscriptPdf(doc_path, pdf_directory)
            if option == "pdf":
                os.remove(doc_path)
    print("All files for option '{}' have been generated!".format(option))

# Merged function to generate all documents (transcripts, certificates, associate)
def generate_all(excel_transcript, template_transcript, excel_certificate, template_certificate, excel_associate, template_associate, option):
    # Generate transcripts
    docx_directory_transcript = 'Transcript_Doc'
    pdf_directory_transcript = 'Transcript_PDF'
    generate_transcripts(excel_transcript, template_transcript, docx_directory_transcript, pdf_directory_transcript, option)
    
    # Generate certificates
    output_folder_cert = 'Certificates'
    generate_certificates(excel_certificate, template_certificate, output_folder_cert)
    
    # Generate associate
    docx_directory_associate = 'Associate_Documents'
    pdf_directory_associate = 'Associate_PDF'
    GeneratAssociate(excel_associate, template_associate, docx_directory_associate, pdf_directory_associate, option)

@app.route('/')
def home():
    return render_template('home.html')

@app.route('/complete-info')
def complete_info():
    return render_template('complete_info.html')

@app.route('/generate', methods=['POST'])
def generate():
    if request.method == 'POST':
        doc_type = request.form.get('doc_type')
        option = request.form.get('option')
        template = request.files.get('template')
        excel = request.files.get('excel')
        

        if not template or not excel:
            flash('Please upload both template and Excel files.')
            return redirect(url_for('complete_info'))
        
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
                generate_transcripts(excel_path, template_path, docx_directory, pdf_directory, option)
            elif doc_type == 'certificate':
                output_folder = 'Certificates'
                generate_certificates(excel_path, template_path, output_folder,template_filename, option)
            else:
                flash('Invalid document type.')
                return redirect(url_for('complete_info'))
            
            flash('Documents generated successfully!')
            return redirect(url_for('preview'))
        except Exception as e:
            flash(f'An error occurred: {str(e)}')
            return redirect(url_for('complete_info'))

@app.route('/preview')
def preview():
    files = os.listdir(app.config['UPLOAD_FOLDER'])
    return render_template('preview.html', files=files)

if __name__ == '__main__':
    app.run(debug=True)
