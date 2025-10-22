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

app.secret_key = 'your_secert_key_here'

upload_folder = 'temp/uploads'
generate_folder = 'temp/generated'
os.makedirs(upload_folder, exist_ok=True)
os.makedirs(generate_folder, exist_ok=True)

# Function for generate certificate
def generate_certificates(output_folder, font_path = "arialdb.ttf", font_size = 100):
    excel_file = os.path.join(upload_folder, 'data.xlsx')
    template_file = os.path.join(upload_folder, 'template.png')
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
            width,height = certificate.size()
            draw = ImageDraw.Draw(certificate)
            draw.text((width//2,700), name, fill="orange", font=font_name, anchor='mm')
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
def TranscriptExcel():
    filename = os.path.join(upload_folder, 'data.xlsx')
    if not os.path.join.exists(filename):
        raise FileNotFoundError("data.xlsx not found in upload folder.")
    workbook = openpyxl.load_workbook(filename)
    sheet = workbook.active
    return list(sheet.values)


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