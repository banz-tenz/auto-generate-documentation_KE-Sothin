import os
import io
import zipfile
from flask import Flask, render_template, request, send_file
import pandas as pd
from PIL import Image, ImageDraw, ImageFont

app = Flask(__name__)

# Folders
UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "generated_certificates"
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

# Font settings
FONT_PATH = "arialbd.ttf"  # Make sure this font exists
FONT_SIZE = 100

# ===== Certificate Generation Function =====
def generate_certificates(excel_file_path, template_file_path):
    data = pd.read_excel(excel_file_path)
    generated_files = []

    for index, row in data.iterrows():
        name = row["Name"]
        certificate = Image.open(template_file_path)
        draw = ImageDraw.Draw(certificate)

        # Position adjustment based on name length
        if len(name) >= 15 and len(name) < 25:
            name_position = (550, 600)
        elif len(name) >= 10 and len(name) < 15:
            name_position = (700, 600)
        else:
            name_position = (730, 600)

        font = ImageFont.truetype(FONT_PATH, FONT_SIZE)
        draw.text(name_position, name, fill="orange", font=font)

        output_path = os.path.join(OUTPUT_FOLDER, f"certificate_{name}.png")
        certificate.save(output_path)
        generated_files.append(output_path)

    return generated_files

# ===== Routes =====
@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        # Check uploaded files
        if 'template_file' not in request.files or 'excel_file' not in request.files:
            return "Please upload both Excel file and template PNG."

        template_file = request.files['template_file']
        excel_file = request.files['excel_file']

        template_path = os.path.join(UPLOAD_FOLDER, template_file.filename)
        excel_path = os.path.join(UPLOAD_FOLDER, excel_file.filename)

        template_file.save(template_path)
        excel_file.save(excel_path)

        # Generate certificates
        files = generate_certificates(excel_path, template_path)

        # Create ZIP file
        zip_buffer = io.BytesIO()
        with zipfile.ZipFile(zip_buffer, "w") as zip_file:
            for file_path in files:
                zip_file.write(file_path, os.path.basename(file_path))

        zip_buffer.seek(0)
        return send_file(zip_buffer, mimetype='application/zip', as_attachment=True, download_name="certificates.zip")

    # GET request → show form
    return render_template("certificateUpload.html")

# ===== Run App =====
if __name__ == "__main__":
    app.run(debug=True)
