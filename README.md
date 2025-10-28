# Automated Document Generator

A Flask-based web application for generating certificates,and transcripts documents from Excel data and templates.

## Description

This application allows users to upload Excel files containing student data and corresponding templates (DOCX, PPTX, or images) to generate personalized documents. It supports generating DOCX files, PDF files, or both. The app is built using Flask and integrates with libraries like `docxtpl`, `docx2pdf`, `PIL`, `openpyxl`, and `python-pptx`.

## Features

- **Document Types**:
  - Certificates (PNG/JPG images or PPTX with placeholders and formatting: Arial 54pt, bold, navy color, justified alignment)
  - Transcripts (DOCX/PDF)
- **Output Formats**: DOCX only, PDF only, or both.
- **PPTX Support**: Use `{{your_name}}` placeholders in PowerPoint templates for certificates, with automatic replacement and styling.
- **Web Interface**: Simple form-based UI for uploading files and selecting options.
- **File Upload**: Secure handling of template and Excel files.
- **Preview**: View generated files in the uploads directory.

## Installation

1. **Clone or Download the Repository**:
   - Place the Python script (e.g., `app.py`) in your project directory.

2. **Install Dependencies**:
   - Ensure you have Python 3.7+ installed.
   - Install required libraries using pip:
     ```
     pip install -r requirements.txt
     ```
   - Note: `docx2pdf` requires LibreOffice or Microsoft Office for PDF conversion. Install LibreOffice if not present.

3. **Font File**:
   - Ensure `arialbd.ttf` (Arial Bold font) is in the project root directory for certificate image generation.

4. **Directory Structure**:
   - Create the following folders in your project directory:
     ```
     AUTO GENERATE DOCUMENTATOIN_FINAL
      ├── static
      │   ├── images
      │   └── style
      │       └── main.css
      ├── template
      └── templates
         ├── complete_certificate_info_multiple.html
         ├── complete_transcript.html
         ├── complete-certificate-individual.html
         ├── complete-certificate.html
         ├── home.html
         └── view.html
      ├── app.py
      ├── helper.py
      ├── README.md
      ├── requirement.txt
      ├── student_data_Transcript.xlsx
      ├── student_names_certificate_test.xlsx
      └── student_names_certificate.xlsx
     ```

## Usage

1. **Run the Application**:
   - Execute the script:
     ```
     python app.py
     ```
   - Access the app at `http://127.0.0.1:5000/`.

2. **Navigate the App**:
   - **Home Page** (`/`): Welcome page with a link to start.
   - **Complete Info Page** (`/complete-info`): Form to select document type, upload template and Excel files, and choose output format.
   - **Generate** (`/generate`): Processes uploads and generates documents.
   - **view** (`/view`): Lists uploaded/generated files.

3. **Excel File Format**:
   - **Certificates**: First column after header should contain names (e.g., "Name").
   - **Transcripts**: Columns as per the code (student_id, first_name, etc., up to 51 columns).
   - **Associate Degrees**: Columns as per the code (id_kh, name_kh, etc.).
   - Ensure headers are in the first row, data in subsequent rows.

4. **Template Files**:
   - **Certificates (Image)**: PNG/JPG template for overlay text.
   - **Certificates (PPTX)**: PPTX template with `{{your_name}}` placeholder in a text box. The replaced text will be styled as Arial 54pt, bold, navy color, justified.
   - **Transcripts/Associate**: DOCX template with placeholders matching the render keys in the code.

5. **Output**:
   - Generated files are saved in folders like `Certificates/`, `Transcript_Doc/`, `Transcript_PDF/`, etc.
   - View them via the Preview page or directly in the directories.

## Requirements

- Python 3.7+
- Flask
- werkzeug
- docxtpl
- docx2pdf
- pillow (PIL)
- openpyxl
- python-pptx
- LibreOffice (for PDF conversion)

## File Structure

```
Auto Generate Documentation/ 
├── app.py # Main Flask application 
├── requirements.txt # Python dependencies 
├── arialbd.ttf # Font file for certificates 
├── templates/ 
    │ 
    ├── home.html # Home page template 
    │ 
    ├── complete_info.html # Form page template 
    │ 
    └── preview.html # Preview page template 
├── static/ 
    │ 
    └── style/ 
        │ 
        └── main.css # CSS styles 
├── uploads/ # Temporary uploads (auto-created) 
├── Certificates/ # Generated certificates (auto-created) 
├── Transcript_Doc/ # Generated transcript DOCX (auto-created) 
├── Transcript_PDF/ # Generated transcript PDF (auto-created)
```

## Troubleshooting (Fixing Problems)

- **TypeError: 'tuple' object is not callable**: Ensure Excel files have correct structure. The app uses openpyxl consistently.
- **PDF Conversion Fails**: Install LibreOffice and ensure it's in your PATH.
- **Font Not Found**: Verify `arialbd.ttf` is in the root directory.
- **PPTX Placeholder Not Replaced**: Ensure `{{your_name}}` is exactly in the PPTX template. Check console debug output for issues.
- **First Certificate Missing Name**: Check Excel data (first column after header) and PPTX template. Use debug prints to troubleshoot.
- **File Upload Errors**: Check file sizes and formats; ensure uploads folder is writable.

## Contributing (Helping Improve)

Feel free to fork the repository and submit pull requests for improvements.

## License

This project is open-source. Use at your own risk. No warranties provided.