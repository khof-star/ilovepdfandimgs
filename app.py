from flask import Flask, render_template, request, redirect, url_for, send_file, send_from_directory, abort, flash
from werkzeug.utils import secure_filename
from PyPDF2 import PdfMerger, PdfReader, PdfWriter
from pdf2image import convert_from_path
from PIL import Image
import img2pdf
from docx2pdf import convert as docx2pdf_convert
from pdf2docx import Converter
from pptx import Presentation
from pptx.util import Inches
import pdfplumber
import pandas as pd
import fitz  # PyMuPDF
import subprocess
import zipfile
import traceback
import os

app = Flask(__name__)

# Folder configurations
UPLOAD_FOLDER = "uploads"
COMPRESS_FOLDER = "compressed_files"
SPLIT_FOLDER = "split_files"
THUMBNAIL_FOLDER = os.path.join("static", "thumbnails")
WORD_FOLDER = os.path.join("static", "word_outputs")
PDF_FOLDER = "converted"
JPG_FOLDER = "converted_jpgs"
PPT_FOLDER = "static/ppt_outputs"
EXCEL_FOLDER = 'static/excel_outputs'
os.makedirs(PPT_FOLDER, exist_ok=True)
# Configure app
app.config["UPLOAD_FOLDER"] = UPLOAD_FOLDER

# Ensure directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(PDF_FOLDER, exist_ok=True)
os.makedirs(SPLIT_FOLDER, exist_ok=True)
os.makedirs(JPG_FOLDER, exist_ok=True)
os.makedirs(COMPRESS_FOLDER, exist_ok=True)
os.makedirs(THUMBNAIL_FOLDER, exist_ok=True)
os.makedirs(EXCEL_FOLDER, exist_ok=True)


@app.route("/home")
def home():
    return render_template("home.html")


@app.route("/")
def index():
    return render_template("home.html")   # Optional: make / go to home


# Show the merge page (GET)
@app.route("/merge", methods=["GET"])
def merge():
    return render_template("merge.html")


@app.route("/uploaded", methods=["POST"])
def uploaded():
    files = []
    uploaded_files = request.files.getlist("pdfs")

    for file in uploaded_files:
        filename = secure_filename(file.filename)
        save_path = os.path.join(UPLOAD_FOLDER, filename)

        file.save(save_path)

        if os.path.getsize(save_path) == 0:
            return f"Error: {filename} is empty or corrupted."

        try:
            images = convert_from_path(
                save_path,
                first_page=1,
                last_page=1,
                poppler_path=r"C:\poppler\poppler-24.08.0\Library\bin"
            )

            thumb_filename = f"{filename}_thumb.jpg"
            thumb_path = os.path.join(THUMBNAIL_FOLDER, thumb_filename)
            images[0].save(thumb_path, "JPEG")

            files.append({
                "filename": filename,
                "thumbnail": f"thumbnails/{thumb_filename}"
            })

        except Exception as e:
            return f"Error generating thumbnail for {filename}: {e}"

    return render_template("merge_result.html", files=files)


@app.route("/merge", methods=["POST"])
def merge_pdfs():
    filenames = request.form.getlist("filenames")
    if not filenames:
        return "Error: No files selected for merging."

    merger = PdfMerger()
    for name in filenames:
        path = os.path.join(UPLOAD_FOLDER, secure_filename(name))
        merger.append(path)

    output_path = os.path.join(UPLOAD_FOLDER, "merged.pdf")
    merger.write(output_path)
    merger.close()

    return send_file(output_path, as_attachment=True)


@app.route("/split")
def split():
    return render_template("split.html")


@app.route("/split_uploaded", methods=["POST"])
def split_uploaded():
    file = request.files.get("pdfs")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    reader = PdfReader(pdf_path)
    num_pages = len(reader.pages)
    split_files = []

    for i in range(num_pages):
        writer = PdfWriter()
        writer.add_page(reader.pages[i])

        split_filename = f"{os.path.splitext(filename)[0]}_page_{i+1}.pdf"
        split_path = os.path.join(SPLIT_FOLDER, split_filename)

        with open(split_path, "wb") as out_file:
            writer.write(out_file)

        images = convert_from_path(
            split_path,
            first_page=1,
            last_page=1,
            poppler_path=r"C:\poppler\poppler-24.08.0\Library\bin"
        )

        thumb_filename = f"{split_filename}_thumb.jpg"
        thumb_path = os.path.join(THUMBNAIL_FOLDER, thumb_filename)
        images[0].save(thumb_path, "JPEG")

        split_files.append({
            "filename": split_filename,
            "thumbnail": f"thumbnails/{thumb_filename}",
            "size": f"{round(os.path.getsize(split_path)/1024,1)} KB"
        })

    return render_template("split_result.html", files=split_files)


@app.route("/split_download", methods=["POST"])
def split_download():
    filenames = request.form.getlist("filenames")
    if not filenames:
        return "No files selected.", 400

    zip_path = os.path.join(SPLIT_FOLDER, "split_pages.zip")
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for name in filenames:
            file_path = os.path.join(SPLIT_FOLDER, name)
            zipf.write(file_path, arcname=name)

    return send_file(zip_path, as_attachment=True)


@app.route("/compress")
def compress():
    """
    Show the upload page for PDF compression.
    """
    return render_template("compress.html")


@app.route("/compress_uploaded", methods=["POST"])
def compress_uploaded():
    """
    Handle the uploaded PDF and prepare compression preview.
    """
    file = request.files.get("pdfs")
    if not file:
        return "No file uploaded.", 400

    # Save the uploaded PDF to disk
    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    # Generate a thumbnail of the first page
    images = convert_from_path(
        pdf_path,
        first_page=1,
        last_page=1,
        poppler_path=r"C:\poppler\poppler-24.08.0\Library\bin"
    )
    thumb_filename = f"{filename}_thumb.jpg"
    thumb_path = os.path.join(THUMBNAIL_FOLDER, thumb_filename)
    images[0].save(thumb_path, "JPEG")

    # Build file metadata for the template
    file_info = {
        "filename": filename,
        "thumbnail": f"thumbnails/{thumb_filename}",
        "size": f"{round(os.path.getsize(pdf_path) / 1024, 1)} KB"
    }

    return render_template("compress_result.html", file=file_info)


@app.route("/compress_download", methods=["POST"])
def compress_download():
    """
    Compress the PDF based on selected compression level and return it.
    """
    filename = request.form.get("filename")
    compression_level = request.form.get("compression_level")

    if not filename:
        return "No file selected.", 400

    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_path = os.path.join(COMPRESS_FOLDER, f"compressed_{filename}")

    # Compress the PDF using PyMuPDF
    compress_pdf(input_path, output_path, compression_level)

    return send_file(output_path, as_attachment=True)


def compress_pdf(input_path, output_path, compression_level):
    """
    Compress a PDF by rasterizing pages to images at a specified DPI.
    This is effective for reducing file size, especially for scanned PDFs.
    """
    # Map compression levels to DPI scaling factors
    dpi_map = {
        "extreme": 0.5,       # Very low quality / highest compression
        "recommended": 0.7,   # Medium quality
        "less": 0.9           # Higher quality
    }
    dpi_factor = dpi_map.get(compression_level, 0.7)

    # Open the original PDF
    doc = fitz.open(input_path)

    # Create a new PDF container
    new_doc = fitz.open()

    for page_num in range(len(doc)):
        page = doc.load_page(page_num)

        # Get the page dimensions
        rect = page.rect
        width = int(rect.width * dpi_factor)
        height = int(rect.height * dpi_factor)

        # Rasterize the page to an image with the specified scale
        pix = page.get_pixmap(matrix=fitz.Matrix(dpi_factor, dpi_factor))

        # Create a temporary single-page PDF from the image
        img_pdf = fitz.open()
        img_pdf.new_page(width=pix.width, height=pix.height)
        img_page = img_pdf[0]
        img_page.insert_image(
            fitz.Rect(0, 0, pix.width, pix.height),
            pixmap=pix
        )

        # Append the rasterized page to the output PDF
        new_doc.insert_pdf(img_pdf)
        img_pdf.close()

    # Save the compressed PDF
    new_doc.save(output_path)
    new_doc.close()
    doc.close()

 

@app.route("/rotate", methods=["GET"])
def rotate():
    return render_template("rotate.html")

@app.route("/rotate_upload", methods=["POST"])
def rotate_upload():
     if "images" not in request.files:
        return "No file part", 400

     file = request.files["images"]
     if file.filename == "":
        return "No selected file", 400

     filename = secure_filename(file.filename)
     file_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
     file.save(file_path)

     image_url = url_for("uploaded_file", filename=filename)

     return render_template("rotate_upload.html", image_url=image_url, filename=filename)
@app.route("/uploads/<filename>")
def uploaded_file(filename):
    return send_from_directory(app.config["UPLOAD_FOLDER"], filename)


@app.route("/protect", methods=["GET", "POST"])
def protect_upload():
    if request.method == "GET":
        # Show the initial Protect PDF page
        return render_template("protect.html")
    else:
        # Handle uploaded file
        file = request.files["pdf_file"]
        filename = secure_filename(file.filename)
        filepath = os.path.join(app.config["UPLOAD_FOLDER"], filename)
        file.save(filepath)
        return render_template("protect_upload.html", filename=filename)

@app.route("/protect/process", methods=["POST"])
def protect_process():
    filename = request.form["filename"]
    password1 = request.form["password1"]
    password2 = request.form["password2"]

    if password1 != password2:
        return "Passwords do not match."

    input_path = os.path.join(app.config["UPLOAD_FOLDER"], filename)
    output_path = os.path.join(app.config["UPLOAD_FOLDER"], f"protected_{filename}")

    reader = PdfReader(input_path)
    writer = PdfWriter()

    for page in reader.pages:
        writer.add_page(page)

    writer.encrypt(password1)

    with open(output_path, "wb") as f:
        writer.write(f)

    return send_file(output_path, as_attachment=True, download_name=f"protected_{filename}")



@app.route("/compressimage")
def compressimage():
    """
    Show the upload page for image compression.
    """
    return render_template("compressimage.html")


@app.route("/compressimage_uploaded", methods=["POST"])
def compressimage_uploaded():
    """
    Handle the uploaded image and prepare compression preview.
    """
    file = request.files.get("images")
    if not file:
        return "No file uploaded.", 400

    # Save uploaded image
    filename = secure_filename(file.filename)
    img_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(img_path)

    # Generate thumbnail (200px wide)
    img = Image.open(img_path)
    img.thumbnail((200, 200))
    thumb_filename = f"{filename}_thumb.jpg"
    thumb_path = os.path.join(THUMBNAIL_FOLDER, thumb_filename)
    img.save(thumb_path, "JPEG")

    # File info for template
    file_info = {
        "filename": filename,
        "thumbnail": f"thumbnails/{thumb_filename}",
        "size": f"{round(os.path.getsize(img_path) / 1024, 1)} KB"
    }

    return render_template("compressimage_result.html", file=file_info)


@app.route("/compressimage_download", methods=["POST"])
def compressimage_download():
    """
    Compress the image based on selected compression level and return it.
    """
    filename = request.form.get("filename")
    compression_level = request.form.get("compression_level")

    if not filename:
        return "No file selected.", 400

    input_path = os.path.join(UPLOAD_FOLDER, filename)
    output_path = os.path.join(COMPRESS_FOLDER, f"compressed_{filename}")

    compress_image(input_path, output_path, compression_level)

    return send_file(output_path, as_attachment=True)


def compress_image(input_path, output_path, compression_level):
    """
    Compress an image by resizing and lowering quality.
    """
    compress_map = {
        "extreme": (40, 0.5),      # Lowest quality & half size
        "recommended": (70, 0.75), # Medium quality
        "less": (90, 1.0)          # High quality
    }
    quality, scale = compress_map.get(compression_level, (70, 0.75))

    img = Image.open(input_path)

    new_size = (
        int(img.width * scale),
        int(img.height * scale)
    )

    # Pillow >=10 compatibility
    try:
        resample = Image.Resampling.LANCZOS
    except AttributeError:
        resample = Image.LANCZOS

    img_resized = img.resize(new_size, resample)

    img_resized.save(output_path, "JPEG", quality=quality)

@app.route("/jpgtopdf")
def jpgtopdf():
    """
    Display the JPG to PDF upload page.
    """
    return render_template("jpgtopdf.html")

@app.route("/jpgtopdf_uploaded", methods=["POST"])
def jpgtopdf_uploaded():
    """
    Handle image upload, convert to A4 PDF, and generate thumbnail.
    """
    file = request.files.get("image")
    if not file:
        return "No file uploaded.", 400

    # Save uploaded image
    filename = secure_filename(file.filename)
    img_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(img_path)

    # Convert to A4 PDF without layout_fun
    pdf_filename = f"{os.path.splitext(filename)[0]}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

    with open(pdf_path, "wb") as f:
        f.write(
            img2pdf.convert(
                img_path,
                pagesize=(
                    img2pdf.in_to_pt(8.27),
                    img2pdf.in_to_pt(11.69)
                ),
                auto_orient=True
            )
        )

    # Generate thumbnail
    thumb_filename = f"{os.path.splitext(filename)[0]}_thumb.jpg"
    thumb_path = os.path.join(THUMBNAIL_FOLDER, thumb_filename)

    img = Image.open(img_path)
    img.thumbnail((200, 200))
    img = img.convert("RGB")
    img.save(thumb_path, "JPEG")

    return render_template(
        "jpgtopdf_result.html",
        filename=filename,
        thumbnail=f"thumbnails/{thumb_filename}",
        pdf_filename=pdf_filename
    )


@app.route("/download_jpg_pdf/<filename>", methods=["POST"])
def download_jpg_pdf(filename):
    """
    Serve the generated PDF for download.
    """
    pdf_path = os.path.join(PDF_FOLDER, filename)
    return send_file(pdf_path, as_attachment=True)


@app.route("/wordtopdf")
def wordtopdf():
    return render_template("wordtopdf.html")

@app.route("/wordtopdf_uploaded", methods=["POST"])
def wordtopdf_uploaded():
    file = request.files.get("wordfile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    doc_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(doc_path)

    return render_template(
        "wordtopdf_result.html",
        filename=filename
    )

@app.route("/convert_word_to_pdf/<filename>", methods=["POST"])
def convert_word_to_pdf(filename):
    doc_path = os.path.join(UPLOAD_FOLDER, filename)
    pdf_filename = f"{os.path.splitext(filename)[0]}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

    converted = False

    # Try docx2pdf
    try:
        print("Trying conversion with docx2pdf...")
        docx2pdf_convert(doc_path, pdf_path)
        converted = True
        print("docx2pdf conversion successful.")
    except Exception as e:
        print("docx2pdf failed:", e)

    # Try soffice in PATH
    if not converted:
        try:
            print("Trying conversion with soffice from PATH...")
            subprocess.run([
                "soffice",
                "--headless",
                "--convert-to",
                "pdf",
                "--outdir",
                PDF_FOLDER,
                doc_path
            ], check=True, capture_output=True, text=True)
            converted = True
            print("LibreOffice soffice conversion successful.")
        except Exception as e:
            print("LibreOffice PATH conversion failed:", e)

    # Try soffice default path
    if not converted:
        try:
            soffice_default_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
            if os.path.exists(soffice_default_path):
                print("Trying conversion with soffice default path...")
                subprocess.run([
                    soffice_default_path,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    PDF_FOLDER,
                    doc_path
                ], check=True, capture_output=True, text=True)
                converted = True
                print("LibreOffice default path conversion successful.")
        except Exception as e:
            print("LibreOffice default path conversion failed:", e)

    if not converted:
        return "Conversion failed. Please ensure either Microsoft Word or LibreOffice is installed.", 500

    return send_file(pdf_path, as_attachment=True)


@app.route("/powerpointtopdf")
def powerpointtopdf():
    return render_template("powerpointtopdf.html")

@app.route("/powerpointtopdf_uploaded", methods=["POST"])
def powerpointtopdf_uploaded():
    file = request.files.get("pptfile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    ppt_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(ppt_path)

    return render_template(
        "powerpointtopdf_result.html",
        filename=filename
    )

@app.route("/convert_powerpoint_to_pdf/<filename>", methods=["POST"])
def convert_powerpoint_to_pdf(filename):
    ppt_path = os.path.join(UPLOAD_FOLDER, filename)
    pdf_filename = f"{os.path.splitext(filename)[0]}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

    converted = False

    # Try LibreOffice (soffice) first (recommended for PPT/PPTX)
    try:
        print("Trying conversion with soffice from PATH...")
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            PDF_FOLDER,
            ppt_path
        ], check=True, capture_output=True, text=True)
        converted = True
        print("LibreOffice soffice conversion successful.")
    except Exception as e:
        print("LibreOffice PATH conversion failed:", e)

    # Try soffice default install directory
    if not converted:
        try:
            soffice_default_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
            if os.path.exists(soffice_default_path):
                print("Trying conversion with soffice default path...")
                subprocess.run([
                    soffice_default_path,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    PDF_FOLDER,
                    ppt_path
                ], check=True, capture_output=True, text=True)
                converted = True
                print("LibreOffice default path conversion successful.")
        except Exception as e:
            print("LibreOffice default path conversion failed:", e)

    if not converted:
        return "Conversion failed. Please ensure LibreOffice is installed.", 500

    return send_file(pdf_path, as_attachment=True)



@app.route("/exceltopdf")
def exceltopdf():
    return render_template("exceltopdf.html")

@app.route("/exceltopdf_uploaded", methods=["POST"])
def exceltopdf_uploaded():
    file = request.files.get("excelfile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    excel_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(excel_path)

    return render_template(
        "exceltopdf_result.html",
        filename=filename
    )

@app.route("/convert_excel_to_pdf/<filename>", methods=["POST"])
def convert_excel_to_pdf(filename):
    excel_path = os.path.join(UPLOAD_FOLDER, filename)
    pdf_filename = f"{os.path.splitext(filename)[0]}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

    converted = False

    # 1️⃣ Try LibreOffice in PATH
    try:
        print("Trying conversion with soffice from PATH...")
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            PDF_FOLDER,
            excel_path
        ], check=True, capture_output=True, text=True)
        converted = True
        print("LibreOffice soffice conversion successful.")
    except Exception as e:
        print("LibreOffice PATH conversion failed:", e)

    # 2️⃣ Try LibreOffice default install directory
    if not converted:
        try:
            soffice_default_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
            if os.path.exists(soffice_default_path):
                print("Trying conversion with soffice default path...")
                subprocess.run([
                    soffice_default_path,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    PDF_FOLDER,
                    excel_path
                ], check=True, capture_output=True, text=True)
                converted = True
                print("LibreOffice default path conversion successful.")
        except Exception as e:
            print("LibreOffice default path conversion failed:", e)

    # 3️⃣ Fallback to Microsoft Excel COM API
    if not converted:
        try:
            print("Trying conversion with Microsoft Excel COM...")
            import win32com.client

            excel = win32com.client.Dispatch("Excel.Application")
            excel.Visible = False

            wb = excel.Workbooks.Open(os.path.abspath(excel_path))
            wb.ExportAsFixedFormat(0, os.path.abspath(pdf_path))
            wb.Close(False)
            excel.Quit()

            converted = True
            print("Microsoft Excel COM conversion successful.")
        except Exception as e:
            print("Microsoft Excel COM conversion failed:", e)

    if not converted:
        return "Conversion failed. Please ensure LibreOffice or Microsoft Office is installed.", 500

    return send_file(pdf_path, as_attachment=True)


@app.route("/htmltopdf")
def htmltopdf():
    return render_template("htmltopdf.html")


@app.route("/htmltopdf_uploaded", methods=["POST"])
def htmltopdf_uploaded():
    file = request.files.get("htmlfile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    html_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(html_path)

    return render_template(
        "htmltopdf_result.html",
        filename=filename
    )


@app.route("/convert_html_to_pdf/<filename>", methods=["POST"])
def convert_html_to_pdf(filename):
    html_path = os.path.join(UPLOAD_FOLDER, filename)
    pdf_filename = f"{os.path.splitext(filename)[0]}.pdf"
    pdf_path = os.path.join(PDF_FOLDER, pdf_filename)

    converted = False

    # 1️⃣ Try LibreOffice in PATH
    try:
        print("Trying conversion with soffice from PATH...")
        subprocess.run([
            "soffice",
            "--headless",
            "--convert-to",
            "pdf",
            "--outdir",
            PDF_FOLDER,
            html_path
        ], check=True, capture_output=True, text=True)
        converted = True
        print("LibreOffice soffice conversion successful.")
    except Exception as e:
        print("LibreOffice PATH conversion failed:", e)

    # 2️⃣ Try LibreOffice default install directory
    if not converted:
        try:
            soffice_default_path = r"C:\Program Files\LibreOffice\program\soffice.exe"
            if os.path.exists(soffice_default_path):
                print("Trying conversion with soffice default path...")
                subprocess.run([
                    soffice_default_path,
                    "--headless",
                    "--convert-to",
                    "pdf",
                    "--outdir",
                    PDF_FOLDER,
                    html_path
                ], check=True, capture_output=True, text=True)
                converted = True
                print("LibreOffice default path conversion successful.")
        except Exception as e:
            print("LibreOffice default path conversion failed:", e)

    if not converted:
        return "Conversion failed. Please ensure LibreOffice is installed.", 500

    return send_file(pdf_path, as_attachment=True)


@app.route("/pdftojpg")
def pdftojpg():
    """Display the PDF to JPG upload page."""
    return render_template("pdftojpg.html")

@app.route("/pdftojpg_uploaded", methods=["POST"])
def pdftojpg_uploaded():
    """Handle PDF upload, convert to JPG images, generate thumbnails."""
    file = request.files.get("pdffile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    # Convert PDF to JPG
    doc = fitz.open(pdf_path)
    jpg_files = []

    for page_num in range(doc.page_count):
        page = doc.load_page(page_num)
        pix = page.get_pixmap(dpi=150)

        # Save main image
        jpg_filename = f"{os.path.splitext(filename)[0]}_page{page_num + 1}.jpg"
        jpg_path = os.path.join(JPG_FOLDER, jpg_filename)
        pix.save(jpg_path)

        # Generate thumbnail
        thumb_filename = f"{os.path.splitext(filename)[0]}_page{page_num + 1}_thumb.jpg"
        thumb_path = os.path.join(THUMBNAIL_FOLDER, thumb_filename)

        img = Image.open(jpg_path)
        img.thumbnail((200, 200))
        img.save(thumb_path, "JPEG")

        jpg_files.append({
            "filename": jpg_filename,
            "thumbnail": f"thumbnails/{thumb_filename}"
        })

    return render_template(
        "pdftojpg_result.html",
        filename=filename,
        jpg_files=jpg_files
    )

@app.route("/download_jpg_image/<filename>")
def download_jpg_image(filename):
    """Serve a single JPG file from jpg_outputs with correct MIME type."""
    try:
        return send_from_directory(
            directory=JPG_FOLDER,
            path=filename,
            as_attachment=True,
            mimetype="image/jpeg"
        )
    except FileNotFoundError:
        abort(404, description="JPG file not found.")

@app.route("/download_all_jpg_zip/<filename>")
def download_all_jpg_zip(filename):
    """Zip all JPGs from this PDF and allow ZIP download."""
    base_name = os.path.splitext(filename)[0]
    zip_filename = f"{base_name}_images.zip"
    zip_path = os.path.join(JPG_FOLDER, zip_filename)

    # Create ZIP
    with zipfile.ZipFile(zip_path, "w") as zipf:
        for f in os.listdir(JPG_FOLDER):
            if f.startswith(base_name) and f.endswith(".jpg"):
                zipf.write(
                    os.path.join(JPG_FOLDER, f),
                    arcname=f
                )

    try:
        return send_from_directory(
            directory=JPG_FOLDER,
            path=zip_filename,
            as_attachment=True,
            mimetype='application/zip'
        )
    except FileNotFoundError:
        abort(404, description="ZIP file not found.")

# Optional PDF test
if __name__ == "__main__":
    test_file = "yourfile.pdf"
    if os.path.exists(test_file):
        doc = fitz.open(test_file)
        print("Pages:", doc.page_count)
    else:
        print(f"Test file '{test_file}' does not exist. Skipping PyMuPDF test.")



# 📤 Upload form
@app.route("/pdftoword")
def pdftoword():
    return render_template("pdftoword.html")

# 📥 Handle upload
@app.route("/pdftoword_uploaded", methods=["POST"])
def pdftoword_uploaded():
    file = request.files.get("pdffile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    return render_template(
        "pdftoword_result.html",
        filename=filename,
        filesize=get_file_size(pdf_path)
    )

# 🔁 Convert PDF → Word
@app.route("/convert_pdf_to_word/<filename>", methods=["POST"])
def convert_pdf_to_word(filename):
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    word_filename = os.path.splitext(filename)[0] + ".docx"
    word_path = os.path.join(WORD_FOLDER, word_filename)

    print(f"🔍 PDF: {pdf_path}")
    print(f"📁 Word output: {word_path}")

    try:
        # ✅ Ensure the output directory exists again before saving
        os.makedirs(os.path.dirname(word_path), exist_ok=True)

        print("🚀 Converting...")
        cv = Converter(pdf_path)
        cv.convert(word_path, start=0, end=None)
        cv.close()

        if not os.path.exists(word_path):
            raise FileNotFoundError(f"Word file was not created at: {word_path}")

        print("✅ Conversion successful.")
    except Exception as e:
        print("❌ Conversion failed:")
        traceback.print_exc()
        return f"Conversion failed: {str(e)}", 500

    return send_file(word_path, as_attachment=True)

# ⬇ Download Word file
@app.route("/download_word/<filename>")
def download_word(filename):
    try:
        return send_from_directory(
            directory=WORD_FOLDER,
            path=filename,
            as_attachment=True,
            mimetype="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )
    except FileNotFoundError:
        abort(404, description="Converted DOCX not found.")

# 📏 File size utility
def get_file_size(path):
    size_bytes = os.path.getsize(path)
    return f"{round(size_bytes / 1024, 2)} KB"





@app.route("/pdftoexcel")
def pdftoexcel():
    return render_template("pdftoexcel.html")

@app.route("/pdftoexcel_uploaded", methods=["POST"])
def pdftoexcel_uploaded():
    file = request.files.get("pdffile")
    if not file:
        return "No file uploaded.", 400

    filename = secure_filename(file.filename)
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    file.save(pdf_path)

    # Get file size
    file_size = round(os.path.getsize(pdf_path) / 1024, 2)
    size_label = f"{file_size} KB" if file_size < 1024 else f"{round(file_size / 1024, 2)} MB"

    return render_template("pdftoexcel_result.html", filename=filename, filesize=size_label)

@app.route("/convert_pdf_to_excel/<filename>", methods=["POST"])
def convert_pdf_to_excel(filename):
    pdf_path = os.path.join(UPLOAD_FOLDER, filename)
    excel_filename = f"{os.path.splitext(filename)[0]}.xlsx"
    excel_path = os.path.join(EXCEL_FOLDER, excel_filename)

    try:
        with pdfplumber.open(pdf_path) as pdf:
            all_tables = []
            for page in pdf.pages:
                tables = page.extract_tables()
                for table in tables:
                    df = pd.DataFrame(table)
                    all_tables.append(df)

            if not all_tables:
                return "No tables found in the PDF file.", 400

            with pd.ExcelWriter(excel_path, engine='openpyxl') as writer:
                for i, df in enumerate(all_tables):
                    df.to_excel(writer, sheet_name=f"Page_{i+1}", index=False)

    except Exception as e:
        return f"Conversion failed: {str(e)}", 500

    return send_file(excel_path, as_attachment=True)



@app.route("/ourstory")
def our_story():
    return render_template("ourstory.html")

blog_posts = [
    {
        "id": 1,
        "date": "Jul 18, 2025",
        "title": "How to delete a page in Word (without the formatting headache)",
        "short_title": "Delete a page in Word",
        "summary": "Struggling with extra pages in Word? Here's how to remove them easily—plus a faster way using ILOVEIMG tools.",
        "image": "images/blog1.png",
        "link": "/blog/delete-page-word"
    },
    {
        "id": 2,
        "date": "Jul 17, 2025",
        "title": "Best ways to compress PDF without losing quality",
        "short_title": "Compress PDF without losing quality",
        "summary": "Need to email a large PDF but it’s too big? Learn how to compress your PDF files without losing any important detail.",
        "image": "images/blog2.png",
        "link": "/blog/compress-pdf-quality"
    }
]

@app.route("/daily_blog")
def daily_blog():
    return render_template("daily_blog.html", posts=blog_posts)

@app.route("/blog/<slug>")
def blog_post(slug):
    return f"<h2>Blog post: {slug}</h2><p>Coming soon!</p>"


@app.route("/legal_privacy")
def legal_privacy():
    return render_template("legal_privacy.html")

@app.route("/privacypolicy")
def privacy_policy():
    return render_template("privacypolicy.html")


@app.route("/aboutus")
def about_us():
    return render_template("aboutus.html")






@app.route("/contactus", methods=["GET", "POST"])
def contact_us():
    if request.method == "POST":
        name = request.form.get("name")
        email = request.form.get("email")
        subject = request.form.get("subject")
        message = request.form.get("message")

        # ✅ Optional: Save to DB, send email, or log
        print(f"New message from {name} ({email}): {subject}\n{message}")

        flash("Thank you! Your message has been received.", "success")
        return redirect("/contact")

    return render_template("contactus.html")




@app.route("/disclaimer")
def disclaimer():
    return render_template("disclaimer.html")



@app.route("/features")
def features():
    return render_template("features.html")



@app.route("/faq")
def faq():
    return render_template("faq.html")






 
if __name__ == "__main__":
    app.run(debug=True)