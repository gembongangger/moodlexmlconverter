from flask import Flask, render_template, request, send_from_directory
import os
import uuid
from converter import create_moodle_xml

BASE_DIR = "/home/gembonganggeredu/mysite"

UPLOAD_FOLDER = os.path.join(BASE_DIR, "uploads")
OUTPUT_FOLDER = os.path.join(BASE_DIR, "outputs")

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")

        if not file or not file.filename.lower().endswith(".docx"):
            return "Upload file DOCX ya 🙂"

        # Rename otomatis
        new_name = f"template_{uuid.uuid4().hex}.docx"
        upload_path = os.path.join(UPLOAD_FOLDER, new_name)
        file.save(upload_path)

        # Output XML
        out_name = new_name.replace(".docx", ".xml")
        output_path = os.path.join(OUTPUT_FOLDER, out_name)

        # Konversi
        create_moodle_xml(upload_path, output_path)

        # DEBUG (cek di error log PythonAnywhere)
        print("UPLOAD:", upload_path)
        print("OUTPUT:", output_path)
        print("EXISTS:", os.path.exists(output_path))

        # Download
        return send_from_directory(
            OUTPUT_FOLDER,
            out_name,
            as_attachment=True
        )

    return render_template("index.html")
