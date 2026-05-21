from flask import Flask, render_template, request, send_from_directory
import os
import uuid
from converter import create_moodle_xml

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

app = Flask(__name__)

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")

        if not file or not file.filename.lower().endswith(".docx"):
            return "Upload file DOCX ya 🙂"

        # Nama file unik
        file_id = uuid.uuid4().hex
        new_name = f"template_{file_id}.docx"
        upload_path = os.path.join(UPLOAD_FOLDER, new_name)
        file.save(upload_path)

        # Path Output
        out_name = new_name.replace(".docx", ".xml")
        output_path = os.path.join(OUTPUT_FOLDER, out_name)

        # Jalankan Konversi & Ambil Statistik
        stats = create_moodle_xml(upload_path, output_path)
        stats["filename"] = out_name
        stats["original_name"] = file.filename

        return render_template("index.html", stats=stats)

    return render_template("index.html")

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
