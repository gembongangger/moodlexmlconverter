import re
from flask import Flask, render_template, request, send_from_directory, redirect, url_for
import os
import uuid
import shutil
import json
from converter import extract_questions_dict, save_to_xml, parse_keys_from_excel, merge_keys
from pdf2docx import Converter

UPLOAD_FOLDER = "uploads"
OUTPUT_FOLDER = "outputs"
DATA_FOLDER = "data"

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)
os.makedirs(DATA_FOLDER, exist_ok=True)

app = Flask(__name__)
app.secret_key = uuid.uuid4().hex

@app.route("/", methods=["GET", "POST"])
def index():
    if request.method == "POST":
        file = request.files.get("file")
        if not file: return "Pilih file dulu ya 🙂"
        
        filename_lower = file.filename.lower()
        if not (filename_lower.endswith(".docx") or filename_lower.endswith(".pdf")):
            return "Hanya dukung file DOCX atau PDF ya 🙂"

        file_id = uuid.uuid4().hex
        ext = ".pdf" if filename_lower.endswith(".pdf") else ".docx"
        temp_name = f"upload_{file_id}{ext}"
        upload_path = os.path.join(UPLOAD_FOLDER, temp_name)
        file.save(upload_path)

        actual_docx_path = upload_path
        if ext == ".pdf":
            actual_docx_path = os.path.join(UPLOAD_FOLDER, f"converted_{file_id}.docx")
            try:
                cv = Converter(upload_path)
                cv.convert(actual_docx_path)
                cv.close()
            except Exception as e: return f"Gagal mengonversi PDF: {e}"

        # Ekstrak data mentah
        data = extract_questions_dict(actual_docx_path)
        data["file_id"] = file_id
        data["original_name"] = file.filename
        
        # Simpan file docx asli/konversi untuk didownload nanti
        docx_out_name = f"soal_{file_id}.docx"
        shutil.copy(actual_docx_path, os.path.join(OUTPUT_FOLDER, docx_out_name))
        data["docx_filename"] = docx_out_name

        # Simpan draft ke JSON
        json_path = os.path.join(DATA_FOLDER, f"{file_id}.json")
        with open(json_path, "w") as f:
            json.dump(data, f)

        return redirect(url_for('review', file_id=file_id))

    return render_template("index.html")

@app.route("/review/<file_id>")
def review(file_id):
    json_path = os.path.join(DATA_FOLDER, f"{file_id}.json")
    if not os.path.exists(json_path): return redirect(url_for('index'))
    with open(json_path, "r") as f:
        data = json.load(f)
    return render_template("review.html", data=data)

@app.route("/upload_keys/<file_id>", methods=["POST"])
def upload_keys(file_id):
    json_path = os.path.join(DATA_FOLDER, f"{file_id}.json")
    if not os.path.exists(json_path): return redirect(url_for('index'))

    file = request.files.get("key_file")
    if not file:
        return "Pilih file kunci dulu ya 🙂", 400

    fn = file.filename.lower()
    if not (fn.endswith(".xlsx") or fn.endswith(".xls")):
        return "Sementara hanya dukung file Excel (.xlsx/.xls) ya. Dukungan Word kunci menyusul 🙂", 400

    paket = request.form.get("paket", "PAKET A")

    tmp_path = os.path.join(UPLOAD_FOLDER, f"keys_{file_id}_{file.filename}")
    file.save(tmp_path)

    with open(json_path, "r") as f:
        data = json.load(f)

    keys = parse_keys_from_excel(tmp_path, paket)

    data["questions"] = merge_keys(data["questions"], keys)
    data["key_filename"] = file.filename

    with open(json_path, "w") as f:
        json.dump(data, f)

    os.unlink(tmp_path)
    return redirect(url_for('review', file_id=file_id))

@app.route("/generate", methods=["POST"])
def generate():
    file_id = request.form.get("file_id")
    selected_indices = request.form.getlist("selected_soal")
    
    json_path = os.path.join(DATA_FOLDER, f"{file_id}.json")
    if not os.path.exists(json_path): return redirect(url_for('index'))
    
    with open(json_path, "r") as f:
        data = json.load(f)
    
    # Filter soal yang dipilih
    selected_questions = [data["questions"][int(idx)] for idx in selected_indices]
    
    # Gunakan images_binary (sudah format base64 string dari JSON)
    images_binary = data.get("images_binary", {})

    final_data = {
        "category": data["category"],
        "questions": selected_questions,
        "images_binary": images_binary
    }
    
    xml_name = f"hasil_{file_id}.xml"
    output_path = os.path.join(OUTPUT_FOLDER, xml_name)
    save_to_xml(final_data, output_path)
    
    # Hitung statistik riil dari soal yang dipilih
    pg_tunggal = sum(1 for q in selected_questions if q.get('type') == 'multichoice' and len(q.get('keys', [])) <= 1)
    pg_kompleks = sum(1 for q in selected_questions if q.get('type') == 'multichoice' and len(q.get('keys', [])) > 1)
    tf_count = sum(1 for q in selected_questions if q.get('type') == 'truefalse')
    total_gambar = sum(len(re.findall(r'@@PLUGINFILE@@', q['text'])) for q in selected_questions)
    total_tabel = sum(len(re.findall(r'<table', q['text'])) for q in selected_questions)

    stats = {
        "total_soal": len(selected_questions),
        "pg_tunggal": pg_tunggal,
        "pg_kompleks": pg_kompleks,
        "truefalse": tf_count,
        "total_gambar": total_gambar,
        "total_tabel": total_tabel,
        "category": data["category"],
        "status": "BERHASIL",
        "xml_filename": xml_name,
        "docx_filename": data.get("docx_filename"),
        "original_name": data["original_name"]
    }
    
    return render_template("index.html", stats=stats)

@app.route("/download/<filename>")
def download_file(filename):
    return send_from_directory(OUTPUT_FOLDER, filename, as_attachment=True)

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
