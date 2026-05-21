import re
import base64
import os
import pypandoc
import xml.etree.ElementTree as ET
from xml.dom import minidom
from docx import Document
import zipfile

def get_docx_images(docx_path):
    images = {}
    try:
        with zipfile.ZipFile(docx_path) as z:
            for file in z.namelist():
                if file.startswith('word/media/'):
                    img_name = os.path.basename(file)
                    images[img_name] = z.read(file)
    except Exception as e:
        pass
    return images

def clean_html(html):
    if not html: return ""
    html = re.sub(r'^\s*<\/(?:p|li|h[1-6]|ul|ol|strong|span|em|i|b)>', '', html, flags=re.IGNORECASE)
    html = re.sub(r'<(?:p|li|h[1-6]|ul|ol|strong|span|em|i|b)[^>]*>\s*$', '', html, flags=re.IGNORECASE)
    html = re.sub(r'<(p|li|ul|ol|span|strong)[^>]*>\s*<\/\1>', '', html, flags=re.IGNORECASE)
    return html.strip()

def create_moodle_xml(docx_path, output_path):
    print(f"\n[1/4] Membaca file: {os.path.basename(docx_path)}...")
    
    stats = {
        "total_soal": 0,
        "pg_tunggal": 0,
        "pg_kompleks": 0,
        "total_gambar": 0,
        "total_tabel": 0,
        "category": "Bank Soal",
        "status": "GAGAL"
    }

    try:
        full_html = pypandoc.convert_file(docx_path, 'html', format='docx', extra_args=['--mathjax', '--wrap=none'])
    except Exception as e:
        print(f"ERROR: Gagal menjalankan Pandoc: {e}")
        return stats

    images_binary = get_docx_images(docx_path)
    
    try:
        doc = Document(docx_path)
        category_name = doc.paragraphs[0].text.strip() if doc.paragraphs else "Bank Soal"
    except:
        category_name = "Bank Soal"
    
    stats["category"] = category_name

    quiz = ET.Element("quiz")
    cat_q = ET.SubElement(quiz, "question", type="category")
    cat_node = ET.SubElement(cat_q, "category")
    ET.SubElement(cat_node, "text").text = f"$course$/{category_name}"

    soal_pattern = r'<(p|li|h[1-6])[^>]*>(?:\s*<(?:strong|span|em)[^>]*>)*\s*(?:\[soal\s+no\s+\d+\]|(?:\d+[\.\)])(?:\s|&nbsp;)*|Soal\s*\d+|Nomor\s*\d+)'
    soal_matches = list(re.finditer(soal_pattern, full_html, flags=re.IGNORECASE))
    
    if not soal_matches:
        soal_pattern = r'<(p|li|h[1-6])[^>]*>(?:\s*<(?:strong|span|em)[^>]*>)*\s*(?:\d+[\.\)])(?:\s|&nbsp;)*'
        soal_matches = list(re.finditer(soal_pattern, full_html, flags=re.IGNORECASE))

    print(f"[2/4] Menganalisis konten... (Ditemukan {len(soal_matches)} calon blok soal)")

    for i in range(len(soal_matches)):
        start = soal_matches[i].end()
        end = soal_matches[i+1].start() if i+1 < len(soal_matches) else len(full_html)
        block = full_html[start:end]
        
        opt_pattern = r'(?:\[opsi\s+[a-e]\]|<(?:li|p|h[1-6])[^>]*>(?:\s*<(?:p|strong|span|em)[^>]*>)?\s*[a-e][\.\)](?:\s|&nbsp;)*)'
        opt_matches = list(re.finditer(opt_pattern, block, flags=re.IGNORECASE))
        
        if not opt_matches: continue

        stats["total_soal"] += 1
        
        tag_text = re.sub(r'<[^>]+>', '', soal_matches[i].group(0))
        num_match = re.search(r'(\d+)', tag_text)
        num = num_match.group(1) if num_match else str(stats["total_soal"])

        key_pattern = r'(?:\[kunci\s*:\s*|Kunci\s*(?:Jawaban)?\s*[:\.]\s*|Jawaban\s*[:\.]\s*)(?:<(?:strong|span|p)[^>]*>)?\s*([a-e,\s]+)'
        key_match = re.search(key_pattern, block, flags=re.IGNORECASE)
        keys = []
        if key_match:
            raw_key = re.sub(r'<[^>]+>', '', key_match.group(1))
            keys = [k.strip().upper() for k in raw_key.split(',')]

        if len(keys) > 1:
            stats["pg_kompleks"] += 1
        else:
            stats["pg_tunggal"] += 1

        question_html_raw = block[:opt_matches[0].start()]
        options = []
        for j in range(len(opt_matches)):
            opt_start = opt_matches[j].end()
            lookahead = r'(?:\[opsi\s+[a-e]\]|<(?:li|p|h[1-6])[^>]*>(?:\s*<(?:p|strong|span|em)[^>]*>)?\s*[a-e][\.\)]|\[kunci\s*:|Kunci\s*(?:Jawaban)?\s*[:\.]|Jawaban\s*[:\.]|Pembahasan)'
            next_search = block[opt_start:]
            opt_end_match = re.search(lookahead, next_search, flags=re.IGNORECASE)
            opt_end = opt_start + (opt_end_match.start() if opt_end_match else len(next_search))
            options.append(block[opt_start:opt_end])

        q = ET.SubElement(quiz, "question", type="multichoice")
        ET.SubElement(ET.SubElement(q, "name"), "text").text = f"Soal {num}"
        qtext_node = ET.SubElement(q, "questiontext", format="html")
        
        processed_qtext = clean_html(question_html_raw)
        stats["total_tabel"] += len(re.findall(r'<table', processed_qtext, flags=re.IGNORECASE))
        img_tags = re.findall(r'<img [^>]*src="([^"]+)"', processed_qtext)
        for img_src in img_tags:
            base_name = os.path.basename(img_src)
            if base_name in images_binary:
                stats["total_gambar"] += 1
                f_tag = ET.SubElement(qtext_node, "file", name=base_name, path="/", encoding="base64")
                f_tag.text = base64.b64encode(images_binary[base_name]).decode('utf-8')
                processed_qtext = processed_qtext.replace(img_src, f'@@PLUGINFILE@@/{base_name}')
        
        processed_qtext = re.sub(r'<table', r'<table border="1" style="border-collapse: collapse; width: 100%;"', processed_qtext)
        processed_qtext = re.sub(r'<td', r'<td style="border: 1px solid black; padding: 5px;"', processed_qtext)
        processed_qtext = re.sub(r'<th', r'<th style="border: 1px solid black; padding: 5px;"', processed_qtext)
        ET.SubElement(qtext_node, "text").text = processed_qtext

        num_keys = len(keys)
        is_single = num_keys <= 1
        ET.SubElement(q, "single").text = "true" if is_single else "false"
        ET.SubElement(q, "answernumbering").text = "abc"
        ET.SubElement(q, "shuffleanswers").text = "1"

        for i, opt_raw in enumerate(options):
            label = chr(65 + i)
            correct = label in keys
            fraction = str(100 / max(1, num_keys)) if correct else ("0" if is_single else str((-1)*100 / (5-max(1, num_keys))))
            ans_node = ET.SubElement(q, "answer", fraction=fraction, format="html")
            clean_opt = re.sub(r'^.*?\[opsi\s+[a-e]\]', '', opt_raw, flags=re.IGNORECASE)
            if clean_opt == opt_raw:
                clean_opt = re.sub(r'^\s*[a-e][\.\)]\s*', '', opt_raw, flags=re.IGNORECASE)
            clean_opt = clean_html(clean_opt)
            ET.SubElement(ans_node, "text").text = f"<span>{clean_opt}</span>"
            ET.SubElement(ET.SubElement(ans_node, "feedback"), "text").text = "Benar" if correct else "Salah"

    print(f"[3/4] Membuat file XML...")
    xml_data = ET.tostring(quiz, encoding='utf-8')
    pretty_xml = minidom.parseString(xml_data).toprettyxml(indent="  ")
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(pretty_xml)
    
    stats["status"] = "BERHASIL"
    print("[4/4] Konversi Selesai.")

    # Tampilkan laporan di layar
    print("\n" + "="*45)
    print("        RINGKASAN HASIL KONVERSI")
    print("="*45)
    print(f" File Input      : {os.path.basename(docx_path)}")
    print(f" Kategori Moodle : {stats['category']}")
    print("-" * 45)
    print(f" Total Soal      : {stats['total_soal']}")
    print(f"  - PG Tunggal   : {stats['pg_tunggal']}")
    print(f"  - PG Kompleks  : {stats['pg_kompleks']}")
    print(f" Total Gambar    : {stats['total_gambar']}")
    print(f" Total Tabel     : {stats['total_tabel']}")
    print("-" * 45)
    print(f" STATUS          : {stats['status']}")
    print(f" File Output     : {os.path.basename(output_path)}")
    print("="*45 + "\n")
    
    return stats

if __name__ == "__main__":
    import sys
    FILE_INPUT = sys.argv[1] if len(sys.argv) > 1 else "template.docx"
    FILE_OUTPUT = sys.argv[2] if len(sys.argv) > 2 else "hasil_moodle.xml"
    if os.path.exists(FILE_INPUT):
        create_moodle_xml(FILE_INPUT, FILE_OUTPUT)
    else:
        print(f"Error: File {FILE_INPUT} tidak ditemukan.")
