import re
import base64
import os
import zipfile
from docx import Document
import xml.etree.ElementTree as ET
from xml.dom import minidom

def get_docx_images(docx_path):
    """Mengekstrak gambar dari docx dan menyimpannya dalam dictionary."""
    images = {}
    with zipfile.ZipFile(docx_path) as z:
        for file in z.namelist():
            if file.startswith('word/media/'):
                img_name = os.path.basename(file)
                images[img_name] = z.read(file)
    return images

def create_moodle_xml_with_images(docx_path, output_path):
    doc = Document(docx_path)
    images_data = get_docx_images(docx_path)
    quiz = ET.Element("quiz")

    # Mapping gambar yang muncul di tiap paragraf
    # Docx menyimpan gambar di dalam relasi (rId), namun cara termudah 
    # untuk script sederhana ini adalah mendeteksi keberadaan gambar di paragraf.
    
    full_content = []
    current_soal = None

    for para in doc.paragraphs:
        text = para.text.strip()
        
        # Deteksi awal soal
        if re.search(r'\[soal no \d+\]', text.lower()):
            if current_soal: full_content.append(current_soal)
            current_soal = {'text': '', 'options': [], 'keys': [], 'images': []}
            current_soal['name'] = re.search(r'\[soal no (\d+)\]', text.lower()).group(1)
        
        elif current_soal is not None:
            if text.startswith('isi soal nomor'):
                current_soal['text'] = re.sub(r'isi soal nomor \d+', '', text).strip()
                # Cek apakah ada gambar di paragraf ini
                if 'Graphic' in para._p.xml or 'Drawing' in para._p.xml:
                    # Ambil semua gambar yang tersedia (asumsi satu gambar per soal untuk kestabilan)
                    for img_name in images_data.keys():
                        if img_name not in [img['name'] for img in current_soal['images']]:
                            current_soal['images'].append({'name': img_name, 'data': images_data[img_name]})
                            break
            elif text.startswith('[opsi'):
                opt_text = re.sub(r'\[opsi [A-E]\]', '', text).strip()
                current_soal['options'].append(opt_text)
            elif text.startswith('[KUNCI:'):
                key_match = re.search(r'\[KUNCI:\s*(.*)\]', text)
                if key_match:
                    current_soal['keys'] = [k.strip() for k in key_match.group(1).split(',')]

    if current_soal: full_content.append(current_soal)

    for item in full_content:
        q_type = "multichoice"
        question = ET.SubElement(quiz, "question", type=q_type)
        
        name = ET.SubElement(question, "name")
        ET.SubElement(name, "text").text = f"Soal {item['name']}"

        # Setup Question Text dengan Gambar
        qtext_element = ET.SubElement(question, "questiontext", format="html")
        
        # Jika ada gambar, tambahkan tag <img> yang mengarah ke internal file
        img_html = ""
        for img in item['images']:
            img_html += f'<p><img src="@@PLUGINFILE@@/{img["name"]}" alt="" /></p>'
        
        # Render matematika sederhana ke LaTeX Moodle
        formatted_text = item['text'].replace("2x2", "2x^2").replace("cos2x", "\\cos^2 x")
        ET.SubElement(qtext_element, "text").text = f"<![CDATA[<p>{formatted_text}</p>{img_html}]]>"

        # Masukkan file gambar (Base64)
        for img in item['images']:
            file_tag = ET.SubElement(qtext_element, "file", name=img['name'], path="/", encoding="base64")
            file_tag.text = base64.b64encode(img['data']).decode('utf-8')

        # Konfigurasi Jawaban
        single = "true" if len(item['keys']) <= 1 else "false"
        ET.SubElement(question, "single").text = single
        
        map_keys = {'A': 0, 'B': 1, 'C': 2, 'D': 3, 'E': 4}
        for i, opt_text in enumerate(item['options']):
            current_label = list(map_keys.keys())[i]
            is_correct = current_label in item['keys']
            fraction = str(100/len(item['keys'])) if is_correct else ("0" if single == "true" else "-50")
            
            answer = ET.SubElement(question, "answer", fraction=fraction)
            ET.SubElement(answer, "text").text = f"<![CDATA[{opt_text}]]>"

    # Save
    xml_str = ET.tostring(quiz, encoding='utf-8')
    pretty_xml = minidom.parseString(xml_str).toprettyxml(indent="  ")
    pretty_xml = pretty_xml.replace("&lt;![CDATA[", "<![CDATA[").replace("]]&gt;", "]]>")
    
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(pretty_xml)

if __name__ == "__main__":
    create_moodle_xml_with_images("template.docx", "questions-rev-01.xml")
    print("Konversi Berhasil dengan Gambar!")