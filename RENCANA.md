# Rencana Implementasi: Dukungan True/False & Matching + Key File Terpisah

## 1. Tujuan

Menambah kemampuan `converter.py` untuk mendeteksi dan menghasilkan XML Moodle untuk:

- **True/False (Benar-Salah)** — soal dengan opsi Benar/Salah, format sub-pernyataan dalam tabel
- **Matching (Mencocokkan)** — soal pasangan item kiri-kanan
- **Key File Terpisah** — upload kunci jawaban via Excel/Word untuk override/melengkapi kunci otomatis

---

## 2. Data Model Baru

Setiap soal sekarang memiliki field `type`:

### Multichoice (existing)

```python
{
    "type": "multichoice",
    "number": "1",
    "text": "<html>...",
    "options": ["opsi A", "opsi B", ...],
    "keys": ["A"]              # atau ["A", "B"] untuk PG kompleks
}
```

### True/False (baru)

```python
{
    "type": "truefalse",
    "number": "37.1",          # sub-nomor dalam grup
    "parent_number": "37",     # nomor grup asal di dokumen
    "text": "<passage+gambar><br>Vektor (3,–2) adalah d̄",
    "keys": ["B"]              # "B" = Benar, "S" = Salah
}
```

### Matching (baru)

```python
{
    "type": "matching",
    "number": "41",
    "text": "<html>Pasangkan...",
    "pairs": [
        {"text": "Jakarta", "answer": "Indonesia"},
        {"text": "Tokyo",  "answer": "Jepang"}
    ]
}
```

---

## 3. Perubahan di `converter.py`

### 3a. `extract_questions_dict()` — Deteksi Heuristik

**True/False:**

1. Setelah parsing PG biasa (soal 1—32/36), cari blok sisa HTML yang mengandung tabel dengan kolom "BENAR" dan "SALAH"
2. Jika ditemukan:
   - Ambil teks/gambar sebelum tabel sebagai shared passage
   - Parse setiap baris tabel: No | Pernyataan | BENAR | SALAH
   - Tiap baris jadi 1 objek `truefalse` dengan `number = "{parent}.{row}"`
   - Passage di-prepend ke `text` tiap sub-pernyataan
3. Key: jika kolom BENAR dicentang/terisi → `keys: ["B"]`, jika SALAH → `keys: ["S"]`

**Matching:**

1. Deteksi tabel dengan 2 kolom (setelah judul "Cocokkan", "Pasangkan", dll.)
2. Atau pola list dengan separator `→`, `[match]`, atau `:` yang jelas
3. Tiap baris jadi 1 entry dalam `pairs` array

### 3b. Fungsi XML Output Baru

**`save_truefalse_xml(q, parent)`**

```python
q_xml = ET.SubElement(parent, "question", type="truefalse")
# name → f"Soal {q['number']}"
# questiontext → q['text'] (dengan gambar via @@PLUGINFILE@@)
# answer fraction=100 → <text>true</text>  (jika keys=['B'])
# answer fraction=0   → <text>false</text> (jika keys=['S'])
# feedback → "Benar" / "Salah"
```

Tidak perlu tag `<single>`, `<shuffleanswers>`, atau `<answernumbering>`.

**`save_matching_xml(q, parent)`**

```python
q_xml = ET.SubElement(parent, "question", type="matching")
# name, questiontext sama
# setiap pair → <subquestion>
#   <text>pair['text']</text>
#   <answer><text>pair['answer']</text></answer>
# </subquestion>
# <shuffleanswers>true</shuffleanswers>
# <correctfeedback>, <incorrectfeedback>, <partiallycorrectfeedback>
```

**`save_to_xml()` — routing berdasarkan type**

```python
for item in data["questions"]:
    if item["type"] == "truefalse":
        save_truefalse_xml(item, quiz)
    elif item["type"] == "matching":
        save_matching_xml(item, quiz)
    else:
        save_multichoice_xml(item, quiz)  # existing
```

### 3c. Fungsi Key Parser Baru

**`parse_keys_from_excel(excel_path) → dict`**

- Baca `.xlsx` via `openpyxl`
- Deteksi kolom secara fleksibel (header atau posisi):
  - Nomor soal → `"1"`, `"37.1"`, dll.
  - Kunci → `"A"`, `"B"`, `"B,S"`, `"B"` (untuk truefalse: B=Benar, S=Salah)
  - (opsional untuk matching) Kiri | Kanan
- Output: `{"1": ["A"], "37.1": ["B"], ...}`

**`parse_keys_from_docx(docx_path) → dict`**

- Baca `.docx` via `python-docx` (existing)
- Dukung format dalam dokumen:
  - List: `1. A`, `Soal 1: B`, `37.1: B`
  - Tabel: 2 kolom (Nomor | Kunci)
  - EQ-style: `1=A, 2=B`
- Output: dict sama seperti Excel

**`merge_keys(questions, keys_dict) → questions`**

- Override `q["keys"]` untuk setiap soal yang nomornya cocok
- Soal tanpa kecocokan tetap pakai key hasil deteksi otomatis
- Khusus truefalse: validasi key hanya `"B"` atau `"S"`

### 3d. Tambahan: `<correctfeedback>` dkk. untuk Multichoice

Sambil nanti implement, tambahkan tag feedback umum ke `save_multichoice_xml()`:

```python
ET.SubElement(q, "correctfeedback").text = "Jawaban Anda benar."
ET.SubElement(q, "partiallycorrectfeedback").text = "Sebagian benar."
ET.SubElement(q, "incorrectfeedback").text = "Jawaban Anda salah."
```

---

## 4. Perubahan di `app.py`

### Route Baru: `POST /upload_keys/<file_id>`

```
Menerima file .xlsx atau .docx
1. Simpan ke UPLOAD_FOLDER
2. Panggil parse_keys_from_excel/docx
3. Panggil merge_keys(questions, parsed_keys)
4. Update JSON data di DATA_FOLDER
5. Redirect ke /review/<file_id>
```

### Route Existing: `/generate`

Update statistik untuk menangani tipe soal baru:

```python
pg_tunggal = sum(1 for q in selected if q['type']=='multichoice' and len(q['keys'])<=1)
pg_kompleks = sum(1 for q in selected if q['type']=='multichoice' and len(q['keys'])>1)
tf_count = sum(1 for q in selected if q['type']=='truefalse')
matching_count = sum(1 for q in selected if q['type']=='matching')
total_gambar = ...
```

---

## 5. Perubahan di `templates/review.html`

### Form Upload Key File (bagian atas halaman)

```html
<div style="background:#fff3cd; padding:15px; border-radius:8px; margin-bottom:20px;">
  <h3>Upload Kunci Terpisah (Opsional)</h3>
  <form action="/upload_keys/{{ data.file_id }}" method="post" enctype="multipart/form-data">
    <input type="file" name="key_file" accept=".xlsx,.docx">
    <button type="submit">Upload & Merge</button>
  </form>
  <p style="font-size:0.85em; color:#856404;">
    Kunci dari file akan override kunci otomatis. Format: kolom No + Kunci.
  </p>
</div>
```

### Render Kartu Soal per Tipe

- **`multichoice`**: tampilkan seperti sekarang (A/B/C/D/E)
- **`truefalse`**: 2 opsi Benar (hijau) / Salah (merah), highlight sesuai kunci
- **`matching`**: tabel 2 kolom, centang hijau pada pasangan benar

### Badge Sumber Kunci

Tiap kartu soal tampilkan sumber kunci:

```
Kunci: A (dari dokumen) | Kunci: B (dari file kunci)
```

---

## 6. Perubahan di `requirements.txt`

```
flask
pypandoc
python-docx
pdf2docx
openpyxl           # baru — untuk baca Excel
```

---

## 7. Prioritas Implementasi

| # | Fitur | Estimasi |
|---|-------|----------|
| 1 | **True/False detection** di `extract_questions_dict()` | 2-3 jam |
| 2 | **True/False XML output** (`save_truefalse_xml`) | 1 jam |
| 3 | **Update review UI** untuk truefalse | 1 jam |
| 4 | **Key parser: Excel** (`parse_keys_from_excel`) | 2 jam |
| 5 | **Key parser: Word** (`parse_keys_from_docx`) | 1 jam |
| 6 | **Merge keys** + route `/upload_keys` | 1 jam |
| 7 | **Matching detection + XML** (butuh sample file) | 2-3 jam |
| 8 | **Update review UI** untuk matching | 1 jam |
| 9 | `<correctfeedback>` dkk. untuk multichoice | 0.5 jam |
| 10 | Uji coba & perbaikan | 2 jam |

**Total estimasi: ~14-16 jam**

---

## 8. Catatan / Edge Cases

| Situasi | Penanganan |
|---|---|
| Key file kosong / format salah | Flash error, tetap di review page |
| Nomor di key file tidak ada di soal | Abaikan (log saja) |
| PG Kompleks (multi-key) | Pake koma di Excel: `A,B` atau `A B` |
| TrueFalse key `B`/`S` | `B`=Benar (`true`), `S`=Salah (`false`) |
| Soal matching tanpa sample file | Tunda implementasi sampai ada contoh nyata |
| Passage dengan gambar di truefalse | Gambar dari passage di-copy ke tiap sub-pernyataan |
