# Word to Moodle XML Converter (Heuristic)

Aplikasi web untuk mengonversi soal dari Microsoft Word (.docx) ke format XML Moodle. Mendukung **PG Tunggal**, **PG Kompleks**, dan **Benar/Salah** dengan upload kunci jawaban dari file Excel.

## Fitur

- **3 Tipe Soal**: PG Tunggal (multichoice single), PG Kompleks (multichoice multiple), Benar/Salah (truefalse)
- **Upload Kunci Jawaban**: Upload file Excel kunci jawaban, override kunci otomatis
- **Auto-detect Format**: Mendukung Word auto-numbering (pandoc `<ol>`) dan format legacy/regex
- **Gambar**: Ekstrak dan embed gambar dari DOCX ke XML (base64)
- **MathJax/LaTeX**: Support persamaan matematika
- **PDF**: Upload PDF otomatis dikonversi ke DOCX via pdf2docx
- **PAKET A/B**: Pilih paket saat upload kunci

## Persyaratan

- **Python** 3.11+
- **Pandoc** — harus terinstall di sistem
- **Pip** (Python package manager)

## Instalasi Pandoc

### Linux (Debian/Ubuntu)
```bash
sudo apt update && sudo apt install pandoc
```

### Linux (Arch)
```bash
sudo pacman -S pandoc
```

### Windows
Download installer dari https://pandoc.org/installing.html, lalu tambahkan ke PATH.

Cek instalasi:
```bash
pandoc --version
```

## Instalasi Aplikasi

### 1. Clone repository
```bash
git clone https://github.com/gembongangger/moodlexmlconverter.git
cd moodlexmlconverter
```

### 2. Buat virtual environment

#### Linux / macOS
```bash
python3 -m venv venv
source venv/bin/activate
```

#### Windows (Command Prompt)
```cmd
python -m venv venv
venv\Scripts\activate
```

#### Windows (PowerShell)
```powershell
python -m venv venv
venv\Scripts\Activate.ps1
```

#### Alternatif dengan Conda (Linux/Windows)
```bash
conda create -n moodle-converter python=3.13
conda activate moodle-converter
```

### 3. Install dependencies
```bash
pip install -r requirements.txt
```

## Menjalankan Aplikasi

```bash
python app.py
```

Buka browser ke **http://localhost:5000**

## Cara Penggunaan

1. Upload file **DOCX** (atau PDF) berisi soal
2. Akan muncul halaman **Review** — centang soal yang ingin dimasukkan
3. (Opsional) Upload **file Excel kunci jawaban**, pilih PAKET A atau B
4. Klik **Konfirmasi & Buat XML**
5. Download file XML dan import ke Moodle

## Format Kunci Jawaban Excel

| Kolom | Isi |
|-------|-----|
| Kolom 2-9 | PG Tunggal (PAKET A, baris genap/ganjil) |
| Kolom 10-15 | PG Kompleks |
| Kolom 16-20 | Benar/Salah (B = Benar, S = Salah) |

Lihat file contoh: `KUNCI_JAWABAN_MAT_TL_(40_SOAL) FIX.xlsx`

## Struktur Proyek

```
.
├── app.py              # Flask web application
├── converter.py        # Core conversion logic
├── requirements.txt    # Python dependencies
├── .gitignore
├── templates/
│   ├── index.html      # Halaman upload utama
│   └── review.html     # Halaman review soal + upload kunci
├── uploads/            # Temporary uploaded files
├── outputs/            # Generated XML files
└── data/               # Session cache (JSON)
```

## Lisensi

MIT
