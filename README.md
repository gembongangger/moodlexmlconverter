# Word to Moodle XML Converter (Heuristic)

A web-based tool to convert Microsoft Word (.docx) documents into Moodle XML format. This converter uses heuristic patterns to identify questions, options, images, and tables, supporting both legacy rigid tags and natural numbering formats.

## Features

- **PDF Auto-Recognition**: Automatically converts uploaded PDF files to DOCX before processing.

- **Smart Detection**: Recognizes question numbers (1., Soal 1, etc.), options (A., [opsi A]), and answer keys automatically.
- **Media Support**: Automatically extracts and embeds images into the XML.
- **Table Formatting**: Converts Word tables into clean HTML tables with borders compatible with Moodle.
- **Math Support**: Supports LaTeX/MathJax equations.
- **Summary Report**: Displays a detailed conversion summary (Total questions, images, tables, etc.) on both CLI and Web UI.
- **Backward Compatible**: Supports legacy tags like [soal no 1] and [opsi A].

## Requirements

- Python 3.11+
- Pandoc installed on your system.

## Installation

1. Clone or download this project.
2. Create and activate a virtual environment (optional but recommended):
   conda create -n word-xml python=3.13
   conda activate word-xml
3. Install dependencies:
   pip install -r requirements.txt

## Usage

### Web Application
To start the web interface:
python app.py
Open your browser and navigate to http://localhost:5000.

### Command Line
To convert a file directly via terminal:
python converter.py path/to/your/file.docx output_name.xml

## Supported Question Formats

### Format 1: Natural (Heuristic)
1. What is the capital of Indonesia?
A. Jakarta
B. Bandung
C. Surabaya
Jawaban: A

### Format 2: Rigid Tags (Legacy)
[soal no 1]
What is the capital of Indonesia?
[opsi A] Jakarta
[opsi B] Bandung
[opsi C] Surabaya
[KUNCI: A]

## Project Structure

- app.py: Main Flask application.
- converter.py: Core conversion logic with heuristic regex.
- templates/: HTML templates for the web UI.
- uploads/: Temporary storage for uploaded Word files.
- outputs/: Storage for generated Moodle XML files.

## License
MIT
# moodlexmlconverter
