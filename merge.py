import os
import re
import sys
import subprocess
import shutil
import yaml
from docx import Document
from docxcompose.composer import Composer
from tqdm import tqdm
from docx.oxml import OxmlElement
from docx.oxml.ns import qn

# --- Load configuration ---
CONFIG_FILE = "config.yaml"

if not os.path.exists(CONFIG_FILE):
    print(f"❌ Configuration file '{CONFIG_FILE}' not found.")
    sys.exit(1)

with open(CONFIG_FILE, "r") as f:
    config = yaml.safe_load(f)

INPUT_DIR = config.get("input_folder", "input")
OUTPUT_DIR = config.get("output_folder", "output")
PAGE_BREAK = config.get("merge_options", {}).get("page_break_between_files", True)
NUMERIC_SORT = config.get("merge_options", {}).get("numeric_sorting", True)
PDF_CONVERT = config.get("pdf_options", {}).get("convert_to_pdf", True)
PDF_FILENAME = config.get("pdf_options", {}).get("pdf_filename", "merged_output.pdf")
DOCX_FILENAME = "merged_output.docx"

# --- Helper functions ---
def numeric_sort(filename):
    numbers = re.findall(r'\d+', filename)
    return int(numbers[0]) if numbers else sys.maxsize

def add_page_break(doc):
    p = OxmlElement('w:p')
    r = OxmlElement('w:r')
    br = OxmlElement('w:br')
    br.set(qn('w:type'), 'page')
    r.append(br)
    p.append(r)
    doc.element.body.append(p)

# --- Folder checks ---
if not os.path.exists(INPUT_DIR):
    print(f"❌ Input folder '{INPUT_DIR}' not found.")
    sys.exit(1)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Load DOCX files ---
files = [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".docx")]
if NUMERIC_SORT:
    files = sorted(files, key=numeric_sort)

if len(files) < 2:
    print("❌ Not enough DOCX files to merge.")
    sys.exit(1)

print(f"📄 Found {len(files)} DOCX files in '{INPUT_DIR}'")
print("🔧 Initializing master document...")

# --- Master document ---
master_path = os.path.join(INPUT_DIR, files[0])
master = Document(master_path)
composer = Composer(master)

# --- Merge remaining files ---
with tqdm(total=len(files) - 1, desc="🔗 Merging documents", unit="file",
          bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]") as pbar:

    for file in files[1:]:
        file_path = os.path.join(INPUT_DIR, file)
        try:
            doc = Document(file_path)
            if PAGE_BREAK:
                add_page_break(composer)
            composer.append(doc)
        except Exception as e:
            print(f"\n⚠️ Error merging {file}: {e}")
        pbar.update(1)

# --- Save merged DOCX ---
output_docx_path = os.path.join(OUTPUT_DIR, DOCX_FILENAME)
print("💾 Saving merged DOCX...")
composer.save(output_docx_path)
print("\n✅ DOCX merge completed!")
print(f"📁 Output DOCX: {output_docx_path}")

# --- PDF conversion ---
if PDF_CONVERT:
    choice = input("\n📄 Convert merged DOCX to PDF? (y/n): ").strip().lower()
    if choice in ("y", "yes"):
        # Auto-detect LibreOffice
        possible_paths = [
            r"C:\Program Files\LibreOffice\program\soffice.exe",
            r"C:\Program Files (x86)\LibreOffice\program\soffice.exe"
        ]
        LIBREOFFICE_PATH = None
        for path in possible_paths:
            if os.path.isfile(path):
                LIBREOFFICE_PATH = path
                break
        if LIBREOFFICE_PATH is None:
            LIBREOFFICE_PATH = shutil.which("soffice")
        if LIBREOFFICE_PATH is None:
            print("❌ LibreOffice not found. Please install it or add it to PATH.")
            sys.exit(1)

        output_pdf_path = os.path.join(OUTPUT_DIR, PDF_FILENAME)
        print(f"🔄 Converting to PDF using LibreOffice: {LIBREOFFICE_PATH}")
        try:
            subprocess.run(
                [LIBREOFFICE_PATH, "--headless", "--convert-to", "pdf",
                 "--outdir", OUTPUT_DIR, output_docx_path],
                check=True
            )
            print("✅ PDF conversion completed!")
            print(f"📁 Output PDF: {output_pdf_path}")
        except Exception as e:
            print("❌ PDF conversion failed.")
            print(e)
    else:
        print("⏭️ PDF conversion skipped.")

print("\n🏁 Process finished.")
