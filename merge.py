from docx import Document
from docxcompose.composer import Composer
from tqdm import tqdm
import os
import re
import sys

INPUT_DIR = "input"
OUTPUT_DIR = "output"
OUTPUT_FILE = "merged_output.docx"

def numeric_sort(filename):
    numbers = re.findall(r'\d+', filename)
    return int(numbers[0]) if numbers else sys.maxsize

# --- Checks ---
if not os.path.exists(INPUT_DIR):
    print(f"❌ Input folder '{INPUT_DIR}' not found.")
    sys.exit(1)

os.makedirs(OUTPUT_DIR, exist_ok=True)

# --- Load and sort files ---
files = sorted(
    [f for f in os.listdir(INPUT_DIR) if f.lower().endswith(".docx")],
    key=numeric_sort
)

if len(files) < 2:
    print("❌ Not enough DOCX files to merge.")
    sys.exit(1)

print(f"📄 Found {len(files)} DOCX files in '{INPUT_DIR}'")
print("🔧 Initializing master document...")

# --- Master document ---
master_path = os.path.join(INPUT_DIR, files[0])
master = Document(master_path)
composer = Composer(master)

# --- Progress bar ---
with tqdm(
    total=len(files) - 1,
    desc="🔗 Merging documents",
    unit="file",
    bar_format="{l_bar}{bar}| {n_fmt}/{total_fmt} [{elapsed}<{remaining}]"
) as pbar:

    for file in files[1:]:
        file_path = os.path.join(INPUT_DIR, file)
        try:
            doc = Document(file_path)
            composer.append(doc)
        except Exception as e:
            print(f"\n⚠️ Error merging {file}: {e}")
        pbar.update(1)

# --- Save output ---
output_path = os.path.join(OUTPUT_DIR, OUTPUT_FILE)
print("💾 Saving final document...")
composer.save(output_path)

print("\n✅ Merge completed successfully!")
print(f"📁 Output file: {output_path}")
