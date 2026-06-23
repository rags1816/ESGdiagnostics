import os
import hashlib
import zipfile
from docx import Document

# Base directory of the toolkit
BASE_DIR = r'C:\\ESG'

# Paths to handover document and output zip
HANDOVER_PATH = os.path.join(BASE_DIR, 'ESG_Toolkit_V4_Handover.docx')
ZIP_PATH = os.path.join(BASE_DIR, 'ESG_Diagnostic_Toolkit_v4.1_Freeze.zip')

# ---------------------------------------------------------------------------
# Update handover guide – add a short note about the visual cue fix
# ---------------------------------------------------------------------------
try:
    doc = Document(HANDOVER_PATH)
    doc.add_paragraph('Update (2026-06-23): Visual cue on the Start Assessment screen now persists until the user selects a mode (Grouped by Theme or One at a Time).')
    doc.save(HANDOVER_PATH)
    print('Handover guide updated.')
except Exception as e:
    print('Failed to update handover guide:', e)

# ---------------------------------------------------------------------------
# Create (or recreate) the freeze zip – include every file under BASE_DIR except .zip files
# ---------------------------------------------------------------------------
if os.path.exists(ZIP_PATH):
    os.remove(ZIP_PATH)

with zipfile.ZipFile(ZIP_PATH, 'w', zipfile.ZIP_DEFLATED) as zipf:
    for foldername, _, filenames in os.walk(BASE_DIR):
        for filename in filenames:
            if filename.lower().endswith('.zip'):
                continue  # skip zip files themselves
            file_path = os.path.join(foldername, filename)
            arcname = os.path.relpath(file_path, BASE_DIR)
            zipf.write(file_path, arcname)
print('Freeze zip created.')

# ---------------------------------------------------------------------------
# Compute SHA-256 hash of the created zip and display it (ASCII hyphen only)
# ---------------------------------------------------------------------------
sha256 = hashlib.sha256()
with open(ZIP_PATH, 'rb') as f:
    for chunk in iter(lambda: f.read(8192), b''):
        sha256.update(chunk)
print('New SHA-256:', sha256.hexdigest())
