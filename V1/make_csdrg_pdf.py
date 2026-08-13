"""
make_csdrg_pdf.py
=================
Converts csdrg.docx -> csdrg.pdf following CSDRG Completion Guidelines v1.4:

  - Uses Microsoft Word COM (ExportAsFixedFormat) to preserve heading bookmarks
  - Sets initial view to show Bookmarks Panel on open
  - Clears Title / Author / Subject / Keywords per CSDRG guidelines (p.21)
  - Verifies PDF version <= 1.7 (FDA eCTD requirement)

Requirements:  pip install pywin32 pikepdf
"""

import os
import sys
import shutil
import time

try:
    import win32com.client
    import pikepdf
    from pikepdf import Dictionary, Name
except ImportError as e:
    print(f"Missing library: {e}")
    print("Run:  python -m pip install pywin32 pikepdf")
    sys.exit(1)

# ── Paths ──────────────────────────────────────────────────────────────────────
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
DOCX_SRC   = os.path.join(SCRIPT_DIR, "csdrg.docx")
FINAL_PDF  = os.path.join(SCRIPT_DIR, "csdrg.pdf")
TEMP_DIR   = os.path.join(os.environ.get("TEMP", "C:\\Temp"), "csdrg_build")
TEMP_DOCX  = os.path.join(TEMP_DIR, "csdrg.docx")
TEMP_PDF   = os.path.join(TEMP_DIR, "csdrg.pdf")

# ── Step 1: Copy docx to temp (avoids OneDrive/space path issues) ──────────────
print("=" * 60)
print("CSDRG PDF Builder")
print("=" * 60)

if not os.path.exists(DOCX_SRC):
    print(f"ERROR: csdrg.docx not found at:\n  {DOCX_SRC}")
    sys.exit(1)

os.makedirs(TEMP_DIR, exist_ok=True)
shutil.copy2(DOCX_SRC, TEMP_DOCX)
print(f"[1/4] Copied docx to temp: {TEMP_DOCX}")

# ── Step 2: Word COM -> PDF with heading bookmarks ─────────────────────────────
print("[2/4] Converting Word -> PDF via Microsoft Word...")
word = win32com.client.Dispatch("Word.Application")
word.Visible = False
try:
    doc = word.Documents.Open(TEMP_DOCX)
    doc.ExportAsFixedFormat(
        OutputFileName    = TEMP_PDF,
        ExportFormat      = 17,     # wdExportFormatPDF
        OpenAfterExport   = False,
        OptimizeFor       = 0,      # wdExportOptimizeForPrint
        Range             = 0,      # wdExportAllDocument
        From              = 1,
        To                = 1,
        Item              = 0,      # wdExportDocumentContent
        IncludeDocProps   = True,
        KeepIRM           = True,
        CreateBookmarks   = 1,      # wdExportCreateHeadingBookmarks
        DocStructureTags  = True,
        BitmapMissingFonts = True,
        UseISO19005_1     = False,
    )
    doc.Close(False)
    print(f"     PDF created: {TEMP_PDF}")
finally:
    word.Quit()

time.sleep(2)

# ── Step 3: pikepdf post-processing ───────────────────────────────────────────
print("[3/4] Setting PDF properties (bookmarks panel + clear metadata)...")
with pikepdf.open(TEMP_PDF, allow_overwriting_input=True) as pdf:

    # Show Bookmarks Panel when opened (CSDRG guideline requirement)
    pdf.Root["/PageMode"] = Name("/UseOutlines")

    # ViewerPreferences: default page layout and magnification
    if "/ViewerPreferences" not in pdf.Root:
        pdf.Root["/ViewerPreferences"] = Dictionary()

    # Clear Title, Author, Subject, Keywords (CSDRG Completion Guidelines p.21)
    info = pdf.docinfo
    for key in ["/Title", "/Author", "/Subject", "/Keywords", "/Creator", "/Producer"]:
        if key in info:
            info[key] = ""

    # Report bookmark count
    if "/Outlines" in pdf.Root:
        print("     Bookmarks: found OK")
    else:
        print("     WARNING: No bookmarks found.")
        print("     Ensure csdrg.docx uses built-in Word Heading 1/2/3 styles.")

    pdf.save(TEMP_PDF)

# ── Step 4: Verify PDF version and copy to destination ────────────────────────
print("[4/4] Verifying PDF version...")
with open(TEMP_PDF, "rb") as f:
    header = f.read(8).decode("ascii", errors="replace").strip()
print(f"     PDF header: {header}")

version_ok = any(v in header for v in ["1.4", "1.5", "1.6", "1.7"])
if version_ok:
    print("     Version <= 1.7: OK")
else:
    print(f"     WARNING: Version may exceed 1.7. Check before submission.")

# Copy to final destination (retry if locked)
for attempt in range(5):
    try:
        shutil.copy2(TEMP_PDF, FINAL_PDF)
        break
    except PermissionError:
        if attempt < 4:
            print(f"     File locked, retrying in 3s... (attempt {attempt+1}/5)")
            time.sleep(3)
        else:
            print(f"\nERROR: Could not write to {FINAL_PDF}")
            print("Please close any program that has csdrg.pdf open, then re-run.")
            sys.exit(1)

print()
print("=" * 60)
print(f"Done!  Output: {FINAL_PDF}")
print()
print("Checklist:")
print(f"  [{'OK' if version_ok else '!!'}] PDF version <= 1.7")
print( "  [OK] Bookmarks panel shown on open")
print( "  [OK] Title/Author/Subject/Keywords cleared")
print( "  [OK] File named csdrg.pdf")
print("=" * 60)
