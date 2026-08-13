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
import shutil
import sys
import time
from pathlib import Path

try:
    import pikepdf
    import win32com.client
    from pikepdf import Dictionary, Name
except ImportError as exc:
    print(f"Missing library: {exc}")
    print("Run: python -m pip install pywin32 pikepdf")
    sys.exit(1)


def get_app_dir():
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


APP_DIR = get_app_dir()
DOCX_SRC = APP_DIR / "csdrg.docx"
FINAL_PDF = APP_DIR / "csdrg.pdf"
TEMP_DIR = Path(os.environ.get("TEMP", "C:\\Temp")) / "csdrg_build"
TEMP_DOCX = TEMP_DIR / "csdrg.docx"
TEMP_PDF = TEMP_DIR / "csdrg.pdf"


def export_word_docx_to_pdf(source_docx, output_pdf):
    word = win32com.client.Dispatch("Word.Application")
    word.Visible = False
    try:
        doc = word.Documents.Open(source_docx)
        doc.ExportAsFixedFormat(
            OutputFileName=output_pdf,
            ExportFormat=17,
            OpenAfterExport=False,
            OptimizeFor=0,
            Range=0,
            From=1,
            To=1,
            Item=0,
            IncludeDocProps=True,
            KeepIRM=True,
            CreateBookmarks=1,
            DocStructureTags=True,
            BitmapMissingFonts=True,
            UseISO19005_1=False,
        )
        doc.Close(False)
    finally:
        word.Quit()


def post_process_pdf(pdf_path):
    with pikepdf.open(pdf_path, allow_overwriting_input=True) as pdf:
        pdf.Root["/PageMode"] = Name("/UseOutlines")

        if "/ViewerPreferences" not in pdf.Root:
            pdf.Root["/ViewerPreferences"] = Dictionary()

        info = pdf.docinfo
        for key in ["/Title", "/Author", "/Subject", "/Keywords", "/Creator", "/Producer"]:
            if key in info:
                info[key] = ""

        if "/Outlines" in pdf.Root:
            print("     Bookmarks: found OK")
        else:
            print("     WARNING: No bookmarks found.")
            print("     Ensure csdrg.docx uses built-in Word Heading 1/2/3 styles.")

        pdf.save(pdf_path)


def verify_pdf_version(pdf_path):
    with open(pdf_path, "rb") as handle:
        header = handle.read(8).decode("ascii", errors="replace").strip()

    print(f"     PDF header: {header}")
    version_ok = any(version in header for version in ["1.4", "1.5", "1.6", "1.7"])
    if version_ok:
        print("     Version <= 1.7: OK")
    else:
        print("     WARNING: Version may exceed 1.7. Check before submission.")
    return version_ok


def copy_with_retry(source_path, destination_path, retries=5):
    for attempt in range(retries):
        try:
            shutil.copy2(source_path, destination_path)
            return
        except PermissionError:
            if attempt < retries - 1:
                print(f"     File locked, retrying in 3s... (attempt {attempt + 1}/{retries})")
                time.sleep(3)
            else:
                raise


def main():
    print("=" * 60)
    print("CSDRG PDF Builder")
    print("=" * 60)

    if not DOCX_SRC.exists():
        raise FileNotFoundError(f"csdrg.docx not found at:\n  {DOCX_SRC}")

    TEMP_DIR.mkdir(parents=True, exist_ok=True)
    shutil.copy2(str(DOCX_SRC), str(TEMP_DOCX))
    print(f"[1/4] Copied docx to temp: {TEMP_DOCX}")

    print("[2/4] Converting Word -> PDF via Microsoft Word...")
    export_word_docx_to_pdf(str(TEMP_DOCX), str(TEMP_PDF))
    print(f"     PDF created: {TEMP_PDF}")

    time.sleep(2)

    print("[3/4] Setting PDF properties (bookmarks panel + clear metadata)...")
    post_process_pdf(str(TEMP_PDF))

    print("[4/4] Verifying PDF version...")
    version_ok = verify_pdf_version(str(TEMP_PDF))

    copy_with_retry(str(TEMP_PDF), str(FINAL_PDF))

    print()
    print("=" * 60)
    print(f"Done!  Output: {FINAL_PDF}")
    print()
    print("Checklist:")
    print(f"  [{'OK' if version_ok else '!!'}] PDF version <= 1.7")
    print("  [OK] Bookmarks panel shown on open")
    print("  [OK] Title/Author/Subject/Keywords cleared")
    print("  [OK] File named csdrg.pdf")
    print("=" * 60)
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
