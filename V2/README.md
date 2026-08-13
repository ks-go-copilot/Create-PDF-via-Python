# V2

This folder contains the exe-oriented workflow.

## Files
- `make_csdrg_pdf.py` - source for `make_csdrg_pdf.exe`
- `pdf_checker.py` - source for `pdf_checker.exe`
- `user_note.md` - description of what the tools do to PDFs

## Usage
- Double-click the generated exe directly.
- No VBS launcher is required.
- `make_csdrg_pdf.exe` still requires Microsoft Word on Windows.

## Packaging recommendation
Use PyInstaller and build one exe per script:

```bash
pyinstaller --onefile --noconfirm --clean make_csdrg_pdf.py
pyinstaller --onefile --noconfirm --clean pdf_checker.py
```

Notes:
- `--onefile` gives the simplest end-user experience.
- The exe size will still be driven mostly by the embedded Python runtime plus `pywin32` and `pikepdf`.
- Keeping the code small helps, but this kind of tool will not be tiny.
