from pathlib import Path
import html
import webbrowser

import pikepdf


ROOT = Path.cwd()
REPORT = ROOT / "Report.html"


def is_linearized(pdf_file):
    try:
        with open(pdf_file, "rb") as handle:
            header = handle.read(4096)
        return b"/Linearized" in header
    except Exception:
        return False


def get_metadata(pdf_file):
    result = {
        "Pages": "",
        "Title": "",
        "Author": "",
        "Subject": "",
        "Keywords": "",
        "Fast Web View": "No",
    }

    with pikepdf.open(pdf_file) as pdf:
        info = pdf.docinfo
        result["Pages"] = len(pdf.pages)
        result["Title"] = str(info.get("/Title", "") or "")
        result["Author"] = str(info.get("/Author", "") or "")
        result["Subject"] = str(info.get("/Subject", "") or "")
        result["Keywords"] = str(info.get("/Keywords", "") or "")

        if is_linearized(pdf_file):
            result["Fast Web View"] = "Yes"

    return result


def main():
    pdfs = sorted(ROOT.glob("*.pdf"))
    rows = []
    pass_count = 0
    fail_count = 0

    for pdf in pdfs:
        m = get_metadata(pdf)

        issues = []
        if m["Title"] != "":
            issues.append("Title")
        if m["Author"] != "":
            issues.append("Author")
        if m["Subject"] != "":
            issues.append("Subject")
        if m["Keywords"] != "":
            issues.append("Keywords")
        if m["Fast Web View"] != "Yes":
            issues.append("Fast Web View")

        result = "PASS" if len(issues) == 0 else "FAIL"
        if result == "PASS":
            pass_count += 1
        else:
            fail_count += 1

        rows.append(
            f"""
<tr>
<td>{html.escape(pdf.name)}</td>
<td align="right">{m["Pages"]}</td>
<td>{m["Fast Web View"]}</td>
<td>{"Empty" if m["Title"] == "" else "Exists"}</td>
<td>{"Empty" if m["Author"] == "" else "Exists"}</td>
<td>{"Empty" if m["Subject"] == "" else "Exists"}</td>
<td>{"Empty" if m["Keywords"] == "" else "Exists"}</td>
<td>{result}</td>
<td>{", ".join(issues)}</td>
</tr>
"""
        )

    html_text = f"""
<!DOCTYPE html>

<html>

<head>

<meta charset="utf-8">

<title>PDF QA Validation Report</title>

<style>

body{{
font-family:Arial;
margin:40px;
}}

table{{
border-collapse:collapse;
width:100%;
}}

th,td{{
border:1px solid black;
padding:8px;
}}

th{{
background:#EAEAEA;
}}

h1{{
margin-bottom:5px;
}}

.summary{{
margin-top:20px;
margin-bottom:20px;
line-height:1.8;
}}

</style>

</head>

<body>

<h1>PDF QA Validation Report</h1>

<hr>

<div class="summary">

<b>Folder</b><br>

{ROOT}

<br><br>

<b>Total PDF :</b> {len(pdfs)}<br>
<b>PASS :</b> {pass_count}<br>
<b>FAIL :</b> {fail_count}

</div>

<table>

<tr>
<th>File</th>
<th>Pages</th>
<th>Fast Web View</th>
<th>Title</th>
<th>Author</th>
<th>Subject</th>
<th>Keywords</th>
<th>Result</th>
<th>Reason</th>
</tr>

{''.join(rows)}

</table>

</body>

</html>
"""

    REPORT.write_text(html_text, encoding="utf-8")
    print("Finished.")
    webbrowser.open(REPORT.as_uri())
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
