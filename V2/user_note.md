# User Note

This workflow is designed to make PDF processing transparent.

## `make_csdrg_pdf.exe`
What it does:
- opens `csdrg.docx`
- uses Microsoft Word COM to export the document to PDF
- sets the PDF to open with the bookmarks panel visible
- clears PDF document metadata fields such as Title, Author, Subject, Keywords, Creator, and Producer
- checks the PDF version header before copying the result to `csdrg.pdf`

What it does not do:
- it does not rewrite the document content
- it does not edit body text, tables, figures, or headings
- it does not change the original `csdrg.docx`

## `pdf_checker.exe`
What it does:
- reads PDF files in the current folder
- checks page count
- checks whether Fast Web View is enabled
- checks whether Title, Author, Subject, and Keywords are empty
- writes `Report.html`

What it does not do:
- it does not modify PDF content
- it does not change PDF metadata
- it only reports what it finds

## Summary
The processing is limited to PDF export settings and PDF properties. The body content is not intentionally modified by the post-processing step.
