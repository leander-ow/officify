# Officify

**Officify** is a Node.js package that lets you modify Office documents in the OpenDocument Format (ODT, ODS, ODP).
With Officify, you can automatically replace text placeholders, replace images, and export your documents as PDFs.

## Features

- Replace placeholders in OpenDocument files (`replacePlaceholders`)
- Replace images in documents (`replaceImages`)
- Export to PDF using LibreOffice (`exportPDF`)

## Installation

```bash
npm install officify
```

---

**Note:** Officify requires LibreOffice to be installed on your system for PDF export. Make sure `soffice` is available in your PATH.
