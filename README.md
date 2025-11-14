# Officify

**Officify** is a Node.js package that lets you **modify Office** documents in the OpenDocument Format (ODT, ODS, ODP).
With Officify, you can automatically replace text placeholders, replace images, and export your documents as PDFs.

## Features

- Replace placeholders in OpenDocument files (`replacePlaceholders`)
- Replace images in documents (`replaceImages`)
- Export to PDF using LibreOffice (`exportPDF`)

## Installation

**Note:** Officify requires LibreOffice to be installed on your system for PDF export. Make sure `soffice` is available in your PATH.

```bash
npm install officify
```

## Usage

```ts
import { Officify } from "officify";
import fs from "fs/promises";

// Load template
const buffer = await fs.readFile("template.odt");
const doc = new Officify(buffer);
await doc.load();

// Replace placeholders
await doc.replacePlaceholders({ "{{name}}": "John Doe" });

// Replace images
const img = await fs.readFile("logo.png");
doc.replaceImages({ "logo.png": img });

// Export to PDF
const pdf = await doc.exportPDF(); // all pages
const pdf = await doc.exportPDF("1-3"); // pages 1-3
await fs.writeFile("output.pdf", pdf);

// Save modified document
const newDoc = await doc.getBuffer();
await fs.writeFile("output.odt", newDoc);
```
