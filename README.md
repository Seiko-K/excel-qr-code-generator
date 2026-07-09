![Version](https://img.shields.io/badge/version-v0.1-blue)
![Status](https://img.shields.io/badge/status-stable-success)
![License](https://img.shields.io/badge/license-MIT-green)
![Excel](https://img.shields.io/badge/Excel-VBA-darkgreen)
![Feature](https://img.shields.io/badge/feature-Batch_QR_Generation-orange)

# Excel QR Code Generator

Reusable Excel VBA toolkit for batch QR code generation.

Built for situations where many QR codes need to be created efficiently from Excel data.

Generate hundreds of QR codes at once, export them as PNG files, and optionally insert them directly into Excel worksheets.

---

## Architecture

![Architecture](images/architecture.svg)

---

## Features

✓ Batch QR Generation

✓ PNG Export

✓ Worksheet Image Insertion

✓ UTF-8 Encoding

✓ BOM Removal

✓ Zero-width Character Cleanup

✓ URL Encoding

✓ Automatic PNG File Naming

✓ Status Reporting (OK / NG)

✓ Re-runnable Workflow

✓ Excel VBA

---

## Use Cases

### Product Labels

- Product QR codes
- Inventory labels
- Asset management

### Event Management

- Registration QR codes
- Check-in systems
- Visitor badges

### Surveys

- Questionnaire links
- Feedback forms
- Customer satisfaction surveys

### Internal Documents

- Digital manuals
- Shared resources
- Company documents

### Batch Processing

Generate hundreds of QR codes automatically from Excel without creating them one by one.

---

## Workflow

Excel Data

↓

Text Cleanup

↓

UTF-8 Encoding

↓

BOM Removal

↓

QR Code Generation

↓

PNG Export

↓

Worksheet Image Insertion

↓

Status Output (OK / NG)

---

## Repository Structure

```text
excel-qr-code-generator/

├── images/
│   └── architecture.svg
│
├── src/
│   ├── QRGenerator.bas
│   └── PictureUtility.bas
│
├── LICENSE
└── README.md
```

---

## Technical Highlights

### QR Generation

- Batch QR code generation
- Quiet zone (margin) control
- PNG image export
- Automatic sequential file naming
- Zero-padded numbering

### Data Processing

- UTF-8 URL encoding
- BOM removal
- Zero-width character cleanup
- Line break removal
- Whitespace normalization

### Excel Automation

- Automatic worksheet image insertion
- Status tracking
- Image resizing
- Cell fitting
- Reusable workflow

---

## Current Version

v0.1

Released: 2026-07-09

Status: Stable Release

---

## License

MIT License
