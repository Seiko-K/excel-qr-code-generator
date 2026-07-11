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

<p align="center">
  <img src="images/architecture.svg" width="800" alt="Excel QR Code Generator Architecture">
</p>

---

## Features

✓ Batch QR Generation

✓ Dynamic Batch Range Detection

✓ Skip Empty Rows

✓ PNG Export

✓ Worksheet Image Insertion

✓ Button-ready Macro Entry Points

✓ UTF-8 Encoding

✓ BOM Removal

✓ Zero-width Character Cleanup

✓ URL Encoding

✓ Automatic PNG File Naming

✓ Status Reporting (OK / NG / Skip)

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

```text
Excel Data
      │
      ▼
Detect Last Data Row
      │
      ▼
Text Cleanup
      │
      ▼
UTF-8 Encoding
      │
      ▼
BOM Removal
      │
      ▼
Generate QR Code
      │
      ▼
Export PNG
      │
      ▼
Insert into Worksheet
      │
      ▼
Status Output
 (OK / NG / Skip)
```

---

## Repository Structure

```text
excel-qr-code-generator/

├── images/
│   └── architecture.svg
│
├── src/
│   └── QRGenerator.bas
│
├── LICENSE
└── README.md
```

---

## Technical Highlights

### QR Generation

- Batch QR code generation
- Dynamic last-row detection
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
- Empty row detection

### Excel Automation

- Automatic worksheet image insertion
- Button-ready macro execution
- Status tracking
- Image resizing
- Cell fitting
- Reusable workflow

---

## Future Roadmap

### v0.2

- Processing progress indicator
- Elapsed time display
- Output folder selection
- Improved error handling

### v0.3

- Configuration worksheet
- Custom QR size
- Custom margin
- Configurable worksheet columns

### v0.4

- URL validation
- Duplicate detection
- Validation report
- Error summary

### v0.5

- Label printing support
- SVG export
- Logo QR generation
- ZIP export

---

## Current Version

**v0.1**

Released: **2026-07-11**

Status: **Stable Release**

---

## License

MIT License
