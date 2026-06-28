# doc-extract-reconstruct

A Python CLI tool that extracts content from `.docx`, `.pdf`, or image files (`.png`/`.jpg`) and reconstructs it as a new `.docx` file, preserving text formatting and math equations (OMML) as closely as possible.

## Features

- **DOCX → DOCX** (near-lossless): Preserves all text formatting (font, size, bold, italic, underline, color, highlight) and copies math equations (OMML) verbatim via XML deep-copy. Cambria Math equations are preserved exactly.
- **PDF → DOCX**: Extracts text with font metadata (name, size, bold/italic flags, color) per span using PyMuPDF. Detects equation regions and converts them via math-OCR → LaTeX → OMML.
- **Image → DOCX**: Runs Tesseract OCR for text extraction with bounding boxes. Uses heuristic styling inference for bold/italic/size. Detects and converts equation regions.
- **Confidence flagging**: Tracks confidence levels (HIGH/MEDIUM/LOW) for every reconstructed element. Low-confidence items are highlighted in yellow in the output document and listed in a companion report file.

## Installation

### Prerequisites

- **Python 3.9+**
- **Pandoc** (for LaTeX → OMML conversion in PDF/image pipelines)
  - Windows: Download from https://pandoc.org/installing.html
  - Or: `choco install pandoc` / `winget install pandoc`
- **Tesseract OCR** (for image/scanned-PDF pipelines)
  - Windows: Download from https://github.com/UB-Mannheim/tesseract/wiki
  - Or: `choco install tesseract`

### Python Dependencies

```bash
pip install -r requirements.txt
```

### Optional: Math OCR Engine

For equation recognition in PDF/image pipelines, install one of:

```bash
# Option A: pix2tex (heavier, ~2GB download, requires PyTorch)
pip install "pix2tex[gui]"

# Option B: rapid-latex-ocr (lighter, ONNX-based)
pip install rapid-latex-ocr
```

If neither is installed, equation regions will be flagged as placeholders.

## Usage

```bash
# Basic usage (auto-detects input type)
python extract.py input.docx -o output.docx

# PDF with verbose logging
python extract.py report.pdf -o report_reconstructed.docx --verbose

# Image with specific math OCR engine
python extract.py scan.png -o scan_output.docx --math-ocr pix2tex

# High-DPI rendering for better quality
python extract.py paper.pdf -o paper.docx --dpi 400
```

### CLI Options

| Option | Default | Description |
|:---|:---|:---|
| `-o, --output` | `<input>_reconstructed.docx` | Output .docx file path |
| `--math-ocr` | `auto` | Math OCR engine: `auto`, `pix2tex`, `rapid`, `none` |
| `--dpi` | `300` | DPI for image/PDF rendering |
| `--confidence-report` | `<output>_confidence_report.txt` | Path for confidence report |
| `-v, --verbose` | off | Enable debug logging |
| `-q, --quiet` | off | Suppress all output except errors |

## How It Works

### DOCX Pipeline (Highest Fidelity)

Operates at the raw XML level using `python-docx` + `lxml`:
1. Walks every `<w:p>` paragraph element in the source document
2. For text runs (`<w:r>`): deep-copies the `<w:rPr>` formatting properties and `<w:t>` text
3. For math equations (`<m:oMath>`, `<m:oMathPara>`): deep-copies the entire XML node verbatim — this preserves Cambria Math equations exactly as they were
4. Copies paragraph properties (alignment, spacing, numbering, indentation)
5. Deep-copies tables, hyperlinks, bookmarks, and other structural elements

### PDF Pipeline

1. Extracts text with font metadata via `PyMuPDF`'s `page.get_text("dict")` (returns font name, size, style flags, color per span)
2. Detects equation regions using font-name heuristics (Cambria Math, Symbol, etc.) and Unicode symbol density
3. Crops equation regions as high-DPI images
4. Converts equations: image → LaTeX (via pix2tex) → OMML (via Pandoc)
5. Falls back to image pipeline if PDF appears scanned (<50 chars of text per page)

### Image Pipeline

1. Preprocesses with OpenCV (grayscale, deskew, adaptive thresholding, noise removal)
2. Runs Tesseract OCR for word-level bounding boxes + confidence
3. Estimates formatting heuristically (see limitations below)
4. Detects and converts equation regions via the same math-OCR pipeline

## Known Limitations

### Fidelity by Input Type

| Property | DOCX Input | PDF (Digital) | PDF (Scanned) | Image |
|:---|:---|:---|:---|:---|
| Font Family | ✅ Exact copy | ✅ From PDF metadata | ❌ Default (Calibri) | ❌ Default (Calibri) |
| Font Size | ✅ Exact copy | ✅ From PDF metadata | ⚠️ Estimated from bbox | ⚠️ Estimated from bbox |
| Bold | ✅ Exact copy | ⚠️ From style flags | ⚠️ Stroke width heuristic | ⚠️ Stroke width heuristic |
| Italic | ✅ Exact copy | ⚠️ From style flags | ⚠️ Not reliably detectable | ⚠️ Not reliably detectable |
| Underline | ✅ Exact copy | ❌ Not in PDF metadata | ❌ Not detectable | ❌ Not detectable |
| Color | ✅ Exact copy | ✅ sRGB from PDF | ❌ Defaults to black | ❌ Defaults to black |
| Equations | ✅ Verbatim OMML copy | ⚠️ OCR + LaTeX → OMML | ⚠️ OCR + LaTeX → OMML | ⚠️ OCR + LaTeX → OMML |
| Tables | ✅ Exact copy | ⚠️ Text only (no structure) | ⚠️ Text only | ⚠️ Text only |
| Images | ✅ Deep-copied | ❌ Not extracted | ❌ Not extracted | N/A |

### Image-Based Font Detection

**True font-family identification from pixels is unreliable.** The image pipeline defaults all text to "Calibri" and flags every font-family property as LOW confidence. This is a fundamental limitation of OCR — character appearance alone cannot reliably distinguish between font families.

### Reconstructed Equations

Equations processed through the math-OCR → LaTeX → OMML pipeline are **approximations**. Complex constructs (multi-line equations, matrices, custom operators) may not convert perfectly. These are always flagged with MEDIUM or LOW confidence.

## Project Structure

```
doc-extract-reconstruct/
├── extract.py                  # CLI entry point
├── requirements.txt            # Python dependencies
├── README.md                   # This file
│
├── doc_extract_reconstruct/    # Main package
│   ├── __init__.py
│   ├── cli.py                  # argparse CLI interface
│   ├── router.py               # Input-type detection & routing
│   ├── confidence.py           # Confidence tracking & reporting
│   ├── omml_utils.py           # LaTeX → OMML conversion via Pandoc
│   ├── writer.py               # Common DOCX output writer
│   ├── docx_pipeline.py        # DOCX → DOCX (highest fidelity)
│   ├── pdf_pipeline.py         # PDF → DOCX
│   └── image_pipeline.py       # Image → DOCX (OCR-based)
│
└── tests/
    ├── test_docx_pipeline.py   # DOCX pipeline tests
    ├── create_test_fixture.py  # Test fixture generator
    └── fixtures/               # Test input files
```

## License

MIT
