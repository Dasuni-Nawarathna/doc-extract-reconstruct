"""
pdf_pipeline.py — Digital PDF to DOCX reconstruction pipeline.

Extracts text, fonts (family, size, bold/italic, color) per span using PyMuPDF,
and reconstructs the document with matching paragraph blocks and formatting runs.
"""

import logging
from typing import Optional, Tuple
from pathlib import Path

try:
    import pymupdf
except ImportError:
    try:
        import fitz as pymupdf
    except ImportError:
        raise ImportError(
            "PyMuPDF is required for PDF parsing. "
            "Install with: pip install PyMuPDF"
        )

from docx import Document
from .writer import (
    create_document, create_paragraph, create_run, save_document,
    set_paragraph_properties, qn
)
from .confidence import ConfidenceReport, Level

logger = logging.getLogger(__name__)

def process_pdf(
    input_path: str,
    output_path: str,
    **kwargs,
) -> str:
    """
    Process a digital PDF file, extracting text and font formatting per span,
    and reconstruct it as a .docx file.

    Args:
        input_path: Path to the input digital PDF.
        output_path: Path for the reconstructed output .docx.

    Returns:
        Path to the output .docx file.
    """
    logger.info("PDF Pipeline: Processing '%s'", input_path)

    report = ConfidenceReport(source_file=input_path, output_file=output_path)
    dst_doc = create_document()

    # Remove default empty paragraph
    body = dst_doc.element.body
    for p in body.findall(qn('w:p')):
        body.remove(p)

    doc = pymupdf.open(input_path)

    for page_idx, page in enumerate(doc):
        logger.info("Processing PDF page %d/%d", page_idx + 1, len(doc))
        text_dict = page.get_text("dict")

        # Sort blocks by vertical position, then horizontal
        blocks = text_dict.get("blocks", [])
        blocks = sorted(blocks, key=lambda b: (b.get("bbox", (0, 0, 0, 0))[1], b.get("bbox", (0, 0, 0, 0))[0]))

        for block in blocks:
            if block.get("type") != 0:  # Skip image blocks
                continue

            dst_p = create_paragraph(dst_doc)

            # Sort lines in block top-to-bottom
            lines = sorted(block.get("lines", []), key=lambda l: l.get("bbox", (0, 0, 0, 0))[1])

            for line_idx, line in enumerate(lines):
                # Sort spans left-to-right
                spans = sorted(line.get("spans", []), key=lambda s: s.get("bbox", (0, 0, 0, 0))[0])

                for span_idx, span in enumerate(spans):
                    text = span.get("text", "")
                    if not text.strip():
                        continue

                    font_raw = span.get("font", "")
                    font_size = span.get("size", 11.0)
                    font_color_int = span.get("color", 0)
                    flags = span.get("flags", 0)

                    # Extract RGB values from PyMuPDF color integer (0xRRGGBB)
                    r = (font_color_int >> 16) & 255
                    g = (font_color_int >> 8) & 255
                    b = font_color_int & 255
                    font_color = (r, g, b)

                    # Normalize font name and formats
                    font_lower = font_raw.lower()

                    # Deduce bold / italic
                    is_bold = bool(flags & 2**4) or "bold" in font_lower or "black" in font_lower or "heavy" in font_lower
                    is_italic = bool(flags & 2**1) or "italic" in font_lower or "oblique" in font_lower

                    # Map raw fonts to standard system fonts
                    font_name = "Calibri"
                    if any(term in font_lower for term in ["times", "roman", "serif", "georgia", "cambria"]):
                        font_name = "Times New Roman"
                    elif any(term in font_lower for term in ["arial", "helvetica", "sans", "calibri"]):
                        font_name = "Arial"
                    elif any(term in font_lower for term in ["courier", "mono", "consolas"]):
                        font_name = "Courier New"
                    else:
                        clean_font = font_raw
                        if "+" in clean_font:
                            clean_font = clean_font.split("+", 1)[1]
                        for suffix in ["-Bold", "-Italic", "Bold", "Italic", "-Regular", "Regular", "-MT", "MT"]:
                            clean_font = clean_font.replace(suffix, "")
                        font_name = clean_font.strip()

                    # Prepend space between spans to avoid run merging without spaces
                    run_text = text
                    if span_idx > 0 or line_idx > 0:
                        if not run_text.startswith(" ") and not run_text.startswith("\n"):
                            run_text = " " + run_text

                    create_run(
                        dst_p,
                        text=run_text,
                        font_name=font_name,
                        font_size=font_size,
                        bold=is_bold or None,
                        italic=is_italic or None,
                        color=font_color,
                    )

                    report.add_simple(
                        location=f"Page {page_idx + 1}, Block {block.get('number', 0)}, Span '{text[:10]}'",
                        property_name="font_style",
                        confidence=Level.HIGH,
                        detail=f"Preserved digital font '{font_raw}' (size={font_size:.1f}pt, bold={is_bold}, italic={is_italic}, color={font_color})",
                        original_value=font_raw,
                        reconstructed_value=font_name
                    )

    doc.close()
    save_document(dst_doc, output_path)
    report.write()
    return output_path
