"""
router.py — Input-type detection and pipeline routing.

Detects the input file type by extension (and for PDFs, by content analysis)
and dispatches to the correct extraction pipeline.
"""

import os
import logging
from pathlib import Path
from enum import Enum, auto
from typing import Callable, Any

logger = logging.getLogger(__name__)


class InputType(Enum):
    """Supported input file types."""
    DOCX = auto()
    PDF_DIGITAL = auto()
    PDF_SCANNED = auto()
    IMAGE = auto()


# File extensions mapped to initial classification
_EXTENSION_MAP = {
    '.docx': InputType.DOCX,
    '.pdf': None,  # Requires content analysis to distinguish digital vs scanned
    '.png': InputType.IMAGE,
    '.jpg': InputType.IMAGE,
    '.jpeg': InputType.IMAGE,
    '.tiff': InputType.IMAGE,
    '.tif': InputType.IMAGE,
    '.bmp': InputType.IMAGE,
}

# Minimum number of meaningful text characters on the first page
# to classify a PDF as "digital" (vs scanned/image-based).
_PDF_DIGITAL_TEXT_THRESHOLD = 50


def detect_input_type(file_path: str) -> InputType:
    """
    Detect the input file type and return the appropriate InputType.

    For PDFs, peeks at the first page using PyMuPDF to determine whether
    it contains extractable digital text or is scanned (image-based).

    Args:
        file_path: Path to the input file.

    Returns:
        InputType enum value.

    Raises:
        FileNotFoundError: If the file does not exist.
        ValueError: If the file extension is not supported.
    """
    path = Path(file_path)

    ext = path.suffix.lower()
    if ext not in _EXTENSION_MAP:
        supported = ', '.join(sorted(_EXTENSION_MAP.keys()))
        raise ValueError(
            f"Unsupported file extension '{ext}'. "
            f"Supported extensions: {supported}"
        )

    if not path.exists():
        raise FileNotFoundError(f"Input file not found: {file_path}")

    # Non-PDF types are determined purely by extension
    initial_type = _EXTENSION_MAP[ext]
    if initial_type is not None:
        logger.info("Detected input type: %s (by extension '%s')", initial_type.name, ext)
        return initial_type

    # PDF: analyze first page to distinguish digital vs scanned
    return _classify_pdf(file_path)


def _classify_pdf(file_path: str) -> InputType:
    """
    Classify a PDF as digital or scanned by extracting text from the first page.

    If the first page yields fewer than _PDF_DIGITAL_TEXT_THRESHOLD characters
    of meaningful (non-whitespace) text, the PDF is classified as scanned.
    """
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            logger.warning(
                "PyMuPDF not installed. Cannot classify PDF — "
                "defaulting to PDF_DIGITAL. Install with: pip install PyMuPDF"
            )
            return InputType.PDF_DIGITAL

    try:
        doc = pymupdf.open(file_path)
        if len(doc) == 0:
            logger.warning("PDF has no pages: %s", file_path)
            doc.close()
            return InputType.PDF_SCANNED

        # Sample the first page (and optionally a middle page for multi-page docs)
        pages_to_check = [0]
        if len(doc) > 2:
            pages_to_check.append(len(doc) // 2)

        total_text_chars = 0
        for page_idx in pages_to_check:
            page = doc[page_idx]
            text = page.get_text("text")
            # Count non-whitespace characters as "meaningful"
            meaningful = len(text.replace(" ", "").replace("\n", "").replace("\r", ""))
            total_text_chars += meaningful

        doc.close()

        avg_chars = total_text_chars / len(pages_to_check)

        if avg_chars >= _PDF_DIGITAL_TEXT_THRESHOLD:
            logger.info(
                "PDF classified as DIGITAL (avg %.0f chars/page on sampled pages)",
                avg_chars,
            )
            return InputType.PDF_DIGITAL
        else:
            logger.info(
                "PDF classified as SCANNED (avg %.0f chars/page — below threshold %d)",
                avg_chars,
                _PDF_DIGITAL_TEXT_THRESHOLD,
            )
            return InputType.PDF_SCANNED

    except Exception as e:
        logger.error("Error analyzing PDF '%s': %s. Defaulting to PDF_DIGITAL.", file_path, e)
        return InputType.PDF_DIGITAL


def route(file_path: str, output_path: str, **kwargs) -> str:
    """
    Detect the input type and dispatch to the appropriate pipeline.

    Args:
        file_path: Path to the input file.
        output_path: Path for the output .docx file.
        **kwargs: Additional options passed to the pipeline (e.g., math_ocr, dpi).

    Returns:
        Path to the generated output .docx file.
    """
    input_type = detect_input_type(file_path)

    if input_type == InputType.DOCX:
        from .docx_pipeline import process_docx
        return process_docx(file_path, output_path, **kwargs)

    elif input_type == InputType.PDF_DIGITAL:
        try:
            from .pdf_pipeline import process_pdf
            return process_pdf(file_path, output_path, **kwargs)
        except ImportError:
            logger.warning("pdf_pipeline is not available. Falling back to scanned PDF (image-based) pipeline.")
            from .image_pipeline import process_image
            return process_image(file_path, output_path, is_scanned_pdf=True, **kwargs)

    elif input_type == InputType.PDF_SCANNED:
        from .image_pipeline import process_image
        # For scanned PDFs, the image pipeline handles page-by-page conversion
        return process_image(file_path, output_path, is_scanned_pdf=True, **kwargs)

    elif input_type == InputType.IMAGE:
        from .image_pipeline import process_image
        return process_image(file_path, output_path, **kwargs)

    else:
        raise ValueError(f"Unhandled input type: {input_type}")
