"""
image_pipeline.py — Image-to-DOCX reconstruction pipeline (OCR-based).

Handles standalone images (.png/.jpg) and scanned PDFs (converted page-by-page).

Uses Tesseract OCR for text + bounding boxes, with heuristic styling inference:

KNOWN LIMITATIONS (documented here and in user-facing output):
  - Font family: CANNOT be reliably inferred from pixels. Defaults to "Calibri".
    All font-family values are flagged as LOW confidence.
  - Bold detection: Approximated via stroke width estimation. MEDIUM confidence.
  - Italic detection: Approximated via bounding box aspect ratio heuristics.
    MEDIUM confidence.
  - Font size: Estimated from bounding box height (pixels → points conversion).
    MEDIUM confidence.
  - Color: Not extracted from OCR (would require per-character pixel analysis).
    Defaults to black.
  - Equation recognition: Depends on math-OCR availability (pix2tex/rapid-latex-ocr).
"""

import logging
import math
import os
import tempfile
import shutil
from dataclasses import dataclass, field
from typing import Optional

import numpy as np
from PIL import Image

from .writer import (
    create_document, create_paragraph, create_run, insert_omml,
    save_document, highlight_low_confidence, insert_comment_text, qn,
)
from .confidence import ConfidenceReport, Level
from .omml_utils import image_to_omml, create_placeholder_run

logger = logging.getLogger(__name__)

# ── Constants ──────────────────────────────────────────────────────────

# Default font when font family cannot be inferred from image
DEFAULT_FONT = "Calibri"

# Assumed DPI for pixel-to-point conversion (standard screen DPI)
DEFAULT_DPI = 96

# Minimum word confidence from Tesseract to include in output
MIN_WORD_CONFIDENCE = 20

# Vertical gap threshold (as ratio of line height) to start a new paragraph
PARAGRAPH_GAP_RATIO = 1.8

# Minimum non-ASCII density in a region to flag as potential equation
MATH_DENSITY_THRESHOLD = 0.25


@dataclass
class OcrWord:
    """A word extracted from OCR with positioning and confidence."""
    text: str
    x: int      # left
    y: int      # top
    w: int      # width
    h: int      # height
    confidence: float  # 0-100
    block_num: int = 0
    par_num: int = 0
    line_num: int = 0
    word_num: int = 0


@dataclass
class OcrLine:
    """A line of words grouped from OCR output."""
    words: list[OcrWord] = field(default_factory=list)
    y: int = 0
    h: int = 0
    is_math: bool = False

    @property
    def text(self) -> str:
        return " ".join(w.text for w in self.words)

    @property
    def bbox(self) -> tuple:
        if not self.words:
            return (0, 0, 0, 0)
        x0 = min(w.x for w in self.words)
        y0 = min(w.y for w in self.words)
        x1 = max(w.x + w.w for w in self.words)
        y1 = max(w.y + w.h for w in self.words)
        return (x0, y0, x1, y1)


@dataclass
class OcrParagraph:
    """A paragraph of lines grouped by vertical proximity."""
    lines: list[OcrLine] = field(default_factory=list)

    @property
    def is_all_math(self) -> bool:
        return all(l.is_math for l in self.lines) if self.lines else False


def process_image(
    input_path: str,
    output_path: str,
    is_scanned_pdf: bool = False,
    math_ocr: str = "auto",
    dpi: int = 300,
    **kwargs,
) -> str:
    """
    Extract content from an image (or scanned PDF) and reconstruct as .docx.

    Args:
        input_path: Path to the image file or scanned PDF.
        output_path: Path for the output .docx file.
        is_scanned_pdf: If True, input is a PDF to be converted page-by-page.
        math_ocr: Math OCR engine ("auto", "pix2tex", "rapid", "none").
        dpi: DPI for PDF page rendering.

    Returns:
        Path to the generated output file.
    """
    logger.info(
        "%s Pipeline: Processing '%s'",
        "Scanned-PDF" if is_scanned_pdf else "Image",
        input_path,
    )

    report = ConfidenceReport(source_file=input_path, output_file=output_path)
    dst_doc = create_document()

    # Remove default empty paragraph
    body = dst_doc.element.body
    for p in body.findall(qn('w:p')):
        body.remove(p)

    tmpdir = tempfile.mkdtemp(prefix="doc_extract_img_")

    try:
        if is_scanned_pdf:
            image_paths = _pdf_to_images(input_path, tmpdir, dpi)
        else:
            image_paths = [input_path]

        for page_idx, img_path in enumerate(image_paths):
            logger.info(
                "Processing %s %d/%d",
                "page" if is_scanned_pdf else "image",
                page_idx + 1,
                len(image_paths),
            )
            _process_single_image(
                img_path, dst_doc, page_idx, report, tmpdir, math_ocr,
            )

    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

    save_document(dst_doc, output_path)
    report.write()

    logger.info(
        "%s Pipeline complete. Output: '%s'",
        "Scanned-PDF" if is_scanned_pdf else "Image",
        output_path,
    )
    return output_path


def _pdf_to_images(pdf_path: str, tmpdir: str, dpi: int) -> list[str]:
    """Convert each page of a scanned PDF to an image."""
    try:
        import pymupdf
    except ImportError:
        try:
            import fitz as pymupdf
        except ImportError:
            raise ImportError(
                "PyMuPDF is required for scanned PDF processing. "
                "Install with: pip install PyMuPDF"
            )

    image_paths = []
    doc = pymupdf.open(pdf_path)

    for page_num in range(len(doc)):
        page = doc[page_num]
        mat = pymupdf.Matrix(dpi / 72, dpi / 72)
        pix = page.get_pixmap(matrix=mat)
        img_path = os.path.join(tmpdir, f"page_{page_num:04d}.png")
        pix.save(img_path)
        image_paths.append(img_path)

    doc.close()
    return image_paths


def _process_single_image(
    img_path: str,
    dst_doc,
    page_idx: int,
    report: ConfidenceReport,
    tmpdir: str,
    math_ocr: str,
) -> None:
    """Process a single image through OCR and write to the document."""
    # ── Preprocessing ──────────────────────────────────────────────
    preprocessed_path = _preprocess_image(img_path, tmpdir, page_idx)

    # ── OCR ────────────────────────────────────────────────────────
    words = _run_ocr(preprocessed_path)

    if not words:
        logger.warning("No text detected in image: %s", img_path)
        p = create_paragraph(dst_doc)
        create_run(p, text="[No text detected in this image]")
        report.add_simple(
            location=f"Page {page_idx + 1}",
            property_name="ocr",
            confidence=Level.LOW,
            detail="OCR returned no words",
        )
        return

    # ── Group into lines and paragraphs ────────────────────────────
    lines = _group_words_into_lines(words)
    lines = _detect_math_lines(lines)
    paragraphs = _group_lines_into_paragraphs(lines)

    # ── Estimate DPI for size conversion ───────────────────────────
    img = Image.open(img_path)
    img_dpi = img.info.get('dpi', (DEFAULT_DPI, DEFAULT_DPI))
    effective_dpi = img_dpi[1] if isinstance(img_dpi, tuple) else img_dpi
    img.close()

    # ── Write to document ──────────────────────────────────────────
    para_idx = 0
    for para in paragraphs:
        para_idx += 1
        location_base = f"Page {page_idx + 1}, Para {para_idx}"

        if para.is_all_math:
            # Equation block — crop and OCR
            _handle_equation_paragraph(
                para, dst_doc, img_path, page_idx, para_idx,
                report, tmpdir, math_ocr,
            )
            continue

        dst_p = create_paragraph(dst_doc)

        for line_idx, line in enumerate(para.lines):
            if line.is_math:
                # Math line within mixed paragraph
                _handle_equation_line(
                    line, dst_p, img_path, page_idx, para_idx,
                    report, tmpdir, math_ocr,
                )
                continue

            for word in line.words:
                if word.confidence < MIN_WORD_CONFIDENCE:
                    continue

                # ── Estimate formatting from image heuristics ──
                est_size = _estimate_font_size(word.h, effective_dpi)
                est_bold = _estimate_bold(word, preprocessed_path)
                est_italic = _estimate_italic(word)

                run = create_run(
                    dst_p,
                    text=word.text + " ",
                    font_name=DEFAULT_FONT,
                    font_size=est_size,
                    bold=est_bold or None,
                    italic=est_italic or None,
                )

                # Flag low confidence items
                if word.confidence < 60:
                    highlight_low_confidence(run)

                # Report confidence for font family (always LOW for images)
                report.add_simple(
                    location=f"{location_base}, Word '{word.text[:20]}'",
                    property_name="font_family",
                    confidence=Level.LOW,
                    detail=(
                        f"Font family cannot be inferred from image pixels. "
                        f"Defaulted to '{DEFAULT_FONT}'. "
                        f"OCR confidence: {word.confidence:.0f}%"
                    ),
                    reconstructed_value=DEFAULT_FONT,
                )


def _preprocess_image(img_path: str, tmpdir: str, page_idx: int) -> str:
    """
    Preprocess an image for better OCR accuracy:
    - Convert to grayscale
    - Deskew
    """
    try:
        import cv2
    except ImportError:
        logger.warning(
            "OpenCV not installed — skipping image preprocessing. "
            "Install with: pip install opencv-python-headless"
        )
        return img_path

    img = cv2.imread(img_path)
    if img is None:
        logger.warning("Could not read image: %s", img_path)
        return img_path

    # Convert to grayscale
    logger.info("Converting page %d to grayscale", page_idx + 1)
    gray = cv2.cvtColor(img, cv2.COLOR_BGR2GRAY)

    # Deskew
    gray = _deskew(gray)

    out_path = os.path.join(tmpdir, f"preprocessed_{page_idx:04d}.png")
    cv2.imwrite(out_path, gray)

    return out_path


def _deskew(image) -> np.ndarray:
    """
    Deskew a grayscale image by detecting the skew angle from contours.
    """
    try:
        import cv2

        coords = np.column_stack(np.where(image < 128))
        if len(coords) < 100:
            return image

        angle = cv2.minAreaRect(coords)[-1]

        # Normalize angle
        if angle < -45:
            angle = -(90 + angle)
        else:
            angle = -angle

        # Only correct if skew is significant but not too extreme
        if abs(angle) < 0.5 or abs(angle) > 15:
            return image

        (h, w) = image.shape[:2]
        center = (w // 2, h // 2)
        M = cv2.getRotationMatrix2D(center, angle, 1.0)
        rotated = cv2.warpAffine(
            image, M, (w, h),
            flags=cv2.INTER_CUBIC,
            borderMode=cv2.BORDER_REPLICATE,
        )

        logger.debug("Deskewed image by %.2f degrees", angle)
        return rotated

    except Exception as e:
        logger.debug("Deskew failed: %s", e)
        return image


def _run_ocr(img_path: str) -> list[OcrWord]:
    """
    Run OCR on an image and return word-level data.
    Attempts to use Tesseract OCR first for highest accuracy across any font style.
    Falls back to Windows Media OCR (winocr) on Windows if Tesseract is not installed.
    """
    tesseract_available = False
    
    # ── 1. Attempt Tesseract OCR first ──────────────────────────────
    try:
        import pytesseract
        import shutil
        
        # Auto-configure Tesseract path on Windows if not already on system PATH
        if os.name == 'nt' and not shutil.which("tesseract"):
            user_path = os.path.expandvars(r"%LOCALAPPDATA%\Programs\Tesseract-OCR\tesseract.exe")
            sys_path = r"C:\Program Files\Tesseract-OCR\tesseract.exe"
            if os.path.exists(user_path):
                pytesseract.pytesseract.tesseract_cmd = user_path
                logger.info("Auto-configured Tesseract path (User): %s", user_path)
            elif os.path.exists(sys_path):
                pytesseract.pytesseract.tesseract_cmd = sys_path
                logger.info("Auto-configured Tesseract path (System): %s", sys_path)

        # Validate that the tesseract command is actually working
        pytesseract.get_tesseract_version()
        tesseract_available = True
    except Exception as e:
        logger.warning("Tesseract OCR is not available/configured: %s. Will attempt WinOCR fallback if on Windows.", e)

    if tesseract_available:
        try:
            from pytesseract import Output
            data = pytesseract.image_to_data(
                img_path, output_type=Output.DICT,
                config='--psm 6',  # Assume uniform block of text
            )
            words = []
            n_boxes = len(data['text'])

            for i in range(n_boxes):
                text = data['text'][i].strip()
                if not text:
                    continue

                conf = float(data['conf'][i])

                words.append(OcrWord(
                    text=text,
                    x=data['left'][i],
                    y=data['top'][i],
                    w=data['width'][i],
                    h=data['height'][i],
                    confidence=conf,
                    block_num=data['block_num'][i],
                    par_num=data['par_num'][i],
                    line_num=data['line_num'][i],
                    word_num=data['word_num'][i],
                ))

            if words:
                logger.info("Tesseract OCR successfully detected %d words", len(words))
                return words
        except Exception as e:
            logger.error("Tesseract OCR execution failed: %s. Attempting fallback.", e)

    # ── 2. Fallback to Windows Media OCR (winocr) on Windows ───────
    if os.name == 'nt':
        logger.info("Falling back to Windows Media OCR (winocr)")
        try:
            import subprocess
            import json
            import sys
            
            # Escape backslashes for python code string
            safe_img_path = img_path.replace('\\', '\\\\')
            code = f"""
import winocr, json, sys
from PIL import Image
try:
    img = Image.open(r'{safe_img_path}')
    res = winocr.recognize_pil_sync(img)
    img.close()
    print(json.dumps(res))
except Exception as e:
    print(json.dumps({{"error": str(e)}}))
    sys.exit(1)
"""
            result = subprocess.run(
                [sys.executable, "-c", code],
                capture_output=True,
                text=True,
                timeout=30
            )
            
            if result.returncode == 0:
                data = json.loads(result.stdout)
                if "error" not in data:
                    words = []
                    block_num = 0
                    par_num = 0
                    
                    for line_num, line in enumerate(data.get('lines', [])):
                        for word_num, word in enumerate(line.get('words', [])):
                            rect = word.get('bounding_rect', {})
                            words.append(OcrWord(
                                text=word.get('text', ''),
                                x=int(rect.get('x', 0)),
                                y=int(rect.get('y', 0)),
                                w=int(rect.get('width', 0)),
                                h=int(rect.get('height', 0)),
                                confidence=99.0,
                                block_num=block_num,
                                par_num=par_num,
                                line_num=line_num,
                                word_num=word_num
                            ))
                    if words:
                        logger.info("Windows Media OCR (subprocess) fallback successfully processed image with %d words", len(words))
                        return words
                else:
                    logger.debug("winocr subprocess returned error: %s", data["error"])
            else:
                logger.debug("winocr subprocess exited with code %d: %s", result.returncode, result.stderr)
        except Exception as e:
            logger.debug("Windows Media OCR subprocess fallback failed: %s", e)

    logger.warning("All OCR options failed or returned no text.")
    return []


def _group_words_into_lines(words: list[OcrWord]) -> list[OcrLine]:
    """
    Group words into lines based on Tesseract's block/par/line hierarchy.
    """
    line_map: dict[tuple, list[OcrWord]] = {}

    for word in words:
        key = (word.block_num, word.par_num, word.line_num)
        if key not in line_map:
            line_map[key] = []
        line_map[key].append(word)

    lines = []
    for key in sorted(line_map.keys()):
        line_words = sorted(line_map[key], key=lambda w: w.x)  # Sort left to right
        if line_words:
            y = min(w.y for w in line_words)
            h = max(w.y + w.h for w in line_words) - y
            lines.append(OcrLine(words=line_words, y=y, h=h))

    return lines


def _detect_math_lines(lines: list[OcrLine]) -> list[OcrLine]:
    """
    Flag lines that appear to contain math equations based on content heuristics.

    LIMITATION: This is a rough heuristic. It checks for high density of
    non-alphanumeric characters and mathematical symbols.
    """
    for line in lines:
        text = line.text
        if not text:
            continue

        # Count mathematical indicators
        math_indicators = 0
        total_chars = len(text.replace(" ", ""))

        if total_chars == 0:
            continue

        for ch in text:
            if ch in '=+−×÷∑∫∏√∂∇∞≈≠≤≥±·∈∉⊂⊃∪∩∧∨¬∀∃':
                math_indicators += 1
            elif ord(ch) > 0x0370 and ord(ch) < 0x03FF:  # Greek letters
                math_indicators += 1
            elif ord(ch) > 0x2200 and ord(ch) < 0x22FF:  # Math operators
                math_indicators += 1

        density = math_indicators / total_chars
        if density > MATH_DENSITY_THRESHOLD:
            line.is_math = True

    return lines


def _group_lines_into_paragraphs(lines: list[OcrLine]) -> list[OcrParagraph]:
    """Group lines into paragraphs based on vertical gaps."""
    if not lines:
        return []

    paragraphs = []
    current = OcrParagraph(lines=[lines[0]])

    for i in range(1, len(lines)):
        prev = lines[i - 1]
        curr = lines[i]

        gap = curr.y - (prev.y + prev.h)
        avg_height = (prev.h + curr.h) / 2

        if avg_height > 0 and gap > avg_height * PARAGRAPH_GAP_RATIO:
            paragraphs.append(current)
            current = OcrParagraph(lines=[curr])
        else:
            current.lines.append(curr)

    paragraphs.append(current)
    return paragraphs


def _estimate_font_size(word_height_px: int, dpi: float) -> float:
    """
    Estimate font size in points from word bounding box height in pixels.

    LIMITATION: This is approximate. The bounding box includes ascenders
    and descenders, so the estimated size may be slightly larger than
    the actual font size.
    """
    # Convert pixels to points: 1 point = 1/72 inch, 1 pixel = 1/dpi inch
    # The bbox height includes ascenders + descenders, roughly 1.2x the font size
    raw_pt = (word_height_px / dpi) * 72
    adjusted_pt = raw_pt / 1.2  # Compensate for ascender/descender space

    # Snap to common font sizes
    common_sizes = [8, 9, 10, 10.5, 11, 12, 14, 16, 18, 20, 22, 24, 26, 28, 36, 48, 72]
    closest = min(common_sizes, key=lambda s: abs(s - adjusted_pt))

    # Only snap if reasonably close
    if abs(closest - adjusted_pt) < 2:
        return closest
    return round(adjusted_pt, 1)


def _estimate_bold(word: OcrWord, img_path: str) -> bool:
    """
    Estimate whether a word is bold based on stroke width analysis.

    LIMITATION: This is a rough heuristic using stroke width transform.
    Accuracy depends heavily on image quality and font. Flagged as
    MEDIUM confidence in the output.
    """
    try:
        import cv2

        img = cv2.imread(img_path, cv2.IMREAD_GRAYSCALE)
        if img is None:
            return False

        # Crop the word region
        x, y, w, h = word.x, word.y, word.w, word.h
        # Clamp to image bounds
        y1 = max(0, y)
        y2 = min(img.shape[0], y + h)
        x1 = max(0, x)
        x2 = min(img.shape[1], x + w)

        if y2 <= y1 or x2 <= x1:
            return False

        crop = img[y1:y2, x1:x2]

        # Threshold
        _, binary = cv2.threshold(crop, 0, 255, cv2.THRESH_BINARY_INV + cv2.THRESH_OTSU)

        # Estimate stroke width via distance transform
        dist = cv2.distanceTransform(binary, cv2.DIST_L2, 5)
        if dist.size == 0:
            return False

        mean_dist = np.mean(dist[dist > 0]) if np.any(dist > 0) else 0

        # Bold text typically has thicker strokes
        # This threshold is empirically tuned and may need adjustment
        threshold = 2.5 * (h / 20)  # Scale with word height
        return mean_dist > threshold

    except Exception:
        return False


def _estimate_italic(word: OcrWord) -> bool:
    """
    Estimate whether a word is italic based on bounding box geometry.

    LIMITATION: Very rough heuristic. True italic detection from pixels
    would require analyzing character slant angles, which is unreliable
    without specialized models. Flagged as MEDIUM confidence.
    """
    # Simple heuristic: italic text tends to have a wider bounding box
    # relative to its height for the same text content.
    # This is too unreliable to use alone, so we default to False.
    # A more sophisticated approach would use character contour analysis.
    return False


def _handle_equation_paragraph(
    para: OcrParagraph,
    dst_doc,
    img_path: str,
    page_idx: int,
    para_idx: int,
    report: ConfidenceReport,
    tmpdir: str,
    math_ocr: str,
) -> None:
    """Handle a paragraph that is entirely a math equation."""
    dst_p = create_paragraph(dst_doc)
    location = f"Page {page_idx + 1}, Para {para_idx}"

    # Crop the equation region from the original image
    all_bboxes = [l.bbox for l in para.lines]
    merged = _merge_bboxes(all_bboxes)

    eq_img_path = _crop_region(img_path, merged, tmpdir, f"eq_p{page_idx}_{para_idx}")

    if eq_img_path:
        omml_result = image_to_omml(eq_img_path, engine=math_ocr)
        if omml_result.elements:
            for elem in omml_result.elements:
                insert_omml(dst_p, elem)
            report.add_simple(
                location=location,
                property_name="equation",
                confidence=omml_result.confidence,
                detail=omml_result.detail,
            )
            return

    # Fallback: insert placeholder
    placeholder = create_placeholder_run()
    run = create_run(dst_p, text=placeholder)
    highlight_low_confidence(run)
    report.add_simple(
        location=location,
        property_name="equation",
        confidence=Level.LOW,
        detail="Equation could not be processed from image",
    )


def _handle_equation_line(
    line: OcrLine,
    dst_p,
    img_path: str,
    page_idx: int,
    para_idx: int,
    report: ConfidenceReport,
    tmpdir: str,
    math_ocr: str,
) -> None:
    """Handle a single line flagged as math within a mixed paragraph."""
    location = f"Page {page_idx + 1}, Para {para_idx}, Math line"

    eq_img_path = _crop_region(
        img_path, line.bbox, tmpdir,
        f"eq_line_p{page_idx}_{para_idx}",
    )

    if eq_img_path:
        omml_result = image_to_omml(eq_img_path, engine=math_ocr)
        if omml_result.elements:
            for elem in omml_result.elements:
                insert_omml(dst_p, elem)
            report.add_simple(
                location=location,
                property_name="equation",
                confidence=omml_result.confidence,
                detail=omml_result.detail,
            )
            return

    # Fallback: insert the OCR text with low confidence
    run = create_run(dst_p, text=line.text + " ")
    highlight_low_confidence(run)
    report.add_simple(
        location=location,
        property_name="equation",
        confidence=Level.LOW,
        detail="Math line inserted as plain text (OCR/conversion failed)",
    )


def _crop_region(
    img_path: str,
    bbox: tuple,
    tmpdir: str,
    name: str,
    padding: int = 10,
) -> Optional[str]:
    """Crop a region from an image and save it."""
    try:
        img = Image.open(img_path)
        x0, y0, x1, y1 = bbox

        # Add padding
        x0 = max(0, x0 - padding)
        y0 = max(0, y0 - padding)
        x1 = min(img.width, x1 + padding)
        y1 = min(img.height, y1 + padding)

        crop = img.crop((x0, y0, x1, y1))
        out_path = os.path.join(tmpdir, f"{name}.png")
        crop.save(out_path)
        img.close()
        return out_path

    except Exception as e:
        logger.error("Failed to crop region: %s", e)
        return None


def _merge_bboxes(bboxes: list[tuple]) -> tuple:
    """Merge multiple bounding boxes into one encompassing all."""
    if not bboxes:
        return (0, 0, 0, 0)
    x0 = min(b[0] for b in bboxes)
    y0 = min(b[1] for b in bboxes)
    x1 = max(b[2] for b in bboxes)
    y1 = max(b[3] for b in bboxes)
    return (x0, y0, x1, y1)
