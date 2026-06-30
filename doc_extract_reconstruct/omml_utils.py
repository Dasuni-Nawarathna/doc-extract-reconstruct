"""
omml_utils.py — Math-OCR-to-OMML conversion utility.

Shared by the PDF and image pipelines. Provides:
  1. Math OCR: image → LaTeX (via pix2tex or rapid-latex-ocr)
  2. LaTeX → OMML: LaTeX → temp .docx via Pandoc → extract OMML XML nodes
  3. Fallback: insert placeholder text if tools are unavailable

KNOWN LIMITATION: Reconstructed equations (via OCR + LaTeX → OMML) are
approximations of the original. Complex multi-line equations, matrices,
or heavily styled equations may not round-trip perfectly.
"""

import logging
import os
import shutil
import subprocess
import tempfile
from copy import deepcopy
from dataclasses import dataclass
from pathlib import Path
from typing import Optional

from lxml import etree

from .writer import qn, NSMAP
from .confidence import Level

logger = logging.getLogger(__name__)


@dataclass
class MathOCRResult:
    """Result from math OCR recognition."""
    latex: str
    confidence: float  # 0.0 to 1.0
    engine: str        # e.g., "pix2tex", "rapid_latex_ocr", "none"
    error: Optional[str] = None


@dataclass
class OmmlResult:
    """Result from LaTeX → OMML conversion."""
    elements: list  # List of lxml <m:oMath> or <m:oMathPara> elements
    confidence: Level
    detail: str
    error: Optional[str] = None


# ── System dependency checks ───────────────────────────────────────────

def _get_pandoc_cmd() -> str:
    """Get the command to execute Pandoc (auto-discovering absolute path on Windows)."""
    cmd = shutil.which("pandoc")
    if cmd is not None:
        return cmd
        
    if os.name == 'nt':
        # Check standard user and system installation paths
        paths = [
            os.path.expandvars(r"%LOCALAPPDATA%\Pandoc\pandoc.exe"),
            os.path.expandvars(r"%LOCALAPPDATA%\Programs\Pandoc\pandoc.exe"),
            r"C:\Program Files\Pandoc\pandoc.exe",
            r"C:\Program Files (x86)\Pandoc\pandoc.exe",
        ]
        for p in paths:
            if os.path.exists(p):
                return p
    return "pandoc"


def _check_pandoc() -> bool:
    """Check if Pandoc is available."""
    cmd = _get_pandoc_cmd()
    if cmd == "pandoc":
        return shutil.which("pandoc") is not None
    return os.path.exists(cmd)


def _check_pix2tex() -> bool:
    """Check if pix2tex is importable."""
    try:
        from pix2tex.cli import LatexOCR  # noqa: F401
        return True
    except ImportError:
        return False


def _check_rapid_latex_ocr() -> bool:
    """Check if rapid-latex-ocr is importable."""
    try:
        from rapid_latex_ocr import LatexOCR as RapidLatexOCR  # noqa: F401
        return True
    except ImportError:
        return False


# ── Math OCR: Image → LaTeX ───────────────────────────────────────────

def image_to_latex(
    image_path: str,
    engine: str = "auto",
) -> MathOCRResult:
    """
    Recognize a math equation from an image and return LaTeX.

    Args:
        image_path: Path to the cropped equation image.
        engine: Which OCR engine to use:
            "auto"  — try pix2tex first, then rapid-latex-ocr
            "pix2tex" — use pix2tex only
            "rapid" — use rapid-latex-ocr only
            "none"  — skip OCR, return placeholder

    Returns:
        MathOCRResult with the recognized LaTeX and confidence score.
    """
    if engine == "none":
        return MathOCRResult(
            latex="",
            confidence=0.0,
            engine="none",
            error="Math OCR disabled by user",
        )

    # Try pix2tex
    if engine in ("auto", "pix2tex") and _check_pix2tex():
        return _ocr_pix2tex(image_path)

    # Try rapid-latex-ocr
    if engine in ("auto", "rapid") and _check_rapid_latex_ocr():
        return _ocr_rapid(image_path)

    # Nothing available
    logger.warning(
        "No math OCR engine available. Install pix2tex or rapid-latex-ocr. "
        "Equation will be inserted as placeholder text."
    )
    return MathOCRResult(
        latex="",
        confidence=0.0,
        engine="none",
        error="No math OCR engine installed (need pix2tex or rapid-latex-ocr)",
    )


def _ocr_pix2tex(image_path: str) -> MathOCRResult:
    """Run pix2tex (LatexOCR) on an equation image."""
    try:
        from pix2tex.cli import LatexOCR
        from PIL import Image

        model = LatexOCR()
        img = Image.open(image_path)
        latex = model(img)

        # pix2tex doesn't provide a confidence score natively,
        # so we use a heuristic: non-empty result = 0.7 confidence
        conf = 0.7 if latex and len(latex.strip()) > 0 else 0.2

        logger.info("pix2tex recognized: %s (confidence: %.2f)", latex[:80], conf)
        return MathOCRResult(latex=latex, confidence=conf, engine="pix2tex")

    except Exception as e:
        logger.error("pix2tex error: %s", e)
        return MathOCRResult(
            latex="", confidence=0.0, engine="pix2tex",
            error=str(e),
        )


def _ocr_rapid(image_path: str) -> MathOCRResult:
    """Run rapid-latex-ocr on an equation image."""
    try:
        from rapid_latex_ocr import LatexOCR as RapidLatexOCR

        model = RapidLatexOCR()
        result, elapse = model(image_path)

        conf = 0.65 if result and len(result.strip()) > 0 else 0.2

        logger.info("rapid-latex-ocr recognized: %s (confidence: %.2f)", result[:80], conf)
        return MathOCRResult(latex=result, confidence=conf, engine="rapid_latex_ocr")

    except Exception as e:
        logger.error("rapid-latex-ocr error: %s", e)
        return MathOCRResult(
            latex="", confidence=0.0, engine="rapid_latex_ocr",
            error=str(e),
        )


# ── LaTeX → OMML conversion via Pandoc ─────────────────────────────────

def latex_to_omml(latex: str) -> OmmlResult:
    """
    Convert a LaTeX math expression to OMML (Office MathML) elements
    using Pandoc as the conversion engine.

    Pipeline:
      1. Wrap LaTeX in a minimal .tex document with math delimiters
      2. Run `pandoc -f latex -t docx -o temp.docx`
      3. Open temp.docx and extract <m:oMath> / <m:oMathPara> elements
      4. Return deep-copied OMML elements

    Args:
        latex: LaTeX math string (without delimiters — they are added here).

    Returns:
        OmmlResult with extracted OMML elements.
    """
    if not _check_pandoc():
        logger.error(
            "Pandoc is not installed. Cannot convert LaTeX to OMML. "
            "Install from https://pandoc.org/installing.html"
        )
        return OmmlResult(
            elements=[],
            confidence=Level.LOW,
            detail="Pandoc not installed — cannot convert LaTeX to OMML",
            error="Pandoc not found on PATH",
        )

    if not latex or not latex.strip():
        return OmmlResult(
            elements=[],
            confidence=Level.LOW,
            detail="Empty LaTeX expression",
            error="Empty LaTeX input",
        )

    # Clean the LaTeX string
    latex = latex.strip()

    # Ensure it has math delimiters — if not already wrapped
    if not latex.startswith('$') and not latex.startswith('\\[') \
       and not latex.startswith('\\begin'):
        # Wrap as display math for block equations, inline for short ones
        if '\n' in latex or '\\\\' in latex or 'begin{' in latex:
            latex_wrapped = f"\\[{latex}\\]"
        else:
            latex_wrapped = f"${latex}$"
    else:
        latex_wrapped = latex

    # Create a minimal LaTeX document
    tex_content = f"\\documentclass{{article}}\n\\begin{{document}}\n{latex_wrapped}\n\\end{{document}}\n"

    tmpdir = tempfile.mkdtemp(prefix="doc_extract_omml_")
    try:
        tex_path = os.path.join(tmpdir, "equation.tex")
        docx_path = os.path.join(tmpdir, "equation.docx")

        with open(tex_path, "w", encoding="utf-8") as f:
            f.write(tex_content)

        # Run Pandoc
        pandoc_cmd = _get_pandoc_cmd()
        cmd = [pandoc_cmd, "-f", "latex", "-t", "docx", "-o", docx_path, tex_path]
        logger.debug("Running Pandoc: %s", " ".join(cmd))

        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            timeout=30,
        )

        if result.returncode != 0:
            error_msg = result.stderr.strip() or "Unknown Pandoc error"
            logger.error("Pandoc failed (exit %d): %s", result.returncode, error_msg)
            return OmmlResult(
                elements=[],
                confidence=Level.LOW,
                detail=f"Pandoc conversion failed: {error_msg}",
                error=error_msg,
            )

        # Extract OMML elements from the generated .docx
        elements = _extract_omml_from_docx(docx_path)

        if elements:
            logger.info("Extracted %d OMML element(s) from Pandoc output", len(elements))
            return OmmlResult(
                elements=elements,
                confidence=Level.MEDIUM,
                detail=f"LaTeX→OMML via Pandoc ({len(elements)} element(s))",
            )
        else:
            logger.warning("No OMML elements found in Pandoc output")
            return OmmlResult(
                elements=[],
                confidence=Level.LOW,
                detail="Pandoc produced .docx but no OMML math elements found",
                error="No m:oMath or m:oMathPara in Pandoc output",
            )

    except subprocess.TimeoutExpired:
        logger.error("Pandoc timed out after 30 seconds")
        return OmmlResult(
            elements=[],
            confidence=Level.LOW,
            detail="Pandoc conversion timed out",
            error="Pandoc timeout",
        )
    except Exception as e:
        logger.error("LaTeX→OMML conversion error: %s", e)
        return OmmlResult(
            elements=[],
            confidence=Level.LOW,
            detail=f"LaTeX→OMML error: {e}",
            error=str(e),
        )
    finally:
        # Clean up temp directory
        shutil.rmtree(tmpdir, ignore_errors=True)


def _extract_omml_from_docx(docx_path: str) -> list:
    """
    Open a .docx file and extract all <m:oMath> and <m:oMathPara> elements
    from the document body. Returns deep-copied elements.
    """
    from docx import Document

    doc = Document(docx_path)
    body = doc.element.body

    m_omath = qn('m:oMath')
    m_omathpara = qn('m:oMathPara')

    elements = []

    # Search recursively for math elements
    for elem in body.iter():
        if elem.tag in (m_omath, m_omathpara):
            # Check that this is not a child of another math element we already captured
            parent = elem.getparent()
            if parent is not None and parent.tag in (m_omath, m_omathpara):
                continue  # Skip nested — the parent will be copied whole
            elements.append(deepcopy(elem))

    return elements


def image_to_omml(
    image_path: str,
    engine: str = "auto",
) -> OmmlResult:
    """
    Full pipeline: image → LaTeX (via OCR) → OMML (via Pandoc).

    This is the main entry point for the PDF and image pipelines when
    they encounter an equation region.

    Args:
        image_path: Path to the cropped equation image.
        engine: Math OCR engine selection (see image_to_latex).

    Returns:
        OmmlResult with OMML elements ready for insertion.
    """
    # Step 1: Image → LaTeX
    ocr_result = image_to_latex(image_path, engine=engine)

    if ocr_result.error or not ocr_result.latex:
        return OmmlResult(
            elements=[],
            confidence=Level.LOW,
            detail=f"Math OCR failed: {ocr_result.error or 'empty result'}",
            error=ocr_result.error,
        )

    # Step 2: LaTeX → OMML
    omml_result = latex_to_omml(ocr_result.latex)

    # Adjust confidence based on OCR confidence
    if ocr_result.confidence < 0.5 and omml_result.confidence > Level.LOW:
        omml_result.confidence = Level.LOW
        omml_result.detail += f" (OCR confidence low: {ocr_result.confidence:.2f})"

    return omml_result


def create_placeholder_run(text: str = "[EQUATION — could not be processed]") -> str:
    """
    Return placeholder text for equations that couldn't be converted.
    The caller should insert this via writer.create_run() with LOW confidence.
    """
    return text
