"""
writer.py — Common DOCX output writer used by all three pipelines.

Provides helper functions for creating runs with full formatting,
inserting OMML math elements, setting paragraph properties, and
flagging low-confidence content with yellow highlighting.

All functions operate at the lxml level (via python-docx's internal
element API) to avoid dropping unknown XML elements.
"""

import logging
from copy import deepcopy
from typing import Optional, Tuple

from docx import Document
from docx.shared import Pt, RGBColor, Emu
from docx.enum.text import WD_ALIGN_PARAGRAPH
from lxml import etree

logger = logging.getLogger(__name__)

# ── XML Namespaces ──────────────────────────────────────────────────────
NSMAP = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'm': 'http://schemas.openxmlformats.org/officeDocument/2006/math',
    'r': 'http://schemas.openxmlformats.org/officeDocument/2006/relationships',
    'wp': 'http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing',
    'a': 'http://schemas.openxmlformats.org/drawingml/2006/main',
}


def qn(tag: str) -> str:
    """
    Convert a namespace-prefixed tag to Clark notation.
    E.g., 'w:rPr' -> '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}rPr'
    """
    prefix, local = tag.split(':', 1)
    return f'{{{NSMAP[prefix]}}}{local}'


def create_document() -> Document:
    """Create a new empty Document for output."""
    return Document()


def append_to_body(doc: Document, element: etree._Element) -> None:
    """
    Append an element to the document body, ensuring it is placed
    before the <w:sectPr> element to maintain OpenXML schema validity.
    """
    body = doc.element.body
    sectPr = body.find(qn('w:sectPr'))
    if sectPr is not None:
        sectPr.addprevious(element)
    else:
        body.append(element)


def create_paragraph(doc: Document) -> etree._Element:
    """
    Create a new <w:p> element and insert it into the document body
    before <w:sectPr>.
    """
    p = etree.Element(qn('w:p'))
    append_to_body(doc, p)
    return p


def set_paragraph_properties(
    p_element: etree._Element,
    alignment: Optional[str] = None,
    spacing_before: Optional[float] = None,
    spacing_after: Optional[float] = None,
    line_spacing: Optional[float] = None,
    indent_left: Optional[float] = None,
    indent_right: Optional[float] = None,
    indent_first_line: Optional[float] = None,
    numbering_id: Optional[int] = None,
    numbering_level: Optional[int] = None,
) -> None:
    """
    Set paragraph-level properties on a <w:p> element.

    All size values are in points (Pt). Alignment values: 'left', 'center',
    'right', 'both' (justify).
    """
    # Get or create <w:pPr>
    pPr = p_element.find(qn('w:pPr'))
    if pPr is None:
        pPr = etree.SubElement(p_element, qn('w:pPr'))
        # Insert pPr as the first child
        p_element.insert(0, pPr)

    # Alignment
    if alignment:
        jc = etree.SubElement(pPr, qn('w:jc'))
        alignment_map = {
            'left': 'left', 'center': 'center',
            'right': 'right', 'both': 'both', 'justify': 'both',
        }
        jc.set(qn('w:val'), alignment_map.get(alignment, alignment))

    # Spacing
    if any(v is not None for v in [spacing_before, spacing_after, line_spacing]):
        spacing = etree.SubElement(pPr, qn('w:spacing'))
        if spacing_before is not None:
            spacing.set(qn('w:before'), str(int(spacing_before * 20)))  # Pt -> twips
        if spacing_after is not None:
            spacing.set(qn('w:after'), str(int(spacing_after * 20)))
        if line_spacing is not None:
            # line_spacing as a multiplier (e.g., 1.5 = 1.5x)
            spacing.set(qn('w:line'), str(int(line_spacing * 240)))
            spacing.set(qn('w:lineRule'), 'auto')

    # Indentation
    if any(v is not None for v in [indent_left, indent_right, indent_first_line]):
        ind = etree.SubElement(pPr, qn('w:ind'))
        if indent_left is not None:
            ind.set(qn('w:left'), str(int(indent_left * 20)))
        if indent_right is not None:
            ind.set(qn('w:right'), str(int(indent_right * 20)))
        if indent_first_line is not None:
            ind.set(qn('w:firstLine'), str(int(indent_first_line * 20)))

    # Numbering (list/bullet)
    if numbering_id is not None:
        numPr = etree.SubElement(pPr, qn('w:numPr'))
        ilvl = etree.SubElement(numPr, qn('w:ilvl'))
        ilvl.set(qn('w:val'), str(numbering_level or 0))
        numId = etree.SubElement(numPr, qn('w:numId'))
        numId.set(qn('w:val'), str(numbering_id))


def copy_paragraph_properties(
    source_pPr: etree._Element,
    target_p: etree._Element,
) -> None:
    """
    Deep-copy <w:pPr> from a source element to a target <w:p>.
    Replaces any existing pPr on the target.
    """
    # Remove existing pPr if present
    existing = target_p.find(qn('w:pPr'))
    if existing is not None:
        target_p.remove(existing)

    new_pPr = deepcopy(source_pPr)
    target_p.insert(0, new_pPr)


def create_run(
    p_element: etree._Element,
    text: str,
    font_name: Optional[str] = None,
    font_size: Optional[float] = None,
    bold: Optional[bool] = None,
    italic: Optional[bool] = None,
    underline: Optional[bool] = None,
    color: Optional[Tuple[int, int, int]] = None,
    highlight: Optional[str] = None,
    strike: Optional[bool] = None,
    superscript: Optional[bool] = None,
    subscript: Optional[bool] = None,
) -> etree._Element:
    """
    Create a <w:r> run element with full formatting and append it to a <w:p>.

    Args:
        p_element: The parent <w:p> element.
        text: The text content for this run.
        font_name: Font family name (e.g., "Calibri", "Cambria Math").
        font_size: Font size in points.
        bold: Whether the text is bold.
        italic: Whether the text is italic.
        underline: Whether the text is underlined.
        color: RGB color as (r, g, b) tuple, each 0-255.
        highlight: Highlight color name (e.g., "yellow", "green").
        strike: Strikethrough.
        superscript: Superscript text.
        subscript: Subscript text.

    Returns:
        The created <w:r> element.
    """
    r = etree.SubElement(p_element, qn('w:r'))

    # Build <w:rPr> (run properties) — only if there are properties to set
    has_props = any(v is not None for v in [
        font_name, font_size, bold, italic, underline, color,
        highlight, strike, superscript, subscript,
    ])

    if has_props:
        rPr = etree.SubElement(r, qn('w:rPr'))

        if font_name:
            rFonts = etree.SubElement(rPr, qn('w:rFonts'))
            rFonts.set(qn('w:ascii'), font_name)
            rFonts.set(qn('w:hAnsi'), font_name)
            rFonts.set(qn('w:cs'), font_name)

        if bold:
            etree.SubElement(rPr, qn('w:b'))

        if italic:
            etree.SubElement(rPr, qn('w:i'))

        if underline:
            u_elem = etree.SubElement(rPr, qn('w:u'))
            u_elem.set(qn('w:val'), 'single')

        if strike:
            etree.SubElement(rPr, qn('w:strike'))

        if color:
            c_elem = etree.SubElement(rPr, qn('w:color'))
            c_elem.set(qn('w:val'), '{:02X}{:02X}{:02X}'.format(*color))

        if font_size is not None:
            sz = etree.SubElement(rPr, qn('w:sz'))
            sz.set(qn('w:val'), str(int(font_size * 2)))  # half-points
            szCs = etree.SubElement(rPr, qn('w:szCs'))
            szCs.set(qn('w:val'), str(int(font_size * 2)))

        if highlight:
            hl = etree.SubElement(rPr, qn('w:highlight'))
            hl.set(qn('w:val'), highlight)

        if superscript:
            va = etree.SubElement(rPr, qn('w:vertAlign'))
            va.set(qn('w:val'), 'superscript')

        if subscript:
            va = etree.SubElement(rPr, qn('w:vertAlign'))
            va.set(qn('w:val'), 'subscript')

    # Build <w:t> text element
    t = etree.SubElement(r, qn('w:t'))
    t.text = text
    # Preserve leading/trailing whitespace
    t.set('{http://www.w3.org/XML/1998/namespace}space', 'preserve')

    return r


def copy_run_properties(
    source_rPr: etree._Element,
    target_r: etree._Element,
) -> None:
    """
    Deep-copy <w:rPr> from a source element into a target <w:r>.
    Replaces any existing rPr on the target run.
    """
    existing = target_r.find(qn('w:rPr'))
    if existing is not None:
        target_r.remove(existing)

    new_rPr = deepcopy(source_rPr)
    # Insert rPr as first child of the run
    target_r.insert(0, new_rPr)


def insert_omml(
    p_element: etree._Element,
    omml_element: etree._Element,
) -> None:
    """
    Insert an OMML math element (<m:oMath> or <m:oMathPara>) into a paragraph.

    The element is deep-copied before insertion to avoid ownership issues
    when moving elements between documents.
    """
    copied = deepcopy(omml_element)
    p_element.append(copied)
    logger.debug("Inserted OMML element <%s> into paragraph", omml_element.tag)


def highlight_low_confidence(run_element: etree._Element) -> None:
    """
    Apply yellow highlighting to a run element to flag low-confidence content.

    If the run already has an <w:rPr>, the highlight is added to it.
    If not, a new <w:rPr> is created.
    """
    rPr = run_element.find(qn('w:rPr'))
    if rPr is None:
        rPr = etree.SubElement(run_element, qn('w:rPr'))
        run_element.insert(0, rPr)

    # Remove existing highlight if any
    existing_hl = rPr.find(qn('w:highlight'))
    if existing_hl is not None:
        rPr.remove(existing_hl)

    hl = etree.SubElement(rPr, qn('w:highlight'))
    hl.set(qn('w:val'), 'yellow')


def insert_comment_text(
    p_element: etree._Element,
    comment_text: str,
) -> None:
    """
    Insert a bracketed comment as a distinct run at the end of a paragraph.
    Styled in red, smaller font, to visually stand out.
    """
    create_run(
        p_element,
        text=f" [{comment_text}]",
        font_size=8,
        color=(255, 0, 0),
        italic=True,
    )


def save_document(doc: Document, output_path: str) -> str:
    """Save the document and return the output path."""
    doc.save(output_path)
    logger.info("Document saved: %s", output_path)
    return output_path
