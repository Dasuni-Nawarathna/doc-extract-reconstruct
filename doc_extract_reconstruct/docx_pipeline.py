"""
docx_pipeline.py — DOCX-to-DOCX reconstruction pipeline (highest fidelity).

Reads every paragraph in the source .docx at the raw XML level, preserving:
  - Text runs with full w:rPr properties (font, size, bold, italic, underline,
    color, highlight, strikethrough, super/subscript)
  - Math equations (<m:oMath>, <m:oMathPara>) via verbatim deep-copy of the XML
    nodes, preserving native Word equations exactly (including Cambria Math)
  - Paragraph properties (alignment, spacing, indentation, numbering)
  - Hyperlinks and other structural elements via deep-copy

This is the highest-fidelity pipeline — near-lossless for text and equations.
"""

import logging
from copy import deepcopy
from typing import Optional

from docx import Document
from lxml import etree

from .writer import (
    qn, NSMAP, create_document, create_paragraph,
    copy_paragraph_properties, create_run, copy_run_properties,
    insert_omml, save_document, highlight_low_confidence,
    insert_comment_text, append_to_body,
)
from .confidence import ConfidenceReport, ConfidenceEntry, Level

logger = logging.getLogger(__name__)

# ── Namespace constants ────────────────────────────────────────────────
W_NS = NSMAP['w']
M_NS = NSMAP['m']

# Tags we recognize and process
W_P = qn('w:p')          # paragraph
W_R = qn('w:r')          # text run
W_RPR = qn('w:rPr')      # run properties
W_PPR = qn('w:pPr')      # paragraph properties
W_T = qn('w:t')          # text
W_TAB = qn('w:tab')      # tab character
W_BR = qn('w:br')        # break (page/line/column)
W_HYPERLINK = qn('w:hyperlink')
W_BOOKMARK_START = qn('w:bookmarkStart')
W_BOOKMARK_END = qn('w:bookmarkEnd')
W_TBL = qn('w:tbl')      # table
W_SDT = qn('w:sdt')      # structured document tag
W_SECTPR = qn('w:sectPr')  # section properties

M_OMATH = qn('m:oMath')
M_OMATHPARA = qn('m:oMathPara')


def process_docx(
    input_path: str,
    output_path: str,
    **kwargs,
) -> str:
    """
    Read a .docx file and reconstruct it as a new .docx preserving all
    formatting and math equations.

    Args:
        input_path: Path to the source .docx file.
        output_path: Path for the reconstructed output .docx file.

    Returns:
        Path to the generated output file.
    """
    logger.info("DOCX Pipeline: Processing '%s'", input_path)

    report = ConfidenceReport(source_file=input_path, output_file=output_path)
    src_doc = Document(input_path)
    dst_doc = create_document()

    # Remove the default empty paragraph that python-docx creates
    _remove_default_paragraph(dst_doc)

    # Process every top-level element in the document body
    body = src_doc.element.body
    para_index = 0

    for child in body:
        tag = child.tag

        if tag == W_P:
            para_index += 1
            _process_paragraph(child, dst_doc, para_index, report)

        elif tag == W_TBL:
            # Tables: deep-copy the entire table element
            _process_table(child, dst_doc, para_index, report)

        elif tag == W_SECTPR:
            # Section properties: deep-copy to preserve page layout
            _copy_section_properties(child, dst_doc)

        elif tag == W_SDT:
            # Structured document tags (e.g., TOC): deep-copy verbatim
            append_to_body(dst_doc, deepcopy(child))
            report.add_simple(
                location=f"Body element (SDT)",
                property_name="structured_document_tag",
                confidence=Level.HIGH,
                detail="Structured document tag deep-copied verbatim",
            )

        else:
            # Unknown elements: deep-copy to preserve them
            append_to_body(dst_doc, deepcopy(child))
            logger.debug("Deep-copied unknown body element: %s", tag)

    # Save output
    save_document(dst_doc, output_path)

    # Write confidence report
    report.write()

    logger.info(
        "DOCX Pipeline complete. %d paragraphs processed. Output: '%s'",
        para_index, output_path,
    )
    return output_path


def _remove_default_paragraph(doc: Document) -> None:
    """Remove the empty default paragraph that python-docx adds to new documents."""
    body = doc.element.body
    paragraphs = body.findall(qn('w:p'))
    for p in paragraphs:
        body.remove(p)


def _process_paragraph(
    src_p: etree._Element,
    dst_doc: Document,
    para_index: int,
    report: ConfidenceReport,
) -> None:
    """
    Process a single <w:p> element from the source, reconstructing it
    in the destination document.
    """
    dst_p = create_paragraph(dst_doc)

    # ── Copy paragraph properties ──────────────────────────────────
    src_pPr = src_p.find(qn('w:pPr'))
    if src_pPr is not None:
        copy_paragraph_properties(src_pPr, dst_p)
        report.add_simple(
            location=f"Paragraph {para_index}",
            property_name="paragraph_properties",
            confidence=Level.HIGH,
            detail="Paragraph properties (alignment, spacing, numbering) deep-copied",
        )

    # ── Process children in document order ─────────────────────────
    run_index = 0
    for child in src_p:
        tag = child.tag

        if tag == W_PPR:
            # Already handled above
            continue

        elif tag == W_R:
            run_index += 1
            _process_run(child, dst_p, para_index, run_index, report)

        elif tag in (M_OMATH, M_OMATHPARA):
            # Math equation: deep-copy verbatim — this preserves native
            # Word equations exactly, including Cambria Math rendering.
            insert_omml(dst_p, child)
            report.add_simple(
                location=f"Paragraph {para_index}, Math element",
                property_name="equation",
                confidence=Level.HIGH,
                detail="OMML equation deep-copied verbatim from source",
            )

        elif tag == W_HYPERLINK:
            # Hyperlinks contain runs — deep-copy the entire element
            # to preserve the relationship reference and formatting
            dst_p.append(deepcopy(child))
            report.add_simple(
                location=f"Paragraph {para_index}, Hyperlink",
                property_name="hyperlink",
                confidence=Level.HIGH,
                detail="Hyperlink element deep-copied verbatim",
            )

        elif tag in (W_BOOKMARK_START, W_BOOKMARK_END):
            dst_p.append(deepcopy(child))

        else:
            # Any other inline element: deep-copy to preserve
            dst_p.append(deepcopy(child))
            logger.debug(
                "Para %d: deep-copied unknown inline element: %s",
                para_index, tag,
            )


def _process_run(
    src_r: etree._Element,
    dst_p: etree._Element,
    para_index: int,
    run_index: int,
    report: ConfidenceReport,
) -> None:
    """
    Process a single <w:r> run element, preserving all formatting properties.
    """
    location = f"Paragraph {para_index}, Run {run_index}"

    # Create the destination run
    dst_r = etree.SubElement(dst_p, qn('w:r'))

    # ── Copy run properties (w:rPr) ────────────────────────────────
    src_rPr = src_r.find(qn('w:rPr'))
    if src_rPr is not None:
        copy_run_properties(src_rPr, dst_r)
        # Log what we preserved
        _log_run_properties(src_rPr, location, report)

    # ── Copy run content (w:t, w:tab, w:br, etc.) ──────────────────
    for child in src_r:
        tag = child.tag

        if tag == W_RPR:
            # Already handled above
            continue

        elif tag == W_T:
            # Text element: create new w:t with same text and xml:space
            dst_t = etree.SubElement(dst_r, qn('w:t'))
            dst_t.text = child.text
            # Preserve xml:space="preserve" attribute
            space_attr = child.get('{http://www.w3.org/XML/1998/namespace}space')
            if space_attr:
                dst_t.set('{http://www.w3.org/XML/1998/namespace}space', space_attr)

        elif tag == W_TAB:
            etree.SubElement(dst_r, qn('w:tab'))

        elif tag == W_BR:
            br = etree.SubElement(dst_r, qn('w:br'))
            # Copy break type attribute if present
            break_type = child.get(qn('w:type'))
            if break_type:
                br.set(qn('w:type'), break_type)

        else:
            # Other run-level content (e.g., drawing, footnote ref): deep-copy
            dst_r.append(deepcopy(child))
            logger.debug(
                "%s: deep-copied run child element: %s",
                location, tag,
            )


def _log_run_properties(
    rPr: etree._Element,
    location: str,
    report: ConfidenceReport,
) -> None:
    """
    Log the formatting properties found in a run for the confidence report.
    All DOCX-to-DOCX run properties are HIGH confidence.
    """
    props_found = []

    rFonts = rPr.find(qn('w:rFonts'))
    if rFonts is not None:
        font = rFonts.get(qn('w:ascii')) or rFonts.get(qn('w:hAnsi')) or "?"
        props_found.append(f"font={font}")

    sz = rPr.find(qn('w:sz'))
    if sz is not None:
        half_pts = sz.get(qn('w:val'))
        if half_pts:
            props_found.append(f"size={int(half_pts) / 2}pt")

    if rPr.find(qn('w:b')) is not None:
        props_found.append("bold")
    if rPr.find(qn('w:i')) is not None:
        props_found.append("italic")
    if rPr.find(qn('w:u')) is not None:
        props_found.append("underline")

    color = rPr.find(qn('w:color'))
    if color is not None:
        props_found.append(f"color=#{color.get(qn('w:val'), '?')}")

    highlight = rPr.find(qn('w:highlight'))
    if highlight is not None:
        props_found.append(f"highlight={highlight.get(qn('w:val'), '?')}")

    if props_found:
        report.add_simple(
            location=location,
            property_name="run_formatting",
            confidence=Level.HIGH,
            detail=f"Formatting deep-copied: {', '.join(props_found)}",
        )


def _process_table(
    src_tbl: etree._Element,
    dst_doc: Document,
    para_index: int,
    report: ConfidenceReport,
) -> None:
    """
    Deep-copy an entire table element to the output document.
    Tables contain their own paragraphs, runs, and potentially math — all
    preserved by the deep-copy.
    """
    append_to_body(dst_doc, deepcopy(src_tbl))
    report.add_simple(
        location=f"Table (after paragraph {para_index})",
        property_name="table",
        confidence=Level.HIGH,
        detail="Entire table deep-copied verbatim (preserves all content and formatting)",
    )


def _copy_section_properties(
    src_sectPr: etree._Element,
    dst_doc: Document,
) -> None:
    """
    Copy section properties (page size, margins, headers/footers).
    Section properties are typically the last child of w:body.
    """
    body = dst_doc.element.body

    # Remove existing sectPr if present
    existing = body.find(qn('w:sectPr'))
    if existing is not None:
        body.remove(existing)

    body.append(deepcopy(src_sectPr))
    logger.debug("Copied section properties to output document")
