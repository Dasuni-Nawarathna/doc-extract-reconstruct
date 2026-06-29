"""
confidence.py — Confidence tracking, in-document flagging, and report generation.

Every pipeline records confidence metadata for each reconstructed element
(text runs, equations, formatting properties). Low-confidence items are
highlighted in the output document and listed in a companion report file.
"""

import logging
from dataclasses import dataclass, field
from enum import IntEnum
from typing import Optional
from pathlib import Path

logger = logging.getLogger(__name__)


class Level(IntEnum):
    """Confidence levels for reconstructed content."""
    HIGH = 3    # Exact copy / lossless (e.g., DOCX-to-DOCX text runs)
    MEDIUM = 2  # Reliable extraction with minor uncertainty (e.g., PDF font flags)
    LOW = 1     # Heuristic / guessed (e.g., image-based font detection)


@dataclass
class ConfidenceEntry:
    """A single confidence record for a reconstructed element."""
    location: str           # Human-readable location, e.g., "Page 1, Para 3, Run 2"
    property_name: str      # What was reconstructed, e.g., "font_family", "equation"
    confidence: Level       # HIGH / MEDIUM / LOW
    detail: str             # Explanation, e.g., "Font inferred from image heuristic"
    original_value: Optional[str] = None  # The original value if known
    reconstructed_value: Optional[str] = None  # What was written to the output


@dataclass
class ConfidenceReport:
    """
    Aggregates confidence entries across an entire document reconstruction.

    Usage:
        report = ConfidenceReport()
        report.add(ConfidenceEntry(...))
        ...
        report.write("output_report.txt")
    """
    entries: list[ConfidenceEntry] = field(default_factory=list)
    source_file: str = ""
    output_file: str = ""

    def add(self, entry: ConfidenceEntry) -> None:
        """Add a confidence entry."""
        self.entries.append(entry)
        if entry.confidence <= Level.LOW:
            logger.debug(
                "LOW confidence: %s — %s: %s",
                entry.location, entry.property_name, entry.detail,
            )

    def add_simple(
        self,
        location: str,
        property_name: str,
        confidence: Level,
        detail: str,
        original_value: Optional[str] = None,
        reconstructed_value: Optional[str] = None,
    ) -> None:
        """Convenience method to add an entry without constructing a dataclass."""
        self.add(ConfidenceEntry(
            location=location,
            property_name=property_name,
            confidence=confidence,
            detail=detail,
            original_value=original_value,
            reconstructed_value=reconstructed_value,
        ))

    @property
    def high_count(self) -> int:
        return sum(1 for e in self.entries if e.confidence == Level.HIGH)

    @property
    def medium_count(self) -> int:
        return sum(1 for e in self.entries if e.confidence == Level.MEDIUM)

    @property
    def low_count(self) -> int:
        return sum(1 for e in self.entries if e.confidence == Level.LOW)

    @property
    def total_count(self) -> int:
        return len(self.entries)

    def get_low_entries(self) -> list[ConfidenceEntry]:
        """Return all LOW confidence entries."""
        return [e for e in self.entries if e.confidence == Level.LOW]

    def get_medium_entries(self) -> list[ConfidenceEntry]:
        """Return all MEDIUM confidence entries."""
        return [e for e in self.entries if e.confidence == Level.MEDIUM]

    def write(self, report_path: Optional[str] = None) -> str:
        """
        Write the confidence report to a text file.

        Args:
            report_path: Path for the report file. If None, derived from output_file.

        Returns:
            Path to the written report file.
        """
        if report_path is None:
            if self.output_file:
                base = Path(self.output_file).stem
                parent = Path(self.output_file).parent
                report_path = str(parent / f"{base}_confidence_report.txt")
            else:
                report_path = "confidence_report.txt"

        lines = []
        lines.append("=" * 72)
        lines.append("DOC-EXTRACT-RECONSTRUCT — CONFIDENCE REPORT")
        lines.append("=" * 72)
        lines.append(f"Source file : {self.source_file}")
        lines.append(f"Output file : {self.output_file}")
        lines.append("")
        lines.append(f"Total elements tracked : {self.total_count}")
        lines.append(f"  HIGH confidence      : {self.high_count}")
        lines.append(f"  MEDIUM confidence    : {self.medium_count}")
        lines.append(f"  LOW confidence       : {self.low_count}")
        lines.append("")

        # Detail LOW confidence items first (most actionable)
        low_entries = self.get_low_entries()
        if low_entries:
            lines.append("-" * 72)
            lines.append("LOW CONFIDENCE ITEMS (highlighted in output document)")
            lines.append("-" * 72)
            for i, entry in enumerate(low_entries, 1):
                lines.append(f"\n  [{i}] {entry.location}")
                lines.append(f"      Property: {entry.property_name}")
                lines.append(f"      Detail  : {entry.detail}")
                if entry.original_value:
                    lines.append(f"      Original: {entry.original_value}")
                if entry.reconstructed_value:
                    lines.append(f"      Reconstructed: {entry.reconstructed_value}")
            lines.append("")

        # Then MEDIUM confidence items
        medium_entries = self.get_medium_entries()
        if medium_entries:
            lines.append("-" * 72)
            lines.append("MEDIUM CONFIDENCE ITEMS")
            lines.append("-" * 72)
            for i, entry in enumerate(medium_entries, 1):
                lines.append(f"\n  [{i}] {entry.location}")
                lines.append(f"      Property: {entry.property_name}")
                lines.append(f"      Detail  : {entry.detail}")
                if entry.original_value:
                    lines.append(f"      Original: {entry.original_value}")
                if entry.reconstructed_value:
                    lines.append(f"      Reconstructed: {entry.reconstructed_value}")
            lines.append("")

        if not low_entries and not medium_entries:
            lines.append("All elements reconstructed with HIGH confidence. ✓")
            lines.append("")

        lines.append("=" * 72)
        lines.append("END OF REPORT")
        lines.append("=" * 72)

        report_text = "\n".join(lines)

        with open(report_path, "w", encoding="utf-8") as f:
            f.write(report_text)

        logger.info("Confidence report written to: %s", report_path)
        return report_path
