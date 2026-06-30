"""
cli.py — Command-line interface for doc-extract-reconstruct.

Usage:
    python extract.py <input_file> -o <output.docx> [options]

Options:
    -o, --output          Output .docx path (default: <input>_reconstructed.docx)
    --math-ocr            Math OCR engine: auto|pix2tex|rapid|none (default: auto)
    --dpi                 DPI for image/PDF rendering (default: 300)
    --confidence-report   Path for confidence report (default: alongside output)
    --verbose             Enable detailed logging
    --quiet               Suppress all output except errors
"""

import argparse
import logging
import sys
import os
from pathlib import Path

from . import __version__
from .router import route, detect_input_type


def build_parser() -> argparse.ArgumentParser:
    """Build the argument parser."""
    parser = argparse.ArgumentParser(
        prog="doc-extract-reconstruct",
        description=(
            "Extract content from .docx, .pdf, or image files and reconstruct "
            "as a new .docx file preserving formatting and math equations."
        ),
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=(
            "Examples:\n"
            "  python extract.py document.docx -o output.docx\n"
            "  python extract.py report.pdf -o report_reconstructed.docx\n"
            "  python extract.py scan.png -o scan_output.docx --math-ocr pix2tex\n"
            "  python extract.py paper.pdf -o paper.docx --dpi 400 --verbose\n"
            "\n"
            "Supported input formats: .docx, .pdf, .png, .jpg, .jpeg, .tiff, .bmp\n"
            "\n"
            "For more details, see the README.md file."
        ),
    )

    parser.add_argument(
        "input",
        help="Input file path (.docx, .pdf, .png, .jpg, etc.)",
    )

    parser.add_argument(
        "-o", "--output",
        help=(
            "Output .docx file path. "
            "Default: <input_basename>_reconstructed.docx"
        ),
        default=None,
    )

    parser.add_argument(
        "--math-ocr",
        choices=["auto", "pix2tex", "rapid", "none"],
        default="auto",
        help=(
            "Math OCR engine for equation recognition. "
            "'auto' tries pix2tex first, then rapid-latex-ocr. "
            "'none' disables equation OCR. (default: auto)"
        ),
    )

    parser.add_argument(
        "--dpi",
        type=int,
        default=300,
        help="DPI for rendering PDF pages and equation crops. (default: 300)",
    )

    parser.add_argument(
        "--confidence-report",
        help=(
            "Path for the confidence report file. "
            "Default: <output_basename>_confidence_report.txt"
        ),
        default=None,
    )

    parser.add_argument(
        "--verbose", "-v",
        action="store_true",
        help="Enable detailed debug logging.",
    )

    parser.add_argument(
        "--quiet", "-q",
        action="store_true",
        help="Suppress all output except errors.",
    )

    parser.add_argument(
        "--version",
        action="version",
        version=f"%(prog)s {__version__}",
    )

    return parser


def _configure_logging(verbose: bool, quiet: bool) -> None:
    """Configure logging based on verbosity flags."""
    if quiet:
        level = logging.ERROR
    elif verbose:
        level = logging.DEBUG
    else:
        level = logging.INFO

    logging.basicConfig(
        level=level,
        format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
        datefmt="%H:%M:%S",
        handlers=[logging.StreamHandler(sys.stderr)],
    )


def _derive_output_path(input_path: str) -> str:
    """Generate a default output path from the input path."""
    p = Path(input_path)
    return str(p.parent / f"{p.stem}_reconstructed.docx")


def main(argv: list[str] | None = None) -> int:
    """
    Main entry point for the CLI.

    Args:
        argv: Command-line arguments (defaults to sys.argv[1:]).

    Returns:
        Exit code (0 for success, 1 for error).
    """
    parser = build_parser()
    args = parser.parse_args(argv)

    _configure_logging(args.verbose, args.quiet)
    logger = logging.getLogger("doc_extract_reconstruct")

    # Validate input file
    if not os.path.isfile(args.input):
        print(f"Error: Input file not found: {args.input}", file=sys.stderr)
        return 1

    # Determine output path
    output_path = args.output or _derive_output_path(args.input)

    # Ensure output directory exists
    output_dir = os.path.dirname(output_path)
    if output_dir:
        os.makedirs(output_dir, exist_ok=True)

    # Detect and log input type
    try:
        input_type = detect_input_type(args.input)
        if not args.quiet:
            print(f"Input type detected: {input_type.name}")
            print(f"Processing: {args.input}")
            print(f"Output:     {output_path}")
    except (FileNotFoundError, ValueError) as e:
        print(f"Error: {e}", file=sys.stderr)
        return 1

    # Run the pipeline
    try:
        result_path = route(
            file_path=args.input,
            output_path=output_path,
            math_ocr=args.math_ocr,
            dpi=args.dpi,
        )

        if not args.quiet:
            print(f"\n[OK] Reconstruction complete: {result_path}")

            # Check if a confidence report was generated
            report_path = args.confidence_report
            if report_path is None:
                report_path = str(
                    Path(output_path).parent /
                    f"{Path(output_path).stem}_confidence_report.txt"
                )
            if os.path.isfile(report_path):
                print(f"[OK] Confidence report: {report_path}")

        return 0

    except ImportError as e:
        print(f"\nError: Missing dependency — {e}", file=sys.stderr)
        print("Run: pip install -r requirements.txt", file=sys.stderr)
        return 1

    except Exception as e:
        logger.exception("Pipeline failed")
        print(f"\nError: {e}", file=sys.stderr)
        return 1
