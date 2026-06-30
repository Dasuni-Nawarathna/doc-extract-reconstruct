#!/usr/bin/env python3
"""
extract.py — Entry point for doc-extract-reconstruct.

Usage:
    python extract.py <input_file> -o <output.docx> [options]

See `python extract.py --help` for full usage information.
"""

import sys

from doc_extract_reconstruct.cli import main

if __name__ == "__main__":
    sys.exit(main())
