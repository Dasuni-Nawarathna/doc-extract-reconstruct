"""
app.py — Flask web application for doc-extract-reconstruct.

Provides a modern web UI for uploading files (.docx, .pdf, .png, .jpg)
and converting them into reconstructed .docx files with preserved formatting.
"""

import os
import uuid
import logging
import time
from pathlib import Path

from flask import (
    Flask, render_template, request, jsonify,
    send_file, after_this_request,
)

from doc_extract_reconstruct.router import route, detect_input_type

# ── App Setup ──────────────────────────────────────────────────────────

app = Flask(__name__, template_folder="templates", static_folder="static")
app.config['MAX_CONTENT_LENGTH'] = 50 * 1024 * 1024  # 50 MB max upload

UPLOAD_DIR = Path(__file__).parent / "uploads"
OUTPUT_DIR = Path(__file__).parent / "outputs"
UPLOAD_DIR.mkdir(exist_ok=True)
OUTPUT_DIR.mkdir(exist_ok=True)

ALLOWED_EXTENSIONS = {'.docx', '.pdf', '.png', '.jpg', '.jpeg', '.tiff', '.bmp'}

logging.basicConfig(
    level=logging.INFO,
    format="%(asctime)s [%(levelname)s] %(name)s: %(message)s",
    datefmt="%H:%M:%S",
)
logger = logging.getLogger(__name__)


# ── Routes ─────────────────────────────────────────────────────────────

@app.route("/")
def index():
    """Serve the main page."""
    return render_template("index.html")


@app.route("/api/convert", methods=["POST"])
def convert():
    """
    Handle file upload and conversion.
    Returns JSON with download URL on success.
    """
    if "file" not in request.files:
        return jsonify({"error": "No file uploaded"}), 400

    file = request.files["file"]

    if not file.filename:
        return jsonify({"error": "No file selected"}), 400

    # Validate extension
    ext = Path(file.filename).suffix.lower()
    if ext not in ALLOWED_EXTENSIONS:
        return jsonify({
            "error": f"Unsupported file type '{ext}'. Accepted: {', '.join(sorted(ALLOWED_EXTENSIONS))}"
        }), 400

    # Save uploaded file with a unique name
    job_id = str(uuid.uuid4())[:8]
    safe_name = f"{job_id}_{Path(file.filename).name}"
    input_path = UPLOAD_DIR / safe_name
    file.save(str(input_path))

    # Determine output path
    output_name = f"{job_id}_{Path(file.filename).stem}_reconstructed.docx"
    output_path = OUTPUT_DIR / output_name

    try:
        # Detect type
        input_type = detect_input_type(str(input_path))

        # Run the conversion pipeline
        start_time = time.time()
        result_path = route(
            file_path=str(input_path),
            output_path=str(output_path),
            math_ocr="auto",
            dpi=300,
        )
        elapsed = time.time() - start_time

        # Read confidence report if it exists
        report_name = f"{Path(output_path).stem}_confidence_report.txt"
        report_path = OUTPUT_DIR / report_name
        report_content = ""
        if report_path.exists():
            report_content = report_path.read_text(encoding="utf-8")

        return jsonify({
            "success": True,
            "job_id": job_id,
            "input_type": input_type.name,
            "output_file": output_name,
            "elapsed": round(elapsed, 2),
            "report": report_content,
        })

    except Exception as e:
        logger.exception("Conversion failed for %s", file.filename)
        return jsonify({"error": str(e)}), 500

    finally:
        # Clean up input file
        if input_path.exists():
            input_path.unlink(missing_ok=True)


@app.route("/api/download/<filename>")
def download(filename):
    """Serve a converted file for download."""
    file_path = OUTPUT_DIR / filename

    if not file_path.exists():
        return jsonify({"error": "File not found"}), 404

    @after_this_request
    def cleanup(response):
        """Clean up the output file after sending."""
        try:
            file_path.unlink(missing_ok=True)
            # Also clean up the confidence report
            report_path = OUTPUT_DIR / f"{file_path.stem}_confidence_report.txt"
            report_path.unlink(missing_ok=True)
        except Exception:
            pass
        return response

    return send_file(
        str(file_path),
        as_attachment=True,
        download_name=filename.split("_", 1)[1] if "_" in filename else filename,
    )


# ── Main ───────────────────────────────────────────────────────────────

if __name__ == "__main__":
    app.run(debug=True, port=5000)
