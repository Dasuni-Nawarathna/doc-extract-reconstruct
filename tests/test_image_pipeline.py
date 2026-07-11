import os
import sys
import pytest
from pathlib import Path
from PIL import Image, ImageDraw, ImageFont

# Add project root to path
sys.path.insert(0, str(Path(__file__).parent.parent))

from docx import Document
from doc_extract_reconstruct.image_pipeline import process_image
from doc_extract_reconstruct.router import detect_input_type, InputType

FIXTURES_DIR = Path(__file__).parent / "fixtures"
TEST_IMG = FIXTURES_DIR / "temp_test_image.png"
TEST_OUT_DOCX = FIXTURES_DIR / "temp_test_output.docx"

@pytest.fixture(autouse=True)
def cleanup_files():
    """Ensure temporary test files are cleaned up."""
    yield
    if TEST_IMG.exists():
        TEST_IMG.unlink(missing_ok=True)
    if TEST_OUT_DOCX.exists():
        TEST_OUT_DOCX.unlink(missing_ok=True)
    for f in FIXTURES_DIR.glob("temp_test_output*"):
        f.unlink(missing_ok=True)

def create_test_image(text, filepath):
    # Create simple image with text using PIL
    img = Image.new('RGB', (800, 150), color='white')
    d = ImageDraw.Draw(img)
    try:
        font = ImageFont.truetype("arial.ttf", 20)
    except IOError:
        font = ImageFont.load_default()
        
    d.text((20, 30), text, fill='black', font=font)
    img.save(str(filepath))

class TestImagePipeline:
    """Test cases for the Image-to-DOCX extraction and reconstruction pipeline."""

    def test_image_type_detection(self):
        create_test_image("Test", TEST_IMG)
        assert detect_input_type(str(TEST_IMG)) == InputType.IMAGE

    def test_image_text_extraction(self):
        test_text = "The quick brown fox jumps over the lazy dog"
        create_test_image(test_text, TEST_IMG)
        
        # Process the image to docx
        process_image(
            str(TEST_IMG),
            str(TEST_OUT_DOCX),
            math_ocr="none"
        )
        
        assert TEST_OUT_DOCX.exists()
        
        # Load and verify docx content
        doc = Document(str(TEST_OUT_DOCX))
        full_text = " ".join([p.text.strip() for p in doc.paragraphs])
        
        # Assert that the extracted text contains key words
        # Lowercase comparisons to avoid minor OCR casing differences
        full_text_lower = full_text.lower()
        assert "quick" in full_text_lower
        assert "brown" in full_text_lower
        assert "jumps" in full_text_lower
        assert "lazy" in full_text_lower
