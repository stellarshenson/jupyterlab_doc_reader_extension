"""Tests for document conversion handlers."""

import os
import tempfile
import pytest
from io import BytesIO
from unittest.mock import MagicMock, patch, PropertyMock

# Import conversion libraries
from pptx import Presentation
from pptx.util import Inches, Pt
from docx import Document
from reportlab.pdfgen import canvas
from reportlab.pdfbase import pdfmetrics
from reportlab.pdfbase.ttfonts import TTFont
from reportlab.lib import colors


def create_mock_handler():
    """Create a handler instance with mocked log property."""
    from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler
    handler = object.__new__(DocumentConverterHandler)
    # Set the _log attribute directly to bypass property
    object.__setattr__(handler, '_log', MagicMock())
    return handler


class TestPPTXConversion:
    """Test PPTX to PDF conversion functionality."""

    @pytest.fixture
    def sample_pptx(self):
        """Create a sample PPTX file for testing."""
        prs = Presentation()

        # Add title slide
        title_slide_layout = prs.slide_layouts[0]
        slide = prs.slides.add_slide(title_slide_layout)
        title = slide.shapes.title
        subtitle = slide.placeholders[1]
        title.text = "Test Presentation"
        subtitle.text = "Created for unit testing"

        # Add content slide
        bullet_slide_layout = prs.slide_layouts[1]
        slide2 = prs.slides.add_slide(bullet_slide_layout)
        shapes = slide2.shapes
        title_shape = shapes.title
        body_shape = shapes.placeholders[1]
        title_shape.text = "Test Slide"
        tf = body_shape.text_frame
        tf.text = "First bullet point"
        p = tf.add_paragraph()
        p.text = "Second bullet point"
        p.level = 1
        p = tf.add_paragraph()
        p.text = "Polish characters: ąćęłńóśźż"
        p.level = 0

        # Save to temp file
        with tempfile.NamedTemporaryFile(suffix='.pptx', delete=False) as f:
            prs.save(f.name)
            yield f.name

        # Cleanup
        os.unlink(f.name)

    def test_pptx_loads_correctly(self, sample_pptx):
        """Test that sample PPTX file loads correctly."""
        prs = Presentation(sample_pptx)
        assert len(prs.slides) == 2

    def test_pptx_has_slide_dimensions(self, sample_pptx):
        """Test that PPTX has valid slide dimensions."""
        prs = Presentation(sample_pptx)
        assert prs.slide_width > 0
        assert prs.slide_height > 0

    def test_pptx_conversion_produces_pdf(self, sample_pptx):
        """Test that PPTX conversion produces valid PDF bytes."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            pdf_bytes = handler._convert_pptx_to_pdf(sample_pptx)

        # Check that we got PDF bytes
        assert pdf_bytes is not None
        assert len(pdf_bytes) > 0
        # PDF files start with %PDF
        assert pdf_bytes[:4] == b'%PDF'

    def test_pptx_conversion_multiple_slides(self, sample_pptx):
        """Test that conversion handles multiple slides."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            pdf_bytes = handler._convert_pptx_to_pdf(sample_pptx)

        # PDF should contain multiple pages
        assert pdf_bytes is not None
        # Check PDF contains page references
        assert b'/Page' in pdf_bytes


class TestDOCXConversion:
    """Test DOCX to PDF conversion functionality."""

    @pytest.fixture
    def sample_docx(self):
        """Create a sample DOCX file for testing."""
        doc = Document()
        doc.add_heading('Test Document', 0)
        doc.add_paragraph('This is a test paragraph.')
        doc.add_paragraph('Polish characters: ąćęłńóśźż')
        doc.add_heading('Section 1', level=1)
        doc.add_paragraph('More content here.')

        # Save to temp file
        with tempfile.NamedTemporaryFile(suffix='.docx', delete=False) as f:
            doc.save(f.name)
            yield f.name

        # Cleanup
        os.unlink(f.name)

    def test_docx_loads_correctly(self, sample_docx):
        """Test that sample DOCX file loads correctly."""
        doc = Document(sample_docx)
        assert len(doc.paragraphs) > 0

    def test_docx_conversion_produces_pdf(self, sample_docx):
        """Test that DOCX conversion produces valid PDF bytes."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            pdf_bytes = handler._convert_docx_to_pdf(sample_docx)

        # Check that we got PDF bytes
        assert pdf_bytes is not None
        assert len(pdf_bytes) > 0
        # PDF files start with %PDF
        assert pdf_bytes[:4] == b'%PDF'


class TestFileTypeRouting:
    """Test file type detection and routing."""

    def test_unsupported_ppt_format(self):
        """Test that legacy PPT format raises appropriate error."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            with pytest.raises(Exception) as exc_info:
                handler._convert_to_pdf('/path/to/file.ppt')
            assert 'PPT format not supported' in str(exc_info.value)

    def test_unsupported_doc_format(self):
        """Test that legacy DOC format raises appropriate error."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            with pytest.raises(Exception) as exc_info:
                handler._convert_to_pdf('/path/to/file.doc')
            assert 'DOC' in str(exc_info.value) or 'not supported' in str(exc_info.value)

    def test_unsupported_rtf_format(self):
        """Test that RTF format raises appropriate error."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            with pytest.raises(Exception) as exc_info:
                handler._convert_to_pdf('/path/to/file.rtf')
            assert 'RTF' in str(exc_info.value) or 'not supported' in str(exc_info.value)

    def test_unknown_file_type(self):
        """Test that unknown file types raise appropriate error."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            with pytest.raises(Exception) as exc_info:
                handler._convert_to_pdf('/path/to/file.xyz')
            assert 'Unsupported file type' in str(exc_info.value)


class TestUnicodeFonts:
    """Test Unicode font registration."""

    def test_font_registration_method_exists(self):
        """Test that font registration method exists."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        # Method should exist and be callable
        assert hasattr(DocumentConverterHandler, '_register_unicode_fonts')
        assert callable(DocumentConverterHandler._register_unicode_fonts)

    def test_font_registration_runs_without_error(self):
        """Test that font registration runs without error."""
        from jupyterlab_doc_reader_extension.handlers import DocumentConverterHandler

        with patch.object(DocumentConverterHandler, 'log', new_callable=PropertyMock) as mock_log:
            mock_log.return_value = MagicMock()
            handler = object.__new__(DocumentConverterHandler)
            # Should not raise
            handler._register_unicode_fonts()


class TestImportAvailability:
    """Test that required imports are available."""

    def test_reportlab_available(self):
        """Test that reportlab imports are available."""
        from jupyterlab_doc_reader_extension.handlers import REPORTLAB_AVAILABLE
        assert REPORTLAB_AVAILABLE is True

    def test_docx_available(self):
        """Test that python-docx imports are available."""
        from jupyterlab_doc_reader_extension.handlers import DOCX_AVAILABLE
        assert DOCX_AVAILABLE is True

    def test_pptx_available(self):
        """Test that python-pptx imports are available."""
        from jupyterlab_doc_reader_extension.handlers import PPTX_AVAILABLE
        assert PPTX_AVAILABLE is True

    def test_pil_available(self):
        """Test that PIL imports are available."""
        from jupyterlab_doc_reader_extension.handlers import PIL_AVAILABLE
        assert PIL_AVAILABLE is True


class TestPDFGeneration:
    """Test basic PDF generation with reportlab."""

    def test_create_simple_pdf(self):
        """Test that we can create a simple PDF."""
        pdf_buffer = BytesIO()
        c = canvas.Canvas(pdf_buffer)
        c.drawString(100, 750, "Test PDF")
        c.showPage()
        c.save()

        pdf_bytes = pdf_buffer.getvalue()
        assert pdf_bytes[:4] == b'%PDF'
        assert len(pdf_bytes) > 100

    def test_pdf_with_colors(self):
        """Test PDF generation with colors."""
        pdf_buffer = BytesIO()
        c = canvas.Canvas(pdf_buffer)
        c.setFillColor(colors.red)
        c.drawString(100, 750, "Red text")
        c.setFillColor(colors.blue)
        c.drawString(100, 700, "Blue text")
        c.showPage()
        c.save()

        pdf_bytes = pdf_buffer.getvalue()
        assert pdf_bytes[:4] == b'%PDF'
